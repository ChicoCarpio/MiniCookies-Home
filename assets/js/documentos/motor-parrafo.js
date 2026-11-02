/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Cortador de párrafos de Word (Fase 6, etapa 6.2). Toma un párrafo del modelo (docx-modelo.js) y lo corta en líneas con las reglas
// de Word que el maquetador del sitio midió contra Word (docx-maquetador.js): sangrías, sangría francesa, etiqueta de lista, tabulaciones
// (paradas propias y por defecto; izquierda, centro, derecha y decimal), alto de línea por fuente, interlineado (múltiplo, «al menos» y
// «exacto»), alineación (izquierda, centro, derecha, justificado), imágenes en línea y campos de página. Unidad: puntos (pt).
//
// Las reglas de corte de palabras son las de MotorTexto (motor-texto.js): codicioso, los espacios del final cuelgan, una palabra más larga
// que la línea se parte por letras (búsqueda binaria, sin partir un par sustituto). Además se corta después de un guion. El texto se mide
// con un canvas (ctx) y las fuentes de PptxTexto (anchos de Office); hay que cargar antes las fuentes con cargarFuentes.
// Cada línea lleva desplazamientos (desde, hasta) dentro del texto plano del párrafo (bloque.texto) para el cursor.
//
// Uso: const motor = MotorParrafo.crear(ctx, modeloEstilos);
//      const r = motor.romper(bloque, anchoDisponiblePt, { pagina: 3, paginas: 10 })  → { lineas, alto, antes, despues }
//      motor.puntoDe(r, desplazamiento) → { x, y, base, alto, linea }     motor.posicionDe(r, x, y) → desplazamiento
(function () {
  'use strict';
  const PX = 96 / 72;                       // px por pt: se mide a tamaño de pantalla y se vuelve a pt
  const ESP = /([  -  -​　]+|\t)/;   // espacios que cortan (no el espacio duro U+00A0)
  const SERIF = /cambria|georgia|garamond|palatino|bookman|century|constantia|baskerville|times|book antiqua|rockwell/i;
  const MONO = /consolas|courier|lucida console|mono|cascadia/i;
  const PARTE_ARRIBA = { exact: 0.8, atLeast: 1, auto: 0 };   // qué parte del espacio de más de una línea queda sobre el texto (medido contra Word)
  const MAX_ATOMOS = 400000;
  const MAX_ANCHO_TAB = 3168;
  const GUION = /(?<=[\p{L}\p{N}][-‐])(?=[\p{L}])|(?<=[–—])(?=[^\s–—])/u;   // dónde se puede cortar dentro de una palabra

  function crear(ctx, m, opciones = {}) {
    const kerning = opciones.kerning || 'auto';
    const metricas = new Map();
    let cache = new WeakMap();
    const invalidar = (b) => { if (b) cache.delete(b); else cache = new WeakMap(); };

    const familia = (nombre) => {
      const n = String(nombre || '').replace(/[^\p{L}\p{M}\p{N} ._&+-]/gu, '').trim() || 'Times New Roman';
      const base = typeof PptxTexto !== 'undefined' && PptxTexto.familia ? PptxTexto.familia(n) : `"${n}"`;
      return base.replace(/, "Segoe UI", sans-serif$/, '') + (MONO.test(n) ? ', monospace' : SERIF.test(n) ? ', serif' : ', "Segoe UI", sans-serif');
    };
    const cssDe = (r, tam) => `${r.i ? 'italic ' : ''}${r.b ? 'bold ' : ''}${(tam * PX).toFixed(3)}px ${familia(m.fuente(r.fuentes))}`;
    const medir = (css, s, espaciado) => {
      ctx.font = css; ctx.fontKerning = kerning; ctx.letterSpacing = espaciado ? espaciado * PX + 'px' : '0px';
      return ctx.measureText(s).width / PX;
    };
    // Subida y bajada naturales de esa fuente (con el tamaño del css) en pt.
    // Se mide a 100 px: a tamaños chicos el navegador redondea la subida y la bajada a píxeles enteros.
    const naturales = (css, tam) => {
      const grande = css.replace(/[0-9.]+px/, '100px');
      let v = metricas.get(grande);
      if (!v) {
        ctx.font = grande;
        const t = ctx.measureText('Hg');
        const asc = t.fontBoundingBoxAscent, desc = t.fontBoundingBoxDescent;
        v = asc + desc > 0 ? { asc: asc / 100, desc: desc / 100 } : { asc: 0.9, desc: 0.3 };
        metricas.set(grande, v);
      }
      return { asc: v.asc * tam, desc: v.desc * tam };
    };

    // Un trozo de formato → lo que hace falta para medir y dibujar.
    function formato(r) {
      const tam = r.tam !== undefined ? r.tam : 10;
      const vert = r.alineacionVertical === 'superscript' ? 1 : r.alineacionVertical === 'subscript' ? -1 : 0;
      const tamL = vert ? tam * 0.65 : r.versalitas && !r.mayusculas ? tam * 0.8 : tam;
      return { tam: tamL, subida: vert > 0 ? 0.33 * tam : vert < 0 ? -0.14 * tam : 0, css: cssDe(r, tamL), esp: r.espaciado || 0, escala: r.escala ? r.escala / 100 : 1, mayus: r.mayusculas || r.versalitas, vert };
    }

    // Piezas del modelo → átomos (palabras, espacios, tabulaciones, saltos, imágenes, campos).
    function atomos(b, opc) {
      const sal = [];
      const poner = (a) => { if (sal.length < MAX_ATOMOS) sal.push(a); };
      const piezas = [];
      // texto de derecha a izquierda: cada átomo lleva su nivel (impar = de derecha a izquierda); sin nada que reordenar, niv es null
      const nivBase = b.p && b.p.derechaAIzquierda ? 1 : 0;
      const niv = typeof Bidi !== 'undefined' ? Bidi.niveles(b.texto || '', nivBase === 1) : null;
      const porNivel = (s, d) => {
        const res = [];
        let ini = 0;
        for (let i = 1; i <= s.length; i++) if (i === s.length || niv[d + i] !== niv[d + ini]) { res.push([s.slice(ini, i), d + ini, niv[d + ini], d + i]); ini = i; }
        return res;
      };
      if (b.etiqueta) {
        const r = { ...b.etiqueta.r };
        if (b.etiqueta.fmt === 'bullet') delete r.fuentes;
        if (b.etiqueta.texto) piezas.push({ t: 'texto', s: b.etiqueta.texto, r, etiqueta: true, desde: 0, hasta: 0 });
        if (b.etiqueta.sufijo === 'tab') piezas.push({ t: 'tab', r, etiqueta: true, desde: 0, hasta: 0 });
        else if (b.etiqueta.sufijo === 'space') piezas.push({ t: 'texto', s: ' ', r, etiqueta: true, desde: 0, hasta: 0 });
      }
      for (const z of b.piezas) piezas.push(z);
      for (const z of piezas) {
        if (z.t === 'ancla') continue;
        const f = formato(z.r);
        const base = { r: z.r, f, desde: z.desde, hasta: z.hasta, etiqueta: !!z.etiqueta, enlace: z.enlace, nota: z.nota, nivel: nivBase };
        const delTexto = niv && z.t === 'texto' && !z.etiqueta;
        if (z.t === 'texto' || z.t === 'campo') {
          const origen = z.t === 'campo' ? String(z.nombre === 'NUMPAGES' ? opc.paginas || 1 : opc.pagina || 1) : z.s;
          const partes = origen.split(ESP);
          let pos = 0;
          for (const parte of partes) {
            if (!parte) continue;
            const desde = z.desde + (z.t === 'campo' ? 0 : pos);
            pos += parte.length;
            const hasta = z.t === 'campo' ? z.hasta : z.desde + pos;
            if (parte === '\t') { poner({ ...base, k: 't', desde, hasta }); continue; }
            if (/^[  -  -​　]+$/.test(parte)) { poner({ ...base, k: 'e', texto: parte, desde, hasta, ancho: medir(f.css, parte, f.esp) * f.escala, nivel: delTexto && niv[desde] !== undefined ? niv[desde] : nivBase }); continue; }
            const subs = parte.split(GUION).filter(Boolean);
            let p2 = 0;
            subs.forEach((s, i) => {
              const d = desde + (z.t === 'campo' ? 0 : p2); p2 += s.length;
              const h = z.t === 'campo' ? z.hasta : desde + p2;
              const trozos = delTexto ? porNivel(s, d) : [[s, d, nivBase, h]];
              trozos.forEach(([ts, td, tn, th], ti) => {
                const mostrado = (f.mayus ? ts.toUpperCase() : ts).replace(/­/g, '');
                poner({ ...base, k: z.t === 'campo' ? 'c' : 'p', texto: mostrado, desde: td, hasta: Math.max(td, Math.min(th, z.hasta)), corte: i < subs.length - 1 && ti === trozos.length - 1, ancho: medir(f.css, mostrado, f.esp) * f.escala, nivel: tn });
              });
            });
          }
        } else if (z.t === 'tab') poner({ ...base, k: 't' });
        else if (z.t === 'br') poner({ ...base, k: 'b' });
        else if (z.t === 'pagina' || z.t === 'columna') poner({ ...base, k: 'pg', columna: z.t === 'columna' });
        else if ((z.t === 'imagen' || z.t === 'cuadro' || z.t === 'dibujo') && z.enLinea) poner({ ...base, k: 'i', pieza: z, ancho: z.ancho, alto: z.alto, corte: true });
        else if (z.t === 'imagen' || z.t === 'cuadro' || z.t === 'dibujo') poner({ ...base, k: 'f', pieza: z, ancho: 0 });   // flotante: no ocupa lugar en la línea
      }
      return sal;
    }

    // Las piezas de una línea en el orden en que se ven medidas desde el borde donde empieza el párrafo (reglas L1 y L2 de Unicode).
    // Sin texto de derecha a izquierda es el orden normal.
    function ordenDeLinea(ln, rtl) {
      const at = ln.atomos, base = rtl ? 1 : 0;
      if (!rtl && !at.some((a) => a.nivel)) return at;
      const niv = at.map((a) => (a.k === 't' ? base : a.nivel === undefined ? base : a.nivel));
      for (let i = at.length - 1; i >= 0 && (at[i].k === 'e' || at[i].k === 'b' || at[i].k === 'pg' || at[i].k === 'f'); i--) niv[i] = base;
      at.forEach((a, i) => { a.nivel = niv[i]; });
      const visual = Bidi.ordenVisual(niv).map((i) => at[i]);
      return rtl ? visual.reverse() : visual;
    }

    function romper(b, anchoDisp, opc = {}) {
      if (opc.flotantes && opc.flotantes.length) return cortar(b, anchoDisp, opc);   // el corte depende de dónde cae el párrafo en la hoja
      const clave = anchoDisp + '|' + (opc.pagina || 0) + '|' + (opc.paginas || 0) + '|' + (opc.limites || '');
      const previo = cache.get(b);
      if (previo && previo.clave === clave) return previo.r;
      const r = cortar(b, anchoDisp, opc);
      cache.set(b, { clave, r });
      return r;
    }

    function cortar(b, anchoDisp, opc) {
      const p = b.p;
      const sang = p.sangria || {};
      const izq = sang.izq || 0, der = sang.der || 0, colg = sang.colgante || 0, prim = colg ? -colg : sang.primera || 0;
      const esp = p.espacio || {};
      const alin = { center: 'c', right: 'r', end: 'r', both: 'j', distribute: 'd' }[p.jc] || 'l';
      const lista = atomos(b, opc);
      const paso = (m.ajustes && m.ajustes.tabPorDefecto) || 36;
      const paradas = (p.tabs || []).filter((t) => t.tipo !== 'clear' && Number.isFinite(t.pos)).sort((a, c) => a.pos - c.pos);
      if (prim < 0) paradas.push({ tipo: 'left', pos: izq });
      paradas.sort((a, c) => a.pos - c.pos);
      const lineas = [];
      let l = null;
      // Flotantes (cuadrado, ajuste estrecho, arriba y abajo) de la hoja: rectángulos { x0, x1, y0, y1, lado } en coordenadas de la hoja;
      // opc.x0 es el borde izquierdo del área de texto y opc.y0 el de arriba del párrafo.
      const flot = opc.flotantes && opc.flotantes.length ? opc.flotantes : null;
      const marcaF = formato(b.marca || b.r), vm = naturales(marcaF.css, marcaF.tam);
      const factor = esp.linea !== undefined && (esp.regla || 'auto') === 'auto' && Math.abs(esp.linea - 240) > 1 ? esp.linea / 240 : 1;
      const estimado = esp.linea !== undefined && esp.regla === 'exact' ? esp.linea / 20 : (vm.asc + vm.desc) * factor;
      let yAcum = 0;
      const nueva = (primera) => {
        let inicio = Math.max(izq + (primera ? prim : 0), 0), limiteDer = anchoDisp - der, gap = 0;
        if (flot) {
          const x0 = opc.x0 || 0;
          for (let vuelta = 0; vuelta < 8; vuelta++) {
            const top = (opc.y0 || 0) + yAcum + gap;
            let movido = false;
            for (const q of flot) if (q.lado === 'arribaAbajo' && top < q.y1 && top + estimado > q.y0) { gap += q.y1 - top; movido = true; break; }
            if (!movido) break;
          }
          const top = (opc.y0 || 0) + yAcum + gap;
          // lo que queda libre de la línea entre los dibujos que toca: el texto va en el hueco más ancho
          let huecos = [[inicio, limiteDer]];
          for (const q of flot) {
            if (q.lado === 'arribaAbajo' || q.lado === 'ninguno' || !(top < q.y1 && top + estimado > q.y0)) continue;
            const a = q.x0 - x0, z = q.x1 - x0;
            huecos = huecos.flatMap(([i, j]) => (z <= i || a >= j ? [[i, j]] : [[i, Math.min(j, a)], [Math.max(i, z), j]].filter(([u, v]) => v - u > 0.5)));
          }
          if (huecos.length) { const mejor = huecos.reduce((m1, h) => (h[1] - h[0] > m1[1] - m1[0] ? h : m1)); inicio = mejor[0]; limiteDer = mejor[1]; }
        }
        l = { atomos: [], inicio, ancho: 0, limite: Math.max(limiteDer - inicio, 1), final: false, primera, gapAntes: gap };
        yAcum += gap + estimado;
        lineas.push(l);
      };
      nueva(true);
      const anchoSinFinal = (ln) => { let w = ln.ancho; for (let i = ln.atomos.length - 1; i >= 0 && ln.atomos[i].k === 'e'; i--) w -= ln.atomos[i].ancho; return w; };
      const esPalabra = (a) => a && (a.k === 'p' || a.k === 'c');
      const inicioToken = (i) => i === 0 || lista[i].k === 'i' || lista[i - 1].corte || !esPalabra(lista[i - 1]) || !esPalabra(lista[i]);
      for (let i = 0; i < lista.length; i++) {
        const a = lista[i];
        if (a.k === 'b') { l.atomos.push(a); l.salto = true; l.final = true; nueva(false); continue; }
        if (a.k === 'pg') { l.atomos.push(a); if (a.columna) l.saltoColumna = true; else l.saltoPagina = true; l.final = true; nueva(false); continue; }
        if (a.k === 'f') { l.atomos.push(a); continue; }
        if (a.k === 'e') { l.atomos.push(a); l.ancho += a.ancho; continue; }
        if (a.k === 't') {
          const x0 = l.inicio + l.ancho;
          const meta = paradas.find((s) => s.pos > x0 + 0.5) || { tipo: 'left', pos: (Math.floor((x0 + 0.5) / paso) + 1) * paso };
          let ancho = meta.pos - x0;
          if (meta.tipo === 'right' || meta.tipo === 'center' || meta.tipo === 'decimal') {
            let siguiente = 0;
            for (let k = i + 1; k < lista.length && lista[k].k !== 't' && lista[k].k !== 'b' && lista[k].k !== 'pg'; k++) {
              const t = lista[k];
              if (meta.tipo === 'decimal' && /[.,]/.test(t.texto || '')) { siguiente += medir(t.f.css, t.texto.slice(0, t.texto.search(/[.,]/)), t.f.esp) * t.f.escala; break; }
              siguiente += t.ancho || 0;
            }
            ancho -= meta.tipo === 'center' ? siguiente / 2 : siguiente;
          }
          ancho = Math.min(Math.max(ancho, 0), MAX_ANCHO_TAB);
          a.ancho = ancho; a.guia = meta.guia; a.paradaTipo = meta.tipo;
          l.atomos.push(a); l.ancho += ancho;
          continue;
        }
        // palabra, campo o imagen
        const inicio = inicioToken(i);
        if (inicio) {
          let w = 0;
          for (let k = i; k < lista.length && (k === i || (esPalabra(lista[k]) && lista[k].k !== 'i' && !lista[k - 1].corte)); k++) w += lista[k].ancho;
          if (l.atomos.some((q) => q.k !== 'e' && q.k !== 'f') && anchoSinFinal(l) > 0 && l.ancho + w > l.limite + 0.01) nueva(false);
        } else if (anchoSinFinal(l) > 0 && l.ancho + a.ancho > l.limite + 0.01) nueva(false);   // un token más largo que una línea: se corta donde llega
        if (a.k === 'p' && a.ancho > l.limite + 0.01) {
          // palabra más larga que la línea: se parte por letras (búsqueda binaria; no se parte un par sustituto)
          let resto = a.texto, restoOrigen = a.hasta - a.desde, consumido = 0;
          const mismoLargo = a.hasta - a.desde === a.texto.length;
          while (resto) {
            let n = resto.length;
            const entra = (k) => l.ancho + medir(a.f.css, resto.slice(0, k), a.f.esp) * a.f.escala <= l.limite + 0.01;
            if (n > 1 && !entra(n)) {
              let lo = 1, hi = n - 1;
              while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (entra(mid)) lo = mid; else hi = mid - 1; }
              n = lo;
            }
            if (n < resto.length && /[\uD800-\uDBFF]/.test(resto[n - 1])) n = n > 1 ? n - 1 : 2;
            const trozo = resto.slice(0, n);
            const w = medir(a.f.css, trozo, a.f.esp) * a.f.escala;
            const d0 = a.desde + (mismoLargo ? consumido : 0);
            l.atomos.push({ ...a, texto: trozo, ancho: w, desde: Math.min(a.hasta, d0), hasta: Math.min(a.hasta, d0 + (mismoLargo ? n : 0)) });
            l.ancho += w; consumido += n;
            resto = resto.slice(n);
            if (resto) nueva(false);
          }
          if (!mismoLargo) { const u = l.atomos[l.atomos.length - 1]; u.hasta = a.hasta; }
          continue;
        }
        l.atomos.push(a); l.ancho += a.ancho;
      }
      l.final = true;

      // desplazamientos continuos
      let cursor = 0;
      for (const ln of lineas) {
        ln.desde = cursor;
        let h = cursor;
        for (const a of ln.atomos) if (!a.etiqueta && a.hasta > h) h = a.hasta;
        ln.hasta = h; cursor = h;
      }

      // altura de cada línea y posición vertical
      const marca = formato(b.marca || b.r);
      const regla = esp.regla || 'auto';
      let y = 0;
      lineas.forEach((ln, il) => {
        ln.yIni = y;   // dónde empieza la línea, sin el hueco que deja un dibujo de «arriba y abajo»
        y += ln.gapAntes || 0;
        const conTexto = ln.atomos.filter((a) => a.k !== 'f' && a.k !== 'b' && a.k !== 'pg');
        let asc = 0, desc = 0;
        for (const a of conTexto) {
          if (a.k === 'i') { asc = Math.max(asc, a.alto); continue; }
          const v = naturales(a.f.css, a.f.tam);
          asc = Math.max(asc, v.asc); desc = Math.max(desc, v.desc);
        }
        // una línea sin texto (párrafo vacío, o solo un salto) usa la fuente de la marca de párrafo, o la del salto
        if (!conTexto.length) {
          const brk = ln.atomos.find((a) => a.k === 'b' || a.k === 'pg');
          const f = brk ? brk.f : marca;
          const v = naturales(f.css, f.tam);
          asc = Math.max(asc, v.asc); desc = Math.max(desc, v.desc);
        }
        const nat = asc + desc;
        let alto = nat, arriba = 0;
        if (esp.linea !== undefined) {
          if (regla === 'exact') alto = esp.linea / 20;
          else if (regla === 'atLeast') alto = Math.max(esp.linea / 20, nat);
          else if (Math.abs(esp.linea - 240) > 1) alto = nat * esp.linea / 240;
          arriba = (alto - nat) * (PARTE_ARRIBA[regla] !== undefined ? PARTE_ARRIBA[regla] : 0);
        }
        ln.alto = alto; ln.y = y; ln.base = arriba + asc; ln.naturalAsc = asc;
        y += alto;
        ln.ultima = il === lineas.length - 1;
      });

      // posición horizontal (alineación) de cada pieza
      const rtl = !!p.derechaAIzquierda;
      for (const ln of lineas) {
        const usado = anchoSinFinal(ln);
        const libre = ln.limite - usado;
        let x = ln.inicio, extra = 0;
        if (alin === 'c') x += libre / 2;
        else if (alin === 'r') x += libre;
        else if ((alin === 'j' && !ln.final) || alin === 'd') {
          let ult = -1;
          for (let i = ln.atomos.length - 1; i >= 0; i--) if (ln.atomos[i].k !== 'e' && ln.atomos[i].k !== 'f') { ult = i; break; }
          const espacios = ln.atomos.slice(0, ult + 1).filter((a) => a.k === 'e').length;
          if (espacios && libre > 0) extra = libre / espacios;
        }
        ln.x0 = x;
        let ult = -1;
        for (let i = ln.atomos.length - 1; i >= 0; i--) if (ln.atomos[i].k !== 'e' && ln.atomos[i].k !== 'f') { ult = i; break; }
        ln.atomos.forEach((a, i) => { a.li = i; });
        // las piezas se colocan desde el borde donde empieza el párrafo, en el orden en que se ven; en un párrafo de derecha a izquierda se espeja
        for (const a of ordenDeLinea(ln, rtl)) { a.x = x; x += a.ancho + (a.k === 'e' && a.li < ult ? extra : 0); }
        if (rtl) { for (const a of ln.atomos) a.x = anchoDisp - (a.x + a.ancho); ln.x0 = anchoDisp - ln.x0; }
      }
      const nAuto = (v, auto) => (auto ? 14 : v || 0);
      // el borde de arriba y el de abajo de un párrafo (con su espacio al texto) ocupan lugar, como en Word
      const bd = p.bordes || {};
      const hueco = (x) => (x && x.tipo && x.tipo !== 'nil' && x.tipo !== 'none' ? (x.espacio || 0) + Math.max(x.ancho === undefined ? 0.5 : x.ancho, 0.25) : 0);
      const bordeArriba = hueco(bd.top), bordeAbajo = hueco(bd.bottom);
      return { lineas, alto: y, antes: nAuto(esp.antes, esp.antesAuto) + bordeArriba, despues: nAuto(esp.despues, esp.despuesAuto) + bordeAbajo, bordeArriba, bordeAbajo, anchoDisp };
    }

    // ───────── Posición ↔ punto (coordenadas del párrafo; x desde el borde izquierdo del área de texto) ─────────
    function lineaDe(r, d) {
      for (let i = 0; i < r.lineas.length; i++) if (d < r.lineas[i].hasta || i === r.lineas.length - 1) return i;
      return 0;
    }
    const impar = (a) => (a.nivel & 1) === 1;
    // x de un desplazamiento dentro de una pieza; en una pieza de derecha a izquierda el principio del texto está a su derecha
    function xEnAtomo(a, d) {
      const k = Math.max(0, Math.min(d - a.desde, a.hasta - a.desde)), der = impar(a);
      const principio = der ? a.x + a.ancho : a.x, final = der ? a.x : a.x + a.ancho;
      if (k <= 0) return principio;
      const n = a.hasta - a.desde;
      if (k >= n) return final;
      // un grupo de espacios, o una palabra cuyo texto dibujado no tiene los mismos caracteres que el del archivo (guion suave, mayúsculas con ß…): el cursor avanza
      // en proporción al ancho, en vez de saltar de un extremo al otro
      if (a.k === 'e' || (a.k === 'p' && n !== a.texto.length)) return der ? a.x + a.ancho - a.ancho * k / n : a.x + a.ancho * k / n;
      if (a.k !== 'p') return final;
      const w = medir(a.f.css, a.texto.slice(0, k), a.f.esp) * a.f.escala;
      return der ? a.x + a.ancho - w : a.x + w;
    }
    // x (desde el borde del área de texto) del desplazamiento d dentro de la línea l; d se acota a los de la línea
    function xEnLinea(l, d) {
      const dd = Math.max(l.desde, Math.min(d, l.hasta));
      const a = l.atomos.find((q) => !q.etiqueta && q.k !== 'f' && dd >= q.desde && dd < q.hasta);
      if (a) return xEnAtomo(a, dd);
      const vis = l.atomos.filter((q) => !q.etiqueta && q.k !== 'f');
      const u = vis[vis.length - 1];
      return u ? (impar(u) ? u.x : u.x + u.ancho) : l.x0;
    }
    // Los tramos [x0, x1] que ocupa el texto entre dos desplazamientos de una línea (con texto mezclado son varios)
    function tramosEnLinea(l, d0, d1) {
      const sal = [];
      for (const a of l.atomos) {
        if (a.etiqueta || a.k === 'f') continue;
        const i = Math.max(d0, a.desde), j = Math.min(d1, a.hasta);
        if (j <= i) continue;
        const x0 = xEnAtomo(a, i), x1 = xEnAtomo(a, j);
        sal.push([Math.min(x0, x1), Math.max(x0, x1)]);
      }
      sal.sort((p, q) => p[0] - q[0]);
      const unidos = [];
      for (const t of sal) { const u = unidos[unidos.length - 1]; if (u && t[0] <= u[1] + 0.5) u[1] = Math.max(u[1], t[1]); else unidos.push(t.slice()); }
      return unidos;
    }
    function puntoDe(r, d) {
      const i = lineaDe(r, d);
      const l = r.lineas[i];
      return { x: xEnLinea(l, d), y: l.y, base: l.y + l.base, alto: l.alto, linea: i };
    }
    function posicionDe(r, x, y) {
      const l = r.lineas.find((q) => y < q.y + q.alto) || r.lineas[r.lineas.length - 1];
      let vis = l.atomos.filter((q) => !q.etiqueta && q.k !== 'f' && q.k !== 'b' && q.k !== 'pg');
      // en una línea cortada, los espacios que cuelgan al final no son un sitio donde poner el cursor
      if (!l.final) { let n = vis.length; while (n > 0 && vis[n - 1].k === 'e') n--; vis = vis.slice(0, n); }
      vis.sort((p, q) => p.x - q.x);
      if (!vis.length) return l.desde;
      // el desplazamiento del borde izquierdo y del derecho de una pieza
      const izq = (a) => (impar(a) ? a.hasta : a.desde), der = (a) => (impar(a) ? a.desde : a.hasta);
      let d = l.desde, hallado = false;
      for (const a of vis) {
        if (x >= a.x + a.ancho) { d = der(a); continue; }
        if (x <= a.x) { d = izq(a); hallado = true; break; }
        let mejor = a.desde, dist = Infinity;
        for (let k = 0; k <= a.hasta - a.desde; k++) {
          if (a.texto && k < a.texto.length && /[\uD800-\uDBFF]/.test(a.texto[k - 1] || '')) continue;
          const dx = Math.abs(xEnAtomo(a, a.desde + k) - x);
          if (dx < dist) { dist = dx; mejor = a.desde + k; }
        }
        d = mejor; hallado = true; break;
      }
      if (!hallado && l.salto) d = Math.min(d, l.hasta - 1);
      return d;
    }

    // Fuentes que usa un conjunto de párrafos, para cargarlas antes de medir.
    function fuentesDe(parrafos) {
      const usos = new Map();
      const ver = (r) => { const n = m.fuente(r.fuentes) || 'Times New Roman'; usos.set(n + '|' + !!r.b + '|' + !!r.i, { fuente: n, negrita: !!r.b, cursiva: !!r.i }); };
      for (const b of parrafos) { ver(b.marca || b.r); if (b.etiqueta) ver(b.etiqueta.r); for (const z of b.piezas) if (z.r) ver(z.r); }
      return [...usos.values()];
    }
    // Ancho mínimo (la palabra más larga) y máximo (todo en una línea) del contenido de un párrafo, para ajustar columnas de tablas.
    function contenido(b, opc = {}) {
      const lista = atomos(b, opc);
      let min = 0, max = 0, token = 0;
      for (const a of lista) {
        if (a.k === 'b' || a.k === 'pg') { max = 0; token = 0; continue; }
        const w = a.k === 't' ? 0 : a.ancho || 0;
        max += w;
        if (a.k === 'e' || a.k === 't') { min = Math.max(min, token); token = 0; }
        else { token += w; if (a.corte) { min = Math.max(min, token); token = 0; } }
      }
      min = Math.max(min, token);
      const sang = (b.p.sangria || {});
      const extra = (sang.izq || 0) + (sang.der || 0);
      return { min: min + extra, max: Math.max(...[max], 0) + extra };
    }
    return { contenido, romper, puntoDe, xEnLinea, tramosEnLinea, posicionDe, lineaDe, invalidar, fuentesDe, naturales, medir, formato };
  }

  globalThis.MotorParrafo = { crear };
})();

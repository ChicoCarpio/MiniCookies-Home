/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Motor de texto de la suite (Fase 4, etapas 4.1 y 4.2): maquetado incremental
// con consultas posición ↔ punto, para editar texto con cursor (Presentaciones
// ahora, Documentos después).
//
// Origen: adaptado de MiniCookies-Home/assets/js/pptx-texto.js (PptxTexto.maquetar,
// el que da el 99,4 % de cortes iguales a PowerPoint) el 2026-10-30. Las reglas
// de corte, interlineado, sangrías, viñetas, alineación y columnas son las mismas
// (no se recalibra nada aquí: un cambio de regla se hace en el original y se
// vuelve a copiar). Lo que se usa tal cual de PptxTexto, que debe estar cargado
// antes: familia, fuenteCss, kerningDe, colorCss, cargarFuentes.
// Diferencias con el original:
//   1. El maquetado se hace por párrafo y se guarda (WeakMap por objeto de
//      párrafo + clave de ancho/ajuste/viñeta): al editar, solo se vuelve a cortar
//      el párrafo cuyo objeto cambió. Los párrafos son inmutables: una edición
//      crea un objeto nuevo (o se llama a invalidar).
//   2. Cada línea, pieza y trozo lleva sus desplazamientos de carácter (desde,
//      hasta) dentro del párrafo: texto de los trozos en orden; un salto de línea
//      suave cuenta como 1 (U+2028).
//   3. Las piezas de cada línea (palabras, espacios, tabuladores) quedan con su x,
//      para puntoDe y posicionDe.
//
// Uso: const motor = MotorTexto.crear(ctx2d);
//      const m = motor.maquetar(texto, anchoPx, altoPx, pxPorEmu);   // { lineas, alto, ptPx }
//      motor.puntoDe(m, indiceParrafo, desplazamiento) → { x, arriba, base, alto, linea }
//      motor.posicionDe(m, x, y) → { parrafo, desplazamiento }

const MotorTexto = (() => {
  const EMU_POR_PT = 12700;
  const ALTO_LINEA = 1.2;   // [calibrar] igual que en pptx-texto.js
  const SIMBOLOS = {
    wingdings: { v: '❖', 'Ø': '➢', '§': '▪', 'ü': '✔', q: '❑', n: '■', l: '●', 'Ÿ': '•', 'à': '➔', 'è': '➜', 'ê': '➢', 'ð': '⇒', o: '□', p: '◻', 'ª': '✦', w: '⬥', u: '◆', s: '⬧', t: '⧫', 'ö': '⬧', 'û': '✗', 'ý': '☒', 'þ': '☑' },
    symbol: { '·': '•', 'Þ': '⇒', '®': '→', 'à': '◊', '¨': '♣', '©': '♦', 'ª': '♥', '«': '♠', '-': '−', 'Ö': '√' },
  };
  const FUENTES_SIMBOLO = '"Segoe UI Symbol", "Segoe UI Emoji", sans-serif';

  // --- Numeración (copia de pptx-texto.js) ---------------------------------------------
  function romano(n) {
    const t = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
    let s = ''; for (const [v, l] of t) while (n >= v) { s += l; n -= v; } return s;
  }
  function alfa(n) { let s = ''; while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); } return s; }
  function numerar(esquema, n) {
    const m = /^(arabic|romanUc|romanLc|alphaUc|alphaLc)(Period|ParenR|ParenBoth|Plain|Minus)?/.exec(esquema || '');
    let base = String(n);
    if (m) {
      if (m[1] === 'romanUc') base = romano(n); else if (m[1] === 'romanLc') base = romano(n).toLowerCase();
      else if (m[1] === 'alphaUc') base = alfa(n); else if (m[1] === 'alphaLc') base = alfa(n).toLowerCase();
      switch (m[2]) {
        case 'ParenR': return base + ')';
        case 'ParenBoth': return '(' + base + ')';
        case 'Plain': return base;
        case 'Minus': return '- ' + base + ' -';
        default: return base + '.';
      }
    }
    if (/^circleNum/.test(esquema) && n <= 20) return String.fromCharCode(0x2460 + n - 1);
    return n + '.';
  }
  function colorDeRelleno(r) {
    if (!r) return null;
    if (r.tipo === 'solido') return r.color;
    if (r.tipo === 'degradado') return r.paradas[0].color;
    if (r.tipo === 'trama') return r.fg;
    return null;
  }
  function ultimaNoEspacio(l) { for (let i = l.piezas.length - 1; i >= 0; i--) if (!l.piezas[i].espacio) return i; return -1; }
  function anchoSinEspacioFinal(l) {
    let w = l.ancho;
    for (let i = l.piezas.length - 1; i >= 0 && l.piezas[i].espacio; i--) w -= l.piezas[i].ancho;
    return w;
  }
  // Largo de un párrafo en desplazamientos: texto de los trozos; un salto suave, 1.
  const largoParrafo = (p) => p.trozos.reduce((n, t) => n + (t.salto ? 1 : (t.texto || '').length), 0);

  function crear(ctx) {
    let cache = new WeakMap();
    const invalidar = (p) => { if (p) cache.delete(p); else cache = new WeakMap(); };

    const medirCon = (t, s, px, ptPx) => {
      ctx.fontKerning = PptxTexto.kerningDe(t, px / ptPx);
      ctx.font = PptxTexto.fuenteCss(t, px);
      ctx.letterSpacing = t.espaciado ? `${t.espaciado * ptPx}px` : '0px';
      return ctx.measureText(s).width;
    };
    const metricas = (t, px) => {
      ctx.font = PptxTexto.fuenteCss(t, px);
      const m = ctx.measureText('Hg');
      return { asc: m.fontBoundingBoxAscent || px * 0.8, desc: m.fontBoundingBoxDescent || px * 0.2 };
    };

    // Corta un párrafo en líneas con y relativa al inicio del párrafo (sin el
    // espacio de antes). cfg: { cuerpo, ptPx, pxPorEmu, auto, escalar, sinAjuste, anchoCol }.
    function romperParrafo(p, cfg, textoVineta) {
      const { ptPx, pxPorEmu, auto, escalar, sinAjuste, anchoCol } = cfg;
      const medir = (t, s, px) => medirCon(t, s, px, ptPx);
      const visibles = p.trozos.filter((t) => !t.salto && t.texto);
      const tamDe = (t) => escalar(t.tamano) * (t.base ? 2 / 3 : 1);
      const primerTam = escalar((visibles[0] || p.fin).tamano);

      // Trozos → piezas (palabras, espacios, tabuladores) con su desplazamiento.
      const piezas = [];
      let off = 0;
      for (const t of p.trozos) {
        if (t.salto) { piezas.push({ salto: true, t, desde: off, hasta: off + 1 }); off += 1; continue; }
        const origen = t.texto || '';
        const px = tamDe(t) * ptPx * (t.mayusculas === 'small' ? 0.8 : 1);
        let pos = 0;
        for (const parte of origen.split(/(\t|[^\S\t]+)/)) {
          if (!parte) continue;
          const desde = off + pos; pos += parte.length;
          const hasta = off + pos;
          if (parte === '\t') { piezas.push({ tab: true, t, px, desde, hasta }); continue; }
          const s = t.mayusculas === 'all' || t.mayusculas === 'small' ? parte.toUpperCase() : parte;
          piezas.push({ texto: s, espacio: /^\s+$/.test(s), t, px, ancho: medir(t, s, px), desde, hasta });
        }
        off += origen.length;
      }

      const margen = p.margen * pxPorEmu, sangria = p.sangria * pxPorEmu;
      let vineta = null;
      let inicioPrimera = margen + sangria;
      if (textoVineta != null) {
        const base = visibles[0] || p.fin;
        const px = (p.vineta.puntos != null ? escalar(p.vineta.puntos) : escalar(base.tamano) * (p.vineta.escala || 1)) * ptPx;
        let car = textoVineta, fam = PptxTexto.familia(p.vineta.fuente || base.fuente);
        const tabla = SIMBOLOS[(p.vineta.fuente || '').toLowerCase()];
        if (tabla && p.vineta.tipo === 'caracter') { car = tabla[car] || car; fam = FUENTES_SIMBOLO; }
        else if (/wingdings|symbol|webdings/i.test(p.vineta.fuente || '')) fam = FUENTES_SIMBOLO;
        const fuenteCss = `${p.vineta.tipo === 'numero' && base.negrita ? 'bold ' : ''}${px.toFixed(3)}px ${p.vineta.tipo === 'numero' ? PptxTexto.familia(base.fuente) : fam}`;
        ctx.font = fuenteCss; ctx.letterSpacing = '0px';
        const w = ctx.measureText(car).width;
        const x = margen + sangria;
        vineta = { texto: car, css: fuenteCss, color: PptxTexto.colorCss(p.vineta.color || colorDeRelleno(base.relleno)), x, px };
        inicioPrimera = sangria < 0 ? Math.max(margen, x + w) : x + w + px * 0.25;
      }

      // Corte de líneas.
      const lineasP = [];
      let actual = null;
      const nueva = (primera) => {
        actual = { piezas: [], inicio: primera ? inicioPrimera : margen, ancho: 0, final: false, vineta: primera ? vineta : null };
        actual.limite = sinAjuste ? Infinity : anchoCol - actual.inicio;
        lineasP.push(actual);
      };
      nueva(true);
      const proximoTab = (x) => {
        const abs = actual.inicio + x;
        for (const tb of p.tabs) if (tb.pos * pxPorEmu > abs + 0.5) return tb.pos * pxPorEmu - actual.inicio;
        const paso = (p.tabulacion || 914400) * pxPorEmu;
        return (Math.floor(abs / paso) + 1) * paso - actual.inicio;
      };
      const esPalabra = (pz) => pz && !pz.salto && !pz.tab && !pz.espacio;
      piezas.forEach((pz, i) => {
        if (pz.salto) { actual.salto = pz.t; actual.saltoHasta = pz.hasta; actual.final = true; nueva(false); return; }
        if (pz.tab) { const x = proximoTab(actual.ancho); actual.piezas.push({ ...pz, texto: '', ancho: x - actual.ancho }); actual.ancho = x; return; }
        if (pz.espacio) { actual.piezas.push(pz); actual.ancho += pz.ancho; return; }
        if (!esPalabra(piezas[i - 1])) {
          let palabra = 0;
          for (let k = i; esPalabra(piezas[k]); k++) palabra += piezas[k].ancho;
          const ocupado = anchoSinEspacioFinal(actual);
          if (actual.piezas.length && ocupado > 0 && actual.ancho + palabra > actual.limite + 0.01) nueva(false);
        }
        if (pz.ancho > actual.limite && !sinAjuste) {
          let resto = pz.texto, consumido = 0;
          while (resto) {
            let n = resto.length;
            // el mayor n que entra (lo mismo que bajar de uno en uno desde el final, como el
            // original, pero sin medir miles de prefijos: una palabra larga se corta por línea
            // en O(log n) medidas y no en O(n), que con 3.000 letras eran segundos por tecla)
            const entra = (k) => actual.ancho + medir(pz.t, resto.slice(0, k), pz.px) <= actual.limite;
            if (n > 1 && !entra(n)) {
              let lo = 1, hi = n - 1;   // invariante: lo «entra» o es 1; todo > hi no entra
              while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (entra(mid)) lo = mid; else hi = mid - 1; }
              n = lo;
            }
            // no partir un par sustituto (emoji y otros fuera del plano básico)
            if (n < resto.length && /[\uD800-\uDBFF]/.test(resto[n - 1])) n = n > 1 ? n - 1 : 2;
            const trozo = resto.slice(0, n);
            const w = medir(pz.t, trozo, pz.px);
            // con mayúsculas el texto mostrado puede ser más largo que el original (ß → SS): los desplazamientos no pasan del de la pieza
            actual.piezas.push({ ...pz, texto: trozo, ancho: w, desde: Math.min(pz.hasta, pz.desde + consumido), hasta: Math.min(pz.hasta, pz.desde + consumido + n) });
            consumido += n;
            actual.ancho += w;
            resto = resto.slice(n);
            if (resto) nueva(false);
          }
          return;
        }
        actual.piezas.push(pz); actual.ancho += pz.ancho;
      });
      actual.final = true;

      // Desplazamientos de cada línea: continuos, cubren espacios y el salto.
      let cursor = 0;
      for (const l of lineasP) {
        l.desde = cursor;
        l.hasta = l.piezas.length ? l.piezas[l.piezas.length - 1].hasta : cursor;
        if (l.saltoHasta != null) l.hasta = Math.max(l.hasta, l.saltoHasta);
        cursor = l.hasta;
      }

      // Altura y posición vertical relativas al párrafo.
      let y = 0;
      lineasP.forEach((l, il) => {
        const conTexto = l.piezas.filter((pz) => !pz.espacio || l.piezas.length === 1);
        const fuentes = conTexto.length ? conTexto : [{ t: l.salto || p.fin, px: tamDe(l.salto || p.fin) * ptPx }];
        let tam = 0, asc = 0, desc = 0;
        for (const pz of fuentes) {
          const m = metricas(pz.t, pz.px);
          tam = Math.max(tam, pz.px); asc = Math.max(asc, m.asc); desc = Math.max(desc, m.desc);
        }
        if (l.vineta) tam = Math.max(tam, Math.min(l.vineta.px, tam || l.vineta.px));
        const unaLinea = tam * ALTO_LINEA;
        const bajoLaBase = unaLinea * (desc / (asc + desc || 1));
        const altoLinea = p.interlineado.pts != null ? p.interlineado.pts * ptPx : unaLinea * Math.max(0, p.interlineado.pct - auto.reduccion);
        l.alto = altoLinea;
        l.arriba0 = y;
        l.base0 = y + altoLinea - bajoLaBase;
        y += altoLinea;
        l.ultimaDelParrafo = il === lineasP.length - 1;
      });

      // Alineación horizontal (x de cada pieza y trozos para dibujar).
      for (const l of lineasP) {
        const alg = p.alineacion;
        const usado = anchoSinEspacioFinal(l);
        const libre = l.limite === Infinity ? 0 : l.limite - usado;
        let x = l.inicio;
        let extraPorEspacio = 0;
        if (l.limite !== Infinity) {
          if (alg === 'ctr') x += libre / 2;
          else if (alg === 'r') x += libre;
          else if ((alg === 'just' && !l.final) || alg === 'dist') {
            const espacios = l.piezas.slice(0, ultimaNoEspacio(l) + 1).filter((pz) => pz.espacio).length;
            if (espacios) extraPorEspacio = libre / espacios;
          }
        } else if (alg === 'ctr' || alg === 'r') {
          const libreCaja = anchoCol - l.inicio - usado;
          x += alg === 'ctr' ? libreCaja / 2 : libreCaja;
        }
        l.trozos = [];
        l.x0 = x;   // donde empieza la línea ya alineada: el cursor de una línea vacía (centrada o a la derecha) va ahí
        const ultima = ultimaNoEspacio(l);
        l.piezas.forEach((pz, i) => {
          pz.x = x;
          if (i <= ultima && !pz.espacio && pz.texto) {
            const t = pz.t;
            l.trozos.push({
              texto: pz.texto, x, ancho: pz.ancho, css: PptxTexto.fuenteCss(t, pz.px), px: pz.px, desde: pz.desde, hasta: pz.hasta,
              espaciado: t.espaciado ? t.espaciado * ptPx : 0, kerning: PptxTexto.kerningDe(t, pz.px / ptPx), color: PptxTexto.colorCss(colorDeRelleno(t.relleno)),
              subrayado: t.subrayado, tachado: t.tachado, resaltado: PptxTexto.colorCss(t.resaltado),
              desplazamiento: -t.base * escalar(t.tamano) * ptPx, sombra: t.sombra,
            });
          }
          x += pz.ancho + (pz.espacio && i <= ultima ? extraPorEspacio : 0);
        });
      }
      return { lineas: lineasP, alto: y, primerTam };
    }

    // texto: { cuerpo, parrafos } de pptx-modelo.js; ancho/alto: área de texto sin márgenes.
    // opciones.vinetaEn: el párrafo con el cursor; si está vacío, igual lleva su viñeta (lo que hace PowerPoint mientras se edita; lo guardado no la dibuja).
    function maquetar(texto, ancho, alto, pxPorEmu, opciones = {}) {
      const { cuerpo, parrafos } = texto;
      const ptPx = EMU_POR_PT * pxPorEmu;
      const auto = cuerpo.autoajuste.tipo === 'reducir' ? cuerpo.autoajuste : { escala: 1, reduccion: 0 };
      const escalar = (tam) => auto.escala === 1 ? tam : Math.max(1, Math.round(tam * auto.escala));
      const sinAjuste = cuerpo.ajuste === 'none';
      const columnas = Math.max(1, cuerpo.columnas);
      const espacioCol = cuerpo.espacioColumnas * pxPorEmu;
      const anchoCol = columnas > 1 ? (ancho - espacioCol * (columnas - 1)) / columnas : ancho;
      const cfg = { cuerpo, ptPx, pxPorEmu, auto, escalar, sinAjuste, anchoCol };
      const claveBase = [anchoCol, pxPorEmu, sinAjuste, auto.escala, auto.reduccion].join('|');
      const contadores = [];
      const lineas = [];
      let y = 0;

      parrafos.forEach((p, ip) => {
        const vacio = !p.trozos.some((t) => !t.salto && t.texto);
        contadores.length = Math.min(contadores.length, p.nivel + 1);
        let textoVineta = null;
        if (p.vineta && (!vacio || ip === opciones.vinetaEn)) {
          if (p.vineta.tipo === 'numero') {
            const c = contadores[p.nivel];
            const n = c && c.esquema === p.vineta.esquema ? c.n + 1 : p.vineta.inicio;
            contadores[p.nivel] = { esquema: p.vineta.esquema, n };
            textoVineta = numerar(p.vineta.esquema, n);
          } else if (p.vineta.tipo === 'caracter') textoVineta = p.vineta.caracter;
        } else if (!vacio) contadores[p.nivel] = null;

        const clave = claveBase + '|' + textoVineta;
        let r = cache.get(p);
        if (!r || r.clave !== clave) { r = { clave, ...romperParrafo(p, cfg, textoVineta) }; cache.set(p, r); }

        const espacio = (e) => e.pts != null ? e.pts * ptPx : e.pct * r.primerTam * ALTO_LINEA * ptPx;
        const antes = ip > 0 || cuerpo.primeroYUltimo ? espacio(p.antes) : 0;
        y += antes;
        const copia = r.lineas.map((l) => ({ ...l, parrafo: p, parrafoIndice: ip, arriba: y + l.arriba0, base: y + l.base0 }));
        y += r.alto;
        const despues = ip < parrafos.length - 1 || cuerpo.primeroYUltimo ? espacio(p.despues) : 0;
        y += despues;
        copia[0].antes = antes; copia[copia.length - 1].despues = despues;
        lineas.push(...copia);
      });

      // Columnas: las líneas pasan a la siguiente columna cuando no entran.
      if (columnas > 1) {
        let col = 0, desplazamiento = 0;
        for (const l of lineas) {
          if (l.base - desplazamiento > alto && col < columnas - 1 && l !== lineas[0]) { col++; desplazamiento = l.base - l.alto * 0.8; }
          l.base -= desplazamiento; l.arriba -= desplazamiento;
          const dx = col * (anchoCol + espacioCol);
          l.columna = col; l.x0 += dx;   // l es la copia de esta maqueta: lo guardado por párrafo no se toca
          l.piezas = l.piezas.map((pz) => ({ ...pz, x: pz.x + dx }));
          l.trozos = l.trozos.map((t) => ({ ...t, x: t.x + dx }));
          if (l.vineta) l.vineta = { ...l.vineta, x: l.vineta.x + dx };
        }
        y = Math.min(y, alto);
      }
      // columnas, anchoCol y espacioCol: para que posicionDe busque la línea en la columna del clic
      return { lineas, alto: y, ptPx, columnas, anchoCol, espacioCol };
    }

    // --- Consultas posición ↔ punto ----------------------------------------------------
    // x desde el borde del área de texto hasta el desplazamiento `d` de la pieza.
    function xEnPieza(pz, d, ptPx) {
      const k = Math.max(0, Math.min(d - pz.desde, pz.hasta - pz.desde));
      if (k <= 0 || pz.tab || !pz.texto) return pz.x + (k > 0 ? pz.ancho : 0);
      if (k >= pz.texto.length) return pz.x + pz.ancho;
      return pz.x + medirCon(pz.t, pz.texto.slice(0, k), pz.px, ptPx);
    }
    // Índice de la línea que contiene el desplazamiento `d` del párrafo `ip`. Un
    // desplazamiento en el corte entre dos líneas va al final de la primera si esa
    // línea termina en salto o es la última del párrafo; si no, al inicio de la
    // siguiente (como Word con el cursor tras un espacio de corte).
    function lineaDe(m, ip, d) {
      let primera = -1, ultima = -1;
      for (let i = 0; i < m.lineas.length; i++) if (m.lineas[i].parrafoIndice === ip) { if (primera < 0) primera = i; ultima = i; }
      if (primera < 0) return -1;
      for (let i = primera; i <= ultima; i++) {
        const l = m.lineas[i];
        if (d < l.hasta || i === ultima) return i;
      }
      return ultima;
    }
    // x del desplazamiento `d` dentro de la línea `i` (sin pasar a la siguiente: al final de una línea cortada da su borde derecho).
    function xEnLinea(m, i, d) {
      const l = m.lineas[i];
      if (!l.piezas.length) return l.x0 != null ? l.x0 : l.inicio;
      const dd = Math.max(l.desde, Math.min(d, l.hasta));
      const pz = l.piezas.find((q) => dd >= q.desde && dd < q.hasta);
      if (pz) return xEnPieza(pz, dd, m.ptPx);
      const u = l.piezas[l.piezas.length - 1];
      return u.x + u.ancho;
    }
    function puntoDe(m, ip, d) {
      const i = lineaDe(m, ip, d);
      if (i < 0) return null;
      const l = m.lineas[i];
      return { x: xEnLinea(m, i, d), arriba: l.arriba, base: l.base, alto: l.alto, linea: i };
    }
    // `linea` (opcional): la línea ya elegida (flechas arriba/abajo); si no, la de la y, en la columna de la x.
    function posicionDe(m, x, y, linea) {
      if (!m.lineas.length) return null;
      let candidatas = m.lineas;
      if (m.columnas > 1 && linea == null) {
        // de las líneas a esa altura, la más cerca en x (una línea que se sale de su columna, p. ej. por un tabulador, sigue siendo suya);
        // fuera de toda línea, las de la columna de la x
        const paso = m.anchoCol + m.espacioCol, ultimaCol = m.lineas[m.lineas.length - 1].columna || 0;
        // distancia en x a la línea; a igual distancia, gana la línea de la columna donde cae la x
        const lejos = (q) => {
          const a = (q.columna || 0) * paso, b = q.piezas.reduce((mx, pz) => Math.max(mx, pz.x + pz.ancho), a + m.anchoCol);
          const enSuColumna = x >= a - m.espacioCol / 2 && x < a + paso - m.espacioCol / 2;
          return (x < a ? a - x : x > b ? x - b : 0) + (enSuColumna ? 0 : 0.5);
        };
        const enY = m.lineas.filter((q) => y >= q.arriba && y < q.arriba + q.alto);
        if (enY.length) candidatas = [enY.reduce((mejor, q) => (lejos(q) < lejos(mejor) ? q : mejor))];
        else {
          const col = Math.max(0, Math.min(ultimaCol, Math.floor((x + m.espacioCol / 2) / paso)));
          candidatas = m.lineas.filter((q) => (q.columna || 0) === col);
        }
      }
      const l = linea != null && m.lineas[linea] ? m.lineas[linea] : candidatas.find((q) => y < q.arriba + q.alto) || candidatas[candidatas.length - 1];
      let d = l.desde;
      const piezas = l.piezas;
      if (piezas.length) {
        let hallado = false;
        for (const pz of piezas) {
          if (x >= pz.x + pz.ancho) { d = pz.hasta; continue; }
          if (x <= pz.x) { d = pz.desde; hallado = true; break; }
          // Dentro de la pieza: el límite de carácter más cercano.
          let mejor = pz.desde, dist = Infinity;
          for (let k = 0; k <= pz.hasta - pz.desde; k++) {
            if (pz.texto && k < pz.texto.length && /[\uD800-\uDBFF]/.test(pz.texto[k - 1] || '')) continue;   // no entre las dos mitades de un emoji
            const dx = Math.abs(xEnPieza(pz, pz.desde + k, m.ptPx) - x);
            if (dx < dist) { dist = dx; mejor = pz.desde + k; }
          }
          d = mejor; hallado = true; break;
        }
        // Pasado el final de una línea que sigue en otra: antes del espacio de corte.
        if (!hallado && !l.final) { const u = ultimaNoEspacio(l); if (u >= 0) d = l.piezas[u].hasta; }
        // Al final de una línea con salto suave, antes del salto.
        if (!hallado && l.saltoHasta != null) d = Math.min(d, l.saltoHasta - 1);
      }
      return { parrafo: l.parrafoIndice, desplazamiento: d };
    }
    return { maquetar, puntoDe, posicionDe, lineaDe, xEnLinea, invalidar };
  }

  return { crear, largoParrafo, numerar };
})();

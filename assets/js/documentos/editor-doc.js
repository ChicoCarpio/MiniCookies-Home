/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Editor de pantalla de Documentos (Fase 6, etapa 6.4): cursor, selección, teclado, ratón, dedo, IME y portapapeles sobre las hojas que dibuja
// el visor. El texto en sí lo cambia EdicionDoc (edicion-doc.js); aquí solo se traducen los gestos a órdenes y se vuelve a maquetar.
//
// El teclado, el IME y el pegado los recibe un <textarea> nativo que no se ve (junto al cursor): así el teclado del celular y la composición de
// japonés o chino funcionan solos. El cursor es un elemento del DOM que parpadea con CSS; la selección se dibuja sobre cada hoja, en su canvas.
//
//   const editor = EditorDoc.crear({ contenedor, visor, ed, mq, pintor, diseno, capa, alCambiar });
//   editor.seleccion() → { ancla, foco } · editor.poner(ancla, foco) · editor.relayout() · editor.destruir()
//   editor.posDePantalla(x, y) · editor.cajaDe(pos) · editor.area (el textarea)
//
// Requiere visor-doc.js, edicion-doc.js, maquetador-doc.js y motor-parrafo.js.
(function () {
  'use strict';
  const SUAVE = String.fromCharCode(0x2028);
  const CENTINELA = String.fromCharCode(0x200b);
  const MAX_AREA = 200000;   // una selección más larga no se copia al cuadro nativo (se copia desde el modelo)

  function crear(opc) {
    const { contenedor, visor, ed, mq } = opc;
    const doc = contenedor.ownerDocument;
    const motor = mq.motor;
    let diseno = opc.diseno;
    let ancla = null, foco = null;
    let xObj = null;                  // columna que mantienen las flechas ↑↓ (pt en la hoja)
    let lineaFija = null;             // { b, o }: el cursor se queda al final de una línea cortada (Fin)
    let enfocado = false, comp = null, vivo = true;
    let registros = null;             // bloque → [{ n, it }] las líneas de cada párrafo en las hojas
    let corriendo = null, otraVuelta = false;
    let hojasMarcadas = new Set();    // hojas donde se dibujó selección la última vez
    let temporizadorCapa = null;
    let lectura = false;              // modo lectura (celular): se mueve el cursor, se selecciona y se copia, pero no se escribe
    let versionCorr = 0;
    let corr = null;                  // { c (CorrectorDoc), idioma }: el corrector, si está encendido
    const erroresPorBloque = new WeakMap();   // bloque → { texto, idioma, lista }
    let busq = null;                  // { consulta, opc, lista: [{ b, o, n }] }: lo que se está buscando, resaltado en las hojas
    let pendiente = null;             // formato de caracteres que llevará lo próximo que se escriba (cursor sin selección: Ctrl+B y compañía)

    if (!doc.getElementById('vd-estilo-cursor')) {
      const st = doc.createElement('style');
      st.id = 'vd-estilo-cursor';
      st.textContent = '@keyframes vd-parpadeo{0%,49%{opacity:1}50%,100%{opacity:0}}';
      doc.head.appendChild(st);
    }
    // ───────── Cuadro nativo oculto ─────────
    const area = doc.createElement('textarea');
    area.className = 'vd-entrada';
    area.setAttribute('aria-label', opc.etiqueta || 'Texto del documento');
    area.setAttribute('autocapitalize', 'off'); area.setAttribute('autocomplete', 'off'); area.setAttribute('autocorrect', 'off'); area.setAttribute('spellcheck', 'false');
    area.style.cssText = 'position:absolute;left:0;top:0;width:2px;height:20px;padding:0;margin:0;border:0;outline:0;resize:none;overflow:hidden;opacity:0;pointer-events:none;background:transparent;color:transparent;caret-color:transparent;z-index:5;white-space:pre';
    contenedor.appendChild(area);
    const sincronizarArea = () => {
      if (comp) return;
      const [a, f] = ed.ordenar(ancla, foco);
      let v = CENTINELA + CENTINELA, desde = 1, hasta = 1;
      if (a.b !== f.b || a.o !== f.o) {
        const t = ed.texto(a, f);
        if (t.length < MAX_AREA) { v = t; desde = 0; hasta = t.length; }
      }
      if (area.value !== v) area.value = v;
      area.setSelectionRange(desde, hasta);
    };

    // ───────── Corrector ─────────
    const idiomaDe = (b) => {
      for (const z of b.piezas) { const i = z.r && z.r.idioma; if (i && /^(es|en)\b/i.test(i)) return i.slice(0, 2).toLowerCase(); }
      return corr.idioma;
    };
    // Las palabras mal escritas de un párrafo (se calculan una vez por bloque; al editarlo el bloque es otro)
    function erroresDe(b) {
      if (!corr) return [];
      const idioma = idiomaDe(b);
      if (!corr.c.listo(idioma)) return [];
      const k = erroresPorBloque.get(b);
      if (k && k.idioma === idioma && k.v === corr.v) return k.lista;
      const lista = corr.c.errores(b.texto.replace(/[\uFFFC\f]/g, ' '), idioma).map((e) => ({ ...e, idioma }));
      erroresPorBloque.set(b, { idioma, lista, v: corr.v });
      return lista;
    }
    const enEscritura = (b, e) => !!foco && foco.b === b && !!ancla && ancla.b === b && ancla.o === foco.o && foco.o >= e.o && foco.o <= e.o + e.n;
    // El idioma o el corrector cambió: se vuelve a calcular todo y se pinta
    function ponerCorrector(c, idioma) {
      corr = c ? { c, idioma: idioma || 'es', v: ++versionCorr } : null;
      if (corr && !corr.c.listo(corr.idioma)) corr.c.cargar(corr.idioma).then(() => { if (vivo) visor.repintar(); }, () => {});
      visor.repintar();
    }
    function errorEn(p) {
      if (!corr || !p) return null;
      return erroresDe(p.b).find((e) => p.o >= e.o && p.o <= e.o + e.n) ? { ...erroresDe(p.b).find((e) => p.o >= e.o && p.o <= e.o + e.n), b: p.b } : null;
    }
    // Cambia una palabra mal escrita por otra
    function reemplazarPalabra(e, texto) {
      const r = ed.reemplazar({ b: e.b, o: e.o }, { b: e.b, o: e.o + e.n }, texto, { nombre: 'Corregir' });
      if (r) listo(r);
      return r;
    }
    // ───────── Dónde está cada línea ─────────
    function indexar() {
      registros = new Map();
      diseno.paginas.forEach((h, n) => {
        for (const it of h.items) {
          if (it.t !== 'linea' || ed.indiceDe(it.b) < 0) continue;
          let a = registros.get(it.b);
          if (!a) registros.set(it.b, a = []);
          a.push({ n, it });
        }
      });
    }
    const valida = (p) => (p && ed.indiceDe(p.b) >= 0 ? p : null);
    // La línea (hoja y elemento) donde está la posición. Un desplazamiento justo en el corte entre dos líneas es el principio de la siguiente,
    // salvo si el cursor se fijó al final de la de arriba (tecla Fin).
    function ubicar(p) {
      const regs = registros && registros.get(p.b);
      if (!regs || !regs.length) return null;
      const fija = lineaFija && lineaFija.b === p.b && lineaFija.o === p.o;
      for (let i = 0; i < regs.length; i++) {
        const l = regs[i].it.l;
        if (p.o < l.hasta || (fija && p.o === l.hasta) || i === regs.length - 1) return regs[i];
      }
      return regs[regs.length - 1];
    }
    // Rectángulo (en pt, en la hoja) del cursor en una posición
    function cajaDe(p) {
      const reg = ubicar(p);
      if (!reg) return null;
      const it = reg.it;
      return { n: reg.n, x: it.x + motor.xEnLinea(it.l, p.o), y: it.y, alto: it.l.alto, it };
    }
    function posEnHoja(n, px, py) {
      const h = diseno.paginas[n];
      let mejor = null, costo = Infinity;
      for (const it of h.items) {
        if (it.t !== 'linea' || ed.indiceDe(it.b) < 0) continue;
        const l = it.l, ancho = it.r.anchoDisp || 0;
        const dy = py < it.y ? it.y - py : py > it.y + l.alto ? py - it.y - l.alto : 0;
        const dx = px < it.x ? it.x - px : px > it.x + ancho ? px - it.x - ancho : 0;
        const c = dy * 1000 + dx;
        if (c < costo) { costo = c; mejor = it; }
      }
      if (!mejor) return null;
      return { b: mejor.b, o: motor.posicionDe(mejor.r, px - mejor.x, mejor.l.y + 0.5) };
    }
    // De un punto de la pantalla a una posición del texto
    function posDePantalla(cx, cy) {
      let mejor = -1, dist = Infinity, caja = null;
      visor.hojas.forEach((h, n) => {
        const r = h.getBoundingClientRect();
        const d = cy < r.top ? r.top - cy : cy > r.bottom ? cy - r.bottom : 0;
        if (d < dist) { dist = d; mejor = n; caja = r; }
      });
      if (mejor < 0) return null;
      const k = visor.escala;
      return posEnHoja(mejor, (cx - caja.left) / k, (cy - caja.top) / k);
    }
    // Dónde cae un punto de la hoja en la pantalla
    function aPantalla(n, x, y) {
      const r = visor.hojas[n].getBoundingClientRect(), k = visor.escala;
      return { x: r.left + x * k, y: r.top + y * k };
    }

    // ───────── Selección y cursor ─────────
    const cursor = doc.createElement('div');
    cursor.className = 'vd-cursor';
    cursor.setAttribute('aria-hidden', 'true');
    cursor.style.cssText = 'position:absolute;width:0;border-left:1.5px solid #000;pointer-events:none;z-index:4;display:none';
    function ponerCursor(visible = true) {
      cursor.remove();
      if (!ancla || !foco) return;
      const c = cajaDe(foco);
      if (!c) { cursor.style.display = 'none'; return; }
      const h = visor.hojas[c.n], k = visor.escala;
      h.appendChild(cursor);
      const [a, f] = ed.ordenar(ancla, foco);
      const colapsada = a.b === f.b && a.o === f.o;
      Object.assign(cursor.style, { display: colapsada && enfocado ? 'block' : 'none', left: c.x * k + 'px', top: c.y * k + 'px', height: c.alto * k + 'px' });
      // reiniciar la animación: el cursor se ve fijo mientras se escribe
      cursor.style.animation = 'none'; void cursor.offsetWidth; cursor.style.animation = 'vd-parpadeo 1.06s steps(1) infinite';
      // el cuadro nativo junto al cursor: ahí aparece la ventana del IME
      const cr = contenedor.getBoundingClientRect(), q = aPantalla(c.n, c.x, c.y);
      area.style.left = Math.max(0, q.x - cr.left + contenedor.scrollLeft) + 'px';
      area.style.top = Math.max(0, q.y - cr.top + contenedor.scrollTop) + 'px';
      area.style.height = Math.max(8, c.alto * k) + 'px';
      area.style.fontSize = Math.max(8, c.alto * k * 0.8) + 'px';
      if (visible) asegurarVisible(c);
    }
    function asegurarVisible(c) {
      const q = aPantalla(c.n, c.x, c.y), r = contenedor.getBoundingClientRect(), alto = c.alto * visor.escala;
      const margen = 24;
      if (q.y < r.top + margen) contenedor.scrollTop -= r.top + margen - q.y;
      else if (q.y + alto > r.bottom - margen) contenedor.scrollTop += q.y + alto - (r.bottom - margen);
      if (q.x < r.left + margen) contenedor.scrollLeft -= r.left + margen - q.x;
      else if (q.x > r.right - margen) contenedor.scrollLeft += q.x - (r.right - margen);
    }
    // La selección, dibujada encima del texto de cada hoja (la llama el visor tras pintar la hoja)
    function sobre(ctx, h, n) {
      if (!ancla || !foco) return;
      const [a, f] = ed.ordenar(ancla, foco);
      if (comp && comp.a) {
        const ia = ed.indiceDe(comp.a.b);
        ctx.fillStyle = '#000';
        for (const it of h.items) {
          if (it.t !== 'linea' || it.b !== comp.a.b || ia < 0) continue;
          const l = it.l, s0 = Math.max(comp.a.o, l.desde), s1 = Math.min(comp.b.o, l.hasta);
          if (s1 > s0) for (const [t0, t1] of motor.tramosEnLinea(l, s0, s1)) ctx.fillRect(it.x + t0, it.y + l.base - 0.5 + 2, t1 - t0, 0.8);
        }
      }
      if (corr) {
        // subrayado ondulado rojo bajo las palabras que no están en el diccionario; la que se está escribiendo (el cursor la toca) no se marca todavía
        ctx.strokeStyle = '#d32f2f'; ctx.lineWidth = 0.7;
        for (const it of h.items) {
          if (it.t !== 'linea' || ed.indiceDe(it.b) < 0) continue;
          for (const e of erroresDe(it.b)) {
            if (enEscritura(it.b, e)) continue;
            const l = it.l, s0 = Math.max(e.o, l.desde), s1 = Math.min(e.o + e.n, l.hasta);
            if (s1 <= s0) continue;
            for (const [t0, t1] of motor.tramosEnLinea(l, s0, s1)) {
              const y = it.y + l.base + 1.6;
              ctx.beginPath(); ctx.moveTo(it.x + t0, y);
              for (let x = it.x + t0, k = 0; x < it.x + t1; x += 1.5, k++) ctx.lineTo(Math.min(x + 1.5, it.x + t1), y + (k % 2 ? 0 : 1.2));
              ctx.stroke();
            }
          }
        }
      }
      if (busq) {
        ctx.fillStyle = 'rgba(255,214,0,0.4)';
        for (const m of busq.lista) {
          if (ed.indiceDe(m.b) < 0) continue;
          for (const it of h.items) {
            if (it.t !== 'linea' || it.b !== m.b) continue;
            const l = it.l, s0 = Math.max(m.o, l.desde), s1 = Math.min(m.o + m.n, l.hasta);
            if (s1 > s0) for (const [t0, t1] of motor.tramosEnLinea(l, s0, s1)) ctx.fillRect(it.x + t0, it.y, t1 - t0, l.alto);
          }
        }
      }
      if (a.b === f.b && a.o === f.o) return;
      const ia = ed.indiceDe(a.b), ifin = ed.indiceDe(f.b);
      ctx.fillStyle = enfocado ? 'rgba(51,144,255,0.38)' : 'rgba(140,140,140,0.35)';
      for (const it of h.items) {
        if (it.t !== 'linea') continue;
        const i = ed.indiceDe(it.b);
        if (i < ia || i > ifin || i < 0) continue;
        const l = it.l, d0 = i === ia ? a.o : 0, d1 = i === ifin ? f.o : Infinity;
        const s0 = Math.max(d0, l.desde), s1 = Math.min(d1, l.hasta);
        if (s1 < s0) continue;
        // la marca de párrafo (o el salto de línea) también se selecciona: un poco más de ancho al final de la línea
        const cubreFin = i < ifin && (l.final || l.salto);
        // con texto mezclado de izquierda a derecha y al revés, lo seleccionado en el texto son varios tramos en la hoja
        const tramos = motor.tramosEnLinea(l, s0, s1).map((t) => [it.x + t[0], it.x + t[1]]);
        if (cubreFin) {
          const rtl = !!it.b.p.derechaAIzquierda;
          if (!tramos.length) { const x = it.x + motor.xEnLinea(l, s0); tramos.push(rtl ? [x - 5, x] : [x, x + 5]); }
          else if (rtl) tramos[0][0] -= 5; else tramos[tramos.length - 1][1] += 5;
        }
        for (const [x0, x1] of tramos) if (x1 > x0) ctx.fillRect(x0, it.y, x1 - x0, l.alto);
      }
    }
    visor.ponerSobre(sobre);
    function hojasDeSeleccion() {
      const sal = new Set();
      if (!ancla || !foco) return sal;
      const [a, f] = ed.ordenar(ancla, foco);
      if (a.b === f.b && a.o === f.o && !comp) return sal;
      const ia = ed.indiceDe(a.b), ifin = ed.indiceDe(f.b);
      diseno.paginas.forEach((h, n) => {
        for (const it of h.items) { if (it.t !== 'linea') continue; const i = ed.indiceDe(it.b); if (i >= ia && i <= ifin && i >= 0) { sal.add(n); break; } }
      });
      return sal;
    }
    function repintarSeleccion() {
      const nuevas = hojasDeSeleccion();
      for (const n of new Set([...hojasMarcadas, ...nuevas])) visor.repintarHoja(n);
      hojasMarcadas = nuevas;
    }
    // Pone la selección: `a` es el ancla y `f` el extremo que se mueve. Todo cambio de selección pasa por aquí.
    function poner(a, f, opciones = {}) {
      ancla = valida(a) || ancla; foco = valida(f) || foco;
      pendiente = null;
      if (!opciones.conservarX) xObj = null;
      if (!opciones.conservarLinea) lineaFija = null;
      if (!opciones.sinCortar) ed.cortarFusion();
      sincronizarArea();
      repintarSeleccion();
      if (corr) { const c0 = cajaDe(foco); if (c0) visor.repintarHoja(c0.n); }
      ponerCursor();
      if (opc.alSeleccionar) opc.alSeleccionar(ancla, foco);
    }

    // ───────── Cambios al documento ─────────
    async function relayout() {
      if (corriendo) { otraVuelta = true; return corriendo; }
      corriendo = (async () => {
        try {
          do {
            otraVuelta = false;
            diseno = await mq.maquetar();
            if (opc.pintor) await opc.pintor.precargar(diseno);
          } while (otraVuelta && vivo);
          if (!vivo) return;
          indexar();
          if (busq) busq.lista = ed.buscar(busq.consulta, busq.opc);
          visor.actualizarDiseno(diseno);
          hojasMarcadas = hojasDeSeleccion();
          if (busq) visor.repintar(); else for (const n of hojasMarcadas) visor.repintarHoja(n);
          if (!valida(foco)) { const t = ed.parrafos(); if (t.length) { ancla = foco = { b: t[0], o: 0 }; } }
          ponerCursor();
          if (opc.alRelayout) opc.alRelayout(diseno);
          programarCapa();
        } finally { corriendo = null; }
      })();
      return corriendo;
    }
    function programarCapa() {
      if (!opc.capa) return;
      clearTimeout(temporizadorCapa);
      temporizadorCapa = setTimeout(() => { if (vivo) opc.capa.actualizar(); }, 350);
    }
    const alCambiarDoc = () => { if (opc.alCambiar) opc.alCambiar(); };
    // un maquetado que falla (documento enorme, archivo raro) se avisa en vez de quedar como rechazo sin atender con la hoja vieja a la vista
    const relayoutSeguro = () => relayout().catch((e) => { console.warn(e); if (opc.alError) opc.alError(e); });
    // Tras una orden de edición que devolvió la posición del cursor
    function listo(pos, extra) {
      if (!pos) return;
      ancla = foco = pos; xObj = null; lineaFija = null; pendiente = null;
      sincronizarArea();
      relayoutSeguro();
      alCambiarDoc();
      return extra;
    }
    function escribir(texto, fusionar = true) {
      const [a, f] = ed.ordenar(ancla, foco);
      const opciones = fusionar ? { fusionar: 'escribir' } : {};
      if (pendiente) opciones.formato = pendiente;
      listo(ed.reemplazar(a, f, texto, opciones));
    }
    function borrarSeleccion() {
      const [a, f] = ed.ordenar(ancla, foco);
      if (a.b === f.b && a.o === f.o) return false;
      listo(ed.reemplazar(a, f, '', { nombre: 'Borrar' }));
      return true;
    }
    const haySeleccion = () => { const [a, f] = ed.ordenar(ancla, foco); return !(a.b === f.b && a.o === f.o); };

    // ───────── Movimiento ─────────
    function inicioDeLinea(p) { const c = ubicar(p); return c ? { b: p.b, o: c.it.l.desde } : p; }
    function finDeLinea(p) {
      const c = ubicar(p);
      if (!c) return p;
      const l = c.it.l;
      let fin = l.hasta;
      if (l.salto) fin = l.hasta - 1;                      // detrás del salto de línea no: antes
      else if (!l.final) { for (let i = l.atomos.length - 1; i >= 0; i--) { const q = l.atomos[i]; if (q.k !== 'e' && !q.etiqueta && q.k !== 'f') { fin = q.hasta; break; } } }
      return { b: p.b, o: fin, fija: !l.final || !!l.salto };
    }
    function vertical(p, dir) {
      const c = ubicar(p);
      if (!c) return p;
      const k = c.it, xAbs = xObj != null ? xObj : k.x + motor.xEnLinea(k.l, p.o);
      xObj = xAbs;
      const hojas = diseno.paginas;
      const elegir = (n, ref, desdeBorde) => {
        let mejor = null, dyMin = Infinity, dxMin = Infinity;
        for (const it of hojas[n].items) {
          if (it.t !== 'linea' || ed.indiceDe(it.b) < 0 || it === k) continue;
          let dy;
          if (desdeBorde) dy = dir > 0 ? it.y : -(it.y + it.l.alto);
          else dy = dir > 0 ? it.y - (k.y + k.l.alto) : k.y - (it.y + it.l.alto);
          if (!desdeBorde && dy < -0.5) continue;
          const ancho = it.r.anchoDisp || 0;
          const dx = xAbs < it.x ? it.x - xAbs : xAbs > it.x + ancho ? xAbs - it.x - ancho : 0;
          if (dy < dyMin - 0.5 || (Math.abs(dy - dyMin) <= 0.5 && dx < dxMin)) { dyMin = Math.min(dy, dyMin); dxMin = dx; mejor = it; }
        }
        return mejor;
      };
      let destino = elegir(c.n, k, false);
      for (let n = c.n + dir; !destino && n >= 0 && n < hojas.length; n += dir) destino = elegir(n, k, true);
      if (!destino) { const t = ed.parrafos(); return dir > 0 ? { b: t[t.length - 1], o: t[t.length - 1].texto.length } : { b: t[0], o: 0 }; }
      return { b: destino.b, o: motor.posicionDe(destino.r, xAbs - destino.x, destino.l.y + 0.5) };
    }
    function porPantalla(p, dir) {
      const c = cajaDe(p);
      if (!c) return p;
      const q = aPantalla(c.n, c.x, c.y + c.alto / 2), alto = (contenedor.clientHeight || 600) * 0.9;
      xObj = xObj != null ? xObj : c.x;
      contenedor.scrollTop += dir * alto;
      const r = posDePantalla(q.x, q.y + dir * alto);
      return r || p;
    }
    const inicioDoc = () => { const t = ed.parrafos(); return { b: t[0], o: 0 }; };
    const finDoc = () => { const t = ed.parrafos(); return { b: t[t.length - 1], o: t[t.length - 1].texto.length }; };

    // ¿El texto junto al cursor se escribe de derecha a izquierda? Entonces ← avanza en el texto y → retrocede.
    function haciaLaDerecha(p) {
      const c = ubicar(p);
      if (!c) return !!p.b.p.derechaAIzquierda;
      const l = c.it.l, vis = l.atomos.filter((q) => !q.etiqueta && q.k !== 'f' && q.k !== 'b' && q.k !== 'pg');
      const a = vis.find((q) => p.o >= q.desde && p.o < q.hasta) || vis[vis.length - 1];
      return a ? (a.nivel & 1) === 1 : !!p.b.p.derechaAIzquierda;
    }
    function teclaDeMovimiento(ev) {
      const ctrl = ev.ctrlKey || ev.metaKey, sh = ev.shiftKey;
      let destino = null, vert = false, fija = false;
      const [a, f] = ed.ordenar(ancla, foco);
      switch (ev.key) {
        case 'ArrowLeft':
          if (!sh && haySeleccion()) destino = a;
          else if (haciaLaDerecha(foco)) destino = ctrl ? { b: foco.b, o: EdicionDoc.palabraDespues(foco.b.texto, foco.o) } : ed.siguiente(foco);
          else destino = ctrl ? { b: foco.b, o: EdicionDoc.palabraAntes(foco.b.texto, foco.o) } : ed.anterior(foco);
          if (ctrl && !(!sh && haySeleccion()) && destino.b === foco.b && destino.o === foco.o) destino = haciaLaDerecha(foco) ? ed.siguiente(foco) : ed.anterior(foco);
          break;
        case 'ArrowRight':
          if (!sh && haySeleccion()) destino = f;
          else if (haciaLaDerecha(foco)) destino = ctrl ? { b: foco.b, o: EdicionDoc.palabraAntes(foco.b.texto, foco.o) } : ed.anterior(foco);
          else destino = ctrl ? { b: foco.b, o: EdicionDoc.palabraDespues(foco.b.texto, foco.o) } : ed.siguiente(foco);
          if (ctrl && !(!sh && haySeleccion()) && destino.b === foco.b && destino.o === foco.o) destino = haciaLaDerecha(foco) ? ed.anterior(foco) : ed.siguiente(foco);
          break;
        case 'ArrowUp': destino = ctrl ? inicioDeParrafoAnterior(foco) : vertical(foco, -1); vert = true; break;
        case 'ArrowDown': destino = ctrl ? inicioDeParrafoSiguiente(foco) : vertical(foco, 1); vert = true; break;
        case 'Home': destino = ctrl ? inicioDoc() : inicioDeLinea(foco); break;
        case 'End': if (ctrl) destino = finDoc(); else { destino = finDeLinea(foco); fija = !!destino.fija; } break;
        case 'PageUp': destino = porPantalla(foco, -1); vert = true; break;
        case 'PageDown': destino = porPantalla(foco, 1); vert = true; break;
        default: return false;
      }
      ev.preventDefault();
      if (!destino) return true;
      const guardadaX = xObj;
      const pos = { b: destino.b, o: destino.o };
      if (fija) lineaFija = { b: pos.b, o: pos.o };
      poner(sh ? ancla : pos, pos, { conservarX: vert, conservarLinea: fija });
      if (!vert) xObj = null; else xObj = guardadaX != null ? guardadaX : xObj;
      return true;
    }
    function inicioDeParrafoAnterior(p) { const i = ed.indiceDe(p.b); return p.o > 0 ? { b: p.b, o: 0 } : { b: ed.parrafos()[Math.max(i - 1, 0)], o: 0 }; }
    function inicioDeParrafoSiguiente(p) { const t = ed.parrafos(), i = ed.indiceDe(p.b); return i < t.length - 1 ? { b: t[i + 1], o: 0 } : { b: p.b, o: p.b.texto.length }; }

    // ───────── Teclado ─────────
    area.addEventListener('keydown', (ev) => {
      if (ev.isComposing || ev.keyCode === 229 || comp) return;
      if (!ancla) return;
      const ctrl = ev.ctrlKey || ev.metaKey;
      if (teclaDeMovimiento(ev)) return;
      if (lectura) {
        const k = ev.key.toLowerCase();
        if (ctrl && !ev.altKey && k === 'a') { ev.preventDefault(); poner(inicioDoc(), finDoc()); }
        else if (!(ctrl && (k === 'c' || k === 'f'))) ev.preventDefault();
        return;
      }
      switch (ev.key) {
        case 'Backspace':
          ev.preventDefault();
          if (!borrarSeleccion()) listo(ed.borrarAtras(foco, ctrl));
          return;
        case 'Delete':
          ev.preventDefault();
          if (!borrarSeleccion()) listo(ed.borrarAdelante(foco, ctrl));
          return;
        case 'Enter':
          ev.preventDefault();
          if (ev.shiftKey) escribir(SUAVE, false);
          else if (ctrl) { const [a, f] = ed.ordenar(ancla, foco); listo(ed.saltoDePagina(a, f)); }
          else { const [a, f] = ed.ordenar(ancla, foco); listo(ed.entrar(a, f)); }
          return;
        case 'Tab': {
          // en una lista, Tab al principio del elemento baja un nivel y Mayús+Tab sube uno; en otro sitio, un tabulador
          const [a, f] = ed.ordenar(ancla, foco), enLista = !!a.b.numPedida, alInicio = enLista && (a.b !== f.b || foco.o === 0);
          // en una tabla, Tab pasa a la celda siguiente (en la última agrega una fila) y Mayús+Tab a la anterior
          if (ed.celdaDe(foco) && !(enLista && (alInicio || ev.shiftKey))) {
            ev.preventDefault();
            const dest = ed.celdaVecina(foco, ev.shiftKey ? -1 : 1);
            if (dest && dest.b) listo(dest);
            return;
          }
          if (ev.shiftKey) { if (enLista) { ev.preventDefault(); aplicar(ed.nivelDeLista(a, f, -1)); } return; }
          ev.preventDefault();
          if (alInicio) aplicar(ed.nivelDeLista(a, f, 1)); else escribir('\t', false);
          return;
        }
        case 'Escape':
          if (haySeleccion()) { ev.preventDefault(); poner(foco, foco); }
          return;
        default: break;
      }
      if (ctrl && !ev.altKey) {
        const k = ev.key.toLowerCase();
        if (k === 'a') { ev.preventDefault(); poner(inicioDoc(), finDoc()); }
        else if (k === 'z') { ev.preventDefault(); deshacerRehacer(!ev.shiftKey); }
        else if (k === 'y') { ev.preventDefault(); deshacerRehacer(false); }
        else if (k === 'b' || k === 'i' || k === 'u') { ev.preventDefault(); alternar(k); }
        else if (k === 'l' && ev.shiftKey) { ev.preventDefault(); ponerLista('vinetas'); }
        else if (k === 'l' || k === 'e' || k === 'r' || k === 'j') { ev.preventDefault(); alinear({ l: 'left', e: 'center', r: 'right', j: 'both' }[k]); }
        else if (k === ' ') { ev.preventDefault(); formato({ limpiar: true }); }
      }
    });
    // ───────── Formato ─────────
    // Lo que devolvió una orden de formato ({ desde, hasta } con los bloques nuevos) pasa a ser la selección, con el mismo sentido
    function aplicar(r) {
      if (!r) return false;
      const anclaPrimero = ed.comparar(ancla, foco) <= 0;
      ancla = anclaPrimero ? r.desde : r.hasta; foco = anclaPrimero ? r.hasta : r.desde;
      xObj = null; lineaFija = null;
      sincronizarArea();
      relayoutSeguro();
      alCambiarDoc();
      return true;
    }
    const rango = () => ed.ordenar(ancla, foco);
    const avisarFormato = () => { if (opc.alSeleccionar) opc.alSeleccionar(ancla, foco); };
    // Formato de caracteres: con selección se aplica ya; con el cursor solo, queda pendiente para lo que se escriba
    function formato(cambios) {
      if (!ancla) return false;
      if (haySeleccion()) { const [a, f] = rango(); return aplicar(ed.formatoCaracter(a, f, cambios)); }
      pendiente = Object.assign({}, pendiente, cambios);
      avisarFormato();
      return true;
    }
    function alternar(prop) {
      if (!ancla) return false;
      if (haySeleccion()) { const [a, f] = rango(); return aplicar(ed.alternarCaracter(a, f, prop)); }
      const poner = estado()[prop] !== true;
      return formato({ [prop]: prop === 'u' ? (poner ? 'single' : 'none') : poner });
    }
    // Lo que valen los caracteres seleccionados (o el que está junto al cursor), con el formato pendiente encima
    function estado() {
      const [a, f] = rango();
      const e = ed.estadoFormato(a, f) || {};
      if (pendiente) {
        const p = pendiente;
        if ('b' in p) e.b = !!p.b; if ('i' in p) e.i = !!p.i; if ('tachado' in p) e.tachado = !!p.tachado;
        if ('u' in p) e.u = !!p.u && p.u !== 'none';
        if ('color' in p) e.color = /^[0-9a-f]{6}$/i.test(p.color || '') ? '#' + p.color.toLowerCase() : null;
        if ('tam' in p) e.tam = p.tam; if ('fuente' in p) e.fuente = p.fuente; if ('resaltado' in p) e.resaltado = p.resaltado === 'none' ? null : p.resaltado;
        if ('vertical' in p) e.vertical = p.vertical === 'baseline' ? null : p.vertical;
      }
      return e;
    }
    const alinear = (v) => { const [a, f] = rango(); return aplicar(ed.alinear(a, f, v)); };
    const sangrar = (d) => { const [a, f] = rango(); return aplicar(ed.sangrar(a, f, d)); };
    const interlineado = (x) => { const [a, f] = rango(); return aplicar(ed.interlineado(a, f, x)); };
    const espaciado = (antes, despues) => { const [a, f] = rango(); return aplicar(ed.espaciado(a, f, antes, despues)); };
    // ───────── Buscar y reemplazar ─────────
    // Qué se busca: se resalta en todas las hojas. Devuelve cuántas hay. Sin consulta, se quita el resaltado.
    function buscar(consulta, opc = {}) {
      busq = consulta ? { consulta, opc, lista: ed.buscar(consulta, opc) } : null;
      visor.repintar();
      return busq ? busq.lista.length : 0;
    }
    const coincidencias = () => (busq ? busq.lista.length : 0);
    // Salta a la siguiente (dir = 1) o anterior (-1) desde el cursor, dando la vuelta al documento; la deja seleccionada
    // con `incluir`, una coincidencia que empieza donde empieza la selección también vale (al escribir en el cuadro, la búsqueda no se va corriendo)
    function irA(dir, incluir) {
      if (!busq || !busq.lista.length || !ancla) return null;
      const [a, f] = ed.ordenar(ancla, foco), L = busq.lista;
      let m = null;
      if (dir > 0) m = L.find((x) => ed.comparar({ b: x.b, o: x.o }, incluir ? a : f) >= 0);
      else for (let i = L.length - 1; i >= 0 && !m; i--) if (ed.comparar({ b: L[i].b, o: L[i].o }, a) < 0) m = L[i];
      if (!m) m = dir > 0 ? L[0] : L[L.length - 1];
      poner({ b: m.b, o: m.o }, { b: m.b, o: m.o + m.n });
      return m;
    }
    const esCoincidencia = () => {
      if (!busq || !ancla || !haySeleccion()) return null;
      const [a, f] = ed.ordenar(ancla, foco);
      return busq.lista.find((m) => m.b === a.b && a.b === f.b && m.o === a.o && m.o + m.n === f.o) || null;
    };
    // Cambia la coincidencia seleccionada (si lo seleccionado no lo es, primero salta a la siguiente) y pasa a la que sigue
    function reemplazarActual(nuevo) {
      if (!busq) return false;
      let m = esCoincidencia();
      if (!m) { irA(1); return false; }
      const c = ed.reemplazarUno(m, nuevo);
      if (!c) return false;
      ancla = foco = c;
      busq.lista = ed.buscar(busq.consulta, busq.opc);
      sincronizarArea(); relayoutSeguro(); alCambiarDoc();
      if (busq.lista.length) irA(1);
      return true;
    }
    function reemplazarTodo(nuevo) {
      if (!busq) return 0;
      const r = ed.reemplazarTodo(busq.consulta, nuevo, busq.opc);
      if (r.cursor) listo(r.cursor);
      return r.n;
    }
    // ───────── Tablas, imágenes y saltos ─────────
    // Una orden que devuelve dónde queda el cursor; con una selección, primero se borra
    function conCursor(fn) {
      if (!ancla) return null;
      if (haySeleccion()) borrarSeleccion();
      const r = fn(foco);
      if (r && r.b) listo(r);
      return r || null;
    }
    const insertarTabla = (filas, columnas) => conCursor((p) => ed.insertarTabla(p, filas, columnas));
    const enTabla = () => (foco ? ed.celdaDe(foco) : null);
    const ordenTabla = (nombre) => { const r = enTabla() ? ed[nombre](foco) : null; if (r && r.b) listo(r); return r; };
    const insertarCampo = (nombre) => conCursor((p) => ed.insertarCampo(p, nombre));
    const saltoDePagina = () => conCursor((p) => ed.saltoDePagina(p, p));
    const saltoDeSeccion = (tipo) => conCursor((p) => ed.saltoDeSeccion(p, tipo));
    async function insertarImagen(bytes, extension, ancho, alto, descripcion) {
      if (!ancla) return null;
      if (haySeleccion()) borrarSeleccion();
      const r = await ed.insertarImagen(foco, bytes, extension, ancho, alto, descripcion);
      if (r && r.b) listo(r);
      return r;
    }
    async function ponerLista(tipo) { const [a, f] = rango(); return aplicar(await ed.ponerLista(a, f, tipo)); }
    async function ponerEstilo(id) { const [a, f] = rango(); return aplicar(await ed.ponerEstilo(a, f, id)); }

    function deshacerRehacer(deshacer) {
      const p = deshacer ? ed.deshacer() : ed.rehacer();
      if (p) listo(p);
    }
    // Lo que escribe el teclado (también el del celular y las herramientas de accesibilidad) llega como «beforeinput»; se aplica al modelo y el
    // cuadro nativo no cambia nunca. La composición del IME es la excepción: se sigue con sus propios eventos.
    area.addEventListener('beforeinput', (ev) => {
      if (lectura) { ev.preventDefault(); return; }
      if (comp || ev.isComposing || !ancla) return;
      const t = ev.inputType, d = ev.data;
      let manejado = true;
      if (t === 'insertText' || t === 'insertReplacementText') escribir(d || (ev.dataTransfer ? ev.dataTransfer.getData('text/plain') : ''));
      else if (t === 'insertLineBreak') escribir(SUAVE, false);
      else if (t === 'insertParagraph') { const [a, f] = ed.ordenar(ancla, foco); listo(ed.entrar(a, f)); }
      else if (t === 'deleteContentBackward') { if (!borrarSeleccion()) listo(ed.borrarAtras(foco, false)); }
      else if (t === 'deleteContentForward') { if (!borrarSeleccion()) listo(ed.borrarAdelante(foco, false)); }
      else if (t === 'deleteWordBackward') { if (!borrarSeleccion()) listo(ed.borrarAtras(foco, true)); }
      else if (t === 'deleteWordForward') { if (!borrarSeleccion()) listo(ed.borrarAdelante(foco, true)); }
      else if (t === 'historyUndo') deshacerRehacer(true);
      else if (t === 'historyRedo') deshacerRehacer(false);
      else if (t === 'deleteByCut' || t === 'deleteContent') borrarSeleccion();
      else if (t === 'insertFromPaste' || t === 'insertFromDrop' || t === 'insertFromYank') manejado = false;
      else manejado = false;
      if (manejado) ev.preventDefault();
    });
    // lo que el cuadro nativo reciba por otra vía se descarta: el modelo manda
    area.addEventListener('input', () => { if (!comp) sincronizarArea(); });

    // ───────── IME ─────────
    area.addEventListener('compositionstart', () => {
      if (!ancla || lectura) return;
      const [a, f] = ed.ordenar(ancla, foco);
      comp = { a: { b: a.b, o: a.o }, b: { b: f.b, o: f.o }, vacia: true, formato: pendiente };
    });
    area.addEventListener('compositionupdate', (ev) => {
      if (!comp) return;
      reemplazarComposicion(ev.data || '');
    });
    area.addEventListener('compositionend', (ev) => {
      if (!comp) return;
      const final = ev.data == null ? null : ev.data;
      if (final !== null) reemplazarComposicion(final);
      comp = null; pendiente = null;
      ed.cortarFusion();
      sincronizarArea();
      ponerCursor();
      repintarSeleccion();
    });
    function reemplazarComposicion(texto) {
      const limpio = EdicionDoc.limpiar(texto).replace(/\n/g, ' ');
      const fin = ed.reemplazar(comp.a, comp.b, limpio, { fusionar: 'ime', nombre: 'Escribir', formato: comp.formato }) || comp.b;
      // la composición queda en un solo párrafo: el principio vuelve a ser el mismo desplazamiento, en el bloque nuevo
      comp = { a: { b: fin.b, o: comp.a.o }, b: { b: fin.b, o: fin.o }, formato: comp.formato };
      ancla = foco = { b: fin.b, o: fin.o };
      xObj = null; lineaFija = null;
      relayoutSeguro();
      alCambiarDoc();
    }

    // ───────── Portapapeles ─────────
    // Al copiar sale el texto y, si hay con qué, el HTML con el formato (Word, Google Docs y los navegadores lo entienden)
    function aPortapapeles(ev) {
      const [a, f] = ed.ordenar(ancla, foco), t = ed.texto(a, f);
      ev.clipboardData.setData('text/plain', t);
      if (globalThis.HtmlDoc && t.length < MAX_AREA) {
        try { ev.clipboardData.setData('text/html', HtmlDoc.escribir(ed.contenidoRico(a, f))); } catch (e) { /* queda el texto */ }
      }
      ev.preventDefault();
    }
    area.addEventListener('copy', (ev) => { if (haySeleccion()) aPortapapeles(ev); });
    area.addEventListener('cut', (ev) => {
      if (!haySeleccion()) return;
      aPortapapeles(ev);
      if (!lectura) borrarSeleccion();
    });
    // Pegar: con HTML se conserva el formato (negrita, color, títulos, listas…); sin él, o si no se entiende, va el texto
    async function pegarRico(rico) {
      const [a, f] = ed.ordenar(ancla, foco);
      const pos = await ed.pegarRico(a, f, rico);
      if (pos) listo(pos);
    }
    area.addEventListener('paste', (ev) => {
      ev.preventDefault();
      if (lectura) return;
      const cd = ev.clipboardData;
      const t = cd ? cd.getData('text/plain') : '', html = cd && globalThis.HtmlDoc ? cd.getData('text/html') : '';
      if (html) {
        let rico = [];
        try { rico = HtmlDoc.leer(html); } catch (e) { rico = []; }
        if (rico.length) { pegarRico(rico); return; }
      }
      if (t) escribir(t, false);
    });
    area.addEventListener('focus', () => { enfocado = true; sincronizarArea(); repintarSeleccion(); ponerCursor(false); });
    area.addEventListener('blur', () => { enfocado = false; repintarSeleccion(); ponerCursor(false); });

    // ───────── Ratón y dedo ─────────
    let arrastre = null, ultimo = { t: 0, x: 0, y: 0, n: 0 };
    const pila = visor.pila;
    pila.addEventListener('pointerdown', (ev) => {
      if (ev.button > 0 || !ev.target.closest || !ev.target.closest('.vd-hoja')) return;
      const p = posDePantalla(ev.clientX, ev.clientY);
      if (!p) return;
      const cerca = ev.timeStamp - ultimo.t < 450 && Math.abs(ev.clientX - ultimo.x) < 6 && Math.abs(ev.clientY - ultimo.y) < 6;
      ultimo = { t: ev.timeStamp, x: ev.clientX, y: ev.clientY, n: cerca ? ultimo.n + 1 : 1 };
      ev.preventDefault();
      area.focus({ preventScroll: true });
      const n = ultimo.n;
      if (n >= 3) { poner({ b: p.b, o: 0 }, { b: p.b, o: p.b.texto.length }); arrastre = null; }
      else if (n === 2) { const [x, y] = EdicionDoc.palabraEn(p.b.texto, p.o); poner({ b: p.b, o: x }, { b: p.b, o: y }); arrastre = null; }
      else {
        poner(ev.shiftKey ? ancla : p, p);
        arrastre = { ancla: { b: ancla.b, o: ancla.o } };
        try { pila.setPointerCapture(ev.pointerId); } catch (e) { /* sin captura */ }
      }
    });
    pila.addEventListener('pointermove', (ev) => {
      if (!arrastre) return;
      const p = posDePantalla(ev.clientX, ev.clientY);
      if (p) poner(arrastre.ancla, p, { sinCortar: true });
      // cerca del borde, la vista sigue al ratón
      const r = contenedor.getBoundingClientRect();
      if (ev.clientY < r.top + 20) contenedor.scrollTop -= 20; else if (ev.clientY > r.bottom - 20) contenedor.scrollTop += 20;
    });
    const soltar = () => { arrastre = null; };
    pila.addEventListener('pointerup', soltar); pila.addEventListener('pointercancel', soltar);

    // ───────── Arranque ─────────
    const hoja0 = ed.parrafos();
    if (hoja0.length) { ancla = foco = { b: hoja0[0], o: 0 }; }
    indexar();
    visor.repintar();
    ponerCursor(false);
    sincronizarArea();

    return {
      area, cursor, relayout, posDePantalla, cajaDe, ubicar, poner,
      seleccion: () => ({ ancla, foco }),
      diseno: () => diseno,
      enfocar: () => area.focus({ preventScroll: true }),
      escribir, pegarRico, deshacerRehacer, aPantalla,
      soloLectura(si) { lectura = !!si; area.inputMode = lectura ? 'none' : ''; area.setAttribute('aria-readonly', String(lectura)); },
      estaEnLectura: () => lectura,
      ponerCorrector, errorEn, reemplazarPalabra, hojaDe: (p) => { const c = cajaDe(p); return c && c.n; },
      buscar, coincidencias, irA, reemplazarActual, reemplazarTodo,
      insertarTabla, insertarCampo, enTabla, ordenTabla, saltoDePagina, saltoDeSeccion, insertarImagen,
      formato, alternar, estado, alinear, sangrar, interlineado, espaciado, ponerLista, ponerEstilo, pendiente: () => pendiente,
      composicion: () => comp && { a: { ...comp.a }, b: { ...comp.b } },
      destruir() { vivo = false; clearTimeout(temporizadorCapa); area.remove(); cursor.remove(); visor.ponerSobre(null); },
    };
  }

  globalThis.EditorDoc = { crear };
})();

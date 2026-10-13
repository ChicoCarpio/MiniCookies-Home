/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Dibujo de una hoja de Excel en un Canvas 2D, como la pinta Excel en pantalla al 100 % (96 dpi): cuadrícula, rellenos y
// tramas, bordes, texto (alineación, sangría, ajuste, desborde, rotación, reducir hasta ajustar) y celdas combinadas.
// Se calibra contra las imágenes que da Excel real (investigacion-motor-excel\arnes-xlsx\exportar-rango.ps1 + comparar-dibujo.mjs).
//
// Convención de píxeles (medida en Excel): la cuadrícula de cada celda es la línea del borde izquierdo y del superior, en el
// píxel x(c) / y(f); el relleno de una celda cubre x(c)…x(c+1) e y(f)…y(f+1) (los dos extremos incluidos: tapa sus dos líneas
// y la del vecino de la derecha y de abajo).
//
// Uso:
//   const g = XlsxGeometria.crear(libro, hoja, { medirDigito: XlsxDibujo.medirDigito });
//   XlsxDibujo.dibujarRango(ctx, libro, hoja, g, { f0, c0, f1, c1 }, { x: 0, y: 0, idioma: 'es', cuadricula: true });
// Requiere xlsx-modelo.js, xlsx-estilos.js, xlsx-formato.js, xlsx-geometria.js y pptx-texto.js (fuentes con las medidas de Office);
// para el formato condicional, también xlsx-expresion.js y xlsx-condicional.js, y xlsx-tablas-estilos.js y xlsx-tablas.js para los estilos de tabla
// (sin ellos se dibuja la hoja sin esas reglas).

(function () {
  const REJILLA = '#E0E0E0';
  const PAD_X = 2;            // píxeles entre la línea de la celda y el texto
  const COLORES_FORMATO = { Black: '#000000', Blue: '#0000FF', Cyan: '#00FFFF', Green: '#00FF00', Magenta: '#FF00FF', Red: '#FF0000', White: '#FFFFFF', Yellow: '#FFFF00' };

  // ---- Fuentes --------------------------------------------------------------------------------------

  // El nombre sale del archivo: sin comillas, barras ni caracteres de control, que romperían la cadena de la fuente (y el Canvas la ignoraría).
  const limpiarNombre = (n) => String(n || 'Calibri').replace(/["\\\u0000-\u001f\u007f]/g, '');
  const fuenteCss = (fu, tamPx) => `${fu.cursiva ? 'italic ' : ''}${fu.negrita ? 'bold ' : ''}${tamPx}px ${PptxTexto.familia(limpiarNombre(fu.nombre))}`;
  // Excel pinta con GDI a un tamaño entero de píxeles: 11 pt → 15 px.
  const tamPx = (pt) => { const t = Number.isFinite(pt) && pt > 0 ? Math.min(409.5, pt) : 11; return Math.max(1, CONFIG.pxExacto ? t * 96 / 72 : Math.round(t * 96 / 72)); }; // 409,5 pt: el máximo de Excel
  let canvasMedida = null;
  function ctxMedida() {
    if (!canvasMedida) canvasMedida = (typeof document !== 'undefined' ? document.createElement('canvas') : new OffscreenCanvas(8, 8)).getContext('2d'); // con los FontFace del documento
    return canvasMedida;
  }
  // Ancho del "0" con esa fuente (para el MDW de la geometría).
  function medirDigito(nombre, px) {
    const c = ctxMedida();
    c.font = fuenteCss({ nombre, negrita: false, cursiva: false }, px);
    return c.measureText('0').width;
  }
  // Carga las fuentes que usa el libro (hay que esperarla antes de medir: si no, el Canvas mide con otra sin avisar).
  function usosDeFuentes(libro) {
    const usos = [];
    const vistos = new Set();
    for (const f of libro.estilos.fuentes) {
      for (const [n, c] of [[false, false], [true, false], [false, true], [true, true]]) {
        const k = (f.nombre || '') + n + c;
        if (!vistos.has(k)) { vistos.add(k); usos.push({ fuente: f.nombre || 'Calibri', negrita: n, cursiva: c }); }
      }
    }
    for (const t of libro.textos || []) { // fuentes que solo usan tramos de texto con formato
      for (const run of t.ricos || []) {
        const nombre = run.fuente && run.fuente.nombre;
        if (!nombre) continue;
        for (const [n, c] of [[false, false], [true, false], [false, true], [true, true]]) {
          const k = nombre + n + c;
          if (!vistos.has(k)) { vistos.add(k); usos.push({ fuente: nombre, negrita: n, cursiva: c }); }
        }
      }
    }
    return usos;
  }
  async function cargarFuentes(libro) { await PptxTexto.cargarFuentes(usosDeFuentes(libro)); }

  // Excel (GDI) coloca cada letra en un píxel entero y no hace kerning; el Canvas usa avances fraccionarios. CONFIG.avance:
  // 0 = los del Canvas, 1 = redondeados, 2 = redondeados hacia arriba. (Se calibra contra Excel: comparar-dibujo.mjs --avance=N.)
  const CONFIG = { avance: 0, pxExacto: false };
  const memoAvance = new Map();
  let DIGITO = 0; // mientras se dibuja un número: ancho fijo del dígito (el MDW, como Excel)
  // Al imprimir (PDF) las unidades no son píxeles enteros: `fino` deja las posiciones sin redondear y `pxDe` da el tamaño de letra.
  let FINO = false, UIMP = 1, PXDE = null, TXT = { x: 0, y: 0 }, DXT = 0; // TXT: desplazamiento del texto según su alineación (x a la izquierda, xDer a la derecha, xCen al centro) y DXT el de la celda que se escribe: al imprimir a escala menor que 100 % el relleno de la celda no se achica en proporción
  const RD = (v) => (FINO ? v : Math.round(v));
  const pxLetra = (pt) => (PXDE ? PXDE(pt) : tamPx(pt));
  function avance(ctx, ch) {
    if (DIGITO && ch >= '0' && ch <= '9') return DIGITO;
    const k = ctx.font + '|' + ch;
    let a = memoAvance.get(k);
    if (a === undefined) {
      a = ctx.measureText(ch).width;
      if (CONFIG.avance === 1) a = Math.round(a); else if (CONFIG.avance === 2) a = Math.ceil(a - 0.2);
      memoAvance.set(k, a);
    }
    return a;
  }
  function anchoTexto(ctx, t) {
    if (!CONFIG.avance && !DIGITO) return ctx.measureText(t).width;
    let w = 0;
    for (const ch of t) w += avance(ctx, ch);
    return w;
  }
  function escribir(ctx, t, x, y) {
    x += DXT; y += TXT.y;
    if (!CONFIG.avance && !DIGITO) { ctx.fillText(t, x, y); return; }
    for (const ch of t) { ctx.fillText(ch, x, y); x += avance(ctx, ch); }
  }

  // ---- Colores y rellenos ----------------------------------------------------------------------------

  const colorDe = (libro, ref, porDefecto) => (ref ? libro.estilos.color(ref) || porDefecto : porDefecto);

  // Tramas de la norma (18.18.55 ST_PatternType): matrices 8×8 (1 = primer plano) medidas en Excel (comparar-dibujo.mjs, caso tramas-puras).
  // Excel las ancla a las coordenadas de la hoja (no al inicio de cada celda), con período 8.
  const TRAMAS = {
    gray125: ['10001000', '00000000', '00100010', '00000000', '10001000', '00000000', '00100010', '00000000'],
    gray0625: ['10000000', '00000000', '00001000', '00000000', '10000000', '00000000', '00001000', '00000000'],
    darkGray: ['11101110', '10111011', '11101110', '10111011', '11101110', '10111011', '11101110', '10111011'],
    mediumGray: ['10101010', '01010101', '10101010', '01010101', '10101010', '01010101', '10101010', '01010101'],
    lightGray: ['10001000', '00100010', '10001000', '00100010', '10001000', '00100010', '10001000', '00100010'],
    darkHorizontal: ['11111111', '11111111', '00000000', '00000000', '11111111', '11111111', '00000000', '00000000'],
    darkVertical: ['11001100', '11001100', '11001100', '11001100', '11001100', '11001100', '11001100', '11001100'],
    darkDown: ['11001100', '01100110', '00110011', '10011001', '11001100', '01100110', '00110011', '10011001'],
    darkUp: ['00110011', '01100110', '11001100', '10011001', '00110011', '01100110', '11001100', '10011001'],
    darkGrid: ['10011001', '01100110', '01100110', '10011001', '10011001', '01100110', '01100110', '10011001'],
    darkTrellis: ['11111111', '01100110', '11111111', '10011001', '11111111', '01100110', '11111111', '10011001'],
    lightHorizontal: ['11111111', '00000000', '00000000', '00000000', '11111111', '00000000', '00000000', '00000000'],
    lightVertical: ['10001000', '10001000', '10001000', '10001000', '10001000', '10001000', '10001000', '10001000'],
    lightDown: ['10001000', '01000100', '00100010', '00010001', '10001000', '01000100', '00100010', '00010001'],
    lightUp: ['00010001', '00100010', '01000100', '10001000', '00010001', '00100010', '01000100', '10001000'],
    lightGrid: ['11111111', '10001000', '10001000', '10001000', '11111111', '10001000', '10001000', '10001000'],
    lightTrellis: ['10001000', '01010101', '00100010', '01010101', '10001000', '01010101', '00100010', '01010101'],
  };
  // Distancia de la línea base al borde inferior de la fila, por tamaño de letra en píxeles (medida en Excel con Calibri, Arial, Times y Courier).
  const DESCENSO = { 19: 5, 21: 6, 24: 7, 27: 7, 32: 9, 37: 10, 48: 13 };
  const cachePatron = new Map();
  function patron(ctx, nombre, fg, bg) {
    const k = nombre + fg + bg;
    if (cachePatron.has(k)) return cachePatron.get(k);
    const m = TRAMAS[nombre];
    if (!m) return null;
    const c = typeof OffscreenCanvas !== 'undefined' ? new OffscreenCanvas(8, 8) : Object.assign(document.createElement('canvas'), { width: 8, height: 8 });
    const x = c.getContext('2d');
    x.fillStyle = bg; x.fillRect(0, 0, 8, 8);
    x.fillStyle = fg;
    for (let j = 0; j < 8; j++) for (let i = 0; i < 8; i++) if (m[j][i] === '1') x.fillRect(i, j, 1, 1);
    const p = ctx.createPattern(c, 'repeat');
    cachePatron.set(k, p);
    return p;
  }
  function estiloRelleno(ctx, libro, relleno, anc) {
    if (!relleno || relleno.tipo === 'ninguno') return null;
    if (relleno.tipo === 'degradado') {
      const p = relleno.degradado && relleno.degradado.paradas;
      return p && p.length ? colorDe(libro, p[0].color, null) : null;
    }
    const fg = colorDe(libro, relleno.fg, '#000000');
    if (relleno.patron === 'solid') return fg;
    const bg = colorDe(libro, relleno.bg, '#FFFFFF');
    const pt = patron(ctx, relleno.patron, fg, bg);
    if (!pt) return fg;
    if (pt.setTransform && anc) pt.setTransform(new DOMMatrix().translate(anc.tx, anc.ty));
    return pt;
  }

  // ---- Bordes ------------------------------------------------------------------------------------------

  // Medido en Excel (pantalla, 100 %): el píxel de la línea es el del borde de la celda; el grosor 2 ocupa [p-1, p], el 3 ocupa
  // [p-1, p+1] y el doble [p-1, p+1] con hueco en p. Los trazos discontinuos se anclan a la coordenada de la hoja (no al inicio de la línea):
  // patron = [raya, hueco, raya, hueco…] en píxeles; fase = desplazamiento del patrón.
  const BORDES = {
    thin: { o: [0] }, medium: { o: [-1, 0] }, thick: { o: [-1, 0, 1] }, double: { o: [-1, 1] }, hair: { o: [0], patron: [1, 1] },
    dashed: { o: [0], patron: [3, 1] }, dotted: { o: [0], patron: [2, 2], fase: 1 }, mediumDashed: { o: [-1, 0], patron: [9, 3] },
    dashDot: { o: [0], patron: [9, 3, 3, 3] }, mediumDashDot: { o: [-1, 0], patron: [9, 3, 3, 3] },
    dashDotDot: { o: [0], patron: [9, 3, 3, 3, 3, 3] }, mediumDashDotDot: { o: [-1, 0], patron: [9, 3, 3, 3, 3, 3] }, slantDashDot: { o: [-1, 0], patron: [9, 3, 3, 3] },
  };

  // (x0,y0)-(x1,y1): línea horizontal o vertical; ini = coordenada de la hoja donde empieza (para anclar el patrón)
  function trazo(ctx, x0, y0, x1, y1, borde, color, ini) {
    const def = BORDES[borde.estilo] || BORDES.thin;
    ctx.fillStyle = color;
    const horizontal = y0 === y1;
    const largo = horizontal ? x1 - x0 : y1 - y0;
    const rayas = def.patron || null;
    const total = rayas ? rayas.reduce((u, v) => u + v, 0) : 0;
    for (const o of def.o) {
      const fx = (d, h) => (horizontal ? ctx.fillRect(x0 + d, y0 + o, h - d, 1) : ctx.fillRect(x0 + o, y0 + d, 1, h - d));
      if (!rayas) { fx(0, largo); continue; }
      // recorre el patrón desde la fase que le toca al píxel de inicio
      let pos = (((ini + (def.fase || 0)) % total) + total) % total, i = 0;
      while (pos >= rayas[i]) { pos -= rayas[i]; i = (i + 1) % rayas.length; }
      let d = 0, resto = rayas[i] - pos;
      while (d < largo) {
        const l = Math.min(resto, largo - d);
        if (i % 2 === 0) fx(d, d + l);
        d += l; i = (i + 1) % rayas.length; resto = rayas[i];
      }
    }
  }

  // ---- Texto -------------------------------------------------------------------------------------------

  // Cuántas letras de t (al menos 1) caben en ancho px: búsqueda binaria sobre prefijos cortos (no se mide nunca el texto entero).
  function prefijo(ctx, t, ancho) {
    const n = t.length;
    const muestra = Math.min(n, 32);
    const media = Math.max(0.5, anchoTexto(ctx, t.slice(0, muestra)) / muestra);
    let hi = Math.min(n, Math.max(2, Math.ceil(ancho / media * 1.5) + 4));
    while (hi < n && anchoTexto(ctx, t.slice(0, hi)) <= ancho) hi = Math.min(n, hi * 2);
    let lo = 1;
    while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (anchoTexto(ctx, t.slice(0, mid)) <= ancho) lo = mid; else hi = mid - 1; }
    return lo;
  }

  // Rompe un texto en líneas que quepan en `ancho` px (por palabras; una palabra más larga que la línea se parte por letras).
  // Como máximo 5000 líneas (una celda de Excel no pasa de 32.767 letras).
  function romper(ctx, texto, ancho) {
    const salida = [];
    if (!(ancho >= 1)) return [String(texto)]; // sin ancho no hay dónde romper
    const MAX = 5000;
    for (const parrafo of String(texto).split(/\r\n|\n|\r/)) {
      if (salida.length >= MAX) break;
      if (parrafo === '') { salida.push(''); continue; }
      const palabras = parrafo.split(/(?<= )/);
      let linea = '';
      for (const p of palabras) {
        if (salida.length >= MAX) break;
        const prueba = linea + p;
        if (linea && anchoTexto(ctx, prueba.replace(/ +$/, '')) > ancho) { salida.push(linea.replace(/ +$/, '')); linea = p; } else linea = prueba;
        // una palabra más larga que la línea: se parte por letras
        while (linea.length > 1 && salida.length < MAX) {
          const k = prefijo(ctx, linea, ancho);
          if (k >= linea.length) break;
          salida.push(linea.slice(0, k)); linea = linea.slice(k);
        }
      }
      salida.push(linea.replace(/ +$/, ''));
    }
    return salida;
  }

  // Texto con formato por tramos (una sola línea). Las propiedades que el tramo no trae salen de la fuente de la celda.
  // Con dibujar = false solo mide. Devuelve el ancho.
  function tramos(ctx, libro, runs, fu, x, y, colorBase, dibujar) {
    let ax = x;
    for (const run of runs) {
      const f = Object.assign({}, fu, run.fuente);
      let tam = f.tam || 11;
      const sup = f.alinVertical === 'superscript', sub = f.alinVertical === 'subscript';
      if (sup || sub) tam = tam * 0.65;
      const px = pxLetra(tam);
      ctx.font = fuenteCss(f, px);
      const w = anchoTexto(ctx, run.texto);
      if (dibujar) {
        ctx.fillStyle = f.color ? colorDe(libro, f.color, colorBase) : colorBase;
        const dy = sup ? -Math.round(px * 0.55) : sub ? Math.round(px * 0.2) : 0;
        escribir(ctx, run.texto, RD(ax), y + dy);
        if (f.subrayado && f.subrayado !== 'none') ctx.fillRect(RD(ax), y + dy + 1, RD(w), 1);
        if (f.tachado) ctx.fillRect(RD(ax), y + dy - Math.round(px * 0.3), RD(w), 1);
      }
      ax += w;
    }
    return ax - x;
  }

  // ---- Dibujo de un rango -----------------------------------------------------------------------------

  // Los estilos de las tablas de una hoja, preparados una vez.
  const TABLAS = new WeakMap();
  function tablasDe(libro, hoja, op) {
    if (op.tablas === false || typeof XlsxTablas === 'undefined' || !(hoja.tablas && hoja.tablas.length)) return null;
    let t = TABLAS.get(hoja);
    if (t === undefined) { t = XlsxTablas.preparar(libro, hoja); TABLAS.set(hoja, t); }
    return t;
  }
  // Las reglas condicionales de una hoja, preparadas una vez (guardan estadísticas de sus rangos).
  const REGLAS = new WeakMap();
  const MINIS = new WeakMap();
  function minigraficosDe(libro, hoja, op) {
    if (op.minigraficos === false || typeof XlsxMinigraficos === 'undefined' || !(hoja.minigraficos && hoja.minigraficos.length)) return null;
    let m = MINIS.get(hoja);
    if (m === undefined) { m = XlsxMinigraficos.preparar(libro, hoja); MINIS.set(hoja, m); }
    return m;
  }
  function condicionalDe(libro, hoja, op) {
    if (op.condicional === false || typeof XlsxCondicional === 'undefined' || !(hoja.formatosCondicionales && hoja.formatosCondicionales.length)) return null;
    let c = REGLAS.get(hoja);
    if (c === undefined) { c = XlsxCondicional.preparar(libro, hoja, { hoy: op.hoy }); REGLAS.set(hoja, c); }
    return c;
  }

  function dibujarRango(ctx, libro, hoja, g, r, op) {
    op = op || {};
    FINO = !!op.fino; UIMP = op.fino ? (op.grosorRejilla || 1.28) : 1; PXDE = op.pxDe || null; TXT = op.desplazarTexto || { x: 0, y: 0 }; DXT = 0;
    const ox = op.x || 0, oy = op.y || 0;
    const X = (c) => ox + g.x(c) - g.x(r.c0), Y = (f) => oy + g.y(f) - g.y(r.f0);
    const C = hoja.celdas;
    const TIPO = LibroXlsx.TIPO;
    const cond = condicionalDe(libro, hoja, op);
    if (cond) cond.reiniciar();
    const tablas = tablasDe(libro, hoja, op);
    const anclaje = { tx: ((ox - g.x(r.c0)) % 8 + 8) % 8, ty: ((oy - g.y(r.f0)) % 8 + 8) % 8 };
    const filas = [], cols = [];
    for (let f = r.f0; f <= r.f1; f++) if (g.altoFila(f) > 0) filas.push(f);
    for (let c = r.c0; c <= r.c1; c++) if (g.anchoCol(c) > 0) cols.push(c);
    const anchoTotal = X(r.c1 + 1) - ox, altoTotal = Y(r.f1 + 1) - oy;
    ctx.save();
    try {
      ctx.fillStyle = '#FFFFFF';
      ctx.fillRect(ox, oy, anchoTotal + 1, altoTotal + 1);
      ctx.beginPath(); ctx.rect(ox, oy, anchoTotal + 1, altoTotal + 1); ctx.clip(); // nada se sale del rango (el texto y las combinadas pueden venir de fuera)

      // combinadas: celda superior izquierda → rectángulo; resto de celdas → la combinada que las cubre. Las repetidas o superpuestas
      // se saltan (un archivo hostil puede traer miles) y hay un tope de trabajo por cuadro.
      const combinada = new Map(), cubierta = new Set(), fuera = [];
      let presupuesto = 5000;
      for (const m of hoja.combinadas) {
        if (m.f1 < r.f0 || m.f0 > r.f1 || m.c1 < r.c0 || m.c0 > r.c1) continue;
        const tl = m.f0 * 16384 + m.c0;
        if (combinada.has(tl) || cubierta.has(tl)) continue;
        if (--presupuesto < 0) break;
        combinada.set(tl, m);
        // si su celda superior izquierda no se dibuja (fuera del rango o en una fila/columna oculta), la combinada entra "desde fuera"
        if (m.f0 < r.f0 || m.c0 < r.c0 || g.altoFila(m.f0) === 0 || g.anchoCol(m.c0) === 0) fuera.push(m);
        for (let f = Math.max(m.f0, r.f0); f <= Math.min(m.f1, r.f1); f++) for (let c = Math.max(m.c0, r.c0); c <= Math.min(m.c1, r.c1); c++) if (f !== m.f0 || c !== m.c0) cubierta.add(f * 16384 + c);
      }

      // 1. cuadrícula (no se ve dentro de una combinada)
      const rejilla = op.cuadricula !== undefined ? op.cuadricula : hoja.vista.cuadricula;
      if (rejilla) {
        const gr = op.grosorRejilla || 1;
        ctx.fillStyle = op.colorRejilla || REJILLA;
        for (const c of cols) ctx.fillRect(X(c), oy, gr, altoTotal + gr);
        ctx.fillRect(X(r.c1 + 1), oy, gr, altoTotal + gr);
        for (const f of filas) ctx.fillRect(ox, Y(f), anchoTotal + gr, gr);
        ctx.fillRect(ox, Y(r.f1 + 1), anchoTotal + gr, gr);
        ctx.fillStyle = '#FFFFFF';
        for (const m of combinada.values()) ctx.fillRect(X(m.c0) + gr, Y(m.f0) + gr, X(m.c1 + 1) - X(m.c0) - gr, Y(m.f1 + 1) - Y(m.f0) - gr); // sin tapar la línea de arriba ni la de la izquierda (impresa mide 0,96 pt, más de 1 u)
      }

      // estilo de cada celda del rango
      // el estilo de la celda, con el de su tabla debajo de su formato propio y el formato condicional encima de todo
      const aplicados = new WeakMap(), aplicadosT = new WeakMap();
      const estiloDe = (f, c) => {
        let xf = libro.estilos.xf(hoja.estiloDe(f, c));
        const t = tablas && tablas.para(f, c);
        if (t) {
          let rt = aplicadosT.get(t);
          if (!rt) { rt = XlsxTablas.aplicar(libro, xf, t); aplicadosT.set(t, rt); }
          xf = rt;
        }
        const e = cond && cond.para(f, c);
        if (!e) return xf;
        let r = aplicados.get(e);
        if (!r) { r = XlsxCondicional.aplicar(xf, e); aplicados.set(e, r); }
        return r;
      };

      // 2. rellenos (en orden de fila: el relleno de la fila de abajo cubre la línea que comparten)
      const rellenar = (f, c, m) => {
        const est = estiloRelleno(ctx, libro, estiloDe(f, c).relleno, anclaje);
        if (!est) return;
        ctx.fillStyle = est;
        const f1 = m ? m.f1 : f, c1 = m ? m.c1 : c;
        ctx.fillRect(X(c), Y(f), X(c1 + 1) - X(c) + 1, Y(f1 + 1) - Y(f) + 1);
      };
      for (const f of filas) {
        for (const c of cols) {
          const clave = f * 16384 + c;
          if (!cubierta.has(clave)) rellenar(f, c, combinada.get(clave));
        }
      }
      for (const m of fuera) rellenar(m.f0, m.c0, m);

      // 2b. barras de datos e iconos de las reglas condicionales (medido en el PDF de Excel: la barra va 2 u adentro de la celda a cada
      // lado y 2 u abajo del borde de arriba, con un degradado del color de la regla a su tinte hacia el blanco; el icono mide 16 u)
      if (cond) {
        for (const f of filas) {
          for (const c of cols) {
            const clave = f * 16384 + c;
            if (cubierta.has(clave)) continue;
            const e = cond.para(f, c);
            if (!e || (!e.barra && !e.icono)) continue;
            const m = combinada.get(clave);
            const bx0 = X(c), by0 = Y(f), bx1 = X((m ? m.c1 : c) + 1), by1 = Y((m ? m.f1 : f) + 1);
            if (e.barra) dibujarBarra(ctx, e.barra, bx0, by0, bx1, by1, oy);
            if (e.icono && typeof XlsxIconos !== 'undefined') {
              const lado = 16, alto = by1 - by0;
              XlsxIconos.dibujar(ctx, e.icono.conjunto, e.icono.indice, bx0 + RD(1.28), by0 + RD(Math.max(0, (alto - lado) * 0.76)), lado);
            }
          }
        }
      }

      // 2c. minigráficos (x14): ocupan la celda, sobre el relleno y las barras
      const mini = minigraficosDe(libro, hoja, op);
      if (mini) {
        for (const f of filas) {
          for (const c of cols) {
            const clave = f * 16384 + c;
            if (cubierta.has(clave)) continue;
            const sp = mini.para(f, c);
            if (!sp) continue;
            const m = combinada.get(clave);
            XlsxMinigraficos.dibujar(ctx, libro, sp, X(c), Y(f), X((m ? m.c1 : c) + 1), Y((m ? m.f1 : f) + 1), UIMP);
          }
        }
      }

      // 3. bordes
      const lado = (f, c, cual) => { if (f < 0 || c < 0 || f > 1048575 || c > 16383) return null; const b = estiloDe(f, c).borde; return b && b[cual] ? b[cual] : null; };
      // la línea que comparten dos celdas: gana el borde más pesado, luego el más oscuro y luego el propio de la celda (no el de un estilo de tabla)
      const mejor = (a, b) => XlsxEstilos.mejorBorde(b, a, (x) => libro.estilos.color(x)); // empate total: gana el de la derecha / de abajo (b)
      const bordes = (f, c, m) => {
        const f1 = m ? m.f1 : f, c1 = m ? m.c1 : c;
        const x0 = X(c), x1 = X(c1 + 1), y0 = Y(f), y1 = Y(f1 + 1);
        const izq = mejor(lado(f, c - 1, 'der'), lado(f, c, 'izq'));
        const arr = mejor(lado(f - 1, c, 'aba'), lado(f, c, 'arr'));
        const der = mejor(lado(f, c1, 'der'), lado(f, c1 + 1, 'izq'));
        const aba = mejor(lado(f1, c, 'aba'), lado(f1 + 1, c, 'arr'));
        const col = (b) => colorDe(libro, b.color, '#000000');
        if (izq) trazo(ctx, x0, y0, x0, y1 + 1, izq, col(izq), g.y(f));
        if (arr) trazo(ctx, x0, y0, x1 + 1, y0, arr, col(arr), g.x(c));
        if (der) trazo(ctx, x1, y0, x1, y1 + 1, der, col(der), g.y(f));
        if (aba) trazo(ctx, x0, y1, x1 + 1, y1, aba, col(aba), g.x(c));
        const d = estiloDe(f, c).borde;
        if (d && d.diagonal) {
          ctx.strokeStyle = col(d.diagonal); ctx.lineWidth = 1; ctx.beginPath();
          if (d.diagAbajo) { ctx.moveTo(x0, y0); ctx.lineTo(x1, y1); }
          if (d.diagArriba) { ctx.moveTo(x0, y1); ctx.lineTo(x1, y0); }
          ctx.stroke();
        }
      };
      for (const f of filas) {
        for (const c of cols) {
          const clave = f * 16384 + c;
          if (!cubierta.has(clave)) bordes(f, c, combinada.get(clave));
        }
      }
      for (const m of fuera) bordes(m.f0, m.c0, m);

      // 4. texto
      const idioma = op.idioma || 'es';
      const mdw = g.mdw;
      const hayCelda = (f, c) => { const i = C.indiceDe(f, c); return i >= 0 && C.tipo[i] !== TIPO.vacio && !(C.tipo[i] === TIPO.texto && C.valor[i] === ''); };
      const libre = (f, c) => !hayCelda(f, c) && !combinada.has(f * 16384 + c) && !cubierta.has(f * 16384 + c);
      // las celdas del rango (por el índice de filas, no recorriendo toda la hoja), la más cercana con contenido a cada lado (su texto
      // puede desbordar hacia el rango) y las combinadas que entran desde fuera
      C._indexar();
      const desborda = (q, haciaDonde) => { // ¿el texto de esta celda vecina puede entrar en el rango? (los números y los textos con ajuste no desbordan)
        if (C.tipo[q] === TIPO.numero || C.tipo[q] === TIPO.logico || C.tipo[q] === TIPO.error) return false;
        const al = libro.estilos.xf(C.estilo[q]).alineacion;
        if (al.ajustar) return false;
        const h = al.h === 'general' ? 'left' : al.h;
        return h === 'center' || h === 'centerContinuous' || h === (haciaDonde === 'der' ? 'right' : 'left') || h === 'distributed' || h === 'justify' || h === 'fill';
      };
      const visibles = [];
      for (const f of filas) {
        const ini = C._inicioFila.get(f);
        if (ini === undefined) continue;
        let lo = ini, hi = C.n; // primera celda de la fila con columna >= c0
        while (lo < hi) { const mid = (lo + hi) >> 1; if (C.fila[mid] < f || (C.fila[mid] === f && C.col[mid] < r.c0)) lo = mid + 1; else hi = mid; }
        let k = lo;
        for (; k < C.n && C.fila[k] === f && C.col[k] <= r.c1; k++) visibles.push(k);
        // la vecina de cada lado solo cuenta si su texto puede desbordar hacia el rango (la de la derecha, alineada a la derecha o al centro; la de la izquierda, a la izquierda o al centro)
        if (k < C.n && C.fila[k] === f) { for (let q = k; q < C.n && C.fila[q] === f && q < k + 8; q++) if (C.tipo[q] !== TIPO.vacio) { if (desborda(q, 'der')) visibles.push(q); break; } }
        for (let q = lo - 1; q >= ini && q > lo - 9; q--) if (C.tipo[q] !== TIPO.vacio) { if (desborda(q, 'izq')) visibles.push(q); break; }
      }
      for (const m of fuera) { const k = C.indiceDe(m.f0, m.c0); if (k >= 0) visibles.push(k); }
      for (const k of visibles) {
        const f = C.fila[k], c = C.col[k];
        if (C.tipo[k] === TIPO.vacio) continue;
        const clave = f * 16384 + c;
        if (cubierta.has(clave)) continue;
        const m = combinada.get(clave);
        if (!m && (g.altoFila(f) === 0 || g.anchoCol(c) === 0)) continue;
        const esNumero = C.tipo[k] === TIPO.numero;
        if (esNumero && C.valor[k] === 0 && hoja.vista.ceros === false) continue; // "mostrar ceros" desactivado
        const efecto = cond && cond.para(f, c);
        if (efecto && ((efecto.barra && !efecto.barra.mostrarValor) || (efecto.icono && !efecto.icono.mostrarValor))) continue; // "mostrar solo la barra / el icono"
        const xf = estiloDe(f, c);
        const f1 = m ? m.f1 : f, c1 = m ? m.c1 : c;
        const x0 = X(c), x1 = X(c1 + 1), y0 = Y(f), y1 = Y(f1 + 1); // bordes reales de la celda o la combinada (el recorte al rango ya está puesto)
        const anchoCelda = g.x(c1 + 1) - g.x(c);
        const fu = xf.fuente;
        let px = pxLetra(fu.tam || 11);
        const al = xf.alineacion;
        const anchoChars = Math.max(1, Math.floor((anchoCelda - 5) / mdw));
        const opcTexto = { idioma, decimal: op.decimal || (idioma === 'es' ? ',' : '.'), miles: op.miles || (idioma === 'es' ? ' ' : ','), ancho: anchoChars, anchoGeneral: Math.min(11, Math.max(1, anchoChars)) };
        if (efecto && efecto.formato) opcTexto.formatoCf = efecto.formato;
        let fm = LibroXlsx.textoCelda(libro, hoja, k, opcTexto);
        let relleno = null;
        if (fm.relleno) { relleno = fm.relleno; fm = LibroXlsx.textoCelda(libro, hoja, k, Object.assign({}, opcTexto, { ancho: 0 })); relleno = fm.relleno || relleno; } // "*x": se reparte por píxeles
        let texto = fm.texto;
        if (texto === '' || texto === undefined) continue;
        ctx.font = fuenteCss(fu, px);
        ctx.textBaseline = 'alphabetic';
        let color = fm.color && COLORES_FORMATO[fm.color] ? COLORES_FORMATO[fm.color] : colorDe(libro, fu.color, '#000000');
        if (fm.color && /^Color\d+$/.test(fm.color)) {
          const paleta = libro.estilos.paletaPropia && libro.estilos.paletaPropia.length ? libro.estilos.paletaPropia : XlsxEstilos.PALETA_INDEXADA;
          color = '#' + (paleta[+fm.color.slice(5) + 7] || '000000');
        }
        ctx.fillStyle = color;
        DIGITO = esNumero ? mdw : 0;
        const interior = anchoCelda - 2 * PAD_X - 1;
        if (al.reducir && !al.ajustar) { // reducir hasta ajustar: la letra baja de tamaño hasta que el texto cabe
          const w0 = anchoTexto(ctx, texto);
          if (w0 > interior && w0 > 0) {
            px = Math.max(4, Math.floor(px * interior / w0));
            ctx.font = fuenteCss(fu, px);
            for (let i = 0; i < 4 && px > 4 && anchoTexto(ctx, texto) > interior; i++) { px--; ctx.font = fuenteCss(fu, px); }
          }
        }
        let h = al.h;
        if (h === 'general') h = esNumero || C.tipo[k] === TIPO.logico || C.tipo[k] === TIPO.error ? (esNumero ? 'right' : 'center') : 'left';
        DXT = h === 'right' ? (TXT.xDer || 0) : h === 'center' || h === 'centerContinuous' ? (TXT.xCen || 0) : (TXT.x || 0);
        const sangria = Math.round((al.sangria || 0) * mdw * 1.2381);
        let lineas;
        if (al.ajustar && !esNumero) lineas = romper(ctx, texto, interior - (h === 'left' || h === 'right' ? sangria : 0));
        else lineas = [texto];
        if (relleno && !al.ajustar) { // antes + relleno repetido hasta llenar el ancho + después
          const ant = texto.slice(0, relleno.indice), des = texto.slice(relleno.indice + 1);
          const wf = anchoTexto(ctx, relleno.caracter) || 1;
          const libreAncho = interior - anchoTexto(ctx, ant) - anchoTexto(ctx, des);
          lineas = [ant + relleno.caracter.repeat(Math.max(1, Math.floor(libreAncho / wf))) + des];
          h = 'left';
        }
        if (h === 'fill' && !esNumero) { const w1 = anchoTexto(ctx, texto); lineas = [texto.repeat(Math.max(1, Math.floor(interior / Math.max(1, w1))))]; }
        // número que no cabe: ####
        if (esNumero && !al.ajustar && !relleno && anchoTexto(ctx, texto) > interior) {
          let n = 1;
          while (n < 255 && anchoTexto(ctx, '#'.repeat(n + 1)) <= interior) n++;
          lineas = ['#'.repeat(n)];
        }
        const asc = Math.round(px * 0.952), desc = DESCENSO[px] !== undefined ? DESCENSO[px] : px <= 16 ? 4 : Math.round(px * 0.269);
        const paso = Math.floor(g.altoTexto(fu.nombre || 'Calibri', fu.tam || 11, fu.negrita, fu.cursiva) * 4 / 3 + 1e-6);
        const altoTexto = (lineas.length - 1) * paso + asc + desc;
        const altoDisp = y1 - y0;
        let v = al.v || 'bottom';
        if (v === 'bottom' && altoTexto > altoDisp && !(FINO && lineas.length === 1)) v = 'top'; // si no cabe, Excel muestra las primeras líneas (impreso, una sola línea queda abajo aunque la fila sea más baja)
        let yBase;
        if (v === 'top') yBase = y0 + 2 + asc;
        else if (v === 'center') yBase = y0 + 1 + Math.floor((altoDisp - altoTexto) / 2) + asc;
        else yBase = y1 - desc - (lineas.length - 1) * paso;
        // texto girado: se alinea la caja que ocupa la línea ya girada (rotación 1-90 sube, 91-180 baja, 255 apila las letras)
        const rot = al.rotacion || 0;
        if (rot && rot !== 255) {
          const grados = rot <= 90 ? rot : -(rot - 90), th = grados * Math.PI / 180;
          const w = anchoTexto(ctx, texto), cs = Math.cos(th), sn = Math.sin(th);
          const esq = [[0, -asc], [w, -asc], [w, desc], [0, desc]].map(([qx, qy]) => [qx * cs + qy * sn, -qx * sn + qy * cs]);
          const mnx = Math.min(...esq.map((q) => q[0])), mxx = Math.max(...esq.map((q) => q[0])), mny = Math.min(...esq.map((q) => q[1])), mxy = Math.max(...esq.map((q) => q[1]));
          const Wp = mxx - mnx, Hp = mxy - mny;
          const bx = h === 'right' ? x1 - PAD_X - Wp : h === 'center' || h === 'centerContinuous' ? x0 + (x1 - x0 - Wp) / 2 : x0 + 1 + PAD_X;
          const by = v === 'top' ? y0 + 2 : v === 'center' ? y0 + (altoDisp - Hp) / 2 : y1 - Hp;
          ctx.save();
          ctx.beginPath(); ctx.rect(x0 + 1, y0 + 1, x1 - x0 - 1, y1 - y0 - 1); ctx.clip();
          ctx.translate(RD(bx - mnx), RD(by - mny)); ctx.rotate(-th);
          escribir(ctx, texto, 0, 0);
          ctx.restore();
          continue;
        }
        // sobre qué ancho se puede desbordar: hasta la próxima celda con contenido (o el borde del rango)
        const ancho = lineas.reduce((u, l) => Math.max(u, anchoTexto(ctx, l)), 0);
        if (!m && (c < r.c0 || c > r.c1)) { // vecina de fuera: solo se dibuja si su texto llega a entrar en el rango
          const ti = h === 'right' ? x1 - PAD_X - ancho : h === 'center' || h === 'centerContinuous' ? x0 + (x1 - x0 - ancho) / 2 : x0 + 1 + PAD_X + sangria;
          if (ti + ancho <= X(r.c0) || ti >= X(r.c1 + 1)) continue;
        }
        let izqLim = x0, derLim = x1;
        if (!al.ajustar && !esNumero && !m) {
          const haciaDerecha = (h === 'left' && ancho + PAD_X + sangria > anchoCelda) || (h === 'center' && ancho + 2 * PAD_X > anchoCelda);
          const haciaIzquierda = (h === 'right' && ancho + PAD_X > anchoCelda) || (h === 'center' && ancho + 2 * PAD_X > anchoCelda);
          if (haciaDerecha) { let cc = c1 + 1; while (cc <= r.c1 && cc - c1 <= 400 && libre(f, cc)) cc++; derLim = cc > r.c1 ? X(r.c1 + 1) + 1 : X(cc); }
          if (haciaIzquierda) { let cc = c - 1; while (cc >= r.c0 && c - cc <= 400 && libre(f, cc)) cc--; izqLim = cc < r.c0 ? X(r.c0) : X(cc + 1); }
        }
        ctx.save();
        ctx.beginPath();
        let recorte = [x0 + 1, y0 + 1, x1 - x0 - 1, y1 - y0 - 1];
        if (!al.ajustar && !esNumero) {
          if (h === 'left') recorte = [x0 + 1, y0 + 1, derLim - x0 - 1, y1 - y0 - 1];
          else if (h === 'right') recorte = [izqLim + 1, y0 + 1, x1 - izqLim - 1, y1 - y0 - 1];
          else if (h === 'center') recorte = [izqLim + 1, y0 + 1, derLim - izqLim - 1, y1 - y0 - 1];
        }
        ctx.rect(recorte[0], recorte[1], recorte[2], recorte[3]);
        ctx.clip();
        for (let i = 0; i < lineas.length; i++) {
          const l = lineas[i];
          const runs = lineas.length === 1 && C.ricos.get(k);
          const base = ctx.font;
          const w = runs ? tramos(ctx, libro, runs, fu, 0, 0, color, false) : anchoTexto(ctx, l);
          ctx.font = base;
          const y = yBase + i * paso;
          let x;
          if (h === 'distributed' && /\S\s+\S/.test(l.trim())) { // reparte el espacio sobrante entre las palabras
            const pal = l.trim().split(/\s+/), wp = pal.map((q) => anchoTexto(ctx, q)), hueco = (interior - wp.reduce((u, q) => u + q, 0)) / (pal.length - 1);
            let xx = x0 + 1 + PAD_X;
            pal.forEach((q, j) => { escribir(ctx, q, RD(xx), y); xx += wp[j] + hueco; });
            continue;
          }
          if (h === 'right') x = x1 - PAD_X - w - sangria;
          else if (h === 'center' || h === 'centerContinuous') x = x0 + (x1 - x0 - w) / 2 + (sangria ? sangria / 2 : 0);
          else x = x0 + 1 + PAD_X + sangria + Math.max(0, Math.round((px - 15) / 10));
          if (runs) { tramos(ctx, libro, runs, fu, RD(x), y, color, true); ctx.font = base; ctx.fillStyle = color; continue; }
          escribir(ctx, l, RD(x), y);
          if (fu.subrayado && fu.subrayado !== 'none') {
            const doble = /double/.test(fu.subrayado);
            ctx.fillRect(RD(x), y + 1, RD(w), 1);
            if (doble) ctx.fillRect(RD(x), y + 3, RD(w), 1);
          }
          if (fu.tachado) ctx.fillRect(RD(x), y - Math.round(px * 0.3), RD(w), 1);
        }
        ctx.restore();
      }
      // 5. objetos de la hoja (imágenes, formas, gráficos…): encima de todo, con el motor de PowerPoint
      if (op.objetos !== false && hoja.objetos && hoja.objetos.length && typeof XlsxObjetos !== 'undefined') XlsxObjetos.dibujar(ctx, hoja, g, r, { x: ox, y: oy, lienzo: op.lienzo });
    } finally {
      DIGITO = 0; FINO = false; PXDE = null; TXT = { x: 0, y: 0 }; DXT = 0;
      ctx.restore();
    }
    // 6. notas de celda: el triángulo rojo (pantalla) y el cuadro de las que se muestran siempre; sin el recorte del rango (al imprimir el cuadro se sale de la hoja)
    if (hoja.comentarios && hoja.comentarios.length && typeof XlsxNotas !== 'undefined') XlsxNotas.dibujar(ctx, libro, hoja, g, r, { x: ox, y: oy, fino: !!op.fino, grosorRejilla: op.grosorRejilla, notas: op.notas, indicadores: op.indicadores });
  }

  // Barra de datos con la extensión de Excel 2010 (medida en el PDF de Excel, 36-x14; u = un píxel de la impresora, 0,96 pt): la región va de 2 u a 1 u del borde de la celda,
  // 1 u más alta que la de 2007; sin eje las barras salen del lado de la dirección y miden (v − mín)/(máx − mín) de la región; con eje
  // (automático: en la proporción del cero; en el medio: al centro) la línea del eje mide 1 u, va a rayas (3 u y 1 u) y las barras positivas
  // y las negativas salen de sus lados, a escala del mayor valor absoluto del lado.
  function dibujarBarraX14(ctx, b, x0, y0, x1, y1, yBase) {
    const e = b.x14;
    const U = UIMP; // un píxel de la impresora (op.grosorRejilla, 0,96 pt); en pantalla, 1 px
    const xa = x0 + 2 * U, ancho = x1 - x0 - 3 * U, ya = y0 + 2 * U, alto = y1 - y0 - 4 * U;
    if (!(ancho > 1) || !(alto > 2)) return;
    const { lo, hi, valor: v } = b;
    const rtl = e.direccion === 'rightToLeft';
    const X = (t) => (rtl ? xa + ancho - t : xa + t);          // t: distancia desde el inicio de la región, en el sentido de la dirección
    const largoDe = (fr) => (e.largoMin + (e.largoMax - e.largoMin) * Math.min(1, Math.max(0, fr))) / 100;
    let ejeT = null, desde, hasta, negativa = false;
    if (e.eje === 'none' || (e.eje === 'automatic' && !(lo < 0))) {
      const fr = hi > lo ? (v - lo) / (hi - lo) : 0;
      desde = 0; hasta = ancho * largoDe(fr);
    } else {
      const cero = e.eje === 'middle' ? ancho / 2 : (hi <= 0 ? ancho : ancho * (-lo) / (hi - lo));
      ejeT = cero;
      const izq = Math.max(0, cero - U / 2), der = Math.max(0, ancho - cero);
      const mayor = e.eje === 'middle' ? Math.max(Math.abs(lo), Math.abs(hi)) : null;
      const mitad = ancho / 2; // la barra no pasa del borde de la región
      if (v >= 0) { const m = mayor !== null ? mayor : hi; const fr = m > 0 ? Math.min(1, v / m) : 0; desde = cero + U / 2; hasta = Math.min(ancho, desde + (mayor !== null ? mitad : der) * largoDe(fr)); }
      else { negativa = true; const m = mayor !== null ? mayor : -lo; const fr = m > 0 ? Math.min(1, -v / m) : 0; hasta = cero - U / 2; desde = Math.max(0, hasta - (mayor !== null ? mitad : izq) * largoDe(fr)); }
    }
    ctx.save();
    if (hasta - desde > 0.2) {
      const xi = Math.min(X(desde), X(hasta)), xd = Math.max(X(desde), X(hasta));
      const col = v < 0 ? e.colorNeg : b.color; // los negativos llevan su color también sin eje
      if (e.gradiente) {
        const c = [1, 3, 5].map((i) => parseInt(col.slice(i, i + 2), 16));
        const tinte = '#' + c.map((q) => (q + Math.trunc((255 - q) * 0.85)).toString(16).padStart(2, '0')).join('');
        // el color en el extremo que sale del eje (o del inicio) y el tinte en el extremo lejano
        const g = ctx.createLinearGradient(xi, 0, xd, 0);
        const alRevesX = (negativa !== rtl);
        g.addColorStop(alRevesX ? 1 : 0, col); g.addColorStop(alRevesX ? 0 : 1, tinte);
        ctx.fillStyle = g;
      } else ctx.fillStyle = col;
      ctx.fillRect(xi, ya, xd - xi, alto);
      const bc = v < 0 ? e.bordeNeg : e.borde;
      if (bc) {
        ctx.fillStyle = bc;
        ctx.fillRect(xi, ya, xd - xi, U); ctx.fillRect(xi, ya + alto - U, xd - xi, U);
        ctx.fillRect(xi, ya, U, alto); ctx.fillRect(xd - U, ya, U, alto);
      }
    }
    if (ejeT !== null) {
      ctx.strokeStyle = e.colorEje; ctx.lineWidth = U; ctx.setLineDash([3 * U, U]); ctx.lineDashOffset = (((y0 + U - yBase - 0.25 * U) % (4 * U)) + 4 * U) % (4 * U); // las rayas siguen una sola pauta desde la primera fila (medido)
      ctx.beginPath(); const xe = X(ejeT); ctx.moveTo(xe, y0 + U); ctx.lineTo(xe, y1 + U); ctx.stroke();
    }
    ctx.restore();
  }

  // Barra de datos (estilo 2007): del color de la regla a su tinte del 85 % hacia el blanco, de izquierda a derecha por la barra.
  function dibujarBarra(ctx, b, x0, y0, x1, y1, yBase) {
    if (b.x14) { dibujarBarraX14(ctx, b, x0, y0, x1, y1, yBase || 0); return; }
    const region = x1 - x0 - 4;
    const largo = region * (b.largoMin + (b.largoMax - b.largoMin) * b.fraccion) / 100;
    if (!(largo > 0.2) || !(y1 - y0 > 6)) return;
    const col = /^#[0-9A-F]{6}$/i.test(b.color) ? b.color : '#638EC6';
    const c = [1, 3, 5].map((i) => parseInt(col.slice(i, i + 2), 16));
    const tinte = '#' + c.map((v) => (v + Math.trunc((255 - v) * 0.85)).toString(16).padStart(2, '0')).join('');
    const g = ctx.createLinearGradient(x0 + 2, 0, x0 + 2 + largo, 0);
    g.addColorStop(0, col); g.addColorStop(1, tinte);
    ctx.fillStyle = g;
    ctx.fillRect(x0 + 2, y0 + 2, largo, y1 - y0 - 5);
  }

  // La geometría de la hoja con las dos medidas que necesitan el Canvas: el ancho del dígito (MDW) y las líneas que ocupa un texto con
  // ajuste (para el alto automático de las filas).
  function geometria(libro, hoja, opc) {
    const c = ctxMedida();
    const C = hoja.celdas;
    return XlsxGeometria.crear(libro, hoja, Object.assign({
      medirDigito,
      lineasDe(i, anchoCol) {
        const fu = libro.estilos.xf(C.estilo[i]).fuente;
        c.font = fuenteCss(fu, tamPx(fu.tam || 11));
        return romper(c, C.valor[i], anchoCol - 2 * PAD_X - 1).length;
      },
    }, opc || {}));
  }

  // ---- Vista con desplazamiento -------------------------------------------------------------------------

  // Dibuja lo que se ve de la hoja con su desplazamiento: encabezados de columna (A, B, C…) y de fila (1, 2, 3…) y solo las celdas
  // visibles (el costo no depende del tamaño de la hoja). v = { x, y (desplazamiento en px de hoja), ancho, alto (px del lienzo),
  // encabezados (true) }. Devuelve { f0, f1, c0, c1, areaX, areaY } (lo visible), útil para saber qué celda está bajo el puntero.
  function dibujarVista(ctx, libro, hoja, g, v, op) {
    op = op || {};
    const conEnc = v.encabezados !== false;
    const altoEnc = conEnc ? 20 : 0;
    const f0 = Math.max(0, g.filaEn(v.y)), c0 = Math.max(0, g.colEn(v.x));
    let f1 = f0, c1 = c0;
    while (f1 < 1048575 && g.y(f1 + 1) - v.y < v.alto - altoEnc) { if (f1 + 1 >= g.filasConocidas && g.altoPorDefecto === 0) break; f1++; }
    const ultima = Math.min(1048576, f1 + 1);
    const anchoEnc = conEnc ? Math.max(26, String(ultima).length * 8 + 12) : 0;
    while (c1 < 16383 && g.x(c1 + 1) - v.x < v.ancho - anchoEnc) { if (c1 + 1 >= g.columnasConocidas && g.anchoPorDefecto === 0) break; c1++; }
    ctx.save();
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, v.ancho, v.alto);
    ctx.beginPath(); ctx.rect(anchoEnc, altoEnc, v.ancho - anchoEnc, v.alto - altoEnc); ctx.clip();
    dibujarRango(ctx, libro, hoja, g, { f0, c0, f1, c1 }, Object.assign({}, op, { x: anchoEnc + g.x(c0) - v.x, y: altoEnc + g.y(f0) - v.y }));
    ctx.restore();
    if (conEnc) {
      ctx.save();
      ctx.fillStyle = '#E6E6E6'; ctx.fillRect(0, 0, v.ancho, altoEnc); ctx.fillRect(0, 0, anchoEnc, v.alto);
      ctx.fillStyle = '#BAB9B9';
      ctx.fillRect(0, altoEnc - 1, v.ancho, 1); ctx.fillRect(anchoEnc - 1, 0, 1, v.alto);
      ctx.font = '15px ' + PptxTexto.familia('Calibri'); ctx.textBaseline = 'alphabetic'; ctx.textAlign = 'center'; ctx.fillStyle = '#444444';
      for (let c = c0; c <= c1; c++) {
        const xa = anchoEnc + g.x(c) - v.x, w = g.anchoCol(c);
        if (w === 0) continue;
        ctx.fillRect(xa + w - 1, 0, 1, altoEnc);
        ctx.fillText(LibroXlsx.numeroACol(c), xa + w / 2, 15);
      }
      for (let f = f0; f <= f1; f++) {
        const ya = altoEnc + g.y(f) - v.y, h = g.altoFila(f);
        if (h === 0) continue;
        ctx.fillRect(0, ya + h - 1, anchoEnc, 1);
        ctx.fillText(String(f + 1), anchoEnc / 2, ya + Math.min(h, 20) - 5);
      }
      ctx.restore();
    }
    return { f0, f1, c0, c1, areaX: anchoEnc, areaY: altoEnc };
  }

  // Lee los objetos de la hoja y carga sus imágenes y fuentes (hay que esperarlo antes de dibujarla; sin él se dibuja sin objetos).
  // Lo que se carga de forma asíncrona antes de dibujar una hoja: los objetos (imágenes, formas, gráficos), las hojas que citan las fórmulas condicionales y las de los datos de los minigráficos.
  const prepararObjetos = async (libro, hoja, lienzo) => {
    if (typeof XlsxCondicional !== 'undefined' && hoja.formatosCondicionales && hoja.formatosCondicionales.length) await XlsxCondicional.cargar(libro, hoja);
    if (typeof XlsxMinigraficos !== 'undefined' && hoja.minigraficos && hoja.minigraficos.length) await XlsxMinigraficos.cargar(libro, hoja);
    return typeof XlsxObjetos === 'undefined' ? [] : XlsxObjetos.preparar(libro, hoja, lienzo);
  };

  globalThis.XlsxDibujo = { geometria, dibujarRango, prepararObjetos, dibujarVista, cargarFuentes, usosDeFuentes, medirDigito, tamPx, fuenteCss, anchoTexto, config: CONFIG };
})();

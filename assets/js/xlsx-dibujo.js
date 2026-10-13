/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Dibujo de una hoja de Excel en un Canvas 2D, como la pinta Excel en pantalla al 100 % (96 dpi): cuadrícula, rellenos y
// tramas, bordes, texto (alineación, sangría, ajuste, desborde, rotación, reducir hasta ajustar) y celdas combinadas.
// Se calibra contra las imágenes que da Excel real (comparar-xlsx\exportar-rango.ps1 + comparar-dibujo.mjs).
//
// Convención de píxeles (medida en Excel): la cuadrícula de cada celda es la línea del borde izquierdo y del superior, en el
// píxel x(c) / y(f); el relleno de una celda cubre x(c)…x(c+1) e y(f)…y(f+1) (los dos extremos incluidos: tapa sus dos líneas
// y la del vecino de la derecha y de abajo).
//
// Uso:
//   const g = XlsxGeometria.crear(libro, hoja, { medirDigito: XlsxDibujo.medirDigito });
//   XlsxDibujo.dibujarRango(ctx, libro, hoja, g, { f0, c0, f1, c1 }, { x: 0, y: 0, idioma: 'es', cuadricula: true });
// Requiere xlsx-modelo.js, xlsx-estilos.js, xlsx-formato.js, xlsx-geometria.js y pptx-texto.js (fuentes con las medidas de Office).

(function () {
  const REJILLA = '#E0E0E0';
  const PAD_X = 2;            // píxeles entre la línea de la celda y el texto
  const COLORES_FORMATO = { Black: '#000000', Blue: '#0000FF', Cyan: '#00FFFF', Green: '#00FF00', Magenta: '#FF00FF', Red: '#FF0000', White: '#FFFFFF', Yellow: '#FFFF00' };

  // ---- Fuentes --------------------------------------------------------------------------------------

  const fuenteCss = (fu, tamPx) => `${fu.cursiva ? 'italic ' : ''}${fu.negrita ? 'bold ' : ''}${tamPx}px ${PptxTexto.familia(fu.nombre)}`;
  // Excel pinta con GDI a un tamaño entero de píxeles: 11 pt → 15 px.
  const tamPx = (pt) => Math.max(1, CONFIG.pxExacto ? pt * 96 / 72 : Math.round(pt * 96 / 72));
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
  async function cargarFuentes(libro) {
    const usos = [];
    const vistos = new Set();
    for (const f of libro.estilos.fuentes) {
      for (const [n, c] of [[false, false], [true, false], [false, true], [true, true]]) {
        const k = (f.nombre || '') + n + c;
        if (!vistos.has(k)) { vistos.add(k); usos.push({ fuente: f.nombre || 'Calibri', negrita: n, cursiva: c }); }
      }
    }
    await PptxTexto.cargarFuentes(usos);
  }

  // Excel (GDI) coloca cada letra en un píxel entero y no hace kerning; el Canvas usa avances fraccionarios. CONFIG.avance:
  // 0 = los del Canvas, 1 = redondeados, 2 = redondeados hacia arriba. (Se calibra contra Excel: comparar-dibujo.mjs --avance=N.)
  const CONFIG = { avance: 0, pxExacto: false };
  const memoAvance = new Map();
  let DIGITO = 0; // mientras se dibuja un número: ancho fijo del dígito (el MDW, como Excel)
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
  const PESO = { hair: 1, dotted: 2, dashed: 3, thin: 4, dashDotDot: 5, dashDot: 6, mediumDashDotDot: 7, slantDashDot: 8, mediumDashDot: 9, mediumDashed: 10, medium: 11, double: 12, thick: 13 };

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

  // Rompe un texto en líneas que quepan en `ancho` px (por palabras; una palabra más larga que la línea se parte por letras).
  function romper(ctx, texto, ancho) {
    const salida = [];
    for (const parrafo of String(texto).split(/\r\n|\n|\r/)) {
      if (parrafo === '') { salida.push(''); continue; }
      const palabras = parrafo.split(/(?<= )/);
      let linea = '';
      for (const p of palabras) {
        const prueba = linea + p;
        if (linea && anchoTexto(ctx, prueba.replace(/ +$/, '')) > ancho) { salida.push(linea.replace(/ +$/, '')); linea = p; } else linea = prueba;
        while (anchoTexto(ctx, linea.replace(/ +$/, '')) > ancho && linea.length > 1) {
          let k = linea.length - 1;
          while (k > 1 && anchoTexto(ctx, linea.slice(0, k)) > ancho) k--;
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
      const px = tamPx(tam);
      ctx.font = fuenteCss(f, px);
      const w = anchoTexto(ctx, run.texto);
      if (dibujar) {
        ctx.fillStyle = f.color ? colorDe(libro, f.color, colorBase) : colorBase;
        const dy = sup ? -Math.round(px * 0.55) : sub ? Math.round(px * 0.2) : 0;
        escribir(ctx, run.texto, Math.round(ax), y + dy);
        if (f.subrayado && f.subrayado !== 'none') ctx.fillRect(Math.round(ax), y + dy + 1, Math.round(w), 1);
        if (f.tachado) ctx.fillRect(Math.round(ax), y + dy - Math.round(px * 0.3), Math.round(w), 1);
      }
      ax += w;
    }
    return ax - x;
  }

  // ---- Dibujo de un rango -----------------------------------------------------------------------------

  function dibujarRango(ctx, libro, hoja, g, r, op) {
    op = op || {};
    const ox = op.x || 0, oy = op.y || 0;
    const X = (c) => ox + g.x(c) - g.x(r.c0), Y = (f) => oy + g.y(f) - g.y(r.f0);
    const C = hoja.celdas;
    const anclaje = { tx: ((ox - g.x(r.c0)) % 8 + 8) % 8, ty: ((oy - g.y(r.f0)) % 8 + 8) % 8 };
    const filas = [], cols = [];
    for (let f = r.f0; f <= r.f1; f++) if (g.altoFila(f) > 0) filas.push(f);
    for (let c = r.c0; c <= r.c1; c++) if (g.anchoCol(c) > 0) cols.push(c);
    const anchoTotal = X(r.c1 + 1) - ox, altoTotal = Y(r.f1 + 1) - oy;
    ctx.save();
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(ox, oy, anchoTotal + 1, altoTotal + 1);

    // combinadas: celda superior izquierda → rectángulo; resto de celdas → la combinada que las cubre
    const combinada = new Map(), cubierta = new Set();
    for (const m of hoja.combinadas) {
      if (m.f1 < r.f0 || m.f0 > r.f1 || m.c1 < r.c0 || m.c0 > r.c1) continue;
      combinada.set(m.f0 * 16384 + m.c0, m);
      for (let f = Math.max(m.f0, r.f0); f <= Math.min(m.f1, r.f1); f++) for (let c = Math.max(m.c0, r.c0); c <= Math.min(m.c1, r.c1); c++) if (f !== m.f0 || c !== m.c0) cubierta.add(f * 16384 + c);
    }

    // 1. cuadrícula
    const rejilla = op.cuadricula !== undefined ? op.cuadricula : hoja.vista.cuadricula;
    if (rejilla) {
      ctx.fillStyle = REJILLA;
      for (const c of cols) ctx.fillRect(X(c), oy, 1, altoTotal + 1);
      ctx.fillRect(X(r.c1 + 1), oy, 1, altoTotal + 1);
      for (const f of filas) ctx.fillRect(ox, Y(f), anchoTotal + 1, 1);
      ctx.fillRect(ox, Y(r.f1 + 1), anchoTotal + 1, 1);
    }

    // estilo de cada celda del rango
    const estiloDe = (f, c) => libro.estilos.xf(hoja.estiloDe(f, c));

    // 2. rellenos (en orden de fila: el relleno de la fila de abajo cubre la línea que comparten)
    for (const f of filas) {
      for (const c of cols) {
        const clave = f * 16384 + c;
        if (cubierta.has(clave)) continue;
        const xf = estiloDe(f, c);
        const est = estiloRelleno(ctx, libro, xf.relleno, anclaje);
        if (!est) continue;
        const m = combinada.get(clave);
        const f1 = m ? Math.min(m.f1, r.f1) : f, c1 = m ? Math.min(m.c1, r.c1) : c;
        ctx.fillStyle = est;
        ctx.fillRect(X(c), Y(f), X(c1 + 1) - X(c) + 1, Y(f1 + 1) - Y(f) + 1);
      }
    }
    for (const m of combinada.values()) { // combinadas que empiezan antes del rango: su relleno entra desde fuera
      if (m.f0 >= r.f0 && m.c0 >= r.c0) continue;
      const est = estiloRelleno(ctx, libro, estiloDe(m.f0, m.c0).relleno, anclaje);
      if (!est) continue;
      ctx.fillStyle = est;
      ctx.fillRect(X(Math.max(m.c0, r.c0)), Y(Math.max(m.f0, r.f0)), X(Math.min(m.c1, r.c1) + 1) - X(Math.max(m.c0, r.c0)) + 1, Y(Math.min(m.f1, r.f1) + 1) - Y(Math.max(m.f0, r.f0)) + 1);
    }

    // 3. bordes
    const lado = (f, c, cual) => { const b = estiloDe(f, c).borde; return b && b[cual] ? b[cual] : null; };
    const mejor = (a, b) => (!a ? b : !b ? a : (PESO[b.estilo] || 0) >= (PESO[a.estilo] || 0) ? b : a);
    for (const f of filas) {
      for (const c of cols) {
        const clave = f * 16384 + c;
        if (cubierta.has(clave)) continue;
        const m = combinada.get(clave);
        const f1 = m ? m.f1 : f, c1 = m ? m.c1 : c;
        const x0 = X(c), x1 = X(Math.min(c1, r.c1) + 1), y0 = Y(f), y1 = Y(Math.min(f1, r.f1) + 1);
        const izq = mejor(c > 0 ? lado(f, c - 1, 'der') : null, lado(f, c, 'izq'));
        const arr = mejor(f > 0 ? lado(f - 1, c, 'aba') : null, lado(f, c, 'arr'));
        const der = mejor(lado(f, c1, 'der'), c1 < 16383 ? lado(f, c1 + 1, 'izq') : null);
        const aba = mejor(lado(f1, c, 'aba'), f1 < 1048575 ? lado(f1 + 1, c, 'arr') : null);
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
      }
    }

    // 4. texto
    const idioma = op.idioma || 'es';
    const mdw = g.mdw;
    const hayCelda = (f, c) => { const i = C.indiceDe(f, c); return i >= 0 && C.tipo[i] !== LibroXlsx.TIPO.vacio && !(C.tipo[i] === LibroXlsx.TIPO.texto && C.valor[i] === ''); };
    // solo las celdas del rango (por el índice de filas, no recorriendo toda la hoja) y las combinadas que empiezan antes del rango
    C._indexar();
    const visibles = [];
    for (const f of filas) {
      let k = C._inicioFila.get(f);
      if (k === undefined) continue;
      for (; k < C.n && C.fila[k] === f; k++) { if (C.col[k] < r.c0) continue; if (C.col[k] > r.c1) break; visibles.push(k); }
    }
    for (const m of combinada.values()) if (m.f0 < r.f0 || m.c0 < r.c0) { const k = C.indiceDe(m.f0, m.c0); if (k >= 0) visibles.push(k); }
    for (const k of visibles) {
      const f = C.fila[k], c = C.col[k];
      if (C.tipo[k] === LibroXlsx.TIPO.vacio) continue;
      if (g.altoFila(f) === 0 || g.anchoCol(c) === 0) continue;
      const clave = f * 16384 + c;
      if (cubierta.has(clave)) continue;
      const xf = libro.estilos.xf(C.estilo[k]);
      const m = combinada.get(clave);
      const f1 = m ? m.f1 : f, c1 = m ? m.c1 : c;
      const x0 = X(c), x1 = X(Math.min(c1, r.c1) + 1), y0 = Y(f), y1 = Y(Math.min(f1, r.f1) + 1);
      const anchoCelda = g.x(c1 + 1) - g.x(c);
      const fu = xf.fuente;
      let px = tamPx(fu.tam || 11);
      const al = xf.alineacion;
      const anchoChars = Math.max(1, Math.floor((anchoCelda - 5) / mdw));
      const fm = LibroXlsx.textoCelda(libro, hoja, k, { idioma, decimal: op.decimal || (idioma === 'es' ? ',' : '.'), miles: op.miles || (idioma === 'es' ? '\u00A0' : ','), ancho: anchoChars, anchoGeneral: Math.min(11, Math.max(1, anchoChars)) });
      let texto = fm.texto;
      if (texto === '' || texto === undefined) continue;
      const esNumero = C.tipo[k] === LibroXlsx.TIPO.numero;
      ctx.font = fuenteCss(fu, px);
      ctx.textBaseline = 'alphabetic';
      let color = fm.color && COLORES_FORMATO[fm.color] ? COLORES_FORMATO[fm.color] : colorDe(libro, fu.color, '#000000');
      if (fm.color && /^Color\d+$/.test(fm.color)) color = '#' + (XlsxEstilos.PALETA_INDEXADA[+fm.color.slice(5) + 7] || '000000');
      ctx.fillStyle = color;
      DIGITO = esNumero ? mdw : 0;
      const interior = anchoCelda - 2 * PAD_X - 1;
      if (al.reducir && !al.ajustar) { // reducir hasta ajustar: la letra baja de tamaño hasta que el texto cabe
        while (px > 4 && anchoTexto(ctx, texto) > interior) { px--; ctx.font = fuenteCss(fu, px); }
      }
      let h = al.h;
      if (h === 'general') h = esNumero || C.tipo[k] === LibroXlsx.TIPO.logico || C.tipo[k] === LibroXlsx.TIPO.error ? (C.tipo[k] === LibroXlsx.TIPO.numero ? 'right' : 'center') : 'left';
      const sangria = Math.round((al.sangria || 0) * mdw * 1.2381);
      let lineas;
      if (al.ajustar && !esNumero) lineas = romper(ctx, texto, interior - (h === 'left' || h === 'right' ? sangria : 0));
      else lineas = [texto];
      if (h === 'fill' && !esNumero) { const w1 = anchoTexto(ctx, texto); lineas = [texto.repeat(Math.max(1, Math.floor(interior / Math.max(1, w1))))]; }
      // número que no cabe: ####
      if (esNumero && !al.ajustar && anchoTexto(ctx, texto) > interior) {
        let n = 1;
        while (anchoTexto(ctx, '#'.repeat(n + 1)) <= interior) n++;
        lineas = ['#'.repeat(n)];
      }
      const altoLinea = Math.round(px * 1.2207) + 0;
      const asc = Math.round(px * 0.952), desc = DESCENSO[px] !== undefined ? DESCENSO[px] : px <= 16 ? 4 : Math.round(px * 0.269);
      const paso = Math.floor(g.altoTexto(fu.nombre || 'Calibri', fu.tam || 11, fu.negrita, fu.cursiva) * 4 / 3 + 1e-6);
      const altoTexto = (lineas.length - 1) * paso + asc + desc;
      const altoDisp = y1 - y0;
      let v = al.v || 'bottom';
      if (v === 'bottom' && altoTexto > altoDisp) v = 'top'; // si no cabe, Excel muestra las primeras líneas
      let yBase;
      if (v === 'top') yBase = y0 + 2 + asc;
      else if (v === 'center') yBase = y0 + 1 + Math.floor((altoDisp - altoTexto) / 2) + asc;
      else yBase = y1 - desc - (lineas.length - 1) * paso;
      // texto girado: se alinea la caja que ocupa la línea ya girada (rotación 1-90 sube, 91-180 baja, 255 apila las letras)
      const rot = al.rotacion || 0;
      if (rot && rot !== 255 && !esNumero || rot && rot !== 255 && esNumero) {
        const grados = rot <= 90 ? rot : -(rot - 90), th = grados * Math.PI / 180;
        const w = anchoTexto(ctx, texto), cs = Math.cos(th), sn = Math.sin(th);
        const esq = [[0, -asc], [w, -asc], [w, desc], [0, desc]].map(([qx, qy]) => [qx * cs + qy * sn, -qx * sn + qy * cs]);
        const mnx = Math.min(...esq.map((q) => q[0])), mxx = Math.max(...esq.map((q) => q[0])), mny = Math.min(...esq.map((q) => q[1])), mxy = Math.max(...esq.map((q) => q[1]));
        const Wp = mxx - mnx, Hp = mxy - mny;
        const bx = h === 'right' ? x1 - PAD_X - Wp : h === 'center' || h === 'centerContinuous' ? x0 + (x1 - x0 - Wp) / 2 : x0 + 1 + PAD_X;
        const by = v === 'top' ? y0 + 2 : v === 'center' ? y0 + (altoDisp - Hp) / 2 : y1 - Hp;
        ctx.save();
        ctx.beginPath(); ctx.rect(x0 + 1, y0 + 1, x1 - x0 - 1, y1 - y0 - 1); ctx.clip();
        ctx.translate(Math.round(bx - mnx), Math.round(by - mny)); ctx.rotate(-th);
        ctx.fillText(texto, 0, 0);
        ctx.restore();
        continue;
      }
      // sobre qué ancho se puede desbordar
      let izqLim = x0, derLim = x1;
      if (!al.ajustar && !esNumero && !m) {
        let cc = c1 + 1;
        while (cc <= r.c1 + 40 && !hayCelda(f, cc) && !combinada.has(f * 16384 + cc) && !cubierta.has(f * 16384 + cc) && cc - c1 < 200) { derLim = Math.max(derLim, X(Math.min(cc, r.c1 + 1) + 1)); if (cc > r.c1) break; cc++; }
        cc = c - 1;
        while (cc >= r.c0 && !hayCelda(f, cc) && !combinada.has(f * 16384 + cc) && !cubierta.has(f * 16384 + cc)) { izqLim = X(cc); cc--; }
      }
      ctx.save();
      ctx.beginPath();
      const ancho = Math.max(...lineas.map((l) => anchoTexto(ctx, l)));
      let recorte = [x0 + 1, y0 + 1, x1 - x0 - 1, y1 - y0 - 1];
      if (!al.ajustar && !esNumero) {
        if (h === 'left' && ancho + PAD_X + sangria > anchoCelda) recorte = [x0 + 1, y0 + 1, derLim - x0 - 1, y1 - y0 - 1];
        else if (h === 'right' && ancho + PAD_X > anchoCelda) recorte = [izqLim + 1, y0 + 1, x1 - izqLim - 1, y1 - y0 - 1];
        else if (h === 'center' && ancho + 2 * PAD_X > anchoCelda) recorte = [izqLim + 1, y0 + 1, derLim - izqLim - 1, y1 - y0 - 1];
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
          pal.forEach((q, j) => { escribir(ctx, q, Math.round(xx), y); xx += wp[j] + hueco; });
          continue;
        }
        if (h === 'right') x = x1 - PAD_X - w - sangria;
        else if (h === 'center' || h === 'centerContinuous') x = x0 + (x1 - x0 - w) / 2 + (sangria ? sangria / 2 : 0);
        else x = x0 + 1 + PAD_X + sangria + Math.max(0, Math.round((px - 15) / 10));
        if (runs) { tramos(ctx, libro, runs, fu, Math.round(x), y, color, true); ctx.font = base; ctx.fillStyle = color; continue; }
        escribir(ctx, l, Math.round(x), y);
        if (fu.subrayado && fu.subrayado !== 'none') {
          const doble = /double/.test(fu.subrayado);
          ctx.fillRect(Math.round(x), y + 1, Math.round(w), 1);
          if (doble) ctx.fillRect(Math.round(x), y + 3, Math.round(w), 1);
        }
        if (fu.tachado) ctx.fillRect(Math.round(x), y - Math.round(px * 0.3), Math.round(w), 1);
      }
      ctx.restore();
    }
    ctx.restore();
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
    while (f1 < 1048575 && g.y(f1 + 1) - v.y < v.alto - altoEnc) f1++;
    const ultima = Math.min(1048576, f1 + 1);
    const anchoEnc = conEnc ? Math.max(26, String(ultima).length * 8 + 12) : 0;
    while (c1 < 16383 && g.x(c1 + 1) - v.x < v.ancho - anchoEnc) c1++;
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

  globalThis.XlsxDibujo = { geometria, dibujarRango, dibujarVista, cargarFuentes, medirDigito, tamPx, config: CONFIG };
})();

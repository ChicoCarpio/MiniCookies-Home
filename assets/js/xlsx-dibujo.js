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
  const tamPx = (pt) => Math.max(1, Math.round(pt * 96 / 72));
  let canvasMedida = null;
  function ctxMedida() {
    if (!canvasMedida) canvasMedida = (typeof OffscreenCanvas !== 'undefined' ? new OffscreenCanvas(8, 8) : document.createElement('canvas')).getContext('2d');
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
  const CONFIG = { avance: 0 };
  const memoAvance = new Map();
  function avance(ctx, ch) {
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
    if (!CONFIG.avance) return ctx.measureText(t).width;
    let w = 0;
    for (const ch of t) w += avance(ctx, ch);
    return w;
  }
  function escribir(ctx, t, x, y) {
    if (!CONFIG.avance) { ctx.fillText(t, x, y); return; }
    for (const ch of t) { ctx.fillText(ch, x, y); x += avance(ctx, ch); }
  }

  // ---- Colores y rellenos ----------------------------------------------------------------------------

  const colorDe = (libro, ref, porDefecto) => (ref ? libro.estilos.color(ref) || porDefecto : porDefecto);

  // Tramas de la norma (18.18.55 ST_PatternType): matrices 8×8 (1 = primer plano). Aproximaciones calibradas contra Excel.
  const TRAMAS = {
    gray125: ['10001000', '00000000', '00100010', '00000000', '10001000', '00000000', '00100010', '00000000'],
    gray0625: ['10000000', '00000000', '00001000', '00000000', '10000000', '00000000', '00001000', '00000000'],
    lightGray: ['10001000', '00100010', '10001000', '00100010', '10001000', '00100010', '10001000', '00100010'],
    mediumGray: ['10101010', '01010101', '10101010', '01010101', '10101010', '01010101', '10101010', '01010101'],
    darkGray: ['01110111', '11011101', '01110111', '11011101', '01110111', '11011101', '01110111', '11011101'],
    darkHorizontal: ['11111111', '11111111', '00000000', '00000000', '11111111', '11111111', '00000000', '00000000'],
    darkVertical: ['11001100', '11001100', '11001100', '11001100', '11001100', '11001100', '11001100', '11001100'],
    darkDown: ['11001100', '01100110', '00110011', '10011001', '11001100', '01100110', '00110011', '10011001'],
    darkUp: ['00110011', '01100110', '11001100', '10011001', '00110011', '01100110', '11001100', '10011001'],
    darkGrid: ['11111111', '11001100', '11001100', '11111111', '11001100', '11001100', '11111111', '11001100'],
    darkTrellis: ['11001100', '11111111', '00110011', '11111111', '11001100', '11111111', '00110011', '11111111'],
    lightHorizontal: ['11111111', '00000000', '00000000', '00000000', '11111111', '00000000', '00000000', '00000000'],
    lightVertical: ['10001000', '10001000', '10001000', '10001000', '10001000', '10001000', '10001000', '10001000'],
    lightDown: ['10001000', '01000100', '00100010', '00010001', '10001000', '01000100', '00100010', '00010001'],
    lightUp: ['00010001', '00100010', '01000100', '10001000', '00010001', '00100010', '01000100', '10001000'],
    lightGrid: ['11111111', '10001000', '10001000', '10001000', '11111111', '10001000', '10001000', '10001000'],
    lightTrellis: ['10001000', '11111111', '00100010', '11111111', '10001000', '11111111', '00100010', '11111111'],
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
  function estiloRelleno(ctx, libro, relleno) {
    if (!relleno || relleno.tipo === 'ninguno') return null;
    if (relleno.tipo === 'degradado') {
      const p = relleno.degradado && relleno.degradado.paradas;
      return p && p.length ? colorDe(libro, p[0].color, null) : null;
    }
    const fg = colorDe(libro, relleno.fg, '#000000');
    if (relleno.patron === 'solid') return fg;
    const bg = colorDe(libro, relleno.bg, '#FFFFFF');
    return patron(ctx, relleno.patron, fg, bg) || fg;
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

  // ---- Dibujo de un rango -----------------------------------------------------------------------------

  function dibujarRango(ctx, libro, hoja, g, r, op) {
    op = op || {};
    const ox = op.x || 0, oy = op.y || 0;
    const X = (c) => ox + g.x(c) - g.x(r.c0), Y = (f) => oy + g.y(f) - g.y(r.f0);
    const C = hoja.celdas;
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
        const est = estiloRelleno(ctx, libro, xf.relleno);
        if (!est) continue;
        const m = combinada.get(clave);
        const f1 = m ? Math.min(m.f1, r.f1) : f, c1 = m ? Math.min(m.c1, r.c1) : c;
        ctx.fillStyle = est;
        ctx.fillRect(X(c), Y(f), X(c1 + 1) - X(c) + 1, Y(f1 + 1) - Y(f) + 1);
      }
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
    for (let k = 0; k < C.n; k++) {
      const f = C.fila[k], c = C.col[k];
      if (f < r.f0 || f > r.f1 || c < r.c0 || c > r.c1 || C.tipo[k] === LibroXlsx.TIPO.vacio) continue;
      if (g.altoFila(f) === 0 || g.anchoCol(c) === 0) continue;
      const clave = f * 16384 + c;
      if (cubierta.has(clave)) continue;
      const xf = libro.estilos.xf(C.estilo[k]);
      const m = combinada.get(clave);
      const f1 = m ? m.f1 : f, c1 = m ? m.c1 : c;
      const x0 = X(c), x1 = X(Math.min(c1, r.c1) + 1), y0 = Y(f), y1 = Y(Math.min(f1, r.f1) + 1);
      const anchoCelda = g.x(c1 + 1) - g.x(c);
      const fu = xf.fuente;
      const px = tamPx(fu.tam || 11);
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
      const interior = anchoCelda - 2 * PAD_X - 1;
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
      const altoTexto = lineas.length * (asc + desc);
      const altoDisp = y1 - y0;
      let v = al.v || 'bottom';
      if (v === 'bottom' && altoTexto > altoDisp) v = 'top'; // si no cabe, Excel muestra las primeras líneas
      let yBase;
      if (v === 'top') yBase = y0 + 2 + asc;
      else if (v === 'center') yBase = y0 + 1 + Math.floor((altoDisp - altoTexto) / 2) + asc;
      else yBase = y1 - desc - (lineas.length - 1) * (asc + desc);
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
        const w = anchoTexto(ctx, l);
        const y = yBase + i * (asc + desc);
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

  globalThis.XlsxDibujo = { dibujarRango, cargarFuentes, medirDigito, tamPx, config: CONFIG };
})();

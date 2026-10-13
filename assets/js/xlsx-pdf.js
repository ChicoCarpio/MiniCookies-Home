/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Hojas de Excel a PDF vectorial con las mismas páginas que imprime Excel (etapa 5 del motor de Excel): el corte lo calcula
// xlsx-paginas.js y cada página se dibuja con xlsx-dibujo.js sobre el lienzo de PDF (pdf-lienzo.js), así que el texto sigue siendo
// texto y las celdas, trazos.
//
// Unidades: la página se dibuja en "u", donde 1 u = 0,75·escala puntos (un píxel de pantalla a esa escala). Como en la impresora las
// columnas salen más anchas y las filas un poco más bajas que en pantalla en relación con la letra (MDW de 600 dpi), las
// columnas y filas de la página miden lo que mide Excel impreso (en u) y la letra se pide al tamaño de pantalla: así la
// proporción celda/letra es la de la impresión.
//
// Uso:
//   const { bytes, paginas } = await XlsxPdf.generar(libro, [0, 1], { idioma: 'es', nombreArchivo: 'Libro.xlsx' });
// Requiere pdf-lib, pdf-lienzo.js, pptx-texto.js, xlsx-modelo.js, xlsx-geometria.js, xlsx-paginas.js y xlsx-dibujo.js.

(function () {
  const PT_PX = 72 / 600;        // puntos por píxel de impresora
  const U_PT = 0.75;             // puntos por u a escala 100 %
  const ENC_ANCHO_PX = 173, ENC_ALTO_PX = 121; // encabezados de fila/columna impresos (píxeles de impresora, medidos)

  // ---- Encabezado y pie de página ---------------------------------------------------------------------
  // Códigos de ECMA-376 18.3.1.40: &L &C &R secciones, &P página, &N total, &D fecha, &T hora, &F archivo, &A hoja, &Z ruta,
  // &B negrita, &I cursiva, &U subrayado, &E doble subrayado, &S tachado, &X &Y superíndice/subíndice, &nn tamaño, &"fuente,estilo", &K color, && el signo.
  function analizarHF(texto) {
    const sec = { izq: [], cen: [], der: [] };
    let actual = 'cen';
    let est = { negrita: false, cursiva: false, subrayado: false, tachado: false, tam: null, fuente: null, color: null };
    let buf = '';
    const vaciar = () => { if (buf) { sec[actual].push(Object.assign({ texto: buf }, est)); buf = ''; } };
    const t = String(texto || '');
    for (let i = 0; i < t.length && i < 4096; i++) {
      const ch = t[i];
      if (ch !== '&') { buf += ch; continue; }
      const n = t[++i];
      if (n === undefined) break;
      if (n === '&') { buf += '&'; continue; }
      if (n === 'L' || n === 'C' || n === 'R') { vaciar(); actual = { L: 'izq', C: 'cen', R: 'der' }[n]; continue; }
      if ('PNDTFAZ'.includes(n)) { vaciar(); sec[actual].push(Object.assign({ campo: n }, est)); continue; }
      if (n === 'B' || n === 'I' || n === 'U' || n === 'E' || n === 'S') {
        vaciar();
        est = Object.assign({}, est);
        if (n === 'B') est.negrita = !est.negrita; else if (n === 'I') est.cursiva = !est.cursiva; else if (n === 'S') est.tachado = !est.tachado; else est.subrayado = !est.subrayado;
        continue;
      }
      if (n === '"') { // &"Arial,Bold"
        const fin = t.indexOf('"', i + 1);
        if (fin < 0) break;
        vaciar();
        const [nombre, estilo] = t.slice(i + 1, fin).split(',');
        est = Object.assign({}, est, { fuente: nombre || null });
        if (estilo !== undefined) { const e = estilo.toLowerCase(); est.negrita = /bold|negrita/.test(e); est.cursiva = /italic|cursiva/.test(e); }
        i = fin;
        continue;
      }
      if (n === 'K') { const m = /^[0-9A-Fa-f]{6}/.exec(t.slice(i + 1)); if (m) { vaciar(); est = Object.assign({}, est, { color: '#' + m[0] }); i += 6; } continue; }
      if (n >= '0' && n <= '9') { let d = n; while (i + 1 < t.length && t[i + 1] >= '0' && t[i + 1] <= '9' && d.length < 3) d += t[++i]; vaciar(); est = Object.assign({}, est, { tam: +d || null }); continue; }
      // &G (imagen), &X, &Y y lo desconocido: se ignora
    }
    vaciar();
    return sec;
  }

  // El encabezado/pie que toca a la página: primera, par/impar o normal.
  function textosHF(hf, numeroPagina, esPrimera) {
    if (!hf) return null;
    const grupo = (hf.distintoPrimera && esPrimera) ? hf.primera : (hf.distintoParImpar && numeroPagina % 2 === 0) ? hf.par : hf.impar;
    return grupo || null;
  }

  // ---- Geometría de impresión ------------------------------------------------------------------------
  // Posiciones (u) de las columnas y filas de la hoja tal como las deja Excel al imprimir a la escala de la hoja.
  function geometriaImpresion(g, p, hasta) {
    const s = p.escala / 100;
    const uPorPx = PT_PX / (U_PT * s);
    const razon = p.mdwImpresora / g.mdw;
    const ancho = (c) => (g.anchoCol(c) > 0 ? Math.round(Math.round(g.anchoCol(c) * razon) * s) * uPorPx : 0);
    const alto = (f) => (g.altoFila(f) > 0 ? Math.round(Math.floor(g.altoFila(f) * 6.05 + 1e-6) * s) * uPorPx : 0);
    const nC = Math.min(16384, hasta.c + 2), nF = Math.min(1048576, hasta.f + 2);
    const X = new Float64Array(nC + 1), Y = new Float64Array(nF + 1);
    for (let c = 0; c < nC; c++) X[c + 1] = X[c] + ancho(c);
    for (let f = 0; f < nF; f++) Y[f + 1] = Y[f] + alto(f);
    return {
      mdw: p.mdwImpresora * PT_PX / U_PT, // sin escala: el dígito mide lo mismo respecto a la letra
      anchoCol: (c) => (c < nC ? X[c + 1] - X[c] : 0), altoFila: (f) => (f < nF ? Y[f + 1] - Y[f] : 0),
      x: (c) => X[Math.min(c, nC)], y: (f) => Y[Math.min(f, nF)],
      altoTexto: g.altoTexto, anchoPorDefecto: ancho(nC), altoPorDefecto: alto(nF), columnasConocidas: nC, filasConocidas: nF,
      encAncho: ENC_ANCHO_PX * PT_PX / (U_PT * s), encAlto: ENC_ALTO_PX * PT_PX / (U_PT * s),
    };
  }

  const fmtFecha = (d, idioma) => d.toLocaleDateString(idioma === 'en' ? 'en-US' : 'es-ES');
  const fmtHora = (d, idioma) => d.toLocaleTimeString(idioma === 'en' ? 'en-US' : 'es-ES');

  // ---- Una página ------------------------------------------------------------------------------------
  function dibujarPagina(ctx, libro, hoja, G, p, pagina, ctxP, estilo) {
    const f = U_PT * (p.escala / 100);               // puntos por u
    const m = p.margenes;
    const mI = m.izq / f, mD = m.der / f, mA = m.arr / f, mB = m.aba / f;
    const anchoU = p.papel.ancho / f, altoU = p.papel.alto / f;
    const op = p.opciones;
    const cabeW = op.encabezados ? G.encAncho : 0, cabeH = op.encabezados ? G.encAlto : 0;
    const tf = pagina.titulos.filas, tc = pagina.titulos.columnas;
    const anchoTit = tc ? G.x(tc.c1 + 1) - G.x(tc.c0) : 0, altoTit = tf ? G.y(tf.f1 + 1) - G.y(tf.f0) : 0;
    const anchoCuerpo = G.x(pagina.c1 + 1) - G.x(pagina.c0), altoCuerpo = G.y(pagina.f1 + 1) - G.y(pagina.f0);
    const totalW = cabeW + anchoTit + anchoCuerpo, totalH = cabeH + altoTit + altoCuerpo;
    const dispW = anchoU - mI - mD, dispH = altoU - mA - mB;
    // centrado: Excel cuenta con las líneas de cierre (medido: 1,6 pt más de ancho y 2,2 pt más de alto)
    const ox = mI + (op.centrarH ? Math.max(0, (dispW - totalW - 1.6 / f) / 2) : 0);
    const oy = mA + (op.centrarV ? Math.max(0, (dispH - totalH - 2.2 / f) / 2) : 0);

    const e = p.escala / 100;
    // a escala menor que 100 % Excel deja el texto más abajo y a la derecha de lo que da la escala exacta (medido: el relleno de la celda no se achica con la escala; ≈ +0,95 pt a la izquierda, +0,25 a la derecha y +1,05 hacia abajo por cada 100 % que falta)
    const desplazarTexto = { x: 0.95 * (1 - e) / f, xDer: 0.25 * (1 - e) / f, xCen: 0, y: 1.05 * (1 - e) / f };
    const comun = { desplazarTexto, fino: true, idioma: ctxP.idioma, cuadricula: !!op.cuadricula, colorRejilla: '#000000', grosorRejilla: 0.96 / f, pxDe: estilo.pxDe };
    const R = (f0, c0, f1, c1, x, y) => XlsxDibujo.dibujarRango(ctx, libro, hoja, G, { f0, c0, f1, c1 }, Object.assign({ x, y }, comun));
    const xCuerpo = ox + cabeW + anchoTit, yCuerpo = oy + cabeH + altoTit;
    R(pagina.f0, pagina.c0, pagina.f1, pagina.c1, xCuerpo, yCuerpo);
    if (tf) R(tf.f0, pagina.c0, tf.f1, pagina.c1, xCuerpo, oy + cabeH);
    if (tc) R(pagina.f0, tc.c0, pagina.f1, tc.c1, ox + cabeW, yCuerpo);
    if (tf && tc) R(tf.f0, tc.c0, tf.f1, tc.c1, ox + cabeW, oy + cabeH);

    if (op.encabezados) { // A, B, C… y 1, 2, 3…: texto centrado entre líneas negras
      ctx.save();
      ctx.fillStyle = '#000000'; ctx.textAlign = 'center'; ctx.textBaseline = 'alphabetic';
      ctx.font = XlsxDibujo.fuenteCss({ nombre: estilo.normal.nombre, negrita: false, cursiva: false }, estilo.pxDe(estilo.normal.tam));
      const gr = 0.96 / f, desc = 4;
      const tramoC = [];
      if (tc) for (let c = tc.c0; c <= tc.c1; c++) tramoC.push([c, ox + cabeW + G.x(c) - G.x(tc.c0)]);
      for (let c = pagina.c0; c <= pagina.c1; c++) tramoC.push([c, xCuerpo + G.x(c) - G.x(pagina.c0)]);
      const tramoF = [];
      if (tf) for (let r = tf.f0; r <= tf.f1; r++) tramoF.push([r, oy + cabeH + G.y(r) - G.y(tf.f0)]);
      for (let r = pagina.f0; r <= pagina.f1; r++) tramoF.push([r, yCuerpo + G.y(r) - G.y(pagina.f0)]);
      const yTop = oy, xLeft = ox;
      ctx.fillRect(xLeft, yTop, totalW + gr, gr);
      ctx.fillRect(xLeft, yTop, gr, totalH + gr);
      ctx.fillRect(xLeft, yTop + cabeH, totalW + gr, gr);
      ctx.fillRect(xLeft + cabeW, yTop, gr, totalH + gr);
      for (const [c, x] of tramoC) {
        const w = G.anchoCol(c);
        if (w <= 0) continue;
        ctx.fillRect(x + w, yTop, gr, cabeH + gr);
        ctx.fillText(LibroXlsx.numeroACol(c), x + w / 2 + 0.9 / f, yTop + cabeH - desc);
      }
      for (const [r, y] of tramoF) {
        const h = G.altoFila(r);
        if (h <= 0) continue;
        ctx.fillRect(xLeft, y + h, cabeW + gr, gr);
        ctx.fillText(String(r + 1), xLeft + cabeW / 2 + 0.9 / f, y + h - desc);
      }
      ctx.restore();
    }

    // encabezado y pie: el texto cuelga del margen de encabezado y el pie sube desde el margen de pie (medido contra Excel:
    // línea base del encabezado = margen + 0,952·letra + 1 pt; la del pie = alto − margen − (0,28·letra + 1,72 pt))
    const grupo = textosHF(hoja.imprimir.encabezadoPie, ctxP.numeroPagina, ctxP.primeraDeLaHoja);
    if (grupo) {
      const escalar = hoja.imprimir.encabezadoPie.escalar !== false;
      const k = escalar ? p.escala / 100 : 1;                       // la letra del papel frente a la que dice el código
      const campo = (c) => ({ P: String(ctxP.numeroPagina), N: String(ctxP.totalPaginas), D: fmtFecha(ctxP.fecha, ctxP.idioma), T: fmtHora(ctxP.fecha, ctxP.idioma), F: ctxP.nombreArchivo || '', A: hoja.nombre || '', Z: '' }[c] || '');
      const dibujarSeccion = (texto, esPie) => {
        if (!texto) return;
        const sec = analizarHF(texto);
        for (const lado of ['izq', 'cen', 'der']) {
          if (!sec[lado].length) continue;
          const lineas = [[]];
          for (const r of sec[lado]) {
            const partes = (r.campo ? campo(r.campo) : r.texto).split(/\r?\n/);
            partes.forEach((pt, i) => { if (i > 0) lineas.push([]); if (pt) lineas[lineas.length - 1].push(Object.assign({}, r, { texto: pt, campo: undefined })); });
          }
          // los trozos seguidos con el mismo formato (texto y campos) se escriben como una sola cadena
          for (let i = 0; i < lineas.length; i++) {
            const mezcla = [];
            for (const r of lineas[i]) {
              const u = mezcla[mezcla.length - 1];
              if (u && u.fuente === r.fuente && u.negrita === r.negrita && u.cursiva === r.cursiva && u.subrayado === r.subrayado && u.tachado === r.tachado && u.tam === r.tam && u.color === r.color) u.texto += r.texto;
              else mezcla.push(Object.assign({}, r));
            }
            lineas[i] = mezcla;
          }
          const tamPt = (l) => (l.reduce((a, r) => Math.max(a, r.tam || 0), 0) || estilo.normal.tam || 11) * k; // pt en el papel
          const anchoDe = (r) => {
            ctx.font = XlsxDibujo.fuenteCss({ nombre: r.fuente || estilo.normal.nombre, negrita: r.negrita, cursiva: r.cursiva }, (r.tam || estilo.normal.tam || 11) * k / f);
            return ctx.measureText(r.texto).width;
          };
          const bases = new Array(lineas.length); // línea base de cada línea (pt desde arriba)
          if (!esPie) {
            let y = m.encabezado;
            lineas.forEach((l, i) => { const t = tamPt(l); bases[i] = y + 0.952 * t + 1.0; y += 1.22 * t; });
          } else {
            const i = lineas.length - 1;
            bases[i] = p.papel.alto - m.pie - (0.28 * tamPt(lineas[i]) + 1.72);
            for (let j = i - 1; j >= 0; j--) bases[j] = bases[j + 1] - 1.22 * tamPt(lineas[j + 1]);
          }
          lineas.forEach((l, i) => {
            const anchos = l.map(anchoDe);
            const total = anchos.reduce((a, b) => a + b, 0);
            let x = lado === 'izq' ? mI + 1 / f : lado === 'cen' ? mI + (dispW - total) / 2 : anchoU - mD - total - 1 / f;
            l.forEach((r, q) => {
              ctx.font = XlsxDibujo.fuenteCss({ nombre: r.fuente || estilo.normal.nombre, negrita: r.negrita, cursiva: r.cursiva }, (r.tam || estilo.normal.tam || 11) * k / f);
              ctx.fillStyle = r.color || '#000000'; ctx.textBaseline = 'alphabetic'; ctx.textAlign = 'left';
              ctx.fillText(r.texto, x, bases[i] / f);
              if (r.subrayado) ctx.fillRect(x, bases[i] / f + 1, anchos[q], 0.8);
              if (r.tachado) ctx.fillRect(x, bases[i] / f - 3, anchos[q], 0.8);
              x += anchos[q];
            });
          });
        }
      };
      ctx.save();
      dibujarSeccion(grupo.encabezado, false);
      dibujarSeccion(grupo.pie, true);
      ctx.restore();
    }
  }

  // ---- Hojas a PDF ------------------------------------------------------------------------------------
  async function generar(libro, indices, opc) {
    opc = opc || {};
    const idioma = opc.idioma || 'es';
    const pdf = await PDFLib.PDFDocument.create();
    await XlsxDibujo.cargarFuentes(libro);
    const usos = XlsxDibujo.usosDeFuentes(libro);
    for (const hf of ['Calibri', 'Arial', 'Times New Roman']) for (const [n, c] of [[false, false], [true, false], [false, true], [true, true]]) usos.push({ fuente: hf, negrita: n, cursiva: c });
    await new LienzoPdf.Lienzo({ ancho: 1, alto: 1, anchoPt: 1, altoPt: 1 }).precargarFuentes(usos); // las deja listas para dibujar sin esperar
    const normal = libro.estilos.xf(0).fuente;
    const estilo = { normal: { nombre: normal.nombre || 'Calibri', tam: normal.tam || 11 } };
    const trabajos = [];
    for (const i of indices) {
      const hoja = await libro.cargarHoja(i);
      const g = XlsxDibujo.geometria(libro, hoja);
      const p = XlsxPaginas.calcular(libro, hoja, g);
      trabajos.push({ hoja, g, p });
    }
    const total = trabajos.reduce((a, t) => a + t.p.paginas.length, 0);
    const fecha = new Date();
    let numero = 0, hechas = 0;
    for (const t of trabajos) {
      const { hoja, g, p } = t;
      if (!hoja.imprimir.pagina || !hoja.imprimir.pagina.usarPrimeraPagina || !(hoja.imprimir.pagina.primeraPagina >= 0)) { if (opc.numeracion === 'porHoja') numero = 0; }
      else numero = hoja.imprimir.pagina.primeraPagina - 1;
      let hasta = { c: 0, f: 0 };
      for (const a of p.areas) { hasta.c = Math.max(hasta.c, a.c1); hasta.f = Math.max(hasta.f, a.f1); }
      if (p.titulos.filas) hasta.f = Math.max(hasta.f, p.titulos.filas.f1);
      if (p.titulos.columnas) hasta.c = Math.max(hasta.c, p.titulos.columnas.c1);
      const G = geometriaImpresion(g, p, hasta);
      const f = U_PT * (p.escala / 100);
      // letra: píxeles de impresora redondeados (11 pt → 92 px → 11,04 pt) vistos en u
      const pxDe = (pt) => { const t = Number.isFinite(pt) && pt > 0 ? Math.min(409.5, pt) : 11; return Math.max(1, Math.round(t * (p.escala / 100) * 600 / 72)) * PT_PX / f; };
      estilo.pxDe = pxDe;
      let primera = true;
      for (const pagina of p.paginas) {
        numero++; hechas++;
        const L = new LienzoPdf.Lienzo({ ancho: p.papel.ancho / f, alto: p.papel.alto / f, anchoPt: p.papel.ancho, altoPt: p.papel.alto });
        dibujarPagina(L, libro, hoja, G, p, pagina, { idioma, numeroPagina: numero, totalPaginas: opc.totalPaginas || total, primeraDeLaHoja: primera, fecha, nombreArchivo: opc.nombreArchivo }, estilo);
        primera = false;
        if (L.fallo) throw L.fallo;
        await L.aPagina(pdf);
        if (opc.progreso) opc.progreso(hechas, total);
        if ((hechas & 7) === 7) await new Promise((r) => setTimeout(r, 0));
      }
    }
    if (!pdf.getPageCount()) pdf.addPage([612, 792]);
    return { bytes: await pdf.save(), paginas: total };
  }

  globalThis.XlsxPdf = { generar, analizarHF, geometriaImpresion };
})();

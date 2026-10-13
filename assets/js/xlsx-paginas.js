/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Paginación de una hoja de Excel al imprimir: en qué páginas cae cada fila y cada columna, igual que Excel (etapa 5 del
// motor de Excel). Reglas medidas contra Excel real (Investigaciones\investigacion-motor-excel\03-mediciones.md, etapa 5):
//  - se cuenta en píxeles de la impresora (600 dpi, "Microsoft Print to PDF"): una columna mide round(px de pantalla · MDWimp / MDW)
//    (MDWimp = ancho del "0" de la fuente Normal a 600 dpi: 47 con Calibri 11) y una fila trunc(px de pantalla · 6,05);
//  - con escala s % cada tamaño se vuelve a redondear: round(px · s / 100);
//  - el área útil es el papel menos los márgenes; la página se llena de a una columna/fila entera y la que no cabe pasa a la siguiente;
//  - "ajustar a N × M páginas" toma la escala entera más grande (10-100) con la que caben en N de ancho y M de alto (0 = libre);
//  - los saltos manuales cortan siempre; el área de impresión y las filas/columnas de título se leen de los nombres definidos.
// No depende del navegador. Requiere xlsx-modelo.js y xlsx-geometria.js.
//
// Uso:
//   const g = XlsxGeometria.crear(libro, hoja);
//   const p = XlsxPaginas.calcular(libro, hoja, g);
//   p.paginas → [{ f0, c0, f1, c1, area, titulos: { filas: {f0, f1}|null, columnas: {c0, c1}|null } }] en el orden de impresión
//   p.escala (%) · p.papel { ancho, alto } (pt, ya girado) · p.margenes (pt) · p.areas · p.opciones

(function () {
  const DPI = 600;
  const PX_FILA = 6.05; // píxeles de impresora por píxel de pantalla, en vertical
  const PT_PX = 72 / DPI;
  const RECORTE = -8; // píxeles que Excel le quita al área útil (medido: 7 a 9)
  const PENALIZACION = 0.81; // "ajustar a" con N ≥ 2 páginas de ancho: por cada salto Excel cuenta como si sobrara esta fracción de la columna "típica" (media ponderada por ancho)
  const MAX_PAGINAS = 20000; // una hoja no puede pasar de ahí (Excel tampoco imprime más de 32767)

  // Papel por código de ECMA-376 (puntos, vertical): lo medido en el PDF de Excel (papel.json) y la norma para el resto.
  const PAPELES = {
    1: [612, 792], 2: [612, 792], 3: [792, 1224], 4: [1224, 792], 5: [612, 1008], 6: [396, 612], 7: [522, 756], 8: [841.92, 1190.52],
    9: [595.32, 841.92], 10: [595.32, 841.92], 11: [419.52, 595.32], 12: [728.52, 1031.76], 13: [515.88, 728.52], 14: [612, 936],
    15: [609.45, 779.53], 16: [720, 1008], 17: [792, 1224], 18: [612, 792], 19: [279, 639], 20: [297, 684], 24: [1224, 1584],
    25: [1584, 2448], 26: [2448, 3168], 27: [311.76, 623.64], 28: [459.24, 649.08], 30: [649.08, 918.48], 34: [498.84, 708.72],
    37: [279, 540], 43: [283.44, 419.52], 66: [1190.52, 1683.72], 70: [297.6, 419.52], 71: [680.28, 941.16],
  };
  // Ancho del "0" como fracción del em (el mismo que usa xlsx-geometria para el MDW de pantalla).
  const DIGITO = { calibri: 0.5068, 'calibri light': 0.5405, arial: 0.5562, 'times new roman': 0.5, 'courier new': 0.6, verdana: 0.6362, tahoma: 0.5459, 'segoe ui': 0.5498, consolas: 0.5498, cambria: 0.5513, georgia: 0.6123, 'arial narrow': 0.456, aptos: 0.546, 'aptos narrow': 0.4829 };
  const ALIAS = { carlito: 'calibri', 'liberation sans': 'arial', arimo: 'arial', 'liberation serif': 'times new roman', tinos: 'times new roman', 'liberation mono': 'courier new', cousine: 'courier new', 'liberation sans narrow': 'arial narrow' };

  // MDW de la impresora: la fuente se pide al tamaño redondeado en píxeles de 600 dpi y el dígito se redondea de nuevo.
  function mdwImpresora(nombre, tam) {
    const k = String(nombre || '').toLowerCase().trim();
    const d = DIGITO[k] || DIGITO[ALIAS[k]] || DIGITO.calibri;
    return Math.max(1, Math.round(d * Math.round((tam > 0 ? tam : 11) * DPI / 72)));
  }

  function papelDe(imprimir) {
    const p = imprimir && imprimir.pagina;
    let ancho, alto;
    const medida = (v) => { const m = /^([\d.]+)(mm|cm|in|pt|pc|pi)$/.exec(v || ''); if (!m) return 0; return parseFloat(m[1]) * { mm: 72 / 25.4, cm: 72 / 2.54, in: 72, pt: 1, pc: 12, pi: 12 }[m[2]]; };
    if (p && medida(p.anchoPapel) > 0 && medida(p.altoPapel) > 0) { ancho = medida(p.anchoPapel); alto = medida(p.altoPapel); }
    else { const t = PAPELES[p ? p.papel : 1] || PAPELES[9]; ancho = t[0]; alto = t[1]; }
    // un archivo puede pedir cualquier medida ("999…9in" da Infinity): se acota a 3–14400 pt (lo que admite una página PDF)
    ancho = Math.min(14400, Math.max(3, ancho)); alto = Math.min(14400, Math.max(3, alto));
    if (p && p.orientacion === 'landscape' && alto > ancho) [ancho, alto] = [alto, ancho];
    else if (p && p.orientacion === 'portrait' && ancho > alto) [ancho, alto] = [alto, ancho];
    return { ancho, alto };
  }

  // Referencias de un nombre definido ("Hoja1!$A$1:$E$20,Hoja1!$H$5:$K$30" o "Hoja1!$1:$2") → rangos propios de esta hoja.
  function rangosDeNombre(libro, hoja, nombre) {
    const salida = [];
    const def = libro.nombres.find((n) => n.nombre === nombre && n.hojaLocal === hoja.indice);
    if (!def || !def.ref) return salida;
    // la expresión es cuadrática con una cadena larga sin "!": se lee hasta 8192 caracteres (el máximo de una fórmula en Excel)
    // y cada rango se acota a la hoja (A1:XFD1048576); uno que queda vacío se descarta
    const ref = String(def.ref).slice(0, 8192);
    const poner = (r) => {
      const q = { f0: Math.max(0, r.f0), c0: Math.max(0, r.c0), f1: Math.min(1048575, r.f1), c1: Math.min(16383, r.c1) };
      if (r.soloFilas) q.soloFilas = true; if (r.soloCols) q.soloCols = true;
      if (q.f0 <= q.f1 && q.c0 <= q.c1) salida.push(q);
    };
    const re =/(?:'((?:[^']|'')*)'|([^!,'()]+))!(\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)/g;
    let m;
    while ((m = re.exec(ref))) {
      const nombreHoja = (m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]).trim();
      const info = libro.hojas[hoja.indice];
      if (info && info.nombre !== undefined && info.nombre !== nombreHoja) continue;
      const r = m[3].replace(/\$/g, '');
      let x = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(r);
      if (x) { const a = LibroXlsx.colANumero(x[1].toUpperCase()), b = LibroXlsx.colANumero(x[2].toUpperCase()); poner({ f0: 0, c0: Math.min(a, b), f1: 1048575, c1: Math.max(a, b), soloCols: true }); continue; }
      x = /^(\d+):(\d+)$/.exec(r);
      if (x) { const a = +x[1] - 1, b = +x[2] - 1; poner({ f0: Math.min(a, b), c0: 0, f1: Math.max(a, b), c1: 16383, soloFilas: true }); continue; }
      const rg = LibroXlsx.decodificarRango(r);
      if (rg) poner(rg);
    }
    return salida;
  }

  // Última fila y columna que Excel imprime sin área: la última con contenido, con relleno o con borde, o el fin de una combinada.
  function extension(libro, hoja) {
    let f1 = -1, c1 = -1;
    const C = hoja.celdas;
    const visible = new Map();
    for (let i = 0; i < C.n; i++) {
      const fi = C.fila[i], ci = C.col[i];
      if (fi <= f1 && ci <= c1) continue;
      let v = C.tipo[i] !== LibroXlsx.TIPO.vacio;
      if (!v) {
        const e = C.estilo[i];
        if (!visible.has(e)) {
          const xf = libro.estilos.xf(e);
          const b = xf.borde || {};
          visible.set(e, (xf.relleno && xf.relleno.tipo !== 'ninguno' && xf.relleno.patron !== 'none') || !!(b.izq || b.der || b.arr || b.aba));
        }
        v = visible.get(e);
      }
      if (v) { if (fi > f1) f1 = fi; if (ci > c1) c1 = ci; }
    }
    for (const m of hoja.combinadas) { if (m.f1 > f1) f1 = m.f1; if (m.c1 > c1) c1 = m.c1; }
    return { f1, c1 };
  }

  // Corta [ini, fin] en bloques: cada uno se llena mientras quepa entero y los saltos manuales (índices donde empieza página) cortan siempre.
  // `titulo` { hasta, px }: las páginas que empiezan después de las filas de título pierden ese espacio.
  function cortar(tam, ini, fin, disponible, manuales, titulo, limite) {
    const bloques = [];
    let i = ini;
    while (i <= fin && bloques.length < limite) {
      const disp = disponible - (titulo && i > titulo.hasta ? titulo.px : 0);
      let suma = 0, j = i;
      while (j <= fin) {
        const t = tam(j);
        if (j > i && (manuales.has(j) || suma + t > disp)) break;
        suma += t; j++;
      }
      bloques.push([i, j - 1]);
      i = j;
    }
    return bloques;
  }

  function calcular(libro, hoja, g, opc) {
    opc = opc || {};
    const imp = hoja.imprimir || {};
    const papel = papelDe(imp);
    const mg = Object.assign({ izq: 0.7, der: 0.7, arr: 0.75, aba: 0.75, encabezado: 0.3, pie: 0.3 }, imp.margenes || {});
    const margenes = {};
    for (const k of Object.keys(mg)) margenes[k] = (Number.isFinite(mg[k]) ? Math.min(200, Math.max(0, mg[k])) : 0) * 72; // hasta 200" (el papel más grande): un margen de 1e300 llegaba al PDF como "1e+303"
    const opciones = imp.opciones || { cuadricula: false, encabezados: false, centrarH: false, centrarV: false };

    // Áreas: las del nombre Print_Area, o desde A1 hasta lo último que tenga algo.
    let areas = rangosDeNombre(libro, hoja, '_xlnm.Print_Area').filter((r) => !r.soloFilas && !r.soloCols);
    if (!areas.length) {
      const e = extension(libro, hoja);
      areas = e.f1 < 0 ? [] : [{ f0: 0, c0: 0, f1: e.f1, c1: e.c1 }];
    }
    areas = areas.slice(0, 256);
    const titulos = rangosDeNombre(libro, hoja, '_xlnm.Print_Titles');
    const tFilas = titulos.find((r) => r.soloFilas) || null, tCols = titulos.find((r) => r.soloCols) || null;

    // Tamaños de la impresora a escala 100 %.
    const normal = libro.estilos.xf(0).fuente;
    const mdwImp = mdwImpresora(normal.nombre, normal.tam);
    const razonCol = mdwImp / g.mdw;
    const baseCol = (c) => Math.round(g.anchoCol(c) * razonCol);
    const baseFila = (f) => Math.floor(g.altoFila(f) * PX_FILA + 1e-6);

    const encFilas = opciones.encabezados ? (opc.altoEncabezado || 121) : 0;
    const encCols = opciones.encabezados ? (opc.anchoEncabezado || 173) : 0;
    const anchoUtil = RECORTE + Math.round(papel.ancho * DPI / 72) - Math.round(margenes.izq * DPI / 72) - Math.round(margenes.der * DPI / 72) - encCols;
    const altoUtil = RECORTE + Math.round(papel.alto * DPI / 72) - Math.round(margenes.arr * DPI / 72) - Math.round(margenes.aba * DPI / 72) - encFilas;

    // Con "ajustar a" Excel ignora los saltos manuales.
    const saltosF = new Set(), saltosC = new Set();
    if (!imp.ajustarAPagina) for (const s of imp.saltosFila || []) if (s.manual !== false && s.id >= 1) saltosF.add(s.id); // el salto `id` es después de la fila id (desde 1): empieza en el índice id
    if (!imp.ajustarAPagina) for (const s of imp.saltosCol || []) if (s.manual !== false && s.id >= 1) saltosC.add(s.id);

    // Páginas de todas las áreas a una escala dada. Con `cuenta` { N, M } solo cuenta las páginas de ancho y de alto, y deja de cortar
    // apenas pasa de N y M (para "ajustar a", que prueba varias escalas). Nunca se recorren más de MAX_PAGINAS páginas en total.
    const paginar = (escala, cuenta) => {
      const k = escala / 100;
      const tc = (c) => Math.round(baseCol(c) * k), tf = (f) => Math.round(baseFila(f) * k);
      const salida = [];
      const resumen = { ancho: 0, alto: 0 };
      const sobre = imp.pagina && imp.pagina.orden === 'overThenDown';
      const lleno = () => salida.length >= MAX_PAGINAS;
      const pxTF = tFilas ? sumar(tf, tFilas.f0, tFilas.f1) : 0, pxTC = tCols ? sumar(tc, tCols.c0, tCols.c1) : 0;
      // cada página lleva solo los títulos que caben en ella (los de más caerían fuera del papel, y con "$1:$100000" serían
      // 100.000 filas que dibujar en cada una de 20.000 páginas)
      const tfHasta = tFilas ? hastaQueCabe(tf, tFilas.f0, tFilas.f1, altoUtil) : 0, tcHasta = tCols ? hastaQueCabe(tc, tCols.c0, tCols.c1, anchoUtil) : 0;
      for (let area = 0; area < areas.length && !lleno(); area++) {
        const a = areas[area];
        const tituloF = tFilas ? { hasta: Math.min(tFilas.f1, a.f1), px: pxTF } : null;
        const tituloC = tCols ? { hasta: Math.min(tCols.c1, a.c1), px: pxTC } : null;
        const limC = cuenta ? (cuenta.N > 0 ? cuenta.N + 1 : 1) : MAX_PAGINAS, limF = cuenta ? (cuenta.M > 0 ? cuenta.M + 1 : 1) : MAX_PAGINAS;
        const cols = cortar(tc, a.c0, a.c1, anchoUtil, saltosC, tituloC, limC);
        const filas = cortar(tf, a.f0, a.f1, altoUtil, saltosF, tituloF, limF);
        if (cols.length > resumen.ancho) resumen.ancho = cols.length;
        if (filas.length > resumen.alto) resumen.alto = filas.length;
        if (cuenta) continue;
        const poner = (cb, fb) => {
          salida.push({
            f0: fb[0], f1: fb[1], c0: cb[0], c1: cb[1], area,
            titulos: { filas: tFilas && fb[0] > tFilas.f1 ? { f0: tFilas.f0, f1: tfHasta } : null, columnas: tCols && cb[0] > tCols.c1 ? { c0: tCols.c0, c1: tcHasta } : null },
          });
        };
        const [externas, internas] = sobre ? [filas, cols] : [cols, filas];
        for (const e of externas) { if (lleno()) break; for (const q of internas) { if (lleno()) break; if (sobre) poner(q, e); else poner(e, q); } }
      }
      return { paginas: salida, resumen };
    };
    const sumar = (tam, a, b) => { let s = 0; for (let i = a; i <= b && i - a < 100000; i++) s += tam(i); return s; };
    // última fila/columna de [a, b] que todavía empieza dentro de `disp` (al menos la primera)
    const hastaQueCabe = (tam, a, b, disp) => { let s = 0, i = a; for (; i < b && i - a < 100000; i++) { s += tam(i); if (s >= disp) break; } return i; };

    let escala = 100;
    const pag = imp.pagina || {};
    if (imp.ajustarAPagina) {
      const N = pag.ajustarAncho === undefined ? 1 : pag.ajustarAncho, M = pag.ajustarAlto === undefined ? 1 : pag.ajustarAlto;
      // Ancho: además de que las páginas salgan, Excel pide que el ancho total (más un salto de sobra por cada página extra) quepa en N páginas.
      const anchoCabe = (s) => {
        if (!(N > 1) || saltosC.size) return true;
        const k = s / 100;
        for (const a of areas) {
          let suma = 0, base = 0, base2 = 0;
          for (let c = a.c0; c <= a.c1; c++) { const b = baseCol(c); suma += Math.round(b * k); base += b; base2 += b * b; }
          const ponderada = base ? base2 / base : 0; // media de los anchos ponderada por el propio ancho: las columnas anchas pesan más
          if (suma + (N - 1) * Math.round(PENALIZACION * ponderada * k) > N * anchoUtil) return false;
        }
        return true;
      };
      const cabe = (s) => { const r = paginar(s, { N, M }).resumen; return (!(N > 0) || r.ancho <= N) && (!(M > 0) || r.alto <= M) && anchoCabe(s); };
      if (!cabe(100)) {
        let lo = 10, hi = 99; // la mayor escala que cabe (si ni a 10 % cabe, se queda en 10)
        if (!cabe(10)) hi = 10;
        else { while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (cabe(mid)) lo = mid; else hi = mid - 1; } hi = lo; }
        escala = hi;
      }
    } else if (pag.escala > 0) escala = Math.min(400, Math.max(10, pag.escala));

    const { paginas } = paginar(escala);
    return { escala, papel, margenes, areas, opciones, titulos: { filas: tFilas, columnas: tCols }, paginas, anchoUtil, altoUtil, mdwImpresora: mdwImp };
  }

  globalThis.XlsxPaginas = { calcular, papelDe, mdwImpresora, PAPELES, cortar };
})();

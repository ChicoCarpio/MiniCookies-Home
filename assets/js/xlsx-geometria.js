/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Geometría de una hoja de Excel en píxeles de pantalla (96 dpi, zoom 100 %): ancho de cada columna, alto de cada
// fila y la posición de cada una. Reglas medidas en Excel real (Investigaciones\investigacion-motor-excel\03-mediciones.md):
//  - ancho guardado `w` (en caracteres, con el relleno incluido) → px = trunc((256·w + trunc(128/MDW)) / 256 · MDW),
//    donde MDW es el ancho en píxeles del dígito "0" de la fuente del estilo Normal;
//  - sin `defaultColWidth`, la columna mide ceil((baseColWidth·MDW + 5) / 8) · 8 px (64 px con MDW 7);
//  - alto de fila en puntos → px = round(pt · 96 / 72); una fila sin alto guardado se ajusta a su fuente más alta
//    (tabla ALTOS, medida con AutoFit y con el alto estándar de cada fuente como estilo Normal).
// No depende del navegador. Requiere xlsx-modelo.js (el modelo de la hoja).
//
// Uso:
//   const g = XlsxGeometria.crear(libro, hoja, { medirDigito });  // medirDigito(nombre, tamPx) → ancho del "0" (opcional)
//   g.anchoCol(c) · g.altoFila(f)        px (columna y fila desde 0)
//   g.x(c) · g.y(f)                      posición del borde izquierdo / superior (px, desde A1)
//   g.colEn(x) · g.filaEn(y)             columna / fila que contiene el píxel
//   g.mdw · g.altoPorDefecto (px) · g.anchoPorDefecto (px)

(function () {
  // Alto de una línea de texto (pt) por fuente, tamaño y variante (n, b negrita, i cursiva, bi): lo que Excel suma por cada línea de una fila. Medido con
  // AutoFit de 2 y 3 líneas (la diferencia es una línea; 48, 60 y 72 pt: AutoFit de una).
  const TAMANOS = [6, 7, 8, 9, 10, 10.5, 11, 12, 13, 14, 15, 16, 18, 20, 22, 24, 26, 28, 32, 36, 40, 48, 60, 72];
  const ALTOS = {
    'calibri': {
      n: [8.25, 9, 11.25, 12, 12.75, 14.25, 15, 15.75, 17.25, 18.75, 19.5, 21, 23.25, 26.25, 28.5, 31.5, 33.75, 36, 42, 46.5, 51, 61.5, 76.5, 92.25],
      b: [8.25, 9, 11.25, 12, 12.75, 14.25, 15, 15.75, 17.25, 18.75, 19.5, 21, 23.25, 26.25, 28.5, 31.5, 33.75, 36, 42, 46.5, 51, 61.5, 76.5, 92.25],
      i: [8.25, 9, 11.25, 12, 12.75, 14.25, 15, 15.75, 17.25, 18.75, 19.5, 21, 23.25, 26.25, 28.5, 31.5, 33.75, 36, 42, 46.5, 51, 61.5, 76.5, 92.25],
      bi: [8.25, 9, 11.25, 12, 12.75, 14.25, 15, 15.75, 17.25, 18.75, 19.5, 21, 23.25, 26.25, 28.5, 31.5, 33.75, 36, 42, 46.5, 51, 61.5, 76.5, 92.25],
    },
    'calibri light': {
      n: [8.25, 9, 11.25, 12, 12.75, 14.25, 15, 15.75, 17.25, 18.75, 19.5, 21, 23.25, 26.25, 28.5, 31.5, 33.75, 36, 42, 46.5, 51, 61.5, 76.5, 92.25],
      b: [8.25, 9, 11.25, 12, 12.75, 14.25, 15, 15.75, 17.25, 18.75, 19.5, 21, 23.25, 26.25, 28.5, 31.5, 33.75, 36, 42, 46.5, 51, 61.5, 76.5, 92.25],
      i: [8.25, 9, 11.25, 12, 12.75, 14.25, 15, 15.75, 17.25, 18.75, 19.5, 21, 23.25, 26.25, 28.5, 31.5, 33.75, 36, 42, 46.5, 51, 61.5, 76.5, 92.25],
      bi: [8.25, 9, 11.25, 12, 12.75, 14.25, 15, 15.75, 17.25, 18.75, 19.5, 21, 23.25, 26.25, 28.5, 31.5, 33.75, 36, 42, 46.5, 51, 61.5, 76.5, 92.25],
    },
    'arial': {
      n: [8.25, 9, 11.25, 12, 12.75, 13.5, 14.25, 15, 16.5, 18, 18.75, 20.25, 23.25, 25.5, 27, 30, 33, 34.5, 40.5, 44.25, 49.5, 59.25, 75, 90],
      b: [8.25, 9, 11.25, 12, 12.75, 13.5, 15, 15.75, 16.5, 18, 19.5, 20.25, 23.25, 26.25, 27.75, 30, 33.75, 35.25, 41.25, 45, 50.25, 60, 75.75, 90.75],
      i: [8.25, 9, 11.25, 12, 12.75, 13.5, 14.25, 15, 16.5, 18.75, 19.5, 20.25, 23.25, 25.5, 27.75, 30.75, 33, 35.25, 40.5, 45, 50.25, 60, 75, 90.75],
      bi: [8.25, 9, 10.5, 12, 12.75, 13.5, 14.25, 15, 16.5, 18.75, 19.5, 20.25, 23.25, 25.5, 27.75, 30, 33, 34.5, 40.5, 44.25, 49.5, 60, 75.75, 90.75],
    },
    'times new roman': {
      n: [8.25, 9, 11.25, 12, 12.75, 13.5, 15, 15.75, 16.5, 18.75, 19.5, 20.25, 23.25, 26.25, 27.75, 30.75, 33, 35.25, 40.5, 45.75, 50.25, 61.5, 76.5, 91.5],
      b: [8.25, 9, 10.5, 12, 12.75, 13.5, 14.25, 15.75, 16.5, 18.75, 19.5, 20.25, 22.5, 25.5, 27, 30, 33, 34.5, 39.75, 45, 49.5, 60.75, 75, 90],
      i: [8.25, 9, 11.25, 12, 12.75, 13.5, 15, 15.75, 16.5, 18.75, 19.5, 20.25, 23.25, 26.25, 27.75, 30.75, 33, 35.25, 40.5, 45.75, 50.25, 61.5, 76.5, 91.5],
      bi: [8.25, 9, 11.25, 12, 13.5, 14.25, 15, 15.75, 17.25, 19.5, 20.25, 20.25, 23.25, 25.5, 27, 30, 33, 34.5, 39.75, 45, 49.5, 60.75, 75, 90],
    },
    'courier new': {
      n: [8.25, 9, 11.25, 12, 13.5, 14.25, 15, 15.75, 17.25, 18.75, 19.5, 21, 24, 26.25, 28.5, 31.5, 34.5, 36, 42, 46.5, 51.75, 62.25, 77.25, 92.25],
      b: [8.25, 9, 11.25, 12.75, 13.5, 14.25, 15.75, 16.5, 17.25, 19.5, 20.25, 21, 24, 27, 29.25, 32.25, 35.25, 36.75, 42.75, 47.25, 52.5, 63, 78.75, 94.5],
      i: [8.25, 9, 12, 12.75, 13.5, 14.25, 15.75, 16.5, 17.25, 18.75, 20.25, 21, 24, 26.25, 28.5, 30.75, 33.75, 36, 42, 46.5, 51, 62.25, 77.25, 92.25],
      bi: [8.25, 9, 11.25, 12.75, 13.5, 14.25, 15.75, 16.5, 17.25, 19.5, 20.25, 21, 24, 27, 29.25, 32.25, 35.25, 36.75, 42.75, 47.25, 52.5, 63, 78.75, 94.5],
    },
    'verdana': {
      n: [8.25, 9, 10.5, 11.25, 12.75, 13.5, 14.25, 15, 15.75, 18, 20.25, 19.5, 22.5, 24.75, 27, 29.25, 32.25, 35.25, 39.75, 45.75, 50.25, 60, 74.25, 87.75],
      b: [8.25, 9, 10.5, 11.25, 12.75, 13.5, 14.25, 15, 15.75, 18, 20.25, 19.5, 22.5, 24.75, 27, 29.25, 32.25, 35.25, 39.75, 45.75, 50.25, 60, 74.25, 87.75],
      i: [8.25, 9, 10.5, 11.25, 12.75, 13.5, 14.25, 15, 15.75, 18, 20.25, 19.5, 22.5, 24.75, 27, 29.25, 32.25, 35.25, 39.75, 45.75, 50.25, 60, 74.25, 87.75],
      bi: [8.25, 9, 10.5, 11.25, 12.75, 13.5, 14.25, 15, 15.75, 18, 20.25, 19.5, 22.5, 24.75, 27, 29.25, 32.25, 35.25, 39.75, 45.75, 50.25, 60, 74.25, 87.75],
    },
    'tahoma': {
      n: [8.25, 9, 10.5, 11.25, 12.75, 13.5, 14.25, 15, 16.5, 18, 18.75, 19.5, 22.5, 25.5, 27, 30, 32.25, 34.5, 39.75, 44.25, 48.75, 58.5, 73.5, 87.75],
      b: [8.25, 9, 10.5, 11.25, 12.75, 13.5, 14.25, 15, 16.5, 18, 18.75, 19.5, 22.5, 25.5, 27, 30, 32.25, 34.5, 39.75, 44.25, 48.75, 58.5, 73.5, 87.75],
      i: [8.25, 9, 10.5, 11.25, 12.75, 13.5, 14.25, 15, 16.5, 18, 18.75, 19.5, 22.5, 25.5, 27, 30, 32.25, 34.5, 39.75, 44.25, 48.75, 58.5, 73.5, 87.75],
      bi: [8.25, 9, 10.5, 11.25, 12.75, 13.5, 14.25, 15, 16.5, 18, 18.75, 19.5, 22.5, 25.5, 27, 30, 32.25, 34.5, 39.75, 44.25, 48.75, 58.5, 73.5, 87.75],
    },
    'segoe ui': {
      n: [8.25, 9, 10.5, 12, 14.25, 15.75, 16.5, 17.25, 18.75, 20.25, 24, 25.5, 26.25, 30.75, 33, 37.5, 38.25, 40.5, 48, 52.5, 57, 69, 84.75, 102],
      b: [8.25, 9, 10.5, 12, 14.25, 15.75, 16.5, 17.25, 18.75, 20.25, 24, 25.5, 26.25, 30.75, 33, 37.5, 38.25, 40.5, 48, 52.5, 57, 69, 84.75, 102],
      i: [8.25, 9, 10.5, 12, 14.25, 15.75, 16.5, 17.25, 18.75, 20.25, 24, 25.5, 26.25, 30.75, 33, 37.5, 38.25, 40.5, 48, 52.5, 57, 69, 84.75, 102],
      bi: [8.25, 9, 10.5, 12, 14.25, 15.75, 16.5, 17.25, 18.75, 20.25, 24, 25.5, 26.25, 30.75, 33, 37.5, 38.25, 40.5, 48, 52.5, 57, 69, 84.75, 102],
    },
    'consolas': {
      n: [8.25, 9, 11.25, 12, 12.75, 14.25, 15, 15.75, 16.5, 18.75, 19.5, 20.25, 23.25, 26.25, 27.75, 30.75, 33.75, 35.25, 41.25, 45.75, 50.25, 60.75, 75.75, 90.75],
      b: [8.25, 9, 11.25, 12, 12.75, 14.25, 15, 15.75, 16.5, 18.75, 19.5, 20.25, 23.25, 26.25, 27.75, 30.75, 33.75, 35.25, 41.25, 45.75, 50.25, 60.75, 75.75, 90.75],
      i: [8.25, 9, 11.25, 12, 12.75, 14.25, 15, 15.75, 16.5, 18.75, 19.5, 20.25, 23.25, 26.25, 27.75, 30.75, 33.75, 35.25, 41.25, 45.75, 50.25, 60.75, 75.75, 90.75],
      bi: [8.25, 9, 11.25, 12, 12.75, 14.25, 15, 15.75, 16.5, 18.75, 19.5, 20.25, 23.25, 26.25, 27.75, 30.75, 33.75, 35.25, 41.25, 45.75, 50.25, 60.75, 75.75, 90.75],
    },
    'cambria': {
      n: [8.25, 9, 10.5, 12, 12.75, 13.5, 14.25, 15.75, 16.5, 18, 18.75, 20.25, 22.5, 25.5, 27, 30, 33, 34.5, 40.5, 45, 49.5, 59.25, 74.25, 88.5],
      b: [8.25, 9, 10.5, 12, 12.75, 13.5, 14.25, 15.75, 16.5, 18, 18.75, 20.25, 22.5, 25.5, 27, 30, 33, 34.5, 40.5, 45, 49.5, 59.25, 74.25, 88.5],
      i: [8.25, 9, 10.5, 12, 12.75, 13.5, 14.25, 15.75, 16.5, 18, 18.75, 20.25, 22.5, 25.5, 27, 30, 33, 34.5, 40.5, 45, 49.5, 59.25, 74.25, 88.5],
      bi: [8.25, 9, 10.5, 12, 12.75, 13.5, 14.25, 15.75, 16.5, 18, 18.75, 20.25, 22.5, 25.5, 27, 30, 33, 34.5, 40.5, 45, 49.5, 59.25, 74.25, 88.5],
    },
    'georgia': {
      n: [8.25, 9, 11.25, 12, 12.75, 13.5, 14.25, 15, 16.5, 18, 19.5, 20.25, 23.25, 25.5, 27, 30, 33, 34.5, 39.75, 45, 49.5, 59.25, 74.25, 88.5],
      b: [8.25, 9, 11.25, 12, 12.75, 13.5, 14.25, 15, 16.5, 18, 19.5, 20.25, 23.25, 25.5, 27, 30, 33, 34.5, 39.75, 45, 49.5, 59.25, 74.25, 88.5],
      i: [8.25, 9, 11.25, 12, 12.75, 13.5, 14.25, 15, 16.5, 18, 19.5, 20.25, 23.25, 25.5, 27, 30, 33, 34.5, 39.75, 45, 49.5, 59.25, 74.25, 88.5],
      bi: [8.25, 9, 11.25, 12, 12.75, 13.5, 14.25, 15, 16.5, 18, 19.5, 20.25, 23.25, 25.5, 27, 30, 33, 34.5, 39.75, 45, 49.5, 59.25, 74.25, 88.5],
    },
    'arial narrow': {
      n: [8.25, 9, 12.75, 13.5, 12.75, 13.5, 16.5, 15.75, 17.25, 18, 19.5, 20.25, 23.25, 25.5, 27, 30, 33.75, 35.25, 40.5, 45.75, 49.5, 60, 74.25, 88.5],
      b: [8.25, 9, 12.75, 13.5, 12.75, 13.5, 16.5, 15.75, 17.25, 18, 19.5, 20.25, 23.25, 25.5, 27, 30, 33.75, 35.25, 40.5, 45.75, 49.5, 60, 74.25, 88.5],
      i: [8.25, 9, 12.75, 13.5, 12.75, 13.5, 16.5, 15.75, 17.25, 18, 19.5, 20.25, 23.25, 25.5, 27, 30, 33.75, 35.25, 40.5, 45.75, 49.5, 60, 74.25, 88.5],
      bi: [8.25, 9, 12.75, 13.5, 12.75, 13.5, 16.5, 15.75, 17.25, 18, 19.5, 20.25, 23.25, 25.5, 27, 30, 33.75, 35.25, 40.5, 45.75, 49.5, 60, 74.25, 88.5],
    },
    'comic sans ms': {
      n: [8.25, 9, 12.75, 14.25, 15, 15.75, 16.5, 19.5, 20.25, 21, 23.25, 24, 27, 31.5, 33, 37.5, 40.5, 42.75, 48.75, 54.75, 60.75, 73.5, 91.5, 110.25],
      b: [8.25, 9, 14.25, 13.5, 16.5, 15.75, 18, 19.5, 21, 22.5, 24.75, 24.75, 29.25, 31.5, 33, 37.5, 40.5, 41.25, 49.5, 54.75, 60.75, 73.5, 90, 110.25],
      i: [8.25, 9, 13.5, 14.25, 16.5, 17.25, 18.75, 19.5, 20.25, 22.5, 24.75, 24, 28.5, 31.5, 33.75, 36.75, 41.25, 42.75, 50.25, 55.5, 61.5, 72, 92.25, 111],
      bi: [8.25, 9, 12.75, 15.75, 16.5, 17.25, 18.75, 19.5, 20.25, 22.5, 24.75, 24, 28.5, 31.5, 33.75, 36.75, 41.25, 42.75, 50.25, 55.5, 61.5, 73.5, 92.25, 111],
    },
    'trebuchet ms': {
      n: [8.25, 9, 13.5, 15, 15, 14.25, 16.5, 18, 18, 18.75, 20.25, 21, 23.25, 27.75, 28.5, 30.75, 33.75, 36, 42, 46.5, 51, 61.5, 76.5, 92.25],
      b: [8.25, 9, 13.5, 15, 15, 14.25, 16.5, 18, 18, 18.75, 20.25, 21, 23.25, 27.75, 28.5, 30.75, 33.75, 36, 42, 46.5, 51, 61.5, 76.5, 92.25],
      i: [8.25, 9, 13.5, 15, 15, 14.25, 16.5, 18, 18, 18.75, 20.25, 21, 23.25, 27.75, 28.5, 30.75, 33.75, 36, 42, 46.5, 51, 61.5, 76.5, 92.25],
      bi: [8.25, 9, 13.5, 15, 15, 14.25, 16.5, 18, 18, 18.75, 20.25, 21, 23.25, 27.75, 28.5, 30.75, 33.75, 36, 42, 46.5, 51, 61.5, 76.5, 92.25],
    },
  };
  const ALIAS = {
    carlito: 'calibri', 'liberation sans': 'arial', arimo: 'arial', 'liberation serif': 'times new roman', tinos: 'times new roman',
    'liberation mono': 'courier new', cousine: 'courier new', 'liberation sans narrow': 'arial narrow', gelasio: 'georgia',
    'aptos narrow': 'calibri', aptos: 'calibri', 'calibri light': 'calibri light', 'segoe ui': 'segoe ui',
  };
  // Ancho del "0" como fracción del em, para sacar el MDW cuando no hay Canvas con la fuente. (Medido contra Excel: Calibri 11 → 7,
  // Calibri 12 → 8, Arial 10 → 7, Arial 11 → 8, Times New Roman 11 → 7, Calibri Light 11 → 8, Aptos 11 → 8, Aptos Narrow 11 → 7.)
  const DIGITO = { calibri: 0.5068, 'calibri light': 0.5405, arial: 0.5562, 'times new roman': 0.5, 'courier new': 0.6, verdana: 0.6362, tahoma: 0.5459, 'segoe ui': 0.5498, consolas: 0.5498, cambria: 0.5513, georgia: 0.6123, 'arial narrow': 0.456, aptos: 0.5469, 'aptos narrow': 0.4795 };

  const clave = (nombre) => { const k = String(nombre || '').toLowerCase().trim(); return ALIAS[k] && !DIGITO[k] ? ALIAS[k] : k; };

  // Alto de texto (pt) de una fuente con ese tamaño, tal como lo deja Excel en la fila.
  function altoTexto(nombre, tam, negrita, cursiva) {
    const fuente = ALTOS[clave(nombre)] || ALTOS[ALIAS[clave(nombre)]] || ALTOS.calibri;
    const tabla = fuente[(negrita ? 'b' : '') + (cursiva ? 'i' : '') || 'n'];
    const t = tam || 11;
    if (t <= TAMANOS[0]) return tabla[0];
    if (t >= TAMANOS[TAMANOS.length - 1]) return tabla[TAMANOS.length - 1] * t / TAMANOS[TAMANOS.length - 1];
    let i = 0;
    while (TAMANOS[i + 1] <= t) i++;
    if (TAMANOS[i] === t) return tabla[i];
    const a = tabla[i], b = tabla[i + 1], f = (t - TAMANOS[i]) / (TAMANOS[i + 1] - TAMANOS[i]);
    return Math.ceil((a + (b - a) * f) / 0.75 - 1e-9) * 0.75;
  }

  // MDW: ancho en píxeles del dígito "0" de la fuente Normal. medirDigito(nombre, tamPx) lo da el Canvas si hay.
  function mdwDe(nombre, tam, medirDigito) {
    const px = (tam || 11) * 96 / 72;
    if (medirDigito) { const m = medirDigito(nombre, px); if (m > 0) return Math.max(1, Math.round(m)); }
    return Math.max(1, Math.round(px * (DIGITO[clave(nombre)] || DIGITO.calibri)));
  }

  const pxAncho = (w, mdw) => Math.trunc(((256 * w + Math.trunc(128 / mdw)) / 256) * mdw);
  const pxAlto = (pt) => Math.floor(pt * 96 / 72 + 1e-6); // Excel trunca al píxel: 8,1 pt (10,8 px) → 10 px

  function crear(libro, hoja, opc) {
    opc = opc || {};
    const normal = libro.estilos.xf(0).fuente;
    const nombreNormal = normal.nombre || 'Calibri', tamNormal = normal.tam || 11;
    const mdw = opc.mdw || mdwDe(nombreNormal, tamNormal, opc.medirDigito);
    const f = hoja.formato;

    // --- columnas ---
    let anchoDef;
    if (f.anchoPorDefecto) anchoDef = pxAncho(f.anchoPorDefecto, mdw);
    else anchoDef = Math.ceil(((f.anchoBase || 8) * mdw + 5) / 8) * 8;
    let topeC = hoja.celdas.n ? 0 : -1;
    hoja.celdas.paraCada((fi, c) => { if (c > topeC) topeC = c; });
    for (const col of hoja.columnas) if (col.max > topeC) topeC = col.max;
    const nC = topeC + 1;
    const anchos = new Int32Array(Math.max(nC, 0)).fill(anchoDef);
    for (const col of hoja.columnas) {
      const px = col.oculta ? 0 : (col.ancho === null || col.ancho === undefined ? anchoDef : pxAncho(col.ancho, mdw));
      for (let c = col.min; c <= col.max && c < nC; c++) anchos[c] = px;
    }
    const izq = new Float64Array(nC + 1);
    for (let c = 0; c < nC; c++) izq[c + 1] = izq[c] + anchos[c];

    // --- filas ---
    const ptDef = f.altoPorDefecto || altoTexto(nombreNormal, tamNormal);
    const altoDef = f.filasOcultasPorDefecto ? 0 : pxAlto(ptDef);
    let topeF = -1;
    for (const k of hoja.filas.keys()) if (k > topeF) topeF = k;
    if (hoja.celdas.n) hoja.celdas.paraCada((fi) => { if (fi > topeF) topeF = fi; });
    const nF = topeF + 1;
    // Fuente más alta (pt) de cada fila sin alto guardado.
    const ptAuto = new Float32Array(nF);
    const C = hoja.celdas;
    const combinadasArriba = new Set(hoja.combinadas.map((m) => m.f0 * 16384 + m.c0));
    for (let i = 0; i < C.n; i++) {
      if (C.tipo[i] === LibroXlsx.TIPO.vacio) continue;
      const fi = C.fila[i];
      const info = hoja.filas.get(fi);
      if (info && info.alto !== null && info.alto !== undefined) continue;
      const fu = libro.estilos.xf(C.estilo[i]).fuente;
      let h = altoTexto(fu.nombre || nombreNormal, fu.tam || tamNormal, fu.negrita, fu.cursiva);
      // Texto con ajuste: una línea por cada salto o por cada vez que no cabe (lo cuenta lineasDe, que sabe medir texto).
      // Excel no ajusta la fila al contenido de una celda combinada.
      if (C.tipo[i] === LibroXlsx.TIPO.texto && libro.estilos.xf(C.estilo[i]).alineacion.ajustar && !combinadasArriba.has(fi * 16384 + C.col[i])) {
        const n = opc.lineasDe ? opc.lineasDe(i, anchos[C.col[i]] !== undefined ? anchos[C.col[i]] : anchoDef) : String(C.valor[i]).split('\n').length;
        h *= Math.max(1, n); // con n líneas Excel da n veces el alto de una (medido al abrir un libro sin altos guardados)
      }
      if (h > ptAuto[fi]) ptAuto[fi] = h;
    }
    const altos = new Int32Array(Math.max(nF, 0));
    for (let fi = 0; fi < nF; fi++) {
      const info = hoja.filas.get(fi);
      if (info && info.oculta) { altos[fi] = 0; continue; }
      if (info && info.alto !== null && info.alto !== undefined) { altos[fi] = pxAlto(info.alto); continue; }
      altos[fi] = f.filasOcultasPorDefecto ? 0 : pxAlto(Math.max(ptDef, ptAuto[fi]));
    }
    const arr = new Float64Array(nF + 1);
    for (let fi = 0; fi < nF; fi++) arr[fi + 1] = arr[fi] + altos[fi];

    const busca = (pref, n, def, v) => {
      if (v < 0) return -1;
      if (v >= pref[n]) return def > 0 ? n + Math.floor((v - pref[n]) / def) : -1;
      let a = 0, b = n - 1;
      while (a < b) { const m = (a + b + 1) >> 1; if (pref[m] <= v) a = m; else b = m - 1; }
      // salta las de tamaño cero: la columna que contiene el píxel es la última con inicio <= v y tamaño > 0
      while (a < n - 1 && pref[a + 1] <= v) a++;
      return a;
    };

    return {
      mdw, anchoPorDefecto: anchoDef, altoPorDefecto: altoDef, columnasConocidas: nC, filasConocidas: nF,
      anchoCol: (c) => (c < nC ? anchos[c] : anchoDef),
      altoFila: (fi) => (fi < nF ? altos[fi] : altoDef),
      x: (c) => (c < nC ? izq[c] : izq[nC] + (c - nC) * anchoDef),
      y: (fi) => (fi < nF ? arr[fi] : arr[nF] + (fi - nF) * altoDef),
      colEn: (x) => busca(izq, nC, anchoDef, x),
      filaEn: (y) => busca(arr, nF, altoDef, y),
      altoTexto, normal: { nombre: nombreNormal, tam: tamNormal },
    };
  }

  globalThis.XlsxGeometria = { crear, altoTexto, mdwDe, pxAncho, pxAlto };
})();

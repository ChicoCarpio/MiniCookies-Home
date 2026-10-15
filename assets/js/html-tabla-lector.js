/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Lector de tablas HTML que se guardan con extensión de Excel (etapa 9 del motor de Excel): muchos sistemas "exportan a Excel" una página HTML con un
// <table> y la llaman .xls (o .xlsx). Cada <table> es una hoja. Usa DOMParser, así que solo corre en el navegador. Requiere csv-planilla.js (decodificar).
//
//   HtmlTablaLector.parece(bytes)   → true si el archivo es HTML con al menos una tabla
//   HtmlTablaLector.leer(bytes)     → { hojas: [{ nombre, filas, anchos: [], combinadas }] }   (mismo formato que OdsLector)
//
// Una celda con x:num="…" (lo que escribe Excel al guardar como página web) entrega ese número; si no, un texto que sea un número "limpio" (1234, -5.25;
// sin ceros a la izquierda) pasa a número y el resto queda como texto. Las celdas con colspan o rowspan quedan como combinadas.
(function () {
  const MAX_FILAS = 1048576, MAX_COLS = 16384, MAX_TEXTO = 32767, MAX_TABLAS = 255, MAX_COMBINADAS = 500000;

  const inicio = (bytes) => {
    let s = '';
    for (let k = 0, n = Math.min(bytes.length, 4096); k < n; k++) { if (bytes[k] === 0) return ''; s += String.fromCharCode(bytes[k]); }
    return s;
  };
  function parece(bytes) {
    const s = inicio(bytes);
    return s !== '' && /<\s*(!doctype\s+html|html|table|head|body)\b/i.test(s) && (/<\s*table\b/i.test(s) || bytes.length > 4096);
  }

  function textoDeCelda(celda) {
    let s = '';
    const recorrer = (nodo) => {
      for (const h of nodo.childNodes) {
        if (h.nodeType === 3) s += h.nodeValue;
        else if (h.nodeType === 1) {
          const t = h.tagName.toLowerCase();
          if (t === 'br') s += '\n';
          else if (t === 'script' || t === 'style') continue;
          else { recorrer(h); if (t === 'p' || t === 'div') s += '\n'; }
        }
      }
    };
    recorrer(celda);
    return s.replace(/[ \t\r\f\v ]+/g, ' ').replace(/ *\n */g, '\n').replace(/^\s+|\s+$/g, '');
  }

  function valorDe(celda, texto) {
    const num = celda.getAttribute('x:num') ?? celda.getAttribute('x:Num');
    if (num !== null && num !== '' && Number.isFinite(Number(num))) return Number(num);
    if (/^[-+]?(\d+)(\.\d+)?$/.test(texto) && !/^[-+]?0\d/.test(texto) && texto.length < 16) return Number(texto);
    return texto === '' ? undefined : (texto.length > MAX_TEXTO ? texto.slice(0, MAX_TEXTO) : texto);
  }

  function leer(bytes) {
    if (typeof DOMParser === 'undefined') throw new LibroXlsx.ErrorXlsx('noZip', 'sin DOMParser');
    const doc = new DOMParser().parseFromString(CsvPlanilla.decodificar(bytes), 'text/html');
    const tablas = [...doc.querySelectorAll('table')].filter((t) => !t.parentElement || !t.parentElement.closest('table')).slice(0, MAX_TABLAS);   // sin las anidadas
    if (!tablas.length) throw new LibroXlsx.ErrorXlsx('sinLibro', 'sin tablas');
    const hojas = [];
    tablas.forEach((tabla, k) => {
      // ocupadaHasta[c]: la última fila que toma, en la columna c, un rowspan de arriba (antes, un Set con cada celda: 30 celdas de rowspan y colspan
      // 1000 en 1 KB de HTML eran 30 millones de entradas)
      const filas = [], combinadas = [], ocupadaHasta = new Int32Array(MAX_COLS).fill(-1);
      let f = 0;
      const nfilas = tabla.rows.length;   // un rowspan no pasa de la última fila de la tabla (rowspan="65534" en la primera: una combinada de 65.534 filas)
      for (const tr of tabla.rows) {
        if (f >= MAX_FILAS) break;
        let c = 0;
        const fila = filas[f] || (filas[f] = []);
        for (const celda of tr.cells) {
          while (c < MAX_COLS && ocupadaHasta[c] >= f) c++;
          if (c >= MAX_COLS) break;
          const texto = textoDeCelda(celda);
          const v = valorDe(celda, texto);
          if (v !== undefined) fila[c] = v;
          const cs = Math.min(Math.max(parseInt(celda.getAttribute('colspan'), 10) || 1, 1), MAX_COLS - c), rs = Math.min(Math.max(parseInt(celda.getAttribute('rowspan'), 10) || 1, 1), MAX_FILAS - f, Math.max(1, nfilas - f));
          if (cs > 1 || rs > 1) {
            if (combinadas.length < MAX_COMBINADAS) combinadas.push(LibroXlsx.numeroACol(c) + (f + 1) + ':' + LibroXlsx.numeroACol(c + cs - 1) + (f + rs));
            if (rs > 1) for (let j = c; j < c + cs; j++) if (ocupadaHasta[j] < f + rs - 1) ocupadaHasta[j] = f + rs - 1;
          }
          c += cs;
        }
        f++;
      }
      for (let r = 0; r < filas.length; r++) if (!filas[r]) filas[r] = [];
      hojas.push({ nombre: tabla.getAttribute('title') || tabla.getAttribute('id') || (tablas.length > 1 ? 'Tabla ' + (k + 1) : 'Hoja1'), filas, anchos: [], combinadas });
    });
    return { hojas };
  }

  globalThis.HtmlTablaLector = { parece, leer };
})();

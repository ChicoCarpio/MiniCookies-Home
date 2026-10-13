/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Lector de CSV del motor propio de Excel: texto (o bytes) → el mismo modelo
// que LibroXlsx (una hoja, sus celdas y el estilo por defecto), para que las
// herramientas traten igual un .csv que un .xlsx. Requiere xlsx-estilos.js y
// xlsx-modelo.js. No depende del navegador.
//
// Reglas (las que ya usaban las herramientas con SheetJS):
// - Codificación: UTF-8 (o UTF-16 con BOM); si no es UTF-8 válido, Windows-1252 (el CSV que guarda Excel en español).
// - Separador: el más frecuente entre , ; tabulador | en las primeras líneas, fuera de comillas.
// - Con ";" la coma es decimal y el punto, de miles (Excel en español); con "," al revés. Con tabulador o "|" decide
//   lo que más aparezca en los datos.
// - Solo son número los textos "limpios" (1234, 1.234,50, -3,5); "00123" (códigos, cédulas) y las fechas quedan como texto.
//   Para los números se conserva el texto original (hoja.celdas.textoOriginal) y así se ve "1.234,50" como venía.
//
// Uso: const libro = CsvPlanilla.leer(bytesOTexto, 'Hoja1');   // mismo modelo que LibroXlsx.abrir()
//      CsvPlanilla.decodificar(bytes) / .detectarSeparador(texto) / .detectarComaDecimal(texto, sep) / .interpretarNumero(texto, comaDecimal)

(function () {
  const MAX_FILAS = 1048576, MAX_COLUMNAS = 16384;

  function decodificar(bytes) {
    if (bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) return new TextDecoder('utf-8').decode(bytes.subarray(3));
    if (bytes[0] === 0xFF && bytes[1] === 0xFE) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
    if (bytes[0] === 0xFE && bytes[1] === 0xFF) return new TextDecoder('utf-16be').decode(bytes.subarray(2));
    try { return new TextDecoder('utf-8', { fatal: true }).decode(bytes); } catch (e) { return new TextDecoder('windows-1252').decode(bytes); }
  }

  // Separador más frecuente en las primeras líneas, contando solo fuera de comillas.
  function detectarSeparador(texto) {
    const muestra = texto.slice(0, 20000);
    const cuentas = { ',': 0, ';': 0, '\t': 0, '|': 0 };
    let entreComillas = false, lineas = 0;
    for (const caracter of muestra) {
      if (caracter === '"') entreComillas = !entreComillas;
      else if (!entreComillas && caracter in cuentas) cuentas[caracter]++;
      else if (!entreComillas && caracter === '\n' && ++lineas >= 20) break;
    }
    return Object.keys(cuentas).reduce((a, b) => (cuentas[b] > cuentas[a] ? b : a), ',');
  }

  function detectarComaDecimal(texto, separador) {
    if (separador === ';') return true;
    if (separador === ',') return false;
    const conComa = (texto.match(/\d,\d/g) || []).length;
    const conPunto = (texto.match(/\d\.\d/g) || []).length;
    return conComa > conPunto;
  }

  // El número que representa el texto, o null si no es un número "limpio".
  function interpretarNumero(texto, comaDecimal) {
    const valor = String(texto).trim();
    if (!valor) return null;
    const miles = comaDecimal ? '\\.' : ',';
    const decimal = comaDecimal ? ',' : '\\.';
    if (!new RegExp('^[-+]?(\\d{1,3}(' + miles + '\\d{3})+|\\d+)(' + decimal + '\\d+)?$').test(valor)) return null;
    if (/^0\d/.test(valor.replace(/^[-+]/, ''))) return null;
    const numero = Number(comaDecimal ? valor.replace(/\./g, '').replace(',', '.') : valor.replace(/,/g, ''));
    return Number.isFinite(numero) ? numero : null;
  }

  // RFC 4180 con tolerancia: comillas dobles, "" dentro de comillas, saltos CRLF/LF/CR dentro de un campo entre comillas.
  // Llama a alCampo(fila, columna, texto) por cada campo no vacío; devuelve { filas, columnas, truncado }.
  function recorrerCsv(texto, separador, alCampo) {
    const n = texto.length;
    let fila = 0, col = 0, i = 0, columnas = 0, truncado = false;
    const sep = separador.charCodeAt(0);
    while (i < n) {
      let campo;
      if (texto.charCodeAt(i) === 34) { // campo entre comillas
        let j = i + 1;
        campo = '';
        for (;;) {
          const k = texto.indexOf('"', j);
          if (k < 0) { campo += texto.slice(j); j = n; break; }
          campo += texto.slice(j, k);
          if (texto.charCodeAt(k + 1) === 34) { campo += '"'; j = k + 2; } else { j = k + 1; break; }
        }
        // lo que quede hasta el separador o el fin de línea se agrega tal cual
        let f = j;
        while (f < n) { const c = texto.charCodeAt(f); if (c === sep || c === 10 || c === 13) break; f++; }
        if (f > j) campo += texto.slice(j, f);
        i = f;
      } else {
        let f = i;
        while (f < n) { const c = texto.charCodeAt(f); if (c === sep || c === 10 || c === 13) break; f++; }
        campo = texto.slice(i, f);
        i = f;
      }
      if (campo !== '' && fila < MAX_FILAS && col < MAX_COLUMNAS) alCampo(fila, col, campo);
      else if (campo !== '') truncado = true;
      col++;
      if (col > columnas) columnas = col;
      const c = texto.charCodeAt(i);
      if (c === sep) { i++; if (i >= n) { /* separador al final: campo vacío */ } continue; }
      if (c === 13) i += texto.charCodeAt(i + 1) === 10 ? 2 : 1; else if (c === 10) i++;
      fila++; col = 0;
    }
    return { filas: col > 0 ? fila + 1 : fila, columnas, truncado };
  }

  function leer(origen, nombreHoja) {
    const texto = typeof origen === 'string' ? origen.replace(/^﻿/, '') : decodificar(origen instanceof Uint8Array ? origen : new Uint8Array(origen));
    const separador = detectarSeparador(texto);
    const comaDecimal = detectarComaDecimal(texto, separador);
    const estilos = XlsxEstilos.leer(null, null);
    // Reutiliza la construcción de la hoja de LibroXlsx: una hoja vacía y se le agregan las celdas.
    const hoja = LibroXlsx.hojaVacia();
    const C = hoja.celdas;
    C.textoOriginal = new Map();
    const resumen = recorrerCsv(texto, separador, (f, c, campo) => {
      const numero = interpretarNumero(campo, comaDecimal);
      if (numero !== null) { const i = C.agregar(f, c, LibroXlsx.TIPO.numero, numero, 0); C.textoOriginal.set(i, campo); } else C.agregar(f, c, LibroXlsx.TIPO.texto, campo, 0);
    });
    hoja.dimension = { f0: 0, c0: 0, f1: Math.max(0, resumen.filas - 1), c1: Math.max(0, resumen.columnas - 1) };
    hoja.nombre = nombreHoja || 'Hoja1';
    hoja.indice = 0;
    const libro = {
      zip: null, hojas: [{ nombre: hoja.nombre, id: 1, estado: 'visible', tipo: 'hoja', ruta: null }], nombres: [], fecha1904: false, estilos, textos: [],
      vistaLibro: { hojaActiva: 0, primeraHoja: 0 }, calculo: null, propiedades: {}, relaciones: [], tieneMacros: false,
      separadorCsv: separador, comaDecimalCsv: comaDecimal, truncadoCsv: resumen.truncado, esCsv: true,
      cargarHoja: async () => hoja,
    };
    return libro;
  }

  globalThis.CsvPlanilla = { leer, decodificar, detectarSeparador, detectarComaDecimal, interpretarNumero };
})();

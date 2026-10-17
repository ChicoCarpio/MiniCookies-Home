/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Formatos de hoja de cálculo que no son .xlsx (etapa 9 del motor de Excel, reemplaza a SheetJS): .ods, .xlsb y .xls. Cada uno tiene su lector
// (ods-lector.js, xlsb-lector.js, xls-lector.js, html-tabla-lector.js) y los tres devuelven lo mismo; este módulo elige el lector por el contenido del archivo (no por la
// extensión) y lo pasa a un .xlsx con XlsxEscritor.crearLibro, que es lo que el resto del motor ya sabe leer, mostrar e imprimir.
//
//   await FormatosAntiguos.detectar(bytes)    → 'ods' | 'xlsb' | 'xls' | 'xml' | 'html' | null   ('html': una página con tablas guardada como .xls, que muchos sistemas "exportan a Excel")
//   await FormatosAntiguos.leer(bytes)        → { hojas: [{ nombre, filas, anchos, combinadas }], fecha1904? }   (ver OdsLector)
//   await FormatosAntiguos.aXlsx(bytes)       → Uint8Array con un .xlsx (valores y formato de número; sin fórmulas; con estilos de celda y alto de fila)
//
// Lanza Error (con .codigo = 'cifrado' si tiene contraseña) con un mensaje listo para mostrar. Requiere xlsx-escritor.js y los lectores que se usen.
(function () {
  const bytesIgual = (b, firma) => firma.every((x, k) => b[k] === x);
  const CFB = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

  async function detectar(bytes) {
    if (bytes.length >= 8 && bytesIgual(bytes, CFB)) return 'xls';
    if (bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4B) {
      let zip;
      try { zip = await ZipPropio.loadAsync(bytes); } catch (e) { return null; }
      if (zip.file('xl/workbook.bin')) return 'xlsb';
      if (zip.file('content.xml') && (zip.file('mimetype') || zip.file('META-INF/manifest.xml'))) return 'ods';
      return null;
    }
    if (typeof Xml2003Lector !== 'undefined' && Xml2003Lector.parece(bytes)) return 'xml';   // antes que 'html': su <Table> también parece una tabla HTML
    if (typeof HtmlTablaLector !== 'undefined' && HtmlTablaLector.parece(bytes)) return 'html';
    return null;
  }

  async function leer(bytes) {
    const tipo = await detectar(bytes);
    if (tipo === 'ods') return OdsLector.leer(bytes);
    if (tipo === 'xlsb') return XlsbLector.leer(bytes);
    if (tipo === 'xls') return XlsLector.leer(bytes);
    if (tipo === 'xml') return Xml2003Lector.leer(bytes);
    if (tipo === 'html') return HtmlTablaLector.leer(bytes);
    throw new Error('el archivo no es una hoja de cálculo que se pueda leer');
  }

  async function aXlsx(bytes) {
    const libro = await leer(bytes);
    if (libro.fecha1904) {
      // la base 1904 se pasa a la 1900 (como hace Excel al abrir el libro), salvo en los formatos de solo horas
      for (const h of libro.hojas) for (const fila of h.filas) for (let c = 0; c < fila.length; c++) {
        const v = fila[c];
        if (!v || typeof v !== 'object' || typeof v.valor !== 'number' || !v.formato) continue;
        const codigo = typeof v.formato === 'number' ? XlsxEstilos.FORMATOS_INTEGRADOS[v.formato] : v.formato;   // un número es un formato integrado
        if (codigo && XlsxFormato.esFecha(codigo) && XlsxFormato.tieneCalendario(codigo)) fila[c] = { valor: v.valor + 1462, formato: v.formato, estilo: v.estilo };
      }
    }
    return XlsxEscritor.crearLibro(libro.hojas.map((h) => ({ nombre: h.nombre, filas: h.filas, anchos: h.anchos, combinadas: h.combinadas, estado: h.estado, columnasOcultas: h.columnasOcultas, filasOcultas: h.filasOcultas, altosFila: h.altosFila })));
  }

  globalThis.FormatosAntiguos = { detectar, leer, aXlsx };
})();

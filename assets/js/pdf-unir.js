/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Unir PDF sin interfaz: copia las páginas de uno o varios PDF a un documento
// de pdf-lib. Lo usan "Unir PDF" y los flujos. Supone PDFLib ya cargado
// (vendor/pdf-lib).
//
// Un PDF cifrado falla igual que antes: PDFDocument.load sin opciones lanza el
// error de pdf-lib y quien llama decide qué mostrar.
(function () {
  // Acepta Uint8Array, ArrayBuffer, File o Blob.
  async function aBytes(origen) {
    if (origen && typeof origen.arrayBuffer === 'function') return origen.arrayBuffer();
    return origen;
  }

  // Agrega todas las páginas de un PDF al final de pdfDestino. Devuelve cuántas.
  async function agregarPdf(pdfDestino, bytesOArchivo) {
    if (typeof PDFLib === 'undefined') throw new Error('PDFLib no está cargado');
    const pdfOrigen = await PDFLib.PDFDocument.load(await aBytes(bytesOArchivo));
    const paginas = await pdfDestino.copyPages(pdfOrigen, pdfOrigen.getPageIndices());
    paginas.forEach((pagina) => pdfDestino.addPage(pagina));
    return paginas.length;
  }

  // Une una lista de PDF en el orden dado y devuelve los bytes del resultado.
  async function unir(lista) {
    if (typeof PDFLib === 'undefined') throw new Error('PDFLib no está cargado');
    const pdfFinal = await PDFLib.PDFDocument.create();
    for (const item of lista) await agregarPdf(pdfFinal, item);
    return pdfFinal.save();
  }

  window.PdfUnir = { agregarPdf, unir };
})();

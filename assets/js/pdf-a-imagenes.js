/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Dibujar las páginas de un PDF como imágenes (PNG o JPG), sin interfaz.
// Lo usan la herramienta PDF → Imágenes y los flujos.
//
// Unidad de tamaño: ESCALA de PDF.js (1 = 72 ppp; la herramienta ofrece 1.5, 2 y 3),
// no DPI. Requiere que pdfjs.js ya esté cargado (usa `abrirPdf`); no carga scripts.
//
// PdfAImagenes.renderizar(archivoOBytes, { escala, formato, calidad, paginas }, alAvanzar)
//   archivoOBytes: File/Blob o ArrayBuffer/Uint8Array con el PDF.
//   escala: número (por defecto 2). formato: 'image/png' (por defecto) o 'image/jpeg'.
//   calidad: 0 a 1, solo cuenta en JPG (por defecto 0.92).
//   paginas: arreglo opcional de números de página (desde 1); sin él, todas.
//   alAvanzar(hecha, total): se llama con (0, total) al empezar y tras cada página.
//   Devuelve [{ pagina, blob, ancho, alto }] (pagina desde 1).
(function () {
  async function renderizar(archivoOBytes, opciones, alAvanzar) {
    if (typeof abrirPdf !== 'function') {
      throw new Error('PdfAImagenes: falta abrirPdf (hay que cargar assets/js/pdfjs.js antes).');
    }
    const op = opciones || {};
    const escala = op.escala || 2;
    const formato = op.formato || 'image/png';
    const calidad = op.calidad == null ? 0.92 : op.calidad;

    const bytes = (archivoOBytes && typeof archivoOBytes.arrayBuffer === 'function')
      ? await archivoOBytes.arrayBuffer()
      : archivoOBytes;
    const pdf = await abrirPdf(bytes);
    try {
      const numeros = Array.isArray(op.paginas) && op.paginas.length
        ? op.paginas
        : Array.from({ length: pdf.numPages }, (_, i) => i + 1);
      const resultados = [];
      if (alAvanzar) alAvanzar(0, numeros.length);

      for (let k = 0; k < numeros.length; k++) {
        const numero = numeros[k];
        if (!Number.isInteger(numero) || numero < 1 || numero > pdf.numPages) {
          throw new Error('PdfAImagenes: la página ' + numero + ' no existe (el PDF tiene ' + pdf.numPages + ').');
        }
        const pagina = await pdf.getPage(numero);
        const viewport = pagina.getViewport({ scale: escala });
        const canvas = document.createElement('canvas');
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        await pagina.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;

        const blob = await new Promise((resolve) => canvas.toBlob(resolve, formato, calidad));
        resultados.push({ pagina: numero, blob, ancho: canvas.width, alto: canvas.height });

        // Liberar la memoria de la página y del canvas antes de pasar a la siguiente
        pagina.cleanup();
        canvas.width = 0;
        canvas.height = 0;
        if (alAvanzar) alAvanzar(k + 1, numeros.length);
      }
      return resultados;
    } finally {
      // PDF.js 6 ya no tiene pdf.destroy(): se libera la tarea de carga (con el hilo de PDF.js de este documento).
      pdf.loadingTask.destroy();
    }
  }

  window.PdfAImagenes = { renderizar };
})();

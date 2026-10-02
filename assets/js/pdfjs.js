/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Carga de PDF.js, compartida por todas las herramientas que leen PDFs.
//
// Desde la 4.x PDF.js se publica solo como módulo ES: ya no hay un <script>
// que deje "pdfjsLib" listo al cargar la página. Este archivo lo importa una
// vez, configura el worker y expone:
//
//   abrirPdf(bytes, opciones)  → el PDFDocumentProxy (espera a que cargue)
//   cargarPdfjs()              → el módulo, para OPS, ImageKind, etc.
//   window.pdfjsLib            → el mismo módulo, una vez cargado
//
// Abrir siempre con abrirPdf: así la ruta del WebAssembly (JPEG 2000, JBIG2 y
// perfiles de color, que antes venían dentro de la librería) está en un solo
// lugar. Como es un módulo, las herramientas de PDF no funcionan abiertas
// directo del disco (file://): hay que usar `npm run preview`.
//
// Requiere assets/js/vendor.js antes.
//
// Las rutas van completas (no armadas con una constante) para que el build
// pueda comprobar que existen: ver referencias-vendor.mjs.
(function () {
  let carga = null;

  window.cargarPdfjs = function () {
    if (!carga) {
      carga = import(rutaVendor('pdf.js/6.3.289/pdf.min.mjs')).then((modulo) => {
        modulo.GlobalWorkerOptions.workerSrc = rutaVendor('pdf.js/6.3.289/pdf.worker.min.mjs');
        window.pdfjsLib = modulo;
        return modulo;
      });
      // Si falla, el próximo intento vuelve a probar en vez de quedar roto.
      carga.catch(() => { carga = null; });
    }
    return carga;
  };

  window.abrirPdf = async function (bytes, opciones) {
    const pdfjs = await cargarPdfjs();
    return pdfjs.getDocument(Object.assign({ data: bytes, wasmUrl: rutaVendor('pdf.js/6.3.289/wasm/') }, opciones)).promise;
  };

  // Se empieza a bajar apenas carga la página, no al elegir el archivo.
  cargarPdfjs().catch(() => {});
})();

/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Ruta a una librería de terceros copiada en /vendor/.
//
// Las etiquetas <script src="../vendor/..."> se resuelven solas, pero algunas
// rutas viven DENTRO del JavaScript (el worker de PDF.js, el motor del OCR, la
// fuente de respaldo de Excel → PDF) y esas el navegador las resuelve contra la
// página, no contra el script. Una misma página existe en /herramientas/,
// en /en/herramientas/ y abierta directo del disco, así que no hay una ruta
// relativa que sirva para las tres. Esta sí: se calcula desde la ubicación de
// este archivo, que siempre está en assets/js/.
//
// Uso:  rutaVendor('pdf.js/6.3.289/pdf.worker.min.mjs')
// Tiene que cargarse antes del script que la usa.
(function () {
  const base = new URL('../../vendor/', document.currentScript.src).href;
  window.rutaVendor = (ruta) => base + ruta;
})();

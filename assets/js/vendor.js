/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
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
// Uso:  rutaVendor('pdf.js/3.11.174/pdf.worker.min.js')
// Tiene que cargarse antes del script que la usa.
(function () {
  const base = new URL('../../vendor/', document.currentScript.src).href;
  window.rutaVendor = (ruta) => base + ruta;
})();

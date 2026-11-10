/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Sello «mini 🍪» en gris, chico, en la esquina inferior derecha de la última página de un PDF.
// No va siempre: cada documento lo lleva al azar (PROBABILIDAD), por decisión del dueño del sitio (2026-11-10).
// Supone PDFLib ya cargado (vendor/pdf-lib). Lo usa el escáner de documentos.
//
//   await SelloMiniCookies.quizasSellar(pdf)  → true si lo puso
//   await SelloMiniCookies.sellar(pdf)        → lo pone siempre (pruebas)
(function () {
  const PROBABILIDAD = 1 / 3;     // uno de cada tres documentos, en promedio
  const ANCHO_PT = 56;            // ~2 cm de ancho
  const MARGEN_PT = 8;
  const OPACIDAD = 0.55;
  const IMAGEN = new URL('../../marca-logo/minicookies_sello_gris.png', document.currentScript.src).href;

  let bytes = null;
  async function bytesDelSello() {
    if (!bytes) {
      const r = await fetch(IMAGEN);
      if (!r.ok) throw new Error('sello ' + r.status);
      bytes = new Uint8Array(await r.arrayBuffer());
    }
    return bytes;
  }

  async function sellar(pdf) {
    const paginas = pdf.getPages();
    if (!paginas.length) return false;
    const pagina = paginas[paginas.length - 1];
    const imagen = await pdf.embedPng(await bytesDelSello());
    const { width: W } = pagina.getSize();
    const { x: x0, y: y0 } = pagina.getMediaBox();
    const ancho = Math.min(ANCHO_PT, W * 0.25);
    const alto = ancho * imagen.height / imagen.width;
    pagina.drawImage(imagen, { x: x0 + W - ancho - MARGEN_PT, y: y0 + MARGEN_PT, width: ancho, height: alto, opacity: OPACIDAD });
    return true;
  }

  // Si el sello no se puede cargar, el documento sale igual, sin sello.
  async function quizasSellar(pdf) {
    if (Math.random() >= PROBABILIDAD) return false;
    try { return await sellar(pdf); } catch (e) { return false; }
  }

  window.SelloMiniCookies = { quizasSellar, sellar, PROBABILIDAD };
})();

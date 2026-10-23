/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Word (.docx) → PDF, compartido por "Word ↔ PDF" y "Unir PDF".
//
// Dos pasos, porque Word → PDF muestra el documento antes de descargarlo:
//  - aHtml(archivo): el maquetador propio (docx-maquetador.js) arma el documento con
//    su diseño (hojas, márgenes, encabezados, tablas, dibujos). El contenido pasa
//    por el mismo sanitizador que el resto del sitio.
//  - aPdf(documento): se vuelve a dibujar fuera de pantalla y html-a-pdf.js
//    escribe el PDF, sin imprimir.
//
// Requiere: ZipPropio, XmlLigero, DocxBase, DocxEstilos, PptxTexto (las fuentes con
// los anchos de Office), DocxMaquetador, emf.js + imagenes-office.js,
// seguridad-html.js, pdf-lib y html-a-pdf.js.
// Uso: const bytes = await WordAPdf.aPdf(await WordAPdf.aHtml(archivo));

(function () {
  async function aHtml(archivo) {
    // Las tablas pegadas de Excel (EMF) pasan antes a PNG: el navegador no las dibuja.
    // Un .doc o un .odt (con formatos-word.js en la página) se pasa a .docx primero
    const datos = typeof FormatosWord !== 'undefined' ? await FormatosWord.comoDocx(await archivo.arrayBuffer()) : await archivo.arrayBuffer();
    const docx = await ImagenesOffice.docxConImagenesComunes(datos);
    const { html } = await DocxMaquetador.aHtml(docx);
    return { nombre: archivo.name.replace(/\.(docx|doc|odt)$/i, ''), html: sanitizarHtml(html) };
  }

  // alAvanzar(hoja, total), como en HtmlAPdf.convertir.
  async function aPdf(documento, alAvanzar) {
    const lienzo = HtmlAPdf.crearLienzo();
    try {
      lienzo.innerHTML = documento.html;
      await Promise.all([...lienzo.querySelectorAll('img')].map((img) => img.decode().catch(() => {})));
      await document.fonts.ready;
      // El maquetador arma un <section class="docx"> por hoja de Word, con el
      // tamaño de hoja como ancho/alto mínimo y los márgenes como relleno.
      const paginas = [...lienzo.querySelectorAll('section.docx')].map((seccion) => {
        const e = getComputedStyle(seccion);
        return {
          elemento: seccion,
          anchoPx: parseFloat(e.width),
          altoPx: parseFloat(e.minHeight) || seccion.getBoundingClientRect().height,
          margenArribaPx: parseFloat(e.paddingTop),
          margenAbajoPx: parseFloat(e.paddingBottom),
          // Encabezados y pies repetidos en cada hoja, y numeración de páginas.
          medirContenido: true,
          repetidos: !!seccion.querySelector('[data-repetido]'),
          paginaPortada: seccion.dataset.titulopg === '1',
          parImpar: seccion.dataset.parimpar === '1',
          numeroInicial: seccion.dataset.pginicio !== undefined ? parseInt(seccion.dataset.pginicio, 10) : undefined,
        };
      });
      return await HtmlAPdf.convertir(paginas, { titulo: documento.nombre, alAvanzar });
    } finally {
      lienzo.remove();
    }
  }

  window.WordAPdf = { aHtml, aPdf };
})();

/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Word (.docx) → PDF, compartido por "Word ↔ PDF" y "Unir PDF".
//
// Dos pasos, porque Word → PDF muestra el documento antes de descargarlo:
//  - aHtml(archivo): docx-preview dibuja el documento con su diseño (páginas,
//    tamaños, posiciones). Su CSS es de la propia librería; el contenido pasa
//    por el mismo sanitizador que el resto del sitio.
//  - aPdf(documento): se vuelve a dibujar fuera de pantalla y html-a-pdf.js
//    escribe el PDF, sin imprimir.
//
// Requiere: JSZip, docx-preview (como window.docxPreview: su nombre global,
// "docx", choca con la librería docx), emf.js + imagenes-office.js,
// seguridad-html.js, pdf-lib y html-a-pdf.js.
// Uso: const bytes = await WordAPdf.aPdf(await WordAPdf.aHtml(archivo));

(function () {
  async function aHtml(archivo) {
    const cuerpo = document.createElement('div');
    const estilos = document.createElement('div');
    // Las tablas pegadas de Excel (EMF) pasan antes a PNG: docx-preview no las dibuja.
    const docx = await ImagenesOffice.docxConImagenesComunes(await archivo.arrayBuffer());
    // useBase64URL: por defecto las imágenes salen como blob:, que el
    // sanitizador descarta (y dejaba los recuadros vacíos); como data:image pasan.
    await docxPreview.renderAsync(docx, cuerpo, estilos, { inWrapper: true, useBase64URL: true });
    return {
      nombre: archivo.name.replace(/\.docx$/i, ''),
      html: `<section>${estilos.innerHTML}${sanitizarHtml(cuerpo.innerHTML)}</section>`,
    };
  }

  // alAvanzar(hoja, total), como en HtmlAPdf.convertir.
  async function aPdf(documento, alAvanzar) {
    const lienzo = HtmlAPdf.crearLienzo();
    try {
      lienzo.innerHTML = documento.html;
      await Promise.all([...lienzo.querySelectorAll('img')].map((img) => img.decode().catch(() => {})));
      await document.fonts.ready;
      // docx-preview arma un <section class="docx"> por sección del Word, con el
      // tamaño de hoja como ancho/alto mínimo y los márgenes como relleno.
      const paginas = [...lienzo.querySelectorAll('section.docx')].map((seccion) => {
        const e = getComputedStyle(seccion);
        return {
          elemento: seccion,
          anchoPx: parseFloat(e.width),
          altoPx: parseFloat(e.minHeight) || seccion.getBoundingClientRect().height,
          margenArribaPx: parseFloat(e.paddingTop),
          margenAbajoPx: parseFloat(e.paddingBottom),
        };
      });
      return await HtmlAPdf.convertir(paginas, { titulo: documento.nombre, alAvanzar });
    } finally {
      lienzo.remove();
    }
  }

  window.WordAPdf = { aHtml, aPdf };
})();

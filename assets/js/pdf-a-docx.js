/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// PDF → Word (.docx) conservando la estructura: no solo el texto.
//
// La lectura del PDF (párrafos, estilos, imágenes, líneas) es de
// pdf-estructura.js. Acá se escribe el Word lo más parecido posible al PDF:
// mismo tamaño de hoja y márgenes, espacio entre párrafos medido,
// tabulaciones, imágenes flotantes en su posición exacta (delante del texto,
// sin apartarlo) y un salto por cada página del PDF.
//
// Requiere: pdfjs.js, pdf-estructura.js y la librería docx (window.docx).
// Uso: const blob = await PdfADocx.convertir(arrayBuffer);

(function () {
  const PT_A_TWIP = 20;
  const PT_A_EMU = 12700;
  const PT_A_PX = 96 / 72;

  async function convertir(bytes) {
    const D = window.docx;
    const { paginas, columna, nivelTitulo: nivelDe, margenArriba, margenAbajo, ancho0, alto0 } = await PdfEstructura.analizar(bytes);
    const nivelTitulo = (tamano) => [D.HeadingLevel.HEADING_1, D.HeadingLevel.HEADING_2, D.HeadingLevel.HEADING_3][nivelDe(tamano)] || null;

    // Imagen en su posición exacta de la hoja, delante del texto y sin apartarlo:
    // la posición ya es la del PDF.
    const imagenFlotante = (img) => new D.ImageRun({
      data: img.bytes,
      transformation: { width: Math.round(img.ancho * PT_A_PX), height: Math.round(img.alto * PT_A_PX) },
      floating: {
        horizontalPosition: { relative: D.HorizontalPositionRelativeFrom.PAGE, offset: Math.round(img.x * PT_A_EMU) },
        verticalPosition: { relative: D.VerticalPositionRelativeFrom.PAGE, offset: Math.round(img.y * PT_A_EMU) },
        wrap: { type: D.TextWrappingType.NONE },
        allowOverlap: true,
        behindDocument: false,
      },
    });

    const hijos = [];
    paginas.forEach((pagina, numero) => {
      // Página sin texto (un escaneo, una portada): sus imágenes van tal cual
      // estaban en la hoja, ancladas a un párrafo vacío.
      if (pagina.sinTexto) {
        hijos.push(new D.Paragraph({ pageBreakBefore: numero > 0, children: pagina.imagenes.map(imagenFlotante) }));
        return;
      }
      const { bloques } = pagina;

      let abajoAnterior = margenArriba;
      bloques.forEach((b, i) => {
        const saltoPagina = numero > 0 && i === 0;
        const antes = Math.max(0, Math.min(72, b.arriba - abajoAnterior));
        if (b.tipo === 'imagen') {
          const img = b.img;
          // Que entre en la columna y en la hoja (una página escaneada entera, por ejemplo).
          const escala = Math.min(1, (columna.x1 - columna.x0) / img.ancho, (alto0 - margenArriba - margenAbajo) / img.alto);
          const centro = img.x + img.ancho / 2;
          const anchoCol = columna.x1 - columna.x0;
          const alinea = Math.abs(centro - (columna.x0 + columna.x1) / 2) < anchoCol * 0.1 ? D.AlignmentType.CENTER
            : img.x > columna.x0 + anchoCol * 0.4 ? D.AlignmentType.RIGHT : D.AlignmentType.LEFT;
          hijos.push(new D.Paragraph({
            pageBreakBefore: saltoPagina,
            alignment: alinea,
            spacing: { before: Math.round((b.pagCompleta ? 0 : antes) * PT_A_TWIP), after: 0 },
            children: [new D.ImageRun({ data: img.bytes, transformation: { width: Math.round(img.ancho * escala * PT_A_PX), height: Math.round(img.alto * escala * PT_A_PX) } })],
          }));
          abajoAnterior = img.y + img.alto;
          return;
        }
        const p = b.p;
        const nivel = nivelTitulo(p.tamano);
        const corridas = p.corridas.map((c) => {
          // Los \t se escriben como tabulaciones reales de Word.
          const partes = c.texto.split('\t');
          const contenido = partes.length > 1
            ? { children: partes.flatMap((parte, k) => (k ? [new D.Tab(), parte] : [parte])) }
            : { text: c.texto };
          return new D.TextRun({
            ...contenido,
            bold: c.negrita,
            italics: c.cursiva,
            size: Math.round(c.tamano * 2),
            font: c.familia || undefined,
            color: '000000',
          });
        });
        for (const img of p.flotantes || []) corridas.unshift(imagenFlotante(img));
        const alineacion = { centro: D.AlignmentType.CENTER, derecha: D.AlignmentType.RIGHT, justificado: D.AlignmentType.JUSTIFIED }[p.alineacion] || D.AlignmentType.LEFT;
        hijos.push(new D.Paragraph({
          pageBreakBefore: saltoPagina,
          heading: nivel || undefined,
          alignment: alineacion,
          indent: { left: Math.round(p.izquierda * PT_A_TWIP), firstLine: p.primeraLinea > 1 ? Math.round(p.primeraLinea * PT_A_TWIP) : undefined, hanging: p.primeraLinea < -1 ? Math.round(-p.primeraLinea * PT_A_TWIP) : undefined },
          spacing: {
            before: Math.round(antes * PT_A_TWIP),
            after: 0,
            line: p.interlineado ? Math.round(p.interlineado * PT_A_TWIP) : undefined,
            lineRule: p.interlineado ? D.LineRuleType.AT_LEAST : undefined,
          },
          tabStops: p.tabs.length ? p.tabs.map((x) => ({ type: D.TabStopType.LEFT, position: Math.round(x * PT_A_TWIP) })) : undefined,
          border: p.bordeAbajo ? { bottom: { style: D.BorderStyle.SINGLE, size: 8, color: '000000', space: 1 } } : undefined,
          children: corridas,
        }));
        abajoAnterior = p.abajo;
      });
    });

    const documento = new D.Document({
      sections: [{
        properties: {
          page: {
            size: { width: Math.round(ancho0 * PT_A_TWIP), height: Math.round(alto0 * PT_A_TWIP) },
            margin: {
              top: Math.round(margenArriba * PT_A_TWIP),
              bottom: Math.round(margenAbajo * PT_A_TWIP),
              left: Math.round(columna.x0 * PT_A_TWIP),
              right: Math.round((ancho0 - columna.x1) * PT_A_TWIP),
            },
          },
        },
        children: hijos.length ? hijos : [new D.Paragraph({ children: [] })],
      }],
    });
    return D.Packer.toBlob(documento);
  }

  window.PdfADocx = { convertir };
})();

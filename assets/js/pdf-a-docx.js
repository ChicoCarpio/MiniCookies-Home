/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// PDF → Word (.docx) conservando la estructura: no solo el texto.
//
// La lectura del PDF (párrafos, estilos, imágenes, líneas) es de
// pdf-estructura.js. Acá se escribe el Word lo más parecido posible al PDF:
// mismo tamaño de hoja y márgenes, espacio entre párrafos medido,
// tabulaciones, imágenes flotantes en su posición exacta (delante del texto,
// sin apartarlo) y un salto por cada página del PDF.
//
// Requiere: pdfjs.js, pdf-estructura.js y escritor-docx.js (EscritorDocx).
// Uso: const blob = await PdfADocx.convertir(arrayBuffer);

(function () {
  const PT_A_TWIP = 20;
  const PT_A_EMU = 12700;
  const PT_A_PX = 96 / 72;

  // Los valores por defecto de un documento de Word (interlineado 1,16 y 8 pt después) cambiarían la
  // distribución del PDF: cada párrafo lleva su espaciado explícito, con interlineado simple si el PDF no mide otro.
  const INTERLINEADO_SIMPLE = { linea: 240, regla: 'auto' };

  async function convertir(bytes) {
    const { paginas, columna, nivelTitulo: nivelDe, margenArriba, margenAbajo, ancho0, alto0 } = await PdfEstructura.analizar(bytes);
    const nivelTitulo = (tamano) => [1, 2, 3][nivelDe(tamano)] || null;

    // Imagen en su posición exacta de la hoja, delante del texto y sin apartarlo:
    // la posición ya es la del PDF.
    const imagenFlotante = (img) => ({
      imagen: {
        datos: img.bytes,
        ancho: Math.round(img.ancho * PT_A_PX),
        alto: Math.round(img.alto * PT_A_PX),
        flotante: { x: Math.round(img.x * PT_A_EMU), y: Math.round(img.y * PT_A_EMU) },
      },
    });

    const hijos = [];
    paginas.forEach((pagina, numero) => {
      // Página sin texto (un escaneo, una portada): sus imágenes van tal cual
      // estaban en la hoja, ancladas a un párrafo vacío.
      if (pagina.sinTexto) {
        hijos.push({ tipo: 'parrafo', saltoAntes: numero > 0, espaciado: { antes: 0, despues: 0, ...INTERLINEADO_SIMPLE }, hijos: pagina.imagenes.map(imagenFlotante) });
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
          const alinea = Math.abs(centro - (columna.x0 + columna.x1) / 2) < anchoCol * 0.1 ? 'centro'
            : img.x > columna.x0 + anchoCol * 0.4 ? 'derecha' : 'izquierda';
          hijos.push({
            tipo: 'parrafo',
            saltoAntes: saltoPagina,
            alineacion: alinea,
            espaciado: { antes: Math.round((b.pagCompleta ? 0 : antes) * PT_A_TWIP), despues: 0, ...INTERLINEADO_SIMPLE },
            hijos: [{ imagen: { datos: img.bytes, ancho: Math.round(img.ancho * escala * PT_A_PX), alto: Math.round(img.alto * escala * PT_A_PX) } }],
          });
          abajoAnterior = img.y + img.alto;
          return;
        }
        const p = b.p;
        const nivel = nivelTitulo(p.tamano);
        // Los \t van tal cual: el escritor los escribe como tabulaciones reales de Word.
        const corridas = p.corridas.map((c) => ({
          texto: c.texto,
          negrita: c.negrita,
          cursiva: c.cursiva,
          tamano: Math.max(2, Math.round(c.tamano * 2)), // 1 pt, el mínimo de Word: un 0 dejaba el texto en 12 pt
          fuente: c.familia || undefined,
          color: '000000',
        }));
        for (const img of p.flotantes || []) corridas.unshift(imagenFlotante(img));
        const alineacion = { centro: 'centro', derecha: 'derecha', justificado: 'justificado' }[p.alineacion] || 'izquierda';
        hijos.push({
          tipo: 'parrafo',
          saltoAntes: saltoPagina,
          titulo: nivel || undefined,
          alineacion,
          sangria: { izquierda: Math.round(p.izquierda * PT_A_TWIP), primeraLinea: p.primeraLinea > 1 ? Math.round(p.primeraLinea * PT_A_TWIP) : undefined, colgante: p.primeraLinea < -1 ? Math.round(-p.primeraLinea * PT_A_TWIP) : undefined },
          espaciado: p.interlineado
            ? { antes: Math.round(antes * PT_A_TWIP), despues: 0, linea: Math.round(p.interlineado * PT_A_TWIP), regla: 'minimo' }
            : { antes: Math.round(antes * PT_A_TWIP), despues: 0, ...INTERLINEADO_SIMPLE },
          tabs: p.tabs.length ? p.tabs.map((x) => ({ tipo: 'izquierda', pos: Math.round(x * PT_A_TWIP) })) : undefined,
          bordeAbajo: p.bordeAbajo ? { tam: 8, color: '000000', espacio: 1 } : undefined,
          hijos: corridas,
        });
        abajoAnterior = p.abajo;
      });
    });

    return EscritorDocx.aBlob({
      pagina: {
        ancho: Math.round(ancho0 * PT_A_TWIP),
        alto: Math.round(alto0 * PT_A_TWIP),
        margen: {
          arriba: Math.round(margenArriba * PT_A_TWIP),
          abajo: Math.round(margenAbajo * PT_A_TWIP),
          izquierda: Math.round(columna.x0 * PT_A_TWIP),
          derecha: Math.round((ancho0 - columna.x1) * PT_A_TWIP),
        },
      },
      hijos: hijos.length ? hijos : [{ tipo: 'parrafo', hijos: [] }],
    });
  }

  window.PdfADocx = { convertir };
})();

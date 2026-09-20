// Helper compartido para renderizar diapositivas de un .pptx a imágenes reales
// (no solo texto) usando PptxViewJS, y empaquetarlas como PDF o .zip.
// Requiere que la página ya haya cargado JSZip, Chart.js, PptxViewJS y pdf-lib,
// y que exista un <canvas id="canvasRender"> oculto en el documento.

// Factor de sobre-muestreo: renderizamos los píxeles a esta escala y después
// los mostramos en el PDF a su tamaño físico real — así el PDF queda en buena
// resolución en vez de 1 píxel de canvas = 1 punto de PDF, que se ve borroso.
const ESCALA_RENDERIZADO = 6;

const EMU_POR_PULGADA = 914400; // unidad interna de PowerPoint (English Metric Units)
const PUNTOS_POR_PULGADA = 72;  // unidad de PDF
const PIXELES_BASE_POR_PULGADA = 96; // referencia estándar de pantalla, antes de aplicar ESCALA_RENDERIZADO

// El modo "actual" de PptxViewJS puede devolver un tamaño de diapositiva base
// incorrecto (se observó 50x25 px reales en un archivo de prueba, cuando debía
// ser del orden de 1000x750 — resultaba en un PDF con páginas del tamaño de una
// estampilla). Para no depender de eso, leemos el tamaño real de la diapositiva
// directo del XML del .pptx (ppt/presentation.xml, elemento <p:sldSz>, en EMUs)
// y fijamos el tamaño del canvas nosotros mismos antes de renderizar.
async function obtenerTamanoDiapositivaEmu(archivo) {
  const zip = await JSZip.loadAsync(archivo);
  const archivoPresentacion = zip.file('ppt/presentation.xml');
  if (!archivoPresentacion) throw new Error('No se pudo leer la estructura del archivo .pptx.');
  const xmlTexto = await archivoPresentacion.async('string');
  const doc = new DOMParser().parseFromString(xmlTexto, 'application/xml');
  const sldSz = doc.getElementsByTagName('p:sldSz')[0];
  if (!sldSz) throw new Error('No se pudo determinar el tamaño de las diapositivas.');
  const cx = parseInt(sldSz.getAttribute('cx'), 10);
  const cy = parseInt(sldSz.getAttribute('cy'), 10);
  if (!cx || !cy) throw new Error('Tamaño de diapositiva inválido.');
  return { cx, cy };
}

// Heurística simple: si al muestrear una grilla de puntos del canvas todos dan
// exactamente el mismo color, asumimos que no se dibujó nada reconocible.
function canvasPareceVacio(canvas) {
  const contexto = canvas.getContext('2d');
  const ancho = canvas.width, alto = canvas.height;
  if (ancho === 0 || alto === 0) return true;
  const puntos = 8;
  let primerColor = null;
  for (let i = 0; i < puntos; i++) {
    for (let j = 0; j < puntos; j++) {
      const x = Math.floor((ancho - 1) * (i / (puntos - 1)));
      const y = Math.floor((alto - 1) * (j / (puntos - 1)));
      const [r, g, b] = contexto.getImageData(x, y, 1, 1).data;
      const color = `${r},${g},${b}`;
      if (primerColor === null) primerColor = color;
      else if (color !== primerColor) return false;
    }
  }
  return true;
}

// Devuelve un array de { dataUrl, anchoPuntos, altoPuntos }, una entrada por
// diapositiva. anchoPuntos/altoPuntos salen del tamaño real de diapositiva leído
// del .pptx (no del tamaño que termine teniendo el canvas), así una diapositiva
// mal reportada por la librería no produce páginas de PDF minúsculas.
// Tira una excepción si PptxViewJS falla, si el canvas termina mucho más chico
// de lo esperado, o si ninguna diapositiva produjo contenido visible — en
// cualquiera de esos casos el llamador debería usar el nivel de respaldo
// (extraerTextoPowerPoint + generarPdfDeTexto).
async function renderizarDiapositivasComoImagenes(archivo, canvas) {
  const { cx, cy } = await obtenerTamanoDiapositivaEmu(archivo);
  const anchoPuntos = (cx / EMU_POR_PULGADA) * PUNTOS_POR_PULGADA;
  const altoPuntos = (cy / EMU_POR_PULGADA) * PUNTOS_POR_PULGADA;
  const anchoPxObjetivo = Math.round((cx / EMU_POR_PULGADA) * PIXELES_BASE_POR_PULGADA * ESCALA_RENDERIZADO);
  const altoPxObjetivo = Math.round((cy / EMU_POR_PULGADA) * PIXELES_BASE_POR_PULGADA * ESCALA_RENDERIZADO);

  canvas.width = anchoPxObjetivo;
  canvas.height = altoPxObjetivo;

  const viewer = new PptxViewJS.PPTXViewer({ canvas, slideSizeMode: 'custom' });
  await viewer.loadFile(archivo);
  const totalDiapositivas = viewer.getSlideCount();
  if (!totalDiapositivas || totalDiapositivas === 0) throw new Error('No se detectaron diapositivas.');

  const imagenes = [];
  let vaciasSeguidas = 0;

  for (let i = 0; i < totalDiapositivas; i++) {
    // Algunos modos de PptxViewJS pueden reajustar el tamaño del canvas al
    // renderizar — lo volvemos a fijar en cada diapositiva por las dudas.
    canvas.width = anchoPxObjetivo;
    canvas.height = altoPxObjetivo;

    await viewer.renderSlide(i, canvas, { scale: ESCALA_RENDERIZADO, quality: 'high' });

    // Si después de renderizar el canvas quedó mucho más chico de lo pedido,
    // la librería ignoró el tamaño "custom" (como pasó con "actual") — mejor
    // frenar acá y caer al respaldo de solo texto que arriesgar un PDF con
    // páginas ilegibles.
    if (canvas.width < anchoPxObjetivo * 0.5 || canvas.height < altoPxObjetivo * 0.5) {
      throw new Error('El renderizado devolvió un tamaño de diapositiva inconsistente.');
    }

    if (canvasPareceVacio(canvas)) {
      vaciasSeguidas++;
      if (vaciasSeguidas === totalDiapositivas) throw new Error('El renderizado no produjo contenido visible.');
    }

    // PNG en vez de JPEG: sin artefactos de compresión en el texto/las formas,
    // que es la mayor parte del contenido típico de una diapositiva.
    imagenes.push({ dataUrl: canvas.toDataURL('image/png'), anchoPuntos, altoPuntos });
  }

  if (viewer.destroy) viewer.destroy();
  return imagenes;
}

async function construirPdfDesdeImagenes(imagenes) {
  const { PDFDocument } = PDFLib;
  const pdf = await PDFDocument.create();
  for (const { dataUrl, anchoPuntos, altoPuntos } of imagenes) {
    const bytes = await (await fetch(dataUrl)).arrayBuffer();
    const imagenEmbebida = await pdf.embedPng(bytes);
    const pagina = pdf.addPage([anchoPuntos, altoPuntos]);
    pagina.drawImage(imagenEmbebida, { x: 0, y: 0, width: anchoPuntos, height: altoPuntos });
  }
  return pdf;
}

async function construirZipDesdeImagenes(imagenes) {
  const zip = new JSZip();
  const cantidadDigitos = String(imagenes.length).length;
  for (let i = 0; i < imagenes.length; i++) {
    const bytes = await (await fetch(imagenes[i].dataUrl)).arrayBuffer();
    const numero = String(i + 1).padStart(cantidadDigitos, '0');
    zip.file(`diapositiva-${numero}.png`, bytes);
  }
  return await zip.generateAsync({ type: 'blob' });
}

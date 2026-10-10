/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
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
  if (!sldSz) throw new Error(window.t ? t('ppt-sin-tamano') : 'No se pudo determinar el tamaño de las diapositivas.');
  const cx = parseInt(sldSz.getAttribute('cx'), 10);
  const cy = parseInt(sldSz.getAttribute('cy'), 10);
  if (!cx || !cy) throw new Error(window.t ? t('ppt-tamano-invalido') : 'Tamaño de diapositiva inválido.');
  return { cx, cy };
}

// Si toda la diapositiva es de un solo color, la librería no dibujó nada
// reconocible. Se recorre la imagen entera con un paso de 3 píxeles (a la
// escala de 6, ningún trazo de letra es tan fino): antes se miraba una grilla
// de 8×8 puntos y un título chico sobre fondo blanco caía entre ellos, así que
// una presentación de texto sobre blanco salía como "vacía" y se convertía
// solo como texto. Lee de los píxeles ya sacados del canvas.
function pixelesParecenVacios({ data, width: ancho, height: alto }) {
  if (ancho === 0 || alto === 0) return true;
  const paso = 3;
  const r = data[0], g = data[1], b = data[2];
  for (let y = 0; y < alto; y += paso) {
    for (let k = y * ancho * 4, finFila = k + ancho * 4; k < finFila; k += paso * 4) {
      if (data[k] !== r || data[k + 1] !== g || data[k + 2] !== b) return false;
    }
  }
  return true;
}

// Los píxeles de la diapositiva, sin canal alfa y comprimidos con el deflate
// del navegador (CompressionStream), listos para ir tal cual al PDF como
// imagen FlateDecode. Por qué así (medido el 2026-10-09 con un .pptx de 22
// diapositivas): con PNG, el navegador lo codificaba y pdf-lib lo volvía a
// decodificar y a comprimir, todo en el hilo principal: unos 30 s con la
// página congelada, un tramo de casi 15 s seguidos. El resultado es el mismo
// (sin pérdida), pero se comprime en pedazos y la página sigue respondiendo.
// Sin CompressionStream (navegadores viejos) devuelve null y se usa el PNG.
async function rgbComprimido(pixeles) {
  if (typeof CompressionStream === 'undefined') return null;
  const { data, width, height } = pixeles;
  const total = width * height;
  const rgb = new Uint8Array(total * 3);
  const PIXELES_POR_TANDA = 2000000;
  for (let inicio = 0; inicio < total; inicio += PIXELES_POR_TANDA) {
    const fin = Math.min(total, inicio + PIXELES_POR_TANDA);
    for (let p = inicio, k = inicio * 4, r = inicio * 3; p < fin; p++, k += 4, r += 3) {
      rgb[r] = data[k];
      rgb[r + 1] = data[k + 1];
      rgb[r + 2] = data[k + 2];
    }
    await new Promise((listo) => setTimeout(listo, 0));
  }
  const flujo = new Blob([rgb]).stream().pipeThrough(new CompressionStream('deflate'));
  return new Uint8Array(await new Response(flujo).arrayBuffer());
}

// Devuelve un array con una entrada por diapositiva:
//   { anchoPuntos, altoPuntos, ancho, alto, rgb, png, url }
// anchoPuntos/altoPuntos salen del tamaño real de diapositiva leído del .pptx
// (no del tamaño que termine teniendo el canvas), así una diapositiva mal
// reportada por la librería no produce páginas de PDF minúsculas. rgb: los
// píxeles para el PDF (ver rgbComprimido). png (Blob) y url (para un <img>):
// solo con conImagenes, que es lo que muestran y empaquetan PowerPoint → PDF
// y PowerPoint → imágenes; Unir PDF no los necesita.
// Tira una excepción si PptxViewJS falla, si el canvas termina mucho más chico
// de lo esperado, o si ninguna diapositiva produjo contenido visible — en
// cualquiera de esos casos el llamador debería usar el nivel de respaldo
// (extraerTextoPowerPoint + generarPdfDeTexto).
async function renderizarDiapositivasComoImagenes(archivo, canvas, alAvanzar, opciones = {}) {
  // Motor propio: el predeterminado desde el 2026-10-10 (mejor o igual que
  // PptxViewJS en las 118 diapositivas de npm run comparar-pptx; ver
  // dibujarConMotorPropio). { motor: 'pptxviewjs' } fuerza la librería (para
  // medirla); las páginas que no cargan el motor (Unir PDF) la usan siempre.
  if (opciones.motor !== 'pptxviewjs' && typeof PptxModelo !== 'undefined') {
    try {
      return await dibujarConMotorPropio(archivo, canvas, alAvanzar, opciones);
    } catch (errorPropio) {
      console.warn('El motor propio no pudo con el PowerPoint, se usa PptxViewJS:', errorPropio);
    }
  }
  // assets/js/pptx-preparar.js corrige, sobre una copia, lo que la librería
  // dibuja mal (fondo, SmartArt, autoajuste, interlineado…). Si preparar o
  // dibujar la copia falla, se dibuja el original, como antes.
  if (typeof PptxPreparar !== 'undefined') {
    let preparado = null;
    try {
      preparado = (await PptxPreparar.preparar(archivo)).buffer;
    } catch (errorPreparar) {
      console.warn('No se pudo preparar el PowerPoint, se dibuja el original:', errorPreparar);
    }
    if (preparado) {
      try {
        return await dibujarDiapositivas(archivo, preparado, canvas, alAvanzar, opciones);
      } catch (errorCopia) {
        console.warn('No se pudo dibujar el PowerPoint preparado, se dibuja el original:', errorCopia);
      }
    }
  }
  return dibujarDiapositivas(archivo, archivo, canvas, alAvanzar, opciones);
}

// archivo: el .pptx del usuario (de ahí sale el tamaño); paraDibujar: lo que
// recibe la librería (el original o la copia preparada).
async function dibujarDiapositivas(archivo, paraDibujar, canvas, alAvanzar, { conImagenes = true } = {}) {
  const { cx, cy } = await obtenerTamanoDiapositivaEmu(archivo);
  const anchoPuntos = (cx / EMU_POR_PULGADA) * PUNTOS_POR_PULGADA;
  const altoPuntos = (cy / EMU_POR_PULGADA) * PUNTOS_POR_PULGADA;
  const anchoPxObjetivo = Math.round((cx / EMU_POR_PULGADA) * PIXELES_BASE_POR_PULGADA * ESCALA_RENDERIZADO);
  const altoPxObjetivo = Math.round((cy / EMU_POR_PULGADA) * PIXELES_BASE_POR_PULGADA * ESCALA_RENDERIZADO);

  canvas.width = anchoPxObjetivo;
  canvas.height = altoPxObjetivo;

  // autoChartRerenderDelayMs: 0 — por defecto la librería vuelve a dibujar
  // la diapositiva 1 sobre el mismo canvas 200 ms después (para completar los
  // gráficos). Con el bucle de abajo eso podía mezclar la diapositiva 1 en la
  // captura de otra, y si caía después de destroy() tiraba "No PPTX loaded".
  // El segundo dibujado para los gráficos se hace a mano en cada diapositiva.
  const viewer = new PptxViewJS.PPTXViewer({ canvas, slideSizeMode: 'custom', autoChartRerenderDelayMs: 0 });
  await viewer.loadFile(paraDibujar);
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
    // Segundo dibujado (lo que hacía la librería sola): los gráficos se
    // terminan de armar después del primero.
    await new Promise((listo) => setTimeout(listo, 200));
    canvas.width = anchoPxObjetivo;
    canvas.height = altoPxObjetivo;
    await viewer.renderSlide(i, canvas, { scale: ESCALA_RENDERIZADO, quality: 'high' });

    // Si después de renderizar el canvas quedó mucho más chico de lo pedido,
    // la librería ignoró el tamaño "custom" (como pasó con "actual") — mejor
    // frenar acá y caer al respaldo de solo texto que arriesgar un PDF con
    // páginas ilegibles.
    if (canvas.width < anchoPxObjetivo * 0.5 || canvas.height < altoPxObjetivo * 0.5) {
      throw new Error(window.t ? t('ppt-render-inconsistente') : 'El renderizado devolvió un tamaño de diapositiva inconsistente.');
    }

    const imagen = await capturarDiapositiva(canvas, anchoPuntos, altoPuntos, conImagenes);
    if (imagen.vacia) {
      vaciasSeguidas++;
      if (vaciasSeguidas === totalDiapositivas) throw new Error('El renderizado no produjo contenido visible.');
    }
    imagenes.push(imagen);
    if (alAvanzar) alAvanzar(i + 1, totalDiapositivas);
  }

  if (viewer.destroy) viewer.destroy();
  return imagenes;
}

// Los píxeles del canvas como una entrada de la lista que devuelve
// renderizarDiapositivasComoImagenes. vacia: todo de un solo color.
async function capturarDiapositiva(canvas, anchoPuntos, altoPuntos, conImagenes) {
  const pixeles = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height);
  const rgb = await rgbComprimido(pixeles);
  const imagen = { anchoPuntos, altoPuntos, ancho: pixeles.width, alto: pixeles.height, rgb, png: null, url: null, vacia: pixelesParecenVacios(pixeles) };
  // PNG en vez de JPEG: sin artefactos de compresión en el texto/las formas,
  // que es la mayor parte del contenido típico de una diapositiva. toBlob y
  // no toDataURL: codifica sin trabar la página.
  if (conImagenes || !rgb) {
    imagen.png = await new Promise((listo) => canvas.toBlob(listo, 'image/png'));
    if (conImagenes) imagen.url = URL.createObjectURL(imagen.png);
  }
  return imagen;
}

// Motor propio (assets/js/pptx-modelo.js, pptx-geometria.js, pptx-texto.js,
// pptx-dibujo.js). Las diapositivas que todavía no sabe dibujar (tablas,
// gráficos) las dibuja PptxViewJS con la copia preparada, como hasta ahora.
// Cada imagen lleva motor: 'propio' | 'pptxviewjs'.
async function dibujarConMotorPropio(archivo, canvas, alAvanzar, { conImagenes = true } = {}) {
  const pres = await PptxModelo.abrir(archivo);
  const anchoPuntos = (pres.ancho / EMU_POR_PULGADA) * PUNTOS_POR_PULGADA;
  const altoPuntos = (pres.alto / EMU_POR_PULGADA) * PUNTOS_POR_PULGADA;
  const anchoPx = Math.round((pres.ancho / EMU_POR_PULGADA) * PIXELES_BASE_POR_PULGADA * ESCALA_RENDERIZADO);
  const altoPx = Math.round((pres.alto / EMU_POR_PULGADA) * PIXELES_BASE_POR_PULGADA * ESCALA_RENDERIZADO);
  let visor = null;
  const imagenes = [];
  let vacias = 0;
  for (let i = 0; i < pres.cantidad; i++) {
    canvas.width = anchoPx;
    canvas.height = altoPx;
    let diapositiva = null;
    try {
      diapositiva = await pres.diapositiva(i);
    } catch (errorDiapositiva) {
      console.warn('Diapositiva ' + (i + 1) + ' ilegible para el motor propio:', errorDiapositiva);
    }
    let motor = 'propio';
    let textos = null;
    if (diapositiva && diapositiva.soportada) {
      ({ textos } = await PptxDibujo.dibujar(pres, diapositiva, canvas));
    } else {
      motor = 'pptxviewjs';
      if (!visor) {
        let paraDibujar = archivo;
        if (typeof PptxPreparar !== 'undefined') {
          try { paraDibujar = (await PptxPreparar.preparar(archivo)).buffer; } catch (e) { /* el original */ }
        }
        visor = new PptxViewJS.PPTXViewer({ canvas, slideSizeMode: 'custom', autoChartRerenderDelayMs: 0 });
        await visor.loadFile(paraDibujar);
      }
      for (let vez = 0; vez < 2; vez++) {
        if (vez) await new Promise((listo) => setTimeout(listo, 200));
        canvas.width = anchoPx;
        canvas.height = altoPx;
        await visor.renderSlide(i, canvas, { scale: ESCALA_RENDERIZADO, quality: 'high' });
      }
    }
    const imagen = await capturarDiapositiva(canvas, anchoPuntos, altoPuntos, conImagenes);
    imagen.motor = motor;
    if (textos) imagen.textos = textos;
    if (diapositiva && !diapositiva.soportada) imagen.motivos = diapositiva.motivos;
    if (imagen.vacia) vacias++;
    imagenes.push(imagen);
    if (alAvanzar) alAvanzar(i + 1, pres.cantidad);
  }
  if (visor && visor.destroy) visor.destroy();
  if (pres.cantidad && vacias === pres.cantidad) throw new Error('El renderizado no produjo contenido visible.');
  return imagenes;
}

async function construirPdfDesdeImagenes(imagenes) {
  const { PDFDocument, pushGraphicsState, concatTransformationMatrix, drawObject, popGraphicsState } = PDFLib;
  const pdf = await PDFDocument.create();
  for (const { anchoPuntos, altoPuntos, ancho, alto, rgb, png } of imagenes) {
    const pagina = pdf.addPage([anchoPuntos, altoPuntos]);
    if (!rgb) {
      const imagenEmbebida = await pdf.embedPng(await png.arrayBuffer());
      pagina.drawImage(imagenEmbebida, { x: 0, y: 0, width: anchoPuntos, height: altoPuntos });
      continue;
    }
    // La imagen ya comprimida va directo como XObject: pdf-lib no la toca.
    const flujo = pdf.context.stream(rgb, {
      Type: 'XObject', Subtype: 'Image', Width: ancho, Height: alto,
      ColorSpace: 'DeviceRGB', BitsPerComponent: 8, Filter: 'FlateDecode',
    });
    const nombre = pagina.node.newXObject('Diapositiva', pdf.context.register(flujo));
    pagina.pushOperators(
      pushGraphicsState(),
      concatTransformationMatrix(anchoPuntos, 0, 0, altoPuntos, 0, 0),
      drawObject(nombre),
      popGraphicsState(),
    );
  }
  return pdf;
}

// PowerPoint → PDF con su respaldo: si las diapositivas no se pueden dibujar,
// un PDF con solo su texto (necesita además assets/js/pdf-texto.js). Lo usan
// "PowerPoint → PDF" y "Unir PDF". Devuelve { pdf, imagenes }: el PDFDocument
// de pdf-lib y las imágenes de cada diapositiva, o null si salió solo texto.
// Sin canvas, usa uno propio. alAvanzar(n, total): por diapositiva dibujada.
// opciones.conImagenes: ver renderizarDiapositivasComoImagenes.
async function powerPointAPdf(archivo, canvas, alAvanzar, opciones) {
  try {
    const imagenes = await renderizarDiapositivasComoImagenes(archivo, canvas || document.createElement('canvas'), alAvanzar, opciones);
    return { pdf: await construirPdfDesdeImagenes(imagenes), imagenes };
  } catch (errorVisual) {
    console.warn('No se pudo renderizar el PowerPoint visualmente, se usa solo texto:', errorVisual);
    return { pdf: await generarPdfDeTexto(await extraerTextoPowerPoint(archivo)), imagenes: null };
  }
}

async function construirZipDesdeImagenes(imagenes) {
  const zip = new JSZip();
  const cantidadDigitos = String(imagenes.length).length;
  for (let i = 0; i < imagenes.length; i++) {
    const numero = String(i + 1).padStart(cantidadDigitos, '0');
    zip.file(`diapositiva-${numero}.png`, imagenes[i].png);
  }
  return await zip.generateAsync({ type: 'blob' });
}

/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Helper compartido para dibujar las diapositivas de un .pptx como imágenes
// reales (no solo texto) con el motor propio (pptx-modelo.js, pptx-dibujo.js…),
// y empaquetarlas como PDF o .zip. Requiere que la página ya haya cargado ZipPropio
// y pdf-lib, y que exista un <canvas id="canvasRender"> oculto en el documento.

// Factor de sobre-muestreo: renderizamos los píxeles a esta escala y después
// los mostramos en el PDF a su tamaño físico real — así el PDF queda en buena
// resolución en vez de 1 píxel de canvas = 1 punto de PDF, que se ve borroso.
const ESCALA_RENDERIZADO = 6;

// El motor propio, en el orden en que se carga: la lista única. Unir PDF lo
// carga de acá (archivo-a-pdf.js) y PowerPoint → PDF e → imágenes lo traen en
// <script> en este mismo orden, lo que el build comprueba (motor-pptx.mjs):
// cuando cada página tenía su copia de la lista, Unir PDF se quedó sin el
// motor propio. Necesita además vendor.js y ZipPropio.
const SCRIPTS_MOTOR_PPTX = ['pptx-xml.js', 'vendor/ecma-376-geometrias/5ed-2016/geometrias.js', 'emf.js', 'pptx-geometria.js', 'pptx-estilos-tabla.js', 'pptx-modelo.js', 'pptx-texto.js', 'pptx-grafico-3d.js', 'pptx-grafico-superficie.js', 'pptx-grafico.js', 'pdf-lienzo.js', 'pptx-dibujo.js', 'pptx-visual.js'];

const EMU_POR_PULGADA = 914400; // unidad interna de PowerPoint (English Metric Units)
const PUNTOS_POR_PULGADA = 72;  // unidad de PDF
const PIXELES_BASE_POR_PULGADA = 96; // referencia estándar de pantalla, antes de aplicar ESCALA_RENDERIZADO

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
// anchoPuntos/altoPuntos salen del tamaño real de diapositiva del .pptx. rgb: los
// píxeles para el PDF (ver rgbComprimido). png (Blob) y url (para un <img>):
// solo con conImagenes, que es lo que muestran y empaquetan PowerPoint → PDF
// y PowerPoint → imágenes; Unir PDF no los necesita.
// Tira una excepción si el archivo no se puede leer o si ninguna diapositiva
// produjo contenido visible — en cualquiera de esos casos el llamador debería
// usar el nivel de respaldo (extraerTextoPowerPoint + generarPdfDeTexto).
async function renderizarDiapositivasComoImagenes(archivo, canvas, alAvanzar, opciones = {}) {
  if (typeof PptxModelo === 'undefined') throw new Error('Falta el motor de PowerPoint (SCRIPTS_MOTOR_PPTX).');
  return dibujarConMotorPropio(archivo, canvas, alAvanzar, opciones);
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

// El motor propio (assets/js/pptx-modelo.js, pptx-geometria.js, pptx-texto.js,
// pptx-grafico.js, pptx-dibujo.js) dibuja todo lo que sabe; lo que no (un
// SmartArt sin su dibujo guardado, un objeto sin imagen, un gráfico que no
// puede leer) se omite y queda anotado en imagen.motivos, y la diapositiva sale
// con el resto. Una diapositiva ilegible sale en blanco.
async function dibujarConMotorPropio(archivo, canvas, alAvanzar, { conImagenes = true, paraPdf = false } = {}) {
  const pres = await PptxModelo.abrir(archivo);
  const anchoPuntos = (pres.ancho / EMU_POR_PULGADA) * PUNTOS_POR_PULGADA;
  const altoPuntos = (pres.alto / EMU_POR_PULGADA) * PUNTOS_POR_PULGADA;
  const anchoPx = Math.round((pres.ancho / EMU_POR_PULGADA) * PIXELES_BASE_POR_PULGADA * ESCALA_RENDERIZADO);
  const altoPx = Math.round((pres.alto / EMU_POR_PULGADA) * PIXELES_BASE_POR_PULGADA * ESCALA_RENDERIZADO);
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
    let textos = null, graficos = null, trozosPdf = null, vector = null;
    if (!diapositiva) {
      // Ilegible: en blanco (un canvas nuevo es transparente y el PDF lo guardaría negro).
      const c2d = canvas.getContext('2d');
      c2d.fillStyle = '#ffffff';
      c2d.fillRect(0, 0, canvas.width, canvas.height);
    } else {
      // Para un PDF, la diapositiva se dibuja además sobre un lienzo de PDF
      // (pdf-lienzo.js): formas, relleno, imágenes y texto reales. Si eso falla,
      // la página va como imagen, como siempre.
      if (paraPdf && typeof LienzoPdf !== 'undefined') {
        const lienzo = new LienzoPdf.Lienzo({ ancho: anchoPx, alto: altoPx, anchoPt: anchoPuntos, altoPt: altoPuntos });
        try {
          await PptxDibujo.dibujar(pres, diapositiva, lienzo);
          if (lienzo.fallo) throw lienzo.fallo;
          vector = lienzo;
        } catch (errorVector) {
          console.warn('Diapositiva ' + (i + 1) + ': no se pudo dibujar como vectores, va como imagen:', errorVector);
        }
      }
      if (!vector || conImagenes) ({ textos, graficos, trozosPdf } = await PptxDibujo.dibujar(pres, diapositiva, canvas));
    }
    const imagen = vector && !conImagenes
      ? { anchoPuntos, altoPuntos, ancho: 0, alto: 0, rgb: null, png: null, url: null, vacia: false }
      : await capturarDiapositiva(canvas, anchoPuntos, altoPuntos, conImagenes);
    if (vector) {
      imagen.vector = vector;
      if (!conImagenes) {
        // Si después no se puede escribir la página de vectores, se dibuja como imagen.
        imagen.comoImagen = async () => {
          const c = document.createElement('canvas');
          c.width = anchoPx; c.height = altoPx;
          await PptxDibujo.dibujar(pres, diapositiva, c);
          return capturarDiapositiva(c, anchoPuntos, altoPuntos, false);
        };
      }
    }
    imagen.motor = 'propio';
    if (textos) imagen.textos = textos;
    if (trozosPdf && trozosPdf.length) imagen.trozosPdf = trozosPdf;
    if (graficos && graficos.length) imagen.graficos = graficos;
    if (diapositiva && !diapositiva.soportada) imagen.motivos = diapositiva.motivos;
    else if (!diapositiva) imagen.motivos = ['diapositiva ilegible'];
    if (imagen.vacia) vacias++;
    imagenes.push(imagen);
    if (alAvanzar) alAvanzar(i + 1, pres.cantidad);
  }
  if (pres.cantidad && vacias === pres.cantidad) throw new Error('El renderizado no produjo contenido visible.');
  return imagenes;
}

async function construirPdfDesdeImagenes(imagenes) {
  const { PDFDocument, pushGraphicsState, concatTransformationMatrix, drawObject, popGraphicsState } = PDFLib;
  const pdf = await PDFDocument.create();
  const capa = imagenes.some((i) => i.trozosPdf && !i.vector) ? await prepararCapaTexto(pdf) : null;
  for (const entrada of imagenes) {
    let { anchoPuntos, altoPuntos, ancho, alto, rgb, png } = entrada;
    const { trozosPdf, vector } = entrada;
    // Página de vectores (etapa 4b). Si no se puede escribir, va como imagen (la
    // que ya hay o, si la diapositiva se dibujó solo como vectores, una nueva).
    if (vector) {
      try {
        await vector.aPagina(pdf);
        continue;
      } catch (errorPagina) {
        if (!rgb && !png) {
          if (!entrada.comoImagen) throw errorPagina;
          ({ ancho, alto, rgb, png } = await entrada.comoImagen());
        }
        console.warn('No se pudo escribir la página como vectores, va como imagen:', errorPagina);
      }
    }
    const pagina = pdf.addPage([anchoPuntos, altoPuntos]);
    if (!rgb) {
      const imagenEmbebida = await pdf.embedPng(await png.arrayBuffer());
      pagina.drawImage(imagenEmbebida, { x: 0, y: 0, width: anchoPuntos, height: altoPuntos });
    } else {
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
    if (capa && trozosPdf) {
      try {
        escribirCapaTexto(pagina, capa, trozosPdf, anchoPuntos / ancho, altoPuntos);
      } catch (errorCapa) {
        console.warn('No se pudo escribir el texto seleccionable de una diapositiva:', errorCapa);
      }
    }
  }
  return pdf;
}

// Capa de texto invisible sobre la imagen de cada diapositiva: el PDF se ve
// igual, pero el texto se puede seleccionar, copiar y buscar. Etapa 4a del
// motor de PowerPoint; la 4b dibuja las formas como vectores y el texto visible.
// Una sola fuente (Liberation Sans, la de Arial) basta: el texto no se ve, solo
// importa a qué carácter corresponde cada glifo (ToUnicode, lo escribe pdf-lib)
// y el ancho, que se ajusta por trozo con Tz. Lo que esa fuente no tiene
// (chino, emojis…) sale mal al copiar. Si no se puede cargar fontkit o la
// fuente, devuelve null y el PDF queda solo con imágenes, como antes.
async function prepararCapaTexto(pdf) {
  try {
    await LienzoPdf.cargarFontkit();
    const respuesta = await fetch(rutaVendor('liberation-fonts/2.1.5/LiberationSans-Regular.ttf'));
    if (!respuesta.ok) throw new Error('No se pudo cargar la fuente del texto seleccionable.');
    pdf.registerFontkit(window.fontkit);
    return { fuente: await pdf.embedFont(await respuesta.arrayBuffer(), { subset: true }) };
  } catch (error) {
    console.warn('El PDF no tendrá texto seleccionable:', error);
    return null;
  }
}

// factor: puntos por píxel del canvas. El origen del canvas es arriba a la
// izquierda y el de PDF abajo a la izquierda; un giro horario en el canvas es
// antihorario con el signo cambiado en PDF.
function escribirCapaTexto(pagina, { fuente }, trozos, factor, altoPuntos) {
  const { beginText, endText, setFontAndSize, setTextMatrix, showText, setTextRenderingMode, setCharacterSqueeze, TextRenderingMode } = PDFLib;
  const clave = pagina.node.newFontDictionary(fuente.name, fuente.ref);
  const operadores = [beginText(), setTextRenderingMode(TextRenderingMode.Invisible)];
  for (const t of trozos) {
    const tamano = t.px * factor;
    if (!(tamano > 0.1)) continue;
    let natural = 0;
    try { natural = fuente.widthOfTextAtSize(t.texto, tamano); } catch (e) { /* glifo sin ancho: queda en 100 % */ }
    const estiramiento = natural > 0 && t.ancho > 0 ? Math.min(300, Math.max(30, 100 * t.ancho * factor / natural)) : 100;
    const cos = Math.cos(t.ang), sen = Math.sin(t.ang);
    operadores.push(
      setFontAndSize(clave, tamano),
      setCharacterSqueeze(estiramiento),
      setTextMatrix(cos, -sen, sen, cos, t.x * factor, altoPuntos - t.y * factor),
      showText(fuente.encodeText(t.texto)),
    );
  }
  operadores.push(endText());
  pagina.pushOperators(...operadores);
}

// PowerPoint → PDF con su respaldo: si las diapositivas no se pueden dibujar,
// un PDF con solo su texto (necesita además assets/js/pdf-texto.js). Lo usan
// "PowerPoint → PDF" y "Unir PDF". Devuelve { pdf, imagenes }: el PDFDocument
// de pdf-lib y las imágenes de cada diapositiva, o null si salió solo texto.
// Sin canvas, usa uno propio. alAvanzar(n, total): por diapositiva dibujada.
// opciones.conImagenes: ver renderizarDiapositivasComoImagenes.
async function powerPointAPdf(archivo, canvas, alAvanzar, opciones) {
  try {
    const imagenes = await renderizarDiapositivasComoImagenes(archivo, canvas || document.createElement('canvas'), alAvanzar, { ...opciones, paraPdf: !(opciones && opciones.vectorial === false) });
    return { pdf: await construirPdfDesdeImagenes(imagenes), imagenes };
  } catch (errorVisual) {
    console.warn('No se pudo renderizar el PowerPoint visualmente, se usa solo texto:', errorVisual);
    return { pdf: await generarPdfDeTexto(await extraerTextoPowerPoint(archivo)), imagenes: null };
  }
}

async function construirZipDesdeImagenes(imagenes) {
  const zip = new ZipPropio();
  const cantidadDigitos = String(imagenes.length).length;
  for (let i = 0; i < imagenes.length; i++) {
    const numero = String(i + 1).padStart(cantidadDigitos, '0');
    zip.file(`diapositiva-${numero}.png`, imagenes[i].png);
  }
  return await zip.generateAsync({ type: 'blob' });
}

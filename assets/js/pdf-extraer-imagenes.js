/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Extraer las imágenes embebidas de un PDF (fotos, logos), sin interfaz.
// La usan la herramienta «Extraer imágenes de PDF» y los flujos.
//
//   PdfExtraerImagenes.extraer(archivoOBytes, opciones, alAvanzar)
//     archivoOBytes: File/Blob, ArrayBuffer o Uint8Array con el PDF.
//     opciones: { tiempoMaximoMs: 5000 }  (espera máxima por cada imagen).
//     alAvanzar(pagina, totalPaginas): se llama antes de revisar cada página.
//     → Promise<Array<{ pagina, numero, nombre, blob (PNG), ancho, alto }>>
//       numero cuenta desde 1 en todo el PDF; nombre = pagina-X-imagen-N.png.
//
// Las imágenes repetidas (mismo objeto común o mismo contenido) salen una vez.
// Requiere antes: assets/js/vendor.js y assets/js/pdfjs.js (abrirPdf).
(function () {
  function imagenPdfACanvas(imagenPdf) {
    const { width, height, data, kind, bitmap } = imagenPdf;
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    const contexto = canvas.getContext('2d');
    // PDF.js entrega casi todas las imágenes ya decodificadas como
    // ImageBitmap (sin "data"): alcanza con dibujarlas.
    if (bitmap) {
      contexto.drawImage(bitmap, 0, 0);
      return canvas;
    }
    if (!data) return null;
    const imageData = contexto.createImageData(width, height);
    const salida = imageData.data;

    if (kind === pdfjsLib.ImageKind.RGBA_32BPP) {
      salida.set(data);
    } else if (kind === pdfjsLib.ImageKind.RGB_24BPP) {
      for (let i = 0, j = 0; i < data.length; i += 3, j += 4) {
        salida[j] = data[i];
        salida[j + 1] = data[i + 1];
        salida[j + 2] = data[i + 2];
        salida[j + 3] = 255;
      }
    } else if (kind === pdfjsLib.ImageKind.GRAYSCALE_1BPP) {
      const bytesPorFila = Math.ceil(width / 8);
      for (let y = 0; y < height; y++) {
        for (let x = 0; x < width; x++) {
          const byte = data[y * bytesPorFila + (x >> 3)];
          const bit = (byte >> (7 - (x % 8))) & 1;
          const valor = bit ? 255 : 0;
          const j = (y * width + x) * 4;
          salida[j] = salida[j + 1] = salida[j + 2] = valor;
          salida[j + 3] = 255;
        }
      }
    } else {
      return null;
    }

    contexto.putImageData(imageData, 0, 0);
    return canvas;
  }

  async function extraer(archivoOBytes, opciones, alAvanzar) {
    if (typeof abrirPdf !== 'function') {
      throw new Error('PdfExtraerImagenes necesita assets/js/pdfjs.js (abrirPdf) cargado antes.');
    }
    const { tiempoMaximoMs = 5000 } = opciones || {};
    const bytes = (archivoOBytes && typeof archivoOBytes.arrayBuffer === 'function')
      ? await archivoOBytes.arrayBuffer()
      : archivoOBytes;
    const pdf = await abrirPdf(bytes);
    const resultado = [];
    const imagenesComunesVistas = new Set();
    const huellasVistas = new Set();

    for (let i = 1; i <= pdf.numPages; i++) {
      if (alAvanzar) alAvanzar(i, pdf.numPages);
      const pagina = await pdf.getPage(i);
      const listaOperadores = await pagina.getOperatorList();
      const nombresImagenes = new Set();

      listaOperadores.fnArray.forEach((funcion, indice) => {
        // Una imagen repetida en mosaico llega como paintImageXObjectRepeat,
        // también con el nombre del objeto como primer argumento.
        if (funcion === pdfjsLib.OPS.paintImageXObject || funcion === pdfjsLib.OPS.paintImageXObjectRepeat) {
          const nombre = listaOperadores.argsArray[indice][0];
          nombresImagenes.add(nombre);
        }
      });

      for (const nombre of nombresImagenes) {
        // Las imágenes que se repiten en varias páginas (un logo, un fondo)
        // PDF.js las guarda una sola vez en commonObjs con nombre "g_...":
        // se extraen una sola vez.
        const esComun = nombre.startsWith('g_');
        if (esComun && imagenesComunesVistas.has(nombre)) continue;
        try {
          const imagenPdf = await new Promise((resolve, reject) => {
            const temporizador = setTimeout(() => reject(new Error('tiempo agotado')), tiempoMaximoMs);
            (esComun ? pagina.commonObjs : pagina.objs).get(nombre, (objeto) => {
              clearTimeout(temporizador);
              resolve(objeto);
            });
          });
          if (esComun) imagenesComunesVistas.add(nombre);
          if (!imagenPdf || (!imagenPdf.data && !imagenPdf.bitmap)) continue;

          const canvas = imagenPdfACanvas(imagenPdf);
          if (!canvas) continue;

          const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
          // La misma imagen usada en otra página (antes de que PDF.js la pase
          // a commonObjs) llega con otro nombre: se descarta por contenido.
          const huella = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', await blob.arrayBuffer()))).join(',');
          if (huellasVistas.has(huella)) continue;
          huellasVistas.add(huella);
          const numero = resultado.length + 1;
          resultado.push({
            pagina: i,
            numero,
            nombre: `pagina-${i}-imagen-${numero}.png`,
            blob,
            ancho: canvas.width,
            alto: canvas.height
          });
        } catch (errorImagen) {
          // Se ignora una imagen puntual que no se pudo leer y se sigue con las demás.
          console.warn('No se pudo extraer una imagen:', errorImagen);
        }
      }
    }
    return resultado;
  }

  window.PdfExtraerImagenes = { extraer };
})();

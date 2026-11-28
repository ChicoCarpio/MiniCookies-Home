/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Leer códigos QR y de barras sin interfaz: desde un lienzo, una imagen o un
// PDF. Lo usan "Lector de códigos" y los flujos. Devuelve datos, nunca textos
// para el usuario. Supone ZXing ya cargado (vendor/@zxing/library) y, para
// dePdf, abrirPdf (assets/js/pdfjs.js); no carga nada por su cuenta.
(function () {
  // getBarcodeFormat() devuelve el número del enum; sin traducirlo, la página
  // mostraba "11" en vez de "QR CODE".
  function nombreFormato(valor) {
    if (typeof ZXing === 'undefined') throw new Error('ZXing no está cargado');
    const nombre = ZXing.BarcodeFormat[valor];
    return nombre ? String(nombre).replace(/_/g, ' ') : String(valor);
  }

  // La lectura trabaja siempre sobre los píxeles de un lienzo, igual en la página y en un hilo de Flujos
  // (flujos-hilo.js, donde no hay <img>).
  const sinDom = typeof HTMLImageElement === 'undefined';
  let lectorBase = null;

  // Un intento con ZXing. Devuelve { texto, formato } o null.
  function intentar(fuente, aFondo) {
    if (!lectorBase) lectorBase = new ZXing.MultiFormatReader();
    lectorBase.setHints(aFondo ? new Map([[ZXing.DecodeHintType.TRY_HARDER, true]]) : undefined);
    try {
      const hallado = lectorBase.decodeWithState(new ZXing.BinaryBitmap(new ZXing.HybridBinarizer(fuente)));
      return { texto: hallado.getText(), formato: nombreFormato(hallado.getBarcodeFormat()) };
    } catch (error) {
      if (error instanceof ZXing.Exception) return null; // no encontró ningún código
      throw error;   // otra cosa (el entorno): que se entere quien llama
    }
  }

  // Un cuarto de vuelta a una imagen en grises (un byte por píxel).
  function girar(grises, ancho, alto) {
    const girada = new Uint8ClampedArray(grises.length);
    for (let y = 0; y < alto; y++) {
      for (let x = 0; x < ancho; x++) girada[(ancho - 1 - x) * alto + y] = grises[y * ancho + x];
    }
    return girada;
  }

  // Una copia de la imagen en grises a otra escala (interpolación bilineal).
  function escalar(grises, ancho, alto, factor) {
    const w = Math.round(ancho * factor);
    const h = Math.round(alto * factor);
    const copia = new Uint8ClampedArray(w * h);
    for (let y = 0; y < h; y++) {
      const sy = Math.min(alto - 1, Math.max(0, (y + 0.5) / factor - 0.5));
      const y0 = sy | 0, y1 = Math.min(alto - 1, y0 + 1), b = sy - y0;
      for (let x = 0; x < w; x++) {
        const sx = Math.min(ancho - 1, Math.max(0, (x + 0.5) / factor - 0.5));
        const x0 = sx | 0, x1 = Math.min(ancho - 1, x0 + 1), a = sx - x0;
        copia[y * w + x] = (grises[y0 * ancho + x0] * (1 - a) + grises[y0 * ancho + x1] * a) * (1 - b)
          + (grises[y1 * ancho + x0] * (1 - a) + grises[y1 * ancho + x1] * a) * b;
      }
    }
    return new ZXing.RGBLuminanceSource(copia, w, h);
  }

  // Los códigos de barras y los Micro QR casi no se corrigen solos: buscando a fondo, ZXing «lee» algunos en el
  // texto o la textura de una foto. Un código de verdad se vuelve a leer igual a otras escalas; uno inventado, no.
  const FIABLES = ['QR CODE', 'DATA MATRIX', 'AZTEC', 'PDF 417'];
  function seRepite(hallado, grises, ancho, alto) {
    let veces = 0;
    return [1.15, 0.9, 1.3].some((factor) => {
      const otro = intentar(escalar(grises, ancho, alto, factor), true);
      if (otro && otro.texto === hallado.texto && otro.formato === hallado.formato) veces++;
      return veces === 2;
    });
  }

  // Una pasada sobre la imagen. La búsqueda rápida se acepta tal cual, salvo los Micro QR.
  function pasada(grises, ancho, alto, aFondo) {
    const hallado = intentar(new ZXing.RGBLuminanceSource(grises, ancho, alto), aFondo);
    if (!hallado || !hallado.texto) return null;
    const dudoso = !FIABLES.includes(hallado.formato) && (aFondo || hallado.formato === 'MICRO QR CODE');
    return !dudoso || seRepite(hallado, grises, ancho, alto) ? hallado : null;
  }

  // Primero la búsqueda rápida. Si no encuentra nada, la exhaustiva (TRY_HARDER): la rápida salta filas según
  // el alto de la imagen y pasa por encima de un código chico dentro de una foto grande (el QR de un recibo).
  // La exhaustiva también se prueba de lado, para los códigos de barras verticales; el giro se hace acá porque
  // el de ZXing necesita la página y en un hilo no existe.
  function leerPixeles(lienzo) {
    if (typeof ZXing === 'undefined') throw new Error('ZXing no está cargado');
    const grises = new ZXing.HTMLCanvasElementLuminanceSource(lienzo).getMatrix();
    const { width: ancho, height: alto } = lienzo;
    return pasada(grises, ancho, alto, false)
      || pasada(grises, ancho, alto, true)
      || pasada(girar(grises, ancho, alto), alto, ancho, true);
  }

  // Un lienzo con la imagen del archivo, o null si no se puede abrir.
  async function lienzoDeImagen(archivo) {
    if (sinDom) {
      const mapa = await createImageBitmap(archivo);   // como <img>: con la orientación EXIF aplicada
      const lienzo = new OffscreenCanvas(mapa.width, mapa.height);
      lienzo.getContext('2d').drawImage(mapa, 0, 0);
      mapa.close();
      return lienzo;
    }
    const url = URL.createObjectURL(archivo);
    try {
      const imagen = new Image();   // <img> y no createImageBitmap: también abre SVG
      imagen.src = url;
      await imagen.decode();
      const lienzo = document.createElement('canvas');
      lienzo.width = imagen.naturalWidth;
      lienzo.height = imagen.naturalHeight;
      if (!lienzo.width || !lienzo.height) return null;
      lienzo.getContext('2d', { willReadFrequently: true }).drawImage(imagen, 0, 0);
      return lienzo;
    } catch (error) {
      return null;
    } finally {
      URL.revokeObjectURL(url);
    }
  }

  // Un intento de lectura sobre un lienzo. Devuelve { texto, formato } o null.
  async function deLienzo(lienzo) {
    return leerPixeles(lienzo);
  }

  // Una imagen (File o Blob). Devuelve una lista de 0 o 1 códigos.
  async function deImagen(archivo, origen) {
    const lienzo = await lienzoDeImagen(archivo);
    const hallado = lienzo && leerPixeles(lienzo);
    return hallado ? [{ ...hallado, origen: origen || '' }] : [];
  }

  // Los códigos chicos (un QR de 2 cm en una página A4) no se reconocen al
  // tamaño natural de la página: se prueba con más resolución antes de darla
  // por perdida. Se empieza por la escala barata.
  const ESCALAS_PDF = [2, 3.5];

  // alAvanzar(n, total) se llama antes de cada página; origenDe(n) devuelve el
  // texto de origen de la página n (la página lo traduce). Cada código trae
  // además 'pagina'.
  async function dePdf(archivo, alAvanzar, origenDe) {
    if (typeof ZXing === 'undefined') throw new Error('ZXing no está cargado');
    const bytes = new Uint8Array(await archivo.arrayBuffer());
    const pdf = await abrirPdf(bytes);
    const encontrados = [];
    try {
      for (let n = 1; n <= pdf.numPages; n++) {
        if (alAvanzar) alAvanzar(n, pdf.numPages);
        const pagina = await pdf.getPage(n);
        for (const escala of ESCALAS_PDF) {
          const viewport = pagina.getViewport({ scale: escala });
          const lienzo = document.createElement('canvas');
          lienzo.width = viewport.width;
          lienzo.height = viewport.height;
          await pagina.render({ canvasContext: lienzo.getContext('2d'), viewport }).promise;
          const hallado = await deLienzo(lienzo);
          if (hallado) {
            encontrados.push({
              texto: hallado.texto,
              formato: hallado.formato,
              origen: origenDe ? origenDe(n) : '',
              pagina: n,
            });
            break; // ya está: no hace falta probar más resolución en esta página
          }
        }
      }
    } finally {
      pdf.loadingTask.destroy();   // libera el PDF y su hilo de PDF.js: en un lote de miles no se acumulan
    }
    return encontrados;
  }

  window.LeerCodigos = { deLienzo, deImagen, dePdf, nombreFormato };
})();

/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Leer códigos QR y de barras sin interfaz: desde un lienzo, una imagen o un
// PDF. Lo usan "Lector de códigos" y los flujos. Devuelve datos, nunca textos
// para el usuario. Supone ZXing ya cargado (vendor/@zxing/library) y, para
// dePdf, abrirPdf (assets/js/pdfjs.js); no carga nada por su cuenta.
(function () {
  let lector = null;

  // El lector se crea la primera vez que se usa.
  function obtenerLector() {
    if (typeof ZXing === 'undefined') throw new Error('ZXing no está cargado');
    if (!lector) lector = new ZXing.BrowserMultiFormatReader();
    return lector;
  }

  // getBarcodeFormat() devuelve el número del enum; sin traducirlo, la página
  // mostraba "11" en vez de "QR CODE".
  function nombreFormato(valor) {
    if (typeof ZXing === 'undefined') throw new Error('ZXing no está cargado');
    const nombre = ZXing.BarcodeFormat[valor];
    return nombre ? String(nombre).replace(/_/g, ' ') : String(valor);
  }

  // En un hilo de Flujos (flujos-hilo.js) no hay <img>: ZXing lee los píxeles del lienzo con las mismas piezas que
  // usa por dentro BrowserMultiFormatReader (gris con HTMLCanvasElementLuminanceSource, HybridBinarizer y
  // MultiFormatReader sin opciones), así que encuentra lo mismo.
  const sinDom = typeof HTMLImageElement === 'undefined';
  let lectorBase = null;
  function leerPixeles(lienzo) {
    if (typeof ZXing === 'undefined') throw new Error('ZXing no está cargado');
    if (!lectorBase) {
      lectorBase = new ZXing.MultiFormatReader();
      lectorBase.setHints(undefined);   // como BrowserMultiFormatReader sin opciones
    }
    try {
      const mapa = new ZXing.BinaryBitmap(new ZXing.HybridBinarizer(new ZXing.HTMLCanvasElementLuminanceSource(lienzo)));
      const hallado = lectorBase.decodeWithState(mapa);
      return { texto: hallado.getText(), formato: nombreFormato(hallado.getBarcodeFormat()) };
    } catch (error) {
      if (error instanceof ZXing.Exception) return null; // en ese lienzo no había ningún código
      throw error;   // otra cosa (el entorno): que se entere quien llama
    }
  }

  // Un intento de lectura sobre un lienzo. Devuelve { texto, formato } o null.
  // ZXing no sabe leer un <canvas> directamente (esta versión no trae
  // decodeFromCanvas), así que se le pasa el lienzo como imagen.
  async function deLienzo(lienzo) {
    if (sinDom) return leerPixeles(lienzo);
    const lec = obtenerLector();
    const blob = await new Promise((resolve) => lienzo.toBlob(resolve, 'image/png'));
    if (!blob) return null;
    const url = URL.createObjectURL(blob);
    try {
      const hallado = await lec.decodeFromImageUrl(url);
      return { texto: hallado.getText(), formato: nombreFormato(hallado.getBarcodeFormat()) };
    } catch (error) {
      return null; // en ese lienzo no había ningún código
    } finally {
      URL.revokeObjectURL(url);
    }
  }

  // Una imagen (File o Blob). Devuelve una lista de 0 o 1 códigos.
  async function deImagen(archivo, origen) {
    if (sinDom) {
      const mapa = await createImageBitmap(archivo);   // como <img>: con la orientación EXIF aplicada
      const lienzo = new OffscreenCanvas(mapa.width, mapa.height);
      lienzo.getContext('2d').drawImage(mapa, 0, 0);
      mapa.close();
      const hallado = leerPixeles(lienzo);
      return hallado ? [{ ...hallado, origen: origen || '' }] : [];
    }
    const lec = obtenerLector();
    const url = URL.createObjectURL(archivo);
    try {
      const hallado = await lec.decodeFromImageUrl(url);
      return [{ texto: hallado.getText(), formato: nombreFormato(hallado.getBarcodeFormat()), origen: origen || '' }];
    } catch (error) {
      return [];
    } finally {
      URL.revokeObjectURL(url);
    }
  }

  // Los códigos chicos (un QR de 2 cm en una página A4) no se reconocen al
  // tamaño natural de la página: se prueba con más resolución antes de darla
  // por perdida. Se empieza por la escala barata.
  const ESCALAS_PDF = [2, 3.5];

  // alAvanzar(n, total) se llama antes de cada página; origenDe(n) devuelve el
  // texto de origen de la página n (la página lo traduce). Cada código trae
  // además 'pagina'.
  async function dePdf(archivo, alAvanzar, origenDe) {
    if (!sinDom) obtenerLector();
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

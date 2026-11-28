/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Leer códigos QR y de barras sin interfaz: desde un lienzo, una imagen, un PDF o la cámara. Lo usan "Lector de
// códigos" y los flujos. Devuelve datos, nunca textos para el usuario. Lee con ZXing C++ compilado a WebAssembly:
// supone vendor.js y vendor/zxing-wasm ya cargados y, para dePdf, abrirPdf (assets/js/pdfjs.js); no carga scripts.
(function () {
  // El nombre de formato que ve la persona (y que Flujos guarda). Uno que no esté acá se muestra como viene.
  const NOMBRES = {
    QRCode: 'QR CODE', QRCodeModel1: 'QR CODE', QRCodeModel2: 'QR CODE', MicroQRCode: 'MICRO QR CODE',
    RMQRCode: 'RMQR CODE', DataMatrix: 'DATA MATRIX', Aztec: 'AZTEC', AztecCode: 'AZTEC', AztecRune: 'AZTEC',
    PDF417: 'PDF 417', CompactPDF417: 'PDF 417', MicroPDF417: 'MICRO PDF 417',
    EAN13: 'EAN 13', EAN8: 'EAN 8', UPCA: 'UPC A', UPCE: 'UPC E', ISBN: 'EAN 13',
    Code128: 'CODE 128', Code39: 'CODE 39', Code39Std: 'CODE 39', Code39Ext: 'CODE 39', Code93: 'CODE 93',
    ITF: 'ITF', ITF14: 'ITF', Codabar: 'CODABAR',
    DataBar: 'RSS 14', DataBarOmni: 'RSS 14', DataBarStk: 'RSS 14', DataBarStkOmni: 'RSS 14', DataBarLtd: 'RSS 14',
    DataBarExp: 'RSS EXPANDED', DataBarExpStk: 'RSS EXPANDED',
  };
  const nombreFormato = (valor) => NOMBRES[valor] || String(valor);

  // Los que se corrigen solos: son los únicos que se buscan por trozos de la imagen. Un código de barras cortado
  // por el borde de un trozo podría leerse a medias.
  const MATRICES = ['QRCode', 'DataMatrix', 'Aztec', 'PDF417'];
  // Code39Std: sin el modo extendido, que junta pares de caracteres («+A» pasa a «a») sin que el código lo diga.
  const FORMATOS = [...MATRICES, 'EANUPC', 'Code128', 'Code39Std', 'Code93', 'ITF', 'Codabar', 'DataBar'];

  let preparada = false;
  function libreria() {
    if (typeof ZXingWASM === 'undefined') throw new Error('ZXing no está cargado');
    if (!preparada) {
      // Sin esto la librería baja su motor de una CDN: acá sale siempre del propio sitio.
      const motor = rutaVendor('zxing-wasm/3.1.8/reader/zxing_reader.wasm');
      ZXingWASM.prepareZXingModule({ overrides: { locateFile: (ruta, base) => (ruta.endsWith('.wasm') ? motor : base + ruta) } });
      preparada = true;
    }
    return ZXingWASM;
  }

  // El texto como lo trae impreso el código. La librería da un UPC como el EAN-13 equivalente (con ceros a la
  // izquierda; la forma corta de un UPC-E va aparte, en `extra`) y antepone «01» al número de un DataBar.
  function salida(r) {
    let texto = r.text, formato = r.format;
    if (formato === 'UPCE') {
      try { texto = JSON.parse(r.extra).UPCE || texto; } catch (error) { /* queda la forma larga */ }
    } else if ((formato === 'EAN13' || formato === 'UPCA') && /^0\d{12}$/.test(texto)) {
      texto = texto.slice(1);
      formato = 'UPCA';
    } else if (nombreFormato(formato) === 'RSS 14' && /^01\d{14}$/.test(texto)) {
      texto = texto.slice(2);
    }
    return { texto, formato: nombreFormato(formato) };
  }

  async function buscarEn(pixeles, formatos, binarizador, opciones) {
    const hallados = await libreria().readBarcodes(pixeles, {
      formats: formatos, binarizer: binarizador || 'LocalAverage', textMode: 'Plain', maxNumberOfSymbols: 1,
      tryHarder: true, tryRotate: true, tryInvert: true, tryDownscale: true, ...opciones,
    });
    return hallados.find((r) => r.isValid && r.text) || null;
  }

  // ITF y Codabar no se comprueban solos: con el código inclinado, dos líneas de lectura pueden coincidir en un
  // pedazo («1234» de «12345678901231»), y a veces se «leen» en el texto de una foto. Solo valen si coinciden más.
  const FRAGILES = ['ITF', 'Codabar'];
  const LINEAS_FRAGILES = 8;

  // Un intento sobre unos píxeles (ImageData). Devuelve { texto, formato } o null.
  async function intentar(pixeles, formatos, binarizador) {
    let hallado = await buscarEn(pixeles, formatos, binarizador);
    if (hallado && /^(ITF|Codabar)/.test(hallado.format)) {
      hallado = await buscarEn(pixeles, FRAGILES, binarizador, { minLineCount: LINEAS_FRAGILES })
        || await buscarEn(pixeles, formatos.filter((f) => !FRAGILES.includes(f)), binarizador);
    }
    // Un DataBar extendido lleva campos: se muestran con su número entre paréntesis, «(01)…(3103)…».
    if (hallado && hallado.format.startsWith('DataBarExp')) {
      hallado = await buscarEn(pixeles, [hallado.format], binarizador, { textMode: 'HRI' }) || hallado;
    }
    return hallado ? salida(hallado) : null;
  }

  // Solo QR, sobre unos píxeles ya preparados (los recortes que limpia vestigator.js), con las dos formas de separar
  // blancos y negros.
  const QR = ['QRCodeModel1', 'QRCodeModel2'];
  const qrEn = async (pixeles) => (await intentar(pixeles, QR)) || intentar(pixeles, QR, 'GlobalHistogram');

  // Los píxeles de un lienzo, con lo transparente sobre blanco (como se ve): si no, un QR en un PNG sin fondo
  // quedaría negro sobre negro.
  function pixelesDe(contexto, x, y, ancho, alto) {
    const imagen = contexto.getImageData(x, y, ancho, alto);
    const d = imagen.data;
    for (let i = 3; i < d.length; i += 4) {
      const a = d[i];
      if (a === 255) continue;
      const blanco = 255 - a;
      for (let c = i - 3; c < i; c++) d[c] = (d[c] * a) / 255 + blanco;
      d[i] = 255;
    }
    return imagen;
  }

  // Los lados de los trozos solapados (cada uno avanza medio lado) con que se insiste en una imagen grande.
  const LADOS = [1000, 600, 380];

  // Primero la imagen entera. Si no aparece nada: la entera con otra forma de separar blancos y negros y, después,
  // por trozos solapados, donde un código chico o con mala luz se separa mejor de lo que lo rodea.
  async function leerLienzo(lienzo, insistir) {
    const { width: ancho, height: alto } = lienzo;
    if (!ancho || !alto) return null;
    const contexto = lienzo.getContext('2d', { willReadFrequently: true });
    const entera = pixelesDe(contexto, 0, 0, ancho, alto);
    let hallado = await intentar(entera, FORMATOS);
    if (hallado || !insistir) return hallado;
    hallado = await intentar(entera, FORMATOS, 'GlobalHistogram');
    for (const lado of LADOS) {
      if (hallado || lado >= Math.max(ancho, alto)) continue;
      const w = Math.min(lado, ancho), h = Math.min(lado, alto), paso = lado >> 1;
      for (let y = 0; !hallado; y += paso) {
        const yy = Math.min(y, alto - h);
        for (let x = 0; !hallado; x += paso) {
          const xx = Math.min(x, ancho - w);
          const trozo = pixelesDe(contexto, xx, yy, w, h);
          hallado = await intentar(trozo, MATRICES) || await intentar(trozo, MATRICES, 'GlobalHistogram');
          if (xx >= ancho - w) break;
        }
        if (yy >= alto - h) break;
      }
    }
    return hallado;
  }

  // La lectura trabaja siempre sobre un lienzo, igual en la página y en un hilo de Flujos (flujos-hilo.js, donde
  // no hay <img>).
  const sinDom = typeof HTMLImageElement === 'undefined';

  // Un lienzo con la imagen del archivo, o null si no se puede abrir.
  async function lienzoDeImagen(archivo) {
    if (sinDom) {
      const mapa = await createImageBitmap(archivo);   // como <img>: con la orientación EXIF aplicada
      const lienzo = new OffscreenCanvas(mapa.width, mapa.height);
      lienzo.getContext('2d', { willReadFrequently: true }).drawImage(mapa, 0, 0);
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

  // Un intento de lectura sobre un lienzo, insistiendo si hace falta. Devuelve { texto, formato } o null.
  const deLienzo = (lienzo) => leerLienzo(lienzo, true);

  // Una imagen (File o Blob). Devuelve una lista de 0 o 1 códigos. `aFondo(lienzo)`, si se pasa (esfuerzo alto),
  // se prueba cuando la lectura normal no encuentra nada.
  async function deImagen(archivo, origen, aFondo) {
    const lienzo = await lienzoDeImagen(archivo);
    const hallado = lienzo && (await deLienzo(lienzo) || (aFondo ? await aFondo(lienzo) : null));
    return hallado ? [{ ...hallado, origen: origen || '' }] : [];
  }

  // Los códigos chicos (un QR de 2 cm en una página A4) no se reconocen al
  // tamaño natural de la página: se prueba con más resolución antes de darla
  // por perdida. Se empieza por la escala barata, y solo la última insiste.
  const ESCALAS_PDF = [2, 3.5];

  // alAvanzar(n, total) se llama antes de cada página; origenDe(n) devuelve el
  // texto de origen de la página n (la página lo traduce). Cada código trae
  // además 'pagina'. `aFondo(lienzo)`, si se pasa, recibe la página en su
  // escala mayor cuando la lectura normal no encontró nada en ella.
  async function dePdf(archivo, alAvanzar, origenDe, aFondo) {
    libreria();
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
          await pagina.render({ canvasContext: lienzo.getContext('2d', { willReadFrequently: true }), viewport }).promise;
          const ultima = escala === ESCALAS_PDF[ESCALAS_PDF.length - 1];
          let hallado = await leerLienzo(lienzo, ultima);
          if (!hallado && ultima && aFondo) hallado = await aFondo(lienzo);
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

  // La cámara. Abre la pedida (o la trasera), la muestra en `video` y lee un cuadro tras otro, una sola pasada por
  // cuadro, llamando a alHallar({ texto, formato }) cada vez que hay un código. Devuelve { detener, deviceId }.
  const PAUSA_MS = 100;
  async function camara(video, deviceId, alHallar) {
    // El motor se carga antes de encender la cámara: si falla, se entera quien llama y no queda una cámara que no lee.
    await intentar(new ImageData(8, 8), FORMATOS);
    const flujo = await navigator.mediaDevices.getUserMedia({
      audio: false,
      video: {
        width: { ideal: 1280 }, height: { ideal: 720 },
        ...(deviceId ? { deviceId: { exact: deviceId } } : { facingMode: 'environment' }),
      },
    });
    let activa = true, reloj = 0;
    function detener() {
      activa = false;
      clearTimeout(reloj);
      flujo.getTracks().forEach((pista) => pista.stop());
      if (video.srcObject === flujo) video.srcObject = null;
    }
    try {
      video.srcObject = flujo;
      await video.play();
    } catch (error) {
      detener();
      throw error;
    }
    const lienzo = document.createElement('canvas');
    async function cuadro() {
      if (!activa) return;
      if (video.videoWidth && video.videoHeight) {
        lienzo.width = video.videoWidth;
        lienzo.height = video.videoHeight;
        lienzo.getContext('2d', { willReadFrequently: true }).drawImage(video, 0, 0);
        const hallado = await leerLienzo(lienzo, false);
        if (hallado && activa) alHallar(hallado);
      }
      if (activa) reloj = setTimeout(cuadro, PAUSA_MS);
    }
    cuadro().catch((error) => console.error(error));
    return { detener, deviceId: flujo.getVideoTracks()[0].getSettings().deviceId };
  }

  window.LeerCodigos = { deLienzo, deImagen, dePdf, camara, qrEn, nombreFormato };
})();

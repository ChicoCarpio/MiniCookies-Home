/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// OCR sin interfaz con Librarius, la IA propia de Mini Cookies (ocr-librarius.js): separado de ocr-pdf.html para que lo usen la
// página y Flujos con el mismo código. Supone ya cargados vendor.js, pdf-lib, ocr-librarius.js y, para PDF, pdfjs.js.
//
//   const lector = await Ocr.crearWorker(['spa', 'eng'])          los idiomas (códigos de ocr-pdf.html); rechaza si el navegador no puede
//   const paginas = await Ocr.paginasDeArchivo(archivo, enderezar)    [{ canvas, escala }] (PDF: una por página)
//   await Ocr.recorrerPaginas(archivo, enderezar, async (pagina) => …)   lo mismo de a una página (suelta cada lienzo al terminar)
//   const n = await Ocr.contarPaginas(archivo)
//   const texto = await Ocr.agregarPagina(pdfNuevo, lector, pagina)
//        reconoce la página, la agrega a pdfNuevo (pdf-lib) con el texto invisible encima y devuelve el texto.
//   Ocr.esPdf(archivo)
(function () {
  // Escala a la que se renderiza cada página de un PDF antes del OCR. pdf.js
  // usa 72 puntos por pulgada como base, así que escala 3.5 equivale a ~250 DPI
  // — el OCR reconoce notablemente mejor a partir de ~200-300 DPI que a los
  // ~180 DPI que daba la escala anterior (2.5). Las imágenes sueltas ya vienen
  // con su propia resolución, así que no se les aplica esta escala (ver
  // ANCHO_MINIMO_IMAGEN más abajo para el caso de una foto chica).
  const ESCALA_RENDERIZADO_PDF = 3.5;

  // Si una imagen suelta viene más angosta que esto (fotos de baja resolución,
  // capturas chicas), se agranda antes del OCR — el texto pequeño se reconoce
  // mucho peor cuando la imagen de origen ya es chica.
  const ANCHO_MINIMO_IMAGEN = 1800;

  const esPdf = (archivo) => archivo.type === 'application/pdf' || /\.pdf$/i.test(archivo.name || '');

  // El lector es Librarius (ocr-librarius.js): lee los 32 idiomas, y una mezcla de grupos (p. ej. ruso y polaco) con varios modelos a
  // la vez. Rechaza con un mensaje claro si el navegador no puede correrlo (sin hilos de trabajo o sin WebAssembly). Devuelve un objeto
  // con `terminate()` (el nombre viene de cuando el motor era otro: Flujos y la página lo usan así).
  async function crearWorker(idiomas) {
    if (!window.OcrLibrarius) throw new Error('OcrLibrarius no está cargado');
    const lector = await OcrLibrarius.crear(idiomas);
    return { librarius: lector, terminate: async () => lector.cerrar() };
  }

  function cargarImagen(archivo) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      const url = URL.createObjectURL(archivo);
      img.onload = () => { URL.revokeObjectURL(url); resolve(img); };
      img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('The source image could not be decoded.')); };   // detalleError lo traduce
      img.src = url;
    });
  }

  // --- Enderezado automático (deskew) ---
  // Técnica clásica de "perfil de proyección": si el texto está bien horizontal,
  // sumar los píxeles oscuros fila por fila da un perfil con picos y valles muy
  // marcados (líneas de texto vs. espacios en blanco entre ellas); si la página
  // está torcida, esas franjas se emborronan. Probamos rotar en un rango de
  // ángulos y nos quedamos con el que da el perfil más marcado (mayor varianza).
  const ANGULO_MAXIMO_DESKEW = 8; // no intentar corregir inclinaciones más extremas: es más probable que la detección se equivoque que que el documento esté así de torcido
  const ANGULO_MINIMO_CORRECCION = 0.3; // por debajo de esto, no vale la pena rotar (ruido de la propia medición)

  // Copia en blanco y negro (a un tamaño chico, para que la búsqueda de ángulo
  // sea rápida) usada solo para medir la inclinación, nunca para el resultado final.
  function crearCopiaBinariaChica(canvasOriginal, anchoProxy = 700) {
    const factor = anchoProxy / canvasOriginal.width;
    const altoProxy = Math.max(1, Math.round(canvasOriginal.height * factor));
    const proxy = document.createElement('canvas');
    proxy.width = anchoProxy;
    proxy.height = altoProxy;
    const ctx = proxy.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, anchoProxy, altoProxy);
    ctx.drawImage(canvasOriginal, 0, 0, anchoProxy, altoProxy);

    const imagenData = ctx.getImageData(0, 0, anchoProxy, altoProxy);
    const datos = imagenData.data;
    const n = anchoProxy * altoProxy;
    const grises = new Float32Array(n);
    let suma = 0;
    for (let i = 0, j = 0; i < datos.length; i += 4, j++) {
      const g = 0.299 * datos[i] + 0.587 * datos[i + 1] + 0.114 * datos[i + 2];
      grises[j] = g;
      suma += g;
    }
    const umbral = (suma / n) * 0.85; // más oscuro que ~85% del promedio de la página = "tinta"
    for (let i = 0, j = 0; i < datos.length; i += 4, j++) {
      const v = grises[j] < umbral ? 0 : 255;
      datos[i] = datos[i + 1] = datos[i + 2] = v;
    }
    ctx.putImageData(imagenData, 0, 0);
    return proxy;
  }

  function rotarCanvas(canvasFuente, anguloGrados) {
    const rad = (anguloGrados * Math.PI) / 180;
    const { width: w, height: h } = canvasFuente;
    const nuevoAncho = Math.ceil(Math.abs(w * Math.cos(rad)) + Math.abs(h * Math.sin(rad)));
    const nuevoAlto = Math.ceil(Math.abs(w * Math.sin(rad)) + Math.abs(h * Math.cos(rad)));
    const destino = document.createElement('canvas');
    destino.width = nuevoAncho;
    destino.height = nuevoAlto;
    const ctx = destino.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, nuevoAncho, nuevoAlto);
    ctx.translate(nuevoAncho / 2, nuevoAlto / 2);
    ctx.rotate(rad);
    ctx.drawImage(canvasFuente, -w / 2, -h / 2);
    return destino;
  }

  // Varianza de la cantidad de píxeles "tinta" por fila: cuanto más marcados
  // los picos/valles (texto horizontal bien alineado), mayor la varianza.
  function varianzaPerfilDeFilas(canvasBinario) {
    const { width, height } = canvasBinario;
    const datos = canvasBinario.getContext('2d').getImageData(0, 0, width, height).data;
    const filas = new Float64Array(height);
    for (let y = 0; y < height; y++) {
      let cuenta = 0;
      const base = y * width * 4;
      for (let x = 0; x < width; x++) {
        if (datos[base + x * 4] < 128) cuenta++;
      }
      filas[y] = cuenta;
    }
    let media = 0;
    for (let y = 0; y < height; y++) media += filas[y];
    media /= height;
    let varianza = 0;
    for (let y = 0; y < height; y++) varianza += (filas[y] - media) ** 2;
    return varianza / height;
  }

  // Devuelve el ángulo (en grados) que hay que rotar la página para enderezarla.
  // Primero una pasada gruesa (cada 1°) y después una fina (cada 0.15°)
  // alrededor del mejor resultado de la gruesa, para afinar sin probar
  // cientos de ángulos.
  function estimarInclinacionGrados(canvasOriginal) {
    const binario = crearCopiaBinariaChica(canvasOriginal);

    function mejorAnguloEnRango(desde, hasta, paso) {
      let mejorAngulo = 0;
      let mejorVarianza = -1;
      for (let angulo = desde; angulo <= hasta + 1e-9; angulo += paso) {
        const varianza = varianzaPerfilDeFilas(rotarCanvas(binario, angulo));
        if (varianza > mejorVarianza) {
          mejorVarianza = varianza;
          mejorAngulo = angulo;
        }
      }
      return mejorAngulo;
    }

    const anguloGrueso = mejorAnguloEnRango(-ANGULO_MAXIMO_DESKEW, ANGULO_MAXIMO_DESKEW, 1);
    return mejorAnguloEnRango(anguloGrueso - 0.9, anguloGrueso + 0.9, 0.15);
  }

  // Si detecta inclinación (por encima del mínimo, y dentro del rango que se
  // anima a corregir), devuelve el canvas ya rotado; si no, devuelve el mismo
  // canvas de entrada sin tocar.
  function corregirInclinacionSiHaceFalta(canvas) {
    const angulo = estimarInclinacionGrados(canvas);
    if (Math.abs(angulo) < ANGULO_MINIMO_CORRECCION) return canvas;
    return rotarCanvas(canvas, angulo);
  }

  // Las páginas de un archivo, cada una {canvas, escala}: cada página de un PDF
  // renderizada a ESCALA_RENDERIZADO_PDF, y una imagen suelta dibujada en un
  // canvas a su resolución original (escala 1, ya que no hay "tamaño de página"
  // original que preservar más chico que la propia foto). Si enderezar está
  // activo, cada canvas se endereza antes de devolverlo — el mismo canvas (ya
  // derecho) es el que se usa después tanto para el OCR como para el PDF final.
  //
  // Las páginas salen de a una (generador): una página de PDF a esta escala ocupa unos 25 MB, y tener las de un PDF
  // de 100 páginas todas juntas (como antes) eran más de 2 GB y la pestaña se caía.
  async function* generarPaginas(archivo, enderezar) {
    if (esPdf(archivo)) {
      const pdfjsDoc = await abrirPdf(await archivo.arrayBuffer());
      try {
        for (let i = 1; i <= pdfjsDoc.numPages; i++) {
          const pagina = await pdfjsDoc.getPage(i);
          const viewport = pagina.getViewport({ scale: ESCALA_RENDERIZADO_PDF });
          let canvas = document.createElement('canvas');
          canvas.width = viewport.width;
          canvas.height = viewport.height;
          await pagina.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
          pagina.cleanup();
          if (enderezar) canvas = corregirInclinacionSiHaceFalta(canvas);
          yield { canvas, escala: ESCALA_RENDERIZADO_PDF };
        }
      } finally {
        try { await pdfjsDoc.destroy(); } catch (e) { /* ya cerrado */ }
      }
      return;
    }
    const img = await cargarImagen(archivo);
    // Si la imagen es angosta, se agranda (escala > 1) para que el OCR
    // tenga más píxeles por letra; el punto PDF final se calcula dividiendo
    // por esa misma escala, así la página resultante mide igual que la
    // imagen original, solo con más resolución interna para el OCR/embebido.
    const escala = img.naturalWidth < ANCHO_MINIMO_IMAGEN ? ANCHO_MINIMO_IMAGEN / img.naturalWidth : 1;
    let canvas = document.createElement('canvas');
    canvas.width = Math.round(img.naturalWidth * escala);
    canvas.height = Math.round(img.naturalHeight * escala);
    const ctx = canvas.getContext('2d');
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    if (enderezar) canvas = corregirInclinacionSiHaceFalta(canvas);
    yield { canvas, escala };
  }

  async function paginasDeArchivo(archivo, enderezar) {
    const paginas = [];
    for await (const pagina of generarPaginas(archivo, enderezar)) paginas.push(pagina);
    return paginas;
  }

  // Página por página: alTenerPagina({ canvas, escala }) y, cuando termina, el lienzo se suelta antes de dibujar la
  // siguiente. Para documentos largos, en vez de paginasDeArchivo.
  async function recorrerPaginas(archivo, enderezar, alTenerPagina) {
    for await (const pagina of generarPaginas(archivo, enderezar)) {
      try { await alTenerPagina(pagina); } finally { pagina.canvas.width = pagina.canvas.height = 0; }
    }
  }

  // Cuántas páginas va a dar un archivo (para el avance), sin dibujarlas.
  async function contarPaginas(archivo) {
    if (!esPdf(archivo)) return 1;
    const pdfjsDoc = await abrirPdf(await archivo.arrayBuffer());
    try { return pdfjsDoc.numPages; } finally { try { await pdfjsDoc.destroy(); } catch (e) { /* ya cerrado */ } }
  }

  // La fuente estándar del PDF (Helvetica, WinAnsi) no tiene cirílico ni griego: esas palabras van con Noto Sans (copia propia en
  // vendor/, la misma de html-a-pdf.js), que se carga solo si una página la necesita. Chino, japonés y coreano no tienen fuente en el
  // sitio: esas palabras quedan en el .txt pero no en la capa de texto del PDF.
  const CJK = /[぀-ヿ㐀-鿿豈-﫿가-힯]/;
  let bytesNoto = null, cargaFontkit = null;
  function cargarScriptOcr(src) {
    return new Promise((ok, mal) => {
      const s = document.createElement('script');
      s.src = src; s.onload = ok; s.onerror = () => mal(new Error('No cargó ' + src));
      document.head.appendChild(s);
    });
  }
  async function fuenteUnicode(pdfNuevo) {
    if (!pdfNuevo._fuenteOcrUnicode) {
      if (!window.fontkit) cargaFontkit = cargaFontkit || cargarScriptOcr(rutaVendor('@pdf-lib/fontkit/1.1.1/dist/fontkit.umd.min.js?v=2'));
      await cargaFontkit;
      if (!bytesNoto) bytesNoto = await (await fetch(rutaVendor('noto-sans/2.015/NotoSans-VF.ttf'))).arrayBuffer();
      pdfNuevo.registerFontkit(window.fontkit);
      pdfNuevo._fuenteOcrUnicode = await pdfNuevo.embedFont(bytesNoto, { subset: false });
    }
    return pdfNuevo._fuenteOcrUnicode;
  }

  // Reconoce una página y la agrega a pdfNuevo: la imagen y encima el texto invisible (seleccionable y buscable), una palabra por vez,
  // con la caja repartida dentro de su línea y la inclinación de la línea. Devuelve el texto reconocido.
  async function agregarPagina(pdfNuevo, worker, { canvas, escala }) {
    const r = await worker.librarius.leerPagina(canvas);
    const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/jpeg', 0.9));
    const imagenEmbebida = await pdfNuevo.embedJpg(new Uint8Array(await blob.arrayBuffer()));
    // El punto PDF de la página final es 1/escala respecto al canvas: la imagen queda a mayor densidad que el tamaño de despliegue
    // cuando viene de un PDF (escala > 1); para imágenes sueltas (escala 1) el punto PDF coincide con el píxel, como en «Unir PDF».
    const anchoPagina = canvas.width / escala;
    const altoPagina = canvas.height / escala;
    const paginaNueva = pdfNuevo.addPage([anchoPagina, altoPagina]);
    paginaNueva.drawImage(imagenEmbebida, { x: 0, y: 0, width: anchoPagina, height: altoPagina });
    if (!pdfNuevo._fuenteOcr) pdfNuevo._fuenteOcr = await pdfNuevo.embedFont(PDFLib.StandardFonts.Helvetica);
    for (const p of r.palabras) {
      if (CJK.test(p.texto)) continue;
      let fuente = pdfNuevo._fuenteOcr, texto = p.texto, ancho1;
      try { ancho1 = fuente.widthOfTextAtSize(texto, 1); } catch (e) {
        // Fuera de WinAnsi (cirílico, griego, signos): Noto Sans; si tampoco puede, la letra base o se omite.
        try { fuente = await fuenteUnicode(pdfNuevo); ancho1 = fuente.widthOfTextAtSize(texto, 1); } catch (e2) {
          fuente = pdfNuevo._fuenteOcr;
          texto = texto.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^ -~]/g, '');
          if (!texto) continue;
          try { ancho1 = fuente.widthOfTextAtSize(texto, 1); } catch (e3) { continue; }
        }
      }
      if (!(ancho1 > 0)) continue;
      // Tamaño que hace que el texto invisible mida lo mismo de ancho que la palabra de la imagen (sin pasar de su alto).
      const tamano = Math.min(p.ancho / escala / ancho1, p.alto / escala);
      if (!(tamano > 0)) continue;
      try {
        paginaNueva.drawText(texto, {
          x: p.x / escala, y: altoPagina - p.y / escala, size: tamano, font: fuente,
          rotate: PDFLib.degrees(p.angulo), color: PDFLib.rgb(0, 0, 0), opacity: 0,
        });
      } catch (errorTexto) {
        console.warn('No se pudo dibujar una palabra del OCR:', texto, errorTexto);
      }
    }
    return r.texto || '';
  }

  window.Ocr = { crearWorker, paginasDeArchivo, recorrerPaginas, contarPaginas, agregarPagina, esPdf };
})();

/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// OCR sin interfaz (sobre Tesseract.js, copia propia en vendor/): separado de ocr-pdf.html para que lo usen la
// página y Flujos con el mismo código. Supone ya cargados vendor.js, pdf-lib, Tesseract.js y, para PDF, pdfjs.js.
//
//   const worker = await Ocr.crearWorker(['spa', 'eng'], logger, motor)  el motor y los idiomas (de vendor/); motor 'librarius' = Beta (ocr-librarius.js)
//   const paginas = await Ocr.paginasDeArchivo(archivo, enderezar)    [{ canvas, escala }] (PDF: una por página)
//   await Ocr.recorrerPaginas(archivo, enderezar, async (pagina) => …)   lo mismo de a una página (suelta cada lienzo al terminar)
//   const n = await Ocr.contarPaginas(archivo)
//   const texto = await Ocr.agregarPagina(pdfNuevo, worker, pagina, preproceso)
//        reconoce la página, la agrega a pdfNuevo (pdf-lib) con el texto invisible encima y devuelve el texto.
//        preproceso: 'ninguno' | 'contraste' | 'otsu'
//   Ocr.esPdf(archivo)
(function () {
  // Escala a la que se renderiza cada página de un PDF antes del OCR. pdf.js
  // usa 72 puntos por pulgada como base, así que escala 3.5 equivale a ~250 DPI
  // — Tesseract reconoce notablemente mejor a partir de ~200-300 DPI que a los
  // ~180 DPI que daba la escala anterior (2.5). Las imágenes sueltas ya vienen
  // con su propia resolución, así que no se les aplica esta escala (ver
  // ANCHO_MINIMO_IMAGEN más abajo para el caso de una foto chica).
  const ESCALA_RENDERIZADO_PDF = 3.5;

  // Si una imagen suelta viene más angosta que esto (fotos de baja resolución,
  // capturas chicas), se agranda antes del OCR — Tesseract reconoce mucho peor
  // texto pequeño cuando la imagen de origen ya es chica.
  const ANCHO_MINIMO_IMAGEN = 1800;

  const esPdf = (archivo) => archivo.type === 'application/pdf' || /\.pdf$/i.test(archivo.name || '');

  // Tesseract baja por su cuenta el worker, el motor (WebAssembly) y el
  // diccionario de cada idioma; por defecto de cdn.jsdelivr.net. Se le
  // indican las copias propias de /vendor/ para no depender de esa CDN.
  // motor 'librarius' (Beta): el lector con IA propia (ocr-librarius.js), solo para idiomas latinos; si no puede cargar
  // (navegador sin Workers o WebAssembly, o un idioma que no lee) se usa Tesseract. Devuelve un objeto con `terminate()`
  // en los dos casos; el de Librarius lleva además `motor: 'librarius'` y `librarius` (el lector).
  async function crearWorker(idiomas, logger, motor) {
    if (motor === 'librarius' && window.OcrLibrarius && OcrLibrarius.soportaIdiomas(idiomas)) {
      try {
        const lector = await OcrLibrarius.crear();
        return { motor: 'librarius', librarius: lector, terminate: async () => lector.cerrar() };
      } catch (error) {
        console.warn('Librarius no cargó; se usa Tesseract.', error);
      }
    }
    return Tesseract.createWorker(idiomas, 1, {
      workerPath: rutaVendor('tesseract.js/7.0.0/dist/worker.min.js?v=2'),
      corePath: rutaVendor('tesseract.js-core/7.0.0'),
      langPath: rutaVendor('tesseract-data/4.0.0_best_int'),
      logger: logger || (() => {}),
    });
  }

  // Extrae una lista plana de palabras {text, bbox:{x0,y0,x1,y1}} sin importar
  // si Tesseract.js devuelve blocks/paragraphs/lines/words anidados o data.words directo.
  function palabrasDesdeResultado(data) {
    if (Array.isArray(data.words) && data.words.length) return data.words;
    const palabras = [];
    const bloques = data.blocks || data.paragraphs || [];
    const recorrer = (nodo) => {
      if (!nodo) return;
      if (Array.isArray(nodo.words)) { palabras.push(...nodo.words); return; }
      if (Array.isArray(nodo.paragraphs)) nodo.paragraphs.forEach(recorrer);
      if (Array.isArray(nodo.lines)) nodo.lines.forEach(recorrer);
    };
    bloques.forEach(recorrer);
    return palabras;
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

  // Escala de grises + estiramiento de contraste (histograma) sobre una copia
  // del canvas, solo para que Tesseract reconozca mejor letras sobre fondos con
  // sombras o poco contraste — el canvas original (a color) es el que se
  // incrusta en el PDF final, este solo se usa para el reconocimiento.
  function prepararParaOcr(canvasOriginal) {
    const canvas = document.createElement('canvas');
    canvas.width = canvasOriginal.width;
    canvas.height = canvasOriginal.height;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(canvasOriginal, 0, 0);
    const imagenData = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const datos = imagenData.data;

    const grises = new Uint8ClampedArray(datos.length / 4);
    let minimo = 255;
    let maximo = 0;
    for (let i = 0, j = 0; i < datos.length; i += 4, j++) {
      const g = 0.299 * datos[i] + 0.587 * datos[i + 1] + 0.114 * datos[i + 2];
      grises[j] = g;
      if (g < minimo) minimo = g;
      if (g > maximo) maximo = g;
    }

    const rango = maximo - minimo || 1;
    for (let i = 0, j = 0; i < datos.length; i += 4, j++) {
      const valor = ((grises[j] - minimo) * 255) / rango;
      datos[i] = datos[i + 1] = datos[i + 2] = valor;
    }

    ctx.putImageData(imagenData, 0, 0);
    return canvas;
  }

  // Blanco y negro puro con el método de Otsu: en vez de un umbral fijo (como
  // el 85% del promedio que usa el deskew solo para detectar líneas), calcula
  // el punto de corte que mejor separa dos grupos de píxeles (tinta vs. papel)
  // a partir del histograma real de grises de la imagen — el umbral clásico
  // para binarización de documentos, sin necesidad de ajustarlo a mano.
  function binarizarOtsu(canvasOriginal) {
    const canvas = document.createElement('canvas');
    canvas.width = canvasOriginal.width;
    canvas.height = canvasOriginal.height;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(canvasOriginal, 0, 0);
    const imagenData = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const datos = imagenData.data;
    const totalPixeles = canvas.width * canvas.height;

    const grises = new Uint8ClampedArray(totalPixeles);
    const histograma = new Uint32Array(256);
    for (let i = 0, j = 0; i < datos.length; i += 4, j++) {
      const g = Math.round(0.299 * datos[i] + 0.587 * datos[i + 1] + 0.114 * datos[i + 2]);
      grises[j] = g;
      histograma[g]++;
    }

    let sumaTotal = 0;
    for (let nivel = 0; nivel < 256; nivel++) sumaTotal += nivel * histograma[nivel];

    let pesoFondo = 0;
    let sumaFondo = 0;
    let mejorVarianza = -1;
    let mejorUmbral = 128;
    for (let nivel = 0; nivel < 256; nivel++) {
      pesoFondo += histograma[nivel];
      if (pesoFondo === 0) continue;
      const pesoObjeto = totalPixeles - pesoFondo;
      if (pesoObjeto === 0) break;
      sumaFondo += nivel * histograma[nivel];
      const mediaFondo = sumaFondo / pesoFondo;
      const mediaObjeto = (sumaTotal - sumaFondo) / pesoObjeto;
      const varianzaEntreClases = pesoFondo * pesoObjeto * (mediaFondo - mediaObjeto) ** 2;
      if (varianzaEntreClases > mejorVarianza) {
        mejorVarianza = varianzaEntreClases;
        mejorUmbral = nivel;
      }
    }

    for (let i = 0, j = 0; i < datos.length; i += 4, j++) {
      const v = grises[j] > mejorUmbral ? 255 : 0;
      datos[i] = datos[i + 1] = datos[i + 2] = v;
    }
    ctx.putImageData(imagenData, 0, 0);
    return canvas;
  }

  // Reconoce una página, la agrega a pdfNuevo con la imagen y el texto invisible
  // encima (seleccionable y buscable) y devuelve el texto reconocido.
  async function agregarPagina(pdfNuevo, worker, { canvas, escala }, preproceso) {
    if (worker.motor === 'librarius') return agregarPaginaLibrarius(pdfNuevo, worker, { canvas, escala });
    let canvasParaOcr = canvas;
    if (preproceso === 'contraste') canvasParaOcr = prepararParaOcr(canvas);
    else if (preproceso === 'otsu') canvasParaOcr = binarizarOtsu(canvas);
    const { data } = await worker.recognize(canvasParaOcr, {}, { blocks: true });

    const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/jpeg', 0.9));
    const imagenEmbebida = await pdfNuevo.embedJpg(new Uint8Array(await blob.arrayBuffer()));

    // El punto PDF de la página final es 1/escala respecto al canvas, así la
    // imagen queda a mayor densidad (nítida) que el tamaño de despliegue
    // cuando viene de un PDF (escala > 1); para imágenes sueltas (escala 1)
    // el punto PDF coincide 1:1 con el píxel, igual que en "Unir PDF".
    const anchoPagina = canvas.width / escala;
    const altoPagina = canvas.height / escala;
    const paginaNueva = pdfNuevo.addPage([anchoPagina, altoPagina]);
    paginaNueva.drawImage(imagenEmbebida, { x: 0, y: 0, width: anchoPagina, height: altoPagina });

    for (const palabra of palabrasDesdeResultado(data)) {
      const texto = (palabra.text || '').trim();
      if (!texto || !palabra.bbox) continue;
      const { x0, y0, y1 } = palabra.bbox;
      const altoPalabra = (y1 - y0) / escala;
      if (altoPalabra <= 0) continue;
      try {
        paginaNueva.drawText(texto, {
          x: x0 / escala,
          y: altoPagina - (y1 / escala),
          size: altoPalabra,
          color: PDFLib.rgb(0, 0, 0),
          opacity: 0, // invisible: se ve la imagen, pero el texto queda seleccionable/buscable
        });
      } catch (errorTexto) {
        // Un carácter fuera del set estándar (WinAnsi) puede fallar al dibujarse; se ignora esa palabra sin romper el resto.
        console.warn('No se pudo dibujar una palabra del OCR:', texto, errorTexto);
      }
    }
    return data.text || '';
  }

  // Lo mismo que agregarPagina, con Librarius: la imagen más el texto invisible, una palabra por vez, con la caja
  // repartida dentro de su línea y la inclinación de la línea. El preproceso no se usa (Librarius trabaja en grises).
  async function agregarPaginaLibrarius(pdfNuevo, worker, { canvas, escala }) {
    const r = await worker.librarius.leerPagina(canvas);
    const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/jpeg', 0.9));
    const imagenEmbebida = await pdfNuevo.embedJpg(new Uint8Array(await blob.arrayBuffer()));
    const anchoPagina = canvas.width / escala;
    const altoPagina = canvas.height / escala;
    const paginaNueva = pdfNuevo.addPage([anchoPagina, altoPagina]);
    paginaNueva.drawImage(imagenEmbebida, { x: 0, y: 0, width: anchoPagina, height: altoPagina });
    if (!pdfNuevo._fuenteOcr) pdfNuevo._fuenteOcr = await pdfNuevo.embedFont(PDFLib.StandardFonts.Helvetica);
    const fuente = pdfNuevo._fuenteOcr;
    for (const p of r.palabras) {
      // La fuente estándar solo cubre WinAnsi: lo que no entra (griego, góticas…) se cambia por su letra base o se omite.
      let texto = p.texto;
      let ancho1;
      try { ancho1 = fuente.widthOfTextAtSize(texto, 1); } catch (e) {
        texto = texto.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^ -~]/g, '');
        if (!texto) continue;
        try { ancho1 = fuente.widthOfTextAtSize(texto, 1); } catch (e2) { continue; }
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

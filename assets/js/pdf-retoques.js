/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Retoques de PDF sin interfaz, sobre pdf-lib: separados de sus páginas para que los usen ellas y Flujos con el
// mismo código. Supone ya cargados pdf-lib y miniaturas.js (marcoVisualPagina).
//
//   PdfRetoques.rotar(pdf, grados)                                           rotar-pdf.html
//   PdfRetoques.numerar(pdf, { inicio, posicion })                           numerar-paginas.html
//   await PdfRetoques.marcaDeAgua(pdf, { texto, tamanoMaximo, opacidad, repeticiones })   marca-de-agua.html
//   await PdfRetoques.optimizar(bytes, { calidad, ladoMaximo }, alAvanzar, alGuardar)   optimizar-pdf.html
//        -> { bytes, procesadas, sinMejora }  (sinMejora: devuelve los bytes originales)
//   PdfRetoques.calcularFilasColumnas(cantidad) -> { filas, columnas }   cuadrícula de la marca de agua
//   PdfRetoques.indicesDeRango("1-3, 5", max) -> [0,1,2,4]; await extraerPaginas(pdf, indices) -> PDFDocument   extraer-paginas.html
//   await PdfRetoques.dividir(pdf) -> [bytes por página]                     dividir-pdf.html
//   PdfRetoques.ponerMetadatos(pdf, { titulo, autor, asunto, palabras })     editar-metadatos-pdf.html
// `pdf` es un PDFDocument de pdf-lib ya abierto; quien llama lo abre y lo guarda.
(function () {
  // ---------- Rotar ----------
  function rotar(pdf, grados) {
    pdf.getPages().forEach((pagina) => {
      pagina.setRotation(PDFLib.degrees(pagina.getRotation().angle + grados));
    });
  }

  // ---------- Numerar ----------
  // posicion: 'bottom-left' | 'bottom-center' | 'bottom-right' | 'top-left' | 'top-center' | 'top-right'
  function calcularX(posicion, width, anchoTexto, margen) {
    if (posicion.endsWith('-left')) return margen;
    if (posicion.endsWith('-center')) return width / 2 - anchoTexto / 2;
    return width - anchoTexto - margen;
  }

  function calcularY(posicion, height, margenInferior, margenSuperior) {
    return posicion.startsWith('top-') ? height - margenSuperior : margenInferior;
  }

  async function numerar(pdf, { inicio = 1, posicion = 'bottom-center' } = {}) {
    const { rgb, StandardFonts, degrees } = PDFLib;
    const fuente = await pdf.embedFont(StandardFonts.Helvetica);
    const tamanoFuente = 12;
    const margen = 30;
    pdf.getPages().forEach((pagina, indice) => {
      // Posición calculada sobre la página tal como se ve (con /Rotate).
      const marco = marcoVisualPagina(pagina);
      const texto = String(inicio + indice);
      const anchoTexto = fuente.widthOfTextAtSize(texto, tamanoFuente);
      const { x, y } = marco.aPagina(
        calcularX(posicion, marco.ancho, anchoTexto, margen),
        calcularY(posicion, marco.alto, 20, margen)
      );
      pagina.drawText(texto, { x, y, rotate: degrees(marco.rotacion), size: tamanoFuente, font: fuente, color: rgb(0.2, 0.2, 0.2) });
    });
  }

  // ---------- Marca de agua ----------
  // Reparte "cantidad" copias en una cuadrícula lo más cuadrada posible
  // (ej. 4 -> 2x2, 6 -> 3x2, 1 -> 1x1) para distribuirlas parejo por la página.
  function calcularFilasColumnas(cantidad) {
    const columnas = Math.ceil(Math.sqrt(cantidad));
    const filas = Math.ceil(cantidad / columnas);
    return { filas, columnas };
  }

  // Encuentra el tamaño de letra más grande (hasta tamanoMaximo) tal que el
  // texto, ya rotado 45°, quepa dentro de la página — para que un mensaje largo
  // se achique en vez de salirse de los bordes.
  function calcularTamanoAjustado(fuente, texto, anchoPagina, altoPagina, tamanoMaximo) {
    const margen = 0.9; // deja un 10% de aire
    let tamano = tamanoMaximo;
    while (tamano > 6) {
      const anchoTexto = fuente.widthOfTextAtSize(texto, tamano);
      const anchoRotado = anchoTexto * Math.SQRT1_2 + tamano * Math.SQRT1_2;
      const altoRotado = anchoTexto * Math.SQRT1_2 + tamano * Math.SQRT1_2;
      if (anchoRotado <= anchoPagina * margen && altoRotado <= altoPagina * margen) return tamano;
      tamano -= 2;
    }
    return tamano;
  }

  async function marcaDeAgua(pdf, { texto, tamanoMaximo = 60, opacidad = 0.3, repeticiones = 1 }) {
    const { rgb, StandardFonts, degrees } = PDFLib;
    const fuente = await pdf.embedFont(StandardFonts.HelveticaBold);
    const { filas, columnas } = calcularFilasColumnas(repeticiones);
    const angulo = 45 * Math.PI / 180;
    pdf.getPages().forEach((pagina) => {
      // Se trabaja sobre la página tal como se ve (con /Rotate): si no, en
      // un PDF escaneado o rotado la marca quedaba girada y descentrada.
      const marco = marcoVisualPagina(pagina);
      const anchoCelda = marco.ancho / columnas;
      const altoCelda = marco.alto / filas;
      const tamanoFuente = calcularTamanoAjustado(fuente, texto, anchoCelda, altoCelda, tamanoMaximo);
      const anchoTexto = fuente.widthOfTextAtSize(texto, tamanoFuente);
      let contador = 0;
      for (let f = 0; f < filas && contador < repeticiones; f++) {
        for (let c = 0; c < columnas && contador < repeticiones; c++) {
          const centroX = anchoCelda * c + anchoCelda / 2;
          const centroY = altoCelda * f + altoCelda / 2;
          // drawText gira desde el inicio de la línea de base: se retrocede
          // medio ancho a lo largo del texto y media altura de mayúscula
          // perpendicular a él, para que el texto quede centrado en la celda.
          const altoLetra = fuente.heightAtSize(tamanoFuente, { descender: false }) * 0.7;
          const origen = marco.aPagina(
            centroX - (anchoTexto / 2) * Math.cos(angulo) + (altoLetra / 2) * Math.sin(angulo),
            centroY - (anchoTexto / 2) * Math.sin(angulo) - (altoLetra / 2) * Math.cos(angulo)
          );
          pagina.drawText(texto, {
            x: origen.x, y: origen.y, size: tamanoFuente, font: fuente,
            color: rgb(0.5, 0.5, 0.5), opacity: opacidad, rotate: degrees(45 + marco.rotacion),
          });
          contador++;
        }
      }
    });
  }

  // ---------- Optimizar (recomprimir las fotos JPEG a color) ----------
  function esEspacioRgb(pdf, espacioColor) {
    const { PDFName, PDFArray } = PDFLib;
    if (!espacioColor) return false;
    const valor = pdf.context.lookup(espacioColor);
    if (valor && valor.toString() === '/DeviceRGB') return true;
    if (valor instanceof PDFArray && valor.size() === 2 && valor.get(0).toString() === '/ICCBased') {
      const perfil = pdf.context.lookup(valor.get(1));
      const canales = perfil && perfil.dict ? perfil.dict.get(PDFName.of('N')) : null;
      return Boolean(canales) && canales.toString() === '3';
    }
    return false;
  }

  function decodificarImagenJpeg(bytes) {
    return new Promise((resolve, reject) => {
      const url = URL.createObjectURL(new Blob([bytes], { type: 'image/jpeg' }));
      const img = new Image();
      img.onload = () => { URL.revokeObjectURL(url); resolve(img); };
      img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('JPEG inválido')); };
      img.src = url;
    });
  }

  // alAvanzar(n): se llama antes de procesar la foto n (1, 2, …); alGuardar(): antes de guardar el PDF.
  async function optimizar(bytesOriginales, { calidad = 0.7, ladoMaximo = 0 } = {}, alAvanzar = () => {}, alGuardar = () => {}) {
    const { PDFDocument, PDFName, PDFNumber, PDFRawStream } = PDFLib;
    const pesoOriginal = bytesOriginales.byteLength;
    const pdf = await PDFDocument.load(bytesOriginales, { updateMetadata: false });
    let procesadas = 0;
    for (const [ref, objeto] of pdf.context.enumerateIndirectObjects()) {
      if (!(objeto instanceof PDFRawStream)) continue;
      const dict = objeto.dict;
      const subtipo = dict.get(PDFName.of('Subtype'));
      if (!subtipo || subtipo.toString() !== '/Image') continue;
      const filtro = dict.get(PDFName.of('Filter'));
      if (!filtro || filtro.toString() !== '/DCTDecode') continue;
      // Color RGB: /DeviceRGB o un perfil ICC de 3 canales (muy común en
      // fotos de cámaras y escáneres). Gris, CMYK e indexadas se dejan igual:
      // el canvas siempre devuelve RGB y cambiarían de color.
      if (!esEspacioRgb(pdf, dict.get(PDFName.of('ColorSpace')))) continue;
      // Un /Decode no estándar invierte o reescala los colores al dibujar.
      if (dict.get(PDFName.of('Decode'))) continue;
      try {
        alAvanzar(procesadas + 1);
        const bytesImagen = objeto.getContents();
        const img = await decodificarImagenJpeg(bytesImagen);
        // Reducción opcional: el lado más largo queda en "ladoMaximo" px.
        // La foto ocupa el mismo lugar en la página (el tamaño en la página
        // lo define el contenido, no los píxeles), solo con menos detalle.
        const escala = ladoMaximo > 0 ? Math.min(1, ladoMaximo / Math.max(img.naturalWidth, img.naturalHeight)) : 1;
        const canvas = document.createElement('canvas');
        canvas.width = Math.max(1, Math.round(img.naturalWidth * escala));
        canvas.height = Math.max(1, Math.round(img.naturalHeight * escala));
        canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
        const blobNuevo = await new Promise((resolve) => canvas.toBlob(resolve, 'image/jpeg', calidad));
        const bytesNuevos = new Uint8Array(await blobNuevo.arrayBuffer());
        if (bytesNuevos.length < bytesImagen.length) {
          dict.set(PDFName.of('Width'), PDFNumber.of(canvas.width));
          dict.set(PDFName.of('Height'), PDFNumber.of(canvas.height));
          dict.set(PDFName.of('BitsPerComponent'), PDFNumber.of(8));
          pdf.context.assign(ref, PDFRawStream.of(dict, bytesNuevos));
          procesadas++;
        }
      } catch (errorImagen) {
        console.warn('Se omitió una imagen que no se pudo optimizar:', errorImagen);
      }
    }
    alGuardar();
    const bytes = await pdf.save();
    // Si no se ganó nada (PDF ya optimizado, o sin fotos para recomprimir),
    // se devuelve el original: guardar con pdf-lib a veces pesa un poco más.
    const sinMejora = bytes.length >= pesoOriginal;
    return { bytes: sinMejora ? new Uint8Array(bytesOriginales) : bytes, procesadas, sinMejora };
  }

  // ---------- Extraer páginas (extraer-paginas.html) ----------
  // "1-3, 5, 9-7" → índices desde 0, en ese orden (un rango al revés va al revés). Lanza un Error con
  // `codigo` 'vacio' | 'formato' | 'rango' (y `trozo`, `max`) para que quien llama lo diga en su idioma.
  function indicesDeRango(texto, maxPagina) {
    const falla = (codigo, trozo) => Object.assign(new Error(codigo), { codigo, trozo, max: maxPagina });
    const indices = [];
    const trozos = String(texto).split(',').map((x) => x.trim()).filter((x) => x !== '');
    if (trozos.length === 0) throw falla('vacio');
    for (const trozo of trozos) {
      const coincidencia = trozo.match(/^(\d+)(?:-(\d+))?$/);
      if (!coincidencia) throw falla('formato', trozo);
      const inicio = parseInt(coincidencia[1], 10);
      const fin = coincidencia[2] ? parseInt(coincidencia[2], 10) : inicio;
      if (inicio < 1 || fin < 1 || inicio > maxPagina || fin > maxPagina) throw falla('rango', trozo);
      if (inicio <= fin) {
        for (let p = inicio; p <= fin; p++) indices.push(p - 1);
      } else {
        for (let p = inicio; p >= fin; p--) indices.push(p - 1);
      }
    }
    return indices;
  }

  // Un PDF nuevo con esas páginas (índices desde 0) de pdfOrigen.
  async function extraerPaginas(pdfOrigen, indices) {
    const pdfNuevo = await PDFLib.PDFDocument.create();
    const paginas = await pdfNuevo.copyPages(pdfOrigen, indices);
    paginas.forEach((pagina) => pdfNuevo.addPage(pagina));
    return pdfNuevo;
  }

  // ---------- Dividir (dividir-pdf.html) ----------
  // Los bytes de un PDF por página, en orden.
  async function dividir(pdfOrigen) {
    const partes = [];
    for (let i = 0; i < pdfOrigen.getPageCount(); i++) partes.push(await (await extraerPaginas(pdfOrigen, [i])).save());
    return partes;
  }

  // ---------- Metadatos (editar-metadatos-pdf.html) ----------
  // Solo cambia los campos que vienen (un campo ausente queda como estaba). `pdf` se abre con
  // { updateMetadata: false }: si no, pdf-lib reemplaza Creador y Productor por "pdf-lib".
  function ponerMetadatos(pdf, { titulo, autor, asunto, palabras } = {}) {
    if (titulo !== undefined) pdf.setTitle(titulo);
    if (autor !== undefined) pdf.setAuthor(autor);
    if (asunto !== undefined) pdf.setSubject(asunto);
    // Un solo elemento: setKeywords une la lista con espacios y se perderían las comas.
    if (palabras !== undefined) pdf.setKeywords(palabras.trim() ? [palabras.trim()] : []);
    pdf.setModificationDate(new Date());
  }

  // calcularFilasColumnas: también la usa la vista previa de marca-de-agua.html (misma cuadrícula que el PDF).
  window.PdfRetoques = {
    rotar, numerar, marcaDeAgua, optimizar, calcularFilasColumnas, indicesDeRango, extraerPaginas, dividir, ponerMetadatos,
  };
})();

/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Los módulos de Flujos (assets/js/flujos.js). Cada uno es un adaptador fino sobre la pieza compartida que
// usa también la herramienta de su página: el flujo y la página hacen lo mismo con el mismo código.
//
//   leer-codigos        LeerCodigos (lector-codigos.html)            imagen|pdf -> igual   (guarda datos.codigos)
//   recortar-documento  EscanerDocumento + Custos (escanear-documento.html)   imagen -> imagen
//   girar-imagen        GirarImagen (rotar-imagen.html)              imagen -> imagen
//   comprimir-imagen    ComprimirImagen (comprimir-imagen.html)      imagen -> imagen
//   imagen-a-pdf        miniaturas.js + pdf-lib (pdf-imagenes.html)  imagen -> pdf (un PDF por imagen; códigos en una tabla debajo)
//   pdf-a-imagenes      PdfAImagenes (pdf-imagenes.html)             pdf -> imágenes (abre: una por página)
//   extraer-imagenes-pdf PdfExtraerImagenes (extraer-imagenes.html)  pdf -> imágenes (abre: las incrustadas)
//   unir-pdf            PdfUnir (unir-pdf.html)                      todos los pdf -> un pdf
// Etapa 2 (2026-11-10):
//   documento-a-pdf     ArchivoAPdf (unir-pdf.html)                  documento|texto -> pdf
//   ocr-pdf, ocr-texto  Ocr (ocr-pdf.html)                           imagen|pdf -> pdf buscable | texto
//   pdf-a-texto         textoDePdf de pdf-texto.js (pdf-texto.html)  pdf -> texto
//   redimensionar-imagen, convertir-imagen, quitar-metadatos   ImagenBasica (sus 3 páginas)   imagen -> imagen
//   rotar-pdf, comprimir-pdf, numerar-pdf, marca-agua-pdf      PdfRetoques (sus 4 páginas)    pdf -> pdf
//
// Un módulo que junta (forma 'juntar', como unir-pdf) es el que se usa para unir varias entradas, pero el tramo
// común acepta cualquier módulo que reciba lo que llega de todas: unir PDF es solo el primero de ese tipo.
//
// Cada módulo carga lo que necesita la primera vez que se usa (preparar). Necesita assets/js/vendor.js antes.
// Agregar una herramienta: separar su lógica en una pieza compartida de assets/js, usarla en su página y
// escribir aquí su ficha (ver el contrato en flujos.js).
(function () {
  const F = window.Flujos;
  const js = (archivo) => F.cargarScript(F.rutaJs(archivo));

  // ---------- utilidades de imagen ----------
  const LADO_MAX_FUENTE = 4096;   // como el escáner: más grande no mejora el resultado y gasta memoria

  async function lienzoDe(blob, ladoMax) {
    const mapa = await createImageBitmap(blob);   // aplica la orientación EXIF de las fotos de celular
    const escala = ladoMax ? Math.min(1, ladoMax / Math.max(mapa.width, mapa.height)) : 1;
    const lienzo = document.createElement('canvas');
    lienzo.width = Math.max(1, Math.round(mapa.width * escala));
    lienzo.height = Math.max(1, Math.round(mapa.height * escala));
    lienzo.getContext('2d').drawImage(mapa, 0, 0, lienzo.width, lienzo.height);
    mapa.close();
    return lienzo;
  }

  function aBlob(lienzo, tipo, calidad) {
    return new Promise((ok, mal) => lienzo.toBlob((b) => (b ? ok(b) : mal(new Error('toBlob'))), tipo, calidad));
  }

  const EXT = { 'image/jpeg': '.jpg', 'image/png': '.png', 'image/webp': '.webp' };
  const conExtension = (nombre, tipo) => F.base(nombre) + (EXT[tipo] || '.png');

  // ---------- 1. Leer códigos QR y de barras ----------
  F.registrar({
    id: 'leer-codigos',
    icono: '🔍',
    nombre: { es: 'Leer códigos QR y de barras', en: 'Read QR and barcodes' },
    descripcion: { es: 'Busca un código en cada archivo y guarda lo que dice para los pasos siguientes.', en: 'Looks for a code in each file and keeps what it says for the next steps.' },
    entra: ['imagen', 'pdf'], sale: 'igual', forma: 'uno',
    opciones: [
      { id: 'sinCodigo', tipo: 'select', defecto: 'seguir', etiqueta: { es: 'Si no hay código', en: 'If there is no code' },
        valores: [{ valor: 'seguir', es: 'Seguir con el archivo', en: 'Keep the file' }, { valor: 'quitar', es: 'Quitarlo del flujo', en: 'Drop it from the flow' }] },
    ],
    async preparar() {
      await F.cargarScript(rutaVendor('@zxing/library/0.23.0/umd/index.min.js'));
      await js('pdfjs.js');
      await js('leer-codigos.js');
    },
    async ejecutar(f, o) {
      const codigos = f.tipo === 'pdf'
        ? await LeerCodigos.dePdf(f.blob)
        : await LeerCodigos.deImagen(f.blob);
      if (!codigos.length && o.sinCodigo === 'quitar') return null;
      const vistos = new Set();   // el mismo código en dos páginas de un PDF se anota una vez
      const nuevos = codigos.filter((c) => !vistos.has(c.formato + '\n' + c.texto) && vistos.add(c.formato + '\n' + c.texto));
      return { ...f, datos: { ...f.datos, codigos: [...(f.datos.codigos || []), ...nuevos.map((c) => ({ texto: c.texto, formato: c.formato }))] } };
    },
  });

  // ---------- 2. Recortar documento (escáner con IA) ----------
  const areaCuadrilatero = (q) => Math.abs(q.reduce((s, p, i) => { const n = q[(i + 1) % 4]; return s + p.x * n.y - n.x * p.y; }, 0)) / 2;
  const MEJORAS = [
    { valor: 'color', es: 'Color realzado', en: 'Enhanced color' },
    { valor: 'original', es: 'Original', en: 'Original' },
    { valor: 'gris', es: 'Escala de grises', en: 'Grayscale' },
    { valor: 'bn', es: 'Blanco y negro', en: 'Black and white' },
  ];
  F.registrar({
    id: 'recortar-documento',
    icono: '📄',
    nombre: { es: 'Recortar documento (escáner con IA)', en: 'Crop document (AI scanner)' },
    descripcion: { es: 'Encuentra la hoja en la foto, la endereza y recorta el fondo.', en: 'Finds the sheet in the photo, straightens it and crops the background.' },
    entra: ['imagen'], sale: 'imagen', forma: 'uno',
    opciones: [
      { id: 'mejora', tipo: 'select', defecto: 'color', etiqueta: { es: 'Mejora', en: 'Enhancement' }, valores: MEJORAS },
      { id: 'sinHoja', tipo: 'select', defecto: 'dejar', etiqueta: { es: 'Si no encuentra la hoja', en: 'If no sheet is found' },
        valores: [{ valor: 'dejar', es: 'Dejar la foto como está', en: 'Keep the photo as is' }, { valor: 'quitar', es: 'Quitarla del flujo', en: 'Drop it from the flow' }] },
    ],
    async preparar() {
      await js('escaner-documento.js');
      await js('custos.js');
      Custos.cargar();
    },
    async ejecutar(f, o) {
      const lienzo = await lienzoDe(f.blob, LADO_MAX_FUENTE);
      const datos = lienzo.getContext('2d', { willReadFrequently: true }).getImageData(0, 0, lienzo.width, lienzo.height);
      // Igual que el escáner: primero el motor clásico y, si Custos está listo, se combinan los dos.
      let hallazgo = EscanerDocumento.detectarEsquinas(datos);
      if (await Custos.esperar(15000)) {
        try { hallazgo = Custos.combinar(await Custos.detectar(lienzo, lienzo.width, lienzo.height), hallazgo); } catch (e) { /* queda el clásico */ }
      }
      if (!hallazgo.encontrado) return o.sinHoja === 'quitar' ? null : f;
      // Un «documento» que ocupa casi toda la foto suele ser el borde de la foto misma (madera, una pared):
      // recortar un par de por ciento no aporta y no es motivo para quitarla. Pasa tal cual.
      if (areaCuadrilatero(hallazgo.esquinas) > 0.92 * lienzo.width * lienzo.height) return f;
      const { ancho, alto } = EscanerDocumento.tamanoSalida(hallazgo.esquinas, 2400);
      const recta = EscanerDocumento.enderezar(datos, hallazgo.esquinas, ancho, alto);
      EscanerDocumento.mejorar(recta, o.mejora, {});
      const salida = document.createElement('canvas');
      salida.width = recta.width; salida.height = recta.height;
      salida.getContext('2d').putImageData(new ImageData(recta.data, recta.width, recta.height), 0, 0);
      return { ...f, blob: await aBlob(salida, 'image/jpeg', 0.88), nombre: conExtension(f.nombre, 'image/jpeg'), tipo: 'imagen' };
    },
  });

  // ---------- 3. Girar imagen ----------
  F.registrar({
    id: 'girar-imagen',
    icono: '🔄',
    nombre: { es: 'Girar o voltear imagen', en: 'Rotate or flip image' },
    descripcion: { es: 'Gira en pasos de 90° o voltea, sin perder calidad.', en: 'Rotates in 90° steps or flips, without losing quality.' },
    entra: ['imagen'], sale: 'imagen', forma: 'uno',
    opciones: [
      { id: 'giro', tipo: 'select', defecto: '90', etiqueta: { es: 'Girar', en: 'Rotate' },
        valores: [{ valor: '0', es: 'No girar', en: 'Do not rotate' }, { valor: '90', es: '90° a la derecha', en: '90° right' },
          { valor: '180', es: '180°', en: '180°' }, { valor: '270', es: '90° a la izquierda', en: '90° left' }] },
      { id: 'voltear', tipo: 'select', defecto: 'no', etiqueta: { es: 'Voltear', en: 'Flip' },
        valores: [{ valor: 'no', es: 'No voltear', en: 'Do not flip' }, { valor: 'h', es: 'Horizontal', en: 'Horizontal' }, { valor: 'v', es: 'Vertical', en: 'Vertical' }] },
    ],
    async preparar() { await js('girar-imagen.js'); },
    async ejecutar(f, o) {
      let lienzo = await lienzoDe(f.blob);
      if (o.giro !== '0') lienzo = GirarImagen.rotar(lienzo, Number(o.giro));
      if (o.voltear !== 'no') lienzo = GirarImagen.voltear(lienzo, o.voltear === 'h');
      const tipo = GirarImagen.formatoDeSalida(f.blob.type || 'image/png', false);
      return { ...f, blob: await aBlob(lienzo, tipo, 0.92), nombre: conExtension(f.nombre, tipo) };
    },
  });

  // ---------- 4. Comprimir imagen ----------
  F.registrar({
    id: 'comprimir-imagen',
    icono: '🗜️',
    nombre: { es: 'Comprimir imagen', en: 'Compress image' },
    descripcion: { es: 'Vuelve a guardar la imagen como JPG o WebP con la calidad elegida.', en: 'Saves the image again as JPG or WebP at the chosen quality.' },
    entra: ['imagen'], sale: 'imagen', forma: 'uno',
    opciones: [
      { id: 'calidad', tipo: 'numero', defecto: 70, min: 10, max: 95, paso: 5, etiqueta: { es: 'Calidad (%)', en: 'Quality (%)' } },
      { id: 'formato', tipo: 'select', defecto: 'image/jpeg', etiqueta: { es: 'Formato', en: 'Format' },
        valores: [{ valor: 'image/jpeg', es: 'JPG', en: 'JPG' }, { valor: 'image/webp', es: 'WebP', en: 'WebP' }] },
    ],
    async preparar() { await js('comprimir-imagen.js'); },
    async ejecutar(f, o) {
      const blob = await ComprimirImagen.comprimir(f.blob, { calidad: o.calidad / 100, formato: o.formato });
      return { ...f, blob, nombre: conExtension(f.nombre, blob.type) };
    },
  });

  // ---------- 5. Imagen a PDF (un PDF por imagen) ----------
  // Cuadro de los códigos leídos: una tabla JUSTO DEBAJO de la imagen, con aspecto de comentario (fondo amarillo
  // claro y borde), a todo el ancho. Una fila por código: tipo | contenido. Pedidos del usuario (2026-11-10): el
  // texto suelto al pie se leía mal, y el cuadro no debe tapar la imagen. En las páginas que miden lo que la imagen
  // («ajustar-*», «original») la página crece hacia abajo; en A4/Carta exactos el contenido se achica para dejarle
  // lugar y la hoja conserva su tamaño estándar. Con Helvetica si el texto entra en su juego de caracteres
  // (queda seleccionable); si trae algo que Helvetica no tiene (emoji, chino…), el texto se dibuja como imagen
  // sobre el mismo cuadro.
  const MAX_LINEAS = 10;
  const TITULO_CUADRO = { es: 'Códigos leídos', en: 'Codes read' };

  function partir(texto, ancho, medir) {
    const lineas = [];
    for (const parrafo of String(texto).split(/\r?\n/)) {
      let actual = '';
      for (const ch of parrafo) {   // por caracter: una URL larga no tiene espacios donde cortar
        if (actual && medir(actual + ch) > ancho) { lineas.push(actual); actual = ''; }
        actual += ch;
      }
      lineas.push(actual);
    }
    return lineas;
  }

  async function ponerCuadro(pdf, pagina, filas, tamano) {
    const { width: W, height: H } = pagina.getSize();
    const tam = Math.max(9, Math.min(16, W * 0.02));
    const pad = tam * 0.6, inter = tam * 1.3, margen = tam, anchoMax = W - 2 * margen;
    const titulo = TITULO_CUADRO[(window.idiomaActual && idiomaActual()) === 'en' ? 'en' : 'es'];
    const [normal, negrita] = await Promise.all([pdf.embedFont(PDFLib.StandardFonts.Helvetica), pdf.embedFont(PDFLib.StandardFonts.HelveticaBold)]);
    let comoTexto = true;
    try { [titulo, ...filas.flat()].forEach((x) => normal.encodeText(x)); } catch (e) { comoTexto = false; }
    const ctx = document.createElement('canvas').getContext('2d');
    const medirCon = (negr) => (comoTexto
      ? (x) => (negr ? negrita : normal).widthOfTextAtSize(x, tam)
      : (x) => { ctx.font = `${negr ? 'bold ' : ''}${tam}px sans-serif`; return ctx.measureText(x).width; });
    const medir = medirCon(false), medirN = medirCon(true);

    // Medidas: columna del tipo al ancho del más largo; la del contenido, lo que quede (partida en líneas).
    const col1 = Math.min(anchoMax * 0.35, Math.max(...filas.map((f) => medirN(f[0]))));
    const col2Max = anchoMax - col1 - 3 * pad;
    let quedan = MAX_LINEAS;
    const celdas = [];
    for (const [tipo, contenido] of filas) {
      if (quedan <= 0) break;
      let lineas = partir(contenido, col2Max, medir);
      if (lineas.length > quedan) { lineas = lineas.slice(0, quedan); lineas[quedan - 1] = lineas[quedan - 1].slice(0, -1) + '…'; }
      quedan -= lineas.length;
      celdas.push({ tipo: partir(tipo, col1, medirN)[0], lineas });
    }
    const anchoCaja = anchoMax;   // a todo el ancho: se lee como una tabla debajo de la imagen
    const altoTitulo = inter + pad;
    const altoFilas = celdas.map((c) => c.lineas.length * inter + pad);
    const altoCaja = altoTitulo + altoFilas.reduce((a, b) => a + b, 0) + pad * 0.5;
    // Lugar para la tabla debajo de la imagen (en PDF, y crece hacia arriba).
    const espacio = altoCaja + 2 * margen;
    let y0;
    if (tamano === 'a4' || tamano === 'carta') {
      // La página con la imagen entra achicada en una página nueva del mismo tamaño. No con scaleContent: lo que
      // se dibuja después cae en el mismo flujo de contenido, dentro de la transformación, y la tabla también se
      // achicaba y se corría encima de la imagen (visto en la prueba del 2026-11-10).
      const s = (H - espacio) / H;
      const incrustada = await pdf.embedPage(pagina);
      const indice = pdf.getPages().indexOf(pagina);
      const nueva = pdf.insertPage(indice + 1, [W, H]);
      nueva.drawPage(incrustada, { x: (W - W * s) / 2, y: espacio, xScale: s, yScale: s });
      pdf.removePage(indice);
      pagina = nueva;
      y0 = margen;
    } else {
      pagina.setMediaBox(0, -espacio, W, H + espacio);
      y0 = -espacio + margen;
    }
    const x0 = margen;

    pagina.drawRectangle({ x: x0, y: y0, width: anchoCaja, height: altoCaja, color: PDFLib.rgb(1, 0.98, 0.8), opacity: 0.86,
      borderColor: PDFLib.rgb(0.78, 0.66, 0.2), borderWidth: Math.max(0.6, tam * 0.06), borderOpacity: 0.95 });

    // Posiciones (desde arriba de la caja) de cada texto, para dibujarlos igual como texto o como imagen.
    const textos = [{ s: titulo, x: pad, yArriba: pad, negr: true }];
    const lineasSep = [];
    let y = altoTitulo;
    celdas.forEach((c, i) => {
      lineasSep.push(y);
      textos.push({ s: c.tipo, x: pad, yArriba: y + pad * 0.5, negr: true });
      c.lineas.forEach((l, k) => textos.push({ s: l, x: pad + col1 + pad, yArriba: y + pad * 0.5 + k * inter, negr: false }));
      y += altoFilas[i];
    });
    const arriba = y0 + altoCaja;
    const colorSep = PDFLib.rgb(0.78, 0.66, 0.2);
    lineasSep.forEach((yy) => pagina.drawLine({ start: { x: x0 + pad * 0.5, y: arriba - yy }, end: { x: x0 + anchoCaja - pad * 0.5, y: arriba - yy }, thickness: 0.5, color: colorSep, opacity: 0.6 }));
    const colorTexto = PDFLib.rgb(0.13, 0.13, 0.13);
    if (comoTexto) {
      textos.forEach((t) => pagina.drawText(t.s, { x: x0 + t.x, y: arriba - t.yArriba - tam * 0.95, size: tam, font: t.negr ? negrita : normal, color: colorTexto }));
      return;
    }
    const k = 3;   // el texto como imagen se dibuja a 3x para que se lea bien impreso; fondo transparente
    const c = document.createElement('canvas');
    c.width = Math.ceil(anchoCaja * k); c.height = Math.ceil(altoCaja * k);
    const g = c.getContext('2d');
    g.fillStyle = '#222'; g.textBaseline = 'top';
    textos.forEach((t) => { g.font = `${t.negr ? 'bold ' : ''}${tam * k}px sans-serif`; g.fillText(t.s, t.x * k, (t.yArriba + tam * 0.1) * k); });
    const png = await pdf.embedPng(new Uint8Array(await (await aBlob(c, 'image/png')).arrayBuffer()));
    pagina.drawImage(png, { x: x0, y: y0, width: anchoCaja, height: altoCaja });
  }

  F.registrar({
    id: 'imagen-a-pdf',
    icono: '🖼️',
    nombre: { es: 'Imagen a PDF', en: 'Image to PDF' },
    descripcion: { es: 'Un PDF por cada imagen. Si un paso anterior leyó un código, lo muestra en un cuadro debajo de la imagen.', en: 'One PDF per image. If an earlier step read a code, it shows it in a box below the image.' },
    entra: ['imagen'], sale: 'pdf', forma: 'uno',
    opciones: [
      { id: 'tamano', tipo: 'select', defecto: 'ajustar-a4', etiqueta: { es: 'Tamaño de página', en: 'Page size' },
        valores: [
          { valor: 'ajustar-a4', es: 'Ancho A4, alto según la imagen', en: 'A4 width, height to fit the image' },
          { valor: 'ajustar-carta', es: 'Ancho Carta, alto según la imagen', en: 'Letter width, height to fit the image' },
          { valor: 'a4', es: 'A4 (con márgenes)', en: 'A4 (with margins)' },
          { valor: 'carta', es: 'Carta (con márgenes)', en: 'Letter (with margins)' },
          { valor: 'original', es: 'Tamaño de la imagen', en: 'Image size' }] },
      { id: 'pie', tipo: 'casilla', defecto: true, etiqueta: { es: 'Mostrar los códigos leídos en un cuadro', en: 'Show the codes read in a box' } },
    ],
    async preparar() {
      await F.cargarScript(rutaVendor('pdf-lib/1.17.1/pdf-lib.min.js'));
      await js('miniaturas.js');
    },
    async ejecutar(f, o) {
      const pdf = await PDFLib.PDFDocument.create();
      agregarPaginaConImagen(pdf, await incrustarImagenEnPdf(pdf, f.blob), o.tamano);
      const codigos = f.datos.codigos || [];
      if (o.pie && codigos.length) {
        const nombre = (c) => (window.LeerCodigos ? LeerCodigos.nombreFormato(c.formato) : c.formato);
        await ponerCuadro(pdf, pdf.getPage(0), codigos.map((c) => [nombre(c), c.texto]), o.tamano);
      }
      if (codigos.length) pdf.setKeywords(codigos.map((c) => c.texto));
      const bytes = await pdf.save();
      return { ...f, blob: new Blob([bytes], { type: 'application/pdf' }), nombre: F.base(f.nombre) + '.pdf', tipo: 'pdf' };
    },
  });

  // ---------- 6. PDF a imágenes (una por página) ----------
  const fichaDeImagen = (f, img, nombre) => ({ ...f, datos: { ...f.datos }, blob: img.blob, tipo: 'imagen', nombre });
  F.registrar({
    id: 'pdf-a-imagenes',
    icono: '🖼️',
    nombre: { es: 'PDF a imágenes (cada página)', en: 'PDF to images (each page)' },
    descripcion: { es: 'Convierte cada página del PDF en una imagen; cada una sigue el flujo por su cuenta.', en: 'Turns each page of the PDF into an image; each one goes on through the flow on its own.' },
    entra: ['pdf'], sale: 'imagen', forma: 'abrir',
    opciones: [
      { id: 'escala', tipo: 'select', defecto: '2', etiqueta: { es: 'Resolución', en: 'Resolution' },
        valores: [{ valor: '1.5', es: 'Normal', en: 'Normal' }, { valor: '2', es: 'Alta', en: 'High' }, { valor: '3', es: 'Muy alta', en: 'Very high' }] },
      { id: 'formato', tipo: 'select', defecto: 'image/png', etiqueta: { es: 'Formato', en: 'Format' },
        valores: [{ valor: 'image/png', es: 'PNG', en: 'PNG' }, { valor: 'image/jpeg', es: 'JPG', en: 'JPG' }] },
    ],
    async preparar() { await js('pdfjs.js'); await js('pdf-a-imagenes.js'); },
    async ejecutar(f, o) {
      const paginas = await PdfAImagenes.renderizar(f.blob, { escala: Number(o.escala), formato: o.formato, calidad: 0.92 });
      const ext = o.formato === 'image/jpeg' ? '.jpg' : '.png', ancho = String(paginas.length).length;
      return paginas.map((p) => fichaDeImagen(f, p, `${F.base(f.nombre)}-pagina-${String(p.pagina).padStart(ancho, '0')}${ext}`));
    },
  });

  // ---------- 7. Extraer las imágenes incrustadas de un PDF ----------
  F.registrar({
    id: 'extraer-imagenes-pdf',
    icono: '🧩',
    nombre: { es: 'Extraer imágenes de PDF (incrustadas)', en: 'Extract images from PDF (embedded)' },
    descripcion: { es: 'Saca las imágenes que el PDF lleva adentro, sin el texto; cada una sigue el flujo por su cuenta.', en: 'Takes out the images the PDF carries inside, without the text; each one goes on through the flow on its own.' },
    entra: ['pdf'], sale: 'imagen', forma: 'abrir',
    opciones: [
      { id: 'sinImagenes', tipo: 'select', defecto: 'avisar', etiqueta: { es: 'Si el PDF no trae imágenes', en: 'If the PDF has no images' },
        valores: [{ valor: 'avisar', es: 'Avisar en el resumen', en: 'Report it in the summary' }, { valor: 'pagina', es: 'Usar cada página como imagen', en: 'Use each page as an image' }] },
    ],
    async preparar() { await js('pdfjs.js'); await js('pdf-extraer-imagenes.js'); await js('pdf-a-imagenes.js'); },
    async ejecutar(f, o) {
      const imagenes = await PdfExtraerImagenes.extraer(f.blob, {});
      if (imagenes.length) return imagenes.map((i) => fichaDeImagen(f, i, `${F.base(f.nombre)}-${i.nombre}`));
      if (o.sinImagenes === 'pagina') {
        const paginas = await PdfAImagenes.renderizar(f.blob, { escala: 2, formato: 'image/png', calidad: 0.92 });
        return paginas.map((p) => fichaDeImagen(f, p, `${F.base(f.nombre)}-pagina-${p.pagina}.png`));
      }
      return [];   // queda en «lo que no pasó» como quitado en este paso
    },
  });

  // ---------- 8. Unir PDF ----------
  F.registrar({
    id: 'unir-pdf',
    icono: '📎',
    nombre: { es: 'Unir en un solo PDF', en: 'Merge into one PDF' },
    descripcion: { es: 'Junta todos los PDF que llegan a este paso, en el orden de los archivos.', en: 'Joins every PDF that reaches this step, in the order of the files.' },
    entra: ['pdf'], sale: 'pdf', forma: 'juntar',
    opciones: [],
    async preparar() {
      await F.cargarScript(rutaVendor('pdf-lib/1.17.1/pdf-lib.min.js'));
      await js('pdf-unir.js');
    },
    async ejecutar(fichas) {
      const bytes = await PdfUnir.unir(fichas.map((f) => f.blob));
      const codigos = fichas.flatMap((f) => f.datos.codigos || []);
      return { nombre: 'unido.pdf', blob: new Blob([bytes], { type: 'application/pdf' }), tipo: 'pdf', datos: codigos.length ? { codigos } : {}, origen: 'unido.pdf' };
    },
  });

  // ================= Etapa 2 =================
  const PDF_LIB = () => F.cargarScript(rutaVendor('pdf-lib/1.17.1/pdf-lib.min.js'));
  // Las piezas que miran archivo.name o archivo.type reciben un File con el nombre y el tipo de la ficha.
  const comoArchivo = (f, tipo) => new File([f.blob], f.nombre, { type: tipo || f.blob.type || '' });
  const fichaPdf = (f, bytes) => ({ ...f, blob: new Blob([bytes], { type: 'application/pdf' }), nombre: F.base(f.nombre) + '.pdf', tipo: 'pdf' });
  const fichaTexto = (f, texto) => ({ ...f, blob: new Blob([texto], { type: 'text/plain;charset=utf-8' }), nombre: F.base(f.nombre) + '.txt', tipo: 'texto' });
  const SEPARADOR_PAGINA = { es: '\n\n--- Página siguiente ---\n\n', en: '\n\n--- Next page ---\n\n' };   // como en ocr-pdf y pdf-texto
  const separador = () => SEPARADOR_PAGINA[(window.idiomaActual && idiomaActual()) === 'en' ? 'en' : 'es'];

  // ---------- 9. Documento a PDF ----------
  F.registrar({
    id: 'documento-a-pdf',
    icono: '📑',
    nombre: { es: 'Documento a PDF', en: 'Document to PDF' },
    descripcion: { es: 'Pasa Word, Excel, CSV, PowerPoint, EPUB, Markdown o texto a PDF, con los mismos motores que Unir PDF.', en: 'Turns Word, Excel, CSV, PowerPoint, EPUB, Markdown or text into PDF, with the same engines as Merge PDF.' },
    entra: ['documento', 'texto'], sale: 'pdf', forma: 'uno',
    opciones: [],
    async preparar() { await PDF_LIB(); },   // archivo-a-pdf.js ya está en la página; cada formato trae su motor al usarse
    async ejecutar(f) {
      return fichaPdf(f, await ArchivoAPdf.convertir(comoArchivo(f)));
    },
  });

  // ---------- 10 y 11. OCR (PDF buscable o texto) ----------
  const IDIOMAS_OCR = [
    { valor: 'spa+eng', es: 'Español e inglés', en: 'Spanish and English' },
    { valor: 'spa', es: 'Español', en: 'Spanish' },
    { valor: 'eng', es: 'Inglés', en: 'English' },
    { valor: 'por', es: 'Portugués', en: 'Portuguese' },
    { valor: 'fra', es: 'Francés', en: 'French' },
    { valor: 'deu', es: 'Alemán', en: 'German' },
    { valor: 'ita', es: 'Italiano', en: 'Italian' },
    { valor: 'chi_sim', es: 'Chino simplificado', en: 'Simplified Chinese' },
    { valor: 'jpn', es: 'Japonés', en: 'Japanese' },
  ];
  const OPCIONES_OCR = [
    { id: 'idiomas', tipo: 'select', defecto: 'spa+eng', etiqueta: { es: 'Idioma del texto', en: 'Text language' }, valores: IDIOMAS_OCR },
    { id: 'enderezar', tipo: 'casilla', defecto: true, etiqueta: { es: 'Enderezar páginas torcidas', en: 'Straighten tilted pages' } },
    { id: 'preproceso', tipo: 'select', defecto: 'contraste', etiqueta: { es: 'Preparar la imagen', en: 'Prepare the image' },
      valores: [{ valor: 'contraste', es: 'Contraste automático', en: 'Automatic contrast' }, { valor: 'otsu', es: 'Blanco y negro puro', en: 'Pure black and white' },
        { valor: 'ninguno', es: 'Tal cual', en: 'As is' }] },
  ];
  // Un motor de Tesseract por combinación de idiomas, compartido por los dos módulos y por todos los archivos de una
  // ejecución (crearlo tarda); terminar() los cierra al final.
  const motoresOcr = new Map();
  const motorOcr = (idiomas) => {
    if (!motoresOcr.has(idiomas)) motoresOcr.set(idiomas, Ocr.crearWorker(idiomas.split('+')));
    return motoresOcr.get(idiomas);
  };
  async function cerrarMotoresOcr() {
    const lista = [...motoresOcr.values()];
    motoresOcr.clear();
    for (const p of lista) { try { await (await p).terminate(); } catch (e) { /* ya cerrado */ } }
  }
  async function reconocer(f, o) {
    const worker = await motorOcr(o.idiomas);
    const pdf = await PDFLib.PDFDocument.create();
    const textos = [];
    for (const pagina of await Ocr.paginasDeArchivo(comoArchivo(f, f.tipo === 'pdf' ? 'application/pdf' : undefined), o.enderezar)) {
      textos.push(await Ocr.agregarPagina(pdf, worker, pagina, o.preproceso));
    }
    return { pdf, texto: textos.join(separador()) };
  }
  const prepararOcr = async () => {
    await PDF_LIB();
    await js('pdfjs.js');
    await F.cargarScript(rutaVendor('tesseract.js/7.0.0/dist/tesseract.min.js?v=2'));
    await js('ocr.js');
  };
  F.registrar({
    id: 'ocr-pdf',
    icono: '🔤',
    nombre: { es: 'OCR: PDF con texto buscable', en: 'OCR: searchable PDF' },
    descripcion: { es: 'Reconoce el texto de fotos y escaneos y entrega un PDF igual al original donde se puede buscar y copiar.', en: 'Recognizes the text in photos and scans and gives back a PDF that looks the same but can be searched and copied.' },
    entra: ['imagen', 'pdf'], sale: 'pdf', forma: 'uno',
    opciones: OPCIONES_OCR,
    preparar: prepararOcr,
    terminar: cerrarMotoresOcr,
    async ejecutar(f, o) {
      const { pdf } = await reconocer(f, o);
      return fichaPdf(f, await pdf.save());
    },
  });
  F.registrar({
    id: 'ocr-texto',
    icono: '🔤',
    nombre: { es: 'OCR: sacar el texto (.txt)', en: 'OCR: get the text (.txt)' },
    descripcion: { es: 'Reconoce el texto de fotos y escaneos y lo guarda en un archivo de texto.', en: 'Recognizes the text in photos and scans and saves it as a text file.' },
    entra: ['imagen', 'pdf'], sale: 'texto', forma: 'uno',
    opciones: OPCIONES_OCR,
    preparar: prepararOcr,
    terminar: cerrarMotoresOcr,
    async ejecutar(f, o) {
      const { texto } = await reconocer(f, o);
      return fichaTexto(f, texto);
    },
  });

  // ---------- 12. PDF a texto ----------
  F.registrar({
    id: 'pdf-a-texto',
    icono: '📝',
    nombre: { es: 'PDF a texto (.txt)', en: 'PDF to text (.txt)' },
    descripcion: { es: 'Saca el texto que el PDF ya trae, con sus párrafos y tablas. Para escaneos, usá el OCR.', en: 'Takes out the text the PDF already has, with its paragraphs and tables. For scans, use OCR.' },
    entra: ['pdf'], sale: 'texto', forma: 'uno',
    opciones: [
      { id: 'sinTexto', tipo: 'select', defecto: 'dejar', etiqueta: { es: 'Si el PDF no tiene texto (escaneo)', en: 'If the PDF has no text (scan)' },
        valores: [{ valor: 'dejar', es: 'Dejar el .txt vacío', en: 'Keep the empty .txt' }, { valor: 'quitar', es: 'Quitarlo del flujo', en: 'Drop it from the flow' }] },
    ],
    async preparar() { await PDF_LIB(); await js('pdfjs.js'); await js('tablas-texto.js'); await js('pdf-texto.js'); },
    async ejecutar(f, o) {
      const { texto, sinTexto } = await textoDePdf(f.blob, separador());
      if (sinTexto && o.sinTexto === 'quitar') return null;
      return fichaTexto(f, texto);
    },
  });

  // ---------- 13. Redimensionar imagen ----------
  F.registrar({
    id: 'redimensionar-imagen',
    icono: '📐',
    nombre: { es: 'Redimensionar imagen', en: 'Resize image' },
    descripcion: { es: 'Cambia el tamaño manteniendo la proporción: por el lado más largo, por el ancho o en porcentaje.', en: 'Changes the size keeping the proportions: by the longest side, by the width or by a percentage.' },
    entra: ['imagen'], sale: 'imagen', forma: 'uno',
    opciones: [
      { id: 'medida', tipo: 'select', defecto: 'lado', etiqueta: { es: 'Medida', en: 'Measure' },
        valores: [{ valor: 'lado', es: 'Lado más largo (px)', en: 'Longest side (px)' }, { valor: 'ancho', es: 'Ancho (px)', en: 'Width (px)' },
          { valor: 'porcentaje', es: 'Porcentaje (%)', en: 'Percentage (%)' }] },
      { id: 'valor', tipo: 'numero', defecto: 1600, min: 1, max: 8000, etiqueta: { es: 'Valor', en: 'Value' } },
      { id: 'agrandar', tipo: 'casilla', defecto: false, etiqueta: { es: 'Agrandar también las imágenes más chicas', en: 'Also enlarge smaller images' } },
      { id: 'suavizado', tipo: 'select', defecto: 'suave', etiqueta: { es: 'Calidad', en: 'Quality' },
        valores: [{ valor: 'suave', es: 'Suave (recomendado)', en: 'Smooth (recommended)' }, { valor: 'nitido', es: 'Nítido (pixel art)', en: 'Sharp (pixel art)' }] },
    ],
    async preparar() { await js('imagen-basica.js'); },
    async ejecutar(f, o) {
      const mapa = await createImageBitmap(f.blob);   // aplica la orientación EXIF de las fotos de celular
      try {
        const w = mapa.width, h = mapa.height;
        let escala = o.medida === 'porcentaje' ? o.valor / 100 : o.medida === 'ancho' ? o.valor / w : o.valor / Math.max(w, h);
        if (!o.agrandar) escala = Math.min(1, escala);
        const ancho = Math.max(1, Math.round(w * escala)), alto = Math.max(1, Math.round(h * escala));
        if (ancho === w && alto === h) return f;
        if (ancho > ImagenBasica.LIMITE_DIMENSION_PX || alto > ImagenBasica.LIMITE_DIMENSION_PX || ancho * alto > ImagenBasica.LIMITE_MEGAPIXELES * 1e6) {
          throw new Error(`${ancho}×${alto} px`);
        }
        const lienzo = ImagenBasica.escalar(mapa, w, h, ancho, alto, o.suavizado);
        const tipo = ImagenBasica.formatoDeSalida(f.blob.type);
        return { ...f, blob: await aBlob(lienzo, tipo, 0.92), nombre: conExtension(f.nombre, tipo) };
      } finally { mapa.close(); }
    },
  });

  // ---------- 14. Convertir formato de imagen ----------
  F.registrar({
    id: 'convertir-imagen',
    icono: '🔁',
    nombre: { es: 'Convertir formato de imagen', en: 'Convert image format' },
    descripcion: { es: 'Guarda la imagen como PNG, JPG o WebP. En JPG lo transparente queda blanco.', en: 'Saves the image as PNG, JPG or WebP. In JPG, transparent areas turn white.' },
    entra: ['imagen'], sale: 'imagen', forma: 'uno',
    opciones: [
      { id: 'formato', tipo: 'select', defecto: 'image/jpeg', etiqueta: { es: 'Formato', en: 'Format' },
        valores: [{ valor: 'image/png', es: 'PNG', en: 'PNG' }, { valor: 'image/jpeg', es: 'JPG', en: 'JPG' }, { valor: 'image/webp', es: 'WebP', en: 'WebP' }] },
      { id: 'calidad', tipo: 'numero', defecto: 85, min: 10, max: 100, paso: 5, etiqueta: { es: 'Calidad JPG/WebP (%)', en: 'JPG/WebP quality (%)' } },
    ],
    async preparar() { await js('imagen-basica.js'); },
    async ejecutar(f, o) {
      const mapa = await createImageBitmap(f.blob);
      try {
        const blob = await ImagenBasica.convertir(mapa, o.formato, o.calidad / 100);
        if (!blob) throw new Error('toBlob');
        return { ...f, blob, nombre: conExtension(f.nombre, blob.type) };   // si el navegador no sabe WebP, sale PNG
      } finally { mapa.close(); }
    },
  });

  // ---------- 15. Quitar metadatos (EXIF) ----------
  F.registrar({
    id: 'quitar-metadatos',
    icono: '🧹',
    nombre: { es: 'Quitar metadatos EXIF', en: 'Remove EXIF metadata' },
    descripcion: { es: 'Saca cámara, fecha y ubicación GPS. Un JPG derecho queda idéntico, sin recomprimir.', en: 'Removes camera, date and GPS location. An upright JPG stays identical, without recompressing.' },
    entra: ['imagen'], sale: 'imagen', forma: 'uno',
    opciones: [],
    async preparar() { await js('imagen-basica.js'); },
    async ejecutar(f) {
      const archivo = comoArchivo(f);
      const mapa = await createImageBitmap(f.blob);   // con la orientación ya aplicada, por si hay que redibujar
      try {
        const blob = await ImagenBasica.quitarMetadatos(archivo, mapa);
        return { ...f, blob, nombre: conExtension(f.nombre, blob.type) };
      } finally { mapa.close(); }
    },
  });

  // ---------- 16 a 19. Retoques de PDF ----------
  const abrirPdfLib = async (f) => PDFLib.PDFDocument.load(await f.blob.arrayBuffer());
  const prepararRetoques = async () => { await PDF_LIB(); await js('miniaturas.js'); await js('pdf-retoques.js'); };

  F.registrar({
    id: 'rotar-pdf',
    icono: '↻',
    nombre: { es: 'Rotar PDF', en: 'Rotate PDF' },
    descripcion: { es: 'Gira todas las páginas.', en: 'Rotates every page.' },
    entra: ['pdf'], sale: 'pdf', forma: 'uno',
    opciones: [
      { id: 'grados', tipo: 'select', defecto: '90', etiqueta: { es: 'Girar', en: 'Rotate' },
        valores: [{ valor: '90', es: '90° a la derecha', en: '90° right' }, { valor: '180', es: '180°', en: '180°' }, { valor: '270', es: '90° a la izquierda', en: '90° left' }] },
    ],
    preparar: prepararRetoques,
    async ejecutar(f, o) {
      const pdf = await abrirPdfLib(f);
      PdfRetoques.rotar(pdf, Number(o.grados));
      return fichaPdf(f, await pdf.save());
    },
  });

  F.registrar({
    id: 'comprimir-pdf',
    icono: '🗜️',
    nombre: { es: 'Comprimir PDF', en: 'Compress PDF' },
    descripcion: { es: 'Recomprime las fotos a color del PDF; el texto no cambia. Si no baja de peso, queda el original.', en: 'Recompresses the color photos in the PDF; the text does not change. If it does not get lighter, the original stays.' },
    entra: ['pdf'], sale: 'pdf', forma: 'uno',
    opciones: [
      { id: 'calidad', tipo: 'numero', defecto: 65, min: 20, max: 90, paso: 5, etiqueta: { es: 'Calidad de las fotos (%)', en: 'Photo quality (%)' } },
      { id: 'ladoMaximo', tipo: 'select', defecto: '1600', etiqueta: { es: 'Tamaño de las fotos', en: 'Photo size' },
        valores: [{ valor: '0', es: 'Sin reducir', en: 'Do not reduce' }, { valor: '2500', es: 'Hasta 2500 px (impresión)', en: 'Up to 2500 px (print)' },
          { valor: '1600', es: 'Hasta 1600 px (pantalla)', en: 'Up to 1600 px (screen)' }, { valor: '1000', es: 'Hasta 1000 px (lo más liviano)', en: 'Up to 1000 px (lightest)' }] },
    ],
    preparar: prepararRetoques,
    async ejecutar(f, o) {
      const { bytes } = await PdfRetoques.optimizar(await f.blob.arrayBuffer(), { calidad: o.calidad / 100, ladoMaximo: Number(o.ladoMaximo) });
      return fichaPdf(f, bytes);
    },
  });

  F.registrar({
    id: 'numerar-pdf',
    icono: '#️⃣',
    nombre: { es: 'Numerar páginas', en: 'Number pages' },
    descripcion: { es: 'Agrega el número de página en la esquina elegida.', en: 'Adds the page number in the chosen corner.' },
    entra: ['pdf'], sale: 'pdf', forma: 'uno',
    opciones: [
      { id: 'posicion', tipo: 'select', defecto: 'bottom-center', etiqueta: { es: 'Posición', en: 'Position' },
        valores: [
          { valor: 'bottom-right', es: 'Abajo a la derecha', en: 'Bottom right' }, { valor: 'bottom-center', es: 'Abajo al centro', en: 'Bottom center' },
          { valor: 'bottom-left', es: 'Abajo a la izquierda', en: 'Bottom left' }, { valor: 'top-right', es: 'Arriba a la derecha', en: 'Top right' },
          { valor: 'top-center', es: 'Arriba al centro', en: 'Top center' }, { valor: 'top-left', es: 'Arriba a la izquierda', en: 'Top left' }] },
      { id: 'inicio', tipo: 'numero', defecto: 1, min: 1, max: 99999, etiqueta: { es: 'Empezar en', en: 'Start at' } },
    ],
    preparar: prepararRetoques,
    async ejecutar(f, o) {
      const pdf = await abrirPdfLib(f);
      await PdfRetoques.numerar(pdf, { inicio: o.inicio, posicion: o.posicion });
      return fichaPdf(f, await pdf.save());
    },
  });

  F.registrar({
    id: 'marca-agua-pdf',
    icono: '💧',
    nombre: { es: 'Marca de agua en PDF', en: 'Watermark on PDF' },
    descripcion: { es: 'Escribe un texto en diagonal y semitransparente sobre cada página.', en: 'Writes a diagonal, semi-transparent text over every page.' },
    entra: ['pdf'], sale: 'pdf', forma: 'uno',
    opciones: [
      { id: 'texto', tipo: 'texto', defecto: 'CONFIDENCIAL', max: 80, etiqueta: { es: 'Texto', en: 'Text' } },
      { id: 'tamano', tipo: 'numero', defecto: 60, min: 20, max: 150, etiqueta: { es: 'Tamaño máximo de letra', en: 'Maximum font size' } },
      { id: 'opacidad', tipo: 'numero', defecto: 15, min: 5, max: 50, etiqueta: { es: 'Opacidad (%)', en: 'Opacity (%)' } },
      { id: 'repeticiones', tipo: 'numero', defecto: 1, min: 1, max: 20, etiqueta: { es: 'Repeticiones por página', en: 'Copies per page' } },
    ],
    preparar: prepararRetoques,
    async ejecutar(f, o) {
      const pdf = await abrirPdfLib(f);
      await PdfRetoques.marcaDeAgua(pdf, { texto: o.texto, tamanoMaximo: o.tamano, opacidad: o.opacidad / 100, repeticiones: o.repeticiones });
      return fichaPdf(f, await pdf.save());
    },
  });
})();

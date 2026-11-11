/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Tachado de texto en un PDF (herramienta "Tachar texto de un PDF"). Tapar con un rectángulo negro NO borra nada: el texto sigue debajo y se
// puede copiar. Por eso, cada página con coincidencias se dibuja con PDF.js, se pintan las cajas negras sobre lo encontrado y la página
// entera se reemplaza por esa imagen: el texto original deja de existir en el archivo. Las páginas sin coincidencias se copian tal cual, y de
// todas se quitan anotaciones, miniaturas y metadatos (pueden contener el texto, o arrastrar la página original). Al final se vuelve a leer
// el PDF resultante y se comprueba que ya no se encuentra nada; si queda algo, no se entrega.
//
// Lo que NO puede hacer: encontrar texto que no es texto (letras convertidas en dibujo, imágenes, escaneos). La página lo avisa.
//
//   PdfTachado.buscar(pdf, reglas, { alProgreso })              → { aciertos, paginasSinTexto, paginasConImagenes }   (pdf: documento de PDF.js)
//       aciertos: [{ id, pagina (desde 1), texto, cajas: [[ {x,y} ×4 ]] }]  coordenadas en puntos de la página tal como se ve (con su giro)
//   PdfTachado.generar(bytes, pdf, aciertos, { reglas, dpi, todo, permitidas, alProgreso }) → { bytes, todo }
//       "aciertos" son los que SÍ se tachan; "permitidas" es cuántas coincidencias se dejan a propósito sin tachar (las que la persona
//       destildó) y todavía se encuentran en el resultado. Si el PDF tiene capas opcionales (pueden estar ocultas) se convierte entero
//       ("todo" vuelve true). Lanza (codigo 'quedan') si después de tachar se sigue encontrando más de "permitidas".
//   PdfTachado.contar(pdf, reglas)                              → cuántas coincidencias quedan en un documento de PDF.js
// "reglas" son RegExp globales (PatronBusqueda.crear). Requiere pdf-lib (PDFLib) y PDF.js (abrirPdf, de pdfjs.js).

const PdfTachado = (() => {
  const PIXELES_MAXIMOS = 24e6;   // una página no se dibuja con más píxeles que esto (se baja la resolución)
  const ACIERTOS_MAXIMOS = 2000;  // más coincidencias que esto es una búsqueda mal planteada (y un navegador trabado)
  // Margen de la caja, en "em" (el tamaño de la letra): a cada lado del texto, y hacia abajo (colas de g, y, p) y hacia arriba (mayúsculas,
  // tildes). El ancho de cada letra se mide con la fuente del PDF; el margen cubre lo que esa medida no acierta (kerning, fuente sustituida).
  const MARGEN_LADO = 0.3, BAJADA = 0.3, SUBIDA = 1.0;
  const CLAVES_A_QUITAR = ['Annots', 'Metadata', 'PieceInfo', 'Thumb', 'AA', 'B', 'StructParents'];

  let lienzoMedir = null;
  function contextoMedir() {
    if (!lienzoMedir) lienzoMedir = document.createElement('canvas').getContext('2d');
    return lienzoMedir;
  }

  // A∘B para matrices [a, b, c, d, e, f] (primero B, después A), como pdfjsLib.Util.transform.
  const multiplicar = (A, B) => [
    A[0] * B[0] + A[2] * B[1], A[1] * B[0] + A[3] * B[1],
    A[0] * B[2] + A[2] * B[3], A[1] * B[2] + A[3] * B[3],
    A[0] * B[4] + A[2] * B[5] + A[4], A[1] * B[4] + A[3] * B[5] + A[5],
  ];

  // Texto de una página como una sola cadena, con el mapa de cada carácter a su fragmento. Un fin de línea cuenta como un espacio, para que
  // un nombre partido en dos renglones se encuentre igual. Cada fragmento se pasa a NFC (la "é" suelta de un PDF hecho en Mac es "e" + tilde).
  async function textoDePagina(pagina, conFuentes) {
    const vista = pagina.getViewport({ scale: 1 });
    const contenido = await pagina.getTextContent();
    let imagenes = false;
    if (conFuentes) {
      // Las fuentes del PDF se cargan en el navegador al armar la lista de operadores; con ellas se miden las letras.
      const lista = await pagina.getOperatorList();
      const OPS = (window.pdfjsLib && pdfjsLib.OPS) || {};
      const deImagen = new Set([OPS.paintImageXObject, OPS.paintInlineImageXObject, OPS.paintImageMaskXObject, OPS.paintJpegXObject].filter((x) => x !== undefined));
      imagenes = lista.fnArray.some((f) => deImagen.has(f));
      const familias = new Set(Object.values(contenido.styles || {}).map((e) => e.fontFamily).filter(Boolean));
      await Promise.all([...familias].map((f) => (document.fonts ? document.fonts.load('16px ' + f).catch(() => {}) : null)));
    }
    let texto = '';
    const trozos = [];
    for (const it of contenido.items) {
      if (typeof it.str !== 'string') continue;
      const s = it.str.normalize('NFC');
      const inicio = texto.length;
      texto += s;
      trozos.push({ it, s, inicio, fin: texto.length, familia: (contenido.styles && contenido.styles[it.fontName] || {}).fontFamily });
      if (it.hasEOL) texto += ' ';
    }
    return { texto, trozos, vista, imagenes };
  }

  // Las cuatro esquinas, en coordenadas de la vista, de la parte [desde, hasta) del texto s de un fragmento.
  function esquinas(vista, tr, desde, hasta) {
    const it = tr.it, s = tr.s;
    const m = multiplicar(vista.transform, it.transform);
    const largoLinea = Math.hypot(m[0], m[1]) || 1;
    const alto = Math.hypot(m[2], m[3]) || 1;
    const ux = m[0] / largoLinea, uy = m[1] / largoLinea;        // sentido del texto
    const vx = m[2] / alto, vy = m[3] / alto;                    // "hacia arriba" de la letra (la vista ya invirtió el eje y)
    const largo = s.length || 1;
    let ancho = it.width;
    if (!(ancho > 0)) ancho = alto * 0.5 * largo;                // fuente sin anchos: una letra mide más o menos media altura
    // Dónde empieza y dónde termina la parte, como fracción del ancho del fragmento: midiendo las letras con la fuente del PDF (así una "i"
    // y una "W" no valen lo mismo); si no se puede medir, reparto parejo.
    let f0 = desde / largo, f1 = hasta / largo;
    if (tr.familia && it.dir !== 'rtl') {
      try {
        const ctx = contextoMedir();
        ctx.font = '100px ' + tr.familia;
        const total = ctx.measureText(s).width;
        if (total > 0) {
          f0 = desde <= 0 ? 0 : ctx.measureText(s.slice(0, desde)).width / total;
          f1 = hasta >= largo ? 1 : ctx.measureText(s.slice(0, hasta)).width / total;
        }
      } catch (error) { /* se queda el reparto parejo */ }
    }
    if (it.dir === 'rtl') { f0 = 0; f1 = 1; }                    // texto de derecha a izquierda: se cubre el fragmento entero
    const a = f0 * ancho - MARGEN_LADO * alto, b = f1 * ancho + MARGEN_LADO * alto;
    const y0 = -BAJADA * alto, y1 = SUBIDA * alto;
    const punto = (u, h) => ({ x: m[4] + ux * u + vx * h, y: m[5] + uy * u + vy * h });
    return [punto(a, y0), punto(b, y0), punto(b, y1), punto(a, y1)];
  }

  async function buscar(pdf, reglas, opciones) {
    const { alProgreso } = opciones || {};
    const aciertos = [];
    const sinTexto = [], conImagenes = [];
    for (let n = 1; n <= pdf.numPages; n++) {
      if (alProgreso) alProgreso(n, pdf.numPages);
      const pagina = await pdf.getPage(n);
      const { texto, trozos, vista, imagenes } = await textoDePagina(pagina, true);
      if (imagenes) conImagenes.push(n);
      if (!texto.trim()) { sinTexto.push(n); pagina.cleanup(); continue; }
      for (const regla of reglas) {
        regla.lastIndex = 0;
        for (const m of texto.matchAll(regla)) {
          if (m[0].length === 0) continue;
          const desde = m.index, hasta = m.index + m[0].length;
          const cajas = [];
          for (const tr of trozos) {
            const d = Math.max(desde, tr.inicio), h = Math.min(hasta, tr.fin);
            if (d >= h || !tr.s.length) continue;
            cajas.push(esquinas(vista, tr, d - tr.inicio, h - tr.inicio));
          }
          if (cajas.length) aciertos.push({ id: 0, pagina: n, texto: m[0], cajas });
          if (aciertos.length > ACIERTOS_MAXIMOS) throw Object.assign(new Error('demasiadas coincidencias'), { codigo: 'demasiados', maximo: ACIERTOS_MAXIMOS });
        }
      }
      pagina.cleanup();
    }
    aciertos.sort((x, y) => x.pagina - y.pagina);
    aciertos.forEach((a, i) => { a.id = i; });
    return { aciertos, paginasSinTexto: sinTexto, paginasConImagenes: conImagenes };
  }

  // Dibuja la página en un canvas (sin anotaciones: su texto no se busca y se vería) y pinta las cajas negras. Devuelve { canvas, ancho, alto } (en puntos).
  async function dibujarPagina(pagina, aciertos, dpi) {
    const vista1 = pagina.getViewport({ scale: 1 });
    let escala = dpi / 72;
    if (vista1.width * vista1.height * escala * escala > PIXELES_MAXIMOS) escala = Math.sqrt(PIXELES_MAXIMOS / (vista1.width * vista1.height));
    const vista = pagina.getViewport({ scale: escala });
    const canvas = document.createElement('canvas');
    canvas.width = Math.ceil(vista.width);
    canvas.height = Math.ceil(vista.height);
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await pagina.render({ canvasContext: ctx, viewport: vista, annotationMode: 0 /* AnnotationMode.DISABLE */ }).promise;
    ctx.fillStyle = '#000';
    for (const a of aciertos) {
      for (const caja of a.cajas) {
        ctx.beginPath();
        caja.forEach((p, i) => { if (i === 0) ctx.moveTo(p.x * escala, p.y * escala); else ctx.lineTo(p.x * escala, p.y * escala); });
        ctx.closePath();
        ctx.fill();
      }
    }
    return { canvas, ancho: vista1.width, alto: vista1.height };
  }

  async function contar(pdf, reglas) {
    let total = 0;
    for (let n = 1; n <= pdf.numPages; n++) {
      const pagina = await pdf.getPage(n);
      const { texto } = await textoDePagina(pagina, false);
      for (const regla of reglas) { regla.lastIndex = 0; for (const m of texto.matchAll(regla)) if (m[0].length) total++; }
      pagina.cleanup();
    }
    return total;
  }

  // Suelta un documento de PDF.js: sus datos y su hilo de PDF.js, así el PDF original (sin tachar) no queda en memoria.
  // PDF.js 6 ya no tiene pdf.destroy(): se destruye la tarea de carga (2026-11-10; antes solo corría cleanup(), que
  // suelta las páginas pero deja vivo el documento). destroy() de versiones viejas y cleanup() quedan de respaldo.
  function liberar(pdf) {
    try {
      if (pdf.loadingTask && pdf.loadingTask.destroy) pdf.loadingTask.destroy();
      else if (pdf.destroy) pdf.destroy();
      else if (pdf.cleanup) pdf.cleanup();
    } catch (error) { /* ya estaba liberado */ }
  }

  // ¿Tiene capas opcionales (pueden estar ocultas)? Al copiar una página a un PDF nuevo se pierde cuáles estaban apagadas.
  async function tieneCapas(pdf) {
    try {
      const config = await pdf.getOptionalContentConfig();
      for (const _grupo of config) return true;
    } catch (error) { /* sin capas, o esta versión no lo informa */ }
    return false;
  }

  async function generar(bytes, pdf, aciertos, opciones) {
    const { dpi = 200, alProgreso, reglas, permitidas = 0 } = opciones || {};
    if (!reglas || !reglas.length) throw new Error('Faltan las reglas para comprobar el resultado');
    const todo = Boolean(opciones && opciones.todo) || await tieneCapas(pdf);
    const { PDFDocument, PDFName } = PDFLib;
    const origen = await PDFDocument.load(bytes, { updateMetadata: false });
    const nuevo = await PDFDocument.create();
    const porPagina = new Map();
    for (const a of aciertos) { if (!porPagina.has(a.pagina)) porPagina.set(a.pagina, []); porPagina.get(a.pagina).push(a); }
    const paginasOrigen = origen.getPages();
    for (let n = 1; n <= pdf.numPages; n++) {
      if (alProgreso) alProgreso(n, pdf.numPages);
      const marcas = porPagina.get(n) || [];
      if (marcas.length || todo) {
        const pagina = await pdf.getPage(n);
        const { canvas, ancho, alto } = await dibujarPagina(pagina, marcas, dpi);
        pagina.cleanup();
        const png = new Uint8Array(await (await new Promise((r) => canvas.toBlob(r, 'image/png'))).arrayBuffer());
        canvas.width = canvas.height = 0;
        const imagen = await nuevo.embedPng(png);
        const hoja = nuevo.addPage([ancho, alto]);
        hoja.drawImage(imagen, { x: 0, y: 0, width: ancho, height: alto });
      } else {
        // Las claves se quitan ANTES de copiar: copyPages sigue todo lo que la página referencia y se llevaría, aunque después se borre la
        // referencia, el comentario, el destino de un enlace (con la página original que se tachó) o los metadatos.
        for (const clave of CLAVES_A_QUITAR) paginasOrigen[n - 1].node.delete(PDFName.of(clave));
        const [copia] = await nuevo.copyPages(origen, [n - 1]);
        nuevo.addPage(copia);
      }
    }
    nuevo.setProducer('Mini Cookies');
    nuevo.setCreator('Mini Cookies');
    nuevo.setCreationDate(new Date());
    nuevo.setModificationDate(new Date());
    const salida = await nuevo.save({ useObjectStreams: false });

    // Comprobación final: se lee el PDF que se va a entregar y no puede quedar nada de lo buscado.
    const verificacion = await abrirPdf(salida.slice(0));
    const quedan = await contar(verificacion, reglas);
    liberar(verificacion);
    if (quedan > permitidas) throw Object.assign(new Error('quedan ' + quedan + ' coincidencias'), { codigo: 'quedan', quedan });
    return { bytes: salida, todo };
  }

  return { buscar, generar, contar, liberar, PIXELES_MAXIMOS, ACIERTOS_MAXIMOS };
})();

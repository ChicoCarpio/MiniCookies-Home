/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
/*
 * Traducir un archivo de principio a fin: lo comparten la página «Traducir documentos» y el módulo de Flujos (una sola copia).
 * Carga lo pesado solo cuando hace falta (el modelo la primera vez, PDF.js, pdf-lib y las fuentes con un PDF, el OCR con una página escaneada).
 * Necesita antes vendor.js, traductor-documentos.js, traductor-docx.js, traductor-epub.js y traductor-hilos.js (los trae `preparar`).
 *
 *   await TraducirArchivo.preparar()                                    carga los módulos (idempotente)
 *   TraducirArchivo.tipoDe(nombre)                                      'docx' | 'pdf' | 'epub' | 'md' | 'txt' | null
 *   const r = await TraducirArchivo.traducir(bytes, tipo, {
 *       destino, origen,          idiomas: es en pt fr de it zh ja ko; `origen` solo para las páginas escaneadas de un PDF
 *       modo,                     PDF: 'quitar' | 'tapar'
 *       progreso(h, total),       oraciones hechas
 *       cargando(bytes, total),   descarga del modelo
 *       cancelado() })            -> true para parar
 *     -> { bytes, mime, ext, estadistica, docx? }                       docx: el PDF traducido como Word (solo PDF)
 *   TraducirArchivo.cerrar()                                            libera los hilos del traductor
 */
(function () {
  const IDIOMAS = ['es', 'en', 'pt', 'fr', 'de', 'it', 'zh', 'ja', 'ko'];
  const OCR_DE = { es: 'spa', en: 'eng', pt: 'por', fr: 'fra', de: 'deu', it: 'ita', zh: 'chi_sim', ja: 'jpn', ko: 'kor' };
  const CJK = { zh: 'NotoSansSC', ja: 'NotoSansJP', ko: 'NotoSansKR' };
  // Archivos de Interpres Phone V1 (nombre elegido por el usuario el 2026-11-16). Mientras el modelo no exista, MODELO es null y se avisa
  // que el traductor todavía no está disponible.
  const MODELO = null; // { vocabulario: 'interpres-vocabulario.json', codificador: 'interpres-phone-v1-codificador-int8.onnx', decodificador: 'interpres-phone-v1-decodificador-int8.onnx', megas: 45 }

  const TIPOS = [[/\.docx$/i, 'docx'], [/\.pdf$/i, 'pdf'], [/\.epub$/i, 'epub'], [/\.(md|markdown)$/i, 'md'], [/\.txt$/i, 'txt']];
  const tipoDe = (nombre) => { const t = TIPOS.find(([re]) => re.test(nombre || '')); return t ? t[1] : null; };

  // Este archivo puede correr dentro de Flujos, que tiene su propio idioma: los avisos salen en el de la página.
  const TEXTOS = {
    es: { sinModelo: 'El traductor todavía no está disponible.',
      navegador: 'Tu navegador no puede ejecutar el traductor (necesita WebAssembly y trabajadores en segundo plano). Probá con una versión reciente de Chrome, Edge, Firefox o Safari.',
      mismoIdioma: 'Elegí un idioma distinto del original.', formato: 'Este archivo no es un Word (.docx), PDF, EPUB, Markdown ni texto.' },
    en: { sinModelo: 'The translator is not available yet.',
      navegador: 'Your browser can’t run the translator (it needs WebAssembly and background workers). Try a recent version of Chrome, Edge, Firefox or Safari.',
      mismoIdioma: 'Choose a language different from the original.', formato: 'This file is not a Word (.docx), PDF, EPUB, Markdown or text file.' },
  };
  const txt = (clave) => TEXTOS[(window.idiomaActual && idiomaActual()) === 'en' ? 'en' : 'es'][clave];

  // Las rutas salen de dónde está este archivo (assets/js/): sirve igual desde herramientas/ que desde /en/herramientas/.
  const BASE = new URL('./', document.currentScript.src).href;
  const absoluta = (ruta) => new URL(ruta, location.href).href;
  const cargar = (src) => new Promise((ok, mal) => {
    const s = document.createElement('script');
    s.src = src; s.onload = ok; s.onerror = () => mal(new Error('No cargó ' + src));
    document.head.appendChild(s);
  });
  const faltan = (...nombres) => nombres.filter((n) => !window[n]);

  async function preparar() {
    const modulos = [['XmlEditar', 'xml-editar.js'], ['TraductorDocumentos', 'traductor-documentos.js'], ['TraductorDocx', 'traductor-docx.js'],
      ['TraductorEpub', 'traductor-epub.js'], ['TraductorHilos', 'traductor-hilos.js']];
    if (!window.ZipPropio) await cargar(BASE + 'zip-propio.js');
    for (const [global, archivo] of modulos) if (!window[global]) await cargar(BASE + archivo);
  }

  let traductor = null, creando = null;
  async function obtenerTraductor(cargando) {
    if (window.traductorDePrueba) return window.traductorDePrueba; // solo lo ponen las pruebas
    if (traductor) return traductor;
    if (!MODELO) throw new Error(txt('sinModelo'));
    if (typeof Worker === 'undefined' || typeof WebAssembly === 'undefined') throw new Error(txt('navegador'));
    if (!creando) {
      creando = TraductorHilos.crear({
        rutaOrt: absoluta(rutaVendor('onnxruntime-web/1.20.1/')), rutaHilo: BASE + 'traductor-hilo.js',
        rutaPiezas: BASE + 'traductor-piezas.js', rutaMotor: BASE + 'traductor-motor.js',
        vocabulario: BASE + '../modelos/' + MODELO.vocabulario, codificador: BASE + '../modelos/' + MODELO.codificador,
        decodificador: BASE + '../modelos/' + MODELO.decodificador,
        hilos: Math.max(1, Math.min(4, (navigator.hardwareConcurrency || 4) - 1)),
        alCargar: (bytes, total) => { if (cargando) cargando(bytes, total || MODELO.megas * 1e6); },
      }).then((t) => { traductor = t; return t; }, (e) => { creando = null; throw e; });
    }
    return creando;
  }
  async function cerrar() {
    const t = traductor;
    traductor = null; creando = null;
    if (t && t.cerrar) await t.cerrar();
  }

  async function pdfListo() {
    if (!window.abrirPdf) await cargar(BASE + 'pdfjs.js');
    if (!window.PDFLib) await cargar(rutaVendor('pdf-lib/1.17.1/pdf-lib.min.js'));
    if (!window.fontkit) await cargar(rutaVendor('@pdf-lib/fontkit/1.1.1/dist/fontkit.umd.min.js'));
    if (!window.EscritorDocx) await cargar(BASE + 'escritor-docx.js');
    if (!window.TraductorPdf || !TraductorPdf.traducir) await cargar(BASE + 'traductor-pdf.js');
  }
  const bytesDe = async (ruta) => new Uint8Array(await (await fetch(ruta)).arrayBuffer());
  async function fuentesPdf(dest) {
    const lib = (n) => bytesDe(rutaVendor(`liberation-fonts/2.1.5/${n}.ttf`));
    const f = { sans: await lib('LiberationSans-Regular'), sansNegrita: await lib('LiberationSans-Bold'),
      serifa: await lib('LiberationSerif-Regular'), serifaNegrita: await lib('LiberationSerif-Bold') };
    if (CJK[dest]) {
      f.cjk = await bytesDe(rutaVendor(`noto-sans-cjk/2.004/${CJK[dest]}-Regular.ttf`));
      f.cjkNegrita = await bytesDe(rutaVendor(`noto-sans-cjk/2.004/${CJK[dest]}-Bold.ttf`));
    }
    return f;
  }
  // El OCR se crea solo si aparece una página escaneada.
  function ocrPerezoso(idioma) {
    let motor = null;
    return {
      async leerPagina(lienzo) {
        if (!motor) {
          if (!window.OcrLibrarius) await cargar(BASE + 'ocr-librarius.js');
          motor = await OcrLibrarius.crear([OCR_DE[idioma]]);
        }
        return motor.leerPagina(lienzo);
      },
      cerrar() { if (motor) motor.cerrar(); },
    };
  }

  async function traducir(entrada, tipo, o) {
    if (!tipo) throw new Error(txt('formato'));
    if (tipo === 'pdf' && o.origen && o.origen === o.destino) throw new Error(txt('mismoIdioma'));
    await preparar();
    const motor = await obtenerTraductor(o.cargando);
    const traducirSegmentos = (segs, d, op = {}) => motor.traducir(segs, d, {
      progreso: (h, total) => { if (o.progreso) o.progreso(h, total); if (op.progreso) op.progreso(h, total); },
      cancelado: o.cancelado || (() => false),
    });
    const op = { traducir: traducirSegmentos, destino: o.destino };
    if (tipo === 'txt' || tipo === 'md') {
      const texto = typeof entrada === 'string' ? entrada : new TextDecoder().decode(entrada);
      const r = await (tipo === 'txt' ? TraductorDocumentos.traducirTexto : TraductorDocumentos.traducirMarkdown)(texto, op);
      return { bytes: new TextEncoder().encode(typeof r === 'string' ? r : r.texto), estadistica: r.estadistica || {},
        mime: tipo === 'md' ? 'text/markdown;charset=utf-8' : 'text/plain;charset=utf-8', ext: tipo };
    }
    if (tipo === 'docx') {
      const r = await TraductorDocx.traducir(entrada, op);
      return { bytes: r.bytes, estadistica: r.estadistica, ext: 'docx', mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' };
    }
    if (tipo === 'epub') {
      const r = await TraductorEpub.traducir(entrada, op);
      return { bytes: r.bytes, estadistica: r.estadistica, ext: 'epub', mime: 'application/epub+zip' };
    }
    await pdfListo();
    const ocr = ocrPerezoso(o.origen || 'es');
    try {
      const r = await TraductorPdf.traducir(new Uint8Array(entrada), { ...op, fuentes: await fuentesPdf(o.destino), ocr, modo: o.modo || 'quitar' });
      if (o.guardando) o.guardando();
      const docx = o.comoDocx === false ? null : await EscritorDocx.generar(TraductorPdf.comoDocx(r.paginas));
      return { bytes: r.bytes, estadistica: r.estadistica, ext: 'pdf', mime: 'application/pdf', docx };
    } finally {
      ocr.cerrar();
    }
  }

  window.TraducirArchivo = { IDIOMAS, tipoDe, preparar, traducir, cerrar, disponible: () => !!MODELO || !!window.traductorDePrueba };
})();

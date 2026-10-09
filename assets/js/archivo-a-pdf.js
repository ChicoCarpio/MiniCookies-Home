/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Un documento de otro formato → PDF, con el mismo motor que usa la
// herramienta de ese formato (no una copia). Lo usa "Unir PDF" para aceptar
// Word, Excel/CSV, PowerPoint, EPUB, Markdown y texto además de PDF e
// imágenes.
//
// Cada motor y sus librerías se cargan recién cuando aparece un archivo de ese
// tipo: juntas pesan varios MB y la mayoría de las veces se unen solo PDF.
//
// Requiere, ya cargados: vendor.js (rutaVendor) y pdf-lib.
// Uso: ArchivoAPdf.formatoDe(archivo) → 'word' | 'excel' | … | null
//      const bytes = await ArchivoAPdf.convertir(archivo, alAvanzar);
//      ArchivoAPdf.EXTENSIONES → '.docx,.xlsx,…' (para el accept del selector)

(function () {
  const TEXTOS = {
    es: {
      'archivo-a-pdf-sin-motor': 'No se pudo cargar el conversor (¿sin conexión?):',
      'archivo-a-pdf-no-soportado': 'Formato no soportado:',
      'archivo-a-pdf-danado': 'el archivo está dañado o no es del formato que dice su extensión',
    },
    en: {
      'archivo-a-pdf-sin-motor': 'Couldn’t load the converter (offline?):',
      'archivo-a-pdf-no-soportado': 'Unsupported format:',
      'archivo-a-pdf-danado': 'the file is damaged or isn’t the format its extension says',
    },
  };
  // Se suman a los textos de la página (los de la página mandan).
  const pagina = window.i18nPagina || (window.i18nPagina = {});
  for (const idioma of Object.keys(TEXTOS)) pagina[idioma] = Object.assign({}, TEXTOS[idioma], pagina[idioma]);

  // Los módulos propios están al lado de este archivo, en assets/js/. Como
  // con rutaVendor, la ruta sale de este script y no de la página (que existe
  // también bajo /en/).
  const base = new URL('./', document.currentScript.src).href;
  const propio = (nombre) => base + nombre;

  const cargas = new Map(); // src → promesa: cada script se pide una sola vez
  function cargarScript(src) {
    if (!cargas.has(src)) {
      cargas.set(src, new Promise((resolver, rechazar) => {
        const s = document.createElement('script');
        s.src = src;
        s.onload = resolver;
        s.onerror = () => {
          cargas.delete(src); // un fallo de red no queda guardado para siempre
          rechazar(new Error(t('archivo-a-pdf-sin-motor') + ' ' + src.split('/').pop().replace(/\?.*$/, '')));
        };
        document.head.appendChild(s);
      }));
    }
    return cargas.get(src);
  }
  async function cargarEnOrden(lista) {
    for (const src of lista) await cargarScript(src);
  }

  // Las rutas de vendor/ van escritas completas: el build las busca así para
  // comprobar que se publican (referencias-vendor.mjs).
  const JSZIP = () => rutaVendor('jszip/3.10.1/jszip.min.js');

  const FORMATOS = {
    word: {
      extensiones: ['docx'],
      async preparar() {
        await cargarEnOrden([JSZIP(), propio('seguridad-html.js'), propio('emf.js'), propio('imagenes-office.js'), propio('html-a-pdf.js')]);
        if (!window.docxPreview) {
          // docx-preview se publica como window.docx, el mismo nombre que la
          // librería docx: se guarda aparte y se devuelve lo que hubiera.
          const anterior = window.docx;
          await cargarScript(rutaVendor('docx-preview/0.4.1/docx-preview.min.js'));
          window.docxPreview = window.docx;
          window.docx = anterior;
        }
        await cargarScript(propio('word-a-pdf.js'));
      },
      convertir: async (archivo) => WordAPdf.aPdf(await WordAPdf.aHtml(archivo)),
    },
    excel: {
      extensiones: ['xlsx', 'xlsm', 'xlsb', 'xls', 'ods', 'csv'],
      preparar: () => cargarEnOrden([
        rutaVendor('xlsx/0.20.3/xlsx.full.min.js'), JSZIP(), rutaVendor('Chart.js/4.5.1/chart.umd.min.js'),
        propio('lector-planillas.js'), propio('excel-a-pdf.js'),
      ]),
      convertir: (archivo) => libroExcelAPdf(archivo),
    },
    powerpoint: {
      extensiones: ['pptx'],
      async preparar() {
        await cargarEnOrden([JSZIP(), rutaVendor('Chart.js/4.5.1/chart.umd.min.js')]);
        if (!window.PptxViewJS) {
          // PptxViewJS pisa el FileReader del navegador con uno propio, y eso
          // rompe JSZip: se devuelve el original (igual que en PowerPoint → PDF).
          const fileReaderNativo = window.FileReader;
          try {
            await cargarScript(rutaVendor('pptxviewjs/1.1.9/dist/PptxViewJS.min.js?v=2'));
          } finally {
            window.FileReader = fileReaderNativo;
          }
        }
        await cargarEnOrden([propio('pdf-texto.js'), propio('pptx-visual.js')]);
      },
      convertir: async (archivo) => (await powerPointAPdf(archivo)).pdf.save(),
    },
    epub: {
      extensiones: ['epub'],
      preparar: () => cargarEnOrden([JSZIP(), propio('seguridad-html.js'), propio('epub.js'), propio('html-a-pdf.js')]),
      convertir: (archivo, alAvanzar) => epubAPdf(archivo, alAvanzar),
    },
    markdown: {
      extensiones: ['md', 'markdown'],
      preparar: () => cargarEnOrden([propio('leer-texto.js'), propio('pdf-texto.js')]),
      convertir: async (archivo) => (await generarPdfDeMarkdown((await leerTextoArchivo(archivo)).texto)).save(),
    },
    texto: {
      extensiones: ['txt'],
      preparar: () => cargarEnOrden([propio('leer-texto.js'), propio('pdf-texto.js')]),
      convertir: async (archivo) => (await generarPdfDeTexto((await leerTextoArchivo(archivo)).texto)).save(),
    },
  };

  function formatoDe(archivo) {
    const extension = (/\.([^.]+)$/.exec(archivo.name) || [])[1];
    if (!extension) return null;
    const minuscula = extension.toLowerCase();
    return Object.keys(FORMATOS).find((f) => FORMATOS[f].extensiones.includes(minuscula)) || null;
  }

  // Devuelve los bytes del PDF. alAvanzar(hoja, total): solo los formatos
  // que se paginan de a poco (EPUB) lo llaman.
  async function convertir(archivo, alAvanzar) {
    const formato = FORMATOS[formatoDe(archivo)];
    if (!formato) throw new Error(t('archivo-a-pdf-no-soportado') + ' ' + archivo.name);
    await formato.preparar();
    try {
      return await formato.convertir(archivo, alAvanzar);
    } catch (error) {
      // Word, Excel, PowerPoint y EPUB son .zip por dentro: si no lo es, JSZip
      // tira un mensaje técnico en inglés (con una URL) que no le sirve a nadie.
      // Lo mismo con un zip que no es del formato (un .docx sin documento
      // adentro): la librería termina leyendo algo que no está, un TypeError.
      if (error instanceof TypeError || /central directory|zip file|corrupted zip|unsupported zip/i.test((error && error.message) || '')) {
        console.warn('No se pudo convertir ' + archivo.name + ':', error);
        throw new Error(t('archivo-a-pdf-danado'));
      }
      throw error;
    }
  }

  const EXTENSIONES = Object.values(FORMATOS).flatMap((f) => f.extensiones.map((e) => '.' + e)).join(',');

  window.ArchivoAPdf = { formatoDe, convertir, EXTENSIONES };
})();

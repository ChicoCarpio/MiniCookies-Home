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
//      const bytes = await ArchivoAPdf.convertir(archivo, alAvanzar); // alAvanzar(fracción 0–1)
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
      // Ya lo puso otro cargador (grafico-suelto.js trae parte del motor de PowerPoint): no repetirlo.
      if (document.querySelector('script[src="' + src + '"]')) cargas.set(src, Promise.resolve());
    }
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

  const JSZIP = () => propio('zip-propio.js');

  const FORMATOS = {
    word: {
      extensiones: ['docx', 'doc', 'odt'],
      async preparar() {
        await cargarEnOrden([
          JSZIP(), propio('xml-ligero.js'), propio('seguridad-html.js'), propio('emf.js'), propio('imagenes-office.js'), propio('html-a-pdf.js'),
          propio('formatos-word.js'), propio('docx-base.js'), propio('docx-estilos.js'), propio('pptx-texto.js'), propio('docx-maquetador.js'),
        ]);
        await cargarScript(propio('word-a-pdf.js'));
      },
      // Dibujar el Word es un solo paso sin avance: la mitad de la barra.
      convertir: async (archivo, alAvanzar) => {
        const documento = await WordAPdf.aHtml(archivo);
        alAvanzar(0.5);
        return WordAPdf.aPdf(documento, (h, th) => alAvanzar(0.5 + 0.5 * (h + 1) / th));
      },
    },
    excel: {
      extensiones: ['xlsx', 'xlsm', 'xlsb', 'xls', 'ods', 'csv'],
      preparar: () => cargarEnOrden([
        JSZIP(), propio('xml-ligero.js'), propio('xlsx-formato.js'), propio('xlsx-estilos.js'), propio('xlsx-modelo.js'),
        propio('csv-planilla.js'), propio('lector-hojas.js'), propio('motor-hojas.js'), propio('excel-a-pdf-motor.js'),
      ]),
      convertir: (archivo, alAvanzar) => ExcelAPdfMotor.libroAPdf(archivo, alAvanzar),
    },
    powerpoint: {
      extensiones: ['pptx'],
      async preparar() {
        await cargarEnOrden([JSZIP()]);
        // El motor propio, con la lista única de pptx-visual.js.
        await cargarEnOrden([propio('pdf-texto.js'), propio('pptx-visual.js')]);
        await cargarEnOrden(SCRIPTS_MOTOR_PPTX.map((x) => (x.startsWith('vendor/') ? rutaVendor(x.slice(7)) : propio(x))));
      },
      // Sin conImagenes: Unir PDF no muestra las diapositivas sueltas ni arma su .zip.
      convertir: async (archivo, alAvanzar) => (await powerPointAPdf(archivo, null, (n, total) => alAvanzar(0.9 * n / total), { conImagenes: false })).pdf.save(),
    },
    epub: {
      extensiones: ['epub'],
      preparar: () => cargarEnOrden([JSZIP(), propio('seguridad-html.js'), propio('epub.js'), propio('html-a-pdf.js')]),
      convertir: (archivo, alAvanzar) => epubAPdf(archivo, (h, th) => alAvanzar((h + 1) / th)),
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

  // Devuelve los bytes del PDF. alAvanzar(fraccion): de 0 a 1 en los formatos
  // que saben cuánto les falta (Word, Excel, PowerPoint, EPUB).
  async function convertir(archivo, alAvanzar = () => {}) {
    const formato = FORMATOS[formatoDe(archivo)];
    if (!formato) throw new Error(t('archivo-a-pdf-no-soportado') + ' ' + archivo.name);
    await formato.preparar();
    try {
      return await formato.convertir(archivo, alAvanzar);
    } catch (error) {
      // Word, Excel, PowerPoint y EPUB son .zip por dentro: si no lo es, el zip
      // propio tira un ErrorZip con un mensaje técnico en inglés que no le sirve
      // a nadie. Lo mismo con un zip que no es del formato (un .docx sin
      // documento adentro): se termina leyendo algo que no está, un TypeError.
      if (error instanceof TypeError || (error && error.name === 'ErrorZip')) {
        console.warn('No se pudo convertir ' + archivo.name + ':', error);
        throw new Error(t('archivo-a-pdf-danado'));
      }
      throw error;
    }
  }

  const EXTENSIONES = Object.values(FORMATOS).flatMap((f) => f.extensiones.map((e) => '.' + e)).join(',');

  window.ArchivoAPdf = { formatoDe, convertir, EXTENSIONES };
})();

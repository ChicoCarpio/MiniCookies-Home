/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Excel / CSV → PDF, compartido por "Excel/CSV → PDF" y "Unir PDF".
// Requiere SheetJS (XLSX), assets/js/lector-planillas.js, pdf-lib, JSZip y
// Chart.js (los gráficos de un .xlsx), y assets/js/vendor.js (rutaVendor).
//
// La tabla se dibuja escribiendo los operadores PDF a mano (ver "Dibujo
// directo"): así un Excel de 20.000 filas entra en memoria. Las funciones son
// globales a propósito: la página de Excel → PDF usa varias por su cuenta
// (leer estilos, gráficos de la hoja) para su vista previa.

(function () {
  const TEXTOS = {
    es: {
      'aviso-caracteres': 'La fuente de respaldo no representa todos los caracteres del archivo ({n} sustituidos); los signos monetarios disponibles se conservaron.',
      'error-sin-datos': 'La hoja no tiene datos.',
      'serie-defecto': 'Serie {n}',
      'leyenda-tipo-no-soportado': ' (tipo de gráfico no soportado — mostrado como barras)',
      'leyenda-grafico': 'Gráfico{titulo}, celda {celda} de "{hoja}"{tipo} — reconstruido a partir de los datos, el estilo puede no ser idéntico al original',
    },
    en: {
      'aviso-caracteres': 'The fallback font does not cover every character in the file ({n} replaced); available currency symbols were preserved.',
      'error-sin-datos': 'The sheet has no data.',
      'serie-defecto': 'Series {n}',
      'leyenda-tipo-no-soportado': ' (chart type not supported — shown as bars)',
      'leyenda-grafico': 'Chart{titulo}, cell {celda} of "{hoja}"{tipo} — rebuilt from the data, the style may not be identical to the original',
    },
  };
  // Se suman a los textos de la página (los de la página mandan).
  const pagina = window.i18nPagina || (window.i18nPagina = {});
  for (const idioma of Object.keys(TEXTOS)) pagina[idioma] = Object.assign({}, TEXTOS[idioma], pagina[idioma]);
})();

let caracteresPdfAdaptados = 0;
let datosFuenteUnicodeCache = null;

const TAMANOS_PAGINA_PTS = {
  '4a0': [1682, 2378].map(milimetros => milimetros * 72 / 25.4),
  '2a0': [1189, 1682].map(milimetros => milimetros * 72 / 25.4),
  'arch-e': [914.4, 1219.2].map(milimetros => milimetros * 72 / 25.4),
  a0: [2384, 3370], a1: [1684, 2384], a2: [1191, 1684], a3: [842, 1191],
  'ansi-e': [863.6, 1117.6].map(milimetros => milimetros * 72 / 25.4),
  a4: [595, 842], a5: [420, 595], a6: [298, 420],
  b4: [729, 1032], b5: [516, 729], carta: [612, 792], legal: [612, 1008],
  oficio: [612, 936], tabloide: [792, 1224], ejecutivo: [522, 756],
  'media-carta': [396, 612], folio: [612, 936],
};

const URL_FONTKIT = rutaVendor('@pdf-lib/fontkit/1.1.1/dist/fontkit.umd.min.js?v=2');
// Noto Sans 2.015 (google/fonts@2984c57), copiada en /vendor/ junto a su
// licencia OFL. Antes venía de raw.githubusercontent.com, rama main: un
// dominio no declarado y una versión que podía cambiar sola.
const URL_FUENTE_ALTERNATIVA = rutaVendor('noto-sans/2.015/NotoSans-VF.ttf');

async function cargarDatosFuenteUnicode() {
  if (datosFuenteUnicodeCache) return datosFuenteUnicodeCache;

  if (!window.fontkit) {
    await new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = URL_FONTKIT;
      script.onload = resolve;
      script.onerror = () => reject(new Error('No se pudo cargar fontkit.'));
      document.head.appendChild(script);
    });
  }

  const respuesta = await fetch(URL_FUENTE_ALTERNATIVA);
  if (!respuesta.ok) throw new Error('No se pudo descargar Noto Sans.');
  const bytes = new Uint8Array(await respuesta.arrayBuffer());
  datosFuenteUnicodeCache = { bytes, fuente: window.fontkit.create(bytes) };
  return datosFuenteUnicodeCache;
}

function columnasVisiblesDeHoja(hoja, inicio, fin, incluirOcultas = true) {
  const columnas = [];
  const metadatos = hoja['!cols'] || [];
  for (let columna = inicio; columna <= fin; columna++) {
    if (incluirOcultas || !metadatos[columna]?.hidden) columnas.push(columna);
  }
  return columnas;
}

const CODIGOS_MONEDA = new Map([
  ['؋', 'AFN'], ['৳', 'BDT'], ['₿', 'BTC'], ['₵', 'GHS'], ['₡', 'CRC'],
  ['₫', 'VND'], ['₭', 'LAK'], ['₦', 'NGN'], ['₱', 'PHP'], ['₲', 'PYG'],
  ['₴', 'UAH'], ['₪', 'ILS'], ['₸', 'KZT'], ['₺', 'TRY'], ['₮', 'MNT'],
  ['₩', 'KRW'], ['₹', 'INR'], ['₽', 'RUB'], ['﷼', 'SAR'],
]);

function necesitaFuenteUnicode(fuente, matriz) {
  return matriz.some(fila => fila.some(celda => {
    try {
      fuente.encodeText(String(celda.texto ?? '').normalize('NFC'));
      return false;
    } catch {
      return true;
    }
  }));
}

// Qué fuente dibuja cada carácter y cuánto mide, calculado UNA vez por
// carácter y por fuente. Antes se volvía a preguntar a la fuente por cada
// letra cada vez que se medía un texto, y partir una celda en líneas medía
// de nuevo la línea entera por cada palabra agregada: con 20.000 filas el PDF
// completo tardaba más de 4 minutos. El caché vive mientras viva la fuente,
// que se crea de nuevo en cada PDF, así que nunca mezcla dos documentos.
const cacheCaracteresPdf = new WeakMap();

function infoCaracterPdf(fuente, caracter, fuenteUnicode, metadatosUnicode) {
  let porCaracter = cacheCaracteresPdf.get(fuente);
  if (!porCaracter) {
    porCaracter = new Map();
    cacheCaracteresPdf.set(fuente, porCaracter);
  }
  let info = porCaracter.get(caracter);
  if (info) return info;

  let fuenteCaracter = fuente;
  let texto = caracter;
  let adaptado = false;
  try {
    fuente.encodeText(caracter);
  } catch {
    if (fuenteUnicode && metadatosUnicode && metadatosUnicode.hasGlyphForCodePoint(caracter.codePointAt(0))) {
      fuenteCaracter = fuenteUnicode;
    } else {
      texto = CODIGOS_MONEDA.get(caracter) || '?';
      adaptado = true;
    }
  }
  // El ancho se suma letra por letra, sin el ajuste entre pares (kerning)
  // que incluye widthOfTextAtSize: drawText tampoco lo aplica al dibujar, así
  // que esto coincide mejor con lo que realmente ocupa el texto en el PDF.
  info = { texto, fuente: fuenteCaracter, ancho: fuenteCaracter.widthOfTextAtSize(texto, 1), adaptado };
  porCaracter.set(caracter, info);
  return info;
}

function prepararTextoPdf(fuente, texto, fuenteUnicode, metadatosUnicode) {
  const caracteres = [];
  const textoNormalizado = String(texto ?? '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/\n/g, ' ');
  for (const caracter of textoNormalizado.normalize('NFC')) {
    caracteres.push(infoCaracterPdf(fuente, caracter, fuenteUnicode, metadatosUnicode));
  }
  return caracteres;
}

function agruparCaracteresPdf(caracteres) {
  const grupos = [];
  caracteres.forEach(({ texto, fuente, ancho }) => {
    const ultimo = grupos[grupos.length - 1];
    if (ultimo && ultimo.fuente === fuente) {
      ultimo.texto += texto;
      ultimo.ancho += ancho;
    } else {
      grupos.push({ texto, fuente, ancho });
    }
  });
  return grupos;
}

function medirTextoPdf(caracteres, tamano) {
  let ancho = 0;
  for (const caracter of caracteres) ancho += caracter.ancho;
  return ancho * tamano;
}

function truncarTexto(fuente, texto, tamano, anchoMaximo, fuenteUnicode, metadatosUnicode) {
  const caracteres = prepararTextoPdf(fuente, texto, fuenteUnicode, metadatosUnicode);
  if (medirTextoPdf(caracteres, tamano) <= anchoMaximo) return caracteres;

  const elipsis = infoCaracterPdf(fuente, '…', fuenteUnicode, metadatosUnicode);
  let ancho = elipsis.ancho * tamano;
  let cantidad = 0;
  while (cantidad < caracteres.length && ancho + caracteres[cantidad].ancho * tamano <= anchoMaximo) {
    ancho += caracteres[cantidad].ancho * tamano;
    cantidad++;
  }
  return caracteres.slice(0, cantidad).concat(elipsis);
}

// --- Dibujo directo sobre el contenido de la página ---
// pdf-lib guarda cada drawText/drawLine/drawRectangle como varios objetos en
// memoria hasta el save(): con 20.000 filas eso pasaba de 1,5 GB y con más la
// pestaña moría. Acá se escriben los mismos operadores PDF que generaría
// pdf-lib, pero como texto, y al cerrar cada página se comprimen en un único
// flujo. Lo que queda en memoria por página es ese flujo, no miles de objetos.
// Las fuentes siguen siendo de pdf-lib (encodeText marca los glifos usados, así
// que el subconjunto de Noto Sans se arma igual al guardar).
function crearLienzoPdf(pdf, pagina) {
  return { pdf, pagina, partes: [], fuentes: new Map() };
}

function numeroPdf(n) {
  return Number.isInteger(n) ? String(n) : n.toFixed(3).replace(/\.?0+$/, '');
}

function colorPdf(color) {
  return `${numeroPdf(color.red)} ${numeroPdf(color.green)} ${numeroPdf(color.blue)}`;
}

function nombreFuenteEnLienzo(lienzo, fuente) {
  let nombre = lienzo.fuentes.get(fuente);
  if (!nombre) {
    nombre = lienzo.pagina.node.newFontDictionary(fuente.name, fuente.ref).toString();
    lienzo.fuentes.set(fuente, nombre);
  }
  return nombre;
}

function textoEnLienzo(lienzo, texto, fuente, tamano, x, y, color) {
  lienzo.partes.push(`BT ${nombreFuenteEnLienzo(lienzo, fuente)} ${numeroPdf(tamano)} Tf ${colorPdf(color)} rg ${numeroPdf(x)} ${numeroPdf(y)} Td ${fuente.encodeText(texto).toString()} Tj ET`);
}

function lineaEnLienzo(lienzo, x1, y1, x2, y2, grosor, color) {
  lienzo.partes.push(`${colorPdf(color)} RG ${numeroPdf(grosor)} w ${numeroPdf(x1)} ${numeroPdf(y1)} m ${numeroPdf(x2)} ${numeroPdf(y2)} l S`);
}

function rectanguloEnLienzo(lienzo, x, y, ancho, alto, color) {
  lienzo.partes.push(`${colorPdf(color)} rg ${numeroPdf(x)} ${numeroPdf(y)} ${numeroPdf(ancho)} ${numeroPdf(alto)} re f`);
}

function cerrarLienzoPdf(lienzo) {
  if (!lienzo.partes.length) return;
  const flujo = lienzo.pdf.context.flateStream(`q\n${lienzo.partes.join('\n')}\nQ\n`);
  lienzo.pagina.node.addContentStream(lienzo.pdf.context.register(flujo));
  lienzo.partes = [];
}

function dibujarTextoPdf(lienzo, caracteres, opciones) {
  // Se cuentan al dibujar, no al medir: un mismo texto se mide varias veces
  // y el aviso inflaba la cantidad de caracteres sustituidos.
  for (const caracter of caracteres) if (caracter.adaptado) caracteresPdfAdaptados++;
  let x = opciones.x;
  agruparCaracteresPdf(caracteres).forEach(grupo => {
    textoEnLienzo(lienzo, grupo.texto, grupo.fuente, opciones.size, x, opciones.y, opciones.color);
    x += grupo.ancho * opciones.size;
  });
  return x - opciones.x;
}

function celdaVacia(texto) {
  return String(texto ?? '').trim() === '';
}

function colorDesdeHex(hex) {
  if (!hex || typeof hex !== 'string') return null;
  const limpio = hex.length === 8 ? hex.slice(2) : hex; // quita el canal alfa si viene ARGB
  if (limpio.length !== 6) return null;
  const r = parseInt(limpio.slice(0, 2), 16) / 255;
  const g = parseInt(limpio.slice(2, 4), 16) / 255;
  const b = parseInt(limpio.slice(4, 6), 16) / 255;
  if ([r, g, b].some(Number.isNaN)) return null;
  return { r, g, b };
}

// El color real de la letra solo se conoce en .xlsx/.xlsm (ver
// leerFuentesDeCeldas). En los demás formatos, si la celda tiene relleno, se
// elige blanco o negro según qué tan clara es, para no terminar con texto
// oscuro invisible sobre un relleno oscuro (muy común en encabezados).
function colorTextoLegible(fondo) {
  const luminancia = 0.299 * fondo.r + 0.587 * fondo.g + 0.114 * fondo.b;
  return luminancia > 0.6 ? { r: 0, g: 0, b: 0 } : { r: 1, g: 1, b: 1 };
}

function leerCeldaConEstilo(hoja, fila, columna) {
  const referencia = XLSX.utils.encode_cell({ r: fila, c: columna });
  const celda = hoja[referencia];
  const estilo = (celda && celda.s) || {};
  // Letra: la lee leerFuentesDeCeldas() del XML (solo .xlsx/.xlsm).
  const fuente = (hoja['!fuentes'] && hoja['!fuentes'].get(referencia)) || {};
  const relleno = estilo.patternType === 'solid' ? colorDesdeHex(estilo.fgColor && estilo.fgColor.rgb) : null;
  return {
    texto: celda ? XLSX.utils.format_cell(celda) : '',
    negrita: !!fuente.negrita,
    cursiva: !!fuente.cursiva,
    subrayado: !!fuente.subrayado,
    color: fuente.color || (relleno ? colorTextoLegible(relleno) : null),
    relleno,
  };
}

function elegirFuente(negrita, cursiva) {
  if (negrita && cursiva) return window.fuenteNegritaCursivaPdf || window.fuenteNegritaCursivaPdf;
  if (negrita) return window.fuenteNegritaPdf || window.fuenteNegritaPdf;
  if (cursiva) return window.fuenteCursivaPdf || window.fuenteCursivaPdf;
  return window.fuenteBasePdf || window.fuenteBasePdf;
}

// Lee estilos celda por celda porque sheet_to_json los descarta.
//
// Recorre solo las celdas que EXISTEN en la hoja, no el rectángulo de '!ref'.
// Antes armaba un objeto por cada posición del rango, vacía o no, y después lo
// filtraba: con un archivo grande —o con un '!ref' inflado porque alguien le
// dio formato a columnas enteras— la pestaña se quedaba sin memoria antes de
// llegar a dibujar nada. Ahora primero se decide qué filas y columnas tienen
// texto y recién después se leen esas, hasta 'maxFilas' (la vista previa no
// necesita la hoja entera).
const CELDA_SIN_DATOS = Object.freeze({ texto: '', negrita: false, cursiva: false, subrayado: false, color: null, relleno: null });

function leerHojaConEstilos(hoja, maxFilas = Infinity, incluirOcultas = true) {
  if (!hoja['!ref']) return [];
  const rango = XLSX.utils.decode_range(hoja['!ref']);
  const conTexto = [];
  const columnasConTexto = new Set();
  for (const referencia in hoja) {
    if (referencia[0] === '!') continue;
    const { r, c } = XLSX.utils.decode_cell(referencia);
    if (r < rango.s.r || r > rango.e.r || c < rango.s.c || c > rango.e.c) continue;
    if (celdaVacia(XLSX.utils.format_cell(hoja[referencia]))) continue;
    conTexto.push(r, c);
    columnasConTexto.add(c);
  }

  const columnasAConservar = [...columnasConTexto]
    .filter(c => incluirOcultas || !hoja['!cols']?.[c]?.hidden)
    .sort((x, y) => x - y);
  const conservar = new Set(columnasAConservar);
  const filasConTexto = new Set();
  for (let i = 0; i < conTexto.length; i += 2) {
    if (conservar.has(conTexto[i + 1])) filasConTexto.add(conTexto[i]);
  }
  const filas = [...filasConTexto].sort((x, y) => x - y);

  const resultado = filas.slice(0, maxFilas).map(r => columnasAConservar.map(c => (
    hoja[XLSX.utils.encode_cell({ r, c })] ? leerCeldaConEstilo(hoja, r, c) : CELDA_SIN_DATOS
  )));
  resultado.indicesColumnas = columnasAConservar;
  resultado.totalFilas = filas.length;
  return resultado;
}

// --- Gráficos embebidos (solo .xlsx) ---
// SheetJS no expone la definición de gráficos, así que se vuelve a abrir el
// mismo archivo con JSZip y se parsean a mano el OPF de relaciones y el XML
// de dibujo — mismo patrón ya usado en este proyecto para PowerPoint/Word/
// EPUB (ver administrar-diapositivas.html, epub.js). No se incluyen imágenes
// embebidas, únicamente celdas (tabla) y gráficos. No se intenta ubicar cada
// gráfico exactamente sobre su celda de origen: la tabla ya filtra filas/
// columnas vacías y usa columnas de ancho uniforme, así que esas posiciones
// no coinciden con las del Excel original. En cambio, se agregan como
// páginas propias después de la tabla, con una leyenda que indica en qué
// celda estaban ancladas.

function carpetaDe(ruta) {
  const partes = ruta.split('/');
  partes.pop();
  return partes.length ? partes.join('/') + '/' : '';
}

function resolverRutaRelativa(carpetaBase, relativa) {
  const partes = (carpetaBase + relativa).split('/');
  const resultado = [];
  partes.forEach((parte) => {
    if (parte === '..') resultado.pop();
    else if (parte !== '.' && parte !== '') resultado.push(parte);
  });
  return resultado.join('/');
}

function rutaRelsDe(rutaXml) {
  const partes = rutaXml.split('/');
  const archivo = partes.pop();
  return [...partes, '_rels', archivo + '.rels'].join('/');
}

async function leerXmlDeZip(zip, ruta) {
  const entrada = zip.file(ruta);
  if (!entrada) return null;
  const texto = await entrada.async('text');
  return new DOMParser().parseFromString(texto, 'application/xml');
}

async function leerRelaciones(zip, rutaRels) {
  const doc = await leerXmlDeZip(zip, rutaRels);
  const mapa = new Map();
  if (!doc) return mapa;
  [...doc.getElementsByTagName('Relationship')].forEach((rel) => {
    mapa.set(rel.getAttribute('Id'), rel.getAttribute('Target'));
  });
  return mapa;
}

// nombre de hoja -> ruta del sheetN.xml dentro del zip
async function mapaHojaARutaXml(zip) {
  const workbookDoc = await leerXmlDeZip(zip, 'xl/workbook.xml');
  const mapa = new Map();
  if (!workbookDoc) return mapa;
  const relaciones = await leerRelaciones(zip, 'xl/_rels/workbook.xml.rels');
  [...workbookDoc.getElementsByTagName('sheet')].forEach((hoja) => {
    const nombre = hoja.getAttribute('name');
    const rId = hoja.getAttribute('r:id');
    const target = rId && relaciones.get(rId);
    if (nombre && target) mapa.set(nombre, resolverRutaRelativa('xl/', target));
  });
  return mapa;
}

// --- Negrita, cursiva, subrayado y color de la letra (solo .xlsx/.xlsm) ---
// SheetJS (versión gratuita) lee de styles.xml la lista de formatos y de
// fuentes (libro.Styles), y hasta resuelve los colores de tema, pero en cada
// celda deja solo el relleno: no dice qué formato usa. Ese dato está en el XML
// de la hoja (<c r="B2" s="5">), así que se lee el atributo y se cruza con la
// lista. Se guarda solo lo que cambia el dibujo, en hoja['!fuentes'] (las
// claves con "!" no son celdas: nadie más las recorre).
function estiloFuenteDeFormato(estilos, indice) {
  const formato = estilos.CellXf[indice];
  const fuente = formato && estilos.Fonts[formato.fontId || 0];
  if (!fuente) return null;
  const color = fuente.color && !fuente.color.auto ? colorDesdeHex(fuente.color.rgb) : null;
  const relleno = estilos.Fills && estilos.Fills[formato.fillId || 0];
  const negro = color && color.r === 0 && color.g === 0 && color.b === 0;
  // Letra negra común: no hace falta guardarla, salvo sobre un relleno (si no,
  // colorTextoLegible() la cambiaría por blanca sobre un fondo oscuro).
  const colorQueImporta = color && (!negro || (relleno && relleno.patternType === 'solid'));
  if (!fuente.bold && !fuente.italic && !fuente.underline && !colorQueImporta) return null;
  return { negrita: !!fuente.bold, cursiva: !!fuente.italic, subrayado: !!fuente.underline, color: colorQueImporta ? color : null };
}

async function leerFuentesDeCeldas(libro, bytes) {
  const estilos = libro.Styles;
  if (!estilos || !estilos.CellXf || !estilos.Fonts) return;
  const zip = await JSZip.loadAsync(bytes);
  const rutas = await mapaHojaARutaXml(zip);
  const porFormato = new Map(); // índice de formato -> estilo o null, calculado una vez
  for (const nombre of libro.SheetNames) {
    const hoja = libro.Sheets[nombre];
    const entrada = rutas.has(nombre) && zip.file(rutas.get(nombre));
    if (!hoja || !entrada) continue;
    const xml = await entrada.async('text');
    const fuentes = new Map();
    // Expresiones regulares y no DOMParser: una hoja de 60.000 filas pesa
    // decenas de MB como XML, y armar su árbol duplicaba la memoria.
    for (const [, atributos] of xml.matchAll(/<c\b([^>]*)>/g)) {
      const s = /\ss="(\d+)"/.exec(atributos);
      if (!s) continue;
      if (!porFormato.has(s[1])) porFormato.set(s[1], estiloFuenteDeFormato(estilos, +s[1]));
      const estilo = porFormato.get(s[1]);
      if (!estilo) continue;
      const referencia = /\sr="([A-Z]+\d+)"/.exec(atributos);
      if (referencia) fuentes.set(referencia[1], estilo);
    }
    hoja['!fuentes'] = fuentes;
  }
}

function primerHijoDirecto(nodo, tag) {
  if (!nodo) return null;
  return [...nodo.children].find((n) => n.tagName === tag) || null;
}

function celdaDesdeAncla(ancla) {
  const from = ancla.getElementsByTagName('xdr:from')[0];
  if (!from) return null;
  const col = parseInt(from.getElementsByTagName('xdr:col')[0]?.textContent || '0', 10);
  const row = parseInt(from.getElementsByTagName('xdr:row')[0]?.textContent || '0', 10);
  return XLSX.utils.encode_cell({ r: row, c: col });
}

async function objetosDeHoja(zip, rutaSheetXml) {
  const vacio = { graficos: [] };
  const hojaDoc = await leerXmlDeZip(zip, rutaSheetXml);
  if (!hojaDoc) return vacio;
  const nodoDrawing = hojaDoc.getElementsByTagName('drawing')[0];
  if (!nodoDrawing) return vacio;

  const relacionesHoja = await leerRelaciones(zip, rutaRelsDe(rutaSheetXml));
  const targetDrawing = relacionesHoja.get(nodoDrawing.getAttribute('r:id'));
  if (!targetDrawing) return vacio;

  const rutaDrawing = resolverRutaRelativa(carpetaDe(rutaSheetXml), targetDrawing);
  const drawingDoc = await leerXmlDeZip(zip, rutaDrawing);
  if (!drawingDoc) return vacio;

  const relacionesDrawing = await leerRelaciones(zip, rutaRelsDe(rutaDrawing));
  const carpetaDrawing = carpetaDe(rutaDrawing);

  const graficos = [];
  const anclas = [...drawingDoc.getElementsByTagName('xdr:twoCellAnchor'), ...drawingDoc.getElementsByTagName('xdr:oneCellAnchor')];

  for (const ancla of anclas) {
    const graphicFrame = ancla.getElementsByTagName('xdr:graphicFrame')[0];
    if (!graphicFrame) continue;
    const chartNodo = graphicFrame.getElementsByTagName('c:chart')[0];
    const rIdGrafico = chartNodo && chartNodo.getAttribute('r:id');
    const targetGrafico = rIdGrafico && relacionesDrawing.get(rIdGrafico);
    if (!targetGrafico) continue;
    const rutaGrafico = resolverRutaRelativa(carpetaDrawing, targetGrafico);
    const graficoDoc = await leerXmlDeZip(zip, rutaGrafico);
    if (graficoDoc) graficos.push({ doc: graficoDoc, celda: celdaDesdeAncla(ancla) });
  }

  return { graficos };
}

const PALETA_GRAFICOS = ['#4f7cff', '#22c55e', '#f97316', '#ef4444', '#a855f7', '#14b8a6', '#eab308', '#64748b'];

function valoresDeCache(nodo) {
  if (!nodo) return [];
  const pts = [...nodo.getElementsByTagName('c:pt')];
  const maxIdx = pts.reduce((m, p) => Math.max(m, parseInt(p.getAttribute('idx') || '0', 10)), -1);
  const arr = new Array(maxIdx + 1).fill('');
  pts.forEach((p) => {
    const idx = parseInt(p.getAttribute('idx') || '0', 10);
    const v = p.getElementsByTagName('c:v')[0];
    arr[idx] = v ? v.textContent : '';
  });
  return arr;
}

function nombreSerie(ser, indice) {
  const tx = primerHijoDirecto(ser, 'c:tx');
  const vDirecto = tx && primerHijoDirecto(tx, 'c:v');
  if (vDirecto) return vDirecto.textContent;
  const strRef = tx && primerHijoDirecto(tx, 'c:strRef');
  const cache = strRef && primerHijoDirecto(strRef, 'c:strCache');
  const valores = cache ? valoresDeCache(cache) : [];
  return valores[0] || t('serie-defecto').replace('{n}', indice + 1);
}

function categoriasDe(ser) {
  const cat = primerHijoDirecto(ser, 'c:cat');
  if (!cat) return [];
  const ref = primerHijoDirecto(cat, 'c:strRef') || primerHijoDirecto(cat, 'c:numRef');
  if (ref) {
    const cache = primerHijoDirecto(ref, 'c:strCache') || primerHijoDirecto(ref, 'c:numCache');
    return cache ? valoresDeCache(cache) : [];
  }
  // Categorías de varios niveles (ej. una tabla de Excel con dos columnas
  // como categoría, "lugar" + "casa") vienen en c:multiLvlStrRef en vez de
  // c:strRef — sin este caso las categorías quedaban vacías y el gráfico de
  // barras no dibujaba nada (0 categorías = barras de ancho 0 en Chart.js).
  const multiRef = primerHijoDirecto(cat, 'c:multiLvlStrRef');
  const multiCache = multiRef && primerHijoDirecto(multiRef, 'c:multiLvlStrCache');
  if (!multiCache) return [];
  const niveles = [...multiCache.children].filter((n) => n.tagName === 'c:lvl').map((nivel) => valoresDeCache(nivel));
  const cantidad = niveles.reduce((max, nivel) => Math.max(max, nivel.length), 0);
  return Array.from({ length: cantidad }, (_, i) => niveles.map((nivel) => nivel[i]).filter(Boolean).join(' / '));
}

function valoresDe(ser) {
  const val = primerHijoDirecto(ser, 'c:val');
  const ref = val && primerHijoDirecto(val, 'c:numRef');
  const cache = ref && primerHijoDirecto(ref, 'c:numCache');
  const bruto = cache ? valoresDeCache(cache) : [];
  return bruto.map((v) => parseFloat(v) || 0);
}

function colorDeSerie(ser, indice) {
  const spPr = primerHijoDirecto(ser, 'c:spPr');
  const solidFill = spPr && primerHijoDirecto(spPr, 'a:solidFill');
  const srgb = solidFill && primerHijoDirecto(solidFill, 'a:srgbClr');
  return srgb ? '#' + srgb.getAttribute('val') : PALETA_GRAFICOS[indice % PALETA_GRAFICOS.length];
}

function extraerTituloGrafico(doc) {
  const chartNodo = doc.getElementsByTagName('c:chart')[0];
  const titleNodo = chartNodo && primerHijoDirecto(chartNodo, 'c:title');
  if (!titleNodo) return '';
  return [...titleNodo.getElementsByTagName('a:t')].map((r) => r.textContent).join('').trim();
}

const TIPOS_GRAFICO_SOPORTADOS = [
  ['c:barChart', 'bar'], ['c:bar3DChart', 'bar'],
  ['c:lineChart', 'line'], ['c:line3DChart', 'line'],
  ['c:areaChart', 'line'], ['c:area3DChart', 'line'],
  ['c:pieChart', 'pie'], ['c:pie3DChart', 'pie'],
  ['c:doughnutChart', 'doughnut'],
  ['c:radarChart', 'radar'],
  ['c:scatterChart', 'line'],
];

function parsearGrafico(doc) {
  let elementoChart = null, tipo = 'bar', soportado = true;
  for (const [tag, t] of TIPOS_GRAFICO_SOPORTADOS) {
    const nodo = doc.getElementsByTagName(tag)[0];
    if (nodo) { elementoChart = nodo; tipo = t; break; }
  }
  if (!elementoChart) {
    const plotArea = doc.getElementsByTagName('c:plotArea')[0];
    const cualquiera = plotArea && [...plotArea.children].find((n) => /Chart$/.test(n.tagName));
    if (!cualquiera) return null;
    elementoChart = cualquiera;
    soportado = false;
  }

  const seriesNodos = [...elementoChart.getElementsByTagName('c:ser')];
  if (seriesNodos.length === 0) return null;

  let categorias = [];
  const series = seriesNodos.map((ser, indice) => {
    const cats = categoriasDe(ser);
    if (cats.length > categorias.length) categorias = cats;
    return { nombre: nombreSerie(ser, indice), valores: valoresDe(ser), color: colorDeSerie(ser, indice) };
  });

  // Si el gráfico no tiene un título explícito, Excel muestra automáticamente
  // el nombre de la única serie como título cuando hay solo una — se imita
  // ese comportamiento en vez de dejar el título vacío.
  const titulo = extraerTituloGrafico(doc) || (series.length === 1 ? series[0].nombre : '');

  return { tipo, soportado, titulo, categorias, series };
}

async function renderizarGraficoComoPng(datosGrafico) {
  const canvas = document.createElement('canvas');
  canvas.width = 900;
  canvas.height = 550;
  // Chart.js necesita el canvas dentro del documento para medir/dibujar bien
  // (con uno desconectado del DOM, en algunos navegadores el gráfico sale en
  // blanco) — se agrega fuera de pantalla y se saca apenas se captura la imagen.
  canvas.style.position = 'fixed';
  canvas.style.left = '-99999px';
  canvas.style.top = '0';
  document.body.appendChild(canvas);

  const esCircular = datosGrafico.tipo === 'pie' || datosGrafico.tipo === 'doughnut';
  const datasets = esCircular
    ? [{ data: datosGrafico.series[0]?.valores || [], backgroundColor: datosGrafico.categorias.map((_, i) => PALETA_GRAFICOS[i % PALETA_GRAFICOS.length]) }]
    : datosGrafico.series.map((serie) => ({
        label: serie.nombre,
        data: serie.valores,
        backgroundColor: serie.color,
        borderColor: serie.color,
        fill: false,
      }));

  const chart = new Chart(canvas.getContext('2d'), {
    type: datosGrafico.tipo,
    data: { labels: datosGrafico.categorias, datasets },
    options: {
      responsive: false,
      animation: false,
      plugins: {
        title: { display: !!datosGrafico.titulo, text: datosGrafico.titulo, font: { size: 18 } },
        legend: { display: datasets.length > 1 || esCircular },
      },
    },
  });

  await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
  chart.destroy();
  canvas.remove();
  return blob.arrayBuffer();
}

function anchoColumnaDesdeHoja(hoja, indiceColumna, fallback = 120) {
  const metadato = hoja['!cols']?.[indiceColumna] || {};
  if (Number(metadato.wpx) > 0) return Number(metadato.wpx);
  if (Number(metadato.wch) > 0) return Math.max(50, Math.ceil(Number(metadato.wch) * 7 + 10));
  return fallback;
}

function dividirTextoEnLineas(fuente, texto, anchoMaximo, tamano, fuenteUnicode, metadatosUnicode) {
  const textoBase = String(texto ?? '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  if (!textoBase) return [''];

  const bloques = textoBase.split('\n');
  const lineas = [];

  for (const bloque of bloques) {
    if (!bloque) {
      lineas.push('');
      continue;
    }

    const tokens = bloque.split(/(\s+)/g).filter((token) => token.length > 0);
    let lineaActual = '';
    let anchoLinea = 0;

    // El ancho de la línea se va sumando palabra por palabra, en vez de volver
    // a medir la línea entera cada vez que se le agrega una.
    for (const token of tokens) {
      const anchoToken = medirTextoPdf(prepararTextoPdf(fuente, token, fuenteUnicode, metadatosUnicode), tamano);
      if (anchoLinea + anchoToken <= anchoMaximo || lineaActual === '') {
        lineaActual += token;
        anchoLinea += anchoToken;
        continue;
      }

      if (lineaActual) {
        lineas.push(lineaActual.trimEnd());
      }
      lineaActual = token;
      anchoLinea = anchoToken;
    }

    if (lineaActual) lineas.push(lineaActual.trimEnd());
  }

  return lineas.length ? lineas : [''];
}

function prepararLineasFilaPDF(fila, anchosColumnas, tamanoFuente, fuenteUnicode, metadatosUnicode, forzarNegrita = false, limitarAOcho = false) {
  return fila.map((celdaInfo, indice) => {
    const ancho = Math.max(12, (anchosColumnas[indice] || 80) - 8);
    const fuenteCelda = elegirFuente(celdaInfo.negrita || forzarNegrita, celdaInfo.cursiva);
    const lineas = dividirTextoEnLineas(fuenteCelda, celdaInfo.texto, ancho, tamanoFuente, fuenteUnicode, metadatosUnicode);
    return limitarAOcho ? lineas.slice(0, 8) : lineas;
  });
}

function calcularAlturaFilaPDF(lineasFila, tamanoFuente) {
  const mayorNumeroLineas = Math.max(1, ...lineasFila.map(lineas => lineas.length));
  return Math.max(tamanoFuente * 1.5 + 8, mayorNumeroLineas * tamanoFuente * 1.35 + 8);
}

function dibujarCeldaPdf(pagina, celdaInfo, x, y, ancho, altoFila, tamanoFuente, colorPorDefecto, fuente, fuenteUnicode, metadatosUnicode, forzarNegrita, rgbFn, lineasPreparadas = null, inicioLinea = 0, cantidadLineas = Infinity) {
  const negrita = celdaInfo.negrita || forzarNegrita;
  const fuenteElegida = elegirFuente(negrita, celdaInfo.cursiva);
  const color = celdaInfo.color ? rgbFn(celdaInfo.color.r, celdaInfo.color.g, celdaInfo.color.b) : colorPorDefecto;
  if (celdaInfo.relleno) {
    rectanguloEnLienzo(pagina, x, y - altoFila, ancho, altoFila, rgbFn(celdaInfo.relleno.r, celdaInfo.relleno.g, celdaInfo.relleno.b));
  }

  const anchoUtilCelda = Math.max(12, ancho - 8);
  const lineas = lineasPreparadas || dividirTextoEnLineas(fuenteElegida, celdaInfo.texto, anchoUtilCelda, tamanoFuente, fuenteUnicode, metadatosUnicode);
  const interlineado = tamanoFuente * 1.35;
  const inicioY = y - 4;
  const lineasVisibles = lineas.slice(inicioLinea, inicioLinea + cantidadLineas);
  lineasVisibles.forEach((linea, indice) => {
    const yLinea = inicioY - (indice * interlineado);
    const textoRecortado = truncarTexto(fuenteElegida, linea, tamanoFuente, anchoUtilCelda, fuenteUnicode, metadatosUnicode);
    dibujarTextoPdf(pagina, textoRecortado, { x: x + 4, y: yLinea - tamanoFuente + 2, size: tamanoFuente, color });
  });

  if (celdaInfo.subrayado && inicioLinea + lineasVisibles.length >= lineas.length) {
    const ultimaLinea = lineasVisibles[lineasVisibles.length - 1] || '';
    const textoUltimaLinea = truncarTexto(fuenteElegida, ultimaLinea, tamanoFuente, anchoUtilCelda, fuenteUnicode, metadatosUnicode);
    const anchoTexto = medirTextoPdf(textoUltimaLinea, tamanoFuente);
    const yLinea = inicioY - ((lineasVisibles.length - 1) * interlineado) - 1;
    lineaEnLienzo(pagina, x + 4, yLinea, x + 4 + anchoTexto, yLinea, 0.6, color);
  }
}

// --- Una hoja → PDF ---
// Lo usan "Excel/CSV → PDF" (con lo que elige la persona, y para su vista
// previa) y "Unir PDF" (con las opciones de fábrica, hoja por hoja). Antes
// vivía dentro de excel-csv-a-pdf.html leyendo sus controles.
//
// opciones:
//   libro, nombreHoja       el libro de SheetJS y la hoja a dibujar
//   bytesArchivo, esXlsx    los gráficos solo existen en .xlsx/.xlsm
//   rangos                  null = la hoja entera; un array = solo esas celdas
//   incluirOcultas          columnas ocultas sí o no
//   horizontal              orientación
//   tamanoPagina            [ancho, alto] en puntos, en vertical
//   tamanoFuente            7, 9 u 11
//   limitarAOchoLineas      corta las celdas largas a 8 líneas
//   limiteFilas             la vista previa no necesita la hoja entera
//   alAvanzar(n, total)     cada 100 filas; alGuardar() antes de escribir el PDF
// Devuelve { bytes, avisos } (avisos: textos ya traducidos).
async function construirPdfDeHoja({
  libro, nombreHoja, bytesArchivo = null, esXlsx = false, rangos = null,
  incluirOcultas = true, horizontal = false, tamanoPagina = TAMANOS_PAGINA_PTS.a4,
  tamanoFuente = 9, limitarAOchoLineas = false, limiteFilas = Infinity,
  alAvanzar = null, alGuardar = null,
}) {
  caracteresPdfAdaptados = 0;
  const mensajesAdvertencia = [];
  const soloCeldas = Array.isArray(rangos);
  const hoja = libro.Sheets[nombreHoja];
  let matriz;
  let encabezado;
  let cuerpoFilas;
  let cantidadColumnas;
  let totalFilasCuerpo;
  if (soloCeldas) {
    if (rangos.length === 0) throw new Error(t('error-sin-rangos'));
    const rangosConCeldas = rangos.map(rango => {
      const columnas = columnasVisiblesDeHoja(hoja, rango.s.c, rango.e.c, incluirOcultas);
      return { rango, columnas, ancho: columnas.length };
    }).filter(item => item.ancho > 0);
    if (rangosConCeldas.length === 0) throw new Error(t('error-rangos-ocultos'));
    cantidadColumnas = Math.max(...rangosConCeldas.map(item => item.ancho));
    const celdaVacia = { texto: '', negrita: false, cursiva: false, subrayado: false, color: null, relleno: null };
    cuerpoFilas = [];
    totalFilasCuerpo = rangosConCeldas.reduce((suma, { rango }, indice) => suma + (indice > 0 ? 1 : 0) + rango.e.r - rango.s.r + 1, 0);
    rangosConCeldas.forEach(({ rango, columnas, ancho }, indiceRango) => {
      if (indiceRango > 0 && cuerpoFilas.length < limiteFilas) cuerpoFilas.push(Array(cantidadColumnas).fill(celdaVacia));
      for (let fila = rango.s.r; fila <= rango.e.r && cuerpoFilas.length < limiteFilas; fila++) {
        const celdas = [];
        for (const columna of columnas) celdas.push(leerCeldaConEstilo(hoja, fila, columna));
        while (celdas.length < cantidadColumnas) celdas.push(celdaVacia);
        cuerpoFilas.push(celdas);
      }
    });
    matriz = cuerpoFilas;
    encabezado = [];
  } else {
    matriz = leerHojaConEstilos(hoja, limiteFilas + 1, incluirOcultas);
    if (matriz.length === 0) throw Object.assign(new Error(t('error-sin-datos')), { sinDatos: true });
    encabezado = matriz[0];
    cuerpoFilas = matriz.slice(1);
    cantidadColumnas = encabezado.length;
    totalFilasCuerpo = matriz.totalFilas - 1;
  }
  if (cuerpoFilas.length < totalFilasCuerpo) {
    mensajesAdvertencia.push(t('aviso-vista-parcial').replace('{n}', cuerpoFilas.length).replace('{total}', totalFilasCuerpo));
  }

  const { PDFDocument, StandardFonts, rgb } = PDFLib;
  const pdf = await PDFDocument.create();
  const fuente = await pdf.embedFont(StandardFonts.Helvetica);
  const fuenteNegrita = await pdf.embedFont(StandardFonts.HelveticaBold);
  const fuenteCursiva = await pdf.embedFont(StandardFonts.HelveticaOblique);
  const fuenteNegritaCursiva = await pdf.embedFont(StandardFonts.HelveticaBoldOblique);
  window.fuenteBasePdf = fuente;
  window.fuenteNegritaPdf = fuenteNegrita;
  window.fuenteCursivaPdf = fuenteCursiva;
  window.fuenteNegritaCursivaPdf = fuenteNegritaCursiva;
  let fuenteUnicode = null;
  let metadatosFuenteUnicode = null;
  if (necesitaFuenteUnicode(fuente, matriz)) {
    try {
      const datos = await cargarDatosFuenteUnicode();
      pdf.registerFontkit(window.fontkit);
      fuenteUnicode = await pdf.embedFont(datos.bytes, { subset: true });
      metadatosFuenteUnicode = datos.fuente;
    } catch (errorFuente) {
      console.warn('No se pudo cargar la fuente de símbolos; se usarán sustituciones:', errorFuente);
    }
  }

  const margen = 30;
  const [anchoVertical, altoVertical] = tamanoPagina;
  if (!Number.isFinite(anchoVertical) || !Number.isFinite(altoVertical) || anchoVertical < 70 || altoVertical < 70) {
    throw new Error(t('error-tamano-personalizado'));
  }
  const [anchoPagina, altoPagina] = horizontal ? [altoVertical, anchoVertical] : [anchoVertical, altoVertical];
  const anchoUtil = anchoPagina - margen * 2;

  const anchosColumnas = soloCeldas
    ? Array.from({ length: cantidadColumnas }, (_, indice) => Math.max(70, anchoUtil / Math.max(1, cantidadColumnas)))
    : Array.from({ length: cantidadColumnas }, (_, indice) => anchoColumnaDesdeHoja(hoja, matriz.indicesColumnas?.[indice] ?? indice, 120));

  const totalAnchoColumnas = anchosColumnas.reduce((suma, ancho) => suma + ancho, 0);
  const escalaColumnas = totalAnchoColumnas > anchoUtil ? anchoUtil / totalAnchoColumnas : 1;
  const anchosAUsar = anchosColumnas.map((ancho) => ancho * escalaColumnas);

  if (anchosAUsar.some((ancho) => ancho < 25)) {
    mensajesAdvertencia.push(t('aviso-columnas').replace('{n}', cantidadColumnas));
  }

  function dibujarEncabezado(pagina, y, altura) {
    let x = margen;
    if (!soloCeldas) rectanguloEnLienzo(pagina, margen, y - altura, anchoUtil, altura, rgb(0.9, 0.93, 0.98));
    encabezado.forEach((celdaInfo, indice) => {
      const ancho = anchosAUsar[indice] || 80;
      dibujarCeldaPdf(pagina, celdaInfo, x, y, ancho, altura, tamanoFuente, rgb(0.1, 0.1, 0.2), fuente, fuenteUnicode, metadatosFuenteUnicode, true, rgb, lineasEncabezado[indice]);
      x += ancho;
    });
  }

  function dibujarLineasColumnas(pagina, yTope, yFondo) {
    let x = margen;
    for (let i = 0; i < anchosAUsar.length; i++) {
      x += anchosAUsar[i];
      lineaEnLienzo(pagina, x, yTope, x, yFondo, 0.5, rgb(0.85, 0.85, 0.85));
    }
  }

  const lineasEncabezado = prepararLineasFilaPDF(encabezado || [], anchosAUsar, tamanoFuente, fuenteUnicode, metadatosFuenteUnicode, true, limitarAOchoLineas);
  const alturaEncabezado = Math.max(26, calcularAlturaFilaPDF(lineasEncabezado, tamanoFuente));
  const interlineadoFila = tamanoFuente * 1.35;
  let pagina;
  let y;
  let yFilasInicio;
  let hayFilasEnPagina = false;

  function abrirPaginaTabla() {
    pagina = crearLienzoPdf(pdf, pdf.addPage([anchoPagina, altoPagina]));
    y = altoPagina - margen;
    if (!soloCeldas) {
      dibujarEncabezado(pagina, y, alturaEncabezado);
      y -= alturaEncabezado;
    }
    yFilasInicio = y;
    hayFilasEnPagina = false;
  }

  function cerrarPaginaTabla() {
    const yTopeTabla = yFilasInicio + (soloCeldas ? 0 : alturaEncabezado);
    dibujarLineasColumnas(pagina, yTopeTabla, y);
    lineaEnLienzo(pagina, margen, y, margen + anchoUtil, y, 0.5, rgb(0.85, 0.85, 0.85));
    cerrarLienzoPdf(pagina);
  }

  abrirPaginaTabla();

  for (let indiceFila = 0; indiceFila < cuerpoFilas.length; indiceFila++) {
    const fila = cuerpoFilas[indiceFila];
    // Se calculan fila por fila: tenerlas todas armadas de antemano era otra
    // copia entera de la hoja en memoria.
    const lineasFila = prepararLineasFilaPDF(fila, anchosAUsar, tamanoFuente, fuenteUnicode, metadatosFuenteUnicode, false, limitarAOchoLineas);
    const totalLineas = Math.max(1, ...lineasFila.map(lineas => lineas.length));
    let inicioLinea = 0;

    while (inicioLinea < totalLineas) {
      const espacioDisponible = y - margen;
      let maxLineasSegmento = Math.floor((espacioDisponible - 8) / interlineadoFila);
      if (maxLineasSegmento < 1 && hayFilasEnPagina) {
        cerrarPaginaTabla();
        abrirPaginaTabla();
        continue;
      }
      maxLineasSegmento = Math.max(1, maxLineasSegmento);
      const cantidadLineas = Math.min(totalLineas - inicioLinea, maxLineasSegmento);
      const alturaSegmento = cantidadLineas * interlineadoFila + 8;

      let x = margen;
      fila.forEach((celdaInfo, indice) => {
        if (indice >= anchosAUsar.length) return;
        const ancho = anchosAUsar[indice] || 80;
        dibujarCeldaPdf(
          pagina, celdaInfo, x, y, ancho, alturaSegmento, tamanoFuente,
          rgb(0.15, 0.15, 0.15), fuente, fuenteUnicode, metadatosFuenteUnicode,
          false, rgb, lineasFila[indice], inicioLinea, cantidadLineas
        );
        x += ancho;
      });
      y -= alturaSegmento;
      lineaEnLienzo(pagina, margen, y, margen + anchoUtil, y, 0.5, rgb(0.85, 0.85, 0.85));
      hayFilasEnPagina = true;
      inicioLinea += cantidadLineas;

      if (inicioLinea < totalLineas) {
        cerrarPaginaTabla();
        abrirPaginaTabla();
      }
    }

    if ((indiceFila + 1) % 100 === 0 || indiceFila + 1 === cuerpoFilas.length) {
      if (alAvanzar) alAvanzar(indiceFila + 1, cuerpoFilas.length);
      await new Promise(resolve => setTimeout(resolve, 0));
    }
  }
  // Siempre: es lo que dibuja las líneas de columna de la última página y
  // vuelca su contenido al PDF.
  cerrarPaginaTabla();

  // Gráficos embebidos en la hoja (solo aplica a .xlsx — .xls y .csv no
  // tienen esta estructura). Si algo falla acá no se pierde la tabla ya
  // generada: se avisa y se sigue con lo que sí se pudo procesar.
  if (esXlsx && bytesArchivo && !soloCeldas) {
    try {
      const zip = await JSZip.loadAsync(bytesArchivo);
      const mapaHojas = await mapaHojaARutaXml(zip);
      const rutaHojaXml = mapaHojas.get(nombreHoja);
      const { graficos } = rutaHojaXml ? await objetosDeHoja(zip, rutaHojaXml) : { graficos: [] };

      function agregarPaginaConImagen(imagenPdf, leyenda) {
        const p = pdf.addPage([anchoPagina, altoPagina]);
        const anchoDisponible = anchoPagina - margen * 2;
        const altoDisponible = altoPagina - margen * 2 - 20;
        const escala = Math.min(anchoDisponible / imagenPdf.width, altoDisponible / imagenPdf.height, 1);
        const anchoFinal = imagenPdf.width * escala;
        const altoFinal = imagenPdf.height * escala;
        p.drawImage(imagenPdf, {
          x: margen + (anchoDisponible - anchoFinal) / 2,
          y: margen + (altoDisponible - altoFinal) / 2 + 20,
          width: anchoFinal,
          height: altoFinal,
        });
        const textoLeyenda = truncarTexto(fuente, leyenda, 9, anchoDisponible, fuenteUnicode, metadatosFuenteUnicode);
        const lienzoLeyenda = crearLienzoPdf(pdf, p);
        dibujarTextoPdf(lienzoLeyenda, textoLeyenda, { x: margen, y: margen, size: 9, color: rgb(0.4, 0.4, 0.4) });
        cerrarLienzoPdf(lienzoLeyenda);
      }

      for (const grafico of graficos) {
        try {
          const datos = parsearGrafico(grafico.doc);
          if (!datos) continue;
          const bytesGrafico = await renderizarGraficoComoPng(datos);
          const imagenPdf = await pdf.embedPng(bytesGrafico);
          const leyendaTipo = datos.soportado ? '' : t('leyenda-tipo-no-soportado');
          const leyenda = t('leyenda-grafico')
            .replace('{titulo}', datos.titulo ? `: ${datos.titulo}` : '')
            .replace('{celda}', grafico.celda || '?')
            .replace('{hoja}', nombreHoja)
            .replace('{tipo}', leyendaTipo);
          agregarPaginaConImagen(imagenPdf, leyenda);
        } catch (errorGrafico) {
          console.warn('No se pudo reconstruir un gráfico de la hoja:', errorGrafico);
        }
      }
    } catch (errorObjetos) {
      console.warn('No se pudieron leer los gráficos del archivo:', errorObjetos);
    }
  }

  if (caracteresPdfAdaptados > 0) {
    mensajesAdvertencia.push(t('aviso-caracteres').replace('{n}', caracteresPdfAdaptados));
  }
  if (alGuardar) alGuardar();
  return { bytes: await pdf.save(), avisos: mensajesAdvertencia };
}

// Un libro entero → un PDF, hoja tras hoja, con las opciones de fábrica
// (lo que usa "Unir PDF"). Las hojas vacías se saltan; si no hay ninguna con
// datos, error. Devuelve los bytes del PDF. alAvanzar(fraccion): de 0 a 1,
// por hojas y filas.
async function libroExcelAPdf(archivo, alAvanzar) {
  const bytesArchivo = await archivo.arrayBuffer();
  const libro = await leerLibroDesdeArchivo(archivo, { cellStyles: true });
  const esXlsx = /\.(xlsx|xlsm)$/i.test(archivo.name);
  if (esXlsx) {
    try { await leerFuentesDeCeldas(libro, bytesArchivo); } catch (error) { console.warn('No se pudieron leer los estilos de letra:', error); }
  }
  const { PDFDocument } = PDFLib;
  const unido = await PDFDocument.create();
  for (const [indiceHoja, nombreHoja] of libro.SheetNames.entries()) {
    const hoja = libro.Sheets[nombreHoja];
    if (!hoja || !hoja['!ref']) continue;
    let bytes;
    try {
      ({ bytes } = await construirPdfDeHoja({
        libro, nombreHoja, bytesArchivo, esXlsx,
        alAvanzar: alAvanzar ? (n, total) => alAvanzar((indiceHoja + n / total) / libro.SheetNames.length) : null,
      }));
    } catch (error) {
      if (error.sinDatos) continue;
      throw error;
    }
    const origen = await PDFDocument.load(bytes);
    (await unido.copyPages(origen, origen.getPageIndices())).forEach((p) => unido.addPage(p));
  }
  if (unido.getPageCount() === 0) throw new Error(t('error-sin-datos'));
  return unido.save();
}

/*! © 2026 Herramientas Ofimáticas (Mini Cookies). Código propio, todos los derechos reservados. Uso sujeto a los términos: https://www.herramientasofimaticas.minicookies.net/menus/terminos.html */
// Helpers compartidos de "texto plano -> PDF", "texto de un PDF con sus
// renglones y párrafos" y "extraer texto de un .pptx".
// Requiere que la página ya haya cargado pdf-lib (global PDFLib) y, para
// extraerTextoPowerPoint, JSZip (global JSZip).

// Las fuentes estándar de PDF (Helvetica) solo traen los caracteres de
// Windows-1252: un solo emoji, una flecha "→" o un ideograma hacían fallar
// toda la conversión ("WinAnsi cannot encode..."). Se reemplazan los más
// comunes por un equivalente y el resto por "?", en vez de fallar.
const EQUIVALENTES_WINANSI = {
  '→': '->', '←': '<-', '↑': '^', '↓': 'v', '⇒': '=>', '≤': '<=', '≥': '>=', '≠': '!=',
  '✓': 'v', '✔': 'v', '✗': 'x', '✘': 'x', '●': '•', '◦': '•', '▪': '•', '■': '•', '\t': '    ',
};
function adaptarTextoAFuente(texto, fuente) {
  const permitidos = new Set(fuente.getCharacterSet());
  return [...texto.normalize('NFC')].map((c) => {
    if (permitidos.has(c.codePointAt(0))) return c;
    if (EQUIVALENTES_WINANSI[c] !== undefined) return EQUIVALENTES_WINANSI[c];
    if (/[\u200B-\u200D\uFE0F\u20E3]/.test(c)) return ''; // caracteres invisibles de unión/emoji
    return '?';
  }).join('');
}

async function generarPdfDeTexto(texto) {
  const { PDFDocument, StandardFonts, rgb } = PDFLib;
  const pdf = await PDFDocument.create();
  const fuente = await pdf.embedFont(StandardFonts.Helvetica);
  texto = adaptarTextoAFuente(texto.replace(/\r\n?/g, '\n'), fuente);
  const tamanoFuente = 11;
  const margen = 50;
  const anchoPagina = 595, altoPagina = 842; // A4
  const anchoUtil = anchoPagina - margen * 2;
  const interlineado = tamanoFuente * 1.4;

  function envolverLinea(linea) {
    const palabras = linea.split(' ');
    const lineas = [];
    let actual = '';
    palabras.forEach((palabra) => {
      const prueba = actual ? actual + ' ' + palabra : palabra;
      if (actual && fuente.widthOfTextAtSize(prueba, tamanoFuente) > anchoUtil) {
        lineas.push(actual);
        actual = palabra;
      } else {
        actual = prueba;
      }
    });
    lineas.push(actual);
    return lineas;
  }

  const lineasFinales = [];
  texto.split('\n').forEach((linea) => envolverLinea(linea).forEach(l => lineasFinales.push(l)));

  let pagina = pdf.addPage([anchoPagina, altoPagina]);
  let y = altoPagina - margen;

  lineasFinales.forEach((linea) => {
    if (y < margen) {
      pagina = pdf.addPage([anchoPagina, altoPagina]);
      y = altoPagina - margen;
    }
    pagina.drawText(linea, { x: margen, y, size: tamanoFuente, font: fuente, color: rgb(0.15, 0.15, 0.15) });
    y -= interlineado;
  });

  return pdf;
}

// --- Texto de un PDF con sus renglones y párrafos ---
//
// getTextContent() de pdf.js devuelve fragmentos sueltos con su posición en la
// página. Unirlos todos con un espacio (como se hacía antes) dejaba cada
// página como un único renglón gigante: se perdían todos los saltos de línea
// y de párrafo. Acá se reconstruyen a partir de las coordenadas: cambio de
// altura de la línea base = renglón nuevo; un salto vertical bastante mayor al
// interlineado habitual de la página = párrafo nuevo.
// También corta el párrafo cuando cambia claramente el tamaño de letra (un
// título seguido de texto normal quedaba pegado en un solo párrafo).
// extraerRenglonesPaginaPdf devuelve [{ texto, y, alto, parrafoNuevo }] (alto =
// tamaño de letra aproximado); extraerTextoPaginaPdf, el texto con los
// renglones separados por "\n" y los párrafos por una línea en blanco ("\n\n").
async function extraerRenglonesPaginaPdf(pagina) {
  const contenido = await pagina.getTextContent();
  const lineas = [];
  let linea = null;
  let finAnterior = null;
  let saltoPendiente = false;
  for (const item of contenido.items) {
    if (typeof item.str !== 'string') continue; // marcas de contenido, sin texto
    const x = item.transform[4];
    const y = item.transform[5];
    const alto = Math.abs(item.height || item.transform[3]) || 10;
    const renglonNuevo = !linea || saltoPendiente || Math.abs(y - linea.y) > alto * 0.5;
    if (renglonNuevo) {
      if (item.str.trim() === '' && !item.hasEOL) { saltoPendiente = false; continue; }
      linea = { texto: '', y, alto };
      lineas.push(linea);
      finAnterior = null;
    } else if (finAnterior !== null && x - finAnterior > alto * 0.15 && !/\s$/.test(linea.texto) && !/^\s/.test(item.str)) {
      // Dos fragmentos del mismo renglón separados por un hueco: va un espacio.
      linea.texto += ' ';
    }
    linea.texto += item.str;
    finAnterior = x + (item.width || 0);
    saltoPendiente = !!item.hasEOL;
  }

  const renglones = lineas.map((l) => ({ ...l, texto: l.texto.replace(/\s+$/, '') })).filter((l) => l.texto !== '');
  if (renglones.length === 0) return [];

  // Interlineado habitual = mediana de la distancia entre renglones seguidos.
  const saltos = [];
  for (let i = 1; i < renglones.length; i++) {
    const salto = renglones[i - 1].y - renglones[i].y;
    if (salto > 0) saltos.push(salto);
  }
  saltos.sort((a, b) => a - b);
  const interlineado = saltos.length ? saltos[Math.floor(saltos.length / 2)] : renglones[0].alto * 1.2;

  renglones[0].parrafoNuevo = true;
  for (let i = 1; i < renglones.length; i++) {
    const salto = renglones[i - 1].y - renglones[i].y;
    const anterior = renglones[i - 1].alto, actual = renglones[i].alto;
    // Un salto mucho mayor que el habitual, hacia arriba (otra columna), o un
    // cambio claro de tamaño de letra (título ↔ texto) separa párrafos.
    renglones[i].parrafoNuevo = salto > interlineado * 1.45 || salto < 0 || Math.max(anterior, actual) / Math.min(anterior, actual) > 1.25;
  }
  return renglones;
}

async function extraerTextoPaginaPdf(pagina) {
  const renglones = await extraerRenglonesPaginaPdf(pagina);
  return renglones.map((r, i) => (i === 0 ? '' : r.parrafoNuevo ? '\n\n' : '\n') + r.texto).join('');
}

// Párrafos de una página con su tamaño de letra (el mayor de sus renglones):
// [{ texto, alto }], con el texto de cada párrafo en una sola línea.
async function extraerParrafosPaginaPdf(pagina) {
  const parrafos = [];
  for (const r of await extraerRenglonesPaginaPdf(pagina)) {
    if (r.parrafoNuevo || !parrafos.length) parrafos.push({ renglones: [], alto: 0 });
    const actual = parrafos[parrafos.length - 1];
    actual.renglones.push(r.texto);
    actual.alto = Math.max(actual.alto, r.alto);
  }
  return parrafos.map((p) => ({ texto: unirRenglonesDeParrafo(p.renglones.join('\n')), alto: p.alto })).filter((p) => p.texto);
}

// Une los renglones de un párrafo en una sola línea de texto corrido (para
// Word, EPUB o Markdown, donde el ancho lo decide quien lo muestra), juntando
// además las palabras cortadas con guion al final de un renglón ("docu-\nmento").
function unirRenglonesDeParrafo(parrafo) {
  return parrafo.split('\n').reduce((acumulado, renglon) => {
    if (!acumulado) return renglon.trim();
    if (/\p{L}-$/u.test(acumulado) && /^\p{Ll}/u.test(renglon.trim())) return acumulado.slice(0, -1) + renglon.trim();
    return acumulado + ' ' + renglon.trim();
  }, '');
}

async function extraerTextoPowerPoint(archivo) {
  const zip = await JSZip.loadAsync(archivo);
  const nombresDiapositivas = Object.keys(zip.files)
    .filter(nombre => /^ppt\/slides\/slide\d+\.xml$/.test(nombre))
    .sort((a, b) => parseInt(a.match(/slide(\d+)/)[1], 10) - parseInt(b.match(/slide(\d+)/)[1], 10));

  if (nombresDiapositivas.length === 0) {
    throw new Error(window.t ? t('ppt-sin-diapositivas') : 'No se encontraron diapositivas legibles en este archivo.');
  }

  const bloques = [];
  for (let i = 0; i < nombresDiapositivas.length; i++) {
    const xmlTexto = await zip.files[nombresDiapositivas[i]].async('string');
    const doc = new DOMParser().parseFromString(xmlTexto, 'application/xml');
    const parrafos = [...doc.getElementsByTagName('a:p')];
    const lineas = parrafos
      .map(p => [...p.getElementsByTagName('a:t')].map(t => t.textContent).join(''))
      .filter(l => l.trim() !== '');
    const tituloDiapositiva = (window.t ? t('ppt-diapositiva-titulo') : '--- Diapositiva {n} ---').replace('{n}', i + 1);
    const sinTexto = window.t ? t('ppt-sin-texto') : '(sin texto)';
    bloques.push(`${tituloDiapositiva}\n` + (lineas.length > 0 ? lineas.join('\n') : sinTexto));
  }
  return bloques.join('\n\n');
}

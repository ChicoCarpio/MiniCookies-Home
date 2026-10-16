/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Helpers compartidos de "texto plano -> PDF", "texto de un PDF con sus
// renglones y párrafos" y "extraer texto de un .pptx".
// Requiere que la página ya haya cargado pdf-lib (global PDFLib), tablas-texto.js
// (para el texto de un PDF con tablas) y, para extraerTextoPowerPoint, ZipPropio.

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
    // El salto de línea no es un carácter de la fuente, pero quien llama parte
    // el texto por "\n" después de adaptarlo: como "?", Texto → PDF salía en
    // un solo renglón (encontrado el 2026-10-09).
    if (c === '\n' || permitidos.has(c.codePointAt(0))) return c;
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

// Markdown sencillo → PDF. Títulos (#, ## y ###) en negrita y más grandes,
// listas con viñeta y párrafos. El formato en línea (negrita, cursiva, código,
// enlaces e imágenes) se quita y queda el texto. No es un renderizado completo
// de Markdown: las tablas y los bloques de código salen como texto plano.
function envolverTexto(texto, fuente, tamano, ancho) {
  const lineas = [];
  let actual = '';
  texto.split(' ').forEach((palabra) => {
    const prueba = actual ? actual + ' ' + palabra : palabra;
    if (actual && fuente.widthOfTextAtSize(prueba, tamano) > ancho) {
      lineas.push(actual);
      actual = palabra;
    } else {
      actual = prueba;
    }
  });
  lineas.push(actual);
  return lineas;
}

function quitarFormatoMarkdown(linea) {
  return linea
    .replace(/!\[([^\]]*)\]\([^)]*\)/g, '$1')
    .replace(/\[([^\]]+)\]\([^)]*\)/g, '$1')
    .replace(/(\*\*|__)(.+?)\1/g, '$2')
    .replace(/(\*|_)(.+?)\1/g, '$2')
    .replace(/`([^`]+)`/g, '$1');
}

async function generarPdfDeMarkdown(markdown) {
  const { PDFDocument, StandardFonts, rgb } = PDFLib;
  const pdf = await PDFDocument.create();
  const normal = await pdf.embedFont(StandardFonts.Helvetica);
  const negrita = await pdf.embedFont(StandardFonts.HelveticaBold);
  const margen = 50;
  const anchoPagina = 595, altoPagina = 842; // A4
  const anchoUtil = anchoPagina - margen * 2;
  let pagina = pdf.addPage([anchoPagina, altoPagina]);
  let y = altoPagina - margen;

  function escribir(texto, { fuente = normal, tamano = 11, sangria = 0, espacioAntes = 0 } = {}) {
    const lineas = envolverTexto(adaptarTextoAFuente(texto, fuente), fuente, tamano, anchoUtil - sangria);
    y -= espacioAntes;
    lineas.forEach((linea) => {
      if (y < margen + tamano) {
        pagina = pdf.addPage([anchoPagina, altoPagina]);
        y = altoPagina - margen;
      }
      pagina.drawText(linea, { x: margen + sangria, y, size: tamano, font: fuente, color: rgb(0.15, 0.15, 0.15) });
      y -= tamano * 1.4;
    });
  }

  let parrafo = [];
  const cerrarParrafo = () => {
    if (parrafo.length) escribir(parrafo.join(' '));
    parrafo = [];
  };
  markdown.replace(/\r\n?/g, '\n').split('\n').forEach((crudo) => {
    const linea = crudo.trim();
    const titulo = linea.match(/^(#{1,3})\s+(.*)$/);
    const item = linea.match(/^[-*+]\s+(.*)$/);
    if (!linea) { cerrarParrafo(); return; }
    if (titulo) {
      cerrarParrafo();
      const tamano = { 1: 20, 2: 16, 3: 13 }[titulo[1].length];
      escribir(quitarFormatoMarkdown(titulo[2]), { fuente: negrita, tamano, espacioAntes: 8 });
      y -= 4;
      return;
    }
    if (item) {
      cerrarParrafo();
      escribir('• ' + quitarFormatoMarkdown(item[1]), { sangria: 14 });
      return;
    }
    parrafo.push(quitarFormatoMarkdown(linea));
  });
  cerrarParrafo();
  return pdf;
}

// --- Texto de un PDF con sus renglones, párrafos y tablas ---
//
// getTextContent() de pdf.js devuelve fragmentos sueltos con su posición en la
// página. Unirlos todos con un espacio (como se hacía antes) dejaba cada
// página como un único renglón gigante: se perdían todos los saltos de línea
// y de párrafo. Acá se reconstruyen a partir de las coordenadas: cambio de
// altura de la línea base = renglón nuevo; un salto vertical bastante mayor al
// interlineado habitual de la página = párrafo nuevo.
// También corta el párrafo cuando cambia claramente el tamaño de letra (un
// título seguido de texto normal quedaba pegado en un solo párrafo).
//
// Tablas: cada renglón guarda además sus "celdas" (los fragmentos separados por
// un hueco horizontal). Varios renglones seguidos con celdas que caen en las
// mismas columnas son una tabla (buscarTablasPdf). Antes las celdas de una fila
// salían pegadas con un espacio y las filas como renglones sueltos.
//
// leerBloquesPaginaPdf devuelve, en orden de lectura, renglones
// { tipo: 'renglon', texto, y, alto, parrafoNuevo } (alto = tamaño de letra
// aproximado) y tablas { tipo: 'tabla', filas, alto }.
const PDF_HUECO_CELDA = 0.55; // hueco entre fragmentos (en alturas de letra) que parte una celda
const PDF_MARCA_DE_LISTA = /^([•·▪◦●○■□✓✔–—*+-]|\(?[0-9a-zA-Z]{1,3}[.)])$/;

async function leerLineasPaginaPdf(pagina) {
  const contenido = await pagina.getTextContent();
  const lineas = [];
  let linea = null;
  let saltoPendiente = false;
  for (const item of contenido.items) {
    if (typeof item.str !== 'string') continue; // marcas de contenido, sin texto
    const x = item.transform[4];
    const y = item.transform[5];
    const alto = Math.abs(item.height || item.transform[3]) || 10;
    const vacio = item.str.trim() === '';
    const renglonNuevo = !linea || saltoPendiente || Math.abs(y - linea.y) > alto * 0.5;
    if (renglonNuevo) {
      if (vacio && !item.hasEOL) { saltoPendiente = false; continue; }
      linea = { texto: '', y, alto, celdas: [], vecinas: [], fin: null, finVisible: null, derecha: -Infinity };
      lineas.push(linea);
    } else if (linea.fin !== null && x - linea.fin > alto * 0.15 && !/\s$/.test(linea.texto) && !/^\s/.test(item.str)) {
      // Dos fragmentos del mismo renglón separados por un hueco: va un espacio.
      linea.texto += ' ';
    }
    linea.texto += item.str;
    linea.fin = x + (item.width || 0);
    if (!vacio) {
      // El ancho de un fragmento incluye los espacios de los extremos: no cuentan como parte de la celda.
      const largo = item.str.length, sinFinal = item.str.replace(/\s+$/, '').length;
      const x1 = x + (item.width || 0) * (sinFinal / largo);
      const x0 = x + (item.width || 0) * ((largo - item.str.replace(/^\s+/, '').length) / largo);
      const texto = item.str.trim();
      const celda = linea.celdas[linea.celdas.length - 1];
      let propia, nueva = false;
      if (celda && linea.finVisible !== null && x0 - linea.finVisible <= alto * PDF_HUECO_CELDA) {
        celda.texto += (x0 - linea.finVisible > alto * 0.15 ? ' ' : '') + texto;
        celda.x1 = x1;
        propia = celda;
      } else {
        propia = { x0, x1, texto, absorbida: false };
        linea.celdas.push(propia);
        nueva = true;
      }
      linea.finVisible = x1;

      // Una celda de varios renglones se dibuja entera antes de la de al lado:
      // el texto de la celda siguiente vuelve a la altura de la fila (de un
      // renglón con dos celdas o más de unos pocos renglones atrás), a la
      // derecha de lo que ya tenía y con solo la continuación de sus celdas de
      // por medio. Es de ese renglón si hay una tabla; se anota aparte
      // (celdasConVecinas) para no alterar el texto corrido, que sigue el orden del archivo.
      linea.derecha = x1;
      let absorbido = false;
      for (let k = lineas.length - 1; k >= Math.max(0, lineas.length - 13); k--) {
        const previa = lineas[k];
        if (previa === linea || previa.celdas.length < 2 || Math.abs(y - previa.y) > alto * 0.5 || x0 < previa.derecha - alto * 0.2) continue;
        // Entre aquel renglón y este solo puede haber lo que sigue de una de sus celdas (empieza donde empieza ella).
        const intermedios = lineas.slice(k + 1).filter((m) => m !== linea && m.celdas.length);
        if (!intermedios.every((m) => previa.celdas.some((c) => Math.abs(c.x0 - m.celdas[0].x0) <= alto * 0.3))) continue;
        const ultima = previa.vecinas[previa.vecinas.length - 1];
        if (ultima && x0 - ultima.x1 <= alto * PDF_HUECO_CELDA) {
          ultima.texto += (x0 - ultima.x1 > alto * 0.15 ? ' ' : '') + texto;
          ultima.x1 = x1;
        } else {
          previa.vecinas.push({ x0, x1, texto });
        }
        previa.derecha = x1;
        absorbido = true;
        break;
      }
      propia.absorbida = nueva ? absorbido : propia.absorbida && absorbido;
    }
    saltoPendiente = !!item.hasEOL;
  }
  return lineas
    .map((l) => ({
      ...l,
      texto: l.texto.replace(/\s+$/, ''),
      // Las celdas del renglón más las que le tocan por altura (las "vecinas"), en orden, para buscar tablas.
      celdasConVecinas: [...l.celdas.map((c) => ({ ...c })), ...l.vecinas].sort((p, q) => p.x0 - q.x0),
      soloVecinas: l.celdas.length > 0 && l.celdas.every((c) => c.absorbida),
    }))
    .filter((l) => l.texto !== '');
}

const medianaPdf = (numeros) => {
  const orden = [...numeros].sort((a, b) => a - b);
  return orden.length ? orden[Math.floor(orden.length / 2)] : 0;
};

// Busca las tablas entre los renglones. Devuelve [{ desde, hasta, filas, alto }]
// con los renglones desde..hasta-1 (hasta no incluido) que forman cada tabla.
function buscarTablasPdf(lineas) {
  const tablas = [];
  const cercano = (previa, l) => previa.y - l.y > 0 && previa.y - l.y <= Math.max(l.alto, previa.alto) * 3 &&
    Math.max(l.alto, previa.alto) / Math.min(l.alto, previa.alto) < 1.25;
  let i = 0;
  while (i < lineas.length) {
    if (lineas[i].celdas.length < 2) { i++; continue; }
    // Hasta donde llegan renglones que pueden ser parte de la tabla: con dos
    // celdas o más, o con una sola celda entre dos de esos (una fila de un solo
    // dato, o lo que sigue de una celda que ocupó varios renglones). Los de una
    // celda al final se incluyen también y armarTablaPdf decide si sobran.
    let fin = i + 1;
    let j = i + 1;
    while (j < lineas.length && cercano(lineas[j - 1], lineas[j])) {
      if (lineas[j].celdas.length >= 2) { fin = ++j; continue; }
      let k = j;
      while (k < lineas.length && k - j < 12 && lineas[k].celdas.length < 2 && cercano(lineas[k - 1], lineas[k])) k++;
      if (k < lineas.length && lineas[k].celdas.length >= 2 && cercano(lineas[k - 1], lineas[k])) { fin = k + 1; j = k + 1; continue; }
      fin = k; // renglones de una celda al final
      break;
    }
    const tabla = armarTablaPdf(lineas, i, fin);
    if (tabla) { tablas.push(tabla); i = tabla.hasta; } else i++;
  }
  return tablas;
}

function armarTablaPdf(lineas, desde, hasta) {
  const run = lineas.slice(desde, hasta);
  // Número de celdas más frecuente entre las filas con dos o más.
  const cuentas = new Map();
  run.forEach((l) => { if (l.celdas.length >= 2) cuentas.set(l.celdas.length, (cuentas.get(l.celdas.length) || 0) + 1); });
  let moda = 0, veces = 0;
  for (const [n, v] of cuentas) if (v > veces || (v === veces && n > moda)) { moda = n; veces = v; }
  if (veces < 2) return null;
  const regulares = run.filter((l) => l.celdas.length === moda);

  // Columna c = el rango que ocupan, en todas las filas regulares, sus celdas número c.
  const columnas = Array.from({ length: moda }, (_, c) => ({
    lo: Math.min(...regulares.map((l) => l.celdas[c].x0)),
    hi: Math.max(...regulares.map((l) => l.celdas[c].x1)),
  }));
  const altoMedio = medianaPdf(run.map((l) => l.alto));
  for (let c = 1; c < moda; c++) if (columnas[c].lo < columnas[c - 1].hi + altoMedio * 0.2) return null; // columnas que se pisan: no es una tabla

  // Una lista con viñetas o números ("1.   Texto") no es una tabla de dos columnas.
  const primeras = regulares.map((l) => l.celdas[0].texto);
  if (primeras.every((t) => PDF_MARCA_DE_LISTA.test(t))) return null;

  const columnaDe = (celda) => {
    let mejor = 0, solape = -Infinity;
    columnas.forEach((col, c) => {
      const s = Math.min(col.hi, celda.x1) - Math.max(col.lo, celda.x0);
      if (s > solape) { solape = s; mejor = c; }
    });
    if (solape <= 0) { // en ninguna: la más cercana por el centro
      const centro = (celda.x0 + celda.x1) / 2;
      let d = Infinity;
      columnas.forEach((col, c) => { const dc = Math.abs((col.lo + col.hi) / 2 - centro); if (dc < d) { d = dc; mejor = c; } });
    }
    return mejor;
  };

  // Separación habitual entre filas, para saber si un renglón de menos celdas
  // es la continuación de la celda de arriba (el texto de una celda que ocupó
  // dos renglones) o una fila nueva.
  const pasos = [];
  for (let k = 1; k < run.length; k++) {
    if (run[k].celdas.length === moda && run[k - 1].celdas.length === moda) pasos.push(run[k - 1].y - run[k].y);
  }
  const pasoFila = medianaPdf(pasos);

  // Columnas de prosa (como en un periódico): renglones tan juntos como los de un
  // párrafo y casi todas las celdas con una frase larga. No es una tabla.
  if (regulares.length >= 4 && pasoFila > 0 && pasoFila <= altoMedio * 1.6) {
    const celdas = regulares.flatMap((l) => l.celdas);
    if (celdas.filter((c) => c.texto.length >= 25).length >= celdas.length * 0.8) return null;
  }

  const filas = [];
  let ultimo = null; // renglón anterior de la tabla
  let hastaReal = desde;
  // Continuación = lo que sigue de una celda: pegado al renglón de arriba y empezando donde empieza una de sus celdas.
  // Con pocas filas seguidas no hay separación medida: menos de ~1,45 veces la letra.
  const esContinuacion = (l, salto) => filas.length > 0 && salto > 0 && (pasoFila > 0 ? salto < pasoFila * 0.85 : salto <= l.alto * 1.45) &&
    l.celdas.every((c) => ultimo.celdas.some((u) => Math.abs(u.x0 - c.x0) <= l.alto * 0.3));
  for (let k = 0; k < run.length; k++) {
    const l = run[k];
    const regular = l.celdas.length === moda;
    const salto = ultimo ? ultimo.y - l.y : Infinity;
    const continuacion = !regular && esContinuacion(l, salto);
    if (!regular && !continuacion && l.celdas.length < 2 && !run.slice(k + 1).some((x) => x.celdas.length >= 2)) break; // renglones sueltos al final: no son de la tabla
    // Una fila con menos celdas que las demás tiene que caer cada una en una columna distinta ("•   texto" no).
    if (!regular && !continuacion && l.celdas.length >= 2) {
      const columnasDeLaFila = l.celdas.map(columnaDe);
      if (columnasDeLaFila.some((c, i) => i > 0 && c <= columnasDeLaFila[i - 1])) break;
    }
    if (continuacion) {
      const fila = filas[filas.length - 1];
      l.celdas.forEach((celda) => { const c = columnaDe(celda); fila[c] = (fila[c] ? fila[c] + ' ' : '') + celda.texto; });
    } else {
      const nueva = Array(moda).fill('');
      l.celdas.forEach((celda) => { const c = columnaDe(celda); nueva[c] = (nueva[c] ? nueva[c] + ' ' : '') + celda.texto; });
      filas.push(nueva);
    }
    ultimo = l;
    hastaReal = desde + k + 1;
  }
  if (filas.length < 2) return null;
  return { desde, hasta: hastaReal, filas, alto: altoMedio };
}

async function leerBloquesPaginaPdf(pagina) {
  const lineas = await leerLineasPaginaPdf(pagina);
  if (lineas.length === 0) return [];

  // Las tablas se buscan en una vista donde las celdas que se dibujaron después
  // pero a la altura de un renglón anterior ya están en ese renglón, y los
  // renglones que solo eran eso no están. El texto corrido sigue el orden del archivo.
  const vista = lineas.map((l, indice) => ({ y: l.y, alto: l.alto, celdas: l.celdasConVecinas, indice, saltar: l.soloVecinas })).filter((l) => !l.saltar);

  // Renglones sueltos y tablas, en orden.
  const bloques = [];
  let desde = 0;
  for (const tabla of buscarTablasPdf(vista)) {
    const primero = vista[tabla.desde].indice;
    let ultimo = vista[tabla.hasta - 1].indice + 1;
    while (ultimo < lineas.length && lineas[ultimo].soloVecinas) ultimo++; // lo que cerraba la última fila
    lineas.slice(desde, primero).forEach((l) => bloques.push({ tipo: 'renglon', texto: l.texto, y: l.y, alto: l.alto }));
    bloques.push({ tipo: 'tabla', filas: tabla.filas, alto: tabla.alto });
    desde = ultimo;
  }
  lineas.slice(desde).forEach((l) => bloques.push({ tipo: 'renglon', texto: l.texto, y: l.y, alto: l.alto }));

  // Interlineado habitual = mediana de la distancia entre renglones seguidos (sin tablas de por medio).
  const saltos = [];
  for (let i = 1; i < bloques.length; i++) {
    if (bloques[i].tipo !== 'renglon' || bloques[i - 1].tipo !== 'renglon') continue;
    const salto = bloques[i - 1].y - bloques[i].y;
    if (salto > 0) saltos.push(salto);
  }
  const primero = bloques.find((b) => b.tipo === 'renglon');
  const interlineado = saltos.length ? medianaPdf(saltos) : (primero ? primero.alto * 1.2 : 12);

  bloques.forEach((b, i) => {
    if (b.tipo !== 'renglon') return;
    const previo = bloques[i - 1];
    if (!previo || previo.tipo !== 'renglon') { b.parrafoNuevo = true; return; }
    const salto = previo.y - b.y;
    // Un salto mucho mayor que el habitual, hacia arriba (otra columna), o un
    // cambio claro de tamaño de letra (título ↔ texto) separa párrafos.
    b.parrafoNuevo = salto > interlineado * 1.45 || salto < 0 || Math.max(previo.alto, b.alto) / Math.min(previo.alto, b.alto) > 1.25;
  });
  return bloques;
}

// Los renglones de una página (sin las tablas), como antes.
async function extraerRenglonesPaginaPdf(pagina) {
  return (await leerBloquesPaginaPdf(pagina)).filter((b) => b.tipo === 'renglon');
}

// El texto de la página: renglones separados por "\n", párrafos por una línea
// en blanco ("\n\n") y las tablas con sus columnas alineadas.
async function extraerTextoPaginaPdf(pagina) {
  const bloques = await leerBloquesPaginaPdf(pagina);
  return bloques.map((b, i) => {
    if (b.tipo === 'tabla') return (i === 0 ? '' : '\n\n') + TablasTexto.aTexto(b.filas);
    const separador = i === 0 ? '' : (b.parrafoNuevo ? '\n\n' : '\n');
    return separador + b.texto;
  }).join('');
}

// Párrafos de una página con su tamaño de letra (el mayor de sus renglones):
// [{ texto, alto }] con el texto de cada párrafo en una sola línea, y las
// tablas como { tabla: filas, alto }. Un renglón que empieza con viñeta o con
// número ("•", "-", "1.", "2)") abre un párrafo propio, con el elemento de
// lista en `item: { marca, resto }` (marca "-" para una viñeta o el número con su punto).
const PDF_ELEMENTO_DE_LISTA = /^([•·▪◦●○■□–—*-]|\d{1,3}[.)])\s+(\S.*)$/;
async function extraerParrafosPaginaPdf(pagina) {
  const resultado = [];
  let parrafo = null;
  for (const b of await leerBloquesPaginaPdf(pagina)) {
    if (b.tipo === 'tabla') { parrafo = null; resultado.push({ tabla: b.filas, alto: b.alto }); continue; }
    const elemento = PDF_ELEMENTO_DE_LISTA.exec(b.texto);
    if (b.parrafoNuevo || elemento || !parrafo) { parrafo = { renglones: [], alto: 0, elemento }; resultado.push(parrafo); }
    parrafo.renglones.push(b.texto);
    parrafo.alto = Math.max(parrafo.alto, b.alto);
  }
  return resultado
    .map((p) => {
      if (p.tabla) return p;
      const texto = unirRenglonesDeParrafo(p.renglones.join('\n'));
      const item = p.elemento ? { marca: /^\d/.test(p.elemento[1]) ? p.elemento[1] : '-', resto: texto.slice(p.elemento[1].length).trim() } : null;
      return item ? { texto, alto: p.alto, item } : { texto, alto: p.alto };
    })
    .filter((p) => p.tabla || p.texto);
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
  const zip = await ZipPropio.loadAsync(archivo);
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

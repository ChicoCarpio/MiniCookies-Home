// Convierte HTML "de verdad" (de un <div contenteditable> o de mammoth.convertToHtml)
// a los objetos Paragraph/Table/TextRun/ImageRun de docx.js. Requiere que la página
// ya haya cargado docx.js (global `docx`) antes de usar contenidoAParrafos.
//
// Es todo asíncrono porque medir el ancho/alto real de una imagen incrustada
// requiere esperar a que el navegador la decodifique.

function leerFormatoDeAncestros(nodo, raiz) {
  const formato = { bold: false, italics: false, underline: false };
  let actual = nodo.parentElement;
  while (actual && actual !== raiz) {
    const etiqueta = actual.tagName;
    if (etiqueta === 'B' || etiqueta === 'STRONG') formato.bold = true;
    if (etiqueta === 'I' || etiqueta === 'EM') formato.italics = true;
    if (etiqueta === 'U') formato.underline = true;
    actual = actual.parentElement;
  }
  return formato;
}

// docx.js acepta el string base64 tal cual para ImageRun.data, sin decodificarlo
// a bytes — alcanza con separar el tipo de imagen y el contenido del data: URI.
function decodificarDataUriImagen(src) {
  const coincidencia = /^data:image\/(png|jpe?g|gif|bmp);base64,(.+)$/i.exec(src || '');
  if (!coincidencia) return null;
  let tipo = coincidencia[1].toLowerCase();
  if (tipo === 'jpeg') tipo = 'jpg';
  return { tipo, base64: coincidencia[2] };
}

function medirImagen(src) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth || 300, height: img.naturalHeight || 200 });
    img.onerror = () => resolve({ width: 300, height: 200 }); // valor de respaldo si no se puede medir
    img.src = src;
  });
}

const ANCHO_MAXIMO_IMAGEN_DOCX = 550; // aprox. el ancho útil de una página A4 con márgenes normales

async function crearImageRunDesdeImg(imgEl) {
  const src = imgEl.getAttribute('src') || '';
  const decodificado = decodificarDataUriImagen(src);
  if (!decodificado) return null; // src externo o formato no soportado (ej. SVG) — se omite sin romper el resto

  const medidas = await medirImagen(src);
  let ancho = medidas.width;
  let alto = medidas.height;
  if (ancho > ANCHO_MAXIMO_IMAGEN_DOCX) {
    alto = Math.round(alto * (ANCHO_MAXIMO_IMAGEN_DOCX / ancho));
    ancho = ANCHO_MAXIMO_IMAGEN_DOCX;
  }

  return new docx.ImageRun({
    data: decodificado.base64,
    type: decodificado.tipo,
    transformation: { width: ancho, height: alto },
  });
}

// omitirListas: en un <li>, las sublistas (<ul>/<ol> anidadas) se
// convierten aparte, como párrafos de nivel siguiente; si no se omiten, su
// texto quedaba pegado al del ítem padre.
async function extraerRuns(elementoBloque, omitirListas) {
  const runs = [];
  async function recorrer(nodo) {
    if (nodo.nodeType === Node.TEXT_NODE) {
      if (nodo.textContent === '') return;
      const formato = leerFormatoDeAncestros(nodo, elementoBloque);
      runs.push(new docx.TextRun({ text: nodo.textContent, ...formato }));
      return;
    }
    if (nodo.nodeType === Node.ELEMENT_NODE) {
      if (nodo.tagName === 'BR') {
        runs.push(new docx.TextRun({ text: '', break: 1 }));
        return;
      }
      if (nodo.tagName === 'IMG') {
        const imageRun = await crearImageRunDesdeImg(nodo);
        if (imageRun) runs.push(imageRun);
        return;
      }
      if (omitirListas && (nodo.tagName === 'UL' || nodo.tagName === 'OL')) return;
      for (const hijo of nodo.childNodes) {
        await recorrer(hijo);
      }
    }
  }
  for (const hijo of elementoBloque.childNodes) {
    await recorrer(hijo);
  }
  if (runs.length === 0) runs.push(new docx.TextRun(''));
  return runs;
}

const NIVELES_TITULO_DOCX = {
  H1: 'HEADING_1', H2: 'HEADING_2', H3: 'HEADING_3',
  H4: 'HEADING_4', H5: 'HEADING_5', H6: 'HEADING_6',
};

// Cada <ol> es una "instancia" distinta de la numeración: sin esto, todas
// las listas numeradas del documento compartían el contador y la segunda
// lista seguía en 4, 5, 6... en vez de empezar en 1.
let instanciaListaDocx = 0;

async function convertirListaAParrafos(listaEl, numerada, nivel) {
  nivel = nivel || 0;
  const instancia = ++instanciaListaDocx;
  const nivelDocx = Math.min(nivel, 8); // Word admite hasta 9 niveles (0-8)
  const parrafos = [];
  for (const item of [...listaEl.children]) {
    if (item.tagName !== 'LI') continue;
    parrafos.push(new docx.Paragraph({
      children: await extraerRuns(item, true),
      numbering: numerada ? { reference: 'lista-numerada', level: nivelDocx, instance: instancia } : undefined,
      bullet: numerada ? undefined : { level: nivelDocx },
    }));
    for (const sublista of [...item.children]) {
      if (sublista.tagName === 'UL' || sublista.tagName === 'OL') {
        parrafos.push(...await convertirListaAParrafos(sublista, sublista.tagName === 'OL', nivel + 1));
      }
    }
  }
  return parrafos;
}

async function convertirTablaADocx(tablaEl) {
  const filas = [];
  for (const filaEl of [...tablaEl.querySelectorAll('tr')]) {
    const celdas = [];
    for (const celdaEl of [...filaEl.children]) {
      celdas.push(new docx.TableCell({
        children: [new docx.Paragraph({ children: await extraerRuns(celdaEl) })],
      }));
    }
    filas.push(new docx.TableRow({ children: celdas }));
  }
  if (filas.length === 0) return null;
  return new docx.Table({ rows: filas, width: { size: 100, type: docx.WidthType.PERCENTAGE } });
}

// Recorre los hijos directos de "contenedor" (un <div contenteditable> o el <body>
// de un HTML parseado) y devuelve un array de Paragraph/Table de docx.js.
async function contenidoAParrafos(contenedor) {
  const bloques = [];
  for (const elemento of [...contenedor.children]) {
    if (elemento.tagName === 'UL') {
      bloques.push(...await convertirListaAParrafos(elemento, false));
    } else if (elemento.tagName === 'OL') {
      bloques.push(...await convertirListaAParrafos(elemento, true));
    } else if (elemento.tagName === 'TABLE') {
      const tabla = await convertirTablaADocx(elemento);
      if (tabla) bloques.push(tabla);
    } else if (NIVELES_TITULO_DOCX[elemento.tagName]) {
      bloques.push(new docx.Paragraph({
        heading: docx.HeadingLevel[NIVELES_TITULO_DOCX[elemento.tagName]],
        children: await extraerRuns(elemento),
      }));
    } else if (elemento.tagName === 'IMG') {
      const imageRun = await crearImageRunDesdeImg(elemento);
      if (imageRun) bloques.push(new docx.Paragraph({ children: [imageRun] }));
    } else {
      bloques.push(new docx.Paragraph({ children: await extraerRuns(elemento) }));
    }
  }
  return bloques;
}

// Config de numeración que hay que pasarle a `new docx.Document({ numbering: ... })`
// para que las listas numeradas (bullet: false) se vean como "1. 2. 3." reales.
// Los niveles 1+ son las sublistas (a., i., 1., ...), con sangría creciente.
const FORMATOS_NIVEL_DOCX = ['decimal', 'lowerLetter', 'lowerRoman'];
const NUMERACION_DOCX_CONFIG = {
  config: [{
    reference: 'lista-numerada',
    levels: Array.from({ length: 9 }, (_, nivel) => ({
      level: nivel,
      format: FORMATOS_NIVEL_DOCX[nivel % 3],
      text: `%${nivel + 1}.`,
      alignment: docx.AlignmentType.START,
      style: { paragraph: { indent: { left: 720 * (nivel + 1), hanging: 360 } } },
    })),
  }],
};

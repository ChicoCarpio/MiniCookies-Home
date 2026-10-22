/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Convierte HTML "de verdad" (de un <div contenteditable> o del lector de Word)
// a los objetos de párrafo, tabla y texto de escritor-docx.js (EscritorDocx). Se usa con
// EscritorDocx.generar({ numeraciones: NUMERACIONES_DOCX, hijos: await contenidoAParrafos(contenedor) }).
//
// Es todo asíncrono porque medir el ancho/alto real de una imagen incrustada
// requiere esperar a que el navegador la decodifique.

function leerFormatoDeAncestros(nodo, raiz) {
  const formato = { negrita: false, cursiva: false, subrayado: false };
  let actual = nodo.parentElement;
  while (actual && actual !== raiz) {
    const etiqueta = actual.tagName;
    if (etiqueta === 'B' || etiqueta === 'STRONG') formato.negrita = true;
    if (etiqueta === 'I' || etiqueta === 'EM') formato.cursiva = true;
    if (etiqueta === 'U') formato.subrayado = true;
    actual = actual.parentElement;
  }
  return formato;
}

// El escritor acepta la imagen en base64 tal cual: alcanza con separar el tipo
// de imagen y el contenido del data: URI.
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

  return { imagen: { datos: decodificado.base64, tipo: decodificado.tipo, ancho, alto } };
}

// El escritor entiende "\n" como salto de línea y "\t" como tabulación; en HTML el espacio en blanco
// se colapsa (el texto de un archivo HTML trae saltos de línea de sangría). Solo dentro de <pre> se conserva.
function textoComoHtml(nodo, raiz) {
  for (let actual = nodo.parentElement; actual && actual !== raiz; actual = actual.parentElement) {
    if (actual.tagName === 'PRE') return nodo.textContent;
  }
  return raiz.tagName === 'PRE' ? nodo.textContent : nodo.textContent.replace(/[ \t\r\n\f]+/g, ' '); // sin tocar los espacios duros (U+00A0)
}

// omitirListas: en un <li>, las sublistas (<ul>/<ol> anidadas) se
// convierten aparte, como párrafos de nivel siguiente; si no se omiten, su
// texto quedaba pegado al del ítem padre.
async function extraerRuns(elementoBloque, omitirListas, nodos) {
  const runs = [];
  async function recorrer(nodo) {
    if (nodo.nodeType === Node.TEXT_NODE) {
      if (nodo.textContent === '') return;
      const formato = leerFormatoDeAncestros(nodo, elementoBloque);
      runs.push({ texto: textoComoHtml(nodo, elementoBloque), ...formato });
      return;
    }
    if (nodo.nodeType === Node.ELEMENT_NODE) {
      if (nodo.tagName === 'BR') {
        runs.push({ texto: '\n' });
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
  for (const hijo of nodos || elementoBloque.childNodes) {
    await recorrer(hijo);
  }
  if (runs.length === 0) runs.push({ texto: '' });
  return runs;
}

const NIVELES_TITULO_DOCX = { H1: 1, H2: 2, H3: 3, H4: 4, H5: 5, H6: 6 };

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
    parrafos.push({
      tipo: 'parrafo',
      hijos: await extraerRuns(item, true),
      lista: numerada ? { id: 'lista-numerada', nivel: nivelDocx, instancia } : { vineta: true, nivel: nivelDocx },
    });
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
      celdas.push({ hijos: [{ tipo: 'parrafo', hijos: await extraerRuns(celdaEl) }] });
    }
    filas.push({ celdas });
  }
  if (filas.length === 0) return null;
  return { tipo: 'tabla', anchoPct: 100, filas };
}

// Un <div>, <p>, <section>, <blockquote>… que contiene listas, tablas, títulos u otros bloques no es un párrafo: se baja por él.
// (El navegador envuelve así las listas de un <div contenteditable>: <div><ul>…</ul><div><ol>…</ol></div></div>.)
// Solo se tiene por texto en línea lo que se sabe que lo es; cualquier otra etiqueta es un bloque.
const EN_LINEA_DOCX = new Set(['A', 'B', 'STRONG', 'I', 'EM', 'U', 'S', 'SPAN', 'BR', 'IMG', 'CODE', 'SMALL', 'SUB', 'SUP', 'MARK', 'FONT', 'ABBR', 'CITE', 'Q', 'TT', 'DEL', 'INS', 'KBD', 'VAR', 'WBR']);
const tieneBloques = (el) => [...el.children].some((h) => !EN_LINEA_DOCX.has(h.tagName));

// Recorre los hijos de "contenedor" (un <div contenteditable> o el <body> de un HTML parseado)
// y devuelve un array de párrafos y tablas para EscritorDocx. El texto suelto entre bloques
// (la primera línea de un editor, sin envoltorio) forma un párrafo.
async function contenidoAParrafos(contenedor) {
  const bloques = [];
  let suelto = [];
  const cerrarSuelto = async () => {
    const hayContenido = suelto.some((n) => (n.nodeType === Node.TEXT_NODE ? n.textContent.trim() !== '' : n.tagName === 'BR' || n.tagName === 'IMG' || n.textContent.trim() !== ''));
    if (hayContenido) bloques.push({ tipo: 'parrafo', hijos: await extraerRuns(contenedor, false, suelto) });
    suelto = [];
  };
  for (const nodo of [...contenedor.childNodes]) {
    if (nodo.nodeType !== Node.ELEMENT_NODE) { suelto.push(nodo); continue; }
    const elemento = nodo;
    const etiqueta = elemento.tagName;
    if (EN_LINEA_DOCX.has(etiqueta)) { suelto.push(elemento); continue; } // texto en línea (<b>, <span>, <br>…): se agrupa
    await cerrarSuelto();
    if (etiqueta === 'UL') {
      bloques.push(...await convertirListaAParrafos(elemento, false));
    } else if (etiqueta === 'OL') {
      bloques.push(...await convertirListaAParrafos(elemento, true));
    } else if (etiqueta === 'TABLE') {
      const tabla = await convertirTablaADocx(elemento);
      if (tabla) bloques.push(tabla);
    } else if (NIVELES_TITULO_DOCX[etiqueta]) {
      bloques.push({ tipo: 'parrafo', titulo: NIVELES_TITULO_DOCX[etiqueta], hijos: await extraerRuns(elemento) });
    } else if (etiqueta !== 'PRE' && tieneBloques(elemento)) {
      bloques.push(...await contenidoAParrafos(elemento));
    } else {
      bloques.push({ tipo: 'parrafo', hijos: await extraerRuns(elemento) });
    }
  }
  await cerrarSuelto();
  return bloques;
}

// Un salto de página entre documentos o capítulos (unir, dividir, administrar, EPUB).
const SALTO_DE_PAGINA_DOCX = () => ({ tipo: 'parrafo', hijos: [{ salto: 'pagina' }] });

// Numeración que hay que pasarle a EscritorDocx.generar({ numeraciones: NUMERACIONES_DOCX, … })
// para que las listas numeradas se vean como "1. 2. 3." reales.
// Los niveles 1+ son las sublistas (a., i., 1., ...), con sangría creciente.
const FORMATOS_NIVEL_DOCX = ['decimal', 'lowerLetter', 'lowerRoman'];
const NUMERACIONES_DOCX = [{
  id: 'lista-numerada',
  niveles: Array.from({ length: 9 }, (_, nivel) => ({
    formato: FORMATOS_NIVEL_DOCX[nivel % 3],
    texto: `%${nivel + 1}.`,
    izquierda: 720 * (nivel + 1),
    colgante: 360,
  })),
}];

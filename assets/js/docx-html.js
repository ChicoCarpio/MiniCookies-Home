/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
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

// Separa el tipo de imagen y el contenido (base64) del data: URI.
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
  let bytes;
  try { bytes = Uint8Array.from(atob(decodificado.base64), (c) => c.charCodeAt(0)); } catch (e) { return null; } // base64 roto: se omite la imagen, no todo el documento

  const medidas = await medirImagen(src);
  let ancho = medidas.width;
  let alto = medidas.height;
  if (ancho > ANCHO_MAXIMO_IMAGEN_DOCX) {
    alto = Math.round(alto * (ANCHO_MAXIMO_IMAGEN_DOCX / ancho));
    ancho = ANCHO_MAXIMO_IMAGEN_DOCX;
  }

  return { imagen: { datos: bytes, tipo: decodificado.tipo, ancho, alto } };
}

// El escritor entiende "\n" como salto de línea y "\t" como tabulación. El salto de línea de un archivo HTML
// (y el sangrado que lo sigue) es solo formato del archivo: se vuelve un espacio, y al principio o al final
// de un renglón no se ve. Los espacios y tabulaciones sin salto se conservan: los trae el texto de un Word
// leído con docx-lector (espacios seguidos, tabulaciones) o de un .txt (sangría). Dentro de <pre>, todo tal cual.
const SALTO_HTML = /[ \t]*[\r\n\f][ \t\r\n\f]*/g;
function dentroDePre(nodo, raiz) {
  for (let actual = nodo.parentElement; actual && actual !== raiz; actual = actual.parentElement) {
    if (actual.tagName === 'PRE') return true;
  }
  return raiz.tagName === 'PRE';
}

// omitirListas: en un <li>, las sublistas (<ul>/<ol> anidadas) se
// convierten aparte, como párrafos de nivel siguiente; si no se omiten, su
// texto quedaba pegado al del ítem padre.
async function extraerRuns(elementoBloque, omitirListas, nodos) {
  const runs = [];
  const originales = new Map(); // run de texto fuera de <pre> → su texto antes de colapsar los saltos
  const saltos = new Set(); // runs de un <br>
  async function recorrer(nodo) {
    if (nodo.nodeType === Node.TEXT_NODE) {
      if (nodo.textContent === '') return;
      const formato = leerFormatoDeAncestros(nodo, elementoBloque);
      const pre = dentroDePre(nodo, elementoBloque);
      const run = { texto: pre ? nodo.textContent : nodo.textContent.replace(SALTO_HTML, ' '), ...formato };
      if (!pre) originales.set(run, nodo.textContent);
      runs.push(run);
      return;
    }
    if (nodo.nodeType === Node.ELEMENT_NODE) {
      if (nodo.tagName === 'BR') {
        const run = { texto: '\n' };
        saltos.add(run);
        runs.push(run);
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
  // Como se ve el HTML: el salto de línea del archivo no se ve al principio de un renglón ni al final del bloque,
  // y un <br> final no agrega renglón (el <div><br></div> de una línea vacía del editor daba dos).
  if (saltos.has(runs[runs.length - 1])) runs.pop();
  let inicioDeRenglon = true;
  for (const run of runs) {
    if (saltos.has(run)) { inicioDeRenglon = true; continue; }
    if (!inicioDeRenglon) continue;
    if (originales.has(run) && /^[ \t]*[\r\n\f]/.test(originales.get(run))) run.texto = run.texto.slice(1);
    if (!originales.has(run) || run.texto) inicioDeRenglon = false;
  }
  for (let i = runs.length - 1; i >= 0 && originales.has(runs[i]); i--) {
    if (/[\r\n\f][ \t]*$/.test(originales.get(runs[i])) && runs[i].texto.endsWith(' ')) runs[i].texto = runs[i].texto.slice(0, -1);
    if (runs[i].texto) break;
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
  for (const filaEl of [...tablaEl.rows]) { // solo las filas de esta tabla: querySelectorAll('tr') repetía las de una tabla anidada
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

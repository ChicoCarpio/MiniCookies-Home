/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. */
// Revisión automática de la interfaz: abre TODAS las páginas publicadas a
// ancho de teléfono y comprueba un puñado de invariantes que no se pueden
// romper nunca.
//
// Por qué existe: la guía de interfaz ya decía "la publicidad no tapa el
// encabezado" y aun así las tres herramientas de pantalla completa lo hacían
// durante semanas — la capa del editor era "position: fixed; inset: 0" y
// cubría el encabezado entero. Nadie lo vio porque no hay forma de mirar 120
// páginas a mano. Un documento describe el resultado deseado; esto lo
// comprueba.
//
// Cómo funciona, y por qué así:
//  - Se revisa dist/, no el código fuente: es lo que realmente se publica, así
//    que también atrapa lo que rompa el propio build.
//  - Cada página se mide DENTRO de un <iframe> de 390px. Es la única forma de
//    fijar el viewport: "--window-size" de Edge no lo controla (deja el layout
//    en 492px y solo recorta la captura, así que todo parece cortado a la
//    derecha cuando en realidad entra bien).
//  - Un solo Edge para todas las páginas, cargándolas una por una en el mismo
//    iframe. Levantar un navegador por página tardaba minutos.
//  - Se bloquea todo lo que no sea localhost: las páginas cargan librerías
//    pesadas de CDN que no hacen falta para medir la distribución, y sin
//    bloquearlas la revisión depende de la red.
//
// Uso: npm run revisar      (y falla con código 1 si algo no cumple)

import { createServer } from 'node:http';
import { spawn } from 'node:child_process';
import { readFileSync, writeFileSync, rmSync, existsSync, readdirSync, statSync } from 'node:fs';
import { join, extname, relative, sep } from 'node:path';

const RAIZ = new URL('.', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1');
const DIST = join(RAIZ, 'dist');
const PUERTO = 8788;
const ANCHO = 390;
const ALTO = 844;

// Edge en Windows. Se puede forzar otro con EDGE=<ruta>.
const CANDIDATOS_EDGE = [
  process.env.EDGE,
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
].filter(Boolean);

const TIPOS = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.webmanifest': 'application/manifest+json',
  '.xml': 'application/xml',
  '.txt': 'text/plain; charset=utf-8',
};

function paginasDe(dir, acumulado = []) {
  for (const nombre of readdirSync(dir)) {
    const ruta = join(dir, nombre);
    if (statSync(ruta).isDirectory()) {
      // /en/ se genera a partir del mismo markup: con revisar el español
      // alcanza para la distribución. (Si algún día el inglés diverge, sacar
      // esta línea.)
      if (nombre === 'en') continue;
      paginasDe(ruta, acumulado);
    } else if (extname(nombre) === '.html') {
      acumulado.push('/' + relative(DIST, ruta).split(sep).join('/'));
    }
  }
  return acumulado;
}

// --- Lo que se comprueba, ejecutado dentro del navegador -------------------
// Se manda como texto porque corre en el contexto de la página, no acá.
const COMPROBACIONES = `
// Decir "la página desborda" no alcanza para arreglarla: hay que decir quién.
// Se buscan los elementos que se salen y se reportan los MÁS PROFUNDOS (los
// que no tienen adentro otro que también se salga), porque ese es el que
// realmente es ancho — sus padres solo lo arrastran.
function culpablesDeDesborde(doc, ancho) {
  // Salirse de la pantalla no siempre ensancha la página: lo que está dentro
  // de algo con overflow distinto de visible queda recortado, y lo que cuelga
  // de un position:fixed no participa del tamaño del documento. El fondo
  // decorativo cumple las dos cosas, así que sin este filtro se lleva la culpa
  // en todas las páginas y el verdadero causante queda tapado.
  const inocente = (el) => {
    for (let p = el; p && p !== doc.body; p = p.parentElement) {
      const e = getComputedStyle(p);
      if (p !== el && e.overflowX !== 'visible') return true;
      if (e.position === 'fixed') return true;
    }
    return false;
  };
  const seSalen = [];
  for (const el of doc.querySelectorAll('body *')) {
    const r = el.getBoundingClientRect();
    if (r.width > 0 && r.height > 0 && Math.round(r.right) > ancho + 1 && !inocente(el)) seSalen.push(el);
  }
  const hojas = seSalen.filter((el) => !seSalen.some((otro) => otro !== el && el.contains(otro)));
  const nombre = (el) => {
    let n = el.tagName.toLowerCase();
    if (el.id) n += '#' + el.id;
    else if (el.className && typeof el.className === 'string') n += '.' + el.className.trim().split(/\\s+/).join('.');
    return n + ' (' + Math.round(el.getBoundingClientRect().right) + 'px)';
  };
  const lista = (hojas.length ? hojas : seSalen).slice(0, 3).map(nombre);
  return lista.length ? lista.join(', ') : 'no se pudo identificar';
}

function medir(doc, win) {
  const fallas = [];
  const anchoVista = doc.documentElement.clientWidth;
  const visible = (el) => {
    if (!el) return false;
    const r = el.getBoundingClientRect();
    return r.width > 0 && r.height > 0 && getComputedStyle(el).visibility !== 'hidden';
  };

  // 1. El encabezado del sitio está, se ve, y nada lo tapa.
  const encabezado = doc.querySelector('header.sitio');
  if (!encabezado) {
    fallas.push('no hay <header class="sitio">');
  } else if (!visible(encabezado)) {
    fallas.push('el encabezado del sitio no se ve');
  } else {
    const r = encabezado.getBoundingClientRect();
    const encima = doc.elementFromPoint(Math.round(r.left + 24), Math.round(r.top + r.height / 2));
    if (encima && !encabezado.contains(encima) && encima !== encabezado) {
      const quien = encima.className || encima.tagName;
      fallas.push('algo tapa el encabezado: ' + quien);
    }
  }

  // 2. La publicidad no ocupa el lugar del encabezado.
  //    (El bug real: la capa del editor cubría el header y la tira de anuncio
  //    quedaba como primera cosa de la pantalla.)
  const anuncio = doc.querySelector('.anuncio');
  if (anuncio && visible(anuncio)) {
    const arriba = doc.elementFromPoint(Math.round(anchoVista / 2), 4);
    if (arriba && (anuncio === arriba || anuncio.contains(arriba))) {
      fallas.push('la publicidad está en el borde superior, donde va el encabezado');
    }
  }

  // 3. Nada se sale de la pantalla a lo ancho.
  const de = doc.documentElement;
  if (de.scrollWidth > de.clientWidth + 1) {
    fallas.push('desborde horizontal: ' + de.scrollWidth + 'px en una pantalla de ' + de.clientWidth
      + 'px — lo causa ' + culpablesDeDesborde(doc, de.clientWidth));
  }

  // 4. Toda página tiene un título visible.
  if (!doc.querySelector('h1')) fallas.push('la página no tiene <h1>');

  return fallas;
}
`;

function paginaDeControl(paginas) {
  return `<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8"><title>revision</title>
<style>html,body{margin:0}iframe{width:${ANCHO}px;height:${ALTO}px;border:0;display:block}</style>
</head><body>
<iframe id="marco"></iframe>
<pre id="salida"></pre>
<script>
const PAGINAS = ${JSON.stringify(paginas)};
${COMPROBACIONES}

const marco = document.getElementById('marco');
const resultados = [];

function esperar(ms) { return new Promise((r) => setTimeout(r, ms)); }

function cargar(ruta) {
  return new Promise((resolver) => {
    let listo = false;
    const seguir = () => { if (!listo) { listo = true; resolver(); } };
    marco.onload = seguir;
    marco.onerror = seguir;
    marco.src = ruta;
    setTimeout(seguir, 4000);
  });
}

// En una computadora la barra de desplazamiento ocupa ancho real y se come
// 10px del iframe, así que páginas que entran bien en un teléfono parecían
// desbordar. En un celular la barra va superpuesta: esto iguala las dos.
function barraSuperpuesta(doc) {
  const s = doc.createElement('style');
  s.textContent = 'html{scrollbar-width:none}html::-webkit-scrollbar{display:none}';
  doc.head.appendChild(s);
}

// Solo las herramientas de pantalla completa tienen un segundo estado que
// revisar. Forzar el editor en una página cualquiera que use ese mismo id
// dispara su código a destiempo y el error es del método, no de la página.
function esPantallaCompleta(doc, win) {
  return !!doc.querySelector('.editor-color-layout, .editor-avanzado-layout')
    || !!doc.getElementById('tiraAnuncioMovil');
}

// Una foto de verdad para las herramientas que no muestran su editor hasta
// que hay un archivo cargado.
function archivoDePrueba(win) {
  return new Promise((resolver) => {
    const c = win.document.createElement('canvas');
    c.width = 400; c.height = 300;
    const cx = c.getContext('2d');
    cx.fillStyle = '#4caf50'; cx.fillRect(0, 0, 400, 300);
    c.toBlob((b) => resolver(new win.File([b], 'prueba.png', { type: 'image/png' })), 'image/png');
  });
}

(async () => {
  for (const ruta of PAGINAS) {
    await cargar(ruta);
    await esperar(60);
    let doc, win;
    try { doc = marco.contentDocument; win = marco.contentWindow; } catch (e) {}
    if (!doc) { resultados.push({ ruta, estado: 'inicial', fallas: ['no se pudo abrir la página'] }); continue; }

    barraSuperpuesta(doc);
    await esperar(30);
    resultados.push({ ruta, estado: 'inicial', fallas: medir(doc, win) });

    // Las herramientas de pantalla completa solo arman su distribución cuando
    // hay contenido. Ahí es justamente donde estaba el bug, así que se las
    // revisa también con el editor abierto.
    if (esPantallaCompleta(doc, win)) {
      try {
        const editor = doc.getElementById('editorCompleto');
        if (editor) {
          editor.style.display = 'block';
          await esperar(150);
          resultados.push({ ruta, estado: 'editor abierto', fallas: medir(doc, win) });
        } else if (typeof win.agregarArchivos === 'function') {
          const archivo = await archivoDePrueba(win);
          const dt = new win.DataTransfer();
          dt.items.add(archivo);
          await win.agregarArchivos(dt.files);
          await esperar(150);
          resultados.push({ ruta, estado: 'con contenido', fallas: medir(doc, win) });
        }
      } catch (e) {
        resultados.push({ ruta, estado: 'editor abierto', fallas: ['no se pudo abrir el editor: ' + e.message] });
      }
    }
  }
  document.getElementById('salida').textContent = JSON.stringify(resultados);
})();
</script></body></html>`;
}

function servir(dir) {
  const servidor = createServer((pedido, respuesta) => {
    const ruta = decodeURIComponent(pedido.url.split('?')[0]);
    const archivo = join(dir, ruta === '/' ? 'index.html' : ruta);
    if (!archivo.startsWith(dir) || !existsSync(archivo) || statSync(archivo).isDirectory()) {
      respuesta.writeHead(404).end('no');
      return;
    }
    respuesta.writeHead(200, { 'Content-Type': TIPOS[extname(archivo)] || 'application/octet-stream' });
    respuesta.end(readFileSync(archivo));
  });
  return new Promise((r) => servidor.listen(PUERTO, () => r(servidor)));
}

function correrEdge(url) {
  const edge = CANDIDATOS_EDGE.find((r) => existsSync(r));
  if (!edge) {
    console.error('No se encontró Edge. Indicá la ruta con la variable EDGE.');
    process.exit(1);
  }
  return new Promise((resolver, rechazar) => {
    const proceso = spawn(edge, [
      '--headless',
      '--disable-gpu',
      '--no-sandbox',
      // Sin esto la revisión depende de que respondan los CDN.
      '--host-resolver-rules=MAP * ~NOTFOUND, EXCLUDE localhost',
      '--virtual-time-budget=600000',
      '--dump-dom',
      url,
    ], { windowsHide: true });
    let salida = '';
    proceso.stdout.on('data', (d) => { salida += d; });
    proceso.on('error', rechazar);
    proceso.on('close', () => resolver(salida));
  });
}

// --- Principal -------------------------------------------------------------
if (!existsSync(DIST)) {
  console.error('No existe dist/. Corré primero: npm run build');
  process.exit(1);
}

const paginas = paginasDe(DIST).sort();
const archivoControl = join(DIST, '__revision.html');
writeFileSync(archivoControl, paginaDeControl(paginas), 'utf8');

const servidor = await servir(DIST);
console.log(`Revisando ${paginas.length} páginas a ${ANCHO}px...`);

let salida;
try {
  salida = await correrEdge(`http://localhost:${PUERTO}/__revision.html`);
} finally {
  servidor.close();
  rmSync(archivoControl, { force: true });
}

// --dump-dom devuelve el DOM entero, con el <script> de control incluido, así
// que hay que sacar el resultado del <pre> y no buscar una marca de texto: la
// marca también aparecería dentro del script.
const bruto = salida.match(/<pre id="salida">([\s\S]*?)<\/pre>/);
if (!bruto) {
  console.error('La revisión no devolvió resultados (¿se quedó cargando alguna página?).');
  process.exit(1);
}

const entidades = { '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&#39;': "'" };
const resultados = JSON.parse(bruto[1].replace(/&(amp|lt|gt|quot|#39);/g, (e) => entidades[e]));
const conFallas = resultados.filter((r) => r.fallas.length);

if (!conFallas.length) {
  console.log(`Interfaz OK — ${resultados.length} comprobaciones sobre ${paginas.length} páginas.`);
  process.exit(0);
}

// Agrupado por tipo de falla: es mucho más útil ver "23 páginas desbordan"
// que una lista plana de 23 líneas iguales.
const porFalla = new Map();
for (const r of conFallas) {
  for (const f of r.fallas) {
    const corte = f.indexOf(': ');
    const clave = corte > 0 ? f.slice(0, corte) : f;
    const detalle = corte > 0 ? f.slice(corte + 2) : '';
    if (!porFalla.has(clave)) porFalla.set(clave, []);
    porFalla.get(clave).push(`${r.ruta} (${r.estado})` + (detalle ? `\n          ${detalle}` : ''));
  }
}

console.error(`\nInterfaz con problemas en ${conFallas.length} de ${resultados.length} comprobaciones:\n`);
for (const [falla, paginasConEsto] of [...porFalla].sort((a, b) => b[1].length - a[1].length)) {
  console.error(`  ${falla} — ${paginasConEsto.length}`);
  for (const p of paginasConEsto.slice(0, 12)) console.error(`      ${p}`);
  if (paginasConEsto.length > 12) console.error(`      ...y ${paginasConEsto.length - 12} más`);
  console.error('');
}
process.exit(1);

/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. */
// Respaldo local de TODO lo que el sitio carga de servidores ajenos: las
// librerías de las CDN, los archivos que algunas librerías descargan solas
// mientras funcionan (motor y diccionarios del OCR, WebAssembly de la IA), los
// modelos de IA y las fuentes de Google Fonts. Con la licencia de cada una.
//
// Por qué existe: si una CDN se cae, borra una versión o cambia de dirección,
// las herramientas que la usan dejan de funcionar y no habría de dónde sacar
// ESA versión exacta. Desde el 2026-10-01 las librerías se sirven desde el
// propio sitio (carpeta vendor/, versionada en git), así que lo que este
// respaldo sigue cubriendo de verdad es lo que todavía viene de afuera: el
// motor de IA (onnxruntime-web, en jsDelivr) y los modelos (Hugging Face),
// que pasan los 25 MB por archivo de Cloudflare Pages. Cualquier dirección de
// CDN que vuelva a aparecer en el código también se respalda sola.
//
// Uso:  node respaldar-librerias.mjs [--sin-modelos] [carpeta-destino]
//       (por defecto ../librerias-respaldo, fuera del repo: con los modelos
//       pesa cientos de MB y no debe ir a git)
//
// Volver a correrlo al agregar o actualizar una librería: detecta solas las
// direcciones escritas en el código; lo que no se puede detectar (lo que
// descargan las librerías por dentro) está en las listas de abajo.

import { readdirSync, readFileSync, writeFileSync, mkdirSync, existsSync, statSync } from 'node:fs';
import { join, dirname, extname } from 'node:path';
import { createHash } from 'node:crypto';

const RAIZ = new URL('.', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1');
const args = process.argv.slice(2);
const SIN_MODELOS = args.includes('--sin-modelos');
const DESTINO = args.find((a) => !a.startsWith('--')) || join(RAIZ, '..', 'librerias-respaldo');

const HOSTS_CDN = ['cdnjs.cloudflare.com', 'cdn.jsdelivr.net', 'fonts.googleapis.com'];

// --- Lo que las librerías descargan por su cuenta --------------------------
// Tesseract (ocr-pdf): el worker, el motor en WebAssembly (elige la variante
// según el navegador, así que van todas) y un diccionario por idioma.
const TESSERACT = '7.0.0';
const IDIOMAS_OCR = ['spa', 'eng', 'por', 'fra', 'deu', 'ita', 'chi_sim', 'jpn']; // los de ocr-pdf.html
// transformers.js (dictado-voz, generador-stickers) baja el WebAssembly de
// onnxruntime-web. La versión sale de las dependencias de transformers.
const TRANSFORMERS = '4.3.0';
// Modelos de IA, con el dtype que pide cada herramienta ('q8' → *_quantized.onnx).
const MODELOS = ['Xenova/whisper-tiny', 'Xenova/whisper-base', 'Xenova/whisper-small', 'xrds/isnet-general-onnx-int8'];

// Paquete de npm de cada librería de cdnjs, para bajar su licencia.
const NPM_DE_CDNJS = {
  jszip: 'jszip', 'pdf.js': 'pdfjs-dist', 'pdf-lib': 'pdf-lib', xlsx: 'xlsx', mammoth: 'mammoth',
  marked: 'marked', 'Chart.js': 'chart.js', turndown: 'turndown', jsdiff: 'diff', jsbarcode: 'jsbarcode',
};
// Las que no están en npm con esa versión: licencia desde su repositorio.
// También las fuentes (OFL), que no son paquetes de npm.
const LICENCIA_DIRECTA = {
  qrcodejs: 'https://cdn.jsdelivr.net/gh/davidshimjs/qrcodejs@04f46c6/LICENSE',
  // fontkit no publica archivo de licencia: el único registro de que es MIT
  // es el campo "license" de su package.json, así que se guarda ese.
  '@pdf-lib__fontkit@1.1.1': 'https://cdn.jsdelivr.net/npm/@pdf-lib/fontkit@1.1.1/package.json',
  'jsbarcode@3.12.3': 'https://cdn.jsdelivr.net/gh/lindell/JsBarcode@v3.12.3/MIT-LICENSE.txt',
  'onnxruntime-web': 'https://cdn.jsdelivr.net/gh/microsoft/onnxruntime@main/LICENSE',
  'Noto Sans': 'https://cdn.jsdelivr.net/gh/google/fonts@2984c575fdce412ee02b2baaba67672b9a9434d8/ofl/notosans/OFL.txt',
  'Public Sans': 'https://cdn.jsdelivr.net/gh/google/fonts@main/ofl/publicsans/OFL.txt',
  'JetBrains Mono': 'https://cdn.jsdelivr.net/gh/google/fonts@main/ofl/jetbrainsmono/OFL.txt',
  tessdata_best: 'https://cdn.jsdelivr.net/gh/tesseract-ocr/tessdata_best@main/LICENSE',
};

// --- Utilidades -------------------------------------------------------------
const manifiesto = [];
let bytesTotales = 0;

async function bajar(url, archivo, tipo) {
  if (existsSync(archivo) && statSync(archivo).size > 0) {
    const datos = readFileSync(archivo);
    registrar(url, archivo, datos, tipo, true);
    return true;
  }
  for (let intento = 1; intento <= 3; intento++) {
    try {
      const r = await fetch(url, { headers: { 'User-Agent': 'Mozilla/5.0 (respaldo Mini Cookies)' } });
      if (!r.ok) throw new Error('HTTP ' + r.status);
      const datos = Buffer.from(await r.arrayBuffer());
      mkdirSync(dirname(archivo), { recursive: true });
      writeFileSync(archivo, datos);
      registrar(url, archivo, datos, tipo, false);
      return true;
    } catch (e) {
      if (intento === 3) {
        console.error(`  ✗ ${url} — ${e.message}`);
        manifiesto.push({ url, tipo, error: e.message });
        return false;
      }
    }
  }
}

function registrar(url, archivo, datos, tipo, yaEstaba) {
  bytesTotales += datos.length;
  manifiesto.push({
    url, tipo,
    archivo: archivo.slice(DESTINO.length + 1).split('\\').join('/'),
    bytes: datos.length,
    sha256: createHash('sha256').update(datos).digest('hex'),
  });
  console.log(`  ${yaEstaba ? '=' : '✓'} ${url} (${(datos.length / 1048576).toFixed(2)} MB)`);
}

// cdn/<host>/<ruta>: la misma ruta que en la CDN.
function rutaEspejo(url) {
  const u = new URL(url);
  let ruta = decodeURIComponent(u.pathname);
  if (u.search) ruta += '__' + createHash('sha1').update(u.search).digest('hex').slice(0, 10) + (extname(ruta) ? '' : '.css');
  return join(DESTINO, 'cdn', u.host, ruta);
}

async function listarJsdelivr(paquete) {
  const r = await fetch(`https://data.jsdelivr.com/v1/packages/npm/${paquete}?structure=flat`);
  if (!r.ok) throw new Error(`no se pudo listar ${paquete}: HTTP ${r.status}`);
  return (await r.json()).files;
}

function archivosFuente(dir, acumulado = []) {
  for (const nombre of readdirSync(dir)) {
    if (['node_modules', 'dist', '.git', '.vscode', '.claude'].includes(nombre)) continue;
    const ruta = join(dir, nombre);
    if (statSync(ruta).isDirectory()) archivosFuente(ruta, acumulado);
    else if (['.html', '.js', '.mjs', '.css'].includes(extname(nombre)) && nombre !== 'respaldar-librerias.mjs') acumulado.push(ruta);
  }
  return acumulado;
}

// --- 1. Direcciones escritas en el código ----------------------------------
console.log(`Respaldo en ${DESTINO}\n`);
const directas = new Set();
const regex = /https:\/\/(?:cdnjs\.cloudflare\.com|cdn\.jsdelivr\.net|fonts\.googleapis\.com)\/[^"'`\s)<>]+/g;
for (const f of archivosFuente(RAIZ)) {
  for (const m of readFileSync(f, 'utf8').matchAll(regex)) {
    if (!m[0].includes('${')) directas.add(m[0].replace(/&amp;/g, '&'));
  }
}
console.log(`1. Librerías referenciadas en el código: ${directas.size}`);
for (const url of [...directas].sort()) {
  await bajar(url, rutaEspejo(url), new URL(url).host === 'fonts.googleapis.com' ? 'fuente-css' : 'libreria');
}

// Google Fonts: el CSS apunta a los archivos de fuente, que también hay que tener.
// Se pide con un User-Agent de navegador moderno para recibir woff2.
for (const url of directas) {
  if (new URL(url).host !== 'fonts.googleapis.com') continue;
  const css = await (await fetch(url, { headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36' } })).text();
  writeFileSync(rutaEspejo(url), css);
  for (const m of css.matchAll(/url\((https:\/\/fonts\.gstatic\.com\/[^)]+)\)/g)) await bajar(m[1], rutaEspejo(m[1]), 'fuente');
}

// --- 2. Lo que descargan las librerías por dentro --------------------------
console.log('\n2. Archivos que las librerías descargan mientras funcionan');
await bajar(`https://cdn.jsdelivr.net/npm/tesseract.js@v${TESSERACT}/dist/worker.min.js`,
  rutaEspejo(`https://cdn.jsdelivr.net/npm/tesseract.js@v${TESSERACT}/dist/worker.min.js`), 'runtime');
for (const f of await listarJsdelivr(`tesseract.js-core@${TESSERACT}`)) {
  if (!/^\/tesseract-core[^/]*\.(js|wasm)$/.test(f.name)) continue;
  const url = `https://cdn.jsdelivr.net/npm/tesseract.js-core@v${TESSERACT}${f.name}`;
  await bajar(url, rutaEspejo(url), 'runtime');
}
for (const idioma of IDIOMAS_OCR) {
  const url = `https://cdn.jsdelivr.net/npm/@tesseract.js-data/${idioma}/4.0.0_best_int/${idioma}.traineddata.gz`;
  await bajar(url, rutaEspejo(url), 'runtime');
}
const ort = (await (await fetch(`https://registry.npmjs.org/@huggingface/transformers/${TRANSFORMERS}`)).json()).dependencies['onnxruntime-web'];
for (const f of await listarJsdelivr(`onnxruntime-web@${ort}`)) {
  if (!/^\/dist\/ort-wasm-simd-threaded[^/]*\.(mjs|wasm)$/.test(f.name)) continue;
  const url = `https://cdn.jsdelivr.net/npm/onnxruntime-web@${ort}${f.name}`;
  await bajar(url, rutaEspejo(url), 'runtime');
}

// --- 3. Modelos de IA --------------------------------------------------------
if (SIN_MODELOS) {
  console.log('\n3. Modelos de IA: omitidos (--sin-modelos)');
} else {
  console.log('\n3. Modelos de IA (Hugging Face)');
  for (const modelo of MODELOS) {
    const info = await (await fetch(`https://huggingface.co/api/models/${modelo}`)).json();
    for (const { rfilename } of info.siblings) {
      const esOnnx = rfilename.endsWith('.onnx') || rfilename.endsWith('.onnx_data');
      // dtype 'q8' → solo los *_quantized; el resto de las variantes no se usa.
      if (esOnnx && !/_quantized\.onnx(_data)?$/.test(rfilename)) continue;
      // Whisper trae el decodificador en tres formas; transformers.js usa
      // solo la "merged". Las otras dos son ~380 MB que nunca se cargan.
      if (/\/decoder_(model|with_past_model)_quantized\.onnx/.test(rfilename)) continue;
      if (!esOnnx && !/\.(json|txt|model|md)$/.test(rfilename) && rfilename !== 'LICENSE') continue;
      await bajar(`https://huggingface.co/${modelo}/resolve/main/${rfilename}`, join(DESTINO, 'modelos', modelo, rfilename), 'modelo');
    }
  }
}

// --- 4. Licencias ------------------------------------------------------------
console.log('\n4. Licencias');
const paquetes = new Map();
for (const url of directas) {
  const jsd = url.match(/cdn\.jsdelivr\.net\/npm\/((?:@[^/]+\/)?[^@/]+)@([^/]+)/);
  if (jsd) paquetes.set(jsd[1], jsd[2].replace(/^v/, ''));
  const cdnjs = url.match(/cdnjs\.cloudflare\.com\/ajax\/libs\/([^/]+)\/([^/]+)\//);
  if (cdnjs && NPM_DE_CDNJS[cdnjs[1]]) paquetes.set(NPM_DE_CDNJS[cdnjs[1]], cdnjs[2]);
}
paquetes.set('tesseract.js-core', TESSERACT);
paquetes.set('onnxruntime-web', ort);
for (const [paquete, version] of [...paquetes].sort()) {
  let archivos;
  try { archivos = await listarJsdelivr(`${paquete}@${version}`); } catch (e) { console.error(`  ✗ ${paquete}@${version} — ${e.message}`); continue; }
  const licencia = archivos.find((f) => /^\/(licen[cs]e|copying)(\.[a-z]+)?$/i.test(f.name));
  if (!licencia) continue; // la trae LICENCIA_DIRECTA
  await bajar(`https://cdn.jsdelivr.net/npm/${paquete}@${version}${licencia.name}`,
    join(DESTINO, 'licencias', `${paquete.replace('/', '__')}@${version}`, licencia.name.slice(1)), 'licencia');
}
for (const [nombre, url] of Object.entries(LICENCIA_DIRECTA)) {
  await bajar(url, join(DESTINO, 'licencias', nombre, url.split('/').pop()), 'licencia');
}

// --- Manifiesto ----------------------------------------------------------------
const errores = manifiesto.filter((m) => m.error);
writeFileSync(join(DESTINO, 'MANIFIESTO.json'), JSON.stringify({
  generado: new Date().toISOString(),
  aviso: 'Copia de respaldo de recursos de terceros que usa minicookies.net. Cada uno conserva su licencia original (carpeta licencias/).',
  totalArchivos: manifiesto.length - errores.length,
  totalBytes: bytesTotales,
  archivos: manifiesto,
}, null, 2));

console.log(`\n${manifiesto.length - errores.length} archivos, ${(bytesTotales / 1048576).toFixed(1)} MB.`);
if (errores.length) {
  console.error(`${errores.length} no se pudieron bajar (ver MANIFIESTO.json).`);
  process.exit(1);
}

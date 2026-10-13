/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. */
// Build de publicación: copia el sitio a dist/ con el JavaScript y el CSS
// minificados (también los <script> y <style> dentro de cada HTML) y sin
// comentarios HTML. El código fuente legible queda solo en el repo privado;
// lo publicado es lo mismo pero comprimido y más difícil de reutilizar.
//
// Cloudflare Pages lo corre en cada push: build command "npm run build",
// output directory "dist". Para probarlo en local: npm install && npm run build.
//
// Los comentarios /*! ... */ (avisos de copyright) se conservan a propósito.
// Si algo no se puede minificar, el build falla y Cloudflare deja publicada
// la versión anterior en vez de subir un sitio roto.
//
// Este mismo archivo está en los dos repos (MiniCookies-Home y
// HerramientasOfimaticas): si se cambia, copiarlo al otro.
import { transformSync } from 'esbuild';
import { readdirSync, readFileSync, writeFileSync, mkdirSync, copyFileSync, rmSync, statSync, existsSync } from 'node:fs';
import { inyectarSeo, generarRobotsYSitemap, indiceDeCategorias, esNoIndex } from './seo.mjs';
import { traducir, diccionarioComun, paginasConIngles, tieneIngles } from './ingles.mjs';
import { DESTINO, ES_SEGURO, SEGURAS, CARPETA_SALIDA, SITIO_PRINCIPAL, esHerramientaSegura, urlSegura } from './destinos.mjs';
import { contarHerramientas, aplicarConteos } from './conteo.mjs';
import { leerCatalogo, aplicarCatalogo } from './catalogo.mjs';
import { revisarTerceros } from './terceros.mjs';
import { revisarReferenciasVendor } from './referencias-vendor.mjs';
import { revisarMotorPptx } from './motor-pptx.mjs';
import { revisarDependenciasScripts } from './dependencias-scripts.mjs';
import { join, relative, extname, dirname } from 'node:path';
import { createHash } from 'node:crypto';

const RAIZ = process.cwd();
const SALIDA = join(RAIZ, CARPETA_SALIDA);

// Lo que no se publica: herramientas del build, notas internas (.md) y
// archivos del repo.
const EXCLUIR_NOMBRES = new Set(['.git', '.github', '.claude', '.wrangler', 'node_modules', 'dist', 'dist-seguro', 'destinos.mjs', 'seguro-index.html',
  'build.mjs', 'seo.mjs', 'ingles.mjs', 'conteo.mjs', 'catalogo.mjs', 'catalogo.json', 'terceros.mjs', 'referencias-vendor.mjs', 'motor-pptx.mjs', 'revisar-interfaz.mjs', 'probar-funciones.mjs', 'respaldar-librerias.mjs', 'comparar-pptx.mjs', 'comparar-pptx', 'package.json', 'package-lock.json', '.gitignore', '.DS_Store', 'Thumbs.db']);
const EXCLUIR_EXTENSIONES = new Set(['.md', '.log']);

// Aviso de copyright de cada sitio: campo "config.aviso" de su package.json.
const AVISO = `/*! ${JSON.parse(readFileSync(join(RAIZ, 'package.json'), 'utf8')).config.aviso} */`;
// Scripts dentro del HTML a partir de este tamaño reciben el aviso de
// copyright (los chiquitos, como el del tema, no lo necesitan).
const MINIMO_PARA_AVISO = 1500;

// Etiqueta de Google AdSense. Va en TODAS las paginas, no solo en la portada:
// ademas de servir los anuncios, es la que verifica la propiedad del sitio y
// la que muestra el mensaje de consentimiento de cookies a los visitantes de
// la UE (Privacy & messaging usa esta misma etiqueta, no un script aparte).
// El ID sale de package.json para no repetirlo en 119 archivos.
const ADSENSE = JSON.parse(readFileSync(join(RAIZ, 'package.json'), 'utf8')).config.adsense;
const ETIQUETA_ADSENSE = ADSENSE
  ? `<script async src="https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js?client=${ADSENSE}" crossorigin="anonymous"></script>`
  : '';

// Unidad de anuncio manual (AdSense > Anuncios > Por unidad de anuncio): su
// "data-ad-slot" va en config.adsenseSlot. Mientras esté vacío, cada .anuncio
// queda con el texto "Espacio publicitario" y nada más.
//
// Cada hueco se convierte en:
//   <div class="anuncio"><span class="anuncio-texto" data-i18n="anuncio">…</span><ins class="adsbygoogle" …></ins></div>
// El texto pasa a un <span> porque idioma.js reemplaza el textContent de todo
// [data-i18n]: dejado en el <div>, borraba el anuncio al cambiar de idioma. El
// texto se ve solo si Google no llena el hueco (ver .anuncio en estilos.css).
// Sin data-ad-format: el tamaño lo fija el CSS del <ins> (alto fijo, regla de
// SISTEMA-DE-INTERFAZ.md) y Google elige un anuncio que entre ahí.
const SLOT_ADSENSE = JSON.parse(readFileSync(join(RAIZ, 'package.json'), 'utf8')).config.adsenseSlot;
const HUECO_ANUNCIO = /<div class="anuncio"([^>]*?) data-i18n="anuncio"([^>]*)>([^<]*)<\/div>/g;

function ponerUnidadesDeAnuncio(html, archivo) {
  let cuantos = 0;
  html = html.replace(HUECO_ANUNCIO, (todo, antes, despues, texto) => {
    cuantos++;
    return `<div class="anuncio"${antes}${despues}><span class="anuncio-texto" data-i18n="anuncio">${texto}</span>`
      + `<ins class="adsbygoogle" data-ad-client="${ADSENSE}" data-ad-slot="${SLOT_ADSENSE}"></ins></div>`;
  });
  // Un hueco escrito de otra forma (sin data-i18n, con otro orden) no se
  // convertiría y quedaría sin anuncio sin que nadie lo note.
  if (/class="anuncio"/.test(html.replace(/<div class="anuncio"[^>]*><span class="anuncio-texto"/g, '')))
    throw new Error('Hueco .anuncio con un formato que el build no reconoce: ' + archivo);
  if (!cuantos) return html;
  // Un push por <ins>, al final del body: cuando corre, todos ya están en la
  // página. Antes del ÚLTIMO </body>: hay páginas con otro dentro de una cadena
  // de su JavaScript (el documento de la vista previa o de impresión), y
  // ponerlo ahí les rompía el script.
  const fin = html.toLowerCase().lastIndexOf('</body>');
  if (fin < 0) throw new Error('Sin </body>, no se pueden pedir los anuncios: ' + archivo);
  return html.slice(0, fin)
    + `<script>document.querySelectorAll('ins.adsbygoogle').forEach(function(){(window.adsbygoogle=window.adsbygoogle||[]).push({})})</script>\n`
    + html.slice(fin);
}

function esModulo(codigo) {
  return /^\s*(import|export)\s/m.test(codigo);
}

function minificarJs(codigo, archivo, forzarModulo = false) {
  return transformSync(codigo, {
    loader: 'js',
    minify: true,
    // Sin "format": los nombres de primer nivel NO se renombran, porque los
    // scripts del sitio se comparten variables globales entre sí
    // (AjustePorColor, FiltrosColor, funciones llamadas desde el HTML...).
    format: forzarModulo || esModulo(codigo) ? 'esm' : undefined,
    legalComments: 'inline',
    charset: 'utf8',
    sourcefile: archivo,
  }).code.trim();
}

function minificarCss(codigo, archivo) {
  return transformSync(codigo, { loader: 'css', minify: true, legalComments: 'inline', charset: 'utf8', sourcefile: archivo }).code.trim();
}

function procesarHtml(html, archivo) {
  // Comentarios HTML (notas de desarrollo): fuera.
  html = html.replace(/<!--(?!\[if)[\s\S]*?-->/g, '');
  // <script> sin src: minificar su contenido. Los de otro tipo (JSON-LD,
  // plantillas...) quedan como están.
  html = html.replace(/<script(\s[^>]*)?>([\s\S]*?)<\/script>/gi, (completo, atributos = '', codigo) => {
    if (/\ssrc\s*=/i.test(atributos) || !codigo.trim()) return completo;
    const tipo = (atributos.match(/\stype\s*=\s*["']?([^"'\s>]+)/i) || [])[1];
    if (tipo && !/^(module|text\/javascript|application\/javascript)$/i.test(tipo)) return completo;
    let min = minificarJs(codigo, `${archivo} <script>`, tipo === 'module');
    if (min.length >= MINIMO_PARA_AVISO && !min.startsWith('/*!')) min = `${AVISO}\n${min}`;
    return `<script${atributos}>${min}</script>`;
  });
  html = html.replace(/<style(\s[^>]*)?>([\s\S]*?)<\/style>/gi, (completo, atributos = '', codigo) =>
    (codigo.trim() ? `<style${atributos}>${minificarCss(codigo, `${archivo} <style>`)}</style>` : completo));
  // Sin AdSense en las páginas que no se indexan (la 404): Google no admite
  // anuncios en páginas de error o sin contenido.
  // Tampoco en las páginas que guardan datos sensibles (administrador de contraseñas): un
  // script de anuncios correría en la misma página que la contraseña maestra.
  // El sitio seguro no lleva anuncios en ninguna página (destinos.mjs).
  if (ETIQUETA_ADSENSE && !ES_SEGURO && !html.includes('adsbygoogle.js') && !esNoIndex(html) && !html.includes('<meta name="sin-terceros"')) {
    if (!/<\/head>/i.test(html)) throw new Error('Sin </head>, no se puede poner AdSense: ' + archivo);
    html = html.replace(/<\/head>/i, ETIQUETA_ADSENSE + '\n</head>');
    if (SLOT_ADSENSE) html = ponerUnidadesDeAnuncio(html, archivo);
  }
  return html;
}

/**
 * Enlaces internos sin ".html".
 *
 * El código fuente enlaza "pagina.html" para que las páginas se puedan abrir
 * directo del disco al desarrollar. Pero Cloudflare Pages sirve la URL sin
 * extensión y responde 308 cuando se la pide con ella, así que cada enlace del
 * sitio costaba dos peticiones y, peor, la URL enlazada no era la misma que la
 * página declara como canónica ni la que está en el sitemap. Un rastreador veía
 * esa incoherencia en las 120 páginas.
 *
 * "index" es la raíz de su carpeta: ../index.html -> ../ (pedir /index también
 * redirige).
 *
 * Solo se tocan los atributos del marcado: el contenido de <script> se aparta
 * antes, para no reescribir cadenas dentro del JavaScript.
 */
function enlacesLimpios(html) {
  const guardados = [];
  const sinScripts = html.replace(/<script[\s\S]*?<\/script>/gi, (bloque) => {
    guardados.push(bloque);
    return `\u0000S${guardados.length - 1}\u0000`;
  });

  const limpio = sinScripts.replace(/(\s(?:href|action)\s*=\s*")([^"]+)(")/gi, (todo, antes, url, despues) => {
    // Absolutas, protocolo relativo, anclas y mailto: no se tocan.
    if (/^[a-z][a-z0-9+.-]*:/i.test(url) || url.startsWith('//') || url.startsWith('#')) return todo;
    if (!/\.html$/i.test(url)) return todo;
    let destino = url.slice(0, -'.html'.length).replace(/(^|\/)index$/i, '$1');
    if (destino === '') destino = './';
    return antes + destino + despues;
  });

  return limpio.replace(/\u0000S(\d+)\u0000/g, (_, i) => guardados[i]);
}

// Copias de librerías de terceros: se publican byte por byte y completas.
// Minificarlas les agregaría nuestro aviso de copyright y podría quitarles el
// suyo; sus licencias (algunas .md, o un package.json cuando el paquete no trae
// otra) tienen que publicarse junto a cada una; y las exclusiones del sitio no
// aplican: muchas guardan sus archivos en una carpeta "dist", que para el
// sitio es la salida del build y se saltea.
function copiarVendor(carpeta) {
  for (const nombre of readdirSync(carpeta)) {
    const origen = join(carpeta, nombre);
    if (statSync(origen).isDirectory()) { copiarVendor(origen); continue; }
    const destino = join(SALIDA, relative(RAIZ, origen));
    mkdirSync(dirname(destino), { recursive: true });
    copyFileSync(origen, destino);
  }
}

// Páginas de la raíz del repo que, en el sitio seguro, apuntan al principal
// (el sitio seguro no tiene portada de herramientas, ni categorías, ni textos
// legales propios: los legales viven una sola vez, en www).
function enlacesAlPrincipal(html, rel, relSalida) {
  const dirOrigen = rel.includes('/') ? rel.slice(0, rel.lastIndexOf('/')) : '';
  // Lo que SÍ existe en el sitio seguro (por su ruta de salida).
  const propias = new Set([...SEGURAS.map((h) => 'herramientas/' + h + '.html'), 'seguro-index.html', '404.html']);
  return html.replace(/(\shref\s*=\s*")([^"#:][^"]*\.html)(")/gi, (todo, antes, url, despues) => {
    if (url.startsWith('/')) return todo;
    const pila = [];
    for (const parte of (dirOrigen ? dirOrigen.split('/') : []).concat(url.split('/'))) {
      if (parte === '.' || parte === '') continue;
      if (parte === '..') pila.pop(); else pila.push(parte);
    }
    const resuelto = pila.join('/');
    if (propias.has(resuelto)) return todo;
    // index.html (la marca del encabezado) también va a www: la portada del sitio seguro es otra cosa.
    const limpio = resuelto.replace(/\.html$/i, '').replace(/(^|\/)index$/i, '$1');
    return antes + SITIO_PRINCIPAL + '/' + limpio + despues;
  });
}

// En la versión en inglés, los enlaces absolutos a www que tienen gemela en
// inglés pasan a /en/ (los de arriba se escriben una vez, en español).
function inglesDelPrincipal(html) {
  return html.replace(new RegExp('(\\shref\\s*=\\s*")' + SITIO_PRINCIPAL.replace(/\./g, '\\.') + '/([a-z0-9-]*)(")', 'gi'), (todo, antes, nombre, despues) => {
    if (nombre === '' || conInglesPrincipal.has(nombre + '.html')) return antes + SITIO_PRINCIPAL + '/en/' + nombre + despues;
    return todo;
  });
}

// Recursos locales (scripts, estilos, imágenes) que cargan las páginas del
// sitio seguro: es lo único que se copia de assets/ y vendor/.
const recursos = new Set();
function anotarRecursos(crudo, rel) {
  const dir = rel.includes('/') ? rel.slice(0, rel.lastIndexOf('/')) : '';
  for (const m of crudo.matchAll(/\s(?:src|href)\s*=\s*"([^"#?:][^"#?]*)/gi)) {
    const pila = [];
    for (const parte of (dir ? dir.split('/') : []).concat(m[1].split('/'))) {
      if (parte === '.' || parte === '') continue;
      if (parte === '..') pila.pop(); else pila.push(parte);
    }
    const ruta = pila.join('/');
    if (/^(assets|marca-logo|vendor)\//.test(ruta)) recursos.add(ruta);
  }
}

// Cuerpos de los <script> en línea que el navegador ejecuta (no los JSON-LD ni las plantillas).
function scriptsEnLinea(html) {
  const cuerpos = [];
  html.replace(/<script(\s[^>]*)?>([\s\S]*?)<\/script>/gi, (todo, atributos = '', codigo) => {
    if (/\ssrc\s*=/i.test(atributos) || !codigo.trim()) return todo;
    const tipo = (atributos.match(/\stype\s*=\s*["']?([^"'\s>]+)/i) || [])[1];
    if (!tipo || /^(module|text\/javascript|application\/javascript)$/i.test(tipo)) cuerpos.push(codigo);
    return todo;
  });
  return cuerpos;
}

// Política de contenido del sitio seguro. En vez de 'unsafe-inline' para los scripts, la lista
// exacta de los que la página trae (por su hash SHA-256): lo que Cloudflare inyecta al servir
// (detección de bots, analítica) y cualquier HTML colado quedan bloqueados. Los estilos sí
// siguen con 'unsafe-inline' (las páginas usan style="…"; no ejecutan código).
function politicaSeguro(nombre, hashes) {
  const fuentes = ["'self'"].concat(hashes.map((h) => "'sha256-" + h + "'")).join(' ');
  const csp = (extra) => "default-src 'self'; script-src " + fuentes + "; style-src 'self' 'unsafe-inline'; "
    + extra + "; form-action 'none'; base-uri 'none'; object-src 'none'";
  const general = { csp: csp("img-src 'self' data:; connect-src 'none'; frame-src 'none'; worker-src 'none'"), permisos: 'camera=(), microphone=(), geolocation=(), clipboard-read=()', extra: [] };
  const politicas = {
    'administrador-contrasenas': Object.assign({}, general, { extra: ['Cache-Control: no-store'] }),
    'enviar-archivos-qr': { csp: csp("img-src 'self' data: blob:; media-src 'self' blob:; connect-src 'self'; frame-src 'self'; worker-src 'self'"), permisos: 'camera=(self), microphone=(), geolocation=(), clipboard-read=()', extra: [] },
  };
  return politicas[nombre] || general;
}

const hashesPorPagina = new Map(); // ruta de salida → hashes de sus scripts en línea

// Pone la CSP (con los hashes de lo que la página trae) en un <meta> y la anota para _headers.
function sellarSeguro(html, relSalida) {
  html = html.replace(/<meta\s+http-equiv="Content-Security-Policy"[^>]*>\s*/gi, '');
  const hashes = [...new Set(scriptsEnLinea(html).map((c) => createHash('sha256').update(c, 'utf8').digest('base64')))];
  hashesPorPagina.set(relSalida.replace(/\\/g, '/'), hashes);
  const nombre = (/herramientas\/([^/]+)\.html$/.exec(relSalida.replace(/\\/g, '/')) || [])[1];
  const meta = '<meta http-equiv="Content-Security-Policy" content="' + politicaSeguro(nombre, hashes).csp + '">';
  return html.replace(/<head>/i, () => '<head>\n' + meta);
}

function procesarArchivo(origen, rel, relSalida = rel) {
  const ext = extname(origen).toLowerCase();
  const destino = join(SALIDA, relSalida);
  mkdirSync(dirname(destino), { recursive: true });
  if (ext === '.js' || ext === '.mjs') writeFileSync(destino, minificarJs(readFileSync(origen, 'utf8'), rel) + '\n');
  else if (ext === '.css') writeFileSync(destino, minificarCss(readFileSync(origen, 'utf8'), rel) + '\n');
  else if (ext === '.html') {
    const rutaWeb = relSalida.replace(/\\/g, '/');
    let fuente = readFileSync(origen, 'utf8');
    if (ES_SEGURO) {
      anotarRecursos(fuente, rel);
      fuente = enlacesAlPrincipal(fuente, rel.replace(/\\/g, '/'), rutaWeb);
      // Este sitio no tiene anuncios: tampoco el hueco que dice "Espacio publicitario".
      fuente = fuente.replace(/<div class="anuncio"[^>]*>[\s\S]*?<\/div>\s*/g, '');
    }
    // Primero el catálogo (genera las tarjetas), después los conteos.
    const crudo = aplicarConteos(aplicarCatalogo(fuente, catalogo, rutaWeb, opcionesCatalogo), conteo, rutaWeb);
    // Al sitemap solo van las páginas que se indexan (no la 404).
    if (!esNoIndex(crudo)) paginas.push(relSalida);
    const final = enlacesLimpios(inyectarSeo(procesarHtml(crudo, relSalida), relSalida, indiceCategorias, conIngles));
    writeFileSync(destino, ES_SEGURO ? sellarSeguro(final, relSalida) : final);

    // Gemela en inglés bajo /en/, con su propia URL para que Google la vea.
    if (conIngles.has(rutaWeb)) {
      let en = traducir(crudo, RAIZ, comun, rutaWeb, conIngles);
      if (!en) throw new Error('No se pudo traducir ' + rutaWeb);
      // Las tarjetas hacia el sitio seguro (y desde él hacia www) en su versión en inglés.
      en = en.split(SITIO_SEGURO_HERRAMIENTAS).join(SITIO_SEGURO_HERRAMIENTAS_EN);
      if (ES_SEGURO) en = inglesDelPrincipal(en);
      const destinoEn = join(SALIDA, 'en', relSalida);
      mkdirSync(dirname(destinoEn), { recursive: true });
      const finalEn = enlacesLimpios(inyectarSeo(procesarHtml(en, 'en/' + relSalida), relSalida, indiceCategorias, conIngles, true));
      writeFileSync(destinoEn, ES_SEGURO ? sellarSeguro(finalEn, 'en/' + relSalida) : finalEn);
      if (!esNoIndex(crudo)) paginasEn.push('en/' + rutaWeb);
    }
  }
  else copyFileSync(origen, destino);
}

function recorrer(carpeta) {
  for (const nombre of readdirSync(carpeta)) {
    const origen = join(carpeta, nombre);
    const rel = relative(RAIZ, origen);
    if (rel === 'vendor') { copiarVendor(origen); continue; }
    if (EXCLUIR_NOMBRES.has(nombre)) continue;
    // Nada oculto se publica (.vscode, .cbmignore, .gitattributes…): son
    // configuración del editor o de herramientas, no parte del sitio.
    if (nombre.startsWith('.')) continue;
    if (statSync(origen).isDirectory()) { recorrer(origen); continue; }
    if (EXCLUIR_EXTENSIONES.has(extname(nombre).toLowerCase())) continue;
    // Las herramientas seguras solo se publican en el sitio seguro.
    if (esHerramientaSegura(rel)) continue;
    procesarArchivo(origen, rel);
  }
}

// El sitio seguro se arma al revés: se parte de sus pocas páginas y se copia
// solo lo que ellas cargan. Menos archivos publicados, menos superficie.
function construirSeguro() {
  const paginasSeguras = [['seguro-index.html', 'index.html'], ['404.html', '404.html'],
    ...SEGURAS.map((h) => ['herramientas/' + h + '.html', 'herramientas/' + h + '.html'])];
  for (const [origen, salida] of paginasSeguras) procesarArchivo(join(RAIZ, origen), origen, salida);
  procesarArchivo(join(RAIZ, 'enviar-descarga-sw.js'), 'enviar-descarga-sw.js');
  // Imágenes, tipos de letra y marca: completos (son chicos y los scripts los nombran por dentro).
  for (const carpeta of ['marca-logo', 'assets/img', 'assets/fonts']) if (existsSync(join(RAIZ, carpeta))) recorrer(join(RAIZ, carpeta));
  const paquetes = new Set();
  for (const ruta of recursos) {
    if (/^assets\/(js|css)\//.test(ruta)) {
      if (existsSync(join(RAIZ, ruta))) procesarArchivo(join(RAIZ, ruta), ruta);
    } else if (ruta.startsWith('vendor/')) {
      const m = /^vendor\/((?:@[^/]+\/)?[^/]+\/[^/]+)\//.exec(ruta);
      if (m) paquetes.add(m[1]);
    }
  }
  for (const paquete of paquetes) copiarVendor(join(RAIZ, 'vendor', paquete));
}

// Cabeceras del sitio seguro, una regla por página (Cloudflare las aplica por
// ruta, sin la extensión). Más estrictas que las de www: acá no hay nada más
// en el origen. La política de contenido también va dentro de la página del
// administrador (meta) para poder probarla sin Cloudflare.
function cabecerasSeguro(rutas) {
  // Lo común va en una sola regla /* (también cubre las respuestas 404 y el service worker). Cloudflare
  // une con coma los valores de una misma cabecera si coinciden varias reglas: por eso no se repite abajo.
  const lineas = ['# Generado por build.mjs (cabecerasSeguro): no editar.',
    '/*', '  X-Frame-Options: DENY', '  Referrer-Policy: no-referrer', '  X-Content-Type-Options: nosniff',
    '  Cross-Origin-Opener-Policy: same-origin', '  Strict-Transport-Security: max-age=31536000', '',
    '# Las librerías de /vendor/ llevan la versión en la ruta: se guardan un año.', '/vendor/*', '  Cache-Control: public, max-age=31536000, immutable', ''];
  for (const r of rutas.map((x) => x.replace(/\\/g, '/')).filter((x) => x !== '404.html' && x !== 'en/404.html')) {
    const ruta = r === 'index.html' ? '/' : r === 'en/index.html' ? '/en/' : '/' + r.replace(/\.html$/, '');
    const nombre = (/herramientas\/([^/]+)\.html$/.exec(r) || [])[1];
    const p = politicaSeguro(nombre, hashesPorPagina.get(r) || []);
    lineas.push(ruta, ...['Content-Security-Policy: ' + p.csp + "; frame-ancestors 'none'", 'Permissions-Policy: ' + p.permisos].concat(p.extra).map((l) => '  ' + l), '');
  }
  writeFileSync(join(SALIDA, '_headers'), lineas.join('\n'));
}

// Antes de publicar el sitio seguro: ningún enlace ni recurso local roto, y nada cargado de afuera.
function revisarSeguro() {
  const problemas = [];
  (function recorrerSalida(dir) {
    for (const n of readdirSync(dir)) {
      const p = join(dir, n);
      if (statSync(p).isDirectory()) { recorrerSalida(p); continue; }
      if (extname(n).toLowerCase() !== '.html') continue;
      const rel = relative(SALIDA, p).replace(/\\/g, '/');
      const html = readFileSync(p, 'utf8');
      for (const m of html.matchAll(/<(?:script|img|iframe|link)\b[^>]*\s(?:src|href)\s*=\s*"(https?:\/\/[^"]+)"[^>]*>/gi)) {
        if (!/rel="(canonical|alternate)"/i.test(m[0]) && !/<link\b[^>]*rel="icon"/i.test(m[0]) && /<(script|img|iframe)\b/i.test(m[0])) problemas.push(rel + ' carga de afuera: ' + m[1]);
        if (/<link\b[^>]*rel="stylesheet"/i.test(m[0])) problemas.push(rel + ' carga un estilo de afuera: ' + m[1]);
      }
      const base = rel === '404.html' || rel === 'en/404.html' ? '' : dirname(rel) === '.' ? '' : dirname(rel);
      for (const m of html.matchAll(/\s(?:href|src)\s*=\s*"([^"]+)"/gi)) {
        const url = m[1].split('#')[0].split('?')[0];
        if (!url || /^[a-z][a-z0-9+.-]*:/i.test(url) || url.startsWith('//')) continue;
        const pila = [];
        const partes = (url.startsWith('/') ? url.slice(1).split('/') : (base ? base.split('/') : []).concat(url.split('/')));
        for (const parte of partes) { if (parte === '.' || parte === '') continue; if (parte === '..') pila.pop(); else pila.push(parte); }
        const ruta = pila.join('/');
        const existe = existsSync(join(SALIDA, ruta)) && !statSync(join(SALIDA, ruta)).isDirectory()
          || existsSync(join(SALIDA, ruta + '.html')) || existsSync(join(SALIDA, ruta, 'index.html')) || ruta === '';
        if (!existe) problemas.push(rel + ' enlaza a ' + m[1] + ', que no existe en el sitio seguro');
      }
    }
  })(SALIDA);
  if (problemas.length) throw new Error('Sitio seguro con problemas:\n  - ' + [...new Set(problemas)].join('\n  - '));
}

// Rutas de todas las páginas publicadas, para armar el sitemap al final.
const paginas = [];
const paginasEn = [];
// Si catalogo.json no coincide con herramientas/, esto corta el build acá.
const catalogo = leerCatalogo(RAIZ);
const indiceCategorias = indiceDeCategorias(catalogo);
const comun = diccionarioComun(RAIZ);
const SITIO_SEGURO_HERRAMIENTAS = 'https://seguro.minicookies.net/herramientas/';
const SITIO_SEGURO_HERRAMIENTAS_EN = 'https://seguro.minicookies.net/en/herramientas/';
// Páginas de www con versión en inglés (para enlazarlas desde el sitio seguro).
const conInglesPrincipal = paginasConIngles(RAIZ, EXCLUIR_NOMBRES);
// En el sitio seguro: solo sus propias páginas tienen gemela.
const conIngles = ES_SEGURO
  ? new Set([['seguro-index.html', 'index.html'], ['404.html', '404.html'], ...SEGURAS.map((h) => ['herramientas/' + h + '.html', 'herramientas/' + h + '.html'])]
    .filter(([origen]) => tieneIngles(readFileSync(join(RAIZ, origen), 'utf8'))).map(([, salida]) => salida))
  : conInglesPrincipal;
const conteo = contarHerramientas(RAIZ, catalogo);
// Las tarjetas de las herramientas seguras apuntan al otro sitio.
const opcionesCatalogo = ES_SEGURO
  ? { seguras: SEGURAS }
  : { hrefDe: (h) => (SEGURAS.includes(h) ? urlSegura(h) : undefined) };

rmSync(SALIDA, { recursive: true, force: true });
mkdirSync(SALIDA);
if (ES_SEGURO) construirSeguro(); else recorrer(RAIZ);
generarRobotsYSitemap(SALIDA, paginas.concat(paginasEn));
if (ES_SEGURO) { cabecerasSeguro(paginas.concat(paginasEn)); revisarSeguro(); }
// Todo servidor ajeno que el sitio carga tiene que estar en la política de
// privacidad. Si no, el build falla y Cloudflare deja la versión anterior.
// (La política vive en www: el sitio seguro la lee del build principal.)
const terceros = revisarTerceros(SALIDA, RAIZ, ES_SEGURO ? { politicaEn: join(RAIZ, 'dist') } : {});
// Todo archivo de vendor/ que las páginas piden tiene que publicarse (y estar
// en git, que es de donde compila Cloudflare).
const rutasVendor = revisarReferenciasVendor(SALIDA, RAIZ);
// Las páginas con el motor de PowerPoint en <script> lo traen completo (motor-pptx.mjs).
if (!ES_SEGURO) revisarMotorPptx(RAIZ);
if (!ES_SEGURO) revisarDependenciasScripts(RAIZ);
console.log(`Build ${DESTINO} listo en ${CARPETA_SALIDA}/ — ${conteo.total} herramientas en el catálogo, ${paginas.length} páginas + ${paginasEn.length} en inglés bajo /en/, sitemap.xml y robots.txt generados; servidores externos declarados: ${[...terceros.keys()].sort().join(', ') || 'ninguno'}; ${rutasVendor} rutas de vendor/ comprobadas`);

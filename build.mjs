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
import { readdirSync, readFileSync, writeFileSync, mkdirSync, copyFileSync, rmSync, statSync } from 'node:fs';
import { inyectarSeo, generarRobotsYSitemap, indiceDeCategorias, esNoIndex } from './seo.mjs';
import { traducir, diccionarioComun, paginasConIngles } from './ingles.mjs';
import { contarHerramientas, aplicarConteos } from './conteo.mjs';
import { leerCatalogo, aplicarCatalogo } from './catalogo.mjs';
import { revisarTerceros } from './terceros.mjs';
import { revisarReferenciasVendor } from './referencias-vendor.mjs';
import { revisarMotorPptx } from './motor-pptx.mjs';
import { join, relative, extname, dirname } from 'node:path';

const RAIZ = process.cwd();
const SALIDA = join(RAIZ, 'dist');

// Lo que no se publica: herramientas del build, notas internas (.md) y
// archivos del repo.
const EXCLUIR_NOMBRES = new Set(['.git', '.github', '.claude', '.wrangler', 'node_modules', 'dist',
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
  if (ETIQUETA_ADSENSE && !html.includes('adsbygoogle.js') && !esNoIndex(html)) {
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
    const ext = extname(nombre).toLowerCase();
    if (EXCLUIR_EXTENSIONES.has(ext)) continue;
    const destino = join(SALIDA, rel);
    mkdirSync(dirname(destino), { recursive: true });
    if (ext === '.js' || ext === '.mjs') writeFileSync(destino, minificarJs(readFileSync(origen, 'utf8'), rel) + '\n');
    else if (ext === '.css') writeFileSync(destino, minificarCss(readFileSync(origen, 'utf8'), rel) + '\n');
    else if (ext === '.html') {
      const rutaWeb = rel.replace(/\\/g, '/');
      // Primero el catálogo (genera las tarjetas), después los conteos.
      const crudo = aplicarConteos(aplicarCatalogo(readFileSync(origen, 'utf8'), catalogo, rutaWeb), conteo, rutaWeb);
      // Al sitemap solo van las páginas que se indexan (no la 404).
      if (!esNoIndex(crudo)) paginas.push(rel);
      writeFileSync(destino, enlacesLimpios(inyectarSeo(procesarHtml(crudo, rel), rel, indiceCategorias, conIngles)));

      // Gemela en inglés bajo /en/, con su propia URL para que Google la vea.
      if (conIngles.has(rutaWeb)) {
        const en = traducir(crudo, RAIZ, comun, rutaWeb, conIngles);
        if (!en) throw new Error('No se pudo traducir ' + rutaWeb);
        const destinoEn = join(SALIDA, 'en', rel);
        mkdirSync(dirname(destinoEn), { recursive: true });
        writeFileSync(destinoEn, enlacesLimpios(inyectarSeo(procesarHtml(en, 'en/' + rel), rel, indiceCategorias, conIngles, true)));
        if (!esNoIndex(crudo)) paginasEn.push('en/' + rutaWeb);
      }
    }
    else copyFileSync(origen, destino);
  }
}

// Rutas de todas las páginas publicadas, para armar el sitemap al final.
const paginas = [];
const paginasEn = [];
// Si catalogo.json no coincide con herramientas/, esto corta el build acá.
const catalogo = leerCatalogo(RAIZ);
const indiceCategorias = indiceDeCategorias(catalogo);
const comun = diccionarioComun(RAIZ);
const conIngles = paginasConIngles(RAIZ, EXCLUIR_NOMBRES);
const conteo = contarHerramientas(RAIZ, catalogo);

rmSync(SALIDA, { recursive: true, force: true });
mkdirSync(SALIDA);
recorrer(RAIZ);
generarRobotsYSitemap(SALIDA, paginas.concat(paginasEn));
// Todo servidor ajeno que el sitio carga tiene que estar en la política de
// privacidad. Si no, el build falla y Cloudflare deja la versión anterior.
const terceros = revisarTerceros(SALIDA, RAIZ);
// Todo archivo de vendor/ que las páginas piden tiene que publicarse (y estar
// en git, que es de donde compila Cloudflare).
const rutasVendor = revisarReferenciasVendor(SALIDA, RAIZ);
// Las páginas con el motor de PowerPoint en <script> lo traen completo (motor-pptx.mjs).
revisarMotorPptx(RAIZ);
console.log(`Build listo en dist/ — ${conteo.total} herramientas, ${paginas.length} páginas + ${paginasEn.length} en inglés bajo /en/, sitemap.xml y robots.txt generados; servidores externos declarados: ${[...terceros.keys()].sort().join(', ')}; ${rutasVendor} rutas de vendor/ comprobadas`);

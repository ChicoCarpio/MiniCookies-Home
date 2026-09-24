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
import { join, relative, extname, dirname } from 'node:path';

const RAIZ = process.cwd();
const SALIDA = join(RAIZ, 'dist');

// Lo que no se publica: herramientas del build, notas internas (.md) y
// archivos del repo.
const EXCLUIR_NOMBRES = new Set(['.git', '.github', '.claude', '.wrangler', 'node_modules', 'dist',
  'build.mjs', 'package.json', 'package-lock.json', '.gitignore', '.DS_Store', 'Thumbs.db']);
const EXCLUIR_EXTENSIONES = new Set(['.md', '.log']);

// Aviso de copyright de cada sitio: campo "config.aviso" de su package.json.
const AVISO = `/*! ${JSON.parse(readFileSync(join(RAIZ, 'package.json'), 'utf8')).config.aviso} */`;
// Scripts dentro del HTML a partir de este tamaño reciben el aviso de
// copyright (los chiquitos, como el del tema, no lo necesitan).
const MINIMO_PARA_AVISO = 1500;

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
  return html;
}

function recorrer(carpeta) {
  for (const nombre of readdirSync(carpeta)) {
    if (EXCLUIR_NOMBRES.has(nombre)) continue;
    const origen = join(carpeta, nombre);
    const rel = relative(RAIZ, origen);
    if (statSync(origen).isDirectory()) { recorrer(origen); continue; }
    const ext = extname(nombre).toLowerCase();
    if (EXCLUIR_EXTENSIONES.has(ext)) continue;
    const destino = join(SALIDA, rel);
    mkdirSync(dirname(destino), { recursive: true });
    if (ext === '.js' || ext === '.mjs') writeFileSync(destino, minificarJs(readFileSync(origen, 'utf8'), rel) + '\n');
    else if (ext === '.css') writeFileSync(destino, minificarCss(readFileSync(origen, 'utf8'), rel) + '\n');
    else if (ext === '.html') writeFileSync(destino, procesarHtml(readFileSync(origen, 'utf8'), rel));
    else copyFileSync(origen, destino);
  }
}

rmSync(SALIDA, { recursive: true, force: true });
mkdirSync(SALIDA);
recorrer(RAIZ);
console.log('Build listo en dist/');

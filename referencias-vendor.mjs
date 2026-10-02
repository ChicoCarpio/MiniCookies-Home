/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. */
// Todo archivo de vendor/ que el sitio pide tiene que publicarse. Esto lo
// comprueba en cada build.
//
// Por qué existe: el .gitignore tenía "dist/" (pensado para la salida del
// build) y eso excluía también toda carpeta dist DENTRO de vendor/. Siete
// librerías (OCR, dictado, stickers, proteger PDF, PowerPoint…) estaban en el
// disco pero nunca llegaron a git: las pruebas locales pasaban y en
// producción esas herramientas no cargaban (encontrado el 2026-10-02).
//
// Dos comprobaciones sobre lo que el build publica (dist/, sin dist/vendor):
// 1. Cada ruta "vendor/…" que aparece en el HTML o el JavaScript existe en
//    dist/vendor. En Cloudflare (que compila desde git) esto solo ya habría
//    cortado el build.
// 2. Si hay git, además git no la ignora: así falla también en la computadora
//    antes del push, donde el archivo sí existe aunque nunca vaya a subirse.
//    (Un archivo nuevo que todavía no se agregó con git add no cuenta: eso es
//    el flujo normal; lo que se perdió en silencio fue lo ignorado.)
//
// Una ruta armada en el JavaScript con variables no se ve: escribirlas
// completas (ver assets/js/pdfjs.js).

import { readdirSync, readFileSync, statSync, existsSync } from 'node:fs';
import { join, extname, relative } from 'node:path';
import { execFileSync } from 'node:child_process';

const EXTENSIONES = new Set(['.html', '.js', '.mjs', '.css']);
// vendor/<ruta> en src, import, url(), href, o rutaVendor('<ruta>').
const PATRONES = [/vendor\/([\w@.+\-\/]+)/g, /rutaVendor\(\s*['"]([^'"]+)['"]\s*\)/g];

function archivos(dir, acc = []) {
  for (const n of readdirSync(dir)) {
    const p = join(dir, n);
    if (statSync(p).isDirectory()) archivos(p, acc);
    else if (EXTENSIONES.has(extname(n).toLowerCase())) acc.push(p);
  }
  return acc;
}

function ignoradosPorGit(raiz) {
  try {
    const salida = execFileSync('git', ['ls-files', '-z', '--others', '--ignored', '--exclude-standard', 'vendor'], { cwd: raiz, encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] });
    return salida.split('\0').filter(Boolean).map((r) => r.slice('vendor/'.length));
  } catch {
    return null; // sin git (o fuera de un repositorio): solo la comprobación 1
  }
}

export function revisarReferenciasVendor(salida, raiz) {
  const vendorPublicado = join(salida, 'vendor');
  const referencias = new Map(); // ruta → primer archivo que la usa
  for (const archivo of archivos(salida)) {
    const rel = relative(salida, archivo).replace(/\\/g, '/');
    if (rel.startsWith('vendor/')) continue;
    const texto = readFileSync(archivo, 'utf8');
    for (const patron of PATRONES) {
      for (const m of texto.matchAll(patron)) {
        // Sin el ?v=… de las rutas que tuvieron que cambiar de URL (ver _headers).
        const ruta = m[1].replace(/\?.*$/, '').replace(/[.\/]+$/, (f) => (f.endsWith('/') ? '/' : ''));
        if (!referencias.has(ruta)) referencias.set(ruta, rel);
      }
    }
  }

  const ignorados = ignoradosPorGit(raiz);
  const errores = [];
  for (const [ruta, usadaEn] of referencias) {
    const destino = join(vendorPublicado, ruta);
    if (!existsSync(destino)) {
      errores.push(`vendor/${ruta} (la pide ${usadaEn}) no existe`);
      continue;
    }
    if (!ignorados) continue;
    const prefijo = ruta.replace(/\/$/, '') + '/';
    const ignorado = ignorados.find((r) => r === ruta || r.startsWith(prefijo));
    if (ignorado) errores.push(`vendor/${ignorado} (la pide ${usadaEn}) está ignorado por git: no se va a publicar (revisar .gitignore)`);
  }

  if (errores.length) {
    console.error('\nArchivos de vendor/ que el sitio pide y no se publicarían:\n  ' + errores.join('\n  '));
    console.error('\nEl build se detiene para que Cloudflare no publique herramientas rotas.\n');
    process.exit(1);
  }
  return referencias.size;
}

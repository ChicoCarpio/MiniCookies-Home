/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. */
// Cada servidor ajeno del que el sitio carga algo tiene que estar nombrado en
// la política de privacidad. Esto lo comprueba en cada build.
//
// Por qué existe: la política es una promesa a quien usa el sitio, y se había
// quedado atrás sin que nadie lo notara. Dictado por voz y el generador de
// stickers bajaban sus modelos de huggingface.co y la política nunca lo
// mencionó (encontrado el 2026-10-01). Un documento no se actualiza solo; esto
// corta el build si aparece un dominio nuevo sin declarar.
//
// Qué cuenta como "cargar": lo que el navegador pide solo al abrir o usar la
// página — <script src>, hojas de estilo, imágenes, iframes, formularios, y en
// el JavaScript fetch/import/Worker/.src. Un <a href> a GitHub NO cuenta (no
// carga nada hasta que alguien hace clic), ni los identificadores de formatos
// XML (www.w3.org, schemas.openxmlformats.org…), que son nombres, no
// direcciones que se visiten.
//
// Lo que una librería de vendor/ pide por dentro no se puede ver leyendo el
// código del sitio: va declarado a mano abajo con `interno: true`. Tampoco lo
// que agrega Cloudflare al servir la página: va con `inyectado`.

import { readdirSync, readFileSync, statSync, existsSync } from 'node:fs';
import { join, extname, relative } from 'node:path';

const PROPIO = /(^|\.)minicookies\.net$/;

// Dominio → qué texto tiene que aparecer en la política (es y en) y por qué.
export const TERCEROS = {
  'pagead2.googlesyndication.com': { menciona: 'AdSense', motivo: 'etiqueta de Google AdSense (la inyecta build.mjs)' },
  'api.web3forms.com': { menciona: 'Web3Forms', motivo: 'envío del formulario de contacto' },
  'huggingface.co': { menciona: 'huggingface.co', motivo: 'modelos de IA que baja transformers.js (dictado, stickers, máscaras de primer plano/fondo y "Elegir objeto" del Editor avanzado)', interno: 'vendor/@huggingface/transformers' },
  'cdn.jsdelivr.net': { menciona: 'cdn.jsdelivr.net', motivo: 'onnxruntime-web, que baja transformers.js', interno: 'vendor/@huggingface/transformers' },
  // No está en el código: Cloudflare inyecta el script de Web Analytics al
  // servir cada página (decisión del usuario, 2026-10-02: se conserva).
  'static.cloudflareinsights.com': { menciona: 'static.cloudflareinsights.com', motivo: 'Cloudflare Web Analytics, que inyecta Cloudflare en cada página', inyectado: 'Cloudflare' },
};

const host = (url) => {
  const m = /^(?:https?:)?\/\/([a-z0-9.-]+\.[a-z]{2,})/i.exec(url);
  return m ? m[1].toLowerCase() : null;
};

function archivos(dir, acc = []) {
  for (const n of readdirSync(dir)) {
    const r = join(dir, n);
    if (statSync(r).isDirectory()) { if (n !== 'vendor') archivos(r, acc); } else acc.push(r);
  }
  return acc;
}

/** Direcciones externas que cargan las páginas y el código propio de `dist`. */
export function cargasExternas(dist) {
  const vistos = new Map(); // host → Set de archivos
  const anotar = (url, archivo) => {
    const h = host(url);
    if (!h || PROPIO.test(h)) return;
    if (!vistos.has(h)) vistos.set(h, new Set());
    vistos.get(h).add(relative(dist, archivo).replace(/\\/g, '/'));
  };
  const enJs = (codigo, archivo) => {
    const re = /(?:\b(?:fetch|import|importScripts|Worker|SharedWorker|sendBeacon|WebSocket|EventSource)\s*\(\s*|\bfrom\s*|\.(?:src|href|action)\s*=\s*)["'`]((?:https?:)?\/\/[^"'`]+)/g;
    for (const m of codigo.matchAll(re)) anotar(m[1], archivo);
  };
  const enCss = (codigo, archivo) => {
    for (const m of codigo.matchAll(/(?:url\(\s*["']?|@import\s+["'])((?:https?:)?\/\/[^"')\s]+)/g)) anotar(m[1], archivo);
  };

  for (const archivo of archivos(dist)) {
    const ext = extname(archivo);
    if (!['.html', '.js', '.mjs', '.css'].includes(ext)) continue;
    const s = readFileSync(archivo, 'utf8');
    if (ext === '.css') { enCss(s, archivo); continue; }
    if (ext !== '.html') { enJs(s, archivo); continue; }

    for (const [, etiqueta, attrs] of s.matchAll(/<(script|link|img|iframe|source|video|audio|embed|object|form)\b([^>]*)>/gi)) {
      if (etiqueta.toLowerCase() === 'link' && !/\brel="?(?:stylesheet|preload|modulepreload|icon|manifest|prefetch)/i.test(attrs)) continue;
      for (const [, valor] of attrs.matchAll(/\b(?:src|href|action|data|poster)="?([^"\s>]+)/gi)) anotar(valor, archivo);
    }
    for (const [, codigo] of s.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)) enJs(codigo, archivo);
    for (const [, codigo] of s.matchAll(/<style\b[^>]*>([\s\S]*?)<\/style>/gi)) enCss(codigo, archivo);
    for (const [, codigo] of s.matchAll(/\bstyle="([^"]*)"/gi)) enCss(codigo, archivo);
  }
  return vistos;
}

/** Corta el build si un dominio cargado no está declarado o la política no lo nombra. */
export function revisarTerceros(dist, raiz) {
  const errores = [];
  const avisos = [];
  const vistos = cargasExternas(dist);

  for (const [h, donde] of vistos) {
    if (!TERCEROS[h]) {
      errores.push(`${h} se carga desde ${[...donde].slice(0, 3).join(', ')}${donde.size > 3 ? '…' : ''}, pero no está declarado. `
        + 'Nombralo en privacidad.html (es y en) y agregalo a TERCEROS en terceros.mjs.');
    }
  }

  const politica = ['privacidad.html', 'en/privacidad.html'].map((p) => [p, existsSync(join(dist, p)) ? readFileSync(join(dist, p), 'utf8') : '']);
  for (const [h, { menciona, interno, inyectado }] of Object.entries(TERCEROS)) {
    for (const [p, texto] of politica) {
      if (!texto.includes(menciona)) errores.push(`${p} no menciona "${menciona}" (por ${h})`);
    }
    if (interno && !existsSync(join(raiz, interno))) avisos.push(`${h} figura como carga de ${interno}, que ya no existe: ¿sigue haciendo falta declararlo?`);
    if (!interno && !inyectado && !vistos.has(h)) avisos.push(`${h} está declarado pero ninguna página lo carga: ¿la política habla de algo que ya no se usa?`);
  }

  for (const a of avisos) console.warn('Aviso (terceros): ' + a);
  if (errores.length) throw new Error('Servidores externos sin declarar:\n  - ' + errores.join('\n  - '));
  return vistos;
}

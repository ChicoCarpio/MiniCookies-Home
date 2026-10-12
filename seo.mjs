/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. */
// SEO del build: lo que hace que Google y los buscadores con IA puedan
// descubrir y entender las páginas del sitio.
//
// Nada de esto se escribe a mano en los HTML. Se genera acá a partir de lo
// que cada página ya declara (su <title>, su <meta description> y su ruta),
// para que no pueda quedar desactualizado respecto al sitio real: si se
// agrega una herramienta, aparece sola en el sitemap y recibe sus etiquetas.
//
// Se agrega por página:
//   - <link rel="canonical">      una sola URL buena por página
//   - Open Graph y Twitter Card   cómo se ve al compartirla en un chat o red
//   - JSON-LD                     qué ES la página, en datos y no en prosa:
//                                 SoftwareApplication para cada herramienta
//                                 (esto es lo que permite que un buscador o
//                                 un asistente entienda "esta página une
//                                 PDFs" en vez de solo indexar palabras)
//   - BreadcrumbList              dónde está dentro del sitio
// Y para el sitio entero: robots.txt y sitemap.xml.

import { readFileSync, writeFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { categoriaPrincipal } from './catalogo.mjs';

import { SITIO_ACTUAL, ES_SEGURO } from './destinos.mjs';

// El sitio que se está construyendo (www o seguro): ver destinos.mjs.
export const SITIO = SITIO_ACTUAL;
const IMAGEN_SOCIAL = SITIO + '/marca-logo/minicookies_expandido.png';

// Las 9 páginas de categoría, para las migas de pan de cada herramienta.
const CATEGORIAS = {
  pdf: 'PDF', word: 'Word', 'excel-csv': 'Excel / CSV', powerpoint: 'PowerPoint',
  epub: 'EPUB', markdown: 'Markdown', texto: 'Texto plano', imagenes: 'Imágenes',
  generales: 'Generales',
};

const escXml = (s) => String(s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;')
  .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');

// El texto de un atributo o etiqueta ya viene escapado en el HTML; para
// meterlo en JSON-LD hay que devolverlo a texto plano primero.
const desescapar = (s) => String(s)
  .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&apos;/g, "'")
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');

/**
 * Ruta de un archivo dentro del sitio -> la URL con la que se sirve.
 *
 * Cloudflare Pages entrega las páginas SIN la extensión y manda un 308 de
 * `/pdf.html` a `/pdf`. Así que la dirección buena —la que va en el canonical,
 * en hreflang, en Open Graph y en el sitemap— es la corta: declarar la otra
 * sería darle a Google como canónica una URL que redirige a otra.
 */
export function urlDe(rel) {
  const limpio = rel.replace(/\\/g, '/').replace(/\.html$/i, '');
  if (limpio === 'index') return SITIO + '/';
  // Una carpeta se sirve por su índice: /en/index -> /en/
  return SITIO + '/' + limpio.replace(/\/index$/, '/');
}

/**
 * Para cada herramienta, la categoría de sus migas de pan, sacada de
 * catalogo.json. Varias herramientas son bidireccionales y están en dos
 * categorías (ej. Word ↔ PDF): se usa la primera según el orden de CATEGORIAS,
 * que alcanza para ubicarla.
 */
export function indiceDeCategorias(catalogo) {
  return categoriaPrincipal(catalogo, Object.keys(CATEGORIAS));
}

const titulos = (html) => ({
  titulo: desescapar((html.match(/<title[^>]*>([\s\S]*?)<\/title>/i) || [, ''])[1].trim()),
  desc: desescapar((html.match(/<meta\s+name="description"[^>]*\scontent="([^"]*)"/i) || [, ''])[1].trim()),
});

/** La página pide no ser indexada: <meta name="robots" content="noindex">. */
export const esNoIndex = (html) => /<meta\s+name="robots"\s+content="[^"]*noindex/i.test(html);

/**
 * Agrega a una página sus etiquetas de SEO justo antes de </head>.
 * Si la página ya trae alguna escrita a mano, se respeta y no se duplica.
 */
export function inyectarSeo(html, rel, indice, conIngles, esIngles = false) {
  if (/<link[^>]+rel="canonical"/i.test(html)) return html;
  // Una página que no se indexa (la 404) no tiene URL propia que declarar.
  if (esNoIndex(html)) return html;

  const ruta = rel.replace(/\\/g, '/');
  const url = urlDe(esIngles ? 'en/' + ruta : ruta);
  const { titulo, desc } = titulos(html);
  if (!titulo) return html;

  const esHerramienta = ruta.startsWith('herramientas/');
  const archivo = ruta.replace(/^herramientas\//, '');
  const esCategoria = !esHerramienta && Object.prototype.hasOwnProperty.call(CATEGORIAS, archivo.replace(/\.html$/, ''));

  const migas = [{ name: 'Inicio', item: SITIO + '/' }];
  if (esHerramienta) {
    const cat = indice.get(archivo);
    // El sitio seguro no tiene páginas de categoría.
    if (cat && !ES_SEGURO) migas.push({ name: CATEGORIAS[cat], item: `${SITIO}/${cat}.html` });
    migas.push({ name: titulo.split('—')[0].split('|')[0].trim(), item: url });
  } else if (ruta !== 'index.html') {
    migas.push({ name: titulo.split('—')[0].split('|')[0].trim(), item: url });
  }

  const datos = [];

  if (esHerramienta) {
    // El tipo que le dice a un buscador que esto es una herramienta que se
    // usa en el navegador, gratis, y qué hace — no un artículo.
    datos.push({
      '@context': 'https://schema.org',
      '@type': 'SoftwareApplication',
      name: titulo.split('—')[0].split('|')[0].trim(),
      description: desc,
      url,
      applicationCategory: 'UtilitiesApplication',
      operatingSystem: 'Cualquiera con navegador web',
      browserRequirements: 'Requiere JavaScript',
      inLanguage: esIngles ? 'en' : 'es',
      isAccessibleForFree: true,
      offers: { '@type': 'Offer', price: '0', priceCurrency: 'USD' },
      publisher: { '@type': 'Organization', name: 'Mini Cookies', url: SITIO + '/' },
    });
  } else if (ruta === 'index.html') {
    datos.push({
      '@context': 'https://schema.org',
      '@type': 'WebSite',
      name: 'Mini Cookies',
      url: SITIO + '/',
      description: desc,
      inLanguage: esIngles ? 'en' : 'es',
      publisher: { '@type': 'Organization', name: 'Mini Cookies', url: SITIO + '/', logo: IMAGEN_SOCIAL },
    });
  } else if (esCategoria) {
    datos.push({
      '@context': 'https://schema.org',
      '@type': 'CollectionPage',
      name: titulo.split('—')[0].split('|')[0].trim(),
      description: desc,
      url,
      inLanguage: esIngles ? 'en' : 'es',
      isPartOf: { '@type': 'WebSite', name: 'Mini Cookies', url: SITIO + '/' },
    });
  }

  if (migas.length > 1) {
    datos.push({
      '@context': 'https://schema.org',
      '@type': 'BreadcrumbList',
      itemListElement: migas.map((m, i) => ({
        '@type': 'ListItem', position: i + 1, name: m.name, item: m.item,
      })),
    });
  }

  // hreflang solo si la página existe de verdad en los dos idiomas. Declarar
  // una alternativa que no está traducida es peor que no declarar ninguna.
  const alternativas = (conIngles && conIngles.has(ruta)) ? [
    `<link rel="alternate" hreflang="es" href="${escXml(urlDe(ruta))}">`,
    `<link rel="alternate" hreflang="en" href="${escXml(urlDe('en/' + ruta))}">`,
    `<link rel="alternate" hreflang="x-default" href="${escXml(urlDe(ruta))}">`,
  ] : [];

  const etiquetas = [
    `<link rel="canonical" href="${escXml(url)}">`,
    ...alternativas,
    `<meta property="og:type" content="website">`,
    `<meta property="og:site_name" content="Mini Cookies">`,
    `<meta property="og:locale" content="${esIngles ? 'en_US' : 'es_ES'}">`,
    `<meta property="og:url" content="${escXml(url)}">`,
    `<meta property="og:title" content="${escXml(titulo)}">`,
    `<meta property="og:description" content="${escXml(desc)}">`,
    `<meta property="og:image" content="${escXml(IMAGEN_SOCIAL)}">`,
    `<meta name="twitter:card" content="summary_large_image">`,
    `<meta name="twitter:title" content="${escXml(titulo)}">`,
    `<meta name="twitter:description" content="${escXml(desc)}">`,
    `<meta name="twitter:image" content="${escXml(IMAGEN_SOCIAL)}">`,
    ...datos.map((d) => `<script type="application/ld+json">${JSON.stringify(d)}</script>`),
  ].join('\n');

  return html.replace(/<\/head>/i, etiquetas + '\n</head>');
}

/**
 * robots.txt y sitemap.xml. Sin esto el sitio depende de que alguien enlace
 * cada página para que un buscador llegue a ella.
 */
export function generarRobotsYSitemap(salida, rutas) {
  const hoy = new Date().toISOString().slice(0, 10);
  const prioridad = (r) => (r === 'index.html' ? '1.0' : r.includes('/') ? '0.7' : '0.8');

  const urls = rutas
    .map((r) => r.replace(/\\/g, '/'))
    .sort((a, b) => (a === 'index.html' ? -1 : b === 'index.html' ? 1 : a.localeCompare(b)))
    .map((r) => `  <url>
    <loc>${escXml(urlDe(r))}</loc>
    <lastmod>${hoy}</lastmod>
    <priority>${prioridad(r)}</priority>
  </url>`)
    .join('\n');

  writeFileSync(join(salida, 'sitemap.xml'),
    `<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n${urls}\n</urlset>\n`);

  // El "no al entrenamiento de IA" va acá y no en Cloudflare: su bloqueo de
  // "AI training" (2026-10-01) también le daba 403 a Googlebot, Bingbot y
  // Applebot, que alimentan IA además de la búsqueda. Estos nombres son solo
  // de entrenamiento; los buscadores y los lectores de IA siguen entrando.
  const soloEntrenamiento = ['GPTBot', 'ClaudeBot', 'CCBot', 'Bytespider',
    'Google-Extended', 'Applebot-Extended', 'meta-externalagent'];
  writeFileSync(join(salida, 'robots.txt'),
    `# Mini Cookies — todo el sitio es público y se puede indexar.\n` +
    `User-agent: *\nAllow: /\n\n` +
    `# No a usar el contenido para entrenar modelos de IA.\n` +
    soloEntrenamiento.map((ua) => `User-agent: ${ua}\n`).join('') +
    `Disallow: /\n\nSitemap: ${SITIO}/sitemap.xml\n`);

  return rutas.length;
}

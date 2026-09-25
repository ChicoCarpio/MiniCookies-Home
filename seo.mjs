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

export const SITIO = 'https://www.minicookies.net';
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

/** Ruta relativa dentro del sitio -> URL absoluta y canónica. */
export function urlDe(rel) {
  const limpio = rel.replace(/\\/g, '/');
  if (limpio === 'index.html') return SITIO + '/';
  return SITIO + '/' + limpio;
}

/**
 * Recorre las páginas de categoría y devuelve, para cada herramienta, a qué
 * categorías pertenece. Varias herramientas son bidireccionales y tienen
 * tarjeta en dos categorías (ej. Word ↔ PDF): para las migas de pan se usa
 * la primera, que alcanza para ubicarla.
 */
export function indiceDeCategorias(raiz) {
  const indice = new Map();
  for (const clave of Object.keys(CATEGORIAS)) {
    let html;
    try { html = readFileSync(join(raiz, clave + '.html'), 'utf8'); } catch { continue; }
    const re = /<a class="tarjeta-herramienta" href="(?:\.\.\/)?herramientas\/([^"]+)"/g;
    let m;
    while ((m = re.exec(html))) {
      if (!indice.has(m[1])) indice.set(m[1], clave);
    }
  }
  return indice;
}

const titulos = (html) => ({
  titulo: desescapar((html.match(/<title[^>]*>([\s\S]*?)<\/title>/i) || [, ''])[1].trim()),
  desc: desescapar((html.match(/<meta\s+name="description"[^>]*\scontent="([^"]*)"/i) || [, ''])[1].trim()),
});

/**
 * Agrega a una página sus etiquetas de SEO justo antes de </head>.
 * Si la página ya trae alguna escrita a mano, se respeta y no se duplica.
 */
export function inyectarSeo(html, rel, indice, conIngles, esIngles = false) {
  if (/<link[^>]+rel="canonical"/i.test(html)) return html;

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
    if (cat) migas.push({ name: CATEGORIAS[cat], item: `${SITIO}/${cat}.html` });
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

  writeFileSync(join(salida, 'robots.txt'),
    `# Mini Cookies — todo el sitio es público y se puede indexar.\n` +
    `User-agent: *\nAllow: /\n\nSitemap: ${SITIO}/sitemap.xml\n`);

  return rutas.length;
}

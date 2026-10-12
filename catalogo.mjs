/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. */
// El catálogo de herramientas: UNA sola fuente para todo lo que lista
// herramientas en el sitio.
//
// Por qué existe: cada herramienta (icono, título y descripción, en español y
// en inglés) estaba escrita a mano en tres lugares —su página de categoría,
// todas.html y la lista del buscador de la portada— y las tres copias se
// habían separado: el buscador no encontraba el generador de stickers,
// todas.html tenía otro orden y otros iconos que las categorías, y el lector de
// códigos tenía dos descripciones distintas. Ahora se escribe una vez, en
// catalogo.json, y el build genera las tres cosas. Agregar una herramienta es
// agregarla ahí.
//
// Cómo se usa desde una página (las dos son plantillas: abiertas directo del
// disco no muestran las tarjetas; para verlas, `npm run preview`):
//   <div class="grid-herramientas" data-catalogo="pdf"></div>   → tarjetas de esa categoría
//   var HERRAMIENTAS = {{CATALOGO_BUSCADOR}};                   → lista del buscador
//
// Los textos de las tarjetas viajan en window.i18nCatalogo (que idioma.js e
// ingles.mjs mezclan con el diccionario de la página), con las mismas claves de
// siempre: tool-<archivo>-titulo y tool-<archivo>-desc.

import { readFileSync, readdirSync, existsSync } from 'node:fs';
import { join } from 'node:path';

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/**
 * Lee catalogo.json y lo valida contra el disco. Cualquier inconsistencia
 * corta el build: es mejor no publicar que publicar un buscador que no
 * encuentra una herramienta.
 */
export function leerCatalogo(raiz) {
  const catalogo = JSON.parse(readFileSync(join(raiz, 'catalogo.json'), 'utf8'));
  const { categorias, herramientas } = catalogo;
  const errores = [];

  const enDisco = new Set(readdirSync(join(raiz, 'herramientas'))
    .filter((n) => n.toLowerCase().endsWith('.html')).map((n) => n.slice(0, -5)));
  for (const h of enDisco) if (!herramientas[h]) errores.push(`herramientas/${h}.html no está en catalogo.json`);
  for (const h of Object.keys(herramientas)) {
    if (!enDisco.has(h)) errores.push(`catalogo.json lista "${h}", pero no existe herramientas/${h}.html`);
    const t = herramientas[h];
    for (const idioma of ['es', 'en']) {
      if (!Array.isArray(t[idioma]) || t[idioma].length !== 2 || t[idioma].some((x) => !x || !String(x).trim())) {
        errores.push(`"${h}": falta título o descripción en "${idioma}"`);
      }
    }
  }

  const conCategoria = new Set();
  for (const [categoria, lista] of Object.entries(categorias)) {
    if (!existsSync(join(raiz, categoria + '.html'))) errores.push(`la categoría "${categoria}" no tiene ${categoria}.html`);
    for (const [h, icono] of Object.entries(lista)) {
      if (!herramientas[h]) errores.push(`la categoría "${categoria}" lista "${h}", que no está en "herramientas"`);
      if (!icono) errores.push(`"${h}" en "${categoria}" no tiene icono`);
      conCategoria.add(h);
    }
  }
  for (const h of Object.keys(herramientas)) {
    if (!conCategoria.has(h)) errores.push(`"${h}" no aparece en ninguna categoría: nadie podría llegar a ella desde un menú`);
  }

  // Subgrupos: opcionales por categoría; si hay, cada herramienta de la categoría va en uno solo.
  for (const [categoria, lista] of Object.entries(catalogo.grupos || {})) {
    if (!categorias[categoria]) { errores.push(`"grupos" habla de "${categoria}", que no es una categoría`); continue; }
    const vistas = new Set();
    for (const g of lista) {
      if (!g.id || !g.es || !g.en) errores.push(`un grupo de "${categoria}" no tiene id, es o en`);
      for (const h of g.herramientas || []) {
        if (!categorias[categoria][h]) errores.push(`el grupo "${g.id}" de "${categoria}" lista "${h}", que no está en esa categoría`);
        if (vistas.has(h)) errores.push(`"${h}" está en dos grupos de "${categoria}"`);
        vistas.add(h);
      }
    }
    for (const h of Object.keys(categorias[categoria])) if (!vistas.has(h)) errores.push(`"${h}" no está en ningún grupo de "${categoria}"`);
  }

  if (errores.length) throw new Error('catalogo.json:\n  - ' + errores.join('\n  - '));
  return catalogo;
}

/** Herramientas distintas por categoría (las bidireccionales cuentan en las dos). */
export function conteoPorCategoria(catalogo) {
  return Object.fromEntries(Object.entries(catalogo.categorias).map(([c, lista]) => [c, Object.keys(lista).length]));
}

/** Primera categoría de cada herramienta, en el orden de `orden` (para las migas de pan). */
export function categoriaPrincipal(catalogo, orden) {
  const indice = new Map();
  for (const c of orden) {
    for (const h of Object.keys(catalogo.categorias[c] || {})) {
      if (!indice.has(h + '.html')) indice.set(h + '.html', c);
    }
  }
  return indice;
}

function tarjeta(h, icono, t, href) {
  return `      <a class="tarjeta-herramienta" href="${href || `herramientas/${h}.html`}">
        <div class="icono">${icono}</div>
        <h3 data-i18n="tool-${h}-titulo">${esc(t.es[0])}</h3>
        <p data-i18n="tool-${h}-desc">${esc(t.es[1])}</p>
      </a>`;
}

/** Resuelve los marcadores del catálogo en una página. */
// opciones.hrefDe(h): dirección de la tarjeta de una herramienta (para las que viven en otro sitio).
// opciones.seguras: lista de herramientas del sitio seguro; habilita data-catalogo="seguras".
export function aplicarCatalogo(html, catalogo, archivo, opciones = {}) {
  const usadas = new Set();
  const gruposUsados = [];

  html = html.replace(/(<div class="grid-herramientas" data-catalogo="([a-z0-9-]+)">)\s*(<\/div>)/g,
    (todo, apertura, categoria, cierre) => {
      const lista = categoria === 'seguras' && opciones.seguras
        ? Object.fromEntries(opciones.seguras.map((h) => [h, catalogo.categorias.generales[h]]))
        : catalogo.categorias[categoria];
      if (!lista) throw new Error(`${archivo}: data-catalogo="${categoria}" no es una categoría de catalogo.json`);
      const tarjetas = (nombres) => nombres.map((h) => {
        usadas.add(h);
        return tarjeta(h, lista[h], catalogo.herramientas[h], opciones.hrefDe && opciones.hrefDe(h));
      });
      // Con subgrupos (h2 en la página de cada categoría; h3 en todas.html, donde h2 es la categoría):
      // un encabezado y su cuadrícula por grupo.
      const grupos = categoria !== 'seguras' && (catalogo.grupos || {})[categoria];
      const nivel = archivo === 'todas.html' ? 3 : 2;
      if (grupos) {
        return grupos.map((g) => {
          gruposUsados.push([categoria, g]);
          return `<div class="grupo-herramientas">
      <h${nivel} class="titulo-grupo" data-i18n="grupo-${categoria}-${g.id}">${esc(g.es)}</h${nivel}>
      ${apertura}
${tarjetas(g.herramientas).join('\n')}
      ${cierre}
    </div>`;
        }).join('\n    ');
      }
      return `${apertura}\n${tarjetas(Object.keys(lista)).join('\n')}\n    ${cierre}`;
    });

  if (html.includes('{{CATALOGO_BUSCADOR}}')) {
    // Una entrada por herramienta, con el icono de su primera categoría.
    const vistas = new Map();
    for (const lista of Object.values(catalogo.categorias)) {
      for (const [h, icono] of Object.entries(lista)) if (!vistas.has(h)) vistas.set(h, icono);
    }
    const lista = [...vistas].map(([h, icono]) => ({
      h: h + '.html', i: icono, u: (opciones.hrefDe && opciones.hrefDe(h)) || undefined, es: catalogo.herramientas[h].es, en: catalogo.herramientas[h].en,
    }));
    html = html.replace('{{CATALOGO_BUSCADOR}}', () => JSON.stringify(lista));
  }

  if (usadas.size) {
    const dic = { es: {}, en: {} };
    for (const [categoria, g] of gruposUsados) {
      for (const idioma of ['es', 'en']) dic[idioma][`grupo-${categoria}-${g.id}`] = g[idioma];
    }
    for (const h of usadas) {
      for (const idioma of ['es', 'en']) {
        dic[idioma][`tool-${h}-titulo`] = catalogo.herramientas[h][idioma][0];
        dic[idioma][`tool-${h}-desc`] = catalogo.herramientas[h][idioma][1];
      }
    }
    // El "\n};" final es lo que buscan idioma.js / ingles.mjs para leer el objeto.
    const script = `<script>window.i18nCatalogo = ${JSON.stringify(dic, null, 1).replace(/<\//g, '<\\/')};</script>\n`;
    if (!/<\/head>/i.test(html)) throw new Error(`${archivo}: usa el catálogo pero no tiene </head>`);
    html = html.replace(/<\/head>/i, () => script + '</head>');
  }
  return html;
}

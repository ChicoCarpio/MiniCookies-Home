/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Versión en inglés del sitio, generada en el build bajo /en/.
//
// En el navegador el idioma lo cambia assets/js/idioma.js, que reemplaza el
// texto de los elementos con data-i18n. Eso le sirve a una persona, pero no
// a un buscador: Google indexa una sola URL por página y vería siempre el
// español. Para que el inglés exista para Google necesita URL propia.
//
// Así que acá se hace lo mismo que hace idioma.js, pero antes de publicar:
// se aplica el diccionario 'en' sobre el HTML y se guarda el resultado en
// /en/…, de modo que esa URL entrega inglés de verdad, sin depender de que
// el crawler ejecute JavaScript (los de los asistentes de IA normalmente no
// lo hacen).
//
// Solo se traducen las páginas que declaran window.i18nPagina con textos en
// inglés. Las que no, no tienen versión /en/ y tampoco reciben hreflang:
// es preferible no tener alternativa a declararle a Google una que no está
// traducida de verdad.

import { readFileSync, readdirSync, statSync } from 'node:fs';
import { join, relative } from 'node:path';

const escHtml = (s) => String(s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const escAttr = (s) => escHtml(s).replace(/"/g, '&quot;');

/** Evalúa un objeto literal declarado en un <script> del HTML o de un .js. */
function leerObjeto(texto, nombre) {
  const m = texto.match(new RegExp('window\\.' + nombre + '\\s*=\\s*(\\{[\\s\\S]*?\\n\\};)'));
  if (!m) return null;
  try { return new Function('return ' + m[1].replace(/;\s*$/, ''))(); } catch { return null; }
}

// Una pieza de assets/js que trae textos propios (p. ej. flujos-pagina.js) los declara en window.i18nPieza = { es, en },
// en la columna 0 y cerrado con «};» en su propia línea; acá se suman a los de cada página que la carga (los de la
// página mandan, como en el navegador).
const piezasLeidas = new Map();
function textosDePiezas(html, raiz) {
  const textos = [];
  for (const m of html.matchAll(/<script[^>]*\ssrc="[^"]*assets\/js\/([\w.-]+\.js)"/g)) {
    if (!piezasLeidas.has(m[1])) {
      let dic = null;
      try { dic = leerObjeto(readFileSync(join(raiz, 'assets/js', m[1]), 'utf8'), 'i18nPieza'); } catch { /* no está */ }
      piezasLeidas.set(m[1], dic);
    }
    const dic = piezasLeidas.get(m[1]);
    if (dic && dic.en) textos.push(dic.en);
  }
  return textos;
}

/** Textos comunes a todas las páginas (nav, pie), de assets/js/idioma.js. */
export function diccionarioComun(raiz) {
  const js = readFileSync(join(raiz, 'assets/js/idioma.js'), 'utf8');
  return leerObjeto(js, 'i18nComun') || { es: {}, en: {} };
}

/** ¿Esta página tiene textos en inglés propios? */
export function tieneIngles(html) {
  const dic = leerObjeto(html, 'i18nPagina');
  return !!(dic && dic.en && Object.keys(dic.en).length);
}

/** Recorre el sitio y devuelve el conjunto de rutas que tendrán versión /en/. */
export function paginasConIngles(raiz, excluir) {
  const encontradas = new Set();
  (function recorrer(carpeta) {
    for (const nombre of readdirSync(carpeta)) {
      if (excluir.has(nombre)) continue;
      const origen = join(carpeta, nombre);
      if (statSync(origen).isDirectory()) { recorrer(origen); continue; }
      if (!nombre.toLowerCase().endsWith('.html')) continue;
      if (tieneIngles(readFileSync(origen, 'utf8'))) {
        encontradas.add(relative(raiz, origen).replace(/\\/g, '/'));
      }
    }
  })(raiz);
  return encontradas;
}

/**
 * Aplica el diccionario en inglés al HTML, igual que haría idioma.js en el
 * navegador, y deja las rutas de assets absolutas para que sigan resolviendo
 * desde /en/ (donde no hay copia de assets/ ni de marca-logo/).
 */
export function traducir(html, raiz, comun, rutaWeb, conIngles) {
  const pagina = leerObjeto(html, 'i18nPagina');
  if (!pagina || !pagina.en) return null;
  // Las tarjetas generadas desde catalogo.json traen sus textos en i18nCatalogo.
  const catalogo = leerObjeto(html, 'i18nCatalogo');
  const dic = Object.assign({}, comun.en, catalogo ? catalogo.en : {}, ...textosDePiezas(html, raiz), pagina.en);

  // Texto de los elementos con data-i18n. idioma.js usa textContent, así que
  // el contenido siempre es texto plano y se puede reemplazar entero.
  html = html.replace(
    /(<([a-zA-Z0-9]+)\b[^>]*\bdata-i18n="([^"]+)"[^>]*>)([\s\S]*?)(<\/\2>)/g,
    (todo, apertura, etiqueta, clave, contenido, cierre) =>
      (dic[clave] === undefined ? todo : apertura + escHtml(dic[clave]) + cierre));

  // Texto con formato: data-i18n-html pone el valor del diccionario como HTML
  // (lo escribimos nosotros; ver idioma.js). Va antes de reescribir los enlaces
  // de más abajo, para que también valgan para los que trae el texto.
  html = html.replace(
    /(<([a-zA-Z0-9]+)\b[^>]*\bdata-i18n-html="([^"]+)"[^>]*>)([\s\S]*?)(<\/\2>)/g,
    (todo, apertura, etiqueta, clave, contenido, cierre) =>
      (dic[clave] === undefined ? todo : apertura + dic[clave] + cierre));

  // Atributos traducibles: data-i18n-attr="content:clave;placeholder:otra"
  html = html.replace(/<[^>]*\bdata-i18n-attr="([^"]+)"[^>]*>/g, (etiqueta, reglas) => {
    let salida = etiqueta;
    for (const regla of reglas.split(';')) {
      const [attr, clave] = regla.split(':');
      if (!attr || !clave || dic[clave] === undefined) continue;
      const re = new RegExp('(\\s' + attr.trim() + '=")[^"]*(")');
      if (re.test(salida)) salida = salida.replace(re, '$1' + escAttr(dic[clave]) + '$2');
    }
    return salida;
  });

  // El idioma declarado del documento.
  html = html.replace(/<html lang="es"/i, '<html lang="en"');
  // Una página que resuelve sus enlaces desde la raíz (la 404, que se sirve en
  // cualquier dirección) los resuelve desde /en/ en su versión en inglés.
  // Los assets ya quedan absolutos más abajo.
  html = html.replace(/<base href="\/">/i, '<base href="/en/">');

  // /en/x.html y /en/herramientas/x.html no tienen assets/ propios al lado:
  // se apuntan a los de la raíz del sitio.
  html = html.replace(/(href|src)="(?:\.\.\/)?(assets|marca-logo|vendor)\//g, '$1="/$2/');
  // Lo mismo para los import de los módulos en línea. Sin esto, dictado por
  // voz y el generador de stickers en inglés importaban /en/assets/js/…, que
  // no existe: Cloudflare devolvía la portada como HTML, el navegador lo
  // rechazaba como módulo y la herramienta no arrancaba.
  html = html.replace(/(\b(?:from|import)\s*\(?\s*['"])(?:\.\.\/)?(assets|vendor)\//g, '$1/$2/');

  // No todas las páginas tienen versión en inglés (la política de privacidad,
  // el contacto y las herramientas que todavía no se tradujeron). Un enlace
  // relativo a una de ellas desde /en/ caería en una URL que no existe, así
  // que se lo manda a la versión en español, que sí está.
  const carpeta = rutaWeb.includes('/') ? rutaWeb.slice(0, rutaWeb.lastIndexOf('/')) : '';
  html = html.replace(/href="([^"#:][^"]*\.html)"/g, (todo, destino) => {
    if (destino.startsWith('/')) return todo;
    const partes = (carpeta ? carpeta.split('/') : []).concat(destino.split('/'));
    const pila = [];
    for (const parte of partes) {
      if (parte === '.' || parte === '') continue;
      if (parte === '..') pila.pop();
      else pila.push(parte);
    }
    const resuelto = pila.join('/');
    return conIngles.has(resuelto) ? todo : `href="/${resuelto}"`;
  });

  // El diccionario embebido ya no hace falta traerlo dos veces, pero se deja:
  // el botón de idioma lo usa para volver al español sin recargar nada más.
  return html;
}

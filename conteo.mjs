/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. */
// Cuántas herramientas tiene el sitio, contadas en el build.
//
// El número aparece en el título, la descripción, el subtítulo de la portada
// y el enlace al catálogo — y los conteos por categoría, en cada tarjeta del
// hub. Escritos a mano quedan viejos al primer agregado: de hecho el hub
// anterior decía "22 herramientas" de PDF cuando ya había 23.
//
// Así que en el HTML se escribe un marcador y acá se cuenta de verdad:
//   {{HERRAMIENTAS}}       -> total de archivos en herramientas/
//   {{HERRAMIENTAS:pdf}}   -> tarjetas de herramienta en pdf.html
//
// Si queda un marcador sin resolver, el build falla: es preferible eso a
// publicar "{{HERRAMIENTAS}} herramientas gratis" en la portada.

import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';

/** Cuenta las herramientas del sitio: total y por categoría. */
export function contarHerramientas(raiz) {
  const total = readdirSync(join(raiz, 'herramientas'))
    .filter((n) => n.toLowerCase().endsWith('.html')).length;

  const porCategoria = {};
  for (const nombre of readdirSync(raiz)) {
    if (!nombre.toLowerCase().endsWith('.html')) continue;
    const clave = nombre.replace(/\.html$/i, '');
    const html = readFileSync(join(raiz, nombre), 'utf8');
    // Una herramienta puede tener tarjeta en dos categorías (las
    // bidireccionales), así que se cuentan destinos distintos, no tarjetas.
    const vistas = new Set();
    const re = /<a class="tarjeta-herramienta" href="(?:\.\.\/)?herramientas\/([^"]+)"/g;
    let m;
    while ((m = re.exec(html))) vistas.add(m[1]);
    if (vistas.size) porCategoria[clave] = vistas.size;
  }
  return { total, porCategoria };
}

/**
 * Reemplaza los marcadores de conteo en una página.
 * Tira un error si queda alguno sin resolver, para que el build se corte
 * antes de publicar un marcador crudo.
 */
export function aplicarConteos(html, conteo, archivo) {
  html = html.replace(/\{\{HERRAMIENTAS(?::([a-z0-9-]+))?\}\}/gi, (todo, categoria) => {
    if (!categoria) return String(conteo.total);
    const n = conteo.porCategoria[categoria];
    if (n === undefined) {
      throw new Error(`${archivo}: ${todo} no corresponde a ninguna categoría con herramientas`);
    }
    return String(n);
  });

  const sobrante = html.match(/\{\{[A-Za-z0-9_:-]+\}\}/);
  if (sobrante) throw new Error(`${archivo}: quedó el marcador ${sobrante[0]} sin resolver`);
  return html;
}

/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. */
// Los dos sitios que salen de este repo.
//
// - principal: www.minicookies.net, con anuncios. Todas las herramientas
//   menos las "seguras".
// - seguro: seguro.minicookies.net, SIN anuncios ni scripts de terceros. Aloja
//   las herramientas donde la seguridad es la prioridad número 1 (contraseñas,
//   envío de archivos por conexión directa).
//
// Por qué un origen aparte: el navegador aísla por ORIGEN, no por página. Si
// las herramientas seguras vivieran en www, cualquier script de otra página
// del sitio (anuncios, analítica, una falla XSS) correría con los mismos
// poderes y leería, por ejemplo, las cajas cifradas del administrador de
// contraseñas (IndexedDB es por origen). Ver CLAUDE.md, "Sitio seguro".
//
// Se elige con un argumento del build: `node build.mjs --destino=seguro`
// (sin argumento, principal). `npm run build` arma los dos: principal en
// dist/ y seguro en dist-seguro/. Cloudflare publica cada uno desde su
// proyecto de Pages (mismo repo, misma orden de build, otra carpeta de salida).

export const SITIO_PRINCIPAL = 'https://www.minicookies.net';
export const SITIO_SEGURO = 'https://seguro.minicookies.net';

// Herramientas que viven SOLO en el sitio seguro (nombre de archivo sin .html).
export const SEGURAS = ['administrador-contrasenas', 'enviar-archivos-qr', 'tachar-pdf'];

export const DESTINO = process.argv.includes('--destino=seguro') ? 'seguro' : 'principal';
export const ES_SEGURO = DESTINO === 'seguro';
export const SITIO_ACTUAL = ES_SEGURO ? SITIO_SEGURO : SITIO_PRINCIPAL;
export const CARPETA_SALIDA = ES_SEGURO ? 'dist-seguro' : 'dist';

/** ¿Esta ruta del repo ("herramientas/x.html") es una herramienta segura? */
export function esHerramientaSegura(rel) {
  const m = /^herramientas\/([^/]+)\.html$/i.exec(String(rel).replace(/\\/g, '/'));
  return !!m && SEGURAS.includes(m[1]);
}

/** Dirección pública (sin .html) de una herramienta segura, en español o inglés. */
export function urlSegura(h, ingles = false) {
  return `${SITIO_SEGURO}/${ingles ? 'en/' : ''}herramientas/${h}`;
}

/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Lo mínimo para recorrer el XML de un .pptx (DrawingML, PresentationML y
// gráficos), compartido por el motor de PowerPoint: pptx-preparar.js,
// pptx-geometria.js, pptx-modelo.js y pptx-grafico.js tenían cada uno su
// copia. Por nombre local, sin mirar el espacio de nombres (a:, p:, c:…).
//
// Uso: const { hijo, hijos, ruta, todos, parsear } = PptxXml;

const PptxXml = (() => {
  // Los hijos directos con ese nombre.
  const hijos = (el, nombre) => el ? [...el.children].filter((h) => h.localName === nombre) : [];
  // El primer hijo directo con ese nombre, o null.
  const hijo = (el, nombre) => { if (!el) return null; for (const h of el.children) if (h.localName === nombre) return h; return null; };
  // Bajando de hijo en hijo: ruta(sp, 'spPr', 'xfrm', 'off'); null si falta uno.
  const ruta = (el, ...nombres) => nombres.reduce((e, n) => hijo(e, n), el);
  // Todos los descendientes con ese nombre.
  const todos = (el, nombre) => el ? [...el.getElementsByTagNameNS('*', nombre)] : [];
  const parsear = (texto) => new DOMParser().parseFromString(texto, 'application/xml');
  return { hijo, hijos, ruta, todos, parsear };
})();

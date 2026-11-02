/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Limpieza de HTML que viene de un archivo del usuario (capítulos de un EPUB,
// Markdown convertido con markdown-propio.js, HTML del lector de Word) antes de meterlo en la
// página con innerHTML. Sin esto, un archivo armado a propósito podía ejecutar
// JavaScript en el sitio: un <script> insertado con innerHTML no corre, pero
// un <img src=x onerror="..."> sí, incluso en un <div> que nunca se agrega a
// la página. El análisis se hace con DOMParser, que arma un documento inerte
// (ahí no se carga ni se ejecuta nada), y recién el resultado limpio se usa.

const ETIQUETAS_PELIGROSAS_HTML = new Set([
  'script', 'iframe', 'frame', 'frameset', 'object', 'embed', 'applet',
  'form', 'meta', 'base', 'link', 'style', 'template', 'noscript',
]);

function escaparHtml(texto) {
  return String(texto ?? '').replace(/[&<>"']/g, (c) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[c]));
}

// Un enlace o recurso solo puede apuntar a http(s), mailto, un ancla interna,
// una ruta relativa o una imagen incrustada (data:image/...). Se descartan
// javascript:, vbscript:, data:text/html, etc. Los espacios y caracteres de
// control se ignoran al comparar porque el navegador también los ignora
// ("java\nscript:" funciona como "javascript:").
function urlEsSegura(valor, esImagen) {
  const limpio = String(valor).replace(/[\u0000-\u0020]/g, '').toLowerCase();
  if (limpio.startsWith('data:')) return esImagen && /^data:image\/(png|jpe?g|gif|webp|bmp|svg\+xml);/.test(limpio);
  const esquema = /^([a-z][a-z0-9+.-]*):/.exec(limpio);
  return !esquema || ['http', 'https', 'mailto', 'tel'].includes(esquema[1]);
}

// Limpia en el lugar un árbol de nodos (de un documento inerte).
function sanitizarArbolHtml(raiz) {
  const elementos = [...raiz.querySelectorAll('*')];
  for (const el of elementos) {
    const etiqueta = el.localName.toLowerCase();
    if (ETIQUETAS_PELIGROSAS_HTML.has(etiqueta)) { el.remove(); continue; }
    for (const atributo of [...el.attributes]) {
      const nombre = atributo.name.toLowerCase();
      if (nombre.startsWith('on') || nombre === 'srcdoc' || nombre === 'formaction') {
        el.removeAttribute(atributo.name);
      } else if (['href', 'src', 'xlink:href', 'action', 'background', 'poster', 'data'].includes(nombre)) {
        if (!urlEsSegura(atributo.value, nombre === 'src' || (etiqueta === 'image' && nombre.endsWith('href')))) {
          el.removeAttribute(atributo.name);
        }
      } else if (nombre === 'style' && /expression\s*\(|url\s*\(\s*['"]?\s*javascript:/i.test(atributo.value)) {
        el.removeAttribute(atributo.name);
      }
    }
    // Los enlaces externos abren en otra pestaña sin darle acceso a esta (los
    // internos, como una nota al pie "#nota1", se dejan como están).
    if (etiqueta === 'a' && /^\s*https?:/i.test(el.getAttribute('href') || '')) {
      el.setAttribute('rel', 'noopener noreferrer');
      el.setAttribute('target', '_blank');
    }
  }
  return raiz;
}

// Recibe un string de HTML y devuelve el mismo HTML ya limpio.
function sanitizarHtml(html) {
  const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
  return sanitizarArbolHtml(doc.body).innerHTML;
}

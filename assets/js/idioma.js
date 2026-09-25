/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Selector de idioma es/en para todo el sitio.
// Cada página define window.i18nPagina = {es:{...}, en:{...}} con sus textos propios
// (títulos de sección, tarjetas de herramientas). Este archivo aporta los textos
// comunes a todas (nav, pie de página, aviso de privacidad) y aplica la traducción.

window.i18nComun = {
  es: {
    'nav-todas': 'Todas las herramientas',
    'nav-todas-cat': 'Todas las categorías',
    'nav-privacidad': 'Privacidad',
    'nav-terminos': 'Términos de uso',
    'anuncio': 'Espacio publicitario',
    'aviso-privacidad': '🔒 Tus archivos nunca se suben a ningún servidor: todo el procesamiento ocurre en tu propia computadora.',
    'footer-tagline': 'Mini Cookies — gratis, sin registro, sin subir tus archivos a ningún lado.',
    'footer-todas': 'Todas las herramientas',
    'footer-privacidad': 'Política de privacidad',
    'footer-terminos': 'Términos de uso',
    'footer-creditos': 'Créditos',
    'footer-contacto': 'Contacto',
    // Compartidas por las 76 páginas de herramientas (header/nav y helpers JS).
    'nav-cat-pdf': 'Herramientas de PDF',
    'nav-cat-excel': 'Herramientas de Excel / CSV',
    'nav-cat-word': 'Herramientas de Word',
    'nav-cat-powerpoint': 'Herramientas de PowerPoint',
    'nav-cat-epub': 'Herramientas de EPUB',
    'nav-cat-markdown': 'Herramientas de Markdown',
    'nav-cat-texto': 'Herramientas de Texto plano',
    'nav-cat-generales': 'Herramientas generales',
    'tema-aria': 'Cambiar a modo oscuro',
    'tema-title': 'Cambiar tema',
    'previsualizar-fallo': '(no se pudo previsualizar)',
    'generando-miniaturas': 'Generando miniaturas...',
    'sufijo-pag': 'pág.',
    'quitar': 'Quitar',
    'vista-cuadricula': '▦ Cuadrícula',
    'vista-lista': '☰ Lista',
    // Usadas por assets/js/pdf-texto.js (extraerTextoPowerPoint), compartida
    // entre powerpoint-a-pdf.html y extraer-texto-powerpoint.html.
    'ppt-sin-diapositivas': 'No se encontraron diapositivas legibles en este archivo.',
    'ppt-diapositiva-titulo': '--- Diapositiva {n} ---',
    'ppt-sin-texto': '(sin texto)',
    // Usadas por assets/js/epub.js (leerEpub/construirEpub), compartidas por
    // todas las herramientas de EPUB.
    'epub-drm': 'Este EPUB está protegido con DRM (cifrado) — no se puede leer su contenido en el navegador.',
    'epub-sin-container': 'El EPUB no tiene un container.xml válido.',
    'epub-sin-titulo': 'Sin título',
    'epub-sin-contenido': 'No se encontró contenido legible dentro del EPUB.',
    // Usadas por assets/js/pptx-visual.js, compartidas por powerpoint-a-pdf.html
    // y powerpoint-a-imagenes.html.
    'ppt-sin-tamano': 'No se pudo determinar el tamaño de las diapositivas.',
    'ppt-tamano-invalido': 'Tamaño de diapositiva inválido.',
    'ppt-render-inconsistente': 'El renderizado devolvió un tamaño de diapositiva inconsistente.',
    // Usada por detalleError() (abajo) en todas las herramientas de PDF.
    'pdf-protegido': 'el PDF está protegido (tiene contraseña o restricciones de edición). Quitale la protección primero con la herramienta "Desbloquear PDF" y volvé a intentarlo.',
    // Mensaje genérico de lectura fallida (las páginas pueden definir uno propio).
    'error-leer': 'No se pudo leer el archivo. Puede estar dañado o no tener un formato válido',
    'xml-invalido': 'XML inválido'
  },
  en: {
    'nav-todas': 'All tools',
    'nav-todas-cat': 'All categories',
    'nav-privacidad': 'Privacy',
    'nav-terminos': 'Terms of use',
    'anuncio': 'Ad space',
    'aviso-privacidad': '🔒 Your files never get uploaded anywhere: everything runs on your own computer.',
    'footer-tagline': 'Mini Cookies — free, no sign-up, your files never leave your device.',
    'footer-todas': 'All tools',
    'footer-privacidad': 'Privacy policy',
    'footer-terminos': 'Terms of use',
    'footer-creditos': 'Credits',
    'footer-contacto': 'Contact',
    'nav-cat-pdf': 'PDF tools',
    'nav-cat-excel': 'Excel / CSV tools',
    'nav-cat-word': 'Word tools',
    'nav-cat-powerpoint': 'PowerPoint tools',
    'nav-cat-epub': 'EPUB tools',
    'nav-cat-markdown': 'Markdown tools',
    'nav-cat-texto': 'Plain text tools',
    'nav-cat-generales': 'General tools',
    'tema-aria': 'Switch to dark mode',
    'tema-title': 'Toggle theme',
    'previsualizar-fallo': '(preview unavailable)',
    'generando-miniaturas': 'Generating thumbnails...',
    'sufijo-pag': 'p.',
    'quitar': 'Remove',
    'vista-cuadricula': '▦ Grid',
    'vista-lista': '☰ List',
    'ppt-sin-diapositivas': 'No readable slides were found in this file.',
    'ppt-diapositiva-titulo': '--- Slide {n} ---',
    'ppt-sin-texto': '(no text)',
    'epub-drm': 'This EPUB is DRM-protected (encrypted) — its content can’t be read in the browser.',
    'epub-sin-container': 'The EPUB doesn’t have a valid container.xml.',
    'epub-sin-titulo': 'Untitled',
    'epub-sin-contenido': 'No readable content was found inside the EPUB.',
    'ppt-sin-tamano': 'Couldn’t determine the size of the slides.',
    'ppt-tamano-invalido': 'Invalid slide size.',
    'ppt-render-inconsistente': 'Rendering returned an inconsistent slide size.',
    'pdf-protegido': 'the PDF is protected (it has a password or editing restrictions). Remove the protection first with the "Unlock PDF" tool and try again.',
    'error-leer': 'Couldn’t read the file. It may be damaged or not in a valid format',
    'xml-invalido': 'Invalid XML'
  }
};

(function () {
  // Las páginas traducidas se publican dos veces, una por idioma y con URL
  // propia (/pdf.html y /en/pdf.html), para que un buscador pueda indexar
  // las dos. El build deja en cada una un <link rel="alternate"> apuntando a
  // su gemela, y de ahí sale a dónde lleva el botón de idioma.
  function alternativa(idioma) {
    var l = document.querySelector('link[rel="alternate"][hreflang="' + idioma + '"]');
    return l ? l.href : null;
  }

  function idiomaActual() {
    // Si la página tiene gemela, manda la URL: el HTML ya viene en ese idioma
    // desde el servidor y reescribirlo según localStorage lo contradiría.
    if (alternativa('en')) {
      return document.documentElement.getAttribute('lang') === 'en' ? 'en' : 'es';
    }
    // Páginas sin versión en inglés todavía: se traduce lo que haya en el acto.
    try { return localStorage.getItem('idioma') || 'es'; } catch (e) { return 'es'; }
  }

  function diccionario(idioma) {
    var pagina = (window.i18nPagina && window.i18nPagina[idioma]) || {};
    var comun = window.i18nComun[idioma] || window.i18nComun.es;
    return Object.assign({}, comun, pagina);
  }

  // Traducción para texto generado por JavaScript (mensajes de botones en
  // proceso, alertas de error, etiquetas de miniaturas, etc.), que data-i18n
  // no puede cubrir porque no existe en el HTML hasta que el script corre.
  // Cada página inline y los helpers compartidos (miniaturas.js, etc.) llaman
  // a t('clave') en vez de escribir el texto fijo en español. Si la clave no
  // está en ningún diccionario, devuelve la clave misma como último recurso.
  window.t = function (clave) {
    var dic = diccionario(idiomaActual());
    return dic[clave] !== undefined ? dic[clave] : clave;
  };
  window.idiomaActual = idiomaActual;

  // Texto para mostrarle al usuario cuando algo falla. Los errores de las
  // librerías vienen en inglés técnico; el caso más común que tiene una
  // explicación clara (PDF protegido: pdf-lib tira "...is encrypted...") se
  // reemplaza por un mensaje traducido que dice qué hacer.
  window.detalleError = function (error) {
    var mensaje = (error && error.message) || String(error);
    if (/encrypted/i.test(mensaje)) return t('pdf-protegido');
    return mensaje;
  };

  function aplicarIdioma(idioma) {
    var dic = diccionario(idioma);

    document.querySelectorAll('[data-i18n]').forEach(function (el) {
      var clave = el.getAttribute('data-i18n');
      if (dic[clave] !== undefined) el.textContent = dic[clave];
    });

    document.querySelectorAll('[data-i18n-attr]').forEach(function (el) {
      el.getAttribute('data-i18n-attr').split(';').forEach(function (regla) {
        var partes = regla.split(':');
        var attr = partes[0], clave = partes[1];
        if (dic[clave] !== undefined) el.setAttribute(attr, dic[clave]);
      });
    });

    document.documentElement.setAttribute('lang', idioma);

    var boton = document.getElementById('botonIdioma');
    if (boton) {
      boton.textContent = idioma === 'es' ? 'EN' : 'ES';
      boton.setAttribute('title', idioma === 'es' ? 'Switch to English' : 'Cambiar a español');
    }

    // Avisa a la página (código inline de cada herramienta) que el idioma
    // cambió, para que pueda re-renderizar texto generado dinámicamente que
    // ya estaba dibujado en pantalla (miniaturas, mensajes de resultado, etc.).
    document.dispatchEvent(new CustomEvent('cambioIdioma', { detail: { idioma: idioma } }));
  }

  document.addEventListener('DOMContentLoaded', function () {
    aplicarIdioma(idiomaActual());

    var boton = document.getElementById('botonIdioma');
    if (!boton) return;
    boton.addEventListener('click', function () {
      var nuevo = idiomaActual() === 'es' ? 'en' : 'es';
      try { localStorage.setItem('idioma', nuevo); } catch (e) {}

      // Si esta página existe en el otro idioma, se va a esa URL en vez de
      // reescribir el texto acá: así la dirección que queda en la barra (y la
      // que alguien copia y comparte) es la del idioma que se está viendo.
      var destino = alternativa(nuevo);
      if (destino && destino !== location.href.split('#')[0]) {
        location.href = destino;
        return;
      }
      aplicarIdioma(nuevo);
    });
  });
})();

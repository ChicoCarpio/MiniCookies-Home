// Selector de idioma es/en. Cada página define window.i18nPagina = {es:{...}, en:{...}}
// con sus textos propios; este archivo aporta los textos comunes (nav, pie de página,
// aviso de privacidad) y aplica la traducción.

window.i18nComun = {
  es: {
    'nav-todas': 'Todas las herramientas',
    'nav-privacidad': 'Privacidad',
    'nav-terminos': 'Términos de uso',
    'aviso-privacidad': '🔒 Tus archivos nunca se suben a ningún servidor: todo el procesamiento ocurre en tu propia computadora.',
    'footer-tagline': 'Mini Cookies — gratis, sin registro, sin subir tus archivos a ningún lado.',
    'footer-privacidad': 'Política de privacidad',
    'footer-terminos': 'Términos de uso',
    'footer-creditos': 'Créditos',
    'footer-contacto': 'Contacto'
  },
  en: {
    'nav-todas': 'All tools',
    'nav-privacidad': 'Privacy',
    'nav-terminos': 'Terms of use',
    'aviso-privacidad': '🔒 Your files never get uploaded anywhere: everything runs on your own computer.',
    'footer-tagline': 'Mini Cookies — free, no sign-up, your files never leave your device.',
    'footer-privacidad': 'Privacy policy',
    'footer-terminos': 'Terms of use',
    'footer-creditos': 'Credits',
    'footer-contacto': 'Contact'
  }
};

(function () {
  function idiomaActual() {
    try { return localStorage.getItem('idioma') || 'es'; } catch (e) { return 'es'; }
  }

  function diccionario(idioma) {
    var pagina = (window.i18nPagina && window.i18nPagina[idioma]) || {};
    var comun = window.i18nComun[idioma] || window.i18nComun.es;
    return Object.assign({}, comun, pagina);
  }

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
  }

  document.addEventListener('DOMContentLoaded', function () {
    aplicarIdioma(idiomaActual());

    var boton = document.getElementById('botonIdioma');
    if (!boton) return;
    boton.addEventListener('click', function () {
      var nuevo = idiomaActual() === 'es' ? 'en' : 'es';
      try { localStorage.setItem('idioma', nuevo); } catch (e) {}
      aplicarIdioma(nuevo);
    });
  });
})();

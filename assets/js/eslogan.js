/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. */
// Sortea uno de los eslóganes de la marca al cargar la página.
//
// Lo usa cualquier página que tenga un elemento con id="eslogan"; hoy la
// portada y la política de privacidad.
//
// En vez de escribir el texto, cambia el data-i18n del elemento a la clave
// sorteada y deja que idioma.js lo traduzca. Así el eslogan sigue cambiando
// de idioma con el botón, sin que este archivo sepa nada de traducciones.
//
// El HTML trae un eslogan escrito de verdad (no un hueco vacío): si el script
// no corre, o lo lee un buscador que no ejecuta JavaScript, igual se ve uno.
//
// Va cargado DESPUÉS de idioma.js: necesita leer window.i18nComun para pintar
// el elegido enseguida, y se ejecuta antes de que idioma.js aplique la
// traducción en DOMContentLoaded, así no se ve el cambio en pantalla.

(function () {
  var el = document.getElementById('eslogan');
  if (!el || !window.i18nComun) return;

  // Cuántos hay se deduce del diccionario, así que agregar o sacar un eslogan
  // en idioma.js alcanza: acá no hay ningún número que actualizar.
  var claves = Object.keys(window.i18nComun.es || {}).filter(function (k) {
    return /^eslogan-\d+$/.test(k);
  });
  if (!claves.length) return;

  var clave = claves[Math.floor(Math.random() * claves.length)];
  el.setAttribute('data-i18n', clave);

  var idioma = 'es';
  try { idioma = window.idiomaActual ? window.idiomaActual() : (localStorage.getItem('idioma') || 'es'); } catch (e) {}
  var dic = window.i18nComun[idioma] || window.i18nComun.es;
  if (dic[clave]) el.textContent = dic[clave];
})();

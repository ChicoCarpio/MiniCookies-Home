/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
(function () {
  function iconoPara(tema) {
    return tema === 'dark' ? '☀️' : '🌙';
  }

  function temaActual() {
    var guardado = document.documentElement.getAttribute('data-theme');
    if (guardado) return guardado;
    return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  }

  document.addEventListener('DOMContentLoaded', function () {
    var boton = document.getElementById('botonTema');
    if (!boton) return;
    boton.textContent = iconoPara(temaActual());
    boton.addEventListener('click', function () {
      var nuevo = temaActual() === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', nuevo);
      try { localStorage.setItem('tema', nuevo); } catch (e) {}
      boton.textContent = iconoPara(nuevo);
    });
  });
})();

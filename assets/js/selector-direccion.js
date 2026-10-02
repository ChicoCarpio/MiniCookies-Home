/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Selector de dirección de las herramientas bidireccionales (PDF → Texto /
// Texto → PDF, etc.): muestra el panel de la dirección elegida y oculta los
// demás.
//
// Por qué existe: diez herramientas tenían el mismo bloque de siete líneas
// copiado, cambiando solo los nombres de los paneles. Ahora es declarativo:
// cada panel dice a qué dirección pertenece y este archivo hace el resto.
//
//   <div class="selector-direccion">
//     <label><input type="radio" name="direccion" value="pdf-a-texto" checked> …</label>
//     <label><input type="radio" name="direccion" value="texto-a-pdf"> …</label>
//   </div>
//   <div data-panel-direccion="pdf-a-texto">…</div>
//   <div data-panel-direccion="texto-a-pdf" style="display:none;">…</div>
//
// (No es data-direccion a secas: pdf-a-pdfa ya usa ese atributo en sus
// tarjetas de opción, y este script las ocultaría.)
// (El style="display:none" del segundo panel evita que se vea un instante
// antes de que cargue este script.) También se sincroniza al cargar y al
// volver a la página: el navegador puede restaurar la opción elegida, y sin
// esto quedaba marcada una dirección mostrando el panel de la otra.

(function () {
  function sincronizar(selector) {
    const marcado = selector.querySelector('input[type="radio"]:checked');
    if (!marcado) return;
    document.querySelectorAll('[data-panel-direccion]').forEach((panel) => {
      panel.style.display = panel.dataset.panelDireccion === marcado.value ? '' : 'none';
    });
  }

  function iniciar() {
    document.querySelectorAll('.selector-direccion').forEach((selector) => {
      selector.querySelectorAll('input[type="radio"]').forEach((radio) => {
        radio.addEventListener('change', () => sincronizar(selector));
      });
      sincronizar(selector);
      window.addEventListener('pageshow', () => sincronizar(selector));
    });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();
})();

/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. */
// Arma la distribución de celular de los editores de imagen (ver el dibujo en
// assets/css/editor-movil.css).
//
// Idea central: NO se duplican los controles. Los paneles que en computadora
// son columnas se mueven —el mismo elemento, no una copia— al panel de abajo
// cuando se elige su categoría, y vuelven a su lugar al cerrarla o al pasar a
// pantalla ancha. Mover un elemento conserva sus escuchadores, así que todo lo
// que ya funcionaba sigue funcionando sin tocar una línea de los editores.
//
// Las categorías salen de lo que la página ya tiene: cada <fieldset
// class="acordeon-grupo"> del panel de ajustes es una, más los paneles de
// filtros y de capas. Si mañana se agrega un grupo de ajustes, aparece solo
// en la tira de abajo.

(function () {
  var ANCHO_CELULAR = '(max-width: 900px)';

  // Ícono por categoría. La clave es el texto del título en español, en
  // minúsculas y sin acentos; si una categoría no está acá, se usa el
  // comodín (no se rompe nada, solo sale un ícono genérico).
  var ICONOS = {
    'luz': '☀️',
    'color': '🎨',
    'luz por color': '🌈',
    'blanco y negro': '◐',
    'efectos': '✨',
    'resplandor': '💫',
    'dispersion de luz': '🔦',
    'textura': '🧱',
    'filtros': '🖼️',
    'capas': '🗂️',
  };

  function sinAcentos(t) {
    return t.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
  }

  function iniciar(shell) {
    var cuerpo = document.body;
    var consulta = window.matchMedia(ANCHO_CELULAR);

    var panelAjustes = shell.querySelector('.panel-ajustes');
    var panelFiltros = shell.querySelector('.panel-filtros');
    var panelCapas = shell.querySelector('.panel-capas');
    var lienzo = shell.querySelector('.panel-lienzo');
    if (!lienzo) return;

    shell.classList.add('em-shell');

    // --- Piezas nuevas: tira de publicidad, hoja, categorías y acciones ---
    var tiraAnuncio = document.createElement('div');
    tiraAnuncio.className = 'em-anuncio';

    var hoja = document.createElement('div');
    hoja.className = 'em-hoja';

    var chips = document.createElement('nav');
    chips.className = 'em-chips';
    chips.setAttribute('aria-label', 'Categorías de ajuste');

    var acciones = document.createElement('div');
    acciones.className = 'em-acciones';

    // El anuncio de la página se sube a la tira de arriba mientras dure el
    // modo celular; se recuerda de dónde salió para devolverlo después.
    var anuncio = document.querySelector('.anuncio');
    var anuncioPadre = anuncio ? anuncio.parentNode : null;
    var anuncioSiguiente = anuncio ? anuncio.nextSibling : null;

    // --- Categorías ---
    // Cada una sabe qué elemento mostrar y, si es un grupo del acordeón, cuál
    // es su contenedor original para poder devolverlo.
    var categorias = [];

    function agregarCategoria(etiqueta, elemento, contenedorOriginal) {
      if (!elemento) return;
      categorias.push({
        etiqueta: etiqueta,
        elemento: elemento,
        padre: contenedorOriginal || elemento.parentNode,
        siguiente: elemento.nextSibling,
      });
    }

    if (panelFiltros) agregarCategoria('Filtros', panelFiltros);

    // Una página puede marcar un bloque suelto como categoría propia con
    // data-em-categoria="Nombre". Se usa para cosas que en computadora viven
    // dentro de otro panel pero que en celular merecen su propio botón (los
    // filtros de capa del editor avanzado, por ejemplo).
    shell.querySelectorAll('[data-em-categoria]').forEach(function (el) {
      agregarCategoria(el.getAttribute('data-em-categoria'), el);
    });

    if (panelAjustes) {
      panelAjustes.querySelectorAll('.acordeon-grupo').forEach(function (grupo) {
        var titulo = grupo.querySelector('.acordeon-titulo');
        if (!titulo) return;
        var texto = titulo.textContent.replace(/[▾▴▲▼]/g, '').trim();
        agregarCategoria(texto, grupo);
      });
    }
    if (panelCapas) agregarCategoria('Capas', panelCapas);
    if (!categorias.length) return;

    var abierta = -1;

    categorias.forEach(function (cat, i) {
      var boton = document.createElement('button');
      boton.type = 'button';
      boton.className = 'em-chip';
      var icono = document.createElement('span');
      icono.className = 'em-icono';
      icono.textContent = ICONOS[sinAcentos(cat.etiqueta)] || '🎚️';
      var nombre = document.createElement('span');
      nombre.textContent = cat.etiqueta;
      boton.appendChild(icono);
      boton.appendChild(nombre);
      boton.addEventListener('click', function () { alternar(i); });
      cat.chip = boton;
      chips.appendChild(boton);
    });

    function cerrarHoja() {
      if (abierta < 0) return;
      var cat = categorias[abierta];
      // De vuelta a su lugar exacto en el panel de origen.
      if (cat.padre) cat.padre.insertBefore(cat.elemento, cat.siguiente);
      cat.chip.classList.remove('activo');
      abierta = -1;
    }

    function alternar(i) {
      if (abierta === i) { cerrarHoja(); return; }
      cerrarHoja();
      var cat = categorias[i];
      hoja.appendChild(cat.elemento);
      cat.chip.classList.add('activo');
      abierta = i;
      // Si el grupo venía plegado en la versión de computadora, acá tiene que
      // verse igual: la hoja ya muestra el contenido, pero la clase mantiene
      // coherente el estado por si se vuelve a pantalla ancha.
      cat.elemento.classList.add('abierto');
    }

    // --- Acciones siempre visibles ---
    // Se reutilizan los botones que ya existen en la página (mismo id, mismos
    // escuchadores): estos solo los "pulsan".
    // Los dos editores nombran distinto los mismos botones (el avanzado les
    // agrega "Avanzado"), así que se prueban los nombres posibles y se usa el
    // primero que exista.
    function accion(texto, ids, clase) {
      var destino = null;
      for (var i = 0; i < ids.length && !destino; i++) destino = document.getElementById(ids[i]);
      if (!destino) return null;
      var b = document.createElement('button');
      b.type = 'button';
      b.className = clase;
      b.textContent = texto;
      b.addEventListener('click', function () { destino.click(); });
      // El botón de descargar arranca deshabilitado hasta que hay imagen:
      // se copia ese estado del original.
      if (destino.disabled !== undefined) {
        b.disabled = destino.disabled;
        new MutationObserver(function () { b.disabled = destino.disabled; })
          .observe(destino, { attributes: true, attributeFilter: ['disabled'] });
      }
      acciones.appendChild(b);
      return b;
    }

    accion('🖼️', ['botonCambiarImagen'], 'em-secundario');
    accion('↺', ['botonRestablecer', 'botonRestablecerAvanzado'], 'em-secundario');
    accion('⤓ Descargar', ['botonDescargar', 'botonDescargarAvanzado'], 'em-principal');

    // --- Encender / apagar según el ancho ---
    function encender() {
      if (cuerpo.classList.contains('em-activo')) return;
      cuerpo.classList.add('em-activo');
      if (anuncio) tiraAnuncio.appendChild(anuncio);
      shell.insertBefore(tiraAnuncio, shell.firstChild);
      shell.appendChild(hoja);
      shell.appendChild(chips);
      shell.appendChild(acciones);
      avisarCambioDeTamano();
    }

    function apagar() {
      if (!cuerpo.classList.contains('em-activo')) return;
      cerrarHoja();
      cuerpo.classList.remove('em-activo');
      // Todo vuelve a donde estaba: la página queda igual que antes.
      if (anuncio && anuncioPadre) anuncioPadre.insertBefore(anuncio, anuncioSiguiente);
      [tiraAnuncio, hoja, chips, acciones].forEach(function (el) {
        if (el.parentNode) el.parentNode.removeChild(el);
      });
      avisarCambioDeTamano();
    }

    // Los editores recalculan el tamaño del lienzo al cambiar la ventana; el
    // modo celular cambia el espacio disponible sin que la ventana cambie, así
    // que hay que avisarles igual.
    function avisarCambioDeTamano() {
      window.dispatchEvent(new Event('resize'));
    }

    function revisar() {
      // Solo tiene sentido cuando hay una imagen cargada: el editor está
      // oculto hasta entonces.
      var visible = shell.offsetParent !== null || shell.getClientRects().length > 0;
      if (consulta.matches && visible) encender(); else apagar();
    }

    if (consulta.addEventListener) consulta.addEventListener('change', revisar);
    else consulta.addListener(revisar);

    // El editor aparece cuando el usuario carga una imagen: se vigila ese
    // cambio de visibilidad en vez de pedirle a cada editor que avise.
    var contenedor = document.getElementById('editorCompleto') || shell;
    new MutationObserver(revisar).observe(contenedor, { attributes: true, attributeFilter: ['style', 'class'] });

    revisar();
  }

  document.addEventListener('DOMContentLoaded', function () {
    var shell = document.querySelector('.editor-color-layout, .editor-avanzado-layout');
    if (shell) iniciar(shell);
  });
})();

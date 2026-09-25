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

    // Las categorías no entran todas a lo ancho de un celular. Se corren con
    // el dedo, pero eso solo no alcanza: con la barra de desplazamiento
    // oculta no se ve que haya más, y con mouse (ventana angosta en una
    // computadora) no había ninguna forma de moverlas.
    function marcarBordes() {
      var resto = chips.scrollWidth - chips.clientWidth;
      chips.classList.toggle('hay-mas-izquierda', chips.scrollLeft > 4);
      chips.classList.toggle('hay-mas-derecha', resto > 4 && chips.scrollLeft < resto - 4);
    }

    chips.addEventListener('scroll', marcarBordes);

    // Arrastrar la tira, agarrándola. No se deja en manos del desplazamiento
    // nativo del navegador porque con mouse (ventana angosta en una
    // computadora, o el emulador de celular) no se arrastra, y ahí las
    // últimas categorías quedaban fuera de alcance. Hecho con eventos de
    // puntero, el mismo gesto sirve para el dedo y para el mouse.
    var arrastre = null;
    var huboArrastre = false;
    var MINIMO_PARA_ARRASTRE = 6; // px: menos que esto es un toque, no un arrastre

    chips.addEventListener('pointerdown', function (e) {
      if (e.pointerType === 'mouse' && e.button !== 0) return;
      huboArrastre = false;
      // Ojo: acá NO se captura el puntero. Capturarlo desde el principio
      // desvía el click al contenedor y entonces tocar una categoría deja de
      // abrirla. Se captura recién cuando el gesto resulta ser un arrastre.
      arrastre = { id: e.pointerId, x: e.clientX, desde: chips.scrollLeft, maximo: 0, capturado: false };
    });

    chips.addEventListener('pointermove', function (e) {
      if (!arrastre || e.pointerId !== arrastre.id) return;
      var dx = e.clientX - arrastre.x;
      if (Math.abs(dx) > arrastre.maximo) arrastre.maximo = Math.abs(dx);
      if (arrastre.maximo > MINIMO_PARA_ARRASTRE) {
        huboArrastre = true;
        chips.classList.add('arrastrando-tira');
        if (!arrastre.capturado) {
          // Ya es un arrastre: ahora sí conviene capturar, para que siga
          // funcionando aunque el dedo se salga de la tira.
          arrastre.capturado = true;
          try { chips.setPointerCapture(e.pointerId); } catch (err) {}
        }
      }
      if (!huboArrastre) return;
      chips.scrollLeft = arrastre.desde - dx;
      marcarBordes();
    });

    function soltarTira(e) {
      if (!arrastre || e.pointerId !== arrastre.id) return;
      if (arrastre.capturado) {
        try { chips.releasePointerCapture(e.pointerId); } catch (err) {}
      }
      arrastre = null;
      chips.classList.remove('arrastrando-tira');
    }
    chips.addEventListener('pointerup', soltarTira);
    chips.addEventListener('pointercancel', soltarTira);

    // Si el dedo se movió, el gesto fue correr la tira: no hay que abrir la
    // categoría que quedó debajo al soltar.
    chips.addEventListener('click', function (e) {
      if (!huboArrastre) return;
      e.preventDefault();
      e.stopPropagation();
    }, true);

    // La rueda del mouse mueve de costado: una rueda común solo da deltaY, y
    // sin esto la tira quedaba inmóvil en una ventana angosta de computadora.
    chips.addEventListener('wheel', function (e) {
      var resto = chips.scrollWidth - chips.clientWidth;
      if (resto <= 0) return;
      var paso = Math.abs(e.deltaY) > Math.abs(e.deltaX) ? e.deltaY : e.deltaX;
      if (!paso) return;
      e.preventDefault();
      chips.scrollLeft += paso;
    }, { passive: false });

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
      // Si la categoría estaba cortada contra el borde, se la trae a la vista.
      if (cat.chip.scrollIntoView) cat.chip.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'nearest' });
      abierta = i;
      // Si el grupo venía plegado en la versión de computadora, acá tiene que
      // verse igual: la hoja ya muestra el contenido, pero la clase mantiene
      // coherente el estado por si se vuelve a pantalla ancha.
      cat.elemento.classList.add('abierto');
      // Un bloque marcado con data-em-categoria puede ser un <details>
      // plegado (el filtro de capa del editor avanzado): elegir la categoría
      // y que no se vea nada no tendría sentido.
      if (cat.elemento.tagName === 'DETAILS') cat.elemento.open = true;
      // El panel arranca desde arriba aunque la categoría anterior estuviera
      // desplazada.
      hoja.scrollTop = 0;
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
      marcarBordes();
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

    window.addEventListener('resize', function () {
      if (cuerpo.classList.contains('em-activo')) marcarBordes();
    });

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

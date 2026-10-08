/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Arma la distribución de celular de los editores de imagen (ver el dibujo en
// assets/css/app-movil.css).
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
  var TEXTOS = {
    es: { filtros: 'Filtros', capas: 'Capas', categorias: 'Categorías de ajuste' },
    en: { filtros: 'Filters', capas: 'Layers', categorias: 'Adjustment categories' },
  };
  var txt = function (clave) {
    var lang = (window.idiomaActual && idiomaActual() === 'en') ? 'en' : 'es';
    return TEXTOS[lang][clave];
  };

  // Por clave data-i18n del titulo del grupo: no cambia con el idioma.
  // (El mapa por texto de abajo queda como respaldo para un grupo que
  // todavia no tenga data-i18n.)
  var ICONOS_POR_CLAVE = {
    'grupo-luz': '☀️',
    'grupo-color': '🎨',
    'grupo-luz-color': '🌈',
    'grupo-bn': '◐',
    'grupo-efectos': '✨',
    'grupo-resplandor': '💫',
    'grupo-dispersion': '🔦',
    'grupo-textura': '🧱',
    'titulo-aplicar': '🎯',
    'titulo-vista': '🔍',
    'titulo-contenido': '🔤',
    'titulo-capas': '🗂️',
    'titulo-textura': '🧱',
    'Original': '👁️',
    'resumen-filtro': '🖼️',
    'grupo-filtros': '🖼️',
    'grupo-acciones': '⤓',
    'campo-color-borde': '🎨',
    // Collage
    'h-diseno': '▦',
    'h-formato': '🖼️',
    'h-estilo': '🎨',
    'h-descargar': '⤓',
  };

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

  // El nombre de una categoría es el texto traducido de su título (el
  // elemento con data-i18n), no todo el título: así lo que el título agregue
  // al lado (la cantidad de capas del editor avanzado, la flecha) no pasa al
  // botón de la tira.
  function textoDeTitulo(titulo) {
    var conClave = titulo.hasAttribute('data-i18n') ? titulo : (titulo.querySelector('[data-i18n]') || titulo);
    return conClave.textContent.replace(/[▾▴▲▼]/g, '').trim();
  }

  function sinAcentos(t) {
    return t.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
  }

  // El anuncio real es un <iframe> de AdSense: con insertBefore/appendChild
  // el navegador lo recarga y queda en blanco. moveBefore (Chrome/Edge 133+)
  // lo mueve sin recargarlo; donde no existe, se mueve como siempre.
  // moveBefore solo mueve entre nodos que ya están en la página: con la
  // tira todavía suelta tiraba HierarchyRequestError, y como cortaba
  // encender() a la mitad, el modo celular quedaba sin categorías ni
  // botones (2026-10-04). Si igual falla, se mueve como siempre.
  // También la usa pantalla-completa.js (window.AppMovil.moverSinRecargar).
  function mover(padre, elemento, referencia) {
    if (padre.moveBefore && padre.isConnected && elemento.isConnected) {
      try { padre.moveBefore(elemento, referencia); return; } catch (e) { /* abajo */ }
    }
    padre.insertBefore(elemento, referencia);
  }
  window.AppMovil = { moverSinRecargar: mover };

  function iniciar(shell) {
    var cuerpo = document.body;
    var consulta = window.matchMedia(ANCHO_CELULAR);

    var panelAjustes = shell.querySelector('.panel-ajustes');
    var panelFiltros = shell.querySelector('.panel-filtros');
    var panelCapas = shell.querySelector('.panel-capas');

    // El escenario es lo que se está editando y tiene que quedar siempre a la
    // vista. Una página lo marca con data-app-escenario; las que vienen de
    // antes usan .panel-lienzo. Si lleva el valor "desplazable", además se
    // comporta como lienzo con zoom propio (ver app-movil.css).
    var escenario = shell.querySelector('[data-app-escenario]') || shell.querySelector('.panel-lienzo');
    if (!escenario) return;
    escenario.classList.add('app-escenario');
    if (escenario.getAttribute('data-app-escenario') === 'desplazable') {
      escenario.classList.add('app-desplazable');
    }

    shell.classList.add('app-shell');

    // --- Piezas nuevas: tira de publicidad, hoja, categorías y acciones ---
    var tiraAnuncio = document.createElement('div');
    tiraAnuncio.className = 'app-anuncio';

    var hoja = document.createElement('div');
    hoja.className = 'app-hoja';

    var chips = document.createElement('nav');
    chips.className = 'app-chips';
    chips.setAttribute('aria-label', txt('categorias'));

    var acciones = document.createElement('div');
    acciones.className = 'app-acciones';

    // El anuncio de la página se sube a la tira de arriba mientras dure el
    // modo celular; se recuerda de dónde salió para devolverlo después.
    var anuncio = document.querySelector('.anuncio');
    var anuncioPadre = anuncio ? anuncio.parentNode : null;
    var anuncioSiguiente = anuncio ? anuncio.nextSibling : null;

    // --- Categorías ---
    // Cada una sabe qué elemento mostrar y, si es un grupo del acordeón, cuál
    // es su contenedor original para poder devolverlo.
    var categorias = [];

    function agregarCategoria(etiqueta, elemento, contenedorOriginal, claveIcono, titulo) {
      if (!elemento) return;
      categorias.push({
        etiqueta: etiqueta,
        claveIcono: claveIcono || etiqueta,
        titulo: titulo || null,
        elemento: elemento,
        padre: contenedorOriginal || elemento.parentNode,
        siguiente: elemento.nextSibling,
      });
    }

    // Agrupa cada <fieldset class="acordeon-grupo"> de "panel" como su propia
    // categoría (texto e icono salen de su propio título) — lo usa tanto
    // panelAjustes como un panelFiltros que pida categorías individuales en
    // vez de viajar entero como un solo bloque "Filtros".
    function agregarAcordeonesDe(panel) {
      if (!panel) return;
      panel.querySelectorAll('.acordeon-grupo').forEach(function (grupo) {
        var titulo = grupo.querySelector('.acordeon-titulo');
        if (!titulo) return;
        var texto = titulo.textContent.replace(/[▾▴▲▼]/g, '').trim();
        var conClave = titulo.querySelector('[data-i18n]');
        var clave = conClave ? conClave.getAttribute('data-i18n') : null;
        agregarCategoria(texto, grupo, undefined, clave || texto, titulo);
      });
    }

    // Los editores con esqueleto (bloques con data-categoria) toman sus
    // categorías de él (ver armarDesdeEsqueleto, más abajo), no de los
    // acordeones ni de data-app-categoria.
    var modoEsqueleto = !!window.EsqueletoEditor && !!shell.querySelector('[data-categoria]');

    if (modoEsqueleto) {
      // nada: se arman al encender
    } else if (panelFiltros) {
      // El Editor de color tiene "Filtros" y "Textura extra" como dos
      // acordeones separados dentro de .panel-filtros (a diferencia del
      // editor avanzado, donde "Filtro" y "Textura" ya son fieldsets sueltos
      // de .panel-ajustes): si el shell pide categorías individuales, cada
      // acordeón de .panel-filtros se trata igual que los de panelAjustes, en
      // vez de moverse los dos juntos como un solo bloque "Filtros".
      if (panelFiltros.hasAttribute('data-app-categorias-individuales')) agregarAcordeonesDe(panelFiltros);
      else agregarCategoria(txt('filtros'), panelFiltros, undefined, 'Filtros');
    }

    // Una página puede marcar un bloque suelto como categoría propia con
    // data-app-categoria="Nombre". Se usa para herramientas que en computadora
    // viven en la misma columna que otra cosa pero son un bloque aparte y en
    // celular merecen su propio botón (el color del borde y el descargar del
    // generador de stickers, separados de "Filtros"; antes también los
    // filtros de capa del editor avanzado).
    if (!modoEsqueleto) shell.querySelectorAll('[data-app-categoria]').forEach(function (el) {
      var clave = el.getAttribute('data-app-categoria');
      // 'Filtros' y 'Capas' los sabe traducir el modulo. Cualquier otro
      // nombre no tiene traducción propia acá: se lee del título del propio
      // acordeón (que idioma.js ya tradujo), igual que hace el panel de
      // ajustes más abajo — así una categoría nueva no necesita agregarse a
      // mano en este archivo.
      var conocido = { Filtros: 'filtros', Capas: 'capas' }[clave];
      // El título da el nombre y el icono de la categoría. En los editores es
      // el botón del acordeón; en una herramienta que solo agrupa opciones bajo
      // un encabezado (el collage) es ese encabezado.
      var titulo = el.querySelector('.acordeon-titulo') || el.querySelector('h3');
      // La clave de traducción puede estar en un hijo del título (un <span>
      // dentro del botón del acordeón) o en el título mismo (un <h3>).
      var conClave = !titulo ? null
        : (titulo.hasAttribute('data-i18n') ? titulo : titulo.querySelector('[data-i18n]'));
      var claveIcono = conClave ? conClave.getAttribute('data-i18n') : clave;
      var etiqueta = conocido ? txt(conocido) : (titulo ? textoDeTitulo(titulo) : clave);
      agregarCategoria(etiqueta, el, undefined, claveIcono, titulo || undefined);
    });

    if (!modoEsqueleto) {
      agregarAcordeonesDe(panelAjustes);
      if (panelCapas && !panelCapas.hasAttribute('data-app-categorias-individuales')) {
        agregarCategoria(txt('capas'), panelCapas, undefined, 'Capas');
      }
      if (!categorias.length) return;
    }

    var abierta = -1;

    function crearChip(cat, i, icono) {
      var boton = document.createElement('button');
      boton.type = 'button';
      boton.className = 'app-chip';
      var ico = document.createElement('span');
      ico.className = 'app-icono';
      ico.setAttribute('aria-hidden', 'true');
      ico.textContent = icono;
      var nombre = document.createElement('span');
      nombre.textContent = cat.etiqueta;
      cat.nombreChip = nombre;
      boton.appendChild(ico);
      boton.appendChild(nombre);
      boton.setAttribute('aria-pressed', 'false');
      boton.addEventListener('click', function () { alternar(i); });
      cat.chip = boton;
      chips.appendChild(boton);
    }

    categorias.forEach(function (cat, i) {
      crearChip(cat, i, ICONOS_POR_CLAVE[cat.claveIcono] || ICONOS[sinAcentos(cat.claveIcono)] || '🎚️');
    });

    // Con esqueleto (los editores de imagen, fase 7): la barra de abajo es la
    // barra de íconos de la PC acostada (mismos íconos, nombres y orden) y la
    // hoja muestra lo mismo que la columna. Se arma al encender, cuando la
    // página ya conectó su esqueleto.
    var esqueleto = null;
    function armarDesdeEsqueleto() {
      if (esqueleto || !modoEsqueleto) return;
      esqueleto = window.EsqueletoEditor.de(shell);
      if (!esqueleto) return;
      esqueleto.movil.categorias().forEach(function (c) {
        if (c.separadorAntes) chips.appendChild(Object.assign(document.createElement('span'), { className: 'app-chips-sep' }));
        var cat = { clave: c.clave, etiqueta: esqueleto.movil.nombre(c.clave) };
        categorias.push(cat);
        crearChip(cat, categorias.length - 1, esqueleto.movil.icono(c.clave));
      });
      esqueleto.movil.alAbrir(function (clave) {
        var i = categorias.findIndex(function (cat) { return cat.clave === clave; });
        if (i >= 0 && i !== abierta) alternar(i);
      });
    }

    // El texto de cada categoria vuelve a salir de su titulo del acordeon
    // (que la pagina ya tradujo); el icono no, porque su clave es fija.
    document.addEventListener('cambioIdioma', function () {
      chips.setAttribute('aria-label', txt('categorias'));
      categorias.forEach(function (cat) {
        if (cat.clave && esqueleto) {
          cat.etiqueta = esqueleto.movil.nombre(cat.clave);
        } else if (cat.titulo) {
          cat.etiqueta = textoDeTitulo(cat.titulo);
        } else if (cat.claveIcono === 'Filtros' || cat.claveIcono === 'Capas') {
          cat.etiqueta = txt(cat.claveIcono === 'Filtros' ? 'filtros' : 'capas');
        }
        if (cat.nombreChip) cat.nombreChip.textContent = cat.etiqueta;
      });
      // "Herramientas de Imágenes" y "Image Tools" no ocupan lo mismo: el
      // enlace de categoría puede pasar de dos líneas a una.
      medirEncabezado();
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

    function cerrarHoja(notificarTamano) {
      if (abierta < 0) return;
      var cat = categorias[abierta];
      // De vuelta a su lugar exacto en el panel de origen.
      if (cat.mudados) {
        for (var j = cat.mudados.length - 1; j >= 0; j--) {
          var m = cat.mudados[j];
          m.padre.insertBefore(m.el, m.siguiente && m.siguiente.parentNode === m.padre ? m.siguiente : null);
        }
        cat.mudados = null;
      } else if (cat.padre) {
        cat.padre.insertBefore(cat.elemento, cat.siguiente);
      }
      cat.chip.classList.remove('activo');
      cat.chip.setAttribute('aria-pressed', 'false');
      abierta = -1;
      if (notificarTamano !== false) avisarCambioDeTamano();
    }

    function alternar(i) {
      if (abierta === i) { cerrarHoja(); return; }
      cerrarHoja(false);
      var cat = categorias[i];
      var claveIcono = cat.claveIcono;
      if (cat.clave) {
        // Una categoría del esqueleto puede tener varios bloques (Efectos =
        // efectos, resplandor y dispersión): van todos, en su orden, y el
        // primero no repite el nombre (ya está en la barra de abajo).
        var bloques = esqueleto.movil.bloques(cat.clave);
        cat.mudados = bloques.map(function (b) { return { el: b, padre: b.parentNode, siguiente: b.nextSibling }; });
        bloques.forEach(function (b) { hoja.appendChild(b); b.classList.add('abierto'); });
        var conClave = bloques[0] && bloques[0].querySelector('.acordeon-titulo [data-i18n], h3[data-i18n]');
        claveIcono = conClave ? conClave.getAttribute('data-i18n') : cat.clave;
      } else {
        hoja.appendChild(cat.elemento);
      }
      cat.chip.classList.add('activo');
      cat.chip.setAttribute('aria-pressed', 'true');
      // Si la categoría estaba cortada contra el borde, se la trae a la vista.
      if (cat.chip.scrollIntoView) cat.chip.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'nearest' });
      abierta = i;
      if (cat.clave) esqueleto.movil.abrio(cat.clave);
      document.dispatchEvent(new CustomEvent('appMovilCategoriaAbierta', {
        detail: { claveIcono: claveIcono, categoria: cat.clave || null },
      }));
      // Si el grupo venía plegado en la versión de computadora, acá tiene que
      // verse igual: la hoja ya muestra el contenido, pero la clase mantiene
      // coherente el estado por si se vuelve a pantalla ancha.
      if (cat.elemento) cat.elemento.classList.add('abierto');
      // El panel arranca desde arriba aunque la categoría anterior estuviera
      // desplazada.
      hoja.scrollTop = 0;
      avisarCambioDeTamano();
    }

    // --- Acciones siempre visibles ---
    // Se reutilizan los botones que ya existen en la página (mismo id, mismos
    // escuchadores): estos solo los "pulsan". Son un atajo para lo que NO está
    // en ninguna categoría: si el botón ya vive en una (como "Acciones" en los
    // editores y el generador de stickers, o "Exportar" en el collage), no se
    // repite abajo. Antes se repetía y cada página lo tapaba a medias con una
    // marca que solo cubría "Descargar": en el editor de color salían los tres
    // botones dos veces.
    // Los dos editores nombran distinto los mismos botones (el avanzado les
    // agrega "Avanzado"), así que se prueban los nombres posibles y se usa el
    // primero que exista.
    function accion(texto, ids, clase) {
      var destino = null;
      for (var i = 0; i < ids.length && !destino; i++) destino = document.getElementById(ids[i]);
      if (!destino) return null;
      if (categorias.some(function (cat) { return cat.elemento.contains(destino); })) return null;
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

    // Con esqueleto no hacen falta: descargar y "⋯" están en su barra de arriba.
    if (!modoEsqueleto) {
      accion('🖼️', ['botonCambiarImagen'], 'app-secundario');
      accion('↺', ['botonRestablecer', 'botonRestablecerAvanzado'], 'app-secundario');
      accion('⤓ Descargar', ['botonDescargar', 'botonDescargarAvanzado'], 'app-principal');
    }

    // Hoja + categorías + acciones van juntas en un solo bloque flotante,
    // superpuesto ABAJO de la imagen (no empujándola hacia arriba): así la
    // imagen ocupa siempre el mismo alto fijo, se abra o no un panel, y lo
    // que tapa el bloque flotante se alcanza paneando/haciendo zoom en la
    // imagen en vez de que esta se achique cada vez. Ver "encajarEnLaVista"
    // en los editores: al no cambiar más el alto de .panel-lienzo, tampoco
    // hace falta re-ajustar el zoom al abrir/cerrar un panel.
    var controles = document.createElement('div');
    controles.className = 'app-controles';
    controles.appendChild(hoja);
    controles.appendChild(chips);
    // Sin botones (una página sin cambiar-imagen/reiniciar ni el atajo de
    // descarga), la tira quedaría vacía: mejor no mostrar la franja.
    if (acciones.children.length) controles.appendChild(acciones);

    // --- El encabezado del sitio queda a la vista, arriba de la capa ---
    // La capa del editor es "position: fixed", así que no la empuja nada:
    // hay que decirle desde dónde empieza. Se mide el encabezado DESPUÉS de
    // agregar "app-activo" (que es lo que lo compacta), porque si no se mide
    // el alto grande de escritorio y queda una franja vacía.
    var encabezado = document.querySelector('header.sitio');

    function medirEncabezado() {
      if (!encabezado || !cuerpo.classList.contains('app-activo')) return;
      var alto = Math.round(encabezado.getBoundingClientRect().height);
      document.documentElement.style.setProperty('--app-alto-encabezado', alto + 'px');
    }

    // --- Encender / apagar según el ancho ---
    // Con esqueleto, el plan 4.3:
    //   barra superior (la de la PC) · Editando: … · IMAGEN ·
    //   anuncio (franja propia, debajo de la imagen) · hoja · barra de íconos
    var piezasEsqueleto = null;

    function encender() {
      if (cuerpo.classList.contains('app-activo')) return;
      armarDesdeEsqueleto();
      cuerpo.classList.add('app-activo');
      document.documentElement.classList.add('app-activo');
      medirEncabezado();
      if (esqueleto) {
        shell.classList.add('app-con-esqueleto');
        piezasEsqueleto = esqueleto.movil.encender();
        shell.insertBefore(piezasEsqueleto.barra, shell.firstChild);
        if (piezasEsqueleto.editando) shell.insertBefore(piezasEsqueleto.editando, escenario);
        if (piezasEsqueleto.menu) shell.appendChild(piezasEsqueleto.menu);
        shell.insertBefore(tiraAnuncio, escenario.nextSibling);
      } else {
        shell.insertBefore(tiraAnuncio, shell.firstChild);
      }
      if (anuncio) mover(tiraAnuncio, anuncio, null);
      shell.appendChild(controles);
      marcarBordes();
      avisarCambioDeTamano();
    }

    function apagar() {
      if (!cuerpo.classList.contains('app-activo')) return;
      cerrarHoja();
      if (esqueleto && piezasEsqueleto) {
        // La barra, el menú y "Editando" los devuelve el esqueleto.
        esqueleto.movil.apagar();
        piezasEsqueleto = null;
        shell.classList.remove('app-con-esqueleto');
      }
      cuerpo.classList.remove('app-activo');
      document.documentElement.classList.remove('app-activo');
      document.documentElement.style.removeProperty('--app-alto-encabezado');
      // Todo vuelve a donde estaba: la página queda igual que antes.
      if (anuncio && anuncioPadre) mover(anuncioPadre, anuncio, anuncioSiguiente);
      [tiraAnuncio, controles].forEach(function (el) {
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
      if (!cuerpo.classList.contains('app-activo')) return;
      marcarBordes();
      // Girar el teléfono puede cambiar cuántas líneas ocupa el enlace de
      // categoría, y con eso el alto del encabezado.
      medirEncabezado();
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
    // data-app-movil es la forma de marcar el contenedor en una página nueva;
    // las dos clases son las herramientas que ya existían cuando esto se
    // escribió para ellas.
    var shell = document.querySelector('[data-app-movil], .editor-color-layout, .editor-avanzado-layout');
    if (shell) iniciar(shell);
  });
})();

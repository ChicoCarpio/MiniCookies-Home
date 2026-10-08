/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Esqueleto de PC de los editores de imagen (plan del rediseño, hallazgos.md, 4.1,
// 4.2 y fase 6). Con una foto cargada y 901 px o más de ancho:
//
//   ┌ barra superior: ↶ ↷ │ 👁 Original │ zoom │ vista │ ⛶ │   [Descargar] ⋯ ┐
//   ├──┬──────────────────┬──────────────────────────────────────────────────┤
//   │🗂│ Editando: …      │                                                  │
//   │🎯│ ☀️ LUZ           │                  la imagen                       │
//   │☀️│ (una categoría)  │                                                  │
//   │  │                  ├──────────────────────────────────────────────────┤
//   │  │                  │        anuncio: franja propia de 90 px           │
//   └──┴──────────────────┴──────────────────────────────────────────────────┘
//
// El editor ocupa toda la ventana debajo del encabezado del sitio, que queda
// fijo arriba (decisión 4, opción A): el título y la descripción quedan
// arriba, fuera de la vista.
//
// Cómo se usa (el contrato, como el de app-movil.js):
// - Cada bloque de controles se marca con data-categoria="luz" (varios
//   bloques pueden ser de la misma categoría). Los bloques NO se mueven: se
//   muestran solo los de la categoría elegida, en una sola columna. Por eso
//   app-movil.js sigue funcionando igual debajo de 900 px.
// - Lo de la barra superior son los botones REALES de la página (mismos id,
//   mismos escuchadores): se mudan a la barra mientras dure y vuelven a su
//   lugar al pasar a celular.
// - El anuncio de la página se muda a su franja (sin recargarlo: ver
//   AppMovil.moverSinRecargar) y también vuelve.
//
// EsqueletoEditor.conectar({
//   contenedor,            el elemento con las columnas (y el de pantalla completa)
//   editor,                el que se muestra al cargar una foto (por defecto, contenedor)
//   orden: ['luz', '|', …] las categorías en el orden de la barra ('|' = separador)
//   inicio: 'luz',
//   barra: { historial: [botones], original: boton | { abajo, arriba },
//            zoom: { menos, mas, encajar, texto } , resolucion: <select>,
//            pantallaCompleta: boton, descargar: boton, enlace: <a>, menu: [botones],
//            notas: [avisos cortos, a la izquierda de Descargar] },
//   editando: (contenedor) => {}   arma "Editando: …" (solo el Editor avanzado)
//   alCambiarCategoria: (categoria) => {}
// }) -> { abrir(categoria), categoria(), activo(), actualizar() }
//
// En celular (fase 7) el marco lo pone app-movil.js, pero las categorías, la
// barra superior y "Editando: …" salen de acá (EsqueletoEditor.de(contenedor)
// y su .movil): mismos íconos, nombres y orden que en PC. La barra de íconos
// de la PC es la de abajo del celular, y la columna es la hoja que sube.

window.EsqueletoEditor = (function () {
  'use strict';

  var ANCHO_PC = '(min-width: 901px)';

  // Mismo nombre, ícono y orden en los tres editores (plan, 4.1).
  var CATEGORIAS = {
    capas: { ico: '🗂️', es: 'Capas', en: 'Layers' },
    texto: { ico: '🔤', es: 'Texto e imagen', en: 'Text and image' },
    zona: { ico: '🎯', es: 'Zona', en: 'Area' },
    borrar: { ico: '🧽', es: 'Borrar', en: 'Erase' },
    luz: { ico: '☀️', es: 'Luz', en: 'Light' },
    color: { ico: '🎨', es: 'Color', en: 'Color' },
    bn: { ico: '◐', es: 'Blanco y negro', en: 'Black and white' },
    efectos: { ico: '✨', es: 'Efectos', en: 'Effects' },
    filtros: { ico: '🎞️', es: 'Filtros', en: 'Filters' },
    textura: { ico: '🧱', es: 'Textura', en: 'Texture' },
    diseno: { ico: '▦', es: 'Diseño', en: 'Layout' },
    formato: { ico: '⬚', es: 'Formato', en: 'Format' },
    estilo: { ico: '🖌️', es: 'Estilo', en: 'Style' },
    fotos: { ico: '📷', es: 'Fotos', en: 'Photos' },
  };
  var TEXTOS = {
    es: {
      barra: 'Herramientas del editor', categorias: 'Categorías', original: 'Original', 'original-title': 'Mantené presionado para ver el original (\\)',
      menos: 'Alejar', mas: 'Acercar', encajar: 'Encajar', 'encajar-title': 'Encajar la imagen en la vista (Ctrl+0)', mas_opciones: 'Más opciones',
      atajo: 'tecla', 'tab-title': 'Tab esconde y muestra la columna (en pantalla completa)',
      zoom: 'Zoom', resolucion: 'Resolución de la vista previa',
    },
    en: {
      barra: 'Editor tools', categorias: 'Categories', original: 'Original', 'original-title': 'Press and hold to see the original (\\)',
      menos: 'Zoom out', mas: 'Zoom in', encajar: 'Fit', 'encajar-title': 'Fit the image in the view (Ctrl+0)', mas_opciones: 'More options',
      atajo: 'key', 'tab-title': 'Tab hides and shows the column (in full screen)',
      zoom: 'Zoom', resolucion: 'Preview resolution',
    },
  };
  function idioma() { return (window.idiomaActual && window.idiomaActual() === 'en') ? 'en' : 'es'; }
  function txt(clave) { return TEXTOS[idioma()][clave]; }
  function nombre(cat) { return CATEGORIAS[cat] ? CATEGORIAS[cat][idioma()] : cat; }

  function el(tag, clase, texto) {
    var e = document.createElement(tag);
    if (clase) e.className = clase;
    if (texto != null) e.textContent = texto;
    return e;
  }

  // Un anuncio que se mueve de la forma común se recarga y queda en blanco.
  function moverSinRecargar(padre, elemento, referencia) {
    if (window.AppMovil) window.AppMovil.moverSinRecargar(padre, elemento, referencia);
    else padre.insertBefore(elemento, referencia || null);
  }

  function categoriasDe(elemento) { return (elemento.getAttribute('data-categoria') || '').split(/\s+/).filter(Boolean); }

  function conectar(opciones) {
    var contenedor = opciones.contenedor;
    if (!contenedor) return null;
    var editor = opciones.editor || contenedor;
    var barraOp = opciones.barra || {};
    var alCambiar = opciones.alCambiarCategoria || function () {};
    var consulta = window.matchMedia(ANCHO_PC);
    var raiz = document.documentElement;

    // Solo las categorías que esta página tiene, sin separadores sobrantes.
    var orden = [];
    (opciones.orden || []).forEach(function (c) {
      if (c === '|') { if (orden.length && orden[orden.length - 1] !== '|') orden.push('|'); return; }
      if (contenedor.querySelector('[data-categoria~="' + c + '"]')) orden.push(c);
    });
    if (orden[orden.length - 1] === '|') orden.pop();
    var categorias = orden.filter(function (c) { return c !== '|'; });
    if (!categorias.length) return null;

    var actual = categorias.indexOf(opciones.inicio) >= 0 ? opciones.inicio : categorias[0];
    var conColumna = true;
    var activo = false;
    var movil = false; // armado para app-movil.js (celular)
    var abrirEnMovil = null; // lo pone app-movil.js: abre la hoja de una categoría
    var yaSeAcomodo = false; // la página se desplaza sola una vez por foto
    var mudanzas = [];

    // --- Piezas ---
    var barra = el('div', 'ed-barra');
    barra.setAttribute('role', 'toolbar');
    var rail = el('nav', 'ed-rail');
    var cabecera = el('div', 'ed-cabecera');
    var editando = el('div', 'ed-editando');
    var titulo = el('div', 'ed-titulo-categoria');
    var tituloIco = el('span', 'ed-titulo-ico');
    tituloIco.setAttribute('aria-hidden', 'true');
    var tituloTexto = el('span');
    titulo.append(tituloIco, tituloTexto);
    if (opciones.editando) cabecera.appendChild(editando);
    cabecera.appendChild(titulo);
    var filaAnuncio = el('div', 'ed-anuncio-fila');

    // Botones de la barra de íconos.
    var botonesRail = {};
    orden.forEach(function (c) {
      if (c === '|') { rail.appendChild(el('span', 'ed-rail-sep')); return; }
      var b = el('button');
      b.type = 'button';
      b.setAttribute('data-ed-categoria', c);
      var ico = el('span', 'ed-ico', CATEGORIAS[c] ? CATEGORIAS[c].ico : '•');
      ico.setAttribute('aria-hidden', 'true');
      b.appendChild(ico);
      b.addEventListener('click', function () {
        if (c === actual && conColumna) { conColumna = false; aplicar(); return; }
        abrir(c);
      });
      botonesRail[c] = b;
      rail.appendChild(b);
    });

    // "Original": el botón real de la página, o uno que se mantiene apretado.
    var botonOriginal = null;
    if (barraOp.original && !(barraOp.original instanceof Element)) {
      botonOriginal = el('button', 'ed-bi');
      botonOriginal.type = 'button';
      var icoOriginal = el('span', 'ed-ico', '👁');
      icoOriginal.setAttribute('aria-hidden', 'true');
      var textoOriginal = el('span');
      botonOriginal.append(icoOriginal, textoOriginal);
      var mantener = barraOp.original;
      botonOriginal.addEventListener('pointerdown', function (e) { e.preventDefault(); mantener.abajo(); });
      ['pointerup', 'pointerleave', 'pointercancel'].forEach(function (ev) { botonOriginal.addEventListener(ev, function () { mantener.arriba(); }); });
      botonOriginal.addEventListener('keydown', function (e) { if ((e.key === ' ' || e.key === 'Enter') && !e.repeat) { e.preventDefault(); mantener.abajo(); } });
      botonOriginal.addEventListener('keyup', function (e) { if (e.key === ' ' || e.key === 'Enter') mantener.arriba(); });
      botonOriginal.addEventListener('contextmenu', function (e) { e.preventDefault(); });
    }

    // Zoom: −, el porcentaje, + y Encajar. En celular la barra no tiene ancho
    // para todo: queda Encajar, y −, % y + van en una fila del menú "⋯" (el
    // pellizco no puede ser la única forma de acercar).
    var zoom = null;
    if (barraOp.zoom) {
      zoom = { menos: el('button', 'ed-bi', '−'), texto: el('span', 'ed-zoom-texto'), mas: el('button', 'ed-bi', '+'), encajar: el('button', 'ed-bi') };
      zoom.menos.type = zoom.mas.type = zoom.encajar.type = 'button';
      var icoEncajar = el('span', 'ed-ico', '⤢');
      icoEncajar.setAttribute('aria-hidden', 'true');
      zoom.encajarTexto = el('span');
      zoom.encajar.append(icoEncajar, zoom.encajarTexto);
      zoom.menos.addEventListener('click', function () { barraOp.zoom.menos(); actualizar(); });
      zoom.mas.addEventListener('click', function () { barraOp.zoom.mas(); actualizar(); });
      zoom.encajar.addEventListener('click', function () { barraOp.zoom.encajar(); actualizar(); });
      zoom.fila = el('div', 'ed-menu-fila');
      zoom.rotulo = el('span', 'ed-menu-rotulo');
    }
    // La resolución de la vista previa, con su rótulo, en el menú de celular.
    var etiquetaResolucion = barraOp.resolucion ? el('label', 'ed-menu-campo') : null;
    var rotuloResolucion = el('span');
    if (etiquetaResolucion) etiquetaResolucion.appendChild(rotuloResolucion);

    // "⋯": un menú que vive dentro del editor (así se ve en pantalla completa).
    var botonMenu = null, menu = null;
    if ((barraOp.menu && barraOp.menu.length) || zoom || etiquetaResolucion) {
      botonMenu = el('button', 'ed-bi ed-boton-menu', '⋯');
      botonMenu.type = 'button';
      botonMenu.setAttribute('aria-haspopup', 'true');
      botonMenu.setAttribute('aria-expanded', 'false');
      menu = el('div', 'ed-menu');
      menu.hidden = true;
      menu.setAttribute('role', 'menu');
      var cerrarMenu = function () { menu.hidden = true; botonMenu.setAttribute('aria-expanded', 'false'); };
      botonMenu.addEventListener('click', function () {
        menu.hidden = !menu.hidden;
        botonMenu.setAttribute('aria-expanded', String(!menu.hidden));
        if (!menu.hidden) { var primero = menu.querySelector('button:not([disabled])'); if (primero) primero.focus(); }
      });
      // Solo las opciones lo cierran: el − y el + del zoom no.
      menu.addEventListener('click', function (e) { var b = e.target.closest('button'); if (b && b.parentNode === menu) cerrarMenu(); });
      document.addEventListener('pointerdown', function (e) { if (!menu.hidden && !menu.contains(e.target) && e.target !== botonMenu) cerrarMenu(); });
      document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && !menu.hidden) { cerrarMenu(); botonMenu.focus(); } });
    }

    // Descargar desde la barra baja el archivo directamente: el botón arma el
    // enlace (como siempre) y acá se lo toca. Solo con un clic de verdad
    // (isTrusted): las pruebas automáticas tocan el botón y leen el enlace,
    // sin bajar nada.
    var esperandoEnlace = false;
    if (barraOp.descargar && barraOp.enlace) {
      barraOp.descargar.addEventListener('click', function (e) { if ((activo || movil) && e.isTrusted) esperandoEnlace = true; }, true);
      new MutationObserver(function () {
        if (!esperandoEnlace || !/^blob:/.test(barraOp.enlace.href)) return;
        esperandoEnlace = false;
        barraOp.enlace.click();
      }).observe(barraOp.enlace, { attributes: true, attributeFilter: ['href'] });
    }

    function textos() {
      barra.setAttribute('aria-label', txt('barra'));
      rail.setAttribute('aria-label', txt('categorias'));
      var n = 0;
      categorias.forEach(function (c) {
        n += 1;
        var b = botonesRail[c];
        var tecla = n <= 10 ? ' (' + txt('atajo') + ' ' + (n % 10) + ')' : '';
        b.title = nombre(c) + tecla;
        b.setAttribute('aria-label', nombre(c));
      });
      if (botonOriginal) {
        botonOriginal.lastChild.textContent = txt('original');
        botonOriginal.title = txt('original-title');
      }
      if (zoom) {
        zoom.menos.title = txt('menos'); zoom.menos.setAttribute('aria-label', txt('menos'));
        zoom.mas.title = txt('mas'); zoom.mas.setAttribute('aria-label', txt('mas'));
        zoom.encajarTexto.textContent = txt('encajar'); zoom.encajar.title = txt('encajar-title');
        zoom.encajar.setAttribute('aria-label', txt('encajar'));
        zoom.rotulo.textContent = txt('zoom');
      }
      if (botonOriginal) botonOriginal.setAttribute('aria-label', txt('original'));
      rotuloResolucion.textContent = txt('resolucion');
      if (botonMenu) { botonMenu.title = txt('mas_opciones'); botonMenu.setAttribute('aria-label', txt('mas_opciones')); }
      tituloTexto.textContent = nombre(actual);
    }

    // --- Mostrar una categoría ---
    function aplicar() {
      contenedor.classList.toggle('ed-sin-columna', !conColumna);
      var paneles = new Set();
      var primero = true;
      contenedor.querySelectorAll('[data-categoria]').forEach(function (bloque) {
        var visible = activo && categoriasDe(bloque).indexOf(actual) >= 0;
        bloque.classList.toggle('ed-cat-visible', visible);
        // El primero no repite el título: ya está arriba de la columna.
        bloque.classList.toggle('ed-cat-primero', visible && primero);
        if (!visible) return;
        primero = false;
        // El panel (hijo directo del contenedor) que lo tiene.
        var p = bloque;
        while (p.parentElement && p.parentElement !== contenedor) p = p.parentElement;
        paneles.add(p);
      });
      Array.prototype.forEach.call(contenedor.children, function (hijo) { hijo.classList.toggle('ed-panel-visible', paneles.has(hijo)); });
      categorias.forEach(function (c) { botonesRail[c].setAttribute('aria-pressed', String(c === actual && conColumna)); });
      tituloIco.textContent = CATEGORIAS[actual] ? CATEGORIAS[actual].ico : '';
      tituloTexto.textContent = nombre(actual);
      if (activo) window.dispatchEvent(new Event('resize'));
    }

    function abrir(cat) {
      if (categorias.indexOf(cat) < 0) return;
      // En celular, abrir una categoría es abrir su hoja (app-movil.js).
      if (movil) { if (abrirEnMovil) abrirEnMovil(cat); return; }
      var cambio = cat !== actual;
      actual = cat;
      conColumna = true;
      aplicar();
      if (cambio || activo) alCambiar(cat);
    }

    function actualizar() {
      if (zoom && barraOp.zoom.texto) zoom.texto.textContent = barraOp.zoom.texto();
      if (opciones.editando && (activo || movil)) opciones.editando(editando, actual);
    }

    // --- Encender y apagar ---
    function mudar(elemento, destino) {
      if (!elemento) return;
      mudanzas.push({ elemento: elemento, padre: elemento.parentNode, siguiente: elemento.nextSibling });
      destino.appendChild(elemento);
    }
    function separador() { barra.appendChild(el('span', 'ed-sep')); }

    // La misma barra en PC y en celular, con lo mismo en el mismo orden. En
    // celular (390 px) no entra todo: −, %, + y la resolución pasan al menú
    // "⋯", junto con los avisos; Original y Encajar quedan solo con su ícono.
    function armarBarra(enCelular) {
      barra.textContent = '';
      barra.classList.toggle('ed-barra-celular', !!enCelular);
      if (menu) menu.textContent = '';
      (barraOp.historial || []).forEach(function (b) { mudar(b, barra); });
      separador();
      if (barraOp.original instanceof Element) mudar(barraOp.original, barra);
      else if (botonOriginal) barra.appendChild(botonOriginal);
      if (zoom && !enCelular) barra.append(zoom.menos, zoom.texto, zoom.mas, zoom.encajar);
      if (zoom && enCelular) {
        barra.appendChild(zoom.encajar);
        zoom.fila.textContent = '';
        zoom.fila.append(zoom.rotulo, zoom.menos, zoom.texto, zoom.mas);
        menu.appendChild(zoom.fila);
      }
      if (barraOp.resolucion && !enCelular) mudar(barraOp.resolucion, barra);
      if (barraOp.resolucion && enCelular) {
        menu.appendChild(etiquetaResolucion);
        mudar(barraOp.resolucion, etiquetaResolucion);
      }
      if (barraOp.pantallaCompleta) mudar(barraOp.pantallaCompleta, barra);
      barra.appendChild(el('span', 'ed-espacio'));
      // Avisos cortos que dependen de lo de la barra ("Sale en PNG…").
      (barraOp.notas || []).forEach(function (n) { mudar(n, enCelular ? menu : barra); });
      if (barraOp.descargar) {
        barraOp.descargar.classList.add('ed-descargar');
        mudar(barraOp.descargar, barra);
      }
      if (botonMenu && (enCelular || (barraOp.menu && barraOp.menu.length))) {
        (barraOp.menu || []).forEach(function (b) { mudar(b, menu); });
        barra.appendChild(botonMenu);
      }
    }

    // Todo lo mudado vuelve a su lugar, en orden inverso (cada uno antes de
    // su vecino).
    function devolver() {
      for (var i = mudanzas.length - 1; i >= 0; i--) {
        var m = mudanzas[i];
        if (m.anuncio) {
          // Si app-movil.js ya se lo llevó a su tira, es suyo.
          if (filaAnuncio.contains(m.elemento)) moverSinRecargar(m.padre, m.elemento, m.siguiente && m.siguiente.parentNode === m.padre ? m.siguiente : null);
          continue;
        }
        m.padre.insertBefore(m.elemento, m.siguiente && m.siguiente.parentNode === m.padre ? m.siguiente : null);
      }
      mudanzas = [];
    }

    function medir() {
      var encabezado = document.querySelector('header.sitio');
      raiz.style.setProperty('--ed-alto-encabezado', (encabezado ? Math.round(encabezado.getBoundingClientRect().height) : 0) + 'px');
      raiz.style.setProperty('--ed-barra-desplazamiento', Math.max(0, window.innerWidth - raiz.clientWidth) + 'px');
    }

    function encender() {
      if (activo) return;
      if (movil) apagarMovil();
      activo = true;
      document.body.classList.add('ed-esqueleto-activo');
      contenedor.classList.add('esqueleto-pc');
      // Los paneles plegables de antes no existen en el esqueleto.
      contenedor.classList.remove('izq-plegado', 'der-plegado');
      medir();
      contenedor.insertBefore(cabecera, contenedor.firstChild);
      contenedor.insertBefore(rail, contenedor.firstChild);
      contenedor.insertBefore(barra, contenedor.firstChild);
      contenedor.appendChild(filaAnuncio);
      if (menu) contenedor.appendChild(menu);
      armarBarra(false);
      var anuncio = document.querySelector('.anuncio');
      if (anuncio && !anuncio.closest('.app-anuncio')) {
        mudanzas.push({ elemento: anuncio, padre: anuncio.parentNode, siguiente: anuncio.nextSibling, anuncio: true });
        moverSinRecargar(filaAnuncio, anuncio, null);
      }
      textos();
      aplicar();
      actualizar();
      alCambiar(actual);
      if (!yaSeAcomodo) {
        yaSeAcomodo = true;
        // El editor justo debajo del encabezado, que queda fijo arriba.
        var arriba = contenedor.getBoundingClientRect().top + window.scrollY - parseFloat(getComputedStyle(raiz).getPropertyValue('--ed-alto-encabezado') || 0);
        window.scrollTo({ top: Math.max(0, arriba), behavior: 'instant' });
      }
    }

    function apagar() {
      if (!activo) return;
      activo = false;
      // Al volver a PC (de celular, o con otra foto) se vuelve a acomodar.
      yaSeAcomodo = false;
      if (menu) menu.hidden = true;
      devolver();
      [barra, rail, cabecera, filaAnuncio, menu].forEach(function (pieza) { if (pieza && pieza.parentNode) pieza.parentNode.removeChild(pieza); });
      contenedor.classList.remove('esqueleto-pc', 'ed-sin-columna');
      document.body.classList.remove('ed-esqueleto-activo');
      aplicar();
      // La página devuelve lo que haya mudado por su cuenta.
      alCambiar(null);
      window.dispatchEvent(new Event('resize'));
    }

    // --- Celular: app-movil.js pone el marco y pide las piezas ---
    // Arma la barra (con los botones reales mudados) y "Editando: …", y las
    // devuelve para que app-movil.js las ubique. Las categorías no se mudan
    // acá: app-movil.js mueve sus bloques a la hoja al abrirla.
    function encenderMovil() {
      if (activo) apagar();
      if (!movil) {
        movil = true;
        document.body.classList.add('ed-esqueleto-movil');
        armarBarra(true);
        textos();
        actualizar();
        // Lo que la página muda por su cuenta (los controles del pincel a
        // "Borrar", la tira de fotos a "Fotos"), igual que en PC.
        alCambiar(actual);
      }
      return { barra: barra, editando: opciones.editando ? editando : null, menu: menu };
    }

    function apagarMovil() {
      if (!movil) return;
      movil = false;
      if (menu) menu.hidden = true;
      devolver();
      [barra, menu].forEach(function (pieza) { if (pieza && pieza.parentNode) pieza.parentNode.removeChild(pieza); });
      // "Editando" vuelve arriba de la columna de PC.
      if (opciones.editando) cabecera.insertBefore(editando, titulo);
      document.body.classList.remove('ed-esqueleto-movil');
      alCambiar(null);
    }

    // app-movil.js avisa qué hoja abrió: es la categoría elegida, también
    // para cuando se vuelva a PC.
    function movilAbrio(cat) {
      if (!movil || categorias.indexOf(cat) < 0) return;
      actual = cat;
      alCambiar(cat);
      actualizar();
    }

    function editorVisible() { return editor.style.display !== 'none' && editor.getClientRects().length > 0; }
    function revisar() {
      if (!editorVisible()) yaSeAcomodo = false;
      if (consulta.matches && editorVisible()) {
        // Al pasar de celular a PC, app-movil.js primero devuelve lo suyo
        // (el anuncio, la hoja): el esqueleto se arma después.
        setTimeout(function () { if (consulta.matches && editorVisible()) encender(); }, 0);
      } else {
        apagar();
      }
    }
    if (consulta.addEventListener) consulta.addEventListener('change', revisar);
    else consulta.addListener(revisar);
    new MutationObserver(revisar).observe(editor, { attributes: true, attributeFilter: ['style', 'class', 'hidden'] });
    window.addEventListener('resize', function () { if (activo) medir(); });
    // La lista de "Editando: …" se cierra al tocar afuera o con Esc.
    document.addEventListener('pointerdown', function (e) {
      var lista = editando.querySelector('.ed-editando-lista');
      if (lista && !lista.hidden && !editando.contains(e.target)) lista.hidden = true;
    });
    document.addEventListener('keydown', function (e) {
      var lista = editando.querySelector('.ed-editando-lista');
      if (e.key === 'Escape' && lista && !lista.hidden) lista.hidden = true;
    });
    document.addEventListener('cambioIdioma', function () {
      if (activo) { textos(); actualizar(); medir(); }
      if (movil) { textos(); actualizar(); }
    });

    // --- Teclado: 1–9 y 0 eligen categoría; Tab esconde la columna, solo
    // en pantalla completa (afuera, Tab sigue recorriendo la página). ---
    if (window.Atajos) {
      var numeros = {};
      categorias.slice(0, 10).forEach(function (c, i) { numeros[String((i + 1) % 10)] = function () { abrir(c); }; });
      window.Atajos.conectar(numeros, function () { return activo; });
      window.Atajos.conectar({ tab: function () { conColumna = !conColumna; aplicar(); } }, function () {
        return activo && document.fullscreenElement === contenedor;
      });
    }

    revisar();
    var api = {
      abrir: abrir,
      categoria: function () { return actual; },
      // En PC o en celular: hay una categoría elegida que se puede abrir.
      activo: function () { return activo || movil; },
      columnaVisible: function () { return activo && conColumna; },
      actualizar: actualizar,
      // Lo que usa app-movil.js (ver arriba, "Celular").
      movil: {
        categorias: function () {
          return orden.map(function (c, i) {
            if (c === '|') return null;
            return { clave: c, separadorAntes: orden[i - 1] === '|' };
          }).filter(Boolean);
        },
        bloques: function (cat) { return Array.prototype.slice.call(contenedor.querySelectorAll('[data-categoria~="' + cat + '"]')); },
        icono: function (cat) { return CATEGORIAS[cat] ? CATEGORIAS[cat].ico : '•'; },
        nombre: nombre,
        encender: encenderMovil,
        apagar: apagarMovil,
        abrio: movilAbrio,
        alAbrir: function (fn) { abrirEnMovil = fn; },
      },
    };
    instancias.push({ contenedor: contenedor, api: api });
    return api;
  }

  // El esqueleto conectado a un contenedor (para app-movil.js).
  var instancias = [];
  function de(elemento) {
    for (var i = 0; i < instancias.length; i++) if (instancias[i].contenedor === elemento) return instancias[i].api;
    return null;
  }

  return { conectar: conectar, de: de, CATEGORIAS: CATEGORIAS };
})();

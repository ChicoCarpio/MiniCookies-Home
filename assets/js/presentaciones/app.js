/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Mini Cookies Presentaciones (fase 5 del PLAN, etapas 1 a 5): la aplicación. Une la presentación editable (presentacion-editable.js), el modelo y el
// dibujo del motor de PowerPoint (pptx-modelo.js, pptx-dibujo.js), el historial, los comandos y la interfaz G (nucleo/oficina.js).
//   Etapa 1: abrir, crear en blanco, panel de miniaturas, escenario a escala, descarga y autoguardado.
//   Etapa 2: seleccionar, mover, redimensionar y girar objetos (tiradores en el DOM encima del lienzo), teclado, inspector con controles, orden.
//   Etapa 3: agregar, duplicar, borrar y reordenar diapositivas (arrastre con ratón y con el dedo, o con comandos) y modo presentación.
//   Etapas 4 y 5: editar el texto de las formas (cuadro nativo encima), su formato, insertar cuadro de texto e imagen, guías y PDF.
(function () {
  'use strict';
  // ---- textos ----------------------------------------------------------------------------------------------------
  const ES = {
    'titulo-pagina': 'Editar PowerPoint (.pptx) en el navegador — Mini Cookies', 'meta-descripcion': 'Abre y edita presentaciones de PowerPoint (.pptx) en el navegador: mueve objetos, ordena diapositivas y presenta, sin instalar nada y sin subir tu archivo. Gratis y sin registro.',
    'h1': 'Editar PowerPoint (Presentaciones)', 'intro': 'Abre una presentación de PowerPoint, mueve sus objetos, ordena las diapositivas y preséntala. Al guardar, lo que no tocaste queda exactamente como estaba.',
    'zona-arrastre': '🎞️ Arrastra tu archivo aquí o haz clic para elegirlo (.pptx)', 'boton-nuevo': 'Presentación en blanco', 'cargando': 'Abriendo la presentación…',
    'borradores': 'Trabajo sin terminar en este navegador', 'recuperar': 'Recuperar', 'descartar': 'Descartar', 'nombre-app': 'Presentaciones', 'volver-inicio': 'Volver al inicio',
    'nombre-archivo': 'Nombre del archivo', 'mas-opciones': 'Más opciones', 'cinta': 'Cinta', 'inspector': 'Propiedades', 'of-herramientas': 'Herramientas',
    'tab-inicio': 'Inicio', 'tab-vista': 'Vista', 'g-diapositivas': 'Diapositivas', 'g-archivo': 'Archivo', 'g-ventana': 'Ventana', 'g-historial': 'Deshacer', 'g-orden': 'Orden', 'g-presentar': 'Presentar',
    'cmd-deshacer': 'Deshacer', 'cmd-rehacer': 'Rehacer', 'cmd-descargar': 'Descargar la presentación', 'cmd-descargar-corto': 'Descargar', 'cmd-nuevo': 'Presentación nueva', 'cmd-abrir': 'Abrir otro archivo',
    'cmd-paleta': 'Buscar un comando', 'cmd-completa': 'Pantalla completa', 'cmd-modo-editar': 'Editar', 'cmd-modo-leer': 'Leer', 'cmd-inspector': 'Panel de propiedades',
    'cmd-anterior': 'Diapositiva anterior', 'cmd-siguiente': 'Diapositiva siguiente', 'cmd-primera': 'Primera diapositiva', 'cmd-ultima': 'Última diapositiva', 'cmd-panel': 'Panel de diapositivas',
    'cmd-nueva-diap': 'Diapositiva nueva', 'cmd-duplicar-diap': 'Duplicar la diapositiva', 'cmd-borrar-diap': 'Borrar la diapositiva', 'cmd-subir-diap': 'Subir la diapositiva', 'cmd-bajar-diap': 'Bajar la diapositiva',
    'cmd-presentar': 'Presentar', 'cmd-presentar-aqui': 'Presentar desde esta diapositiva', 'cmd-eliminar-forma': 'Eliminar el objeto', 'cmd-frente': 'Traer al frente', 'cmd-fondo': 'Enviar al fondo', 'cmd-deseleccionar': 'Quitar la selección',
    'guardado-local': 'Guardado en este navegador a las {0}', 'sin-cambios': 'Sin cambios', 'cambios-sin-descargar': 'Cambios sin descargar', 'descargado': 'Descargado',
    'honesta-edita': 'Aquí se editan los objetos (posición, tamaño, giro, relleno, línea), el texto de las formas y las diapositivas.', 'honesta-conserva': 'Se conserva al guardar, sin editar aquí: {0}.', 'honesta-motivos': 'En la diapositiva {0} hay cosas que se dibujan de forma aproximada: {1}.',
    'x-notas': 'notas del orador', 'x-comentarios': 'comentarios', 'x-animaciones': 'animaciones', 'x-transiciones': 'transiciones', 'x-graficos': 'gráficos', 'x-smartart': 'diagramas SmartArt', 'x-incrustados': 'objetos incrustados', 'x-multimedia': 'audio y video', 'x-macros': 'macros', 'x-tablas': 'el texto de tablas y grupos',
    'cmd-editar-texto': 'Editar el texto', 'cmd-negrita': 'Negrita', 'cmd-cursiva': 'Cursiva', 'cmd-subrayado': 'Subrayado', 'cmd-tamano-mas': 'Aumentar el tamaño del texto', 'cmd-tamano-menos': 'Disminuir el tamaño del texto', 'cmd-color-texto': 'Color del texto',
    'cmd-insertar-texto': 'Cuadro de texto', 'cmd-insertar-forma': 'Formas', 'cmd-insertar-imagen': 'Imagen', 'cmd-pdf': 'Descargar en PDF', 'g-texto': 'Texto', 'g-insertar': 'Insertar', 'aviso-pdf': 'Preparando el PDF…', 'aviso-pdf-listo': 'PDF listo: {0} diapositivas', 'aviso-imagen': 'Elige una imagen PNG, JPEG o GIF', 'texto-nuevo': 'Texto', 'g-parrafo': 'Párrafo', 'cmd-vinetas': 'Viñetas', 'cmd-numeracion': 'Numeración', 'cmd-nivel-mas': 'Aumentar el nivel de la lista', 'cmd-nivel-menos': 'Disminuir el nivel de la lista', 'hint-titulo': 'Haz clic para agregar título', 'hint-subtitulo': 'Haz clic para agregar subtítulo', 'hint-texto': 'Haz clic para agregar texto',
    'i-diapositiva': 'Diapositiva', 'i-numero': 'Número', 'i-tamano': 'Tamaño', 'i-formas': 'Formas', 'i-archivo': 'El archivo', 'i-diapositivas': 'Diapositivas', 'i-oculta': 'Oculta',
    'i-objeto': 'Objeto', 'i-nombre': 'Nombre', 'i-tipo': 'Tipo', 'i-x': 'Izquierda (cm)', 'i-y': 'Arriba (cm)', 'i-ancho': 'Ancho (cm)', 'i-alto': 'Alto (cm)', 'i-giro': 'Giro (°)', 'i-relleno': 'Relleno', 'i-linea': 'Línea', 'i-quitar': 'Quitar',
    'obj-forma': 'Forma', 'obj-imagen': 'Imagen', 'obj-grupo': 'Grupo', 'obj-tabla': 'Tabla', 'obj-grafico': 'Gráfico', 'obj-otro': 'Objeto',
    'si': 'Sí', 'no': 'No', 'diapositiva-n': 'Diapositiva {0}', 'diapositiva-de': 'Diapositiva {0} de {1}',
    'aviso-no-se-pudo': 'No se pudo hacer', 'aviso-no-editable': 'Este objeto se conserva al guardar, pero no se edita aquí', 'aviso-ultima': 'Una presentación necesita al menos una diapositiva',
    'pres-salir': 'Salir de la presentación', 'pres-fin': 'Fin de la presentación', 'pres-n': '{0} / {1}',
    'error-abrir': 'No se pudo abrir el archivo', 'salir-sin-descargar': 'Hay cambios sin descargar.',
  };
  const EN = {
    'titulo-pagina': 'Edit PowerPoint (.pptx) files in your browser — Mini Cookies', 'meta-descripcion': 'Open and edit PowerPoint presentations (.pptx) in your browser: move objects, reorder slides and present, without installing anything or uploading your file. Free, no sign-up.',
    'h1': 'Presentations', 'intro': 'Open a PowerPoint presentation, move its objects, reorder the slides and present it. When you save, everything you didn’t touch stays exactly as it was.',
    'zona-arrastre': '🎞️ Drop your file here or click to choose it (.pptx)', 'boton-nuevo': 'Blank presentation', 'cargando': 'Opening the presentation…',
    'borradores': 'Unfinished work in this browser', 'recuperar': 'Recover', 'descartar': 'Discard', 'nombre-app': 'Presentations', 'volver-inicio': 'Back to start',
    'nombre-archivo': 'File name', 'mas-opciones': 'More options', 'cinta': 'Ribbon', 'inspector': 'Properties', 'of-herramientas': 'Tools',
    'tab-inicio': 'Home', 'tab-vista': 'View', 'g-diapositivas': 'Slides', 'g-archivo': 'File', 'g-ventana': 'Window', 'g-historial': 'Undo', 'g-orden': 'Order', 'g-presentar': 'Present',
    'cmd-deshacer': 'Undo', 'cmd-rehacer': 'Redo', 'cmd-descargar': 'Download the presentation', 'cmd-descargar-corto': 'Download', 'cmd-nuevo': 'New presentation', 'cmd-abrir': 'Open another file',
    'cmd-paleta': 'Search for a command', 'cmd-completa': 'Full screen', 'cmd-modo-editar': 'Edit', 'cmd-modo-leer': 'Read', 'cmd-inspector': 'Properties panel',
    'cmd-anterior': 'Previous slide', 'cmd-siguiente': 'Next slide', 'cmd-primera': 'First slide', 'cmd-ultima': 'Last slide', 'cmd-panel': 'Slide panel',
    'cmd-nueva-diap': 'New slide', 'cmd-duplicar-diap': 'Duplicate the slide', 'cmd-borrar-diap': 'Delete the slide', 'cmd-subir-diap': 'Move the slide up', 'cmd-bajar-diap': 'Move the slide down',
    'cmd-presentar': 'Present', 'cmd-presentar-aqui': 'Present from this slide', 'cmd-eliminar-forma': 'Delete the object', 'cmd-frente': 'Bring to front', 'cmd-fondo': 'Send to back', 'cmd-deseleccionar': 'Clear the selection',
    'guardado-local': 'Saved in this browser at {0}', 'sin-cambios': 'No changes', 'cambios-sin-descargar': 'Changes not downloaded', 'descargado': 'Downloaded',
    'honesta-edita': 'Here you edit objects (position, size, rotation, fill, line), the text of shapes and the slides.', 'honesta-conserva': 'Kept as is when saving, not editable here: {0}.', 'honesta-motivos': 'Slide {0} has things that are drawn approximately: {1}.',
    'x-notas': 'speaker notes', 'x-comentarios': 'comments', 'x-animaciones': 'animations', 'x-transiciones': 'transitions', 'x-graficos': 'charts', 'x-smartart': 'SmartArt diagrams', 'x-incrustados': 'embedded objects', 'x-multimedia': 'audio and video', 'x-macros': 'macros', 'x-tablas': 'the text of tables and groups',
    'cmd-editar-texto': 'Edit the text', 'cmd-negrita': 'Bold', 'cmd-cursiva': 'Italic', 'cmd-subrayado': 'Underline', 'cmd-tamano-mas': 'Increase text size', 'cmd-tamano-menos': 'Decrease text size', 'cmd-color-texto': 'Text color',
    'cmd-insertar-texto': 'Text box', 'cmd-insertar-forma': 'Shapes', 'cmd-insertar-imagen': 'Picture', 'cmd-pdf': 'Download as PDF', 'g-texto': 'Text', 'g-insertar': 'Insert', 'aviso-pdf': 'Preparing the PDF…', 'aviso-pdf-listo': 'PDF ready: {0} slides', 'aviso-imagen': 'Choose a PNG, JPEG or GIF picture', 'texto-nuevo': 'Text', 'g-parrafo': 'Paragraph', 'cmd-vinetas': 'Bullets', 'cmd-numeracion': 'Numbering', 'cmd-nivel-mas': 'Increase list level', 'cmd-nivel-menos': 'Decrease list level', 'hint-titulo': 'Click to add title', 'hint-subtitulo': 'Click to add subtitle', 'hint-texto': 'Click to add text',
    'i-diapositiva': 'Slide', 'i-numero': 'Number', 'i-tamano': 'Size', 'i-formas': 'Shapes', 'i-archivo': 'The file', 'i-diapositivas': 'Slides', 'i-oculta': 'Hidden',
    'i-objeto': 'Object', 'i-nombre': 'Name', 'i-tipo': 'Type', 'i-x': 'Left (cm)', 'i-y': 'Top (cm)', 'i-ancho': 'Width (cm)', 'i-alto': 'Height (cm)', 'i-giro': 'Rotation (°)', 'i-relleno': 'Fill', 'i-linea': 'Line', 'i-quitar': 'None',
    'obj-forma': 'Shape', 'obj-imagen': 'Picture', 'obj-grupo': 'Group', 'obj-tabla': 'Table', 'obj-grafico': 'Chart', 'obj-otro': 'Object',
    'si': 'Yes', 'no': 'No', 'diapositiva-n': 'Slide {0}', 'diapositiva-de': 'Slide {0} of {1}',
    'aviso-no-se-pudo': 'Couldn’t do it', 'aviso-no-editable': 'This object is kept when saving, but it can’t be edited here', 'aviso-ultima': 'A presentation needs at least one slide',
    'pres-salir': 'Exit the presentation', 'pres-fin': 'End of the presentation', 'pres-n': '{0} / {1}',
    'error-abrir': 'The file could not be opened', 'salir-sin-descargar': 'There are changes not downloaded.',
  };
  const pagina = window.i18nPagina || (window.i18nPagina = {});
  pagina.es = Object.assign(pagina.es || {}, ES);
  pagina.en = Object.assign(pagina.en || {}, EN);
  const tx = (k, ...a) => { let s = typeof window.t === 'function' ? window.t(k) : (ES[k] || k); a.forEach((x, i) => { s = s.replace('{' + i + '}', x); }); return s; };
  const idioma = () => (typeof window.idiomaActual === 'function' && window.idiomaActual() === 'en' ? 'en' : 'es');

  // ---- estado ------------------------------------------------------------------------------------------------------
  const $ = (id) => document.getElementById(id);
  const EMU_POR_CM = 360000, EMU_POR_PX = 9525, MAX_EMU = 51206400 * 4, MINIMO = 22860;
  const crearEl = Oficina.crearEl;
  let pe = null, modelo = null, hist = null, docId = null, nombreDoc = 'Presentación', lectura = false, actual = 0, cantidad = 0;
  let diapActual = null, sel = null, mostrada = -1;   // sel: { clave, camino, bloqueada }
  const autoguardado = Autoguardado.crear('presentaciones');
  const comandos = Comandos.crear();

  const panel = crearEl('div', { class: 'pr-panel', id: 'panelDiapositivas', role: 'listbox', 'aria-label': tx('cmd-panel'), 'data-i18n-attr': 'aria-label:cmd-panel' });
  const lienzoPrincipal = crearEl('canvas', { id: 'lienzoDiapositiva', role: 'img' });
  // la capa de interacción: la caja de selección con sus tiradores va en el DOM encima del lienzo
  const ASAS = [['nw', -1, -1], ['n', 0, -1], ['ne', 1, -1], ['e', 1, 0], ['se', 1, 1], ['s', 0, 1], ['sw', -1, 1], ['w', -1, 0]];
  const caja = crearEl('div', { class: 'pr-caja', hidden: true }, [
    ...ASAS.map(([n]) => crearEl('span', { class: 'pr-asa pr-asa-' + n, dataset: { asa: n } })),
    crearEl('span', { class: 'pr-asa pr-asa-giro', dataset: { asa: 'giro' } }),
  ]);
  const hints = crearEl('div', { class: 'pr-hints', 'aria-hidden': 'true' });
  const sobre = crearEl('div', { class: 'pr-sobre', id: 'sobreDiapositiva', tabindex: '0' }, [hints, caja]);
  const marco = crearEl('div', { class: 'pr-diapositiva', id: 'diapositiva' }, [lienzoPrincipal, sobre]);
  const escenario = crearEl('div', { class: 'pr-escenario', id: 'escenario' }, [marco]);
  const interfaz = Oficina.montar($('app'), { nombre: 'nombre-app', inicio: '../', riel: false });
  interfaz.lienzo.append(crearEl('div', { class: 'pr-cuerpo' }, [panel, escenario]));
  const insp = Oficina.inspector(interfaz.inspector);

  // ---- abrir -----------------------------------------------------------------------------------------------------
  async function abrirArchivo(archivo) {
    mostrarCargando(true);
    try { await abrirBytes(new Uint8Array(await archivo.arrayBuffer()), archivo.name.replace(/\.[^.]+$/, '')); } catch (e) { mostrarError(e); }
    mostrarCargando(false);
  }
  async function abrirBytes(bytes, nombre, opciones) {
    const o = opciones || {};
    const nuevo = await PresentacionEditable.abrir(bytes);
    const m = await PptxModelo.abrir(null, { zip: nuevo.zip });
    pe = nuevo; modelo = m; cantidad = m.cantidad; actual = 0; sel = null; mostrada = -1; diapActual = null; rutasAntes = m.rutas.join('|'); tocadas = new Set(); pedirIr = null;
    nombreDoc = nombre || 'Presentación';
    docId = o.id || Autoguardado.nuevoId();
    hist = HistorialDoc.crear({ aplicar: aplicarOp, alCambiar: alCambiarHistorial });
    $('inicio').hidden = true; $('app').hidden = false; document.documentElement.classList.add('of-abierta');
    $('nombreDoc').value = nombreDoc;
    lectura = matchMedia('(max-width: 900px)').matches && !o.nuevo;   // en el celular se abre para leer; "Editar" activa la edición
    aplicarModo();
    armarPanel();
    cosas = await cosasDelArchivo().catch(() => new Set());
    await mostrarDiapositiva(0);
    actualizarGuardado(); comandos.refrescar();
    if (o.recuperado) autoguardado.programar(datosAutoguardado, 10);
  }
  async function presentacionNueva() {
    mostrarCargando(true);
    try { const p = await PresentacionEditable.nueva({ idioma: idioma() }); await abrirBytes(await p.generar(), idioma() === 'en' ? 'Presentation1' : 'Presentacion1', { nuevo: true }); } catch (e) { mostrarError(e); }
    mostrarCargando(false);
  }
  function mostrarCargando(si) { $('cargando').hidden = !si; if (si) $('errorInicio').hidden = true; }
  function mostrarError(e) { console.error(e); const el = $('errorInicio'); el.textContent = tx('error-abrir') + ': ' + (e && e.message ? e.message : e); el.hidden = false; $('inicio').hidden = false; $('app').hidden = true; document.documentElement.classList.remove('of-abierta'); }

  // ---- dibujo: una cola, porque el modelo se comparte ---------------------------------------------------------------
  let cola = Promise.resolve();
  const enCola = (fn) => { const p = cola.then(fn, fn); cola = p.catch(() => {}); return p; };
  async function dibujarEn(canvas, i, anchoPx) {
    const m = modelo;
    const diap = await m.diapositiva(i);
    canvas.width = Math.max(1, Math.round(anchoPx));
    canvas.height = Math.max(1, Math.round(anchoPx * m.alto / m.ancho));
    await PptxDibujo.dibujar(m, diap, canvas);
    return diap;
  }
  function tamanoEn(contenedor, relleno) {
    const cs = getComputedStyle(contenedor), pad = relleno === false ? 0 : parseFloat(cs.paddingLeft) + parseFloat(cs.paddingRight), padV = relleno === false ? 0 : parseFloat(cs.paddingTop) + parseFloat(cs.paddingBottom);
    const aw = Math.max(80, contenedor.clientWidth - pad), ah = Math.max(60, contenedor.clientHeight - padV), razon = modelo.ancho / modelo.alto;
    const w = Math.min(aw, ah * razon);
    return { w, h: w / razon };
  }
  let ordenDibujo = 0;
  async function mostrarDiapositiva(i, enfocar) {
    if (!modelo) return;
    if (editor && Math.max(0, Math.min(cantidad - 1, i)) !== editor.diap) await cerrarEditor(true);
    actual = Math.max(0, Math.min(cantidad - 1, i));
    if (actual !== mostrada) { sel = null; mostrada = actual; }
    const orden = ++ordenDibujo;
    const { w, h } = tamanoEn(escenario);
    marco.style.width = w + 'px'; marco.style.height = h + 'px';
    marco.classList.add('cargando');
    marcarActual(enfocar);
    let diap = null;
    try {
      diap = await enCola(() => (orden === ordenDibujo ? dibujarEn(lienzoPrincipal, actual, w * Math.min(2, window.devicePixelRatio || 1)) : null));
    } catch (e) { console.warn(e); }
    if (orden !== ordenDibujo) return;
    marco.classList.remove('cargando');
    diapActual = diap;
    marcarRepetidos(diap);
    lienzoPrincipal.setAttribute('aria-label', tx('diapositiva-de', actual + 1, cantidad));
    reponerSeleccion();
    lineaHonesta(diap);
    actualizarInspector();
    comandos.refrescar();
  }
  new ResizeObserver(() => { if (modelo && !$('app').hidden) mostrarDiapositiva(actual); }).observe(escenario);

  // ---- panel de diapositivas ---------------------------------------------------------------------------------------------
  const MINI_PX = 260;   // ancho del lienzo de una miniatura (se ve a 130 px: el doble para pantallas densas)
  let observador = null;
  function armarPanel() {
    if (observador) observador.disconnect();
    panel.textContent = '';
    observador = new IntersectionObserver((entradas) => {
      for (const e of entradas) {
        if (!e.isIntersecting) continue;
        observador.unobserve(e.target);
        const k = +e.target.dataset.indice;
        enCola(() => dibujarEn(e.target.querySelector('canvas'), k, MINI_PX)).catch((err) => console.warn(err));
      }
    }, { root: panel, rootMargin: '200px' });
    for (let k = 0; k < cantidad; k++) {
      const c = crearEl('canvas', { width: MINI_PX, height: Math.round(MINI_PX * modelo.alto / modelo.ancho), 'aria-hidden': 'true' });
      const b = crearEl('button', { type: 'button', class: 'pr-mini', role: 'option', dataset: { indice: String(k) }, 'aria-label': tx('diapositiva-n', k + 1), draggable: !lectura }, [crearEl('span', { class: 'pr-mini-n' }, [String(k + 1)]), c]);
      b.addEventListener('click', () => mostrarDiapositiva(k));
      panel.append(b);
      observador.observe(b);
    }
    marcarActual();
  }
  function redibujarMini(k) {
    const b = panel.children[k];
    if (b) enCola(() => dibujarEn(b.querySelector('canvas'), k, MINI_PX)).catch((err) => console.warn(err));
  }
  function marcarActual(enfocar) {
    for (const b of panel.children) {
      const si = +b.dataset.indice === actual;
      if (si) b.setAttribute('aria-current', 'true'); else b.removeAttribute('aria-current');
      b.setAttribute('aria-selected', String(si));
      if (si) { b.scrollIntoView({ block: 'nearest', inline: 'nearest' }); if (enfocar) b.focus(); }
    }
  }
  panel.addEventListener('keydown', (e) => {
    const paso = { ArrowDown: 1, ArrowRight: 1, ArrowUp: -1, ArrowLeft: -1 }[e.key];
    if (paso) { e.preventDefault(); mostrarDiapositiva(actual + paso, true); }
  });
  // reordenar arrastrando: ratón (arrastre HTML5) y dedo (reordenar-tactil.js). Los dos reacomodan los botones y al final se lee el orden nuevo.
  let arrastrando = null;
  panel.addEventListener('dragstart', (e) => { const b = e.target.closest('.pr-mini'); if (!b || lectura) { e.preventDefault(); return; } arrastrando = b; b.classList.add('arrastrando'); e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', b.dataset.indice); });
  panel.addEventListener('dragover', (e) => {
    if (!arrastrando) return;
    e.preventDefault();
    const objetivo = e.target.closest('.pr-mini');
    if (objetivo && objetivo !== arrastrando) insertarSegunPosicion(arrastrando, objetivo, e.clientX, e.clientY);
  });
  panel.addEventListener('drop', (e) => e.preventDefault());
  panel.addEventListener('dragend', () => { if (arrastrando) arrastrando.classList.remove('arrastrando'); arrastrando = null; sincronizarOrden(); });
  habilitarReordenTactil(panel, '.pr-mini', () => sincronizarOrden());
  // de un solo cambio de lugar (a → b) se deduce cuál se movió comparando el orden de los botones con el de antes
  function sincronizarOrden() {
    const orden = [...panel.children].map((b) => +b.dataset.indice);
    const p = orden.findIndex((v, k) => v !== k);
    if (p < 0) return;
    const de = orden[p] === p + 1 ? p : orden[p], a = orden[p] === p + 1 ? orden.indexOf(p) : p;
    if (lectura) { armarPanel(); return; }
    moverDiapositiva(de, a);
  }

  // ---- hacer cambios: todo pasa por el historial ---------------------------------------------------------------------------------------------
  const editable = () => !!pe && !!modelo && !lectura;
  // crearOp: función que devuelve la promesa de la operación. Los cambios van en fila: cada uno lee el XML después de que se aplicó el
  // anterior (dos cambios seguidos sobre la misma diapositiva, p. ej. Ctrl+B y un color, no se pisan)
  let filaCambios = Promise.resolve(), directo = false;
  function hacer(nombre, crearOp, extra) {
    if (!editable()) return Promise.resolve(null);
    const doc = pe;
    const p = filaCambios.then(async () => {
      if (pe !== doc || !editable()) return null;
      try {
        const op = await crearOp();
        if (!op || pe !== doc) return null;
        directo = true;
        try { hist.hacer(nombre, [op.t ? op : op.op], extra); } finally { directo = false; }
        return op;
      } catch (e) {
        console.warn(e);
        avisar(tx('aviso-no-se-pudo') + ': ' + (e && e.message ? e.message : e));
        armarPanel();
        return null;
      }
    });
    filaCambios = p.catch(() => {});
    return p;
  }

  // ---- selección: buscar el objeto, dibujar la caja y los tiradores --------------------------------------------------------------------------
  const escala = () => marco.clientWidth / modelo.ancho;   // píxeles por EMU
  // la clave de un objeto en la diapositiva; si otro tiene el mismo id (archivos de otros programas), se le suma cuál es (#1, #2…)
  const claveDe = (e) => e.origen.camino.join('/') + (e.origen.repetido ? '#' + e.origen.repetido : '');
  function marcarRepetidos(diap) {
    const vistos = new Map();
    for (const e of diap ? diap.elementos : []) {
      if (!e.origen) continue;
      const k = e.origen.capa + ':' + e.origen.camino.join('/'), n = vistos.get(k) || 0;
      if (n) e.origen.repetido = n;
      vistos.set(k, n + 1);
    }
  }
  const seleccionables = () => (diapActual ? diapActual.elementos.filter((e) => e.origen && e.origen.capa === 'diapositiva' && e.origen.camino.length === 1 && e.xfrm) : []);
  const elementoSel = () => (sel ? seleccionables().find((e) => claveDe(e) === sel.clave) || null : null);
  function tipoDe(e) { return { forma: 'obj-forma', imagen: 'obj-imagen', grupo: 'obj-grupo', tabla: 'obj-tabla', grafico: 'obj-grafico' }[e.tipo] || 'obj-otro'; }
  function reponerSeleccion() { if (sel && !elementoSel()) sel = null; pintarCaja(); pintarMarcadores(); }
  // Como PowerPoint: un marcador (título, subtítulo, contenido) sin texto muestra «Haz clic para agregar…» al editar; no sale en la presentación ni en el PDF
  function pintarMarcadores() {
    hints.textContent = '';
    if (lectura || !modelo || !diapActual) return;
    const s = escala(), pt = marco.clientWidth / (modelo.ancho / 12700);
    for (const e of seleccionables()) {
      if (!e.origen.marcador || !e.origen.texto || e.texto || (editor && editor.camino.join('/') === e.origen.camino.join('/'))) continue;
      const clave = { title: 'hint-titulo', ctrTitle: 'hint-titulo', subTitle: 'hint-subtitulo' }[e.origen.marcadorTipo] || 'hint-texto', r = e.xfrm;
      const d = crearEl('div', { class: 'pr-hint' }, [tx(clave)]);
      Object.assign(d.style, { left: r.x * s + 'px', top: r.y * s + 'px', width: r.w * s + 'px', height: r.h * s + 'px', transform: 'rotate(' + (r.rot || 0) + 'deg)', fontSize: Math.max(10, 18 * pt * 0.6) + 'px' });
      hints.append(d);
    }
  }
  function pintarCaja(forzada) {
    const e = elementoSel();
    if (!e || lectura) { caja.hidden = true; return; }
    const r = forzada || e.xfrm, s = escala();
    caja.hidden = false;
    caja.classList.toggle('bloqueada', !!sel.bloqueada);
    Object.assign(caja.style, { left: r.x * s + 'px', top: r.y * s + 'px', width: Math.max(2, r.w * s) + 'px', height: Math.max(2, r.h * s) + 'px', transform: 'rotate(' + (r.rot || 0) + 'deg)' });
  }
  async function seleccionar(e) {
    rango = null; fmtRango = null;
    if (!e) { sel = null; pintarCaja(); actualizarInspector(); comandos.refrescar(); return; }
    sel = { clave: claveDe(e), camino: e.origen.camino.slice(), bloqueada: !!(await pe.editable(actual, e.origen.camino)) };
    if (sel.bloqueada) avisar(tx('aviso-no-editable'));
    pintarCaja(); actualizarInspector(); comandos.refrescar();
  }
  // ¿qué objeto hay bajo el punto (en EMU de la diapositiva)? De arriba hacia abajo; se pasa al marco local de cada objeto para contar su giro
  function objetoEn(ex, ey) {
    const tol = 6 / escala(), lista = seleccionables();
    for (let k = lista.length - 1; k >= 0; k--) {
      const r = lista[k].xfrm, a = (r.rot || 0) * Math.PI / 180, dx = ex - (r.x + r.w / 2), dy = ey - (r.y + r.h / 2);
      const lx = Math.cos(a) * dx + Math.sin(a) * dy, ly = -Math.sin(a) * dx + Math.cos(a) * dy;
      if (Math.abs(lx) <= r.w / 2 + tol && Math.abs(ly) <= r.h / 2 + tol) return lista[k];
    }
    return null;
  }
  const aEmu = (ev) => { const b = marco.getBoundingClientRect(), s = escala(); return { x: (ev.clientX - b.left) / s, y: (ev.clientY - b.top) / s, px: ev.clientX - b.left, py: ev.clientY - b.top }; };

  // mover, cambiar el tamaño y girar con el puntero (ratón, dedo o lápiz). Mientras se arrastra solo se mueve la caja; al soltar se aplica el cambio.
  let gesto = null;
  sobre.addEventListener('pointerdown', async (ev) => {
    if (!editable() || ev.button > 0) return;
    sobre.focus({ preventScroll: true });
    rango = null; fmtRango = null;
    const asa = ev.target.closest('[data-asa]');
    const p = aEmu(ev);
    let e = elementoSel();
    if (asa && e && !sel.bloqueada) {
      gesto = { modo: asa.dataset.asa === 'giro' ? 'giro' : 'tam', asa: asa.dataset.asa, p0: p, r0: Object.assign({}, e.xfrm), r: Object.assign({}, e.xfrm), movido: false };
    } else {
      const hit = objetoEn(p.x, p.y);
      if (!hit) { await seleccionar(null); return; }
      if (!e || claveDe(hit) !== sel.clave) { await seleccionar(hit); e = hit; }
      if (sel.bloqueada) return;
      gesto = { modo: 'mover', p0: p, r0: Object.assign({}, e.xfrm), r: Object.assign({}, e.xfrm), movido: false };
    }
    sobre.setPointerCapture(ev.pointerId);
    ev.preventDefault();
  });
  sobre.addEventListener('pointermove', (ev) => {
    if (!gesto) return;
    const p = aEmu(ev), g = gesto, r0 = g.r0;
    if (!g.movido && Math.hypot(p.px - g.p0.px, p.py - g.p0.py) < 3) return;
    g.movido = true;
    let dx = p.x - g.p0.x, dy = p.y - g.p0.y;
    if (g.modo === 'mover') {
      if (ev.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0; }
      g.r = ajustarGuias(Object.assign({}, r0, { x: Math.round(r0.x + dx), y: Math.round(r0.y + dy) }));
    } else if (g.modo === 'giro') {
      const cx = r0.x + r0.w / 2, cy = r0.y + r0.h / 2;
      let grados = Math.atan2(p.y - cy, p.x - cx) * 180 / Math.PI + 90;
      if (ev.shiftKey) grados = Math.round(grados / 15) * 15;
      g.r = Object.assign({}, r0, { rot: ((Math.round(grados * 10) / 10) % 360 + 360) % 360 });
    } else {
      g.r = tamanoConAsa(r0, g.asa, dx, dy, ev.shiftKey);
    }
    pintarCaja(g.r);
  });
  const terminarGesto = async (cancelar) => {
    const g = gesto; gesto = null; quitarGuias();
    if (!g) return;
    if (cancelar || !g.movido) {
      pintarCaja();
      const e = elementoSel();
      if (!cancelar && g.modo === 'mover' && e && e.origen.marcador && e.origen.texto && !e.texto && haySel()) editarTexto();
      return;
    }
    await aplicarCaja(g.r, g.modo === 'giro' ? 'Girar' : g.modo === 'tam' ? 'Tamaño' : 'Mover');
  };
  sobre.addEventListener('pointerup', () => terminarGesto(false));
  sobre.addEventListener('pointercancel', () => terminarGesto(true));
  // el tamaño nuevo al arrastrar un tirador: el lado o la esquina opuesta se queda quieta, también con el objeto girado
  function tamanoConAsa(r0, asa, dx, dy, proporcion) {
    const [, hx, hy] = ASAS.find((a) => a[0] === asa), ang = (r0.rot || 0) * Math.PI / 180, c = Math.cos(ang), s = Math.sin(ang);
    const ldx = c * dx + s * dy, ldy = -s * dx + c * dy;
    let w = hx ? Math.max(MINIMO, r0.w + hx * ldx) : r0.w, h = hy ? Math.max(MINIMO, r0.h + hy * ldy) : r0.h;
    if (proporcion && hx && hy && r0.w > 0 && r0.h > 0) { const k = Math.max(w / r0.w, h / r0.h); w = r0.w * k; h = r0.h * k; }
    const lcx = hx * (w - r0.w) / 2, lcy = hy * (h - r0.h) / 2;
    const cx = r0.x + r0.w / 2 + c * lcx - s * lcy, cy = r0.y + r0.h / 2 + s * lcx + c * lcy;
    return Object.assign({}, r0, { x: Math.round(cx - w / 2), y: Math.round(cy - h / 2), w: Math.round(w), h: Math.round(h) });
  }
  async function aplicarCaja(r, nombre, fusionar) {
    const e = elementoSel();
    if (!e || sel.bloqueada) return;
    const caja_ = { x: r.x, y: r.y, cx: r.w, cy: r.h };
    if (r.rot !== (e.xfrm.rot || 0)) caja_.rot = Math.round((r.rot || 0) * 60000);
    const i = actual, camino = sel.camino;
    const op = await hacer(nombre, () => pe.opMover(i, camino, caja_), fusionar ? { fusionar } : undefined);
    if (!op) pintarCaja();
  }
  sobre.addEventListener('dblclick', async (ev) => {
    if (!editable()) return;
    const p = aEmu(ev), h = objetoEn(p.x, p.y);
    if (!h || !h.origen.texto) return;
    if (!sel || sel.clave !== claveDe(h)) await seleccionar(h);
    editarTexto();
  });
  // teclado sobre el escenario: flechas mueven (Mayús: más), Supr elimina, Tab recorre los objetos, Esc quita la selección
  sobre.addEventListener('keydown', async (ev) => {
    if (!editable()) return;
    const lista = seleccionables();
    if (ev.key === 'Tab' && lista.length) {
      // pasado el último objeto (o antes del primero), Tab sale del escenario: no es una trampa de foco
      const k = sel ? lista.findIndex((x) => claveDe(x) === sel.clave) : -1;
      const sig = ev.shiftKey ? (k < 0 ? -1 : k - 1) : k + 1;
      if (sig < 0 || sig >= lista.length) { if (sel) await seleccionar(null); return; }
      ev.preventDefault();
      await seleccionar(lista[sig]);
      return;
    }
    if (ev.key === 'Escape') { await seleccionar(null); return; }
    if ((ev.key === 'Enter' || ev.key === 'F2') && sel && !sel.bloqueada) { ev.preventDefault(); editarTexto(); return; }
    const paso = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[ev.key];
    if (paso && elementoSel() && !sel.bloqueada) {
      ev.preventDefault();
      // en fila y con el modelo ya recargado: cada pulsación parte de donde dejó la anterior (si no, al repetir la tecla se perdían pasos)
      if (flechasPendientes >= 5) return;   // con la tecla apretada no se acumulan pasos para después de soltarla
      flechasPendientes++;
      const d = EMU_POR_PX * (ev.shiftKey ? 10 : 1);
      filaFlechas = filaFlechas.then(async () => {
        await cadena;
        const e = elementoSel();
        if (e && !sel.bloqueada) await aplicarCaja(Object.assign({}, e.xfrm, { x: e.xfrm.x + paso[0] * d, y: e.xfrm.y + paso[1] * d }), 'Mover', 'mov:' + sel.clave);
      }).catch((err) => console.warn(err)).finally(() => { flechasPendientes--; });
    }
  });
  let filaFlechas = Promise.resolve(), flechasPendientes = 0;

  // ---- texto: un cuadro de texto nativo sobre el objeto (el teclado del celular y el IME funcionan solos) --------------------------------------------------------
  let editor = null;   // { area, camino, diap, inicial: [párrafos] }
  const trozoDe = (e) => { const ps = e && e.texto && e.texto.parrafos; const t = ps && ps.flatMap((p) => p.trozos || []).find((x) => x.texto); return t || (ps && ps[0] && ps[0].trozos && ps[0].trozos[0]) || null; };
  async function editarTexto() {
    const e = elementoSel();
    if (editor || !e || !e.origen.texto || !haySel()) return;
    const inicial = await pe.textoEditable(actual, sel.camino);
    const r = e.xfrm, s = escala(), pt = marco.clientWidth / (modelo.ancho / 12700), t = trozoDe(e) || {};
    const area = crearEl('textarea', { class: 'pr-editor-texto', 'aria-label': tx('cmd-editar-texto'), 'data-i18n-attr': 'aria-label:cmd-editar-texto', spellcheck: 'true' });
    area.value = inicial.join('\n');
    Object.assign(area.style, { left: r.x * s + 'px', top: r.y * s + 'px', width: Math.max(60, r.w * s) + 'px', height: Math.max(30, r.h * s) + 'px', transform: 'rotate(' + (r.rot || 0) + 'deg)', fontSize: Math.max(8, (t.tamano || 18) * pt) + 'px', fontFamily: (t.fuente ? '"' + t.fuente + '", ' : '') + 'Carlito, Calibri, "Liberation Sans", Arial, sans-serif', fontWeight: t.negrita ? '700' : '400', fontStyle: t.cursiva ? 'italic' : 'normal' });
    editor = { area, camino: sel.camino.slice(), diap: actual, inicial };
    pintarMarcadores();
    marco.append(area);
    area.addEventListener('keydown', (ev) => {
      const atajo = (ev.ctrlKey || ev.metaKey) && !ev.altKey && { b: 'negrita', i: 'cursiva', u: 'subrayado' }[ev.key.toLowerCase()];
      if (ev.key === 'Escape') { ev.preventDefault(); cerrarEditor(true); } else if (ev.key === 'Enter' && (ev.ctrlKey || ev.metaKey)) { ev.preventDefault(); cerrarEditor(true); }
      else if (ev.key === 'Enter' && ev.shiftKey && editor && editor.entrada) { ev.preventDefault(); area.setRangeText('\u2028', area.selectionStart, area.selectionEnd, 'end'); area.dispatchEvent(new Event('input', { bubbles: true })); }   // como PowerPoint: salto de línea dentro del párrafo (<a:br>); el cuadro nativo lo mostraría como un espacio, así que ahí sigue dando otra línea
      else if (atajo) { ev.preventDefault(); comandos.ejecutar(atajo); }
      else if (ev.key === 'Tab' && !ev.ctrlKey && !ev.altKey) { ev.preventDefault(); comandos.ejecutar(ev.shiftKey ? 'nivelMenos' : 'nivelMas'); }
      ev.stopPropagation();
    });
    area.addEventListener('blur', () => cerrarEditor(true, true));
    for (const nombre of ['select', 'keyup', 'mouseup', 'input']) area.addEventListener(nombre, leerFormatoRango);
    rango = null; fmtRango = null;
    area.focus();
    area.setSelectionRange(area.value.length, area.value.length);
    // con el motor de texto el cuadro nativo no se ve: el texto se dibuja con su formato y su corte reales (texto vertical, girado o espejado sigue con el cuadro nativo);
    // una forma sin texto usa el formato que tendrá lo que se escriba
    const conModelo = e.texto || !e.textoVacio ? e : { ...e, texto: e.textoVacio };
    if (EntradaTexto.admite(conModelo)) { try { editor.entrada = EntradaTexto.crear({ marco, area, elemento: conModelo, escala: s }); } catch (err) { console.warn(err); } }
    leerFormatoRango();
  }
  // La selección dentro del cuadro de texto, como en PowerPoint: con texto seleccionado, el formato va a lo seleccionado; con el cursor dentro de una
  // palabra, a la palabra; los números son posiciones en el texto de la forma (párrafos unidos con \n y saltos de línea como U+2028).
  let rango = null, fmtRango = null;   // rango: la última selección del cuadro ya cerrado, para el formato que llega después (selector de color)
  function seleccionDe(area) {
    let ini = area.selectionStart, fin = area.selectionEnd;
    if (ini === fin) {
      const v = area.value, esLetra = (c) => /[\p{L}\p{N}_]/u.test(c);
      while (ini > 0 && esLetra(v[ini - 1])) ini--;
      while (fin < v.length && esLetra(v[fin])) fin++;
    }
    return { ini, fin };
  }
  // El formato efectivo de lo seleccionado sale del modelo (ya trae lo heredado del diseño y del patrón): los tramos de texto con su posición.
  function leerFormatoRango() {
    const ed = editor, e = elementoSel();
    if (!ed || !e || (!e.texto && !ed.entrada)) return;
    const { ini, fin } = seleccionDe(ed.area);
    const tramos = [];
    let pos = 0;
    // con el motor de texto, el modelo vivo (lo que se ve): el del archivo todavía no tiene lo escrito; un salto suave cuenta 1 y no da formato
    const parrafos = ed.entrada ? ed.entrada.modelo() : e.texto.parrafos;
    for (const p of parrafos) { for (const z of p.trozos) { const n = z.salto ? 1 : (z.texto || '').length; if (n && !z.salto && pos < fin && pos + n > ini) tramos.push(z); pos += n; } pos += 1; }
    if (fin <= ini || !tramos.length) { fmtRango = null; comandos.refrescar(); return; }
    fmtRango = { negrita: tramos.every((z) => z.negrita), cursiva: tramos.every((z) => z.cursiva), subrayado: tramos.every((z) => z.subrayado), tamano: tramos[0].tamano };
    comandos.refrescar();
  }
  async function cerrarEditor(guardar, porPerderFoco) {
    const ed = editor;
    if (!ed) return;
    const s = seleccionDe(ed.area);
    rango = guardar && porPerderFoco && s.fin > s.ini ? { diap: ed.diap, camino: ed.camino, ini: s.ini, fin: s.fin, hasta: Date.now() + 15000 } : null;   // solo si se cerró al perder el foco (p. ej. el selector de color) y por poco tiempo
    editor = null;
    pintarMarcadores();
    const nuevo = ed.area.value.split('\n');
    if (ed.entrada) ed.entrada.destruir();
    ed.area.remove();
    if (guardar && nuevo.join('\n') !== ed.inicial.join('\n')) await hacer('Texto', () => pe.opTexto(ed.diap, ed.camino, nuevo), { fusionar: 'texto:' + ed.camino.join('/') });
    sobre.focus({ preventScroll: true });
  }
  // formato del texto de todo el objeto seleccionado
  const formatoSel = () => trozoDe(elementoSel()) || {};
  const conTexto = () => haySel() && !!elementoSel() && !!elementoSel().origen.texto;
  const enSel = (nombre, fn) => { const i = actual, c = sel.camino; return hacer(nombre, () => fn(i, c)); };   // un cambio del objeto seleccionado ahora
  const formatoActual = () => (editor && fmtRango) || formatoSel();
  // El formato va a lo seleccionado en el cuadro de texto (o a la palabra bajo el cursor); sin cuadro de texto, a todo el texto del objeto.
  async function formatear(f) {
    const r = editor ? Object.assign({ diap: editor.diap, camino: editor.camino }, seleccionDe(editor.area)) : (rango && rango.hasta > Date.now() ? rango : null);
    if (r && sel && r.diap === actual && r.camino.join('/') === sel.camino.join('/')) {
      if (r.fin <= r.ini) return;   // el cursor no está en una palabra: no hay a qué darle formato
      if (editor) await cerrarEditor(true);   // primero se guarda el texto escrito: las posiciones siguen valiendo
      rango = null;
      await hacer('Formato', () => pe.opFormatoRango(r.diap, r.camino, r.ini, r.fin, f));
      await cadena;
      await editarTexto();   // como PowerPoint, se sigue editando con lo mismo seleccionado
      if (editor) { editor.area.setSelectionRange(r.ini, r.fin); leerFormatoRango(); }
      return;
    }
    return enSel('Formato', (i, c) => pe.opFormatoTexto(i, c, f));
  }

  // Viñetas, numeración y nivel: por párrafo. Con el cuadro de texto abierto, los párrafos que toca la selección; si no, todos los del objeto.
  const parrafoActual = () => {
    const e = elementoSel();
    if (!e || (!e.texto && !(editor && editor.entrada))) return null;
    const k = editor ? editor.area.value.slice(0, editor.area.selectionStart).split('\n').length - 1 : 0;
    const ps = editor && editor.entrada ? editor.entrada.modelo() : e.texto.parrafos;   // con el motor, el modelo vivo: tiene los párrafos recién escritos
    return ps[Math.min(k, ps.length - 1)] || null;
  };
  const tipoVineta = () => { const p = parrafoActual(); return p && p.vineta ? p.vineta.tipo : null; };
  async function formatearParrafos(f) {
    if (!haySel()) return;
    const a = editor && editor.area, v = a ? a.value : '';
    const desde = a ? v.slice(0, a.selectionStart).split('\n').length - 1 : 0, hasta = a ? v.slice(0, a.selectionEnd).split('\n').length - 1 : 1e6;
    const camino = sel.camino.slice(), diap = actual, ini = a && a.selectionStart, fin = a && a.selectionEnd, estaba = !!editor;
    if (estaba) await cerrarEditor(true);
    await hacer('Párrafo', () => pe.opFormatoParrafo(diap, camino, desde, hasta, f));
    if (estaba) { await cadena; await editarTexto(); if (editor) { editor.area.setSelectionRange(ini, fin); leerFormatoRango(); } }
  }

  // ---- insertar: cuadro de texto e imagen ---------------------------------------------------------------------------------------------------------------------
  async function insertarCuadroTexto() {
    if (!editable()) return;
    const cx = Math.round(modelo.ancho * 0.4), cy = Math.round(modelo.alto * 0.12);
    const i = actual, caja_ = { x: Math.round((modelo.ancho - cx) / 2), y: Math.round((modelo.alto - cy) / 2), cx, cy };
    const r = await hacer('Cuadro de texto', () => pe.opCuadroTexto(i, caja_, tx('texto-nuevo')));
    if (!r) return;
    await cadena;
    await seleccionarPorId(r.id);
    editarTexto();
  }
  // Insertar → Formas: la galería de formas predefinidas de la norma; la forma queda en el centro de la diapositiva, seleccionada, con el estilo del tema
  async function insertarForma() {
    if (!editable()) return;
    const f = await GaleriaFormas.abrir({ idioma: idioma() });
    if (!f) return;
    const cx = Math.round(f.ancho * EMU_POR_CM), cy = Math.round(f.alto * EMU_POR_CM), i = actual;
    const r = await hacer('Forma', () => pe.opForma(i, { x: Math.round((modelo.ancho - cx) / 2), y: Math.round((modelo.alto - cy) / 2), cx, cy }, f.prst, null, { nombre: f.nombre, flecha: f.flecha }));
    if (r) { await cadena; await seleccionarPorId(r.id); }
  }
  async function insertarImagen(archivo) {
    if (!editable() || !archivo) return;
    const tipo = archivo.type;
    if (!/^image\/(png|jpeg|gif)$/.test(tipo)) { avisar(tx('aviso-imagen')); return; }
    try {
      const bytes = new Uint8Array(await archivo.arrayBuffer());
      const bmp = await createImageBitmap(archivo);
      const maxW = modelo.ancho * 0.6, maxH = modelo.alto * 0.6, k = Math.min(9525, maxW / bmp.width, maxH / bmp.height);
      const cx = Math.round(bmp.width * k), cy = Math.round(bmp.height * k);
      const i = actual, caja_ = { x: Math.round((modelo.ancho - cx) / 2), y: Math.round((modelo.alto - cy) / 2), cx, cy };
      const r = await hacer('Imagen', () => pe.opImagen(i, bytes, tipo, caja_));
      if (r) { await cadena; await seleccionarPorId(r.id); }
    } catch (e) { console.warn(e); avisar(tx('aviso-no-se-pudo') + ': ' + (e && e.message ? e.message : e)); }
  }
  async function seleccionarPorId(id) { const e = seleccionables().find((x) => x.origen.camino[0] === String(id)); if (e) await seleccionar(e); }

  // ---- guías: al mover, el objeto se pega a los bordes y centros de la diapositiva y de los demás objetos ------------------------------------------------------------------------
  const guiaV = crearEl('div', { class: 'pr-guia pr-guia-v', hidden: true }), guiaH = crearEl('div', { class: 'pr-guia pr-guia-h', hidden: true });
  sobre.append(guiaV, guiaH);
  function quitarGuias() { guiaV.hidden = true; guiaH.hidden = true; }
  function ajustarGuias(r) {
    const tol = 6 / escala(), otros = seleccionables().filter((x) => !sel || claveDe(x) !== sel.clave);
    const xs = [0, modelo.ancho / 2, modelo.ancho], ys = [0, modelo.alto / 2, modelo.alto];
    for (const o of otros) { const q = o.xfrm; xs.push(q.x, q.x + q.w / 2, q.x + q.w); ys.push(q.y, q.y + q.h / 2, q.y + q.h); }
    const mejor = (propios, candidatas) => { let m = null; for (const p of propios) for (const c of candidatas) { const d = c - p; if (Math.abs(d) <= tol && (!m || Math.abs(d) < Math.abs(m.d))) m = { d, en: c }; } return m; };
    const mx = mejor([r.x, r.x + r.w / 2, r.x + r.w], xs), my = mejor([r.y, r.y + r.h / 2, r.y + r.h], ys), s = escala();
    const n = Object.assign({}, r);
    if (mx) n.x = Math.round(r.x + mx.d); if (my) n.y = Math.round(r.y + my.d);
    guiaV.hidden = !mx; guiaH.hidden = !my;
    if (mx) guiaV.style.left = mx.en * s + 'px';
    if (my) guiaH.style.top = my.en * s + 'px';
    return n;
  }

  // ---- exportar a PDF: vectores (pdf-lienzo.js); si una diapositiva no sale así, va como imagen ------------------------------------------------------------------
  async function exportarPdf() {
    if (!hayPres()) return;
    avisar(tx('aviso-pdf'));
    try {
      const pdf = await PDFLib.PDFDocument.create(), m = modelo;
      const ptW = m.ancho / 12700, ptH = m.alto / 12700, pxW = 1920, pxH = Math.round(1920 * m.alto / m.ancho);
      let hechas = 0;
      for (let k = 0; k < cantidad; k++) {
        await enCola(async () => {
          const diap = await m.diapositiva(k);
          if (diap.oculta) return;   // como PowerPoint, las ocultas no van al PDF
          try {
            const lienzo = new LienzoPdf.Lienzo({ ancho: pxW, alto: pxH, anchoPt: ptW, altoPt: ptH });
            await PptxDibujo.dibujar(m, diap, lienzo);
            if (lienzo.fallo) throw lienzo.fallo;
            await lienzo.aPagina(pdf);
          } catch (err) {
            console.warn('Diapositiva ' + (k + 1) + ' como imagen:', err);
            const c = document.createElement('canvas'); c.width = pxW; c.height = pxH;
            await PptxDibujo.dibujar(m, diap, c);
            const png = await new Promise((res) => c.toBlob(res, 'image/png'));
            const pagina = pdf.addPage([ptW, ptH]);
            pagina.drawImage(await pdf.embedPng(await png.arrayBuffer()), { x: 0, y: 0, width: ptW, height: ptH });
          }
          hechas++;
        });
      }
      descargarBlob(new Blob([await pdf.save()], { type: 'application/pdf' }), nombreDoc + '.pdf');
      avisar(tx('aviso-pdf-listo', hechas));
    } catch (e) { console.warn(e); avisar(tx('aviso-no-se-pudo') + ': ' + (e && e.message ? e.message : e)); }
  }

  // ---- inspector (con controles), línea honesta ----------------------------------------------------------------------------------------------------------
  const cm = (emu) => (emu / EMU_POR_CM).toLocaleString(idioma() === 'en' ? 'en-US' : 'es-CR', { maximumFractionDigits: 2 });
  const aHex = (c) => (c ? '#' + [c.r, c.g, c.b].map((v) => Math.max(0, Math.min(255, v | 0)).toString(16).padStart(2, '0')).join('') : null);
  const colorDe = (rell) => (rell && rell.tipo === 'solido' ? aHex(rell.color) : null);
  function actualizarInspector() {
    if (!modelo) return;
    const diap = diapActual;
    const secciones = [
      { titulo: tx('i-diapositiva'), filas: [['i-numero', actual + 1 + ' / ' + cantidad], ['i-formas', diap ? String(seleccionables().length) : '…'], ['i-oculta', diap && diap.oculta ? tx('si') : tx('no')]] },
      { titulo: tx('i-archivo'), filas: [['i-tamano', cm(modelo.ancho) + ' × ' + cm(modelo.alto) + ' cm'], ['i-diapositivas', String(cantidad)]] },
    ];
    const e = elementoSel();
    if (e) secciones.unshift({ titulo: tx('i-objeto'), filas: [['i-nombre', e.nombre || '—'], ['i-tipo', tx(tipoDe(e))]] });
    insp.mostrar(secciones);
    if (e && !sel.bloqueada && !lectura) controlesObjeto(e);
  }
  function controlesObjeto(e) {
    const r = e.xfrm;
    const campo = (clave, valor, alCambiar, paso) => {
      const entrada = crearEl('input', { type: 'number', step: paso || '0.01', value: String(valor), id: 'ctl-' + clave, 'aria-label': tx(clave) });
      entrada.addEventListener('change', () => { const v = parseFloat(entrada.value); if (Number.isFinite(v)) alCambiar(v); else pintarCaja(); });
      return crearEl('label', { class: 'pr-campo' }, [crearEl('span', {}, [tx(clave)]), entrada]);
    };
    const redondear = (v) => Math.round(v * 100) / 100;
    const con = (cambios) => aplicarCaja(Object.assign({}, r, cambios), 'Propiedades');
    const rejilla = crearEl('div', { class: 'pr-campos' }, [
      campo('i-x', redondear(r.x / EMU_POR_CM), (v) => con({ x: Math.round(v * EMU_POR_CM) })),
      campo('i-y', redondear(r.y / EMU_POR_CM), (v) => con({ y: Math.round(v * EMU_POR_CM) })),
      campo('i-ancho', redondear(r.w / EMU_POR_CM), (v) => con({ w: Math.max(MINIMO, Math.round(v * EMU_POR_CM)) })),
      campo('i-alto', redondear(r.h / EMU_POR_CM), (v) => con({ h: Math.max(MINIMO, Math.round(v * EMU_POR_CM)) })),
      campo('i-giro', redondear(r.rot || 0), (v) => con({ rot: ((v % 360) + 360) % 360 }), '1'),
    ]);
    const nodos = [rejilla];
    if (e.origen.texto) {
      const bt = crearEl('button', { type: 'button', class: 'boton', id: 'ctl-editar-texto' }, [tx('cmd-editar-texto')]);
      bt.addEventListener('click', () => editarTexto());
      nodos.push(bt);
      const color = (clave, actualHex, aplicar) => {
        const entrada = crearEl('input', { type: 'color', value: actualHex || '#ffffff', id: 'ctl-' + clave, 'aria-label': tx(clave) });
        entrada.addEventListener('change', () => aplicar(entrada.value.slice(1).toUpperCase()));
        const quitar = crearEl('button', { type: 'button', class: 'boton boton-secundario' }, [tx('i-quitar')]);
        quitar.addEventListener('click', () => aplicar(null));
        return crearEl('div', { class: 'pr-campo pr-color' }, [crearEl('span', {}, [tx(clave)]), entrada, quitar]);
      };
      nodos.push(
        color('i-relleno', colorDe(e.relleno), (hex) => enSel('Relleno', (i, c) => pe.opRelleno(i, c, hex))),
        color('i-linea', colorDe(e.linea && e.linea.relleno), (hex) => enSel('Línea', (i, c) => pe.opLinea(i, c, hex))),
      );
    }
    interfaz.inspector.append(crearEl('h2', { class: 'of-inspector-sep' }, [tx('i-objeto')]), ...nodos);
  }
  // Lo que el archivo trae y aquí no se edita (se mira una vez, al abrir): se conserva tal cual al guardar.
  let cosas = new Set();
  async function cosasDelArchivo() {
    const c = new Set();
    const nombres = Object.keys(pe.zip.files);
    const hay = (re) => nombres.some((n) => re.test(n));
    if (hay(/^ppt\/notesSlides\//)) c.add('x-notas');
    if (hay(/^ppt\/(comments|authors)|^ppt\/comments\//i)) c.add('x-comentarios');
    if (hay(/^ppt\/charts\//)) c.add('x-graficos');
    if (hay(/^ppt\/diagrams\//)) c.add('x-smartart');
    if (hay(/^ppt\/embeddings\//)) c.add('x-incrustados');
    if (hay(/\.(mp3|mp4|wav|m4a|mov|wmv|avi|wma)$/i)) c.add('x-multimedia');
    if (hay(/vbaProject\.bin$/i)) c.add('x-macros');
    for (const r of modelo.rutas) {
      const xml = (await pe.paquete.leer(r)) || '';
      if (/<(?:[\w.-]+:)?timing\b/.test(xml)) c.add('x-animaciones');
      if (/<(?:[\w.-]+:)?transition\b/.test(xml)) c.add('x-transiciones');
      if (/<(?:[\w.-]+:)?(?:tbl|grpSp)\b/.test(xml)) c.add('x-tablas');
    }
    return c;
  }
  function lineaHonesta(diap) {
    const partes = [tx('honesta-edita')];
    if (cosas.size) partes.push(tx('honesta-conserva', [...cosas].map((k) => tx(k)).join(', ')));
    const motivos = diap && diap.motivos ? diap.motivos.map((m) => (typeof m === 'string' ? m : m && (m.motivo || m.mensaje || m.tipo)) || '').filter(Boolean) : [];
    if (motivos.length) partes.push(tx('honesta-motivos', actual + 1, [...new Set(motivos)].slice(0, 4).join(', ')));
    $('lineaHonesta').textContent = partes.join(' ');
    $('lineaHonesta').title = $('lineaHonesta').textContent;
  }

  // ---- historial, guardado y descarga -------------------------------------------------------------------------------------------
  // Cada cambio dice qué partes tocó: así se sabe qué miniaturas rehacer y a qué diapositiva ir (también al deshacer o rehacer).
  let tocadas = new Set(), rutasAntes = '', pedirIr = null, cadena = Promise.resolve();
  function aplicarOp(op) { for (const c of op.cambios) tocadas.add(c.ruta); return pe.aplicar(op); }
  async function recargarModelo() { modelo = await PptxModelo.abrir(null, { zip: pe.zip }); cantidad = modelo.cantidad; }
  function alCambiarHistorial() {
    // un cambio hecho aquí (no deshacer/rehacer) no mueve de diapositiva: el usuario pudo haber pasado a otra mientras se recarga el modelo
    const t = tocadas, destino = pedirIr, quedarse = directo; tocadas = new Set(); pedirIr = null;
    actualizarGuardado(); comandos.refrescar();
    autoguardado.programar(datosAutoguardado, 2500, (fecha) => { if (fecha) actualizarGuardado(fecha); });
    cadena = cadena.then(async () => {
      await recargarModelo();
      const ahora = modelo.rutas.join('|');
      const estructura = ahora !== rutasAntes || t.has('ppt/presentation.xml');
      rutasAntes = ahora;
      const tocadasIdx = [...t].map((r) => modelo.rutas.indexOf(r)).filter((k) => k >= 0);
      if (estructura) armarPanel(); else tocadasIdx.forEach(redibujarMini);
      if (destino !== null) actual = destino; else if (!quedarse && tocadasIdx.length && !tocadasIdx.includes(actual)) actual = tocadasIdx[0];
      await mostrarDiapositiva(actual);
    }).catch((e) => console.warn(e));
  }
  async function datosAutoguardado() { if (!pe || !hist.sucio()) return null; return { id: docId, nombre: nombreDoc, bytes: await pe.generar() }; }
  let ultimoAutoguardado = null;
  function actualizarGuardado(fecha) {
    if (fecha) ultimoAutoguardado = fecha;
    const el = $('estadoGuardado');
    if (!hist || !hist.sucio()) { el.textContent = tx('sin-cambios'); return; }
    el.textContent = ultimoAutoguardado ? tx('guardado-local', new Date(ultimoAutoguardado).toLocaleTimeString(idioma() === 'en' ? 'en-US' : 'es-CR', { hour: '2-digit', minute: '2-digit' })) : tx('cambios-sin-descargar');
  }
  function avisar(texto) { Oficina.aviso($('avisoEstado'), texto, 5000); }
  async function descargar() {
    const bytes = await pe.generar();
    const conMacros = !!pe.zip.file('ppt/vbaProject.bin'); // como PowerPoint: una presentación con macros se guarda como .pptm
    descargarBlob(new Blob([bytes], { type: conMacros ? 'application/vnd.ms-powerpoint.presentation.macroEnabled.12' : 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }), nombreDoc + (conMacros ? '.pptm' : '.pptx'));
    hist.marcarGuardado();
    actualizarGuardado();
    avisar(tx('descargado'));
    autoguardado.borrar(docId);
  }

  // ---- diapositivas: agregar, duplicar, borrar, mover ------------------------------------------------------------------------------------------------------------
  // la diapositiva a la que ir se pide justo antes de aplicar (dentro de la fila de cambios), así otro cambio no la consume
  const conDestino = (destino, crear) => async () => { const op = await crear(); if (op) pedirIr = destino; return op; };
  async function nuevaDiapositiva() { if (!editable()) return; const i = actual; await hacer('Diapositiva nueva', conDestino(i + 1, () => pe.opNuevaDiapositiva(i))); }
  async function duplicarDiapositiva() { if (!editable()) return; const i = actual; await hacer('Duplicar', conDestino(i + 1, () => pe.opDuplicarDiapositiva(i))); }
  async function borrarDiapositiva() {
    if (!editable()) return;
    if (cantidad <= 1) { avisar(tx('aviso-ultima')); return; }
    const i = actual;
    await hacer('Borrar diapositiva', conDestino(Math.min(i, cantidad - 2), () => pe.opBorrarDiapositiva(i)));
  }
  async function moverDiapositiva(de, a) {
    if (!editable() || a < 0 || a >= cantidad || de === a) { armarPanel(); return; }
    await hacer('Mover diapositiva', conDestino(a, () => pe.opMoverDiapositiva(de, a)));
  }

  // ---- modo presentación ------------------------------------------------------------------------------------------------------------------------------------------------------------------
  let pres = null;   // { el, canvas, contador, k, orden }
  async function oculta(k) { try { return !!(await modelo.diapositiva(k)).oculta; } catch (e) { return false; } }
  async function dibujarPresentacion() {
    const p = pres, orden = ++p.orden;
    const { w } = tamanoEn(p.el, false);
    p.canvas.style.width = w + 'px';
    p.canvas.style.height = (w * modelo.alto / modelo.ancho) + 'px';
    try { await enCola(() => (orden === p.orden ? dibujarEn(p.canvas, p.k, w * Math.min(2, window.devicePixelRatio || 1)) : null)); } catch (e) { console.warn(e); }
    if (orden !== p.orden) return;
    p.canvas.setAttribute('aria-label', tx('diapositiva-de', p.k + 1, cantidad));
    p.contador.textContent = tx('pres-n', p.k + 1, cantidad);
  }
  async function irPresentacion(paso) {
    const p = pres;
    if (!p) return;
    let k = p.k + paso;
    while (k >= 0 && k < cantidad && await oculta(k)) k += paso;
    if (k < 0) return;
    if (k >= cantidad) { p.contador.textContent = tx('pres-fin'); return; }
    p.k = k;
    await dibujarPresentacion();
  }
  function salirPresentacion() {
    if (!pres) return;
    const p = pres; pres = null;
    window.removeEventListener('keydown', teclasPresentacion, true);
    window.removeEventListener('resize', redimensionarPresentacion);
    document.removeEventListener('fullscreenchange', alCambiarPantalla);
    p.el.remove();
    if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
    mostrarDiapositiva(p.k);
  }
  const alCambiarPantalla = () => { if (pres && !document.fullscreenElement) salirPresentacion(); };
  const redimensionarPresentacion = () => { if (pres) dibujarPresentacion(); };
  function teclasPresentacion(ev) {
    const adelante = ['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter'], atras = ['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace'];
    if (ev.key === 'Escape') salirPresentacion();
    else if (adelante.includes(ev.key)) irPresentacion(1);
    else if (atras.includes(ev.key)) irPresentacion(-1);
    else if (ev.key === 'Home') { pres.k = -1; irPresentacion(1); }
    else if (ev.key === 'End') { pres.k = cantidad; irPresentacion(-1); }
    else return;
    ev.preventDefault(); ev.stopPropagation();
  }
  async function presentar(desde) {
    if (!modelo || pres) return;
    const canvas = crearEl('canvas', { class: 'pr-pres-lienzo', role: 'img' });
    const contador = crearEl('span', { class: 'pr-pres-contador', 'aria-live': 'polite' });
    const cerrar = crearEl('button', { type: 'button', class: 'pr-pres-cerrar', 'aria-label': tx('pres-salir') }, ['✕']);
    const el = crearEl('div', { class: 'pr-presentacion', role: 'dialog', 'aria-modal': 'true', 'aria-label': tx('cmd-presentar') }, [canvas, contador, cerrar]);
    pres = { el, canvas, contador, k: Math.max(0, Math.min(cantidad - 1, desde)), orden: 0 };
    cerrar.addEventListener('click', salirPresentacion);
    // tocar o hacer clic: el tercio de la izquierda retrocede, el resto avanza
    el.addEventListener('click', (ev) => { if (ev.target === cerrar) return; irPresentacion(ev.clientX < el.clientWidth / 3 ? -1 : 1); });
    document.body.append(el);
    window.addEventListener('keydown', teclasPresentacion, true);
    window.addEventListener('resize', redimensionarPresentacion);
    document.addEventListener('fullscreenchange', alCambiarPantalla);
    try { if (el.requestFullscreen) await el.requestFullscreen(); } catch (e) { /* sin pantalla completa se queda como capa que ocupa todo */ }
    cerrar.focus();
    await dibujarPresentacion();
  }

  // ---- modo lectura (celular) -----------------------------------------------------------------------------------------------
  function aplicarModo() {
    $('app').classList.toggle('of-lectura', lectura);
    $('botonModo').textContent = lectura ? '✎ ' + tx('cmd-modo-editar') : '👁 ' + tx('cmd-modo-leer');
    for (const b of panel.children) b.draggable = !lectura;
    if (lectura && sel) sel = null;
    pintarCaja();
    pintarMarcadores();
    actualizarInspector();
  }

  // ---- comandos ---------------------------------------------------------------------------------------------------------
  const hayPres = () => !!pe && !!modelo;
  const reg = (c) => comandos.registrar(Object.assign({ habilitado: hayPres }, c));
  const haySel = () => editable() && !!sel && !sel.bloqueada;
  reg({ id: 'deshacer', texto: 'cmd-deshacer', atajo: 'mod+z', ejecutar: () => hist.deshacer(), habilitado: () => editable() && hist.puedeDeshacer() });
  reg({ id: 'rehacer', texto: 'cmd-rehacer', atajo: ['mod+y', 'mod+shift+z'], ejecutar: () => hist.rehacer(), habilitado: () => editable() && hist.puedeRehacer() });
  reg({ id: 'descargar', texto: 'cmd-descargar', atajo: 'mod+s', ejecutar: descargar });
  reg({ id: 'nuevo', texto: 'cmd-nuevo', habilitado: () => true, ejecutar: () => { if (hist && hist.sucio() && !confirm(tx('salir-sin-descargar'))) return; presentacionNueva(); } });
  reg({ id: 'abrir', texto: 'cmd-abrir', habilitado: () => true, ejecutar: () => { $('entradaArchivo').click(); } });
  reg({ id: 'anterior', texto: 'cmd-anterior', atajo: ['pageup', 'alt+arrowup'], ejecutar: () => mostrarDiapositiva(actual - 1), habilitado: () => hayPres() && actual > 0 });
  reg({ id: 'siguiente', texto: 'cmd-siguiente', atajo: ['pagedown', 'alt+arrowdown'], ejecutar: () => mostrarDiapositiva(actual + 1), habilitado: () => hayPres() && actual < cantidad - 1 });
  reg({ id: 'primera', texto: 'cmd-primera', atajo: 'mod+home', ejecutar: () => mostrarDiapositiva(0), habilitado: () => hayPres() && actual > 0 });
  reg({ id: 'ultima', texto: 'cmd-ultima', atajo: 'mod+end', ejecutar: () => mostrarDiapositiva(cantidad - 1), habilitado: () => hayPres() && actual < cantidad - 1 });
  reg({ id: 'nuevaDiap', texto: 'cmd-nueva-diap', atajo: 'mod+m', ejecutar: nuevaDiapositiva, habilitado: editable });
  reg({ id: 'duplicarDiap', texto: 'cmd-duplicar-diap', atajo: 'mod+shift+d', ejecutar: duplicarDiapositiva, habilitado: editable });
  reg({ id: 'borrarDiap', texto: 'cmd-borrar-diap', ejecutar: borrarDiapositiva, habilitado: () => editable() && cantidad > 1 });
  reg({ id: 'subirDiap', texto: 'cmd-subir-diap', ejecutar: () => moverDiapositiva(actual, actual - 1), habilitado: () => editable() && actual > 0 });
  reg({ id: 'bajarDiap', texto: 'cmd-bajar-diap', ejecutar: () => moverDiapositiva(actual, actual + 1), habilitado: () => editable() && actual < cantidad - 1 });
  reg({ id: 'eliminarForma', texto: 'cmd-eliminar-forma', atajo: ['delete', 'backspace'], ejecutar: async () => { const p = enSel('Eliminar', (i, c) => pe.opBorrar(i, c)); sel = null; await p; }, habilitado: haySel });
  reg({ id: 'frente', texto: 'cmd-frente', atajo: 'mod+shift+]', ejecutar: () => enSel('Orden', (i, c) => pe.opOrden(i, c, 'frente')), habilitado: haySel });
  reg({ id: 'fondo', texto: 'cmd-fondo', atajo: 'mod+shift+[', ejecutar: () => enSel('Orden', (i, c) => pe.opOrden(i, c, 'fondo')), habilitado: haySel });
  reg({ id: 'editarTexto', texto: 'cmd-editar-texto', atajo: 'f2', ejecutar: editarTexto, habilitado: () => conTexto() && !editor });
  reg({ id: 'negrita', texto: 'cmd-negrita', atajo: 'mod+b', ejecutar: () => formatear({ negrita: !formatoActual().negrita }), habilitado: conTexto, activo: () => !!formatoActual().negrita });
  reg({ id: 'cursiva', texto: 'cmd-cursiva', atajo: 'mod+i', ejecutar: () => formatear({ cursiva: !formatoActual().cursiva }), habilitado: conTexto, activo: () => !!formatoActual().cursiva });
  reg({ id: 'subrayado', texto: 'cmd-subrayado', atajo: 'mod+u', ejecutar: () => formatear({ subrayado: !formatoActual().subrayado }), habilitado: conTexto, activo: () => !!formatoActual().subrayado });
  reg({ id: 'vinetas', texto: 'cmd-vinetas', ejecutar: () => formatearParrafos({ vineta: tipoVineta() === 'caracter' ? 'ninguna' : 'punto' }), habilitado: conTexto, activo: () => tipoVineta() === 'caracter' });
  reg({ id: 'numeracion', texto: 'cmd-numeracion', ejecutar: () => formatearParrafos({ vineta: tipoVineta() === 'numero' ? 'ninguna' : 'numero' }), habilitado: conTexto, activo: () => tipoVineta() === 'numero' });
  reg({ id: 'nivelMas', texto: 'cmd-nivel-mas', ejecutar: () => formatearParrafos({ nivel: 1 }), habilitado: conTexto });
  reg({ id: 'nivelMenos', texto: 'cmd-nivel-menos', ejecutar: () => formatearParrafos({ nivel: -1 }), habilitado: conTexto });
  reg({ id: 'tamanoMas', texto: 'cmd-tamano-mas', ejecutar: () => formatear({ tamano: Math.min(400, (formatoActual().tamano || 18) + 2) }), habilitado: conTexto });
  reg({ id: 'tamanoMenos', texto: 'cmd-tamano-menos', ejecutar: () => formatear({ tamano: Math.max(6, (formatoActual().tamano || 18) - 2) }), habilitado: conTexto });
  reg({ id: 'colorTexto', texto: 'cmd-color-texto', ejecutar: (hex) => hex && formatear({ color: String(hex).replace('#', '') }), habilitado: conTexto });
  reg({ id: 'insertarTexto', texto: 'cmd-insertar-texto', ejecutar: insertarCuadroTexto, habilitado: editable });
  reg({ id: 'insertarForma', texto: 'cmd-insertar-forma', ejecutar: insertarForma, habilitado: editable });
  reg({ id: 'insertarImagen', texto: 'cmd-insertar-imagen', ejecutar: () => entradaImagen.click(), habilitado: editable });
  reg({ id: 'pdf', texto: 'cmd-pdf', ejecutar: exportarPdf, habilitado: hayPres });
  reg({ id: 'deseleccionar', texto: 'cmd-deseleccionar', ejecutar: () => seleccionar(null), habilitado: () => !!sel, oculto: true });
  reg({ id: 'presentar', texto: 'cmd-presentar', atajo: 'f5', ejecutar: () => presentar(0), habilitado: hayPres });
  reg({ id: 'presentarAqui', texto: 'cmd-presentar-aqui', atajo: 'shift+f5', ejecutar: () => presentar(actual), habilitado: hayPres });
  reg({ id: 'inspector', texto: 'cmd-inspector', ejecutar: () => insp.alternar(), activo: () => insp.visible(), habilitado: () => true });
  reg({ id: 'panel', texto: 'cmd-panel', ejecutar: () => { panel.hidden = !panel.hidden; }, activo: () => !panel.hidden, habilitado: () => true });
  reg({ id: 'paleta', texto: 'cmd-paleta', atajo: 'mod+k', enTexto: true, ejecutar: () => Oficina.paleta(comandos), oculto: true, habilitado: () => true });
  reg({ id: 'completa', texto: 'cmd-completa', atajo: 'f11', ejecutar: () => { if (document.fullscreenElement) document.exitFullscreen(); else $('app').requestFullscreen && $('app').requestFullscreen(); }, habilitado: () => true });
  reg({ id: 'modoEdicion', texto: 'cmd-modo-editar', habilitado: hayPres, ejecutar: () => { lectura = !lectura; aplicarModo(); comandos.refrescar(); } });

  const entradaImagen = crearEl('input', { type: 'file', accept: 'image/png,image/jpeg,image/gif', hidden: true, id: 'entradaImagen' });
  entradaImagen.addEventListener('change', () => { const a = entradaImagen.files[0]; entradaImagen.value = ''; insertarImagen(a); });
  document.body.append(entradaImagen);
  // los botones de la cinta no le quitan el foco al cuadro de texto: el formato va a lo seleccionado sin cerrarlo
  $('cinta').addEventListener('mousedown', (ev) => { if (editor && ev.target.closest('button')) ev.preventDefault(); });
  Oficina.cinta($('pestanas'), $('cinta'), [
    { id: 'inicio', texto: 'tab-inicio', grupos: [
      { texto: 'g-diapositivas', items: [{ comando: 'nuevaDiap', icono: '＋', conTexto: true }, { comando: 'duplicarDiap', icono: '⧉' }, { comando: 'borrarDiap', icono: '🗑' }, { comando: 'subirDiap', icono: '▲' }, { comando: 'bajarDiap', icono: '▼' }] },
      { texto: 'g-texto', items: [{ comando: 'negrita', icono: 'B' }, { comando: 'cursiva', icono: 'I' }, { comando: 'subrayado', icono: 'U' }, { comando: 'tamanoMenos', icono: 'A−' }, { comando: 'tamanoMas', icono: 'A+' }, { comando: 'colorTexto', tipo: 'color', icono: 'A', porDefecto: '#000000' }, { comando: 'editarTexto', icono: '✎' }] },
      { texto: 'g-parrafo', items: [{ comando: 'vinetas', icono: '•≡' }, { comando: 'numeracion', icono: '1.' }, { comando: 'nivelMenos', icono: '⇤' }, { comando: 'nivelMas', icono: '⇥' }] },
      { texto: 'g-insertar', items: [{ comando: 'insertarTexto', icono: '𝐓', conTexto: true }, { comando: 'insertarForma', icono: '◇', conTexto: true }, { comando: 'insertarImagen', icono: '🖼', conTexto: true }] },
      { texto: 'g-orden', items: [{ comando: 'frente', icono: '⤒' }, { comando: 'fondo', icono: '⤓' }, { comando: 'eliminarForma', icono: '✕' }] },
      { texto: 'g-presentar', items: [{ comando: 'presentar', icono: '▶', conTexto: true }, { comando: 'presentarAqui', icono: '⏵' }] },
      { texto: 'g-archivo', solo: 'pc', items: [{ comando: 'nuevo', icono: '📄' }, { comando: 'abrir', icono: '📂' }, { comando: 'descargar', icono: '⤓' }, { comando: 'pdf', icono: '⎙' }] },
    ] },
    { id: 'vista', texto: 'tab-vista', grupos: [
      { texto: 'g-diapositivas', items: [{ comando: 'anterior', icono: '◀', conTexto: true }, { comando: 'siguiente', icono: '▶', conTexto: true }] },
      { texto: 'g-ventana', items: [{ comando: 'panel', icono: '▤', conTexto: true }, { comando: 'inspector', icono: '☰', conTexto: true }, { comando: 'completa', icono: '⛶', conTexto: true }, { comando: 'paleta', icono: '⌘', conTexto: true }] },
    ] },
  ], comandos);
  comandos.conectarBotones(document.querySelector('.of-acciones'));
  Oficina.menu($('botonMas'), [{ comando: 'nuevo' }, { comando: 'abrir' }, '-', { comando: 'descargar' }, { comando: 'pdf' }, '-', { comando: 'presentar' }, { comando: 'presentarAqui' }, '-', { comando: 'duplicarDiap' }, { comando: 'borrarDiap' }, { comando: 'primera' }, { comando: 'ultima' }, { comando: 'panel' }, { comando: 'paleta' }, { comando: 'completa' }, { comando: 'inspector' }], comandos);
  comandos.conectarTeclado(document);

  $('nombreDoc').addEventListener('change', () => { nombreDoc = $('nombreDoc').value.trim().replace(/[\\/:*?"<>|]+/g, '-') || nombreDoc; $('nombreDoc').value = nombreDoc; });
  $('nombreDoc').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); e.target.blur(); } });

  // ---- inicio: abrir, nuevo, borradores ------------------------------------------------------------------------------------
  conectarZonaArrastre($('zonaArrastre'), $('entradaArchivo'), (archivos) => { if (archivos[0]) abrirArchivo(archivos[0]); });
  $('botonNuevo').addEventListener('click', presentacionNueva);
  async function mostrarBorradores() {
    const lista = await autoguardado.listar();
    const ul = $('listaBorradores');
    ul.textContent = '';
    $('seccionBorradores').hidden = !lista.length;
    for (const b of lista) {
      const li = document.createElement('li');
      const info = document.createElement('span');
      const nom = document.createElement('b'), fecha = document.createElement('small');
      nom.textContent = b.nombre;
      fecha.textContent = new Date(b.fecha).toLocaleString(idioma() === 'en' ? 'en-US' : 'es-CR');
      info.append(nom, ' ', fecha);
      const rec = crearEl('button', { type: 'button', class: 'boton' }, [tx('recuperar')]);
      rec.addEventListener('click', async () => { const r = await autoguardado.leer(b.id); if (r) { mostrarCargando(true); try { await abrirBytes(new Uint8Array(r.bytes), r.nombre, { id: r.id, recuperado: true }); } catch (e) { mostrarError(e); } mostrarCargando(false); } });
      const des = crearEl('button', { type: 'button', class: 'boton boton-secundario' }, [tx('descartar')]);
      des.addEventListener('click', async () => { await autoguardado.borrar(b.id); mostrarBorradores(); });
      const acciones = document.createElement('span'); acciones.append(rec, ' ', des);
      li.append(info, acciones);
      ul.append(li);
    }
  }
  mostrarBorradores();
  // El idioma se cambia sin recargar la página (botón del inicio): lo que se escribió con tx() al abrir se vuelve a escribir
  document.addEventListener('cambioIdioma', () => {
    mostrarBorradores();
    if (!modelo) return;
    aplicarModo(); actualizarGuardado(); actualizarInspector(); pintarMarcadores();
    lineaHonesta(diapActual);
    lienzoPrincipal.setAttribute('aria-label', tx('diapositiva-de', actual + 1, cantidad));
    for (const b of panel.querySelectorAll('.pr-mini')) b.setAttribute('aria-label', tx('diapositiva-n', +b.dataset.indice + 1));
  });
  window.addEventListener('beforeunload', (e) => { if (hist && hist.sucio()) { e.preventDefault(); e.returnValue = ''; } });

  // para las pruebas de navegador
  window.PresentacionesApp = {
    abrirBytes, comandos, seleccionar, insertarForma, insertarImagen, exportarPdf, get editor() { return editor; }, get cosas() { return cosas; }, objetoEn, presentar, salirPresentacion, moverDiapositiva, tamanoConAsa, claveDe,
    get pe() { return pe; }, get modelo() { return modelo; }, get hist() { return hist; }, get actual() { return actual; }, get cantidad() { return cantidad; },
    get sel() { return sel; }, get elementos() { return seleccionables(); }, get presentando() { return !!pres; }, descargar: () => pe.generar(),
  };
})();

/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Documentos (Fase 6, etapa 6.9): la aplicación completa en el esqueleto del arquetipo G (suite/nucleo/oficina.js). Abre un .docx o uno en blanco, lo muestra
// hoja a hoja en Canvas y deja escribirlo, darle formato, insertar tablas, imágenes, saltos, encabezados y pies, buscar y reemplazar, pegar con formato,
// revisar la ortografía y descargarlo como .docx o PDF. El trabajo se guarda solo en este navegador (borradores) y en el celular se abre para leer.
// Toda la lógica de edición vive en los motores (edicion-doc, formato-doc…) y en editor-doc.js; aquí solo se conectan a la cinta, los atajos y la página.
(function () {
  'use strict';
  const ES = {
    'titulo-pagina': 'Editar Word (.docx) y OpenDocument (.odt) en el navegador — Mini Cookies', 'meta-descripcion': 'Edita documentos de Word (.docx) y de LibreOffice (.odt) en el navegador: escribe, da formato, inserta tablas e imágenes, revisa la ortografía y descarga un PDF, sin instalar nada y sin subir tu archivo. Gratis y sin registro.',
    'h1': 'Editar Word (Documentos)', 'intro': 'Abre un documento de Word o de LibreOffice, léelo y edítalo. Al guardar, lo que no tocaste queda exactamente como estaba.',
    'zona-arrastre': '📄 Arrastra tu archivo aquí o haz clic para elegirlo (.docx, .odt)', 'boton-nuevo': 'Documento en blanco', 'cargando': 'Abriendo el documento…', 'anuncio': 'Espacio publicitario',
    'borradores': 'Trabajo sin terminar en este navegador', 'recuperar': 'Recuperar', 'descartar': 'Descartar', 'nombre-app': 'Documentos', 'volver-inicio': 'Volver al inicio',
    'nombre-archivo': 'Nombre del archivo', 'mas-opciones': 'Más opciones', 'cinta': 'Cinta', 'inspector': 'Propiedades', 'cmd-descargar-corto': 'Descargar',
    'tema-aria': 'Cambiar a modo oscuro', 'tema-title': 'Cambiar tema',
    'tab-inicio': 'Inicio', 'tab-insertar': 'Insertar', 'tab-revisar': 'Revisar', 'tab-vista': 'Vista',
    'g-archivo': 'Archivo', 'g-estilo': 'Estilo', 'g-fuente': 'Fuente', 'g-parrafo': 'Párrafo', 'g-ilustraciones': 'Ilustraciones', 'g-saltos': 'Saltos', 'g-encabezado': 'Encabezado y pie',
    'g-tabla-cambiar': 'Filas y columnas', 'g-buscar': 'Buscar', 'g-ortografia': 'Ortografía', 'g-zoom': 'Zoom', 'g-ventana': 'Ventana',
    'cmd-nuevo': 'Documento en blanco', 'cmd-abrir': 'Abrir…', 'cmd-descargar': 'Descargar', 'cmd-convertir-odt': 'Descargar como .odt (conversión)', 'cmd-convertir-docx': 'Descargar como .docx (conversión)', 'error-otro-odf': 'Este archivo de OpenDocument no es un documento de texto (.odt). Una hoja se abre en Hojas y una presentación en Presentaciones.', 'aviso-conversion': 'Convertido. La conversión entre Word y OpenDocument no conserva todo (estilos finos, notas, campos): revisa el resultado.', 'cosa-indices': 'índices', 'cosa-protegidas': 'secciones protegidas', 'cosa-firmas': 'firmas', 'cosa-ruby': 'texto ruby', 'cmd-pdf': 'Descargar PDF', 'cmd-deshacer': 'Deshacer', 'cmd-rehacer': 'Rehacer',
    'cmd-modo-editar': 'Editar', 'cmd-modo-leer': 'Leer', 'cmd-estilo': 'Estilo de párrafo', 'cmd-fuente': 'Fuente', 'cmd-tamano': 'Tamaño de letra',
    'cmd-negrita': 'Negrita', 'cmd-cursiva': 'Cursiva', 'cmd-subrayado': 'Subrayado', 'cmd-tachado': 'Tachado', 'cmd-color-texto': 'Color del texto', 'cmd-resaltado': 'Resaltado',
    'cmd-superindice': 'Superíndice', 'cmd-subindice': 'Subíndice', 'cmd-limpiar': 'Quitar formato',
    'cmd-izquierda': 'Alinear a la izquierda', 'cmd-centro': 'Centrar', 'cmd-derecha': 'Alinear a la derecha', 'cmd-justificar': 'Justificar', 'cmd-interlineado': 'Interlineado',
    'cmd-vinetas': 'Viñetas', 'cmd-numeracion': 'Numeración', 'cmd-sangria-mas': 'Aumentar sangría', 'cmd-sangria-menos': 'Disminuir sangría',
    'cmd-tabla': 'Tabla…', 'cmd-imagen': 'Imagen…', 'cmd-salto-pagina': 'Salto de página', 'cmd-salto-seccion': 'Salto de sección (página siguiente)', 'cmd-salto-seccion-continuo': 'Salto de sección (continuo)',
    'cmd-encabezado': 'Encabezado', 'cmd-pie': 'Pie de página', 'cmd-num-pagina': 'Número de página', 'cmd-total-paginas': 'Total de páginas', 'cmd-cerrar-parte': 'Cerrar encabezado y pie',
    'cmd-fila-antes': 'Fila arriba', 'cmd-fila-despues': 'Fila abajo', 'cmd-col-antes': 'Columna a la izquierda', 'cmd-col-despues': 'Columna a la derecha',
    'cmd-quitar-fila': 'Quitar fila', 'cmd-quitar-col': 'Quitar columna', 'cmd-quitar-tabla': 'Quitar tabla',
    'cmd-buscar': 'Buscar y reemplazar', 'cmd-ortografia': 'Corrector ortográfico', 'cmd-idioma-orto': 'Idioma del corrector', 'cmd-mas-zoom': 'Acercar', 'cmd-menos-zoom': 'Alejar', 'cmd-ajustar': 'Ajustar al ancho',
    'cmd-inspector': 'Panel de propiedades', 'cmd-paleta': 'Paleta de comandos', 'cmd-completa': 'Pantalla completa',
    'est-normal': 'Normal', 'est-titulo': 'Título', 'est-subtitulo': 'Subtítulo', 'est-t1': 'Título 1', 'est-t2': 'Título 2', 'est-t3': 'Título 3',
    'idioma-es': 'Español', 'idioma-en': 'English',
    'b-buscar': 'Buscar', 'b-reemplazar-con': 'Reemplazar con', 'b-mayus': 'Mayúsculas', 'b-palabra': 'Palabra completa', 'b-anterior': 'Coincidencia anterior', 'b-siguiente': 'Coincidencia siguiente',
    'b-reemplazar': 'Reemplazar', 'b-todo': 'Reemplazar todo', 'b-cerrar': 'Cerrar', 'b-una': '1 coincidencia', 'b-varias': '{0} coincidencias', 'b-ninguna': 'Sin coincidencias', 'b-reemplazos': '{0} reemplazos', 'b-un-reemplazo': '1 reemplazo',
    'p-encabezado': 'Encabezado', 'p-pie': 'Pie de página', 'p-primera': ' de la primera página', 'p-pares': ' de las páginas pares', 'p-heredado': ' (el de la sección anterior)',
    'o-sin-sugerencias': 'Sin sugerencias', 'o-omitir': 'Omitir todo', 'o-agregar': 'Agregar al diccionario',
    'tabla-titulo': 'Insertar tabla', 'tabla-filas': 'Filas', 'tabla-columnas': 'Columnas', 'aceptar': 'Aceptar', 'cancelar': 'Cancelar',
    'estado-listo': 'Listo', 'sin-cambios': 'Sin cambios', 'guardado-local': 'Guardado en este navegador a las {0}', 'cambios-sin-descargar': 'Cambios sin descargar',
    'honesta-edita': 'Aquí se edita el texto, su formato, las tablas, las imágenes en línea, los saltos, los encabezados y los pies.', 'honesta-conserva': 'Se conserva al guardar, sin editar aquí: {0}.',
    'honesta-aprox': 'Se dibuja de forma aproximada: {0}.',
    'cosa-notas': 'notas al pie', 'cosa-comentarios': 'comentarios', 'cosa-cambios': 'cambios con control de cambios', 'cosa-cuadros': 'cuadros de texto y formas', 'cosa-ecuaciones': 'ecuaciones',
    'cosa-controles': 'controles de contenido', 'cosa-columnas': 'columnas', 'cosa-graficos': 'gráficos y diagramas', 'cosa-macros': 'macros', 'cosa-flotantes': 'imágenes flotantes', 'cosa-campos': 'campos (índices, fechas…)', 'cosa-incrustados': 'objetos incrustados',
    'insp-documento': 'Documento', 'insp-seleccion': 'Selección', 'insp-hojas': 'Hojas', 'insp-palabras': 'Palabras', 'insp-estilo': 'Estilo', 'insp-fuente': 'Fuente', 'insp-tamano': 'Tamaño', 'insp-alineacion': 'Alineación', 'insp-modo': 'Modo', 'insp-edicion': 'Edición', 'insp-lectura': 'Lectura',
    'aviso-pdf': 'Preparando el PDF…', 'aviso-pdf-listo': 'PDF listo: {0} hojas', 'aviso-pdf-listo-1': 'PDF listo: 1 hoja', 'error-hojas': 'El documento tiene demasiadas hojas.', 'aviso-no-se-pudo': 'No se pudo', 'aviso-tabla': 'Aquí no se puede insertar una tabla (solo en el cuerpo del documento).',
    'aviso-salto-parte': 'En un encabezado o un pie no se pueden poner saltos ni imágenes.', 'aviso-una-seccion': 'Este documento tiene varias secciones: el {0} se crea en la última.', 'aviso-estructura': 'Esta tabla tiene una estructura que el editor todavía no cambia.',
    'aviso-secciones': 'Los saltos de sección solo se insertan en el cuerpo del documento.', 'aviso-imagen': 'No se pudo leer la imagen.', 'aviso-imagen-grande': 'La imagen pesa más de 20 MB.', 'aviso-orto': 'No se pudo cargar el corrector',
    'error-abrir': 'El archivo no se pudo abrir', 'salir-sin-descargar': 'Hay cambios sin descargar.',
    'capa-imagen': 'Imagen', 'capa-dibujo': 'Dibujo', 'capa-notas': 'Notas al pie',
  };
  const EN = {
    'titulo-pagina': 'Edit Word (.docx) and OpenDocument (.odt) files in your browser — Mini Cookies', 'meta-descripcion': 'Edit Word (.docx) and LibreOffice (.odt) documents in your browser: type, format, insert tables and pictures, check spelling and download a PDF, without installing anything or uploading your file. Free, no sign-up.',
    'h1': 'Edit Word (Documents)', 'intro': 'Open a Word or LibreOffice document, read it and edit it. When you save, everything you didn’t touch stays exactly as it was.',
    'zona-arrastre': '📄 Drop your file here or click to choose it (.docx, .odt)', 'boton-nuevo': 'Blank document', 'cargando': 'Opening the document…', 'anuncio': 'Advertising space',
    'borradores': 'Unfinished work in this browser', 'recuperar': 'Recover', 'descartar': 'Discard', 'nombre-app': 'Documents', 'volver-inicio': 'Back to start',
    'nombre-archivo': 'File name', 'mas-opciones': 'More options', 'cinta': 'Ribbon', 'inspector': 'Properties', 'cmd-descargar-corto': 'Download',
    'tema-aria': 'Switch to dark mode', 'tema-title': 'Change theme',
    'tab-inicio': 'Home', 'tab-insertar': 'Insert', 'tab-revisar': 'Review', 'tab-vista': 'View',
    'g-archivo': 'File', 'g-estilo': 'Style', 'g-fuente': 'Font', 'g-parrafo': 'Paragraph', 'g-ilustraciones': 'Illustrations', 'g-saltos': 'Breaks', 'g-encabezado': 'Header & footer',
    'g-tabla-cambiar': 'Rows & columns', 'g-buscar': 'Find', 'g-ortografia': 'Spelling', 'g-zoom': 'Zoom', 'g-ventana': 'Window',
    'cmd-nuevo': 'Blank document', 'cmd-abrir': 'Open…', 'cmd-descargar': 'Download', 'cmd-convertir-odt': 'Download as .odt (conversion)', 'cmd-convertir-docx': 'Download as .docx (conversion)', 'error-otro-odf': 'This OpenDocument file is not a text document (.odt). Open spreadsheets in Sheets and presentations in Presentations.', 'aviso-conversion': 'Converted. Converting between Word and OpenDocument doesn’t keep everything (fine styles, notes, fields): check the result.', 'cosa-indices': 'indexes', 'cosa-protegidas': 'protected sections', 'cosa-firmas': 'signatures', 'cosa-ruby': 'ruby text', 'cmd-pdf': 'Download PDF', 'cmd-deshacer': 'Undo', 'cmd-rehacer': 'Redo',
    'cmd-modo-editar': 'Edit', 'cmd-modo-leer': 'Read', 'cmd-estilo': 'Paragraph style', 'cmd-fuente': 'Font', 'cmd-tamano': 'Font size',
    'cmd-negrita': 'Bold', 'cmd-cursiva': 'Italic', 'cmd-subrayado': 'Underline', 'cmd-tachado': 'Strikethrough', 'cmd-color-texto': 'Text color', 'cmd-resaltado': 'Highlight',
    'cmd-superindice': 'Superscript', 'cmd-subindice': 'Subscript', 'cmd-limpiar': 'Clear formatting',
    'cmd-izquierda': 'Align left', 'cmd-centro': 'Center', 'cmd-derecha': 'Align right', 'cmd-justificar': 'Justify', 'cmd-interlineado': 'Line spacing',
    'cmd-vinetas': 'Bullets', 'cmd-numeracion': 'Numbering', 'cmd-sangria-mas': 'Increase indent', 'cmd-sangria-menos': 'Decrease indent',
    'cmd-tabla': 'Table…', 'cmd-imagen': 'Picture…', 'cmd-salto-pagina': 'Page break', 'cmd-salto-seccion': 'Section break (next page)', 'cmd-salto-seccion-continuo': 'Section break (continuous)',
    'cmd-encabezado': 'Header', 'cmd-pie': 'Footer', 'cmd-num-pagina': 'Page number', 'cmd-total-paginas': 'Page count', 'cmd-cerrar-parte': 'Close header and footer',
    'cmd-fila-antes': 'Row above', 'cmd-fila-despues': 'Row below', 'cmd-col-antes': 'Column to the left', 'cmd-col-despues': 'Column to the right',
    'cmd-quitar-fila': 'Delete row', 'cmd-quitar-col': 'Delete column', 'cmd-quitar-tabla': 'Delete table',
    'cmd-buscar': 'Find and replace', 'cmd-ortografia': 'Spell checker', 'cmd-idioma-orto': 'Spell checker language', 'cmd-mas-zoom': 'Zoom in', 'cmd-menos-zoom': 'Zoom out', 'cmd-ajustar': 'Fit to width',
    'cmd-inspector': 'Properties panel', 'cmd-paleta': 'Command palette', 'cmd-completa': 'Full screen',
    'est-normal': 'Normal', 'est-titulo': 'Title', 'est-subtitulo': 'Subtitle', 'est-t1': 'Heading 1', 'est-t2': 'Heading 2', 'est-t3': 'Heading 3',
    'idioma-es': 'Español', 'idioma-en': 'English',
    'b-buscar': 'Find', 'b-reemplazar-con': 'Replace with', 'b-mayus': 'Match case', 'b-palabra': 'Whole word', 'b-anterior': 'Previous match', 'b-siguiente': 'Next match',
    'b-reemplazar': 'Replace', 'b-todo': 'Replace all', 'b-cerrar': 'Close', 'b-una': '1 match', 'b-varias': '{0} matches', 'b-ninguna': 'No matches', 'b-reemplazos': '{0} replacements', 'b-un-reemplazo': '1 replacement',
    'p-encabezado': 'Header', 'p-pie': 'Footer', 'p-primera': ' of the first page', 'p-pares': ' of even pages', 'p-heredado': ' (the one from the previous section)',
    'o-sin-sugerencias': 'No suggestions', 'o-omitir': 'Ignore all', 'o-agregar': 'Add to dictionary',
    'tabla-titulo': 'Insert table', 'tabla-filas': 'Rows', 'tabla-columnas': 'Columns', 'aceptar': 'OK', 'cancelar': 'Cancel',
    'estado-listo': 'Ready', 'sin-cambios': 'No changes', 'guardado-local': 'Saved in this browser at {0}', 'cambios-sin-descargar': 'Changes not downloaded',
    'honesta-edita': 'Here you edit text, its formatting, tables, inline pictures, breaks, headers and footers.', 'honesta-conserva': 'Kept as is when saving, not editable here: {0}.',
    'honesta-aprox': 'Drawn approximately: {0}.',
    'cosa-notas': 'footnotes', 'cosa-comentarios': 'comments', 'cosa-cambios': 'tracked changes', 'cosa-cuadros': 'text boxes and shapes', 'cosa-ecuaciones': 'equations',
    'cosa-controles': 'content controls', 'cosa-columnas': 'columns', 'cosa-graficos': 'charts and diagrams', 'cosa-macros': 'macros', 'cosa-flotantes': 'floating pictures', 'cosa-campos': 'fields (indexes, dates…)', 'cosa-incrustados': 'embedded objects',
    'insp-documento': 'Document', 'insp-seleccion': 'Selection', 'insp-hojas': 'Pages', 'insp-palabras': 'Words', 'insp-estilo': 'Style', 'insp-fuente': 'Font', 'insp-tamano': 'Size', 'insp-alineacion': 'Alignment', 'insp-modo': 'Mode', 'insp-edicion': 'Editing', 'insp-lectura': 'Reading',
    'aviso-pdf': 'Preparing the PDF…', 'aviso-pdf-listo': 'PDF ready: {0} pages', 'aviso-pdf-listo-1': 'PDF ready: 1 page', 'error-hojas': 'The document has too many pages.', 'aviso-no-se-pudo': 'Couldn’t do it', 'aviso-tabla': 'A table can’t be inserted here (only in the body of the document).',
    'aviso-salto-parte': 'Breaks and pictures can’t go in a header or footer.', 'aviso-una-seccion': 'This document has several sections: the {0} is created in the last one.', 'aviso-estructura': 'This table has a structure the editor can’t change yet.',
    'aviso-secciones': 'Section breaks can only be inserted in the body of the document.', 'aviso-imagen': 'The picture couldn’t be read.', 'aviso-imagen-grande': 'The picture is larger than 20 MB.', 'aviso-orto': 'The spell checker couldn’t be loaded',
    'error-abrir': 'The file could not be opened', 'salir-sin-descargar': 'There are changes not downloaded.',
    'capa-imagen': 'Picture', 'capa-dibujo': 'Drawing', 'capa-notas': 'Footnotes',
  };
  const pagina = window.i18nPagina || (window.i18nPagina = {});
  pagina.es = Object.assign(pagina.es || {}, ES);
  pagina.en = Object.assign(pagina.en || {}, EN);
  const tx = (k, ...a) => { let s = typeof window.t === 'function' ? window.t(k) : (ES[k] || k); a.forEach((x, i) => { s = s.replace('{' + i + '}', x); }); return s; };
  const idioma = () => (typeof window.idiomaActual === 'function' && window.idiomaActual() === 'en' ? 'en' : 'es');
  const $ = (id) => document.getElementById(id);
  const crearEl = Oficina.crearEl;

  // ───────── Estado ─────────
  let modelo = null, edDoc = null, editorDoc = null, mqDoc = null, pintorDoc = null, visor = null, capa = null;
  let ed = null, editor = null;           // los de lo que se está editando ahora: el documento o, mientras está abierto, su encabezado o pie
  let parte = null;                       // { vista, ed, editor, visor, tipo }
  let docId = null, nombreDoc = 'Documento', lectura = false, partesCambiadas = false, cosas = new Set(), aproximado = new Set();
  const autoguardado = Autoguardado.crear('documentos');
  const comandos = Comandos.crear();
  const sucio = () => !!(edDoc && edDoc.sucio()) || partesCambiadas;

  // ───────── La interfaz (arquetipo G) ─────────
  const interfaz = Oficina.montar($('app'), { nombre: 'nombre-app', inicio: '../', riel: false });
  const insp = Oficina.inspector(interfaz.inspector);
  const barraBuscar = crearEl('div', { class: 'dc-buscar', id: 'barraBuscar', role: 'search', hidden: true });
  const bTexto = crearEl('input', { type: 'search', id: 'bTexto', placeholder: tx('b-buscar'), 'aria-label': tx('b-buscar'), autocomplete: 'off', spellcheck: 'false' });
  const bNuevo = crearEl('input', { type: 'text', id: 'bNuevo', name: 'reemplazo', placeholder: tx('b-reemplazar-con'), 'aria-label': tx('b-reemplazar-con'), autocomplete: 'off', spellcheck: 'false' });
  const bMayus = crearEl('input', { type: 'checkbox', id: 'bMayus' }), bPalabra = crearEl('input', { type: 'checkbox', id: 'bPalabra' });
  const bCuenta = crearEl('span', { id: 'bCuenta', 'aria-live': 'polite' });
  const botonB = (id, texto, aria) => crearEl('button', { type: 'button', id, 'aria-label': aria || texto }, [texto]);
  const bAnterior = botonB('bAnterior', '↑', tx('b-anterior')), bSiguiente = botonB('bSiguiente', '↓', tx('b-siguiente')), bReemplazar = botonB('bReemplazar', tx('b-reemplazar')), bTodo = botonB('bTodo', tx('b-todo')), bCerrar = botonB('bCerrar', '✕', tx('b-cerrar'));
  barraBuscar.append(bTexto, bNuevo, crearEl('label', {}, [bMayus, ' ' + tx('b-mayus')]), crearEl('label', {}, [bPalabra, ' ' + tx('b-palabra')]), bAnterior, bSiguiente, bReemplazar, bTodo, bCuenta, bCerrar);
  const panelParte = crearEl('section', { class: 'dc-parte', id: 'panelParte', hidden: true });
  const parteTitulo = crearEl('strong', { id: 'parteTitulo' });
  const parteBarra = crearEl('div', { class: 'dc-parte-barra' }, [parteTitulo]);
  const contParte = crearEl('div', { class: 'dc-visor dc-visor-parte', id: 'visorParte', tabindex: '0' });
  panelParte.append(parteBarra, contParte);
  const cont = crearEl('div', { class: 'dc-visor', id: 'visor', tabindex: '0' });
  interfaz.lienzo.append(barraBuscar, panelParte, cont);
  interfaz.honesta.textContent = tx('estado-listo');

  // ───────── Abrir ─────────
  const mostrarCargando = (si) => { $('cargando').hidden = !si; if (si) $('errorInicio').hidden = true; };
  function mostrarError(e) {
    console.error(e);
    const el = $('errorInicio');
    el.textContent = tx('error-abrir') + ': ' + textoError(e);
    el.hidden = false; $('inicio').hidden = false; $('app').hidden = true; document.documentElement.classList.remove('of-abierta');
  }
  async function abrirArchivo(archivo) {
    mostrarCargando(true);
    try { await abrirBytes(new Uint8Array(await archivo.arrayBuffer()), archivo.name.replace(/\.[^.]+$/, '')); } catch (e) { mostrarError(e); }
    mostrarCargando(false);
  }
  async function documentoNuevo() {
    mostrarCargando(true);
    try { await abrirBytes(await DocumentoNuevo.crear({ idioma: idioma() }), idioma() === 'en' ? 'Document1' : 'Documento1', { nuevo: true }); } catch (e) { mostrarError(e); }
    mostrarCargando(false);
  }
  // una hoja (.ods) o una presentación (.odp) no son un documento de texto: se dice con claridad en vez de quejarse de Word
  async function avisarSiOtroOpenDocument(bytes) {
    let tipo = '';
    try { const zip = await ZipPropio.loadAsync(bytes); tipo = zip.file('mimetype') ? (await zip.file('mimetype').async('string')).trim() : ''; } catch (e) { return; }
    if (/^application\/vnd\.oasis\.opendocument\./.test(tipo)) throw new Error(tx('error-otro-odf'));
  }
  async function abrirBytes(bytes, nombre, opciones) {
    const o = opciones || {};
    const formato = o.formato || (await OdtLector.parece(bytes) ? 'odt' : 'docx');
    if (formato === 'docx') await avisarSiOtroOpenDocument(bytes);
    const m = formato === 'odt' ? await OdtModelo.abrir(bytes) : await DocxModelo.abrir(bytes);
    if (parte) await cerrarParte(true);
    // lo último del documento anterior va ya a su borrador: el autoguardado programado correría después con el documento nuevo
    autoguardado.cancelar();
    if (edDoc && sucio()) { try { const d = await datosAutoguardado(); if (d) await autoguardado.guardar(d); } catch (e) { console.warn(e); } }
    ultimoAutoguardado = null;
    if (editorDoc) editorDoc.destruir();
    if (visor) visor.destruir();
    if (capa) capa.destruir();
    modelo = m;
    nombreDoc = nombre || 'Documento';
    docId = o.id || Autoguardado.nuevoId();
    partesCambiadas = false;
    $('inicio').hidden = true; $('app').hidden = false; document.documentElement.classList.add('of-abierta');
    $('nombreDoc').value = nombreDoc;
    mqDoc = MaquetadorDoc.crear(modelo, document.createElement('canvas').getContext('2d'), {});
    edDoc = EdicionDoc.crear(modelo, { invalidar: mqDoc.invalidar, alCambiar: alCambiarDoc });
    const diseno = await mqDoc.maquetar();
    pintorDoc = PintorDoc.crear(modelo);
    await pintorDoc.precargar(diseno);
    visor = VisorDoc.crear(cont, { diseno, pintor: pintorDoc, zoom: 1, alCambiarHoja: () => programarInspector() });
    capa = CapaAccesible.crear(cont, modelo, { titulo: nombreDoc, textos: { imagen: tx('capa-imagen'), dibujo: tx('capa-dibujo'), notas: tx('capa-notas') } });
    editorDoc = EditorDoc.crear({ contenedor: cont, visor, ed: edDoc, mq: mqDoc, pintor: pintorDoc, diseno, capa, alCambiar: alCambiarDoc, alSeleccionar, etiqueta: tx('nombre-app'), alError: (e) => avisar(tx('aviso-no-se-pudo') + ': ' + textoError(e)) });
    ed = edDoc; editor = editorDoc;
    lectura = matchMedia('(max-width: 900px)').matches && !o.nuevo;   // en el celular se abre para leer; «Editar» activa la edición
    aplicarModo();
    ajustarAlAncho(false);
    cosas = cosasDelArchivo(modelo);
    aproximado = aproximadoDelArchivo(modelo);
    pintarHonesta();
    aplicarCorrector();
    barraBuscar.hidden = true;
    actualizarGuardado(); refrescarTodo();
    if (o.recuperado) autoguardado.programar(datosAutoguardado, 10);
    if (!lectura) editorDoc.enfocar();
  }

  // ───────── Lo que cambia ─────────
  let temporizadorInsp = null, cacheEstado = null;
  const estadoActual = () => { if (!cacheEstado && editor) cacheEstado = editor.estado(); return cacheEstado || {}; };
  function alSeleccionar() { cacheEstado = null; comandos.refrescar(); programarInspector(); }
  function alCambiarDoc() {
    cacheEstado = null;
    if (parte && ed !== edDoc) { alCambiarParte(); return; }
    actualizarGuardado();
    comandos.refrescar();
    programarInspector();
    autoguardado.programar(datosAutoguardado, 2500, (fecha) => { if (fecha) actualizarGuardado(fecha); });
  }
  function refrescarTodo() { cacheEstado = null; comandos.refrescar(); pintarInspector(); }
  async function datosAutoguardado() { if (!edDoc || !sucio()) return null; await escribiendoParte; return { id: docId, nombre: nombreDoc, bytes: await edDoc.guardar() }; }
  let ultimoAutoguardado = null;
  function actualizarGuardado(fecha) {
    if (fecha) ultimoAutoguardado = fecha;
    const el = $('estadoGuardado');
    if (!edDoc || !sucio()) { el.textContent = tx('sin-cambios'); return; }
    el.textContent = ultimoAutoguardado ? tx('guardado-local', new Date(ultimoAutoguardado).toLocaleTimeString(idioma() === 'en' ? 'en-US' : 'es-CR', { hour: '2-digit', minute: '2-digit' })) : tx('cambios-sin-descargar');
  }
  const avisar = (texto) => Oficina.aviso($('avisoEstado'), texto, 5000);

  // ───────── Línea honesta: lo que se conserva sin poder editarse aquí ─────────
  const COSAS_ODT = { 'notas-pie': 'cosa-notas', 'notas-finales': 'cosa-notas', comentarios: 'cosa-comentarios', cambios: 'cosa-cambios', cuadros: 'cosa-cuadros', objetos: 'cosa-incrustados', dibujos: 'cosa-cuadros',
    'imagenes-flotantes': 'cosa-flotantes', campos: 'cosa-campos', macros: 'cosa-macros', indices: 'cosa-indices', 'secciones-protegidas': 'cosa-protegidas', firmas: 'cosa-firmas', ruby: 'cosa-ruby' };
  function cosasDelArchivo(m) {
    if (m.formato === 'odt') return new Set([...m.cosas].map((c) => COSAS_ODT[c]).filter(Boolean));
    const sal = new Set(), x = m.xml, rutas = m.paquete.rutas();
    if (m.notas && m.notas.size) sal.add('cosa-notas');
    if (rutas.includes('word/comments.xml')) sal.add('cosa-comentarios');
    if (/<w:(?:ins|del)\s/.test(x)) sal.add('cosa-cambios');
    if (/<w:txbxContent|<wps:wsp|<v:shape\b|<v:rect\b/.test(x)) sal.add('cosa-cuadros');
    if (/<m:oMath/.test(x)) sal.add('cosa-ecuaciones');
    if (/<w:sdt>/.test(x)) sal.add('cosa-controles');
    if (rutas.some((r) => /^word\/(charts|diagrams)\//.test(r))) sal.add('cosa-graficos');
    if (rutas.some((r) => /vbaProject\.bin$/.test(r))) sal.add('cosa-macros');
    if (rutas.some((r) => /^word\/embeddings\//.test(r))) sal.add('cosa-incrustados');
    if (/<wp:anchor\b/.test(x)) sal.add('cosa-flotantes');
    for (const f of x.matchAll(/(?:w:instr="|<w:instrText[^>]*>)\s*([A-Z]+)/g)) if (!/^(PAGE|NUMPAGES)$/.test(f[1])) { sal.add('cosa-campos'); break; }
    return sal;
  }
  function aproximadoDelArchivo(m) {
    const sal = new Set();
    if (m.secciones.some((s) => s.columnas)) sal.add('cosa-columnas');
    return sal;
  }
  function pintarHonesta() {
    if (!modelo) { interfaz.honesta.textContent = tx('estado-listo'); return; }
    const partes = [tx('honesta-edita')];
    if (cosas.size) partes.push(tx('honesta-conserva', [...cosas].map((k) => tx(k)).join(', ')));
    if (aproximado.size) partes.push(tx('honesta-aprox', [...aproximado].map((k) => tx(k)).join(', ')));
    interfaz.honesta.textContent = partes.join(' ');
  }

  // ───────── Inspector ─────────
  function programarInspector() { clearTimeout(temporizadorInsp); temporizadorInsp = setTimeout(pintarInspector, 150); }
  function pintarInspector() {
    if (!edDoc) { insp.vaciar(); return; }
    const e = estadoActual(), hojas = editorDoc.diseno().paginas.length;
    const palabras = edDoc.parrafos().reduce((n, b) => n + (b.texto.match(/[\p{L}\p{N}]+/gu) || []).length, 0);
    const alin = { left: '⯇', center: '≡', right: '⯈', both: '☰' }[e.jc] || '';
    insp.mostrar([
      { titulo: tx('insp-documento'), filas: [['insp-hojas', String(hojas)], ['insp-palabras', String(palabras)], ['insp-modo', tx(lectura ? 'insp-lectura' : 'insp-edicion')]] },
      { titulo: tx('insp-seleccion'), filas: [['insp-estilo', e.estilo || tx('est-normal')], ['insp-fuente', e.fuente || '—'], ['insp-tamano', e.tam ? String(e.tam).replace('.', ',') + ' pt' : '—'], ['insp-alineacion', alin]] },
    ]);
  }

  // ───────── Modo lectura ─────────
  function aplicarModo() {
    $('app').classList.toggle('of-lectura', lectura);
    $('botonModo').textContent = lectura ? '✎ ' + tx('cmd-modo-editar') : '👁 ' + tx('cmd-modo-leer');
    if (editorDoc) editorDoc.soloLectura(lectura);
    if (parte) parte.editor.soloLectura(lectura);
  }
  const editable = () => !!editor && !!ed && !lectura;

  // ───────── Zoom ─────────
  function ajustarAlAncho(siempre) {
    if (!visor || !editorDoc) return;
    const d = editorDoc.diseno().paginas[0];
    if (!d) return;
    const util = cont.clientWidth - 32, anchoPx = d.ancho * 96 / 72;
    const z = Math.min(siempre ? 3 : 1, Math.max(0.25, util / anchoPx));
    if (siempre || z < 1) visor.zoom(z);
    editorDoc.relayout();
  }
  function cambiarZoom(f) { if (!visor) return; visor.zoom(visor.zoom() * f); editorDoc.relayout(); }

  // los errores del maquetador traen un código y se dicen en el idioma de la página; el resto viene con su mensaje
  const textoError = (e) => (e && e.codigo === 'hojas' ? tx('error-hojas') : e && e.message ? e.message : String(e));
  // ───────── Descargar y PDF ─────────
  let escribiendoParte = Promise.resolve();
  async function descargar() {
    if (!edDoc) return;
    // lo que se escriba mientras el archivo se arma no está en él: se anota la versión de antes, y solo hasta ahí queda «guardado»
    let version, huboPartes = false;
    try {
      await escribiendoParte;
      version = edDoc.historial.version; huboPartes = partesCambiadas; partesCambiadas = false;
      const odt = modelo.formato === 'odt';
      const conMacros = !odt && modelo.paquete.rutas().some((r) => /vbaProject\.bin$/.test(r));
      const tipo = odt ? 'application/vnd.oasis.opendocument.text' : conMacros ? 'application/vnd.ms-word.document.macroEnabled.12' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
      descargarBlob(new Blob([await edDoc.guardar()], { type: tipo }), nombreDoc + (odt ? '.odt' : conMacros ? '.docm' : '.docx'));
      edDoc.historial.marcarGuardado(version);
      actualizarGuardado(); refrescarTodo();
    } catch (e) { partesCambiadas = partesCambiadas || huboPartes; console.warn(e); avisar(tx('aviso-no-se-pudo') + ': ' + textoError(e)); }
  }
  // Word ↔ OpenDocument es una conversión, no un guardado: se baja una copia y el documento abierto sigue como estaba
  async function convertir() {
    if (!edDoc) return;
    try {
      await escribiendoParte;
      const bytes = await edDoc.guardar();
      let salida, tipo, ext;
      if (modelo.formato === 'odt') { salida = (await OdtLector.aDocx(bytes)).bytes; tipo = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'; ext = '.docx'; }
      else { const { contrato } = await DocxAContrato.leer(bytes); salida = (await OdtEscritor.generar(contrato)).bytes; tipo = 'application/vnd.oasis.opendocument.text'; ext = '.odt'; }
      descargarBlob(new Blob([salida], { type: tipo }), nombreDoc + ext);
      avisar(tx('aviso-conversion'));
    } catch (e) { console.warn(e); avisar(tx('aviso-no-se-pudo') + ': ' + textoError(e)); }
  }
  const ESCALA_PDF = 3;
  async function exportarPdf() {
    if (!edDoc) return;
    avisar(tx('aviso-pdf'));
    try {
      await escribiendoParte;
      await editorDoc.relayout();
      const diseno = editorDoc.diseno();
      const pdf = await PDFLib.PDFDocument.create();
      const parrafos = new Set();
      for (const h of diseno.paginas) for (const grupo of [h.items, h.encabezado && h.encabezado.items, h.pie && h.pie.items, ...(h.notas || []).map((n) => n.items || [])]) for (const it of grupo || []) if (it.b && it.b.tipo === 'p') parrafos.add(it.b);
      const primera = diseno.paginas[0];
      await new LienzoPdf.Lienzo({ ancho: primera.ancho, alto: primera.alto, anchoPt: primera.ancho, altoPt: primera.alto }).precargarFuentes(mqDoc.motor.fuentesDe([...parrafos]));
      let comoImagen = 0;
      for (const hoja of diseno.paginas) {
        try {
          // la hoja se dibuja al triple de resolución: el texto que las fuentes incluidas no tienen (árabe, chino, emoji…) va como imagen y a 1 píxel por punto se ve borroso
          const lienzo = new LienzoPdf.Lienzo({ ancho: hoja.ancho * ESCALA_PDF, alto: hoja.alto * ESCALA_PDF, anchoPt: hoja.ancho, altoPt: hoja.alto });
          lienzo.scale(ESCALA_PDF, ESCALA_PDF);
          pintorDoc.pintarHoja(lienzo, hoja);
          if (lienzo.fallo) throw lienzo.fallo;
          await lienzo.aPagina(pdf);
        } catch (err) {
          console.warn('Hoja ' + hoja.numero + ' como imagen:', err);
          comoImagen++;
          const c = document.createElement('canvas');
          c.width = Math.round(hoja.ancho * 2); c.height = Math.round(hoja.alto * 2);
          const x = c.getContext('2d');
          x.setTransform(2, 0, 0, 2, 0, 0);
          pintorDoc.pintarHoja(x, hoja);
          const png = await new Promise((res) => c.toBlob(res, 'image/png'));
          pdf.addPage([hoja.ancho, hoja.alto]).drawImage(await pdf.embedPng(await png.arrayBuffer()), { x: 0, y: 0, width: hoja.ancho, height: hoja.alto });
        }
      }
      descargarBlob(new Blob([await pdf.save()], { type: 'application/pdf' }), nombreDoc + '.pdf');
      window.ultimoPdf = { hojas: diseno.paginas.length, comoImagen };
      avisar(diseno.paginas.length === 1 ? tx('aviso-pdf-listo-1') : tx('aviso-pdf-listo', diseno.paginas.length));
    } catch (e) { console.warn(e); avisar(tx('aviso-no-se-pudo') + ': ' + textoError(e)); }
  }

  // ───────── Encabezado y pie ─────────
  // Se editan como un documento aparte en un panel sobre las hojas; mientras tanto `ed` y `editor` son los del encabezado, así la cinta y los atajos
  // actúan sobre él. Cada cambio se escribe en el paquete y las hojas del documento se vuelven a maquetar al instante.
  // mientras se abre uno no se abre otro (un doble clic en el botón crearía dos partes y dos paneles)
  let abriendoParte = false;
  async function abrirParte(tipo, cual, seccion) {
    if (!editorDoc || parte || abriendoParte) return;
    abriendoParte = true;
    try { await abrirParteYa(tipo, cual, seccion); } finally { abriendoParte = false; }
  }
  async function abrirParteYa(tipo, cual, seccion) {
    const pos = editorDoc.seleccion().foco;
    if (seccion === undefined) seccion = edDoc.seccionDe(pos);
    if (cual === undefined) { const c = editorDoc.cajaDe(pos), hoja = c && editorDoc.diseno().paginas[c.n]; cual = (hoja && hoja.tipo) || 'default'; }
    let r = null;
    try { r = await EncabezadoDoc.abrir(modelo, seccion, tipo, cual, { crear: true }); } catch (e) { avisar(tx('aviso-no-se-pudo') + ': ' + (e && e.message ? e.message : e)); return; }
    if (!r) { avisar(tx('aviso-una-seccion', tx(tipo === 'pie' ? 'p-pie' : 'p-encabezado').toLowerCase())); return; }
    const vista = r.vista;
    vista.secciones[0].alto = 600;
    const mq = MaquetadorDoc.crear(vista, document.createElement('canvas').getContext('2d'), {});
    const edP = EdicionDoc.crear(vista, { invalidar: mq.invalidar, alCambiar: alCambiarDoc });
    const diseno = await mq.maquetar();
    const pintor = PintorDoc.crear(vista);
    await pintor.precargar(diseno);
    panelParte.hidden = false;
    const visorP = VisorDoc.crear(contParte, { diseno, pintor, zoom: 1 });
    parteTitulo.textContent = tx(tipo === 'pie' ? 'p-pie' : 'p-encabezado') + (cual === 'first' ? tx('p-primera') : cual === 'even' ? tx('p-pares') : '') + (r.heredada ? tx('p-heredado') : '');
    const editorP = EditorDoc.crear({ contenedor: contParte, visor: visorP, ed: edP, mq, pintor, diseno, alCambiar: alCambiarDoc, alSeleccionar, alRelayout: ajustarAltoParte, etiqueta: parteTitulo.textContent });
    parte = { vista, ed: edP, editor: editorP, visor: visorP, tipo };
    ed = edP; editor = editorP;
    editorP.soloLectura(lectura);
    if (corrector.activo && corrector.c) editorP.ponerCorrector(corrector.c, corrector.idioma);
    await editorP.relayout();
    if (!lectura) editorP.enfocar();
    refrescarTodo();
  }
  function ajustarAltoParte(diseno) {
    if (!parte) return;
    let fin = 0;
    for (const h of diseno.paginas) for (const it of h.items) fin = Math.max(fin, it.t === 'linea' ? it.y + it.l.alto : (it.y || 0) + (it.alto || 0));
    const alto = Math.min(600, Math.max(40, Math.ceil(fin) + 16)), s = parte.vista.secciones[0];
    if (Math.abs(s.alto - alto) > 1) { s.alto = alto; parte.editor.relayout(); }
  }
  function alCambiarParte() {
    const p = parte;
    if (!p) return;
    partesCambiadas = true;
    // si una escritura falla, la cadena sigue viva: una promesa rechazada dejaría sin poder descargar, convertir ni cerrar el encabezado
    escribiendoParte = escribiendoParte.then(async () => { await EncabezadoDoc.guardar(p.vista, p.ed); if (editorDoc) await editorDoc.relayout(); }).catch((e) => { console.warn(e); avisar(tx('aviso-no-se-pudo') + ': ' + (e && e.message ? e.message : e)); });
    actualizarGuardado(); comandos.refrescar();
    autoguardado.programar(datosAutoguardado, 2500, (fecha) => { if (fecha) actualizarGuardado(fecha); });
  }
  async function cerrarParte(sinFoco) {
    if (!parte) return;
    const p = parte;
    await escribiendoParte;
    parte = null;
    p.editor.destruir(); p.visor.destruir();
    panelParte.hidden = true;
    ed = edDoc; editor = editorDoc;
    if (!sinFoco) { await editor.relayout(); if (!lectura) editor.enfocar(); }
    refrescarTodo();
  }
  // doble clic en el margen de arriba o de abajo de una hoja, como en Word
  cont.addEventListener('dblclick', (ev) => {
    if (parte || !visor || !editorDoc || lectura) return;
    const hoja = ev.target.closest && ev.target.closest('.vd-hoja');
    if (!hoja) return;
    const n = visor.hojas.indexOf(hoja), h = editorDoc.diseno().paginas[n];
    if (!h) return;
    const y = (ev.clientY - hoja.getBoundingClientRect().top) / visor.escala;
    const arriba = Math.max(h.seccion.arriba, h.hEnc ? h.seccion.encabezado + h.hEnc : 0), abajo = Math.max(h.seccion.abajo, h.hPie ? h.seccion.pie + h.hPie : 0);
    if (y < arriba) abrirParte('encabezado', h.tipo, h.secIdx); else if (y > h.alto - abajo) abrirParte('pie', h.tipo, h.secIdx);
  });
  const botonParte = (id, texto, fn) => { const b = crearEl('button', { type: 'button', id, 'data-i18n': texto }, [tx(texto)]); b.addEventListener('click', fn); parteBarra.append(b); return b; };
  botonParte('parteNumero', 'cmd-num-pagina', () => { if (parte && !lectura) { parte.editor.insertarCampo('PAGE'); parte.editor.enfocar(); } });
  botonParte('parteTotal', 'cmd-total-paginas', () => { if (parte && !lectura) { parte.editor.insertarCampo('NUMPAGES'); parte.editor.enfocar(); } });
  botonParte('parteCerrar', 'cmd-cerrar-parte', () => cerrarParte());

  // ───────── Ortografía ─────────
  const GUARDADO = 'mc-doc-ortografia';
  const guardadoLocal = (k, v) => { try { if (v === undefined) return JSON.parse(localStorage.getItem(GUARDADO + k)); localStorage.setItem(GUARDADO + k, JSON.stringify(v)); } catch (e) { /* sin almacenamiento */ } return null; };
  const corrector = { c: null, activo: guardadoLocal('Activo') === true, idioma: guardadoLocal('Idioma') || (idioma() === 'en' ? 'en' : 'es') };
  const leerDiccionario = async (id) => {
    const [aff, dic] = await Promise.all(['aff', 'dic'].map(async (x) => {
      const r = await fetch(rutaVendor('dictionary-' + id + '/4.0.0/index.' + x));
      if (!r.ok) throw new Error('HTTP ' + r.status);
      return r.text();
    }));
    return { aff, dic };
  };
  function aplicarCorrector() {
    const editores = () => [editorDoc, parte && parte.editor].filter(Boolean);
    if (!corrector.activo) { for (const e of editores()) e.ponerCorrector(null); return; }
    if (!corrector.c) corrector.c = CorrectorDoc.crear({ NSpell, leer: leerDiccionario, personales: guardadoLocal('Personales') || [], alAgregar: (l) => guardadoLocal('Personales', l) });
    for (const e of editores()) e.ponerCorrector(corrector.c, corrector.idioma);
    const idiomaPedido = corrector.idioma;
    corrector.c.cargar(idiomaPedido).then(() => { if (corrector.activo && corrector.idioma === idiomaPedido) for (const e of editores()) e.ponerCorrector(corrector.c, idiomaPedido); }, (e) => avisar(tx('aviso-orto') + ': ' + (e && e.message ? e.message : e)));
  }
  const menuOrto = $('menuOrto');
  const cerrarMenuOrto = () => { menuOrto.hidden = true; menuOrto.textContent = ''; };
  document.addEventListener('pointerdown', (ev) => { if (!menuOrto.hidden && !ev.target.closest('#menuOrto')) cerrarMenuOrto(); });
  document.addEventListener('keydown', (ev) => { if (ev.key === 'Escape' && !menuOrto.hidden) cerrarMenuOrto(); });
  // clic derecho (o pulsación larga, que el celular manda como «contextmenu») sobre una palabra marcada: sugerencias, omitir y agregar
  document.addEventListener('contextmenu', (ev) => {
    const donde = ev.target.closest && ev.target.closest('.vd-hoja');
    if (!donde || !corrector.c || !corrector.activo || lectura) return;
    const e2 = parte && parte.visor.hojas.includes(donde) ? parte.editor : editorDoc;
    if (!e2) return;
    const e = e2.errorEn(e2.posDePantalla(ev.clientX, ev.clientY));
    if (!e) return;
    ev.preventDefault();
    cerrarMenuOrto();
    const boton = (texto, accion) => {
      const b = crearEl('button', { type: 'button', role: 'menuitem' }, [texto]);
      b.addEventListener('click', () => { cerrarMenuOrto(); accion(); e2.enfocar(); });
      menuOrto.append(b);
    };
    const sug = corrector.c.sugerencias(e.palabra, e.idioma);
    if (sug.length) for (const s of sug) boton(s, () => e2.reemplazarPalabra(e, s));
    else menuOrto.append(crearEl('div', { style: 'padding:6px 10px;opacity:.7;font-size:.85rem' }, [tx('o-sin-sugerencias')]));
    menuOrto.append(crearEl('hr'));
    boton(tx('o-omitir'), () => { corrector.c.ignorar(e.palabra); aplicarCorrector(); });
    boton(tx('o-agregar'), () => { corrector.c.agregar(e.palabra); aplicarCorrector(); });
    menuOrto.hidden = false;
    menuOrto.style.left = Math.max(4, Math.min(ev.clientX, innerWidth - 200)) + 'px';
    menuOrto.style.top = Math.max(4, Math.min(ev.clientY, innerHeight - menuOrto.offsetHeight - 8)) + 'px';
    menuOrto.firstChild.focus();
  });

  // ───────── Buscar y reemplazar ─────────
  const opcionesBusqueda = () => ({ mayusculas: bMayus.checked, palabra: bPalabra.checked });
  function cuenta() {
    const n = editor ? editor.coincidencias() : 0;
    bCuenta.textContent = !bTexto.value ? '' : n === 0 ? tx('b-ninguna') : n === 1 ? tx('b-una') : tx('b-varias', n);
  }
  function buscarAhora(saltar) {
    if (!editor) return;
    editor.buscar(bTexto.value, opcionesBusqueda());
    if (saltar) editor.irA(1, true);
    cuenta();
  }
  function abrirBuscar() {
    if (!editor) return;
    barraBuscar.hidden = false;
    const s = editor.seleccion();
    if (s.ancla && s.foco) {
      const [a, f] = ed.ordenar(s.ancla, s.foco);
      if (a.b === f.b && f.o > a.o && f.o - a.o < 100) bTexto.value = a.b.texto.slice(a.o, f.o);
    }
    bTexto.focus(); bTexto.select();
    buscarAhora(false);
  }
  function cerrarBuscar() { barraBuscar.hidden = true; for (const e of [editorDoc, parte && parte.editor]) if (e) e.buscar('', {}); if (editor && !lectura) editor.enfocar(); }
  barraBuscar.addEventListener('mousedown', (ev) => { if (!ev.target.closest('input, label')) ev.preventDefault(); });
  bTexto.addEventListener('input', () => buscarAhora(true));
  bMayus.addEventListener('change', () => buscarAhora(true));
  bPalabra.addEventListener('change', () => buscarAhora(true));
  bTexto.addEventListener('keydown', (ev) => { if (ev.key === 'Enter') { ev.preventDefault(); if (editor) editor.irA(ev.shiftKey ? -1 : 1); } });
  bSiguiente.addEventListener('click', () => { if (editor) editor.irA(1); });
  bAnterior.addEventListener('click', () => { if (editor) editor.irA(-1); });
  bReemplazar.addEventListener('click', () => { if (editable()) { editor.reemplazarActual(bNuevo.value); cuenta(); } });
  bTodo.addEventListener('click', () => {
    if (!editable()) return;
    const n = editor.reemplazarTodo(bNuevo.value);
    bCuenta.textContent = n === 0 ? tx('b-ninguna') : n === 1 ? tx('b-un-reemplazo') : tx('b-reemplazos', n);
  });
  bCerrar.addEventListener('click', cerrarBuscar);
  barraBuscar.addEventListener('keydown', (ev) => { if (ev.key === 'Escape') { ev.preventDefault(); cerrarBuscar(); } });
  document.addEventListener('keydown', (ev) => { if (ev.key === 'F3' && editor && !barraBuscar.hidden) { ev.preventDefault(); editor.irA(ev.shiftKey ? -1 : 1); } });

  // ───────── Imagen ─────────
  const entradaImagen = crearEl('input', { type: 'file', accept: 'image/*', hidden: true, id: 'entradaImagen' });
  document.body.append(entradaImagen);
  entradaImagen.addEventListener('change', async () => {
    const archivo = entradaImagen.files[0];
    entradaImagen.value = '';
    if (!archivo || !editable()) return;
    if (parte) { avisar(tx('aviso-salto-parte')); return; }
    if (archivo.size > 20 * 1024 * 1024) { avisar(tx('aviso-imagen-grande')); return; }
    try {
      const bmp = await createImageBitmap(archivo);
      let ext = (archivo.name.split('.').pop() || '').toLowerCase(), bytes = new Uint8Array(await archivo.arrayBuffer());
      if (!['png', 'jpg', 'jpeg', 'gif', 'bmp'].includes(ext)) {
        // Word no abre webp ni svg: se convierte a PNG
        const lienzo = document.createElement('canvas');
        lienzo.width = bmp.width; lienzo.height = bmp.height;
        lienzo.getContext('2d').drawImage(bmp, 0, 0);
        bytes = new Uint8Array(await (await new Promise((res) => lienzo.toBlob(res, 'image/png'))).arrayBuffer());
        ext = 'png';
      }
      const sec = modelo.secciones[0], util = sec.ancho - sec.izquierda - sec.derecha;
      let ancho = bmp.width * 0.75, alto = bmp.height * 0.75;
      const k = Math.min(1, util / ancho, 600 / alto);
      ancho *= k; alto *= k;
      const r = await editor.insertarImagen(bytes, ext, ancho, alto, archivo.name.replace(/\.[^.]+$/, ''));
      if (!r) avisar(tx('aviso-imagen'));
    } catch (e) { avisar(tx('aviso-imagen') + ' ' + (e && e.message ? e.message : '')); }
    editor.enfocar();
  });

  // ───────── Comandos ─────────
  const reg = (c) => comandos.registrar(c);
  const hay = () => !!edDoc;
  const conFoco = (fn) => (...a) => { const r = fn(...a); if (editor && !lectura) editor.enfocar(); return r; };
  const RESALTES = { yellow: [255, 255, 0], green: [0, 255, 0], cyan: [0, 255, 255], magenta: [255, 0, 255], blue: [0, 0, 255], red: [255, 0, 0], darkBlue: [0, 0, 139], darkCyan: [0, 139, 139], darkGreen: [0, 100, 0], darkMagenta: [139, 0, 139], darkRed: [139, 0, 0], darkYellow: [128, 128, 0], darkGray: [128, 128, 128], lightGray: [211, 211, 211], black: [0, 0, 0], white: [255, 255, 255] };
  function resalteMasCercano(hex) {
    const v = parseInt(String(hex).replace('#', ''), 16), c = [v >> 16, (v >> 8) & 255, v & 255];
    let mejor = 'yellow', d = Infinity;
    for (const [n, r] of Object.entries(RESALTES)) { const dd = (r[0] - c[0]) ** 2 + (r[1] - c[1]) ** 2 + (r[2] - c[2]) ** 2; if (dd < d) { d = dd; mejor = n; } }
    return mejor;
  }
  const alternar = (prop) => conFoco(() => editor.alternar(prop));
  reg({ id: 'nuevo', texto: 'cmd-nuevo', ejecutar: () => { if (sucio() && !confirm(tx('salir-sin-descargar'))) return; documentoNuevo(); } });
  reg({ id: 'abrir', texto: 'cmd-abrir', atajo: 'mod+o', enTexto: true, ejecutar: () => $('entradaArchivo').click() });
  reg({ id: 'descargar', texto: 'cmd-descargar', atajo: 'mod+s', enTexto: true, ejecutar: descargar, habilitado: hay });
  reg({ id: 'pdf', texto: 'cmd-pdf', ejecutar: exportarPdf, habilitado: hay });
  reg({ id: 'convertirOdt', texto: 'cmd-convertir-odt', ejecutar: convertir, habilitado: () => hay() && modelo.formato !== 'odt' });
  reg({ id: 'convertirDocx', texto: 'cmd-convertir-docx', ejecutar: convertir, habilitado: () => hay() && modelo.formato === 'odt' });
  reg({ id: 'deshacer', texto: 'cmd-deshacer', atajo: 'mod+z', ejecutar: () => editor.deshacerRehacer(true), habilitado: () => editable() && ed.puedeDeshacer() });
  reg({ id: 'rehacer', texto: 'cmd-rehacer', atajo: 'mod+y', ejecutar: () => editor.deshacerRehacer(false), habilitado: () => editable() && ed.puedeRehacer() });
  reg({ id: 'modoEdicion', texto: 'cmd-modo-editar', habilitado: hay, ejecutar: () => { lectura = !lectura; aplicarModo(); refrescarTodo(); if (!lectura) editor.enfocar(); } });
  // estilo, fuente y tamaño: listas que reflejan lo seleccionado
  const estiloActual = () => { const e = estadoActual().estilo || ''; const m = /^(?:Heading|Ttulo)([1-3])$/.exec(e); return m ? 'Heading' + m[1] : e === 'Title' || e === 'Subtitle' ? e : 'Normal'; };
  reg({ id: 'estilo', texto: 'cmd-estilo', ejecutar: conFoco(async (v) => { if (v !== undefined) { await editor.ponerEstilo(v === 'Normal' ? null : v); cacheEstado = null; comandos.refrescar(); } }), habilitado: editable, valor: estiloActual });
  reg({ id: 'fuente', texto: 'cmd-fuente', ejecutar: conFoco((v) => { if (v) editor.formato({ fuente: v }); }), habilitado: editable, valor: () => estadoActual().fuente || '' });
  reg({ id: 'tamano', texto: 'cmd-tamano', ejecutar: conFoco((v) => { if (v) editor.formato({ tam: parseFloat(v) }); }), habilitado: editable, valor: () => (estadoActual().tam ? String(estadoActual().tam) : '') });
  reg({ id: 'negrita', texto: 'cmd-negrita', ejecutar: alternar('b'), habilitado: editable, activo: () => estadoActual().b === true });
  reg({ id: 'cursiva', texto: 'cmd-cursiva', ejecutar: alternar('i'), habilitado: editable, activo: () => estadoActual().i === true });
  reg({ id: 'subrayado', texto: 'cmd-subrayado', ejecutar: alternar('u'), habilitado: editable, activo: () => estadoActual().u === true });
  reg({ id: 'tachado', texto: 'cmd-tachado', ejecutar: alternar('tachado'), habilitado: editable, activo: () => estadoActual().tachado === true });
  reg({ id: 'colorTexto', texto: 'cmd-color-texto', ejecutar: (hex) => { if (hex) editor.formato({ color: String(hex).replace('#', '') }); }, habilitado: editable });
  reg({ id: 'resaltado', texto: 'cmd-resaltado', ejecutar: (hex) => { if (hex) editor.formato({ resaltado: estadoActual().resaltado === resalteMasCercano(hex) ? 'none' : resalteMasCercano(hex) }); }, habilitado: editable });
  reg({ id: 'superindice', texto: 'cmd-superindice', ejecutar: conFoco(() => editor.formato({ vertical: estadoActual().vertical === 'superscript' ? 'baseline' : 'superscript' })), habilitado: editable, activo: () => estadoActual().vertical === 'superscript' });
  reg({ id: 'subindice', texto: 'cmd-subindice', ejecutar: conFoco(() => editor.formato({ vertical: estadoActual().vertical === 'subscript' ? 'baseline' : 'subscript' })), habilitado: editable, activo: () => estadoActual().vertical === 'subscript' });
  reg({ id: 'limpiar', texto: 'cmd-limpiar', ejecutar: conFoco(() => editor.formato({ limpiar: true })), habilitado: editable });
  for (const [id, texto, valor] of [['izquierda', 'cmd-izquierda', 'left'], ['centro', 'cmd-centro', 'center'], ['derecha', 'cmd-derecha', 'right'], ['justificar', 'cmd-justificar', 'both']]) {
    reg({ id, texto, ejecutar: conFoco(() => editor.alinear(valor)), habilitado: editable, activo: () => estadoActual().jc === valor });
  }
  reg({ id: 'interlineado', texto: 'cmd-interlineado', ejecutar: conFoco((v) => { if (v) editor.interlineado(parseFloat(v)); }), habilitado: editable,
    valor: () => { const i = estadoActual().interlineado; return i === null || i === undefined ? '' : String(Math.round(i * 100) / 100); } });
  reg({ id: 'vinetas', texto: 'cmd-vinetas', ejecutar: conFoco(() => editor.ponerLista('vinetas')), habilitado: editable, activo: () => estadoActual().lista === 'vinetas' });
  reg({ id: 'numeracion', texto: 'cmd-numeracion', ejecutar: conFoco(() => editor.ponerLista('numeros')), habilitado: editable, activo: () => estadoActual().lista === 'numeros' });
  reg({ id: 'sangriaMas', texto: 'cmd-sangria-mas', ejecutar: conFoco(() => editor.sangrar(1)), habilitado: editable });
  reg({ id: 'sangriaMenos', texto: 'cmd-sangria-menos', ejecutar: conFoco(() => editor.sangrar(-1)), habilitado: editable });
  // insertar
  reg({ id: 'tabla', texto: 'cmd-tabla', habilitado: editable, ejecutar: async () => {
    const v = await Dialogo.formulario({ titulo: tx('tabla-titulo'), campos: [{ nombre: 'filas', etiqueta: tx('tabla-filas'), tipo: 'numero', valor: 3 }, { nombre: 'columnas', etiqueta: tx('tabla-columnas'), tipo: 'numero', valor: 3 }], aceptar: tx('aceptar'), cancelar: tx('cancelar') });
    if (!v) { editor.enfocar(); return; }
    const r = editor.insertarTabla(Math.min(50, Math.max(1, parseInt(v.filas, 10) || 3)), Math.min(20, Math.max(1, parseInt(v.columnas, 10) || 3)));
    if (!r) avisar(tx('aviso-tabla'));
    editor.enfocar();
  } });
  reg({ id: 'imagen', texto: 'cmd-imagen', ejecutar: () => { if (parte) avisar(tx('aviso-salto-parte')); else entradaImagen.click(); }, habilitado: editable });
  reg({ id: 'saltoPagina', texto: 'cmd-salto-pagina', ejecutar: conFoco(() => { if (parte) avisar(tx('aviso-salto-parte')); else editor.saltoDePagina(); }), habilitado: editable });
  const saltoSeccion = (tipo) => conFoco(() => { if (parte) { avisar(tx('aviso-salto-parte')); return; } if (!editor.saltoDeSeccion(tipo)) avisar(tx('aviso-secciones')); });
  const conSecciones = () => editable() && (!ed || ed.puede('secciones'));
  reg({ id: 'saltoSeccion', texto: 'cmd-salto-seccion', ejecutar: saltoSeccion('nextPage'), habilitado: conSecciones });
  reg({ id: 'saltoSeccionContinuo', texto: 'cmd-salto-seccion-continuo', ejecutar: saltoSeccion('continuous'), habilitado: conSecciones });
  reg({ id: 'encabezado', texto: 'cmd-encabezado', habilitado: hay, activo: () => !!parte && parte.tipo === 'encabezado', ejecutar: async () => { const mismo = parte && parte.tipo === 'encabezado'; if (parte) await cerrarParte(); if (!mismo) await abrirParte('encabezado'); } });
  reg({ id: 'pie', texto: 'cmd-pie', habilitado: hay, activo: () => !!parte && parte.tipo === 'pie', ejecutar: async () => { const mismo = parte && parte.tipo === 'pie'; if (parte) await cerrarParte(); if (!mismo) await abrirParte('pie'); } });
  reg({ id: 'numPagina', texto: 'cmd-num-pagina', ejecutar: conFoco(() => editor.insertarCampo('PAGE')), habilitado: () => !!parte && !lectura });
  reg({ id: 'totalPaginas', texto: 'cmd-total-paginas', ejecutar: conFoco(() => editor.insertarCampo('NUMPAGES')), habilitado: () => !!parte && !lectura });
  // tabla
  const enTabla = () => editable() && !!editor.enTabla();
  const ordenTabla = (nombre) => conFoco(() => { if (!editor.ordenTabla(nombre)) avisar(tx('aviso-estructura')); });
  for (const [id, texto, nombre] of [['filaAntes', 'cmd-fila-antes', 'filaAntes'], ['filaDespues', 'cmd-fila-despues', 'filaDespues'], ['colAntes', 'cmd-col-antes', 'columnaAntes'], ['colDespues', 'cmd-col-despues', 'columnaDespues'], ['quitarFila', 'cmd-quitar-fila', 'quitarFila'], ['quitarCol', 'cmd-quitar-col', 'quitarColumna'], ['quitarTabla', 'cmd-quitar-tabla', 'quitarTabla']]) {
    reg({ id, texto, ejecutar: ordenTabla(nombre), habilitado: enTabla });
  }
  // revisar
  reg({ id: 'buscar', texto: 'cmd-buscar', atajo: 'mod+f', enTexto: true, ejecutar: abrirBuscar, habilitado: hay });
  reg({ id: 'ortografia', texto: 'cmd-ortografia', habilitado: hay, activo: () => corrector.activo, ejecutar: () => { corrector.activo = !corrector.activo; guardadoLocal('Activo', corrector.activo); aplicarCorrector(); } });
  reg({ id: 'idiomaOrto', texto: 'cmd-idioma-orto', habilitado: hay, valor: () => corrector.idioma, ejecutar: (v) => { if (v === 'es' || v === 'en') { corrector.idioma = v; guardadoLocal('Idioma', v); if (corrector.activo) aplicarCorrector(); } } });
  // vista
  reg({ id: 'masZoom', texto: 'cmd-mas-zoom', atajo: 'mod+=', ejecutar: () => cambiarZoom(1.25), habilitado: hay });
  reg({ id: 'menosZoom', texto: 'cmd-menos-zoom', atajo: 'mod+-', ejecutar: () => cambiarZoom(0.8), habilitado: hay });
  reg({ id: 'ajustar', texto: 'cmd-ajustar', ejecutar: () => ajustarAlAncho(true), habilitado: hay });
  reg({ id: 'inspector', texto: 'cmd-inspector', ejecutar: () => insp.alternar(), activo: () => insp.visible(), habilitado: () => true });
  reg({ id: 'paleta', texto: 'cmd-paleta', atajo: 'mod+k', enTexto: true, ejecutar: () => Oficina.paleta(comandos), oculto: true, habilitado: () => true });
  reg({ id: 'completa', texto: 'cmd-completa', atajo: 'f11', ejecutar: () => { if (document.fullscreenElement) document.exitFullscreen(); else $('app').requestFullscreen && $('app').requestFullscreen(); }, habilitado: () => true });

  // ───────── La cinta ─────────
  const FUENTES = ['Arial', 'Calibri', 'Cambria', 'Comic Sans MS', 'Consolas', 'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'];
  const TAMANOS = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];
  Oficina.cinta($('pestanas'), $('cinta'), [
    { id: 'inicio', texto: 'tab-inicio', grupos: [
      { texto: 'g-estilo', items: [{ comando: 'estilo', tipo: 'select', opciones: [['Normal', 'est-normal'], ['Title', 'est-titulo'], ['Subtitle', 'est-subtitulo'], ['Heading1', 'est-t1'], ['Heading2', 'est-t2'], ['Heading3', 'est-t3']] }] },
      { texto: 'g-fuente', items: [
        { comando: 'fuente', tipo: 'select', opciones: [['', ''], ...FUENTES.map((f) => [f, f])] }, { comando: 'tamano', tipo: 'select', opciones: [['', ''], ...TAMANOS.map((t) => [String(t), String(t).replace('.', ',')])] },
        { comando: 'negrita', icono: 'N' }, { comando: 'cursiva', icono: 'C' }, { comando: 'subrayado', icono: 'S' }, { comando: 'tachado', icono: 'T̶' },
        { comando: 'colorTexto', tipo: 'color', icono: 'A', porDefecto: '#000000' }, { comando: 'resaltado', tipo: 'color', icono: '▇', porDefecto: '#ffff00' },
        { comando: 'superindice', icono: 'x²' }, { comando: 'subindice', icono: 'x₂' }, { comando: 'limpiar', icono: 'Tx' },
      ] },
      { texto: 'g-parrafo', items: [
        { comando: 'izquierda', icono: '⯇' }, { comando: 'centro', icono: '≡' }, { comando: 'derecha', icono: '⯈' }, { comando: 'justificar', icono: '☰' },
        { comando: 'interlineado', tipo: 'select', opciones: [['', ''], ['1', '1'], ['1.15', '1,15'], ['1.5', '1,5'], ['2', '2'], ['3', '3']] },
        { comando: 'vinetas', icono: '•≡' }, { comando: 'numeracion', icono: '1.' }, { comando: 'sangriaMenos', icono: '⇤' }, { comando: 'sangriaMas', icono: '⇥' },
      ] },
      { texto: 'g-archivo', solo: 'pc', items: [{ comando: 'nuevo', icono: '📄' }, { comando: 'abrir', icono: '📂' }, { comando: 'descargar', icono: '⤓' }, { comando: 'pdf', icono: '⎙' }] },
    ] },
    { id: 'insertar', texto: 'tab-insertar', grupos: [
      { texto: 'g-ilustraciones', items: [{ comando: 'tabla', icono: '▦', conTexto: true }, { comando: 'imagen', icono: '🖼', conTexto: true }] },
      { texto: 'g-tabla-cambiar', items: [{ comando: 'filaAntes', icono: '⤒' }, { comando: 'filaDespues', icono: '⤓' }, { comando: 'colAntes', icono: '⇤' }, { comando: 'colDespues', icono: '⇥' }, { comando: 'quitarFila', icono: '−▭' }, { comando: 'quitarCol', icono: '−▯' }, { comando: 'quitarTabla', icono: '🗑' }] },
      { texto: 'g-saltos', items: [{ comando: 'saltoPagina', icono: '⏎', conTexto: true }, { comando: 'saltoSeccion', icono: '§' }, { comando: 'saltoSeccionContinuo', icono: '§⋯' }] },
      { texto: 'g-encabezado', items: [{ comando: 'encabezado', icono: '⬒', conTexto: true }, { comando: 'pie', icono: '⬓', conTexto: true }, { comando: 'numPagina', icono: '#' }, { comando: 'totalPaginas', icono: '#/∑' }] },
    ] },
    { id: 'revisar', texto: 'tab-revisar', grupos: [
      { texto: 'g-buscar', items: [{ comando: 'buscar', icono: '🔍', conTexto: true }] },
      { texto: 'g-ortografia', items: [{ comando: 'ortografia', icono: 'abc', conTexto: true }, { comando: 'idiomaOrto', tipo: 'select', opciones: [['es', 'idioma-es'], ['en', 'idioma-en']] }] },
    ] },
    { id: 'vista', texto: 'tab-vista', grupos: [
      { texto: 'g-zoom', items: [{ comando: 'menosZoom', icono: '−' }, { comando: 'masZoom', icono: '+' }, { comando: 'ajustar', icono: '↔', conTexto: true }] },
      { texto: 'g-ventana', items: [{ comando: 'inspector', icono: '☰', conTexto: true }, { comando: 'completa', icono: '⛶', conTexto: true }, { comando: 'paleta', icono: '⌘', conTexto: true }] },
    ] },
  ], comandos);
  comandos.conectarBotones(document.querySelector('.of-acciones'));
  Oficina.menu($('botonMas'), [{ comando: 'nuevo' }, { comando: 'abrir' }, '-', { comando: 'descargar' }, { comando: 'convertirOdt' }, { comando: 'convertirDocx' }, { comando: 'pdf' }, '-', { comando: 'buscar' }, { comando: 'ortografia' }, '-', { comando: 'saltoPagina' }, { comando: 'encabezado' }, { comando: 'pie' }, '-', { comando: 'masZoom' }, { comando: 'menosZoom' }, { comando: 'ajustar' }, { comando: 'inspector' }, { comando: 'completa' }, { comando: 'paleta' }], comandos);
  comandos.conectarTeclado(document);
  // los botones de la cinta no le quitan el foco ni la selección al documento
  $('cinta').addEventListener('mousedown', (ev) => { if (editor && !lectura && ev.target.closest('button')) ev.preventDefault(); });
  $('nombreDoc').addEventListener('change', () => { nombreDoc = $('nombreDoc').value.trim().replace(/[\\/:*?"<>|]+/g, '-') || nombreDoc; $('nombreDoc').value = nombreDoc; });
  $('nombreDoc').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); e.target.blur(); } });
  window.addEventListener('resize', () => { if (visor && editorDoc && matchMedia('(max-width: 900px)').matches) ajustarAlAncho(false); });

  // ───────── Inicio: abrir, nuevo, borradores ─────────
  conectarZonaArrastre($('zonaArrastre'), $('entradaArchivo'), (archivos) => { if (archivos[0]) abrirArchivo(archivos[0]); });
  $('botonNuevo').addEventListener('click', documentoNuevo);
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
  window.addEventListener('beforeunload', (e) => { if (sucio()) { e.preventDefault(); e.returnValue = ''; } });
  document.addEventListener('dragover', (e) => e.preventDefault());
  document.addEventListener('drop', (e) => { e.preventDefault(); if (e.dataTransfer.files[0]) abrirArchivo(e.dataTransfer.files[0]); });

  // El idioma se cambia sin recargar la página (botón del inicio): lo que se escribió con tx() al cargar o al abrir se vuelve a escribir
  document.addEventListener('cambioIdioma', () => {
    for (const [el, k] of [[bTexto, 'b-buscar'], [bNuevo, 'b-reemplazar-con']]) { el.placeholder = tx(k); el.setAttribute('aria-label', tx(k)); }
    bMayus.parentNode.lastChild.textContent = ' ' + tx('b-mayus'); bPalabra.parentNode.lastChild.textContent = ' ' + tx('b-palabra');
    for (const [el, k] of [[bAnterior, 'b-anterior'], [bSiguiente, 'b-siguiente'], [bCerrar, 'b-cerrar'], [bReemplazar, 'b-reemplazar'], [bTodo, 'b-todo']]) el.setAttribute('aria-label', tx(k));
    bReemplazar.textContent = tx('b-reemplazar'); bTodo.textContent = tx('b-todo');
    pintarHonesta(); mostrarBorradores();
    if (editorDoc) editorDoc.area.setAttribute('aria-label', tx('nombre-app'));
    if (modelo) { aplicarModo(); actualizarGuardado(); cuenta(); refrescarTodo(); }
  });

  // para las pruebas de navegador
  window.DocumentosApp = {
    abrirBytes, comandos, abrirParte, cerrarParte, descargar, exportarPdf, aplicarModo, mostrarBorradores,
    get modelo() { return modelo; }, get ed() { return edDoc; }, get editor() { return editorDoc; }, get visor() { return visor; }, get parte() { return parte; },
    get actual() { return { ed, editor }; }, get lectura() { return lectura; }, get cosas() { return cosas; }, get interfaz() { return interfaz; }, get docId() { return docId; },
  };
})();

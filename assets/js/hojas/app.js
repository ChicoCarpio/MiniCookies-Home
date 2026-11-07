/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Mini Cookies Hojas (fase 2 del PLAN): la aplicación. Une el libro editable (libro-editable.js), la rejilla (motores/rejilla-hoja.js),
// el historial, los comandos y la interfaz G. Aquí viven el editor de celda, la barra de fórmulas, la cinta, el portapapeles y el autoguardado.
(function () {
  'use strict';
  // ---- textos ----------------------------------------------------------------------------------------------------
  const ES = {
    'titulo-pagina': 'Editar Excel y .ods en el navegador — Mini Cookies', 'meta-descripcion': 'Edita libros de Excel (.xlsx) y hojas de LibreOffice (.ods) en el navegador: fórmulas, formato, gráficos y deshacer, sin instalar nada y sin subir tu archivo. Gratis y sin registro.',
    'h1': 'Editar Excel (Hojas)', 'intro': 'Abre un libro de Excel y edítalo como en una hoja de cálculo: valores, fórmulas, formato, deshacer. Al guardar, lo que no tocaste queda exactamente como estaba.',
    'zona-arrastre': '📊 Arrastra tu archivo aquí o haz clic para elegirlo (.xlsx, .xlsm, .ods, .csv, .tsv, .txt, .xls, .xlsb)', 'boton-nuevo': 'Libro en blanco', 'cargando': 'Abriendo el libro…',
    'borradores': 'Trabajo sin terminar en este navegador', 'recuperar': 'Recuperar', 'descartar': 'Descartar', 'nombre-app': 'Hojas', 'volver-inicio': 'Volver al inicio',
    'nombre-archivo': 'Nombre del archivo', 'mas-opciones': 'Más opciones', 'cinta': 'Cinta', 'cuadro-nombre': 'Cuadro de nombre', 'contenido-celda': 'Contenido de la celda', 'inspector': 'Propiedades',
    'tab-inicio': 'Inicio', 'tab-insertar': 'Insertar', 'tab-datos': 'Datos', 'tab-vista': 'Vista',
    'g-historial': 'Deshacer', 'g-portapapeles': 'Portapapeles', 'g-fuente': 'Fuente', 'g-alineacion': 'Alineación', 'g-numero': 'Número', 'g-celdas': 'Celdas', 'g-edicion': 'Edición', 'g-funciones': 'Funciones', 'g-archivo': 'Archivo', 'g-zoom': 'Zoom', 'g-ventana': 'Ventana', 'g-buscar': 'Buscar',
    'cmd-deshacer': 'Deshacer', 'cmd-rehacer': 'Rehacer', 'cmd-cortar': 'Cortar', 'cmd-copiar': 'Copiar', 'cmd-pegar': 'Pegar', 'cmd-pegar-valores': 'Pegar valores', 'cmd-negrita': 'Negrita', 'cmd-cursiva': 'Cursiva', 'cmd-subrayado': 'Subrayado', 'cmd-tachado': 'Tachado',
    'cmd-color': 'Color de fuente', 'cmd-relleno': 'Color de relleno', 'cmd-sin-relleno': 'Sin relleno', 'cmd-izquierda': 'Alinear a la izquierda', 'cmd-centro': 'Centrar', 'cmd-derecha': 'Alinear a la derecha', 'cmd-ajustar': 'Ajustar texto',
    'cmd-formato': 'Formato de número', 'cmd-mas-decimales': 'Aumentar decimales', 'cmd-menos-decimales': 'Disminuir decimales', 'cmd-bordes': 'Todos los bordes', 'cmd-sin-bordes': 'Sin bordes',
    'cmd-borrar': 'Borrar contenido', 'cmd-borrar-formato': 'Borrar formato', 'cmd-autosuma': 'Autosuma', 'cmd-buscar': 'Buscar', 'cmd-descargar': 'Descargar el libro', 'cmd-descargar-corto': 'Descargar', 'cmd-csv': 'Exportar la hoja como .csv',
    'cmd-nuevo': 'Libro nuevo', 'cmd-abrir': 'Abrir otro archivo', 'cmd-paleta': 'Buscar un comando', 'cmd-zoom-mas': 'Acercar', 'cmd-zoom-menos': 'Alejar', 'cmd-completa': 'Pantalla completa', 'cmd-editar-celda': 'Editar la celda',
    'cmd-modo-editar': 'Editar', 'cmd-modo-leer': 'Leer', 'cmd-inspector': 'Panel de propiedades', 'cmd-expandir-barra': 'Expandir la barra de fórmulas', 'cmd-hoy': 'Fecha de hoy', 'cmd-funcion': 'Insertar función…',
    'fmt-general': 'General', 'fmt-numero': 'Número', 'fmt-moneda': 'Moneda', 'fmt-porcentaje': 'Porcentaje', 'fmt-fecha': 'Fecha corta', 'fmt-hora': 'Hora', 'fmt-texto': 'Texto', 'fmt-miles': 'Miles',
    'guardado-local': 'Guardado en este navegador a las {0}', 'sin-cambios': 'Sin cambios', 'cambios-sin-descargar': 'Cambios sin descargar', 'descargado': 'Descargado',
    'aviso-protegida': 'Esa celda es parte de una tabla o de una fórmula matricial: se edita en Excel', 'aviso-formula': 'La fórmula tiene un error: revísala (paréntesis, comillas, separadores)',
    'aviso-no-soportadas': '{0} fórmulas usan funciones que aún no se calculan aquí: conservan su valor y Excel las recalcula al abrir', 'aviso-circular': 'Hay una referencia circular: la fórmula da 0',
    'aviso-pegado-grande': 'Lo pegado es demasiado grande (más de {0} celdas)', 'aviso-copiado': 'Copiado', 'aviso-cortado': 'Cortado: se mueve al pegar. Esc cancela', 'aviso-mover-combinada': 'No se puede mover una parte de una celda combinada', 'aviso-mover-matriz': 'No se puede cambiar una parte de una matriz', 'aviso-mover-limite': 'No se puede mover: se saldría de la hoja', 'aviso-valores-cortado': 'Pegar valores no se puede usar con lo cortado', 'aviso-portapapeles': 'El navegador no deja leer el portapapeles: usa Ctrl+V',
    'honesta-conserva': 'Se conserva al guardar, sin editar aquí: {0}.', 'honesta-convertido': 'Este archivo venía en otro formato: se guarda como .xlsx.', 'honesta-todo': 'Todo el contenido del archivo se puede editar aquí.',
    'x-graficos': 'imágenes, formas o gráficos que no se editan', 'x-macros': 'macros', 'x-notas': 'notas', 'x-tablas': 'encabezados de tablas', 'x-dinamicas': 'tablas dinámicas', 'x-condicional': 'reglas de formato de Excel 2010', 'x-validacion': 'validación de datos', 'x-minigraficos': 'minigráficos', 'x-imagenes': 'imágenes en celdas', 'x-matriciales': 'fórmulas matriciales de Ctrl+Mayús+Intro',
    'i-celda': 'Celda', 'i-tipo': 'Tipo', 'i-valor': 'Valor', 'i-formula': 'Fórmula', 'i-formato': 'Formato', 'i-fuente': 'Fuente', 'i-archivo': 'El archivo', 'i-hojas': 'Hojas', 'i-formulas': 'Fórmulas',
    't-numero': 'Número', 't-texto': 'Texto', 't-logico': 'Lógico', 't-error': 'Error', 't-vacio': 'Vacía', 'error-abrir': 'No se pudo abrir el archivo', 'salir-sin-descargar': 'Hay cambios sin descargar.',
    'funcion-pregunta': 'Escribe el nombre de la función (por ejemplo SUMA, PROMEDIO, BUSCARV):',
    'cmd-insertar-columnas': 'Insertar columnas', 'cmd-quitar-columnas': 'Quitar columnas', 'g-filas': 'Filas y columnas', 'g-hojas': 'Hojas', 'g-paneles': 'Paneles', 'cmd-insertar-filas': 'Insertar filas', 'cmd-quitar-filas': 'Quitar filas', 'cmd-insertar-hoja': 'Hoja nueva', 'cmd-renombrar-hoja': 'Renombrar la hoja',
    'cmd-quitar-hoja': 'Quitar la hoja', 'cmd-hoja-izquierda': 'Mover la hoja a la izquierda', 'cmd-hoja-derecha': 'Mover la hoja a la derecha', 'cmd-combinar': 'Combinar celdas', 'cmd-separar': 'Separar celdas',
    'cmd-inmovilizar': 'Inmovilizar paneles', 'cmd-movilizar': 'Movilizar paneles', 'cmd-ancho-columna': 'Ancho de columna', 'pregunta-nombre-hoja': 'Nombre de la hoja:', 'pregunta-quitar-hoja': '¿Quitar la hoja «{0}» y todo lo que tiene?',
    'pregunta-ancho': 'Ancho de la columna (en caracteres):', 'aviso-estructura': 'No se pudo hacer', 'aviso-formulas-ods': '{0} fórmulas no se pudieron escribir en .ods: quedó solo su valor', 'x-formato-condicional': 'formato condicional', 'x-filtros': 'filtros automáticos', 'x-formulas-sin-traducir': 'fórmulas que aquí solo se ven como valor',
    'cmd-alto-fila': 'Alto de fila', 'pregunta-alto': 'Alto de la fila (en puntos, de 0 a 409; vacío = automático):',
    'cmd-reemplazar': 'Buscar y reemplazar', 'cmd-ordenar-asc': 'Ordenar de A a Z', 'cmd-ordenar-desc': 'Ordenar de Z a A', 'cmd-pdf': 'Descargar la hoja en PDF', 'cmd-pdf-libro': 'Descargar el libro en PDF', 'g-ordenar': 'Ordenar',
    'r-buscar': 'Buscar', 'r-reemplazar': 'Reemplazar con', 'r-mayus': 'Distinguir mayúsculas', 'r-entera': 'Celda completa', 'r-libro': 'En todo el libro', 'r-todo': 'Reemplazar todo', 'cerrar': 'Cerrar', 'r-hechos': '{0} celdas cambiadas',
    'aviso-pdf': 'Preparando el PDF…', 'aviso-pdf-listo': 'PDF listo: {0} páginas', 'aviso-ordenar-combinadas': 'No se puede ordenar: hay celdas combinadas en el rango', 'aviso-ordenar-matriz': 'No se puede ordenar: el rango incluye una matriz dinámica',
  };
  const EN = {
    'titulo-pagina': 'Edit Excel and .ods files in your browser — Mini Cookies', 'meta-descripcion': 'Edit Excel workbooks (.xlsx) and LibreOffice spreadsheets (.ods) in your browser: formulas, formatting, charts and undo, without installing anything or uploading your file. Free, no sign-up.',
    'h1': 'Edit Excel (Sheets)', 'intro': 'Open an Excel workbook and edit it like a spreadsheet: values, formulas, formatting, undo. When you save, everything you didn’t touch stays exactly as it was.',
    'zona-arrastre': '📊 Drop your file here or click to choose it (.xlsx, .xlsm, .ods, .csv, .tsv, .txt, .xls, .xlsb)', 'boton-nuevo': 'Blank workbook', 'cargando': 'Opening the workbook…',
    'borradores': 'Unfinished work in this browser', 'recuperar': 'Recover', 'descartar': 'Discard', 'nombre-app': 'Sheets', 'volver-inicio': 'Back to start',
    'nombre-archivo': 'File name', 'mas-opciones': 'More options', 'cinta': 'Ribbon', 'cuadro-nombre': 'Name box', 'contenido-celda': 'Cell contents', 'inspector': 'Properties',
    'tab-inicio': 'Home', 'tab-insertar': 'Insert', 'tab-datos': 'Data', 'tab-vista': 'View',
    'g-historial': 'Undo', 'g-portapapeles': 'Clipboard', 'g-fuente': 'Font', 'g-alineacion': 'Alignment', 'g-numero': 'Number', 'g-celdas': 'Cells', 'g-edicion': 'Editing', 'g-funciones': 'Functions', 'g-archivo': 'File', 'g-zoom': 'Zoom', 'g-ventana': 'Window', 'g-buscar': 'Find',
    'cmd-deshacer': 'Undo', 'cmd-rehacer': 'Redo', 'cmd-cortar': 'Cut', 'cmd-copiar': 'Copy', 'cmd-pegar': 'Paste', 'cmd-pegar-valores': 'Paste values', 'cmd-negrita': 'Bold', 'cmd-cursiva': 'Italic', 'cmd-subrayado': 'Underline', 'cmd-tachado': 'Strikethrough',
    'cmd-color': 'Font color', 'cmd-relleno': 'Fill color', 'cmd-sin-relleno': 'No fill', 'cmd-izquierda': 'Align left', 'cmd-centro': 'Center', 'cmd-derecha': 'Align right', 'cmd-ajustar': 'Wrap text',
    'cmd-formato': 'Number format', 'cmd-mas-decimales': 'Increase decimals', 'cmd-menos-decimales': 'Decrease decimals', 'cmd-bordes': 'All borders', 'cmd-sin-bordes': 'No borders',
    'cmd-borrar': 'Clear contents', 'cmd-borrar-formato': 'Clear formats', 'cmd-autosuma': 'AutoSum', 'cmd-buscar': 'Find', 'cmd-descargar': 'Download the workbook', 'cmd-descargar-corto': 'Download', 'cmd-csv': 'Export sheet as .csv',
    'cmd-nuevo': 'New workbook', 'cmd-abrir': 'Open another file', 'cmd-paleta': 'Search for a command', 'cmd-zoom-mas': 'Zoom in', 'cmd-zoom-menos': 'Zoom out', 'cmd-completa': 'Full screen', 'cmd-editar-celda': 'Edit cell',
    'cmd-modo-editar': 'Edit', 'cmd-modo-leer': 'Read', 'cmd-inspector': 'Properties panel', 'cmd-expandir-barra': 'Expand the formula bar', 'cmd-hoy': 'Today’s date', 'cmd-funcion': 'Insert function…',
    'fmt-general': 'General', 'fmt-numero': 'Number', 'fmt-moneda': 'Currency', 'fmt-porcentaje': 'Percent', 'fmt-fecha': 'Short date', 'fmt-hora': 'Time', 'fmt-texto': 'Text', 'fmt-miles': 'Thousands',
    'guardado-local': 'Saved in this browser at {0}', 'sin-cambios': 'No changes', 'cambios-sin-descargar': 'Changes not downloaded', 'descargado': 'Downloaded',
    'aviso-protegida': 'That cell belongs to a table or an array formula: edit it in Excel', 'aviso-formula': 'The formula has an error: check parentheses, quotes and separators',
    'aviso-no-soportadas': '{0} formulas use functions not calculated here yet: they keep their value and Excel recalculates them on open', 'aviso-circular': 'There is a circular reference: the formula returns 0',
    'aviso-pegado-grande': 'What you pasted is too large (more than {0} cells)', 'aviso-copiado': 'Copied', 'aviso-cortado': 'Cut: it moves when you paste. Esc cancels', 'aviso-mover-combinada': 'Can’t move part of a merged cell', 'aviso-mover-matriz': 'You can’t change part of an array', 'aviso-mover-limite': 'Can’t move: it would go off the sheet', 'aviso-valores-cortado': 'Paste values can’t be used with what you cut', 'aviso-portapapeles': 'The browser won’t let us read the clipboard: use Ctrl+V',
    'honesta-conserva': 'Kept as is when saving, not editable here: {0}.', 'honesta-convertido': 'This file came in another format: it is saved as .xlsx.', 'honesta-todo': 'Everything in this file can be edited here.',
    'x-graficos': 'images, shapes or charts that can’t be edited', 'x-macros': 'macros', 'x-notas': 'notes', 'x-tablas': 'table headers', 'x-dinamicas': 'pivot tables', 'x-condicional': 'Excel 2010 formatting rules', 'x-validacion': 'data validation', 'x-minigraficos': 'sparklines', 'x-imagenes': 'images in cells', 'x-matriciales': 'Ctrl+Shift+Enter array formulas',
    'i-celda': 'Cell', 'i-tipo': 'Type', 'i-valor': 'Value', 'i-formula': 'Formula', 'i-formato': 'Format', 'i-fuente': 'Font', 'i-archivo': 'The file', 'i-hojas': 'Sheets', 'i-formulas': 'Formulas',
    't-numero': 'Number', 't-texto': 'Text', 't-logico': 'Logical', 't-error': 'Error', 't-vacio': 'Empty', 'error-abrir': 'The file could not be opened', 'salir-sin-descargar': 'There are changes not downloaded.',
    'funcion-pregunta': 'Type the function name (for example SUM, AVERAGE, VLOOKUP):',
    'cmd-insertar-columnas': 'Insert columns', 'cmd-quitar-columnas': 'Delete columns', 'g-filas': 'Rows & columns', 'g-hojas': 'Sheets', 'g-paneles': 'Panes', 'cmd-insertar-filas': 'Insert rows', 'cmd-quitar-filas': 'Delete rows', 'cmd-insertar-hoja': 'New sheet', 'cmd-renombrar-hoja': 'Rename sheet',
    'cmd-quitar-hoja': 'Delete sheet', 'cmd-hoja-izquierda': 'Move sheet left', 'cmd-hoja-derecha': 'Move sheet right', 'cmd-combinar': 'Merge cells', 'cmd-separar': 'Unmerge cells',
    'cmd-inmovilizar': 'Freeze panes', 'cmd-movilizar': 'Unfreeze panes', 'cmd-ancho-columna': 'Column width', 'pregunta-nombre-hoja': 'Sheet name:', 'pregunta-quitar-hoja': 'Delete the sheet “{0}” and everything in it?',
    'pregunta-ancho': 'Column width (in characters):', 'aviso-estructura': 'Couldn’t do it', 'aviso-formulas-ods': '{0} formulas couldn’t be written to .ods: only their value was kept', 'x-formato-condicional': 'conditional formatting', 'x-filtros': 'autofilters', 'x-formulas-sin-traducir': 'formulas shown here only as values',
    'cmd-alto-fila': 'Row height', 'pregunta-alto': 'Row height (in points, 0 to 409; empty = automatic):',
    'cmd-reemplazar': 'Find and replace', 'cmd-ordenar-asc': 'Sort A to Z', 'cmd-ordenar-desc': 'Sort Z to A', 'cmd-pdf': 'Download sheet as PDF', 'cmd-pdf-libro': 'Download workbook as PDF', 'g-ordenar': 'Sort',
    'r-buscar': 'Find', 'r-reemplazar': 'Replace with', 'r-mayus': 'Match case', 'r-entera': 'Match entire cell', 'r-libro': 'Whole workbook', 'r-todo': 'Replace all', 'cerrar': 'Close', 'r-hechos': '{0} cells changed',
    'aviso-pdf': 'Preparing the PDF…', 'aviso-pdf-listo': 'PDF ready: {0} pages', 'aviso-ordenar-combinadas': 'Can’t sort: there are merged cells in the range', 'aviso-ordenar-matriz': 'Can’t sort: the range includes a dynamic array',
  };
  const pagina = window.i18nPagina || (window.i18nPagina = {});
  pagina.es = Object.assign(pagina.es || {}, ES);
  pagina.en = Object.assign(pagina.en || {}, EN);
  const tx = (k, ...a) => { let s = typeof window.t === 'function' ? window.t(k) : (ES[k] || k); a.forEach((x, i) => { s = s.replace('{' + i + '}', x); }); return s; };
  const idioma = () => (typeof window.idiomaActual === 'function' && window.idiomaActual() === 'en' ? 'en' : 'es');

  // ---- estado ------------------------------------------------------------------------------------------------------
  const $ = (id) => document.getElementById(id);
  const MAX_PEGAR = 200000;
  let ed = null, hist = null, rejilla = null, docId = null, nombreDoc = 'Libro', convertido = false, lectura = false, cosasOds = null;
  const autoguardado = Autoguardado.crear('hojas');
  const comandos = Comandos.crear();
  // esqueleto del editor (arquetipo G, núcleo): lo propio de Hojas es la barra de fórmulas y la rejilla en el lienzo
  const crearEl = Oficina.crearEl;
  const interfaz = Oficina.montar($('app'), {
    nombre: 'nombre-app', inicio: '../',
    extra: crearEl('div', { class: 'hc-formula' }, [
      crearEl('input', { class: 'hc-nombre', id: 'cajaNombre', 'aria-label': tx('cuadro-nombre'), 'data-i18n-attr': 'aria-label:cuadro-nombre', spellcheck: 'false', autocomplete: 'off' }),
      crearEl('span', { class: 'hc-fx', 'aria-hidden': 'true' }, ['fx']),
      crearEl('textarea', { class: 'hc-contenido', id: 'barraContenido', rows: '1', 'aria-label': tx('contenido-celda'), 'data-i18n-attr': 'aria-label:contenido-celda', spellcheck: 'false', autocomplete: 'off' }),
      crearEl('button', { class: 'hc-expandir', type: 'button', 'data-comando': 'expandirBarra', 'aria-expanded': 'false', 'aria-label': tx('cmd-expandir-barra'), 'data-i18n-attr': 'aria-label:cmd-expandir-barra;title:cmd-expandir-barra', title: tx('cmd-expandir-barra') }, ['⌄']),
    ]),
  });
  interfaz.lienzo.append(crearEl('div', { class: 'hc-rejilla', id: 'rejilla' }));
  const insp = Oficina.inspector(interfaz.inspector);
  let portapapeles = null;   // lo último copiado aquí: PegadoEspecial.capturar + { texto, cortado, familia } (para pegar fórmulas y formato sin pérdida)
  let familia = 0;           // el documento abierto: al deshacer o cambiar la estructura sigue siendo el mismo
  let ultimoPegado = null;   // para cambiar lo pegado con el botón que aparece al lado: { h, s, datos }
  let repegando = false;
  let pegadoDeCorte = null;  // el texto de lo cortado que ya se pegó
  let teclaValores = false;  // Ctrl+Mayús+V: el evento de pegar que sigue es de valores
  document.addEventListener('keydown', (e) => { if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === 'v') teclaValores = true; }, true);
  document.addEventListener('keyup', (e) => { if (e.key.toLowerCase() === 'v') setTimeout(() => { teclaValores = false; }, 0); }, true);

  // ---- abrir ----------------------------------------------------------------------------------------------------
  const esXlsx = (n) => /\.(xlsx|xlsm|ods)$/i.test(n);   // los que se abren y se guardan en su propio formato (.ods con el backend de ODS)
  async function bytesDeArchivo(archivo) {
    if (esXlsx(archivo.name)) return { bytes: new Uint8Array(await archivo.arrayBuffer()), convertido: false };
    // CSV, .ods, .xls, .xlsb: el lector del sitio los pasa a .xlsx (con sus valores y formatos de número)
    const lector = await LectorHojas.abrir(archivo);
    if (lector.bytesXlsx) return { bytes: lector.bytesXlsx, convertido: true };
    if (lector.libroXlsx) return { bytes: new Uint8Array(await archivo.arrayBuffer()), convertido: false };
    const hojas = [];
    for (const h of lector.hojas) { const x = await lector.cargar(h.indice); hojas.push({ nombre: h.nombre, filas: x.matrizValores(true) }); }
    return { bytes: await XlsxEscritor.crearLibro(hojas), convertido: true };
  }
  async function abrirArchivo(archivo) {
    mostrarCargando(true);
    try {
      const { bytes, convertido: c } = await bytesDeArchivo(archivo);
      await abrirBytes(bytes, archivo.name.replace(/\.[^.]+$/, ''), { convertido: c });
    } catch (e) { mostrarError(e); }
    mostrarCargando(false);
  }
  // Lo que trae un .ods y aquí no se edita (para la línea honesta); se mira una vez al abrir, en el texto de content.xml
  async function cosasDeOds() {
    const c = new Set(), t = (await ed.paquete.leer('content.xml')) || '';
    if (t.includes('<office:annotation')) c.add('x-notas');
    // los gráficos con objeto incrustado se dibujan y se editan; quedan por conservar los que no se saben dibujar, las imágenes y las formas
    const conObjeto = ed.hojas.reduce((n, h) => n + ((h && h._marcos) ? h._marcos.length : 0), 0);
    const sinDibujar = ed.libro._graficos ? ed.libro._graficos.graficos.some((g) => !g.virtual) : false;
    if (sinDibujar || (t.match(/<draw:frame\b/g) || []).length > conObjeto || /<draw:(?:custom-shape|rect|ellipse|line|polyline|polygon|path|control|g|caption|connector|measure)\b/.test(t)) c.add('x-graficos');
    if (ed.paquete.rutas().some((r) => /^(Basic|Scripts)\//.test(r))) c.add('x-macros');
    if (ed.libro.avisos && ed.libro.avisos.formulasSinTraducir) c.add('x-formulas-sin-traducir');
    for (const hoja of ed.hojas) for (const fx of hoja.celdas.formulas.values()) if (fx.tipo === 'array' && !fx.dinamica) { c.add('x-matriciales'); break; }
    return c;
  }
  async function abrirBytes(...args) { familia++; return abrirBytesDe(...args); }
  async function abrirBytesDe(bytes, nombre, opciones) {
    const o = opciones || {};
    await MotorHojas.cargarDibujo();
    const nuevo = await LibroEditable.abrir(bytes, { idioma: idioma() });
    ed = nuevo;
    convertido = !!o.convertido;
    cosasOds = ed.formato === 'ods' ? await cosasDeOds() : null;
    nombreDoc = nombre || 'Libro';
    docId = o.id || Autoguardado.nuevoId();
    hist = HistorialDoc.crear({ aplicar: aplicarApp, alCambiar: alCambiarHistorial });
    $('inicio').hidden = true; $('app').hidden = false; document.documentElement.classList.add('of-abierta');
    $('nombreDoc').value = nombreDoc;
    await crearRejilla();
    lectura = matchMedia('(max-width: 900px)').matches && !o.nuevo;   // en el celular se abre para leer; "Editar" activa la edición
    aplicarModo();
    lineaHonesta();
    actualizarBarra(); actualizarInspector(); actualizarGuardado(); comandos.refrescar();
    rejilla.enfocar();
    if (o.recuperado) autoguardado.programar(datosAutoguardado, 10);
  }
  async function crearRejilla(vista) {
    if (rejilla) { ultimaVista = { hoja: rejilla.hojaActual(), sel: rejilla.seleccion() }; rejilla.destruir(); rejilla = null; }
    const lector = { libroMotor: ed.libro, hojas: ed.nombres().filter((x) => x.tipo === 'hoja').map((x) => ({ nombre: x.nombre, indice: x.indice, estado: x.estado })), opciones: { idioma: idioma(), decimal: idioma() === 'en' ? '.' : ',', miles: idioma() === 'en' ? ',' : '.' } };
    rejilla = await RejillaHoja.crear($('rejilla'), {
      libro: lector, editable: true,
      alEditar: (x) => abrirEditor(x), alBorrar: (r) => borrar(r), alCopiar: (e, cortar) => copiar(e, cortar), alPegar: (e) => { const valores = teclaValores; teclaValores = false; pegar(e.clipboardData, { valores }); }, alEscape: () => { soltarCorte(); for (const g of ganchos.alEscape) g(); }, alSoltar: () => { for (const g of ganchos.alSoltar) g(); },
      alTecla: () => false,
    });
    rejilla.alCambiar(() => { cerrarEditorSiOtraCelda(); actualizarBarra(); actualizarInspector(); comandos.refrescar(); for (const g of ganchos.alSeleccionar) g(); });
    if (vista) {
      const k = Math.min(vista.hoja, ed.libro.hojas.length - 1);
      if (k !== rejilla.hojaActual() && ed.libro.hojas[k] && ed.libro.hojas[k].tipo === 'hoja') await rejilla.mostrarHoja(k);
      if (vista.sel) rejilla.seleccionar(vista.sel.f0, vista.sel.c0, Math.min(vista.sel.f1, 1048575), Math.min(vista.sel.c1, 16383));
    }
  }
  // ---- cambios de estructura: se transforma una copia del paquete y el historial cambia de un libro al otro --------------------
  let ocupado = false, vistaPendiente = null;
  // mientras la rejilla se vuelve a crear (después de un cambio de estructura) no hay rejilla: la hoja y la selección son las de antes
  let rejillaLista = Promise.resolve(), ultimaVista = { hoja: 0, sel: { f0: 0, c0: 0, f1: 0, c1: 0, hoja: 0, nombre: 'A1' } };
  function aplicarApp(op) {
    if (op.t === 'libro') {
      const previo = { t: 'libro', ed, vista: { hoja: hojaActiva(), sel: seleccion() } };
      ed = op.ed;
      vistaPendiente = op.vista;
      return previo;
    }
    return ed.aplicar(op);
  }
  async function estructural(nombre, transformar, vista) {
    await rejillaLista;
    if (lectura || ocupado || !ed) return;
    if (editando) confirmar(editando.modo === 'barra' ? barra.value : editor.value, 0, 0);
    ocupado = true;
    try {
      const bytes = await ed.guardar();
      const r = await transformar(bytes);
      const nuevo = await LibroEditable.abrir(r.bytes || r, { idioma: idioma() });
      nuevo.recalcularTodo();   // los valores del archivo transformado son los de antes del cambio
      hist.hacer(nombre, [{ t: 'libro', ed: nuevo, vista: Object.assign({ hoja: hojaActiva(), sel: seleccion() }, vista || {}, r.indice !== undefined ? { hoja: r.indice } : {}) }]);
    } catch (e) { console.warn(e); avisar(tx('aviso-estructura') + ': ' + (e && e.message ? e.message : e)); }
    ocupado = false;
  }
  async function libroNuevo() {
    mostrarCargando(true);
    try { const e2 = await LibroEditable.nuevo({ idioma: idioma() }); await abrirBytes(await e2.guardar(), idioma() === 'en' ? 'Book1' : 'Libro1', { nuevo: true }); } catch (e) { mostrarError(e); }
    mostrarCargando(false);
  }
  function mostrarCargando(si) { $('cargando').hidden = !si; if (si) $('errorInicio').hidden = true; }
  function mostrarError(e) { console.error(e); const el = $('errorInicio'); el.textContent = tx('error-abrir') + ': ' + (e && e.message ? e.message : e); el.hidden = false; $('inicio').hidden = false; $('app').hidden = true; document.documentElement.classList.remove('of-abierta'); }

  // ---- la línea honesta: lo que el archivo tiene y aquí no se edita ---------------------------------------------------
  function lineaHonesta() {
    const cosas = new Set();
    let anclasXlsx = 0;   // objetos de dibujo de un .xlsx: si hay más que gráficos (partes xl/charts), el resto son imágenes o formas que no se editan
    for (const hoja of ed.hojas) {
      if (!hoja || hoja.sinCeldas) continue;
      if (ed.formato !== 'ods' && hoja.anclas) anclasXlsx += hoja.anclas.length;
      if (hoja.comentarios && hoja.comentarios.length) cosas.add('x-notas');
      if (hoja.tablas && hoja.tablas.length) cosas.add('x-tablas');
      // formato condicional y validación se editan (fase 3.2); las reglas de Excel 2010 que no, las marca el administrador de reglas
      if ((hoja.formatosCondicionales || []).some((cf) => cf.reglas.some((r) => r.dxfXml !== undefined || (r.barras && r.barras.x14 && !r.barras.cfvo)))) cosas.add('x-condicional');
      if (hoja.minigraficos && hoja.minigraficos.length) cosas.add('x-minigraficos');
      if (hoja.imagenesCelda && hoja.imagenesCelda.size) cosas.add('x-imagenes');
      for (const fx of hoja.celdas.formulas.values()) if (fx.tipo === 'array' && !fx.dinamica) { cosas.add('x-matriciales'); break; }
    }
    if (cosasOds) for (const k of cosasOds) cosas.add(k);
    if (anclasXlsx && anclasXlsx > ed.paquete.rutas().filter((r) => /^xl\/charts\/chart\d+\.xml$/.test(r)).length) cosas.add('x-graficos');
    if (ed.libro.tieneMacros) cosas.add('x-macros');
    if (ed.paquete.rutas().some((r) => /pivotTable/i.test(r))) cosas.add('x-dinamicas');
    const partes = [];
    if (convertido) partes.push(tx('honesta-convertido'));
    partes.push(cosas.size ? tx('honesta-conserva', [...cosas].map((k) => tx(k)).join(', ')) : tx('honesta-todo'));
    $('lineaHonesta').textContent = partes.join(' ');
    $('lineaHonesta').title = $('lineaHonesta').textContent;
  }

  // ---- historial → recálculo, pintado, autoguardado ---------------------------------------------------------------------
  function alCambiarHistorial() {
    soltarCorte();
    const r = ed.terminar();
    if (vistaPendiente) { const v = vistaPendiente; vistaPendiente = null; rejillaLista = crearRejilla(v).then(() => { lineaHonesta(); actualizarBarra(); actualizarInspector(); comandos.refrescar(); rejilla.enfocar(); }).catch((e) => console.warn(e)); }
    else if (rejilla) rejilla.refrescar();
    if (r.circulares) avisar(tx('aviso-circular'));
    else if (r.noSoportadas && r.noSoportadas.length) avisar(tx('aviso-no-soportadas', r.noSoportadas.length));
    actualizarBarra(); actualizarInspector(); actualizarGuardado(); comandos.refrescar();
    autoguardado.programar(datosAutoguardado, 2500, (fecha) => { if (fecha) actualizarGuardado(fecha); });
    for (const g of ganchos.alCambiar) g(r);
  }
  async function datosAutoguardado() { if (!ed || !hist.sucio()) return null; return { id: docId, nombre: nombreDoc, bytes: await ed.guardar() }; }
  let ultimoAutoguardado = null;
  function actualizarGuardado(fecha) {
    if (fecha) ultimoAutoguardado = fecha;
    const el = $('estadoGuardado');
    if (!hist || !hist.sucio()) { el.textContent = tx('sin-cambios'); return; }
    el.textContent = ultimoAutoguardado ? tx('guardado-local', new Date(ultimoAutoguardado).toLocaleTimeString(idioma() === 'en' ? 'en-US' : 'es-CR', { hour: '2-digit', minute: '2-digit' })) : tx('cambios-sin-descargar');
  }
  function avisar(texto) { Oficina.aviso($('avisoEstado'), texto, 5000); }
  const ganchos = { antesDeEscribir: [], despuesDeEscribir: [], alSeleccionar: [], cuadroNombre: [], alCambiar: [], alPegar: [], alEscape: [], alSoltar: [] };
  function hacer(nombre, op, extra) {
    if (!op) return false;
    if (op.error) { avisar(tx(op.error === 'protegida' ? 'aviso-protegida' : 'aviso-formula')); return false; }
    if (!op.celdas.length) return false;
    return hist.hacer(nombre, [op], extra);
  }

  // ---- selección y celdas -------------------------------------------------------------------------------------------------
  const hojaActiva = () => (rejilla ? rejilla.hojaActual() : ultimaVista.hoja);
  const seleccion = () => (rejilla ? rejilla.seleccion() : ultimaVista.sel);
  const aA1 = (f, c) => LibroXlsx.numeroACol(c) + (f + 1);
  function textoEdicion(h, f, c) {
    const x = ed.celda(h, f, c);
    if (!x) return '';
    const xf = ed.libro.estilos.xf(x.estilo);
    const formato = xf.formato || 'General';
    return EntradaCelda.textoParaEditar(x.formula ? Object.assign({}, x, { formula: ed.formulaVista(h, f, c) }) : x, {
      idioma: idioma(), esFecha: typeof x.valor === 'number' && XlsxFormato.esFecha(formato), esPorcentaje: /%/.test(formato.replace(/"[^"]*"/g, '')),
      textoFecha: (v) => XlsxFormato.formatear(v, /[hs]/i.test(formato.replace(/"[^"]*"/g, '')) ? (idioma() === 'en' ? 'm/d/yyyy h:mm' : 'dd/mm/yyyy h:mm') : (idioma() === 'en' ? 'm/d/yyyy' : 'dd/mm/yyyy'), { idioma: idioma(), fecha1904: ed.libro.fecha1904 }).texto,
    });
  }
  function actualizarBarra() {
    if (!rejilla || editando) return;
    const s = seleccion();
    $('cajaNombre').value = s.nombre;
    const a = rejilla.activa();
    // una celda que desbordó una matriz dinámica muestra la fórmula de su ancla, apagada (como Excel)
    const ancla = ed.anclaDe ? ed.anclaDe(s.hoja, a.f, a.c) : null;
    const t = textoEdicion(s.hoja, ancla ? ancla.f : a.f, ancla ? ancla.c : a.c);
    const barra = $('barraContenido');
    barra.value = t; barra.classList.toggle('es-formula', t.charAt(0) === '='); barra.classList.toggle('es-derrame', !!ancla);
    barra.readOnly = lectura;
  }

  // ---- editor de celda ----------------------------------------------------------------------------------------------------
  const editor = document.createElement('textarea');
  editor.className = 'hc-editor'; editor.hidden = true; editor.spellcheck = false; editor.setAttribute('autocomplete', 'off'); editor.setAttribute('autocapitalize', 'off');
  editor.setAttribute('aria-label', tx('contenido-celda'));
  let editando = null;   // { h, f, c, modo: 'tecla'|'f2'|'doble'|'barra' }
  function abrirEditor(x) {
    if (lectura || !ed) return;
    soltarCorte();
    const h = hojaActiva();
    if (ed.protegida(h, x.f, x.c)) { avisar(tx('aviso-protegida')); return; }
    editando = { h, f: x.f, c: x.c, modo: x.modo };
    const raiz = $('rejilla').querySelector('.rh'), vista = raiz.querySelector('.rh-vista');
    if (editor.parentNode !== raiz) raiz.append(editor);
    const rr = raiz.getBoundingClientRect(), rv = vista.getBoundingClientRect(), rc = rejilla.rectCelda(x.f, x.c);
    const zoom = rejilla.zoom() / 100;
    editor.style.left = (rv.left - rr.left + rc.x - 1) + 'px'; editor.style.top = (rv.top - rr.top + rc.y - 1) + 'px';
    editor.style.minWidth = (rc.ancho + 3) + 'px'; editor.style.width = (rc.ancho + 3) + 'px'; editor.style.height = Math.max(rc.alto + 3, 22) + 'px';
    editor.style.fontSize = Math.round(15 * zoom) + 'px';
    const inicial = x.inicial === null || x.inicial === undefined ? textoEdicion(h, x.f, x.c) : '';
    editor.value = inicial;
    editor.hidden = false;
    editor.focus({ preventScroll: true });
    editor.setSelectionRange(editor.value.length, editor.value.length);
    if (x.inicial) {   // la tecla que abrió el editor: el navegador la escribe en el editor recién enfocado; si no llegó, se pone a mano
      const esperado = x.inicial;
      setTimeout(() => { if (editando && editor.value === '' && document.activeElement === editor) { editor.value = esperado; ajustarEditor(); sincronizarBarra(); } }, 0);
    }
    ajustarEditor(); sincronizarBarra();
  }
  function ajustarEditor() {
    editor.style.width = editor.style.minWidth;
    editor.style.width = Math.max(parseFloat(editor.style.minWidth), editor.scrollWidth + 4) + 'px';
    const filas = editor.value.split('\n').length;
    if (filas > 1) editor.style.height = Math.max(parseFloat(editor.style.height), editor.scrollHeight + 2) + 'px';
  }
  function sincronizarBarra() { const b = $('barraContenido'); b.value = editor.value; b.classList.toggle('es-formula', editor.value.charAt(0) === '='); }
  function cerrarEditor(enfocar) {
    if (!editando) return;
    editando = null; editor.hidden = true;
    if (enfocar !== false) rejilla.enfocar();
    actualizarBarra();
  }
  function confirmar(texto, df, dc) {
    if (!editando) return true;
    const { h, f, c } = editando;
    if (!ganchos.antesDeEscribir.every((g) => g(h, f, c, texto))) return false;   // validación de datos (app-datos.js)
    const op = ed.opEscribir(h, f, c, texto);
    if (op.error) { avisar(tx(op.error === 'protegida' ? 'aviso-protegida' : 'aviso-formula')); return false; }
    cerrarEditor();
    hacer('Escribir', op);
    for (const g of ganchos.despuesDeEscribir) g(h, f, c);
    if (df || dc) rejilla.irA(Math.max(0, f + df), Math.max(0, c + dc));
    return true;
  }
  function cerrarEditorSiOtraCelda() { if (editando && editando.modo !== 'barra') { const a = rejilla.activa(); if (a.f !== editando.f || a.c !== editando.c) confirmar(editor.value, 0, 0); } }
  editor.addEventListener('input', () => { ajustarEditor(); sincronizarBarra(); });
  editor.addEventListener('keydown', (e) => {
    if (e.isComposing) return;
    if (e.key === 'Enter' && !e.altKey) { e.preventDefault(); confirmar(editor.value, e.shiftKey ? -1 : 1, 0); return; }
    if (e.key === 'Enter' && e.altKey) { e.preventDefault(); const p = editor.selectionStart; editor.setRangeText('\n', p, editor.selectionEnd, 'end'); ajustarEditor(); sincronizarBarra(); return; }
    if (e.key === 'Tab') { e.preventDefault(); confirmar(editor.value, 0, e.shiftKey ? -1 : 1); return; }
    if (e.key === 'Escape') { e.preventDefault(); cerrarEditor(); return; }
    if (editando && editando.modo === 'tecla' && editor.value.charAt(0) !== '=' && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
      e.preventDefault();   // como Excel: al escribir sobre una celda, las flechas confirman y se mueven
      const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
      confirmar(editor.value, d[0], d[1]);
    }
  });
  editor.addEventListener('blur', () => { setTimeout(() => { if (editando && editando.modo !== 'barra' && document.activeElement !== editor && document.activeElement !== $('barraContenido')) confirmar(editor.value, 0, 0); }, 0); });

  // barra de fórmulas: también edita (en el celular es el editor principal)
  const barra = $('barraContenido');
  barra.addEventListener('focus', () => {
    if (lectura || !ed || editando) { if (editando && editando.modo !== 'barra') { editando.modo = 'barra'; } return; }
    const a = rejilla.activa(), h = hojaActiva();
    if (ed.protegida(h, a.f, a.c)) { avisar(tx('aviso-protegida')); barra.blur(); return; }
    editando = { h, f: a.f, c: a.c, modo: 'barra' };
  });
  barra.addEventListener('input', () => { if (editando && !editor.hidden) { editor.value = barra.value; ajustarEditor(); } barra.classList.toggle('es-formula', barra.value.charAt(0) === '='); });
  barra.addEventListener('keydown', (e) => {
    if (e.isComposing) return;
    if (e.key === 'Enter' && !e.altKey) { e.preventDefault(); if (confirmar(barra.value, e.shiftKey ? -1 : 1, 0)) editor.hidden = true; }
    else if (e.key === 'Escape') { e.preventDefault(); cerrarEditor(); }
    else if (e.key === 'Tab') { e.preventDefault(); confirmar(barra.value, 0, e.shiftKey ? -1 : 1); }
  });
  barra.addEventListener('blur', () => { setTimeout(() => { if (editando && editando.modo === 'barra' && document.activeElement !== barra && document.activeElement !== editor) confirmar(barra.value, 0, 0); }, 0); });

  // cuadro de nombre: ir a una celda o rango (A1, B2:D9, Hoja2!A1)
  $('cajaNombre').addEventListener('keydown', (e) => {
    if (e.key === 'Escape') { e.preventDefault(); actualizarBarra(); rejilla.enfocar(); return; }
    if (e.key !== 'Enter') return;
    e.preventDefault();
    let t = $('cajaNombre').value.trim().replace(/\$/g, '');
    const m = /^(?:'?([^'!]+)'?!)?([A-Za-z]{1,3}\d+)(?::([A-Za-z]{1,3}\d+))?$/.exec(t);
    if (!m) { for (const g of ganchos.cuadroNombre) if (g(t)) return; actualizarBarra(); return; }   // un nombre definido: ir o crearlo (app-datos.js)
    const ir = () => { const a = LibroXlsx.decodificarDireccion(m[2].toUpperCase()), b = m[3] ? LibroXlsx.decodificarDireccion(m[3].toUpperCase()) : a; if (a && b) rejilla.seleccionar(a.f, a.c, b.f, b.c); rejilla.enfocar(); };
    if (m[1]) { const k = ed.nombres().find((x) => x.nombre.toLowerCase() === m[1].toLowerCase()); if (k) rejilla.mostrarHoja(k.indice).then(ir); else actualizarBarra(); } else ir();
  });

  // ---- borrar, formato ----------------------------------------------------------------------------------------------------
  function celdasDe(r, fn) {
    const h = hojaActiva(), C = ed.hojas[h].celdas;
    const f1 = Math.min(r.f1, ed.limite(h).filas - 1), c1 = Math.min(r.c1, ed.limite(h).cols - 1);
    if ((f1 - r.f0 + 1) * (c1 - r.c0 + 1) <= 50000) { for (let f = r.f0; f <= f1; f++) for (let c = r.c0; c <= c1; c++) fn(f, c); }
    else C.paraCada((f, c) => { if (f >= r.f0 && f <= f1 && c >= r.c0 && c <= c1) fn(f, c); });
  }
  function borrar(r) {
    if (lectura) return;
    const h = hojaActiva(), celdas = [];
    celdasDe(r || seleccion(), (f, c) => { const x = ed.estadoDe(h, f, c); if (x && (x.tipo !== 'vacio' || x.formula)) celdas.push({ f, c, v: x.estilo ? { tipo: 'vacio', valor: null, formula: null, estilo: x.estilo } : null }); });
    hacer('Borrar', ed.opEstado(h, celdas));
  }
  function borrarFormato() {
    const h = hojaActiva(), celdas = [];
    celdasDe(seleccion(), (f, c) => { const x = ed.estadoDe(h, f, c); if (x && x.estilo) celdas.push({ f, c, v: x.tipo === 'vacio' && !x.formula ? null : Object.assign({}, x, { estilo: 0 }) }); });
    hacer('Borrar formato', ed.opEstado(h, celdas));
  }
  function estilo(nombre, cambios) {
    if (lectura || !ed) return;
    const s = seleccion();
    hacer(nombre, ed.opEstilo(hojaActiva(), s, cambios));
  }
  const propsActivas = () => { if (!ed || !rejilla) return {}; const a = rejilla.activa(); const x = ed.estadoDe(hojaActiva(), a.f, a.c); return ed.estilos.propiedades(x ? x.estilo : ed.hojas[hojaActiva()].estiloDe(a.f, a.c)); };
  const alternar = (prop, nombre) => () => estilo(nombre, { [prop]: !propsActivas()[prop] });
  const FORMATOS = { general: 'General', numero: '#,##0.00', moneda: '"$"#,##0.00', miles: '#,##0', porcentaje: '0%', fecha: 'dd/mm/yyyy', hora: 'h:mm', texto: '@' };
  function claveFormato(codigo) { for (const k in FORMATOS) if (FORMATOS[k] === codigo) return k; return ''; }
  function decimales(delta) {
    estilo('Decimales', (base) => {
      return { formato: EstilosLibro.cambiarDecimales(ed.libro.estilos.xf(base).formato || 'General', delta) };
    });
  }
  function autosuma() {
    if (lectura) return;
    const h = hojaActiva(), a = rejilla.activa();
    // el bloque de números justo arriba (o a la izquierda), como Excel
    const esNum = (f, c) => { const x = ed.celda(h, f, c); return x && x.tipo === 'numero'; };
    let f0 = a.f - 1; while (f0 >= 0 && esNum(f0, a.c)) f0--;
    let rango;
    if (f0 < a.f - 1) rango = aA1(f0 + 1, a.c) + ':' + aA1(a.f - 1, a.c);
    else { let c0 = a.c - 1; while (c0 >= 0 && esNum(a.f, c0)) c0--; if (c0 < a.c - 1) rango = aA1(a.f, c0 + 1) + ':' + aA1(a.f, a.c - 1); }
    abrirEditor({ f: a.f, c: a.c, inicial: null, modo: 'f2' });
    editor.value = '=' + (idioma() === 'en' ? 'SUM' : 'SUMA') + '(' + (rango || '') + ')';
    editor.setSelectionRange(editor.value.length - 1, editor.value.length - 1);
    ajustarEditor(); sincronizarBarra();
  }

  // ---- portapapeles ---------------------------------------------------------------------------------------------------------
  const escHtml = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  function contenidoCopia(r) {
    const h = hojaActiva(), lim = ed.limite(h);
    const f1 = Math.min(r.f1, Math.max(r.f0, lim.filas - 1)), c1 = Math.min(r.c1, Math.max(r.c0, lim.cols - 1));
    if ((f1 - r.f0 + 1) * (c1 - r.c0 + 1) > MAX_PEGAR) return null;
    const hoja = ed.hojas[h], C = hoja.celdas, filasTexto = [], filasHtml = [], capt = PegadoEspecial.capturar(ed, h, r);
    for (const f of capt.filas) {
      const t = [], hh = [];
      for (let c = r.c0; c <= capt.r.c1; c++) {
        const i = C.indiceDe(f, c);
        const texto = i < 0 ? '' : LibroXlsx.textoCelda(ed.libro, hoja, i, { idioma: idioma(), decimal: idioma() === 'en' ? '.' : ',', miles: idioma() === 'en' ? ',' : '.' }).texto;
        t.push(texto.replace(/[\t\r\n]+/g, ' '));
        const x = ed.estadoDe(h, f, c);
        const p = ed.estilos.propiedades(x ? x.estilo : hoja.estiloDe(f, c));
        const css = [p.negrita ? 'font-weight:bold' : '', p.cursiva ? 'font-style:italic' : '', p.subrayado ? 'text-decoration:underline' : '', p.tachado ? 'text-decoration:line-through' : '', p.colorFuente ? 'color:#' + p.colorFuente : '', p.relleno ? 'background:#' + p.relleno : '', p.h !== 'general' ? 'text-align:' + p.h : '', p.tamano && p.tamano !== 11 ? 'font-size:' + p.tamano + 'pt' : '', p.formato && p.formato !== 'General' ? 'mso-number-format:\'' + String(p.formato).replace(/'/g, '') + '\'' : ''].filter(Boolean).join(';');
        hh.push('<td' + (css ? ' style="' + escHtml(css) + '"' : '') + '>' + escHtml(texto) + '</td>');
      }
      filasTexto.push(t.join('\t')); filasHtml.push('<tr>' + hh.join('') + '</tr>');
    }
    const texto = filasTexto.join('\r\n');
    return { texto, html: '<table>' + filasHtml.join('') + '</table>', interno: Object.assign(capt, { texto, familia }) };
  }
  function copiar(e, cortar) {
    if (!ed) return false;
    const c = contenidoCopia(seleccion());
    if (!c) { avisar(tx('aviso-pegado-grande', MAX_PEGAR)); return true; }
    if (e && e.clipboardData) { e.clipboardData.setData('text/plain', c.texto); e.clipboardData.setData('text/html', c.html); }
    else {
      try { navigator.clipboard.write([new ClipboardItem({ 'text/plain': new Blob([c.texto], { type: 'text/plain' }), 'text/html': new Blob([c.html], { type: 'text/html' }) })]); }
      catch (x) { navigator.clipboard && navigator.clipboard.writeText(c.texto); }
    }
    soltarCorte();
    pegadoDeCorte = null;
    portapapeles = c.interno;
    // cortar no vacía nada: lo marca, y se mueve al pegar (como Excel). Hasta entonces, Esc o cualquier cambio lo cancelan
    if (cortar && !lectura) { portapapeles.cortado = true; rejilla.marcarCorte(Object.assign({ h: portapapeles.h }, portapapeles.r)); avisar(tx('aviso-cortado')); }
    else avisar(tx('aviso-copiado'));
    return true;
  }
  function soltarCorte() {
    if (!portapapeles || !portapapeles.cortado) return;
    pegadoDeCorte = portapapeles.texto;   // lo cancelado ya no se pega: el portapapeles del sistema todavía lo tiene y no se puede vaciar sin permiso
    portapapeles = null;
    if (rejilla) rejilla.marcarCorte(null);
  }
  function pegarCorte(P, h, s) {
    const m = ed.opMover(P.h, P.r, h, s.f0, s.c0);
    if (m.error) { avisar(tx(m.error === 'protegida' ? 'aviso-protegida' : m.error === 'combinada' ? 'aviso-mover-combinada' : m.error === 'matriz' ? 'aviso-mover-matriz' : 'aviso-mover-limite')); return; }
    soltarCorte();   // se pega una sola vez
    const destino = { f0: s.f0, c0: s.c0, f1: s.f0 + P.r.f1 - P.r.f0, c1: s.c0 + P.r.c1 - P.r.c0 };
    // lo que no es una celda (nombres, reglas, gráficos, combinadas) se reescribe en el archivo, como un cambio de estructura
    if (m.objetos) return estructural('Mover', (b) => MoverObjetos.aplicar(b, P.h, P.r, h, s.f0, s.c0, { idioma: idioma() }), { hoja: h, sel: destino });
    if (!m.ops.length || !hist.hacer('Mover', m.ops)) return;
    rejilla.seleccionar(destino.f0, destino.c0, destino.f1, destino.c1);
  }
  // matriz de textos desde lo pegado: tabla HTML (Excel, Google, la web) o texto separado por tabuladores
  function matrizDe(html, texto) {
    if (html && /<table/i.test(html)) {
      const doc = new DOMParser().parseFromString(html, 'text/html');   // documento inerte: no corre scripts ni carga imágenes
      const tabla = doc.querySelector('table');
      if (tabla) {
        const filas = [];
        for (const tr of tabla.querySelectorAll('tr')) {
          const fila = [];
          for (const td of tr.querySelectorAll('td,th')) { fila.push(td.innerText !== undefined && td.innerText !== '' ? td.innerText : td.textContent.replace(/\s+/g, ' ').trim()); const n = Math.min(100, +(td.getAttribute('colspan') || 1)); for (let k = 1; k < n; k++) fila.push(''); }
          filas.push(fila);
        }
        if (filas.length) return filas;
      }
    }
    if (!texto) return [];
    const lineas = texto.replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n');
    return lineas.map((l) => l.split('\t'));
  }
  async function pegarDesdeBoton(opciones) {
    const o = opcionesDePegado(opciones);
    if (o.tipo !== 'todo' && portapapeles && portapapeles.cortado) { avisar(tx('aviso-valores-cortado')); return; }
    try {
      const items = await navigator.clipboard.read();
      let html = '', texto = '';
      for (const it of items) {
        if (it.types.includes('text/html')) html = await (await it.getType('text/html')).text();
        if (it.types.includes('text/plain')) texto = await (await it.getType('text/plain')).text();
      }
      pegar({ getData: (t) => (t === 'text/html' ? html : t === 'text/plain' ? texto : '') }, o);
    } catch (e) {
      // sin permiso para leer el portapapeles: lo copiado aquí mismo sigue a mano
      if (portapapeles) { const P = portapapeles; pegar({ getData: (t) => (t === 'text/plain' ? P.texto : '') }, o); }
      else avisar(tx('aviso-portapapeles'));
    }
  }
  function opcionesDePegado(opciones) { const o = Object.assign({}, opciones || {}); if (o.valores) o.tipo = 'valores'; if (!o.tipo) o.tipo = 'todo'; return o; }
  const AVISOS_PEGAR = { limite: 'aviso-pegar-limite', solape: 'aviso-pegar-solape', combinada: 'aviso-pegar-combinada', matriz: 'aviso-pegar-matriz', vinculo: 'aviso-pegar-matriz', protegida: 'aviso-protegida' };
  function despuesDePegar(h, s, datos, destino) {
    ultimoPegado = datos ? { h, s, datos } : null;
    if (datos) for (const g of ganchos.alPegar) g(destino);
  }
  function aplicarPlan(plan, h, s, datos, P) {
    if (plan.error) { if (plan.error === 'grande') avisar(tx('aviso-pegado-grande', MAX_PEGAR)); else avisar(tx(AVISOS_PEGAR[plan.error] || 'aviso-estructura')); return; }
    if (plan.objetos) {
      const opc = { idioma: idioma() };
      // lo que no es una celda se escribe en el archivo, como un cambio de estructura
      return estructural('Pegar', async (b) => PegadoEspecial.ejecutar(await LibroEditable.abrir(b, opc), plan, opc), { hoja: h, sel: plan.destino })
        .then(() => rejillaLista).then(() => { if (P && P.familia === familia) P.ed = ed; despuesDePegar(h, s, datos, plan.destino); });
    }
    if (!plan.celdas.length) { rejilla.seleccionar(plan.destino.f0, plan.destino.c0, plan.destino.f1, plan.destino.c1); return; }
    if (hacer('Pegar', ed.opEstado(h, plan.celdas))) {
      rejilla.seleccionar(plan.destino.f0, plan.destino.c0, plan.destino.f1, plan.destino.c1);
      despuesDePegar(h, s, datos, plan.destino);
    }
  }
  function pegarInterno(P, opciones, datos) {
    if (lectura || !ed) return;
    if (P.familia === familia) P.ed = ed;
    return aplicarPlan(PegadoEspecial.planear(ed, P, hojaActiva(), seleccion(), opcionesDePegado(opciones)), hojaActiva(), seleccion(), datos, P);
  }
  // cambiar lo que se acaba de pegar: se deshace y se pega de la otra forma
  async function repegar(opciones) {
    const u = ultimoPegado;
    if (!u) return pegarDesdeBoton(opciones);
    repegando = true;
    try {
      hist.deshacer();
      await rejillaLista;
      if (hojaActiva() !== u.h) await rejilla.mostrarHoja(u.h);
      rejilla.seleccionar(u.s.f0, u.s.c0, u.s.f1, u.s.c1);
      await pegar(u.datos, opciones);
    } finally { repegando = false; }
  }
  function pegar(datos, opciones) {
    if (lectura || !ed || !datos) return;
    const o = opcionesDePegado(opciones);
    const texto = datos.getData('text/plain') || '', html = datos.getData('text/html') || '';
    const h = hojaActiva(), s = seleccion(), celdas = [];
    const igual = (a, b) => a.replace(/\r\n?/g, '\n') === b.replace(/\r\n?/g, '\n');
    if (pegadoDeCorte !== null && igual(texto, pegadoDeCorte)) return;
    if (portapapeles && igual(texto, portapapeles.texto)) {
      const P = portapapeles;
      if (P.cortado && P.ed === ed) {
        if (o.tipo !== 'todo') avisar(tx('aviso-valores-cortado')); else pegarCorte(P, h, s);
        return;
      }
      if (!o.soloTexto) return pegarInterno(P, o, { getData: (k) => (k === 'text/plain' ? texto : k === 'text/html' ? html : '') });
    }
    if (o.op || o.trans || o.vinculo || o.blancos || (o.tipo !== 'todo' && o.tipo !== 'valores')) { avisar(tx('aviso-pegar-externo')); return; }
    const copiaDatos = { getData: (k) => (k === 'text/plain' ? texto : k === 'text/html' ? html : '') };
    if (o.tipo === 'todo' && html && !o.soloTexto) {
      const plan = PegadoHtml.planear(ed, h, s, html, { idioma: idioma() });
      if (plan) return aplicarPlan(plan, h, s, copiaDatos);
    }
    {
      const M = matrizDe(o.soloTexto ? '' : html, texto);
      if (!M.length) return;
      const n = M.reduce((a, f) => a + f.length, 0);
      if (n > MAX_PEGAR) { avisar(tx('aviso-pegado-grande', MAX_PEGAR)); return; }
      const uno = M.length === 1 && M[0].length === 1;
      const alto = uno ? s.f1 - s.f0 + 1 : M.length, ancho = uno ? s.c1 - s.c0 + 1 : Math.max(...M.map((f) => f.length));
      if (alto * ancho > MAX_PEGAR) { avisar(tx('aviso-pegado-grande', MAX_PEGAR)); return; }
      for (let i = 0; i < alto; i++) for (let j = 0; j < ancho; j++) {
        const t = uno ? M[0][0] : (M[i][j] === undefined ? '' : M[i][j]);
        const op = ed.opEscribir(h, s.f0 + i, s.c0 + j, t);
        if (!op.error) celdas.push(...op.celdas);
      }
    }
    if (hacer('Pegar', ed.opEstado(h, celdas))) {
      const alto = Math.max(...celdas.map((x) => x.f)), ancho = Math.max(...celdas.map((x) => x.c));
      rejilla.seleccionar(s.f0, s.c0, Math.max(s.f0, alto), Math.max(s.c0, ancho));
      despuesDePegar(h, s, copiaDatos, { f0: s.f0, c0: s.c0, f1: Math.max(s.f0, alto), c1: Math.max(s.c0, ancho) });
    }
  }

  // ---- descargar y exportar ---------------------------------------------------------------------------------------------
  async function descargar() {
    if (editando) confirmar(editando.modo === 'barra' ? barra.value : editor.value, 0, 0);
    const bytes = await ed.guardar();
    const destino = ed.destino();
    descargarBlob(new Blob([bytes], { type: destino.mime }), nombreDoc + '.' + destino.extension);
    hist.marcarGuardado();
    avisar(tx('descargado'));
    if (ed.libro.avisos && ed.libro.avisos.formulasNoGuardadas) avisar(tx('aviso-formulas-ods', ed.libro.avisos.formulasNoGuardadas));
    autoguardado.borrar(docId);
  }
  function exportarCsv() {
    const h = hojaActiva(), lim = ed.limite(h), hoja = ed.hojas[h], C = hoja.celdas, sep = idioma() === 'en' ? ',' : ';';
    const lineas = [];
    for (let f = 0; f < lim.filas; f++) {
      const fila = [];
      for (let c = 0; c < lim.cols; c++) {
        const i = C.indiceDe(f, c);
        let t = i < 0 ? '' : LibroXlsx.textoCelda(ed.libro, hoja, i, { idioma: idioma(), decimal: idioma() === 'en' ? '.' : ',', miles: '' }).texto;
        if (/[",;\r\n]/.test(t) || t.includes(sep)) t = '"' + t.replace(/"/g, '""') + '"';
        fila.push(t);
      }
      lineas.push(fila.join(sep));
    }
    descargarTexto(lineas.join('\r\n'), nombreDoc + (lim.filas ? '' : '') + '.csv', 'text/csv', true);
  }

  // ---- PDF (el motor de impresión del sitio: mismas páginas que Excel) ---------------------------------------------------------
  async function exportarPdf(todas) {
    if (editando) confirmar(editando.modo === 'barra' ? barra.value : editor.value, 0, 0);
    avisar(tx('aviso-pdf'));
    try {
      await ExcelAPdfMotor.cargarMotor();
      const hojas = todas ? ed.nombres().filter((x) => x.tipo === 'hoja' && x.estado === 'visible').map((x) => x.indice) : [hojaActiva()];
      const r = await ExcelAPdfMotor.generar(ed.libro, { hojas, idioma: idioma(), nombreArchivo: nombreDoc });
      descargarBlob(new Blob([r.bytes], { type: 'application/pdf' }), nombreDoc + '.pdf');
      avisar(tx('aviso-pdf-listo', r.paginas));
    } catch (e) { console.warn(e); avisar(tx('aviso-estructura') + ': ' + (e && e.message ? e.message : e)); }
  }

  // ---- ordenar (de verdad, en el archivo) ----------------------------------------------------------------------------------
  // La región: la selección si es de más de una celda; si no, la región actual alrededor (como Ctrl+* de Excel). Ordena por la columna activa;
  // si la primera fila parece encabezado (textos sobre números), queda arriba. Las fórmulas se corren con su fila.
  function ordenar(descendente) {
    if (lectura) return;
    const h = hojaActiva(), a = rejilla.activa();
    let r = seleccion();
    if (r.f0 === r.f1 && r.c0 === r.c1) r = RejillaOrdenFiltro.regionActual(ed.hojas[h], a.f, a.c);
    const lim = ed.limite(h);
    r = { f0: r.f0, c0: r.c0, f1: Math.min(r.f1, lim.filas - 1), c1: Math.min(r.c1, lim.cols - 1) };
    if (r.f1 <= r.f0) return;
    if ((ed.hojas[h].combinadas || []).some((m) => !(m.f1 < r.f0 || m.f0 > r.f1 || m.c1 < r.c0 || m.c0 > r.c1))) { avisar(tx('aviso-ordenar-combinadas')); return; }
    if (ed.hayDesborde(h, r)) { avisar(tx('aviso-ordenar-matriz')); return; }
    const tipoDe = (f, c) => { const x = ed.celda(h, f, c); return x ? x.tipo : 'vacio'; };
    const encabezado = tipoDe(r.f0, a.c) === 'texto' && tipoDe(r.f0 + 1, a.c) !== 'texto' && tipoDe(r.f0 + 1, a.c) !== 'vacio';
    const f0 = encabezado ? r.f0 + 1 : r.f0;
    const filas = [];
    for (let f = f0; f <= r.f1; f++) { const fila = []; for (let c = r.c0; c <= r.c1; c++) fila.push(ed.estadoDe(h, f, c)); filas.push({ f, fila, clave: ed.celda(h, f, a.c) }); }
    const rango = (x) => (!x || x.tipo === 'vacio' ? 4 : x.tipo === 'numero' ? 0 : x.tipo === 'texto' ? 1 : x.tipo === 'logico' ? 2 : 3);
    const cmp = new Intl.Collator(idioma() === 'en' ? 'en-US' : 'es-CR', { numeric: true, sensitivity: 'base' }).compare;
    filas.sort((p, q) => {
      const a1 = p.clave, b1 = q.clave, ra = rango(a1), rb = rango(b1);
      if (ra === 4 || rb === 4) return ra - rb;   // las vacías siempre al final
      let d = ra - rb;
      if (!d) d = ra === 0 ? a1.valor - b1.valor : ra === 1 ? cmp(String(a1.valor), String(b1.valor)) : ra === 2 ? (a1.valor ? 1 : 0) - (b1.valor ? 1 : 0) : 0;
      return (descendente ? -d : d) || p.f - q.f;
    });
    const celdas = [];
    filas.forEach((x, k) => {
      const destino = f0 + k;
      x.fila.forEach((v, j) => {
        let nv = v ? Object.assign({}, v) : null;
        if (nv) delete nv.fija;
        if (nv && nv.formula && destino !== x.f) nv.formula = LibroXlsx.desplazarFormula(nv.formula, destino - x.f, 0);
        celdas.push({ f: destino, c: r.c0 + j, v: nv });
      });
    });
    hacer(descendente ? 'Ordenar Z → A' : 'Ordenar A → Z', ed.opEstado(h, celdas));
  }

  // ---- buscar y reemplazar ----------------------------------------------------------------------------------------------
  let dialogoReemplazar = null;
  function reemplazarDialogo() {
    if (!dialogoReemplazar) {
      dialogoReemplazar = Oficina.crearEl('dialog', { class: 'of-dialogo', 'aria-label': tx('cmd-reemplazar') });
      dialogoReemplazar.innerHTML = '<form method="dialog" class="hc-reemplazar"><h2></h2><label><span></span><input name="buscar" required autocomplete="off"></label><label><span></span><input name="reemplazar" autocomplete="off"></label>' +
        '<label class="hc-check"><input type="checkbox" name="mayus"> <span></span></label><label class="hc-check"><input type="checkbox" name="entera"> <span></span></label><label class="hc-check"><input type="checkbox" name="libro"> <span></span></label>' +
        '<p class="hc-resultado" role="status"></p><div class="hc-botones"><button type="button" class="boton" value="todo"></button><button type="button" class="boton boton-secundario" value="cerrar"></button></div></form>';
      document.body.append(dialogoReemplazar);
      const q = (s2) => dialogoReemplazar.querySelector(s2);
      q('h2').textContent = tx('cmd-reemplazar');
      const spans = dialogoReemplazar.querySelectorAll('label > span');
      ['r-buscar', 'r-reemplazar', 'r-mayus', 'r-entera', 'r-libro'].forEach((k, i) => { spans[i].textContent = tx(k); });
      q('button[value="todo"]').textContent = tx('r-todo'); q('button[value="cerrar"]').textContent = tx('cerrar');
      q('button[value="cerrar"]').addEventListener('click', () => dialogoReemplazar.close());
      q('button[value="todo"]').addEventListener('click', () => {
        const n = reemplazarTodo(q('[name=buscar]').value, q('[name=reemplazar]').value, { mayus: q('[name=mayus]').checked, entera: q('[name=entera]').checked, libro: q('[name=libro]').checked });
        q('.hc-resultado').textContent = tx('r-hechos', n);
      });
      dialogoReemplazar.addEventListener('close', () => rejilla && rejilla.enfocar());
    }
    dialogoReemplazar.querySelector('.hc-resultado').textContent = '';
    dialogoReemplazar.showModal();
  }
  // reemplaza en los textos y en las fórmulas (en la forma en que se ven en el idioma de la página); los números no se tocan
  function reemplazarTodo(buscar, nuevo, o) {
    if (!buscar || lectura) return 0;
    const esc = buscar.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(o.entera ? '^' + esc + '$' : esc, o.mayus ? 'g' : 'gi');
    const hojas = o.libro ? ed.nombres().filter((x) => x.tipo === 'hoja').map((x) => x.indice) : [hojaActiva()];
    let total = 0;
    const ops = [];
    for (const h of hojas) {
      const C = ed.hojas[h].celdas, celdas = [];
      C.paraCada((f, c) => {
        const x = ed.estadoDe(h, f, c);
        if (!x || x.derrame || ed.protegida(h, f, c)) return;   // lo que desborda no se escribe: es de la fórmula del ancla
        if (x.formula) {
          const vis = EntradaCelda.formulaVisible(ed.formulaVista(h, f, c), idioma());
          re.lastIndex = 0;
          if (!re.test(vis)) return;
          const op = ed.opEscribir(h, f, c, '=' + vis.replace(re, () => nuevo));
          if (!op.error) { celdas.push(...op.celdas); total++; }
        } else if (x.tipo === 'texto') {
          re.lastIndex = 0;
          if (!re.test(String(x.valor))) return;
          celdas.push({ f, c, v: Object.assign({}, x, { valor: String(x.valor).replace(re, () => nuevo) }) });
          total++;
        }
      });
      if (celdas.length) ops.push(ed.opEstado(h, celdas));
    }
    if (ops.length) hist.hacer('Reemplazar', ops);
    return total;
  }

  // ---- inspector ----------------------------------------------------------------------------------------------------------
  function actualizarInspector() {
    if (!ed || !rejilla) { insp.vaciar(); return; }
    const a = rejilla.activa(), h = hojaActiva(), x = ed.celda(h, a.f, a.c);
    const p = propsActivas();
    const tipo = x ? tx('t-' + x.tipo) : tx('t-vacio');
    insp.mostrar([
      { titulo: tx('i-celda') + ' ' + aA1(a.f, a.c), filas: [
        ['i-tipo', tipo], ['i-valor', x && x.valor !== null ? String(x.valor) : '—'], ['i-formula', x && x.formula ? '=' + EntradaCelda.formulaVisible(ed.formulaVista(h, a.f, a.c) || x.formula, idioma()) : '—'],
        ['i-formato', p.formato || 'General'], ['i-fuente', (p.fuente || '') + ' ' + (p.tamano || '') + (p.negrita ? ' · B' : '') + (p.cursiva ? ' · I' : '')],
      ] },
      { titulo: tx('i-archivo'), filas: [
        ['i-hojas', String(ed.nombres().length)], ['i-formulas', String(ed.motor.cantidad || ed.hojas.reduce((n, hj) => n + (hj && hj.celdas ? hj.celdas.formulas.size : 0), 0))],
      ] },
    ]);
  }

  // ---- modo lectura (celular) -----------------------------------------------------------------------------------------------
  function aplicarModo() {
    $('app').classList.toggle('of-lectura', lectura);
    const b = $('botonModo');
    b.textContent = lectura ? '✎ ' + tx('cmd-modo-editar') : '👁 ' + tx('cmd-modo-leer');
    actualizarBarra();
  }

  // ---- comandos ---------------------------------------------------------------------------------------------------------
  const hayLibro = () => !!ed && !!rejilla;
  const EST = () => (ed && ed.formato === 'ods' ? EstructuraOds : EstructuraLibro);   // los cambios de estructura de cada formato
  const esOds = () => !!ed && ed.formato === 'ods';
  // lo que las extensiones aún no hacen en un .ods (etapa 6 del plan, a medida que se hace sale de la lista)
  const SOLO_XLSX = new Set(['tabla']);
  const editable = () => hayLibro() && !lectura;
  const reg = (c) => comandos.registrar(Object.assign({ habilitado: editable }, c));
  reg({ id: 'deshacer', texto: 'cmd-deshacer', atajo: 'mod+z', ejecutar: () => hist.deshacer(), habilitado: () => editable() && hist.puedeDeshacer() });
  reg({ id: 'rehacer', texto: 'cmd-rehacer', atajo: ['mod+y', 'mod+shift+z'], ejecutar: () => hist.rehacer(), habilitado: () => editable() && hist.puedeRehacer() });
  reg({ id: 'cortar', texto: 'cmd-cortar', ejecutar: () => copiar(null, true) });
  reg({ id: 'copiar', texto: 'cmd-copiar', ejecutar: () => copiar(null, false), habilitado: hayLibro });
  reg({ id: 'pegar', texto: 'cmd-pegar', ejecutar: () => pegarDesdeBoton() });
  reg({ id: 'pegarValores', texto: 'cmd-pegar-valores', ejecutar: () => pegarDesdeBoton({ tipo: 'valores' }) });
  reg({ id: 'negrita', texto: 'cmd-negrita', atajo: 'mod+b', ejecutar: alternar('negrita', 'Negrita'), activo: () => !!propsActivas().negrita });
  reg({ id: 'cursiva', texto: 'cmd-cursiva', atajo: 'mod+i', ejecutar: alternar('cursiva', 'Cursiva'), activo: () => !!propsActivas().cursiva });
  reg({ id: 'subrayado', texto: 'cmd-subrayado', atajo: 'mod+u', ejecutar: alternar('subrayado', 'Subrayado'), activo: () => !!propsActivas().subrayado });
  reg({ id: 'tachado', texto: 'cmd-tachado', atajo: 'mod+5', ejecutar: alternar('tachado', 'Tachado'), activo: () => !!propsActivas().tachado });
  reg({ id: 'colorFuente', texto: 'cmd-color', ejecutar: (v) => estilo('Color', { colorFuente: String(v || '#000000').slice(1) }) });
  reg({ id: 'relleno', texto: 'cmd-relleno', ejecutar: (v) => estilo('Relleno', { relleno: String(v || '#FFFF00').slice(1) }) });
  reg({ id: 'sinRelleno', texto: 'cmd-sin-relleno', ejecutar: () => estilo('Sin relleno', { relleno: null }) });
  reg({ id: 'izquierda', texto: 'cmd-izquierda', ejecutar: () => estilo('Alinear', { h: propsActivas().h === 'left' ? 'general' : 'left' }), activo: () => propsActivas().h === 'left' });
  reg({ id: 'centro', texto: 'cmd-centro', atajo: 'mod+e', ejecutar: () => estilo('Centrar', { h: propsActivas().h === 'center' ? 'general' : 'center' }), activo: () => propsActivas().h === 'center' });
  reg({ id: 'derecha', texto: 'cmd-derecha', ejecutar: () => estilo('Alinear', { h: propsActivas().h === 'right' ? 'general' : 'right' }), activo: () => propsActivas().h === 'right' });
  reg({ id: 'ajustar', texto: 'cmd-ajustar', ejecutar: () => estilo('Ajustar texto', { ajustar: !propsActivas().ajustar }), activo: () => !!propsActivas().ajustar });
  reg({ id: 'formato', texto: 'cmd-formato', ejecutar: (v) => { if (FORMATOS[v]) estilo('Formato', { formato: FORMATOS[v] }); rejilla.enfocar(); }, valor: () => claveFormato(propsActivas().formato) });
  reg({ id: 'masDecimales', texto: 'cmd-mas-decimales', ejecutar: () => decimales(1) });
  reg({ id: 'menosDecimales', texto: 'cmd-menos-decimales', ejecutar: () => decimales(-1) });
  reg({ id: 'bordes', texto: 'cmd-bordes', ejecutar: () => estilo('Bordes', { bordes: 'todos' }) });
  reg({ id: 'sinBordes', texto: 'cmd-sin-bordes', ejecutar: () => estilo('Sin bordes', { bordes: 'ninguno' }) });
  reg({ id: 'borrar', texto: 'cmd-borrar', ejecutar: () => borrar() });
  reg({ id: 'borrarFormato', texto: 'cmd-borrar-formato', ejecutar: borrarFormato });
  reg({ id: 'autosuma', texto: 'cmd-autosuma', atajo: 'alt+=', ejecutar: autosuma });
  reg({ id: 'hoy', texto: 'cmd-hoy', atajo: 'mod+;', ejecutar: () => { const a = rejilla.activa(); const d = new Date(); hacer('Fecha', ed.opEscribir(hojaActiva(), a.f, a.c, idioma() === 'en' ? (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear() : d.getDate() + '/' + (d.getMonth() + 1) + '/' + d.getFullYear())); } });
  reg({ id: 'funcion', texto: 'cmd-funcion', ejecutar: () => { const n = prompt(tx('funcion-pregunta')); if (!n) return; const a = rejilla.activa(); abrirEditor({ f: a.f, c: a.c, inicial: null, modo: 'f2' }); editor.value = '=' + n.trim().toUpperCase() + '()'; editor.setSelectionRange(editor.value.length - 1, editor.value.length - 1); ajustarEditor(); sincronizarBarra(); } });
  reg({ id: 'editarCelda', texto: 'cmd-editar-celda', ejecutar: () => { const a = rejilla.activa(); abrirEditor({ f: a.f, c: a.c, inicial: null, modo: 'f2' }); } });
  reg({ id: 'buscar', texto: 'cmd-buscar', ejecutar: () => rejilla.buscar(''), habilitado: hayLibro });
  reg({ id: 'reemplazar', texto: 'cmd-reemplazar', atajo: 'mod+h', ejecutar: reemplazarDialogo });
  reg({ id: 'ordenarAsc', texto: 'cmd-ordenar-asc', ejecutar: () => ordenar(false) });
  reg({ id: 'ordenarDesc', texto: 'cmd-ordenar-desc', ejecutar: () => ordenar(true) });
  reg({ id: 'pdf', texto: 'cmd-pdf', atajo: 'mod+p', enTexto: true, ejecutar: () => exportarPdf(false), habilitado: hayLibro });
  reg({ id: 'pdfLibro', texto: 'cmd-pdf-libro', ejecutar: () => exportarPdf(true), habilitado: hayLibro });
  reg({ id: 'descargar', texto: 'cmd-descargar', atajo: 'mod+s', enTexto: true, ejecutar: descargar, habilitado: hayLibro });
  reg({ id: 'csv', texto: 'cmd-csv', ejecutar: exportarCsv, habilitado: hayLibro });
  reg({ id: 'nuevo', texto: 'cmd-nuevo', ejecutar: () => { if (!hist || !hist.sucio() || confirm(tx('salir-sin-descargar'))) libroNuevo(); }, habilitado: () => true });
  reg({ id: 'abrir', texto: 'cmd-abrir', ejecutar: () => $('entradaArchivo').click(), habilitado: () => true });
  reg({ id: 'paleta', texto: 'cmd-paleta', atajo: 'mod+k', enTexto: true, ejecutar: () => Oficina.paleta(comandos, () => rejilla && rejilla.enfocar()), habilitado: hayLibro, oculto: true });
  reg({ id: 'zoomMas', texto: 'cmd-zoom-mas', ejecutar: () => rejilla.zoom(rejilla.zoom() + 10), habilitado: hayLibro });
  reg({ id: 'zoomMenos', texto: 'cmd-zoom-menos', ejecutar: () => rejilla.zoom(rejilla.zoom() - 10), habilitado: hayLibro });
  reg({ id: 'completa', texto: 'cmd-completa', ejecutar: () => { if (document.fullscreenElement) document.exitFullscreen(); else document.documentElement.requestFullscreen().catch(() => {}); }, habilitado: hayLibro });
  // vista ampliada de la barra de fórmulas (como Excel: Ctrl+Mayús+U); se recuerda en este navegador
  const barraFormulas = document.querySelector('.hc-formula');
  const barraExpandida = () => barraFormulas.classList.contains('expandida');
  function ponerBarraExpandida(si) {
    barraFormulas.classList.toggle('expandida', si);
    const b = barraFormulas.querySelector('.hc-expandir'); if (b) b.setAttribute('aria-expanded', si ? 'true' : 'false');
    try { localStorage.setItem('hojas-barra-expandida', si ? '1' : '0'); } catch (e) { /* sin almacenamiento: no se recuerda */ }
  }
  try { if (localStorage.getItem('hojas-barra-expandida') === '1') ponerBarraExpandida(true); } catch (e) { /* ... */ }
  reg({ id: 'expandirBarra', texto: 'cmd-expandir-barra', atajo: 'mod+shift+u', enTexto: true, ejecutar: () => ponerBarraExpandida(!barraExpandida()), activo: barraExpandida, habilitado: hayLibro });
  reg({ id: 'inspector', texto: 'cmd-inspector', ejecutar: () => insp.alternar(), activo: () => insp.visible(), habilitado: hayLibro });
  reg({ id: 'modoEdicion', texto: 'cmd-modo-editar', ejecutar: () => { lectura = !lectura; aplicarModo(); comandos.refrescar(); }, habilitado: hayLibro });

  const filasSel = () => { const s = seleccion(); return { f0: s.f0, f1: Math.min(s.f1, 1048575), c0: s.c0, c1: Math.min(s.c1, 16383) }; };
  reg({ id: 'insertarFilas', texto: 'cmd-insertar-filas', ejecutar: () => { const s = filasSel(); estructural('Insertar filas', (b) => EST().insertarFilas(b, hojaActiva(), s.f0, s.f1 - s.f0 + 1)); } });
  reg({ id: 'quitarFilas', texto: 'cmd-quitar-filas', ejecutar: () => { const s = filasSel(); const filas = []; for (let f = s.f0; f <= Math.min(s.f1, s.f0 + 100000); f++) filas.push(f); estructural('Quitar filas', (b) => EST().quitarFilas(b, hojaActiva(), filas)); } });
  reg({ id: 'insertarColumnas', texto: 'cmd-insertar-columnas', ejecutar: () => { const s = filasSel(); estructural('Insertar columnas', (b) => EST().insertarColumnas(b, hojaActiva(), s.c0, s.c1 - s.c0 + 1, { nombreColumna: idioma() === 'en' ? 'Column' : 'Columna' })); } });
  reg({ id: 'quitarColumnas', texto: 'cmd-quitar-columnas', ejecutar: () => { const s = filasSel(); const cols = []; for (let c = s.c0; c <= s.c1; c++) cols.push(c); estructural('Quitar columnas', (b) => EST().quitarColumnas(b, hojaActiva(), cols)); } });
  reg({ id: 'insertarHoja', texto: 'cmd-insertar-hoja', ejecutar: () => {
    const usados = ed.libro.hojas.map((x) => x.nombre.toLowerCase()); let k = ed.libro.hojas.length + 1; const base = idioma() === 'en' ? 'Sheet' : 'Hoja';
    while (usados.includes((base + k).toLowerCase())) k++;
    estructural('Insertar hoja', (b) => EST().agregarHoja(b, base + k, hojaActiva()));
  } });
  reg({ id: 'renombrarHoja', texto: 'cmd-renombrar-hoja', ejecutar: () => { const h = hojaActiva(); const n = prompt(tx('pregunta-nombre-hoja'), ed.libro.hojas[h].nombre); if (n && n !== ed.libro.hojas[h].nombre) estructural('Renombrar hoja', (b) => EST().renombrarHoja(b, h, n)); } });
  reg({ id: 'quitarHoja', texto: 'cmd-quitar-hoja', ejecutar: () => { const h = hojaActiva(); if (confirm(tx('pregunta-quitar-hoja', ed.libro.hojas[h].nombre))) estructural('Quitar hoja', (b) => EST().quitarHoja(b, h), { hoja: Math.max(0, h - 1), sel: null }); }, habilitado: () => editable() && ed.libro.hojas.length > 1 });
  reg({ id: 'hojaIzquierda', texto: 'cmd-hoja-izquierda', ejecutar: () => { const h = hojaActiva(); estructural('Mover hoja', (b) => EST().moverHoja(b, h, h - 1), { hoja: h - 1 }); }, habilitado: () => editable() && hojaActiva() > 0 });
  reg({ id: 'hojaDerecha', texto: 'cmd-hoja-derecha', ejecutar: () => { const h = hojaActiva(); estructural('Mover hoja', (b) => EST().moverHoja(b, h, h + 1), { hoja: h + 1 }); }, habilitado: () => editable() && hojaActiva() < ed.libro.hojas.length - 1 });
  reg({ id: 'combinar', texto: 'cmd-combinar', ejecutar: () => { const s = filasSel(); if (s.f0 === s.f1 && s.c0 === s.c1) return; estructural('Combinar', (b) => EST().combinar(b, hojaActiva(), s)); } });
  reg({ id: 'separar', texto: 'cmd-separar', ejecutar: () => { const s = filasSel(); estructural('Separar', (b) => EST().separar(b, hojaActiva(), s)); } });
  reg({ id: 'inmovilizar', texto: 'cmd-inmovilizar', ejecutar: () => { const a = rejilla.activa(); estructural('Inmovilizar', (b) => EST().inmovilizar(b, hojaActiva(), a.f, a.c)); } });
  reg({ id: 'movilizar', texto: 'cmd-movilizar', ejecutar: () => estructural('Movilizar', (b) => EST().inmovilizar(b, hojaActiva(), 0, 0)) });
  reg({ id: 'altoFila', texto: 'cmd-alto-fila', ejecutar: () => {
    const s = filasSel(), h = hojaActiva();
    const actual = (ed.hojas[h].filas && ed.hojas[h].filas.get(s.f0)) || {};
    const v = prompt(tx('pregunta-alto'), actual.alto != null ? String(actual.alto).replace('.', idioma() === 'en' ? '.' : ',') : '15');
    if (v === null) return;
    const n = String(v).trim() === '' ? null : parseFloat(String(v).replace(',', '.'));
    if (n !== null && !(n >= 0 && n <= 409)) return;
    // una selección de filas enteras llega hasta la última fila usada (Excel guarda el alto fila por fila)
    const f1 = Math.min(s.f1, Math.max(s.f0, ed.limite(h).filas - 1), s.f0 + 99999);
    estructural('Alto de fila', (b) => EST().altoFilas(b, h, s.f0, f1, n));
  } });
  reg({ id: 'anchoColumna', texto: 'cmd-ancho-columna', ejecutar: () => { const s = filasSel(); const v = prompt(tx('pregunta-ancho'), '10'); const n = parseFloat(String(v || '').replace(',', '.')); if (n > 0 && n <= 255) estructural('Ancho de columna', (b) => EST().anchoColumnas(b, hojaActiva(), s.c0, Math.min(s.c1, s.c0 + 16383), n)); } });

  // ---- extensiones (app-datos.js: nombres, tablas, validación, formato condicional, filtros, gráficos, página) ----------------
  const extension = {
    reg: (c) => reg(Object.assign({}, c, { habilitado: () => (c.habilitado || editable)() && !(esOds() && SOLO_XLSX.has(c.id)) })), DAT: () => (esOds() ? DatosOds : DatosLibro), FIL: () => (esOds() ? FiltrosOds : FiltrosLibro), PAG: () => (esOds() ? PaginaOds : PaginaLibro), GRA: () => (esOds() ? GraficosOds : GraficosLibro), estructural, tx, avisar, hacer, hojaActiva, seleccion, filasSel, idioma, comandos, ganchos, aA1,
    get ed() { return ed; }, get rejilla() { return rejilla; }, get hist() { return hist; }, get lectura() { return lectura; }, editable,
    // el rectángulo de la celda en la pantalla (para poner algo encima: la flecha de una lista, un aviso)
    rectEnPantalla: (f, c) => { const raiz = $('rejilla').querySelector('.rh'), vista = raiz && raiz.querySelector('.rh-vista'); if (!vista) return null; const rv = vista.getBoundingClientRect(), rc = rejilla.rectCelda(f, c); if (!rc) return null; return { x: rv.left + rc.x, y: rv.top + rc.y, w: rc.ancho, h: rc.alto, vista: rv }; },
    cinta: [], menu: [],
    // el PDF de la hoja activa (o de todas las visibles), como «Descargar en PDF»
    pdf: async (todas) => { await ExcelAPdfMotor.cargarMotor(); const hojas = todas ? ed.nombres().filter((x) => x.tipo === 'hoja' && x.estado === 'visible').map((x) => x.indice) : [hojaActiva()]; return ExcelAPdfMotor.generar(ed.libro, { hojas, idioma: idioma(), nombreArchivo: nombreDoc }); },
    confirmarEdicion: () => { if (editando) confirmar(editando.modo === 'barra' ? barra.value : editor.value, 0, 0); },
    pegarDesdeBoton, pegarInterno, repegar, hayCopiaInterna: () => !!(portapapeles && !portapapeles.cortado), pegando: () => repegando, despuesDeCinta: [],
    capturarSeleccion: () => { const c = contenidoCopia(seleccion()); return c ? c.interno : (avisar(tx('aviso-pegado-grande', MAX_PEGAR)), null); },
  };
  for (const ext of window.HojasExtensiones || []) ext(extension);
  const CINTA = [
    { id: 'inicio', texto: 'tab-inicio', grupos: [
      { texto: 'g-portapapeles', items: [{ comando: 'pegar', icono: '📋' }, { comando: 'pegarMenu', icono: '▾' }, { comando: 'pegarValores', icono: '123' }, { comando: 'cortar', icono: '✂' }, { comando: 'copiar', icono: '⧉' }, { comando: 'copiarFormato', icono: '🖌' }] },
      { texto: 'g-fuente', items: [{ comando: 'negrita', icono: 'B' }, { comando: 'cursiva', icono: 'I' }, { comando: 'subrayado', icono: 'U' }, { comando: 'tachado', icono: 'S̶', solo: 'pc' }, { comando: 'colorFuente', tipo: 'color', icono: 'A', porDefecto: '#C00000' }, { comando: 'relleno', tipo: 'color', icono: '◧', porDefecto: '#FFFF00' }, { comando: 'sinRelleno', icono: '⊘', solo: 'pc' }] },
      { texto: 'g-alineacion', items: [{ comando: 'izquierda', icono: '⇤' }, { comando: 'centro', icono: '↔' }, { comando: 'derecha', icono: '⇥' }, { comando: 'ajustar', icono: '↵', solo: 'pc' }] },
      { texto: 'g-numero', items: [{ comando: 'formato', tipo: 'select', opciones: [['', 'cmd-formato'], ['general', 'fmt-general'], ['numero', 'fmt-numero'], ['miles', 'fmt-miles'], ['moneda', 'fmt-moneda'], ['porcentaje', 'fmt-porcentaje'], ['fecha', 'fmt-fecha'], ['hora', 'fmt-hora'], ['texto', 'fmt-texto']] }, { comando: 'menosDecimales', icono: '.0←', solo: 'pc' }, { comando: 'masDecimales', icono: '.00→', solo: 'pc' }] },
      { texto: 'g-celdas', solo: 'pc', items: [{ comando: 'combinar', icono: '⇔' }, { comando: 'bordes', icono: '▦' }, { comando: 'sinBordes', icono: '□' }, { comando: 'borrarFormato', icono: '⌫' }, { comando: 'anchoColumna', icono: '↔|' }, { comando: 'altoFila', icono: '↕' }] },
      { texto: 'g-edicion', items: [{ comando: 'autosuma', icono: 'Σ' }, { comando: 'borrar', icono: '✕' }, { comando: 'buscar', icono: '🔍' }, { comando: 'editarCelda', icono: '✎' }] },
    ] },
    { id: 'insertar', texto: 'tab-insertar', grupos: [
      { texto: 'g-filas', items: [{ comando: 'insertarFilas', icono: '⊞', conTexto: true }, { comando: 'quitarFilas', icono: '⊟', conTexto: true }, { comando: 'insertarColumnas', icono: '⊞', conTexto: true }, { comando: 'quitarColumnas', icono: '⊟', conTexto: true }] },
      { texto: 'g-hojas', items: [{ comando: 'insertarHoja', icono: '＋', conTexto: true }, { comando: 'renombrarHoja', icono: '✎' }, { comando: 'hojaIzquierda', icono: '◀' }, { comando: 'hojaDerecha', icono: '▶' }, { comando: 'quitarHoja', icono: '🗑' }] },
      { texto: 'g-funciones', items: [{ comando: 'autosuma', icono: 'Σ', conTexto: true }, { comando: 'funcion', icono: 'ƒx', conTexto: true }, { comando: 'hoy', icono: '📅', conTexto: true }] },
    ] },
    { id: 'datos', texto: 'tab-datos', grupos: [
      { texto: 'g-ordenar', items: [{ comando: 'ordenarAsc', icono: 'A→Z', conTexto: true }, { comando: 'ordenarDesc', icono: 'Z→A', conTexto: true }] },
      { texto: 'g-buscar', items: [{ comando: 'buscar', icono: '🔍', conTexto: true }, { comando: 'reemplazar', icono: '⇄', conTexto: true }] },
      { texto: 'g-archivo', items: [{ comando: 'descargar', icono: '⤓', conTexto: true }, { comando: 'csv', icono: '⤓', conTexto: true }, { comando: 'pdf', icono: '⎙', conTexto: true }] },
    ] },
    { id: 'vista', texto: 'tab-vista', grupos: [
      { texto: 'g-zoom', items: [{ comando: 'zoomMenos', icono: '−' }, { comando: 'zoomMas', icono: '+' }] },
      { texto: 'g-paneles', items: [{ comando: 'inmovilizar', icono: '❄', conTexto: true }, { comando: 'movilizar', icono: '☀', conTexto: true }] },
      { texto: 'g-ventana', items: [{ comando: 'completa', icono: '⛶', conTexto: true }, { comando: 'inspector', icono: '☰', conTexto: true }, { comando: 'expandirBarra', icono: '⌄', conTexto: true }, { comando: 'paleta', icono: '⌘', conTexto: true }] },
    ] },
  ];
  for (const [tab, grupo] of extension.cinta) { const t = CINTA.find((x) => x.id === tab); if (t) t.grupos.push(grupo); }
  Oficina.cinta($('pestanas'), $('cinta'), CINTA, comandos);
  for (const f of extension.despuesDeCinta) f();
  comandos.conectarBotones(document.querySelector('.of-acciones'));
  comandos.conectarBotones(barraFormulas);
  Oficina.menu($('botonMas'), [{ comando: 'nuevo' }, { comando: 'abrir' }, '-', { comando: 'descargar' }, { comando: 'csv' }, { comando: 'pdf' }, { comando: 'pdfLibro' }, ...extension.menu, '-', { comando: 'reemplazar' }, { comando: 'paleta' }, { comando: 'completa' }, { comando: 'inspector' }], comandos);
  comandos.conectarTeclado(document);

  // nombre del documento
  $('nombreDoc').addEventListener('change', () => { nombreDoc = $('nombreDoc').value.trim().replace(/[\\/:*?"<>|]+/g, '-') || nombreDoc; $('nombreDoc').value = nombreDoc; });
  $('nombreDoc').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); rejilla && rejilla.enfocar(); } });

  // ---- inicio: abrir, nuevo, borradores ------------------------------------------------------------------------------------
  conectarZonaArrastre($('zonaArrastre'), $('entradaArchivo'), (archivos) => { if (archivos[0]) abrirArchivo(archivos[0]); });
  $('botonNuevo').addEventListener('click', libroNuevo);
  async function mostrarBorradores() {
    const lista = await autoguardado.listar();
    const ul = $('listaBorradores');
    ul.textContent = '';
    $('seccionBorradores').hidden = !lista.length;
    for (const b of lista) {
      const li = document.createElement('li');
      const info = document.createElement('span');
      info.innerHTML = '<b></b> <small></small>';
      info.querySelector('b').textContent = b.nombre;
      info.querySelector('small').textContent = new Date(b.fecha).toLocaleString(idioma() === 'en' ? 'en-US' : 'es-CR');
      const rec = Oficina.crearEl('button', { type: 'button', class: 'boton' }, [tx('recuperar')]);
      rec.addEventListener('click', async () => { const r = await autoguardado.leer(b.id); if (r) { mostrarCargando(true); try { await abrirBytes(new Uint8Array(r.bytes), r.nombre, { id: r.id, recuperado: true }); } catch (e) { mostrarError(e); } mostrarCargando(false); } });
      const des = Oficina.crearEl('button', { type: 'button', class: 'boton boton-secundario' }, [tx('descartar')]);
      des.addEventListener('click', async () => { await autoguardado.borrar(b.id); mostrarBorradores(); });
      const acciones = document.createElement('span'); acciones.append(rec, ' ', des);
      li.append(info, acciones);
      ul.append(li);
    }
  }
  mostrarBorradores();
  window.addEventListener('beforeunload', (e) => { if (hist && hist.sucio()) { e.preventDefault(); e.returnValue = ''; } });

  // para las pruebas de navegador
  window.HojasApp = { ordenar, reemplazarTodo, abrirBytes, get ed() { return ed; }, get hist() { return hist; }, get rejilla() { return rejilla; }, comandos, pegar, copiar, descargar: () => ed.guardar() };
})();

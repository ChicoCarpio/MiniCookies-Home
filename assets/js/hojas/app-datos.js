/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Mini Cookies Hojas, fase 3.2 del PLAN: la interfaz de los datos del libro. Administrar nombres (y crearlos desde el cuadro de nombre), dar
// formato como tabla, validación de datos (el diálogo, la comprobación al escribir y la flecha con la lista) y el administrador de reglas de
// formato condicional. Se engancha a app.js como extensión (window.HojasExtensiones): usa sus comandos, su historial y sus avisos; lo que
// cambia el archivo lo hace DatosLibro (datos-libro.js) y se deshace como los cambios de estructura.
(function () {
  'use strict';
  const ES = {
    'g-datos-libro': 'Datos del libro', 'cmd-nombres': 'Administrar nombres', 'cmd-tabla': 'Dar formato como tabla', 'cmd-validacion': 'Validación de datos', 'cmd-reglas': 'Formato condicional', 'cmd-nueva-regla': 'Nueva regla de formato',
    'n-titulo': 'Nombres definidos', 'n-vacio': 'El libro no tiene nombres definidos.', 'n-nuevo': 'Nombre nuevo', 'n-editar': 'Editar', 'n-quitar': 'Quitar', 'n-nombre': 'Nombre', 'n-ref': 'Se refiere a', 'n-ambito': 'Ámbito', 'n-libro': 'Libro',
    'n-no-valido': 'Ese nombre no sirve: empieza con una letra, sin espacios, y no puede parecer una celda (A1) ni R1C1', 'n-pregunta-quitar': '¿Quitar el nombre «{0}»? Las fórmulas que lo usan darán #¿NOMBRE?',
    't-titulo': 'Dar formato como tabla', 't-rango': 'Rango de la tabla', 't-encabezados': 'La tabla tiene encabezados', 't-estilo': 'Estilo', 't-hecha': 'Tabla {0} creada', 't-rango-malo': 'El rango no es válido',
    'e-claro': 'Claro', 'e-medio': 'Medio', 'e-oscuro': 'Oscuro',
    'v-titulo': 'Validación de datos', 'v-tipo': 'Permitir', 'v-cualquiera': 'Cualquier valor (quitar la validación)', 'v-list': 'Lista', 'v-whole': 'Número entero', 'v-decimal': 'Decimal', 'v-date': 'Fecha', 'v-time': 'Hora', 'v-textLength': 'Longitud del texto', 'v-custom': 'Personalizada (fórmula)',
    'v-operador': 'Datos', 'v-valor1': 'Valor', 'v-valor2': 'Hasta', 'v-origen': 'Origen (separa con ; o usa =A1:A9)', 'v-formula': 'Fórmula (debe dar VERDADERO)', 'v-vacias': 'Omitir celdas en blanco', 'v-flecha': 'Mostrar la flecha con la lista',
    'v-estilo': 'Si se escribe otra cosa', 'v-stop': 'Detener', 'v-warning': 'Advertir', 'v-information': 'Informar', 'v-titulo-error': 'Título del mensaje', 'v-error': 'Mensaje de error', 'v-entrada': 'Mensaje al seleccionar la celda (opcional)',
    'v-defecto': 'El valor no coincide con las restricciones de validación definidas para esta celda.', 'v-continuar': '¿Continuar?', 'v-falta-valor': 'Falta el valor',
    'op-between': 'entre', 'op-notBetween': 'no está entre', 'op-equal': 'igual a', 'op-notEqual': 'no igual a', 'op-greaterThan': 'mayor que', 'op-lessThan': 'menor que', 'op-greaterThanOrEqual': 'mayor o igual que', 'op-lessThanOrEqual': 'menor o igual que',
    'r-titulo': 'Reglas de formato condicional de la hoja', 'r-vacio': 'La hoja no tiene reglas.', 'r-nueva': 'Nueva regla', 'r-subir': '↑', 'r-bajar': '↓', 'r-solo-excel': '(de Excel 2010: se conserva, se edita en Excel)',
    'r-aplicar': 'Se aplica a', 'r-tipo': 'Tipo de regla', 'r-cellIs': 'Valor de la celda', 'r-containsText': 'El texto contiene', 'r-notContainsText': 'El texto no contiene', 'r-beginsWith': 'Comienza con', 'r-endsWith': 'Termina con',
    'r-containsBlanks': 'Celdas en blanco', 'r-notContainsBlanks': 'Celdas no en blanco', 'r-top10': 'Superiores o inferiores', 'r-aboveAverage': 'Encima o debajo del promedio', 'r-duplicateValues': 'Valores duplicados', 'r-uniqueValues': 'Valores únicos',
    'r-expression': 'Usar una fórmula', 'r-colorScale': 'Escala de color', 'r-dataBar': 'Barra de datos', 'r-iconSet': 'Conjunto de iconos', 'r-texto': 'Texto', 'r-rango': 'Cuántos', 'r-porcentaje': 'En porcentaje', 'r-abajo': 'Los inferiores',
    'r-encima': 'Posición', 'r-por-encima': 'Por encima del promedio', 'r-por-debajo': 'Por debajo del promedio', 'r-formula': 'Fórmula (VERDADERO = dar formato)', 'r-color1': 'Color del mínimo', 'r-color2': 'Color del medio', 'r-color3': 'Color del máximo', 'r-tres': 'Tres colores',
    'r-color-barra': 'Color de la barra', 'r-iconos': 'Iconos', 'r-relleno': 'Relleno', 'r-fuente': 'Color de la fuente', 'r-negrita': 'Negrita', 'r-con-relleno': 'Con relleno',
    'aceptar': 'Aceptar', 'cancelar': 'Cancelar', 'lista-flecha': 'Elegir de la lista',
  };
  const EN = {
    'g-datos-libro': 'Workbook data', 'cmd-nombres': 'Name manager', 'cmd-tabla': 'Format as table', 'cmd-validacion': 'Data validation', 'cmd-reglas': 'Conditional formatting', 'cmd-nueva-regla': 'New formatting rule',
    'n-titulo': 'Defined names', 'n-vacio': 'The workbook has no defined names.', 'n-nuevo': 'New name', 'n-editar': 'Edit', 'n-quitar': 'Delete', 'n-nombre': 'Name', 'n-ref': 'Refers to', 'n-ambito': 'Scope', 'n-libro': 'Workbook',
    'n-no-valido': 'That name won’t work: start with a letter, no spaces, and it can’t look like a cell (A1) or R1C1', 'n-pregunta-quitar': 'Delete the name “{0}”? Formulas using it will return #NAME?',
    't-titulo': 'Format as table', 't-rango': 'Table range', 't-encabezados': 'My table has headers', 't-estilo': 'Style', 't-hecha': 'Table {0} created', 't-rango-malo': 'The range is not valid',
    'e-claro': 'Light', 'e-medio': 'Medium', 'e-oscuro': 'Dark',
    'v-titulo': 'Data validation', 'v-tipo': 'Allow', 'v-cualquiera': 'Any value (remove validation)', 'v-list': 'List', 'v-whole': 'Whole number', 'v-decimal': 'Decimal', 'v-date': 'Date', 'v-time': 'Time', 'v-textLength': 'Text length', 'v-custom': 'Custom (formula)',
    'v-operador': 'Data', 'v-valor1': 'Value', 'v-valor2': 'To', 'v-origen': 'Source (separate with , or use =A1:A9)', 'v-formula': 'Formula (must return TRUE)', 'v-vacias': 'Ignore blank', 'v-flecha': 'In-cell dropdown',
    'v-estilo': 'If something else is entered', 'v-stop': 'Stop', 'v-warning': 'Warning', 'v-information': 'Information', 'v-titulo-error': 'Message title', 'v-error': 'Error message', 'v-entrada': 'Message when the cell is selected (optional)',
    'v-defecto': 'This value doesn’t match the data validation restrictions defined for this cell.', 'v-continuar': 'Continue?', 'v-falta-valor': 'The value is missing',
    'op-between': 'between', 'op-notBetween': 'not between', 'op-equal': 'equal to', 'op-notEqual': 'not equal to', 'op-greaterThan': 'greater than', 'op-lessThan': 'less than', 'op-greaterThanOrEqual': 'greater than or equal to', 'op-lessThanOrEqual': 'less than or equal to',
    'r-titulo': 'Conditional formatting rules of the sheet', 'r-vacio': 'The sheet has no rules.', 'r-nueva': 'New rule', 'r-subir': '↑', 'r-bajar': '↓', 'r-solo-excel': '(Excel 2010 rule: kept, edit it in Excel)',
    'r-aplicar': 'Applies to', 'r-tipo': 'Rule type', 'r-cellIs': 'Cell value', 'r-containsText': 'Text contains', 'r-notContainsText': 'Text does not contain', 'r-beginsWith': 'Begins with', 'r-endsWith': 'Ends with',
    'r-containsBlanks': 'Blank cells', 'r-notContainsBlanks': 'Non-blank cells', 'r-top10': 'Top or bottom', 'r-aboveAverage': 'Above or below average', 'r-duplicateValues': 'Duplicate values', 'r-uniqueValues': 'Unique values',
    'r-expression': 'Use a formula', 'r-colorScale': 'Color scale', 'r-dataBar': 'Data bar', 'r-iconSet': 'Icon set', 'r-texto': 'Text', 'r-rango': 'How many', 'r-porcentaje': 'As a percentage', 'r-abajo': 'Bottom ones',
    'r-encima': 'Position', 'r-por-encima': 'Above average', 'r-por-debajo': 'Below average', 'r-formula': 'Formula (TRUE = format)', 'r-color1': 'Minimum color', 'r-color2': 'Midpoint color', 'r-color3': 'Maximum color', 'r-tres': 'Three colors',
    'r-color-barra': 'Bar color', 'r-iconos': 'Icons', 'r-relleno': 'Fill', 'r-fuente': 'Font color', 'r-negrita': 'Bold', 'r-con-relleno': 'With fill',
    'aceptar': 'OK', 'cancelar': 'Cancel', 'lista-flecha': 'Pick from the list',
  };
  const pagina = window.i18nPagina || (window.i18nPagina = {});
  pagina.es = Object.assign(pagina.es || {}, ES);
  pagina.en = Object.assign(pagina.en || {}, EN);

  (window.HojasExtensiones = window.HojasExtensiones || []).push((A) => {
    const tx = A.tx, ed = () => A.ed;
    const vis = (f) => EntradaCelda.formulaVisible(String(f).replace(/^=/, ''), A.idioma());
    const alm = (t) => EntradaCelda.formulaAlmacenada(String(t).replace(/^=/, ''), A.idioma());
    const hojaCitada = (n) => (/^[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'");
    const absoluta = (h, r) => hojaCitada(ed().libro.hojas[h].nombre) + '!$' + LibroXlsx.numeroACol(r.c0) + '$' + (r.f0 + 1) + (r.f0 === r.f1 && r.c0 === r.c1 ? '' : ':$' + LibroXlsx.numeroACol(r.c1) + '$' + (r.f1 + 1));
    const relativa = (r) => LibroXlsx.numeroACol(r.c0) + (r.f0 + 1) + (r.f0 === r.f1 && r.c0 === r.c1 ? '' : ':' + LibroXlsx.numeroACol(r.c1) + (r.f1 + 1));
    const comun = { aceptar: tx('aceptar'), cancelar: tx('cancelar'), alCerrar: () => A.rejilla && A.rejilla.enfocar() };

    // ---- nombres ------------------------------------------------------------------------------------------------------
    async function editarNombre(previo) {
      const h = A.hojaActiva();
      const hojas = ed().nombres().filter((x) => x.tipo === 'hoja');
      const v = await Dialogo.formulario(Object.assign({}, comun, {
        titulo: previo ? tx('n-editar') + ': ' + previo.nombre : tx('n-nuevo'),
        campos: [
          { nombre: 'nombre', etiqueta: tx('n-nombre'), tipo: 'texto', valor: previo ? previo.nombre : '', requerido: true },
          { nombre: 'ref', etiqueta: tx('n-ref'), tipo: 'texto', valor: previo ? '=' + vis(previo.ref) : '=' + absoluta(h, A.filasSel()), requerido: true },
          { nombre: 'ambito', etiqueta: tx('n-ambito'), tipo: 'select', valor: previo && previo.hojaLocal !== null && previo.hojaLocal !== undefined ? String(previo.hojaLocal) : '', opciones: [['', tx('n-libro')], ...hojas.map((x) => [String(x.indice), x.nombre])] },
        ],
        validar: (x) => (A.DAT().nombreValido(x.nombre.trim()) ? null : tx('n-no-valido')),
      }));
      if (!v) return;
      const nuevo = { nombre: v.nombre.trim(), ref: alm(v.ref), hoja: v.ambito === '' ? null : +v.ambito };
      await A.estructural('Nombre', (b) => A.DAT().ponerNombre(b, nuevo, previo ? { nombre: previo.nombre, hoja: previo.hojaLocal } : null));
    }
    async function administrarNombres() {
      const lista = ed().libro.nombres.filter((x) => !/^_xlnm\./i.test(x.nombre) && !x.oculto);
      const hojas = ed().libro.hojas;
      const r = await Dialogo.lista(Object.assign({}, comun, {
        titulo: tx('n-titulo'), vacio: tx('n-vacio'), cerrar: tx('cerrar'),
        filas: lista.map((x) => ({ texto: x.nombre + (x.hojaLocal !== null && x.hojaLocal !== undefined && hojas[x.hojaLocal] ? ' · ' + hojas[x.hojaLocal].nombre : ''), detalle: '=' + vis(x.ref), acciones: [{ valor: 'editar', texto: tx('n-editar') }, { valor: 'quitar', texto: tx('n-quitar') }] })),
        botones: [{ valor: 'nuevo', texto: tx('n-nuevo') }],
      }));
      if (!r) return;
      if (r.accion === 'nuevo') return editarNombre(null);
      const x = lista[r.indice];
      if (r.accion === 'editar') return editarNombre(x);
      if (r.accion === 'quitar' && confirm(tx('n-pregunta-quitar', x.nombre))) await A.estructural('Quitar nombre', (b) => A.DAT().quitarNombre(b, { nombre: x.nombre, hoja: x.hojaLocal }));
    }
    // cuadro de nombre: un nombre que existe lleva a su rango; uno nuevo y válido se crea para la selección (como Excel)
    A.ganchos.cuadroNombre.push((t) => {
      if (!ed() || !t) return false;
      const h = A.hojaActiva();
      const x = ed().libro.nombres.find((n) => n.nombre.toLowerCase() === t.toLowerCase() && (n.hojaLocal === h || n.hojaLocal === null || n.hojaLocal === undefined));
      if (x) {
        let nodo; try { nodo = XlsFormulas._analizar(XlsFormulas._tokenizar(x.ref.replace(/^=/, ''))); } catch (e) { return false; }
        if (!nodo || nodo.t !== 'ref' || (nodo.modo !== 'celda' && nodo.modo !== 'area')) return false;
        const hh = nodo.hoja ? ed().libro.hojas.findIndex((y) => y.nombre.toLowerCase() === nodo.hoja.a.toLowerCase()) : h;
        if (hh < 0) return false;
        const f1 = nodo.modo === 'area' ? nodo.f1 : nodo.f0, c1 = nodo.modo === 'area' ? nodo.c1 : nodo.c0;
        const ir = () => { A.rejilla.seleccionar(Math.min(nodo.f0, f1), Math.min(nodo.c0, c1), Math.max(nodo.f0, f1), Math.max(nodo.c0, c1)); A.rejilla.enfocar(); };
        if (hh !== h) A.rejilla.mostrarHoja(hh).then(ir); else ir();
        return true;
      }
      if (!A.editable() || !A.DAT().nombreValido(t)) return false;
      A.estructural('Nombre', (b) => A.DAT().ponerNombre(b, { nombre: t, ref: absoluta(h, A.filasSel()), hoja: null }));
      return true;
    });

    // ---- tablas ---------------------------------------------------------------------------------------------------------
    // la región alrededor de una celda (como Ctrl+A en Excel): crece mientras haya algo escrito pegado al borde
    function regionActual(h, f, c) {
      const vacia = (ff, cc) => { const x = ed().celda(h, ff, cc); return !x || x.tipo === 'vacio' || x.valor === null || x.valor === ''; };
      let r = { f0: f, c0: c, f1: f, c1: c };
      for (let vuelta = 0; vuelta < 2000; vuelta++) {
        let creció = false;
        const borde = (ff, cc) => ff >= 0 && cc >= 0 && ff <= 1048575 && cc <= 16383 && !vacia(ff, cc);
        for (let cc = r.c0 - 1; cc <= r.c1 + 1 && !creció; cc++) { if (borde(r.f0 - 1, cc)) { r.f0--; creció = true; } else if (borde(r.f1 + 1, cc)) { r.f1++; creció = true; } }
        for (let ff = r.f0 - 1; ff <= r.f1 + 1 && !creció; ff++) { if (borde(ff, r.c0 - 1)) { r.c0--; creció = true; } else if (borde(ff, r.c1 + 1)) { r.c1++; creció = true; } }
        if (!creció) break;
      }
      return r;
    }
    const ESTILOS_TABLA = [['TableStyleLight1', 'e-claro', 1], ['TableStyleLight9', 'e-claro', 9], ['TableStyleLight15', 'e-claro', 15], ['TableStyleMedium2', 'e-medio', 2], ['TableStyleMedium4', 'e-medio', 4], ['TableStyleMedium7', 'e-medio', 7], ['TableStyleMedium9', 'e-medio', 9], ['TableStyleMedium14', 'e-medio', 14], ['TableStyleDark1', 'e-oscuro', 1], ['TableStyleDark9', 'e-oscuro', 9]];
    async function formatoTabla() {
      const h = A.hojaActiva();
      let s = A.filasSel();
      if (s.f0 === s.f1 && s.c0 === s.c1) s = regionActual(h, s.f0, s.c0);
      let todasTexto = true;
      for (let c = s.c0; c <= s.c1; c++) { const x = ed().celda(h, s.f0, c); if (!x || x.tipo !== 'texto') todasTexto = false; }
      const v = await Dialogo.formulario(Object.assign({}, comun, {
        titulo: tx('t-titulo'),
        campos: [
          { nombre: 'rango', etiqueta: tx('t-rango'), tipo: 'texto', valor: relativa(s), requerido: true },
          { nombre: 'encabezados', etiqueta: tx('t-encabezados'), tipo: 'check', valor: todasTexto },
          { nombre: 'estilo', etiqueta: tx('t-estilo'), tipo: 'select', valor: 'TableStyleMedium2', opciones: ESTILOS_TABLA.map(([v2, k, n]) => [v2, tx(k) + ' ' + n]) },
        ],
        validar: (x) => (LibroXlsx.decodificarRango(x.rango.replace(/\$/g, '').toUpperCase()) ? null : tx('t-rango-malo')),
      }));
      if (!v) return;
      const r = LibroXlsx.decodificarRango(v.rango.replace(/\$/g, '').toUpperCase());
      let nombre = '';
      await A.estructural('Tabla', async (b) => { const x = await A.DAT().crearTabla(b, h, r, { encabezados: v.encabezados, estilo: v.estilo, nombreColumna: A.idioma() === 'en' ? 'Column' : 'Columna', nombreBase: A.idioma() === 'en' ? 'Table' : 'Tabla' }); nombre = x.nombre; return x.bytes; });
      if (nombre) A.avisar(tx('t-hecha', nombre));
    }

    // ---- validación de datos --------------------------------------------------------------------------------------------
    const OPS = ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'];
    // lo que se escribe en el diálogo → lo que se guarda: =fórmula, una lista "a,b,c" o un número
    function valorGuardado(t, tipo) {
      const s = String(t === null || t === undefined ? '' : t).trim();
      if (!s) return '';
      if (s[0] === '=') return alm(s);
      if (tipo === 'list') { const partes = s.split(s.includes(';') ? ';' : ','); return '"' + partes.map((x) => x.trim()).filter(Boolean).join(',').replace(/"/g, '""') + '"'; }
      if (tipo === 'custom') return alm(s);
      const x = EntradaCelda.interpretar(s, ed().opciones);
      if (x.tipo === 'numero') return String(x.valor);
      return alm(s);
    }
    function valorVisible(f, tipo) {
      if (f === null || f === undefined || f === '') return '';
      const s = String(f);
      if (tipo === 'list' && /^".*"$/.test(s)) return s.slice(1, -1).replace(/""/g, '"').split(',').join(A.idioma() === 'en' ? ',' : ';');
      if (/^-?\d+(\.\d+)?$/.test(s)) return A.idioma() === 'en' ? s : s.replace('.', ',');
      return '=' + vis(s);
    }
    async function validacion() {
      const h = A.hojaActiva(), s = A.filasSel(), a = A.rejilla.activa();
      const v0 = A.DAT().validacionEn(ed(), h, a.f, a.c);
      const tipoActual = v0 && v0.tipo !== 'none' ? v0.tipo : 'list';
      const conOperador = (x) => ['whole', 'decimal', 'date', 'time', 'textLength'].includes(x.tipo);
      const v = await Dialogo.formulario(Object.assign({}, comun, {
        titulo: tx('v-titulo') + ' · ' + relativa(s),
        campos: [
          { nombre: 'tipo', etiqueta: tx('v-tipo'), tipo: 'select', valor: tipoActual, opciones: [['ninguna', tx('v-cualquiera')], ...['list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom'].map((k) => [k, tx('v-' + k)])] },
          { nombre: 'operador', etiqueta: tx('v-operador'), tipo: 'select', valor: (v0 && v0.operador) || 'between', opciones: OPS.map((k) => [k, tx('op-' + k)]), visible: conOperador },
          { nombre: 'valor1', etiqueta: tx('v-valor1'), tipo: 'texto', valor: v0 ? valorVisible(v0.formula1, v0.tipo) : '', visible: (x) => x.tipo !== 'ninguna' && x.tipo !== 'list' && x.tipo !== 'custom' },
          { nombre: 'origen', etiqueta: tx('v-origen'), tipo: 'texto', valor: v0 && v0.tipo === 'list' ? valorVisible(v0.formula1, 'list') : '', visible: (x) => x.tipo === 'list' },
          { nombre: 'formula', etiqueta: tx('v-formula'), tipo: 'texto', valor: v0 && v0.tipo === 'custom' ? valorVisible(v0.formula1, 'custom') : '', visible: (x) => x.tipo === 'custom' },
          { nombre: 'valor2', etiqueta: tx('v-valor2'), tipo: 'texto', valor: v0 ? valorVisible(v0.formula2, v0.tipo) : '', visible: (x) => conOperador(x) && (x.operador === 'between' || x.operador === 'notBetween') },
          { nombre: 'vacias', etiqueta: tx('v-vacias'), tipo: 'check', valor: v0 ? v0.permitirVacio : true, visible: (x) => x.tipo !== 'ninguna' },
          { nombre: 'flecha', etiqueta: tx('v-flecha'), tipo: 'check', valor: v0 ? !v0.ocultarFlecha : true, visible: (x) => x.tipo === 'list' },
          { nombre: 'estilo', etiqueta: tx('v-estilo'), tipo: 'select', valor: (v0 && v0.estiloError) || 'stop', opciones: ['stop', 'warning', 'information'].map((k) => [k, tx('v-' + k)]), visible: (x) => x.tipo !== 'ninguna' },
          { nombre: 'tituloError', etiqueta: tx('v-titulo-error'), tipo: 'texto', valor: v0 ? v0.tituloError : '', visible: (x) => x.tipo !== 'ninguna' },
          { nombre: 'error', etiqueta: tx('v-error'), tipo: 'area', valor: v0 ? v0.error : '', visible: (x) => x.tipo !== 'ninguna' },
          { nombre: 'entrada', etiqueta: tx('v-entrada'), tipo: 'area', valor: v0 ? v0.entrada : '', visible: (x) => x.tipo !== 'ninguna' },
        ],
        validar: (x) => {
          if (x.tipo === 'ninguna') return null;
          const uno = x.tipo === 'list' ? x.origen : x.tipo === 'custom' ? x.formula : x.valor1;
          if (!String(uno).trim()) return tx('v-falta-valor');
          if (conOperador(x) && (x.operador === 'between' || x.operador === 'notBetween') && !String(x.valor2).trim()) return tx('v-falta-valor');
          return null;
        },
      }));
      if (!v) return;
      const regla = v.tipo === 'ninguna' ? null : {
        tipo: v.tipo, operador: conOperador(v) ? v.operador : null,
        formula1: valorGuardado(v.tipo === 'list' ? v.origen : v.tipo === 'custom' ? v.formula : v.valor1, v.tipo), formula2: conOperador(v) ? valorGuardado(v.valor2, v.tipo) : null,
        permitirVacio: v.vacias, ocultarFlecha: v.tipo === 'list' && !v.flecha, estiloError: v.estilo, tituloError: v.tituloError, error: v.error, mostrarEntrada: true, entrada: v.entrada, mostrarError: true,
      };
      await A.estructural('Validación', (b) => A.DAT().ponerValidacion(b, h, [s], regla), { sel: A.seleccion() });
    }
    // al escribir: detener, advertir o informar (como Excel)
    function avisoValidacion(r) {
      const m = (r.titulo ? r.titulo + '\n\n' : '') + (r.mensaje || tx('v-defecto'));
      if (r.estilo === 'warning') return confirm(m + '\n\n' + tx('v-continuar'));
      alert(m);
      return r.estilo === 'information';
    }
    A.ganchos.antesDeEscribir.push((h, f, c, texto) => {
      if (!ed() || !ed().hojas[h].validaciones || !ed().hojas[h].validaciones.length) return true;
      const r = A.DAT().comprobarEntrada(ed(), h, f, c, texto);
      return !r || avisoValidacion(r);
    });
    A.ganchos.despuesDeEscribir.push((h, f, c) => {
      if (!ed() || !ed().hojas[h].validaciones || !ed().hojas[h].validaciones.length) return;
      const r = A.DAT().comprobarDespues(ed(), h, f, c);
      if (r && !avisoValidacion(r)) { A.hist.deshacer(); ed().terminar(); A.rejilla.redibujar && A.rejilla.redibujar(); }
    });

    // la flecha de las listas y el mensaje de entrada, sobre la celda activa
    const flecha = Object.assign(document.createElement('button'), { type: 'button', className: 'hc-flecha-lista', textContent: '▾', hidden: true });
    flecha.setAttribute('aria-label', tx('lista-flecha'));
    const menuLista = Object.assign(document.createElement('div'), { className: 'hc-menu-lista', hidden: true });
    menuLista.setAttribute('role', 'listbox');
    const notaEntrada = Object.assign(document.createElement('div'), { className: 'hc-nota-entrada', hidden: true });
    document.body.append(flecha, menuLista, notaEntrada);
    let celdaFlecha = null;
    function ocultarFlecha() { flecha.hidden = true; menuLista.hidden = true; notaEntrada.hidden = true; celdaFlecha = null; }
    function ponerFlecha() {
      ocultarFlecha();
      if (!ed() || !A.rejilla) return;
      const h = A.hojaActiva(), a = A.rejilla.activa();
      const v = A.DAT().validacionEn(ed(), h, a.f, a.c);
      if (!v) return;
      const r = A.rectEnPantalla(a.f, a.c);
      if (!r || r.x + r.w < r.vista.left || r.y + r.h < r.vista.top || r.x > r.vista.right || r.y > r.vista.bottom) return;
      if (v.tipo === 'list' && !v.ocultarFlecha && A.editable()) {
        flecha.style.left = Math.min(r.x + r.w + 1, r.vista.right - 20) + 'px'; flecha.style.top = r.y + 'px'; flecha.style.height = Math.max(18, r.h) + 'px';
        flecha.hidden = false; celdaFlecha = { h, f: a.f, c: a.c, r };
      }
      if (v.mostrarEntrada && (v.entrada || v.tituloEntrada)) {
        notaEntrada.textContent = '';
        if (v.tituloEntrada) notaEntrada.append(Object.assign(document.createElement('strong'), { textContent: v.tituloEntrada }));
        if (v.entrada) notaEntrada.append(Object.assign(document.createElement('span'), { textContent: v.entrada }));
        notaEntrada.style.left = (r.x + 8) + 'px'; notaEntrada.style.top = (r.y + r.h + 4) + 'px';
        notaEntrada.hidden = false;
      }
    }
    flecha.addEventListener('mousedown', (e) => e.preventDefault());
    flecha.addEventListener('click', () => {
      if (!celdaFlecha) return;
      const { h, f, c, r } = celdaFlecha;
      const ops = A.DAT().listaDe(ed(), h, f, c) || [];
      menuLista.textContent = '';
      for (const o of ops.slice(0, 1000)) {
        const b = Object.assign(document.createElement('button'), { type: 'button', textContent: o });
        b.setAttribute('role', 'option');
        b.addEventListener('click', () => { menuLista.hidden = true; A.hacer('Lista', ed().opEscribir(h, f, c, o)); A.rejilla.enfocar(); });
        menuLista.append(b);
      }
      menuLista.style.left = r.x + 'px'; menuLista.style.top = (r.y + r.h + 1) + 'px'; menuLista.style.minWidth = Math.max(r.w, 120) + 'px';
      notaEntrada.hidden = true;
      menuLista.hidden = !ops.length;
      const primero = menuLista.querySelector('button'); if (primero) primero.focus();
    });
    menuLista.addEventListener('keydown', (e) => {
      const bs = [...menuLista.querySelectorAll('button')], k = bs.indexOf(document.activeElement);
      if (e.key === 'Escape') { e.preventDefault(); menuLista.hidden = true; A.rejilla.enfocar(); }
      else if (e.key === 'ArrowDown' && k < bs.length - 1) { e.preventDefault(); bs[k + 1].focus(); }
      else if (e.key === 'ArrowUp' && k > 0) { e.preventDefault(); bs[k - 1].focus(); }
    });
    document.addEventListener('mousedown', (e) => { if (!menuLista.hidden && !menuLista.contains(e.target) && e.target !== flecha) menuLista.hidden = true; });
    A.ganchos.alSeleccionar.push(ponerFlecha);
    // al desplazarse o cambiar el tamaño, la flecha se vuelve a poner donde está la celda
    let pendiente = 0;
    const reponer = () => { if (pendiente) return; pendiente = requestAnimationFrame(() => { pendiente = 0; if (!menuLista.hidden) return; ponerFlecha(); }); };
    document.addEventListener('scroll', reponer, true); window.addEventListener('resize', reponer); document.addEventListener('wheel', reponer, { passive: true, capture: true });
    // alt+flecha abajo abre la lista, como Excel
    document.addEventListener('keydown', (e) => { if (e.altKey && e.key === 'ArrowDown' && !flecha.hidden && !e.defaultPrevented) { e.preventDefault(); flecha.click(); } });

    // ---- formato condicional --------------------------------------------------------------------------------------------
    const TIPOS_REGLA = ['cellIs', 'containsText', 'notContainsText', 'beginsWith', 'endsWith', 'containsBlanks', 'notContainsBlanks', 'top10', 'aboveAverage', 'duplicateValues', 'uniqueValues', 'expression', 'colorScale', 'dataBar', 'iconSet'];
    const ICONOS = ['3TrafficLights1', '3Arrows', '3ArrowsGray', '3Flags', '3Symbols', '3Signs', '4Arrows', '4Rating', '4TrafficLights', '5Arrows', '5Rating', '5Quarters'];
    const sinFormato = (t) => t === 'colorScale' || t === 'dataBar' || t === 'iconSet';
    function describir(x) {
      const r = x.regla, t = tx('r-' + x.tipo) === 'r-' + x.tipo ? x.tipo : tx('r-' + x.tipo);
      if (x.tipo === 'cellIs') return t + ' ' + tx('op-' + (r.operador || 'between')) + ' ' + r.formulas.map((f) => valorVisible(f)).join(' · ');
      if (r.texto) return t + ' «' + r.texto + '»';
      if (x.tipo === 'top10') return t + ': ' + (r.abajo ? '↓ ' : '↑ ') + r.rango + (r.porcentaje ? ' %' : '');
      if (x.tipo === 'expression' && r.formulas[0]) return t + ': =' + vis(r.formulas[0]);
      return t;
    }
    async function editarRegla(previa) {
      const h = A.hojaActiva();
      const r0 = previa ? previa.regla : null;
      const valorDeRegla = (f) => (f === undefined || f === null ? '' : /^-?\d+(\.\d+)?$/.test(f) ? valorVisible(f) : /^".*"$/.test(f) ? f.slice(1, -1).replace(/""/g, '"') : '=' + vis(f));
      const conOp = (x) => x.tipo === 'cellIs';
      const v = await Dialogo.formulario(Object.assign({}, comun, {
        titulo: previa ? tx('n-editar') : tx('cmd-nueva-regla'),
        campos: [
          { nombre: 'sqref', etiqueta: tx('r-aplicar'), tipo: 'texto', valor: previa ? previa.sqref : relativa(A.filasSel()), requerido: true },
          { nombre: 'tipo', etiqueta: tx('r-tipo'), tipo: 'select', valor: r0 ? r0.tipo : 'cellIs', opciones: TIPOS_REGLA.map((k) => [k, tx('r-' + k)]) },
          { nombre: 'operador', etiqueta: tx('v-operador'), tipo: 'select', valor: (r0 && r0.operador) || 'greaterThan', opciones: ['greaterThan', 'lessThan', 'between', 'notBetween', 'equal', 'notEqual', 'greaterThanOrEqual', 'lessThanOrEqual'].map((k) => [k, tx('op-' + k)]), visible: conOp },
          { nombre: 'valor1', etiqueta: tx('v-valor1'), tipo: 'texto', valor: r0 && r0.tipo === 'cellIs' ? valorDeRegla(r0.formulas[0]) : '', visible: conOp },
          { nombre: 'valor2', etiqueta: tx('v-valor2'), tipo: 'texto', valor: r0 && r0.tipo === 'cellIs' ? valorDeRegla(r0.formulas[1]) : '', visible: (x) => conOp(x) && (x.operador === 'between' || x.operador === 'notBetween') },
          { nombre: 'texto', etiqueta: tx('r-texto'), tipo: 'texto', valor: (r0 && r0.texto) || '', visible: (x) => ['containsText', 'notContainsText', 'beginsWith', 'endsWith'].includes(x.tipo) },
          { nombre: 'rango', etiqueta: tx('r-rango'), tipo: 'numero', valor: (r0 && r0.rango) || 10, visible: (x) => x.tipo === 'top10' },
          { nombre: 'porcentaje', etiqueta: tx('r-porcentaje'), tipo: 'check', valor: !!(r0 && r0.porcentaje), visible: (x) => x.tipo === 'top10' },
          { nombre: 'abajo', etiqueta: tx('r-abajo'), tipo: 'check', valor: !!(r0 && r0.abajo), visible: (x) => x.tipo === 'top10' },
          { nombre: 'encima', etiqueta: tx('r-encima'), tipo: 'select', valor: r0 && r0.encima === false ? 'no' : 'si', opciones: [['si', tx('r-por-encima')], ['no', tx('r-por-debajo')]], visible: (x) => x.tipo === 'aboveAverage' },
          { nombre: 'formula', etiqueta: tx('r-formula'), tipo: 'texto', valor: r0 && r0.tipo === 'expression' && r0.formulas[0] ? '=' + vis(r0.formulas[0]) : '=', visible: (x) => x.tipo === 'expression' },
          { nombre: 'color1', etiqueta: tx('r-color1'), tipo: 'color', valor: '#' + ((r0 && r0.tipo === 'colorScale' && r0.colores[0]) || 'F8696B'), visible: (x) => x.tipo === 'colorScale' },
          { nombre: 'tres', etiqueta: tx('r-tres'), tipo: 'check', valor: !r0 || r0.tipo !== 'colorScale' || r0.colores.length === 3, visible: (x) => x.tipo === 'colorScale' },
          { nombre: 'color2', etiqueta: tx('r-color2'), tipo: 'color', valor: '#' + ((r0 && r0.tipo === 'colorScale' && r0.colores.length === 3 && r0.colores[1]) || 'FFEB84'), visible: (x) => x.tipo === 'colorScale' && x.tres },
          { nombre: 'color3', etiqueta: tx('r-color3'), tipo: 'color', valor: '#' + ((r0 && r0.tipo === 'colorScale' && r0.colores[r0.colores.length - 1]) || '63BE7B'), visible: (x) => x.tipo === 'colorScale' },
          { nombre: 'barra', etiqueta: tx('r-color-barra'), tipo: 'color', valor: '#' + ((r0 && r0.tipo === 'dataBar' && r0.colores[0]) || '638EC6'), visible: (x) => x.tipo === 'dataBar' },
          { nombre: 'iconos', etiqueta: tx('r-iconos'), tipo: 'select', valor: (r0 && r0.iconos) || '3TrafficLights1', opciones: ICONOS.map((k) => [k, k]), visible: (x) => x.tipo === 'iconSet' },
          { nombre: 'conRelleno', etiqueta: tx('r-con-relleno'), tipo: 'check', valor: true, visible: (x) => !sinFormato(x.tipo) },
          { nombre: 'relleno', etiqueta: tx('r-relleno'), tipo: 'color', valor: '#FFC7CE', visible: (x) => !sinFormato(x.tipo) && x.conRelleno },
          { nombre: 'fuente', etiqueta: tx('r-fuente'), tipo: 'color', valor: '#9C0006', visible: (x) => !sinFormato(x.tipo) },
          { nombre: 'negrita', etiqueta: tx('r-negrita'), tipo: 'check', valor: false, visible: (x) => !sinFormato(x.tipo) },
        ],
        validar: (x) => {
          if (!String(x.sqref).trim().split(/[\s;,]+/).every((p) => LibroXlsx.decodificarRango(p.replace(/\$/g, '').toUpperCase()))) return tx('t-rango-malo');
          if (conOp(x) && !String(x.valor1).trim()) return tx('v-falta-valor');
          return null;
        },
      }));
      if (!v) return;
      // en la regla, un valor escrito: número, =fórmula o texto (entre comillas, como lo guarda Excel)
      const valorRegla = (t) => { const s = String(t || '').trim(); if (!s) return ''; if (s[0] === '=') return alm(s); const x = EntradaCelda.interpretar(s, ed().opciones); return x.tipo === 'numero' ? String(x.valor) : '"' + s.replace(/"/g, '""') + '"'; };
      const regla = {
        tipo: v.tipo, operador: v.operador, formulas: v.tipo === 'cellIs' ? [valorRegla(v.valor1), valorRegla(v.valor2)].filter(Boolean) : v.tipo === 'expression' ? [alm(v.formula)] : [],
        texto: v.texto, rango: v.rango, porcentaje: v.porcentaje, abajo: v.abajo, encima: v.encima !== 'no', iconos: v.iconos,
        colores: v.tipo === 'colorScale' ? (v.tres ? [v.color1, v.color2, v.color3] : [v.color1, v.color3]) : v.tipo === 'dataBar' ? [v.barra] : [],
        estilo: { relleno: v.conRelleno ? v.relleno : null, color: v.fuente, negrita: v.negrita },
      };
      const sqref = String(v.sqref).trim().toUpperCase().replace(/\$/g, '').split(/[\s;,]+/).join(' ');
      await A.estructural(previa ? 'Cambiar regla' : 'Nueva regla', (b) => (previa ? A.DAT().cambiarRegla(b, h, previa.id, sqref, regla) : A.DAT().agregarRegla(b, h, sqref, regla)));
    }
    async function administrarReglas() {
      const h = A.hojaActiva();
      const reglas = await A.DAT().reglasCondicionales(await ed().guardar(), h);
      const r = await Dialogo.lista(Object.assign({}, comun, {
        titulo: tx('r-titulo'), vacio: tx('r-vacio'), cerrar: tx('cerrar'),
        filas: reglas.map((x, i) => ({
          texto: describir(x) + (x.editable ? '' : ' ' + tx('r-solo-excel')), detalle: x.sqref,
          acciones: x.editable ? [{ valor: 'editar', texto: tx('n-editar') }, { valor: 'subir', texto: tx('r-subir'), deshabilitado: i === 0 }, { valor: 'bajar', texto: tx('r-bajar'), deshabilitado: i === reglas.length - 1 }, { valor: 'quitar', texto: tx('n-quitar') }] : [],
        })),
        botones: [{ valor: 'nueva', texto: tx('r-nueva') }],
      }));
      if (!r) return;
      if (r.accion === 'nueva') return editarRegla(null);
      const x = reglas[r.indice];
      if (r.accion === 'editar') return editarRegla(x);
      if (r.accion === 'quitar') await A.estructural('Quitar regla', (b) => A.DAT().quitarRegla(b, h, x.id));
      if (r.accion === 'subir' || r.accion === 'bajar') { await A.estructural('Mover regla', (b) => A.DAT().moverRegla(b, h, x.id, r.accion === 'subir' ? -1 : 1)); return administrarReglas(); }
    }

    A.reg({ id: 'nombres', texto: 'cmd-nombres', atajo: 'mod+f3', ejecutar: administrarNombres });
    A.reg({ id: 'tabla', texto: 'cmd-tabla', atajo: 'mod+t', ejecutar: formatoTabla });
    A.reg({ id: 'validacion', texto: 'cmd-validacion', ejecutar: validacion });
    A.reg({ id: 'reglas', texto: 'cmd-reglas', ejecutar: administrarReglas });
    A.reg({ id: 'nuevaRegla', texto: 'cmd-nueva-regla', ejecutar: () => editarRegla(null) });
    A.cinta.push(['datos', { texto: 'g-datos-libro', items: [{ comando: 'tabla', icono: '▤', conTexto: true }, { comando: 'validacion', icono: '✓', conTexto: true }, { comando: 'reglas', icono: '◧', conTexto: true }, { comando: 'nombres', icono: '🏷', conTexto: true }] }]);
    A.cinta.push(['inicio', { texto: 'g-datos-libro', solo: 'pc', items: [{ comando: 'nuevaRegla', icono: '◧' }, { comando: 'tabla', icono: '▤' }] }]);
  });
})();

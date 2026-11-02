/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Datos del libro de Mini Cookies Hojas (fase 3.2 del PLAN): nombres definidos, dar formato como tabla, validación de datos y reglas de
// formato condicional. Como EstructuraLibro, cada operación toma los bytes de un .xlsx y devuelve los nuevos editando solo lo que cambia
// (xml-editar): el resto del paquete sale byte a byte. El editor las deshace volviendo al libro anterior. Sin DOM.
// Las fórmulas (de nombres, validaciones y reglas) van como las guarda Excel: en inglés, con comas.
//
//   DatosLibro.nombreValido('Ventas') → true                                  (sintaxis de Excel: no A1, no R1C1, sin espacios…)
//   await DatosLibro.ponerNombre(bytes, { nombre, ref, hoja: null | índice }, anterior?: { nombre, hoja })
//   await DatosLibro.quitarNombre(bytes, { nombre, hoja })
//   await DatosLibro.crearTabla(bytes, hoja, { f0, c0, f1, c1 }, { encabezados: true, estilo: 'TableStyleMedium2', nombreColumna: 'Columna', nombreBase: 'Tabla' })
//     → { bytes, nombre }
//   await DatosLibro.ponerValidacion(bytes, hoja, [rangos], regla | null)       (null quita la validación del rango)
//     regla = { tipo: 'list'|'whole'|'decimal'|'date'|'time'|'textLength'|'custom', operador, formula1, formula2, permitirVacio, ocultarFlecha,
//               estiloError: 'stop'|'warning'|'information', tituloError, error, mostrarEntrada, tituloEntrada, entrada }
//   await DatosLibro.reglasCondicionales(bytes, hoja) → [{ id, sqref, tipo, prioridad, regla, editable }]
//   await DatosLibro.agregarRegla(bytes, hoja, sqref, regla)                   (la nueva queda primera, como en Excel)
//   await DatosLibro.cambiarRegla(bytes, hoja, id, sqref, regla) · quitarRegla(bytes, hoja, id) · moverRegla(bytes, hoja, id, -1 | 1)
//     regla = { tipo: 'cellIs'|'containsText'|'notContainsText'|'beginsWith'|'endsWith'|'containsBlanks'|'top10'|'aboveAverage'|'duplicateValues'|
//               'uniqueValues'|'expression'|'colorScale'|'dataBar'|'iconSet', operador, formulas: [], texto, rango, porcentaje, abajo, encima,
//               colores: ['F8696B', …], iconos: '3Arrows', estilo: { relleno, color, negrita, cursiva } }
// Las reglas y validaciones de Excel 2010 (en extLst) se muestran pero no se editan: se conservan tal cual.
(function () {
  const X = () => globalThis.XmlEditar;
  const ErrorDatos = (m) => new XlsxEscritor.ErrorEscritor(m);
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const REL_TABLA = NS_R + '/table';
  const TIPO_TABLA = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';
  const ORDEN_HOJA = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
  const ORDEN_LIBRO = ['fileVersion', 'fileSharing', 'workbookPr', 'workbookProtection', 'bookViews', 'sheets', 'functionGroups', 'externalReferences', 'definedNames', 'calcPr', 'oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'];
  const ORDEN_ESTILOS = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'];
  const col = (n) => LibroXlsx.numeroACol(n);
  const refDe = (r) => col(r.c0) + (r.f0 + 1) + (r.f0 === r.f1 && r.c0 === r.c1 ? '' : ':' + col(r.c1) + (r.f1 + 1));
  const texto = async (zip, ruta) => { const e = zip.file(ruta); return e ? e.async('string') : null; };

  async function conLibro(bytes, fn) {
    const libro = await LibroXlsx.abrir(bytes);
    const r = await fn(libro);
    const salida = await libro.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    return r && typeof r === 'object' ? Object.assign({}, r, { bytes: salida }) : salida;
  }
  const rutaHoja = (libro, h) => { const m = libro.hojas[h]; if (!m || m.tipo !== 'hoja' || !m.ruta) throw ErrorDatos('la hoja no es de celdas'); return m.ruta; };
  // pone `xml` como hijo de la raíz en su lugar según el orden de la norma (después de los del mismo nombre, si ya hay)
  function insertarEnOrden(doc, orden, nombre, xml, profundidad) {
    const raiz = X().analizar(doc, { profundidad: profundidad || 2 });
    const iguales = raiz.hijos.filter((h) => h.local === nombre);
    let pos;
    if (iguales.length) pos = iguales[iguales.length - 1].fin;
    else { const k = orden.indexOf(nombre); const sig = raiz.hijos.find((h) => orden.indexOf(h.local) > k); pos = sig ? sig.ini : raiz.iniCierre; }
    if (pos < 0) throw ErrorDatos('documento sin cierre');
    return X().aplicar(doc, [{ ini: pos, fin: pos, texto: xml }]);
  }

  // ---- nombres definidos ---------------------------------------------------------------------------------------------
  function nombreValido(nombre) {
    const n = String(nombre || '');
    if (!n || n.length > 255 || !/^[A-Za-z_\\À-￿][A-Za-z0-9_.\\?À-￿]*$/.test(n)) return false;
    if (/^[A-Za-z]{1,3}\d+$/.test(n) && LibroXlsx.colANumero(n.replace(/\d+$/, '').toUpperCase()) <= 16383 && +n.replace(/^[A-Za-z]+/, '') <= 1048576) return false;   // parece una celda
    if (/^([Rr]\d*)?([Cc]\d*)?$/.test(n)) return false;   // R, C, RC, R1C1… (referencias del estilo R1C1)
    if (/^(TRUE|FALSE)$/i.test(n)) return false;
    return true;
  }
  async function ponerNombre(bytes, nuevo, anterior) {
    if (!nombreValido(nuevo.nombre)) throw ErrorDatos('nombre no válido: ' + nuevo.nombre);
    const ref = String(nuevo.ref || '').replace(/^=/, '');
    if (!ref) throw ErrorDatos('el nombre necesita una referencia o fórmula');
    try { XlsFormulas._analizar(XlsFormulas._tokenizar(ref)); } catch (e) { throw ErrorDatos('la referencia del nombre no es una fórmula válida'); }
    return conLibro(bytes, async (libro) => {
      const hoja = Number.isInteger(nuevo.hoja) ? nuevo.hoja : null;
      if (hoja !== null && !libro.hojas[hoja]) throw ErrorDatos('hoja inexistente');
      const mismo = (x, nombre, h) => x.nombre.toLowerCase() === String(nombre).toLowerCase() && (x.hojaLocal === undefined ? null : x.hojaLocal) === h;
      const esAnterior = (x) => anterior && mismo(x, anterior.nombre, Number.isInteger(anterior.hoja) ? anterior.hoja : null);
      if (libro.nombres.some((x) => mismo(x, nuevo.nombre, hoja) && !esAnterior(x))) throw ErrorDatos('ya hay un nombre «' + nuevo.nombre + '» con ese ámbito');
      // un nombre no puede llamarse como una tabla
      for (let i = 0; i < libro.hojas.length; i++) { const m = libro.hojas[i].tipo === 'hoja' ? await libro.cargarHoja(i) : null; if (m && (m.tablas || []).some((t) => t.nombre.toLowerCase() === nuevo.nombre.toLowerCase())) throw ErrorDatos('ya hay una tabla «' + nuevo.nombre + '»'); }
      let wb = await texto(libro.zip, libro.rutaLibro);
      const raiz = X().analizar(wb, { profundidad: 3 }), dn = X().hijo(raiz, 'definedNames');
      const local = hoja !== null ? ` localSheetId="${hoja}"` : '';
      const elemento = (abrir) => abrir + X().esc(ref) + '</definedName>';
      if (anterior) {
        const lista = dn ? X().hijos(dn, 'definedName') : [];
        const h0 = Number.isInteger(anterior.hoja) ? String(anterior.hoja) : undefined;
        const x = lista.find((d) => (d.attrs.name || '').toLowerCase() === anterior.nombre.toLowerCase() && d.attrs.localSheetId === h0);
        if (!x) throw ErrorDatos('no está el nombre ' + anterior.nombre);
        // se conservan sus otros atributos (comment, hidden…): cambian nombre, ámbito y contenido
        let abrir = wb.slice(x.ini, x.finApertura).replace(/\s*\/>$/, '>');
        abrir = X().ponerAtributo(X().ponerAtributo(abrir, 'name', nuevo.nombre), 'localSheetId', hoja !== null ? String(hoja) : null);
        wb = X().aplicar(wb, [{ ini: x.ini, fin: x.fin, texto: elemento(abrir) }]);
      } else {
        const nodo = elemento(`<definedName name="${X().escAttr(nuevo.nombre)}"${local}>`);
        if (dn && dn.iniCierre >= 0) wb = X().aplicar(wb, [{ ini: dn.iniCierre, fin: dn.iniCierre, texto: nodo }]);
        else if (dn) wb = X().aplicar(wb, [{ ini: dn.ini, fin: dn.fin, texto: '<definedNames>' + nodo + '</definedNames>' }]);
        else wb = insertarEnOrden(wb, ORDEN_LIBRO, 'definedNames', '<definedNames>' + nodo + '</definedNames>');
      }
      libro.zip.file(libro.rutaLibro, wb);
      await XlsxPaquete.marcarRecalculo(libro.zip);
    });
  }
  async function quitarNombre(bytes, q) {
    return conLibro(bytes, async (libro) => {
      let wb = await texto(libro.zip, libro.rutaLibro);
      const raiz = X().analizar(wb, { profundidad: 3 }), dn = X().hijo(raiz, 'definedNames');
      const h0 = Number.isInteger(q.hoja) ? String(q.hoja) : undefined;
      const lista = dn ? X().hijos(dn, 'definedName') : [];
      const x = lista.find((d) => (d.attrs.name || '').toLowerCase() === String(q.nombre).toLowerCase() && d.attrs.localSheetId === h0);
      if (!x) throw ErrorDatos('no está el nombre ' + q.nombre);
      wb = X().aplicar(wb, [lista.length === 1 ? { ini: dn.ini, fin: dn.fin, texto: '' } : { ini: x.ini, fin: x.fin, texto: '' }]);
      libro.zip.file(libro.rutaLibro, wb);
      await XlsxPaquete.marcarRecalculo(libro.zip);
    });
  }

  // ---- dar formato como tabla --------------------------------------------------------------------------------------------
  const solapa = (a, b) => a.f0 <= b.f1 && a.f1 >= b.f0 && a.c0 <= b.c1 && a.c1 >= b.c0;
  async function crearTabla(bytes, h, r, opciones) {
    const o = Object.assign({ encabezados: true, estilo: 'TableStyleMedium2', nombreColumna: 'Columna', nombreBase: 'Tabla' }, opciones || {});
    if (r.f1 - r.f0 < (o.encabezados ? 1 : 0) || r.c1 - r.c0 > 16383) throw ErrorDatos('una tabla con encabezados necesita al menos dos filas');
    return conLibro(bytes, async (libro) => {
      const ruta = rutaHoja(libro, h), zip = libro.zip, P = PaqueteEditable.desdeZip(zip);
      const hoja = await libro.cargarHoja(h);
      if ((hoja.tablas || []).some((t) => solapa(t.ref, r))) throw ErrorDatos('el rango se superpone con otra tabla');
      if ((hoja.combinadas || []).some((m) => solapa(m, r))) throw ErrorDatos('el rango tiene celdas combinadas');
      let xml = await texto(zip, ruta);
      const af = /<autoFilter\b[^>]*\bref="([^"]+)"/.exec(xml);
      if (af && LibroXlsx.decodificarRango(af[1]) && solapa(LibroXlsx.decodificarRango(af[1]), r)) throw ErrorDatos('el rango tiene un autofiltro: quítalo primero');
      for (const fx of hoja.celdas.formulas.values()) if (fx.tipo === 'array' && fx.ref && !fx.dinamica) { const g = LibroXlsx.decodificarRango(fx.ref); if (g && solapa(g, r)) throw ErrorDatos('el rango tiene una fórmula matricial'); }
      // nombres e id: únicos en todo el libro (también frente a los nombres definidos)
      const usados = new Set(libro.nombres.map((x) => x.nombre.toLowerCase()));
      let maxId = 0;
      const partes = [];
      zip.forEach((p) => { if (/^xl\/tables\/[^/]+\.xml$/i.test(p)) partes.push(p); });
      for (const p of partes) { const t = await texto(zip, p); const a = /<table\b[^>]*>/.exec(t || ''); if (!a) continue; const at = X().analizar(a[0] + '</table>').attrs; usados.add(String(at.name || '').toLowerCase()); usados.add(String(at.displayName || '').toLowerCase()); maxId = Math.max(maxId, +at.id || 0); }
      let k = 1; while (usados.has((o.nombreBase + k).toLowerCase())) k++;
      const nombre = o.nombreTabla && nombreValido(o.nombreTabla) && !usados.has(o.nombreTabla.toLowerCase()) ? o.nombreTabla : o.nombreBase + k;
      // encabezados: textos únicos (los vacíos se llaman Columna1…; los números pasan a texto, como en Excel)
      const nombres = [], cambios = [], vistos = new Set();
      for (let c = r.c0; c <= r.c1; c++) {
        const x = o.encabezados ? hoja.celdas.obtener(r.f0, c) : null;
        let t = x && x.valor !== null && x.valor !== undefined && x.tipo !== LibroXlsx.TIPO.vacio ? (typeof x.valor === 'number' ? String(x.valor) : String(x.valor)) : '';
        t = t.replace(/[\r\n]+/g, ' ').trim();
        const base = t || o.nombreColumna + (c - r.c0 + 1);
        let n = base, j = 2;
        while (vistos.has(n.toLowerCase())) n = base + j++;
        vistos.add(n.toLowerCase());
        nombres.push(n);
        if (o.encabezados && (!x || x.tipo !== LibroXlsx.TIPO.texto || x.valor !== n || hoja.celdas.formulas.has(hoja.celdas.indiceDe(r.f0, c)))) cambios.push({ f: r.f0, c, valor: n, estilo: x ? x.estilo : undefined });
      }
      if (cambios.length) { await XlsxEscritor.editarCeldas(libro, h, cambios); xml = await texto(zip, ruta); }
      // la parte de la tabla
      let n = partes.length + 1; while (zip.file('xl/tables/table' + n + '.xml')) n++;
      const rutaTabla = 'xl/tables/table' + n + '.xml';
      const refTabla = refDe(r);
      const columnas = nombres.map((x, i) => `<tableColumn id="${i + 1}" name="${X().escAttr(x)}"/>`).join('');
      const tabla = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${maxId + 1}" name="${nombre}" displayName="${nombre}" ref="${refTabla}"`
        + (o.encabezados ? '' : ' headerRowCount="0"') + ' totalsRowShown="0">' + (o.encabezados ? `<autoFilter ref="${refTabla}"/>` : '')
        + `<tableColumns count="${nombres.length}">${columnas}</tableColumns><tableStyleInfo name="${X().escAttr(o.estilo)}" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`;
      P.escribir(rutaTabla, tabla);
      await P.asegurarTipo('/' + rutaTabla, TIPO_TABLA);
      const relativo = '../tables/table' + n + '.xml';
      const id = await P.agregarRelacion(ruta, REL_TABLA, relativo);
      // <tableParts> de la hoja (con el prefijo r declarado)
      const raiz = X().analizar(xml, { profundidad: 2 });
      const prefijo = Object.keys(raiz.attrs).find((a) => a.startsWith('xmlns:') && raiz.attrs[a] === NS_R);
      const pr = prefijo ? prefijo.slice(6) : 'r';
      if (!prefijo) xml = X().aplicar(xml, [{ ini: raiz.ini, fin: raiz.finApertura, texto: X().ponerAtributo(xml.slice(raiz.ini, raiz.finApertura), 'xmlns:r', NS_R) }]);
      const parte = `<tablePart ${pr}:id="${id}"/>`;
      const r2 = X().analizar(xml, { profundidad: 2 }), tp = X().hijo(r2, 'tableParts');
      if (tp && tp.iniCierre >= 0) {
        const cuenta = X().hijos(tp, 'tablePart').length + 1;
        xml = X().aplicar(xml, [{ ini: tp.ini, fin: tp.finApertura, texto: X().ponerAtributo(xml.slice(tp.ini, tp.finApertura), 'count', String(cuenta)) }, { ini: tp.iniCierre, fin: tp.iniCierre, texto: parte }]);
      } else if (tp) xml = X().aplicar(xml, [{ ini: tp.ini, fin: tp.fin, texto: `<tableParts count="1">${parte}</tableParts>` }]);
      else xml = insertarEnOrden(xml, ORDEN_HOJA, 'tableParts', `<tableParts count="1">${parte}</tableParts>`);
      zip.file(ruta, xml);
      if (libro._hojasCargadas) libro._hojasCargadas.clear();
      return { nombre };
    });
  }

  // ---- validación de datos -----------------------------------------------------------------------------------------------
  // a menos b (rectángulos) → hasta cuatro rectángulos
  function restar(a, b) {
    if (!solapa(a, b)) return [a];
    const r = [];
    if (a.f0 < b.f0) r.push({ f0: a.f0, c0: a.c0, f1: b.f0 - 1, c1: a.c1 });
    if (a.f1 > b.f1) r.push({ f0: b.f1 + 1, c0: a.c0, f1: a.f1, c1: a.c1 });
    const f0 = Math.max(a.f0, b.f0), f1 = Math.min(a.f1, b.f1);
    if (a.c0 < b.c0) r.push({ f0, c0: a.c0, f1, c1: b.c0 - 1 });
    if (a.c1 > b.c1) r.push({ f0, c0: b.c1 + 1, f1, c1: a.c1 });
    return r;
  }
  const TIPOS_VAL = new Set(['list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom']);
  const OPERADORES = new Set(['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual']);
  function xmlValidacion(rangos, v) {
    if (!TIPOS_VAL.has(v.tipo)) throw ErrorDatos('tipo de validación desconocido');
    const op = v.tipo === 'list' || v.tipo === 'custom' ? null : OPERADORES.has(v.operador) ? v.operador : 'between';
    const at = [['type', v.tipo], op && op !== 'between' ? ['operator', op] : null, v.permitirVacio !== false ? ['allowBlank', '1'] : null, v.ocultarFlecha ? ['showDropDown', '1'] : null,
      v.mostrarEntrada !== false ? ['showInputMessage', '1'] : null, v.mostrarError !== false ? ['showErrorMessage', '1'] : null,
      v.estiloError && v.estiloError !== 'stop' ? ['errorStyle', v.estiloError] : null, v.tituloError ? ['errorTitle', v.tituloError] : null, v.error ? ['error', v.error] : null,
      v.tituloEntrada ? ['promptTitle', v.tituloEntrada] : null, v.entrada ? ['prompt', v.entrada] : null, ['sqref', rangos.map(refDe).join(' ')]].filter(Boolean);
    const f1 = String(v.formula1 === undefined || v.formula1 === null ? '' : v.formula1).replace(/^=/, '');
    const f2 = String(v.formula2 === undefined || v.formula2 === null ? '' : v.formula2).replace(/^=/, '');
    if (!f1) throw ErrorDatos('la validación necesita un valor o una fórmula');
    if ((op === 'between' || op === 'notBetween') && !f2) throw ErrorDatos('«entre» necesita dos valores');
    return '<dataValidation ' + at.map(([k, x]) => `${k}="${X().escAttr(String(x))}"`).join(' ') + '><formula1>' + X().esc(f1) + '</formula1>' + ((op === 'between' || op === 'notBetween') ? '<formula2>' + X().esc(f2) + '</formula2>' : '') + '</dataValidation>';
  }
  async function ponerValidacion(bytes, h, rangos, regla) {
    if (!rangos || !rangos.length) throw ErrorDatos('sin rango');
    return conLibro(bytes, async (libro) => {
      const ruta = rutaHoja(libro, h);
      let xml = await texto(libro.zip, ruta);
      const raiz = X().analizar(xml, { profundidad: 3 }), dvs = X().hijo(raiz, 'dataValidations');
      const ed = [];
      let quedan = 0;
      // las validaciones que ya cubrían parte del rango pierden esa parte (como Excel: una celda tiene una sola validación)
      for (const dv of dvs ? X().hijos(dvs, 'dataValidation') : []) {
        const antes = String(dv.attrs.sqref || '').split(/\s+/).filter(Boolean).map((x) => LibroXlsx.decodificarRango(x)).filter(Boolean);
        let despues = antes;
        for (const b of rangos) despues = [].concat(...despues.map((a) => restar(a, b)));
        if (despues.length === antes.length && despues.every((x, i) => x === antes[i])) { quedan++; continue; }
        if (!despues.length) { ed.push({ ini: dv.ini, fin: dv.fin, texto: '' }); continue; }
        quedan++;
        let et = X().ponerAtributo(xml.slice(dv.ini, dv.finApertura), 'sqref', despues.map(refDe).join(' '));
        // sus fórmulas son relativas a la primera celda: si esa celda cambió, se corren para que sigan diciendo lo mismo
        const a0 = antes[0], b0 = despues[0];
        let cuerpo = dv.iniCierre >= 0 ? xml.slice(dv.finApertura, dv.iniCierre) : '';
        if (a0.f0 !== b0.f0 || a0.c0 !== b0.c0) cuerpo = cuerpo.replace(/(<formula[12]>)([\s\S]*?)(<\/formula[12]>)/g, (m, x, f, y) => x + X().esc(LibroXlsx.desplazarFormula(X().decodificar(f), b0.f0 - a0.f0, b0.c0 - a0.c0)) + y);
        ed.push({ ini: dv.ini, fin: dv.fin, texto: dv.iniCierre >= 0 ? et + cuerpo + '</dataValidation>' : et });
      }
      const nueva = regla ? xmlValidacion(rangos, regla) : '';
      if (regla) quedan++;
      if (dvs) {
        if (!quedan) ed.push({ ini: dvs.ini, fin: dvs.fin, texto: '' });
        else {
          ed.push({ ini: dvs.ini, fin: dvs.finApertura, texto: X().ponerAtributo(xml.slice(dvs.ini, dvs.finApertura), 'count', String(quedan)) });
          if (nueva) ed.push({ ini: dvs.iniCierre, fin: dvs.iniCierre, texto: nueva });
        }
        xml = X().aplicar(xml, ed);
      } else if (nueva) xml = insertarEnOrden(xml, ORDEN_HOJA, 'dataValidations', `<dataValidations count="1">${nueva}</dataValidations>`);
      libro.zip.file(ruta, xml);
    });
  }

  // ---- formato condicional -----------------------------------------------------------------------------------------------
  const rgb = (c) => { const x = String(c || '').replace(/^#/, '').toUpperCase(); if (!/^[0-9A-F]{6}$/.test(x)) throw ErrorDatos('color no válido: ' + c); return 'FF' + x; };
  function xmlDxf(e) {
    const s = e || {};
    let fuente = '';
    if (s.negrita) fuente += '<b/>';
    if (s.cursiva) fuente += '<i/>';
    if (s.color) fuente += `<color rgb="${rgb(s.color)}"/>`;
    return '<dxf>' + (fuente ? '<font>' + fuente + '</font>' : '') + (s.relleno ? `<fill><patternFill><bgColor rgb="${rgb(s.relleno)}"/></patternFill></fill>` : '') + '</dxf>';
  }
  // agrega el dxf a styles.xml y devuelve su índice
  async function agregarDxf(libro, estilo) {
    const rel = (libro.relaciones || []).find((r) => /(^|\/)styles$/.test(r.tipo) && !r.externa);
    const ruta = rel ? rel.destino : 'xl/styles.xml';
    let xml = await texto(libro.zip, ruta);
    if (xml === null) throw ErrorDatos('el libro no tiene estilos');
    const raiz = X().analizar(xml, { profundidad: 2 }), dx = X().hijo(raiz, 'dxfs');
    const nuevo = xmlDxf(estilo);
    let i;
    if (dx && dx.iniCierre >= 0) {
      i = X().hijos(X().analizar(xml, { profundidad: 3 }).hijos.find((h) => h.local === 'dxfs'), 'dxf').length;
      xml = X().aplicar(xml, [{ ini: dx.ini, fin: dx.finApertura, texto: X().ponerAtributo(xml.slice(dx.ini, dx.finApertura), 'count', String(i + 1)) }, { ini: dx.iniCierre, fin: dx.iniCierre, texto: nuevo }]);
    } else if (dx) { i = 0; xml = X().aplicar(xml, [{ ini: dx.ini, fin: dx.fin, texto: `<dxfs count="1">${nuevo}</dxfs>` }]); }
    else { i = 0; xml = insertarEnOrden(xml, ORDEN_ESTILOS, 'dxfs', `<dxfs count="1">${nuevo}</dxfs>`); }
    libro.zip.file(ruta, xml);
    return i;
  }
  const TIPOS_CF = new Set(['cellIs', 'containsText', 'notContainsText', 'beginsWith', 'endsWith', 'containsBlanks', 'notContainsBlanks', 'top10', 'aboveAverage', 'duplicateValues', 'uniqueValues', 'expression', 'colorScale', 'dataBar', 'iconSet']);
  const OPERADORES_CF = new Set(['lessThan', 'lessThanOrEqual', 'equal', 'notEqual', 'greaterThanOrEqual', 'greaterThan', 'between', 'notBetween']);
  const ICONOS = new Set(['3Arrows', '3ArrowsGray', '3Flags', '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2', '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights', '5Arrows', '5ArrowsGray', '5Rating', '5Quarters']);
  // texto de cfRule (sin dxfId cuando la regla no lleva formato: escalas, barras, iconos). `celda`: la primera del rango (las fórmulas de texto la citan)
  function xmlRegla(regla, prioridad, dxfId, celda) {
    const t = regla.tipo;
    if (!TIPOS_CF.has(t)) throw ErrorDatos('tipo de regla desconocido');
    const dx = dxfId === null || dxfId === undefined ? '' : ` dxfId="${dxfId}"`;
    const f = (x) => '<formula>' + X().esc(String(x).replace(/^=/, '')) + '</formula>';
    const lit = (s) => '"' + String(s).replace(/"/g, '""') + '"';
    switch (t) {
      case 'cellIs': {
        const op = OPERADORES_CF.has(regla.operador) ? regla.operador : 'greaterThan';
        const fs = (regla.formulas || []).filter((x) => x !== '' && x !== null && x !== undefined);
        if (!fs.length || ((op === 'between' || op === 'notBetween') && fs.length < 2)) throw ErrorDatos('la regla necesita su valor');
        return `<cfRule type="cellIs"${dx} priority="${prioridad}" operator="${op}">${fs.slice(0, op === 'between' || op === 'notBetween' ? 2 : 1).map(f).join('')}</cfRule>`;
      }
      case 'containsText': case 'notContainsText': case 'beginsWith': case 'endsWith': {
        const s = String(regla.texto || '');
        if (!s) throw ErrorDatos('la regla necesita un texto');
        const op = { containsText: 'containsText', notContainsText: 'notContains', beginsWith: 'beginsWith', endsWith: 'endsWith' }[t];
        const fo = { containsText: `NOT(ISERROR(SEARCH(${lit(s)},${celda})))`, notContainsText: `ISERROR(SEARCH(${lit(s)},${celda}))`, beginsWith: `LEFT(${celda},LEN(${lit(s)}))=${lit(s)}`, endsWith: `RIGHT(${celda},LEN(${lit(s)}))=${lit(s)}` }[t];
        return `<cfRule type="${t}"${dx} priority="${prioridad}" operator="${op}" text="${X().escAttr(s)}">${f(fo)}</cfRule>`;
      }
      case 'containsBlanks': return `<cfRule type="containsBlanks"${dx} priority="${prioridad}">${f(`LEN(TRIM(${celda}))=0`)}</cfRule>`;
      case 'notContainsBlanks': return `<cfRule type="notContainsBlanks"${dx} priority="${prioridad}">${f(`LEN(TRIM(${celda}))>0`)}</cfRule>`;
      case 'top10': {
        const n = Math.trunc(+regla.rango || 10);
        if (!(n >= 1 && n <= 1000)) throw ErrorDatos('el rango de «superiores» va de 1 a 1000');
        return `<cfRule type="top10"${dx} priority="${prioridad}"${regla.porcentaje ? ' percent="1"' : ''}${regla.abajo ? ' bottom="1"' : ''} rank="${n}"/>`;
      }
      case 'aboveAverage': return `<cfRule type="aboveAverage"${dx} priority="${prioridad}"${regla.encima === false ? ' aboveAverage="0"' : ''}/>`;
      case 'duplicateValues': case 'uniqueValues': return `<cfRule type="${t}"${dx} priority="${prioridad}"/>`;
      case 'expression': {
        const fs = (regla.formulas || []).filter(Boolean);
        if (!fs.length) throw ErrorDatos('la regla necesita una fórmula');
        return `<cfRule type="expression"${dx} priority="${prioridad}">${f(fs[0])}</cfRule>`;
      }
      case 'colorScale': {
        const cs = (regla.colores && regla.colores.length ? regla.colores : ['F8696B', 'FFEB84', '63BE7B']).slice(0, 3);
        if (cs.length < 2) throw ErrorDatos('una escala necesita dos o tres colores');
        const cfvo = cs.length === 3 ? '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>' : '<cfvo type="min"/><cfvo type="max"/>';
        return `<cfRule type="colorScale" priority="${prioridad}"><colorScale>${cfvo}${cs.map((c) => `<color rgb="${rgb(c)}"/>`).join('')}</colorScale></cfRule>`;
      }
      case 'dataBar': return `<cfRule type="dataBar" priority="${prioridad}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="${rgb((regla.colores || [])[0] || '638EC6')}"/></dataBar></cfRule>`;
      case 'iconSet': {
        const nombre = ICONOS.has(regla.iconos) ? regla.iconos : '3TrafficLights1';
        const n = +nombre[0];
        const cfvo = Array.from({ length: n }, (_, i) => `<cfvo type="percent" val="${Math.round(100 * i / n)}"/>`).join('');
        return `<cfRule type="iconSet" priority="${prioridad}"><iconSet${nombre === '3TrafficLights1' ? '' : ` iconSet="${nombre}"`}>${cfvo}</iconSet></cfRule>`;
      }
      default: throw ErrorDatos('tipo de regla desconocido');
    }
  }
  const sinDxf = (t) => t === 'colorScale' || t === 'dataBar' || t === 'iconSet';
  // las reglas de la hoja en el orden del archivo: las de <conditionalFormatting> (editables) y las de Excel 2010 en extLst (solo lectura)
  function leerReglas(xml) {
    const raiz = X().analizar(xml, { profundidad: 3 });
    const lista = [];
    raiz.hijos.filter((h) => h.local === 'conditionalFormatting').forEach((cf, k) => {
      X().hijos(cf, 'cfRule').forEach((r, j) => {
        const a = r.attrs, interior = r.iniCierre >= 0 ? xml.slice(r.finApertura, r.iniCierre) : '';
        const formulas = [...interior.matchAll(/<formula>([\s\S]*?)<\/formula>/g)].map((m) => X().decodificar(m[1]));
        const colores = [...interior.matchAll(/<color\b[^>]*\brgb="(?:[0-9A-Fa-f]{2})?([0-9A-Fa-f]{6})"/g)].map((m) => m[1].toUpperCase());
        lista.push({
          id: k + ':' + j, nodo: r, cf, sqref: cf.attrs.sqref || '', tipo: a.type, prioridad: +a.priority || 0, editable: TIPOS_CF.has(a.type),
          regla: { tipo: a.type, operador: a.operator || null, formulas, texto: a.text ? X().decodificar(a.text) : null, rango: a.rank ? +a.rank : null, porcentaje: a.percent === '1', abajo: a.bottom === '1', encima: a.aboveAverage !== '0', colores, iconos: (/<iconSet\b[^>]*\biconSet="([^"]+)"/.exec(interior) || [])[1] || (a.type === 'iconSet' ? '3TrafficLights1' : null), dxfId: a.dxfId !== undefined ? +a.dxfId : null },
        });
      });
    });
    for (const m of xml.matchAll(/<x14:conditionalFormatting\b[\s\S]*?<\/x14:conditionalFormatting>/g)) {
      const sq = (/<xm:sqref>([\s\S]*?)<\/xm:sqref>/.exec(m[0]) || [])[1] || '';
      for (const r of m[0].matchAll(/<x14:cfRule\b([^>]*)>/g)) {
        const a = X().analizar('<r ' + r[1].replace(/\/$/, '') + '/>').attrs;
        lista.push({ id: 'x14:' + m.index + ':' + r.index, sqref: sq, tipo: a.type, prioridad: +a.priority || 0, editable: false, regla: { tipo: a.type } });
      }
    }
    return lista;
  }
  async function reglasCondicionales(bytes, h) {
    const libro = await LibroXlsx.abrir(bytes);
    const xml = await texto(libro.zip, rutaHoja(libro, h));
    return leerReglas(xml).map((x) => ({ id: x.id, sqref: x.sqref, tipo: x.tipo, prioridad: x.prioridad, editable: x.editable, regla: x.regla })).sort((a, b) => a.prioridad - b.prioridad);
  }
  const corrPrioridades = (xml, fn) => xml.replace(/(<(?:x14:)?cfRule\b[^>]*?\bpriority=")(\d+)(")/g, (m, a, n, c) => a + fn(+n) + c);
  function sqrefValido(sqref) {
    const partes = String(sqref || '').trim().split(/\s+/).filter(Boolean);
    const rs = partes.map((x) => LibroXlsx.decodificarRango(x.replace(/\$/g, '')));
    if (!partes.length || rs.some((x) => !x)) throw ErrorDatos('rango no válido: ' + sqref);
    return { texto: rs.map(refDe).join(' '), primera: col(rs[0].c0) + (rs[0].f0 + 1) };
  }
  async function agregarRegla(bytes, h, sqref, regla, prioridad) {
    return conLibro(bytes, async (libro) => {
      const ruta = rutaHoja(libro, h);
      let xml = await texto(libro.zip, ruta);
      const sq = sqrefValido(sqref);
      const dxf = sinDxf(regla.tipo) ? null : await agregarDxf(libro, regla.estilo || { relleno: 'FFC7CE', color: '9C0006' });
      let p = prioridad;
      if (!p) { xml = corrPrioridades(xml, (n) => n + 1); p = 1; }   // la nueva va primero; las demás bajan un lugar
      const nodo = `<conditionalFormatting sqref="${sq.texto}">${xmlRegla(regla, p, dxf, sq.primera)}</conditionalFormatting>`;
      xml = insertarEnOrden(xml, ORDEN_HOJA, 'conditionalFormatting', nodo);
      libro.zip.file(ruta, xml);
    });
  }
  // quita la regla (y su <conditionalFormatting> si queda vacío); devuelve el xml y la prioridad que tenía
  function sacar(xml, id) {
    const r = leerReglas(xml).find((x) => x.id === id);
    if (!r || !r.editable) throw ErrorDatos('esa regla no se puede editar aquí');
    const otras = X().hijos(r.cf, 'cfRule').length;
    return { xml: X().aplicar(xml, [otras === 1 ? { ini: r.cf.ini, fin: r.cf.fin, texto: '' } : { ini: r.nodo.ini, fin: r.nodo.fin, texto: '' }]), prioridad: r.prioridad };
  }
  async function quitarRegla(bytes, h, id) {
    return conLibro(bytes, async (libro) => {
      const ruta = rutaHoja(libro, h);
      const { xml } = sacar(await texto(libro.zip, ruta), id);
      libro.zip.file(ruta, xml);
    });
  }
  async function cambiarRegla(bytes, h, id, sqref, regla) {
    const libro = await LibroXlsx.abrir(bytes);
    const ruta = rutaHoja(libro, h);
    const { xml, prioridad } = sacar(await texto(libro.zip, ruta), id);
    libro.zip.file(ruta, xml);
    const intermedio = await libro.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    return agregarRegla(intermedio, h, sqref, regla, prioridad);   // en el mismo lugar del orden
  }
  // sube (-1) o baja (+1) la regla un lugar en el orden de prioridad: intercambia su prioridad con la vecina
  async function moverRegla(bytes, h, id, delta) {
    return conLibro(bytes, async (libro) => {
      const ruta = rutaHoja(libro, h);
      let xml = await texto(libro.zip, ruta);
      const lista = leerReglas(xml).sort((a, b) => a.prioridad - b.prioridad);
      const k = lista.findIndex((x) => x.id === id), j = k + delta;
      if (k < 0 || j < 0 || j >= lista.length) return;
      const a = lista[k].prioridad, b = lista[j].prioridad;
      if (a === b) return;
      xml = corrPrioridades(xml, (n) => (n === a ? b : n === b ? a : n));
      libro.zip.file(ruta, xml);
    });
  }

  // ---- validación al escribir (con el libro abierto en el editor) ----------------------------------------------------------
  // la validación que cubre la celda (la última de la lista, como Excel cuando hay dos) o null
  function validacionEn(ed, h, f, c) {
    const hoja = ed.hojas[h];
    if (!hoja || !hoja.validaciones) return null;
    let v = null;
    for (const x of hoja.validaciones) if (x.tipo !== 'none' && x.rangos.some((r) => f >= r.f0 && f <= r.f1 && c >= r.c0 && c <= r.c1)) v = x;
    return v;
  }
  // una fórmula de la validación, evaluada en la celda: sus referencias relativas se cuentan desde la primera celda del primer rango
  function evaluarEn(ed, v, formula, h, f, c) {
    if (formula === null || formula === undefined || formula === '') return null;
    const r0 = v.rangos[0];
    const t = LibroXlsx.desplazarFormula(String(formula).replace(/^=/, ''), f - r0.f0, c - r0.c0);
    ed.asegurarMotor();
    return ed.motor.evaluar(t, { h, f, c });
  }
  // las opciones de una lista: "a,b,c" (en el archivo siempre con comas) o los valores de un rango o nombre
  function listaDe(ed, h, f, c) {
    const v = validacionEn(ed, h, f, c);
    if (!v || v.tipo !== 'list' || !v.formula1) return null;
    const f1 = v.formula1.trim();
    if (/^".*"$/.test(f1)) return f1.slice(1, -1).replace(/""/g, '"').split(',').map((x) => x.trim()).filter((x) => x !== '');
    const r0 = v.rangos[0];
    const texto = LibroXlsx.desplazarFormula(f1.replace(/^=/, ''), f - r0.f0, c - r0.c0);
    // el rango de la lista: sus valores en orden (las celdas vacías no salen)
    const analizado = (() => { try { return XlsFormulas._analizar(XlsFormulas._tokenizar(texto)); } catch (e) { return null; } })();
    if (!analizado) return [];
    const valores = [];
    const libro = ed.libro;
    const hojaDe = (n) => (n ? libro.hojas.findIndex((x) => x.nombre.toLowerCase() === n.toLowerCase()) : h);
    let nodo = analizado;
    if (nodo.t === 'nombre') { const def = libro.nombres.find((x) => x.nombre.toLowerCase() === String(nodo.nombre).toLowerCase() && (x.hojaLocal === null || x.hojaLocal === h)); if (!def) return []; try { nodo = XlsFormulas._analizar(XlsFormulas._tokenizar(def.ref.replace(/^=/, ''))); } catch (e) { return []; } }
    if (nodo.t !== 'ref') return [];
    const hh = hojaDe(nodo.hoja && nodo.hoja.a);
    if (hh < 0) return [];
    const f0 = Math.min(nodo.f0, nodo.modo === 'area' ? nodo.f1 : nodo.f0), f1b = Math.max(nodo.f0, nodo.modo === 'area' ? nodo.f1 : nodo.f0);
    const c0 = Math.min(nodo.c0, nodo.modo === 'area' ? nodo.c1 : nodo.c0), c1 = Math.max(nodo.c0, nodo.modo === 'area' ? nodo.c1 : nodo.c0);
    for (let ff = f0; ff <= Math.min(f1b, f0 + 10000); ff++) for (let cc = c0; cc <= c1; cc++) {
      const x = ed.celda(hh, ff, cc);
      if (x && x.valor !== null && x.valor !== '' && x.tipo !== 'vacio') valores.push(typeof x.valor === 'number' ? CalculoHojas.formatoGeneral(x.valor) : String(x.valor));
    }
    return valores;
  }
  // ¿se acepta lo escrito? → null si sí; si no, { estilo: 'stop'|'warning'|'information', titulo, mensaje }. Las fórmulas no se comprueban aquí
  // (Excel comprueba su resultado); la personalizada se comprueba después de escribir con comprobarDespues.
  function comprobarEntrada(ed, h, f, c, texto) {
    const v = validacionEn(ed, h, f, c);
    if (!v || !v.mostrarError) return null;
    const x = EntradaCelda.interpretar(texto, ed.opciones);
    if (x.tipo === 'vacio') return null;   // borrar siempre se puede (como Excel)
    if (x.tipo === 'formula' || v.tipo === 'custom') return null;
    const malo = { estilo: v.estiloError || 'stop', titulo: v.tituloError || '', mensaje: v.error || '' };
    if (v.tipo === 'list') {
      const ops = listaDe(ed, h, f, c) || [];
      const t = x.tipo === 'numero' ? CalculoHojas.formatoGeneral(x.valor) : String(x.valor);
      return ops.some((o) => o.toLowerCase() === t.toLowerCase() || (x.tipo === 'numero' && EntradaCelda.interpretar(o, { idioma: 'en' }).valor === x.valor)) ? null : malo;
    }
    let valor;
    if (v.tipo === 'textLength') valor = (x.tipo === 'numero' ? CalculoHojas.formatoGeneral(x.valor) : String(x.valor)).length;
    else { if (x.tipo !== 'numero') return malo; valor = x.valor; if (v.tipo === 'whole' && !Number.isInteger(valor)) return malo; }
    const num = (y) => (typeof y === 'number' ? y : typeof y === 'string' && y.trim() !== '' && !Number.isNaN(+y) ? +y : null);
    const a = num(evaluarEn(ed, v, v.formula1, h, f, c)), b = num(evaluarEn(ed, v, v.formula2, h, f, c));
    if (a === null) return null;   // la condición no se puede calcular: Excel la deja pasar
    const op = v.operador || 'between';
    const bien = { between: () => b !== null && valor >= Math.min(a, b) && valor <= Math.max(a, b), notBetween: () => b !== null && (valor < Math.min(a, b) || valor > Math.max(a, b)), equal: () => valor === a, notEqual: () => valor !== a,
      greaterThan: () => valor > a, lessThan: () => valor < a, greaterThanOrEqual: () => valor >= a, lessThanOrEqual: () => valor <= a }[op];
    return !bien || bien() ? null : malo;
  }
  // la validación personalizada (una fórmula que debe dar VERDADERO), ya con el valor escrito en la celda
  function comprobarDespues(ed, h, f, c) {
    const v = validacionEn(ed, h, f, c);
    if (!v || v.tipo !== 'custom' || !v.mostrarError) return null;
    const x = ed.celda(h, f, c);
    if (!x || x.tipo === 'vacio') return null;
    const r = evaluarEn(ed, v, v.formula1, h, f, c);
    return r === true || (typeof r === 'number' && r !== 0) ? null : { estilo: v.estiloError || 'stop', titulo: v.tituloError || '', mensaje: v.error || '' };
  }

  globalThis.DatosLibro = { nombreValido, ponerNombre, quitarNombre, crearTabla, ponerValidacion, validacionEn, listaDe, comprobarEntrada, comprobarDespues, reglasCondicionales, agregarRegla, cambiarRegla, quitarRegla, moverRegla, _interno: { restar, leerReglas, xmlRegla } };
})();

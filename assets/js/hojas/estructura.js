/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Cambios de estructura de un libro de Mini Cookies Hojas (fase 2 del PLAN): insertar y quitar filas, agregar, renombrar, quitar y mover hojas,
// combinar celdas, inmovilizar paneles y ancho de columnas. Cada uno toma los bytes de un .xlsx y devuelve los bytes nuevos, editando el paquete
// en el lugar (xml-editar): lo que no tiene que ver con el cambio sale tal cual. El editor los deshace volviendo al libro anterior. Sin DOM.
//
// Requiere zip-propio, xml-editar, xlsx-modelo, xlsx-paquete, motores/xlsx-escritor, motores/xlsx-filas y motores/xlsx-columnas.
//
//   await EstructuraLibro.insertarFilas(bytes, hoja, fila, cuantas)
//   await EstructuraLibro.quitarFilas(bytes, hoja, [filas])
//   await EstructuraLibro.insertarColumnas(bytes, hoja, columna, cuantas, { nombreColumna }) · quitarColumnas(bytes, hoja, [columnas])
//     (dentro de una tabla, las columnas nuevas se llaman nombreColumna + número: «Column1» en Excel en inglés, «Columna1» en español)
//   await EstructuraLibro.agregarHoja(bytes, nombre, despuesDe)          → { bytes, indice }
//   await EstructuraLibro.renombrarHoja(bytes, hoja, nombreNuevo)
//   await EstructuraLibro.quitarHoja(bytes, hoja)
//   await EstructuraLibro.moverHoja(bytes, hoja, posicionNueva)
//   await EstructuraLibro.combinar(bytes, hoja, { f0, c0, f1, c1 })     (deja solo el valor de arriba a la izquierda, como Excel)
//   await EstructuraLibro.separar(bytes, hoja, { f0, c0, f1, c1 })      (quita las combinadas que tocan el rango)
//   await EstructuraLibro.inmovilizar(bytes, hoja, filas, columnas)    (0, 0 = quitar)
//   await EstructuraLibro.anchoColumnas(bytes, hoja, c0, c1, ancho)     (ancho en caracteres, como Excel)
//   await EstructuraLibro.altoFilas(bytes, hoja, f0, f1, alto)          (alto en puntos, como Excel; null = automático)
(function () {
  const X = () => globalThis.XmlEditar;
  const NS_REL_HOJA = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
  const TIPO_HOJA = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
  // Orden de los hijos de <worksheet> (CT_Worksheet): un elemento nuevo va antes del primero que lo sigue
  const ORDEN = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
  const col = (n) => LibroXlsx.numeroACol(n);
  const ref = (r) => col(r.c0) + (r.f0 + 1) + (r.f0 === r.f1 && r.c0 === r.c1 ? '' : ':' + col(r.c1) + (r.f1 + 1));

  async function conLibro(bytes, fn) {
    const libro = await LibroXlsx.abrir(bytes);
    const r = await fn(libro);
    const salida = await libro.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    return r && r.indice !== undefined ? { bytes: salida, indice: r.indice } : salida;
  }
  const texto = async (zip, ruta) => { const e = zip.file(ruta); return e ? e.async('string') : null; };
  // inserta `xml` como hijo de <worksheet> en su lugar según el orden de la norma (o reemplaza el elemento si ya está)
  function ponerEnHoja(hojaXml, nombre, xml) {
    const raiz = X().analizar(hojaXml, { profundidad: 2 });
    const existente = X().hijo(raiz, nombre);
    if (existente) return X().aplicar(hojaXml, [{ ini: existente.ini, fin: existente.fin, texto: xml }]);
    const k = ORDEN.indexOf(nombre);
    const sig = raiz.hijos.find((h) => ORDEN.indexOf(h.local) > k);
    const pos = sig ? sig.ini : raiz.iniCierre;
    return X().aplicar(hojaXml, [{ ini: pos, fin: pos, texto: xml }]);
  }
  function quitarDeHoja(hojaXml, nombre) {
    const raiz = X().analizar(hojaXml, { profundidad: 2 });
    const n = X().hijo(raiz, nombre);
    return n ? X().aplicar(hojaXml, [{ ini: n.ini, fin: n.fin, texto: '' }]) : hojaXml;
  }
  const rutaDe = (libro, h) => { const m = libro.hojas[h]; if (!m || m.tipo !== 'hoja' || !m.ruta) throw new XlsxEscritor.ErrorEscritor('la hoja no es de celdas'); return m.ruta; };

  // ---- filas ------------------------------------------------------------------------------------------------------
  const insertarFilas = (bytes, h, fila, n) => conLibro(bytes, (libro) => XlsxEscritor.insertarFilas(libro, h, fila, n));
  const quitarFilas = (bytes, h, filas) => conLibro(bytes, (libro) => XlsxEscritor.quitarFilas(libro, h, filas));
  const insertarColumnas = (bytes, h, col, n, opciones) => conLibro(bytes, (libro) => XlsxEscritor.insertarColumnas(libro, h, col, n, opciones));
  const quitarColumnas = (bytes, h, cols) => conLibro(bytes, (libro) => XlsxEscritor.quitarColumnas(libro, h, cols));

  // ---- hojas ------------------------------------------------------------------------------------------------------
  function nombreValido(nombre, otros) {
    const n = String(nombre || '').trim();
    if (!n || n.length > 31 || /[:\\/?*[\]]/.test(n) || /^'|'$/.test(n) || n.toLowerCase() === 'history') return null;
    if (otros.some((o) => o.toLowerCase() === n.toLowerCase())) return null;
    return n;
  }
  async function agregarHoja(bytes, nombre, despuesDe) {
    return conLibro(bytes, async (libro) => {
      const zip = libro.zip, P = PaqueteEditable.desdeZip(zip);
      const nombres = libro.hojas.map((x) => x.nombre);
      const n = nombreValido(nombre, nombres);
      if (!n) throw new XlsxEscritor.ErrorEscritor('nombre de hoja no válido o repetido');
      let k = 1; while (zip.file('xl/worksheets/sheet' + k + '.xml')) k++;
      const ruta = 'xl/worksheets/sheet' + k + '.xml';
      zip.file(ruta, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>');
      await P.asegurarTipo('/' + ruta, TIPO_HOJA);
      const rutaLibro = libro.rutaLibro;
      const dir = rutaLibro.slice(0, rutaLibro.lastIndexOf('/') + 1);
      const relativa = ruta.startsWith(dir) ? ruta.slice(dir.length) : '/' + ruta;
      const rid = await P.agregarRelacion(rutaLibro, NS_REL_HOJA, relativa);
      let wb = await texto(zip, rutaLibro);
      const raiz = X().analizar(wb, { profundidad: 3 }), sheets = X().hijo(raiz, 'sheets'), lista = X().hijos(sheets, 'sheet');
      const id = Math.max(0, ...lista.map((s) => +s.attrs.sheetId || 0)) + 1;
      const prefijoR = (/xmlns:(\w+)="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships"/.exec(wb) || [, 'r'])[1];
      const nodo = `<sheet name="${X().escAttr(n)}" sheetId="${id}" ${prefijoR}:id="${rid}"/>`;
      const tras = despuesDe === undefined || despuesDe === null ? lista[lista.length - 1] : lista[Math.min(lista.length - 1, Math.max(-1, despuesDe))];
      const pos = despuesDe === -1 ? lista[0].ini : tras ? tras.fin : sheets.iniCierre;
      wb = X().aplicar(wb, [{ ini: pos, fin: pos, texto: nodo }]);
      zip.file(rutaLibro, wb);
      const indice = despuesDe === undefined || despuesDe === null ? lista.length : Math.max(0, despuesDe + 1);
      // las posiciones (activeTab, localSheetId) de las hojas que quedan después de la nueva corren un lugar
      if (indice < lista.length) await correrPosiciones(zip, rutaLibro, indice, +1);
      return { indice };
    });
  }
  // localSheetId de los nombres definidos y activeTab / firstSheet son posiciones: insertar o quitar una hoja las corre
  async function correrPosiciones(zip, rutaLibro, desde, delta) {
    let wb = await texto(zip, rutaLibro);
    wb = wb.replace(/(<definedName\b[^>]*\blocalSheetId=")(\d+)(")/g, (m, a, n, c) => (+n >= desde ? a + (+n + delta) + c : m));
    wb = wb.replace(/(<workbookView\b[^>]*\b(?:activeTab|firstSheet)=")(\d+)(")/g, (m, a, n, c) => (+n >= desde ? a + Math.max(0, +n + delta) + c : m));
    zip.file(rutaLibro, wb);
  }
  async function renombrarHoja(bytes, h, nombreNuevo) {
    return conLibro(bytes, async (libro) => {
      const zip = libro.zip, viejo = libro.hojas[h].nombre;
      const n = nombreValido(nombreNuevo, libro.hojas.filter((x, i) => i !== h).map((x) => x.nombre));
      if (!n) throw new XlsxEscritor.ErrorEscritor('nombre de hoja no válido o repetido');
      if (n === viejo) return;
      const mapa = new Map([[viejo.toLowerCase(), n]]);   // renombrarHojas busca por nombre en minúsculas
      const cambiar = (f) => XlsxEscritor.renombrarHojas(f, mapa);
      // workbook.xml: el nombre y los nombres definidos
      let wb = await texto(zip, libro.rutaLibro);
      const raiz = X().analizar(wb, { profundidad: 3 }), sheet = X().hijos(X().hijo(raiz, 'sheets'), 'sheet')[h];
      wb = X().aplicar(wb, [{ ini: sheet.ini, fin: sheet.finApertura, texto: X().ponerAtributo(wb.slice(sheet.ini, sheet.finApertura), 'name', n) }]);
      wb = wb.replace(/(<definedName\b[^>]*>)([\s\S]*?)(<\/definedName>)/g, (m, a, f, c) => a + X().esc(cambiar(X().decodificar(f))) + c);
      zip.file(libro.rutaLibro, wb);
      // fórmulas de todas las hojas, de las tablas y de los gráficos que citan la hoja
      const rutas = [];
      zip.forEach((r) => { if (/^xl\/(worksheets|charts|tables)\/[^/]+\.xml$/i.test(r)) rutas.push(r); });
      const marca = viejo.toLowerCase();
      for (const r of rutas) {
        const t = await texto(zip, r);
        if (t === null || !t.toLowerCase().includes(X().esc(marca).replace(/'/g, '&apos;')) && !t.toLowerCase().includes(marca)) continue;
        const u = t.replace(/(<(?:\w+:)?(?:f|calculatedColumnFormula|totalsRowFormula)\b[^>]*>)([\s\S]*?)(<\/(?:\w+:)?(?:f|calculatedColumnFormula|totalsRowFormula)>)/g, (m, a, f, c) => a + X().esc(cambiar(X().decodificar(f))) + c);
        if (u !== t) zip.file(r, u);
      }
      // docProps/app.xml lista los nombres de las hojas
      const app = await texto(zip, 'docProps/app.xml');
      if (app) zip.file('docProps/app.xml', app.replace(new RegExp('(<vt:lpstr>)' + X().esc(viejo).replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(</vt:lpstr>)'), '$1' + X().esc(n) + '$2'));
    });
  }
  async function quitarHoja(bytes, h) {
    return conLibro(bytes, async (libro) => {
      const nombres = libro.hojas.map((x) => x.nombre);
      if (nombres.length <= 1) throw new XlsxEscritor.ErrorEscritor('un libro necesita al menos una hoja');
      await XlsxPaquete.reordenarHojas(libro.zip, nombres.filter((x, i) => i !== h));
    });
  }
  async function moverHoja(bytes, h, posicion) {
    return conLibro(bytes, async (libro) => {
      const nombres = libro.hojas.map((x) => x.nombre);
      const [n] = nombres.splice(h, 1);
      nombres.splice(Math.max(0, Math.min(nombres.length, posicion)), 0, n);
      await XlsxPaquete.reordenarHojas(libro.zip, nombres);
    });
  }

  // ---- combinar, inmovilizar, anchos ----------------------------------------------------------------------------------
  async function combinar(bytes, h, r) {
    return conLibro(bytes, async (libro) => {
      const ruta = rutaDe(libro, h);
      const hoja = await libro.cargarHoja(h);
      for (const m of hoja.combinadas || []) if (!(m.f1 < r.f0 || m.f0 > r.f1 || m.c1 < r.c0 || m.c0 > r.c1)) throw new XlsxEscritor.ErrorEscritor('el rango toca otra celda combinada');
      // como Excel: queda el valor de la esquina de arriba a la izquierda; las demás celdas se vacían (conservan su formato)
      const vaciar = [];
      hoja.celdas.paraCada((f, c, i) => { if (f >= r.f0 && f <= r.f1 && c >= r.c0 && c <= r.c1 && !(f === r.f0 && c === r.c0) && hoja.celdas.tipo[i] !== LibroXlsx.TIPO.vacio) vaciar.push({ f, c, valor: null }); });
      if (vaciar.length) await XlsxEscritor.editarCeldas(libro, h, vaciar);
      let xml = await texto(libro.zip, ruta);
      const raiz = X().analizar(xml, { profundidad: 2 }), mc = X().hijo(raiz, 'mergeCells');
      if (mc) {
        const n = (xml.slice(mc.finApertura, mc.iniCierre < 0 ? mc.finApertura : mc.iniCierre).match(/<(?:\w+:)?mergeCell\b/g) || []).length + 1;
        const apertura = X().ponerAtributo(xml.slice(mc.ini, mc.finApertura).replace(/\s*\/>$/, '>'), 'count', String(n));
        xml = X().aplicar(xml, mc.iniCierre < 0 ? [{ ini: mc.ini, fin: mc.fin, texto: apertura + `<mergeCell ref="${ref(r)}"/></mergeCells>` }] : [{ ini: mc.ini, fin: mc.finApertura, texto: apertura }, { ini: mc.iniCierre, fin: mc.iniCierre, texto: `<mergeCell ref="${ref(r)}"/>` }]);
      } else xml = ponerEnHoja(xml, 'mergeCells', `<mergeCells count="1"><mergeCell ref="${ref(r)}"/></mergeCells>`);
      libro.zip.file(ruta, xml);
    });
  }
  async function separar(bytes, h, r) {
    return conLibro(bytes, async (libro) => {
      const ruta = rutaDe(libro, h);
      let xml = await texto(libro.zip, ruta);
      const raiz = X().analizar(xml, { profundidad: 3 }), mc = X().hijo(raiz, 'mergeCells');
      if (!mc) return;
      const quedan = [], ed = [];
      for (const m of X().hijos(mc, 'mergeCell')) {
        const g = LibroXlsx.decodificarRango(m.attrs.ref || '');
        if (g && !(g.f1 < r.f0 || g.f0 > r.f1 || g.c1 < r.c0 || g.c0 > r.c1)) ed.push({ ini: m.ini, fin: m.fin, texto: '' }); else quedan.push(m);
      }
      if (!ed.length) return;
      xml = quedan.length ? X().aplicar(xml, ed.concat([{ ini: mc.ini, fin: mc.finApertura, texto: X().ponerAtributo(xml.slice(mc.ini, mc.finApertura), 'count', String(quedan.length)) }])) : X().aplicar(xml, [{ ini: mc.ini, fin: mc.fin, texto: '' }]);
      libro.zip.file(ruta, xml);
    });
  }
  async function inmovilizar(bytes, h, filas, columnas) {
    return conLibro(bytes, async (libro) => {
      const ruta = rutaDe(libro, h);
      let xml = await texto(libro.zip, ruta);
      const raiz = X().analizar(xml, { profundidad: 4 }), vistas = X().hijo(raiz, 'sheetViews');
      const vista = vistas && X().hijos(vistas, 'sheetView')[0];
      const pane = filas || columnas
        ? `<pane${columnas ? ` xSplit="${columnas}"` : ''}${filas ? ` ySplit="${filas}"` : ''} topLeftCell="${col(columnas) + (filas + 1)}" activePane="${filas && columnas ? 'bottomRight' : filas ? 'bottomLeft' : 'topRight'}" state="frozen"/>`
        : '';
      if (!vista) { xml = ponerEnHoja(xml, 'sheetViews', `<sheetViews><sheetView workbookViewId="0">${pane}</sheetView></sheetViews>`); libro.zip.file(ruta, xml); return; }
      const viejo = X().hijo(vista, 'pane');
      const ed = [];
      if (viejo) ed.push({ ini: viejo.ini, fin: viejo.fin, texto: pane });
      else if (pane) ed.push(vista.iniCierre < 0 ? { ini: vista.ini, fin: vista.fin, texto: xml.slice(vista.ini, vista.finApertura).replace(/\s*\/>$/, '>') + pane + '</sheetView>' } : { ini: vista.finApertura, fin: vista.finApertura, texto: pane });
      // las selecciones de paneles que ya no existen confunden a Excel: se dejan solo las del panel nuevo
      for (const s of X().hijos(vista, 'selection')) if (s.attrs.pane) ed.push({ ini: s.ini, fin: s.fin, texto: '' });
      if (ed.length) libro.zip.file(ruta, X().aplicar(xml, ed));
    });
  }
  async function anchoColumnas(bytes, h, c0, c1, ancho) {
    return conLibro(bytes, async (libro) => {
      const ruta = rutaDe(libro, h);
      let xml = await texto(libro.zip, ruta);
      const raiz = X().analizar(xml, { profundidad: 3 }), cols = X().hijo(raiz, 'cols');
      const nuevos = [];
      const lista = cols ? X().hijos(cols, 'col') : [];
      // los <col> que tocan [c0, c1] se parten: fuera del tramo conservan lo suyo; adentro, ancho nuevo con el resto de sus atributos
      const tramo = [];
      for (const k of lista) {
        const a = +k.attrs.min - 1, b = +k.attrs.max - 1, et = xml.slice(k.ini, k.fin);
        if (b < c0 || a > c1) { nuevos.push({ a, x: et }); continue; }
        if (a < c0) nuevos.push({ a, x: X().ponerAtributo(et, 'max', String(c0)) });
        if (b > c1) nuevos.push({ a: c1 + 1, x: X().ponerAtributo(et, 'min', String(c1 + 2)) });
        tramo.push({ a: Math.max(a, c0), b: Math.min(b, c1), et });
      }
      for (let c = c0; c <= c1;) {
        const t = tramo.find((x) => c >= x.a && c <= x.b);
        const fin = t ? t.b : Math.min(c1, ...tramo.filter((x) => x.a > c).map((x) => x.a - 1), c1);
        let et = t ? t.et : '<col/>';
        for (const [k, v] of [['min', c + 1], ['max', fin + 1], ['width', ancho], ['customWidth', 1]]) et = X().ponerAtributo(et, k, String(v));
        nuevos.push({ a: c, x: et });
        c = fin + 1;
      }
      nuevos.sort((p, q) => p.a - q.a);
      const bloque = '<cols>' + nuevos.map((n) => n.x).join('') + '</cols>';
      xml = cols ? X().aplicar(xml, [{ ini: cols.ini, fin: cols.fin, texto: bloque }]) : ponerEnHoja(xml, 'cols', bloque);
      libro.zip.file(ruta, xml);
    });
  }

  // Alto de las filas f0..f1 en puntos (como Excel: 0–409). Las filas que ya están en <sheetData> solo cambian ht y customHeight en su
  // etiqueta de apertura (sus celdas y demás atributos quedan igual); las que no están se agregan vacías, en su lugar, con el alto.
  // alto = null vuelve al alto automático (quita ht y customHeight).
  const MAX_FILAS_ALTO = 100000;
  async function altoFilas(bytes, h, f0, f1, alto) {
    if (alto !== null && !(alto >= 0 && alto <= 409)) throw new XlsxEscritor.ErrorEscritor('el alto de fila va de 0 a 409 puntos');
    if (f1 - f0 + 1 > MAX_FILAS_ALTO) throw new XlsxEscritor.ErrorEscritor('demasiadas filas a la vez');
    return conLibro(bytes, async (libro) => {
      const ruta = rutaDe(libro, h);
      let xml = await texto(libro.zip, ruta);
      const raiz = X().analizar(xml, { profundidad: 3 });
      let datos = X().hijo(raiz, 'sheetData');
      if (!datos) throw new XlsxEscritor.ErrorEscritor('la hoja no tiene sheetData');
      if (datos.iniCierre < 0) {   // <sheetData/>: se abre para poder agregar filas
        xml = X().aplicar(xml, [{ ini: datos.ini, fin: datos.fin, texto: '<sheetData></sheetData>' }]);
        datos = X().hijo(X().analizar(xml, { profundidad: 3 }), 'sheetData');
      }
      const filas = X().hijos(datos, 'row');
      const ediciones = [];
      const etiqueta = (r) => xml.slice(r.ini, r.iniCierre < 0 ? r.fin : r.finApertura);
      const conAlto = (et) => (alto === null ? X().ponerAtributo(X().ponerAtributo(et, 'ht', null), 'customHeight', null) : X().ponerAtributo(X().ponerAtributo(et, 'ht', String(alto)), 'customHeight', '1'));
      let k = 0, siguiente = 0;   // filas sin r: Excel las numera una después de la otra
      const numero = filas.map((r) => { const n = r.attrs.r ? parseInt(r.attrs.r, 10) - 1 : siguiente; siguiente = n + 1; return n; });
      for (let f = f0; f <= f1; f++) {
        while (k < filas.length && numero[k] < f) k++;
        if (k < filas.length && numero[k] === f) {
          const r = filas[k], et = etiqueta(r);
          ediciones.push({ ini: r.ini, fin: r.ini + et.length, texto: conAlto(et) });
        } else if (alto !== null) {
          const pos = k < filas.length ? filas[k].ini : datos.iniCierre;
          ediciones.push({ ini: pos, fin: pos, texto: '<row r="' + (f + 1) + '" ht="' + alto + '" customHeight="1"/>' });
        }
      }
      if (ediciones.length) xml = X().aplicar(xml, ediciones);
      libro.zip.file(ruta, xml);
    });
  }

  globalThis.EstructuraLibro = { _ponerEnHoja: ponerEnHoja, insertarFilas, quitarFilas, insertarColumnas, quitarColumnas, agregarHoja, renombrarHoja, quitarHoja, moverHoja, combinar, separar, inmovilizar, anchoColumnas, altoFilas, nombreValido };
})();

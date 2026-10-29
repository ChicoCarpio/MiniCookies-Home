/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Autofiltro de verdad de Mini Cookies Hojas (fase 3.3 del PLAN): poner y quitar el autofiltro de la hoja (o usar el de una tabla) y filtrar
// sus columnas como Excel: el criterio queda en <filterColumn> y las filas que no cumplen quedan ocultas en el archivo (hidden="1"), así
// SUBTOTALES y AGREGAR las saltan y Excel abre el libro igual. Como DatosLibro, toma bytes y devuelve bytes editando solo lo que cambia. Sin DOM.
//
//   await FiltrosLibro.ponerAutofiltro(bytes, hoja, { f0, c0, f1, c1 } | null)   (null lo quita y vuelve a mostrar las filas que ocultaba)
//   FiltrosLibro.filtroEn(hojaModelo, f, c) → { tipo: 'hoja'|'tabla', ref, tabla?, columnas: Map(colId → criterio) } | null
//   await FiltrosLibro.filtrar(bytes, hoja, objetivo, colId, criterio | null, textoDe)   (null borra el filtro de esa columna)
//     criterio = { tipo: 'valores', valores: [textos], vacias }                      (los textos son como se ven en la celda)
//              | { tipo: 'condiciones', y: true|false, condiciones: [{ operador, valor }] }   (equal, notEqual, greaterThan, …; * y ? en textos)
//              | { tipo: 'top10', n, porcentaje, abajo }
//     textoDe(f, c) → { texto, valor }  (lo que se ve y el valor de la celda; lo da el editor)
(function () {
  const X = () => globalThis.XmlEditar;
  const Err = (m) => new XlsxEscritor.ErrorEscritor(m);
  const ORDEN_HOJA = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
  const col = (n) => LibroXlsx.numeroACol(n);
  const refDe = (r) => col(r.c0) + (r.f0 + 1) + ':' + col(r.c1) + (r.f1 + 1);
  const absDe = (r) => '$' + col(r.c0) + '$' + (r.f0 + 1) + ':$' + col(r.c1) + '$' + (r.f1 + 1);
  const texto = async (zip, ruta) => { const e = zip.file(ruta); return e ? e.async('string') : null; };
  const OPERADORES = new Set(['equal', 'notEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual']);

  async function conLibro(bytes, fn) {
    const libro = await LibroXlsx.abrir(bytes);
    await fn(libro);
    return libro.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }
  const rutaHoja = (libro, h) => { const m = libro.hojas[h]; if (!m || m.tipo !== 'hoja' || !m.ruta) throw Err('la hoja no es de celdas'); return m.ruta; };
  const citarHoja = (n) => (/^[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'");

  // filas ocultas o visibles: cambia hidden en la etiqueta de cada <row> (las que no están se agregan vacías y ocultas)
  function ponerOcultas(xml, filas) {
    if (!filas.size) return xml;
    let raiz = X().analizar(xml, { profundidad: 3 });
    let datos = X().hijo(raiz, 'sheetData');
    if (!datos) throw Err('la hoja no tiene sheetData');
    if (datos.iniCierre < 0) { xml = X().aplicar(xml, [{ ini: datos.ini, fin: datos.fin, texto: '<sheetData></sheetData>' }]); datos = X().hijo(X().analizar(xml, { profundidad: 3 }), 'sheetData'); }
    const lista = X().hijos(datos, 'row');
    let sig = 0;
    const numero = lista.map((r) => { const n = r.attrs.r ? parseInt(r.attrs.r, 10) - 1 : sig; sig = n + 1; return n; });
    const porFila = new Map(numero.map((n, i) => [n, lista[i]]));
    const ed = [];
    const orden = [...filas.keys()].sort((a, b) => a - b);
    let k = 0;
    for (const f of orden) {
      const oculta = filas.get(f), r = porFila.get(f);
      if (r) {
        const et = xml.slice(r.ini, r.iniCierre < 0 ? r.fin : r.finApertura);
        const yaOculta = r.attrs.hidden === '1' || r.attrs.hidden === 'true';
        if (yaOculta !== oculta) ed.push({ ini: r.ini, fin: r.ini + et.length, texto: X().ponerAtributo(et, 'hidden', oculta ? '1' : null) });
      } else if (oculta) {
        while (k < lista.length && numero[k] < f) k++;
        const pos = k < lista.length ? lista[k].ini : datos.iniCierre;
        ed.push({ ini: pos, fin: pos, texto: `<row r="${f + 1}" hidden="1"/>` });
      }
    }
    return ed.length ? X().aplicar(xml, ed) : xml;
  }
  // el nombre oculto que Excel pone con el rango del autofiltro de la hoja
  async function nombreBaseFiltro(libro, h, ref) {
    let wb = await texto(libro.zip, libro.rutaLibro);
    const raiz = X().analizar(wb, { profundidad: 3 }), dn = X().hijo(raiz, 'definedNames');
    const lista = dn ? X().hijos(dn, 'definedName') : [];
    const x = lista.find((d) => d.attrs.name === '_xlnm._FilterDatabase' && d.attrs.localSheetId === String(h));
    const valor = ref ? citarHoja(libro.hojas[h].nombre) + '!' + absDe(ref) : null;
    if (x) {
      if (!valor) wb = X().aplicar(wb, [lista.length === 1 ? { ini: dn.ini, fin: dn.fin, texto: '' } : { ini: x.ini, fin: x.fin, texto: '' }]);
      else wb = X().aplicar(wb, [{ ini: x.finApertura, fin: x.iniCierre, texto: X().esc(valor) }]);
    } else if (valor) {
      const nodo = `<definedName name="_xlnm._FilterDatabase" localSheetId="${h}" hidden="1">${X().esc(valor)}</definedName>`;
      if (dn && dn.iniCierre >= 0) wb = X().aplicar(wb, [{ ini: dn.iniCierre, fin: dn.iniCierre, texto: nodo }]);
      else {
        const orden = ['fileVersion', 'fileSharing', 'workbookPr', 'workbookProtection', 'bookViews', 'sheets', 'functionGroups', 'externalReferences', 'definedNames'];
        const sig = raiz.hijos.find((n) => orden.indexOf(n.local) < 0);
        const pos = sig ? sig.ini : raiz.iniCierre;
        wb = X().aplicar(wb, [{ ini: pos, fin: pos, texto: '<definedNames>' + nodo + '</definedNames>' }]);
      }
    }
    libro.zip.file(libro.rutaLibro, wb);
  }

  function filtroEn(hoja, f, c) {
    for (const t of hoja.tablas || []) if (f >= t.ref.f0 && f <= t.ref.f1 && c >= t.ref.c0 && c <= t.ref.c1) return t.filtro && t.encabezados ? { tipo: 'tabla', ref: t.ref, tabla: t } : null;
    const fa = hoja.filtroAuto && LibroXlsx.decodificarRango(hoja.filtroAuto.ref || '');
    if (fa && c >= fa.c0 && c <= fa.c1 && f >= fa.f0 && f <= fa.f1) return { tipo: 'hoja', ref: fa };
    return null;
  }

  async function ponerAutofiltro(bytes, h, r) {
    return conLibro(bytes, async (libro) => {
      const ruta = rutaHoja(libro, h);
      let xml = await texto(libro.zip, ruta);
      const raiz = X().analizar(xml, { profundidad: 2 }), af = X().hijo(raiz, 'autoFilter');
      if (!r) {
        if (!af) return;
        // quitar el filtro vuelve a mostrar las filas que ocultaba (las del rango, como Excel)
        const viejo = LibroXlsx.decodificarRango(af.attrs.ref || '');
        const ss = X().hijo(raiz, 'sortState');
        xml = X().aplicar(xml, [{ ini: af.ini, fin: af.fin, texto: '' }].concat(ss ? [{ ini: ss.ini, fin: ss.fin, texto: '' }] : []));
        if (viejo) {
          const filas = new Map(); for (let f = viejo.f0 + 1; f <= viejo.f1; f++) filas.set(f, false);
          xml = ponerOcultas(xml, filas);
        }
        libro.zip.file(ruta, xml);
        await nombreBaseFiltro(libro, h, null);
        return;
      }
      if (r.f1 <= r.f0) throw Err('el filtro necesita un encabezado y al menos una fila');
      const hoja = await libro.cargarHoja(h);
      if ((hoja.tablas || []).some((t) => t.ref.f0 <= r.f1 && t.ref.f1 >= r.f0 && t.ref.c0 <= r.c1 && t.ref.c1 >= r.c0)) throw Err('el rango toca una tabla: las tablas ya traen su filtro');
      const nodo = `<autoFilter ref="${refDe(r)}"/>`;
      if (af) xml = X().aplicar(xml, [{ ini: af.ini, fin: af.fin, texto: nodo }]);
      else {
        const k = ORDEN_HOJA.indexOf('autoFilter'), sig = raiz.hijos.find((n) => ORDEN_HOJA.indexOf(n.local) > k);
        const pos = sig ? sig.ini : raiz.iniCierre;
        xml = X().aplicar(xml, [{ ini: pos, fin: pos, texto: nodo }]);
      }
      libro.zip.file(ruta, xml);
      await nombreBaseFiltro(libro, h, r);
    });
  }

  // ---- criterios ---------------------------------------------------------------------------------------------------------
  const comodin = (s) => new RegExp('^' + String(s).replace(/~([*?~])|([*?])|([.+^${}()|[\]\\])/g, (t, e, w, m) => (e ? '\\' + e : w === '*' ? '[\\s\\S]*' : w === '?' ? '[\\s\\S]' : '\\' + m)) + '$', 'i');
  function cumpleCondicion(cond, x) {
    const num = typeof x.valor === 'number' ? x.valor : null;
    const v = String(cond.valor === undefined || cond.valor === null ? '' : cond.valor);
    const nv = v.trim() !== '' && !Number.isNaN(+v) ? +v : null;
    const t = String(x.texto || '');
    switch (cond.operador) {
      case 'equal': return nv !== null && num !== null ? num === nv : comodin(v).test(t);
      case 'notEqual': return nv !== null && num !== null ? num !== nv : !comodin(v).test(t);
      default: {
        if (nv !== null) { if (num === null) return false; return { greaterThan: num > nv, greaterThanOrEqual: num >= nv, lessThan: num < nv, lessThanOrEqual: num <= nv }[cond.operador]; }
        const a = t.toLowerCase(), b = v.toLowerCase();
        return { greaterThan: a > b, greaterThanOrEqual: a >= b, lessThan: a < b, lessThanOrEqual: a <= b }[cond.operador];
      }
    }
  }
  // ¿la fila pasa el criterio de la columna? `todos`: los valores de la columna (para los superiores)
  function cumple(criterio, x, umbral) {
    if (!criterio) return true;
    if (criterio.tipo === 'valores') {
      if (x.texto === '' || x.texto === null || x.texto === undefined) return !!criterio.vacias;
      return criterio._set.has(String(x.texto).toLowerCase());
    }
    if (criterio.tipo === 'condiciones') {
      const r = criterio.condiciones.map((c) => cumpleCondicion(c, x));
      return criterio.y ? r.every(Boolean) : r.some(Boolean);
    }
    if (criterio.tipo === 'top10') return typeof x.valor === 'number' && (criterio.abajo ? x.valor <= umbral : x.valor >= umbral);
    return true;
  }
  function xmlColumna(colId, cr) {
    if (cr.tipo === 'valores') {
      const vals = [...new Set(cr.valores.map(String))];
      return `<filterColumn colId="${colId}"><filters${cr.vacias ? ' blank="1"' : ''}>${vals.map((v) => `<filter val="${X().escAttr(v)}"/>`).join('')}</filters></filterColumn>`;
    }
    if (cr.tipo === 'condiciones') {
      const cs = cr.condiciones.filter((c) => OPERADORES.has(c.operador) && String(c.valor).trim() !== '').slice(0, 2);
      if (!cs.length) throw Err('el filtro necesita un valor');
      return `<filterColumn colId="${colId}"><customFilters${cr.y && cs.length > 1 ? ' and="1"' : ''}>${cs.map((c) => `<customFilter${c.operador === 'equal' ? '' : ` operator="${c.operador}"`} val="${X().escAttr(String(c.valor))}"/>`).join('')}</customFilters></filterColumn>`;
    }
    if (cr.tipo === 'top10') {
      const n = Math.max(1, Math.min(500, Math.trunc(+cr.n || 10)));
      return `<filterColumn colId="${colId}"><top10${cr.abajo ? ' top="0"' : ''}${cr.porcentaje ? ' percent="1"' : ''} val="${n}"/></filterColumn>`;
    }
    throw Err('criterio de filtro desconocido');
  }
  // los criterios que ya tiene el autofiltro, leídos del XML (para volver a calcular qué filas se ven)
  function leerCriterios(afXml) {
    const m = new Map();
    for (const fc of afXml.matchAll(/<filterColumn\b[^>]*\bcolId="(\d+)"[^>]*>([\s\S]*?)<\/filterColumn>/g)) {
      const id = +fc[1], cuerpo = fc[2];
      if (/<filters\b/.test(cuerpo)) m.set(id, { tipo: 'valores', vacias: /<filters\b[^>]*\bblank="(1|true)"/.test(cuerpo), valores: [...cuerpo.matchAll(/<filter\b[^>]*\bval="([^"]*)"/g)].map((x) => X().decodificar(x[1])) });
      else if (/<customFilters\b/.test(cuerpo)) m.set(id, { tipo: 'condiciones', y: /<customFilters\b[^>]*\band="(1|true)"/.test(cuerpo), condiciones: [...cuerpo.matchAll(/<customFilter\b([^>]*)\/?>/g)].map((x) => ({ operador: (/\boperator="(\w+)"/.exec(x[1]) || [])[1] || 'equal', valor: X().decodificar((/\bval="([^"]*)"/.exec(x[1]) || [])[1] || '') })) });
      else if (/<top10\b/.test(cuerpo)) { const a = /<top10\b([^>]*)/.exec(cuerpo)[1]; m.set(id, { tipo: 'top10', n: +((/\bval="([\d.]+)"/.exec(a) || [])[1] || 10), porcentaje: /\bpercent="(1|true)"/.test(a), abajo: /\btop="(0|false)"/.test(a) }); }
      else m.set(id, { tipo: 'otro' });   // por color, dinámico…: se conserva y no se calcula
    }
    return m;
  }

  // qué filas se ven con todas las columnas filtradas a la vez: Map(fila → oculta) de la primera fila de datos a `ultima`
  function calcularOcultas(criterios, r, textoDe, ultima) {
    for (const cr of criterios.values()) if (cr.tipo === 'valores') cr._set = new Set(cr.valores.map((v) => v.toLowerCase()));
    const umbrales = new Map();
    // los superiores se cuentan entre las filas que pasan los demás filtros (los datos que se ven, como en Excel)
    const pasaOtros = (f, id) => { for (const [j, cr] of criterios) { if (j === id || cr.tipo === 'otro' || cr.tipo === 'top10') continue; if (!cumple(cr, textoDe(f, r.c0 + j))) return false; } return true; };
    for (const [id, cr] of criterios) if (cr.tipo === 'top10') {
      const nums = [];
      for (let f = r.f0 + 1; f <= r.f1; f++) { if (!pasaOtros(f, id)) continue; const x = textoDe(f, r.c0 + id); if (typeof x.valor === 'number') nums.push(x.valor); }
      nums.sort((a, b) => (cr.abajo ? a - b : b - a));
      const n = cr.porcentaje ? Math.max(1, Math.floor(nums.length * cr.n / 100)) : cr.n;
      umbrales.set(id, nums.length ? nums[Math.min(nums.length, n) - 1] : Infinity);
    }
    const filas = new Map();
    for (let f = r.f0 + 1; f <= ultima; f++) {
      let ve = true;
      for (const [id, cr] of criterios) { if (cr.tipo === 'otro') continue; if (!cumple(cr, textoDe(f, r.c0 + id), umbrales.get(id))) { ve = false; break; } }
      filas.set(f, !ve);
    }
    return filas;
  }

  async function filtrar(bytes, h, objetivo, colId, criterio, textoDe) {
    return conLibro(bytes, async (libro) => {
      const ruta = rutaHoja(libro, h);
      let xml = await texto(libro.zip, ruta);
      let parte = ruta, doc = xml;
      if (objetivo.tipo === 'tabla') {
        const rels = await XlsxPaquete.leerRels(libro.zip, ruta);
        for (const rel of rels.filter((x) => x.tipo === 'table')) { const t = await texto(libro.zip, rel.destino); const m = t && /<table\b[^>]*\bref="([^"]+)"/.exec(t); const g = m && LibroXlsx.decodificarRango(m[1]); if (g && g.f0 === objetivo.ref.f0 && g.c0 === objetivo.ref.c0) { parte = rel.destino; doc = t; break; } }
        if (parte === ruta) throw Err('no se encontró la tabla');
      }
      const raiz = X().analizar(doc, { profundidad: 2 }), af = X().hijo(raiz, 'autoFilter');
      if (!af) throw Err('no hay autofiltro');
      const r = objetivo.ref;
      if (colId < 0 || colId > r.c1 - r.c0) throw Err('columna fuera del filtro');
      // el <filterColumn> de esa columna (en orden de colId)
      const cuerpo = af.iniCierre >= 0 ? doc.slice(af.finApertura, af.iniCierre) : '';
      const cols = [...cuerpo.matchAll(/<filterColumn\b[^>]*\bcolId="(\d+)"[^>]*(?:\/>|>[\s\S]*?<\/filterColumn>)/g)].map((m) => ({ id: +m[1], xml: m[0] }));
      const otras = cols.filter((x) => x.id !== colId);
      if (criterio) otras.push({ id: colId, xml: xmlColumna(colId, criterio) });
      otras.sort((a, b) => a.id - b.id);
      const resto = cuerpo.replace(/<filterColumn\b[^>]*\bcolId="\d+"[^>]*(?:\/>|>[\s\S]*?<\/filterColumn>)/g, '');   // sortState y extLst del filtro quedan
      const apertura = doc.slice(af.ini, af.iniCierre >= 0 ? af.finApertura : af.fin).replace(/\s*\/>$/, '>');
      const nuevoAf = otras.length || resto ? apertura + otras.map((x) => x.xml).join('') + resto + '</autoFilter>' : apertura.replace(/>$/, '/>');
      doc = X().aplicar(doc, [{ ini: af.ini, fin: af.fin, texto: nuevoAf }]);
      if (parte !== ruta) libro.zip.file(parte, doc); else xml = doc;
      const ultima = objetivo.tipo === 'tabla' && objetivo.tabla && objetivo.tabla.totales ? r.f1 - 1 : r.f1;   // la fila de totales no se filtra
      xml = ponerOcultas(xml, calcularOcultas(leerCriterios(nuevoAf), r, textoDe, ultima));
      libro.zip.file(ruta, xml);
      if (libro._hojasCargadas) libro._hojasCargadas.clear();
    });
  }

  // los criterios del filtro (hoja o tabla) como objetos: Map(colId → criterio)
  async function criterios(bytes, h, objetivo) {
    const libro = await LibroXlsx.abrir(bytes);
    const ruta = rutaHoja(libro, h);
    let doc = await texto(libro.zip, ruta);
    if (objetivo.tipo === 'tabla') {
      for (const rel of (await XlsxPaquete.leerRels(libro.zip, ruta)).filter((x) => x.tipo === 'table')) { const t = await texto(libro.zip, rel.destino); const m = t && /<table\b[^>]*\bref="([^"]+)"/.exec(t); const g = m && LibroXlsx.decodificarRango(m[1]); if (g && g.f0 === objetivo.ref.f0 && g.c0 === objetivo.ref.c0) { doc = t; break; } }
    }
    const af = X().hijo(X().analizar(doc, { profundidad: 2 }), 'autoFilter');
    return af ? leerCriterios(doc.slice(af.ini, af.fin)) : new Map();
  }

  globalThis.FiltrosLibro = { criterios, ponerAutofiltro, filtroEn, filtrar, _interno: { ponerOcultas, leerCriterios, cumple, cumpleCondicion, calcularOcultas, OPERADORES } };
})();

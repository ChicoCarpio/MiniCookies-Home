/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Cambios de estructura de un libro OpenDocument (.ods) para Mini Cookies Hojas (F3.4a, etapa 4). Misma interfaz que `EstructuraLibro` (hojas/estructura.js): cada uno toma
// los bytes de un .ods y devuelve los bytes nuevos, editando content.xml (y settings.xml) en el lugar con xml-editar; lo que no tiene que ver con el cambio sale tal cual.
// Las referencias de TODAS las fórmulas del libro se acomodan con los mismos reescritores que usa el .xlsx (se traduce la fórmula a la sintaxis de Excel, se acomoda y se
// vuelve a ODF solo si cambió), y los nombres definidos (`table:named-range`, `table:named-expression`) también. Sin DOM.
//
//   await EstructuraOds.insertarFilas(bytes, hoja, fila, cuantas)          · quitarFilas(bytes, hoja, [filas])
//   await EstructuraOds.insertarColumnas(bytes, hoja, columna, cuantas)    · quitarColumnas(bytes, hoja, [columnas])
//   await EstructuraOds.agregarHoja(bytes, nombre, despuesDe) → { bytes, indice } · renombrarHoja(bytes, hoja, nombre) · quitarHoja(bytes, hoja) · moverHoja(bytes, hoja, posicion)
//   await EstructuraOds.combinar(bytes, hoja, rango) · separar(bytes, hoja, rango) · inmovilizar(bytes, hoja, filas, columnas)
//   await EstructuraOds.anchoColumnas(bytes, hoja, c0, c1, ancho)  (en caracteres, como Excel)  · altoFilas(bytes, hoja, f0, f1, alto)  (en puntos; null = automático)
//
// Formato condicional, validación de datos, autofiltro, áreas de impresión y gráficos (con sus rangos en `Object N/`) se acomodan con el cambio. Se rechazan (ErrorEscritor, sin
// tocar nada) los cambios que dejarían mal algo que no se sabe mover: objetos OLE, formas y dibujos anclados a celdas (`table:end-cell-address`), insertar o quitar columnas
// dentro de un filtro con criterios, y las filas o columnas que parten una combinada o una fórmula matricial.
// Requiere zip-propio, xml-editar, paquete-editable, ods-formulas, motores/xlsx-escritor (ErrorEscritor), xlsx-filas, xlsx-columnas y ods-escritor-celdas.
(function () {
  const X = () => globalThis.XmlEditar;
  const U = () => globalThis.OdsEscritorCeldas._util;
  const OF = () => globalThis.OdsFormulas;
  const MAX_FILAS = 1048576, MAX_COLS = 16384, MAX_AREA = 250000;
  const Err = (m) => new XlsxEscritor.ErrorEscritor(m);

  async function conPaquete(bytes, fn) {
    const p = await PaqueteEditable.abrir(bytes);
    const r = await fn(p);
    const salida = await p.generar();
    return r && r.indice !== undefined ? { bytes: salida, indice: r.indice } : salida;
  }
  function documento(texto) {
    const raiz = X().analizar(texto, { profundidad: 5 });
    const hojaCalc = X().hijo(X().hijo(raiz, 'body'), 'spreadsheet');
    if (!hojaCalc) throw Err('el archivo no es una hoja de cálculo');
    return { raiz, hojaCalc, tablas: X().hijos(hojaCalc, 'table') };
  }
  const tablaDe = (d, h) => { const t = d.tablas[h]; if (!t) throw Err('la hoja no existe'); return t; };
  const nombreDe = (t) => t.attrs['table:name'] || '';
  const entero = (s) => { const n = parseInt(s, 10); return n > 0 ? n : 0; };
  const parte = (texto, n) => texto.slice(n.ini, n.fin);

  // Lo que aún no se sabe mover: si el archivo lo tiene, el cambio se rechaza en vez de dejarlo apuntando a celdas que ya no son
  function comprobarReferencias(texto, que) {
    const cosas = [[/<draw:object-ole\b/, 'objetos OLE incrustados'], [/table:end-cell-address=/, 'dibujos anclados a celdas']];
    for (const [re, nombre] of cosas) if (re.test(texto)) throw Err('el archivo tiene ' + nombre + ': todavía no se pueden ' + que + ' en un .ods con eso');
  }
  // Un rango de base de datos con contenido (criterios de filtro u orden por columnas) guarda columnas relativas a su rango: mover columnas dentro de él las confundiría.
  // `dentro(c0, c1)` dice si el cambio de columnas toca ese rango.
  function comprobarColumnasBd(texto, que, nombreHoja, dentro) {
    for (const m of texto.matchAll(/<table:database-range\b[^>]*[^\/]>/g)) {
      const mt = RE_DB_DIRECCION.exec(m[0]);
      const rg = mt && /^\$?(?:'((?:[^']|'')+)'|([^.']+))\.\$?([A-Za-z]{1,3})\$?\d+(?::\$?(?:'(?:[^']|'')+'|[^.']+)?\.\$?([A-Za-z]{1,3})\$?\d+)?$/.exec(X().decodificar(mt[2]));
      if (!rg) continue;
      const hoja = rg[1] !== undefined ? rg[1].replace(/''/g, "'") : rg[2];
      if (hoja.toLowerCase() !== nombreHoja.toLowerCase()) continue;
      const c0 = globalThis.LibroXlsx.colANumero(rg[3].toUpperCase()), c1 = rg[4] ? globalThis.LibroXlsx.colANumero(rg[4].toUpperCase()) : c0;
      if (dentro(c0, c1)) throw Err('hay un filtro u orden con criterios en esas columnas: todavía no se pueden ' + que + ' ahí en un .ods');
    }
  }

  // Combinadas y matrices de una tabla: [{ tipo, f0, c0, f1, c1 }] (desde 0)
  function bloques(texto, tabla) {
    const salida = [];
    for (const { nodo, f0 } of U().enumerarFilas(texto, tabla)) {
      const t = parte(texto, nodo);
      if (!/number-(?:columns|rows)-spanned|number-matrix/.test(t)) continue;
      const r = X().analizar(t, { profundidad: 2 });
      let c = 0;
      for (const h of r.hijos) {
        if (h.local !== 'table-cell' && h.local !== 'covered-table-cell') continue;
        const rep = U().repeticiones(h, 'table:number-columns-repeated');
        const cs = entero(h.attrs['table:number-columns-spanned']), rs = entero(h.attrs['table:number-rows-spanned']);
        const mc = entero(h.attrs['table:number-matrix-columns-spanned']), mr = entero(h.attrs['table:number-matrix-rows-spanned']);
        if (cs > 1 || rs > 1) salida.push({ tipo: 'combinada', f0, c0: c, f1: f0 + Math.max(rs, 1) - 1, c1: c + Math.max(cs, 1) - 1 });
        if (mc || mr) salida.push({ tipo: 'matriz', f0, c0: c, f1: f0 + Math.max(mr, 1) - 1, c1: c + Math.max(mc, 1) - 1 });
        c += rep;
      }
    }
    return salida;
  }
  const NOMBRE_BLOQUE = { combinada: 'una combinada', matriz: 'una fórmula matricial' };

  const RE_DB = /<table:database-range\b[^>]*>/g;
  // las direcciones de un atributo que trae varias («Hoja.A1:Hoja.C3 'Mi hoja'.E1:'Mi hoja'.F9»): el nombre de la hoja puede llevar espacios si va entre comillas
  const RE_DIRECCIONES = /\$?(?:'(?:[^']|'')+'|[^\s.']+)\.\$?[A-Za-z]{1,3}\$?\d+(?::\$?(?:'(?:[^']|'')+'|[^\s.']+)?\.\$?[A-Za-z]{1,3}\$?\d+)?/g;
  const RE_DB_DIRECCION = /(\stable:target-range-address=")([^"]*)(")/;
  // «Hoja.A1:Hoja.B6» (como la escribe ODF en las bases de datos) ↔ «Hoja!A1:B6»
  function direccionAExcel(d) {
    const m = /^\$?(?:'((?:[^']|'')+)'|([^.']+))\.(\$?[A-Za-z]{1,3}\$?\d+)(?::\$?(?:'(?:[^']|'')+'|[^.']+)?\.(\$?[A-Za-z]{1,3}\$?\d+))?$/.exec(d);
    if (!m) return null;
    const hoja = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
    return citar(hoja) + '!' + m[3] + (m[4] ? ':' + m[4] : '');
  }
  function direccionAOds(xl) {
    const m = /^('(?:[^']|'')+'|[^!]+)!(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?$/.exec(xl);
    if (!m) return xl;
    return m[1] + '.' + m[2] + (m[3] ? ':' + m[1] + '.' + m[3] : '');
  }

  // ---- fórmulas y nombres: acomodar las referencias de todo el libro --------------------------------------------------------------------
  // rw(textoExcel, enObjetivo) → textoExcel nuevo; enObjetivo: la fórmula vive en la hoja que cambia
  function rehacer(of, enObjetivo, rw) {
    if (/\['[^']*'#/.test(of)) return of;   // referencia a otro archivo (['file:///…'#$Hoja.A1]): no la cambia mover celdas de este libro
    const xl = OF().aExcel(of);
    if (xl === null || xl === undefined) {
      if (/\[/.test(of)) throw Err('una fórmula no se puede traducir para acomodarla: ' + of.slice(0, 80));
      return of;
    }
    const nuevo = rw(xl, enObjetivo);
    if (nuevo === xl) return of;
    const ab = OF().aOpenFormula(nuevo);
    if (ab === null || ab === undefined) throw Err('una fórmula no se puede escribir de nuevo: ' + nuevo.slice(0, 80));
    return 'of:=' + ab;
  }
  function reescribirTodo(texto, h, rw, soloObjetos) {
    const d = documento(texto);
    const rangos = d.tablas.map((t) => [t.ini, t.fin]);
    const dueno = (pos) => { for (let k = 0; k < rangos.length; k++) if (pos >= rangos[k][0] && pos < rangos[k][1]) return k; return -1; };
    const eds = [];
    // table:formula="of:=…" de cada celda
    const reF = /(\stable:formula=")([^"]*)(")/g;
    let m;
    while (!soloObjetos && (m = reF.exec(texto))) {
      const of = X().decodificar(m[2]);
      if (!of.startsWith('of:=')) { if (/\[/.test(of)) throw Err('una fórmula no está en OpenFormula: ' + of.slice(0, 60)); continue; }
      const nuevo = rehacer(of, dueno(m.index) === h, rw);
      if (nuevo !== of) eds.push({ ini: m.index + m[1].length, fin: m.index + m[1].length + m[2].length, texto: X().escAttr(nuevo) });
    }
    // nombres definidos: table:named-range (una dirección) y table:named-expression (una fórmula sin «of:=»)
    const reN = /<table:named-(range|expression)\b[^>]*>/g;
    while ((m = reN.exec(texto))) {
      const etiqueta = m[0], ini = m.index;
      const attr = m[1] === 'range' ? 'table:cell-range-address' : 'table:expression';
      const ma = new RegExp('(\\s' + attr + '=")([^"]*)(")').exec(etiqueta);
      if (!ma) continue;
      const valor = X().decodificar(ma[2]);
      const original = m[1] === 'range' ? 'of:=[' + valor + ']' : 'of:=' + valor;
      const rehecho = rehacer(original, false, rw);
      if (rehecho === original) continue;
      if (/#REF!/.test(rehecho) && m[1] === 'range') { eds.push({ ini, fin: ini + etiqueta.length, texto: '' }); continue; }   // el nombre apuntaba a lo que se quitó
      const nuevo = m[1] === 'range' ? rehecho.slice(5, -1) : rehecho.slice(4);
      eds.push({ ini: ini + ma.index + ma[1].length, fin: ini + ma.index + ma[1].length + ma[2].length, texto: X().escAttr(nuevo) });
    }
    // rangos de bases de datos (los de «dar formato como tabla» o con botones de filtro, con o sin criterios): solo hay que acomodar su dirección;
    // los criterios y el orden usan columnas relativas al rango (las operaciones de columnas que los tocan se rechazan antes)
    const reB = new RegExp(RE_DB.source, "g");
    while ((m = reB.exec(texto))) {
      const etiqueta = m[0], ini = m.index;
      const mt = RE_DB_DIRECCION.exec(etiqueta);
      if (!mt) continue;
      const xl = direccionAExcel(X().decodificar(mt[2]));
      if (xl === null) continue;
      const nuevoXl = rw(xl, false);
      if (nuevoXl === xl) continue;
      if (/#REF!/.test(nuevoXl)) {   // el rango desapareció: se va el elemento entero
        const cierre = etiqueta.endsWith('/>') ? -1 : texto.indexOf('</table:database-range>', ini);
        eds.push({ ini, fin: cierre < 0 ? ini + etiqueta.length : cierre + 23, texto: '' });
        if (cierre >= 0) reB.lastIndex = cierre + 23;
        continue;
      }
      eds.push({ ini: ini + mt.index + mt[1].length, fin: ini + mt.index + mt[1].length + mt[2].length, texto: X().escAttr(direccionAOds(nuevoXl)) });
    }
    // áreas de impresión: table:print-ranges="Hoja.A1:Hoja.C40 Hoja.E1:Hoja.F9" (si el área desaparece, se va el atributo)
    const reP = /(\s(?:table:print-ranges|draw:notify-on-update-of-ranges)=")([^"]*)(")/g;
    while ((m = reP.exec(texto))) {
      const partes = X().decodificar(m[2]).match(RE_DIRECCIONES) || [];
      const nuevas = partes.map((q) => { const xl = direccionAExcel(q); if (xl === null) return q; const n = rw(xl, false); return /#REF!/.test(n) ? null : n === xl ? q : direccionAOds(n); }).filter((q) => q !== null);
      if (nuevas.join(' ') === partes.join(' ')) continue;
      if (nuevas.length) eds.push({ ini: m.index + m[1].length, fin: m.index + m[1].length + m[2].length, texto: X().escAttr(nuevas.join(' ')) });
      else eds.push({ ini: m.index, fin: m.index + m[0].length, texto: '' });
    }
    // validación de datos: la condición (con referencias entre corchetes) y la celda base, que dice a qué hoja pertenecen las referencias sin hoja
    const reV = /<table:content-validation\b[^>]*>/g;
    const nombreObjetivo = h >= 0 && d.tablas[h] ? nombreDe(d.tablas[h]) : null;
    while ((m = reV.exec(texto))) {
      const etiqueta = m[0], ini = m.index;
      const mc = /(\stable:condition=")([^"]*)(")/.exec(etiqueta), mb = /(\stable:base-cell-address=")([^"]*)(")/.exec(etiqueta);
      const baseTexto = mb ? X().decodificar(mb[2]) : '';
      const hojaBase = /^\$?(?:'((?:[^']|'')+)'|([^.']+))\./.exec(baseTexto);
      const enObj = !!hojaBase && nombreObjetivo !== null && (hojaBase[1] !== undefined ? hojaBase[1].replace(/''/g, "'") : hojaBase[2]).toLowerCase() === nombreObjetivo.toLowerCase();
      if (mc) {
        const cond = X().decodificar(mc[2]);
        const nueva = cond.replace(/\[[^\]]*\]/g, (g) => { const r = rehacer('of:=' + g, enObj, rw); return r.slice(4); });
        if (nueva !== cond) eds.push({ ini: ini + mc.index + mc[1].length, fin: ini + mc.index + mc[1].length + mc[2].length, texto: X().escAttr(nueva) });
      }
      if (mb) {
        const xl = direccionAExcel(baseTexto);
        if (xl !== null) {
          const nuevoXl = rw(xl, false);
          if (nuevoXl !== xl && !/#REF!/.test(nuevoXl)) eds.push({ ini: ini + mb.index + mb[1].length, fin: ini + mb.index + mb[1].length + mb[2].length, texto: X().escAttr(direccionAOds(nuevoXl)) });
        }
      }
    }
    // formato condicional: los rangos a los que se aplica (se acomodan o se van) y, en cada condición, sus referencias y su celda base
    const reC = /<calcext:conditional-format\b[^>]*>/g;
    while ((m = reC.exec(texto))) {
      const etiqueta = m[0], ini = m.index;
      const mt = /(\scalcext:target-range-address=")([^"]*)(")/.exec(etiqueta);
      if (!mt) continue;
      const partes = X().decodificar(mt[2]).match(RE_DIRECCIONES) || [];   // (no por espacios: «'Mi hoja'.A1:'Mi hoja'.A3» es una sola)
      const nuevas = [];
      let cambio = false;
      for (const parte of partes) {
        const xl = direccionAExcel(parte);
        if (xl === null) { nuevas.push(parte); continue; }
        const nuevoXl = rw(xl, false);
        if (nuevoXl === xl) { nuevas.push(parte); continue; }
        cambio = true;
        if (!/#REF!/.test(nuevoXl)) nuevas.push(direccionAOds(nuevoXl));
      }
      if (!cambio) continue;
      if (!nuevas.length) {   // se fue todo el rango: el elemento entero
        const cierre = texto.indexOf('</calcext:conditional-format>', ini);
        if (cierre >= 0) eds.push({ ini, fin: cierre + '</calcext:conditional-format>'.length, texto: '' });
      } else eds.push({ ini: ini + mt.index + mt[1].length, fin: ini + mt.index + mt[1].length + mt[2].length, texto: X().escAttr(nuevas.join(' ')) });
    }
    const reCond = /<calcext:condition\b[^>]*>/g;
    while ((m = reCond.exec(texto))) {
      const etiqueta = m[0], ini = m.index;
      const mv = /(\scalcext:value=")([^"]*)(")/.exec(etiqueta), mb = /(\scalcext:base-cell-address=")([^"]*)(")/.exec(etiqueta);
      const baseTexto = mb ? X().decodificar(mb[2]) : '';
      const hojaBase = /^\$?(?:'((?:[^']|'')+)'|([^.']+))\./.exec(baseTexto);
      const nombreObj = h >= 0 && d.tablas[h] ? nombreDe(d.tablas[h]) : null;
      const enObj = !!hojaBase && nombreObj !== null && (hojaBase[1] !== undefined ? hojaBase[1].replace(/''/g, "'") : hojaBase[2]).toLowerCase() === nombreObj.toLowerCase();
      if (mv) {
        const valor = X().decodificar(mv[2]);
        const nuevo = valor.replace(/\[[^\]]*\]/g, (g) => rehacer('of:=' + g, enObj, rw).slice(4));
        if (nuevo !== valor) eds.push({ ini: ini + mv.index + mv[1].length, fin: ini + mv.index + mv[1].length + mv[2].length, texto: X().escAttr(nuevo) });
      }
      if (mb) {
        const xl = direccionAExcel(baseTexto);
        if (xl !== null) { const nuevoXl = rw(xl, false); if (nuevoXl !== xl && !/#REF!/.test(nuevoXl)) eds.push({ ini: ini + mb.index + mb[1].length, fin: ini + mb.index + mb[1].length + mb[2].length, texto: X().escAttr(direccionAOds(nuevoXl)) }); }
      }
    }
    return eds.length ? X().aplicar(texto, eds) : texto;
  }

  // ---- qué fila (columna) va a dónde ------------------------------------------------------------------------------------------------------------
  const mapaInsercion = (r, n) => {
    const f = (x) => (x < r ? x : x + n);
    return { fila: f, col: f, rango: (a, b) => { const x = f(a), y = f(b); return y > 1048575 ? (x > 1048575 ? null : [x, 1048575]) : [x, y]; }, quita: () => 0, antes: (x) => (x < r ? 0 : -n), hasta: (x) => (x < r ? 0 : -n), lista: [], insercion: { r, c: r, n } };
  };
  const rwFilas = (M, nombre) => (xl, enObjetivo) => globalThis.XlsxEscritor._quitarFilasInterno.reescribirFormula(xl, enObjetivo, nombre, M);
  const rwColumnas = (M, nombre) => (xl, enObjetivo) => globalThis.XlsxEscritor._columnasInterno.reescribirFormula(xl, enObjetivo, nombre, M);

  // ---- estilos automáticos nuevos (fila, columna) ---------------------------------------------------------------------------------------------
  function nombresUsados(texto) { const s = new Set(); for (const m of texto.matchAll(/style:name="([^"]*)"/g)) s.add(m[1]); return s; }
  function agregarEstilos(texto, xml) {
    const auto = X().hijo(documento(texto).raiz, 'automatic-styles');
    if (!auto) throw Err('el archivo no tiene office:automatic-styles');
    if (auto.iniCierre >= 0) return X().aplicar(texto, [{ ini: auto.iniCierre, fin: auto.iniCierre, texto: xml }]);
    return X().aplicar(texto, [{ ini: auto.ini, fin: auto.fin, texto: U().sinCierre(texto.slice(auto.ini, auto.finApertura)) + xml + '</' + auto.nombre + '>' }]);
  }

  // ---- cola de filas / columnas vacías: se recorta para que lo que se corre no pase del límite --------------------------------------------------
  const sinContenido = (t) => !/office:value-type|table:formula|<office:annotation|<draw:/.test(t);
  function recortarColaFilas(texto, tabla, exceso) {
    const filas = U().enumerarFilas(texto, tabla), u = filas[filas.length - 1];
    if (!u || u.n <= exceso || !sinContenido(parte(texto, u.nodo))) throw Err('no caben las filas nuevas: empujarían datos fuera de la hoja');
    return X().aplicar(texto, [{ ini: u.nodo.ini, fin: u.nodo.fin, texto: U().copiaConRepeticion(texto, u.nodo, 'table:number-rows-repeated', u.n - exceso) }]);
  }

  // ---- gráficos: sus rangos viven en Object N/content.xml (no en content.xml) --------------------------------------------------------------------------
  // rw(textoExcel, enObjetivo) como en reescribirTodo; un rango que desaparece se quita del gráfico
  async function reescribirObjetos(p, rw) {
    for (const ruta of p.rutas()) {
      if (!/^Object \d+\/content\.xml$/.test(ruta)) continue;
      const texto = await p.leer(ruta);
      if (texto === null || !/cell-range-address|label-cell-address/.test(texto)) continue;
      const eds = [];
      for (const m of texto.matchAll(/(\s(?:table:cell-range-address|chart:values-cell-range-address|chart:label-cell-address)=")([^"]*)(")/g)) {
        const partes = X().decodificar(m[2]).match(RE_DIRECCIONES) || [];
        const nuevas = partes.map((q) => { const xl = direccionAExcel(q); if (xl === null) return q; const n = rw(xl, false); return /#REF!/.test(n) ? null : n === xl ? q : direccionAOds(n); }).filter((q) => q !== null);
        if (nuevas.join(' ') === partes.join(' ')) continue;
        eds.push(nuevas.length ? { ini: m.index + m[1].length, fin: m.index + m[1].length + m[2].length, texto: X().escAttr(nuevas.join(' ')) } : { ini: m.index, fin: m.index + m[0].length, texto: '' });
      }
      if (eds.length) p.escribir(ruta, X().aplicar(texto, eds));
    }
  }
  // los objetos incrustados que ya no tiene ningún marco (se quitó la hoja o las celdas donde estaban) se van del paquete y del manifiesto
  async function limpiarObjetos(p) {
    const texto = await p.leer('content.xml');
    if (texto === null) return;
    const numeros = new Set(p.rutas().map((r) => /^Object (\d+)\//.exec(r)).filter(Boolean).map((m) => m[1]));
    let manifiesto = null;
    for (const n of numeros) {
      const enlazado = (s) => texto.indexOf(s + '"') >= 0 || texto.indexOf(s + '/') >= 0;
      if (enlazado('Object ' + n) || enlazado('Object%20' + n)) continue;   // (cualquier forma de enlazarlo: ./Object 1, #Object 1, ./Object%201…)
      for (const r of p.rutas()) if (r.startsWith('Object ' + n + '/') || r === 'ObjectReplacements/Object ' + n) p.quitar(r);
      if (manifiesto === null) manifiesto = await p.leer('META-INF/manifest.xml');
      if (manifiesto) manifiesto = manifiesto.replace(new RegExp('<manifest:file-entry\\b[^>]*manifest:full-path="(?:Object ' + n + '/[^"]*|ObjectReplacements/Object ' + n + ')"[^>]*/>', 'g'), '');
    }
    if (manifiesto !== null) p.escribir('META-INF/manifest.xml', manifiesto);
  }

  // lo que apunta a celdas sin ser una celda (nombres, validación, formato condicional, gráficos): lo usa mover celdas (hojas/mover-objetos.js)
  const reescribirReferencias = (bytes, h, rw) => conPaquete(bytes, async (p) => {
    const texto = await p.leer('content.xml');
    p.escribir('content.xml', reescribirTodo(texto, h, rw, true));
    await reescribirObjetos(p, rw);
  });

  // ======================================================== filas ========================================================
  async function insertarFilas(bytes, h, fila, n) {
    if (!(n >= 1)) throw Err('cuántas filas');
    return conPaquete(bytes, async (p) => {
      let texto = await p.leer('content.xml');
      comprobarReferencias(texto, 'insertar filas');
      let d = documento(texto), tabla = tablaDe(d, h);
      const nombre = nombreDe(tabla);
      for (const b of bloques(texto, tabla)) if (b.f0 < fila && b.f1 >= fila) throw Err('no se pueden insertar filas dentro de ' + NOMBRE_BLOQUE[b.tipo]);
      let filas = U().enumerarFilas(texto, tabla);
      const total = filas.length ? filas[filas.length - 1].f0 + filas[filas.length - 1].n : 0;
      if (fila < total) {
        if (total + n > MAX_FILAS) { texto = recortarColaFilas(texto, tabla, total + n - MAX_FILAS); d = documento(texto); tabla = tablaDe(d, h); filas = U().enumerarFilas(texto, tabla); }
        const el = filas.find((x) => x.f0 <= fila && fila < x.f0 + x.n);
        const vacia = '<table:table-row' + (n > 1 ? ' table:number-rows-repeated="' + n + '"' : '') + '><table:table-cell table:number-columns-repeated="' + MAX_COLS + '"/></table:table-row>';
        const titulos = X().hijos(tabla, 'table-header-rows')[0];
        if (el && titulos && fila === el.f0 && el.nodo.ini === titulos.finApertura) {   // delante de las filas que se repiten al imprimir: las filas nuevas quedan fuera del grupo (los títulos se corren, como en Excel)
          texto = X().aplicar(texto, [{ ini: titulos.ini, fin: titulos.ini, texto: vacia }]);
        } else if (el) {
          const piezas = [];
          if (fila > el.f0) piezas.push(U().copiaConRepeticion(texto, el.nodo, 'table:number-rows-repeated', fila - el.f0));
          piezas.push(vacia);
          piezas.push(U().copiaConRepeticion(texto, el.nodo, 'table:number-rows-repeated', el.f0 + el.n - fila));
          texto = X().aplicar(texto, [{ ini: el.nodo.ini, fin: el.nodo.fin, texto: piezas.join('') }]);
        }
      }
      p.escribir('content.xml', reescribirTodo(texto, h, rwFilas(mapaInsercion(fila, n), nombre)));
      await reescribirObjetos(p, rwFilas(mapaInsercion(fila, n), nombre));
      await limpiarObjetos(p);
    });
  }
  async function quitarFilas(bytes, h, lista) {
    const quitar = [...new Set(lista)].filter((f) => f >= 0 && f < MAX_FILAS).sort((a, b) => a - b);
    if (!quitar.length) throw Err('qué filas');
    return conPaquete(bytes, async (p) => {
      let texto = await p.leer('content.xml');
      comprobarReferencias(texto, 'quitar filas');
      const d = documento(texto), tabla = tablaDe(d, h), nombre = nombreDe(tabla);
      const set = new Set(quitar);
      for (const b of bloques(texto, tabla)) {
        let dentro = 0;
        for (let f = b.f0; f <= b.f1; f++) if (set.has(f)) dentro++;
        if (dentro && dentro < b.f1 - b.f0 + 1) throw Err('las filas que quitas parten ' + NOMBRE_BLOQUE[b.tipo]);
      }
      const eds = [];
      for (const el of U().enumerarFilas(texto, tabla)) {
        const dentro = quitar.filter((f) => f >= el.f0 && f < el.f0 + el.n);
        if (!dentro.length) continue;
        const piezas = [];
        let pos = el.f0;
        for (const f of dentro) { if (f > pos) piezas.push(U().copiaConRepeticion(texto, el.nodo, 'table:number-rows-repeated', f - pos)); pos = f + 1; }
        if (pos < el.f0 + el.n) piezas.push(U().copiaConRepeticion(texto, el.nodo, 'table:number-rows-repeated', el.f0 + el.n - pos));
        eds.push({ ini: el.nodo.ini, fin: el.nodo.fin, texto: piezas.join('') });
      }
      texto = X().aplicar(texto, eds);
      p.escribir('content.xml', reescribirTodo(texto, h, rwFilas(globalThis.XlsxEscritor._quitarFilasInterno.mapeo(quitar), nombre)));
      await reescribirObjetos(p, rwFilas(globalThis.XlsxEscritor._quitarFilasInterno.mapeo(quitar), nombre));
      await limpiarObjetos(p);
    });
  }

  // ======================================================== columnas ========================================================
  // Cambia las celdas de una fila (texto, con su raíz analizada): fn(cursor de columnas, nodo de celda, repeticiones) → texto nuevo | null (sin cambio)
  function filaPorCeldas(t, fn, alFinal) {
    const r = X().analizar(t, { profundidad: 2 });
    if (r.iniCierre < 0) { const cola = alFinal ? alFinal(0) : ''; return cola ? U().sinCierre(t) + cola + '</' + r.nombre + '>' : t; }
    const eds = [];
    let col = 0;
    for (const h of r.hijos) {
      if (h.local !== 'table-cell' && h.local !== 'covered-table-cell') continue;
      const rep = U().repeticiones(h, 'table:number-columns-repeated');
      const nuevo = fn(col, h, rep, t);
      if (nuevo !== null && nuevo !== undefined) eds.push({ ini: h.ini, fin: h.fin, texto: nuevo });
      col += rep;
    }
    if (alFinal) { const cola = alFinal(col); if (cola) eds.push({ ini: r.iniCierre, fin: r.iniCierre, texto: cola }); }
    return eds.length ? X().aplicar(t, eds) : t;
  }
  const anchoDeFila = (t) => { const r = X().analizar(t, { profundidad: 2 }); let c = 0; for (const h of r.hijos) if (h.local === 'table-cell' || h.local === 'covered-table-cell') c += U().repeticiones(h, 'table:number-columns-repeated'); return c; };

  // las definiciones de columna: aplica `piezasDe(def)` → texto nuevo | null
  function porDefiniciones(texto, tabla, piezasDe) {
    const eds = [];
    for (const def of U().enumerarColumnas(texto, tabla)) { const nuevo = piezasDe(def); if (nuevo !== null && nuevo !== undefined) eds.push({ ini: def.nodo.ini, fin: def.nodo.fin, texto: nuevo }); }
    return eds;
  }
  const filasDeTabla = (texto, tabla) => U().enumerarFilas(texto, tabla);

  async function insertarColumnas(bytes, h, c, n) {
    if (!(n >= 1)) throw Err('cuántas columnas');
    return conPaquete(bytes, async (p) => {
      let texto = await p.leer('content.xml');
      comprobarReferencias(texto, 'insertar columnas');
      let d = documento(texto), tabla = tablaDe(d, h);
      const nombre = nombreDe(tabla);
      comprobarColumnasBd(texto, 'insertar columnas', nombre, (c0, c1) => c > c0 && c <= c1);
      for (const b of bloques(texto, tabla)) if (b.c0 < c && b.c1 >= c) throw Err('no se pueden insertar columnas dentro de ' + NOMBRE_BLOQUE[b.tipo]);
      // 1) las definiciones de columna
      const defs = U().enumerarColumnas(texto, tabla);
      const totalDefs = defs.length ? defs[defs.length - 1].f0 + defs[defs.length - 1].n : 0;
      let eds = [];
      if (c < totalDefs) {
        const exceso = totalDefs + n - MAX_COLS;
        const ultima = defs[defs.length - 1];
        if (exceso > 0 && (ultima.n <= exceso || !sinContenido(parte(texto, ultima.nodo)))) throw Err('no caben las columnas nuevas');
        for (const def of defs) {
          const esUltima = def === ultima;
          const cubre = def.f0 <= c && c < def.f0 + def.n;
          if (!cubre && !(esUltima && exceso > 0)) continue;
          let derecha = cubre ? def.f0 + def.n - c : 0, izquierda = cubre ? c - def.f0 : def.n;
          const piezas = [];
          if (cubre) {
            if (izquierda > 0) piezas.push(U().copiaConRepeticion(texto, def.nodo, 'table:number-columns-repeated', izquierda));
            piezas.push(U().copiaConRepeticion(texto, def.nodo, 'table:number-columns-repeated', n));
            if (esUltima && exceso > 0) derecha -= exceso;
            if (derecha > 0) piezas.push(U().copiaConRepeticion(texto, def.nodo, 'table:number-columns-repeated', derecha));
          } else piezas.push(U().copiaConRepeticion(texto, def.nodo, 'table:number-columns-repeated', def.n - exceso));
          eds.push({ ini: def.nodo.ini, fin: def.nodo.fin, texto: piezas.join('') });
        }
      }
      // 2) las celdas de cada fila
      for (const el of filasDeTabla(texto, tabla)) {
        const t = parte(texto, el.nodo);
        let nueva = filaPorCeldas(t, (col, hijo, rep, tt) => {
          if (c < col || c >= col + rep) return null;
          const izq = c - col, der = col + rep - c;
          return (izq > 0 ? U().copiaConRepeticion(tt, hijo, 'table:number-columns-repeated', izq) : '') + U().celdaVacia(n) + U().copiaConRepeticion(tt, hijo, 'table:number-columns-repeated', der);
        });
        if (nueva !== t) {
          const ancho = anchoDeFila(nueva);
          if (ancho > MAX_COLS) {
            const exceso = ancho - MAX_COLS;
            const r2 = X().analizar(nueva, { profundidad: 2 });
            const celdas = r2.hijos.filter((x) => x.local === 'table-cell' || x.local === 'covered-table-cell');
            const u = celdas[celdas.length - 1];
            if (!u || U().repeticiones(u, 'table:number-columns-repeated') <= exceso || !sinContenido(nueva.slice(u.ini, u.fin))) throw Err('insertar columnas empujaría datos fuera de la hoja');
            nueva = X().aplicar(nueva, [{ ini: u.ini, fin: u.fin, texto: U().copiaConRepeticion(nueva, u, 'table:number-columns-repeated', U().repeticiones(u, 'table:number-columns-repeated') - exceso) }]);
          }
          eds.push({ ini: el.nodo.ini, fin: el.nodo.fin, texto: nueva });
        }
      }
      texto = X().aplicar(texto, eds);
      p.escribir('content.xml', reescribirTodo(texto, h, rwColumnas(globalThis.XlsxEscritor._columnasInterno.mapeoInsercion(c, n), nombre)));
      await reescribirObjetos(p, rwColumnas(globalThis.XlsxEscritor._columnasInterno.mapeoInsercion(c, n), nombre));
      await limpiarObjetos(p);
      void d;
    });
  }
  async function quitarColumnas(bytes, h, lista) {
    const quitar = [...new Set(lista)].filter((x) => x >= 0 && x < MAX_COLS).sort((a, b) => a - b);
    if (!quitar.length) throw Err('qué columnas');
    return conPaquete(bytes, async (p) => {
      let texto = await p.leer('content.xml');
      comprobarReferencias(texto, 'quitar columnas');
      const d = documento(texto), tabla = tablaDe(d, h), nombre = nombreDe(tabla);
      comprobarColumnasBd(texto, 'quitar columnas', nombre, (c0, c1) => quitar.some((x) => x >= c0 && x <= c1));
      const set = new Set(quitar);
      for (const b of bloques(texto, tabla)) {
        let dentro = 0;
        for (let x = b.c0; x <= b.c1; x++) if (set.has(x)) dentro++;
        if (dentro && dentro < b.c1 - b.c0 + 1) throw Err('las columnas que quitas parten ' + NOMBRE_BLOQUE[b.tipo]);
      }
      const resto = (col, rep) => {   // los trozos [a, b) de la repetición que se quedan
        const trozos = []; let pos = col;
        for (const x of quitar) { if (x < col || x >= col + rep) continue; if (x > pos) trozos.push(x - pos); pos = x + 1; }
        if (pos < col + rep) trozos.push(col + rep - pos);
        return trozos;
      };
      const eds = porDefiniciones(texto, tabla, (def) => {
        if (!quitar.some((x) => x >= def.f0 && x < def.f0 + def.n)) return null;
        return resto(def.f0, def.n).map((k) => U().copiaConRepeticion(texto, def.nodo, 'table:number-columns-repeated', k)).join('');
      });
      for (const el of filasDeTabla(texto, tabla)) {
        const t = parte(texto, el.nodo);
        const nueva = filaPorCeldas(t, (col, hijo, rep, tt) => {
          if (!quitar.some((x) => x >= col && x < col + rep)) return null;
          return resto(col, rep).map((k) => U().copiaConRepeticion(tt, hijo, 'table:number-columns-repeated', k)).join('');
        });
        if (nueva !== t) eds.push({ ini: el.nodo.ini, fin: el.nodo.fin, texto: nueva });
      }
      texto = X().aplicar(texto, eds);
      p.escribir('content.xml', reescribirTodo(texto, h, rwColumnas(globalThis.XlsxEscritor._columnasInterno.mapeoQuitar(quitar), nombre)));
      await reescribirObjetos(p, rwColumnas(globalThis.XlsxEscritor._columnasInterno.mapeoQuitar(quitar), nombre));
      await limpiarObjetos(p);
    });
  }

  // ======================================================== hojas ========================================================
  const NOMBRE_SIMPLE = /^[A-Za-zÀ-￿_][A-Za-z0-9_.À-￿]*$/;
  const citar = (n) => (NOMBRE_SIMPLE.test(n) ? n : "'" + n.replace(/'/g, "''") + "'");
  function validarNombre(d, nombre, salvo) {
    const n = String(nombre || '').trim();
    if (!n) throw Err('el nombre de la hoja está vacío');
    if (/[\[\]*?:\/\\]/.test(n) || /^'|'$/.test(n)) throw Err('nombre de hoja no válido');
    if (d.tablas.some((t, k) => k !== salvo && nombreDe(t).toLowerCase() === n.toLowerCase())) throw Err('nombre de hoja repetido');
    return n;
  }
  // recorre un texto de fórmula de Excel saltando las cadenas entre comillas dobles
  function fueraDeCadenas(xl, fn) {
    return xl.split(/("(?:[^"]|"")*")/).map((trozo, k) => (k % 2 ? trozo : fn(trozo))).join('');
  }
  const RE_PREFIJO = (nombre) => new RegExp("(^|[^A-Za-z0-9_.'À-￿$])(?:'" + nombre.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/'/g, "''") + "'|" + nombre.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ")!", 'gi');
  const renombrarEn = (viejo, nuevo) => (xl) => fueraDeCadenas(xl, (t) => t.replace(RE_PREFIJO(viejo), (m, previo) => previo + citar(nuevo) + '!'));
  const quitarEn = (viejo) => (xl) => fueraDeCadenas(xl, (t) => t.replace(new RegExp(RE_PREFIJO(viejo).source + '(\\$?[A-Za-z]{1,3}\\$?\\d+(?::\\$?[A-Za-z]{1,3}\\$?\\d+)?|\\$?\\d+:\\$?\\d+|\\$?[A-Za-z]{1,3}:\\$?[A-Za-z]{1,3})', 'gi'), (m, previo) => previo + '#REF!'));

  const TABLA_VACIA = (nombre) => '<table:table table:name="' + X().escAttr(nombre) + '"><table:table-column table:number-columns-repeated="' + MAX_COLS + '"/><table:table-row table:number-rows-repeated="' + MAX_FILAS
    + '"><table:table-cell table:number-columns-repeated="' + MAX_COLS + '"/></table:table-row></table:table>';
  async function agregarHoja(bytes, nombre, despuesDe) {
    return conPaquete(bytes, async (p) => {
      const texto = await p.leer('content.xml');
      const d = documento(texto);
      const n = validarNombre(d, nombre, -1);
      const k = despuesDe === undefined || despuesDe === null ? d.tablas.length - 1 : Math.max(-1, Math.min(despuesDe, d.tablas.length - 1));
      const pos = k < 0 ? d.tablas[0].ini : d.tablas[k].fin;
      p.escribir('content.xml', X().aplicar(texto, [{ ini: pos, fin: pos, texto: TABLA_VACIA(n) }]));
      return { indice: k + 1 };
    });
  }
  async function renombrarHoja(bytes, h, nombre) {
    return conPaquete(bytes, async (p) => {
      let texto = await p.leer('content.xml');
      comprobarReferencias(texto, 'renombrar hojas');
      const d = documento(texto), t = tablaDe(d, h);
      const nuevo = validarNombre(d, nombre, h), viejo = nombreDe(t);
      if (nuevo === viejo) return;
      texto = X().aplicar(texto, [{ ini: t.ini, fin: t.finApertura, texto: X().ponerAtributo(texto.slice(t.ini, t.finApertura), 'table:name', nuevo) }]);
      p.escribir('content.xml', reescribirTodo(texto, h, renombrarEn(viejo, nuevo)));
      await reescribirObjetos(p, renombrarEn(viejo, nuevo));
      await limpiarObjetos(p);
      const ajustes = await p.leer('settings.xml');
      // (solo las entradas de hoja: una hoja que se llame «Tables» o «Views» no debe renombrar los mapas de los ajustes)
      const entrada = (n) => '<config:config-item-map-entry config:name="' + X().escAttr(n) + '"';
      const activa = (n) => 'config:name="ActiveTable" config:type="string">' + X().esc(n) + '</config:config-item>';
      if (ajustes && (ajustes.includes(entrada(viejo)) || ajustes.includes(activa(viejo)))) p.escribir('settings.xml', ajustes.split(entrada(viejo)).join(entrada(nuevo)).split(activa(viejo)).join(activa(nuevo)));
    });
  }
  async function quitarHoja(bytes, h) {
    return conPaquete(bytes, async (p) => {
      let texto = await p.leer('content.xml');
      comprobarReferencias(texto, 'quitar hojas');
      const d = documento(texto), t = tablaDe(d, h);
      if (d.tablas.length < 2) throw Err('un libro necesita al menos una hoja');
      const viejo = nombreDe(t);
      texto = X().aplicar(texto, [{ ini: t.ini, fin: t.fin, texto: '' }]);
      p.escribir('content.xml', reescribirTodo(texto, -1, quitarEn(viejo)));
      await reescribirObjetos(p, quitarEn(viejo));
      await limpiarObjetos(p);
    });
  }
  async function moverHoja(bytes, h, posicion) {
    return conPaquete(bytes, async (p) => {
      const texto = await p.leer('content.xml');
      const d = documento(texto), t = tablaDe(d, h);
      const destino = Math.max(0, Math.min(posicion, d.tablas.length - 1));
      if (destino === h) return;
      const otra = d.tablas[destino];
      const eds = [{ ini: t.ini, fin: t.fin, texto: '' }, destino > h ? { ini: otra.fin, fin: otra.fin, texto: parte(texto, t) } : { ini: otra.ini, fin: otra.ini, texto: parte(texto, t) }];
      p.escribir('content.xml', X().aplicar(texto, eds));
    });
  }

  // ======================================================== combinar, separar ========================================================
  const sinAtributos = (etiqueta, lista) => lista.reduce((e, a) => X().ponerAtributo(e, a, null), etiqueta);
  const ATR_VALOR = ['office:value-type', 'office:value', 'office:date-value', 'office:time-value', 'office:boolean-value', 'office:string-value', 'office:currency', 'calcext:value-type', 'table:formula',
    'table:number-matrix-columns-spanned', 'table:number-matrix-rows-spanned', 'table:number-columns-repeated'];
  async function combinar(bytes, h, r) {
    if ((r.f1 - r.f0 + 1) * (r.c1 - r.c0 + 1) > MAX_AREA) throw Err('el rango es demasiado grande para combinarlo');
    if (r.f0 === r.f1 && r.c0 === r.c1) return bytes;
    return conPaquete(bytes, async (p) => {
      const texto = await p.leer('content.xml');
      const d = documento(texto), tabla = tablaDe(d, h);
      for (const b of bloques(texto, tabla)) if (b.f0 <= r.f1 && b.f1 >= r.f0 && b.c0 <= r.c1 && b.c1 >= r.c0) throw Err('el rango toca ' + NOMBRE_BLOQUE[b.tipo]);
      const lista = [];
      for (let f = r.f0; f <= r.f1; f++) for (let c = r.c0; c <= r.c1; c++) lista.push({ f, c });
      const eds = U().crearMotor(texto).editarTabla(tabla, lista, (orig, t, e) => {
        if (e.f === r.f0 && e.c === r.c0) {   // la de arriba a la izquierda conserva su valor y todo lo suyo
          const base = orig ? t.slice(orig.ini, orig.finApertura) : '<table:table-cell/>';
          const et = X().ponerAtributo(X().ponerAtributo(X().ponerAtributo(base, 'table:number-columns-repeated', null), 'table:number-columns-spanned', String(r.c1 - r.c0 + 1)), 'table:number-rows-spanned', String(r.f1 - r.f0 + 1));
          if (!orig || orig.iniCierre < 0) return et;
          return U().sinCierre(et) + t.slice(orig.finApertura, orig.fin);
        }
        const estilo = orig && orig.attrs['table:style-name'];
        return '<table:covered-table-cell' + (estilo ? ' table:style-name="' + X().escAttr(estilo) + '"' : '') + '/>';
      });
      p.escribir('content.xml', X().aplicar(texto, eds));
    });
  }
  async function separar(bytes, h, r) {
    return conPaquete(bytes, async (p) => {
      const texto = await p.leer('content.xml');
      const d = documento(texto), tabla = tablaDe(d, h);
      const tocadas = bloques(texto, tabla).filter((b) => b.tipo === 'combinada' && b.f0 <= r.f1 && b.f1 >= r.f0 && b.c0 <= r.c1 && b.c1 >= r.c0);
      if (!tocadas.length) return;
      const lista = [];
      for (const b of tocadas) for (let f = b.f0; f <= b.f1; f++) for (let c = b.c0; c <= b.c1; c++) lista.push({ f, c, ancla: f === b.f0 && c === b.c0 });
      if (lista.length > MAX_AREA) throw Err('las combinadas son demasiado grandes para separarlas');
      const eds = U().crearMotor(texto).editarTabla(tabla, lista, (orig, t, e) => {
        if (e.ancla) {
          const et = X().ponerAtributo(X().ponerAtributo(X().ponerAtributo(t.slice(orig.ini, orig.finApertura), 'table:number-columns-spanned', null), 'table:number-rows-spanned', null), 'table:number-columns-repeated', null);
          return orig.iniCierre < 0 ? et : U().sinCierre(et) + t.slice(orig.finApertura, orig.fin);
        }
        const estilo = orig && orig.attrs['table:style-name'];
        return '<table:table-cell' + (estilo ? ' table:style-name="' + X().escAttr(estilo) + '"' : '') + '/>';
      });
      p.escribir('content.xml', X().aplicar(texto, eds));
    });
  }

  // ======================================================== ancho y alto ========================================================
  async function anchoColumnas(bytes, h, c0, c1, ancho) {
    if (!(ancho > 0)) throw Err('ancho no válido');
    return conPaquete(bytes, async (p) => {
      let texto = await p.leer('content.xml');
      const d = documento(texto), tabla = tablaDe(d, h);
      const usados = nombresUsados(texto);
      let n = 0; while (usados.has('mcc' + (++n))); const nombre = 'mcc' + n;
      const cm = (ancho * 7 / 96 * 2.54).toFixed(3) + 'cm';
      const defs = U().enumerarColumnas(texto, tabla);
      const total = defs.length ? defs[defs.length - 1].f0 + defs[defs.length - 1].n : 0;
      const eds = [];
      for (const def of defs) {
        const a = Math.max(c0, def.f0), b = Math.min(c1, def.f0 + def.n - 1);
        if (a > b) continue;
        const piezas = [];
        if (a > def.f0) piezas.push(U().copiaConRepeticion(texto, def.nodo, 'table:number-columns-repeated', a - def.f0));
        piezas.push(X().ponerAtributo(U().copiaConRepeticion(texto, def.nodo, 'table:number-columns-repeated', b - a + 1), 'table:style-name', nombre));
        if (b < def.f0 + def.n - 1) piezas.push(U().copiaConRepeticion(texto, def.nodo, 'table:number-columns-repeated', def.f0 + def.n - 1 - b));
        eds.push({ ini: def.nodo.ini, fin: def.nodo.fin, texto: piezas.join('') });
      }
      if (c1 >= total) {   // columnas más allá de las que declara la tabla
        const desde = Math.max(c0, total);
        const nuevo = '<table:table-column table:style-name="' + nombre + '"' + (c1 - desde + 1 > 1 ? ' table:number-columns-repeated="' + (c1 - desde + 1) + '"' : '') + '/>';
        const relleno = desde > total ? '<table:table-column' + (desde - total > 1 ? ' table:number-columns-repeated="' + (desde - total) + '"' : '') + '/>' : '';
        const pos = defs.length ? defs[defs.length - 1].nodo.fin : tabla.finApertura;
        eds.push({ ini: pos, fin: pos, texto: relleno + nuevo });
      }
      texto = X().aplicar(texto, eds);
      p.escribir('content.xml', agregarEstilos(texto, '<style:style style:name="' + nombre + '" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="' + cm + '"/></style:style>'));
    });
  }
  async function altoFilas(bytes, h, f0, f1, alto) {
    return conPaquete(bytes, async (p) => {
      let texto = await p.leer('content.xml');
      const d = documento(texto), tabla = tablaDe(d, h);
      const usados = nombresUsados(texto);
      let n = 0; while (usados.has('mcr' + (++n))); const nombre = 'mcr' + n;
      const prop = alto === null || alto === undefined ? '<style:table-row-properties style:use-optimal-row-height="true"/>'
        : '<style:table-row-properties style:row-height="' + (alto / 72 * 2.54).toFixed(3) + 'cm" style:use-optimal-row-height="false"/>';
      let filas = U().enumerarFilas(texto, tabla);
      const total = filas.length ? filas[filas.length - 1].f0 + filas[filas.length - 1].n : 0;
      let eds = [];
      if (f1 >= total) {   // filas más allá de las que tiene la tabla
        const pos = filas.length ? U().finFuera(tabla, filas[filas.length - 1].nodo) : tabla.finApertura;
        const falta = f1 + 1 - total;
        eds.push({ ini: pos, fin: pos, texto: '<table:table-row table:number-rows-repeated="' + falta + '"><table:table-cell table:number-columns-repeated="' + MAX_COLS + '"/></table:table-row>' });
        texto = X().aplicar(texto, eds); eds = [];
        filas = U().enumerarFilas(texto, tablaDe(documento(texto), h));
      }
      for (const el of filas) {
        const a = Math.max(f0, el.f0), b = Math.min(f1, el.f0 + el.n - 1);
        if (a > b) continue;
        const piezas = [];
        if (a > el.f0) piezas.push(U().copiaConRepeticion(texto, el.nodo, 'table:number-rows-repeated', a - el.f0));
        piezas.push(X().ponerAtributo(U().copiaConRepeticion(texto, el.nodo, 'table:number-rows-repeated', b - a + 1), 'table:style-name', nombre));
        if (b < el.f0 + el.n - 1) piezas.push(U().copiaConRepeticion(texto, el.nodo, 'table:number-rows-repeated', el.f0 + el.n - 1 - b));
        eds.push({ ini: el.nodo.ini, fin: el.nodo.fin, texto: piezas.join('') });
      }
      texto = X().aplicar(texto, eds);
      p.escribir('content.xml', agregarEstilos(texto, '<style:style style:name="' + nombre + '" style:family="table-row">' + prop + '</style:style>'));
    });
  }

  // ======================================================== inmovilizar (settings.xml) ========================================================
  const item = (nombre, tipo, valor) => '<config:config-item config:name="' + nombre + '" config:type="' + tipo + '">' + valor + '</config:config-item>';
  async function inmovilizar(bytes, h, filas, columnas) {
    return conPaquete(bytes, async (p) => {
      const contenido = await p.leer('content.xml');
      const nombre = nombreDe(tablaDe(documento(contenido), h));
      let s = await p.leer('settings.xml');
      if (s === null) throw Err('este archivo no guarda paneles inmovilizados (no tiene settings.xml)');
      const ajustes = { HorizontalSplitMode: ['short', columnas > 0 ? 2 : 0], VerticalSplitMode: ['short', filas > 0 ? 2 : 0], HorizontalSplitPosition: ['int', columnas], VerticalSplitPosition: ['int', filas],
        ActiveSplitRange: ['short', filas > 0 && columnas > 0 ? 3 : columnas > 0 ? 1 : 2], PositionLeft: ['int', 0], PositionRight: ['int', columnas], PositionTop: ['int', 0], PositionBottom: ['int', filas] };
      const entrada = '<config:config-item-map-entry config:name="' + X().escAttr(nombre) + '">' + Object.entries(ajustes).map(([k, [t, v]]) => item(k, t, v)).join('') + '</config:config-item-map-entry>';
      const raiz = X().analizar(s, { profundidad: 8 });
      const ajustesVista = X().hijos(X().hijo(raiz, 'settings'), 'config-item-set').find((x) => x.attrs['config:name'] === 'ooo:view-settings');
      const ECONFIG = /xmlns:config=/.test(s.slice(0, 3000)) ? [] : [{ ini: raiz.ini, fin: raiz.finApertura, texto: X().ponerAtributo(s.slice(raiz.ini, raiz.finApertura), 'xmlns:config', 'urn:oasis:names:tc:opendocument:xmlns:config:1.0') }];
      if (!ajustesVista) {   // sin ajustes de vista (archivos de otros programas): se crean con la vista de esta hoja
        const cuerpo = X().hijo(raiz, 'settings');
        if (!cuerpo) throw Err('settings.xml no tiene office:settings');
        const conjunto = '<config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>' + item('ViewId', 'string', 'view1')
          + '<config:config-item-map-named config:name="Tables">' + entrada + '</config:config-item-map-named></config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set>';
        const e = cuerpo.iniCierre >= 0 ? { ini: cuerpo.iniCierre, fin: cuerpo.iniCierre, texto: conjunto }
          : { ini: cuerpo.ini, fin: cuerpo.fin, texto: U().sinCierre(s.slice(cuerpo.ini, cuerpo.finApertura)) + conjunto + '</' + cuerpo.nombre + '>' };
        p.escribir('settings.xml', X().aplicar(s, [e, ...ECONFIG]));
        return;
      }
      let vistas = X().hijos(ajustesVista, 'config-item-map-indexed').find((x) => x.attrs['config:name'] === 'Views');
      let eds;
      if (!vistas) {   // el archivo no guardó vistas (openpyxl, LibreOffice sin abrir la hoja): se crea la vista con su mapa de hojas
        const vista = '<config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>' + item('ViewId', 'string', 'view1') + '<config:config-item-map-named config:name="Tables">' + entrada + '</config:config-item-map-named></config:config-item-map-entry></config:config-item-map-indexed>';
        eds = [{ ini: ajustesVista.iniCierre, fin: ajustesVista.iniCierre, texto: vista }];
      } else {
        const vista = X().hijo(vistas, 'config-item-map-entry');
        const tablas = vista && X().hijos(vista, 'config-item-map-named').find((x) => x.attrs['config:name'] === 'Tables');
        if (!tablas) eds = [{ ini: vista.iniCierre, fin: vista.iniCierre, texto: '<config:config-item-map-named config:name="Tables">' + entrada + '</config:config-item-map-named>' }];
        else {
          const hoja = X().hijos(tablas, 'config-item-map-entry').find((x) => x.attrs['config:name'] === nombre);
          if (!hoja) eds = [{ ini: tablas.iniCierre, fin: tablas.iniCierre, texto: entrada }];
          else {
            eds = [];
            const presentes = new Set();
            for (const it of X().hijos(hoja, 'config-item')) {
              const k = it.attrs['config:name'];
              if (!(k in ajustes)) continue;
              presentes.add(k);
              eds.push({ ini: it.finApertura, fin: it.iniCierre, texto: String(ajustes[k][1]) });
            }
            const faltan = Object.entries(ajustes).filter(([k]) => !presentes.has(k)).map(([k, [t, v]]) => item(k, t, v)).join('');
            if (faltan) eds.push({ ini: hoja.iniCierre, fin: hoja.iniCierre, texto: faltan });
          }
        }
      }
      p.escribir('settings.xml', X().aplicar(s, eds.concat(ECONFIG)));
    });
  }

  globalThis.EstructuraOds = { reescribirReferencias, insertarFilas, quitarFilas, insertarColumnas, quitarColumnas, agregarHoja, renombrarHoja, quitarHoja, moverHoja, combinar, separar, inmovilizar, anchoColumnas, altoFilas,
    _interno: { conPaquete, documento, tablaDe, nombreDe, direccionAExcel, direccionAOds, rehacer, reescribirTodo, citar, comprobarReferencias, nombresUsados, agregarEstilos, bloques, Err } };
})();

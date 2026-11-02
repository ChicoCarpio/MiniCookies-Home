/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Autofiltro de un libro OpenDocument (.ods) para Mini Cookies Hojas (F3.4b, etapa 6), con la misma interfaz que `FiltrosLibro` (hojas/filtros-libro.js).
// ODF guarda el autofiltro como un rango de base de datos (`table:database-range` con `table:display-filter-buttons="true"`) cuyo `table:filter` lleva las
// condiciones de todas las columnas (`table:field-number` es relativo al rango), y las filas que no pasan llevan `table:visibility="filter"`. Cada operación toma
// los bytes de un .ods y devuelve los nuevos editando solo content.xml. Sin DOM.
//
//   await FiltrosOds.ponerAutofiltro(bytes, hoja, { f0, c0, f1, c1 } | null)   (null lo quita y vuelve a mostrar las filas que ocultaba)
//   FiltrosOds.filtroEn(hojaModelo, f, c)                                       (igual que FiltrosLibro)
//   await FiltrosOds.filtrar(bytes, hoja, objetivo, colId, criterio | null, textoDe)    (los criterios son los de FiltrosLibro)
//   await FiltrosOds.criterios(bytes, hoja, objetivo) → Map(colId → criterio)
//
// Lo que ODF dice distinto de Excel: «empieza con / termina con / contiene» son operadores propios (en Excel son comodines: «x*», «*x», «*x*»), y las condiciones de varias
// columnas viven en un solo `table:filter` (un Y de grupos, cada grupo con su O). Una condición que no se puede llevar a Excel (expresión regular, color…) queda como
// criterio `{ tipo: 'otro' }`: se conserva y no se calcula, y editar el filtro de esa hoja se rechaza para no perderla.
// Requiere xml-editar, paquete-editable, ods-estructura, ods-escritor-celdas y hojas/filtros-libro.
(function () {
  const X = () => globalThis.XmlEditar;
  const I = () => globalThis.EstructuraOds._interno;
  const U = () => globalThis.OdsEscritorCeldas._util;
  const F = () => globalThis.FiltrosLibro;
  const Err = (m) => new XlsxEscritor.ErrorEscritor(m);
  const OPERADORES = { '=': 'equal', '!=': 'notEqual', '>': 'greaterThan', '>=': 'greaterThanOrEqual', '<': 'lessThan', '<=': 'lessThanOrEqual' };
  const OPERADORES_ODF = Object.fromEntries(Object.entries(OPERADORES).map(([k, v]) => [v, k]));
  const SUPERIORES = { 'top values': { abajo: false, porcentaje: false }, 'bottom values': { abajo: true, porcentaje: false }, 'top percent': { abajo: false, porcentaje: true }, 'bottom percent': { abajo: true, porcentaje: true } };
  const SIN_COMODINES = /^[^*?~]*$/;
  const esNumero = (v) => String(v).trim() !== '' && !Number.isNaN(+v);

  // «Hoja.A1:Hoja.C6» → { hoja, r: { f0, c0, f1, c1 } }
  function direccion(d) {
    const xl = I().direccionAExcel(d);
    if (!xl) return null;
    const i = xl.lastIndexOf('!');
    let hoja = xl.slice(0, i);
    if (hoja[0] === "'") hoja = hoja.slice(1, -1).replace(/''/g, "'");
    const r = globalThis.LibroXlsx.decodificarRango(xl.slice(i + 1).replace(/\$/g, ''));
    return r ? { hoja, r } : null;
  }
  const refExcel = (r) => globalThis.LibroXlsx.codificarDireccion(r.f0, r.c0) + ':' + globalThis.LibroXlsx.codificarDireccion(r.f1, r.c1);
  const destinoOds = (nombre, r) => I().direccionAOds(I().citar(nombre) + '!' + refExcel(r));

  // ---- los rangos de base de datos del libro -----------------------------------------------------------------------------------------------------
  // [{ nodo, ini, fin, finApertura, base, attrs, destino, nombreSinDefinir }]; las posiciones son del texto completo
  function bases(texto, d) {
    const lista = [];
    for (const c of X().hijos(d.hojaCalc, 'database-ranges')) {
      if (c.iniCierre < 0) continue;
      const sub = X().analizar(texto.slice(c.ini, c.fin), { profundidad: 8 });
      for (const x of sub.hijos) {
        if (x.local !== 'database-range') continue;
        lista.push({ cont: c, nodo: x, base: c.ini, ini: c.ini + x.ini, fin: c.ini + x.fin, finApertura: c.ini + x.finApertura, attrs: x.attrs, destino: direccion(x.attrs['table:target-range-address'] || '') });
      }
    }
    return lista;
  }
  const delFiltro = (b, nombreHoja) => b.destino && b.attrs['table:display-filter-buttons'] === 'true' && b.destino.hoja.toLowerCase() === nombreHoja.toLowerCase();
  const buscar = (texto, d, h) => { const nombre = I().nombreDe(I().tablaDe(d, h)); return bases(texto, d).find((b) => delFiltro(b, nombre)) || null; };

  // ---- leer: table:filter → criterios de Excel ----------------------------------------------------------------------------------------------------
  const condicion = (n) => ({ campo: parseInt(n.attrs['table:field-number'], 10) || 0, operador: n.attrs['table:operator'] || '=', valor: n.attrs['table:value'] === undefined ? '' : n.attrs['table:value'],
    items: n.hijos.filter((x) => x.local === 'filter-set-item').map((x) => x.attrs['table:value'] || '') });
  const escapar = (v) => String(v).replace(/([*?~])/g, '~$1');
  // una condición suelta de ODF → { operador, valor } de Excel, o null si Excel no la tiene
  function simple(c) {
    if (OPERADORES[c.operador]) return { operador: OPERADORES[c.operador], valor: c.valor };
    switch (c.operador) {
      case 'begins-with': return { operador: 'equal', valor: escapar(c.valor) + '*' };
      case 'ends-with': return { operador: 'equal', valor: '*' + escapar(c.valor) };
      case 'contains': return { operador: 'equal', valor: '*' + escapar(c.valor) + '*' };
      case 'does-not-contain': return { operador: 'notEqual', valor: '*' + escapar(c.valor) + '*' };
      case '!empty': return { operador: 'notEqual', valor: '' };
      default: return null;
    }
  }
  function traducir(grupo) {
    const cs = grupo.conds;
    if (cs.length === 1) {
      const c = cs[0];
      if (c.operador === 'empty') return { tipo: 'valores', valores: [], vacias: true };
      if (SUPERIORES[c.operador]) { const n = Math.trunc(+c.valor); return n >= 1 ? Object.assign({ tipo: 'top10', n }, SUPERIORES[c.operador]) : { tipo: 'otro' }; }
      if (c.operador === '=' && c.items.length) return { tipo: 'valores', valores: c.items, vacias: false };
      const s = simple(c);
      return s ? { tipo: 'condiciones', y: false, condiciones: [s] } : { tipo: 'otro' };   // (con una sola condición «y» no cuenta: así lo lee también Excel)
    }
    if (cs.length === 2) {
      // «estos valores o las vacías»: así lo escribe Excel cuando se marca «(Vacías)» junto a otros valores
      const vacia = cs.find((c) => c.operador === 'empty'), resto = cs.find((c) => c !== vacia);
      if (vacia && !grupo.y && resto && resto.operador === '=') return { tipo: 'valores', valores: resto.items.length ? resto.items : [resto.valor], vacias: true };
      const t = cs.map(simple);
      if (t.every(Boolean)) return { tipo: 'condiciones', y: grupo.y, condiciones: t };
    }
    return { tipo: 'otro' };
  }
  // el nodo table:filter → Map(columna → criterio)
  function leerFiltro(nodo) {
    const m = new Map();
    const filtro = nodo && nodo.hijos.find((x) => x.local === 'filter');
    const raiz = filtro && filtro.hijos.find((x) => x.local === 'filter-and' || x.local === 'filter-or' || x.local === 'filter-condition');
    if (!raiz) return m;
    const grupos = new Map();   // columna → { y, conds }
    const marcar = (campo) => grupos.set(campo, { rota: true, conds: [] });
    const poner = (c, y) => {
      const g = grupos.get(c.campo);
      if (!g) grupos.set(c.campo, { y, conds: [c] });
      else if (g.rota || g.y !== y) marcar(c.campo);
      else g.conds.push(c);
    };
    const recorrer = (n, y) => {
      for (const h of n.hijos) {
        if (h.local === 'filter-condition') poner(condicion(h), y);
        else if (h.local === 'filter-and' || h.local === 'filter-or') {
          const hijoY = h.local === 'filter-and', conds = h.hijos.filter((x) => x.local === 'filter-condition').map(condicion);
          if (h.hijos.some((x) => x.local !== 'filter-condition')) { for (const c of conds) marcar(c.campo); continue; }   // anidado más hondo: no se lleva
          if (hijoY === y) { for (const c of conds) poner(c, y); continue; }
          const campos = new Set(conds.map((c) => c.campo));
          if (campos.size === 1 && !grupos.has(conds[0].campo)) grupos.set(conds[0].campo, { y: hijoY, conds });
          else for (const c of conds) marcar(c.campo);   // un grupo con varias columnas no se puede llevar a criterios por columna
        }
      }
    };
    if (raiz.local === 'filter-condition') poner(condicion(raiz), true);
    else if (raiz.local === 'filter-and') recorrer(raiz, true);
    else {   // un O de primer nivel: solo vale si todas son de la misma columna
      const conds = raiz.hijos.filter((x) => x.local === 'filter-condition').map(condicion);
      if (conds.length === raiz.hijos.length && new Set(conds.map((c) => c.campo)).size === 1) grupos.set(conds[0].campo, { y: false, conds });
      else for (const x of raiz.hijos) if (x.local === 'filter-condition') marcar(condicion(x).campo);
    }
    for (const [campo, g] of grupos) m.set(campo, g.rota ? { tipo: 'otro' } : traducir(g));
    return m;
  }

  // ---- escribir: criterios de Excel → table:filter ------------------------------------------------------------------------------------------------
  const atributo = (n, v) => ' ' + n + '="' + X().escAttr(String(v)) + '"';
  const nodoCondicion = (col, op, valor, numero, items) => '<table:filter-condition' + (numero ? atributo('table:data-type', 'number') : '') + atributo('table:value', valor) + atributo('table:operator', op) + atributo('table:field-number', col)
    + (items && items.length ? '>' + items.map((v) => '<table:filter-set-item' + atributo('table:value', v) + '/>').join('') + '</table:filter-condition>' : '/>');
  // «x*», «*x», «*x*» y el texto sin comodines → el operador de ODF (~ escapa un comodín); null si hay comodines en medio
  function patron(v) {
    let i = 0, fin = false, cuerpo = '', fino = v.length;
    const ini = v[0] === '*';
    if (ini) i = 1;
    if (fino > i && v[fino - 1] === '*') { let k = fino - 2, t = 0; while (k >= i && v[k] === '~') { t++; k--; } if (t % 2 === 0) { fin = true; fino--; } }
    for (let j = i; j < fino; j++) {
      const ch = v[j];
      if (ch === '~' && j + 1 < fino) { cuerpo += v[++j]; continue; }
      if (ch === '*' || ch === '?') return null;
      cuerpo += ch;
    }
    return { ini, fin, cuerpo };
  }
  function aOperadorOds(operador, valor) {
    const v = String(valor);
    if (operador !== 'equal' && operador !== 'notEqual') return { op: OPERADORES_ODF[operador], valor: v };
    const igual = operador === 'equal', p = patron(v);
    if (!p) throw Err('el filtro «' + v + '» usa comodines que no se pueden guardar en .ods: solo vale «x*», «*x» o «*x*»');
    if (!p.ini && !p.fin) return { op: igual ? '=' : '!=', valor: p.cuerpo };
    if (igual) return { op: p.ini && p.fin ? 'contains' : p.fin ? 'begins-with' : 'ends-with', valor: p.cuerpo };
    if (p.ini && p.fin) return { op: 'does-not-contain', valor: p.cuerpo };
    throw Err('el filtro «' + v + '» no se puede guardar en .ods: usa «no contiene» o un texto exacto');
  }
  function xmlColumna(col, cr) {
    if (cr.tipo === 'valores') {
      const vals = [...new Set(cr.valores.map(String))];
      if (!vals.length && !cr.vacias) throw Err('el filtro necesita al menos un valor');
      const vacia = nodoCondicion(col, 'empty', '', false);
      if (!vals.length) return vacia;
      const principal = nodoCondicion(col, '=', vals[0], false, vals.length > 1 ? vals : null);
      return cr.vacias ? '<table:filter-or>' + principal + vacia + '</table:filter-or>' : principal;
    }
    if (cr.tipo === 'condiciones') {
      const cs = cr.condiciones.filter((c) => F()._interno.OPERADORES.has(c.operador) && String(c.valor).trim() !== '').slice(0, 2);
      if (!cs.length) throw Err('el filtro necesita un valor');
      const partes = cs.map((c) => { const o = aOperadorOds(c.operador, c.valor); return nodoCondicion(col, o.op, o.valor, false); });
      return partes.length > 1 && !cr.y ? '<table:filter-or>' + partes.join('') + '</table:filter-or>' : partes.join('');
    }
    if (cr.tipo === 'top10') {
      const n = Math.max(1, Math.min(500, Math.trunc(+cr.n || 10)));
      return nodoCondicion(col, (cr.abajo ? 'bottom ' : 'top ') + (cr.porcentaje ? 'percent' : 'values'), n, true);
    }
    throw Err('criterio de filtro desconocido');
  }
  function xmlFiltro(criterios) {
    const cols = [...criterios.keys()].sort((a, b) => a - b);
    if (!cols.length) return '';
    const partes = cols.map((c) => xmlColumna(c, criterios.get(c)));
    // una sola columna con «o»: como la escribe LibreOffice, el grupo O es el de primer nivel
    if (partes.length === 1 && partes[0].startsWith('<table:filter-or>')) return '<table:filter>' + partes[0] + '</table:filter>';
    return '<table:filter><table:filter-and>' + partes.join('') + '</table:filter-and></table:filter>';
  }

  // el elemento database-range con otra dirección, otros botones o otro filtro; lo demás (orden, subtotales, origen) se conserva
  function elementoNuevo(texto, b, destino, filtro) {
    let ap = U().sinCierre(texto.slice(b.ini, b.finApertura));
    if (destino) ap = X().ponerAtributo(ap, 'table:target-range-address', destino);
    ap = X().ponerAtributo(ap, 'table:display-filter-buttons', 'true');
    const hijos = b.nodo.iniCierre < 0 ? [] : b.nodo.hijos.filter((x) => x.local !== 'filter').map((x) => ({ local: x.local, texto: texto.slice(b.base + x.ini, b.base + x.fin) }));
    const k = hijos.findIndex((x) => x.local === 'sort' || x.local === 'subtotal-rules');
    if (filtro) hijos.splice(k < 0 ? hijos.length : k, 0, { local: 'filter', texto: filtro });
    return hijos.length ? ap + hijos.map((x) => x.texto).join('') + '</table:database-range>' : ap.replace(/>$/, '/>');
  }
  function sinFiltro(texto, b, d, eds) {
    const anonimo = /^__Anonymous_Sheet_DB__/.test(b.attrs['table:name'] || '');
    const otros = b.nodo.iniCierre >= 0 && b.nodo.hijos.some((x) => x.local !== 'filter');
    if (anonimo && !otros) {   // el rango anónimo que solo servía de filtro se va con él
      const hermanos = bases(texto, d).filter((x) => x.cont.ini === b.cont.ini);
      if (hermanos.length === 1) eds.push({ ini: b.cont.ini, fin: b.cont.fin, texto: '' }); else eds.push({ ini: b.ini, fin: b.fin, texto: '' });
      return;
    }
    let ap = U().sinCierre(texto.slice(b.ini, b.finApertura));
    ap = X().ponerAtributo(ap, 'table:display-filter-buttons', null);
    const hijos = b.nodo.iniCierre < 0 ? [] : b.nodo.hijos.filter((x) => x.local !== 'filter').map((x) => texto.slice(b.base + x.ini, b.base + x.fin));
    eds.push({ ini: b.ini, fin: b.fin, texto: hijos.length ? ap + hijos.join('') + '</table:database-range>' : ap.replace(/>$/, '/>') });
  }
  function contenedorNuevo(texto, d, elemento) {
    const existente = X().hijos(d.hojaCalc, 'database-ranges')[0];
    if (existente && existente.iniCierre >= 0) return { ini: existente.iniCierre, fin: existente.iniCierre, texto: elemento };
    if (existente) return { ini: existente.ini, fin: existente.fin, texto: '<table:database-ranges>' + elemento + '</table:database-ranges>' };
    // después de la última hoja y de los nombres del libro
    let pos = d.tablas[d.tablas.length - 1].fin;
    for (const n of X().hijos(d.hojaCalc, 'named-expressions')) pos = Math.max(pos, n.fin);
    return { ini: pos, fin: pos, texto: '<table:database-ranges>' + elemento + '</table:database-ranges>' };
  }

  // ---- filas: table:visibility="filter" en las ocultas ---------------------------------------------------------------------------------------------
  // `ocultas`: Map(fila → true|false); las filas que no están no se tocan. Las visibles pierden también un «collapse» (como en Excel, filtrar las vuelve a mostrar).
  function edicionesFilas(texto, tabla, ocultas) {
    if (!ocultas.size) return [];
    const eds = [];
    let lo = Infinity, hi = -Infinity;
    for (const f of ocultas.keys()) { if (f < lo) lo = f; if (f > hi) hi = f; }
    for (const el of U().enumerarFilas(texto, tabla)) {
      // solo se recorren las filas del elemento que caen entre la primera y la última de `ocultas` (una repetición puede cubrir millones de filas)
      const k0 = Math.max(0, lo - el.f0), k1 = Math.min(el.n - 1, hi - el.f0);
      if (k0 > k1) continue;
      const tramos = [];   // [estado, cuántas]
      const poner = (e, n) => { const u = tramos[tramos.length - 1]; if (u && u[0] === e) u[1] += n; else tramos.push([e, n]); };
      if (k0 > 0) poner('igual', k0);
      let cambia = false;
      const visibilidad = el.nodo.attrs['table:visibility'];
      for (let k = k0; k <= k1; k++) {
        const f = el.f0 + k;
        if (!ocultas.has(f)) { poner('igual', 1); continue; }
        const e = ocultas.get(f) ? 'filter' : 'visible';
        if (e === 'filter' ? visibilidad !== 'filter' : visibilidad === 'filter' || visibilidad === 'collapse') cambia = true;
        poner(e, 1);
      }
      if (k1 < el.n - 1) poner('igual', el.n - 1 - k1);
      if (!cambia) continue;
      const abre = texto.slice(el.nodo.ini, el.nodo.finApertura), resto = texto.slice(el.nodo.finApertura, el.nodo.fin);
      const piezas = [];
      for (const [estado, n] of tramos) {
        let a = X().ponerAtributo(abre, 'table:number-rows-repeated', n > 1 ? String(n) : null);
        if (estado !== 'igual') a = X().ponerAtributo(a, 'table:visibility', estado === 'filter' ? 'filter' : null);
        piezas.push(a + resto);
      }
      eds.push({ ini: el.nodo.ini, fin: el.nodo.fin, texto: piezas.join('') });
    }
    return eds;
  }

  // ---- operaciones -------------------------------------------------------------------------------------------------------------------------------------
  function filtroEn(hoja, f, c) {
    const fa = hoja.filtroAuto && globalThis.LibroXlsx.decodificarRango(hoja.filtroAuto.ref || '');
    if (fa && c >= fa.c0 && c <= fa.c1 && f >= fa.f0 && f <= fa.f1) return { tipo: 'hoja', ref: fa };
    return null;
  }

  async function ponerAutofiltro(bytes, h, r) {
    return I().conPaquete(bytes, async (p) => {
      const texto = await p.leer('content.xml');
      const d = I().documento(texto), tabla = I().tablaDe(d, h), nombre = I().nombreDe(tabla);
      const actual = buscar(texto, d, h);
      const eds = [];
      if (!r) {
        if (!actual) return;
        sinFiltro(texto, actual, d, eds);
        const viejo = actual.destino.r, mostrar = new Map();
        for (let f = viejo.f0 + 1; f <= viejo.f1; f++) mostrar.set(f, false);
        eds.push(...edicionesFilas(texto, tabla, mostrar));
        p.escribir('content.xml', X().aplicar(texto, eds));
        return;
      }
      if (r.f1 <= r.f0) throw Err('el filtro necesita un encabezado y al menos una fila');
      const destino = destinoOds(nombre, r);
      if (actual) {
        // otro rango: se quitan los criterios y se vuelven a mostrar las filas que ocultaba el anterior
        eds.push({ ini: actual.ini, fin: actual.fin, texto: elementoNuevo(texto, actual, destino, '') });
        const mostrar = new Map();
        for (let f = actual.destino.r.f0 + 1; f <= actual.destino.r.f1; f++) mostrar.set(f, false);
        eds.push(...edicionesFilas(texto, tabla, mostrar));
      } else {
        const usados = new Set(bases(texto, d).map((b) => b.attrs['table:name']));
        let n = h; while (usados.has('__Anonymous_Sheet_DB__' + n)) n++;
        eds.push(contenedorNuevo(texto, d, '<table:database-range table:name="__Anonymous_Sheet_DB__' + n + '" table:target-range-address="' + X().escAttr(destino) + '" table:display-filter-buttons="true"/>'));
      }
      p.escribir('content.xml', X().aplicar(texto, eds));
    });
  }

  async function criterios(bytes, h) {
    const p = await PaqueteEditable.abrir(bytes);
    const texto = await p.leer('content.xml');
    const d = I().documento(texto);
    const b = buscar(texto, d, h);
    return b ? leerFiltro(b.nodo) : new Map();
  }

  async function filtrar(bytes, h, objetivo, colId, criterio, textoDe) {
    return I().conPaquete(bytes, async (p) => {
      const texto = await p.leer('content.xml');
      const d = I().documento(texto), tabla = I().tablaDe(d, h);
      const b = buscar(texto, d, h);
      if (!b) throw Err('no hay autofiltro');
      const r = b.destino.r;
      if (colId < 0 || colId > r.c1 - r.c0) throw Err('columna fuera del filtro');
      const lista = leerFiltro(b.nodo);
      if ([...lista.values()].some((c) => c.tipo === 'otro')) throw Err('el filtro de esta hoja tiene condiciones que solo LibreOffice sabe editar (expresiones regulares, orden…): bórralo allí');
      if (criterio) lista.set(colId, criterio); else lista.delete(colId);
      const eds = [{ ini: b.ini, fin: b.fin, texto: elementoNuevo(texto, b, null, xmlFiltro(lista)) }];
      eds.push(...edicionesFilas(texto, tabla, F()._interno.calcularOcultas(lista, r, textoDe, r.f1)));
      p.escribir('content.xml', X().aplicar(texto, eds));
    });
  }

  globalThis.FiltrosOds = { criterios, ponerAutofiltro, filtroEn, filtrar, _interno: { leerFiltro, xmlFiltro, aOperadorOds, patron, simple } };
})();

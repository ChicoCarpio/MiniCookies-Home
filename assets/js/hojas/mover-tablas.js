/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Cortar y pegar en .xlsx: las tablas. Medido en Excel (pruebas/oraculo/excel-mover-limites.ps1):
// - Una tabla entera dentro de lo movido se mueve, también a otra hoja, y pierde su filtro (se ven las filas que ocultaba; igual el autofiltro de la hoja). Si llega encima de otra tabla entera, esa se pierde; si la toca a medias, la que se pierde es la movida.
// - Lo demás son celdas comunes, y la tabla se queda donde está: un encabezado que se va o queda vacío pasa a ColumnaN; lo pegado en un encabezado es el nombre nuevo (como texto), y otra columna
//   con ese nombre pasa a nombre2. Las fórmulas siguen a la columna, no al nombre.
// - La fila de totales: lo que se va deja la columna sin total, y lo pegado queda como etiqueta o fórmula propia.
// - Una columna calculada sin ninguna de sus celdas originales deja de serlo.
// - Pegar junto a la tabla la hace crecer: debajo si no hay fila de totales, a la derecha también con ella, siempre que lo pegado no se salga de su ancho o su alto, traiga algo y lo demás de esa
//   fila o columna esté vacío. Una columna calculada se llena en las filas nuevas, y los rangos que abarcan una columna hasta su última fila crecen con ella.
// - Escribir, borrar y copiar y pegar en el encabezado, los totales o junto a la tabla hacen lo mismo (medido en pruebas/oraculo/excel-tablas-editar.ps1): el cambio va sobre el archivo.
// - El nombre de un encabezado que queda vacío (excel-medir-encabezados.ps1): al borrar o escribir, el primer ColumnaN libre; al pegar (cortando o copiando), el que sigue
//   al de su izquierda si termina en número (x5 → x6, Columna01 → Columna02) y, si ya está, el de la izquierda con 2, 3…; si no, el primer ColumnaN libre.
//   MoverTablas.planear(ed, m, nombresHoja, idioma, modo) → plan (aplica las celdas en ed) | null;  await MoverTablas.escribir(libro, plan, M, ft)   (libro: LibroXlsx abierto; escribe en su zip)
// - Dos encabezados con el mismo nombre: se resuelve de izquierda a derecha, y el de la derecha pasa a nombre2 (o el número que siga libre en toda la fila).
// - Una fila de solo fórmulas que llega justo debajo de una tabla sin totales (al escribir, copiar o cortar) es su fila de totales; SUBTOTAL(109,[su columna]) queda como «suma».
// - Borrar (Supr) la tabla entera, con su fila de totales, la quita: lo que la usaba queda en #REF!.
// - Una sola fórmula, escrita o pegada, en una columna de la tabla que no tiene nada más (también la que acaba de crecer) la vuelve calculada: se llena entera.
//   Pegar encima de todas las celdas de una columna calculada le quita la fórmula.
//   modo: 'mover' (por omisión), 'pegar', 'escribir' o 'borrar'. MoverTablas.toca(ed, h, r, formula) → ¿el cambio de r necesita las tablas?
//   await MoverTablas.guardar(ed, h, r, modo, idioma) → bytes (con las celdas ya cambiadas en ed);  MoverTablas.nombresYGraficos(zip, f) (los nombres definidos y los gráficos, por f)
(function () {
  const X = () => XmlEditar;
  const REL_TABLA = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table';
  const toca = (a, b) => !(a.f1 < b.f0 || a.f0 > b.f1 || a.c1 < b.c0 || a.c0 > b.c1);
  const dentro = (a, b) => a.f0 >= b.f0 && a.f1 <= b.f1 && a.c0 >= b.c0 && a.c1 <= b.c1;
  const textoRango = (r) => MoverObjetos._textoRango(r);
  const escapar = (n) => String(n).replace(/(['\[\]#])/g, "'$1");
  const sinEscape = (n) => String(n).replace(/'(.)/g, '$1');

  // Las referencias a una tabla dentro de una fórmula: cambia sus columnas o, si la tabla ya no existe, pone #REF!
  function enTabla(f, nombre, columna, quitada) {
    const texto = String(f), bajo = nombre.toLowerCase();
    let i = 0, salida = '';
    while (i < texto.length) {
      if (texto[i] === '"') { let j = i + 1; while (j < texto.length) { if (texto[j] === '"') { if (texto[j + 1] === '"') j += 2; else break; } else j++; } salida += texto.slice(i, j + 1); i = j + 1; continue; }
      const antes = i ? texto[i - 1] : '';
      if (texto.substr(i, nombre.length).toLowerCase() === bajo && texto[i + nombre.length] === '[' && !/[A-Za-z0-9_.\\À-￿]/.test(antes)) {
        let j = i + nombre.length, nivel = 0;
        for (; j < texto.length; j++) { const c = texto[j]; if (c === "'") { j++; continue; } if (c === '[') nivel++; else if (c === ']' && --nivel === 0) break; }
        const dentroTxt = texto.slice(i + nombre.length, j + 1);
        if (quitada) salida += '#REF!';
        else {
          const uno = (p) => { if (p[0] === '#') return p; const n = columna(sinEscape(p)); return n === null ? p : escapar(n); };
          const inner = dentroTxt.slice(1, -1);
          salida += texto.slice(i, i + nombre.length) + '[' + (inner[0] === '[' ? inner.replace(/\[((?:[^\[\]']|'.)*)\]/g, (x, p) => '[' + uno(p) + ']') : uno(inner)) + ']';
        }
        i = j + 1;
        continue;
      }
      salida += texto[i++];
    }
    return salida;
  }
  // Los rangos que abarcan una columna de la tabla hasta su última fila crecen con ella (y una sola fila o celda, no)
  function extender(f, h, cambios, nombresHoja) {
    const tapado = String(f).replace(/"(?:[^"]|"")*"|\[[^\]]*\]/g, (m) => (m[0] === '"' ? m : '_'.repeat(m.length)));
    let toks;
    try { toks = XlsFormulas._tokenizar(tapado); } catch (e) { return f; }
    let s = String(f);
    for (let k = toks.length - 1; k >= 0; k--) {
      const t = toks[k];
      if (t.k !== 'ref' || t.modo !== 'area' || (t.hoja && t.hoja.a !== t.hoja.b)) continue;
      const hr = t.hoja ? nombresHoja.findIndex((n) => n.toLowerCase() === t.hoja.a.toLowerCase()) : h;
      const c = cambios.find((x) => x.h === hr && t.c0 >= x.c0 && t.c1 <= x.c1 && t.f1 === x.ultima && (t.f0 === x.arriba || t.f0 === x.datos));
      if (!c) continue;
      const o = s.slice(t.i, t.i + t.l);
      s = s.slice(0, t.i) + o.replace(/(\$?)(\d+)$/, (x, a) => a + (c.nueva + 1)) + s.slice(t.i + t.l);
    }
    return s;
  }

  function planear(ed, m, nombresHoja, idioma, modo) {
    const hojas = ed.hojas, M = m.M, editar = !!modo && modo !== 'mover', serie = !editar || modo === 'pegar';
    if (!hojas.some((x) => x && !x.sinCeldas && (x.tablas || []).length)) return null;
    const S = { f0: M.f0, c0: M.c0, f1: M.f1, c1: M.c1 }, D = { f0: M.f0 + M.df, c0: M.c0 + M.dc, f1: M.f1 + M.df, c1: M.c1 + M.dc };
    const nombreDe = (t) => t.nombreVisible || t.nombre;
    const palabra = idioma === 'en' ? 'Column' : 'Columna';
    const vacia = (h, f, c) => { const x = ed.estadoDe(h, f, c); return !x || (!x.formula && (x.tipo === 'vacio' || x.valor === null || x.valor === undefined || x.valor === '')); };
    // un número o un valor lógico que llega a un encabezado se queda con el texto que se ve, con su formato (medido: 8 con formato «0.08» → «0.08»)
    const comoSeVe = (x) => {
      if (typeof x.valor !== 'number' && typeof x.valor !== 'boolean') return String(x.valor === null || x.valor === undefined ? '' : x.valor);
      try { return XlsxFormato.formatear(x.valor, (ed.libro.estilos.xf(x.estilo || 0) || {}).formato || 'General', { idioma: idioma === 'en' ? 'en' : 'es', fecha1904: ed.libro.fecha1904 }).texto; } catch (e) { return String(x.valor); }
    };
    const afectada = (h, f, c) => { const p = { f0: f, f1: f, c0: c, c1: c }; return (h === M.h0 && dentro(p, S)) || (h === M.h1 && dentro(p, D)); };
    const plan = { tablas: new Map(), quitar: new Set(), renombres: [], cambios: [], formula: null };
    const movidas = editar ? [] : (hojas[M.h0].tablas || []).filter((t) => dentro(t.ref, S));
    const fijas = (hojas[M.h1].tablas || []).filter((t) => !(M.h0 === M.h1 && movidas.includes(t)));
    let sigue = movidas.slice();
    if (movidas.length) for (const o of fijas) {
      if (!toca(o.ref, D)) continue;
      if (dentro(o.ref, D)) plan.quitar.add(nombreDe(o));                      // queda entera debajo: se pierde
      else sigue = [];                                                         // a medias: la que se pierde es la movida
    }
    if (movidas.length && !sigue.length) for (const t of movidas) plan.quitar.add(nombreDe(t));
    if (modo === 'borrar') for (const t of hojas[M.h1].tablas || []) if (dentro(t.ref, D)) plan.quitar.add(nombreDe(t));
    for (const t of sigue) plan.tablas.set(nombreDe(t), { ref: { f0: t.ref.f0 + M.df, c0: t.ref.c0 + M.dc, f1: t.ref.f1 + M.df, c1: t.ref.c1 + M.dc }, hojaVieja: M.h0, hojaNueva: M.h1, movida: true });
    const celdas = new Map();
    const poner = (h, f, c, v) => { if (!celdas.has(h)) celdas.set(h, []); celdas.get(h).push({ f, c, v }); };

    hojas.forEach((hoja, h) => {
      if (!hoja || hoja.sinCeldas) return;
      for (const t of hoja.tablas || []) {
        const n = nombreDe(t);
        if (plan.quitar.has(n) || (movidas.includes(t) && h === M.h0)) continue;
        if (!(h === M.h0 || h === M.h1)) continue;
        const r = Object.assign({}, t.ref), viejas = t.nombresColumnas || [];
        if (!viejas.length) continue;
        // crecer
        let f1 = r.f1, c1 = r.c1, filaTotales = -1;
        const otras = (hoja.tablas || []).filter((o) => o !== t);
        if (h === M.h1 && !movidas.length) {
          const solapa = toca(D, r);
          if (!t.totales && D.f1 > r.f1 && (solapa || (D.f0 === r.f1 + 1 && D.c0 >= r.c0 && D.c1 <= r.c1))) {
            const zona = { f0: r.f1 + 1, f1: D.f1, c0: r.c0, c1: r.c1 };
            let algo = false, libre = true, formulas = true;
            for (let i = zona.f0; i <= zona.f1 && libre; i++) for (let j = zona.c0; j <= zona.c1 && libre; j++) { const v = vacia(h, i, j), enD = i >= D.f0 && i <= D.f1 && j >= D.c0 && j <= D.c1; if (enD && !v) { algo = true; if (!ed.estadoDe(h, i, j).formula) formulas = false; } if (!enD && !v) libre = false; }
            if (algo && libre && !otras.some((o) => toca(o.ref, zona))) { f1 = D.f1; if (zona.f0 === zona.f1 && formulas) filaTotales = f1; }
          }
          if (D.c1 > r.c1 && (solapa || (D.c0 === r.c1 + 1 && D.f0 >= r.f0 && D.f1 <= r.f1))) {
            const zona = { f0: r.f0, f1: r.f1, c0: r.c1 + 1, c1: D.c1 };
            let algo = false, libre = true;
            for (let i = zona.f0; i <= zona.f1 && libre; i++) for (let j = zona.c0; j <= zona.c1 && libre; j++) { const v = vacia(h, i, j), enD = i >= D.f0 && i <= D.f1 && j >= D.c0 && j <= D.c1; if (enD && !v) algo = true; if (!enD && !v) libre = false; }
            if (algo && libre && !otras.some((o) => toca(o.ref, zona))) c1 = D.c1;
          }
        }
        const nuevaRef = { f0: r.f0, c0: r.c0, f1, c1 };
        const anchas = c1 - r.c0 + 1;
        // encabezados
        const nombres = [];
        if (t.encabezados) {
          const estados = [];
          for (let k = 0; k < anchas; k++) {
            const c = r.c0 + k, tocada = afectada(h, r.f0, c) || k >= viejas.length;
            const x = ed.estadoDe(h, r.f0, c);
            let txt = null;
            if (tocada && !vacia(h, r.f0, c)) txt = (x.tipo === 'error' ? String(x.valor) : comoSeVe(x)).slice(0, 255);   // Excel no deja más de 255 caracteres en un encabezado
            estados.push({ tocada, txt, vieja: viejas[k] });
          }
          const usados = new Set(), propio = (e) => (e.tocada ? e.txt : e.vieja);
          estados.forEach((e, k) => {
            const b = propio(e);
            if (!b) return;
            if (!usados.has(b.toLowerCase())) e.nombre = b;
            else {
              const despues = new Set(estados.slice(k + 1).map((o) => String(propio(o) || '').toLowerCase()));
              let z = 2;
              while (usados.has((b + z).toLowerCase()) || despues.has((b + z).toLowerCase())) z++;
              e.nombre = b + z;
            }
            usados.add(e.nombre.toLowerCase());
          });
          const primeroLibre = () => { let z = 1; while (usados.has((palabra + z).toLowerCase())) z++; return palabra + z; };
          estados.forEach((e, k) => {
            if (e.nombre !== undefined) return;
            const izq = k ? estados[k - 1].nombre : undefined, q = serie && izq !== undefined ? /^([\s\S]*?)(\d+)$/.exec(izq) : null;
            if (q) {
              const sig = q[1] + (BigInt(q[2]) + 1n).toString().padStart(q[2].length, '0');
              if (!usados.has(sig.toLowerCase())) e.nombre = sig;
              else { let z = 2; while (usados.has((izq + z).toLowerCase())) z++; e.nombre = izq + z; }
            } else e.nombre = primeroLibre();
            usados.add(e.nombre.toLowerCase());
          });
          const mapa = new Map();
          estados.forEach((e, k) => {
            nombres.push(e.nombre);
            if (e.vieja !== undefined && e.vieja !== e.nombre) mapa.set(e.vieja.toLowerCase(), e.nombre);
            const x = ed.estadoDe(h, r.f0, r.c0 + k);
            if (!x || x.tipo !== 'texto' || x.valor !== e.nombre || x.formula) poner(h, r.f0, r.c0 + k, { tipo: 'texto', valor: e.nombre, formula: null, estilo: x ? x.estilo : 0 });
          });
          if (mapa.size) plan.renombres.push({ nombre: n, mapa });
        } else for (let k = 0; k < anchas; k++) nombres.push(viejas[k] !== undefined ? viejas[k] : palabra + (k + 1));
        // fila de totales
        const totales = new Map();
        if (filaTotales >= 0) for (let k = 0; k < anchas; k++) {
          const x = ed.estadoDe(h, filaTotales, r.c0 + k);
          if (x && x.formula) totales.set(k, funcionTotal(x.formula, n, nombres[k]) || { formula: x.formula });
        }
        if (t.totales) for (let k = 0; k < anchas; k++) {
          const c = r.c0 + k;
          if (!afectada(h, r.f1, c)) continue;
          const x = ed.estadoDe(h, r.f1, c);
          totales.set(k, vacia(h, r.f1, c) ? { nada: true } : x.formula ? { formula: x.formula } : { etiqueta: String(x.valor) });
          if (!vacia(h, r.f1, c) && !x.formula && x.tipo !== 'texto') poner(h, r.f1, c, { tipo: 'texto', valor: String(x.valor), formula: null, estilo: x.estilo });
        }
        // columnas calculadas: todas sus celdas se fueron o quedaron tapadas
        const datosF0 = r.f0 + (t.encabezados ? 1 : 0), datosF1 = r.f1 - (t.totales ? 1 : 0);
        const sinCalc = new Set();
        for (let k = 0; k < viejas.length; k++) {
          let todas = datosF0 <= datosF1;
          for (let i = datosF0; todas && i <= datosF1; i++) todas = afectada(h, i, r.c0 + k);
          if (todas) sinCalc.add(k);
        }
        const crece = f1 > r.f1 && filaTotales < 0;
        // una sola fórmula en una columna sin nada más la vuelve calculada
        const nuevasCalc = new Map();
        if (editar && modo !== 'borrar' && h === M.h1 && D.f0 === D.f1 && D.c0 === D.c1) {
          const k = D.c0 - r.c0, x = ed.estadoDe(h, D.f0, D.c0), hasta = crece ? f1 : datosF1;
          if (k >= 0 && k < anchas && D.f0 >= datosF0 && D.f0 <= hasta && x && x.formula && !(t.calculadas && t.calculadas.has(k))) {
            let sola = true;
            for (let i = datosF0; sola && i <= hasta; i++) if (i !== D.f0 && !vacia(h, i, D.c0)) sola = false;
            if (sola) {
              nuevasCalc.set(k, LibroXlsx.desplazarFormula(x.formula, datosF0 - D.f0, 0));
              for (let i = datosF0; i <= hasta; i++) if (i !== D.f0) { const y = ed.estadoDe(h, i, D.c0); poner(h, i, D.c0, { tipo: 'vacio', valor: null, formula: LibroXlsx.desplazarFormula(x.formula, i - D.f0, 0), estilo: y ? y.estilo : 0 }); }
            }
          }
        }
        plan.tablas.set(n, { ref: nuevaRef, hojaVieja: h, hojaNueva: h, nombres, totales, sinCalc, datosF1, calculadas: t.calculadas || new Map(), autollenar: crece ? { desde: r.f1 + 1, hasta: f1, datosF0 } : null, h, nuevosTotales: filaTotales >= 0, nuevasCalc });
        if (crece) plan.cambios.push({ h, c0: r.c0, c1: r.c1, ultima: r.f1, arriba: r.f0, datos: datosF0, nueva: f1 });
        t.ref = nuevaRef; t.nombresColumnas = nombres;
        if (filaTotales >= 0) t.totales = true;
        if (nuevasCalc.size) { t.calculadas = new Map(t.calculadas || []); for (const [k, fx] of nuevasCalc) t.calculadas.set(k, fx); }
      }
    });
    for (const t of sigue) { t.ref = plan.tablas.get(nombreDe(t)).ref; }
    if (M.h0 !== M.h1 && sigue.length) {   // el motor busca las tablas por hoja
      hojas[M.h0].tablas = hojas[M.h0].tablas.filter((t) => !sigue.includes(t));
      hojas[M.h1].tablas = (hojas[M.h1].tablas || []).concat(sigue);
    }
    for (const nombre of plan.quitar) hojas.forEach((x) => { if (x && x.tablas) x.tablas = x.tablas.filter((t) => nombreDe(t) !== nombre); });
    plan.formula = (f, h) => {
      let u = f;
      for (const nombre of plan.quitar) u = enTabla(u, nombre, null, true);
      for (const r of plan.renombres) u = enTabla(u, r.nombre, (c) => { const v = r.mapa.get(c.toLowerCase()); return v === undefined ? null : v; }, false);
      return plan.cambios.length ? extender(u, h, plan.cambios, nombresHoja) : u;
    };
    // lo que cambia en las celdas
    const hayNombres = plan.quitar.size || plan.renombres.length || plan.cambios.length;
    if (hayNombres) hojas.forEach((hoja, h) => {
      if (!hoja || hoja.sinCeldas) return;
      const C = hoja.celdas;
      for (const [k] of C.formulas) {
        const i = C.fila[k], j = C.col[k], x = ed.estadoDe(h, i, j);
        if (!x || !x.formula) continue;
        const u = plan.formula(x.formula, h);
        if (u === x.formula) continue;
        const v = Object.assign({}, x, { formula: u });
        delete v.matriz;
        poner(h, i, j, v);
      }
    });
    // una columna calculada se llena en las filas nuevas (lo pegado ahí se queda)
    for (const p of plan.tablas.values()) {
      if (!p.autollenar) continue;
      for (const [k, formula] of p.calculadas) {
        if (p.sinCalc.has(k)) continue;
        const c = p.ref.c0 + k;
        for (let i = p.autollenar.desde; i <= p.autollenar.hasta; i++) if (vacia(p.h, i, c)) poner(p.h, i, c, { tipo: 'vacio', valor: null, formula: LibroXlsx.desplazarFormula(formula, i - p.autollenar.datosF0, 0), estilo: 0 });
      }
    }
    plan.celdas = celdas;
    for (const [h, l] of celdas) if (l.length) ed.aplicar({ t: 'celdas', h, celdas: l });
    return plan;
  }

  // SUBTOTAL(10x, [la columna]) en una fila de totales nueva es una de las funciones de la lista de Excel (medido: 109 → sum)
  const FUNCIONES_TOTAL = { 101: 'average', 102: 'countNums', 103: 'count', 104: 'max', 105: 'min', 107: 'stdDev', 109: 'sum', 110: 'var' };
  function funcionTotal(formula, tabla, columna) {
    const m = /^\s*SUBTOTAL\(\s*(\d+)\s*,\s*((?:[^\[\]]*)\[((?:[^\[\]']|'.)*)\])\s*\)\s*$/i.exec(String(formula));
    if (!m || !FUNCIONES_TOTAL[m[1]] || columna === undefined) return null;
    const de = m[2].slice(0, m[2].indexOf('['));
    if (de && de.toLowerCase() !== String(tabla).toLowerCase()) return null;
    return sinEscape(m[3]).toLowerCase() === String(columna).toLowerCase() ? { funcion: FUNCIONES_TOTAL[m[1]] } : null;
  }
  function nombreAttr(n) { return X().escAttr(n); }
  function tablaXml(xml, p, ft, formula, h) {
    let u = xml;
    const ref = textoRango(p.ref), r = p.ref;
    u = u.replace(/(<table\b[^>]*?\bref\s*=\s*")[^"]*(")/, (x, a, c) => a + ref + c);
    if (p.nuevosTotales) u = u.replace(/<table\b[^>]*>/, (e) => e.replace(/\s+totalsRow(?:Shown|Count)\s*=\s*"[^"]*"/g, '').replace(/^<table\b/, '<table totalsRowCount="1"'));
    const totalesFila = /\btotalsRowCount\s*=\s*"[1-9]/.test((/<table\b[^>]*>/.exec(u) || [''])[0]);
    const filtro = textoRango({ f0: r.f0, c0: r.c0, f1: r.f1 - (totalesFila ? 1 : 0), c1: r.c1 });
    u = u.replace(/(<autoFilter\b[^>]*?\bref\s*=\s*")[^"]*(")/, (x, a, c) => a + filtro + c);
    if (p.movida) {
      const M = p.M;
      const corrida = (t) => { const g = LibroXlsx.decodificarRango(t); return g ? textoRango({ f0: g.f0 + M.df, c0: g.c0 + M.dc, f1: g.f1 + M.df, c1: g.c1 + M.dc }) : t; };
      u = u.replace(/(<(?:sortState|sortCondition)\b[^>]*?\bref\s*=\s*")([^"]*)(")/g, (x, a, t, c) => a + corrida(t) + c);
      // movida entera, pierde el filtro (medido en Excel); las filas que ocultaba se muestran en escribir()
      u = u.replace(/<filterColumn\b[^>]*?(?:\/>|>[\s\S]*?<\/filterColumn>)/g, '').replace(/(<autoFilter\b[^>]*?)\s*>\s*<\/autoFilter>/, '$1/>');
    }
    const fx = (t) => X().esc(formula(ft(X().decodificar(t), p.hojaVieja, p.hojaVieja), h));
    u = u.replace(/(<(calculatedColumnFormula|totalsRowFormula)\b[^>]*(?<!\/)>)([\s\S]*?)(<\/\2>)/g, (x, a, n, t, c) => a + fx(t) + c);
    if (!p.nombres && !p.totales && !p.sinCalc && !(p.nuevasCalc && p.nuevasCalc.size)) return u;
    const m = /<tableColumns\b[^>]*>([\s\S]*?)<\/tableColumns>/.exec(u);
    if (!m) return u;
    const cols = m[1].match(/<tableColumn\b[^>]*?(?:\/>|>[\s\S]*?<\/tableColumn>)/g) || [];
    const nuevas = cols.map((col, k) => {
      let e = col;
      if (p.nombres && p.nombres[k] !== undefined) e = e.replace(/(\bname\s*=\s*")[^"]*(")/, (x, a, c) => a + nombreAttr(p.nombres[k]) + c);
      if (p.sinCalc && p.sinCalc.has(k)) e = e.replace(/<calculatedColumnFormula\b[^>]*(?:\/>|>[\s\S]*?<\/calculatedColumnFormula>)/, '');
      const tot = p.totales && p.totales.get(k);
      if (tot) {
        e = e.replace(/\s+totalsRow(?:Function|Label)\s*=\s*"[^"]*"/g, '').replace(/<totalsRowFormula\b[^>]*(?:\/>|>[\s\S]*?<\/totalsRowFormula>)/, '');
        if (tot.etiqueta !== undefined) e = e.replace(/^<tableColumn\b/, () => '<tableColumn totalsRowLabel="' + nombreAttr(tot.etiqueta) + '"');
        if (tot.funcion !== undefined) e = e.replace(/^<tableColumn\b/, () => '<tableColumn totalsRowFunction="' + tot.funcion + '"');
        if (tot.formula !== undefined) {
          e = e.replace(/^<tableColumn\b/, '<tableColumn totalsRowFunction="custom"');
          const hijo = '<totalsRowFormula>' + X().esc(tot.formula) + '</totalsRowFormula>';
          e = /\/>$/.test(e) ? e.replace(/\/>$/, () => '>' + hijo + '</tableColumn>') : e.replace('</tableColumn>', () => hijo + '</tableColumn>');
        }
      }
      return e.replace(/<tableColumn\b([^>]*)>\s*<\/tableColumn>/, '<tableColumn$1/>');
    });
    if (p.nombres) {
      let id = 0;
      for (const c of cols) { const q = /\bid\s*=\s*"(\d+)"/.exec(c); if (q) id = Math.max(id, +q[1]); }
      for (let k = cols.length; k < p.nombres.length; k++) nuevas.push(`<tableColumn id="${++id}" name="${nombreAttr(p.nombres[k])}"/>`);
    }
    if (p.nuevasCalc) for (const [k, fx] of p.nuevasCalc) {
      if (nuevas[k] === undefined) continue;
      const hijo = '<calculatedColumnFormula>' + X().esc(fx) + '</calculatedColumnFormula>';
      nuevas[k] = /\/>$/.test(nuevas[k]) ? nuevas[k].replace(/\/>$/, () => '>' + hijo + '</tableColumn>') : nuevas[k].replace(/(<tableColumn\b[^>]*>)/, (a) => a + hijo);
    }
    return u.replace(m[0], () => m[0].slice(0, m[0].indexOf('>') + 1).replace(/\bcount\s*=\s*"\d+"/, 'count="' + nuevas.length + '"') + nuevas.join('') + '</tableColumns>');
  }

  async function escribir(libro, plan, M, ft) {
    const P = PaqueteEditable.desdeZip(libro.zip), hechas = new Set();
    for (let h = 0; h < libro.hojas.length; h++) {
      if (libro.hojas[h].tipo !== 'hoja' || !libro.hojas[h].ruta) continue;
      const ruta = libro.hojas[h].ruta, rels = await P.relaciones(ruta);
      for (const rel of rels.filter((r) => r.tipo === REL_TABLA)) {
        if (hechas.has(rel.destino)) continue;
        hechas.add(rel.destino);
        const xml = await P.leer(rel.destino);
        if (xml === null) continue;
        const n = /\bdisplayName\s*=\s*"([^"]*)"/.exec(xml), nombre = n ? X().decodificar(n[1]) : null;
        if (plan.quitar.has(nombre)) { await sacarDeHoja(P, ruta, rel.id); await MoverNotas._quitarPartes(P, [rel.destino]); await MoverNotas._quitarRelaciones(P, ruta, new Set([rel.id])); continue; }
        const p = plan.tablas.get(nombre);
        if (!p) { const u = xml.replace(/(<(calculatedColumnFormula|totalsRowFormula)\b[^>]*(?<!\/)>)([\s\S]*?)(<\/\2>)/g, (x, a, q, t, c) => a + X().esc(plan.formula(ft(X().decodificar(t), h, h), h)) + c); if (u !== xml) P.escribir(rel.destino, u); continue; }
        p.M = M;
        P.escribir(rel.destino, tablaXml(xml, p, ft, plan.formula, h));
        if (p.movida && /<filterColumn\b/.test(xml)) {
          const filas = new Map(), viejaRuta = libro.hojas[p.hojaVieja].ruta;
          for (let f = p.ref.f0 - M.df + 1; f <= p.ref.f1 - M.df; f++) filas.set(f, false);
          P.escribir(viejaRuta, FiltrosLibro._interno.ponerOcultas(await P.leer(viejaRuta), filas));
        }
        if (p.hojaNueva !== p.hojaVieja && h === p.hojaVieja) {
          await sacarDeHoja(P, ruta, rel.id);
          await MoverNotas._quitarRelaciones(P, ruta, new Set([rel.id]));
          const r1 = libro.hojas[p.hojaNueva].ruta, id = await P.agregarRelacion(r1, REL_TABLA, '../' + rel.destino.replace(/^xl\//, ''));
          await ponerEnHoja(P, r1, id);
        }
      }
    }
  }
  async function sacarDeHoja(P, ruta, id) {
    const xml = await P.leer(ruta);
    const u = xml.replace(new RegExp('<tablePart\\b[^>]*?\\bid\\s*=\\s*"' + id + '"[^>]*?/>'), '');
    const k = (u.match(/<tablePart\b/g) || []).length;
    P.escribir(ruta, k ? u.replace(/(<tableParts\b[^>]*?\bcount\s*=\s*")\d+(")/, (x, a, c) => a + k + c) : u.replace(/<tableParts\b[^>]*>\s*<\/tableParts>/, ''));
  }
  async function ponerEnHoja(P, ruta, id) {
    const con = MoverNotas._conPrefijoR(await P.leer(ruta));
    let u = con.xml;
    const parte = `<tablePart ${con.p}:id="${id}"/>`;
    if (/<tableParts\b[^>]*\/>/.test(u)) u = u.replace(/<tableParts\b[^>]*\/>/, '<tableParts count="1">' + parte + '</tableParts>');
    else if (/<tableParts\b/.test(u)) { u = u.replace('</tableParts>', () => parte + '</tableParts>'); u = u.replace(/(<tableParts\b[^>]*?\bcount\s*=\s*")\d+(")/, (x, a, c) => a + (u.match(/<tablePart\b/g) || []).length + c); }
    else u = EstructuraLibro._ponerEnHoja(u, 'tableParts', '<tableParts count="1">' + parte + '</tableParts>');
    P.escribir(ruta, u);
  }

  // ¿escribir, borrar o pegar en r toca un encabezado o unos totales, queda junto a una tabla (donde puede crecer), tapa una columna calculada entera o es una
  // sola fórmula (formula: true) en una columna sin nada más? Solo en .xlsx
  function tocaTabla(ed, h, r, formula) {
    if (ed.formato === 'ods' || !ed.hojas[h] || ed.hojas[h].sinCeldas) return false;
    const vacia = (f, c) => { const x = ed.estadoDe(h, f, c); return !x || (!x.formula && (x.tipo === 'vacio' || x.valor === null || x.valor === undefined || x.valor === '')); };
    return (ed.hojas[h].tablas || []).some((t) => {
      const g = t.ref, d0 = g.f0 + (t.encabezados ? 1 : 0), d1 = g.f1 - (t.totales ? 1 : 0);
      if (!toca(r, { f0: g.f0, c0: g.c0, f1: g.f1 + 1, c1: g.c1 }) && !toca(r, { f0: g.f0, c0: g.c0, f1: g.f1, c1: g.c1 + 1 })) return false;   // la esquina en diagonal no la hace crecer
      if (!dentro(r, { f0: d0, c0: g.c0, f1: d1, c1: g.c1 })) return true;
      if (r.f0 <= d0 && r.f1 >= d1 && [...(t.calculadas || new Map()).keys()].some((k) => g.c0 + k >= r.c0 && g.c0 + k <= r.c1)) return true;
      if (!formula || r.f0 !== r.f1 || r.c0 !== r.c1 || (t.calculadas && t.calculadas.has(r.c0 - g.c0))) return false;
      for (let i = d0; i <= d1; i++) if (i !== r.f0 && !vacia(i, r.c0)) return false;
      return true;
    });
  }
  // los nombres definidos y los gráficos: cada fórmula pasa por f
  async function nombresYGraficos(zip, f) {
    const rutas = [];
    zip.forEach((n) => { if (/^xl\/workbook\.xml$|^xl\/charts\/chart[^/]*\.xml$|^xl\/charts\/chartEx[^/]*\.xml$/i.test(n)) rutas.push(n); });
    for (const ruta of rutas) {
      const t = await zip.file(ruta).async('string');
      const u = /workbook\.xml$/i.test(ruta)
        ? t.replace(/(<definedName\b[^>]*>)([\s\S]*?)(<\/definedName>)/g, (x, a, q, c) => a + X().esc(f(X().decodificar(q))) + c)
        : t.replace(/(<(?:c|cx):f\b[^>]*>)([\s\S]*?)(<\/(?:c|cx):f>)/g, (x, a, q, c) => a + X().esc(f(X().decodificar(q))) + c);
      if (u !== t) zip.file(ruta, u);
    }
  }
  // las celdas de r ya cambiaron en ed (sin terminar): las tablas se ponen al día y se guarda
  async function guardar(ed, h, r, modo, idioma) {
    ed.terminar();
    const M = { h0: h, h1: h, f0: r.f0, c0: r.c0, f1: r.f1, c1: r.c1, df: 0, dc: 0 };
    const plan = planear(ed, { M }, ed.libro.hojas.map((x) => x.nombre), idioma, modo);
    ed.terminar();
    if (!plan) return ed.guardar();
    for (const x of ed.hojas) if (x && !x.sinCeldas) x.tablas = [];   // los encabezados y totales los escribe la tabla
    const libro = await LibroXlsx.abrir(await ed.guardar());
    await escribir(libro, plan, M, (t) => t);
    await nombresYGraficos(libro.zip, (t) => plan.formula(t, -1));
    return libro.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  globalThis.MoverTablas = { planear, escribir, toca: tocaTabla, guardar, nombresYGraficos, _enTabla: enTabla };
})();

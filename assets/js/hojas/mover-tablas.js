/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Cortar y pegar en .xlsx: las tablas. Medido en Excel (pruebas/oraculo/excel-mover-limites.ps1):
// - Una tabla entera dentro de lo movido se mueve con su filtro, también a otra hoja. Si llega encima de otra tabla entera, esa se pierde; si la toca a medias, la que se pierde es la movida.
// - Lo demás son celdas comunes, y la tabla se queda donde está: un encabezado que se va o queda vacío pasa a ColumnaN; lo pegado en un encabezado es el nombre nuevo (como texto), y otra columna
//   con ese nombre pasa a nombre2. Las fórmulas siguen a la columna, no al nombre.
// - La fila de totales: lo que se va deja la columna sin total, y lo pegado queda como etiqueta o fórmula propia.
// - Una columna calculada sin ninguna de sus celdas originales deja de serlo.
// - Pegar junto a la tabla la hace crecer: debajo si no hay fila de totales, a la derecha también con ella, siempre que lo pegado no se salga de su ancho o su alto, traiga algo y lo demás de esa
//   fila o columna esté vacío. Una columna calculada se llena en las filas nuevas, y los rangos que abarcan una columna hasta su última fila crecen con ella.
//   MoverTablas.planear(ed, m, nombresHoja, idioma) → plan (aplica las celdas en ed) | null;  await MoverTablas.escribir(libro, plan, M, ft)   (libro: LibroXlsx abierto; escribe en su zip)
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

  function planear(ed, m, nombresHoja, idioma) {
    const hojas = ed.hojas, M = m.M;
    if (!hojas.some((x) => x && !x.sinCeldas && (x.tablas || []).length)) return null;
    const S = { f0: M.f0, c0: M.c0, f1: M.f1, c1: M.c1 }, D = { f0: M.f0 + M.df, c0: M.c0 + M.dc, f1: M.f1 + M.df, c1: M.c1 + M.dc };
    const nombreDe = (t) => t.nombreVisible || t.nombre;
    const palabra = idioma === 'en' ? 'Column' : 'Columna';
    const vacia = (h, f, c) => { const x = ed.estadoDe(h, f, c); return !x || (!x.formula && (x.tipo === 'vacio' || x.valor === null || x.valor === undefined || x.valor === '')); };
    const afectada = (h, f, c) => { const p = { f0: f, f1: f, c0: c, c1: c }; return (h === M.h0 && dentro(p, S)) || (h === M.h1 && dentro(p, D)); };
    const plan = { tablas: new Map(), quitar: new Set(), renombres: [], cambios: [], formula: null };
    const movidas = (hojas[M.h0].tablas || []).filter((t) => dentro(t.ref, S));
    const fijas = (hojas[M.h1].tablas || []).filter((t) => !(M.h0 === M.h1 && movidas.includes(t)));
    let sigue = movidas.slice();
    if (movidas.length) for (const o of fijas) {
      if (!toca(o.ref, D)) continue;
      if (dentro(o.ref, D)) plan.quitar.add(nombreDe(o));                      // queda entera debajo: se pierde
      else sigue = [];                                                         // a medias: la que se pierde es la movida
    }
    if (movidas.length && !sigue.length) for (const t of movidas) plan.quitar.add(nombreDe(t));
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
        let f1 = r.f1, c1 = r.c1;
        const otras = (hoja.tablas || []).filter((o) => o !== t);
        if (h === M.h1 && !movidas.length) {
          const solapa = toca(D, r);
          if (!t.totales && D.f1 > r.f1 && (solapa || (D.f0 === r.f1 + 1 && D.c0 >= r.c0 && D.c1 <= r.c1))) {
            const zona = { f0: r.f1 + 1, f1: D.f1, c0: r.c0, c1: r.c1 };
            let algo = false, libre = true;
            for (let i = zona.f0; i <= zona.f1 && libre; i++) for (let j = zona.c0; j <= zona.c1 && libre; j++) { const v = vacia(h, i, j), enD = i >= D.f0 && i <= D.f1 && j >= D.c0 && j <= D.c1; if (enD && !v) algo = true; if (!enD && !v) libre = false; }
            if (algo && libre && !otras.some((o) => toca(o.ref, zona))) f1 = D.f1;
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
            if (tocada && !vacia(h, r.f0, c)) txt = x.tipo === 'error' ? String(x.valor) : String(x.valor === null || x.valor === undefined ? '' : x.valor);
            estados.push({ tocada, txt, vieja: viejas[k] });
          }
          const usados = new Set();
          const libreCon = (b) => { if (!usados.has(b.toLowerCase())) return b; let z = 2; while (usados.has((b + z).toLowerCase())) z++; return b + z; };
          for (const e of estados) if (e.tocada && e.txt) { e.nombre = libreCon(e.txt); usados.add(e.nombre.toLowerCase()); }
          for (const e of estados) if (!e.tocada) { e.nombre = libreCon(e.vieja); usados.add(e.nombre.toLowerCase()); }
          for (const e of estados) if (e.nombre === undefined) { let z = 1; while (usados.has((palabra + z).toLowerCase())) z++; e.nombre = palabra + z; usados.add(e.nombre.toLowerCase()); }
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
        plan.tablas.set(n, { ref: nuevaRef, hojaVieja: h, hojaNueva: h, nombres, totales, sinCalc, datosF1, calculadas: t.calculadas || new Map(), autollenar: f1 > r.f1 ? { desde: r.f1 + 1, hasta: f1, datosF0 } : null, h });
        if (f1 > r.f1) plan.cambios.push({ h, c0: r.c0, c1: r.c1, ultima: r.f1, arriba: r.f0, datos: datosF0, nueva: f1 });
        t.ref = nuevaRef; t.nombresColumnas = nombres;
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

  function nombreAttr(n) { return X().escAttr(n); }
  function tablaXml(xml, p, ft, formula, h) {
    let u = xml;
    const ref = textoRango(p.ref), r = p.ref;
    u = u.replace(/(<table\b[^>]*?\bref\s*=\s*")[^"]*(")/, (x, a, c) => a + ref + c);
    const totalesFila = /\btotalsRowCount\s*=\s*"[1-9]/.test(u.slice(0, u.indexOf('>') + 1));
    const filtro = textoRango({ f0: r.f0, c0: r.c0, f1: r.f1 - (totalesFila ? 1 : 0), c1: r.c1 });
    u = u.replace(/(<autoFilter\b[^>]*?\bref\s*=\s*")[^"]*(")/, (x, a, c) => a + filtro + c);
    if (p.movida) {
      const M = p.M;
      const corrida = (t) => { const g = LibroXlsx.decodificarRango(t); return g ? textoRango({ f0: g.f0 + M.df, c0: g.c0 + M.dc, f1: g.f1 + M.df, c1: g.c1 + M.dc }) : t; };
      u = u.replace(/(<(?:sortState|sortCondition)\b[^>]*?\bref\s*=\s*")([^"]*)(")/g, (x, a, t, c) => a + corrida(t) + c);
    }
    const fx = (t) => X().esc(formula(ft(X().decodificar(t), p.hojaVieja, p.hojaVieja), h));
    u = u.replace(/(<(calculatedColumnFormula|totalsRowFormula)\b[^>]*(?<!\/)>)([\s\S]*?)(<\/\2>)/g, (x, a, n, t, c) => a + fx(t) + c);
    if (!p.nombres && !p.totales && !p.sinCalc) return u;
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

  globalThis.MoverTablas = { planear, escribir, _enTabla: enTabla };
})();

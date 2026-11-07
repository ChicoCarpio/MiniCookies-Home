/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Cortar y pegar en .ods: lo mismo que en .xlsx (hojas/mover-objetos.js), con las reglas medidas en Excel. Diferencias del formato:
// - La validación está en cada celda (table:content-validation-name) y su definición tiene una celda base: una regla partida son definiciones nuevas, con la base en su esquina.
// - Las notas (office:annotation) y los enlaces (text:a) viven dentro de la celda, así que se copian al destino antes de escribir el valor.
// - No hay tablas como las de Excel: un rango de base de datos solo sigue a lo movido si está entero adentro (LibreOffice no le pone nombres a los encabezados).
//   await MoverOds.leerExtras(bytes, h0, h1, M) → extras (antes de guardar);  await MoverOds.aplicar(bytes, M, nombresHoja, extras) → Uint8Array
(function () {
  const X = () => XmlEditar;
  const EO = () => EstructuraOds._interno, U = () => OdsEscritorCeldas._util, OF = () => OdsFormulas;
  const adentro = (p, r) => p.f >= r.f0 && p.f <= r.f1 && p.c >= r.c0 && p.c <= r.c1;

  // recorre las celdas de las filas [f0, f1] de una tabla: fn(f, c, nodo, textoFila, repeticionesDeCol, repeticionesDeFila)
  function recorrer(texto, tabla, f0, f1, fn) {
    for (const { nodo, f0: fi, n } of U().enumerarFilas(texto, tabla)) {
      if (fi > f1 || fi + n - 1 < f0) continue;
      const t = texto.slice(nodo.ini, nodo.fin), r = X().analizar(t, { profundidad: 2 });
      let c = 0;
      for (const h of r.hijos) {
        if (h.local !== 'table-cell' && h.local !== 'covered-table-cell') continue;
        const rep = U().repeticiones(h, 'table:number-columns-repeated');
        fn(fi, c, h, t, rep, n);
        c += rep;
      }
    }
  }
  const NOTA = /<office:annotation\b[\s\S]*?<\/office:annotation>/;
  function extraDe(h, t) {
    const x = t.slice(h.ini, h.fin);
    if (h.iniCierre < 0) return {};
    const nota = NOTA.exec(x), p = /<text:a\b/.test(x.replace(NOTA, '')) ? (x.replace(NOTA, '').match(/<text:p\b[\s\S]*?<\/text:p>/g) || []).join('') : null;
    return { nota: nota ? nota[0] : null, parrafos: p };
  }
  async function leerExtras(bytes, h0, h1, M) {
    const p = await PaqueteEditable.abrir(bytes), texto = await p.leer('content.xml');
    const d = EO().documento(texto), mapa = new Map();
    recorrer(texto, EO().tablaDe(d, h0), M.f0, M.f1, (f, c, h, t, rep, n) => {
      for (let i = Math.max(f, M.f0); i <= Math.min(f + n - 1, M.f1); i++) for (let j = Math.max(c, M.c0); j <= Math.min(c + rep - 1, M.c1); j++) {
        const e = extraDe(h, t);
        if (e.nota || e.parrafos) mapa.set(i + ',' + j, e);
      }
    });
    return mapa;
  }

  // ---- reglas ------------------------------------------------------------------------------------------------------------------------------------------------------
  const RE_CORCHETE = /\[[^\]]*\]/g;
  const aExcel = (g) => { const x = OF().aExcel('of:=' + g); return x === null || x === undefined ? null : String(x).replace(/^=/, ''); };
  const aOds = (x) => { const a = OF().aOpenFormula(x); return a === null || a === undefined ? null : a; };
  const hojaDeBase = (b) => { const m = /^\$?(?:'((?:[^']|'')+)'|([^.']+))\./.exec(b); return m ? (m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]) : null; };
  const baseA1 = (b) => { const m = /\.\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(b); return m ? { f: +m[2] - 1, c: LibroXlsx.colANumero(m[1].toUpperCase()) } : null; };
  const baseTexto = (nombre, e) => EO().citar(nombre).replace(/^'|'$/g, (q) => q) + '.' + LibroXlsx.numeroACol(e.c) + (e.f + 1);
  // una condición: las referencias de su texto, relativas a la esquina del rango
  function formulasDe(valor, base, e) {
    const l = (valor.match(RE_CORCHETE) || []).map(aExcel);
    return l.some((x) => x === null) ? null : l.map((x) => LibroXlsx.desplazarFormula(x, e.f - base.f, e.c - base.c));
  }
  const conFormulas = (valor, l) => { let k = 0; return valor.replace(RE_CORCHETE, (g) => { const a = aOds(l[k++]); return a === null ? g : a; }); };

  function reglasCondicionales(texto, M, nombresHoja, ft) {
    const d = EO().documento(texto), eds = [], llegan = [];
    d.tablas.forEach((tabla, h) => {
      const x = texto.slice(tabla.ini, tabla.fin);
      for (const m of x.matchAll(/<calcext:conditional-format\b([^>]*)(?<!\/)>([\s\S]*?)<\/calcext:conditional-format>/g)) {
        const at = /calcext:target-range-address="([^"]*)"/.exec(m[1]);
        if (!at) continue;
        const partes = (X().decodificar(at[1]).match(EO().RE_DIRECCIONES) || []).map(EO().direccionAExcel);
        if (!partes.length || partes.some((q) => q === null)) continue;
        const sq = partes.map((q) => q.slice(q.lastIndexOf('!') + 1).replace(/\$/g, '')).join(' ');
        const rangos = sq.split(' ').map((q) => LibroXlsx.decodificarRango(q)), e = MoverObjetos._esquina(rangos);
        // las condiciones, con todas sus referencias en una lista
        const conds = [...m[2].matchAll(/<calcext:condition\b[^>]*\/>|<calcext:condition\b[^>]*>[\s\S]*?<\/calcext:condition>/g)].map((c) => c[0]);
        const lista = [], cortes = [];
        let bien = true;
        for (const c of conds) {
          const v = /calcext:value="([^"]*)"/.exec(c), b = /calcext:base-cell-address="([^"]*)"/.exec(c);
          const base = b ? baseA1(X().decodificar(b[1])) : e;
          const fx = v ? formulasDe(X().decodificar(v[1]), base || e, e) : [];
          if (fx === null) { bien = false; break; }
          cortes.push(fx.length); lista.push(...fx);
        }
        if (!bien) continue;
        const salida = MoverObjetos._planRegla(sq, lista, h, M, ft, false, (i) => nombresHoja[i]);
        const piezas = [];
        for (const p of salida) {
          const eP = MoverObjetos._esquina(p.sqref.split(' ').map((q) => LibroXlsx.decodificarRango(q))), nombre = nombresHoja[p.h];
          let k = 0, cuerpo = m[2];
          for (let i = 0; i < conds.length; i++) {
            const trozo = p.formulas.slice(k, k + cortes[i]);
            k += cortes[i];
            const nueva = conds[i].replace(/(calcext:value=")([^"]*)(")/, (y, a, v, c) => a + X().escAttr(conFormulas(X().decodificar(v), trozo)) + c).replace(/(calcext:base-cell-address=")[^"]*(")/, (y, a, c) => a + X().escAttr(baseTexto(nombre, eP)) + c);
            cuerpo = cuerpo.replace(conds[i], () => nueva);
          }
          const dir = p.sqref.split(' ').map((q) => { const [a, b = a] = q.split(':'); return EO().citar(nombre) + '.' + a + ':' + EO().citar(nombre) + '.' + b; }).join(' ');
          const e2 = `<calcext:conditional-format${m[1].replace(at[0], 'calcext:target-range-address="' + X().escAttr(dir) + '"')}>${cuerpo}</calcext:conditional-format>`;
          if (p.h === h) piezas.push(e2); else llegan.push(e2);
        }
        eds.push({ ini: tabla.ini + m.index, fin: tabla.ini + m.index + m[0].length, texto: piezas.join('') });
      }
    });
    let u = eds.length ? X().aplicar(texto, eds) : texto;
    u = u.replace(/<calcext:conditional-formats\b[^>]*>\s*<\/calcext:conditional-formats>/g, '');
    if (llegan.length) {
      const dd = EO().documento(u), t = dd.tablas[M.h1], x = u.slice(t.ini, t.fin), k = x.indexOf('</calcext:conditional-formats>');
      if (k >= 0) u = u.slice(0, t.ini + k) + llegan.join('') + u.slice(t.ini + k);
      else {
        // el lugar de LibreOffice: detrás de las filas y antes de los nombres de la hoja
        const n = x.indexOf('<table:named-expressions'), pos = t.ini + (n >= 0 ? n : x.lastIndexOf('</table:table>'));
        u = u.slice(0, pos) + '<calcext:conditional-formats>' + llegan.join('') + '</calcext:conditional-formats>' + u.slice(pos);
      }
    }
    return u;
  }

  // ---- validación: una definición por regla, y el nombre en cada celda --------------------------------------------------------------------------------------------------
  function validacionesDeHoja(texto, tabla) {
    const porNombre = new Map();
    const x = texto.slice(tabla.ini, tabla.fin);
    if (x.indexOf('content-validation-name') < 0) return porNombre;
    recorrer(texto, tabla, 0, 1048575, (f, c, h, t, rep, n) => {
      const v = h.attrs['table:content-validation-name'];
      if (!v) return;
      const nombre = X().decodificar(v);
      if (!porNombre.has(nombre)) porNombre.set(nombre, []);
      porNombre.get(nombre).push({ f0: f, c0: c, f1: f + n - 1, c1: c + rep - 1 });
    });
    return porNombre;
  }
  function validaciones(texto, M, nombresHoja, ft) {
    let d = EO().documento(texto);
    const defs = new Map();
    for (const m of texto.matchAll(/<table:content-validation\b[^>]*?(?:\/>|>[\s\S]*?<\/table:content-validation>)/g)) { const n = /table:name="([^"]*)"/.exec(m[0]); if (n) defs.set(X().decodificar(n[1]), m[0]); }
    if (!defs.size) return { texto, celdas: new Map() };
    const usados = new Set(defs.keys());
    let nuevas = '', cuenta = 0;
    const celdas = new Map();   // hoja → Map('f,c' → nombre | null)
    const poner = (h, f, c, n) => { if (!celdas.has(h)) celdas.set(h, new Map()); celdas.get(h).set(f + ',' + c, n); };
    const ponerDespues = [];
    const deOrigen = validacionesDeHoja(texto, d.tablas[M.h0]);
    const pisadas = MoverObjetos._pisadasDe([...deOrigen.values()].map((rs) => rs.map(MoverObjetos._textoRango).join(' ')), M);
    d.tablas.forEach((tabla, h) => {
      for (const [nombre, rs] of validacionesDeHoja(texto, tabla)) {
        const def = defs.get(nombre);
        if (!def) continue;
        const sq = rs.map(MoverObjetos._textoRango).join(' ');
        const e = MoverObjetos._esquina(rs), mc = /table:condition="([^"]*)"/.exec(def), mb = /table:base-cell-address="([^"]*)"/.exec(def);
        const base = (mb && baseA1(X().decodificar(mb[1]))) || e;
        const fx = mc ? formulasDe(X().decodificar(mc[1]), base, e) : [];
        if (fx === null) continue;
        const salida = MoverObjetos._planRegla(sq, fx, h, M, ft, pisadas, (i) => nombresHoja[i]);
        const antes = rs.reduce((a, r) => a + (r.f1 - r.f0 + 1) * (r.c1 - r.c0 + 1), 0);
        const quedan = salida.reduce((a, p) => a + p.sqref.split(' ').reduce((b, q) => { const g = LibroXlsx.decodificarRango(q); return b + (g.f1 - g.f0 + 1) * (g.c1 - g.c0 + 1); }, 0), 0);
        const igual = salida.length === 1 && salida[0].h === h && salida[0].sqref.split(' ').sort().join(' ') === sq.split(' ').sort().join(' ') && antes === quedan && JSON.stringify(salida[0].formulas) === JSON.stringify(fx);
        if (igual) continue;
        // las celdas que tenían esta regla la pierden; y cada parte, una definición nueva con base en su esquina
        for (const r of rs) if ((r.f1 - r.f0 + 1) * (r.c1 - r.c0 + 1) <= 200000) for (let i = r.f0; i <= r.f1; i++) for (let j = r.c0; j <= r.c1; j++) poner(h, i, j, null);
        for (const p of salida) {
          const rg = p.sqref.split(' ').map((q) => LibroXlsx.decodificarRango(q)), eP = MoverObjetos._esquina(rg);
          let n; do { n = 'mcv' + (++cuenta); } while (usados.has(n));
          usados.add(n);
          let copia = def.replace(/(table:name=")[^"]*(")/, (y, a, c) => a + n + c);
          if (mc) copia = copia.replace(/(table:condition=")[^"]*(")/, (y, a, c) => a + X().escAttr(conFormulas(X().decodificar(mc[1]), p.formulas)) + c);
          const bt = baseTexto(nombresHoja[p.h], eP);
          copia = mb ? copia.replace(/(table:base-cell-address=")[^"]*(")/, (y, a, c) => a + X().escAttr(bt) + c) : copia.replace(/^<table:content-validation\b/, () => '<table:content-validation table:base-cell-address="' + X().escAttr(bt) + '"');
          nuevas += copia;
          ponerDespues.push([p.h, rg, n]);
        }
      }
    });
    for (const [h, rg, n] of ponerDespues) for (const r of rg) for (let i = r.f0; i <= r.f1; i++) for (let j = r.c0; j <= r.c1; j++) poner(h, i, j, n);
    let u = texto;
    if (nuevas) u = u.replace('</table:content-validations>', () => nuevas + '</table:content-validations>');
    return { texto: u, celdas };
  }

  // ---- escribir lo de cada celda: nota, enlace, validación ---------------------------------------------------------------------------------------------------------------
  function escribirCeldas(texto, h, cambios) {
    const d = EO().documento(texto), tabla = EO().tablaDe(d, h), lista = [];
    for (const [k, e] of cambios) { const [f, c] = k.split(',').map(Number); lista.push(Object.assign({ f, c }, e)); }
    if (!lista.length) return texto;
    const eds = U().crearMotor(texto).editarTabla(tabla, lista, (orig, t, e) => {
      let et = orig ? X().ponerAtributo(t.slice(orig.ini, orig.finApertura), 'table:number-columns-repeated', null) : '<table:table-cell/>';
      if ('validacion' in e) et = X().ponerAtributo(et, 'table:content-validation-name', e.validacion);
      let dentro = orig && orig.iniCierre >= 0 ? t.slice(orig.finApertura, orig.iniCierre) : '';
      if ('nota' in e) dentro = (e.nota || '') + dentro.replace(NOTA, '');
      if (e.parrafos) dentro = dentro.replace(/<text:p\b[\s\S]*?<\/text:p>|<text:p\b[^>]*\/>/g, '') + e.parrafos;
      et = U().sinCierre(et);
      return dentro ? et + dentro + '</' + (orig ? orig.nombre : 'table:table-cell') + '>' : et.replace(/>$/, '/>');
    });
    return eds.length ? X().aplicar(texto, eds) : texto;
  }

  async function aplicar(bytes, M, nombresHoja, extras) {
    const ft = (t, hF, hNueva) => LibroEditable._formulaTrasMover(t, hF, hNueva, M, nombresHoja);
    const rw = (t, enObjetivo) => LibroEditable._formulaTrasMover(t, enObjetivo ? M.h0 : -1, enObjetivo ? M.h0 : -1, M, nombresHoja);
    // un rango de base de datos solo sigue a lo movido si está entero adentro
    const rwBase = (t) => {
      const m = /^('(?:[^']|'')+'|[^!]+)!\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(t);
      if (!m) return t;
      const hoja = nombresHoja.findIndex((n) => n.toLowerCase() === m[1].replace(/^'|'$/g, '').replace(/''/g, "'").toLowerCase());
      const r = { f0: +m[3] - 1, c0: LibroXlsx.colANumero(m[2].toUpperCase()), f1: +(m[5] || m[3]) - 1, c1: LibroXlsx.colANumero((m[4] || m[2]).toUpperCase()) };
      if (hoja !== M.h0 || r.f0 < M.f0 || r.f1 > M.f1 || r.c0 < M.c0 || r.c1 > M.c1) return t;
      const fin = { f0: r.f0 + M.df, c0: r.c0 + M.dc, f1: r.f1 + M.df, c1: r.c1 + M.dc };
      return XlsFormulas._citarHoja(nombresHoja[M.h1]) + '!$' + LibroXlsx.numeroACol(fin.c0) + '$' + (fin.f0 + 1) + ':$' + LibroXlsx.numeroACol(fin.c1) + '$' + (fin.f1 + 1);
    };
    const b = await EstructuraOds.reescribirReferencias(bytes, M.h0, rw, rwBase);
    return EO().conPaquete(b, async (p) => {
      let texto = await p.leer('content.xml');
      texto = reglasCondicionales(texto, M, nombresHoja, ft);
      const v = validaciones(texto, M, nombresHoja, ft);
      texto = v.texto;
      const porHoja = new Map(), de = (h) => { if (!porHoja.has(h)) porHoja.set(h, new Map()); return porHoja.get(h); };
      const uno = (h, k, c) => de(h).set(k, Object.assign(de(h).get(k) || {}, c));
      for (const [h, m] of v.celdas) for (const [k, n] of m) uno(h, k, { validacion: n });
      // lo que estaba en el origen se va; lo del destino se pierde y llega lo movido
      if (extras) {
        for (let i = M.f0; i <= M.f1; i++) for (let j = M.c0; j <= M.c1; j++) {
          if (extras.has(i + ',' + j)) uno(M.h0, i + ',' + j, { nota: null });
          uno(M.h1, (i + M.df) + ',' + (j + M.dc), { nota: null });
        }
        for (const [k, e] of extras) { const [i, j] = k.split(',').map(Number); uno(M.h1, (i + M.df) + ',' + (j + M.dc), { nota: e.nota, parrafos: e.parrafos }); }
      }
      for (const [h, m] of porHoja) texto = escribirCeldas(texto, h, m);
      p.escribir('content.xml', texto);
    });
  }

  function extrasDe(texto, h, M) {
    const d = EO().documento(texto), mapa = new Map();
    recorrer(texto, EO().tablaDe(d, h), M.f0, M.f1, (f, c, nodo, t, rep, n) => {
      for (let i = Math.max(f, M.f0); i <= Math.min(f + n - 1, M.f1); i++) for (let j = Math.max(c, M.c0); j <= Math.min(c + rep - 1, M.c1); j++) {
        const e = extraDe(nodo, t);
        if (e.nota || e.parrafos) mapa.set(i + ',' + j, e);
      }
    });
    return mapa;
  }
  // el texto que se ve en una celda
  function textoDe(texto, h, f, c) {
    let s = null;
    recorrer(texto, EO().tablaDe(EO().documento(texto), h), f, f, (fi, ci, nodo, t, rep) => { if (s === null && c >= ci && c < ci + rep) s = (t.slice(nodo.ini, nodo.fin).replace(NOTA, '').match(/<text:p\b[\s\S]*?<\/text:p>/g) || []).join('\n').replace(/<[^>]+>/g, ''); });
    return s || '';
  }
  globalThis.MoverOds = { aplicar, leerExtras, _interno: { baseA1, baseTexto, formulasDe, conFormulas, validacionesDeHoja, escribirCeldas, extrasDe, textoDe } };
})();

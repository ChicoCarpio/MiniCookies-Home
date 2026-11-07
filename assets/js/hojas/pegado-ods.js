/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Copiar y pegar en .ods: lo que no es una celda, con las mismas reglas que en .xlsx (pegado-objetos.js). Lo propio del formato está en mover-ods.js: la validación va en cada celda, con su definición y su
// celda base, y las notas y los enlaces viven dentro de la celda.
//   await PegadoOds.aplicar(bytes, objetos, { idioma }) → Uint8Array
(function () {
  const X = () => XmlEditar, EO = () => EstructuraOds._interno, MO = () => MoverOds._interno, MI = () => MoverObjetos._interno, PO = () => PegadoObjetos;
  const dentro = (p, r) => p.f >= r.f0 && p.f <= r.f1 && p.c >= r.c0 && p.c <= r.c1;

  function reglasCondicionales(texto, O, nombres) {
    const d = EO().documento(texto), eds = [], llegan = [];
    const cortes = O.cf === 'reemplaza' ? PO()._cortesDe(O, (p) => !p.blanca) : null;
    d.tablas.forEach((tabla, h) => {
      if (h !== O.h0 && h !== O.h1) return;
      const x = texto.slice(tabla.ini, tabla.fin);
      for (const m of x.matchAll(/<calcext:conditional-format\b([^>]*)(?<!\/)>([\s\S]*?)<\/calcext:conditional-format>/g)) {
        const at = /calcext:target-range-address="([^"]*)"/.exec(m[1]);
        if (!at) continue;
        const partes = (X().decodificar(at[1]).match(EO().RE_DIRECCIONES) || []).map(EO().direccionAExcel);
        if (!partes.length || partes.some((q) => q === null)) continue;
        const sq = partes.map((q) => q.slice(q.lastIndexOf('!') + 1).replace(/\$/g, '')).join(' ');
        const rangos = sq.split(' ').map((q) => LibroXlsx.decodificarRango(q)), e = MoverObjetos._esquina(rangos);
        const conds = [...m[2].matchAll(/<calcext:condition\b[^>]*\/>|<calcext:condition\b[^>]*>[\s\S]*?<\/calcext:condition>/g)].map((c) => c[0]);
        const lista = [], cuantas = [];
        let bien = true;
        for (const c of conds) {
          const v = /calcext:value="([^"]*)"/.exec(c), b = /calcext:base-cell-address="([^"]*)"/.exec(c);
          const base = b ? MO().baseA1(X().decodificar(b[1])) : e;
          const fx = v ? MO().formulasDe(X().decodificar(v[1]), base || e, e) : [];
          if (fx === null) { bien = false; break; }
          cuantas.push(fx.length); lista.push(...fx);
        }
        if (!bien) continue;
        const r = PO()._planRegla(sq, lista, h, O, cortes, nombres[O.h1]);
        const elemento = (p, hoja) => {
          const q = p.sq.split(' ').map((z) => LibroXlsx.decodificarRango(z)), eP = MoverObjetos._esquina(q), nombre = nombres[hoja];
          let k = 0, cuerpo = m[2];
          for (let i = 0; i < conds.length; i++) {
            const trozo = p.formulas.slice(k, k + cuantas[i]);
            k += cuantas[i];
            const nueva = conds[i].replace(/(calcext:value=")([^"]*)(")/, (y, a, v, c) => a + X().escAttr(MO().conFormulas(X().decodificar(v), trozo)) + c).replace(/(calcext:base-cell-address=")[^"]*(")/, (y, a, c) => a + X().escAttr(MO().baseTexto(nombre, eP)) + c);
            cuerpo = cuerpo.replace(conds[i], () => nueva);
          }
          const dir = p.sq.split(' ').map((z) => { const [a, b = a] = z.split(':'); return EO().citar(nombre) + '.' + a + ':' + EO().citar(nombre) + '.' + b; }).join(' ');
          return `<calcext:conditional-format${m[1].replace(at[0], 'calcext:target-range-address="' + X().escAttr(dir) + '"')}>${cuerpo}</calcext:conditional-format>`;
        };
        let piezas = '';
        if (r.llega) { if (h === O.h1) piezas += elemento(r.llega, h); else llegan.push(elemento(r.llega, O.h1)); }
        if (r.queda) piezas += elemento(r.queda, h);
        eds.push({ ini: tabla.ini + m.index, fin: tabla.ini + m.index + m[0].length, texto: piezas });
      }
    });
    let u = eds.length ? X().aplicar(texto, eds) : texto;
    u = u.replace(/<calcext:conditional-formats\b[^>]*>\s*<\/calcext:conditional-formats>/g, '');
    if (llegan.length) {
      const dd = EO().documento(u), t = dd.tablas[O.h1], x = u.slice(t.ini, t.fin), k = x.indexOf('</calcext:conditional-formats>');
      if (k >= 0) u = u.slice(0, t.ini + k) + llegan.join('') + u.slice(t.ini + k);
      else { const n = x.indexOf('<table:named-expressions'), pos = t.ini + (n >= 0 ? n : x.lastIndexOf('</table:table>')); u = u.slice(0, pos) + '<calcext:conditional-formats>' + llegan.join('') + '</calcext:conditional-formats>' + u.slice(pos); }
    }
    return u;
  }

  function validaciones(texto, O, nombres, poner) {
    const d = EO().documento(texto), defs = new Map();
    for (const m of texto.matchAll(/<table:content-validation\b[^>]*?(?:\/>|>[\s\S]*?<\/table:content-validation>)/g)) { const n = /table:name="([^"]*)"/.exec(m[0]); if (n) defs.set(X().decodificar(n[1]), m[0]); }
    if (!defs.size) return texto;
    const usados = new Set(defs.keys());
    let nuevas = '', cuenta = 0;
    const deOrigen = MO().validacionesDeHoja(texto, d.tablas[O.h0]);
    const conDv = O.pares.filter((p) => { for (const rs of deOrigen.values()) if (rs.some((r) => dentro({ f: p.sf, c: p.sc }, r))) return true; return false; });
    const cortes = O.blancos ? PO()._unirCeldas(conDv.map((p) => ({ f: p.df, c: p.dc }))) : [O.D];
    const trabajo = O.h0 === O.h1 ? [O.h0] : [O.h0, O.h1], llegan = [];
    for (const h of trabajo) {
      for (const [nombre, rs] of MO().validacionesDeHoja(texto, d.tablas[h])) {
        const def = defs.get(nombre);
        if (!def) continue;
        const sq = rs.map(MoverObjetos._textoRango).join(' '), e = MoverObjetos._esquina(rs), mc = /table:condition="([^"]*)"/.exec(def), mb = /table:base-cell-address="([^"]*)"/.exec(def);
        const base = (mb && MO().baseA1(X().decodificar(mb[1]))) || e;
        const fx = mc ? MO().formulasDe(X().decodificar(mc[1]), base, e) : [];
        if (fx === null) continue;
        const r = PO()._planRegla(sq, fx, h, O, h === O.h1 ? cortes : null, nombres[O.h1]);
        const parte = (p, hoja) => {
          const rg = p.sq.split(' ').map((q) => LibroXlsx.decodificarRango(q)), eP = MoverObjetos._esquina(rg);
          let n; do { n = 'mcv' + (++cuenta); } while (usados.has(n));
          usados.add(n);
          let copia = def.replace(/(table:name=")[^"]*(")/, (y, a, c) => a + n + c);
          if (mc) copia = copia.replace(/(table:condition=")[^"]*(")/, (y, a, c) => a + X().escAttr(MO().conFormulas(X().decodificar(mc[1]), p.formulas)) + c);
          const bt = MO().baseTexto(nombres[hoja], eP);
          copia = mb ? copia.replace(/(table:base-cell-address=")[^"]*(")/, (y, a, c) => a + X().escAttr(bt) + c) : copia.replace(/^<table:content-validation\b/, () => '<table:content-validation table:base-cell-address="' + X().escAttr(bt) + '"');
          nuevas += copia;
          for (const g of rg) for (let i = g.f0; i <= g.f1; i++) for (let j = g.c0; j <= g.c1; j++) poner(hoja, i, j, { validacion: n });
        };
        // lo que se corta pierde la regla; lo que queda sigue con la suya, que ya es relativa a su base
        if (h === O.h1) for (const z of cortes) for (const g of rs) { const f0 = Math.max(z.f0, g.f0), f1 = Math.min(z.f1, g.f1), c0 = Math.max(z.c0, g.c0), c1 = Math.min(z.c1, g.c1); for (let i = f0; i <= f1; i++) for (let j = c0; j <= c1; j++) { if (!(h === O.h0 && r.queda && r.queda.sq.split(' ').some((q) => dentro({ f: i, c: j }, LibroXlsx.decodificarRango(q))))) poner(h, i, j, { validacion: null }); } }
        if (r.llega) llegan.push(() => parte(r.llega, O.h1));
      }
    }
    for (const hazla of llegan) hazla();
    return nuevas ? texto.replace('</table:content-validations>', () => nuevas + '</table:content-validations>') : texto;
  }

  async function aplicar(bytes, O, opciones) {
    const nombres = (await LibroEditable.abrir(bytes, opciones)).libro.hojas.map((x) => x.nombre);
    let b = await EO().conPaquete(bytes, async (p) => {
      let texto = await p.leer('content.xml');
      if (O.cf) texto = reglasCondicionales(texto, O, nombres);
      const porHoja = new Map(), poner = (h, f, c, e) => { if (!porHoja.has(h)) porHoja.set(h, new Map()); const k = f + ',' + c, m = porHoja.get(h); m.set(k, Object.assign(m.get(k) || {}, e)); };
      if (O.dv) texto = validaciones(texto, O, nombres, poner);
      if (O.notas || O.enlaces) {
        const M = { f0: O.S.f0, c0: O.S.c0, f1: O.S.f1, c1: O.S.c1 };
        const extras = await MO().extrasDe(texto, O.h0, M);
        if (O.notas) for (const q of O.pares) {
          const e = extras.get(q.sf + ',' + q.sc);
          if (e && e.nota) poner(O.h1, q.df, q.dc, { nota: e.nota });
          else if (!O.blancos) poner(O.h1, q.df, q.dc, { nota: null });
        }
        if (O.enlaces) for (const q of O.pares) { const e = extras.get(q.sf + ',' + q.sc); if (e && e.parrafos) poner(O.h1, q.df, q.dc, { parrafos: e.parrafos }); }
      }
      const seQueda = (k) => !(O.enlaces && !O.blancos);
      for (const [k, e] of O.enlacesPrevios || []) {
        if (!e.parrafos || !seQueda(k)) continue;
        const [f, c] = k.split(',').map(Number), a = /<text:a\b[^>]*>/.exec(e.parrafos);
        if (a) poner(O.h1, f, c, { parrafos: '<text:p>' + a[0] + X().esc(X().decodificar(MO().textoDe(texto, O.h1, f, c))) + '</text:a></text:p>' });
      }
      for (const q of O.externos || []) {
        const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
        poner(O.h1, q.f, q.c, { parrafos: '<text:p><text:a xlink:href="' + esc(q.url) + '" xlink:type="simple">' + esc(q.texto || q.url) + '</text:a></text:p>' });
      }
      for (const [h, m] of porHoja) texto = MO().escribirCeldas(texto, h, m);
      p.escribir('content.xml', texto);
    });
    if (O.anchos.length || O.altos.length) {
      const ed = await LibroEditable.abrir(bytes, opciones), H0 = ed.hojas[O.h0];
      for (const q of O.anchos) {
        const k = (H0.columnas || []).find((x) => q.de >= x.min && q.de <= x.max);
        if (k && k.ancho > 0) b = await EstructuraOds.anchoColumnas(b, O.h1, q.c, q.c, k.ancho);
      }
      for (const q of O.altos) { const x = H0.filas && H0.filas.get(q.de); if (x && x.alto) b = await EstructuraOds.altoFilas(b, O.h1, q.f, q.f, x.alto); }
    }
    for (const m of O.comb.quitar) b = await EstructuraOds.separar(b, O.h1, m);
    for (const m of O.comb.poner) b = await EstructuraOds.combinar(b, O.h1, m);
    return b;
  }

  globalThis.PegadoOds = { aplicar };
})();

/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Copiar y pegar: lo que no es una celda. Formato condicional, validación, notas, hipervínculos, combinadas, anchos de columna, altos de fila y tablas, sobre el archivo. Medido en Excel:
// pruebas/oraculo/excel-pegado-especial.ps1. Formato condicional y validación: una regla llega a su destino diciendo lo mismo para cada celda (relativa a ella), y lo que había debajo se corta, aunque no llegue nada;
// con «omitir blancos» solo se corta donde llega algo. Una celda a la que le llega la regla que ya tenía se queda como está. En .ods: pegado-ods.js.
//   PegadoObjetos.planear(ed, P, h, D, T, o, pares, G, tabla) → objetos | null   (null: no hay nada que escribir, se pega solo con celdas)
//   await PegadoObjetos.aplicar(bytes, objetos, { idioma }) → Uint8Array
(function () {
  const X = () => XmlEditar, MI = () => MoverObjetos._interno;
  const toca = (a, b) => a.f0 <= b.f1 && b.f0 <= a.f1 && a.c0 <= b.c1 && b.c0 <= a.c1;
  const dentro = (a, b) => a.f0 >= b.f0 && a.f1 <= b.f1 && a.c0 >= b.c0 && a.c1 <= b.c1;
  const enRangos = (l, f, c) => l.some((r) => f >= r.f0 && f <= r.f1 && c >= r.c0 && c <= r.c1);
  const COLS = 16384, FILAS = 1048576;

  // ---- ¿hay algo que escribir? -------------------------------------------------------------------------------------------------------------------------------------------------------------
  function planear(ed, P, h, D, T, o, pares, G, tabla) {
    const h0 = P.h, H0 = ed.hojas[h0], H1 = ed.hojas[h], S = { f0: P.filas[0], c0: P.r.c0, f1: P.filas[P.filas.length - 1], c1: P.r.c1 };
    const hayEn = (l, a, b) => (l || []).some((x) => toca(a(x), b));
    const conRangos = (x) => x.rangos || [];
    const tocaLista = (lista, r) => (lista || []).some((x) => conRangos(x).some((g) => toca(g, r)));
    const cf = T.cf && (tocaLista(H0.formatosCondicionales, S) || (T.cf === 'reemplaza' && tocaLista(H1.formatosCondicionales, D)));
    const dv = T.dv && (tocaLista(H0.validaciones, S) || tocaLista(H1.validaciones, D));
    const notaEn = (x) => ({ f0: x.f, c0: x.c, f1: x.f, c1: x.c });
    const hilos = (hh) => hh && hh.partes && hh.partes.hilos && hh.partes.hilos.length;
    const notas = T.notas && (hayEn(H0.comentarios, notaEn, S) || hayEn(H1.comentarios, notaEn, D) || ed.libro.extrasDeCelda || hilos(H0) || hilos(H1));
    const enlaces = T.enlaces && (hayEn(H0.hipervinculos, (x) => x, S) || hayEn(H1.hipervinculos, (x) => x, D) || ed.libro.extrasDeCelda);
    // combinadas
    const comb = { quitar: [], poner: [] };
    if (T.comb) {
      const desde = (H0.combinadas || []).filter((m) => dentro(m, P.r));
      const base = new Map();
      for (const m of desde) {
        const lista = pares.filter((p) => p.sf === m.f0 && p.sc === m.c0);
        for (const p of lista) {
          const g = o.trans ? { f0: p.df, c0: p.dc, f1: p.df + (m.c1 - m.c0), c1: p.dc + (m.f1 - m.f0) } : { f0: p.df, c0: p.dc, f1: p.df + (m.f1 - m.f0), c1: p.dc + (m.c1 - m.c0) };
          if (g.f1 <= D.f1 && g.c1 <= D.c1) base.set(g.f0 + ',' + g.c0, g);
        }
      }
      comb.poner = [...base.values()];
      for (const m of H1.combinadas || []) {
        if (!toca(m, D)) continue;
        const iguales = comb.poner.some((g) => g.f0 === m.f0 && g.c0 === m.c0 && g.f1 === m.f1 && g.c1 === m.c1);
        if (iguales) { comb.poner = comb.poner.filter((g) => !(g.f0 === m.f0 && g.c0 === m.c0 && g.f1 === m.f1 && g.c1 === m.c1)); continue; }
        const unaSobreLaSuya = P.celdas.length === 1 && P.celdas[0].length === 1 && o.tipo !== 'formatos';
        if (!unaSobreLaSuya) comb.quitar.push(m);
      }
    }
    // anchos de columna: Anchos, y una columna entera con su formato. Altos: una fila entera con su formato
    const anchos = [], altos = [];
    if (!P.filaEntera && (T.anchos || (P.colEntera && T.formato && T.formato !== 'numero'))) {
      const n = P.r.c1 - P.r.c0 + 1;   // también al transponer: se recorren las columnas de lo copiado
      for (let c = D.c0; c <= D.c1 && c < COLS; c++) anchos.push({ c, de: P.r.c0 + ((c - D.c0) % n) });
    }
    if (P.filaEntera && T.formato && T.formato !== 'numero') {
      for (let k = 0; k < P.filas.length; k++) altos.push({ f: D.f0 + k, de: P.filas[k] });
    }
    const tablas = tabla && tabla.lista ? tabla.lista : [];
    const tapadas = tablas.length ? (H1.tablas || []).filter((t) => dentro(t.ref, D) && !(h0 === h && dentro(t.ref, P.r))) : [];
    // en .ods un enlace vive en el texto de la celda: al escribirla hay que volver a ponerlo
    const conservarEnlaces = ed.formato === 'ods' && ed.libro.extrasDeCelda && (T.contenido || T.formato) && !T.enlaces;
    if (!cf && !dv && !notas && !enlaces && !conservarEnlaces && !comb.quitar.length && !comb.poner.length && !anchos.length && !altos.length && !tablas.length) return null;
    return { h0, h1: h, D, S: { f0: S.f0, c0: S.c0, f1: S.f1, c1: S.c1 }, pares, cf: cf ? T.cf : null, dv: !!dv, notas: !!notas, enlaces: !!enlaces, blancos: !!o.blancos, comb, anchos, altos, tablas, tapadas, formato: ed.formato, trans: !!o.trans };
  }

  // ---- reglas ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  const unirCeldas = (l) => MI().unir(l.map((p) => ({ f0: p.f, c0: p.c, f1: p.f, c1: p.c })));
  // Por dónde se corta una regla debajo del destino: todo el destino, o con «omitir blancos» solo lo que dice `corta`
  function cortesDe(O, porCelda) {
    if (!O.blancos) return [O.D];
    return unirCeldas(O.pares.filter(porCelda).map((p) => ({ f: p.df, c: p.dc })));
  }
  // Una regla (rangos y fórmulas, relativas a su esquina). → { queda: { rangos, formulas } | null, llega: { rangos, formulas } | null }
  function planRegla(sq, formulas, h, O, cortes, nombreDestino) {
    const rangos = MI().leerRangos(sq);
    if (!rangos.length || rangos.some((r) => !r)) return { queda: { sq, formulas }, llega: null };
    const e = MI().esquina(rangos);
    // lo que llega: las celdas de destino cuyo origen tiene la regla. Las que ya la tenían no cambian
    let llega = [];
    const propias = [];
    if (h === O.h0) for (const p of O.pares) if (enRangos(rangos, p.sf, p.sc)) { if (O.h0 === O.h1 && enRangos(rangos, p.df, p.dc)) propias.push(p); else llega.push({ f: p.df, c: p.dc }); }
    let quedan = rangos;
    if (h === O.h1 && cortes) {
      let corta = cortes;
      if (propias.length) { const m = new Set(propias.map((p) => p.df + ',' + p.dc)); const l = []; for (const c of cortes) for (let i = c.f0; i <= c.f1; i++) for (let j = c.c0; j <= c.c1; j++) if (!m.has(i + ',' + j)) l.push({ f: i, c: j }); corta = unirCeldas(l); }
      for (const z of corta) quedan = quedan.flatMap((x) => MI().restar(x, z));
    }
    const salida = { queda: null, llega: null };
    if (quedan.length) {
      const e2 = MI().esquina(quedan);
      salida.queda = { sq: quedan.map(MI().textoRango).join(' '), formulas: formulas.map((f) => LibroXlsx.desplazarFormula(f, e2.f - e.f, e2.c - e.c)) };
    }
    if (llega.length) {
      const g = unirCeldas(llega), e3 = MI().esquina(g);
      salida.llega = { sq: g.map(MI().textoRango).join(' '), formulas: formulas.map((f) => {
        const u = LibroXlsx.desplazarFormula(f, e3.f - e.f, e3.c - e.c);
        return O.h0 !== O.h1 ? MI().sinHojaPropia(u, nombreDestino) : u;
      }) };
    }
    return salida;
  }

  function reglas(xml, h, O, nombreDestino, llevar) {
    const I = MI();
    let u = xml;
    const cortesCF = O.cf === 'reemplaza' ? cortesDe(O, (p) => !p.blanca) : null;
    if (O.cf) {
      const idsNuevos = new Map();
      u = u.replace(/<conditionalFormatting\b([^>]*?)\bsqref\s*=\s*"([^"]*)"([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g, (todo, a, sq, b, hijos) => {
        const r = planRegla(X().decodificar(sq), [...hijos.matchAll(I.RE_FORMULA)].map((m) => X().decodificar(m[2])), h, O, cortesCF, nombreDestino);
        let e = '';
        if (r.llega) {
          // el destino pasa delante de su origen (la misma prioridad; se renumera al final). Una barra de datos lleva su pareja en x14, con un identificador propio
          const hijos2 = I.conFormulas(hijos, r.llega.formulas).replace(/(<x14:id>\{)([0-9A-Fa-f-]{36})(\}<\/x14:id>)/g, (y, p, id, c) => { const n = I.idNuevo(id); idsNuevos.set(id.toUpperCase(), n); return p + n + c; });
          const el = `<conditionalFormatting${a}sqref="${r.llega.sq}"${b}>${hijos2}</conditionalFormatting>`;
          if (h === O.h1) e += el; else llevar.cf.push(el);
        }
        if (r.queda) e += `<conditionalFormatting${a}sqref="${r.queda.sq}"${b}>${I.conFormulas(hijos, r.queda.formulas)}</conditionalFormatting>`;
        return e;
      });
      u = u.replace(/<x14:conditionalFormattings\b([^>]*)>([\s\S]*?)<\/x14:conditionalFormattings>/g, (todo, a, hijos) => {
        let n = 0;
        const v = hijos.replace(/<x14:conditionalFormatting\b([^>]*)>([\s\S]*?)<\/x14:conditionalFormatting>/g, (x, b, cuerpo) => {
          const sq = I.RE_XMSQ.exec(cuerpo);
          if (!sq) { n++; return x; }
          const r = planRegla(X().decodificar(sq[2]), [...cuerpo.matchAll(I.RE_XMF)].map((m) => X().decodificar(m[2])), h, O, cortesCF, nombreDestino);
          const con = (p) => I.conXm(cuerpo, p.formulas).replace(I.RE_XMSQ, (y, c1, s, c3) => c1 + p.sq + c3);
          let e = '';
          if (r.llega) {
            const c2 = con(r.llega).replace(/(<x14:cfRule\b[^>]*?\bid\s*=\s*"\{)([0-9A-Fa-f-]{36})(\})/g, (y, p, id, c) => p + (idsNuevos.get(id.toUpperCase()) || I.idNuevo(id)) + c);
            const el = '<x14:conditionalFormatting' + b + '>' + c2 + '</x14:conditionalFormatting>';
            if (h === O.h1) { e += el; n++; } else llevar.x14cf.push(el);
          }
          if (r.queda) { e += '<x14:conditionalFormatting' + b + '>' + con(r.queda) + '</x14:conditionalFormatting>'; n++; }
          return e;
        });
        return n ? '<x14:conditionalFormattings' + a + '>' + v + '</x14:conditionalFormattings>' : '';
      });
    }
    if (O.dv) {
      const rangosDv = [];
      for (const m of xml.matchAll(/<dataValidation\b[^>]*?\bsqref\s*=\s*"([^"]*)"/g)) rangosDv.push(...I.leerRangos(X().decodificar(m[1])));
      for (const m of xml.matchAll(/<x14:dataValidation\b[\s\S]*?<xm:sqref>([^<]*)<\/xm:sqref>/g)) rangosDv.push(...I.leerRangos(X().decodificar(m[1])));
      if (h === O.h0) O.conDv = O.pares.filter((p) => enRangos(rangosDv.filter(Boolean), p.sf, p.sc));
      const cortesDv = O.blancos ? unirCeldas((O.conDv || []).map((p) => ({ f: p.df, c: p.dc }))) : [O.D];
      let quedan = null;
      u = u.replace(/<dataValidation\b([^>]*?)(\/>|>([\s\S]*?)<\/dataValidation>)/g, (todo, attrs, resto, hijos) => {
        const ms = /\bsqref\s*=\s*"([^"]*)"/.exec(attrs);
        if (!ms) return todo;
        const r = planRegla(X().decodificar(ms[1]), hijos ? [...hijos.matchAll(I.RE_FORMULA)].map((m) => X().decodificar(m[2])) : [], h, O, h === O.h1 ? cortesDv : null, nombreDestino);
        quedan = quedan === null ? 0 : quedan;
        const hazla = (p, nueva) => { const a2 = (nueva ? attrs.replace(/\s+xr:uid\s*=\s*"[^"]*"/, '') : attrs).replace(ms[0], `sqref="${p.sq}"`); return hijos === undefined ? `<dataValidation${a2}/>` : `<dataValidation${a2}>${I.conFormulas(hijos, p.formulas)}</dataValidation>`; };
        let e = '';
        if (r.queda) { e += hazla(r.queda, false); quedan++; }
        if (r.llega) { const el = hazla(r.llega, true); if (h === O.h1) { e += el; quedan++; } else llevar.dv.push(el); }
        return e;
      });
      if (quedan === 0) u = u.replace(/<dataValidations\b[^>]*>\s*<\/dataValidations>/, '');
      else if (quedan !== null) u = u.replace(/(<dataValidations\b[^>]*?\bcount\s*=\s*")\d+(")/, (x, a, c) => a + quedan + c);
      u = u.replace(/<x14:dataValidations\b([^>]*)>([\s\S]*?)<\/x14:dataValidations>/g, (todo, a, hijos) => {
        let n = 0;
        const v = hijos.replace(/<x14:dataValidation\b([^>]*)>([\s\S]*?)<\/x14:dataValidation>/g, (x, b, cuerpo) => {
          const sq = I.RE_XMSQ.exec(cuerpo);
          if (!sq) { n++; return x; }
          const r = planRegla(X().decodificar(sq[2]), [...cuerpo.matchAll(I.RE_XMF)].map((m) => X().decodificar(m[2])), h, O, h === O.h1 ? cortesDv : null, nombreDestino);
          const con = (p, nueva) => '<x14:dataValidation' + (nueva ? b.replace(/\s+xr:uid\s*=\s*"[^"]*"/, '') : b) + '>' + I.conXm(cuerpo, p.formulas).replace(I.RE_XMSQ, (y, c1, s, c3) => c1 + p.sq + c3) + '</x14:dataValidation>';
          let e = '';
          if (r.queda) { e += con(r.queda, false); n++; }
          if (r.llega) { if (h === O.h1) { e += con(r.llega, true); n++; } else llevar.x14dv.push(con(r.llega, true)); }
          return e;
        });
        return n ? '<x14:dataValidations' + a.replace(/\bcount\s*=\s*"\d+"/, 'count="' + n + '"') + '>' + v + '</x14:dataValidations>' : '';
      });
    }
    u = u.replace(/<ext\b[^>]*>\s*<\/ext>/g, '').replace(/<extLst>\s*<\/extLst>/g, '');
    return O.cf ? I.renumerar(u) : u;
  }

  // ---- anchos: se copia lo que la columna de origen dice en el archivo ------------------------------------------------------------------------------------------------------------------------
  function colsDe(xml) {
    const raiz = X().analizar(xml, { profundidad: 3 }), cols = X().hijo(raiz, 'cols'), l = [];
    if (cols) for (const k of X().hijos(cols, 'col')) l.push({ a: +k.attrs.min - 1, b: +k.attrs.max - 1, et: xml.slice(k.ini, k.fin), attrs: k.attrs });
    return { cols, l };
  }
  function escribirAnchos(x0, x1, anchos) {
    const q = colsDe(x0), d = colsDe(x1);
    const datosDe = (c) => q.l.find((k) => c >= k.a && c <= k.b) || null;
    const cambia = new Map(anchos.map((p) => [p.c, datosDe(p.de)]));   // columna de destino → lo de su origen, o null (ancho normal)
    const piezas = [];
    for (const k of d.l) {
      let i = k.a;
      for (let c = k.a; c <= k.b + 1; c++) {
        if (c <= k.b && !cambia.has(c)) continue;
        if (c > i) piezas.push({ a: i, b: c - 1, et: k.et });
        i = c + 1;
      }
    }
    for (const [c, de] of cambia) {
      const base = d.l.find((k) => c >= k.a && c <= k.b);
      if (!de) { if (base && base.attrs.style) piezas.push({ a: c, b: c, et: X().ponerAtributo(X().ponerAtributo(X().ponerAtributo(base.et, 'width', null), 'customWidth', null), 'hidden', null) }); continue; }
      let et = base ? base.et : '<col/>';
      et = X().ponerAtributo(et, 'width', de.attrs.width !== undefined ? String(de.attrs.width) : null);
      et = X().ponerAtributo(et, 'customWidth', de.attrs.customWidth !== undefined ? String(de.attrs.customWidth) : null);
      et = X().ponerAtributo(et, 'hidden', de.attrs.hidden !== undefined ? String(de.attrs.hidden) : null);
      piezas.push({ a: c, b: c, et });
    }
    piezas.sort((p, r) => p.a - r.a);
    const bloque = piezas.map((p) => X().ponerAtributo(X().ponerAtributo(p.et, 'min', String(p.a + 1)), 'max', String(p.b + 1))).join('');
    if (d.cols) return X().aplicar(x1, [{ ini: d.cols.ini, fin: d.cols.fin, texto: bloque ? '<cols>' + bloque + '</cols>' : '' }]);
    return bloque ? EstructuraLibro._ponerEnHoja(x1, 'cols', '<cols>' + bloque + '</cols>') : x1;
  }

  // ---- tablas -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  const REL_TABLA = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table', TIPO_TABLA = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';
  const NOMBRE_RE = (n) => new RegExp('(^|[^A-Za-z0-9_.])' + n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\[', 'g');
  async function clonarTablas(libro, O) {
    const P = PaqueteEditable.desdeZip(libro.zip), r0 = libro.hojas[O.h0].ruta, r1 = libro.hojas[O.h1].ruta;
    const rels = await P.relaciones(r0), partes = rels.filter((r) => r.tipo === REL_TABLA);
    // las que quedan enteras debajo se van
    if (O.tapadas.length) {
      const rels1 = await P.relaciones(r1);
      let x1 = await P.leer(r1);
      const quita = [], ids = new Set();
      for (const r of rels1.filter((q) => q.tipo === REL_TABLA)) {
        const t = await P.leer(r.destino), a = /<table\b[^>]*?\bref="([^"]*)"/.exec(t), g = a && LibroXlsx.decodificarRango(a[1]);
        if (g && O.tapadas.some((m) => m.ref.f0 === g.f0 && m.ref.c0 === g.c0 && m.ref.f1 === g.f1 && m.ref.c1 === g.c1)) { quita.push(r.destino); ids.add(r.id); x1 = x1.replace(new RegExp('<tablePart\\b[^>]*?\\bid\\s*=\\s*"' + r.id + '"[^>]*/>'), ''); }
      }
      x1 = x1.replace(/(<tableParts\b[^>]*?)\bcount="\d+"([^>]*>)([\s\S]*?)(<\/tableParts>)/, (m, a, b, c, d) => { const n = (c.match(/<tablePart\b/g) || []).length; return n ? a + 'count="' + n + '"' + b + c + d : ''; });
      P.escribir(r1, x1);
      await MoverNotas._quitarPartes(P, quita);
      await MoverNotas._quitarRelaciones(P, r1, ids);
    }
    let maxId = 0, k = 1;
    for (const n of P.rutas()) if (/^xl\/tables\/[^/]+\.xml$/i.test(n)) { const a = /<table\b[^>]*?\bid="(\d+)"/.exec(await P.leer(n)); if (a) maxId = Math.max(maxId, +a[1]); }
    for (const l of O.tablas) {
      let molde = null;
      for (const r of partes) { const t = await P.leer(r.destino); const a = t && /<table\b[^>]*?\bref="([^"]*)"/.exec(t), g = a && LibroXlsx.decodificarRango(a[1]); if (g && g.f0 === l.t.ref.f0 && g.c0 === l.t.ref.c0 && g.f1 === l.t.ref.f1 && g.c1 === l.t.ref.c1) { molde = t; break; } }
      if (molde === null) continue;
      const nombreViejo = l.t.nombreVisible || l.t.nombre, ref = MoverObjetos._textoRango(l.ref);
      let u = molde.replace(/(<table\b[^>]*?\bid=")\d+(")/, (m, a, c) => a + (++maxId) + c).replace(/(<table\b[^>]*?\bname=")[^"]*(")/, (m, a, c) => a + l.nombre + c).replace(/(<table\b[^>]*?\bdisplayName=")[^"]*(")/, (m, a, c) => a + l.nombre + c)
        .replace(/(<table\b[^>]*?\bref=")[^"]*(")/, (m, a, c) => a + ref + c).replace(/\s+xr3?:uid="[^"]*"/g, '').replace(/(<calculatedColumnFormula\b[^>]*>)([\s\S]*?)(<\/calculatedColumnFormula>)/g, (m, a, f, c) => a + X().decodificar(f).replace(NOMBRE_RE(nombreViejo), (y, q) => q + l.nombre + '[').replace(/&/g, '&amp;').replace(/</g, '&lt;') + c);
      const fin = l.t.totales ? { f0: l.ref.f0, c0: l.ref.c0, f1: l.ref.f1 - 1, c1: l.ref.c1 } : l.ref;
      u = u.replace(/(<autoFilter\b[^>]*?\bref=")[^"]*(")/, (m, a, c) => a + MoverObjetos._textoRango(fin) + c);
      while (P.existe('xl/tables/table' + k + '.xml')) k++;
      const ruta = 'xl/tables/table' + k + '.xml';
      P.escribir(ruta, u);
      await P.asegurarTipo('/' + ruta, TIPO_TABLA);
      const id = await P.agregarRelacion(r1, REL_TABLA, '../tables/table' + k + '.xml');
      const conR = MoverNotas._conPrefijoR(await P.leer(r1));
      let x = conR.xml;
      const parte = `<tablePart ${conR.p}:id="${id}"/>`;
      if (/<tableParts\b/.test(x)) x = x.replace(/<tableParts\b([^>]*?)\/>/, '<tableParts$1></tableParts>').replace('</tableParts>', () => parte + '</tableParts>').replace(/(<tableParts\b[^>]*?\bcount=")(\d+)(")/, (m, a, n, c) => a + (+n + 1) + c);
      else x = EstructuraLibro._ponerEnHoja(x, 'tableParts', `<tableParts count="1">${parte}</tableParts>`);
      P.escribir(r1, x);
    }
  }

  // ---- enlaces de lo pegado desde otro programa ---------------------------------------------------------------------------------------------------------------------------------------------
  async function enlacesExternos(libro, O) {
    const P = PaqueteEditable.desdeZip(libro.zip), r1 = libro.hojas[O.h1].ruta;
    const fuera = (e) => { const m = /\bref\s*=\s*"([^"]*)"/.exec(e), r = m && LibroXlsx.decodificarRango(m[1]); return r && dentro(r, O.D); };
    let x = (await P.leer(r1)).replace(/<hyperlink\b[^>]*?\/>/g, (e) => (fuera(e) ? '' : e)).replace(/<hyperlinks\b[^>]*>\s*<\/hyperlinks>/, '');
    const piezas = [], conR = MoverNotas._conPrefijoR(x);
    x = conR.xml;
    for (const q of O.externos || []) {
      const id = await P.agregarRelacion(r1, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', q.url);
      const k = r1.lastIndexOf('/'), rr = r1.slice(0, k + 1) + '_rels/' + r1.slice(k + 1) + '.rels', t = await P.leer(rr);
      P.escribir(rr, t.replace(new RegExp('(<Relationship\\b[^>]*?\\bId="' + id + '"[^>]*?)/>'), '$1 TargetMode="External"/>'));
      piezas.push('<hyperlink ref="' + LibroXlsx.numeroACol(q.c) + (q.f + 1) + '" ' + conR.p + ':id="' + id + '"/>');
    }
    if (piezas.length) {
      if (/<hyperlinks\b/.test(x)) x = x.replace('</hyperlinks>', () => piezas.join('') + '</hyperlinks>');
      else x = EstructuraLibro._ponerEnHoja(x, 'hyperlinks', '<hyperlinks>' + piezas.join('') + '</hyperlinks>');
    }
    P.escribir(r1, x);
  }

  // ---- todo junto --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  async function aplicarXlsx(bytes, O, opciones) {
    const libro = await LibroXlsx.abrir(bytes), zip = libro.zip;
    const leer = async (r) => { const e = zip.file(r); return e ? e.async('string') : null; };
    const r0 = libro.hojas[O.h0].ruta, r1 = libro.hojas[O.h1].ruta, nombre = libro.hojas[O.h1].nombre;
    if (O.cf || O.dv) {
      const llevar = { cf: [], dv: [], x14cf: [], x14dv: [] };
      const x0 = await leer(r0), u0 = reglas(x0, O.h0, O, nombre, llevar);
      if (O.h0 === O.h1) zip.file(r0, u0);
      else {
        zip.file(r0, u0);
        const x1 = await leer(r1);
        zip.file(r1, MI().recibir(reglas(x1, O.h1, O, nombre, llevar), llevar));
      }
    }
    if (O.anchos.length) zip.file(r1, escribirAnchos(await leer(r0), await leer(r1), O.anchos));
    if (O.notas || O.enlaces) await MoverNotas.copiar(libro, O);
    if (O.tablas.length) await clonarTablas(libro, O);
    if (O.externos || O.limpiarEnlaces) await enlacesExternos(libro, O);
    let b = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    if (O.altos.length) {
      const ed = await LibroEditable.abrir(bytes, opciones), filas = ed.hojas[O.h0].filas;
      for (const q of O.altos) { const x = filas && filas.get(q.de); if (x && x.alto !== null && x.alto !== undefined && (x.personalizada || x.oculta)) b = await EstructuraLibro.altoFilas(b, O.h1, q.f, q.f, x.oculta ? 0 : x.alto); }
    }
    for (const m of O.comb.quitar) b = await EstructuraLibro.separar(b, O.h1, m);
    if (O.comb.poner.length) b = await ponerCombinadas(b, O.h1, O.comb.poner);
    return b;
  }
  // pegar una combinada no borra lo que hay debajo (Excel lo guarda oculto)
  async function ponerCombinadas(bytes, h, lista) {
    const libro = await LibroXlsx.abrir(bytes), ruta = libro.hojas[h].ruta, zip = libro.zip;
    let x = await zip.file(ruta).async('string');
    const el = lista.map((m) => '<mergeCell ref="' + MoverObjetos._textoRango(m) + '"/>').join('');
    if (/<mergeCells\b/.test(x)) x = x.replace('</mergeCells>', () => el + '</mergeCells>').replace(/(<mergeCells\b[^>]*?\bcount=")(\d+)(")/, (y, a, n, c) => a + (+n + lista.length) + c);
    else x = EstructuraLibro._ponerEnHoja(x, 'mergeCells', '<mergeCells count="' + lista.length + '">' + el + '</mergeCells>');
    zip.file(ruta, x);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }
  async function aplicar(bytes, O, opciones) {
    if (O.formato === 'ods') return PegadoOds.aplicar(bytes, O, opciones);
    return aplicarXlsx(bytes, O, opciones);
  }

  globalThis.PegadoObjetos = { planear, aplicar, _planRegla: planRegla, _cortesDe: cortesDe, _unirCeldas: unirCeldas };
})();

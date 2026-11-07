/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Cortar y pegar: lo que no es una celda y también sigue a lo movido (nombres definidos, formato condicional, validación, gráficos y celdas combinadas). Trabaja sobre el archivo:
// aplica el movimiento de celdas (LibroEditable.opMover), guarda y reescribe lo demás. Medido en Excel: pruebas/oraculo/excel-medir-mover-objetos.ps1.
//   await MoverObjetos.aplicar(bytes, h0, rango, h1, fila, columna, { idioma }) → Uint8Array
// Formato condicional: viaja con la celda y lo que estaba en el destino se queda. Validación: viaja con la celda y la del destino se pierde. En las dos, la fórmula es relativa a la
// esquina de su rango; si solo algunas de sus celdas apuntan a lo movido, la regla se parte. Las reglas de Excel 2010 (x14) no se parten y no cambian de hoja.
// En .ods la validación y el formato condicional solo corren sus referencias y su rango si se mueve entero (EstructuraOds.reescribirReferencias).
(function () {
  const X = () => XmlEditar;
  const TOPE_CELDAS = 20000;   // una regla con fórmula relativa se parte celda por celda: si es enorme, se deja como está
  const BASE_F = 524288, BASE_C = 8192;
  const textoRango = (r) => { const a = LibroXlsx.numeroACol(r.c0) + (r.f0 + 1), b = LibroXlsx.numeroACol(r.c1) + (r.f1 + 1); return a === b ? a : a + ':' + b; };
  const leerRangos = (sq) => String(sq).split(/\s+/).filter(Boolean).map((p) => LibroXlsx.decodificarRango(p.replace(/\$/g, '')));
  const interseccion = (a, b) => { const f0 = Math.max(a.f0, b.f0), f1 = Math.min(a.f1, b.f1), c0 = Math.max(a.c0, b.c0), c1 = Math.min(a.c1, b.c1); return f0 <= f1 && c0 <= c1 ? { f0, c0, f1, c1 } : null; };
  function restar(a, b) {
    const i = interseccion(a, b);
    if (!i) return [a];
    const l = [];
    if (a.f0 < i.f0) l.push({ f0: a.f0, c0: a.c0, f1: i.f0 - 1, c1: a.c1 });
    if (i.c0 > a.c0) l.push({ f0: i.f0, c0: a.c0, f1: i.f1, c1: i.c0 - 1 });
    if (i.c1 < a.c1) l.push({ f0: i.f0, c0: i.c1 + 1, f1: i.f1, c1: a.c1 });
    if (i.f1 < a.f1) l.push({ f0: i.f1 + 1, c0: a.c0, f1: a.f1, c1: a.c1 });
    return l;
  }
  const esquina = (l) => l.reduce((e, r) => ({ f: Math.min(e.f, r.f0), c: Math.min(e.c, r.c0) }), { f: Infinity, c: Infinity });
  const conRelativas = (f) => {
    try { return XlsFormulas._tokenizar(String(f).replace(/\[[^\]]*\]/g, (m) => '_'.repeat(m.length))).some((t) => t.k === 'ref' && (!t.fa0 || !t.ca0 || (t.modo === 'area' && (!t.fa1 || !t.ca1)))); } catch (e) { return false; }
  };

  // Una regla (rangos y fórmulas, relativas a su esquina) después de mover → [{ h, sqref, formulas }]. pisa: lo que cae en el destino se pierde (validación).
  function planRegla(sq, formulas, h, M, ft, pisa) {
    const partes = leerRangos(sq);
    if (!partes.length || partes.some((r) => !r)) return [{ h, sqref: sq, formulas: formulas.map((f) => ft(f, h, h)) }];
    const S = { f0: M.f0, c0: M.c0, f1: M.f1, c1: M.c1 }, D = { f0: M.f0 + M.df, c0: M.c0 + M.dc, f1: M.f1 + M.df, c1: M.c1 + M.dc };
    const antes = esquina(partes);
    let quedan = [];
    const llevadas = [];
    for (const r of partes) {
      let resto = [r];
      if (h === M.h0) {
        const i = interseccion(r, S);
        if (i) { llevadas.push({ f0: i.f0 + M.df, c0: i.c0 + M.dc, f1: i.f1 + M.df, c1: i.c1 + M.dc }); resto = restar(r, S); }
      }
      if (pisa && h === M.h1) resto = resto.flatMap((x) => restar(x, D));
      quedan.push(...resto);
    }
    const grupos = M.h0 === M.h1 || h !== M.h0 ? [{ h, rangos: quedan.concat(llevadas) }] : [{ h, rangos: quedan }, { h: M.h1, rangos: llevadas }];
    const salida = [];
    for (const g of grupos) {
      if (!g.rangos.length) continue;
      const e = esquina(g.rangos);
      const fb = formulas.map((f) => LibroXlsx.desplazarFormula(f, e.f - antes.f, e.c - antes.c));
      const celdas = g.rangos.reduce((n, r) => n + (r.f1 - r.f0 + 1) * (r.c1 - r.c0 + 1), 0);
      const relativa = celdas > 1 && fb.some(conRelativas);
      if (!relativa || celdas > TOPE_CELDAS) { salida.push({ h: g.h, sqref: g.rangos.map(textoRango).join(' '), formulas: fb.map((f) => (relativa ? f : ft(f, h, g.h))) }); continue; }
      // cada celda con su fórmula: las que quedan iguales relativas vuelven a ser una sola regla
      const porClave = new Map();
      let cambio = false;
      for (const r of g.rangos) for (let f = r.f0; f <= r.f1; f++) for (let c = r.c0; c <= r.c1; c++) {
        const propia = fb.map((x) => LibroXlsx.desplazarFormula(x, f - e.f, c - e.c)), nueva = propia.map((x) => ft(x, h, g.h));
        if (nueva.some((x, k) => x !== propia[k])) cambio = true;
        const normal = nueva.map((x) => LibroXlsx.desplazarFormula(x, BASE_F - f, BASE_C - c)), k = normal.join('\u0000');
        if (!porClave.has(k)) porClave.set(k, { normal, rangos: [] });
        const l = porClave.get(k).rangos, u = l[l.length - 1];
        if (u && u.c0 === c && u.c1 === c && u.f1 === f - 1) u.f1 = f; else l.push({ f0: f, c0: c, f1: f, c1: c });
      }
      if (!cambio) { salida.push({ h: g.h, sqref: g.rangos.map(textoRango).join(' '), formulas: fb }); continue; }
      for (const v of porClave.values()) {
        const t = esquina(v.rangos);
        salida.push({ h: g.h, sqref: v.rangos.map(textoRango).join(' '), formulas: v.normal.map((x) => LibroXlsx.desplazarFormula(x, t.f - BASE_F, t.c - BASE_C)) });
      }
    }
    return salida;
  }

  const RE_FORMULA = /(<formula[12]?\b[^>]*(?<!\/)>)([\s\S]*?)(<\/formula[12]?>)/g;
  const conFormulas = (cuerpo, l) => { let k = 0; return cuerpo.replace(RE_FORMULA, (x, a, f, c) => a + X().esc(l[k++]) + c); };
  const renumerar = (u) => {
    const RE_P = /(<(?:x14:)?cfRule\b[^>]*?\bpriority\s*=\s*")(\d+)(")/g;
    const ps = [...u.matchAll(RE_P)].map((m, i) => [parseInt(m[2], 10), i]).sort((x, y) => x[0] - y[0] || x[1] - y[1]);
    const nueva = [];
    ps.forEach(([, i], k) => { nueva[i] = k + 1; });
    let i = 0;
    return u.replace(RE_P, (x, a, p, c) => a + nueva[i++] + c);
  };

  // La hoja h: reescribe sus reglas; lo que cambia de hoja queda en `llevar`
  function hojaXml(xml, h, M, ft, llevar) {
    let u = xml, partida = false;
    u = u.replace(/<conditionalFormatting\b([^>]*?)\bsqref\s*=\s*"([^"]*)"([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g, (todo, a, sq, b, hijos) => {
      const l = planRegla(X().decodificar(sq), [...hijos.matchAll(RE_FORMULA)].map((m) => X().decodificar(m[2])), h, M, ft, false);
      if (l.length !== 1) partida = true;
      return l.map((p) => {
        const e = `<conditionalFormatting${a}sqref="${p.sqref}"${b}>${conFormulas(hijos, p.formulas)}</conditionalFormatting>`;
        if (p.h === h) return e;
        llevar.cf.push(e);
        return '';
      }).join('');
    });
    let quedan = null;
    u = u.replace(/<dataValidation\b([^>]*?)(\/>|>([\s\S]*?)<\/dataValidation>)/g, (todo, attrs, resto, hijos) => {
      const ms = /\bsqref\s*=\s*"([^"]*)"/.exec(attrs);
      if (!ms) return todo;
      const l = planRegla(X().decodificar(ms[1]), hijos ? [...hijos.matchAll(RE_FORMULA)].map((m) => X().decodificar(m[2])) : [], h, M, ft, true);
      quedan = quedan === null ? 0 : quedan;
      return l.map((p) => {
        const a2 = attrs.replace(ms[0], `sqref="${p.sqref}"`);
        const e = hijos === undefined ? `<dataValidation${a2}/>` : `<dataValidation${a2}>${conFormulas(hijos, p.formulas)}</dataValidation>`;
        if (p.h === h) { quedan++; return e; }
        llevar.dv.push(e);
        return '';
      }).join('');
    });
    if (quedan === 0) u = u.replace(/<dataValidations\b[^>]*>\s*<\/dataValidations>/, '');
    else if (quedan !== null) u = u.replace(/(<dataValidations\b[^>]*?\bcount\s*=\s*")\d+(")/, (x, a, c) => a + quedan + c);
    // Excel 2010: mismo rango y fórmulas, sin partir y sin cambiar de hoja
    if (u.indexOf('x14:') >= 0) {
      const x14 = (cuerpo) => cuerpo.replace(/(<xm:f>)([\s\S]*?)(<\/xm:f>)/g, (x, a, f, c) => a + X().esc(ft(X().decodificar(f), h, h)) + c);
      const conRango = (cuerpo) => {
        const sq = /(<xm:sqref>)([^<]*)(<\/xm:sqref>)/.exec(cuerpo);
        if (!sq || M.h0 !== M.h1 || h !== M.h0) return cuerpo;
        const p = planRegla(X().decodificar(sq[2]), [], h, M, ft, false)[0];
        return p ? cuerpo.replace(sq[0], sq[1] + p.sqref + sq[3]) : cuerpo;
      };
      u = u.replace(/<x14:conditionalFormatting\b[^>]*>[\s\S]*?<\/x14:conditionalFormatting>/g, (t) => conRango(x14(t)));
      u = u.replace(/<x14:dataValidation\b[^>]*>[\s\S]*?<\/x14:dataValidation>/g, (t) => conRango(x14(t)));
      u = u.replace(/<x14:sparkline\s*>[\s\S]*?<\/x14:sparkline>/g, x14);
    }
    return partida ? renumerar(u) : u;
  }
  // lo que llega de otra hoja: al final de su lista, o donde corresponde si la hoja no tiene ninguna
  function recibir(xml, llevar) {
    let u = xml;
    if (llevar.cf.length) {
      const k = u.lastIndexOf('</conditionalFormatting>');
      if (k >= 0) { const p = k + '</conditionalFormatting>'.length; u = u.slice(0, p) + llevar.cf.join('') + u.slice(p); }
      else u = EstructuraLibro._ponerEnHoja(u, 'conditionalFormatting', llevar.cf.join(''));
      u = renumerar(u);
    }
    if (llevar.dv.length) {
      if (/<dataValidations\b/.test(u)) u = u.replace('</dataValidations>', () => llevar.dv.join('') + '</dataValidations>').replace(/(<dataValidations\b[^>]*?\bcount\s*=\s*")(\d+)(")/, (x, a, n, c) => a + (+n + llevar.dv.length) + c);
      else u = EstructuraLibro._ponerEnHoja(u, 'dataValidations', `<dataValidations count="${llevar.dv.length}">${llevar.dv.join('')}</dataValidations>`);
    }
    return u;
  }

  async function reescribirXlsx(bytes, M, nombresHoja) {
    const libro = await LibroXlsx.abrir(bytes), zip = libro.zip;
    const ft = (t, hF, hNueva) => LibroEditable._formulaTrasMover(t, hF, hNueva, M, nombresHoja);
    const leer = async (r) => { const e = zip.file(r); return e ? e.async('string') : null; };
    const llevar = { cf: [], dv: [] };
    for (let h = 0; h < libro.hojas.length; h++) {
      const ruta = libro.hojas[h].ruta, xml = ruta && libro.hojas[h].tipo === 'hoja' ? await leer(ruta) : null;
      if (xml === null) continue;
      const u = hojaXml(xml, h, M, ft, llevar);
      if (u !== xml) zip.file(ruta, u);
    }
    if (llevar.cf.length || llevar.dv.length) {
      const ruta = libro.hojas[M.h1].ruta, xml = await leer(ruta);
      zip.file(ruta, recibir(xml, llevar));
    }
    const rutas = [];
    zip.forEach((n) => { if (/^xl\/workbook\.xml$|^xl\/charts\/chart[^/]*\.xml$|^xl\/charts\/chartEx[^/]*\.xml$/i.test(n)) rutas.push(n); });
    for (const ruta of rutas) {
      const t = await leer(ruta);
      const u = /workbook\.xml$/i.test(ruta)
        ? t.replace(/(<definedName\b[^>]*>)([\s\S]*?)(<\/definedName>)/g, (x, a, f, c) => a + X().esc(ft(X().decodificar(f), -1, -1)) + c)
        : t.replace(/(<(?:c|cx):f\b[^>]*>)([\s\S]*?)(<\/(?:c|cx):f>)/g, (x, a, f, c) => a + X().esc(ft(X().decodificar(f), -1, -1)) + c);
      if (u !== t) zip.file(ruta, u);
    }
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  async function aplicar(bytes, h0, r, h1, f, c, opciones) {
    const ed = await LibroEditable.abrir(bytes, opciones);
    const m = ed.opMover(h0, r, h1, f, c);
    if (m.error) throw new Error(m.error);
    for (const op of m.ops) ed.aplicar(op);
    ed.terminar();
    const nombresHoja = ed.libro.hojas.map((x) => x.nombre);
    const esOds = ed.formato === 'ods', EST = esOds ? EstructuraOds : EstructuraLibro;
    let b = await ed.guardar();
    const rw = (t, enObjetivo) => LibroEditable._formulaTrasMover(t, enObjetivo ? m.M.h0 : -1, enObjetivo ? m.M.h0 : -1, m.M, nombresHoja);
    b = esOds ? await EST.reescribirReferencias(b, h0, rw) : await reescribirXlsx(b, m.M, nombresHoja);
    for (const g of m.combinadas.mover) b = await EST.separar(b, h0, g);
    for (const g of m.combinadas.separarDestino) b = await EST.separar(b, h1, g);
    for (const g of m.combinadas.mover) b = await EST.combinar(b, h1, { f0: g.f0 + m.M.df, c0: g.c0 + m.M.dc, f1: g.f1 + m.M.df, c1: g.c1 + m.M.dc });
    return b;
  }

  globalThis.MoverObjetos = { aplicar, _planRegla: planRegla };
})();

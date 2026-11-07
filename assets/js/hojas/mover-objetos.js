/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Cortar y pegar: lo que no es una celda y también sigue a lo movido (nombres, formato condicional, validación, gráficos, minigráficos, combinadas, notas, hipervínculos y tablas). Trabaja
// sobre el archivo: aplica el movimiento de celdas (LibroEditable.opMover), guarda y reescribe lo demás. Medido en Excel: pruebas/oraculo/excel-medir-mover-objetos.ps1 y excel-mover-limites.ps1.
//   await MoverObjetos.aplicar(bytes, h0, rango, h1, fila, columna, { idioma }) → Uint8Array
// Formato condicional: viaja con la celda y lo que estaba en el destino se queda. Validación: viaja con la celda y la del destino se pierde. En las dos, la fórmula es relativa a la
// esquina de su rango; si solo algunas de sus celdas apuntan a lo movido, la regla se parte; una regla que cambia de hoja sigue diciendo lo mismo, ahora en su hoja nueva.
// Las notas e hipervínculos van en mover-notas.js, las tablas en mover-tablas.js y lo de .ods en mover-ods.js.
(function () {
  const X = () => XmlEditar;
  const MAX_ZONAS = 200000, MAX_PARTES = 5000;   // zonas de una regla que se miran y partes en que se parte: un archivo hostil no cuelga la página; pasado el tope, la regla queda como está
  const BASE_F = 524288, BASE_C = 8192, FILAS = 1048576, COLS = 16384;
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
  const tapar = (f) => String(f).replace(/"(?:[^"]|"")*"|\[[^\]]*\]/g, (m) => (m[0] === '"' ? m : '_'.repeat(m.length)));
  const tokensRef = (f) => { try { return XlsFormulas._tokenizar(tapar(f)).filter((t) => t.k === 'ref' && (t.modo === 'celda' || t.modo === 'area')); } catch (e) { return []; } };
  const conRelativas = (f) => tokensRef(f).some((t) => !t.fa0 || !t.ca0 || (t.modo === 'area' && (!t.fa1 || !t.ca1)));
  // una regla que llega a su hoja no la cita: Excel escribe $A$1, no Otra!$A$1
  function sinHojaPropia(f, nombre) {
    let s = String(f);
    const toks = tokensRef(f);
    for (let k = toks.length - 1; k >= 0; k--) {
      const t = toks[k];
      if (!t.hoja || t.hoja.a !== t.hoja.b || t.hoja.a.toLowerCase() !== nombre.toLowerCase()) continue;
      const o = s.slice(t.i, t.i + t.l);
      s = s.slice(0, t.i) + o.slice(o.lastIndexOf('!') + 1) + s.slice(t.i + t.l);
    }
    return s;
  }

  // ---- partir una regla: cada celda tiene su propia fórmula, y las que quedan iguales (relativas a la celda) son una sola regla -----------------------------------------------------
  // No se mira celda por celda: lo que le pasa a una referencia cambia solo donde cruza los bordes de lo movido, así que el rango se corta en zonas por esos cortes y cada zona se resuelve de
  // una vez. Una zona con esquinas distintas (un rango cuyo borde queda dentro de lo movido) se sigue partiendo por la mitad.
  function cortes(f, e, M, D) {
    const fs = new Set(), cs = new Set();
    const umbralesF = [M.f0, M.f1, D.f0, D.f1, M.f0 - M.df, M.f1 - M.df, M.f1 + 1 - M.df, M.f0 - 1 - M.df];
    const umbralesC = [M.c0, M.c1, D.c0, D.c1, M.c0 - M.dc, M.c1 - M.dc, M.c1 + 1 - M.dc, M.c0 - 1 - M.dc];
    const anotar = (s, a, abs, base, umbrales, max) => {
      if (abs) return;
      for (const v of umbrales) for (const w of [v - 1, v, v + 1, v + 2]) s.add(w - a + base);
      s.add(base - a); s.add(base - a + max);   // donde la referencia da la vuelta a la hoja
    };
    for (const t of tokensRef(f)) {
      const area = t.modo === 'area';
      anotar(fs, t.f0, t.fa0, e.f, umbralesF, FILAS); anotar(cs, t.c0, t.ca0, e.c, umbralesC, COLS);
      if (area) { anotar(fs, t.f1, t.fa1, e.f, umbralesF, FILAS); anotar(cs, t.c1, t.ca1, e.c, umbralesC, COLS); }
    }
    return { fs, cs };
  }
  const tramos = (a0, a1, s) => { const l = [...s].filter((v) => v > a0 && v <= a1).sort((x, y) => x - y); const r = []; let p = a0; for (const v of l) { r.push([p, v - 1]); p = v; } r.push([p, a1]); return r; };
  function unir(rs) {
    const v = new Map();   // primero hacia abajo, después hacia el costado
    for (const r of rs) { const k = r.c0 + ':' + r.c1; if (!v.has(k)) v.set(k, []); v.get(k).push(r); }
    const m1 = [];
    for (const l of v.values()) { l.sort((a, b) => a.f0 - b.f0); for (const r of l) { const u = m1.length && m1[m1.length - 1]; if (u && u.c0 === r.c0 && u.c1 === r.c1 && u.f1 + 1 === r.f0) u.f1 = r.f1; else m1.push(Object.assign({}, r)); } }
    const h = new Map();
    for (const r of m1) { const k = r.f0 + ':' + r.f1; if (!h.has(k)) h.set(k, []); h.get(k).push(r); }
    const m2 = [];
    for (const l of h.values()) { l.sort((a, b) => a.c0 - b.c0); for (const r of l) { const u = m2.length && m2[m2.length - 1]; if (u && u.f0 === r.f0 && u.f1 === r.f1 && u.c1 + 1 === r.c0) u.c1 = r.c1; else m2.push(Object.assign({}, r)); } }
    return m2;
  }
  // → [{ normal, rangos }] por cada fórmula distinta (normal: relativa a una celda fija), o null si es demasiado
  function zonas(listas, e, fb, M, ftCelda) {
    const D = { f0: M.f0 + M.df, c0: M.c0 + M.dc, f1: M.f1 + M.df, c1: M.c1 + M.dc };
    const fs = new Set(), cs = new Set();
    for (const f of fb) { const c = cortes(f, e, M, D); c.fs.forEach((x) => fs.add(x)); c.cs.forEach((x) => cs.add(x)); }
    const porClave = new Map();
    let vistas = 0, cambio = false;
    const clave = (f, c, dF, dC) => {
      const propia = fb.map((x) => LibroXlsx.desplazarFormula(x, f - e.f, c - e.c)), nueva = propia.map(ftCelda);
      if (nueva.some((x, k) => x !== propia[k])) cambio = true;
      const normal = nueva.map((x) => LibroXlsx.desplazarFormula(x, BASE_F - (f + dF), BASE_C - (c + dC)));
      return { k: normal.join('\u0000'), normal };
    };
    const resolver = (r, dF, dC) => {
      if (++vistas > MAX_ZONAS) return false;
      const a = clave(r.f0, r.c0, dF, dC), b = clave(r.f1, r.c1, dF, dC), c = clave(r.f0, r.c1, dF, dC), d = clave(r.f1, r.c0, dF, dC);
      if (a.k === b.k && a.k === c.k && a.k === d.k) { if (!porClave.has(a.k)) porClave.set(a.k, { normal: a.normal, rangos: [] }); porClave.get(a.k).rangos.push({ f0: r.f0 + dF, c0: r.c0 + dC, f1: r.f1 + dF, c1: r.c1 + dC }); return true; }
      if (r.f1 - r.f0 >= r.c1 - r.c0) { const m = (r.f0 + r.f1) >> 1; return resolver({ f0: r.f0, c0: r.c0, f1: m, c1: r.c1 }, dF, dC) && resolver({ f0: m + 1, c0: r.c0, f1: r.f1, c1: r.c1 }, dF, dC); }
      const m = (r.c0 + r.c1) >> 1;
      return resolver({ f0: r.f0, c0: r.c0, f1: r.f1, c1: m }, dF, dC) && resolver({ f0: r.f0, c0: m + 1, f1: r.f1, c1: r.c1 }, dF, dC);
    };
    for (const l of listas) for (const g of l.rangos) for (const tf of tramos(g.f0, g.f1, fs)) for (const tc of tramos(g.c0, g.c1, cs)) if (!resolver({ f0: tf[0], c0: tc[0], f1: tf[1], c1: tc[1] }, l.dF, l.dC)) return null;
    if (porClave.size > MAX_PARTES) return null;
    return { cambio, grupos: [...porClave.values()].map((v) => ({ normal: v.normal, rangos: unir(v.rangos) })) };
  }

  // Una regla (rangos y fórmulas, relativas a su esquina) después de mover → [{ h, sqref, formulas }]. pisa (validación): las zonas del destino a las que llega otra validación; además, una regla entera
  // debajo del destino se pierde. A medias, si lo que llega no trae validación, se queda (medido).
  // ft(f, hF, hNueva): formulaTrasMover; nombreHoja(h): para quitar la cita de la hoja a la que llega.
  function planRegla(sq, formulas, h, M, ft, pisa, nombreHoja) {
    const partes = leerRangos(sq);
    if (!partes.length || partes.some((r) => !r)) return [{ h, sqref: sq, formulas: formulas.map((f) => ft(f, h, h)) }];
    const S = { f0: M.f0, c0: M.c0, f1: M.f1, c1: M.c1 }, D = { f0: M.f0 + M.df, c0: M.c0 + M.dc, f1: M.f1 + M.df, c1: M.c1 + M.dc };
    const antes = esquina(partes);
    let quedan = [];
    const llevadasOrigen = [];
    for (const r of partes) {
      let resto = [r];
      if (h === M.h0) {
        const i = interseccion(r, S);
        if (i) { llevadasOrigen.push(i); resto = restar(r, S); }
      }
      quedan.push(...resto);
    }
    if (pisa && h === M.h1) {
      if (quedan.every((x) => x.f0 >= D.f0 && x.f1 <= D.f1 && x.c0 >= D.c0 && x.c1 <= D.c1)) quedan = [];
      else for (const z of pisa) quedan = quedan.flatMap((x) => restar(x, z));
    }
    const corrido = (r) => ({ f0: r.f0 + M.df, c0: r.c0 + M.dc, f1: r.f1 + M.df, c1: r.c1 + M.dc });
    const llevadas = llevadasOrigen.map(corrido);
    const salida = [];
    const relativa = formulas.some(conRelativas);
    // en su misma hoja, una celda movida conserva lo que decía su fórmula: =B2 sigue siendo =B2, salvo que B2 también se mueva
    const propias = [{ rangos: quedan, dF: 0, dC: 0 }];
    if (M.h0 === M.h1 && llevadas.length) propias.push({ rangos: llevadasOrigen, dF: M.df, dC: M.dc });
    const finales = propias.flatMap((l) => l.rangos.map((r) => (l.dF || l.dC ? corrido(r) : r)));
    if (finales.length) {
      const e = esquina(finales), texto = finales.map(textoRango).join(' ');
      const ftH = (f) => ft(f, h, h);
      if (!relativa) salida.push({ h, sqref: texto, formulas: formulas.map(ftH) });
      else {
        const z = zonas(propias, antes, formulas, M, ftH);
        if (!z || (!z.cambio && propias.length === 1)) salida.push({ h, sqref: texto, formulas: formulas.map((f) => LibroXlsx.desplazarFormula(f, e.f - antes.f, e.c - antes.c)) });
        else for (const v of z.grupos) {
          const t = esquina(v.rangos);
          salida.push({ h, sqref: v.rangos.map(textoRango).join(' '), formulas: v.normal.map((x) => LibroXlsx.desplazarFormula(x, t.f - BASE_F, t.c - BASE_C)) });
        }
      }
    }
    // a otra hoja: la regla sigue diciendo lo mismo, relativa a su esquina, y lo que no cita hoja ahora es de la hoja nueva
    if (M.h0 !== M.h1 && llevadas.length) {
      const e = esquina(llevadas), corr = (f) => LibroXlsx.desplazarFormula(f, M.df, M.dc);
      const ftG = (f) => (nombreHoja ? sinHojaPropia(ft(f, h, h), nombreHoja(M.h1)) : ft(f, h, h));
      const eo = esquina(llevadasOrigen);
      const base = formulas.map((f) => LibroXlsx.desplazarFormula(f, eo.f - antes.f, eo.c - antes.c));
      salida.push({ h: M.h1, sqref: llevadas.map(textoRango).join(' '), formulas: base.map((f) => ftG(corr(f))) });
      void e;
    }
    return salida;
  }

  // lo que las validaciones de lo movido tapan en el destino
  function pisadasDe(sqrefs, M) {
    const S = { f0: M.f0, c0: M.c0, f1: M.f1, c1: M.c1 }, l = [];
    for (const sq of sqrefs) for (const r of leerRangos(sq)) { const i = r && interseccion(r, S); if (i) l.push({ f0: i.f0 + M.df, c0: i.c0 + M.dc, f1: i.f1 + M.df, c1: i.c1 + M.dc }); }
    return l;
  }

  const RE_FORMULA = /(<formula[12]?\b[^>]*(?<!\/)>)([\s\S]*?)(<\/formula[12]?>)/g;
  const RE_XMF = /(<xm:f>)([\s\S]*?)(<\/xm:f>)/g, RE_XMSQ = /(<xm:sqref>)([^<]*)(<\/xm:sqref>)/;
  const URI_CF = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}', URI_DV = '{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}';
  const NS_X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main', NS_XM = 'http://schemas.microsoft.com/office/excel/2006/main';
  const conFormulas = (cuerpo, l) => { let k = 0; return cuerpo.replace(RE_FORMULA, (x, a, f, c) => a + X().esc(l[k++]) + c); };
  const conXm = (cuerpo, l) => { let k = 0; return cuerpo.replace(RE_XMF, (x, a, f, c) => a + X().esc(l[k++]) + c); };
  const renumerar = (u) => {
    const RE_P = /(<(?:x14:)?cfRule\b[^>]*?\bpriority\s*=\s*")(\d+)(")/g;
    const ps = [...u.matchAll(RE_P)].map((m, i) => [parseInt(m[2], 10), i]).sort((x, y) => x[0] - y[0] || x[1] - y[1]);
    const nueva = [];
    ps.forEach(([, i], k) => { nueva[i] = k + 1; });
    let i = 0;
    return u.replace(RE_P, (x, a, p, c) => a + nueva[i++] + c);
  };
  // una regla partida necesita su propio identificador (el de Excel 2010)
  let contadorId = 0;
  const idNuevo = (id) => { contadorId++; const h = id.replace(/[^0-9A-Fa-f]/g, '').toUpperCase(), n = ((parseInt(h.slice(-8), 16) + contadorId * 2654435761) >>> 0).toString(16).toUpperCase().padStart(8, '0'), t = h.slice(0, -8) + n; return t.slice(0, 8) + '-' + t.slice(8, 12) + '-' + t.slice(12, 16) + '-' + t.slice(16, 20) + '-' + t.slice(20); };
  const idsNuevos = (cuerpo) => cuerpo.replace(/(<x14:cfRule\b[^>]*?\bid\s*=\s*"\{)([0-9A-Fa-f-]{36})(\})/g, (y, a, id, c) => a + idNuevo(id) + c);

  // La hoja h: reescribe sus reglas; lo que cambia de hoja queda en `llevar`
  function hojaXml(xml, h, M, ft, llevar, nombreHoja, pisadas) {
    let u = xml, partida = false;
    u = u.replace(/<conditionalFormatting\b([^>]*?)\bsqref\s*=\s*"([^"]*)"([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g, (todo, a, sq, b, hijos) => {
      const l = planRegla(X().decodificar(sq), [...hijos.matchAll(RE_FORMULA)].map((m) => X().decodificar(m[2])), h, M, ft, false, nombreHoja);
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
      const l = planRegla(X().decodificar(ms[1]), hijos ? [...hijos.matchAll(RE_FORMULA)].map((m) => X().decodificar(m[2])) : [], h, M, ft, pisadas, nombreHoja);
      quedan = quedan === null ? 0 : quedan;
      return l.map((p, k) => {
        const a2 = (k ? attrs.replace(/\s+xr:uid\s*=\s*"[^"]*"/, '') : attrs).replace(ms[0], `sqref="${p.sqref}"`);
        const e = hijos === undefined ? `<dataValidation${a2}/>` : `<dataValidation${a2}>${conFormulas(hijos, p.formulas)}</dataValidation>`;
        if (p.h === h) { quedan++; return e; }
        llevar.dv.push(e);
        return '';
      }).join('');
    });
    if (quedan === 0) u = u.replace(/<dataValidations\b[^>]*>\s*<\/dataValidations>/, '');
    else if (quedan !== null) u = u.replace(/(<dataValidations\b[^>]*?\bcount\s*=\s*")\d+(")/, (x, a, c) => a + quedan + c);
    // Excel 2010: las mismas reglas, en las extensiones (las que citan otra hoja, las barras de datos, las listas de otra hoja) y los minigráficos
    if (u.indexOf('x14:') >= 0) {
      const partesX14 = (cuerpo, pisa) => {
        const sq = RE_XMSQ.exec(cuerpo);
        if (!sq) return null;
        return planRegla(X().decodificar(sq[2]), [...cuerpo.matchAll(RE_XMF)].map((m) => X().decodificar(m[2])), h, M, ft, pisa, nombreHoja);
      };
      const conRango = (cuerpo, p) => conXm(cuerpo, p.formulas).replace(RE_XMSQ, (x, a, s, c) => a + p.sqref + c);
      u = u.replace(/<x14:conditionalFormattings\b([^>]*)>([\s\S]*?)<\/x14:conditionalFormattings>/g, (todo, a, hijos) => {
        let n = 0;
        const v = hijos.replace(/<x14:conditionalFormatting\b([^>]*)>([\s\S]*?)<\/x14:conditionalFormatting>/g, (x, b, cuerpo) => {
          const l = partesX14(cuerpo, false);
          if (!l) { n++; return x; }
          if (l.length !== 1) partida = true;
          return l.map((p, k) => {
            const e = '<x14:conditionalFormatting' + b + '>' + (k ? idsNuevos(conRango(cuerpo, p)) : conRango(cuerpo, p)) + '</x14:conditionalFormatting>';
            if (p.h === h) { n++; return e; }
            llevar.x14cf.push(e);
            return '';
          }).join('');
        });
        return n ? '<x14:conditionalFormattings' + a + '>' + v + '</x14:conditionalFormattings>' : '';
      });
      u = u.replace(/<x14:dataValidations\b([^>]*)>([\s\S]*?)<\/x14:dataValidations>/g, (todo, a, hijos) => {
        let n = 0;
        const v = hijos.replace(/<x14:dataValidation\b([^>]*)>([\s\S]*?)<\/x14:dataValidation>/g, (x, b, cuerpo) => {
          const l = partesX14(cuerpo, pisadas);
          if (!l) { n++; return x; }
          return l.map((p, k) => {
            const e = '<x14:dataValidation' + (k ? b.replace(/\s+xr:uid\s*=\s*"[^"]*"/, '') : b) + '>' + conRango(cuerpo, p) + '</x14:dataValidation>';
            if (p.h === h) { n++; return e; }
            llevar.x14dv.push(e);
            return '';
          }).join('');
        });
        return n ? '<x14:dataValidations' + a.replace(/\bcount\s*=\s*"\d+"/, 'count="' + n + '"') + '>' + v + '</x14:dataValidations>' : '';
      });
      // un minigráfico es de su celda: se mueve con ella en la misma hoja (a otra hoja se queda, medido), y el que estaba en el destino se pierde
      const S = { f0: M.f0, c0: M.c0, f1: M.f1, c1: M.c1 }, D = { f0: M.f0 + M.df, c0: M.c0 + M.dc, f1: M.f1 + M.df, c1: M.c1 + M.dc };
      const adentro = (r, a) => r && r.f0 >= a.f0 && r.f1 <= a.f1 && r.c0 >= a.c0 && r.c1 <= a.c1;
      u = u.replace(/<x14:sparklineGroups\b([^>]*)>([\s\S]*?)<\/x14:sparklineGroups>/g, (todo, a, hijos) => {
        let grupos = 0;
        const v = hijos.replace(/<x14:sparklineGroup\b([^>]*)>([\s\S]*?)<\/x14:sparklineGroup>/g, (x, b, cuerpo) => {
          let vivos = 0;
          const c2 = cuerpo.replace(/<x14:sparkline\s*>([\s\S]*?)<\/x14:sparkline>/g, (y, dentro) => {
            const sq = RE_XMSQ.exec(dentro), r = sq && LibroXlsx.decodificarRango(X().decodificar(sq[2]));
            const va = h === M.h0 && M.h0 === M.h1 && adentro(r, S);
            if (h === M.h1 && !va && adentro(r, D)) return '';
            vivos++;
            let w = dentro.replace(RE_XMF, (z, c1, f, c3) => c1 + X().esc(ft(X().decodificar(f), h, h)) + c3);
            if (va) w = w.replace(RE_XMSQ, (z, c1, s, c3) => c1 + textoRango({ f0: r.f0 + M.df, c0: r.c0 + M.dc, f1: r.f1 + M.df, c1: r.c1 + M.dc }) + c3);
            return '<x14:sparkline>' + w + '</x14:sparkline>';
          });
          if (!vivos) return '';
          grupos++;
          return '<x14:sparklineGroup' + b + '>' + c2 + '</x14:sparklineGroup>';
        });
        return grupos ? '<x14:sparklineGroups' + a + '>' + v + '</x14:sparklineGroups>' : '';
      });
      u = u.replace(/<ext\b[^>]*>\s*<\/ext>/g, '').replace(/<extLst>\s*<\/extLst>/g, '');
    }
    return partida ? renumerar(u) : u;
  }
  // lo que llega de otra hoja: al final de su lista, o donde corresponde si la hoja no tiene ninguna
  function recibir(xml, llevar) {
    let u = xml;
    if (llevar.cf.length) {
      const k = u.lastIndexOf('</conditionalFormatting>');
      if (k >= 0 && k > u.indexOf('</sheetData>')) { const p = k + '</conditionalFormatting>'.length; u = u.slice(0, p) + llevar.cf.join('') + u.slice(p); }
      else u = EstructuraLibro._ponerEnHoja(u, 'conditionalFormatting', llevar.cf.join(''));
    }
    if (llevar.dv.length) {
      if (/<dataValidations\b/.test(u)) u = u.replace('</dataValidations>', () => llevar.dv.join('') + '</dataValidations>').replace(/(<dataValidations\b[^>]*?\bcount\s*=\s*")(\d+)(")/, (x, a, n, c) => a + (+n + llevar.dv.length) + c);
      else u = EstructuraLibro._ponerEnHoja(u, 'dataValidations', `<dataValidations count="${llevar.dv.length}">${llevar.dv.join('')}</dataValidations>`);
    }
    const enExt = (uri, ext) => {
      if (/<extLst>/.test(u)) u = u.replace('</extLst>', () => ext + '</extLst>');
      else u = EstructuraLibro._ponerEnHoja(u, 'extLst', '<extLst>' + ext + '</extLst>');
    };
    if (llevar.x14cf.length) {
      if (/<x14:conditionalFormattings\b/.test(u)) u = u.replace('</x14:conditionalFormattings>', () => llevar.x14cf.join('') + '</x14:conditionalFormattings>');
      else enExt(URI_CF, `<ext uri="${URI_CF}" xmlns:x14="${NS_X14}"><x14:conditionalFormattings>${llevar.x14cf.join('')}</x14:conditionalFormattings></ext>`);
    }
    if (llevar.x14dv.length) {
      if (/<x14:dataValidations\b/.test(u)) u = u.replace('</x14:dataValidations>', () => llevar.x14dv.join('') + '</x14:dataValidations>').replace(/(<x14:dataValidations\b[^>]*?\bcount\s*=\s*")(\d+)(")/, (x, a, n, c) => a + (+n + llevar.x14dv.length) + c);
      else enExt(URI_DV, `<ext uri="${URI_DV}" xmlns:x14="${NS_X14}"><x14:dataValidations count="${llevar.x14dv.length}" xmlns:xm="${NS_XM}">${llevar.x14dv.join('')}</x14:dataValidations></ext>`);
    }
    return llevar.cf.length || llevar.x14cf.length ? renumerar(u) : u;
  }

  async function reescribirXlsx(bytes, M, nombresHoja, planTablas) {
    const libro = await LibroXlsx.abrir(bytes), zip = libro.zip;
    const ft = (t, hF, hNueva) => LibroEditable._formulaTrasMover(t, hF, hNueva, M, nombresHoja);
    const leer = async (r) => { const e = zip.file(r); return e ? e.async('string') : null; };
    const llevar = { cf: [], dv: [], x14cf: [], x14dv: [] };
    const nombreHoja = (h) => nombresHoja[h];
    const xml0 = await leer(libro.hojas[M.h0].ruta);
    const pisadas = pisadasDe([...xml0.matchAll(/<dataValidation\b[^>]*?\bsqref\s*=\s*"([^"]*)"/g), ...xml0.matchAll(/<x14:dataValidation\b[\s\S]*?<xm:sqref>([^<]*)<\/xm:sqref>/g)].map((m) => X().decodificar(m[1])), M);
    for (let h = 0; h < libro.hojas.length; h++) {
      const ruta = libro.hojas[h].ruta, xml = ruta && libro.hojas[h].tipo === 'hoja' ? await leer(ruta) : null;
      if (xml === null) continue;
      const u = hojaXml(xml, h, M, ft, llevar, nombreHoja, pisadas);
      if (u !== xml) zip.file(ruta, u);
    }
    if (llevar.cf.length || llevar.dv.length || llevar.x14cf.length || llevar.x14dv.length) {
      const ruta = libro.hojas[M.h1].ruta, xml = await leer(ruta);
      zip.file(ruta, recibir(xml, llevar));
    }
    await MoverNotas.aplicar(libro, M);
    if (planTablas) await MoverTablas.escribir(libro, planTablas, M, ft);
    const fin = planTablas ? (t) => planTablas.formula(t, -1) : (t) => t;
    const rutas = [];
    zip.forEach((n) => { if (/^xl\/workbook\.xml$|^xl\/charts\/chart[^/]*\.xml$|^xl\/charts\/chartEx[^/]*\.xml$/i.test(n)) rutas.push(n); });
    for (const ruta of rutas) {
      const t = await leer(ruta);
      const u = /workbook\.xml$/i.test(ruta)
        ? t.replace(/(<definedName\b[^>]*>)([\s\S]*?)(<\/definedName>)/g, (x, a, f, c) => a + X().esc(fin(ft(X().decodificar(f), -1, -1))) + c)
        : t.replace(/(<(?:c|cx):f\b[^>]*>)([\s\S]*?)(<\/(?:c|cx):f>)/g, (x, a, f, c) => a + X().esc(fin(ft(X().decodificar(f), -1, -1))) + c);
      if (u !== t) zip.file(ruta, u);
    }
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  async function aplicar(bytes, h0, r, h1, f, c, opciones) {
    const ed = await LibroEditable.abrir(bytes, opciones);
    const m = ed.opMover(h0, r, h1, f, c);
    if (m.error) throw new Error(m.error);
    for (const op of m.ops) ed.aplicar(op);
    const esOds = ed.formato === 'ods', EST = esOds ? EstructuraOds : EstructuraLibro;
    const nombresHoja = ed.libro.hojas.map((x) => x.nombre);
    const planTablas = esOds ? null : MoverTablas.planear(ed, m, nombresHoja, opciones && opciones.idioma);
    ed.terminar();
    let b;
    if (esOds) {
      const extras = await MoverOds.leerExtras(bytes, h0, h1, m.M);
      b = await ed.guardar();
      b = await MoverOds.aplicar(b, m.M, nombresHoja, extras);
    } else {
      // los encabezados y totales los cuida la tabla, que se escribe enseguida
      for (const x of ed.hojas) if (x && !x.sinCeldas) x.tablas = [];
      b = await ed.guardar();
      b = await reescribirXlsx(b, m.M, nombresHoja, planTablas);
    }
    for (const g of m.combinadas.mover) b = await EST.separar(b, h0, g);
    for (const g of m.combinadas.separarDestino) b = await EST.separar(b, h1, g);
    for (const g of m.combinadas.mover) b = await EST.combinar(b, h1, { f0: g.f0 + m.M.df, c0: g.c0 + m.M.dc, f1: g.f1 + m.M.df, c1: g.c1 + m.M.dc });
    return valoresDeGraficos(b, esOds, opciones);
  }
  // un gráfico dibuja los valores que guarda: si su rango se estiró o se achicó, hay que ponerlos al día
  async function valoresDeGraficos(b, esOds, opciones) {
    const G = esOds ? globalThis.GraficosOds : globalThis.GraficosLibro;
    if (!G) return b;
    const ed = await LibroEditable.abrir(b, opciones);
    let hay = false;
    ed.libro.zip.forEach((n) => { if (/^xl\/charts\/chart[^/]*\.xml$|^Object \d+\/content\.xml$/i.test(n)) hay = true; });
    if (!hay) return b;
    ed.recalcularTodo();
    if (!(await G.refrescarCaches(ed.libro.zip, G.lectorDe(ed)))) return b;
    return ed.libro.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  globalThis.MoverObjetos = { aplicar, _planRegla: planRegla, _pisadasDe: pisadasDe, _textoRango: textoRango, _esquina: esquina,
    _interno: { leerRangos, restar, esquina, unir, textoRango, sinHojaPropia, RE_FORMULA, RE_XMF, RE_XMSQ, conFormulas, conXm, renumerar, idNuevo, recibir } };
})();

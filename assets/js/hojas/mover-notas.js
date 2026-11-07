/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Cortar y pegar en .xlsx: las notas, los comentarios en hilo y los hipervínculos. Medido en Excel (pruebas/oraculo/excel-mover-limites.ps1):
// una nota o un hilo viajan con su celda, también a otra hoja, y los del destino se pierden aunque la celda que llega no traiga ninguno. Un hipervínculo viaja si está entero dentro de lo
// movido y se pierde si queda entero debajo; a medias se queda como está. El destino interno de un hipervínculo no sigue a las celdas.
//   await MoverNotas.aplicar(libro, M)   (libro: LibroXlsx abierto; escribe en su zip)
(function () {
  const X = () => XmlEditar;
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  const REL_NOTAS = REL + 'comments', REL_VML = REL + 'vmlDrawing', REL_HIPER = REL + 'hyperlink', REL_HILOS = 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment';
  const TIPO_NOTAS = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml', TIPO_HILOS = 'application/vnd.ms-excel.threadedcomments+xml';
  const RE_NOTA = /<comment\b([^>]*?)(?:\/>|>[\s\S]*?<\/comment>)/g, RE_HILO = /<threadedComment\b([^>]*?)(?:\/>|>[\s\S]*?<\/threadedComment>)/g, RE_FORMA = /<((?:\w+:)?)shape\b[^>]*>[\s\S]*?<\/\1shape>/g;
  const celdaDe = (attrs) => { const m = /\bref\s*=\s*"([A-Za-z]{1,3}\d+)"/.exec(attrs); return m ? LibroXlsx.decodificarDireccion(m[1]) : null; };
  const dentro = (p, r) => p && p.f >= r.f0 && p.f <= r.f1 && p.c >= r.c0 && p.c <= r.c1;
  const conRef = (e, p) => e.replace(/(\bref\s*=\s*")[A-Za-z]{1,3}\d+(")/, (x, a, c) => a + LibroXlsx.numeroACol(p.c) + (p.f + 1) + c);
  // la celda de una forma del dibujo de notas (las que no son notas se quedan donde están)
  function celdaForma(s) {
    if (!/<(?:\w+:)?ClientData\b[^>]*?\bObjectType\s*=\s*["']Note["']/.test(s)) return null;
    const f = /<(?:\w+:)?Row>\s*(\d+)\s*<\//.exec(s), c = /<(?:\w+:)?Column>\s*(\d+)\s*<\//.exec(s);
    return f && c ? { f: +f[1], c: +c[1] } : null;
  }
  function formaMovida(s, p, df, dc) {
    let u = s.replace(/(<(?:\w+:)?Row>\s*)\d+/, (x, a) => a + p.f).replace(/(<(?:\w+:)?Column>\s*)\d+/, (x, a) => a + p.c);
    return u.replace(/(<(?:\w+:)?Anchor>)([^<]*)(<\/)/, (x, a, v, c) => {
      const q = v.split(',');
      if (q.length !== 8) return x;
      const dF = Math.max(df, -Math.min(+q[2], +q[6])), dC = Math.max(dc, -Math.min(+q[0], +q[4]));   // el cuadro no sale de la hoja
      for (const i of [2, 6]) q[i] = q[i].replace(/\d+/, (n) => String(+n + dF));
      for (const i of [0, 4]) q[i] = q[i].replace(/\d+/, (n) => String(+n + dC));
      return a + q.join(',') + c;
    });
  }

  async function quitarRelaciones(P, ruta, ids) {
    const k = ruta.lastIndexOf('/'), rr = ruta.slice(0, k + 1) + '_rels/' + ruta.slice(k + 1) + '.rels', t = await P.leer(rr);
    if (t === null) return;
    const u = t.replace(/<Relationship\b[^>]*?\/>/g, (r) => { const id = /\bId\s*=\s*"([^"]*)"/.exec(r); return id && ids.has(id[1]) ? '' : r; });
    if (!/<Relationship\b/.test(u)) P.quitar(rr); else if (u !== t) P.escribir(rr, u);
  }
  async function quitarPartes(P, rutas) {
    for (const r of rutas) P.quitar(r);
    const ct = await P.leer('[Content_Types].xml'), idx = new Set(rutas.map((r) => '/' + r.toLowerCase()));
    P.escribir('[Content_Types].xml', ct.replace(/<Override\b[^>]*?\/>/g, (o) => { const p = /\bPartName\s*=\s*"([^"]*)"/.exec(o); return p && idx.has(p[1].toLowerCase()) ? '' : o; }));
  }
  const prefijoR = (xml) => { const m = /\bxmlns:(\w+)\s*=\s*"http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships"/.exec(xml.slice(0, 4000)); return m ? m[1] : null; };
  function conPrefijoR(xml) {
    const p = prefijoR(xml);
    if (p) return { xml, p };
    const raiz = X().analizar(xml, { profundidad: 1 });
    return { xml: X().aplicar(xml, [{ ini: raiz.ini, fin: raiz.finApertura, texto: X().ponerAtributo(xml.slice(raiz.ini, raiz.finApertura), 'xmlns:r', REL.replace(/\/$/, '')) }]), p: 'r' };
  }
  function libre(P, carpeta, base, ext) { let n = 1; while (P.existe(carpeta + base + n + ext)) n++; return carpeta + base + n + ext; }
  const relativa = (ruta) => '../' + ruta.replace(/^xl\//, '');

  async function partesDe(P, ruta) {
    const rels = await P.relaciones(ruta), xml = await P.leer(ruta);
    const ld = /<legacyDrawing\b[^>]*?\bid\s*=\s*"([^"]*)"/.exec(xml);
    const com = rels.find((r) => r.tipo === REL_NOTAS), hil = rels.find((r) => r.tipo === REL_HILOS), vml = ld ? rels.find((r) => r.id === ld[1]) : rels.find((r) => r.tipo === REL_VML);
    return { rels, com, hil, vml, tCom: com ? await P.leer(com.destino) : null, tHil: hil ? await P.leer(hil.destino) : null, tVml: vml ? await P.leer(vml.destino) : null };
  }
  // lo que queda vacío se va, con su relación y su <legacyDrawing>
  async function guardar(P, ruta, q) {
    const notas = q.tCom !== null && RE_NOTA.test(q.tCom); RE_NOTA.lastIndex = 0;
    const hilos = q.tHil !== null && /<threadedComment\b/.test(q.tHil);
    const otras = q.tVml !== null && (q.tVml.match(RE_FORMA) || []).some((s) => !/ObjectType\s*=\s*["']Note["']/.test(s));
    const borrar = [], ids = new Set();
    if (q.com) { if (notas) P.escribir(q.com.destino, q.tCom); else { borrar.push(q.com.destino); ids.add(q.com.id); } }
    if (q.hil) { if (hilos) P.escribir(q.hil.destino, q.tHil); else { borrar.push(q.hil.destino); ids.add(q.hil.id); } }
    if (q.vml) {
      if (notas || otras) P.escribir(q.vml.destino, q.tVml);
      else {
        borrar.push(q.vml.destino); ids.add(q.vml.id);
        const xml = await P.leer(ruta);
        P.escribir(ruta, xml.replace(/<legacyDrawing\b[^>]*?\/>/, ''));
      }
    }
    if (borrar.length) { await quitarPartes(P, borrar); await quitarRelaciones(P, ruta, ids); }
  }

  // agrega elementos a la lista de notas de otra hoja, con su autor
  function conNotas(destino, notas, origen) {
    const autoresDe = (t) => { const m = /<authors>([\s\S]*?)<\/authors>/.exec(t); return m ? m[1].match(/<author\b[^>]*?(?:\/>|>[\s\S]*?<\/author>)/g) || [] : []; };
    const de = autoresDe(origen), a = autoresDe(destino);
    const movidas = notas.map((e) => e.replace(/\s+data-autor="[^"]*"/, '').replace(/(\bauthorId\s*=\s*")(\d+)(")/, (x, p, i, c) => {
      const propio = /\sdata-autor="([^"]*)"/.exec(e);
      const au = propio ? '<author>' + propio[1] + '</author>' : de[+i];
      if (au === undefined) return x;
      let k = a.indexOf(au);
      if (k < 0) { a.push(au); k = a.length - 1; }
      return p + k + c;
    }));
    let u = destino.replace(/<authors>[\s\S]*?<\/authors>|<authors\s*\/>/, () => '<authors>' + a.join('') + '</authors>');
    if (/<commentList\s*\/>/.test(u)) return u.replace(/<commentList\s*\/>/, () => '<commentList>' + movidas.join('') + '</commentList>');
    return u.replace('</commentList>', () => movidas.join('') + '</commentList>');
  }
  const conIdsFormas = (formas, base, vistos) => formas.map((s) => {
    let k = 1;
    while (vistos.has(base * 1024 + k)) k++;
    vistos.add(base * 1024 + k);
    return s.replace(/(\bid\s*=\s*"_x0000_s)\d+(")/, (x, a, c) => a + (base * 1024 + k) + c);
  });
  const idsVml = (t) => new Set([...t.matchAll(/\bid\s*=\s*"_x0000_s(\d+)"/g)].map((m) => +m[1]));
  const baseVml = (t) => { const m = /<o:idmap\b[^>]*?\bdata\s*=\s*"(\d+)/.exec(t); return m ? +m[1] : 1; };

  async function notas(P, libro, M) {
    const r0 = libro.hojas[M.h0].ruta, r1 = libro.hojas[M.h1].ruta;
    const S = { f0: M.f0, c0: M.c0, f1: M.f1, c1: M.c1 }, D = { f0: M.f0 + M.df, c0: M.c0 + M.dc, f1: M.f1 + M.df, c1: M.c1 + M.dc };
    const q0 = await partesDe(P, r0), q1 = M.h0 === M.h1 ? q0 : await partesDe(P, r1);
    if (!q0.com && !q0.hil && !q1.com && !q1.hil) return;
    const nueva = (p) => ({ f: p.f + M.df, c: p.c + M.dc });
    if (M.h0 === M.h1) {
      // en la misma hoja: lo que está debajo del destino se va y lo movido cambia de lugar
      const uno = (e, attrs) => { const p = celdaDe(attrs); return dentro(p, S) ? conRef(e, nueva(p)) : dentro(p, D) ? '' : e; };
      if (q0.tCom !== null) q0.tCom = q0.tCom.replace(RE_NOTA, uno);
      if (q0.tHil !== null) q0.tHil = q0.tHil.replace(RE_HILO, uno);
      if (q0.tVml !== null) q0.tVml = q0.tVml.replace(RE_FORMA, (s) => { const p = celdaForma(s); return dentro(p, S) ? formaMovida(s, nueva(p), M.df, M.dc) : dentro(p, D) ? '' : s; });
      await guardar(P, r0, q0);
      return;
    }
    const llevar = { com: [], hil: [], vml: [] };
    if (q0.tCom !== null) q0.tCom = q0.tCom.replace(RE_NOTA, (e, attrs) => { const p = celdaDe(attrs); if (!dentro(p, S)) return e; llevar.com.push(conRef(e, nueva(p))); return ''; });
    if (q0.tHil !== null) q0.tHil = q0.tHil.replace(RE_HILO, (e, attrs) => { const p = celdaDe(attrs); if (!dentro(p, S)) return e; llevar.hil.push(conRef(e, nueva(p))); return ''; });
    if (q0.tVml !== null) q0.tVml = q0.tVml.replace(RE_FORMA, (s) => { const p = celdaForma(s); if (!dentro(p, S)) return s; llevar.vml.push(formaMovida(s, nueva(p), M.df, M.dc)); return ''; });
    const fuera = (e, attrs) => (dentro(celdaDe(attrs), D) ? '' : e);
    if (q1.tCom !== null) q1.tCom = q1.tCom.replace(RE_NOTA, fuera);
    if (q1.tHil !== null) q1.tHil = q1.tHil.replace(RE_HILO, fuera);
    if (q1.tVml !== null) q1.tVml = q1.tVml.replace(RE_FORMA, (s) => (dentro(celdaForma(s), D) ? '' : s));
    await agregar(P, q0, q1, r1, llevar);
    await guardar(P, r0, q0);
    await guardar(P, r1, q1);
  }

  async function agregar(P, q0, q1, r1, llevar) {
    if (llevar.com.length) {
      const origen = q0.tCom;
      if (q1.tCom !== null) q1.tCom = conNotas(q1.tCom, llevar.com, origen);
      else {
        const raiz = /<comments\b[^>]*>/.exec(origen)[0];
        const ruta = libre(P, 'xl/', 'comments', '.xml');
        q1.tCom = conNotas(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${raiz}<authors></authors><commentList></commentList></comments>`, llevar.com, origen);
        P.escribir(ruta, q1.tCom);
        await P.asegurarTipo('/' + ruta, TIPO_NOTAS);
        q1.com = { destino: ruta, id: await P.agregarRelacion(r1, REL_NOTAS, relativa(ruta)) };
      }
      if (q1.tVml !== null) {
        const base = baseVml(q1.tVml), formas = conIdsFormas(llevar.vml, base, idsVml(q1.tVml));
        const k = q1.tVml.lastIndexOf('</xml>');
        q1.tVml = q1.tVml.slice(0, k) + formas.join('') + q1.tVml.slice(k);
      } else {
        // un dibujo nuevo: el molde del de origen, con su propio número de formas
        const usados = new Set();
        for (const n of P.rutas()) if (/\.vml$/i.test(n)) usados.add(baseVml(await P.leer(n)));
        let base = 1; while (usados.has(base)) base++;
        const molde = (await P.leer(q0.vml.destino)).replace(RE_FORMA, '').replace(/(<o:idmap\b[^>]*?\bdata\s*=\s*")[^"]*(")/, (x, a, c) => a + base + c);
        const k = molde.lastIndexOf('</xml>');
        q1.tVml = molde.slice(0, k) + conIdsFormas(llevar.vml, base, new Set()).join('') + molde.slice(k);
        const ruta = libre(P, 'xl/drawings/', 'vmlDrawing', '.vml');
        P.escribir(ruta, q1.tVml);
        q1.vml = { destino: ruta, id: await P.agregarRelacion(r1, REL_VML, relativa(ruta)) };
        const conR = conPrefijoR(await P.leer(r1));
        P.escribir(r1, EstructuraLibro._ponerEnHoja(conR.xml, 'legacyDrawing', `<legacyDrawing ${conR.p}:id="${q1.vml.id}"/>`));
      }
    }
    if (llevar.hil.length) {
      if (q1.tHil !== null) q1.tHil = q1.tHil.replace('</ThreadedComments>', () => llevar.hil.join('') + '</ThreadedComments>');
      else {
        const raiz = /<ThreadedComments\b[^>]*>/.exec(q0.tHil)[0];
        const ruta = libre(P, 'xl/threadedComments/', 'threadedComment', '.xml');
        q1.tHil = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${raiz}${llevar.hil.join('')}</ThreadedComments>`;
        P.escribir(ruta, q1.tHil);
        await P.asegurarTipo('/' + ruta, TIPO_HILOS);
        q1.hil = { destino: ruta, id: await P.agregarRelacion(r1, REL_HILOS, relativa(ruta)) };
      }
    }
  }

  async function hipervinculos(P, libro, M) {
    const r0 = libro.hojas[M.h0].ruta, r1 = libro.hojas[M.h1].ruta;
    const S = { f0: M.f0, c0: M.c0, f1: M.f1, c1: M.c1 }, D = { f0: M.f0 + M.df, c0: M.c0 + M.dc, f1: M.f1 + M.df, c1: M.c1 + M.dc };
    const adentro = (r, a) => r && r.f0 >= a.f0 && r.f1 <= a.f1 && r.c0 >= a.c0 && r.c1 <= a.c1;
    const RE = /<hyperlink\b[^>]*?\/>/g;
    const rangoDe = (e) => { const m = /\bref\s*=\s*"([^"]*)"/.exec(e); return m ? LibroXlsx.decodificarRango(m[1]) : null; };
    const conRango = (e, r) => e.replace(/(\bref\s*=\s*")[^"]*(")/, (x, a, c) => a + MoverObjetos._textoRango({ f0: r.f0 + M.df, c0: r.c0 + M.dc, f1: r.f1 + M.df, c1: r.c1 + M.dc }) + c);
    const sinVacios = (t) => t.replace(/<hyperlinks\b[^>]*>\s*<\/hyperlinks>/, '');
    let x0 = await P.leer(r0);
    if (M.h0 === M.h1) {
      if (x0.indexOf('<hyperlink') < 0) return;
      const u = sinVacios(x0.replace(RE, (e) => { const r = rangoDe(e); return adentro(r, S) ? conRango(e, r) : adentro(r, D) ? '' : e; }));
      if (u !== x0) P.escribir(r0, u);
      return;
    }
    const llevar = [];
    x0 = sinVacios(x0.replace(RE, (e) => { const r = rangoDe(e); if (!adentro(r, S)) return e; llevar.push(conRango(e, r)); return ''; }));
    let x1 = await P.leer(r1);
    x1 = sinVacios(x1.replace(RE, (e) => (adentro(rangoDe(e), D) ? '' : e)));
    if (llevar.length) {
      // el destino de un enlace externo es una relación de la hoja: va con él
      const rels = await P.relaciones(r0), quitar = new Set(), piezas = [];
      const conR = conPrefijoR(x1);
      x1 = conR.xml;
      for (const e of llevar) {
        const m = /\b(\w+):id\s*=\s*"([^"]*)"/.exec(e), rel = m && rels.find((q) => q.id === m[2]);
        if (!rel) { piezas.push(e); continue; }
        const id = await P.agregarRelacion(r1, REL_HIPER, rel.destino);
        const k = r1.lastIndexOf('/'), rr = r1.slice(0, k + 1) + '_rels/' + r1.slice(k + 1) + '.rels';
        const t = await P.leer(rr);
        P.escribir(rr, t.replace(new RegExp('(<Relationship\\b[^>]*?\\bId="' + id + '"[^>]*?)/>'), rel.externa ? '$1 TargetMode="External"/>' : '$1/>'));
        quitar.add(m[2]);
        piezas.push(e.replace(m[0], conR.p + ':id="' + id + '"'));
      }
      if (/<hyperlinks\b/.test(x1)) x1 = x1.replace('</hyperlinks>', () => piezas.join('') + '</hyperlinks>');
      else x1 = EstructuraLibro._ponerEnHoja(x1, 'hyperlinks', '<hyperlinks>' + piezas.join('') + '</hyperlinks>');
      if (quitar.size) await quitarRelaciones(P, r0, quitar);
    }
    P.escribir(r0, x0);
    P.escribir(r1, x1);
  }

  // ---- copiar (no mover): la nota, el hilo o el enlace de cada celda de origen llega a su destino, con lo de abajo cortado. O.pares: { sf, sc, df, dc }; con «omitir blancos» solo se pisa donde llega uno
  async function copiar(libro, O) {
    const P = PaqueteEditable.desdeZip(libro.zip), r0 = libro.hojas[O.h0].ruta, r1 = libro.hojas[O.h1].ruta;
    if (O.notas) await copiarNotas(P, O, r0, r1);
    if (O.enlaces) await copiarEnlaces(P, O, r0, r1);
  }
  async function copiarNotas(P, O, r0, r1) {
    const q0 = await partesDe(P, r0), q1 = O.h0 === O.h1 ? q0 : await partesDe(P, r1);
    if (!q0.com && !q0.hil && !q1.com && !q1.hil) return;
    const destinos = new Map();
    for (const p of O.pares) { const k = p.sf + ',' + p.sc; if (!destinos.has(k)) destinos.set(k, []); destinos.get(k).push(p); }
    const llevar = { com: [], hil: [], vml: [] }, quita = new Set(), idsHilo = new Map();
    const de = (p) => (p ? destinos.get(p.f + ',' + p.c) || [] : []);
    if (q0.tHil !== null) for (const m of q0.tHil.matchAll(RE_HILO)) {
      const p = celdaDe(m[1]);
      for (const d of de(p)) {
        quita.add(d.df + ',' + d.dc);
        const k = p.f + ',' + p.c + '>' + d.df + ',' + d.dc;
        const id = /\bid\s*=\s*"\{([0-9A-Fa-f-]{36})\}"/.exec(m[1]);
        const nuevo = id ? MoverObjetos._interno.idNuevo(id[1]) : null;
        if (id) idsHilo.set(k + '|' + id[1].toUpperCase(), nuevo);
        let e = conRef(m[0], { f: d.df, c: d.dc });
        if (id) e = e.replace(id[0], id[0].replace(id[1], nuevo));
        const padre = /\bparentId\s*=\s*"\{([0-9A-Fa-f-]{36})\}"/.exec(m[1]);
        if (padre && idsHilo.has(k + '|' + padre[1].toUpperCase())) e = e.replace(padre[0], padre[0].replace(padre[1], idsHilo.get(k + '|' + padre[1].toUpperCase())));
        llevar.hil.push(e);
      }
    }
    if (q0.tCom !== null) for (const m of q0.tCom.matchAll(RE_NOTA)) {
      const p = celdaDe(m[1]);
      for (const d of de(p)) {
        quita.add(d.df + ',' + d.dc);
        let e = conRef(m[0], { f: d.df, c: d.dc }).replace(/\bxr:uid\s*=\s*"\{([0-9A-Fa-f-]{36})\}"/, (x, u) => 'xr:uid="{' + MoverObjetos._interno.idNuevo(u) + '}"');
        // la nota que acompaña a un hilo se llama tc={id del hilo}: la copia, con el id de la suya
        const k = p.f + ',' + p.c + '>' + d.df + ',' + d.dc + '|';
        for (const [kk, n] of idsHilo) if (kk.startsWith(k)) { e = e.replace(/^<((?:\w+:)?)comment\b/, (x) => x + ' data-autor="tc={' + n + '}"'); break; }
        llevar.com.push(e);
      }
    }
    if (q0.tVml !== null) for (const s of q0.tVml.match(RE_FORMA) || []) {
      const p = celdaForma(s);
      for (const d of de(p)) llevar.vml.push(formaMovida(s, { f: d.df, c: d.dc }, d.df - p.f, d.dc - p.c));
    }
    const origenCom = q0.tCom, moldeVml = q0.vml;
    const fuera = (p) => p && (O.blancos ? quita.has(p.f + ',' + p.c) : dentro(p, O.D));
    if (q1.tCom !== null) q1.tCom = q1.tCom.replace(RE_NOTA, (e, attrs) => (fuera(celdaDe(attrs)) ? '' : e));
    if (q1.tHil !== null) q1.tHil = q1.tHil.replace(RE_HILO, (e, attrs) => (fuera(celdaDe(attrs)) ? '' : e));
    if (q1.tVml !== null) q1.tVml = q1.tVml.replace(RE_FORMA, (s) => (fuera(celdaForma(s)) ? '' : s));
    void origenCom; void moldeVml;
    await agregar(P, O.h0 === O.h1 ? Object.assign({}, q0) : q0, q1, r1, llevar);
    await guardar(P, r1, q1);
  }
  async function copiarEnlaces(P, O, r0, r1) {
    const RE = /<hyperlink\b[^>]*?\/>/g;
    const rangoDe = (e) => { const m = /\bref\s*=\s*"([^"]*)"/.exec(e); return m ? LibroXlsx.decodificarRango(m[1]) : null; };
    const x0 = await P.leer(r0);
    const llegan = [], tocadas = [];
    for (const m of x0.match(RE) || []) {
      const r = rangoDe(m);
      if (!r) continue;
      const l = O.pares.filter((p) => p.sf >= r.f0 && p.sf <= r.f1 && p.sc >= r.c0 && p.sc <= r.c1);
      if (!l.length) continue;
      for (const g of MoverObjetos._interno.unir(l.map((p) => ({ f0: p.df, c0: p.dc, f1: p.df, c1: p.dc })))) { llegan.push({ e: m, g }); tocadas.push(g); }
    }
    let x1 = O.h0 === O.h1 ? x0 : await P.leer(r1);
    const adentro = (r, a) => r && r.f0 >= a.f0 && r.f1 <= a.f1 && r.c0 >= a.c0 && r.c1 <= a.c1;
    const fuera = (e) => { const r = rangoDe(e); return O.blancos ? tocadas.some((g) => adentro(r, g)) : adentro(r, O.D); };
    x1 = x1.replace(RE, (e) => (fuera(e) ? '' : e)).replace(/<hyperlinks\b[^>]*>\s*<\/hyperlinks>/, '');
    if (llegan.length) {
      const rels = await P.relaciones(r0), piezas = [], conR = conPrefijoR(x1);
      x1 = conR.xml;
      for (const { e, g } of llegan) {
        let u = e.replace(/(\bref\s*=\s*")[^"]*(")/, (x, a, c) => a + MoverObjetos._textoRango(g) + c).replace(/\s+xr:uid\s*=\s*"[^"]*"/, '');
        const m = /\b(\w+):id\s*=\s*"([^"]*)"/.exec(u), rel = m && rels.find((q) => q.id === m[2]);
        if (rel && O.h0 !== O.h1) {
          const id = await P.agregarRelacion(r1, REL_HIPER, rel.destino);
          const k = r1.lastIndexOf('/'), rr = r1.slice(0, k + 1) + '_rels/' + r1.slice(k + 1) + '.rels', t = await P.leer(rr);
          P.escribir(rr, t.replace(new RegExp('(<Relationship\\b[^>]*?\\bId="' + id + '"[^>]*?)/>'), rel.externa ? '$1 TargetMode="External"/>' : '$1/>'));
          u = u.replace(m[0], conR.p + ':id="' + id + '"');
        }
        piezas.push(u);
      }
      if (/<hyperlinks\b/.test(x1)) x1 = x1.replace('</hyperlinks>', () => piezas.join('') + '</hyperlinks>');
      else x1 = EstructuraLibro._ponerEnHoja(x1, 'hyperlinks', '<hyperlinks>' + piezas.join('') + '</hyperlinks>');
    }
    P.escribir(r1, x1);
  }

  async function aplicar(libro, M) {
    const P = PaqueteEditable.desdeZip(libro.zip);
    await notas(P, libro, M);
    await hipervinculos(P, libro, M);
  }

  globalThis.MoverNotas = { aplicar, copiar, _quitarPartes: quitarPartes, _quitarRelaciones: quitarRelaciones, _conPrefijoR: conPrefijoR };
})();

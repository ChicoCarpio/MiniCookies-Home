/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Edición del paquete de un libro de Excel (.xlsx, .xlsm) con ZipPropio, sin DOM: quitar y reordenar hojas dejando el libro que Excel abre sin reparar
// (etapa 7). Reemplaza a reordenarHojasExcel de ooxml-paquete.js, que no trataba las tablas dinámicas, las segmentaciones ni las escalas de tiempo de
// las hojas quitadas (Excel pedía reparar el archivo). Lo que no se toca queda byte a byte igual: las entradas intactas del zip se copian comprimidas.
//
// Uso:
//   const r = await XlsxPaquete.reordenarHojas(zip, ['Hoja3', 'Hoja1']);   // deja solo esas hojas, en ese orden; zip: ZipPropio ya abierto
//   // r = { quitadas: [nombres], conservadas: [nombres], partesBorradas: [rutas] }
// Hace, para cada hoja quitada: su entrada, su relación, sus partes y su .rels; reasigna localSheetId, activeTab y firstSheet; borra los nombres locales y deja
// #REF! en los que apuntan a ella; sustituye por su valor las fórmulas que la usan; quita las tablas dinámicas, segmentaciones y escalas de tiempo que
// dependían de ella; borra calcChain, la miniatura y las partes que ya nadie referencia; y actualiza docProps/app.xml y los textos compartidos.
(function () {
  const X = XmlEditar;
  const R_OFICINA = 'officeDocument', MAX_PARTES = 200000;

  // ---- rutas y relaciones --------------------------------------------------------------------------------
  const rutaRels = (parte) => { const b = parte.lastIndexOf('/'); return parte.slice(0, b + 1) + '_rels/' + parte.slice(b + 1) + '.rels'; };
  function resolver(parte, destino) {
    let d = destino;
    try { d = decodeURI(destino); } catch (e) { /* tal cual */ }
    d = d.replace(/\\/g, '/');
    if (d.startsWith('/')) return d.slice(1);
    const out = [];
    (parte.slice(0, parte.lastIndexOf('/') + 1) + d).split('/').forEach((s) => { if (s === '..') out.pop(); else if (s !== '.' && s !== '') out.push(s); });
    return out.join('/');
  }
  const textoDe = async (zip, ruta) => { const f = ruta ? zip.file(ruta) : null; return f ? f.async('string') : null; };
  const sufijo = (tipo) => tipo.slice(tipo.lastIndexOf('/') + 1);
  async function leerRels(zip, parte) {
    const t = await textoDe(zip, rutaRels(parte));
    if (t === null) return [];
    const salida = [];
    XmlLigero.recorrer(t, { abrir(n, a) {
      if (n === 'Relationship' && salida.length < MAX_PARTES) {
        const externa = a.TargetMode === 'External';
        salida.push({ id: a.Id || '', tipo: sufijo(a.Type || ''), externa, destino: externa ? (a.Target || '') : resolver(parte, a.Target || '') });
      }
    } });
    return salida;
  }
  async function quitarRels(zip, parte, ids) {
    if (!ids.size) return;
    const ruta = rutaRels(parte), t = await textoDe(zip, ruta);
    if (t === null) return;
    const ed = [];
    for (const n of X.analizar(t, { profundidad: 2 }).hijos) if (n.local === 'Relationship' && ids.has(n.attrs.Id)) ed.push({ ini: n.ini, fin: n.fin, texto: '' });
    if (ed.length) zip.file(ruta, X.aplicar(t, ed));
  }

  // Partes a las que se llega por relaciones desde la raíz; el resto se borra, con su Override. Si un .rels no se puede leer no se borra nada.
  async function limpiarHuerfanas(zip) {
    const alc = new Set(), pend = [''];
    while (pend.length) {
      const parte = pend.pop();
      const r = rutaRels(parte);
      if (!zip.file(r)) continue;
      alc.add(r);
      let rels;
      try { rels = await leerRels(zip, parte); } catch (e) { return []; }
      for (const x of rels) if (!x.externa && !alc.has(x.destino)) { alc.add(x.destino); pend.push(x.destino); }
    }
    const borradas = [];
    zip.forEach((ruta, ent) => { if (!ent.dir && ruta !== '[Content_Types].xml' && !alc.has(ruta)) borradas.push(ruta); });
    borradas.forEach((r) => zip.remove(r));
    const ct = await textoDe(zip, '[Content_Types].xml');
    if (ct !== null && borradas.length) {
      const idx = new Set(borradas.map((r) => '/' + r.toLowerCase()));
      const ed = [];
      for (const n of X.analizar(ct, { profundidad: 2 }).hijos) if (n.local === 'Override' && idx.has((n.attrs.PartName || '').toLowerCase())) ed.push({ ini: n.ini, fin: n.fin, texto: '' });
      if (ed.length) zip.file('[Content_Types].xml', X.aplicar(ct, ed));
    }
    return borradas;
  }

  // ---- referencias a hojas ----------------------------------------------------------------------------------
  const escRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  function detector(nombres) { // ¿esta fórmula/ruta usa alguna de esas hojas? (un solo patrón: no una pasada por nombre)
    if (!nombres.length) return () => false;
    const alt = nombres.map(escRe).join('|'), altQ = nombres.map((n) => escRe(n.replace(/'/g, "''"))).join('|');
    const re = new RegExp("'(?:" + altQ + ")'!|(?:^|[^\\w.'])(?:" + alt + ')!');
    return (t) => re.test(t);
  }

  // ---- análisis de lo que depende de las hojas que quedan ------------------------------------------------------
  // Tablas dinámicas, segmentaciones y escalas de tiempo: se conservan solo si todo de lo que dependen sigue en el libro.
  async function dependencias(zip, rutaLibro, relsLibro, partesHojas, sheetIdPorHoja) {
    const cachesVivas = new Set(), pivotesVivos = new Set(), tablasVivas = new Set();
    for (const [nombre, parte] of partesHojas) {
      const rels = await leerRels(zip, parte);
      for (const r of rels) {
        if (r.externa) continue;
        if (r.tipo === 'pivotTable') {
          const t = await textoDe(zip, r.destino);
          if (t !== null) { try { pivotesVivos.add(sheetIdPorHoja.get(nombre) + '|' + X.attr(X.analizar(t, { profundidad: 1 }), 'name')); } catch (e) { /* ilegible */ } }
          for (const rc of await leerRels(zip, r.destino)) if (rc.tipo === 'pivotCacheDefinition') cachesVivas.add(rc.destino);
        } else if (r.tipo === 'table') {
          const t = await textoDe(zip, r.destino);
          if (t !== null) { try { tablasVivas.add(X.attr(X.analizar(t, { profundidad: 1 }), 'id')); } catch (e) { /* ilegible */ } }
        }
      }
    }
    return { cachesVivas, pivotesVivos, tablasVivas };
  }

  // ---- docProps/app.xml: nombres de hojas (y rangos con nombre) ----------------------------------------------
  async function actualizarApp(zip, orig, conservar, quita) {
    const t = await textoDe(zip, 'docProps/app.xml');
    if (t === null) return;
    let raiz;
    try { raiz = X.analizar(t, { profundidad: 6 }); } catch (e) { return; }
    const pares = X.hijo(raiz, 'HeadingPairs'), titulos = X.hijo(raiz, 'TitlesOfParts');
    const vp = pares && X.hijo(pares, 'vector'), vt = titulos && X.hijo(titulos, 'vector');
    if (!vp || !vt) return;
    const vars = X.hijos(vp, 'variant'), entradas = vt.hijos.filter((n) => n.local === 'lpstr');
    if (vars.length < 2 || vars.length % 2) return;
    const pre = vt.hijos.length ? vt.hijos[0].nombre.replace(/lpstr$/, '') : 'vt:';
    const grupos = [];
    let pos = 0;
    for (let i = 0; i + 1 < vars.length; i += 2) {
      const cant = X.hijo(vars[i + 1], 'i4');
      if (!cant) return;
      const n = parseInt(X.interior(t, cant), 10) || 0;
      grupos.push({ nombre: vars[i], textos: entradas.slice(pos, pos + n).map((e) => X.decodificar(X.interior(t, e))) });
      pos += n;
    }
    const sinHojas = (txt) => !quita(txt);
    const nombresHojas = new Set(orig);
    // un grupo es de hojas si todos sus textos son nombres de hojas del libro (hojas de cálculo, hojas de gráfico…): se reordena; los demás (rangos con nombre) se filtran
    const nuevos = grupos.map((g) => (g.textos.length && g.textos.every((x) => nombresHojas.has(x)) ? conservar.filter((n) => g.textos.includes(n)) : g.textos.filter(sinHojas)));
    let ph = '', pt = '';
    grupos.forEach((g, i) => {
      if (!nuevos[i].length) return;
      ph += t.slice(g.nombre.ini, g.nombre.fin) + '<' + pre + 'variant><' + pre + 'i4>' + nuevos[i].length + '</' + pre + 'i4></' + pre + 'variant>';
      nuevos[i].forEach((x) => { pt += '<' + pre + 'lpstr>' + X.esc(x) + '</' + pre + 'lpstr>'; });
    });
    const nv = nuevos.reduce((a, g) => a + (g.length ? 2 : 0), 0), ne = nuevos.reduce((a, g) => a + g.length, 0);
    zip.file('docProps/app.xml', X.aplicar(t, [
      { ini: vp.ini, fin: vp.finApertura, texto: X.ponerAtributo(t.slice(vp.ini, vp.finApertura), 'size', String(nv)) },
      { ini: vp.finApertura, fin: vp.iniCierre, texto: ph },
      { ini: vt.ini, fin: vt.finApertura, texto: X.ponerAtributo(t.slice(vt.ini, vt.finApertura), 'size', String(ne)) },
      { ini: vt.finApertura, fin: vt.iniCierre, texto: pt },
    ]));
  }

  // ---- textos compartidos: solo los que usan las hojas que quedan ---------------------------------------------
  const RE_CELDA = /<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g;
  async function compactarTextos(zip, rutaLibro, rutasHojas) {
    const rel = (await leerRels(zip, rutaLibro)).find((r) => r.tipo === 'sharedStrings');
    const t = rel && await textoDe(zip, rel.destino);
    if (!t) return;
    const raiz = X.analizar(t, { profundidad: 2 });
    const si = raiz.hijos.filter((n) => n.local === 'si');
    const xmls = new Map(), nuevo = new Map();
    let total = 0;
    for (const ruta of rutasHojas) {
      const xml = await textoDe(zip, ruta);
      if (xml === null) continue;
      xmls.set(ruta, xml);
      xml.replace(RE_CELDA, (celda) => {
        const ini = celda.match(/^<c\b[^>]*>/)[0];
        if (!/\st="s"/.test(ini)) return celda;
        const v = celda.match(/<v>(\d+)<\/v>/);
        if (v) { total++; if (!nuevo.has(+v[1])) nuevo.set(+v[1], nuevo.size); }
        return celda;
      });
    }
    if (nuevo.size === si.length) return;
    for (const [ruta, xml] of xmls) {
      zip.file(ruta, xml.replace(RE_CELDA, (celda) => {
        const ini = celda.match(/^<c\b[^>]*>/)[0];
        if (!/\st="s"/.test(ini)) return celda;
        return celda.replace(/<v>(\d+)<\/v>/, (_, n) => '<v>' + (nuevo.has(+n) ? nuevo.get(+n) : n) + '</v>');
      }));
    }
    const orden = [...nuevo.entries()].sort((a, b) => a[1] - b[1]).map(([viejo]) => si[viejo]).filter(Boolean);
    let ab = X.ponerAtributo(t.slice(raiz.ini, raiz.finApertura), 'uniqueCount', String(orden.length));
    ab = X.ponerAtributo(ab, 'count', String(total));
    zip.file(rel.destino, t.slice(0, raiz.ini) + ab + orden.map((n) => t.slice(n.ini, n.fin)).join('') + t.slice(raiz.iniCierre));
  }

  // ---- una hoja que se queda: fórmulas y objetos que miran a hojas quitadas -------------------------------------
  function arreglarHoja(xml, usa, quitadas, activa, esActiva) {
    const menciona = quitadas.some((n) => xml.includes(n) || xml.includes(X.esc(n)));
    if (menciona) {
      const rotas = new Set();
      xml.replace(/<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/g, (todo, at, f) => {
        const si = at.match(/\bsi="(\d+)"/);
        if (f !== undefined && si && /\bt="shared"/.test(at) && usa(X.decodificar(f))) rotas.add(si[1]);
        return todo;
      });
      xml = xml.replace(RE_CELDA, (celda) => {
        let rota = false;
        const sin = celda.replace(/<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/, (todo, at, f) => {
          const si = at.match(/\bsi="(\d+)"/);
          rota = (f !== undefined && usa(X.decodificar(f))) || Boolean(si && rotas.has(si[1]));
          return rota ? '' : todo;
        });
        if (!rota) return celda;
        if (/^<c\b[^>]*\bt="str"/.test(sin)) return sin.replace(/^(<c\b[^>]*\bt=")str"/, '$1inlineStr"').replace(/<v>([\s\S]*?)<\/v>/, '<is><t xml:space="preserve">$1</t></is>');
        return sin;
      });
      // formatos condicionales y validaciones que apuntan a una hoja quitada: Excel los tomaría por vínculos a otro libro
      let quitoVal = 0;
      xml = xml.replace(/<(?:\w+:)?conditionalFormatting\b[\s\S]*?<\/(?:\w+:)?conditionalFormatting>/g, (b) => (usa(X.decodificar(b.replace(/<[^>]*>/g, ' '))) ? '' : b));
      xml = xml.replace(/<dataValidation\b[\s\S]*?<\/dataValidation>/g, (b) => (usa(X.decodificar(b.replace(/<[^>]*>/g, ' '))) ? (quitoVal++, '') : b));
      if (quitoVal) {
        xml = xml.replace(/<dataValidations\b([^>]*)>([\s\S]*?)<\/dataValidations>/, (todo, at, dentro) => {
          const n = (dentro.match(/<dataValidation\b/g) || []).length;
          return n ? '<dataValidations' + X.ponerAtributo('<x' + at + '>', 'count', String(n)).slice(2) + dentro + '</dataValidations>' : '';
        });
      }
      xml = xml.replace(/<(?:\w+:)?(?:conditionalFormatting|dataValidation)\b[^>]*>(?:(?!<\/)[\s\S])*?<(?:\w+:)?f>[^<]*<\/(?:\w+:)?f>[\s\S]*?<\/(?:\w+:)?(?:conditionalFormatting|dataValidation)>/g, (b) => (usa(X.decodificar(b.replace(/<[^>]*>/g, ' '))) ? '' : b));
    }
    // solo la pestaña activa queda seleccionada
    let primera = true;
    xml = xml.replace(/<sheetView\b[^>]*>/g, (et) => {
      let l = et.replace(/\s+tabSelected="[^"]*"/, '');
      if (primera && esActiva) l = l.replace(/^<sheetView\b/, '<sheetView tabSelected="1"');
      primera = false;
      return l;
    });
    return xml;
  }

  // ---- la operación ---------------------------------------------------------------------------------------------
  async function reordenarHojas(zip, nombresEnOrden) {
    const raizRels = await leerRels(zip, '');
    const principal = raizRels.find((r) => r.tipo === R_OFICINA);
    const rutaLibro = principal ? principal.destino : 'xl/workbook.xml';
    const wb = await textoDe(zip, rutaLibro);
    if (wb === null) throw new Error(rutaLibro);
    const raiz = X.analizar(wb, { profundidad: 7 });
    const listaN = X.hijo(raiz, 'sheets');
    const hojas = X.hijos(listaN, 'sheet').map((n) => ({ n, nombre: n.attrs.name, id: X.attr(n, 'id'), sheetId: n.attrs.sheetId, estado: n.attrs.state || '' }));
    const porNombre = new Map(hojas.map((h) => [h.nombre, h]));
    const conservar = nombresEnOrden.filter((n, i) => porNombre.has(n) && nombresEnOrden.indexOf(n) === i);
    if (!conservar.length) throw new Error('sin hojas');
    const orig = hojas.map((h) => h.nombre);
    const quitadas = orig.filter((n) => !conservar.includes(n));
    const usa = detector(quitadas);
    const relsLibro = await leerRels(zip, rutaLibro);
    const rutaPorId = new Map(relsLibro.map((r) => [r.id, r]));
    const parteDe = (h) => { const r = rutaPorId.get(h.id); return r && !r.externa ? r.destino : null; };
    const partesVivas = new Map(conservar.map((n) => [n, parteDe(porNombre.get(n))]).filter(([, p]) => p));
    const sheetIdPorHoja = new Map(hojas.map((h) => [h.nombre, h.sheetId]));

    const ed = [], quitarRelLibro = new Set();
    const indice = (n) => conservar.indexOf(n);

    // hoja activa y visibilidad
    const vista = X.hijo(X.hijo(raiz, 'bookViews'), 'workbookView');
    const activaOrig = orig[parseInt(vista && vista.attrs.activeTab, 10) || 0];
    const visible = (n) => !/hidden/i.test(porNombre.get(n).estado);
    const activa = conservar.includes(activaOrig) && visible(activaOrig) ? activaOrig : (conservar.find(visible) || conservar[0]);
    const sinVisible = !conservar.some(visible);   // Excel exige una hoja visible
    ed.push({ ini: listaN.finApertura, fin: listaN.iniCierre, texto: conservar.map((n, k) => {
      const h = porNombre.get(n).n, ab = wb.slice(h.ini, h.finApertura);
      return (sinVisible && k === 0 ? X.ponerAtributo(ab, 'state', null) : ab) + wb.slice(h.finApertura, h.fin);
    }).join('') });
    if (vista) {
      let et = wb.slice(vista.ini, vista.finApertura);
      et = X.ponerAtributo(et, 'activeTab', String(indice(activa)));
      et = X.ponerAtributo(et, 'firstSheet', null);
      ed.push({ ini: vista.ini, fin: vista.finApertura, texto: et });
    }

    // dependencias (solo si se quita algo)
    let cachesVivas = new Set();
    const nombresCacheQuitados = new Set(), segmentacionesQuitadas = new Set();
    if (quitadas.length) {
      const dep = await dependencias(zip, rutaLibro, relsLibro, partesVivas, sheetIdPorHoja);
      cachesVivas = dep.cachesVivas;
      const pc = X.hijo(raiz, 'pivotCaches');
      for (const c of X.hijos(pc, 'pivotCache')) {
        const r = rutaPorId.get(X.attr(c, 'id'));
        if (!r || !cachesVivas.has(r.destino)) { ed.push({ ini: c.ini, fin: c.fin, texto: '' }); quitarRelLibro.add(X.attr(c, 'id')); }
      }
      // segmentaciones y escalas de tiempo (extLst del libro)
      const ext = X.hijo(raiz, 'extLst');
      for (const e of X.hijos(ext, 'ext')) {
        for (const lista of e.hijos) {
          if (lista.local !== 'slicerCaches' && lista.local !== 'timelineCacheRefs') continue;
          const refs = lista.hijos;
          for (const ref of refs) {
            const r = rutaPorId.get(X.attr(ref, 'id'));
            const t = r ? await textoDe(zip, r.destino) : null;
            let vive = false, nombre = null;
            if (t !== null) {
              try {
                const cr = X.analizar(t, { profundidad: 8 });
                nombre = cr.attrs.name || null;
                const pivs = [], tabs = [];
                (function rec(n) { for (const h of n.hijos) { if (h.local === 'pivotTable' && h.attrs.tabId !== undefined) pivs.push(h); else if (h.local === 'tableSlicerCache') tabs.push(h); rec(h); } })(cr);
                const idPorSheetId = new Map(hojas.map((h) => [h.sheetId, h.nombre]));
                vive = (pivs.length > 0 || tabs.length > 0)
                  && pivs.every((p) => { const hn = idPorSheetId.get(p.attrs.tabId); return hn && conservar.includes(hn) && dep.pivotesVivos.has(p.attrs.tabId + '|' + p.attrs.name); })
                  && tabs.every((p) => dep.tablasVivas.has(p.attrs.tableId));
              } catch (e) { vive = false; }
            }
            if (!vive) { ed.push({ ini: ref.ini, fin: ref.fin, texto: '' }); if (nombre) { nombresCacheQuitados.add(nombre); segmentacionesQuitadas.add(nombre); } if (X.attr(ref, 'id')) quitarRelLibro.add(X.attr(ref, 'id')); }
          }
        }
      }
    }

    // nombres definidos
    const defs = X.hijo(raiz, 'definedNames');
    const indiceViejo = new Map(orig.map((n, i) => [i, n]));
    if (defs) {
      let quedan = 0;
      for (const d of X.hijos(defs, 'definedName')) {
        const loc = d.attrs.localSheetId;
        const valor = X.decodificar(X.interior(wb, d));
        if (nombresCacheQuitados.has(d.attrs.name)) { ed.push({ ini: d.ini, fin: d.fin, texto: '' }); continue; }
        if (loc !== undefined) {
          const hn = indiceViejo.get(parseInt(loc, 10));
          if (!hn || !conservar.includes(hn)) { ed.push({ ini: d.ini, fin: d.fin, texto: '' }); continue; }
          ed.push({ ini: d.ini, fin: d.finApertura, texto: X.ponerAtributo(wb.slice(d.ini, d.finApertura), 'localSheetId', String(indice(hn))) });
          if (usa(valor) && d.iniCierre >= 0) ed.push({ ini: d.finApertura, fin: d.iniCierre, texto: '#REF!' });
          quedan++;
        } else if (usa(valor)) { ed.push({ ini: d.finApertura, fin: d.iniCierre, texto: '#REF!' }); quedan++; } else quedan++;
      }
      if (!quedan) ed.push({ ini: defs.ini, fin: defs.fin, texto: '' });
    }
    // vistas personalizadas del libro: guardan hojas por posición o identificador
    if (quitadas.length) { const v = X.hijo(raiz, 'customWorkbookViews'); if (v) ed.push({ ini: v.ini, fin: v.fin, texto: '' }); }

    // relaciones del libro: hojas quitadas, calcChain, miniatura
    for (const n of quitadas) quitarRelLibro.add(porNombre.get(n).id);
    const calc = relsLibro.find((r) => r.tipo === 'calcChain');
    if (calc && quitadas.length) quitarRelLibro.add(calc.id);

    // Si una ext quedó sin segmentaciones, se quita entera (Excel no admite la lista vacía).
    const fusionadas = [];
    {
      const ext = X.hijo(raiz, 'extLst');
      for (const e of X.hijos(ext, 'ext')) {
        for (const lista of e.hijos) {
          if ((lista.local === 'slicerCaches' || lista.local === 'timelineCacheRefs') && lista.hijos.length && lista.hijos.every((r) => ed.some((x) => x.ini === r.ini && x.fin === r.fin && x.texto === ''))) {
            fusionadas.push({ ini: e.ini, fin: e.fin });
          }
        }
      }
    }
    for (const f of fusionadas) { for (let k = ed.length - 1; k >= 0; k--) if (ed[k].ini >= f.ini && ed[k].fin <= f.fin) ed.splice(k, 1); ed.push({ ini: f.ini, fin: f.fin, texto: '' }); }
    {
      const pc = X.hijo(raiz, 'pivotCaches');
      if (pc && pc.hijos.length && pc.hijos.every((c) => ed.some((x) => x.ini === c.ini && x.fin === c.fin && x.texto === ''))) {
        for (let k = ed.length - 1; k >= 0; k--) if (ed[k].ini >= pc.ini && ed[k].fin <= pc.fin) ed.splice(k, 1);
        ed.push({ ini: pc.ini, fin: pc.fin, texto: '' });
      }
    }
    zip.file(rutaLibro, X.aplicar(wb, ed));
    await quitarRels(zip, rutaLibro, quitarRelLibro);

    // hojas que se quedan
    const rutasHojas = [];
    for (const n of conservar) {
      const parte = partesVivas.get(n);
      const rel = parte && relsLibro.find((r) => r.destino === parte);
      if (!parte || !rel || rel.tipo !== 'worksheet') { if (parte && quitadas.length && segmentacionesQuitadas.size) await quitarSegmentaciones(zip, parte, segmentacionesQuitadas); continue; }
      const xml = await textoDe(zip, parte);
      if (xml === null) continue;
      rutasHojas.push(parte);
      const nuevo = arreglarHoja(xml, usa, quitadas, activa, n === activa);
      if (nuevo !== xml) zip.file(parte, nuevo);
      if (quitadas.length && segmentacionesQuitadas.size) await quitarSegmentaciones(zip, parte, segmentacionesQuitadas);
    }
    // las vistas personalizadas de cada hoja dependen de las del libro
    if (quitadas.length) for (const p of rutasHojas) { const x = await textoDe(zip, p); const y = x.replace(/<customSheetViews\b[\s\S]*?<\/customSheetViews>/, ''); if (y !== x) zip.file(p, y); }

    await actualizarApp(zip, orig, conservar, (t) => quitadas.includes(t) || usa(t));
    // la miniatura es de la primera hoja del original
    const raizR = await leerRels(zip, '');
    await quitarRels(zip, '', new Set(raizR.filter((r) => r.tipo === 'thumbnail').map((r) => r.id)));
    const borradas = await limpiarHuerfanas(zip);
    if (quitadas.length) await compactarTextos(zip, rutaLibro, rutasHojas);
    return { quitadas, conservadas: conservar, partesBorradas: borradas };
  }

  // Segmentaciones de esta hoja cuya caché se quitó: se borran de la lista, de su parte y del dibujo.
  async function quitarSegmentaciones(zip, parteHoja, cachesQuitadas) {
    const rels = await leerRels(zip, parteHoja);
    const nombres = new Set();
    const idsRel = new Set();
    for (const r of rels) {
      if (r.externa || r.tipo !== 'slicer') continue;
      const t = await textoDe(zip, r.destino);
      if (t === null) continue;
      const raiz = X.analizar(t, { profundidad: 3 });
      const ed = [];
      let quedan = 0;
      for (const s of X.hijos(raiz, 'slicer')) { if (cachesQuitadas.has(s.attrs.cache)) { ed.push({ ini: s.ini, fin: s.fin, texto: '' }); nombres.add(s.attrs.name); } else quedan++; }
      if (ed.length) zip.file(r.destino, X.aplicar(t, ed));
      if (!quedan) idsRel.add(r.id);
    }
    if (!nombres.size) return;
    // el dibujo: se quita el objeto entero (con su mc:AlternateContent si lo tiene)
    for (const r of rels) {
      if (r.externa || r.tipo !== 'drawing') continue;
      const t = await textoDe(zip, r.destino);
      if (t === null) continue;
      const raiz = X.analizar(t, { profundidad: 40 });
      const ed = [], hechos = new Set();
      (function rec(n) { for (const h of n.hijos) { if (h.local === 'slicer' && nombres.has(h.attrs.name)) { let a = h; while (a.padre && a.padre !== raiz) a = a.padre; if (!hechos.has(a)) { hechos.add(a); ed.push({ ini: a.ini, fin: a.fin, texto: '' }); } } rec(h); } })(raiz);
      if (ed.length) zip.file(r.destino, X.aplicar(t, ed));
    }
    // la lista de segmentaciones de la hoja (extLst) y sus relaciones
    if (idsRel.size) {
      const xml = await textoDe(zip, parteHoja);
      if (xml !== null) {
        const y = xml.replace(/<ext\b[^>]*>(?:(?!<\/ext>)[\s\S])*?<(?:\w+:)?slicerList\b[\s\S]*?<\/ext>/g, '');
        if (y !== xml) zip.file(parteHoja, y.replace(/<extLst>\s*<\/extLst>/, ''));
      }
      await quitarRels(zip, parteHoja, idsRel);
    }
  }

  globalThis.XlsxPaquete = { reordenarHojas, leerRels, resolver, limpiarHuerfanas };
})();

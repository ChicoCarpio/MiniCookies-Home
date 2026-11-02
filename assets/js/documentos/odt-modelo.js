/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Modelo de un .odt para el editor de Documentos (Fase 6b, etapa 6b.1). Sin DOM: corre igual en Node.
//
// Lee content.xml (y los encabezados y pies de styles.xml) a los MISMOS bloques que DocxModelo (docx-modelo.js): párrafos y tablas con el formato
// efectivo ya resuelto y con la POSICIÓN de cada uno en el texto del XML. El maquetador, el pintor y la capa accesible no saben que el archivo es ODT.
// Lo que el modelo no entiende (notas, comentarios, cambios con control de cambios, índices, marcos, campos) se conserva sin tocar.
//
// Diferencias con Word que este modelo resuelve al leer (diseño en docs/fase-6b.md):
//   · Los espacios se colapsan (ODF 1.2 §6.1.2): la pieza de texto lleva `s` (colapsado) y la posición del texto crudo (`ini`, `fin`); `text:s` es otra pieza.
//   · Las listas envuelven a los párrafos (`text:list`): se aplanan; cada párrafo lleva `lista` y, si toca, su `etiqueta` (viñeta o número).
//   · Los encabezados y pies viven en styles.xml: `secciones[].referencias` trae ids sintéticos `odt:<maestra>:encabezado:default` y `rels.get(id).destino`
//     = `styles.xml#<maestra>/encabezado/default`; `parte(destino)` lee los bloques de ahí.
//   · Una sección empieza en un párrafo cuyo estilo trae `master-page-name` (`abreSeccion`); `seccion` es siempre null.
//   · Sin `text:p`/`text:h` en un índice o sección protegida: el bloque lleva `protegido: true`.
//
// Requiere zip-propio.js, xml-editar.js, paquete-editable.js, docx-base.js, docx-estilos.js y odt-estilos.js.
//
//   const modelo = await OdtModelo.abrir(bytes)      (misma forma que DocxModelo.abrir: bloques, secciones, notas, rels, parrafos(), parte(), releer()…)
(function () {
  'use strict';
  const X = () => globalThis.XmlEditar;
  const E = () => globalThis.OdtEstilos;
  const A = (n) => E().A(n);
  const MAX_PROFUNDIDAD = 8, MAX_INLINE = 60, MAX_CARACTERES = 50 * 1024 * 1024, MAX_COLUMNAS = 128, MAX_PT = 1584;
  const acotar = (x, min, max) => (Number.isFinite(x) ? Math.min(Math.max(x, min), max) : min);
  const hijos = (n, local) => (n ? n.hijos.filter((h) => h.local === local) : []);
  const hijo = (n, local) => (n ? n.hijos.find((h) => h.local === local) || null : null);
  const INDICES = new Set(['table-of-content', 'alphabetical-index', 'user-index', 'illustration-index', 'table-index', 'object-index', 'bibliography']);
  const FORMATOS = { '1': 'decimal', a: 'lowerLetter', A: 'upperLetter', i: 'lowerRoman', I: 'upperRoman' };
  const textoDe = (xml, n) => X().decodificar(X().interior(xml, n));

  // ───────── Texto en línea ─────────
  // Los blancos de un trozo de texto crudo: se colapsan a un espacio, y el primero de un párrafo (o tras otro espacio) se descarta.
  function colapsar(crudo, ctx) {
    let s = '';
    for (const ch of crudo) {
      if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') { if (!ctx.espacio) { s += ' '; ctx.espacio = true; } }
      else { s += ch; ctx.espacio = false; }
    }
    return s;
  }
  // El texto de un hueco entre elementos: sin comentarios ni instrucciones; el CDATA va tal cual.
  function huecoATexto(crudo) {
    if (crudo.indexOf('<') < 0) return X().decodificar(crudo);
    let sal = '', i = 0;
    while (i < crudo.length) {
      const lt = crudo.indexOf('<', i);
      if (lt < 0) { sal += X().decodificar(crudo.slice(i)); break; }
      sal += X().decodificar(crudo.slice(i, lt));
      if (crudo.startsWith('<![CDATA[', lt)) { const f = crudo.indexOf(']]>', lt); sal += crudo.slice(lt + 9, f < 0 ? crudo.length : f); i = f < 0 ? crudo.length : f + 3; }
      else if (crudo.startsWith('<!--', lt)) { const f = crudo.indexOf('-->', lt); i = f < 0 ? crudo.length : f + 3; }
      else { const f = crudo.indexOf('>', lt); i = f < 0 ? crudo.length : f + 1; }
    }
    return sal;
  }
  const hrefSeguro = (h) => {
    if (!h) return '';
    if (globalThis.DocxBase && globalThis.DocxBase.hrefSeguro) return globalThis.DocxBase.hrefSeguro(h);
    return /^(https?:|mailto:|#)/i.test(h) ? h : '';
  };

  function leerMarco(c, e, r) {
    const a = A(e);
    const ancho = acotar(E().pt(a.width) || 0, 0, MAX_PT), alto = acotar(E().pt(a.height) || 0, 0, MAX_PT);
    const dentro = e.hijos.find((x) => ['image', 'text-box', 'object', 'object-ole', 'plugin', 'applet', 'floating-frame'].includes(x.local));
    const desc = hijo(e, 'desc') || hijo(e, 'title');
    const enLinea = a['anchor-type'] === 'as-char';
    const base = { r, ini: e.ini, fin: e.fin, ei: e.ini, ef: e.fin, ancho, alto, enLinea, descr: desc ? textoDe(c.xml, desc) : '' };
    if (!enLinea) {
      base.ancla = {
        h: { rel: 'column', align: null, offset: E().pt(a.x) || 0 }, v: { rel: 'paragraph', align: null, offset: E().pt(a.y) || 0 },
        envoltura: 'wrapTopAndBottom', detras: false, dist: { t: 0, b: 0, l: 0, r: 0 },
      };
      c.cosas.add('imagenes-flotantes');
    }
    if (dentro && dentro.local === 'image') {
      const href = A(dentro).href || '';
      const interno = href && !/^[a-z][a-z0-9+.-]*:/i.test(href) && !/\.\.|^\//.test(href) ? href.replace(/^\.\//, '') : null;
      if (interno) return { t: 'imagen', ...base, destino: interno };
      return { t: 'dibujo', ...base };
    }
    if (dentro && dentro.local === 'text-box') { c.cosas.add('cuadros'); return { t: 'cuadro', ...base }; }
    c.cosas.add(dentro ? 'objetos' : 'dibujos');
    return { t: 'dibujo', ...base };
  }

  function leerInline(c, nodo, r, salida, ctx, prof = 0, enlace = null, opaco = null) {
    if (prof > MAX_INLINE) return;
    const poner = (z) => { if (enlace) z.enlace = enlace; if (opaco) z.opaco = opaco; salida.push(z); };
    const hueco = (a, b) => {
      if (b <= a) return;
      const crudo = c.xml.slice(a, b);
      const s = colapsar(huecoATexto(crudo), ctx);
      if (!s) return;
      if ((c.caracteres += s.length) > MAX_CARACTERES) throw new Error('El documento es demasiado grande para abrirlo aquí.');
      poner({ t: 'texto', s, r, ini: a, fin: b, ei: a, ef: b, ti: a, tf: b, hueco: true, rara: crudo.indexOf('<') >= 0 });
    };
    let pos = nodo.finApertura;
    for (const e of nodo.hijos) {
      hueco(pos, e.ini);
      pos = e.fin;
      const a = A(e);
      const dato = { r, ini: e.ini, fin: e.fin, ei: e.ini, ef: e.fin };
      switch (e.local) {
        case 'span': leerInline(c, e, c.est.sumarR(r, c.est.resolverCaracter(a['style-name'])), salida, ctx, prof + 1, enlace, opaco); break;
        case 'a': {
          const href = a.href ? (a.href.startsWith('#') ? a.href : hrefSeguro(a.href)) : '';
          leerInline(c, e, c.est.sumarR(r, c.est.resolverCaracter(a['style-name'])), salida, ctx, prof + 1, href || enlace, opaco);
          break;
        }
        case 's': {
          const n = acotar(parseInt(a.c, 10) || 1, 1, 100000);
          if ((c.caracteres += n) > MAX_CARACTERES) throw new Error('El documento es demasiado grande para abrirlo aquí.');
          ctx.espacio = false;
          poner({ t: 'texto', s: ' '.repeat(n), ...dato, espacio: true });
          break;
        }
        case 'tab': ctx.espacio = false; poner({ t: 'tab', ...dato }); break;
        case 'line-break': ctx.espacio = false; poner({ t: 'br', ...dato }); break;
        case 'page-number': poner({ t: 'campo', nombre: 'PAGE', ...dato }); ctx.espacio = false; break;
        case 'page-count': poner({ t: 'campo', nombre: 'NUMPAGES', ...dato }); ctx.espacio = false; break;
        case 'bookmark': case 'bookmark-start': if (a.name) poner({ t: 'ancla', nombre: a.name, ini: e.ini, fin: e.fin }); break;
        case 'note': {
          const fin = a['note-class'] === 'endnote';
          const clave = (fin ? 'e' : 'f') + (a.id || c.notasPend.size + 1);
          if (!c.notasPend.has(clave)) { c.notasPend.set(clave, { cuerpo: hijo(e, 'note-body'), n: c.notasPend.size + 1 }); c.cosas.add(fin ? 'notas-finales' : 'notas-pie'); }
          poner({ t: 'texto', s: String(c.notasPend.get(clave).n), ...dato, r: { ...r, alineacionVertical: 'superscript' }, nota: clave });
          ctx.espacio = false;
          break;
        }
        case 'frame': { const z = leerMarco(c, e, r); poner(z); ctx.espacio = false; break; }
        case 'ruby': leerInline(c, hijo(e, 'ruby-base') || { hijos: [], finApertura: e.finApertura }, r, salida, ctx, prof + 1, enlace, opaco); c.cosas.add('ruby'); break;
        case 'annotation': c.cosas.add('comentarios'); break;
        case 'annotation-end': case 'soft-page-break': case 'change': case 'change-start': case 'change-end': case 'tracked-changes': break;
        case 'meta': case 'meta-field': leerInline(c, e, r, salida, ctx, prof + 1, enlace, opaco); break;
        default:
          // un campo (fecha, título, variable…) u otro elemento en línea: se muestra su texto guardado y no se edita
          if (e.hijos.length || e.iniCierre >= 0) { c.cosas.add('campos'); leerInline(c, e, r, salida, ctx, prof + 1, enlace, opaco || e.local); }
          break;
      }
    }
    if (nodo.iniCierre >= 0) hueco(pos, nodo.iniCierre);
  }

  function asignarTexto(piezas) {
    let texto = '';
    for (const z of piezas) {
      z.desde = texto.length;
      if (z.t === 'texto') texto += z.s;
      else if (z.t === 'tab') texto += '\t';
      else if (z.t === 'br') texto += ' ';
      else if (z.t === 'pagina' || z.t === 'columna') texto += '\f';
      else if (z.t !== 'ancla') texto += '￼';
      z.hasta = texto.length;
    }
    return texto;
  }

  // ───────── Listas ─────────
  const viñeta = (ch) => (!ch ? '•' : /[-]/.test(ch) ? '•' : ch);
  function etiquetaDe(c, inst, nivelN, valorInicial) {
    const niveles = c.est.lista(inst.estilo) || [];
    const def = niveles[Math.min(nivelN, niveles.length - 1)] || { tipo: 'number', formato: '1', prefijo: '', sufijo: '.', inicio: 1, niveles: 1, sepa: 'tab', p: {}, r: {} };
    const real = niveles[nivelN] || def;
    const cont = inst.cont;
    if (valorInicial !== undefined) cont[nivelN] = valorInicial;
    else cont[nivelN] = cont[nivelN] === undefined ? (real.inicio || 1) : cont[nivelN] + 1;
    cont.length = nivelN + 1;
    if (real.tipo === 'bullet') return { texto: viñeta(real.caracter), fmt: 'bullet', sufijo: real.sepa, nivel: real };
    if (real.tipo === 'image') return { texto: '•', fmt: 'bullet', sufijo: real.sepa, nivel: real };
    const partes = [];
    for (let k = Math.max(0, nivelN - (real.niveles - 1)); k <= nivelN; k++) {
      const dk = niveles[k] || real, n = cont[k] === undefined ? (dk.inicio || 1) : cont[k];
      partes.push(dk.formato === '' ? '' : globalThis.DocxEstilos.formatear(n, FORMATOS[dk.formato] || 'decimal'));
    }
    const num = partes.filter((x) => x !== '').join('.');
    return { texto: real.formato === '' && !real.prefijo && !real.sufijo ? '' : (real.prefijo || '') + num + (real.sufijo || ''), fmt: FORMATOS[real.formato] || 'none', sufijo: real.sepa, nivel: real };
  }

  function leerLista(c, nodo, medio, prof, sal, padre, protegido) {
    if (prof > MAX_PROFUNDIDAD * 4) return;
    const a = A(nodo);
    const estilo = a['style-name'] || (padre && padre.inst.estilo) || null;
    // la lista de nivel superior es la unidad que se reescribe entera al cambiar su estructura (odt-listas.js): su posición, su etiqueta de apertura y si es de las que no se pueden regenerar
    const unidad = padre ? padre.unidad : { rango: { ini: nodo.ini, fin: nodo.fin }, tag: c.xml.slice(nodo.ini, nodo.finApertura), continua: !!(a['continue-list'] || a['continue-numbering'] === 'true'), estilo, compleja: false };
    const tagLista = c.xml.slice(nodo.ini, nodo.finApertura);
    let inst, nivelN;
    if (padre) { inst = padre.inst; nivelN = padre.nivel + 1; }
    else {
      nivelN = 0;
      const idRef = a['continue-list'], seguir = a['continue-numbering'] === 'true';
      inst = (idRef && c.listasPorId.get(idRef)) || (seguir && c.ultimaLista.get(estilo)) || { estilo, cont: [], id: 'l' + (++c.contadorListas) };
      if (a.id && !c.listasPorId.has(a.id)) c.listasPorId.set(a.id, inst);
      c.ultimaLista.set(estilo, inst);
    }
    for (const x of nodo.hijos) {
      const item = { tag: c.xml.slice(x.ini, x.finApertura), local: x.local, tagLista };
      if (x.local === 'list-header') { leerContenidoItem(c, x, medio, prof, sal, { inst, nivel: nivelN, estilo, unidad, item }, protegido, null); continue; }
      if (x.local !== 'list-item') { unidad.compleja = true; continue; }
      const inicio = A(x)['start-value'] !== undefined ? parseInt(A(x)['start-value'], 10) : undefined;
      const et = etiquetaDe(c, inst, nivelN, Number.isFinite(inicio) ? inicio : undefined);
      leerContenidoItem(c, x, medio, prof, sal, { inst, nivel: nivelN, estilo, unidad, item }, protegido, et);
    }
  }
  function leerContenidoItem(c, item, medio, prof, sal, info, protegido, etiqueta) {
    let primero = true;
    for (const x of item.hijos) {
      if (x.local === 'p' || x.local === 'h') { sal.push(parrafo(c, x, medio, { lista: { id: info.inst.id, nivel: info.nivel, estilo: info.estilo, cabecera: !etiqueta, item: info.item, unidad: info.unidad }, etiqueta: primero ? etiqueta : null, protegido })); primero = false; }
      else if (x.local === 'list') leerLista(c, x, medio, prof + 1, sal, info, protegido);
      else if (x.local === 'table') { info.unidad.compleja = true; sal.push(tabla(c, x, prof + 1, protegido)); }
      else info.unidad.compleja = true;
    }
  }

  // ───────── Párrafo ─────────
  function parrafo(c, p, medio, extra = {}) {
    const m = c.m, a = A(p);
    const nombre = a['style-name'] || null;
    const est = m.estilos.parrafo(nombre, medio);
    const res = c.est.resolverParrafo(nombre);
    let efectivo = est.p;
    if (p.local === 'h' && efectivo.nivelEsquema === undefined) efectivo = m.mezclar(efectivo, { nivelEsquema: acotar((parseInt(a['outline-level'], 10) || 1) - 1, 0, 9) });
    let etiqueta = null;
    if (extra.etiqueta && extra.etiqueta.texto !== undefined) {
      const e = extra.etiqueta, n = e.nivel;
      efectivo = m.mezclar(efectivo, { sangria: { izq: (n.p.sangria && n.p.sangria.izq) || 0, primera: (n.p.sangria && n.p.sangria.primera) || 0, colgante: (n.p.sangria && n.p.sangria.colgante) || 0 } });
      if (n.p.tabs) efectivo = m.mezclar(efectivo, { tabs: n.p.tabs });
      etiqueta = { texto: e.texto, fmt: e.fmt, sufijo: e.sufijo, nivel: n, rNivel: n.r };
    }
    if (c.ajustes.tabsRelativos && efectivo.tabs && !(extra.lista)) {
      const izq = (efectivo.sangria && efectivo.sangria.izq) || 0;
      if (izq) efectivo = { ...efectivo, tabs: efectivo.tabs.map((t) => ({ ...t, pos: t.pos + izq })) };
    }
    const baseR = est.r;
    const marca = baseR;
    if (etiqueta) { etiqueta.r = m.mezclar(marca, etiqueta.rNivel); delete etiqueta.rNivel; }
    const piezas = [];
    leerInline(c, p, baseR, piezas, { espacio: true });
    const texto = asignarTexto(piezas);
    const b = { tipo: 'p', medio: medio || null, ini: p.ini, fin: p.fin, estilo: nombre, nombreEtiqueta: p.local, numPedida: extra.lista ? { numId: extra.lista.id, nivel: String(extra.lista.nivel) } : null, p: efectivo, r: baseR, marca, etiqueta, piezas, texto, seccion: null };
    if (extra.lista) b.lista = extra.lista;
    if (extra.protegido) b.protegido = true;
    if (res.maestra) b.abreSeccion = res.maestra;
    if (res.ex && res.ex.saltoDespues) b.saltoDespues = true;
    return b;
  }

  // ───────── Tablas ─────────
  function recolectar(n, local, sal, prof = 0) {
    for (const x of n.hijos) {
      if (x.local === local) sal.push(x);
      else if (prof < 4 && ['table-header-rows', 'table-rows', 'table-row-group', 'table-columns', 'table-header-columns', 'table-column-group'].includes(x.local)) recolectar(x, local, sal, prof + 1);
    }
    return sal;
  }
  function tabla(c, t, prof, protegido) {
    const m = c.m, a = A(t), est = c.est.tabla(a['style-name']);
    // columnas
    const anchos = [], rel = [];
    for (const col of recolectar(t, 'table-column', [])) {
      const q = A(col), cs = c.est.columna(q['style-name']), n = acotar(parseInt(q['number-columns-repeated'], 10) || 1, 1, MAX_COLUMNAS);
      for (let k = 0; k < n && anchos.length < MAX_COLUMNAS; k++) { anchos.push(cs.ancho === undefined ? 0 : cs.ancho); rel.push(cs.relAncho || 0); }
    }
    const total = est.anchoTabla || (est.relAnchoTabla ? c.anchoUtil * est.relAnchoTabla / 100 : c.anchoUtil);
    const sinAncho = anchos.filter((w) => !w).length;
    if (sinAncho) {
      const conocido = anchos.reduce((s, w) => s + w, 0), sumaRel = rel.reduce((s, w, i) => s + (anchos[i] ? 0 : w), 0), resto = Math.max(total - conocido, 36 * sinAncho);
      anchos.forEach((w, i) => { if (!w) anchos[i] = sumaRel ? resto * rel[i] / sumaRel : resto / sinAncho; });
    }
    const grid = anchos.map((w) => Math.round(w * 20) / 20);
    // filas
    const filasXml = [];
    const recFilas = (n, enc, prof2 = 0) => {
      for (const x of n.hijos) {
        if (x.local === 'table-row') filasXml.push({ nodo: x, enc });
        else if (prof2 < 4 && x.local === 'table-header-rows') recFilas(x, true, prof2 + 1);
        else if (prof2 < 4 && (x.local === 'table-rows' || x.local === 'table-row-group')) recFilas(x, enc, prof2 + 1);
      }
    };
    recFilas(t, false);
    const cubiertasV = new Map();     // fila → columnas cubiertas por una celda combinada de una fila de arriba
    const infos = filasXml.map(({ nodo: f }, i) => {
      let col = 0;
      const lista = [];
      for (const cel of f.hijos) {
        if (cel.local !== 'table-cell' && cel.local !== 'covered-table-cell') continue;
        const q = A(cel), rep = acotar(parseInt(q['number-columns-repeated'], 10) || 1, 1, MAX_COLUMNAS);
        for (let k = 0; k < rep && col < MAX_COLUMNAS; k++) {
          if (cel.local === 'covered-table-cell') {
            if ((cubiertasV.get(i) || new Set()).has(col)) lista.push({ ini: cel.ini, fin: cel.fin, nodo: cel, col, span: 1, filas: 1, omitir: true, estilo: q['style-name'] });
            col++;
            continue;
          }
          const span = acotar(parseInt(q['number-columns-spanned'], 10) || 1, 1, MAX_COLUMNAS), filas = acotar(parseInt(q['number-rows-spanned'], 10) || 1, 1, 1000);
          for (let dr = 1; dr < filas; dr++) { const s = cubiertasV.get(i + dr) || new Set(); for (let dc = 0; dc < span; dc++) s.add(col + dc); cubiertasV.set(i + dr, s); }
          lista.push({ ini: cel.ini, fin: cel.fin, nodo: cel, col, span, filas, omitir: false, estilo: q['style-name'] });
          col += span;
        }
      }
      return lista;
    });
    const nCols = Math.max(infos.reduce((mx, f) => Math.max(mx, f.reduce((s, x) => Math.max(s, x.col + x.span), 0)), 0), grid.length, 1);
    while (grid.length < nCols) grid.push(grid.length ? grid[grid.length - 1] : 72);
    const filas = infos.map((fila, i) => {
      const fx = filasXml[i], fe = c.est.fila(A(fx.nodo)['style-name']);
      const tr = {};
      if (fe.alto !== undefined) { tr.alto = fe.alto; tr.regla = fe.reglaAlto || 'atLeast'; }
      if (fx.enc) tr.encabezado = true;
      return {
        ini: fx.nodo.ini, fin: fx.nodo.fin, tr,
        celdas: fila.map((info) => {
          const tc = c.est.celda(info.estilo).tc || {};
          const medio = { p: {}, r: {} };
          return { ini: info.ini, fin: info.fin, col: info.col, span: info.span, filas: info.filas, omitir: info.omitir, tc, medio, anchoDecl: undefined, bloques: info.omitir || prof + 1 > MAX_PROFUNDIDAD ? [] : bloquesDe(c, info.nodo, medio, prof + 1, protegido) };
        }),
      };
    });
    const tbl = {};
    if (est.anchoTabla) tbl.ancho = { tipo: 'dxa', valor: Math.round(est.anchoTabla * 20) };
    if (est.alineaTabla) tbl.jc = est.alineaTabla;
    if (est.sangriaTabla) tbl.sangria = { tipo: 'dxa', valor: Math.round(est.sangriaTabla * 20) };
    const r = { tipo: 'tabla', ini: t.ini, fin: t.fin, tbl, grid, nCols, anchoDecl: tbl.ancho, filas };
    if (protegido) r.protegido = true;
    return r;
  }

  // ───────── Bloques de un contenedor ─────────
  function bloquesDe(c, contenedor, medio, prof = 0, protegido = false, sal = []) {
    if (prof > MAX_PROFUNDIDAD * 4) return sal;
    for (const e of contenedor.hijos) {
      const antes = sal.length;
      switch (e.local) {
        case 'p': case 'h': sal.push(parrafo(c, e, medio, { protegido })); break;
        case 'list': leerLista(c, e, medio, prof, sal, null, protegido); break;
        case 'table': sal.push(tabla(c, e, prof, protegido)); break;
        case 'section': {
          const a = A(e);
          const prot = a.protected === 'true' || a['protected'] === '1';
          if (prot) c.cosas.add('secciones-protegidas');
          const cols = c.est.seccion(a['style-name']).columnas || null;
          const k = sal.length;
          bloquesDe(c, e, medio, prof + 1, protegido || prot, sal);
          if (cols) for (let i = k; i < sal.length; i++) if (!sal[i].columnasSeccion) sal[i].columnasSeccion = cols;
          break;
        }
        case 'tracked-changes': c.cosas.add('cambios'); break;
        default:
          if (INDICES.has(e.local)) { c.cosas.add('indices'); const cuerpo = hijo(e, 'index-body'); if (cuerpo) bloquesDe(c, cuerpo, medio, prof + 1, true, sal); }
          else if (e.local === 'index-title') bloquesDe(c, e, medio, prof + 1, true, sal);
          break;
      }
      // un salto de página DESPUÉS de un párrafo equivale a uno ANTES del siguiente
      if (antes > 0 && sal[antes - 1] && sal[antes - 1].saltoDespues && sal[antes] && sal[antes].tipo === 'p') { sal[antes].p = { ...sal[antes].p, saltoAntes: true }; sal[antes - 1].saltoDespues = false; }
    }
    return sal;
  }

  // ───────── Secciones (páginas maestras) ─────────
  function configuracionDe(est, maestra) {
    const mp = maestra && est.maestras.get(maestra);
    const pag = mp && est.pagina(A(mp)['page-layout-name']);
    const a = pag ? pag.a : {};
    const pt = E().pt;
    const v = (x, d) => { const n = pt(x); return n === null ? d : n; };
    const ancho = acotar(v(a['page-width'], 612), 72, MAX_PT), alto = acotar(v(a['page-height'], 792), 72, MAX_PT);
    const mt = v(a['margin-top'], 72), mb = v(a['margin-bottom'], 72);
    let izquierda = Math.max(v(a['margin-left'], 72), 0), derecha = Math.max(v(a['margin-right'], 72), 0);
    const conCab = mp && (hijo(mp, 'header') || hijo(mp, 'header-first') || hijo(mp, 'header-left')), conPie = mp && (hijo(mp, 'footer') || hijo(mp, 'footer-first') || hijo(mp, 'footer-left'));
    let arriba = mt + (conCab && pag && pag.cab ? v(pag.cab['min-height'], 0) + v(pag.cab['margin-bottom'], 0) : 0);
    let abajo = mb + (conPie && pag && pag.pie ? v(pag.pie['min-height'], 0) + v(pag.pie['margin-top'], 0) : 0);
    if (arriba + abajo > alto * 0.75) { const f = (alto * 0.75) / (arriba + abajo); arriba *= f; abajo *= f; }
    if (izquierda + derecha > ancho * 0.75) { const f = (ancho * 0.75) / (izquierda + derecha); izquierda *= f; derecha *= f; }
    const referencias = [];
    if (mp) {
      for (const [tag, cual] of [['', 'default'], ['-left', 'even'], ['-first', 'first']]) {
        if (hijo(mp, 'header' + tag)) referencias.push({ tipo: 'encabezado', cual, id: 'odt:' + maestra + ':encabezado:' + cual });
        if (hijo(mp, 'footer' + tag)) referencias.push({ tipo: 'pie', cual, id: 'odt:' + maestra + ':pie:' + cual });
      }
    }
    const primera = parseInt(a['first-page-number'], 10);
    return { ancho, alto, arriba, derecha, abajo, izquierda, encabezado: mt, pie: mb, continua: false, columnas: null, portada: referencias.some((r) => r.cual === 'first'),
      inicioPaginas: Number.isFinite(primera) ? acotar(primera, 0, 1e6) : undefined, referencias, maestra };
  }
  function calcularSecciones(c, bloques) {
    // una sección nueva empieza donde cambia la página maestra o donde empiezan o terminan las columnas de una sección de texto (continua, como en Word)
    const clave = (b) => (b && b.columnasSeccion ? b.columnasSeccion.n + ':' + b.columnasSeccion.espacio : '');
    const inicios = [{ idx: 0, maestra: c.est.primeraMaestra, columnas: clave(bloques[0]) ? bloques[0].columnasSeccion : null, continua: false }];
    bloques.forEach((b, i) => {
      if (i === 0) { if (b.tipo === 'p' && b.abreSeccion && c.est.maestras.has(b.abreSeccion)) inicios[0].maestra = b.abreSeccion; return; }
      const ultima = inicios[inicios.length - 1];
      if (b.tipo === 'p' && b.abreSeccion && c.est.maestras.has(b.abreSeccion)) inicios.push({ idx: i, maestra: b.abreSeccion, columnas: b.columnasSeccion || null, continua: false });
      else if (clave(b) !== clave(bloques[i - 1])) inicios.push({ idx: i, maestra: ultima.maestra, columnas: b.columnasSeccion || null, continua: true });
    });
    return inicios.map((s, k) => ({ ...configuracionDe(c.est, s.maestra), columnas: s.columnas, continua: s.continua, hasta: (k + 1 < inicios.length ? inicios[k + 1].idx : bloques.length) - 1 }));
  }

  // ───────── Entrada ─────────
  async function abrir(bytes) {
    if (typeof Blob !== 'undefined' && bytes instanceof Blob) bytes = await bytes.arrayBuffer();
    const zip = await ZipPropio.loadAsync(bytes);
    const paquete = PaqueteEditable.desdeZip(zip);
    const mime = await paquete.leer('mimetype');
    if (mime !== null && !/^\s*application\/vnd\.oasis\.opendocument\.text\s*$/.test(mime)) throw new Error('No es un documento de texto de OpenDocument (.odt).');
    const manifiesto = await paquete.leer('META-INF/manifest.xml');
    if (manifiesto && /encryption-data/.test(manifiesto)) { const e = new Error('El documento está protegido con contraseña; ábrelo en LibreOffice y guárdalo sin contraseña.'); e.codigo = 'cifrado'; throw e; }
    const xml = await paquete.leer('content.xml');
    if (xml === null) throw new Error('No es un documento de OpenDocument válido (falta content.xml).');
    const raiz = X().analizar(xml);
    const cuerpo = hijo(hijo(raiz, 'body'), 'text');
    if (!cuerpo) throw new Error('El documento no tiene cuerpo.');
    const xmlEstilos = await paquete.leer('styles.xml');
    let raizEstilos = xmlEstilos ? X().analizar(xmlEstilos) : null;
    const ajustesXml = await paquete.leer('settings.xml');
    const est = E().crear({ contenido: raiz, estilos: raizEstilos, settings: ajustesXml });
    const firmas = Object.keys(zip.files || {}).some((n) => /^META-INF\/documentsignatures\.xml$/i.test(n));
    const cosas = new Set();
    if (firmas) cosas.add('firmas');
    if (Object.keys(zip.files || {}).some((n) => /^(Basic|Scripts)\//.test(n))) cosas.add('macros');
    let regiones = 0;
    const c = { xml, est, m: est.m, ajustes: est.ajustes, cosas, notasPend: new Map(), listasPorId: new Map(), ultimaLista: new Map(), contadorListas: 0, caracteres: 0, anchoUtil: 468, rels: new Map() };
    const cfg0 = configuracionDe(est, est.primeraMaestra);
    c.anchoUtil = cfg0.ancho - cfg0.izquierda - cfg0.derecha;
    const bloques = bloquesDe(c, cuerpo, null);
    const secciones = calcularSecciones(c, bloques);
    // notas al pie y finales
    const notas = new Map();
    for (const [clave, n] of c.notasPend) notas.set(clave, { id: clave, n: n.n, bloques: n.cuerpo ? bloquesDe(c, n.cuerpo, null) : [] });
    // referencias de encabezados y pies: ids sintéticos que apuntan a un elemento de styles.xml
    const rels = c.rels;
    const registrar = () => { for (const s of secciones) for (const r of s.referencias) { const [, mp, tipo, cual] = r.id.split(':'); rels.set(r.id, { destino: 'styles.xml#' + mp + '/' + tipo + '/' + cual, externo: false }); } };
    registrar();

    const modelo = {
      formato: 'odt', notas, paquete, zip, m: c.m, est, ruta: 'content.xml', xml, xmlEstilos, raizContenido: raiz, rels, bloques, secciones, seccionFinal: secciones[secciones.length - 1], cosas,
      parrafos() {
        const sal = [];
        const rec = (lista) => { for (const b of lista) { if (b.tipo === 'p') sal.push(b); else for (const f of b.filas) for (const cel of f.celdas) rec(cel.bloques); } };
        rec(bloques);
        return sal;
      },
      releer(fragmento, base) {
        const r = X().analizar(fragmento);
        const piezas = [];
        leerInline({ ...c, xml: fragmento }, r, base, piezas, { espacio: true });
        return { raiz: r, piezas, texto: asignarTexto(piezas), seccion: null };
      },
      leerCompleto(fragmento, medio, previo) {
        const r = X().analizar(fragmento);
        return parrafo({ ...c, xml: fragmento }, r, medio || null, previo && previo.lista ? { lista: previo.lista, etiqueta: previo.etiqueta ? { texto: previo.etiqueta.texto, fmt: previo.etiqueta.fmt, sufijo: previo.etiqueta.sufijo, nivel: previo.etiqueta.nivel } : null, protegido: previo.protegido } : {});
      },
      leerBloque(fragmento, medio) {
        const r = X().analizar(fragmento);
        const c2 = { ...c, xml: fragmento };
        const b = r.local === 'table' ? tabla(c2, r, 0, false) : parrafo(c2, r, medio || null);
        const marcar = (x) => { x.base = fragmento; if (x.tipo === 'tabla') for (const f of x.filas) for (const ce of f.celdas) ce.bloques.forEach(marcar); };
        marcar(b);
        return b;
      },
      // Varios párrafos y listas sueltos (lo que sale de reescribir una lista, odt-listas.js): se leen como si fueran un contenedor propio; la base es el texto envuelto
      leerRegion(xml, medio) {
        const base = '<w>' + xml + '</w>', r = X().analizar(base);
        regiones += 100000;
        const c2 = { ...c, xml: base, listasPorId: new Map(), ultimaLista: new Map(), contadorListas: regiones, notasPend: new Map(), cosas: new Set() };
        const bs = bloquesDe(c2, r, medio || null);
        const marcar = (x) => { x.base = base; if (x.tipo === 'tabla') for (const f of x.filas) for (const ce of f.celdas) ce.bloques.forEach(marcar); };
        bs.forEach(marcar);
        return { base, bloques: bs };
      },
      renumerar() { return []; },
      recalcularSecciones() { modelo.secciones.splice(0, modelo.secciones.length, ...calcularSecciones(c, bloques)); modelo.seccionFinal = modelo.secciones[modelo.secciones.length - 1]; registrar(); },
      async recargar() {},
      // Un encabezado o un pie de styles.xml: sus bloques, con los estilos automáticos de ESE archivo
      // (lee el styles.xml de ahora, con lo que ya se haya escrito en otros encabezados; las posiciones de los bloques son las de ese texto)
      async parte(destino) {
        const m = /^styles\.xml#([^/]+)\/(encabezado|pie)\/(default|even|first)$/.exec(destino || '');
        const xmlAhora = m ? await paquete.leer('styles.xml') : null;
        if (!m || xmlAhora === null) return null;
        const raizAhora = X().analizar(xmlAhora);
        const mp = hijos(hijo(raizAhora, 'master-styles'), 'master-page').find((n) => A(n).name === m[1]);
        const nombre = (m[2] === 'encabezado' ? 'header' : 'footer') + (m[3] === 'even' ? '-left' : m[3] === 'first' ? '-first' : '');
        const el = hijo(mp, nombre);
        if (!el) return null;
        // los estilos automáticos de ESTE archivo, más los que se crearon al editar
        const estParte = E().crear({ contenido: null, estilos: raizAhora, settings: ajustesXml });
        if (modelo.dialecto && modelo.dialecto.estilosCreados) for (const x of modelo.dialecto.estilosCreados()) estParte.registrar(x.xml, x.nombrado);
        const c2 = { ...c, xml: xmlAhora, est: estParte, m: estParte.m, notasPend: new Map(), cosas: new Set() };
        return { ruta: 'styles.xml', xml: xmlAhora, rels: new Map(), bloques: bloquesDe(c2, el, null) };
      },
      // Abre un encabezado o pie como un documento aparte (la misma vista que DocxModelo.editarParte)
      async editarParte(destino, secIdx) {
        const p = await modelo.parte(destino);
        if (!p) return null;
        const base = modelo.secciones[Math.min(Math.max(secIdx || 0, 0), modelo.secciones.length - 1)];
        const cfg = { ...base, arriba: 8, abajo: 8, encabezado: 0, pie: 0, continua: false, columnas: null, portada: false, inicioPaginas: undefined, referencias: [] };
        const vista = Object.create(modelo);
        const flota = (l, sal) => { for (const b of l) { if (b.tipo === 'p') sal.push(b); else for (const f of b.filas) for (const ce of f.celdas) flota(ce.bloques, sal); } return sal; };
        // el mismo dialecto, pero los estilos que crea son para styles.xml (no reutiliza los automáticos de content.xml)
        const d = modelo.dialecto;
        if (d && d.enEstilos) {
          const enParte = (f) => (...a) => d.enEstilos(() => d[f](...a));
          vista.dialecto = Object.assign(Object.create(d), { formatoCaracter: enParte('formatoCaracter'), formatoParrafo: enParte('formatoParrafo'), estiloConservando: enParte('estiloConservando') });
        }
        Object.assign(vista, {
          ruta: p.ruta, xml: p.xml, rels: p.rels, bloques: p.bloques, notas: new Map(), seccionFinal: cfg,
          secciones: [{ ...cfg, hasta: p.bloques.length - 1 }],
          parrafos: () => flota(vista.bloques, []),
          renumerar: () => [],
          recalcularSecciones() { vista.secciones[0].hasta = vista.bloques.length - 1; },
        });
        return vista;
      },
      // Un encabezado o pie que el documento no tiene, en la página maestra de la última sección: la página (style:page-layout) se copia con la altura del
      // encabezado o pie, la maestra apunta a la copia y recibe el elemento vacío. Devuelve el destino de `parte`, o null.
      async crearParte(tipo, cual) {
        const maestra = modelo.seccionFinal.maestra, xml = await paquete.leer('styles.xml');
        if (!maestra || xml === null) return null;
        const r = X().analizar(xml);
        const mp = hijos(hijo(r, 'master-styles'), 'master-page').find((n) => A(n).name === maestra);
        if (!mp) return null;
        const esEnc = tipo === 'encabezado', loc = (esEnc ? 'header' : 'footer') + (cual === 'even' ? '-left' : cual === 'first' ? '-first' : '');
        const destino = 'styles.xml#' + maestra + '/' + tipo + '/' + cual;
        if (hijo(mp, loc)) return destino;
        const pfS = mp.nombre.slice(0, mp.nombre.indexOf(':') + 1);
        const uri = (u, def) => { const k = Object.keys(r.attrs).find((q) => q.startsWith('xmlns:') && r.attrs[q] === u); return k ? k.slice(6) : null || def; };
        const URI_FO = 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0', URI_TEXT = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0';
        const ed = [];
        let nuevoLayout = null;
        const capa = A(mp)['page-layout-name'], as = hijo(r, 'automatic-styles');
        const layout = hijos(as, 'page-layout').find((n) => A(n).name === capa);
        const sty = esEnc ? 'header-style' : 'footer-style';
        if (layout && !hijo(layout, sty)) {
          let n = 1;
          while (xml.includes('style:name="MCpm' + n + '"')) n++;
          const nuevo = 'MCpm' + n, claveN = Object.keys(layout.attrs).find((q) => q.slice(q.indexOf(':') + 1) === 'name');
          const pl = layout.nombre.slice(0, layout.nombre.indexOf(':') + 1);
          const props = `<${pl}${sty} xmlns:fo="${URI_FO}"><${pl}header-footer-properties fo:min-height="0.6cm" fo:${esEnc ? 'margin-bottom' : 'margin-top'}="0.2cm"/></${pl}${sty}>`;
          let copia = xml.slice(layout.ini, layout.fin);
          const rel = layout.iniCierre < 0 ? null : layout.iniCierre - layout.ini;
          const pieEl = hijo(layout, 'footer-style');
          if (rel === null) copia = copia.replace(/\s*\/>$/, '>') + props + `</${layout.nombre}>`;
          else { const donde = esEnc && pieEl ? pieEl.ini - layout.ini : rel; copia = copia.slice(0, donde) + props + copia.slice(donde); }
          copia = copia.replace(/^<[^>]*>/, (t) => X().ponerAtributo(t, claveN, nuevo));
          ed.push({ ini: layout.fin, fin: layout.fin, texto: copia });
          nuevoLayout = nuevo;
        }
        const ORDEN = ['header', 'header-left', 'header-first', 'footer', 'footer-left', 'footer-first'];
        const rango = ORDEN.indexOf(loc);
        const despues = mp.hijos.find((h) => ORDEN.indexOf(h.local) > rango);
        const donde = despues ? despues.ini : mp.iniCierre;
        const elem = `<${pfS}${loc} xmlns:text="${URI_TEXT}"><text:p/></${pfS}${loc}>`;
        let tagM = xml.slice(mp.ini, mp.finApertura);
        if (nuevoLayout) tagM = X().ponerAtributo(tagM, Object.keys(mp.attrs).find((q) => q.slice(q.indexOf(':') + 1) === 'page-layout-name'), nuevoLayout);
        if (mp.iniCierre < 0) ed.push({ ini: mp.ini, fin: mp.fin, texto: tagM.replace(/\s*\/>$/, '>') + elem + `</${mp.nombre}>` });
        else { if (nuevoLayout) ed.push({ ini: mp.ini, fin: mp.finApertura, texto: tagM }); ed.push({ ini: donde, fin: donde, texto: elem }); }
        void uri;
        const nuevoXml = X().aplicar(xml, ed);
        paquete.escribir('styles.xml', nuevoXml);
        raizEstilos = X().analizar(nuevoXml);
        est.releerPaginas(raizEstilos);
        modelo.recalcularSecciones();
        return destino;
      },
    };
    // cómo se corta el XML de un párrafo al editar (odt-parrafo.js); sin él el modelo es de solo lectura
    if (globalThis.OdtParrafo) modelo.dialecto = globalThis.OdtParrafo.crear(modelo);
    return modelo;
  }

  globalThis.OdtModelo = { abrir };
})();

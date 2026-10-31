/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Modelo de un .docx para el editor de Documentos (Fase 6, etapa 6.1). Sin DOM: corre igual en Node.
//
// Lee word/document.xml (y sus encabezados y pies) a BLOQUES con el formato efectivo ya resuelto (estilos con herencia, tema, numeración,
// estilo de tabla) y con la POSICIÓN de cada párrafo, tabla, fila, celda y corrida en el texto del XML (`ini`, `fin`). Esa posición es lo
// que permite editar solo lo editado: lo que el modelo no entiende (marcadores, controles, comentarios, revisiones, dibujos) no se toca.
// El formato sale de DocxEstilos (docx-estilos.js, copia del sitio); el XML con posiciones, de XmlEditar.
//
// Requiere zip-propio.js, xml-editar.js, paquete-editable.js, docx-base.js y docx-estilos.js.
//
//   const modelo = await DocxModelo.abrir(bytes)
//   modelo.bloques        → [bloque]       (párrafos y tablas del cuerpo, en orden; un control de contenido (sdt) se desenvuelve)
//   modelo.secciones      → [{ ...cfg, hasta }]   cfg: ancho, alto, arriba, derecha, abajo, izquierda (pt), encabezado, pie, continua, columnas, portada, referencias
//   modelo.parrafos()     → todos los párrafos (también los de tablas), en orden de documento
//   await modelo.parte(ruta) → { ruta, xml, bloques } de un encabezado o pie (word/headerN.xml)
//   modelo.paquete        → PaqueteEditable (guardar: await modelo.paquete.generar())
//
// Párrafo: { tipo:'p', ini, fin, p (formato efectivo), r (formato base del texto), marca (de la marca de párrafo), estilo, etiqueta, piezas, texto, seccion }
//   etiqueta: null o { texto, fmt, sufijo, r, nivel } (viñeta o número de lista)
//   pieza: { t, r, ini, fin, desde, hasta, ... } con t = texto | tab | br | pagina | columna | campo | imagen | cuadro | dibujo | ancla
//     ini/fin: la corrida (w:r) en el XML; desde/hasta: posición dentro de `texto` (UTF-16). campo, imagen, cuadro y dibujo valen un carácter (U+FFFC).
//   texto: el texto plano del párrafo (tab = \t, salto de línea = U+2028, salto de página = \f)
// Tabla: { tipo:'tabla', ini, fin, tbl, grid (pt), nCols, filas:[{ ini, fin, tr, celdas:[{ ini, fin, col, span, filas, omitir, tc, medio, anchoDecl, bloques }] }] }
(function () {
  'use strict';
  const X = () => globalThis.XmlEditar;
  const D = () => globalThis.DocxBase;
  const MAX_PROFUNDIDAD = 8;       // tablas anidadas y controles de contenido
  const MAX_PT = 1584;
  const acotar = (x, min, max) => (Number.isFinite(x) ? Math.min(Math.max(x, min), max) : min);
  const emuPt = (v) => (Number.isFinite(parseFloat(v)) ? parseFloat(v) / 12700 : 0);

  // Árbol sencillo { n, a, h, t } (el que lee DocxEstilos) a partir de un nodo con posición.
  function aArbol(xml, nodo) {
    if (!nodo) return null;
    const a = {};
    for (const k of Object.keys(nodo.attrs)) a[k.slice(k.indexOf(':') + 1)] = nodo.attrs[k];
    const e = { n: nodo.local, a, h: nodo.hijos.map((x) => aArbol(xml, x)), t: '' };
    if (!nodo.hijos.length && nodo.iniCierre >= 0) e.t = X().decodificar(X().interior(xml, nodo));
    return e;
  }
  const hijos = (n, local) => (n ? n.hijos.filter((h) => h.local === local) : []);
  const hijo = (n, local) => (n ? n.hijos.find((h) => h.local === local) || null : null);
  // Hijos con ese nombre, también los de un control de contenido (sdt > sdtContent).
  function desenvolver(n, local, prof = 0) {
    const sal = [];
    for (const x of n.hijos) {
      if (x.local === local) sal.push(x);
      else if (x.local === 'sdt' && prof < MAX_PROFUNDIDAD) { const c = hijo(x, 'sdtContent'); if (c) sal.push(...desenvolver(c, local, prof + 1)); }
    }
    return sal;
  }

  // Formato de un trozo: base del párrafo + estilo de carácter + formato directo.
  const formatoRun = (c, baseR, rPr) => {
    const directo = c.m.leerRPr(rPr ? aArbol(c.xml, rPr) : null);
    return c.m.mezclar(c.m.mezclar(baseR, c.m.estilos.caracter(directo.estilo)), directo);
  };

  // ───────── Texto en línea ─────────
  function leerDibujo(c, nodo, r, run) {
    const ar = aArbol(c.xml, nodo);
    const { hijo: h, buscar } = D();
    const raiz = ar.n === 'AlternateContent' ? (h(ar, 'Choice') || h(ar, 'Fallback')) : ar;
    const dib = raiz && (h(raiz, 'drawing') || raiz);
    const inl = dib && h(dib, 'inline'), anc = dib && h(dib, 'anchor');
    const caja = inl || anc;
    const ext = caja && h(caja, 'extent');
    if (!ext) return null;
    const ancho = acotar(emuPt(ext.a.cx), 0, MAX_PT), alto = acotar(emuPt(ext.a.cy), 0, MAX_PT);
    const doc = h(caja, 'docPr');
    const base = { r, ini: run.ini, fin: run.fin, ancho, alto, enLinea: !!inl, descr: doc ? doc.a.descr || '' : '' };
    if (anc) {
      // posición y envoltura de un dibujo flotante (wp:anchor)
      const pos = (p) => { if (!p) return null; const al = h(p, 'align'), of = h(p, 'posOffset'); return { rel: p.a.relativeFrom, align: al ? al.t.trim() : null, offset: of ? emuPt(of.t) : 0 }; };
      const envoltura = anc.h.map((n) => n.n).find((n) => /^wrap/.test(n)) || 'wrapNone';
      base.ancla = {
        h: pos(h(anc, 'positionH')), v: pos(h(anc, 'positionV')), envoltura, detras: anc.a.behindDoc === '1',
        dist: { t: emuPt(anc.a.distT), b: emuPt(anc.a.distB), l: emuPt(anc.a.distL), r: emuPt(anc.a.distR) },
      };
    }
    const blip = buscar(caja, 'blip');
    if (blip && blip.a.embed) {
      const rel = c.rels.get(blip.a.embed);
      return { t: 'imagen', ...base, rid: blip.a.embed, destino: rel && !rel.externo ? rel.destino : null };
    }
    const wsp = buscar(caja, 'wsp');
    if (wsp && buscar(wsp, 'txbxContent')) return { t: 'cuadro', ...base };
    return { t: 'dibujo', ...base };
  }

  function leerRun(c, run, baseR, salida, ctx) {
    const r = formatoRun(c, baseR, hijo(run, 'rPr'));
    const dato0 = { r, ini: run.ini, fin: run.fin };
    for (const e of run.hijos) {
      // ei/ef: el elemento hijo de la corrida que da esta pieza (para editar el texto sin tocar el resto)
      const dato = { ...dato0, ei: e.ini, ef: e.fin };
      switch (e.local) {
        case 't': if (!ctx.enInstruccion) salida.push({ t: 'texto', s: X().decodificar(X().interior(c.xml, e)), ...dato, tn: e.nombre, ti: e.iniCierre >= 0 ? e.finApertura : undefined, tf: e.iniCierre >= 0 ? e.iniCierre : undefined }); break;
        case 'tab': if (!ctx.enInstruccion) salida.push({ t: 'tab', ...dato }); break;
        case 'br': case 'cr': {
          const tipo = X().attr(e, 'type');
          if (tipo === 'page') salida.push({ t: 'pagina', ...dato });
          else if (tipo === 'column') salida.push({ t: 'columna', ...dato });
          else salida.push({ t: 'br', limpiar: X().attr(e, 'clear'), ...dato });
          break;
        }
        case 'noBreakHyphen': salida.push({ t: 'texto', s: '‑', ...dato }); break;
        case 'softHyphen': salida.push({ t: 'texto', s: '­', ...dato }); break;
        case 'fldChar': {
          const tipo = X().attr(e, 'fldCharType');
          if (tipo === 'begin') ctx.campos.push({ fase: 'instr', instr: '' });
          else if (tipo === 'separate') {
            const f = ctx.campos[ctx.campos.length - 1];
            if (f) {
              f.fase = 'resultado';
              const nombre = /^\s*(PAGE|NUMPAGES)\b/i.exec(f.instr);
              if (nombre) { f.campo = true; salida.push({ t: 'campo', nombre: nombre[1].toUpperCase(), ...dato }); }   // el número vale al pintar cada hoja, no el guardado
            }
          } else if (tipo === 'end') ctx.campos.pop();
          ctx.enInstruccion = ctx.campos.some((f) => f.fase === 'instr' || f.campo);
          break;
        }
        case 'instrText': { const f = ctx.campos[ctx.campos.length - 1]; if (f) f.instr += X().decodificar(X().interior(c.xml, e)); break; }
        case 'footnoteReference': case 'endnoteReference': {
          const fin = e.local === 'endnoteReference';
          const clave = (fin ? 'e' : 'f') + X().attr(e, 'id');
          salida.push({ t: 'texto', s: String((c.numNotas && c.numNotas.get(clave)) || '*'), ...dato, r: { ...r, alineacionVertical: 'superscript' }, nota: clave });
          break;
        }
        case 'footnoteRef': case 'endnoteRef':
          salida.push({ t: 'texto', s: String(c.numeroNota || '*'), ...dato, r: { ...r, alineacionVertical: 'superscript' } });
          break;
        case 'drawing': case 'AlternateContent': {
          const img = leerDibujo(c, e, r, run);
          if (img) salida.push({ ...img, ei: e.ini, ef: e.fin });
          break;
        }
        default: break;   // instrText, delText, lastRenderedPageBreak, pict, rPr…
      }
    }
  }

  function leerInline(c, nodo, baseR, salida, ctx, prof = 0) {
    if (prof > 60) return;
    for (const e of nodo.hijos) {
      switch (e.local) {
        case 'r': leerRun(c, e, baseR, salida, ctx); break;
        case 'hyperlink': {
          const k = salida.length;
          leerInline(c, e, baseR, salida, ctx, prof + 1);
          const id = X().attr(e, 'id'), rel = id ? c.rels.get(id) : null;
          const href = X().attr(e, 'anchor') ? '#' + X().attr(e, 'anchor') : rel ? D().hrefSeguro(rel.destino) : '';
          if (href) for (let i = k; i < salida.length; i++) salida[i].enlace = href;
          break;
        }
        case 'fldSimple': {
          const nombre = /^\s*(PAGE|NUMPAGES)\b/i.exec(X().attr(e, 'instr') || '');
          if (nombre) salida.push({ t: 'campo', nombre: nombre[1].toUpperCase(), r: formatoRun(c, baseR, hijo(hijo(e, 'r'), 'rPr')), ini: e.ini, fin: e.fin, ei: e.ini, ef: e.fin });
          else leerInline(c, e, baseR, salida, ctx, prof + 1);
          break;
        }
        case 'smartTag': case 'sdt': case 'sdtContent': case 'ins': case 'moveTo': case 'customXml': case 'bdo': case 'dir':
          leerInline(c, e, baseR, salida, ctx, prof + 1); break;
        case 'bookmarkStart': { const nombre = X().attr(e, 'name'); if (nombre && !/^_/.test(nombre)) salida.push({ t: 'ancla', nombre, ini: e.ini, fin: e.fin }); break; }
        case 'AlternateContent': { const dentro = hijo(e, 'Choice') || hijo(e, 'Fallback'); if (dentro) leerInline(c, dentro, baseR, salida, ctx, prof + 1); break; }
        default: break;   // pPr, proofErr, del, moveFrom, comentarios, marcas de revisión…
      }
    }
  }

  // ───────── Párrafo ─────────
  // El texto plano de un párrafo a partir de sus piezas; deja en cada una su `desde` y `hasta`.
  function asignarTexto(piezas) {
    let texto = '';
    for (const z of piezas) {
      z.desde = texto.length;
      if (z.t === 'texto') texto += z.s;
      else if (z.t === 'tab') texto += '\t';
      else if (z.t === 'br') texto += '\u2028';
      else if (z.t === 'pagina' || z.t === 'columna') texto += '\f';
      else if (z.t !== 'ancla') texto += '\uFFFC';
      z.hasta = texto.length;
    }
    return texto;
  }
  const simboloViñeta = (s) => (/^[○◦□▫o]$/.test(s) ? 'o' : '•');
  function parrafo(c, p, medio) {
    const m = c.m;
    const pPr = hijo(p, 'pPr');
    const pDirecto = m.leerPPr(pPr ? aArbol(c.xml, pPr) : null);
    const est = m.estilos.parrafo(pDirecto.estilo, medio);
    const pedida = pDirecto.numeracion || est.p.numeracion;
    let efectivo = m.mezclar(est.p, null), etiqueta = null;
    if (pedida && pedida.numId && pedida.numId !== '0' && c.numerador) {
      const e = c.numerador.siguiente(pedida.numId, pedida.nivel || '0');
      if (e) {
        efectivo = m.mezclar(efectivo, e.nivel.p);
        etiqueta = { texto: e.nivel.fmt === 'bullet' ? simboloViñeta(e.texto) : e.texto, fmt: e.nivel.fmt, sufijo: e.nivel.sufijo || 'tab', nivel: e.nivel, rNivel: e.nivel.r };
      }
    }
    efectivo = m.mezclar(efectivo, pDirecto);
    const baseR = est.r;
    const marca = m.mezclar(baseR, efectivo.marca);
    if (etiqueta) { etiqueta.r = m.mezclar(marca, etiqueta.rNivel); delete etiqueta.rNivel; }
    const piezas = [];
    leerInline(c, p, baseR, piezas, { campos: [], enInstruccion: false });
    const texto = asignarTexto(piezas);
    const sectPr = pPr && hijo(pPr, 'sectPr');
    return { tipo: 'p', medio: medio || null, ini: p.ini, fin: p.fin, estilo: pDirecto.estilo || null, numPedida: pedida && pedida.numId && pedida.numId !== '0' ? pedida : null, p: efectivo, r: baseR, marca, etiqueta, piezas, texto, seccion: sectPr ? leerSeccion(c, sectPr) : null };
  }

  // ───────── Tablas ─────────
  const leerAncho = (e) => {
    if (!e || e.a.w === undefined) return undefined;
    const crudo = String(e.a.w);
    return { tipo: e.a.type || 'dxa', valor: parseFloat(crudo), porcentaje: /%$/.test(crudo) ? parseFloat(crudo) : undefined };
  };
  // Formatos condicionales del estilo de tabla que valen para la celda (de menor a mayor prioridad).
  function condicionalesDe(estiloT, aspecto, tblPr, i, j, nFilas, nCols) {
    const sal = [];
    const k = estiloT.condicionales || {};
    const a = aspecto || { primeraFila: true, primeraColumna: true };
    const bandaF = tblPr.bandasFila || 1, bandaC = tblPr.bandasColumna || 1;
    const pF = a.primeraFila && i === 0, uF = a.ultimaFila && i === nFilas - 1, pC = a.primeraColumna && j === 0, uC = a.ultimaColumna && j === nCols - 1;
    if (k.wholeTable) sal.push(k.wholeTable);
    if (!a.sinBandasV && !pC && !uC) { const b = Math.floor((j - (a.primeraColumna ? 1 : 0)) / bandaC) % 2 === 0 ? k.band1Vert : k.band2Vert; if (b) sal.push(b); }
    if (!a.sinBandasH && !pF && !uF) { const b = Math.floor((i - (a.primeraFila ? 1 : 0)) / bandaF) % 2 === 0 ? k.band1Horz : k.band2Horz; if (b) sal.push(b); }
    if (pC && k.firstCol) sal.push(k.firstCol);
    if (uC && k.lastCol) sal.push(k.lastCol);
    if (pF && k.firstRow) sal.push(k.firstRow);
    if (uF && k.lastRow) sal.push(k.lastRow);
    if (pF && pC && k.nwCell) sal.push(k.nwCell);
    if (pF && uC && k.neCell) sal.push(k.neCell);
    if (uF && pC && k.swCell) sal.push(k.swCell);
    if (uF && uC && k.seCell) sal.push(k.seCell);
    return sal;
  }

  function tabla(c, t, prof) {
    const m = c.m;
    const tblPr = hijo(t, 'tblPr');
    const directo = m.leerTblPr(tblPr ? aArbol(c.xml, tblPr) : null);
    const estiloT = m.estilos.tabla(directo.estilo);
    const tbl = m.mezclar(estiloT.tbl, directo);
    const anchoDecl = leerAncho(aArbol(c.xml, hijo(tblPr, 'tblW')));
    const grid = hijos(hijo(t, 'tblGrid'), 'gridCol').map((g) => (parseFloat(X().attr(g, 'w')) || 0) / 20);
    const filasXml = desenvolver(t, 'tr');
    const infos = filasXml.map((f) => {
      let col = 0;
      return desenvolver(f, 'tc').map((cel) => {
        const tcPr = hijo(cel, 'tcPr');
        const pr = m.leerTcPr(tcPr ? aArbol(c.xml, tcPr) : null);
        const span = Math.min(Math.max(Math.floor(pr.columnas) || 1, 1), Math.max(grid.length, 64));
        const info = { ini: cel.ini, fin: cel.fin, nodo: cel, pr, col, span, filas: 1, omitir: false, anchoDecl: leerAncho(aArbol(c.xml, hijo(tcPr, 'tcW'))) };
        col += span;
        return info;
      });
    });
    const nCols = Math.max(infos.reduce((mx, f) => Math.max(mx, f.reduce((a, x) => a + x.span, 0)), 0), grid.length, 1);
    infos.forEach((fila, i) => fila.forEach((info) => {
      if (info.pr.combinadaVertical === 'restart') {
        for (let k = i + 1; k < infos.length; k++) {
          const sig = infos[k].find((x) => x.col === info.col);
          if (sig && sig.pr.combinadaVertical === 'continue') { sig.omitir = true; info.filas++; } else break;
        }
      }
    }));
    const filas = infos.map((fila, i) => {
      const trPr = hijo(filasXml[i], 'trPr');
      return {
        ini: filasXml[i].ini, fin: filasXml[i].fin, tr: m.leerTrPr(trPr ? aArbol(c.xml, trPr) : null),
        celdas: fila.map((info) => {
          let tc = {}, medioP = {}, medioR = {};
          for (const k of condicionalesDe(estiloT, tbl.aspecto, tbl, i, info.col, infos.length, nCols)) { tc = m.mezclar(tc, k.tc); medioP = m.mezclar(medioP, k.p); medioR = m.mezclar(medioR, k.r); }
          tc = m.mezclar(m.mezclar(estiloT.tc, tc), info.pr);
          const medio = { p: m.mezclar(estiloT.p, medioP), r: m.mezclar(estiloT.r, medioR) };
          return { ini: info.ini, fin: info.fin, col: info.col, span: info.span, filas: info.filas, omitir: info.omitir, tc, medio, anchoDecl: info.anchoDecl, bloques: bloquesDe(c, info.nodo, medio, prof + 1) };
        }),
      };
    });
    return { tipo: 'tabla', ini: t.ini, fin: t.fin, tbl, grid, nCols, anchoDecl, filas };
  }

  // Bloques de un contenedor (cuerpo, celda, encabezado): párrafos y tablas; un control de contenido se desenvuelve.
  function bloquesDe(c, contenedor, medio, prof = 0) {
    const sal = [];
    if (prof > MAX_PROFUNDIDAD) return sal;
    for (const e of contenedor.hijos) {
      if (e.local === 'p') sal.push(parrafo(c, e, medio));
      else if (e.local === 'tbl') sal.push(tabla(c, e, prof));
      else if (e.local === 'sdt') { const dentro = hijo(e, 'sdtContent'); if (dentro) sal.push(...bloquesDe(c, dentro, medio, prof + 1)); }
    }
    return sal;
  }

  // ───────── Secciones ─────────
  function leerSeccion(c, sectPr) {
    const ar = sectPr ? aArbol(c.xml, sectPr) : null;
    const { hijo: h, lista } = D();
    const sz = ar && h(ar, 'pgSz'), mg = ar && h(ar, 'pgMar'), cols = ar && h(ar, 'cols'), num = ar && h(ar, 'pgNumType');
    const v = (x, defecto) => { const n = parseFloat(x); return Number.isFinite(n) ? n / 20 : defecto; };
    // Un archivo hostil puede pedir una hoja de 0 o de kilómetros, o márgenes más altos que la hoja: la hoja va de 1 a 22 pulgadas
    // y al cuerpo le queda al menos un cuarto de ella. Un margen negativo arriba o abajo es el margen «exacto»: vale su valor absoluto.
    const ancho = acotar(v(sz && sz.a.w, 612), 72, MAX_PT), alto = acotar(v(sz && sz.a.h, 792), 72, MAX_PT);
    let arriba = Math.abs(v(mg && mg.a.top, 72)), abajo = Math.abs(v(mg && mg.a.bottom, 72));
    let izquierda = Math.max(v(mg && mg.a.left, 72), 0), derecha = Math.max(v(mg && mg.a.right, 72), 0);
    if (arriba + abajo > alto * 0.75) { const f = (alto * 0.75) / (arriba + abajo); arriba *= f; abajo *= f; }
    if (izquierda + derecha > ancho * 0.75) { const f = (ancho * 0.75) / (izquierda + derecha); izquierda *= f; derecha *= f; }
    const inicio = num ? parseInt(num.a.start, 10) : NaN;
    return {
      ancho, alto, arriba, derecha, abajo, izquierda,
      encabezado: Math.min(v(mg && mg.a.header, 36), alto / 4), pie: Math.min(v(mg && mg.a.footer, 36), alto / 4),   // un encabezado a 2.000.000 twips sacaba el cuerpo de la hoja
      continua: !!(ar && h(ar, 'type') && h(ar, 'type').a.val === 'continuous'),
      columnas: cols && parseInt(cols.a.num, 10) > 1 ? { n: parseInt(cols.a.num, 10), espacio: v(cols.a.space, 36), linea: cols.a.sep === '1' || cols.a.sep === 'true' } : null,
      portada: !!(ar && h(ar, 'titlePg')),
      inicioPaginas: Number.isFinite(inicio) ? acotar(inicio, 0, 1e6) : undefined,
      referencias: ar ? lista(ar, 'headerReference').map((x) => ({ tipo: 'encabezado', cual: x.a.type || 'default', id: x.a.id }))
        .concat(lista(ar, 'footerReference').map((x) => ({ tipo: 'pie', cual: x.a.type || 'default', id: x.a.id }))) : [],
    };
  }

  // ───────── Entrada ─────────
  async function abrir(bytes) {
    if (typeof Blob !== 'undefined' && bytes instanceof Blob) bytes = await bytes.arrayBuffer();
    const zip = await ZipPropio.loadAsync(bytes);
    const paquete = PaqueteEditable.desdeZip(zip);
    let ruta = 'word/document.xml';
    const raices = await paquete.leer('_rels/.rels');
    if (raices) {
      const r = X().hijos(X().analizar(raices), 'Relationship').find((x) => /\/officeDocument$/.test(x.attrs.Type || ''));
      if (r && r.attrs.Target) ruta = PaqueteEditable.resolver('x', r.attrs.Target);
    }
    let xml = await paquete.leer(ruta);
    if (xml === null) throw new Error('No es un documento de Word válido (falta word/document.xml).');
    // Un cuerpo sin ningún párrafo (solo sectPr, como lo deja python-docx al empezar) no deja dónde poner el cursor: se le agrega el párrafo vacío que Word exige
    const pre = /<(\w+):body[ >]/.exec(xml);
    if (pre && !new RegExp('<' + pre[1] + ':(p|tbl|sdt|customXml)[ >/]').test(xml.slice(pre.index))) {
      const finCuerpo = xml.lastIndexOf('</' + pre[1] + ':body>'), sect = xml.lastIndexOf('<' + pre[1] + ':sectPr');
      const donde = sect > pre.index ? sect : finCuerpo;
      if (donde > 0) { xml = xml.slice(0, donde) + '<' + pre[1] + ':p/>' + xml.slice(donde); paquete.escribir(ruta, xml); }
    }
    const rels = await D().leerRels(zip, ruta);
    const m = await DocxEstilos.cargar({ zip, rels });
    const cuerpo = hijo(X().analizar(xml), 'body');
    if (!cuerpo) throw new Error('El documento no tiene cuerpo.');
    // Notas al pie: se numeran por orden de aparición en el cuerpo; su texto está en word/footnotes.xml
    const notas = new Map();
    const rutaNotas = D().porTipo(rels, '/footnotes');
    const textoNotas = rutaNotas ? await paquete.leer(rutaNotas) : null;
    const nodosNotas = new Map();
    let raizNotas = null;
    if (textoNotas) {
      raizNotas = X().analizar(textoNotas);
      for (const n of hijos(raizNotas, 'footnote')) { const t = X().attr(n, 'type'); if (!t || t === 'normal') nodosNotas.set('f' + X().attr(n, 'id'), n); }
    }
    const numNotas = new Map();
    if (nodosNotas.size) {
      const re = /<(?:[\w.-]+:)?footnoteReference\b[^>]*?\bid="(-?\d+)"/g;
      let mm;
      while ((mm = re.exec(xml))) { const k = 'f' + mm[1]; if (nodosNotas.has(k) && !numNotas.has(k)) numNotas.set(k, numNotas.size + 1); }
    }
    const c = { m, xml, rels, numerador: m.numeracion.numerador(), numNotas };
    const bloques = bloquesDe(c, cuerpo, null);
    if (numNotas.size) {
      const relsN = await D().leerRels(zip, rutaNotas);
      for (const [k, n] of numNotas) notas.set(k, { id: k, n, bloques: bloquesDe({ m, xml: textoNotas, rels: relsN, numerador: null, numeroNota: n }, nodosNotas.get(k), null) });
      const sep = hijos(raizNotas, 'footnote').find((nn) => X().attr(nn, 'type') === 'separator');
      if (sep) notas.separador = bloquesDe({ m, xml: textoNotas, rels: relsN, numerador: null }, sep, null);
    }

    // Secciones: un párrafo con sectPr cierra la suya; la última está en el propio cuerpo.
    const secciones = [];
    bloques.forEach((b, i) => { if (b.tipo === 'p' && b.seccion) secciones.push({ ...b.seccion, hasta: i }); });
    const final = hijo(cuerpo, 'sectPr');
    const seccionFinal = leerSeccion(c, final);
    secciones.push({ ...seccionFinal, hasta: bloques.length - 1 });

    const modelo = {
      notas, paquete, zip, m, ruta, xml, rels, bloques, secciones, seccionFinal,
      parrafos() {
        const sal = [];
        const rec = (lista) => { for (const b of lista) { if (b.tipo === 'p') sal.push(b); else for (const f of b.filas) for (const cel of f.celdas) rec(cel.bloques); } };
        rec(bloques);
        return sal;
      },
      // Lee de nuevo un solo párrafo (su texto XML, <w:p>…</w:p>) tras editarlo: piezas, texto plano y sección. `base` es el formato base del texto.
      releer(fragmento, base) {
        const raiz = X().analizar(fragmento);
        const c2 = { m, xml: fragmento, rels, numerador: null, numNotas };
        const piezas = [];
        leerInline(c2, raiz, base, piezas, { campos: [], enInstruccion: false });
        const texto = asignarTexto(piezas);
        const pPr = hijo(raiz, 'pPr'), sectPr = pPr && hijo(pPr, 'sectPr');
        return { raiz, piezas, texto, seccion: sectPr ? leerSeccion(c2, sectPr) : null };
      },
      // Lee de nuevo un párrafo entero (formato efectivo incluido) desde su XML: para cuando cambia el estilo o la numeración. La etiqueta de lista
      // sale con el número que tocaría a un párrafo suelto; `renumerar` la corrige después.
      leerCompleto(fragmento, medio) {
        const raiz = X().analizar(fragmento);
        return parrafo({ m, xml: fragmento, rels, numerador: m.numeracion.numerador(), numNotas }, raiz, medio || null);
      },
      // Vuelve a numerar las listas de todo el documento (tras partir, unir o borrar párrafos). Devuelve los párrafos cuya etiqueta cambió.
      renumerar() {
        const num = m.numeracion.numerador(), cambiados = [];
        for (const b of modelo.parrafos()) {
          if (!b.numPedida || !b.etiqueta) continue;
          const e = num.siguiente(b.numPedida.numId, b.numPedida.nivel || '0');
          if (!e) continue;
          const t = e.nivel.fmt === 'bullet' ? simboloViñeta(e.texto) : e.texto;
          if (b.etiqueta.texto !== t) { b.etiqueta.texto = t; cambiados.push(b); }
        }
        return cambiados;
      },
      // Lee un bloque nuevo (un <w:p> o un <w:tbl>) desde su propio XML. Sus posiciones valen en ESE texto, que queda en `base` de cada bloque.
      leerBloque(fragmento, medio) {
        const raiz = X().analizar(fragmento);
        const c = { m, xml: fragmento, rels, numerador: m.numeracion.numerador(), numNotas };
        const b = raiz.local === 'tbl' ? tabla(c, raiz, 0) : parrafo(c, raiz, medio || null);
        const marcar = (x) => { x.base = fragmento; if (x.tipo === 'tabla') for (const f of x.filas) for (const ce of f.celdas) ce.bloques.forEach(marcar); };
        marcar(b);
        return b;
      },
      // Las secciones salen de los párrafos que cierran una (w:sectPr) y de la sección final del cuerpo; se vuelven a calcular tras cada cambio
      recalcularSecciones() {
        const sal = [];
        bloques.forEach((b, i) => { if (b.tipo === 'p' && b.seccion) sal.push({ ...b.seccion, hasta: i }); });
        sal.push({ ...modelo.seccionFinal, hasta: bloques.length - 1 });
        modelo.secciones.splice(0, modelo.secciones.length, ...sal);
      },
      // Vuelve a leer los estilos y la numeración del paquete (tras agregar un estilo o una lista)
      async recargar() { const nuevo = await DocxEstilos.cargar({ zip, rels }); m.estilos = nuevo.estilos; m.numeracion = nuevo.numeracion; },
      // Un encabezado o un pie: con sus propias relaciones (imágenes). No lleva numeración propia.
      async parte(destino) {
        const t = await paquete.leer(destino);
        if (t === null) return null;
        const raiz = X().analizar(t);
        const relsP = await D().leerRels(zip, destino);
        return { ruta: destino, xml: t, rels: relsP, bloques: bloquesDe({ m, xml: t, rels: relsP, numerador: null }, raiz, null) };
      },
      // Un encabezado o un pie visto como un documento aparte (para editarlo con EdicionDoc y mostrarlo con MaquetadorDoc): comparte el paquete, los estilos
      // y las fuentes, pero tiene su propio XML, sus bloques y una sola «sección» del ancho de la hoja de `secIdx`. Su alto lo ajusta quien lo muestra.
      async editarParte(destino, secIdx) {
        const p = await modelo.parte(destino);
        if (!p) return null;
        const base = modelo.secciones[Math.min(Math.max(secIdx || 0, 0), modelo.secciones.length - 1)];
        const cfg = { ...base, arriba: 8, abajo: 8, encabezado: 0, pie: 0, continua: false, columnas: null, portada: false, inicioPaginas: undefined, referencias: [] };
        const vista = Object.create(modelo);
        const flota = (l, sal) => { for (const b of l) { if (b.tipo === 'p') sal.push(b); else for (const f of b.filas) for (const ce of f.celdas) flota(ce.bloques, sal); } return sal; };
        Object.assign(vista, {
          ruta: p.ruta, xml: p.xml, rels: p.rels, bloques: p.bloques, notas: new Map(), seccionFinal: cfg,
          secciones: [{ ...cfg, hasta: p.bloques.length - 1 }],
          parrafos: () => flota(vista.bloques, []),
          renumerar: () => [],
          recalcularSecciones() { vista.secciones[0].hasta = vista.bloques.length - 1; },
        });
        return vista;
      },
    };
    return modelo;
  }

  globalThis.DocxModelo = { abrir, aArbol };
})();

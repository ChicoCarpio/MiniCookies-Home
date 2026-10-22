/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Maquetador de Word propio (reemplaza a docx-preview). Convierte un .docx en HTML
// con estilos en línea: un <section class="docx"> por sección de Word, con el
// tamaño de la hoja como ancho y alto mínimo (border-box) y los márgenes como
// relleno. No corta líneas ni páginas: el navegador calcula el diseño y
// html-a-pdf.js lo mide, lo pagina y escribe el PDF (ver word-a-pdf.js).
//
// Lo que sale del HTML tiene que sobrevivir a seguridad-html.js (sin <style> ni
// clases con CSS) y a html-a-pdf.js (que solo ve palabras, imágenes, fondos y
// bordes reales: nada de ::before; la numeración y las guías de tabulación son
// texto de verdad).
//
// Requiere DocxBase, DocxEstilos, ZipPropio, XmlLigero. Para las fuentes con los
// anchos de Office usa PptxTexto (pptx-texto.js, que ya las tiene) si está.
// Las tabulaciones se calculan después de maquetar, midiendo en el navegador.
//
// Uso: const { html, mensajes } = await DocxMaquetador.aHtml(datos);

(function () {
  'use strict';
  const { hijo, lista, abrirPaquete } = DocxBase;
  const PX_POR_PT = 96 / 72;

  const RESALTADOS = {
    yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000', darkBlue: '#000080',
    darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080',
    lightGray: '#c0c0c0', black: '#000000', white: '#ffffff',
  };
  const TIPOS_IMAGEN = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp' };
  // Qué parte del espacio de más de una línea pone Word por encima del texto (CSS pone la mitad). Medido contra Word:
  // interlineado exacto 0,8; "al menos" 1 (todo arriba); múltiplo (sencillo, 1,5, doble) 0 (todo abajo).
  const PARTE_ARRIBA = { exact: 0.8, atLeast: 1, auto: 0 };
  const GUIAS = { dot: '.', hyphen: '-', underscore: '_', heavy: '_', middleDot: '·' };

  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const n3 = (x) => String(Math.round(x * 1000) / 1000);
  const pt = (x) => n3(x) + 'pt';
  const css = (o) => Object.keys(o).filter((k) => o[k] !== undefined && o[k] !== null && o[k] !== '').map((k) => `${k}:${o[k]}`).join(';');
  const hrefSeguro = (h) => (/^\s*(javascript|vbscript|data):/i.test(h || '') ? '' : h || '');

  function base64(bytes) {
    let s = '';
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  }

  // ───────── Fuentes ─────────
  const razones = new Map();
  // Alto natural de una línea de esa fuente, en veces el tamaño (Calibri 1,22; Arial 1,15…).
  function razonLinea(familiaCss) {
    if (razones.has(familiaCss)) return razones.get(familiaCss);
    let razon = 1.2;
    try {
      const c = document.createElement('canvas').getContext('2d');
      c.font = `100px ${familiaCss}`;
      const m = c.measureText('Hg');
      if (m.fontBoundingBoxAscent + m.fontBoundingBoxDescent > 0) razon = (m.fontBoundingBoxAscent + m.fontBoundingBoxDescent) / 100;
    } catch (e) { /* sin lienzo (Node): queda el valor de respaldo */ }
    razones.set(familiaCss, razon);
    return razon;
  }
  const SERIF = /cambria|georgia|garamond|palatino|bookman|century|constantia|baskerville|times|book antiqua|rockwell/i;
  const MONO = /consolas|courier|lucida console|mono|cascadia/i;
  function familiaCss(nombre) {
    const n = nombre || 'Times New Roman';
    const base = typeof PptxTexto !== 'undefined' && PptxTexto.familia ? PptxTexto.familia(n) : `"${n}"`;
    // Comillas simples: la familia va dentro de un atributo style="…".
    return (base.replace(/, "Segoe UI", sans-serif$/, '') + (MONO.test(n) ? ', monospace' : SERIF.test(n) ? ', serif' : ', "Segoe UI", sans-serif')).replace(/"/g, "'");
  }

  // ───────── Contexto del documento ─────────
  class Documento {
    constructor(paquete, modelo, mensajes) {
      this.zip = paquete.zip; this.rels = paquete.rels; this.m = modelo; this.mensajes = mensajes;
      this.numerador = modelo.numeracion.numerador();
      this.fuentes = new Map();
      this.imagenes = new Map();
      this.sinNombre = 0;
    }
    usarFuente(nombre, negrita, cursiva) {
      const k = `${nombre}|${!!negrita}|${!!cursiva}`;
      if (!this.fuentes.has(k)) this.fuentes.set(k, { fuente: nombre, negrita: !!negrita, cursiva: !!cursiva });
    }
    aviso(texto) { if (!this.mensajes.some((x) => x.message === texto)) this.mensajes.push({ type: 'warning', message: texto }); }

    async imagenDe(idRel) {
      const rel = this.rels.get(idRel);
      if (!rel || rel.externo) { this.aviso('Hay una imagen vinculada que no se puede incluir.'); return null; }
      if (this.imagenes.has(rel.destino)) return this.imagenes.get(rel.destino);
      const f = this.zip.file(rel.destino);
      let src = null;
      if (f) {
        const tipo = TIPOS_IMAGEN[(rel.destino.split('.').pop() || '').toLowerCase()];
        if (tipo) src = `data:${tipo};base64,${base64(await f.async('uint8array'))}`;
        else this.aviso('Hay una imagen en un formato que el navegador no muestra.');
      }
      this.imagenes.set(rel.destino, src);
      return src;
    }

    // Formato efectivo de un trozo de texto: base del párrafo + estilo de carácter + formato directo.
    formatoRun(baseR, rPr) {
      const directo = this.m.leerRPr(rPr);
      return this.m.mezclar(this.m.mezclar(baseR, this.m.estilos.caracter(directo.estilo)), directo);
    }
  }

  // CSS de un trozo de texto (todo lo que Word llama formato de carácter).
  function cssRun(d, r, sinFuente, bajar) {
    const nombreFuente = d.m.fuente(r.fuentes) || null;
    const tam = r.tam !== undefined ? r.tam : 10;
    const sup = r.alineacionVertical === 'superscript', sub = r.alineacionVertical === 'subscript';
    d.usarFuente(nombreFuente || 'Times New Roman', r.b, r.i);
    const deco = [];
    if (r.u && r.u !== 'none') deco.push('underline');
    if (r.tachado || r.tachadoDoble) deco.push('line-through');
    const estiloU = /double/.test(r.u || '') ? 'double' : /dotted|dotDash/.test(r.u || '') ? 'dotted' : /dash/i.test(r.u || '') ? 'dashed' : /wave/.test(r.u || '') ? 'wavy' : undefined;
    const fondo = r.resaltado && RESALTADOS[r.resaltado] ? RESALTADOS[r.resaltado]
      : r.sombreado && r.sombreado.relleno && r.sombreado.relleno !== 'auto' ? '#' + r.sombreado.relleno : undefined;
    return css({
      'font-family': sinFuente ? undefined : familiaCss(nombreFuente),
      'font-size': pt(sup || sub ? tam * 0.65 : tam),
      'font-weight': r.b ? 'bold' : undefined, 'font-style': r.i ? 'italic' : undefined,
      color: d.m.color(r.color) || undefined,
      'background-color': fondo,
      'text-decoration-line': deco.length ? deco.join(' ') : undefined,
      'text-decoration-style': deco.length ? estiloU : undefined,
      'text-transform': r.mayusculas ? 'uppercase' : undefined,
      'font-variant': r.versalitas ? 'small-caps' : undefined,
      'vertical-align': sup ? 'super' : sub ? 'sub' : undefined,
      'letter-spacing': r.espaciado ? pt(r.espaciado) : undefined,
      position: r.posicion || bajar ? 'relative' : undefined, top: r.posicion || bajar ? pt((bajar || 0) - (r.posicion || 0)) : undefined,
      display: r.oculto ? 'none' : undefined,
    });
  }

  // ───────── Contenido en línea de un párrafo ─────────
  // Devuelve una lista de elementos: { t:'texto', s, r } | { t:'tab', r } | { t:'br', r } | { t:'pagina' } | { t:'html', html, r }
  async function leerInline(d, nodo, baseR, salida, contexto) {
    for (const e of nodo.h) {
      switch (e.n) {
        case 'r': await leerRun(d, e, baseR, salida, contexto); break;
        case 'hyperlink': {
          const ini = salida.length;
          await leerInline(d, e, baseR, salida, contexto);
          const rel = e.a.id ? d.rels.get(e.a.id) : null;
          const href = e.a.anchor ? '#' + e.a.anchor : rel ? hrefSeguro(rel.destino) : '';
          if (href) for (let i = ini; i < salida.length; i++) salida[i].enlace = href;
          break;
        }
        case 'fldSimple': await leerInline(d, e, baseR, salida, contexto); break;
        case 'smartTag': case 'sdt': case 'sdtContent': case 'ins': case 'customXml': case 'bdo': case 'dir':
          await leerInline(d, e, baseR, salida, contexto); break;
        case 'bookmarkStart': if (e.a.name && !/^_/.test(e.a.name)) salida.push({ t: 'ancla', nombre: e.a.name }); break;
        case 'AlternateContent': {
          const dentro = hijo(e, 'Choice') || hijo(e, 'Fallback');
          if (dentro) await leerInline(d, dentro, baseR, salida, contexto);
          break;
        }
        default: break; // pPr, proofErr, del, comentarios, marcadores de revisión…
      }
    }
  }

  async function leerRun(d, run, baseR, salida, contexto) {
    const r = d.formatoRun(baseR, hijo(run, 'rPr'));
    for (const e of run.h) {
      switch (e.n) {
        case 't': if (!contexto.enInstruccion) salida.push({ t: 'texto', s: e.t, r }); break;
        case 'tab': if (!contexto.enInstruccion) salida.push({ t: 'tab', r }); break;
        case 'br': case 'cr':
          if (e.a.type === 'page') salida.push({ t: 'pagina' });
          else if (e.a.type === 'column') salida.push({ t: 'columna' });
          else salida.push({ t: 'br', r, limpiar: e.a.clear });
          break;
        case 'noBreakHyphen': salida.push({ t: 'texto', s: '‑', r }); break;
        case 'softHyphen': salida.push({ t: 'texto', s: '­', r }); break;
        case 'fldChar':
          if (e.a.fldCharType === 'begin') contexto.campos.push({ fase: 'instr' });
          else if (e.a.fldCharType === 'separate') { const c = contexto.campos[contexto.campos.length - 1]; if (c) c.fase = 'resultado'; }
          else if (e.a.fldCharType === 'end') contexto.campos.pop();
          contexto.enInstruccion = contexto.campos.some((c) => c.fase === 'instr');
          break;
        case 'drawing': case 'AlternateContent': {
          const img = await leerDibujo(d, e, r);
          if (img) salida.push(img);
          break;
        }
        default: break; // instrText, delText, lastRenderedPageBreak, pict, footnoteReference (fase 5)…
      }
    }
  }

  const EMU_POR_PT = 12700;
  // Fase 2: solo imágenes en línea. Las flotantes (wp:anchor) llegan en la fase 4.
  async function leerDibujo(d, e, r) {
    const raiz = e.n === 'AlternateContent' ? (hijo(e, 'Choice') || hijo(e, 'Fallback')) : e;
    const dib = raiz && (hijo(raiz, 'drawing') || raiz);
    const inl = dib && hijo(dib, 'inline');
    if (!inl) return null;
    const ext = hijo(inl, 'extent');
    const blip = buscarProf(inl, 'blip');
    if (!ext || !blip || !blip.a.embed) return null;
    const src = await d.imagenDe(blip.a.embed);
    if (!src) return null;
    const doc = hijo(inl, 'docPr');
    const ancho = parseFloat(ext.a.cx) / EMU_POR_PT, alto = parseFloat(ext.a.cy) / EMU_POR_PT;
    return { t: 'html', r, html: `<img src="${esc(src)}" alt="${esc(doc ? doc.a.descr || '' : '')}" style="${css({ width: pt(ancho), height: pt(alto), 'vertical-align': 'baseline', 'max-width': 'none' })}">` };
  }
  function buscarProf(e, nombre) {
    for (const x of e.h) {
      if (x.n === nombre) return x;
      const r = buscarProf(x, nombre);
      if (r) return r;
    }
    return null;
  }

  // ───────── Párrafos ─────────
  const ALINEACION = { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify', justify: 'justify' };

  function cssBorde(d, b) {
    if (!b || !b.tipo || b.tipo === 'nil' || b.tipo === 'none') return null;
    const ancho = b.ancho === undefined ? 0.5 : Math.max(b.ancho, 0.25);
    const estilo = /double/i.test(b.tipo) ? 'double' : /dotted/i.test(b.tipo) ? 'dotted' : /dash/i.test(b.tipo) ? 'dashed' : 'solid';
    return { texto: `${pt(ancho)} ${estilo} ${d.m.color(b.color) || '#000000'}`, ancho };
  }

  function agruparRuns(items) {
    // Une elementos de texto seguidos con el mismo formato en un solo trozo.
    const sal = [];
    for (const it of items) {
      const ult = sal[sal.length - 1];
      if (it.t === 'texto' && ult && ult.t === 'texto' && ult.r === it.r && ult.enlace === it.enlace) ult.s += it.s;
      else sal.push({ ...it });
    }
    return sal;
  }

  function htmlInline(d, items, esDefecto, bajar) {
    let html = '';
    let anclaEnlace = null;
    const cerrarEnlace = () => { if (anclaEnlace) { html += '</a>'; anclaEnlace = null; } };
    for (const it of agruparRuns(items)) {
      if (it.t === 'ancla') { html += `<a id="${esc(it.nombre)}"></a>`; continue; }
      if (it.enlace !== anclaEnlace) { cerrarEnlace(); if (it.enlace) { html += `<a href="${esc(it.enlace)}" style="color:inherit">`; anclaEnlace = it.enlace; } }
      if (it.t === 'texto') html += `<span style="${cssRun(d, it.r, esDefecto, bajar)}">${esc(it.s)}</span>`;
      else if (it.t === 'tab') html += `<span data-tab="1" style="${cssRun(d, it.r, false)};display:inline-block;white-space:pre">\t</span>`;
      else if (it.t === 'br') html += `<br${it.limpiar ? ` style="clear:${it.limpiar === 'all' ? 'both' : it.limpiar}"` : ''}>`;
      else if (it.t === 'html') html += it.html;
    }
    cerrarEnlace();
    return html;
  }

  // El PDF solo puede escribir las viñetas que están en WinAnsi o en Noto Sans (html-a-pdf.js): el resto
  // (●, ■, ➢…) saldría en blanco. Los círculos vacíos quedan como "o"; los demás, como "•".
  const simboloEscribible = (s) => (/^[○◦□▫o]$/.test(s) ? 'o' : '•');

  // Un párrafo de Word → uno o varios <p> (un salto de página parte el párrafo en dos).
  async function parrafo(d, p) {
    const m = d.m;
    const pDirecto = m.leerPPr(hijo(p, 'pPr'));
    const est = m.estilos.parrafo(pDirecto.estilo);
    const numeracionPedida = pDirecto.numeracion || est.p.numeracion;
    let nivel = null;
    let etiqueta = null;
    let efectivo = m.mezclar(est.p, null);
    if (numeracionPedida && numeracionPedida.numId && numeracionPedida.numId !== '0') {
      const e = d.numerador.siguiente(numeracionPedida.numId, numeracionPedida.nivel || '0');
      if (e) { nivel = e.nivel; etiqueta = nivel.fmt === 'bullet' ? simboloEscribible(e.texto) : e.texto; efectivo = m.mezclar(efectivo, nivel.p); }
    }
    efectivo = m.mezclar(efectivo, pDirecto);
    // El formato de la marca de párrafo vale para el párrafo vacío y para la etiqueta de la lista.
    const baseR = est.r;

    const items = [];
    const contexto = { campos: [], enInstruccion: false };
    await leerInline(d, p, baseR, items, contexto);
    const marcaR = m.mezclar(baseR, efectivo.marca);

    const sangria = efectivo.sangria || {};
    const izq = sangria.izq || 0, der = sangria.der || 0;
    const colgante = sangria.colgante || 0, primera = colgante ? -colgante : sangria.primera || 0;
    const esp = efectivo.espacio || {};
    const antes = esp.antesAuto ? 14 : esp.antes || 0, despues = esp.despuesAuto ? 14 : esp.despues || 0;

    // Etiqueta de la lista, como texto de verdad, seguida de tabulación, espacio o nada.
    if (etiqueta !== null) {
      const rEt = m.mezclar(marcaR, nivel.r);
      if (nivel.fmt === 'bullet') delete rEt.fuentes; // Symbol/Wingdings: el símbolo ya se cambió por uno que el PDF escribe
      const aj = [];
      if (etiqueta) aj.push({ t: 'texto', s: etiqueta, r: rEt, esEtiqueta: true });
      if ((nivel.sufijo || 'tab') === 'tab') aj.push({ t: 'tab', r: rEt });
      else if (nivel.sufijo === 'space') aj.push({ t: 'texto', s: ' ', r: rEt });
      items.unshift(...aj);
    }

    // Fuente que fija la altura de la línea: la del primer trozo de texto (o la de la marca si no hay).
    const rTextos = items.filter((x) => x.t === 'texto' || x.t === 'tab').map((x) => x.r);
    const rBase = rTextos[0] || marcaR;
    const fuenteBase = m.fuente(rBase.fuentes) || null;
    const tamMax = rTextos.reduce((mx, r) => Math.max(mx, r.tam !== undefined ? r.tam : 10), rBase.tam !== undefined ? rBase.tam : 10);
    const famBase = familiaCss(fuenteBase);
    const natural = razonLinea(famBase);
    let lineHeight, altoLinea;
    const linea = esp.linea;
    if (linea !== undefined) {
      if (esp.regla === 'exact') altoLinea = linea / 20;
      else if (esp.regla === 'atLeast') altoLinea = Math.max(linea / 20, natural * tamMax);
      else if (Math.abs(linea - 240) > 1) altoLinea = natural * tamMax * linea / 240;
    }
    if (altoLinea !== undefined) lineHeight = pt(altoLinea);
    // Word pone el espacio de más de la línea arriba del texto; CSS lo reparte mitad y mitad:
    // se mueve el texto lo que haga falta (puede ser hacia arriba).
    const bajar = altoLinea !== undefined ? (altoLinea - natural * tamMax) * (PARTE_ARRIBA[esp.regla] ?? PARTE_ARRIBA.auto) - (altoLinea - natural * tamMax) / 2 : 0;

    // Bordes y sombreado del párrafo.
    const b = efectivo.bordes || {};
    const bordes = {};
    for (const lado of ['top', 'right', 'bottom', 'left']) bordes[lado] = cssBorde(d, b[lado]);
    const rellenoP = efectivo.sombreado && efectivo.sombreado.relleno && efectivo.sombreado.relleno !== 'auto' ? '#' + efectivo.sombreado.relleno : undefined;
    const espacioLado = (lado) => (b[lado] && b[lado].espacio) || 0;
    const margenIzq = izq - (bordes.left ? espacioLado('left') + bordes.left.ancho : 0);
    const margenDer = der - (bordes.right ? espacioLado('right') + bordes.right.ancho : 0);

    const estilo = (extra) => css({
      margin: `${pt(antes)} ${pt(Math.max(margenDer, 0))} ${pt(despues)} ${pt(Math.max(margenIzq, 0))}`,
      'text-indent': primera ? pt(primera) : undefined,
      'text-align': ALINEACION[efectivo.jc] || 'left',
      'text-align-last': efectivo.jc === 'distribute' ? 'justify' : undefined,
      'white-space': 'pre-wrap', 'overflow-wrap': 'break-word',
      'font-family': famBase, 'font-size': pt(tamMax), 'line-height': lineHeight,
      'padding-top': bordes.top ? pt(espacioLado('top')) : undefined, 'padding-bottom': bordes.bottom ? pt(espacioLado('bottom')) : undefined,
      'padding-left': bordes.left ? pt(espacioLado('left')) : undefined, 'padding-right': bordes.right ? pt(espacioLado('right')) : undefined,
      'border-top': bordes.top && bordes.top.texto, 'border-bottom': bordes.bottom && bordes.bottom.texto,
      'border-left': bordes.left && bordes.left.texto, 'border-right': bordes.right && bordes.right.texto,
      'background-color': rellenoP,
      direction: efectivo.derechaAIzquierda ? 'rtl' : undefined,
      'break-before': efectivo.saltoAntes ? 'page' : undefined,
      'break-inside': efectivo.mantenerLineas ? 'avoid' : undefined,
      ...extra,
    });
    const atributos = `data-izq="${n3(izq)}" data-der="${n3(der)}" data-tabs='${esc(JSON.stringify(efectivo.tabs || []))}'`;
    d.usarFuente(fuenteBase || 'Times New Roman', rBase.b, rBase.i);

    // Parte en trozos cada salto de página.
    const trozos = [[]];
    for (const it of items) { if (it.t === 'pagina') trozos.push([]); else trozos[trozos.length - 1].push(it); }
    let html = '';
    trozos.forEach((trozo, i) => {
      const interior = htmlInline(d, trozo, false, bajar) || `<span style="${cssRun(d, marcaR, true)}"><br></span>`;
      html += `<p ${atributos} style="${estilo(i > 0 ? { 'break-before': 'page', 'text-indent': undefined } : {})}">${interior}</p>`;
    });
    const sectPr = hijo(hijo(p, 'pPr'), 'sectPr');
    return { html, sectPr };
  }

  // ───────── Tablas (versión mínima; la fase 3 la completa) ─────────
  async function tabla(d, t) {
    const m = d.m;
    const tblPr = m.leerTblPr(hijo(t, 'tblPr'));
    const grid = lista(hijo(t, 'tblGrid'), 'gridCol').map((c) => parseFloat(c.a.w) / 20);
    const filas = lista(t, 'tr');
    const bordesT = tblPr.bordes || {};
    const margenesT = tblPr.margenes || {};
    const celdasFila = filas.map((f) => lista(f, 'tc'));
    // Posición de columna de cada celda y alto (rowspan) de las combinadas verticalmente.
    const infos = celdasFila.map((celdas) => {
      let col = 0;
      return celdas.map((c) => {
        const pr = m.leerTcPr(hijo(c, 'tcPr'));
        const span = pr.columnas || 1;
        const info = { c, pr, col, span };
        col += span;
        return info;
      });
    });
    infos.forEach((fila, i) => fila.forEach((info) => {
      info.filas = 1;
      if (info.pr.combinadaVertical === 'restart') {
        for (let k = i + 1; k < infos.length; k++) {
          const sig = infos[k].find((x) => x.col === info.col);
          if (sig && sig.pr.combinadaVertical === 'continue') { sig.omitir = true; info.filas++; } else break;
        }
      }
    }));
    let html = `<table style="${css({ 'border-collapse': 'collapse', 'table-layout': 'fixed', width: pt(grid.reduce((a, b) => a + b, 0)), margin: '0' })}"><colgroup>${grid.map((g) => `<col style="width:${pt(g)}">`).join('')}</colgroup>`;
    for (const fila of infos) {
      html += '<tr>';
      for (const info of fila) {
        if (info.omitir) continue;
        const bc = info.pr.bordes || {};
        const lado = (nombre, interior) => cssBorde(d, bc[nombre] || bordesT[nombre] || bordesT[interior]);
        const bt = lado('top', 'insideH'), bb = lado('bottom', 'insideH'), bl = lado('left', 'insideV'), br = lado('right', 'insideV');
        const relleno = info.pr.sombreado && info.pr.sombreado.relleno && info.pr.sombreado.relleno !== 'auto' ? '#' + info.pr.sombreado.relleno : undefined;
        const mc = { ...{ top: 0, bottom: 0, left: 5.4, right: 5.4 }, ...margenesT, ...(info.pr.margenes || {}) };
        const dentro = await bloques(d, info.c);
        html += `<td${info.span > 1 ? ` colspan="${info.span}"` : ''}${info.filas > 1 ? ` rowspan="${info.filas}"` : ''} style="${css({
          'vertical-align': info.pr.alineacionVertical === 'center' ? 'middle' : info.pr.alineacionVertical === 'bottom' ? 'bottom' : 'top',
          padding: `${pt(mc.top)} ${pt(mc.right)} ${pt(mc.bottom)} ${pt(mc.left)}`,
          'border-top': bt && bt.texto, 'border-bottom': bb && bb.texto, 'border-left': bl && bl.texto, 'border-right': br && br.texto,
          'background-color': relleno, 'overflow-wrap': 'break-word',
        })}"><div style="display:flex;flex-direction:column">${dentro.html}</div></td>`;
      }
      html += '</tr>';
    }
    return `${html}</table>`;
  }

  // ───────── Bloques y secciones ─────────
  // Bloques de una celda (o de cualquier contenedor sin secciones).
  async function bloques(d, contenedor) {
    let html = '';
    for (const e of contenedor.h) {
      if (e.n === 'p') html += (await parrafo(d, e)).html;
      else if (e.n === 'tbl') html += await tabla(d, e);
      else if (e.n === 'sdt') { const c = hijo(e, 'sdtContent'); if (c) html += (await bloques(d, c)).html; }
    }
    return { html };
  }

  function leerSeccion(m, sectPr) {
    const sz = sectPr && hijo(sectPr, 'pgSz');
    const mg = sectPr && hijo(sectPr, 'pgMar');
    const v = (x, defecto) => { const n = parseFloat(x); return Number.isFinite(n) ? n / 20 : defecto; };
    return {
      ancho: v(sz && sz.a.w, 612), alto: v(sz && sz.a.h, 792),
      arriba: v(mg && mg.a.top, 72), derecha: v(mg && mg.a.right, 72), abajo: v(mg && mg.a.bottom, 72), izquierda: v(mg && mg.a.left, 72),
      continua: !!(sectPr && hijo(sectPr, 'type') && hijo(sectPr, 'type').a.val === 'continuous'),
    };
  }

  function htmlSeccion(s, interior) {
    return `<section class="docx" style="${css({
      'box-sizing': 'border-box', display: 'block', position: 'relative', width: pt(s.ancho), 'min-height': pt(s.alto),
      padding: `${pt(s.arriba)} ${pt(s.derecha)} ${pt(s.abajo)} ${pt(s.izquierda)}`, margin: '0', background: '#ffffff', color: '#000000',
    })}"><div style="display:flex;flex-direction:column">${interior}</div></section>`;
  }

  // ───────── Tabulaciones (después de maquetar) ─────────
  async function resolverTabulaciones(html, d) {
    if (typeof document === 'undefined' || html.indexOf('data-tab') < 0) return html;
    const lienzo = document.createElement('div');
    lienzo.style.cssText = 'position:absolute;left:-100000px;top:0;visibility:hidden';
    lienzo.innerHTML = html;
    document.body.appendChild(lienzo);
    try {
      if (document.fonts && document.fonts.ready) await document.fonts.ready;
      const tabDefecto = d.m.ajustes.tabPorDefecto;
      const ctx = document.createElement('canvas').getContext('2d');
      for (const p of lienzo.querySelectorAll('p[data-izq]')) {
        const tabs = p.querySelectorAll('[data-tab]');
        if (!tabs.length) continue;
        const izq = parseFloat(p.dataset.izq) * PX_POR_PT;
        const e = getComputedStyle(p);
        let personalizadas = [];
        try { personalizadas = JSON.parse(p.dataset.tabs || '[]'); } catch (x) { personalizadas = []; }
        const paradas = personalizadas.filter((t) => t.tipo !== 'clear' && t.pos !== undefined).map((t) => ({ ...t, px: t.pos * PX_POR_PT }));
        const sangria = parseFloat(p.style.textIndent) || 0;
        if (sangria < 0) paradas.push({ tipo: 'left', px: izq }); // la sangría francesa también es una parada
        for (const tab of tabs) {
          const rp = p.getBoundingClientRect();
          const origen = rp.left + parseFloat(e.borderLeftWidth) + parseFloat(e.paddingLeft) - izq;
          tab.style.width = '0px'; tab.textContent = '​';
          const x0 = tab.getBoundingClientRect().left - origen;
          const parada = paradas.filter((s) => s.px > x0 + 0.5).sort((a, b) => a.px - b.px)[0];
          const paso = tabDefecto * PX_POR_PT;
          const porDefecto = (Math.floor((x0 + 0.5) / paso) + 1) * paso;
          // Las paradas por defecto solo valen después de la última parada propia.
          const meta = parada || { tipo: 'left', px: porDefecto };
          let ancho = meta.px - x0;
          if (meta.tipo === 'right' || meta.tipo === 'center' || meta.tipo === 'decimal') {
            let siguiente = 0;
            for (let n = tab.nextSibling; n; n = n.nextSibling) {
              if (n.nodeType === 1 && n.hasAttribute('data-tab')) break;
              if (n.nodeType === 1 && n.tagName === 'BR') break;
              siguiente += n.nodeType === 1 ? n.getBoundingClientRect().width : 0;
            }
            ancho -= meta.tipo === 'center' ? siguiente / 2 : siguiente;
          }
          ancho = Math.max(ancho, 0);
          const caracter = GUIAS[meta.guia];
          tab.style.width = n3(ancho) + 'px';
          if (caracter && ancho > 0) {
            const ec = getComputedStyle(tab);
            ctx.font = `${ec.fontStyle} ${ec.fontWeight} ${ec.fontSize} ${ec.fontFamily}`;
            const w = ctx.measureText(caracter).width || 1;
            tab.textContent = caracter.repeat(Math.max(Math.floor(ancho / w) - 0, 0));
          }
        }
      }
      return lienzo.innerHTML;
    } finally {
      lienzo.remove();
    }
  }

  // ───────── Entrada ─────────
  globalThis.DocxMaquetador = { aHtml };

  async function aHtml(datos, opciones) {
    const mensajes = [];
    const paquete = await abrirPaquete(datos);
    const modelo = await DocxEstilos.cargar(paquete);
    const d = new Documento(paquete, modelo, mensajes);
    const lectura = await bloquesCuerpo(d, paquete.cuerpo);
    if (typeof PptxTexto !== 'undefined' && PptxTexto.cargarFuentes) await PptxTexto.cargarFuentes([...d.fuentes.values()]);
    const secciones = [];
    for (const t of lectura) {
      const cfg = leerSeccion(modelo, t.sectPr);
      const previa = secciones[secciones.length - 1];
      if (previa && cfg.continua && previa.cfg.ancho === cfg.ancho && previa.cfg.alto === cfg.alto) previa.html += t.html;
      else secciones.push({ cfg, html: t.html });
    }
    const html = await resolverTabulaciones(secciones.map((s) => htmlSeccion(s.cfg, s.html)).join(''), d);
    return { html, mensajes };
  }

  // Recorre el cuerpo y lo reparte en secciones: cada párrafo con sectPr cierra una; la última es la del cuerpo.
  async function bloquesCuerpo(d, cuerpo) {
    const sal = [];
    let actual = '';
    const procesar = async (contenedor) => {
      for (const e of contenedor.h) {
        if (e.n === 'p') {
          const r = await parrafo(d, e);
          actual += r.html;
          if (r.sectPr) { sal.push({ html: actual, sectPr: r.sectPr }); actual = ''; }
        } else if (e.n === 'tbl') actual += await tabla(d, e);
        else if (e.n === 'sdt') { const c = hijo(e, 'sdtContent'); if (c) await procesar(c); }
        else if (e.n === 'customXml' || e.n === 'ins') await procesar(e);
      }
    };
    await procesar(cuerpo);
    sal.push({ html: actual, sectPr: lista(cuerpo, 'sectPr').pop() || null });
    // Si el cuerpo termina con un párrafo que ya cerró su sección y no queda nada más, esa última sección vacía sobra.
    if (sal.length > 1 && !sal[sal.length - 1].html && !lista(cuerpo, 'sectPr').length) sal.pop();
    return sal;
  }
})();

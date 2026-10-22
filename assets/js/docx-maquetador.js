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
  const { hijo, lista, buscar, base64, hrefSeguro, abrirPaquete } = DocxBase;
  const PX_POR_PT = 96 / 72;
  // Tablas de consulta sin prototipo: un valor del archivo como "constructor" o "__proto__" no encuentra nada.
  const sinProto = (o) => Object.assign(Object.create(null), o);

  const RESALTADOS = sinProto({
    yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000', darkBlue: '#000080',
    darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080',
    lightGray: '#c0c0c0', black: '#000000', white: '#ffffff',
  });
  const TIPOS_IMAGEN = sinProto({ png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp' });
  // Qué parte del espacio de más de una línea pone Word por encima del texto (CSS pone la mitad). Medido contra Word:
  // interlineado exacto 0,8; "al menos" 1 (todo arriba); múltiplo (sencillo, 1,5, doble) 0 (todo abajo).
  const PARTE_ARRIBA = sinProto({ exact: 0.8, atLeast: 1, auto: 0 });
  const GUIAS = sinProto({ dot: '.', hyphen: '-', underscore: '_', heavy: '_', middleDot: '·' });

  // Todo lo que viene del XML y va a un atributo pasa por esc (también la comilla simple: data-tabs va entre '…').
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  const n3 = (x) => String(Math.round(x * 1000) / 1000);
  // Toda medida del CSS sale de aquí: sin NaN ni Infinity, y sin pasar del doble de la hoja más grande
  // (un espacio antes de 10^9 twips serían miles de hojas en blanco; Word no deja más de 1584 pt).
  const pt = (x) => n3(Number.isFinite(x) ? Math.min(Math.max(x, -2 * MAX_PT), 2 * MAX_PT) : 0) + 'pt';
  const css = (o) => Object.keys(o).filter((k) => o[k] !== undefined && o[k] !== null && o[k] !== '').map((k) => `${k}:${o[k]}`).join(';');
  // Un relleno (w:shd w:fill) va a style="…": solo un hexadecimal de 6 cifras; "auto" u otra cosa, sin relleno.
  const rellenoCss = (sh) => (sh && /^[0-9a-fA-F]{6}$/.test(sh.relleno || '') ? '#' + sh.relleno.toLowerCase() : undefined);
  // Tope de medidas que vienen del archivo: 22 pulgadas, lo más grande que Word deja poner a una hoja o un dibujo.
  const MAX_PT = 1584;
  const acotar = (x, min, max) => (Number.isFinite(x) ? Math.min(Math.max(x, min), max) : min);

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
    // El nombre viene del archivo y va dentro de font-family:'…': sin comillas, barras ni ";" que cierren la cadena.
    const n = String(nombre || '').replace(/[^\p{L}\p{M}\p{N} ._&+-]/gu, '').trim() || 'Times New Roman';
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
      this.notas = [];          // notas al pie y finales citadas en el texto: { id, fin }
      this.numeroNotaActual = 0;
      this.nParrafo = 0;  // párrafos leídos hasta ahora
      this.flotIzq = null; // último dibujo flotante a la izquierda: { borde (pt desde el margen), n (párrafo) }
      this.cfg = null; // geometría de la sección que se está leyendo (para los dibujos flotantes)
    }
    usarFuente(nombre, negrita, cursiva) {
      const k = `${nombre}|${!!negrita}|${!!cursiva}`;
      if (!this.fuentes.has(k)) this.fuentes.set(k, { fuente: nombre, negrita: !!negrita, cursiva: !!cursiva });
    }
    // Registra una nota citada en el texto y devuelve su número (se numeran en orden de aparición).
    notaNueva(id, fin) {
      if (id === undefined || id === '-1' || id === '0') return 0;
      const clave = (fin ? 'e' : 'f') + id;
      let i = this.notas.findIndex((n) => n.clave === clave);
      if (i < 0) { this.notas.push({ clave, id, fin }); i = this.notas.length - 1; }
      return i + 1;
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
    const fondo = r.resaltado && RESALTADOS[r.resaltado] ? RESALTADOS[r.resaltado] : rellenoCss(r.sombreado);
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
        case 'fldSimple': {
          const nombre = /^\s*(PAGE|NUMPAGES)\b/i.exec(e.a.instr || '');
          if (nombre) salida.push({ t: 'campo', nombre: nombre[1].toUpperCase(), r: d.formatoRun(baseR, hijo(hijo(e, 'r'), 'rPr')) });
          else await leerInline(d, e, baseR, salida, contexto);
          break;
        }
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
          if (e.a.fldCharType === 'begin') contexto.campos.push({ fase: 'instr', instr: '' });
          else if (e.a.fldCharType === 'separate') {
            const c = contexto.campos[contexto.campos.length - 1];
            if (c) {
              c.fase = 'resultado';
              // Número de página y total: se escriben al dibujar cada hoja (html-a-pdf.js); el resultado guardado no vale.
              const nombre = /^\s*(PAGE|NUMPAGES)\b/i.exec(c.instr);
              if (nombre) { c.campo = true; salida.push({ t: 'campo', nombre: nombre[1].toUpperCase(), r }); }
            }
          } else if (e.a.fldCharType === 'end') contexto.campos.pop();
          contexto.enInstruccion = contexto.campos.some((c) => c.fase === 'instr' || c.campo);
          break;
        case 'instrText': { const c = contexto.campos[contexto.campos.length - 1]; if (c) c.instr += e.t; break; }
        case 'footnoteReference': case 'endnoteReference': {
          const n = d.notaNueva(e.a.id, e.n === 'endnoteReference');
          if (n) salida.push({ t: 'texto', s: String(n), r: { ...r, alineacionVertical: 'superscript' } });
          break;
        }
        case 'footnoteRef': case 'endnoteRef':
          if (d.numeroNotaActual) salida.push({ t: 'texto', s: String(d.numeroNotaActual), r: { ...r, alineacionVertical: 'superscript' } });
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
  const emuPt = (v) => (v === undefined || v === '' || Number.isNaN(parseFloat(v)) ? 0 : parseFloat(v) / EMU_POR_PT);

  // Dibujos: imágenes y cuadros de texto, en línea (wp:inline) o flotantes (wp:anchor).
  async function leerDibujo(d, e, r) {
    const raiz = e.n === 'AlternateContent' ? (hijo(e, 'Choice') || hijo(e, 'Fallback')) : e;
    const dib = raiz && (hijo(raiz, 'drawing') || raiz);
    const inl = dib && hijo(dib, 'inline');
    const anc = dib && hijo(dib, 'anchor');
    const caja = inl || anc;
    const ext = caja && hijo(caja, 'extent');
    if (!ext) return null;
    const ancho = acotar(emuPt(ext.a.cx), 0, MAX_PT), alto = acotar(emuPt(ext.a.cy), 0, MAX_PT);
    const contenido = await contenidoDibujo(d, caja, ancho, alto, !!inl);
    if (!contenido) return null;
    // Un cuadro de texto en línea lleva <p> adentro: no cabe en un <p>, parrafo() lo deja entre dos trozos.
    if (inl) return contenido.startsWith('<div') ? { t: 'bloque', r, html: contenido } : { t: 'html', r, html: contenido };
    // Va aparte (antes del párrafo): un <div> o un <p> dentro de un <p> cierra el párrafo al leer el HTML.
    return { t: 'flotante', r, html: colocarAncla(d, anc, ancho, alto, contenido) };
  }

  // La imagen o el cuadro de texto, sin posición.
  async function contenidoDibujo(d, caja, ancho, alto, enLinea) {
    const doc = hijo(caja, 'docPr');
    const blip = buscar(caja, 'blip');
    if (blip && blip.a.embed) {
      const src = await d.imagenDe(blip.a.embed);
      if (!src) return null;
      return `<img src="${esc(src)}" alt="${esc(doc ? doc.a.descr || '' : '')}" style="${css({ width: pt(ancho), height: pt(alto), display: enLinea ? 'inline-block' : 'block', 'max-width': 'none', 'vertical-align': enLinea ? 'baseline' : undefined })}">`;
    }
    const wsp = buscar(caja, 'wsp');
    const txbx = wsp && buscar(wsp, 'txbxContent');
    if (txbx) {
      const sp = hijo(wsp, 'spPr');
      const bp = hijo(wsp, 'bodyPr');
      const ins = (n, defecto) => (bp && bp.a[n] !== undefined ? emuPt(bp.a[n]) : defecto);
      const ln = sp && hijo(sp, 'ln');
      const colorDe = (e) => { const c = e && hijo(e, 'solidFill'); const x = c && c.h[0]; return x && x.a.val && /^[0-9a-fA-F]{6}$/.test(x.a.val) ? '#' + x.a.val.toLowerCase() : undefined; };
      const relleno = colorDe(sp);
      const colorLinea = ln && colorDe(ln);
      const interior = await bloques(d, txbx, Math.max(ancho - ins('lIns', 7.2) - ins('rIns', 7.2), 10));
      // En línea (wp:inline) el cuadro va dentro del <p>: ahí no puede haber <p> ni <div> (cerrarían el párrafo al leer
      // el HTML), así que se escribe con <span> en bloque. Con una tabla adentro no se puede: queda entre dos trozos.
      const dentroDeP = enLinea && !/<table/.test(interior.html);
      const etiqueta = dentroDeP ? 'span' : 'div';
      const cuerpo = dentroDeP ? interior.html.replace(/<(p|div)\s([^>]*?)style="/g, '<span $2style="display:block;').replace(/<(p|div)>/g, '<span style="display:block">').replace(/<\/(p|div)>/g, '</span>') : interior.html;
      return `<${etiqueta} style="${css({
        display: dentroDeP ? 'inline-block' : undefined, 'vertical-align': dentroDeP ? 'bottom' : undefined,
        'box-sizing': 'border-box', width: pt(ancho), 'min-height': pt(alto), padding: `${pt(ins('tIns', 3.6))} ${pt(ins('rIns', 7.2))} ${pt(ins('bIns', 3.6))} ${pt(ins('lIns', 7.2))}`,
        'background-color': relleno, border: colorLinea ? `${pt(ln.a.w ? emuPt(ln.a.w) : 0.75)} solid ${colorLinea}` : undefined,
        'white-space': 'normal', 'text-indent': '0', 'text-align': 'left', 'line-height': 'normal', 'letter-spacing': 'normal', 'font-weight': 'normal', 'font-style': 'normal',
        'text-decoration': 'none', 'text-transform': 'none', 'font-variant': 'normal',
      })}">${cuerpo}</${etiqueta}>`;
    }
    return null;
  }

  // Coloca un dibujo flotante. Con envoltura de texto y casi pegado a su párrafo es un float de CSS (el texto lo
  // rodea, también el de los párrafos que siguen); en cualquier otro caso se posiciona en la hoja.
  function colocarAncla(d, anc, w, h, contenido) {
    const cfg = d.cfg || { ancho: 612, alto: 792, arriba: 72, derecha: 72, abajo: 72, izquierda: 72 };
    const ph = hijo(anc, 'positionH'), pv = hijo(anc, 'positionV');
    const desdeH = ph ? ph.a.relativeFrom : 'column', desdeV = pv ? pv.a.relativeFrom : 'paragraph';
    const alinH = ph && hijo(ph, 'align') ? hijo(ph, 'align').t.trim() : null, alinV = pv && hijo(pv, 'align') ? hijo(pv, 'align').t.trim() : null;
    const despH = ph && hijo(ph, 'posOffset') ? emuPt(hijo(ph, 'posOffset').t) : 0, despV = pv && hijo(pv, 'posOffset') ? emuPt(hijo(pv, 'posOffset').t) : 0;
    const envoltura = ['wrapNone', 'wrapSquare', 'wrapTight', 'wrapThrough', 'wrapTopAndBottom'].find((n) => hijo(anc, n)) || 'wrapNone';
    const lado = envoltura !== 'wrapNone' && hijo(anc, envoltura).a.wrapText;
    const dist = { t: emuPt(anc.a.distT), b: emuPt(anc.a.distB), l: emuPt(anc.a.distL), r: emuPt(anc.a.distR) };
    const detras = anc.a.behindDoc === '1' || anc.a.behindDoc === 'true';
    const { ancho: W, alto: H, arriba: mt, derecha: mr, abajo: mb, izquierda: ml } = cfg;
    const refsH = sinProto({ page: [0, W], margin: [ml, W - mr], column: [ml, W - mr], character: [ml, W - mr], leftMargin: [0, ml], rightMargin: [W - mr, W], insideMargin: [ml, W - mr], outsideMargin: [ml, W - mr] });
    const refsV = sinProto({ page: [0, H], margin: [mt, H - mb], topMargin: [0, mt], bottomMargin: [H - mb, H], insideMargin: [mt, H - mb], outsideMargin: [mt, H - mb] });
    const eje = (refs, desde, alin, desp, tam) => {
      const [a0, a1] = refs[desde] || refs.margin || refs.page;
      if (alin === 'right' || alin === 'bottom' || alin === 'outside') return a1 - tam;
      if (alin === 'center') return (a0 + a1) / 2 - tam / 2;
      return a0 + desp;
    };
    const x = eje(refsH, desdeH, alinH, despH, w);
    const relativoAlParrafo = desdeV === 'paragraph' || desdeV === 'line';

    const conEnvoltura = envoltura !== 'wrapNone' && relativoAlParrafo && Math.abs(despV) < 30;
    if (conEnvoltura) {
      if (envoltura === 'wrapTopAndBottom') {
        const alinear = Math.abs(x + w / 2 - (ml + (W - mr)) / 2) < 8 ? 'center' : x + w / 2 > (ml + W - mr) / 2 ? 'right' : 'left';
        return `<div style="${css({ float: 'left', clear: 'both', width: '100%', 'text-align': alinear, 'padding-top': pt(dist.t + Math.max(despV, 0)), 'padding-bottom': pt(dist.b), 'line-height': '0', 'text-indent': '0', 'box-sizing': 'content-box' })}">${contenido.replace('display:block', 'display:inline-block')}</div>`;
      }
      const centro = x + w / 2, mitad = (ml + W - mr) / 2;
      const aLaDerecha = lado === 'left' ? true : lado === 'right' ? false : centro > mitad;
      if (!aLaDerecha) d.flotIzq = { borde: x - ml + w + dist.r, n: d.nParrafo };
      return `<div style="${css({
        float: aLaDerecha ? 'right' : 'left', 'text-indent': '0', 'line-height': '0',
        margin: `${pt(dist.t + Math.max(despV, 0))} ${pt(aLaDerecha ? Math.max(W - mr - (x + w), 0) : dist.r)} ${pt(dist.b)} ${pt(aLaDerecha ? dist.l : Math.max(x - ml, 0))}`,
      })}">${contenido}</div>`;
    }
    // Posición en la hoja: la sección es el bloque contenedor (position:relative).
    let vertical;
    if (relativoAlParrafo) vertical = `top:auto;margin-top:${pt(despV)}`;
    else vertical = `top:${pt(eje(refsV, desdeV, alinV, despV, h))}`;
    return `<div style="position:absolute;left:${pt(x)};${vertical};z-index:${detras ? -1 : 1};text-indent:0;line-height:0;white-space:normal">${contenido}</div>`;
  }

  // ───────── Párrafos ─────────
  const ALINEACION = sinProto({ left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify', justify: 'justify' });

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

  // ☐ ☒ ☑ (casillas de formulario) no están en las fuentes del PDF: se dibujan como un cuadrito con borde (y una X dentro).
  const CASILLA = /[☐☑☒]/;
  function conCasillas(texto) {
    if (!CASILLA.test(texto)) return esc(texto);
    return texto.split(/([☐☑☒])/).map((t) => {
      if (!CASILLA.test(t)) return esc(t);
      const marcada = t !== '☐';
      return `<span style="display:inline-block;box-sizing:border-box;width:0.8em;height:0.8em;border:0.08em solid #000000;margin:0 0.1em;vertical-align:-0.05em;line-height:0.62em;font-size:0.8em;text-align:center;font-family:Arial,sans-serif;font-weight:bold">${marcada ? 'X' : ''}</span>`;
    }).join('');
  }

  function htmlInline(d, items, esDefecto, bajar) {
    let html = '';
    let anclaEnlace = null;
    const cerrarEnlace = () => { if (anclaEnlace) { html += '</a>'; anclaEnlace = null; } };
    for (const it of agruparRuns(items)) {
      if (it.t === 'ancla') { html += `<a id="${esc(it.nombre)}"></a>`; continue; }
      if (it.enlace !== anclaEnlace) { cerrarEnlace(); if (it.enlace) { html += `<a href="${esc(it.enlace)}" style="color:inherit">`; anclaEnlace = it.enlace; } }
      if (it.t === 'texto') html += `<span style="${cssRun(d, it.r, esDefecto, bajar)}">${conCasillas(it.s)}</span>`;
      else if (it.t === 'tab') html += `<span data-tab="1" style="${cssRun(d, it.r, false)};display:inline-block;white-space:pre">\t</span>`;
      else if (it.t === 'br') { const lado = { all: 'both', left: 'left', right: 'right' }[it.limpiar]; html += `<br${typeof lado === 'string' ? ` style="clear:${lado}"` : ''}>`; }
      else if (it.t === 'campo') html += `<span data-campo="${it.nombre}" style="${cssRun(d, it.r, false, bajar)}">1</span>`;
      else if (it.t === 'html') html += it.html;
      // 'flotante': los dibujos flotantes los pone parrafo() antes del <p>
    }
    cerrarEnlace();
    return html;
  }

  // El PDF solo puede escribir las viñetas que están en WinAnsi o en Noto Sans (html-a-pdf.js): el resto
  // (●, ■, ➢…) saldría en blanco. Los círculos vacíos quedan como "o"; los demás, como "•".
  const simboloEscribible = (s) => (/^[○◦□▫o]$/.test(s) ? 'o' : '•');

  // Un párrafo de Word → uno o varios <p> (un salto de página parte el párrafo en dos).
  async function parrafo(d, p, medio) {
    const m = d.m;
    d.nParrafo++;
    const pDirecto = m.leerPPr(hijo(p, 'pPr'));
    const est = m.estilos.parrafo(pDirecto.estilo, medio);
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
    const rellenoP = rellenoCss(efectivo.sombreado);
    const espacioLado = (lado) => (b[lado] && b[lado].espacio) || 0;
    const margenIzq = izq - (bordes.left ? espacioLado('left') + bordes.left.ancho : 0);
    const margenDer = der - (bordes.right ? espacioLado('right') + bordes.right.ancho : 0);

    const yDesde = (f) => (f.y ? parseFloat(f.y) / 20 : 0);
    // Párrafo en un marco (w:framePr): flotante si va pegado a un lado, o colocado en la hoja.
    let marco = null;
    const fr = efectivo.marco;
    if (fr && (fr.w || fr.x !== undefined || fr.xAlign)) {
      const cfg = d.cfg || { ancho: 612, arriba: 72, izquierda: 72, derecha: 72 };
      const ancho = fr.w ? parseFloat(fr.w) / 20 : undefined;
      const alto = fr.h ? parseFloat(fr.h) / 20 : undefined;
      const hs = fr.hSpace ? parseFloat(fr.hSpace) / 20 : 0, vs = fr.vSpace ? parseFloat(fr.vSpace) / 20 : 0;
      const estiloAncho = { width: ancho ? pt(ancho) : undefined, height: alto && fr.hRule === 'exact' ? pt(alto) : undefined, 'min-height': alto && fr.hRule !== 'exact' ? pt(alto) : undefined, 'box-sizing': 'border-box' };
      const alrededor = fr.wrap !== 'none' && fr.wrap !== 'notBeside' && fr.vAnchor !== 'page' && fr.vAnchor !== 'margin';
      if (alrededor) {
        // Junto a un dibujo flotante a la izquierda (la foto de un currículum) el marco empieza donde ese termina.
        const previo = d.flotIzq && d.nParrafo - d.flotIzq.n <= 2 ? d.flotIzq.borde : 0;
        const x0 = fr.xAlign ? 0 : Math.max(parseFloat(fr.x || 0) / 20 - previo, 0);
        marco = { ...estiloAncho, float: fr.xAlign === 'right' ? 'right' : 'left', margin: `${pt(yDesde(fr))} ${pt(hs)} ${pt(vs)} ${pt(x0)}` };
      } else {
        const x0 = fr.hAnchor === 'page' ? 0 : cfg.izquierda;
        const yDesp = fr.y ? parseFloat(fr.y) / 20 : 0;
        const vertical = fr.vAnchor === 'page' ? { top: pt(yDesp) } : fr.vAnchor === 'margin' ? { top: pt(cfg.arriba + yDesp) } : { top: 'auto', 'margin-top': pt(yDesp) };
        marco = { ...estiloAncho, position: 'absolute', left: pt(x0 + (fr.xAlign === 'center' ? ((cfg.ancho - cfg.izquierda - cfg.derecha) - (ancho || 0)) / 2 : parseFloat(fr.x || 0) / 20)), ...vertical, 'z-index': 1, margin: undefined };
        if (vertical['margin-top']) marco['margin-top'] = vertical['margin-top'];
      }
    }
    const estilo = (extra) => css({
      margin: `0 ${pt(Math.max(margenDer, 0))} 0 ${pt(Math.max(margenIzq, 0))}`,
      'text-indent': primera ? pt(primera) : undefined,
      'text-align': ALINEACION[efectivo.jc] || 'left',
      'text-align-last': efectivo.jc === 'distribute' ? 'justify' : undefined,
      'white-space': 'pre-wrap', 'overflow-wrap': 'break-word',
      'font-family': famBase, 'font-size': pt(tamMax), 'line-height': lineHeight || 'normal',
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

    // Parte el párrafo en trozos: un salto de página empieza una hoja nueva; un cuadro de texto en línea (que lleva
    // párrafos adentro y no cabe en un <p>) queda entre dos trozos, en su propia línea.
    const trozos = [{ items: [], pagina: false }];
    for (const it of items) {
      if (it.t === 'pagina') trozos.push({ items: [], pagina: true });
      else if (it.t === 'bloque') { trozos.push({ bloque: it.html }); trozos.push({ items: [], pagina: false, tras: true }); }
      else trozos[trozos.length - 1].items.push(it);
    }
    const espaciador = (h) => (h > 0 ? `<div style="height:${pt(h)};line-height:0;font-size:0"></div>` : '');
    let html = marco ? '' : espaciador(antes);
    html += items.filter((x) => x.t === 'flotante').map((x) => x.html).join('');
    const hayBloques = trozos.some((x) => x.bloque !== undefined);
    trozos.forEach((trozo, i) => {
      if (trozo.bloque !== undefined) { html += trozo.bloque; return; }
      const vacio = !trozo.items.some((x) => x.t !== 'ancla');
      if (hayBloques && vacio && !trozo.pagina) return; // lo que queda antes o después de un cuadro de texto, si no tiene nada
      const interior = htmlInline(d, trozo.items, false, bajar) || `<span style="${cssRun(d, marcaR, true)}"><br></span>`;
      html += `<p ${atributos} style="${estilo({ ...(marco || {}), ...(trozo.pagina ? { 'break-before': 'page' } : {}), ...(i > 0 ? { 'text-indent': undefined } : {}) })}">${interior}</p>`;
    });
    if (!marco) html += espaciador(despues);
    const sectPr = hijo(hijo(p, 'pPr'), 'sectPr');
    return { html, sectPr };
  }

  // ───────── Tablas ─────────
  // Ancho de una tabla o celda en pt: dxa = veinteavos de punto; pct = 50avos de por ciento (o "100%").
  function anchoPt(w, disponible) {
    if (!w || w.valor === undefined || Number.isNaN(w.valor)) return undefined;
    if (w.tipo === 'dxa') return w.valor / 20;
    if (w.tipo === 'pct') return disponible * (w.porcentaje !== undefined ? w.porcentaje : w.valor / 50) / 100;
    return undefined;
  }
  function leerAncho(e) { // como DocxEstilos, pero deja pasar "100%"
    if (!e || e.a.w === undefined) return undefined;
    const crudo = String(e.a.w);
    return { tipo: e.a.type || 'dxa', valor: parseFloat(crudo), porcentaje: /%$/.test(crudo) ? parseFloat(crudo) : undefined };
  }

  // Formatos condicionales del estilo de tabla que valen para la celda (de menor a mayor prioridad).
  function condicionalesDe(estiloT, aspecto, tblPr, i, j, nFilas, nCols) {
    const sal = [];
    const c = estiloT.condicionales || {};
    const a = aspecto || { primeraFila: true, primeraColumna: true };
    const bandaF = tblPr.bandasFila || 1, bandaC = tblPr.bandasColumna || 1;
    const esPrimeraFila = a.primeraFila && i === 0, esUltimaFila = a.ultimaFila && i === nFilas - 1;
    const esPrimeraCol = a.primeraColumna && j === 0, esUltimaCol = a.ultimaColumna && j === nCols - 1;
    if (c.wholeTable) sal.push(c.wholeTable);
    if (!a.sinBandasV && !esPrimeraCol && !esUltimaCol) {
      const k = Math.floor((j - (a.primeraColumna ? 1 : 0)) / bandaC);
      const banda = k % 2 === 0 ? c.band1Vert : c.band2Vert;
      if (banda) sal.push(banda);
    }
    if (!a.sinBandasH && !esPrimeraFila && !esUltimaFila) {
      const k = Math.floor((i - (a.primeraFila ? 1 : 0)) / bandaF);
      const banda = k % 2 === 0 ? c.band1Horz : c.band2Horz;
      if (banda) sal.push(banda);
    }
    if (esPrimeraCol && c.firstCol) sal.push(c.firstCol);
    if (esUltimaCol && c.lastCol) sal.push(c.lastCol);
    if (esPrimeraFila && c.firstRow) sal.push(c.firstRow);
    if (esUltimaFila && c.lastRow) sal.push(c.lastRow);
    if (esPrimeraFila && esPrimeraCol && c.nwCell) sal.push(c.nwCell);
    if (esPrimeraFila && esUltimaCol && c.neCell) sal.push(c.neCell);
    if (esUltimaFila && esPrimeraCol && c.swCell) sal.push(c.swCell);
    if (esUltimaFila && esUltimaCol && c.seCell) sal.push(c.seCell);
    return sal;
  }

  // Hijos con ese nombre, también los que están dentro de un control de contenido (sdt > sdtContent).
  function desenvolver(e, nombre, prof = 0) {
    const sal = [];
    for (const x of e.h) {
      if (x.n === nombre) sal.push(x);
      else if (x.n === 'sdt' && prof < 8) { const c = hijo(x, 'sdtContent'); if (c) sal.push(...desenvolver(c, nombre, prof + 1)); }
    }
    return sal;
  }

  async function tabla(d, t, disponible) {
    const m = d.m;
    const tblDirecto = m.leerTblPr(hijo(t, 'tblPr'));
    const wDirecto = leerAncho(hijo(hijo(t, 'tblPr'), 'tblW'));
    const estiloT = m.estilos.tabla(tblDirecto.estilo);
    const tbl = m.mezclar(estiloT.tbl, tblDirecto);
    const grid = lista(hijo(t, 'tblGrid'), 'gridCol').map((c) => parseFloat(c.a.w) / 20);
    // Los controles de contenido (w:sdt) de un formulario pueden envolver filas o celdas enteras.
    const filas = desenvolver(t, 'tr');
    // Posición de columna de cada celda y alto (rowspan) de las combinadas verticalmente.
    const infos = filas.map((f) => {
      let col = 0;
      return desenvolver(f, 'tc').map((c) => {
        const pr = m.leerTcPr(hijo(c, 'tcPr'));
        // gridSpan viene del archivo: entero, de 1 a lo que tenga la rejilla (Word no pasa de 63 columnas).
        const span = Math.min(Math.max(Math.floor(pr.columnas) || 1, 1), Math.max(grid.length, 64));
        const info = { c, pr, col, span, w: leerAncho(hijo(hijo(c, 'tcPr'), 'tcW')) };
        col += span;
        return info;
      });
    });
    const nCols = infos.reduce((mx, f) => Math.max(mx, f.reduce((a, x) => a + x.span, 0)), Math.max(grid.length, 1)); // sin "...": con 10^5 filas no caben como argumentos
    while (grid.length < nCols) grid.push(0);
    infos.forEach((fila, i) => fila.forEach((info) => {
      info.filas = 1;
      if (info.pr.combinadaVertical === 'restart') {
        for (let k = i + 1; k < infos.length; k++) {
          const sig = infos[k].find((x) => x.col === info.col);
          if (sig && sig.pr.combinadaVertical === 'continue') { sig.omitir = true; info.filas++; } else break;
        }
      }
    }));

    // Anchos: la tabla ocupa lo que dice tblW (si lo dice) y las columnas se reparten en proporción a la rejilla.
    const wTabla = anchoPt(wDirecto || (tbl.ancho && { tipo: tbl.ancho.tipo, valor: tbl.ancho.valor }), disponible);
    let cols = grid.slice();
    // Rejilla sin anchos reales (las tablas que la escriben en 100): se usan los de las celdas de la primera fila.
    const sumaGrid = cols.reduce((a, b) => a + b, 0);
    if (infos[0]) {
      const preferidas = new Array(nCols).fill(0);
      let completas = true;
      for (const info of infos[0]) {
        const w = anchoPt(info.w, wTabla || disponible);
        if (w === undefined) { completas = false; continue; }
        for (let k = 0; k < info.span; k++) preferidas[info.col + k] = w / info.span;
      }
      if (completas && preferidas.every((x) => x > 0) && (!wTabla || sumaGrid < 1 || Math.abs(sumaGrid - preferidas.reduce((a, b) => a + b, 0)) > 1)) cols = preferidas;
    }
    let total = cols.reduce((a, b) => a + b, 0);
    if (total <= 0) { cols = new Array(nCols).fill((wTabla || disponible) / nCols); total = cols.reduce((a, b) => a + b, 0); }
    if (wTabla !== undefined && Math.abs(wTabla - total) > 0.5) { const f = wTabla / total; cols = cols.map((x) => x * f); total = wTabla; }
    else if (wTabla === undefined && total > disponible && disponible > 0) { const f = disponible / total; cols = cols.map((x) => x * f); total = disponible; }

    // Sin ancho en ninguna celda y con una rejilla de mentira (la escriben así algunas librerías): Word reparte las
    // columnas según el contenido; CSS hace lo mismo con table-layout:auto.
    const rejillaFalsa = infos.every((f) => f.every((x) => !x.w || !(x.w.valor > 0))) && (sumaGrid / nCols) < 30;
    // Word ajusta las columnas al contenido salvo que la tabla sea de ancho fijo: una palabra que no cabe ensancha su columna.
    const autoajuste = tbl.disposicion !== 'fixed';
    const margenesT = { top: 0, bottom: 0, left: 0, right: 0, ...(tbl.margenes || {}) }; // sin estilo de tabla (TableNormal) Word no pone margen de celda
    const bT = tbl.bordes || {};
    // Antes de Word 2013 (modo de compatibilidad < 15) el borde de la tabla sale al margen izquierdo de la celda
    // hacia afuera, para que el texto quede alineado con el del párrafo.
    const sangria = (tbl.sangria && tbl.sangria.tipo !== 'pct' ? (tbl.sangria.valor || 0) / 20 : 0) - (m.ajustes.modoCompat < 15 ? margenesT.left : 0);
    const alineacion = tbl.jc === 'center' ? 'center' : tbl.jc === 'right' || tbl.jc === 'end' ? 'flex-end' : 'flex-start';
    let html = `<table style="${css({
      // Las hojas de estilo de la página ponen cosas a table/td (interlineado, relleno, bordes): aquí se fija todo.
      'border-collapse': 'collapse', 'border-spacing': '0', 'table-layout': autoajuste ? 'auto' : 'fixed', width: rejillaFalsa && wTabla === undefined ? undefined : pt(total), 'max-width': rejillaFalsa ? pt(disponible) : undefined, margin: '0', 'align-self': alineacion, 'line-height': 'normal', 'text-align': 'left', background: 'none', border: 'none',
      'margin-left': alineacion === 'flex-start' && sangria ? pt(sangria) : undefined,
    })}">${rejillaFalsa ? '' : `<colgroup>${cols.map((g) => `<col style="width:${pt(g)}">`).join('')}</colgroup>`}`;
    for (let i = 0; i < infos.length; i++) {
      const trPr = m.leerTrPr(hijo(filas[i], 'trPr'));
      html += `<tr style="${css({ height: trPr.alto ? pt(trPr.alto) : undefined, 'break-inside': trPr.noPartir ? 'avoid' : undefined, background: 'none' })}">`;
      for (const info of infos[i]) {
        if (info.omitir) continue;
        const conds = condicionalesDe(estiloT, tbl.aspecto, tbl, i, info.col, infos.length, nCols);
        let tc = {}, medioP = {}, medioR = {};
        for (const c of conds) { tc = m.mezclar(tc, c.tc); medioP = m.mezclar(medioP, c.p); medioR = m.mezclar(medioR, c.r); }
        tc = m.mezclar(m.mezclar(estiloT.tc, tc), info.pr);
        const medio = { p: m.mezclar(estiloT.p, medioP), r: m.mezclar(estiloT.r, medioR) };
        const ultimaCol = info.col + info.span >= nCols, ultimaFila = i + info.filas >= infos.length;
        const bc = tc.bordes || {};
        // Borde de cada lado: el de la celda; si no, el de la tabla (en el interior son insideH / insideV).
        const lado = (nombre, interior, exterior) => cssBorde(d, bc[nombre] || (exterior ? bT[nombre] : bT[interior]));
        const bt = lado('top', 'insideH', i === 0), bb = lado('bottom', 'insideH', ultimaFila);
        const bl = lado('left', 'insideV', info.col === 0), br = lado('right', 'insideV', ultimaCol);
        const mc = { ...margenesT, ...(tc.margenes || {}) };
        const anchoCelda = cols.slice(info.col, info.col + info.span).reduce((a, b) => a + b, 0);
        const dentro = await bloques(d, info.c, Math.max(anchoCelda - mc.left - mc.right, 20), medio);
        html += `<td${info.span > 1 ? ` colspan="${info.span}"` : ''}${info.filas > 1 ? ` rowspan="${info.filas}"` : ''} style="${css({
          'vertical-align': tc.alineacionVertical === 'center' ? 'middle' : tc.alineacionVertical === 'bottom' ? 'bottom' : 'top',
          padding: `${pt(mc.top)} ${pt(mc.right)} ${pt(mc.bottom)} ${pt(mc.left)}`,
          'border-top': bt && bt.texto, 'border-bottom': bb && bb.texto, 'border-left': bl && bl.texto, 'border-right': br && br.texto,
          'background-color': rellenoCss(tc.sombreado) || rellenoCss(tbl.sombreado), 'overflow-wrap': 'break-word', 'line-height': 'normal', 'text-align': 'left',
          'font-weight': 'normal', 'font-size': '10pt', color: 'inherit',
        })}">${dentro.html}</td>`;
      }
      html += '</tr>';
    }
    return `${html}</table>`;
  }

  // ───────── Bloques y secciones ─────────
  // Bloques de una celda (o de cualquier contenedor sin secciones). `medio`: formato del estilo de tabla.
  async function bloques(d, contenedor, disponible, medio) {
    let html = '';
    for (const e of contenedor.h) {
      if (e.n === 'p') html += (await parrafo(d, e, medio)).html;
      else if (e.n === 'tbl') html += await tabla(d, e, disponible);
      else if (e.n === 'sdt') { const c = hijo(e, 'sdtContent'); if (c) html += (await bloques(d, c, disponible, medio)).html; }
    }
    return { html };
  }

  function leerSeccion(m, sectPr) {
    const sz = sectPr && hijo(sectPr, 'pgSz');
    const mg = sectPr && hijo(sectPr, 'pgMar');
    const cols = sectPr && hijo(sectPr, 'cols');
    const num = sectPr && hijo(sectPr, 'pgNumType');
    const v = (x, defecto) => { const n = parseFloat(x); return Number.isFinite(n) ? n / 20 : defecto; };
    // Un archivo armado a propósito puede pedir una hoja de 0 o de kilómetros, o márgenes más altos que la hoja (y
    // html-a-pdf.js no podría avanzar): la hoja va de 1 a 22 pulgadas y al cuerpo le queda al menos un cuarto de ella.
    // Un margen negativo arriba o abajo en Word es el mismo margen "exacto": vale su valor absoluto.
    const ancho = acotar(v(sz && sz.a.w, 612), 72, MAX_PT), alto = acotar(v(sz && sz.a.h, 792), 72, MAX_PT);
    let arriba = Math.abs(v(mg && mg.a.top, 72)), abajo = Math.abs(v(mg && mg.a.bottom, 72));
    let izquierda = Math.max(v(mg && mg.a.left, 72), 0), derecha = Math.max(v(mg && mg.a.right, 72), 0);
    if (arriba + abajo > alto * 0.75) { const f = (alto * 0.75) / (arriba + abajo); arriba *= f; abajo *= f; }
    if (izquierda + derecha > ancho * 0.75) { const f = (ancho * 0.75) / (izquierda + derecha); izquierda *= f; derecha *= f; }
    const inicio = num ? parseInt(num.a.start, 10) : NaN;
    return {
      ancho, alto, arriba, derecha, abajo, izquierda,
      encabezado: v(mg && mg.a.header, 36), pie: v(mg && mg.a.footer, 36),
      continua: !!(sectPr && hijo(sectPr, 'type') && hijo(sectPr, 'type').a.val === 'continuous'),
      columnas: cols && parseInt(cols.a.num, 10) > 1 ? { n: parseInt(cols.a.num, 10), espacio: v(cols.a.space, 36), linea: cols.a.sep === '1' || cols.a.sep === 'true' } : null,
      portada: !!(sectPr && hijo(sectPr, 'titlePg')),
      inicioPaginas: Number.isFinite(inicio) ? acotar(inicio, 0, 1e6) : undefined,
      referencias: sectPr ? lista(sectPr, 'headerReference').map((x) => ({ tipo: 'encabezado', cual: x.a.type || 'default', id: x.a.id }))
        .concat(lista(sectPr, 'footerReference').map((x) => ({ tipo: 'pie', cual: x.a.type || 'default', id: x.a.id }))) : [],
    };
  }

  // Contenido de un encabezado o pie (word/headerN.xml, con sus propias relaciones para las imágenes).
  async function leerEncabezado(d, id, ancho) {
    const rel = d.rels.get(id);
    const f = rel && d.zip.file(rel.destino);
    if (!f) return '';
    const raiz = DocxBase.arbol(await f.async('string'));
    const relsPropias = await DocxBase.leerRels(d.zip, rel.destino);
    const guardadas = d.rels;
    d.rels = relsPropias;
    try { return (await bloques(d, raiz, ancho)).html; } finally { d.rels = guardadas; }
  }

  // Los encabezados y pies de la sección: los propios; si no hay, los de la sección anterior.
  async function repetidosDe(d, cfg) {
    const previos = d.repetidosPrevios || {};
    const actuales = { ...previos };
    const ancho = cfg.ancho - cfg.izquierda - cfg.derecha;
    for (const r of cfg.referencias) {
      const html = await leerEncabezado(d, r.id, ancho);
      actuales[`${r.tipo}:${r.cual}`] = html;
    }
    d.repetidosPrevios = actuales;
    let html = '';
    for (const clave of Object.keys(actuales)) {
      const [tipo, cual] = clave.split(':');
      if (cual === 'first' && !cfg.portada) continue;
      html += `<div data-repetido="${tipo}" data-cual="${esc(cual)}" data-dist="${n3(tipo === 'encabezado' ? cfg.encabezado : cfg.pie)}" style="position:absolute;left:${pt(cfg.izquierda)};top:0;width:${pt(ancho)};line-height:normal">${actuales[clave]}</div>`;
    }
    return html;
  }

  function htmlSeccion(s, partes, repetidos, parImpar) {
    // Con columnas, CSS reparte el texto en partes iguales (como Word al final de una sección continua).
    const cuerpo = partes.map(({ columnas: c, html }) => (c ? `<div style="${css({ 'column-count': c.n, 'column-gap': pt(c.espacio), 'column-rule': c.linea ? '0.5pt solid #000000' : undefined })}">${html}</div>` : html)).join('');
    return `<section class="docx" data-titulopg="${s.portada ? 1 : 0}" data-parimpar="${parImpar ? 1 : 0}"${s.inicioPaginas !== undefined ? ` data-pginicio="${s.inicioPaginas}"` : ''} style="${css({
      'box-sizing': 'border-box', display: 'block', position: 'relative', 'z-index': '0', width: pt(s.ancho), 'min-height': pt(s.alto),
      padding: `${pt(s.arriba)} ${pt(s.derecha)} ${pt(s.abajo)} ${pt(s.izquierda)}`, margin: '0', background: '#ffffff', color: '#000000',
    })}">${repetidos || ''}${cuerpo}</section>`;
  }

  // ───────── Tabulaciones (después de maquetar) ─────────
  async function resolverTabulaciones(html, d) {
    if (typeof document === 'undefined' || html.indexOf('data-tab="') < 0) return html;
    const lienzo = document.createElement('div');
    lienzo.style.cssText = 'position:absolute;left:-100000px;top:0;visibility:hidden';
    // El HTML se mete en la página para medirlo: antes pasa por el sanitizador (como al final, en word-a-pdf.js),
    // para que nada del archivo llegue a correr aunque algún valor se haya escapado mal.
    lienzo.innerHTML = typeof sanitizarHtml === 'function' ? sanitizarHtml(html) : html;
    document.body.appendChild(lienzo);
    try {
      if (document.fonts && document.fonts.ready) await document.fonts.ready;
      const paso = d.m.ajustes.tabPorDefecto * PX_POR_PT;
      const maxAncho = 2 * MAX_PT * PX_POR_PT;
      const ctx = document.createElement('canvas').getContext('2d');
      const parrafos = [];
      let maxTabs = 0;
      for (const p of lienzo.querySelectorAll('p[data-izq]')) {
        const tabs = p.querySelectorAll('[data-tab]');
        if (!tabs.length) continue;
        const izq = parseFloat(p.dataset.izq) * PX_POR_PT;
        let personalizadas = [];
        try { personalizadas = JSON.parse(p.dataset.tabs || '[]'); } catch (x) { personalizadas = []; }
        const paradas = personalizadas.filter((t) => t.tipo !== 'clear' && Number.isFinite(t.pos)).map((t) => ({ ...t, px: t.pos * PX_POR_PT }));
        const sangria = parseFloat(p.style.textIndent) || 0;
        if (sangria < 0) paradas.push({ tipo: 'left', px: izq }); // la sangría francesa también es una parada
        parrafos.push({ p, tabs, izq, e: getComputedStyle(p), paradas });
        maxTabs = Math.max(maxTabs, tabs.length);
      }
      // Por rondas: la tabulación k de todos los párrafos a la vez (primero se escribe, después se mide, después se
      // escribe el ancho). Medir después de cada cambio obligaba al navegador a rehacer el documento entero por cada
      // tabulación: con 5000 elementos de lista (cada uno con su tabulación) eran 6,5 s.
      for (let k = 0; k < maxTabs; k++) {
        const ronda = parrafos.filter((x) => k < x.tabs.length);
        for (const { tabs } of ronda) { tabs[k].style.width = '0px'; tabs[k].textContent = '​'; }
        const medidas = ronda.map(({ p, tabs, izq, e }) => {
          const tab = tabs[k];
          const rp = p.getBoundingClientRect();
          const origen = rp.left + parseFloat(e.borderLeftWidth) + parseFloat(e.paddingLeft) - izq;
          let siguiente = 0;
          for (let n = tab.nextSibling; n; n = n.nextSibling) {
            if (n.nodeType === 1 && n.hasAttribute('data-tab')) break;
            if (n.nodeType === 1 && n.tagName === 'BR') break;
            siguiente += n.nodeType === 1 ? n.getBoundingClientRect().width : 0;
          }
          return { x0: tab.getBoundingClientRect().left - origen, siguiente };
        });
        ronda.forEach(({ tabs, paradas }, i) => {
          const tab = tabs[k];
          const { x0, siguiente } = medidas[i];
          const parada = paradas.filter((s) => s.px > x0 + 0.5).sort((a, b) => a.px - b.px)[0];
          const porDefecto = (Math.floor((x0 + 0.5) / paso) + 1) * paso;
          // Las paradas por defecto solo valen después de la última parada propia.
          const meta = parada || { tipo: 'left', px: porDefecto };
          let ancho = meta.px - x0;
          if (meta.tipo === 'right' || meta.tipo === 'center' || meta.tipo === 'decimal') ancho -= meta.tipo === 'center' ? siguiente / 2 : siguiente;
          // Una parada a 10^11 twips pedía mil millones de puntos de guía: el ancho no pasa del de dos hojas.
          ancho = Math.min(Math.max(ancho, 0), maxAncho);
          const caracter = GUIAS[meta.guia];
          tab.style.width = n3(ancho) + 'px';
          if (caracter && ancho > 0) {
            const ec = getComputedStyle(tab);
            ctx.font = `${ec.fontStyle} ${ec.fontWeight} ${ec.fontSize} ${ec.fontFamily}`;
            const w = ctx.measureText(caracter).width || 1;
            tab.textContent = caracter.repeat(Math.max(Math.floor(ancho / w), 0));
          }
        });
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
    const secciones = [];
    for (const t of lectura) {
      const cfg = leerSeccion(modelo, t.sectPr);
      d.cfg = cfg;
      const previa = secciones[secciones.length - 1];
      // Una sección continua (sin salto de página) comparte la hoja con la anterior; cada parte lleva sus columnas.
      if (previa && cfg.continua && previa.cfg.ancho === cfg.ancho && previa.cfg.alto === cfg.alto && !cfg.referencias.length) previa.partes.push({ columnas: cfg.columnas, html: t.html });
      else secciones.push({ cfg, partes: [{ columnas: cfg.columnas, html: t.html }], repetidos: await repetidosDe(d, cfg) });
    }
    // Notas al pie y finales: al final del documento (Word las pone al pie de cada página).
    if (d.notas.length && secciones.length) {
      const ultima = secciones[secciones.length - 1];
      ultima.partes.push({ columnas: null, html: await htmlNotas(d, ultima.cfg.ancho - ultima.cfg.izquierda - ultima.cfg.derecha) });
    }
    if (typeof PptxTexto !== 'undefined' && PptxTexto.cargarFuentes) await PptxTexto.cargarFuentes([...d.fuentes.values()]);
    const html = await resolverTabulaciones(secciones.map((x) => htmlSeccion(x.cfg, x.partes, x.repetidos, modelo.ajustes.parImpar)).join(''), d);
    return { html, mensajes };
  }

  async function htmlNotas(d, ancho) {
    let html = '<div style="width:144pt;border-top:0.5pt solid #000000;margin-top:14pt;margin-bottom:6pt"></div>';
    for (const tipo of [false, true]) {
      const rel = [...d.rels.values()].find((r) => r.tipo.endsWith(tipo ? '/endnotes' : '/footnotes'));
      const f = d.zip.file(rel ? rel.destino : tipo ? 'word/endnotes.xml' : 'word/footnotes.xml');
      if (!f || !d.notas.some((n) => n.fin === tipo)) continue;
      const raiz = DocxBase.arbol(await f.async('string'));
      for (const [i, n] of d.notas.entries()) {
        if (n.fin !== tipo) continue;
        const nota = lista(raiz, tipo ? 'endnote' : 'footnote').find((x) => x.a.id === n.id);
        if (!nota) continue;
        d.numeroNotaActual = i + 1;
        html += (await bloques(d, nota, ancho)).html;
      }
    }
    d.numeroNotaActual = 0;
    return html;
  }

  // Recorre el cuerpo y lo reparte en secciones: cada párrafo con sectPr cierra una; la última es la del cuerpo.
  async function bloquesCuerpo(d, cuerpo) {
    const sectFinal = lista(cuerpo, 'sectPr').pop() || null;
    // Primero se ve a qué sección pertenece cada bloque (para saber el ancho disponible de las tablas).
    const seccionDe = new Map();
    let pendientes = [];
    for (const e of cuerpo.h) {
      if (e.n === 'sectPr') continue;
      pendientes.push(e);
      const sp = e.n === 'p' ? hijo(hijo(e, 'pPr'), 'sectPr') : null;
      if (sp) { for (const x of pendientes) seccionDe.set(x, sp); pendientes = []; }
    }
    for (const x of pendientes) seccionDe.set(x, sectFinal);
    const sal = [];
    let actual = '';
    for (const e of cuerpo.h) {
      if (e.n === 'sectPr') continue;
      const cfg = leerSeccion(d.m, seccionDe.get(e));
      const disponible = cfg.ancho - cfg.izquierda - cfg.derecha;
      d.cfg = cfg;
      if (e.n === 'p') {
        const r = await parrafo(d, e);
        actual += r.html;
        if (r.sectPr) { sal.push({ html: actual, sectPr: r.sectPr }); actual = ''; }
      } else if (e.n === 'tbl') actual += await tabla(d, e, disponible);
      else if (e.n === 'sdt') { const c = hijo(e, 'sdtContent'); if (c) actual += (await bloques(d, c, disponible)).html; }
    }
    sal.push({ html: actual, sectPr: sectFinal });
    // Si el cuerpo termina con un párrafo que ya cerró su sección y no queda nada más, esa última sección vacía sobra.
    if (sal.length > 1 && !sal[sal.length - 1].html && !sectFinal) sal.pop();
    return sal;
  }
})();

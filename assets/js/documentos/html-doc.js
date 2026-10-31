/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// El HTML del portapapeles de Documentos (Fase 6, etapa 6.7). Sin DOM: corre igual en Node, y lee el HTML tal como lo manda Word, LibreOffice,
// Google Docs o un navegador (etiquetas sin cerrar, comentarios condicionales de Office, estilos «mso-»).
//
//   HtmlDoc.leer(html) → [{ texto, tramos: [{ d, h, f }], jc, estilo, lista }]      un elemento por párrafo
//   HtmlDoc.escribir(parrafos) → HTML        (lo contrario; lo que se copia al portapapeles)
//
//   tramo.f: formato de caracteres con los mismos nombres que EdicionDoc.formatoCaracter: b, i, u ('single'), tachado, color ('RRGGBB'), tam (pt),
//            fuente, resaltado (uno de los 16 de Word), vertical ('superscript' | 'subscript')
//   jc: 'center' | 'right' | 'both' (la izquierda es lo normal y no se anota) · estilo: 'Heading1'…'Heading3' · lista: { tipo: 'vinetas'|'numeros', nivel }
//
// No lee tablas (cada celda sale como un párrafo), imágenes ni enlaces (queda el texto). El salto de línea es U+2028, como en el resto del editor.
(function () {
  'use strict';
  const SUAVE = String.fromCharCode(0x2028), NBSP = String.fromCharCode(0xA0);
  const ENTIDADES = {
    nbsp: NBSP, amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", iexcl: '¡', iquest: '¿', laquo: '«', raquo: '»', ndash: '–', mdash: '—', hellip: '…',
    euro: '€', copy: '©', reg: '®', deg: '°', middot: '·', bull: '•', lsquo: '‘', rsquo: '’', ldquo: '“', rdquo: '”', times: '×', ordm: 'º', ordf: 'ª',
    aacute: 'á', eacute: 'é', iacute: 'í', oacute: 'ó', uacute: 'ú', Aacute: 'Á', Eacute: 'É', Iacute: 'Í', Oacute: 'Ó', Uacute: 'Ú', ntilde: 'ñ', Ntilde: 'Ñ',
    uuml: 'ü', Uuml: 'Ü', agrave: 'à', egrave: 'è', ccedil: 'ç', auml: 'ä', ouml: 'ö', szlig: 'ß',
  };
  const decodificar = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|[a-z][a-z0-9]*);/gi, (m, e) => {
    if (e[0] === '#') { const n = e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10); return n > 0 && n < 0x110000 ? String.fromCodePoint(n) : ''; }
    return Object.prototype.hasOwnProperty.call(ENTIDADES, e) ? ENTIDADES[e] : m;
  });
  const BLOQUE = new Set(['p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'pre', 'blockquote', 'tr', 'td', 'th', 'ul', 'ol', 'table', 'section', 'article', 'header', 'footer', 'main', 'aside', 'nav', 'form', 'fieldset', 'dl', 'dt', 'dd', 'figure', 'figcaption', 'address', 'center']);
  const VACIAS = new Set(['br', 'img', 'hr', 'meta', 'link', 'input', 'wbr', 'col', 'area', 'base', 'param', 'source', 'track', 'embed']);
  const OCULTAS = new Set(['script', 'style', 'head', 'title', 'template', 'xml', 'object', 'svg', 'noscript', 'iframe']);
  const RESALTES = { yellow: [255, 255, 0], green: [0, 255, 0], cyan: [0, 255, 255], magenta: [255, 0, 255], blue: [0, 0, 255], red: [255, 0, 0], darkBlue: [0, 0, 139], darkCyan: [0, 139, 139], darkGreen: [0, 100, 0], darkMagenta: [139, 0, 139], darkRed: [139, 0, 0], darkYellow: [128, 128, 0], darkGray: [128, 128, 128], lightGray: [211, 211, 211], black: [0, 0, 0], white: [255, 255, 255] };
  const NOMBRES_COLOR = { red: 'FF0000', blue: '0000FF', green: '008000', yellow: 'FFFF00', orange: 'FFA500', purple: '800080', gray: '808080', grey: '808080', white: 'FFFFFF', navy: '000080', teal: '008080', maroon: '800000', silver: 'C0C0C0', lime: '00FF00', aqua: '00FFFF', fuchsia: 'FF00FF', olive: '808000', brown: 'A52A2A', pink: 'FFC0CB' };

  // ───────── Estilos en línea ─────────
  function estilos(s) {
    const m = {};
    for (const par of String(s || '').split(';')) { const i = par.indexOf(':'); if (i > 0) m[par.slice(0, i).trim().toLowerCase()] = par.slice(i + 1).trim(); }
    return m;
  }
  function color(v) {
    v = String(v || '').trim().toLowerCase();
    let m = /^#([0-9a-f]{6})$/.exec(v);
    if (m) return m[1].toUpperCase();
    if ((m = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(v))) return (m[1] + m[1] + m[2] + m[2] + m[3] + m[3]).toUpperCase();
    if ((m = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)/.exec(v))) return [m[1], m[2], m[3]].map((x) => Math.min(255, +x).toString(16).padStart(2, '0')).join('').toUpperCase();
    return NOMBRES_COLOR[v] || null;
  }
  function resaltado(v) {
    v = String(v || '').trim();
    if (RESALTES[v]) return v;
    const c = color(v);
    if (!c) return null;
    const rgb = [parseInt(c.slice(0, 2), 16), parseInt(c.slice(2, 4), 16), parseInt(c.slice(4, 6), 16)];
    if (rgb.every((x) => x > 245)) return null;   // el fondo blanco no es un resaltado
    let mejor = 'yellow', d = Infinity;
    for (const [n, r] of Object.entries(RESALTES)) { const dd = (r[0] - rgb[0]) ** 2 + (r[1] - rgb[1]) ** 2 + (r[2] - rgb[2]) ** 2; if (dd < d) { d = dd; mejor = n; } }
    return mejor;
  }
  const tamPt = (v) => {
    const m = /^([\d.]+)\s*(pt|px)$/i.exec(String(v || '').trim());
    if (!m) return null;
    const pt = m[2].toLowerCase() === 'px' ? parseFloat(m[1]) * 0.75 : parseFloat(m[1]);
    return pt >= 1 && pt <= 400 ? Math.round(pt * 2) / 2 : null;
  };
  function familia(v) {
    const primera = String(v || '').split(',')[0].replace(/["']/g, '').trim();
    if (!primera || /^(serif|sans-serif|monospace|cursive|fantasy|system-ui|inherit|initial)$/i.test(primera)) return /^monospace$/i.test(primera) ? 'Consolas' : null;
    return primera.slice(0, 64);
  }
  // El formato que añade una etiqueta con sus atributos y su estilo
  function formatoDe(f, tag, attrs, st) {
    f = { ...f };
    if (tag === 'b' || tag === 'strong') f.b = true;
    if (tag === 'i' || tag === 'em' || tag === 'cite' || tag === 'dfn') f.i = true;
    if (tag === 'u' || tag === 'ins') f.u = 'single';
    if (tag === 's' || tag === 'strike' || tag === 'del') f.tachado = true;
    if (tag === 'sup') f.vertical = 'superscript';
    if (tag === 'sub') f.vertical = 'subscript';
    if (tag === 'code' || tag === 'tt' || tag === 'kbd' || tag === 'samp' || tag === 'pre') f.fuente = 'Consolas';
    if (/^h[4-6]$/.test(tag)) f.b = true;
    if (tag === 'font') {
      const c = color(attrs.color); if (c && c !== '000000') f.color = c;
      if (attrs.face) { const q = familia(attrs.face); if (q) f.fuente = q; }
    }
    if (st['font-weight']) { const w = st['font-weight'].toLowerCase(); f.b = w === 'bold' || w === 'bolder' || (/^\d+$/.test(w) && +w >= 600); }
    if (st['font-style']) f.i = /italic|oblique/i.test(st['font-style']);
    const deco = (st['text-decoration'] || '') + ' ' + (st['text-decoration-line'] || '');
    if (/underline/i.test(deco)) f.u = 'single'; else if (/\bnone\b/i.test(deco) && !/line-through/i.test(deco)) { if (f.u) delete f.u; }
    if (/line-through/i.test(deco)) f.tachado = true;
    if (st.color) { const c = color(st.color); if (c && c !== '000000') f.color = c; else if (c) delete f.color; }
    const t = tamPt(st['font-size']); if (t) f.tam = t;
    if (st['font-family']) { const q = familia(st['font-family']); if (q) f.fuente = q; }
    const r = resaltado(st['mso-highlight'] || st['background-color'] || (/^[^ ]*$/.test(st.background || '') ? st.background : ''));
    if (r) f.resaltado = r;
    if (st['vertical-align']) { const v = st['vertical-align'].toLowerCase(); if (v === 'super') f.vertical = 'superscript'; else if (v === 'sub') f.vertical = 'subscript'; else if (v === 'baseline') delete f.vertical; }
    for (const k of Object.keys(f)) if (f[k] === false) delete f[k];
    return f;
  }

  // ───────── Leer ─────────
  function atributos(s) {
    const a = {}, re = /([a-zA-Z_:][\w:.-]*)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>]+)))?/g;
    let m;
    while ((m = re.exec(s))) a[m[1].toLowerCase()] = decodificar(m[2] !== undefined ? m[2] : m[3] !== undefined ? m[3] : m[4] || '');
    return a;
  }
  const mismoFormato = (a, b) => JSON.stringify(a) === JSON.stringify(b);

  function leer(html) {
    html = String(html || '');
    const ini = html.indexOf('<!--StartFragment-->'), fin = html.indexOf('<!--EndFragment-->');
    if (ini >= 0) html = html.slice(ini + 20, fin > ini ? fin : undefined);
    const salida = [];
    const pila = [{ tag: '#', f: {}, oculta: 0, pre: false, lista: null }];
    let cur = null, glifo = null, ocultas = 0;
    const tope = () => pila[pila.length - 1];
    const nuevo = () => ({ texto: '', tramos: [], jc: null, estilo: null, lista: null, contenido: false });
    function agregar(t, f) {
      if (!t) return;
      if (!cur) cur = nuevo();
      cur.contenido = true;
      const d = cur.texto.length;
      cur.texto += t;
      const u = cur.tramos[cur.tramos.length - 1];
      if (u && u.h === d && mismoFormato(u.f, f)) u.h = cur.texto.length; else cur.tramos.push({ d, h: cur.texto.length, f: { ...f } });
    }
    function cerrar() {
      if (!cur) return;
      const p = cur; cur = null;
      if (!p.contenido) return;
      // los espacios de los extremos no cuentan; un párrafo de solo espacios duros es un párrafo vacío
      let t = p.texto;
      const lead = /^ +/.exec(t), tail = / +$/.exec(t);
      let cortaIni = lead ? lead[0].length : 0, cortaFin = tail && tail.index >= cortaIni ? tail[0].length : 0;
      if (/^[\s ]*$/.test(t.slice(cortaIni, t.length - cortaFin).replace(/ /g, ' '))) { p.texto = ''; p.tramos = []; }
      else if (cortaIni || cortaFin) {
        p.texto = t.slice(cortaIni, t.length - cortaFin);
        p.tramos = p.tramos.map((x) => ({ d: Math.max(0, x.d - cortaIni), h: Math.min(p.texto.length, x.h - cortaIni), f: x.f })).filter((x) => x.h > x.d);
      }
      delete p.contenido;
      salida.push(p);
    }
    const abrirParrafo = (tag, a, st, ctx) => {
      cerrar();
      cur = nuevo();
      const al = (st['text-align'] || a.align || '').toLowerCase();
      cur.jc = al === 'center' ? 'center' : al === 'right' || al === 'end' ? 'right' : al === 'justify' ? 'both' : null;
      if (tag === 'h1' || tag === 'h2' || tag === 'h3') cur.estilo = 'Heading' + tag[1];
      if (tag === 'li') {
        // el dueño y la profundidad de listas viajan en cada elemento de la pila: recorrerla para cada <li> era cuadrático con miles de etiquetas sin cerrar
        const t = tope();
        cur.lista = { tipo: t.dueno === 'ol' ? 'numeros' : 'vinetas', nivel: Math.min(8, Math.max(0, (t.niveles || 0) - 1)) };
      }
      const ml = /level(\d+)/.exec(st['mso-list'] || '');
      if (ml) cur.lista = { tipo: 'vinetas', nivel: Math.min(8, Math.max(0, +ml[1] - 1)), glifo: true };
      void ctx;
    };
    const re = /<!--([\s\S]*?)-->|<!\[([^\]]*)\]>|<(\/?)([a-zA-Z][\w:-]*)((?:"[^"]*"|'[^']*'|[^>"'<])*)>|([^<]+)|</g;
    let m;
    while ((m = re.exec(html))) {
      if (m[1] !== undefined || m[2] !== undefined) {
        // los marcadores de Office que encierran la viñeta o el número de mentira de una lista
        const c = (m[1] !== undefined ? m[1] : m[2]).replace(/^\[|\]$/g, '').trim();
        if (/^if !supportLists/i.test(c)) glifo = '';
        else if (/^endif/i.test(c) && glifo !== null) {
          if (cur && cur.lista && cur.lista.glifo) { cur.lista.tipo = /^(\d+[.)]?|[a-zA-Z][.)]|[ivxlcdmIVXLCDM]+[.)])$/.test(glifo.replace(/[\s ]/g, '')) ? 'numeros' : 'vinetas'; delete cur.lista.glifo; }
          glifo = null;
        }
        continue;
      }
      if (m[4] !== undefined) {
        const cierre = m[3] === '/', tag = m[4].toLowerCase();
        if (!cierre) {
          const a = atributos(m[5] || ''), st = estilos(a.style);
          if (tag === 'br') { if (!ocultas && glifo === null) { if (!cur) cur = nuevo(); agregar(SUAVE, tope().f); } continue; }
          if (VACIAS.has(tag)) continue;
          const sup = tope();
          const oculta = OCULTAS.has(tag) || /display\s*:\s*none/i.test(a.style || '') || st['mso-hide'] === 'all';
          if (/\/\s*$/.test(m[5] || '')) continue;
          if (BLOQUE.has(tag)) { if (tag === 'ul' || tag === 'ol') cerrar(); else abrirParrafo(tag, a, st); }
          const esLista = tag === 'ul' || tag === 'ol';
          pila.push({ tag, f: formatoDe(sup.f, tag, a, st), pre: sup.pre || tag === 'pre', dueno: esLista ? tag : sup.dueno, niveles: (sup.niveles || 0) + (esLista ? 1 : 0) });
          if (oculta) ocultas++, pila[pila.length - 1].oculta = true;
        } else {
          let i = pila.length - 1;
          while (i > 0 && pila[i].tag !== tag) i--;
          if (i <= 0) continue;
          while (pila.length > i) { const q = pila.pop(); if (q.oculta) ocultas--; if (BLOQUE.has(q.tag)) cerrar(); }
        }
        continue;
      }
      if (m[6] !== undefined) {
        if (ocultas) continue;
        let t = decodificar(m[6]);
        if (glifo !== null) { glifo += t; continue; }
        const f = tope().f;
        if (tope().pre) {
          const lineas = t.replace(/\r\n?/g, '\n').split('\n');
          lineas.forEach((l, k) => { if (k > 0) { const lista = cur && cur.lista, jc = cur && cur.jc; cerrar(); cur = nuevo(); cur.lista = lista; cur.jc = jc; cur.contenido = true; } agregar(l, f); });
        } else {
          t = t.replace(/[\t\r\n ]+/g, ' ');
          if (cur && /( |^)$/.test(cur.texto) && t[0] === ' ') t = t.slice(1);
          else if (!cur && t[0] === ' ') t = t.slice(1);
          if (t) agregar(t, f);
        }
      }
    }
    cerrar();
    return salida;
  }

  // ───────── Escribir ─────────
  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  function estiloCss(f) {
    const css = [];
    if (f.b) css.push('font-weight:bold');
    if (f.i) css.push('font-style:italic');
    const deco = [f.u && f.u !== 'none' ? 'underline' : '', f.tachado ? 'line-through' : ''].filter(Boolean).join(' ');
    if (deco) css.push('text-decoration:' + deco);
    if (f.color && /^[0-9a-f]{6}$/i.test(f.color)) css.push('color:#' + f.color.toLowerCase());
    if (f.tam) css.push('font-size:' + f.tam + 'pt');
    if (f.fuente) css.push('font-family:' + JSON.stringify(String(f.fuente).replace(/["<>&]/g, '')).replace(/"/g, "'"));
    if (f.resaltado && RESALTES[f.resaltado]) css.push('background-color:rgb(' + RESALTES[f.resaltado].join(',') + ')');
    if (f.vertical === 'superscript') css.push('vertical-align:super');
    if (f.vertical === 'subscript') css.push('vertical-align:sub');
    return css.join(';');
  }
  function escribir(parrafos) {
    const sal = [], abiertas = [];
    const cerrarListas = (n) => { while (abiertas.length > n) sal.push('</' + abiertas.pop() + '>'); };
    for (const p of parrafos) {
      let cuerpo = '';
      for (const x of p.tramos && p.tramos.length ? p.tramos : [{ d: 0, h: p.texto.length, f: {} }]) {
        const t = p.texto.slice(x.d, x.h);
        const html = t.split(SUAVE).map((l) => esc(l).replace(/\t/g, '&emsp;')).join('<br>');
        const css = estiloCss(x.f || {});
        cuerpo += css ? `<span style="${css}">${html}</span>` : html;
      }
      if (!cuerpo) cuerpo = '<br>';
      const al = p.jc ? `text-align:${p.jc === 'both' ? 'justify' : p.jc}` : '';
      if (p.lista) {
        const tag = p.lista.tipo === 'numeros' ? 'ol' : 'ul', nivel = Math.min(8, Math.max(0, p.lista.nivel || 0));
        cerrarListas(nivel + 1);
        if (abiertas.length === nivel + 1 && abiertas[nivel] !== tag) cerrarListas(nivel);
        while (abiertas.length < nivel + 1) { abiertas.push(tag); sal.push('<' + tag + '>'); }
        sal.push(`<li${al ? ` style="${al}"` : ''}>${cuerpo}</li>`);
        continue;
      }
      cerrarListas(0);
      const h = /^Heading([1-3])$/.exec(p.estilo || '');
      const t = h ? 'h' + h[1] : 'p';
      sal.push(`<${t}${al ? ` style="${al}"` : ''}>${cuerpo}</${t}>`);
    }
    cerrarListas(0);
    return sal.join('');
  }

  globalThis.HtmlDoc = { leer, escribir, decodificar };
})();

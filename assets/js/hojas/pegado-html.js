/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Pegar desde otro programa (Excel, Google Sheets, una página web) con su formato: lo que Excel hace con el HTML del portapapeles, medido con pruebas/oraculo/excel-pegado-html.ps1. Se lee en un documento inerte
// (DOMParser: no corre scripts ni carga imágenes); lo que llega al libro son textos y números, y los colores, fuentes y formatos pasan por el escritor de estilos, que los escapa.
//   PegadoHtml.leer(html) → { filas: [{ celdas: [{ c, texto, cambios, enlace, valor, formato, fin }] }], combinadas: [rango relativo], alto, ancho }
//   PegadoHtml.planear(ed, h, seleccion, html, { idioma }) → plan de PegadoEspecial (.celdas, .destino, .objetos) | { error } | null (no trae nada)
(function () {
  const MAX_CELDAS = 200000, MAX_SPAN = 1000, ESQUEMAS = /^(https?:|mailto:|ftp:)/i;
  const NOMBRES = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', navy: '000080', gray: '808080', grey: '808080', silver: 'C0C0C0', maroon: '800000', olive: '808000', purple: '800080', teal: '008080', aqua: '00FFFF', cyan: '00FFFF', fuchsia: 'FF00FF', magenta: 'FF00FF', lime: '00FF00', orange: 'FFA500', windowtext: '000000', window: 'FFFFFF' };
  function color(t) {
    const s = String(t || '').trim().toLowerCase();
    if (!s || s === 'transparent' || s === 'none' || s === 'inherit' || s === 'auto' || s === 'initial') return null;
    if (NOMBRES[s]) return NOMBRES[s];
    let m = /^#([0-9a-f]{3})$/.exec(s);
    if (m) return [...m[1]].map((x) => x + x).join('').toUpperCase();
    m = /^#([0-9a-f]{6})$/.exec(s);
    if (m) return m[1].toUpperCase();
    m = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)(?:[\s,/]+([\d.]+%?))?\s*\)$/.exec(s);
    if (m) { if (m[4] !== undefined && parseFloat(m[4]) === 0) return null; return [m[1], m[2], m[3]].map((x) => Math.min(255, +x).toString(16).padStart(2, '0')).join('').toUpperCase(); }
    return undefined;
  }
  const FORMATOS_MSO = { 'general': null, 'short date': 'mm-dd-yy', 'medium date': 'd-mmm-yy', 'long date': 'd-mmm-yy', 'short time': 'h:mm', 'medium time': 'h:mm AM/PM', 'long time': 'h:mm:ss AM/PM', 'percent': '0.00%', 'scientific': '0.00E+00', 'standard': '#,##0.00', 'fixed': '0.00' };
  function formatoMso(v) {
    let t = String(v || '').trim();
    if (/^".*"$|^'.*'$/.test(t)) t = t.slice(1, -1);
    const n = t.toLowerCase();
    if (n in FORMATOS_MSO) return FORMATOS_MSO[n];
    return t.replace(/\\(.)/g, (m, c) => (/[.,#0%@/:\- ]/.test(c) ? c : m)).slice(0, 200) || null;
  }
  // ancho → grosor de borde de Excel: medido (.5pt fino, 1px y 2px medio, 3px grueso; la raya también cambia con el grueso)
  function borde(valor) {
    const s = String(valor || '').toLowerCase();
    if (!s || /\b(none|hidden)\b/.test(s)) return null;
    const ancho = /(\d*\.?\d+)\s*(px|pt)/.exec(s), pt = ancho ? (ancho[2] === 'px' ? +ancho[1] * 0.75 : +ancho[1]) : /\bthin\b/.test(s) ? 0.5 : /\bthick\b/.test(s) ? 2.25 : 0.75;
    if (pt <= 0) return null;
    const c = (s.match(/#[0-9a-f]{3,6}\b|rgba?\([^)]*\)|\b[a-z]+\b/g) || []).map(color).find((x) => x);
    const clase = /\bdouble\b/.test(s) ? 'double' : /\bdashed\b/.test(s) ? (pt >= 0.75 ? 'mediumDashed' : 'dashed') : /\bdotted\b/.test(s) ? 'dotted' : pt < 0.75 ? 'thin' : pt >= 2 ? 'thick' : 'medium';
    return { estilo: clase, color: c || null };
  }
  function declaraciones(t) {
    const o = {};
    for (const p of String(t || '').split(';')) { const k = p.indexOf(':'); if (k > 0) o[p.slice(0, k).trim().toLowerCase()] = p.slice(k + 1).replace(/\s*!important\s*$/i, '').trim(); }
    return o;
  }
  function leerHojaEstilos(doc) {
    const reglas = [];
    for (const st of doc.querySelectorAll('style')) {
      const t = st.textContent.replace(/<!--|-->/g, '').replace(/\/\*[\s\S]*?\*\//g, '').slice(0, 2000000);
      for (const m of t.matchAll(/([^{}@]+)\{([^{}]*)\}/g)) {
        const d = declaraciones(m[2]);
        for (const sel of m[1].split(',')) {
          const q = /^\s*([a-zA-Z][\w-]*)?(?:\.([\w-]+))?(?:#([\w-]+))?\s*$/.exec(sel);
          if (q && (q[1] || q[2] || q[3])) reglas.push({ tag: q[1] ? q[1].toLowerCase() : null, clase: q[2] || null, id: q[3] || null, d });
        }
      }
      if (reglas.length > 5000) break;
    }
    return (el) => {
      let o = {};
      const clases = (el.getAttribute('class') || '').split(/\s+/);
      for (const peso of [0, 1, 2]) for (const r of reglas) {
        const p = r.id ? 2 : r.clase ? 1 : 0;
        if (p !== peso) continue;
        if ((r.tag && r.tag !== el.tagName.toLowerCase()) || (r.clase && !clases.includes(r.clase)) || (r.id && r.id !== el.getAttribute('id'))) continue;
        o = Object.assign(o, r.d);
      }
      return Object.assign(o, declaraciones(el.getAttribute('style')));
    };
  }

  function aCambios(css, base) {
    const c = Object.assign({}, base);
    const fw = css['font-weight'];
    if (fw) { if (/bold/.test(fw) || +fw >= 600) c.negrita = true; else if (fw === 'normal' || +fw > 0) delete c.negrita; }
    if (css['font-style']) { if (/italic|oblique/.test(css['font-style'])) c.cursiva = true; else delete c.cursiva; }
    const td = css['text-decoration'] || css['text-decoration-line'];
    if (td) { if (/underline/.test(td)) c.subrayado = true; if (/line-through/.test(td)) c.tachado = true; if (/none/.test(td)) { delete c.subrayado; delete c.tachado; } }
    if (css.color) { const k = color(css.color); if (k) c.colorFuente = k; }
    const fs = /^(\d*\.?\d+)\s*(pt|px)$/.exec(css['font-size'] || '');
    if (fs) { const pt = fs[2] === 'px' ? +fs[1] * 0.75 : +fs[1]; if (pt >= 1 && pt <= 409) c.tamano = pt; }
    if (css['font-family']) { const f = css['font-family'].split(',')[0].trim().replace(/^["']|["']$/g, ''); if (f && !/^(sans-serif|serif|monospace|inherit|initial|cursive|fantasy|system-ui)$/i.test(f)) c.fuente = f.slice(0, 31); }
    if (css['background-color'] || css.background) { const k = color(css['background-color'] || (css.background || '').split(/\s+(?![^(]*\))/)[0]); if (k) c.relleno = k; else if (k === null) delete c.relleno; }
    const ta = (css['text-align'] || '').toLowerCase();
    if (/^(left|center|right)$/.test(ta)) c.h = ta; else if (ta === 'general' || ta === 'start') delete c.h;
    const va = (css['vertical-align'] || '').toLowerCase();
    if (va === 'top') c.v = 'top'; else if (va === 'middle') c.v = 'center'; else if (va === 'bottom') delete c.v;
    if (css['white-space']) { if (/nowrap|pre\b/.test(css['white-space'])) delete c.ajustar; else c.ajustar = true; }
    return c;
  }
  const LADOS = [['left', 'izq'], ['right', 'der'], ['top', 'arr'], ['bottom', 'aba']];
  function bordesDe(css, c) {
    let bo = c.borde ? Object.assign({}, c.borde) : {};
    if (css.border) { const b = borde(css.border); for (const [, k] of LADOS) { if (b) bo[k] = b; else delete bo[k]; } }
    for (const [n, k] of LADOS) if (css['border-' + n]) { const b = borde(css['border-' + n]); if (b) bo[k] = b; else delete bo[k]; }
    if (Object.keys(bo).length) c.borde = bo; else delete c.borde;
  }
  const TAM_FONT = { 1: 7.5, 2: 10, 3: 12, 4: 13.5, 5: 18, 6: 24, 7: 36 };
  const EN_LINEA = (cssDe, el, c0) => {   // estilo de lo que hay dentro de una celda: se juntan todos los tramos, como hace Excel
    const c = Object.assign({}, c0);
    let enlace = null;
    const ver = (n, prof) => {
      if (prof > 60) return;
      for (const k of n.children) {
        const t = k.tagName.toLowerCase();
        if (t === 'table') continue;
        if (t === 'b' || t === 'strong') c.negrita = true;
        if (t === 'i' || t === 'em') c.cursiva = true;
        if (t === 'u' || t === 'ins') c.subrayado = true;
        if (t === 's' || t === 'strike' || t === 'del') c.tachado = true;
        if (t === 'a' && k.getAttribute('href') && !enlace) {
          const u = k.getAttribute('href').trim();
          if (ESQUEMAS.test(u)) { enlace = u.slice(0, 2000); c.subrayado = true; c.colorFuente = '467886'; }
        }
        if (t === 'font') {
          const co = color(k.getAttribute('color')); if (co) c.colorFuente = co;
          const fa = k.getAttribute('face'); if (fa) c.fuente = fa.split(',')[0].trim().slice(0, 31);
          const s = TAM_FONT[+k.getAttribute('size')]; if (s) c.tamano = s;
        }
        const u = aCambios(cssDe(k), c);
        const conservar = t === 'a' && enlace ? { subrayado: true, colorFuente: c.colorFuente } : null;
        for (const p in u) if (t === 'span' || t === 'font' || t === 'a' || u[p] !== undefined) c[p] = u[p];
        if (conservar) Object.assign(c, conservar);
        ver(k, prof + 1);
      }
    };
    ver(el, 0);
    return { c, enlace };
  };
  // los renglones de una celda: <br> y los párrafos los separan (entre párrafos queda uno vacío), salvo mso-data-placement:same-cell
  function renglones(el, cssDe) {
    const l = [];
    let actual = '';
    const cortar = () => { l.push(actual.replace(/\s+/g, ' ').trim()); actual = ''; };
    const ver = (n, prof) => {
      if (prof > 60) return;
      for (const k of n.childNodes) {
        if (k.nodeType === 3) { actual += k.textContent; continue; }
        if (k.nodeType !== 1) continue;
        const t = k.tagName.toLowerCase();
        if (t === 'style' || t === 'script' || t === 'img') continue;
        if (t === 'br') { if (/same-cell/.test(cssDe(k)['mso-data-placement'] || '')) actual += '\n'; else cortar(); continue; }
        if (t === 'table') { if (actual.trim()) cortar(); else actual = ''; for (const tr of k.querySelectorAll('tr')) { actual = [...tr.children].map((td) => td.textContent).join(' '); cortar(); } continue; }
        if (t === 'p' || t === 'div' || t === 'li') {
          if (actual.trim()) cortar(); else actual = '';
          if (t === 'p' && l.length) l.push('');
          ver(k, prof + 1);
          if (actual.trim()) cortar(); else actual = '';
          continue;
        }
        ver(k, prof + 1);
      }
    };
    ver(el, 0);
    if (actual.trim() || !l.length) cortar();
    return l;
  }
  const SIN_ESPACIOS = (s) => s.split('\n').map((x) => x.replace(/\s+/g, ' ').trim()).join('\n');

  function leer(html) {
    const doc = new DOMParser().parseFromString(String(html), 'text/html');
    const cssDe = leerHojaEstilos(doc);
    const salida = [], combinadas = [];   // salida: { f, c, celda }
    let fila = 0, ancho = 0;
    const heredar = (cadena) => cadena.reduce((a, el) => { const c = aCambios(cssDe(el), a); const bg = el.getAttribute('bgcolor'); if (bg && color(bg)) c.relleno = color(bg); return c; }, {});
    function tabla(t) {
      const trs = [...t.querySelectorAll('tr')].filter((tr) => tr.closest('table') === t);   // en el orden del documento: el tfoot donde esté
      const cap = t.querySelector(':scope > caption');
      const baseT = heredar([t]);
      const largoT = trs.reduce((a, tr) => Math.max(a, [...tr.cells].reduce((x, td) => x + Math.min(MAX_SPAN, +td.colSpan || 1), 0)), 1);
      if (cap) {
        const texto = cap.textContent.replace(/\s+/g, ' ').trim();
        salida.push({ f: fila, c: 0, texto, cambios: Object.assign({ h: 'center', v: 'center' }, heredar([t, cap])), enlace: null });
        for (let c = 1; c < largoT; c++) salida.push({ f: fila, c, texto: '', cambios: {}, enlace: null });
        if (largoT > 1) combinadas.push({ f0: fila, c0: 0, f1: fila, c1: largoT - 1 });
        fila++;
      }
      // cuadrícula: quién ocupa qué, con el alto real de cada fila (sus renglones)
      const ocup = new Map(), filasT = [];
      let cuenta = 0;
      trs.forEach((tr, i) => {
        let c = 0;
        const celdas = [];
        for (const td of tr.cells) {
          while (ocup.has(i + ',' + c)) c++;
          const cs = Math.min(MAX_SPAN, +td.colSpan || 1), rs = Math.min(MAX_SPAN, +td.rowSpan || 1);
          cuenta += cs * rs;
          if (cuenta > MAX_CELDAS) throw new RangeError('grande');   // un HTML hostil no cuelga la página
          for (let a = 0; a < rs; a++) for (let b = 0; b < cs; b++) ocup.set((i + a) + ',' + (c + b), 1);
          celdas.push({ td, c, cs, rs, lineas: renglones(td, cssDe) });
          ancho = Math.max(ancho, c + cs);
          c += cs;
        }
        filasT.push(celdas);
      });
      const altos = filasT.map((l) => l.reduce((a, x) => (x.rs === 1 ? Math.max(a, x.lineas.length) : a), 1));
      const inicio = []; let acum = fila;
      for (const a of altos) { inicio.push(acum); acum += a; }
      filasT.forEach((l, i) => {
        const tr = trs[i];
        for (const x of l) {
          const td = x.td, cadena = [t];
          if (tr.parentElement && tr.parentElement !== t) cadena.push(tr.parentElement);
          cadena.push(tr, td);
          let c = heredar(cadena.slice(0, -1));
          c.v = 'center'; c.ajustar = true;
          if (td.tagName === 'TH') { c.negrita = true; c.h = 'center'; }
          c = aCambios(cssDe(td), c);
          const bg = td.getAttribute('bgcolor'); if (bg && color(bg)) c.relleno = color(bg);
          const al = (td.getAttribute('align') || '').toLowerCase(); if (/^(left|center|right)$/.test(al)) c.h = al;
          const val = (td.getAttribute('valign') || '').toLowerCase(); if (val === 'top') c.v = 'top'; else if (val === 'bottom') delete c.v;
          const cssT = cssDe(td);
          bordesDe(cssT, c);
          const adentro = EN_LINEA(cssDe, td, c);
          const f0 = inicio[i], altoTotal = altos.slice(i, i + x.rs).reduce((a, b) => a + b, 0);
          let formato;
          if (cssT['mso-number-format']) formato = formatoMso(cssT['mso-number-format']);
          let valor;
          const dsv = td.getAttribute('data-sheets-value');
          if (dsv) { try { const j = JSON.parse(dsv); if (j[1] === 3 && Number.isFinite(j[3])) valor = { tipo: 'numero', valor: j[3] }; else if (j[1] === 2 && typeof j[2] === 'string') valor = { tipo: 'texto', valor: j[2] }; else if (j[1] === 4) valor = { tipo: 'logico', valor: !!j[4] }; } catch (e) { /* no era JSON */ } }
          const dsf = td.getAttribute('data-sheets-numberformat');
          if (dsf) { try { const j = JSON.parse(dsf); if (typeof j[2] === 'string') formato = j[2].slice(0, 200); } catch (e) { /* no era JSON */ } }
          x.lineas.forEach((texto, k) => { salida.push({ f: f0 + k, c: x.c, texto: SIN_ESPACIOS(texto), cambios: adentro.c, enlace: k === 0 ? adentro.enlace : null, valor: k === 0 ? valor : undefined, formato }); });
          const usadas = x.lineas.length;
          // en una combinada, cada celda con los bordes de su lado: los de adentro no se ven
          const conLados = (a, b) => {
            if (!adentro.c.borde || (x.cs === 1 && altoTotal === 1)) return adentro.c;
            const bo = Object.assign({}, adentro.c.borde);
            if (b > 0) delete bo.izq;
            if (b < x.cs - 1) delete bo.der;
            if (a > 0) delete bo.arr;
            if (a < altoTotal - 1) delete bo.aba;
            const q = Object.assign({}, adentro.c);
            if (Object.keys(bo).length) q.borde = bo; else delete q.borde;
            return q;
          };
          if (x.cs > 1 || altoTotal > usadas) salida[salida.length - 1].cambios = conLados(usadas - 1, 0);
          const relleno = (a, b) => ({ texto: '', cambios: conLados(a, b), enlace: null });
          for (let a = usadas; a < altoTotal; a++) salida.push(Object.assign({ f: f0 + a, c: x.c }, relleno(a, 0)));
          for (let b = 1; b < x.cs; b++) for (let a = 0; a < altoTotal; a++) salida.push(Object.assign({ f: f0 + a, c: x.c + b }, relleno(a, b)));
          if (x.cs > 1 || altoTotal > usadas) {
            if (usadas === 1 || x.cs > 1) combinadas.push({ f0: f0 + usadas - 1, c0: x.c, f1: f0 + altoTotal - 1, c1: x.c + x.cs - 1 });
            else combinadas.push({ f0: f0 + usadas - 1, c0: x.c, f1: f0 + altoTotal - 1, c1: x.c });
          }
        }
      });
      fila = acum;
    }
    let bloquePrevio = null, enLinea = null;
    const renglonVacio = () => { salida.push({ f: fila, c: 0, texto: '', cambios: {}, enlace: null }); fila++; };
    const cerrarEnLinea = () => {
      if (!enLinea) return;
      const texto = enLinea.texto.replace(/\s+/g, ' ').trim();
      if (texto) { salida.push({ f: fila, c: 0, texto, cambios: enLinea.c, enlace: enLinea.enlace }); ancho = Math.max(ancho, 1); fila++; bloquePrevio = 'texto'; }
      enLinea = null;
    };
    const parrafo = (el, cambios) => {
      const r = renglones(el, cssDe), a = EN_LINEA(cssDe, el, cambios);
      r.forEach((t, k) => { if (t !== '' || k) { salida.push({ f: fila, c: 0, texto: t, cambios: a.c, enlace: k === 0 ? a.enlace : null }); } fila++; });
      ancho = Math.max(ancho, 1);
    };
    const flujo = (raiz, prof) => {
      if (prof > 60) return;
      for (const n of raiz.childNodes) {
        if (n.nodeType === 3) { if (n.textContent.trim() || enLinea) { if (!enLinea) enLinea = { texto: '', c: {}, enlace: null }; enLinea.texto += n.textContent; } continue; }
        if (n.nodeType !== 1) continue;
        const t = n.tagName.toLowerCase();
        if (t === 'style' || t === 'script' || t === 'head' || t === 'meta' || t === 'link' || t === 'title') continue;
        if (t === 'table') {
          cerrarEnLinea();
          if (bloquePrevio === 'p') renglonVacio();
          tabla(n);
          bloquePrevio = 'tabla';
        } else if (t === 'p' || /^h[1-6]$/.test(t) || t === 'div' || t === 'blockquote' || t === 'pre') {
          cerrarEnLinea();
          if (n.querySelector('table')) { flujo(n, prof + 1); continue; }
          if (t !== 'div' && bloquePrevio) renglonVacio();
          else if (t === 'div' && (bloquePrevio === 'p' || bloquePrevio === 'tabla')) renglonVacio();
          const base = aCambios(cssDe(n), /^h[1-6]$/.test(t) ? { negrita: true } : {});
          parrafo(n, base);
          bloquePrevio = t === 'div' ? 'texto' : 'p';
        } else if (t === 'ul' || t === 'ol' || t === 'dl') {
          cerrarEnLinea();
          for (const li of n.children) parrafo(li, Object.assign({ h: 'left', v: 'center' }, aCambios(cssDe(li), {})));
          bloquePrevio = 'lista';
        } else if (t === 'br') {
          cerrarEnLinea();
        } else if (t === 'google-sheets-html-origin' || t === 'tbody' || t === 'section' || t === 'article' || t === 'main' || t === 'body' || t === 'html') {
          flujo(n, prof + 1);
        } else {
          // lo que va en línea, juntos en una celda: bold, span, a…
          if (!enLinea) enLinea = { texto: '', c: {}, enlace: null };
          enLinea.texto += n.textContent;
          const envuelve = document.createElement ? n.ownerDocument.createElement('div') : null;
          if (envuelve) { envuelve.append(n.cloneNode(true)); const q = EN_LINEA(cssDe, envuelve, enLinea.c); enLinea.c = q.c; enLinea.enlace = enLinea.enlace || q.enlace; }
        }
      }
    };
    try { flujo(doc.body || doc.documentElement, 0); } catch (e) { if (e instanceof RangeError) return { celdas: [{}], combinadas: [], alto: 0, ancho: 0, truncada: true }; throw e; }
    cerrarEnLinea();
    return { celdas: salida.slice(0, MAX_CELDAS), combinadas, alto: fila, ancho, truncada: salida.length > MAX_CELDAS };
  }

  function planear(ed, h, s, html, opciones) {
    const L = leer(html);
    if (!L.celdas.length) return null;
    if (L.truncada) return { error: 'grande' };
    const D = { f0: s.f0, c0: s.c0, f1: s.f0 + Math.max(0, L.alto - 1), c1: s.c0 + Math.max(0, L.ancho - 1) };
    if (D.f1 >= 1048576 || D.c1 >= 16384) return { error: 'limite' };
    const hoja = ed.hojas[h];
    if ((hoja.combinadas || []).some((m) => m.f0 <= D.f1 && D.f0 <= m.f1 && m.c0 <= D.c1 && D.c0 <= m.c1 && !(m.f0 >= D.f0 && m.f1 <= D.f1 && m.c0 >= D.c0 && m.c1 <= D.c1))) return { error: 'combinada' };
    const memo = new Map(), base = 0;
    const estilo = (c, formato) => {
      const k = JSON.stringify(c) + '|' + (formato || '');
      if (!memo.has(k)) {
        const q = Object.assign({}, c);
        if (formato) q.formato = formato;
        memo.set(k, Object.keys(q).length && memo.size < 4000 ? ed.estilos.derivar(base, q) : 0);   // con miles de estilos distintos, el resto sin formato
      }
      return memo.get(k);
    };
    const opc = { idioma: (opciones && opciones.idioma) === 'en' ? 'en' : 'es', fecha1904: ed.libro.fecha1904 };
    const celdas = [], enlaces = [];
    const vistas = new Set();
    for (const x of L.celdas) {
      const f = s.f0 + x.f, c = s.c0 + x.c;
      if (ed.protegida(h, f, c) && !ed.deTabla(h, f, c)) return { error: 'protegida' };
      const k = f * 16384 + c;
      if (vistas.has(k)) continue;
      vistas.add(k);
      let est = estilo(x.cambios || {}, x.formato);
      let v;
      if (x.valor) v = { tipo: x.valor.tipo, valor: x.valor.valor, formula: null, estilo: est };
      else if (x.formato === '@') v = x.texto === '' ? { tipo: 'vacio', valor: null, formula: null, estilo: est } : { tipo: 'texto', valor: x.texto, formula: null, estilo: est };
      else {
        const e = EntradaCelda.interpretar(x.texto, opc);
        if (e.tipo === 'vacio') v = { tipo: 'vacio', valor: null, formula: null, estilo: est };
        else if (e.tipo === 'formula') v = { tipo: 'numero', valor: 0, formula: e.formula, estilo: est };
        else if (e.tipo === 'error') v = { tipo: 'error', valor: e.valor, formula: null, estilo: est };
        else {
          if (e.formato && e.tipo === 'numero' && !x.formato) est = ed.estilos.derivar(est, { formato: e.formato });
          v = { tipo: e.tipo, valor: e.valor, formula: null, estilo: est };
        }
      }
      if (!est && v.tipo === 'vacio') v = null;
      celdas.push({ f, c, v });
      if (x.enlace) enlaces.push({ f, c, url: x.enlace, texto: x.texto });
    }
    const poner = L.combinadas.map((m) => ({ f0: m.f0 + s.f0, c0: m.c0 + s.c0, f1: m.f1 + s.f0, c1: m.c1 + s.c0 })).filter((m) => m.f1 > m.f0 || m.c1 > m.c0);
    const quitar = (hoja.combinadas || []).filter((m) => m.f0 >= D.f0 && m.f1 <= D.f1 && m.c0 >= D.c0 && m.c1 <= D.c1);
    const hayEnlacesAbajo = (hoja.hipervinculos || []).some((m) => m.f0 >= D.f0 && m.f1 <= D.f1 && m.c0 >= D.c0 && m.c1 <= D.c1);
    const objetos = poner.length || quitar.length || enlaces.length || hayEnlacesAbajo
      ? { h0: h, h1: h, D, S: D, pares: [], cf: null, dv: false, notas: false, enlaces: false, blancos: false, comb: { quitar, poner }, anchos: [], altos: [], tablas: [], tapadas: [], formato: ed.formato, trans: false, externos: enlaces, limpiarEnlaces: hayEnlacesAbajo }
      : null;
    return { h, celdas, destino: D, objetos, tablas: MoverTablas.toca(ed, h, D) };
  }

  globalThis.PegadoHtml = { leer, planear, _color: color, _borde: borde, _formatoMso: formatoMso };
})();

/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Escritor de hojas de cálculo OpenDocument (.ods, el formato de LibreOffice y de Google Sheets) del motor propio de formatos antiguos (fase 3). Requiere zip-propio.js y, para
// llevar fórmulas, ods-formulas.js. No depende del navegador.
//
//   const { bytes, avisos } = await OdsEscritor.crearLibro(hojas, opciones);
//
// `hojas` y `opciones` son los de XlsEscritor (ver xls-escritor.js): el contrato de celdas que devuelven los lectores. opciones: { formulas: true, fecha1904: bool, nombres: [{ nombre, hoja?,
// formula }] }. avisos: [{ tipo, cantidad, ejemplo? }] con 'formula-sin-llevar' (quedó el valor), 'formato' (un formato de número que ODF no expresa igual) y 'texto-largo'.
// Un .ods lleva el mimetype primero y sin comprimir, content.xml (hojas, estilos de celda, columnas y filas), styles.xml y META-INF/manifest.xml.
(function () {
  const MAX_TEXTO = 32767;
  const ERRORES = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA']);
  const BORDES = { thin: ['0.74pt', 'solid'], medium: ['1.76pt', 'solid'], thick: ['2.49pt', 'solid'], hair: ['0.05pt', 'solid'], double: ['1.76pt', 'double'], dashed: ['0.74pt', 'dashed'], dotted: ['0.74pt', 'dotted'],
    mediumDashed: ['1.76pt', 'dashed'], dashDot: ['0.74pt', 'dashed'], mediumDashDot: ['1.76pt', 'dashed'], dashDotDot: ['0.74pt', 'dotted'], mediumDashDotDot: ['1.76pt', 'dotted'], slantDashDot: ['1.76pt', 'dashed'] };
  const esc = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const colANumero = (s) => { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
  const numeroACol = (n) => { let s = ''; for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
  function parseRango(t) {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(String(t).trim());
    if (!m) return null;
    const c0 = colANumero(m[1]), f0 = +m[2] - 1, c1 = m[3] ? colANumero(m[3]) : c0, f1 = m[4] ? +m[4] - 1 : f0;
    return f0 < 0 || f1 < f0 || c1 < c0 ? null : { f0, c0, f1, c1 };
  }
  const NOMBRE_HOJA_SIMPLE = /^[A-Za-zÀ-￿_][A-Za-z0-9_.À-￿]*$/;
  const citarHoja = (n) => (NOMBRE_HOJA_SIMPLE.test(n) ? n : "'" + n.replace(/'/g, "''") + "'");

  // El texto de un párrafo: saltos de línea, tabulaciones y espacios repetidos tienen etiquetas propias en ODF.
  function parrafos(texto) {
    return texto.split('\n').map((linea) => {
      let s = esc(linea.replace(/\r/g, '')).replace(/\t/g, '<text:tab/>');
      s = s.replace(/^ +| {2,}| +$/g, (m) => (m.length === 1 ? ' ' : (m[0] === ' ' && m.length > 1 ? ' ' : '') + `<text:s text:c="${m.length - 1}"/>`));
      return '<text:p>' + s + '</text:p>';
    }).join('');
  }

  // ---- fechas: número de serie de Excel → ISO ----------------------------------------------------------------------------------
  function isoDeSerie(serie, fecha1904) {
    const dias = Math.floor(serie + 1e-9), resto = serie - dias;
    let base = fecha1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
    let d = dias;
    if (!fecha1904 && dias < 61) d = dias + 1;   // antes del 1900-03-01 Excel cuenta un día de menos (y el 29 de febrero de 1900 no existió)
    const t = new Date(base + d * 86400000);
    const seg = Math.round(resto * 86400);
    const p = (n, l = 2) => String(n).padStart(l, '0');
    const y = t.getUTCFullYear();
    return { fecha: `${p(y, 4)}-${p(t.getUTCMonth() + 1)}-${p(t.getUTCDate())}`, hora: `${p(Math.floor(seg / 3600) % 24)}:${p(Math.floor(seg / 60) % 60)}:${p(seg % 60)}`, seg };
  }

  // ---- formatos de número: código de Excel → estilo de número de ODF ----------------------------------------------------------------------------------
  // Devuelve { tipo: 'fecha' | 'hora' | 'numero' | 'texto', xml (el cuerpo de <number:…-style>), etiqueta } o null si es «General».
  function traducirSeccion(sec) {
    const partes = []; let i = 0, hayFecha = false, hayHora = false, porcentaje = false, texto = false, entera = '', dec = '', cientifica = null, miles = false, vistoPunto = false, extra = 0, hayDigitos = false;
    const COLORES = { red: 'ff0000', black: '000000', blue: '0000ff', cyan: '00ffff', green: '00ff00', magenta: 'ff00ff', white: 'ffffff', yellow: 'ffff00' };
    const mc = /\[(Red|Black|Blue|Cyan|Green|Magenta|White|Yellow)\]/i.exec(sec), mcond = /\[(<=|>=|<>|<|>|=)\s*(-?[\d.]+)\]/.exec(sec);
    const color = mc ? COLORES[mc[1].toLowerCase()] : null, cond = mcond ? { op: mcond[1], val: mcond[2] } : null;
    const t = sec.replace(/\[(?:Red|Black|Blue|Cyan|Green|Magenta|White|Yellow|Color\s*\d+)\]/gi, '').replace(/\[[<>=]+[^\]]*\]/g, '');
    const fr = /^\s*(#|0)?\s*(\?+|#+|0+)\/(\?+|#+|0+|\d+)\s*$/.exec(t);
    if (fr) return { partes: [{ k: 'fraccion', entero: fr[1] ? (fr[1] === '0' ? 1 : 0) : 0, num: fr[2].length, den: /^\d+$/.test(fr[3]) ? null : fr[3].length, fijo: /^\d+$/.test(fr[3]) ? fr[3] : null }], hayFecha: false, hayHora: false, porcentaje: false, texto: false, entera: '', dec: '', miles: false, cientifica: null, extra: 0, color, cond };
    const lit = (s) => { if (s) partes.push({ k: 'text', s }); };
    const hora = /h/i.test(t.replace(/"[^"]*"/g, '')) || /AM\/PM|A\/P/i.test(t);
    while (i < t.length) {
      const ch = t[i], resto = t.slice(i);
      if (ch === '"') { const j = t.indexOf('"', i + 1); lit(t.slice(i + 1, j < 0 ? t.length : j)); i = j < 0 ? t.length : j + 1; continue; }
      if (ch === '\\') { lit(t[i + 1] || ''); i += 2; continue; }
      if (ch === '_') { lit(' '); i += 2; continue; }
      if (ch === '*') { i += 2; continue; }
      if (ch === '[') {
        const j = t.indexOf(']', i); const c = t.slice(i + 1, j);
        if (/^h+$/i.test(c)) { partes.push({ k: 'hours', largo: c.length > 1, elapsed: true }); hayHora = true; }
        else if (/^m+$/i.test(c)) { partes.push({ k: 'minutes', largo: c.length > 1, elapsed: true }); hayHora = true; }
        else if (/^s+$/i.test(c)) { partes.push({ k: 'seconds', largo: c.length > 1, elapsed: true }); hayHora = true; }
        else { const mm = /^\$([^-\]]*)/.exec(c); if (mm) lit(mm[1]); }
        i = j < 0 ? t.length : j + 1; continue;
      }
      let m;
      if ((m = /^(AM\/PM|A\/P)/i.exec(resto))) { partes.push({ k: 'ampm' }); hayHora = true; i += m[0].length; continue; }
      if ((m = /^(y{1,4})/i.exec(resto))) { partes.push({ k: 'year', largo: m[0].length > 2 }); hayFecha = true; i += m[0].length; continue; }
      if ((m = /^(d{1,4})/i.exec(resto))) { partes.push(m[0].length > 2 ? { k: 'dow', largo: m[0].length === 4 } : { k: 'day', largo: m[0].length === 2 }); hayFecha = true; i += m[0].length; continue; }
      if ((m = /^(m{1,5})/i.exec(resto))) {
        const n = m[0].length;
        // m minúscula: minutos si viene detrás de horas o delante de segundos
        const antes = partes.filter((p) => p.k !== 'text').pop(), despues = /^m+[^a-z]*s/i.test(resto.replace(/"[^"]*"/g, ''));
        const minutos = n <= 2 && ((antes && (antes.k === 'hours')) || despues);
        if (minutos) { partes.push({ k: 'minutes', largo: n === 2 }); hayHora = true; } else { partes.push({ k: 'month', textual: n >= 3, largo: n === 2 || n === 4 }); hayFecha = true; }
        i += n; continue;
      }
      if ((m = /^(h{1,2})/i.exec(resto))) { partes.push({ k: 'hours', largo: m[0].length === 2 }); hayHora = true; i += m[0].length; continue; }
      if ((m = /^(s{1,2})/i.exec(resto))) {
        let dp = 0; const mm = /^\.(0+)/.exec(t.slice(i + m[0].length)); if (mm) { dp = mm[1].length; i += mm[0].length; }
        partes.push({ k: 'seconds', largo: m[0].length === 2, dp }); hayHora = true; i += m[0].length; continue;
      }
      if (ch === '@') { texto = true; partes.push({ k: 'textcontent' }); i++; continue; }
      if (ch === '%') { porcentaje = true; i++; continue; }
      if (/[0#?]/.test(ch)) {
        hayDigitos = true;
        if (!partes.some((p) => p.k === 'num')) partes.push({ k: 'num' });
        if (vistoPunto) dec += ch; else entera += ch;
        i++;
        if ((m = /^E[+-]/i.exec(t.slice(i)))) { const e = /^E([+-])(0+)/i.exec(t.slice(i)); if (e) { cientifica = { signo: e[1] === '+', digitos: e[2].length }; i += e[0].length; } }
        continue;
      }
      if (ch === '.' && hayDigitos && !vistoPunto && /[0#?]/.test(t[i + 1] || '')) { vistoPunto = true; i++; continue; }
      if (ch === '.' && !hayDigitos && /[0#?]/.test(t[i + 1] || '')) { hayDigitos = true; vistoPunto = true; partes.push({ k: 'num' }); i++; continue; }
      if (ch === ',') { if (hayDigitos && /[0#?]/.test(t[i + 1] || '') && !vistoPunto) miles = true; else if (hayDigitos && !vistoPunto) extra++; else lit(','); i++; continue; }
      if (ch === '/' && hayDigitos) { extra += 100; lit('/'); i++; continue; }
      lit(ch); i++;
    }
    return { partes, hayFecha, hayHora, porcentaje, texto, entera, dec, miles, cientifica, extra, hora, color, cond, elapsed: partes.some((p) => p.elapsed) };
  }
  function xmlSeccion(s) {
    let x = '';
    for (const p of s.partes) {
      switch (p.k) {
        case 'text': x += `<number:text>${esc(p.s)}</number:text>`; break;
        case 'num': {
          const minInt = (s.entera.match(/0/g) || []).length, dp = s.dec.length, minDec = (s.dec.match(/0/g) || []).length;
          if (s.cientifica) x += `<number:scientific-number number:decimal-places="${dp}" number:min-decimal-places="${minDec}" number:min-integer-digits="${Math.max(1, minInt)}" number:min-exponent-digits="${s.cientifica.digitos}"${s.cientifica.signo ? ' number:forced-exponent-sign="true"' : ''}/>`;
          else x += `<number:number number:decimal-places="${dp}" number:min-decimal-places="${minDec}" number:min-integer-digits="${Math.max(minInt, /[0]/.test(s.entera) ? 1 : 0)}"${s.miles ? ' number:grouping="true"' : ''}/>`;
          break;
        }
        case 'fraccion': x += `<number:fraction number:min-integer-digits="${p.entero}" number:min-numerator-digits="${p.num}"${p.fijo ? ` number:denominator-value="${p.fijo}"` : ` number:min-denominator-digits="${p.den}"`}/>`; break;
        case 'textcontent': x += '<number:text-content/>'; break;
        case 'year': x += `<number:year${p.largo ? ' number:style="long"' : ''}/>`; break;
        case 'month': x += `<number:month${p.textual ? ' number:textual="true"' : ''}${p.largo ? ' number:style="long"' : ''}/>`; break;
        case 'day': x += `<number:day${p.largo ? ' number:style="long"' : ''}/>`; break;
        case 'dow': x += `<number:day-of-week${p.largo ? ' number:style="long"' : ''}/>`; break;
        case 'hours': x += `<number:hours${p.largo ? ' number:style="long"' : ''}/>`; break;
        case 'minutes': x += `<number:minutes${p.largo ? ' number:style="long"' : ''}/>`; break;
        case 'seconds': x += `<number:seconds${p.largo ? ' number:style="long"' : ''}${p.dp ? ` number:decimal-places="${p.dp}"` : ''}/>`; break;
        case 'ampm': x += '<number:am-pm/>'; break;
        default: break;
      }
    }
    if (s.porcentaje) x += '<number:text>%</number:text>';
    return x;
  }
  function estiloNumero(codigo, nombre, avisar) {
    if (!codigo || codigo === 'General') return null;
    const secs = []; { let cur = '', q = false, br = false; for (const ch of codigo) { if (ch === '"') q = !q; else if (ch === '[') br = true; else if (ch === ']') br = false; if (ch === ';' && !q && !br) { secs.push(cur); cur = ''; } else cur += ch; } secs.push(cur); }
    const ts = secs.slice(0, 3).map(traducirSeccion);
    if (secs.length > 3) avisar('formato', codigo);
    const base = ts[ts.length - 1];
    const tipo = ts.some((s) => s.texto) ? 'text' : ts.some((s) => s.hayFecha) ? 'date' : ts.some((s) => s.hayHora) ? 'time' : ts.some((s) => s.porcentaje) ? 'percentage' : 'number';
    const etiq = 'number:' + tipo + '-style';
    const cuerpoDe = (s, nom, mapas) => '<' + etiq + ' style:name="' + nom + '"' + (s.elapsed ? ' number:truncate-on-overflow="false"' : '') + '>' + (s.color ? '<style:text-properties fo:color="#' + s.color + '"/>' : '') + xmlSeccion(s) + (mapas || '') + '</' + etiq + '>';
    if (!xmlSeccion(base)) return null;
    const dato = { tipo, fecha: tipo === 'date', hora: tipo === 'time', calendario: ts.some((s) => s.hayFecha) };
    if (ts.length === 1) return { ...dato, xml: cuerpoDe(base, nombre) };
    // varias secciones: cada una menos la última va en un estilo aparte y el principal (la última) las elige con style:map
    const OPS = { '<': '&lt;', '>': '&gt;', '<=': '&lt;=', '>=': '&gt;=', '=': '=', '<>': '!=' };
    let xml = '', mapas = '';
    ts.slice(0, -1).forEach((s, k) => {
      const c = s.cond ? s.cond : ts.length === 2 ? { op: '>=', val: '0' } : k === 0 ? { op: '>', val: '0' } : { op: '<', val: '0' };
      xml += cuerpoDe(s, nombre + 'P' + k);
      mapas += '<style:map style:condition="value()' + OPS[c.op] + c.val + '" style:apply-style-name="' + nombre + 'P' + k + '"/>';
    });
    return { ...dato, xml: xml + cuerpoDe(base, nombre, mapas) };
  }

  // ---- el libro ----------------------------------------------------------------------------------------------------------------
  async function crearLibro(hojasEntrada, opciones) {
    opciones = opciones || {};
    const conFormulas = opciones.formulas !== false && typeof OdsFormulas !== 'undefined';
    const f1904 = !!opciones.fecha1904;
    const avisosMapa = new Map();
    const avisar = (tipo, ejemplo) => { const a = avisosMapa.get(tipo); if (a) a.cantidad++; else avisosMapa.set(tipo, { tipo, cantidad: 1, ...(ejemplo ? { ejemplo } : {}) }); };
    if (!Array.isArray(hojasEntrada) || !hojasEntrada.length) throw new Error('el libro necesita al menos una hoja');

    const usados = new Set();
    const nombreUnico = (n, k) => {
      let t = String(n == null ? '' : n).replace(/[\[\]:*?\/\\]/g, '_').replace(/^'+|'+$/g, '').slice(0, 31) || 'Hoja' + (k + 1), b = t, i = 2;
      while (usados.has(t.toLowerCase())) { const suf = '(' + i++ + ')'; t = b.slice(0, 31 - suf.length) + suf; }
      usados.add(t.toLowerCase()); return t;
    };
    const hojas = hojasEntrada.map((h, k) => ({ ...h, nombre: nombreUnico(h.nombre, k) }));

    // estilos de número y de celda, repartidos por identidad y por contenido
    const numeros = new Map(), estilosNumero = [];
    const numeroDe = (cod) => {
      if (cod === undefined || cod === null || cod === 'General' || cod === 0) return null;
      if (typeof cod === 'number') cod = (typeof XlsxEstilos !== 'undefined' && XlsxEstilos.FORMATOS_INTEGRADOS[cod]) || null;
      if (!cod || cod === 'General') return null;
      if (numeros.has(cod)) return numeros.get(cod);
      const r = estiloNumero(cod, 'N' + (numeros.size + 1), avisar);
      if (r) { r.nombre = 'N' + (numeros.size + 1); estilosNumero.push(r); }
      numeros.set(cod, r);
      return r;
    };
    const estilosCelda = [], mapaCelda = new Map(), cacheObj = new WeakMap();
    function estiloCeldaDe(formato, estilo) {
      const por = estilo && typeof estilo === 'object' ? (cacheObj.get(estilo) || cacheObj.set(estilo, new Map()).get(estilo)) : null;
      if (por && por.has(formato)) return por.get(formato);
      const num = numeroDe(formato), e = estilo || {};
      const clave = JSON.stringify([num && num.nombre, e]);
      let r = mapaCelda.get(clave);
      if (!r) {
        const nombre = 'ce' + (estilosCelda.length + 1);
        let celda = '', texto = '', parrafo = '';
        const r0 = e.relleno;
        const fg = typeof r0 === 'string' ? r0 : r0 && typeof r0 === 'object' ? (r0.patron === 'solid' ? r0.fg : (r0.bg || r0.fg)) : null;
        if (fg) celda += ` fo:background-color="#${fg.toLowerCase()}"`;
        if (e.bordes) for (const [k, lado] of [['izq', 'left'], ['der', 'right'], ['arr', 'top'], ['aba', 'bottom']]) { const b = e.bordes[k]; if (b && BORDES[b.e]) celda += ` fo:border-${lado}="${BORDES[b.e][0]} ${BORDES[b.e][1]} #${(b.c || '000000').toLowerCase()}"`; }
        if (e.alinea && e.alinea.ajustar) celda += ' fo:wrap-option="wrap"';
        if (e.alinea && e.alinea.v) celda += ` style:vertical-align="${{ top: 'top', center: 'middle', bottom: 'bottom', justify: 'automatic', distributed: 'automatic' }[e.alinea.v] || 'bottom'}"`;
        if (e.alinea && e.alinea.rotacion) celda += ` style:rotation-angle="${e.alinea.rotacion === 255 ? 0 : e.alinea.rotacion > 90 ? 360 - (e.alinea.rotacion - 90) : e.alinea.rotacion}"`;
        const f = e.fuente;
        if (f) {
          if (f.nombre) texto += ` style:font-name="${esc(f.nombre)}" fo:font-family="${esc(f.nombre)}"`;
          if (f.tam) texto += ` fo:font-size="${f.tam}pt"`;
          if (f.negrita) texto += ' fo:font-weight="bold"';
          if (f.cursiva) texto += ' fo:font-style="italic"';
          if (f.tachado) texto += ' style:text-line-through-style="solid"';
          if (f.subrayado) texto += ` style:text-underline-style="solid" style:text-underline-type="${f.subrayado === 'double' ? 'double' : 'single'}"`;
          if (f.color) texto += ` fo:color="#${f.color.toLowerCase()}"`;
        }
        const h = e.alinea && e.alinea.h;
        if (h && h !== 'general') parrafo += ` fo:text-align="${{ left: 'start', center: 'center', right: 'end', justify: 'justify', fill: 'start', centerContinuous: 'center', distributed: 'justify' }[h] || 'start'}"`;
        if (e.alinea && e.alinea.sangria) parrafo += ` fo:margin-left="${(e.alinea.sangria * 0.2).toFixed(2)}cm"`;
        r = { nombre, xml: `<style:style style:name="${nombre}" style:family="table-cell" style:parent-style-name="Default"${num ? ` style:data-style-name="${num.nombre}"` : ''}>` +
          (celda ? `<style:table-cell-properties${celda}/>` : '') + (parrafo ? `<style:paragraph-properties${parrafo}/>` : '') + (texto ? `<style:text-properties${texto}/>` : '') + '</style:style>', num };
        estilosCelda.push(r); mapaCelda.set(clave, r);
      }
      if (por) por.set(formato, r);
      return r;
    }

    const cmDeAncho = (w) => (((w * 7 + 5) / 96) * 2.54).toFixed(3);
    const estilosColumna = new Map(), estilosFila = new Map();
    const estiloColumna = (w) => { const k = w ? cmDeAncho(w) : ''; if (!estilosColumna.has(k)) estilosColumna.set(k, 'co' + (estilosColumna.size + 1)); return estilosColumna.get(k); };
    const estiloFila = (p) => { const k = p ? (p / 72 * 2.54).toFixed(3) : ''; if (!estilosFila.has(k)) estilosFila.set(k, 'ro' + (estilosFila.size + 1)); return estilosFila.get(k); };
    const indiceHoja = new Map(); hojas.forEach((h, k) => indiceHoja.set(h.nombre.toLowerCase(), k));

    // ---- cada hoja ----
    let cuerpo = '';
    hojas.forEach((h, ih) => {
      const filas = Array.isArray(h.filas) ? h.filas : [];
      const ocultasC = new Set(h.columnasOcultas || []), ocultasF = new Set(h.filasOcultas || []), altos = h.altosFila || {}, anchos = Array.isArray(h.anchos) ? h.anchos : [];
      const combinadas = (h.combinadas || []).map(parseRango).filter((r) => r && (r.f1 > r.f0 || r.c1 > r.c0));
      const origen = new Map(), cubiertas = new Set();
      for (const r of combinadas) { origen.set(r.f0 * 16384 + r.c0, r); for (let f = r.f0; f <= r.f1; f++) for (let c = r.c0; c <= r.c1; c++) if (f !== r.f0 || c !== r.c0) cubiertas.add(f * 16384 + c); }
      // matrices: celda de origen → rango; el resto del rango son celdas cubiertas
      const matrices = new Map();
      if (conFormulas) filas.forEach((fila, f) => (fila || []).forEach((x, c) => { if (x && typeof x === 'object' && x.formula && x.matriz) { const m = parseRango(x.matriz); if (m) { matrices.set(f * 16384 + c, m); for (let ff = m.f0; ff <= m.f1; ff++) for (let cc = m.c0; cc <= m.c1; cc++) if (ff !== f || cc !== c) cubiertas.add(ff * 16384 + cc); } } }));
      let maxC = 0; for (const fila of filas) if (fila) maxC = Math.max(maxC, fila.length);
      for (const r of combinadas) maxC = Math.max(maxC, r.c1 + 1);
      let xml = `<table:table table:name="${esc(h.nombre)}" table:style-name="${h.estado === 'oculta' || h.estado === 'muyOculta' ? 'taH' : 'ta1'}">`;
      // columnas (en tramos iguales)
      const colDe = (c) => estiloColumna(anchos[c]) + (ocultasC.has(c) ? '|h' : '');
      for (let c = 0; c < maxC;) { const k = colDe(c); let d = c; while (d + 1 < maxC && colDe(d + 1) === k) d++; const [est, h2] = k.split('|'); xml += `<table:table-column table:style-name="${est}"${d > c ? ` table:number-columns-repeated="${d - c + 1}"` : ''}${h2 ? ' table:visibility="collapse"' : ''} table:default-cell-style-name="Default"/>`; c = d + 1; }
      if (!maxC) xml += '<table:table-column table:style-name="co0" table:default-cell-style-name="Default"/>';
      const total = Math.max(filas.length, ...Object.keys(altos).map((k) => +k + 1), ...[...ocultasF].map((k) => k + 1), ...combinadas.map((r) => r.f1 + 1), 0);
      for (let f = 0; f < total;) {
        const fila = filas[f] || [];
        const vacia = !fila.some((x) => x !== undefined && x !== null) && altos[f] === undefined && !ocultasF.has(f) && ![...origen.keys()].some((k) => Math.floor(k / 16384) === f) && ![...cubiertas].some((k) => Math.floor(k / 16384) === f);
        if (vacia) { let g = f; while (g + 1 < total && !(filas[g + 1] || []).some((x) => x !== undefined && x !== null) && altos[g + 1] === undefined && !ocultasF.has(g + 1) && ![...origen.keys()].some((k) => Math.floor(k / 16384) === g + 1) && ![...cubiertas].some((k) => Math.floor(k / 16384) === g + 1)) g++; xml += `<table:table-row table:style-name="ro0"${g > f ? ` table:number-rows-repeated="${g - f + 1}"` : ''}><table:table-cell table:number-columns-repeated="${Math.max(1, maxC)}"/></table:table-row>`; f = g + 1; continue; }
        xml += `<table:table-row table:style-name="${estiloFila(altos[f])}"${ocultasF.has(f) ? ' table:visibility="collapse"' : ''}>`;
        let c = 0;
        while (c < maxC) {
          const clave = f * 16384 + c;
          const x = fila[c];
          const celda = celdaXml(x, f, c, origen.get(clave), matrices.get(clave), cubiertas.has(clave), ih);
          if (celda === null) { let d = c; while (d + 1 < maxC && fila[d + 1] === undefined && !cubiertas.has(f * 16384 + d + 1) && !origen.has(f * 16384 + d + 1)) d++; xml += `<table:table-cell${d > c ? ` table:number-columns-repeated="${d - c + 1}"` : ''}/>`; c = d + 1; }
          else { xml += celda; c++; }
        }
        xml += '</table:table-row>'; f++;
      }
      cuerpo += xml + '</table:table>';
    });

    // una celda: valor, fórmula, estilo, combinada o cubierta
    function celdaXml(x, f, c, comb, matriz, cubierta, ih) {
      if (cubierta) {
        // el valor de una celda cubierta por una matriz se guarda; el de una combinada, no
        const n = norm(x);
        return '<table:covered-table-cell' + valorAtributos(n, false).atrib + estiloAtr(n) + '>' + valorAtributos(n, false).texto + '</table:covered-table-cell>';
      }
      if (x === undefined || x === null) {
        if (!comb) return null;
        return `<table:table-cell table:number-columns-spanned="${comb.c1 - comb.c0 + 1}" table:number-rows-spanned="${comb.f1 - comb.f0 + 1}"/>`;
      }
      const n = norm(x);
      let formula = '';
      if (n.formula && conFormulas) {
        const o = OdsFormulas.aOpenFormula(n.formula);
        if (o) formula = ` table:formula="of:=${esc(o)}"`; else avisar('formula-sin-llevar', numeroACol(c) + (f + 1));
      } else if (n.formula) avisar('formula-sin-llevar', numeroACol(c) + (f + 1));
      const v = valorAtributos(n, !!formula);
      if (v.error && !formula) formula = ` table:formula="of:=${v.error}"`;   // ODF no tiene valor de error: LibreOffice lo guarda como la fórmula del error
      const span = comb ? ` table:number-columns-spanned="${comb.c1 - comb.c0 + 1}" table:number-rows-spanned="${comb.f1 - comb.f0 + 1}"` : '';
      const mx = matriz && formula ? ` table:number-matrix-columns-spanned="${matriz.c1 - matriz.c0 + 1}" table:number-matrix-rows-spanned="${matriz.f1 - matriz.f0 + 1}"` : '';
      if (v.vacio && !formula && !n.estilo && !n.formato) return null;
      return '<table:table-cell' + estiloAtr(n) + span + mx + formula + v.atrib + '>' + v.texto + '</table:table-cell>';
    }
    function norm(x) {
      if (x instanceof Date) { const t = x.getTime(); return Number.isFinite(t) ? { valor: (t - x.getTimezoneOffset() * 60000) / 86400000 + 25569, formato: 14 } : { valor: null }; }
      if (x && typeof x === 'object' && typeof x.error !== 'string') {
        let val = x.valor, formato = x.formato, estilo = x.estilo;
        if (val && typeof val === 'object' && typeof val.error !== 'string' && 'valor' in val && !(val instanceof Date)) { if (formato === undefined) formato = val.formato; if (estilo === undefined) estilo = val.estilo; val = val.valor; }
        if (val instanceof Date) { const t = val.getTime(); val = Number.isFinite(t) ? (t - val.getTimezoneOffset() * 60000) / 86400000 + 25569 : null; }
        return { valor: val === undefined ? null : val, formato, estilo, formula: typeof x.formula === 'string' && x.formula ? x.formula : null };
      }
      return { valor: x === undefined ? null : x };
    }
    function estiloAtr(n) {
      if (n.estilo === undefined && (n.formato === undefined || n.formato === null || n.formato === 'General' || n.formato === 0)) return '';
      return ` table:style-name="${estiloCeldaDe(n.formato, n.estilo).nombre}"`;
    }
    function valorAtributos(n, hayFormula) {
      const v = n.valor;
      if (v === null || v === undefined) return { vacio: true, atrib: '', texto: '' };
      if (typeof v === 'number') {
        if (!Number.isFinite(v)) return { atrib: ' office:value-type="string" office:string-value="" calcext:value-type="error"', texto: '<text:p>#NUM!</text:p>' };
        const num = numeroDe(n.formato);
        if (num && num.fecha && num.calendario !== false && (f1904 || v >= 61 || v < 0 ? true : true) && v >= 0) {
          const t = isoDeSerie(v, f1904);
          return { atrib: ` office:value-type="date" office:date-value="${t.fecha}${t.seg % 86400 || v % 1 ? 'T' + t.hora : ''}" calcext:value-type="date"`, texto: `<text:p>${t.fecha}</text:p>` };
        }
        if (num && num.hora && v >= 0) {
          const seg = Math.round(v * 86400), p = (q) => String(q).padStart(2, '0');
          return { atrib: ` office:value-type="time" office:time-value="PT${p(Math.floor(seg / 3600))}H${p(Math.floor(seg / 60) % 60)}M${p(seg % 60)}S" calcext:value-type="time"`, texto: `<text:p>${p(Math.floor(seg / 3600))}:${p(Math.floor(seg / 60) % 60)}:${p(seg % 60)}</text:p>` };
        }
        const pct = num && num.tipo === 'percentage';
        return { atrib: ` office:value-type="${pct ? 'percentage' : 'float'}" office:value="${String(v)}" calcext:value-type="${pct ? 'percentage' : 'float'}"`, texto: `<text:p>${esc(String(v))}</text:p>` };
      }
      if (typeof v === 'boolean') return { atrib: ` office:value-type="boolean" office:boolean-value="${v}" calcext:value-type="boolean"`, texto: `<text:p>${v ? 'TRUE' : 'FALSE'}</text:p>` };
      if (typeof v === 'object' && typeof v.error === 'string') {
        const e = ERRORES.has(v.error) ? v.error : '#VALUE!';
        return { atrib: ' office:value-type="string" office:string-value="" calcext:value-type="error"', texto: `<text:p>${e}</text:p>`, error: e };
      }
      let t = String(v);
      if (t.length > MAX_TEXTO) { t = t.slice(0, MAX_TEXTO); avisar('texto-largo'); }
      return { atrib: ' office:value-type="string" calcext:value-type="string"', texto: parrafos(t), vacio: t === '' };
    }
    // las celdas de error sin fórmula llevan la fórmula del error (como LibreOffice)
    // (se resuelve en celdaXml: si el valor es un error y no hay fórmula, se agrega)

    // ---- nombres definidos ----
    let nombresXml = '';
    for (const nm of opciones.nombres || []) {
      if (!nm || typeof nm.nombre !== 'string' || !nm.nombre || typeof nm.formula !== 'string' || !nm.formula || !conFormulas) continue;
      const o = OdsFormulas.aOpenFormula(nm.formula); if (!o) continue;
      const rango = /^\[(\$[^\]]*)\]$/.exec(o);
      const base = '$' + citarHoja(hojas[0].nombre) + '.$A$1';
      if (rango && /^\$[^.]*\.\$?[A-Z]+\$?\d+(:\.\$?[A-Z]+\$?\d+)?$/.test(rango[1])) nombresXml += `<table:named-range table:name="${esc(nm.nombre)}" table:base-cell-address="${esc(base)}" table:cell-range-address="${esc(rango[1].replace(/\./g, '.').replace(/:\./, ':.'))}"/>`;
      else nombresXml += `<table:named-expression table:name="${esc(nm.nombre)}" table:base-cell-address="${esc(base)}" table:expression="${esc(o)}"/>`;
    }

    // ---- content.xml ----
    const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"';
    let auto = '';
    for (const [cm, nombre] of estilosColumna) auto += `<style:style style:name="${nombre}" style:family="table-column"><style:table-column-properties fo:break-before="auto" ${cm ? `style:column-width="${cm}cm"` : ''}/></style:style>`;
    auto += '<style:style style:name="co0" style:family="table-column"><style:table-column-properties fo:break-before="auto"/></style:style>';
    for (const [cm, nombre] of estilosFila) auto += `<style:style style:name="${nombre}" style:family="table-row"><style:table-row-properties fo:break-before="auto" ${cm ? `style:row-height="${cm}cm" style:use-optimal-row-height="false"` : 'style:use-optimal-row-height="true"'}/></style:style>`;
    auto += '<style:style style:name="ro0" style:family="table-row"><style:table-row-properties fo:break-before="auto" style:use-optimal-row-height="true"/></style:style>';
    auto += '<style:style style:name="ta1" style:family="table" style:master-page-name="Default"><style:table-properties table:display="true" style:writing-mode="lr-tb"/></style:style><style:style style:name="taH" style:family="table"><style:table-properties table:display="false" style:writing-mode="lr-tb"/></style:style>';
    for (const n of estilosNumero) auto += n.xml;
    for (const e of estilosCelda) auto += e.xml;
    const content = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ${NS} office:version="1.3"><office:automatic-styles>${auto}</office:automatic-styles><office:body><office:spreadsheet>` +
      `<table:calculation-settings table:case-sensitive="false" table:automatic-find-labels="false" table:use-regular-expressions="false" table:use-wildcards="true">${f1904 ? '<table:null-date table:date-value="1904-01-01"/>' : ''}</table:calculation-settings>` +
      cuerpo + (nombresXml ? `<table:named-expressions>${nombresXml}</table:named-expressions>` : '') + '</office:spreadsheet></office:body></office:document-content>';
    const styles = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles ${NS} office:version="1.3"><office:styles><style:default-style style:family="table-cell"><style:text-properties style:font-name="Calibri" fo:font-family="Calibri" fo:font-size="11pt"/></style:default-style><style:style style:name="Default" style:family="table-cell"/></office:styles><office:automatic-styles><style:page-layout style:name="Mpm1"/></office:automatic-styles><office:master-styles><style:master-page style:name="Default" style:page-layout-name="Mpm1"/></office:master-styles></office:document-styles>`;
    const manifest = '<?xml version="1.0" encoding="UTF-8"?>\n<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/></manifest:manifest>';
    const zip = new ZipPropio();
    zip.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet', { compression: 'STORE' });
    zip.file('content.xml', content);
    zip.file('styles.xml', styles);
    zip.file('META-INF/manifest.xml', manifest);
    const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    return { bytes, avisos: [...avisosMapa.values()] };
  }

  globalThis.OdsEscritor = { crearLibro };
})();

/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Escritor de hojas de cálculo OpenDocument (.ods, el formato de LibreOffice y de Google Sheets) del motor propio de formatos antiguos (fase 3). Requiere zip-propio.js y, para
// llevar fórmulas, ods-formulas.js. No depende del navegador.
//
//   const { bytes, avisos } = await OdsEscritor.crearLibro(hojas, opciones);
//
// `hojas` y `opciones` son los de XlsEscritor (ver xls-escritor.js): el contrato de celdas que devuelven los lectores. opciones: { formulas: true, fecha1904: bool, nombres: [{ nombre, hoja?,
// formula }] }. avisos: [{ tipo, cantidad, ejemplo? }] con 'formula-sin-llevar' (quedó el valor), 'formato' (un formato de número que ODF no expresa igual), 'texto-largo'
// y 'combinadas' (combinadas que pasan del tope de MAX_FILAS_AREAS: quedan separadas).
// Un .ods lleva el mimetype primero y sin comprimir, content.xml (hojas, estilos de celda, columnas y filas), styles.xml y META-INF/manifest.xml.
(function () {
  const MAX_TEXTO = 32767, MAX_FILAS = 1048576, MAX_COLS = 16384;
  const MAX_FILAS_AREAS = 8000000;   // suma de las filas que abarcan las combinadas y matrices de todo el libro (cada fila de un área se mira al escribir esa fila)
  const ERRORES = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA']);
  const BORDES = { thin: ['0.74pt', 'solid'], medium: ['1.76pt', 'solid'], thick: ['2.49pt', 'solid'], hair: ['0.05pt', 'solid'], double: ['1.76pt', 'double'], dashed: ['0.74pt', 'dashed'], dotted: ['0.74pt', 'dotted'],
    mediumDashed: ['1.76pt', 'dashed'], dashDot: ['0.74pt', 'dashed'], mediumDashDot: ['1.76pt', 'dashed'], dashDotDot: ['0.74pt', 'dotted'], mediumDashDotDot: ['1.76pt', 'dotted'], slantDashDot: ['1.76pt', 'dashed'] };
  // (los caracteres de control que XML 1.0 no admite se quitan: con uno solo, LibreOffice no abre el archivo)
  const esc = (t) => String(t).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const colANumero = (s) => { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
  const numeroACol = (n) => { let s = ''; for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
  function parseRango(t) {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(String(t).trim());
    if (!m) return null;
    const c0 = colANumero(m[1]), f0 = +m[2] - 1, c1 = m[3] ? colANumero(m[3]) : c0, f1 = m[4] ? +m[4] - 1 : f0;
    return f0 < 0 || f1 < f0 || c1 < c0 ? null : { f0, c0, f1, c1 };
  }
  // Las columnas de una fila que tienen algo (ni undefined ni null), en orden. Una fila ancha y rala (una celda en XFD) es un arreglo disperso: se recorren sus índices,
  // no sus 16.384 lugares.
  function presentes(fila) {
    const r = [];
    if (fila.length <= 1024) { for (let c = 0; c < fila.length; c++) if (fila[c] !== undefined && fila[c] !== null) r.push(c); return r; }
    for (const k in fila) { const c = +k; if (Number.isInteger(c) && c >= 0 && fila[c] !== undefined && fila[c] !== null) r.push(c); }
    return r;
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
    let dias = Math.floor(serie + 1e-9), seg = Math.round((serie - dias) * 86400);
    if (seg >= 86400) { dias++; seg -= 86400; }   // 23:59:59.6 se redondea al día siguiente, no a las 00:00:00 del mismo
    let base = fecha1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
    let d = dias;
    if (!fecha1904 && dias < 61) d = dias + 1;   // antes del 1900-03-01 Excel cuenta un día de menos (y el 29 de febrero de 1900 no existió)
    const t = new Date(base + d * 86400000);
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

    const usados = new Set(), sufijo = new Map();   // sufijo: el próximo número que se prueba para cada nombre repetido (miles de hojas iguales no vuelven a probar desde 2)
    const nombreUnico = (n, k) => {
      let t = String(n == null ? '' : n).replace(/[\[\]:*?\/\\]/g, '_').replace(/^'+|'+$/g, '').slice(0, 31) || 'Hoja' + (k + 1), b = t, i = sufijo.get(b.toLowerCase()) || 2;
      while (usados.has(t.toLowerCase())) { const suf = '(' + i++ + ')'; t = b.slice(0, 31 - suf.length) + suf; }
      sufijo.set(b.toLowerCase(), i);
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
    let cuerpo = '', filasAreas = 0;
    hojas.forEach((h, ih) => {
      const filas = Array.isArray(h.filas) ? h.filas : [];
      const ocultasC = new Set(h.columnasOcultas || []), ocultasF = new Set(h.filasOcultas || []), altos = h.altosFila || {}, anchos = Array.isArray(h.anchos) ? h.anchos : [];
      // combinadas y matrices de varias celdas: áreas que cubren todas sus celdas menos la de origen. Se recorren por fila (las que siguen abiertas), sin armar la
      // lista de celdas cubiertas: una combinada A1:XFD1048576 son 17.000 millones. El tope es la suma de las filas que abarcan, en todo el libro.
      const recortar = (r) => (r && r.f0 < MAX_FILAS && r.c0 < MAX_COLS ? { ...r, f1: Math.min(r.f1, MAX_FILAS - 1), c1: Math.min(r.c1, MAX_COLS - 1) } : null);
      const areas = [], origen = new Map(), matrices = new Map();
      const agregar = (r) => { const k = r.f1 - r.f0 + 1; if (filasAreas + k > MAX_FILAS_AREAS) return false; filasAreas += k; areas.push(r); return true; };
      const combinadas = [];
      for (const t of h.combinadas || []) {
        const r = recortar(parseRango(t));
        if (!r || (r.f1 === r.f0 && r.c1 === r.c0)) continue;
        if (agregar(r)) { combinadas.push(r); origen.set(r.f0 * 16384 + r.c0, r); } else avisar('combinadas', t);
      }
      // matrices: celda de origen → rango; el resto del rango son celdas cubiertas
      if (conFormulas) for (let f = 0; f < filas.length; f++) {
        const fila = filas[f];
        if (!fila) continue;
        for (const c of presentes(fila)) {
          const x = fila[c];
          if (!x || typeof x !== 'object' || !x.formula || !x.matriz) continue;
          const m = recortar(parseRango(x.matriz));
          if (!m || m.f0 !== f || m.c0 !== c) continue;
          if (m.f1 === m.f0 && m.c1 === m.c0) matrices.set(f * 16384 + c, m);
          else if (agregar(m)) matrices.set(f * 16384 + c, m);
          else matrices.set(f * 16384 + c, null);   // no cabe: queda el valor (celdaXml)
        }
      }
      areas.sort((a, b) => a.f0 - b.f0);
      let maxC = 0; for (const fila of filas) if (fila && fila.length > maxC) maxC = fila.length;
      for (const r of combinadas) if (r.c1 + 1 > maxC) maxC = r.c1 + 1;
      let xml = `<table:table table:name="${esc(h.nombre)}" table:style-name="${h.estado === 'oculta' || h.estado === 'muyOculta' ? 'taH' : 'ta1'}">`;
      // columnas (en tramos iguales)
      const colDe = (c) => estiloColumna(anchos[c]) + (ocultasC.has(c) ? '|h' : '');
      for (let c = 0; c < maxC;) { const k = colDe(c); let d = c; while (d + 1 < maxC && colDe(d + 1) === k) d++; const [est, h2] = k.split('|'); xml += `<table:table-column table:style-name="${est}"${d > c ? ` table:number-columns-repeated="${d - c + 1}"` : ''}${h2 ? ' table:visibility="collapse"' : ''} table:default-cell-style-name="Default"/>`; c = d + 1; }
      if (!maxC) xml += '<table:table-column table:style-name="co0" table:default-cell-style-name="Default"/>';
      // las filas que se escriben aunque no tengan nada: con alto propio u ocultas
      const marcadas = [];
      for (const k of Object.keys(altos)) { const f = +k; if (Number.isInteger(f) && f >= 0 && f < MAX_FILAS) marcadas.push(f); }
      for (const f of ocultasF) if (Number.isInteger(f) && f >= 0 && f < MAX_FILAS && altos[f] === undefined) marcadas.push(f);
      marcadas.sort((a, b) => a - b);
      let total = Math.min(filas.length, MAX_FILAS);
      if (marcadas.length && marcadas[marcadas.length - 1] + 1 > total) total = marcadas[marcadas.length - 1] + 1;
      for (const r of combinadas) if (r.f1 + 1 > total) total = r.f1 + 1;
      const conValores = (fila) => !!fila && presentes(fila).length > 0;
      const marcada = (f) => altos[f] !== undefined || ocultasF.has(f);
      // las filas iguales sin valores (cubiertas por una combinada alta, por ejemplo) se juntan en una repetida
      let previa = null, repPrevia = 0;
      const soltar = () => { if (previa) xml += previa.abre + (repPrevia > 1 ? ` table:number-rows-repeated="${repPrevia}"` : '') + '>' + previa.celdas + '</table:table-row>'; previa = null; repPrevia = 0; };
      let activas = [], ia = 0, im = 0;
      for (let f = 0; f < total;) {
        while (ia < areas.length && areas[ia].f0 <= f) activas.push(areas[ia++]);
        if (activas.length) activas = activas.filter((a) => a.f1 >= f);
        const fila = filas[f] || [];
        if (!activas.length && !marcada(f) && !conValores(fila)) {
          // un tramo de filas vacías, hasta la próxima con valores, marcada o con un área
          const tope = ia < areas.length ? Math.min(total, areas[ia].f0) : total;
          let g = f + 1;
          while (g < tope && g < filas.length && !marcada(g) && !conValores(filas[g])) g++;
          if (g >= filas.length && g < tope && !marcada(g)) { while (im < marcadas.length && marcadas[im] < g) im++; g = im < marcadas.length ? Math.min(tope, marcadas[im]) : tope; }
          soltar();
          xml += `<table:table-row table:style-name="ro0"${g - f > 1 ? ` table:number-rows-repeated="${g - f}"` : ''}><table:table-cell table:number-columns-repeated="${Math.max(1, maxC)}"/></table:table-row>`;
          f = g; continue;
        }
        // las columnas cubiertas de esta fila (tramos ordenados y juntos) y dónde empiezan las áreas que nacen en ella
        const tramos = [], inicios = [];
        for (const a of activas) { const d = a.f0 === f ? a.c0 + 1 : a.c0; if (a.f0 === f) inicios.push(a.c0); if (d <= a.c1) tramos.push([d, a.c1]); }
        tramos.sort((a, b) => a[0] - b[0]); inicios.sort((a, b) => a - b);
        const cub = [];
        for (const t of tramos) { const u = cub[cub.length - 1]; if (u && t[0] <= u[1] + 1) { if (t[1] > u[1]) u[1] = t[1]; } else cub.push(t); }
        // hasta dónde llega un tramo de columnas vacías que empieza en c (sin pasar de tope): hasta la próxima celda con algo, que se busca en la lista de la fila
        const pres = presentes(fila);
        let ip = 0;
        const finVacias = (c, tope) => { while (ip < pres.length && pres[ip] <= c) ip++; return ip < pres.length ? Math.min(tope, pres[ip] - 1) : tope; };
        let celdas = '', fija = inicios.length > 0, c = 0, ic = 0, io = 0;
        while (c < maxC) {
          while (ic < cub.length && cub[ic][1] < c) ic++;
          const x = fila[c];
          if (ic < cub.length && cub[ic][0] <= c) {
            if (x === undefined || x === null) { const d = finVacias(c, Math.min(cub[ic][1], maxC - 1)); celdas += `<table:covered-table-cell${d > c ? ` table:number-columns-repeated="${d - c + 1}"` : ''}/>`; c = d + 1; continue; }
            celdas += celdaXml(x, f, c, null, null, true, ih); fija = true; c++; continue;
          }
          const clave = f * 16384 + c;
          const celda = celdaXml(x, f, c, origen.get(clave), matrices.get(clave), false, ih);
          if (celda === null) {
            while (io < inicios.length && inicios[io] <= c) io++;
            const d = finVacias(c, Math.min(maxC, ic < cub.length ? cub[ic][0] : maxC, io < inicios.length ? inicios[io] : maxC) - 1);
            celdas += `<table:table-cell${d > c ? ` table:number-columns-repeated="${d - c + 1}"` : ''}/>`; c = d + 1;
          } else { celdas += celda; if (x !== undefined && x !== null) fija = true; c++; }
        }
        const abre = `<table:table-row table:style-name="${estiloFila(altos[f])}"${ocultasF.has(f) ? ' table:visibility="collapse"' : ''}`;
        if (!fija && previa && previa.abre === abre && previa.celdas === celdas) repPrevia++;
        else { soltar(); if (fija) xml += abre + '>' + celdas + '</table:table-row>'; else { previa = { abre, celdas }; repPrevia = 1; } }
        f++;
      }
      soltar();
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
      if (n.formula && conFormulas && matriz !== null) {   // null: una matriz que no cupo en el tope de áreas
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

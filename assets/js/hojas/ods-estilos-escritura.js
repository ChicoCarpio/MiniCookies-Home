/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Escritura de estilos y valores de celda de ODF para el libro editable de la suite (F3.4a, etapa 3).
// ORIGEN: copiado de `ods-escritor.js` del sitio (copia del 2026-10-28): `parrafos`, `isoDeSerie`, `traducirSeccion`, `xmlSeccion` y `estiloNumero` (código de formato de Excel →
// estilo de número de ODF), sin cambios; `esc` es el mismo. Lo nuevo (al final): `estiloCeldaXml`, que escribe un estilo de celda a partir de las propiedades mezcladas de
// `ods-estilos.js` (claves cortas), y `tipoDeFormato`. Sin DOM.
//
//   OdsEstilosEscritura.estiloNumero(codigo, nombre, avisar) → { tipo, xml, … } | null     (xml = <number:…-style style:name="nombre">…)
//   OdsEstilosEscritura.estiloCeldaXml(nombre, props, nombreNumero) → <style:style … family="table-cell">…</style:style>
//   OdsEstilosEscritura.isoDeSerie(serie, fecha1904) → { fecha, hora, seg };   parrafos(texto) → '<text:p>…</text:p>'
//   OdsEstilosEscritura.tipoDeFormato(codigo) → 'fecha' | 'hora' | 'porcentaje' | 'numero'
(function () {
  // (los caracteres de control que XML 1.0 no admite se quitan: con uno solo, LibreOffice no abre el archivo)
  const CONTROL = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g;
  const esc = (t) => String(t).replace(CONTROL, "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

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


  // ---- nuevo en la suite --------------------------------------------------------------------------------------------------
  // Qué clase de valor es un número con ese formato: decide si se escribe como date, time, percentage o float.
  function tipoDeFormato(codigo) {
    if (!codigo || codigo === 'General') return 'numero';
    const sec = String(codigo).split(";")[0].replace(/"[^"]*"/g, "").replace(/\[[^\]]*\]/g, "").replace(/\\./g, "").replace(/_./g, "");
    if (/[yd]/i.test(sec) || (/m{3,}/i.test(sec))) return 'fecha';
    if (/[hs]/i.test(sec) || /AM\/PM|A\/P/i.test(sec)) return "hora";
    if (/%/.test(sec)) return 'porcentaje';
    return 'numero';
  }
  // Las propiedades (claves cortas de ods-estilos.js) otra vez como atributos de ODF.
  const CELDA = { bg: 'fo:background-color', bi: 'fo:border-left', bd: 'fo:border-right', ba: 'fo:border-top', bb: 'fo:border-bottom', aj: 'fo:wrap-option', ro: 'style:rotation-angle', di: 'style:direction', va: 'style:vertical-align', rc: 'style:repeat-content', tas: 'style:text-align-source' };
  const PARRAFO = { ta: 'fo:text-align', ml: 'fo:margin-left', mr: 'fo:margin-right' };
  const TEXTO = { fn: 'style:font-name', ff: 'fo:font-family', fs: 'fo:font-size', fw: 'fo:font-weight', fi: 'fo:font-style', us: 'style:text-underline-style', ut: 'style:text-underline-type', lt: 'style:text-line-through-style', co: 'fo:color' };
  function estiloCeldaXml(nombre, p, nombreNumero) {
    const atributos = (mapa) => Object.keys(mapa).filter((k) => p[k] !== undefined).map((k) => ' ' + mapa[k] + '="' + esc(p[k]) + '"').join('');
    const c = atributos(CELDA), pa = atributos(PARRAFO), t = atributos(TEXTO);
    return '<style:style style:name="' + nombre + '" style:family="table-cell" style:parent-style-name="Default"' + (nombreNumero ? ' style:data-style-name="' + nombreNumero + '"' : '') + '>'
      + (c ? '<style:table-cell-properties' + c + '/>' : '') + (pa ? '<style:paragraph-properties' + pa + '/>' : '') + (t ? '<style:text-properties' + t + '/>' : '') + '</style:style>';
  }

  globalThis.OdsEstilosEscritura = { estiloNumero, estiloCeldaXml, isoDeSerie, parrafos, tipoDeFormato, esc };
})();

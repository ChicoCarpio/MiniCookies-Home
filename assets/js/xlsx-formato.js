/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Formato de número de Excel (los "códigos de formato" de la celda: 0.00, #,##0, 0%, d/m/yyyy, [h]:mm,
// # ?/?, 0.00E+00, "texto"@, secciones, condiciones y colores) del motor propio. Valor + código →
// el texto que muestra Excel. Se probó contra 1.996 casos medidos en Excel real
// (investigacion-motor-excel\arnes-xlsx\medidas\formatos.json). No depende del navegador.
//
// Uso:
//   const r = XlsxFormato.formatear(valor, codigo, opciones);   // → { texto, color, relleno, desborde }
//   valor: número | texto | booleano.   codigo: el de la celda ('General', '#,##0.00', …).
//   opciones: { idioma: 'es'|'en'      nombres de días y meses, AM/PM, VERDADERO (por defecto 'en')
//               decimal: '.'  miles: ','  separadores del idioma de quien mira
//               fecha1904: false          base de los números de serie de fecha
//               ancho: 0                  ancho de la columna en caracteres (para "*" y para decidir cuándo un número no cabe)
//               anchoGeneral: 11 }        cifras de "General" (Excel las reduce si la columna es angosta)
//   r.texto    el texto (con un solo carácter de relleno si hay "*x"; con ancho se expande)
//   r.color    'Red' | 'Blue' | … | 'Color3' | null (el color de fuente que pide el código)
//   r.desborde true si Excel mostraría "####" (fecha negativa o fuera de rango)
//   XlsxFormato.esFecha(codigo)  → true si el código formatea fechas u horas
//
// Lo que no cubre todavía (se trata como General o se deja como literal): los calendarios no gregorianos
// (salvo la era japonesa), [DBNum], y la máscara numérica combinada con fracción ("#0#00??/??").

(function () {
  const CLAVE_MAXIMA = 500;
  const cache = new Map();

  const NOMBRES = {
    es: {
      meses: ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'],
      mesesCortos: ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sept', 'oct', 'nov', 'dic'],
      dias: ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'],
      diasCortos: ['dom', 'lun', 'mar', 'mié', 'jue', 'vie', 'sáb'],
      // Excel en español separa "a. m." y "p. m." con un espacio fino no separable (U+202F)
      am: 'a.\u202Fm.', pm: 'p.\u202Fm.', verdadero: 'VERDADERO', falso: 'FALSO',
      fechaCorta: 'd/m/yyyy', fechaLarga: 'dddd, d "de" mmmm "de" yyyy', horaLarga: 'h:mm:ss',
      errores: { '#DIV/0!': '#¡DIV/0!', '#N/A': '#N/D', '#NAME?': '#¿NOMBRE?', '#NULL!': '#¡NULO!', '#NUM!': '#¡NUM!', '#REF!': '#¡REF!', '#VALUE!': '#¡VALOR!' },
    },
    en: {
      meses: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
      mesesCortos: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
      dias: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
      diasCortos: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
      am: 'AM', pm: 'PM', verdadero: 'TRUE', falso: 'FALSE',
      fechaCorta: 'm/d/yyyy', fechaLarga: 'dddd, mmmm d, yyyy', horaLarga: 'h:mm:ss AM/PM',
      errores: {},
    },
  };
  const COLORES = /^(black|blue|cyan|green|magenta|red|white|yellow|color\s*\d+)$/i;

  // 'red' / 'RED' → 'Red'; 'color 3' → 'Color3' (Color1 a Color56); otro → null
  function colorCanonico(nombre) {
    const limpio = nombre.replace(/\s+/g, '').toLowerCase();
    const m = /^color(\d+)$/.exec(limpio);
    if (m) return +m[1] >= 1 && +m[1] <= 56 ? 'Color' + +m[1] : null;
    return limpio.charAt(0).toUpperCase() + limpio.slice(1);
  }

  // ---- Análisis del código -----------------------------------------------------------------------------

  // Parte el código en secciones por ";" que no estén entre comillas, corchetes ni tras una barra.
  function partirSecciones(codigo) {
    const secciones = [];
    let actual = '', comillas = false, corchetes = false;
    for (let i = 0; i < codigo.length; i++) {
      const c = codigo[i];
      if (c === '\\' && !comillas) { actual += c + (codigo[i + 1] || ''); i++; continue; }
      if (c === '"') comillas = !comillas;
      else if (!comillas && c === '[') corchetes = true;
      else if (!comillas && c === ']') corchetes = false;
      if (c === ';' && !comillas && !corchetes) { secciones.push(actual); actual = ''; } else actual += c;
    }
    secciones.push(actual);
    return secciones;
  }

  function tokenizar(texto) {
    texto = texto.replace(/[\u0001\u0002]/g, ''); // las marcas internas (relleno) no pueden venir del archivo
    const tokens = [];
    let cond = null, color = null, lcid = null;
    const n = texto.length;
    let i = 0;
    while (i < n) {
      const c = texto[i];
      const resto = texto.slice(i);
      if (c === '"') { const f = texto.indexOf('"', i + 1); const fin = f < 0 ? n : f; tokens.push({ t: 'lit', v: texto.slice(i + 1, fin) }); i = fin + 1; continue; }
      if (c === '\\') { tokens.push({ t: 'lit', v: texto[i + 1] || '' }); i += 2; continue; }
      if (c === '_') { tokens.push({ t: 'esp', v: texto[i + 1] || '' }); i += 2; continue; }
      if (c === '*') { tokens.push({ t: 'fill', v: texto[i + 1] || ' ' }); i += 2; continue; }
      if (c === '[') {
        const f = texto.indexOf(']', i);
        if (f < 0) { i = n; continue; } // corchete sin cerrar: no es un código
        const dentro = texto.slice(i + 1, f);
        i = f + 1;
        let m;
        if (COLORES.test(dentro)) { color = colorCanonico(dentro); if (color) continue; }
        if ((m = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+(?:E[+-]?\d+)?)$/i.exec(dentro))) { cond = { op: m[1], n: parseFloat(m[2]) }; continue; }
        if ((m = /^\$([^-\]]*)(?:-([0-9A-Fa-f]+))?$/.exec(dentro))) { // [$€-2], [$-409], [$$-409]
          if (m[1]) tokens.push({ t: 'lit', v: m[1] });
          if (m[2]) lcid = parseInt(m[2], 16);
          continue;
        }
        if ((m = /^(h+|m+|s+)$/i.exec(dentro))) { tokens.push({ t: 'tr', u: m[1][0].toLowerCase(), n: m[1].length }); continue; }
        continue; // [DBNum1], [Hijri]… se ignoran
      }
      if (/^general/i.test(resto)) { tokens.push({ t: 'gen' }); i += 7; continue; }
      if (/^AM\/PM/i.test(resto)) { tokens.push({ t: 'ampm', forma: 'AM/PM' }); i += 5; continue; }
      if (/^A\/P/i.test(resto)) { tokens.push({ t: 'ampm', forma: 'A/P' }); i += 3; continue; }
      if (/^E[+-]/i.test(resto)) { tokens.push({ t: 'exp', signo: resto[1] }); i += 2; continue; }
      if (c === '0' || c === '#' || c === '?') { tokens.push({ t: 'dig', c }); i++; continue; }
      if (c === '.') { tokens.push({ t: 'pto' }); i++; continue; }
      if (c === ',') { tokens.push({ t: 'coma' }); i++; continue; }
      if (c === '%') { tokens.push({ t: 'pct' }); i++; continue; }
      if (c === '/') { tokens.push({ t: 'barra' }); i++; continue; }
      if (c === '@') { tokens.push({ t: 'txt' }); i++; continue; }
      const letra = c.toLowerCase();
      if ('ymdhs'.includes(letra) && /[A-Za-z]/.test(c)) {
        let k = i;
        while (k < n && texto[k].toLowerCase() === letra) k++;
        tokens.push({ t: letra, n: k - i });
        i = k; continue;
      }
      if (letra === 'e' && /[A-Za-z]/.test(c)) { let k = i; while (k < n && texto[k].toLowerCase() === 'e') k++; tokens.push({ t: 'e', n: k - i }); i = k; continue; }
      if (letra === 'g' && /[A-Za-z]/.test(c)) { let k = i; while (k < n && texto[k].toLowerCase() === 'g') k++; tokens.push({ t: 'g', n: k - i }); i = k; continue; }
      tokens.push({ t: 'lit', v: c }); i++;
    }
    return { tokens, cond, color, lcid };
  }

  // "m" es minuto (no mes) si va justo después de una hora o justo antes de segundos.
  function resolverMinutos(tokens) {
    const idx = tokens.map((t, i) => (['y', 'm', 'd', 'h', 's', 'tr', 'e', 'g', 'ampm'].includes(t.t) ? i : -1)).filter((i) => i >= 0);
    idx.forEach((i, k) => {
      const t = tokens[i];
      if (t.t !== 'm') return;
      const antes = tokens[idx[k - 1]], despues = tokens[idx[k + 1]];
      if ((antes && (antes.t === 'h' || (antes.t === 'tr' && antes.u === 'h'))) || (despues && (despues.t === 's' || (despues.t === 'tr' && despues.u === 's')))) t.t = 'min';
    });
    // ".0" detrás de segundos (o de una hora, como lo hace Excel) son décimas de segundo
    const salida = [];
    let previos = 0;
    for (let i = 0; i < tokens.length; i++) {
      const t = tokens[i];
      if (['h', 'min', 's', 'tr'].includes(t.t)) previos++;
      if (t.t === 'pto' && previos) {
        let k = i + 1;
        while (tokens[k] && tokens[k].t === 'dig' && tokens[k].c === '0') k++;
        if (k > i + 1) { salida.push({ t: 'frs', n: k - i - 1 }); i = k - 1; continue; }
      }
      salida.push(t);
    }
    tokens.length = 0;
    for (const t of salida) tokens.push(t);
    return tokens;
  }

  function analizar(codigo) {
    if (codigo.length > 255) codigo = 'General'; // el máximo de Excel; evita análisis enormes con códigos hostiles
    if (cache.has(codigo)) return cache.get(codigo);
    const secciones = partirSecciones(codigo).map((texto) => {
      const a = tokenizar(texto);
      const esFecha = a.tokens.some((t) => ['y', 'm', 'd', 'h', 's', 'tr', 'ampm', 'e', 'g'].includes(t.t));
      if (esFecha) resolverMinutos(a.tokens);
      return { tokens: a.tokens, cond: a.cond, color: a.color, lcid: a.lcid, esFecha, vacia: texto === '', soloMeta: texto !== '' && a.tokens.length === 0 };
    });
    const r = { secciones, codigo };
    if (cache.size > CLAVE_MAXIMA) cache.clear();
    cache.set(codigo, r);
    return r;
  }

  function esFecha(codigo) {
    const a = analizar(codigo || 'General');
    return a.secciones.some((s) => s.esFecha);
  }

  // ---- Aritmética decimal (Excel redondea el valor de 15 cifras, no el binario) ---------------------------

  // |v| → { digitos: '123450000000000' (15 cifras), exp: 2 }   (v = 0.digitos × 10^(exp+1))
  function descomponer(v) {
    if (v === 0) return { digitos: '0'.repeat(15), exp: 0, cero: true };
    const [m, e] = Math.abs(v).toExponential(14).split('e');
    return { digitos: m.replace('.', ''), exp: parseInt(e, 10), cero: false };
  }
  // Redondea (mitad hacia afuera) |v| a dec decimales → { ent: '123', frac: '45' }
  function redondear(v, dec) {
    const { digitos, exp, cero } = descomponer(v);
    if (cero) return { ent: '0', frac: '0'.repeat(dec) };
    // dígitos con la coma después de (exp+1) cifras
    let ent, frac;
    if (exp >= 0) { const k = exp + 1; ent = digitos.slice(0, k).padEnd(k, '0'); frac = digitos.slice(k); } else { ent = '0'; frac = '0'.repeat(-exp - 1) + digitos; }
    let todo = ent + frac.padEnd(dec, '0').slice(0, dec);
    const siguiente = frac.length > dec ? frac.charCodeAt(dec) - 48 : 0;
    if (siguiente >= 5) todo = sumarUno(todo);
    ent = todo.slice(0, todo.length - dec) || '0';
    ent = ent.replace(/^0+(?=\d)/, '');
    return { ent, frac: todo.slice(todo.length - dec) };
  }
  function sumarUno(s) {
    const a = s.split('');
    let i = a.length - 1;
    while (i >= 0) { if (a[i] === '9') { a[i] = '0'; i--; } else { a[i] = String(+a[i] + 1); break; } }
    return (i < 0 ? '1' : '') + a.join('');
  }

  // ---- General ----------------------------------------------------------------------------------------------

  function general(v, opc) {
    if (v === 0) return '0';
    const max = opc.anchoGeneral || 11;
    const abs = Math.abs(v);
    const signo = v < 0 ? '-' : '';
    const { digitos, exp } = descomponer(v);
    const sigValor = Number(abs.toPrecision(6)).toExponential(5).split('e')[0].replace('.', '').replace(/0+$/, '').length || 1;
    const dosPuntos = (t) => t.replace('.', opc.decimal);
    // notación científica con hasta 6 cifras significativas
    const cientifica = () => {
      // hasta 6 cifras; en una columna angosta (max < 11) las que quepan: cifras + coma + E+dd
      const largoExp = Math.abs(parseInt(abs.toExponential(5).split('e')[1], 10)) >= 100 ? 5 : 4;
      const cifras = Math.max(1, Math.min(6, max - largoExp - 1));
      const m = abs.toExponential(cifras - 1).split('e');
      let mant = m[0].replace(/0+$/, '').replace(/\.$/, '');
      const e = parseInt(m[1], 10);
      return signo + dosPuntos(mant) + 'E' + (e < 0 ? '-' : '+') + String(Math.abs(e)).padStart(2, '0');
    };
    // decimal fijo en `max` caracteres (la coma cuenta)
    const longEnt = exp >= 0 ? exp + 1 : 1;
    if (longEnt > max) return cientifica();
    const dec = Math.max(0, max - longEnt - 1);
    const r = redondear(abs, dec);
    if (r.ent.length > max) return cientifica();
    const fracLimpia = r.frac.replace(/0+$/, '');
    const sigFijo = (r.ent === '0' ? '' : r.ent).replace(/^0+/, '').length + (r.ent === '0' ? fracLimpia.replace(/^0+/, '').length : fracLimpia.length);
    if (r.ent === '0' && fracLimpia === '') return cientifica();
    if (sigFijo < sigValor) return cientifica();
    return signo + r.ent + (fracLimpia ? dosPuntos('.' + fracLimpia) : '');
  }

  // ---- Números con marcas de posición ------------------------------------------------------------------

  // Reparte los dígitos de la parte entera entre las marcas 0 # ? (de derecha a izquierda), con los literales entre ellas.
  function emitirEntera(tokens, digitos, miles, opc) {
    const marcas = tokens.filter((t) => t.t === 'dig').length;
    let rango = 0; // posición desde la derecha de la próxima marca
    // se recorre de derecha a izquierda para saber el rango de cada marca
    const rangos = new Map();
    let k = 0;
    for (let i = tokens.length - 1; i >= 0; i--) if (tokens[i].t === 'dig') rangos.set(i, k++);
    const primera = tokens.findIndex((t) => t.t === 'dig');
    let salida = '';
    for (let i = 0; i < tokens.length; i++) {
      const t = tokens[i];
      if (t.t === 'dig') {
        const r = rangos.get(i);
        const emitirHasta = i === primera ? digitos.length - 1 : r; // la primera marca saca también los dígitos sobrantes
        for (let q = Math.max(emitirHasta, r); q >= r; q--) {
          const d = digitos.length - 1 - q >= 0 && q < digitos.length ? digitos[digitos.length - 1 - q] : null;
          let ch;
          if (d !== null) ch = d; else ch = t.c === '0' ? '0' : t.c === '?' ? ' ' : '';
          salida += ch;
          if (miles && q > 0 && q % 3 === 0 && ch !== '' && ch !== ' ') salida += opc.miles;
        }
        rango = r;
      } else salida += textoLiteral(t);
    }
    return salida;
  }
  function textoLiteral(t) {
    if (t.t === 'lit') return t.v;
    if (t.t === 'esp') return ' ';
    if (t.t === 'fill') return '\u0001' + t.v + '\u0001';
    if (t.t === 'pct') return '%';
    if (t.t === 'coma') return ',';
    if (t.t === 'pto') return '.';
    if (t.t === 'barra') return '/';
    return '';
  }

  function renderNumero(tokens, v, opc) {
    // % multiplica por 100; las comas finales dividen por mil
    let pct = tokens.filter((t) => t.t === 'pct').length;
    let ult = -1;
    tokens.forEach((t, i) => { if (t.t === 'dig') ult = i; });
    let escala = 0;
    for (let i = ult + 1; i < tokens.length && tokens[i].t === 'coma'; i++) escala++;
    let iPunto = tokens.findIndex((t) => t.t === 'pto');
    const iBarra = tokens.findIndex((t) => t.t === 'barra');
    let valor = Math.abs(v);
    if (pct) valor = Number((valor * Math.pow(100, pct)).toPrecision(15));
    if (escala) valor = Number((valor / Math.pow(1000, escala)).toPrecision(15));
    if (!Number.isFinite(valor) || valor >= 1e300) return { texto: '#', desborde: true };
    // las comas entre marcas (no las finales de escala) agrupan miles
    const primeraDig = tokens.findIndex((t) => t.t === 'dig');
    const miles = tokens.some((t, i) => t.t === 'coma' && i > primeraDig && i < ult + 1 && tokens.slice(0, i).some((x) => x.t === 'dig') && tokens.slice(i + 1, ult + 1).some((x) => x.t === 'dig'));
    const sinEscala = tokens.filter((t, i) => !(t.t === 'coma' && (i > ult || i < primeraDig)));
    const sinComas = sinEscala.filter((t) => t.t !== 'coma');
    iPunto = sinComas.findIndex((t) => t.t === 'pto');
    if (iBarra >= 0) return renderFraccion(sinComas, valor, v < 0, opc);
    const iE = sinComas.findIndex((t) => t.t === 'exp');
    if (iE >= 0) return renderCientifica(sinComas, valor, miles, opc);
    const nDec = iPunto >= 0 ? sinComas.slice(iPunto).filter((t) => t.t === 'dig').length : 0;
    const r = redondear(valor, nDec);
    let digitosEnt = r.ent === '0' ? '' : r.ent;
    const parteEntera = iPunto >= 0 ? sinComas.slice(0, iPunto) : sinComas;
    const parteDec = iPunto >= 0 ? sinComas.slice(iPunto + 1) : [];
    const hayMarcasEnt = parteEntera.some((t) => t.t === 'dig');
    let texto;
    if (hayMarcasEnt) texto = emitirEntera(parteEntera, digitosEnt, miles, opc);
    else texto = parteEntera.map(textoLiteral).join('') + (iPunto >= 0 ? digitosEnt : digitosEnt === '' && iPunto < 0 ? '' : '');
    if (iPunto >= 0) {
      texto += opc.decimal;
      let j = 0;
      for (const t of parteDec) {
        if (t.t === 'dig') {
          const d = r.frac[j++];
          const resto = r.frac.slice(j - 1);
          if (t.c === '0') texto += d; else if (/^0*$/.test(resto)) texto += t.c === '?' ? ' ' : ''; else texto += d;
        } else if (t.t === 'pct') texto += '%';
        else texto += textoLiteral(t);
      }
    }
    const algoMostrado = /[1-9]/.test(r.ent + r.frac);
    return { texto, negativoVisible: algoMostrado };
  }

  // ---- Científica ---------------------------------------------------------------------------------------

  function renderCientifica(tokens, valor, miles, opc) {
    const iE = tokens.findIndex((t) => t.t === 'exp');
    const iPunto = tokens.findIndex((t, i) => t.t === 'pto' && i < iE);
    const antes = tokens.slice(0, iE);
    const parteEnt = iPunto >= 0 ? antes.slice(0, iPunto) : antes;
    const parteDec = iPunto >= 0 ? antes.slice(iPunto + 1) : [];
    const despues = tokens.slice(iE + 1);
    const nEnt = Math.max(1, parteEnt.filter((t) => t.t === 'dig').length);
    const nDec = parteDec.filter((t) => t.t === 'dig').length;
    const nExp = despues.filter((t) => t.t === 'dig').length;
    let e = 0;
    const d = descomponer(valor);
    if (!d.cero) e = Math.floor(d.exp / nEnt) * nEnt;
    let mant;
    for (let intento = 0; intento < 3; intento++) {
      // mantisa = valor / 10^e, con el decimal corrido sobre las 15 cifras
      const desplazado = d.cero ? 0 : Number(d.digitos[0] + '.' + d.digitos.slice(1) + 'e' + (d.exp - e));
      mant = redondear(desplazado, nDec);
      if (mant.ent.length > nEnt && !d.cero) { e += nEnt; continue; }
      break;
    }
    let texto = emitirEntera(parteEnt, d.cero ? '0'.repeat(nEnt) : mant.ent === '0' && nEnt > 0 && parteEnt.every((t) => t.t !== 'dig' || t.c === '#') ? '' : mant.ent, miles, opc);
    if (iPunto >= 0) {
      texto += opc.decimal;
      let j = 0;
      for (const t of parteDec) { if (t.t === 'dig') { const dd = mant.frac[j++]; texto += t.c === '0' || !/^0*$/.test(mant.frac.slice(j - 1)) ? dd : (t.c === '?' ? ' ' : ''); } else texto += textoLiteral(t); }
    }
    const signoExp = e < 0 ? '-' : tokens[iE].signo === '+' ? '+' : '';
    texto += 'E' + signoExp + String(Math.abs(e)).padStart(nExp, '0');
    for (const t of despues) if (t.t !== 'dig') texto += textoLiteral(t);
    return { texto, negativoVisible: !d.cero };
  }

  // ---- Fracciones ---------------------------------------------------------------------------------------

  // Excel usa la última convergente de la fracción continua cuyo denominador cabe (12,3 con ?/? → 1/3, no 2/7).
  function mejorFraccion(f, maxDen) {
    let h0 = 0, k0 = 1, h1 = 1, k1 = 0, x = f;
    let mejor = { n: 0, d: 1 };
    for (let i = 0; i < 40; i++) {
      const a = Math.floor(x + 1e-12);
      const h2 = a * h1 + h0, k2 = a * k1 + k0;
      if (k2 > maxDen) break;
      mejor = { n: h2, d: k2 };
      h0 = h1; k0 = k1; h1 = h2; k1 = k2;
      const resto = x - a;
      if (resto < 1e-10 || Math.abs(f - h2 / k2) < 1e-12) break;
      x = 1 / resto;
    }
    return mejor;
  }

  function renderFraccion(tokens, valor, negativo, opc) {
    const iBarra = tokens.findIndex((t) => t.t === 'barra');
    const antes = tokens.slice(0, iBarra), despues = tokens.slice(iBarra + 1);
    // numerador: las marcas inmediatamente antes de "/"; lo anterior (si hay un espacio de por medio) es la parte entera
    let iNum = antes.length;
    while (iNum > 0 && antes[iNum - 1].t === 'dig') iNum--;
    const marcasNum = antes.slice(iNum);
    const prefijo = antes.slice(0, iNum);
    const conEntera = prefijo.some((t) => t.t === 'dig');
    // denominador: marcas, o dígitos literales (denominador fijo)
    let iDen = 0;
    while (iDen < despues.length && (despues[iDen].t === 'dig' || (despues[iDen].t === 'lit' && /^\d$/.test(despues[iDen].v)))) iDen++;
    const tokDen = despues.slice(0, iDen), cola = despues.slice(iDen);
    // Denominador fijo: hay un dígito 1-9 literal (?/10, ??/100, ?/50: el 0 también viene como marca, pero aquí es una cifra)
    const fijo = tokDen.some((t) => t.t === 'lit') && tokDen.every((t) => t.t === 'lit' || (t.t === 'dig' && t.c === '0')) ? parseInt(tokDen.map((t) => (t.t === 'lit' ? t.v : '0')).join(''), 10) : null;
    const nMarcasDen = tokDen.filter((t) => t.t === 'dig').length;
    const maxDen = fijo || Math.pow(10, nMarcasDen) - 1;
    const aTexto = (x) => (x >= 1e15 ? BigInt(Math.floor(x)).toString() : String(x)); // String(1e21) daría "1e+21"
    let ent = conEntera ? Math.floor(valor + 1e-12) : 0;
    let f = valor - ent;
    if (f < 0) f = 0;
    let num, den;
    if (fijo) { den = fijo; num = Math.round(f * fijo); } else { const m = mejorFraccion(f, Math.max(1, maxDen)); num = m.n; den = m.d; }
    if (num >= den && conEntera && den > 0) { ent += Math.floor(num / den); num -= Math.floor(num / den) * den; }
    const anchoNum = marcasNum.length, anchoDen = nMarcasDen || String(den).length;
    const pad = (s, ancho, marcas, izq) => {
      const m = marcas.map((t) => t.c);
      let r = s;
      while (r.length < ancho) {
        const pos = izq ? ancho - r.length - 1 : r.length;
        const ch = m[pos] === '0' ? '0' : ' ';
        r = izq ? ch + r : r + ch;
      }
      return r;
    };
    let texto = '';
    if (conEntera) texto += emitirEntera(prefijo, ent === 0 && num !== 0 && prefijo.every((t) => t.t !== 'dig' || t.c !== '0') ? '' : aTexto(ent), false, opc); // 0 con fracción 0 se ve como "0"
    else texto += prefijo.map(textoLiteral).join('');
    const sinFraccion = conEntera && num === 0;
    if (sinFraccion) {
      // Excel deja el hueco de la fracción en blanco
      texto += ' '.repeat(marcasNum.length + 1 + (nMarcasDen || String(den).length));
    } else {
      const numerador = conEntera ? num : ent * den + num;
      const numTexto = conEntera ? pad(aTexto(numerador), anchoNum, marcasNum, true) : (aTexto(numerador).length > anchoNum ? aTexto(numerador) : pad(aTexto(numerador), anchoNum, marcasNum, true));
      texto += numTexto + '/' + pad(String(den), fijo ? String(fijo).length : anchoDen, tokDen.filter((t) => t.t === 'dig'), false);
    }
    for (const t of cola) texto += textoLiteral(t);
    // Sin parte entera, un valor que se redondea a 0/1 pierde el signo (Excel muestra "0/1"); con parte entera lo conserva ("-0    ")
    return { texto, negativoVisible: conEntera ? valor !== 0 : ent * den + num !== 0 };
  }

  // ---- Fechas y horas -----------------------------------------------------------------------------------

  const MS_DIA = 86400000;

  function componentesFecha(serial, fecha1904) {
    const dias = Math.floor(serial);
    let a, m, d;
    if (fecha1904) {
      const f = new Date(Date.UTC(1904, 0, 1) + dias * MS_DIA);
      a = f.getUTCFullYear(); m = f.getUTCMonth() + 1; d = f.getUTCDate();
    } else if (dias < 61) {
      a = 1900;
      if (dias <= 0) { m = 1; d = 0; } else if (dias <= 31) { m = 1; d = dias; } else if (dias <= 59) { m = 2; d = dias - 31; } else { m = 2; d = 29; }
    } else {
      const f = new Date(Date.UTC(1899, 11, 30) + dias * MS_DIA);
      a = f.getUTCFullYear(); m = f.getUTCMonth() + 1; d = f.getUTCDate();
    }
    const semana = fecha1904 ? new Date(Date.UTC(1904, 0, 1) + dias * MS_DIA).getUTCDay() : ((dias - 1) % 7 + 7) % 7;
    return { a, m, d, semana, dias };
  }

  function renderFecha(sec, valor, opc, nombres) {
    const tokens = sec.tokens;
    if (!(valor >= 0) || valor >= 2958466 && !opc.fecha1904) return { texto: '#', desborde: true };
    if (opc.fecha1904 && valor >= 2957004) return { texto: '#', desborde: true };
    // unidad de redondeo según lo más fino que se muestre
    const frs = tokens.find((t) => t.t === 'frs');
    const hayTiempo = tokens.some((t) => ['h', 'min', 's', 'tr'].includes(t.t));
    let unidad = MS_DIA;
    if (hayTiempo) {
      // Excel redondea al segundo (o a la décima, centésima o milésima que se muestre) y luego trunca: 13:37,63 con h:mm → 13:37.
      unidad = frs ? Math.max(1, Math.pow(10, 3 - frs.n)) : 1000;
    }
    let dias = Math.floor(valor);
    let ms = Math.round((valor - dias) * MS_DIA);
    if (hayTiempo) ms = Math.round(ms / unidad) * unidad;
    if (ms >= MS_DIA) { ms -= MS_DIA; dias += 1; }
    if (opc.fecha1904 ? dias >= 2957004 : dias >= 2958466) return { texto: '#', desborde: true };
    const totalMs = dias * MS_DIA + ms;
    const c = componentesFecha(dias, opc.fecha1904);
    const h = Math.floor(ms / 3600000), mi = Math.floor(ms / 60000) % 60, s = Math.floor(ms / 1000) % 60, mil = ms % 1000;
    const dos = (n) => String(n).padStart(2, '0');
    let salida = '';
    const enIngles = nombres === NOMBRES.en;
    const marcaAmpm = (t) => t.t === 'ampm' && (t.forma === 'AM/PM' || enIngles);
    const ampm = tokens.some(marcaAmpm);
    for (let i = 0; i < tokens.length; i++) {
      const t = tokens[i];
      switch (t.t) {
        case 'y': salida += t.n <= 2 ? dos(c.a % 100) : String(c.a).padStart(4, '0'); break;
        case 'e': salida += tokens.some((x) => x.t === 'g') ? String(eraJaponesa(c, t).anio) : String(c.a); break;
        case 'g': salida += eraJaponesa(c, t).era; break;
        case 'm': salida += t.n === 1 ? c.m : t.n === 2 ? dos(c.m) : t.n === 3 ? nombres.mesesCortos[c.m - 1] : t.n === 4 ? nombres.meses[c.m - 1] : nombres.meses[c.m - 1][0]; break;
        case 'd': salida += t.n === 1 ? c.d : t.n === 2 ? dos(c.d) : t.n === 3 ? nombres.diasCortos[c.semana] : nombres.dias[c.semana]; break;
        case 'h': { const hh = ampm ? (h % 12 || 12) : h; salida += t.n === 1 ? hh : dos(hh); break; }
        case 'min': salida += t.n === 1 ? mi : dos(mi); break;
        case 's': salida += t.n === 1 ? s : dos(s); break;
        case 'frs': salida += opc.decimal + String(mil).padStart(3, '0').slice(0, t.n); break;
        case 'ampm': salida += !marcaAmpm(t) ? 'A/P' : t.forma === 'AM/PM' ? (h < 12 ? nombres.am : nombres.pm) : (h < 12 ? 'A' : 'P'); break;
        case 'tr': {
          const total = Math.floor(totalMs / ({ h: 3600000, m: 60000, s: 1000 })[t.u]);
          const n = t.u === 'h' ? total : t.u === 'm' ? total : total;
          salida += String(n).padStart(t.n, '0'); break;
        }
        default: salida += textoLiteral(t);
      }
    }
    return { texto: salida };
  }

  const formateadoresEra = [null, null];
  function eraJaponesa(c, t) {
    try {
      const f = new Date(Date.UTC(c.a, c.m - 1, Math.max(1, c.d)));
      const largo = t.t === 'g' && t.n >= 3;
      if (!formateadoresEra[largo ? 1 : 0]) formateadoresEra[largo ? 1 : 0] = new Intl.DateTimeFormat('ja-JP-u-ca-japanese', { era: largo ? 'long' : 'short', year: 'numeric', timeZone: 'UTC' });
      const partes = formateadoresEra[largo ? 1 : 0].formatToParts(f);
      const era = partes.find((p) => p.type === 'era'), anio = partes.find((p) => p.type === 'year');
      return { era: era ? era.value : '', anio: anio ? anio.value : c.a };
    } catch (e) { return { era: '', anio: c.a }; }
  }

  // ---- Texto y errores ----------------------------------------------------------------------------------

  function renderTexto(sec, texto, opc) {
    if (!sec.tokens.some((t) => t.t === 'txt')) return sec.vacia || sec.tokens.length === 0 ? '' : sec.tokens.map(textoLiteral).join('') + (sec.tokens.some((t) => t.t === 'gen') ? texto : '');
    return sec.tokens.map((t) => (t.t === 'txt' ? texto : textoLiteral(t))).join('');
  }

  // ---- Elegir sección y armar el resultado -------------------------------------------------------------

  function cumple(cond, v) {
    switch (cond.op) {
      case '<': return v < cond.n; case '>': return v > cond.n; case '=': return v === cond.n;
      case '<=': return v <= cond.n; case '>=': return v >= cond.n; default: return v !== cond.n;
    }
  }

  function elegir(secciones, v) {
    const n = secciones.length;
    const conCond = secciones.some((s) => s.cond);
    if (conCond) {
      for (let i = 0; i < n - 1 && i < 2; i++) {
        const cond = secciones[i].cond || (i === 0 ? { op: '>', n: 0 } : { op: '<', n: 0 });
        if (cumple(cond, v)) return { sec: secciones[i], absoluto: false };
      }
      return { sec: secciones[Math.min(n - 1, 2)], absoluto: false };
    }
    if (n === 1) return { sec: secciones[0], absoluto: false };
    if (n === 2) return v < 0 ? { sec: secciones[1], absoluto: true } : { sec: secciones[0], absoluto: false };
    if (v > 0) return { sec: secciones[0], absoluto: false };
    if (v < 0) return { sec: secciones[1], absoluto: true };
    return { sec: secciones[2], absoluto: false };
  }

  function rellenar(texto, ancho) {
    // el relleno "*x" está marcado con \u0001x\u0001; el primero se expande hasta el ancho de la columna (o queda en uno) y los demás se quitan
    const m = /\u0001([^\u0001])\u0001/.exec(texto);
    if (!m) return { texto: texto.replace(/[\u0001\u0002]/g, ''), relleno: null };
    const sinMarcas = texto.replace(/\u0001[^\u0001]\u0001/g, '');
    const faltan = Math.max(1, Math.min(255, ancho) - sinMarcas.length);
    const antes = texto.slice(0, m.index).replace(/\u0001[^\u0001]\u0001/g, ''), despues = texto.slice(m.index + m[0].length).replace(/\u0001[^\u0001]\u0001/g, '');
    return { texto: antes + m[1].repeat(faltan) + despues, relleno: { caracter: m[1], indice: antes.length } };
  }

  function formatear(valor, codigo, opciones) {
    const opc = Object.assign({ idioma: 'en', decimal: '.', miles: ',', fecha1904: false, ancho: 0, anchoGeneral: 11 }, opciones || {});
    opc.ancho = Math.max(0, Math.min(255, Number(opc.ancho) || 0)); // el máximo de una columna de Excel
    codigo = codigo === undefined || codigo === null || codigo === '' ? 'General' : String(codigo);
    const nombres = NOMBRES[opc.idioma] || NOMBRES.en;
    if (valor === null || valor === undefined) return { texto: '', color: null };
    if (typeof valor === 'boolean') return { texto: valor ? nombres.verdadero : nombres.falso, color: null };
    // numFmtId 14 (mm-dd-yy en la norma): Excel lo muestra con la fecha corta de quien mira
    if (opc.idFormato === 14) codigo = nombres.fechaCorta;
    const analisis = analizar(codigo);
    const secciones = analisis.secciones;

    if (typeof valor === 'string') {
      valor = valor.replace(/[\u0001\u0002]/g, '');
      const sec = secciones.length >= 4 ? secciones[3] : (secciones.length === 1 || secciones.every((s) => !s.tokens.some((t) => t.t === 'txt'))) && secciones.length < 4 ? (secciones.length === 1 && secciones[0].tokens.some((t) => t.t === 'txt') ? secciones[0] : null) : secciones.find((s) => s.tokens.some((t) => t.t === 'txt')) || null;
      if (!sec) return { texto: valor, color: null };
      const r = rellenar(renderTexto(sec, valor, opc), opc.ancho);
      return { texto: r.texto, color: sec.color, relleno: r.relleno };
    }
    if (!Number.isFinite(valor)) return { texto: nombres.errores['#NUM!'] || '#NUM!', color: null };

    const { sec, absoluto } = elegir(secciones, valor);
    const v = valor;
    // lcid de la sección: cambia el idioma de los nombres
    let nom = nombres;
    if (sec.lcid !== null && sec.lcid !== undefined) {
      const bajo = sec.lcid & 0xFF;
      if (bajo === 0x0A) nom = NOMBRES.es; else if (bajo === 0x09) nom = NOMBRES.en;
      else if (sec.lcid === 0xF800 || sec.lcid === 0xF400) {
        const sustituto = analizar(sec.lcid === 0xF800 ? nombres.fechaLarga : nombres.horaLarga).secciones[0];
        const r = renderFecha(sustituto, v, opc, nombres);
        if (r.desborde) return { texto: '#'.repeat(Math.max(1, opc.ancho || 8)), color: sec.color, desborde: true };
        return { texto: r.texto, color: sec.color };
      }
    }
    if (sec.soloMeta) return { texto: general(absoluto ? Math.abs(v) : v, opc), color: sec.color };
    if (sec.vacia || (sec.tokens.length === 0)) return { texto: '', color: sec.color };
    if (sec.tokens.length === 1 && sec.tokens[0].t === 'gen') return { texto: general(absoluto ? Math.abs(v) : v, opc), color: sec.color };
    if (sec.esFecha) {
      const r = renderFecha(sec, v, opc, nom);
      if (r.desborde) return { texto: '#'.repeat(Math.max(1, opc.ancho || 8)), color: sec.color, desborde: true };
      const f = rellenar(r.texto, opc.ancho);
      return { texto: f.texto, color: sec.color, relleno: f.relleno };
    }
    if (sec.tokens.some((t) => t.t === 'gen')) {
      const antes = []; let despues = [];
      let visto = false;
      for (const t of sec.tokens) { if (t.t === 'gen') visto = true; else (visto ? despues : antes).push(t); }
      const g = general(absoluto ? Math.abs(v) : v, opc);
      const rg = rellenar(antes.map(textoLiteral).join('') + g + despues.map(textoLiteral).join(''), opc.ancho);
      return { texto: rg.texto, color: sec.color, relleno: rg.relleno };
    }
    if (sec.tokens.some((t) => t.t === 'txt') && !sec.tokens.some((t) => t.t === 'dig')) return { texto: general(absoluto ? Math.abs(v) : v, opc), color: sec.color };
    const r = renderNumero(sec.tokens, v, opc);
    if (r.desborde) return { texto: '#'.repeat(Math.max(1, opc.ancho || 8)), color: sec.color, desborde: true };
    const conSigno = v < 0 && !absoluto && r.negativoVisible;
    // el signo va al principio, antes de cualquier literal
    const f = rellenar((conSigno ? '-' : '') + r.texto, opc.ancho);
    return { texto: f.texto, color: sec.color, relleno: f.relleno };
  }


  globalThis.XlsxFormato = { formatear, esFecha, general, NOMBRES };
})();

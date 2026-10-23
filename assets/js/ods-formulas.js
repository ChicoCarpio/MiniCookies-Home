/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Fórmulas de OpenDocument (OpenFormula, la sintaxis de `table:formula="of:=SUM([.A1:.B2])"`) ↔ fórmulas como las escribe un .xlsx (`SUM(A1:B2)`). No depende del navegador.
//
//   OdsFormulas.aOpenFormula("SUM(Datos!A1:A3)")        → "SUM([$Datos.A1:.A3])"            (sin el "of:=")
//   OdsFormulas.aExcel("SUM([$Datos.A1:.A3])")          → "SUM(Datos!A1:A3)"                (sin el "of:="; null si no se puede llevar)
//
// Las diferencias que cubre (medidas con LibreOffice, `Investigaciones\investigacion-formatos-antiguos\arnes\pares-formulas-ods.mjs`): las referencias van entre corchetes con punto
// (`[$Hoja.A1:.B2]`, la hoja con `$` solo en la primera mitad, `[.A1]` en la misma hoja); los argumentos se separan con `;`; la unión es `~` y la intersección `!` (en Excel `,`
// entre paréntesis y un espacio); las constantes lógicas llevan paréntesis (`TRUE()`); las matrices en línea separan columnas con `;` y filas con `|`; las funciones que Excel guarda
// como `_xlfn.NOMBRE` llevan en LibreOffice el prefijo `COM.MICROSOFT.` (salvo las que ya están en OpenFormula 1.2).
(function () {
  // <tabla>   (generado por Investigaciones\investigacion-formatos-antiguos\arnes\generar-tabla-ods.mjs)
  // Funciones que LibreOffice escribe con el prefijo COM.MICROSOFT.
  const T_COM = 'FLOOR,CEILING,AGGREGATE,BETA.DIST,BETA.INV,BINOM.DIST,BINOM.INV,CEILING.MATH,CEILING.PRECISE,CHISQ.DIST,CHISQ.DIST.RT,CHISQ.INV,CHISQ.INV.RT,CHISQ.TEST,CHOOSECOLS,CHOOSEROWS,CONCAT,CONFIDENCE.NORM,CONFIDENCE.T,COVARIANCE.P,COVARIANCE.S,DROP,ENCODEURL,ERF.PRECISE,ERFC.PRECISE,EXPAND,EXPON.DIST,F.DIST.RT,F.INV.RT,F.TEST,FILTER,FILTERXML,FLOOR.MATH,FLOOR.PRECISE,FORECAST.ETS,FORECAST.ETS.CONFINT,FORECAST.ETS.SEASONALITY,FORECAST.ETS.STAT,FORECAST.LINEAR,GAMMA.DIST,GAMMA.INV,GAMMALN.PRECISE,HSTACK,HYPGEOM.DIST,IFS,ISO.CEILING,LOGNORM.DIST,LOGNORM.INV,MAXIFS,MINIFS,MODE.MULT,MODE.SNGL,NEGBINOM.DIST,NETWORKDAYS.INTL,NORM.DIST,NORM.INV,NORM.S.DIST,NORM.S.INV,PERCENTILE.EXC,PERCENTILE.INC,PERCENTRANK.EXC,PERCENTRANK.INC,POISSON.DIST,QUARTILE.EXC,QUARTILE.INC,RANDARRAY,RANK.AVG,RANK.EQ,SEQUENCE,SORT,SORTBY,STDEV.P,STDEV.S,SWITCH,T.DIST,T.DIST.2T,T.DIST.RT,T.INV,T.INV.2T,T.TEST,TAKE,TEXTAFTER,TEXTBEFORE,TEXTJOIN,TEXTSPLIT,TOCOL,TOROW,UNIQUE,VAR.P,VAR.S,VSTACK,WEBSERVICE,WEIBULL.DIST,WORKDAY.INTL,WRAPCOLS,WRAPROWS,XLOOKUP,XMATCH,Z.TEST';
  // Funciones que un .xlsx escribe con el prefijo _xlfn. (y _xlfn._xlws. las dos de abajo)
  const T_XLFN = 'ACOT,ACOTH,AGGREGATE,ARABIC,ARRAYTOTEXT,BASE,BETA.DIST,BETA.INV,BINOM.DIST,BINOM.DIST.RANGE,BINOM.INV,BITAND,BITLSHIFT,BITOR,BITRSHIFT,BITXOR,BYCOL,BYROW,CEILING.MATH,CEILING.PRECISE,CHISQ.DIST,CHISQ.DIST.RT,CHISQ.INV,CHISQ.INV.RT,CHISQ.TEST,CHOOSECOLS,CHOOSEROWS,COMBINA,CONCAT,CONFIDENCE.NORM,CONFIDENCE.T,COT,COTH,COVARIANCE.P,COVARIANCE.S,CSC,CSCH,DAYS,DECIMAL,DROP,ENCODEURL,ERF.PRECISE,ERFC.PRECISE,EXPAND,EXPON.DIST,F.DIST,F.DIST.RT,F.INV,F.INV.RT,F.TEST,FIELDVALUE,FILTER,FILTERXML,FLOOR.MATH,FLOOR.PRECISE,FORECAST.ETS,FORECAST.ETS.CONFINT,FORECAST.ETS.SEASONALITY,FORECAST.ETS.STAT,FORECAST.LINEAR,GAMMA,GAMMA.DIST,GAMMA.INV,GAMMALN.PRECISE,GAUSS,GROUPBY,HSTACK,HYPGEOM.DIST,IFNA,IFS,IMAGE,IMCOSH,IMCOT,IMCSC,IMCSCH,IMSEC,IMSECH,IMSINH,IMTAN,ISFORMULA,ISOMITTED,ISOWEEKNUM,LOGNORM.DIST,LOGNORM.INV,MAKEARRAY,MAP,MAXIFS,MINIFS,MODE.MULT,MODE.SNGL,MUNIT,NEGBINOM.DIST,NORM.DIST,NORM.INV,NORM.S.DIST,NORM.S.INV,NUMBERVALUE,PDURATION,PERCENTILE.EXC,PERCENTILE.INC,PERCENTRANK.EXC,PERCENTRANK.INC,PERMUTATIONA,PHI,PIVOTBY,POISSON.DIST,QUARTILE.EXC,QUARTILE.INC,RANDARRAY,RANK.AVG,RANK.EQ,REDUCE,REGEXEXTRACT,REGEXREPLACE,REGEXTEST,RRI,SCAN,SEC,SECH,SEQUENCE,SINGLE,SKEW.P,SORT,SORTBY,STDEV.P,STDEV.S,STOCKHISTORY,SWITCH,T.DIST,T.DIST.2T,T.DIST.RT,T.INV,T.INV.2T,T.TEST,TAKE,TEXTAFTER,TEXTBEFORE,TEXTJOIN,TEXTSPLIT,TOCOL,TOROW,TRIMRANGE,UNICHAR,UNICODE,UNIQUE,VALUETOTEXT,VAR.P,VAR.S,VSTACK,WEBSERVICE,WEIBULL.DIST,WRAPCOLS,WRAPROWS,XLOOKUP,XMATCH,XOR,Z.TEST';
  // Funciones que LibreOffice escribe con otro nombre: EXCEL=ODF
  const T_DIST = 'CHIDIST=LEGACY.CHIDIST,CHIINV=LEGACY.CHIINV,FDIST=LEGACY.FDIST,FINV=LEGACY.FINV,NORMSDIST=LEGACY.NORMSDIST,NORMSINV=LEGACY.NORMSINV,TDIST=LEGACY.TDIST,CHITEST=LEGACY.CHITEST,ARRAYTOTEXT=_xlfn.arraytotext,BINOM.DIST.RANGE=_xlfn.binom.dist.range,BYCOL=_xlfn.bycol,BYROW=_xlfn.byrow,F.DIST=FDIST,F.INV=FINV,FIELDVALUE=_xlfn.fieldvalue,GROUPBY=_xlfn.groupby,IMAGE=_xlfn.image,IMCOSH=_xlfn.imcosh,IMCOT=_xlfn.imcot,IMCSC=_xlfn.imcsc,IMCSCH=_xlfn.imcsch,IMSEC=_xlfn.imsec,IMSECH=_xlfn.imsech,IMSINH=_xlfn.imsinh,IMTAN=_xlfn.imtan,ISOMITTED=_xlfn.isomitted,MAKEARRAY=_xlfn.makearray,MAP=_xlfn.map,PIVOTBY=_xlfn.pivotby,REDUCE=_xlfn.reduce,REGEXEXTRACT=_xlfn.regexextract,REGEXREPLACE=_xlfn.regexreplace,REGEXTEST=_xlfn.regextest,SCAN=_xlfn.scan,SINGLE=_xlfn.single,SKEW.P=SKEWP,STOCKHISTORY=_xlfn.stockhistory,TRIMRANGE=_xlfn.trimrange,VALUETOTEXT=_xlfn.valuetotext';
  // </tabla>
  const COM = new Set(T_COM ? T_COM.split(',') : []);
  const XLFN = new Set(T_XLFN ? T_XLFN.split(',') : []);
  const DIST = new Map(), DIST_INV = new Map();
  for (const par of T_DIST ? T_DIST.split(',') : []) { const [a, b] = par.split('='); DIST.set(a, b); DIST_INV.set(b.toUpperCase(), a); }
  COM.add('LET'); XLFN.add('LET');   // LET no está en la tabla de funciones de .xls (no existe en Excel 97-2003) pero sí en los .xlsx y en LibreOffice
  const SIN_ZONA = new Set(['FILTER', 'SORT']);

  const ID = /[A-Za-z0-9_.À-￿]/;
  const NOMBRE_SIMPLE = /^[A-Za-zÀ-￿_][A-Za-z0-9_.À-￿]*$/;
  const CELDA = '\\$?[A-Za-z]{1,3}\\$?\\d+';
  const COLS = '\\$?[A-Za-z]{1,3}', FILAS = '\\$?\\d+';
  const RANGO = `(?:${CELDA}(?::${CELDA})?|${COLS}:${COLS}|${FILAS}:${FILAS})`;
  const HOJA_EXCEL = "(?:'(?:[^']|'')+'|[A-Za-zÀ-￿_][A-Za-z0-9_.À-￿]*)";
  const RE_REF_EXCEL = new RegExp(`^(?:(${HOJA_EXCEL})(?::(${HOJA_EXCEL}))?!)?(${RANGO})(?![A-Za-z0-9_.(À-￿])`);
  const RE_REF_3D_QUOTED = new RegExp(`^'((?:[^']|'')+)'!(${RANGO})(?![A-Za-z0-9_.(À-￿])`);

  const desquotar = (h) => (h[0] === "'" ? h.slice(1, -1).replace(/''/g, "'") : h);
  const citarOds = (n) => (NOMBRE_SIMPLE.test(n) ? n : "'" + n.replace(/'/g, "''") + "'");
  const citarExcel = (n) => (NOMBRE_SIMPLE.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) && !/^[Rr]\d*[Cc]\d*$/.test(n) && !/^(?:TRUE|FALSE)$/i.test(n) ? n : "'" + n.replace(/'/g, "''") + "'");

  // ---- Excel → OpenFormula ------------------------------------------------------------------------------------------------
  function aOpenFormula(texto) {
    let s = String(texto).replace(/^=/, ''), i = 0, out = '';
    const pila = [];   // 'f' llamada a función, 'g' paréntesis de agrupar
    let ultimo = '';   // el último token significativo: 'ref', ')', 'otro'
    const n = s.length;
    while (i < n) {
      const ch = s[i];
      if (ch === '"') {
        let j = i + 1;
        while (j < n) { if (s[j] === '"') { if (s[j + 1] === '"') j += 2; else break; } else j++; }
        out += s.slice(i, j + 1); i = j + 1; ultimo = 'otro'; continue;
      }
      if (ch === '{') {   // matriz en línea: columnas «;», filas «|»
        let j = i + 1, t = '{';
        while (j < n && s[j] !== '}') {
          if (s[j] === '"') { let k = j + 1; while (k < n) { if (s[k] === '"') { if (s[k + 1] === '"') k += 2; else break; } else k++; } t += s.slice(j, k + 1); j = k + 1; continue; }
          t += s[j] === ',' ? ';' : s[j] === ';' ? '|' : s[j]; j++;
        }
        out += t + '}'; i = j + 1; ultimo = 'otro'; continue;
      }
      if (ch === ' ') {
        let j = i; while (j < n && s[j] === ' ') j++;
        // un espacio entre dos referencias (o un paréntesis que cierra y una referencia) es la intersección
        if ((ultimo === 'ref' || ultimo === ')') && j < n && (RE_REF_EXCEL.test(s.slice(j)) || RE_REF_3D_QUOTED.test(s.slice(j)) || s[j] === '(' )) { out += '!'; ultimo = 'otro'; } else out += s.slice(i, j);
        i = j; continue;
      }
      if (ch === '(') {
        const funcion = out.length > 0 && (ID.test(out[out.length - 1]) || out[out.length - 1] === ')') && ultimo !== 'otro-op';
        pila.push(funcion ? 'f' : 'g'); out += '('; i++; ultimo = 'otro-op'; continue;
      }
      if (ch === ')') { pila.pop(); out += ')'; i++; ultimo = ')'; continue; }
      if (ch === ',') { out += pila.length && pila[pila.length - 1] === 'f' ? ';' : '~'; i++; ultimo = 'otro-op'; continue; }
      if (!ID.test(ch) && ch !== "'" && ch !== '$' && ch !== '_') {
        out += ch; i++; ultimo = ch === '%' ? 'otro' : 'otro-op'; continue;
      }
      // referencia (con o sin hoja), nombre, función o constante
      const resto = s.slice(i);
      let m = RE_REF_EXCEL.exec(resto);
      if (m && (i === 0 || !ID.test(s[i - 1]))) {
        const hojaA = m[1] ? desquotar(m[1]) : null;
        let hojaB = m[2] ? desquotar(m[2]) : null;
        let partes = m[3].split(':');
        // «'Hoja A:Hoja B'!»: una hoja entre comillas con dos puntos dentro es una referencia 3D
        let a = hojaA;
        if (a && !hojaB && m[1][0] === "'" && a.indexOf(':') > 0) { const k = a.indexOf(':'); hojaB = a.slice(k + 1); a = a.slice(0, k); }
        if (a === null) out += '[.' + partes[0] + (partes[1] ? ':.' + partes[1] : '') + ']';
        else if (hojaB) out += '[$' + citarOds(a) + '.' + partes[0] + ':$' + citarOds(hojaB) + '.' + (partes[1] || partes[0]) + ']';
        else out += '[$' + citarOds(a) + '.' + partes[0] + (partes[1] ? ':.' + partes[1] : '') + ']';
        i += m[0].length; ultimo = 'ref'; continue;
      }
      m = RE_REF_3D_QUOTED.exec(resto);
      if (m) { /* ya cubierto arriba */ }
      // identificador: función, nombre definido o constante lógica
      let j = i; while (j < n && (ID.test(s[j]) || s[j] === '$')) j++;
      if (j === i) { out += ch; i++; ultimo = 'otro'; continue; }
      let palabra = s.slice(i, j);
      if (s[j] === '[') return null;   // referencia estructurada de tabla (Tabla[Columna]): no existe en ODF
      if (s[j] === '(') {
        let base = palabra.replace(/^_xlfn\.(_xlws\.)?/i, '').replace(/^_xludf\./i, '');
        const u = base.toUpperCase();
        palabra = DIST.has(u) ? DIST.get(u) : COM.has(u) ? 'COM.MICROSOFT.' + base : base;
      } else if (/^(TRUE|FALSE)$/i.test(palabra) && s[j] !== '!') {
        palabra = palabra.toUpperCase() + '()';
      }
      out += palabra; i = j; ultimo = 'otro';
    }
    return out;
  }

  // ---- OpenFormula → Excel ------------------------------------------------------------------------------------------------
  // Una referencia entre corchetes: «$Hoja.A1:.B2», «.A1», «$'Mi hoja'.A1:$'Otra'.A1», «.A:.A», «.1:.1». null si no se entiende.
  function referenciaAExcel(interior) {
    if (/#REF!/.test(interior)) return '#REF!';
    const partes = []; let actual = '', comillas = false;
    for (const ch of interior) { if (ch === "'") comillas = !comillas; if (ch === ':' && !comillas) { partes.push(actual); actual = ''; } else actual += ch; }
    partes.push(actual);
    if (partes.length > 2) return null;
    const una = (p) => {
      const m = /^(\$?)((?:'(?:[^']|'')+')|[^.']*)\.(.+)$/.exec(p);
      if (!m) return null;
      return { hoja: m[2] ? desquotar(m[2]) : null, celda: m[3] };
    };
    const a = una(partes[0]); if (!a) return null;
    const b = partes[1] === undefined ? null : una(partes[1]); if (partes[1] !== undefined && !b) return null;
    const celdaOk = new RegExp(`^(?:${CELDA}|${COLS}|${FILAS})$`);
    if (!celdaOk.test(a.celda) || (b && !celdaOk.test(b.celda))) return null;
    if (!b) return (a.hoja ? citarExcel(a.hoja) + '!' : '') + a.celda;
    if (b.hoja && a.hoja && b.hoja !== a.hoja) {
      if (a.celda !== b.celda) return null;   // un rango entre hojas distintas con celdas distintas no existe en Excel
      return (/^[A-Za-zÀ-￿_][A-Za-z0-9_.À-￿]*$/.test(a.hoja) && /^[A-Za-zÀ-￿_][A-Za-z0-9_.À-￿]*$/.test(b.hoja) ? a.hoja + ':' + b.hoja : "'" + (a.hoja + ':' + b.hoja).replace(/'/g, "''") + "'") + '!' + a.celda;
    }
    return (a.hoja ? citarExcel(a.hoja) + '!' : '') + a.celda + ':' + b.celda;
  }

  function aExcel(texto) {
    const s = String(texto).replace(/^of:=/, '').replace(/^=/, '');
    let i = 0, out = '';
    const n = s.length;
    const pila = [];   // { tipo: 'f' | 'g', inicioArg: posición en `out` donde empieza el argumento, union: bool }
    const cerrarArg = (p) => { if (p.union) { out = out.slice(0, p.inicioArg) + '(' + out.slice(p.inicioArg) + ')'; p.union = false; } };
    while (i < n) {
      const ch = s[i];
      if (ch === '"') {
        let j = i + 1;
        while (j < n) { if (s[j] === '"') { if (s[j + 1] === '"') j += 2; else break; } else j++; }
        out += s.slice(i, j + 1); i = j + 1; continue;
      }
      if (ch === '[') {
        const j = s.indexOf(']', i); if (j < 0) return null;
        const r = referenciaAExcel(s.slice(i + 1, j)); if (r === null) return null;
        out += r; i = j + 1; continue;
      }
      if (ch === '{') {
        let j = i + 1, t = '{';
        while (j < n && s[j] !== '}') {
          if (s[j] === '"') { let k = j + 1; while (k < n) { if (s[k] === '"') { if (s[k + 1] === '"') k += 2; else break; } else k++; } t += s.slice(j, k + 1); j = k + 1; continue; }
          t += s[j] === ';' ? ',' : s[j] === '|' ? ';' : s[j]; j++;
        }
        out += t + '}'; i = j + 1; continue;
      }
      if (ch === '(') {
        const prev = out[out.length - 1];
        pila.push({ tipo: prev !== undefined && (ID.test(prev) || prev === ')') ? 'f' : 'g', inicioArg: out.length + 1, union: false }); out += '('; i++; continue;
      }
      if (ch === ')') { const p = pila.pop(); if (p) cerrarArg(p); out += ')'; i++; continue; }
      if (ch === ';') { const p = pila[pila.length - 1]; if (p) { cerrarArg(p); } out += ','; if (p) p.inicioArg = out.length; i++; continue; }
      if (ch === '~') { const p = pila[pila.length - 1]; out += ','; if (p && p.tipo === 'f') p.union = true; i++; continue; }
      if (ch === '!') { out += ' '; i++; continue; }
      if (!ID.test(ch)) { out += ch; i++; continue; }
      let j = i; while (j < n && ID.test(s[j])) j++;
      let palabra = s.slice(i, j);
      if (s[j] === '(') {
        if (/^(TRUE|FALSE)$/i.test(palabra) && s[j + 1] === ')') { out += palabra.toUpperCase(); i = j + 2; continue; }   // TRUE() → TRUE
        let u = palabra.toUpperCase();
        if (DIST_INV.has(u)) { palabra = DIST_INV.get(u); u = palabra.toUpperCase(); }
        if (/^(ORG\.(OPENOFFICE|LIBREOFFICE)\.|_xlfn\.lambda$)/i.test(palabra)) return null;
        if (u.startsWith('COM.MICROSOFT.')) { const base = palabra.slice(14); const b = base.toUpperCase(); palabra = XLFN.has(b) || !XLFN.size ? (SIN_ZONA.has(b) ? '_xlfn._xlws.' : '_xlfn.') + base : base; }
        else if (XLFN.has(u)) palabra = (SIN_ZONA.has(u) ? '_xlfn._xlws.' : '_xlfn.') + palabra;
      }
      out += palabra; i = j;
    }
    return out;
  }

  globalThis.OdsFormulas = { aOpenFormula, aExcel };
})();

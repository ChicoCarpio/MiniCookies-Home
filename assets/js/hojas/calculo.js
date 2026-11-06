/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Motor de cálculo de Mini Cookies Hojas (fase 2 del PLAN): evalúa fórmulas de Excel, lleva el grafo de dependencias y recalcula solo lo que
// depende de lo que cambió, en orden (Kahn). Sin DOM: corre igual en Node y en un Worker.
//
// El análisis de la fórmula (texto → árbol, con las precedencias reales de Excel) es el de xls-formulas.js (XlsFormulas._tokenizar/_analizar),
// medido con 4.339 fórmulas reales. Aquí se agrega la evaluación. Requiere xls-formulas.js y, para TEXT, xlsx-formato.js.
//
//   const motor = CalculoHojas.crear(fuente);
//   fuente = {
//     hojas: () => ['Hoja1', …],                  // nombres en el orden del libro (el índice es la hoja)
//     valor: (h, f, c) => número | texto | lógico | null | ErrorCalc,
//     poner: (h, f, c, valor, soloVista) => …,    // el motor deja aquí el resultado de una fórmula (soloVista: cambió solo por ser volátil)
//     formulas: () => iterable de { h, f, c, texto, valor },   // todas las fórmulas del libro (texto sin "="); valor: el guardado
//     limite: (h) => { filas, cols },              // hasta dónde llega lo usado (para A:A o 1:1)
//     nombre: (nombre, h) => texto | null,         // un nombre definido (Hoja1!$A$1:$B$5), primero el local de la hoja h
//     fecha1904: false, hoy: () => new Date(), idioma: 'es' | 'en' }
//   motor.construir()                              // lee todas las fórmulas y arma el grafo (no recalcula: se confía en los valores guardados)
//   motor.ponerFormula(h, f, c, texto | null)      // la celda pasa a tener (o deja de tener) esa fórmula
//   motor.recalcular([{ h, f, c }, …])            // las celdas que cambiaron → { calculadas, noSoportadas: [...], circulares }
//   motor.recalcularTodo()                         // todas las fórmulas (después de un cambio de estructura)
//   motor.derrame(h, f, c) → { h, f0, c0, f1, c1, ok } | undefined   // el rango que desborda una matriz dinámica (ok: false = #SPILL!)
// Matrices dinámicas (fase 3.1): una fórmula con dinamica (escrita en el editor, o guardada por Excel 365 con cm) que da una matriz la
// DESBORDA: el resultado va a las celdas vecinas por fuente.derramar(h, f, c, valor | undefined, soloVista); si alguna está ocupada
// (fuente.ocupada) o es de otro desborde, la celda da #SPILL!. Las fórmulas del archivo sin cm siguen la intersección implícita.
//   motor.evaluar(texto, { h, f, c })             // evalúa una fórmula suelta (no la guarda)
//   CalculoHojas.ErrorCalc / CalculoHojas.error('#N/A') / CalculoHojas.SOPORTADAS
// Lo que no sabe calcular (una función conocida por Excel que aún no está aquí, matrices dinámicas) deja el valor que tenía y se informa
// en noSoportadas; Excel la recalcula al abrir el archivo (el escritor marca fullCalcOnLoad).
(function () {
  const MAX_FILAS = 1048576, MAX_COLS = 16384;
  const MAX_EXPANSIONES = 5000000;   // nombres definidos evaluados por fórmula: más que esto es un archivo hecho para colgar la página
  const FORMA_ENTERA = 65536;   // celdas de un rango: hasta aquí una matriz conserva su forma aunque pase de lo usado; más grande, se corta en lo usado
  const ERRORES = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!'];
  class ErrorCalc { constructor(codigo) { this.codigo = codigo; } toString() { return this.codigo; } }
  const ERR = {}; for (const e of ERRORES) ERR[e] = new ErrorCalc(e);
  const error = (codigo) => ERR[codigo] || new ErrorCalc(codigo);
  const E = { div: ERR['#DIV/0!'], valor: ERR['#VALUE!'], ref: ERR['#REF!'], nombre: ERR['#NAME?'], num: ERR['#NUM!'], na: ERR['#N/A'], nulo: ERR['#NULL!'] };
  const esError = (v) => v instanceof ErrorCalc;
  class NoSoportada extends Error {}
  class Lanzado { constructor(e) { this.e = e; } }   // un error de Excel que corta la evaluación de un argumento
  const lanzar = (e) => { throw new Lanzado(e); };

  // "B2:D9" o "B2" → { f0, c0, f1, c1 } (desde 0); null si no es una referencia A1
  function rangoTexto(t) {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(String(t).trim());
    if (!m) return null;
    const col = (x) => { let n = 0; for (const ch of x.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
    const a = { f: +m[2] - 1, c: col(m[1]) }, b = m[3] ? { f: +m[4] - 1, c: col(m[3]) } : a;
    return { f0: Math.min(a.f, b.f), c0: Math.min(a.c, b.c), f1: Math.max(a.f, b.f), c1: Math.max(a.c, b.c) };
  }
  const clave = (h, f, c) => (h * MAX_FILAS + f) * MAX_COLS + c;
  const deClave = (k) => { const c = k % MAX_COLS, r = (k - c) / MAX_COLS, f = r % MAX_FILAS; return { h: (r - f) / MAX_FILAS, f, c }; };

  // ---- conversiones de Excel --------------------------------------------------------------------------------------
  const esRef = (v) => v && typeof v === 'object' && v.tipo === 'ref';
  const esMat = (v) => v && typeof v === 'object' && v.tipo === 'mat';
  const esUnion = (v) => v && typeof v === 'object' && v.tipo === 'union';
  const RE_NUMERO = /^\s*[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?\s*$/;
  function textoANumero(t) {
    const s = t.trim();
    if (RE_NUMERO.test(s)) return parseFloat(s);
    const pct = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+))\s*%$/.exec(s);
    if (pct) return parseFloat(pct[1]) / 100;
    const miles = /^[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?$/.exec(s);
    if (miles) return parseFloat(s.replace(/,/g, ''));
    const fecha = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
    if (fecha) return serialDe(+fecha[1], +fecha[2], +fecha[3]);
    return null;
  }
  function formatoGeneral(n) {
    if (!Number.isFinite(n)) return '#NUM!';
    if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
    let s = Number(n.toPrecision(15)).toString();
    if (/e/.test(s)) { const [m, x] = s.split('e'); s = m + 'E' + (x[0] === '-' ? '-' : '+') + x.replace(/^[+-]/, '').padStart(2, '0'); }
    return s;
  }

  // ---- fechas (número de serie de Excel, con el 29-feb-1900 que no existió) -------------------------------------------
  const DIA = 86400000;
  const BASE_1900 = Date.UTC(1899, 11, 30), BASE_1904 = Date.UTC(1904, 0, 1);
  let fecha1904 = false;
  function serialDe(a, m, d) {
    const ms = Date.UTC(a, m - 1, d);
    if (fecha1904) return Math.round((ms - BASE_1904) / DIA);
    let s = Math.round((ms - BASE_1900) / DIA);
    if (s <= 60) s -= 1;   // antes del 1-mar-1900 Excel cuenta un día menos (por el 29-feb-1900 inventado)
    return s;
  }
  function fechaDe(serial) {
    let s = Math.floor(serial);
    if (fecha1904) return new Date(BASE_1904 + s * DIA);
    if (s === 60) return { a: 1900, m: 2, d: 29, ds: 3 };
    if (s < 60) s += 1;
    const d = new Date(BASE_1900 + s * DIA);
    return d;
  }
  function partes(serial) {
    const d = fechaDe(serial);
    if (!(d instanceof Date)) return d;
    return { a: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), ds: d.getUTCDay() };
  }

  // ---- qué argumentos se aplican a cada elemento, medido en Excel 365 (pruebas/oraculo/excel-medir-matrices.ps1) ------------------
  //   s  rango o matriz: se aplica a cada elemento
  //   x  matriz: a cada elemento; rango de varias celdas: #VALUE! (las del antiguo complemento de análisis)
  //   f  matriz: a cada elemento; rango: la función lo recibe entero (N y T toman su primera celda)
  //   v  rama: si otro argumento es matriz, se toma el elemento de la misma posición; si no, se devuelve entero
  //   e  como s, pero si la función devuelve una matriz o una referencia, ese elemento da #VALUE!
  //   a  la función lo recibe entero
  //   z  la función lo recibe entero, pero con un rango Excel la guarda con cm
  // Patrón: una letra por posición; lo que sigue a «|» se repite para el resto. Lo que no está en la tabla recibe todo entero.
  const ELEVA = new Map();
  const elevan = (patron, nombres) => { for (const n of nombres.split(' ')) ELEVA.set(n, patron); };
  elevan('|s', 'ABS ACOS ACOSH ACOT ACOTH ADDRESS ARABIC ASIN ASINH ATAN ATAN2 ATANH BASE BETA.DIST BETA.INV BETADIST BETAINV BINOM.DIST BINOM.INV BINOMDIST BITAND BITLSHIFT BITOR BITRSHIFT BITXOR CEILING CEILING.MATH CEILING.PRECISE CHAR CHIDIST CHIINV CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CLEAN CODE COMBIN COMBINA CONCATENATE CONFIDENCE CONFIDENCE.NORM CONFIDENCE.T COS COSH COT COTH CRITBINOM CSC CSCH DATE DATEDIF DATEVALUE DAY DAYS DAYS360 DB DDB DECIMAL DEGREES DOLLAR ERROR.TYPE EVEN EXACT EXP EXPON.DIST EXPONDIST F.DIST F.DIST.RT F.INV F.INV.RT FACT FDIST FIND FINV FISHER FISHERINV FIXED FLOOR FLOOR.MATH FLOOR.PRECISE FV GAMMA GAMMA.DIST GAMMA.INV GAMMADIST GAMMAINV GAMMALN GAMMALN.PRECISE GAUSS HOUR HYPERLINK HYPGEOM.DIST HYPGEOMDIST INT IPMT ISBLANK ISERR ISERROR ISLOGICAL ISNA ISNONTEXT ISNUMBER ISO.CEILING ISOWEEKNUM ISPMT ISTEXT LEFT LEFTB LEN LENB LN LOG LOG10 LOGINV LOGNORM.DIST LOGNORM.INV LOGNORMDIST LOWER MID MIDB MINUTE MOD MONTH NEGBINOM.DIST NEGBINOMDIST NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NORMDIST NORMINV NORMSDIST NORMSINV NOT NPER NUMBERVALUE ODD PDURATION PERMUT PERMUTATIONA PHI PMT POISSON POISSON.DIST POWER PPMT PROPER PV RADIANS RANDBETWEEN RATE REGEXREPLACE REGEXTEST REPLACE REPT RIGHT RIGHTB ROMAN ROUND ROUNDDOWN ROUNDUP RRI SEARCH SEC SECH SECOND SIGN SIN SINH SLN SQRT SQRTPI STANDARDIZE SUBSTITUTE SYD T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T TAN TANH TDIST TEXT TIME TIMEVALUE TINV TRIM TRUNC UNICHAR UNICODE UPPER VALUE VALUETOTEXT WEEKDAY WEIBULL WEIBULL.DIST YEAR');
  elevan('|x', 'BIN2DEC BIN2HEX BIN2OCT CONVERT CUMIPMT CUMPRINC DEC2BIN DEC2HEX DEC2OCT DELTA DOLLARDE DOLLARFR EDATE EFFECT EOMONTH ERF ERF.PRECISE ERFC ERFC.PRECISE FACTDOUBLE GESTEP HEX2BIN HEX2DEC HEX2OCT ISEVEN ISODD MROUND NOMINAL OCT2BIN OCT2DEC OCT2HEX QUOTIENT WEEKNUM YEARFRAC');
  elevan('xxxa', 'SERIESSUM');
  elevan('xa', 'FVSCHEDULE XNPV');
  elevan('xxa', 'NETWORKDAYS WORKDAY');
  elevan('xxsa', 'NETWORKDAYS.INTL WORKDAY.INTL');
  elevan('ax', 'DAVERAGE DCOUNT DCOUNTA DGET DMAX DMIN DPRODUCT DSTDEV DSTDEVP DSUM DVAR DVARP');
  elevan('f', 'N T');
  elevan('as', 'IRR LARGE PERCENTILE PERCENTILE.EXC PERCENTILE.INC QUARTILE QUARTILE.EXC QUARTILE.INC SMALL TRIMMEAN COUNTIF');
  elevan('asa', 'AVERAGEIF SUMIF');
  elevan('ass', 'MIRR PERCENTRANK PERCENTRANK.EXC PERCENTRANK.INC Z.TEST ZTEST');
  elevan('|as', 'COUNTIFS');
  elevan('a|as', 'AVERAGEIFS MAXIFS MINIFS SUMIFS');
  elevan('aass', 'T.TEST TTEST');
  elevan('saa', 'FORECAST FORECAST.LINEAR LOOKUP');
  elevan('sas', 'RANK RANK.AVG RANK.EQ');
  elevan('sax', 'MATCH');
  elevan('sass', 'HLOOKUP VLOOKUP XMATCH');
  elevan('saazss', 'XLOOKUP');
  elevan('asss', 'INDEX');
  elevan('ssax', 'AGGREGATE');
  elevan('sa', 'SUBTOTAL');
  elevan('saaa', 'NPV');
  elevan('as', 'TEXTJOIN');
  elevan('szssss', 'TEXTAFTER TEXTBEFORE');
  // las que devuelven una matriz, con una matriz en un argumento de valor suelto: unas dan el primer valor de cada resultado (s) y otras #VALUE! (e)
  elevan('|s', 'MUNIT REGEXEXTRACT SEQUENCE');
  elevan('ass', 'TOCOL TOROW UNIQUE WRAPCOLS WRAPROWS');
  elevan('assa', 'EXPAND');
  elevan('aaas', 'SORT');
  elevan('aee', 'DROP TAKE');
  elevan('|e', 'INDIRECT');
  elevan('a|e', 'OFFSET');
  elevan('z', 'COLUMN ROW TYPE');
  elevan('svv', 'IF');
  elevan('sv', 'IFERROR IFNA');
  elevan('s|v', 'CHOOSE');
  elevan('|sv', 'IFS');
  elevan('s|sv', 'SWITCH');
  const cacheCiclo = new Map();
  function posicionDe(patron, k) {
    let p = cacheCiclo.get(patron);
    if (!p) { const [fijo, ciclo] = patron.split('|'); p = { fijo, ciclo: ciclo || '' }; cacheCiclo.set(patron, p); }
    if (k < p.fijo.length) return p.fijo[k];
    return p.ciclo ? p.ciclo[(k - p.fijo.length) % p.ciclo.length] : 'a';
  }
  const DEVUELVEN_MATRIZ = new Set('CHOOSECOLS CHOOSEROWS DROP EXPAND FILTER FREQUENCY GROWTH HSTACK LINEST LOGEST MODE.MULT RANDARRAY REGEXEXTRACT SEQUENCE SORT SORTBY TAKE TEXTSPLIT TOCOL TOROW TRANSPOSE TREND UNIQUE VSTACK WRAPCOLS WRAPROWS'.split(' '));
  const SIEMPRE_DINAMICA = new Set(['IFS', 'SWITCH']);   // Excel las guarda con cm aunque todos sus argumentos sean valores sueltos
  const DENTRO_SIN_CM = new Set(['ANCHORARRAY', 'MINVERSE', 'MMULT', 'MUNIT']);   // solas llevan cm; recibidas enteras por otra función (INDEX(MMULT(…);1;1), SUM(A1#)) no

  // ---- el motor -----------------------------------------------------------------------------------------------------
  function crear(fuente) {
    fecha1904 = !!fuente.fecha1904;
    const formulas = new Map();      // clave → { h, f, c, texto, ast, celdas: [clave], rangos: [r], volatil, mala }
    const celdaDeps = new Map();     // clave de celda → Set de claves de fórmulas que la leen
    let rangoDeps = [];              // { h, f0, c0, f1, c1, dep } (las que leen un rango)
    let rangosMuertos = 0;
    const cacheAst = new Map();
    const derrames = new Map();      // clave del ancla → { h, f0, c0, f1, c1, ok }
    let pendientesDerrame = [];      // celdas que un desborde dejó vacías al quitar su fórmula (sus dependientes se recalculan)
    const hoy = fuente.hoy || (() => new Date());
    let nombresHoja = null;
    const indiceHoja = (nombre) => {
      if (!nombresHoja) nombresHoja = new Map(fuente.hojas().map((n, i) => [n.toLowerCase(), i]));
      const i = nombresHoja.get(String(nombre).toLowerCase());
      return i === undefined ? -1 : i;
    };

    function analizar(texto, h, f, c) {
      if (texto.indexOf('[') >= 0 && h !== undefined) texto = expandirEstructuradas(texto, h, f, c);
      let ast = cacheAst.get(texto);
      if (ast !== undefined) return ast;
      try { ast = XlsFormulas._analizar(XlsFormulas._tokenizar(texto)); } catch (e) { ast = { t: 'mala', motivo: e.message }; }
      if (cacheAst.size > 50000) cacheAst.clear();
      cacheAst.set(texto, ast);
      return ast;
    }

    // ---- referencias estructuradas (tablas): Tabla[Col], Tabla[[#This Row],[Col]], [@Col], Tabla[#Totals]… → 'Hoja'!$C$2:$C$9 -------------
    let tablasLibro = null;
    const listaTablas = () => { if (!tablasLibro) tablasLibro = fuente.tablas ? fuente.tablas() : []; return tablasLibro; };
    const citar = (n) => (/^[A-Za-z_\u00C0-\uFFFF][A-Za-z0-9_.\u00C0-\uFFFF]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'");
    const colLetras = (n) => { let s = ''; for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
    function expandirEstructuradas(texto, h, f) {
      let salida = '', i = 0;
      while (i < texto.length) {
        const ch = texto[i];
        if (ch === '"') { let j = i + 1; while (j < texto.length) { if (texto[j] === '"') { if (texto[j + 1] === '"') j += 2; else break; } else j++; } salida += texto.slice(i, j + 1); i = j + 1; continue; }
        if (ch === "'") { let j = i + 1; while (j < texto.length) { if (texto[j] === "'") { if (texto[j + 1] === "'") j += 2; else break; } else j++; } salida += texto.slice(i, j + 1); i = j + 1; continue; }
        if (ch === '[') {
          // [n]Hoja!A1 (vínculo a otro libro) → '[n]Hoja'!A1, que el analizador lee como un nombre de hoja
          const ext = /^\[(\d+)\]([A-Za-z_\u00C0-\uFFFF][A-Za-z0-9_.\u00C0-\uFFFF]*)!/.exec(texto.slice(i));
          if (ext && !/[A-Za-z0-9_.\]]$/.test(salida)) { salida += "'[" + ext[1] + ']' + ext[2] + "'!"; i += ext[0].length; continue; }
          // el nombre de la tabla va justo antes (puede faltar: la tabla que contiene la celda)
          const m = /([A-Za-z_\\\u00C0-\uFFFF][A-Za-z0-9_.\u00C0-\uFFFF]*)$/.exec(salida);
          let nombre = m ? m[1] : '';
          let nivel = 0, j = i;
          for (; j < texto.length; j++) { if (texto[j] === "'" ) { j++; continue; } if (texto[j] === '[') nivel++; else if (texto[j] === ']') { nivel--; if (nivel === 0) break; } }
          const interior = texto.slice(i + 1, j);
          const tabla = nombre ? listaTablas().find((t) => t.nombre.toUpperCase() === nombre.toUpperCase()) : listaTablas().find((t) => t.h === h && f >= t.ref.f0 && f <= t.ref.f1);
          const ref = tabla ? resolverEstructurada(tabla, interior, f) : null;
          if (nombre && tabla) salida = salida.slice(0, salida.length - nombre.length);
          salida += ref || '#REF!';
          i = j + 1;
          continue;
        }
        salida += ch; i++;
      }
      return salida;
    }
    function resolverEstructurada(t, interior, f) {
      const items = [];
      const s = interior.trim();
      if (s.startsWith('[')) { const re = /\[((?:'.|[^\]])*)\]|:|,/g; let m; while ((m = re.exec(s))) items.push(m[0] === ':' ? ':' : m[0] === ',' ? ',' : m[1]); }
      else if (s.startsWith('@')) { items.push('#This Row'); if (s.length > 1) items.push(',', s.slice(1).replace(/^\[|\]$/g, '')); }
      else items.push(s);
      const datos0 = t.ref.f0 + (t.encabezados ? 1 : 0), datos1 = t.ref.f1 - (t.totales ? 1 : 0);
      let filas = null, c0 = null, c1 = null;
      const desesc = (x) => x.replace(/'(.)/g, '$1').trim();
      const colDe = (nombre) => { const k = t.columnas.findIndex((x) => x.toUpperCase() === desesc(nombre).toUpperCase()); return k < 0 ? null : t.ref.c0 + k; };
      for (let k = 0; k < items.length; k++) {
        const it = items[k];
        if (it === ',' || it === ':') continue;
        const u = it.trim().toUpperCase();
        if (u === '#ALL') filas = [t.ref.f0, t.ref.f1];
        else if (u === '#DATA') filas = filas && filas[0] === t.ref.f0 ? [t.ref.f0, datos1] : [datos0, datos1];
        else if (u === '#HEADERS') { if (!t.encabezados) return null; filas = filas ? [t.ref.f0, filas[1]] : [t.ref.f0, t.ref.f0]; }
        else if (u === '#TOTALS') { if (!t.totales) return null; filas = filas ? [filas[0], t.ref.f1] : [t.ref.f1, t.ref.f1]; }
        else if (u === '#THIS ROW') { if (f < datos0 || f > datos1) return null; filas = [f, f]; }
        else {
          const c = colDe(it);
          if (c === null) return null;
          if (items[k - 1] === ':' && c0 !== null) c1 = c; else { c0 = c; c1 = c; }
        }
      }
      if (!filas) filas = [datos0, datos1];
      if (c0 === null) { c0 = t.ref.c0; c1 = t.ref.c1; }
      const hoja = citar(t.nombreHoja);
      const a = '$' + colLetras(Math.min(c0, c1)) + '$' + (filas[0] + 1), b = '$' + colLetras(Math.max(c0, c1)) + '$' + (filas[1] + 1);
      return hoja + '!' + a + (a === b ? '' : ':' + b);
    }

    // ---- referencias del árbol --------------------------------------------------------------------------------------
    // Una referencia del árbol a { h, f0, c0, f1, c1 } (o null si es #REF!). h: la hoja de la fórmula salvo que nombre otra.
    function rangoDe(nodo, h) {
      let hojas = [h];
      if (nodo.hoja) {
        const a = indiceHoja(nodo.hoja.a), b = indiceHoja(nodo.hoja.b);
        if (a < 0 || b < 0) return null;
        hojas = []; for (let k = Math.min(a, b); k <= Math.max(a, b); k++) hojas.push(k);
      }
      let f0, c0, f1, c1;
      switch (nodo.modo) {
        case 'celda': f0 = f1 = nodo.f0; c0 = c1 = nodo.c0; break;
        case 'area': f0 = Math.min(nodo.f0, nodo.f1); f1 = Math.max(nodo.f0, nodo.f1); c0 = Math.min(nodo.c0, nodo.c1); c1 = Math.max(nodo.c0, nodo.c1); break;
        case 'cols': f0 = 0; f1 = MAX_FILAS - 1; c0 = Math.min(nodo.c0, nodo.c1); c1 = Math.max(nodo.c0, nodo.c1); break;
        case 'filas': c0 = 0; c1 = MAX_COLS - 1; f0 = Math.min(nodo.f0, nodo.f1); f1 = Math.max(nodo.f0, nodo.f1); break;
        default: return null;
      }
      if (f1 >= MAX_FILAS || c1 >= MAX_COLS) return null;
      return hojas.map((hh) => ({ tipo: 'ref', h: hh, f0, c0, f1, c1 }));
    }
    // Lo que una fórmula lee (para el grafo): celdas sueltas y rangos; y si es volátil
    function recolectar(ast, h, salida, profundidad) {
      if (!ast || profundidad > 200) return;
      switch (ast.t) {
        case 'ref': { const rs = rangoDe(ast, h); if (rs) salida.refs.push(...rs); return; }
        case 'nombre': {
          const def = fuente.nombre ? fuente.nombre(ast.nombre, h) : null;
          if (def && profundidad < 20) recolectar(analizar(def.replace(/^=/, '')), h, salida, profundidad + 1);
          return;
        }
        case 'fn': {
          const n = nombreFuncion(ast.nombre);
          if (VOLATILES.has(n)) salida.volatil = true;
          // FILA(A1) en A1 no es circular: estas miran dónde está la referencia, no lo que tiene
          for (const a of ast.args) if (!(SOLO_POSICION.has(n) && a && a.t === 'ref')) recolectar(a, h, salida, profundidad + 1);
          return;
        }
        case 'bin': recolectar(ast.a, h, salida, profundidad + 1); recolectar(ast.b, h, salida, profundidad + 1); return;
        case 'un': case 'paren': recolectar(ast.a, h, salida, profundidad + 1); return;
        default: return;
      }
    }

    function quitarDeGrafo(k) {
      const fx = formulas.get(k);
      if (!fx) return;
      for (const ck of fx.celdas) { const s = celdaDeps.get(ck); if (s) { s.delete(k); if (!s.size) celdaDeps.delete(ck); } }
      for (const r of fx.rangos) { r.dep = -1; rangosMuertos++; }
      formulas.delete(k);
      if (rangosMuertos > 1000 && rangosMuertos > rangoDeps.length / 2) { rangoDeps = rangoDeps.filter((r) => r.dep >= 0); rangosMuertos = 0; }
    }
    function ponerFormula(h, f, c, texto, opc) {
      const k = clave(h, f, c);
      quitarDeGrafo(k);
      const dinamica = !!(opc && opc.dinamica);
      if (derrames.has(k) && (texto === null || texto === undefined || !dinamica)) limpiarDerrame(k, null, false);
      if (texto === null || texto === undefined) return;
      const ast = analizar(texto, h, f, c);
      const fx = { h, f, c, texto, ast, celdas: [], rangos: [], volatil: false, mala: ast.t === 'mala', dinamica };
      const s = { refs: [], volatil: false };
      recolectar(ast, h, s, 0);
      fx.volatil = s.volatil;
      for (const r of s.refs) {
        if (r.f0 === r.f1 && r.c0 === r.c1) {
          const ck = clave(r.h, r.f0, r.c0);
          let set = celdaDeps.get(ck); if (!set) celdaDeps.set(ck, set = new Set());
          set.add(k); fx.celdas.push(ck);
        } else {
          const e = { h: r.h, f0: r.f0, c0: r.c0, f1: r.f1, c1: r.c1, dep: k };
          rangoDeps.push(e); fx.rangos.push(e);
        }
      }
      formulas.set(k, fx);
    }

    // opc.soloVista: lo que se calcule ahora se muestra pero no cuenta como cambio (al abrir un archivo guardado sin valores)
    function construir(opc) {
      formulas.clear(); celdaDeps.clear(); rangoDeps = []; rangosMuertos = 0; nombresHoja = null; tablasLibro = null;
      const sinValor = [];
      derrames.clear(); pendientesDerrame = [];
      for (const x of fuente.formulas()) {
        ponerFormula(x.h, x.f, x.c, x.texto, x);
        if (x.dinamica && x.ref) { const r = rangoTexto(x.ref); if (r && (r.f1 > r.f0 || r.c1 > r.c0)) derrames.set(clave(x.h, x.f, x.c), { h: x.h, f0: r.f0, c0: r.c0, f1: r.f1, c1: r.c1, ok: !esError(x.valor) }); }
        if (x.valor === null || x.valor === undefined) sinValor.push({ h: x.h, f: x.f, c: x.c });
      }
      return sinValor.length ? recalcular(sinValor, opc) : { calculadas: 0, noSoportadas: [], circulares: 0 };
    }

    // las fórmulas que leen la celda k
    function dependientesDe(k, cb) {
      const s = celdaDeps.get(k);
      if (s) for (const d of s) cb(d);
      const dr = derrames.get(k);
      if (dr && dr.ok) dependientesDeRango(dr, k, cb);   // lo que lee las celdas que desborda el ancla
      if (rangoDeps.length) {
        const { h, f, c } = deClave(k);
        for (const r of rangoDeps) if (r.dep >= 0 && r.h === h && f >= r.f0 && f <= r.f1 && c >= r.c0 && c <= r.c1) cb(r.dep);
      }
    }

    // las fórmulas que leen alguna celda del rango (salvo la celda excluida)
    function dependientesDeRango(r, excluida, cb) {
      for (let f = r.f0; f <= r.f1; f++) for (let c = r.c0; c <= r.c1; c++) {
        const ck = clave(r.h, f, c);
        if (ck === excluida) continue;
        const s = celdaDeps.get(ck);
        if (s) for (const d of s) cb(d);
      }
      for (const x of rangoDeps) if (x.dep >= 0 && x.h === r.h && x.f0 <= r.f1 && x.f1 >= r.f0 && x.c0 <= r.c1 && x.c1 >= r.c0) cb(x.dep);
    }
    // quita lo que desbordó el ancla k, salvo lo que queda dentro de `queda`; devuelve las celdas vaciadas
    function limpiarDerrame(k, queda, soloVista) {
      const d = derrames.get(k);
      derrames.delete(k);
      const vaciadas = [];
      if (!d || !d.ok || !fuente.derramar) return vaciadas;
      const ak = deClave(k);
      for (let f = d.f0; f <= d.f1; f++) for (let c = d.c0; c <= d.c1; c++) {
        if (f === ak.f && c === ak.c) continue;
        if (queda && f >= queda.f0 && f <= queda.f1 && c >= queda.c0 && c <= queda.c1) continue;
        if (formulas.has(clave(d.h, f, c)) || (fuente.ocupada && fuente.ocupada(d.h, f, c))) continue;   // lo que escribió el usuario encima no se borra
        fuente.derramar(d.h, f, c, undefined, soloVista);
        vaciadas.push({ h: d.h, f, c });
      }
      if (!queda) pendientesDerrame.push(...vaciadas);
      return vaciadas;
    }
    // el ancla de un desborde (ok o con #SPILL!) que cubre la celda, sin contar la propia ancla
    function anclasEn(h, f, c) {
      const l = [];
      for (const [k, d] of derrames) if (d.h === h && f >= d.f0 && f <= d.f1 && c >= d.c0 && c <= d.c1 && k !== clave(h, f, c)) l.push(k);
      return l;
    }

    function recalcular(cambiadas, opc, prof) {
      const todoVista = !!(opc && opc.soloVista);
      nombresHoja = null; tablasLibro = null;
      const afectadas = new Set(), cola = [];
      const visitadas = new Set();
      if (pendientesDerrame.length) { cambiadas = cambiadas.concat(pendientesDerrame); pendientesDerrame = []; }
      for (const x of cambiadas) {
        const k = clave(x.h, x.f, x.c); cola.push(k); if (formulas.has(k)) afectadas.add(k);
        // escribir en (o vaciar) una celda del desborde de otra fórmula: el ancla se recalcula (#SPILL! o vuelve a desbordar)
        if (derrames.size) for (const a of anclasEn(x.h, x.f, x.c)) if (formulas.has(a)) { afectadas.add(a); cola.push(a); }
      }
      // lo que depende de lo que cambió de verdad (sin contar las volátiles): eso se guarda; lo que cambia solo por volatilidad es de vista
      const reales = new Set(afectadas);
      { const c2 = cola.slice(), v2 = new Set(); while (c2.length) { const k = c2.pop(); if (v2.has(k)) continue; v2.add(k); dependientesDe(k, (d) => { reales.add(d); if (!v2.has(d)) c2.push(d); }); } }
      for (const [k, fx] of formulas) if (fx.volatil && !afectadas.has(k)) { afectadas.add(k); cola.push(k); }
      // aristas X → F dentro de lo afectado (para el orden)
      const salientes = new Map(), entrantes = new Map();
      while (cola.length) {
        const k = cola.pop();
        if (visitadas.has(k)) continue;
        visitadas.add(k);
        dependientesDe(k, (d) => {
          if (!afectadas.has(d)) afectadas.add(d);
          if (afectadas.has(k) && k !== d) {
            let s = salientes.get(k); if (!s) salientes.set(k, s = new Set());
            if (!s.has(d)) { s.add(d); entrantes.set(d, (entrantes.get(d) || 0) + 1); }
          } else if (k === d) entrantes.set(d, (entrantes.get(d) || 0) + 1);   // se lee a sí misma: circular
          if (!visitadas.has(d)) cola.push(d);
        });
      }
      // Kahn
      const listas = [];
      for (const k of afectadas) if (!entrantes.get(k)) listas.push(k);
      const noSoportadas = [];
      let calculadas = 0;
      const crecidas = [];   // celdas a las que un desborde llegó recién: lo que las lee no estaba en este orden
      ctxRecalculo = { crecidas, todoVista };
      while (listas.length) {
        const k = listas.pop();
        calcularFormula(k, noSoportadas, todoVista || !reales.has(k));
        calculadas++;
        const s = salientes.get(k);
        if (s) for (const d of s) { const n = entrantes.get(d) - 1; entrantes.set(d, n); if (n === 0) listas.push(d); }
        entrantes.delete(k);
      }
      // INDIRECT y OFFSET leen celdas que el grafo no conoce: pudieron calcularse antes que ellas; una segunda pasada las pone al día
      for (const k of afectadas) { const fx = formulas.get(k); if (fx && fx.volatil && !entrantes.get(k)) calcularFormula(k, [], todoVista || !reales.has(k)); }
      // lo que queda con entradas pendientes está en un ciclo: Excel muestra 0 y avisa
      let circulares = 0;
      for (const [k, n] of entrantes) if (n > 0 && formulas.has(k)) { const fx = formulas.get(k); fuente.poner(fx.h, fx.f, fx.c, 0); circulares++; }
      ctxRecalculo = null;
      const vaciadas = pendientesDerrame; pendientesDerrame = [];
      const otra = crecidas.concat(vaciadas);
      if (otra.length && (prof || 0) < 4) {
        const r = recalcular(otra, opc, (prof || 0) + 1);
        calculadas += r.calculadas; noSoportadas.push(...r.noSoportadas); circulares += r.circulares;
      }
      return { calculadas, noSoportadas, circulares };
    }

    let ctxRecalculo = null;
    let expansiones = 0;   // nombres definidos evaluados en esta fórmula
    function calcularFormula(k, noSoportadas, soloVista) {
      const fx = formulas.get(k);
      if (!fx) return;
      expansiones = 0;
      if (fx.mala) { noSoportadas.push({ h: fx.h, f: fx.f, c: fx.c, motivo: fx.ast.motivo }); return; }
      try {
        let v = evaluarNodo(fx.ast, { h: fx.h, f: fx.f, c: fx.c, prof: 0, matriz: fx.dinamica });
        if (fx.dinamica) { desbordar(k, fx, v, soloVista); return; }
        v = resultadoDeCelda(v, fx);
        fuente.poner(fx.h, fx.f, fx.c, v, soloVista);
      } catch (e) {
        if (e instanceof Lanzado) { if (fx.dinamica) limpiarDerrame(k, null, soloVista); fuente.poner(fx.h, fx.f, fx.c, e.e, soloVista); return; }
        if (e instanceof NoSoportada) { noSoportadas.push({ h: fx.h, f: fx.f, c: fx.c, motivo: e.message }); return; }
        if (e instanceof RangeError) { fuente.poner(fx.h, fx.f, fx.c, E.num); return; }   // anidado demasiado hondo
        throw e;
      }
    }
    // matriz dinámica: el resultado cubre un rango desde el ancla; si cabe, va a las celdas vecinas; si no, #SPILL!
    function desbordar(k, fx, v, soloVista) {
      let M = null;
      if (esRef(v) && !(v.f0 === v.f1 && v.c0 === v.c1)) M = matrizDe(v);
      else if (esMat(v)) M = v.filas;
      const viejo = derrames.get(k);
      if (!M || (M.length === 1 && M[0].length === 1)) {
        if (viejo) { if (fuente.antesDeDerrame) fuente.antesDeDerrame(fx.h, fx.f, fx.c, soloVista); limpiarDerrame(k, null, soloVista); }
        fuente.poner(fx.h, fx.f, fx.c, resultadoDeCelda(M ? M[0][0] : v, fx), soloVista);
        return;
      }
      const r = { h: fx.h, f0: fx.f, c0: fx.c, f1: fx.f + M.length - 1, c1: fx.c + M[0].length - 1 };
      let libre = r.f1 < MAX_FILAS && r.c1 < MAX_COLS && (r.f1 - r.f0 + 1) * (r.c1 - r.c0 + 1) <= 1000000;
      if (libre) {
        // ocupada: un valor o fórmula que no es de un desborde; o una celda del desborde de otra ancla
        for (let f = r.f0; f <= r.f1 && libre; f++) for (let c = r.c0; c <= r.c1; c++) {
          if (f === fx.f && c === fx.c) continue;
          if (formulas.has(clave(fx.h, f, c)) || (fuente.ocupada && fuente.ocupada(fx.h, f, c))) { libre = false; break; }
        }
        if (libre) for (const [a, d] of derrames) if (a !== k && d.ok && d.h === r.h && d.f0 <= r.f1 && d.f1 >= r.f0 && d.c0 <= r.c1 && d.c1 >= r.c0) { libre = false; break; }
      }
      const igual = viejo && viejo.ok === libre && viejo.f1 === r.f1 && viejo.c1 === r.c1;
      if (!igual && fuente.antesDeDerrame) fuente.antesDeDerrame(fx.h, fx.f, fx.c, soloVista);
      if (!libre) {
        if (viejo) limpiarDerrame(k, null, soloVista);
        derrames.set(k, Object.assign(r, { ok: false }));
        fuente.poner(fx.h, fx.f, fx.c, error('#SPILL!'), soloVista);
        return;
      }
      if (viejo) limpiarDerrame(k, r, soloVista);
      derrames.set(k, Object.assign(r, { ok: true }));
      const valorDe = (x) => (x === null || x === undefined ? 0 : typeof x === 'number' && !Number.isFinite(x) ? E.num : esRef(x) || esMat(x) ? E.valor : x);
      for (let i = 0; i < M.length; i++) for (let j = 0; j < M[0].length; j++) {
        const f = r.f0 + i, c = r.c0 + j, x = valorDe(M[i][j]);
        if (i === 0 && j === 0) { fuente.poner(fx.h, f, c, x, soloVista); continue; }
        fuente.derramar(fx.h, f, c, x, soloVista);
        if (ctxRecalculo && (!viejo || !viejo.ok || f > viejo.f1 || c > viejo.c1)) ctxRecalculo.crecidas.push({ h: fx.h, f, c });
      }
    }
    // lo que queda guardado en la celda: un escalar (una matriz o un rango se reducen como en Excel sin matrices dinámicas)
    function resultadoDeCelda(v, ctx) {
      if (esRef(v)) v = interseccion(v, ctx);
      else if (esMat(v)) v = v.filas[0][0];
      else if (esUnion(v)) v = E.valor;
      if (v === null || v === undefined) return 0;   // =A1 con A1 vacía muestra 0
      if (typeof v === 'number' && !Number.isFinite(v)) return E.num;
      return v;
    }

    // ---- evaluación ---------------------------------------------------------------------------------------------------
    function valorCelda(h, f, c) {
      const v = fuente.valor(h, f, c);
      return v === undefined ? null : v;
    }
    // un rango usado como un solo valor: la celda de su misma fila (o columna); si no, #VALUE! (intersección implícita)
    function interseccion(r, ctx) {
      if (r.f0 === r.f1 && r.c0 === r.c1) return valorCelda(r.h, r.f0, r.c0);
      if (r.c0 === r.c1 && ctx.f >= r.f0 && ctx.f <= r.f1) return valorCelda(r.h, ctx.f, r.c0);
      if (r.f0 === r.f1 && ctx.c >= r.c0 && ctx.c <= r.c1) return valorCelda(r.h, r.f0, ctx.c);
      return E.valor;
    }
    function escalar(v, ctx) {
      if (esRef(v)) return interseccion(v, ctx);
      if (esMat(v)) return v.filas[0][0];
      if (esUnion(v)) return E.valor;
      return v;
    }
    const recortar = (r) => {
      const lim = fuente.limite ? fuente.limite(r.h) : { filas: MAX_FILAS, cols: MAX_COLS };
      return { f1: Math.min(r.f1, lim.filas - 1), c1: Math.min(r.c1, lim.cols - 1) };
    };
    // matriz de valores de un rango: con la forma escrita (=LARGO(A1:A5) da 5 filas aunque lo usado llegue a la 3); uno enorme (A:A), solo hasta lo usado
    function matrizDe(v) {
      if (esMat(v)) return v.filas;
      if (esRef(v)) {
        const { f1, c1 } = recortar(v);
        const entero = (v.f1 - v.f0 + 1) * (v.c1 - v.c0 + 1) <= FORMA_ENTERA;
        const fFin = entero ? v.f1 : f1, cFin = entero ? v.c1 : c1;
        if (fFin < v.f0 || cFin < v.c0) return [[null]];
        if ((fFin - v.f0 + 1) * (cFin - v.c0 + 1) > 5000000) throw new NoSoportada('rango demasiado grande');
        const filas = [];
        for (let f = v.f0; f <= fFin; f++) {
          const fila = [];
          for (let c = v.c0; c <= cFin; c++) fila.push(f <= f1 && c <= c1 ? valorCelda(v.h, f, c) : null);
          filas.push(fila);
        }
        return filas;
      }
      if (esUnion(v)) throw new Lanzado(E.valor);
      return [[v]];
    }
    // recorre los valores de los argumentos: cb(valor, deRango). Los de un rango o matriz llegan con deRango = true (SUM no cuenta textos de un rango)
    function recorrer(args, ctx, cb) {
      for (const a of args) {
        if (a === null) { cb(null, false); continue; }
        const v = evaluarNodo(a, ctx);
        recorrerValor(v, cb);
      }
    }
    function recorrerValor(v, cb) {
      if (esUnion(v)) { for (const r of v.refs) recorrerValor(r, cb); return; }
      if (esRef(v)) {
        const { f1, c1 } = recortar(v);
        for (let f = v.f0; f <= f1; f++) for (let c = v.c0; c <= c1; c++) cb(valorCelda(v.h, f, c), true);
        return;
      }
      if (esMat(v)) { for (const fila of v.filas) for (const x of fila) cb(x, true); return; }
      cb(v, false);
    }

    function aNumero(v) {
      if (typeof v === 'number') return v;
      if (v === null) return 0;
      if (typeof v === 'boolean') return v ? 1 : 0;
      if (esError(v)) lanzar(v);
      if (typeof v === 'string') { const n = textoANumero(v); if (n === null) lanzar(E.valor); return n; }
      lanzar(E.valor);
    }
    function aTexto(v) {
      if (typeof v === 'string') return v;
      if (v === null) return '';
      if (typeof v === 'number') { const t = formatoGeneral(v); return fuente.idioma === 'es' ? t.replace('.', ',') : t; }
      if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
      if (esError(v)) lanzar(v);
      return String(v);
    }
    function aLogico(v) {
      if (typeof v === 'boolean') return v;
      if (typeof v === 'number') return v !== 0;
      if (v === null) return false;
      if (esError(v)) lanzar(v);
      if (typeof v === 'string') { const u = v.toUpperCase(); if (u === 'TRUE' || u === 'VERDADERO') return true; if (u === 'FALSE' || u === 'FALSO') return false; }
      lanzar(E.valor);
    }
    const num = (a, ctx) => aNumero(escalar(evaluarNodo(a, ctx), ctx));
    const txt = (a, ctx) => aTexto(escalar(evaluarNodo(a, ctx), ctx));
    const log = (a, ctx) => aLogico(escalar(evaluarNodo(a, ctx), ctx));
    const val = (a, ctx) => { const v = escalar(evaluarNodo(a, ctx), ctx); if (esError(v)) lanzar(v); return v; };
    const opcional = (a, ctx, porDefecto, conv) => (a === undefined || a === null ? porDefecto : conv(a, ctx));

    // comparación de Excel: número < texto < lógico; textos sin distinguir mayúsculas; vacía = 0 o "" según el otro
    function comparar(a, b) {
      if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
      if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
      const ta = typeof a === 'number' ? 0 : typeof a === 'string' ? 1 : 2, tb = typeof b === 'number' ? 0 : typeof b === 'string' ? 1 : 2;
      if (ta !== tb) return ta - tb;
      if (ta === 1) { const x = a.toUpperCase(), y = b.toUpperCase(); return x < y ? -1 : x > y ? 1 : 0; }
      if (ta === 2) return (a ? 1 : 0) - (b ? 1 : 0);
      return a < b ? -1 : a > b ? 1 : 0;
    }

    // operación elemento a elemento (para matrices: SUMPRODUCT((A1:A3>2)*B1:B3))
    function binaria(op, x, y, ctx) {
      const mx = esRef(x) && !(x.f0 === x.f1 && x.c0 === x.c1) ? matrizDe(x) : esMat(x) ? x.filas : null;
      const my = esRef(y) && !(y.f0 === y.f1 && y.c0 === y.c1) ? matrizDe(y) : esMat(y) ? y.filas : null;
      if ((mx || my) && (ctx.matriz || esMat(x) || esMat(y))) {
        const A = mx || [[escalar(x, ctx)]], B = my || [[escalar(y, ctx)]];
        const nf = Math.max(A.length, B.length), nc = Math.max(A[0].length, B[0].length), filas = [];
        if (nf * nc > 5000000) return E.valor;   // una columna por una fila enormes (FILA(A1:A999999)*COLUMNA(1:1)) agotarían la memoria
        for (let i = 0; i < nf; i++) {
          const fila = [];
          for (let j = 0; j < nc; j++) {
            const a = (A[A.length === 1 ? 0 : i] || [])[A[0].length === 1 ? 0 : j], b = (B[B.length === 1 ? 0 : i] || [])[B[0].length === 1 ? 0 : j];
            fila.push(a === undefined || b === undefined ? E.na : operar(op, a, b));
          }
          filas.push(fila);
        }
        return { tipo: 'mat', filas };
      }
      return operar(op, escalar(x, ctx), escalar(y, ctx));
    }
    function operar(op, a, b) {
      try {
        if (esError(a)) return a;
        if (esError(b)) return b;
        switch (op) {
          case '+': return aNumero(a) + aNumero(b);
          case '-': return aNumero(a) - aNumero(b);
          case '*': return aNumero(a) * aNumero(b);
          case '/': { const d = aNumero(b); const n = aNumero(a); if (d === 0) return E.div; return n / d; }
          case '^': { const r = Math.pow(aNumero(a), aNumero(b)); return Number.isFinite(r) ? r : E.num; }
          case '&': return aTexto(a) + aTexto(b);
          case '=': return comparar(a, b) === 0;
          case '<>': return comparar(a, b) !== 0;
          case '<': return comparar(a, b) < 0;
          case '<=': return comparar(a, b) <= 0;
          case '>': return comparar(a, b) > 0;
          case '>=': return comparar(a, b) >= 0;
          default: throw new NoSoportada('operador ' + op);
        }
      } catch (e) { if (e instanceof Lanzado) return e.e; throw e; }
    }

    // [n]Hoja!A1:B2 → los valores guardados del otro libro: una celda o una matriz; #REF! si el vínculo o la hoja no están guardados
    function valorExterno(n) {
      const m = /^\[(\d+)\](.*)$/.exec(n.hoja.a);
      if (n.hoja.b !== n.hoja.a) throw new NoSoportada('referencia 3D a otro libro');
      const datos = fuente.externo ? fuente.externo(+m[1], m[2]) : null;
      if (!datos) return E.ref;
      let f0 = n.f0, c0 = n.c0, f1 = n.f0, c1 = n.c0;
      if (n.modo === 'area') { f0 = Math.min(n.f0, n.f1); f1 = Math.max(n.f0, n.f1); c0 = Math.min(n.c0, n.c1); c1 = Math.max(n.c0, n.c1); }
      else if (n.modo !== 'celda') throw new NoSoportada('filas o columnas enteras de otro libro');
      const leer = (f, c) => { const v = datos.celdas.get(f * MAX_COLS + c); return v === undefined ? (datos.errorAlActualizar ? E.ref : null) : v; };
      if (f0 === f1 && c0 === c1) { const v = leer(f0, c0); return v === null ? 0 : v; }
      if ((f1 - f0 + 1) * (c1 - c0 + 1) > 100000) throw new NoSoportada('rango grande de otro libro');
      const filas = [];
      for (let f = f0; f <= f1; f++) { const fila = []; for (let c = c0; c <= c1; c++) fila.push(leer(f, c)); filas.push(fila); }
      return { tipo: 'mat', filas };
    }

    function evaluarNodo(n, ctx) {
      if (++ctx.prof > 400) { ctx.prof = 0; throw new RangeError('profundidad'); }
      try {
        switch (n.t) {
          case 'num': return n.v;
          case 'valor': return n.v;
          case 'str': return n.v;
          case 'bool': return n.v;
          case 'err': return error(n.v);
          case 'ref': {
            if (n.hoja && /^\[\d+\]/.test(n.hoja.a)) return valorExterno(n);
            const rs = rangoDe(n, ctx.h);
            if (!rs) return E.ref;
            return rs.length === 1 ? rs[0] : { tipo: 'union', refs: rs };
          }
          case 'paren': return evaluarNodo(n.a, ctx);
          case 'arr': return { tipo: 'mat', filas: n.filas.map((f) => f.map((x) => evaluarNodo(x, ctx))) };
          case 'nombre': {
            const local = String(n.nombre).toUpperCase().replace(/^_XLPM\./, '');
            if (ctx.ambito && ctx.ambito.has(local)) return ctx.ambito.get(local);
            const def = fuente.nombre ? fuente.nombre(n.nombre, ctx.h) : null;
            if (!def) return E.nombre;
            if (++expansiones > MAX_EXPANSIONES) throw new RangeError('nombres');   // N1=N2+N2, N2=N3+N3…: 34 nombres son 2^33 pasos
            return evaluarNodo(analizar(def.replace(/^=/, ''), ctx.h, ctx.f, ctx.c), ctx);
          }
          case 'un': {
            const v = evaluarNodo(n.a, ctx);
            if (n.op === '%') return binaria('/', v, 100, ctx);
            if (n.op === '-') return binaria('*', v, -1, ctx);
            return ctx.matriz ? v : escalar(v, ctx);
          }
          case 'bin': {
            if (n.op === ':') {   // A1:B2 armado con dos referencias (o INDEX(...):B5)
              const a = evaluarNodo(n.a, ctx), b = evaluarNodo(n.b, ctx);
              if (esError(a)) return a;
              if (esError(b)) return b;
              if (!esRef(a) || !esRef(b) || a.h !== b.h) return E.valor;
              return { tipo: 'ref', h: a.h, f0: Math.min(a.f0, b.f0), c0: Math.min(a.c0, b.c0), f1: Math.max(a.f1, b.f1), c1: Math.max(a.c1, b.c1) };
            }
            if (n.op === ',') { const a = evaluarNodo(n.a, ctx), b = evaluarNodo(n.b, ctx); const refs = []; for (const x of [a, b]) { if (esUnion(x)) refs.push(...x.refs); else if (esRef(x)) refs.push(x); else return E.valor; } return { tipo: 'union', refs }; }
            if (n.op === ' ') {   // intersección
              const a = evaluarNodo(n.a, ctx), b = evaluarNodo(n.b, ctx);
              if (!esRef(a) || !esRef(b) || a.h !== b.h) return E.valor;
              const f0 = Math.max(a.f0, b.f0), f1 = Math.min(a.f1, b.f1), c0 = Math.max(a.c0, b.c0), c1 = Math.min(a.c1, b.c1);
              if (f0 > f1 || c0 > c1) return E.nulo;
              return { tipo: 'ref', h: a.h, f0, c0, f1, c1 };
            }
            return binaria(n.op, evaluarNodo(n.a, ctx), evaluarNodo(n.b, ctx), ctx);
          }
          case 'fn': return llamar(n, ctx);
          default: throw new NoSoportada('nodo ' + n.t);
        }
      } finally { ctx.prof--; }
    }

    function llamar(n, ctx) {
      const nombre = nombreFuncion(n.nombre);
      if (nombre === 'LET') {
        if (n.args.length < 3 || n.args.length % 2 === 0) return E.valor;
        const ambito = new Map(ctx.ambito || []);
        const sub = Object.assign({}, ctx, { ambito });
        for (let k = 0; k + 1 < n.args.length; k += 2) {
          const nodo = n.args[k];
          if (!nodo || nodo.t !== 'nombre') return E.valor;
          const v = evaluarNodo(n.args[k + 1], sub);
          ambito.set(String(nodo.nombre).toUpperCase().replace(/^_XLPM\./, ''), v);
        }
        const r = evaluarNodo(n.args[n.args.length - 1], sub);
        ctx.prof = sub.prof;
        return r;
      }
      const fn = FUNCIONES[nombre];
      if (!fn) {
        if (CONOCIDAS.has(nombre)) throw new NoSoportada(nombre);
        return E.nombre;
      }
      try {
        const patron = ELEVA.get(nombre);
        if (patron !== undefined && puedeSerMatriz(n)) return elevar(fn, n, ctx, patron);
        return fn(n.args, ctx, API);
      } catch (e) {
        if (e instanceof Lanzado) return e.e;
        if (e instanceof NoSoportada || e instanceof RangeError) throw e;
        throw new NoSoportada(nombre + ': ' + (e && e.message));
      }
    }

    // ¿alguno de los argumentos puede dar más de un valor? Lo de todos los días (A1, 2, "x") no paga la evaluación por adelantado. Se guarda en el nodo.
    function puedeSerMatriz(n) {
      if (n.quizaMatriz === undefined) n.quizaMatriz = n.args.some((a) => a && !(a.t === 'num' || a.t === 'str' || a.t === 'bool' || (a.t === 'ref' && a.modo === 'celda')));
      return n.quizaMatriz;
    }
    const varias = (v) => esRef(v) && !(v.f0 === v.f1 && v.c0 === v.c1);
    // Una función de valores sueltos que recibe una matriz: se aplica a cada elemento (Excel 365). Lo que se puede aplicar por posición está en ELEVA.
    // Sin matrices dinámicas (ctx.matriz falso) un rango se reduce a un valor, como en las fórmulas de archivos sin cm; una matriz constante se aplica igual.
    function elevar(fn, n, ctx, patron) {
      const args = n.args.slice(), matrices = [];
      for (let k = 0; k < args.length; k++) {
        const tipo = posicionDe(patron, k);
        if (tipo === 'a' || tipo === 'z' || tipo === 'v' || !args[k]) continue;
        const v = evaluarNodo(args[k], ctx);
        if (tipo === 'x' && ctx.matriz && varias(v)) return E.valor;
        args[k] = { t: 'valor', v };
        if (esMat(v) || ((tipo === 's' || tipo === 'e') && ctx.matriz && varias(v))) matrices.push({ k, M: matrizDe(v), e: tipo === 'e' });
      }
      if (!matrices.length) return fn(args, ctx, API);   // como la función la espera: IF(TRUE;A1:A3) devuelve el rango entero
      for (let k = 0; k < args.length; k++) {
        const nodo = args[k];
        if (!nodo || nodo.t === 'valor') continue;
        const tipo = posicionDe(patron, k);
        // las ramas de IF y compañía: se toma el elemento de la misma posición
        if (tipo === 'v') {
          const v = evaluarNodo(nodo, ctx);
          if (esMat(v) || varias(v)) matrices.push({ k, M: matrizDe(v) });   // también sin matrices dinámicas: VLOOKUP(1;IF({1\0};A1:A9;B1:B9);2;FALSE)
          args[k] = { t: 'valor', v };
        } else if (nodo.t === 'fn' || nodo.t === 'bin' || nodo.t === 'un' || nodo.t === 'paren') args[k] = { t: 'valor', v: evaluarNodo(nodo, ctx) };   // no se recalcula por cada elemento
      }
      let nf = 1, nc = 1;
      for (const { M } of matrices) { nf = Math.max(nf, M.length); nc = Math.max(nc, M[0].length); }
      if (nf * nc > 1000000) return E.valor;
      const hayE = matrices.some((m) => m.e), conRamas = patron.indexOf('v') >= 0, filas = [];
      for (let i = 0; i < nf; i++) {
        const fila = [];
        for (let j = 0; j < nc; j++) {
          let falta = false;
          for (const { k, M } of matrices) {
            const x = (M[M.length === 1 ? 0 : i] || [])[M[0].length === 1 ? 0 : j];
            if (x === undefined) { falta = true; break; }
            args[k] = { t: 'valor', v: x };
          }
          if (falta) { fila.push(E.na); continue; }
          let r;
          try { r = fn(args, ctx, API); } catch (e) { if (e instanceof Lanzado) r = e.e; else throw e; }
          r = hayE && (esMat(r) || esRef(r)) ? E.valor : esMat(r) ? r.filas[0][0] : escalar(r, ctx);
          fila.push(r === null && conRamas ? 0 : r);   // medido: en una matriz, la celda vacía que elige IF ya no es vacía (MIN da 0, COUNT la cuenta, LEN da 1)
        }
        filas.push(fila);
      }
      return { tipo: 'mat', filas };
    }
    // ¿Excel guarda esta fórmula como matriz dinámica (cm)? Sí si en algún punto se usa un rango o una matriz donde va un valor suelto, o si llama a una función que
    // devuelve matrices. Es por lo escrito, no por lo que dé: LEN(A1:A1) y SEQUENCE(1) dan un solo valor y también se guardan así. Si no se marcara, al abrirla sería otra fórmula.
    function necesitaMatriz(ast, h) {
      const nombres = new Map(), vistos = new Map();
      // cada definición, una vez por tipo: un nombre que se usa a sí mismo (LET(x;x+x;x)) o una cadena N1=N2+N2… serían exponenciales. Si una ya dio sí, no se llega a repetirla
      const unaVez = (nodo, tipo, prof) => {
        let t = vistos.get(nodo);
        if (!t) vistos.set(nodo, t = new Set());
        if (t.has(tipo)) return false;
        t.add(tipo);
        return visitar(nodo, tipo, prof + 1);
      };
      const visitar = (n, tipo, prof) => {
        if (!n || prof > 200) return false;
        switch (n.t) {
          case 'ref': return n.modo !== 'celda' && tipo !== 'a' && tipo !== 'x' && tipo !== 'f';
          case 'arr': return tipo !== 'a' && tipo !== 'z';
          case 'paren': return visitar(n.a, tipo, prof + 1);
          case 'un': return visitar(n.a, 's', prof + 1);
          case 'bin':
            if (n.op === ':' || n.op === ' ' || n.op === ',') return tipo === 'r' || tipo === 's';
            return visitar(n.a, 's', prof + 1) || visitar(n.b, 's', prof + 1);
          case 'nombre': {
            const local = String(n.nombre).toUpperCase().replace(/^_XLPM\./, '');
            if (nombres.has(local)) return unaVez(nombres.get(local), tipo, prof);
            const def = fuente.nombre ? fuente.nombre(n.nombre, h) : null;
            return def ? unaVez(analizar(def.replace(/^=/, '')), tipo, prof) : false;
          }
          case 'fn': {
            const nombre = nombreFuncion(n.nombre);
            if (DEVUELVEN_MATRIZ.has(nombre) || SIEMPRE_DINAMICA.has(nombre)) return true;
            if (DENTRO_SIN_CM.has(nombre) && (tipo === 'r' || tipo === 's')) return true;
            if ((nombre === 'INDIRECT' && (tipo === 'r' || tipo === 's' || tipo === 'z')) || (nombre === 'OFFSET' && tipo === 's')) return true;
            if (nombre === 'LET') {
              for (let k = 0; k + 1 < n.args.length; k += 2) {
                if (visitar(n.args[k + 1], 'a', prof + 1)) return true;
                if (n.args[k] && n.args[k].t === 'nombre') nombres.set(String(n.args[k].nombre).toUpperCase().replace(/^_XLPM\./, ''), n.args[k + 1]);
              }
              return visitar(n.args[n.args.length - 1], tipo, prof + 1);
            }
            const patron = ELEVA.get(nombre);
            return n.args.some((a, k) => {
              const p = patron === undefined ? 'a' : posicionDe(patron, k);
              return visitar(a, p === 'v' ? tipo : p, prof + 1);   // una rama vale lo que valga la fórmula: IF(A1;B1;C1) no, IF(A1;B1:B3) sí
            });
          }
          default: return false;
        }
      };
      return visitar(ast, 'r', 0);
    }

    const API = { num, txt, log, val, opcional, evaluarNodo, escalar, matrizDe, recorrer, recorrerValor, aNumero, aTexto, aLogico, comparar, valorCelda, esRef, esMat, esUnion, recortar,
      hoy, fuente, E, error, lanzar, NoSoportada, interseccion, serialDe, partes, formatoGeneral, textoANumero, analizar, rangoDe, indiceHoja,
      derrame: (h, f, c) => derrames.get(clave(h, f, c)),
      esMatrizDeUnValor: (h, f, c) => { const fx = formulas.get(clave(h, f, c)); return !!fx && !!fx.dinamica && !fx.mala && necesitaMatriz(fx.ast, h); } };

    function evaluar(texto, ctx) {
      const x = Object.assign({ h: 0, f: 0, c: 0 }, ctx);
      const ast = analizar(String(texto).replace(/^=/, ''), x.h, x.f, x.c);
      if (ast.t === 'mala') return E.nombre;
      expansiones = 0;
      try { return resultadoDeCelda(evaluarNodo(ast, Object.assign({ h: 0, f: 0, c: 0, prof: 0 }, ctx)), ctx || { f: 0, c: 0 }); } catch (e) {
        if (e instanceof Lanzado) return e.e;
        if (e instanceof NoSoportada) return null;
        if (e instanceof RangeError) return E.num;   // anidado demasiado hondo (un nombre que se usa a sí mismo), igual que en una celda
        throw e;
      }
    }

    // las funciones de fecha leen fecha1904 del módulo: cada entrada del motor la pone con la de su libro (con dos libros abiertos, cada uno la suya)
    const conSuFecha = (fn) => function () { fecha1904 = !!fuente.fecha1904; return fn.apply(this, arguments); };
    // todas las fórmulas, como si todo hubiera cambiado (después de insertar o quitar filas, los valores guardados en el archivo ya no valen)
    const recalcularTodo = () => recalcular([...formulas.values()].map((fx) => ({ h: fx.h, f: fx.f, c: fx.c })));
    return { necesitaMatriz: (h, f, c) => { const fx = formulas.get(clave(h, f, c)); return !!fx && !fx.mala && necesitaMatriz(fx.ast, h); }, construir: conSuFecha(construir), ponerFormula: conSuFecha(ponerFormula), recalcular: conSuFecha(recalcular), recalcularTodo: conSuFecha(recalcularTodo), evaluar: conSuFecha(evaluar), get cantidad() { return formulas.size; }, tieneFormula: (h, f, c) => formulas.has(clave(h, f, c)),
      derrame: (h, f, c) => derrames.get(clave(h, f, c)), get derrames() { return derrames; } };
  }

  // ---- funciones ----------------------------------------------------------------------------------------------------
  const nombreFuncion = (s) => String(s).toUpperCase().replace(/^_XLFN\._XLWS\.|^_XLFN\.|^_XLWS\./, '');
  const VOLATILES = new Set(['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'RANDARRAY', 'OFFSET', 'INDIRECT', 'CELL', 'INFO']);
  const SOLO_POSICION = new Set(['ROW', 'COLUMN', 'ROWS', 'COLUMNS']);
  const CONOCIDAS = new Set((typeof XlsFormulas !== 'undefined' ? XlsFormulas.FUNCIONES : []).map((f) => String(f.nombre).toUpperCase().replace(/^_XLFN\./, '')));

  const redondear = (x, d) => { const p = Math.pow(10, d); const y = Math.abs(x) * p; const r = Math.round(Number(y.toPrecision(15))); return Math.sign(x) * r / p; };
  const truncar = (x, d) => { const p = Math.pow(10, d); return Math.trunc(Number((x * p).toPrecision(15))) / p; };
  const arriba = (x, d) => { const p = Math.pow(10, d); return Math.sign(x) * Math.ceil(Number((Math.abs(x) * p).toPrecision(15))) / p; };

  // reúne los números de los argumentos (SUM, AVERAGE…): de un rango solo los números; sueltos, también textos numéricos y lógicos
  function numeros(args, ctx, A, opc) {
    const lista = [];
    A.recorrer(args, ctx, (v, deRango) => {
      if (A.esRef(v)) return;
      if (typeof v === 'number') lista.push(v);
      else if (esError(v)) A.lanzar(v);
      else if (deRango) { if (opc && opc.conTodo) { if (typeof v === 'boolean') lista.push(v ? 1 : 0); else if (typeof v === 'string') lista.push(0); } }
      else if (v !== null) lista.push(A.aNumero(v));
      else if (opc && opc.vaciaComoCero) lista.push(0);
    });
    return lista;
  }
  const sumar = (l) => { let s = 0; for (const x of l) s += x; return s; };

  // valores de los argumentos de SUBTOTAL/AGGREGATE; de los rangos salta lo que pida el filtro: celdas con SUBTOTAL/AGGREGATE (anidadas),
  // filas ocultas y errores. Devuelve la lista cruda (números, textos, lógicos, errores; null = vacía)
  const RE_ANIDADA = /\b(?:_xlfn\.)?(?:SUBTOTAL|AGGREGATE)\s*\(/i;
  function valoresFiltrados(args, ctx, A, filtro) {
    const lista = [];
    const poner = (v, deRango) => { if (filtro.errores && esError(v)) return; lista.push(deRango ? { v } : { v, suelto: true }); };
    for (const arg of args) {
      if (arg === null) continue;
      const v = A.evaluarNodo(arg, ctx);
      const refs = A.esUnion(v) ? v.refs : A.esRef(v) ? [v] : null;
      if (!refs) { A.recorrerValor(v, poner); continue; }
      for (const r of refs) {
        const { f1, c1 } = A.recortar(r);
        for (let f = r.f0; f <= f1; f++) {
          if (filtro.ocultas && A.fuente.filaOculta && A.fuente.filaOculta(r.h, f)) continue;
          for (let c = r.c0; c <= c1; c++) {
            if (filtro.anidadas && A.fuente.formulaDe) { const fx = A.fuente.formulaDe(r.h, f, c); if (fx && RE_ANIDADA.test(fx)) continue; }
            poner(A.valorCelda(r.h, f, c), true);
          }
        }
      }
    }
    return lista;
  }
  // función 1–13 de AGGREGATE (1–11 de SUBTOTAL) sobre los valores ya filtrados
  function agregarLista(fn, lista, A) {
    if (fn === 3) return lista.filter((x) => x.v !== null).length;
    const l = [];
    for (const x of lista) {
      if (esError(x.v)) return x.v;
      if (typeof x.v === 'number') l.push(x.v);
      else if (x.suelto && x.v !== null) { const n = A.aNumero(x.v); l.push(n); }
    }
    switch (fn) {
      case 1: return l.length ? sumar(l) / l.length : A.E.div;
      case 2: return l.length;
      case 4: return l.length ? Math.max(...l) : 0;
      case 5: return l.length ? Math.min(...l) : 0;
      case 6: return l.reduce((p, x) => p * x, 1);
      case 7: return desviacion(l, 1, A);
      case 8: return desviacion(l, 0, A);
      case 9: return sumar(l);
      case 10: return varianza(l, 1, A);
      case 11: return varianza(l, 0, A);
      case 12: { if (!l.length) return A.E.num; l.sort((x, y) => x - y); const m = l.length >> 1; return l.length % 2 ? l[m] : (l[m - 1] + l[m]) / 2; }
      case 13: { const cuenta = new Map(); let mejor = null, n = 1; for (const x of l) { const k = (cuenta.get(x) || 0) + 1; cuenta.set(x, k); if (k > n) { n = k; mejor = x; } } return mejor === null ? A.E.na : mejor; }
      default: return A.E.valor;
    }
  }
  // percentil inclusivo (PERCENTILE.INC) o exclusivo (PERCENTILE.EXC), con la interpolación de Excel
  function percentil(l, k, inclusivo, A) {
    const n = l.length;
    if (!n || k < 0 || k > 1) return A.E.num;
    const x = l.slice().sort((p, q) => p - q);
    let pos;
    if (inclusivo) pos = k * (n - 1);
    else { pos = k * (n + 1) - 1; if (pos < 0 || pos > n - 1) return A.E.num; }
    const i = Math.floor(pos), d = pos - i;
    return i + 1 < n ? x[i] + d * (x[i + 1] - x[i]) : x[i];
  }

  // criterio de COUNTIF / SUMIF: ">5", "<>x", "ab*", 5, "" …
  function criterio(c, A) {
    if (esError(c)) A.lanzar(c);
    if (typeof c === 'number' || typeof c === 'boolean') return (v) => v === c || (typeof c === 'number' && typeof v === 'string' && A.textoANumero(v) === c);
    if (c === null) return (v) => v === 0;
    let s = String(c), op = '=';
    const m = /^(<=|>=|<>|<|>|=)/.exec(s);
    if (m) { op = m[1]; s = s.slice(op.length); }
    const n = s.trim() !== '' ? A.textoANumero(s) : null;
    if (n !== null) {
      return (v) => {
        // "=5" también cuenta el texto "5"; las desigualdades (">3") solo comparan números de verdad, como Excel
        const x = typeof v === 'number' ? v : (typeof v === 'string' && (op === '=' || op === '<>') ? A.textoANumero(v) : null);
        if (op === '<>') return x === null || x !== n;
        if (x === null || typeof v === 'boolean') return false;
        switch (op) { case '=': return x === n; case '<': return x < n; case '<=': return x <= n; case '>': return x > n; case '>=': return x >= n; default: return false; }
      };
    }
    if (s === '') {
      if (op === '=') return (v) => v === null || v === '';
      if (op === '<>') return (v) => v !== null && v !== '';
    }
    const u = s.toUpperCase();
    if (/^(TRUE|FALSE)$/.test(u) && (op === '=' || op === '<>')) { const b = u === 'TRUE'; return (v) => (op === '=' ? v === b : v !== b); }
    const comodin = /[*?]/.test(s) && (op === '=' || op === '<>');
    let re = null;
    if (comodin) re = new RegExp('^' + s.replace(/~([*?~])|([*?])|([.+^${}()|[\]\\])/g, (t, esc, w, meta) => (esc ? '\\' + esc : w === '*' ? '[\\s\\S]*' : w === '?' ? '[\\s\\S]' : '\\' + meta)) + '$', 'i');
    const literal = s.replace(/~([*?~])/g, '$1').toUpperCase();
    return (v) => {
      if (typeof v !== 'string') return op === '<>' ? true : false;
      if (re) return op === '=' ? re.test(v) : !re.test(v);
      const x = v.toUpperCase();
      switch (op) { case '=': return x === literal; case '<>': return x !== literal; case '<': return x < literal; case '<=': return x <= literal; case '>': return x > literal; case '>=': return x >= literal; default: return false; }
    };
  }
  // pares (rango, criterio) de COUNTIFS/SUMIFS; devuelve las posiciones [i, j] que cumplen todo
  function filtrarPares(args, desde, ctx, A) {
    let forma = null, ok = null;
    for (let k = desde; k + 1 < args.length + 1 && k < args.length; k += 2) {
      const M = A.matrizDe(A.evaluarNodo(args[k], ctx));
      const cr = criterio(A.escalar(A.evaluarNodo(args[k + 1], ctx), ctx), A);
      if (!forma) { forma = [M.length, M[0].length]; ok = M.map((f) => f.map(() => true)); }
      else if (M.length !== forma[0] || M[0].length !== forma[1]) A.lanzar(A.E.valor);
      for (let i = 0; i < M.length; i++) for (let j = 0; j < M[i].length; j++) if (ok[i][j] && !cr(M[i][j])) ok[i][j] = false;
    }
    return ok;
  }
  // el rango de valores de SUMIF(rango, criterio, [rango_suma]): del mismo tamaño que el rango, desde la esquina del rango_suma
  function rangoSuma(nodo, base, ctx, A) {
    const v = A.evaluarNodo(nodo, ctx);
    if (A.esRef(v) && A.esRef(base)) return A.matrizDe({ tipo: 'ref', h: v.h, f0: v.f0, c0: v.c0, f1: v.f0 + (base.f1 - base.f0), c1: v.c0 + (base.c1 - base.c0) });
    return A.matrizDe(v);
  }

  // SUMIFS / AVERAGEIFS / MAXIFS / MINIFS: los números del primer rango donde se cumplen todos los pares (rango, criterio)
  function conjunto(args, ctx, A, fn) {
    const S = A.matrizDe(A.evaluarNodo(args[0], ctx)), ok = filtrarPares(args, 1, ctx, A);
    if (!ok || ok.length !== S.length || ok[0].length !== S[0].length) return A.E.valor;
    const l = [];
    for (let i = 0; i < S.length; i++) for (let j = 0; j < S[i].length; j++) if (ok[i][j] && typeof S[i][j] === 'number') l.push(S[i][j]);
    return fn(l);
  }
  function buscarAprox(lista, x, A, descendente) {   // la última posición <= x (lista ascendente), como MATCH tipo 1
    let lo = 0, hi = lista.length - 1, mejor = -1;
    while (hi >= 0 && lista[hi] === null) hi--;   // las vacías del final no cuentan: MATCH(9E+307;A1:A99) da la última con datos
    while (lo <= hi) {
      const m = (lo + hi) >> 1, v = lista[m];
      const c = v === null ? -1 : A.comparar(v, x);
      if (descendente ? c >= 0 : c <= 0) { mejor = m; lo = m + 1; } else hi = m - 1;
    }
    while (mejor >= 0 && lista[mejor] === null) mejor--;
    return mejor;
  }
  function coincide(lista, x, tipo, A) {
    if (tipo === 0) {
      if (typeof x === 'string' && /[*?~]/.test(x)) { const cr = criterio(x, A); for (let i = 0; i < lista.length; i++) if (typeof lista[i] === 'string' && cr(lista[i])) return i; return -1; }
      for (let i = 0; i < lista.length; i++) if (lista[i] !== null && A.comparar(lista[i], x) === 0 && typeof lista[i] === typeof x) return i;
      if (typeof x === 'number') for (let i = 0; i < lista.length; i++) if (lista[i] === x) return i;
      return -1;
    }
    return buscarAprox(lista, x, A, tipo < 0);
  }
  const columna = (M, j) => M.map((f) => f[j]);

  // TEXT lee el código de formato con la configuración regional del equipo, como Excel: el separador decimal regional es el decimal, el de
  // miles regional agrupa y un '.' o ',' que no es ninguno de los dos queda como carácter literal (en es-CR, que agrupa con espacio, el
  // '.' de "0.00" es literal: TEXT(4;"0.00") da "0.04"). Se traduce a la forma canónica (punto decimal, coma de miles) para formatear.
  // Si se agrupa con espacio, un espacio sin escapar después de un dígito es esa coma: "# ###" agrupa y "0 " divide entre mil. Con una fracción en cualquier
  // sección da #VALUE! (salvo justo antes de la barra, donde no cuenta). Medido en Excel: pruebas/unidad/hojas-excel-medido.prueba.mjs. Devuelve null si Excel lo rechaza.
  const SEPARADORES = { es: { decimal: ',', miles: '\u00a0' }, en: { decimal: '.', miles: ',' } };
  function codigoCanonico(codigo, sep) {
    if (sep.decimal === '.' && sep.miles === ',') return { canon: codigo };
    const agrupaConEspacio = sep.miles === '\u00a0' || sep.miles === ' ';
    const esFraccion = (s) => { const v = s.replace(/"[^"]*"|\\.|_.|\*.|\[[^\]]*\]/g, ''); return v.includes('/') && !XlsxFormato.esFecha(s); };
    let r = '', inicio = 0, hayDigito = false, espacioEnSeccion = false, espacio = false;
    const cerrarSeccion = () => { const mala = espacioEnSeccion && esFraccion(r.slice(inicio)); inicio = r.length; hayDigito = false; espacioEnSeccion = false; return mala; };
    for (let i = 0; i < codigo.length; i++) {
      const ch = codigo[i];
      if (ch === ';' && cerrarSeccion()) return null;
      if (ch === '0' || ch === '#' || ch === '?') hayDigito = true;
      if (ch === ' ' && agrupaConEspacio && hayDigito) {
        if (codigo[i + 1] === '/') continue;
        r += ','; espacio = true; espacioEnSeccion = true;
        continue;
      }
      if (ch === ';') { r += ch; inicio = r.length; continue; }
      if (ch === '"') { const k = codigo.indexOf('"', i + 1); const fin = k < 0 ? codigo.length : k + 1; r += codigo.slice(i, fin); i = fin - 1; }
      else if (ch === '\\' || ch === '_' || ch === '*') { r += codigo.slice(i, i + 2); i++; }
      else if (ch === '[') { const k = codigo.indexOf(']', i); const fin = k < 0 ? codigo.length : k + 1; r += codigo.slice(i, fin); i = fin - 1; }
      else if (ch === sep.decimal) r += '.';
      else if (ch === sep.miles) r += ',';
      else if (ch === '.' || ch === ',') r += '"' + ch + '"';
      else r += ch;
    }
    if (cerrarSeccion()) return null;
    return { canon: r, espacio };
  }
  const textoFormato = (v, codigo, A) => {
    if (typeof XlsxFormato === 'undefined') throw new A.NoSoportada('TEXT');
    const idioma = A.fuente.idioma === 'en' ? 'en' : 'es';
    const sep = A.fuente.separadores || SEPARADORES[idioma];
    const traducido = codigoCanonico(codigo, sep);
    if (!traducido) return A.E.valor;
    const canon = traducido.canon;
    if (typeof v === 'number' && v < 0 && XlsxFormato.esFecha(canon)) return A.E.valor;   // Excel no tiene fechas negativas: TEXT da #¡VALOR!
    return XlsxFormato.formatear(v, canon, { idioma, decimal: sep.decimal, miles: traducido.espacio ? ' ' : sep.miles, fecha1904 }).texto;
  };

  const FUNCIONES = {
    // matemáticas
    SUM: (a, c, A) => sumar(numeros(a, c, A)),
    PRODUCT: (a, c, A) => numeros(a, c, A).reduce((p, x) => p * x, 1),
    AVERAGE: (a, c, A) => { const l = numeros(a, c, A); return l.length ? sumar(l) / l.length : A.E.div; },
    AVERAGEA: (a, c, A) => { const l = numeros(a, c, A, { conTodo: true }); return l.length ? sumar(l) / l.length : A.E.div; },
    MIN: (a, c, A) => { const l = numeros(a, c, A); return l.length ? Math.min(...l) : 0; },
    MAX: (a, c, A) => { const l = numeros(a, c, A); return l.length ? Math.max(...l) : 0; },
    COUNT: (a, c, A) => { let n = 0; A.recorrer(a, c, (v, r) => { if (typeof v === 'number') n++; else if (!r && (typeof v === 'boolean' || (typeof v === 'string' && A.textoANumero(v) !== null))) n++; }); return n; },
    COUNTA: (a, c, A) => { let n = 0; A.recorrer(a, c, (v) => { if (v !== null) n++; }); return n; },
    COUNTBLANK: (a, c, A) => {
      let n = 0;
      for (const x of a) {
        const v = x === null ? null : A.evaluarNodo(x, c);
        // lo que pasa de lo usado también está vacío (recorrerValor no llega ahí)
        if (A.esRef(v)) { const { f1, c1 } = A.recortar(v); n += (v.f1 - v.f0 + 1) * (v.c1 - v.c0 + 1) - Math.max(0, f1 - v.f0 + 1) * Math.max(0, c1 - v.c0 + 1); }
        A.recorrerValor(v, (y) => { if (y === null || y === '') n++; });
      }
      return n;
    },
    ABS: (a, c, A) => Math.abs(A.num(a[0], c)),
    INT: (a, c, A) => Math.floor(A.num(a[0], c)),
    TRUNC: (a, c, A) => truncar(A.num(a[0], c), A.opcional(a[1], c, 0, A.num)),
    ROUND: (a, c, A) => redondear(A.num(a[0], c), Math.trunc(A.num(a[1], c))),
    ROUNDUP: (a, c, A) => arriba(A.num(a[0], c), Math.trunc(A.num(a[1], c))),
    ROUNDDOWN: (a, c, A) => truncar(A.num(a[0], c), Math.trunc(A.num(a[1], c))),
    MROUND: (a, c, A) => { const x = A.num(a[0], c), m = A.num(a[1], c); if (m === 0) return 0; if (Math.sign(x) * Math.sign(m) < 0) return A.E.num; return redondear(Math.round(Number((x / m).toPrecision(15))) * m, 10); },
    CEILING: (a, c, A) => { const x = A.num(a[0], c), m = A.opcional(a[1], c, 1, A.num); if (m === 0) return 0; if (x > 0 && m < 0) return A.E.num; return Math.ceil(Number((x / m).toPrecision(15))) * m; },
    'CEILING.MATH': (a, c, A) => { const x = A.num(a[0], c), m = Math.abs(A.opcional(a[1], c, 1, A.num)); if (m === 0) return 0; return Math.ceil(Number((x / m).toPrecision(15))) * m; },
    FLOOR: (a, c, A) => { const x = A.num(a[0], c), m = A.opcional(a[1], c, 1, A.num); if (m === 0) return A.E.div; if (x > 0 && m < 0) return A.E.num; return Math.floor(Number((x / m).toPrecision(15))) * m; },
    'FLOOR.MATH': (a, c, A) => { const x = A.num(a[0], c), m = Math.abs(A.opcional(a[1], c, 1, A.num)); if (m === 0) return 0; return Math.floor(Number((x / m).toPrecision(15))) * m; },
    MOD: (a, c, A) => { const x = A.num(a[0], c), d = A.num(a[1], c); if (d === 0) return A.E.div; const r = x - d * Math.floor(x / d); return Number(r.toPrecision(15)); },
    POWER: (a, c, A) => { const r = Math.pow(A.num(a[0], c), A.num(a[1], c)); return Number.isFinite(r) ? r : A.E.num; },
    SQRT: (a, c, A) => { const x = A.num(a[0], c); return x < 0 ? A.E.num : Math.sqrt(x); },
    EXP: (a, c, A) => Math.exp(A.num(a[0], c)),
    LN: (a, c, A) => { const x = A.num(a[0], c); return x <= 0 ? A.E.num : Math.log(x); },
    LOG10: (a, c, A) => { const x = A.num(a[0], c); return x <= 0 ? A.E.num : Math.log10(x); },
    LOG: (a, c, A) => { const x = A.num(a[0], c), b = A.opcional(a[1], c, 10, A.num); return x <= 0 || b <= 0 || b === 1 ? A.E.num : Math.log(x) / Math.log(b); },
    PI: () => Math.PI,
    SIGN: (a, c, A) => Math.sign(A.num(a[0], c)),
    EVEN: (a, c, A) => { const x = A.num(a[0], c); const r = Math.ceil(Math.abs(x) / 2) * 2; return x < 0 ? -r : r; },
    ODD: (a, c, A) => { const x = A.num(a[0], c); let r = Math.ceil(Math.abs(x)); if (r % 2 === 0) r += 1; return x < 0 ? -r : r; },
    FACT: (a, c, A) => { const n = Math.floor(A.num(a[0], c)); if (n < 0) return A.E.num; let r = 1; for (let k = 2; k <= n; k++) r *= k; return Number.isFinite(r) ? r : A.E.num; },
    QUOTIENT: (a, c, A) => { const d = A.num(a[1], c); if (d === 0) return A.E.div; return Math.trunc(A.num(a[0], c) / d); },
    GCD: (a, c, A) => { const l = numeros(a, c, A).map(Math.trunc); const g = (x, y) => (y ? g(y, x % y) : x); return l.reduce((p, x) => g(p, x), 0); },
    LCM: (a, c, A) => { const l = numeros(a, c, A).map(Math.trunc); const g = (x, y) => (y ? g(y, x % y) : x); return l.reduce((p, x) => (p && x ? (p * x) / g(p, x) : 0), 1); },
    RAND: () => Math.random(),
    RANDBETWEEN: (a, c, A) => { const lo = Math.ceil(A.num(a[0], c)), hi = Math.floor(A.num(a[1], c)); if (lo > hi) return A.E.num; return lo + Math.floor(Math.random() * (hi - lo + 1)); },
    SIN: (a, c, A) => Math.sin(A.num(a[0], c)), COS: (a, c, A) => Math.cos(A.num(a[0], c)), TAN: (a, c, A) => Math.tan(A.num(a[0], c)),
    ASIN: (a, c, A) => Math.asin(A.num(a[0], c)), ACOS: (a, c, A) => Math.acos(A.num(a[0], c)), ATAN: (a, c, A) => Math.atan(A.num(a[0], c)),
    ATAN2: (a, c, A) => { const x = A.num(a[0], c), y = A.num(a[1], c); if (x === 0 && y === 0) return A.E.div; return Math.atan2(y, x); },
    DEGREES: (a, c, A) => A.num(a[0], c) * 180 / Math.PI, RADIANS: (a, c, A) => A.num(a[0], c) * Math.PI / 180,
    SUMSQ: (a, c, A) => sumar(numeros(a, c, A).map((x) => x * x)),
    SUMPRODUCT: (a, c, A) => {
      const ms = a.map((x) => A.matrizDe(A.evaluarNodo(x, Object.assign({}, c, { matriz: true }))));
      for (const m of ms) if (m.length === 1 && m[0].length === 1 && esError(m[0][0])) return m[0][0];   // un argumento que es un error (Datos!#REF!) lo propaga
      const nf = ms[0].length, nc = ms[0][0].length;
      if (ms.some((m) => m.length !== nf || m[0].length !== nc)) return A.E.valor;
      let s = 0;
      for (let i = 0; i < nf; i++) for (let j = 0; j < nc; j++) {
        let p = 1;
        for (const m of ms) { const v = m[i][j]; if (esError(v)) return v; p *= typeof v === 'number' ? v : typeof v === 'boolean' && ms.length === 1 ? 0 : (typeof v === 'boolean' ? (v ? 1 : 0) : 0); }
        s += p;
      }
      return s;
    },
    // SUBTOTAL nunca cuenta otros SUBTOTAL/AGGREGATE del rango; 101–111 además saltan las filas ocultas
    SUBTOTAL: (a, c, A) => {
      const n = Math.trunc(A.num(a[0], c)), tipo = n % 100;
      if (!(n >= 1 && n <= 11) && !(n >= 101 && n <= 111)) return A.E.valor;
      return agregarLista(tipo, valoresFiltrados(a.slice(1), c, A, { anidadas: true, ocultas: n > 100, errores: false }), A);
    },
    // AGGREGATE(función; opciones; ref1; …) u AGGREGATE(función; opciones; matriz; k). Opciones: 0–3 saltan SUBTOTAL/AGGREGATE anidados,
    // 1, 3, 5 y 7 saltan filas ocultas, 2, 3, 6 y 7 saltan errores.
    AGGREGATE: (a, c, A) => {
      const fn = Math.trunc(A.num(a[0], c)), op = Math.trunc(A.opcional(a[1], c, 0, A.num));
      if (fn < 1 || fn > 19 || op < 0 || op > 7 || a.length < 3) return A.E.valor;
      const filtro = { anidadas: op <= 3, ocultas: op % 2 === 1, errores: op === 2 || op === 3 || op === 6 || op === 7 };
      if (fn <= 13) {   // forma de referencia: solo acepta rangos (con una matriz, Excel da #¡VALOR!)
        for (const x of a.slice(2)) { if (x === null) continue; const v = A.evaluarNodo(x, c); if (esError(v)) return v; if (!A.esRef(v) && !A.esUnion(v)) return A.E.valor; }
        return agregarLista(fn, valoresFiltrados(a.slice(2), c, A, filtro), A);
      }
      if (a.length < 4) return A.E.valor;
      const lista = valoresFiltrados([a[2]], Object.assign({}, c, { matriz: true }), A, filtro);   // forma de matriz: 1/A1:A9 vale
      const l = [];
      for (const { v } of lista) { if (esError(v)) return v; if (typeof v === 'number') l.push(v); }
      const k = A.num(a[3], c);
      if (fn === 14 || fn === 15) { l.sort((x, y) => (fn === 14 ? y - x : x - y)); const i = Math.ceil(k); return i < 1 || i > l.length ? A.E.num : l[i - 1]; }
      if (fn === 16) return percentil(l, k, true, A);
      if (fn === 17) return k < 0 || k > 4 ? A.E.num : percentil(l, Math.trunc(k) / 4, true, A);
      if (fn === 18) return percentil(l, k, false, A);
      return k <= 0 || k >= 4 ? A.E.num : percentil(l, Math.trunc(k) / 4, false, A);
    },
    PERCENTILE: (a, c, A) => percentil(numeros([a[0]], c, A), A.num(a[1], c), true, A),
    'PERCENTILE.INC': (a, c, A) => percentil(numeros([a[0]], c, A), A.num(a[1], c), true, A),
    'PERCENTILE.EXC': (a, c, A) => percentil(numeros([a[0]], c, A), A.num(a[1], c), false, A),
    QUARTILE: (a, c, A) => { const q = Math.trunc(A.num(a[1], c)); return q < 0 || q > 4 ? A.E.num : percentil(numeros([a[0]], c, A), q / 4, true, A); },
    'QUARTILE.INC': (a, c, A) => { const q = Math.trunc(A.num(a[1], c)); return q < 0 || q > 4 ? A.E.num : percentil(numeros([a[0]], c, A), q / 4, true, A); },
    'QUARTILE.EXC': (a, c, A) => { const q = Math.trunc(A.num(a[1], c)); return q <= 0 || q >= 4 ? A.E.num : percentil(numeros([a[0]], c, A), q / 4, false, A); },
    // estadística
    MEDIAN: (a, c, A) => { const l = numeros(a, c, A).sort((x, y) => x - y); if (!l.length) return A.E.num; const m = l.length >> 1; return l.length % 2 ? l[m] : (l[m - 1] + l[m]) / 2; },
    MODE: (a, c, A) => { const l = numeros(a, c, A); const cuenta = new Map(); let mejor = null, n = 1; for (const x of l) { const k = (cuenta.get(x) || 0) + 1; cuenta.set(x, k); if (k > n) { n = k; mejor = x; } } return mejor === null ? A.E.na : mejor; },
    STDEV: (a, c, A) => desviacion(numeros(a, c, A), 1, A), 'STDEV.S': (a, c, A) => desviacion(numeros(a, c, A), 1, A),
    STDEVP: (a, c, A) => desviacion(numeros(a, c, A), 0, A), 'STDEV.P': (a, c, A) => desviacion(numeros(a, c, A), 0, A),
    VAR: (a, c, A) => varianza(numeros(a, c, A), 1, A), 'VAR.S': (a, c, A) => varianza(numeros(a, c, A), 1, A),
    VARP: (a, c, A) => varianza(numeros(a, c, A), 0, A), 'VAR.P': (a, c, A) => varianza(numeros(a, c, A), 0, A),
    LARGE: (a, c, A) => { const l = numeros([a[0]], c, A).sort((x, y) => y - x), k = Math.ceil(A.num(a[1], c)); return k < 1 || k > l.length ? A.E.num : l[k - 1]; },
    SMALL: (a, c, A) => { const l = numeros([a[0]], c, A).sort((x, y) => x - y), k = Math.ceil(A.num(a[1], c)); return k < 1 || k > l.length ? A.E.num : l[k - 1]; },
    RANK: (a, c, A) => rango(a, c, A), 'RANK.EQ': (a, c, A) => rango(a, c, A),
    // lógicas
    IF: (a, c, A) => {
      const cond = A.escalar(A.evaluarNodo(a[0], c), c);
      if (esError(cond)) return cond;
      const si = A.aLogico(cond);
      const rama = si ? a[1] : a[2];
      if (rama === undefined) return si ? true : false;
      if (rama === null) return 0;
      return A.evaluarNodo(rama, c);
    },
    IFS: (a, c, A) => { for (let k = 0; k + 1 < a.length; k += 2) if (A.log(a[k], c)) return A.evaluarNodo(a[k + 1], c); return A.E.na; },
    IFERROR: (a, c, A) => { const v = A.escalar(A.evaluarNodo(a[0], c), c); return esError(v) ? A.evaluarNodo(a[1], c) : v; },
    IFNA: (a, c, A) => { const v = A.escalar(A.evaluarNodo(a[0], c), c); return v === A.E.na ? A.evaluarNodo(a[1], c) : v; },
    AND: (a, c, A) => { let r = true, hubo = false; A.recorrer(a, c, (v, deRango) => { if (esError(v)) A.lanzar(v); if (deRango && (v === null || typeof v === 'string')) return; hubo = true; if (!A.aLogico(v)) r = false; }); return hubo ? r : A.E.valor; },
    OR: (a, c, A) => { let r = false, hubo = false; A.recorrer(a, c, (v, deRango) => { if (esError(v)) A.lanzar(v); if (deRango && (v === null || typeof v === 'string')) return; hubo = true; if (A.aLogico(v)) r = true; }); return hubo ? r : A.E.valor; },
    XOR: (a, c, A) => { let n = 0; A.recorrer(a, c, (v, deRango) => { if (esError(v)) A.lanzar(v); if (deRango && (v === null || typeof v === 'string')) return; if (A.aLogico(v)) n++; }); return n % 2 === 1; },
    NOT: (a, c, A) => !A.log(a[0], c),
    TRUE: () => true, FALSE: () => false,
    SWITCH: (a, c, A) => { const x = A.val(a[0], c); for (let k = 1; k + 1 < a.length; k += 2) if (A.comparar(x, A.val(a[k], c)) === 0) return A.evaluarNodo(a[k + 1], c); return a.length % 2 === 0 ? A.evaluarNodo(a[a.length - 1], c) : A.E.na; },
    // información
    ISBLANK: (a, c, A) => A.escalar(A.evaluarNodo(a[0], c), c) === null,
    ISNUMBER: (a, c, A) => typeof A.escalar(A.evaluarNodo(a[0], c), c) === 'number',
    ISTEXT: (a, c, A) => typeof A.escalar(A.evaluarNodo(a[0], c), c) === 'string',
    ISNONTEXT: (a, c, A) => typeof A.escalar(A.evaluarNodo(a[0], c), c) !== 'string',
    ISLOGICAL: (a, c, A) => typeof A.escalar(A.evaluarNodo(a[0], c), c) === 'boolean',
    ISERROR: (a, c, A) => esError(A.escalar(A.evaluarNodo(a[0], c), c)),
    ISERR: (a, c, A) => { const v = A.escalar(A.evaluarNodo(a[0], c), c); return esError(v) && v !== A.E.na; },
    ISNA: (a, c, A) => A.escalar(A.evaluarNodo(a[0], c), c) === A.E.na,
    ISEVEN: (a, c, A) => Math.trunc(A.num(a[0], c)) % 2 === 0,
    ISODD: (a, c, A) => Math.abs(Math.trunc(A.num(a[0], c)) % 2) === 1,
    ISREF: (a, c, A) => A.esRef(A.evaluarNodo(a[0], c)),
    ISFORMULA: (a, c, A) => { const r = A.evaluarNodo(a[0], c); if (!A.esRef(r)) return A.E.valor; return !!(A.fuente.esFormula && A.fuente.esFormula(r.h, r.f0, r.c0)); },
    NA: (a, c, A) => A.E.na,
    N: (a, c, A) => { let v = A.evaluarNodo(a[0], c); v = A.esRef(v) ? A.valorCelda(v.h, v.f0, v.c0) : A.escalar(v, c); return typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : esError(v) ? v : 0; },
    T: (a, c, A) => { let v = A.evaluarNodo(a[0], c); v = A.esRef(v) ? A.valorCelda(v.h, v.f0, v.c0) : A.escalar(v, c); return typeof v === 'string' ? v : esError(v) ? v : ''; },
    TYPE: (a, c, A) => { const v = A.evaluarNodo(a[0], c); if (A.esMat(v) || (A.esRef(v) && (v.f0 !== v.f1 || v.c0 !== v.c1))) return 64; const e = A.escalar(v, c); return typeof e === 'number' || e === null ? 1 : typeof e === 'string' ? 2 : typeof e === 'boolean' ? 4 : 16; },
    'ERROR.TYPE': (a, c, A) => { const v = A.escalar(A.evaluarNodo(a[0], c), c); const k = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'].indexOf(esError(v) ? v.codigo : ''); return k < 0 ? A.E.na : k + 1; },
    // condicionales
    COUNTIF: (a, c, A) => { const cr = criterio(A.escalar(A.evaluarNodo(a[1], c), c), A); let n = 0; A.recorrerValor(A.evaluarNodo(a[0], c), (v) => { if (cr(v)) n++; }); return n; },
    COUNTIFS: (a, c, A) => { const ok = filtrarPares(a, 0, c, A); let n = 0; for (const f of ok) for (const x of f) if (x) n++; return n; },
    SUMIF: (a, c, A) => {
      const base = A.evaluarNodo(a[0], c), M = A.matrizDe(base), cr = criterio(A.escalar(A.evaluarNodo(a[1], c), c), A);
      const S = a[2] ? rangoSuma(a[2], base, c, A) : M;
      let s = 0;
      for (let i = 0; i < M.length; i++) for (let j = 0; j < M[i].length; j++) if (cr(M[i][j])) { const v = (S[i] || [])[j]; if (esError(v)) return v; if (typeof v === 'number') s += v; }
      return s;
    },
    SUMIFS: (a, c, A) => conjunto(a, c, A, (l) => sumar(l)),
    AVERAGEIFS: (a, c, A) => conjunto(a, c, A, (l) => (l.length ? sumar(l) / l.length : A.E.div)),
    MAXIFS: (a, c, A) => conjunto(a, c, A, (l) => (l.length ? Math.max(...l) : 0)),
    MINIFS: (a, c, A) => conjunto(a, c, A, (l) => (l.length ? Math.min(...l) : 0)),
    AVERAGEIF: (a, c, A) => {
      const base = A.evaluarNodo(a[0], c), M = A.matrizDe(base), cr = criterio(A.escalar(A.evaluarNodo(a[1], c), c), A);
      const S = a[2] ? rangoSuma(a[2], base, c, A) : M;
      let s = 0, n = 0;
      for (let i = 0; i < M.length; i++) for (let j = 0; j < M[i].length; j++) if (cr(M[i][j])) { const v = (S[i] || [])[j]; if (typeof v === 'number') { s += v; n++; } }
      return n ? s / n : A.E.div;
    },
    // búsqueda y referencia
    VLOOKUP: (a, c, A) => {
      const x = A.val(a[0], c), M = A.matrizDe(A.evaluarNodo(a[1], c)), k = Math.trunc(A.num(a[2], c)), aprox = A.opcional(a[3], c, true, A.log);
      if (k < 1) return A.E.valor; if (k > M[0].length) return A.E.ref;
      const i = coincide(columna(M, 0), x, aprox ? 1 : 0, A);
      return i < 0 ? A.E.na : M[i][k - 1];
    },
    HLOOKUP: (a, c, A) => {
      const x = A.val(a[0], c), M = A.matrizDe(A.evaluarNodo(a[1], c)), k = Math.trunc(A.num(a[2], c)), aprox = A.opcional(a[3], c, true, A.log);
      if (k < 1) return A.E.valor; if (k > M.length) return A.E.ref;
      const i = coincide(M[0], x, aprox ? 1 : 0, A);
      return i < 0 ? A.E.na : M[k - 1][i];
    },
    LOOKUP: (a, c, A) => {
      const x = A.val(a[0], c), M = A.matrizDe(A.evaluarNodo(a[1], c));
      const vertical = M.length >= M[0].length;
      const busca = vertical ? columna(M, 0) : M[0];
      const i = coincide(busca, x, 1, A);
      if (i < 0) return A.E.na;
      if (a[2]) { const R = A.matrizDe(A.evaluarNodo(a[2], c)); const plano = R.length === 1 ? R[0] : columna(R, 0); return i < plano.length ? plano[i] : A.E.na; }
      return vertical ? M[i][M[0].length - 1] : M[M.length - 1][i];
    },
    MATCH: (a, c, A) => {
      const x = A.val(a[0], c), M = A.matrizDe(A.evaluarNodo(a[1], c)), tipo = Math.sign(A.opcional(a[2], c, 1, A.num));
      if (M.length > 1 && M[0].length > 1) return A.E.na;
      const lista = M.length === 1 ? M[0] : columna(M, 0);
      const i = coincide(lista, x, tipo, A);
      return i < 0 ? A.E.na : i + 1;
    },
    XMATCH: (a, c, A) => {
      const x = A.val(a[0], c), M = A.matrizDe(A.evaluarNodo(a[1], c)), modo = Math.trunc(A.opcional(a[2], c, 0, A.num));
      const lista = M.length === 1 ? M[0] : columna(M, 0);
      const i = xbuscar(lista, x, modo, Math.trunc(A.opcional(a[3], c, 1, A.num)), A);
      return i < 0 ? A.E.na : i + 1;
    },
    XLOOKUP: (a, c, A) => {
      const x = A.val(a[0], c), M = A.matrizDe(A.evaluarNodo(a[1], c)), R = A.matrizDe(A.evaluarNodo(a[2], c));
      const vertical = M[0].length === 1;
      const lista = vertical ? columna(M, 0) : M[0];
      const i = xbuscar(lista, x, Math.trunc(A.opcional(a[4], c, 0, A.num)), Math.trunc(A.opcional(a[5], c, 1, A.num)), A);
      if (i < 0) return a[3] !== undefined && a[3] !== null ? A.evaluarNodo(a[3], c) : A.E.na;
      if (vertical) return R.length === 1 ? R[0][0] : (R[i] ? (R[0].length === 1 ? R[i][0] : { tipo: 'mat', filas: [R[i]] }) : A.E.na);
      return R[0].length === 1 ? R[0][0] : (R[0][i] !== undefined ? (R.length === 1 ? R[0][i] : { tipo: 'mat', filas: R.map((f) => [f[i]]) }) : A.E.na);
    },
    INDEX: (a, c, A) => {
      const v = A.evaluarNodo(a[0], c);
      if (esError(v)) return v;
      const r = A.opcional(a[1], c, 0, A.num), k = A.opcional(a[2], c, 0, A.num);
      const fila = Math.trunc(r), col = Math.trunc(k);
      if (A.esRef(v)) {
        const nf = v.f1 - v.f0 + 1, nc = v.c1 - v.c0 + 1;
        let ff = fila, cc = col;
        if (nf === 1 && a[2] === undefined) { cc = fila; ff = 1; }
        if (ff < 0 || cc < 0 || ff > nf || cc > nc) return A.E.ref;
        return { tipo: 'ref', h: v.h, f0: ff ? v.f0 + ff - 1 : v.f0, f1: ff ? v.f0 + ff - 1 : v.f1, c0: cc ? v.c0 + cc - 1 : v.c0, c1: cc ? v.c0 + cc - 1 : v.c1 };
      }
      const M = A.matrizDe(v);
      let ff = fila, cc = col;
      if (M.length === 1 && a[2] === undefined) { cc = fila; ff = 1; }
      if (!ff) ff = 1; if (!cc) cc = 1;
      return M[ff - 1] && M[ff - 1][cc - 1] !== undefined ? M[ff - 1][cc - 1] : A.E.ref;
    },
    CHOOSE: (a, c, A) => { const k = Math.trunc(A.num(a[0], c)); if (k < 1 || k >= a.length) return A.E.valor; return A.evaluarNodo(a[k], c); },
    ROW: (a, c, A) => {
      if (!a.length || a[0] === null) return c.f + 1;
      const v = A.evaluarNodo(a[0], c);
      if (!A.esRef(v)) return A.E.valor;
      if (!c.matriz || v.f0 === v.f1) return v.f0 + 1;
      const fin = v.f1 - v.f0 < 1000000 ? v.f1 : Math.max(v.f0, A.recortar(v).f1);   // no lee celdas: la forma escrita, salvo la columna entera
      const filas = []; for (let f = v.f0; f <= fin; f++) filas.push([f + 1]);
      return { tipo: 'mat', filas };
    },
    COLUMN: (a, c, A) => {
      if (!a.length || a[0] === null) return c.c + 1;
      const v = A.evaluarNodo(a[0], c);
      if (!A.esRef(v)) return A.E.valor;
      if (!c.matriz || v.c0 === v.c1) return v.c0 + 1;
      const fila = []; for (let k = v.c0; k <= v.c1; k++) fila.push(k + 1);
      return { tipo: 'mat', filas: [fila] };
    },
    ROWS: (a, c, A) => { const v = A.evaluarNodo(a[0], c); if (A.esRef(v)) return v.f1 - v.f0 + 1; return A.matrizDe(v).length; },
    COLUMNS: (a, c, A) => { const v = A.evaluarNodo(a[0], c); if (A.esRef(v)) return v.c1 - v.c0 + 1; return A.matrizDe(v)[0].length; },
    OFFSET: (a, c, A) => {
      const v = A.evaluarNodo(a[0], c);
      if (!A.esRef(v)) return A.E.valor;
      const df = Math.trunc(A.num(a[1], c)), dc = Math.trunc(A.num(a[2], c));
      const alto = Math.trunc(A.opcional(a[3], c, v.f1 - v.f0 + 1, A.num)), ancho = Math.trunc(A.opcional(a[4], c, v.c1 - v.c0 + 1, A.num));
      const f0 = v.f0 + df, c0 = v.c0 + dc;
      if (f0 < 0 || c0 < 0 || alto < 1 || ancho < 1 || f0 + alto > MAX_FILAS || c0 + ancho > MAX_COLS) return A.E.ref;
      return { tipo: 'ref', h: v.h, f0, c0, f1: f0 + alto - 1, c1: c0 + ancho - 1 };
    },
    // texto
    LEN: (a, c, A) => A.txt(a[0], c).length,
    LEFT: (a, c, A) => { const n = A.opcional(a[1], c, 1, A.num); if (n < 0) return A.E.valor; return A.txt(a[0], c).slice(0, Math.trunc(n)); },
    RIGHT: (a, c, A) => { const n = A.opcional(a[1], c, 1, A.num); if (n < 0) return A.E.valor; const t = A.txt(a[0], c); return Math.trunc(n) ? t.slice(-Math.trunc(n)) : ''; },
    MID: (a, c, A) => { const t = A.txt(a[0], c), i = Math.trunc(A.num(a[1], c)), n = Math.trunc(A.num(a[2], c)); if (i < 1 || n < 0) return A.E.valor; return t.substr(i - 1, n); },
    UPPER: (a, c, A) => A.txt(a[0], c).toUpperCase(),
    LOWER: (a, c, A) => A.txt(a[0], c).toLowerCase(),
    PROPER: (a, c, A) => A.txt(a[0], c).toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (m, p, l) => p + l.toUpperCase()),
    TRIM: (a, c, A) => A.txt(a[0], c).replace(/ +/g, ' ').trim(),
    CLEAN: (a, c, A) => A.txt(a[0], c).replace(/[\u0000-\u001F]/g, ''),
    CONCATENATE: (a, c, A) => a.map((x) => A.txt(x, c)).join(''),
    CONCAT: (a, c, A) => { let s = ''; A.recorrer(a, c, (v) => { s += A.aTexto(v); }); return s; },
    HYPERLINK: (a, c, A) => { const url = A.val(a[0], c); if (esError(url)) return url; if (a.length < 2 || a[1] === null) return A.aTexto(url); const v = A.val(a[1], c); return v === null ? 0 : v; },
    // TEXTBEFORE/TEXTAFTER(texto; delimitador; [instancia]; [mayúsculas: 0 sí, 1 no]; [final cuenta como delimitador]; [si no se encuentra])
    TEXTBEFORE: (a, c, A) => textoAntesDespues(a, c, A, true),
    TEXTAFTER: (a, c, A) => textoAntesDespues(a, c, A, false),
    TEXTJOIN: (a, c, A) => {
      const seps = []; A.recorrerValor(A.evaluarNodo(a[0], c), (v) => seps.push(A.aTexto(v)));
      const saltar = A.log(a[1], c); const partes = [];
      A.recorrer(a.slice(2), c, (v) => { const t = A.aTexto(v); if (!(saltar && t === '')) partes.push(t); });
      return partes.reduce((s, t, k) => (k ? s + seps[(k - 1) % seps.length] + t : t), '');
    },
    REPT: (a, c, A) => { const n = Math.trunc(A.num(a[1], c)); if (n < 0) return A.E.valor; const t = A.txt(a[0], c); if (t.length * n > 32767) return A.E.valor; return t.repeat(n); },
    SUBSTITUTE: (a, c, A) => {
      const t = A.txt(a[0], c), viejo = A.txt(a[1], c), nuevo = A.txt(a[2], c);
      if (!viejo) return t;
      if (a[3] === undefined) return t.split(viejo).join(nuevo);
      const k = Math.trunc(A.num(a[3], c)); if (k < 1) return A.E.valor;
      let pos = -1; for (let i = 0; i < k; i++) { pos = t.indexOf(viejo, pos + 1); if (pos < 0) return t; }
      return t.slice(0, pos) + nuevo + t.slice(pos + viejo.length);
    },
    REPLACE: (a, c, A) => { const t = A.txt(a[0], c), i = Math.trunc(A.num(a[1], c)), n = Math.trunc(A.num(a[2], c)); if (i < 1 || n < 0) return A.E.valor; return t.slice(0, i - 1) + A.txt(a[3], c) + t.slice(i - 1 + n); },
    FIND: (a, c, A) => { const q = A.txt(a[0], c), t = A.txt(a[1], c), i = Math.trunc(A.opcional(a[2], c, 1, A.num)); if (i < 1 || i > t.length + 1) return A.E.valor; const p = t.indexOf(q, i - 1); return p < 0 ? A.E.valor : p + 1; },
    SEARCH: (a, c, A) => {
      const q = A.txt(a[0], c), t = A.txt(a[1], c), i = Math.trunc(A.opcional(a[2], c, 1, A.num));
      if (i < 1 || i > t.length + 1) return A.E.valor;
      const re = new RegExp(q.replace(/~([*?~])|([*?])|([.+^${}()|[\]\\])/g, (m, esc, w, meta) => (esc ? '\\' + esc : w === '*' ? '[\\s\\S]*?' : w === '?' ? '[\\s\\S]' : '\\' + meta)), 'i');
      const m = re.exec(t.slice(i - 1));
      return m ? m.index + i : A.E.valor;
    },
    EXACT: (a, c, A) => A.txt(a[0], c) === A.txt(a[1], c),
    // VALUE lee el texto con la configuración regional, como Excel: en es-CR (coma decimal) «2,5» es un número y «12.5» no lo es (oráculo de Excel del 2026-10-25)
    VALUE: (a, c, A) => {
      const v = A.val(a[0], c);
      if (typeof v === 'number') return v;
      let t = A.aTexto(v);
      if (A.fuente.idioma !== 'en') {
        if (/\d\.\d/.test(t)) return A.E.valor;
        t = t.replace(/(\d)[\s ](?=\d{3}(?!\d))/g, '$1').replace(/(\d),(\d)/g, '$1.$2');
      }
      const n = A.textoANumero(t);
      return n === null ? A.E.valor : n;
    },
    NUMBERVALUE: (a, c, A) => { const t = A.txt(a[0], c), dec = A.opcional(a[1], c, '.', A.txt), mil = A.opcional(a[2], c, ',', A.txt); const n = A.textoANumero(t.split(mil).join('').split(dec).join('.')); return n === null ? A.E.valor : n; },
    TEXT: (a, c, A) => { const v = A.val(a[0], c), f = A.txt(a[1], c); const n = typeof v === 'number' ? v : typeof v === 'string' ? A.textoANumero(v) : null; return textoFormato(n === null ? v : n, f, A); },
    FIXED: (a, c, A) => { const d = Math.trunc(A.opcional(a[1], c, 2, A.num)), sin = A.opcional(a[2], c, false, A.log); const x = redondear(A.num(a[0], c), d), t = x.toLocaleString('en-US', { minimumFractionDigits: Math.max(0, d), maximumFractionDigits: Math.max(0, d), useGrouping: !sin }); return A.fuente.idioma === 'es' ? t.replace(/[.,]/g, (m) => (m === '.' ? ',' : ' ')) : t; },
    CHAR: (a, c, A) => { const n = Math.trunc(A.num(a[0], c)); if (n < 1 || n > 255) return A.E.valor; return String.fromCharCode(n); },
    UNICHAR: (a, c, A) => { const n = Math.trunc(A.num(a[0], c)); if (n < 1 || n > 0x10FFFF) return A.E.valor; return String.fromCodePoint(n); },
    CODE: (a, c, A) => { const t = A.txt(a[0], c); return t ? t.charCodeAt(0) : A.E.valor; },
    UNICODE: (a, c, A) => { const t = A.txt(a[0], c); return t ? t.codePointAt(0) : A.E.valor; },
    // fechas
    DATE: (a, c, A) => {
      let y = Math.trunc(A.num(a[0], c)); const m = Math.trunc(A.num(a[1], c)), d = Math.trunc(A.num(a[2], c));
      if (y < 0 || y > 9999) return A.E.num;
      if (y < 1900) y += 1900;
      const ms = Date.UTC(y, m - 1, 1);
      const base = new Date(ms);
      const s = serialDe(base.getUTCFullYear(), base.getUTCMonth() + 1, 1) + d - 1;
      return s < 0 ? A.E.num : s;
    },
    TIME: (a, c, A) => { const h = A.num(a[0], c), m = A.num(a[1], c), s = A.num(a[2], c); const t = (Math.trunc(h) * 3600 + Math.trunc(m) * 60 + Math.trunc(s)) / 86400; return t < 0 ? A.E.num : t - Math.floor(t); },
    YEAR: (a, c, A) => partesNum(a, c, A).a,
    MONTH: (a, c, A) => partesNum(a, c, A).m,
    DAY: (a, c, A) => partesNum(a, c, A).d,
    HOUR: (a, c, A) => { const x = A.num(a[0], c); return Math.floor(Number(((x - Math.floor(x)) * 24).toPrecision(12))); },
    MINUTE: (a, c, A) => { const x = A.num(a[0], c); const s = Math.round((x - Math.floor(x)) * 86400); return Math.floor(s / 60) % 60; },
    SECOND: (a, c, A) => { const x = A.num(a[0], c); const s = Math.round((x - Math.floor(x)) * 86400); return s % 60; },
    WEEKDAY: (a, c, A) => {
      const x = A.num(a[0], c), tipo = Math.trunc(A.opcional(a[1], c, 1, A.num));
      const ds = x < 1 && !fecha1904 ? 6 : partes(x).ds;   // 0 = sábado 0-ene-1900 según Excel
      switch (tipo) { case 1: case 17: return ds + 1; case 2: case 11: return ((ds + 6) % 7) + 1; case 3: return (ds + 6) % 7; case 12: return ((ds + 5) % 7) + 1; case 13: return ((ds + 4) % 7) + 1; case 14: return ((ds + 3) % 7) + 1; case 15: return ((ds + 2) % 7) + 1; case 16: return ((ds + 1) % 7) + 1; default: return A.E.num; }
    },
    TODAY: (a, c, A) => { const d = A.hoy(); return serialDe(d.getFullYear(), d.getMonth() + 1, d.getDate()); },
    NOW: (a, c, A) => { const d = A.hoy(); return serialDe(d.getFullYear(), d.getMonth() + 1, d.getDate()) + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400; },
    EDATE: (a, c, A) => { const p = partesNum(a, c, A), n = Math.trunc(A.num(a[1], c)); const ultimo = new Date(Date.UTC(p.a, p.m - 1 + n + 1, 0)).getUTCDate(); const d = new Date(Date.UTC(p.a, p.m - 1 + n, Math.min(p.d, ultimo))); const r = serialDe(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()); return r < 0 ? A.E.num : r; },
    EOMONTH: (a, c, A) => { const p = partesNum(a, c, A), n = Math.trunc(A.num(a[1], c)); const d = new Date(Date.UTC(p.a, p.m + n, 0)); const r = serialDe(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()); return r < 0 ? A.E.num : r; },
    DAYS: (a, c, A) => Math.trunc(A.num(a[0], c)) - Math.trunc(A.num(a[1], c)),
    DATEDIF: (a, c, A) => {
      const x = Math.trunc(A.num(a[0], c)), y = Math.trunc(A.num(a[1], c)), u = A.txt(a[2], c).toUpperCase();
      if (x > y) return A.E.num;
      const p = partes(x), q = partes(y);
      const meses = (q.a - p.a) * 12 + (q.m - p.m) - (q.d < p.d ? 1 : 0);
      switch (u) {
        case 'D': return y - x;
        case 'M': return meses;
        case 'Y': return Math.floor(meses / 12);
        case 'YM': return meses % 12;
        case 'MD': { let d = q.d - p.d; if (d < 0) d += new Date(Date.UTC(q.a, q.m - 1, 0)).getUTCDate(); return d; }
        case 'YD': { let d = y - serialDe(q.a, p.m, p.d); if (d < 0) d = y - serialDe(q.a - 1, p.m, p.d); return d; }
        default: return A.E.num;
      }
    },
    NETWORKDAYS: (a, c, A) => {
      let x = Math.trunc(A.num(a[0], c)), y = Math.trunc(A.num(a[1], c));
      const feriados = new Set(a[2] ? numeros([a[2]], c, A).map(Math.trunc) : []);
      const signo = x <= y ? 1 : -1; if (signo < 0) [x, y] = [y, x];
      let n = 0;
      for (let d = x; d <= y; d++) { const ds = partes(d).ds; if (ds !== 0 && ds !== 6 && !feriados.has(d)) n++; }
      return n * signo;
    },
    // fin de semana de las versiones .INTL: un número (1 = sáb y dom … 7 = vie y sáb, 11–17 = un solo día) o un texto de 7 ceros y unos (lunes primero)
    // DATEVALUE: el texto de una fecha como se escribiría en una celda (en el idioma del libro) → su número de serie, sin la hora
    DATEVALUE: (a, c, A) => {
      const t = A.txt(a[0], c);
      if (typeof EntradaCelda === 'undefined') throw new A.NoSoportada('DATEVALUE');
      const x = EntradaCelda.interpretar(t.trim(), { idioma: A.fuente.idioma === 'en' ? 'en' : 'es', fecha1904 });
      if (x.tipo !== 'numero' || !x.formato || !XlsxFormato.esFecha(x.formato)) return A.E.valor;
      return Math.floor(x.valor);
    },
    // YEARFRAC con las cinco bases de Excel (0: 30/360 de EE. UU., 1: real/real, 2: real/360, 3: real/365, 4: 30/360 europea)
    YEARFRAC: (a, c, A) => {
      let x = Math.trunc(A.num(a[0], c)), y = Math.trunc(A.num(a[1], c));
      const base = Math.trunc(A.opcional(a[2], c, 0, A.num));
      if (x < 0 || y < 0 || base < 0 || base > 4) return A.E.num;
      if (x > y) [x, y] = [y, x];
      const p = partes(x), q = partes(y);
      const bisiesto = (n) => (n % 4 === 0 && n % 100 !== 0) || n % 400 === 0;
      const finFeb = (d) => d.m === 2 && d.d === (bisiesto(d.a) ? 29 : 28);
      if (base === 0 || base === 4) {
        let d1 = p.d, d2 = q.d;
        if (base === 0) {
          if (finFeb(p)) { if (finFeb(q)) d2 = 30; d1 = 30; }
          if (d2 === 31 && d1 >= 30) d2 = 30;
          if (d1 === 31) d1 = 30;
        } else { if (d1 === 31) d1 = 30; if (d2 === 31) d2 = 30; }
        return ((q.a - p.a) * 360 + (q.m - p.m) * 30 + (d2 - d1)) / 360;
      }
      const dias = y - x;
      if (base === 2) return dias / 360;
      if (base === 3) return dias / 365;
      // real/real: dentro de un año, 366 si el periodo toca un 29 de febrero; si no, el promedio de los años que abarca
      const dentroDeUnAnio = p.a === q.a || (q.a === p.a + 1 && (p.m > q.m || (p.m === q.m && p.d >= q.d)));
      if (dentroDeUnAnio) {
        let den = 365;
        if (p.a === q.a) { if (bisiesto(p.a)) den = 366; }
        else {
          const tiene29 = (n) => { if (!bisiesto(n)) return false; const s29 = serialDe(n, 2, 29); return s29 >= x && s29 <= y; };
          if (tiene29(p.a) || tiene29(q.a)) den = 366;
        }
        return dias / den;
      }
      let total = 0;
      for (let n = p.a; n <= q.a; n++) total += bisiesto(n) ? 366 : 365;
      return dias / (total / (q.a - p.a + 1));
    },
    'NETWORKDAYS.INTL': (a, c, A) => {
      let x = Math.trunc(A.num(a[0], c)), y = Math.trunc(A.num(a[1], c));
      const libre = finDeSemana(a[2] ? A.val(a[2], c) : 1, A);
      if (!libre) return A.E.num;
      const feriados = new Set(a[3] ? numeros([a[3]], c, A).map(Math.trunc) : []);
      const signo = x <= y ? 1 : -1; if (signo < 0) [x, y] = [y, x];
      let n = 0;
      for (let d = x; d <= y; d++) if (!libre[partes(d).ds] && !feriados.has(d)) n++;
      return n * signo;
    },
    'WORKDAY.INTL': (a, c, A) => {
      let d = Math.trunc(A.num(a[0], c)), n = Math.trunc(A.num(a[1], c));
      const libre = finDeSemana(a[2] ? A.val(a[2], c) : 1, A);
      if (!libre || libre.every(Boolean)) return A.E.num;
      const feriados = new Set(a[3] ? numeros([a[3]], c, A).map(Math.trunc) : []);
      const paso = n >= 0 ? 1 : -1;
      while (n !== 0) { d += paso; if (!libre[partes(d).ds] && !feriados.has(d)) n -= paso; }
      return d;
    },
    WORKDAY: (a, c, A) => {
      let d = Math.trunc(A.num(a[0], c)), n = Math.trunc(A.num(a[1], c));
      const feriados = new Set(a[2] ? numeros([a[2]], c, A).map(Math.trunc) : []);
      const paso = n >= 0 ? 1 : -1;
      while (n !== 0) { d += paso; const ds = partes(d).ds; if (ds !== 0 && ds !== 6 && !feriados.has(d)) n -= paso; }
      return d;
    },
    WEEKNUM: (a, c, A) => { const x = Math.trunc(A.num(a[0], c)), tipo = Math.trunc(A.opcional(a[1], c, 1, A.num)); const p = partes(x); const primero = serialDe(p.a, 1, 1); const ds = partes(primero).ds; const inicio = tipo === 2 ? (ds + 6) % 7 : ds; return Math.floor((x - primero + inicio) / 7) + 1; },
    // finanzas básicas
    PMT: (a, c, A) => { const r = A.num(a[0], c), n = A.num(a[1], c), pv = A.num(a[2], c), fv = A.opcional(a[3], c, 0, A.num), tipo = A.opcional(a[4], c, 0, A.num) ? 1 : 0; if (n === 0) return A.E.num; if (r === 0) return -(pv + fv) / n; const p = Math.pow(1 + r, n); return -(r * (pv * p + fv)) / ((1 + r * tipo) * (p - 1)); },
    FV: (a, c, A) => { const r = A.num(a[0], c), n = A.num(a[1], c), pmt = A.num(a[2], c), pv = A.opcional(a[3], c, 0, A.num), tipo = A.opcional(a[4], c, 0, A.num) ? 1 : 0; if (r === 0) return -(pv + pmt * n); const p = Math.pow(1 + r, n); return -(pv * p + pmt * (1 + r * tipo) * (p - 1) / r); },
    PV: (a, c, A) => { const r = A.num(a[0], c), n = A.num(a[1], c), pmt = A.num(a[2], c), fv = A.opcional(a[3], c, 0, A.num), tipo = A.opcional(a[4], c, 0, A.num) ? 1 : 0; if (r === 0) return -(fv + pmt * n); const p = Math.pow(1 + r, n); return -(fv + pmt * (1 + r * tipo) * (p - 1) / r) / p; },
    // A1# (Excel lo guarda como _xlfn.ANCHORARRAY(A1)): todo el rango que desborda la matriz dinámica de A1
    ANCHORARRAY: (a, c, A) => {
      const r = A.evaluarNodo(a[0], c);
      if (!A.esRef(r)) return A.E.ref;
      const d = A.derrame(r.h, r.f0, r.c0);
      if (!d && A.esMatrizDeUnValor(r.h, r.f0, r.c0)) return r;   // medido: A1# de SEQUENCE(1) es A1
      if (!d || !d.ok) return A.E.ref;
      return { tipo: 'ref', h: d.h, f0: d.f0, c0: d.c0, f1: d.f1, c1: d.c1 };
    },
    // @A1:A9 de Excel 365 (_xlfn.SINGLE): la intersección implícita de antes
    SINGLE: (a, c, A) => { const v = A.evaluarNodo(a[0], c); return A.esMat(v) ? v.filas[0][0] : A.escalar(v, c); },
    NPV: (a, c, A) => { const r = A.num(a[0], c); const l = numeros(a.slice(1), c, A); let s = 0; l.forEach((x, i) => { s += x / Math.pow(1 + r, i + 1); }); return s; },
  };
  function desviacion(l, ddof, A) { const v = varianza(l, ddof, A); return esError(v) ? v : Math.sqrt(v); }
  function varianza(l, ddof, A) { if (l.length - ddof <= 0) return A.E.div; const m = sumar(l) / l.length; let s = 0; for (const x of l) s += (x - m) * (x - m); return s / (l.length - ddof); }
  function rango(a, c, A) { const x = A.num(a[0], c), l = numeros([a[1]], c, A), asc = A.opcional(a[2], c, 0, A.num) !== 0; if (!l.includes(x)) return A.E.na; let n = 1; for (const y of l) if (asc ? y < x : y > x) n++; return n; }
  function textoAntesDespues(a, c, A, antes) {
    const t = A.txt(a[0], c);
    const delims = []; A.recorrerValor(A.evaluarNodo(a[1], c), (v) => { if (esError(v)) A.lanzar(v); delims.push(A.aTexto(v)); });
    const inst = Math.trunc(A.opcional(a[2], c, 1, A.num)), sinMay = A.opcional(a[3], c, 0, A.num) === 1, finCuenta = A.opcional(a[4], c, 0, A.num) === 1;
    if (inst === 0 || Math.abs(inst) > t.length + 1) return A.E.valor;
    const base = sinMay ? t.toLowerCase() : t, ds = delims.map((d) => (sinMay ? d.toLowerCase() : d));
    // todas las apariciones de cualquiera de los delimitadores, sin solaparse, de izquierda a derecha
    const hallados = [];
    for (let i = 0; i <= base.length;) {
      let largo = -1;
      for (const d of ds) if (base.startsWith(d, i) && d.length > largo) largo = d.length;
      if (largo > 0) { hallados.push([i, i + largo]); i += largo; } else if (largo === 0) { hallados.push([i, i]); i++; } else i++;
    }
    if (finCuenta) { if (inst > 0) hallados.push([t.length, t.length]); else hallados.unshift([0, 0]); }   // el final (o el principio, buscando desde atrás) cuenta como delimitador
    const k = inst > 0 ? inst - 1 : hallados.length + inst;
    if (k < 0 || k >= hallados.length) return a[5] !== undefined && a[5] !== null ? A.val(a[5], c) : A.E.na;
    const [i0, i1] = hallados[k];
    return antes ? t.slice(0, i0) : t.slice(i1);
  }

  // días libres de NETWORKDAYS.INTL/WORKDAY.INTL indexados por día de la semana (0 = domingo); null si el código no vale
  function finDeSemana(codigo, A) {
    const libre = [false, false, false, false, false, false, false];
    if (typeof codigo === 'string') {
      if (!/^[01]{7}$/.test(codigo)) return null;
      for (let k = 0; k < 7; k++) libre[(k + 1) % 7] = codigo[k] === '1';
      return libre;
    }
    const n = Math.trunc(A.aNumero(codigo));
    if (n >= 1 && n <= 7) { libre[(n + 5) % 7] = true; libre[(n + 6) % 7] = true; return libre; }   // 1: sáb y dom, 2: dom y lun … 7: vie y sáb
    if (n >= 11 && n <= 17) { libre[(n - 11) % 7] = true; return libre; }                         // 11: domingo … 17: sábado
    return null;
  }

  function partesNum(a, c, A) { const x = A.num(a[0], c); if (x < 0) A.lanzar(A.E.num); return partes(x); }
  // XLOOKUP / XMATCH: modo 0 exacto, -1 exacto o el menor siguiente, 1 exacto o el mayor siguiente, 2 comodines; búsqueda 1 adelante, -1 atrás
  function xbuscar(lista, x, modo, direccion, A) {
    const orden = direccion === -1 ? [...lista.keys()].reverse() : [...lista.keys()];
    if (modo === 2) { const cr = criterio(x, A); for (const i of orden) if (cr(lista[i])) return i; return -1; }
    let mejor = -1;
    for (const i of orden) {
      const v = lista[i];
      if (v === null) continue;
      const d = A.comparar(v, x);
      if (d === 0 && typeof v === typeof x) return i;
      if (modo === -1 && d < 0 && (mejor < 0 || A.comparar(v, lista[mejor]) > 0)) mejor = i;
      if (modo === 1 && d > 0 && (mejor < 0 || A.comparar(v, lista[mejor]) < 0)) mejor = i;
    }
    return modo === 0 ? -1 : mejor;
  }

  // lo que usan los módulos que agregan funciones (calculo-funciones.js): CalculoHojas.ampliar((ayuda) => ({ NOMBRE: (args, ctx, A) => … }))
  const AYUDA = { numeros, sumar, criterio, filtrarPares, varianza, desviacion, percentil, redondear, truncar, arriba, coincide, columna, esError, error, textoFormato, nombreFuncion, Lanzado };
  const ampliar = (fabrica) => Object.assign(FUNCIONES, fabrica(AYUDA));
  globalThis.CalculoHojas = { crear, ErrorCalc, error, esError, get SOPORTADAS() { return Object.keys(FUNCIONES).sort(); }, ampliar, formatoGeneral, serialDe: (a, m, d, en1904) => { const v = fecha1904; fecha1904 = !!en1904; const s = serialDe(a, m, d); fecha1904 = v; return s; } };
})();

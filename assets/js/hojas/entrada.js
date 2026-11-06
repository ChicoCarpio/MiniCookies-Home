/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Lo que se escribe en una celda de Mini Cookies Hojas → el valor que se guarda (fase 2 del PLAN). Es lo contrario de xlsx-formato.js:
// "1.234,5" (es) o "1,234.5" (en) → 1234.5; "25/10/2026" → la fecha como número de serie con formato de fecha; "15%" → 0,15 con formato
// de porcentaje; "=SUMA(A1;A2)" → la fórmula como la guarda Excel ("SUM(A1,A2)"); "'007" → el texto "007". Sin DOM.
//
//   EntradaCelda.interpretar(texto, { idioma: 'es'|'en', fecha1904 }) →
//     { tipo: 'vacio' } | { tipo: 'numero', valor, formato? } | { tipo: 'texto', valor } | { tipo: 'logico', valor } | { tipo: 'error', valor: '#N/A' }
//     | { tipo: 'formula', formula }   (formula: en inglés, con comas, sin "=")
//   EntradaCelda.formulaAlmacenada('SUMA(A1;B1)', 'es') → 'SUM(A1,B1)'   (también acepta la forma inglesa)
//   EntradaCelda.formulaVisible('SUM(A1,B1)', 'es') → 'SUMA(A1;B1)'      (la forma en que la ve quien escribe en español)
//   EntradaCelda.textoParaEditar(celda, opciones) → el texto que aparece al editar una celda (fórmula con "=", número sin formato)
(function () {
  // Nombres en español (los de Excel en español) de las funciones más usadas → el nombre que se guarda
  const ES = {
    SUMA: 'SUM', PROMEDIO: 'AVERAGE', 'PROMEDIOA': 'AVERAGEA', MIN: 'MIN', MAX: 'MAX', CONTAR: 'COUNT', CONTARA: 'COUNTA', 'CONTAR.BLANCO': 'COUNTBLANK', PRODUCTO: 'PRODUCT',
    REDONDEAR: 'ROUND', 'REDONDEAR.MAS': 'ROUNDUP', 'REDONDEAR.MENOS': 'ROUNDDOWN', 'REDOND.MULT': 'MROUND', ENTERO: 'INT', TRUNCAR: 'TRUNC', ABS: 'ABS', RESIDUO: 'MOD', POTENCIA: 'POWER', RAIZ: 'SQRT',
    EXP: 'EXP', LN: 'LN', LOG: 'LOG', LOG10: 'LOG10', PI: 'PI', SIGNO: 'SIGN', 'MULTIPLO.SUPERIOR': 'CEILING', 'MULTIPLO.INFERIOR': 'FLOOR', 'MULTIPLO.SUPERIOR.MAT': 'CEILING.MATH', 'MULTIPLO.INFERIOR.MAT': 'FLOOR.MATH',
    ALEATORIO: 'RAND', 'ALEATORIO.ENTRE': 'RANDBETWEEN', FACT: 'FACT', COCIENTE: 'QUOTIENT', 'M.C.D': 'GCD', 'M.C.M': 'LCM', 'SUMA.CUADRADOS': 'SUMSQ', SUMAPRODUCTO: 'SUMPRODUCT', SUBTOTALES: 'SUBTOTAL',
    'REDONDEA.PAR': 'EVEN', 'REDONDEA.IMPAR': 'ODD', PAR: 'EVEN', IMPAR: 'ODD', SENO: 'SIN', COS: 'COS', TAN: 'TAN', GRADOS: 'DEGREES', RADIANES: 'RADIANS',
    MEDIANA: 'MEDIAN', MODA: 'MODE', DESVEST: 'STDEV', 'DESVEST.M': 'STDEV.S', DESVESTP: 'STDEVP', 'DESVEST.P': 'STDEV.P', VAR: 'VAR', 'VAR.S': 'VAR.S', VARP: 'VARP', 'VAR.P': 'VAR.P',
    'K.ESIMO.MAYOR': 'LARGE', 'K.ESIMO.MENOR': 'SMALL', JERARQUIA: 'RANK', 'JERARQUIA.EQV': 'RANK.EQ',
    SI: 'IF', 'SI.CONJUNTO': 'IFS', 'SI.ERROR': 'IFERROR', 'SI.ND': 'IFNA', Y: 'AND', O: 'OR', XO: 'XOR', NO: 'NOT', VERDADERO: 'TRUE', FALSO: 'FALSE', CAMBIAR: 'SWITCH',
    ESBLANCO: 'ISBLANK', ESNUMERO: 'ISNUMBER', ESTEXTO: 'ISTEXT', ESNOTEXTO: 'ISNONTEXT', ESLOGICO: 'ISLOGICAL', ESERROR: 'ISERROR', ESERR: 'ISERR', ESNOD: 'ISNA', 'ES.PAR': 'ISEVEN', 'ES.IMPAR': 'ISODD', ESREF: 'ISREF', ESFORMULA: 'ISFORMULA', NOD: 'NA', N: 'N', T: 'T', TIPO: 'TYPE', 'TIPO.DE.ERROR': 'ERROR.TYPE',
    'CONTAR.SI': 'COUNTIF', 'CONTAR.SI.CONJUNTO': 'COUNTIFS', 'SUMAR.SI': 'SUMIF', 'SUMAR.SI.CONJUNTO': 'SUMIFS', 'PROMEDIO.SI': 'AVERAGEIF', 'PROMEDIO.SI.CONJUNTO': 'AVERAGEIFS', 'MAX.SI.CONJUNTO': 'MAXIFS', 'MIN.SI.CONJUNTO': 'MINIFS',
    BUSCARV: 'VLOOKUP', BUSCARH: 'HLOOKUP', BUSCAR: 'LOOKUP', COINCIDIR: 'MATCH', 'COINCIDIRX': 'XMATCH', BUSCARX: 'XLOOKUP', INDICE: 'INDEX', ELEGIR: 'CHOOSE', FILA: 'ROW', COLUMNA: 'COLUMN', FILAS: 'ROWS', COLUMNAS: 'COLUMNS', DESREF: 'OFFSET', INDIRECTO: 'INDIRECT',
    LARGO: 'LEN', IZQUIERDA: 'LEFT', DERECHA: 'RIGHT', EXTRAE: 'MID', MAYUSC: 'UPPER', MINUSC: 'LOWER', NOMPROPIO: 'PROPER', ESPACIOS: 'TRIM', LIMPIAR: 'CLEAN', CONCATENAR: 'CONCATENATE', CONCAT: 'CONCAT', UNIRCADENAS: 'TEXTJOIN',
    REPETIR: 'REPT', SUSTITUIR: 'SUBSTITUTE', REEMPLAZAR: 'REPLACE', ENCONTRAR: 'FIND', HALLAR: 'SEARCH', IGUAL: 'EXACT', VALOR: 'VALUE', 'VALOR.NUMERO': 'NUMBERVALUE', TEXTO: 'TEXT', DECIMAL: 'FIXED', 'DECIMAL.FIJO': 'FIXED', CARACTER: 'CHAR', UNICAR: 'UNICHAR', CODIGO: 'CODE', UNICODE: 'UNICODE',
    FECHA: 'DATE', NSHORA: 'TIME', 'AÑO': 'YEAR', ANO: 'YEAR', MES: 'MONTH', DIA: 'DAY', HORA: 'HOUR', MINUTO: 'MINUTE', SEGUNDO: 'SECOND', DIASEM: 'WEEKDAY', HOY: 'TODAY', AHORA: 'NOW', 'FECHA.MES': 'EDATE', 'FIN.MES': 'EOMONTH', DIAS: 'DAYS', SIFECHA: 'DATEDIF',
    'DIAS.LAB': 'NETWORKDAYS', 'DIA.LAB': 'WORKDAY', 'DIAS.LAB.INTL': 'NETWORKDAYS.INTL', 'DIA.LAB.INTL': 'WORKDAY.INTL', 'NUM.DE.SEMANA': 'WEEKNUM', 'FRAC.AÑO': 'YEARFRAC', 'FRAC.ANO': 'YEARFRAC', FECHANUMERO: 'DATEVALUE',
    PAGO: 'PMT', VF: 'FV', VA: 'PV', VNA: 'NPV', AGREGAR: 'AGGREGATE', HIPERVINCULO: 'HYPERLINK', TEXTOANTES: 'TEXTBEFORE', TEXTODESPUES: 'TEXTAFTER', LET: 'LET',
    // fase 3.1
    ASENO: 'ASIN', ACOS: 'ACOS', ATAN: 'ATAN', ATAN2: 'ATAN2', SENOH: 'SINH', COSH: 'COSH', TANH: 'TANH', ASENOH: 'ASINH', ACOSH: 'ACOSH', ATANH: 'ATANH', COT: 'COT', COTH: 'COTH', CSC: 'CSC', CSCH: 'CSCH', SEC: 'SEC', SECH: 'SECH', ACOT: 'ACOT', ACOTH: 'ACOTH', RAIZ2PI: 'SQRTPI',
    'FACT.DOBLE': 'FACTDOUBLE', COMBINAT: 'COMBIN', COMBINA: 'COMBINA', PERMUTACIONES: 'PERMUT', 'PERMUTACIONES.A': 'PERMUTATIONA', MULTINOMIAL: 'MULTINOMIAL', 'MULTIPLO.SUPERIOR.EXACTO': 'CEILING.PRECISE', 'MULTIPLO.INFERIOR.EXACTO': 'FLOOR.PRECISE', 'MULTIPLO.SUPERIOR.ISO': 'ISO.CEILING',
    'NUMERO.ROMANO': 'ROMAN', 'NUMERO.ARABE': 'ARABIC', BASE: 'BASE', 'CONV.DECIMAL': 'DECIMAL', 'DEC.A.BIN': 'DEC2BIN', 'DEC.A.OCT': 'DEC2OCT', 'DEC.A.HEX': 'DEC2HEX', 'BIN.A.DEC': 'BIN2DEC', 'OCT.A.DEC': 'OCT2DEC', 'HEX.A.DEC': 'HEX2DEC',
    'BIN.A.OCT': 'BIN2OCT', 'BIN.A.HEX': 'BIN2HEX', 'OCT.A.BIN': 'OCT2BIN', 'OCT.A.HEX': 'OCT2HEX', 'HEX.A.BIN': 'HEX2BIN', 'HEX.A.OCT': 'HEX2OCT', 'BIT.Y': 'BITAND', 'BIT.O': 'BITOR', 'BIT.XO': 'BITXOR', 'BIT.DESPLIZQDA': 'BITLSHIFT', 'BIT.DESPLDCHA': 'BITRSHIFT',
    DELTA: 'DELTA', 'MAYOR.O.IGUAL': 'GESTEP', SUMAX2MENOSY2: 'SUMX2MY2', SUMAX2MASY2: 'SUMX2PY2', SUMAXMENOSY2: 'SUMXMY2', 'SUMA.SERIES': 'SERIESSUM', MMULT: 'MMULT', MDETERM: 'MDETERM', MINVERSA: 'MINVERSE', 'M.UNIDAD': 'MUNIT', TRANSPONER: 'TRANSPOSE',
    MAXA: 'MAXA', MINA: 'MINA', DESVESTA: 'STDEVA', DESVESTPA: 'STDEVPA', VARA: 'VARA', VARPA: 'VARPA', DESVPROM: 'AVEDEV', DESVIA2: 'DEVSQ', 'MEDIA.GEOM': 'GEOMEAN', 'MEDIA.ARMO': 'HARMEAN', CURTOSIS: 'KURT', 'COEFICIENTE.ASIMETRIA': 'SKEW', 'COEFICIENTE.ASIMETRIA.P': 'SKEW.P', 'MEDIA.ACOTADA': 'TRIMMEAN',
    'COEF.DE.CORREL': 'CORREL', PEARSON: 'PEARSON', 'COEFICIENTE.R2': 'RSQ', COVAR: 'COVAR', 'COVARIANCE.P': 'COVARIANCE.P', 'COVARIANZA.M': 'COVARIANCE.S', PENDIENTE: 'SLOPE', 'INTERSECCION.EJE': 'INTERCEPT', PRONOSTICO: 'FORECAST', 'PRONOSTICO.LINEAL': 'FORECAST.LINEAR', 'ERROR.TIPICO.XY': 'STEYX', NORMALIZACION: 'STANDARDIZE',
    'MODA.UNO': 'MODE.SNGL', 'MODA.VARIOS': 'MODE.MULT', 'JERARQUIA.MEDIA': 'RANK.AVG', 'RANGO.PERCENTIL': 'PERCENTRANK', 'RANGO.PERCENTIL.INC': 'PERCENTRANK.INC', 'RANGO.PERCENTIL.EXC': 'PERCENTRANK.EXC', FRECUENCIA: 'FREQUENCY',
    'DISTR.NORM.N': 'NORM.DIST', 'DISTR.NORM': 'NORMDIST', 'DISTR.NORM.ESTAND.N': 'NORM.S.DIST', 'DISTR.NORM.ESTAND': 'NORMSDIST', 'INV.NORM': 'NORM.INV', 'DISTR.NORM.INV': 'NORMINV', 'INV.NORM.ESTAND': 'NORM.S.INV', 'DISTR.NORM.ESTAND.INV': 'NORMSINV', FI: 'PHI', GAUSS: 'GAUSS',
    'DISTR.LOGNORM': 'LOGNORM.DIST', 'DISTR.LOG.NORM': 'LOGNORMDIST', 'INV.LOGNORM': 'LOGNORM.INV', 'DISTR.LOG.INV': 'LOGINV', 'DISTR.EXP.N': 'EXPON.DIST', 'DISTR.EXP': 'EXPONDIST', 'POISSON.DIST': 'POISSON.DIST', POISSON: 'POISSON',
    'DISTR.BINOM.N': 'BINOM.DIST', 'DISTR.BINOM': 'BINOMDIST', 'INV.BINOM': 'BINOM.INV', 'BINOM.CRIT': 'CRITBINOM', 'NEGBINOM.DIST': 'NEGBINOM.DIST', NEGBINOMDIST: 'NEGBINOMDIST', 'DISTR.HIPERGEOM.N': 'HYPGEOM.DIST', 'DISTR.HIPERGEOM': 'HYPGEOMDIST',
    'DISTR.WEIBULL': 'WEIBULL.DIST', 'DIST.WEIBULL': 'WEIBULL', 'DISTR.GAMMA.N': 'GAMMA.DIST', 'DISTR.GAMMA': 'GAMMADIST', 'INV.GAMMA': 'GAMMA.INV', 'DISTR.GAMMA.INV': 'GAMMAINV', GAMMA: 'GAMMA', 'GAMMA.LN': 'GAMMALN', 'GAMMA.LN.EXACTO': 'GAMMALN.PRECISE',
    'DISTR.BETA.N': 'BETA.DIST', 'DISTR.BETA': 'BETADIST', 'INV.BETA.N': 'BETA.INV', 'DISTR.BETA.INV': 'BETAINV', 'DISTR.T.N': 'T.DIST', 'DISTR.T.2C': 'T.DIST.2T', 'DISTR.T.CD': 'T.DIST.RT', 'DISTR.T': 'TDIST', 'INV.T': 'T.INV', 'INV.T.2C': 'T.INV.2T', 'DISTR.T.INV': 'TINV',
    'DISTR.CHICUAD': 'CHISQ.DIST', 'DISTR.CHICUAD.CD': 'CHISQ.DIST.RT', 'DISTR.CHI': 'CHIDIST', 'INV.CHICUAD': 'CHISQ.INV', 'INV.CHICUAD.CD': 'CHISQ.INV.RT', 'PRUEBA.CHI.INV': 'CHIINV', 'DISTR.F.N': 'F.DIST', 'DISTR.F.CD': 'F.DIST.RT', 'DISTR.F': 'FDIST', 'INV.F': 'F.INV', 'INV.F.CD': 'F.INV.RT', 'DISTR.F.INV': 'FINV',
    FISHER: 'FISHER', 'PRUEBA.FISHER.INV': 'FISHERINV', 'INTERVALO.CONFIANZA': 'CONFIDENCE', 'INTERVALO.CONFIANZA.NORM': 'CONFIDENCE.NORM', 'INTERVALO.CONFIANZA.T': 'CONFIDENCE.T', 'FUN.ERROR': 'ERF', 'FUN.ERROR.EXACTO': 'ERF.PRECISE', 'FUN.ERROR.COMPL': 'ERFC', 'FUN.ERROR.COMPL.EXACTO': 'ERFC.PRECISE',
    'PRUEBA.Z.N': 'Z.TEST', 'PRUEBA.Z': 'ZTEST', 'PRUEBA.T.N': 'T.TEST', 'PRUEBA.T': 'TTEST', 'PRUEBA.F.N': 'F.TEST', 'PRUEBA.F': 'FTEST', 'PRUEBA.CHICUAD': 'CHISQ.TEST', 'PRUEBA.CHI': 'CHITEST',
    NPER: 'NPER', TASA: 'RATE', TIR: 'IRR', TIRM: 'MIRR', 'VNA.NO.PER': 'XNPV', 'TIR.NO.PER': 'XIRR', PAGOINT: 'IPMT', PAGOPRIN: 'PPMT', 'PAGO.INT.ENTRE': 'CUMIPMT', 'PAGO.PRINC.ENTRE': 'CUMPRINC', 'INT.PAGO.DIR': 'ISPMT', SLN: 'SLN', SYD: 'SYD', DDB: 'DDB', DB: 'DB',
    'INT.EFECTIVO': 'EFFECT', 'TASA.NOMINAL': 'NOMINAL', 'P.DURACION': 'PDURATION', RRI: 'RRI', 'VF.PLAN': 'FVSCHEDULE', 'MONEDA.DEC': 'DOLLARDE', 'MONEDA.FRAC': 'DOLLARFR',
    MONEDA: 'DOLLAR', HORANUMERO: 'TIMEVALUE', LARGOB: 'LENB', IZQUIERDAB: 'LEFTB', DERECHAB: 'RIGHTB', EXTRAEB: 'MIDB', DIVIDIRTEXTO: 'TEXTSPLIT', VALORATEXTO: 'VALUETOTEXT', FORMULATEXTO: 'FORMULATEXT',
    DIAS360: 'DAYS360', 'ISO.NUM.DE.SEMANA': 'ISOWEEKNUM', DIRECCION: 'ADDRESS', AREAS: 'AREAS', HOJA: 'SHEET', HOJAS: 'SHEETS',
    FILTRAR: 'FILTER', ORDENAR: 'SORT', ORDENARPOR: 'SORTBY', UNICOS: 'UNIQUE', SECUENCIA: 'SEQUENCE', MATRIZALEAT: 'RANDARRAY', APILARV: 'VSTACK', APILARH: 'HSTACK', TOMAR: 'TAKE', EXCLUIR: 'DROP',
    ELEGIRCOLS: 'CHOOSECOLS', ELEGIRFILAS: 'CHOOSEROWS', ENCOL: 'TOCOL', ENFILA: 'TOROW', AJUSTARFILAS: 'WRAPROWS', AJUSTARCOLS: 'WRAPCOLS', EXPANDIR: 'EXPAND',
    BDSUMA: 'DSUM', BDCONTAR: 'DCOUNT', BDCONTARA: 'DCOUNTA', BDPROMEDIO: 'DAVERAGE', BDMAX: 'DMAX', BDMIN: 'DMIN', BDPRODUCTO: 'DPRODUCT', BDEXTRAER: 'DGET', BDDESVEST: 'DSTDEV', BDDESVESTP: 'DSTDEVP', BDVAR: 'DVAR', BDVARP: 'DVARP',
    // medidos en Excel en español (2026-11-06); las de la primera línea, el motor no las calcula: se muestran con su nombre y conservan su valor
    CONVERTIR: 'CONVERT', CRECIMIENTO: 'GROWTH', 'ESTIMACION.LINEAL': 'LINEST', 'ESTIMACION.LOGARITMICA': 'LOGEST', TENDENCIA: 'TREND',
    REGEXEXTRACCION: 'REGEXEXTRACT', REGEXREEMPLAZAR: 'REGEXREPLACE', REGEXPRUEBA: 'REGEXTEST',
    PERCENTIL: 'PERCENTILE', 'PERCENTIL.INC': 'PERCENTILE.INC', 'PERCENTIL.EXC': 'PERCENTILE.EXC', CUARTIL: 'QUARTILE', 'CUARTIL.INC': 'QUARTILE.INC', 'CUARTIL.EXC': 'QUARTILE.EXC',
  };
  const EN = {}; for (const k in ES) if (!(ES[k] in EN)) EN[ES[k]] = k;   // el primero de cada una (AÑO antes que ANO)
  // En inglés también se aceptan los nombres en español, salvo el que en inglés es otra función: DECIMAL (en español, FIXED)
  const ES_EN = Object.assign({}, ES); for (const k in ES) if (k in EN && ES[k] !== k) delete ES_EN[k];

  // recorre la fórmula fuera de los textos "…" y de los nombres de hoja '…': cb(trozo de código) → reemplazo
  function fueraDeTextos(t, cb) {
    let salida = '', i = 0, codigo = '';
    const soltar = () => { if (codigo) { salida += cb(codigo); codigo = ''; } };
    while (i < t.length) {
      const ch = t[i];
      if (ch === '"' || ch === "'") {
        soltar();
        let j = i + 1;
        while (j < t.length) { if (t[j] === ch) { if (t[j + 1] === ch) j += 2; else break; } else j++; }
        salida += t.slice(i, j + 1); i = j + 1; continue;
      }
      codigo += ch; i++;
    }
    soltar();
    return salida;
  }
  // nombres de función seguidos de "(" según un diccionario
  // El "(" se mira sin consumirlo: si no, en SUM(FILTER( el "(" ya lo gastó SUM y FILTER no se encuentra (quedaba sin traducir ni prefijo).
  const RE_FUNCION = /(^|[^A-Za-z0-9_.À-￿$])([A-Za-zÀ-￿][A-Za-z0-9_.À-￿]*)(\s*)(?=\()/g;
  const traducirNombres = (codigo, dic) => codigo.replace(RE_FUNCION, (m, antes, nombre, paren) => { const u = nombre.toUpperCase(); return antes + (dic[u] || nombre) + paren; });

  // VERDADERO y FALSO sin paréntesis (como los escribe Excel en español) ↔ TRUE y FALSE; un nombre seguido de «(» ya lo traduce traducirNombres
  const RE_LOGICO = (a, b) => new RegExp('(^|[^A-Za-z0-9_.!$À-￿\'\\]])(' + a + '|' + b + ')(?![A-Za-z0-9_.(!\\[À-￿])', 'gi');
  const RE_LOGICO_ES = RE_LOGICO('VERDADERO', 'FALSO'), RE_LOGICO_EN = RE_LOGICO('TRUE', 'FALSE');
  const logicosEn = (c) => c.replace(RE_LOGICO_ES, (m, antes, n) => antes + (n.toUpperCase() === 'FALSO' ? 'FALSE' : 'TRUE'));
  const logicosEs = (c) => c.replace(RE_LOGICO_EN, (m, antes, n) => antes + (n.toUpperCase() === 'FALSE' ? 'FALSO' : 'VERDADERO'));

  // Las funciones posteriores a Excel 2007 se guardan con prefijo (_xlfn.XLOOKUP): sin él, Excel no las reconoce al abrir el archivo
  let futuras = null;
  const SIN_ZONA = new Set(['FILTER', 'SORT']);   // estas llevan _xlfn._xlws.
  function prefijos() {
    if (futuras) return futuras;
    futuras = new Set();
    const tabla = (typeof XlsFormulas !== 'undefined' && XlsFormulas.FUNCIONES) || [];
    for (const d of tabla) if (d && d.tipo === 'futura' && d.prefijo) futuras.add(String(d.nombre).toUpperCase().replace(/^_XLFN\.(_XLWS\.)?/, ''));
    for (const n of ['LET', 'LAMBDA', 'FORMULATEXT', 'SHEET', 'SHEETS']) futuras.add(n);   // no están en la tabla medida
    return futuras;
  }
  // Los nombres que declara LET se guardan con _xlpm. (Excel: _xlfn.LET(_xlpm.x,2,_xlpm.x*3)); sin él, Excel da #¿NOMBRE?.
  // Se aplica a la fórmula entera (textos incluidos en el recorrido, para saltarlos) después de traducir y poner los prefijos.
  function conParametros(f) {
    const re = /(^|[^A-Za-z0-9_.])(?:_xlfn\.)?LET\(/gi;
    let m, salida = f;
    const llamadas = [];
    while ((m = re.exec(f))) llamadas.push(m.index + m[0].length);   // posición justo después de "LET("
    for (let k = llamadas.length - 1; k >= 0; k--) {   // de la última a la primera: las posiciones anteriores no se mueven
      const ini = llamadas[k];
      const args = [];
      let nivel = 0, j = ini, desde = ini;
      for (; j < salida.length; j++) {
        const ch = salida[j];
        if (ch === '"' || ch === "'") { j++; while (j < salida.length && !(salida[j] === ch && salida[j + 1] !== ch)) j += salida[j] === ch ? 2 : 1; continue; }
        if (ch === '(' || ch === '{') nivel++;
        else if (ch === ')' || ch === '}') { if (nivel === 0) break; nivel--; }
        else if (ch === ',' && nivel === 0) { args.push([desde, j]); desde = j + 1; }
      }
      args.push([desde, j]);
      const nombres = [];
      for (let i = 0; i + 1 < args.length; i += 2) { const n = salida.slice(args[i][0], args[i][1]).trim(); if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(n)) nombres.push(n); }
      if (!nombres.length) continue;
      const reNombre = new RegExp('(^|[^A-Za-z0-9_.$!\'\\]])(' + nombres.join('|') + ')(?![A-Za-z0-9_.(!\\[])', 'gi');
      const cuerpo = fueraDeTextos(salida.slice(ini, j), (c) => c.replace(reNombre, (x, antes, n) => antes + '_xlpm.' + n));
      salida = salida.slice(0, ini) + cuerpo + salida.slice(j);
    }
    return salida;
  }
  const conPrefijo = (codigo) => codigo.replace(RE_FUNCION, (m, antes, nombre, paren) => {
    const u = nombre.toUpperCase();
    if (/^_XL/.test(u) || !prefijos().has(u)) return m;
    return antes + (SIN_ZONA.has(u) ? '_xlfn._xlws.' : '_xlfn.') + nombre + paren;
  });
  const sinPrefijo = (codigo) => codigo.replace(/_xlfn\.ANCHORARRAY\(([^()]+)\)/gi, '$1#').replace(/_xlfn\.(?:_xlws\.)?|_xlws\.|_xlpm\./gi, '');
  // A1# (todo lo que desborda la matriz dinámica de A1) se guarda como _xlfn.ANCHORARRAY(A1), como Excel
  const conAncla = (codigo) => codigo.replace(/(^|[^A-Za-z0-9_.!$])((?:[A-Za-z_][A-Za-z0-9_.]*!)?\$?[A-Za-z]{1,3}\$?\d+)#/g, (m, antes, ref) => antes + '_xlfn.ANCHORARRAY(' + ref + ')');

  // Separadores dentro de las matrices constantes {…}: en inglés «,» entre columnas y «;» entre filas; en español (punto y coma entre
  // argumentos, coma decimal) «\» entre columnas y «;» entre filas, como Excel en español. Marca cada separador de matriz de un trozo de
  // código con \u0002 (columna) o \u0003 (fila); `nivel` lleva la cuenta de las llaves abiertas entre trozos (los textos parten la fórmula).
  function marcarMatrices(c, estado, deColumna) {
    let r = '';
    for (const ch of c) {
      if (ch === '{') estado.nivel++;
      else if (ch === '}') estado.nivel = Math.max(0, estado.nivel - 1);
      if (estado.nivel > 0 && ch === deColumna) { r += '\u0002'; continue; }
      if (estado.nivel > 0 && ch === ';') { r += '\u0003'; continue; }
      r += ch;
    }
    return r;
  }
  // los elementos especiales de las referencias a tablas tienen nombre propio en Excel en español
  const ESPECIALES = [['#Todo', '#All'], ['#Datos', '#Data'], ['#Encabezados', '#Headers'], ['#Totales', '#Totals'], ['#Esta fila', '#This Row']];
  const especialesEn = (c) => (c.indexOf('#') < 0 ? c : ESPECIALES.reduce((x, [es, en]) => x.replace(new RegExp(es.replace(' ', '\\s'), 'gi'), en), c));
  const especialesEs = (c) => (c.indexOf('#') < 0 ? c : ESPECIALES.reduce((x, [es, en]) => x.replace(new RegExp(en.replace(' ', '\\s'), 'gi'), es), c));
  // dónde termina el operando de @: un nombre o referencia (A1:A3, A:A, 1:1, Hoja!A1, T1[Col]) y, si sigue un paréntesis, hasta cerrarlo
  function finOperando(f, ini) {
    let j = ini;
    const m = /^(?:'(?:[^']|'')+'!|[A-Za-z_À-￿][\w.À-￿]*!)?(?:\$?\d+:\$?\d+|[$A-Za-z_À-￿][$\w.À-￿]*(?::\$?[A-Za-z]+\$?\d*)?)/.exec(f.slice(j));
    if (m) j += m[0].length;
    if (f[j] === '[') {
      let nivel = 0;
      for (; j < f.length; j++) {
        if (f[j] === "'") { j++; continue; }
        if (f[j] === '[') nivel++;
        else if (f[j] === ']' && --nivel === 0) { j++; break; }
      }
    } else if (f[j] === '(') {
      let nivel = 0;
      for (; j < f.length; j++) {
        if (f[j] === '"') { j++; while (j < f.length && !(f[j] === '"' && f[j + 1] !== '"')) j += f[j] === '"' ? 2 : 1; continue; }
        if (f[j] === '(') nivel++;
        else if (f[j] === ')' && --nivel === 0) { j++; break; }
      }
    }
    return j;
  }
  // _xlfn.SINGLE(…) → @…, también con paréntesis o comillas dentro; si lo de adentro no es un solo operando (SINGLE(A1+1)) queda como función
  function sinInterseccion(f) {
    const PRE = '_xlfn.single(';
    if (f.length > 8192 || f.toLowerCase().indexOf(PRE) < 0) return f;   // más larga que el límite de Excel: viene de un archivo raro y no se recorre
    let salida = '', i = 0;
    while (i < f.length) {
      const ch = f[i];
      if (ch === '"' || ch === "'") { let j = i + 1; while (j < f.length && !(f[j] === ch && f[j + 1] !== ch)) j += f[j] === ch ? 2 : 1; salida += f.slice(i, j + 1); i = j + 1; continue; }
      if (f.slice(i, i + PRE.length).toLowerCase() !== PRE) { salida += ch; i++; continue; }
      const fin = finOperando(f, i + PRE.length - 1);   // desde el paréntesis: hasta el que lo cierra
      if (f[fin - 1] !== ')') { salida += f.slice(i); break; }   // sin cerrar: fórmula mala, se muestra como está
      const dentro = f.slice(i + PRE.length, fin - 1);
      if (dentro && finOperando(dentro, 0) === dentro.length) salida += '@' + sinInterseccion(dentro);
      else salida += f.slice(i, i + PRE.length) + sinInterseccion(dentro) + ')';
      i = fin;
    }
    return salida;
  }
  // @A1:A3 o @F(x) (intersección implícita) se guarda como _xlfn.SINGLE(…), como Excel. La @ de una referencia a tabla ([@Col]) no se toca.
  function conInterseccion(f) {
    if (f.indexOf('@') < 0) return f;
    let salida = '', i = 0, corchetes = 0;
    while (i < f.length) {
      const ch = f[i];
      if (ch === '"' || ch === "'") { let j = i + 1; while (j < f.length && !(f[j] === ch && f[j + 1] !== ch)) j += f[j] === ch ? 2 : 1; salida += f.slice(i, j + 1); i = j + 1; continue; }
      if (ch === '[') corchetes++; else if (ch === ']') corchetes--;
      if (ch !== '@' || corchetes > 0) { salida += ch; i++; continue; }
      const j = finOperando(f, i + 1);
      if (j === i + 1) { salida += ch; i++; continue; }
      salida += '_xlfn.SINGLE(' + conInterseccion(f.slice(i + 1, j)) + ')';
      i = j;
    }
    return salida;
  }
  // En español la coma entre cifras es decimal, también sin ningún «;» (=A1*1,13): lo demás con comas (SUM(1,2,3), ROUND(A1,2), {1,2}) es la forma inglesa
  const decimalesEs = (c) => c.replace(/\{[^{}]*\}|(^|[^A-Za-z0-9_.$,!:])(\d+),(\d+)(?![\d,A-Za-z_.(:!])/g, (m, antes, a, b) => (a === undefined ? m : antes + a + '.' + b));
  // forma que escribe quien usa el editor → forma de Excel en el archivo
  function formulaAlmacenada(texto, idioma) {
    let f = conInterseccion(String(texto).replace(/^=/, ''));
    // con la página en inglés se escribe como en Excel en inglés; en español se aceptan las dos formas (con «;» o «\» es la española)
    const usaPuntoYComa = idioma !== 'en' && fueraDeTextos(f, (c) => (/[;\\]/.test(c) ? '\u0001' : '')).includes('\u0001');
    const estado = { nivel: 0 };
    f = fueraDeTextos(f, (c) => {
      let x = idioma === 'en' ? traducirNombres(c, ES_EN) : traducirNombres(logicosEn(especialesEn(c)), ES);
      if (usaPuntoYComa) x = marcarMatrices(x, estado, '\\').replace(/(\d),(\d)/g, '$1.$2').replace(/;/g, ',').replace(/\u0002/g, ',').replace(/\u0003/g, ';');   // en español: coma decimal y punto y coma entre argumentos
      else if (idioma !== 'en') x = decimalesEs(x);
      return conPrefijo(conAncla(x));
    });
    return /LET\(/i.test(f) ? conParametros(f) : f;
  }
  // forma del archivo → forma para mostrar y editar en el idioma de la página
  function formulaVisible(formula, idioma) {
    const f = sinInterseccion(String(formula).replace(/^=/, ''));
    if (idioma !== 'es') return fueraDeTextos(f, sinPrefijo);
    const estado = { nivel: 0 };
    return fueraDeTextos(f, (c) => marcarMatrices(logicosEs(especialesEs(traducirNombres(sinPrefijo(c), EN))), estado, ',').replace(/,/g, ';').replace(/(\d)\.(\d)/g, '$1,$2').replace(/\u0002/g, '\\').replace(/\u0003/g, ';'));
  }

  const ERRORES = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#¡NULO!', '#¡DIV/0!', '#¡VALOR!', '#¡REF!', '#¿NOMBRE?', '#¡NUM!', '#N/D'];
  const ERROR_ES = { '#¡NULO!': '#NULL!', '#¡DIV/0!': '#DIV/0!', '#¡VALOR!': '#VALUE!', '#¡REF!': '#REF!', '#¿NOMBRE?': '#NAME?', '#¡NUM!': '#NUM!', '#N/D': '#N/A' };

  function serial(a, m, d, fecha1904) {
    if (!(m >= 1 && m <= 12 && d >= 1 && d <= 31 && a >= 1900 && a <= 9999)) return null;
    const ms = Date.UTC(a, m - 1, d);
    const x = new Date(ms);
    if (x.getUTCMonth() !== m - 1) return null;   // 31/02
    if (fecha1904) return Math.round((ms - Date.UTC(1904, 0, 1)) / 86400000);
    let s = Math.round((ms - Date.UTC(1899, 11, 30)) / 86400000);
    if (s <= 60) s -= 1;
    return s;
  }

  function interpretar(texto, opciones) {
    const o = opciones || {};
    const es = o.idioma !== 'en';
    const t = String(texto === undefined || texto === null ? '' : texto);
    if (t === '') return { tipo: 'vacio' };
    if (t[0] === "'") return { tipo: 'texto', valor: t.slice(1) };
    if (t[0] === '=' && t.length > 1) return { tipo: 'formula', formula: formulaAlmacenada(t, o.idioma) };
    const s = t.trim();
    const u = s.toUpperCase();
    if (u === 'TRUE' || u === 'VERDADERO') return { tipo: 'logico', valor: true };
    if (u === 'FALSE' || u === 'FALSO') return { tipo: 'logico', valor: false };
    if (ERRORES.includes(u)) return { tipo: 'error', valor: ERROR_ES[u] || u };
    const dec = es ? ',' : '.', mil = es ? '.' : ',';
    // número con signo, miles, decimales, %, moneda y notación científica
    const m = new RegExp('^([-+]?)\\s*([$€₡£¥]?)\\s*(\\d{1,3}(?:\\' + mil + '\\d{3})+|\\d*)(?:\\' + dec + '(\\d*))?(?:[eE]([-+]?\\d+))?\\s*(%?)\\s*([$€₡£¥]?)$').exec(s);
    if (m && (m[3] !== '' || (m[4] || '') !== '') && !(m[2] && m[7])) {
      const entero = m[3].split(mil).join('');
      let v = parseFloat((entero || '0') + '.' + (m[4] || '0') + (m[5] ? 'e' + m[5] : ''));
      if (m[1] === '-') v = -v;
      const conMiles = m[3].includes(mil);
      const decimales = (m[4] || '').length;
      if (m[6]) return { tipo: 'numero', valor: v / 100, formato: decimales ? '0.' + '0'.repeat(decimales) + '%' : '0%' };
      const moneda = m[2] || m[7];
      if (moneda) return { tipo: 'numero', valor: v, formato: (m[2] ? '"' + moneda + '"' : '') + '#,##0' + (decimales ? '.' + '0'.repeat(Math.max(2, decimales)) : '.00') + (m[7] ? ' "' + moneda + '"' : '') };
      if (conMiles) return { tipo: 'numero', valor: v, formato: '#,##0' + (decimales ? '.' + '0'.repeat(decimales) : '') };
      if (m[5]) return { tipo: 'numero', valor: v, formato: '0.00E+00' };
      return { tipo: 'numero', valor: v };
    }
    // también se acepta el punto decimal en español cuando no puede ser separador de miles ("3.5")
    if (es && /^[-+]?\d+\.\d{1,2}$|^[-+]?\d+\.\d{4,}$/.test(s)) return { tipo: 'numero', valor: parseFloat(s) };
    // fechas: d/m/aaaa (es) o m/d/aaaa (en), aaaa-mm-dd, con hora opcional
    let f = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s);
    let a, mes, d, h = null, mi = 0, se = 0;
    if (f) { a = +f[1]; mes = +f[2]; d = +f[3]; if (f[4] !== undefined) { h = +f[4]; mi = +f[5]; se = +(f[6] || 0); } }
    else if ((f = /^(\d{1,2})[/-](\d{1,2})[/-](\d{2}|\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s))) {
      a = +f[3]; if (f[3].length === 2) a += a < 30 ? 2000 : 1900;
      if (es) { d = +f[1]; mes = +f[2]; } else { mes = +f[1]; d = +f[2]; }
      if (f[4] !== undefined) { h = +f[4]; mi = +f[5]; se = +(f[6] || 0); }
    }
    if (a !== undefined) {
      const sd = serial(a, mes, d, o.fecha1904);
      if (sd !== null && (h === null || (h < 24 && mi < 60 && se < 60))) {
        if (h === null) return { tipo: 'numero', valor: sd, formato: es ? 'dd/mm/yyyy' : 'm/d/yyyy' };
        return { tipo: 'numero', valor: sd + (h * 3600 + mi * 60 + se) / 86400, formato: es ? 'dd/mm/yyyy h:mm' : 'm/d/yyyy h:mm' };
      }
    }
    // horas: 14:30, 14:30:05, 2:30 p. m.
    f = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([ap]\.?\s?m\.?)?$/i.exec(s);
    if (f) {
      let hh = +f[1]; const mm = +f[2], ss = +(f[3] || 0);
      if (f[4]) { const pm = /^p/i.test(f[4]); if (hh === 12) hh = pm ? 12 : 0; else if (pm) hh += 12; }
      if (hh < 24 && mm < 60 && ss < 60) return { tipo: 'numero', valor: (hh * 3600 + mm * 60 + ss) / 86400, formato: f[3] !== undefined ? 'h:mm:ss' : 'h:mm' };
    }
    return { tipo: 'texto', valor: t };
  }

  const ES_DE_ERROR = { '#NULL!': '#¡NULO!', '#DIV/0!': '#¡DIV/0!', '#VALUE!': '#¡VALOR!', '#REF!': '#¡REF!', '#NAME?': '#¿NOMBRE?', '#NUM!': '#¡NUM!', '#N/A': '#N/D' };
  // lo que aparece en el editor al empezar a editar una celda
  function textoParaEditar(celda, opciones) {
    const o = opciones || {};
    if (!celda) return '';
    if (celda.formula) return '=' + formulaVisible(celda.formula, o.idioma === 'en' ? 'en' : 'es');
    const v = celda.valor;
    if (v === null || v === undefined) return '';
    if (typeof v === 'boolean') return o.idioma === 'en' ? (v ? 'TRUE' : 'FALSE') : (v ? 'VERDADERO' : 'FALSO');
    if (typeof v === 'number') {
      if (o.esFecha) return o.textoFecha ? o.textoFecha(v) : String(v);
      if (o.esPorcentaje) return numeroTexto(v * 100, o) + '%';
      return numeroTexto(v, o);
    }
    if (celda.tipo === 'error') return o.idioma === 'en' ? v : (ES_DE_ERROR[v] || v);
    const s = String(v);
    return interpretar(s, o).tipo !== 'texto' ? "'" + s : s;   // un texto que parece número se edita con apóstrofo para que siga siendo texto
  }
  function numeroTexto(v, o) {
    let s = Number(v.toPrecision(15)).toString();
    if (/e/.test(s)) return s.replace('e', 'E');
    return o.idioma === 'en' ? s : s.replace('.', ',');
  }

  globalThis.EntradaCelda = { interpretar, formulaAlmacenada, formulaVisible, textoParaEditar, NOMBRES_ES: ES };
})();

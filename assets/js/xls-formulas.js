/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Fórmulas de Excel 97-2003 (BIFF8, [MS-XLS] 2.5.198): pasa las fórmulas de un .xls de su forma binaria (una lista de "tokens" Ptg en notación polaca inversa, el rgce) a
// texto como lo escribe un .xlsx (sin "=", funciones en inglés, hojas entre comillas simples si hace falta), y al revés. No depende del navegador.
//
//   const texto = XlsFormulas.decodificar(rgce, rgcb, ctx);       // null si la fórmula no se puede llevar (libro externo, tabla de datos, función desconocida…)
//   ctx = { f, c,                                  // celda donde está la fórmula (desde 0): las referencias relativas de las fórmulas compartidas salen de aquí
//           hojas: ['Hoja1', …],                   // nombres de las hojas, en el orden del libro
//           xti: [{ libro, primera, ultima }, …],  // tabla EXTERNSHEET: libro = índice en `libros`; primera/ultima = hojas (-1: borrada, -2: del complemento)
//           libros: [{ tipo: 'interno'|'complemento'|'externo', nombres: [...] }, …],    // los SUPBOOK, con los nombres de sus EXTERNNAME
//           nombres: [{ nombre, integrado }, …] }  // los registros NAME, en orden
//   XlsFormulas.FUNCIONES    // la tabla: [{ nombre, id, min, max, fija, args, ref, tipo }]
//
// Las funciones posteriores a Excel 97 se guardan como llamadas a una función "de usuario" cuyo nombre es `_xlfn.NOMBRE` (el mismo prefijo que usa un .xlsx);
// las del Analysis ToolPak (EDATE, EOMONTH, WORKDAY…) como llamadas a un nombre del libro de complemento, escrito en el idioma de Excel que guardó el archivo.
(function () {
  // <tabla>   (generado por Investigaciones\investigacion-formatos-antiguos\arnes\inyectar-tabla.mjs; id,nombre,mín,máx,fija,clases,devuelveReferencia)
  const T_FTAB = '0,COUNT,1,30,0,R,0;1,IF,2,3,0,VR,1;2,ISNA,1,1,1,V,0;3,ISERROR,1,1,1,V,0;4,SUM,1,30,0,,0;5,AVERAGE,1,30,0,R,0;6,MIN,1,30,0,R,0;7,MAX,1,30,0,R,0;8,ROW,0,0,0,,0;9,COLUMN,0,0,0,,0;10,NA,0,0,1,,0;11,NPV,2,30,0,VR,0;12,STDEV,1,30,0,R,0;13,DOLLAR,1,2,0,V,0;14,FIXED,1,3,0,V,0;15,SIN,1,1,1,V,0;16,COS,1,1,1,V,0;17,TAN,1,1,1,V,0;18,ATAN,1,1,1,V,0;19,PI,0,0,1,,0;20,SQRT,1,1,1,V,0;21,EXP,1,1,1,V,0;22,LN,1,1,1,V,0;23,LOG10,1,1,1,V,0;24,ABS,1,1,1,V,0;25,INT,1,1,1,V,0;26,SIGN,1,1,1,V,0;27,ROUND,2,2,1,V,0;28,LOOKUP,2,3,0,VR,0;29,INDEX,2,4,0,RV,1;30,REPT,2,2,1,V,0;31,MID,3,3,1,V,0;32,LEN,1,1,1,V,0;33,VALUE,1,1,1,V,0;34,TRUE,0,0,1,,0;35,FALSE,0,0,1,,0;36,AND,1,30,0,R,0;37,OR,1,30,0,R,0;38,NOT,1,1,1,V,0;39,MOD,2,2,1,V,0;40,DCOUNT,3,3,1,R,0;41,DSUM,3,3,1,R,0;42,DAVERAGE,3,3,1,R,0;43,DMIN,3,3,1,R,0;44,DMAX,3,3,1,R,0;45,DSTDEV,3,3,1,R,0;46,VAR,1,30,0,R,0;47,DVAR,3,3,1,R,0;48,TEXT,2,2,1,V,0;49,LINEST,1,4,0,R,0;50,TREND,1,4,0,R,0;51,LOGEST,1,4,0,R,0;52,GROWTH,1,4,0,R,0;56,PV,3,5,0,V,0;57,FV,3,5,0,V,0;58,NPER,3,5,0,V,0;59,PMT,3,5,0,V,0;60,RATE,3,6,0,V,0;61,MIRR,3,3,1,RV,0;62,IRR,1,2,0,RV,0;63,RAND,0,0,1,,0;64,MATCH,2,3,0,VR,0;65,DATE,3,3,1,V,0;66,TIME,3,3,1,V,0;67,DAY,1,1,1,V,0;68,MONTH,1,1,1,V,0;69,YEAR,1,1,1,V,0;70,WEEKDAY,1,2,0,V,0;71,HOUR,1,1,1,V,0;72,MINUTE,1,1,1,V,0;73,SECOND,1,1,1,V,0;74,NOW,0,0,1,,0;75,AREAS,1,1,1,R,0;76,ROWS,1,1,1,R,0;77,COLUMNS,1,1,1,R,0;78,OFFSET,3,5,0,RV,1;82,SEARCH,2,3,0,V,0;83,TRANSPOSE,1,1,1,V,0;86,TYPE,1,1,1,V,0;92,SERIESSUM,4,4,1,VVVA,0;97,ATAN2,2,2,1,V,0;98,ASIN,1,1,1,V,0;99,ACOS,1,1,1,V,0;100,CHOOSE,2,30,0,VR,1;101,HLOOKUP,3,4,0,VRRV,0;102,VLOOKUP,3,4,0,VRRV,0;105,ISREF,1,1,1,R,0;109,LOG,1,2,0,V,0;111,CHAR,1,1,1,V,0;112,LOWER,1,1,1,V,0;113,UPPER,1,1,1,V,0;114,PROPER,1,1,1,V,0;115,LEFT,1,2,0,V,0;116,RIGHT,1,2,0,V,0;117,EXACT,2,2,1,V,0;118,TRIM,1,1,1,V,0;119,REPLACE,4,4,1,V,0;120,SUBSTITUTE,3,4,0,V,0;121,CODE,1,1,1,V,0;124,FIND,2,3,0,V,0;125,CELL,1,1,0,V,0;126,ISERR,1,1,1,V,0;127,ISTEXT,1,1,1,V,0;128,ISNUMBER,1,1,1,V,0;129,ISBLANK,1,1,1,V,0;130,T,1,1,1,R,0;131,N,1,1,1,R,0;140,DATEVALUE,1,1,1,V,0;141,TIMEVALUE,1,1,1,V,0;142,SLN,3,3,1,V,0;143,SYD,4,4,1,V,0;144,DDB,4,5,0,V,0;148,INDIRECT,1,2,0,V,1;162,CLEAN,1,1,1,V,0;163,MDETERM,1,1,1,A,0;164,MINVERSE,1,1,1,A,0;165,MMULT,2,2,1,A,0;167,IPMT,4,6,0,V,0;168,PPMT,4,6,0,V,0;169,COUNTA,1,30,0,R,0;183,PRODUCT,1,30,0,R,0;184,FACT,1,1,1,V,0;189,DPRODUCT,3,3,1,R,0;190,ISNONTEXT,1,1,1,V,0;193,STDEVP,1,30,0,R,0;194,VARP,1,30,0,R,0;195,DSTDEVP,3,3,1,R,0;196,DVARP,3,3,1,R,0;197,TRUNC,1,2,0,V,0;198,ISLOGICAL,1,1,1,V,0;199,DCOUNTA,3,3,1,R,0;204,USDOLLAR,1,2,0,V,0;205,FINDB,2,3,0,V,0;206,SEARCHB,2,3,0,V,0;207,REPLACEB,4,4,1,V,0;208,LEFTB,1,2,0,V,0;209,RIGHTB,1,2,0,V,0;210,MIDB,3,3,1,V,0;211,LENB,1,1,1,V,0;212,ROUNDUP,2,2,1,V,0;213,ROUNDDOWN,2,2,1,V,0;214,ASC,1,1,1,V,0;215,DBCS,1,1,1,V,0;216,RANK,2,3,0,VRV,0;219,ADDRESS,2,5,0,V,0;220,DAYS360,2,3,0,V,0;221,TODAY,0,0,1,,0;222,VDB,5,7,0,V,0;227,MEDIAN,1,30,0,R,0;228,SUMPRODUCT,1,30,0,A,0;229,SINH,1,1,1,V,0;230,COSH,1,1,1,V,0;231,TANH,1,1,1,V,0;232,ASINH,1,1,1,V,0;233,ACOSH,1,1,1,V,0;234,ATANH,1,1,1,V,0;235,DGET,3,3,1,R,0;244,INFO,1,1,1,V,0;247,DB,4,5,0,V,0;252,FREQUENCY,2,2,1,R,0;261,ERROR.TYPE,1,1,1,V,0;269,AVEDEV,1,30,0,R,0;270,BETADIST,3,5,0,V,0;271,GAMMALN,1,1,1,V,0;272,BETAINV,3,5,0,V,0;273,BINOMDIST,4,4,1,V,0;274,CHIDIST,2,2,1,V,0;275,CHIINV,2,2,1,V,0;276,COMBIN,2,2,1,V,0;277,CONFIDENCE,3,3,1,V,0;278,CRITBINOM,3,3,1,V,0;279,EVEN,1,1,1,V,0;280,EXPONDIST,3,3,1,V,0;281,FDIST,3,3,1,V,0;282,FINV,3,3,1,V,0;283,FISHER,1,1,1,V,0;284,FISHERINV,1,1,1,V,0;285,FLOOR,2,2,1,V,0;286,GAMMADIST,4,4,1,V,0;287,GAMMAINV,3,3,1,V,0;288,CEILING,2,2,1,V,0;289,HYPGEOMDIST,4,4,1,V,0;290,LOGNORMDIST,3,3,1,V,0;291,LOGINV,3,3,1,V,0;292,NEGBINOMDIST,3,3,1,V,0;293,NORMDIST,4,4,1,V,0;294,NORMSDIST,1,1,1,V,0;295,NORMINV,3,3,1,V,0;296,NORMSINV,1,1,1,V,0;297,STANDARDIZE,3,3,1,V,0;298,ODD,1,1,1,V,0;299,PERMUT,2,2,1,V,0;300,POISSON,3,3,1,V,0;301,TDIST,3,3,1,V,0;302,WEIBULL,4,4,1,V,0;303,SUMXMY2,2,2,1,A,0;304,SUMX2MY2,2,2,1,A,0;305,SUMX2PY2,2,2,1,A,0;306,CHITEST,2,2,1,A,0;307,CORREL,2,2,1,A,0;308,COVAR,2,2,1,A,0;309,FORECAST,3,3,1,VA,0;310,FTEST,2,2,1,A,0;311,INTERCEPT,2,2,1,A,0;312,PEARSON,2,2,1,A,0;313,RSQ,2,2,1,A,0;314,STEYX,2,2,1,A,0;315,SLOPE,2,2,1,A,0;316,TTEST,4,4,1,AAV,0;317,PROB,3,4,0,AAV,0;318,DEVSQ,1,30,0,R,0;319,GEOMEAN,1,30,0,R,0;320,HARMEAN,1,30,0,R,0;321,SUMSQ,1,30,0,R,0;322,KURT,1,30,0,R,0;323,SKEW,1,30,0,R,0;324,ZTEST,2,3,0,RV,0;325,LARGE,2,2,1,RV,0;326,SMALL,2,2,1,RV,0;327,QUARTILE,2,2,1,RV,0;328,PERCENTILE,2,2,1,RV,0;329,PERCENTRANK,2,3,0,RV,0;330,MODE,1,30,0,A,0;331,TRIMMEAN,2,2,1,RV,0;332,TINV,2,2,1,V,0;336,CONCATENATE,1,30,0,V,0;337,POWER,2,2,1,V,0;342,RADIANS,1,1,1,V,0;343,DEGREES,1,1,1,V,0;344,SUBTOTAL,2,30,0,VR,0;345,SUMIF,2,3,0,RVR,0;346,COUNTIF,2,2,1,RV,0;347,COUNTBLANK,1,1,1,R,0;350,ISPMT,4,4,1,V,0;351,DATEDIF,3,3,1,V,0;352,DATESTRING,1,1,1,V,0;353,NUMBERSTRING,2,2,1,V,0;354,ROMAN,1,2,0,V,0;358,GETPIVOTDATA,2,30,0,RRV,0;359,HYPERLINK,1,2,0,V,0;360,PHONETIC,1,1,1,R,0;361,AVERAGEA,1,30,0,R,0;362,MAXA,1,30,0,R,0;363,MINA,1,30,0,R,0;364,STDEVPA,1,30,0,R,0;365,VARPA,1,30,0,R,0;366,STDEVA,1,30,0,R,0;367,VARA,1,30,0,R,0;368,BAHTTEXT,1,1,1,V,0;369,THAIDAYOFWEEK,1,1,1,V,0;370,THAIDIGIT,1,1,1,V,0;371,THAIMONTHOFYEAR,1,1,1,V,0;372,THAINUMSOUND,1,1,1,V,0;373,THAINUMSTRING,1,1,1,V,0;374,THAISTRINGLENGTH,1,1,1,V,0;375,ISTHAIDIGIT,1,1,1,V,0;376,ROUNDBAHTDOWN,1,1,1,V,0;377,ROUNDBAHTUP,1,1,1,V,0;378,THAIYEAR,1,1,1,V,0;379,RTD,2,5,0,V,0';
  // nombre,mín,máx,clases,devuelveReferencia,llevaPrefijoEnXlsx
  const T_FUTURAS = 'ACOT,1,1,V,0,1;ACOTH,1,1,V,0,1;AGGREGATE,3,4,VVR,0,1;ARABIC,1,1,V,0,1;ARRAYTOTEXT,1,2,AV,0,1;AVERAGEIF,2,3,RVR,0,0;AVERAGEIFS,3,30,RRV,0,0;BASE,2,3,V,0,1;BETA.DIST,4,6,V,0,1;BETA.INV,3,5,V,0,1;BINOM.DIST,4,4,V,0,1;BINOM.DIST.RANGE,3,4,V,0,1;BINOM.INV,3,3,V,0,1;BITAND,2,2,V,0,1;BITLSHIFT,2,2,V,0,1;BITOR,2,2,V,0,1;BITRSHIFT,2,2,V,0,1;BITXOR,2,2,V,0,1;BYCOL,1,2,RV,0,1;BYROW,1,2,RV,0,1;CEILING.MATH,1,3,V,0,1;CEILING.PRECISE,1,2,V,0,1;CHISQ.DIST,3,3,V,0,1;CHISQ.DIST.RT,2,2,V,0,1;CHISQ.INV,2,2,V,0,1;CHISQ.INV.RT,2,2,V,0,1;CHISQ.TEST,2,2,A,0,1;CHOOSECOLS,2,30,RV,0,1;CHOOSEROWS,2,30,RV,0,1;COMBINA,2,2,V,0,1;CONCAT,1,30,R,0,1;CONFIDENCE.NORM,3,3,V,0,1;CONFIDENCE.T,3,3,V,0,1;COT,1,1,V,0,1;COTH,1,1,V,0,1;COUNTIFS,2,30,RV,0,0;COVARIANCE.P,2,2,A,0,1;COVARIANCE.S,2,2,A,0,1;CSC,1,1,V,0,1;CSCH,1,1,V,0,1;CUBEKPIMEMBER,3,4,V,0,0;CUBEMEMBER,2,3,VRV,0,0;CUBEMEMBERPROPERTY,3,3,V,0,0;CUBERANKEDMEMBER,3,4,V,0,0;CUBESET,2,5,VRV,0,0;CUBESETCOUNT,1,1,V,0,0;CUBEVALUE,1,30,VR,0,0;DAYS,2,2,V,0,1;DECIMAL,2,2,V,0,1;DROP,2,3,RV,1,1;ENCODEURL,1,1,V,0,1;ERF.PRECISE,1,1,R,0,1;ERFC.PRECISE,1,1,R,0,1;EUROCONVERT,0,30,R,1,0;EXPAND,2,4,AV,0,1;EXPON.DIST,3,3,V,0,1;F.DIST,4,4,V,0,1;F.DIST.RT,3,3,V,0,1;F.INV,3,3,V,0,1;F.INV.RT,3,3,V,0,1;F.TEST,2,2,A,0,1;FIELDVALUE,2,2,V,0,1;FILTER,2,3,RRV,0,1;FILTERXML,2,2,V,0,1;FLOOR.MATH,1,3,V,0,1;FLOOR.PRECISE,1,2,V,0,1;FORECAST.ETS,3,6,VRRV,0,1;FORECAST.ETS.CONFINT,3,7,VRRV,0,1;FORECAST.ETS.SEASONALITY,2,4,RRV,0,1;FORECAST.ETS.STAT,3,6,RRV,0,1;FORECAST.LINEAR,3,3,VA,0,1;GAMMA,1,1,V,0,1;GAMMA.DIST,4,4,V,0,1;GAMMA.INV,3,3,V,0,1;GAMMALN.PRECISE,1,1,V,0,1;GAUSS,1,1,V,0,1;GROUPBY,3,8,RAAVVAAV,0,1;HSTACK,1,30,R,0,1;HYPGEOM.DIST,5,5,V,0,1;IFERROR,2,2,VR,0,0;IFNA,2,2,VR,0,1;IFS,2,30,VR,1,1;IMAGE,1,5,V,0,1;IMCOSH,1,1,R,0,1;IMCOT,1,1,R,0,1;IMCSC,1,1,R,0,1;IMCSCH,1,1,R,0,1;IMSEC,1,1,R,0,1;IMSECH,1,1,R,0,1;IMSINH,1,1,R,0,1;IMTAN,1,1,R,0,1;ISFORMULA,1,1,R,0,1;ISO.CEILING,1,2,V,0,0;ISOMITTED,1,1,R,0,1;ISOWEEKNUM,1,1,V,0,1;JIS,0,30,R,1,0;LOGNORM.DIST,4,4,V,0,1;LOGNORM.INV,3,3,V,0,1;MAKEARRAY,3,3,VVR,0,1;MAP,2,30,R,0,1;MAXIFS,3,30,RRV,0,1;MINIFS,3,30,RRV,0,1;MODE.MULT,1,30,A,0,1;MODE.SNGL,1,30,A,0,1;MUNIT,1,1,V,0,1;NEGBINOM.DIST,4,4,V,0,1;NETWORKDAYS.INTL,2,4,RRVR,0,0;NORM.DIST,4,4,V,0,1;NORM.INV,3,3,V,0,1;NORM.S.DIST,2,2,V,0,1;NORM.S.INV,1,1,V,0,1;NUMBERVALUE,1,3,V,0,1;PDURATION,3,3,V,0,1;PERCENTILE.EXC,2,2,RV,0,1;PERCENTILE.INC,2,2,RV,0,1;PERCENTRANK.EXC,2,3,RV,0,1;PERCENTRANK.INC,2,3,RV,0,1;PERMUTATIONA,2,2,V,0,1;PHI,1,1,V,0,1;PIVOTBY,4,11,RRAAVVAV,0,1;POISSON.DIST,3,3,V,0,1;QUARTILE.EXC,2,2,RV,0,1;QUARTILE.INC,2,2,RV,0,1;RANDARRAY,0,5,V,0,1;RANK.AVG,2,3,VRV,0,1;RANK.EQ,2,3,VRV,0,1;REDUCE,3,3,RRV,1,1;REGEXEXTRACT,2,4,V,0,1;REGEXREPLACE,3,5,V,0,1;REGEXTEST,2,3,V,0,1;RRI,3,3,V,0,1;SCAN,3,3,RRV,0,1;SEC,1,1,V,0,1;SECH,1,1,V,0,1;SEQUENCE,1,4,V,0,1;SINGLE,1,1,V,1,1;SKEW.P,1,30,R,0,1;SORT,1,4,AAAV,0,1;SORTBY,2,30,AAV,0,1;SQL.REQUEST,0,30,R,1,0;STDEV.P,1,30,R,0,1;STDEV.S,1,30,R,0,1;STOCKHISTORY,2,30,V,0,1;SUMIFS,3,30,RRV,0,0;SWITCH,3,30,VVR,1,1;T.DIST,3,3,V,0,1;T.DIST.2T,2,2,V,0,1;T.DIST.RT,2,2,V,0,1;T.INV,2,2,V,0,1;T.INV.2T,2,2,V,0,1;T.TEST,4,4,AAV,0,1;TAKE,2,3,RV,1,1;TEXTAFTER,2,6,V,0,1;TEXTBEFORE,2,6,V,0,1;TEXTJOIN,3,30,RVR,0,1;TEXTSPLIT,2,6,V,0,1;TOCOL,1,3,RV,0,1;TOROW,1,3,RV,0,1;TRIMRANGE,1,3,RV,1,1;UNICHAR,1,1,V,0,1;UNICODE,1,1,V,0,1;UNIQUE,1,3,AV,0,1;VALUETOTEXT,1,2,V,0,1;VAR.P,1,30,R,0,1;VAR.S,1,30,R,0,1;VSTACK,1,30,R,0,1;WEBSERVICE,1,1,V,0,1;WEIBULL.DIST,4,4,V,0,1;WORKDAY.INTL,2,4,RRVR,0,0;WRAPCOLS,2,3,AV,0,1;WRAPROWS,2,3,AV,0,1;XLOOKUP,3,6,VRRRV,1,1;XMATCH,2,4,VRV,0,1;XOR,1,30,R,0,1;Z.TEST,2,3,RV,0,1';
  // nombre,mín,máx,clases,devuelveReferencia,nombreEnEspañol
  const T_COMPLEMENTO = 'ACCRINT,6,8,R,0,INT.ACUM;ACCRINTM,4,5,R,0,INT.ACUM.V;AMORDEGRC,6,7,R,0,AMORTIZ.PROGRE;AMORLINC,6,7,R,0,AMORTIZ.LIN;BESSELI,2,2,R,0,BESSELI;BESSELJ,2,2,R,0,BESSELJ;BESSELK,2,2,R,0,BESSELK;BESSELY,2,2,R,0,BESSELY;BIN2DEC,1,1,R,0,BIN.A.DEC;BIN2HEX,1,2,R,0,BIN.A.HEX;BIN2OCT,1,2,R,0,BIN.A.OCT;COMPLEX,2,3,R,0,COMPLEJO;CONVERT,3,3,R,0,CONVERTIR;COUPDAYBS,3,4,R,0,CUPON.DIAS.L1;COUPDAYS,3,4,R,0,CUPON.DIAS;COUPDAYSNC,3,4,R,0,CUPON.DIAS.L2;COUPNCD,3,4,R,0,CUPON.FECHA.L2;COUPNUM,3,4,R,0,CUPON.NUM;COUPPCD,3,4,R,0,CUPON.FECHA.L1;CUMIPMT,6,6,R,0,PAGO.INT.ENTRE;CUMPRINC,6,6,R,0,PAGO.PRINC.ENTRE;DEC2BIN,1,2,R,0,DEC.A.BIN;DEC2HEX,1,2,R,0,DEC.A.HEX;DEC2OCT,1,2,R,0,DEC.A.OCT;DELTA,1,2,R,0,DELTA;DISC,4,5,R,0,TASA.DESC;DOLLARDE,2,2,R,0,MONEDA.DEC;DOLLARFR,2,2,R,0,MONEDA.FRAC;DURATION,5,6,R,0,DURACION;EDATE,2,2,R,0,FECHA.MES;EFFECT,2,2,R,0,INT.EFECTIVO;EOMONTH,2,2,R,0,FIN.MES;ERF,1,2,R,0,FUN.ERROR;ERFC,1,1,R,0,FUN.ERROR.COMPL;FACTDOUBLE,1,1,R,0,FACT.DOBLE;FVSCHEDULE,2,2,R,0,VF.PLAN;GCD,1,30,R,0,M.C.D;GESTEP,1,2,R,0,MAYOR.O.IGUAL;HEX2BIN,1,2,R,0,HEX.A.BIN;HEX2DEC,1,1,R,0,HEX.A.DEC;HEX2OCT,1,2,R,0,HEX.A.OCT;IMABS,1,1,R,0,IM.ABS;IMAGINARY,1,1,R,0,IMAGINARIO;IMARGUMENT,1,1,R,0,IM.ANGULO;IMCONJUGATE,1,1,R,0,IM.CONJUGADA;IMCOS,1,1,R,0,IM.COS;IMDIV,2,2,R,0,IM.DIV;IMEXP,1,1,R,0,IM.EXP;IMLN,1,1,R,0,IM.LN;IMLOG10,1,1,R,0,IM.LOG10;IMLOG2,1,1,R,0,IM.LOG2;IMPOWER,2,2,R,0,IM.POT;IMPRODUCT,1,30,R,0,IM.PRODUCT;IMREAL,1,1,R,0,IM.REAL;IMSIN,1,1,R,0,IM.SENO;IMSQRT,1,1,R,0,IM.RAIZ2;IMSUB,2,2,R,0,IM.SUSTR;IMSUM,1,30,R,0,IM.SUM;INTRATE,4,5,R,0,TASA.INT;ISEVEN,1,1,R,0,ES.PAR;ISODD,1,1,R,0,ES.IMPAR;LCM,1,30,R,0,M.C.M;MDURATION,5,6,R,0,DURACION.MODIF;MROUND,2,2,R,0,REDOND.MULT;MULTINOMIAL,1,30,R,0,MULTINOMIAL;NETWORKDAYS,2,3,R,0,DIAS.LAB;NOMINAL,2,2,R,0,TASA.NOMINAL;OCT2BIN,1,2,R,0,OCT.A.BIN;OCT2DEC,1,1,R,0,OCT.A.DEC;OCT2HEX,1,2,R,0,OCT.A.HEX;ODDFPRICE,8,9,R,0,PRECIO.PER.IRREGULAR.1;ODDFYIELD,8,9,R,0,RENDTO.PER.IRREGULAR.1;ODDLPRICE,7,8,R,0,PRECIO.PER.IRREGULAR.2;ODDLYIELD,7,8,R,0,RENDTO.PER.IRREGULAR.2;PRICE,6,7,R,0,PRECIO;PRICEDISC,4,5,R,0,PRECIO.DESCUENTO;PRICEMAT,5,6,R,0,PRECIO.VENCIMIENTO;QUOTIENT,2,2,R,0,COCIENTE;RANDBETWEEN,2,2,R,0,ALEATORIO.ENTRE;RECEIVED,4,5,R,0,CANTIDAD.RECIBIDA;SERIESSUM,4,4,R,0,SUMA.SERIES;SQRTPI,1,1,R,0,RAIZ2PI;TBILLEQ,3,3,R,0,LETRA.DE.TEST.EQV.A.BONO;TBILLPRICE,3,3,R,0,LETRA.DE.TES.PRECIO;TBILLYIELD,3,3,R,0,LETRA.DE.TES.RENDTO;WEEKNUM,1,2,R,0,NUM.DE.SEMANA;WORKDAY,2,3,R,0,DIA.LAB;XIRR,2,3,R,0,TIR.NO.PER;XNPV,3,3,R,0,VNA.NO.PER;YEARFRAC,2,3,R,0,FRAC.AÑO;YIELDDISC,4,5,R,0,RENDTO.DESC;YIELDMAT,5,6,R,0,RENDTO.VENCTO';
  // </tabla>

  class NoSoportado extends Error {}
  const no = (motivo) => { throw new NoSoportado(motivo); };
  const u16 = (b, p) => b[p] | (b[p + 1] << 8);
  const i16 = (b, p) => (u16(b, p) << 16) >> 16;
  const u32 = (b, p) => (b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24)) >>> 0;
  const doble = (b, p) => { if (p < 0 || p + 8 > b.length) no('número cortado'); return new DataView(b.buffer, b.byteOffset + p, 8).getFloat64(0, true); };

  // ---- tabla de funciones -------------------------------------------------------------------------------------------------
  // args: la clase de cada argumento (R referencia, V valor, A matriz); la última se repite. ref: puede devolver una referencia.
  const FUNCIONES = [], POR_ID = new Map(), POR_NOMBRE = new Map(), FUTURA = new Map(), COMPLEMENTO = new Map(), COMPLEMENTO_LOCAL = new Map();
  const alta = (d) => { FUNCIONES.push(d); return d; };
  for (const r of T_FTAB.split(';')) {
    if (!r) continue;
    const [id, nombre, min, max, fija, args, ref] = r.split(',');
    const d = alta({ nombre, id: +id, min: +min, max: +max, fija: fija === '1', args, ref: ref === '1', tipo: 'ftab' });
    POR_ID.set(d.id, d); POR_NOMBRE.set(nombre, d);
  }
  for (const r of T_FUTURAS.split(';')) {
    if (!r) continue;
    const [nombre, min, max, args, ref, prefijo] = r.split(',');
    FUTURA.set(nombre, alta({ nombre, id: 255, min: +min, max: +max, fija: false, args, ref: ref === '1', tipo: 'futura', prefijo: prefijo === '1' }));
  }
  for (const r of T_COMPLEMENTO.split(';')) {
    if (!r) continue;
    const [nombre, min, max, args, ref, local] = r.split(',');
    const d = alta({ nombre, id: 255, min: +min, max: +max, fija: false, args, ref: ref === '1', tipo: 'complemento', local });
    COMPLEMENTO.set(nombre, d); COMPLEMENTO_LOCAL.set(local, d);
  }

  // ---- texto de los operandos ---------------------------------------------------------------------------------------------
  const ERRORES = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0F: '#VALUE!', 0x17: '#REF!', 0x1D: '#NAME?', 0x24: '#NUM!', 0x2A: '#N/A', 0x2B: '#GETTING_DATA' };
  const nombreCol = (c) => { let s = ''; for (c++; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s; return s; };

  // Un número como lo muestra Excel en una fórmula: hasta 15 cifras, y 1E+300 / 1E-07 para los muy grandes o chicos.
  function numeroATexto(n) {
    if (!Number.isFinite(n)) no('número no finito');
    if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
    let s = String(Number(n.toPrecision(15)));
    const m = /^(-?[\d.]+)e([+-])(\d+)$/.exec(s);
    if (m) s = m[1] + 'E' + m[2] + (m[3].length < 2 ? '0' : '') + m[3];
    return s;
  }

  // El nombre de una hoja en una referencia: entre comillas simples (y '' dentro) salvo que sea un nombre simple que no se parezca a una celda.
  function citarHoja(n) {
    if (/^[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) && !/^[Rr]\d*[Cc]\d*$/.test(n) && !/^(?:TRUE|FALSE)$/i.test(n)) return n;
    return "'" + n.replace(/'/g, "''") + "'";
  }

  // Una celda: fila y columna ya resueltas, con los $ de lo que es absoluto.
  const celdaATexto = (f, c, fRel, cRel) => (cRel ? '' : '$') + nombreCol(c) + (fRel ? '' : '$') + (f + 1);

  // Los pedazos de un tipo de referencia (PtgRef, PtgRefN y los de las áreas): [fila, columna, filaRelativa, columnaRelativa]. `n`: relativa a la celda de la fórmula.
  function ubicacion(fila, col, ctx, n) {
    const cRel = !!(col & 0x4000), fRel = !!(col & 0x8000);
    let c = col & 0x3FFF;
    let f = fila;
    if (n) {   // fórmula compartida o nombre: lo relativo es un desplazamiento (fila de 16 bits con signo, columna de 8)
      if (fRel) f = (ctx.f + ((fila << 16) >> 16)) & 0xFFFF;
      if (cRel) c = (ctx.c + ((c << 24) >> 24)) & 0xFF;
    }
    return [f, c, fRel, cRel];
  }
  function areaATexto(a, b) {
    const [f0, c0, fr0, cr0] = a, [f1, c1, fr1, cr1] = b;
    if (f0 === 0 && f1 === 0xFFFF && !(c0 === 0 && c1 >= 0xFF)) return (cr0 ? '' : '$') + nombreCol(c0) + ':' + (cr1 ? '' : '$') + nombreCol(c1);   // columnas enteras
    if (c0 === 0 && c1 >= 0xFF && !(f0 === 0 && f1 === 0xFFFF)) return (fr0 ? '' : '$') + (f0 + 1) + ':' + (fr1 ? '' : '$') + (f1 + 1);                       // filas enteras
    return celdaATexto(f0, c0, fr0, cr0) + ':' + celdaATexto(f1, c1, fr1, cr1);
  }

  // El prefijo de hoja de una referencia 3D: Hoja1!, 'Mi hoja'!, Hoja1:Hoja3!, #REF!
  function prefijoHoja(ixti, ctx) {
    const x = ctx.xti && ctx.xti[ixti];
    if (!x) no('ixti fuera de la tabla');
    const lib = ctx.libros && ctx.libros[x.libro];
    if (!lib) no('libro');
    if (lib.tipo === 'externo') no('referencia a otro libro');
    if (lib.tipo === 'complemento') no('referencia al complemento');
    if (x.primera < 0 || x.ultima < 0) return '#REF!';
    const a = ctx.hojas[x.primera], b = ctx.hojas[x.ultima];
    if (a === undefined || b === undefined) no('hoja que no existe');
    if (x.primera === x.ultima) return citarHoja(a) + '!';
    // varias hojas: Hoja1:Hoja3!, o todo entre comillas si alguna lo necesita ('Mi hoja:Otra'!)
    return (citarHoja(a) === a && citarHoja(b) === b ? a + ':' + b : "'" + (a + ':' + b).replace(/'/g, "''") + "'") + '!';
  }

  // Un nombre: el de la lista NAME, o el de una función de usuario. Los nombres integrados (Print_Area…) no se usan en fórmulas de celda.
  function nombreDe(i, ctx) {
    const n = ctx.nombres && ctx.nombres[i - 1];
    if (!n || !n.nombre || n.integrado) no('nombre');
    return n.nombre;
  }

  // El nombre de una función posterior a 97 tal como lo escribe un .xlsx: `_xlfn.IFS`, pero IFERROR y compañía (de 2007) van sin prefijo.
  function nombreFuturo(n) {
    const m = /^_xlfn\.(.+)$/.exec(n);
    const d = m && FUTURA.get(m[1]);
    return d && !d.prefijo ? m[1] : n;
  }

  // ---- decodificador -------------------------------------------------------------------------------------------------------
  const BINARIOS = { 0x03: '+', 0x04: '-', 0x05: '*', 0x06: '/', 0x07: '^', 0x08: '&', 0x09: '<', 0x0A: '<=', 0x0B: '=', 0x0C: '>=', 0x0D: '>', 0x0E: '<>', 0x0F: ' ', 0x10: ',', 0x11: ':' };
  const MAX_TOKENS = 8192;

  function arregloATexto(rgcb, pos) {   // PtgArray: en el rgcb, columnas-1 (1 byte), filas-1 (2 bytes) y los valores por fila
    const cols = rgcb[pos] + 1, filas = u16(rgcb, pos + 1) + 1;
    let p = pos + 3;
    if (cols * filas > 4096) no('matriz muy grande');
    const filasTxt = [];
    for (let i = 0; i < filas; i++) {
      const fila = [];
      for (let j = 0; j < cols; j++) {
        const t = rgcb[p++];
        if (t === 0x01) { fila.push(numeroATexto(doble(rgcb, p))); p += 8; }
        else if (t === 0x02) { const cch = u16(rgcb, p), ancho = rgcb[p + 2] & 1; fila.push('"' + new TextDecoder(ancho ? 'utf-16le' : 'latin1').decode(rgcb.subarray(p + 3, p + 3 + cch * (ancho + 1))).replace(/"/g, '""') + '"'); p += 3 + cch * (ancho + 1); }
        else if (t === 0x04) { fila.push(rgcb[p] ? 'TRUE' : 'FALSE'); p += 8; }
        else if (t === 0x10) { fila.push(ERRORES[rgcb[p]] || no('error')); p += 8; }
        else if (t === 0x00) { fila.push(''); p += 8; }
        else no('valor de matriz');
      }
      filasTxt.push(fila.join(','));
    }
    if (p > rgcb.length) no('matriz cortada');
    return { texto: '{' + filasTxt.join(';') + '}', fin: p };
  }

  function decodificar(rgce, rgcb, ctx) {
    try { return decod(rgce, rgcb || new Uint8Array(0), ctx || {}); } catch (e) { if (e instanceof NoSoportado) return null; throw e; }
  }

  function decod(r, rgcb, ctx) {
    const pila = [];
    const sacar = () => (pila.length ? pila.pop() : no('pila vacía'));
    const sacarN = (n) => { if (n > pila.length) no('pila corta'); return pila.splice(pila.length - n, n); };
    let posExtra = 0, tokens = 0, p = 0;
    for (; p < r.length;) {
      if (++tokens > MAX_TOKENS) no('fórmula muy larga');
      const ptg = r[p];
      if (ptg < 0x20) {
        if (BINARIOS[ptg]) { const b = sacar(), a = sacar(); pila.push(a + BINARIOS[ptg] + b); p += 1; continue; }
        switch (ptg) {
          case 0x01: case 0x02: no('fórmula compartida, de matriz o tabla de datos sin resolver');
          case 0x12: pila.push('+' + sacar()); p += 1; break;
          case 0x13: pila.push('-' + sacar()); p += 1; break;
          case 0x14: pila.push(sacar() + '%'); p += 1; break;
          case 0x15: pila.push('(' + sacar() + ')'); p += 1; break;
          case 0x16: pila.push(''); p += 1; break;   // argumento omitido
          case 0x17: { const cch = r[p + 1], ancho = r[p + 2] & 1; pila.push('"' + new TextDecoder(ancho ? 'utf-16le' : 'latin1').decode(r.subarray(p + 3, p + 3 + cch * (ancho + 1))).replace(/"/g, '""') + '"'); p += 3 + cch * (ancho + 1); break; }
          case 0x19: {   // PtgAttr: casi todos son pistas para calcular más rápido (saltos de IF y CHOOSE, espacios); solo SUM de un argumento cuenta
            const t = r[p + 1];
            if (t === 0x04) p += 4 + (u16(r, p + 2) + 1) * 2;
            else { if (t === 0x10) pila.push('SUM(' + sacar() + ')'); p += 4; }
            break;
          }
          case 0x1C: pila.push(ERRORES[r[p + 1]] || no('error')); p += 2; break;
          case 0x1D: pila.push(r[p + 1] ? 'TRUE' : 'FALSE'); p += 2; break;
          case 0x1E: pila.push(String(u16(r, p + 1))); p += 3; break;
          case 0x1F: pila.push(numeroATexto(doble(r, p + 1))); p += 9; break;
          default: no('token 0x' + ptg.toString(16));
        }
        continue;
      }
      switch (ptg & 0x1F) {
        case 0x00: { const a = arregloATexto(rgcb, posExtra); posExtra = a.fin; pila.push(a.texto); p += 8; break; }
        case 0x01: {   // PtgFunc: cantidad fija de argumentos
          const d = POR_ID.get(u16(r, p + 1));
          if (!d) no('función ' + u16(r, p + 1));
          pila.push(d.nombre + '(' + sacarN(d.max).join(',') + ')'); p += 3; break;
        }
        case 0x02: {   // PtgFuncVar
          const n = r[p + 1] & 0x7F, id = u16(r, p + 2) & 0x7FFF;
          const args = sacarN(n);
          if (id === 255) {   // función con nombre: el primer operando es el nombre
            if (!args.length) no('llamada sin nombre');
            pila.push(args[0] + '(' + args.slice(1).join(',') + ')');
          } else {
            const d = POR_ID.get(id);
            if (!d) no('función ' + id);
            pila.push(d.nombre + '(' + args.join(',') + ')');
          }
          p += 4; break;
        }
        case 0x03: pila.push(nombreFuturo(nombreDe(u32(r, p + 1), ctx))); p += 5; break;
        case 0x04: case 0x0C: { const [f, c, fr, cr] = ubicacion(u16(r, p + 1), u16(r, p + 3), ctx, (ptg & 0x1F) === 0x0C); pila.push(celdaATexto(f, c, fr, cr)); p += 5; break; }
        case 0x05: case 0x0D: {
          const n = (ptg & 0x1F) === 0x0D;
          pila.push(areaATexto(ubicacion(u16(r, p + 1), u16(r, p + 5), ctx, n), ubicacion(u16(r, p + 3), u16(r, p + 7), ctx, n))); p += 9; break;
        }
        case 0x06: { const n = u16(rgcb, posExtra); posExtra += 2 + n * 8; p += 7; break; }   // PtgMemArea: pista; sus referencias van en el rgcb
        case 0x07: case 0x08: case 0x09: p += (ptg & 0x1F) === 0x09 ? 3 : 7; break;          // PtgMemErr, PtgMemNoMem, PtgMemFunc: pistas
        case 0x0A: pila.push('#REF!'); p += 5; break;
        case 0x0B: pila.push('#REF!'); p += 9; break;
        case 0x19: {   // PtgNameX: nombre de un libro (complemento o interno)
          const x = ctx.xti && ctx.xti[u16(r, p + 1)], lib = x && ctx.libros && ctx.libros[x.libro];
          if (!lib) no('nombre externo');
          const nombre = lib.nombres && lib.nombres[u32(r, p + 3) - 1];
          if (lib.tipo === 'complemento') {
            const d = COMPLEMENTO.get(nombre) || COMPLEMENTO_LOCAL.get(nombre);
            if (!d) no('función de complemento desconocida: ' + nombre);
            pila.push(d.nombre);
          } else if (lib.tipo === 'interno' && ctx.nombres) pila.push(nombreFuturo(nombreDe(u32(r, p + 3), ctx)));
          else no('nombre de otro libro');
          p += 7; break;
        }
        case 0x1A: case 0x1B: case 0x1C: case 0x1D: {
          const pre = prefijoHoja(u16(r, p + 1), ctx), k = ptg & 0x1F;
          if (k === 0x1A) { const [f, c, fr, cr] = ubicacion(u16(r, p + 3), u16(r, p + 5), ctx, false); pila.push(pre + celdaATexto(f, c, fr, cr)); p += 7; }
          else if (k === 0x1B) { pila.push(pre + areaATexto(ubicacion(u16(r, p + 3), u16(r, p + 7), ctx, false), ubicacion(u16(r, p + 5), u16(r, p + 9), ctx, false))); p += 11; }
          else { pila.push(pre + '#REF!'); p += k === 0x1C ? 7 : 11; }
          break;
        }
        default: no('token 0x' + ptg.toString(16));
      }
    }
    if (p > r.length) no('token cortado');
    if (pila.length !== 1) no('la fórmula no deja un solo resultado');
    return pila[0];
  }

  globalThis.XlsFormulas = { decodificar, FUNCIONES, _numeroATexto: numeroATexto, _citarHoja: citarHoja };
})();

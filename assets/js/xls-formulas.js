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
  // <tabla>   (generado por Investigaciones\investigacion-formatos-antiguos\arnes\inyectar-tabla.mjs; id,nombre,mín,máx,fija,clases,qué clase devuelve dentro de SUM (R, A o V),volátil)
  const T_FTAB = '0,COUNT,1,30,0,RRRRRRRR,V,0,1;1,IF,2,3,0,VRR,R,0,0;2,ISNA,1,1,1,V,V,0,1;3,ISERROR,1,1,1,V,V,0,1;4,SUM,1,30,0,R,V,0,1;5,AVERAGE,1,30,0,RRRRRRRR,V,0,1;6,MIN,1,30,0,RRRRRRRR,V,0,1;7,MAX,1,30,0,RRRRRRRR,V,0,1;8,ROW,0,1,0,R,V,0,1;9,COLUMN,0,1,0,R,V,0,1;10,NA,0,0,1,,V,0,1;11,NPV,2,30,0,VRRRRRRR,V,0,1;12,STDEV,1,30,0,RRRRRRRR,V,0,1;13,DOLLAR,1,2,0,VV,V,0,1;14,FIXED,1,3,0,VVV,V,0,1;15,SIN,1,1,1,V,V,0,1;16,COS,1,1,1,V,V,0,1;17,TAN,1,1,1,V,V,0,1;18,ATAN,1,1,1,V,V,0,1;19,PI,0,0,1,,V,0,1;20,SQRT,1,1,1,V,V,0,1;21,EXP,1,1,1,V,V,0,1;22,LN,1,1,1,V,V,0,1;23,LOG10,1,1,1,V,V,0,1;24,ABS,1,1,1,V,V,0,1;25,INT,1,1,1,V,V,0,1;26,SIGN,1,1,1,V,V,0,1;27,ROUND,2,2,1,VV,V,0,1;28,LOOKUP,2,3,0,VRR,V,0,1;29,INDEX,2,4,0,RVVV,R,0,0;30,REPT,2,2,1,VV,V,0,1;31,MID,3,3,1,VVV,V,0,1;32,LEN,1,1,1,V,V,0,1;33,VALUE,1,1,1,V,V,0,1;34,TRUE,0,0,1,,V,0,1;35,FALSE,0,0,1,,V,0,1;36,AND,1,30,0,RRRRRRRR,V,0,1;37,OR,1,30,0,RRRRRRRR,V,0,1;38,NOT,1,1,1,V,V,0,1;39,MOD,2,2,1,VV,V,0,1;40,DCOUNT,3,3,1,RRR,V,0,1;41,DSUM,3,3,1,RRR,V,0,1;42,DAVERAGE,3,3,1,RRR,V,0,1;43,DMIN,3,3,1,RRR,V,0,1;44,DMAX,3,3,1,RRR,V,0,1;45,DSTDEV,3,3,1,RRR,V,0,1;46,VAR,1,30,0,RRRRRRRR,V,0,1;47,DVAR,3,3,1,RRR,V,0,1;48,TEXT,2,2,1,VV,V,0,1;49,LINEST,1,4,0,RRRR,A,0,1;50,TREND,1,4,0,RRRR,A,0,1;51,LOGEST,1,4,0,RRRR,A,0,1;52,GROWTH,1,4,0,RRRR,A,0,1;56,PV,3,5,0,VVVVV,V,0,1;57,FV,3,5,0,VVVVV,V,0,1;58,NPER,3,5,0,VVVVV,V,0,1;59,PMT,3,5,0,VVVVV,V,0,1;60,RATE,3,6,0,VVVVVV,V,0,1;61,MIRR,3,3,1,RVV,V,0,1;62,IRR,1,2,0,RV,V,0,1;63,RAND,0,0,1,,V,1,1;64,MATCH,2,3,0,VRR,V,0,1;65,DATE,3,3,1,VVV,V,0,1;66,TIME,3,3,1,VVV,V,0,1;67,DAY,1,1,1,V,V,0,1;68,MONTH,1,1,1,V,V,0,1;69,YEAR,1,1,1,V,V,0,1;70,WEEKDAY,1,2,0,VV,V,0,1;71,HOUR,1,1,1,V,V,0,1;72,MINUTE,1,1,1,V,V,0,1;73,SECOND,1,1,1,V,V,0,1;74,NOW,0,0,1,,V,1,1;75,AREAS,1,1,1,R,V,0,1;76,ROWS,1,1,1,R,V,0,1;77,COLUMNS,1,1,1,R,V,0,1;78,OFFSET,3,5,0,RVVVV,R,1,1;82,SEARCH,2,3,0,VVV,V,0,1;83,TRANSPOSE,1,1,1,V,A,0,0;86,TYPE,1,1,1,V,V,0,1;92,SERIESSUM,4,4,1,VVVA,V,0,1;97,ATAN2,2,2,1,VV,V,0,1;98,ASIN,1,1,1,V,V,0,1;99,ACOS,1,1,1,V,V,0,1;100,CHOOSE,2,30,0,VRRRRRRR,R,0,0;101,HLOOKUP,3,4,0,VRRV,V,0,0;102,VLOOKUP,3,4,0,VRRV,V,0,0;105,ISREF,1,1,1,R,V,0,1;109,LOG,1,2,0,VV,V,0,1;111,CHAR,1,1,1,V,V,0,1;112,LOWER,1,1,1,V,V,0,1;113,UPPER,1,1,1,V,V,0,1;114,PROPER,1,1,1,V,V,0,1;115,LEFT,1,2,0,VV,V,0,1;116,RIGHT,1,2,0,VV,V,0,1;117,EXACT,2,2,1,VV,V,0,1;118,TRIM,1,1,1,V,V,0,1;119,REPLACE,4,4,1,VVVV,V,0,1;120,SUBSTITUTE,3,4,0,VVVV,V,0,1;121,CODE,1,1,1,V,V,0,1;124,FIND,2,3,0,VVV,V,0,1;125,CELL,1,2,0,VR,V,1,0;126,ISERR,1,1,1,V,V,0,1;127,ISTEXT,1,1,1,V,V,0,1;128,ISNUMBER,1,1,1,V,V,0,1;129,ISBLANK,1,1,1,V,V,0,1;130,T,1,1,1,R,V,0,1;131,N,1,1,1,R,V,0,1;140,DATEVALUE,1,1,1,V,V,0,1;141,TIMEVALUE,1,1,1,V,V,0,1;142,SLN,3,3,1,VVV,V,0,1;143,SYD,4,4,1,VVVV,V,0,1;144,DDB,4,5,0,VVVVV,V,0,1;148,INDIRECT,1,2,0,VV,R,1,1;162,CLEAN,1,1,1,V,V,0,1;163,MDETERM,1,1,1,A,V,0,1;164,MINVERSE,1,1,1,A,A,0,1;165,MMULT,2,2,1,AA,A,0,1;167,IPMT,4,6,0,VVVVVV,V,0,1;168,PPMT,4,6,0,VVVVVV,V,0,1;169,COUNTA,1,30,0,RRRRRRRR,V,0,1;183,PRODUCT,1,30,0,RRRRRRRR,V,0,1;184,FACT,1,1,1,V,V,0,1;189,DPRODUCT,3,3,1,RRR,V,0,1;190,ISNONTEXT,1,1,1,V,V,0,1;193,STDEVP,1,30,0,RRRRRRRR,V,0,1;194,VARP,1,30,0,RRRRRRRR,V,0,1;195,DSTDEVP,3,3,1,RRR,V,0,1;196,DVARP,3,3,1,RRR,V,0,1;197,TRUNC,1,2,0,VV,V,0,1;198,ISLOGICAL,1,1,1,V,V,0,1;199,DCOUNTA,3,3,1,RRR,V,0,1;204,USDOLLAR,1,2,0,VV,V,0,1;205,FINDB,2,3,0,VVV,V,0,1;206,SEARCHB,2,3,0,VVV,V,0,1;207,REPLACEB,4,4,1,VVVV,V,0,1;208,LEFTB,1,2,0,VV,V,0,1;209,RIGHTB,1,2,0,VV,V,0,1;210,MIDB,3,3,1,VVV,V,0,1;211,LENB,1,1,1,V,V,0,1;212,ROUNDUP,2,2,1,VV,V,0,1;213,ROUNDDOWN,2,2,1,VV,V,0,1;214,ASC,1,1,1,V,V,0,1;215,DBCS,1,1,1,V,V,0,1;216,RANK,2,3,0,VRV,V,0,1;219,ADDRESS,2,5,0,VVVVV,V,0,1;220,DAYS360,2,3,0,VVV,V,0,1;221,TODAY,0,0,1,,V,1,1;222,VDB,5,7,0,VVVVVVV,V,0,1;227,MEDIAN,1,30,0,RRRRRRRR,V,0,1;228,SUMPRODUCT,1,30,0,AAAAAAAA,V,0,1;229,SINH,1,1,1,V,V,0,1;230,COSH,1,1,1,V,V,0,1;231,TANH,1,1,1,V,V,0,1;232,ASINH,1,1,1,V,V,0,1;233,ACOSH,1,1,1,V,V,0,1;234,ATANH,1,1,1,V,V,0,1;235,DGET,3,3,1,RRR,V,0,1;244,INFO,1,1,1,V,V,1,1;247,DB,4,5,0,VVVVV,V,0,1;252,FREQUENCY,2,2,1,RR,A,0,1;261,ERROR.TYPE,1,1,1,V,V,0,1;269,AVEDEV,1,30,0,RRRRRRRR,V,0,1;270,BETADIST,3,5,0,VVVVV,V,0,1;271,GAMMALN,1,1,1,V,V,0,1;272,BETAINV,3,5,0,VVVVV,V,0,1;273,BINOMDIST,4,4,1,VVVV,V,0,1;274,CHIDIST,2,2,1,VV,V,0,1;275,CHIINV,2,2,1,VV,V,0,1;276,COMBIN,2,2,1,VV,V,0,1;277,CONFIDENCE,3,3,1,VVV,V,0,1;278,CRITBINOM,3,3,1,VVV,V,0,1;279,EVEN,1,1,1,V,V,0,1;280,EXPONDIST,3,3,1,VVV,V,0,1;281,FDIST,3,3,1,VVV,V,0,1;282,FINV,3,3,1,VVV,V,0,1;283,FISHER,1,1,1,V,V,0,1;284,FISHERINV,1,1,1,V,V,0,1;285,FLOOR,2,2,1,VV,V,0,1;286,GAMMADIST,4,4,1,VVVV,V,0,1;287,GAMMAINV,3,3,1,VVV,V,0,1;288,CEILING,2,2,1,VV,V,0,1;289,HYPGEOMDIST,4,4,1,VVVV,V,0,1;290,LOGNORMDIST,3,3,1,VVV,V,0,1;291,LOGINV,3,3,1,VVV,V,0,1;292,NEGBINOMDIST,3,3,1,VVV,V,0,1;293,NORMDIST,4,4,1,VVVV,V,0,1;294,NORMSDIST,1,1,1,V,V,0,1;295,NORMINV,3,3,1,VVV,V,0,1;296,NORMSINV,1,1,1,V,V,0,1;297,STANDARDIZE,3,3,1,VVV,V,0,1;298,ODD,1,1,1,V,V,0,1;299,PERMUT,2,2,1,VV,V,0,1;300,POISSON,3,3,1,VVV,V,0,1;301,TDIST,3,3,1,VVV,V,0,1;302,WEIBULL,4,4,1,VVVV,V,0,1;303,SUMXMY2,2,2,1,AA,V,0,1;304,SUMX2MY2,2,2,1,AA,V,0,1;305,SUMX2PY2,2,2,1,AA,V,0,1;306,CHITEST,2,2,1,AA,V,0,1;307,CORREL,2,2,1,AA,V,0,1;308,COVAR,2,2,1,AA,V,0,1;309,FORECAST,3,3,1,VAA,V,0,1;310,FTEST,2,2,1,AA,V,0,1;311,INTERCEPT,2,2,1,AA,V,0,1;312,PEARSON,2,2,1,AA,V,0,1;313,RSQ,2,2,1,AA,V,0,1;314,STEYX,2,2,1,AA,V,0,1;315,SLOPE,2,2,1,AA,V,0,1;316,TTEST,4,4,1,AAVV,V,0,1;317,PROB,3,4,0,AAVV,V,0,1;318,DEVSQ,1,30,0,RRRRRRRR,V,0,1;319,GEOMEAN,1,30,0,RRRRRRRR,V,0,1;320,HARMEAN,1,30,0,RRRRRRRR,V,0,1;321,SUMSQ,1,30,0,RRRRRRRR,V,0,1;322,KURT,1,30,0,RRRRRRRR,V,0,1;323,SKEW,1,30,0,RRRRRRRR,V,0,1;324,ZTEST,2,3,0,RVV,V,0,1;325,LARGE,2,2,1,RV,V,0,1;326,SMALL,2,2,1,RV,V,0,1;327,QUARTILE,2,2,1,RV,V,0,1;328,PERCENTILE,2,2,1,RV,V,0,1;329,PERCENTRANK,2,3,0,RVV,V,0,1;330,MODE,1,30,0,AAAAAAAA,V,0,1;331,TRIMMEAN,2,2,1,RV,V,0,1;332,TINV,2,2,1,VV,V,0,1;336,CONCATENATE,1,30,0,VVVVVVVV,V,0,1;337,POWER,2,2,1,VV,V,0,1;342,RADIANS,1,1,1,V,V,0,1;343,DEGREES,1,1,1,V,V,0,1;344,SUBTOTAL,2,30,0,VRRRRRRR,V,0,1;345,SUMIF,2,3,0,RVR,V,0,1;346,COUNTIF,2,2,1,RV,V,0,1;347,COUNTBLANK,1,1,1,R,V,0,1;350,ISPMT,4,4,1,VVVV,V,0,1;351,DATEDIF,3,3,1,VVV,V,0,1;352,DATESTRING,1,1,1,V,V,0,1;353,NUMBERSTRING,2,2,1,VV,V,0,1;354,ROMAN,1,2,0,VV,V,0,1;358,GETPIVOTDATA,2,30,0,RRVVVVVV,V,0,1;359,HYPERLINK,1,2,0,VV,V,0,0;360,PHONETIC,1,1,1,R,V,0,1;361,AVERAGEA,1,30,0,RRRRRRRR,V,0,1;362,MAXA,1,30,0,RRRRRRRR,V,0,1;363,MINA,1,30,0,RRRRRRRR,V,0,1;364,STDEVPA,1,30,0,RRRRRRRR,V,0,1;365,VARPA,1,30,0,RRRRRRRR,V,0,1;366,STDEVA,1,30,0,RRRRRRRR,V,0,1;367,VARA,1,30,0,RRRRRRRR,V,0,1;368,BAHTTEXT,1,1,1,V,V,0,1;369,THAIDAYOFWEEK,1,1,1,V,V,0,1;370,THAIDIGIT,1,1,1,V,V,0,1;371,THAIMONTHOFYEAR,1,1,1,V,V,0,1;372,THAINUMSOUND,1,1,1,V,V,0,1;373,THAINUMSTRING,1,1,1,V,V,0,1;374,THAISTRINGLENGTH,1,1,1,V,V,0,1;375,ISTHAIDIGIT,1,1,1,V,V,0,1;376,ROUNDBAHTDOWN,1,1,1,V,V,0,1;377,ROUNDBAHTUP,1,1,1,V,V,0,1;378,THAIYEAR,1,1,1,V,V,0,1;379,RTD,2,5,0,V,V,0,1';
  // nombre,mín,máx,clases,clase que devuelve,volátil,pasaArgumentosAMatriz,llevaPrefijoEnXlsx
  const T_FUTURAS = 'ACOT,1,1,V,V,0,1,1;ACOTH,1,1,V,V,0,1,1;AGGREGATE,3,30,VVRRRRRR,V,0,1,1;ARABIC,1,1,V,V,0,1,1;ARRAYTOTEXT,1,2,AV,V,0,1,1;AVERAGEIF,2,3,RVR,V,0,1,0;AVERAGEIFS,3,29,RRVRVRVR,V,0,1,0;BAHTTEXT,1,1,V,V,0,1,0;BASE,2,3,VVV,V,0,1,1;BETA.DIST,4,6,VVVVVV,V,0,1,1;BETA.INV,3,5,VVVVV,V,0,1,1;BINOM.DIST,4,4,VVVV,V,0,1,1;BINOM.DIST.RANGE,3,4,VVVV,V,0,1,1;BINOM.INV,3,3,VVV,V,0,1,1;BITAND,2,2,VV,V,0,1,1;BITLSHIFT,2,2,VV,V,0,1,1;BITOR,2,2,VV,V,0,1,1;BITRSHIFT,2,2,VV,V,0,1,1;BITXOR,2,2,VV,V,0,1,1;BYCOL,1,2,RV,A,0,1,1;BYROW,1,2,RV,A,0,1,1;CEILING.MATH,1,3,VVV,V,0,1,1;CEILING.PRECISE,1,2,VV,V,0,1,1;CHISQ.DIST,3,3,VVV,V,0,1,1;CHISQ.DIST.RT,2,2,VV,V,0,1,1;CHISQ.INV,2,2,VV,V,0,1,1;CHISQ.INV.RT,2,2,VV,V,0,1,1;CHISQ.TEST,2,2,AA,V,0,1,1;CHOOSECOLS,2,30,RVVVVVVV,A,0,0,1;CHOOSEROWS,2,30,RVVVVVVV,A,0,0,1;COMBINA,2,2,VV,V,0,1,1;CONCAT,1,30,RRRRRRRR,V,0,1,1;CONFIDENCE.NORM,3,3,VVV,V,0,1,1;CONFIDENCE.T,3,3,VVV,V,0,1,1;COT,1,1,V,V,0,1,1;COTH,1,1,V,V,0,1,1;COUNTIFS,2,30,RVRVRVRV,V,0,1,0;COVARIANCE.P,2,2,AA,V,0,1,1;COVARIANCE.S,2,2,AA,V,0,1,1;CSC,1,1,V,V,0,1,1;CSCH,1,1,V,V,0,1,1;CUBEKPIMEMBER,3,4,VVVV,V,0,1,0;CUBEMEMBER,2,3,VRV,V,0,1,0;CUBEMEMBERPROPERTY,3,3,VVV,V,0,1,0;CUBERANKEDMEMBER,3,4,VVVV,V,0,1,0;CUBESET,2,5,VRVVV,V,0,1,0;CUBESETCOUNT,1,1,V,V,0,1,0;CUBEVALUE,1,30,VRRRRRRR,V,0,1,0;DAYS,2,2,VV,V,0,1,1;DECIMAL,2,2,VV,V,0,1,1;DROP,2,3,RVV,R,0,0,1;ENCODEURL,1,1,V,V,0,1,1;ERF.PRECISE,1,1,R,V,0,1,1;ERFC.PRECISE,1,1,R,V,0,1,1;EXPAND,2,4,AVVV,A,0,0,1;EXPON.DIST,3,3,VVV,V,0,1,1;F.DIST,4,4,VVVV,V,0,1,1;F.DIST.RT,3,3,VVV,V,0,1,1;F.INV,3,3,VVV,V,0,1,1;F.INV.RT,3,3,VVV,V,0,1,1;F.TEST,2,2,AA,V,0,1,1;FIELDVALUE,2,2,VV,A,0,0,1;FILTER,2,3,RRV,A,0,1,1;FILTERXML,2,2,VV,V,0,0,1;FLOOR.MATH,1,3,VVV,V,0,1,1;FLOOR.PRECISE,1,2,VV,V,0,1,1;FORECAST.ETS,3,6,VRRVVV,V,0,1,1;FORECAST.ETS.CONFINT,3,7,VRRVVVV,V,0,1,1;FORECAST.ETS.SEASONALITY,2,4,RRVV,V,0,1,1;FORECAST.ETS.STAT,3,6,RRVVVV,V,0,1,1;FORECAST.LINEAR,3,3,VAA,V,0,1,1;GAMMA,1,1,V,V,0,1,1;GAMMA.DIST,4,4,VVVV,V,0,1,1;GAMMA.INV,3,3,VVV,V,0,1,1;GAMMALN.PRECISE,1,1,V,V,0,1,1;GAUSS,1,1,V,V,0,1,1;GROUPBY,3,8,RAAVVAAV,A,0,1,1;HSTACK,1,30,RRRRRRRR,A,0,1,1;HYPGEOM.DIST,5,5,VVVVV,V,0,1,1;IFERROR,2,2,VR,V,0,0,0;IFNA,2,2,VR,V,0,0,1;IFS,2,30,VRVRVRVR,R,0,0,1;IMAGE,1,5,VVVVV,V,0,1,1;IMCOSH,1,1,R,V,0,1,1;IMCOT,1,1,R,V,0,1,1;IMCSC,1,1,R,V,0,1,1;IMCSCH,1,1,R,V,0,1,1;IMSEC,1,1,R,V,0,1,1;IMSECH,1,1,R,V,0,1,1;IMSINH,1,1,R,V,0,1,1;IMTAN,1,1,R,V,0,1,1;ISFORMULA,1,1,R,V,0,1,1;ISO.CEILING,1,2,VV,V,0,1,0;ISOMITTED,1,1,R,V,0,1,1;ISOWEEKNUM,1,1,V,V,0,1,1;ISTHAIDIGIT,1,1,V,V,0,1,0;LOGNORM.DIST,4,4,VVVV,V,0,1,1;LOGNORM.INV,3,3,VVV,V,0,1,1;MAKEARRAY,3,3,VVR,A,0,0,1;MAP,2,30,RRRRRRRR,A,0,1,1;MAXIFS,3,29,RRVRVRVR,V,0,1,1;MINIFS,3,29,RRVRVRVR,V,0,1,1;MODE.MULT,1,30,AAAAAAAA,A,0,1,1;MODE.SNGL,1,30,AAAAAAAA,V,0,1,1;MUNIT,1,1,V,A,0,0,1;NEGBINOM.DIST,4,4,VVVV,V,0,1,1;NETWORKDAYS.INTL,2,4,RRVR,V,0,1,0;NORM.DIST,4,4,VVVV,V,0,1,1;NORM.INV,3,3,VVV,V,0,1,1;NORM.S.DIST,2,2,VV,V,0,1,1;NORM.S.INV,1,1,V,V,0,1,1;NUMBERVALUE,1,3,VVV,V,0,1,1;PDURATION,3,3,VVV,V,0,1,1;PERCENTILE.EXC,2,2,RV,V,0,1,1;PERCENTILE.INC,2,2,RV,V,0,1,1;PERCENTRANK.EXC,2,3,RVV,V,0,1,1;PERCENTRANK.INC,2,3,RVV,V,0,1,1;PERMUTATIONA,2,2,VV,V,0,1,1;PHI,1,1,V,V,0,1,1;PIVOTBY,4,11,RRAAVVAV,A,0,1,1;POISSON.DIST,3,3,VVV,V,0,1,1;QUARTILE.EXC,2,2,RV,V,0,1,1;QUARTILE.INC,2,2,RV,V,0,1,1;RANDARRAY,0,5,VVVVV,A,1,1,1;RANK.AVG,2,3,VRV,V,0,1,1;RANK.EQ,2,3,VRV,V,0,1,1;REDUCE,3,3,RRV,R,0,0,1;REGEXEXTRACT,2,4,VVVV,A,0,0,1;REGEXREPLACE,3,5,VVVVV,V,0,1,1;REGEXTEST,2,3,VVV,V,0,1,1;ROUNDBAHTDOWN,1,1,V,V,0,1,0;ROUNDBAHTUP,1,1,V,V,0,1,0;RRI,3,3,VVV,V,0,1,1;RTD,3,30,VVVVVVVV,V,0,1,0;SCAN,3,3,RRV,A,0,0,1;SEC,1,1,V,V,0,1,1;SECH,1,1,V,V,0,1,1;SEQUENCE,1,4,VVVV,A,0,0,1;SINGLE,1,1,V,R,0,0,1;SKEW.P,1,30,RRRRRRRR,V,0,1,1;SORT,1,4,AAAV,A,0,1,1;SORTBY,2,29,AAVAVAVA,A,0,1,1;STDEV.P,1,30,RRRRRRRR,V,0,1,1;STDEV.S,1,30,RRRRRRRR,V,0,1,1;STOCKHISTORY,2,30,VVVVVVVV,A,0,0,1;SUMIFS,3,29,RRVRVRVR,V,0,1,0;SWITCH,3,30,VVRRRRRR,R,0,0,1;T.DIST,3,3,VVV,V,0,1,1;T.DIST.2T,2,2,VV,V,0,1,1;T.DIST.RT,2,2,VV,V,0,1,1;T.INV,2,2,VV,V,0,1,1;T.INV.2T,2,2,VV,V,0,1,1;T.TEST,4,4,AAVV,V,0,1,1;TAKE,2,3,RVV,R,0,0,1;TEXTAFTER,2,6,VVVVVV,V,0,1,1;TEXTBEFORE,2,6,VVVVVV,V,0,1,1;TEXTJOIN,3,30,RVRRRRRR,V,0,1,1;TEXTSPLIT,2,6,VVVVVV,A,0,0,1;THAIDAYOFWEEK,1,1,V,V,0,1,0;THAIDIGIT,1,1,V,V,0,1,0;THAIMONTHOFYEAR,1,1,V,V,0,1,0;THAINUMSOUND,1,1,V,V,0,1,0;THAINUMSTRING,1,1,V,V,0,1,0;THAISTRINGLENGTH,1,1,V,V,0,1,0;THAIYEAR,1,1,V,V,0,1,0;TOCOL,1,3,RVV,A,0,1,1;TOROW,1,3,RVV,A,0,1,1;TRIMRANGE,1,3,RVV,R,0,1,1;UNICHAR,1,1,V,V,0,1,1;UNICODE,1,1,V,V,0,1,1;UNIQUE,1,3,AVV,A,0,1,1;VALUETOTEXT,1,2,VV,V,0,1,1;VAR.P,1,30,RRRRRRRR,V,0,1,1;VAR.S,1,30,RRRRRRRR,V,0,1,1;VSTACK,1,30,RRRRRRRR,A,0,1,1;WEBSERVICE,1,1,V,V,0,1,1;WEIBULL.DIST,4,4,VVVV,V,0,1,1;WORKDAY.INTL,2,4,RRVR,V,0,1,0;WRAPCOLS,2,3,AVV,A,0,0,1;WRAPROWS,2,3,AVV,A,0,0,1;XLOOKUP,3,6,VRRRVV,R,0,0,1;XMATCH,2,4,VRVV,V,0,1,1;XOR,1,30,RRRRRRRR,V,0,1,1;Z.TEST,2,3,RVV,V,0,1,1';
  // nombre,mín,máx,clases,clase que devuelve,volátil,pasaArgumentosAMatriz,nombreEnEspañol
  const T_COMPLEMENTO = 'ACCRINT,6,8,RRRRRRRR,V,0,1,INT.ACUM;ACCRINTM,4,5,RRRRR,V,0,1,INT.ACUM.V;AMORDEGRC,6,7,RRRRRRR,V,0,1,AMORTIZ.PROGRE;AMORLINC,6,7,RRRRRRR,V,0,1,AMORTIZ.LIN;BESSELI,2,2,RR,V,0,1,BESSELI;BESSELJ,2,2,RR,V,0,1,BESSELJ;BESSELK,2,2,RR,V,0,1,BESSELK;BESSELY,2,2,RR,V,0,1,BESSELY;BIN2DEC,1,1,R,V,0,1,BIN.A.DEC;BIN2HEX,1,2,RR,V,0,1,BIN.A.HEX;BIN2OCT,1,2,RR,V,0,1,BIN.A.OCT;COMPLEX,2,3,RRR,V,0,1,COMPLEJO;CONVERT,3,3,RRR,V,0,1,CONVERTIR;COUPDAYBS,3,4,RRRR,V,0,1,CUPON.DIAS.L1;COUPDAYS,3,4,RRRR,V,0,1,CUPON.DIAS;COUPDAYSNC,3,4,RRRR,V,0,1,CUPON.DIAS.L2;COUPNCD,3,4,RRRR,V,0,1,CUPON.FECHA.L2;COUPNUM,3,4,RRRR,V,0,1,CUPON.NUM;COUPPCD,3,4,RRRR,V,0,1,CUPON.FECHA.L1;CUMIPMT,6,6,RRRRRR,V,0,1,PAGO.INT.ENTRE;CUMPRINC,6,6,RRRRRR,V,0,1,PAGO.PRINC.ENTRE;DEC2BIN,1,2,RR,V,0,1,DEC.A.BIN;DEC2HEX,1,2,RR,V,0,1,DEC.A.HEX;DEC2OCT,1,2,RR,V,0,1,DEC.A.OCT;DELTA,1,2,RR,V,0,1,DELTA;DISC,4,5,RRRRR,V,0,1,TASA.DESC;DOLLARDE,2,2,RR,V,0,1,MONEDA.DEC;DOLLARFR,2,2,RR,V,0,1,MONEDA.FRAC;DURATION,5,6,RRRRRR,V,0,1,DURACION;EDATE,2,2,RR,V,0,1,FECHA.MES;EFFECT,2,2,RR,V,0,1,INT.EFECTIVO;EOMONTH,2,2,RR,V,0,1,FIN.MES;ERF,1,2,RR,V,0,1,FUN.ERROR;ERFC,1,1,R,V,0,1,FUN.ERROR.COMPL;FACTDOUBLE,1,1,R,V,0,1,FACT.DOBLE;FVSCHEDULE,2,2,RR,V,0,1,VF.PLAN;GCD,1,30,RRRRRRRR,V,0,1,M.C.D;GESTEP,1,2,RR,V,0,1,MAYOR.O.IGUAL;HEX2BIN,1,2,RR,V,0,1,HEX.A.BIN;HEX2DEC,1,1,R,V,0,1,HEX.A.DEC;HEX2OCT,1,2,RR,V,0,1,HEX.A.OCT;IMABS,1,1,R,V,0,1,IM.ABS;IMAGINARY,1,1,R,V,0,1,IMAGINARIO;IMARGUMENT,1,1,R,V,0,1,IM.ANGULO;IMCONJUGATE,1,1,R,V,0,1,IM.CONJUGADA;IMCOS,1,1,R,V,0,1,IM.COS;IMDIV,2,2,RR,V,0,1,IM.DIV;IMEXP,1,1,R,V,0,1,IM.EXP;IMLN,1,1,R,V,0,1,IM.LN;IMLOG10,1,1,R,V,0,1,IM.LOG10;IMLOG2,1,1,R,V,0,1,IM.LOG2;IMPOWER,2,2,RR,V,0,1,IM.POT;IMPRODUCT,1,30,RRRRRRRR,V,0,1,IM.PRODUCT;IMREAL,1,1,R,V,0,1,IM.REAL;IMSIN,1,1,R,V,0,1,IM.SENO;IMSQRT,1,1,R,V,0,1,IM.RAIZ2;IMSUB,2,2,RR,V,0,1,IM.SUSTR;IMSUM,1,30,RRRRRRRR,V,0,1,IM.SUM;INTRATE,4,5,RRRRR,V,0,1,TASA.INT;ISEVEN,1,1,R,V,0,1,ES.PAR;ISODD,1,1,R,V,0,1,ES.IMPAR;LCM,1,30,RRRRRRRR,V,0,1,M.C.M;MDURATION,5,6,RRRRRR,V,0,1,DURACION.MODIF;MROUND,2,2,RR,V,0,1,REDOND.MULT;MULTINOMIAL,1,30,RRRRRRRR,V,0,1,MULTINOMIAL;NETWORKDAYS,2,3,RRR,V,0,1,DIAS.LAB;NOMINAL,2,2,RR,V,0,1,TASA.NOMINAL;OCT2BIN,1,2,RR,V,0,1,OCT.A.BIN;OCT2DEC,1,1,R,V,0,1,OCT.A.DEC;OCT2HEX,1,2,RR,V,0,1,OCT.A.HEX;ODDFPRICE,8,9,RRRRRRRR,V,0,1,PRECIO.PER.IRREGULAR.1;ODDFYIELD,8,9,RRRRRRRR,V,0,1,RENDTO.PER.IRREGULAR.1;ODDLPRICE,7,8,RRRRRRRR,V,0,1,PRECIO.PER.IRREGULAR.2;ODDLYIELD,7,8,RRRRRRRR,V,0,1,RENDTO.PER.IRREGULAR.2;PRICE,6,7,RRRRRRR,V,0,1,PRECIO;PRICEDISC,4,5,RRRRR,V,0,1,PRECIO.DESCUENTO;PRICEMAT,5,6,RRRRRR,V,0,1,PRECIO.VENCIMIENTO;QUOTIENT,2,2,RR,V,0,1,COCIENTE;RANDBETWEEN,2,2,RR,V,1,1,ALEATORIO.ENTRE;RECEIVED,4,5,RRRRR,V,0,1,CANTIDAD.RECIBIDA;SERIESSUM,4,4,RRRR,V,0,1,SUMA.SERIES;SQRTPI,1,1,R,V,0,1,RAIZ2PI;TBILLEQ,3,3,RRR,V,0,1,LETRA.DE.TEST.EQV.A.BONO;TBILLPRICE,3,3,RRR,V,0,1,LETRA.DE.TES.PRECIO;TBILLYIELD,3,3,RRR,V,0,1,LETRA.DE.TES.RENDTO;WEEKNUM,1,2,RR,V,0,1,NUM.DE.SEMANA;WORKDAY,2,3,RRR,V,0,1,DIA.LAB;XIRR,2,3,RRR,V,0,1,TIR.NO.PER;XNPV,3,3,RRR,V,0,1,VNA.NO.PER;YEARFRAC,2,3,RRR,V,0,1,FRAC.AÑO;YIELDDISC,4,5,RRRRR,V,0,1,RENDTO.DESC;YIELDMAT,5,6,RRRRRR,V,0,1,RENDTO.VENCTO';
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
  // La clase del argumento i: las medidas (hasta 8) y, más allá, la cola que se repite (período 1: SUM; período 2: COUNTIFS, SWITCH).
  function periodoDe(a) {
    for (let p = 1; p <= 4 && p < a.length; p++) {
      let ok = true;
      for (let i = Math.max(p, a.length - 4); i < a.length; i++) if (a[i] !== a[i - p]) ok = false;
      if (ok) return p;
    }
    return 1;
  }
  const alta = (d) => { d.periodo = periodoDe(d.args); FUNCIONES.push(d); return d; };
  const claseDeArg = (d, i) => (!d ? 'R' : !d.args ? 'V' : i < d.args.length ? d.args[i] : d.args[d.args.length - d.periodo + ((i - d.args.length) % d.periodo)]);
  for (const r of T_FTAB.split(';')) {
    if (!r) continue;
    const [id, nombre, min, max, fija, args, ref, vol, lift] = r.split(',');
    const d = alta({ nombre, id: +id, min: +min, max: +max, fija: fija === '1', args, ret: ref, vol: vol === '1', lift: lift === '1', tipo: 'ftab' });
    POR_ID.set(d.id, d); POR_NOMBRE.set(nombre, d);
  }
  for (const r of T_FUTURAS.split(';')) {
    if (!r) continue;
    const [nombre, min, max, args, ref, vol, lift, prefijo] = r.split(',');
    FUTURA.set(nombre, alta({ nombre, id: 255, min: +min, max: +max, fija: false, args, ret: ref, vol: vol === '1', lift: lift === '1', tipo: 'futura', prefijo: prefijo === '1' }));
  }
  for (const r of T_COMPLEMENTO.split(';')) {
    if (!r) continue;
    const [nombre, min, max, args, ref, vol, lift, local] = r.split(',');
    const d = alta({ nombre, id: 255, min: +min, max: +max, fija: false, args, ret: ref, vol: vol === '1', lift: lift === '1', tipo: 'complemento', local });
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

  // ---- codificador ---------------------------------------------------------------------------------------------------------
  // Del texto de una fórmula (sin "=") a los bytes que Excel espera, con las mismas reglas que sigue Excel al guardar un .xls:
  //  - la clase de cada token (R referencia, V valor, A matriz) la fija quien lo usa: un argumento de SUM pide referencia, el de ABS pide valor, el de SUMPRODUCT matriz;
  //    los operadores aritméticos piden valor (matriz si están dentro de una matriz); : , y el espacio piden referencia;
  //  - IF y CHOOSE llevan saltos (PtgAttrIf, PtgAttrGoto, PtgAttrChoose) y SUM de un argumento es PtgAttrSum;
  //  - una fórmula con una función volátil (NOW, RAND, OFFSET…) empieza con PtgAttrSemi;
  //  - un número negativo escrito directamente (-1) es un número, no un menos delante;
  //  - una referencia fuera de 65.536 filas × 256 columnas es #REF!.
  // Los nombres, las hojas de las referencias 3D y las funciones de complemento se piden a `ctx` (el que arma el .xls decide qué índice les toca):
  //   ctx = { hoja(nombre) → índice de la hoja o -1,  ixti(primera, ultima) → índice en EXTERNSHEET,  nombre(texto) → índice (desde 1) en la lista NAME o null,
  //           complemento(def) → { ixti, idx } del nombre de la función en el libro de complemento, o null }
  // Devuelve { rgce, rgcb } (Uint8Array) o null si no se puede (referencia a una tabla, nombre que el libro no define, más de 30 argumentos, texto de más de 255 caracteres…).
  const ERROR_A_CODIGO = { '#NULL!': 0x00, '#DIV/0!': 0x07, '#VALUE!': 0x0F, '#REF!': 0x17, '#NAME?': 0x1D, '#NUM!': 0x24, '#N/A': 0x2A, '#GETTING_DATA': 0x2B };
  const CLASE = { R: 0x20, V: 0x40, A: 0x60 };
  const MAX_RGCE = 8000, MAX_TEXTO_LIT = 255, MAX_PROFUNDIDAD = 400;   // la profundidad del árbol: una suma de 400 términos ya es más de lo que cabe en una fórmula de 97
  const PREFIJO_HOJA = "(?:(?:'((?:[^']|'')+)'|([A-Za-z_\\u00C0-\\uFFFF][A-Za-z0-9_.\\u00C0-\\uFFFF]*))!)?";
  const CELDA = '(\\$?)([A-Za-z]{1,3})(\\$?)(\\d{1,7})';
  const RE_REF = new RegExp('^' + PREFIJO_HOJA + '(?:' + CELDA + '(?::' + CELDA + ')?|(\\$?)([A-Za-z]{1,3}):(\\$?)([A-Za-z]{1,3})|(\\$?)(\\d{1,7}):(\\$?)(\\d{1,7})|(#REF!))(?![A-Za-z0-9_.(\\u00C0-\\uFFFF])');
  const RE_NUM = /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/;
  const RE_ID = /^[A-Za-z_À-￿\\][A-Za-z0-9_.À-￿]*/;
  const RE_ERR = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA)/i;
  const colANum = (s) => { let n = 0; for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) & 31); return n - 1; };

  // El texto a tokens. Cada uno lleva `esp`: ¿había un espacio antes? (entre dos referencias, un espacio es el operador de intersección).
  function tokenizar(t) {
    const out = [];
    let i = 0, esp = false;
    if (t.length > 8192) no('fórmula muy larga');
    while (i < t.length) {
      const ch = t[i];
      if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') { esp = true; i++; continue; }
      const resto = t.slice(i, i + 400);
      let m;
      const poner = (tok, largo) => { tok.esp = esp; esp = false; out.push(tok); i += largo; };
      if (ch === '"') {
        let j = i + 1, s = '';
        for (;;) { if (j >= t.length) no('texto sin cerrar'); if (t[j] === '"') { if (t[j + 1] === '"') { s += '"'; j += 2; continue; } break; } s += t[j++]; }
        poner({ k: 'str', v: s }, j + 1 - i); continue;
      }
      if ((m = RE_REF.exec(resto))) {
        const tok = { k: 'ref', hoja: null, fuera: false };
        if (m[1] !== undefined || m[2] !== undefined) {
          const nombre = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2], p = nombre.indexOf(':');
          tok.hoja = p < 0 ? { a: nombre, b: nombre } : { a: nombre.slice(0, p), b: nombre.slice(p + 1) };
        }
        if (m[19]) tok.modo = 'error';
        else if (m[3] !== undefined) {   // celda o área
          tok.f0 = +m[6] - 1; tok.c0 = colANum(m[4].toUpperCase()); tok.fa0 = m[5] === '$'; tok.ca0 = m[3] === '$';
          if (m[7] !== undefined) { tok.modo = 'area'; tok.f1 = +m[10] - 1; tok.c1 = colANum(m[8].toUpperCase()); tok.fa1 = m[9] === '$'; tok.ca1 = m[7] === '$'; } else tok.modo = 'celda';
        } else if (m[11] !== undefined || m[12] !== undefined) {   // columnas enteras A:C
          tok.modo = 'cols'; tok.c0 = colANum(m[12].toUpperCase()); tok.ca0 = m[11] === '$'; tok.c1 = colANum(m[14].toUpperCase()); tok.ca1 = m[13] === '$';
        } else {   // filas enteras 1:3
          tok.modo = 'filas'; tok.f0 = +m[16] - 1; tok.fa0 = m[15] === '$'; tok.f1 = +m[18] - 1; tok.fa1 = m[17] === '$';
        }
        poner(tok, m[0].length); continue;
      }
      if ((m = RE_ERR.exec(resto))) { poner({ k: 'err', v: m[0].toUpperCase() }, m[0].length); continue; }
      if ((m = RE_NUM.exec(resto))) { poner({ k: 'num', v: parseFloat(m[0]) }, m[0].length); continue; }
      if ((m = RE_ID.exec(resto))) {
        const sig = /^\s*\(/.test(t.slice(i + m[0].length, i + m[0].length + 20));
        poner({ k: 'id', v: m[0], fn: sig }, m[0].length); continue;
      }
      const dos = t.slice(i, i + 2);
      if (dos === '<=' || dos === '>=' || dos === '<>') { poner({ k: 'op', v: dos }, 2); continue; }
      if ('+-*/^&=<>%:,;(){}'.includes(ch)) { poner({ k: 'op', v: ch }, 1); continue; }
      no('carácter ' + ch);   // [ de tablas estructuradas, @, !, etc.
    }
    return out;
  }

  // Los tokens a un árbol, con las precedencias de Excel: (de mayor a menor) : espacio , — menos delante — % — ^ — * / — + - — & — comparaciones.
  function analizar(toks) {
    let p = 0, profundidad = 0;
    const ver = () => toks[p], op = (v) => toks[p] && toks[p].k === 'op' && toks[p].v === v;
    const pedir = (v) => { if (!op(v)) no('se esperaba ' + v); p++; };
    const inicioOperando = (t) => t && (t.k === 'ref' || t.k === 'id' || t.k === 'num' || t.k === 'str' || (t.k === 'op' && t.v === '('));
    function cmp() {
      let a = concat();
      while (toks[p] && toks[p].k === 'op' && ['=', '<>', '<', '<=', '>', '>='].includes(toks[p].v)) { const o = toks[p++].v; a = { t: 'bin', op: o, a, b: concat() }; }
      return a;
    }
    function concat() { let a = suma(); while (op('&')) { p++; a = { t: 'bin', op: '&', a, b: suma() }; } return a; }
    function suma() { let a = prod(); while (op('+') || op('-')) { const o = toks[p++].v; a = { t: 'bin', op: o, a, b: prod() }; } return a; }
    function prod() { let a = pot(); while (op('*') || op('/')) { const o = toks[p++].v; a = { t: 'bin', op: o, a, b: pot() }; } return a; }
    function pot() { let a = pct(); while (op('^')) { p++; a = { t: 'bin', op: '^', a, b: pct() }; } return a; }
    function pct() { let a = neg(); while (op('%')) { p++; a = { t: 'un', op: '%', a }; } return a; }
    let signos = 0;   // los signos seguidos (-----1): cada uno es una llamada más
    function neg() {
      if (op('-')) {
        p++;
        if (ver() && ver().k === 'num' && !ver().esp) return { t: 'num', v: -toks[p++].v };   // -1 es un número negativo, no un menos delante del 1
        if (++signos > 200) no('demasiados signos seguidos');
        const a = neg(); signos--;
        return { t: 'un', op: '-', a };
      }
      if (op('+')) {
        p++;
        if (++signos > 200) no('demasiados signos seguidos');
        const a = neg(); signos--;
        return { t: 'un', op: '+', a };
      }
      return inter();
    }
    function inter() {
      let a = rango();
      while (inicioOperando(ver()) && ver().esp) a = { t: 'bin', op: ' ', a, b: rango() };
      return a;
    }
    function rango() { let a = primario(); while (op(':')) { p++; a = { t: 'bin', op: ':', a, b: primario() }; } return a; }
    function primario() {
      if (++profundidad > 60) no('anidado muy hondo');
      const t = toks[p++];
      if (!t) no('fin inesperado');
      let r;
      if (t.k === 'num') r = { t: 'num', v: t.v };
      else if (t.k === 'str') r = { t: 'str', v: t.v };
      else if (t.k === 'err') r = { t: 'err', v: t.v };
      else if (t.k === 'ref') r = { t: 'ref', ...t };
      else if (t.k === 'op' && t.v === '(') {
        let a = cmp();
        while (op(',')) { p++; a = { t: 'bin', op: ',', a, b: cmp() }; }   // una lista entre paréntesis es la unión de referencias
        pedir(')');
        r = { t: 'paren', a };
      } else if (t.k === 'op' && t.v === '{') r = matriz();
      else if (t.k === 'id' && t.fn) {
        pedir('(');
        const args = [];
        if (!op(')')) {
          for (;;) {
            args.push(op(',') || op(')') ? null : cmp());   // un argumento omitido
            if (op(',')) { p++; continue; }
            break;
          }
        }
        pedir(')');
        if (args.length > 255) no('demasiados argumentos');
        r = { t: 'fn', nombre: t.v, args };
      } else if (t.k === 'id') {
        const u = t.v.toUpperCase();
        r = u === 'TRUE' || u === 'FALSE' ? { t: 'bool', v: u === 'TRUE' } : { t: 'nombre', nombre: t.v };
      } else no('token ' + t.v);
      profundidad--;
      return r;
    }
    function matriz() {   // {1,2;3,4}: números (con signo), textos, lógicos y errores
      const filas = [[]];
      for (;;) {
        let neg = false;
        if (op('-')) { p++; neg = true; }
        const t = toks[p++];
        if (!t) no('matriz sin cerrar');
        if (t.k === 'num') filas[filas.length - 1].push({ t: 'num', v: neg ? -t.v : t.v });
        else if (neg) no('menos delante de algo que no es un número');
        else if (t.k === 'str') filas[filas.length - 1].push({ t: 'str', v: t.v });
        else if (t.k === 'err') filas[filas.length - 1].push({ t: 'err', v: t.v });
        else if (t.k === 'id' && /^(?:TRUE|FALSE)$/i.test(t.v)) filas[filas.length - 1].push({ t: 'bool', v: /^true$/i.test(t.v) });
        else no('elemento de matriz');
        if (op(',')) { p++; continue; }
        if (op(';')) { p++; filas.push([]); continue; }
        pedir('}');
        break;
      }
      const ancho = filas[0].length;
      if (filas.some((f) => f.length !== ancho) || ancho > 256 || filas.length > 65536 || ancho * filas.length > 4096) no('matriz irregular o muy grande');
      return { t: 'arr', filas };
    }
    const raiz = cmp();
    if (p !== toks.length) no('sobra texto');
    return raiz;
  }

  const u16b = (n) => [n & 255, (n >> 8) & 255];
  const u32b = (n) => [n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >>> 24) & 255];
  const f64b = (v) => Array.from(new Uint8Array(new Float64Array([v]).buffer));
  const cadenaBytes = (s) => { const ancho = /[^\u0000-ÿ]/.test(s); const b = []; for (let k = 0; k < s.length; k++) { const c = s.charCodeAt(k); b.push(c & 255); if (ancho) b.push(c >> 8); } return { ancho, b }; };

  // Busca una función por su nombre: en la tabla con número, entre las posteriores a 97 (nombre `_xlfn.…`), entre las del complemento, o es una función de usuario.
  const SIN_ZONA = new Set(['FILTER', 'SORT']);   // las únicas que Excel guarda como `_xlfn._xlws.NOMBRE`
  function buscarFuncion(texto) {
    const u = texto.toUpperCase();
    if (!u.startsWith('_XLFN.') && !u.startsWith('_XLWS.')) {
      if (COMPLEMENTO.has(u)) return { def: COMPLEMENTO.get(u), tipo: 'complemento' };   // antes que la tabla: SERIESSUM tiene número en la norma, pero Excel la guarda como de complemento
      if (FUTURA.has(u)) return { def: FUTURA.get(u), tipo: 'futura', nombre: (SIN_ZONA.has(u) ? '_xlfn._xlws.' : '_xlfn.') + u };
      const d = POR_NOMBRE.get(u);
      if (d) return { def: d, tipo: 'ftab' };
      return { def: null, tipo: 'usuario', nombre: texto };
    }
    const base = u.replace(/^_XLFN\./, '').replace(/^_XLWS\./, '');
    if (FUTURA.has(base)) return { def: FUTURA.get(base), tipo: 'futura', nombre: texto };
    return { def: null, tipo: 'usuario', nombre: texto };   // un `_xlfn.` que no conocemos se conserva como llamada por nombre
  }

  // Devuelve { rgce, rgcb, avisos } o, si no se puede, { error: 'motivo' }. Los avisos son lo que cambia al llevar la fórmula a Excel 97-2003:
  //   'futura:NOMBRE' (Excel 97-2003 la deja como #NAME? al recalcular; uno más nuevo la entiende), 'fuera' (una referencia pasaba de 65.536 filas × 256 columnas y quedó #REF!),
  //   'anidado' (más de 7 funciones una dentro de otra: Excel 97-2003 no deja editarla).
  function codificar(texto, ctx, opciones) {
    try { return codif(texto, ctx || {}, (opciones && opciones.clase) || 'V'); } catch (e) { if (e instanceof NoSoportado) return { error: e.message }; throw e; }
  }

  function codif(texto, ctx, claseRaiz) {
    const arbol = analizar(tokenizar(texto.replace(/^=/, '')));
    const E = { ctx, rgcb: [], volatil: false, tokens: 0, prof: 0, nivel: 0, avisos: new Set() };
    const rgce = emitir(arbol, claseRaiz, E, true);
    const cabeza = E.volatil ? [0x19, 0x01, 0, 0] : [];
    const todo = cabeza.concat(rgce);
    if (todo.length > MAX_RGCE || E.rgcb.length > MAX_RGCE) no('fórmula muy larga');
    return { rgce: Uint8Array.from(todo), rgcb: Uint8Array.from(E.rgcb), avisos: [...E.avisos] };
  }

  // Una referencia (celda, área, columnas o filas enteras) a bytes. Fuera de la cuadrícula del .xls es #REF!.
  function emitirRef(n, cls, E) {
    let ixti = null;
    if (n.hoja) {
      const a = E.ctx.hoja && E.ctx.hoja(n.hoja.a), b = E.ctx.hoja && E.ctx.hoja(n.hoja.b);
      if (a === undefined || b === undefined || a < 0 || b < 0) no('hoja que no existe: ' + n.hoja.a);
      ixti = E.ctx.ixti(Math.min(a, b), Math.max(a, b));
    }
    const pre = ixti === null ? [] : u16b(ixti);
    const R = n.hoja ? 0x1A : 0x04, A = n.hoja ? 0x1B : 0x05;
    if (n.modo === 'error') return n.hoja ? [0x1C | CLASE.R, ...pre, 0, 0, 0, 0xC0] : [0x0A | CLASE[cls], 0, 0, 0, 0];   // con hoja Excel lo guarda siempre como referencia y con la marca C0
    const fuera = (f, c) => f < 0 || f > 65535 || c < 0 || c > 255;
    if (n.modo === 'celda') {
      if (fuera(n.f0, n.c0)) { E.avisos.add('fuera'); return [(n.hoja ? 0x1C : 0x0A) | CLASE.R, ...pre, 0, 0, 0, 0xC0]; }
      return [R | CLASE[cls], ...pre, ...u16b(n.f0), ...u16b(n.c0 | (n.ca0 ? 0 : 0x4000) | (n.fa0 ? 0 : 0x8000))];
    }
    let f0, f1, c0, c1, fr0, fr1, cr0, cr1;
    if (n.modo === 'area') { f0 = n.f0; f1 = n.f1; c0 = n.c0; c1 = n.c1; fr0 = !n.fa0; fr1 = !n.fa1; cr0 = !n.ca0; cr1 = !n.ca1; }
    else if (n.modo === 'cols') { f0 = 0; f1 = 65535; c0 = n.c0; c1 = n.c1; fr0 = false; fr1 = false; cr0 = !n.ca0; cr1 = !n.ca1; }
    else { f0 = n.f0; f1 = n.f1; c0 = 0; c1 = 255; fr0 = !n.fa0; fr1 = !n.fa1; cr0 = false; cr1 = false; }
    if (fuera(f0, c0) || fuera(f1, c1)) { E.avisos.add('fuera'); return [(n.hoja ? 0x1D : 0x0B) | CLASE.R, ...pre, 0, 0, 0, 0, 0, 0xC0, 0, 0xC0]; }
    if (f0 > f1) { [f0, f1] = [f1, f0]; [fr0, fr1] = [fr1, fr0]; }
    if (c0 > c1) { [c0, c1] = [c1, c0]; [cr0, cr1] = [cr1, cr0]; }
    return [A | CLASE[cls], ...pre, ...u16b(f0), ...u16b(f1), ...u16b(c0 | (cr0 ? 0x4000 : 0) | (fr0 ? 0x8000 : 0)), ...u16b(c1 | (cr1 ? 0x4000 : 0) | (fr1 ? 0x8000 : 0))];
  }

  const CODIGO_BIN = { '+': 0x03, '-': 0x04, '*': 0x05, '/': 0x06, '^': 0x07, '&': 0x08, '<': 0x09, '<=': 0x0A, '=': 0x0B, '>=': 0x0C, '>': 0x0D, '<>': 0x0E, ' ': 0x0F, ',': 0x10, ':': 0x11 };

  // `raiz`: es la fórmula entera. Una constante de matriz sola lleva la clase de la fórmula; dentro de otra cosa siempre es de clase matriz.
  function emitir(n, cls, E, raiz) {
    if (++E.tokens > 4096) no('fórmula muy larga');
    if (E.prof >= MAX_PROFUNDIDAD) no('fórmula muy honda');
    E.prof++;
    try { return emitir2(n, cls, E, raiz); } finally { E.prof--; }
  }
  function emitir2(n, cls, E, raiz) {
    switch (n.t) {
      case 'num': {
        if (Number.isInteger(n.v) && n.v >= 0 && n.v <= 65535 && !Object.is(n.v, -0)) return [0x1E, ...u16b(n.v)];
        if (!Number.isFinite(n.v)) no('número no finito');
        return [0x1F, ...f64b(n.v)];
      }
      case 'str': {
        if (n.v.length > MAX_TEXTO_LIT) no('texto de más de 255 caracteres');
        const s = cadenaBytes(n.v);
        return [0x17, n.v.length, s.ancho ? 1 : 0, ...s.b];
      }
      case 'bool': return [0x1D, n.v ? 1 : 0];
      case 'err': return n.v === '#REF!' ? [0x0A | CLASE[cls], 0, 0, 0, 0] : [0x1C, ERROR_A_CODIGO[n.v]];
      case 'arr': {
        const filas = n.filas.length, cols = n.filas[0].length;
        E.rgcb.push(cols - 1, ...u16b(filas - 1));
        for (const fila of n.filas) for (const v of fila) {
          if (v.t === 'num') E.rgcb.push(0x01, ...f64b(v.v));
          else if (v.t === 'str') { const s = cadenaBytes(v.v); E.rgcb.push(0x02, ...u16b(v.v.length), s.ancho ? 1 : 0, ...s.b); }
          else if (v.t === 'bool') E.rgcb.push(0x04, v.v ? 1 : 0, 0, 0, 0, 0, 0, 0);
          else E.rgcb.push(0x10, ERROR_A_CODIGO[v.v], 0, 0, 0, 0, 0, 0);
        }
        return [raiz ? CLASE[cls] : 0x60, cols - 1, ...u16b(filas - 1), 0, 0, 0, 0];
      }
      case 'ref': return emitirRef(n, cls, E);
      case 'nombre': {
        const idx = E.ctx.nombre && E.ctx.nombre(n.nombre);
        if (!idx) no('nombre que el libro no define: ' + n.nombre);
        return [0x03 | CLASE[cls], ...u32b(idx)];
      }
      case 'paren': return [...emitir(n.a, cls, E, false), 0x15];
      case 'un': {
        const c = cls === 'A' ? 'A' : 'V';
        return [...emitir(n.a, c, E, false), n.op === '-' ? 0x13 : n.op === '+' ? 0x12 : 0x14];
      }
      case 'bin': {
        const c = n.op === ':' || n.op === ' ' || n.op === ',' ? 'R' : cls === 'A' ? 'A' : 'V';
        return [...emitir(n.a, c, E, false), ...emitir(n.b, c, E, false), CODIGO_BIN[n.op]];
      }
      case 'fn': return emitirFuncion(n, cls, E);
      default: no('nodo ' + n.t);
    }
  }

  function emitirFuncion(n, cls, E) {
    const b = buscarFuncion(n.nombre), d = b.def, argc = n.args.length;
    if (argc > 30) no('más de 30 argumentos');
    if (d) {
      if (argc < d.min || argc > d.max) no('cantidad de argumentos de ' + d.nombre);
      if (d.vol) E.volatil = true;
      if (b.tipo === 'futura') E.avisos.add('futura:' + d.nombre);
    }
    if (++E.nivel > 7) E.avisos.add('anidado');
    try { return emitirFuncion2(n, cls, E, b, d, argc); } finally { E.nivel--; }
  }

  function emitirFuncion2(n, cls, E, b, d, argc) {
    // dentro de una matriz, los argumentos de valor de una función también se calculan como matriz (SUMPRODUCT(ABS(A1:A3),1) evalúa ABS celda por celda)
    const arg = (i) => { if (n.args[i] === null) return [0x16]; const c = claseDeArg(d, i); return emitir(n.args[i], cls === 'A' && c === 'V' && d && d.lift ? 'A' : c, E, false); };
    // la clase de la función misma: la que pide quien la usa; si pide una referencia, la que la función puede dar (INDEX, OFFSET: referencia; BYCOL: matriz; el resto: valor)
    const clase = cls === 'R' ? (d ? d.ret || 'V' : 'R') : cls;   // una función de usuario puede devolver cualquier cosa: se le da lo que se pida
    if (b.tipo === 'ftab') {
      if (d.id === 1) {   // IF: la condición, un salto si es falsa, la rama verdadera, un salto al final, la rama falsa y otro salto
        const cond = arg(0), si = arg(1), no2 = argc === 3 ? arg(2) : null;
        const ptgFunc = [0x02 | CLASE[clase], argc, ...u16b(1)];
        if (no2) return [...cond, 0x19, 0x02, ...u16b(si.length + 4), ...si, 0x19, 0x08, ...u16b(no2.length + 4 + ptgFunc.length - 1), ...no2, 0x19, 0x08, ...u16b(ptgFunc.length - 1), ...ptgFunc];
        return [...cond, 0x19, 0x02, ...u16b(si.length + 4), ...si, 0x19, 0x08, ...u16b(ptgFunc.length - 1), ...ptgFunc];
      }
      if (d.id === 100) {   // CHOOSE: el índice, una tabla de saltos y cada valor seguido de un salto al final
        const indice = arg(0), valores = []; for (let i = 1; i < argc; i++) valores.push(arg(i));
        const ptgFunc = [0x02 | CLASE[clase], argc, ...u16b(100)];
        const tabla = 2 * (valores.length + 1);
        const offs = []; let acum = tabla;
        for (const v of valores) { offs.push(acum); acum += v.length + 4; }
        offs.push(acum);
        let cuerpo = [];
        valores.forEach((v, i) => {
          let resto = ptgFunc.length - 1; for (let j = i + 1; j < valores.length; j++) resto += valores[j].length + 4;
          cuerpo = cuerpo.concat(v, [0x19, 0x08, ...u16b(resto)]);
        });
        return [...indice, 0x19, 0x04, ...u16b(valores.length), ...offs.flatMap(u16b), ...cuerpo, ...ptgFunc];
      }
      if (d.id === 4 && argc === 1) return [...arg(0), 0x19, 0x10, 0, 0];   // SUM de un solo argumento: PtgAttrSum
      const args = []; for (let i = 0; i < argc; i++) args.push(...arg(i));
      if (d.fija) return [...args, 0x01 | CLASE[clase], ...u16b(d.id)];
      return [...args, 0x02 | CLASE[clase], argc, ...u16b(d.id)];
    }
    // sin número en la tabla: una llamada por nombre. El nombre va primero (como referencia), después los argumentos y al final la función 255 con un argumento de más.
    let nombre;
    if (b.tipo === 'complemento') {
      const c = E.ctx.complemento && E.ctx.complemento(d);
      if (!c) no('el libro no admite funciones de complemento');
      nombre = [0x39, ...u16b(c.ixti), ...u32b(c.idx)];
    } else {
      const idx = E.ctx.nombre && E.ctx.nombre(b.nombre, true);
      if (!idx) no('función que el libro no define: ' + n.nombre);
      nombre = [0x23, ...u32b(idx)];
    }
    const args = []; for (let i = 0; i < argc; i++) args.push(...arg(i));
    return [...nombre, ...args, 0x02 | CLASE[clase], argc + 1, 0xFF, 0x00];
  }

  globalThis.XlsFormulas = { decodificar, codificar, FUNCIONES, _numeroATexto: numeroATexto, _citarHoja: citarHoja, _tokenizar: tokenizar, _analizar: analizar };
})();

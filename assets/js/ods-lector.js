/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Lector de hojas de cálculo OpenDocument (.ods, de LibreOffice, Google Sheets y Excel) del motor propio de Excel (etapa 9). Un .ods es un zip con
// content.xml (las hojas) y styles.xml (los estilos con nombre). Requiere zip-propio.js, xml-ligero.js y xlsx-modelo.js (para los errores); no depende del navegador.
//
//   const libro = await OdsLector.leer(bytes);
//   libro = { hojas: [{ nombre, filas: [[valor…]…], anchos: [caracteres…], combinadas: ['A1:B2'…], estado, columnasOcultas, filasOcultas, altosFila: { fila desde 0: puntos } }] }
//   valor: undefined (vacía) | número | texto | lógico | { error } | { valor, formato?, estilo? }   (lo que entiende XlsxEscritor.crearLibro; una celda vacía
//   con relleno o bordes es { valor: null, estilo })
//
// Qué lee: valores (número, porcentaje, moneda, fecha, hora, lógico, texto, error), el formato de número de cada celda (traducido a un código de Excel),
// combinadas, anchos de columna y el idioma-neutro de los textos con espacios y saltos. Las fórmulas solo se leen con leer(bytes, { formulas: true }), traducidas
// por ods-formulas.js (la sintaxis de ODF — "of:=[.A1]+1" — no es la de Excel); si no, queda el último valor, que es lo que el archivo guarda. También lleva el estilo de cada celda (fuente, color, relleno,
// bordes y alineación, con la herencia de style:parent-style-name y el estilo por defecto de columna o fila) y el alto propio de las filas
// (style:row-height cuando no es "óptimo"). Hay un objeto de estilo por cada combinación distinta: las celdas que la comparten apuntan al mismo, que es lo
// que espera crearLibro. Lo que Excel no tiene (reducir hasta ajustar, los colores de un formato de número) no se lleva.
(function () {
  const MAX_FILAS = 1048576, MAX_COLS = 16384, MAX_TEXTO = 32767;
  const MAX_CELDAS = 10000000;    // en todo el libro: crearLibro no arma un .xlsx de más de ~8 M celdas (una fila repetida un millón de veces cabe en 200 bytes)
  const MAX_COMBINADAS = 500000;  // por hoja, como LibroXlsx
  const MAX_ESTILOS = 100000;     // estilos de celda, de fila y de columna que se guardan (un archivo hostil puede declarar millones)
  const MAX_VACIAS = 1000000;     // celdas vacías con relleno o bordes en todo el libro (una fila de relleno repetida un millón de veces no se materializa)
  const MAX_ALTOS = 100000;       // filas con alto propio por hoja
  // Lugares de fila que se arman en todo el libro detrás de la última celda con contenido (lo que mide cada fila, con sus huecos, por las veces que se repite):
  // una fila de 16.384 con una vacía con relleno al final, repetida, armaba cien millones con 550 bytes. Lo que pasa del tope pierde las vacías con estilo.
  const MAX_RANURAS = MAX_CELDAS;
  const MAX_HOJAS = 4096, MAX_FORMULA = 32768;   // tablas que se leen; largo de una fórmula de OpenFormula que se traduce (la de Excel no pasa de 8.192)
  const MAX_DIGITOS = 30;         // cifras de un formato de número (decimales, enteros mínimos): Excel no muestra más de 30 decimales
  const SIN_ESTILO = { est: null, cod: null };
  const MAX_FUSIONES = 500;   // combinadas con estilo que se siguen a la vez (las que aún cubren la fila de abajo)
  const seVe = (e) => !!e && !!(e.relleno || e.bordes);
  const ERRORES = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA']);
  const ERRORES_ES = { '#¡NULO!': '#NULL!', '#¡DIV/0!': '#DIV/0!', '#¡VALOR!': '#VALUE!', '#¡REF!': '#REF!', '#¿NOMBRE?': '#NAME?', '#¡NUM!': '#NUM!', '#N/D': '#N/A' };   // un LibreOffice en español los guarda así

  const ent = (a, k) => (a[k] === undefined ? null : a[k]);
  const entero = (s, defecto) => { const n = parseInt(s, 10); return n > 0 ? n : defecto; };
  const cifras = (s, defecto) => Math.min(entero(s, defecto), MAX_DIGITOS);
  const comillas = (t) => '"' + t.replace(/"/g, '""') + '"';

  // ---- fechas y horas de ODF → número de serie de Excel (base 1900) -----------------------------------------------
  function serieDeFecha(texto) {
    const m = /^(-?\d{1,6})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2}(?:\.\d+)?))?/.exec(texto);
    if (!m) return null;
    const a = +m[1], mes = +m[2], d = +m[3];
    let dias;
    if (a === 1900 && mes === 2 && d === 29) dias = 60;   // el día que Excel inventó
    else {
      const t = new Date(0); t.setUTCFullYear(a, mes - 1, d); t.setUTCHours(0, 0, 0, 0);
      dias = Math.round(t.getTime() / 86400000) + 25569;   // días desde 1899-12-30
      if (dias < 61) dias -= 1;                             // antes del 1900-03-01 Excel cuenta un día menos
    }
    const hora = m[4] === undefined ? 0 : (+m[4] * 3600 + +m[5] * 60 + parseFloat(m[6])) / 86400;
    return { serie: dias + hora, conHora: m[4] !== undefined && (+m[4] || +m[5] || +m[6]) };
  }
  function serieDeHora(texto) {
    const m = /^(-)?P(?:(\d+)D)?(?:T(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?)?$/.exec(texto);
    if (!m) return null;
    const s = (+(m[2] || 0)) * 86400 + (+(m[3] || 0)) * 3600 + (+(m[4] || 0)) * 60 + +(m[5] || 0);
    return m[1] ? -s / 86400 : s / 86400;
  }

  // ---- ancho de columna ---------------------------------------------------------------------------------------------
  function caracteresDeAncho(texto) {
    const m = /^([\d.]+)\s*(cm|mm|in|pt|px)?$/.exec(texto || '');
    if (!m) return 0;
    const n = parseFloat(m[1]), u = m[2] || 'cm';
    const px = u === 'cm' ? n / 2.54 * 96 : u === 'mm' ? n / 25.4 * 96 : u === 'in' ? n * 96 : u === 'pt' ? n * 96 / 72 : n;
    return px > 5 ? Math.round((px - 5) / 7 * 100) / 100 : 0;   // Calibri 11: 7 px por carácter y 5 de margen
  }

  // ---- formatos de número de ODF → código de Excel -------------------------------------------------------------
  // estilo = { tipo, partes: [{ k, a }], mapas: [{ cond, estilo }] }; se traduce al final, cuando ya están todos.
  // El estilo al que lleva un style:map se traduce sin sus propios mapas: Excel admite dos condiciones y una sección más, y una cadena hostil de miles
  // de mapas desbordaba la pila.
  function traducirEstilo(estilos, nombre, vistos, sinMapas) {
    const e = estilos.datos.get(nombre);
    if (!e) return null;
    vistos = vistos || new Set();
    if (vistos.has(nombre)) return null;
    vistos.add(nombre);
    const porcentaje = e.tipo === 'percentage-style';
    let s = '';
    for (const p of e.partes) {
      const a = p.a;
      switch (p.k) {
        case 'text': s += porcentaje && p.t === '%' ? '%' : (p.t ? (/^[-\/:. ,]+$/.test(p.t) && /date|time/.test(e.tipo) ? p.t : comillas(p.t)) : ''); break;
        case 'currency-symbol': s += p.t ? comillas(p.t) : ''; break;
        case 'text-content': s += '@'; break;
        case 'fill-character': s += p.t ? '*' + p.t.charAt(0) : ''; break;
        case 'number': {
          const dp = a['decimal-places'] === undefined ? null : cifras(a['decimal-places'], 0);
          const minDec = a['min-decimal-places'] === undefined ? dp : cifras(a['min-decimal-places'], 0);
          const minInt = Math.max(1, cifras(a['min-integer-digits'], 1));
          if (dp === null && a.grouping !== 'true' && minInt === 1) { s += 'General'; break; }
          s += a.grouping === 'true' ? '#,' + '#'.repeat(Math.max(0, 3 - minInt)) + '0'.repeat(minInt) : '0'.repeat(minInt);
          if (dp) s += '.' + '0'.repeat(minDec || 0) + '#'.repeat(Math.max(0, dp - (minDec || 0)));
          break;
        }
        case 'scientific-number': {
          const dp = cifras(a['decimal-places'], 0), minDec = a['min-decimal-places'] === undefined ? dp : cifras(a['min-decimal-places'], 0);
          s += '0' + (dp ? '.' + '0'.repeat(minDec) + '#'.repeat(Math.max(0, dp - minDec)) : '') + 'E' + (a['forced-exponent-sign'] === 'true' ? '+' : '-') + '0'.repeat(cifras(a['min-exponent-digits'], 1));
          break;
        }
        case 'fraction': {
          const fijo = /^\d{1,9}$/.test(a['denominator-value'] || '') ? a['denominator-value'] : null, max = entero(a['max-denominator-value'], 0);
          const dig = fijo ? 0 : Math.min(MAX_DIGITOS, Math.max(cifras(a['min-denominator-digits'], 1), String(max || 9).length));
          s += (entero(a['min-integer-digits'], 0) > 0 ? '0 ' : '# ') + '?'.repeat(Math.max(1, cifras(a['min-numerator-digits'], 1))) + '/' + (fijo ? fijo : '?'.repeat(dig));
          break;
        }
        case 'day': s += a.style === 'long' ? 'dd' : 'd'; break;
        case 'day-of-week': s += a.style === 'long' ? 'dddd' : 'ddd'; break;
        case 'month': s += a.textual === 'true' ? (a.style === 'long' ? 'mmmm' : 'mmm') : (a.style === 'long' ? 'mm' : 'm'); break;
        case 'year': s += a.style === 'long' ? 'yyyy' : 'yy'; break;
        case 'quarter': s += a.style === 'long' ? '"Q"0' : '"Q"0'; break;
        case 'hours': s += e.a['truncate-on-overflow'] === 'false' ? '[h]' : (a.style === 'long' ? 'hh' : 'h'); break;
        case 'minutes': s += e.a['truncate-on-overflow'] === 'false' && !e.partes.some((q) => q.k === 'hours') ? '[mm]' : (a.style === 'long' ? 'mm' : 'm'); break;
        case 'seconds': s += (e.a['truncate-on-overflow'] === 'false' && !e.partes.some((q) => q.k === 'hours' || q.k === 'minutes') ? '[ss]' : (a.style === 'long' ? 'ss' : 's')) + (cifras(a['decimal-places'], 0) ? '.' + '0'.repeat(cifras(a['decimal-places'], 0)) : ''); break;
        case 'am-pm': s += 'AM/PM'; break;
        default: break;
      }
    }
    if (porcentaje && !/%/.test(s)) s += '%';
    if (e.tipo === 'boolean-style') s = 'General';
    if (e.tipo === 'text-style' && !s) s = '@';
    // style:map: "si el valor cumple la condición, usa el otro estilo"; Excel lo expresa como [condición]formato1;formato2
    const secciones = [];
    for (const m of sinMapas ? [] : e.mapas.slice(0, 3)) {
      const c = /^value\(\)\s*(&gt;=|&lt;=|>=|<=|!=|<>|=|>|<)\s*(-?[\d.]+)$/.exec(m.cond.replace(/&gt;/g, '>').replace(/&lt;/g, '<'));
      const otro = c ? traducirEstilo(estilos, m.estilo, vistos, true) : null;
      // "valor >= 0" es el reparto normal de Excel (positivos;negativos): la segunda sección muestra el valor absoluto, y el estilo base ya trae su "-"
      if (c && otro) secciones.push((c[1] === '>=' && c[2] === '0' ? '' : '[' + (c[1] === '!=' ? '<>' : c[1]) + c[2] + ']') + otro);
    }
    return secciones.length ? secciones.join(';') + ';' + s : (s || null);
  }

  // ---- estilo de celda: de las propiedades de ODF al estilo de crearLibro ----------------------------------------------
  // Una longitud de ODF ("0.74pt", "1.059cm", "15.75pt") a puntos; null si no es una medida.
  function puntosDe(texto) {
    const m = /^(-?[\d.]+)\s*(cm|mm|in|pt|px|pc)?$/.exec(texto || '');
    if (!m) return null;
    const n = parseFloat(m[1]), u = m[2] || 'pt';
    return u === 'cm' ? n * 72 / 2.54 : u === 'mm' ? n * 72 / 25.4 : u === 'in' ? n * 72 : u === 'px' ? n * 0.75 : u === 'pc' ? n * 12 : n;
  }
  // De un elemento de propiedades se guarda solo lo que se usa, con claves cortas; lo definido más tarde (el hijo) pisa lo anterior.
  function capturarPropiedades(p, elemento, a) {
    if (elemento === 'text-properties') {
      if (a['font-name'] !== undefined) p.fn = a['font-name'];
      if (a['font-family'] !== undefined) p.ff = a['font-family'];
      if (a['font-size'] !== undefined) p.fs = a['font-size'];
      if (a['font-weight'] !== undefined) p.fw = a['font-weight'];
      if (a['font-style'] !== undefined) p.fi = a['font-style'];
      if (a['text-underline-style'] !== undefined) p.us = a['text-underline-style'];
      if (a['text-underline-type'] !== undefined) p.ut = a['text-underline-type'];
      if (a['text-line-through-style'] !== undefined) p.lt = a['text-line-through-style'];
      if (a.color !== undefined) p.co = a.color;
    } else if (elemento === 'table-cell-properties') {
      if (a['background-color'] !== undefined) p.bg = a['background-color'];
      if (a.border !== undefined) p.bi = p.bd = p.ba = p.bb = a.border;   // el general primero y los de un lado después
      if (a['border-left'] !== undefined) p.bi = a['border-left'];
      if (a['border-right'] !== undefined) p.bd = a['border-right'];
      if (a['border-top'] !== undefined) p.ba = a['border-top'];
      if (a['border-bottom'] !== undefined) p.bb = a['border-bottom'];
      if (a['diagonal-tl-br'] !== undefined) p.dA = a['diagonal-tl-br'];
      if (a['diagonal-bl-tr'] !== undefined) p.dB = a['diagonal-bl-tr'];
      if (a['wrap-option'] !== undefined) p.aj = a['wrap-option'];
      if (a['rotation-angle'] !== undefined) p.ro = a['rotation-angle'];
      if (a.direction !== undefined) p.di = a.direction;
      if (a['vertical-align'] !== undefined) p.va = a['vertical-align'];
      if (a['repeat-content'] !== undefined) p.rc = a['repeat-content'];
      if (a['text-align-source'] !== undefined) p.tas = a['text-align-source'];
    } else {
      if (a['text-align'] !== undefined) p.ta = a['text-align'];
      if (a['margin-left'] !== undefined) p.ml = a['margin-left'];
      if (a['margin-right'] !== undefined) p.mr = a['margin-right'];
    }
  }
  // "0.74pt solid #000000" → { e, c } de crearLibro; null si no hay borde. LibreOffice usa 0.05 / 0.74 / 1.76 / 2.49 pt para hair / thin / medium / thick,
  // y Excel escribe thin, medium o thick; las líneas dobles, de puntos y de rayas se reconocen por su palabra.
  function bordeDe(texto) {
    if (!texto || texto === 'none' || texto === 'hidden') return null;
    let ancho = null, forma = 'solid', color;
    for (const t of texto.trim().split(/\s+/).slice(0, 4)) {
      if (t.charAt(0) === '#') { if (/^#[0-9a-fA-F]{6}$/.test(t)) color = t.slice(1).toUpperCase(); }
      else if (t === 'thin') ancho = 0.75;
      else if (t === 'medium') ancho = 1.75;
      else if (t === 'thick') ancho = 2.5;
      else if (/^(solid|dotted|dashed|double|dot-dash|dot-dot-dash|none|hidden)$/.test(t)) forma = t;
      else if (/^double-thin$/.test(t)) forma = 'double';          // LibreOffice: «1.76pt double-thin» (Excel la lee como doble)
      else if (/^(fine-dashed|long-dash|dash-dot|dash-dot-dot|dash|ultrafine-dashed)$/.test(t)) forma = /dot-dot/.test(t) ? 'dot-dot-dash' : /dot/.test(t) ? 'dot-dash' : 'dashed';   // LibreOffice
      else { const pt = puntosDe(t); if (pt !== null) ancho = pt; }
    }
    if (forma === 'none' || forma === 'hidden') return null;
    if (ancho === null) ancho = 0.75;
    const grueso = ancho >= 1.2;
    const grosor = ancho <= 0.4 ? 'hair' : ancho < 1.2 ? 'thin' : ancho < 2.5 ? 'medium' : 'thick';   // 2.49 pt (el «grueso» de LibreOffice) lo lee Excel como medio: medido con LibreOffice 26.8
    let e;
    if (forma === 'double') e = 'double';
    else if (forma === 'dotted') e = 'dotted';
    else if (forma === 'dashed') e = grueso ? 'mediumDashed' : 'dashed';
    else if (forma === 'dot-dash') e = grueso ? 'mediumDashDot' : 'dashDot';
    else if (forma === 'dot-dot-dash') e = grueso ? 'mediumDashDotDot' : 'dashDotDot';
    else e = grosor;
    return color ? { e, c: color } : { e };
  }
  const familiaDe = (t) => { const f = String(t).split(',')[0].trim().replace(/^["']+|["']+$/g, ''); return f || null; };

  // El estilo de una celda con nombre (null si queda como la celda de siempre): la cadena de style:parent-style-name, de la raíz al hijo, sobre el estilo por
  // defecto de las celdas. La cadena se corta a 30 eslabones y no repite nombres (un ciclo en un archivo hostil no la cuelga).
  function estiloDeCelda(estilos, nombre, memo) {
    if (!nombre) return null;
    if (memo.has(nombre)) return memo.get(nombre);
    const cadena = [], vistos = new Set();
    for (let n = nombre; n && !vistos.has(n) && cadena.length < 30;) {
      vistos.add(n);
      const e = estilos.celdas.get(n);
      if (!e) break;
      cadena.push(e);
      n = e.padre;
    }
    const r = {};
    if (estilos.porDefecto) Object.assign(r, estilos.porDefecto.p);
    for (let k = cadena.length - 1; k >= 0; k--) Object.assign(r, cadena[k].p);
    const resultado = construirEstilo(estilos, r);
    memo.set(nombre, resultado);
    return resultado;
  }
  // Los estilos de iguales propiedades comparten el objeto (crearLibro reparte los iguales por identidad, sin volver a mirarlos).
  function construirEstilo(estilos, r) {
    const e = {}, f = {};
    const nombreFuente = r.fn !== undefined ? familiaDe(estilos.fuentes.get(r.fn) || r.fn) : r.ff !== undefined ? familiaDe(r.ff) : null;
    if (nombreFuente && nombreFuente !== 'Calibri') f.nombre = nombreFuente;
    const tam = r.fs !== undefined ? puntosDe(r.fs) : null;
    if (tam && tam > 0 && Math.abs(tam - 11) > 0.001) f.tam = tam;
    if (r.fw === 'bold' || parseInt(r.fw, 10) >= 600) f.negrita = true;
    if (r.fi === 'italic' || r.fi === 'oblique') f.cursiva = true;
    if (r.us !== undefined && r.us !== 'none') f.subrayado = r.ut === 'double' ? 'double' : true;
    if (r.lt !== undefined && r.lt !== 'none') f.tachado = true;
    if (typeof r.co === 'string' && /^#[0-9a-fA-F]{6}$/.test(r.co) && r.co.toUpperCase() !== '#000000') f.color = r.co.slice(1).toUpperCase();
    if (Object.keys(f).length) e.fuente = f;
    if (typeof r.bg === 'string' && /^#[0-9a-fA-F]{6}$/.test(r.bg)) e.relleno = r.bg.slice(1).toUpperCase();
    const bordes = {};
    for (const [clave, lado] of [['bi', 'izq'], ['bd', 'der'], ['ba', 'arr'], ['bb', 'aba'], ['dA', 'diagA'], ['dB', 'diagB']]) { const b = bordeDe(r[clave]); if (b) bordes[lado] = b; }
    if (Object.keys(bordes).length) e.bordes = bordes;
    const a = {};
    let h = r.tas === 'value-type' ? null : r.ta === 'start' || r.ta === 'left' ? 'left' : r.ta === 'end' || r.ta === 'right' ? 'right' : r.ta === 'center' ? 'center' : r.ta === 'justify' ? 'justify' : null;
    if (r.rc === 'true') h = 'fill';
    if (h) a.h = h;
    if (r.va === 'top') a.v = 'top'; else if (r.va === 'middle') a.v = 'center';
    if (r.aj === 'wrap') a.ajustar = true;
    const sangria = puntosDe(h === 'right' ? r.mr : r.ml);
    if (sangria && sangria >= 5) a.sangria = Math.min(250, Math.round(sangria / 10));   // un nivel de sangría de Excel son 10 pt
    const ang = r.ro === undefined ? 0 : ((Math.round(parseFloat(r.ro)) % 360) + 360) % 360;
    if (ang >= 1 && ang <= 90) a.rotacion = ang;
    else if (ang >= 270) a.rotacion = 90 + (360 - ang);   // Excel cuenta el giro horario del 91 al 180
    else if (r.di === 'ttb') a.rotacion = 255;            // texto apilado
    if (Object.keys(a).length) e.alinea = a;
    if (!Object.keys(e).length) return null;
    const clave = JSON.stringify(e);
    if (estilos.unicos.has(clave)) return estilos.unicos.get(clave);
    estilos.unicos.set(clave, e);
    return e;
  }

  // ---- lectura -----------------------------------------------------------------------------------------------------
  function leerEstilos(texto, estilos) {
    // estilos.datos: nombre → { tipo, partes, mapas };  estilos.celdas: nombre → { dato, padre };  estilos.columnas: nombre → ancho;
    // estilos.hojasOcultas: nombre del estilo de tabla → true
    let actual = null, parteTexto = null;
    const TIPOS = new Set(['number-style', 'percentage-style', 'currency-style', 'date-style', 'time-style', 'boolean-style', 'text-style']);
    const PARTES = new Set(['number', 'scientific-number', 'fraction', 'day', 'day-of-week', 'month', 'year', 'quarter', 'hours', 'minutes', 'seconds', 'am-pm', 'text-content', 'fill-character']);
    XmlLigero.recorrer(texto, {
      abrir(nombre, a) {
        if (TIPOS.has(nombre)) { actual = { tipo: nombre, partes: [], mapas: [], a }; estilos.datos.set(a.name, actual); return; }
        if (actual) {
          if (nombre === 'text' || nombre === 'currency-symbol' || nombre === 'fill-character') parteTexto = { k: nombre, a, t: '' };
          else if (nombre === 'map') actual.mapas.push({ cond: a.condition || '', estilo: a['apply-style-name'] });
          else if (PARTES.has(nombre)) actual.partes.push({ k: nombre, a });
          return;
        }
        if (nombre === 'font-face' && a.name !== undefined && a['font-family'] !== undefined) { if (estilos.fuentes.size < MAX_ESTILOS) estilos.fuentes.set(a.name, a['font-family']); return; }
        if ((nombre === 'style' && a.name !== undefined) || (nombre === 'default-style' && a.family === 'table-cell')) {
          estilos._celda = null;
          if (nombre === 'default-style') { estilos._celda = estilos.porDefecto = { dato: null, padre: null, p: {} }; return; }
          if (a.family === 'table-cell' && estilos.celdas.size < MAX_ESTILOS) {
            estilos._celda = { dato: a['data-style-name'] || null, padre: a['parent-style-name'] || null, p: {} };
            estilos.celdas.set(a.name, estilos._celda);
          }
          estilos._nombre = a.name;
        } else if (nombre === 'table-column-properties' && a['column-width'] !== undefined) { if (estilos.columnas.size < MAX_ESTILOS) estilos.columnas.set(estilos._nombre, caracteresDeAncho(a['column-width'])); }
        else if (nombre === 'table-row-properties' && a['row-height'] !== undefined) {
          const pt = puntosDe(a['row-height']);
          if (pt > 0 && estilos.filas.size < MAX_ESTILOS) estilos.filas.set(estilos._nombre, { alto: pt, optimo: a['use-optimal-row-height'] === 'true' });
        } else if (nombre === 'table-properties' && a.display === 'false') estilos.hojasOcultas.add(estilos._nombre);
        else if (estilos._celda && (nombre === 'text-properties' || nombre === 'table-cell-properties' || nombre === 'paragraph-properties')) capturarPropiedades(estilos._celda.p, nombre, a);
      },
      cerrar(nombre) {
        if (nombre === 'style' || nombre === 'default-style') estilos._celda = null;
        if (TIPOS.has(nombre)) actual = null;
        else if (parteTexto && nombre === parteTexto.k) { actual.partes.push(parteTexto); parteTexto = null; }
      },
      texto(t) { if (parteTexto) parteTexto.t += t; },
    });
  }

  function formatoDeCelda(estilos, nombre, memo) {
    if (!nombre) return null;
    if (memo.has(nombre)) return memo.get(nombre);
    let dato = null, n = nombre, saltos = 0;
    while (n && saltos++ < 20) {
      const e = estilos.celdas.get(n);
      if (!e) break;
      if (e.dato) { dato = e.dato; break; }
      n = e.padre;
    }
    const codigo = dato ? traducirEstilo(estilos, dato) : null;
    memo.set(nombre, codigo);
    return codigo;
  }

  async function leer(bytes, opciones) {
    const conFormulas = !!(opciones && opciones.formulas) && typeof OdsFormulas !== 'undefined';
    let zip;
    try { zip = await ZipPropio.loadAsync(bytes); } catch (e) { throw new LibroXlsx.ErrorXlsx('noZip', 'no es OpenDocument'); }
    const contenido = zip.file('content.xml');
    // LibreOffice 26 (ODF 1.4) cifra el paquete entero: un solo archivo «encrypted-package» y ningún content.xml a la vista (medido con LibreOffice 26.8, Argon2id + AES-256-GCM)
    if (!contenido && zip.file('encrypted-package')) throw new LibroXlsx.ErrorXlsx('cifrado');
    if (!contenido) throw new LibroXlsx.ErrorXlsx('noZip', 'no es OpenDocument');
    if (zip.file('META-INF/manifest.xml')) {
      const man = await zip.file('META-INF/manifest.xml').async('string');
      // con cualquier prefijo (o ninguno) del espacio de nombres del manifiesto: LibreOffice usa "manifest:", otros programas no siempre
      if (/<(?:[\w.-]+:)?encryption-data\b/.test(man)) throw new LibroXlsx.ErrorXlsx('cifrado');
    }
    const estilos = { datos: new Map(), celdas: new Map(), columnas: new Map(), hojasOcultas: new Set(), fuentes: new Map(), filas: new Map(), unicos: new Map(), porDefecto: null };
    if (zip.file('styles.xml')) leerEstilos(await zip.file('styles.xml').async('string'), estilos);
    const texto = await contenido.async('string');
    // Los estilos automáticos van antes del cuerpo: se leen en la misma pasada que las hojas.
    const hojas = [];
    const memo = new Map(), memoEstilo = new Map();
    let hoja = null, fila = 0, col = 0, repFila = 1, filaActual = null, celda = null;
    let enCelda = false, parrafos = [], parrafoActual = null, profundidadTabla = 0, enNota = 0, celdas = 0, vacias = 0, ranuras = 0;

    // El estilo y el formato de número por defecto de una columna: la hoja guarda tramos de columnas seguidas con lo mismo ({ ini, tiene, est, cod }),
    // así una fila de 16.000 celdas vacías repetidas se mira por tramos y no por columnas.
    // (hoja.tramos lleva también el formato de número; hoja.tramosEst solo el estilo, y junta las columnas sin estilo, que son casi todas)
    const tramoDe = (t, c) => {
      let a = 0, b = t.length - 1, r = -1;
      while (a <= b) { const m = (a + b) >> 1; if (t[m].ini <= c) { r = m; a = m + 1; } else b = m - 1; }
      return r;
    };
    // Los estilos automáticos de content.xml se leen aparte, con el lector de estilos, antes de recorrer el cuerpo.
    const a0 = texto.indexOf('<office:automatic-styles'), a1 = texto.indexOf('</office:automatic-styles>');
    if (a0 >= 0 && a1 > a0) leerEstilos(texto.slice(a0, a1 + 26), estilos);
    const cuerpo = a0 >= 0 && a1 > a0 ? texto.slice(a1 + 26) : texto;

    // Recorre las columnas [x, fin) por tramos de la lista y llama a f(desde, hasta, tramo|null) para cada trozo (null: columnas sin definir).
    const trozos = (lista, x, fin, f) => {
      const nCol = hoja.anchos.length;
      while (x < fin) {
        const i = x < nCol ? tramoDe(lista, x) : -1;
        let hasta = fin, t = null;
        if (i >= 0) { t = lista[i]; hasta = Math.min(fin, i + 1 < lista.length ? lista[i + 1].ini : nCol); }
        f(x, hasta, t);
        x = hasta;
      }
    };

    // n celdas vacías con estilo desde desde hasta hasta (sin incluir), dentro del tope de todo el libro; dice si puso alguna
    const ponerVacias = (desde, hasta, est) => {
      const n = Math.min(hasta - desde, MAX_VACIAS - vacias);
      if (n <= 0) return false;
      const v = { valor: null, estilo: est };
      for (let k = 0; k < n; k++) filaActual.celdas[desde + k] = v;
      vacias += n; filaActual.vacias += n;
      return true;
    };

    const cierraCelda = () => {
      const c = celda; celda = null; enCelda = false;
      const rep = c.rep;
      const tipo = c.tipo;
      let base, def;   // el valor sin estilo y el formato de número que lleva si el archivo no dice otro
      const cadena = () => { let t = c.a['string-value'] !== undefined ? c.a['string-value'] : parrafos.join('\n'); if (t.length > MAX_TEXTO) t = t.slice(0, MAX_TEXTO); return t; };
      switch (tipo) {
        case 'float': case 'currency': { const n = parseFloat(c.a.value); if (Number.isFinite(n)) base = n; break; }
        case 'percentage': { const n = parseFloat(c.a.value); if (Number.isFinite(n)) { base = n; def = '0%'; } break; }
        case 'date': { const d = serieDeFecha(c.a['date-value'] || ''); if (d) { base = d.serie; def = d.conHora ? 'yyyy-mm-dd hh:mm:ss' : 'yyyy-mm-dd'; } break; }
        case 'time': { const h = serieDeHora(c.a['time-value'] || ''); if (h !== null) { base = h; def = 'hh:mm:ss'; } break; }
        case 'boolean': base = c.a['boolean-value'] === 'true'; break;
        case 'error': { const e = c.a['string-value'] || parrafos.join(''); base = { error: ERRORES.has(e) ? e : ERRORES_ES[e] || '#VALUE!' }; break; }
        case 'string': { const t = cadena(); if (t !== '' || (conFormulas && c.a.formula)) base = t; break; }   // una fórmula que da "" sigue siendo fórmula
        default: break;
      }
      parrafos = [];
      // el valor de la celda con el estilo y el formato de `par` ({ est, cod }): crearLibro reparte los estilos iguales por el objeto, que es el mismo en todas
      const hacer = (par) => {
        if (base === undefined) return par.est ? { valor: null, estilo: par.est } : undefined;
        const f = typeof base === 'number' ? (par.cod && par.cod !== 'General' ? par.cod : def) : undefined;
        if (!f && !par.est) return base;
        const o = { valor: base };
        if (f) o.formato = f;
        if (par.est) o.estilo = par.est;
        return o;
      };
      const propio = c.estilo ? { est: estiloDeCelda(estilos, c.estilo, memoEstilo), cod: formatoDeCelda(estilos, c.estilo, memo) } : null;
      const fin = Math.min(col + rep, MAX_COLS);
      let hayAlgo = false;
      // las celdas cubiertas por una combinada se ven como la primera: su estilo se apunta para cuando aparezcan (en la misma fila o en las de abajo)
      if ((c.fs > 1 || c.cs > 1) && hoja.fusiones.length < MAX_FUSIONES && col < MAX_COLS) {
        let par = propio;
        if (!par) trozos(hoja.tramos, col, col + 1, (x, y, t) => { par = t && t.tiene ? t : filaActual.defecto; });
        if (par && par.est) hoja.fusiones.push({ c0: col, c1: Math.min(col + c.cs - 1, MAX_COLS - 1), f0: fila, f1: fila + c.fs - 1, est: par.est });
      }
      if (base === undefined) {
        // vacía: solo se escribe si se ve (relleno o bordes; una fuente o una alineación no dibujan nada), y con un tope en todo el libro
        const poner = (desde, hasta, par) => { if (ponerVacias(desde, hasta, par.est)) hayAlgo = true; };
        if (propio) { if (seVe(propio.est)) poner(col, fin, propio); }
        else if (vacias < MAX_VACIAS) {
          trozos(hoja.tramosEst, col, fin, (x, y, t) => { const par = t && t.est ? t : filaActual.defecto; if (seVe(par.est)) poner(x, y, par); });
        }
        if (hayAlgo && (c.fs > 1 || c.cs > 1)) filaActual.combinadas.push({ c: col, f: fila, nc: c.cs, nf: c.fs });
        col += rep;
        return;
      }
      if (propio) { const v = hacer(propio); for (let x = col; x < fin; x++) filaActual.celdas[x] = v; }
      else trozos(hoja.tramos, col, fin, (x, y, t) => { const v = hacer(t && t.tiene ? t : filaActual.defecto); for (let j = x; j < y; j++) filaActual.celdas[j] = v; });
      // la fórmula (OpenFormula → texto de Excel) va en la primera celda; el valor guardado queda como último resultado
      if (conFormulas && c.a.formula && col < MAX_COLS && c.a.formula.length <= MAX_FORMULA && c.a.formula.startsWith('of:=') && !/^of:=#[A-Za-z\/0!?]+$/.test(c.a.formula)) {   // un error suelto (of:=#N/A) es como LibreOffice guarda una celda de error, no una fórmula
        const texto = OdsFormulas.aExcel(c.a.formula);
        if (texto && texto.length < 8000) {
          const v = filaActual.celdas[col];
          const o = v && typeof v === 'object' && !('error' in v) ? { ...v } : { valor: v === undefined ? null : v };
          o.formula = texto;
          const mc = entero(c.a['number-matrix-columns-spanned'], 0), mf = entero(c.a['number-matrix-rows-spanned'], 0);
          if (mc && mf) o.matriz = LibroXlsx.numeroACol(col) + (fila + 1) + (mc > 1 || mf > 1 ? ':' + LibroXlsx.numeroACol(Math.min(col + mc - 1, MAX_COLS - 1)) + Math.min(fila + mf, MAX_FILAS) : '');
          filaActual.celdas[col] = o;
        }
      }
      if (fin > col) { filaActual.llenas += fin - col; if (fin > filaActual.largo) filaActual.largo = fin; }
      if (c.fs > 1 || c.cs > 1) filaActual.combinadas.push({ c: col, f: fila, nc: c.cs, nf: c.fs });
      col += rep;
    };
    const cierraFila = () => {
      const f = filaActual; filaActual = null;
      if (f.celdas.length && fila < MAX_FILAS) {
        let n = Math.max(1, Math.min(repFila, MAX_FILAS - fila));
        if (f.llenas) {
          n = Math.max(1, Math.min(n, Math.floor(5e6 / f.largo)));
          if ((celdas += n * f.largo) > MAX_CELDAS) throw new LibroXlsx.ErrorXlsx('grande', 'más de ' + MAX_CELDAS + ' celdas');
          if (f.vacias && n > 1) {
            // las celdas vacías con estilo de una fila repetida cuentan en su tope; si no caben, la fila pierde esas celdas
            // (y la fila entera, que con una vacía al final mide 16.384, se copia n veces: sin el tope, 1 M de copias agotaba la memoria)
            if (vacias + (n - 1) * f.vacias <= MAX_VACIAS && n * f.celdas.length <= 5e6) vacias += (n - 1) * f.vacias;
            else {
              for (let k = 0; k < f.celdas.length; k++) if (f.celdas[k] && f.celdas[k].valor === null) f.celdas[k] = undefined;
              while (f.celdas.length && f.celdas[f.celdas.length - 1] === undefined) f.celdas.length--;
              vacias -= f.vacias;
            }
          }
        } else {
          n = Math.max(1, Math.min(n, 1 + Math.floor((MAX_VACIAS - vacias) / f.vacias), Math.floor(5e6 / f.celdas.length)));   // una sola vacía en XFD repetida 1 M de veces: 1 M de filas de 16.384
          vacias += (n - 1) * f.vacias;
        }
        // el tope de lugares después de la última celda con contenido (los arman las vacías con estilo): si no cabe, la fila pierde esas vacías
        // (y si no le queda nada, no se arma); hasta f.largo ya lo topa MAX_CELDAS
        if (f.celdas.length > f.largo && ranuras + n * (f.celdas.length - f.largo) > MAX_RANURAS) f.celdas.length = f.largo;
        ranuras += n * (f.celdas.length - f.largo);
        // una fila repetida con contenido se copia (el contenido es el mismo)
        if (f.celdas.length) for (let k = 0; k < n; k++) hoja.filas[fila + k] = k === 0 ? f.celdas : f.celdas.slice();
      }
      if (f.oculta) for (let k = 0; k < Math.min(repFila, 100000) && fila + k < MAX_FILAS; k++) hoja.filasOcultas.push(fila + k);
      if (f.alto) for (let k = 0; k < repFila && fila + k < MAX_FILAS && hoja.nAltos < MAX_ALTOS; k++, hoja.nAltos++) hoja.altosFila[fila + k] = f.alto;
      for (const m of f.combinadas) if (hoja.combinadasBrutas.length < MAX_COMBINADAS) hoja.combinadasBrutas.push(m);
      fila += repFila;
    };

    XmlLigero.recorrer(cuerpo, {
      abrir(nombre, a, vacio) {
        void vacio;
        if (nombre === 'table' ) {
          if (profundidadTabla++ > 0) return;   // una tabla dentro de otra (objeto incrustado): no es una hoja
          if (hojas.length >= MAX_HOJAS) throw new LibroXlsx.ErrorXlsx('grande', 'más de ' + MAX_HOJAS + ' hojas');
          hoja = { nombre: a.name || ('Hoja' + (hojas.length + 1)), filas: [], anchos: [], combinadasBrutas: [], tramos: [], tramosEst: [], fusiones: [], altosFila: {}, nAltos: 0, columnasOcultas: [], filasOcultas: [], oculta: estilos.hojasOcultas.has(a['style-name']) };
          hojas.push(hoja); fila = 0; col = 0; return;
        }
        if (!hoja || profundidadTabla !== 1) return;
        if (nombre === 'annotation') { enNota++; parrafoActual = null; return; }   // el comentario de una celda no es su texto
        if (enNota) return;
        switch (nombre) {
          case 'table-column': {
            const rep = Math.min(entero(a['number-columns-repeated'], 1), MAX_COLS);
            const w = a['style-name'] ? estilos.columnas.get(a['style-name']) || 0 : 0;
            const oculta = a.visibility === 'collapse' || a.visibility === 'filter';
            const nom = a['default-cell-style-name'] || null;
            const ini = hoja.anchos.length;
            if (ini < MAX_COLS) {
              const est = nom ? estiloDeCelda(estilos, nom, memoEstilo) : null, cod = nom ? formatoDeCelda(estilos, nom, memo) : null;
              const u = hoja.tramos[hoja.tramos.length - 1], ue = hoja.tramosEst[hoja.tramosEst.length - 1];
              if (!u || u.tiene !== !!nom || u.est !== est || u.cod !== cod) hoja.tramos.push({ ini, tiene: !!nom, est, cod });
              if (!ue || ue.est !== est) hoja.tramosEst.push({ ini, est });
            }
            for (let k = 0; k < rep && hoja.anchos.length < MAX_COLS; k++) {
              if (oculta) hoja.columnasOcultas.push(hoja.anchos.length);
              hoja.anchos.push(w);
            }
            break;
          }
          case 'table-row': {
            repFila = Math.min(entero(a['number-rows-repeated'], 1), MAX_FILAS);
            const ef = a['style-name'] ? estilos.filas.get(a['style-name']) : null;
            const nomDef = a['default-cell-style-name'];
            filaActual = {
              celdas: [], combinadas: [], oculta: a.visibility === 'collapse' || a.visibility === 'filter',
              defecto: nomDef ? { est: estiloDeCelda(estilos, nomDef, memoEstilo), cod: formatoDeCelda(estilos, nomDef, memo) } : SIN_ESTILO,   // el de las columnas que no traen uno
              alto: ef && !ef.optimo ? ef.alto : 0, largo: 0, llenas: 0, vacias: 0,
            };
            col = 0;
            if (hoja.fusiones.length) hoja.fusiones = hoja.fusiones.filter((m) => m.f1 >= fila);
            break;
          }
          case 'table-cell':
            if (!filaActual) return;
            celda = {
              tipo: a['value-type'] || null, a, estilo: a['style-name'] || null,
              rep: Math.min(entero(a['number-columns-repeated'], 1), MAX_COLS),
              cs: entero(a['number-columns-spanned'], 1), fs: entero(a['number-rows-spanned'], 1),
            };
            enCelda = true; parrafoActual = null;
            break;
          case 'covered-table-cell': {
            if (!filaActual) return;
            // una celda cubierta con valor es parte del resultado de una fórmula de matriz (las cubiertas por una combinada no traen valor): se lee como cualquier celda
            if (a['value-type'] && !hoja.fusiones.some((m) => m.f0 <= fila && fila <= m.f1 && m.c0 <= col && col <= m.c1)) {
              celda = { tipo: a['value-type'], a, estilo: a['style-name'] || null, rep: Math.min(entero(a['number-columns-repeated'], 1), MAX_COLS), cs: 1, fs: 1, cubierta: true };
              enCelda = true; parrafoActual = null;
              break;
            }
            const rep = Math.min(entero(a['number-columns-repeated'], 1), MAX_COLS);
            for (const m of hoja.fusiones) if (m.f0 <= fila && fila <= m.f1) ponerVacias(Math.max(col, m.c0), Math.min(col + rep, m.c1 + 1), m.est);
            col += rep;
            break;
          }
          case 'p':
            if (enCelda) parrafoActual = '';
            break;
          case 's':
            if (parrafoActual !== null) parrafoActual += ' '.repeat(Math.min(entero(a.c, 1), 10000));
            break;
          case 'tab': if (parrafoActual !== null) parrafoActual += '\t'; break;
          case 'line-break': if (parrafoActual !== null) parrafoActual += '\n'; break;
          default: break;
        }
      },
      cerrar(nombre) {
        if (nombre === 'table') { if (--profundidadTabla === 0) hoja = null; return; }
        if (!hoja || profundidadTabla !== 1) return;
        if (nombre === 'annotation') { enNota--; return; }
        if (enNota) return;
        if (nombre === 'p' && enCelda && parrafoActual !== null) { parrafos.push(parrafoActual); parrafoActual = null; }
        else if ((nombre === 'table-cell' || (nombre === 'covered-table-cell' && celda)) && celda) cierraCelda();
        else if (nombre === 'table-row' && filaActual) cierraFila();
      },
      texto(t) { if (parrafoActual !== null) parrafoActual += t; },
    });

    if (!hojas.length) throw new LibroXlsx.ErrorXlsx('sinLibro', 'sin hojas');
    return {
      hojas: hojas.map((h) => {
        const filas = [];
        let ancho = 0;
        for (let r = 0; r < h.filas.length; r++) { const fila = h.filas[r] || []; filas.push(fila); if (fila.length > ancho) ancho = fila.length; }
        const rango = (m) => {
          const col = LibroXlsx_numeroACol(m.c + 1), col2 = LibroXlsx_numeroACol(Math.min(m.c + m.nc, MAX_COLS));
          return col + (m.f + 1) + ':' + col2 + (Math.min(m.f + m.nf, MAX_FILAS));
        };
        return { nombre: h.nombre, filas, anchos: h.anchos.slice(0, ancho), combinadas: h.combinadasBrutas.map(rango), estado: h.oculta ? 'oculta' : 'visible', columnasOcultas: h.columnasOcultas.filter((c) => c < ancho), filasOcultas: h.filasOcultas.filter((r) => r < filas.length), altosFila: h.altosFila };
      }),
    };
  }

  function LibroXlsx_numeroACol(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }

  globalThis.OdsLector = { leer, _traducirEstilo: traducirEstilo, _serieDeFecha: serieDeFecha };
})();

/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Estilos de celda de una hoja OpenDocument (.ods) para el libro editable de la suite (F3.4a, etapa 2).
// ORIGEN: adaptado de `ods-lector.js` del sitio (copia del 2026-10-28): `leerEstilos`, `traducirEstilo` (formato de número de ODF → código de Excel),
// `capturarPropiedades`, `bordeDe`. Diferencias, para devolverlo al madre (decisión D8) en `docs/motores-adaptados.md`:
//   - los anchos se guardan en píxeles (el modelo de Excel los lleva en caracteres: ancho = px / 7) y los altos en puntos, con su marca de «alto óptimo»;
//   - el estilo de una celda se devuelve como propiedades ya mezcladas con su cadena de padres (`propiedades`), no como el objeto de crearLibro;
//   - `TablaXf` arma un styles.xml VIRTUAL con un xf por cada combinación distinta (propiedades + formato de número) y lo lee `XlsxEstilos.leer`:
//     así `estilos.xf(i)` da lo mismo que en un .xlsx y la rejilla, el PDF y la cinta no saben que el archivo es ODF.
// Requiere xml-ligero.js. Sin DOM.
//
//   const estilos = OdsEstilos.crear();  OdsEstilos.leer(textoXml, estilos);          // styles.xml y luego las automáticas de content.xml
//   OdsEstilos.codigoFormato(estilos, nombreEstiloCelda) → código de Excel | null      (el formato de número del estilo o de su cadena de padres)
//   OdsEstilos.propiedades(estilos, nombreEstiloCelda) → { fn, ff, fs, fw, … }          (claves cortas de capturarPropiedades; null/'' = el estilo por defecto)
//   const tabla = OdsEstilos.tablaXf(estilos);  tabla.indice(nombreEstilo, formatoPorDefecto) → número de xf;  tabla.xml() → styles.xml virtual
(function () {
  const MAX_ESTILOS = 100000;     // estilos de celda, de fila y de columna que se guardan (un archivo hostil puede declarar millones)
  const MAX_DIGITOS = 30;         // cifras de un formato de número: Excel no muestra más de 30 decimales
  const MAX_XFS = 60000;          // xf distintos de un libro (el .xlsx admite 64.000)

  const entero = (s, defecto) => { const n = parseInt(s, 10); return n > 0 ? n : defecto; };
  const cifras = (s, defecto) => Math.min(entero(s, defecto), MAX_DIGITOS);
  const comillas = (t) => '"' + t.replace(/"/g, '""') + '"';
  const escXml = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  // Una longitud de ODF ("0.74pt", "1.059cm", "15.75pt") a puntos; null si no es una medida.
  function puntosDe(texto) {
    const m = /^(-?[\d.]+)\s*(cm|mm|in|pt|px|pc)?$/.exec(texto || '');
    if (!m) return null;
    const n = parseFloat(m[1]), u = m[2] || 'pt';
    return u === 'cm' ? n * 72 / 2.54 : u === 'mm' ? n * 72 / 25.4 : u === 'in' ? n * 72 : u === 'px' ? n * 0.75 : u === 'pc' ? n * 12 : n;
  }
  const pxDe = (texto) => { const pt = puntosDe(texto); return pt === null ? 0 : pt * 96 / 72; };

  // ---- formatos de número de ODF → código de Excel ---------------------------------------------------------------------------
  // estilo = { tipo, partes: [{ k, a }], mapas: [{ cond, estilo }] }; se traduce al final, cuando ya están todos.
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
        case 'quarter': s += '"Q"0'; break;
        case 'hours': s += e.a['truncate-on-overflow'] === 'false' ? '[h]' : (a.style === 'long' ? 'hh' : 'h'); break;
        case 'minutes': s += e.a['truncate-on-overflow'] === 'false' && !e.partes.some((q) => q.k === 'hours') ? '[mm]' : (a.style === 'long' ? 'mm' : 'm'); break;
        case 'seconds': s += (e.a['truncate-on-overflow'] === 'false' && !e.partes.some((q) => q.k === 'hours' || q.k === 'minutes') ? '[ss]' : (a.style === 'long' ? 'ss' : 's')) + (cifras(a['decimal-places'], 0) > 0 && a['decimal-places'] !== undefined ? '.' + '0'.repeat(cifras(a['decimal-places'], 0)) : ''); break;
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
      if (c && otro) secciones.push((c[1] === '>=' && c[2] === '0' ? '' : '[' + (c[1] === '!=' ? '<>' : c[1]) + c[2] + ']') + otro);
    }
    return secciones.length ? secciones.join(';') + ';' + s : (s || null);
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

  // "0.74pt solid #000000" → { e, c } con e = nombre del borde de Excel (thin, medium, double, dashed…); null si no hay borde.
  function bordeDe(texto) {
    if (!texto || texto === 'none' || texto === 'hidden') return null;
    let ancho = null, forma = 'solid', color;
    for (const t of texto.trim().split(/\s+/).slice(0, 4)) {
      if (t.charAt(0) === '#') { if (/^#[0-9a-fA-F]{6}$/.test(t)) color = t.slice(1).toUpperCase(); }
      else if (t === 'thin') ancho = 0.75;
      else if (t === 'medium') ancho = 1.75;
      else if (t === 'thick') ancho = 2.5;
      else if (/^(solid|dotted|dashed|double|dot-dash|dot-dot-dash|none|hidden)$/.test(t)) forma = t;
      else if (/^double-thin$/.test(t)) forma = 'double';
      else if (/^(fine-dashed|long-dash|dash-dot|dash-dot-dot|dash|ultrafine-dashed)$/.test(t)) forma = /dot-dot/.test(t) ? 'dot-dot-dash' : /dot/.test(t) ? 'dot-dash' : 'dashed';
      else { const pt = puntosDe(t); if (pt !== null) ancho = pt; }
    }
    if (forma === 'none' || forma === 'hidden') return null;
    if (ancho === null) ancho = 0.75;
    const grueso = ancho >= 1.2;
    const grosor = ancho <= 0.4 ? 'hair' : ancho < 1.2 ? 'thin' : ancho < 2.5 ? 'medium' : 'thick';   // 2.49 pt (el «grueso» de LibreOffice) lo lee Excel como medio
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

  // ---- lectura de los estilos ------------------------------------------------------------------------------------------------
  function crear() {
    // datos: nombre → { tipo, partes, mapas, a };  celdas: nombre → { dato, padre, p };  columnas: nombre → px;  filas: nombre → { alto (pt), optimo };  hojasOcultas: estilos de tabla con display=false;  paginaDeTabla: estilo de tabla → página maestra
    return { alias: new Map(), datos: new Map(), celdas: new Map(), columnas: new Map(), filas: new Map(), hojasOcultas: new Set(), paginaDeTabla: new Map(), fuentes: new Map(), porDefecto: null, _celda: null, _nombre: null };
  }
  function leer(texto, estilos) {
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
          if (a.family === 'table-cell' && a['display-name'] && a['display-name'] !== a.name) estilos.alias.set(a['display-name'], a.name);   // los formatos condicionales nombran el estilo por su nombre visible
          if (a.family === 'table-cell' && estilos.celdas.size < MAX_ESTILOS) {
            estilos._celda = { dato: a['data-style-name'] || null, padre: a['parent-style-name'] || null, p: {} };
            estilos.celdas.set(a.name, estilos._celda);
          }
          estilos._nombre = a.name;
          if (a.family === 'table' && a['master-page-name'] && estilos.paginaDeTabla.size < MAX_ESTILOS) estilos.paginaDeTabla.set(a.name, a['master-page-name']);   // la página maestra (papel, márgenes, encabezado) de la hoja
        } else if (nombre === 'table-column-properties' && a['column-width'] !== undefined) { if (estilos.columnas.size < MAX_ESTILOS) estilos.columnas.set(estilos._nombre, pxDe(a['column-width'])); }
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

  // El formato de número (código de Excel) del estilo de celda o del primero de sus padres que tenga uno.
  function codigoFormato(estilos, nombre) {
    if (!nombre) return null;
    let dato = null, n = nombre, saltos = 0;
    while (n && saltos++ < 20) {
      const e = estilos.celdas.get(n);
      if (!e) break;
      if (e.dato) { dato = e.dato; break; }
      n = e.padre;
    }
    return dato ? traducirEstilo(estilos, dato) : null;
  }
  // Las propiedades del estilo de celda mezcladas con las de toda su cadena de padres, de la raíz al hijo, sobre el estilo por defecto de las celdas.
  // La cadena se corta a 30 eslabones y no repite nombres (un ciclo en un archivo hostil no la cuelga).
  function propiedades(estilos, nombre) {
    const r = {};
    if (estilos.porDefecto) Object.assign(r, estilos.porDefecto.p);
    if (!nombre && estilos.celdas.has('Default')) nombre = 'Default';   // en LibreOffice «Default» es el estilo de toda celda sin estilo
    if (!nombre) return r;
    const cadena = [], vistos = new Set();
    for (let n = nombre; n && !vistos.has(n) && cadena.length < 30;) {
      vistos.add(n);
      const e = estilos.celdas.get(n);
      if (!e) break;
      cadena.push(e);
      n = e.padre;
    }
    for (let k = cadena.length - 1; k >= 0; k--) Object.assign(r, cadena[k].p);
    return r;
  }

  // Solo lo que el estilo (y sus padres, salvo «Default») define: es lo que un formato condicional pone encima de la celda.
  function propiedadesPropias(estilos, nombre) {
    const r = {};
    const cadena = [], vistos = new Set();
    for (let n = nombre; n && n !== 'Default' && !vistos.has(n) && cadena.length < 30;) {
      vistos.add(n);
      const e = estilos.celdas.get(n);
      if (!e) break;
      cadena.push(e);
      n = e.padre;
    }
    for (let k = cadena.length - 1; k >= 0; k--) Object.assign(r, cadena[k].p);
    return r;
  }

  // ---- styles.xml virtual -------------------------------------------------------------------------------------------------------
  const hex6 = (c) => (typeof c === 'string' && /^#[0-9a-fA-F]{6}$/.test(c) ? c.slice(1).toUpperCase() : null);
  function tablaXf(estilos) {
    const fuentes = [], rellenos = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'], bordes = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    const codigos = new Map();   // código de formato → numFmtId (164 en adelante)
    const xfs = [], clavesXf = new Map(), unico = (lista, mapa, xml) => { let i = mapa.get(xml); if (i === undefined) { i = lista.length; lista.push(xml); mapa.set(xml, i); } return i; };
    const mapaF = new Map(), mapaR = new Map([[rellenos[0], 0], [rellenos[1], 1]]), mapaB = new Map([[bordes[0], 0]]);
    const porDefecto = estilos.porDefecto ? estilos.porDefecto.p : {};

    function xmlFuente(r) {
      const nombreCrudo = r.fn !== undefined ? (estilos.fuentes.get(r.fn) || r.fn) : r.ff !== undefined ? r.ff : null;
      const nombre = (nombreCrudo && familiaDe(nombreCrudo)) || 'Liberation Sans';
      const tam = r.fs !== undefined ? puntosDe(r.fs) : null;
      let x = '<font>';
      if (r.fw === 'bold' || parseInt(r.fw, 10) >= 600) x += '<b/>';
      if (r.fi === 'italic' || r.fi === 'oblique') x += '<i/>';
      if (r.lt !== undefined && r.lt !== 'none') x += '<strike/>';
      if (r.us !== undefined && r.us !== 'none') x += r.ut === 'double' ? '<u val="double"/>' : '<u/>';
      x += '<sz val="' + (tam && tam > 0 ? Math.round(tam * 100) / 100 : 10) + '"/>';
      const co = hex6(r.co);
      if (co) x += '<color rgb="FF' + co + '"/>';
      return x + '<name val="' + escXml(nombre) + '"/></font>';
    }
    function xmlBorde(r) {
      let hay = false;
      const lado = (etiqueta, clave) => {
        const b = bordeDe(r[clave]);
        if (!b) return '<' + etiqueta + '/>';
        hay = true;
        return '<' + etiqueta + ' style="' + b.e + '">' + (b.c ? '<color rgb="FF' + b.c + '"/>' : '<color auto="1"/>') + '</' + etiqueta + '>';
      };
      const x = '<border>' + lado('left', 'bi') + lado('right', 'bd') + lado('top', 'ba') + lado('bottom', 'bb') + '<diagonal/></border>';
      return hay ? x : bordes[0];
    }
    function xmlAlineacion(r) {
      const a = [];
      let h = r.tas === 'value-type' ? null : r.ta === 'start' || r.ta === 'left' ? 'left' : r.ta === 'end' || r.ta === 'right' ? 'right' : r.ta === 'center' ? 'center' : r.ta === 'justify' ? 'justify' : null;
      if (r.rc === 'true') h = 'fill';
      if (h) a.push('horizontal="' + h + '"');
      if (r.va === 'top') a.push('vertical="top"'); else if (r.va === 'middle') a.push('vertical="center"');
      if (r.aj === 'wrap') a.push('wrapText="1"');
      const sangria = puntosDe(h === 'right' ? r.mr : r.ml);
      if (sangria && sangria >= 5) a.push('indent="' + Math.min(250, Math.round(sangria / 10)) + '"');
      const ang = r.ro === undefined ? 0 : ((Math.round(parseFloat(r.ro)) % 360) + 360) % 360;
      if (ang >= 1 && ang <= 90) a.push('textRotation="' + ang + '"');
      else if (ang >= 270) a.push('textRotation="' + (90 + (360 - ang)) + '"');
      else if (r.di === 'ttb') a.push('textRotation="255"');
      return a.length ? '<alignment ' + a.join(' ') + '/>' : '';
    }

    // Cada xf recuerda de dónde salió: { r (propiedades mezcladas), codigo, nombre (estilo con nombre del archivo, o null), nuevo (lo creó derivar), auto (nombre del estilo
    // automático que le toca al guardar), escrito }. Al guardar solo los xf `nuevo` hacen falta como estilos automáticos; los demás ya están en el archivo.
    const defs = [];
    function definir(r, codigo, nombre, nuevo) {
      const clave = JSON.stringify([r, codigo]);
      let i = clavesXf.get(clave);
      if (i !== undefined) return i;
      if (xfs.length >= MAX_XFS) return 0;
      let idFmt = 0;
      if (codigo !== 'General') { idFmt = codigos.get(codigo); if (idFmt === undefined) { idFmt = 164 + codigos.size; codigos.set(codigo, idFmt); } }
      const fuente = unico(fuentes, mapaF, xmlFuente(r));
      const bg = hex6(r.bg);
      const relleno = bg ? unico(rellenos, mapaR, '<fill><patternFill patternType="solid"><fgColor rgb="FF' + bg + '"/><bgColor indexed="64"/></patternFill></fill>') : 0;
      const borde = unico(bordes, mapaB, xmlBorde(r));
      const al = xmlAlineacion(r);
      i = xfs.length;
      xfs.push('<xf numFmtId="' + idFmt + '" fontId="' + fuente + '" fillId="' + relleno + '" borderId="' + borde + '" xfId="0"' + (idFmt ? ' applyNumberFormat="1"' : '') + (al ? ' applyAlignment="1">' + al + '</xf>' : '/>'));
      defs.push({ r, codigo, nombre: nombre && nombre !== 'Default' ? nombre : null, nuevo: !!nuevo, auto: null, escrito: false });
      clavesXf.set(clave, i);
      return i;
    }
    // Número de xf para un estilo de celda (null/'' = el de las celdas sin estilo) y el formato que le toca si el estilo no trae uno (porcentaje, fecha, hora).
    function indice(nombre, formatoPorDefecto) {
      const cod = codigoFormato(estilos, nombre);
      return definir(propiedades(estilos, nombre), cod && cod !== 'General' ? cod : (formatoPorDefecto || 'General'), nombre, false);
    }
    indice(null, null);   // el xf 0 es el de las celdas sin estilo
    void porDefecto;

    // Un xf nuevo a partir de otro, con los cambios de la cinta (los mismos nombres que EstilosLibro.derivar). Se escribe como estilo automático al guardar.
    const BORDE_ODF = { thin: '0.74pt solid #000000', medium: '1.76pt solid #000000', thick: '2.49pt solid #000000' };
    function derivar(base, c) {
      const d = defs[base] || defs[0];
      const r = Object.assign({}, d.r);
      let codigo = d.codigo;
      const poner = (k, v) => { if (v === null || v === undefined) delete r[k]; else r[k] = v; };
      if ('negrita' in c) poner('fw', c.negrita ? 'bold' : 'normal');
      if ('cursiva' in c) poner('fi', c.cursiva ? 'italic' : 'normal');
      if ('subrayado' in c) { poner('us', c.subrayado ? 'solid' : 'none'); poner('ut', c.subrayado === 'double' ? 'double' : c.subrayado ? 'single' : null); }
      if ('tachado' in c) poner('lt', c.tachado ? 'solid' : 'none');
      if ('colorFuente' in c) poner('co', c.colorFuente ? '#' + String(c.colorFuente).toLowerCase() : null);
      if ('tamano' in c && c.tamano) poner('fs', +c.tamano + 'pt');
      if ('fuente' in c && c.fuente) { poner('ff', c.fuente); delete r.fn; }
      if ('relleno' in c) poner('bg', c.relleno ? '#' + String(c.relleno).toLowerCase() : 'transparent');
      if ('h' in c) { poner('ta', { left: 'start', center: 'center', right: 'end', justify: 'justify' }[c.h] || null); poner('rc', c.h === 'fill' ? 'true' : null); }
      if ('v' in c) poner('va', { top: 'top', center: 'middle', bottom: 'bottom' }[c.v] || null);
      if ('ajustar' in c) poner('aj', c.ajustar ? 'wrap' : 'no-wrap');
      if ('formato' in c) codigo = c.formato || 'General';
      if (c.bordes === 'todos') for (const k of ['bi', 'bd', 'ba', 'bb']) r[k] = BORDE_ODF.thin;
      if (c.bordes === 'ninguno') for (const k of ['bi', 'bd', 'ba', 'bb']) r[k] = 'none';
      if (c.borde) { const m = { izq: 'bi', der: 'bd', arr: 'ba', aba: 'bb' }; for (const k in c.borde) if (m[k]) r[m[k]] = c.borde[k] ? (BORDE_ODF[c.borde[k]] || BORDE_ODF.thin) : 'none'; }
      const antes = xfs.length, i = definir(r, codigo, null, true);
      if (i >= antes) defs[i].auto = null;   // el nombre del estilo automático lo pone quien guarda (conoce los que ya existen en el archivo)
      return i;
    }
    const def = (i) => defs[i] || defs[0];

    // formatos diferenciales (<dxf>) de los estilos que usan los formatos condicionales: índice de dxfs en el styles.xml virtual
    const dxfs = [], clavesDxf = new Map();
    function dxf(nombre) {
      nombre = estilos.alias.get(nombre) || nombre;
      if (clavesDxf.has(nombre)) return clavesDxf.get(nombre);
      const r = propiedadesPropias(estilos, nombre);
      let fuente = '';
      if (r.fw === 'bold' || parseInt(r.fw, 10) >= 600) fuente += '<b/>';
      if (r.fi === 'italic' || r.fi === 'oblique') fuente += '<i/>';
      if (r.lt !== undefined && r.lt !== 'none') fuente += '<strike/>';
      if (r.us !== undefined && r.us !== 'none') fuente += r.ut === 'double' ? '<u val="double"/>' : '<u/>';
      if (r.fs !== undefined && puntosDe(r.fs)) fuente += '<sz val="' + puntosDe(r.fs) + '"/>';
      const co = hex6(r.co);
      if (co) fuente += '<color rgb="FF' + co + '"/>';
      const bg = hex6(r.bg);
      let borde = '';
      for (const [k, lado] of [['bi', 'left'], ['bd', 'right'], ['ba', 'top'], ['bb', 'bottom']]) {
        const b = bordeDe(r[k]);
        if (b) borde += '<' + lado + ' style="' + b.e + '">' + (b.c ? '<color rgb="FF' + b.c + '"/>' : '<color auto="1"/>') + '</' + lado + '>';
      }
      const cod = codigoFormato(estilos, nombre);
      const x = '<dxf>' + (fuente ? '<font>' + fuente + '</font>' : '') + (cod && cod !== 'General' ? '<numFmt numFmtId="' + (300 + dxfs.length) + '" formatCode="' + escXml(cod) + '"/>' : '')
        + (bg ? '<fill><patternFill patternType="solid"><fgColor rgb="FF' + bg + '"/><bgColor rgb="FF' + bg + '"/></patternFill></fill>' : '') + (borde ? '<border>' + borde + '</border>' : '') + '</dxf>';
      dxfs.push(x);
      clavesDxf.set(nombre, dxfs.length - 1);
      return dxfs.length - 1;
    }

    function xml() {
      const fmts = [...codigos].map(([c, id]) => '<numFmt numFmtId="' + id + '" formatCode="' + escXml(c) + '"/>').join('');
      return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + (fmts ? '<numFmts count="' + codigos.size + '">' + fmts + '</numFmts>' : '')
        + '<fonts count="' + fuentes.length + '">' + fuentes.join('') + '</fonts><fills count="' + rellenos.length + '">' + rellenos.join('') + '</fills><borders count="' + bordes.length + '">' + bordes.join('') + '</borders>'
        + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="' + xfs.length + '">' + xfs.join('') + '</cellXfs>'
        + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' + (dxfs.length ? '<dxfs count="' + dxfs.length + '">' + dxfs.join('') + '</dxfs>' : '') + '</styleSheet>';
    }
    return { indice, derivar, def, dxf, xml, cantidad: () => xfs.length };
  }

  globalThis.OdsEstilos = { crear, leer, codigoFormato, propiedades, propiedadesPropias, tablaXf, traducirEstilo, bordeDe, puntosDe, pxDe };
})();

/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Lector de hojas de cálculo OpenDocument (.ods, de LibreOffice, Google Sheets y Excel) del motor propio de Excel (etapa 9). Un .ods es un zip con
// content.xml (las hojas) y styles.xml (los estilos con nombre). Requiere zip-propio.js, xml-ligero.js y xlsx-modelo.js (para los errores); no depende del navegador.
//
//   const libro = await OdsLector.leer(bytes);
//   libro = { hojas: [{ nombre, filas: [[valor…]…], anchos: [caracteres…], combinadas: ['A1:B2'…], estado, columnasOcultas, filasOcultas }] }
//   valor: undefined (vacía) | número | texto | lógico | { error } | { valor, formato }   (lo que entiende XlsxEscritor.crearLibro)
//
// Qué lee: valores (número, porcentaje, moneda, fecha, hora, lógico, texto, error), el formato de número de cada celda (traducido a un código de Excel),
// combinadas, anchos de columna y el idioma-neutro de los textos con espacios y saltos. Las fórmulas no se leen: queda el último valor, que es lo que
// el archivo guarda (el sintaxis de fórmulas de ODF — "of:=[.A1]+1" — no es la de Excel). Formatos con color, relleno o bordes no se llevan.
(function () {
  const MAX_FILAS = 1048576, MAX_COLS = 16384, MAX_TEXTO = 32767;
  const MAX_CELDAS = 10000000;    // en todo el libro: crearLibro no arma un .xlsx de más de ~8 M celdas (una fila repetida un millón de veces cabe en 200 bytes)
  const MAX_COMBINADAS = 500000;  // por hoja, como LibroXlsx
  const ERRORES = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA']);

  const ent = (a, k) => (a[k] === undefined ? null : a[k]);
  const entero = (s, defecto) => { const n = parseInt(s, 10); return n > 0 ? n : defecto; };
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
          const dp = a['decimal-places'] === undefined ? null : entero(a['decimal-places'], 0);
          const minDec = a['min-decimal-places'] === undefined ? dp : entero(a['min-decimal-places'], 0);
          const minInt = Math.max(1, entero(a['min-integer-digits'], 1));
          if (dp === null && a.grouping !== 'true' && minInt === 1) { s += 'General'; break; }
          s += a.grouping === 'true' ? '#,' + '#'.repeat(Math.max(0, 3 - minInt)) + '0'.repeat(minInt) : '0'.repeat(minInt);
          if (dp) s += '.' + '0'.repeat(minDec || 0) + '#'.repeat(Math.max(0, dp - (minDec || 0)));
          break;
        }
        case 'scientific-number': {
          const dp = entero(a['decimal-places'], 0), minDec = a['min-decimal-places'] === undefined ? dp : entero(a['min-decimal-places'], 0);
          s += '0' + (dp ? '.' + '0'.repeat(minDec) + '#'.repeat(Math.max(0, dp - minDec)) : '') + 'E' + (a['forced-exponent-sign'] === 'true' ? '+' : '-') + '0'.repeat(entero(a['min-exponent-digits'], 1));
          break;
        }
        case 'fraction': {
          const fijo = a['denominator-value'], max = entero(a['max-denominator-value'], 0);
          const dig = fijo ? 0 : Math.max(entero(a['min-denominator-digits'], 1), String(max || 9).length);
          s += (entero(a['min-integer-digits'], 0) > 0 ? '0 ' : '# ') + '?'.repeat(Math.max(1, entero(a['min-numerator-digits'], 1))) + '/' + (fijo ? fijo : '?'.repeat(dig));
          break;
        }
        case 'day': s += a.style === 'long' ? 'dd' : 'd'; break;
        case 'day-of-week': s += a.style === 'long' ? 'dddd' : 'ddd'; break;
        case 'month': s += a.textual === 'true' ? (a.style === 'long' ? 'mmmm' : 'mmm') : (a.style === 'long' ? 'mm' : 'm'); break;
        case 'year': s += a.style === 'long' ? 'yyyy' : 'yy'; break;
        case 'quarter': s += a.style === 'long' ? '"Q"0' : '"Q"0'; break;
        case 'hours': s += e.a['truncate-on-overflow'] === 'false' ? '[h]' : (a.style === 'long' ? 'hh' : 'h'); break;
        case 'minutes': s += e.a['truncate-on-overflow'] === 'false' && !e.partes.some((q) => q.k === 'hours') ? '[mm]' : (a.style === 'long' ? 'mm' : 'm'); break;
        case 'seconds': s += (e.a['truncate-on-overflow'] === 'false' && !e.partes.some((q) => q.k === 'hours' || q.k === 'minutes') ? '[ss]' : (a.style === 'long' ? 'ss' : 's')) + (entero(a['decimal-places'], 0) ? '.' + '0'.repeat(entero(a['decimal-places'], 0)) : ''); break;
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
        if (nombre === 'style' && a.name !== undefined) {
          if (a.family === 'table-cell') { estilos.celdas.set(a.name, { dato: a['data-style-name'] || null, padre: a['parent-style-name'] || null }); }
          estilos._nombre = a.name;
        } else if (nombre === 'table-column-properties' && a['column-width'] !== undefined) estilos.columnas.set(estilos._nombre, caracteresDeAncho(a['column-width']));
        else if (nombre === 'table-properties' && a.display === 'false') estilos.hojasOcultas.add(estilos._nombre);
      },
      cerrar(nombre) {
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

  async function leer(bytes) {
    let zip;
    try { zip = await ZipPropio.loadAsync(bytes); } catch (e) { throw new LibroXlsx.ErrorXlsx('noZip', 'no es OpenDocument'); }
    const contenido = zip.file('content.xml');
    if (!contenido) throw new LibroXlsx.ErrorXlsx('noZip', 'no es OpenDocument');
    if (zip.file('META-INF/manifest.xml')) {
      const man = await zip.file('META-INF/manifest.xml').async('string');
      // con cualquier prefijo (o ninguno) del espacio de nombres del manifiesto: LibreOffice usa "manifest:", otros programas no siempre
      if (/<(?:[\w.-]+:)?encryption-data\b/.test(man)) throw new LibroXlsx.ErrorXlsx('cifrado');
    }
    const estilos = { datos: new Map(), celdas: new Map(), columnas: new Map(), hojasOcultas: new Set() };
    if (zip.file('styles.xml')) leerEstilos(await zip.file('styles.xml').async('string'), estilos);
    const texto = await contenido.async('string');
    // Los estilos automáticos van antes del cuerpo: se leen en la misma pasada que las hojas.
    const hojas = [];
    const memo = new Map();
    let hoja = null, fila = 0, col = 0, repFila = 1, filaActual = null, celda = null;
    let enCelda = false, parrafos = [], parrafoActual = null, profundidadTabla = 0, enNota = 0, celdas = 0;
    // Los estilos automáticos de content.xml se leen aparte, con el lector de estilos, antes de recorrer el cuerpo.
    const a0 = texto.indexOf('<office:automatic-styles'), a1 = texto.indexOf('</office:automatic-styles>');
    if (a0 >= 0 && a1 > a0) leerEstilos(texto.slice(a0, a1 + 26), estilos);
    const cuerpo = a0 >= 0 && a1 > a0 ? texto.slice(a1 + 26) : texto;

    const cierraCelda = () => {
      const c = celda; celda = null; enCelda = false;
      const rep = c.rep;
      const tipo = c.tipo;
      let valor;
      const codigo = formatoDeCelda(estilos, c.estilo || filaActual.estiloColumna(col), memo);
      const conFormato = (n, def) => { const f = codigo && codigo !== 'General' ? codigo : def; return f ? { valor: n, formato: f } : n; };
      const cadena = () => { let t = c.a['string-value'] !== undefined ? c.a['string-value'] : parrafos.join('\n'); if (t.length > MAX_TEXTO) t = t.slice(0, MAX_TEXTO); return t; };
      switch (tipo) {
        case 'float': { const n = parseFloat(c.a.value); valor = Number.isFinite(n) ? conFormato(n) : undefined; break; }
        case 'percentage': { const n = parseFloat(c.a.value); valor = Number.isFinite(n) ? conFormato(n, '0%') : undefined; break; }
        case 'currency': { const n = parseFloat(c.a.value); valor = Number.isFinite(n) ? conFormato(n) : undefined; break; }
        case 'date': { const d = serieDeFecha(c.a['date-value'] || ''); valor = d ? conFormato(d.serie, d.conHora ? 'yyyy-mm-dd hh:mm:ss' : 'yyyy-mm-dd') : undefined; break; }
        case 'time': { const h = serieDeHora(c.a['time-value'] || ''); valor = h === null ? undefined : conFormato(h, 'hh:mm:ss'); break; }
        case 'boolean': valor = c.a['boolean-value'] === 'true'; break;
        case 'error': { const e = c.a['string-value'] || parrafos.join(''); valor = { error: ERRORES.has(e) ? e : '#VALUE!' }; break; }
        case 'string': { const t = cadena(); valor = t === '' ? undefined : t; break; }
        default: valor = undefined;
      }
      parrafos = [];
      if (valor === undefined) { col += rep; return; }
      const tope = Math.min(rep, MAX_COLS - col);
      for (let k = 0; k < tope; k++) filaActual.celdas[col + k] = valor;
      if (c.fs > 1 || c.cs > 1) filaActual.combinadas.push({ c: col, f: fila, nc: c.cs, nf: c.fs });
      col += rep;
    };
    const cierraFila = () => {
      const f = filaActual; filaActual = null;
      if (f.celdas.length && fila < MAX_FILAS) {
        const n = Math.max(1, Math.min(repFila, MAX_FILAS - fila, Math.floor(5e6 / f.celdas.length)));
        if ((celdas += n * f.celdas.length) > MAX_CELDAS) throw new LibroXlsx.ErrorXlsx('grande', 'más de ' + MAX_CELDAS + ' celdas');
        // una fila repetida con contenido se copia (el contenido es el mismo)
        for (let k = 0; k < n; k++) hoja.filas[fila + k] = k === 0 ? f.celdas : f.celdas.slice();
      }
      if (f.oculta) for (let k = 0; k < Math.min(repFila, 100000) && fila + k < MAX_FILAS; k++) hoja.filasOcultas.push(fila + k);
      for (const m of f.combinadas) if (hoja.combinadasBrutas.length < MAX_COMBINADAS) hoja.combinadasBrutas.push(m);
      fila += repFila;
    };

    XmlLigero.recorrer(cuerpo, {
      abrir(nombre, a, vacio) {
        void vacio;
        if (nombre === 'table' ) {
          if (profundidadTabla++ > 0) return;   // una tabla dentro de otra (objeto incrustado): no es una hoja
          hoja = { nombre: a.name || ('Hoja' + (hojas.length + 1)), filas: [], anchos: [], combinadasBrutas: [], columnasEstilo: [], columnasOcultas: [], filasOcultas: [], oculta: estilos.hojasOcultas.has(a['style-name']) };
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
            for (let k = 0; k < rep && hoja.columnasEstilo.length < MAX_COLS; k++) {
              if (oculta) hoja.columnasOcultas.push(hoja.columnasEstilo.length);
              hoja.anchos.push(w);
              hoja.columnasEstilo.push(a['default-cell-style-name'] || null);
            }
            break;
          }
          case 'table-row':
            repFila = Math.min(entero(a['number-rows-repeated'], 1), MAX_FILAS);
            filaActual = { celdas: [], combinadas: [], oculta: a.visibility === 'collapse' || a.visibility === 'filter', estiloColumna: (k) => hoja.columnasEstilo[k] || null };
            col = 0;
            break;
          case 'table-cell':
            if (!filaActual) return;
            celda = {
              tipo: a['value-type'] || null, a, estilo: a['style-name'] || null,
              rep: Math.min(entero(a['number-columns-repeated'], 1), MAX_COLS),
              cs: entero(a['number-columns-spanned'], 1), fs: entero(a['number-rows-spanned'], 1),
            };
            enCelda = true; parrafoActual = null;
            break;
          case 'covered-table-cell':
            if (!filaActual) return;
            col += Math.min(entero(a['number-columns-repeated'], 1), MAX_COLS);
            break;
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
        else if (nombre === 'table-cell' && celda) cierraCelda();
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
        return { nombre: h.nombre, filas, anchos: h.anchos.slice(0, ancho), combinadas: h.combinadasBrutas.map(rango), estado: h.oculta ? 'oculta' : 'visible', columnasOcultas: h.columnasOcultas.filter((c) => c < ancho), filasOcultas: h.filasOcultas.filter((r) => r < filas.length) };
      }),
    };
  }

  function LibroXlsx_numeroACol(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }

  globalThis.OdsLector = { leer, _traducirEstilo: traducirEstilo, _serieDeFecha: serieDeFecha };
})();

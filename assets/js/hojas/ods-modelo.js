/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Modelo de un libro OpenDocument (.ods) con la MISMA FORMA que `LibroXlsx` (F3.4a, etapa 2): la rejilla, el motor de cálculo, el PDF y la cinta lo leen
// sin saber que el archivo es ODF. Se lee content.xml una vez (más styles.xml para los estilos con nombre) y cada hoja sale con su almacén de celdas
// (`CeldasHoja`), sus filas y columnas, sus combinadas y sus fórmulas traducidas a la sintaxis de Excel (`ods-formulas.js`).
//
//   const libro = await OdsModelo.abrir(bytes);          // → { zip, hojas: [{ nombre, id, estado, tipo, ruta }], nombres, fecha1904, estilos, cargarHoja(i), formato: 'ods', avisos, … }
//   const hoja = await libro.cargarHoja(0);               // → como la de LibroXlsx: celdas, filas (Map), columnas, combinadas, formato, estiloDe(f, c), …
//
// Lo que NO se edita se queda en el archivo tal cual (nada de esto se reescribe al guardar): imágenes, objetos incrustados, anotaciones, validación,
// formato condicional, gráficos, estilos de página. Las repeticiones de ODF (`number-columns-repeated`, `number-rows-repeated`) se expanden en celdas
// reales; las filas y columnas vacías de la cola no se materializan. Las celdas vacías solo se guardan si tienen un estilo propio distinto del de su
// columna o fila. Los valores de fecha usan la base de `table:null-date` (1904-01-01 → `fecha1904`; otra → base 1900, con la diferencia anotada en `avisos`).
// El estilo de cada celda es un número de xf de un styles.xml VIRTUAL (ver `ods-estilos.js`): `libro.estilos.xf(i)` da lo mismo que en un .xlsx.
//
// Requiere zip-propio, xml-ligero, xlsx-estilos, xlsx-modelo, ods-formulas (del sitio) y motores/ods-estilos. Sin DOM.
(function () {
  const MAX_FILAS = 1048576, MAX_COLS = 16384, MAX_TEXTO = 32767;
  const MAX_CELDAS = 10000000;    // en todo el libro
  const MAX_VACIAS = 1000000;     // celdas vacías con estilo propio en todo el libro
  const MAX_ALTOS = 100000;       // filas con propiedades por hoja
  const MAX_HOJAS = 4096, MAX_FORMULA = 32768;
  const MAX_COMBINADAS = 500000;
  const ERRORES = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA']);
  const T = () => LibroXlsx.TIPO;

  const entero = (s, defecto) => { const n = parseInt(s, 10); return n > 0 ? n : defecto; };
  // las direcciones de un atributo que trae varias («H.A1:H.A5 'Mi hoja'.C1»): la hoja entre comillas puede llevar espacios, así que no se corta por espacios
  const RE_DIRECCIONES = /\$?(?:'(?:[^']|'')+'|[^\s.']+)\.\$?[A-Za-z]{1,3}\$?\d+(?::\$?(?:'(?:[^']|'')+'|[^\s.']+)?\.\$?[A-Za-z]{1,3}\$?\d+)?/g;
  function rangosDe(texto) {   // «H.A1:H.A5 H.C1» → [{ f0, c0, f1, c1 }]
    const salida = [];
    for (const parte of String(texto || '').match(RE_DIRECCIONES) || []) {
      const celdas = [...parte.replace(/'(?:[^']|'')*'/g, '').matchAll(/\.\$?([A-Za-z]{1,3})\$?(\d+)/g)].map((m) => m[1] + m[2]);
      const r = celdas.length ? LibroXlsx.decodificarRango(celdas.length > 1 ? celdas[0] + ':' + celdas[celdas.length - 1] : celdas[0]) : null;
      if (r) salida.push(r);
    }
    return salida;
  }

  // ---- fechas y horas de ODF → número de serie ------------------------------------------------------------------------------------
  // base: 'iso' de table:null-date. La base 1904 cuenta los días desde 1904-01-01; la base 1900 de Excel desde 1899-12-30 y regala el 29-feb-1900 que no existió.
  function diasDesde(a, mes, d, baseMs) {
    const t = new Date(0); t.setUTCFullYear(a, mes - 1, d); t.setUTCHours(0, 0, 0, 0);
    return Math.round((t.getTime() - baseMs) / 86400000);
  }
  function serieDeFecha(texto, fecha1904) {
    const m = /^(-?\d{1,6})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2}(?:\.\d+)?))?/.exec(texto);
    if (!m) return null;
    const a = +m[1], mes = +m[2], d = +m[3];
    let dias;
    if (fecha1904) dias = diasDesde(a, mes, d, Date.UTC(1904, 0, 1));
    else if (a === 1900 && mes === 2 && d === 29) dias = 60;   // el día que Excel inventó
    else { dias = diasDesde(a, mes, d, Date.UTC(1899, 11, 30)); if (dias < 61) dias -= 1; }   // antes del 1900-03-01 Excel cuenta un día menos
    const hora = m[4] === undefined ? 0 : (+m[4] * 3600 + +m[5] * 60 + parseFloat(m[6])) / 86400;
    return { serie: dias + hora, conHora: m[4] !== undefined && (+m[4] || +m[5] || +m[6]) };
  }
  function serieDeHora(texto) {
    const m = /^(-)?P(?:(\d+)D)?(?:T(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?)?$/.exec(texto);
    if (!m) return null;
    const s = (+(m[2] || 0)) * 86400 + (+(m[3] || 0)) * 3600 + (+(m[4] || 0)) * 60 + +(m[5] || 0);
    return m[1] ? -s / 86400 : s / 86400;
  }

  async function abrir(origen) {
    const bytes = origen instanceof Uint8Array ? origen : new Uint8Array(await origen.arrayBuffer());
    let zip;
    try { zip = await ZipPropio.loadAsync(bytes); } catch (e) { throw new LibroXlsx.ErrorXlsx('noZip', 'no es OpenDocument'); }
    const contenido = zip.file('content.xml');
    // LibreOffice 26 (ODF 1.4) cifra el paquete entero: un solo archivo «encrypted-package»
    if (!contenido && zip.file('encrypted-package')) throw new LibroXlsx.ErrorXlsx('cifrado');
    if (!contenido) throw new LibroXlsx.ErrorXlsx('noZip', 'no es OpenDocument');
    if (zip.file('META-INF/manifest.xml')) {
      const man = await zip.file('META-INF/manifest.xml').async('string');
      if (/<(?:[\w.-]+:)?encryption-data\b/.test(man)) throw new LibroXlsx.ErrorXlsx('cifrado');
    }
    const texto = await contenido.async('string');
    const estilos = OdsEstilos.crear();
    if (zip.file('styles.xml')) OdsEstilos.leer(await zip.file('styles.xml').async('string'), estilos);
    // los estilos automáticos van antes del cuerpo (se leen aparte, con el lector de estilos)
    const a0 = texto.indexOf('<office:automatic-styles'), a1 = texto.indexOf('</office:automatic-styles>');
    if (a0 >= 0 && a1 > a0) OdsEstilos.leer(texto.slice(a0, a1 + 26), estilos);
    const cuerpo = a0 >= 0 && a1 > a0 ? texto.slice(a1 + 26) : texto;

    const tabla = OdsEstilos.tablaXf(estilos);
    const avisos = { formulasSinTraducir: 0, baseDeFechas: null, celdasPerdidas: 0 };
    const libro = {
      formato: 'ods', zip, rutaLibro: 'content.xml', relaciones: [], hojas: [], nombres: [], fecha1904: false, estilos: null, textos: [], vistaLibro: { hojaActiva: 0, primeraHoja: 0 },
      calculo: null, propiedades: {}, tieneMacros: false, rutaTema: null, rutaPersonas: null, _hojasCargadas: new Map(), avisos, _tablaXf: tabla, _estilosOds: estilos,
    };
    const modelos = [];

    // validación de datos: las definiciones (table:content-validation) y las celdas que las usan (table:content-validation-name)
    const defsValidacion = new Map(), usosValidacion = [];
    let valActual = null, mensajeVal = null;
    // formato condicional (calcext:conditional-format): un bloque por elemento, con sus reglas en el orden del archivo (el primero manda)
    let cfActual = null, cfDestino = null, prioridadCf = 0;
    let dbActual = null; const bdsFiltro = [];
    const infoHojas = [];   // por hoja: su estilo de tabla (la página maestra), el área de impresión y las filas que se repiten
    let titulosDesde = -1;
    let marcoActual = null;   // un draw:frame con un objeto incrustado (un gráfico): dónde está y cuánto mide
    const colorRef = (hex) => (/^#[0-9a-fA-F]{6}$/.test(hex || '') ? { tipo: 'rgb', valor: hex.slice(1).toUpperCase(), tint: 0 } : null);
    const reglaCf = (tipo, extra) => Object.assign({ tipo, prioridad: ++prioridadCf, dxf: null, operador: null, texto: null, detener: false, encima: true, igual: false, rango: null, porcentaje: false, abajo: false, periodo: null, desvEst: null, formulas: [], escala: null, barras: null, iconos: null }, extra || {});
    let hoja = null, C = null, fila = 0, col = 0, repFila = 1, filaActual = null, celda = null, profundidadTabla = 0, enNota = 0;
    let enCelda = false, parrafos = [], parrafoActual = null, totalCeldas = 0, totalVacias = 0;
    let colRuns = [], nCols = 0, filasConPropiedades = 0;
    const combinadasVivas = [];   // combinadas que aún cubren la fila de abajo (para distinguir las celdas cubiertas de las partes de una matriz)

    const runDe = (c) => {
      let a = 0, b = colRuns.length - 1, r = null;
      while (a <= b) { const m = (a + b) >> 1; if (colRuns[m].ini <= c) { r = colRuns[m]; a = m + 1; } else b = m - 1; }
      return r && c < r.ini + r.n ? r : null;
    };
    // el nombre del estilo que le toca a una celda sin estilo propio: el de su columna si la columna trae uno; si no, el de su fila
    const nombreDefecto = (c) => { const r = runDe(c); return r && r.nombre ? r.nombre : filaActual && filaActual.nombre ? filaActual.nombre : null; };

    const nuevaHoja = (a) => {
      if (modelos.length >= MAX_HOJAS) throw new LibroXlsx.ErrorXlsx('grande', 'más de ' + MAX_HOJAS + ' hojas');
      const h = LibroXlsx.hojaVacia();
      h.sinCeldas = false;
      h.formato.altoPorDefecto = null;
      libro.hojas.push({ nombre: a.name || ('Hoja' + (libro.hojas.length + 1)), id: libro.hojas.length + 1, estado: estilos.hojasOcultas.has(a['style-name']) ? 'oculta' : 'visible', tipo: 'hoja', ruta: 'content.xml', rid: null });
      h.nombre = libro.hojas[libro.hojas.length - 1].nombre;
      h.indice = libro.hojas.length - 1;
      modelos.push(h);
      h._marcos = [];
      infoHojas.push({ estilo: a['style-name'] || null, impresion: a['print-ranges'] || null, titulos: null });
      hoja = h; C = h.celdas; fila = 0; col = 0; colRuns = []; nCols = 0; filasConPropiedades = 0; combinadasVivas.length = 0;
      const celdasH = h.celdas;   // (C cambia al abrir la hoja siguiente: la función se queda con las de esta)
      h.estiloDe = (f, c) => {
        const k = celdasH.indiceDe(f, c);
        if (k >= 0) return celdasH.estilo[k];
        const x = h.filas.get(f);
        if (x && x.formatoPersonalizado) return x.estilo;
        const r = runDe(c);
        return r ? r.xf : 0;
      };
    };
    const cerrarHoja = () => {
      for (const r of colRuns) hoja.columnas.push({ min: r.ini, max: r.ini + r.n - 1, ancho: r.px > 0 ? r.px / 7 : null, oculta: r.oculta, personalizado: r.px > 0, estilo: r.xf, nivel: 0, colapsada: false, ajusteAuto: false });
      C._indexar();
      hoja = null; C = null;
    };

    const cierraCelda = () => {
      const c = celda; celda = null; enCelda = false;
      const rep = c.rep, a = c.a;
      if (a['content-validation-name'] && col < MAX_COLS) usosValidacion.push({ nombre: a['content-validation-name'], hoja: modelos.length - 1, f0: fila, f1: Math.min(fila + repFila - 1, MAX_FILAS - 1), c0: col, c1: Math.min(col + rep, MAX_COLS) - 1 });
      let tipo = T().vacio, valor = null, def = null, hayValor = false;
      const cadena = () => { let t = a['string-value'] !== undefined ? a['string-value'] : parrafos.join('\n'); if (t.length > MAX_TEXTO) t = t.slice(0, MAX_TEXTO); return t; };
      switch (c.tipo) {
        case 'float': case 'currency': { const n = parseFloat(a.value); if (Number.isFinite(n)) { tipo = T().numero; valor = n; hayValor = true; } break; }
        case 'percentage': { const n = parseFloat(a.value); if (Number.isFinite(n)) { tipo = T().numero; valor = n; def = '0%'; hayValor = true; } break; }
        case 'date': { const d = serieDeFecha(a['date-value'] || '', libro.fecha1904); if (d) { tipo = T().numero; valor = d.serie; def = d.conHora ? 'yyyy-mm-dd hh:mm:ss' : 'yyyy-mm-dd'; hayValor = true; } break; }
        case 'time': { const h = serieDeHora(a['time-value'] || ''); if (h !== null) { tipo = T().numero; valor = h; def = 'hh:mm:ss'; hayValor = true; } break; }
        case 'boolean': tipo = T().logico; valor = a['boolean-value'] === 'true'; hayValor = true; break;
        case 'error': { const e = a['string-value'] || parrafos.join(''); tipo = T().error; valor = ERRORES.has(e) ? e : '#VALUE!'; hayValor = true; break; }
        case 'string': { const t = cadena(); if (t !== '' || a.formula) { tipo = T().texto; valor = t; hayValor = true; } break; }
        default: break;
      }
      parrafos = [];
      // la fórmula (OpenFormula → texto de Excel); si no se puede llevar, queda el último valor y se anota
      let formula = null, matriz = null;
      if (a.formula) {
        if (a.formula.length <= MAX_FORMULA && a.formula.startsWith('of:=') && !/^of:=#[A-Za-z\/0!?]+$/.test(a.formula)) {
          const t = OdsFormulas.aExcel(a.formula);
          if (t && t.length < 8000) {
            formula = t;
            const mc = entero(a['number-matrix-columns-spanned'], 0), mf = entero(a['number-matrix-rows-spanned'], 0);
            if (mc && mf) matriz = LibroXlsx.numeroACol(col) + (fila + 1) + (mc > 1 || mf > 1 ? ':' + LibroXlsx.numeroACol(Math.min(col + mc - 1, MAX_COLS - 1)) + Math.min(fila + mf, MAX_FILAS) : '');
          } else avisos.formulasSinTraducir++;
        } else if (!/^of:=#/.test(a.formula)) avisos.formulasSinTraducir++;
      }
      // las celdas combinadas: la ancla anota su rango
      if ((c.fs > 1 || c.cs > 1) && !c.cubierta && hoja.combinadas.length < MAX_COMBINADAS && col < MAX_COLS) {
        const m = { f0: fila, c0: col, f1: Math.min(fila + c.fs - 1, MAX_FILAS - 1), c1: Math.min(col + c.cs - 1, MAX_COLS - 1) };
        hoja.combinadas.push(m); combinadasVivas.push(m);
      }
      const fin = Math.min(col + rep, MAX_COLS);
      if (!hayValor && !formula) {
        // vacía: solo se guarda con un estilo propio distinto del que le toca (los vacíos de la cola con estilo de columna no se materializan)
        if (c.estilo) {
          const xf = tabla.indice(c.estilo, null), xfDef = tabla.indice(nombreDefecto(col), null);
          if (xf !== xfDef) {
            const n = Math.min(fin - col, MAX_VACIAS - totalVacias);
            for (let k = 0; k < n; k++) filaActual.celdas.push({ c: col + k, tipo: T().vacio, valor: null, xf, formula: null, matriz: null });
            if (n > 0) { totalVacias += n; if (n < fin - col) avisos.celdasPerdidas++; }
          }
        }
        col += rep;
        return;
      }
      const xf = tabla.indice(c.estilo || nombreDefecto(col), def);
      for (let x = col; x < fin; x++) filaActual.celdas.push({ c: x, tipo, valor, xf, formula, matriz: x === col ? matriz : null });
      col += rep;
    };

    const cierraFila = () => {
      const f = filaActual; filaActual = null;
      let n = Math.max(1, Math.min(repFila, MAX_FILAS - fila));
      if (fila >= MAX_FILAS) { fila += repFila; return; }
      if (f.celdas.length) {
        n = Math.min(n, Math.max(1, Math.floor((MAX_CELDAS - totalCeldas) / f.celdas.length)));
        if ((totalCeldas += n * f.celdas.length) > MAX_CELDAS) throw new LibroXlsx.ErrorXlsx('grande', 'más de ' + MAX_CELDAS + ' celdas');
        for (let k = 0; k < n; k++) {
          for (const x of f.celdas) {
            const i = C.agregar(fila + k, x.c, x.tipo, x.valor, x.xf);
            if (x.formula) {
              // una matriz de varias celdas se lee como dinámica (LibreOffice las recalcula): al cambiar sus datos se vuelve a calcular y desborda, y sus celdas no quedan protegidas
              const dinamica = !!(x.matriz && k === 0 && x.matriz.indexOf(':') > 0);
              C.formulas.set(i, { texto: x.formula, tipo: x.matriz ? 'array' : 'normal', ref: k === 0 ? x.matriz : null, si: null, dinamica });
              if (dinamica) {
                const g = LibroXlsx.decodificarRango(x.matriz);
                if (g) {
                  if (!hoja.derramadas) hoja.derramadas = new Set();
                  for (let ff = g.f0; ff <= g.f1; ff++) for (let cc = g.c0; cc <= g.c1; cc++) if (ff !== g.f0 || cc !== g.c0) hoja.derramadas.add(ff * MAX_COLS + cc);
                }
              }
            }
          }
        }
      }
      if ((f.oculta || f.alto || f.nombre) && filasConPropiedades < MAX_ALTOS) {
        const copias = Math.min(repFila, MAX_ALTOS - filasConPropiedades);
        for (let k = 0; k < copias && fila + k < MAX_FILAS; k++, filasConPropiedades++) {
          hoja.filas.set(fila + k, { alto: f.alto || null, oculta: f.oculta, personalizada: !!f.alto, estilo: f.nombre ? tabla.indice(f.nombre, null) : 0, formatoPersonalizado: !!f.nombre, nivel: 0, colapsada: false });
        }
      }
      fila += repFila;
      for (let k = combinadasVivas.length - 1; k >= 0; k--) if (combinadasVivas[k].f1 < fila) combinadasVivas.splice(k, 1);
    };

    // fecha base: table:null-date viene antes de las hojas
    XmlLigero.recorrer(cuerpo, {
      abrir(nombre, a) {
        if (profundidadTabla === 0) {   // las definiciones de validación van antes de las hojas
          // autofiltro: un rango de base de datos con botones de filtro (los criterios están en sus table:filter-condition)
          if (nombre === 'database-range') { dbActual = a['display-filter-buttons'] === 'true' && a['target-range-address'] ? { destino: a['target-range-address'], campos: new Set() } : null; if (dbActual) bdsFiltro.push(dbActual); return; }
          if (nombre === 'filter-condition' && dbActual) { const k = parseInt(a['field-number'], 10); if (k >= 0 && dbActual.campos.size < 16384) dbActual.campos.add(k); return; }
          if (nombre === 'content-validation' && a.name) { valActual = { a, ayuda: null, error: null }; defsValidacion.set(a.name, valActual); return; }
          if (valActual && (nombre === 'help-message' || nombre === 'error-message')) { const m = { a, texto: '' }; if (nombre === 'help-message') valActual.ayuda = m; else valActual.error = m; mensajeVal = m; return; }
        }
        if (nombre === 'null-date') {
          const v = a['date-value'] || '';
          if (v.slice(0, 10) === '1904-01-01') libro.fecha1904 = true;
          else if (v.slice(0, 10) !== '1899-12-30') avisos.baseDeFechas = v;   // base poco común (p. ej. 1900-01-01): se lee como base 1900
          return;
        }
        if (nombre === 'table') {
          if (profundidadTabla++ > 0) return;   // una tabla dentro de otra (objeto incrustado): no es una hoja
          nuevaHoja(a);
          return;
        }
        if (nombre === 'named-range' && a['cell-range-address'] && a.name) {
          const ref = OdsFormulas.aExcel('[' + a['cell-range-address'] + ']');
          if (ref) libro.nombres.push({ nombre: a.name, hojaLocal: profundidadTabla ? modelos.length - 1 : null, oculto: false, comentario: '', ref });
          return;
        }
        if (nombre === 'named-expression' && a.expression && a.name) {   // un nombre que vale una fórmula («Tasa» = 0,13)
          const ref = OdsFormulas.aExcel(a.expression.startsWith('of:=') ? a.expression : 'of:=' + a.expression);
          if (ref) libro.nombres.push({ nombre: a.name, hojaLocal: profundidadTabla ? modelos.length - 1 : null, oculto: false, comentario: '', ref });
          return;
        }
        if (!hoja || profundidadTabla !== 1) return;
        if (nombre === 'table-header-rows') { titulosDesde = fila; return; }
        if (nombre === 'annotation') { enNota++; parrafoActual = null; return; }   // el comentario de una celda no es su texto
        if (enNota) return;
        switch (nombre) {
          case 'table-column': {
            const rep = Math.min(entero(a['number-columns-repeated'], 1), MAX_COLS - nCols);
            if (rep <= 0) break;
            const px = a['style-name'] ? estilos.columnas.get(a['style-name']) || 0 : 0;
            const oculta = a.visibility === 'collapse' || a.visibility === 'filter';
            const nombreEst = a['default-cell-style-name'] || null;
            const xf = nombreEst ? tabla.indice(nombreEst, null) : 0;
            const u = colRuns[colRuns.length - 1];
            if (u && u.px === px && u.oculta === oculta && u.nombre === nombreEst) u.n += rep;
            else colRuns.push({ ini: nCols, n: rep, px, oculta, nombre: nombreEst, xf });
            nCols += rep;
            break;
          }
          case 'table-row': {
            repFila = Math.min(entero(a['number-rows-repeated'], 1), MAX_FILAS);
            const ef = a['style-name'] ? estilos.filas.get(a['style-name']) : null;
            filaActual = { celdas: [], oculta: a.visibility === 'collapse' || a.visibility === 'filter', nombre: a['default-cell-style-name'] || null, alto: ef && !ef.optimo ? ef.alto : 0 };
            col = 0;
            break;
          }
          case 'table-cell':
            if (!filaActual) return;
            celda = { tipo: a['value-type'] || null, a, estilo: a['style-name'] || null, rep: Math.min(entero(a['number-columns-repeated'], 1), MAX_COLS), cs: entero(a['number-columns-spanned'], 1), fs: entero(a['number-rows-spanned'], 1) };
            enCelda = true; parrafoActual = null;
            break;
          case 'covered-table-cell': {
            if (!filaActual) return;
            // una celda cubierta con valor es parte del resultado de una fórmula de matriz (las cubiertas por una combinada no traen valor): se lee como cualquier celda
            if (a['value-type'] && !combinadasVivas.some((m) => m.f0 <= fila && fila <= m.f1 && m.c0 <= col && col <= m.c1)) {
              celda = { tipo: a['value-type'], a, estilo: a['style-name'] || null, rep: Math.min(entero(a['number-columns-repeated'], 1), MAX_COLS), cs: 1, fs: 1, cubierta: true };
              enCelda = true; parrafoActual = null;
              break;
            }
            col += Math.min(entero(a['number-columns-repeated'], 1), MAX_COLS);
            break;
          }
          case 'frame': marcoActual = { fila, col, enCelda: !!celda, x: a.x, y: a.y, w: a.width, h: a.height, nombre: a.name || '', href: null, z: a['z-index'] }; break;
          case 'object': if (marcoActual && a.href && !marcoActual.href) marcoActual.href = a.href; break;
          case 'conditional-format': cfActual = globalThis.OdsCondicional ? { rangos: rangosDe(a['target-range-address']), reglas: [] } : null; break;
          case 'condition':
            if (cfActual && globalThis.OdsCondicional) {
              const r = OdsCondicional.leerValor(a.value);
              if (!r) break;
              const base = /\.\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a['base-cell-address'] || '');
              const r0 = cfActual.rangos[0];
              const df = base && r0 ? r0.f0 - (+base[2] - 1) : 0, dc = base && r0 ? r0.c0 - LibroXlsx.colANumero(base[1].toUpperCase()) : 0;
              const mover = (x) => (df || dc ? LibroXlsx.desplazarFormula(x, df, dc) : x);
              let formulasCf = (r.formulas || []).map(mover);
              if (r.texto !== undefined && cfActual.rangos[0]) {
                const celda = LibroXlsx.codificarDireccion(cfActual.rangos[0].f0, cfActual.rangos[0].c0), lit = '"' + String(r.texto).replace(/"/g, '""') + '"';
                formulasCf = [{ containsText: 'NOT(ISERROR(SEARCH(' + lit + ',' + celda + ')))', notContainsText: 'ISERROR(SEARCH(' + lit + ',' + celda + '))', beginsWith: 'LEFT(' + celda + ',LEN(' + lit + '))=' + lit, endsWith: 'RIGHT(' + celda + ',LEN(' + lit + '))=' + lit }[r.tipo]];
              }
              cfActual.reglas.push(reglaCf(r.tipo, { dxf: a['apply-style-name'] ? tabla.dxf(a['apply-style-name']) : null, operador: r.operador || null, texto: r.texto === undefined ? null : r.texto, encima: r.encima !== false,
                igual: !!r.igual, rango: r.rango === undefined ? null : r.rango, porcentaje: !!r.porcentaje, abajo: !!r.abajo, formulas: formulasCf }));
            }
            break;
          case 'color-scale': if (cfActual) cfDestino = { tipo: 'colorScale', escala: { cfvo: [], colores: [] } }; break;
          case 'color-scale-entry': if (cfDestino && cfDestino.escala) { cfDestino.escala.cfvo.push({ tipo: OdsCondicional.aCfvo(a.type), valor: a.value !== undefined ? a.value : null, gte: true }); cfDestino.escala.colores.push(colorRef(a.color)); } break;
          case 'data-bar': if (cfActual) cfDestino = { tipo: 'dataBar', barras: { cfvo: [], color: colorRef(a['positive-color']), mostrarValor: true, min: a['min-length'] !== undefined ? parseInt(a['min-length'], 10) : 10, max: a['max-length'] !== undefined ? parseInt(a['max-length'], 10) : 90 } }; break;
          case 'icon-set': if (cfActual) cfDestino = { tipo: 'iconSet', iconos: { conjunto: a['icon-set-type'] || '3TrafficLights1', cfvo: [], mostrarValor: a['show-value'] !== 'false', invertir: false, soloIcono: false } }; break;
          case 'formatting-entry': if (cfDestino && (cfDestino.barras || cfDestino.iconos)) (cfDestino.barras || cfDestino.iconos).cfvo.push({ tipo: OdsCondicional.aCfvo(a.type), valor: a.value !== undefined ? a.value : null, gte: true }); break;
          case 'p': if (enCelda) parrafoActual = ''; break;
          case 's': if (parrafoActual !== null) parrafoActual += ' '.repeat(Math.min(entero(a.c, 1), 10000)); break;
          case 'tab': if (parrafoActual !== null) parrafoActual += '\t'; break;
          case 'line-break': if (parrafoActual !== null) parrafoActual += '\n'; break;
          default: break;
        }
      },
      cerrar(nombre) {
        if (nombre === 'help-message' || nombre === 'error-message') mensajeVal = null;
        else if (nombre === 'content-validation') valActual = null;
        else if (nombre === 'database-range') dbActual = null;
        if (nombre === 'table') { if (--profundidadTabla === 0 && hoja) cerrarHoja(); return; }
        if (!hoja || profundidadTabla !== 1) return;
        if (nombre === 'annotation') { enNota--; return; }
        if (nombre === 'table-header-rows') { if (titulosDesde >= 0 && fila > titulosDesde) infoHojas[infoHojas.length - 1].titulos = { f0: titulosDesde, f1: fila - 1 }; titulosDesde = -1; return; }
        if (enNota) return;
        if (nombre === 'frame') { if (marcoActual && marcoActual.href && hoja._marcos.length < 500) hoja._marcos.push(marcoActual); marcoActual = null; return; }
        if (nombre === 'color-scale' || nombre === 'data-bar' || nombre === 'icon-set') { if (cfActual && cfDestino) cfActual.reglas.push(reglaCf(cfDestino.tipo, { escala: cfDestino.escala || null, barras: cfDestino.barras || null, iconos: cfDestino.iconos || null })); cfDestino = null; return; }
        if (nombre === 'conditional-format') { if (cfActual && cfActual.rangos.length && cfActual.reglas.length && hoja.formatosCondicionales.length < MAX_COMBINADAS) hoja.formatosCondicionales.push(cfActual); cfActual = null; return; }
        if (nombre === 'p' && enCelda && parrafoActual !== null) { parrafos.push(parrafoActual); parrafoActual = null; }
        else if ((nombre === 'table-cell' || (nombre === 'covered-table-cell' && celda)) && celda) cierraCelda();
        else if (nombre === 'table-row' && filaActual) cierraFila();
      },
      texto(t) { if (parrafoActual !== null) parrafoActual += t; else if (mensajeVal) mensajeVal.texto += t; },
    });

    if (!modelos.length) throw new LibroXlsx.ErrorXlsx('sinLibro', 'sin hojas');
    // validación de datos: los rectángulos de celdas de cada definición, por hoja (las filas seguidas con las mismas columnas se juntan)
    if (usosValidacion.length && globalThis.OdsValidacion) {
      const porClave = new Map();
      for (const u of usosValidacion) { const k = u.nombre + '|' + u.hoja; let l = porClave.get(k); if (!l) porClave.set(k, l = []); l.push(u); }
      for (const [k, lista] of porClave) {
        const def = defsValidacion.get(lista[0].nombre);
        if (!def) continue;
        lista.sort((x, y) => x.c0 - y.c0 || x.c1 - y.c1 || x.f0 - y.f0);
        const rangos = [];
        for (const u of lista) {
          const ult = rangos[rangos.length - 1];
          if (ult && ult.c0 === u.c0 && ult.c1 === u.c1 && u.f0 <= ult.f1 + 1) ult.f1 = Math.max(ult.f1, u.f1); else rangos.push({ f0: u.f0, c0: u.c0, f1: u.f1, c1: u.c1 });
        }
        const r = OdsValidacion.leer(def.a.condition);
        if (!r) continue;
        const base = /\$?([A-Za-z]+)\$?(\d+)$/.exec(def.a['base-cell-address'] || '');
        const df = base ? rangos[0].f0 - (+base[2] - 1) : 0, dc = base ? rangos[0].c0 - LibroXlsx.colANumero(base[1].toUpperCase()) : 0;
        const mover = (x) => (x && (df || dc) ? LibroXlsx.desplazarFormula(x, df, dc) : x);
        const hoja = modelos[lista[0].hoja];
        if (hoja.validaciones.length >= MAX_COMBINADAS) continue;
        hoja.validaciones.push({ rangos, tipo: r.tipo, operador: r.operador, permitirVacio: def.a['allow-empty-cell'] !== 'false', ocultarFlecha: def.a['display-list'] === 'none', mostrarEntrada: !!def.ayuda && def.ayuda.a.display === 'true',
          mostrarError: !!def.error && def.error.a.display !== 'false', estiloError: (def.error && def.error.a['message-type']) || 'stop', tituloError: (def.error && def.error.a.title) || '', error: def.error ? def.error.texto : '',
          tituloEntrada: (def.ayuda && def.ayuda.a.title) || '', entrada: def.ayuda ? def.ayuda.texto : '', formula1: mover(r.formula1), formula2: mover(r.formula2), nombreOds: def.a.name });
      }
    }
    // autofiltro de la hoja: el primer rango con botones de cada hoja (ref y las columnas con criterios)
    for (const b of bdsFiltro) {
      const m = /^\$?(?:'((?:[^']|'')+)'|([^.']+))\.(\$?[A-Za-z]{1,3}\$?\d+)(?::\$?(?:'(?:[^']|'')+'|[^.']+)?\.(\$?[A-Za-z]{1,3}\$?\d+))?$/.exec(b.destino);
      if (!m) continue;
      const nombreHoja = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
      const h = modelos.find((x) => x.nombre.toLowerCase() === nombreHoja.toLowerCase());
      const ref = (m[3] + (m[4] ? ':' + m[4] : '')).replace(/\$/g, '');
      if (!h || h.filtroAuto || !LibroXlsx.decodificarRango(ref)) continue;
      h.filtroAuto = { ref, columnas: new Map([...b.campos].map((k) => [k, { col: k, oculto: false, filtrada: true }])) };
    }
    // página e impresión: la página maestra de cada hoja (styles.xml) da el papel, los márgenes, el encabezado y el pie; el área y los títulos vienen de la tabla
    if (globalThis.OdsPagina && zip.file('styles.xml')) {
      const { maestras, layouts } = OdsPagina.leerLibro(await zip.file('styles.xml').async('string'));
      const primera = maestras.size ? maestras.keys().next().value : null;
      const refDe = (r, hojaNombre) => { const n = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(hojaNombre) ? hojaNombre : "'" + hojaNombre.replace(/'/g, "''") + "'"; return n + '!$' + LibroXlsx.numeroACol(r.c0) + '$' + (r.f0 + 1) + ':$' + LibroXlsx.numeroACol(r.c1) + '$' + (r.f1 + 1); };
      modelos.forEach((hj, i) => {
        const info = infoHojas[i], nombreM = (info.estilo && estilos.paginaDeTabla.get(info.estilo)) || primera;
        const maestra = nombreM ? maestras.get(nombreM) || null : null;
        if (maestra) {
          const cfg = OdsPagina.configuracion(maestra, maestra.layout ? layouts.get(maestra.layout) || null : null);
          hj.imprimir = Object.assign(hj.imprimir, OdsPagina.imprimirDe(cfg, maestra));
        }
        const areas = OdsPagina.areaDe(info.impresion).map((a) => LibroXlsx.decodificarRango(a)).filter(Boolean);
        if (areas.length) libro.nombres.push({ nombre: '_xlnm.Print_Area', hojaLocal: i, oculto: false, comentario: '', ref: areas.map((r) => refDe(r, hj.nombre)).join(',') });
        if (info.titulos) libro.nombres.push({ nombre: '_xlnm.Print_Titles', hojaLocal: i, oculto: false, comentario: '', ref: hj.nombre.replace(/^(?![A-Za-z_][A-Za-z0-9_.]*$)(.*)$/, (m) => "'" + m.replace(/'/g, "''") + "'") + '!$' + (info.titulos.f0 + 1) + ':$' + (info.titulos.f1 + 1) });
      });
    }
    // paneles inmovilizados: settings.xml guarda por hoja el modo de división (2 = fijo) y cuántas columnas y filas quedan fijas
    const ajustes = zip.file('settings.xml');
    if (ajustes) {
      const porNombre = new Map(modelos.map((h) => [h.nombre, h]));
      let enTablas = false, hojaAjustes = null, item = null, valores = null;
      XmlLigero.recorrer(await ajustes.async('string'), {
        abrir(nombre, a) {
          if (nombre === 'config-item-map-named') enTablas = a.name === 'Tables';
          else if (nombre === 'config-item-map-entry' && enTablas && a.name !== undefined) { hojaAjustes = porNombre.get(a.name) || null; valores = {}; }
          else if (nombre === 'config-item' && hojaAjustes) item = { nombre: a.name, texto: '' };
        },
        texto(t) { if (item) item.texto += t; },
        cerrar(nombre) {
          if (nombre === 'config-item' && item) { if (valores) valores[item.nombre] = item.texto.trim(); item = null; }
          else if (nombre === 'config-item-map-entry' && hojaAjustes) {
            const columnas = valores.HorizontalSplitMode === '2' ? parseInt(valores.HorizontalSplitPosition, 10) || 0 : 0, filas = valores.VerticalSplitMode === '2' ? parseInt(valores.VerticalSplitPosition, 10) || 0 : 0;
            if (columnas || filas) hojaAjustes.vista.paneles = { x: columnas, y: filas, celdaSuperior: LibroXlsx.codificarDireccion(filas, columnas), activo: 'bottomRight', estado: 'frozen' };
            hojaAjustes = null; valores = null;
          } else if (nombre === 'config-item-map-named') enTablas = false;
        },
      });
    }
    libro.estilos = XlsxEstilos.leer(tabla.xml(), null);
    modelos.forEach((h, i) => libro._hojasCargadas.set(i, h));
    libro.cargarHoja = async (i) => {
      if (!libro._hojasCargadas.has(i)) throw new RangeError('hoja ' + i);
      return libro._hojasCargadas.get(i);
    };
    return libro;
  }

  globalThis.OdsModelo = { abrir, serieDeFecha, serieDeHora, rangosDe, RE_DIRECCIONES };
})();

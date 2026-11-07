/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Guardado SIN PÉRDIDA de celdas editadas en content.xml de un .ods (F3.4a, etapa 3). Lo que no se edita sale byte a byte como vino: solo se reemplazan los elementos
// de las celdas tocadas, partiendo las repeticiones de ODF (`number-rows-repeated`, `number-columns-repeated`) en tres trozos —los de antes, la celda editada y los de
// después— que copian los demás atributos. La celda editada conserva lo que no maneja el editor (combinación, validación, protección, anotaciones y dibujos).
//
//   const r = OdsEscritorCeldas.guardar({ texto, porHoja, libro, hojas, tabla, contadores, estilos, idioma })
//     texto: content.xml;  porHoja: [[indiceHoja, [{ f, c, valor, estilo }]]] (lo que arma LibroEditable.guardar);  tabla: OdsEstilos.tablaXf;  hojas: modelos de hoja
//     contadores: { estilo, numero } (persisten entre guardados, para no repetir nombres de estilos automáticos);  estilos: OdsEstilos.crear() (nombres ya usados)
//   → { texto (content.xml nuevo), avisos: { formulasNoGuardadas } }
//
// Se agregan al FINAL de office:automatic-styles los estilos de celda y de número que se hayan creado (nombres propios `mcsN`, `mcnN`): los de LibreOffice no se renumeran.
// Requiere xml-editar, ods-formulas, motores/ods-estilos-escritura y, para el texto que se ve en la celda, xlsx-formato (si está).
(function () {
  const X = () => globalThis.XmlEditar;
  const E = () => globalThis.OdsEstilosEscritura;
  const MAX_COLS = 16384;
  const GRUPOS_FILAS = new Set(['table-row-group', 'table-header-rows', 'table-rows']);
  // atributos de una celda que depende del valor o del número de repeticiones: se quitan al reescribirla y se vuelven a poner; el resto se conserva
  const ATRIBUTOS_VALOR = ['office:value-type', 'office:value', 'office:date-value', 'office:time-value', 'office:boolean-value', 'office:string-value', 'office:currency',
    'calcext:value-type', 'table:formula', 'table:number-matrix-columns-spanned', 'table:number-matrix-rows-spanned', 'table:number-columns-repeated', 'table:style-name'];

  const repeticiones = (nodo, atributo) => { const n = parseInt(nodo.attrs[atributo], 10); return n > 0 ? n : 1; };
  const sinCierre = (etiqueta) => etiqueta.replace(/\s*\/\s*>$/, '>');

  // Los elementos `elemento` de una tabla en orden (filas o definiciones de columna), cada uno con el número de la primera fila (columna) que cubre y cuántas cubre;
  // las posiciones son del texto completo. Entra en los grupos (`table-row-group`, `table-header-rows`…) aunque el primer análisis no llegara a sus hijos.
  function enumerar(texto, tabla, elemento, grupos, atributo) {
    const salida = [];
    let numero = 0;
    const recorrer = (nodo, base) => {
      for (const h of nodo.hijos) {
        if (h.local === elemento) {
          const n = repeticiones(h, atributo);
          salida.push({ nodo: { nombre: h.nombre, attrs: h.attrs, local: h.local, ini: h.ini + base, finApertura: h.finApertura + base, iniCierre: h.iniCierre < 0 ? -1 : h.iniCierre + base, fin: h.fin + base }, f0: numero, n });
          numero += n;
        } else if (grupos.has(h.local)) {
          if (h.guardado === false || !h.hijos.length) {
            const sub = X().analizar(texto.slice(h.ini + base, h.fin + base), { profundidad: 4 });
            recorrer(sub, h.ini + base);
          } else recorrer(h, base);
        }
      }
    };
    recorrer(tabla, 0);
    return salida;
  }
  const GRUPOS_COLUMNAS = new Set(['table-columns', 'table-header-columns', 'table-column-group']);
  const enumerarFilas = (texto, tabla) => enumerar(texto, tabla, 'table-row', GRUPOS_FILAS, 'table:number-rows-repeated');
  const enumerarColumnas = (texto, tabla) => enumerar(texto, tabla, 'table-column', GRUPOS_COLUMNAS, 'table:number-columns-repeated');
  // dónde van las filas nuevas detrás de la última: fuera del grupo que la envuelva (filas de título, grupo de esquema), para no meterlas en él
  const finFuera = (tabla, nodo) => { let pos = nodo.fin; for (const h of tabla.hijos) if (h.ini <= nodo.ini && nodo.fin <= h.fin) pos = Math.max(pos, h.fin); return pos; };

  // ---- el motor: reescribe las celdas de una lista (f, c, …) dentro de una tabla; `render(original, textoOriginal, edicion)` da el XML de cada celda ----------------
  const celdaVacia = (n) => '<table:table-cell' + (n > 1 ? ' table:number-columns-repeated="' + n + '"' : '') + '/>';
  // una copia del elemento con otro número de repeticiones
  const copiaConRepeticion = (t, nodo, atributo, n) => {
    const etiqueta = X().ponerAtributo(t.slice(nodo.ini, nodo.finApertura), atributo, n > 1 ? String(n) : null);
    return etiqueta + t.slice(nodo.finApertura, nodo.fin);
  };

  function crearMotor(texto) {
    function filaEditada(nodo, edits, render) {
      const t = texto.slice(nodo.ini, nodo.fin);
      const raizF = X().analizar(t, { profundidad: 2 });
      const apertura = X().ponerAtributo(t.slice(0, raizF.finApertura), 'table:number-rows-repeated', null);
      const columnas = [...edits.keys()].sort((a, b) => a - b);
      const ed = [];
      let col = 0, k = 0;   // k: la siguiente columna editada
      for (const hijo of raizF.hijos) {
        if (hijo.local !== 'table-cell' && hijo.local !== 'covered-table-cell') continue;
        const rep = repeticiones(hijo, 'table:number-columns-repeated');
        const piezas = [];
        let pos = col;
        while (k < columnas.length && columnas[k] < col + rep) {
          const c = columnas[k++];
          if (c > pos) piezas.push(copiaConRepeticion(t, hijo, 'table:number-columns-repeated', c - pos));
          piezas.push(render(hijo, t, edits.get(c)));
          pos = c + 1;
        }
        if (piezas.length) {
          if (pos < col + rep) piezas.push(copiaConRepeticion(t, hijo, 'table:number-columns-repeated', col + rep - pos));
          ed.push({ ini: hijo.ini, fin: hijo.fin, texto: piezas.join('') });
        }
        col += rep;
      }
      // columnas que están más allá de la última celda de la fila
      let cola = '';
      for (; k < columnas.length; k++) {
        const c = columnas[k];
        if (c > col) cola += celdaVacia(c - col);
        cola += render(null, null, edits.get(c));
        col = c + 1;
      }
      if (raizF.iniCierre < 0) return sinCierre(apertura) + cola + '</' + raizF.nombre + '>';
      if (cola) ed.push({ ini: raizF.iniCierre, fin: raizF.iniCierre, texto: cola });
      ed.push({ ini: 0, fin: raizF.finApertura, texto: apertura });
      return X().aplicar(t, ed);
    }
    function filaNueva(edits, render) {
      let col = 0, cuerpo = '';
      for (const c of [...edits.keys()].sort((a, b) => a - b)) {
        if (c > col) cuerpo += celdaVacia(c - col);
        cuerpo += render(null, null, edits.get(c));
        col = c + 1;
      }
      return '<table:table-row>' + cuerpo + '</table:table-row>';
    }


    function editarTabla(tablaXml, lista, render) {
      const filas = enumerarFilas(texto, tablaXml);
      const porFila = new Map();
      for (const e of lista) {
        let m = porFila.get(e.f); if (!m) porFila.set(e.f, m = new Map());
        m.set(e.c, e);
      }
      const orden = [...porFila.keys()].sort((a, b) => a - b);
      const ediciones = [];
      const porElemento = new Map();
      const extra = [];
      let k = 0;
      for (const f of orden) {
        while (k < filas.length && filas[k].f0 + filas[k].n <= f) k++;
        if (k < filas.length) { let a = porElemento.get(k); if (!a) porElemento.set(k, a = []); a.push(f); } else extra.push(f);
      }
      for (const [i, fs] of porElemento) {
        const el = filas[i], piezas = [];
        let pos = el.f0;
        for (const f of fs) {
          if (f > pos) piezas.push(copiaConRepeticion(texto, el.nodo, 'table:number-rows-repeated', f - pos));
          piezas.push(filaEditada(el.nodo, porFila.get(f), render));
          pos = f + 1;
        }
        if (pos < el.f0 + el.n) piezas.push(copiaConRepeticion(texto, el.nodo, 'table:number-rows-repeated', el.f0 + el.n - pos));
        ediciones.push({ ini: el.nodo.ini, fin: el.nodo.fin, texto: piezas.join('') });
      }
      if (extra.length) {
        // filas más allá de la última que tiene la tabla: filas vacías de relleno y las nuevas, justo después de la última fila (o de las columnas)
        const ultima = filas.length ? filas[filas.length - 1] : null;
        let libre = ultima ? ultima.f0 + ultima.n : 0;
        let texto2 = '';
        for (const f of extra) {
          if (f > libre) texto2 += '<table:table-row' + (f - libre > 1 ? ' table:number-rows-repeated="' + (f - libre) + '"' : '') + '><table:table-cell table:number-columns-repeated="' + MAX_COLS + '"/></table:table-row>';
          texto2 += filaNueva(porFila.get(f), render);
          libre = f + 1;
        }
        let pos = tablaXml.finApertura;
        if (ultima) pos = finFuera(tablaXml, ultima.nodo);
        else for (const hijo of tablaXml.hijos) if (/^(forms|shapes|table-columns?|table-header-columns|table-column-group)$/.test(hijo.local)) pos = Math.max(pos, hijo.fin);
        ediciones.push({ ini: pos, fin: pos, texto: texto2 });
      }
      return ediciones;
    }

    return { editarTabla };
  }

  function guardar(ctx) {
    const { texto, libro, hojas, tabla, contadores } = ctx;
    const f1904 = !!libro.fecha1904;
    const tieneCalcext = /xmlns:calcext=/.test(texto.slice(0, 6000));
    const avisos = { formulasNoGuardadas: 0 };
    let declararCalcext = false;   // el archivo no declara calcext y se escribe un error (ODF no tiene cómo): se declara en la raíz
    const usadosNuevos = new Set();   // xf nuevos que quedaron escritos en alguna celda de este guardado

    const nombreLibre = (prefijo, clave, existe) => { let n; do { n = prefijo + (++contadores[clave]); } while (existe(n)); return n; };
    const numeroDe = new Map();   // código de formato → { nombre, xml } de los estilos de número creados
    const estilosNumero = [];
    function nombreNumero(codigo) {
      if (!codigo || codigo === 'General') return null;
      if (numeroDe.has(codigo)) return numeroDe.get(codigo) && numeroDe.get(codigo).nombre;
      const nombre = nombreLibre('mcn', 'numero', (n) => ctx.estilos.datos.has(n));
      const r = E().estiloNumero(codigo, nombre, () => {});
      numeroDe.set(codigo, r ? { nombre, xml: r.xml } : null);
      if (r) estilosNumero.push(r.xml);
      return r ? nombre : null;
    }
    function nombreEstilo(xf) {
      const d = tabla.def(xf);
      if (d.nuevo) {
        if (!d.auto) d.auto = nombreLibre('mcs', 'estilo', (n) => ctx.estilos.celdas.has(n));
        usadosNuevos.add(xf);
        return d.auto;
      }
      return d.nombre;
    }

    // ---- una celda nueva (o reescrita) --------------------------------------------------------------------------------------------
    const textoVisible = (v, codigo) => {
      try { if (globalThis.XlsxFormato) { const t = globalThis.XlsxFormato.formatear(v, codigo || 'General', { idioma: ctx.idioma || 'en', fecha1904: f1904 }).texto; if (typeof t === 'string') return t; } } catch (e) { /* queda el número crudo */ }
      return String(v);
    };
    // atributos y párrafos de un valor ya calculado
    function valorODF(v, codigo) {
      if (typeof v === 'number' && Number.isFinite(v)) {
        const clase = E().tipoDeFormato(codigo);
        const texto = textoVisible(v, codigo);
        if (clase === 'fecha' && v >= 0) {
          const iso = E().isoDeSerie(v, f1904);
          const conHora = iso.seg !== 0 || /[hs]/i.test(String(codigo).replace(/"[^"]*"/g, ''));
          return { tipo: 'date', attrs: [['office:date-value', conHora ? iso.fecha + 'T' + iso.hora : iso.fecha]], p: E().parrafos(texto) };
        }
        if (clase === 'hora' && v >= 0) {
          const s = Math.round(v * 86400);
          return { tipo: 'time', attrs: [['office:time-value', 'PT' + Math.floor(s / 3600) + 'H' + Math.floor(s / 60) % 60 + 'M' + s % 60 + 'S']], p: E().parrafos(texto) };
        }
        return { tipo: clase === 'porcentaje' ? 'percentage' : 'float', attrs: [['office:value', String(v)]], p: E().parrafos(texto) };
      }
      if (typeof v === 'number') return { tipo: 'string', calc: 'error', attrs: [['office:string-value', '']], p: E().parrafos(Number.isNaN(v) ? '#NUM!' : '#NUM!') };
      if (typeof v === 'boolean') return { tipo: 'boolean', attrs: [['office:boolean-value', v ? 'true' : 'false']], p: E().parrafos(v ? 'TRUE' : 'FALSE') };
      if (v && typeof v === 'object' && v.error) return { tipo: 'string', calc: 'error', attrs: [['office:string-value', '']], p: E().parrafos(v.error) };
      if (typeof v === 'string') return { tipo: 'string', attrs: [], p: E().parrafos(v.length > 32767 ? v.slice(0, 32767) : v) };
      return null;
    }

    function xfPorDefecto(h, f, c) {
      const hoja = hojas[h];
      const col = hoja.columnas.find((k) => k.min <= c && c <= k.max);
      if (col && col.estilo) return col.estilo;
      const fila = hoja.filas.get(f);
      return fila && fila.formatoPersonalizado ? fila.estilo : 0;
    }

    // original: nodo de la celda que se reemplaza (o null); cubiertaPorMatriz: la celda cae dentro de una matriz que se guarda ahora
    function celdaXml(original, tOriginal, e, h, cubiertaPorMatriz) {
      const valor = e.valor;
      const conFormula = valor && typeof valor === 'object' && !valor.error && 'formula' in valor;
      const v = conFormula ? valor.valor : valor;
      const nombreOriginal = original ? original.nombre : 'table:table-cell';
      const cubierta = cubiertaPorMatriz || (original && original.local === 'covered-table-cell' && (!original.attrs['office:value-type'] || v !== null));
      const nombre = cubierta ? 'table:covered-table-cell' : (original && original.local === 'table-cell' ? nombreOriginal : 'table:table-cell');
      let etiqueta = original ? tOriginal.slice(original.ini, original.finApertura) : '<' + nombre + '/>';
      if (original && nombre !== original.nombre) etiqueta = etiqueta.replace(/^<[^\s/>]+/, '<' + nombre);
      for (const a of ATRIBUTOS_VALOR) etiqueta = X().ponerAtributo(etiqueta, a, null);
      // el estilo: el de su xf; si coincide con el que ya le da su columna o fila, no hace falta; el estilo «Default» de LibreOffice si el xf es el de siempre
      const xf = e.estilo || 0, xfDef = xfPorDefecto(h, e.f, e.c);
      const estilo = xf === xfDef ? null : (nombreEstilo(xf) || 'Default');
      etiqueta = X().ponerAtributo(etiqueta, 'table:style-name', estilo);   // sin estilo propio: el que tenía se va
      const codigo = tabla.def(xf).codigo;
      const infoValor = v === null || v === undefined ? null : valorODF(v, codigo);
      if (conFormula) {
        // A1# (anchorarray) no existe en OpenFormula: LibreOffice conserva el nombre con su prefijo _xlfn., y aOpenFormula se lo quita
        const abierta = (() => { const a = globalThis.OdsFormulas.aOpenFormula(valor.formula); return a && /_xlfn\.anchorarray\(/i.test(valor.formula) ? a.replace(/(^|[^A-Za-z0-9_.])(anchorarray\()/gi, '$1_xlfn.$2') : a; })();
        if (abierta) etiqueta = X().ponerAtributo(etiqueta, 'table:formula', 'of:=' + abierta);
        else avisos.formulasNoGuardadas++;
        const m = valor.matriz && /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(valor.matriz);
        if (abierta && m) {
          const c0 = globalThis.LibroXlsx.colANumero(m[1]), c1 = m[3] ? globalThis.LibroXlsx.colANumero(m[3]) : c0, f0 = +m[2], f1 = m[4] ? +m[4] : f0;
          etiqueta = X().ponerAtributo(X().ponerAtributo(etiqueta, 'table:number-matrix-columns-spanned', String(c1 - c0 + 1)), 'table:number-matrix-rows-spanned', String(f1 - f0 + 1));
        }
      }
      if (infoValor) {
        etiqueta = X().ponerAtributo(etiqueta, 'office:value-type', infoValor.tipo);
        for (const [k, val] of infoValor.attrs) etiqueta = X().ponerAtributo(etiqueta, k, val);
        if (conFormula && infoValor.tipo === 'string' && !infoValor.attrs.length) etiqueta = X().ponerAtributo(etiqueta, 'office:string-value', typeof v === 'string' ? v : '');
        if (tieneCalcext || infoValor.calc === 'error') {
          etiqueta = X().ponerAtributo(etiqueta, 'calcext:value-type', infoValor.calc || infoValor.tipo);
          if (!tieneCalcext) declararCalcext = true;
        }
      }
      // lo que la celda ya tenía y el editor no maneja (anotación, dibujos, rango de origen) se queda, delante de los párrafos
      let conservado = '';
      if (original && original.iniCierre >= 0) {
        const interior = X().analizar(tOriginal.slice(original.ini, original.fin), { profundidad: 2 });
        for (const hijo of interior.hijos) if (/^(office|draw|table|svg|dr3d|chart):/.test(hijo.nombre)) conservado += tOriginal.slice(original.ini + hijo.ini, original.ini + hijo.fin);
      }
      const contenido = conservado + (infoValor ? infoValor.p : '');
      etiqueta = sinCierre(etiqueta);
      return contenido ? etiqueta + contenido + '</' + nombre + '>' : etiqueta.replace(/>$/, '/>');
    }


    // ---- una tabla ---------------------------------------------------------------------------------------------------------------------------
    const motor = crearMotor(texto);
    function editarTabla(tablaXml, h, lista) {
      // las celdas de una matriz se guardan como celdas comunes, igual que LibreOffice (no como «cubiertas»)
      return motor.editarTabla(tablaXml, lista, (original, t, e) => celdaXml(original, t, e, h, false));
    }


    const raiz = X().analizar(texto, { profundidad: 5 });
    const hojaCalc = X().hijo(X().hijo(raiz, 'body'), 'spreadsheet');
    const tablas = X().hijos(hojaCalc, 'table');
    let ediciones = [];
    for (const [h, lista] of ctx.porHoja) {
      if (!tablas[h]) throw new Error('ods: no se encuentra la tabla de la hoja ' + h);
      ediciones = ediciones.concat(editarTabla(tablas[h], h, lista));
    }
    if (declararCalcext) ediciones.push({ ini: raiz.ini, fin: raiz.finApertura, texto: X().ponerAtributo(texto.slice(raiz.ini, raiz.finApertura), 'xmlns:calcext', 'urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0') });
    // los estilos nuevos van al final de office:automatic-styles
    let estilosCelda = '';
    for (const xf of usadosNuevos) {
      const d = tabla.def(xf);
      if (d.escrito) continue;
      d.escrito = true;
      estilosCelda += E().estiloCeldaXml(d.auto, d.r, nombreNumero(d.codigo));   // crea también el estilo de número si hace falta
    }
    const estilosNuevos = estilosNumero.join('') + estilosCelda;
    if (estilosNuevos) {
      const auto = X().hijo(raiz, 'automatic-styles');
      if (auto) {
        if (auto.iniCierre >= 0) ediciones.push({ ini: auto.iniCierre, fin: auto.iniCierre, texto: estilosNuevos });
        else ediciones.push({ ini: auto.ini, fin: auto.fin, texto: sinCierre(texto.slice(auto.ini, auto.finApertura)) + estilosNuevos + '</' + auto.nombre + '>' });
      } else {
        const antes = X().hijo(raiz, 'body');
        ediciones.push({ ini: antes.ini, fin: antes.ini, texto: '<office:automatic-styles>' + estilosNuevos + '</office:automatic-styles>' });
      }
    }
    return { texto: X().aplicar(texto, ediciones), avisos };
  }

  globalThis.OdsEscritorCeldas = { guardar, _util: { enumerarFilas, enumerarColumnas, finFuera, crearMotor, copiaConRepeticion, celdaVacia, repeticiones, sinCierre, MAX_COLS } };
})();

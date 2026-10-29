/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Datos de un libro OpenDocument (.ods) para Mini Cookies Hojas (F3.4b, etapa 6): nombres definidos, validación de datos y reglas de formato condicional, con la misma interfaz
// que `DatosLibro` (hojas/datos-libro.js). Cada operación toma los bytes de un .ods y devuelve los nuevos editando solo lo que cambia (content.xml, con xml-editar).
// Lo que se lee del modelo (validacionEn, listaDe, comprobarEntrada…) no depende del formato y se reusa de DatosLibro.
//
//   DatosOds.nombreValido(n) · await DatosOds.ponerNombre(bytes, { nombre, ref, hoja: null | índice }, anterior?) · await DatosOds.quitarNombre(bytes, { nombre, hoja })
//
// Requiere xml-editar, paquete-editable, ods-formulas, hojas/datos-libro (lecturas), ods-modelo (rangosDe) y ods-estructura (utilidades).
(function () {
  const X = () => globalThis.XmlEditar;
  const I = () => globalThis.EstructuraOds._interno;
  const U = () => globalThis.OdsEscritorCeldas._util;
  const Err = (m) => new XlsxEscritor.ErrorEscritor(m);

  // ---- nombres definidos: table:named-range (una dirección) y table:named-expression (una fórmula) ---------------------------------------------------
  const RE_RANGO = /^('(?:[^']|'')+'|[^!']+)!(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?$/;
  const absoluta = (c) => c.replace(/^\$?([A-Za-z]+)\$?(\d+)$/, (m, a, b) => '$' + a.toUpperCase() + '$' + b);
  const hojaOds = (h) => '$' + (/^'/.test(h) ? h : I().citar(h));
  // los contenedores de nombres: el del libro (hijo de office:spreadsheet) y el de cada hoja (hijo de su table:table), con sus nombres
  function contenedores(texto) {
    const d = I().documento(texto);
    const salida = [];
    const mirar = (nodo, hoja) => {
      for (const c of X().hijos(nodo, 'named-expressions')) {
        const sub = c.iniCierre >= 0 ? X().analizar(texto.slice(c.ini, c.fin), { profundidad: 2 }) : null;
        const elementos = sub ? sub.hijos.filter((h) => h.local === 'named-range' || h.local === 'named-expression').map((h) => ({ nombre: h.attrs['table:name'] || '', ini: c.ini + h.ini, fin: c.ini + h.fin })) : [];
        salida.push({ nodo: c, hoja, elementos });
      }
    };
    mirar(d.hojaCalc, null);
    d.tablas.forEach((t, i) => mirar(t, i));
    return { d, salida };
  }
  const valido = (n) => globalThis.DatosLibro.nombreValido(n);

  async function ponerNombre(bytes, nuevo, anterior) {
    if (!valido(nuevo.nombre)) throw Err('nombre no válido: ' + nuevo.nombre);
    const ref = String(nuevo.ref || '').replace(/^=/, '');
    if (!ref) throw Err('el nombre necesita una referencia o fórmula');
    try { globalThis.XlsFormulas._analizar(globalThis.XlsFormulas._tokenizar(ref)); } catch (e) { throw Err('la referencia del nombre no es una fórmula válida'); }
    return I().conPaquete(bytes, async (p) => {
      const texto = await p.leer('content.xml');
      const { d, salida } = contenedores(texto);
      const hoja = Number.isInteger(nuevo.hoja) ? nuevo.hoja : null;
      if (hoja !== null && !d.tablas[hoja]) throw Err('hoja inexistente');
      const hojaDe = (x) => (x === undefined ? null : x);
      const mismo = (e, c, nombre, h) => e.nombre.toLowerCase() === String(nombre).toLowerCase() && hojaDe(c.hoja) === h;
      const hAnt = anterior ? (Number.isInteger(anterior.hoja) ? anterior.hoja : null) : null;
      let viejo = null;
      for (const c of salida) for (const e of c.elementos) {
        if (anterior && mismo(e, c, anterior.nombre, hAnt)) viejo = e;
        else if (mismo(e, c, nuevo.nombre, hoja)) throw Err('ya hay un nombre «' + nuevo.nombre + '» con ese ámbito');
      }
      if (anterior && !viejo) throw Err('no está el nombre ' + anterior.nombre);
      // la dirección o la fórmula, en ODF; la base para las referencias relativas es A1 de la primera hoja (o la del nombre)
      const base = hojaOds(I().citar(I().nombreDe(d.tablas[hoja === null ? 0 : hoja]))) + '.$A$1';
      const m = RE_RANGO.exec(ref);
      let elemento;
      if (m) {
        const h = m[1];
        const dir = hojaOds(h) + '.' + absoluta(m[2]) + (m[3] ? ':.' + absoluta(m[3]) : '');
        elemento = '<table:named-range table:name="' + X().escAttr(nuevo.nombre) + '" table:base-cell-address="' + X().escAttr(base) + '" table:cell-range-address="' + X().escAttr(dir) + '"/>';
      } else {
        const ab = globalThis.OdsFormulas.aOpenFormula(ref);
        if (ab === null || ab === undefined) throw Err('la fórmula del nombre no se puede escribir en .ods');
        elemento = '<table:named-expression table:name="' + X().escAttr(nuevo.nombre) + '" table:base-cell-address="' + X().escAttr(base) + '" table:expression="' + X().escAttr(ab) + '"/>';
      }
      const eds = [];
      const mismoAmbito = viejo && salida.some((c) => hojaDe(c.hoja) === hoja && c.elementos.includes(viejo));
      if (viejo && mismoAmbito) eds.push({ ini: viejo.ini, fin: viejo.fin, texto: elemento });
      else {
        if (viejo) eds.push({ ini: viejo.ini, fin: viejo.fin, texto: '' });
        const c = salida.find((x) => hojaDe(x.hoja) === hoja);
        if (c && c.nodo.iniCierre >= 0) eds.push({ ini: c.nodo.iniCierre, fin: c.nodo.iniCierre, texto: elemento });
        else if (c) eds.push({ ini: c.nodo.ini, fin: c.nodo.fin, texto: '<table:named-expressions>' + elemento + '</table:named-expressions>' });
        else if (hoja === null) {   // el libro sin nombres: el contenedor va justo después de la última hoja
          const t = d.tablas[d.tablas.length - 1];
          eds.push({ ini: t.fin, fin: t.fin, texto: '<table:named-expressions>' + elemento + '</table:named-expressions>' });
        } else {   // la hoja sin nombres: al final de la tabla
          const t = d.tablas[hoja];
          eds.push({ ini: t.iniCierre, fin: t.iniCierre, texto: '<table:named-expressions>' + elemento + '</table:named-expressions>' });
        }
      }
      p.escribir('content.xml', X().aplicar(texto, eds));
    });
  }
  async function quitarNombre(bytes, q) {
    return I().conPaquete(bytes, async (p) => {
      const texto = await p.leer('content.xml');
      const { salida } = contenedores(texto);
      const h0 = Number.isInteger(q.hoja) ? q.hoja : null;
      for (const c of salida) {
        if ((c.hoja === undefined ? null : c.hoja) !== h0) continue;
        const e = c.elementos.find((x) => x.nombre.toLowerCase() === String(q.nombre).toLowerCase());
        if (e) { p.escribir('content.xml', X().aplicar(texto, [{ ini: e.ini, fin: e.fin, texto: '' }])); return; }
      }
      throw Err('no está el nombre ' + q.nombre);
    });
  }

  // ---- validación de datos: una definición (table:content-validation) y, en cada celda que la usa, table:content-validation-name ------------------------------------
  const MAX_CELDAS_VALIDACION = 100000;
  // la celda con ese atributo puesto o quitado (lo demás de la celda se conserva); `orig` es el elemento que ya había, o null
  function celdaConAtributo(orig, t, atributo, valor) {
    let et = orig ? t.slice(orig.ini, orig.finApertura) : '<table:table-cell/>';
    et = X().ponerAtributo(X().ponerAtributo(et, 'table:number-columns-repeated', null), atributo, valor);
    if (!orig || orig.iniCierre < 0) return et;
    return U().sinCierre(et) + t.slice(orig.finApertura, orig.fin);
  }
  const mensaje = (nombre, titulo, mostrar, texto, extra) => '<table:' + nombre + (extra || '') + ' table:title="' + X().escAttr(titulo || '') + '" table:display="' + (mostrar ? 'true' : 'false') + '">' + (texto ? '<text:p>' + X().esc(texto) + '</text:p>' : '') + '</table:' + nombre + '>';
  async function ponerValidacion(bytes, h, rangos, regla) {
    if (!rangos || !rangos.length) throw Err('sin rango');
    const area = rangos.reduce((n, r) => n + (r.f1 - r.f0 + 1) * (r.c1 - r.c0 + 1), 0);
    if (area > MAX_CELDAS_VALIDACION) throw Err('el rango es demasiado grande para ponerle validación en un .ods');
    return I().conPaquete(bytes, async (p) => {
      let texto = await p.leer('content.xml');
      const d = I().documento(texto), tabla = I().tablaDe(d, h);
      let nombreVal = null, definicion = '';
      if (regla && regla.tipo && regla.tipo !== 'none') {
        const cond = globalThis.OdsValidacion.escribir(regla);
        const usados = new Set([...texto.matchAll(/<table:content-validation\b[^>]*\stable:name="([^"]*)"/g)].map((m) => m[1]));
        let n = 0; while (usados.has('mcv' + (++n)));
        nombreVal = 'mcv' + n;
        const r0 = rangos[0], base = I().citar(I().nombreDe(tabla)) + '.' + globalThis.LibroXlsx.codificarDireccion(r0.f0, r0.c0);
        definicion = '<table:content-validation table:name="' + nombreVal + '" table:condition="' + X().escAttr(cond) + '" table:allow-empty-cell="' + (regla.permitirVacio === false ? 'false' : 'true') + '" table:base-cell-address="' + X().escAttr(base) + '"'
          + (regla.tipo === 'list' ? ' table:display-list="' + (regla.ocultarFlecha ? 'none' : 'unsorted') + '"' : '') + '>'
          + (regla.mostrarEntrada || regla.entrada || regla.tituloEntrada ? mensaje('help-message', regla.tituloEntrada, regla.mostrarEntrada, regla.entrada) : '')
          + (regla.mostrarError === false && !regla.error ? '' : mensaje('error-message', regla.tituloError, regla.mostrarError !== false, regla.error, ' table:message-type="' + (regla.estiloError || 'stop') + '"')) + '</table:content-validation>';
      }
      const lista = [];
      for (const r of rangos) for (let f = r.f0; f <= r.f1; f++) for (let c = r.c0; c <= r.c1; c++) lista.push({ f, c });
      const eds = U().crearMotor(texto).editarTabla(tabla, lista, (orig, t) => celdaConAtributo(orig, t, 'table:content-validation-name', nombreVal));
      texto = X().aplicar(texto, eds);
      // las definiciones que ya nadie usa se van, y la nueva entra en el contenedor (antes de la primera hoja si no hay)
      const d2 = I().documento(texto);
      const contenedor = X().hijo(d2.hojaCalc, 'content-validations');
      const eds2 = [];
      const referenciadas = new Set([...texto.matchAll(/\stable:content-validation-name="([^"]*)"/g)].map((m) => m[1]));
      if (contenedor && contenedor.iniCierre >= 0) {
        const sub = X().analizar(texto.slice(contenedor.ini, contenedor.fin), { profundidad: 2 });
        for (const x of sub.hijos) if (x.local === 'content-validation' && !referenciadas.has(x.attrs['table:name'])) eds2.push({ ini: contenedor.ini + x.ini, fin: contenedor.ini + x.fin, texto: '' });
        if (definicion) eds2.push({ ini: contenedor.iniCierre, fin: contenedor.iniCierre, texto: definicion });
      } else if (definicion) {
        const nuevoContenedor = '<table:content-validations>' + definicion + '</table:content-validations>';
        if (contenedor) eds2.push({ ini: contenedor.ini, fin: contenedor.fin, texto: nuevoContenedor });
        else eds2.push({ ini: d2.tablas[0].ini, fin: d2.tablas[0].ini, texto: nuevoContenedor });
      }
      p.escribir('content.xml', eds2.length ? X().aplicar(texto, eds2) : texto);
    });
  }

  // ---- formato condicional: calcext:conditional-formats (hijo de la tabla), un calcext:conditional-format por rango con una condición, una escala, una barra o un conjunto de iconos ----
  const NS_CALCEXT = 'urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0';
  const hex = (c) => String(c || '').replace(/^#/, '').toUpperCase();
  const ICONOS_N = (n) => { const k = parseInt(n, 10); return k >= 3 && k <= 5 ? k : 3; };
  const refExcel = (rangos) => rangos.map((r) => globalThis.LibroXlsx.codificarDireccion(r.f0, r.c0) + (r.f0 === r.f1 && r.c0 === r.c1 ? '' : ':' + globalThis.LibroXlsx.codificarDireccion(r.f1, r.c1))).join(' ');
  const direccionOds = (hoja, sqref) => String(sqref).split(/\s+/).filter(Boolean).map((p) => { const [a, b] = p.split(':'); return I().citar(hoja) + '.' + a.replace(/\$/g, '') + (b ? ':' + I().citar(hoja) + '.' + b.replace(/\$/g, '') : ''); }).join(' ');
  const rangosDe = (texto) => globalThis.OdsModelo.rangosDe(texto);   // (la hoja entre comillas puede llevar espacios)
  // los estilos del archivo (styles.xml y los automáticos de content.xml), para leer el formato de las reglas
  async function estilosDe(p, texto) {
    const est = globalThis.OdsEstilos.crear();
    const st = await p.leer('styles.xml');
    if (st) globalThis.OdsEstilos.leer(st, est);
    const a0 = texto.indexOf('<office:automatic-styles'), a1 = texto.indexOf('</office:automatic-styles>');
    if (a0 >= 0 && a1 > a0) globalThis.OdsEstilos.leer(texto.slice(a0, a1 + 26), est);
    return est;
  }
  // las unidades de formato condicional de una tabla, en orden de prioridad: { ini, fin, destino (target-range-address), hijo: { local, attrs, xml } }
  function unidades(texto, tabla) {
    const cont = X().hijos(tabla, 'conditional-formats')[0];
    if (!cont) return { cont: null, lista: [] };
    const lista = [];
    if (cont.iniCierre >= 0) {
      const sub = X().analizar(texto.slice(cont.ini, cont.fin), { profundidad: 3 });
      for (const cf of sub.hijos.filter((x) => x.local === 'conditional-format')) {
        for (const h of cf.hijos) lista.push({ destino: cf.attrs['calcext:target-range-address'] || '', hijo: { local: h.local, attrs: h.attrs, xml: texto.slice(cont.ini + h.ini, cont.ini + h.fin) } });
      }
    }
    return { cont, lista };
  }
  const serializar = (u) => '<calcext:conditional-format calcext:target-range-address="' + X().escAttr(u.destino) + '">' + u.hijo.xml + '</calcext:conditional-format>';
  const baseDe = (u) => { const m = /\.\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(u.hijo.attrs['calcext:base-cell-address'] || ''); return m ? { f: +m[2] - 1, c: globalThis.LibroXlsx.colANumero(m[1].toUpperCase()) } : null; };

  // la regla como la usa el diálogo (las mismas de Excel): tipo, operador, fórmulas relativas a la primera celda del primer rango, estilo…
  function reglaDe(u, est) {
    const h = u.hijo, rangos = rangosDe(u.destino), r0 = rangos[0];
    if (h.local === 'condition') {
      const r = globalThis.OdsCondicional.leerValor(h.attrs['calcext:value']);
      if (!r) return null;
      const base = baseDe(u);
      const df = base && r0 ? r0.f0 - base.f : 0, dc = base && r0 ? r0.c0 - base.c : 0;
      const mover = (x) => (df || dc ? globalThis.LibroXlsx.desplazarFormula(x, df, dc) : x);
      const nombre = h.attrs['calcext:apply-style-name'];
      const p = nombre ? globalThis.OdsEstilos.propiedadesPropias(est, est.alias.get(nombre) || nombre) : {};
      const estilo = {};
      if (/^#[0-9a-fA-F]{6}$/.test(p.bg || '')) estilo.relleno = hex(p.bg);
      if (/^#[0-9a-fA-F]{6}$/.test(p.co || '')) estilo.color = hex(p.co);
      if (p.fw === 'bold') estilo.negrita = true;
      if (p.fi === 'italic') estilo.cursiva = true;
      return Object.assign({}, r, { formulas: (r.formulas || []).map(mover), estilo });
    }
    const entradas = X().analizar(h.xml, { profundidad: 2 }).hijos;
    if (h.local === 'color-scale') return { tipo: 'colorScale', colores: entradas.map((e) => hex(e.attrs['calcext:color'])) };
    if (h.local === 'data-bar') return { tipo: 'dataBar', colores: [hex(h.attrs['calcext:positive-color'] || '#638EC6')] };
    if (h.local === 'icon-set') return { tipo: 'iconSet', iconos: h.attrs['calcext:icon-set-type'] || '3TrafficLights1' };
    return null;
  }
  const TIPO_UNIDAD = (r) => (r ? r.tipo : 'otra');

  async function reglasCondicionales(bytes, h) {
    const p = await PaqueteEditable.abrir(bytes);
    const texto = await p.leer('content.xml');
    const tabla = I().tablaDe(I().documento(texto), h);
    const est = await estilosDe(p, texto);
    return unidades(texto, tabla).lista.map((u, i) => {
      const regla = reglaDe(u, est);
      return { id: String(i), sqref: refExcel(rangosDe(u.destino)), tipo: TIPO_UNIDAD(regla), prioridad: i + 1, regla, editable: !!regla };
    });
  }

  // el XML de la unidad nueva (y su estilo automático) a partir de la regla del diálogo
  function unidadNueva(nombreHoja, sqref, regla, usados, estilosXml) {
    const rangos = String(sqref).split(/\s+/).filter(Boolean).map((x) => globalThis.LibroXlsx.decodificarRango(x)).filter(Boolean);
    if (!rangos.length) throw Err('rango no válido');
    const r0 = rangos[0], baseA1 = globalThis.LibroXlsx.codificarDireccion(r0.f0, r0.c0);
    const destino = direccionOds(nombreHoja, sqref);
    let xml;
    if (regla.tipo === 'colorScale') {
      const c = (regla.colores && regla.colores.length ? regla.colores : ['F8696B', 'FFEB84', '63BE7B']).slice(0, 3);
      const tipos = c.length === 2 ? [['0', 'minimum'], ['0', 'maximum']] : [['0', 'minimum'], ['50', 'percentile'], ['0', 'maximum']];
      xml = '<calcext:color-scale>' + c.map((col, i) => '<calcext:color-scale-entry calcext:value="' + tipos[i][0] + '" calcext:type="' + tipos[i][1] + '" calcext:color="#' + hex(col).toLowerCase() + '"/>').join('') + '</calcext:color-scale>';
    } else if (regla.tipo === 'dataBar') {
      const col = hex((regla.colores || [])[0] || '638EC6').toLowerCase();
      xml = '<calcext:data-bar calcext:min-length="10" calcext:max-length="90" calcext:negative-color="#ff0000" calcext:axis-position="none" calcext:positive-color="#' + col + '" calcext:axis-color="#000000"><calcext:formatting-entry calcext:value="0" calcext:type="minimum"/><calcext:formatting-entry calcext:value="0" calcext:type="maximum"/></calcext:data-bar>';
    } else if (regla.tipo === 'iconSet') {
      const nombre = regla.iconos || '3TrafficLights1', n = ICONOS_N(nombre);
      xml = '<calcext:icon-set calcext:icon-set-type="' + X().escAttr(nombre) + '">' + Array.from({ length: n }, (_, i) => '<calcext:formatting-entry calcext:value="' + Math.round(i * 100 / n) + '" calcext:type="percent"/>').join('') + '</calcext:icon-set>';
    } else {
      let valor = globalThis.OdsCondicional.escribirValor(regla);
      if (valor === null) throw Err('este tipo de regla no se puede escribir en .ods: ' + regla.tipo);
      valor = valor.split('#BASE#').join(baseA1);
      let nombreEstilo = '';
      const e = regla.estilo || {};
      if (e.relleno || e.color || e.negrita || e.cursiva) {
        let n = 0; while (usados.has('mccf' + (++n)));
        nombreEstilo = 'mccf' + n; usados.add(nombreEstilo);
        const col = (c) => '#' + hex(c).toLowerCase();
        estilosXml.push('<style:style style:name="' + nombreEstilo + '" style:family="table-cell" style:parent-style-name="Default">'
          + (e.relleno ? '<style:table-cell-properties fo:background-color="' + col(e.relleno) + '"/>' : '')
          + (e.color || e.negrita || e.cursiva ? '<style:text-properties' + (e.color ? ' fo:color="' + col(e.color) + '"' : '') + (e.negrita ? ' fo:font-weight="bold"' : '') + (e.cursiva ? ' fo:font-style="italic"' : '') + '/>' : '') + '</style:style>');
      }
      xml = '<calcext:condition' + (nombreEstilo ? ' calcext:apply-style-name="' + nombreEstilo + '"' : '') + ' calcext:value="' + X().escAttr(valor) + '" calcext:base-cell-address="' + X().escAttr(I().citar(nombreHoja) + '.' + baseA1) + '"/>';
    }
    const attrs = {}; for (const m of xml.slice(0, xml.indexOf('>')).matchAll(/(calcext:[\w-]+)="([^"]*)"/g)) attrs[m[1]] = X().decodificar(m[2]);
    return { destino, hijo: { local: xml.slice(9, 9 + xml.slice(9).search(/[\s>/]/)), attrs, xml } };
  }

  // reescribe el contenedor de la tabla con las unidades dadas (una unidad por elemento) y añade los estilos y el espacio de nombres que falten
  async function guardarUnidades(p, texto, h, lista, estilosXml) {
    const d = I().documento(texto), tabla = I().tablaDe(d, h);
    const { cont } = unidades(texto, tabla);
    const nuevo = lista.length ? '<calcext:conditional-formats>' + lista.map(serializar).join('') + '</calcext:conditional-formats>' : '';
    const eds = [];
    if (cont) eds.push({ ini: cont.ini, fin: cont.fin, texto: nuevo });
    else if (nuevo) eds.push({ ini: tabla.iniCierre, fin: tabla.iniCierre, texto: nuevo });
    if (nuevo && !/xmlns:calcext=/.test(texto.slice(0, 6000))) eds.push({ ini: d.raiz.ini, fin: d.raiz.finApertura, texto: X().ponerAtributo(texto.slice(d.raiz.ini, d.raiz.finApertura), 'xmlns:calcext', NS_CALCEXT) });
    let t = X().aplicar(texto, eds);
    if (estilosXml.length) t = I().agregarEstilos(t, estilosXml.join(''));
    p.escribir('content.xml', t);
  }
  async function conUnidades(bytes, h, fn) {
    return I().conPaquete(bytes, async (p) => {
      const texto = await p.leer('content.xml');
      const tabla = I().tablaDe(I().documento(texto), h);
      const { lista } = unidades(texto, tabla);
      const estilosXml = [];
      const usados = I().nombresUsados(texto);
      const nombreHoja = I().nombreDe(tabla);
      const nueva = await fn(lista, { nombreHoja, usados, estilosXml, estilos: () => estilosDe(p, texto) });
      await guardarUnidades(p, texto, h, nueva || lista, estilosXml);
    });
  }
  const indiceDe = (lista, id) => { const i = parseInt(id, 10); if (!(i >= 0 && i < lista.length)) throw Err('no está la regla ' + id); return i; };
  const agregarRegla = (bytes, h, sqref, regla) => conUnidades(bytes, h, (lista, c) => [unidadNueva(c.nombreHoja, sqref, regla, c.usados, c.estilosXml), ...lista]);
  const cambiarRegla = (bytes, h, id, sqref, regla) => conUnidades(bytes, h, (lista, c) => { const i = indiceDe(lista, id); const nueva = lista.slice(); nueva[i] = unidadNueva(c.nombreHoja, sqref, regla, c.usados, c.estilosXml); return nueva; });
  const quitarRegla = (bytes, h, id) => conUnidades(bytes, h, (lista) => { const i = indiceDe(lista, id); return lista.filter((_, k) => k !== i); });
  const moverRegla = (bytes, h, id, delta) => conUnidades(bytes, h, (lista) => {
    const i = indiceDe(lista, id), j = i + (delta < 0 ? -1 : 1);
    if (j < 0 || j >= lista.length) return lista;
    const nueva = lista.slice(); [nueva[i], nueva[j]] = [nueva[j], nueva[i]]; return nueva;
  });

  globalThis.DatosOds = {
    nombreValido: valido, ponerNombre, quitarNombre, ponerValidacion, reglasCondicionales, agregarRegla, cambiarRegla, quitarRegla, moverRegla,
    // lo que lee del modelo, igual para todos los formatos
    validacionEn: (...a) => globalThis.DatosLibro.validacionEn(...a), listaDe: (...a) => globalThis.DatosLibro.listaDe(...a),
    comprobarEntrada: (...a) => globalThis.DatosLibro.comprobarEntrada(...a), comprobarDespues: (...a) => globalThis.DatosLibro.comprobarDespues(...a),
  };
})();

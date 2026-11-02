/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Página e impresión de un libro OpenDocument (.ods) para Mini Cookies Hojas (F3.4b, etapa 6): papel, orientación, márgenes, escala o ajuste a páginas, área de impresión, filas que se
// repiten, encabezado y pie, cuadrícula y centrado, con la misma interfaz que `PaginaLibro` (hojas/pagina-libro.js). En ODF la página es un estilo de página (`style:page-layout`, en styles.xml)
// al que apunta una página maestra (`style:master-page`, con el encabezado y el pie) que la tabla elige por su estilo (`style:master-page-name`); el área de impresión es el atributo
// `table:print-ranges` de la hoja y las filas que se repiten son las que envuelve `table:table-header-rows`. Una página maestra la comparten las hojas que la nombran: configurar una hoja
// que comparte la suya le crea una propia (`mcpN`) y deja las demás como estaban. El encabezado y el pie se llevan al código de Excel (`&L &C &R`, `&P` página, `&N` total, `&A` hoja…).
//
//   OdsPagina.leerLibro(stylesXml) → { maestras: Map, layouts: Map }                                  (lo que el modelo necesita de styles.xml)
//   OdsPagina.configuracion(maestra, layout) → la configuración de PaginaLibro.leer (lo que no está, con los valores de Excel)
//   OdsPagina.imprimirDe(config) → los campos de `hoja.imprimir` del modelo
//   await PaginaOds.leer(bytes, hoja) · await PaginaOds.configurar(bytes, hoja, config)
//
// Límites: «ajustar a N páginas en total» (`style:scale-to-pages`) se lee como 100 %; los saltos de página manuales, las filas o columnas de título repartidas en grupos y el
// encabezado de la primera página distinto se conservan pero no se editan.
(function () {
  const X = () => globalThis.XmlEditar;
  const I = () => globalThis.EstructuraOds._interno;
  const U = () => globalThis.OdsEscritorCeldas._util;
  const Err = (m) => new XlsxEscritor.ErrorEscritor(m);
  // tamaños de papel de Excel (paperSize) en pulgadas, de menor a mayor lado
  const TAMANOS = { 1: [8.5, 11], 5: [8.5, 14], 9: [8.2677, 11.6929], 8: [11.6929, 16.5354], 11: [5.8268, 8.2677], 7: [7.25, 10.5] };
  const PAPELES = { 1: 'Carta', 5: 'Oficio (Legal)', 9: 'A4', 8: 'A3', 11: 'A5', 7: 'Ejecutivo' };
  const DEFECTO = { orientacion: 'portrait', papel: 1, margenes: { izq: 0.7, der: 0.7, sup: 0.75, inf: 0.75, enc: 0.3, pie: 0.3 }, ajuste: null, escala: 100, area: null, titulosFilas: null, encabezado: '', pie: '', cuadricula: false, centrarH: false, centrarV: false };
  const ALTO_TEXTO = 0.1528;   // una línea de 11 pt, en pulgadas
  // una dirección de ODF («Hoja.A1:Hoja.C3», con la hoja entre comillas si lleva espacios); un atributo puede traer varias separadas por espacios
  const RE_DIRECCION = /\$?(?:'(?:[^']|'')+'|[^\s.']+)\.\$?[A-Za-z]{1,3}\$?\d+(?::\$?(?:'(?:[^']|'')+'|[^\s.']+)?\.\$?[A-Za-z]{1,3}\$?\d+)?/g;
  const REGIONES = { left: 'izq', center: 'cen', right: 'der' };
  const CAMPOS = { 'sheet-name': '&A', 'page-number': '&P', 'page-count': '&N', date: '&D', time: '&T', 'file-name': '&F', title: '&F' };

  // «8.5in», «21cm», «0.5in» → pulgadas (undefined si no hay medida)
  function pulgadas(v) {
    const m = /^(-?[\d.]+)\s*(in|cm|mm|pt|pc)?$/.exec(String(v === undefined ? '' : v).trim());
    if (!m) return undefined;
    const n = parseFloat(m[1]);
    return n * { in: 1, cm: 1 / 2.54, mm: 1 / 25.4, pt: 1 / 72, pc: 1 / 6 }[m[2] || 'in'];
  }
  const redondear = (n) => Math.round(n * 1000) / 1000;
  const enPulgadas = (n) => Math.round(n * 10000) / 10000 + 'in';
  const papelDe = (a, b) => {
    const lo = Math.min(a, b), hi = Math.max(a, b);
    for (const [id, [x, y]] of Object.entries(TAMANOS)) if (Math.abs(x - lo) < 0.05 && Math.abs(y - hi) < 0.05) return +id;
    return null;
  };

  // ---- leer styles.xml: estilos de página y páginas maestras --------------------------------------------------------------------------------------------
  function leerLibro(texto) {
    const layouts = new Map(), maestras = new Map();
    let L = null, seccion = null, M = null, parte = null, region = 'cen', parrafo = null, enCampo = 0;
    if (!texto) return { maestras, layouts };
    globalThis.XmlLigero.recorrer(texto, {
      abrir(n, a) {
        if (n === 'page-layout' && a.name !== undefined) { L = { nombre: a.name, p: {}, cab: null, pie: null }; layouts.set(a.name, L); return; }
        if (L) {
          if (n === 'page-layout-properties') L.p = a;
          else if (n === 'header-style') seccion = 'cab';
          else if (n === 'footer-style') seccion = 'pie';
          else if (n === 'header-footer-properties' && seccion) L[seccion] = a;
          return;
        }
        if (n === 'master-page' && a.name !== undefined) { M = { nombre: a.name, layout: a['page-layout-name'] || null, cab: null, pie: null }; maestras.set(a.name, M); parte = null; return; }
        if (!M) return;
        if (n === 'header' || n === 'footer') { parte = n === 'header' ? 'cab' : 'pie'; M[parte] = a.display === 'false' ? null : { izq: '', cen: '', der: '' }; region = 'cen'; return; }
        if (!parte || !M[parte]) return;
        if (n === 'region-left' || n === 'region-center' || n === 'region-right') { region = REGIONES[n.slice(7)]; return; }
        if (n === 'p') { parrafo = ''; return; }
        if (parrafo === null) return;
        if (CAMPOS[n]) { parrafo += CAMPOS[n]; enCampo++; }
        else if (n === 's') parrafo += ' '.repeat(Math.min(parseInt(a.c, 10) || 1, 200));
        else if (n === 'tab') parrafo += '\t';
        else if (n === 'line-break') parrafo += '\n';
      },
      cerrar(n) {
        if (n === 'page-layout') { L = null; seccion = null; return; }
        if (n === 'header-style' || n === 'footer-style') { seccion = null; return; }
        if (n === 'master-page') { M = null; parte = null; return; }
        if (n === 'header' || n === 'footer') { parte = null; parrafo = null; return; }
        if (!M || !parte || !M[parte]) return;
        if (CAMPOS[n]) { enCampo = Math.max(0, enCampo - 1); return; }
        if (n === 'p' && parrafo !== null) { M[parte][region] += (M[parte][region] ? '\n' : '') + parrafo; parrafo = null; }
      },
      texto(t) { if (parrafo !== null && !enCampo) parrafo += t.replace(/&/g, '&&'); },
    });
    return { maestras, layouts };
  }
  const codigoHF = (g) => (g ? (g.izq ? '&L' + g.izq : '') + (g.cen ? '&C' + g.cen : '') + (g.der ? '&R' + g.der : '') : '');
  const hayTexto = (g) => !!g && !!(g.izq || g.cen || g.der);

  // la configuración de página (la del diálogo y la de PaginaLibro.leer) de una página maestra y su estilo de página
  function configuracion(maestra, layout) {
    const c = JSON.parse(JSON.stringify(DEFECTO)), p = (layout && layout.p) || {};
    const w = pulgadas(p['page-width']), h = pulgadas(p['page-height']);
    c.orientacion = p['print-orientation'] === 'landscape' ? 'landscape' : 'portrait';
    if (w !== undefined && h !== undefined) { const id = papelDe(w, h); c.papel = id || 1; c._papelPropio = id ? null : { ancho: p['page-width'], alto: p['page-height'] }; }
    const cab = !!(maestra && maestra.cab), pie = !!(maestra && maestra.pie);
    const m = (k, d) => { const v = pulgadas(p['margin-' + k]); return v === undefined ? d : redondear(v); };
    const altoCab = layout && layout.cab ? pulgadas(layout.cab['min-height'] || layout.cab.height) : undefined, altoPie = layout && layout.pie ? pulgadas(layout.pie['min-height'] || layout.pie.height) : undefined;
    // en ODF el margen de arriba/abajo es hasta el encabezado/pie y éste tiene su altura; en Excel es el margen del texto y el encabezado cuelga de más arriba
    const arr = m('top', undefined), aba = m('bottom', undefined);
    c.margenes = {
      izq: m('left', DEFECTO.margenes.izq), der: m('right', DEFECTO.margenes.der),
      sup: arr === undefined ? DEFECTO.margenes.sup : cab && altoCab !== undefined ? redondear(arr + altoCab) : arr, enc: arr === undefined || !cab ? DEFECTO.margenes.enc : arr,
      inf: aba === undefined ? DEFECTO.margenes.inf : pie && altoPie !== undefined ? redondear(aba + altoPie) : aba, pie: aba === undefined || !pie ? DEFECTO.margenes.pie : aba,
    };
    const pct = /^(\d+(?:\.\d+)?)\s*%$/.exec(p['scale-to'] || '');
    if (pct) c.escala = Math.max(10, Math.min(400, Math.round(+pct[1])));
    else if (p['scale-to-X'] !== undefined || p['scale-to-Y'] !== undefined) c.ajuste = { ancho: parseInt(p['scale-to-X'], 10) || 0, alto: parseInt(p['scale-to-Y'], 10) || 0 };
    const imp = String(p.print || '').split(/\s+/);
    c.cuadricula = imp.includes('grid');
    c.centrarH = p['table-centering'] === 'horizontal' || p['table-centering'] === 'both';
    c.centrarV = p['table-centering'] === 'vertical' || p['table-centering'] === 'both';
    // (como en PaginaLibro.leer: el diálogo ve un « & » suelto y los códigos &P &N… tal cual)
    c.encabezado = maestra && maestra.cab ? maestra.cab.cen.replace(/&&/g, '&') : '';
    c.pie = maestra && maestra.pie ? maestra.pie.cen.replace(/&&/g, '&') : '';
    c._primeraPagina = /^\d+$/.test(p['first-page-number'] || '') ? +p['first-page-number'] : null;
    return c;
  }
  // los campos del modelo (`hoja.imprimir`) que lee el motor de PDF
  function imprimirDe(c, maestra) {
    const aStr = (v) => (v === undefined || v === null ? null : v);
    return {
      margenes: { izq: c.margenes.izq, der: c.margenes.der, arr: c.margenes.sup, aba: c.margenes.inf, encabezado: c.margenes.enc, pie: c.margenes.pie },
      pagina: { papel: c.papel, anchoPapel: aStr(c._papelPropio && c._papelPropio.ancho), altoPapel: aStr(c._papelPropio && c._papelPropio.alto), orientacion: c.orientacion, escala: c.escala,
        ajustarAncho: c.ajuste ? c.ajuste.ancho : 1, ajustarAlto: c.ajuste ? c.ajuste.alto : 1, orden: 'downThenOver', primeraPagina: c._primeraPagina, usarPrimeraPagina: c._primeraPagina !== null, blancoYNegro: false, borrador: false, comentarios: 'none', errores: 'displayed' },
      opciones: { cuadricula: c.cuadricula, encabezados: false, centrarH: c.centrarH, centrarV: c.centrarV },
      encabezadoPie: maestra && (hayTexto(maestra.cab) || hayTexto(maestra.pie)) ? { distintoPrimera: false, distintoParImpar: false, escalar: true, alinear: true, impar: { encabezado: codigoHF(maestra.cab), pie: codigoHF(maestra.pie) }, par: {}, primera: {} } : null,
      ajustarAPagina: !!c.ajuste,
    };
  }

  // ---- escribir ------------------------------------------------------------------------------------------------------------------------------------------
  const conAtributos = (etiqueta, atributos) => { let t = etiqueta; for (const [k, v] of Object.entries(atributos)) t = X().ponerAtributo(t, k, v === null ? null : String(v)); return t; };
  // el código de encabezado de Excel (&C… con &P &N &D &T &F &A) → los párrafos de ODF
  function parrafosHF(codigo) {
    const lineas = String(codigo || '').split('\n');
    return lineas.map((linea) => {
      let xml = '', i = 0;
      const texto = (t) => { if (t) xml += X().esc(t); };
      while (i < linea.length) {
        const ch = linea[i];
        if (ch !== '&') { let j = i; while (j < linea.length && linea[j] !== '&') j++; texto(linea.slice(i, j)); i = j; continue; }
        const sig = linea[i + 1];
        if (sig === '&') { texto('&'); i += 2; }
        else if (sig === 'P') { xml += '<text:page-number>1</text:page-number>'; i += 2; }
        else if (sig === 'N') { xml += '<text:page-count>99</text:page-count>'; i += 2; }
        else if (sig === 'A') { xml += '<text:sheet-name>?</text:sheet-name>'; i += 2; }
        else if (sig === 'D') { xml += '<text:date>01/01/2000</text:date>'; i += 2; }
        else if (sig === 'T') { xml += '<text:time>00:00:00</text:time>'; i += 2; }
        else if (sig === 'F') { xml += '<text:file-name text:display="name">?</text:file-name>'; i += 2; }
        else if (sig === '"') { const j = linea.indexOf('"', i + 2); i = j < 0 ? linea.length : j + 1; }   // &"Fuente,Estilo"
        else if (/[0-9]/.test(sig || '')) { let j = i + 1; while (j < linea.length && /[0-9]/.test(linea[j])) j++; i = j; }   // &12: tamaño de letra
        else if (/[BIUSXYEK]/.test(sig || '')) i += 2;   // &B &I &U &S &X &Y &E &K: formato que aquí no se lleva
        else { texto('&'); i += 1; }   // un « & » suelto es un « & »
      }
      return '<text:p>' + xml + '</text:p>';
    }).join('');
  }
  // el contenido de style:header / style:footer: tres regiones si hay izquierda o derecha, y si no un solo párrafo
  function contenidoHF(g) {
    if (!hayTexto(g)) return '';
    if (!g.izq && !g.der) return parrafosHF(g.cen);
    return ['izq', 'cen', 'der'].map((k, i) => (g[k] ? '<style:region-' + ['left', 'center', 'right'][i] + '>' + parrafosHF(g[k]) + '</style:region-' + ['left', 'center', 'right'][i] + '>' : '')).join('');
  }
  // los códigos de Excel de una sección (&L…&C…&R…) → { izq, cen, der }
  function dividirCodigo(codigo) {
    const g = { izq: '', cen: '', der: '' };
    const t = String(codigo || '');
    if (!/&[LCR]/.test(t)) { g.cen = t; return g; }
    const re = /&([LCR])((?:[^&]|&[^LCR])*)/g;
    let m;
    while ((m = re.exec(t))) g[{ L: 'izq', C: 'cen', R: 'der' }[m[1]]] += m[2];
    return g;
  }

  const nombreLibre = (usados, prefijo) => { let n = 0; while (usados.has(prefijo + (++n))); usados.add(prefijo + n); return prefijo + n; };
  const maestraDeTabla = (estilosTablas, tabla, primera) => (tabla.attrs['table:style-name'] && estilosTablas.get(tabla.attrs['table:style-name'])) || primera;
  // estilo de tabla → página maestra, leído de content.xml
  function estilosDeTablas(contenido) {
    const m = new Map();
    for (const e of contenido.matchAll(/<style:style\b[^>]*>/g)) {
      const nombre = /\sstyle:name="([^"]*)"/.exec(e[0]), familia = /\sstyle:family="table"/.test(e[0]), maestra = /\sstyle:master-page-name="([^"]*)"/.exec(e[0]);
      if (nombre && familia && maestra) m.set(X().decodificar(nombre[1]), X().decodificar(maestra[1]));
    }
    return m;
  }

  async function datosHoja(p, h) {
    const contenido = await p.leer('content.xml'), styles = (await p.leer('styles.xml')) || '';
    const d = I().documento(contenido), tabla = I().tablaDe(d, h);
    const { maestras, layouts } = leerLibro(styles);
    const primera = maestras.size ? maestras.keys().next().value : null;
    const tablas = estilosDeTablas(contenido);
    const nombreM = maestraDeTabla(tablas, tabla, primera);
    const maestra = nombreM ? maestras.get(nombreM) || null : null;
    return { contenido, styles, d, tabla, maestras, layouts, tablas, primera, maestra, layout: maestra && maestra.layout ? layouts.get(maestra.layout) || null : null };
  }

  // el rango de «A1:C40» del atributo table:print-ranges («Uno.A1:Uno.C40 Uno.E1:Uno.F9») → el primero que es de esta hoja
  function areaDe(atributo) {
    const rangos = [];
    for (const parte of String(atributo || '').match(RE_DIRECCION) || []) {
      const celdas = [...parte.matchAll(/\.\$?([A-Za-z]{1,3})\$?(\d+)/g)].map((m) => m[1] + m[2]);
      if (celdas.length) rangos.push(celdas.length > 1 ? celdas[0] + ':' + celdas[celdas.length - 1] : celdas[0]);
    }
    return rangos;
  }
  // las filas que envuelve table:table-header-rows (desde 0) → { f0, f1 } | null
  function titulosDe(texto, tabla) {
    const grupo = X().hijos(tabla, 'table-header-rows')[0];
    if (!grupo) return null;
    const filas = U().enumerarFilas(texto, tabla).filter((f) => f.nodo.ini >= grupo.ini && f.nodo.fin <= grupo.fin);
    return filas.length ? { f0: filas[0].f0, f1: filas[filas.length - 1].f0 + filas[filas.length - 1].n - 1 } : null;
  }

  async function leer(bytes, h) {
    const p = await PaqueteEditable.abrir(bytes);
    const x = await datosHoja(p, h);
    const c = configuracion(x.maestra, x.layout);
    delete c._papelPropio; delete c._primeraPagina;
    const areas = areaDe(x.tabla.attrs['table:print-ranges']);
    c.area = areas.length ? areas[0] : null;
    const t = titulosDe(x.contenido, x.tabla);
    c.titulosFilas = t ? (t.f0 + 1) + ':' + (t.f1 + 1) : null;
    return c;
  }

  // ---- configurar -----------------------------------------------------------------------------------------------------------------------------------------
  // los estilos de página y las páginas maestras de styles.xml, con su posición
  function elementosDe(styles) {
    const raiz = X().analizar(styles, { profundidad: 6 });
    const auto = X().hijo(raiz, 'automatic-styles'), maestros = X().hijo(raiz, 'master-styles');
    const lista = (cont, local, atributo) => (cont ? X().hijos(cont, local) : []).map((n) => ({ nombre: n.attrs[atributo], nodo: n }));
    return { raiz, auto, maestros, layouts: lista(auto, 'page-layout', 'style:name'), paginas: lista(maestros, 'master-page', 'style:name') };
  }

  async function configurar(bytes, h, c) {
    return I().conPaquete(bytes, async (p) => {
      const x = await datosHoja(p, h);
      let { contenido, styles } = x;
      if (!styles) throw Err('el archivo no tiene styles.xml');
      const actual = configuracion(x.maestra, x.layout);
      const todas = elementosDe(styles);
      if (!todas.auto || !todas.maestros) throw Err('el archivo no tiene estilos de página');
      // ¿la página maestra es sólo de esta hoja? Si la comparten, esta hoja se queda con una copia
      const nombreM = x.maestra ? x.maestra.nombre : null;
      let usan = 0;
      const primera = x.primera;
      for (const t of x.d.tablas) if (maestraDeTabla(x.tablas, t, primera) === nombreM) usan++;
      const clonar = !x.maestra || usan > 1;
      let layoutOrigen = x.layout ? todas.layouts.find((l) => l.nombre === x.layout.nombre) : null;
      const maestraOrigen = nombreM ? todas.paginas.find((m) => m.nombre === nombreM) : null;
      const usadosLayout = maestraOrigen && layoutOrigen ? [...x.maestras.values()].filter((m) => m.layout === x.layout.nombre).length : 0;
      const clonarLayout = clonar || usadosLayout > 1 || !layoutOrigen;

      // lo que cambia
      const papelId = c.papel !== undefined ? +c.papel : actual.papel;
      if (c.papel !== undefined && !(Number.isInteger(papelId) && papelId >= 1 && papelId <= 118)) throw Err('tamaño de papel no válido');
      const orient = c.orientacion !== undefined ? (c.orientacion === 'landscape' ? 'landscape' : 'portrait') : actual.orientacion;
      const margenes = Object.assign({}, actual.margenes, c.margenes || {});
      for (const k in margenes) if (!(margenes[k] >= 0 && margenes[k] <= 49)) throw Err('margen no válido');
      const cab = Object.assign({ izq: '', cen: '', der: '' }, x.maestra && x.maestra.cab ? x.maestra.cab : {});
      const pie = Object.assign({ izq: '', cen: '', der: '' }, x.maestra && x.maestra.pie ? x.maestra.pie : {});
      if (c.encabezado !== undefined) cab.cen = String(c.encabezado || '');
      if (c.pie !== undefined) pie.cen = String(c.pie || '');
      const conCab = hayTexto(cab), conPie = hayTexto(pie);

      const propiedades = {};
      let ancho, alto;
      if (c.papel !== undefined || c.orientacion !== undefined) {
        if (TAMANOS[papelId]) [ancho, alto] = TAMANOS[papelId];
        else { const w = pulgadas(x.layout && x.layout.p['page-width']), hh = pulgadas(x.layout && x.layout.p['page-height']); [ancho, alto] = w !== undefined && hh !== undefined ? [Math.min(w, hh), Math.max(w, hh)] : TAMANOS[1]; }
        if (orient === 'landscape') [ancho, alto] = [alto, ancho];
        propiedades['fo:page-width'] = enPulgadas(ancho); propiedades['fo:page-height'] = enPulgadas(alto);
      }
      propiedades['style:print-orientation'] = orient;
      propiedades['fo:margin-left'] = enPulgadas(margenes.izq); propiedades['fo:margin-right'] = enPulgadas(margenes.der);
      propiedades['fo:margin-top'] = enPulgadas(conCab ? margenes.enc : margenes.sup);
      propiedades['fo:margin-bottom'] = enPulgadas(conPie ? margenes.pie : margenes.inf);
      if (c.ajuste) { propiedades['style:scale-to'] = null; propiedades['style:scale-to-pages'] = null; propiedades['style:scale-to-X'] = Math.max(0, Math.trunc(+c.ajuste.ancho || 0)); propiedades['style:scale-to-Y'] = Math.max(0, Math.trunc(+c.ajuste.alto || 0)); }
      else if (c.escala !== undefined) {
        const e = Math.trunc(+c.escala);
        if (!(e >= 10 && e <= 400)) throw Err('la escala va de 10 a 400 %');
        propiedades['style:scale-to-X'] = null; propiedades['style:scale-to-Y'] = null; propiedades['style:scale-to-pages'] = null; propiedades['style:scale-to'] = e === 100 ? null : e + '%';
      }
      if (c.cuadricula !== undefined) {
        const fichas = new Set(String((x.layout && x.layout.p.print) || 'charts drawings objects zero-values').split(/\s+/).filter(Boolean));
        if (c.cuadricula) fichas.add('grid'); else fichas.delete('grid');
        propiedades['style:print'] = [...fichas].join(' ') || null;
      }
      if (c.centrarH !== undefined || c.centrarV !== undefined) {
        const hh = c.centrarH !== undefined ? c.centrarH : actual.centrarH, vv = c.centrarV !== undefined ? c.centrarV : actual.centrarV;
        propiedades['style:table-centering'] = hh && vv ? 'both' : hh ? 'horizontal' : vv ? 'vertical' : null;
      }

      // el estilo de página: se reescribe a partir del que hay (o de uno vacío)
      const usadosNombres = new Set([...todas.layouts.map((l) => l.nombre), ...todas.paginas.map((m) => m.nombre), ...I().nombresUsados(contenido)]);
      const nombreLayout = clonarLayout ? nombreLibre(usadosNombres, 'mcpl') : x.layout.nombre;
      const alturaCab = Math.max(0, margenes.sup - margenes.enc), alturaPie = Math.max(0, margenes.inf - margenes.pie);   // alto total: el texto más el espacio hasta los datos
      const origen = layoutOrigen ? layoutOrigen.nodo : null;
      const textoLayout = () => {
        const abre = origen ? U().sinCierre(styles.slice(origen.ini, origen.finApertura)) : '<style:page-layout>';
        const hijos = origen && origen.iniCierre >= 0 ? origen.hijos : [];
        let xml = X().ponerAtributo(abre, 'style:name', nombreLayout);
        const trozo = (n) => styles.slice(n.ini, n.fin);
        const props = hijos.find((n) => n.local === 'page-layout-properties');
        const propsXml = props ? conAtributos(U().sinCierre(styles.slice(props.ini, props.finApertura)), propiedades) : conAtributos('<style:page-layout-properties>', propiedades);
        const propsCompleto = props && props.iniCierre >= 0 ? propsXml + styles.slice(props.finApertura, props.fin) : propsXml.replace(/>$/, '/>');
        const hf = (local, nombreSec, presente, altura, lado) => {
          const viejo = hijos.find((n) => n.local === local);
          const hfp = viejo && viejo.iniCierre >= 0 ? viejo.hijos.find((n) => n.local === 'header-footer-properties') : null;
          const base = hfp ? U().sinCierre(styles.slice(hfp.ini, hfp.finApertura)) : '<style:header-footer-properties>';
          const at = { 'fo:min-height': enPulgadas(presente ? Math.max(altura, 0.05) : (hfp && hfp.attrs['fo:min-height'] ? pulgadas(hfp.attrs['fo:min-height']) : 0.2953)) };
          if (presente) at['fo:margin-' + lado] = enPulgadas(Math.max(0, Math.max(altura, 0.05) - ALTO_TEXTO));   // LibreOffice lo exporta a Excel como «margen + altura del texto + este espacio»
          let t = conAtributos(base, at);
          t = hfp && hfp.iniCierre >= 0 ? t + styles.slice(hfp.finApertura, hfp.fin) : t.replace(/>$/, '/>');
          return '<style:' + nombreSec + '>' + t + '</style:' + nombreSec + '>';
        };
        const otros = hijos.filter((n) => !['page-layout-properties', 'header-style', 'footer-style'].includes(n.local)).map(trozo).join('');
        xml += propsCompleto + hf('header-style', 'header-style', conCab, alturaCab, 'bottom') + hf('footer-style', 'footer-style', conPie, alturaPie, 'top') + otros + '</style:page-layout>';
        return xml;
      };
      // la página maestra: el encabezado y el pie se reemplazan, lo demás (primera página, izquierda) se conserva
      const nombreMaestra = clonar ? nombreLibre(usadosNombres, 'mcp') : nombreM;
      const textoMaestra = () => {
        const o = maestraOrigen ? maestraOrigen.nodo : null;
        let abre = o ? U().sinCierre(styles.slice(o.ini, o.finApertura)) : '<style:master-page>';
        abre = conAtributos(abre, { 'style:name': nombreMaestra, 'style:page-layout-name': nombreLayout, 'style:display-name': clonar ? null : (o && o.attrs['style:display-name']) || null });
        const hijos = o && o.iniCierre >= 0 ? o.hijos : [];
        const seccion = (local, g) => (hayTexto(g) ? '<style:' + local + '>' + contenidoHF(g) + '</style:' + local + '>' : '<style:' + local + ' style:display="false"/>');
        const otros = hijos.filter((n) => n.local !== 'header' && n.local !== 'footer').map((n) => styles.slice(n.ini, n.fin));
        const lados = otros.filter((e) => /^<style:(header|footer)-/.test(e)), resto = otros.filter((e) => !/^<style:(header|footer)-/.test(e));
        const orden = [seccion('header', cab), ...lados.filter((e) => /^<style:header-/.test(e)), seccion('footer', pie), ...lados.filter((e) => /^<style:footer-/.test(e)), ...resto];
        return abre + orden.join('') + '</style:master-page>';
      };

      // styles.xml: el estilo y la página maestra (en su lugar o al final de su sección)
      const ediciones = [];
      const nuevoLayout = textoLayout(), nuevaMaestra = textoMaestra();
      if (!clonarLayout) ediciones.push({ ini: origen.ini, fin: origen.fin, texto: nuevoLayout });
      else if (todas.auto.iniCierre >= 0) ediciones.push({ ini: todas.auto.iniCierre, fin: todas.auto.iniCierre, texto: nuevoLayout });
      else ediciones.push({ ini: todas.auto.ini, fin: todas.auto.fin, texto: U().sinCierre(styles.slice(todas.auto.ini, todas.auto.finApertura)) + nuevoLayout + '</' + todas.auto.nombre + '>' });
      if (!clonar) ediciones.push({ ini: maestraOrigen.nodo.ini, fin: maestraOrigen.nodo.fin, texto: nuevaMaestra });
      else if (todas.maestros.iniCierre >= 0) ediciones.push({ ini: todas.maestros.iniCierre, fin: todas.maestros.iniCierre, texto: nuevaMaestra });
      else ediciones.push({ ini: todas.maestros.ini, fin: todas.maestros.fin, texto: U().sinCierre(styles.slice(todas.maestros.ini, todas.maestros.finApertura)) + nuevaMaestra + '</' + todas.maestros.nombre + '>' });
      styles = X().aplicar(styles, ediciones);

      // content.xml: el estilo de la tabla que apunta a la página (si se clonó), el área de impresión y las filas que se repiten
      let tabla = x.tabla;
      if (clonar) {
        const nombreEstilo = nombreLibre(usadosNombres, 'mct');
        const viejo = tabla.attrs['table:style-name'];
        const m = viejo ? new RegExp('<style:style\\b[^>]*\\sstyle:name="' + viejo.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"[^>]*?(?:/>|>[\\s\\S]*?</style:style>)').exec(contenido) : null;   // (sin hijos va cerrado en sí mismo: no tomar el estilo de al lado)
        let estilo = m ? m[0] : '<style:style style:family="table"><style:table-properties table:display="true" style:writing-mode="lr-tb"/></style:style>';
        estilo = estilo.replace(/^<style:style\b[^>]*>/, (e) => conAtributos(e, { 'style:name': nombreEstilo, 'style:master-page-name': nombreMaestra, 'style:family': 'table' }));
        contenido = I().agregarEstilos(contenido, estilo);
        tabla = I().tablaDe(I().documento(contenido), h);
        contenido = X().aplicar(contenido, [{ ini: tabla.ini, fin: tabla.finApertura, texto: X().ponerAtributo(contenido.slice(tabla.ini, tabla.finApertura), 'table:style-name', nombreEstilo) }]);
        tabla = I().tablaDe(I().documento(contenido), h);
      }
      if (c.area !== undefined) {
        const r = c.area ? globalThis.LibroXlsx.decodificarRango(String(c.area).replace(/\$/g, '').toUpperCase()) : null;
        if (c.area && !r) throw Err('el área de impresión no es un rango');
        const nombreHoja = I().nombreDe(tabla), cit = I().citar(nombreHoja);
        const dir = r ? cit + '.' + globalThis.LibroXlsx.codificarDireccion(r.f0, r.c0) + ':' + cit + '.' + globalThis.LibroXlsx.codificarDireccion(r.f1, r.c1) : null;
        contenido = X().aplicar(contenido, [{ ini: tabla.ini, fin: tabla.finApertura, texto: X().ponerAtributo(contenido.slice(tabla.ini, tabla.finApertura), 'table:print-ranges', dir) }]);
        tabla = I().tablaDe(I().documento(contenido), h);
      }
      if (c.titulosFilas !== undefined) {
        const m = c.titulosFilas ? /^\$?(\d+)(?::\$?(\d+))?$/.exec(String(c.titulosFilas).trim()) : null;
        if (c.titulosFilas && !m) throw Err('las filas que se repiten van como 1:2');
        contenido = ponerTitulos(contenido, tabla, m ? { f0: +m[1] - 1, f1: +(m[2] || m[1]) - 1 } : null);
      }
      p.escribir('styles.xml', styles);
      p.escribir('content.xml', contenido);
    });
  }

  // las filas f0..f1 dentro de table:table-header-rows (y las que ya estaban, fuera); null las saca todas
  function ponerTitulos(texto, tabla, r) {
    if (/<table:table-row-group\b/.test(texto.slice(tabla.ini, tabla.fin))) throw Err('esta hoja tiene filas agrupadas: todavía no se pueden cambiar las filas que se repiten en un .ods');
    const grupo = X().hijos(tabla, 'table-header-rows')[0];
    if (grupo) {   // se desenvuelve: las filas quedan donde estaban
      const dentro = grupo.iniCierre >= 0 ? texto.slice(grupo.finApertura, grupo.iniCierre) : '';
      texto = X().aplicar(texto, [{ ini: grupo.ini, fin: grupo.fin, texto: dentro }]);
      tabla = I().tablaDe(I().documento(texto), I().documento(texto).tablas.findIndex((t) => t.ini === tabla.ini));
    }
    if (!r) return texto;
    if (r.f1 < r.f0 || r.f1 > 1048575) throw Err('las filas que se repiten no son válidas');
    const filas = U().enumerarFilas(texto, tabla);
    const ult = filas.length ? filas[filas.length - 1].f0 + filas[filas.length - 1].n : 0;
    if (r.f1 >= ult) throw Err('las filas que se repiten tienen que estar dentro de los datos de la hoja');
    const eds = [];
    for (const el of filas) {
      const a = el.f0, b = el.f0 + el.n - 1;
      if (b < r.f0 || a > r.f1) continue;
      const lo = Math.max(a, r.f0), hi = Math.min(b, r.f1);
      if (lo === a && hi === b) continue;   // el elemento entero está dentro
      const piezas = [];
      if (lo > a) piezas.push(U().copiaConRepeticion(texto, el.nodo, 'table:number-rows-repeated', lo - a));
      piezas.push(U().copiaConRepeticion(texto, el.nodo, 'table:number-rows-repeated', hi - lo + 1));
      if (hi < b) piezas.push(U().copiaConRepeticion(texto, el.nodo, 'table:number-rows-repeated', b - hi));
      eds.push({ ini: el.nodo.ini, fin: el.nodo.fin, texto: piezas.join('') });
    }
    texto = eds.length ? X().aplicar(texto, eds) : texto;
    // ahora cada elemento cae entero dentro o fuera: se envuelven los de dentro
    const doc = I().documento(texto), idx = doc.tablas.findIndex((t) => t.ini === tabla.ini);
    const nuevaTabla = doc.tablas[idx];
    const dentro = U().enumerarFilas(texto, nuevaTabla).filter((el) => el.f0 >= r.f0 && el.f0 + el.n - 1 <= r.f1);
    if (!dentro.length) throw Err('no se encontraron las filas que se repiten');
    const ini = dentro[0].nodo.ini, fin = dentro[dentro.length - 1].nodo.fin;
    return X().aplicar(texto, [{ ini, fin, texto: '<table:table-header-rows>' + texto.slice(ini, fin) + '</table:table-header-rows>' }]);
  }

  globalThis.OdsPagina = { leerLibro, configuracion, imprimirDe, areaDe, titulosDe, pulgadas, PAPELES, DEFECTO };
  globalThis.PaginaOds = { leer, configurar, PAPELES, DEFECTO };
})();

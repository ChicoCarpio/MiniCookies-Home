/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Gráficos de un libro OpenDocument (.ods) para Mini Cookies Hojas (F3.4b, etapa 6), con la interfaz de `GraficosLibro` (hojas/graficos-libro.js).
// En ODF un gráfico es un objeto incrustado: en la celda donde está anclado (o en `table:shapes`) hay un `draw:frame` con un `draw:object` que apunta a una carpeta
// «Object N/» del paquete, y su `content.xml` guarda el gráfico (`chart:chart`: clase, título, leyenda, ejes, series con los rangos de celdas) más una tabla local con
// los valores. Para dibujarlo se usa el motor que ya dibuja los gráficos de Excel: de cada gráfico se arma un dibujo virtual (drawingN.xml con su ancla y chartM.xml en
// DrawingML) que sólo vive en la vista del paquete que ve el motor (`libro.zip`); lo que se guarda en el archivo es el paquete tal cual, con las ediciones de ODF.
//
//   await OdsGraficos.preparar(libro, hojas, paquete, lector?)    arma la vista del paquete y las anclas de cada hoja (lo llama OdsBackend.abrir)
//   GraficosOds.listar(bytes, hoja) · cambiarSeries · moverAncla · insertar · refrescarCaches · lectorDe   (misma interfaz que GraficosLibro)
//
// Se dibujan: columnas y barras (agrupadas, apiladas y al 100 %), líneas, áreas, circular, anillo y dispersión, y la combinación de columnas con líneas. Los demás
// (cotizaciones, burbujas, radar, 3D…) se conservan en el archivo pero no se dibujan.
// Requiere xml-ligero, xml-editar, graficos-libro (el generador de caché) y ods-estructura (utilidades de direcciones).
(function () {
  const EMU = { in: 914400, cm: 360000, mm: 36000, pt: 12700, pc: 152400 };
  const emu = (v) => { const m = /^(-?[\d.]+)\s*(in|cm|mm|pt|pc)?$/.exec(String(v === undefined || v === null ? '' : v).trim()); return m ? Math.round(parseFloat(m[1]) * EMU[m[2] || 'in']) : 0; };
  const PALETA = ['004586', 'FF420E', 'FFD320', '579D1C', '7E0021', '83CAFF', '314004', 'AECF00', '4B1F6F', 'FF950E'];
  const CLASES = { 'chart:bar': 'bar', 'chart:line': 'line', 'chart:area': 'area', 'chart:circle': 'circle', 'chart:ring': 'ring', 'chart:scatter': 'scatter' };
  const X = () => globalThis.XmlEditar;
  const G = () => globalThis.GraficosLibro._interno;
  const I = () => globalThis.EstructuraOds._interno;
  const U = () => globalThis.OdsEscritorCeldas._util;
  const hex = (c) => (/^#[0-9a-fA-F]{6}$/.test(c || '') ? c.slice(1).toUpperCase() : null);

  // ---- leer la definición de un gráfico (Object N/content.xml) ------------------------------------------------------------------------------------
  function leerDefinicion(texto) {
    const estilos = new Map();   // nombre → { fondo, linea, apilado, porcentaje, vertical, simbolos, grosor }
    let estilo = null, def = null, prof = 0, profChart = -1, serie = null, ejeActual = null, titulo = null, enTitulo = false, enTabla = 0;
    globalThis.XmlLigero.recorrer(texto, {
      abrir(n, a) {
        prof++;
        if (enTabla) { enTabla++; return; }
        if (n === 'table' && a.name === 'local-table') { enTabla = 1; return; }
        if (n === 'style' && a.family === 'chart' && a.name) { estilo = { nombre: a.name }; estilos.set(a.name, estilo); return; }
        if (estilo) {
          if (n === 'chart-properties') {
            if (a.stacked !== undefined) estilo.apilado = a.stacked === 'true';
            if (a.percentage !== undefined) estilo.porcentaje = a.percentage === 'true';
            if (a.vertical !== undefined) estilo.vertical = a.vertical === 'true';
            if (a['symbol-type'] !== undefined) estilo.simbolos = a['symbol-type'] !== 'none';
            if (a['three-dimensional'] === 'true') estilo.tresD = true;
            if (a.lines !== undefined) estilo.lineas = a.lines === 'true';
            if (a.visible === 'false') estilo.oculto = true;
          } else if (n === 'graphic-properties') {
            if (a['fill-color'] !== undefined && a.fill !== 'none') estilo.fondo = hex(a['fill-color']);
            if (a['stroke-color'] !== undefined) estilo.linea = hex(a['stroke-color']);
            if (a['stroke-width'] !== undefined) estilo.grosor = emu(a['stroke-width']);
            if (a.stroke === 'none') estilo.sinLinea = true;
          }
          return;
        }
        if (n === 'chart' && !def && a.class !== undefined) { profChart = prof; def = { clase: CLASES[a.class] || null, claseOds: a.class || '', titulo: '', leyenda: null, plano: null, categorias: null, series: [], ancho: emu(a.width), alto: emu(a.height) }; return; }
        if (!def) return;
        if (n === 'title' && prof === profChart + 1) { enTitulo = true; titulo = ''; }
        else if (n === 'legend') def.leyenda = a['legend-position'] || 'end';
        else if (n === 'plot-area') def.plano = a['style-name'] || null;
        else if (n === 'axis') ejeActual = { dimension: a.dimension, nombre: a.name, estilo: a['style-name'] };
        else if (n === 'categories' && ejeActual && ejeActual.dimension === 'x' && !def.categorias) def.categorias = a['cell-range-address'] || null;
        else if (n === 'series') { serie = { clase: CLASES[a.class] || def.clase, valores: a['values-cell-range-address'] || null, etiqueta: a['label-cell-address'] || null, estilo: a['style-name'] || null, dominio: null, puntos: [], eje: a['attached-axis'] || 'primary-y' }; def.series.push(serie); }
        else if (n === 'domain' && serie && !serie.dominio) serie.dominio = a['cell-range-address'] || null;
        else if (n === 'data-point' && serie) { const k = Math.max(1, Math.min(parseInt(a.repeated, 10) || 1, 1000)); for (let i = 0; i < k && serie.puntos.length < 5000; i++) serie.puntos.push(a['style-name'] || null); }
      },
      cerrar(n) {
        const nivel = prof--;
        if (enTabla) { enTabla--; return; }
        if (n === 'style') estilo = null;
        if (!def) return;
        if (n === 'title') { enTitulo = false; if (nivel === profChart + 1) def.titulo = titulo || ''; }
        else if (n === 'axis') ejeActual = null;
        else if (n === 'series') serie = null;
      },
      texto(t) { if (enTitulo && titulo !== null && !enTabla) titulo += t; },
    });
    if (!def) return null;
    def.estilos = estilos;
    return def;
  }
  // «Hoja.B2:Hoja.B4» → «Hoja!$B$2:$B$4» (la referencia de Excel que lee el motor); null si no es una dirección
  function aExcel(dir) {
    if (!dir) return null;
    const xl = I().direccionAExcel(String(dir).trim());
    if (!xl) return null;
    const i = xl.lastIndexOf('!');
    const rango = xl.slice(i + 1).replace(/\$/g, '').split(':').map((c) => c.replace(/^([A-Za-z]+)(\d+)$/, '$$$1$$$2'));
    return xl.slice(0, i + 1) + (rango.length === 2 && rango[0] === rango[1] ? rango[0] : rango.join(':'));   // «B1:B1» es la celda B1
  }

  // ---- armar el XML de DrawingML ------------------------------------------------------------------------------------------------------------------
  function xmlSerie(i, def, s, leer, clase) {
    const gi = G(), esc = (t) => X().esc(t);
    const est = (s.estilo && def.estilos.get(s.estilo)) || {};
    const color = est.fondo || est.linea || PALETA[i % PALETA.length];
    const refTexto = (dir) => { const f = aExcel(dir); if (!f) return ''; const d = leer(f); return `<c:strRef><c:f>${esc(f)}</c:f>${gi.cacheStr('c:', d ? d.textos : [])}</c:strRef>`; };
    const nombre = s.etiqueta ? `<c:tx>${refTexto(s.etiqueta)}</c:tx>` : '';
    const fv = aExcel(s.valores), dv = fv ? leer(fv) : null;
    const val = fv ? `<c:numRef><c:f>${esc(fv)}</c:f>${gi.cacheNum('c:', dv ? dv.valores : [], dv ? dv.formato : 'General')}</c:numRef>` : '';
    const cat = def.categorias ? `<c:cat>${refTexto(def.categorias)}</c:cat>` : '';
    const relleno = `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;
    if (clase === 'bar' || clase === 'area') return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${nombre}<c:spPr>${relleno}</c:spPr>${clase === 'bar' ? '<c:invertIfNegative val="0"/>' : ''}${cat}<c:val>${val}</c:val></c:ser>`;
    if (clase === 'line') {
      const w = Math.max(9525, Math.min(est.grosor || 28575, 76200));
      const marca = def.simbolos(s) ? `<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>${relleno}</c:spPr></c:marker>` : '<c:marker><c:symbol val="none"/></c:marker>';
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${nombre}<c:spPr><a:ln w="${w}" cap="rnd"><a:solidFill><a:srgbClr val="${est.linea || color}"/></a:solidFill><a:round/></a:ln></c:spPr>${marca}${cat}<c:val>${val}</c:val><c:smooth val="0"/></c:ser>`;
    }
    if (clase === 'circle' || clase === 'ring') {
      const n = dv ? dv.valores.flat().length : 0;
      const puntos = Array.from({ length: n }, (_, k) => { const pe = (s.puntos[k] && def.estilos.get(s.puntos[k])) || {}; return `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${pe.fondo || PALETA[k % PALETA.length]}"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`; }).join('');
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${nombre}${puntos}${cat}<c:val>${val}</c:val></c:ser>`;
    }
    // dispersión: x del dominio, y de los valores
    const fx = aExcel(s.dominio), dx = fx ? leer(fx) : null;
    const xv = fx ? `<c:xVal><c:numRef><c:f>${esc(fx)}</c:f>${gi.cacheNum('c:', dx ? dx.valores : [], dx ? dx.formato : 'General')}</c:numRef></c:xVal>` : (def.categorias ? `<c:xVal>${refTexto(def.categorias)}</c:xVal>` : '');
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${nombre}<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr>${relleno}</c:spPr></c:marker>${xv}<c:yVal>${val}</c:yVal><c:smooth val="0"/></c:ser>`;
  }
  const EJE_CAT = (pos) => `<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
  const EJE_VAL = (id, cruce, pos, rejilla) => `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/>${rejilla ? '<c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>' : ''}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cruce}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  const POS_LEYENDA = { bottom: 'b', end: 'r', top: 't', start: 'l', right: 'r', left: 'l' };

  // El gráfico en DrawingML con los valores de ahora (leer(f) → { valores, textos, formato } | null); null si la clase no se dibuja
  function aOoxml(def, leer) {
    if (!def || !def.series.length) return null;
    const plano = (def.plano && def.estilos.get(def.plano)) || {};
    const clases = [...new Set(def.series.map((s) => s.clase))].filter(Boolean);
    if (!clases.length || clases.some((c) => !['bar', 'line', 'area', 'circle', 'ring', 'scatter'].includes(c))) return null;
    if (plano.tresD) return null;
    const horizontal = !!plano.vertical && clases.includes('bar');
    const grupo = plano.porcentaje ? 'percentStacked' : plano.apilado ? 'stacked' : 'clustered';
    const lineal = plano.porcentaje ? 'percentStacked' : plano.apilado ? 'stacked' : 'standard';
    def.simbolos = (s) => { const e = (s.estilo && def.estilos.get(s.estilo)) || {}; return e.simbolos === undefined ? false : e.simbolos; };
    const sers = (clase) => def.series.map((s, i) => (s.clase === clase ? xmlSerie(i, def, s, leer, clase) : '')).join('');
    let cuerpo = '';
    const ejes = '<c:axId val="500000001"/><c:axId val="500000002"/>';
    for (const clase of clases) {
      if (clase === 'bar') cuerpo += `<c:barChart><c:barDir val="${horizontal ? 'bar' : 'col'}"/><c:grouping val="${grupo}"/><c:varyColors val="0"/>${sers('bar')}<c:gapWidth val="150"/>${grupo === 'clustered' ? '' : '<c:overlap val="100"/>'}${ejes}</c:barChart>`;
      else if (clase === 'line') cuerpo += `<c:lineChart><c:grouping val="${lineal}"/><c:varyColors val="0"/>${sers('line')}<c:marker val="1"/>${ejes}</c:lineChart>`;
      else if (clase === 'area') cuerpo += `<c:areaChart><c:grouping val="${lineal}"/><c:varyColors val="0"/>${sers('area')}${ejes}</c:areaChart>`;
      else if (clase === 'circle') cuerpo += `<c:pieChart><c:varyColors val="1"/>${sers('circle')}<c:firstSliceAng val="0"/></c:pieChart>`;
      else if (clase === 'ring') cuerpo += `<c:doughnutChart><c:varyColors val="1"/>${sers('ring')}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>`;
      else cuerpo += `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers('scatter')}${ejes}</c:scatterChart>`;
    }
    const sinEjes = clases.every((c) => c === 'circle' || c === 'ring');
    if (!sinEjes) {
      if (clases.includes('scatter')) cuerpo += EJE_VAL('500000001', '500000002', 'b', false) + EJE_VAL('500000002', '500000001', 'l', true);
      else cuerpo += EJE_CAT(horizontal ? 'l' : 'b') + EJE_VAL('500000002', '500000001', horizontal ? 'b' : 'l', true);
    }
    const titulo = def.titulo ? `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${X().esc(def.titulo)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` : '';
    const leyenda = def.leyenda ? `<c:legend><c:legendPos val="${POS_LEYENDA[def.leyenda] || 'r'}"/><c:overlay val="0"/></c:legend>` : '';
    const NS = G();
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${NS.NS_C}" xmlns:a="${NS.NS_A}" xmlns:r="${NS.NS_R}"><c:roundedCorners val="0"/><c:chart>${titulo}<c:autoTitleDeleted val="${def.titulo ? 0 : 1}"/><c:plotArea><c:layout/>${cuerpo}</c:plotArea>${leyenda}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>`
      + '<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:chartSpace>';
  }

  // ---- los valores de las celdas -------------------------------------------------------------------------------------------------------------------
  // Un lector sobre el modelo (los valores guardados): sirve mientras el editor no da los calculados. Misma forma que GraficosLibro.lectorDe.
  function lectorModelo(libro, hojas) {
    return (f) => {
      const m = /^('(?:[^']|'')+'|[^!']+)!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(String(f));
      if (!m) return null;
      const nombre = m[1][0] === "'" ? m[1].slice(1, -1).replace(/''/g, "'") : m[1];
      const h = libro.hojas.findIndex((x) => x.nombre.toLowerCase() === nombre.toLowerCase());
      if (h < 0 || !hojas[h]) return null;
      const L = globalThis.LibroXlsx;
      const r0 = +m[3] - 1, c0 = L.colANumero(m[2]), r1 = m[5] ? +m[5] - 1 : r0, c1 = m[4] ? L.colANumero(m[4]) : c0;
      if ((Math.abs(r1 - r0) + 1) * (Math.abs(c1 - c0) + 1) > 100000) return null;
      const C = hojas[h].celdas, valores = [], textos = [];
      for (let f2 = Math.min(r0, r1); f2 <= Math.max(r0, r1); f2++) {
        const fv = [], ft = [];
        for (let c2 = Math.min(c0, c1); c2 <= Math.max(c0, c1); c2++) {
          const i = C.indiceDe(f2, c2), v = i < 0 || C.tipo[i] === L.TIPO.vacio || C.tipo[i] === L.TIPO.error ? null : C.valor[i];
          fv.push(v === undefined ? null : v); ft.push(v === null || v === undefined ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v));
        }
        valores.push(fv); textos.push(ft);
      }
      return { valores, textos, formato: 'General' };
    };
  }

  // ---- la vista del paquete que ve el motor ------------------------------------------------------------------------------------------------------
  const REL_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
  function xmlDibujo(marcos, primeraRel) {
    const NS = G();
    const ancla = (m, i) => {
      const marco = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${i + 2}" name="${X().escAttr(m.nombre || 'Gráfico ' + (i + 1))}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${NS.NS_C}"><c:chart xmlns:c="${NS.NS_C}" xmlns:r="${NS.NS_R}" r:id="rId${primeraRel + i}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>`;
      if (m.absoluta) return `<xdr:absoluteAnchor><xdr:pos x="${m.x}" y="${m.y}"/><xdr:ext cx="${m.w}" cy="${m.h}"/>${marco}</xdr:absoluteAnchor>`;
      return `<xdr:oneCellAnchor><xdr:from><xdr:col>${m.col}</xdr:col><xdr:colOff>${m.x}</xdr:colOff><xdr:row>${m.fila}</xdr:row><xdr:rowOff>${m.y}</xdr:rowOff></xdr:from><xdr:ext cx="${m.w}" cy="${m.h}"/>${marco}</xdr:oneCellAnchor>`;
    };
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${NS.NS_XDR}" xmlns:a="${NS.NS_A}">${marcos.map(ancla).join('')}</xdr:wsDr>`;
  }
  function crearVista(zip, virtuales) {
    const vista = Object.create(zip);
    const entrada = (ruta, texto) => ({ name: ruta, dir: false, async: async (tipo) => (tipo === 'string' || tipo === 'text' ? texto : tipo === 'uint8array' ? new TextEncoder().encode(texto) : tipo === 'base64' ? btoa(unescape(encodeURIComponent(texto))) : texto) });
    vista.file = function (ruta, ...resto) {
      if (resto.length || typeof ruta !== 'string') return zip.file(ruta, ...resto);
      const f = virtuales.get(ruta);
      return f ? entrada(ruta, f()) : zip.file(ruta);
    };
    vista.forEach = function (fn) { zip.forEach(fn); for (const r of virtuales.keys()) fn(r, vista.file(r)); };
    return vista;
  }

  // Lee los gráficos de cada hoja del modelo (las definiciones y las anclas de los que se saben dibujar). `paquete`: el paquete original (de ahí salen los Object N)
  async function reunir(libro, hojas, paquete) {
    const estado = { graficos: [], lector: lectorModelo(libro, hojas), ultimo: new Map(), virtuales: new Map(), dibujables: hojas.map(() => []) };
    let n = 0;
    for (let h = 0; h < hojas.length; h++) {
      const lista = [];
      hojas[h].anclas = [];
      for (const m of hojas[h]._marcos || []) {
        const ruta = String(m.href || '').replace(/^\.\//, '').replace(/\/$/, '');
        const t = await paquete.leer(ruta + '/content.xml');
        const def = t ? leerDefinicion(t) : null;
        const g = { hoja: h, objeto: ruta, def, marco: m, nombre: m.nombre, virtual: null };
        estado.graficos.push(g);
        if (def && aOoxml(def, () => null) !== null) lista.push(g);   // sólo se dibujan los que se saben dibujar; el resto queda en el archivo
      }
      estado.dibujables[h] = lista;
      lista.forEach((g, i) => {
        const m = g.marco, absoluta = !m.enCelda;
        const a = { nombre: m.nombre, absoluta, col: m.col || 0, fila: m.fila || 0, x: emu(m.x), y: emu(m.y), w: emu(m.w) || g.def.ancho, h: emu(m.h) || g.def.alto };
        g.virtual = { ruta: 'xl/charts/chart' + (++n) + '.xml', ancla: a, indice: i, rutaDibujo: 'xl/drawings/drawing' + (h + 1) + '.xml' };
        hojas[h].anclas.push(absoluta ? { tipo: 'abs', de: null, a: null, ext: { w: a.w, h: a.h }, pos: { x: a.x, y: a.y }, imprimible: true, oculto: false }
          : { tipo: 'uno', de: { col: a.col, fila: a.fila, dx: a.x, dy: a.y }, a: null, ext: { w: a.w, h: a.h }, pos: null, imprimible: true, oculto: false });
      });
      if (!lista.length) continue;
      const rutaD = 'xl/drawings/drawing' + (h + 1) + '.xml', rutaR = 'xl/drawings/_rels/drawing' + (h + 1) + '.xml.rels';
      estado.virtuales.set(rutaD, () => xmlDibujo(lista.map((g) => g.virtual.ancla), 1));
      estado.virtuales.set(rutaR, () => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + lista.map((g, i) => `<Relationship Id="rId${i + 1}" Type="${REL_CHART}" Target="../charts/${g.virtual.ruta.split('/').pop()}"/>`).join('') + '</Relationships>');
      for (const g of lista) estado.virtuales.set(g.virtual.ruta, () => aOoxml(g.def, estado.lector) || '');
      hojas[h].partes.dibujo = rutaD;
      hojas[h].objetos = undefined;
    }
    return estado;
  }
  // Lee los gráficos de cada hoja, arma el dibujo virtual y se lo pone a `libro.zip` (el paquete que se guarda es el original: `paquete`).
  async function preparar(libro, hojas, paquete) {
    const estado = await reunir(libro, hojas, paquete);
    libro._graficos = estado;
    if (estado.virtuales.size) { libro.zip = crearVista(libro.zip, estado.virtuales); libro.zip._graficosEstado = estado; }
    return estado;
  }
  // Pone al día lo que dibuja el motor con los valores de ahora: devuelve cuántos gráficos cambiaron (el motor vuelve a leerlos cuando se lo piden)
  async function refrescarCaches(zip, lector) {
    const est = zip && zip._graficosEstado;
    if (!est) return 0;
    if (lector) est.lector = lector;
    let cambiados = 0;
    for (const g of est.graficos) {
      if (!g.virtual) continue;
      const xml = aOoxml(g.def, est.lector) || '';
      if (est.ultimo.get(g.virtual.ruta) !== xml) { est.ultimo.set(g.virtual.ruta, xml); cambiados++; }
    }
    return cambiados;
  }

  // ---- la interfaz de GraficosLibro sobre los bytes de un .ods ----------------------------------------------------------------------------------
  const Err = (m) => new XlsxEscritor.ErrorEscritor(m);
  const TIPO_XLSX = { bar: 'barChart', line: 'lineChart', area: 'areaChart', circle: 'pieChart', ring: 'doughnutChart', scatter: 'scatterChart' };
  async function abrirTodo(bytes) {
    const libro = await OdsModelo.abrir(bytes);
    const hojas = [];
    for (let i = 0; i < libro.hojas.length; i++) hojas[i] = await libro.cargarHoja(i);
    const paquete = PaqueteEditable.desdeZip(libro.zip);
    const estado = await reunir(libro, hojas, paquete);
    return { libro, hojas, paquete, estado };
  }
  const excelDe = (d) => aExcel(d);
  async function listar(bytes, h) {
    const { estado } = await abrirTodo(bytes);
    return (estado.dibujables[h] || []).map((g) => ({
      indice: g.virtual.indice, rutaDibujo: g.virtual.rutaDibujo, rutaGrafico: g.virtual.ruta, tipo: TIPO_XLSX[g.def.clase] || 'otro',
      series: g.def.series.map((s) => ({ nombre: excelDe(s.etiqueta), categorias: excelDe(s.dominio || g.def.categorias), valores: excelDe(s.valores) })),
    }));
  }
  const alOds = (f) => (f === null || f === undefined || f === '' ? null : I().direccionAOds(String(f).replace(/\$/g, '')));
  const pulgadas = (e) => Math.round(e / EMU.in * 1000000) / 1000000 + 'in';

  // la tabla local del gráfico (los valores que guarda el objeto): una columna por serie, con las categorías al principio
  function tablaLocal(series, categorias, leer) {
    const gi = G(), esc = (t) => X().esc(t);
    const lectura = (dir) => { const f = aExcel(dir); return f ? leer(f) : null; };
    const cols = series.map((s) => ({ nombre: lectura(s.etiqueta), valores: lectura(s.valores) }));
    const cat = lectura(categorias);
    const n = Math.max(0, ...cols.map((c) => (c.valores ? c.valores.valores.flat().length : 0)), cat ? cat.textos.flat().length : 0);
    const celdaTexto = (t) => '<table:table-cell office:value-type="string"><text:p>' + esc(t) + '</text:p></table:table-cell>';
    const celdaNum = (v) => (typeof v === 'number' && Number.isFinite(v) ? `<table:table-cell office:value-type="float" office:value="${v}"><text:p>${v}</text:p></table:table-cell>` : '<table:table-cell/>');
    void gi;
    let filas = '';
    for (let i = 0; i < n; i++) {
      filas += '<table:table-row>' + celdaTexto(cat && cat.textos.flat()[i] !== undefined ? cat.textos.flat()[i] : String(i + 1)) + cols.map((c) => celdaNum(c.valores ? c.valores.valores.flat()[i] : null)).join('') + '</table:table-row>';
    }
    return '<table:table table:name="local-table"><table:table-header-columns><table:table-column/></table:table-header-columns><table:table-columns><table:table-column table:number-columns-repeated="' + Math.max(1, series.length) + '"/></table:table-columns>'
      + '<table:table-header-rows><table:table-row><table:table-cell><text:p/></table:table-cell>' + cols.map((c) => celdaTexto(c.nombre ? c.nombre.textos.flat()[0] || '' : '')).join('') + '</table:table-row></table:table-header-rows>'
      + '<table:table-rows>' + filas + '</table:table-rows></table:table>';
  }
  const rangosDe = (def) => [...new Set([def.categorias, ...def.series.flatMap((s) => [s.etiqueta, s.valores, s.dominio])].filter(Boolean))];

  async function cambiarSeries(bytes, h, k, series, leer) {
    return I().conPaquete(bytes, async (p) => {
      const libro = await OdsModelo.abrir(bytes);
      const hojas = []; for (let i = 0; i < libro.hojas.length; i++) hojas[i] = await libro.cargarHoja(i);
      const estado = await reunir(libro, hojas, PaqueteEditable.desdeZip(libro.zip));
      const g = (estado.dibujables[h] || [])[k];
      if (!g) throw Err('no hay un gráfico ahí');
      let t = await p.leer(g.objeto + '/content.xml');
      const sers = [...t.matchAll(/<chart:series\b[^>]*>/g)];
      if (series.length !== sers.length) throw Err('cambiar la cantidad de series se hace en LibreOffice');
      const eds = [];
      const nuevas = g.def.series.map((s, i) => ({ etiqueta: series[i].nombre === undefined ? s.etiqueta : alOds(series[i].nombre), valores: series[i].valores === undefined ? s.valores : alOds(series[i].valores), dominio: s.dominio }));
      const dadas = series.map((x) => x.categorias).filter((c) => c !== undefined);
      const primera = dadas.find((c) => c);
      const cat = dadas.length ? (primera ? alOds(primera) : null) : g.def.categorias;
      const esDispersion = g.def.clase === 'scatter';
      sers.forEach((m, i) => {
        let tag = m[0];
        tag = X().ponerAtributo(tag, 'chart:values-cell-range-address', nuevas[i].valores);
        tag = X().ponerAtributo(tag, 'chart:label-cell-address', nuevas[i].etiqueta);
        eds.push({ ini: m.index, fin: m.index + m[0].length, texto: tag });
      });
      // las categorías (o el eje x de una dispersión) son del eje, no de cada serie
      const mc = /<chart:categories\b[^>]*\/?>/.exec(t);
      if (mc && !esDispersion) eds.push({ ini: mc.index, fin: mc.index + mc[0].length, texto: cat ? X().ponerAtributo(mc[0], 'table:cell-range-address', cat) : '' });
      else if (!mc && cat && !esDispersion) {
        const ejeX = /<chart:axis\b[^>]*chart:dimension="x"[^>]*>/.exec(t);
        if (ejeX && !/\/>$/.test(ejeX[0])) eds.push({ ini: ejeX.index + ejeX[0].length, fin: ejeX.index + ejeX[0].length, texto: '<chart:categories table:cell-range-address="' + X().escAttr(cat) + '"/>' });
      }
      const ml = /<table:table table:name="local-table"[\s\S]*?<\/table:table>/.exec(t);
      const defNueva = { categorias: esDispersion ? g.def.categorias : cat, series: nuevas };
      if (ml) eds.push({ ini: ml.index, fin: ml.index + ml[0].length, texto: tablaLocal(nuevas, defNueva.categorias, leer) });
      t = X().aplicar(t, eds);
      p.escribir(g.objeto + '/content.xml', t);
      // el marco avisa a LibreOffice de qué rangos depende (así lo vuelve a dibujar cuando cambian)
      const contenido = await p.leer('content.xml');
      const re = new RegExp('<draw:object\\b[^>]*xlink:href="\\./' + g.objeto.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"[^>]*>');
      const mo = re.exec(contenido);
      if (mo) p.escribir('content.xml', X().aplicar(contenido, [{ ini: mo.index, fin: mo.index + mo[0].length, texto: X().ponerAtributo(mo[0], 'draw:notify-on-update-of-ranges', rangosDe(defNueva.categorias === undefined ? g.def : { categorias: defNueva.categorias, series: nuevas }).join(' ')) }]));
    });
  }

  // ---- mover y cambiar el tamaño ----------------------------------------------------------------------------------------------------------------
  const RE_MARCO = /<draw:frame\b[^>]*>[\s\S]*?<\/draw:frame>/g;
  // el draw:frame que apunta a «Object N» en content.xml: { ini, fin, texto }
  function marcoDe(contenido, objeto) {
    for (const m of contenido.matchAll(RE_MARCO)) if (m[0].includes('xlink:href="./' + objeto + '"')) return { ini: m.index, fin: m.index + m[0].length, texto: m[0] };
    return null;
  }
  // el marco con otro tamaño y otra posición (relativa a su celda) y sin las direcciones de celda de las formas ancladas
  function marcoConPosicion(texto, x, y, w, h) {
    const fin = texto.indexOf('>') + 1;
    let tag = texto.slice(0, fin);
    for (const a of ['table:end-cell-address', 'table:end-x', 'table:end-y']) tag = X().ponerAtributo(tag, a, null);
    if (w > 0) tag = X().ponerAtributo(tag, 'svg:width', pulgadas(w));
    if (h > 0) tag = X().ponerAtributo(tag, 'svg:height', pulgadas(h));
    tag = X().ponerAtributo(X().ponerAtributo(tag, 'svg:x', pulgadas(Math.max(0, x))), 'svg:y', pulgadas(Math.max(0, y)));
    return tag + texto.slice(fin);
  }
  // la celda con el marco puesto al final de su contenido, o sin él (orig: el elemento de la celda, o null si todavía no existe)
  function celdaConMarco(orig, t, marco) {
    let et = orig ? t.slice(orig.ini, orig.finApertura) : '<table:table-cell/>';
    et = X().ponerAtributo(et, 'table:number-columns-repeated', null);
    if (!orig) return U().sinCierre(et) + marco + '</table:table-cell>';
    if (orig.iniCierre < 0) return U().sinCierre(et) + marco + '</' + orig.nombre + '>';
    return et + t.slice(orig.finApertura, orig.iniCierre) + marco + t.slice(orig.iniCierre, orig.fin);
  }
  function celdaSinMarco(orig, t, objeto) {
    const et = X().ponerAtributo(t.slice(orig.ini, orig.finApertura), 'table:number-columns-repeated', null);
    if (orig.iniCierre < 0) return et;
    const cuerpo = t.slice(orig.finApertura, orig.iniCierre);
    const m = marcoDe(cuerpo, objeto);
    return et + (m ? cuerpo.slice(0, m.ini) + cuerpo.slice(m.fin) : cuerpo) + t.slice(orig.iniCierre, orig.fin);
  }

  async function moverAncla(bytes, h, k, pos) {
    const { estado } = await abrirTodo(bytes);
    const g = (estado.dibujables[h] || [])[k];
    if (!g) throw Err('no hay un gráfico ahí');
    return I().conPaquete(bytes, async (p) => {
      let texto = await p.leer('content.xml');
      const marco = marcoDe(texto, g.objeto);
      if (!marco) throw Err('no se encontró el gráfico en la hoja');
      const w = pos.ext ? pos.ext.w : emu(g.marco.w), alto = pos.ext ? pos.ext.h : emu(g.marco.h);
      const de = pos.de || { col: g.marco.col, fila: g.marco.fila, dx: emu(g.marco.x), dy: emu(g.marco.y) };
      const nuevo = marcoConPosicion(marco.texto, de.dx, de.dy, w, alto);
      const m = g.marco, mismaCelda = m.enCelda && m.fila === Math.trunc(de.fila) && m.col === Math.trunc(de.col);
      if (mismaCelda) { p.escribir('content.xml', X().aplicar(texto, [{ ini: marco.ini, fin: marco.fin, texto: nuevo }])); return; }
      // otra celda: el marco sale de donde estaba y entra en la celda nueva
      const d = I().documento(texto), tabla = I().tablaDe(d, h);
      const lista = [{ f: Math.trunc(de.fila), c: Math.trunc(de.col), poner: nuevo }];
      if (m.enCelda) lista.push({ f: m.fila, c: m.col, quitar: true });
      const eds = U().crearMotor(texto).editarTabla(tabla, lista, (orig, t, e) => (e.quitar ? celdaSinMarco(orig, t, g.objeto) : celdaConMarco(orig, t, e.poner)));
      if (!m.enCelda) eds.push({ ini: marco.ini, fin: marco.fin, texto: '' });   // estaba en table:shapes
      texto = X().aplicar(texto, eds);
      p.escribir('content.xml', texto);
    });
  }

  // ---- insertar un gráfico nuevo ----------------------------------------------------------------------------------------------------------------
  const NS_CHART = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"';
  const CLASE_NUEVA = { col: 'chart:bar', bar: 'chart:bar', line: 'chart:line', pie: 'chart:circle' };
  // el content.xml de un gráfico nuevo: `series` = [{ nombre, categorias, valores }] con referencias de Excel
  function contenidoGrafico(tipo, series, leer, ancho, alto, titulo) {
    const clase = CLASE_NUEVA[tipo];
    if (!clase) throw Err('tipo de gráfico desconocido');
    const colores = G().COLORES.map((c) => '#' + c.toLowerCase());
    let estilos = '', n = 0;
    const nuevo = (props) => { estilos += '<style:style style:name="ch' + (++n) + '" style:family="chart">' + props + '</style:style>'; return 'ch' + n; };
    const sChart = nuevo('<style:graphic-properties draw:stroke="solid" svg:stroke-width="0.026cm" svg:stroke-color="#d9d9d9"/>');
    const sLeyenda = nuevo('<style:chart-properties chart:auto-position="true"/><style:graphic-properties draw:stroke="none" draw:fill="none"/><style:text-properties fo:font-size="10pt"/>');
    const sPlano = nuevo('<style:chart-properties' + (tipo === 'bar' ? ' chart:vertical="true"' : '') + ' chart:include-hidden-cells="false" chart:auto-position="true" chart:auto-size="true" chart:treat-empty-cells="leave-gap"/>');
    const sEjeX = nuevo('<style:chart-properties chart:display-label="true" chart:tick-marks-major-inner="false" chart:tick-marks-major-outer="false" chart:axis-position="0" chart:tick-mark-position="at-axis"/><style:graphic-properties draw:stroke="none" svg:stroke-color="#b3b3b3"/><style:text-properties fo:font-size="10pt"/>');
    const sEjeY = nuevo('<style:chart-properties chart:display-label="true" chart:tick-marks-major-inner="false" chart:tick-marks-major-outer="false" chart:axis-position="0" chart:tick-mark-position="at-axis"/><style:graphic-properties draw:stroke="none" svg:stroke-color="#b3b3b3"/><style:text-properties fo:font-size="10pt"/>');
    const sRejilla = nuevo('<style:graphic-properties svg:stroke-width="0.026cm" svg:stroke-color="#d9d9d9"/>');
    const sTitulo = nuevo('<style:chart-properties chart:auto-position="true"/><style:text-properties fo:font-size="13pt"/>');
    const cat = alOds(series.find((s) => s.categorias) && series.find((s) => s.categorias).categorias);
    const sers = series.map((s, i) => {
      const color = colores[i % colores.length];
      const est = tipo === 'line' ? nuevo('<style:chart-properties chart:symbol-type="none"/><style:graphic-properties svg:stroke-width="0.079cm" svg:stroke-color="' + color + '" svg:stroke-linecap="round" draw:fill-color="' + color + '"/>')
        : nuevo('<style:graphic-properties draw:stroke="none" draw:fill-color="' + color + '"/>');
      let puntos = '';
      if (tipo === 'pie') {
        const d = leer(s.valores), nPts = d ? d.valores.flat().length : 0;
        for (let q = 0; q < nPts; q++) puntos += '<chart:data-point chart:style-name="' + nuevo('<style:graphic-properties draw:stroke="solid" svg:stroke-width="0.053cm" svg:stroke-color="#ffffff" draw:fill-color="' + colores[q % colores.length] + '"/>') + '"/>';
      }
      return '<chart:series chart:style-name="' + est + '" chart:values-cell-range-address="' + X().escAttr(alOds(s.valores)) + '"' + (s.nombre ? ' chart:label-cell-address="' + X().escAttr(alOds(s.nombre)) + '"' : '') + ' chart:class="' + clase + '">' + puntos + '</chart:series>';
    }).join('');
    const ejes = tipo === 'pie' ? '' : '<chart:axis chart:dimension="x" chart:name="primary-x" chart:style-name="' + sEjeX + '">' + (cat ? '<chart:categories table:cell-range-address="' + X().escAttr(cat) + '"/>' : '') + '</chart:axis>'
      + '<chart:axis chart:dimension="y" chart:name="primary-y" chart:style-name="' + sEjeY + '"><chart:grid chart:style-name="' + sRejilla + '" chart:class="major"/></chart:axis>';
    const catPie = tipo === 'pie' && cat ? '<chart:axis chart:dimension="x" chart:name="primary-x" chart:style-name="' + sEjeX + '"><chart:categories table:cell-range-address="' + X().escAttr(cat) + '"/></chart:axis>' : '';
    const serie0 = series[0] && series[0].nombre ? leer(series[0].nombre) : null;
    const textoTitulo = titulo || (tipo === 'pie' && serie0 ? serie0.textos.flat()[0] : '');
    const xmlTitulo = textoTitulo ? '<chart:title chart:style-name="' + sTitulo + '"><text:p>' + X().esc(textoTitulo) + '</text:p></chart:title>' : '';
    const sLocal = tablaLocal(series.map((s) => ({ etiqueta: alOds(s.nombre), valores: alOds(s.valores) })), cat, leer);
    return '<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ' + NS_CHART + ' office:version="1.2"><office:automatic-styles>' + estilos + '</office:automatic-styles><office:body><office:chart>'
      + '<chart:chart svg:width="' + pulgadas(ancho) + '" svg:height="' + pulgadas(alto) + '" xlink:href=".." xlink:type="simple" chart:class="' + clase + '" chart:style-name="' + sChart + '">' + xmlTitulo
      + '<chart:legend chart:legend-position="' + (tipo === 'pie' ? 'end' : 'bottom') + '" style:legend-expansion="' + (tipo === 'pie' ? 'high' : 'wide') + '" chart:style-name="' + sLeyenda + '"/>'
      + '<chart:plot-area chart:style-name="' + sPlano + '">' + (tipo === 'pie' ? catPie : ejes) + sers + '</chart:plot-area>' + sLocal + '</chart:chart></office:chart></office:body></office:document-content>';
  }
  const ESTILOS_OBJETO = '<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles ' + NS_CHART + ' office:version="1.2"><office:styles/></office:document-styles>';

  async function insertar(bytes, h, r, tipo, leer, pos, nombres, esTexto) {
    const { libro } = await abrirTodo(bytes);
    const nombreHoja = libro.hojas[h].nombre;
    const series = G().seriesDeRango(nombreHoja, r, esTexto);
    if (tipo === 'pie') series.splice(1);
    const res = await I().conPaquete(bytes, async (p) => {
      let n = 1; while (p.existe('Object ' + n + '/content.xml') || p.existe('Object ' + n + '/')) n++;
      const objeto = 'Object ' + n;
      const ancho = pos && pos.ext ? pos.ext.w : Math.round(4.6807 * EMU.in), alto = pos && pos.ext ? pos.ext.h : Math.round(2.9165 * EMU.in);
      p.escribir(objeto + '/content.xml', contenidoGrafico(tipo, series, leer, ancho, alto, null));
      p.escribir(objeto + '/styles.xml', ESTILOS_OBJETO);
      // el manifiesto: el objeto es un paquete de tipo gráfico dentro del paquete
      const man = await p.leer('META-INF/manifest.xml');
      if (man) {
        const entradas = '<manifest:file-entry manifest:full-path="' + objeto + '/content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="' + objeto + '/styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="' + objeto + '/" manifest:media-type="application/vnd.oasis.opendocument.chart"/>';
        p.escribir('META-INF/manifest.xml', man.replace('</manifest:manifest>', () => entradas + '</manifest:manifest>'));
      }
      // el marco, en la celda de la posición
      const texto = await p.leer('content.xml');
      const d = I().documento(texto), tabla = I().tablaDe(d, h);
      const idsZ = [...texto.matchAll(/draw:z-index="(\d+)"/g)].map((x) => +x[1]);
      const rangos = [...new Set([series.find((s) => s.categorias) && series.find((s) => s.categorias).categorias, ...series.flatMap((s) => [s.nombre, s.valores])].filter(Boolean).map(alOds))];
      const marco = '<draw:frame draw:z-index="' + (Math.max(-1, ...idsZ) + 1) + '" draw:name="' + X().escAttr(((nombres && nombres.grafico) || 'Gráfico') + ' ' + (n)) + '" svg:width="' + pulgadas(ancho) + '" svg:height="' + pulgadas(alto) + '" svg:x="' + pulgadas(pos && pos.de ? pos.de.dx : 0) + '" svg:y="' + pulgadas(pos && pos.de ? pos.de.dy : 0) + '">'
        + '<draw:object draw:notify-on-update-of-ranges="' + X().escAttr(rangos.join(' ')) + '" xlink:href="./' + objeto + '" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>';
      const f = Math.max(0, Math.trunc(pos && pos.de ? pos.de.fila : r.f0)), c = Math.max(0, Math.trunc(pos && pos.de ? pos.de.col : r.c1 + 2));
      const eds = U().crearMotor(texto).editarTabla(tabla, [{ f, c }], (orig, t) => celdaConMarco(orig, t, marco));
      p.escribir('content.xml', X().aplicar(texto, eds));
      return { indice: objeto };
    });
    const despues = await abrirTodo(res.bytes);
    const g = (despues.estado.dibujables[h] || []).find((x) => x.objeto === res.indice);
    if (!g) throw Err('el gráfico nuevo no se pudo leer');
    return { bytes: res.bytes, indice: g.virtual.indice };
  }

  globalThis.GraficosOds = { listar, cambiarSeries, moverAncla, insertar, refrescarCaches, lectorDe: (ed) => globalThis.GraficosLibro.lectorDe(ed) };
  globalThis.OdsGraficos = { leerDefinicion, aOoxml, aExcel, lectorModelo, reunir, preparar, refrescarCaches, emu, PALETA, CLASES };
})();

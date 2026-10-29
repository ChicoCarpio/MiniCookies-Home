/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Gráficos de Mini Cookies Hojas (fase 3.3 del PLAN): leer los gráficos de una hoja (sus series y rangos), cambiar el rango de las series,
// mover o cambiar el tamaño del ancla, insertar un gráfico nuevo (columnas, barras, líneas, circular) y poner al día los valores que el
// gráfico guarda (c:numCache / c:strCache) con lo que tienen ahora las celdas. El gráfico lo dibuja el motor del sitio (XlsxObjetos con el
// de PowerPoint). Edita en el lugar solo las partes del gráfico y su dibujo. Sin DOM.
//
//   await GraficosLibro.listar(bytes, hoja) → [{ indice, rutaDibujo, rutaGrafico, tipo, series: [{ nombre, categorias, valores }] }]  (fórmulas c:f)
//   await GraficosLibro.cambiarSeries(bytes, hoja, indice, series, leer)   (series = [{ nombre, categorias, valores }]; leer(f) → [[valores]] | null)
//   await GraficosLibro.moverAncla(bytes, hoja, indice, { de: { col, dx, fila, dy }, a: {…}, ext: { w, h } })   (EMU)
//   await GraficosLibro.insertar(bytes, hoja, { f0, c0, f1, c1 }, tipo, leer, { de, a }, nombres)  → { bytes, indice }
//   await GraficosLibro.refrescarCaches(zip, leer, rutas?) → cantidad de gráficos que cambiaron (en el zip, sin regenerar el paquete)
(function () {
  const X = () => globalThis.XmlEditar;
  const Err = (m) => new XlsxEscritor.ErrorEscritor(m);
  const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
  const REL_DIBUJO = NS_R + '/drawing', REL_GRAFICO = NS_R + '/chart';
  const TIPO_DIBUJO = 'application/vnd.openxmlformats-officedocument.drawing+xml', TIPO_GRAFICO = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
  const ORDEN_HOJA = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
  const TIPOS = /<(?:\w+:)?(barChart|bar3DChart|lineChart|line3DChart|pieChart|pie3DChart|doughnutChart|areaChart|area3DChart|scatterChart|radarChart|bubbleChart|stockChart|surfaceChart|surface3DChart|ofPieChart)\b/;
  const texto = async (zip, ruta) => { const e = zip.file(ruta); return e ? e.async('string') : null; };
  const rutaHoja = (libro, h) => { const m = libro.hojas[h]; if (!m || m.tipo !== 'hoja' || !m.ruta) throw Err('la hoja no es de celdas'); return m.ruta; };
  const citarHoja = (n) => (/^[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'");
  const col = (n) => LibroXlsx.numeroACol(n);
  const abs = (hoja, r) => citarHoja(hoja) + '!$' + col(r.c0) + '$' + (r.f0 + 1) + (r.f0 === r.f1 && r.c0 === r.c1 ? '' : ':$' + col(r.c1) + '$' + (r.f1 + 1));
  async function conLibro(bytes, fn) {
    const libro = await LibroXlsx.abrir(bytes);
    const r = await fn(libro);
    const salida = await libro.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    return r && typeof r === 'object' ? Object.assign({}, r, { bytes: salida }) : salida;
  }

  // las anclas del dibujo en orden (las mismas que lee xlsx-modelo: hoja.anclas[k]) con el gráfico de cada una
  function anclasDe(xml) {
    const lista = [];
    const re = /<((?:\w+:)?)(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[^>]*>[\s\S]*?<\/\1\2>/g;
    let m;
    while ((m = re.exec(xml))) {
      const id = /<(?:\w+:)?chart\b[^>]*?\b(?:\w+:)?id="([^"]+)"/.exec(m[0]);
      lista.push({ ini: m.index, fin: m.index + m[0].length, xml: m[0], prefijo: m[1], tipo: m[2], idGrafico: id ? id[1] : null });
    }
    return lista;
  }
  async function dibujoDe(libro, h) {
    const ruta = rutaHoja(libro, h);
    const rels = await XlsxPaquete.leerRels(libro.zip, ruta);
    const rel = rels.find((r) => r.tipo === 'drawing' && !r.externa);
    if (!rel) return null;
    const xml = await texto(libro.zip, rel.destino);
    if (xml === null) return null;
    const relsDib = await XlsxPaquete.leerRels(libro.zip, rel.destino);
    return { ruta: rel.destino, xml, relsDib, anclas: anclasDe(xml) };
  }
  // series del XML del gráfico: c:tx (nombre), c:cat / c:xVal (categorías), c:val / c:yVal (valores)
  function seriesDe(xml) {
    const series = [];
    for (const m of xml.matchAll(/<(\w+:)?ser>([\s\S]*?)<\/\1?ser>/g)) {
      const s = m[2];
      const f = (etiquetas) => { for (const e of etiquetas) { const x = new RegExp('<(?:\\w+:)?' + e + '>[\\s\\S]*?<(?:\\w+:)?f>([\\s\\S]*?)</(?:\\w+:)?f>[\\s\\S]*?</(?:\\w+:)?' + e + '>').exec(s); if (x) return X().decodificar(x[1]); } return null; };
      series.push({ nombre: f(['tx']), categorias: f(['cat', 'xVal']), valores: f(['val', 'yVal']) });
    }
    return series;
  }
  async function listar(bytes, h) {
    const libro = await LibroXlsx.abrir(bytes);
    const d = await dibujoDe(libro, h);
    if (!d) return [];
    const lista = [];
    for (let k = 0; k < d.anclas.length; k++) {
      const a = d.anclas[k];
      if (!a.idGrafico) continue;
      const rel = d.relsDib.find((r) => r.id === a.idGrafico);
      const xml = rel ? await texto(libro.zip, rel.destino) : null;
      if (!xml) continue;
      lista.push({ indice: k, rutaDibujo: d.ruta, rutaGrafico: rel.destino, tipo: (TIPOS.exec(xml) || [])[1] || 'otro', series: seriesDe(xml) });
    }
    return lista;
  }

  // ---- cachés ----------------------------------------------------------------------------------------------------------
  const num = (v) => (typeof v === 'number' && Number.isFinite(v) ? v : typeof v === 'boolean' ? (v ? 1 : 0) : null);
  function cacheNum(p, valores, formato) {
    const plano = valores.flat();
    const pts = plano.map((v, i) => { const n = num(v); return n === null ? '' : `<${p}pt idx="${i}"><${p}v>${n}</${p}v></${p}pt>`; }).join('');
    return `<${p}numCache><${p}formatCode>${X().esc(formato || 'General')}</${p}formatCode><${p}ptCount val="${plano.length}"/>${pts}</${p}numCache>`;
  }
  function cacheStr(p, textos) {
    const plano = textos.flat();
    const pts = plano.map((v, i) => (v === null || v === undefined || v === '' ? '' : `<${p}pt idx="${i}"><${p}v>${X().esc(String(v))}</${p}v></${p}pt>`)).join('');
    return `<${p}strCache><${p}ptCount val="${plano.length}"/>${pts}</${p}strCache>`;
  }
  // reescribe el caché de cada c:numRef / c:strRef con lo que dice leer(f) → { valores: [[...]], textos: [[...]], formato } | null
  function refrescarXml(xml, leer) {
    return xml.replace(/<((?:\w+:)?)(numRef|strRef)>([\s\S]*?)<\/\1\2>/g, (todo, p, tipo, cuerpo) => {
      const fm = new RegExp('<' + p + 'f>([\\s\\S]*?)</' + p + 'f>').exec(cuerpo);
      if (!fm) return todo;
      const d = leer(X().decodificar(fm[1]));
      if (!d) return todo;
      const ext = (new RegExp('<' + p + 'extLst>[\\s\\S]*?</' + p + 'extLst>').exec(cuerpo) || [''])[0];
      const formato = (new RegExp('<' + p + 'formatCode>([\\s\\S]*?)</' + p + 'formatCode>').exec(cuerpo) || [])[1];
      // si lo guardado ya dice lo mismo (cantidad y cada punto), el gráfico queda tal cual
      const pts = new Map([...cuerpo.matchAll(new RegExp('<' + p + 'pt\\b[^>]*\\bidx="(\\d+)"[^>]*>\\s*<' + p + 'v>([\\s\\S]*?)</' + p + 'v>', 'g'))].map((m) => [+m[1], X().decodificar(m[2])]));
      const cuenta = +((new RegExp('<' + p + 'ptCount\\b[^>]*\\bval="(\\d+)"').exec(cuerpo) || [])[1] || -1);
      const nuevos = (tipo === 'numRef' ? d.valores : d.textos).flat();
      const igual = cuenta === nuevos.length && nuevos.every((v, i) => { const a = pts.get(i); if (tipo === 'numRef') { const n = num(v); return n === null ? a === undefined : a !== undefined && Math.abs(+a - n) <= 1e-12 * Math.max(1, Math.abs(n)); } return v === '' || v === null ? a === undefined || a === '' : a === String(v); });
      if (igual) return todo;
      const cache = tipo === 'numRef' ? cacheNum(p, d.valores, formato ? X().decodificar(formato) : d.formato) : cacheStr(p, d.textos);
      return `<${p}${tipo}>${fm[0]}${cache}${ext}</${p}${tipo}>`;
    });
  }
  async function refrescarCaches(zip, leer, rutas) {
    const lista = rutas || [];
    if (!rutas) zip.forEach((r) => { if (/^xl\/charts\/chart[^/]*\.xml$/i.test(r)) lista.push(r); });
    let n = 0;
    for (const r of lista) {
      const xml = await texto(zip, r);
      if (!xml) continue;
      const u = refrescarXml(xml, leer);
      if (u !== xml) { zip.file(r, u); n++; }
    }
    return n;
  }

  // ---- cambiar las series ------------------------------------------------------------------------------------------------
  // cambia el texto de c:f (y su caché) dentro del elemento `etiqueta` de la serie; si no estaba, lo agrega como referencia nueva
  function ponerRef(s, p, etiquetas, f, leer, numerico) {
    if (f === undefined) return s;
    for (const e of etiquetas) {
      const re = new RegExp('<' + p + e + '>([\\s\\S]*?)</' + p + e + '>');
      const m = re.exec(s);
      if (!m) continue;
      if (f === null || f === '') return s.replace(re, () => '');
      const conRef = /<(?:\w+:)?(?:numRef|strRef)>/.test(m[1]);
      const tipo = conRef ? (/<(?:\w+:)?numRef>/.test(m[1]) ? 'numRef' : 'strRef') : (numerico ? 'numRef' : 'strRef');
      const d = leer(f);
      const cache = tipo === 'numRef' ? cacheNum(p, d ? d.valores : [], d ? d.formato : 'General') : cacheStr(p, d ? d.textos : []);
      const nuevo = `<${p}${e}><${p}${tipo}><${p}f>${X().esc(f)}</${p}f>${cache}</${p}${tipo}></${p}${e}>`;
      return s.replace(re, () => nuevo);   // con función: los $ de las referencias no son patrones de reemplazo
    }
    return s;
  }
  async function cambiarSeries(bytes, h, k, series, leer) {
    return conLibro(bytes, async (libro) => {
      const d = await dibujoDe(libro, h);
      const a = d && d.anclas[k];
      const rel = a && a.idGrafico && d.relsDib.find((r) => r.id === a.idGrafico);
      if (!rel) throw Err('no hay un gráfico ahí');
      let xml = await texto(libro.zip, rel.destino);
      const sers = [...xml.matchAll(/<(\w+:)?ser>[\s\S]*?<\/\1?ser>/g)];
      if (series.length !== sers.length) throw Err('cambiar la cantidad de series se hace en Excel');
      const ed = sers.map((m, i) => {
        const p = m[1] || '';
        let s = m[0];
        const x = series[i];
        s = ponerRef(s, p, ['tx'], x.nombre, leer, false);
        s = ponerRef(s, p, ['cat', 'xVal'], x.categorias, leer, false);
        s = ponerRef(s, p, ['val', 'yVal'], x.valores, leer, true);
        return { ini: m.index, fin: m.index + m[0].length, texto: s };
      });
      xml = X().aplicar(xml, ed);
      libro.zip.file(rel.destino, xml);
    });
  }

  // ---- mover y cambiar el tamaño ------------------------------------------------------------------------------------------
  function marcador(p, nombre, m) {
    return `<${p}${nombre}><${p}col>${Math.max(0, Math.trunc(m.col))}</${p}col><${p}colOff>${Math.max(0, Math.round(m.dx))}</${p}colOff><${p}row>${Math.max(0, Math.trunc(m.fila))}</${p}row><${p}rowOff>${Math.max(0, Math.round(m.dy))}</${p}rowOff></${p}${nombre}>`;
  }
  async function moverAncla(bytes, h, k, pos) {
    return conLibro(bytes, async (libro) => {
      const d = await dibujoDe(libro, h);
      const a = d && d.anclas[k];
      if (!a) throw Err('no hay un objeto ahí');
      const p = a.prefijo;
      let t = a.xml;
      const reDe = new RegExp('<' + p + 'from>[\\s\\S]*?</' + p + 'from>'), reA = new RegExp('<' + p + 'to>[\\s\\S]*?</' + p + 'to>');
      const reExt = new RegExp('<' + p + 'ext\\b[^>]*/>'), rePos = new RegExp('<' + p + 'pos\\b[^>]*/>');
      if (a.tipo === 'twoCellAnchor') { t = t.replace(reDe, marcador(p, 'from', pos.de)).replace(reA, marcador(p, 'to', pos.a)); }
      else if (a.tipo === 'oneCellAnchor') { t = t.replace(reDe, marcador(p, 'from', pos.de)); if (pos.ext) t = t.replace(reExt, `<${p}ext cx="${Math.round(pos.ext.w)}" cy="${Math.round(pos.ext.h)}"/>`); }
      else { if (pos.abs) t = t.replace(rePos, `<${p}pos x="${Math.round(pos.abs.x)}" y="${Math.round(pos.abs.y)}"/>`); if (pos.ext) t = t.replace(reExt, `<${p}ext cx="${Math.round(pos.ext.w)}" cy="${Math.round(pos.ext.h)}"/>`); }
      libro.zip.file(d.ruta, d.xml.slice(0, a.ini) + t + d.xml.slice(a.fin));
    });
  }

  // ---- insertar un gráfico nuevo -------------------------------------------------------------------------------------------
  const COLORES = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '264478', '9E480E', '636363', '997300'];
  function xmlSerie(i, tipo, s, leer) {
    const d = (f) => (f ? leer(f) : null);
    const nombre = s.nombre ? `<c:tx><c:strRef><c:f>${X().esc(s.nombre)}</c:f>${cacheStr('c:', d(s.nombre) ? d(s.nombre).textos : [])}</c:strRef></c:tx>` : '';
    const color = COLORES[i % COLORES.length];
    const sp = tipo === 'line' ? `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>`
      : tipo === 'pie' ? '' : `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>`;
    const inv = tipo === 'col' || tipo === 'bar' ? '<c:invertIfNegative val="0"/>' : '';
    const pts = tipo === 'pie' ? (d(s.valores) ? d(s.valores).valores.flat() : []).map((_, k) => `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${COLORES[k % COLORES.length]}"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join('') : '';
    const cat = s.categorias ? `<c:cat><c:strRef><c:f>${X().esc(s.categorias)}</c:f>${cacheStr('c:', d(s.categorias) ? d(s.categorias).textos : [])}</c:strRef></c:cat>` : '';
    const dv = d(s.valores);
    const val = `<c:val><c:numRef><c:f>${X().esc(s.valores)}</c:f>${cacheNum('c:', dv ? dv.valores : [], dv ? dv.formato : 'General')}</c:numRef></c:val>`;
    const fin = tipo === 'line' ? '<c:smooth val="0"/>' : '';
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${nombre}${sp}${inv}${pts}${cat}${val}${fin}</c:ser>`;
  }
  const EJES = '<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="{{CAT}}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
    + '<c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="{{VAL}}"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>';
  function xmlGrafico(tipo, series, leer) {
    const sers = series.map((s, i) => xmlSerie(i, tipo, s, leer)).join('');
    let cuerpo;
    if (tipo === 'col' || tipo === 'bar') cuerpo = `<c:barChart><c:barDir val="${tipo}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${sers}<c:gapWidth val="219"/><c:overlap val="-27"/><c:axId val="500000001"/><c:axId val="500000002"/></c:barChart>` + EJES.replace('{{CAT}}', tipo === 'bar' ? 'l' : 'b').replace('{{VAL}}', tipo === 'bar' ? 'b' : 'l');
    else if (tipo === 'line') cuerpo = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}<c:marker val="1"/><c:axId val="500000001"/><c:axId val="500000002"/></c:lineChart>` + EJES.replace('{{CAT}}', 'b').replace('{{VAL}}', 'l');
    else if (tipo === 'pie') cuerpo = `<c:pieChart><c:varyColors val="1"/>${sers}<c:firstSliceAng val="0"/></c:pieChart>`;
    else throw Err('tipo de gráfico desconocido');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><c:roundedCorners val="0"/><c:chart><c:autoTitleDeleted val="${series.length > 1 && tipo !== 'pie' ? 1 : 0}"/><c:plotArea><c:layout/>${cuerpo}</c:plotArea>`
      + `<c:legend><c:legendPos val="${tipo === 'pie' ? 'r' : 'b'}"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:chartSpace>`;
  }
  // las series de un rango: la primera fila y la primera columna son nombres y categorías si son textos (como Excel); una serie por columna
  function seriesDeRango(hojaNombre, r, esTexto) {
    let filaNombres = true;
    for (let c = r.c0 + 1; c <= r.c1; c++) if (!esTexto(r.f0, c)) filaNombres = false;
    if (r.c0 === r.c1) filaNombres = esTexto(r.f0, r.c0);
    let colCategorias = r.c1 > r.c0;
    for (let f = r.f0 + (filaNombres ? 1 : 0); f <= r.f1 && colCategorias; f++) if (!esTexto(f, r.c0) && !(f === r.f0)) colCategorias = false;
    const f0 = r.f0 + (filaNombres ? 1 : 0), c0 = r.c0 + (colCategorias ? 1 : 0);
    if (f0 > r.f1 || c0 > r.c1) throw Err('el rango no tiene números para graficar');
    const series = [];
    for (let c = c0; c <= r.c1; c++) series.push({
      nombre: filaNombres ? abs(hojaNombre, { f0: r.f0, c0: c, f1: r.f0, c1: c }) : null,
      categorias: colCategorias ? abs(hojaNombre, { f0, c0: r.c0, f1: r.f1, c1: r.c0 }) : null,
      valores: abs(hojaNombre, { f0, c0: c, f1: r.f1, c1: c }),
    });
    return series;
  }
  async function insertar(bytes, h, r, tipo, leer, pos, nombres, esTexto) {
    return conLibro(bytes, async (libro) => {
      const zip = libro.zip, P = PaqueteEditable.desdeZip(zip);
      const rutaH = rutaHoja(libro, h);
      const series = seriesDeRango(libro.hojas[h].nombre, r, esTexto);
      if (tipo === 'pie') series.splice(1);
      // la parte del gráfico
      let n = 1; while (zip.file('xl/charts/chart' + n + '.xml')) n++;
      const rutaG = 'xl/charts/chart' + n + '.xml';
      P.escribir(rutaG, xmlGrafico(tipo, series, leer));
      await P.asegurarTipo('/' + rutaG, TIPO_GRAFICO);
      // el dibujo de la hoja (nuevo si no tiene)
      let d = await dibujoDe(libro, h);
      if (!d) {
        let m = 1; while (zip.file('xl/drawings/drawing' + m + '.xml')) m++;
        const rutaD = 'xl/drawings/drawing' + m + '.xml';
        P.escribir(rutaD, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}"></xdr:wsDr>`);
        await P.asegurarTipo('/' + rutaD, TIPO_DIBUJO);
        const id = await P.agregarRelacion(rutaH, REL_DIBUJO, '../drawings/drawing' + m + '.xml');
        let xml = await texto(zip, rutaH);
        const raiz = X().analizar(xml, { profundidad: 2 });
        const pref = Object.keys(raiz.attrs).find((a) => a.startsWith('xmlns:') && raiz.attrs[a] === NS_R);
        if (!pref) xml = X().aplicar(xml, [{ ini: raiz.ini, fin: raiz.finApertura, texto: X().ponerAtributo(xml.slice(raiz.ini, raiz.finApertura), 'xmlns:r', NS_R) }]);
        const r2 = X().analizar(xml, { profundidad: 2 });
        const k = ORDEN_HOJA.indexOf('drawing'), sig = r2.hijos.find((x) => ORDEN_HOJA.indexOf(x.local) > k);
        const posX = sig ? sig.ini : r2.iniCierre;
        xml = X().aplicar(xml, [{ ini: posX, fin: posX, texto: `<drawing ${pref ? pref.slice(6) : 'r'}:id="${id}"/>` }]);
        zip.file(rutaH, xml);
        d = await dibujoDe(libro, h);
      }
      const relId = await P.agregarRelacion(d.ruta, REL_GRAFICO, '../charts/chart' + n + '.xml');
      const raizD = X().analizar(d.xml, { profundidad: 1 });
      const p = (/^<(\w+:)?wsDr/.exec(d.xml.slice(raizD.ini)) || [])[1] || '';
      const ids = [...d.xml.matchAll(/<(?:\w+:)?cNvPr\b[^>]*\bid="(\d+)"/g)].map((x) => +x[1]);
      const idObj = Math.max(1, ...ids) + 1;
      const nombre = (nombres && nombres.grafico ? nombres.grafico : 'Gráfico') + ' ' + (d.anclas.length + 1);
      const ancla = `<${p}twoCellAnchor editAs="oneCell" xmlns:a="${NS_A}">${marcador(p, 'from', pos.de)}${marcador(p, 'to', pos.a)}<${p}graphicFrame macro=""><${p}nvGraphicFramePr><${p}cNvPr id="${idObj}" name="${X().escAttr(nombre)}"/><${p}cNvGraphicFramePr/></${p}nvGraphicFramePr>`
        + `<${p}xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></${p}xfrm><a:graphic xmlns:a="${NS_A}"><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" xmlns:r="${NS_R}" r:id="${relId}"/></a:graphicData></a:graphic></${p}graphicFrame><${p}clientData/></${p}twoCellAnchor>`;
      const xmlD = raizD.iniCierre >= 0 ? X().aplicar(d.xml, [{ ini: raizD.iniCierre, fin: raizD.iniCierre, texto: ancla }]) : X().aplicar(d.xml, [{ ini: raizD.ini, fin: raizD.fin, texto: d.xml.slice(raizD.ini, raizD.finApertura).replace(/\s*\/>$/, '>') + ancla + `</${p}wsDr>` }]);
      zip.file(d.ruta, xmlD);
      if (libro._hojasCargadas) libro._hojasCargadas.clear();
      return { indice: d.anclas.length };
    });
  }

  // lo que leen los gráficos, desde el libro abierto en el editor: leer('Hoja1!$B$2:$B$5') → { valores, textos, formato } (fila por fila)
  function lectorDe(ed) {
    return (f) => {
      let n;
      try { n = XlsFormulas._analizar(XlsFormulas._tokenizar(String(f).replace(/^=/, ''))); } catch (e) { return null; }
      if (n && n.t === 'paren') n = n.a;
      if (!n || n.t !== 'ref' || (n.modo !== 'celda' && n.modo !== 'area')) return null;
      const h = n.hoja ? ed.libro.hojas.findIndex((x) => x.nombre.toLowerCase() === String(n.hoja.a).toLowerCase()) : -1;
      if (h < 0 || !ed.hojas[h] || ed.hojas[h].sinCeldas) return null;
      const f1 = n.modo === 'area' ? n.f1 : n.f0, c1 = n.modo === 'area' ? n.c1 : n.c0;
      const r = { f0: Math.min(n.f0, f1), f1: Math.max(n.f0, f1), c0: Math.min(n.c0, c1), c1: Math.max(n.c0, c1) };
      if ((r.f1 - r.f0 + 1) * (r.c1 - r.c0 + 1) > 100000) return null;
      const valores = [], textos = [];
      let formato = null;
      for (let ff = r.f0; ff <= r.f1; ff++) {
        const fv = [], ft = [];
        for (let cc = r.c0; cc <= r.c1; cc++) {
          const x = ed.celda(h, ff, cc);
          const v = !x || x.tipo === 'vacio' ? null : x.tipo === 'error' ? null : x.valor;
          fv.push(v);
          let t = '';
          if (v !== null) {
            const fmt = (ed.libro.estilos.xf(x.estilo) || {}).formato || 'General';
            if (formato === null && typeof v === 'number') formato = fmt;
            t = typeof v === 'number' && typeof XlsxFormato !== 'undefined' ? XlsxFormato.formatear(v, fmt, { idioma: 'en', fecha1904: ed.libro.fecha1904 }).texto : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
          }
          ft.push(t);
        }
        valores.push(fv); textos.push(ft);
      }
      return { valores, textos, formato: formato || 'General' };
    };
  }

  globalThis.GraficosLibro = { lectorDe, listar, cambiarSeries, moverAncla, insertar, refrescarCaches, _interno: { seriesDe, refrescarXml, anclasDe, seriesDeRango, xmlGrafico, cacheNum, cacheStr, marcador, EJES, COLORES, NS_C, NS_A, NS_R, NS_XDR } };
})();

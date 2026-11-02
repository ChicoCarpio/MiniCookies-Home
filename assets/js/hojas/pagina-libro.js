/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Configurar página de Mini Cookies Hojas (fase 3.3 del PLAN): orientación, tamaño del papel, márgenes, ajustar a N páginas o escala, área de
// impresión, filas que se repiten, encabezado y pie, líneas de la cuadrícula y centrar. Queda en <pageSetup>, <pageMargins>, <printOptions>,
// <headerFooter>, <sheetPr><pageSetUpPr fitToPage> y los nombres _xlnm.Print_Area / _xlnm.Print_Titles, como los guarda Excel; el PDF del sitio
// (xlsx-paginas, xlsx-pdf) los respeta. Toma bytes y devuelve bytes editando solo lo que cambia. Sin DOM.
//
//   await PaginaLibro.leer(bytes, hoja) → configuración actual (lo que no está, con los valores de Excel)
//   await PaginaLibro.configurar(bytes, hoja, { orientacion: 'portrait'|'landscape', papel: 1|5|9…, margenes: { izq, der, sup, inf, enc, pie } (pulgadas),
//          ajuste: { ancho, alto } | null, escala: 10–400, area: 'A1:F40' | null, titulosFilas: '1:2' | null, encabezado, pie, cuadricula, centrarH, centrarV })
(function () {
  const X = () => globalThis.XmlEditar;
  const Err = (m) => new XlsxEscritor.ErrorEscritor(m);
  const ORDEN_HOJA = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
  const ORDEN_SHEETPR = ['tabColor', 'outlinePr', 'pageSetUpPr'];
  const texto = async (zip, ruta) => { const e = zip.file(ruta); return e ? e.async('string') : null; };
  const rutaHoja = (libro, h) => { const m = libro.hojas[h]; if (!m || m.tipo !== 'hoja' || !m.ruta) throw Err('la hoja no es de celdas'); return m.ruta; };
  const citarHoja = (n) => (/^[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'");
  const PAPELES = { 1: 'Carta', 5: 'Oficio (Legal)', 9: 'A4', 8: 'A3', 11: 'A5', 7: 'Ejecutivo' };
  const DEFECTO = { orientacion: 'portrait', papel: 1, margenes: { izq: 0.7, der: 0.7, sup: 0.75, inf: 0.75, enc: 0.3, pie: 0.3 }, ajuste: null, escala: 100, area: null, titulosFilas: null, encabezado: '', pie: '', cuadricula: false, centrarH: false, centrarV: false };

  // pone (o reemplaza, o quita con xml = '') un hijo de <worksheet> en su lugar
  function ponerEnHoja(doc, nombre, xml) {
    const raiz = X().analizar(doc, { profundidad: 2 });
    const n = X().hijo(raiz, nombre);
    if (n) return X().aplicar(doc, [{ ini: n.ini, fin: n.fin, texto: xml }]);
    if (!xml) return doc;
    const k = ORDEN_HOJA.indexOf(nombre), sig = raiz.hijos.find((h) => ORDEN_HOJA.indexOf(h.local) > k);
    const pos = sig ? sig.ini : raiz.iniCierre;
    return X().aplicar(doc, [{ ini: pos, fin: pos, texto: xml }]);
  }
  // el texto de &L &C &R de Excel → la sección central (lo más común); un texto propio se escribe en el centro
  const seccionCentral = (t) => { if (!t) return ''; const m = /&C((?:[^&]|&[^LCR])*)/.exec(t); return (m ? m[1] : /&[LR]/.test(t) ? '' : t).replace(/&&/g, '&'); };
  const aCodigo = (t) => (t ? '&C' + String(t).replace(/&(?![PNDTFAZG"\d&KSUXYEBI])/g, '&&') : '');

  async function leer(bytes, h) {
    const libro = await LibroXlsx.abrir(bytes);
    const xml = await texto(libro.zip, rutaHoja(libro, h));
    const raiz = X().analizar(xml, { profundidad: 3 });
    const ps = X().hijo(raiz, 'pageSetup'), pm = X().hijo(raiz, 'pageMargins'), po = X().hijo(raiz, 'printOptions'), hf = X().hijo(raiz, 'headerFooter'), sp = X().hijo(raiz, 'sheetPr');
    const psu = sp ? X().hijo(sp, 'pageSetUpPr') : null;
    const c = JSON.parse(JSON.stringify(DEFECTO));
    if (ps) {
      if (ps.attrs.orientation) c.orientacion = ps.attrs.orientation === 'landscape' ? 'landscape' : 'portrait';
      if (ps.attrs.paperSize) c.papel = +ps.attrs.paperSize;
      if (ps.attrs.scale) c.escala = +ps.attrs.scale;
      if (psu && (psu.attrs.fitToPage === '1' || psu.attrs.fitToPage === 'true')) c.ajuste = { ancho: ps.attrs.fitToWidth !== undefined ? +ps.attrs.fitToWidth : 1, alto: ps.attrs.fitToHeight !== undefined ? +ps.attrs.fitToHeight : 1 };
    }
    if (pm) c.margenes = { izq: +pm.attrs.left || 0, der: +pm.attrs.right || 0, sup: +pm.attrs.top || 0, inf: +pm.attrs.bottom || 0, enc: +pm.attrs.header || 0, pie: +pm.attrs.footer || 0 };
    if (po) { c.cuadricula = po.attrs.gridLines === '1' || po.attrs.gridLines === 'true'; c.centrarH = po.attrs.horizontalCentered === '1' || po.attrs.horizontalCentered === 'true'; c.centrarV = po.attrs.verticalCentered === '1' || po.attrs.verticalCentered === 'true'; }
    if (hf) {
      const de = (n) => { const x = X().hijo(hf, n); return x && x.iniCierre >= 0 ? X().decodificar(xml.slice(x.finApertura, x.iniCierre)) : ''; };
      c.encabezado = seccionCentral(de('oddHeader')); c.pie = seccionCentral(de('oddFooter'));
    }
    const nombre = (n) => libro.nombres.find((x) => x.nombre.toLowerCase() === n && x.hojaLocal === h);
    const area = nombre('_xlnm.print_area'), tit = nombre('_xlnm.print_titles');
    if (area) c.area = area.ref.replace(/^.*!/, '').replace(/\$/g, '');
    if (tit) { const m = /!\$?(\d+):\$?(\d+)/.exec(tit.ref); if (m) c.titulosFilas = m[1] + ':' + m[2]; }
    return c;
  }

  async function configurar(bytes, h, c) {
    const libro = await LibroXlsx.abrir(bytes);
    const ruta = rutaHoja(libro, h);
    let xml = await texto(libro.zip, ruta);
    const raiz = X().analizar(xml, { profundidad: 3 });
    const ps = X().hijo(raiz, 'pageSetup');
    // pageSetup: se conservan sus otros atributos (r:id de la impresora, copias, calidad…)
    let et = ps ? xml.slice(ps.ini, ps.iniCierre >= 0 ? ps.finApertura : ps.fin) : '<pageSetup/>';
    const cierre = ps && ps.iniCierre >= 0 ? xml.slice(ps.finApertura, ps.fin) : '';
    const poner = (k, v) => { et = X().ponerAtributo(et, k, v === null || v === undefined ? null : String(v)); };
    if (c.papel !== undefined) { if (!(Number.isInteger(+c.papel) && +c.papel >= 1 && +c.papel <= 118)) throw Err('tamaño de papel no válido'); poner('paperSize', +c.papel === 1 ? null : +c.papel); }
    if (c.orientacion !== undefined) poner('orientation', c.orientacion === 'landscape' ? 'landscape' : 'portrait');
    if (c.ajuste) { poner('fitToWidth', Math.max(0, Math.trunc(+c.ajuste.ancho || 0)) === 1 ? null : Math.max(0, Math.trunc(+c.ajuste.ancho || 0))); poner('fitToHeight', Math.max(0, Math.trunc(+c.ajuste.alto || 0)) === 1 ? null : Math.max(0, Math.trunc(+c.ajuste.alto || 0))); poner('scale', null); }
    else if (c.escala !== undefined) { const e = Math.trunc(+c.escala); if (!(e >= 10 && e <= 400)) throw Err('la escala va de 10 a 400 %'); poner('scale', e === 100 ? null : e); poner('fitToWidth', null); poner('fitToHeight', null); }
    const conAtributos = /\s[\w:]+=/.test(et);
    xml = ponerEnHoja(xml, 'pageSetup', conAtributos || cierre ? (cierre ? et + cierre : et.replace(/\s*\/?>$/, '/>')) : '');
    // márgenes: los seis son obligatorios en <pageMargins>
    if (c.margenes) {
      const m = Object.assign({}, DEFECTO.margenes, c.margenes);
      for (const k in m) if (!(m[k] >= 0 && m[k] <= 49)) throw Err('margen no válido');
      xml = ponerEnHoja(xml, 'pageMargins', `<pageMargins left="${m.izq}" right="${m.der}" top="${m.sup}" bottom="${m.inf}" header="${m.enc}" footer="${m.pie}"/>`);
    }
    // printOptions: solo si algo no es lo de siempre
    if (c.cuadricula !== undefined || c.centrarH !== undefined || c.centrarV !== undefined) {
      const at = [c.centrarH ? 'horizontalCentered="1"' : '', c.centrarV ? 'verticalCentered="1"' : '', c.cuadricula ? 'gridLines="1"' : ''].filter(Boolean).join(' ');
      xml = ponerEnHoja(xml, 'printOptions', at ? `<printOptions ${at}/>` : '');
    }
    // encabezado y pie: el texto va al centro; lo demás que tuvieran (izquierda, derecha, primera página) se conserva
    if (c.encabezado !== undefined || c.pie !== undefined) {
      const r2 = X().analizar(xml, { profundidad: 3 }), hf = X().hijo(r2, 'headerFooter');
      const viejo = (n) => { const x = hf && X().hijo(hf, n); return x && x.iniCierre >= 0 ? X().decodificar(xml.slice(x.finApertura, x.iniCierre)) : ''; };
      const conCentro = (antes, nuevo) => { const sinC = antes.replace(/&C(?:[^&]|&[^LCR])*/g, ''); return (sinC + aCodigo(nuevo)) || ''; };
      const cab = c.encabezado !== undefined ? conCentro(viejo('oddHeader'), c.encabezado) : viejo('oddHeader');
      const pie = c.pie !== undefined ? conCentro(viejo('oddFooter'), c.pie) : viejo('oddFooter');
      let resto = '';
      if (hf && hf.iniCierre >= 0) resto = xml.slice(hf.finApertura, hf.iniCierre).replace(/<oddHeader>[\s\S]*?<\/oddHeader>|<oddHeader\/>|<oddFooter>[\s\S]*?<\/oddFooter>|<oddFooter\/>/g, '');
      const abre = hf ? xml.slice(hf.ini, hf.iniCierre >= 0 ? hf.finApertura : hf.fin).replace(/\s*\/>$/, '>') : '<headerFooter>';
      const cuerpo = (cab ? `<oddHeader>${X().esc(cab)}</oddHeader>` : '') + (pie ? `<oddFooter>${X().esc(pie)}</oddFooter>` : '') + resto;
      xml = ponerEnHoja(xml, 'headerFooter', cuerpo || (hf && /\s\w/.test(abre.replace(/^<headerFooter/, '').replace(/>$/, ''))) ? abre + cuerpo + '</headerFooter>' : '');
    }
    // sheetPr/pageSetUpPr fitToPage
    if (c.ajuste !== undefined) {
      const r3 = X().analizar(xml, { profundidad: 3 }), sp = X().hijo(r3, 'sheetPr');
      const psu = sp ? X().hijo(sp, 'pageSetUpPr') : null;
      if (psu) xml = X().aplicar(xml, [{ ini: psu.ini, fin: psu.fin, texto: X().ponerAtributo(xml.slice(psu.ini, psu.iniCierre >= 0 ? psu.finApertura : psu.fin).replace(/\s*\/?>$/, '/>'), 'fitToPage', c.ajuste ? '1' : null) }]);
      else if (c.ajuste) {
        if (sp && sp.iniCierre >= 0) { const k = ORDEN_SHEETPR.indexOf('pageSetUpPr'), sig = X().hijos(sp).find((x) => ORDEN_SHEETPR.indexOf(x.local) > k); const pos = sig ? sig.ini : sp.iniCierre; xml = X().aplicar(xml, [{ ini: pos, fin: pos, texto: '<pageSetUpPr fitToPage="1"/>' }]); }
        else if (sp) xml = X().aplicar(xml, [{ ini: sp.ini, fin: sp.fin, texto: xml.slice(sp.ini, sp.fin).replace(/\s*\/>$/, '>') + '<pageSetUpPr fitToPage="1"/></sheetPr>' }]);
        else xml = ponerEnHoja(xml, 'sheetPr', '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
      }
    }
    libro.zip.file(ruta, xml);
    // área de impresión y títulos: nombres de la hoja
    let wb = await texto(libro.zip, libro.rutaLibro);
    const hoja = citarHoja(libro.hojas[h].nombre);
    const nombres = [];
    if (c.area !== undefined) {
      const r = c.area ? LibroXlsx.decodificarRango(String(c.area).replace(/\$/g, '').toUpperCase()) : null;
      if (c.area && !r) throw Err('el área de impresión no es un rango');
      nombres.push(['_xlnm.Print_Area', r ? hoja + '!$' + LibroXlsx.numeroACol(r.c0) + '$' + (r.f0 + 1) + ':$' + LibroXlsx.numeroACol(r.c1) + '$' + (r.f1 + 1) : null]);
    }
    if (c.titulosFilas !== undefined) {
      const m = c.titulosFilas ? /^\$?(\d+)(?::\$?(\d+))?$/.exec(String(c.titulosFilas).trim()) : null;
      if (c.titulosFilas && !m) throw Err('las filas que se repiten van como 1:2');
      nombres.push(['_xlnm.Print_Titles', m ? hoja + '!$' + m[1] + ':$' + (m[2] || m[1]) : null]);
    }
    for (const [n, valor] of nombres) {
      const raizW = X().analizar(wb, { profundidad: 3 }), dn = X().hijo(raizW, 'definedNames');
      const lista = dn ? X().hijos(dn, 'definedName') : [];
      const x = lista.find((d) => (d.attrs.name || '').toLowerCase() === n.toLowerCase() && d.attrs.localSheetId === String(h));
      if (x) wb = X().aplicar(wb, [valor === null ? (lista.length === 1 ? { ini: dn.ini, fin: dn.fin, texto: '' } : { ini: x.ini, fin: x.fin, texto: '' }) : { ini: x.finApertura, fin: x.iniCierre, texto: X().esc(valor) }]);
      else if (valor !== null) {
        const nodo = `<definedName name="${n}" localSheetId="${h}">${X().esc(valor)}</definedName>`;
        if (dn && dn.iniCierre >= 0) wb = X().aplicar(wb, [{ ini: dn.iniCierre, fin: dn.iniCierre, texto: nodo }]);
        else { const orden = ['fileVersion', 'fileSharing', 'workbookPr', 'workbookProtection', 'bookViews', 'sheets', 'functionGroups', 'externalReferences', 'definedNames']; const sig = raizW.hijos.find((q) => orden.indexOf(q.local) < 0); const pos = sig ? sig.ini : raizW.iniCierre; wb = X().aplicar(wb, [{ ini: pos, fin: pos, texto: '<definedNames>' + nodo + '</definedNames>' }]); }
      }
    }
    if (nombres.length) libro.zip.file(libro.rutaLibro, wb);
    if (libro._hojasCargadas) libro._hojasCargadas.clear();
    return libro.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  globalThis.PaginaLibro = { leer, configurar, PAPELES, DEFECTO };
})();

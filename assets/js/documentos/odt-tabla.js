/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Tablas de un .odt para EdicionDoc (Fase 6b, etapa 6b.5): crear una tabla y cambiar su estructura (filas y columnas) sobre el XML de un <table:table>. Sin DOM.
// La misma interfaz que TablaDoc (tabla-doc.js, Word): nueva, insertarFila, eliminarFila, insertarColumna, eliminarColumna. Solo entiende tablas «planas»:
// columnas y filas directas, celdas sin combinar ni repetidas (`number-*-repeated`, `*-spanned`, celdas cubiertas). Con otra estructura devuelve null y el editor
// lo dice: mejor no tocar que dañar. Los estilos de tabla, columna y celda que se crean son automáticos (`MCTb`, `MCCo`, `MCCe`, ver odt-formato.js).
//
// Requiere xml-editar.js y odt-estilos.js. `crear(modelo, { formato })` lo llama OdtParrafo.
(function () {
  'use strict';
  const X = () => globalThis.XmlEditar;
  const A = (n) => globalThis.OdtEstilos.A(n);
  const NS = { table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0', text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0', office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
    style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0', fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0' };
  const decl = (...ns) => ns.map((n) => ` xmlns:${n}="${NS[n]}"`).join('');
  const redondo = (v) => Math.round(v * 100) / 100;

  function crear(modelo, { formato }) {
    const est = modelo.est;
    const estiloTabla = (ancho) => formato.registrarXml('MCTb', 'tabla|' + ancho, (n) => `<style:style${decl('style', 'table')} style:name="${n}" style:family="table"><style:table-properties style:width="${ancho}pt" table:align="margins"/></style:style>`);
    const estiloColumna = (ancho) => formato.registrarXml('MCCo', 'columna|' + ancho, (n) => `<style:style${decl('style')} style:name="${n}" style:family="table-column"><style:table-column-properties style:column-width="${ancho}pt"/></style:style>`);
    const estiloCelda = () => formato.registrarXml('MCCe', 'celda', (n) => `<style:style${decl('style', 'fo')} style:name="${n}" style:family="table-cell"><style:table-cell-properties fo:padding="2.8pt" fo:border="0.5pt solid #000000"/></style:style>`);
    let nTablas = 0;
    const nombreTabla = () => { let n; do n = 'MCTabla' + (++nTablas); while (modelo.xml.includes('table:name="' + n + '"')); return n; };

    function nueva(filas, columnas, anchoPt) {
      filas = Math.min(Math.max(Math.floor(filas) || 1, 1), 200); columnas = Math.min(Math.max(Math.floor(columnas) || 1, 1), 63);
      const total = redondo(Math.max(columnas * 10, anchoPt)), w = redondo(total / columnas);
      const celda = `<table:table-cell table:style-name="${estiloCelda()}" office:value-type="string"><text:p/></table:table-cell>`;
      const col = `<table:table-column table:style-name="${estiloColumna(w)}"/>`;
      return `<table:table${decl('table', 'text', 'office')} table:name="${nombreTabla()}" table:style-name="${estiloTabla(redondo(w * columnas))}">${col.repeat(columnas)}`
        + `<table:table-row>${celda.repeat(columnas)}</table:table-row>`.repeat(filas) + '</table:table>';
    }

    // La tabla plana: { raiz, cols, filas: [{ tr, celdas }] } o null
    function analizar(xml) {
      const raiz = X().analizar(xml);
      if (raiz.local !== 'table') return null;
      const cols = [], filas = [];
      for (const h of raiz.hijos) {
        if (h.local === 'table-column') { const a = A(h); if (parseInt(a['number-columns-repeated'], 10) > 1) return null; cols.push(h); }
        else if (h.local === 'table-row') {
          const a = A(h);
          if (parseInt(a['number-rows-repeated'], 10) > 1) return null;
          const celdas = [];
          for (const c of h.hijos) {
            if (c.local !== 'table-cell') return null;
            const q = A(c);
            if (parseInt(q['number-columns-repeated'], 10) > 1 || parseInt(q['number-columns-spanned'], 10) > 1 || parseInt(q['number-rows-spanned'], 10) > 1) return null;
            celdas.push(c);
          }
          filas.push({ tr: h, celdas });
        } else if (h.local === 'table-header-rows' || h.local === 'table-rows' || h.local === 'table-columns' || h.local === 'table-row-group' || h.local === 'table-header-columns') return null;
      }
      if (!cols.length || !filas.length || filas.some((f) => f.celdas.length !== cols.length)) return null;
      return { raiz, cols, filas };
    }
    const prefijoTexto = (xml) => { const m = /<([\w.-]+:)p[\s>/]/.exec(xml); return m ? m[1] : 'text:'; };
    // Una celda nueva y vacía con el estilo de `modelo`
    function celdaVacia(xml, c) {
      const nom = c.nombre, p = prefijoTexto(xml), a = Object.keys(c.attrs).find((k) => k.slice(k.indexOf(':') + 1) === 'style-name');
      return `<${nom}${a ? ` ${a}="${X().escAttr(c.attrs[a])}"` : ''} office:value-type="string" xmlns:office="${NS.office}"><${p}p/></${nom}>`;
    }
    function filaVacia(xml, f) {
      const tag = xml.slice(f.tr.ini, f.tr.finApertura);
      return tag + f.celdas.map((c) => celdaVacia(xml, c)).join('') + `</${f.tr.nombre}>`;
    }

    function insertarFila(xml, i, despues) {
      const t = analizar(xml);
      if (!t || i < 0 || i >= t.filas.length) return null;
      const f = t.filas[i], donde = despues ? f.tr.fin : f.tr.ini;
      return X().aplicar(xml, [{ ini: donde, fin: donde, texto: filaVacia(xml, f) }]);
    }
    function eliminarFila(xml, i) {
      const t = analizar(xml);
      if (!t || i < 0 || i >= t.filas.length) return null;
      if (t.filas.length === 1) return '';
      return X().aplicar(xml, [{ ini: t.filas[i].tr.ini, fin: t.filas[i].tr.fin, texto: '' }]);
    }
    const anchoDe = (col) => { const s = A(col)['style-name'], v = s ? est.columna(s).ancho : undefined; return v > 0 ? v : null; };
    // Vuelve a repartir el ancho total entre las columnas que quedan (`nuevas`: ancho por columna existente; null = se queda igual)
    function reajustar(t, anchos) {
      const ed = [];
      t.cols.forEach((c, k) => {
        if (anchos[k] === null || anchos[k] === undefined) return;
        const clave = Object.keys(c.attrs).find((q) => q.slice(q.indexOf(':') + 1) === 'style-name');
        const tag = xml0.slice(c.ini, c.finApertura), nuevo = X().ponerAtributo(tag, clave || 'table:style-name', estiloColumna(redondo(anchos[k])));
        ed.push({ ini: c.ini, fin: c.finApertura, texto: nuevo });
      });
      return ed;
    }
    let xml0 = '';
    function insertarColumna(xml, c, despues) {
      const t = analizar(xml);
      if (!t || c < 0 || c >= t.cols.length) return null;
      xml0 = xml;
      const n = t.cols.length, ws = t.cols.map(anchoDe), total = ws.every((w) => w) ? ws.reduce((s, w) => s + w, 0) : null;
      const nuevoAncho = total ? total / (n + 1) : null;
      const ed = total ? reajustar(t, ws.map((w) => w * n / (n + 1))) : [];
      const col = t.cols[c], tagCol = xml.slice(col.ini, col.fin);
      const dc = despues ? col.fin : col.ini;
      const clave = Object.keys(col.attrs).find((q) => q.slice(q.indexOf(':') + 1) === 'style-name');
      ed.push({ ini: dc, fin: dc, texto: nuevoAncho ? X().ponerAtributo(tagCol, clave || 'table:style-name', estiloColumna(redondo(nuevoAncho))) : tagCol });
      for (const f of t.filas) { const ce = f.celdas[c], d = despues ? ce.fin : ce.ini; ed.push({ ini: d, fin: d, texto: celdaVacia(xml, ce) }); }
      return aplicarSinChoque(xml, ed);
    }
    function eliminarColumna(xml, c) {
      const t = analizar(xml);
      if (!t || c < 0 || c >= t.cols.length) return null;
      if (t.cols.length === 1) return '';
      xml0 = xml;
      const n = t.cols.length, ws = t.cols.map(anchoDe), total = ws.every((w) => w) ? ws.reduce((s, w) => s + w, 0) : null;
      const resto = total ? total - ws[c] : null;
      const ed = total ? reajustar(t, ws.map((w, k) => (k === c ? null : w * total / resto))) : [];
      ed.push({ ini: t.cols[c].ini, fin: t.cols[c].fin, texto: '' });
      for (const f of t.filas) ed.push({ ini: f.celdas[c].ini, fin: f.celdas[c].fin, texto: '' });
      void n;
      return aplicarSinChoque(xml, ed);
    }
    // (el estilo de la columna donde se inserta cambia de etiqueta y a la vez recibe una copia: se juntan las ediciones que tocan la misma etiqueta)
    function aplicarSinChoque(xml, ed) { return X().aplicar(xml, ed); }

    return { nueva, analizar, insertarFila, eliminarFila, insertarColumna, eliminarColumna };
  }

  globalThis.OdtTabla = { crear };
})();

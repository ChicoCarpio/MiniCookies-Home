/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Tablas de Documentos (Fase 6, etapa 6.6): crear una tabla y cambiar su estructura (filas y columnas) sobre el XML de un <w:tbl>. Sin DOM.
//
// Solo entiende tablas «planas»: tblPr, tblGrid y filas (w:tr) directas, con celdas (w:tc) directas, sin gridBefore/gridAfter. Con otra
// estructura (filas dentro de un control de contenido, columnas desplazadas) las funciones devuelven null y el editor lo dice: mejor no tocar
// que dañar. Las celdas combinadas (gridSpan, vMerge) se respetan.
//
//   TablaDoc.nueva(filas, columnas, anchoPt) → XML de una tabla vacía con bordes
//   TablaDoc.analizar(xml) → { raiz, grid:[tw], filas:[{ tr, celdas:[{ tc, col, span, vm }] }] } | null
//   TablaDoc.insertarFila(xml, i, despues) · eliminarFila(xml, i) · insertarColumna(xml, c, despues) · eliminarColumna(xml, c)
//     → XML nuevo, o null si no se puede; eliminar la última fila o columna devuelve '' (la tabla desaparece)
//
// Requiere xml-editar.js y formato-doc.js.
(function () {
  'use strict';
  const X = () => globalThis.XmlEditar;
  const ORDEN_TCPR = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark'];
  const hijoDe = (n, local) => (n ? n.hijos.find((h) => h.local === local) || null : null);
  const val = (n) => (n ? X().attr(n, 'val') : null);
  const pfDe = (raiz) => { const i = raiz.nombre.indexOf(':'); return i < 0 ? '' : raiz.nombre.slice(0, i + 1); };

  function nueva(filas, columnas, anchoPt) {
    filas = Math.min(Math.max(Math.floor(filas) || 1, 1), 200); columnas = Math.min(Math.max(Math.floor(columnas) || 1, 1), 63);
    const total = Math.max(columnas * 200, Math.round(anchoPt * 20)), w = Math.floor(total / columnas);
    const borde = (l) => `<w:${l} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`;
    let x = `<w:tbl><w:tblPr><w:tblW w:w="${w * columnas}" w:type="dxa"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(borde).join('')}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid>`;
    x += `<w:gridCol w:w="${w}"/>`.repeat(columnas) + '</w:tblGrid>';
    x += `<w:tr>${`<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`.repeat(columnas)}</w:tr>`.repeat(filas);
    return x + '</w:tbl>';
  }

  function analizar(xml) {
    const raiz = X().analizar(xml);
    if (raiz.local !== 'tbl') return null;
    const gridEl = hijoDe(raiz, 'tblGrid');
    if (!gridEl) return null;
    const grid = gridEl.hijos.filter((h) => h.local === 'gridCol').map((g) => parseInt(X().attr(g, 'w'), 10) || 0);
    const filas = [];
    for (const h of raiz.hijos) {
      if (h.local === 'tblPr' || h.local === 'tblGrid' || h.local === 'tblPrChange' || h.local === 'bookmarkStart' || h.local === 'bookmarkEnd') continue;
      if (h.local !== 'tr') return null;
      const trPr = hijoDe(h, 'trPr');
      if (trPr && trPr.hijos.some((q) => q.local === 'gridBefore' || q.local === 'gridAfter')) return null;
      const celdas = [];
      let col = 0;
      for (const c of h.hijos) {
        if (c.local === 'trPr' || c.local === 'tblPrEx') continue;
        if (c.local !== 'tc') return null;
        const tcPr = hijoDe(c, 'tcPr');
        const span = Math.max(parseInt(val(hijoDe(tcPr, 'gridSpan')), 10) || 1, 1);
        const vmEl = hijoDe(tcPr, 'vMerge');
        celdas.push({ tc: c, tcPr, col, span, vm: vmEl ? (val(vmEl) === 'restart' ? 'restart' : 'continue') : null });
        col += span;
      }
      if (col !== grid.length) return null;
      filas.push({ tr: h, celdas });
    }
    return filas.length && grid.length ? { raiz, grid, filas } : null;
  }

  // El <w:tcPr> de una celda con otros elementos (null quita uno)
  function tcPrCon(xml, c, elementos, pf) {
    const actual = c.tcPr ? xml.slice(c.tcPr.ini, c.tcPr.fin) : '';
    return globalThis.FormatoDoc.conElementos(actual, 'tcPr', ORDEN_TCPR, elementos, pf);
  }
  const cambioTcPr = (xml, c, elementos, pf) => {
    const nuevo = tcPrCon(xml, c, elementos, pf);
    if (c.tcPr) return { ini: c.tcPr.ini, fin: c.tcPr.fin, texto: nuevo };
    return { ini: c.tc.finApertura, fin: c.tc.finApertura, texto: nuevo };
  };
  const parrafoVacio = (xml, c, pf) => {
    const p = c.tc.hijos.find((h) => h.local === 'p');
    const pPr = hijoDe(p, 'pPr');
    const pPrXml = pPr ? xml.slice(pPr.ini, pPr.fin).replace(/<(?:\w+:)?sectPr\b[\s\S]*?<\/(?:\w+:)?sectPr>/, '') : '';
    return pPrXml ? `<${pf}p>${pPrXml}</${pf}p>` : `<${pf}p/>`;
  };
  const anchoDe = (grid, c) => { let t = 0; for (let k = c.col; k < c.col + c.span; k++) t += grid[k] || 0; return t; };

  function insertarFila(xml, i, despues) {
    const t = analizar(xml);
    if (!t || i < 0 || i >= t.filas.length) return null;
    const pf = pfDe(t.raiz), ref = t.filas[i], sig = t.filas[i + 1];
    const celdas = ref.celdas.map((c) => {
      // dentro de una combinación vertical, la fila nueva sigue en ella; la celda de abajo se busca por columna (con celdas combinadas en horizontal el índice no coincide)
      const abajo = sig && sig.celdas.find((x) => x.col === c.col);
      const sigue = despues ? c.vm && abajo && abajo.vm === 'continue' : c.vm === 'continue';
      const el = new Map([['vMerge', sigue ? `<${pf}vMerge/>` : null]]);
      const tcPr = tcPrCon(xml, c, el, pf);
      return `<${pf}tc>${tcPr}${parrafoVacio(xml, c, pf)}</${pf}tc>`;
    });
    const trPr = hijoDe(ref.tr, 'trPr');
    const fila = `<${pf}tr>${trPr ? xml.slice(trPr.ini, trPr.fin) : ''}${celdas.join('')}</${pf}tr>`;
    const donde = despues ? ref.tr.fin : ref.tr.ini;
    return X().aplicar(xml, [{ ini: donde, fin: donde, texto: fila }]);
  }

  function eliminarFila(xml, i) {
    const t = analizar(xml);
    if (!t || i < 0 || i >= t.filas.length) return null;
    if (t.filas.length === 1) return '';
    const pf = pfDe(t.raiz), ed = [];
    // si la fila abría una combinación vertical, la de abajo pasa a abrirla
    const sig = t.filas[i + 1];
    t.filas[i].celdas.forEach((c) => {
      const abajo = sig && sig.celdas.find((x) => x.col === c.col);
      if (c.vm === 'restart' && abajo && abajo.vm === 'continue') ed.push(cambioTcPr(xml, abajo, new Map([['vMerge', `<${pf}vMerge ${pf}val="restart"/>`]]), pf));
    });
    const f = t.filas[i].tr;
    ed.push({ ini: f.ini, fin: f.fin, texto: '' });
    return X().aplicar(xml, ed);
  }

  // Los anchos nuevos de las columnas (conservando el ancho total) y las ediciones que los escriben en tblGrid, tblW y cada tcW
  function escribirAnchos(xml, t, grid, filasCeldas, pf, ed, total) {
    const g = t.raiz.hijos.find((h) => h.local === 'tblGrid');
    ed.push({ ini: g.ini, fin: g.fin, texto: `<${pf}tblGrid>${grid.map((w) => `<${pf}gridCol ${pf}w="${w}"/>`).join('')}</${pf}tblGrid>` });
    const tblPr = t.raiz.hijos.find((h) => h.local === 'tblPr'), tblW = hijoDe(tblPr, 'tblW');
    if (tblW && (X().attr(tblW, 'type') || 'dxa') === 'dxa') ed.push({ ini: tblW.ini, fin: tblW.fin, texto: `<${pf}tblW ${pf}w="${total}" ${pf}type="dxa"/>` });
    return filasCeldas;
  }

  function insertarColumna(xml, c, despues) {
    const t = analizar(xml);
    if (!t || c < 0 || c >= t.grid.length) return null;
    const pf = pfDe(t.raiz), borde = despues ? c + 1 : c;
    const total = t.grid.reduce((a, b) => a + b, 0), nuevoAncho = Math.max(Math.round(total / t.grid.length), 200);
    const grid = t.grid.slice();
    grid.splice(borde, 0, nuevoAncho);
    const suma = grid.reduce((a, b) => a + b, 0), esc = total / suma;
    const g2 = grid.map((w) => Math.max(Math.round(w * esc), 100));
    const ed = [];
    escribirAnchos(xml, t, g2, null, pf, ed, g2.reduce((a, b) => a + b, 0));
    for (const fila of t.filas) {
      // la celda que cubre la columna de referencia aporta el formato; si el borde cae dentro de una combinada, esa se ensancha
      let nueva = null;
      for (const ce of fila.celdas) {
        const dentro = ce.col < borde && borde < ce.col + ce.span;
        const nuevoCol = ce.col >= borde ? ce.col + 1 : ce.col, nuevoSpan = dentro ? ce.span + 1 : ce.span;
        const el = new Map();
        if (dentro) el.set('gridSpan', `<${pf}gridSpan ${pf}val="${nuevoSpan}"/>`);
        const w = g2.slice(nuevoCol, nuevoCol + nuevoSpan).reduce((a, b) => a + b, 0);
        el.set('tcW', `<${pf}tcW ${pf}w="${w}" ${pf}type="dxa"/>`);
        ed.push(cambioTcPr(xml, ce, el, pf));
        if (!nueva && ce.col >= borde) nueva = ce;
      }
      const ref = fila.celdas.find((ce) => ce.col <= c && c < ce.col + ce.span) || fila.celdas[fila.celdas.length - 1];
      if (ref.col < borde && borde < ref.col + ref.span) continue;     // ya cubierta por la combinada
      const tcPr = tcPrCon(xml, ref, new Map([['gridSpan', null], ['vMerge', null], ['tcW', `<${pf}tcW ${pf}w="${g2[borde]}" ${pf}type="dxa"/>`]]), pf);
      const celda = `<${pf}tc>${tcPr}${parrafoVacio(xml, ref, pf)}</${pf}tc>`;
      const donde = nueva ? nueva.tc.ini : fila.tr.iniCierre;
      ed.push({ ini: donde, fin: donde, texto: celda });
    }
    return X().aplicar(xml, ed);
  }

  function eliminarColumna(xml, c) {
    const t = analizar(xml);
    if (!t || c < 0 || c >= t.grid.length) return null;
    if (t.grid.length === 1) return '';
    const pf = pfDe(t.raiz), ed = [];
    const grid = t.grid.slice();
    grid.splice(c, 1);
    escribirAnchos(xml, t, grid, null, pf, ed, grid.reduce((a, b) => a + b, 0));
    for (const fila of t.filas) {
      for (const ce of fila.celdas) {
        const cubre = ce.col <= c && c < ce.col + ce.span;
        if (cubre && ce.span === 1) { ed.push({ ini: ce.tc.ini, fin: ce.tc.fin, texto: '' }); continue; }
        const nuevoSpan = cubre ? ce.span - 1 : ce.span, nuevoCol = ce.col > c ? ce.col - 1 : ce.col;
        const el = new Map([['gridSpan', nuevoSpan > 1 ? `<${pf}gridSpan ${pf}val="${nuevoSpan}"/>` : null]]);
        el.set('tcW', `<${pf}tcW ${pf}w="${grid.slice(nuevoCol, nuevoCol + nuevoSpan).reduce((a, b) => a + b, 0)}" ${pf}type="dxa"/>`);
        ed.push(cambioTcPr(xml, ce, el, pf));
      }
    }
    return X().aplicar(xml, ed);
  }

  globalThis.TablaDoc = { nueva, analizar, insertarFila, eliminarFila, insertarColumna, eliminarColumna, anchoDe };
})();

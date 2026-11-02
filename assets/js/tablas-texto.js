/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Tablas en texto plano y en Markdown, para las herramientas que sacan el texto
// de un documento (PDF, Word, EPUB, PowerPoint → TXT o Markdown). Antes cada
// celda salía como un renglón suelto o pegada a la de al lado y la tabla se
// volvía ilegible.
//
//   TablasTexto.aTexto(filas)             → la tabla con las columnas alineadas
//   TablasTexto.aMarkdown(filas, escapar) → tabla de Markdown (la de GitHub, con "|")
//   TablasTexto.htmlATexto(raiz)          → el texto de un trozo de HTML (títulos,
//                                           párrafos, listas y tablas)
//   TablasTexto.diapositivaARenglones(xml) → los renglones de una diapositiva de PowerPoint,
//                                           con sus tablas
//
// `filas` es un arreglo de filas y cada fila un arreglo de textos; la primera
// fila es el encabezado. Una fila más corta que las demás se completa con celdas
// vacías. Una celda numérica va alineada a la derecha.
const TablasTexto = (() => {
  const NUMERO = /^[-−+(]?\s?[$€£¥₡]?\s?\d[\d.,'\s]*\s?(%|[A-Za-z]{1,3})?\)?$/;

  function normalizar(filas) {
    const columnas = Math.max(0, ...filas.map((f) => f.length));
    return filas.map((f) => Array.from({ length: columnas }, (_, i) => String(f[i] ?? '').replace(/\s*\n\s*/g, ' ').trim()));
  }

  // Una columna es numérica si todas sus celdas llenas (menos el encabezado) lo son.
  function columnasNumericas(filas) {
    return (filas[0] || []).map((_, c) => {
      const celdas = filas.slice(1).map((f) => f[c]).filter((x) => x !== '');
      return celdas.length > 0 && celdas.every((x) => NUMERO.test(x));
    });
  }

  const ancho = (s) => [...s].length;

  function aTexto(filasCrudas) {
    const filas = normalizar(filasCrudas);
    if (!filas.length || !filas[0].length) return '';
    const numericas = columnasNumericas(filas);
    const anchos = filas[0].map((_, c) => Math.max(1, ...filas.map((f) => ancho(f[c]))));
    const linea = (fila, esEncabezado) => fila.map((celda, c) => {
      const relleno = ' '.repeat(anchos[c] - ancho(celda));
      return numericas[c] && !esEncabezado ? relleno + celda : celda + relleno;
    }).join(' | ').replace(/\s+$/, '');
    const regla = anchos.map((a) => '-'.repeat(a)).join('-+-');
    return [linea(filas[0], true), regla, ...filas.slice(1).map((f) => linea(f, false))].join('\n');
  }

  const escaparCelda = (s) => s.replace(/\\/g, '\\\\').replace(/\|/g, '\\|');

  function aMarkdown(filasCrudas, escapar) {
    const filas = normalizar(filasCrudas);
    if (!filas.length || !filas[0].length) return '';
    const numericas = columnasNumericas(filas);
    const celda = (s) => (escapar ? escapar(s) : escaparCelda(s));
    const linea = (fila) => '| ' + fila.map(celda).join(' | ') + ' |';
    const regla = '| ' + numericas.map((n) => (n ? '---:' : '---')).join(' | ') + ' |';
    return [linea(filas[0]), regla, ...filas.slice(1).map(linea)].join('\n');
  }

  // --- HTML → texto ---
  const BLOQUES = new Set(['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'BLOCKQUOTE', 'SECTION', 'ARTICLE', 'PRE', 'HR',
    'FIGURE', 'FIGCAPTION', 'HEADER', 'FOOTER', 'ASIDE', 'NAV', 'MAIN', 'DL', 'DT', 'DD', 'ADDRESS']);
  const OMITIR = new Set(['SCRIPT', 'STYLE', 'HEAD', 'TITLE', 'NOSCRIPT', 'TEMPLATE']);
  const MARCA = '\u0001';

  function htmlATexto(raiz) {
    const tablas = [];
    let salida = '';
    let enItem = 0; // dentro de un <li>, los párrafos no dejan línea en blanco
    let finVineta = -1; // donde termina la viñeta del último <li>: ahí no va un salto
    const bloque = () => { if (salida.length !== finVineta) salida += '\n\n'; };
    const renglon = () => { if (salida.length !== finVineta && !salida.endsWith('\n')) salida += '\n'; };

    function recorrer(nodo, enPre) {
      if (nodo.nodeType === 3) {
        salida += enPre ? nodo.nodeValue : nodo.nodeValue.replace(/\s+/g, ' ');
        return;
      }
      if (nodo.nodeType !== 1 && nodo.nodeType !== 9 && nodo.nodeType !== 11) return;
      const etiqueta = nodo.nodeName;
      if (OMITIR.has(etiqueta)) return;
      if (etiqueta === 'BR') { salida += '\n'; return; }
      if (etiqueta === 'TABLE') {
        const filas = [...nodo.rows].map((fila) => [...fila.cells].flatMap((celda) => {
          const huecos = Math.max(1, parseInt(celda.getAttribute('colspan'), 10) || 1) - 1;
          return [textoDeCelda(celda), ...Array(Math.min(huecos, 50)).fill('')];
        })).filter((f) => f.length);
        bloque();
        salida += MARCA + (tablas.push(aTexto(filas)) - 1) + MARCA;
        bloque();
        return;
      }
      if (etiqueta === 'UL' || etiqueta === 'OL') {
        renglon();
        let n = parseInt(nodo.getAttribute('start'), 10) || 1;
        for (const hijo of nodo.children) {
          if (hijo.nodeName !== 'LI') { recorrer(hijo, enPre); continue; }
          renglon();
          salida += etiqueta === 'OL' ? (n++) + '. ' : '- ';
          finVineta = salida.length;
          enItem++;
          recorrer(hijo, enPre);
          enItem--;
        }
        bloque();
        return;
      }
      const esBloque = BLOQUES.has(etiqueta) || /^H[1-6]$/.test(etiqueta);
      if (esBloque) (enItem ? renglon() : bloque());
      if (etiqueta === 'HR') { salida += '-----'; bloque(); return; }
      for (const hijo of nodo.childNodes) recorrer(hijo, enPre || etiqueta === 'PRE');
      if (esBloque) (enItem ? renglon() : bloque());
    }

    // El texto de una celda: sus párrafos en una sola línea ("a / b" si hay una lista o varios párrafos).
    function textoDeCelda(celda) {
      const copia = celda.cloneNode(true);
      copia.querySelectorAll('table').forEach((anidada) => {
        anidada.replaceWith(copia.ownerDocument.createTextNode(' ' + [...anidada.rows]
          .map((f) => [...f.cells].map((c) => c.textContent.trim()).filter(Boolean).join(' / ')).filter(Boolean).join('; ') + ' '));
      });
      copia.querySelectorAll('br').forEach((br) => br.replaceWith(copia.ownerDocument.createTextNode(' ')));
      const sub = [];
      const hojas = copia.querySelectorAll('p, li, div');
      if (hojas.length > 1) {
        hojas.forEach((h) => { if (![...h.children].some((x) => /^(P|LI|DIV)$/.test(x.nodeName))) sub.push(h.textContent.replace(/\s+/g, ' ').trim()); });
        const lista = sub.filter(Boolean);
        if (lista.length) return lista.join(' / ');
      }
      return copia.textContent.replace(/\s+/g, ' ').trim();
    }

    recorrer(raiz, false);
    return salida
      .replace(/[ \t]*\n[ \t]*/g, '\n')
      .replace(/\n{3,}/g, '\n\n')
      .replace(new RegExp(MARCA + '(\\d+)' + MARCA, 'g'), (_, i) => tablas[+i])
      .trim();
  }

  // --- Diapositiva de PowerPoint (el XML de ppt/slides/slideN.xml ya leído) → renglones ---
  // Los párrafos de las formas, uno por renglón, y las tablas (a:tbl) con sus
  // columnas alineadas, en el orden en que aparecen en la diapositiva.
  function diapositivaARenglones(doc) {
    const renglones = [];
    const textoDe = (nodo) => [...nodo.getElementsByTagName('a:t')].map((t) => t.textContent).join('');
    function visitar(nodo) {
      if (nodo.nodeName === 'a:tbl') {
        const filas = [...nodo.getElementsByTagName('a:tr')].map((fila) => {
          const celdas = [];
          [...fila.childNodes].filter((c) => c.nodeName === 'a:tc').forEach((celda) => {
            celdas.push([...celda.getElementsByTagName('a:p')].map(textoDe).filter((x) => x.trim()).join(' '));
          });
          return celdas;
        }).filter((f) => f.length);
        if (filas.length && filas.some((f) => f.some((c) => c.trim()))) {
          if (renglones.length) renglones.push('');
          renglones.push(aTexto(filas));
          renglones.push('');
        }
        return;
      }
      if (nodo.nodeName === 'a:p') {
        const texto = textoDe(nodo);
        if (texto.trim() !== '') renglones.push(texto);
        return;
      }
      for (const hijo of nodo.childNodes) if (hijo.nodeType === 1) visitar(hijo);
    }
    visitar(doc.documentElement);
    while (renglones.length && renglones[renglones.length - 1] === '') renglones.pop();
    return renglones;
  }

  return { aTexto, aMarkdown, htmlATexto, diapositivaARenglones };
})();

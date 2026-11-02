/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// HTML → Markdown propio (reemplaza a Turndown y a turndown-tablas.js). Lo usan
// Word → Markdown, HTML → Markdown y EPUB → Markdown.
//
// Uso:  HtmlAMarkdown.convertir(htmlComoTexto)  →  texto Markdown
//
// Estilo de salida (el que ya usaba el sitio): títulos con `#`, listas con `*`
// (tres espacios tras la viñeta), énfasis con `_x_`, negrita con `**x**`,
// código cercado con ```, enlaces en línea, línea horizontal `* * *`, salto de
// línea con dos espacios.
//
// Trabaja sobre el DOM del navegador (DOMParser): primero colapsa los espacios
// como lo hace un navegador, luego recorre el árbol convirtiendo cada etiqueta.
//
// Tablas: Markdown no tiene celdas combinadas ni tablas anidadas. Una celda que
// abarca varias columnas deja las de al lado vacías, una que abarca varias
// filas queda solo en la primera, y una tabla dentro de una celda pasa a texto
// ("a / b; c / d"). Los saltos de línea de una celda pasan a <br>, que GitHub
// muestra como salto.

(function () {
  'use strict';

  const BLOQUE = new Set(['ADDRESS', 'ARTICLE', 'ASIDE', 'AUDIO', 'BLOCKQUOTE', 'BODY', 'CANVAS', 'CENTER', 'DD', 'DIR', 'DIV', 'DL', 'DT', 'FIELDSET', 'FIGCAPTION', 'FIGURE', 'FOOTER', 'FORM', 'FRAMESET', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'HEADER', 'HGROUP', 'HR', 'HTML', 'ISINDEX', 'LI', 'MAIN', 'MENU', 'NAV', 'NOFRAMES', 'NOSCRIPT', 'OL', 'OUTPUT', 'P', 'PRE', 'SECTION', 'TABLE', 'TBODY', 'TD', 'TFOOT', 'TH', 'THEAD', 'TR', 'UL']);
  const VACIO = new Set(['AREA', 'BASE', 'BR', 'COL', 'COMMAND', 'EMBED', 'HR', 'IMG', 'INPUT', 'KEYGEN', 'LINK', 'META', 'PARAM', 'SOURCE', 'TRACK', 'WBR']);
  const SIGNIFICATIVO_VACIO = new Set(['A', 'TABLE', 'THEAD', 'TBODY', 'TFOOT', 'TH', 'TD', 'IFRAME', 'SCRIPT', 'AUDIO', 'VIDEO']);

  const esBloque = (n) => BLOQUE.has(n.nodeName);
  const esVacio = (n) => VACIO.has(n.nodeName);
  const esPre = (n) => n.nodeName === 'PRE';
  const tieneVacio = (n) => !!n.querySelector && [...VACIO].some((t) => n.getElementsByTagName(t).length > 0);
  const tieneSignificativo = (n) => !!n.querySelector && [...SIGNIFICATIVO_VACIO].some((t) => n.getElementsByTagName(t).length > 0);

  // ───────── Espacios, como los colapsa un navegador ─────────
  function quitar(nodo) {
    const sig = nodo.nextSibling || nodo.parentNode;
    nodo.parentNode.removeChild(nodo);
    return sig;
  }
  function siguiente(prev, actual) {
    if ((prev && prev.parentNode === actual) || esPre(actual)) return actual.nextSibling || actual.parentNode;
    return actual.firstChild || actual.nextSibling || actual.parentNode;
  }
  function colapsarEspacios(raiz) {
    if (!raiz.firstChild || esPre(raiz)) return;
    let textoPrevio = null;
    let conservarEspacioInicial = false;
    let prev = null;
    let nodo = siguiente(prev, raiz);
    while (nodo !== raiz) {
      if (nodo.nodeType === 3 || nodo.nodeType === 4) {
        let texto = nodo.data.replace(/[ \r\n\t]+/g, ' ');
        if ((!textoPrevio || / $/.test(textoPrevio.data)) && !conservarEspacioInicial && texto[0] === ' ') texto = texto.substr(1);
        if (!texto) { nodo = quitar(nodo); continue; }
        nodo.data = texto;
        textoPrevio = nodo;
      } else if (nodo.nodeType === 1) {
        if (esBloque(nodo) || nodo.nodeName === 'BR') {
          if (textoPrevio) textoPrevio.data = textoPrevio.data.replace(/ $/, '');
          textoPrevio = null;
          conservarEspacioInicial = false;
        } else if (esVacio(nodo) || esPre(nodo)) {
          textoPrevio = null;
          conservarEspacioInicial = true;
        } else if (textoPrevio) {
          conservarEspacioInicial = false;
        }
      } else {
        nodo = quitar(nodo);
        continue;
      }
      const sig = siguiente(prev, nodo);
      prev = nodo;
      nodo = sig;
    }
    if (textoPrevio) {
      textoPrevio.data = textoPrevio.data.replace(/ $/, '');
      if (!textoPrevio.data) quitar(textoPrevio);
    }
  }

  // ───────── Información de cada nodo ─────────
  const esCodigo = (n) => (n.nodeType === 1 && n.nodeName === 'CODE') || (n.parentNode && esCodigo(n.parentNode));
  function esVacioDeContenido(n) {
    return !esVacio(n) && !SIGNIFICATIVO_VACIO.has(n.nodeName) && /^\s*$/i.test(n.textContent) && !tieneVacio(n) && !tieneSignificativo(n);
  }
  function bordes(cadena) {
    const m = /^(([ \t\r\n]*)(\s*))(?:(?=\S)[\s\S]*\S)?((\s*?)([ \t\r\n]*))$/.exec(cadena);
    return { inicio: m[1], inicioAscii: m[2], inicioOtro: m[3], fin: m[4], finOtro: m[5], finAscii: m[6] };
  }
  function flanqueadoPorEspacio(lado, nodo) {
    let hermano, re;
    if (lado === 'izq') { hermano = nodo.previousSibling; re = / $/; } else { hermano = nodo.nextSibling; re = /^ /; }
    if (hermano) {
      if (hermano.nodeType === 3) return re.test(hermano.nodeValue);
      if (hermano.nodeType === 1 && !esBloque(hermano)) return re.test(hermano.textContent);
    }
    return false;
  }
  function espaciosAlBorde(nodo) {
    if (esBloque(nodo)) return { inicio: '', fin: '' };
    const e = bordes(nodo.textContent);
    if (e.inicioAscii && flanqueadoPorEspacio('izq', nodo)) e.inicio = e.inicioOtro;
    if (e.finAscii && flanqueadoPorEspacio('der', nodo)) e.fin = e.finOtro;
    return { inicio: e.inicio, fin: e.fin };
  }

  // ───────── Texto ─────────
  const ESCAPES = [
    [/\\/g, '\\\\'], [/\*/g, '\\*'], [/^-/g, '\\-'], [/^\+ /g, '\\+ '], [/^(=+)/g, '\\$1'], [/^(#{1,6}) /g, '\\$1 '],
    [/`/g, '\\`'], [/^~~~/g, '\\~~~'], [/\[/g, '\\['], [/\]/g, '\\]'], [/^>/g, '\\>'], [/_/g, '\\_'], [/^(\d+)\. /g, '$1\\. '],
  ];
  const escapar = (s) => ESCAPES.reduce((r, [re, por]) => r.replace(re, por), s);
  const limpiarAtributo = (a) => (a ? a.replace(/(\n+\s*)+/g, '\n') : '');

  const quitarSaltosFinales = (s) => s.replace(/\n+$/, '');
  const quitarSaltosIniciales = (s) => s.replace(/^\n*/, '');
  function unir(salida, reemplazo) {
    const s1 = quitarSaltosFinales(salida);
    const s2 = quitarSaltosIniciales(reemplazo);
    const n = Math.max(salida.length - s1.length, reemplazo.length - s2.length);
    return s1 + '\n\n'.substring(0, n) + s2;
  }

  // ───────── Reglas ─────────
  function celdaAMarkdown(celda) {
    const copia = celda.cloneNode(true);
    copia.querySelectorAll('table').forEach((anidada) => {
      const texto = [...anidada.rows]
        .map((fila) => [...fila.cells].map((c) => c.textContent.trim()).filter(Boolean).join(' / '))
        .filter(Boolean).join('; ');
      anidada.replaceWith(copia.ownerDocument.createTextNode(' ' + texto + ' '));
    });
    return convertir(copia.innerHTML)
      .trim()
      .replace(/\n+/g, '<br>')
      .replace(/\|/g, '\\|');
  }

  // Cada regla: filtro(nodo) y reemplazo(contenido, nodo). La primera que coincide gana.
  const REGLAS = [
    { // tablas
      filtro: (n) => n.nodeName === 'TABLE',
      reemplazo(_c, tabla) {
        const filas = [...tabla.rows].map((fila) => {
          const celdas = [];
          for (const celda of fila.cells) {
            celdas.push(celdaAMarkdown(celda));
            for (let i = 1; i < (parseInt(celda.getAttribute('colspan'), 10) || 1); i++) celdas.push('');
          }
          return celdas;
        }).filter((celdas) => celdas.length);
        if (!filas.length) return '';
        const columnas = Math.max(...filas.map((f) => f.length));
        const linea = (celdas) => '| ' + Array.from({ length: columnas }, (_, i) => celdas[i] || '').join(' | ') + ' |';
        const lineas = [linea(filas[0]), linea(Array(columnas).fill('---')), ...filas.slice(1).map(linea)];
        return '\n\n' + lineas.join('\n') + '\n\n';
      },
    },
    { filtro: (n) => n.nodeName === 'P', reemplazo: (c) => '\n\n' + c + '\n\n' },
    { filtro: (n) => n.nodeName === 'BR', reemplazo: () => '  \n' },
    { filtro: (n) => /^H[1-6]$/.test(n.nodeName), reemplazo: (c, n) => '\n\n' + '#'.repeat(Number(n.nodeName.charAt(1))) + ' ' + c + '\n\n' },
    { filtro: (n) => n.nodeName === 'BLOCKQUOTE', reemplazo: (c) => '\n\n' + c.replace(/^\n+|\n+$/g, '').replace(/^/gm, '> ') + '\n\n' },
    {
      filtro: (n) => n.nodeName === 'UL' || n.nodeName === 'OL',
      reemplazo(c, n) {
        const padre = n.parentNode;
        if (padre.nodeName === 'LI' && padre.lastElementChild === n) return '\n' + c;
        return '\n\n' + c + '\n\n';
      },
    },
    {
      filtro: (n) => n.nodeName === 'LI',
      reemplazo(c, n) {
        c = c.replace(/^\n+/, '').replace(/\n+$/, '\n').replace(/\n/gm, '\n    ');
        let prefijo = '*   ';
        const padre = n.parentNode;
        if (padre.nodeName === 'OL') {
          const inicio = padre.getAttribute('start');
          const indice = Array.prototype.indexOf.call(padre.children, n);
          prefijo = (inicio ? Number(inicio) + indice : indice + 1) + '.  ';
        }
        return prefijo + c + (n.nextSibling && !/\n$/.test(c) ? '\n' : '');
      },
    },
    {
      filtro: (n) => n.nodeName === 'PRE' && n.firstChild && n.firstChild.nodeName === 'CODE',
      reemplazo(_c, n) {
        const clase = n.firstChild.getAttribute('class') || '';
        const lenguaje = (clase.match(/language-(\S+)/) || [null, ''])[1];
        const codigo = n.firstChild.textContent;
        let largo = 3;
        const re = /^`{3,}/gm;
        let m;
        while ((m = re.exec(codigo))) if (m[0].length >= largo) largo = m[0].length + 1;
        const cerca = '`'.repeat(largo);
        return '\n\n' + cerca + lenguaje + '\n' + codigo.replace(/\n$/, '') + '\n' + cerca + '\n\n';
      },
    },
    { filtro: (n) => n.nodeName === 'HR', reemplazo: () => '\n\n* * *\n\n' },
    {
      filtro: (n) => n.nodeName === 'A' && n.getAttribute('href'),
      reemplazo(c, n) {
        let href = n.getAttribute('href');
        if (href) href = href.replace(/([()])/g, '\\$1');
        let titulo = limpiarAtributo(n.getAttribute('title'));
        if (titulo) titulo = ' "' + titulo.replace(/"/g, '\\"') + '"';
        return '[' + c + '](' + href + titulo + ')';
      },
    },
    { filtro: (n) => n.nodeName === 'EM' || n.nodeName === 'I', reemplazo: (c) => (c.trim() ? '_' + c + '_' : '') },
    { filtro: (n) => n.nodeName === 'STRONG' || n.nodeName === 'B', reemplazo: (c) => (c.trim() ? '**' + c + '**' : '') },
    {
      filtro(n) {
        const conHermanos = n.previousSibling || n.nextSibling;
        const esBloqueDeCodigo = n.parentNode.nodeName === 'PRE' && !conHermanos;
        return n.nodeName === 'CODE' && !esBloqueDeCodigo;
      },
      reemplazo(c) {
        if (!c) return '';
        c = c.replace(/\r?\n|\r/g, ' ');
        const espacio = /^`|^ .*?[^ ].* $|`$/.test(c) ? ' ' : '';
        let delimitador = '`';
        const coincidencias = c.match(/`+/gm) || [];
        while (coincidencias.indexOf(delimitador) !== -1) delimitador += '`';
        return delimitador + espacio + c + espacio + delimitador;
      },
    },
    {
      filtro: (n) => n.nodeName === 'IMG',
      reemplazo(_c, n) {
        const alt = limpiarAtributo(n.getAttribute('alt'));
        const src = n.getAttribute('src') || '';
        const titulo = limpiarAtributo(n.getAttribute('title'));
        return src ? '![' + alt + '](' + src + (titulo ? ' "' + titulo + '"' : '') + ')' : '';
      },
    },
  ];
  const REGLA_VACIA = { reemplazo: (_c, n) => (esBloque(n) ? '\n\n' : '') };
  const REGLA_POR_DEFECTO = { reemplazo: (c, n) => (esBloque(n) ? '\n\n' + c + '\n\n' : c) };

  function reglaPara(nodo) {
    if (esVacioDeContenido(nodo)) return REGLA_VACIA;
    return REGLAS.find((r) => r.filtro(nodo)) || REGLA_POR_DEFECTO;
  }

  function procesar(padre) {
    let salida = '';
    for (const nodo of padre.childNodes) {
      let reemplazo = '';
      if (nodo.nodeType === 3) reemplazo = esCodigo(nodo) ? nodo.nodeValue : escapar(nodo.nodeValue);
      else if (nodo.nodeType === 1) {
        const regla = reglaPara(nodo);
        let contenido = procesar(nodo);
        const espacios = espaciosAlBorde(nodo);
        if (espacios.inicio || espacios.fin) contenido = contenido.trim();
        reemplazo = espacios.inicio + regla.reemplazo(contenido, nodo) + espacios.fin;
      }
      salida = unir(salida, reemplazo);
    }
    return salida;
  }

  function convertir(html) {
    if (html === '' || html == null) return '';
    const doc = new DOMParser().parseFromString('<x-conversor id="raiz-conversor">' + html + '</x-conversor>', 'text/html');
    const raiz = doc.getElementById('raiz-conversor');
    colapsarEspacios(raiz);
    return procesar(raiz).replace(/^[\t\r\n]+/, '').replace(/[\t\r\n\s]+$/, '');
  }

  globalThis.HtmlAMarkdown = { convertir };
})();

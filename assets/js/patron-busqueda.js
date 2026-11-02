/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Patrón de "buscar y reemplazar": del texto que escribe la persona y tres opciones (expresión regular, palabra completa, mayúsculas)
// a una RegExp lista para usar, y la expansión del texto de reemplazo ($1, $&, $<nombre>). Lo comparten Buscar y reemplazar en Word,
// en Excel y en texto, para que las tres se comporten igual.
//
//   PatronBusqueda.crear('ab.', { regex: false, palabra: true, mayusculas: false })  → { regex } | { error }
//   PatronBusqueda.expandir('$2-$1', coincidencia, true)                              → texto del reemplazo para esa coincidencia
//   PatronBusqueda.reemplazarEn(texto, regex, reemplazo, esRegex)                     → { texto, cantidad }
//
// Siempre con la bandera "u" (Unicode): así una letra con tilde o un emoji cuentan como una sola cosa y \p{L} funciona.
// "Palabra completa" no usa \b (que no entiende las tildes): exige que antes y después no haya letra, número ni guion bajo.
(function (global) {
  var LARGO_MAXIMO = 500;
  var BORDE = '[\\p{L}\\p{N}_]';

  function escapar(texto) {
    return texto.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  // Devuelve { regex } o { error: 'largo' | 'invalida', detalle }. Un patrón vacío no es un error: no hay nada que buscar (regex: null).
  function crear(texto, opciones) {
    opciones = opciones || {};
    if (!texto) return { regex: null };
    if (texto.length > LARGO_MAXIMO) return { error: 'largo' };
    var cuerpo = opciones.regex ? texto : escapar(texto);
    if (opciones.palabra) cuerpo = '(?<!' + BORDE + ')(?:' + cuerpo + ')(?!' + BORDE + ')';
    try {
      return { regex: new RegExp(cuerpo, opciones.mayusculas ? 'gu' : 'giu') };
    } catch (error) {
      return { error: 'invalida', detalle: error.message };
    }
  }

  // El texto de reemplazo para una coincidencia (un resultado de matchAll, o los argumentos de un replace). Sin expresión regular es
  // literal: un "$1" escrito se queda como "$1". Con ella, igual que en JavaScript: $$ → $, $& → todo lo encontrado, $1…$99 → ese
  // grupo, $<nombre> → grupo con nombre. Un grupo que no existe se deja tal cual escrito.
  function expandir(reemplazo, coincidencia, esRegex) {
    if (!esRegex || reemplazo.indexOf('$') < 0) return reemplazo;
    var grupos = coincidencia.groups || null;
    return reemplazo.replace(/\$(\$|&|<([^>]*)>|\d{1,2})/g, function (todo, que, nombre) {
      if (que === '$') return '$';
      if (que === '&') return coincidencia[0];
      if (nombre !== undefined) return grupos && nombre in grupos ? (grupos[nombre] === undefined ? '' : grupos[nombre]) : todo;
      var n = Number(que);
      if (n >= 1 && n < coincidencia.length) return coincidencia[n] === undefined ? '' : coincidencia[n];
      // "$12" cuando solo hay 5 grupos: se prueba como "$1" seguido de un "2" (así lo hace JavaScript)
      if (que.length === 2 && Number(que[0]) >= 1 && Number(que[0]) < coincidencia.length) {
        var primero = coincidencia[Number(que[0])];
        return (primero === undefined ? '' : primero) + que[1];
      }
      return todo;
    });
  }

  // Reemplaza todas las coincidencias de un texto. Una coincidencia vacía (por ejemplo, "a*" en "bbb") no se cuenta ni se reemplaza.
  function reemplazarEn(texto, regex, reemplazo, esRegex) {
    var cantidad = 0;
    var nuevo = texto.replace(regex, function () {
      var args = Array.prototype.slice.call(arguments);
      var ultimo = args[args.length - 1];
      var grupos = ultimo !== null && typeof ultimo === 'object' ? ultimo : undefined;
      var m = args.slice(0, grupos ? -3 : -2);
      if (grupos) m.groups = grupos;
      if (m[0].length === 0) return m[0];
      cantidad++;
      return expandir(reemplazo, m, esRegex);
    });
    return { texto: nuevo, cantidad: cantidad };
  }

  global.PatronBusqueda = { crear: crear, expandir: expandir, reemplazarEn: reemplazarEn, LARGO_MAXIMO: LARGO_MAXIMO };
})(typeof window !== 'undefined' ? window : this);

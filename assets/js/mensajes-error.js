/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Errores de los motores en el idioma de la página (plan: Investigaciones/investigacion-mensajes-error/00-plan.md).
//
// Un motor no conoce el idioma de la página (en un hilo de Flujos ni siquiera hay página): dice QUÉ pasó con un código y sus datos, y la
// página lo cuenta en su idioma.
//
//   throw errorMC('lut-fila', { linea: 12 }, 'Fila no válida en la línea {linea}')
//        → un Error con .codigo y .datos; su .message es el texto en español ya rellenado (igual que antes: consolas, pruebas y la página
//          en español no cambian).
//   detalleError(error)  (idioma.js) → si MensajesError tiene 'error-<codigo>' en el idioma de la página, ese texto con los datos; si no,
//          error.message como siempre.
//
// Reglas (las comprueba el build, mensajes-error.mjs):
//   - Cada código usado en errorMC('…') tiene su texto en «es» y en «en» acá, y el texto en español de la llamada es el mismo de acá.
//   - Una página que carga un script que usa errorMC carga antes este archivo (y un hilo lo importa antes que sus piezas).
//   - Los códigos llevan delante el motor («lut-», «word-», «pdf-»…): no chocan con los .codigo cortos que ya leen algunas páginas
//     («rango», «formato»…).
// Huecos: {nombre} se reemplaza por datos.nombre (con una función: un «$&» del usuario no se interpreta).
(function (g) {
  'use strict';

  function rellenar(texto, datos) {
    return String(texto).replace(/\{(\w+)\}/g, function (todo, clave) {
      return datos && datos[clave] !== undefined && datos[clave] !== null ? String(datos[clave]) : todo;
    });
  }

  function errorMC(codigo, datos, textoEs) {
    var error = new Error(rellenar(textoEs || codigo, datos));
    error.codigo = codigo;
    error.datos = datos || {};
    return error;
  }

  g.MensajesError = {
    rellenar: rellenar,
    es: {
      // Tablas de color .cube (lut-cube.js; Editor de color)
      'error-lut-dominio': 'DOMAIN inválido en la línea {linea}',
      'error-lut-dominio-orden': 'DOMAIN_MAX debe ser mayor que DOMAIN_MIN',
      'error-lut-dos-tablas': 'El archivo declara dos tablas',
      'error-lut-tamano': 'Tamaño de tabla no válido: {tamano}',
      'error-lut-fila': 'Fila no válida en la línea {linea}',
      'error-lut-colores-antes': 'Hay colores antes de LUT_3D_SIZE o LUT_1D_SIZE',
      'error-lut-sobran-filas': 'Sobran filas respecto del tamaño declarado',
      'error-lut-no-es-cube': 'No es un archivo .cube (falta LUT_3D_SIZE o LUT_1D_SIZE)',
      'error-lut-faltan-filas': 'Faltan filas: hay {hay} de {esperado}',
    },
    en: {
      'error-lut-dominio': 'Invalid DOMAIN on line {linea}',
      'error-lut-dominio-orden': 'DOMAIN_MAX must be greater than DOMAIN_MIN',
      'error-lut-dos-tablas': 'The file declares two tables',
      'error-lut-tamano': 'Invalid table size: {tamano}',
      'error-lut-fila': 'Invalid row on line {linea}',
      'error-lut-colores-antes': 'There are colors before LUT_3D_SIZE or LUT_1D_SIZE',
      'error-lut-sobran-filas': 'There are more rows than the declared size',
      'error-lut-no-es-cube': 'Not a .cube file (LUT_3D_SIZE or LUT_1D_SIZE is missing)',
      'error-lut-faltan-filas': 'Rows are missing: there are {hay} of {esperado}',
    },
  };
  g.errorMC = errorMC;
})(typeof self !== 'undefined' ? self : this);

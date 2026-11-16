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
      // Word, OpenDocument texto y RTF (doc-lector.js, odt-lector.js, docx-a-contrato.js, escritores…)
      'error-xml-demasiados-elementos': 'El documento tiene demasiados elementos y no se puede leer.',
      'error-xml-demasiados-niveles': 'El documento tiene demasiados niveles anidados y no se puede leer.',
      'error-word-demasiado-grande': 'El documento es demasiado grande para abrirlo aquí.',
      'error-word-demasiados-parrafos': 'El documento tiene demasiados párrafos y no se puede leer.',
      'error-word-demasiados-parrafos-aqui': 'El documento tiene demasiados párrafos para abrirlo aquí.',
      'error-word-demasiadas-imagenes': 'El documento tiene demasiadas imágenes y no se puede leer.',
      'error-word-demasiadas-imagenes-aqui': 'El documento tiene demasiadas imágenes para abrirlo aquí.',
      'error-word-imagenes-ocupan': 'Las imágenes del documento ocupan demasiado y no se puede leer.',
      'error-word-sin-cuerpo': 'El documento no tiene cuerpo.',
      'error-word-sin-document-xml': 'No es un documento de Word válido (falta word/document.xml).',
      'error-word-sin-documento-legible': 'El archivo no tiene un documento de Word legible.',
      'error-word-no-es-docx': 'El archivo no es un documento de Word (.docx).',
      'error-word-no-legible': 'el archivo no es un documento de Word que se pueda leer (.docx, .doc u .odt)',
      'error-word-lector-no-carga': 'No se pudo cargar el lector de este tipo de documento ({archivo}). Revisá tu conexión y volvé a intentarlo.',
      'error-doc-cfb-danado': 'El archivo está dañado y no se puede abrir ({detalle}).',
      'error-doc-no-valido': 'El archivo no es un documento de Word 97-2003 válido.',
      'error-doc-cifrado': 'El documento está protegido con contraseña. Sacale la contraseña en Word y volvé a intentarlo.',
      'error-doc-no-es-doc': 'El archivo no es un documento de Word (.doc).',
      'error-doc-antiguo': 'Este documento es de una versión de Word anterior a Word 97 y no se puede leer. Abrilo en Word y guardalo como .docx.',
      'error-doc-sin-tabla': 'El documento está dañado: falta la tabla de estructura.',
      'error-doc-sin-piezas': 'El documento está dañado: no tiene tabla de piezas de texto.',
      'error-doc-piezas-invalidas': 'El documento está dañado: tabla de piezas inválida.',
      'error-doc-piezas-desordenadas': 'El documento está dañado: piezas de texto desordenadas.',
      'error-doc-piezas-huecos': 'El documento está dañado: piezas de texto con huecos.',
      'error-doc-demasiados-formatos': 'El documento está dañado: demasiados formatos.',
      'error-doc-demasiado-grande-guardar': 'El documento es demasiado grande para guardarlo como .doc.',
      'error-doc-parrafo-propiedades': 'Un párrafo tiene demasiadas propiedades para guardarlo como .doc.',
      'error-odt-no-es-odt': 'El archivo no es un documento de texto de OpenDocument (.odt).',
      'error-odt-paquete-invalido': 'El archivo no es un .odt válido (no se puede abrir como un paquete).',
      'error-odt-cifrado': 'El documento está protegido con contraseña. Sacale la contraseña en LibreOffice y volvé a intentarlo.',
      'error-odt-sin-content': 'El archivo no es un documento de OpenDocument (.odt).',
      'error-odt-tabla-filas': 'La tabla tiene demasiadas filas y no se puede leer.',
      'error-odt-no-es-texto': 'No es un documento de texto de OpenDocument (.odt).',
      'error-odt-cifrado-editor': 'El documento está protegido con contraseña; ábrelo en LibreOffice y guárdalo sin contraseña.',
      'error-odt-sin-content-xml': 'No es un documento de OpenDocument válido (falta content.xml).',
      'error-odt-demasiado-grande-guardar': 'El documento es demasiado grande para guardarlo como .odt.',
      'error-rtf-demasiado-grande-guardar': 'El documento es demasiado grande para guardarlo como .rtf.',
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
      // Word, OpenDocument text and RTF
      'error-xml-demasiados-elementos': 'The document has too many elements to be read.',
      'error-xml-demasiados-niveles': 'The document has too many nested levels to be read.',
      'error-word-demasiado-grande': 'The document is too large to open here.',
      'error-word-demasiados-parrafos': 'The document has too many paragraphs to be read.',
      'error-word-demasiados-parrafos-aqui': 'The document has too many paragraphs to open here.',
      'error-word-demasiadas-imagenes': 'The document has too many images to be read.',
      'error-word-demasiadas-imagenes-aqui': 'The document has too many images to open here.',
      'error-word-imagenes-ocupan': 'The images in the document take up too much space to be read.',
      'error-word-sin-cuerpo': 'The document has no body.',
      'error-word-sin-document-xml': 'Not a valid Word document (word/document.xml is missing).',
      'error-word-sin-documento-legible': 'The file does not contain a readable Word document.',
      'error-word-no-es-docx': 'The file is not a Word document (.docx).',
      'error-word-no-legible': 'the file is not a Word document that can be read (.docx, .doc or .odt)',
      'error-word-lector-no-carga': 'Could not load the reader for this type of document ({archivo}). Check your connection and try again.',
      'error-doc-cfb-danado': 'The file is damaged and cannot be opened.',
      'error-doc-no-valido': 'The file is not a valid Word 97-2003 document.',
      'error-doc-cifrado': 'The document is password-protected. Remove the password in Word and try again.',
      'error-doc-no-es-doc': 'The file is not a Word document (.doc).',
      'error-doc-antiguo': 'This document is from a version of Word older than Word 97 and cannot be read. Open it in Word and save it as .docx.',
      'error-doc-sin-tabla': 'The document is damaged: the structure table is missing.',
      'error-doc-sin-piezas': 'The document is damaged: it has no text pieces table.',
      'error-doc-piezas-invalidas': 'The document is damaged: invalid text pieces table.',
      'error-doc-piezas-desordenadas': 'The document is damaged: text pieces are out of order.',
      'error-doc-piezas-huecos': 'The document is damaged: text pieces have gaps.',
      'error-doc-demasiados-formatos': 'The document is damaged: too many formats.',
      'error-doc-demasiado-grande-guardar': 'The document is too large to save as .doc.',
      'error-doc-parrafo-propiedades': 'A paragraph has too many properties to save as .doc.',
      'error-odt-no-es-odt': 'The file is not an OpenDocument text document (.odt).',
      'error-odt-paquete-invalido': 'The file is not a valid .odt (it cannot be opened as a package).',
      'error-odt-cifrado': 'The document is password-protected. Remove the password in LibreOffice and try again.',
      'error-odt-sin-content': 'The file is not an OpenDocument document (.odt).',
      'error-odt-tabla-filas': 'The table has too many rows to be read.',
      'error-odt-no-es-texto': 'Not an OpenDocument text document (.odt).',
      'error-odt-cifrado-editor': 'The document is password-protected; open it in LibreOffice and save it without a password.',
      'error-odt-sin-content-xml': 'Not a valid OpenDocument document (content.xml is missing).',
      'error-odt-demasiado-grande-guardar': 'The document is too large to save as .odt.',
      'error-rtf-demasiado-grande-guardar': 'The document is too large to save as .rtf.',
    },
  };
  g.errorMC = errorMC;
})(typeof self !== 'undefined' ? self : this);

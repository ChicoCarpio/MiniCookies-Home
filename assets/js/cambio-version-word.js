/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Cambio de versión de documentos de Word: pasa un documento de cualquier formato que el sitio sabe leer (.docx, .doc de Word 97-2003, .odt de LibreOffice) a .docx, .doc, .odt o .rtf,
// conservando el texto, los estilos, las listas, las tablas, las imágenes, los enlaces, los encabezados y pies y las secciones. Es el gemelo de CambioVersion (Excel): los tres lectores
// entregan el mismo «contrato de documento» y los cuatro escritores lo reciben, así que cualquier origen sirve a cualquier destino.
//
//   await CambioVersionWord.detectar(bytes)                  → 'docx' | 'doc' | 'odt' | null
//   await CambioVersionWord.convertir(bytes, 'docx' | 'doc' | 'odt' | 'rtf')
//        → { bytes: Uint8Array, origen, avisos: [{ tipo, cantidad }] }      (lanza Error con .codigo = 'mismo-formato' si el origen ya está en el destino)
//
// Los avisos son lo que no se pudo llevar: tipo 'notas', 'dibujos', 'imagen-formato', 'imagen-falta', 'imagen-externa', 'enlace-descartado'.
// Requiere formatos-word.js (con sus lectores y escritores, que se piden solos).
(function () {
  'use strict';
  const DESTINOS = ['docx', 'doc', 'odt', 'rtf'];

  const detectar = (bytes) => FormatosWord.detectar(bytes);

  async function convertir(bytes, destino) {
    if (!DESTINOS.includes(destino)) throw new Error('formato de destino desconocido: ' + destino);
    const mismo = () => { const e = new Error('el documento ya está en ese formato'); e.codigo = 'mismo-formato'; return e; };
    if ((await detectar(bytes)) === destino) throw mismo();   // antes de leer todo el documento
    const { contrato, origen, mensajes } = await FormatosWord.aContrato(bytes);
    if (origen === destino) throw mismo();
    const r = await FormatosWord.escribir(contrato, destino);
    const tipos = [...new Set([...mensajes, ...r.mensajes])];
    return { bytes: r.bytes, origen, avisos: tipos.map((tipo) => ({ tipo, cantidad: 1 })) };
  }

  globalThis.CambioVersionWord = { detectar, convertir, DESTINOS };
})();

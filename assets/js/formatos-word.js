/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Los formatos de documento de Word que el sitio lee: .docx, .doc (Word 97-2003) y .odt (LibreOffice / OpenOffice). Es el gemelo de formatos-antiguos.js (hojas de cálculo):
// las herramientas de Word saben leer .docx, así que lo que no es .docx se pasa a .docx antes y el resto de la herramienta no cambia.
//
//   FormatosWord.esDocumento(nombre)           → true si el nombre acaba en .docx, .doc u .odt
//   FormatosWord.sinExtension(nombre)          → el nombre sin .docx, .doc ni .odt
//   await FormatosWord.detectar(bytes)         → 'docx' | 'doc' | 'odt' | null    (por el contenido, no por la extensión)
//   await FormatosWord.aDocx(bytes)            → { bytes: Uint8Array, origen, mensajes: [clave…] }    (un .docx se devuelve tal cual)
//   await FormatosWord.leerArchivo(archivo)    → { buffer: ArrayBuffer con un .docx, bytes, origen, mensajes }   (archivo: File o Blob)
//   await FormatosWord.comoDocx(datos)         → datos (ArrayBuffer, Uint8Array o Blob) listos para leer como .docx; DocxBase.abrirPaquete lo llama solo, así que
//                                                DocxLector, DocxMaquetador y todo lo que pasa por ellos acepta .doc y .odt sin más. Un .docx no se toca (ni se examina).
//
// Lanza Error (con .codigo = 'cifrado' si el documento tiene contraseña) con un mensaje listo para mostrar.
// mensajes: claves de lo que no se pudo llevar al .docx: 'notas', 'dibujos', 'imagen-formato', 'imagen-falta', 'imagen-externa'.
//
// Los lectores de .doc y .odt (cfb-propio.js, doc-lector.js, odt-lector.js, docx-desde-contrato.js, xml-ligero.js) se piden la primera vez que llega un archivo
// que no es .docx: una página que solo recibe .docx no los baja. Se buscan al lado de este archivo, en assets/js/ (como archivo-a-pdf.js).
// En Node (las pruebas) se cargan antes con vm y este módulo no pide nada. Requiere zip-propio.js.
(function () {
  'use strict';
  const EXTENSIONES = /\.(docx|doc|odt)$/i;
  const base = typeof document !== 'undefined' && document.currentScript && document.currentScript.src ? new URL('./', document.currentScript.src).href : null;

  const esDocumento = (nombre) => EXTENSIONES.test(String(nombre || ''));
  const sinExtension = (nombre) => String(nombre || '').replace(EXTENSIONES, '');

  // Cada módulo se pide una sola vez; si ya está en la página (otro cargador o una etiqueta <script>) no se repite.
  const cargas = new Map();
  function cargar(global, archivo) {
    if (typeof globalThis[global] !== 'undefined') return Promise.resolve();
    if (!base) return Promise.reject(new Error('Falta el módulo ' + archivo + ' para leer este tipo de documento.'));
    const src = base + archivo;
    if (!cargas.has(src)) {
      cargas.set(src, new Promise((resolver, rechazar) => {
        const s = document.createElement('script');
        s.src = src;
        s.onload = resolver;
        s.onerror = () => { cargas.delete(src); rechazar(new Error('No se pudo cargar el lector de este tipo de documento (' + archivo + '). Revisá tu conexión y volvé a intentarlo.')); };
        document.head.appendChild(s);
      }));
    }
    return cargas.get(src);
  }
  async function prepararDoc() {
    await cargar('ZipPropio', 'zip-propio.js');
    await cargar('CfbPropio', 'cfb-propio.js');
    await cargar('DocxDesdeContrato', 'docx-desde-contrato.js');
    await cargar('DocLector', 'doc-lector.js');
  }
  async function prepararOdt() {
    await cargar('ZipPropio', 'zip-propio.js');
    await cargar('XmlLigero', 'xml-ligero.js');
    await cargar('DocxDesdeContrato', 'docx-desde-contrato.js');
    await cargar('OdtLector', 'odt-lector.js');
  }

  const CFB = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
  const esCfb = (b) => b.length >= 8 && CFB.every((x, i) => b[i] === x);
  // Un .odt guarda «mimetype» como primera entrada del zip, sin comprimir: su nombre está en el byte 30. Un .docx empieza con [Content_Types].xml (o con otra cosa).
  const MIMETYPE = [0x6D, 0x69, 0x6D, 0x65, 0x74, 0x79, 0x70, 0x65];
  const pareceOdt = (b) => b.length > 40 && b[0] === 0x50 && b[1] === 0x4B && MIMETYPE.every((x, i) => b[30 + i] === x);

  async function detectar(bytes) {
    if (esCfb(bytes)) { await prepararDoc(); return DocLector.parece(bytes) ? 'doc' : null; }
    if (bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4B) {
      await cargar('ZipPropio', 'zip-propio.js');
      let zip;
      try { zip = await ZipPropio.loadAsync(bytes); } catch (e) { return null; }
      if (zip.file('word/document.xml')) return 'docx';
      if (zip.file('content.xml') && (zip.file('mimetype') || zip.file('META-INF/manifest.xml'))) { await prepararOdt(); return (await OdtLector.parece(bytes)) ? 'odt' : null; }
      if (zip.file('encrypted-package')) return 'docx';   // el lector de .docx avisa que tiene contraseña
    }
    return null;
  }

  async function aDocx(bytes) {
    const origen = await detectar(bytes);
    if (origen === 'docx') return { bytes, origen, mensajes: [] };
    if (origen === 'doc') { const r = await DocLector.aDocx(bytes); return { bytes: r.bytes, origen, mensajes: r.mensajes }; }
    if (origen === 'odt') { const r = await OdtLector.aDocx(bytes); return { bytes: r.bytes, origen, mensajes: r.mensajes }; }
    if (esCfb(bytes)) await DocLector.aDocx(bytes);   // otro contenedor de Office (un .docx con contraseña, un .xls renombrado): el lector dice el motivo
    throw new Error('el archivo no es un documento de Word que se pueda leer (.docx, .doc u .odt)');
  }

  const aBytes = async (datos) => (typeof Blob !== 'undefined' && datos instanceof Blob ? new Uint8Array(await datos.arrayBuffer()) : datos instanceof Uint8Array ? datos : new Uint8Array(datos));

  async function comoDocx(datos) {
    const b = await aBytes(datos);
    if (esCfb(b)) { await prepararDoc(); return (await DocLector.aDocx(b)).bytes; }
    if (pareceOdt(b)) { await prepararOdt(); if (await OdtLector.parece(b)) return (await OdtLector.aDocx(b)).bytes; }
    return datos;
  }

  async function leerArchivo(archivo) {
    const bytes = new Uint8Array(await archivo.arrayBuffer());
    const r = await aDocx(bytes);
    const buffer = r.bytes.byteOffset === 0 && r.bytes.byteLength === r.bytes.buffer.byteLength ? r.bytes.buffer : r.bytes.slice().buffer;
    return { buffer, bytes: r.bytes, origen: r.origen, mensajes: r.mensajes };
  }

  globalThis.FormatosWord = { esDocumento, sinExtension, detectar, aDocx, comoDocx, leerArchivo };
})();

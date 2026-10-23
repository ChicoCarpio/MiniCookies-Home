/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Los formatos de presentación de PowerPoint que el sitio lee: .pptx, .ppt (PowerPoint 97-2003) y .odp (LibreOffice / OpenOffice). Es el gemelo de formatos-word.js:
// las herramientas de PowerPoint saben leer .pptx, así que lo que no es .pptx se pasa a .pptx antes y el resto de la herramienta no cambia.
//
//   FormatosPowerpoint.esPresentacion(nombre)         → true si el nombre acaba en .pptx, .ppt u .odp
//   FormatosPowerpoint.sinExtension(nombre)           → el nombre sin .pptx, .ppt ni .odp
//   await FormatosPowerpoint.detectar(bytes)          → 'pptx' | 'ppt' | 'odp' | null    (por el contenido, no por la extensión)
//   await FormatosPowerpoint.aPptx(bytes)             → { bytes: Uint8Array, origen, mensajes: [clave…] }    (un .pptx se devuelve tal cual)
//   await FormatosPowerpoint.leerArchivo(archivo)     → { buffer: ArrayBuffer con un .pptx, bytes, origen, mensajes }   (archivo: File o Blob)
//   await FormatosPowerpoint.comoPptx(datos)          → datos (ArrayBuffer, Uint8Array o Blob) listos para leer como .pptx; PptxModelo.abrir lo llama solo, así que
//                                                       PowerPoint → PDF, → imágenes y lo que pasa por el motor acepta .ppt y .odp sin más. Un .pptx no se toca (ni se examina).
//
// Lanza Error (con .codigo = 'cifrado' si la presentación tiene contraseña) con un mensaje listo para mostrar.
// mensajes: claves de lo que no se pudo llevar al .pptx: 'formas', 'objetos', 'wordart', 'imagen-formato', 'imagen-falta', 'imagen-externa', 'enlace-descartado', 'grupos-anidados'.
//
// Los lectores de .ppt y .odp (cfb-propio.js, ppt-lector.js, odp-lector.js, pptx-desde-contrato.js, escritor-pptx.js, xml-ligero.js) se piden la primera vez que llega un
// archivo que no es .pptx: una página que solo recibe .pptx no los baja. Se buscan al lado de este archivo, en assets/js/. En Node (las pruebas) se cargan antes con vm
// y este módulo no pide nada. Requiere zip-propio.js.
(function () {
  'use strict';
  const EXTENSIONES = /\.(pptx|ppt|odp)$/i;
  const base = typeof document !== 'undefined' && document.currentScript && document.currentScript.src ? new URL('./', document.currentScript.src).href : null;

  const esPresentacion = (nombre) => EXTENSIONES.test(String(nombre || ''));
  const sinExtension = (nombre) => String(nombre || '').replace(EXTENSIONES, '');

  // Cada módulo se pide una sola vez; si ya está en la página (otro cargador o una etiqueta <script>) no se repite.
  const cargas = new Map();
  function cargar(global, archivo) {
    if (typeof globalThis[global] !== 'undefined') return Promise.resolve();
    if (!base) return Promise.reject(new Error('Falta el módulo ' + archivo + ' para leer este tipo de presentación.'));
    const src = base + archivo;
    if (!cargas.has(src)) {
      cargas.set(src, new Promise((resolver, rechazar) => {
        const s = document.createElement('script');
        s.src = src;
        s.onload = resolver;
        s.onerror = () => { cargas.delete(src); rechazar(new Error('No se pudo cargar el lector de este tipo de presentación (' + archivo + '). Revisá tu conexión y volvé a intentarlo.')); };
        document.head.appendChild(s);
      }));
    }
    return cargas.get(src);
  }
  async function prepararPpt() {
    await cargar('ZipPropio', 'zip-propio.js');
    await cargar('CfbPropio', 'cfb-propio.js');
    await cargar('EscritorPptx', 'escritor-pptx.js');
    await cargar('PptxDesdeContrato', 'pptx-desde-contrato.js');
    await cargar('PptLector', 'ppt-lector.js');
  }
  async function prepararOdp() {
    await cargar('ZipPropio', 'zip-propio.js');
    await cargar('XmlLigero', 'xml-ligero.js');
    await cargar('EscritorPptx', 'escritor-pptx.js');
    await cargar('PptxDesdeContrato', 'pptx-desde-contrato.js');
    await cargar('OdpLector', 'odp-lector.js');
  }

  const CFB = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
  const esCfb = (b) => b.length >= 8 && CFB.every((x, i) => b[i] === x);
  // Un .odp guarda «mimetype» como primera entrada del zip, sin comprimir: su nombre está en el byte 30. Un .pptx empieza con [Content_Types].xml (o con otra cosa).
  const MIMETYPE = [0x6D, 0x69, 0x6D, 0x65, 0x74, 0x79, 0x70, 0x65];
  const pareceOdp = (b) => b.length > 40 && b[0] === 0x50 && b[1] === 0x4B && MIMETYPE.every((x, i) => b[30 + i] === x);

  async function detectar(bytes) {
    if (esCfb(bytes)) { await prepararPpt(); return PptLector.parece(bytes) ? 'ppt' : null; }
    if (bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4B) {
      await cargar('ZipPropio', 'zip-propio.js');
      let zip;
      try { zip = await ZipPropio.loadAsync(bytes); } catch (e) { return null; }
      if (zip.file('ppt/presentation.xml')) return 'pptx';
      if (zip.file('content.xml') && (zip.file('mimetype') || zip.file('META-INF/manifest.xml'))) { await prepararOdp(); return (await OdpLector.parece(bytes)) ? 'odp' : null; }
      if (zip.file('encrypted-package')) return 'pptx';   // el lector de .pptx avisa que tiene contraseña
    }
    return null;
  }

  async function aPptx(bytes) {
    const origen = await detectar(bytes);
    if (origen === 'pptx') return { bytes, origen, mensajes: [] };
    if (origen === 'ppt') { const r = await PptLector.aPptx(bytes); return { bytes: r.bytes, origen, mensajes: r.mensajes }; }
    if (origen === 'odp') { const r = await OdpLector.aPptx(bytes); return { bytes: r.bytes, origen, mensajes: r.mensajes }; }
    if (esCfb(bytes)) await PptLector.aPptx(bytes);   // otro contenedor de Office (un .pptx con contraseña, un .doc renombrado): el lector dice el motivo
    throw new Error('el archivo no es una presentación de PowerPoint que se pueda leer (.pptx, .ppt u .odp)');
  }

  const aBytes = async (datos) => (typeof Blob !== 'undefined' && datos instanceof Blob ? new Uint8Array(await datos.arrayBuffer()) : datos instanceof Uint8Array ? datos : new Uint8Array(datos));

  async function comoPptx(datos) {
    const b = await aBytes(datos);
    if (esCfb(b)) { await prepararPpt(); return (await PptLector.aPptx(b)).bytes; }
    if (pareceOdp(b)) { await prepararOdp(); if (await OdpLector.parece(b)) return (await OdpLector.aPptx(b)).bytes; }
    return datos;
  }

  async function leerArchivo(archivo) {
    const bytes = new Uint8Array(await archivo.arrayBuffer());
    const r = await aPptx(bytes);
    const buffer = r.bytes.byteOffset === 0 && r.bytes.byteLength === r.bytes.buffer.byteLength ? r.bytes.buffer : r.bytes.slice().buffer;
    return { buffer, bytes: r.bytes, origen: r.origen, mensajes: r.mensajes };
  }

  globalThis.FormatosPowerpoint = { esPresentacion, sinExtension, detectar, aPptx, comoPptx, leerArchivo };
})();

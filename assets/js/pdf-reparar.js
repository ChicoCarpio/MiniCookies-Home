/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Reparar un PDF dañado: la lógica de «Reparar PDF» (reparar-pdf.html), para que la página y su módulo de Flujos
// (flujos-modulos.js) hagan lo mismo. Solo usa lienzos: también corre en los hilos de trabajo de Flujos.
// Necesita pdf-lib (PDFLib) y assets/js/pdfjs.js (abrirPdf).
//
//   const r = await PdfReparar.reparar(buffer, { alAvanzar })
//        buffer: ArrayBuffer del archivo. alAvanzar({ nivel, pagina, total }): nivel 1 al empezar; nivel 2 (sin página)
//        cuando el 1 no alcanzó, y después una vez por página.
//        -> { bytes, nivel, recuperadas, posibles }
//           nivel 1: se volvió a escribir la estructura (el texto sigue siendo texto);
//           nivel 2: cada página quedó como imagen. recuperadas de posibles: páginas que salieron de las que tenía.
//        Lanza un Error con `codigo`: 'sin-paginas' (nivel 1) o 'sin-nada' (no se pudo rescatar ninguna página). Un PDF
//        protegido lanza el error de pdf-lib («…is encrypted…»): no está dañado, hay que desbloquearlo primero.
(function () {
  const OPCIONES_TOLERANTES = { ignoreEncryption: true, throwOnInvalidObject: false, updateMetadata: false };

  // Un PDF cortado (descarga o copia interrumpida) pierde el final: la tabla
  // xref y el "trailer", que es lo que indica dónde empieza el documento. Sin
  // eso ni pdf-lib ni PDF.js lo abren, aunque las páginas sigan todas ahí. Si
  // falta, se descarta el último objeto incompleto y se agrega un trailer que
  // apunta al catálogo (/Type /Catalog); después los dos niveles trabajan sobre
  // eso. Devuelve null si el archivo no parece cortado o no tiene catálogo.
  function completarFinalFaltante(buffer) {
    const bytes = new Uint8Array(buffer);
    // Texto byte a byte (sin decodificar), para buscar marcadores y cortar en
    // la misma posición de los bytes originales.
    let texto = '';
    for (let i = 0; i < bytes.length; i += 0x8000) texto += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    if (/startxref\s+\d+\s*%%EOF\s*$/.test(texto.slice(-2048))) return null;

    let catalogo = null, mayorNumero = 0;
    for (const m of texto.matchAll(/(\d+)\s+(\d+)\s+obj\b([\s\S]*?)endobj/g)) {
      mayorNumero = Math.max(mayorNumero, parseInt(m[1], 10));
      if (/\/Type\s*\/Catalog\b/.test(m[3])) catalogo = `${m[1]} ${m[2]} R`;
    }
    const finUltimoObjeto = texto.lastIndexOf('endobj');
    if (!catalogo || finUltimoObjeto < 0) return null;

    const cola = `\ntrailer\n<< /Size ${mayorNumero + 1} /Root ${catalogo} >>\nstartxref\n0\n%%EOF\n`;
    const resultado = new Uint8Array(finUltimoObjeto + 6 + cola.length);
    resultado.set(bytes.subarray(0, finUltimoObjeto + 6));
    for (let i = 0; i < cola.length; i++) resultado[finUltimoObjeto + 6 + i] = cola.charCodeAt(i);
    return resultado.buffer;
  }

  // Cuántas páginas tiene el original según lo que se puede ver: lo que encuentra
  // el visor y los objetos de página que hay en los bytes (el mayor de los dos).
  async function paginasDelOriginal(bytes) {
    const paginasVisor = (await abrirPdf(bytes.slice(0), { stopAtErrors: false })).numPages;
    const objetosPagina = (new TextDecoder('latin1').decode(bytes).match(/\/Type\s*\/Page(?!s)/g) || []).length;
    return Math.max(paginasVisor, objetosPagina);
  }

  async function intentarNivel1(bytes) {
    const pdf = await PDFLib.PDFDocument.load(bytes, OPCIONES_TOLERANTES);
    if (pdf.getPageCount() === 0) throw Object.assign(new Error('El PDF no tiene páginas'), { codigo: 'sin-paginas' });
    const guardado = await pdf.save();
    // El nivel 1 puede "funcionar" y dejar solo unas páginas (un archivo cortado a
    // la mitad). Se compara con las páginas que el visor encuentra en el original:
    // si el nivel 1 pierde páginas, no se da por bueno y pasa al nivel 2.
    const referencia = await paginasDelOriginal(bytes);
    const paginasGuardadas = (await PDFLib.PDFDocument.load(guardado, OPCIONES_TOLERANTES)).getPageCount();
    if (paginasGuardadas < referencia) throw new Error(`Nivel 1: ${paginasGuardadas} de ${referencia} páginas`);
    return guardado;
  }

  async function intentarNivel2(bytes, alAvanzar) {
    const pdfjsDoc = await abrirPdf(bytes, { stopAtErrors: false });
    if (pdfjsDoc.numPages === 0) throw Object.assign(new Error('No se pudo rescatar ninguna página'), { codigo: 'sin-nada' });
    const pdfNuevo = await PDFLib.PDFDocument.create();
    for (let i = 1; i <= pdfjsDoc.numPages; i++) {
      alAvanzar({ nivel: 2, pagina: i, total: pdfjsDoc.numPages });
      const pagina = await pdfjsDoc.getPage(i);
      const viewport = pagina.getViewport({ scale: 2 });
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      await pagina.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
      const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/jpeg', 0.9));
      const imagenEmbebida = await pdfNuevo.embedJpg(new Uint8Array(await blob.arrayBuffer()));
      const paginaNueva = pdfNuevo.addPage([viewport.width, viewport.height]);
      paginaNueva.drawImage(imagenEmbebida, { x: 0, y: 0, width: viewport.width, height: viewport.height });
      canvas.width = canvas.height = 0;
    }
    return pdfNuevo.save();
  }

  async function reparar(buffer, { alAvanzar } = {}) {
    const avisar = alAvanzar || (() => {});
    // Si el archivo está cortado, se le reconstruye el final antes de todo.
    const original = completarFinalFaltante(buffer) || buffer;

    // Un PDF protegido no está dañado: el nivel 1 (ignoreEncryption) lo
    // "reparaba" guardando el contenido todavía cifrado, y el resultado no
    // abría. Se avisa que hay que desbloquearlo primero.
    try {
      await PDFLib.PDFDocument.load(original.slice(0), { throwOnInvalidObject: false, updateMetadata: false });
    } catch (errorCarga) {
      if (/encrypted/i.test(errorCarga.message)) throw errorCarga;
    }

    let bytes, nivel = 1;
    avisar({ nivel: 1 });
    try {
      bytes = await intentarNivel1(original.slice(0));
    } catch (errorNivel1) {
      console.warn('Nivel 1 falló, probando nivel 2:', errorNivel1);
      nivel = 2;
      avisar({ nivel: 2 });
      bytes = await intentarNivel2(original.slice(0), avisar);
    }
    // Cuántas páginas salieron de cuántas tiene el original: un nivel puede
    // "funcionar" dejando solo unas pocas.
    const recuperadas = (await PDFLib.PDFDocument.load(bytes.slice(0), OPCIONES_TOLERANTES)).getPageCount();
    const posibles = await paginasDelOriginal(original.slice(0));
    return { bytes, nivel, recuperadas, posibles };
  }

  window.PdfReparar = { reparar };
})();

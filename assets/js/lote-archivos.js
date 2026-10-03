/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Convertir varios archivos a la vez. Cada archivo se procesa de uno en uno
// (un PDF grande no ocupa la memoria de todos juntos), un fallo no frena a los
// demás y el resultado sale en un solo .zip, con un archivo por entrada.
//
// convertir(archivo) devuelve { nombre, blob } o lanza un error con el mensaje
// ya traducido. Se usa con JSZip, que ya está en vendor/.
(function () {
  // Dos archivos con el mismo nombre (libro.pdf de dos carpetas) no pueden
  // pisarse dentro del zip: el segundo sale como libro-2.md.
  function nombreUnico(nombre, usados) {
    if (!usados.has(nombre)) { usados.add(nombre); return nombre; }
    const punto = nombre.lastIndexOf('.');
    const base = punto > 0 ? nombre.slice(0, punto) : nombre;
    const extension = punto > 0 ? nombre.slice(punto) : '';
    let n = 2;
    while (usados.has(`${base}-${n}${extension}`)) n++;
    const unico = `${base}-${n}${extension}`;
    usados.add(unico);
    return unico;
  }

  async function convertirLote(archivos, convertir, alAvanzar) {
    const usados = new Set();
    const resultados = [];
    for (let i = 0; i < archivos.length; i++) {
      const archivo = archivos[i];
      if (alAvanzar) alAvanzar(i, archivos.length, archivo);
      try {
        const salida = await convertir(archivo);
        resultados.push({ archivo, ok: true, nombre: nombreUnico(salida.nombre, usados), blob: salida.blob });
      } catch (error) {
        resultados.push({ archivo, ok: false, error: error.message || String(error) });
      }
    }
    return resultados;
  }

  async function zipDeResultados(resultados) {
    const zip = new JSZip();
    resultados.filter((r) => r.ok).forEach((r) => zip.file(r.nombre, r.blob));
    return zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
  }

  window.LoteArchivos = { convertirLote, zipDeResultados };
})();

/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Convertir varios archivos a la vez. Cada archivo se procesa de uno en uno
// (un PDF grande no ocupa la memoria de todos juntos), un fallo no frena a los
// demás y el resultado sale en un solo .zip, con un archivo por entrada.
//
// convertir(archivo) devuelve { nombre, blob } o lanza un error con el mensaje
// ya traducido. Opcional: extras, archivos que acompañan al resultado dentro del
// zip ([{ nombre: 'Doc-imagenes/imagen-1.png', datos }], datos en lo que acepte
// ZipPropio). Se usa con ZipPropio, que ya está en vendor/.
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
        resultados.push({ archivo, ok: true, nombre: nombreUnico(salida.nombre, usados), blob: salida.blob, extras: salida.extras || [] });
      } catch (error) {
        resultados.push({ archivo, ok: false, error: error.message || String(error) });
      }
    }
    return resultados;
  }

  async function zipDeResultados(resultados) {
    const zip = new ZipPropio();
    resultados.filter((r) => r.ok).forEach((r) => {
      zip.file(r.nombre, r.blob);
      (r.extras || []).forEach((extra) => zip.file(extra.nombre, extra.datos));
    });
    return zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
  }

  // Resumen dentro de un contenedor: cuántos salieron y, si hubo fallos, cuáles
  // y por qué. Se arma con DOM y no con HTML: el motivo de un error puede traer
  // texto del propio archivo.
  function pintarResumen(contenedor, resultados, t) {
    const ok = resultados.filter((r) => r.ok).length;
    contenedor.innerHTML = '';
    const resumen = document.createElement('p');
    resumen.textContent = t('lote-convertidos').replace('{ok}', ok).replace('{total}', resultados.length);
    contenedor.appendChild(resumen);
    const fallos = resultados.filter((r) => !r.ok);
    if (fallos.length) {
      const titulo = document.createElement('p');
      titulo.textContent = t('lote-fallos');
      contenedor.appendChild(titulo);
      const ul = document.createElement('ul');
      fallos.forEach((r) => {
        const li = document.createElement('li');
        li.textContent = r.archivo.name + ': ' + r.error;
        ul.appendChild(li);
      });
      contenedor.appendChild(ul);
    }
    contenedor.style.display = 'block';
  }

  window.LoteArchivos = { convertirLote, zipDeResultados, pintarResumen };
})();

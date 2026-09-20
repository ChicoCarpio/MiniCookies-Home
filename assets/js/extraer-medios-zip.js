// Helper compartido para extraer archivos embebidos (imágenes) de formatos que
// son en el fondo un .zip con estructura interna (.docx, .xlsx, .pptx, .epub).
// Requiere que la página ya haya cargado JSZip.

// Devuelve [{ nombre, blob }] para cada entrada del zip cuyo nombre de ruta
// haga match con filtro(rutaCompleta).
async function extraerArchivosDeZip(archivo, filtro) {
  const zip = await JSZip.loadAsync(archivo);
  const rutas = Object.keys(zip.files).filter(ruta => !zip.files[ruta].dir && filtro(ruta));
  const resultado = [];
  for (const ruta of rutas) {
    const blob = await zip.files[ruta].async('blob');
    resultado.push({ nombre: ruta.split('/').pop(), blob });
  }
  return resultado;
}

// Empaqueta una lista de { nombre, blob } en un .zip nuevo, deduplicando
// nombres repetidos (ej. dos imágenes que se llamaban igual en carpetas
// distintas del original).
async function construirZipDeArchivos(archivos) {
  const zip = new JSZip();
  const contadorPorNombre = new Map();
  for (const { nombre, blob } of archivos) {
    let nombreFinal = nombre;
    if (contadorPorNombre.has(nombre)) {
      const contador = contadorPorNombre.get(nombre) + 1;
      contadorPorNombre.set(nombre, contador);
      const punto = nombre.lastIndexOf('.');
      nombreFinal = punto === -1 ? `${nombre}-${contador}` : `${nombre.slice(0, punto)}-${contador}${nombre.slice(punto)}`;
    } else {
      contadorPorNombre.set(nombre, 0);
    }
    zip.file(nombreFinal, blob);
  }
  return zip.generateAsync({ type: 'blob' });
}

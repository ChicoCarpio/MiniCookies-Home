/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Helper compartido para extraer archivos embebidos (imágenes) de formatos que
// son en el fondo un .zip con estructura interna (.docx, .xlsx, .pptx, .epub).
// Requiere que la página ya haya cargado ZipPropio.

// ¿La ruta es una imagen? La carpeta de medios de Office guarda también audio
// y video (un .mp4 de una diapositiva): sin esto salían en el .zip de
// "imágenes". EMF, WMF y TIFF sí cuentan: son imágenes, aunque el navegador
// no las muestre.
function esRutaDeImagen(ruta) {
  return /\.(png|jpe?g|jfif|gif|bmp|tiff?|svg|webp|emf|wmf|ico)$/i.test(ruta);
}

// Devuelve [{ nombre, blob }] para cada entrada del zip cuyo nombre de ruta
// haga match con filtro(rutaCompleta).
async function extraerArchivosDeZip(archivo, filtro) {
  const zip = await ZipPropio.loadAsync(archivo);
  const rutas = Object.keys(zip.files).filter(ruta => !zip.files[ruta].dir && filtro(ruta));
  const resultado = [];
  for (const ruta of rutas) {
    const blob = await zip.files[ruta].async('blob');
    resultado.push({ nombre: ruta.split('/').pop(), blob });
  }
  return resultado;
}

// Las imágenes de cada formato, [{ nombre, blob }], con el mismo filtro que su página (2026-11-10: separadas de
// extraer-imagenes-word/excel/powerpoint/epub.html para que las usen ellas y Flujos). Word y PowerPoint aceptan
// también sus formatos viejos y los de LibreOffice: necesitan formatos-word.js / formatos-powerpoint.js cargados.
async function imagenesDeWord(archivo) {
  const { buffer } = await FormatosWord.leerArchivo(archivo);
  const comoDocx = new File([buffer], FormatosWord.sinExtension(archivo.name) + '.docx');
  return extraerArchivosDeZip(comoDocx, (ruta) => ruta.startsWith('word/media/') && esRutaDeImagen(ruta));
}
function imagenesDeExcel(archivo) {
  return extraerArchivosDeZip(archivo, (ruta) => (ruta.startsWith('xl/media/') || ruta.startsWith('Pictures/')) && esRutaDeImagen(ruta));
}
async function imagenesDePowerPoint(archivo) {
  const buffer = await FormatosPowerpoint.comoPptx(await archivo.arrayBuffer());
  const comoPptx = new File([buffer], FormatosPowerpoint.sinExtension(archivo.name) + '.pptx');
  return extraerArchivosDeZip(comoPptx, (ruta) => ruta.startsWith('ppt/media/') && esRutaDeImagen(ruta));
}
function imagenesDeEpub(archivo) {
  return extraerArchivosDeZip(archivo, (ruta) => /\.(png|jpe?g|gif|svg|webp|bmp)$/i.test(ruta));
}
// Elige por la extensión. Lanza un Error con codigo 'formato' si no es Word, Excel, PowerPoint ni EPUB.
function imagenesDeDocumento(archivo) {
  const extension = ((/\.([^.]+)$/.exec(archivo.name) || [])[1] || '').toLowerCase();
  if (FormatosWord.esDocumento(archivo.name)) return imagenesDeWord(archivo);
  if (['xlsx', 'xlsm', 'xlsb', 'ods'].includes(extension)) return imagenesDeExcel(archivo);
  if (['pptx', 'ppt', 'odp'].includes(extension)) return imagenesDePowerPoint(archivo);
  if (extension === 'epub') return imagenesDeEpub(archivo);
  return Promise.reject(Object.assign(new Error('formato'), { codigo: 'formato' }));
}

// Empaqueta una lista de { nombre, blob } en un .zip nuevo, deduplicando
// nombres repetidos (ej. dos imágenes que se llamaban igual en carpetas
// distintas del original).
async function construirZipDeArchivos(archivos) {
  const zip = new ZipPropio();
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

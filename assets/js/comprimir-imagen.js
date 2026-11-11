/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Comprimir una imagen (JPG o WebP) sin cambiar sus medidas. Lo usan
// «Comprimir imagen» y los flujos. La orientación EXIF la aplica el navegador
// al decodificar la imagen; aquí no se toca.
(function () {
  // calidad: 0..1. formato: 'image/jpeg' o 'image/webp'. Devuelve el Blob;
  // su tipo es el real (un navegador sin WebP devuelve PNG).
  async function comprimir(archivo, { calidad, formato }) {
    // createImageBitmap (no <img>): da los mismos píxeles en la página y en los hilos de Flujos (flujos-hilo.js).
    const img = await createImageBitmap(archivo);
    const canvas = document.createElement('canvas');
    canvas.width = img.width;
    canvas.height = img.height;
    const contexto = canvas.getContext('2d');
    // JPG no tiene canal alfa: sin este relleno, la transparencia de un PNG
    // quedaba pintada de negro.
    if (formato === 'image/jpeg') {
      contexto.fillStyle = '#ffffff';
      contexto.fillRect(0, 0, canvas.width, canvas.height);
    }
    contexto.drawImage(img, 0, 0);
    img.close();
    const blob = await new Promise((resolve) => canvas.toBlob(resolve, formato, calidad));
    if (!blob) throw new Error(typeof t === 'function' ? t('error-sin-imagen') : 'No se pudo generar la imagen');
    return blob;
  }

  window.ComprimirImagen = { comprimir };
})();

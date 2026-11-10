/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Comprimir una imagen (JPG o WebP) sin cambiar sus medidas. Lo usan
// «Comprimir imagen» y los flujos. La orientación EXIF la aplica el navegador
// al decodificar la imagen; aquí no se toca.
(function () {
  function cargarImagen(archivo) {
    return new Promise((resolve, reject) => {
      const url = URL.createObjectURL(archivo);
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = reject;
      img.src = url;
    });
  }

  // calidad: 0..1. formato: 'image/jpeg' o 'image/webp'. Devuelve el Blob;
  // su tipo es el real (un navegador sin WebP devuelve PNG).
  async function comprimir(archivo, { calidad, formato }) {
    const img = await cargarImagen(archivo);
    const canvas = document.createElement('canvas');
    canvas.width = img.naturalWidth;
    canvas.height = img.naturalHeight;
    const contexto = canvas.getContext('2d');
    // JPG no tiene canal alfa: sin este relleno, la transparencia de un PNG
    // quedaba pintada de negro.
    if (formato === 'image/jpeg') {
      contexto.fillStyle = '#ffffff';
      contexto.fillRect(0, 0, canvas.width, canvas.height);
    }
    contexto.drawImage(img, 0, 0);
    const blob = await new Promise((resolve) => canvas.toBlob(resolve, formato, calidad));
    if (!blob) throw new Error(typeof t === 'function' ? t('error-sin-imagen') : 'No se pudo generar la imagen');
    return blob;
  }

  window.ComprimirImagen = { comprimir };
})();

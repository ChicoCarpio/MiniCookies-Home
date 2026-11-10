/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Girar, voltear e inclinar una imagen. Funciones puras: reciben un lienzo y
// devuelven un lienzo nuevo, sin tocar el original ni guardar estado. Las usan
// «Rotar imagen» y los flujos.
(function () {
  function lienzoNuevo(ancho, alto) {
    const lienzo = document.createElement('canvas');
    lienzo.width = ancho;
    lienzo.height = alto;
    return lienzo;
  }

  // Giros rectos (90, 180, 270, -90...): no pierden calidad.
  function rotar(lienzo, grados) {
    const cambiaEjes = Math.abs(grados) % 180 !== 0;
    const salida = lienzoNuevo(cambiaEjes ? lienzo.height : lienzo.width, cambiaEjes ? lienzo.width : lienzo.height);
    const cx = salida.getContext('2d');
    cx.translate(salida.width / 2, salida.height / 2);
    cx.rotate(grados * Math.PI / 180);
    cx.translate(-lienzo.width / 2, -lienzo.height / 2);
    cx.drawImage(lienzo, 0, 0);
    return salida;
  }

  function voltear(lienzo, horizontal) {
    const salida = lienzoNuevo(lienzo.width, lienzo.height);
    const cx = salida.getContext('2d');
    if (horizontal) { cx.translate(lienzo.width, 0); cx.scale(-1, 1); }
    else { cx.translate(0, lienzo.height); cx.scale(1, -1); }
    cx.drawImage(lienzo, 0, 0);
    return salida;
  }

  // El rectángulo más grande, con los lados derechos, que entra en una imagen
  // de ancho × alto girada `radianes` (para recortar las esquinas vacías).
  function rectanguloSinEsquinas(ancho, alto, radianes) {
    const seno = Math.abs(Math.sin(radianes)), coseno = Math.abs(Math.cos(radianes));
    const largo = Math.max(ancho, alto), corto = Math.min(ancho, alto);
    if (corto <= 2 * seno * coseno * largo || Math.abs(seno - coseno) < 1e-10) {
      // Toca los dos lados largos: medio lado corto en cada dirección.
      const medio = corto / 2;
      return ancho >= alto ? [medio / seno, medio / coseno] : [medio / coseno, medio / seno];
    }
    const coseno2 = coseno * coseno - seno * seno;
    return [(ancho * coseno - alto * seno) / coseno2, (alto * coseno - ancho * seno) / coseno2];
  }

  // Ángulo libre. esquinas: 'transparente' (las esquinas vacías quedan sin
  // pintar), 'color' (se rellenan con `color`) o 'recortar' (se recorta al
  // mayor rectángulo sin esquinas vacías; con `color` se pinta el fondo
  // detrás solo si se pasa `color`). Un múltiplo de 90° no recorta.
  function anguloLibre(lienzo, grados, opciones) {
    const { esquinas = 'transparente', color } = opciones || {};
    const ancho = lienzo.width, alto = lienzo.height;
    const radianes = grados * Math.PI / 180;
    const seno = Math.abs(Math.sin(radianes)), coseno = Math.abs(Math.cos(radianes));
    let anchoFinal = Math.round(ancho * coseno + alto * seno);
    let altoFinal = Math.round(ancho * seno + alto * coseno);
    if (esquinas === 'recortar' && grados % 90 !== 0) {
      [anchoFinal, altoFinal] = rectanguloSinEsquinas(ancho, alto, radianes).map((v) => Math.max(1, Math.floor(v)));
    }
    const salida = lienzoNuevo(anchoFinal, altoFinal);
    const cx = salida.getContext('2d');
    if (esquinas === 'color' || (esquinas === 'recortar' && color)) {
      cx.fillStyle = color || '#ffffff';
      cx.fillRect(0, 0, anchoFinal, altoFinal);
    }
    cx.save();
    cx.translate(anchoFinal / 2, altoFinal / 2);
    cx.rotate(radianes);
    cx.drawImage(lienzo, -ancho / 2, -alto / 2);
    cx.restore();
    return salida;
  }

  // Formato de salida: se respetan PNG y WebP (guardan transparencia) y JPG.
  // Cualquier otro formato de entrada (GIF, BMP, SVG...) sale como PNG. Un JPG
  // no tiene transparencia: con esquinas transparentes sale en PNG.
  const FORMATOS = ['image/png', 'image/jpeg', 'image/webp'];
  function formatoDeSalida(tipoMime, quedanTransparentes) {
    const formato = FORMATOS.includes(tipoMime) ? tipoMime : 'image/png';
    return formato === 'image/jpeg' && quedanTransparentes ? 'image/png' : formato;
  }

  window.GirarImagen = { rotar, voltear, anguloLibre, rectanguloSinEsquinas, formatoDeSalida };
})();

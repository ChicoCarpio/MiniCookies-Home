/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Operaciones simples de imagen sin interfaz, separadas de sus páginas para que las usen ellas y Flujos con el
// mismo código:
//   ImagenBasica.escalar(fuente, anchoF, altoF, anchoD, altoD, modo) -> canvas      redimensionar-imagen.html
//   ImagenBasica.LIMITE_DIMENSION_PX, LIMITE_MEGAPIXELES                             (ídem)
//   await ImagenBasica.convertir(fuente, formato, calidad) -> Blob                   convertir-formato-imagen.html
//   await ImagenBasica.leerExif(archivo) -> { marca, modelo, fecha, orientacion, gps } | null   quitar-metadatos-imagen.html
//   await ImagenBasica.quitarMetadatos(archivo, fuente, orientacion?) -> Blob        (ídem)
//   ImagenBasica.formatoDeSalida(tipoMime), ImagenBasica.EXTENSION_FORMATO
// `fuente` es algo que se puede dibujar en un canvas (img, canvas, ImageBitmap) con su tamaño natural.
(function () {
  const LIMITE_DIMENSION_PX = 8000;
  const LIMITE_MEGAPIXELES = 40;

  // Formato de la imagen descargada: se respetan PNG y WebP (guardan
  // transparencia) y JPG. Cualquier otro formato de entrada (GIF, BMP, SVG...)
  // sale como PNG: antes todo lo que no era PNG salía como JPG, y la
  // transparencia de un WebP/GIF quedaba pintada de negro.
  const EXTENSION_FORMATO = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/webp': 'webp' };
  const formatoDeSalida = (tipoMime) => (EXTENSION_FORMATO[tipoMime] ? tipoMime : 'image/png');

  const aBlob = (canvas, tipo, calidad) => new Promise((resolve) => canvas.toBlob(resolve, tipo, calidad));
  const anchoDe = (f) => f.naturalWidth || f.width;
  const altoDe = (f) => f.naturalHeight || f.height;

  // ---------- Redimensionar ----------
  // "nitido": vecino más cercano (sin suavizado) — bloques duros, ideal pixel art.
  // "suave": si se achica mucho (más de 2x), reescala en pasos a la mitad antes
  // del final en vez de un solo salto — evita el aliasing feo de una reducción
  // grande de una sola vez, con suavizado de alta calidad en cada paso.
  function escalar(fuente, anchoFuente, altoFuente, anchoDestino, altoDestino, modo) {
    if (modo === 'nitido') {
      const canvas = document.createElement('canvas');
      canvas.width = anchoDestino;
      canvas.height = altoDestino;
      const ctx = canvas.getContext('2d');
      ctx.imageSmoothingEnabled = false;
      ctx.drawImage(fuente, 0, 0, anchoFuente, altoFuente, 0, 0, anchoDestino, altoDestino);
      return canvas;
    }

    let fuenteActual = fuente;
    let anchoActual = anchoFuente;
    let altoActual = altoFuente;

    while (anchoActual > anchoDestino * 2 && altoActual > altoDestino * 2) {
      const anchoSiguiente = Math.max(anchoDestino, Math.floor(anchoActual / 2));
      const altoSiguiente = Math.max(altoDestino, Math.floor(altoActual / 2));
      const canvasPaso = document.createElement('canvas');
      canvasPaso.width = anchoSiguiente;
      canvasPaso.height = altoSiguiente;
      const ctxPaso = canvasPaso.getContext('2d');
      ctxPaso.imageSmoothingEnabled = true;
      ctxPaso.imageSmoothingQuality = 'high';
      ctxPaso.drawImage(fuenteActual, 0, 0, anchoActual, altoActual, 0, 0, anchoSiguiente, altoSiguiente);
      fuenteActual = canvasPaso;
      anchoActual = anchoSiguiente;
      altoActual = altoSiguiente;
    }

    const canvasFinal = document.createElement('canvas');
    canvasFinal.width = anchoDestino;
    canvasFinal.height = altoDestino;
    const ctxFinal = canvasFinal.getContext('2d');
    ctxFinal.imageSmoothingEnabled = true;
    ctxFinal.imageSmoothingQuality = 'high';
    ctxFinal.drawImage(fuenteActual, 0, 0, anchoActual, altoActual, 0, 0, anchoDestino, altoDestino);
    return canvasFinal;
  }

  // ---------- Convertir formato ----------
  // Si el navegador no sabe codificar el formato pedido (p. ej. WebP), toBlob devuelve PNG: quien llama mira blob.type.
  async function convertir(fuente, formato, calidad) {
    const canvas = document.createElement('canvas');
    canvas.width = anchoDe(fuente);
    canvas.height = altoDe(fuente);
    const ctx = canvas.getContext('2d');
    // JPG no tiene canal alfa — sin este relleno, el navegador pinta la
    // transparencia de negro en vez de dejarla blanca.
    if (formato === 'image/jpeg') {
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
    }
    ctx.drawImage(fuente, 0, 0);
    return aBlob(canvas, formato, formato === 'image/png' ? undefined : calidad);
  }

  // ---------- Metadatos (EXIF) ----------
  // Lector manual de EXIF (solo lectura, para mostrar qué tiene la foto antes
  // de sacarlo — la eliminación en sí no depende de este parseo: pasar la
  // imagen por un <canvas> y reexportarla ya la deja sin ningún metadato).
  function leerIfd(vista, inicioIfd, inicioTiff, esLittleEndian) {
    const entradas = {};
    const cantidad = vista.getUint16(inicioIfd, esLittleEndian);
    for (let i = 0; i < cantidad; i++) {
      const offset = inicioIfd + 2 + i * 12;
      const tag = vista.getUint16(offset, esLittleEndian);
      const tipo = vista.getUint16(offset + 2, esLittleEndian);
      const conteo = vista.getUint32(offset + 4, esLittleEndian);
      const tamanosPorTipo = { 1: 1, 2: 1, 3: 2, 4: 4, 5: 8 };
      const tamanoTotal = (tamanosPorTipo[tipo] || 1) * conteo;
      const posicionValor = tamanoTotal <= 4 ? offset + 8 : inicioTiff + vista.getUint32(offset + 8, esLittleEndian);

      if (tipo === 2) { // ASCII
        let texto = '';
        for (let j = 0; j < conteo - 1; j++) texto += String.fromCharCode(vista.getUint8(posicionValor + j));
        entradas[tag] = texto;
      } else if (tipo === 3) { // SHORT
        entradas[tag] = vista.getUint16(posicionValor, esLittleEndian);
      } else if (tipo === 4) { // LONG
        entradas[tag] = vista.getUint32(posicionValor, esLittleEndian);
      } else if (tipo === 5 && conteo === 1) { // RATIONAL
        entradas[tag] = vista.getUint32(posicionValor, esLittleEndian) / vista.getUint32(posicionValor + 4, esLittleEndian);
      } else if (tipo === 5 && conteo === 3) { // 3 RATIONAL (coordenadas GPS)
        entradas[tag] = [0, 1, 2].map((j) =>
          vista.getUint32(posicionValor + j * 8, esLittleEndian) / vista.getUint32(posicionValor + j * 8 + 4, esLittleEndian)
        );
      }
    }
    return entradas;
  }

  function gradosDesdeGps(valores, referencia) {
    const [grados, minutos, segundos] = valores;
    let decimal = grados + minutos / 60 + segundos / 3600;
    if (referencia === 'S' || referencia === 'W') decimal *= -1;
    return decimal;
  }

  async function leerExif(archivo) {
    if (archivo.type !== 'image/jpeg') return null; // EXIF solo existe en JPEG
    const buffer = await archivo.arrayBuffer();
    const vista = new DataView(buffer);
    if (vista.getUint16(0) !== 0xFFD8) return null; // no es un JPEG válido

    let offset = 2;
    while (offset < vista.byteLength - 4) {
      const marcador = vista.getUint16(offset);
      // Puede haber varios APP1 (p. ej. XMP antes del EXIF): se revisan todos.
      if (marcador === 0xFFE1 && vista.getUint32(offset + 4) === 0x45786966) { // "Exif"
        const inicioApp1 = offset + 4;
        const inicioTiff = inicioApp1 + 6;
        const esLittleEndian = vista.getUint16(inicioTiff) === 0x4949;
        const inicioIfd0 = inicioTiff + vista.getUint32(inicioTiff + 4, esLittleEndian);
        const ifd0 = leerIfd(vista, inicioIfd0, inicioTiff, esLittleEndian);

        let ifdExif = {};
        if (ifd0[0x8769]) {
          ifdExif = leerIfd(vista, inicioTiff + ifd0[0x8769], inicioTiff, esLittleEndian);
        }
        let gps = null;
        if (ifd0[0x8825]) {
          const ifdGps = leerIfd(vista, inicioTiff + ifd0[0x8825], inicioTiff, esLittleEndian);
          if (ifdGps[2] && ifdGps[4]) {
            gps = {
              lat: gradosDesdeGps(ifdGps[2], ifdGps[1]),
              lon: gradosDesdeGps(ifdGps[4], ifdGps[3]),
            };
          }
        }

        return {
          marca: ifd0[0x010F] || null,
          modelo: ifd0[0x0110] || null,
          fecha: ifdExif[0x9003] || ifd0[0x0132] || null,
          orientacion: ifd0[0x0112] || null,
          gps,
        };
      }
      if ((marcador & 0xFF00) !== 0xFF00) break; // ya no es un marcador válido
      if (marcador === 0xFFDA) break; // arrancó el contenido de la imagen, no hay más APPn
      offset += 2 + vista.getUint16(offset + 2);
    }
    return null;
  }

  // Quita los metadatos de un JPEG SIN recomprimirlo: recorre los segmentos del
  // archivo y descarta APP1 (EXIF/XMP), APP13 (IPTC/Photoshop), el resto de
  // APPn de fabricantes y los comentarios (COM). Se conservan APP0 (JFIF) y el
  // perfil de color ICC (APP2 "ICC_PROFILE"), que no identifica a nadie pero
  // cambia cómo se ven los colores. También se descarta todo lo que venga
  // después del final de la imagen principal: los celulares guardan ahí
  // imágenes extra (miniaturas, mapas de profundidad) con su propio EXIF.
  // La imagen comprimida en sí queda idéntica byte a byte.
  // Devuelve null si el archivo no tiene la estructura esperada.
  function quitarMetadatosJpegSinPerdida(buffer) {
    const bytes = new Uint8Array(buffer);
    const vista = new DataView(buffer);
    const largoTotal = bytes.length;
    if (largoTotal < 4 || vista.getUint16(0) !== 0xFFD8) return null;
    const partes = [bytes.subarray(0, 2)];
    let offset = 2;
    // 1) Cabecera: segmentos con largo hasta el primer SOS
    while (true) {
      if (offset + 4 > largoTotal) return null;
      const marcador = vista.getUint16(offset);
      if ((marcador & 0xFF00) !== 0xFF00) return null;
      if (marcador === 0xFFDA) break;
      const largo = vista.getUint16(offset + 2);
      if (largo < 2 || offset + 2 + largo > largoTotal) return null;
      const esAppN = marcador >= 0xFFE0 && marcador <= 0xFFEF;
      const esIcc = marcador === 0xFFE2 && String.fromCharCode(...bytes.subarray(offset + 4, offset + 15)) === 'ICC_PROFILE';
      const seQueda = !(esAppN || marcador === 0xFFFE) || marcador === 0xFFE0 || esIcc;
      if (seQueda) partes.push(bytes.subarray(offset, offset + 2 + largo));
      offset += 2 + largo;
    }
    // 2) Imagen: desde el primer SOS hasta su EOI (FFD9). Los datos comprimidos
    // no tienen largo, así que se recorren hasta el siguiente marcador real
    // (FF00 es un FF "escapado" y FFD0–FFD7 son marcas de reinicio). Las JPEG
    // progresivas tienen varios SOS con tablas en el medio: se siguen igual.
    const inicioImagen = offset;
    let fin = -1;
    while (offset + 2 <= largoTotal) {
      const marcador = vista.getUint16(offset);
      if (marcador === 0xFFD9) { fin = offset + 2; break; }
      if ((marcador & 0xFF00) !== 0xFF00 || offset + 4 > largoTotal) return null;
      const largo = vista.getUint16(offset + 2);
      offset += 2 + largo;
      if (marcador !== 0xFFDA) continue;
      while (offset + 1 < largoTotal) {
        if (bytes[offset] !== 0xFF) { offset++; continue; }
        const siguiente = bytes[offset + 1];
        if (siguiente === 0x00 || (siguiente >= 0xD0 && siguiente <= 0xD7)) { offset += 2; continue; }
        if (siguiente === 0xFF) { offset++; continue; } // bytes de relleno antes de un marcador
        break;
      }
    }
    if (fin < 0) return null;
    partes.push(bytes.subarray(inicioImagen, fin));
    const salida = new Uint8Array(partes.reduce((suma, p) => suma + p.length, 0));
    let posicion = 0;
    for (const parte of partes) { salida.set(parte, posicion); posicion += parte.length; }
    return salida;
  }

  // JPEG con orientación normal: se sacan los metadatos sin tocar la imagen
  // (sin pérdida de calidad). Si la foto depende de la etiqueta de
  // orientación del EXIF para verse derecha (típico de celulares), borrarla
  // la dejaría girada: en ese caso, igual que con PNG/WebP, se redibuja en
  // un canvas (que ya la aplica) y se exporta de nuevo — una imagen
  // completamente nueva, sin ningún bloque EXIF/XMP/IPTC del original.
  // orientacion: la ya leída con leerExif (undefined: se lee aquí; null: no se pudo leer, va por el canvas).
  async function quitarMetadatos(archivo, fuente, orientacion) {
    if (orientacion === undefined) {
      try { const d = await leerExif(archivo); orientacion = d && d.orientacion ? d.orientacion : 1; } catch (e) { orientacion = null; }
    }
    if (archivo.type === 'image/jpeg' && orientacion === 1) {
      const limpio = quitarMetadatosJpegSinPerdida(await archivo.arrayBuffer());
      if (limpio) return new Blob([limpio], { type: 'image/jpeg' });
    }
    const canvas = document.createElement('canvas');
    canvas.width = anchoDe(fuente);
    canvas.height = altoDe(fuente);
    canvas.getContext('2d').drawImage(fuente, 0, 0);
    return aBlob(canvas, formatoDeSalida(archivo.type), 0.95);
  }

  window.ImagenBasica = {
    escalar, convertir, leerExif, quitarMetadatos, formatoDeSalida, EXTENSION_FORMATO, LIMITE_DIMENSION_PX, LIMITE_MEGAPIXELES,
  };
})();

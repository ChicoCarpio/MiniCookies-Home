/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Imágenes que vienen dentro de un Word/Excel/PowerPoint → un formato común.
//
// Office guarda las tablas y gráficos pegados desde Excel como EMF (dibujo
// vectorial de Windows) que ningún navegador muestra y casi nada fuera de
// Office abre. Acá:
//  - PNG, JPEG, GIF y SVG quedan como están.
//  - EMF se dibuja con emf.js (lector propio) y sale como PNG al doble de
//    tamaño, para que el texto de una tabla se lea.
//  - Otros que el navegador sí abre (BMP, WebP...) pasan a PNG.
//  - Lo que no se puede abrir (WMF, TIFF) se devuelve tal cual.
//
// Requiere: emf.js.
// Uso: const { bytes, extension, tipo } = await ImagenesOffice.aFormatoComun(bytes, contentType);

(function () {
  function tipoDe(bytes, contentType) {
    const b = bytes;
    if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return 'png';
    if (b[0] === 0xff && b[1] === 0xd8) return 'jpg';
    if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return 'gif';
    if (b[0] === 0x01 && b[1] === 0x00 && b[2] === 0x00 && b[3] === 0x00 && b[40] === 0x20 && b[41] === 0x45 && b[42] === 0x4d && b[43] === 0x46) return 'emf'; // " EMF" en el encabezado
    if ((b[0] === 0xd7 && b[1] === 0xcd && b[2] === 0xc6 && b[3] === 0x9a) || (b[0] === 0x01 && b[1] === 0x00 && b[2] === 0x09 && b[3] === 0x00)) return 'wmf';
    const ct = (contentType || '').toLowerCase();
    if (/svg/.test(ct)) return 'svg';
    if (/emf/.test(ct)) return 'emf';
    if (/wmf/.test(ct)) return 'wmf';
    return ct.split('/')[1] || 'bin';
  }

  async function canvasAPng(dibujar, ancho, alto) {
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round(ancho));
    canvas.height = Math.max(1, Math.round(alto));
    const ctx = canvas.getContext('2d');
    await dibujar(ctx, canvas.width, canvas.height);
    const blob = await new Promise((r) => canvas.toBlob(r, 'image/png'));
    return new Uint8Array(await blob.arrayBuffer());
  }

  async function aFormatoComun(entrada, contentType) {
    const bytes = entrada instanceof Uint8Array ? entrada : new Uint8Array(entrada);
    const tipo = tipoDe(bytes, contentType);
    if (['png', 'jpg', 'gif', 'svg'].includes(tipo)) {
      return { bytes, extension: tipo, tipo: tipo === 'jpg' ? 'image/jpeg' : tipo === 'svg' ? 'image/svg+xml' : `image/${tipo}` };
    }
    try {
      let png;
      if (tipo === 'emf') {
        const canvas = await Emf.dibujar(bytes, { escalaMaxima: 2, ladoMaximo: 3000 });
        png = await canvasAPng((ctx) => ctx.drawImage(canvas, 0, 0), canvas.width, canvas.height);
      } else if (tipo === 'wmf') {
        throw new Error('WMF sin lector');
      } else {
        const mapa = await createImageBitmap(new Blob([bytes]));
        png = await canvasAPng((ctx) => ctx.drawImage(mapa, 0, 0), mapa.width, mapa.height);
      }
      return { bytes: png, extension: 'png', tipo: 'image/png' };
    } catch {
      // No se pudo convertir: mejor el original que perder la imagen.
      return { bytes, extension: tipo, tipo: contentType || 'application/octet-stream' };
    }
  }

  window.ImagenesOffice = { aFormatoComun };
})();

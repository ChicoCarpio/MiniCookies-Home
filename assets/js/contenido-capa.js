/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Capas con contenido propio del Editor avanzado: texto o una imagen (otra
// foto, un logo, o un recorte de la misma foto), que se mueven, escalan y
// rotan sobre la imagen.
//
// Todo se mide en píxeles de la foto ORIGINAL: el centro va en 0..1 del
// ancho/alto, y "escala" 1 es el tamaño natural del contenido en esa foto. Así
// el mismo contenido sale igual en la vista previa (más chica) y en la
// descarga a resolución completa: quien dibuja pasa "escalaVista", la
// relación entre el lienzo de trabajo y la foto original.
//
// Un texto se dibuja primero en su propio lienzo (la "fuente", que solo
// cambia si cambia el texto o su estilo) y después se trata igual que una
// imagen: mover o rotar no vuelve a componer las letras.

window.ContenidoCapa = (function () {
  'use strict';

  // Fuentes del sistema: no se descarga ninguna. Cada una con su respaldo
  // genérico por si falta en el equipo.
  const FAMILIAS = [
    ['Arial, Helvetica, sans-serif', 'Arial'],
    ['Verdana, Geneva, sans-serif', 'Verdana'],
    ['"Trebuchet MS", sans-serif', 'Trebuchet'],
    ['Impact, "Arial Black", sans-serif', 'Impact'],
    ['Georgia, serif', 'Georgia'],
    ['"Times New Roman", Times, serif', 'Times'],
    ['"Courier New", Courier, monospace', 'Courier'],
    ['"Comic Sans MS", "Comic Neue", cursive', 'Comic'],
  ];

  function transformacionInicial() {
    return { centro: { x: 0.5, y: 0.5 }, escala: 1, rotacion: 0, voltearH: false, voltearV: false };
  }

  // ladoMenor: el lado menor de la foto original, para que la letra arranque
  // de un tamaño que se vea en cualquier foto.
  function textoNuevo(texto, ladoMenor) {
    return {
      tipo: 'texto',
      texto,
      familia: FAMILIAS[0][0],
      px: Math.max(12, Math.round(ladoMenor * 0.1)),
      color: '#ffffff',
      negrita: true,
      cursiva: false,
      alineacion: 'center', // left | center | right
      contorno: '#000000',
      grosorContorno: 8, // % del tamaño de la letra
      sombra: false,
      ...transformacionInicial(),
    };
  }

  let siguienteIdLienzo = 1;
  function imagenNueva(lienzo) {
    return { tipo: 'imagen', lienzo, _idLienzo: siguienteIdLienzo++, ...transformacionInicial() };
  }

  // Lo que define cómo se ven las letras (no dónde están).
  function firmaFuente(c) {
    if (c.tipo !== 'texto') return 'imagen#' + c._idLienzo;
    return JSON.stringify([c.texto, c.familia, c.px, c.color, c.negrita, c.cursiva, c.alineacion, c.contorno, c.grosorContorno, c.sombra]);
  }

  // Todo lo que cambia el resultado (fuente + transformación).
  function firma(c) {
    return firmaFuente(c) + JSON.stringify([c.centro, c.escala, c.rotacion, c.voltearH, c.voltearV]);
  }

  function fuenteDeTexto(c) {
    const fuente = `${c.cursiva ? 'italic ' : ''}${c.negrita ? 'bold ' : ''}${c.px}px ${c.familia}`;
    const lineas = String(c.texto || ' ').split('\n');
    const medir = document.createElement('canvas').getContext('2d');
    medir.font = fuente;
    const anchos = lineas.map((l) => medir.measureText(l || ' ').width);
    const alto = c.px * 1.25;
    const contorno = c.grosorContorno > 0 ? c.px * c.grosorContorno / 100 : 0;
    const margen = Math.ceil(contorno + (c.sombra ? c.px * 0.15 : 0) + c.px * 0.1);
    const ancho = Math.ceil(Math.max(...anchos, 1));
    const lienzo = document.createElement('canvas');
    lienzo.width = ancho + margen * 2;
    lienzo.height = Math.ceil(alto * lineas.length) + margen * 2;
    const ctx = lienzo.getContext('2d');
    ctx.font = fuente;
    ctx.textBaseline = 'middle';
    ctx.lineJoin = 'round';
    ctx.textAlign = c.alineacion;
    const x = c.alineacion === 'left' ? margen : (c.alineacion === 'right' ? margen + ancho : margen + ancho / 2);
    lineas.forEach((linea, i) => {
      const y = margen + alto * (i + 0.5);
      if (c.sombra) {
        ctx.save();
        ctx.shadowColor = 'rgba(0,0,0,0.6)';
        ctx.shadowBlur = c.px * 0.08;
        ctx.shadowOffsetX = ctx.shadowOffsetY = c.px * 0.04;
        ctx.fillStyle = c.color;
        ctx.fillText(linea, x, y);
        ctx.restore();
      }
      if (contorno) {
        ctx.strokeStyle = c.contorno;
        ctx.lineWidth = contorno * 2;
        ctx.strokeText(linea, x, y);
      }
      ctx.fillStyle = c.color;
      ctx.fillText(linea, x, y);
    });
    return lienzo;
  }

  // Centro, medio ancho/alto ya escalados y giro, en píxeles del lienzo.
  function caja(c, fuente, ancho, alto, escalaVista) {
    const s = c.escala * escalaVista;
    return {
      cx: c.centro.x * ancho,
      cy: c.centro.y * alto,
      mx: fuente.width * s / 2,
      my: fuente.height * s / 2,
      giro: c.rotacion * Math.PI / 180,
    };
  }

  // Un punto del lienzo, en coordenadas de la caja (sin giro, centro en 0,0).
  function puntoLocal(k, x, y) {
    const dx = x - k.cx, dy = y - k.cy;
    const cos = Math.cos(k.giro), sin = Math.sin(k.giro);
    return { x: dx * cos + dy * sin, y: -dx * sin + dy * cos };
  }

  // Y al revés: de la caja al lienzo.
  function puntoLienzo(k, x, y) {
    const cos = Math.cos(k.giro), sin = Math.sin(k.giro);
    return { x: k.cx + x * cos - y * sin, y: k.cy + x * sin + y * cos };
  }

  // Las cuatro esquinas (para escalar) y la manija de girar, por encima del
  // borde de arriba. "separacion" en píxeles del lienzo.
  function manijas(k, separacion) {
    return {
      esquinas: [[-1, -1], [1, -1], [1, 1], [-1, 1]].map(([sx, sy]) => puntoLienzo(k, sx * k.mx, sy * k.my)),
      bordeArriba: puntoLienzo(k, 0, -k.my),
      girar: puntoLienzo(k, 0, -k.my - separacion),
    };
  }

  // Dibuja el contenido transformado en "destino" (de ancho x alto).
  function dibujar(c, fuente, destino, escalaVista) {
    const ctx = destino.getContext('2d');
    const k = caja(c, fuente, destino.width, destino.height, escalaVista);
    ctx.save();
    ctx.imageSmoothingQuality = 'high';
    ctx.translate(k.cx, k.cy);
    ctx.rotate(k.giro);
    ctx.scale(c.voltearH ? -1 : 1, c.voltearV ? -1 : 1);
    ctx.drawImage(fuente, -k.mx, -k.my, k.mx * 2, k.my * 2);
    ctx.restore();
    return destino;
  }

  // Recorta un lienzo a lo que no es transparente. Devuelve el recorte y
  // dónde estaba, o null si no queda nada.
  function recortarATransparencia(lienzo) {
    const { width: w, height: h } = lienzo;
    const d = lienzo.getContext('2d').getImageData(0, 0, w, h).data;
    let x0 = w, y0 = h, x1 = -1, y1 = -1;
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        if (d[(y * w + x) * 4 + 3] > 8) {
          if (x < x0) x0 = x;
          if (x > x1) x1 = x;
          if (y < y0) y0 = y;
          if (y > y1) y1 = y;
        }
      }
    }
    if (x1 < 0) return null;
    const recorte = document.createElement('canvas');
    recorte.width = x1 - x0 + 1; recorte.height = y1 - y0 + 1;
    recorte.getContext('2d').drawImage(lienzo, -x0, -y0);
    return { lienzo: recorte, x: x0, y: y0 };
  }

  return {
    FAMILIAS, textoNuevo, imagenNueva, firma, firmaFuente, fuenteDeTexto,
    caja, puntoLocal, puntoLienzo, manijas, dibujar, recortarATransparencia,
  };
})();

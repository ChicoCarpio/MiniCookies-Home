/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Efectos de textura (pixelado y variantes de desenfoque) reusados por
// censurar-imagen.html y Editor-de-color-de-imagen.html.
//
// Dos formas de llamarlos:
// - aplicarRegion(nombre, ctxDestino, fuente, sx, sy, sw, sh, intensidad):
//   aplica el efecto sobre un rectángulo de ctxDestino. Si se pasa "fuente"
//   (una imagen u otro canvas) y es distinta del propio canvas de destino,
//   primero se copia ese rectángulo desde "fuente" antes de procesar — así
//   sirve tanto para pintar sobre un canvas recién creado (pasando la imagen
//   original) como para re-procesar un canvas que ya tiene contenido
//   (pasando undefined, o el propio canvas, como fuente).
// - aplicarCompleto(nombre, ctxDestino, fuente, ancho, alto, intensidad):
//   atajo para aplicar sobre el canvas entero (0,0,ancho,alto).
window.EfectosTextura = (function () {
  function pixelar(ctx, sx, sy, sw, sh, bloque) {
    const datos = ctx.getImageData(sx, sy, sw, sh);
    const d = datos.data;
    for (let by = 0; by < sh; by += bloque) {
      for (let bx = 0; bx < sw; bx += bloque) {
        const anchoBloque = Math.min(bloque, sw - bx);
        const altoBloque = Math.min(bloque, sh - by);
        let r = 0, g = 0, b = 0, n = 0;
        for (let y = 0; y < altoBloque; y++) {
          for (let x = 0; x < anchoBloque; x++) {
            const i = ((by + y) * sw + (bx + x)) * 4;
            r += d[i]; g += d[i + 1]; b += d[i + 2]; n++;
          }
        }
        r = Math.round(r / n); g = Math.round(g / n); b = Math.round(b / n);
        for (let y = 0; y < altoBloque; y++) {
          for (let x = 0; x < anchoBloque; x++) {
            const i = ((by + y) * sw + (bx + x)) * 4;
            d[i] = r; d[i + 1] = g; d[i + 2] = b;
          }
        }
      }
    }
    ctx.putImageData(datos, sx, sy);
  }

  // Igual que pixelar, pero dibuja un punto (círculo) de cada color promedio
  // sobre un fondo oscurecido del mismo bloque — mosaico de puntos.
  function pixelarCirculos(ctx, sx, sy, sw, sh, bloque) {
    const datos = ctx.getImageData(sx, sy, sw, sh);
    const d = datos.data;
    const bloques = [];
    for (let by = 0; by < sh; by += bloque) {
      for (let bx = 0; bx < sw; bx += bloque) {
        const anchoBloque = Math.min(bloque, sw - bx);
        const altoBloque = Math.min(bloque, sh - by);
        let r = 0, g = 0, b = 0, n = 0;
        for (let y = 0; y < altoBloque; y++) {
          for (let x = 0; x < anchoBloque; x++) {
            const i = ((by + y) * sw + (bx + x)) * 4;
            r += d[i]; g += d[i + 1]; b += d[i + 2]; n++;
          }
        }
        r = Math.round(r / n); g = Math.round(g / n); b = Math.round(b / n);
        const fondoR = Math.round(r * 0.55), fondoG = Math.round(g * 0.55), fondoB = Math.round(b * 0.55);
        for (let y = 0; y < altoBloque; y++) {
          for (let x = 0; x < anchoBloque; x++) {
            const i = ((by + y) * sw + (bx + x)) * 4;
            d[i] = fondoR; d[i + 1] = fondoG; d[i + 2] = fondoB;
          }
        }
        bloques.push({ bx, by, anchoBloque, altoBloque, r, g, b });
      }
    }
    ctx.putImageData(datos, sx, sy);
    for (const bl of bloques) {
      const cx = sx + bl.bx + bl.anchoBloque / 2;
      const cy = sy + bl.by + bl.altoBloque / 2;
      const radio = Math.min(bl.anchoBloque, bl.altoBloque) * 0.42;
      ctx.beginPath();
      ctx.arc(cx, cy, radio, 0, Math.PI * 2);
      ctx.fillStyle = `rgb(${bl.r},${bl.g},${bl.b})`;
      ctx.fill();
    }
  }

  // Pixela y suma ruido aleatorio por píxel — más difícil de "promediar"
  // hacia atrás que un pixelado liso.
  function ruido(ctx, sx, sy, sw, sh, bloque) {
    pixelar(ctx, sx, sy, sw, sh, bloque);
    const datos = ctx.getImageData(sx, sy, sw, sh);
    const d = datos.data;
    const fuerza = Math.min(90, bloque * 2.2);
    for (let i = 0; i < d.length; i += 4) {
      const n = (Math.random() - 0.5) * 2 * fuerza;
      d[i] = Math.min(255, Math.max(0, d[i] + n));
      d[i + 1] = Math.min(255, Math.max(0, d[i + 1] + n));
      d[i + 2] = Math.min(255, Math.max(0, d[i + 2] + n));
    }
    ctx.putImageData(datos, sx, sy);
  }

  // "fuente" es de donde se toman los píxeles a desenfocar. Se recorta con
  // un margen extra antes de desenfocar y se vuelve a recortar al tamaño
  // real al pegar, así el desenfoque no aclara los bordes trayendo "aire" de
  // afuera del recorte.
  function difuminar(ctxDestino, fuente, sx, sy, sw, sh, radio) {
    const margen = Math.max(1, Math.round(radio)) * 2;
    const canvasChico = document.createElement('canvas');
    canvasChico.width = sw + margen * 2;
    canvasChico.height = sh + margen * 2;
    const ctxChico = canvasChico.getContext('2d');
    ctxChico.drawImage(fuente, sx - margen, sy - margen, sw + margen * 2, sh + margen * 2, 0, 0, sw + margen * 2, sh + margen * 2);

    const canvasBorroso = document.createElement('canvas');
    canvasBorroso.width = canvasChico.width;
    canvasBorroso.height = canvasChico.height;
    const ctxBorroso = canvasBorroso.getContext('2d');
    ctxBorroso.filter = `blur(${radio}px)`;
    ctxBorroso.drawImage(canvasChico, 0, 0);

    ctxDestino.drawImage(canvasBorroso, margen, margen, sw, sh, sx, sy, sw, sh);
  }

  // Desenfoque de movimiento: superpone la región varias veces con un
  // pequeño desplazamiento horizontal creciente y opacidad baja.
  function direccional(ctxDestino, fuente, sx, sy, sw, sh, fuerza) {
    const distancia = Math.max(2, Math.round(fuerza));
    const margen = distancia + 2;
    const canvasChico = document.createElement('canvas');
    canvasChico.width = sw + margen * 2;
    canvasChico.height = sh + margen * 2;
    const ctxChico = canvasChico.getContext('2d');
    ctxChico.drawImage(fuente, sx - margen, sy - margen, sw + margen * 2, sh + margen * 2, 0, 0, sw + margen * 2, sh + margen * 2);

    const canvasSmear = document.createElement('canvas');
    canvasSmear.width = canvasChico.width;
    canvasSmear.height = canvasChico.height;
    const ctxSmear = canvasSmear.getContext('2d');
    const pasos = 14;
    ctxSmear.globalAlpha = 1 / pasos;
    for (let i = 0; i < pasos; i++) {
      const offset = (i / (pasos - 1) - 0.5) * distancia * 2;
      ctxSmear.drawImage(canvasChico, offset, 0);
    }
    ctxSmear.globalAlpha = 1;

    ctxDestino.drawImage(canvasSmear, margen, margen, sw, sh, sx, sy, sw, sh);
  }

  // Vidrio esmerilado: desenfoca de base y, por cada píxel de salida, toma
  // el color de un vecino elegido al azar (en vez de un promedio suave) —
  // resultado granulado, como mirar a través de vidrio esmerilado.
  function esmerilado(ctxDestino, fuente, sx, sy, sw, sh, fuerza) {
    const radioBase = Math.max(1, Math.round(fuerza / 3));
    const margen = radioBase * 2 + 4;
    const base = document.createElement('canvas');
    base.width = sw + margen * 2;
    base.height = sh + margen * 2;
    const ctxBase = base.getContext('2d');
    ctxBase.drawImage(fuente, sx - margen, sy - margen, sw + margen * 2, sh + margen * 2, 0, 0, sw + margen * 2, sh + margen * 2);

    const suave = document.createElement('canvas');
    suave.width = base.width;
    suave.height = base.height;
    const ctxSuave = suave.getContext('2d');
    ctxSuave.filter = `blur(${radioBase}px)`;
    ctxSuave.drawImage(base, 0, 0);

    const datosOrigen = ctxSuave.getImageData(0, 0, suave.width, suave.height);
    const resultado = document.createElement('canvas');
    resultado.width = sw;
    resultado.height = sh;
    const ctxResultado = resultado.getContext('2d');
    const datosDestino = ctxResultado.createImageData(sw, sh);
    const dispersion = Math.max(2, Math.round(fuerza * 0.6));
    for (let y = 0; y < sh; y++) {
      for (let x = 0; x < sw; x++) {
        const dx = Math.round((Math.random() - 0.5) * dispersion);
        const dy = Math.round((Math.random() - 0.5) * dispersion);
        const ox = Math.max(0, Math.min(suave.width - 1, margen + x + dx));
        const oy = Math.max(0, Math.min(suave.height - 1, margen + y + dy));
        const iOrigen = (oy * suave.width + ox) * 4;
        const iDestino = (y * sw + x) * 4;
        datosDestino.data[iDestino] = datosOrigen.data[iOrigen];
        datosDestino.data[iDestino + 1] = datosOrigen.data[iOrigen + 1];
        datosDestino.data[iDestino + 2] = datosOrigen.data[iOrigen + 2];
        datosDestino.data[iDestino + 3] = 255;
      }
    }
    ctxResultado.putImageData(datosDestino, 0, 0);
    ctxDestino.drawImage(resultado, sx, sy);
  }

  const EN_LUGAR = { pixelar: pixelar, 'pixelar-circulos': pixelarCirculos, ruido: ruido };
  const CON_FUENTE = { direccional: direccional, esmerilado: esmerilado };

  function aplicarRegion(nombre, ctxDestino, fuente, sx, sy, sw, sh, intensidad) {
    const bloque = Math.max(2, Math.round(intensidad));
    if (EN_LUGAR[nombre]) {
      if (fuente && fuente !== ctxDestino.canvas) {
        ctxDestino.drawImage(fuente, sx, sy, sw, sh, sx, sy, sw, sh);
      }
      EN_LUGAR[nombre](ctxDestino, sx, sy, sw, sh, bloque);
      return;
    }
    const fuenteReal = fuente || ctxDestino.canvas;
    (CON_FUENTE[nombre] || difuminar)(ctxDestino, fuenteReal, sx, sy, sw, sh, intensidad);
  }

  function aplicarCompleto(nombre, ctxDestino, fuente, ancho, alto, intensidad) {
    aplicarRegion(nombre, ctxDestino, fuente, 0, 0, ancho, alto, intensidad);
  }

  return {
    pixelar, pixelarCirculos, ruido, difuminar, direccional, esmerilado,
    aplicarRegion, aplicarCompleto,
    NOMBRES: ['pixelar', 'pixelar-circulos', 'ruido', 'difuminar', 'direccional', 'esmerilado'],
  };
})();

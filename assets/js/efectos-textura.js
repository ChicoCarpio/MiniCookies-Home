/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
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
//
// Los mosaicos (pixel art, punto de cruz, cuentas, pop art…) y el cómic viven
// en assets/js/efectos-mosaico.js y entran por acá mismo: solo están en las
// páginas que cargan ese módulo, ignoran la intensidad y reciben un último
// argumento "opciones" (ver ese archivo).
window.EfectosTextura = (function () {
  // Ancho de referencia contra el que se calibra la intensidad/tamaño de
  // toda textura con margen (difuminar, movimiento, esmerilado) o bloque
  // (pixelar, puntos, grano/ruido): el deslizador pasa a representar un
  // porcentaje del ancho de LA IMAGEN QUE SE ESTÉ PROCESANDO en cada momento
  // (ancho / REFERENCIA_TEXTURA), no un número de píxeles fijo. Antes era un
  // número de píxeles calibrado contra fotos de alrededor de este tamaño (el
  // tope del generador de stickers): la misma posición del deslizador se
  // veía mucho más débil en una foto grande que en un sticker, y hasta entre
  // dos stickers de tamaño real distinto (uno recortado de una foto más
  // chica que el tope de 512px).
  const REFERENCIA_TEXTURA = 512;

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
  // hacia atrás que un pixelado liso. Usada SOLO por censurar-imagen.html: el
  // pixelado ahí es la parte que realmente tapa la imagen (el ruido nomás
  // dificulta reconstruirla), así que esta función no se toca aunque no se
  // use como filtro decorativo en ningún otro lado — ver grano() abajo.
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

  // Grano fino: el ruido de arriba, pero SIN pixelar antes — ruido por píxel
  // derecho, para un efecto de grano de película en vez de mosaico con
  // estática. Es la textura "Ruido" de los editores y el generador de
  // stickers (ver CLAVES_NOMBRE/crearGaleria más abajo); censurar-imagen.html
  // sigue usando ruido() tal cual, sin pasar por acá.
  function grano(ctx, sx, sy, sw, sh, fuerza) {
    const datos = ctx.getImageData(sx, sy, sw, sh);
    const d = datos.data;
    const amplitud = Math.min(90, fuerza * 2.2);
    for (let i = 0; i < d.length; i += 4) {
      const n = (Math.random() - 0.5) * 2 * amplitud;
      d[i] = Math.min(255, Math.max(0, d[i] + n));
      d[i + 1] = Math.min(255, Math.max(0, d[i + 1] + n));
      d[i + 2] = Math.min(255, Math.max(0, d[i + 2] + n));
    }
    ctx.putImageData(datos, sx, sy);
  }

  // Cuando el margen de arriba se sale de los límites reales de "fuente"
  // (pasa siempre que sx,sy,sw,sh es la imagen ENTERA, o cerca de su borde:
  // ahí no hay nada más que recortar), ese excedente quedaba transparente en
  // canvasChico/base — el desenfoque mezclaba esa "nada" con los píxeles
  // reales cercanos y oscurecía/atenuaba una franja del ancho del margen en
  // todo el borde de la imagen (mucho más visible en el generador de
  // stickers, donde el margen puede llegar a ser casi la mitad del lado del
  // sticker). Esto estira el borde real hacia ese excedente en vez de
  // dejarlo vacío, sin tocar la parte del margen que sí cae dentro de
  // "fuente" (ahí sigue siendo contexto real, como antes — por eso
  // censurar-imagen.html, que sí suele tener ese contexto real alrededor de
  // la zona a tapar, no cambia de resultado).
  function extenderBordes(ctxChico, anchoFuente, altoFuente, sx, sy, sw, sh, margen) {
    const c = ctxChico.canvas;
    const faltaIzq = Math.max(0, margen - sx);
    const faltaDer = Math.max(0, (sx + sw + margen) - anchoFuente);
    const faltaArriba = Math.max(0, margen - sy);
    const faltaAbajo = Math.max(0, (sy + sh + margen) - altoFuente);
    if (faltaIzq > 0) ctxChico.drawImage(c, faltaIzq, 0, 1, c.height, 0, 0, faltaIzq, c.height);
    if (faltaDer > 0) ctxChico.drawImage(c, c.width - faltaDer - 1, 0, 1, c.height, c.width - faltaDer, 0, faltaDer, c.height);
    if (faltaArriba > 0) ctxChico.drawImage(c, 0, faltaArriba, c.width, 1, 0, 0, c.width, faltaArriba);
    if (faltaAbajo > 0) ctxChico.drawImage(c, 0, c.height - faltaAbajo - 1, c.width, 1, 0, c.height - faltaAbajo, c.width, faltaAbajo);
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
    extenderBordes(ctxChico, fuente.width, fuente.height, sx, sy, sw, sh, margen);

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
    extenderBordes(ctxChico, fuente.width, fuente.height, sx, sy, sw, sh, margen);

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
    extenderBordes(ctxBase, fuente.width, fuente.height, sx, sy, sw, sh, margen);

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

  const EN_LUGAR = { pixelar: pixelar, 'pixelar-circulos': pixelarCirculos, ruido: ruido, grano: grano };
  const CON_FUENTE = { direccional: direccional, esmerilado: esmerilado };

  // ¿Es un estilo de efectos-mosaico.js? (deja transparente lo que no es figura)
  function esMosaico(nombre) {
    return !!(window.EfectosMosaico && EfectosMosaico.TIPOS.includes(nombre));
  }

  function aplicarRegion(nombre, ctxDestino, fuente, sx, sy, sw, sh, intensidad, opciones) {
    if (esMosaico(nombre)) {
      if (fuente && fuente !== ctxDestino.canvas) {
        ctxDestino.drawImage(fuente, sx, sy, sw, sh, sx, sy, sw, sh);
      }
      EfectosMosaico.aplicar(nombre, ctxDestino, sx, sy, sw, sh, opciones);
      return;
    }
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

  function aplicarCompleto(nombre, ctxDestino, fuente, ancho, alto, intensidad, opciones) {
    aplicarRegion(nombre, ctxDestino, fuente, 0, 0, ancho, alto, intensidad, opciones);
  }

  // --- Galería de texturas -------------------------------------------------
  // La arman igual los dos editores y el generador de stickers: antes los
  // botones y sus nombres estaban escritos a mano en cada página. Los nombres
  // viven acá (se suman a los textos de la página, que mandan); los de
  // efectos-mosaico.js los trae ese módulo con la clave 'textura-' + tipo.
  // "ruido" (pixela + ruido) no se ofrece acá a propósito: es solo para
  // censurar-imagen.html. La textura "Ruido" de la galería usa grano() (ruido
  // de grano fino, sin pixelar) — mismo texto, mismo lugar, otra función.
  const CLAVES_NOMBRE = {
    ninguno: 'textura-ninguno', pixelar: 'textura-pixelar', 'pixelar-circulos': 'textura-puntos',
    grano: 'textura-ruido', difuminar: 'textura-difuminar', direccional: 'textura-movimiento', esmerilado: 'textura-esmerilado',
  };
  const TEXTOS = {
    es: {
      'textura-ninguno': 'Ninguna', 'textura-pixelar': 'Pixelado', 'textura-puntos': 'Puntos', 'textura-ruido': 'Ruido',
      'textura-difuminar': 'Difuminado', 'textura-movimiento': 'Movimiento', 'textura-esmerilado': 'Esmerilado',
    },
    en: {
      'textura-ninguno': 'None', 'textura-pixelar': 'Pixelate', 'textura-puntos': 'Dots', 'textura-ruido': 'Noise',
      'textura-difuminar': 'Blur', 'textura-movimiento': 'Motion', 'textura-esmerilado': 'Frosted',
    },
  };
  if (window.i18nPagina) {
    for (const idioma of Object.keys(TEXTOS)) window.i18nPagina[idioma] = Object.assign({}, TEXTOS[idioma], window.i18nPagina[idioma]);
  }

  // Llena "galeria" con un botón por textura (clase miniatura-textura, con
  // un <canvas> para la miniatura y data-textura = tipo): "Ninguna" (activo),
  // las de este archivo y, si la página cargó efectos-mosaico.js, las suyas.
  // Hay que llamarla ANTES de conectar los clics de la galería.
  function crearGaleria(galeria) {
    const texto = (clave) => (window.t ? window.t(clave) : TEXTOS.es[clave] || clave);
    const tipos = Object.entries(CLAVES_NOMBRE);
    if (window.EfectosMosaico) for (const tipo of EfectosMosaico.TIPOS) tipos.push([tipo, 'textura-' + tipo]);
    galeria.textContent = '';
    for (const [tipo, clave] of tipos) {
      const boton = document.createElement('button');
      boton.type = 'button';
      boton.className = 'miniatura-textura' + (tipo === 'ninguno' ? ' activo' : '');
      boton.dataset.textura = tipo;
      boton.title = texto(clave);
      boton.setAttribute('data-i18n-attr', 'title:' + clave);
      const nombre = document.createElement('span');
      nombre.dataset.i18n = clave;
      nombre.textContent = texto(clave);
      boton.append(document.createElement('canvas'), nombre);
      galeria.append(boton);
    }
  }

  return {
    pixelar, pixelarCirculos, ruido, grano, difuminar, direccional, esmerilado,
    aplicarRegion, aplicarCompleto, esMosaico, crearGaleria,
    NOMBRES: ['pixelar', 'pixelar-circulos', 'grano', 'difuminar', 'direccional', 'esmerilado'],
    REFERENCIA_TEXTURA,
  };
})();

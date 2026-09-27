/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Motor de filtros de color, extraído de Editor-de-color-de-imagen.html para
// poder aplicar los mismos ~90 looks de assets/js/filtros-color.js fuera de
// los dos editores grandes (ver herramientas/generador-stickers.html).
// procesarCanvas/renderizarFiltro no tocan el canal alfa del canvas: un
// recorte con fondo transparente sigue transparente después de aplicarlos.
// Necesita, cargados antes como <script>: ajuste-por-color.js,
// efectos-textura.js, filtros-color.js.
window.ProcesarFiltros = (function () {
  const CLAVES_MEZCLA_BN = AjustePorColor.RANGOS.map((r) => 'mezcla' + r.clave[0].toUpperCase() + r.clave.slice(1));

  // Valores neutros de cada ajuste — cada preset parte de esta base y sólo
  // declara en qué se diferencia (ver assets/js/filtros-color.js).
  const AJUSTES_NEUTROS = {
    bn: false, exposicion: 0, contraste: 0, sombras: 0, reflejos: 0, saturacion: 0, intensidad: 0, tono: 0,
    colorRojo: 0, colorVerde: 0, colorAzul: 0,
    canalRojo: 100, canalVerde: 100, canalAzul: 100, claridad: 0, grano: 0, sepia: 0, invertir: 0, desenfoque: 0,
    dispersionIntensidad: 0, dispersionMatiz: 0, dispersionTamano: 50,
    aberracionIntensidad: 0, aberracionAngulo: 0, aberracionMatiz: 0, intensidadFiltro: 100,
    hsl: null,
  };
  CLAVES_MEZCLA_BN.forEach((clave) => { AJUSTES_NEUTROS[clave] = 0; });

  function crearPreset(cambios, extra = {}) {
    return { ...AJUSTES_NEUTROS, ...cambios, capasExtra: extra.capas || [], textura: extra.textura || null };
  }

  const PRESETS = Object.fromEntries(Object.entries(FiltrosColor.FILTROS).map(([clave, f]) =>
    [clave, crearPreset(f.cambios, { capas: f.capas, textura: f.textura })]));

  const SIN_TEXTURA = { tipo: 'ninguno', intensidad: 18, orden: 'debajo' };

  // Matriz de rotación de tono, misma fórmula que usan los navegadores para
  // el filtro CSS hue-rotate (feColorMatrix "hueRotate" de SVG/CSS Filter
  // Effects), para poder combinarla en un único paso por píxel junto al resto.
  function matrizTono(gradosGrados) {
    const a = gradosGrados * Math.PI / 180;
    const c = Math.cos(a), s = Math.sin(a);
    return [
      0.213 + c * 0.787 - s * 0.213, 0.715 - c * 0.715 - s * 0.715, 0.072 - c * 0.072 + s * 0.928,
      0.213 - c * 0.213 + s * 0.143, 0.715 + c * 0.285 + s * 0.140, 0.072 - c * 0.072 - s * 0.283,
      0.213 - c * 0.213 - s * 0.787, 0.715 - c * 0.715 + s * 0.715, 0.072 + c * 0.928 + s * 0.072,
    ];
  }

  // Truco barato de desenfoque: reescalar la imagen a un tamaño muy chico y
  // volver a agrandarla usando el suavizado nativo del canvas — sirve como
  // aproximación de un desenfoque gaussiano fuerte, sin implementar
  // convolución a mano. Se usa como referencia de "detalle local" para la
  // claridad (unsharp mask) y como base del resplandor de la dispersión de luz.
  function generarVersionBorrosa(canvasFuente, ladoChico = 40) {
    const ancho = canvasFuente.width, alto = canvasFuente.height;
    const anchoChico = Math.max(1, Math.round(ancho * ladoChico / Math.max(ancho, alto)));
    const altoChico = Math.max(1, Math.round(alto * ladoChico / Math.max(ancho, alto)));

    const canvasChico = document.createElement('canvas');
    canvasChico.width = anchoChico;
    canvasChico.height = altoChico;
    const ctxChico = canvasChico.getContext('2d');
    ctxChico.imageSmoothingEnabled = true;
    ctxChico.imageSmoothingQuality = 'high';
    ctxChico.drawImage(canvasFuente, 0, 0, anchoChico, altoChico);

    const canvasBorroso = document.createElement('canvas');
    canvasBorroso.width = ancho;
    canvasBorroso.height = alto;
    const ctxBorroso = canvasBorroso.getContext('2d');
    ctxBorroso.imageSmoothingEnabled = true;
    ctxBorroso.imageSmoothingQuality = 'high';
    ctxBorroso.drawImage(canvasChico, 0, 0, anchoChico, altoChico, 0, 0, ancho, alto);
    return ctxBorroso.getImageData(0, 0, ancho, alto).data;
  }

  // Aplica todo el pipeline de ajustes sobre el canvas recibido (que ya debe
  // contener la imagen sin procesar) y lo deja mutado con el resultado. Solo
  // toca r,g,b de cada píxel — nunca el canal alfa (datos[i+3]).
  function procesarCanvas(canvasFuente, ajustes) {
    const ctx = canvasFuente.getContext('2d');
    const ancho = canvasFuente.width, alto = canvasFuente.height;
    const imageData = ctx.getImageData(0, 0, ancho, alto);
    const datos = imageData.data;

    // Dispersión de luz (separación de canales / aberración cromática): a
    // diferencia del resto de ajustes, que son por-píxel, este es geométrico
    // — desplaza el canal rojo hacia un lado y el azul hacia el opuesto,
    // dejando el verde en su lugar, igual que la dispersión óptica de un
    // prisma. Por eso corre antes que todo lo demás, sobre una copia del
    // original.
    const ampAberracion = ajustes.aberracionIntensidad / 100;
    if (ampAberracion > 0) {
      const anguloRad = (ajustes.aberracionAngulo / 180) * Math.PI;
      const desplazamiento = ampAberracion * Math.max(2, Math.round(Math.min(ancho, alto) * 0.02));
      const dx = Math.round(Math.cos(anguloRad) * desplazamiento);
      const dy = Math.round(Math.sin(anguloRad) * desplazamiento);
      const original = new Uint8ClampedArray(datos);
      const mTonoAberracion = ajustes.aberracionMatiz !== 0 ? matrizTono(ajustes.aberracionMatiz) : null;
      for (let y = 0; y < alto; y++) {
        const yr = y + dy < 0 ? 0 : y + dy >= alto ? alto - 1 : y + dy;
        const yb = y - dy < 0 ? 0 : y - dy >= alto ? alto - 1 : y - dy;
        for (let x = 0; x < ancho; x++) {
          const xr = x + dx < 0 ? 0 : x + dx >= ancho ? ancho - 1 : x + dx;
          const xb = x - dx < 0 ? 0 : x - dx >= ancho ? ancho - 1 : x - dx;
          const iDestino = (y * ancho + x) * 4;
          const iR = (yr * ancho + xr) * 4;
          const iB = (yb * ancho + xb) * 4;
          if (mTonoAberracion) {
            const rr = original[iR], rg = original[iR + 1], rb = original[iR + 2];
            const br = original[iB], bg = original[iB + 1], bb = original[iB + 2];
            datos[iDestino] = mTonoAberracion[0] * rr + mTonoAberracion[1] * rg + mTonoAberracion[2] * rb;
            datos[iDestino + 2] = mTonoAberracion[6] * br + mTonoAberracion[7] * bg + mTonoAberracion[8] * bb;
          } else {
            datos[iDestino] = original[iR];
            datos[iDestino + 2] = original[iB + 2];
          }
        }
      }
    }

    const factorExposicion = Math.pow(2, (ajustes.exposicion / 100) * 2);
    const c = ajustes.contraste;
    const factorContraste = (259 * (c + 255)) / (255 * (259 - c));
    const amplitudSombras = (ajustes.sombras / 100) * 70;
    const amplitudReflejos = (ajustes.reflejos / 100) * 70;
    const factorSaturacion = 1 + ajustes.saturacion / 100;
    const nIntensidad = ajustes.intensidad / 100;
    const necesitaTono = ajustes.tono !== 0;
    const mTono = necesitaTono ? matrizTono(ajustes.tono) : null;
    const deltaColorRojo = ajustes.colorRojo;
    const deltaColorVerde = ajustes.colorVerde;
    const deltaColorAzul = ajustes.colorAzul;
    const bn = ajustes.bn;
    const wr = ajustes.canalRojo / 100, wg = ajustes.canalVerde / 100, wb = ajustes.canalAzul / 100;
    const sumaCanales = (wr + wg + wb) || 1;
    const mezclaGris = AjustePorColor.prepararMezclaGris(CLAVES_MEZCLA_BN.map((clave) => ajustes[clave] || 0));
    const refColor = (AjustePorColor.preparar(ajustes.hsl) || mezclaGris) ? AjustePorColor.referenciaColor(canvasFuente) : null;
    const ampSepia = ajustes.sepia / 100;
    const ampInvertir = ajustes.invertir / 100;
    const ampClaridad = (ajustes.claridad / 100) * 1.3;
    const necesitaClaridad = ampClaridad !== 0;
    const borrosos = necesitaClaridad ? generarVersionBorrosa(canvasFuente) : null;
    const ampGrano = (ajustes.grano / 100) * 45;
    const ajustePorColor = AjustePorColor.preparar(ajustes.hsl);
    const factorIntensidadFiltro = ajustes.intensidadFiltro / 100;

    const ampDispersion = (ajustes.dispersionIntensidad / 100) * 1.6;
    let capaDispersion = null;
    let mMatizDispersion = null;
    if (ampDispersion > 0) {
      const brillante = new Uint8ClampedArray(datos.length);
      for (let i = 0; i < datos.length; i += 4) {
        const lum = 0.299 * datos[i] + 0.587 * datos[i + 1] + 0.114 * datos[i + 2];
        const factor = Math.max(0, Math.min(1, (lum - 170) / 60));
        brillante[i] = datos[i] * factor;
        brillante[i + 1] = datos[i + 1] * factor;
        brillante[i + 2] = datos[i + 2] * factor;
        brillante[i + 3] = 255;
      }
      const canvasBrillante = document.createElement('canvas');
      canvasBrillante.width = ancho;
      canvasBrillante.height = alto;
      canvasBrillante.getContext('2d').putImageData(new ImageData(brillante, ancho, alto), 0, 0);
      const ladoChicoDispersion = Math.round(140 - (ajustes.dispersionTamano / 100) * 130);
      capaDispersion = generarVersionBorrosa(canvasBrillante, ladoChicoDispersion);
      if (ajustes.dispersionMatiz !== 0) mMatizDispersion = matrizTono(ajustes.dispersionMatiz);
    }

    for (let i = 0; i < datos.length; i += 4) {
      let r = datos[i], g = datos[i + 1], b = datos[i + 2];
      const rOriginal = r, gOriginal = g, bOriginal = b;

      r *= factorExposicion; g *= factorExposicion; b *= factorExposicion;
      if (r < 0) r = 0; else if (r > 255) r = 255;
      if (g < 0) g = 0; else if (g > 255) g = 255;
      if (b < 0) b = 0; else if (b > 255) b = 255;

      r = factorContraste * (r - 128) + 128;
      g = factorContraste * (g - 128) + 128;
      b = factorContraste * (b - 128) + 128;
      if (r < 0) r = 0; else if (r > 255) r = 255;
      if (g < 0) g = 0; else if (g > 255) g = 255;
      if (b < 0) b = 0; else if (b > 255) b = 255;

      if (amplitudSombras !== 0 || amplitudReflejos !== 0) {
        const lum = 0.299 * r + 0.587 * g + 0.114 * b;
        const mascaraSombra = Math.max(0, Math.min(1, 1 - lum / 128));
        const mascaraReflejo = Math.max(0, Math.min(1, (lum - 128) / 127));
        const delta = amplitudSombras * mascaraSombra + amplitudReflejos * mascaraReflejo;
        r += delta; g += delta; b += delta;
        if (r < 0) r = 0; else if (r > 255) r = 255;
        if (g < 0) g = 0; else if (g > 255) g = 255;
        if (b < 0) b = 0; else if (b > 255) b = 255;
      }

      if (mTono) {
        const nr = mTono[0] * r + mTono[1] * g + mTono[2] * b;
        const ng = mTono[3] * r + mTono[4] * g + mTono[5] * b;
        const nb = mTono[6] * r + mTono[7] * g + mTono[8] * b;
        r = nr; g = ng; b = nb;
        if (r < 0) r = 0; else if (r > 255) r = 255;
        if (g < 0) g = 0; else if (g > 255) g = 255;
        if (b < 0) b = 0; else if (b > 255) b = 255;
      }

      if (factorSaturacion !== 1 || nIntensidad !== 0) {
        const lum = 0.299 * r + 0.587 * g + 0.114 * b;
        const maxC = Math.max(r, g, b), minC = Math.min(r, g, b);
        const saturacionActual = maxC > 0 ? (maxC - minC) / maxC : 0;
        const factorVibrance = 1 + nIntensidad * (1 - saturacionActual) * 1.5;
        const factorTotal = factorSaturacion * factorVibrance;
        r = lum + (r - lum) * factorTotal;
        g = lum + (g - lum) * factorTotal;
        b = lum + (b - lum) * factorTotal;
        if (r < 0) r = 0; else if (r > 255) r = 255;
        if (g < 0) g = 0; else if (g > 255) g = 255;
        if (b < 0) b = 0; else if (b > 255) b = 255;
      }

      if (ajustePorColor) {
        [r, g, b] = refColor
          ? AjustePorColor.aplicar(r, g, b, ajustePorColor, refColor[i], refColor[i + 1], refColor[i + 2])
          : AjustePorColor.aplicar(r, g, b, ajustePorColor);
      }

      if (deltaColorRojo !== 0) { r += deltaColorRojo; if (r < 0) r = 0; else if (r > 255) r = 255; }
      if (deltaColorVerde !== 0) { g += deltaColorVerde; if (g < 0) g = 0; else if (g > 255) g = 255; }
      if (deltaColorAzul !== 0) { b += deltaColorAzul; if (b < 0) b = 0; else if (b > 255) b = 255; }

      if (bn) {
        let gris = (r * wr + g * wg + b * wb) / sumaCanales;
        if (mezclaGris) gris = AjustePorColor.mezclarGris(r, g, b, gris, mezclaGris, refColor[i], refColor[i + 1], refColor[i + 2]);
        r = gris; g = gris; b = gris;
      } else if (mezclaGris) {
        const lum = 0.299 * r + 0.587 * g + 0.114 * b;
        if (lum > 0) {
          const factor = AjustePorColor.mezclarGris(r, g, b, lum, mezclaGris, refColor[i], refColor[i + 1], refColor[i + 2]) / lum;
          r *= factor; g *= factor; b *= factor;
          if (r > 255) r = 255;
          if (g > 255) g = 255;
          if (b > 255) b = 255;
        }
      }

      if (ampSepia !== 0) {
        const tr = 0.393 * r + 0.769 * g + 0.189 * b;
        const tg = 0.349 * r + 0.686 * g + 0.168 * b;
        const tb = 0.272 * r + 0.534 * g + 0.131 * b;
        r += (tr - r) * ampSepia; g += (tg - g) * ampSepia; b += (tb - b) * ampSepia;
        if (r < 0) r = 0; else if (r > 255) r = 255;
        if (g < 0) g = 0; else if (g > 255) g = 255;
        if (b < 0) b = 0; else if (b > 255) b = 255;
      }

      if (ampInvertir !== 0) {
        r += (255 - r - r) * ampInvertir;
        g += (255 - g - g) * ampInvertir;
        b += (255 - b - b) * ampInvertir;
      }

      if (capaDispersion) {
        let br = capaDispersion[i], bg = capaDispersion[i + 1], bb = capaDispersion[i + 2];
        if (mMatizDispersion) {
          const nr = mMatizDispersion[0] * br + mMatizDispersion[1] * bg + mMatizDispersion[2] * bb;
          const ng = mMatizDispersion[3] * br + mMatizDispersion[4] * bg + mMatizDispersion[5] * bb;
          const nb = mMatizDispersion[6] * br + mMatizDispersion[7] * bg + mMatizDispersion[8] * bb;
          br = nr < 0 ? 0 : nr > 255 ? 255 : nr;
          bg = ng < 0 ? 0 : ng > 255 ? 255 : ng;
          bb = nb < 0 ? 0 : nb > 255 ? 255 : nb;
        }
        const pantallaR = 255 - (255 - r) * (255 - br) / 255;
        const pantallaG = 255 - (255 - g) * (255 - bg) / 255;
        const pantallaB = 255 - (255 - b) * (255 - bb) / 255;
        r += (pantallaR - r) * ampDispersion;
        g += (pantallaG - g) * ampDispersion;
        b += (pantallaB - b) * ampDispersion;
        if (r < 0) r = 0; else if (r > 255) r = 255;
        if (g < 0) g = 0; else if (g > 255) g = 255;
        if (b < 0) b = 0; else if (b > 255) b = 255;
      }

      if (borrosos) {
        const lumOriginal = 0.299 * rOriginal + 0.587 * gOriginal + 0.114 * bOriginal;
        const lumBorrosa = 0.299 * borrosos[i] + 0.587 * borrosos[i + 1] + 0.114 * borrosos[i + 2];
        const delta = (lumOriginal - lumBorrosa) * ampClaridad;
        r += delta; g += delta; b += delta;
      }

      if (ampGrano !== 0) {
        const ruido = (Math.random() - 0.5) * ampGrano;
        r += ruido; g += ruido; b += ruido;
      }

      if (factorIntensidadFiltro !== 1) {
        r = rOriginal + (r - rOriginal) * factorIntensidadFiltro;
        g = gOriginal + (g - gOriginal) * factorIntensidadFiltro;
        b = bOriginal + (b - bOriginal) * factorIntensidadFiltro;
      }

      datos[i] = r < 0 ? 0 : r > 255 ? 255 : r;
      datos[i + 1] = g < 0 ? 0 : g > 255 ? 255 : g;
      datos[i + 2] = b < 0 ? 0 : b > 255 ? 255 : b;
    }

    ctx.putImageData(imageData, 0, 0);
    return canvasFuente;
  }

  function copiarCanvas(canvas) {
    const copia = document.createElement('canvas');
    copia.width = canvas.width; copia.height = canvas.height;
    copia.getContext('2d').drawImage(canvas, 0, 0);
    return copia;
  }

  // Pipeline completo de un filtro sobre "canvas" (que ya tiene la imagen
  // dibujada): textura "debajo" → ajustes → capas extra → mezcla con
  // "Intensidad del filtro" → textura "encima" → desenfoque. "escala" adapta
  // las medidas pensadas en píxeles de la imagen real (textura, desenfoque)
  // a un canvas más chico (usar 1 para tamaño completo).
  function renderizarFiltro(canvas, ajustes, capasExtra, textura, escala) {
    const ctx = canvas.getContext('2d');
    const intensidadTextura = Math.max(1, textura.intensidad * escala);
    if (textura.tipo !== 'ninguno' && textura.orden === 'debajo') {
      EfectosTextura.aplicarCompleto(textura.tipo, ctx, undefined, canvas.width, canvas.height, intensidadTextura);
    }
    const base = copiarCanvas(canvas);

    procesarCanvas(canvas, { ...ajustes, intensidadFiltro: 100 });

    FiltrosColor.aplicarCapasExtra(canvas, base, capasExtra, (c, aj) => procesarCanvas(c, { ...AJUSTES_NEUTROS, ...aj }));

    const factorIntensidad = ajustes.intensidadFiltro / 100;
    if (factorIntensidad < 1) {
      const resultadoCompleto = copiarCanvas(canvas);
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(base, 0, 0);
      ctx.globalAlpha = factorIntensidad;
      ctx.drawImage(resultadoCompleto, 0, 0);
      ctx.globalAlpha = 1;
    }

    if (textura.tipo !== 'ninguno' && textura.orden === 'encima') {
      EfectosTextura.aplicarCompleto(textura.tipo, ctx, undefined, canvas.width, canvas.height, intensidadTextura);
    }

    if (ajustes.desenfoque > 0) {
      const sinDesenfoque = copiarCanvas(canvas);
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      ctx.filter = `blur(${ajustes.desenfoque * escala}px)`;
      ctx.drawImage(sinDesenfoque, 0, 0);
      ctx.filter = 'none';
    }
    return canvas;
  }

  return { AJUSTES_NEUTROS, PRESETS, SIN_TEXTURA, crearPreset, procesarCanvas, copiarCanvas, renderizarFiltro };
})();

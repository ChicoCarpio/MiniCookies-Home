/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Motor de filtros de color: el único. Lo usan el Editor de color, el Editor
// avanzado (una vez por capa) y el generador de stickers. Antes cada editor
// tenía su copia y se separaron (la caché de la vista previa estaba en uno
// solo, la mezcla de "Intensidad del filtro" era distinta en cada uno):
// cualquier cambio al pipeline va acá, no en las páginas.
// procesarCanvas no toca el canal alfa; renderizarFiltro tampoco, salvo lo
// que hagan la textura elegida o las capas extra del filtro: un recorte con
// fondo transparente sigue transparente después de aplicarlos.
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
    // 2026-10-04: balance de blancos, puntos blanco y negro, nitidez y viñeta.
    temperatura: 0, matizBlancos: 0, blancos: 0, negros: 0, nitidez: 0, vineta: 0,
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
  function generarVersionBorrosa(canvasFuente, ladoChico = 40, scratch = null) {
    const ancho = canvasFuente.width, alto = canvasFuente.height;
    const anchoChico = Math.max(1, Math.round(ancho * ladoChico / Math.max(ancho, alto)));
    const altoChico = Math.max(1, Math.round(alto * ladoChico / Math.max(ancho, alto)));

    const canvasChico = scratch ? scratch.chico : document.createElement('canvas');
    canvasChico.width = anchoChico;
    canvasChico.height = altoChico;
    const ctxChico = canvasChico.getContext('2d');
    ctxChico.imageSmoothingEnabled = true;
    ctxChico.imageSmoothingQuality = 'high';
    ctxChico.drawImage(canvasFuente, 0, 0, anchoChico, altoChico);

    const canvasBorroso = scratch ? scratch.borroso : document.createElement('canvas');
    canvasBorroso.width = ancho;
    canvasBorroso.height = alto;
    const ctxBorroso = canvasBorroso.getContext('2d');
    ctxBorroso.imageSmoothingEnabled = true;
    ctxBorroso.imageSmoothingQuality = 'high';
    ctxBorroso.drawImage(canvasChico, 0, 0, anchoChico, altoChico, 0, 0, ancho, alto);
    return ctxBorroso.getImageData(0, 0, ancho, alto).data;
  }

  // "memo" (opcional) guarda lo que sale de la imagen de entrada y no de los
  // deslizadores —el desenfoque de Claridad, las zonas brillantes de la
  // dispersión y su desenfoque, que solo depende de "Tamaño"—, para no
  // recalcularlo en cada fotograma mientras se arrastra otro control. Quien
  // llama garantiza que el memo es de esta misma imagen de entrada (lo vacía
  // con {} cuando cambia); acá solo se controla el tamaño.
  function memoDe(memo, ancho, alto) {
    if (!memo) return null;
    if (memo.ancho !== ancho || memo.alto !== alto) {
      memo.ancho = ancho; memo.alto = alto;
      memo.claridad = null; memo.nitidez = null; memo.dispersionBrillante = null; memo.dispersionBlur = null;
    }
    return memo;
  }

  // Aplica todo el pipeline de ajustes sobre el canvas recibido (que ya debe
  // contener la imagen sin procesar) y lo deja mutado con el resultado. Solo
  // toca r,g,b de cada píxel — nunca el canal alfa (datos[i+3]). "scratch":
  // canvases de repuesto (crearScratch); "memo": ver memoDe.
  function procesarCanvas(canvasFuente, ajustes, scratch = null, memo = null) {
    const ctx = canvasFuente.getContext('2d');
    const ancho = canvasFuente.width, alto = canvasFuente.height;
    const imageData = ctx.getImageData(0, 0, ancho, alto);
    const datos = imageData.data;
    memo = memoDe(memo, ancho, alto);

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
    // Desenfoque de referencia de Claridad: sale de la imagen sin procesar,
    // con tamaño fijo; no depende de ningún deslizador.
    let borrosos = null;
    if (necesitaClaridad) {
      borrosos = memo && memo.claridad;
      if (!borrosos) {
        borrosos = generarVersionBorrosa(canvasFuente, 40, scratch ? scratch.scratchBorrosoClaridad : null);
        if (memo) memo.claridad = borrosos;
      }
    }
    // Ajustes del 2026-10-04. Se leen con "|| 0": un proyecto o un filtro de
    // antes no los trae, y un undefined volvería NaN cada píxel.
    // Balance de blancos: Temperatura (frío ↔ cálido) y Matiz (verde ↔
    // magenta), como factores por canal antes de todo lo demás.
    const nTemperatura = (ajustes.temperatura || 0) / 100;
    const nMatizBlancos = (ajustes.matizBlancos || 0) / 100;
    const necesitaBalance = nTemperatura !== 0 || nMatizBlancos !== 0;
    const kR = 1 + 0.25 * nTemperatura + 0.08 * nMatizBlancos;
    const kG = 1 - 0.18 * nMatizBlancos;
    const kB = 1 - 0.25 * nTemperatura + 0.08 * nMatizBlancos;
    // Blancos y Negros: como Sombras y Reflejos, pero solo en los extremos
    // (mueven el punto blanco y el punto negro).
    const ampBlancos = ((ajustes.blancos || 0) / 100) * 60;
    const ampNegros = ((ajustes.negros || 0) / 100) * 60;
    // Nitidez: máscara de enfoque de radio chico (Claridad es la de radio
    // grande). El desenfoque de referencia solo depende de la imagen.
    const ampNitidez = ((ajustes.nitidez || 0) / 100) * 2.5;
    let finos = null;
    if (ampNitidez !== 0) {
      finos = memo && memo.nitidez;
      if (!finos) {
        finos = generarVersionBorrosa(canvasFuente, Math.max(8, Math.round(Math.max(ancho, alto) / 3)), scratch ? scratch.scratchBorrosoClaridad : null);
        if (memo) memo.nitidez = finos;
      }
    }
    // Viñeta: negativa oscurece los bordes, positiva los aclara.
    const ampVineta = (ajustes.vineta || 0) / 100;
    const centroX = (ancho - 1) / 2, centroY = (alto - 1) / 2;

    const ampGrano = (ajustes.grano / 100) * 45;
    const ajustePorColor = AjustePorColor.preparar(ajustes.hsl);
    const factorIntensidadFiltro = ajustes.intensidadFiltro / 100;

    const ampDispersion = (ajustes.dispersionIntensidad / 100) * 1.6;
    let capaDispersion = null;
    let mMatizDispersion = null;
    if (ampDispersion > 0) {
      // Dispersión de luz (bloom): las zonas brillantes del original (no
      // dependen de ningún deslizador), muy desenfocadas (eso sí depende de
      // "Tamaño"), se mezclan en modo "screen" más abajo.
      let brillante = memo && memo.dispersionBrillante;
      if (!brillante) {
        brillante = new Uint8ClampedArray(datos.length);
        for (let i = 0; i < datos.length; i += 4) {
          const lum = 0.299 * datos[i] + 0.587 * datos[i + 1] + 0.114 * datos[i + 2];
          const factor = Math.max(0, Math.min(1, (lum - 170) / 60));
          brillante[i] = datos[i] * factor;
          brillante[i + 1] = datos[i + 1] * factor;
          brillante[i + 2] = datos[i + 2] * factor;
          brillante[i + 3] = 255;
        }
        if (memo) memo.dispersionBrillante = brillante;
      }
      const ladoChicoDispersion = Math.round(140 - (ajustes.dispersionTamano / 100) * 130);
      if (memo && memo.dispersionBlur && memo.dispersionBlur.tamano === ladoChicoDispersion) {
        capaDispersion = memo.dispersionBlur.datos;
      } else {
        const canvasBrillante = scratch ? scratch.canvasBrillante : document.createElement('canvas');
        canvasBrillante.width = ancho;
        canvasBrillante.height = alto;
        canvasBrillante.getContext('2d').putImageData(new ImageData(brillante, ancho, alto), 0, 0);
        capaDispersion = generarVersionBorrosa(canvasBrillante, ladoChicoDispersion, scratch ? scratch.scratchBorrosoDispersion : null);
        if (memo) memo.dispersionBlur = { tamano: ladoChicoDispersion, datos: capaDispersion };
      }
      if (ajustes.dispersionMatiz !== 0) mMatizDispersion = matrizTono(ajustes.dispersionMatiz);
    }

    for (let i = 0; i < datos.length; i += 4) {
      let r = datos[i], g = datos[i + 1], b = datos[i + 2];
      const rOriginal = r, gOriginal = g, bOriginal = b;

      r *= factorExposicion; g *= factorExposicion; b *= factorExposicion;
      if (necesitaBalance) { r *= kR; g *= kG; b *= kB; }
      if (r < 0) r = 0; else if (r > 255) r = 255;
      if (g < 0) g = 0; else if (g > 255) g = 255;
      if (b < 0) b = 0; else if (b > 255) b = 255;

      r = factorContraste * (r - 128) + 128;
      g = factorContraste * (g - 128) + 128;
      b = factorContraste * (b - 128) + 128;
      if (r < 0) r = 0; else if (r > 255) r = 255;
      if (g < 0) g = 0; else if (g > 255) g = 255;
      if (b < 0) b = 0; else if (b > 255) b = 255;

      if (amplitudSombras !== 0 || amplitudReflejos !== 0 || ampBlancos !== 0 || ampNegros !== 0) {
        const lum = 0.299 * r + 0.587 * g + 0.114 * b;
        const mascaraSombra = Math.max(0, Math.min(1, 1 - lum / 128));
        const mascaraReflejo = Math.max(0, Math.min(1, (lum - 128) / 127));
        const extremoBlanco = Math.max(0, (lum - 170) / 85), extremoNegro = Math.max(0, (85 - lum) / 85);
        const delta = amplitudSombras * mascaraSombra + amplitudReflejos * mascaraReflejo
          + ampBlancos * extremoBlanco * extremoBlanco + ampNegros * extremoNegro * extremoNegro;
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

      if (finos) {
        const lumOriginal = 0.299 * rOriginal + 0.587 * gOriginal + 0.114 * bOriginal;
        const lumFina = 0.299 * finos[i] + 0.587 * finos[i + 1] + 0.114 * finos[i + 2];
        const delta = (lumOriginal - lumFina) * ampNitidez;
        r += delta; g += delta; b += delta;
      }

      if (ampVineta !== 0) {
        const p = i >> 2;
        const x = p % ancho, y = (p - x) / ancho;
        const dx = (x - centroX) / (centroX || 1), dy = (y - centroY) / (centroY || 1);
        // 0 en el centro, 1 en las esquinas; empieza a notarse a mitad de camino.
        let t = (Math.sqrt(dx * dx + dy * dy) / Math.SQRT2 - 0.35) / 0.65;
        if (t > 0) {
          if (t > 1) t = 1;
          // Al máximo, 85%: la esquina conserva algo de detalle.
          const m = t * t * (3 - 2 * t) * ampVineta * 0.85;
          if (m < 0) { r *= 1 + m; g *= 1 + m; b *= 1 + m; }
          else { r += (255 - r) * m; g += (255 - g) * m; b += (255 - b) * m; }
        }
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

  function copiarCanvas(canvas, destino = null) {
    const copia = destino || document.createElement('canvas');
    copia.width = canvas.width; copia.height = canvas.height;
    copia.getContext('2d').drawImage(canvas, 0, 0);
    return copia;
  }

  // Pipeline completo de un filtro sobre "canvas" (que ya tiene la imagen
  // dibujada), el mismo para la vista previa, la descarga y las miniaturas:
  // textura "debajo" → ajustes → capas extra → mezcla con "Intensidad del
  // filtro" → textura "encima" → desenfoque. "textura" es { tipo,
  // intensidad, orden, opciones } (opciones: las de los mosaicos, ver
  // efectos-mosaico.js). "escala" adapta el desenfoque, pensado en píxeles de
  // la imagen real, a un canvas más chico (1 para tamaño completo).
  // "scratch" y "memo": solo para la vista previa en vivo (ver crearScratch y
  // memoDe).
  function renderizarFiltro(canvas, ajustes, capasExtra, textura, escala, scratch = null, memo = null) {
    const ctx = canvas.getContext('2d');
    // Porcentaje del ancho de este canvas, no píxeles fijos — ver
    // EfectosTextura.REFERENCIA_TEXTURA: la misma posición del deslizador se
    // ve igual de fuerte en la vista previa, la miniatura y la descarga.
    const intensidadTextura = Math.max(1, textura.intensidad * (canvas.width / EfectosTextura.REFERENCIA_TEXTURA));
    if (textura.tipo !== 'ninguno' && textura.orden === 'debajo') {
      EfectosTextura.aplicarCompleto(textura.tipo, ctx, undefined, canvas.width, canvas.height, intensidadTextura, textura.opciones);
    }
    // Referencia sin ajustes, para las máscaras de las capas extra y para la
    // mezcla de intensidad. Sin ninguna de las dos no hace falta copiarla.
    const hayCapas = capasExtra && capasExtra.length > 0;
    const factorIntensidad = ajustes.intensidadFiltro / 100;
    const base = hayCapas || factorIntensidad < 1 ? copiarCanvas(canvas, scratch ? scratch.base : null) : null;

    procesarCanvas(canvas, { ...ajustes, intensidadFiltro: 100 }, scratch, memo);

    if (hayCapas) FiltrosColor.aplicarCapasExtra(canvas, base, capasExtra, (c, aj) => procesarCanvas(c, { ...AJUSTES_NEUTROS, ...aj }));

    // Mezcla píxel a píxel con el original, alfa incluido. No con globalAlpha:
    // dibujar el resultado encima de la base suma opacidad, y las zonas
    // semitransparentes de un PNG/WebP quedaban más opacas en el Editor de
    // color (alfa 128 → 153 con la intensidad en 40%).
    if (factorIntensidad < 1) {
      const original = base.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
      const resultado = ctx.getImageData(0, 0, canvas.width, canvas.height);
      const datos = resultado.data;
      for (let i = 0; i < datos.length; i++) datos[i] = original[i] + (datos[i] - original[i]) * factorIntensidad;
      ctx.putImageData(resultado, 0, 0);
    }

    if (textura.tipo !== 'ninguno' && textura.orden === 'encima') {
      EfectosTextura.aplicarCompleto(textura.tipo, ctx, undefined, canvas.width, canvas.height, intensidadTextura, textura.opciones);
    }

    if (ajustes.desenfoque > 0) {
      const sinDesenfoque = copiarCanvas(canvas, scratch ? scratch.sinDesenfoque : null);
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      ctx.filter = `blur(${ajustes.desenfoque * escala}px)`;
      ctx.drawImage(sinDesenfoque, 0, 0);
      ctx.filter = 'none';
    }
    return canvas;
  }

  // Canvases de repuesto para la vista previa en vivo: cada
  // document.createElement('canvas') es una textura de GPU que el navegador
  // reserva y suelta, y crear varias por fotograma es lo que se sentía como
  // "lag" al arrastrar un deslizador. No guardan ningún valor entre llamadas
  // (eso es el memo). Para la descarga o las miniaturas no hacen falta.
  function crearScratch() {
    return {
      base: document.createElement('canvas'),
      sinDesenfoque: document.createElement('canvas'),
      canvasBrillante: document.createElement('canvas'),
      scratchBorrosoClaridad: { chico: document.createElement('canvas'), borroso: document.createElement('canvas') },
      scratchBorrosoDispersion: { chico: document.createElement('canvas'), borroso: document.createElement('canvas') },
    };
  }

  return { AJUSTES_NEUTROS, CLAVES_MEZCLA_BN, PRESETS, SIN_TEXTURA, crearPreset, procesarCanvas, copiarCanvas, renderizarFiltro, crearScratch };
})();

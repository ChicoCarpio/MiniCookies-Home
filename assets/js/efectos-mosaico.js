/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Texturas "de manualidad", pop art y cómic, que entran por la misma puerta
// que las demás texturas (EfectosTextura.aplicarRegion, efectos-textura.js):
//
// - Mosaicos: pixel art, cuadrícula, punto de cruz, cuentas (hama beads),
//   diamantes (diamond painting), abalorios y pop art (puntos). La imagen se
//   lleva a una cuadrícula, se reduce a pocos colores (paleta-color.js) y
//   cada celda se dibuja con el "sprite" de su estilo. Lo que hace que se
//   parezcan a los de verdad es la paleta corta, no tanto la separación.
// - Cómic: suavizado Kuwahara, pocos colores, limpieza de manchas y cada zona
//   de color vectorizada (ImageTracer, vendor/imagetracerjs) para que los
//   bordes sean curvas y no escalones; la línea va solo entre colores que se
//   diferencian de verdad.
//
// No usan la "intensidad" de las demás texturas: todo sale de "opciones"
// (ver OPCIONES), medido respecto de la imagen y no en píxeles, así la vista
// previa y la descarga a tamaño completo se ven iguales.
//
// Además arma el panel con sus controles, que comparten los dos editores y
// el generador de stickers. Los botones de la galería los arma
// EfectosTextura.crearGaleria (con los nombres de TEXTOS de acá).
//
// Necesita, cargados antes: paleta-color.js y, para el cómic,
// vendor/imagetracerjs/1.2.6/imagetracer_v1.2.6.js.
window.EfectosMosaico = (function () {
  const { A_LINEAL, linealAOklab, oklabARgb, tono, dist2, masCercano, paleta: calcularPaleta } = PaletaColor;

  // Cada tipo: estilo de celda y colores por defecto. 'auto' en la
  // separación es lo propio de cada estilo (ver ESTILOS).
  const TIPOS = {
    'pixel-art': { estilo: 'liso', colores: 16 },
    cuadricula: { estilo: 'cuadricula', colores: 12 },
    'punto-cruz': { estilo: 'puntoCruz', colores: 12 },
    cuentas: { estilo: 'cuentas', colores: 12 },
    diamante: { estilo: 'diamante', colores: 14 },
    abalorios: { estilo: 'abalorio', colores: 12 },
    'pop-art': { estilo: 'popArt', colores: 8 },
    comic: { colores: 12 },
  };

  // Opciones de fábrica. Mosaicos: columnas (celdas de ancho), tramado,
  // limpiar (celdas sueltas), separacion ('auto' | color CSS) y grosor de la
  // separación (fracción de la celda, solo cuadrícula). Cómic: suavizado
  // (radio Kuwahara a la resolución de trabajo), limpieza (radio del filtro
  // de moda), detalle (lado de la resolución de trabajo, px), separacion
  // ('auto' | color | 'ninguna': la línea) y grosor de línea (px de trabajo).
  const OPCIONES = {
    mosaico: { columnas: 64, tramado: false, limpiar: true, separacion: 'auto', grosor: 0.1 },
    comic: { suavizado: 3, limpieza: 2, detalle: 640, separacion: 'auto', grosor: 1.6 },
  };
  const deFabrica = (tipo) => ({ colores: TIPOS[tipo].colores, ...OPCIONES[tipo === 'comic' ? 'comic' : 'mosaico'] });

  // --- Mosaicos --------------------------------------------------------

  // Una celda por posición: { lab, alfa }, con el color promedio en luz
  // lineal. En "ladrillo" las filas impares se corren media celda.
  function muestrear(canvas, columnas, ladrillo, realce) {
    const ancho = canvas.width, alto = canvas.height;
    const d = canvas.getContext('2d').getImageData(0, 0, ancho, alto).data;
    const lado = ancho / columnas;
    const filas = Math.max(1, Math.round(alto / lado));
    const celdas = [];
    for (let f = 0; f < filas; f++) {
      const corrida = ladrillo && f % 2 ? lado / 2 : 0;
      for (let c = 0; c < columnas; c++) {
        const x0 = Math.floor(c * lado + corrida), x1 = Math.min(ancho, Math.floor((c + 1) * lado + corrida));
        const y0 = Math.floor(f * lado), y1 = Math.min(alto, Math.floor((f + 1) * lado));
        let r = 0, g = 0, b = 0, peso = 0, n = 0;
        for (let y = y0; y < y1; y++) {
          for (let x = x0; x < x1; x++) {
            const i = (y * ancho + x) * 4, a = d[i + 3] / 255;
            r += A_LINEAL[d[i]] * a; g += A_LINEAL[d[i + 1]] * a; b += A_LINEAL[d[i + 2]] * a;
            peso += a; n++;
          }
        }
        const lab = peso ? linealAOklab(r / peso, g / peso, b / peso) : [0, 0, 0];
        lab[1] *= realce; lab[2] *= realce;
        celdas.push({ lab, alfa: n ? peso / n : 0 });
      }
    }
    return { columnas, filas, celdas };
  }

  // Color de la paleta para cada celda (-1 = vacía, sin dibujar). Con
  // tramado, una celda entre dos colores parecidos toma uno u otro según una
  // matriz de Bayer: el degradado se imita con un patrón de puntos, como en
  // el pixel art clásico. Entre colores muy distintos no se trama (sería ruido).
  const BAYER4 = [0, 8, 2, 10, 12, 4, 14, 6, 3, 11, 1, 9, 15, 7, 13, 5].map((v) => (v + 0.5) / 16);
  function asignar(malla, paleta, tramado) {
    const indice = new Int16Array(malla.celdas.length);
    malla.celdas.forEach((celda, i) => {
      if (celda.alfa < 0.5) { indice[i] = -1; return; }
      const p = celda.lab;
      let a = 0, b = 0, dA = Infinity, dB = Infinity;
      for (let j = 0; j < paleta.length; j++) {
        const dd = dist2(p, paleta[j]);
        if (dd < dA) { b = a; dB = dA; a = j; dA = dd; } else if (dd < dB) { b = j; dB = dd; }
      }
      if (tramado && paleta.length > 1) {
        const A = paleta[a], B = paleta[b], ab = dist2(A, B);
        if (ab > 0 && ab < 0.02) {
          const t = ((p[0] - A[0]) * (B[0] - A[0]) + (p[1] - A[1]) * (B[1] - A[1]) + (p[2] - A[2]) * (B[2] - A[2])) / ab;
          const f = Math.floor(i / malla.columnas), c = i % malla.columnas;
          if (t > BAYER4[(f % 4) * 4 + (c % 4)]) a = b;
        }
      }
      indice[i] = a;
    });
    return indice;
  }

  // Una celda cuyo color no comparte con ningún vecino directo toma el color
  // más frecuente a su alrededor: saca el "confeti" que deja la reducción.
  function limpiarSueltas(indice, columnas, filas) {
    const salida = indice.slice();
    for (let f = 0; f < filas; f++) {
      for (let c = 0; c < columnas; c++) {
        const v = indice[f * columnas + c];
        if (v < 0) continue;
        const cuenta = new Map();
        let acompanada = false;
        for (let dy = -1; dy <= 1; dy++) {
          for (let dx = -1; dx <= 1; dx++) {
            if (!dx && !dy) continue;
            const y = f + dy, x = c + dx;
            if (y < 0 || x < 0 || y >= filas || x >= columnas) continue;
            const w = indice[y * columnas + x];
            if (w === v && (dx === 0 || dy === 0)) acompanada = true;
            if (w >= 0) cuenta.set(w, (cuenta.get(w) || 0) + 1);
          }
        }
        if (acompanada) continue;
        let mejor = v, max = 0;
        for (const [w, n] of cuenta) if (n > max) { max = n; mejor = w; }
        if (max >= 3) salida[f * columnas + c] = mejor;
      }
    }
    return salida;
  }

  const linea = (ctx, x0, y0, x1, y1) => { ctx.beginPath(); ctx.moveTo(x0, y0); ctx.lineTo(x1, y1); ctx.stroke(); };
  const circulo = (ctx, x, y, r) => { ctx.beginPath(); ctx.arc(x, y, r, 0, Math.PI * 2); ctx.fill(); };
  const redondeado = (ctx, x, y, w, h, r) => { ctx.beginPath(); ctx.roundRect(x, y, w, h, r); ctx.fill(); };

  // Cada estilo dibuja UNA celda de lado S (un sprite por color de la paleta).
  // sep: 'auto' (lo propio de cada estilo) o un color CSS.
  const ESTILOS = {
    liso(ctx, S, lab) {
      ctx.fillStyle = tono(lab, 1);
      ctx.fillRect(0, 0, S, S);
    },
    // Línea fina entre celdas: por defecto, un tono más oscuro de cada una
    cuadricula(ctx, S, lab, sep, grosor) {
      ctx.fillStyle = sep === 'auto' ? tono(lab, 0.6) : sep;
      ctx.fillRect(0, 0, S, S);
      const g = Math.max(1, Math.round(S * grosor));
      ctx.fillStyle = tono(lab, 1);
      ctx.fillRect(g / 2, g / 2, S - g, S - g);
    },
    // Una X de hilo sobre tela (aida). Entre puntadas la tela queda en la
    // sombra del hilo; por defecto la tela es crema.
    puntoCruz(ctx, S, lab, sep) {
      ctx.fillStyle = sep === 'auto' ? '#ece5d6' : sep;
      ctx.fillRect(0, 0, S, S);
      ctx.globalAlpha = 0.65;
      ctx.fillStyle = tono(lab, 0.45);
      ctx.fillRect(0, 0, S, S);
      ctx.globalAlpha = 1;
      const m = S * 0.12, w = S * 0.46;
      ctx.lineCap = 'round';
      const brazo = (x0, y0, x1, y1) => {
        ctx.lineWidth = w; ctx.strokeStyle = tono(lab, 0.62);
        linea(ctx, x0, y0, x1, y1);
        ctx.lineWidth = w * 0.72; ctx.strokeStyle = tono(lab, 1);
        linea(ctx, x0, y0, x1, y1);
        // Hilo retorcido: rayitas cortas cruzadas a lo largo del brazo
        ctx.lineWidth = Math.max(0.6, w * 0.12); ctx.strokeStyle = tono(lab, 0.8);
        const nx = -(y1 - y0), ny = x1 - x0, largo = Math.hypot(nx, ny);
        const ox = (nx / largo) * w * 0.3, oy = (ny / largo) * w * 0.3;
        for (let k = 1; k < 4; k++) {
          const t = k / 4, cx = x0 + (x1 - x0) * t, cy = y0 + (y1 - y0) * t;
          linea(ctx, cx - ox - (x1 - x0) * 0.06, cy - oy - (y1 - y0) * 0.06, cx + ox + (x1 - x0) * 0.06, cy + oy + (y1 - y0) * 0.06);
        }
        ctx.lineWidth = Math.max(0.6, w * 0.18); ctx.strokeStyle = tono(lab, 1, 0.12);
        linea(ctx, x0 + w * 0.08, y0 - w * 0.08, x1 + w * 0.08, y1 - w * 0.08);
      };
      brazo(m, S - m, S - m, m); // "/" abajo
      brazo(m, m, S - m, S - m); // "\" encima
    },
    // Cuenta de plástico con agujero, sobre el tablero (blanco por defecto)
    cuentas(ctx, S, lab, sep) {
      ctx.fillStyle = sep === 'auto' ? '#f1f1f1' : sep;
      ctx.fillRect(0, 0, S, S);
      ctx.fillStyle = tono(lab, 0.7);
      circulo(ctx, S / 2, S / 2, S * 0.48);
      ctx.fillStyle = tono(lab, 1);
      circulo(ctx, S / 2 - S * 0.02, S / 2 - S * 0.02, S * 0.43);
      ctx.fillStyle = tono(lab, 0.55);
      circulo(ctx, S / 2, S / 2, S * 0.15);
      ctx.strokeStyle = tono(lab, 1, 0.18); ctx.lineWidth = Math.max(1, S * 0.07); ctx.lineCap = 'round';
      ctx.beginPath(); ctx.arc(S / 2, S / 2, S * 0.3, Math.PI * 1.05, Math.PI * 1.45); ctx.stroke();
    },
    // Piedra redonda con brillo (diamond painting), sobre fondo oscuro
    diamante(ctx, S, lab, sep) {
      ctx.fillStyle = sep === 'auto' ? '#1b1b1f' : sep;
      ctx.fillRect(0, 0, S, S);
      const r = S * 0.47;
      const grad = ctx.createRadialGradient(S * 0.36, S * 0.34, S * 0.04, S / 2, S / 2, r);
      grad.addColorStop(0, tono(lab, 1, 0.3));
      grad.addColorStop(0.35, tono(lab, 1, 0.06));
      grad.addColorStop(0.8, tono(lab, 0.85));
      grad.addColorStop(1, tono(lab, 0.6));
      ctx.fillStyle = grad;
      circulo(ctx, S / 2, S / 2, r);
      ctx.fillStyle = 'rgba(255,255,255,0.85)';
      circulo(ctx, S * 0.34, S * 0.32, Math.max(0.6, S * 0.06));
    },
    // Cuenta ovalada; las filas van corridas media cuenta (ver "ladrillo")
    abalorio(ctx, S, lab, sep) {
      ctx.fillStyle = sep === 'auto' ? tono(lab, 0.45) : sep;
      ctx.fillRect(0, 0, S, S);
      const w = S * 0.92, h = S * 0.8, x = (S - w) / 2, y = (S - h) / 2;
      ctx.fillStyle = tono(lab, 0.75);
      redondeado(ctx, x, y, w, h, h * 0.45);
      ctx.fillStyle = tono(lab, 1);
      redondeado(ctx, x + S * 0.04, y + S * 0.03, w - S * 0.1, h - S * 0.12, h * 0.4);
      ctx.fillStyle = tono(lab, 1, 0.2);
      redondeado(ctx, x + w * 0.2, y + h * 0.18, w * 0.35, h * 0.16, h * 0.08);
    },
    // Pop art: un punto de color por celda (tramado Ben-Day) sobre el mismo
    // color oscurecido, como la textura "Puntos" pero con la paleta corta
    popArt(ctx, S, lab, sep) {
      ctx.fillStyle = sep === 'auto' ? tono(lab, 0.55) : sep;
      ctx.fillRect(0, 0, S, S);
      ctx.fillStyle = tono(lab, 1);
      circulo(ctx, S / 2, S / 2, S * 0.42);
    },
  };

  function mosaico(fuente, tipo, o) {
    const nombreEstilo = TIPOS[tipo].estilo;
    const ladrillo = nombreEstilo === 'abalorio';
    // Celdas de ancho, pero sin que una puntada o una cuenta baje de 6 px en
    // pantalla (en una miniatura chica ya no se distinguirían)
    const minimo = nombreEstilo === 'liso' ? 2 : 6;
    const columnas = Math.max(4, Math.min(Math.round(o.columnas), Math.floor(fuente.width / minimo)));
    const S = Math.max(minimo, Math.round(fuente.width / columnas));
    const malla = muestrear(fuente, columnas, ladrillo, 1.3);
    const puntos = [];
    for (const c of malla.celdas) if (c.alfa >= 0.5) puntos.push(c.lab);
    const paleta = calcularPaleta(puntos, o.colores);
    let indice = asignar(malla, paleta, o.tramado);
    // El tramado deja celdas sueltas a propósito: no se limpian
    if (o.limpiar && !o.tramado) indice = limpiarSueltas(indice, malla.columnas, malla.filas);

    const sep = o.separacion === 'ninguna' ? 'auto' : o.separacion;
    const sprites = paleta.map((lab) => {
      const cv = document.createElement('canvas');
      cv.width = cv.height = S;
      ESTILOS[nombreEstilo](cv.getContext('2d'), S, lab, sep, o.grosor);
      return cv;
    });
    const salida = document.createElement('canvas');
    salida.width = malla.columnas * S + (ladrillo ? Math.round(S / 2) : 0);
    salida.height = malla.filas * S;
    const ctx = salida.getContext('2d');
    for (let f = 0; f < malla.filas; f++) {
      const corrida = ladrillo && f % 2 ? Math.round(S / 2) : 0;
      for (let c = 0; c < malla.columnas; c++) {
        const v = indice[f * malla.columnas + c];
        if (v >= 0) ctx.drawImage(sprites[v], c * S + corrida, f * S);
      }
    }
    return salida;
  }

  // --- Cómic -------------------------------------------------------------

  // Kuwahara con imágenes integrales: cada píxel toma el promedio del menos
  // variable de sus 4 cuadrantes de radio r. Aplana las zonas sin cruzar los
  // bordes (un desenfoque común los borronea).
  function kuwahara(d, w, h, r) {
    const W = w + 1, N = W * (h + 1);
    const sR = new Float64Array(N), sG = new Float64Array(N), sB = new Float64Array(N), sL = new Float64Array(N), sL2 = new Float64Array(N);
    for (let y = 0; y < h; y++) {
      let aR = 0, aG = 0, aB = 0, aL = 0, aL2 = 0;
      for (let x = 0; x < w; x++) {
        const i = (y * w + x) * 4, R = d[i], G = d[i + 1], B = d[i + 2];
        const L = 0.299 * R + 0.587 * G + 0.114 * B;
        aR += R; aG += G; aB += B; aL += L; aL2 += L * L;
        const j = (y + 1) * W + x + 1, k = y * W + x + 1;
        sR[j] = sR[k] + aR; sG[j] = sG[k] + aG; sB[j] = sB[k] + aB; sL[j] = sL[k] + aL; sL2[j] = sL2[k] + aL2;
      }
    }
    const suma = (s, x0, y0, x1, y1) => s[y1 * W + x1] - s[y0 * W + x1] - s[y1 * W + x0] + s[y0 * W + x0];
    const salida = new Uint8ClampedArray(d.length);
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        let mejor = Infinity, mR = 0, mG = 0, mB = 0;
        for (let q = 0; q < 4; q++) {
          const x0 = q & 1 ? x : Math.max(0, x - r), x1 = q & 1 ? Math.min(w, x + r + 1) : x + 1;
          const y0 = q & 2 ? y : Math.max(0, y - r), y1 = q & 2 ? Math.min(h, y + r + 1) : y + 1;
          const n = (x1 - x0) * (y1 - y0);
          const mL = suma(sL, x0, y0, x1, y1) / n;
          const v = suma(sL2, x0, y0, x1, y1) / n - mL * mL;
          if (v < mejor) {
            mejor = v;
            mR = suma(sR, x0, y0, x1, y1) / n; mG = suma(sG, x0, y0, x1, y1) / n; mB = suma(sB, x0, y0, x1, y1) / n;
          }
        }
        const i = (y * w + x) * 4;
        salida[i] = mR; salida[i + 1] = mG; salida[i + 2] = mB; salida[i + 3] = d[i + 3];
      }
    }
    return salida;
  }

  function cuantizar(d, w, h, colores) {
    const labs = new Array(w * h);
    const puntos = [];
    const paso = Math.max(1, Math.floor(Math.sqrt((w * h) / 60000)));
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const i = y * w + x;
        if (d[i * 4 + 3] < 128) { labs[i] = null; continue; }
        const lab = linealAOklab(A_LINEAL[d[i * 4]], A_LINEAL[d[i * 4 + 1]], A_LINEAL[d[i * 4 + 2]]);
        lab[1] *= 1.25; lab[2] *= 1.25;
        labs[i] = lab;
        if (x % paso === 0 && y % paso === 0) puntos.push(lab);
      }
    }
    const paleta = calcularPaleta(puntos, colores);
    return { paleta, indice: Int16Array.from(labs, (lab) => (lab ? masCercano(lab, paleta) : -1)) };
  }

  // Filtro de moda: cada píxel toma el color más frecuente de su entorno.
  // Redondea los bordes dentados que deja la reducción de colores.
  function moda(indice, w, h, k, r) {
    const salida = new Int16Array(indice.length);
    const cuenta = new Int32Array(k + 1);
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        cuenta.fill(0);
        for (let dy = -r; dy <= r; dy++) {
          const fila = Math.min(h - 1, Math.max(0, y + dy)) * w;
          for (let dx = -r; dx <= r; dx++) cuenta[indice[fila + Math.min(w - 1, Math.max(0, x + dx))] + 1]++;
        }
        let mejor = 0;
        for (let c = 1; c <= k; c++) if (cuenta[c] > cuenta[mejor]) mejor = c;
        salida[y * w + x] = mejor - 1;
      }
    }
    return salida;
  }

  // Cada zona de menos de "minimo" píxeles toma el color del vecino con el
  // que más borde comparte: las manchitas, vectorizadas, serían ruido. Si
  // contrasta mucho con ese vecino (un ojo, una nariz, un brillo) se conserva
  // mientras no sea diminuta: sin esto el elefante de prueba perdía el ojo.
  function absorberManchas(indice, w, h, minimo, paleta) {
    const visto = new Uint8Array(indice.length);
    const pila = new Int32Array(indice.length);
    for (let inicio = 0; inicio < indice.length; inicio++) {
      if (visto[inicio]) continue;
      const color = indice[inicio];
      let tope = 0, tam = 0;
      pila[tope++] = inicio; visto[inicio] = 1;
      const miembros = [];
      const vecinos = new Map();
      while (tope) {
        const p = pila[--tope];
        tam++;
        if (tam <= minimo) miembros.push(p);
        const x = p % w, y = (p - x) / w;
        for (const q of [x > 0 ? p - 1 : -1, x < w - 1 ? p + 1 : -1, y > 0 ? p - w : -1, y < h - 1 ? p + w : -1]) {
          if (q < 0) continue;
          if (indice[q] === color) {
            if (!visto[q]) { visto[q] = 1; pila[tope++] = q; }
          } else if (tam <= minimo) vecinos.set(indice[q], (vecinos.get(indice[q]) || 0) + 1);
        }
      }
      if (color < 0 || tam >= minimo || !vecinos.size) continue;
      let mejor = color, max = 0;
      for (const [c, v] of vecinos) if (c >= 0 && v > max) { max = v; mejor = c; }
      if (mejor === color) continue;
      const contrasta = Math.sqrt(dist2(paleta[color], paleta[mejor])) > 0.22;
      if (!contrasta || tam < minimo / 12) for (const p of miembros) indice[p] = mejor;
    }
  }

  // Agrupa los colores parecidos: entre ellos no hay línea, como un
  // dibujante que marca el contorno de las cosas y no cada sombra. Dos grupos
  // se unen solo si TODOS sus colores están cerca (enlace completo): con
  // enlace simple los tonos intermedios encadenan el negro con el verde claro
  // y no queda ninguna línea.
  function gruposDeLinea(paleta, umbral) {
    const lista = paleta.map((_, i) => [i]);
    const lejania = (A, B) => {
      let m = 0;
      for (const a of A) for (const b of B) m = Math.max(m, Math.sqrt(dist2(paleta[a], paleta[b])));
      return m;
    };
    for (;;) {
      let par = null, menor = umbral;
      for (let i = 0; i < lista.length; i++) {
        for (let j = i + 1; j < lista.length; j++) {
          const dd = lejania(lista[i], lista[j]);
          if (dd < menor) { menor = dd; par = [i, j]; }
        }
      }
      if (!par) break;
      lista[par[0]] = lista[par[0]].concat(lista[par[1]]);
      lista.splice(par[1], 1);
    }
    const grupo = new Array(paleta.length);
    lista.forEach((miembros, g) => { for (const c of miembros) grupo[c] = g; });
    return grupo;
  }

  // Figuras vectoriales de cada color. Se le pasa a ImageTracer el mapa de
  // colores ya limpio (su formato: una matriz con un borde de -1), así no
  // repite su propia reducción de colores. Devuelve [{ color, d }] con "d" en
  // el formato de SVG, de la figura más grande a la más chica.
  function vectorizar(indice, w, h, k) {
    if (!window.ImageTracer) throw new Error('El cómic necesita vendor/imagetracerjs cargado antes que efectos-mosaico.js');
    const array = [];
    for (let y = -1; y <= h; y++) {
      const fila = new Array(w + 2).fill(-1);
      if (y >= 0 && y < h) for (let x = 0; x < w; x++) fila[x + 1] = indice[y * w + x];
      array.push(fila);
    }
    const ii = { array, palette: [] };
    const op = ImageTracer.checkoptions({ ltres: 1, qtres: 1, pathomit: 8, rightangleenhance: false, roundcoords: -1, scale: 1 });
    const capas = [];
    for (let c = 0; c < k; c++) {
      capas.push(ImageTracer.batchtracepaths(
        ImageTracer.internodes(ImageTracer.pathscan(ImageTracer.layeringstep(ii, c), op.pathomit), op), op.ltres, op.qtres));
    }
    const td = { layers: capas, palette: capas.map(() => ({ r: 0, g: 0, b: 0, a: 255 })) };
    const figuras = [];
    capas.forEach((capa, c) => capa.forEach((fig, f) => {
      if (fig.isholepath) return;
      const m = /d="([^"]*)"/.exec(ImageTracer.svgpathstring(td, c, f, op));
      const [x0, y0, x1, y1] = fig.boundingbox;
      if (m) figuras.push({ color: c, d: m[1], area: (x1 - x0) * (y1 - y0) });
    }));
    // Las chicas encima: tapan las uniones de las grandes
    return figuras.sort((a, b) => b.area - a.area);
  }

  function comic(fuente, o) {
    const ancho = fuente.width, alto = fuente.height;
    // Se procesa a "detalle" px como mucho: el resultado es vectorial, así que
    // se dibuja nítido al tamaño real igual, y el tiempo no crece con la foto.
    const trabajo = Math.min(o.detalle, Math.max(ancho, alto));
    const esc = trabajo / Math.max(ancho, alto);
    const w = Math.max(1, Math.round(ancho * esc)), h = Math.max(1, Math.round(alto * esc));
    const cv = document.createElement('canvas');
    cv.width = w; cv.height = h;
    const ctx0 = cv.getContext('2d', { willReadFrequently: true });
    ctx0.drawImage(fuente, 0, 0, w, h);
    let d = ctx0.getImageData(0, 0, w, h).data;

    // Suavizado, limpieza y grosor están pensados para la resolución de
    // trabajo completa; en una imagen más chica (una miniatura) se achican
    // en proporción, así se parece a lo que sale al descargar.
    const proporcion = trabajo / o.detalle;
    const suavizado = Math.round(o.suavizado * proporcion);
    if (suavizado > 0) {
      d = kuwahara(d, w, h, suavizado);
      d = kuwahara(d, w, h, Math.max(1, suavizado >> 1));
    }
    const { paleta, indice: crudo } = cuantizar(d, w, h, o.colores);
    const indice = o.limpieza > 0 ? moda(crudo, w, h, paleta.length, Math.max(1, Math.round(o.limpieza * proporcion))) : crudo;
    absorberManchas(indice, w, h, Math.round(w * h * 0.001), paleta);

    const salida = document.createElement('canvas');
    salida.width = ancho; salida.height = alto;
    const ctx = salida.getContext('2d');
    ctx.setTransform(ancho / w, 0, 0, alto / h, -ancho / w, -alto / h); // el borde de -1 de ImageTracer
    ctx.lineJoin = 'round'; ctx.lineCap = 'round';
    // Relleno + un trazo fino del mismo color, que tapa las rendijas
    const figuras = vectorizar(indice, w, h, paleta.length);
    for (const f of figuras) {
      const camino = new Path2D(f.d);
      ctx.fillStyle = ctx.strokeStyle = tono(paleta[f.color], 1);
      ctx.lineWidth = 1;
      ctx.fill(camino, 'evenodd');
      ctx.stroke(camino);
    }
    if (o.separacion !== 'ninguna') {
      const grupo = gruposDeLinea(paleta, 0.16);
      const nGrupos = Math.max(...grupo) + 1;
      // La línea automática es el color más oscuro de su grupo, más oscuro aún
      const oscuro = [];
      paleta.forEach((lab, c) => { const g = grupo[c]; if (!oscuro[g] || lab[0] < oscuro[g][0]) oscuro[g] = lab; });
      const porGrupo = Int16Array.from(indice, (c) => (c < 0 ? -1 : grupo[c]));
      ctx.lineWidth = Math.max(0.4, o.grosor * proporcion);
      for (const f of vectorizar(porGrupo, w, h, nGrupos)) {
        ctx.strokeStyle = o.separacion === 'auto' ? tono(oscuro[f.color], 0.4) : o.separacion;
        ctx.stroke(new Path2D(f.d));
      }
    }
    return salida;
  }

  // --- Entrada común -----------------------------------------------------

  // Reemplaza el rectángulo (sx, sy, sw, sh) de ctx por su versión en el
  // estilo "tipo". Lo que queda fuera de la figura (celdas o zonas
  // transparentes) queda transparente: respetan el recorte de un sticker.
  function aplicar(tipo, ctx, sx, sy, sw, sh, opciones) {
    if (sw < 1 || sh < 1) return;
    const o = { ...deFabrica(tipo), ...opciones };
    const fuente = document.createElement('canvas');
    fuente.width = sw; fuente.height = sh;
    fuente.getContext('2d').drawImage(ctx.canvas, sx, sy, sw, sh, 0, 0, sw, sh);
    const resultado = tipo === 'comic' ? comic(fuente, o) : mosaico(fuente, tipo, o);
    ctx.save();
    ctx.clearRect(sx, sy, sw, sh);
    ctx.imageSmoothingEnabled = tipo !== 'pixel-art';
    ctx.drawImage(resultado, sx, sy, sw, sh);
    ctx.restore();
  }

  // --- Interfaz compartida por los dos editores -------------------------

  const TEXTOS = {
    es: {
      'textura-pixel-art': 'Pixel art', 'textura-cuadricula': 'Cuadrícula', 'textura-punto-cruz': 'Punto de cruz',
      'textura-cuentas': 'Cuentas', 'textura-diamante': 'Diamantes', 'textura-abalorios': 'Abalorios',
      'textura-pop-art': 'Pop art', 'textura-comic': 'Cómic',
      'mosaico-columnas': 'Celdas de ancho', 'mosaico-colores': 'Colores', 'mosaico-tramado': 'Tramado',
      'mosaico-limpiar': 'Quitar celdas sueltas', 'mosaico-separacion': 'Separación', 'mosaico-grosor': 'Grosor de la separación',
      'comic-suavizado': 'Suavizado', 'comic-limpieza': 'Limpieza', 'comic-detalle': 'Detalle',
      'comic-linea': 'Línea', 'comic-grosor': 'Grosor de línea',
      'mosaico-auto': 'Automática', 'mosaico-negra': 'Negra', 'mosaico-blanca': 'Blanca', 'mosaico-sin-linea': 'Sin línea',
      'mosaico-elegir-color': 'Elegir color',
    },
    en: {
      'textura-pixel-art': 'Pixel art', 'textura-cuadricula': 'Grid', 'textura-punto-cruz': 'Cross-stitch',
      'textura-cuentas': 'Beads', 'textura-diamante': 'Diamonds', 'textura-abalorios': 'Seed beads',
      'textura-pop-art': 'Pop art', 'textura-comic': 'Comic',
      'mosaico-columnas': 'Cells across', 'mosaico-colores': 'Colors', 'mosaico-tramado': 'Dithering',
      'mosaico-limpiar': 'Remove stray cells', 'mosaico-separacion': 'Gaps', 'mosaico-grosor': 'Gap thickness',
      'comic-suavizado': 'Smoothing', 'comic-limpieza': 'Cleanup', 'comic-detalle': 'Detail',
      'comic-linea': 'Line', 'comic-grosor': 'Line thickness',
      'mosaico-auto': 'Automatic', 'mosaico-negra': 'Black', 'mosaico-blanca': 'White', 'mosaico-sin-linea': 'No line',
      'mosaico-elegir-color': 'Pick a color',
    },
  };
  // Se suman a los textos de la página (los de la página mandan) antes de
  // que idioma.js aplique el idioma, que espera a DOMContentLoaded.
  (function () {
    const pagina = window.i18nPagina || (window.i18nPagina = {});
    for (const idioma of Object.keys(TEXTOS)) pagina[idioma] = Object.assign({}, TEXTOS[idioma], pagina[idioma]);
  })();
  const texto = (clave) => (window.t ? window.t(clave) : TEXTOS.es[clave]);

  // Panel con los controles de estos estilos, al final de "contenedor" (los
  // controles de textura de la página). Solo se ve con un tipo de este
  // módulo, y entonces oculta "campoIntensidad" (el deslizador de intensidad
  // de las demás texturas), que estos estilos no usan. alCambiar() se llama
  // cuando la persona toca algo; leer() devuelve las opciones del tipo
  // visible y poner(opciones, tipo) las muestra (null = las de fábrica).
  function crearPanel(contenedor, { alCambiar, campoIntensidad }) {
    const panel = document.createElement('div');
    panel.className = 'opciones-mosaico';
    contenedor.append(panel);

    // Controles con las piezas de los editores (.campo-deslizador,
    // .campo-check, .modos). data-grupo: 'mosaico', 'comic' o 'ambos'.
    const controles = {};
    function deslizador(clave, etiqueta, min, max, paso, grupo, formato = String) {
      const campo = document.createElement('label');
      campo.className = 'campo-deslizador';
      campo.dataset.grupo = grupo;
      campo.innerHTML = `<span class="etiqueta-deslizador"><span data-i18n="${etiqueta}"></span><span></span></span><input type="range" min="${min}" max="${max}" step="${paso}">`;
      panel.append(campo);
      const entrada = campo.querySelector('input');
      const valor = campo.querySelector('.etiqueta-deslizador span:last-child');
      const rotularValor = () => { valor.textContent = formato(Number(entrada.value)); };
      controles[clave] = { leer: () => Number(entrada.value), poner: (v) => { entrada.value = v; rotularValor(); } };
      entrada.addEventListener('input', () => { rotularValor(); alCambiar(); });
      return campo;
    }
    function casilla(clave, etiqueta, grupo) {
      const campo = document.createElement('label');
      campo.className = 'campo-check';
      campo.dataset.grupo = grupo;
      campo.innerHTML = `<input type="checkbox"> <span data-i18n="${etiqueta}"></span>`;
      panel.append(campo);
      const entrada = campo.querySelector('input');
      controles[clave] = { leer: () => entrada.checked, poner: (v) => { entrada.checked = !!v; } };
      entrada.addEventListener('change', alCambiar);
    }

    deslizador('columnas', 'mosaico-columnas', 8, 200, 1, 'mosaico');
    deslizador('colores', 'mosaico-colores', 2, 32, 1, 'ambos');
    casilla('tramado', 'mosaico-tramado', 'mosaico');
    casilla('limpiar', 'mosaico-limpiar', 'mosaico');
    deslizador('suavizado', 'comic-suavizado', 0, 10, 1, 'comic');
    deslizador('limpieza', 'comic-limpieza', 0, 4, 1, 'comic');
    deslizador('detalle', 'comic-detalle', 320, 1024, 32, 'comic', (v) => v + ' px');

    // Separación (mosaicos) o línea (cómic): automática, negra, blanca, sin
    // línea (solo cómic) o un color elegido
    const grupoSeparacion = document.createElement('div');
    grupoSeparacion.dataset.grupo = 'ambos';
    grupoSeparacion.innerHTML = '<span class="etiqueta-modos"></span><div class="modos">' +
      '<button type="button" data-separacion="auto" data-i18n="mosaico-auto"></button>' +
      '<button type="button" data-separacion="#000000" data-i18n="mosaico-negra"></button>' +
      '<button type="button" data-separacion="#ffffff" data-i18n="mosaico-blanca"></button>' +
      '<button type="button" data-separacion="ninguna" data-i18n="mosaico-sin-linea"></button>' +
      '<input type="color" value="#5a3e8c" data-i18n-attr="aria-label:mosaico-elegir-color;title:mosaico-elegir-color">' +
      '</div>';
    panel.append(grupoSeparacion);
    const campoGrosorSeparacion = deslizador('grosorSeparacion', 'mosaico-grosor', 4, 30, 1, 'mosaico', (v) => v + '%');
    deslizador('grosorLinea', 'comic-grosor', 4, 50, 1, 'comic', (v) => (v / 10).toFixed(1));

    panel.querySelectorAll('[data-i18n]').forEach((el) => { el.textContent = texto(el.dataset.i18n); });
    const etiquetaSeparacion = grupoSeparacion.querySelector('.etiqueta-modos');
    const botonSinLinea = grupoSeparacion.querySelector('[data-separacion="ninguna"]');
    const entradaColor = grupoSeparacion.querySelector('input[type="color"]');
    entradaColor.title = texto('mosaico-elegir-color');
    entradaColor.setAttribute('aria-label', entradaColor.title);
    let tipoActual = null;
    let separacion = 'auto';

    function marcarSeparacion() {
      const elegido = !['auto', '#000000', '#ffffff', 'ninguna'].includes(separacion);
      grupoSeparacion.querySelectorAll('[data-separacion]').forEach((b) => {
        const activo = b.dataset.separacion === separacion;
        b.classList.toggle('activo', activo);
        b.setAttribute('aria-pressed', String(activo));
      });
      entradaColor.classList.toggle('activo', elegido);
      if (elegido) entradaColor.value = separacion;
    }
    function rotular() {
      etiquetaSeparacion.textContent = texto(tipoActual === 'comic' ? 'comic-linea' : 'mosaico-separacion');
    }
    grupoSeparacion.querySelectorAll('[data-separacion]').forEach((b) => b.addEventListener('click', () => {
      separacion = b.dataset.separacion;
      marcarSeparacion(); alCambiar();
    }));
    entradaColor.addEventListener('input', () => { separacion = entradaColor.value; marcarSeparacion(); alCambiar(); });
    document.addEventListener('cambioIdioma', rotular);

    function mostrar(tipo) {
      tipoActual = tipo;
      const esPropio = !!TIPOS[tipo];
      panel.hidden = !esPropio;
      if (campoIntensidad) campoIntensidad.hidden = esPropio;
      if (!esPropio) return;
      const grupo = tipo === 'comic' ? 'comic' : 'mosaico';
      panel.querySelectorAll('[data-grupo]').forEach((el) => {
        el.hidden = el.dataset.grupo !== 'ambos' && el.dataset.grupo !== grupo;
      });
      // El pixel art no tiene separación; el grosor solo lo usa la cuadrícula
      if (tipo === 'pixel-art') grupoSeparacion.hidden = true;
      if (tipo !== 'cuadricula') campoGrosorSeparacion.hidden = true;
      botonSinLinea.hidden = grupo !== 'comic';
      if (separacion === 'ninguna' && grupo !== 'comic') separacion = 'auto';
      rotular(); marcarSeparacion();
    }
    // Las opciones del otro grupo (mosaico / cómic) quedan en sus valores de
    // fábrica: cada capa o cada foto guarda solo las del tipo que usa.
    function poner(opciones, tipo) {
      const esComic = tipo === 'comic';
      const o = { ...(TIPOS[tipo] ? deFabrica(tipo) : deFabrica('cuadricula')), ...opciones };
      const mosaico = esComic ? OPCIONES.mosaico : o;
      const comicOp = esComic ? o : OPCIONES.comic;
      controles.colores.poner(o.colores);
      controles.columnas.poner(mosaico.columnas);
      controles.tramado.poner(mosaico.tramado);
      controles.limpiar.poner(mosaico.limpiar);
      controles.grosorSeparacion.poner(Math.round(mosaico.grosor * 100));
      controles.suavizado.poner(comicOp.suavizado);
      controles.limpieza.poner(comicOp.limpieza);
      controles.detalle.poner(comicOp.detalle);
      controles.grosorLinea.poner(Math.round(comicOp.grosor * 10));
      separacion = o.separacion;
      mostrar(tipo);
    }
    function leer() {
      const comun = { colores: controles.colores.leer(), separacion };
      if (tipoActual === 'comic') {
        return { ...comun, suavizado: controles.suavizado.leer(), limpieza: controles.limpieza.leer(), detalle: controles.detalle.leer(), grosor: controles.grosorLinea.leer() / 10 };
      }
      return { ...comun, columnas: controles.columnas.leer(), tramado: controles.tramado.leer(), limpiar: controles.limpiar.leer(), grosor: controles.grosorSeparacion.leer() / 100 };
    }
    poner(null, null);
    return { mostrar, poner, leer };
  }

  return {
    TIPOS: Object.keys(TIPOS),
    aplicar,
    crearPanel,
  };
})();

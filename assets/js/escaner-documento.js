/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Motor del escáner de documentos: de una foto de una hoja a una página plana
// y legible, sin librerías de terceros.
//
//   detectarEsquinas(imagen)           -> { esquinas, confianza, encontrado }
//   tamanoSalida(esquinas, maxLado)    -> { ancho, alto }
//   enderezar(imagen, esquinas, w, h)  -> imagen nueva (corrige la perspectiva)
//   girar(imagen, cuartos)             -> imagen nueva girada en cuartos de vuelta
//   mejorar(imagen, modo, opciones)    -> la misma imagen, ya mejorada
//
// `imagen` es cualquier objeto { data, width, height } con `data` RGBA de 8
// bits (un ImageData sirve): así el motor no depende del DOM y se prueba en
// Node con imágenes sintéticas. Las esquinas van siempre en el orden
// arriba-izquierda, arriba-derecha, abajo-derecha, abajo-izquierda.
//
// Cómo busca la hoja: baja la foto a ~320 px, y para tres separaciones posibles
// (más clara que el fondo, más oscura, distinta del color del borde) toma la
// mancha conexa más grande, le saca la envolvente convexa y busca el
// cuadrilátero de mayor área entre sus vértices. Gana la mancha que mejor
// "parece una hoja" (llena su cuadrilátero, no toca los bordes de la foto). Si
// ninguna convence, devuelve el cuadro entero con `encontrado: false` y la
// persona ajusta las esquinas a mano.

const EscanerDocumento = (() => {
  const LADO_ANALISIS = 320;

  const limitar = (v, a, b) => (v < a ? a : v > b ? b : v);

  // ---------- Reducción y color ----------

  /** Reduce promediando bloques; devuelve { data, width, height, factor }. */
  function reducir(imagen, ladoMax) {
    const { data, width, height } = imagen;
    const factor = Math.max(1, Math.max(width, height) / ladoMax);
    const w = Math.max(1, Math.round(width / factor));
    const h = Math.max(1, Math.round(height / factor));
    const salida = new Uint8ClampedArray(w * h * 4);
    for (let y = 0; y < h; y++) {
      const y0 = Math.floor((y * height) / h);
      const y1 = Math.max(y0 + 1, Math.floor(((y + 1) * height) / h));
      for (let x = 0; x < w; x++) {
        const x0 = Math.floor((x * width) / w);
        const x1 = Math.max(x0 + 1, Math.floor(((x + 1) * width) / w));
        let r = 0, g = 0, b = 0, n = 0;
        for (let yy = y0; yy < y1; yy++) {
          for (let xx = x0; xx < x1; xx++) {
            const i = (yy * width + xx) * 4;
            r += data[i]; g += data[i + 1]; b += data[i + 2]; n++;
          }
        }
        const o = (y * w + x) * 4;
        salida[o] = r / n; salida[o + 1] = g / n; salida[o + 2] = b / n; salida[o + 3] = 255;
      }
    }
    return { data: salida, width: w, height: h, factor: width / w };
  }

  function aGris(imagen) {
    const { data, width, height } = imagen;
    const gris = new Uint8Array(width * height);
    for (let i = 0, p = 0; p < gris.length; i += 4, p++) {
      gris[p] = (data[i] * 299 + data[i + 1] * 587 + data[i + 2] * 114) / 1000;
    }
    return gris;
  }

  /** Desenfoque de caja de radio 1, separable, sobre un canal de 8 bits. */
  function desenfocar(canal, w, h, pasadas = 1) {
    let a = Float32Array.from(canal);
    let b = new Float32Array(a.length);
    for (let p = 0; p < pasadas; p++) {
      for (let y = 0; y < h; y++) {
        for (let x = 0; x < w; x++) {
          const i = y * w + x;
          b[i] = (a[x > 0 ? i - 1 : i] + a[i] + a[x < w - 1 ? i + 1 : i]) / 3;
        }
      }
      for (let y = 0; y < h; y++) {
        for (let x = 0; x < w; x++) {
          const i = y * w + x;
          a[i] = (b[y > 0 ? i - w : i] + b[i] + b[y < h - 1 ? i + w : i]) / 3;
        }
      }
    }
    return a;
  }

  /** Umbral de Otsu de un canal de 8 bits (0..255). */
  function otsu(canal) {
    const hist = new Float64Array(256);
    for (let i = 0; i < canal.length; i++) hist[limitar(Math.round(canal[i]), 0, 255)]++;
    const total = canal.length;
    let suma = 0;
    for (let t = 0; t < 256; t++) suma += t * hist[t];
    let sumaB = 0, pesoB = 0, mejor = 0, umbral = 128;
    for (let t = 0; t < 256; t++) {
      pesoB += hist[t];
      if (!pesoB) continue;
      const pesoF = total - pesoB;
      if (!pesoF) break;
      sumaB += t * hist[t];
      const mB = sumaB / pesoB, mF = (suma - sumaB) / pesoF;
      const varianza = pesoB * pesoF * (mB - mF) * (mB - mF);
      if (varianza > mejor) { mejor = varianza; umbral = t; }
    }
    return umbral;
  }

  // ---------- Máscara, mancha y envolvente ----------

  /** Dilata (true) o erosiona (false) una máscara con un cuadrado de 3x3. */
  function morfologia(mascara, w, h, dilatar) {
    const sal = new Uint8Array(mascara.length);
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        let hay = false, todo = true;
        for (let dy = -1; dy <= 1; dy++) {
          for (let dx = -1; dx <= 1; dx++) {
            const xx = x + dx, yy = y + dy;
            // Fuera de la imagen cuenta como vacío: una hoja que toca el borde
            // no se "come" por el borde al erosionar.
            const v = xx >= 0 && yy >= 0 && xx < w && yy < h ? mascara[yy * w + xx] : (dilatar ? 0 : 1);
            if (v) hay = true; else todo = false;
          }
        }
        sal[y * w + x] = dilatar ? (hay ? 1 : 0) : (todo ? 1 : 0);
      }
    }
    return sal;
  }

  /**
   * Rellena los huecos cerrados de la máscara: lo que no se alcanza desde el
   * borde de la foto pasando por vacío está rodeado y es parte de la hoja. Sin
   * esto, un dibujo con zonas oscuras y claras deja la hoja hecha un colador
   * y no parece un rectángulo.
   */
  function rellenarHuecos(mascara, w, h) {
    const fuera = new Uint8Array(w * h);
    const pila = new Int32Array(w * h);
    let tope = 0;
    const empujar = (i) => { if (!mascara[i] && !fuera[i]) { fuera[i] = 1; pila[tope++] = i; } };
    for (let x = 0; x < w; x++) { empujar(x); empujar((h - 1) * w + x); }
    for (let y = 0; y < h; y++) { empujar(y * w); empujar(y * w + w - 1); }
    while (tope) {
      const i = pila[--tope];
      const x = i % w, y = (i / w) | 0;
      if (x > 0) empujar(i - 1);
      if (x < w - 1) empujar(i + 1);
      if (y > 0) empujar(i - w);
      if (y < h - 1) empujar(i + w);
    }
    const sal = new Uint8Array(w * h);
    for (let i = 0; i < sal.length; i++) sal[i] = mascara[i] || !fuera[i] ? 1 : 0;
    return sal;
  }

  /** Mancha conexa (4 vecinos) más grande: devuelve sus píxeles y sus datos. */
  function manchaMasGrande(mascara, w, h) {
    const etiqueta = new Int32Array(w * h);
    const pila = new Int32Array(w * h);
    let mejor = null, siguiente = 0;
    for (let inicio = 0; inicio < mascara.length; inicio++) {
      if (!mascara[inicio] || etiqueta[inicio]) continue;
      siguiente++;
      let tope = 0, area = 0;
      pila[tope++] = inicio;
      etiqueta[inicio] = siguiente;
      while (tope) {
        const i = pila[--tope];
        area++;
        const x = i % w, y = (i / w) | 0;
        if (x > 0 && mascara[i - 1] && !etiqueta[i - 1]) { etiqueta[i - 1] = siguiente; pila[tope++] = i - 1; }
        if (x < w - 1 && mascara[i + 1] && !etiqueta[i + 1]) { etiqueta[i + 1] = siguiente; pila[tope++] = i + 1; }
        if (y > 0 && mascara[i - w] && !etiqueta[i - w]) { etiqueta[i - w] = siguiente; pila[tope++] = i - w; }
        if (y < h - 1 && mascara[i + w] && !etiqueta[i + w]) { etiqueta[i + w] = siguiente; pila[tope++] = i + w; }
      }
      if (!mejor || area > mejor.area) mejor = { id: siguiente, area };
    }
    if (!mejor) return null;
    const puntos = [];
    let tocaBorde = 0, bordeTotal = 0;
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        if (etiqueta[y * w + x] !== mejor.id) continue;
        // Solo los píxeles del contorno hacen falta para la envolvente.
        const dentro = x > 0 && y > 0 && x < w - 1 && y < h - 1
          && etiqueta[y * w + x - 1] === mejor.id && etiqueta[y * w + x + 1] === mejor.id
          && etiqueta[(y - 1) * w + x] === mejor.id && etiqueta[(y + 1) * w + x] === mejor.id;
        if (!dentro) puntos.push([x, y]);
        if (x === 0 || y === 0 || x === w - 1 || y === h - 1) tocaBorde++;
      }
    }
    bordeTotal = 2 * (w + h);
    return { puntos, area: mejor.area, fraccionBorde: tocaBorde / bordeTotal };
  }

  const cruz = (o, a, b) => (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0]);

  /** Envolvente convexa (cadena monótona de Andrew), en sentido horario en pantalla. */
  function envolvente(puntos) {
    const p = puntos.slice().sort((a, b) => a[0] - b[0] || a[1] - b[1]);
    if (p.length < 3) return p;
    const inf = [];
    for (const q of p) {
      while (inf.length >= 2 && cruz(inf[inf.length - 2], inf[inf.length - 1], q) <= 0) inf.pop();
      inf.push(q);
    }
    const sup = [];
    for (let i = p.length - 1; i >= 0; i--) {
      const q = p[i];
      while (sup.length >= 2 && cruz(sup[sup.length - 2], sup[sup.length - 1], q) <= 0) sup.pop();
      sup.push(q);
    }
    inf.pop(); sup.pop();
    return inf.concat(sup);
  }

  function areaPoligono(pts) {
    let a = 0;
    for (let i = 0; i < pts.length; i++) {
      const p = pts[i], q = pts[(i + 1) % pts.length];
      a += p[0] * q[1] - q[0] * p[1];
    }
    return Math.abs(a) / 2;
  }

  /** Deja a lo sumo `max` vértices, quitando siempre el que menos área aporta. */
  function simplificar(pts, max) {
    const v = pts.slice();
    while (v.length > max) {
      let peor = 0, menor = Infinity;
      for (let i = 0; i < v.length; i++) {
        const a = v[(i + v.length - 1) % v.length], b = v[i], c = v[(i + 1) % v.length];
        const ar = Math.abs(cruz(a, b, c)) / 2;
        if (ar < menor) { menor = ar; peor = i; }
      }
      v.splice(peor, 1);
    }
    return v;
  }

  /** El cuadrilátero de mayor área con vértices en la envolvente. */
  function mejorCuadrilatero(hull) {
    const v = simplificar(hull, 28);
    if (v.length <= 4) return v.length === 4 ? v : null;
    let mejor = null, areaMejor = -1;
    const n = v.length;
    for (let a = 0; a < n - 3; a++) {
      for (let b = a + 1; b < n - 2; b++) {
        for (let c = b + 1; c < n - 1; c++) {
          for (let d = c + 1; d < n; d++) {
            const ar = areaPoligono([v[a], v[b], v[c], v[d]]);
            if (ar > areaMejor) { areaMejor = ar; mejor = [v[a], v[b], v[c], v[d]]; }
          }
        }
      }
    }
    return mejor;
  }

  /** Ordena cuatro puntos como arriba-izq., arriba-der., abajo-der., abajo-izq. */
  function ordenarEsquinas(pts) {
    const cx = pts.reduce((s, p) => s + p[0], 0) / 4;
    const cy = pts.reduce((s, p) => s + p[1], 0) / 4;
    const porAngulo = pts.slice().sort((p, q) => Math.atan2(p[1] - cy, p[0] - cx) - Math.atan2(q[1] - cy, q[0] - cx));
    // Por ángulo ascendente en pantalla (y hacia abajo) el giro es horario.
    let inicio = 0, menor = Infinity;
    porAngulo.forEach((p, i) => { if (p[0] + p[1] < menor) { menor = p[0] + p[1]; inicio = i; } });
    return [0, 1, 2, 3].map((k) => porAngulo[(inicio + k) % 4]);
  }

  /** Una hoja fotografiada no tiene ángulos de 40° ni de 150°: eso es un cuadrilátero torcido por una protuberancia. */
  function angulosRazonables(q) {
    let signo = 0;
    for (let k = 0; k < 4; k++) {
      const c = cruz(q[k], q[(k + 1) % 4], q[(k + 2) % 4]);
      if (!c || (signo && Math.sign(c) !== signo)) return false; // no convexo o cruzado
      signo = Math.sign(c);
    }
    for (let k = 0; k < 4; k++) {
      const p = q[k], a = q[(k + 3) % 4], c = q[(k + 1) % 4];
      const v1x = a[0] - p[0], v1y = a[1] - p[1], v2x = c[0] - p[0], v2y = c[1] - p[1];
      const coseno = (v1x * v2x + v1y * v2y) / (Math.hypot(v1x, v1y) * Math.hypot(v2x, v2y) || 1);
      const grados = (Math.acos(Math.max(-1, Math.min(1, coseno))) * 180) / Math.PI;
      if (grados < 55 || grados > 125) return false;
    }
    // Los lados opuestos de una hoja fotografiada de frente son casi paralelos; uno inclinado 25° y su
    // opuesto recto es una mancha cortada en diagonal, no una hoja.
    const rumbo = (a, b) => Math.atan2(b[1] - a[1], b[0] - a[0]);
    const difer = (r1, r2) => { let d = Math.abs(r1 - r2) % Math.PI; if (d > Math.PI / 2) d = Math.PI - d; return (d * 180) / Math.PI; };
    if (difer(rumbo(q[0], q[1]), rumbo(q[3], q[2])) > 18 || difer(rumbo(q[0], q[3]), rumbo(q[1], q[2])) > 18) return false;
    return true;
  }

  // ---------- Mapa de bordes ----------

  /**
   * Bordes finos en color: gradiente de Sobel en el canal (R, G o B) donde el cambio es mayor
   * —un papel amarillo y una mesa gris pueden tener la misma luminosidad—, adelgazado a un píxel
   * como en Canny. `prob` es la fuerza del borde de 0 a 1 (1 = de los más marcados de la foto).
   */
  function mapaDeBordes(pequena, w, h) {
    const n = w * h;
    const canales = [0, 1, 2].map((c) => {
      const ch = new Float32Array(n);
      for (let p = 0; p < n; p++) ch[p] = pequena.data[p * 4 + c];
      return desenfocar(ch, w, h, 2);
    });
    const gx = new Float32Array(n), gy = new Float32Array(n), mag = new Float32Array(n);
    for (let y = 1; y < h - 1; y++) {
      for (let x = 1; x < w - 1; x++) {
        const i = y * w + x;
        let mejor = -1;
        for (const g of canales) {
          const ax = (g[i - w + 1] + 2 * g[i + 1] + g[i + w + 1]) - (g[i - w - 1] + 2 * g[i - 1] + g[i + w - 1]);
          const ay = (g[i + w - 1] + 2 * g[i + w] + g[i + w + 1]) - (g[i - w - 1] + 2 * g[i - w] + g[i - w + 1]);
          const m = ax * ax + ay * ay;
          if (m > mejor) { mejor = m; gx[i] = ax; gy[i] = ay; }
        }
        mag[i] = Math.sqrt(mejor);
      }
    }
    // Supresión de no máximos: queda solo la cresta de cada borde, de un píxel de ancho.
    const prob = new Float32Array(n);
    const valores = [];
    for (let y = 2; y < h - 2; y++) {
      for (let x = 2; x < w - 2; x++) {
        const i = y * w + x, m = mag[i];
        if (m < 6) continue;
        const ax = Math.abs(gx[i]), ay = Math.abs(gy[i]);
        const d = ay <= ax * 0.4142 ? 1 : ax <= ay * 0.4142 ? w : (gx[i] > 0) === (gy[i] > 0) ? w + 1 : w - 1;
        if (m >= mag[i - d] && m >= mag[i + d]) { prob[i] = m; valores.push(m); }
      }
    }
    valores.sort((a, b) => a - b);
    const tope = Math.max(16, valores.length ? valores[Math.floor(valores.length * 0.7)] : 16);
    for (let i = 0; i < n; i++) if (prob[i]) prob[i] = Math.min(1, prob[i] / tope);
    return { gx, gy, prob, canales, w, h };
  }

  // ---------- Rectas (Hough) ----------

  const COS = new Float32Array(180), SEN = new Float32Array(180);
  for (let t = 0; t < 180; t++) { COS[t] = Math.cos((t * Math.PI) / 180); SEN[t] = Math.sin((t * Math.PI) / 180); }

  /** Recta x·cos t + y·sen t = rho. `virtual` marca un borde de la foto ('arriba', 'abajo', 'izq', 'der'). */
  const nuevaRecta = (t, rho, votos, virtual) => ({ t, rho, c: COS[t], s: SEN[t], votos, virtual });

  /** Las rectas más votadas de los bordes finos; cada píxel vota solo cerca de la dirección de su borde. */
  function rectasDeHough(mapa) {
    const { gx, gy, prob, w, h } = mapa;
    const diag = Math.ceil(Math.hypot(w, h)), NR = 2 * diag + 1;
    const acc = new Float32Array(180 * NR);
    for (let y = 2; y < h - 2; y++) {
      for (let x = 2; x < w - 2; x++) {
        const i = y * w + x, p = prob[i];
        if (p < 0.2) continue;
        const t0 = Math.round((Math.atan2(gy[i], gx[i]) * 180) / Math.PI);
        for (let d = -3; d <= 3; d++) {
          const t = (((t0 + d) % 180) + 180) % 180;
          acc[t * NR + Math.round(x * COS[t] + y * SEN[t]) + diag] += p;
        }
      }
    }
    const minimo = 0.12 * Math.min(w, h);
    const indices = [];
    for (let k = 0; k < acc.length; k++) if (acc[k] >= minimo) indices.push(k);
    indices.sort((a, b) => acc[b] - acc[a]);
    const lineas = [];
    for (const k of indices) {
      if (lineas.length >= 40) break;
      const t = Math.floor(k / NR), rho = (k % NR) - diag;
      const repetida = lineas.some((l) => {
        let dt = Math.abs(l.t - t), dr = Math.abs(l.rho - rho);
        if (dt > 90) { dt = 180 - dt; dr = Math.abs(l.rho + rho); } // t y t+180 son la misma recta con rho invertido
        return dt <= 3 && dr <= 4;
      });
      if (!repetida) lineas.push(nuevaRecta(t, rho, acc[k]));
    }
    return lineas;
  }

  // ---------- Respaldo de un lado ----------

  /**
   * Cuánto escalón de color hay en (px, py) a través de la normal: el promedio de 3 a 6 px de un
   * lado contra el del otro, en el canal donde más cambia (0..1, 1 = 12 niveles o más). El borde
   * de un papel es un escalón (papel de un lado, fondo del otro); un renglón, una línea impresa o
   * un trazo de texto es una línea fina con lo mismo a los dos lados, y da casi 0.
   */
  function escalon(mapa, px, py, nx, ny) {
    const { canales, w, h } = mapa;
    let mayor = 0;
    for (const g of canales) {
      let afuera = 0, adentro = 0, na = 0, nd = 0;
      for (let k = 3; k <= 6; k++) {
        const xa = Math.round(px + nx * k), ya = Math.round(py + ny * k);
        const xd = Math.round(px - nx * k), yd = Math.round(py - ny * k);
        if (xa >= 0 && ya >= 0 && xa < w && ya < h) { afuera += g[ya * w + xa]; na++; }
        if (xd >= 0 && yd >= 0 && xd < w && yd < h) { adentro += g[yd * w + xd]; nd++; }
      }
      if (na && nd) mayor = Math.max(mayor, Math.abs(afuera / na - adentro / nd));
    }
    return Math.min(1, mayor / 12);
  }

  /** Mejor borde alineado con la normal (nx, ny) a ±1,5 px de (px, py), separado por sentido y pesado por el escalón. */
  function bordeCerca(mapa, px, py, nx, ny) {
    const { gx, gy, prob, w, h } = mapa;
    let mp = 0, mn = 0;
    for (let o = -1.5; o <= 1.5; o += 0.75) {
      const x = Math.round(px + nx * o), y = Math.round(py + ny * o);
      if (x < 1 || y < 1 || x >= w - 1 || y >= h - 1) continue;
      const i = y * w + x, p = prob[i];
      if (!p) continue;
      const coseno = (gx[i] * nx + gy[i] * ny) / Math.hypot(gx[i], gy[i]);
      if (coseno >= 0.9) { if (p > mp) mp = p; } else if (coseno <= -0.9 && p > mn) mn = p;
    }
    if (!mp && !mn) return [0, 0];
    const e = escalon(mapa, px, py, nx, ny);
    return [mp * e, mn * e];
  }

  /**
   * Respaldo (0..1) del segmento a-b: qué parte tiene un borde fino detrás, alineado con él y
   * siempre con el mismo sentido (claro→oscuro, o al revés). El sentido constante es lo que separa
   * el borde de un papel de los trazos del texto o del dibujo, que van y vuelven.
   */
  function respaldoSegmento(mapa, a, b) {
    const dx = b[0] - a[0], dy = b[1] - a[1], largo = Math.hypot(dx, dy);
    if (largo < 3) return 0;
    const nx = dy / largo, ny = -dx / largo;
    const pasos = Math.max(8, Math.round(largo));
    let pos = 0, neg = 0;
    for (let k = 0; k < pasos; k++) {
      const f = 0.03 + 0.94 * (k / (pasos - 1));
      const [mp, mn] = bordeCerca(mapa, a[0] + dx * f, a[1] + dy * f, nx, ny);
      pos += mp; neg += mn;
    }
    return Math.max(pos, neg) / pasos;
  }

  /** Lo mismo que respaldoSegmento, pero precalculado a lo largo de una recta (sumas acumuladas). */
  function perfilDeRecta(mapa, l) {
    const diag = Math.ceil(Math.hypot(mapa.w, mapa.h)), n = 2 * diag + 1;
    const pos = new Float32Array(n + 1), neg = new Float32Array(n + 1);
    const x0 = l.rho * l.c, y0 = l.rho * l.s;
    for (let k = 0; k < n; k++) {
      const t = k - diag;
      const [mp, mn] = bordeCerca(mapa, x0 - l.s * t, y0 + l.c * t, l.c, l.s);
      pos[k + 1] = pos[k] + mp; neg[k + 1] = neg[k] + mn;
    }
    l.perfil = { pos, neg, diag, n };
  }
  const paramEn = (l, p) => -p[0] * l.s + p[1] * l.c;
  function respaldoEnRecta(l, pa, pb) {
    let ta = paramEn(l, pa), tb = paramEn(l, pb);
    if (ta > tb) [ta, tb] = [tb, ta];
    const recorte = (tb - ta) * 0.03, { pos, neg, diag, n } = l.perfil;
    const a = limitar(Math.round(ta + recorte) + diag, 0, n), b = limitar(Math.round(tb - recorte) + diag, 0, n);
    return b > a ? Math.max(pos[b] - pos[a], neg[b] - neg[a]) / (b - a) : 0;
  }

  // ---------- Puntaje de un cuadrilátero ----------

  /** ¿De qué borde de la foto es el lado a-b, si está pegado a uno? */
  function ladoDeMarco(a, b, w, h) {
    const e = 1.2;
    if (a[1] <= e && b[1] <= e) return 'arriba';
    if (a[1] >= h - 1 - e && b[1] >= h - 1 - e) return 'abajo';
    if (a[0] <= e && b[0] <= e) return 'izq';
    if (a[0] >= w - 1 - e && b[0] >= w - 1 - e) return 'der';
    return null;
  }

  const alcanceDeVecinos = (alcances) => alcances.reduce((x, y) => x + y, 0) / alcances.length;

  /**
   * Un solo puntaje para todos los candidatos, vengan de rectas o de manchas: respaldo de los
   * cuatro lados (el promedio y el peor), un poco de premio al tamaño y un poco de castigo por
   * cada lado que va sobre el borde de la foto. Ese lado (la hoja sigue fuera del encuadre) vale
   * solo si los lados vecinos llegan de verdad hasta el borde: si no, cualquier rectángulo del
   * tamaño de la foto ganaría.
   * `respaldo(i, desde, hasta)` da el respaldo del lado i entre dos fracciones de su largo.
   */
  function puntuar(q, w, h, respaldo) {
    if (!angulosRazonables(q)) return -1;
    const area = areaPoligono(q) / (w * h);
    if (area < 0.06) return -1;
    const marco = [0, 1, 2, 3].map((i) => ladoDeMarco(q[i], q[(i + 1) % 4], w, h));
    const virtuales = marco.filter(Boolean).length;
    if (virtuales === 4) return -1;
    const lados = [0, 1, 2, 3].map((i) => (marco[i] ? null : respaldo(i, 0, 1)));
    for (let i = 0; i < 4; i++) {
      if (!marco[i]) continue;
      // Vecinos: el lado anterior termina en la esquina i; el siguiente empieza en la esquina i+1.
      const anterior = (i + 3) % 4, siguiente = (i + 1) % 4;
      const alcances = [];
      if (!marco[anterior]) alcances.push(respaldo(anterior, 0.8, 1));
      if (!marco[siguiente]) alcances.push(respaldo(siguiente, 0, 0.2));
      // Basta con que lleguen en promedio: el borde de una hoja contra una pared de color parecido
      // puede borrarse justo en una esquina.
      lados[i] = alcances.length && alcanceDeVecinos(alcances) >= 0.4 ? 0.8 : 0;
    }
    const media = (lados[0] + lados[1] + lados[2] + lados[3]) / 4, minimo = Math.min(...lados);
    return (0.6 * media + 0.4 * minimo) * (0.75 + 0.25 * Math.sqrt(area)) - 0.03 * virtuales;
  }

  /** Todos los cuadriláteros que se arman con dos rectas casi horizontales y dos casi verticales. */
  function candidatosDeRectas(mapa, lineas) {
    const { w, h } = mapa;
    const pos = (l) => (Math.abs(l.s) > 0.7 ? (l.rho - (w / 2) * l.c) / l.s : (l.rho - (h / 2) * l.s) / l.c);
    const horiz = lineas.filter((l) => Math.abs(l.s) > 0.7).slice(0, 12);
    const vert = lineas.filter((l) => Math.abs(l.s) <= 0.7).slice(0, 12);
    // Los bordes de la foto, para la hoja que se sale del encuadre.
    horiz.push(nuevaRecta(90, 0, 0, 'arriba'), nuevaRecta(90, h - 1, 0, 'abajo'));
    vert.push(nuevaRecta(0, 0, 0, 'izq'), nuevaRecta(0, w - 1, 0, 'der'));
    for (const l of horiz.concat(vert)) { l.pos = pos(l); perfilDeRecta(mapa, l); }
    horiz.sort((a, b) => a.pos - b.pos);
    vert.sort((a, b) => a.pos - b.pos);
    const cruce = (l1, l2) => {
      const det = l1.c * l2.s - l1.s * l2.c;
      if (Math.abs(det) < 0.2) return null;
      return [(l1.rho * l2.s - l1.s * l2.rho) / det, (l1.c * l2.rho - l1.rho * l2.c) / det];
    };
    const mx = 0.02 * w, my = 0.02 * h;
    const fuera = (p) => p[0] < -mx || p[0] > w - 1 + mx || p[1] < -my || p[1] > h - 1 + my;
    const resultado = [];
    for (let a = 0; a < horiz.length; a++) for (let b = a + 1; b < horiz.length; b++) {
      for (let c = 0; c < vert.length; c++) for (let d = c + 1; d < vert.length; d++) {
        const T = horiz[a], B = horiz[b], L = vert[c], R = vert[d];
        const q = [cruce(T, L), cruce(T, R), cruce(B, R), cruce(B, L)];
        if (q.some((p) => !p || fuera(p))) continue;
        const rectas = [T, R, B, L];
        const puntaje = puntuar(q, w, h, (i, desde, hasta) => {
          const p0 = q[i], p1 = q[(i + 1) % 4];
          const en = (f) => [p0[0] + (p1[0] - p0[0]) * f, p0[1] + (p1[1] - p0[1]) * f];
          return respaldoEnRecta(rectas[i], en(desde), en(hasta));
        });
        if (puntaje > 0) resultado.push({ cuad: q, puntaje });
      }
    }
    resultado.sort((x, y) => y.puntaje - x.puntaje);
    return resultado.slice(0, 20);
  }

  // ---------- Refinamiento por bordes ----------

  function muestrear(canal, w, h, x, y) {
    x = limitar(x, 0, w - 1); y = limitar(y, 0, h - 1);
    const x0 = Math.floor(x), y0 = Math.floor(y), x1 = Math.min(w - 1, x0 + 1), y1 = Math.min(h - 1, y0 + 1);
    const fx = x - x0, fy = y - y0;
    return (canal[y0 * w + x0] * (1 - fx) + canal[y0 * w + x1] * fx) * (1 - fy)
      + (canal[y1 * w + x0] * (1 - fx) + canal[y1 * w + x1] * fx) * fy;
  }

  /** Recta (nx, ny, c) con nx*x + ny*y = c que mejor pasa por los puntos. */
  function ajustarRecta(pts) {
    const n = pts.length;
    let mx = 0, my = 0;
    for (const p of pts) { mx += p[0]; my += p[1]; }
    mx /= n; my /= n;
    let sxx = 0, syy = 0, sxy = 0;
    for (const p of pts) { const dx = p[0] - mx, dy = p[1] - my; sxx += dx * dx; syy += dy * dy; sxy += dx * dy; }
    const ang = 0.5 * Math.atan2(2 * sxy, sxx - syy); // dirección de la recta
    const nx = -Math.sin(ang), ny = Math.cos(ang);
    return { nx, ny, c: nx * mx + ny * my };
  }

  /**
   * Mueve cada lado del cuadrilátero hasta el borde más marcado que haya
   * cerca. La ventana mira poco hacia adentro (ahí está el texto de la hoja,
   * con mucho más contraste que el borde) y más hacia afuera, que es donde
   * queda el borde real cuando una sombra deja la hoja del color de la mesa.
   */
  function refinarLados(gris, w, h, cuad) {
    const cx = cuad.reduce((s, p) => s + p[0], 0) / 4, cy = cuad.reduce((s, p) => s + p[1], 0) / 4;
    const afuera = Math.max(8, Math.round(Math.min(w, h) * 0.09)), adentro = 3;
    const rectas = [];
    for (let i = 0; i < 4; i++) {
      const a = cuad[i], b = cuad[(i + 1) % 4];
      const largo = Math.hypot(b[0] - a[0], b[1] - a[1]);
      if (largo < 10) return null;
      const dx = (b[0] - a[0]) / largo, dy = (b[1] - a[1]) / largo;
      let nx = dy, ny = -dx;
      if (nx * ((a[0] + b[0]) / 2 - cx) + ny * ((a[1] + b[1]) / 2 - cy) < 0) { nx = -nx; ny = -ny; }
      let pts = [];
      for (let k = 0; k < 28; k++) {
        const t = 0.1 + 0.8 * (k / 27);
        const px = a[0] + (b[0] - a[0]) * t, py = a[1] + (b[1] - a[1]) * t;
        const perfil = [];
        let gMax = 0;
        for (let s = -adentro; s <= afuera; s += 0.5) {
          const g = Math.abs(muestrear(gris, w, h, px + nx * (s + 1.5), py + ny * (s + 1.5)) - muestrear(gris, w, h, px + nx * (s - 1.5), py + ny * (s - 1.5)));
          perfil.push([s, g]);
          if (g > gMax) gMax = g;
        }
        // El borde significativo más cercano al lado actual, no el más fuerte de la
        // ventana: el más fuerte suele ser otro objeto (un cuadro, una ventana).
        let cerca = -1;
        for (let k = 0; k < perfil.length; k++) {
          if (perfil[k][1] >= gMax * 0.5 && (cerca < 0 || Math.abs(perfil[k][0]) < Math.abs(perfil[cerca][0]))) cerca = k;
        }
        if (gMax >= 4 && cerca >= 0) {
          // Dentro de ese mismo borde, el punto más marcado (la rampa empieza antes del pico).
          let a = cerca, b = cerca;
          while (a > 0 && perfil[a - 1][1] >= gMax * 0.5) a--;
          while (b < perfil.length - 1 && perfil[b + 1][1] >= gMax * 0.5) b++;
          let pico = a;
          for (let k = a; k <= b; k++) if (perfil[k][1] > perfil[pico][1]) pico = k;
          pts.push([px + nx * perfil[pico][0], py + ny * perfil[pico][0]]);
        }
      }
      if (pts.length < 10) return null;
      let r = ajustarRecta(pts);
      // Descarta lo que se aparta de la recta (otro objeto, un trazo) y vuelve a ajustar.
      for (let vuelta = 0; vuelta < 2; vuelta++) {
        const res = pts.map((p) => Math.abs(r.nx * p[0] + r.ny * p[1] - r.c));
        const orden = res.slice().sort((p, q) => p - q);
        const tope = Math.max(1.5, orden[orden.length >> 1] * 2.5);
        const buenos = pts.filter((_, k) => res[k] <= tope);
        if (buenos.length < 8) return null;
        pts = buenos;
        r = ajustarRecta(pts);
      }
      rectas.push(r);
    }
    const nuevas = [];
    for (let i = 0; i < 4; i++) {
      const r1 = rectas[(i + 3) % 4], r2 = rectas[i];
      const det = r1.nx * r2.ny - r1.ny * r2.nx;
      if (Math.abs(det) < 1e-3) return null;
      nuevas.push([(r1.c * r2.ny - r1.ny * r2.c) / det, (r1.nx * r2.c - r1.c * r2.nx) / det]);
    }
    for (let i = 0; i < 4; i++) {
      if (Math.hypot(nuevas[i][0] - cuad[i][0], nuevas[i][1] - cuad[i][1]) > afuera * 2.2) return null;
    }
    const razon = areaPoligono(nuevas) / areaPoligono(cuad);
    return razon > 0.92 && razon < 1.4 ? nuevas : null;
  }

  // ---------- Detección ----------

  function candidatosDeMascara(mascara, w, h) {
    let m = morfologia(mascara, w, h, false);
    m = morfologia(m, w, h, true);
    m = morfologia(m, w, h, true);
    m = morfologia(m, w, h, false);
    m = rellenarHuecos(m, w, h);
    const mancha = manchaMasGrande(m, w, h);
    if (!mancha || mancha.area < w * h * 0.08) return null;
    const hull = envolvente(mancha.puntos);
    if (hull.length < 3) return null;
    const cuad = mejorCuadrilatero(hull);
    if (!cuad) return null;
    const areaHull = areaPoligono(hull);
    const areaCuad = areaPoligono(cuad);
    const fraccion = areaCuad / (w * h);
    if (fraccion < 0.1 || fraccion > 0.985) return null;
    const solidez = Math.min(1, mancha.area / areaHull);
    const ajuste = Math.min(1, areaHull / areaCuad);
    const puntaje = solidez * ajuste * (1 - Math.min(0.5, mancha.fraccionBorde * 2.5));
    return { cuad: ordenarEsquinas(cuad), puntaje };
  }

  /** Tres separaciones posibles del papel: más claro que el resto, más oscuro, o distinto del color del borde. */
  function mascarasDeColor(pequena, gris, w, h) {
    const t = otsu(gris);
    // Color típico del borde de la foto: casi siempre es fondo.
    const borde = [[], [], []];
    const grosor = Math.max(2, Math.round(Math.min(w, h) * 0.03));
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        if (x < grosor || y < grosor || x >= w - grosor || y >= h - grosor) {
          const i = (y * w + x) * 4;
          for (let c = 0; c < 3; c++) borde[c].push(pequena.data[i + c]);
        }
      }
    }
    const mediana = borde.map((v) => { v.sort((a, b) => a - b); return v[v.length >> 1]; });
    const distancia = new Uint8Array(w * h);
    for (let p = 0; p < w * h; p++) {
      const i = p * 4;
      const d = Math.abs(pequena.data[i] - mediana[0]) + Math.abs(pequena.data[i + 1] - mediana[1]) + Math.abs(pequena.data[i + 2] - mediana[2]);
      distancia[p] = Math.min(255, d / 1.5);
    }
    const distanciaSuave = desenfocar(distancia, w, h, 2);
    const td = otsu(distanciaSuave);
    return [
      Uint8Array.from(gris, (v) => (v > t ? 1 : 0)),
      Uint8Array.from(gris, (v) => (v <= t ? 1 : 0)),
      Uint8Array.from(distanciaSuave, (v) => (v > td ? 1 : 0)),
    ];
  }

  // Por debajo de este puntaje no hay hoja creíble: se devuelve la foto entera (no se recorta nada).
  const PUNTAJE_MINIMO = 0.4;

  /**
   * opciones.previa: las esquinas detectadas en el cuadro anterior de la cámara; un candidato
   * parecido recibe un pequeño premio, así el marco no salta entre dos parecidos.
   * opciones.punto: {x, y} en píxeles de la imagen que la persona tocó; se elige la hoja que lo contiene.
   */
  function detectarEsquinas(imagen, opciones = {}) {
    const pequena = reducir(imagen, LADO_ANALISIS);
    const { width: w, height: h } = pequena;
    const fx = imagen.width / w, fy = imagen.height / h;
    const gris = desenfocar(aGris(pequena), w, h, 2);
    const mapa = mapaDeBordes(pequena, w, h);

    // Candidatos de rectas (con los bordes de la foto incluidos)…
    const candidatos = candidatosDeRectas(mapa, rectasDeHough(mapa));
    // …y de manchas de color, tal cual y con los lados afinados contra el borde real.
    for (const m of mascarasDeColor(pequena, gris, w, h)) {
      const c = candidatosDeMascara(m, w, h);
      if (!c) continue;
      candidatos.push({ cuad: c.cuad });
      const afinadas = refinarLados(gris, w, h, c.cuad);
      if (afinadas) candidatos.push({ cuad: ordenarEsquinas(afinadas) });
    }

    // El mismo juez para todos, midiendo ahora directo sobre el mapa de bordes.
    const previa = opciones.previa && opciones.previa.map((p) => [p.x / fx - 0.5, p.y / fy - 0.5]);
    const cerca = 0.05 * Math.hypot(w, h);
    let mejor = null;
    const puntuados = [];
    for (const { cuad: q } of candidatos) {
      let puntaje = puntuar(q, w, h, (i, desde, hasta) => {
        const p0 = q[i], p1 = q[(i + 1) % 4];
        const en = (f) => [p0[0] + (p1[0] - p0[0]) * f, p0[1] + (p1[1] - p0[1]) * f];
        return respaldoSegmento(mapa, en(desde), en(hasta));
      });
      if (puntaje > 0 && previa && q.every((p, i) => Math.hypot(p[0] - previa[i][0], p[1] - previa[i][1]) < cerca)) puntaje += 0.05;
      if (puntaje > 0) puntuados.push({ cuad: q, puntaje });
      if (!mejor || puntaje > mejor.puntaje) mejor = { cuad: q, puntaje };
    }
    // Otras hojas posibles (un ticket sobre un formulario son dos): las mejores que no se parezcan
    // a una ya elegida. La pantalla las ofrece con «Probar otra hoja».
    puntuados.sort((a, b) => b.puntaje - a.puntaje);
    // La persona tocó la hoja que quiere: gana la hoja más chica que contiene ese punto, entre
    // las que tienen un puntaje razonable (un ticket sobre un formulario: tocar el ticket da el ticket).
    if (opciones.punto) {
      const p = [opciones.punto.x / fx - 0.5, opciones.punto.y / fy - 0.5];
      const contiene = (q) => [0, 1, 2, 3].every((i) => cruz(q[i], q[(i + 1) % 4], p) >= 0);
      const conPunto = puntuados.filter((c) => c.puntaje >= PUNTAJE_MINIMO * 0.75 && contiene(c.cuad));
      if (conPunto.length) {
        const tope = conPunto[0].puntaje * 0.5;
        mejor = conPunto.filter((c) => c.puntaje >= tope)
          .reduce((a, b) => (areaPoligono(b.cuad) < areaPoligono(a.cuad) ? b : a));
        mejor = { cuad: mejor.cuad, puntaje: Math.max(mejor.puntaje, PUNTAJE_MINIMO) };
      }
    }
    const distintas = [];
    for (const c of puntuados) {
      if (distintas.length >= 4 || c.puntaje < PUNTAJE_MINIMO) break;
      if (distintas.every((d) => d.cuad.some((p, i) => Math.hypot(p[0] - c.cuad[i][0], p[1] - c.cuad[i][1]) > 0.08 * Math.hypot(w, h)))) distintas.push(c);
    }
    const aImagen = (v, n, f, tope) => (v <= 0.6 ? 0 : v >= n - 1.6 ? tope : limitar((v + 0.5) * f, 0, tope));
    const aPuntos = (q) => q.map(([x, y]) => ({ x: aImagen(x, w, fx, imagen.width - 1), y: aImagen(y, h, fy, imagen.height - 1) }));

    const W = imagen.width - 1, H = imagen.height - 1;
    if (!mejor || mejor.puntaje < PUNTAJE_MINIMO) {
      return {
        esquinas: [{ x: 0, y: 0 }, { x: W, y: 0 }, { x: W, y: H }, { x: 0, y: H }],
        confianza: mejor ? Math.max(0, mejor.puntaje) : 0,
        encontrado: false,
      };
    }
    // +0.5: el centro de un píxel reducido cae a medio píxel del origen. Lo que está sobre el
    // borde de la foto va exactamente al borde.
    return {
      esquinas: aPuntos(mejor.cuad),
      confianza: mejor.puntaje,
      encontrado: true,
      alternativas: distintas.map((c) => aPuntos(c.cuad)),
    };
  }

  // ---------- Perspectiva ----------

  const dist = (a, b) => Math.hypot(a.x - b.x, a.y - b.y);

  /**
   * Tamaño de la página enderezada: el largo de los lados de la hoja, con el
   * lado mayor limitado a `maxLado`. Si la proporción queda a menos de un 5 %
   * de A4 o de Carta se ajusta a ella (la perspectiva la deforma un poco).
   */
  function tamanoSalida(esquinas, maxLado = 2400, ajustarProporcion = true) {
    const [a, b, c, d] = esquinas;
    let ancho = (dist(a, b) + dist(d, c)) / 2;
    let alto = (dist(a, d) + dist(b, c)) / 2;
    if (ajustarProporcion && ancho > 0 && alto > 0) {
      const mayor = Math.max(ancho, alto), menor = Math.min(ancho, alto);
      for (const razon of [Math.SQRT2, 11 / 8.5]) {
        if (Math.abs(mayor / menor - razon) / razon < 0.05) {
          if (ancho >= alto) alto = ancho / razon; else ancho = alto / razon;
          break;
        }
      }
    }
    const escala = Math.min(1, maxLado / Math.max(ancho, alto, 1));
    return { ancho: Math.max(1, Math.round(ancho * escala)), alto: Math.max(1, Math.round(alto * escala)) };
  }

  /** Coeficientes de la homografía que lleva el cuadrado unidad al cuadrilátero. */
  function cuadradoACuadrilatero(q) {
    const [p0, p1, p2, p3] = q;
    const sx = p0.x - p1.x + p2.x - p3.x;
    const sy = p0.y - p1.y + p2.y - p3.y;
    if (Math.abs(sx) < 1e-9 && Math.abs(sy) < 1e-9) {
      return { a: p1.x - p0.x, b: p3.x - p0.x, c: p0.x, d: p1.y - p0.y, e: p3.y - p0.y, f: p0.y, g: 0, h: 0 };
    }
    const dx1 = p1.x - p2.x, dx2 = p3.x - p2.x, dy1 = p1.y - p2.y, dy2 = p3.y - p2.y;
    const det = dx1 * dy2 - dx2 * dy1;
    const g = (sx * dy2 - dx2 * sy) / det;
    const h = (dx1 * sy - sx * dy1) / det;
    return {
      a: p1.x - p0.x + g * p1.x, b: p3.x - p0.x + h * p3.x, c: p0.x,
      d: p1.y - p0.y + g * p1.y, e: p3.y - p0.y + h * p3.y, f: p0.y, g, h,
    };
  }

  /**
   * Recorta el cuadrilátero y lo deja rectangular, de `ancho` x `alto` píxeles.
   * `margen` (fracción del lado) se queda un pelo adentro del borde marcado:
   * en el borde mismo se mezcla un poco de mesa y deja una línea punteada.
   */
  function enderezar(imagen, esquinas, ancho, alto, margen = 0.003) {
    const m = cuadradoACuadrilatero(esquinas);
    const { data, width, height } = imagen;
    const salida = new Uint8ClampedArray(ancho * alto * 4);
    for (let y = 0; y < alto; y++) {
      const v = margen + (1 - 2 * margen) * ((y + 0.5) / alto);
      for (let x = 0; x < ancho; x++) {
        const u = margen + (1 - 2 * margen) * ((x + 0.5) / ancho);
        const k = m.g * u + m.h * v + 1;
        const sx = (m.a * u + m.b * v + m.c) / k - 0.5;
        const sy = (m.d * u + m.e * v + m.f) / k - 0.5;
        const x0 = Math.floor(sx), y0 = Math.floor(sy);
        const fx = sx - x0, fy = sy - y0;
        const xa = limitar(x0, 0, width - 1), xb = limitar(x0 + 1, 0, width - 1);
        const ya = limitar(y0, 0, height - 1), yb = limitar(y0 + 1, 0, height - 1);
        const i00 = (ya * width + xa) * 4, i10 = (ya * width + xb) * 4;
        const i01 = (yb * width + xa) * 4, i11 = (yb * width + xb) * 4;
        const o = (y * ancho + x) * 4;
        for (let c = 0; c < 3; c++) {
          salida[o + c] = (data[i00 + c] * (1 - fx) + data[i10 + c] * fx) * (1 - fy)
            + (data[i01 + c] * (1 - fx) + data[i11 + c] * fx) * fy;
        }
        salida[o + 3] = 255;
      }
    }
    return { data: salida, width: ancho, height: alto };
  }

  /** Gira la imagen `cuartos` cuartos de vuelta en sentido horario (devuelve una nueva). */
  function girar(imagen, cuartos) {
    const q = ((cuartos % 4) + 4) % 4;
    if (q === 0) return imagen;
    const { data, width: w, height: h } = imagen;
    const ancho = q === 2 ? w : h, alto = q === 2 ? h : w;
    const salida = new Uint8ClampedArray(ancho * alto * 4);
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const nx = q === 1 ? h - 1 - y : q === 2 ? w - 1 - x : y;
        const ny = q === 1 ? x : q === 2 ? h - 1 - y : w - 1 - x;
        const i = (y * w + x) * 4, o = (ny * ancho + nx) * 4;
        salida[o] = data[i]; salida[o + 1] = data[i + 1]; salida[o + 2] = data[i + 2]; salida[o + 3] = data[i + 3];
      }
    }
    return { data: salida, width: ancho, height: alto };
  }

  // ---------- Mejora tipo escáner ----------

  /**
   * Iluminación de fondo de un canal: máximo por bloques (ignora los trazos
   * oscuros del texto), suavizado, y ampliado de vuelta con interpolación.
   */
  function fondoDeCanal(canal, w, h) {
    const bloque = Math.max(6, Math.round(Math.min(w, h) / 50));
    const bw = Math.max(2, Math.ceil(w / bloque)), bh = Math.max(2, Math.ceil(h / bloque));
    const celdas = new Float32Array(bw * bh);
    for (let by = 0; by < bh; by++) {
      for (let bx = 0; bx < bw; bx++) {
        let max = 0;
        const y1 = Math.min(h, (by + 1) * bloque), x1 = Math.min(w, (bx + 1) * bloque);
        for (let y = by * bloque; y < y1; y++) {
          for (let x = bx * bloque; x < x1; x++) { const v = canal[y * w + x]; if (v > max) max = v; }
        }
        celdas[by * bw + bx] = max;
      }
    }
    const suave = desenfocar(celdas, bw, bh, 3);
    const fondo = new Float32Array(w * h);
    for (let y = 0; y < h; y++) {
      const gy = limitar(y / bloque - 0.5, 0, bh - 1);
      const y0 = Math.floor(gy), y1 = Math.min(bh - 1, y0 + 1), fy = gy - y0;
      for (let x = 0; x < w; x++) {
        const gx = limitar(x / bloque - 0.5, 0, bw - 1);
        const x0 = Math.floor(gx), x1 = Math.min(bw - 1, x0 + 1), fx = gx - x0;
        fondo[y * w + x] = (suave[y0 * bw + x0] * (1 - fx) + suave[y0 * bw + x1] * fx) * (1 - fy)
          + (suave[y1 * bw + x0] * (1 - fx) + suave[y1 * bw + x1] * fx) * fy;
      }
    }
    return fondo;
  }

  /**
   * Mejora la imagen en el lugar.
   * modo: 'original' | 'color' | 'gris' | 'bn'.
   * opciones.brillo y opciones.contraste, de -100 a 100 (0 = sin cambio). En
   * blanco y negro el brillo mueve el umbral y el contraste no se usa.
   */
  function mejorar(imagen, modo, opciones = {}) {
    const brillo = opciones.brillo || 0;
    const contraste = opciones.contraste || 0;
    const { data, width: w, height: h } = imagen;
    const n = w * h;
    const factorContraste = 1 + contraste / 100;
    const ajustar = (v) => limitar((v - 128) * factorContraste + 128 + brillo * 1.28, 0, 255);

    if (modo === 'original') {
      if (!brillo && !contraste) return imagen;
      for (let i = 0; i < data.length; i += 4) {
        data[i] = ajustar(data[i]); data[i + 1] = ajustar(data[i + 1]); data[i + 2] = ajustar(data[i + 2]);
      }
      return imagen;
    }

    if (modo === 'color') {
      // Cada canal se divide por su propia iluminación de fondo: quita sombras
      // y el tinte de la luz, y el papel queda blanco.
      const canales = [0, 1, 2].map((c) => {
        const ch = new Uint8Array(n);
        for (let p = 0; p < n; p++) ch[p] = data[p * 4 + c];
        return ch;
      });
      const fondos = canales.map((ch) => fondoDeCanal(ch, w, h));
      for (let p = 0; p < n; p++) {
        for (let c = 0; c < 3; c++) {
          const normal = limitar((canales[c][p] / Math.max(1, fondos[c][p])) * 255, 0, 255);
          // Un poco de curva para que el texto gane fuerza sin quemar los colores.
          data[p * 4 + c] = ajustar(normal < 255 ? 255 * Math.pow(normal / 255, 1.35) : 255);
        }
      }
      return imagen;
    }

    // Gris y blanco y negro parten de la luminancia normalizada.
    const gris = aGris(imagen);
    const fondo = fondoDeCanal(gris, w, h);
    const normal = new Uint8Array(n);
    for (let p = 0; p < n; p++) normal[p] = limitar((gris[p] / Math.max(1, fondo[p])) * 255, 0, 255);

    if (modo === 'gris') {
      for (let p = 0; p < n; p++) {
        const v = normal[p] < 255 ? ajustar(255 * Math.pow(normal[p] / 255, 1.35)) : 255;
        data[p * 4] = data[p * 4 + 1] = data[p * 4 + 2] = v;
      }
      return imagen;
    }

    // Blanco y negro: umbral de Otsu sobre la imagen ya sin sombras, acotado
    // para que ni el papel gris ni un lápiz suave se pierdan.
    const base = limitar(otsu(normal), 90, 215);
    const umbral = limitar(base + brillo * 0.8, 40, 245);
    for (let p = 0; p < n; p++) {
      const v = normal[p] > umbral ? 255 : 0;
      data[p * 4] = data[p * 4 + 1] = data[p * 4 + 2] = v;
    }
    return imagen;
  }

  return { detectarEsquinas, tamanoSalida, enderezar, girar, mejorar };
})();

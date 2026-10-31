/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
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
        let mejor = 0, gMax = 0;
        for (let s = -adentro; s <= afuera; s += 0.5) {
          const g = Math.abs(muestrear(gris, w, h, px + nx * (s + 1.5), py + ny * (s + 1.5)) - muestrear(gris, w, h, px + nx * (s - 1.5), py + ny * (s - 1.5)));
          if (g > gMax) { gMax = g; mejor = s; }
        }
        if (gMax >= 4) pts.push([px + nx * mejor, py + ny * mejor]);
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

  function detectarEsquinas(imagen) {
    const pequena = reducir(imagen, LADO_ANALISIS);
    const { width: w, height: h } = pequena;
    const gris = desenfocar(aGris(pequena), w, h, 2);
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

    const mascaras = [
      gris.map((v) => (v > t ? 1 : 0)),
      gris.map((v) => (v <= t ? 1 : 0)),
      distanciaSuave.map((v) => (v > td ? 1 : 0)),
    ];
    let mejor = null;
    for (const m of mascaras) {
      const cand = candidatosDeMascara(Uint8Array.from(m), w, h);
      if (cand && (!mejor || cand.puntaje > mejor.puntaje)) mejor = cand;
    }
    if (!mejor || mejor.puntaje < 0.6) {
      const mx = imagen.width * 0.04, my = imagen.height * 0.04;
      return {
        esquinas: [
          { x: mx, y: my }, { x: imagen.width - mx, y: my },
          { x: imagen.width - mx, y: imagen.height - my }, { x: mx, y: imagen.height - my },
        ],
        confianza: mejor ? mejor.puntaje : 0,
        encontrado: false,
      };
    }
    const f = pequena.factor;
    // `gris` ya está suavizado; sirve para buscar el borde real de cada lado.
    const afinadas = refinarLados(gris, w, h, mejor.cuad);
    if (afinadas) mejor.cuad = ordenarEsquinas(afinadas);
    return {
      // +0.5: el centro de un píxel reducido cae a medio píxel del origen.
      esquinas: mejor.cuad.map(([x, y]) => ({
        x: limitar((x + 0.5) * f, 0, imagen.width - 1),
        y: limitar((y + 0.5) * (imagen.height / h), 0, imagen.height - 1),
      })),
      confianza: mejor.puntaje,
      encontrado: true,
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

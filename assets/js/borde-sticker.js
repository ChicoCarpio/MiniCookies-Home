/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Borde de los stickers como figura matemática, no como píxeles (2026-10-08,
// idea del usuario: sacar el contorno del recorte como líneas, agrandarlo,
// unir los tramos con curvas y rellenar).
//
// Antes el borde se hacía "estampando" la silueta 24 a 48 veces en círculo:
// a 512 px salía escalonado y con ondas, y no se podía hacer a la resolución
// de la foto sin trabar la página. Ahora:
//
//  1. Campo de distancias exacto (Felzenszwalb y Huttenlocher, 2012): para
//     cada punto, a qué distancia está del recorte. El borde son los puntos a
//     menos del grosor: la curva paralela al contorno, del mismo grosor en
//     todos lados, con las esquinas que salen ya redondeadas. Desplazar
//     tramos de línea hacia afuera da lo mismo en una figura simple, pero en
//     el pelo o entre los dedos los tramos se cruzan entre sí; el campo no.
//  2. Suavidad: un "cierre". Se agranda el grosor más la suavidad y se vuelve
//     a achicar la suavidad, también con un campo de distancias. Las entradas
//     más angostas que el doble de la suavidad se rellenan y cada rincón
//     queda redondeado con ese radio: el contorno limpio de un sticker.
//  3. Contorno a líneas: "marching squares" con interpolación, así cada punto
//     cae entre píxeles, donde de verdad cruza el contorno.
//  4. Curvas: cada tramo se une al siguiente con una curva cuadrática que
//     pasa por los puntos medios, y se rellena como figura (Path2D).
//
// El cálculo se hace a lo sumo a LADO_CALCULO px y la figura se escala al
// tamaño real: como son curvas, no pierden nitidez, y una foto de 4000 px no
// necesita cientos de MB de memoria en el celular.
//
// Uso:
//   const figura = BordeSticker.figura(canvasDelRecorte, radio, suavidad);
//   ctx.fill(figura, 'evenodd');   // en coordenadas del canvas del recorte
// radio y suavidad en píxeles de ese canvas.

const BordeSticker = (() => {
  const LADO_CALCULO = 2048;
  const INFINITO = 1e20;

  // Transformada de distancia en una dimensión (distancias al cuadrado), del
  // artículo citado arriba: la envolvente inferior de las parábolas.
  function distancias1D(f, n, d, v, z) {
    let k = 0;
    v[0] = 0; z[0] = -INFINITO; z[1] = INFINITO;
    for (let q = 1; q < n; q++) {
      let s = ((f[q] + q * q) - (f[v[k]] + v[k] * v[k])) / (2 * q - 2 * v[k]);
      while (s <= z[k]) {
        k--;
        s = ((f[q] + q * q) - (f[v[k]] + v[k] * v[k])) / (2 * q - 2 * v[k]);
      }
      k++;
      v[k] = q; z[k] = s; z[k + 1] = INFINITO;
    }
    k = 0;
    for (let q = 0; q < n; q++) {
      while (z[k + 1] < q) k++;
      const dq = q - v[k];
      d[q] = dq * dq + f[v[k]];
    }
  }

  // Distancia (euclídea, exacta) de cada píxel al píxel marcado más cercano.
  function campoDeDistancias(marcado, w, h) {
    const g = new Float32Array(w * h);
    for (let i = 0; i < g.length; i++) g[i] = marcado[i] ? 0 : INFINITO;
    const n = Math.max(w, h);
    const f = new Float64Array(n), d = new Float64Array(n);
    const v = new Int32Array(n), z = new Float64Array(n + 1);
    for (let x = 0; x < w; x++) {
      for (let y = 0; y < h; y++) f[y] = g[y * w + x];
      distancias1D(f, h, d, v, z);
      for (let y = 0; y < h; y++) g[y * w + x] = d[y];
    }
    for (let y = 0; y < h; y++) {
      const fila = y * w;
      for (let x = 0; x < w; x++) f[x] = g[fila + x];
      distancias1D(f, w, d, v, z);
      for (let x = 0; x < w; x++) g[fila + x] = Math.sqrt(d[x]);
    }
    return g;
  }

  // Contornos donde el campo cruza "umbral" (adentro: valor < umbral), como
  // listas cerradas de puntos [x0, y0, x1, y1, …]. Cada punto se interpola
  // sobre el lado de la celda que cruza el contorno.
  function contornos(campo, w, h, umbral) {
    const puntos = new Map(); // id de lado → [x, y]
    const vecinos = new Map(); // id de lado → ids de los lados unidos a él
    const punto = (id, x, y) => { if (!puntos.has(id)) puntos.set(id, [x, y]); return id; };
    const ladoArriba = (x, y) => {
      const a = campo[y * w + x], b = campo[y * w + x + 1];
      return punto((y * w + x) * 2, x + (umbral - a) / (b - a), y);
    };
    const ladoIzquierdo = (x, y) => {
      const a = campo[y * w + x], d = campo[(y + 1) * w + x];
      return punto((y * w + x) * 2 + 1, x, y + (umbral - a) / (d - a));
    };
    const unir = (p, q) => {
      if (!vecinos.has(p)) vecinos.set(p, []);
      if (!vecinos.has(q)) vecinos.set(q, []);
      vecinos.get(p).push(q);
      vecinos.get(q).push(p);
    };
    for (let y = 0; y < h - 1; y++) {
      for (let x = 0; x < w - 1; x++) {
        const a = campo[y * w + x] < umbral, b = campo[y * w + x + 1] < umbral;
        const c = campo[(y + 1) * w + x + 1] < umbral, d = campo[(y + 1) * w + x] < umbral;
        const caso = (a ? 1 : 0) | (b ? 2 : 0) | (c ? 4 : 0) | (d ? 8 : 0);
        if (caso === 0 || caso === 15) continue;
        const arriba = () => ladoArriba(x, y), abajo = () => ladoArriba(x, y + 1);
        const izquierda = () => ladoIzquierdo(x, y), derecha = () => ladoIzquierdo(x + 1, y);
        switch (caso) {
          case 1: case 14: unir(arriba(), izquierda()); break;
          case 2: case 13: unir(arriba(), derecha()); break;
          case 3: case 12: unir(izquierda(), derecha()); break;
          case 4: case 11: unir(derecha(), abajo()); break;
          case 6: case 9: unir(arriba(), abajo()); break;
          case 7: case 8: unir(izquierda(), abajo()); break;
          case 5: case 10: {
            // Esquinas opuestas: decide el centro de la celda.
            const centro = (campo[y * w + x] + campo[y * w + x + 1] + campo[(y + 1) * w + x + 1] + campo[(y + 1) * w + x]) / 4 < umbral;
            if ((caso === 5) === centro) { unir(arriba(), derecha()); unir(izquierda(), abajo()); }
            else { unir(arriba(), izquierda()); unir(derecha(), abajo()); }
            break;
          }
        }
      }
    }
    // Recorrer cada lazo: cada lado cruzado tiene exactamente dos vecinos.
    const usados = new Set();
    const lazos = [];
    for (const inicio of vecinos.keys()) {
      if (usados.has(inicio)) continue;
      const lazo = [];
      let anterior = -1, actual = inicio;
      while (actual !== undefined && !usados.has(actual)) {
        usados.add(actual);
        const [x, y] = puntos.get(actual);
        lazo.push(x, y);
        const [p, q] = vecinos.get(actual);
        const siguiente = p !== anterior ? p : q;
        anterior = actual;
        actual = siguiente;
      }
      if (lazo.length >= 6) lazos.push(lazo);
    }
    return lazos;
  }

  // Une los puntos de cada lazo con curvas cuadráticas por sus puntos medios
  // (cada punto queda como control): sin quiebres, siguiendo la forma.
  function curvas(lazos, escala, desplazamiento) {
    const figura = new Path2D();
    // El punto i de la grilla es el centro del píxel i, que en el canvas está
    // en i + 0,5: sin sumarlo, el borde salía corrido medio píxel.
    const X = (v) => (v + 0.5 - desplazamiento) * escala;
    for (const l of lazos) {
      const n = l.length / 2;
      const medio = (i) => {
        const a = (i % n) * 2, b = ((i + 1) % n) * 2;
        return [X((l[a] + l[b]) / 2), X((l[a + 1] + l[b + 1]) / 2)];
      };
      const [mx, my] = medio(n - 1);
      figura.moveTo(mx, my);
      for (let i = 0; i < n; i++) {
        const [nx, ny] = medio(i);
        figura.quadraticCurveTo(X(l[i * 2]), X(l[i * 2 + 1]), nx, ny);
      }
      figura.closePath();
    }
    return figura;
  }

  function figura(canvasRecorte, radio, suavidad = 0) {
    const ancho = canvasRecorte.width, alto = canvasRecorte.height;
    const k = Math.min(1, LADO_CALCULO / Math.max(ancho, alto));
    const r = radio * k, s = suavidad * k;
    const margen = Math.ceil(r + s) + 3;
    const w = Math.max(1, Math.round(ancho * k)) + margen * 2;
    const h = Math.max(1, Math.round(alto * k)) + margen * 2;

    // La silueta (alfa ≥ 128) a la resolución de cálculo, con margen vacío.
    const lienzo = document.createElement('canvas');
    lienzo.width = w; lienzo.height = h;
    const ctx = lienzo.getContext('2d', { willReadFrequently: true });
    ctx.drawImage(canvasRecorte, margen, margen, w - margen * 2, h - margen * 2);
    const alfa = ctx.getImageData(0, 0, w, h).data;
    const silueta = new Uint8Array(w * h);
    for (let i = 0; i < silueta.length; i++) silueta[i] = alfa[i * 4 + 3] >= 128 ? 1 : 0;

    // Distancia al recorte (−0,5: del centro del píxel a su borde).
    const distancia = campoDeDistancias(silueta, w, h);
    let lazos;
    if (s < 0.5) {
      lazos = contornos(distancia, w, h, r + 0.5);
    } else {
      // Cierre: agrandar r + s, y achicar s midiendo desde lo que quedó afuera.
      const afuera = new Uint8Array(w * h);
      for (let i = 0; i < afuera.length; i++) afuera[i] = distancia[i] - 0.5 > r + s ? 1 : 0;
      const haciaAfuera = campoDeDistancias(afuera, w, h);
      for (let i = 0; i < haciaAfuera.length; i++) haciaAfuera[i] = -haciaAfuera[i];
      lazos = contornos(haciaAfuera, w, h, -(s + 0.5));
    }
    return curvas(lazos, 1 / k, margen);
  }

  return { figura, LADO_CALCULO };
})();

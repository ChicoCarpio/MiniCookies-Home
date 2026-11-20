/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Trazador de contornos propio (2026-11-19): de un mapa de colores por píxel
// (cada celda dice a qué color de la paleta pertenece) saca las figuras de
// cada color como caminos de SVG, con rectas y curvas en vez de escalones.
// Reemplaza a imagetracerjs (Unlicense), de la que se usaba solo esta parte.
//
// Pasos, por cada color:
//  1. Nodos de borde: cada esquina entre cuatro píxeles recibe un número del
//     0 al 15 según cuáles de los cuatro son de este color.
//  2. Recorrido: se siguen los nodos con una tabla de giros; sale un contorno
//     cerrado por figura y por agujero. Los contornos de menos de "omitir"
//     puntos se descartan (manchas), y cada agujero se asigna a la figura que
//     lo contiene.
//  3. Se toman los puntos medios entre esquinas y la dirección de cada tramo
//     (8 direcciones).
//  4. Tramos de solo dos direcciones se ajustan con una recta; si falla
//     (error de distancia al cuadrado > RECTA), con una curva cuadrática
//     por el punto de mayor error; si falla (> CURVA), se parte ahí y se
//     repite con cada mitad.
//
// Uso: TrazadorContornos.vectorizar(indice, w, h, k) con "indice" una matriz
// plana de w*h (valores 0..k-1; negativo = sin color). Devuelve
// [{ color, d, area }] con "d" en formato de SVG, de la figura más grande a la
// más chica. El origen está desplazado (-1, -1): el borde de la matriz con
// relleno cuenta como una celda (quien dibuja resta ese desplazamiento).
window.TrazadorContornos = (function () {
  const OMITIR = 8;   // contornos con menos puntos que esto se descartan
  const RECTA = 1;    // error máximo (al cuadrado) para ajustar una recta
  const CURVA = 1;    // idem para una curva cuadrática

  // Tabla de giros del recorrido: [nodo][dirección] = [nodo que queda, nueva
  // dirección, desplazamiento en x, desplazamiento en y]. Direcciones:
  // 0 derecha, 1 arriba, 2 izquierda, 3 abajo. Los nodos 0 y 15 no existen
  // en un contorno.
  const N = [-1, -1, -1, -1];
  const GIROS = [
    [N, N, N, N],
    [[0, 1, 0, -1], N, N, [0, 2, -1, 0]],
    [N, N, [0, 1, 0, -1], [0, 0, 1, 0]],
    [[0, 0, 1, 0], N, [0, 2, -1, 0], N],
    [N, [0, 0, 1, 0], [0, 3, 0, 1], N],
    [[13, 3, 0, 1], [13, 2, -1, 0], [7, 1, 0, -1], [7, 0, 1, 0]],
    [N, [0, 1, 0, -1], N, [0, 3, 0, 1]],
    [[0, 3, 0, 1], [0, 2, -1, 0], N, N],
    [[0, 3, 0, 1], [0, 2, -1, 0], N, N],
    [N, [0, 1, 0, -1], N, [0, 3, 0, 1]],
    [[11, 1, 0, -1], [14, 0, 1, 0], [14, 3, 0, 1], [11, 2, -1, 0]],
    [N, [0, 0, 1, 0], [0, 3, 0, 1], N],
    [[0, 0, 1, 0], N, [0, 2, -1, 0], N],
    [N, N, [0, 1, 0, -1], [0, 0, 1, 0]],
    [[0, 1, 0, -1], N, N, [0, 2, -1, 0]],
    [N, N, N, N],
  ];

  // Matriz con una celda de relleno (-1) alrededor.
  function conBorde(indice, w, h) {
    const W = w + 2, H = h + 2;
    const a = new Int16Array(W * H).fill(-1);
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) a[(y + 1) * W + x + 1] = indice[y * w + x];
    return { a, W, H };
  }

  // 1. Número de nodo (0..15) de cada esquina para el color c.
  function nodos({ a, W, H }, c) {
    const nodo = new Uint8Array(W * H);
    for (let j = 1; j < H; j++) {
      for (let i = 1; i < W; i++) {
        nodo[j * W + i] = (a[(j - 1) * W + i - 1] === c ? 1 : 0) + (a[(j - 1) * W + i] === c ? 2 : 0)
          + (a[j * W + i - 1] === c ? 8 : 0) + (a[j * W + i] === c ? 4 : 0);
      }
    }
    return nodo;
  }

  function cajaContiene(grande, chica) {
    return grande[0] < chica[0] && grande[1] < chica[1] && grande[2] > chica[2] && grande[3] > chica[3];
  }

  function puntoEnPoligono(px, py, xs, ys) {
    let dentro = false;
    for (let i = 0, j = xs.length - 1; i < xs.length; j = i++) {
      if ((ys[i] > py) !== (ys[j] > py) && px < (xs[j] - xs[i]) * (py - ys[i]) / (ys[j] - ys[i]) + xs[i]) dentro = !dentro;
    }
    return dentro;
  }

  // 2. Recorre los nodos y arma los contornos cerrados.
  function recorrer(nodo, W, H) {
    const contornos = [];
    for (let j = 0; j < H; j++) {
      for (let i = 0; i < W; i++) {
        const t0 = nodo[j * W + i];
        if (t0 !== 4 && t0 !== 11) continue;
        const agujero = t0 === 11;
        const xs = [], ys = [];
        const caja = [i, j, i, j];
        let px = i, py = j, dir = 1;
        for (;;) {
          xs.push(px - 1); ys.push(py - 1);
          if (px - 1 < caja[0]) caja[0] = px - 1;
          if (px - 1 > caja[2]) caja[2] = px - 1;
          if (py - 1 < caja[1]) caja[1] = py - 1;
          if (py - 1 > caja[3]) caja[3] = py - 1;
          const g = GIROS[nodo[py * W + px]][dir];
          nodo[py * W + px] = g[0]; dir = g[1]; px += g[2]; py += g[3];
          if (px - 1 === xs[0] && py - 1 === ys[0]) break;
        }
        if (xs.length < OMITIR) continue;
        const c = { xs, ys, caja, agujero, hijos: [] };
        if (agujero) {
          // La figura que lo contiene: la más interior que tenga su primer punto adentro.
          let padre = 0, cajaPadre = [-1, -1, W + 1, H + 1];
          for (let k = 0; k < contornos.length; k++) {
            const o = contornos[k];
            if (!o.agujero && cajaContiene(o.caja, caja) && cajaContiene(cajaPadre, o.caja) && puntoEnPoligono(xs[0], ys[0], o.xs, o.ys)) {
              padre = k; cajaPadre = o.caja;
            }
          }
          contornos[padre].hijos.push(contornos.length);
        }
        contornos.push(c);
      }
    }
    return contornos;
  }

  // Dirección de un tramo entre dos puntos: 0 E, 1 SE, 2 S, 3 SO, 4 O, 5 NO, 6 N, 7 NE, 8 ninguna.
  function direccion(x1, y1, x2, y2) {
    if (x1 < x2) return y1 < y2 ? 1 : y1 > y2 ? 7 : 0;
    if (x1 > x2) return y1 < y2 ? 3 : y1 > y2 ? 5 : 4;
    return y1 < y2 ? 2 : y1 > y2 ? 6 : 8;
  }

  // 3. Puntos medios entre las esquinas, con la dirección hacia el siguiente.
  function puntosMedios(c) {
    const n = c.xs.length;
    const x = new Float64Array(n), y = new Float64Array(n), dir = new Uint8Array(n);
    for (let i = 0; i < n; i++) {
      const s = (i + 1) % n;
      x[i] = (c.xs[i] + c.xs[s]) / 2;
      y[i] = (c.ys[i] + c.ys[s]) / 2;
    }
    for (let i = 0; i < n; i++) {
      const s = (i + 1) % n;
      dir[i] = direccion(x[i], y[i], x[s], y[s]);
    }
    return { x, y, dir, n };
  }

  // 4. Ajuste de rectas y curvas: recorre el contorno por tramos de solo dos
  // direcciones y ajusta cada uno.
  function trazar(p) {
    const { dir, n } = p;
    const segmentos = [];
    let pcnt = 0;
    while (pcnt < n) {
      const tipo1 = dir[pcnt];
      let tipo2 = -1, fin = pcnt + 1;
      while ((dir[fin] === tipo1 || dir[fin] === tipo2 || tipo2 === -1) && fin < n - 1) {
        if (dir[fin] !== tipo1 && tipo2 === -1) tipo2 = dir[fin];
        fin++;
      }
      if (fin === n - 1) fin = 0;
      ajustar(p, pcnt, fin, segmentos);
      pcnt = fin > 0 ? fin : n;
    }
    return segmentos;
  }

  function ajustar(p, ini, fin, salida) {
    const { x, y, n } = p;
    if (fin > n || fin < 0) return;
    let tl = fin - ini; if (tl < 0) tl += n;
    const vx = (x[fin] - x[ini]) / tl, vy = (y[fin] - y[ini]) / tl;
    let errorPunto = ini, errorValor = 0, pasa = true;

    // Recta
    let k = (ini + 1) % n;
    while (k !== fin) {
      let pl = k - ini; if (pl < 0) pl += n;
      const qx = x[ini] + vx * pl, qy = y[ini] + vy * pl;
      const d2 = (x[k] - qx) * (x[k] - qx) + (y[k] - qy) * (y[k] - qy);
      if (d2 > RECTA) pasa = false;
      if (d2 > errorValor) { errorPunto = k; errorValor = d2; }
      k = (k + 1) % n;
    }
    if (pasa) { salida.push({ x1: x[ini], y1: y[ini], x2: x[fin], y2: y[fin] }); return; }

    // Curva cuadrática que pasa por el punto de mayor error
    const medio = errorPunto; pasa = true; errorValor = 0;
    let t = (medio - ini) / tl, t1 = (1 - t) * (1 - t), t2 = 2 * (1 - t) * t, t3 = t * t;
    const cx = (t1 * x[ini] + t3 * x[fin] - x[medio]) / -t2;
    const cy = (t1 * y[ini] + t3 * y[fin] - y[medio]) / -t2;
    k = ini + 1;
    while (k !== fin) {
      t = (k - ini) / tl; t1 = (1 - t) * (1 - t); t2 = 2 * (1 - t) * t; t3 = t * t;
      const qx = t1 * x[ini] + t2 * cx + t3 * x[fin];
      const qy = t1 * y[ini] + t2 * cy + t3 * y[fin];
      const d2 = (x[k] - qx) * (x[k] - qx) + (y[k] - qy) * (y[k] - qy);
      if (d2 > CURVA) pasa = false;
      if (d2 > errorValor) { errorPunto = k; errorValor = d2; }
      k = (k + 1) % n;
    }
    if (pasa) { salida.push({ x1: x[ini], y1: y[ini], x2: cx, y2: cy, x3: x[fin], y3: y[fin] }); return; }

    // Se parte en el punto del ajuste y se repite con cada mitad
    ajustar(p, ini, medio, salida);
    ajustar(p, medio, fin, salida);
  }

  // Camino de SVG de una figura, con sus agujeros recorridos al revés.
  function camino(seg, huecos) {
    let s = `M ${seg[0].x1} ${seg[0].y1} `;
    for (const g of seg) s += g.x3 === undefined ? `L ${g.x2} ${g.y2} ` : `Q ${g.x2} ${g.y2} ${g.x3} ${g.y3} `;
    s += 'Z ';
    for (const h of huecos) {
      const u = h[h.length - 1];
      s += u.x3 === undefined ? `M ${u.x2} ${u.y2} ` : `M ${u.x3} ${u.y3} `;
      for (let i = h.length - 1; i >= 0; i--) {
        const g = h[i];
        s += g.x3 === undefined ? `L ${g.x1} ${g.y1} ` : `Q ${g.x2} ${g.y2} ${g.x1} ${g.y1} `;
      }
      s += 'Z ';
    }
    return s;
  }

  function vectorizar(indice, w, h, k) {
    const m = conBorde(indice, w, h);
    const figuras = [];
    for (let c = 0; c < k; c++) {
      const contornos = recorrer(nodos(m, c), m.W, m.H);
      const trazados = contornos.map((o) => trazar(puntosMedios(o)));
      contornos.forEach((o, i) => {
        if (o.agujero) return;
        const [x0, y0, x1, y1] = o.caja;
        figuras.push({ color: c, d: camino(trazados[i], o.hijos.map((j) => trazados[j])), area: (x1 - x0) * (y1 - y0) });
      });
    }
    // Las chicas encima: tapan las uniones de las grandes
    return figuras.sort((p, q) => q.area - p.area);
  }

  return { vectorizar };
})();

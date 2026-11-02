/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Gráficos de superficie de PowerPoint (superficie 3D, de alambre, contorno y
// contorno de alambre): la malla de datos hecha triángulos, cortada por las
// bandas del eje de valores, con su sombreado, y las líneas del alambre. Lo
// usa pptx-grafico.js, que pone la cámara y dibuja.
//
// Medido en 26-graficos-superficie (imágenes de PowerPoint): cada celda de la
// malla (categoría × serie) se parte en dos triángulos por la diagonal que va
// de (categoría i, serie j+1) a (i+1, j); las bandas son los tramos del eje
// de valores (una por unidad principal) y cada triángulo, cortado por ellas,
// se pinta plano con el color de su banda por un factor de luz.
//
// Mundo (como PptxGrafico3D): x a la derecha, y hacia arriba, z al fondo.

const PptxSuperficie = (() => {
  // z[j][i]: valor de la serie j en la categoría i (null si falta).
  // niveles: los límites de las bandas, de menor a mayor.
  // mundo: { X(i), Z(j), Y(v) } → coordenadas del mundo.
  const EPS = 1e-9;
  const celdas = (z) => {
    const ns = z.length, nc = Math.max(0, ...z.map((f) => f.length));
    const lista = [];
    for (let j = 0; j + 1 < ns; j++) {
      for (let i = 0; i + 1 < nc; i++) {
        const v = (a, b) => z[b][a];
        const p00 = v(i, j), p10 = v(i + 1, j), p01 = v(i, j + 1), p11 = v(i + 1, j + 1);
        if ([p00, p10, p01, p11].some((x) => x == null || !Number.isFinite(x))) continue;
        // La diagonal va de (i, j+1) a (i+1, j).
        lista.push([[i, j], [i + 1, j], [i, j + 1]], [[i + 1, j], [i + 1, j + 1], [i, j + 1]]);
      }
    }
    return lista;
  };

  // Un polígono [{i, j, v}] recortado a lo que tiene v en [lo, hi] (Sutherland-Hodgman).
  function recortar(poli, lo, hi) {
    const cortar = (p, dentro, v0) => {
      const out = [];
      for (let k = 0; k < p.length; k++) {
        const a = p[k], b = p[(k + 1) % p.length];
        const da = dentro(a), db = dentro(b);
        if (da) out.push(a);
        if (da !== db) {
          const t = (v0 - a.v) / (b.v - a.v);
          out.push({ i: a.i + (b.i - a.i) * t, j: a.j + (b.j - a.j) * t, v: v0 });
        }
      }
      return out;
    };
    let p = cortar(poli, (q) => q.v >= lo - EPS, lo);
    if (p.length) p = cortar(p, (q) => q.v <= hi + EPS, hi);
    return p;
  }

  // La luz de una faceta según hacia dónde mira (medida en el contorno:
  // 0,937 + 0,76·nx + 0,14·nz (ajustado contra el corpus 26) con la normal del promedio de sus vértices;
  // [calibrar] la forma exacta: dos triángulos de una celda a veces tienen el
  // mismo brillo).
  const LUZ = JSON.parse('{"a":0.937,"bx":0.76,"bz":0.14,"lo":0.55,"hi":1.2}');
  const luminosidad = (nx, nz) => Math.max(LUZ.lo, Math.min(LUZ.hi, LUZ.a + LUZ.bx * nx + LUZ.bz * nz));

  // Normal suavizada de un vértice (diferencias centradas en el mundo).
  function normalDe(z, i, j, mundo) {
    const nc = z[j].length, ns = z.length;
    const alto = (a, b) => (z[b] && z[b][a] != null && Number.isFinite(z[b][a]) ? mundo.Y(z[b][a]) : null);
    const i0 = Math.max(0, i - 1), i1 = Math.min(nc - 1, i + 1), j0 = Math.max(0, j - 1), j1 = Math.min(ns - 1, j + 1);
    const ya = alto(i0, j), yb = alto(i1, j), yc = alto(i, j0), yd = alto(i, j1);
    const gx = ya != null && yb != null && i1 > i0 ? (yb - ya) / (mundo.X(i1) - mundo.X(i0)) : 0;
    const gz = yc != null && yd != null && j1 > j0 ? (yd - yc) / (mundo.Z(j1) - mundo.Z(j0)) : 0;
    const l = Math.hypot(gx, 1, gz);
    return [-gx / l, 1 / l, -gz / l];
  }

  // Las caras de la superficie: [{ pts: [[x, y, z]…], banda, luz }].
  function facetas(z, niveles, mundo) {
    const caras = [];
    const nb = niveles.length - 1;
    for (const tri of celdas(z)) {
      const vs = tri.map(([i, j]) => ({ i, j, v: z[j][i] }));
      const luz = tri.reduce((a, [i, j]) => { const n = normalDe(z, i, j, mundo); return a + luminosidad(n[0], n[2]) / 3; }, 0);
      const vmin = Math.min(...vs.map((q) => q.v)), vmax = Math.max(...vs.map((q) => q.v));
      for (let k = 0; k < nb; k++) {
        const lo = niveles[k], hi = niveles[k + 1];
        if (vmax < lo - EPS || vmin > hi + EPS) continue;
        const p = recortar(vs, lo, hi);
        if (p.length < 3) continue;
        const pts = p.map((q) => [mundo.X(q.i), mundo.Y(q.v), mundo.Z(q.j)]);
        // Sin área (un triángulo que solo toca un nivel): fuera.
        let area = 0; for (let m = 0; m < pts.length; m++) { const a = pts[m], b = pts[(m + 1) % pts.length]; area += a[0] * b[2] - b[0] * a[2]; }
        if (Math.abs(area) < 1e-9) continue;
        caras.push({ pts, banda: k, luz });
      }
    }
    return caras;
  }

  // El alambre: las aristas de la malla (horizontales y verticales, sin la
  // diagonal) partidas por banda y las líneas de nivel (el color de la banda
  // de abajo): [{ a: [x, y, z], b, banda }].
  function alambre(z, niveles, mundo) {
    const segs = [];
    const nb = niveles.length - 1;
    const banda = (v) => { for (let k = 0; k < nb; k++) if (v <= niveles[k + 1] + EPS) return v < niveles[k] - EPS ? -1 : k; return -1; };
    const P = (i, j, v) => [mundo.X(i), mundo.Y(v), mundo.Z(j)];
    const arista = (a, b) => {
      // a, b: [i, j]; partida por los niveles que cruza.
      const va = z[a[1]][a[0]], vb = z[b[1]][b[0]];
      if (va == null || vb == null || !Number.isFinite(va) || !Number.isFinite(vb)) return;
      const ts = [0, 1];
      for (const L of niveles) if ((va - L) * (vb - L) < 0) ts.push((L - va) / (vb - va));
      ts.sort((p, q) => p - q);
      for (let m = 0; m + 1 < ts.length; m++) {
        const t0 = ts[m], t1 = ts[m + 1], tm = (t0 + t1) / 2;
        const k = banda(va + (vb - va) * tm);
        if (k < 0) continue;
        const en = (t) => P(a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t, va + (vb - va) * t);
        segs.push({ a: en(t0), b: en(t1), banda: k });
      }
    };
    const ns = z.length, nc = Math.max(0, ...z.map((f) => f.length));
    for (let j = 0; j < ns; j++) for (let i = 0; i + 1 < nc; i++) arista([i, j], [i + 1, j]);
    for (let j = 0; j + 1 < ns; j++) for (let i = 0; i < nc; i++) arista([i, j], [i, j + 1]);
    // Líneas de nivel dentro de cada triángulo.
    for (const tri of celdas(z)) {
      const vs = tri.map(([i, j]) => ({ i, j, v: z[j][i] }));
      for (let k = 1; k < nb; k++) {
        const L = niveles[k], pts = [];
        for (let m = 0; m < 3; m++) {
          const a = vs[m], b = vs[(m + 1) % 3];
          if ((a.v - L) * (b.v - L) < 0) { const t = (L - a.v) / (b.v - a.v); pts.push(P(a.i + (b.i - a.i) * t, a.j + (b.j - a.j) * t, L)); }
        }
        if (pts.length === 2) segs.push({ a: pts[0], b: pts[1], banda: k - 1 });
      }
    }
    return segs;
  }

  return { facetas, alambre, luminosidad };
})();

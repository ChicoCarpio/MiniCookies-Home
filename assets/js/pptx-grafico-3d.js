/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Gráficos 3D de PowerPoint: la cámara (oblicua con ejes en ángulo recto, o
// en perspectiva) y los sólidos de las barras (caja, cilindro, cono, pirámide
// y sus variantes "hasta el máximo"), con sus caras ya proyectadas, sombreadas
// y en el orden en que se pintan. Lo usa pptx-grafico.js.
//
// Mundo: x a la derecha, y hacia arriba, z hacia el fondo. Una cámara es
// { p([x,y,z]) → [X,Y] de pantalla, ojo([x,y,z]) → vector hacia quien mira,
// lejania(eje, valor) → cuánto más lejos está ese valor de ese eje }.
//
// De dónde salen las reglas (investigacion-motor-powerpoint\06-graficos.md,
// fuera del repo; sondeo-3d.ps1 y sondeo-perspectiva.ps1, medidos por COM y
// en las imágenes de PowerPoint): ver cada función.

const PptxGrafico3D = (() => {
  const rad = (g) => g * Math.PI / 180;

  // Ejes en ángulo recto (rAngAx): proyección oblicua en la que el mundo es
  // la pantalla en el frente (en pt, y hacia arriba) y z corre z·sx a la
  // derecha y z·sy hacia arriba.
  function oblicua(sx, sy) {
    const v = [sx, sy, -1];
    return {
      p: ([x, y, z]) => [x + z * sx, -y - z * sy],
      ojo: () => v,
      lejania: (eje, valor) => -v[eje] * valor,
    };
  }

  // En perspectiva (rAngAx = 0), medido en 26 casos de sondeo-perspectiva
  // (las cajas de COM coinciden al centésimo con la cámara libre):
  //  - se gira la caja en su centro rotY alrededor de la vertical y después
  //    rotX hacia quien mira;
  //  - el ojo queda a L = 1,011 · altoProyectado / sen(perspective / 2) del
  //    centro (perspective en el XML es el doble del ángulo de la interfaz),
  //    donde altoProyectado es el alto de la caja sin perspectiva;
  //  - la imagen se escala para que la caja quepa en el 97,25 % del trazado y
  //    se centra en el centro del trazado + CORRE (en el sondeo, sin título
  //    ni leyenda, + (4; 3,93); en 25-graficos-3d, con los dos, ver CORRE).
  // [calibrar] con perspectiva fuerte (80 y más) y giros de 60° o más queda a
  // 2-4 pt: el ojo de PowerPoint no mira justo al centro de la caja. Y con la
  // elevación en 0, PowerPoint aleja algo más el ojo (×1,15).
  // caja: { W, H, D } en unidades del mundo (la caja va de 0 a W, 0 a H, 0 a D).
  // trazado: { x, y, w, h } en pt.
  const CORRE = { x: 0, y: 0 };
  // ajuste (opcional): { w, h, dx, dy }, el tamaño en que entra la caja y cuánto se corre
  // del centro del trazado (las superficies: el ancho +4,7 y el alto −5,06).
  function perspectiva(rotX, rotY, persp, caja, trazado, ajuste) {
    const { W, H, D } = caja;
    const a = rad(rotY), e = rad(-rotX);
    const ca = Math.cos(a), sa = Math.sin(a), ce = Math.cos(e), se = Math.sin(e);
    const altoProy = H * Math.cos(rad(rotX)) + (W * Math.abs(sa) + D * Math.abs(ca)) * Math.abs(Math.sin(rad(rotX)));
    // El ojo nunca adentro de la caja ni detrás de un vértice (con muchas
    // filas de fondo y rotX 0 o perspectiva fuerte la regla lo dejaba ahí y
    // la imagen salía espejada): al menos 1,5 × lo que la caja se acerca.
    let cerca = 0;
    for (const x of [-W / 2, W / 2]) for (const y of [-H / 2, H / 2]) for (const z of [-D / 2, D / 2]) cerca = Math.max(cerca, -(y * se + (-x * sa + z * ca) * ce));
    const L = persp > 0 ? Math.max(1.011 * altoProy / Math.sin(rad(Math.min(239, persp) / 2)), 1.5 * cerca) : 0;
    const prof = ([x, y, z]) => { x -= W / 2; y -= H / 2; z -= D / 2; const Z = -x * sa + z * ca; return y * se + Z * ce; };
    const vista = ([x, y, z]) => {
      x -= W / 2; y -= H / 2; z -= D / 2;
      const X = x * ca + z * sa; let Z = -x * sa + z * ca;
      const Y = y * ce - Z * se; Z = y * se + Z * ce;
      const k = L ? L / (L + Z) : 1;
      return [X * k, Y * k];
    };
    // El ojo en el mundo: el centro + el giro deshecho de (0, 0, −L).
    const ojoMundo = L ? [W / 2 + L * ce * sa, H / 2 - L * se, D / 2 - L * ce * ca] : null;
    const dirOjo = [Math.cos(rad(rotX)) * sa, Math.sin(rad(rotX)), -Math.cos(rad(rotX)) * ca];
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    for (const x of [0, W]) for (const y of [0, H]) for (const z of [0, D]) {
      const [X, Y] = vista([x, y, z]);
      x0 = Math.min(x0, X); x1 = Math.max(x1, X); y0 = Math.min(y0, -Y); y1 = Math.max(y1, -Y);
    }
    const tw = ajuste ? Math.max(1, ajuste.w) : 0.9725 * trazado.w, th = ajuste ? Math.max(1, ajuste.h) : 0.9725 * trazado.h;
    const s = Math.min(tw / Math.max(1e-9, x1 - x0), th / Math.max(1e-9, y1 - y0));
    const cx = trazado.x + trazado.w / 2 + (ajuste ? ajuste.dx : CORRE.x) - s * (x0 + x1) / 2, cy = trazado.y + trazado.h / 2 + (ajuste ? ajuste.dy : CORRE.y) - s * (y0 + y1) / 2;
    return {
      p: (q) => { const [X, Y] = vista(q); return [cx + s * X, cy - s * Y]; },
      ojo: (q) => (ojoMundo ? [ojoMundo[0] - q[0], ojoMundo[1] - q[1], ojoMundo[2] - q[2]] : dirOjo),
      lejania: (eje, valor) => (ojoMundo ? Math.abs(valor - ojoMundo[eje]) : -dirOjo[eje] * valor),
      escala: s,
      prof,
    };
  }

  // --- Sombreado (medido en 25-graficos-3d, rotX 15 y rotY 20) -------------------
  // Caras planas: el color × un factor según hacia dónde mira la cara; con la
  // normal en diagonal, la mezcla de los factores por el cuadrado de cada
  // componente. Medidos: frente 1, arriba 0,75, costado derecho 0,64; las
  // caras de la pirámide, × 1,05 de eso. [calibrar] izquierda, fondo y abajo.
  const LUZ = { frente: 1, fondo: 1, derecha: 0.64, izquierda: 0.64, arriba: 0.75, abajo: 0.5 };
  function luz([nx, ny, nz]) {
    const l = Math.hypot(nx, ny, nz) || 1; nx /= l; ny /= l; nz /= l;
    return nx * nx * (nx > 0 ? LUZ.derecha : LUZ.izquierda) + ny * ny * (ny > 0 ? LUZ.arriba : LUZ.abajo) + nz * nz * (nz < 0 ? LUZ.frente : LUZ.fondo);
  }
  // Superficies curvas: el factor a lo ancho de la silueta, de −1 a 1 (de
  // izquierda a derecha en las de pie; de arriba abajo en las acostadas),
  // medido píxel a píxel en el canal verde.
  const PERFIL = {
    cilindroV: [[-1, 0.45], [-0.9, 0.65], [-0.8, 0.76], [-0.6, 0.89], [-0.4, 0.965], [-0.2, 1.02], [0, 1.06], [0.1, 1.073], [0.2, 1.065], [0.4, 1.035], [0.6, 0.97], [0.8, 0.885], [0.9, 0.82], [1, 0.74]],
    cilindroH: [[-1, 0.86], [-0.8, 0.99], [-0.53, 1.07], [-0.27, 1.16], [-0.13, 1.17], [0, 1.1], [0.27, 0.98], [0.53, 0.85], [0.8, 0.625], [1, 0.4]],
    conoV: [[-1, 0.58], [-0.94, 0.625], [-0.72, 0.854], [-0.5, 0.97], [-0.28, 1.083], [-0.06, 1.2], [0.17, 1.156], [0.39, 1.05], [0.61, 0.94], [0.83, 0.8], [1, 0.64]],
    conoH: [[-1, 0.74], [-0.76, 0.97], [-0.56, 1.05], [-0.37, 1.125], [-0.17, 1.21], [0.02, 1.17], [0.22, 1.03], [0.41, 0.95], [0.61, 0.83], [0.8, 0.67], [1, 0.48]],
  };

  // --- Sólidos ---------------------------------------------------------------------
  // def: { forma: 'box'|'cylinder'|'cone'|'pyramid', eje: 'v' (columna) o 'h'
  // (barra), c0, c1: centros de la base y del tope en el mundo, r0, r1: medio
  // ancho a lo ancho (x en las de pie, y en las acostadas), q0, q1: medio
  // ancho en z }. Devuelve { caras: [{ puntos, luz } o { puntos, perfil, angulo }],
  // caja: { x, y, w, h } en pantalla }.
  const LADOS = 48;
  function solido(def, cam) {
    const { forma, eje, c0, c1, r0, r1, q0, q1 } = def;
    const ancho = eje === 'v' ? [1, 0, 0] : [0, 1, 0], hondo = [0, 0, 1];
    const punto = (c, r, q, u, v) => [c[0] + ancho[0] * r * u + hondo[0] * q * v, c[1] + ancho[1] * r * u + hondo[1] * q * v, c[2] + ancho[2] * r * u + hondo[2] * q * v];
    // Hacia dónde va el eje del sólido (una placa sin grosor no tiene largo).
    const k = eje === 'v' ? 1 : 0, sentido = Math.sign(c1[k] - c0[k]) || 1;
    const largo = eje === 'v' ? [0, sentido, 0] : [sentido, 0, 0];
    const caras = [];
    const visible = (pts3, n) => { const c = medio(pts3); const o = cam.ojo(c); return n[0] * o[0] + n[1] * o[1] + n[2] * o[2] > 1e-9; };
    // La normal de una cara, hacia afuera: la de sus vértices, dada vuelta
    // si apunta al revés del lado que es (de lado: abajo, arriba, frente…).
    const hacia = (pts3, lado) => {
      const n = newell(pts3);
      if (Math.hypot(n[0], n[1], n[2]) < 1e-12) return lado;
      return n[0] * lado[0] + n[1] * lado[1] + n[2] * lado[2] < 0 ? [-n[0], -n[1], -n[2]] : n;
    };
    const pico = r1 < 1e-9 && q1 < 1e-9;
    if (forma === 'cylinder' || forma === 'cone') {
      const anillo = (c, r, q) => Array.from({ length: LADOS }, (_, i) => { const t = 2 * Math.PI * i / LADOS; return punto(c, r, q, Math.cos(t), Math.sin(t)); });
      const abajo = anillo(c0, r0, q0), arriba = pico ? [c1] : anillo(c1, r1, q1);
      const lado = envolvente([...abajo, ...arriba].map(cam.p));
      const perfil = PERFIL[(forma === 'cone' ? 'cono' : 'cilindro') + (eje === 'v' ? 'V' : 'H')];
      caras.push({ puntos: lado, perfil, angulo: eje === 'v' ? 0 : 90 });
      // Las tapas, si se ven (de pie: la de arriba si se mira desde arriba).
      for (const [anilloT, signo] of [[abajo, -1], [pico ? null : arriba, 1]]) {
        if (!anilloT) continue;
        const n = largo.map((v) => v * signo);
        if (visible(anilloT, n)) caras.push({ puntos: anilloT.map(cam.p), luz: luz(n) });
      }
    } else {
      // Esquinas en el orden (−,−) (+,−) (+,+) (−,+): el costado i va del
      // frente (i = 0) a la derecha, el fondo y la izquierda.
      const esq = (c, r, q) => [punto(c, r, q, -1, -1), punto(c, r, q, 1, -1), punto(c, r, q, 1, 1), punto(c, r, q, -1, 1)];
      const b = esq(c0, r0, q0), t = pico ? null : esq(c1, r1, q1);
      const lados = [[0, 0, -1], ancho, [0, 0, 1], ancho.map((v) => -v)];
      const lista = [[b, largo.map((v) => -v)]];
      if (t) lista.push([t, largo]);
      for (let i = 0; i < 4; i++) {
        const j = (i + 1) % 4;
        lista.push([t ? [b[i], b[j], t[j], t[i]] : [b[i], b[j], c1], lados[i]]);
      }
      const extra = forma === 'pyramid' ? 1.05 : 1;
      for (const [pts3, lado] of lista) {
        const n = hacia(pts3, lado);
        if (visible(pts3, n)) caras.push({ puntos: pts3.map(cam.p), luz: Math.min(1.3, luz(n) * extra) });
      }
    }
    // La caja, como la da COM: la del prisma de base cuadrada que lo contiene.
    const cuadrado = (c, r, q) => [punto(c, r, q, -1, -1), punto(c, r, q, 1, -1), punto(c, r, q, 1, 1), punto(c, r, q, -1, 1)];
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    for (const v of [...cuadrado(c0, r0, q0), ...(pico ? [c1] : cuadrado(c1, r1, q1))]) {
      const [x, y] = cam.p(v); x0 = Math.min(x0, x); x1 = Math.max(x1, x); y0 = Math.min(y0, y); y1 = Math.max(y1, y);
    }
    return { caras, caja: { x: x0, y: y0, w: x1 - x0, h: y1 - y0 } };
  }

  // Orden de pintado de las barras de una grilla: por filas de profundidad,
  // después a lo largo de las categorías y por último en la pila, siempre de
  // lo más lejano a lo más cercano (cada barra tapa el costado de la de atrás).
  // celda: { z, a, s } (centros en el mundo); ejes: { a, s } (0 = x, 1 = y).
  function ordenar(lista, cam, ejes) {
    const clave = (c) => [cam.lejania(2, c.z), cam.lejania(ejes.a, c.a), cam.lejania(ejes.s, c.s)];
    const conClave = lista.map((x) => ({ x, k: clave(x.celda) }));
    conClave.sort((p, q) => {
      for (let i = 0; i < 3; i++) if (Math.abs(p.k[i] - q.k[i]) > 1e-6) return q.k[i] - p.k[i];
      return 0;
    });
    return conClave.map((c) => c.x);
  }

  // Un relleno de degradado para el costado curvo: el color del relleno liso
  // por el perfil medido. Sin color liso, nada (se pinta tal cual).
  function degradadoDe(rell, perfil, angulo) {
    if (!rell || rell.tipo !== 'solido') return null;
    const c = rell.color, k = (f) => Math.max(0, Math.min(255, Math.round(f)));
    return {
      tipo: 'degradado', angulo, escalado: false, rotarConForma: false,
      paradas: perfil.map(([u, f]) => ({ pos: (u + 1) / 2, color: { r: k(c.r * f), g: k(c.g * f), b: k(c.b * f), a: c.a } })),
    };
  }

  // --- Geometría ---------------------------------------------------------------------
  function medio(pts) { const n = pts.length; return pts.reduce((a, p) => [a[0] + p[0] / n, a[1] + p[1] / n, a[2] + p[2] / n], [0, 0, 0]); }
  function newell(pts) {
    let x = 0, y = 0, z = 0;
    for (let i = 0; i < pts.length; i++) {
      const a = pts[i], b = pts[(i + 1) % pts.length];
      x += (a[1] - b[1]) * (a[2] + b[2]); y += (a[2] - b[2]) * (a[0] + b[0]); z += (a[0] - b[0]) * (a[1] + b[1]);
    }
    return [x, y, z];
  }
  // Envolvente convexa (cadena monótona), en sentido horario de pantalla.
  function envolvente(pts) {
    const p = pts.slice().sort((a, b) => a[0] - b[0] || a[1] - b[1]);
    if (p.length < 3) return p;
    const cruz = (o, a, b) => (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0]);
    const abajo = [], arriba = [];
    for (const q of p) { while (abajo.length >= 2 && cruz(abajo[abajo.length - 2], abajo[abajo.length - 1], q) <= 0) abajo.pop(); abajo.push(q); }
    for (let i = p.length - 1; i >= 0; i--) { const q = p[i]; while (arriba.length >= 2 && cruz(arriba[arriba.length - 2], arriba[arriba.length - 1], q) <= 0) arriba.pop(); arriba.push(q); }
    abajo.pop(); arriba.pop();
    return abajo.concat(arriba);
  }

  return { oblicua, perspectiva, solido, ordenar, degradadoDe, luz, envolvente };
})();

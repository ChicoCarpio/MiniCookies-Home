/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Color percibido (OKLab) y reducción de una imagen a pocos colores, para los
// efectos de assets/js/efectos-mosaico.js (mosaicos y cómic).
//
// Por qué OKLab y no RGB: reducir colores en RGB, canal por canal, inventa
// tonos que la foto no tiene (los parches rojos y verdes del primer intento
// de "dibujo animado"). En OKLab la distancia entre dos colores se parece a
// cuánto se diferencian a la vista, así que los promedios y la elección del
// color más cercano dan el tono real.
window.PaletaColor = (function () {
  // sRGB (0-255) -> luz lineal (0-1), precalculado
  const A_LINEAL = new Float32Array(256);
  for (let i = 0; i < 256; i++) {
    const c = i / 255;
    A_LINEAL[i] = c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  }
  function aSrgb(c) {
    c = c <= 0.0031308 ? 12.92 * c : 1.055 * c ** (1 / 2.4) - 0.055;
    return Math.max(0, Math.min(255, Math.round(c * 255)));
  }

  // Luz lineal -> OKLab [L, a, b] (fórmulas de Björn Ottosson)
  function linealAOklab(r, g, b) {
    const l = Math.cbrt(0.4122214708 * r + 0.5363325363 * g + 0.0514459929 * b);
    const m = Math.cbrt(0.2119034982 * r + 0.6806995451 * g + 0.1073969566 * b);
    const s = Math.cbrt(0.0883024619 * r + 0.2817188376 * g + 0.6299787005 * b);
    return [
      0.2104542553 * l + 0.793617785 * m - 0.0040720468 * s,
      1.9779984951 * l - 2.428592205 * m + 0.4505937099 * s,
      0.0259040371 * l + 0.7827717662 * m - 0.808675766 * s,
    ];
  }
  function oklabARgb(L, a, b) {
    const l = (L + 0.3963377774 * a + 0.2158037573 * b) ** 3;
    const m = (L - 0.1055613458 * a - 0.0638541728 * b) ** 3;
    const s = (L - 0.0894841775 * a - 1.291485548 * b) ** 3;
    return [
      aSrgb(4.0767416621 * l - 3.3077115913 * m + 0.2309699292 * s),
      aSrgb(-1.2684380046 * l + 2.6097574011 * m - 0.3413193965 * s),
      aSrgb(-0.0041960863 * l - 0.7034186147 * m + 1.707614701 * s),
    ];
  }

  // El mismo color más claro u oscuro (solo cambia la luminosidad), como CSS
  function tono(lab, factorL, sumaL = 0) {
    const [r, g, b] = oklabARgb(Math.max(0, Math.min(1, lab[0] * factorL + sumaL)), lab[1], lab[2]);
    return `rgb(${r},${g},${b})`;
  }

  const dist2 = (p, q) => (p[0] - q[0]) ** 2 + (p[1] - q[1]) ** 2 + (p[2] - q[2]) ** 2;

  function masCercano(p, centros) {
    let mejor = 0, mejorD = Infinity;
    for (let i = 0; i < centros.length; i++) {
      const dd = dist2(p, centros[i]);
      if (dd < mejorD) { mejorD = dd; mejor = i; }
    }
    return mejor;
  }

  // mulberry32: azar con semilla, para que la misma foto dé siempre la misma
  // paleta (si no, cada redibujado al mover un control cambiaría los colores)
  function azar(semilla) {
    return () => {
      semilla |= 0; semilla = (semilla + 0x6d2b79f5) | 0;
      let t = Math.imul(semilla ^ (semilla >>> 15), 1 | semilla);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  // Los colores parecidos se juntan primero (cubos de 0.03 en OKLab) y cada
  // grupo pesa la RAÍZ de su tamaño, no su tamaño. Si no, el fondo grande se
  // lleva todos los colores y lo chico (ojos, una lengua, brillos) desaparece.
  function agrupar(puntos) {
    const grupos = new Map();
    for (const p of puntos) {
      const clave = Math.round(p[0] / 0.03) + ',' + Math.round(p[1] / 0.03) + ',' + Math.round(p[2] / 0.03);
      const g = grupos.get(clave);
      if (g) { g.s[0] += p[0]; g.s[1] += p[1]; g.s[2] += p[2]; g.n++; } else grupos.set(clave, { s: p.slice(), n: 1 });
    }
    const centros = [], pesos = [];
    for (const g of grupos.values()) {
      centros.push([g.s[0] / g.n, g.s[1] / g.n, g.s[2] / g.n]);
      pesos.push(Math.sqrt(g.n));
    }
    return { centros, pesos };
  }

  // Paleta de "k" colores para "puntos" (colores OKLab): k-means con arranque
  // k-means++ (cada centro nuevo, lejos de los anteriores).
  function paleta(puntos, k) {
    if (!puntos.length) return [[0, 0, 0]];
    const { centros: unicos, pesos } = agrupar(puntos);
    const rnd = azar(12345);
    k = Math.max(1, Math.min(k, unicos.length));
    let mayor = 0;
    for (let i = 1; i < pesos.length; i++) if (pesos[i] > pesos[mayor]) mayor = i;
    const centros = [unicos[mayor].slice()];
    const minD = unicos.map((p) => dist2(p, centros[0]));
    while (centros.length < k) {
      let total = 0;
      for (let j = 0; j < unicos.length; j++) total += minD[j] * pesos[j];
      let objetivo = rnd() * total, i = 0;
      while (i < unicos.length - 1 && (objetivo -= minD[i] * pesos[i]) > 0) i++;
      const nuevo = unicos[i].slice();
      centros.push(nuevo);
      for (let j = 0; j < unicos.length; j++) minD[j] = Math.min(minD[j], dist2(unicos[j], nuevo));
    }
    const asignado = new Int32Array(unicos.length).fill(-1);
    for (let vuelta = 0; vuelta < 20; vuelta++) {
      let cambios = 0;
      for (let j = 0; j < unicos.length; j++) {
        const m = masCercano(unicos[j], centros);
        if (m !== asignado[j]) { asignado[j] = m; cambios++; }
      }
      if (!cambios) break;
      const suma = centros.map(() => [0, 0, 0, 0]);
      for (let j = 0; j < unicos.length; j++) {
        const s = suma[asignado[j]], p = unicos[j], w = pesos[j];
        s[0] += p[0] * w; s[1] += p[1] * w; s[2] += p[2] * w; s[3] += w;
      }
      suma.forEach((s, i) => { if (s[3]) centros[i] = [s[0] / s[3], s[1] / s[3], s[2] / s[3]]; });
    }
    return centros;
  }

  return { A_LINEAL, linealAOklab, oklabARgb, tono, dist2, masCercano, paleta };
})();

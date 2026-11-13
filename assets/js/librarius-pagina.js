/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
/*
 * Portado de ia/versus/leer_pagina.py (que es la referencia: este archivo tiene que dar lo mismo). Sin dependencias: recibe las dos sesiones de
 * onnxruntime-web (Versus, el detector de líneas, y Librarius, el lector) ya creadas.
 *
 *   const lector = LibrariusPagina.crear({ ort, versus, librarius, alfabeto });
 *   const r = await lector.leer(gris, ancho, alto);     // gris: Uint8Array de ancho*alto (0 negro, 255 blanco)
 *   r.texto, r.lineas[{ texto, seguridad, centro, largo, alto, u }], r.giro (grados), r.tiempos (ms por etapa)
 *
 * Etapas: 1) reducir la página a 1.024 px de lado largo y correr Versus; 2) componentes conexos, eje (PCA) y largo de cada línea, unir trozos;
 * 3) recortar cada línea enderezada y leerla; 4) orientación: votan líneas leídas al derecho y al revés; si la página estaba girada 90/180/270°,
 * se endereza y se detecta de nuevo; 5) orden de lectura por corte recursivo X-Y.
 */
(function (global) {
  'use strict';

  // ------------------------------------------------------------------ utilidades numéricas
  function redondearPar(x) {                       // como round() de Python (mitades al par)
    const f = Math.floor(x), d = x - f;
    if (d > 0.5) return f + 1;
    if (d < 0.5) return f;
    return f % 2 === 0 ? f : f + 1;
  }

  function mediana(valores) {
    const a = Array.from(valores).sort((x, y) => x - y);
    const n = a.length;
    return n % 2 ? a[(n - 1) / 2] : (a[n / 2 - 1] + a[n / 2]) / 2;
  }

  function percentil(gris, p) {                    // como np.percentile (interpolación lineal), sobre bytes
    const hist = new Uint32Array(256);
    for (let i = 0; i < gris.length; i++) hist[gris[i]]++;
    const pos = (p / 100) * (gris.length - 1), lo = Math.floor(pos), hi = Math.ceil(pos);
    const valorK = (k) => { let acum = 0; for (let v = 0; v < 256; v++) { acum += hist[v]; if (k < acum) return v; } return 255; };
    return valorK(lo) + (valorK(hi) - valorK(lo)) * (pos - lo);
  }

  // ------------------------------------------------------------------ cambio de tamaño (área al reducir; bilineal o bicúbico al ampliar)
  function tablaEje(s, d, cubico) {
    const escala = s / d, encoge = escala > 1, maxN = encoge ? Math.ceil(escala) + 2 : 4;
    const idx = new Int32Array(d * maxN), pes = new Float32Array(d * maxN), cnt = new Int32Array(d);
    for (let i = 0; i < d; i++) {
      let n = 0;
      if (encoge) {
        const a = i * escala, b = (i + 1) * escala, k0 = Math.floor(a), k1 = Math.min(s - 1, Math.ceil(b) - 1);
        let total = 0;
        for (let k = k0; k <= k1; k++) {
          const w = Math.min(b, k + 1) - Math.max(a, k);
          if (w > 1e-9) { idx[i * maxN + n] = k; pes[i * maxN + n] = w; total += w; n++; }
        }
        for (let j = 0; j < n; j++) pes[i * maxN + j] /= total;
      } else {
        const x = (i + 0.5) * escala - 0.5, x0 = Math.floor(x), f = x - x0;
        const pesos = cubico
          ? (() => {
              const A = -0.75;
              const w0 = ((A * (f + 1) - 5 * A) * (f + 1) + 8 * A) * (f + 1) - 4 * A;
              const w1 = ((A + 2) * f - (A + 3)) * f * f + 1;
              const w2 = ((A + 2) * (1 - f) - (A + 3)) * (1 - f) * (1 - f) + 1;
              return [w0, w1, w2, 1 - w0 - w1 - w2];
            })()
          : [0, 1 - f, f, 0];
        for (let j = 0; j < 4; j++) {
          if (!cubico && (j === 0 || j === 3)) continue;
          idx[i * maxN + n] = Math.min(s - 1, Math.max(0, x0 - 1 + j)); pes[i * maxN + n] = pesos[j]; n++;
        }
      }
      cnt[i] = n;
    }
    return { idx, pes, cnt, maxN };
  }

  function redimensionar(src, sw, sh, dw, dh, cubico) {
    const tx = tablaEje(sw, dw, cubico), ty = tablaEje(sh, dh, cubico);
    const medio = new Float32Array(sh * dw);
    for (let y = 0; y < sh; y++) {
      const fila = y * sw;
      for (let x = 0; x < dw; x++) {
        let acum = 0; const base = x * tx.maxN;
        for (let j = 0; j < tx.cnt[x]; j++) acum += src[fila + tx.idx[base + j]] * tx.pes[base + j];
        medio[y * dw + x] = acum;
      }
    }
    const salida = new Uint8Array(dw * dh);
    for (let y = 0; y < dh; y++) {
      const base = y * ty.maxN;
      for (let x = 0; x < dw; x++) {
        let acum = 0;
        for (let j = 0; j < ty.cnt[y]; j++) acum += medio[ty.idx[base + j] * dw + x] * ty.pes[base + j];
        salida[y * dw + x] = Math.max(0, Math.min(255, redondearPar(acum)));
      }
    }
    return salida;
  }

  function girarPagina(gris, w, h, k) {            // k = 1: 90° antihorario, 2: 180°, 3: 90° horario (como cv2.rotate); devuelve {gris, w, h}
    if (k === 2) { const s = new Uint8Array(gris.length); for (let i = 0; i < gris.length; i++) s[i] = gris[gris.length - 1 - i]; return { gris: s, w, h }; }
    const s = new Uint8Array(gris.length), nw = h, nh = w;
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const v = gris[y * w + x];
        if (k === 1) s[(w - 1 - x) * nw + y] = v;               // antihorario: (x, y) -> (y, w-1-x)
        else s[x * nw + (h - 1 - y)] = v;                        // horario: (x, y) -> (h-1-y, x)
      }
    }
    return { gris: s, w: nw, h: nh };
  }

  // ------------------------------------------------------------------ componentes conexos (8 vecinos) y forma de cada línea
  function componentes(mascara, w, h, areaMinima) {
    const etiqueta = new Uint8Array(w * h), lista = [], pila = new Int32Array(w * h);
    for (let inicio = 0; inicio < w * h; inicio++) {
      if (!mascara[inicio] || etiqueta[inicio]) continue;
      let tope = 0; pila[tope++] = inicio; etiqueta[inicio] = 1;
      const xs = [], ys = [];
      while (tope) {
        const p = pila[--tope], x = p % w, y = (p - x) / w;
        xs.push(x); ys.push(y);
        for (let dy = -1; dy <= 1; dy++) {
          const yy = y + dy; if (yy < 0 || yy >= h) continue;
          for (let dx = -1; dx <= 1; dx++) {
            const xx = x + dx; if (xx < 0 || xx >= w) continue;
            const q = yy * w + xx;
            if (mascara[q] && !etiqueta[q]) { etiqueta[q] = 1; pila[tope++] = q; }
          }
        }
      }
      if (xs.length >= areaMinima) lista.push({ xs, ys });
    }
    return lista;
  }

  function forma(px, py, f) {                      // eje (PCA), centro, largo y alto de una nube de puntos
    const n = px.length;
    let mx = 0, my = 0;
    for (let i = 0; i < n; i++) { mx += px[i]; my += py[i]; }
    mx /= n; my /= n;
    let sxx = 0, sxy = 0, syy = 0;
    for (let i = 0; i < n; i++) { const a = px[i] - mx, b = py[i] - my; sxx += a * a; sxy += a * b; syy += b * b; }
    const theta = 0.5 * Math.atan2(2 * sxy, sxx - syy);
    let ux = Math.cos(theta), uy = Math.sin(theta);
    const vx = -uy, vy = ux;
    let tmin = Infinity, tmax = -Infinity, smin = Infinity, smax = -Infinity;
    for (let i = 0; i < n; i++) {
      const a = px[i] - mx, b = py[i] - my, t = a * ux + b * uy, s = a * vx + b * vy;
      if (t < tmin) tmin = t; if (t > tmax) tmax = t; if (s < smin) smin = s; if (s > smax) smax = s;
    }
    const cx = mx + ux * (tmax + tmin) / 2 + vx * (smax + smin) / 2, cy = my + uy * (tmax + tmin) / 2 + vy * (smax + smin) / 2;
    if (ux < 0 || (Math.abs(ux) < 1e-9 && uy < 0)) { ux = -ux; uy = -uy; }
    return { centro: [cx, cy], u: [ux, uy], largo: tmax - tmin + f, alto: smax - smin + f };
  }

  // ------------------------------------------------------------------ orden de lectura (corte recursivo X-Y)
  function ordenarLineas(lineas, giro) {
    const c = Math.cos(-giro), s = Math.sin(-giro);
    const centros = lineas.map((l) => [c * l.centro[0] - s * l.centro[1], s * l.centro[0] + c * l.centro[1]]);
    const cajas = lineas.map((l, i) => [centros[i][0] - l.largo / 2, centros[i][1] - l.alto / 2, centros[i][0] + l.largo / 2, centros[i][1] + l.alto / 2]);
    const hMed = mediana(lineas.map((l) => l.alto));

    function hueco(ini, fin, minimo) {              // posición de un hueco >= minimo que ninguna línea cruza, o null
      const o = ini.map((_, i) => i).sort((a, b) => ini[a] - ini[b]);
      let borde = fin[o[0]], mejor = null;
      for (let j = 1; j < o.length; j++) {
        const k = o[j], g = ini[k] - borde;
        if (g >= minimo && (mejor === null || g > mejor[0])) mejor = [g, (ini[k] + borde) / 2];
        borde = Math.max(borde, fin[k]);
      }
      return mejor === null ? null : mejor[1];
    }

    function recursivo(ids) {
      if (ids.length <= 1) return ids.slice();
      let corte = hueco(ids.map((i) => cajas[i][0]), ids.map((i) => cajas[i][2]), 0.6 * hMed);          // columnas
      if (corte !== null) {
        const izq = ids.filter((i) => (cajas[i][0] + cajas[i][2]) / 2 < corte), der = ids.filter((i) => !izq.includes(i));
        if (izq.length && der.length) return recursivo(izq).concat(recursivo(der));
      }
      corte = hueco(ids.map((i) => cajas[i][1]), ids.map((i) => cajas[i][3]), 2.0 * hMed);                // bloques
      if (corte !== null) {
        const arriba = ids.filter((i) => (cajas[i][1] + cajas[i][3]) / 2 < corte), abajo = ids.filter((i) => !arriba.includes(i));
        if (arriba.length && abajo.length) return recursivo(arriba).concat(recursivo(abajo));
      }
      const porY = ids.slice().sort((a, b) => centros[a][1] - centros[b][1]);                            // filas
      const filas = []; let actual = [porY[0]];
      for (let j = 1; j < porY.length; j++) {
        const i = porY[j], media = actual.reduce((a, k) => a + centros[k][1], 0) / actual.length;
        if (Math.abs(centros[i][1] - media) <= 0.6 * hMed) actual.push(i); else { filas.push(actual); actual = [i]; }
      }
      filas.push(actual);
      return filas.flatMap((fila) => fila.slice().sort((a, b) => centros[a][0] - centros[b][0]));
    }
    return recursivo(lineas.map((_, i) => i));
  }

  // ------------------------------------------------------------------ piezas de la tubería (las usa `crear` y también los trabajadores)
  // `tiempos` es opcional: un objeto donde se suman los milisegundos de cada etapa.
  const marca = (tiempos, nombre, t0) => { if (tiempos) tiempos[nombre] = (tiempos[nombre] || 0) + performance.now() - t0; };

  // ctx: { ort, versus, ladoDet, umbralLogit, unir }  ->  líneas [{ centro, u, largo, alto }]
  async function detectarLineas(ctx, gris, w, h, tiempos) {
    const { ort, versus, ladoDet, umbralLogit, unir } = ctx;
    let t0 = performance.now();
    const e = ladoDet / Math.max(h, w);
    const pw = Math.max(32, redondearPar(w * e)), ph = Math.max(32, redondearPar(h * e));
    const pequena = redimensionar(gris, w, h, pw, ph, false);
    const aw = pw + ((32 - (pw % 32)) % 32), ah = ph + ((32 - (ph % 32)) % 32);
    const entrada = new Float32Array(aw * ah).fill(1);
    for (let y = 0; y < ph; y++) for (let x = 0; x < pw; x++) entrada[y * aw + x] = pequena[y * pw + x] / 255;
    marca(tiempos, 'preparar_versus', t0);
    t0 = performance.now();
    const salida = await versus.run({ pagina: new ort.Tensor('float32', entrada, [1, 1, ah, aw]) });
    marca(tiempos, 'versus', t0);
    t0 = performance.now();
    const mapa = salida.mapa.data, mw = aw / 2, mh = (ph + 1) >> 1, mcw = (pw + 1) >> 1;
    const mascara = new Uint8Array(mcw * mh);
    for (let y = 0; y < mh; y++) for (let x = 0; x < mcw; x++) mascara[y * mcw + x] = mapa[y * mw + x] >= umbralLogit ? 1 : 0;
    const f = 2.0 / e;
    const nubes = componentes(mascara, mcw, mh, 12).map((c) => ({ px: c.xs.map((x) => x * f + f / 2), py: c.ys.map((y) => y * f + f / 2) }));
    const formas = nubes.map((n) => forma(n.px, n.py, f));
    // unir trozos de una misma línea (mismo eje, mismo renglón y hueco pequeño)
    const padre = formas.map((_, i) => i);
    const raiz = (i) => { while (padre[i] !== i) { padre[i] = padre[padre[i]]; i = padre[i]; } return i; };
    for (let i = 0; i < formas.length; i++) {
      for (let j = i + 1; j < formas.length; j++) {
        const a = formas[i], b = formas[j];
        if (Math.abs(a.u[0] * b.u[0] + a.u[1] * b.u[1]) < 0.985) continue;
        const dx = b.centro[0] - a.centro[0], dy = b.centro[1] - a.centro[1];
        const longitudinal = Math.abs(dx * a.u[0] + dy * a.u[1]), lateral = Math.abs(dx * -a.u[1] + dy * a.u[0]);
        const caja = a.alto + b.alto, hueco = longitudinal - (a.largo + b.largo) / 2;
        if (lateral <= 0.3 * caja && hueco <= unir * caja) padre[raiz(j)] = raiz(i);
      }
    }
    const grupos = new Map();
    nubes.forEach((n, i) => { const r = raiz(i); if (!grupos.has(r)) grupos.set(r, []); grupos.get(r).push(n); });
    const lineas = [];
    for (const lista of grupos.values()) {
      const fm = lista.length > 1 ? forma([].concat(...lista.map((n) => n.px)), [].concat(...lista.map((n) => n.py)), f) : forma(lista[0].px, lista[0].py, f);
      if (fm.largo < 2.0 * fm.alto || fm.alto < 3) continue;
      lineas.push(fm);
    }
    marca(tiempos, 'componentes', t0);
    return lineas;
  }

  // Sentido común del eje: en líneas verticales el signo de cada eje sale al azar; se alinea con el eje dominante de la página (o con `eje` si se da).
  function alinearEje(lineas, eje) {
    let txx = 0, txy = 0, tyy = 0;
    for (const l of lineas) { txx += l.largo * l.u[0] * l.u[0]; txy += l.largo * l.u[0] * l.u[1]; tyy += l.largo * l.u[1] * l.u[1]; }
    let dominante = eje;
    if (!dominante) { const th = 0.5 * Math.atan2(2 * txy, txx - tyy); dominante = [Math.cos(th), Math.sin(th)]; }
    for (const l of lineas) if (l.u[0] * dominante[0] + l.u[1] * dominante[1] < 0) l.u = [-l.u[0], -l.u[1]];
    return lineas;
  }

  // Líneas que votan la orientación: las más largas y otras repartidas por la página. Devuelve índices de `lineas`.
  function muestraOrientacion(lineas, k) {
    const orden = lineas.map((_, i) => i).sort((a, b) => lineas[b].largo - lineas[a].largo).slice(0, 4 * k);
    const mitad = k >> 1, paso = Math.max(1, Math.floor((orden.length - mitad) / (mitad + 1)));
    const otras = []; for (let i = mitad; i < orden.length; i += paso) otras.push(orden[i]);
    return orden.slice(0, mitad).concat(otras.slice(0, mitad));
  }

  // ctx: { alto, alturaRel, rellenoRel }  ->  { px: Uint8Array(alto * ancho), ancho }
  function recortarLinea(ctx, gris, w, h, linea, invertir, relleno) {
    const { alto, alturaRel, rellenoRel } = ctx;
    let ux = linea.u[0], uy = linea.u[1], vx = -uy, vy = ux;
    if (invertir) { ux = -ux; uy = -uy; vx = -vx; vy = -vy; }
    const altoReal = Math.max(linea.alto * alturaRel, 8), ancho = linea.largo + 2 * rellenoRel * altoReal;
    const W = Math.max(16, redondearPar(ancho)), H = Math.max(8, redondearPar(altoReal));
    const c0 = linea.centro[0] - ux * W / 2 - vx * H / 2, c1 = linea.centro[1] - uy * W / 2 - vy * H / 2;
    const a = new Uint8Array(W * H), wm = w - 1, hm = h - 1;
    for (let Y = 0; Y < H; Y++) {
      let x = vx * Y + c0, y = vy * Y + c1;                                  // la columna X = 0 de esta fila; se avanza sumando u
      for (let X = 0; X < W; X++, x += ux, y += uy) {
        const x0 = Math.floor(x), y0 = Math.floor(y), fx = x - x0, fy = y - y0;
        let p00, p10, p01, p11;
        if (x0 >= 0 && y0 >= 0 && x0 < wm && y0 < hm) {                      // caso común: los 4 vecinos están dentro de la página
          const k = y0 * w + x0; p00 = gris[k]; p10 = gris[k + 1]; p01 = gris[k + w]; p11 = gris[k + w + 1];
        } else {                                                             // borde: lo de afuera vale `relleno`
          p00 = (x0 < 0 || y0 < 0 || x0 > wm || y0 > hm) ? relleno : gris[y0 * w + x0];
          p10 = (x0 + 1 < 0 || y0 < 0 || x0 + 1 > wm || y0 > hm) ? relleno : gris[y0 * w + x0 + 1];
          p01 = (x0 < 0 || y0 + 1 < 0 || x0 > wm || y0 + 1 > hm) ? relleno : gris[(y0 + 1) * w + x0];
          p11 = (x0 + 1 < 0 || y0 + 1 < 0 || x0 + 1 > wm || y0 + 1 > hm) ? relleno : gris[(y0 + 1) * w + x0 + 1];
        }
        a[Y * W + X] = Math.floor(p00 * (1 - fx) * (1 - fy) + p10 * fx * (1 - fy) + p01 * (1 - fx) * fy + p11 * fx * fy + 0.5);
      }
    }
    const nw = Math.max(8, redondearPar(W * alto / H));
    return { px: redimensionar(a, W, H, nw, alto, H <= alto), ancho: nw };
  }

  // ctx: { ort, librarius, alfabeto, alto }; r = { px, ancho }  ->  { texto, seguridad }
  async function leerRecorte(ctx, r, tiempos) {
    const { ort, librarius, alfabeto, alto } = ctx;
    const t0 = performance.now();
    const ancho4 = r.ancho + ((4 - (r.ancho % 4)) % 4);
    const borde = [];
    for (let y = 0; y < alto; y++) { borde.push(r.px[y * r.ancho], r.px[y * r.ancho + 1], r.px[y * r.ancho + r.ancho - 2], r.px[y * r.ancho + r.ancho - 1]); }
    const relleno = Math.trunc(mediana(borde));
    const x = new Float32Array(alto * ancho4).fill(relleno / 255);
    for (let y = 0; y < alto; y++) for (let c = 0; c < r.ancho; c++) x[y * ancho4 + c] = r.px[y * r.ancho + c] / 255;
    const t1 = performance.now();
    const salida = (await librarius.run({ imagen: new ort.Tensor('float32', x, [1, 1, alto, ancho4]) })).logits;
    const t2 = performance.now();
    const T = salida.dims[1], C = salida.dims[2], datos = salida.data, tValido = Math.min(T, (r.ancho >> 2) + 1);
    let texto = '', previo = 0, suma = 0, nNoBlanco = 0;
    for (let t = 0; t < tValido; t++) {
      let mejor = 0, valor = -Infinity;
      for (let k = 0; k < C; k++) { const v = datos[t * C + k]; if (v > valor) { valor = v; mejor = k; } }
      if (mejor !== 0) {
        let denom = 0; for (let k = 0; k < C; k++) denom += Math.exp(datos[t * C + k] - valor);
        suma += 1 / denom; nNoBlanco++;
      }
      if (mejor !== previo && mejor !== 0) texto += alfabeto[mejor - 1];
      previo = mejor;
    }
    if (tiempos) {
      tiempos.preparar_lector = (tiempos.preparar_lector || 0) + (t1 - t0);
      tiempos.lector = (tiempos.lector || 0) + (t2 - t1);
      tiempos.decodificar = (tiempos.decodificar || 0) + (performance.now() - t2);
    }
    return { texto, seguridad: nNoBlanco ? suma / nNoBlanco : 0 };
  }

  const DEFECTOS = { ladoDet: 1024, alto: 24, alturaRel: 1.8, rellenoRel: 0.35, unir: 0.3, kOrientacion: 8, umbral: 0.5 };   // kOrientacion: 4, 8, 12 y 24 dieron lo mismo en 570 páginas

  function contextos(opciones) {
    const o = Object.assign({}, DEFECTOS, opciones);
    return {
      det: { ort: o.ort, versus: o.versus, ladoDet: o.ladoDet, umbralLogit: Math.log(o.umbral / (1 - o.umbral)), unir: o.unir },
      rec: { alto: o.alto, alturaRel: o.alturaRel, rellenoRel: o.rellenoRel },
      lec: { ort: o.ort, librarius: o.librarius, alfabeto: o.alfabeto, alto: o.alto },
      kOrientacion: o.kOrientacion,
    };
  }

  // Dirección media de lectura de las líneas (radianes), ponderada por el largo.
  function giroDeLineas(lineas) {
    let sx = 0, sy = 0; for (const l of lineas) { sx += l.largo * l.u[0]; sy += l.largo * l.u[1]; }
    return Math.atan2(sy, sx);
  }
  function votoInvertir(s0, s1) {                  // s0, s1: seguridades de leer al derecho y al revés las mismas líneas
    let votos1 = 0, votos0 = 0, suma0 = 0, suma1 = 0;
    for (let i = 0; i < s0.length; i++) { suma0 += s0[i]; suma1 += s1[i]; if (s1[i] > s0[i]) votos1++; else if (s0[i] > s1[i]) votos0++; }
    return votos1 > votos0 || (votos1 === votos0 && suma1 > suma0);
  }
  const cuartosDeVuelta = (giroGrados) => ((Math.round(giroGrados / 90) % 4) + 4) % 4;

  // ------------------------------------------------------------------ la tubería en un solo hilo
  function crear(opciones) {
    const ctx = contextos(opciones);
    let tiempos = {};

    async function pasada(gris, w, h, orientar, eje) {
      const relleno = Math.trunc(percentil(gris, 80));
      const lineas = await detectarLineas(ctx.det, gris, w, h, tiempos);
      if (!lineas.length) return { texto: '', lineas: [], giro: 0, invertida: false };
      alinearEje(lineas, eje);
      let invertir = false;
      if (orientar) {
        const s0 = [], s1 = [];
        for (const i of muestraOrientacion(lineas, ctx.kOrientacion)) {
          const t0 = performance.now();
          const n0 = recortarLinea(ctx.rec, gris, w, h, lineas[i], false, relleno), n1 = recortarLinea(ctx.rec, gris, w, h, lineas[i], true, relleno);
          marca(tiempos, 'recortar', t0);
          s0.push((await leerRecorte(ctx.lec, n0, tiempos)).seguridad); s1.push((await leerRecorte(ctx.lec, n1, tiempos)).seguridad);
        }
        invertir = votoInvertir(s0, s1);
        tiempos.orientacion_lecturas = s0.length * 2;
      }
      for (const l of lineas) {
        const t0 = performance.now();
        const r = recortarLinea(ctx.rec, gris, w, h, l, invertir, relleno);
        marca(tiempos, 'recortar', t0);
        const lec = await leerRecorte(ctx.lec, r, tiempos);
        l.texto = lec.texto; l.seguridad = lec.seguridad;
        if (invertir) l.u = [-l.u[0], -l.u[1]];
      }
      const giro = giroDeLineas(lineas);
      const t0 = performance.now();
      const orden = ordenarLineas(lineas, giro);
      marca(tiempos, 'orden', t0);
      const ordenadas = orden.map((i) => lineas[i]);
      return { texto: ordenadas.map((l) => l.texto).filter((t) => t.trim()).join('\n'), lineas: ordenadas, giro: giro * 180 / Math.PI, invertida: invertir };
    }

    async function leer(gris, w, h, o) {
      tiempos = {};
      const orientar = !o || o.orientar !== false;
      const r = await pasada(gris, w, h, orientar, null);
      let resultado = r;
      if (orientar && r.lineas.length) {
        const k = cuartosDeVuelta(r.giro);
        if (k) {
          const t0 = performance.now();
          const g = girarPagina(gris, w, h, k);
          marca(tiempos, 'girar', t0);
          const r2 = await pasada(g.gris, g.w, g.h, false, [1, 0]);
          if (r2.lineas.length) { r2.giroPrimeraPasada = r.giro; resultado = r2; }
        }
      }
      resultado.tiempos = tiempos;
      return resultado;
    }

    return { leer, ordenarLineas };
  }

  global.LibrariusPagina = {
    crear,
    _nucleo: { contextos, detectarLineas, alinearEje, muestraOrientacion, recortarLinea, leerRecorte, giroDeLineas, votoInvertir, cuartosDeVuelta, girarPagina, ordenarLineas, percentil, DEFECTOS },
    _interno: { redimensionar, componentes, forma, ordenarLineas, girarPagina, percentil, redondearPar },
  };
})(typeof window !== 'undefined' ? window : globalThis);

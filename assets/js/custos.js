/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Custos Phone V2 (experimento E8, 2026-11-10; reemplaza a V1 = E3): red neuronal propia que encuentra las 4 esquinas
// de un documento en una foto. También el área de objetos rectangulares: pizarras, pantallas, cuadros (el lienzo sin el
// marco), pósters, rótulos, proyecciones, portadas y hojas de libros y cuadernos.
//
//   Custos.cargar()                        -> Promise<boolean>  (descarga el modelo y onnxruntime-web una vez)
//   Custos.listo()                         -> boolean
//   Custos.esperar(ms)                     -> Promise<boolean>  (listo antes de ms milisegundos)
//   Custos.detectar(fuente, ancho, alto)   -> Promise<{ esquinas, duda, encontrado }>
//   Custos.combinar(custos, clasico)       -> { esquinas, encontrado, origen }
//
// `fuente` es algo que se pueda dibujar en un canvas (canvas, video, imagen) y ancho/alto son sus
// medidas: las esquinas vuelven en píxeles de esa medida, en el orden arriba-izquierda, arriba-derecha,
// abajo-derecha, abajo-izquierda (el mismo que EscanerDocumento).
//
// Cómo funciona: la foto se pone sobre un margen gris del 25 % por lado (así una esquina cortada por el
// borde cae dentro del cuadro), se reduce a 224x224 y la red (MobileNetV3-Small entrenada desde cero con
// 7.371 fotos reales de documentos y 1.840 de objetos marcadas, más sintéticas; ~1 millón de parámetros) devuelve
// un mapa de calor por esquina; la
// esquina es el promedio del mapa y su dispersión es la «duda». Con duda alta no hay hoja a la vista.
// Corre con onnxruntime-web (WebAssembly, un hilo) dentro del navegador: la foto no sale del equipo.
// Si algo falla al cargar, listo() queda en false y la página sigue con el motor clásico.
//
// Combinación con el motor clásico: si los dos cuadriláteros coinciden (IoU >= 0,9) se promedian; si no, manda Custos.
// V2 contra V1 (mismas fotos nunca usadas para entrenar): 196 de documentos IoU 0,908 / 81 % bien (V1 0,901 / 74 %);
// banco de 186 objetos, como el sitio, 0,855 / 69 % (V1 0,762 / 39 %); banco de 17: 13 bien (V1 11). Igual de rápida
// (~9 ms en PC, ~62 ms con la CPU frenada x6) y del mismo peso (3,9 MB).
// Investigación y mediciones: Investigaciones\investigacion-escaner\05-estado-y-como-retomar.md
// Necesita assets/js/vendor.js antes.

const Custos = (() => {
  const S = 224, MARGEN = 0.25;
  // Por encima: «no hay hoja». V2 con 0,0045: detecta el 96,8 % de los fondos sin papel y pierde el 3,5 % de las fotos
  // con documento, que salían mal de todos modos (IoU medio 0,67). V1 usaba 0,006 (95,5 % / 2,0 %).
  const DUDA_MAX = 0.0045;
  const ACUERDO = 0.9;
  const MEDIA = [0.485, 0.456, 0.406], DESV = [0.229, 0.224, 0.225];
  const MODELO = new URL('../modelos/custos-phone-v2.onnx', document.currentScript.src).href;

  let sesion = null, promesa = null, cola = Promise.resolve();
  let lienzo = null, ctx = null;
  const entrada = new Float32Array(3 * S * S);

  function cargarScript(src) {
    return new Promise((ok, mal) => {
      const s = document.createElement('script');
      s.src = src; s.onload = ok; s.onerror = () => mal(new Error('No cargó ' + src));
      document.head.appendChild(s);
    });
  }

  function cargar() {
    if (!promesa) promesa = (async () => {
      try {
        if (typeof WebAssembly !== 'object') return false;
        if (!window.ort) await cargarScript(rutaVendor('onnxruntime-web/1.20.1/ort.wasm.min.js'));
        ort.env.wasm.numThreads = 1;   // el sitio no tiene aislamiento de origen (sin SharedArrayBuffer)
        ort.env.wasm.wasmPaths = {
          mjs: rutaVendor('onnxruntime-web/1.20.1/ort-wasm-simd-threaded.mjs'),
          wasm: rutaVendor('onnxruntime-web/1.20.1/ort-wasm-simd-threaded.wasm'),
        };
        sesion = await ort.InferenceSession.create(MODELO, { executionProviders: ['wasm'], graphOptimizationLevel: 'all' });
        return true;
      } catch (error) {
        console.warn('Custos no cargó; se usa el motor clásico.', error);
        sesion = null;
        return false;
      }
    })();
    return promesa;
  }

  const listo = () => !!sesion;

  function esperar(ms) {
    if (sesion) return Promise.resolve(true);
    return Promise.race([cargar(), new Promise((ok) => setTimeout(() => ok(false), ms))]).then(() => !!sesion);
  }

  // ---------- geometría ----------
  /** 4 puntos [x, y] en sentido horario (en pantalla) empezando por el más cercano a arriba-izquierda. */
  function ordenar(p) {
    const cx = (p[0][0] + p[1][0] + p[2][0] + p[3][0]) / 4, cy = (p[0][1] + p[1][1] + p[2][1] + p[3][1]) / 4;
    const r = p.slice().sort((a, b) => Math.atan2(a[1] - cy, a[0] - cx) - Math.atan2(b[1] - cy, b[0] - cx));
    let k = 0;
    r.forEach((q, i) => { if (q[0] + q[1] < r[k][0] + r[k][1]) k = i; });
    return r.slice(k).concat(r.slice(0, k));
  }

  const area = (q) => Math.abs(q.reduce((s, p, i) => { const n = q[(i + 1) % q.length]; return s + p[0] * n[1] - n[0] * p[1]; }, 0)) / 2;

  /** Sutherland-Hodgman: intersección de un polígono con otro convexo. */
  function recortar(sujeto, clip) {
    const dentro = (p, a, b) => (b[0] - a[0]) * (p[1] - a[1]) - (b[1] - a[1]) * (p[0] - a[0]) >= 0;
    const cruce = (p, q, a, b) => {
      const d = (p[0] - q[0]) * (a[1] - b[1]) - (p[1] - q[1]) * (a[0] - b[0]);
      if (Math.abs(d) < 1e-12) return q;
      const t = ((p[0] - a[0]) * (a[1] - b[1]) - (p[1] - a[1]) * (a[0] - b[0])) / d;
      return [p[0] + t * (q[0] - p[0]), p[1] + t * (q[1] - p[1])];
    };
    let salida = sujeto.slice();
    for (let i = 0; i < clip.length && salida.length; i++) {
      const a = clip[i], b = clip[(i + 1) % clip.length], ent = salida;
      salida = [];
      for (let j = 0; j < ent.length; j++) {
        const p = ent[j], q = ent[(j + 1) % ent.length];
        if (dentro(q, a, b)) { if (!dentro(p, a, b)) salida.push(cruce(p, q, a, b)); salida.push(q); }
        else if (dentro(p, a, b)) salida.push(cruce(p, q, a, b));
      }
    }
    return salida;
  }

  /** Intersección sobre unión de dos cuadriláteros convexos. */
  function iou(a, b) {
    const u = [b[1][0] - b[0][0], b[1][1] - b[0][1]], v = [b[2][0] - b[1][0], b[2][1] - b[1][1]];
    if (u[0] * v[1] - u[1] * v[0] < 0) b = b.slice().reverse();
    const inter = recortar(a, b), ia = inter.length >= 3 ? area(inter) : 0, union = area(a) + area(b) - ia;
    return union > 0 ? ia / union : 0;
  }

  const aPares = (esq) => esq.map((p) => [p.x, p.y]);
  const aPuntos = (q) => q.map(([x, y]) => ({ x, y }));

  // ---------- red ----------
  function preparar(fuente) {
    if (!lienzo) {
      lienzo = document.createElement('canvas');
      lienzo.width = lienzo.height = S;
      ctx = lienzo.getContext('2d', { willReadFrequently: true });
    }
    ctx.fillStyle = 'rgb(114, 114, 114)';
    ctx.fillRect(0, 0, S, S);
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    const m = (S * MARGEN) / (1 + 2 * MARGEN);
    ctx.drawImage(fuente, m, m, S - 2 * m, S - 2 * m);
    const d = ctx.getImageData(0, 0, S, S).data, n = S * S;
    for (let i = 0; i < n; i++)
      for (let c = 0; c < 3; c++) entrada[c * n + i] = (d[i * 4 + c] / 255 - MEDIA[c]) / DESV[c];
    return new ort.Tensor('float32', entrada, [1, 3, S, S]);
  }

  function detectar(fuente, ancho, alto) {
    // Una sola corrida a la vez: la sesión de WebAssembly no admite dos en paralelo.
    const tarea = cola.then(async () => {
      if (!sesion) throw new Error('Custos no está cargado');
      const r = await sesion.run({ imagen: preparar(fuente) });
      const xy = r.xy.data, v = r.var.data, q = [];
      for (let k = 0; k < 4; k++) {
        const fx = Math.min(1, Math.max(0, xy[k * 2] * (1 + 2 * MARGEN) - MARGEN));
        const fy = Math.min(1, Math.max(0, xy[k * 2 + 1] * (1 + 2 * MARGEN) - MARGEN));
        q.push([Math.min(ancho - 1, fx * ancho), Math.min(alto - 1, fy * alto)]);
      }
      const duda = (v[0] + v[1] + v[2] + v[3]) / 4;
      return { esquinas: aPuntos(ordenar(q)), duda, encontrado: duda <= DUDA_MAX };
    });
    cola = tarea.catch(() => {});
    return tarea;
  }

  /** custos: lo que devolvió detectar(); clasico: lo que devolvió EscanerDocumento.detectarEsquinas(). */
  function combinar(custos, clasico) {
    const c = custos && custos.encontrado ? ordenar(aPares(custos.esquinas)) : null;
    const v = clasico && clasico.encontrado ? ordenar(aPares(clasico.esquinas)) : null;
    if (c && v && iou(c, v) >= ACUERDO) return { esquinas: aPuntos(c.map((p, i) => [(p[0] + v[i][0]) / 2, (p[1] + v[i][1]) / 2])), encontrado: true, origen: 'ambos' };
    if (c) return { esquinas: aPuntos(c), encontrado: true, origen: 'custos' };
    if (v) return { esquinas: aPuntos(v), encontrado: true, origen: 'clasico' };
    return { esquinas: null, encontrado: false, origen: 'ninguno' };
  }

  return { cargar, listo, esperar, detectar, combinar, iou };
})();

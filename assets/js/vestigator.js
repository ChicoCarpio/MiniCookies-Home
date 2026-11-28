/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// «Buscar QR a fondo»: dos redes neuronales propias para los códigos QR que la lectura normal no consigue.
//   Vestigator Phone V1   encuentra los QR de una foto y marca sus 4 esquinas.
//   Restitutor Phone V1   limpia el recorte enderezado (borroso, desteñido, con ruido).
// El recorte limpio se le vuelve a dar a ZXing (leer-codigos.js): las redes no «adivinan» el contenido, que sale
// siempre de un código QR leído con su corrección de errores.
//
//   Vestigator.cargar()                  -> Promise   (descarga los dos modelos y onnxruntime-web una vez; rechaza si falla)
//   Vestigator.buscar(lienzo, sigue)     -> Promise<{ texto, formato } | null>
//
// `lienzo` es un canvas con la imagen. `sigue()` (opcional) se consulta entre un paso y otro: si devuelve false, la
// búsqueda se abandona. Corre con WebAssembly dentro del navegador: la imagen no sale del equipo.
// Necesita antes assets/js/vendor.js, ort-pagina.js y leer-codigos.js.
//
// Investigación y mediciones: Investigaciones\investigacion-lector-qr\00-resumen.md
const Vestigator = (() => {
  const base = document.currentScript.src;
  const MODELOS = ['vestigator-phone-v1.onnx', 'restitutor-phone-v1.onnx'].map((n) => new URL('../modelos/' + n, base).href);
  const PASADAS = [{ lado: 640, umbral: 0.2 }, { lado: 1280, umbral: 0.35 }];   // lado mayor con que mira la foto
  const MAXIMO_HALLAZGOS = 3;
  const L = 256;                        // lado del recorte de Restitutor
  const MARGENES = [0.1, 0.07, 0.14];   // cuánto fondo rodea al QR en el recorte, respecto de su lado

  let sesiones = null, promesa = null, cola = Promise.resolve();

  function cargar() {
    if (!promesa) promesa = (async () => {
      await cargarOrt();
      const [vestigator, restitutor] = await Promise.all(MODELOS.map((url) => ort.InferenceSession.create(url, { executionProviders: ['wasm'], graphOptimizationLevel: 'all' })));
      sesiones = { vestigator, restitutor };
    })().catch((error) => { promesa = null; throw error; });
    return promesa;
  }

  // ---------- imágenes en grises ({ d: Uint8ClampedArray, w, h }) ----------
  function grisDe(lienzo) {
    const { width: w, height: h } = lienzo;
    const p = lienzo.getContext('2d', { willReadFrequently: true }).getImageData(0, 0, w, h).data;
    const d = new Uint8ClampedArray(w * h);
    for (let i = 0, o = 0; i < d.length; i++, o += 4) {
      const gris = (p[o] * 19595 + p[o + 1] * 38470 + p[o + 2] * 7471 + 32768) >> 16;
      const a = p[o + 3];
      d[i] = a === 255 ? gris : (gris * a) / 255 + (255 - a);   // lo transparente, sobre blanco
    }
    return { d, w, h };
  }

  // Cambio de tamaño en dos pasadas, como el que se usó al entrenar: al achicar promedia (caja), al agrandar
  // interpola. `nucleo` y `radio` definen el filtro.
  const CAJA = { radio: 0.5, nucleo: (t) => (t > -0.5 && t <= 0.5 ? 1 : 0) };
  const LINEAL = { radio: 1, nucleo: (t) => Math.max(0, 1 - Math.abs(t)) };
  const CUBICO = {
    radio: 2,
    nucleo: (t) => {
      const x = Math.abs(t);
      if (x < 1) return (1.5 * x - 2.5) * x * x + 1;
      return x < 2 ? (((x - 5) * x + 8) * x - 4) * -0.5 : 0;
    },
  };

  function pesos(origen, destino, filtro) {
    const escala = origen / destino, ancho = Math.max(escala, 1), radio = filtro.radio * ancho;
    const salida = new Array(destino);
    for (let i = 0; i < destino; i++) {
      const centro = (i + 0.5) * escala;
      const desde = Math.max(0, Math.floor(centro - radio + 0.5)), hasta = Math.min(origen, Math.floor(centro + radio + 0.5));
      const w = new Float64Array(hasta - desde);
      let suma = 0;
      for (let x = desde; x < hasta; x++) suma += w[x - desde] = filtro.nucleo((x - centro + 0.5) / ancho);
      if (suma) for (let k = 0; k < w.length; k++) w[k] /= suma;
      salida[i] = { desde, w };
    }
    return salida;
  }

  function redimensionar(im, w, h, filtro) {
    const px = pesos(im.w, w, filtro), py = pesos(im.h, h, filtro);
    const medio = new Uint8ClampedArray(w * im.h);
    for (let y = 0; y < im.h; y++) {
      const fila = y * im.w;
      for (let x = 0; x < w; x++) {
        const { desde, w: p } = px[x];
        let v = 0;
        for (let k = 0; k < p.length; k++) v += im.d[fila + desde + k] * p[k];
        medio[y * w + x] = v;
      }
    }
    const d = new Uint8ClampedArray(w * h);
    for (let y = 0; y < h; y++) {
      const { desde, w: p } = py[y];
      for (let x = 0; x < w; x++) {
        let v = 0;
        for (let k = 0; k < p.length; k++) v += medio[(desde + k) * w + x] * p[k];
        d[y * w + x] = v;
      }
    }
    return { d, w, h };
  }

  // La ventana [x0, x0+w) × [y0, y0+h); lo que cae fuera de la imagen queda negro.
  function recortar(im, x0, y0, w, h) {
    const d = new Uint8ClampedArray(w * h);
    const desde = Math.max(0, -x0), hasta = Math.min(w, im.w - x0);
    for (let y = Math.max(0, -y0); y < Math.min(h, im.h - y0) && hasta > desde; y++) {
      const inicio = (y0 + y) * im.w + x0;
      d.set(im.d.subarray(inicio + desde, inicio + hasta), y * w + desde);
    }
    return { d, w, h };
  }

  // ---------- geometría ----------
  const centro = (q) => [(q[0][0] + q[1][0] + q[2][0] + q[3][0]) / 4, (q[0][1] + q[1][1] + q[2][1] + q[3][1]) / 4];
  const extension = (q) => Math.max(...[0, 1].map((e) => Math.max(...q.map((p) => p[e])) - Math.min(...q.map((p) => p[e]))));
  const distancia = (a, b) => Math.hypot(a[0] - b[0], a[1] - b[1]);

  /** 4 puntos en sentido horario (en pantalla), empezando por el de menor x+y. */
  function horario(q) {
    const [cx, cy] = centro(q);
    const r = q.slice().sort((a, b) => Math.atan2(a[1] - cy, a[0] - cx) - Math.atan2(b[1] - cy, b[0] - cx));
    let k = 0;
    r.forEach((p, i) => { if (p[0] + p[1] < r[k][0] + r[k][1]) k = i; });
    return r.slice(k).concat(r.slice(0, k));
  }

  // Los 8 coeficientes de la perspectiva que lleva los 4 puntos `de` a los 4 puntos `a` (eliminación de Gauss).
  function homografia(de, a) {
    const m = [];
    for (let i = 0; i < 4; i++) {
      const [x, y] = de[i], [u, v] = a[i];
      m.push([x, y, 1, 0, 0, 0, -u * x, -u * y, u], [0, 0, 0, x, y, 1, -v * x, -v * y, v]);
    }
    for (let c = 0; c < 8; c++) {
      let mayor = c;
      for (let f = c + 1; f < 8; f++) if (Math.abs(m[f][c]) > Math.abs(m[mayor][c])) mayor = f;
      [m[c], m[mayor]] = [m[mayor], m[c]];
      if (Math.abs(m[c][c]) < 1e-9) return null;
      for (let f = 0; f < 8; f++) {
        if (f === c) continue;
        const k = m[f][c] / m[c][c];
        for (let j = c; j < 9; j++) m[f][j] -= k * m[c][j];
      }
    }
    return m.map((fila, i) => fila[8] / fila[i]);
  }

  // El QR derecho en el cuadrado de 256 px de Restitutor, con `margen` alrededor. Fuera de la foto, blanco.
  function enderezar(im, q, margen) {
    const lado = Math.round(L / (1 + 2 * margen)), borde = Math.round(lado * margen), total = lado + 2 * borde;
    const a = borde, b = borde + lado;
    const h = homografia([[a, a], [b, a], [b, b], [a, b]], q);
    if (!h) return null;
    const salida = new Float32Array(L * L);
    for (let j = 0; j < L; j++) {
      const y = ((j + 0.5) * total) / L;
      for (let i = 0; i < L; i++) {
        const x = ((i + 0.5) * total) / L, z = h[6] * x + h[7] * y + 1;
        const u = (h[0] * x + h[1] * y + h[2]) / z, v = (h[3] * x + h[4] * y + h[5]) / z;
        let valor = 255;
        if (u >= 0 && v >= 0 && u < im.w && v < im.h) {
          const fu = Math.max(0, u - 0.5), fv = Math.max(0, v - 0.5);
          const x0 = Math.floor(fu), y0 = Math.floor(fv), x1 = Math.min(x0 + 1, im.w - 1), y1 = Math.min(y0 + 1, im.h - 1);
          const dx = fu - x0, dy = fv - y0;
          const arriba = im.d[y0 * im.w + x0] * (1 - dx) + im.d[y0 * im.w + x1] * dx;
          const abajo = im.d[y1 * im.w + x0] * (1 - dx) + im.d[y1 * im.w + x1] * dx;
          valor = arriba * (1 - dy) + abajo * dy;
        }
        salida[j * L + i] = valor / 255;
      }
    }
    return salida;
  }

  // ---------- redes ----------
  // Los QR que ve Vestigator: [{ puntaje, q }] en píxeles de la imagen, del más seguro al menos.
  async function localizar(im, lado, umbral) {
    const escala = lado / Math.max(im.w, im.h);
    const w = Math.max(1, Math.round(im.w * escala)), h = Math.max(1, Math.round(im.h * escala));
    const chica = redimensionar(im, w, h, escala < 1 ? CAJA : LINEAL);
    const W = Math.ceil(w / 32) * 32, H = Math.ceil(h / 32) * 32;
    const entrada = new Float32Array(W * H).fill(0.5);
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) entrada[y * W + x] = chica.d[y * w + x] / 255;
    const r = await sesiones.vestigator.run({ imagen: new ort.Tensor('float32', entrada, [1, 1, H, W]) });
    const picos = r.picos.data, desp = r.esquinas.data, ancho = r.picos.dims[3], celdas = ancho * r.picos.dims[2];
    const indices = [];
    for (let i = 0; i < celdas; i++) if (picos[i] >= umbral) indices.push(i);
    indices.sort((a, b) => picos[b] - picos[a]);
    const salida = [];
    for (const i of indices.slice(0, 8)) {
      const x = i % ancho, y = (i / ancho) | 0;
      const q = horario([0, 1, 2, 3].map((k) => [
        ((desp[2 * k * celdas + i] * 16 + x + 0.5) * 4) / escala,
        ((desp[(2 * k + 1) * celdas + i] * 16 + y + 0.5) * 4) / escala,
      ]));
      const c = centro(q);
      if (salida.some((o) => distancia(c, centro(o.q)) < 0.35 * extension(o.q))) continue;
      salida.push({ puntaje: picos[i], q });
    }
    return salida;
  }

  // Segunda mirada, de cerca, alrededor de un hallazgo: esquinas más precisas (o null).
  async function afinar(im, q) {
    const c = centro(q), r = extension(q) * 1.1;
    const x0 = Math.trunc(c[0] - r), y0 = Math.trunc(c[1] - r);
    const w = Math.trunc(c[0] + r) - x0, h = Math.trunc(c[1] + r) - y0;
    if (w < 8 || h < 8 || w > 4 * im.w || h > 4 * im.h) return null;
    let mejor = null, cerca = r * 0.5;
    for (const hallado of await localizar(recortar(im, x0, y0, w, h), 384, 0.15)) {
      const movido = hallado.q.map(([x, y]) => [x + x0, y + y0]);
      const d = distancia(centro(movido), c);
      if (d < cerca) { cerca = d; mejor = movido; }
    }
    return mejor;
  }

  async function limpiar(recorte) {
    const r = await sesiones.restitutor.run({ recorte: new ort.Tensor('float32', recorte, [1, 1, L, L]) });
    const oscuro = r.oscuro.data, d = new Uint8ClampedArray(L * L);
    for (let i = 0; i < d.length; i++) d[i] = Math.floor((1 - oscuro[i]) * 255);
    return { d, w: L, h: L };
  }

  function aImageData(im) {
    const salida = new ImageData(im.w, im.h);
    for (let i = 0, o = 0; i < im.d.length; i++, o += 4) {
      salida.data[o] = salida.data[o + 1] = salida.data[o + 2] = im.d[i];
      salida.data[o + 3] = 255;
    }
    return salida;
  }

  // El recorte limpio, tal cual, al doble de tamaño y en blanco y negro puro.
  const FORMAS = [
    (im) => im,
    (im) => redimensionar(im, im.w * 2, im.h * 2, CUBICO),
    (im) => ({ d: im.d.map((v) => (v > 127 ? 255 : 0)), w: im.w, h: im.h }),
  ];

  // Deja que la página se redibuje entre una red y la siguiente.
  const respirar = () => new Promise((ok) => setTimeout(ok, 0));

  async function buscarEnGris(im, sigue) {
    const vistos = [];
    for (const { lado, umbral } of PASADAS) {
      await respirar();
      if (!sigue()) return null;
      for (const { q } of await localizar(im, lado, umbral)) {
        const c = centro(q);
        if (vistos.some((o) => distancia(c, centro(o)) < 0.4 * extension(o))) continue;
        vistos.push(q);
        const fino = await afinar(im, q);
        for (const esquinas of fino ? [fino, q] : [q]) {
          for (const margen of MARGENES) {
            await respirar();
            if (!sigue()) return null;
            const recorte = enderezar(im, esquinas, margen);
            if (!recorte) continue;
            const limpio = await limpiar(recorte);
            for (const forma of FORMAS) {
              const hallado = await LeerCodigos.qrEn(aImageData(forma(limpio)));
              if (hallado) return hallado;
            }
          }
        }
        if (vistos.length >= MAXIMO_HALLAZGOS) return null;
      }
    }
    return null;
  }

  function buscar(lienzo, sigue = () => true) {
    // Una búsqueda a la vez: una sesión de WebAssembly no admite dos corridas en paralelo.
    const tarea = cola.then(async () => {
      if (!sigue()) return null;
      await cargar();
      return buscarEnGris(grisDe(lienzo), sigue);
    });
    cola = tarea.catch(() => {});
    return tarea;
  }

  return { cargar, buscar };
})();

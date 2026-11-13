/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Librarius (Beta): el lector de texto con IA propia de Mini Cookies (Librarius Phone V1 + Versus Phone V1), como motor
// alternativo a Tesseract para ocr.js. Corre en hilos de trabajo (Web Workers) con onnxruntime-web: el documento no sale del equipo.
//
//   OcrLibrarius.soportaIdiomas(['spa', 'eng'])   -> boolean (solo idiomas latinos; chino y japonés quedan en Tesseract)
//   const motor = await OcrLibrarius.crear()      -> { leerPagina(canvas), cerrar() }   (rechaza si el navegador no puede)
//   const r = await motor.leerPagina(canvas)      -> { texto, palabras: [{ texto, x, y, ancho, alto, angulo }] }
//        x, y: esquina inferior izquierda de la palabra sobre la línea base, en píxeles del canvas (y hacia abajo);
//        ancho/alto en píxeles; angulo: grados antihorarios (para dibujar texto invisible inclinado en el PDF).
//        Las cajas por palabra se reparten dentro de cada línea según la cantidad de caracteres (aproximadas).
// Supone cargados vendor.js (rutaVendor) y los tres archivos librarius-*.js (esta pieza los carga si faltan).
(function () {
  'use strict';
  const IDIOMAS = new Set(['spa', 'eng', 'por', 'fra', 'deu', 'ita']);
  const LADO_MAXIMO = 3200;   // lado largo máximo de la página que se le pasa a Librarius (más no aporta: las líneas se reescalan a 24 px)

  const soportaIdiomas = (idiomas) => idiomas.length > 0 && idiomas.every((i) => IDIOMAS.has(i));

  const BASE = document.currentScript ? new URL('.', document.currentScript.src).href : new URL('assets/js/', location.href).href;
  const absoluta = (ruta) => new URL(ruta, location.href).href;

  function cargarScript(src) {
    return new Promise((ok, mal) => {
      if ([...document.scripts].some((s) => s.src === src)) return ok();
      const s = document.createElement('script');
      s.src = src; s.onload = ok; s.onerror = () => mal(new Error('No cargó ' + src));
      document.head.appendChild(s);
    });
  }

  let promesaCarga = null, motorUnico = null;

  async function crear() {
    if (typeof Worker !== 'function' || typeof WebAssembly !== 'object') throw new Error('Este navegador no puede correr Librarius.');
    if (motorUnico) return motorUnico;
    if (!promesaCarga) promesaCarga = (async () => {
      if (!window.LibrariusPagina) await cargarScript(BASE + 'librarius-pagina.js');
      if (!window.LibrariusPaginaHilos) await cargarScript(BASE + 'librarius-pagina-hilos.js');
      const alfabeto = (await (await fetch(BASE + '../modelos/librarius-alfabeto.json')).json()).alfabeto;
      const lector = await LibrariusPaginaHilos.crear({
        rutaOrt: absoluta(rutaVendor('onnxruntime-web/1.20.1/')),
        rutaPagina: BASE + 'librarius-pagina.js',
        rutaHilo: BASE + 'librarius-hilo.js',
        modeloVersus: BASE + '../modelos/versus-phone-v1.onnx',
        modeloLibrarius: BASE + '../modelos/librarius-phone-v1-int8.onnx',
        alfabeto,
        hilos: Math.max(1, Math.min(4, (navigator.hardwareConcurrency || 4) - 1)),
      });
      return lector;
    })();
    let lector;
    try { lector = await promesaCarga; } catch (e) { promesaCarga = null; throw e; }

    // Gris de la página (≤ LADO_MAXIMO) y factor para volver a los píxeles del canvas original.
    function aGris(canvas) {
      const f = Math.min(1, LADO_MAXIMO / Math.max(canvas.width, canvas.height));
      const w = Math.max(1, Math.round(canvas.width * f)), h = Math.max(1, Math.round(canvas.height * f));
      const c = document.createElement('canvas'); c.width = w; c.height = h;
      const x = c.getContext('2d', { willReadFrequently: true });
      x.fillStyle = '#fff'; x.fillRect(0, 0, w, h);
      x.imageSmoothingQuality = 'high';
      x.drawImage(canvas, 0, 0, w, h);
      const d = x.getImageData(0, 0, w, h).data, g = new Uint8Array(w * h);
      for (let i = 0, j = 0; j < g.length; i += 4, j++) g[j] = (d[i] * 77 + d[i + 1] * 150 + d[i + 2] * 29) >> 8;
      return { gris: g, w, h, factor: f };
    }

    // Las líneas vienen en el marco de la página girada (si Librarius la enderezó k·90°): se pasan al marco original.
    function aMarcoOriginal(l, k, w, h) {
      let [cx, cy] = l.centro, [ux, uy] = l.u;
      if (k === 1) { [cx, cy, ux, uy] = [w - 1 - cy, cx, -uy, ux]; }
      else if (k === 2) { [cx, cy, ux, uy] = [w - 1 - cx, h - 1 - cy, -ux, -uy]; }
      else if (k === 3) { [cx, cy, ux, uy] = [cy, h - 1 - cx, uy, -ux]; }
      return { cx, cy, ux, uy };
    }

    async function leerPagina(canvas) {
      const g = aGris(canvas);
      const r = await lector.leer(g.gris, g.w, g.h);
      const k = r.giroPrimeraPasada !== undefined ? LibrariusPagina._nucleo.cuartosDeVuelta(r.giroPrimeraPasada) : 0;
      const palabras = [];
      for (const l of r.lineas) {
        const texto = (l.texto || '').trim();
        if (!texto) continue;
        const { cx, cy, ux, uy } = aMarcoOriginal(l, k, g.w, g.h);
        const vx = -uy, vy = ux;                              // perpendicular hacia "abajo" de la línea
        const largo = l.largo / g.factor, alto = l.alto / g.factor;
        const total = l.texto.length || 1;
        const angulo = Math.atan2(-uy, ux) * 180 / Math.PI;   // antihorario, con la y de la imagen hacia abajo
        const re = /\S+/g; let m;
        while ((m = re.exec(l.texto))) {
          const t0 = m.index / total * largo, t1 = (m.index + m[0].length) / total * largo;
          const ox = cx / g.factor + ux * (t0 - largo / 2) + vx * alto * 0.9, oy = cy / g.factor + uy * (t0 - largo / 2) + vy * alto * 0.9;
          palabras.push({ texto: m[0], x: ox, y: oy, ancho: t1 - t0, alto: alto * 1.8, angulo });
        }
      }
      return { texto: r.texto, palabras };
    }

    motorUnico = { leerPagina, cerrar() { lector.cerrar(); motorUnico = null; promesaCarga = null; } };
    return motorUnico;
  }

  window.OcrLibrarius = { soportaIdiomas, crear };
})();

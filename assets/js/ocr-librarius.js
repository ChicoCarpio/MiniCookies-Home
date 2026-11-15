/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Librarius (Beta): el lector de texto con IA propia de Mini Cookies (Librarius Phone V1, E V1, C V1 u O V1 + Versus Phone V1), como motor
// alternativo a Tesseract para ocr.js. Corre en hilos de trabajo (Web Workers) con onnxruntime-web: el documento no sale del equipo.
//
//   OcrLibrarius.soportaIdiomas(['spa', 'eng'])   -> boolean (los seis primeros con, a lo sumo, UNA familia: EE1, EC1 o EO1)
//   const motor = await OcrLibrarius.crear(['spa', 'swe'])   -> { leerPagina(canvas), cerrar() }   (rechaza si el navegador no puede)
//        Con chino, japonés o coreano usa Librarius O V1; con cirílico, Librarius C V1; con algún idioma de IDIOMAS_EE1,
//        Librarius E V1 (23 idiomas); con solo los seis de Librarius Phone V1, ese. Cada modelo trae el alfabeto de los seis primeros.
//   const r = await motor.leerPagina(canvas)      -> { texto, palabras: [{ texto, x, y, ancho, alto, angulo }] }
//        x, y: esquina inferior izquierda de la palabra sobre la línea base, en píxeles del canvas (y hacia abajo);
//        ancho/alto en píxeles; angulo: grados antihorarios (para dibujar texto invisible inclinado en el PDF).
//        Las cajas por palabra se reparten dentro de cada línea según la cantidad de caracteres (aproximadas).
// Supone cargados vendor.js (rutaVendor) y los tres archivos librarius-*.js (esta pieza los carga si faltan).
(function () {
  'use strict';
  const IDIOMAS = new Set(['spa', 'eng', 'por', 'fra', 'deu', 'ita']);
  // Librarius E V1 (EE1) lee estos 17 además de los seis anteriores (códigos de Tesseract).
  const IDIOMAS_EE1 = new Set(['nld', 'swe', 'dan', 'nor', 'fin', 'pol', 'ces', 'slk', 'hun', 'ron', 'tur', 'cat', 'hrv', 'slv', 'lit', 'lav', 'est']);
  // Librarius O V1 (EO1) lee chino simplificado, japonés y coreano, y también el alfabeto latino de los seis primeros (no los acentos de EE1).
  // Entrada de 32 px de alto (los demás, 24).
  const IDIOMAS_EO1 = new Set(['chi_sim', 'jpn', 'kor']);
  // Librarius C V1 (EC1) lee ruso (también con la ortografía anterior a 1918), ucraniano, bielorruso, búlgaro, serbio y macedonio.
  const IDIOMAS_EC1 = new Set(['rus', 'ukr', 'bel', 'bul', 'srp', 'mkd']);
  const FAMILIAS = [['eo1', IDIOMAS_EO1], ['ec1', IDIOMAS_EC1], ['ee1', IDIOMAS_EE1]];
  const familiasDe = (idiomas) => FAMILIAS.filter(([, s]) => (idiomas || []).some((i) => s.has(i))).map(([c]) => c);
  const modeloPara = (idiomas) => familiasDe(idiomas)[0] || 'e4';
  const MODELOS = {
    e4: { lector: 'librarius-phone-v1-int8.onnx', alfabeto: 'librarius-alfabeto.json', alto: 24 },
    ee1: { lector: 'librarius-ee1-int8.onnx', alfabeto: 'librarius-alfabeto-ee1.json', alto: 24 },
    ec1: { lector: 'librarius-ec1-int8.onnx', alfabeto: 'librarius-alfabeto-ec1.json', alto: 24 },
    eo1: { lector: 'librarius-eo1-int8.onnx', alfabeto: 'librarius-alfabeto-eo1.json', alto: 32 },
  };
  const UMBRAL_FONDO_OSCURO = 110;   // mediana del gris por debajo de esto = fondo oscuro, se invierte antes de leer
  const LADO_MAXIMO = 3200;   // lado largo máximo de la página que se le pasa a Librarius (más no aporta: las líneas se reescalan a 24 o 32 px)

  // Los seis primeros idiomas se mezclan con cualquiera (todos los modelos traen su alfabeto); dos familias a la vez (p. ej. ruso y
  // japonés, o ruso y polaco) no las lee ningún modelo: ahí se usa Tesseract.
  const soportaIdiomas = (idiomas) => {
    if (!idiomas.length || !idiomas.every((i) => IDIOMAS.has(i) || FAMILIAS.some(([, s]) => s.has(i)))) return false;
    return familiasDe(idiomas).length <= 1;
  };

  const BASE = document.currentScript ? new URL('.', document.currentScript.src).href : new URL('assets/js/', location.href).href;
  const absoluta = (ruta) => new URL(ruta, location.href).href;
  // Versión en la dirección de los scripts: el sitio deja que el navegador los guarde horas, y una copia vieja
  // mezclada con una nueva rompe la lectura. Subirla cuando cambie librarius-pagina.js o librarius-pagina-hilos.js.
  const V = '?v=6';

  function cargarScript(src) {
    return new Promise((ok, mal) => {
      if ([...document.scripts].some((s) => s.src === src)) return ok();
      const s = document.createElement('script');
      s.src = src; s.onload = ok; s.onerror = () => mal(new Error('No cargó ' + src));
      document.head.appendChild(s);
    });
  }

  const promesasCarga = {}, motores = {};      // por modelo ('e4' o 'ee1')
  // Las páginas pasan de a una por el único lector: su detector guarda la página que está leyendo, y Flujos pide varias a la vez.
  let fila = Promise.resolve();

  async function crear(idiomas) {
    const clave = modeloPara(idiomas);
    if (typeof Worker !== 'function' || typeof WebAssembly !== 'object') throw new Error('Este navegador no puede correr Librarius.');
    if (motores[clave]) return motores[clave];
    if (!promesasCarga[clave]) promesasCarga[clave] = (async () => {
      if (!window.LibrariusPagina) await cargarScript(BASE + 'librarius-pagina.js' + V);
      if (!window.LibrariusPaginaHilos) await cargarScript(BASE + 'librarius-pagina-hilos.js' + V);
      const alfabeto = (await (await fetch(BASE + '../modelos/' + MODELOS[clave].alfabeto)).json()).alfabeto;
      const lector = await LibrariusPaginaHilos.crear({
        rutaOrt: absoluta(rutaVendor('onnxruntime-web/1.20.1/')),
        rutaPagina: BASE + 'librarius-pagina.js' + V,
        rutaHilo: BASE + 'librarius-hilo.js',
        modeloVersus: BASE + '../modelos/versus-phone-v1.onnx' + V,
        modeloLibrarius: BASE + '../modelos/' + MODELOS[clave].lector + V,
        alfabeto,
        opciones: { alto: MODELOS[clave].alto },
        hilos: Math.max(1, Math.min(4, (navigator.hardwareConcurrency || 4) - 1)),
      });
      return lector;
    })();
    let lector;
    try { lector = await promesasCarga[clave]; } catch (e) { promesasCarga[clave] = null; throw e; }

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
      const hist = new Uint32Array(256);
      for (let i = 0, j = 0; j < g.length; i += 4, j++) hist[g[j] = (d[i] * 77 + d[i + 1] * 150 + d[i + 2] * 29) >> 8]++;
      // Versus y Librarius aprendieron con letra oscura sobre fondo claro: con fondo oscuro (capturas en modo oscuro,
      // carteles) casi no encuentran líneas. Si la mediana de la página es oscura, se invierte el gris.
      let acumulado = 0, mediana = 0;
      while (mediana < 255 && (acumulado += hist[mediana]) < g.length / 2) mediana++;
      if (mediana < UMBRAL_FONDO_OSCURO) for (let j = 0; j < g.length; j++) g[j] = 255 - g[j];
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

    function leerPagina(canvas) {
      const turno = fila.then(() => leerUna(canvas));
      fila = turno.catch(() => {});
      return turno;
    }

    async function leerUna(canvas) {
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

    // Varios pedidos simultáneos esperan la misma carga: todos reciben el mismo motor (y la misma fila).
    if (motores[clave]) return motores[clave];
    motores[clave] = { leerPagina, cerrar() { lector.cerrar(); motores[clave] = null; promesasCarga[clave] = null; } };
    return motores[clave];
  }

  window.OcrLibrarius = { soportaIdiomas, crear };
})();

/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Librarius: el lector de texto con IA propia de Mini Cookies (Librarius Phone V3, Phone E V1, Phone C V1 y Phone O V1, con Versus Phone V2),
// el motor de OCR del sitio (ocr.js). Corre en hilos de trabajo (Web Workers) con onnxruntime-web: el documento no sale del equipo.
//
//   OcrLibrarius.soportaIdiomas(['spa', 'rus'])   -> boolean (cualquier combinación de los 32 idiomas; una mezcla de grupos usa varios modelos)
//   const motor = await OcrLibrarius.crear(['spa', 'swe'])   -> { leerPagina(canvas), cerrar() }   (rechaza si el navegador no puede)
//        Con chino, japonés o coreano usa Librarius Phone O V1; con cirílico, Librarius Phone C V1; con algún idioma de IDIOMAS_PHONE_E_V1,
//        Librarius Phone E V1 (23 idiomas); con solo los seis de Librarius Phone V3, ese. Cada modelo trae el alfabeto de los seis primeros;
//        solo Phone V3 lee además los signos nuevos (₡ y otras monedas, casillas, viñetas, flechas, fracciones).
//   const r = await motor.leerPagina(canvas)      -> { texto, palabras: [{ texto, x, y, ancho, alto, angulo }] }
//        x, y: esquina inferior izquierda de la palabra sobre la línea base, en píxeles del canvas (y hacia abajo);
//        ancho/alto en píxeles; angulo: grados antihorarios (para dibujar texto invisible inclinado en el PDF).
//        Las cajas por palabra se reparten dentro de cada línea según la cantidad de caracteres (aproximadas).
// Supone cargados vendor.js (rutaVendor) y los tres archivos librarius-*.js (esta pieza los carga si faltan).
(function () {
  'use strict';
  const IDIOMAS = new Set(['spa', 'eng', 'por', 'fra', 'deu', 'ita']);
  // Librarius Phone E V1 lee estos 17 además de los seis anteriores (códigos de idioma ISO 639-2, los de ocr-pdf.html).
  const IDIOMAS_PHONE_E_V1 = new Set(['nld', 'swe', 'dan', 'nor', 'fin', 'pol', 'ces', 'slk', 'hun', 'ron', 'tur', 'cat', 'hrv', 'slv', 'lit', 'lav', 'est']);
  // Librarius Phone O V1 lee chino simplificado, japonés y coreano, y también el alfabeto latino de los seis primeros (no los acentos de Phone E V1).
  // Entrada de 32 px de alto (los demás, 24).
  const IDIOMAS_PHONE_O_V1 = new Set(['chi_sim', 'jpn', 'kor']);
  // Librarius Phone C V1 lee ruso (también con la ortografía anterior a 1918), ucraniano, bielorruso, búlgaro, serbio y macedonio.
  const IDIOMAS_PHONE_C_V1 = new Set(['rus', 'ukr', 'bel', 'bul', 'srp', 'mkd']);
  // Cada clave es el nombre público del modelo (los archivos de assets/modelos/ se llaman igual): nunca el código interno del entrenamiento.
  const FAMILIAS = [['phone-o-v1', IDIOMAS_PHONE_O_V1], ['phone-c-v1', IDIOMAS_PHONE_C_V1], ['phone-e-v1', IDIOMAS_PHONE_E_V1]];
  const familiasDe = (idiomas) => FAMILIAS.filter(([, s]) => (idiomas || []).some((i) => s.has(i))).map(([c]) => c);
  const MODELOS = {
    'phone-v3': { lector: 'librarius-phone-v3-int8.onnx', alfabeto: 'librarius-alfabeto-phone-v3.json', alto: 24 },
    'phone-e-v1': { lector: 'librarius-phone-e-v1-int8.onnx', alfabeto: 'librarius-alfabeto-phone-e-v1.json', alto: 24 },
    'phone-c-v1': { lector: 'librarius-phone-c-v1-int8.onnx', alfabeto: 'librarius-alfabeto-phone-c-v1.json', alto: 24 },
    'phone-o-v1': { lector: 'librarius-phone-o-v1-int8.onnx', alfabeto: 'librarius-alfabeto-phone-o-v1.json', alto: 32 },
  };
  const LADO_MAXIMO = 3200;   // lado largo máximo de la página que se le pasa a Librarius (más no aporta: las líneas se reescalan a 24 o 32 px)

  // Cualquier combinación de los 32 idiomas: los seis primeros van con cualquier modelo (todos traen su alfabeto) y una mezcla de
  // grupos (p. ej. ruso y polaco, o japonés y ruso) se lee con un modelo por grupo y se elige, línea por línea, la lectura más segura.
  // Si se pidió alguno de los seis primeros junto con otro grupo, también lee Phone V3: es el único con los signos nuevos (₡…) y el mejor en esos seis.
  const soportaIdiomas = (idiomas) => idiomas.length > 0 && idiomas.every((i) => IDIOMAS.has(i) || FAMILIAS.some(([, s]) => s.has(i)));
  const modelosPara = (idiomas) => {
    const f = familiasDe(idiomas);
    return !f.length ? ['phone-v3'] : (idiomas || []).some((i) => IDIOMAS.has(i)) ? ['phone-v3', ...f] : f;
  };
  const enIngles = () => /^en/i.test(document.documentElement.lang || '');

  const BASE = document.currentScript ? new URL('.', document.currentScript.src).href : new URL('assets/js/', location.href).href;
  const absoluta = (ruta) => new URL(ruta, location.href).href;
  // Versión en la dirección de los scripts: el sitio deja que el navegador los guarde horas, y una copia vieja
  // mezclada con una nueva rompe la lectura. Subirla cuando cambie librarius-pagina.js o librarius-pagina-hilos.js.
  const V = '?v=8';

  function cargarScript(src) {
    return new Promise((ok, mal) => {
      if ([...document.scripts].some((s) => s.src === src)) return ok();
      const s = document.createElement('script');
      s.src = src; s.onload = ok; s.onerror = () => mal(errorMC('carga-no-cargo', { archivo: src }, 'No cargó {archivo}'));
      document.head.appendChild(s);
    });
  }

  const promesasCarga = {}, motores = {};      // carga por modelo ('phone-v3', 'phone-e-v1'…) y motor por combinación de modelos
  const alfabetos = {};                        // letras que conoce cada modelo (Set), para combinar lecturas
  // Las páginas pasan de a una por el único lector: su detector guarda la página que está leyendo, y Flujos pide varias a la vez.
  let fila = Promise.resolve();

  function cargarModelo(clave) {
    if (!promesasCarga[clave]) promesasCarga[clave] = (async () => {
      if (!window.LibrariusPagina) await cargarScript(BASE + 'librarius-pagina.js' + V);
      if (!window.LibrariusPaginaHilos) await cargarScript(BASE + 'librarius-pagina-hilos.js' + V);
      const alfabeto = (await (await fetch(BASE + '../modelos/' + MODELOS[clave].alfabeto)).json()).alfabeto;
      alfabetos[clave] = new Set(alfabeto);
      return LibrariusPaginaHilos.crear({
        rutaOrt: absoluta(rutaVendor('onnxruntime-web/1.20.1/')),
        rutaPagina: BASE + 'librarius-pagina.js' + V,
        rutaHilo: BASE + 'librarius-hilo.js',
        modeloVersus: BASE + '../modelos/versus-phone-v2.onnx' + V,
        modeloLibrarius: BASE + '../modelos/' + MODELOS[clave].lector + V,
        alfabeto,
        opciones: { alto: MODELOS[clave].alto },
        hilos: Math.max(1, Math.min(4, (navigator.hardwareConcurrency || 4) - 1)),
      });
    })().catch((e) => { promesasCarga[clave] = null; throw e; });
    return promesasCarga[clave];
  }

  // Con varios modelos: la página se lee con cada uno, la base es la lectura más segura en conjunto (su orientación y sus líneas) y
  // cada línea se queda con el texto del modelo que la leyó con más seguridad (la línea del otro modelo es la de centro más cercano).
  // Excepción: si una lectura difiere de la otra solo en letras que el otro modelo no conoce (p. ej. «₡12.500» contra «(12.500»),
  // gana esa aunque la otra parezca más segura: el otro modelo no podía leerlas y puso la letra suya más parecida.
  function distancia(a, b) {
    a = [...a]; b = [...b];
    let previa = Array.from({ length: b.length + 1 }, (_, j) => j);
    for (let i = 1; i <= a.length; i++) {
      const fila = [i];
      for (let j = 1; j <= b.length; j++) fila[j] = Math.min(previa[j] + 1, fila[j - 1] + 1, previa[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      previa = fila;
    }
    return previa[b.length];
  }
  function soloLetrasAjenas(a, b, alfabetoB) {
    if (!a || !b || !alfabetoB) return false;
    const ajenas = [...a].filter((c) => c !== ' ' && !alfabetoB.has(c)).length;
    return ajenas > 0 && distancia(a, b) <= ajenas;
  }
  function combinar(lecturas) {
    const media = (r) => r.lineas.length ? r.lineas.reduce((s, l) => s + (l.seguridad || 0) * l.largo, 0) / r.lineas.reduce((s, l) => s + l.largo, 0) : 0;
    const orden = [...lecturas].sort((a, b) => media(b) - media(a));
    const base = orden[0];
    const lineas = base.lineas.map((l) => {
      let mejor = l, deMejor = base;
      for (const otra of orden.slice(1)) {
        let cerca = null, d = Infinity;
        for (const m of otra.lineas) {
          const dd = Math.hypot(m.centro[0] - l.centro[0], m.centro[1] - l.centro[1]);
          if (dd < d) { d = dd; cerca = m; }
        }
        if (!cerca || d >= 0.5 * Math.max(l.alto, 1)) continue;
        const ajenaCerca = soloLetrasAjenas(cerca.texto, mejor.texto, deMejor.alfabeto);
        const ajenaMejor = soloLetrasAjenas(mejor.texto, cerca.texto, otra.alfabeto);
        if (ajenaCerca || (!ajenaMejor && (cerca.seguridad || 0) > (mejor.seguridad || 0))) {
          mejor = { ...l, texto: cerca.texto, seguridad: cerca.seguridad }; deMejor = otra;
        }
      }
      return mejor;
    });
    return { ...base, lineas, texto: lineas.map((l) => l.texto).filter((x) => x && x.trim()).join('\n') };
  }

  async function crear(idiomas) {
    if (typeof Worker !== 'function' || typeof WebAssembly !== 'object') {
      throw new Error(enIngles() ? 'This browser cannot run the text recognizer (it needs Web Workers and WebAssembly). Try an up-to-date Chrome, Edge, Firefox or Safari.'
        : 'Este navegador no puede correr el lector de texto (necesita Web Workers y WebAssembly). Probá con Chrome, Edge, Firefox o Safari actualizados.');
    }
    const claves = modelosPara(idiomas), clave = claves.join('+');
    if (motores[clave]) return motores[clave];
    const lectores = await Promise.all(claves.map(cargarModelo));

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
      // Ya no se invierte la página entera cuando es oscura: Versus Phone V2 encuentra las líneas con cualquier polaridad y cada línea de letra clara
      // sobre fondo oscuro se invierte sola al recortarla (letraClaraLinea en librarius-pagina.js), así también se leen las páginas con fondos mezclados.
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
      const lecturas = [];
      for (let i = 0; i < lectores.length; i++) {      // de a uno: comparten los núcleos
        lecturas.push({ ...(await lectores[i].leer(g.gris, g.w, g.h)), alfabeto: alfabetos[claves[i]] });
      }
      const r = lecturas.length > 1 ? combinar(lecturas) : lecturas[0];
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
    motores[clave] = { leerPagina, cerrar() { claves.forEach((c, i) => { lectores[i].cerrar(); promesasCarga[c] = null; }); motores[clave] = null; } };
    return motores[clave];
  }

  window.OcrLibrarius = { soportaIdiomas, crear };
})();

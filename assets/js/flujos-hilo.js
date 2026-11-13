/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Hilo de trabajo de Flujos (Web Worker). flujos.js abre varios, uno por núcleo, y les reparte los pasos de los
// módulos marcados con `hilo: true`: así un lote grande usa todos los núcleos y no uno solo.
//
// El hilo carga las MISMAS piezas que la página (vendor.js, flujos.js, flujos-modulos.js y lo que cada módulo trae en
// preparar()): no hay una copia del código para el hilo. Para eso imita lo poco de la página que esas piezas usan:
//   document.createElement('canvas')  -> un OffscreenCanvas (con toBlob, como el lienzo de la página)
//   new Image() / createElement('img') -> una imagen que se decodifica con createImageBitmap; drawImage la acepta
//   createElement('script') + head.appendChild -> importScripts, con document.currentScript apuntando al script
//   document.fonts -> las fuentes del hilo (PDF.js carga ahí las de cada PDF)
// Lo que no se imita (DOMParser, elementos de verdad) no está: los módulos que lo necesitan no llevan `hilo: true`
// y corren en la página. Si algo falla en el hilo por el entorno, flujos.js vuelve a hacer ese paso en la página.
//
// Mensajes:  { tipo: 'iniciar', scripts: [url…], idioma }        -> { tipo: 'listo' } | { tipo: 'error', mensaje }
//            { tipo: 'paso', id, modulo, ficha, opciones, contexto } -> { tipo: 'hecho', id, r, vacio }
//                                                                   | { tipo: 'fallo', id, enPreparar, nombre, mensaje }
(function () {
  self.window = self;
  let idioma = 'es';
  self.idiomaActual = () => idioma;

  // ---------- scripts ----------
  let actual = null;
  const cargados = [];
  function importar(src) {
    const previo = actual;
    actual = { src };
    try { importScripts(src); } finally { actual = previo; }
    cargados.push({ src });
  }

  // ---------- lienzo ----------
  const pLienzo = OffscreenCanvas.prototype;
  if (!pLienzo.toBlob) {
    pLienzo.toBlob = function (alListo, tipo, calidad) {
      this.convertToBlob({ type: tipo || 'image/png', quality: calidad }).then(alListo, () => alListo(null));
    };
  }
  if (!('style' in pLienzo)) Object.defineProperty(pLienzo, 'style', { get() { return this._estilo || (this._estilo = {}); } });

  // ---------- imagen ----------
  class ImagenHilo {
    constructor() {
      this.width = this.height = this.naturalWidth = this.naturalHeight = 0;
      this.complete = false;
      this.onload = this.onerror = null;
      this.style = {};
      this._src = '';
      this._mapa = null;
      this._carga = null;
    }
    get src() { return this._src; }
    set src(url) {
      this._src = url;
      this.complete = false;
      this._carga = fetch(url).then((r) => r.blob()).then((b) => createImageBitmap(b)).then((mapa) => {
        this._mapa = mapa;
        this.width = this.naturalWidth = mapa.width;
        this.height = this.naturalHeight = mapa.height;
        this.complete = true;
        if (this.onload) this.onload({ type: 'load', target: this });
      }, (error) => {
        if (this.onerror) this.onerror(error);
        throw error;
      });
      this._carga.catch(() => {});
    }
    decode() { return this._carga ? this._carga.then(() => undefined) : Promise.reject(new Error('Imagen sin src')); }
    addEventListener(tipo, fn) { if (tipo === 'load') this.onload = fn; else if (tipo === 'error') this.onerror = fn; }
    removeEventListener() {}
    setAttribute(k, v) { if (k === 'src') this.src = v; }
  }
  self.Image = ImagenHilo;
  const desenvolver = (x) => (x instanceof ImagenHilo ? x._mapa : x);
  const pCtx = OffscreenCanvasRenderingContext2D.prototype;
  const dibujarOriginal = pCtx.drawImage, patronOriginal = pCtx.createPattern, mapaOriginal = self.createImageBitmap;
  pCtx.drawImage = function (fuente, ...resto) { return dibujarOriginal.call(this, desenvolver(fuente), ...resto); };
  pCtx.createPattern = function (fuente, repeticion) { return patronOriginal.call(this, desenvolver(fuente), repeticion); };
  self.createImageBitmap = function (fuente, ...resto) { return mapaOriginal.call(self, desenvolver(fuente), ...resto); };

  // ---------- document ----------
  const nodo = (tag) => ({
    tagName: tag.toUpperCase(), style: {}, dataset: {}, children: [],
    setAttribute(k, v) { this[k] = v; }, getAttribute(k) { return this[k] ?? null; },
    appendChild(x) { this.children.push(x); return x; }, remove() {}, addEventListener() {}, removeEventListener() {},
  });
  const head = nodo('head');
  head.appendChild = function (x) {
    if (x && x.tagName === 'SCRIPT') {
      setTimeout(() => {
        try { importar(x.src); } catch (error) { if (x.onerror) x.onerror(error); return; }
        if (x.onload) x.onload({ type: 'load', target: x });
      });
    }
    return x;
  };
  self.document = {
    get currentScript() { return actual; },
    get scripts() { return cargados; },
    head, body: nodo('body'), documentElement: nodo('html'),
    fonts: self.fonts,
    createElement(tag) {
      const t = String(tag).toLowerCase();
      if (t === 'canvas') return new OffscreenCanvas(300, 150);
      if (t === 'img') return new ImagenHilo();
      return nodo(t);
    },
    addEventListener() {}, removeEventListener() {},
  };

  // ---------- pasos ----------
  const preparados = new Map();
  const mensajeDe = (e) => (e && typeof e.message === 'string' && e.message) || (e && e.type ? 'error: ' + e.type : String(e));

  self.onmessage = async (evento) => {
    const d = evento.data || {};
    if (d.tipo === 'iniciar') {
      idioma = d.idioma === 'en' ? 'en' : 'es';
      document.documentElement.lang = idioma;
      try {
        d.scripts.forEach(importar);
        postMessage({ tipo: 'listo' });
      } catch (error) {
        postMessage({ tipo: 'error', mensaje: mensajeDe(error) });
      }
      return;
    }
    if (d.tipo !== 'paso') return;
    let enPreparar = true;
    try {
      const m = self.Flujos && Flujos.modulo(d.modulo);
      if (!m || !m.hilo) throw new Error('Módulo sin hilo: ' + d.modulo);
      if (!preparados.has(m.id)) preparados.set(m.id, Promise.resolve().then(() => m.preparar && m.preparar()));
      await preparados.get(m.id);
      enPreparar = false;
      const r = await m.ejecutar(d.ficha, d.opciones, d.contexto);
      postMessage({ tipo: 'hecho', id: d.id, r: r === undefined ? null : r, vacio: r === undefined });
    } catch (error) {
      postMessage({ tipo: 'fallo', id: d.id, enPreparar, nombre: (error && error.name) || 'Error', mensaje: mensajeDe(error) });
    }
  };
})();

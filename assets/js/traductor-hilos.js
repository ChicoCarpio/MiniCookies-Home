/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
/*
 * Traductor de documentos con trabajadores (Web Workers): la página solo orquesta y no se congela. Cada trabajador (traductor-hilo.js) tiene
 * su motor; los segmentos se reparten en una cola (el primero que se desocupa toma el siguiente). El resultado es el mismo que en un solo
 * hilo: cada segmento se traduce solo (lote 1) y con su contexto, que es el texto de origen anterior, no su traducción.
 *
 *   const t = await TraductorHilos.crear({ rutaOrt: '/vendor/onnxruntime-web/1.20.1/', rutaHilo: '/assets/js/traductor-hilo.js',
 *                                          rutaPiezas: '/assets/js/traductor-piezas.js', rutaMotor: '/assets/js/traductor-motor.js',
 *                                          vocabulario: '/assets/modelos/…json', codificador: '/assets/modelos/…onnx', decodificador: '…onnx',
 *                                          hilos: 4, alCargar: (bytes, total) => {} })
 *   const textos = await t.traducir(segmentos, 'es', { haz: 4, progreso: (hechos, total) => {}, cancelado: () => false })
 *        segmentos: [{ texto, contexto? }]; un segmento vacío o solo con espacios vuelve igual sin pasar por el modelo.
 *   t.cerrar()
 */
(function (global) {
  'use strict';

  function Trabajador(url) {
    this.w = new Worker(url);
    this.siguiente = 1;
    this.pendientes = new Map();
    this.w.onmessage = (e) => {
      const m = e.data, p = this.pendientes.get(m.id);
      if (!p) return;
      this.pendientes.delete(m.id);
      if (m.error) p.rej(new Error(m.error)); else p.res(m);
    };
    this.w.onerror = (e) => {
      for (const p of this.pendientes.values()) p.rej(new Error(e.message || 'error en el trabajador del traductor'));
      this.pendientes.clear();
    };
  }
  Trabajador.prototype.pedir = function (mensaje, transferir) {
    return new Promise((res, rej) => {
      const id = this.siguiente++;
      this.pendientes.set(id, { res, rej });
      this.w.postMessage(Object.assign({ id }, mensaje), transferir || []);
    });
  };

  // Baja un archivo informando el avance (bytes recibidos) para la barra de carga del modelo.
  // `url` puede ser una lista de trozos del mismo archivo (Cloudflare Pages no admite archivos de más de 25 MiB): se bajan en orden y se juntan.
  async function bajar(url, alAvanzar) {
    if (Array.isArray(url)) {
      const partes = [];
      for (const u of url) partes.push(await bajar(u, alAvanzar));
      const datos = new Uint8Array(partes.reduce((s, p) => s + p.length, 0));
      let k = 0;
      for (const p of partes) { datos.set(p, k); k += p.length; }
      return datos;
    }
    const r = await fetch(url);
    if (!r.ok) throw new Error(`No se pudo bajar ${url} (${r.status})`);
    const total = Number(r.headers.get('content-length')) || 0;
    if (!r.body || !alAvanzar) return new Uint8Array(await r.arrayBuffer());
    const lector = r.body.getReader();
    const trozos = [];
    let recibidos = 0;
    for (;;) {
      const { done, value } = await lector.read();
      if (done) break;
      trozos.push(value); recibidos += value.length;
      alAvanzar(value.length, total);
    }
    const datos = new Uint8Array(recibidos);
    let k = 0;
    for (const t of trozos) { datos.set(t, k); k += t.length; }
    return datos;
  }

  async function crear(op) {
    const hilos = op.hilos || Math.max(1, Math.min(4, (navigator.hardwareConcurrency || 4) - 1));
    const direccion = `${op.rutaHilo}?ort=${encodeURIComponent(op.rutaOrt)}&piezas=${encodeURIComponent(op.rutaPiezas)}&motor=${encodeURIComponent(op.rutaMotor)}`;
    const t0 = performance.now();
    let llevados = 0;
    const avance = op.alCargar ? (n) => { llevados += n; op.alCargar(llevados); } : null;
    const [vocabulario, codificador, decodificador] = await Promise.all([
      fetch(op.vocabulario).then((r) => { if (!r.ok) throw new Error(`No se pudo bajar ${op.vocabulario}`); return r.json(); }),
      bajar(op.codificador, avance),
      bajar(op.decodificador, avance),
    ]);
    const trabajadores = Array.from({ length: hilos }, () => new Trabajador(direccion));
    // Cada trabajador recibe una copia de los modelos (structured clone): la página no transfiere los suyos para poder dárselos a todos.
    await Promise.all(trabajadores.map((t) => t.pedir({ tipo: 'cargar', vocabulario, codificador, decodificador })));
    const cargaMs = performance.now() - t0;
    const libres = trabajadores.slice();
    const cola = [];

    function despachar() {
      while (libres.length && cola.length) {
        const t = libres.pop(), tarea = cola.shift();
        if (tarea.cancelado && tarea.cancelado()) { libres.push(t); tarea.rej(new Error('cancelado')); continue; }
        t.pedir(tarea.mensaje).then(
          (m) => { libres.push(t); tarea.res(m.texto); despachar(); },
          (e) => { libres.push(t); tarea.rej(e); despachar(); });
      }
    }

    function traducirUno(segmento, destino, haz, alfa, cancelado) {
      return new Promise((res, rej) => {
        cola.push({ mensaje: { tipo: 'traducir', segmento, destino, haz, alfa }, res, rej, cancelado });
        despachar();
      });
    }

    async function traducir(segmentos, destino, { haz = 4, alfa = 0.6, progreso, cancelado } = {}) {
      let hechos = 0;
      const total = segmentos.length;
      return Promise.all(segmentos.map(async (s) => {
        const texto = !s.texto || !s.texto.trim() ? s.texto : await traducirUno(s, destino, haz, alfa, cancelado);
        hechos++;
        if (progreso) progreso(hechos, total);
        return texto;
      }));
    }

    function cerrar() { trabajadores.forEach((t) => t.w.terminate()); }
    return { traducir, cerrar, hilos, cargaMs };
  }

  global.TraductorHilos = { crear };
})(typeof window !== 'undefined' ? window : globalThis);

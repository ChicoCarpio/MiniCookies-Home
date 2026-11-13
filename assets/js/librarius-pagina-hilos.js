/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
/*
 * Librarius página con trabajadores (Web Workers): la misma tubería que librarius-pagina.js, pero la página principal solo orquesta y no se congela.
 *   - un trabajador DETECTOR: Versus, componentes, ejes, recortes enderezados, giro de la página;
 *   - N trabajadores LECTORES: cada uno con su sesión de Librarius; los recortes se reparten en una cola (el que se desocupa toma el siguiente).
 * El detector recorta por trozos de 16 líneas mientras los lectores ya leen los anteriores. Resultado idéntico al de un solo hilo.
 *
 *   const lector = await LibrariusPaginaHilos.crear({ rutaOrt: '/ort/', rutaPagina: '/js/librarius-pagina.js', rutaHilo: '/js/librarius-hilo.js',
 *                                                   modeloVersus: '/modelo/versus-phone-v1.onnx', modeloLibrarius: '/modelo/librarius-phone-v1-int8.onnx', alfabeto, hilos: 4 });
 *   const r = await lector.leer(gris, ancho, alto);   // { texto, lineas, giro, tiempos }
 *   lector.cerrar();
 * Requiere que librarius-pagina.js esté cargado en la página (usa sus funciones de geometría y de orden).
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
    this.w.onerror = (e) => { for (const p of this.pendientes.values()) p.rej(new Error(e.message || 'error en el trabajador')); this.pendientes.clear(); };
  }
  Trabajador.prototype.pedir = function (mensaje, transferir) {
    return new Promise((res, rej) => {
      const id = this.siguiente++;
      this.pendientes.set(id, { res, rej });
      this.w.postMessage(Object.assign({ id }, mensaje), transferir || []);
    });
  };

  // Cola de recortes por leer repartida entre los lectores: el primero que se desocupa toma el siguiente.
  function Lectores(trabajadores) { this.libres = trabajadores.slice(); this.cola = []; }
  Lectores.prototype.leer = function (recorte) {
    return new Promise((res, rej) => { this.cola.push({ recorte, res, rej }); this.despachar(); });
  };
  Lectores.prototype.despachar = function () {
    while (this.libres.length && this.cola.length) {
      const t = this.libres.pop(), tarea = this.cola.shift();
      t.pedir({ tipo: 'leer', px: tarea.recorte.px, ancho: tarea.recorte.ancho }, [tarea.recorte.px.buffer]).then(
        (m) => { this.libres.push(t); tarea.res(m); this.despachar(); },
        (e) => { this.libres.push(t); tarea.rej(e); this.despachar(); });
    }
  };

  async function crear(op) {
    const N = global.LibrariusPagina._nucleo;
    const hilos = op.hilos || Math.max(1, Math.min(4, (navigator.hardwareConcurrency || 4) - 1));
    const opciones = Object.assign({}, N.DEFECTOS, op.opciones);
    const direccion = (rol, modelo) => `${op.rutaHilo}?rol=${rol}&ort=${encodeURIComponent(op.rutaOrt)}&pag=${encodeURIComponent(op.rutaPagina)}&modelo=${encodeURIComponent(modelo)}`;
    const t0 = performance.now();
    const detector = new Trabajador(direccion('detector', op.modeloVersus));
    const trabajadores = Array.from({ length: hilos }, () => new Trabajador(direccion('lector', op.modeloLibrarius)));
    await Promise.all([detector.pedir({ tipo: 'cargar', opciones, alfabeto: op.alfabeto }), ...trabajadores.map((t) => t.pedir({ tipo: 'cargar', opciones, alfabeto: op.alfabeto }))]);
    const cargaMs = performance.now() - t0;
    const lectores = new Lectores(trabajadores);
    const TROZO = 16;

    async function pasada(gris, w, h, orientar, eje, tiempos) {
      let t = performance.now();
      const det = await detector.pedir(gris ? { tipo: 'detectar', gris, w, h, eje } : { tipo: 'detectar', eje }, gris ? [gris.buffer] : []);
      tiempos.detectar = (tiempos.detectar || 0) + performance.now() - t;
      const lineas = det.lineas;
      if (!lineas.length) return { texto: '', lineas: [], giro: 0, invertida: false };
      let invertir = false;
      if (orientar) {
        t = performance.now();
        const indices = N.muestraOrientacion(lineas, opciones.kOrientacion);
        const [c0, c1] = [await detector.pedir({ tipo: 'recortes', indices, invertir: false }), await detector.pedir({ tipo: 'recortes', indices, invertir: true })];
        const [l0, l1] = await Promise.all([Promise.all(c0.crops.map((c) => lectores.leer(c))), Promise.all(c1.crops.map((c) => lectores.leer(c)))]);
        invertir = N.votoInvertir(l0.map((x) => x.seguridad), l1.map((x) => x.seguridad));
        tiempos.orientacion = (tiempos.orientacion || 0) + performance.now() - t;
      }
      t = performance.now();
      const lecturas = new Array(lineas.length);
      for (let i = 0; i < lineas.length; i += TROZO) {           // el detector recorta un trozo mientras los lectores leen los anteriores
        const indices = []; for (let j = i; j < Math.min(lineas.length, i + TROZO); j++) indices.push(j);
        const c = await detector.pedir({ tipo: 'recortes', indices, invertir });
        c.crops.forEach((crop, j) => { lecturas[indices[j]] = lectores.leer(crop); });
      }
      const leidas = await Promise.all(lecturas);
      tiempos.lectura = (tiempos.lectura || 0) + performance.now() - t;
      lineas.forEach((l, i) => { l.texto = leidas[i].texto; l.seguridad = leidas[i].seguridad; if (invertir) l.u = [-l.u[0], -l.u[1]]; });
      const giro = N.giroDeLineas(lineas);
      const orden = N.ordenarLineas(lineas, giro).map((i) => lineas[i]);
      return { texto: orden.map((l) => l.texto).filter((x) => x.trim()).join('\n'), lineas: orden, giro: giro * 180 / Math.PI, invertida: invertir };
    }

    // Una página por vez: el detector guarda la página y las líneas que está leyendo; dos lecturas simultáneas se pisarían.
    let fila = Promise.resolve();
    function leer(gris, w, h, o) {
      const turno = fila.then(() => leerUna(gris, w, h, o));
      fila = turno.catch(() => {});
      return turno;
    }

    async function leerUna(gris, w, h, o) {
      const tiempos = {}, inicio = performance.now();
      const orientar = !o || o.orientar !== false;
      const r = await pasada(gris.slice(), w, h, orientar, null, tiempos);       // se manda una copia: el buffer viaja al trabajador
      let resultado = r;
      if (orientar && r.lineas.length) {
        const k = N.cuartosDeVuelta(r.giro);
        if (k) {
          await detector.pedir({ tipo: 'girar', k });
          const r2 = await pasada(null, 0, 0, false, [1, 0], tiempos);
          if (r2.lineas.length) { r2.giroPrimeraPasada = r.giro; resultado = r2; }
        }
      }
      tiempos.total = performance.now() - inicio;
      resultado.tiempos = tiempos;
      return resultado;
    }

    function cerrar() { detector.w.terminate(); trabajadores.forEach((t) => t.w.terminate()); }
    return { leer, cerrar, hilos, cargaMs };
  }

  global.LibrariusPaginaHilos = { crear };
})(typeof window !== 'undefined' ? window : globalThis);

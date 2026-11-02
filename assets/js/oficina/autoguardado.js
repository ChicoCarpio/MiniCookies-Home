/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Autoguardado local de los editores de oficina (núcleo de la suite, fase 1 del PLAN). El documento se guarda en el navegador (IndexedDB) para
// no perder trabajo si se cierra la pestaña: nunca sale de la computadora. Lo que se guarda es el archivo completo ya escrito por el editor
// (los mismos bytes que se descargarían), así recuperar es abrirlo otra vez.
//
// Uso:
//   const ag = Autoguardado.crear('hojas');
//   await ag.guardar({ id, nombre, bytes })          // id: uno por documento abierto (Autoguardado.nuevoId())
//   await ag.listar() → [{ id, nombre, fecha, tamano }] (lo más nuevo primero)
//   await ag.leer(id) → { id, nombre, fecha, bytes } | null
//   await ag.borrar(id)
//   ag.programar(fnQueDevuelveDatos, ms)            // guarda un rato después del último cambio (se llama en cada cambio)
// Si el navegador no deja usar IndexedDB (ventana privada, sin espacio), todo devuelve vacío sin lanzar: el editor sigue funcionando.
(function () {
  const BASE = 'mc-oficina', VERSION = 1, ALMACEN = 'borradores';
  const MAX_BORRADORES = 10;
  let conexion = null;

  function abrirBase() {
    if (conexion) return conexion;
    conexion = new Promise((resolver) => {
      try {
        if (typeof indexedDB === 'undefined') { resolver(null); return; }
        const pedido = indexedDB.open(BASE, VERSION);
        pedido.onupgradeneeded = () => {
          const db = pedido.result;
          if (!db.objectStoreNames.contains(ALMACEN)) db.createObjectStore(ALMACEN, { keyPath: 'id' }).createIndex('app', 'app');
        };
        pedido.onsuccess = () => resolver(pedido.result);
        pedido.onerror = () => resolver(null);
        pedido.onblocked = () => resolver(null);
      } catch (e) { resolver(null); }
    });
    return conexion;
  }
  function operar(modo, fn) {
    return abrirBase().then((db) => new Promise((resolver) => {
      if (!db) { resolver(null); return; }
      try {
        const tx = db.transaction(ALMACEN, modo);
        const almacen = tx.objectStore(ALMACEN);
        let resultado = null;
        const r = fn(almacen);
        if (r) r.onsuccess = () => { resultado = r.result; };
        tx.oncomplete = () => resolver(resultado);
        tx.onerror = tx.onabort = () => resolver(null);
      } catch (e) { resolver(null); }
    }));
  }

  const nuevoId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);

  function crear(app) {
    let temporizador = 0, enCurso = Promise.resolve();
    async function listar() {
      const todos = (await operar('readonly', (a) => a.index('app').getAll(app))) || [];
      return todos.map((x) => ({ id: x.id, nombre: x.nombre, fecha: x.fecha, tamano: x.bytes ? x.bytes.byteLength : 0 })).sort((a, b) => b.fecha - a.fecha);
    }
    async function guardar(datos) {
      const registro = { id: datos.id, app, nombre: datos.nombre, fecha: Date.now(), bytes: datos.bytes };
      await operar('readwrite', (a) => a.put(registro));
      const lista = await listar();   // solo los más recientes: el espacio del navegador es de todos los sitios
      for (const viejo of lista.slice(MAX_BORRADORES)) await borrar(viejo.id);
      return registro.fecha;
    }
    const leer = (id) => operar('readonly', (a) => a.get(id));
    const borrar = (id) => operar('readwrite', (a) => a.delete(id));
    function programar(obtenerDatos, ms, alGuardar) {
      clearTimeout(temporizador);
      temporizador = setTimeout(() => {
        enCurso = enCurso.then(async () => {
          try {
            const datos = await obtenerDatos();
            if (datos) { const fecha = await guardar(datos); if (alGuardar) alGuardar(fecha); }
          } catch (e) { if (alGuardar) alGuardar(null, e); }
        });
      }, ms === undefined ? 2000 : ms);
    }
    const cancelar = () => clearTimeout(temporizador);
    return { listar, guardar, leer, borrar, programar, cancelar, esperar: () => enCurso };
  }

  globalThis.Autoguardado = { crear, nuevoId };
})();

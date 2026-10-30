/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Deshacer y rehacer de los editores de oficina (núcleo de la suite, fase 1 del PLAN). Sin DOM: corre igual en Node.
//
// A diferencia del historial de los editores de imagen (que guarda fotos del estado), aquí cada cambio es una OPERACIÓN que sabe deshacerse:
// el documento aplica la operación y devuelve la operación inversa. Una transacción es una lista de operaciones que se deshace de una vez
// (p. ej. pegar 200 celdas). Las operaciones son objetos simples (JSON), así se pueden probar y guardar.
//
// Uso:
//   const h = HistorialDoc.crear({ aplicar: (op) => inversa, alCambiar: () => …, limite: 500 });
//   h.hacer('Escribir', [op1, op2], { fusionar: 'celda:A1' })   // aplica y registra; con la misma clave de fusión y < 1 s, se suma a la anterior
//   h.deshacer() / h.rehacer()                                  // true si había algo
//   h.puedeDeshacer() / h.puedeRehacer() / h.nombreDeshacer() / h.nombreRehacer()
//   h.marcarGuardado() / h.sucio()                              // ¿cambió algo desde el último guardado?
//   h.limpiar()
// `aplicar` debe devolver la inversa de la operación aplicada (o null si la operación no cambió nada: no se registra).
(function () {
  const VENTANA_FUSION = 1000;

  function crear(opciones) {
    const aplicar = opciones.aplicar;
    const alCambiar = opciones.alCambiar || (() => {});
    const limite = opciones.limite || 500;
    const reloj = opciones.reloj || (() => Date.now());
    let hechas = [];      // [{ nombre, adelante: [op], atras: [op], clave, hora, version }]
    let deshechas = [];
    let version = 0, versionGuardada = 0, contador = 0;

    function correr(ops) {
      const inversas = [];
      for (const op of ops) {
        const inv = aplicar(op);
        if (inv) inversas.push(inv);
      }
      return inversas.reverse();   // se deshace en orden contrario
    }

    function hacer(nombre, ops, extra) {
      if (!ops || !ops.length) return false;
      const atras = correr(ops);
      if (!atras.length) return false;
      const clave = extra && extra.fusionar;
      const ahora = reloj();
      const ultima = hechas[hechas.length - 1];
      deshechas = [];
      if (clave && ultima && ultima.clave === clave && ahora - ultima.hora < VENTANA_FUSION && ultima.version === version) {
        ultima.adelante.push(...ops);
        ultima.atras = atras.concat(ultima.atras);
        ultima.hora = ahora;
        version = ultima.version = ++contador;
      } else {
        version = ++contador;
        hechas.push({ nombre, adelante: ops.slice(), atras, clave: clave || null, hora: ahora, version });
        if (hechas.length > limite) hechas.shift();
      }
      alCambiar();
      return true;
    }

    function deshacer() {
      const t = hechas.pop();
      if (!t) return false;
      correr(t.atras);
      deshechas.push(t);
      version = hechas.length ? hechas[hechas.length - 1].version : 0;
      alCambiar();
      return true;
    }

    function rehacer() {
      const t = deshechas.pop();
      if (!t) return false;
      t.atras = correr(t.adelante);
      t.clave = null;   // lo rehecho no se fusiona con lo que se escriba después
      hechas.push(t);
      version = t.version;
      alCambiar();
      return true;
    }

    return {
      hacer, deshacer, rehacer,
      puedeDeshacer: () => hechas.length > 0,
      puedeRehacer: () => deshechas.length > 0,
      nombreDeshacer: () => (hechas.length ? hechas[hechas.length - 1].nombre : null),
      nombreRehacer: () => (deshechas.length ? deshechas[deshechas.length - 1].nombre : null),
      marcarGuardado() { versionGuardada = version; alCambiar(); },
      sucio: () => version !== versionGuardada,
      cortarFusion() { const u = hechas[hechas.length - 1]; if (u) u.clave = null; },
      limpiar() { hechas = []; deshechas = []; version = versionGuardada = 0; alCambiar(); },
      get cantidad() { return hechas.length; },
    };
  }

  globalThis.HistorialDoc = { crear };
})();

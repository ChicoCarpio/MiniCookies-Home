/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Deshacer y rehacer de los editores de imagen (Editor avanzado, Editor de
// color y collage; plan del rediseño, hallazgos/editores-rediseno.md, fase 3).
//
// Cada editor dice solo dos cosas: cómo sacar una foto de su estado
// (capturar, con una "firma" que la resume) y cómo volver a una
// (restaurar). El módulo no cablea cada control: después de cada clic,
// cambio, toque o tecla compara la firma con la del último paso, y si cambió,
// ese es un paso nuevo. Mover un deslizador es UN paso (al soltarlo), no uno
// por cada movimiento.
//
// Los lienzos (pincel, borrador) se copian solo cuando cambiaron: el editor
// pide la copia con crearCopiador(), que reusa la anterior mientras la
// versión del lienzo sea la misma. Hay un tope de memoria: pasado el tope se
// olvidan los pasos más viejos.

window.Historial = (function () {
  'use strict';

  // Tope de la memoria que ocupan las copias de los lienzos guardados.
  const LIMITE_BYTES = (navigator.deviceMemory && navigator.deviceMemory <= 4) || window.matchMedia('(max-width: 900px)').matches
    ? 192e6 : 512e6;

  function copiar(lienzo, ancho, alto) {
    const c = document.createElement('canvas');
    c.width = ancho || lienzo.width; c.height = alto || lienzo.height;
    c.getContext('2d').drawImage(lienzo, 0, 0, c.width, c.height);
    return c;
  }

  // Copias de lienzos que cambian (el pincel de una capa, lo borrado…), por
  // clave: mientras el lienzo sea el mismo y su versión no cambie, devuelve la
  // copia ya hecha.
  function crearCopiador() {
    const ultimo = new Map(); // clave -> { version, lienzo, copia }
    return {
      copia(clave, version, lienzo) {
        if (!lienzo) return null;
        const u = ultimo.get(clave);
        if (u && u.version === version && u.lienzo === lienzo) return u.copia;
        const copia = copiar(lienzo);
        ultimo.set(clave, { version, lienzo, copia });
        return copia;
      },
      // Un lienzo nuevo (el editor lo va a modificar) con lo de "copia",
      // opcionalmente a otro tamaño. Queda anotado: sacarle una foto enseguida
      // no vuelve a copiarlo.
      restaurar(clave, version, copia, ancho, alto) {
        if (!copia) { ultimo.delete(clave); return null; }
        const nuevo = copiar(copia, ancho, alto);
        if (nuevo.width === copia.width && nuevo.height === copia.height) ultimo.set(clave, { version, lienzo: nuevo, copia });
        else ultimo.delete(clave);
        return nuevo;
      },
      olvidar() { ultimo.clear(); },
    };
  }

  // opciones: { capturar() -> { firma, lienzos? }, restaurar(paso), alCambiar(),
  //             limitePasos, limiteBytes }
  // "lienzos" (las copias que guarda el paso) es para medir la memoria.
  function crear(opciones) {
    const { capturar, restaurar, alCambiar = () => {}, limitePasos = 100, limiteBytes = LIMITE_BYTES } = opciones;
    let pasos = [], pos = -1, firmaActual = null, pendiente = null;

    function memoria() {
      const vistos = new Set();
      let total = 0;
      for (const paso of pasos) {
        for (const c of paso.lienzos || []) {
          if (c && !vistos.has(c)) { vistos.add(c); total += c.width * c.height * 4; }
        }
      }
      return total;
    }
    function recortar() {
      while (pasos.length > 2 && pos > 0 && (pasos.length > limitePasos || memoria() > limiteBytes)) {
        pasos.shift();
        pos -= 1;
      }
    }
    function cancelarPendiente() { if (pendiente) { clearTimeout(pendiente); pendiente = null; } }

    // Empieza de nuevo (otra imagen, otro proyecto): lo de ahora es el paso 0.
    function reiniciar() {
      cancelarPendiente();
      const paso = capturar();
      pasos = [paso]; pos = 0; firmaActual = paso.firma;
      alCambiar();
    }
    // Si algo cambió desde el último paso, lo guarda como uno nuevo (y
    // descarta lo que se podía rehacer).
    function comprobar() {
      cancelarPendiente();
      if (pos < 0) return false;
      const paso = capturar();
      if (paso.firma === firmaActual) return false;
      pasos = pasos.slice(0, pos + 1);
      pasos.push(paso);
      pos += 1;
      firmaActual = paso.firma;
      recortar();
      alCambiar();
      return true;
    }
    function programar(ms = 60) {
      if (pos < 0) return;
      cancelarPendiente();
      pendiente = setTimeout(comprobar, ms);
    }
    function ir(nueva) {
      pos = nueva;
      restaurar(pasos[pos]);
      // Restaurar puede renumerar versiones: la firma de "ahora" es la nueva.
      firmaActual = capturar().firma;
      alCambiar();
    }
    function deshacer() {
      comprobar(); // lo que estaba sin guardar es lo primero que se deshace
      if (pos <= 0) return false;
      ir(pos - 1);
      return true;
    }
    function rehacer() {
      comprobar();
      if (pos >= pasos.length - 1) return false;
      ir(pos + 1);
      return true;
    }
    return {
      reiniciar, comprobar, programar, deshacer, rehacer,
      puedeDeshacer: () => pos > 0,
      puedeRehacer: () => pos >= 0 && pos < pasos.length - 1,
      cantidad: () => pasos.length,
    };
  }

  // Ctrl+Z deshace; Ctrl+Shift+Z y Ctrl+Y rehacen. En un campo de texto no se
  // tocan: ahí deshacen lo escrito, como siempre (la regla es de atajos.js).
  function conectarTeclado(historial, activo) {
    document.addEventListener('keydown', (e) => {
      if (!(e.ctrlKey || e.metaKey) || e.altKey || !activo()) return;
      if (window.Atajos ? window.Atajos.enCampoDeTexto(e) : /^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName)) return;
      const tecla = e.key.toLowerCase();
      if (tecla === 'z' && !e.shiftKey) { e.preventDefault(); historial.deshacer(); }
      else if ((tecla === 'z' && e.shiftKey) || tecla === 'y') { e.preventDefault(); historial.rehacer(); }
    });
  }

  // Tocar con dos dedos a la vez (sin moverlos, como en ibisPaint) deshace.
  // Convive con el pellizco de la página: un pellizco mueve los dedos.
  function conectarToqueDosDedos(elemento, accion) {
    const toques = new Map();
    let inicio = 0, valido = false, maximo = 0;
    elemento.addEventListener('pointerdown', (e) => {
      if (e.pointerType !== 'touch') return;
      if (!toques.size) { inicio = e.timeStamp; valido = true; maximo = 0; }
      toques.set(e.pointerId, { x: e.clientX, y: e.clientY });
      maximo = Math.max(maximo, toques.size);
      if (toques.size > 2) valido = false;
    });
    elemento.addEventListener('pointermove', (e) => {
      const t = toques.get(e.pointerId);
      if (t && Math.hypot(e.clientX - t.x, e.clientY - t.y) > 12) valido = false;
    });
    const fin = (e) => {
      if (!toques.delete(e.pointerId) || toques.size) return;
      if (valido && maximo === 2 && e.timeStamp - inicio < 400) accion();
    };
    elemento.addEventListener('pointerup', fin);
    elemento.addEventListener('pointercancel', (e) => { toques.delete(e.pointerId); valido = false; });
  }

  // Todo junto, para un editor: el historial, los botones ↶ ↷ (se habilitan
  // solos), el teclado, el toque con dos dedos sobre "lienzo" y la vigilancia
  // de los cambios en toda la página. "activo()" dice si hay algo abierto.
  function conectar(opciones) {
    const { deshacer: botonDeshacer, rehacer: botonRehacer, lienzo, activo = () => true } = opciones;
    const historial = crear({
      ...opciones,
      alCambiar: () => {
        if (botonDeshacer) botonDeshacer.disabled = !historial.puedeDeshacer();
        if (botonRehacer) botonRehacer.disabled = !historial.puedeRehacer();
        if (opciones.alCambiar) opciones.alCambiar();
      },
    });
    if (botonDeshacer) botonDeshacer.addEventListener('click', () => historial.deshacer());
    if (botonRehacer) botonRehacer.addEventListener('click', () => historial.rehacer());
    conectarTeclado(historial, activo);
    if (lienzo) conectarToqueDosDedos(lienzo, () => { if (activo()) historial.deshacer(); });
    // Se mira en todo el documento: en celular los paneles viven en la hoja
    // de app-movil.js, fuera del contenedor del editor.
    const ahora = () => { if (activo()) historial.programar(); };
    for (const tipo of ['change', 'pointerup', 'click', 'drop']) document.addEventListener(tipo, ahora);
    document.addEventListener('keyup', (e) => { if (!e.ctrlKey && !e.metaKey && activo()) historial.programar(500); });
    document.addEventListener('wheel', () => { if (activo()) historial.programar(500); }, { passive: true });
    return historial;
  }

  return { crear, crearCopiador, conectar, conectarTeclado, conectarToqueDosDedos };
})();

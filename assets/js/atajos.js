/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Atajos de teclado de los editores de imagen (plan del rediseño, hallazgos.md, 4.1),
// con una sola regla para todos: nunca mientras se escribe en un campo de
// texto (ahí cada tecla es del texto), y la barra espaciadora nunca sobre un
// botón o un enlace (ahí los activa, y eso es del teclado de siempre).
//
// Atajos.conectar(mapa, activo): mapa = { 'f': fn, 'ctrl+0': fn,
// ' ': { abajo, arriba } }. Una función corre al apretar la tecla; un par
// { abajo, arriba }, mientras se la mantiene (ver el original, mover la vista).
// "activo()" dice si el editor está abierto.

window.Atajos = (function () {
  'use strict';

  const SIN_TEXTO = ['range', 'checkbox', 'radio', 'button', 'color', 'file', 'submit', 'reset'];

  function enCampoDeTexto(e) {
    const t = e.target;
    if (!t || !t.tagName) return false;
    if (t.isContentEditable || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT') return true;
    return t.tagName === 'INPUT' && !SIN_TEXTO.includes(t.type);
  }

  function claveDe(e) {
    const tecla = e.key === ' ' ? ' ' : e.key.toLowerCase();
    return ((e.ctrlKey || e.metaKey) ? 'ctrl+' : '') + (e.shiftKey && tecla.length > 1 ? 'shift+' : '') + tecla;
  }

  function conectar(mapa, activo = () => true) {
    const mantenidas = new Map(); // tecla -> { arriba }
    document.addEventListener('keydown', (e) => {
      if (e.altKey || !activo() || enCampoDeTexto(e)) return;
      const clave = claveDe(e);
      const accion = mapa[clave];
      if (!accion) return;
      if (clave === ' ' && e.target.closest && e.target.closest('button, a, summary, [role="button"]')) return;
      e.preventDefault();
      if (typeof accion === 'function') { if (!e.repeat) accion(e); return; }
      if (e.repeat || mantenidas.has(clave)) return;
      mantenidas.set(clave, accion);
      if (accion.abajo) accion.abajo(e);
    });
    const soltar = (clave, e) => {
      const accion = mantenidas.get(clave);
      if (!accion) return;
      mantenidas.delete(clave);
      if (accion.arriba) accion.arriba(e);
    };
    document.addEventListener('keyup', (e) => {
      // Al soltar importa la tecla, no los modificadores.
      soltar(e.key === ' ' ? ' ' : e.key.toLowerCase(), e);
    });
    // Si la ventana pierde el foco con una tecla apretada, se suelta.
    window.addEventListener('blur', () => { [...mantenidas.keys()].forEach((clave) => soltar(clave, null)); });
  }

  return { conectar, enCampoDeTexto };
})();

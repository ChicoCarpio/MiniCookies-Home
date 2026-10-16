/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Pantalla completa de los editores de imagen (plan del rediseño, hallazgos/editores-rediseno.md,
// fase 4): el botón ⛶ y la tecla F ponen el contenedor del editor en
// pantalla completa con la API del navegador; Esc (o el mismo botón) sale.
//
// En pantalla completa se ve SOLO lo que está adentro del contenedor. Por eso:
// - El anuncio va adentro, en una franja propia debajo del editor, de alto
//   fijo y sin nada encima (plan 4.4). En celular ya está adentro (la tira de
//   app-movil.js); en PC se mueve a la franja mientras dure y vuelve a su
//   lugar al salir, con AppMovil.moverSinRecargar (un anuncio que se mueve
//   de la forma común se recarga y queda en blanco).
// - Lo que se abre aparte (las burbujas "?" de ayuda-tip.js) se cuelga de
//   document.fullscreenElement en vez de <body>.
//
// Donde el navegador no la permite para un elemento cualquiera (iPhone) el
// botón no aparece.

window.PantallaCompleta = (function () {
  'use strict';

  function disponible() {
    return !!(document.fullscreenEnabled && Element.prototype.requestFullscreen);
  }

  // opciones: { contenedor, boton, activo() }
  function conectar(opciones) {
    const { contenedor, boton, activo = () => true } = opciones;
    if (!contenedor) return null;
    if (!disponible()) { if (boton) boton.hidden = true; return null; }

    let franja = null, origen = null; // de dónde salió el anuncio en PC

    const mover = (padre, el, ref) => (window.AppMovil ? window.AppMovil.moverSinRecargar(padre, el, ref) : padre.insertBefore(el, ref));
    const anuncio = () => document.querySelector('.anuncio');

    // El anuncio, adentro de lo que está en pantalla completa.
    function meterAnuncio() {
      const a = anuncio();
      if (!a || contenedor.contains(a)) return;
      if (!franja) {
        franja = document.createElement('div');
        franja.className = 'franja-anuncio-completa';
        contenedor.appendChild(franja);
      }
      if (!origen) origen = { padre: a.parentNode, siguiente: a.nextSibling };
      mover(franja, a, null);
    }
    function sacarAnuncio() {
      const a = anuncio();
      if (a && franja && franja.contains(a) && origen) mover(origen.padre, a, origen.siguiente);
      if (franja) franja.remove();
      franja = null; origen = null;
    }

    function enPantallaCompleta() { return document.fullscreenElement === contenedor; }

    function alternar() {
      if (enPantallaCompleta()) { document.exitFullscreen().catch(() => {}); return; }
      if (!activo()) return;
      contenedor.requestFullscreen().catch(() => {});
    }

    function actualizarBoton() {
      if (!boton) return;
      const dentro = enPantallaCompleta();
      boton.setAttribute('aria-pressed', String(dentro));
      boton.classList.toggle('activo', dentro);
      const clave = dentro ? 'title-salir-pantalla-completa' : 'title-pantalla-completa';
      const texto = typeof window.t === 'function' ? window.t(clave) : '';
      if (texto && texto !== clave) { boton.title = texto; boton.setAttribute('aria-label', texto); }
      boton.setAttribute('data-i18n-attr', 'title:' + clave + ';aria-label:' + clave);
    }

    document.addEventListener('fullscreenchange', () => {
      const dentro = enPantallaCompleta();
      contenedor.classList.toggle('pantalla-completa-activa', dentro);
      if (dentro) meterAnuncio(); else sacarAnuncio();
      actualizarBoton();
      // Los editores recalculan el lienzo con "resize" (el tamaño cambió).
      requestAnimationFrame(() => window.dispatchEvent(new Event('resize')));
    });
    // Al pasar de celular a PC (o al revés) en pantalla completa, app-movil
    // devuelve o se lleva el anuncio: se vuelve a poner adentro.
    window.addEventListener('resize', () => {
      if (enPantallaCompleta()) requestAnimationFrame(meterAnuncio);
    });

    if (boton) {
      boton.hidden = false;
      boton.addEventListener('click', alternar);
      actualizarBoton();
      document.addEventListener('cambioIdioma', actualizarBoton);
    }
    if (window.Atajos) window.Atajos.conectar({ f: alternar }, () => activo() || enPantallaCompleta());
    return { alternar, enPantallaCompleta };
  }

  // Dónde colgar algo que se abre por encima de todo.
  function capa() { return document.fullscreenElement || document.body; }

  return { disponible, conectar, capa };
})();

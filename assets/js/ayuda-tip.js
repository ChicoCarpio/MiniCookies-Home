/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Comportamiento del icono "?" con burbuja de ayuda (ver assets/css/ayuda-tip.css
// para el markup esperado y por qué la burbuja usa position:fixed).
//
// Al mostrarla, la burbuja se saca de donde esté (puede estar dentro de un
// acordeón angosto, una columna con scroll propio, etc.) y se cuelga directo
// de <body>, con su posición ya calculada en píxeles de pantalla — así queda
// SIEMPRE por encima de todo y nunca la recorta un "overflow: hidden" ni el
// borde de un panel. Al cerrarla vuelve a su lugar original en el DOM (no se
// pierde su posición ni sus escuchadores).
//
// Con mouse de verdad se muestra al pasar el cursor (y al enfocarla con
// teclado); en pantallas táctiles, sin hover real, se muestra con un toque y
// se oculta con el siguiente — por eso NO hay una regla CSS de ":hover" para
// esto: mostrarla por CSS la dejaría sin la posición que calcula esta función,
// y aparecería pegada a su lugar original en vez de arriba/abajo del botón.
(function () {
  var contador = 0;
  var abierta = null; // { tip, boton, burbuja, padre, siguiente }
  var esMouse = window.matchMedia('(hover: hover) and (pointer: fine)').matches;

  function colocarBurbuja(boton, burbuja) {
    var r = boton.getBoundingClientRect();
    var centroX = r.left + r.width / 2;
    var margen = 8;
    var medioAncho = 130; // mitad de los 260px de max-width (ver CSS)
    centroX = Math.min(Math.max(centroX, medioAncho + margen), window.innerWidth - medioAncho - margen);

    // Por defecto arriba del botón; si no entra (queda muy arriba de la
    // pantalla), se muestra abajo en su lugar.
    var vaArriba = r.top > 140;
    burbuja.classList.toggle('abajo', !vaArriba);
    burbuja.style.left = centroX + 'px';
    if (vaArriba) {
      burbuja.style.bottom = (window.innerHeight - r.top + 9) + 'px';
      burbuja.style.top = 'auto';
    } else {
      burbuja.style.top = (r.bottom + 9) + 'px';
      burbuja.style.bottom = 'auto';
    }
  }

  function cerrar() {
    if (!abierta) return;
    abierta.tip.classList.remove('abierto');
    abierta.burbuja.classList.remove('visible');
    // Vuelve exactamente a donde estaba, para no perder su lugar en el DOM.
    abierta.padre.insertBefore(abierta.burbuja, abierta.siguiente);
    abierta = null;
  }

  // Idempotente: si ya está abierta, solo recalcula su posición (sirve para
  // hover/foco, donde no tiene sentido que un segundo evento la cierre).
  function mostrar(tip, boton, burbuja) {
    if (abierta && abierta.burbuja === burbuja) { colocarBurbuja(boton, burbuja); return; }
    cerrar();
    abierta = { tip: tip, boton: boton, burbuja: burbuja, padre: burbuja.parentNode, siguiente: burbuja.nextSibling };
    // En pantalla completa solo se ve lo que está adentro del elemento.
    (document.fullscreenElement || document.body).appendChild(burbuja);
    colocarBurbuja(boton, burbuja);
    burbuja.classList.add('visible');
    tip.classList.add('abierto');
  }

  // Para el toque: si ya está abierta, la cierra; si no, la abre.
  function alternar(tip, boton, burbuja) {
    if (abierta && abierta.burbuja === burbuja) { cerrar(); return; }
    mostrar(tip, boton, burbuja);
  }

  function iniciar(tip) {
    var boton = tip.querySelector('.ayuda-boton');
    var burbuja = tip.querySelector('.ayuda-burbuja');
    if (!boton || !burbuja) return;
    if (!burbuja.id) burbuja.id = 'ayuda-burbuja-' + (++contador);
    boton.setAttribute('aria-describedby', burbuja.id);
    burbuja.setAttribute('role', 'tooltip');

    // El foco por teclado (Tab) cuenta en cualquier dispositivo. El que llega
    // por un toque o un clic, no: en pantallas táctiles el toque primero da el
    // foco (y la abría) y enseguida el clic la alternaba, o sea, la cerraba.
    // Pasaba en el emulador de celular de Edge y en Chrome de Android.
    var porPuntero = false;
    boton.addEventListener('pointerdown', function () { porPuntero = true; });
    boton.addEventListener('focus', function () {
      if (porPuntero) { porPuntero = false; return; }
      mostrar(tip, boton, burbuja);
    });
    boton.addEventListener('click', function () { porPuntero = false; });
    boton.addEventListener('blur', cerrar);

    if (esMouse) {
      tip.addEventListener('mouseenter', function () { mostrar(tip, boton, burbuja); });
      tip.addEventListener('mouseleave', cerrar);
      // Sin esto, un click sobre el botón (algo que puede pasar sin querer
      // mientras ya está mostrada por hover) burbujeaba hasta el listener de
      // document y la cerraba al toque, aunque el mouse siguiera encima.
      boton.addEventListener('click', function (e) { e.stopPropagation(); });
    } else {
      boton.addEventListener('click', function (e) {
        e.stopPropagation();
        alternar(tip, boton, burbuja);
      });
    }
  }

  document.addEventListener('click', cerrar);
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape') cerrar(); });
  // Si la página scrollea (la ventana o un panel lateral con su propio
  // scroll), la burbuja ya fijada en pantalla dejaría de apuntar al botón:
  // más simple cerrarla que recalcular su posición en cada scroll.
  window.addEventListener('scroll', cerrar, true);
  window.addEventListener('resize', cerrar);

  document.querySelectorAll('.ayuda-tip').forEach(iniciar);
  // para los "?" que una pieza crea después de cargada la página (la rejilla de hoja)
  window.AyudaTipIniciar = iniciar;
})();

/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Reordenar elementos arrastrándolos, compartido por las herramientas que
// tienen listas o cuadrículas reordenables (páginas, hojas, capítulos...).
//
// El arrastre HTML5 (draggable / dragover) no funciona con el dedo en la
// mayoría de los celulares: acá se suma el arrastre táctil. Para no pelear
// con el desplazamiento normal de la página, un elemento se "levanta"
// manteniéndolo apretado un momento (~0,35 s); si el dedo se mueve antes,
// es un desplazamiento común.

// Inserta "arrastrando" antes o después de "objetivo" según dónde está el
// puntero: en una cuadrícula (el objetivo comparte fila con un vecino) cuenta
// la posición horizontal; en una lista (un elemento por fila), la vertical.
function insertarSegunPosicion(arrastrando, objetivo, x, y) {
  const rect = objetivo.getBoundingClientRect();
  const mismaFila = (vecino) => {
    if (!vecino || vecino === arrastrando) return false;
    const r = vecino.getBoundingClientRect();
    return Math.abs(r.top - rect.top) < rect.height / 2;
  };
  const estilo = getComputedStyle(objetivo.parentNode);
  const filaFlex = estilo.display.includes('flex') && estilo.flexDirection.startsWith('row');
  const enCuadricula = filaFlex || mismaFila(objetivo.previousElementSibling) || mismaFila(objetivo.nextElementSibling);
  const despues = enCuadricula ? x > rect.left + rect.width / 2 : y > rect.top + rect.height / 2;
  const referencia = despues ? objetivo.nextSibling : objetivo;
  if (referencia !== arrastrando && arrastrando.nextSibling !== referencia) {
    objetivo.parentNode.insertBefore(arrastrando, referencia);
    return true;
  }
  return false;
}

function habilitarReordenTactil(contenedor, selectorItem, alReordenar) {
  const ESPERA_MS = 350;
  const TOLERANCIA_PX = 10;
  let item = null;
  let temporizador = null;
  let inicio = null;
  let activo = false;
  let ultimoToque = 0;

  function cancelar() {
    clearTimeout(temporizador);
    temporizador = null;
    if (item) item.classList.remove('arrastrando');
    item = null;
    activo = false;
  }

  contenedor.addEventListener('touchstart', (e) => {
    ultimoToque = Date.now();
    if (e.touches.length !== 1) { cancelar(); return; }
    const objetivo = e.target.closest(selectorItem);
    if (!objetivo || !contenedor.contains(objetivo)) return;
    if (e.target.closest('button, input, select, textarea, a, label')) return;
    item = objetivo;
    inicio = { x: e.touches[0].clientX, y: e.touches[0].clientY };
    temporizador = setTimeout(() => {
      activo = true;
      item.classList.add('arrastrando');
      if (navigator.vibrate) navigator.vibrate(20);
    }, ESPERA_MS);
  }, { passive: true });

  contenedor.addEventListener('touchmove', (e) => {
    if (!item) return;
    const toque = e.touches[0];
    if (!activo) {
      if (Math.hypot(toque.clientX - inicio.x, toque.clientY - inicio.y) > TOLERANCIA_PX) cancelar();
      return;
    }
    e.preventDefault(); // mientras se arrastra, la página no se desplaza
    const debajo = document.elementFromPoint(toque.clientX, toque.clientY);
    const objetivo = debajo && debajo.closest(selectorItem);
    if (objetivo && objetivo !== item && objetivo.parentNode === item.parentNode) {
      insertarSegunPosicion(item, objetivo, toque.clientX, toque.clientY);
    }
    // Cerca del borde de la pantalla, la página se desplaza sola.
    if (toque.clientY < 60) window.scrollBy(0, -12);
    else if (toque.clientY > window.innerHeight - 60) window.scrollBy(0, 12);
  }, { passive: false });

  function terminar() {
    const huboArrastre = activo;
    cancelar();
    if (huboArrastre) alReordenar();
  }
  contenedor.addEventListener('touchend', terminar);
  contenedor.addEventListener('touchcancel', terminar);

  // Mantener apretado abre el menú del sistema (guardar imagen, etc.) y
  // algunos navegadores arrancan su propio arrastre: se evitan ambos.
  contenedor.addEventListener('contextmenu', (e) => {
    if (item) e.preventDefault();
  });
  contenedor.addEventListener('dragstart', (e) => {
    if (Date.now() - ultimoToque < 1500) e.preventDefault();
  }, true);
}

// Genera el fondo decorativo del sitio: unos pocos íconos sueltos, cada uno
// en una celda de una grilla floja con una posición/tamaño/rotación al azar
// dentro de esa celda (no una grilla pareja), salteando celdas al azar para
// que quede espaciado y no saturado. Requiere un <div id="fondoDecorativo">
// en el <body> y assets/css/fondo-decorativo.css cargado.
(function () {
  // Resuelve las rutas de los íconos relativas a este mismo script, así
  // funciona igual desde la raíz del sitio que desde herramientas/*.html.
  const scriptActual = document.currentScript;
  const baseIconos = new URL('../img/fondo/', scriptActual.src).href;

  const nombresIconos = [
    'icon-perro-transparente.png', 'hombre-de-gengibre.png', 'galletas.png', 'icon-otono-transparente.png', 'base-de-datos.png',
    'bombilla.png', 'carta.png', 'checks.png', 'code.png', 'engranaje.png',
    'folder.png', 'hoja.png', 'lupa.png', 'minicookies-logo.png',
  ];
  const iconos = nombresIconos.map((nombre) => baseIconos + nombre);

  function generar() {
    const contenedor = document.getElementById('fondoDecorativo');
    if (!contenedor) return;
    contenedor.innerHTML = '';

    const anchoCelda = 240;
    const altoCelda = 220;
    const probabilidadPorCelda = 0.7; // deja algún hueco para que no se sienta en grilla pareja
    const columnas = Math.ceil(window.innerWidth / anchoCelda);
    const filas = Math.ceil(window.innerHeight / altoCelda);

    for (let f = 0; f < filas; f++) {
      for (let c = 0; c < columnas; c++) {
        if (Math.random() > probabilidadPorCelda) continue;

        const tamano = 70 + Math.random() * 70; // 70–140px
        const margen = tamano / 2 + 10;
        const x = c * anchoCelda + margen + Math.random() * (anchoCelda - margen * 2);
        const y = f * altoCelda + margen + Math.random() * (altoCelda - margen * 2);
        const rotacion = Math.round(Math.random() * 40 - 20); // -20° a 20°

        const img = document.createElement('img');
        img.src = iconos[Math.floor(Math.random() * iconos.length)];
        img.alt = '';
        img.style.width = `${tamano}px`;
        img.style.left = `${x}px`;
        img.style.top = `${y}px`;
        img.style.transform = `rotate(${rotacion}deg)`;
        contenedor.appendChild(img);
      }
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', generar);
  } else {
    generar();
  }
})();

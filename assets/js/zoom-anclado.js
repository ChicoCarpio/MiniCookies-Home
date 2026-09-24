/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Zoom "anclado": el punto de la imagen que está bajo el puntero (o el centro
// de lo que se ve, si se usa un deslizador) queda en el mismo lugar de la
// pantalla después de acercar o alejar — como en cualquier visor de fotos.
//
// Las herramientas hacen zoom agrandando el elemento (canvas/imagen) dentro de
// un contenedor con scroll; sin ajustar el scroll, el zoom siempre quedaba
// anclado a la esquina superior izquierda de la imagen.
//
// contenedor: el elemento con overflow:auto que scrollea.
// elemento:   lo que se agranda (su getBoundingClientRect cambia con el zoom).
// clientX/Y:  punto de la pantalla que tiene que quedar fijo; si se omiten,
//             se usa el centro de la parte visible del contenedor.
// cambiarTamano: función que aplica el nuevo tamaño (sincrónica).
function zoomAnclado(contenedor, elemento, cambiarTamano, clientX, clientY) {
  const vista = contenedor.getBoundingClientRect();
  if (clientX === undefined || clientY === undefined) {
    clientX = vista.left + vista.width / 2;
    clientY = vista.top + vista.height / 2;
  }
  const antes = elemento.getBoundingClientRect();
  // Posición del punto dentro del elemento, como fracción (0..1).
  const fx = antes.width ? (clientX - antes.left) / antes.width : 0.5;
  const fy = antes.height ? (clientY - antes.top) / antes.height : 0.5;
  cambiarTamano();
  const despues = elemento.getBoundingClientRect(); // fuerza el nuevo layout
  contenedor.scrollLeft += despues.left + fx * despues.width - clientX;
  contenedor.scrollTop += despues.top + fy * despues.height - clientY;
}

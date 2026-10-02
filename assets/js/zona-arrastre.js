/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// La zona de "arrastrá tu archivo acá", una sola vez para todo el sitio.
//
// Por qué existe: el mismo cableado (clic, elegir, arrastrar, soltar) estaba
// copiado a mano en 77 herramientas y las copias ya se habían separado: 53 no
// limpiaban el selector después de elegir, así que volver a elegir EL MISMO
// archivo no hacía nada. Y ninguna se podía usar con teclado: la zona es un
// <div> y el <input type="file"> real está oculto, así que con Tab no se
// llegaba y con Enter no pasaba nada — la herramienta era inutilizable sin
// mouse.
//
// Uso:
//   conectarZonaArrastre(zona, entrada, (archivos) => { … });
// `archivos` es siempre un Array de File (del selector o de lo soltado).
//
// Además, al cargar, TODA .zona-arrastre de la página se vuelve accesible con
// teclado (foco con Tab, Enter/Espacio para elegir), también las que todavía
// conectan sus eventos a mano: Enter dispara su propio clic.

function conectarZonaArrastre(zona, entrada, alElegir) {
  zona.addEventListener('click', () => entrada.click());
  entrada.addEventListener('change', () => {
    // Se copia antes de limpiar: vaciar el selector vacía también su lista.
    const archivos = Array.from(entrada.files);
    // Limpiar permite volver a elegir el mismo archivo (si no, el navegador
    // no avisa del cambio porque "es el mismo").
    entrada.value = '';
    if (archivos.length) alElegir(archivos);
  });
  zona.addEventListener('dragover', (e) => { e.preventDefault(); zona.classList.add('activa'); });
  zona.addEventListener('dragleave', () => zona.classList.remove('activa'));
  zona.addEventListener('drop', (e) => {
    e.preventDefault();
    zona.classList.remove('activa');
    const archivos = Array.from(e.dataTransfer.files);
    if (archivos.length) alElegir(archivos);
  });
}

(function () {
  function hacerAccesible(zona) {
    if (zona.dataset.zonaAccesible) return;
    zona.dataset.zonaAccesible = '1';
    if (!zona.hasAttribute('tabindex')) zona.tabIndex = 0;
    if (!zona.hasAttribute('role')) zona.setAttribute('role', 'button');
    zona.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter' && e.key !== ' ') return;
      e.preventDefault(); // Espacio no debe desplazar la página
      zona.click();
    });
  }
  function todas() { document.querySelectorAll('.zona-arrastre').forEach(hacerAccesible); }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', todas);
  else todas();
})();

/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Tablas HTML → tablas Markdown (las de GitHub, con "|") para TurndownService.
// Turndown solo, sin reglas para tablas, dejaba cada celda como un párrafo
// suelto: una tabla de Word pasada a Markdown perdía filas y columnas.
// La usan Word → Markdown, HTML → Markdown y EPUB → Markdown.
//
// Markdown no tiene celdas combinadas ni tablas anidadas: una celda que
// abarca varias columnas deja las de al lado vacías, una que abarca varias
// filas queda solo en la primera, y una tabla dentro de una celda pasa a
// texto ("a / b; c / d"). Los saltos de línea de una celda pasan a <br>,
// que GitHub muestra como salto.

function agregarTablasMarkdown(servicio) {
  function celdaAMarkdown(celda) {
    const copia = celda.cloneNode(true);
    copia.querySelectorAll('table').forEach((anidada) => {
      const texto = [...anidada.rows]
        .map((fila) => [...fila.cells].map((c) => c.textContent.trim()).filter(Boolean).join(' / '))
        .filter(Boolean).join('; ');
      anidada.replaceWith(copia.ownerDocument.createTextNode(' ' + texto + ' '));
    });
    return servicio.turndown(copia.innerHTML)
      .trim()
      .replace(/\n+/g, '<br>')
      .replace(/\|/g, '\\|');
  }

  servicio.addRule('tablaMarkdown', {
    filter: 'table',
    replacement(_contenido, tabla) {
      const filas = [...tabla.rows].map((fila) => {
        const celdas = [];
        for (const celda of fila.cells) {
          celdas.push(celdaAMarkdown(celda));
          for (let i = 1; i < (parseInt(celda.getAttribute('colspan'), 10) || 1); i++) celdas.push('');
        }
        return celdas;
      }).filter((celdas) => celdas.length);
      if (!filas.length) return '';
      const columnas = Math.max(...filas.map((f) => f.length));
      const linea = (celdas) => '| ' + Array.from({ length: columnas }, (_, i) => celdas[i] || '').join(' | ') + ' |';
      const lineas = [linea(filas[0]), linea(Array(columnas).fill('---')), ...filas.slice(1).map(linea)];
      return '\n\n' + lineas.join('\n') + '\n\n';
    },
  });
  return servicio;
}

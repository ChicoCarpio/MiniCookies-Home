/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// PDF → EPUB conservando la estructura: no solo el texto.
//
// La lectura del PDF es de pdf-estructura.js. Un EPUB no tiene hojas: el texto
// fluye según la pantalla del lector. Por eso, a diferencia de PDF → Word:
//  - Se quitan los encabezados y pies que se repiten en cada hoja (número de
//    página, título corrido).
//  - Un párrafo que el PDF cortó al final de una hoja se une con su
//    continuación en la siguiente.
//  - Los títulos pasan a <h1>/<h2>/<h3>; cada <h1> abre un capítulo.
//  - Negritas y cursivas como <strong>/<em>; centrado y a la derecha se
//    conservan; las imágenes van en su lugar del texto (sin posición fija).
//
// Requiere: pdfjs.js, pdf-estructura.js, epub.js (construirEpub,
// dividirHtmlEnCapitulosEpub) y seguridad-html.js (escaparHtml).
// Uso: const blob = await PdfAEpub.convertir(arrayBuffer, { titulo, autor, alAvanzar });

(function () {
  function base64(bytes) {
    let binario = '';
    for (let i = 0; i < bytes.length; i += 0x8000) binario += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(binario);
  }

  function imagenHtml(img) {
    const tipo = img.tipo === 'png' ? 'image/png' : 'image/jpeg';
    return `<p style="text-align:center"><img src="data:${tipo};base64,${base64(img.bytes)}" alt=""/></p>`;
  }

  // Corridas de un párrafo → HTML en línea (las tabulaciones, que un EPUB no
  // tiene, pasan a un espacio largo).
  function corridasHtml(corridas) {
    return corridas.map((c) => {
      let html = escaparHtml(c.texto.replace(/\t/g, ' '));
      if (c.cursiva) html = `<em>${html}</em>`;
      if (c.negrita) html = `<strong>${html}</strong>`;
      return html;
    }).join('');
  }

  const terminaFrase = (texto) => /[.!?…:;»"”)\]]\s*$/.test(texto);
  const textoDe = (p) => p.corridas.map((c) => c.texto).join('').trim();

  async function convertir(bytes, opciones = {}) {
    const doc = await PdfEstructura.analizar(bytes, { alAvanzar: opciones.alAvanzar });
    // Nivel "en bruto" de un párrafo: el de su tamaño de letra (0 = el más
    // grande), 3 para un título menor (un renglón corto en negrita, algo más
    // grande que el texto), -1 si no es título. Sin letras ("* * *", un
    // separador de escena) nunca es título.
    const nivelBruto = (p) => {
      const texto = textoDe(p);
      if (!/\p{L}/u.test(texto)) return -1;
      const porTamano = p.renglones.length <= 3 ? doc.nivelTitulo(p.tamano) : -1;
      if (porTamano >= 0) return porTamano;
      const menor = p.renglones.length === 1 && texto.length <= 90 && p.corridas.every((c) => c.negrita) && p.tamano >= doc.tamanoNormal * 1.05;
      return menor ? 3 : -1;
    };
    const bloques = []; // { html } o { parrafo, corridas } (los párrafos se arman al final para poder unirlos)

    for (const pagina of doc.paginas) {
      if (pagina.sinTexto) {
        pagina.imagenes.forEach((img) => bloques.push({ html: imagenHtml(img) }));
        continue;
      }
      let primeroDeLaHoja = true;
      for (const b of pagina.bloques) {
        if (b.tipo === 'imagen') {
          bloques.push({ html: imagenHtml(b.img) });
          primeroDeLaHoja = false;
          continue;
        }
        const p = b.p;
        if (p.repetido) continue;
        (p.flotantes || []).forEach((img) => bloques.push({ html: imagenHtml(img) }));
        const nivel = nivelBruto(p);
        // ¿Sigue el párrafo de la hoja anterior? Aquel llegaba al margen derecho
        // sin terminar la frase, y este empieza la hoja con la misma letra.
        const anterior = bloques.at(-1);
        if (primeroDeLaHoja && nivel < 0 && anterior && anterior.parrafo && anterior.nivel < 0
            && Math.abs(anterior.parrafo.tamano - p.tamano) < 0.6 && p.alineacion !== 'centro'
            && anterior.parrafo.renglones.at(-1).x1 > anterior.columna.x1 - 4 - (anterior.columna.x1 - anterior.columna.x0) * 0.06
            && !terminaFrase(anterior.corridas.at(-1).texto)) {
          const ultima = anterior.corridas.at(-1);
          const primera = p.corridas[0];
          if (/[a-záéíóúñü]-$/i.test(ultima.texto) && /^[a-záéíóúñü]/.test(primera.texto)) ultima.texto = ultima.texto.slice(0, -1);
          else ultima.texto += ' ';
          anterior.corridas.push(...p.corridas);
          primeroDeLaHoja = false;
          continue;
        }
        bloques.push({ parrafo: p, corridas: p.corridas.map((c) => ({ ...c })), nivel, columna: pagina.columna });
        primeroDeLaHoja = false;
      }
    }

    // El nivel de capítulo es el más alto que se repite (un título único enorme,
    // como el del libro en la portada, no se queda con todos los capítulos): ese
    // y los de arriba son <h1> (cada uno abre capítulo); los de abajo, <h2>/<h3>.
    const usados = new Map();
    for (const b of bloques) if (b.nivel >= 0) usados.set(b.nivel, (usados.get(b.nivel) || 0) + 1);
    const niveles = [...usados.keys()].sort((a, b) => a - b);
    const capitulo = niveles.find((n) => usados.get(n) >= 2) ?? niveles[0];
    const etiqueta = (nivel) => (nivel <= capitulo ? 'h1' : niveles.filter((n) => n > capitulo).indexOf(nivel) === 0 ? 'h2' : 'h3');

    const html = bloques.map((b) => {
      if (b.html) return b.html;
      const contenido = corridasHtml(b.corridas);
      if (b.nivel >= 0) { const h = etiqueta(b.nivel); return `<${h}>${contenido}</${h}>`; }
      const alinea = { centro: ' style="text-align:center"', derecha: ' style="text-align:right"' }[b.parrafo.alineacion] || '';
      return `<p${alinea}>${contenido}</p>`;
    }).join('\n');

    const titulo = opciones.titulo || 'Libro';
    const capitulos = dividirHtmlEnCapitulosEpub(html, titulo);
    return construirEpub({ titulo, autor: opciones.autor || '', capitulos });
  }

  window.PdfAEpub = { convertir };
})();

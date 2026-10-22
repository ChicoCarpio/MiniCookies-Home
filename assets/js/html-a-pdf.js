/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// HTML ya dibujado por el navegador → PDF vectorial (pdf-lib), sin imprimir.
//
// Por qué: la impresión del navegador no deja guardar varios archivos ni
// armar un .zip, y con libros grandes guardaba PDF vacíos. Acá el navegador
// solo calcula el diseño (posición de cada palabra, imagen y borde) y el PDF
// se escribe a mano: el texto queda seleccionable y cada documento es su
// propio archivo.
//
// Cómo:
//  - Se recogen piezas con su rectángulo: palabras (Range.getClientRects),
//    imágenes, fondos y bordes.
//  - El documento se mide con su propia fuente. En el PDF se escribe con las
//    fuentes estándar (Helvetica, Times, Courier: 0 bytes, con negrita y
//    cursiva reales) y cada palabra se estira o encoge (operador Tz) al ancho
//    exacto que midió el navegador, así nada se pisa.
//  - Lo que esas fuentes no tienen (griego, cirílico...) va con Noto Sans
//    completa. No con subset: fontkit no sabe recortar una fuente variable y
//    las letras salían como cuadrados vacíos.
//  - El contenido puede ser más alto que una página (docx-preview no corta
//    páginas; un EPUB es un solo rollo): se pagina buscando cortes que no
//    partan ninguna línea ni imagen.
//
// Uso: const bytes = await HtmlAPdf.convertir([{ elemento, anchoPx, altoPx,
//        margenArribaPx, margenAbajoPx }], { titulo, alAvanzar(hoja, total) });
// Un break-before: page dentro del elemento fuerza una hoja nueva.
// El elemento tiene que estar en el documento y visible (puede estar fuera de
// pantalla) para que el navegador lo mida.
// La hoja para eso: const lienzo = HtmlAPdf.crearLienzo('libro'), y al final
// lienzo.remove().

(function () {
  const PX_A_PT = 0.75; // 96 px = 72 pt
  const URL_FUENTE = rutaVendor('noto-sans/2.015/NotoSans-VF.ttf');
  const URL_FONTKIT = rutaVendor('@pdf-lib/fontkit/1.1.1/dist/fontkit.umd.min.js?v=2');

  let fontkitListo = null;
  let bytesNoto = null;

  function cargarScript(src) {
    return new Promise((resolver, rechazar) => {
      const s = document.createElement('script');
      s.src = src;
      s.onload = resolver;
      s.onerror = () => rechazar(new Error('No se pudo cargar ' + src));
      document.head.appendChild(s);
    });
  }

  // Solo hace falta si aparece un carácter que las fuentes estándar no tienen.
  async function cargarNoto() {
    if (!window.fontkit) await (fontkitListo ||= cargarScript(URL_FONTKIT));
    if (!bytesNoto) bytesNoto = new Uint8Array(await (await fetch(URL_FUENTE)).arrayBuffer());
    return bytesNoto;
  }

  // Dónde cae la línea base dentro del rectángulo de una palabra: depende de la
  // fuente que usó el navegador, y lo dice el propio canvas.
  const lienzoMedida = document.createElement('canvas').getContext('2d');
  const proporcionAscenso = new Map();
  function ascensoDe(fuenteCss) {
    if (!proporcionAscenso.has(fuenteCss)) {
      lienzoMedida.font = fuenteCss;
      const m = lienzoMedida.measureText('Hg');
      const total = m.fontBoundingBoxAscent + m.fontBoundingBoxDescent;
      proporcionAscenso.set(fuenteCss, total > 0 ? m.fontBoundingBoxAscent / total : 0.8);
    }
    return proporcionAscenso.get(fuenteCss);
  }

  // Familia estándar del PDF más parecida a la que pidió el documento.
  function familiaEstandar(familiaCss) {
    const primera = familiaCss.split(',')[0].toLowerCase();
    if (/mono|courier|consolas|menlo/.test(primera)) return 'Courier';
    if (/times|georgia|cambria|garamond|book|palatino|serif/.test(primera) && !/sans/.test(primera)) return 'Times';
    return 'Helvetica';
  }

  function color(css) {
    const m = /rgba?\(([^)]+)\)/.exec(css || '');
    if (!m) return null;
    const p = m[1].split(/[ ,/]+/).filter(Boolean).map(Number);
    if (p.length > 3 && p[3] === 0) return null;
    return PDFLib.rgb(p[0] / 255, p[1] / 255, p[2] / 255);
  }

  function visible(elemento) {
    const e = getComputedStyle(elemento);
    return e.visibility !== 'hidden' && e.display !== 'none' && Number(e.opacity) !== 0;
  }

  // Piezas de una página lógica, en px relativos al borde superior izquierdo del contenido.
  function recogerPiezas(raiz, origen) {
    const piezas = [];
    const saltos = [];
    const relativo = (r) => ({ x: r.left - origen.x, y: r.top - origen.y, ancho: r.width, alto: r.height });

    // Fondos y bordes.
    for (const el of raiz.querySelectorAll('*')) {
      if (el.tagName === 'IMG' || !visible(el)) continue;
      const r = el.getBoundingClientRect();
      if (!r.width || !r.height) continue;
      const e = getComputedStyle(el);
      const caja = relativo(r);
      // break-before: page (un capítulo nuevo) → corte de página forzado.
      if (/^(page|always|left|right)$/.test(e.breakBefore) || e.pageBreakBefore === 'always') saltos.push(caja.y);
      const fondo = color(e.backgroundColor);
      if (fondo) piezas.push({ tipo: 'fondo', ...caja, color: fondo });
      for (const lado of ['Top', 'Right', 'Bottom', 'Left']) {
        const grosor = parseFloat(e[`border${lado}Width`]);
        if (!grosor || e[`border${lado}Style`] === 'none' || e[`border${lado}Style`] === 'hidden') continue;
        const c = color(e[`border${lado}Color`]);
        if (!c) continue;
        const b = { tipo: 'fondo', color: c };
        if (lado === 'Top') Object.assign(b, { x: caja.x, y: caja.y, ancho: caja.ancho, alto: grosor });
        if (lado === 'Bottom') Object.assign(b, { x: caja.x, y: caja.y + caja.alto - grosor, ancho: caja.ancho, alto: grosor });
        if (lado === 'Left') Object.assign(b, { x: caja.x, y: caja.y, ancho: grosor, alto: caja.alto });
        if (lado === 'Right') Object.assign(b, { x: caja.x + caja.ancho - grosor, y: caja.y, ancho: grosor, alto: caja.alto });
        piezas.push(b);
      }
    }

    // Imágenes.
    for (const img of raiz.querySelectorAll('img')) {
      if (!visible(img) || !img.complete || !img.naturalWidth) continue;
      const r = img.getBoundingClientRect();
      if (r.width && r.height) piezas.push({ tipo: 'imagen', ...relativo(r), img });
    }

    // Palabras: un rectángulo por palabra, medido por el propio navegador.
    const recorrido = document.createTreeWalker(raiz, NodeFilter.SHOW_TEXT);
    const rango = document.createRange();
    for (let nodo = recorrido.nextNode(); nodo; nodo = recorrido.nextNode()) {
      const padre = nodo.parentElement;
      if (!padre || !nodo.data.trim() || !visible(padre) || padre.closest('style, script')) continue;
      const e = getComputedStyle(padre);
      const estilo = {
        color: color(e.color) || PDFLib.rgb(0, 0, 0),
        tamano: parseFloat(e.fontSize),
        familia: familiaEstandar(e.fontFamily),
        ascenso: ascensoDe(`${e.fontStyle} ${e.fontWeight} ${e.fontSize} ${e.fontFamily}`),
        negrita: parseInt(e.fontWeight, 10) >= 600,
        cursiva: e.fontStyle === 'italic' || e.fontStyle === 'oblique',
        subrayado: (e.textDecorationLine || '').includes('underline'),
      };
      for (const m of nodo.data.matchAll(/\S+/g)) {
        rango.setStart(nodo, m.index);
        rango.setEnd(nodo, m.index + m[0].length);
        const rects = rango.getClientRects();
        if (!rects.length) continue;
        // Una palabra partida en dos renglones (guion, salto forzado): se usa el primero.
        const r = rects[0];
        if (!r.width) continue;
        piezas.push({ tipo: 'texto', ...relativo(r), texto: m[0], ...estilo });
      }
    }
    return { piezas, saltos };
  }

  // Cortes de página en el eje vertical: cada corte cae donde ninguna pieza
  // queda partida (si una pieza es más alta que la página, se la corta igual),
  // o en un salto forzado. Las piezas van ordenadas por y: un libro tiene
  // cientos de miles de palabras y recorrerlas todas por cada corte no termina.
  function calcularCortes(ordenadas, altoContenido, altoTotal, saltos) {
    const altoMaximo = ordenadas.reduce((m, p) => Math.max(m, p.alto), 0);
    const forzados = [...new Set(saltos)].sort((a, b) => a - b);
    const cortes = [0];
    let inicio = 0;
    // 1 px de tolerancia: el alto calculado de una hoja A4 (1122,53) y el medido
    // (1122,53125) difieren en milésimas y agregaban una hoja en blanco.
    while (inicio + altoContenido < altoTotal - 1 || forzados.some((f) => f > inicio + 1)) {
      const limite = inicio + altoContenido;
      const forzado = forzados.find((f) => f > inicio + 1 && f <= limite);
      let corte = forzado ?? limite;
      if (forzado === undefined) {
        for (let vuelta = 0; vuelta < 50; vuelta++) {
          // Piezas que empiezan antes del corte y terminan después: se busca la de más arriba.
          let partida = null;
          for (let k = indiceAntesDe(ordenadas, corte); k >= 0 && ordenadas[k].y > corte - altoMaximo; k--) {
            const p = ordenadas[k];
            if (p.y > inicio && p.y + p.alto > corte && (!partida || p.y < partida.y)) partida = p;
          }
          if (!partida) break;
          corte = partida.y;
        }
        // Si respetar la pieza deja la hoja casi vacía (una imagen casi tan alta
        // como la hoja), se corta en el lugar normal.
        if (corte - inicio < altoContenido * 0.15) corte = limite;
      }
      if (corte >= altoTotal - 1 && forzado === undefined) break;
      // Un salto forzado sobre una hoja que todavía no tiene nada (un capítulo
      // vacío seguido de otro) no deja una hoja en blanco: la hoja empieza ahí.
      const ultima = indiceAntesDe(ordenadas, corte);
      const hayContenido = ultima >= 0 && ordenadas[ultima].y >= inicio;
      if (forzado !== undefined && !hayContenido && cortes.length > 1) cortes[cortes.length - 1] = corte;
      else cortes.push(corte);
      inicio = corte;
    }
    return cortes;
  }

  // Último índice con y < valor (búsqueda binaria; -1 si no hay).
  function indiceAntesDe(ordenadas, valor) {
    let bajo = 0;
    let alto = ordenadas.length;
    while (bajo < alto) {
      const medio = (bajo + alto) >> 1;
      if (ordenadas[medio].y < valor) bajo = medio + 1; else alto = medio;
    }
    return bajo - 1;
  }

  async function bytesDeImagen(img) {
    const src = img.currentSrc || img.src;
    if (/^data:image\//i.test(src)) {
      // El tipo se mira en los bytes, no en el data: (un .docx puede declarar
      // PNG para una foto JPEG: la librería docx guarda todo como .png).
      const bytes = await (await fetch(src)).arrayBuffer();
      const b = new Uint8Array(bytes, 0, 4);
      if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return { bytes, png: true };
      if (b[0] === 0xff && b[1] === 0xd8) return { bytes, png: false };
    }
    // GIF, WebP, SVG...: se pasan a PNG con un canvas.
    const canvas = document.createElement('canvas');
    canvas.width = img.naturalWidth;
    canvas.height = img.naturalHeight;
    canvas.getContext('2d').drawImage(img, 0, 0);
    const blob = await new Promise((r) => canvas.toBlob(r, 'image/png'));
    return { bytes: await blob.arrayBuffer(), png: true };
  }

  async function convertir(paginas, opciones = {}) {
    const pdf = await PDFLib.PDFDocument.create();
    if (opciones.titulo) pdf.setTitle(opciones.titulo);

    const fuentes = new Map();
    async function fuenteEstandar(familia, negrita, cursiva) {
      const S = PDFLib.StandardFonts;
      const nombre = {
        Helvetica: [S.Helvetica, S.HelveticaBold, S.HelveticaOblique, S.HelveticaBoldOblique],
        Times: [S.TimesRoman, S.TimesRomanBold, S.TimesRomanItalic, S.TimesRomanBoldItalic],
        Courier: [S.Courier, S.CourierBold, S.CourierOblique, S.CourierBoldOblique],
      }[familia][(negrita ? 1 : 0) + (cursiva ? 2 : 0)];
      if (!fuentes.has(nombre)) fuentes.set(nombre, await pdf.embedFont(nombre));
      return fuentes.get(nombre);
    }
    let noto = null;
    async function fuenteNoto() {
      if (!noto) {
        const bytes = await cargarNoto();
        pdf.registerFontkit(window.fontkit);
        noto = await pdf.embedFont(bytes, { subset: false });
      }
      return noto;
    }

    const imagenesIncrustadas = new Map();

    for (const pagina of paginas) {
      const { elemento, anchoPx, altoPx } = pagina;
      const margenArriba = pagina.margenArribaPx || 0;
      const margenAbajo = pagina.margenAbajoPx || 0;
      const r = elemento.getBoundingClientRect();
      const origen = { x: r.left, y: r.top + margenArriba };
      const { piezas, saltos } = recogerPiezas(elemento, origen);
      const fondos = piezas.filter((p) => p.tipo === 'fondo');
      const ordenadas = piezas.filter((p) => p.tipo !== 'fondo').sort((a, b) => a.y - b.y);
      const altoContenido = altoPx - margenArriba - margenAbajo;
      const altoTotal = Math.max(r.height - margenArriba - margenAbajo, 1);
      const cortes = calcularCortes(ordenadas, altoContenido, altoTotal, saltos.filter((y) => y > 0));
      // Cada pieza a su hoja una sola vez (lo de antes del primer corte, como un
      // encabezado, va a la primera).
      const porHoja = cortes.map(() => []);
      let hojaActual = 0;
      for (const p of ordenadas) {
        while (hojaActual + 1 < cortes.length && p.y >= cortes[hojaActual + 1]) hojaActual++;
        porHoja[hojaActual].push(p);
      }

      for (let i = 0; i < cortes.length; i++) {
        if (opciones.alAvanzar) opciones.alAvanzar(i, cortes.length);
        // Respiro para la página cada tanto: un libro son cientos de hojas.
        if (i % 20 === 19) await new Promise((r) => setTimeout(r, 0));
        const desde = cortes[i];
        const hasta = i + 1 < cortes.length ? cortes[i + 1] : Infinity;
        const hoja = pdf.addPage([anchoPx * PX_A_PT, altoPx * PX_A_PT]);
        const altoHoja = altoPx * PX_A_PT;
        // px de contenido → pt de la hoja (pdf-lib mide desde abajo).
        const aX = (x) => x * PX_A_PT;
        const aY = (y) => altoHoja - (y - desde + margenArriba) * PX_A_PT;

        for (const p of [...fondos, ...porHoja[i]]) {
          if (p.tipo === 'fondo') {
            // Un fondo puede ocupar varias hojas: se recorta a esta. El relleno
            // de la propia hoja (margen) no se recorta: es el papel.
            const arriba = Math.max(p.y, i === 0 ? -margenArriba : desde);
            const abajo = Math.min(p.y + p.alto, hasta);
            if (abajo <= arriba) continue;
            hoja.drawRectangle({ x: aX(p.x), y: aY(abajo), width: p.ancho * PX_A_PT, height: (abajo - arriba) * PX_A_PT, color: p.color });
            continue;
          }
          if (p.tipo === 'imagen') {
            let incrustada = imagenesIncrustadas.get(p.img);
            if (!incrustada) {
              const datos = await bytesDeImagen(p.img);
              incrustada = datos.png ? await pdf.embedPng(datos.bytes) : await pdf.embedJpg(datos.bytes);
              imagenesIncrustadas.set(p.img, incrustada);
            }
            hoja.drawImage(incrustada, { x: aX(p.x), y: aY(p.y + p.alto), width: p.ancho * PX_A_PT, height: p.alto * PX_A_PT });
            continue;
          }
          const tamano = p.tamano * PX_A_PT;
          const base = aY(p.y + p.ascenso * p.alto);
          let fuente = await fuenteEstandar(p.familia, p.negrita, p.cursiva);
          let anchoNatural;
          try {
            anchoNatural = fuente.widthOfTextAtSize(p.texto, tamano);
          } catch {
            // Un carácter fuera de WinAnsi: la fuente estándar no lo puede escribir.
            fuente = await fuenteNoto();
            anchoNatural = fuente.widthOfTextAtSize(p.texto, tamano);
          }
          // Tz: la palabra ocupa en el PDF el mismo ancho que en pantalla.
          const estiramiento = anchoNatural > 0 ? Math.min(200, Math.max(50, (100 * p.ancho * PX_A_PT) / anchoNatural)) : 100;
          hoja.pushOperators(PDFLib.setCharacterSqueeze(estiramiento));
          hoja.drawText(p.texto, { x: aX(p.x), y: base, size: tamano, font: fuente, color: p.color });
          hoja.pushOperators(PDFLib.setCharacterSqueeze(100));
          if (p.subrayado) {
            hoja.drawLine({ start: { x: aX(p.x), y: base - tamano * 0.12 }, end: { x: aX(p.x + p.ancho), y: base - tamano * 0.12 }, thickness: Math.max(0.5, tamano * 0.06), color: p.color });
          }
        }
      }
    }
    return pdf.save();
  }

  // La hoja fuera de pantalla donde se arma el documento para medirlo
  // (.lienzo-pdf en estilos.css; con clase 'libro', una hoja A4 con letra de
  // libro). Quien la pide la saca con remove() al terminar. Antes cada página
  // tenía su propio <div id="lienzoPdf"> con sus estilos.
  function crearLienzo(clase) {
    const lienzo = document.createElement('div');
    lienzo.className = 'lienzo-pdf' + (clase ? ' ' + clase : '');
    lienzo.setAttribute('aria-hidden', 'true');
    document.body.appendChild(lienzo);
    return lienzo;
  }

  window.HtmlAPdf = { convertir, crearLienzo };
})();

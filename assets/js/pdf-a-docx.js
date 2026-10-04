/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// PDF → Word (.docx) conservando la estructura: no solo el texto.
//
// Un PDF no guarda párrafos ni estilos: guarda letras en posiciones fijas,
// imágenes y trazos. Acá se reconstruye lo que un Word necesita:
//  - Texto con su tamaño, negrita, cursiva y fuente (de la fuente del PDF).
//  - Renglones agrupados en párrafos (Word los vuelve a cortar según el ancho),
//    con alineación (izquierda, centrada, derecha, justificada), sangría,
//    espacio entre párrafos y títulos (letra bastante más grande que la normal).
//  - Huecos grandes dentro de un renglón ("Nombre:      Juan") como
//    tabulaciones con su posición.
//  - Imágenes: si hay texto al lado (una foto junto a un título) van flotando
//    en su posición exacta de la página; si no, en su propio párrafo.
//  - Líneas horizontales finas (subrayados de sección) como borde inferior
//    del párrafo de arriba.
//  - Un salto de página por cada página del PDF, mismo tamaño de hoja y márgenes.
//
// Lo que no reconstruye (todavía): tablas (salen como renglones con
// tabulaciones), colores de texto, columnas.
//
// Requiere: pdfjs.js (abrirPdf, cargarPdfjs) y la librería docx (window.docx).
// Uso: const blob = await PdfADocx.convertir(arrayBuffer);

(function () {
  const PT_A_TWIP = 20;
  const PT_A_EMU = 12700;
  const PT_A_PX = 96 / 72;

  // Producto de matrices de PDF [a b c d e f].
  function multiplicar(m, n) {
    return [
      m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1],
      m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3],
      m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5],
    ];
  }
  function aplicar(m, x, y) {
    return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
  }
  function cajaDe(m, x0, y0, x1, y1) {
    const p = [aplicar(m, x0, y0), aplicar(m, x1, y0), aplicar(m, x0, y1), aplicar(m, x1, y1)];
    const xs = p.map((q) => q[0]);
    const ys = p.map((q) => q[1]);
    return { x: Math.min(...xs), y: Math.min(...ys), ancho: Math.max(...xs) - Math.min(...xs), alto: Math.max(...ys) - Math.min(...ys) };
  }

  // "ABCDEF+TimesNewRomanPS-BoldItalicMT" → { familia: 'Times New Roman', negrita, cursiva }
  function leerFuente(nombre, fuente) {
    const crudo = (nombre || '').replace(/^[A-Z]{6}\+/, '');
    const negrita = /bold|black|heavy|semibold|demi/i.test(crudo) || !!(fuente && (fuente.bold || fuente.black));
    const cursiva = /italic|oblique/i.test(crudo) || !!(fuente && fuente.italic);
    let familia = crudo.split(/[-,]/)[0]
      .replace(/(PSMT|PS|MT)$/, '')
      .replace(/(Bold|Italic|Oblique|Regular|Black)+$/i, '')
      .replace(/([a-z])([A-Z])/g, '$1 $2')
      .trim();
    if (!familia || /^(g_|f\d|T\d)/i.test(familia)) familia = '';
    return { familia, negrita, cursiva };
  }

  function objetoPdf(contenedor, id) {
    return new Promise((resolver) => {
      try {
        if (contenedor.has(id)) { resolver(contenedor.get(id)); return; }
        contenedor.get(id, resolver);
      } catch {
        resolver(null);
      }
    });
  }

  // Imagen de PDF.js (bitmap o píxeles crudos) → bytes PNG o JPEG.
  async function imagenABytes(img, pdfjs) {
    if (!img || !img.width || !img.height) return null;
    const canvas = document.createElement('canvas');
    canvas.width = img.width;
    canvas.height = img.height;
    const ctx = canvas.getContext('2d');
    if (img.bitmap) {
      ctx.drawImage(img.bitmap, 0, 0);
    } else if (img.data) {
      const datos = ctx.createImageData(img.width, img.height);
      const destino = datos.data;
      const origen = img.data;
      const K = pdfjs.ImageKind;
      if (img.kind === K.RGBA_32BPP) {
        destino.set(origen.subarray(0, destino.length));
      } else if (img.kind === K.RGB_24BPP) {
        for (let i = 0, j = 0; j < destino.length; i += 3, j += 4) {
          destino[j] = origen[i]; destino[j + 1] = origen[i + 1]; destino[j + 2] = origen[i + 2]; destino[j + 3] = 255;
        }
      } else if (img.kind === K.GRAYSCALE_1BPP) {
        const porFila = (img.width + 7) >> 3;
        for (let y = 0; y < img.height; y++) {
          for (let x = 0; x < img.width; x++) {
            const bit = (origen[y * porFila + (x >> 3)] >> (7 - (x & 7))) & 1;
            const j = (y * img.width + x) * 4;
            const v = bit ? 255 : 0;
            destino[j] = destino[j + 1] = destino[j + 2] = v; destino[j + 3] = 255;
          }
        }
      } else {
        return null;
      }
      ctx.putImageData(datos, 0, 0);
    } else {
      return null;
    }
    // Con transparencia, PNG; si no, JPEG (una foto en PNG pesa varias veces más).
    const pixeles = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
    let transparente = false;
    for (let i = 3; i < pixeles.length; i += 4) { if (pixeles[i] < 250) { transparente = true; break; } }
    const tipo = transparente ? 'image/png' : 'image/jpeg';
    const blob = await new Promise((r) => canvas.toBlob(r, tipo, 0.9));
    return { bytes: new Uint8Array(await blob.arrayBuffer()), tipo: transparente ? 'png' : 'jpg' };
  }

  // Imágenes y líneas horizontales de una página, en pt desde arriba a la izquierda.
  async function leerGraficos(pagina, vista, pdfjs) {
    const ops = await pagina.getOperatorList();
    const O = pdfjs.OPS;
    const imagenes = [];
    const lineas = [];
    let trazos = 0;
    let ctm = [1, 0, 0, 1, 0, 0];
    const pila = [];
    for (let i = 0; i < ops.fnArray.length; i++) {
      const fn = ops.fnArray[i];
      const args = ops.argsArray[i];
      if (fn === O.save) pila.push(ctm);
      else if (fn === O.restore) ctm = pila.pop() || [1, 0, 0, 1, 0, 0];
      else if (fn === O.transform) ctm = multiplicar(ctm, args);
      else if (fn === O.paintFormXObjectBegin) { pila.push(ctm); if (args[0]) ctm = multiplicar(ctm, args[0]); }
      else if (fn === O.paintFormXObjectEnd) ctm = pila.pop() || [1, 0, 0, 1, 0, 0];
      else if (fn === O.paintImageXObject || fn === O.paintInlineImageXObject) {
        const caja = cajaDe(multiplicar(vista.transform, ctm), 0, 0, 1, 1);
        if (caja.ancho < 8 || caja.alto < 8) continue;
        const img = fn === O.paintInlineImageXObject
          ? args[0]
          : await objetoPdf(String(args[0]).startsWith('g_') ? pagina.commonObjs : pagina.objs, args[0]);
        const datos = await imagenABytes(img, pdfjs);
        if (datos) imagenes.push({ ...caja, ...datos });
      } else if (fn === O.constructPath) {
        trazos++;
        // El último argumento con 4 números es la caja del trazo (sin transformar).
        const minMax = [...args].reverse().find((a) => a && a.length === 4 && typeof a[0] === 'number');
        if (!minMax || !isFinite(minMax[0])) continue;
        const caja = cajaDe(multiplicar(vista.transform, ctm), minMax[0], minMax[1], minMax[2], minMax[3]);
        if (caja.alto <= 3 && caja.ancho >= 40) lineas.push({ x: caja.x, y: caja.y + caja.alto / 2, ancho: caja.ancho });
      }
    }
    return { imagenes, lineas, trazos };
  }

  // La página entera como imagen: para lo que no se puede reconstruir (letras
  // dibujadas como trazos, por ejemplo una página "impresa" desde un navegador).
  async function paginaComoImagen(pagina) {
    const vista = pagina.getViewport({ scale: 2 });
    const canvas = document.createElement('canvas');
    canvas.width = Math.round(vista.width);
    canvas.height = Math.round(vista.height);
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await pagina.render({ canvasContext: ctx, viewport: vista }).promise;
    const blob = await new Promise((r) => canvas.toBlob(r, 'image/jpeg', 0.88));
    return { bytes: new Uint8Array(await blob.arrayBuffer()), tipo: 'jpg' };
  }

  // Trozos de texto con estilo, en pt desde arriba a la izquierda (y = línea base).
  async function leerTexto(pagina, vista) {
    const contenido = await pagina.getTextContent();
    const trozos = [];
    for (const item of contenido.items) {
      if (!item.str || !item.str.trim()) continue;
      const m = multiplicar(vista.transform, item.transform);
      if (Math.abs(m[1]) > 0.01 || Math.abs(m[2]) > 0.01) continue; // texto girado: no se reconstruye
      const tamano = Math.hypot(m[2], m[3]) || Math.abs(m[3]);
      let fuente = null;
      try { if (pagina.commonObjs.has(item.fontName)) fuente = pagina.commonObjs.get(item.fontName); } catch {}
      const estilo = leerFuente(fuente && fuente.name, fuente);
      if (!estilo.familia) {
        const css = contenido.styles[item.fontName];
        estilo.familia = css && /serif/i.test(css.fontFamily) && !/sans/i.test(css.fontFamily) ? 'Times New Roman' : '';
      }
      trozos.push({ texto: item.str, x: m[4], y: m[5], ancho: item.width * (vista.scale || 1), tamano, ...estilo });
    }
    return trozos;
  }

  // Trozos → renglones (misma línea base), con los huecos grandes como tabulaciones.
  function armarRenglones(trozos) {
    const ordenados = [...trozos].sort((a, b) => a.y - b.y || a.x - b.x);
    const renglones = [];
    for (const t of ordenados) {
      const r = renglones.find((q) => Math.abs(q.y - t.y) <= Math.max(q.tamano, t.tamano) * 0.35);
      if (r) { r.trozos.push(t); r.tamano = Math.max(r.tamano, t.tamano); } else renglones.push({ y: t.y, tamano: t.tamano, trozos: [t] });
    }
    for (const r of renglones) {
      r.trozos.sort((a, b) => a.x - b.x);
      r.x0 = r.trozos[0].x;
      r.x1 = Math.max(...r.trozos.map((t) => t.x + t.ancho));
      r.corridas = [];
      r.tabs = [];
      let fin = null;
      for (const t of r.trozos) {
        let texto = t.texto;
        if (fin !== null) {
          const hueco = t.x - fin;
          if (hueco > t.tamano * 1.5) { r.tabs.push(t.x); texto = '\t' + texto.replace(/^\s+/, ''); }
          else if (hueco > t.tamano * 0.15 && !/\s$/.test(r.corridas.at(-1)?.texto || '') && !/^\s/.test(texto)) texto = ' ' + texto;
        }
        const anterior = r.corridas.at(-1);
        if (anterior && anterior.negrita === t.negrita && anterior.cursiva === t.cursiva && Math.abs(anterior.tamano - t.tamano) < 0.3 && anterior.familia === t.familia) {
          anterior.texto += texto;
        } else {
          r.corridas.push({ texto, negrita: t.negrita, cursiva: t.cursiva, tamano: t.tamano, familia: t.familia });
        }
        fin = t.x + t.ancho;
      }
      r.corridas.at(-1).texto = r.corridas.at(-1).texto.replace(/\s+$/, '');
    }
    return renglones;
  }

  // Renglones → párrafos, con alineación y sangría según la columna de texto.
  function armarParrafos(renglones, columna) {
    const anchoCol = columna.x1 - columna.x0;
    const centroCol = (columna.x0 + columna.x1) / 2;
    const alineacion = (r) => {
      if (r.tabs.length) return 'izquierda';
      const centro = (r.x0 + r.x1) / 2;
      if (r.x1 - r.x0 < anchoCol * 0.92 && Math.abs(centro - centroCol) < anchoCol * 0.04 && r.x0 > columna.x0 + anchoCol * 0.04) return 'centro';
      if (r.x1 > columna.x1 - anchoCol * 0.02 && r.x0 > columna.x0 + anchoCol * 0.3) return 'derecha';
      return 'izquierda';
    };
    const parrafos = [];
    for (const r of renglones) {
      r.alineacion = alineacion(r);
      const p = parrafos.at(-1);
      const previo = p && p.renglones.at(-1);
      let unir = false;
      if (previo && !r.tabs.length && !previo.tabs.length && Math.abs(previo.tamano - r.tamano) < 0.6 && previo.alineacion === r.alineacion) {
        const hueco = r.y - previo.y;
        const esperado = p.interlineado || r.tamano * 1.25;
        const huecoNormal = hueco > r.tamano * 0.9 && hueco < esperado * 1.3 && hueco < r.tamano * 2.2;
        if (r.alineacion === 'izquierda') {
          // El renglón anterior llegaba al margen derecho (se cortó por ancho) y este arranca en el margen del párrafo.
          // Y la primera palabra de este no cabía al final de aquel: si cabía,
          // el renglón terminó a propósito (un punto de una lista, por ejemplo).
          const t0 = r.trozos[0];
          const primeraPalabra = (t0.texto.trim().split(/\s+/)[0] || '').length;
          const anchoPalabra = t0.ancho * (primeraPalabra / Math.max(1, t0.texto.trim().length));
          const llenaba = previo.x1 > columna.x1 - anchoCol * 0.12 && previo.x1 + r.tamano * 0.3 + anchoPalabra > columna.x1 - 4;
          const margen = p.renglones.length > 1 ? p.renglones[1].x0 : r.x0;
          unir = huecoNormal && llenaba && Math.abs(r.x0 - margen) < r.tamano * 1.5;
        } else if (r.alineacion === 'centro') {
          unir = huecoNormal && previo.x1 - previo.x0 > anchoCol * 0.7;
        }
        if (unir && !p.interlineado) p.interlineado = hueco;
      }
      if (unir) p.renglones.push(r);
      else parrafos.push({ renglones: [r] });
    }
    for (const p of parrafos) {
      const primero = p.renglones[0];
      const resto = p.renglones.slice(1);
      p.alineacion = primero.alineacion;
      if (p.alineacion === 'izquierda' && resto.length && p.renglones.slice(0, -1).every((r) => r.x1 > columna.x1 - anchoCol * 0.008)) p.alineacion = 'justificado';
      p.izquierda = p.alineacion === 'centro' || p.alineacion === 'derecha' ? 0 : Math.max(0, (resto[0] || primero).x0 - columna.x0);
      p.primeraLinea = resto.length ? primero.x0 - resto[0].x0 : 0;
      p.arriba = primero.y - primero.tamano * 0.85;
      p.abajo = p.renglones.at(-1).y + p.renglones.at(-1).tamano * 0.25;
      p.tamano = primero.tamano;
      p.tabs = primero.tabs.map((x) => x - columna.x0);
      // Corridas de todos los renglones, uniendo con espacio (o quitando el guion de corte).
      p.corridas = [];
      p.renglones.forEach((r, i) => {
        r.corridas.forEach((c, j) => {
          let texto = c.texto;
          if (i > 0 && j === 0) {
            const ultima = p.corridas.at(-1);
            if (ultima && /[a-záéíóúñü]-$/i.test(ultima.texto) && /^[a-záéíóúñü]/.test(texto)) ultima.texto = ultima.texto.slice(0, -1);
            else texto = ' ' + texto;
          }
          const ultima = p.corridas.at(-1);
          if (ultima && ultima.negrita === c.negrita && ultima.cursiva === c.cursiva && Math.abs(ultima.tamano - c.tamano) < 0.3 && ultima.familia === c.familia) ultima.texto += texto;
          else p.corridas.push({ ...c, texto });
        });
      });
    }
    return parrafos;
  }

  async function convertir(bytes) {
    const pdfjs = await cargarPdfjs();
    const pdf = await abrirPdf(new Uint8Array(bytes));
    const D = window.docx;
    const paginas = [];
    for (let n = 1; n <= pdf.numPages; n++) {
      const pagina = await pdf.getPage(n);
      const vista = pagina.getViewport({ scale: 1 });
      const graficos = await leerGraficos(pagina, vista, pdfjs); // primero: deja las fuentes cargadas
      const trozos = await leerTexto(pagina, vista);
      // Sin texto: si lo que hay son fotos (un escaneo), se usan tal cual; si hay
      // dibujos o nada extraíble, la página se guarda como imagen para no perderla.
      if (!trozos.length && (!graficos.imagenes.length || graficos.trazos > 40)) {
        graficos.imagenes = [{ x: 0, y: 0, ancho: vista.width, alto: vista.height, ...(await paginaComoImagen(pagina)) }];
      }
      paginas.push({ ancho: vista.width, alto: vista.height, trozos, ...graficos });
      pagina.cleanup();
    }

    // Columna de texto del documento entero (márgenes izquierdo y derecho) y tamaño normal.
    const todos = paginas.flatMap((p) => p.trozos);
    const ancho0 = paginas[0]?.ancho || 612;
    const alto0 = paginas[0]?.alto || 792;
    const columna = todos.length
      ? { x0: Math.min(...todos.map((t) => t.x)), x1: Math.max(...todos.map((t) => t.x + t.ancho)) }
      : { x0: 72, x1: ancho0 - 72 };
    columna.x0 = Math.max(18, Math.min(columna.x0, ancho0 / 3));
    // Unos puntos de holgura: un renglón que en el PDF llega justo al margen no
    // tiene que partirse en Word por diferencias de redondeo.
    columna.x1 = Math.min(ancho0 - 18, Math.max(columna.x1 + 4, (ancho0 * 2) / 3));
    const conteo = new Map();
    for (const t of todos) { const k = Math.round(t.tamano * 2) / 2; conteo.set(k, (conteo.get(k) || 0) + t.texto.length); }
    const tamanoNormal = [...conteo.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || 11;
    const tamanosTitulo = [...new Set(todos.filter((t) => t.tamano >= tamanoNormal * 1.35).map((t) => Math.round(t.tamano)))].sort((a, b) => b - a).slice(0, 3);
    const nivelTitulo = (tamano) => {
      const i = tamanosTitulo.indexOf(Math.round(tamano));
      return i < 0 ? null : [D.HeadingLevel.HEADING_1, D.HeadingLevel.HEADING_2, D.HeadingLevel.HEADING_3][i];
    };
    const margenArriba = Math.max(18, Math.min(108, Math.min(...paginas.map((p) => Math.min(...p.trozos.map((t) => t.y - t.tamano), ...p.imagenes.map((i) => i.y), p.alto))) || 72));

    const margenAbajo = Math.max(18, Math.min(72, Math.min(...paginas.map((p) => p.alto - Math.max(0, ...p.trozos.map((t) => t.y + t.tamano * 0.3), ...p.imagenes.map((i) => i.y + i.alto)))) || 36));

    // Imagen en su posición exacta de la hoja, delante del texto y sin apartarlo:
    // la posición ya es la del PDF.
    const imagenFlotante = (img) => new D.ImageRun({
      data: img.bytes,
      transformation: { width: Math.round(img.ancho * PT_A_PX), height: Math.round(img.alto * PT_A_PX) },
      floating: {
        horizontalPosition: { relative: D.HorizontalPositionRelativeFrom.PAGE, offset: Math.round(img.x * PT_A_EMU) },
        verticalPosition: { relative: D.VerticalPositionRelativeFrom.PAGE, offset: Math.round(img.y * PT_A_EMU) },
        wrap: { type: D.TextWrappingType.NONE },
        allowOverlap: true,
        behindDocument: false,
      },
    });

    const hijos = [];
    paginas.forEach((pagina, numero) => {
      // Página sin texto (un escaneo, una portada): sus imágenes van tal cual
      // estaban en la hoja, ancladas a un párrafo vacío.
      if (!pagina.trozos.length) {
        hijos.push(new D.Paragraph({ pageBreakBefore: numero > 0, children: pagina.imagenes.map(imagenFlotante) }));
        return;
      }
      const parrafos = armarParrafos(armarRenglones(pagina.trozos), columna);
      // Cada línea horizontal pasa a borde inferior del párrafo que tiene justo arriba.
      for (const linea of pagina.lineas) {
        if (linea.ancho < (columna.x1 - columna.x0) * 0.4) continue;
        const arriba = parrafos.filter((p) => p.abajo <= linea.y + 2 && linea.y - p.abajo < 30).at(-1);
        if (arriba) arriba.bordeAbajo = true;
      }
      // Imágenes: flotantes si comparten altura con texto que está a su lado; si no, en línea.
      const bloques = parrafos.map((p) => ({ tipo: 'parrafo', arriba: p.arriba, p }));
      const flotantes = new Map();
      for (const img of pagina.imagenes) {
        const pagCompleta = img.ancho * img.alto > pagina.ancho * pagina.alto * 0.6;
        const alLado = !pagCompleta && pagina.trozos.some((t) => t.y > img.y && t.y - t.tamano < img.y + img.alto && (t.x + t.ancho < img.x - 2 || t.x > img.x + img.ancho + 2));
        if (alLado) {
          const ancla = parrafos.find((p) => p.abajo >= img.y) || parrafos.at(-1);
          if (ancla) { if (!flotantes.has(ancla)) flotantes.set(ancla, []); flotantes.get(ancla).push(img); continue; }
        }
        bloques.push({ tipo: 'imagen', arriba: img.y, img, pagCompleta });
      }
      bloques.sort((a, b) => a.arriba - b.arriba);

      let abajoAnterior = margenArriba;
      bloques.forEach((b, i) => {
        const saltoPagina = numero > 0 && i === 0;
        const antes = Math.max(0, Math.min(72, b.arriba - abajoAnterior));
        if (b.tipo === 'imagen') {
          const img = b.img;
          // Que entre en la columna y en la hoja (una página escaneada entera, por ejemplo).
          const escala = Math.min(1, (columna.x1 - columna.x0) / img.ancho, (alto0 - margenArriba - margenAbajo) / img.alto);
          const centro = img.x + img.ancho / 2;
          const anchoCol = columna.x1 - columna.x0;
          const alinea = Math.abs(centro - (columna.x0 + columna.x1) / 2) < anchoCol * 0.1 ? D.AlignmentType.CENTER
            : img.x > columna.x0 + anchoCol * 0.4 ? D.AlignmentType.RIGHT : D.AlignmentType.LEFT;
          hijos.push(new D.Paragraph({
            pageBreakBefore: saltoPagina,
            alignment: alinea,
            spacing: { before: Math.round((b.pagCompleta ? 0 : antes) * PT_A_TWIP), after: 0 },
            children: [new D.ImageRun({ data: img.bytes, transformation: { width: Math.round(img.ancho * escala * PT_A_PX), height: Math.round(img.alto * escala * PT_A_PX) } })],
          }));
          abajoAnterior = img.y + img.alto;
          return;
        }
        const p = b.p;
        const nivel = nivelTitulo(p.tamano);
        const corridas = p.corridas.map((c) => {
          // Los \t se escriben como tabulaciones reales de Word.
          const partes = c.texto.split('\t');
          const contenido = partes.length > 1
            ? { children: partes.flatMap((parte, k) => (k ? [new D.Tab(), parte] : [parte])) }
            : { text: c.texto };
          return new D.TextRun({
            ...contenido,
            bold: c.negrita,
            italics: c.cursiva,
            size: Math.round(c.tamano * 2),
            font: c.familia || undefined,
            color: '000000',
          });
        });
        for (const img of flotantes.get(p) || []) corridas.unshift(imagenFlotante(img));
        const alineacion = { centro: D.AlignmentType.CENTER, derecha: D.AlignmentType.RIGHT, justificado: D.AlignmentType.JUSTIFIED }[p.alineacion] || D.AlignmentType.LEFT;
        hijos.push(new D.Paragraph({
          pageBreakBefore: saltoPagina,
          heading: nivel || undefined,
          alignment: alineacion,
          indent: { left: Math.round(p.izquierda * PT_A_TWIP), firstLine: p.primeraLinea > 1 ? Math.round(p.primeraLinea * PT_A_TWIP) : undefined, hanging: p.primeraLinea < -1 ? Math.round(-p.primeraLinea * PT_A_TWIP) : undefined },
          spacing: {
            before: Math.round(antes * PT_A_TWIP),
            after: 0,
            line: p.interlineado ? Math.round(p.interlineado * PT_A_TWIP) : undefined,
            lineRule: p.interlineado ? D.LineRuleType.AT_LEAST : undefined,
          },
          tabStops: p.tabs.length ? p.tabs.map((x) => ({ type: D.TabStopType.LEFT, position: Math.round(x * PT_A_TWIP) })) : undefined,
          border: p.bordeAbajo ? { bottom: { style: D.BorderStyle.SINGLE, size: 8, color: '000000', space: 1 } } : undefined,
          children: corridas,
        }));
        abajoAnterior = p.abajo;
      });
    });

    const documento = new D.Document({
      sections: [{
        properties: {
          page: {
            size: { width: Math.round(ancho0 * PT_A_TWIP), height: Math.round(alto0 * PT_A_TWIP) },
            margin: {
              top: Math.round(margenArriba * PT_A_TWIP),
              bottom: Math.round(margenAbajo * PT_A_TWIP),
              left: Math.round(columna.x0 * PT_A_TWIP),
              right: Math.round((ancho0 - columna.x1) * PT_A_TWIP),
            },
          },
        },
        children: hijos.length ? hijos : [new D.Paragraph({ children: [] })],
      }],
    });
    if (pdf.cleanup) pdf.cleanup();
    return D.Packer.toBlob(documento);
  }

  window.PdfADocx = { convertir };
})();

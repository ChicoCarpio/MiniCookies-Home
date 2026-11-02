/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Lectura de la estructura de un PDF: compartida por PDF → Word (pdf-a-docx.js)
// y PDF → EPUB (pdf-a-epub.js).
//
// Un PDF no guarda párrafos ni estilos: guarda letras en posiciones fijas,
// imágenes y trazos. Acá se reconstruye:
//  - Texto con su tamaño, negrita, cursiva y fuente (de la fuente del PDF).
//  - Renglones agrupados en párrafos (no se une un renglón si la primera
//    palabra del siguiente cabía: un punto de lista), con alineación, sangría
//    y títulos (letra bastante más grande que la normal).
//  - Huecos grandes dentro de un renglón como tabulaciones con su posición.
//  - Imágenes; las que tienen texto al lado se marcan como flotantes.
//  - Líneas horizontales finas como borde inferior del párrafo de arriba.
//  - Encabezados y pies que se repiten en las hojas (número de página, título
//    corrido) marcados como "repetido": Word los deja en su lugar, un EPUB no.
//  - Página sin texto real: sus imágenes tal cual, o la página entera como
//    imagen si las letras son trazos.
//
// Requiere: pdfjs.js (abrirPdf, cargarPdfjs).
// Uso: const doc = await PdfEstructura.analizar(arrayBuffer, { alAvanzar(n, total) });

(function () {

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

  // Ancho de un fragmento dentro de un trozo de texto del PDF: la proporción que
  // ocupa medida con un canvas, aplicada al ancho real del trozo. Contar letras
  // fallaba con palabras de letras anchas ("muñeca" salía corta y un párrafo
  // quedaba partido); la proporción entre letras se parece de una fuente a otra.
  const lienzoMedida = document.createElement('canvas').getContext('2d');
  // Artículos, preposiciones y conjunciones (es/en): una frase no termina en ellas.
  const PALABRAS_DE_ENLACE = new Set(('el la los las lo un una unos unas su sus mi mis tu tus al del de en y e o u a con por para sin sobre entre que como ' +
    'the a an of to in on at by for with and or but from into').split(' '));
  function anchoDentroDe(fragmento, trozo) {
    const texto = trozo.texto;
    if (!fragmento || !texto.trim()) return 0;
    lienzoMedida.font = `${trozo.negrita ? 'bold ' : ''}${trozo.cursiva ? 'italic ' : ''}20px ${/times|serif|georgia|garamond|book/i.test(trozo.familia) ? 'serif' : 'sans-serif'}`;
    const total = lienzoMedida.measureText(texto).width;
    return total > 0 ? (trozo.ancho * lienzoMedida.measureText(fragmento).width) / total : 0;
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
  // columna: la de esta hoja (alineación, unir renglones); margen: el izquierdo
  // del documento, desde donde se miden sangría y tabulaciones.
  function armarParrafos(renglones, columna, margen = columna.x0) {
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
          const palabra = t0.texto.trim().split(/\s+/)[0] || '';
          const anchoPalabra = anchoDentroDe(palabra, t0);
          const anchoEspacio = anchoDentroDe(' ', t0) || r.tamano * 0.25;
          // Sobrante < 0: la palabra no cabía. Al límite (menos de medio carácter),
          // un renglón que sigue con mayúscula tras otro sin coma ni guion es un
          // punto de lista, no un corte por ancho.
          const sobrante = columna.x1 - 4 - (previo.x1 + anchoEspacio + anchoPalabra);
          const finPrevio = previo.trozos.at(-1).texto.trim();
          // (Salvo que aquel termine en una palabra que nunca cierra una frase:
          // "en el / Calabozo" sigue siendo el mismo párrafo.)
          const ultimaPalabra = (finPrevio.split(/\s+/).at(-1) || '').toLowerCase().replace(/[^\p{L}]/gu, '');
          const alLimite = sobrante > -r.tamano * 0.6 && /^\p{Lu}/u.test(palabra) && !/[,;\-–]$/.test(finPrevio) && !PALABRAS_DE_ENLACE.has(ultimaPalabra);
          const llenaba = previo.x1 > columna.x1 - anchoCol * 0.25 && sobrante < 0 && !alLimite;
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
      p.izquierda = p.alineacion === 'centro' || p.alineacion === 'derecha' ? 0 : Math.max(0, (resto[0] || primero).x0 - margen);
      p.primeraLinea = resto.length ? primero.x0 - resto[0].x0 : 0;
      p.arriba = primero.y - primero.tamano * 0.85;
      p.abajo = p.renglones.at(-1).y + p.renglones.at(-1).tamano * 0.25;
      p.tamano = primero.tamano;
      p.tabs = primero.tabs.map((x) => x - margen);
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

  // Mínimo y máximo sin Math.min(...lista): un libro tiene cientos de miles
  // de trozos y tantos argumentos rompen la pila.
  function minimo(lista, valor, porDefecto) {
    let m = Infinity;
    for (const x of lista) { const v = valor(x); if (v < m) m = v; }
    return m === Infinity ? porDefecto : m;
  }
  function maximo(lista, valor, porDefecto) {
    let m = -Infinity;
    for (const x of lista) { const v = valor(x); if (v > m) m = v; }
    return m === -Infinity ? porDefecto : m;
  }

  // Encabezados y pies: párrafos de un renglón cerca del borde de la hoja que
  // se repiten en 3 o más hojas (sin contar números), o que son solo un número.
  function marcarRepetidos(paginas) {
    const texto = (p) => p.corridas.map((c) => c.texto).join('').trim();
    const clave = (p) => texto(p).replace(/\d+/g, '').replace(/\s+/g, ' ').trim().toLowerCase();
    const enBorde = (p, pagina) => p.renglones.length === 1 && (p.abajo < pagina.alto * 0.1 || p.arriba > pagina.alto * 0.9);
    const veces = new Map();
    for (const pagina of paginas) {
      for (const k of new Set((pagina.parrafos || []).filter((p) => enBorde(p, pagina)).map(clave))) veces.set(k, (veces.get(k) || 0) + 1);
    }
    for (const pagina of paginas) {
      for (const p of pagina.parrafos || []) {
        if (!enBorde(p, pagina)) continue;
        if (/^[-–—\s]*([0-9]+|[ivxlcdm]+)[-–—\s]*$/i.test(texto(p)) || (clave(p) && veces.get(clave(p)) >= 3)) p.repetido = true;
      }
    }
  }

  async function analizar(bytes, opciones = {}) {
    const pdfjs = await cargarPdfjs();
    const pdf = await abrirPdf(new Uint8Array(bytes));
    const paginas = [];
    for (let n = 1; n <= pdf.numPages; n++) {
      if (opciones.alAvanzar) opciones.alAvanzar(n - 1, pdf.numPages);
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
    if (pdf.cleanup) pdf.cleanup();

    // Columna de texto del documento entero (márgenes izquierdo y derecho) y tamaño normal.
    const todos = paginas.flatMap((p) => p.trozos);
    const ancho0 = paginas[0]?.ancho || 612;
    const alto0 = paginas[0]?.alto || 792;
    const columna = todos.length
      ? { x0: minimo(todos, (t) => t.x), x1: maximo(todos, (t) => t.x + t.ancho) }
      : { x0: 72, x1: ancho0 - 72 };
    columna.x0 = Math.max(18, Math.min(columna.x0, ancho0 / 3));
    // Unos puntos de holgura: un renglón que en el PDF llega justo al margen no
    // tiene que partirse en Word por diferencias de redondeo.
    columna.x1 = Math.min(ancho0 - 18, Math.max(columna.x1 + 4, (ancho0 * 2) / 3));
    const conteo = new Map();
    for (const t of todos) { const k = Math.round(t.tamano * 2) / 2; conteo.set(k, (conteo.get(k) || 0) + t.texto.length); }
    const tamanoNormal = [...conteo.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || 11;
    const tamanosTitulo = [...new Set(todos.filter((t) => t.tamano >= tamanoNormal * 1.35).map((t) => Math.round(t.tamano)))].sort((a, b) => b - a).slice(0, 3);
    // 0, 1 o 2 según el tamaño (de mayor a menor); -1 si no es título.
    const nivelTitulo = (tamano) => tamanosTitulo.indexOf(Math.round(tamano));
    const margenArriba = Math.max(18, Math.min(108, minimo(paginas, (p) => Math.min(minimo(p.trozos, (t) => t.y - t.tamano, p.alto), minimo(p.imagenes, (i) => i.y, p.alto)), 72)));
    const margenAbajo = Math.max(18, Math.min(72, minimo(paginas, (p) => p.alto - Math.max(0, maximo(p.trozos, (t) => t.y + t.tamano * 0.3, 0), maximo(p.imagenes, (i) => i.y + i.alto, 0)), 36)));

    for (const pagina of paginas) {
      pagina.sinTexto = !pagina.trozos.length;
      if (pagina.sinTexto) continue;
      // Columna de esta hoja: un libro alterna márgenes entre páginas pares e
      // impares, y contra la del documento entero un renglón lleno parecía corto
      // (y un párrafo quedaba partido). Hoja con poco texto: la del documento.
      const renglones = armarRenglones(pagina.trozos);
      pagina.columna = renglones.length >= 6
        ? { x0: minimo(renglones, (r) => r.x0), x1: maximo(renglones, (r) => r.x1) + 4 }
        : columna;
      if (pagina.columna.x1 - pagina.columna.x0 < (columna.x1 - columna.x0) * 0.6) pagina.columna = columna;
      const parrafos = armarParrafos(renglones, pagina.columna, columna.x0);
      // Cada línea horizontal pasa a borde inferior del párrafo que tiene justo arriba.
      for (const linea of pagina.lineas) {
        if (linea.ancho < (columna.x1 - columna.x0) * 0.4) continue;
        const arriba = parrafos.filter((p) => p.abajo <= linea.y + 2 && linea.y - p.abajo < 30).at(-1);
        if (arriba) arriba.bordeAbajo = true;
      }
      // Imágenes: flotantes si comparten altura con texto que está a su lado; si no, en el flujo.
      const bloques = parrafos.map((p) => ({ tipo: 'parrafo', arriba: p.arriba, p }));
      for (const img of pagina.imagenes) {
        const pagCompleta = img.ancho * img.alto > pagina.ancho * pagina.alto * 0.6;
        const alLado = !pagCompleta && pagina.trozos.some((t) => t.y > img.y && t.y - t.tamano < img.y + img.alto && (t.x + t.ancho < img.x - 2 || t.x > img.x + img.ancho + 2));
        if (alLado) {
          const ancla = parrafos.find((p) => p.abajo >= img.y) || parrafos.at(-1);
          if (ancla) { (ancla.flotantes ||= []).push(img); continue; }
        }
        bloques.push({ tipo: 'imagen', arriba: img.y, img, pagCompleta });
      }
      bloques.sort((a, b) => a.arriba - b.arriba);
      pagina.parrafos = parrafos;
      pagina.bloques = bloques;
    }
    marcarRepetidos(paginas);
    return { paginas, columna, tamanoNormal, nivelTitulo, margenArriba, margenAbajo, ancho0, alto0 };
  }

  window.PdfEstructura = { analizar };
})();

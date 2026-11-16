/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Traductor de documentos · PDF. Un PDF no tiene párrafos: tiene trozos de texto sueltos con su posición. Acá se arman los párrafos, se
// traducen y se escribe la traducción en el mismo lugar, tapando el original con un recuadro del color del fondo y achicando la letra si
// no entra. Una página escaneada (sin texto) pasa antes por Librarius (OCR) y sigue igual. También se puede sacar como Word o texto.
//
// Parte sin navegador (la prueban las pruebas en Node):
//   TraductorPdf.bloques(items, { anchoPagina })   items = [{ texto, x, y, ancho, tam, negrita? }]  (x, y: línea base, PDF: y hacia arriba)
//        -> [{ texto, x0, x1, y0, y1, tam, interlinea, lineas }]   un bloque por párrafo, en orden de lectura
//   TraductorPdf.acomodar(texto, { ancho, alto, tam, interlinea, medir, asiatico })   medir(texto, tam) -> ancho
//        -> { tam, lineas: [texto], cabe }   la letra baja de a 5 % hasta que entra (hasta la mitad); si no entra, cabe = false
// Parte con navegador (PDF.js para leer, pdf-lib para escribir; ver traducir):
//   const { bytes, texto, bloques, estadistica } = await TraductorPdf.traducir(datos, { traducir, destino, progreso, fuentes, ocr })
(function (raiz) {
  'use strict';

  const CJK = /[぀-ヿ㐀-鿿豈-﫿가-힯　-〿＀-￯]/;
  const FIN_ORACION = /[.!?:;。！？…]["'”’»)]*$/;

  // ───────── Líneas y bloques ─────────
  // 1) Las piezas se juntan por altura (misma línea base ± 35 % del tamaño: las palabras del OCR traen cada una la suya, apenas distinta);
  // 2) dentro de cada altura, de izquierda a derecha; 3) un hueco de más de 2,5 veces el tamaño separa columnas (dos líneas distintas).
  function lineas(items) {
    const alturas = [];
    for (const it of items.filter((i) => i.texto && i.texto.trim() && i.tam > 0).sort((a, b) => b.y - a.y)) {
      const g = alturas.find((x) => Math.abs(x.y - it.y) < 0.35 * Math.max(x.tam, it.tam));
      if (g) { g.items.push(it); g.tam = Math.max(g.tam, it.tam); } else alturas.push({ y: it.y, tam: it.tam, items: [it] });
    }
    const ordenados = alturas.flatMap((g) => g.items.sort((a, b) => a.x - b.x).map((it) => ({ ...it, yLinea: g.y })));
    const ls = [];
    for (const it of ordenados) {
      const l = ls.find((x) => x.y === it.yLinea && it.x >= x.x1 - 0.5 * it.tam && it.x - x.x1 < 2.5 * it.tam);
      if (l) {
        const hueco = it.x - l.x1;
        const pegar = hueco < 0.15 * it.tam || /\s$/.test(l.texto) || /^\s/.test(it.texto) || (CJK.test(l.texto.slice(-1)) && CJK.test(it.texto[0]));
        l.texto += (pegar ? '' : ' ') + it.texto;
        l.x1 = Math.max(l.x1, it.x + it.ancho);
        l.tam = Math.max(l.tam, it.tam);
        l.negritas += it.negrita ? it.texto.length : 0; l.letras += it.texto.length;
      } else {
        ls.push({ texto: it.texto, x0: it.x, x1: it.x + it.ancho, y: it.yLinea, tam: it.tam, negritas: it.negrita ? it.texto.length : 0, letras: it.texto.length });
      }
    }
    for (const l of ls) l.texto = l.texto.replace(/\s+/g, ' ').trim();
    return ls.sort((a, b) => b.y - a.y || a.x0 - b.x0);
  }

  function unirLinea(a, b) {
    // «infor-» + «mación» -> «información» (guion de fin de línea antes de minúscula); en CJK, sin espacio.
    if (/\p{L}-$/u.test(a) && /^\p{Ll}/u.test(b)) return a.slice(0, -1) + b;
    if (CJK.test(a.slice(-1)) && CJK.test(b[0])) return a + b;
    return a + ' ' + b;
  }

  function bloques(items) {
    const bs = [];
    for (const l of lineas(items)) {
      const negrita = l.negritas > l.letras / 2;
      // Un bloque abierto que siga: mismo tamaño (± 20 %), justo debajo (hueco de menos de 1,8 líneas), que se superpongan en horizontal,
      // y que no sea un párrafo nuevo (la línea anterior cerró con punto y fue corta, o la nueva tiene sangría).
      const b = bs.slice().reverse().find((x) => {
        if (x.cerrado) return false;
        const u = x.ult;
        const hueco = u.y - l.y;
        if (Math.abs(x.tam - l.tam) > 0.2 * x.tam || hueco <= 0 || hueco > 1.8 * x.tam) return false;
        if (l.x0 > x.x1 || l.x1 < x.x0) return false;
        if (negrita !== x.negrita) return false;
        if (FIN_ORACION.test(u.texto) && (u.x1 < x.x0 + 0.75 * (x.x1 - x.x0) || l.x0 > x.x0 + 1.2 * x.tam)) return false;
        return true;
      });
      if (b) {
        b.interlinea = b.lineas.length === 1 ? b.ult.y - l.y : (b.interlinea * (b.lineas.length - 1) + (b.ult.y - l.y)) / b.lineas.length;
        b.texto = unirLinea(b.texto, l.texto);
        b.lineas.push(l); b.ult = l;
        b.x0 = Math.min(b.x0, l.x0); b.x1 = Math.max(b.x1, l.x1);
        b.y0 = l.y - 0.25 * l.tam;
      } else {
        // Las líneas por encima de esta que ya no pueden seguir (quedaron arriba de un hueco grande) se cierran.
        for (const x of bs) if (!x.cerrado && x.ult.y - l.y > 1.8 * x.tam && !(l.x0 > x.x1 || l.x1 < x.x0)) x.cerrado = true;
        bs.push({ texto: l.texto, lineas: [l], ult: l, x0: l.x0, x1: l.x1, y1: l.y + 0.85 * l.tam, y0: l.y - 0.25 * l.tam, tam: l.tam, interlinea: 1.2 * l.tam, negrita });
      }
    }
    // Orden de lectura: por columnas (bloques que se superponen en horizontal más de la mitad del más angosto van juntos, de arriba abajo),
    // las columnas de izquierda a derecha. Un título que cruza las dos columnas queda con la primera.
    const columnas = [];
    for (const b of bs) {
      const c = columnas.find((k) => Math.min(k.x1, b.x1) - Math.max(k.x0, b.x0) > 0.5 * Math.min(k.x1 - k.x0, b.x1 - b.x0));
      if (c) { c.bloques.push(b); c.x0 = Math.min(c.x0, b.x0); c.x1 = Math.max(c.x1, b.x1); } else columnas.push({ x0: b.x0, x1: b.x1, bloques: [b] });
    }
    columnas.sort((a, b) => a.x0 - b.x0);
    const ordenados = columnas.flatMap((c) => c.bloques.sort((a, b) => b.y1 - a.y1));
    return ordenados.map(({ texto, x0, x1, y0, y1, tam, interlinea, lineas: ls, negrita }) => ({ texto, x0, x1, y0, y1, tam, interlinea, negrita, lineas: ls.length }));
  }

  // ───────── Acomodar la traducción ─────────
  // Chino y japonés se cortan por carácter, con las reglas de su tipografía: una línea no empieza con un signo de cierre (。、，」）
  // ni con un kana chico (っ, ャ…), ni termina con uno de apertura (「（); las palabras latinas y los números no se parten.
  const CIERRE_CJK = '、。，．！？：；）」』】〉》〕｝ー…‥・ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ!?.,:;)]}%';
  const APERTURA_CJK = '（「『【〈《〔｛([{';
  const RE_UNIDAD_CJK = new RegExp(`\\s+|[${APERTURA_CJK.replace(/[[\]]/g, '\\$&')}]*(?:[\\p{Script=Latin}\\p{Script=Greek}\\p{Script=Cyrillic}\\p{Nd}'’-]+|\\S)[${CIERRE_CJK.replace(/[[\]]/g, '\\$&')}]*`, 'gu');
  function cortar(texto, ancho, tam, medir, asiatico) {
    const palabras = asiatico ? texto.replace(/\s+/g, ' ').match(RE_UNIDAD_CJK) || [] : texto.split(/\s+/).filter(Boolean);
    const sep = asiatico ? '' : ' ';
    const ls = [];
    let actual = '';
    for (const p of palabras) {
      const prueba = actual ? actual + sep + p : p;
      if (actual && medir(prueba, tam) > ancho) { ls.push(actual); actual = p.trim() ? p : ''; } else actual = prueba;
    }
    if (actual) ls.push(actual);
    return ls;
  }

  function acomodar(texto, { ancho, alto, tam, interlinea, medir, asiatico }) {
    const proporcion = (interlinea || 1.2 * tam) / tam;
    let t = tam;
    for (let i = 0; i <= 14; i++) {
      const ls = cortar(texto, ancho, t, medir, asiatico);
      const largo = ls.reduce((m, l) => Math.max(m, medir(l, t)), 0);
      // Alto que ocupa: las líneas con su interlineado (la última solo con su cuerpo).
      const ocupa = (ls.length - 1) * proporcion * t + 1.1 * t;
      if (ocupa <= alto * 1.02 + 0.5 && largo <= ancho * 1.02 + 0.5) return { tam: t, interlinea: proporcion * t, lineas: ls, cabe: true };
      t *= 0.95;
    }
    return { tam: t, interlinea: proporcion * t, lineas: cortar(texto, ancho, t, medir, asiatico), cabe: false };
  }

  // ───────── Con navegador ─────────
  // Trozos de una página de PDF.js -> items. Se descarta el texto girado (se deja como está) y el invisible (capa de OCR de otro programa,
  // modo de dibujo 3: no se ve, no se tapa ni se traduce encima).
  async function itemsDePagina(pagina) {
    const contenido = await pagina.getTextContent();
    const items = [];
    for (const it of contenido.items) {
      if (!it.str || !it.str.trim() || !it.transform) continue;
      const [a, b, c, d, e, f] = it.transform;
      const tam = Math.hypot(c, d);
      if (!tam || Math.abs(b) > 0.02 * tam || Math.abs(c) > 0.02 * tam || a < 0) continue;
      let nombre = '';
      try { nombre = pagina.commonObjs.has(it.fontName) ? String(pagina.commonObjs.get(it.fontName).name || '') : ''; } catch (err) { nombre = ''; }
      const familia = (contenido.styles[it.fontName] || {}).fontFamily || '';
      items.push({ texto: it.str, x: e, y: f, ancho: it.width, tam, negrita: /bold|black|heavy|semibold/i.test(nombre),
        serifa: /serif(?!-?sans)|times|roman|georgia|garamond|cambria|book|minion/i.test(nombre + ' ' + familia) && !/sans/i.test(nombre + ' ' + familia) });
    }
    return items;
  }

  // Colores de fondo y de letra de cada bloque, mirando la página dibujada: el fondo = el color más común en un borde alrededor del
  // bloque; la letra = el más oscuro (o más distinto del fondo) adentro.
  // aPixel(x, y): de puntos del PDF a píxeles del lienzo (viewport.convertToViewportPoint: tiene en cuenta el giro y el origen de la página).
  function colores(ctx, aPixel, b) {
    const W = ctx.canvas.width, H = ctx.canvas.height;
    const [ax, ay] = aPixel(b.x0 - 2, b.y1 + 2), [bx, by] = aPixel(b.x1 + 2, b.y0 - 2);
    const lim = (v, M) => Math.max(0, Math.min(M - 1, Math.round(v)));
    const x0 = lim(Math.min(ax, bx), W), x1 = lim(Math.max(ax, bx), W), y0 = lim(Math.min(ay, by), H), y1 = lim(Math.max(ay, by), H);
    const datos = ctx.getImageData(x0, y0, Math.max(1, x1 - x0 + 1), Math.max(1, y1 - y0 + 1));
    const w = datos.width, h = datos.height, d = datos.data;
    // Los colores del borde se agrupan en escalones para encontrar el más común; el fondo es el PROMEDIO real de los píxeles de ese grupo
    // (el centro del escalón daba 252 en vez de 255 y el tapado se veía como un recuadro gris tenue).
    const grupos = new Map();
    const borde = (i, j) => i < 2 || j < 2 || i >= w - 2 || j >= h - 2;
    for (let j = 0; j < h; j++) for (let i = 0; i < w; i++) {
      if (!borde(i, j)) continue;
      const k = (j * w + i) * 4;
      const c = ((d[k] >> 3) << 10) | ((d[k + 1] >> 3) << 5) | (d[k + 2] >> 3);
      const g = grupos.get(c) || [0, 0, 0, 0];
      g[0]++; g[1] += d[k]; g[2] += d[k + 1]; g[3] += d[k + 2];
      grupos.set(c, g);
    }
    let fondo = [255, 255, 255], max = 0;
    for (const g of grupos.values()) if (g[0] > max) { max = g[0]; fondo = [g[1] / g[0], g[2] / g[0], g[3] / g[0]]; }
    let letra = null, dist = 0;
    for (let k = 0; k < d.length; k += 4) {
      const dd = Math.abs(d[k] - fondo[0]) + Math.abs(d[k + 1] - fondo[1]) + Math.abs(d[k + 2] - fondo[2]);
      if (dd > dist) { dist = dd; letra = [d[k], d[k + 1], d[k + 2]]; }
    }
    if (!letra || dist < 90) letra = (fondo[0] + fondo[1] + fondo[2]) / 3 > 128 ? [0, 0, 0] : [255, 255, 255];
    return { fondo: fondo.map((v) => Math.min(255, v) / 255), letra: letra.map((v) => v / 255) };
  }

  const ASIATICO = new Set(['zh', 'ja', 'ko']);

  // ───────── Quitar el texto original (camino «quitar») ─────────
  // Lee las instrucciones de dibujo de una página (su «content stream») y borra las que escriben texto (Tj, TJ, ' y "), salvo el texto
  // girado o espejado (no se traduce, queda). Todo lo demás (imágenes, fondos, líneas, colores) queda igual, byte por byte.
  // ' y " además pasan de línea: se cambian por T* (y Tw/Tc para ") para que el texto que queda no se corra.
  const BLANCO = new Set([0, 9, 10, 12, 13, 32]);
  const DELIM = new Set('()<>[]{}/%'.split('').map((c) => c.charCodeAt(0)));

  function quitarTextoDeContenido(b) {
    const n = b.length;
    const salida = [];
    let copiadoHasta = 0;
    let i = 0;
    let operandos = [];       // { tipo, valor, ini }
    let iniOperandos = -1;
    const pilaCtm = [];
    let ctm = [1, 0, 0, 1, 0, 0], tm = [1, 0, 0, 1, 0, 0];
    const mult = (m, k) => [m[0] * k[0] + m[1] * k[2], m[0] * k[1] + m[1] * k[3], m[2] * k[0] + m[3] * k[2], m[2] * k[1] + m[3] * k[3],
      m[4] * k[0] + m[5] * k[2] + k[4], m[4] * k[1] + m[5] * k[3] + k[5]];
    const recto = () => {
      const m = mult(tm, ctm);
      const e = Math.hypot(m[0], m[1]) || 1;
      return Math.abs(m[1]) < 0.02 * e && Math.abs(m[2]) < 0.02 * e && m[0] > 0 && m[3] > 0;
    };
    let quitados = 0, conservados = 0;
    const saltarBlanco = () => {
      while (i < n) {
        if (BLANCO.has(b[i])) i++;
        else if (b[i] === 37) { while (i < n && b[i] !== 10 && b[i] !== 13) i++; }   // comentario
        else break;
      }
    };
    const leerCadena = () => {          // ( … ) con anidado y escapes
      let prof = 0;
      for (; i < n; i++) {
        if (b[i] === 92) { i++; continue; }
        if (b[i] === 40) prof++;
        else if (b[i] === 41) { prof--; if (prof === 0) { i++; return; } }
      }
    };
    const leerHasta = (cierre) => { while (i < n && b[i] !== cierre) i++; i++; };
    const leerArreglo = () => {        // [ … ] (con cadenas adentro)
      i++;
      for (;;) {
        saltarBlanco();
        if (i >= n) return;
        if (b[i] === 93) { i++; return; }
        leerObjeto();
      }
    };
    const leerDiccionario = () => {    // << … >>
      i += 2;
      for (;;) {
        saltarBlanco();
        if (i >= n) return;
        if (b[i] === 62 && b[i + 1] === 62) { i += 2; return; }
        leerObjeto();
      }
    };
    function leerObjeto() {
      const ini = i, c = b[i];
      if (c === 40) { leerCadena(); return { tipo: 'cadena', ini }; }
      if (c === 60 && b[i + 1] === 60) { leerDiccionario(); return { tipo: 'dic', ini }; }
      if (c === 60) { leerHasta(62); return { tipo: 'cadena', ini }; }
      if (c === 91) { leerArreglo(); return { tipo: 'arreglo', ini }; }
      if (c === 47) { i++; while (i < n && !BLANCO.has(b[i]) && !DELIM.has(b[i])) i++; return { tipo: 'nombre', ini }; }
      let s = '';
      while (i < n && !BLANCO.has(b[i]) && !DELIM.has(b[i])) s += String.fromCharCode(b[i++]);
      if (!s) { i++; return { tipo: 'otro', ini }; }
      if (/^[+-]?(\d+\.?\d*|\.\d+)$/.test(s)) return { tipo: 'numero', valor: Number(s), ini };
      if (s === 'true' || s === 'false' || s === 'null') return { tipo: 'otro', ini };
      return { tipo: 'operador', valor: s, ini };
    }
    while (i < n) {
      saltarBlanco();
      if (i >= n) break;
      const o = leerObjeto();
      if (o.tipo !== 'operador') { if (iniOperandos < 0) iniOperandos = o.ini; operandos.push(o); continue; }
      const op = o.valor;
      const nums = operandos.map((x) => x.valor);
      if (op === 'q') pilaCtm.push(ctm);
      else if (op === 'Q') ctm = pilaCtm.pop() || [1, 0, 0, 1, 0, 0];
      else if (op === 'cm' && nums.length === 6) ctm = mult(nums, ctm);
      else if (op === 'BT') tm = [1, 0, 0, 1, 0, 0];
      else if (op === 'Tm' && nums.length === 6) tm = nums;
      else if (op === 'BI') {
        // Imagen en línea: BI … ID <binario> EI. El binario se salta entero (puede tener cualquier byte).
        const id = (() => { for (let k = i; k < n - 1; k++) if (b[k] === 73 && b[k + 1] === 68 && BLANCO.has(b[k + 2])) return k; return n; })();
        let k = id + 3;
        while (k < n - 1 && !(BLANCO.has(b[k - 1]) && b[k] === 69 && b[k + 1] === 73 && (k + 2 >= n || BLANCO.has(b[k + 2]) || DELIM.has(b[k + 2])))) k++;
        i = Math.min(n, k + 2);
      } else if (op === 'Tj' || op === 'TJ' || op === "'" || op === '"') {
        if (recto()) {
          const ini = iniOperandos >= 0 ? iniOperandos : o.ini;
          salida.push(b.subarray(copiadoHasta, ini));
          if (op === "'") salida.push(new TextEncoder().encode(' T* '));
          else if (op === '"' && operandos.length === 3) {
            salida.push(b.subarray(operandos[0].ini, operandos[1].ini), new TextEncoder().encode(' Tw '),
              b.subarray(operandos[1].ini, operandos[2].ini), new TextEncoder().encode(' Tc T* '));
          } else salida.push(new TextEncoder().encode(' '));
          copiadoHasta = i;
          quitados++;
        } else conservados++;
      }
      operandos = []; iniOperandos = -1;
    }
    salida.push(b.subarray(copiadoHasta));
    const total = salida.reduce((s, x) => s + x.length, 0);
    const r = new Uint8Array(total);
    let k = 0;
    for (const x of salida) { r.set(x, k); k += x.length; }
    return { bytes: r, quitados, conservados };
  }

  // Las instrucciones de una página (uno o varios streams, ya descomprimidos y juntos) -> sin el texto, en un stream nuevo comprimido.
  function quitarTextoDePagina(doc, pag) {
    const PDFLib = raiz.PDFLib;
    const { PDFName, PDFArray, PDFRawStream, decodePDFRawStream } = PDFLib;
    const contenido = pag.node.get(PDFName.of('Contents'));
    if (!contenido) return { quitados: 0, conservados: 0 };
    const refs = contenido instanceof PDFArray ? contenido.asArray() : [contenido];
    const partes = [];
    for (const ref of refs) {
      const s = doc.context.lookup(ref);
      if (!(s instanceof PDFRawStream)) return { quitados: 0, conservados: 0, error: 'contenido raro' };
      partes.push(decodePDFRawStream(s).decode(), new Uint8Array([10]));
    }
    const junto = new Uint8Array(partes.reduce((t, p) => t + p.length, 0));
    let k = 0;
    for (const p of partes) { junto.set(p, k); k += p.length; }
    const r = quitarTextoDeContenido(junto);
    pag.node.set(PDFName.of('Contents'), doc.context.register(doc.context.flateStream(r.bytes)));
    return r;
  }

  // opciones: traducir, destino, progreso; abrir(bytes) -> PDFDocumentProxy (por defecto window.abrirPdf); fuentes = { sans, sansNegrita,
  // serifa, serifaNegrita, cjk, cjkNegrita } (bytes .ttf; cjk obligatoria para zh/ja/ko: Noto Sans SC/JP/KR FIJAS de
  // preparar-fuentes-cjk.py, nunca la variable, que fontkit recorta mal); ocr = { leerPagina(canvas) } para las páginas escaneadas;
  // modo: 'tapar' (por defecto: recuadro del color del fondo sobre el original) o 'quitar' (borra el texto original; tapa solo donde no pudo).
  async function traducir(datos, opciones) {
    const { traducir: traductor, destino, progreso, fuentes = {}, ocr } = opciones;
    const PDFLib = raiz.PDFLib;
    if (ASIATICO.has(destino) && !fuentes.cjk) throw new Error('Falta la fuente para chino, japonés o coreano');
    const bytes = datos instanceof Uint8Array ? datos : new Uint8Array(datos instanceof ArrayBuffer ? datos : await datos.arrayBuffer());
    const abrir = opciones.abrir || raiz.abrirPdf;
    const pdf = await abrir(bytes.slice());
    const paginas = [];
    const ESCALA = 1.5;
    for (let n = 1; n <= pdf.numPages; n++) {
      const pagina = await pdf.getPage(n);
      const vista = pagina.getViewport({ scale: ESCALA });
      const lienzo = document.createElement('canvas');
      lienzo.width = Math.ceil(vista.width); lienzo.height = Math.ceil(vista.height);
      const ctx = lienzo.getContext('2d', { willReadFrequently: true });
      await pagina.render({ canvasContext: ctx, viewport: vista }).promise;
      let items = await itemsDePagina(pagina);
      let escaneada = false;
      // Escaneada: casi sin texto. Librarius lee la página dibujada; sus palabras (en píxeles, y hacia abajo) pasan a puntos del PDF.
      if (items.reduce((s, i) => s + i.texto.trim().length, 0) < 20 && ocr) {
        const r = await ocr.leerPagina(lienzo);
        items = r.palabras.filter((p) => p.texto && Math.abs(p.angulo || 0) < 2).map((p) => {
          const [x, y] = vista.convertToPdfPoint(p.x, p.y);
          return { texto: p.texto, x, y, ancho: p.ancho / ESCALA, tam: p.alto / ESCALA * 1.15 };
        });
        escaneada = items.length > 0;
      }
      const bs = bloques(items);
      const aPixel = (x, y) => vista.convertToViewportPoint(x, y);
      for (const b of bs) Object.assign(b, colores(ctx, aPixel, b));
      // Serifa: la mayoría de los trozos de la página.
      const conSerifa = items.filter((i) => i.serifa).length > items.length / 2;
      for (const b of bs) b.serifa = conSerifa;
      paginas.push({ bloques: bs, escaneada });
      lienzo.width = lienzo.height = 0;
    }
    const todos = paginas.flatMap((p) => p.bloques);
    const r = await raiz.TraductorDocumentos.traducirUnidades(todos.map((b) => [{ texto: b.texto, formato: '' }]), {
      traducir: (segs, d, avisar) => traductor(segs, d, { progreso: avisar }), destino, progreso,
    });
    todos.forEach((b, i) => { b.traduccion = r[i].map((t) => t.texto).join('').trim(); });

    // Modo «quitar»: se borra el texto original de cada página (no de las escaneadas: ahí el texto es parte de la imagen) y se vuelve a
    // leer con PDF.js: un bloque que todavía tenga texto debajo (texto dentro de un objeto reutilizado, Form XObject) se tapa igual.
    // Modo «tapar» (por defecto): todos los bloques se tapan con un recuadro del color del fondo.
    const quitar = opciones.modo === 'quitar';
    const aTapar = new Set();
    let base = bytes;
    let quitados = 0;
    if (quitar) {
      const limpio = await PDFLib.PDFDocument.load(bytes, { updateMetadata: false });
      paginas.forEach((p, n) => {
        if (p.escaneada) { p.bloques.forEach((b) => aTapar.add(b)); return; }
        const r2 = quitarTextoDePagina(limpio, limpio.getPage(n));
        quitados += r2.quitados;
        if (r2.error) p.bloques.forEach((b) => aTapar.add(b));
      });
      base = await limpio.save();
      const revisado = await abrir(base.slice());
      for (let n = 0; n < paginas.length; n++) {
        const quedan = await itemsDePagina(await revisado.getPage(n + 1));
        for (const b of paginas[n].bloques) {
          if (quedan.some((it) => it.x < b.x1 && it.x + it.ancho > b.x0 && it.y < b.y1 && it.y + 0.7 * it.tam > b.y0)) aTapar.add(b);
        }
      }
    } else todos.forEach((b) => aTapar.add(b));

    // Escribir: la traducción en el lugar de cada bloque (con un recuadro del color del fondo debajo cuando hay que tapar).
    const doc = await PDFLib.PDFDocument.load(base, { updateMetadata: false });
    doc.registerFontkit(raiz.fontkit);
    const cache = new Map();
    const fuente = async (clave) => {
      if (!cache.has(clave)) cache.set(clave, doc.embedFont(fuentes[clave], { subset: true }));
      return cache.get(clave);
    };
    const asiatico = ASIATICO.has(destino);
    const sinCaber = [];
    for (let n = 0; n < paginas.length; n++) {
      const pag = doc.getPage(n);
      for (const b of paginas[n].bloques) {
        if (!b.traduccion) continue;
        const clave = asiatico ? (b.negrita && fuentes.cjkNegrita ? 'cjkNegrita' : 'cjk') : `${b.serifa ? 'serifa' : 'sans'}${b.negrita ? 'Negrita' : ''}`;
        const f = await fuente(fuentes[clave] ? clave : (asiatico ? 'cjk' : 'sans'));
        // Letras que la fuente no tiene (signos raros): fuera, para que pdf-lib no falle.
        if (!f._letras) f._letras = new Set(f.getCharacterSet());
        const texto = f._letras.size ? [...b.traduccion].filter((c) => /\s/.test(c) || f._letras.has(c.codePointAt(0))).join('') : b.traduccion;
        const ancho = b.x1 - b.x0, altoB = b.y1 - b.y0;
        // El coreano separa palabras con espacios: se corta por palabras, como el español.
        const ac = acomodar(texto, { ancho, alto: altoB, tam: b.tam, interlinea: b.interlinea, medir: (s, t) => f.widthOfTextAtSize(s, t),
          asiatico: asiatico && destino !== 'ko' });
        if (!ac.cabe) sinCaber.push(b);
        // El recuadro tapa el original (con un margen chico) y también lo que ocupe de más la traducción que no cupo.
        const extra = ac.cabe ? 0 : Math.max(0, (ac.lineas.length - 1) * ac.interlinea + 1.1 * ac.tam - altoB);
        if (aTapar.has(b) || extra) {
          pag.drawRectangle({ x: b.x0 - 1, y: b.y0 - 1 - extra, width: ancho + 2, height: altoB + 2 + extra, color: PDFLib.rgb(...b.fondo) });
        }
        ac.lineas.forEach((l, i) => {
          pag.drawText(l, { x: b.x0, y: b.y1 - 0.85 * ac.tam - i * ac.interlinea, size: ac.tam, font: f, color: PDFLib.rgb(...b.letra) });
        });
      }
    }
    const salida = await doc.save();
    const texto = paginas.map((p) => p.bloques.map((b) => b.traduccion).filter(Boolean).join('\n\n')).join('\n\n\f\n\n');
    return { bytes: salida, texto, paginas, estadistica: { ...r.estadistica, bloques: todos.length, sinCaber: sinCaber.length,
      escaneadas: paginas.filter((p) => p.escaneada).length, modo: quitar ? 'quitar' : 'tapar', tapados: aTapar.size, textosQuitados: quitados } };
  }

  // Los bloques traducidos como documento para el escritor de Word del sitio (EscritorDocx): títulos por tamaño de letra.
  function comoDocx(paginas) {
    const tams = paginas.flatMap((p) => p.bloques.map((b) => b.tam)).sort((a, b) => a - b);
    const normal = tams[Math.floor(tams.length / 2)] || 11;
    const hijos = [];
    paginas.forEach((p, n) => {
      p.bloques.forEach((b, i) => {
        if (!b.traduccion) return;
        const titulo = b.tam > 1.5 * normal ? 1 : b.tam > 1.2 * normal ? 2 : undefined;
        hijos.push({ tipo: 'parrafo', titulo, saltoAntes: n > 0 && i === 0 ? true : undefined, hijos: [{ texto: b.traduccion, negrita: b.negrita && !titulo }] });
      });
    });
    return { hijos };
  }

  raiz.TraductorPdf = Object.assign(raiz.TraductorPdf || {}, { bloques, acomodar, traducir, comoDocx,
    _nucleo: { lineas, unirLinea, cortar, colores, itemsDePagina, quitarTextoDeContenido } });
})(typeof self !== 'undefined' ? self : globalThis);

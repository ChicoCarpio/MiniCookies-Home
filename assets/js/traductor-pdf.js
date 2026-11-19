/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Traductor de documentos · PDF. Un PDF no tiene párrafos: tiene trozos de texto sueltos con su posición. Acá se arman los párrafos, se
// traducen y se escribe la traducción en el mismo lugar, tapando el original con un recuadro del color del fondo y achicando la letra si
// no entra. Una página escaneada (sin texto) pasa antes por Librarius (OCR) y sigue igual. También se puede sacar como Word o texto.
//
// Parte sin navegador (la prueban las pruebas en Node):
//   TraductorPdf.bloques(items, { reglas })   items = [{ texto, x, y, ancho, tam, negrita? }]  (x, y: línea base, PDF: y hacia arriba)
//        reglas = { v: [{ x, y0, y1 }], h: [{ y, x0, x1 }] }: las rayas verticales y horizontales dibujadas en la página (bordes de tabla). Un
//        texto no se junta con el de otra celda: ninguna línea cruza una raya vertical y ningún párrafo cruza una horizontal.
//        -> [{ texto, x0, x1, y0, y1, tam, interlinea, lineas, piezas, celda? }]   un bloque por párrafo, en orden de lectura;
//        celda = { x1?, y0? }: hasta dónde llega la celda (la raya de la derecha y la de abajo), para que la traducción use todo su lugar
//   TraductorPdf.esTraducible(texto)   false si es solo números, un correo, una dirección o un código (eso no se traduce)
//   TraductorPdf.acomodar(texto, { ancho, alto, tam, interlinea, medir, asiatico, pasos })   medir(texto, tam) -> ancho; pasos = cuántas veces probar con 5 % menos (0: solo el tamaño original)
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
  const HUECO_COLUMNA = 1.6;   // un hueco de más de 1,6 veces el tamaño separa celdas o columnas (el espacio entre palabras es 0,25-0,4)

  // ¿Hay una raya vertical entre xa y xb, a la altura de esta línea?
  function rayaVerticalEntre(reglas, xa, xb, y, tam) {
    if (!reglas) return false;
    const ym = y + 0.3 * tam;
    return reglas.v.some((r) => r.x > xa - 0.5 && r.x < xb + 0.5 && r.y0 <= ym && r.y1 >= ym);
  }
  // ¿Hay una raya horizontal entre una línea de arriba (u) y una de abajo (l), sobre el texto de las dos?
  function rayaHorizontalEntre(reglas, u, l) {
    if (!reglas) return false;
    const xa = Math.max(u.x0, l.x0), xb = Math.min(u.x1, l.x1);
    return reglas.h.some((r) => r.y < u.y - 0.35 * u.tam && r.y > l.y + 0.5 * l.tam && r.x0 < xb && r.x1 > xa);
  }

  function lineas(items, reglas) {
    const alturas = [];
    for (const it of items.filter((i) => i.texto && i.texto.trim() && i.tam > 0).sort((a, b) => b.y - a.y)) {
      const g = alturas.find((x) => Math.abs(x.y - it.y) < 0.35 * Math.max(x.tam, it.tam));
      if (g) { g.items.push(it); g.tam = Math.max(g.tam, it.tam); } else alturas.push({ y: it.y, tam: it.tam, items: [it] });
    }
    const ordenados = alturas.flatMap((g) => g.items.sort((a, b) => a.x - b.x).map((it) => ({ ...it, yLinea: g.y })));
    const ls = [];
    for (const it of ordenados) {
      const l = ls.find((x) => x.y === it.yLinea && it.x >= x.x1 - 0.5 * it.tam && it.x - x.x1 < HUECO_COLUMNA * it.tam && !rayaVerticalEntre(reglas, x.x1, it.x, x.y, x.tam));
      if (l) {
        const hueco = it.x - l.x1;
        const pegar = hueco < 0.15 * it.tam || /\s$/.test(l.texto) || /^\s/.test(it.texto) || (CJK.test(l.texto.slice(-1)) && CJK.test(it.texto[0]));
        l.texto += (pegar ? '' : ' ') + it.texto;
        l.x1 = Math.max(l.x1, it.x + it.ancho);
        l.tam = Math.max(l.tam, it.tam);
        l.negritas += it.negrita ? it.texto.length : 0; l.letras += it.texto.length;
        l.piezas.push({ texto: it.texto, x: it.x, y: it.y, tam: it.tam, cursiva: it.cursiva, mono: it.mono, serifa: it.serifa });
      } else {
        ls.push({ texto: it.texto, x0: it.x, x1: it.x + it.ancho, y: it.yLinea, tam: it.tam, negritas: it.negrita ? it.texto.length : 0, letras: it.texto.length,
          piezas: [{ texto: it.texto, x: it.x, y: it.y, tam: it.tam, cursiva: it.cursiva, mono: it.mono, serifa: it.serifa }] });
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

  function bloques(items, { reglas } = {}) {
    const bs = [];
    for (const l of lineas(items, reglas)) {
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
        if (rayaHorizontalEntre(reglas, u, l)) return false;
        if (FIN_ORACION.test(u.texto) && (u.x1 < x.x0 + 0.75 * (x.x1 - x.x0) || l.x0 > x.x0 + 1.2 * x.tam)) return false;
        return true;
      });
      if (b) {
        b.interlinea = b.lineas.length === 1 ? b.ult.y - l.y : (b.interlinea * (b.lineas.length - 1) + (b.ult.y - l.y)) / b.lineas.length;
        b.texto = unirLinea(b.texto, l.texto);
        b.lineas.push(l); b.ult = l; b.piezas.push(...l.piezas);
        b.x0 = Math.min(b.x0, l.x0); b.x1 = Math.max(b.x1, l.x1);
        b.y0 = l.y - 0.25 * l.tam;
      } else {
        // Las líneas por encima de esta que ya no pueden seguir (quedaron arriba de un hueco grande) se cierran.
        for (const x of bs) if (!x.cerrado && x.ult.y - l.y > 1.8 * x.tam && !(l.x0 > x.x1 || l.x1 < x.x0)) x.cerrado = true;
        bs.push({ texto: l.texto, lineas: [l], ult: l, piezas: l.piezas.slice(), x0: l.x0, x1: l.x1, y1: l.y + 0.85 * l.tam, y0: l.y - 0.25 * l.tam, tam: l.tam, interlinea: 1.2 * l.tam, negrita });
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
    return ordenados.map(({ texto, x0, x1, y0, y1, tam, interlinea, lineas: ls, negrita, piezas }) =>
      ({ texto, x0, x1, y0, y1, tam, interlinea, negrita, lineas: ls.length, piezas, celda: celdaDe(reglas, { x0, x1, y0, y1, tam }),
        ...estiloDe(piezas) }));
  }

  // Cursiva, monoespaciada y serifa de un bloque: lo que tenga la mayoría de sus letras.
  function estiloDe(piezas) {
    const total = piezas.reduce((t, p) => t + p.texto.length, 0) || 1;
    const mayoria = (k) => piezas.reduce((t, p) => t + (p[k] ? p.texto.length : 0), 0) > total / 2;
    return { cursiva: mayoria('cursiva'), mono: mayoria('mono'), serifa: mayoria('serifa') };
  }

  // Lugar libre alrededor de un bloque (sin otro texto ni raya de tabla): hasta dónde puede crecer la traducción hacia la derecha y hacia abajo.
  function espacioLibre(b, otros, caja) {
    let derecha = caja ? caja[2] - 18 : Infinity, abajo = caja ? caja[1] + 18 : -Infinity;
    if (b.celda && b.celda.x1 !== undefined) derecha = Math.min(derecha, b.celda.x1 - 1.5);
    if (b.celda && b.celda.y0 !== undefined) abajo = Math.max(abajo, b.celda.y0 + 1);
    for (const o of otros) {
      if (o === b) continue;
      if (o.y0 < b.y1 && o.y1 > b.y0 && o.x0 >= b.x1 - 1) derecha = Math.min(derecha, o.x0 - 3);
      if (o.x0 < b.x1 && o.x1 > b.x0 && o.y1 <= b.y0 + 1) abajo = Math.max(abajo, o.y1 + 1);
    }
    return { derecha, abajo };
  }

  // Hasta dónde llega la celda de un bloque: la primera raya vertical a su derecha y la primera horizontal debajo (si la hay).
  function celdaDe(reglas, b) {
    if (!reglas) return undefined;
    const ym = (b.y0 + b.y1) / 2, xm = (b.x0 + b.x1) / 2;
    let x1 = Infinity, y0 = -Infinity;
    for (const r of reglas.v) if (r.x >= b.x1 - 1 && r.x < x1 && r.y0 <= ym && r.y1 >= ym) x1 = r.x;
    for (const r of reglas.h) if (r.y <= b.y0 + 0.3 * b.tam && r.y > y0 && r.x0 <= xm && r.x1 >= xm) y0 = r.y;
    return x1 === Infinity && y0 === -Infinity ? undefined : { x1: x1 === Infinity ? undefined : x1, y0: y0 === -Infinity ? undefined : y0 };
  }

  // Lo que no se traduce: solo números o signos, correos, direcciones web, rutas y códigos (letras y cifras juntas, como «ICEPERBARAM2» o «A4»).
  function identificador(t) {
    return !/\p{L}/u.test(t) || /@|^(?:https?:\/\/|www\.)/i.test(t) || /^[\p{L}\d_-]+(?:[./\\_][\p{L}\d_-]+)+$/u.test(t) || (/\d/.test(t) && /\p{L}/u.test(t));
  }
  function esTraducible(texto) {
    const trozos = String(texto || '').trim().split(/\s+/).filter(Boolean);
    return trozos.length > 0 && !trozos.every(identificador);
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

  function acomodar(texto, { ancho, alto, tam, interlinea, medir, asiatico, pasos = 14 }) {
    const proporcion = (interlinea || 1.2 * tam) / tam;
    let t = tam;
    for (let i = 0; i <= pasos; i++) {
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
        cursiva: /italic|oblique|[-,]it\b|ital\b/i.test(nombre), mono: /mono|courier|consolas|menlo|lucida ?console|typewriter/i.test(nombre + ' ' + familia),
        serifa: /serif(?!-?sans)|times|roman|georgia|garamond|cambria|book|minion/i.test(nombre + ' ' + familia) && !/sans/i.test(nombre + ' ' + familia) });
    }
    return items;
  }

  // Las rayas de la página (bordes de tablas, líneas de gráficos): del dibujo del PDF. Cada trazo recto horizontal o vertical de un contorno,
  // y los rectángulos rellenos muy finos (una raya hecha con un rectángulo). Coordenadas del PDF (y hacia arriba).
  const TRAZAR = new Set([20, 21, 24, 25, 26, 27]), RELLENAR = new Set([22, 23, 24, 25, 26, 27]);
  async function reglasDePagina(pagina) {
    const reglas = { v: [], h: [] };
    let lista;
    try { lista = await pagina.getOperatorList(); } catch (err) { return reglas; }
    const O = raiz.pdfjsLib && raiz.pdfjsLib.OPS;
    if (!O) return reglas;
    const mul = (m1, m2) => [m1[0] * m2[0] + m1[1] * m2[2], m1[0] * m2[1] + m1[1] * m2[3], m1[2] * m2[0] + m1[3] * m2[2], m1[2] * m2[1] + m1[3] * m2[3],
      m1[4] * m2[0] + m1[5] * m2[2] + m2[4], m1[4] * m2[1] + m1[5] * m2[3] + m2[5]];
    let ctm = [1, 0, 0, 1, 0, 0];
    const pila = [];
    const MIN = 6, MAX = 3000;
    const recta = (a, b) => {
      const dx = Math.abs(a[0] - b[0]), dy = Math.abs(a[1] - b[1]);
      if (dx < 0.5 && dy >= MIN) reglas.v.push({ x: (a[0] + b[0]) / 2, y0: Math.min(a[1], b[1]), y1: Math.max(a[1], b[1]) });
      else if (dy < 0.5 && dx >= MIN) reglas.h.push({ y: (a[1] + b[1]) / 2, x0: Math.min(a[0], b[0]), x1: Math.max(a[0], b[0]) });
    };
    for (let i = 0; i < lista.fnArray.length && reglas.v.length + reglas.h.length < MAX; i++) {
      const f = lista.fnArray[i], args = lista.argsArray[i];
      if (f === O.save) pila.push(ctm);
      else if (f === O.restore) { if (pila.length) ctm = pila.pop(); }
      else if (f === O.transform) ctm = mul(args, ctm);
      else if (f === O.constructPath) {
        const pintar = args[0];
        const datos = args[1] && args[1].length && typeof args[1][0] === 'object' ? args[1][0] : args[1];
        if (!datos || (!TRAZAR.has(pintar) && !RELLENAR.has(pintar))) continue;
        const pt = (x, y) => [x * ctm[0] + y * ctm[2] + ctm[4], x * ctm[1] + y * ctm[3] + ctm[5]];
        let sub = [], cerrada = false, curva = false;
        const cerrarSub = () => {
          if (sub.length >= 2 && !curva) {
            const n = sub.length;
            if (TRAZAR.has(pintar)) {
              for (let k = 1; k < n; k++) recta(sub[k - 1], sub[k]);
              if (cerrada && n > 2) recta(sub[n - 1], sub[0]);
            }
            if (RELLENAR.has(pintar) && (n === 4 || n === 5)) {
              const xs = sub.map((q) => q[0]), ys = sub.map((q) => q[1]);
              const x0 = Math.min(...xs), x1 = Math.max(...xs), y0 = Math.min(...ys), y1 = Math.max(...ys);
              const w = x1 - x0, h = y1 - y0;
              if (Math.min(w, h) <= 3 && Math.max(w, h) >= MIN) {
                if (w >= h) recta([x0, (y0 + y1) / 2], [x1, (y0 + y1) / 2]); else recta([(x0 + x1) / 2, y0], [(x0 + x1) / 2, y1]);
              }
            }
          }
          sub = []; cerrada = false; curva = false;
        };
        for (let k = 0; k < datos.length;) {
          const op = datos[k++];
          if (op === 0) { cerrarSub(); sub = [pt(datos[k], datos[k + 1])]; k += 2; }
          else if (op === 1) { sub.push(pt(datos[k], datos[k + 1])); k += 2; }
          else if (op === 2) { curva = true; k += 6; }
          else if (op === 3) { curva = true; k += 4; }
          else if (op === 4) cerrada = true;
          else break;
        }
        cerrarSub();
      }
    }
    return reglas;
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
      const bs = bloques(items, { reglas: escaneada ? undefined : await reglasDePagina(pagina) });
      for (const b of bs) b.conservar = !esTraducible(b.texto);
      const aPixel = (x, y) => vista.convertToViewportPoint(x, y);
      for (const b of bs) Object.assign(b, colores(ctx, aPixel, b));
      paginas.push({ bloques: bs, escaneada, caja: pagina.view });
      lienzo.width = lienzo.height = 0;
    }
    const todos = paginas.flatMap((p) => p.bloques);
    // Lo que no se traduce (números, correos, códigos) se deja como está.
    const aTraducir = todos.filter((b) => !b.conservar);
    const r = await raiz.TraductorDocumentos.traducirUnidades(aTraducir.map((b) => [{ texto: b.texto, formato: '' }]), {
      traducir: (segs, d, avisar) => traductor(segs, d, { progreso: avisar }), destino, progreso,
    });
    aTraducir.forEach((b, i) => { b.traduccion = r[i].map((t) => t.texto).join('').trim(); });
    todos.filter((b) => b.conservar).forEach((b) => { b.traduccion = b.texto; });

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
        if (p.escaneada) { p.bloques.filter((b) => !b.conservar).forEach((b) => aTapar.add(b)); return; }
        const r2 = quitarTextoDePagina(limpio, limpio.getPage(n));
        quitados += r2.quitados;
        if (r2.error) p.bloques.forEach((b) => { if (!b.conservar) aTapar.add(b); });
        else p.bloques.forEach((b) => { if (b.conservar) b.redibujar = true; });
      });
      base = await limpio.save();
      const revisado = await abrir(base.slice());
      for (let n = 0; n < paginas.length; n++) {
        const quedan = await itemsDePagina(await revisado.getPage(n + 1));
        for (const b of paginas[n].bloques) {
          if (quedan.some((it) => it.x < b.x1 && it.x + it.ancho > b.x0 && it.y < b.y1 && it.y + 0.7 * it.tam > b.y0)) {
            if (b.conservar) b.redibujar = false; else aTapar.add(b);
          }
        }
      }
    } else todos.forEach((b) => { if (!b.conservar) aTapar.add(b); });

    // Escribir: la traducción en el lugar de cada bloque (con un recuadro del color del fondo debajo cuando hay que tapar).
    const doc = await PDFLib.PDFDocument.load(base, { updateMetadata: false });
    doc.registerFontkit(raiz.fontkit);
    const cache = new Map();
    const fuente = async (clave) => {
      if (!cache.has(clave)) cache.set(clave, (async () => doc.embedFont(typeof fuentes[clave] === 'function' ? await fuentes[clave]() : fuentes[clave], { subset: true }))());
      return cache.get(clave);
    };
    const asiatico = ASIATICO.has(destino);
    const sinCaber = [];
    for (let n = 0; n < paginas.length; n++) {
      const pag = doc.getPage(n);
      for (const b of paginas[n].bloques) {
        if (!b.traduccion) continue;
        if (b.conservar && !b.redibujar) continue;
        // La fuente más parecida a la original: monoespaciada, con o sin serifa, negrita y cursiva (la que falte se reemplaza por la más cercana).
        const base = b.mono && fuentes.mono ? 'mono' : b.serifa ? 'serifa' : 'sans';
        const nb = b.negrita ? 'Negrita' : '', cu = b.cursiva ? 'Cursiva' : '';
        const clave = asiatico ? (b.negrita && fuentes.cjkNegrita ? 'cjkNegrita' : 'cjk') : [base + nb + cu, base + nb, base, 'sans'].find((k) => fuentes[k]);
        const f = await fuente(clave || 'cjk');
        // Letras que la fuente no tiene (signos raros): fuera, para que pdf-lib no falle.
        if (!f._letras) f._letras = new Set(f.getCharacterSet());
        const texto = f._letras.size ? [...b.traduccion].filter((c) => /\s/.test(c) || f._letras.has(c.codePointAt(0))).join('') : b.traduccion;
        if (b.conservar) {
          // Números, correos y códigos: se vuelven a poner donde estaban, pieza por pieza (el texto original se quitó).
          for (const pz of b.piezas) {
            const t = f._letras.size ? [...pz.texto].filter((c) => /\s/.test(c) || f._letras.has(c.codePointAt(0))).join('') : pz.texto;
            if (t.trim()) pag.drawText(t, { x: pz.x, y: pz.y, size: pz.tam, font: f, color: PDFLib.rgb(...b.letra) });
          }
          continue;
        }
        const ancho = b.x1 - b.x0, altoB = b.y1 - b.y0;
        // En una tabla la traducción puede usar toda la celda (hasta la raya de la derecha y la de abajo), no solo lo que ocupaba el original.
        const cajaAncho = b.celda && b.celda.x1 !== undefined ? Math.max(ancho, b.celda.x1 - 1.5 - b.x0) : ancho;
        const cajaAlto = b.celda && b.celda.y0 !== undefined ? Math.max(altoB, b.y1 - b.celda.y0 - 1) : altoB;
        // Se mantiene el tamaño de la letra: primero en el lugar del original (o su celda); si no entra, en el lugar libre de alrededor (un poco
        // más ancho, hasta el siguiente texto o raya, y hacia abajo); solo si todavía no entra se achica la letra. El coreano separa palabras con
        // espacios: se corta por palabras, como el español.
        const dat = { tam: b.tam, interlinea: b.interlinea, medir: (s, t) => f.widthOfTextAtSize(s, t), asiatico: asiatico && destino !== 'ko' };
        const libre = espacioLibre(b, paginas[n].bloques, paginas[n].caja);
        const anchoL = Math.max(cajaAncho, Math.min(libre.derecha - b.x0, Math.max(1.5 * ancho, ancho + 6 * b.tam)));
        const altoL = Math.max(cajaAlto, b.y1 - libre.abajo);
        // Cada intento crece un poco más: el lugar del original, el ancho de la celda, el lugar libre a la derecha y, por último, más líneas.
        let ac = null;
        for (const [w, h] of [[ancho, altoB], [cajaAncho, altoB], [anchoL, altoB], [anchoL, altoL]]) {
          ac = acomodar(texto, { ...dat, ancho: w, alto: h, pasos: 0 });
          if (ac.cabe) break;
        }
        if (!ac.cabe) ac = acomodar(texto, { ...dat, ancho: anchoL, alto: altoL });
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
    marcarIA(doc);
    const salida = await doc.save();
    const texto = paginas.map((p) => p.bloques.map((b) => b.traduccion).filter(Boolean).join('\n\n')).join('\n\n\f\n\n');
    return { bytes: salida, texto, paginas, estadistica: { ...r.estadistica, bloques: todos.length, sinCaber: sinCaber.length,
      escaneadas: paginas.filter((p) => p.escaneada).length, modo: quitar ? 'quitar' : 'tapar', tapados: aTapar.size, textosQuitados: quitados } };
  }

  // Marca de contenido generado con IA, solo en los metadatos del PDF (Asunto y Palabras clave), como en Word y EPUB.
  function marcarIA(doc) {
    doc.setSubject('Translated with AI (Interpres Phone V1), minicookies.net');
    doc.setKeywords(['AI-translated', 'Interpres Phone V1', 'Mini Cookies']);
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

  raiz.TraductorPdf = Object.assign(raiz.TraductorPdf || {}, { bloques, acomodar, esTraducible, reglasDePagina, espacioLibre, traducir, comoDocx, marcarIA,
    _nucleo: { lineas, unirLinea, cortar, colores, itemsDePagina, quitarTextoDeContenido } });
})(typeof self !== 'undefined' ? self : globalThis);

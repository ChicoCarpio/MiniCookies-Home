/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Objetos de una hoja de Excel (etapa 6): imágenes, formas, líneas, cuadros de texto, grupos y gráficos del dibujo de la hoja
// (xl/drawings/drawingN.xml). Su contenido es DrawingML igual que el de una diapositiva, así que se resuelve y se dibuja con el motor de PowerPoint
// (pptx-modelo.js: PptxModelo.dibujoExcel, y pptx-dibujo.js: PptxDibujo.dibujarElementos); este módulo solo pone cada objeto en su sitio según el
// ancla (twoCellAnchor, oneCellAnchor o absoluteAnchor: hoja.anclas, que lee xlsx-modelo.js) y la geometría de la hoja. Necesita DOM (el lector de DrawingML lo usa): sin él no hay objetos.
// Requiere xlsx-modelo.js, pptx-xml.js, pptx-modelo.js, pptx-dibujo.js (y lo que ellos piden: pptx-geometria.js, pptx-texto.js, pptx-grafico*.js).
//
// Uso:
//   await XlsxObjetos.preparar(libro, hoja, lienzo);      // una vez por hoja: lee los objetos, carga imágenes y fuentes (lienzo: el de PDF, si se imprime)
//   XlsxObjetos.dibujar(ctx, hoja, g, { f0, c0, f1, c1 }, { x, y, lienzo });  // síncrono; lo llama XlsxDibujo.dibujarRango

(function () {
  const EMU_PX = 9525;           // un píxel de pantalla a 96 dpi (y una unidad "u" del PDF a escala 100 %)
  const MAX_OBJETOS = 5000;
  const MAX_IMAGENES_CELDA = 2000;

  async function preparar(libro, hoja, lienzo) {
    if (hoja.objetos) return hoja.objetos;
    hoja.objetos = [];
    hoja._objetosImagenes = new Map();
    if (typeof PptxModelo === 'undefined' || typeof PptxDibujo === 'undefined' || typeof DOMParser === 'undefined' || !hoja.partes || !hoja.partes.dibujo) return hoja.objetos;
    let dib;
    try { dib = await PptxModelo.dibujoExcel(libro.zip, hoja.partes.dibujo, libro.rutaTema); } catch (e) { return hoja.objetos; }
    if (!dib) return hoja.objetos;
    const todos = [];
    for (let k = 0; k < dib.anclajes.length && k < MAX_OBJETOS; k++) {
      const ancla = dib.anclajes[k];
      const motivos = [];
      let elementos = [];
      try { elementos = await dib.elementos(ancla, motivos); } catch (e) { continue; }
      const a = (hoja.anclas || [])[k];
      if (elementos.length && a && !a.oculto) { hoja.objetos.push({ ancla: a, elementos }); todos.push(...elementos); }
    }
    try { hoja._objetosImagenes = await PptxDibujo.precargar((r) => dib.imagen(r), null, todos, lienzo || null); } catch (e) { /* sin imágenes */ }
    return hoja.objetos;
  }

  // Las imágenes dentro de las celdas (hoja.imagenesCelda: celda → ruta): se cargan una vez por ruta y se dibujan con el mismo motor.
  async function prepararCeldas(libro, hoja, lienzo) {
    if (hoja._celdaImagenes) return hoja._celdaImagenes;
    hoja._celdaImagenes = new Map();
    if (!hoja.imagenesCelda || !hoja.imagenesCelda.size || typeof PptxModelo === 'undefined' || typeof PptxDibujo === 'undefined' || !PptxModelo.cargadorImagenes) return hoja._celdaImagenes;
    const rutas = [...new Set(hoja.imagenesCelda.values())].slice(0, MAX_IMAGENES_CELDA);
    const elementos = rutas.map((ruta) => ({ tipo: 'imagen', ruta, xfrm: { x: 0, y: 0, w: 0, h: 0, rot: 0, flipH: false, flipV: false }, geometria: null, relleno: null, linea: null, sombra: null,
      imagen: { tipo: 'imagen', ruta, recorte: null, efectos: [], alfa: 1, rotarConForma: true } }));
    try {
      const imagenes = await PptxDibujo.precargar(PptxModelo.cargadorImagenes(libro.zip), null, elementos, lienzo || null);
      elementos.forEach((el) => { const im = imagenes.get(el.ruta); if (im && im.ancho > 0 && im.alto > 0) hoja._celdaImagenes.set(el.ruta, { el, im, imagenes }); });
    } catch (e) { /* sin imágenes */ }
    return hoja._celdaImagenes;
  }

  // La imagen de una celda: cabe entera en la celda (sin deformarse), pegada a su borde de abajo y a su borde izquierdo (medido en el PDF de Excel).
  // x0, y0, x1, y1: la celda (o la combinada) en unidades del lienzo; px: unidades del lienzo por píxel de pantalla.
  function dibujarCelda(ctx, hoja, ruta, x0, y0, x1, y1, op) {
    const d = hoja._celdaImagenes && hoja._celdaImagenes.get(ruta);
    if (!d || typeof PptxDibujo === 'undefined') return;
    op = op || {};
    const W = x1 - x0 - (op.margenIzq || 0), H = y1 - y0;
    if (!(W > 0.5) || !(H > 0.5)) return;
    const k = Math.min(W / d.im.ancho, H / d.im.alto);
    const w = d.im.ancho * k, h = d.im.alto * k;
    const el = { ...d.el, xfrm: { ...d.el.xfrm, x: 0, y: 0, w: w * EMU_PX, h: h * EMU_PX } };
    ctx.save();
    ctx.beginPath(); ctx.rect(x0, y0, x1 - x0, y1 - y0); ctx.clip();
    ctx.translate(x0 + (op.margenIzq || 0), y1 - h);
    try { PptxDibujo.dibujarElementos(ctx, [el], d.imagenes, 1 / EMU_PX, op.lienzo || null); } catch (e) { /* una imagen que no se puede dibujar no tumba la hoja */ }
    ctx.restore();
  }

  // Dibuja los objetos que tocan el rango r, con el mismo origen que dibujarRango (op.x, op.y: dónde cae la esquina del rango).
  // En pantalla la geometría es la de la hoja. Al imprimir, `g` es la de impresión y trae `pantalla` (la de pantalla) y `escalaX`/`escalaY` (de
  // píxeles de pantalla a unidades de la página): Excel coloca cada objeto con su posición en pantalla escalada, no con las celdas impresas.
  function dibujar(ctx, hoja, g, r, op) {
    const objs = hoja.objetos;
    if (!objs || !objs.length || typeof PptxDibujo === 'undefined') return;
    op = op || {};
    const P = g.pantalla || g, ex = g.escalaX || 1, ey = g.escalaY || 1;
    const px = op.x || 0, py = op.y || 0;
    const sx0 = P.x(r.c0), sy0 = P.y(r.f0);
    const ancho = g.x(r.c1 + 1) - g.x(r.c0), alto = g.y(r.f1 + 1) - g.y(r.f0);
    for (const o of objs) {
      const q = XlsxGeometria.rectanguloAncla(o.ancla, P);
      if (!q) continue;
      const x0 = px + (q.x0 - sx0) * ex, y0 = py + (q.y0 - sy0) * ey, w = (q.x1 - q.x0) * ex, h = (q.y1 - q.y0) * ey;
      if (!(w > 0) || !(h > 0) || x0 + w < px - 1 || x0 > px + ancho + 1 || y0 + h < py - 1 || y0 > py + alto + 1) continue; // fuera del rango
      const els = o.elementos.map((e) => ({ ...e, xfrm: { ...e.xfrm, x: 0, y: 0, w: w * EMU_PX, h: h * EMU_PX } }));
      ctx.save();
      ctx.translate(x0, y0);
      try { PptxDibujo.dibujarElementos(ctx, els, hoja._objetosImagenes, 1 / EMU_PX, op.lienzo || null); } catch (e) { /* un objeto que no se puede dibujar no tumba la hoja */ }
      ctx.restore();
    }
  }

  globalThis.XlsxObjetos = { preparar, prepararCeldas, dibujar, dibujarCelda };
})();

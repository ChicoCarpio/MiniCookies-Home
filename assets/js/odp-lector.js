/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Lector de .odp propio (OpenDocument Presentation: LibreOffice, OpenOffice, Google Presentaciones, PowerPoint al guardar como OpenDocument). Lo entrega como el «contrato de
// presentación» (ver pptx-desde-contrato.js) y, con aPptx, como un .pptx, que es lo que ya saben leer, mostrar e imprimir las herramientas de PowerPoint.
//
//   await OdpLector.parece(bytes)            → true si es un .odp (zip con content.xml y mimetype application/vnd.oasis.opendocument.presentation)
//   await OdpLector.leer(bytes)              → { contrato, mensajes: [clave…] }
//   await OdpLector.aPptx(bytes)             → { bytes, mensajes }
//
// Lanza Error (con .codigo = 'cifrado' si el .odp tiene contraseña) con un mensaje listo para mostrar.
// Se llevan: diapositivas con su página maestra (fondo y formas), marcos de texto y marcadores de posición, formas (rectángulo, elipse, línea, polígono, trazados, formas predefinidas),
// relleno (color, degradado, imagen), línea, sombra, giro, grupos, texto con su formato (fuente, tamaño, color, negrita, cursiva, subrayado, tachado, superíndice, subíndice,
// alineación, sangrías, espaciado, interlineado, viñetas y numeración, enlaces), imágenes, tablas, notas del orador, diapositivas ocultas, encabezado/pie/número/fecha.
// No se llevan: animaciones y transiciones, sonido y vídeo, objetos incrustados (gráficos, fórmulas), conectores con quiebres, formas con ecuaciones.
// Requiere zip-propio.js y xml-ligero.js (y pptx-desde-contrato.js para aPptx).
(function () {
  'use strict';
  const MAX_PROFUNDIDAD = 64;
  const MAX_ELEMENTOS = 3000000;
  const MAX_TEXTO_XML = 256 * 1024 * 1024;
  const MAX_IMAGENES_BYTES = 256 * 1024 * 1024;
  const MAX_IMAGENES = 20000;
  const MAX_FORMAS = 400000;
  const MAX_DIAPOSITIVAS = 20000;
  const MAX_CARACTERES = 64 * 1024 * 1024;
  const MAX_CELDAS = 200000;
  const MAX_PUNTOS = 4000000;
  const sinProto = (o) => Object.assign(Object.create(null), o);
  const acotar = (x, min, max) => Math.min(Math.max(x, min), max);

  // ───────── XML ─────────
  const arbol = (xml) => XmlLigero.arbol(xml, MAX_ELEMENTOS, MAX_PROFUNDIDAD * 4, true);
  const hijos = (e, n) => (e ? e.h.filter((x) => typeof x !== 'string' && (n === undefined || x.n === n)) : []);
  const hijo = (e, n) => (e ? e.h.find((x) => typeof x !== 'string' && x.n === n) || null : null);
  const atr = (e, n) => (e && e.a[n] !== undefined ? e.a[n] : undefined);
  const textoDe = (e) => (e ? e.h.map((x) => (typeof x === 'string' ? x : textoDe(x))).join('') : '');

  // ───────── Medidas ─────────
  const EMU = { cm: 360000, mm: 36000, in: 914400, pt: 12700, px: 9525, pc: 152400 };
  // "2.5cm", "1in", "12pt", "50%" (de base) → EMU (null si no es una medida)
  function emu(t, base) {
    const m = /^\s*(-?[\d.]+(?:e-?\d+)?)\s*(cm|mm|in|pt|px|pc|%)?\s*$/i.exec(t || '');
    if (!m) return null;
    const n = parseFloat(m[1]);
    if (!Number.isFinite(n)) return null;
    const u = (m[2] || 'cm').toLowerCase();
    if (u === '%') return base == null ? null : Math.round(base * n / 100);
    return Math.round(n * EMU[u]);
  }
  const puntos = (t, base) => { const m = /^\s*(-?[\d.]+)\s*(pt|%|cm|mm|in|px)?\s*$/i.exec(t || ''); if (!m) return null; const n = parseFloat(m[1]); if (!Number.isFinite(n)) return null; const u = (m[2] || 'pt').toLowerCase(); return u === '%' ? (base == null ? null : base * n / 100) : u === 'pt' ? n : n * EMU[u] / 12700; };
  const color = (c) => (/^#[0-9a-fA-F]{6}$/.test(c || '') ? c.slice(1).toUpperCase() : null);
  const porcentaje = (t, def) => { const m = /^\s*(-?[\d.]+)\s*%?\s*$/.exec(t || ''); return m ? parseFloat(m[1]) : def; };

  // ───────── Estilos ─────────
  const GRUPOS = { 'style:graphic-properties': 'g', 'loext:graphic-properties': 'g', 'style:text-properties': 't', 'style:paragraph-properties': 'p', 'style:drawing-page-properties': 'dp', 'style:table-cell-properties': 'c', 'style:table-column-properties': 'col', 'style:table-row-properties': 'fila', 'style:table-properties': 'tabla', 'style:chart-properties': 'ch' };
  const nuevoEstilo = () => ({ g: {}, t: {}, p: {}, dp: {}, c: {}, col: {}, fila: {}, tabla: {} });
  const unir = (a, b) => {
    const r = {};
    for (const k of Object.keys(a)) r[k] = Object.assign({}, a[k], b[k]);
    return r;
  };

  class Estilos {
    constructor() {
      this.auto = { c: new Map(), m: new Map() };      // automáticos de content.xml y de styles.xml
      this.comun = new Map();                            // estilos con nombre (office:styles de styles.xml y content.xml)
      this.predet = new Map();                           // style:default-style por familia
      this.fuentes = new Map();
      this.gradientes = new Map();
      this.guiones = new Map();
      this.rellenosImagen = new Map();
      this.listas = { c: new Map(), m: new Map(), comun: new Map() };
      this.paginas = new Map();                          // style:page-layout → { ancho, alto }
      this.maestras = new Map();                         // style:master-page → { nodo, estilo, pagina }
      this.cache = new Map();
    }
    leerFuentes(raiz) {
      for (const d of hijos(hijo(raiz, 'office:font-face-decls'), 'style:font-face')) {
        const n = atr(d, 'style:name');
        const f = (atr(d, 'svg:font-family') || n || '').replace(/^['"]|['"]$/g, '');
        if (n && f) this.fuentes.set(n, f);
      }
    }
    leerDefinicionesDe(contenedor) {
      for (const e of hijos(contenedor)) {
        const n = atr(e, 'draw:name');
        if (e.n === 'draw:gradient' && n) this.gradientes.set(n, e);
        else if (e.n === 'draw:stroke-dash' && n) this.guiones.set(n, e);
        else if (e.n === 'draw:fill-image' && n) this.rellenosImagen.set(n, e);
        else if (e.n === 'text:list-style') { const ln = atr(e, 'style:name'); if (ln) this.listas.comun.set(ln, e); }
      }
    }
    leerEstilosDe(contenedor, ambito, esAuto) {
      for (const e of hijos(contenedor)) {
        if (e.n === 'text:list-style') { const ln = atr(e, 'style:name'); if (ln) (esAuto ? this.listas[ambito] : this.listas.comun).set(ln, e); continue; }
        if (e.n === 'style:page-layout') {
          const p = hijo(e, 'style:page-layout-properties');
          const n = atr(e, 'style:name');
          if (n && p) this.paginas.set(n, { ancho: emu(atr(p, 'fo:page-width')), alto: emu(atr(p, 'fo:page-height')) });
          continue;
        }
        if (e.n !== 'style:style' && e.n !== 'style:default-style') continue;
        const familia = atr(e, 'style:family');
        if (!familia) continue;
        const est = nuevoEstilo();
        for (const c of hijos(e)) {
          const g = GRUPOS[c.n];
          if (g) Object.assign(est[g], c.a);
          // las listas que van dentro de un estilo de presentación (los niveles de esquema)
          if (c.n === 'style:graphic-properties') { const ls = hijo(c, 'text:list-style'); if (ls) est.lista = ls; }
        }
        est.padre = atr(e, 'style:parent-style-name');
        est.listaNombre = atr(e, 'style:list-style-name') || est.p['style:list-style-name'];
        if (e.n === 'style:default-style') { this.predet.set(familia, est); continue; }
        const nombre = atr(e, 'style:name');
        if (!nombre) continue;
        (esAuto ? this.auto[ambito] : this.comun).set(familia + '|' + nombre, est);
      }
    }
    leerMaestras(contenedor) {
      for (const e of hijos(contenedor, 'style:master-page')) {
        const n = atr(e, 'style:name');
        if (n) this.maestras.set(n, e);
      }
    }
    // Un estilo con toda su herencia (los de la familia por defecto debajo de todo)
    resolver(familia, nombre, ambito, sinBase) {
      const clave = ambito + '|' + familia + '|' + nombre + (sinBase ? '|propio' : '');
      if (this.cache.has(clave)) return this.cache.get(clave);
      const base = sinBase ? nuevoEstilo() : this.predet.get(familia) || nuevoEstilo();
      let r = unir(nuevoEstilo(), base);
      // un estilo gráfico (o ninguno) hereda del estilo «standard» (el predeterminado de dibujo)
      if (!nombre && familia === 'graphic' && !sinBase && this.comun.has('graphic|standard')) r = unir(r, this.comun.get('graphic|standard'));
      if (nombre) {
        const cadena = [];
        let n = nombre, k = 0;
        const vistos = new Set();
        while (n && k++ < 32 && !vistos.has(n)) {
          vistos.add(n);
          const e = this.auto[ambito].get(familia + '|' + n) || this.comun.get(familia + '|' + n);
          if (!e) break;
          cadena.unshift(e);
          n = e.padre;
        }
        // un estilo gráfico sin padre hereda del estilo «standard» (el predeterminado de dibujo)
        if (familia === 'graphic' && !sinBase && nombre !== 'standard' && !vistos.has('standard') && this.comun.has('graphic|standard')) cadena.unshift(this.comun.get('graphic|standard'));
        for (const e of cadena) {
          const antes = r;
          r = unir(r, e);
          // el tamaño de letra en porcentaje es relativo al del estilo del que se hereda
          const t = e.t['fo:font-size'];
          if (t && /%\s*$/.test(t) && antes.t['fo:font-size']) { const b = puntos(antes.t['fo:font-size']); if (b != null) r.t['fo:font-size'] = (b * porcentaje(t, 100) / 100) + 'pt'; }
          if (e.lista) r.lista = e.lista;
          if (e.listaNombre) r.listaNombre = e.listaNombre;
        }
      }
      this.cache.set(clave, r);
      return r;
    }
    lista(nombre, ambito) { return this.listas[ambito].get(nombre) || this.listas.comun.get(nombre) || null; }
    fuente(t) {
      const n = t['style:font-name'];
      if (n) return this.fuentes.get(n) || n;
      const f = t['fo:font-family'];
      return f ? f.replace(/^['"]|['"]$/g, '') : null;
    }
  }

  // ───────── Geometrías ─────────
  // Tipos de forma de LibreOffice → forma de ECMA-376; "ooxml-X" ya trae el nombre de ECMA
  const FORMAS = sinProto({
    rectangle: 'rect', 'round-rectangle': 'roundRect', ellipse: 'ellipse', diamond: 'diamond', 'isosceles-triangle': 'triangle', 'right-triangle': 'rtTriangle', parallelogram: 'parallelogram', trapezoid: 'trapezoid',
    hexagon: 'hexagon', octagon: 'octagon', cross: 'plus', 'right-arrow': 'rightArrow', 'left-arrow': 'leftArrow', 'up-arrow': 'upArrow', 'down-arrow': 'downArrow', 'left-right-arrow': 'leftRightArrow', 'up-down-arrow': 'upDownArrow',
    'quad-arrow': 'quadArrow', 'bent-arrow': 'bentArrow', 'u-turn-arrow': 'uturnArrow', chevron: 'chevron', pentagon: 'homePlate', 'pentagon-right': 'homePlate', star4: 'star4', star5: 'star5', star6: 'star6', star8: 'star8', star12: 'star12', star24: 'star24',
    ring: 'donut', 'forbidden': 'noSmoking', 'smiley': 'smileyFace', heart: 'heart', 'lightning': 'lightningBolt', sun: 'sun', moon: 'moon', cloud: 'cloud', 'flower': 'sun', 'cube': 'cube', can: 'can', 'frame': 'frame', 'bevel': 'bevel',
    'flowchart-process': 'flowChartProcess', 'flowchart-alternate-process': 'flowChartAlternateProcess', 'flowchart-decision': 'flowChartDecision', 'flowchart-data': 'flowChartInputOutput', 'flowchart-predefined-process': 'flowChartPredefinedProcess',
    'flowchart-document': 'flowChartDocument', 'flowchart-terminator': 'flowChartTerminator', 'flowchart-preparation': 'flowChartPreparation', 'flowchart-connector': 'flowChartConnector', 'flowchart-card': 'flowChartPunchedCard',
    'flowchart-sequential-access': 'flowChartMagneticTape', 'flowchart-stored-data': 'flowChartOnlineStorage', 'flowchart-delay': 'flowChartDelay', 'flowchart-display': 'flowChartDisplay', 'flowchart-merge': 'flowChartMerge',
    'flowchart-extract': 'flowChartExtract', 'flowchart-summing-junction': 'flowChartSummingJunction', 'flowchart-or': 'flowChartOr', 'flowchart-collate': 'flowChartCollate', 'flowchart-sort': 'flowChartSort',
    'rectangular-callout': 'wedgeRectCallout', 'round-rectangular-callout': 'wedgeRoundRectCallout', 'round-callout': 'wedgeEllipseCallout', 'cloud-callout': 'cloudCallout',
    'horizontal-scroll': 'horizontalScroll', 'vertical-scroll': 'verticalScroll', 'doorplate': 'plaque', 'diagonal-stripe': 'diagStripe', 'left-bracket': 'leftBracket', 'right-bracket': 'rightBracket', 'left-brace': 'leftBrace', 'right-brace': 'rightBrace',
    'bracket-pair': 'bracketPair', 'brace-pair': 'bracePair', 'block-arc': 'blockArc', 'circular-arrow': 'circularArrow', 'notched-right-arrow': 'notchedRightArrow', 'striped-right-arrow': 'stripedRightArrow',
    'mso-spt1': 'rect', 'mso-spt2': 'roundRect', 'mso-spt3': 'ellipse', 'mso-spt4': 'diamond', 'mso-spt5': 'triangle', 'mso-spt6': 'rtTriangle', 'mso-spt7': 'parallelogram', 'mso-spt8': 'trapezoid', 'mso-spt9': 'hexagon', 'mso-spt10': 'octagon',
  });
  const ECMA_VALIDAS = new Set(['rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle', 'diamond', 'parallelogram', 'trapezoid', 'pentagon', 'hexagon', 'heptagon', 'octagon', 'decagon', 'dodecagon', 'star4', 'star5', 'star6', 'star7', 'star8', 'star10', 'star12', 'star16', 'star24', 'star32',
    'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'leftRightArrow', 'upDownArrow', 'quadArrow', 'bentArrow', 'uturnArrow', 'homePlate', 'chevron', 'cube', 'can', 'bevel', 'donut', 'noSmoking', 'blockArc', 'foldedCorner', 'smileyFace', 'heart', 'lightningBolt', 'sun', 'moon', 'cloud', 'arc',
    'plaque', 'frame', 'halfFrame', 'corner', 'diagStripe', 'chord', 'pie', 'teardrop', 'plus', 'wedgeRectCallout', 'wedgeRoundRectCallout', 'wedgeEllipseCallout', 'cloudCallout', 'ribbon', 'ribbon2', 'verticalScroll', 'horizontalScroll', 'wave', 'doubleWave',
    'flowChartProcess', 'flowChartAlternateProcess', 'flowChartDecision', 'flowChartInputOutput', 'flowChartPredefinedProcess', 'flowChartInternalStorage', 'flowChartDocument', 'flowChartMultidocument', 'flowChartTerminator', 'flowChartPreparation',
    'flowChartManualInput', 'flowChartManualOperation', 'flowChartConnector', 'flowChartPunchedCard', 'flowChartPunchedTape', 'flowChartSummingJunction', 'flowChartOr', 'flowChartCollate', 'flowChartSort', 'flowChartExtract', 'flowChartMerge', 'flowChartOnlineStorage',
    'flowChartDelay', 'flowChartMagneticTape', 'flowChartMagneticDisk', 'flowChartMagneticDrum', 'flowChartDisplay', 'leftBracket', 'rightBracket', 'leftBrace', 'rightBrace', 'bracketPair', 'bracePair', 'circularArrow', 'notchedRightArrow', 'stripedRightArrow',
    'mathPlus', 'mathMinus', 'mathMultiply', 'mathDivide', 'mathEqual', 'mathNotEqual', 'gear6', 'gear9', 'funnel', 'snip1Rect', 'snip2SameRect', 'snip2DiagRect', 'snipRoundRect', 'round1Rect', 'round2SameRect', 'round2DiagRect', 'leftUpArrow', 'bentUpArrow',
    'curvedRightArrow', 'curvedLeftArrow', 'curvedUpArrow', 'curvedDownArrow', 'ellipseRibbon', 'ellipseRibbon2', 'rightArrowCallout', 'leftArrowCallout', 'upArrowCallout', 'downArrowCallout', 'leftRightArrowCallout', 'upDownArrowCallout', 'quadArrowCallout']);

  // Trazado SVG ("M 0 0 L 10 10 C …") → órdenes del contrato (M, L, C, Q, Z) en coordenadas absolutas; los arcos pasan a curvas
  function trazadoSvg(d, tope) {
    const t = String(d || '').match(/[a-zA-Z]|-?\d*\.?\d+(?:e[-+]?\d+)?/g) || [];
    const cmds = [];
    let i = 0, cmd = '', x = 0, y = 0, ix = 0, iy = 0, cx = 0, cy = 0, ultimo = '';
    const n = () => { const v = parseFloat(t[i++]); return Number.isFinite(v) ? v : 0; };
    let seguridad = 0;
    while (i < t.length && seguridad++ < tope) {
      if (/^[a-zA-Z]$/.test(t[i])) cmd = t[i++];
      else if (cmd === 'M') cmd = 'L'; else if (cmd === 'm') cmd = 'l';
      const rel = cmd === cmd.toLowerCase();
      const C = cmd.toUpperCase();
      if (C === 'M') { x = n() + (rel ? x : 0); y = n() + (rel ? y : 0); ix = x; iy = y; cmds.push(['M', x, y]); }
      else if (C === 'L') { x = n() + (rel ? x : 0); y = n() + (rel ? y : 0); cmds.push(['L', x, y]); }
      else if (C === 'H') { x = n() + (rel ? x : 0); cmds.push(['L', x, y]); }
      else if (C === 'V') { y = n() + (rel ? y : 0); cmds.push(['L', x, y]); }
      else if (C === 'C') { const a = n() + (rel ? x : 0), b = n() + (rel ? y : 0), c = n() + (rel ? x : 0), e = n() + (rel ? y : 0); const f = n() + (rel ? x : 0), g = n() + (rel ? y : 0); cmds.push(['C', a, b, c, e, f, g]); cx = c; cy = e; x = f; y = g; }
      else if (C === 'S') { const c = n() + (rel ? x : 0), e = n() + (rel ? y : 0), f = n() + (rel ? x : 0), g = n() + (rel ? y : 0); const a = ultimo === 'C' ? 2 * x - cx : x, b = ultimo === 'C' ? 2 * y - cy : y; cmds.push(['C', a, b, c, e, f, g]); cx = c; cy = e; x = f; y = g; }
      else if (C === 'Q') { const a = n() + (rel ? x : 0), b = n() + (rel ? y : 0), f = n() + (rel ? x : 0), g = n() + (rel ? y : 0); cmds.push(['Q', a, b, f, g]); x = f; y = g; }
      else if (C === 'A') {
        const rx = Math.abs(n()), ry = Math.abs(n()), rot = n(), gr = n(), bar = n(), px = n() + (rel ? x : 0), py = n() + (rel ? y : 0);
        for (const c of arcoACurvas(x, y, rx, ry, rot, gr, bar, px, py)) cmds.push(c);
        x = px; y = py;
      }
      else if (C === 'Z') { cmds.push(['Z']); x = ix; y = iy; }
      else { i++; }
      ultimo = C;
    }
    return cmds;
  }
  // Arco elíptico de SVG → curvas de Bézier cúbicas (apéndice F.6 de SVG)
  function arcoACurvas(x1, y1, rx, ry, rotGrados, grande, barrido, x2, y2) {
    if (!rx || !ry || (x1 === x2 && y1 === y2)) return [['L', x2, y2]];
    const fi = rotGrados * Math.PI / 180, cf = Math.cos(fi), sf = Math.sin(fi);
    const dx = (x1 - x2) / 2, dy = (y1 - y2) / 2;
    const x1p = cf * dx + sf * dy, y1p = -sf * dx + cf * dy;
    const l = x1p * x1p / (rx * rx) + y1p * y1p / (ry * ry);
    if (l > 1) { rx *= Math.sqrt(l); ry *= Math.sqrt(l); }
    const s = (grande === barrido ? -1 : 1) * Math.sqrt(Math.max(0, (rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p) / (rx * rx * y1p * y1p + ry * ry * x1p * x1p)));
    const cxp = s * rx * y1p / ry, cyp = -s * ry * x1p / rx;
    const cx = cf * cxp - sf * cyp + (x1 + x2) / 2, cy = sf * cxp + cf * cyp + (y1 + y2) / 2;
    const ang = (ux, uy, vx, vy) => { const a = Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy); return a; };
    const t1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
    let dt = ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
    if (!barrido && dt > 0) dt -= 2 * Math.PI; else if (barrido && dt < 0) dt += 2 * Math.PI;
    const n = Math.max(1, Math.ceil(Math.abs(dt) / (Math.PI / 2)));
    const paso = dt / n, k = 4 / 3 * Math.tan(paso / 4);
    const out = [];
    let t = t1;
    const pt = (a) => [cf * rx * Math.cos(a) - sf * ry * Math.sin(a) + cx, sf * rx * Math.cos(a) + cf * ry * Math.sin(a) + cy];
    const dv = (a) => [-cf * rx * Math.sin(a) - sf * ry * Math.cos(a), -sf * rx * Math.sin(a) + cf * ry * Math.cos(a)];
    for (let i = 0; i < n; i++) {
      const p0 = pt(t), p1 = pt(t + paso), d0 = dv(t), d1 = dv(t + paso);
      out.push(['C', p0[0] + k * d0[0], p0[1] + k * d0[1], p1[0] - k * d1[0], p1[1] - k * d1[1], p1[0], p1[1]]);
      t += paso;
    }
    return out;
  }

  // draw:transform ("rotate (0.5) translate (1cm 2cm)" …) → lista de operaciones que se aplican en el orden escrito (la primera opera sobre la forma sin mover)
  // rotate: radianes, en sentido contrario al reloj; translate: medidas con unidad
  function transformacion(t) {
    const ops = [];
    if (!t) return ops;
    const re = /(rotate|translate|scale|skewX|skewY|matrix)\s*\(([^)]*)\)/g;
    let m;
    while ((m = re.exec(t)) && ops.length < 16) {
      const v = m[2].trim().split(/[\s,]+/);
      const n = (k, def) => { const x = parseFloat(v[k]); return Number.isFinite(x) ? x : def; };
      if (m[1] === 'rotate') ops.push({ r: n(0, 0) });
      else if (m[1] === 'translate') ops.push({ tx: emu(v[0]) || 0, ty: emu(v[1]) || 0 });
      else if (m[1] === 'scale') ops.push({ sx: n(0, 1), sy: n(1, n(0, 1)) });
      else if (m[1] === 'skewX') ops.push({ kx: n(0, 0) });
      else if (m[1] === 'skewY') ops.push({ ky: n(0, 0) });
      else if (m[1] === 'matrix') ops.push({ m: [n(0, 1), n(1, 0), n(2, 0), n(3, 1), emu(v[4]) || 0, emu(v[5]) || 0] });
    }
    return ops;
  }
  // Aplica las operaciones a un punto
  function aplicarOps(ops, x, y) {
    for (const o of ops) {
      if (o.r !== undefined) { const c = Math.cos(o.r), s = Math.sin(o.r); [x, y] = [x * c + y * s, -x * s + y * c]; }
      else if (o.tx !== undefined) { x += o.tx; y += o.ty; }
      else if (o.sx !== undefined) { x *= o.sx; y *= o.sy; }
      else if (o.kx !== undefined) { x += y * Math.tan(o.kx); }
      else if (o.ky !== undefined) { y += x * Math.tan(o.ky); }
      else if (o.m) { const [a, b, c, d, e, f] = o.m; [x, y] = [a * x + c * y + e, b * x + d * y + f]; }
    }
    return [x, y];
  }

  // ───────── Conversión ─────────
  const CLASES_MARCADOR = new Set(['title', 'outline', 'subtitle', 'notes', 'graphic', 'object', 'chart', 'table', 'orgchart', 'page', 'handout', 'text']);
  const CLASES_PIE = sinProto({ footer: 'pie', 'date-time': 'fecha', 'page-number': 'numero', header: 'encabezado' });
  const FORMATO_NUM = sinProto({ '1': 'arabic', a: 'alphaLc', A: 'alphaUc', i: 'romanLc', I: 'romanUc' });

  class Conversor {
    constructor(estilos, zip, contenidoRaiz) {
      this.e = estilos;
      this.zip = zip;
      this.mensajes = new Set();
      this.formas = 0;
      this.caracteres = 0;
      this.bytesImagenes = 0;
      this.imagenes = new Map();     // ruta → { entrada }
      this.pendientes = [];
      this.puntos = 0;
      this.pie = new Map();          // presentation:footer-decl / date-time-decl
      this.contenido = contenidoRaiz;
    }
    aviso(c) { this.mensajes.add(c); }
    contar() { if (++this.formas > MAX_FORMAS) throw new Error('La presentación tiene demasiadas formas y no se puede leer.'); }

    // ───── propiedades de texto y párrafo ─────
    runDe(t) {
      const r = {};
      const f = this.e.fuente(t); if (f) r.fuente = f;
      const tam = puntos(t['fo:font-size']); if (tam != null && tam > 0) r.tamano = Math.round(tam * 100) / 100;
      const pesoN = t['fo:font-weight']; if (pesoN === 'bold' || (parseInt(pesoN, 10) || 0) >= 600) r.negrita = true;
      if (t['fo:font-style'] === 'italic' || t['fo:font-style'] === 'oblique') r.cursiva = true;
      const u = t['style:text-underline-style']; if (u && u !== 'none') r.subrayado = true;
      const tt = t['style:text-line-through-style']; if (tt && tt !== 'none') r.tachado = true;
      const c = color(t['fo:color']); if (c) r.color = c;
      const pos = t['style:text-position'];
      if (pos) { const m = /^(super|sub|[-\d.]+%)/.exec(pos.trim()); if (m) { if (m[1] === 'super' || parseFloat(m[1]) > 0) r.sup = true; else if (m[1] === 'sub' || parseFloat(m[1]) < 0) r.sub = true; } }
      return r;
    }
    // número de la lista: nivel de un text:list-style
    nivelLista(lista, nivel) {
      if (!lista) return null;
      const n = String(nivel);
      for (const l of hijos(lista)) if (atr(l, 'text:level') === n) return l;
      return null;
    }
    vinetaDe(nivelEl, ctxT) {
      if (!nivelEl) return undefined;
      if (nivelEl.n === 'text:list-level-style-bullet') {
        const ch = atr(nivelEl, 'text:bullet-char');
        if (!ch) return null;
        const v = { car: Array.from(ch)[0], fuente: 'Arial' };
        const t = hijo(nivelEl, 'style:text-properties');
        if (t) {
          const f = this.e.fuente(t.a); if (f && !/opensymbol/i.test(f)) v.fuente = f;
          const c = color(t.a['fo:color']); if (c) v.color = c;
          const rel = t.a['fo:font-size']; if (rel && /%/.test(rel)) { const p = porcentaje(rel, 100); if (p && p !== 100) v.tamano = p; }
        }
        // los símbolos de LibreOffice (OpenSymbol) y de Office se pasan a su carácter Unicode
        const rel2 = atr(nivelEl, 'text:bullet-relative-size'); if (rel2) { const p = porcentaje(rel2, 100); if (p && p !== 100) v.tamano = p; }
        return v;
      }
      if (nivelEl.n === 'text:list-level-style-number') {
        const fmt = atr(nivelEl, 'style:num-format');
        if (fmt === undefined || fmt === '') return null;
        const base = FORMATO_NUM[fmt]; if (!base) return null;
        const pre = atr(nivelEl, 'style:num-prefix') || '', suf = atr(nivelEl, 'style:num-suffix') || '';
        let tipo;
        if (base === 'arabic') tipo = pre === '(' && suf === ')' ? 'arabicParenBoth' : suf === ')' ? 'arabicParenR' : suf === '' ? 'arabicPlain' : 'arabicPeriod';
        else if (base === 'alphaLc' || base === 'alphaUc' || base === 'romanLc' || base === 'romanUc') tipo = pre === '(' && suf === ')' ? base + 'ParenBoth' : suf === ')' ? base + 'ParenR' : base + 'Period';
        const ini = parseInt(atr(nivelEl, 'text:start-value'), 10);
        return { num: tipo, inicio: ini > 1 ? ini : undefined };
      }
      return undefined;
    }

    // ───── texto ─────
    // Párrafos de un contenedor (text-box, celda, forma): text:p, text:h y text:list
    parrafos(cont, ctx, ambito, base, nivelLista, lista, out) {
      out = out || [];
      for (const el of hijos(cont)) {
        if (out.length > 200000) break;
        if (el.n === 'text:p' || el.n === 'text:h') out.push(this.parrafo(el, ctx, ambito, base, nivelLista, lista));
        else if (el.n === 'text:list') {
          const nombreLista = atr(el, 'text:style-name');
          const ls = nombreLista ? this.e.lista(nombreLista, ambito) : lista;
          for (const item of hijos(el, 'text:list-item').concat(hijos(el, 'text:list-header'))) this.parrafos(item, ctx, ambito, base, nivelLista + 1, ls || lista, out);
        } else if (el.n === 'text:section') this.parrafos(el, ctx, ambito, base, nivelLista, lista, out);
      }
      return out;
    }
    parrafo(el, ctx, ambito, base, nivelLista, lista) {
      const est = this.e.resolver('paragraph', atr(el, 'text:style-name'), ambito);
      const t0 = Object.assign({}, base.t, est.t);
      const p0 = Object.assign({}, base.p, est.p);
      const p = { hijos: [] };
      const al = p0['fo:text-align'];
      p.alin = al === 'center' ? 'ctr' : al === 'end' || al === 'right' ? 'r' : al === 'justify' ? 'just' : 'l';
      const nivelVineta = nivelLista > 0 ? nivelLista : 0;
      if (nivelVineta > 1) p.nivel = Math.min(8, nivelVineta - 1);
      // sangrías: las del párrafo; en una lista, también las del nivel de la lista
      let izq = emu(p0['fo:margin-left']), ind = emu(p0['fo:text-indent']);
      const tamPunto = puntos(t0['fo:font-size']) || 18;
      let nivelEl = null;
      if (nivelVineta > 0 && lista) {
        nivelEl = this.nivelLista(lista, nivelVineta);
        if (nivelEl) {
          const lp = hijo(nivelEl, 'style:list-level-properties');
          const alin = lp ? hijo(lp, 'style:list-level-label-alignment') : null;
          if (alin) { if (atr(alin, 'fo:margin-left') != null) izq = emu(atr(alin, 'fo:margin-left')) + (izq || 0); if (atr(alin, 'fo:text-indent') != null) ind = emu(atr(alin, 'fo:text-indent')); }
          else if (lp) { const sb = emu(atr(lp, 'text:space-before')) || 0, ml = emu(atr(lp, 'text:min-label-width')) || 0; izq = (izq || 0) + sb + ml; ind = -ml; }
        }
      }
      if (izq != null) p.izq = izq;
      if (ind != null) p.primera = ind;
      const lh = p0['fo:line-height'];
      if (lh && /%/.test(lh)) p.interlineado = { pct: porcentaje(lh, 100) };
      else if (p0['style:line-height-at-least'] || (lh && lh !== 'normal')) { const pt = puntos(p0['style:line-height-at-least'] || lh); if (pt) p.interlineado = { pt }; }
      const antes = puntos(p0['fo:margin-top']), desp = puntos(p0['fo:margin-bottom']);
      if (antes) p.antes = antes; if (desp) p.despues = desp;
      if (nivelVineta > 0) {
        const v = this.vinetaDe(nivelEl, t0);
        if (v) p.vineta = v;
        else if (v === null) p.vineta = null;
      }
      // contenido
      const salida = p.hijos;
      this.contenidoParrafo(el, ctx, ambito, t0, salida, null);
      const primero = salida.find((x) => x.r && x.r.tamano);
      return p;
    }
    contenidoParrafo(el, ctx, ambito, t, salida, enlace) {
      const runBase = () => this.runDe(t);
      const poner = (hijo) => { if (enlace && hijo.texto) hijo.enlace = enlace; salida.push(hijo); };
      for (const x of el.h) {
        if (typeof x === 'string') {
          const s = x.replace(/[\r\n\t]+/g, ' ');
          if ((this.caracteres += s.length) > MAX_CARACTERES) throw new Error('La presentación tiene demasiado texto y no se puede leer.');
          if (s === '') continue;
          const r = runBase();
          poner(Object.keys(r).length ? { texto: s, r } : { texto: s });
          continue;
        }
        switch (x.n) {
          case 'text:span': {
            const est = this.e.resolver('text', atr(x, 'text:style-name'), ambito);
            const t2 = Object.assign({}, t, est.t);
            if (t.fontSizeRel) delete t2.fontSizeRel;
            // porcentaje de tamaño relativo al del párrafo
            const fs = est.t['fo:font-size'];
            if (fs && /%/.test(fs) && t['fo:font-size']) { const b = puntos(t['fo:font-size']); if (b) t2['fo:font-size'] = (b * porcentaje(fs, 100) / 100) + 'pt'; }
            this.contenidoParrafo(x, ctx, ambito, t2, salida, enlace);
            break;
          }
          case 'text:a': {
            const href = atr(x, 'xlink:href');
            let destino = null;
            if (href) {
              if (/^(https?:|mailto:)/i.test(href)) destino = href;
              else if (href.startsWith('#')) { const m = /(\d+)$/.exec(href); destino = m ? { diapositiva: +m[1] } : null; }
            }
            const est = this.e.resolver('text', atr(x, 'text:style-name'), ambito);
            this.contenidoParrafo(x, ctx, ambito, Object.assign({}, t, est.t), salida, destino || enlace);
            break;
          }
          case 'text:line-break': salida.push({ salto: true }); break;
          case 'text:tab': poner({ texto: '\t', r: runBase() }); break;
          case 'text:s': { const n = acotar(parseInt(atr(x, 'text:c'), 10) || 1, 1, 10000); if ((this.caracteres += n) > MAX_CARACTERES) throw new Error('La presentación tiene demasiado texto y no se puede leer.'); poner({ texto: ' '.repeat(n), r: runBase() }); break; }
          case 'text:page-number': case 'presentation:page-number': salida.push({ campo: 'numero', texto: '‹#›', r: runBase() }); break;
          case 'text:date': case 'text:time': case 'presentation:date-time': salida.push({ campo: 'fecha', texto: textoDe(x) || '', r: runBase() }); break;
          case 'presentation:footer': { const f = ctx.pie && ctx.pie.pie; if (f) poner({ texto: f, r: runBase() }); break; }
          case 'presentation:header': { const f = ctx.pie && ctx.pie.encabezado; if (f) poner({ texto: f, r: runBase() }); break; }
          case 'text:bookmark-start': case 'text:bookmark-end': case 'text:bookmark': case 'office:annotation': case 'text:soft-page-break': break;
          default: this.contenidoParrafo(x, ctx, ambito, t, salida, enlace);
        }
      }
    }

    // ───── formas ─────
    relleno(g, ambito, ctx) {
      const f = g['draw:fill'];
      if (f === undefined) return undefined;
      if (f === 'none') return null;
      const op = g['draw:opacity'] !== undefined ? acotar(porcentaje(g['draw:opacity'], 100) / 100, 0, 1) : 1;
      if (f === 'solid') { const c = color(g['draw:fill-color']); return c ? (op < 1 ? { color: c, alfa: op } : { color: c }) : undefined; }
      if (f === 'gradient') { const gr = this.degradado(g['draw:fill-gradient-name']); if (gr) return { degradado: gr }; const c = color(g['draw:fill-color']); return c ? { color: c } : undefined; }
      if (f === 'bitmap') {
        const im = this.e.rellenosImagen.get(g['draw:fill-image-name']);
        const h = im && atr(im, 'xlink:href');
        const img = h ? this.imagen(h) : null;
        if (img) return { imagen: img };
        return undefined;
      }
      if (f === 'hatch') { const c = color(g['draw:fill-color']); return c ? { color: c } : undefined; }
      return undefined;
    }
    degradado(nombre) {
      const d = nombre && this.e.gradientes.get(nombre);
      if (!d) return null;
      const estilo = atr(d, 'draw:style') || 'linear';
      const ini = color(atr(d, 'draw:start-color')) || 'FFFFFF', fin = color(atr(d, 'draw:end-color')) || 'FFFFFF';
      const paradas = [];
      const stops = hijos(d, 'loext:gradient-stop');
      if (stops.length >= 2) for (const s of stops) { const c = color(atr(s, 'loext:color-value')) || color(atr(s, 'svg:stop-color')); if (c) paradas.push({ pos: acotar(parseFloat(atr(s, 'svg:offset')) * 100, 0, 100), color: c }); }
      const borde = acotar(porcentaje(atr(d, 'draw:border'), 0), 0, 100);
      const anguloStr = atr(d, 'draw:angle') || '0';
      let grados = parseFloat(anguloStr) || 0;
      if (!/deg/.test(anguloStr)) grados = grados / 10;
      if (estilo === 'linear' || estilo === 'axial') {
        const lista = paradas.length >= 2 ? paradas : (estilo === 'axial' ? [{ pos: 0, color: fin }, { pos: 50, color: ini }, { pos: 100, color: fin }] : [{ pos: borde, color: ini }, { pos: 100, color: fin }]);
        if (estilo === 'linear' && paradas.length < 2 && borde > 0) lista.unshift({ pos: 0, color: ini });
        return { angulo: ((90 - grados) % 360 + 360) % 360, paradas: lista };
      }
      // radial, elipsoide, cuadrado, rectangular: del centro hacia fuera (LibreOffice pone el color final en el centro)
      return { radial: true, paradas: [{ pos: 0, color: fin }, { pos: 100, color: ini }] };
    }
    linea(g, ambito) {
      const s = g['draw:stroke'];
      if (s === undefined) return undefined;
      if (s === 'none') return null;
      const l = { color: color(g['svg:stroke-color']) || '000000', ancho: emu(g['svg:stroke-width']) || 0 };
      if (s === 'dash') {
        const d = this.e.guiones.get(g['draw:stroke-dash']);
        const estilo = d && atr(d, 'draw:style');
        const puntos1 = d ? parseFloat(atr(d, 'draw:dots1')) || 1 : 1, long1 = d ? atr(d, 'draw:dots1-length') : null;
        l.guion = d && puntos1 > 0 && (!long1 || /^\s*[\d.]+\s*%/.test(long1)) && parseFloat(long1) < 150 ? (atr(d, 'draw:dots2') ? 'dashDot' : 'sysDot') : 'dash';
      }
      const op = g['svg:stroke-opacity']; if (op !== undefined) { const a = porcentaje(op, 100) / 100; if (a < 1) l.alfa = acotar(a, 0, 1); }
      const j = g['draw:stroke-linejoin']; if (j === 'round') l.union = 'round'; else if (j === 'bevel') l.union = 'bevel';
      const ms = g['draw:marker-start'], me = g['draw:marker-end'];
      const flecha = (m) => (m ? (/oval|circle/i.test(m) ? 'oval' : /diamond/i.test(m) ? 'diamond' : /square/i.test(m) ? 'diamond' : /arrow|triangle|line/i.test(m) ? 'triangle' : 'triangle') : null);
      if (ms) l.inicio = flecha(ms); if (me) l.fin = flecha(me);
      return l;
    }
    sombra(g) {
      if (g['draw:shadow'] !== 'visible') return undefined;
      const s = { color: color(g['draw:shadow-color']) || '000000', x: emu(g['draw:shadow-offset-x']) || 0, y: emu(g['draw:shadow-offset-y']) || 0 };
      const op = g['draw:shadow-opacity']; if (op !== undefined) s.alfa = acotar(porcentaje(op, 100) / 100, 0, 1);
      const bl = emu(g['draw:shadow-blur']); if (bl) s.desenfoque = bl;
      return s;
    }
    imagen(href) {
      if (!href || /^(https?:|file:|\/\/|data:)/i.test(href)) { this.aviso('imagen-externa'); return null; }
      const ruta = href.replace(/^\.\//, '');
      const entrada = this.zip.file(ruta);
      if (!entrada) { this.aviso('imagen-falta'); return null; }
      if (this.imagenes.has(ruta)) return this.imagenes.get(ruta);
      if (this.imagenes.size >= MAX_IMAGENES) throw new Error('La presentación tiene demasiadas imágenes y no se puede leer.');
      const im = { datos: null, _entrada: entrada, _ruta: ruta };
      this.imagenes.set(ruta, im);
      this.pendientes.push(im);
      return im;
    }
    // x, y, ancho, alto de un elemento de dibujo (con draw:transform si lo trae: se transforman las esquinas y se saca el centro, el giro y el tamaño)
    caja(el) {
      const x = emu(atr(el, 'svg:x')) || 0, y = emu(atr(el, 'svg:y')) || 0, w = emu(atr(el, 'svg:width')) || 0, h = emu(atr(el, 'svg:height')) || 0;
      const ops = transformacion(atr(el, 'draw:transform'));
      if (!ops.length) return { x, y, w, h };
      const p0 = aplicarOps(ops, x, y), p1 = aplicarOps(ops, x + w, y), p3 = aplicarOps(ops, x, y + h), p2 = aplicarOps(ops, x + w, y + h);
      const cx = (p0[0] + p2[0]) / 2, cy = (p0[1] + p2[1]) / 2;
      const ancho = Math.hypot(p1[0] - p0[0], p1[1] - p0[1]), alto = Math.hypot(p3[0] - p0[0], p3[1] - p0[1]);
      const c = { x: Math.round(cx - ancho / 2), y: Math.round(cy - alto / 2), w: Math.round(ancho), h: Math.round(alto) };
      const giro = Math.atan2(p1[1] - p0[1], p1[0] - p0[0]) * 180 / Math.PI;
      if (Math.abs(giro) > 0.01) c.rot = (giro % 360 + 360) % 360;
      return c;
    }
    // Estilos de una forma
    estiloForma(el, ambito) {
      const g = this.e.resolver('graphic', atr(el, 'draw:style-name'), ambito);
      const pres = atr(el, 'presentation:style-name') ? this.e.resolver('presentation', atr(el, 'presentation:style-name'), ambito) : null;
      // predeterminados < estilo de presentación < estilo propio de la forma (sin lo que este hereda de «standard»)
      const propio = atr(el, 'draw:style-name') ? this.e.resolver('graphic', atr(el, 'draw:style-name'), ambito, true) : null;
      const std = this.e.resolver('graphic', 'standard', ambito);
      const gg = pres ? Object.assign({}, std.g, pres.g, propio ? propio.g : {}) : g.g;
      return { g: gg, base: this.baseTexto(el, ambito, pres, g) };
    }
    // Propiedades de texto por defecto de una forma: estilo de presentación → draw:text-style-name → estilo gráfico
    baseTexto(el, ambito, pres, graf) {
      const ts = atr(el, 'draw:text-style-name') ? this.e.resolver('paragraph', atr(el, 'draw:text-style-name'), ambito) : null;
      const t = Object.assign({}, graf ? graf.t : {}, pres ? pres.t : {}, ts ? ts.t : {});
      const p = Object.assign({}, graf ? graf.p : {}, pres ? pres.p : {}, ts ? ts.p : {});
      return { t, p, lista: pres && pres.lista ? pres.lista : null };
    }
    cuerpoTexto(cont, el, ambito, g, base, ctx) {
      const ps = this.parrafos(cont, ctx, ambito, base, 0, base.lista, []);
      const hayTexto = ps.some((p) => p.hijos.some((h) => (h.texto && h.texto.trim() !== '') || h.campo));
      if (!hayTexto) return null;
      const t = { parrafos: ps };
      const ancla = g['draw:textarea-vertical-align'];
      t.ancla = ancla === 'middle' ? 'ctr' : ancla === 'bottom' ? 'b' : 't';
      const m = { i: emu(g['fo:padding-left']), s: emu(g['fo:padding-top']), d: emu(g['fo:padding-right']), a: emu(g['fo:padding-bottom']) };
      t.margenes = { i: m.i == null ? 91440 : m.i, s: m.s == null ? 45720 : m.s, d: m.d == null ? 91440 : m.d, a: m.a == null ? 45720 : m.a };
      if (g['fo:wrap-option'] === 'no-wrap') t.ajuste = false;
      if (g['draw:auto-grow-height'] === 'true' && g['draw:auto-grow-width'] !== 'true') t.autoajuste = 'forma';
      else if (g['style:shrink-to-fit'] === 'true') t.autoajuste = 'texto';
      if (g['style:writing-mode'] === 'tb-rl') t.vertical = true;
      return t;
    }
    convertir(el, ambito, ctx, prof) {
      if (prof > MAX_PROFUNDIDAD) { this.aviso('grupos-anidados'); return null; }
      this.contar();
      switch (el.n) {
        case 'draw:frame': return this.marco(el, ambito, ctx, prof);
        case 'draw:custom-shape': return this.formaPersonalizada(el, ambito, ctx);
        case 'draw:rect': return this.formaSimple(el, ambito, ctx, 'rect');
        case 'draw:ellipse': case 'draw:circle': return this.formaSimple(el, ambito, ctx, 'ellipse');
        case 'draw:line': return this.lineaSimple(el, ambito, ctx);
        case 'draw:connector': return this.lineaSimple(el, ambito, ctx);
        case 'draw:polyline': case 'draw:polygon': case 'draw:path': return this.trazado(el, ambito, ctx);
        case 'draw:g': return this.grupo(el, ambito, ctx, prof);
        case 'draw:page-thumbnail': case 'draw:control': case 'draw:plugin': case 'draw:applet': return null;
        default: return null;
      }
    }
    decorar(f, el, ambito, est, ctx, nombrar) {
      const nombre = atr(el, 'draw:name'); if (nombre) f.nombre = nombre.slice(0, 80);
      const desc = hijo(el, 'svg:desc'); if (desc) f.descripcion = textoDe(desc).slice(0, 500);
      return f;
    }
    formaSimple(el, ambito, ctx, geom) {
      const est = this.estiloForma(el, ambito);
      const c = this.caja(el);
      const f = Object.assign({ tipo: 'forma', geom }, c);
      if (geom === 'rect') { const r = emu(atr(el, 'draw:corner-radius')) || emu(est.g['draw:corner-radius']); if (r) { f.geom = 'roundRect'; f.ajustes = [Math.round(acotar(r / Math.min(c.w, c.h) * 100000, 0, 50000))]; } }
      this.aplicar(f, est, el, ambito, ctx, el);
      return this.decorar(f, el, ambito, est, ctx);
    }
    aplicar(f, est, el, ambito, ctx, contenedorTexto) {
      const r = this.relleno(est.g, ambito, ctx); if (r !== undefined) f.relleno = r;
      const l = this.linea(est.g, ambito); if (l !== undefined) f.linea = l;
      const s = this.sombra(est.g); if (s) f.sombra = s;
      const t = this.cuerpoTexto(contenedorTexto, el, ambito, est.g, est.base, ctx);
      if (t) f.texto = t;
      const hl = this.enlaceDe(el);
      if (hl) f.enlace = hl;
    }
    enlaceDe(el) {
      const ev = hijo(hijo(el, 'office:event-listeners'), 'presentation:event-listener');
      const h = ev && atr(ev, 'xlink:href');
      if (h && /^(https?:|mailto:)/i.test(h)) return h;
      return undefined;
    }
    lineaSimple(el, ambito, ctx) {
      const est = this.estiloForma(el, ambito);
      const x1 = emu(atr(el, 'svg:x1')) || 0, y1 = emu(atr(el, 'svg:y1')) || 0, x2 = emu(atr(el, 'svg:x2')) || 0, y2 = emu(atr(el, 'svg:y2')) || 0;
      const f = { tipo: 'linea', x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1) };
      if (x2 < x1) f.volteoH = true; if (y2 < y1) f.volteoV = true;
      if ((x2 < x1) !== (y2 < y1) && false) f.volteoV = true;
      // la línea va de (x1,y1) a (x2,y2): la caja con volteo da la diagonal que corresponde
      f.volteoH = x2 < x1; f.volteoV = y2 < y1;
      const l = this.linea(est.g, ambito);
      f.linea = l === undefined ? { color: '000000', ancho: 0 } : l || { color: '000000', ancho: 0 };
      const s = this.sombra(est.g); if (s) f.sombra = s;
      return this.decorar(f, el, ambito, est, ctx);
    }
    trazado(el, ambito, ctx) {
      const est = this.estiloForma(el, ambito);
      const caja = this.caja(el);
      const vb = (atr(el, 'svg:viewBox') || '').trim().split(/[\s,]+/).map(parseFloat);
      const w = vb.length === 4 && vb[2] > 0 ? vb[2] : caja.w || 1, h = vb.length === 4 && vb[3] > 0 ? vb[3] : caja.h || 1;
      let cmds;
      if (el.n === 'draw:path') cmds = trazadoSvg(atr(el, 'svg:d'), 200000);
      else {
        const pts = (atr(el, 'draw:points') || '').trim().split(/\s+/).map((p) => p.split(',').map(parseFloat)).filter((p) => p.length === 2 && Number.isFinite(p[0]) && Number.isFinite(p[1]));
        cmds = pts.slice(0, 200000).map((p, i) => [i ? 'L' : 'M', p[0], p[1]]);
        if (el.n === 'draw:polygon' && cmds.length) cmds.push(['Z']);
      }
      if ((this.puntos += cmds.length) > MAX_PUNTOS) throw new Error('La presentación tiene demasiados trazos y no se puede leer.');
      if (!cmds.length) return null;
      const cerrado = cmds.some((c) => c[0] === 'Z');
      const f = Object.assign({ tipo: 'forma', geom: { rutas: [{ w, h, cmds, sinRelleno: !cerrado && el.n !== 'draw:path' }] } }, caja);
      this.aplicar(f, est, el, ambito, ctx, el);
      if (!cerrado && el.n === 'draw:polyline') f.relleno = null;
      return this.decorar(f, el, ambito, est, ctx);
    }
    formaPersonalizada(el, ambito, ctx) {
      const est = this.estiloForma(el, ambito);
      const caja = this.caja(el);
      const geo = hijo(el, 'draw:enhanced-geometry');
      const tipo = geo ? atr(geo, 'draw:type') || '' : '';
      let geom = 'rect';
      if (/^ooxml-/.test(tipo)) geom = ECMA_VALIDAS.has(tipo.slice(6)) ? tipo.slice(6) : 'rect';
      else if (FORMAS[tipo]) geom = FORMAS[tipo];
      else if (tipo && tipo !== 'non-primitive') this.aviso('formas');
      else if (tipo === 'non-primitive') this.aviso('formas');
      const f = Object.assign({ tipo: 'forma', geom }, caja);
      if (geo) { if (atr(geo, 'draw:mirror-horizontal') === 'true') f.volteoH = !f.volteoH; if (atr(geo, 'draw:mirror-vertical') === 'true') f.volteoV = !f.volteoV; }
      if (geom === 'roundRect') { const m = parseFloat((atr(geo, 'draw:modifiers') || '').split(/\s+/)[0]); if (Number.isFinite(m)) f.ajustes = [Math.round(acotar(m / 21600 * 100000, 0, 50000))]; }
      this.aplicar(f, est, el, ambito, ctx, el);
      return this.decorar(f, el, ambito, est, ctx);
    }
    grupo(el, ambito, ctx, prof) {
      const hijosF = [];
      for (const h of hijos(el)) { const f = this.convertir(h, ambito, ctx, prof + 1); if (f) hijosF.push(f); }
      if (!hijosF.length) return null;
      // los hijos de un grupo de OpenDocument tienen coordenadas absolutas: la caja del grupo es la que los envuelve
      const cajaDe = (f) => { const rad = (f.rot || 0) * Math.PI / 180; const cw = Math.abs(f.w * Math.cos(rad)) + Math.abs(f.h * Math.sin(rad)), ch = Math.abs(f.w * Math.sin(rad)) + Math.abs(f.h * Math.cos(rad)); const cx = f.x + f.w / 2, cy = f.y + f.h / 2; return [cx - cw / 2, cy - ch / 2, cx + cw / 2, cy + ch / 2]; };
      let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
      for (const f of hijosF) { const b = cajaDe(f); x1 = Math.min(x1, b[0]); y1 = Math.min(y1, b[1]); x2 = Math.max(x2, b[2]); y2 = Math.max(y2, b[3]); }
      const f = { tipo: 'grupo', x: Math.round(x1), y: Math.round(y1), w: Math.round(x2 - x1), h: Math.round(y2 - y1), hijos: hijosF };
      return this.decorar(f, el, ambito, null, ctx);
    }
    marco(el, ambito, ctx, prof) {
      const clase = atr(el, 'presentation:class');
      if (clase === 'notes') return null;                 // las notas van aparte
      if (CLASES_PIE[clase] && ctx.master) return this.marcoPie(el, ambito, ctx, clase);
      const est = this.estiloForma(el, ambito);
      let caja = this.caja(el);
      const imagen = hijo(el, 'draw:image');
      const tabla = hijo(el, 'table:table');
      const tb = hijo(el, 'draw:text-box');
      let f;
      if (tabla) f = this.tabla(tabla, caja, ambito, ctx, est);
      else if (imagen && !hijo(el, 'draw:object') && !hijo(el, 'draw:object-ole')) {
        const im = this.imagen(atr(imagen, 'xlink:href'));
        if (!im) return null;
        f = Object.assign({ tipo: 'imagen', imagen: im }, caja);
        const l = this.linea(est.g, ambito); if (l) f.linea = l;
        const rc = this.recorte(est.g['fo:clip'], im); if (rc) f.recorte = rc;
        const hl = this.enlaceDe(el); if (hl) f.enlace = hl;
      } else if (hijo(el, 'draw:object') || hijo(el, 'draw:object-ole') || hijo(el, 'draw:plugin')) {
        if (imagen) { const im = this.imagen(atr(imagen, 'xlink:href')); if (im) { f = Object.assign({ tipo: 'imagen', imagen: im }, caja); this.aviso('objetos'); } else return null; }
        else { this.aviso('objetos'); return null; }
      } else if (tb) {
        f = Object.assign({ tipo: 'forma', geom: 'rect' }, caja);
        const r = this.relleno(est.g, ambito, ctx); f.relleno = r === undefined ? null : r;
        const l = this.linea(est.g, ambito); f.linea = l === undefined ? null : l;
        const s = this.sombra(est.g); if (s) f.sombra = s;
        const t = this.cuerpoTexto(tb, el, ambito, est.g, est.base, ctx);
        if (t) f.texto = t;
        else if (!f.relleno && !f.linea) return null;
        const hl = this.enlaceDe(el); if (hl) f.enlace = hl;
      } else return null;
      return this.decorar(f, el, ambito, est, ctx);
    }
    // fo:clip="rect(arriba, derecha, abajo, izquierda)": distancias desde el borde de la imagen original
    recorte(clip, im) { return null; }
    marcoPie(el, ambito, ctx, clase) {
      // se escribe una copia del marcador del patrón con el texto de la diapositiva
      const tipo = CLASES_PIE[clase];
      const visible = ctx.visibles[tipo];
      if (!visible) return null;
      const est = this.estiloForma(el, ambito);
      const caja = this.caja(el);
      const tb = hijo(el, 'draw:text-box');
      if (!tb) return null;
      const f = Object.assign({ tipo: 'forma', geom: 'rect' }, caja);
      f.relleno = null; f.linea = null;
      const t = this.cuerpoTexto(tb, el, ambito, est.g, est.base, ctx);
      if (!t) return null;
      f.texto = t;
      return f;
    }

    // ───── tablas ─────
    tabla(el, caja, ambito, ctx, estMarco) {
      const columnas = [];
      const filas = [];
      const llenarColumnas = (c) => {
        for (const x of hijos(c)) {
          if (x.n === 'table:table-column') {
            const est = this.e.resolver('table-column', atr(x, 'table:style-name'), ambito);
            const w = emu(est.col['style:column-width']) || 914400;
            const n = acotar(parseInt(atr(x, 'table:number-columns-repeated'), 10) || 1, 1, 63);
            for (let i = 0; i < n && columnas.length < 63; i++) columnas.push(w);
          } else if (x.n === 'table:table-column-group' || x.n === 'table:table-columns' || x.n === 'table:table-header-columns') llenarColumnas(x);
        }
      };
      llenarColumnas(el);
      if (!columnas.length) return null;
      const spans = [];                     // [fila][col] = restantes de un rowSpan
      let celdasTotal = 0;
      const procesarFila = (fe) => {
        if (filas.length >= 5000) return;
        const estF = this.e.resolver('table-row', atr(fe, 'table:style-name'), ambito);
        const alto = emu(estF.fila['style:row-height'] || estF.fila['style:min-row-height']) || 365760;
        const celdas = [];
        const rep = acotar(parseInt(atr(fe, 'table:number-rows-repeated'), 10) || 1, 1, 1000);
        for (const ce of hijos(fe)) {
          if (ce.n !== 'table:table-cell' && ce.n !== 'table:covered-table-cell') continue;
          const nRep = acotar(parseInt(atr(ce, 'table:number-columns-repeated'), 10) || 1, 1, 63);
          for (let k = 0; k < nRep && celdas.length < columnas.length; k++) {
            if (++celdasTotal > MAX_CELDAS) throw new Error('La tabla tiene demasiadas celdas y no se puede leer.');
            const cubierta = ce.n === 'table:covered-table-cell';
            const c = {};
            if (cubierta) c.fusionada = true;
            else {
              const est = this.e.resolver('table-cell', atr(ce, 'table:style-name'), ambito);
              const cs = parseInt(atr(ce, 'table:number-columns-spanned'), 10) || 1, rs = parseInt(atr(ce, 'table:number-rows-spanned'), 10) || 1;
              if (cs > 1) c.columnas = Math.min(cs, 63); if (rs > 1) c.filas = Math.min(rs, 5000);
              const bg = est.c['fo:background-color'] || est.g['fo:background-color'];
              if (color(bg)) c.relleno = { color: color(bg) };
              const bordes = {};
              const borde = (v) => { if (v === undefined) return undefined; if (v === 'none' || v === 'hidden') return null; const m = /^\s*([\d.]+\w*)\s+(\w+)\s+(#[0-9a-fA-F]{6})/.exec(v); return m ? { color: color(m[3]), ancho: emu(m[1]) || 6350 } : undefined; };
              const todos = borde(est.c['fo:border']);
              for (const [k2, a2] of [['izquierda', 'fo:border-left'], ['derecha', 'fo:border-right'], ['arriba', 'fo:border-top'], ['abajo', 'fo:border-bottom']]) { const b = borde(est.c[a2]); const v = b !== undefined ? b : todos; if (v !== undefined) bordes[k2] = v; }
              if (Object.keys(bordes).length) c.bordes = bordes;
              const va = est.c['style:vertical-align']; if (va === 'middle') c.ancla = 'ctr'; else if (va === 'bottom') c.ancla = 'b';
              const pad = emu(est.c['fo:padding']); const lado = (a2) => emu(est.c[a2]);
              c.margenes = { i: lado('fo:padding-left') != null ? lado('fo:padding-left') : pad, d: lado('fo:padding-right') != null ? lado('fo:padding-right') : pad, s: lado('fo:padding-top') != null ? lado('fo:padding-top') : pad, a: lado('fo:padding-bottom') != null ? lado('fo:padding-bottom') : pad };
              for (const kk of Object.keys(c.margenes)) if (c.margenes[kk] == null) delete c.margenes[kk];
              const base = { t: {}, p: {}, lista: null };
              const ps = this.parrafos(ce, ctx, ambito, base, 0, null, []);
              if (ps.length) c.texto = { parrafos: ps };
            }
            celdas.push(c);
          }
        }
        while (celdas.length < columnas.length) celdas.push({ fusionada: true });
        for (let r = 0; r < rep && filas.length < 5000; r++) filas.push({ alto, celdas: celdas.map((c) => Object.assign({}, c)) });
      };
      const recorrerFilas = (c) => {
        for (const x of hijos(c)) {
          if (x.n === 'table:table-row') procesarFila(x);
          else if (x.n === 'table:table-header-rows' || x.n === 'table:table-rows' || x.n === 'table:table-row-group') recorrerFilas(x);
        }
      };
      recorrerFilas(el);
      if (!filas.length) return null;
      // las celdas cubiertas por una combinación vertical: marca 'v'
      for (let y = 0; y < filas.length; y++) for (let x = 0; x < columnas.length; x++) {
        const c = filas[y].celdas[x];
        if (c.filas > 1 || c.columnas > 1) {
          for (let dy = 0; dy < (c.filas || 1); dy++) for (let dx = 0; dx < (c.columnas || 1); dx++) {
            if (!dy && !dx) continue;
            const d = filas[y + dy] && filas[y + dy].celdas[x + dx];
            if (d) d.fusionada = dy > 0 ? (dx > 0 ? true : 'v') : 'h';
          }
        }
      }
      const ancho = columnas.reduce((a, b) => a + b, 0);
      return { tipo: 'tabla', x: caja.x, y: caja.y, w: ancho, h: filas.reduce((a, f) => a + f.alto, 0), columnas, filas };
    }
  }

  // ───────── Documento ─────────
  const tamanoDe = (entrada) => (entrada && entrada._origen && entrada._origen.real) || 0;

  async function abrir(bytes) {
    const datos = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    let zip;
    try { zip = await ZipPropio.loadAsync(datos); } catch (e) { throw new Error('El archivo no es un .odp válido (no se puede abrir como un paquete).'); }
    if (zip.file('encrypted-package') || (zip.file('META-INF/manifest.xml') && /encryption-data/.test(await zip.file('META-INF/manifest.xml').async('string')))) {
      const err = new Error('La presentación está protegida con contraseña. Sacale la contraseña en LibreOffice y volvé a intentarlo.');
      err.codigo = 'cifrado';
      throw err;
    }
    if (!zip.file('content.xml')) throw new Error('El archivo no es una presentación de OpenDocument (.odp).');
    return zip;
  }
  async function parece(bytes) {
    try {
      const datos = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
      if (datos.length < 4 || datos[0] !== 0x50 || datos[1] !== 0x4B) return false;
      const zip = await ZipPropio.loadAsync(datos);
      if (!zip.file('content.xml')) return false;
      const mt = zip.file('mimetype') ? (await zip.file('mimetype').async('string')).trim() : '';
      return mt === 'application/vnd.oasis.opendocument.presentation' || mt === 'application/vnd.oasis.opendocument.presentation-template' || (!mt && /<office:presentation/.test((await zip.file('content.xml').async('string')).slice(0, 60000)));
    } catch (e) { return false; }
  }

  async function leer(bytes) {
    const zip = await abrir(bytes);
    const texto = async (ruta) => {
      const f = zip.file(ruta);
      if (!f) return null;
      if (tamanoDe(f) > MAX_TEXTO_XML) throw new Error('La presentación es demasiado grande para abrirla aquí.');
      const t = await f.async('string');
      if (t.length > MAX_TEXTO_XML) throw new Error('La presentación es demasiado grande para abrirla aquí.');
      return t;
    };
    const contenido = arbol(await texto('content.xml'));
    const estilosXml = await texto('styles.xml');
    const estilosRaiz = estilosXml ? arbol(estilosXml) : null;
    if (!contenido || contenido.n !== 'office:document-content') throw new Error('El archivo no es una presentación de OpenDocument (.odp).');
    const est = new Estilos();
    if (estilosRaiz) {
      est.leerFuentes(estilosRaiz);
      est.leerDefinicionesDe(hijo(estilosRaiz, 'office:styles'));
      est.leerEstilosDe(hijo(estilosRaiz, 'office:styles'), 'c', false);
      est.leerEstilosDe(hijo(estilosRaiz, 'office:automatic-styles'), 'm', true);
      est.leerMaestras(hijo(estilosRaiz, 'office:master-styles'));
    }
    est.leerFuentes(contenido);
    est.leerDefinicionesDe(hijo(contenido, 'office:styles'));
    est.leerEstilosDe(hijo(contenido, 'office:styles'), 'c', false);
    est.leerEstilosDe(hijo(contenido, 'office:automatic-styles'), 'c', true);
    const conv = new Conversor(est, zip, contenido);
    const presentacion = hijo(hijo(contenido, 'office:body'), 'office:presentation');
    if (!presentacion) throw new Error('El archivo no es una presentación de OpenDocument (.odp).');
    // declaraciones de pie y fecha
    const decl = { pie: new Map(), fecha: new Map() };
    for (const d of hijos(presentacion, 'presentation:footer-decl')) decl.pie.set(atr(d, 'presentation:name'), textoDe(d));
    for (const d of hijos(presentacion, 'presentation:date-time-decl')) decl.fecha.set(atr(d, 'presentation:name'), textoDe(d));
    const paginas = hijos(presentacion, 'draw:page').slice(0, MAX_DIAPOSITIVAS);
    if (!paginas.length) throw new Error('La presentación no tiene diapositivas.');

    // tamaño: el de la página maestra de la primera diapositiva
    const nombreMaestra0 = atr(paginas[0], 'draw:master-page-name');
    const m0 = est.maestras.get(nombreMaestra0) || [...est.maestras.values()][0];
    const pag0 = m0 ? est.paginas.get(atr(m0, 'style:page-layout-name')) : null;
    const ancho = (pag0 && pag0.ancho) || 9144000, alto = (pag0 && pag0.alto) || 6858000;
    if (ancho > 51206400 || alto > 51206400 || ancho < 1 || alto < 1) throw new Error('El archivo está dañado (tamaño de la diapositiva).');

    // páginas maestras usadas
    const indiceMaestra = new Map();
    const maestros = [];
    const maestraDe = (nombre) => {
      const clave = est.maestras.has(nombre) ? nombre : [...est.maestras.keys()][0];
      if (indiceMaestra.has(clave)) return indiceMaestra.get(clave);
      const mp = est.maestras.get(clave);
      const m = { fondo: null, formas: [] };
      if (mp) {
        const dp = est.resolver('drawing-page', atr(mp, 'draw:style-name'), 'm');
        const r = conv.relleno(dp.dp, 'm', {});
        if (r) m.fondo = r.degradado ? { degradado: r.degradado } : r.imagen ? { imagen: r.imagen } : r.color ? { color: r.color } : null;
        for (const h of hijos(mp)) {
          if (h.n === 'draw:frame' && (CLASES_PIE[atr(h, 'presentation:class')] || CLASES_MARCADOR.has(atr(h, 'presentation:class')))) continue;   // los marcadores no se dibujan en el patrón
          if (h.n === 'presentation:notes') continue;
          const f = conv.convertir(h, 'm', { pie: null, master: false, visibles: {} }, 0);
          if (f) m.formas.push(f);
        }
      }
      maestros.push(m);
      indiceMaestra.set(clave, maestros.length - 1);
      return maestros.length - 1;
    };

    const diapositivas = [];
    for (let i = 0; i < paginas.length; i++) {
      const pg = paginas[i];
      const nombreM = atr(pg, 'draw:master-page-name');
      const mi = maestraDe(nombreM);
      const mp = est.maestras.get(indiceMaestraNombre(indiceMaestra, mi));
      const dpE = est.resolver('drawing-page', atr(pg, 'draw:style-name'), 'c');
      const dp = dpE.dp;
      const d = { maestro: mi, formas: [] };
      if (atr(pg, 'presentation:visibility') === 'hidden') d.oculta = true;
      if (dp['presentation:background-objects-visible'] === 'false') d.ocultarPatron = true;
      const rf = conv.relleno(dp, 'c', {});
      if (rf) d.fondo = rf.degradado ? { degradado: rf.degradado } : rf.imagen ? { imagen: rf.imagen } : rf.color ? { color: rf.color } : undefined;
      else if (rf === null && dp['draw:fill'] === 'none') d.fondo = undefined;
      const visibles = { pie: dp['presentation:display-footer'] === 'true' || atr(pg, 'presentation:use-footer-name') && dp['presentation:display-footer'] !== 'false', fecha: dp['presentation:display-date-time'] === 'true', numero: dp['presentation:display-page-number'] === 'true' };
      const ctx = { pie: { pie: decl.pie.get(atr(pg, 'presentation:use-footer-name')) || '', encabezado: '' }, master: false, visibles };
      if (mp && (visibles.pie || visibles.fecha || visibles.numero)) {
        const ctxM = { pie: ctx.pie, master: true, visibles };
        for (const h of hijos(mp, 'draw:frame')) { if (CLASES_PIE[atr(h, 'presentation:class')]) { const f = conv.convertir(h, 'm', ctxM, 0); if (f) d.formas.push(f); } }
      }
      let notas = null;
      for (const h of hijos(pg)) {
        if (h.n === 'presentation:notes') {
          for (const fr of hijos(h, 'draw:frame')) {
            if (atr(fr, 'presentation:class') !== 'notes') continue;
            const tb = hijo(fr, 'draw:text-box');
            if (!tb) continue;
            const ps = conv.parrafos(tb, ctx, 'c', { t: {}, p: {}, lista: null }, 0, null, []);
            if (ps.some((p) => p.hijos.length)) notas = ps.map((p) => ({ hijos: p.hijos.map((x) => (x.texto !== undefined ? { texto: x.texto } : x)) }));
          }
          continue;
        }
        const f = conv.convertir(h, 'c', ctx, 0);
        if (f) d.formas.push(f);
      }
      if (notas) d.notas = notas;
      diapositivas.push(d);
    }

    const contrato = { ancho, alto, maestros, diapositivas, propiedades: {} };
    // Las imágenes se leen una por una (con tope de bytes)
    for (const im of conv.pendientes) {
      if (conv.bytesImagenes + tamanoDe(im._entrada) > MAX_IMAGENES_BYTES) throw new Error('Las imágenes de la presentación ocupan demasiado y no se puede leer.');
      im.datos = await im._entrada.async('uint8array');
      conv.bytesImagenes += im.datos.length;
      if (conv.bytesImagenes > MAX_IMAGENES_BYTES) throw new Error('Las imágenes de la presentación ocupan demasiado y no se puede leer.');
      delete im._entrada; delete im._ruta;
    }
    const meta = await texto('meta.xml');
    if (meta) {
      const m = arbol(meta);
      const md = hijo(hijo(m, 'office:meta'), 'dc:title');
      const t = md ? textoDe(md) : '';
      if (t) contrato.propiedades.titulo = t.slice(0, 255);
    }
    return { contrato, mensajes: [...conv.mensajes] };
  }
  const indiceMaestraNombre = (mapa, i) => { for (const [k, v] of mapa) if (v === i) return k; return null; };

  async function aPptx(bytes) {
    const { contrato, mensajes } = await leer(bytes);
    const r = await PptxDesdeContrato.generar(contrato);
    return { bytes: r.bytes, mensajes: [...new Set([...mensajes, ...r.mensajes])] };
  }

  globalThis.OdpLector = { parece, leer, aPptx };
})();

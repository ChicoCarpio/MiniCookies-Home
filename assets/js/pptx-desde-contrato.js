/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Una presentación descrita con el «contrato de presentación» (el que entregan PptLector y OdpLector, y el que recibirán los escritores de .ppt y .odp) escrita como .pptx.
// Es el gemelo de DocxDesdeContrato para PowerPoint: cualquier lector → cualquier escritor, sin un conversor por pareja, y todas las herramientas de PowerPoint
// (que leen .pptx) aceptan los formatos viejos sin cambios.
//
//   const { bytes, mensajes } = await PptxDesdeContrato.generar(contrato);   // Uint8Array con el .pptx
//
//   contrato = { propiedades?: { titulo, autor }, ancho, alto,                          // tamaño de la diapositiva en EMU (1 pulgada = 914400)
//                fuente?: 'Arial', tamano?: 18,                                          // fuente y tamaño (pt) de lo que no dice otra cosa
//                maestros?: [{ fondo?, formas: [forma] }],                               // el patrón: lo que se dibuja detrás de las diapositivas que lo usan
//                diapositivas: [{ maestro?: índice, oculta?, ocultarPatron? (no dibuja las formas del patrón), fondo?, formas: [forma], notas?: [parrafo] }] }
//   fondo  = { color: 'RRGGBB' } | { degradado } | { imagen }
//   forma  = { tipo: 'forma', geom?: 'rect'|… (nombre de forma predefinida de ECMA-376) | { rutas: [{ w, h, cmds: [['M',x,y]|['L',x,y]|['C',x1,y1,x2,y2,x,y]|['Q',x1,y1,x,y]|['Z']], sinRelleno?, sinTrazo? }] } (forma libre), ajustes?: [enteros de ECMA], relleno?, linea?, sombra?, texto?, ...caja }
//          | { tipo: 'imagen', imagen, recorte?: { izq, arr, der, aba } (fracciones 0..1), linea?, geom?, ...caja }
//          | { tipo: 'linea', linea?, ...caja }                                         // de la esquina superior izquierda a la inferior derecha de la caja (con volteo, la otra diagonal)
//          | { tipo: 'tabla', columnas: [ancho EMU], filas: [{ alto, celdas: [celda] }], ...caja }
//          | { tipo: 'grupo', hijos: [forma], hx?, hy?, hw?, hh?, ...caja }              // h*: el sistema de coordenadas de los hijos (por defecto, el de la caja)
//   caja   = { x, y, w, h, rot?: grados, volteoH?, volteoV?, nombre?, descripcion?, enlace? }    (todo en EMU)
//   enlace = 'https://…' | 'mailto:…' | { diapositiva: número desde 1 }
//   celda  = { texto?, relleno?, bordes?: { arriba|abajo|izquierda|derecha: linea|null }, columnas?: n, filas?: n, fusionada?: true, ancla?, margenes? }   // fusionada: la cubre otra celda
//   relleno = null (sin relleno) | { color: 'RRGGBB', alfa?: 0..1 } | { degradado } | { imagen }
//   degradado = { angulo?: grados (0 = de izquierda a derecha, 90 = de arriba abajo), radial?: true, paradas: [{ pos: 0..100, color, alfa? }] }
//   linea  = null | { color: 'RRGGBB', alfa?, ancho (EMU), guion?: 'solid'|'dash'|'dot'|'dashDot'|'lgDash'|'sysDash'|'sysDot', inicio?: 'triangle'|'arrow'|'stealth'|'diamond'|'oval', fin?, union?: 'round'|'bevel'|'miter', cap?: 'rnd'|'sq'|'flat' }
//   sombra = { color, alfa?, x, y (EMU, desplazamiento), desenfoque? (EMU) }
//   imagen = { datos: Uint8Array, tipo?: 'png'|'jpeg'|'gif'|'bmp' }                      // el mismo arreglo (por identidad) se guarda una sola vez
//   texto  = { parrafos: [parrafo], ancla?: 't'|'ctr'|'b', margenes?: { i, s, d, a } (EMU), ajuste?: false (sin partir líneas), autoajuste?: 'forma'|'texto', vertical?: true, escala?: 0..1 (texto reducido) }
//   parrafo = { alin?: 'l'|'ctr'|'r'|'just', nivel?: 0..8, izq?, primera? (EMU; primera < 0 = sangría francesa), antes?, despues? (pt), interlineado?: { pct } | { pt },
//               vineta?: null | { car, fuente?, color?, tamano? (% del texto) } | { num: 'arabicPeriod'|'alphaLcPeriod'|'alphaUcPeriod'|'romanLcPeriod'|'romanUcPeriod'|'arabicParenR'…, inicio? },
//               hijos: [run], r?: propsR del fin del párrafo }
//   run    = { texto, r? } | { salto: true } | { campo: 'numero'|'fecha', texto?, r? }     // r y enlace en el propio run: { texto, r, enlace }
//   propsR = { fuente?, tamano? (pt), negrita?, cursiva?, subrayado?, tachado?, color?: 'RRGGBB', alfa?, sup?, sub? }
//
// Requiere zip-propio.js y escritor-pptx.js (el tema de Office). El patrón escrito no trae marcadores de posición: lo que el lector resolvió (posición, tamaños, viñetas)
// va en cada forma, así que la diapositiva se ve igual sin depender de la herencia de PowerPoint.
(function () {
  'use strict';
  const CABECERA = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const CT = 'application/vnd.openxmlformats-officedocument.presentationml';
  const INVALIDOS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;
  const limpiar = (s) => String(s == null ? '' : s).replace(INVALIDOS, '');
  const esc = (s) => limpiar(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const escAttr = (s) => esc(s).replace(/"/g, '&quot;');
  const entero = (n, min = -2147483647, max = 2147483647) => { const v = Math.round(Number(n)); return Number.isFinite(v) ? Math.max(min, Math.min(max, v)) : 0; };
  const hex6 = (c) => (/^[0-9a-fA-F]{6}$/.test(c || '') ? c.toUpperCase() : null);
  const sinProto = (o) => Object.assign(Object.create(null), o);
  const MAX_DIAPOSITIVAS = 10000;
  const MAX_FORMAS = 150000;          // formas de toda la presentación (cada grupo, celda de tabla y hoja de cálculo cuenta)
  const MAX_PROFUNDIDAD = 16;
  const MAX_TEXTO = 64 * 1024 * 1024;
  const MAX_CELDAS = 200000;
  const MAX_UNIDADES = 3000000;       // párrafos y tramos de texto de toda la presentación
  const MAX_PUNTOS = 1000000;         // órdenes de trazo de todas las formas libres

  const GEOMETRIAS = new Set(['rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle', 'diamond', 'parallelogram', 'trapezoid', 'pentagon', 'hexagon', 'heptagon', 'octagon', 'decagon', 'dodecagon', 'star4', 'star5', 'star6', 'star7', 'star8', 'star10', 'star12', 'star16', 'star24', 'star32',
    'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'leftRightArrow', 'upDownArrow', 'quadArrow', 'leftRightUpArrow', 'bentArrow', 'uturnArrow', 'leftUpArrow', 'bentUpArrow', 'curvedRightArrow', 'curvedLeftArrow', 'curvedUpArrow', 'curvedDownArrow', 'stripedRightArrow', 'notchedRightArrow', 'homePlate', 'chevron',
    'rightArrowCallout', 'leftArrowCallout', 'upArrowCallout', 'downArrowCallout', 'leftRightArrowCallout', 'upDownArrowCallout', 'quadArrowCallout', 'circularArrow', 'cube', 'can', 'bevel', 'donut', 'noSmoking', 'blockArc', 'foldedCorner', 'smileyFace', 'heart', 'lightningBolt', 'sun', 'moon', 'cloud', 'arc',
    'plaque', 'frame', 'halfFrame', 'corner', 'diagStripe', 'chord', 'pie', 'teardrop', 'plus', 'bracketPair', 'bracePair', 'leftBracket', 'rightBracket', 'leftBrace', 'rightBrace', 'wedgeRectCallout', 'wedgeRoundRectCallout', 'wedgeEllipseCallout', 'cloudCallout',
    'borderCallout1', 'borderCallout2', 'borderCallout3', 'accentCallout1', 'callout1', 'accentBorderCallout1', 'ribbon', 'ribbon2', 'ellipseRibbon', 'ellipseRibbon2', 'verticalScroll', 'horizontalScroll', 'wave', 'doubleWave', 'flowChartProcess', 'flowChartAlternateProcess', 'flowChartDecision',
    'flowChartInputOutput', 'flowChartPredefinedProcess', 'flowChartInternalStorage', 'flowChartDocument', 'flowChartMultidocument', 'flowChartTerminator', 'flowChartPreparation', 'flowChartManualInput', 'flowChartManualOperation', 'flowChartConnector', 'flowChartOffpageConnector', 'flowChartPunchedCard',
    'flowChartPunchedTape', 'flowChartSummingJunction', 'flowChartOr', 'flowChartCollate', 'flowChartSort', 'flowChartExtract', 'flowChartMerge', 'flowChartOnlineStorage', 'flowChartDelay', 'flowChartMagneticTape', 'flowChartMagneticDisk', 'flowChartMagneticDrum', 'flowChartDisplay',
    'actionButtonBlank', 'actionButtonHome', 'actionButtonHelp', 'actionButtonInformation', 'actionButtonBackPrevious', 'actionButtonForwardNext', 'actionButtonBeginning', 'actionButtonEnd', 'actionButtonReturn', 'actionButtonDocument', 'actionButtonSound', 'actionButtonMovie',
    'line', 'straightConnector1', 'bentConnector3', 'curvedConnector3', 'mathPlus', 'mathMinus', 'mathMultiply', 'mathDivide', 'mathEqual', 'mathNotEqual', 'gear6', 'gear9', 'funnel', 'cornerTabs', 'squareTabs', 'plaqueTabs', 'snip1Rect', 'snip2SameRect', 'snip2DiagRect', 'snipRoundRect', 'round1Rect', 'round2SameRect', 'round2DiagRect', 'bracePair']);
  const ANCLAS = sinProto({ t: 't', ctr: 'ctr', b: 'b' });
  const ALINEACION = sinProto({ l: 'l', ctr: 'ctr', r: 'r', just: 'just' });
  const GUIONES = new Set(['solid', 'dash', 'dot', 'dashDot', 'lgDash', 'lgDashDot', 'sysDash', 'sysDot', 'sysDashDot', 'lgDashDotDot', 'sysDashDotDot']);
  const FLECHAS = new Set(['triangle', 'arrow', 'stealth', 'diamond', 'oval']);
  const NUMERACIONES = new Set(['arabicPeriod', 'arabicParenR', 'arabicParenBoth', 'arabicPlain', 'alphaLcPeriod', 'alphaUcPeriod', 'alphaLcParenR', 'alphaUcParenR', 'romanLcPeriod', 'romanUcPeriod', 'romanLcParenR', 'romanUcParenR', 'alphaLcParenBoth', 'alphaUcParenBoth', 'romanLcParenBoth', 'romanUcParenBoth', 'circleNumDbPlain']);
  const MIME = sinProto({ png: 'image/png', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', emf: 'image/x-emf', wmf: 'image/x-wmf' });
  const URL_PERMITIDA = /^(https?:\/\/|mailto:)/i;

  function tipoDeImagen(b, indicado) {
    if (b.length > 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4E && b[3] === 0x47) return 'png';
    if (b.length > 3 && b[0] === 0xFF && b[1] === 0xD8) return 'jpeg';
    if (b.length > 6 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return 'gif';
    if (b.length > 2 && b[0] === 0x42 && b[1] === 0x4D) return 'bmp';
    return indicado === 'jpg' ? 'jpeg' : MIME[indicado] ? indicado : null;
  }

  class Contexto {
    constructor(contrato) {
      this.contrato = contrato;
      this.mensajes = new Set();
      this.formas = 0;
      this.puntos = 0;
      this.unidades = 0;
      this.texto = 0;
      this.medios = [];                       // { nombre, bytes, tipo }
      this.mediosPorDatos = new Map();
    }
    aviso(clave) { this.mensajes.add(clave); }
    contar(n = 1) { if ((this.formas += n) > MAX_FORMAS) throw new Error('La presentación tiene demasiadas formas y no se puede escribir.'); }
    medio(imagen) {
      if (!imagen || !(imagen.datos instanceof Uint8Array) || !imagen.datos.length) return null;
      const tipo = tipoDeImagen(imagen.datos, imagen.tipo);
      if (!tipo) { this.aviso('imagen-formato'); return null; }
      let m = this.mediosPorDatos.get(imagen.datos);
      if (!m) { m = { nombre: `image${this.medios.length + 1}.${tipo}`, bytes: imagen.datos, tipo }; this.medios.push(m); this.mediosPorDatos.set(imagen.datos, m); }
      return m;
    }
  }
  // Las relaciones de una parte (diapositiva, patrón o notas): imágenes y enlaces
  class Rels {
    constructor(ctx, base) { this.ctx = ctx; this.lista = []; this.porDestino = new Map(); this.base = base || []; }
    agregar(tipo, destino, externo) {
      const clave = tipo + '|' + destino + '|' + (externo ? 1 : 0);
      let id = this.porDestino.get(clave);
      if (!id) { id = 'rId' + (this.lista.length + 1 + this.base.length); this.lista.push([id, tipo, destino, externo]); this.porDestino.set(clave, id); }
      return id;
    }
    imagen(imagen) { const m = this.ctx.medio(imagen); return m ? this.agregar(`${REL}/image`, `../media/${m.nombre}`) : null; }
    xml() {
      return `${CABECERA}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${this.base.concat(this.lista).map(([id, tipo, destino, externo]) => `<Relationship Id="${id}" Type="${tipo}" Target="${escAttr(destino)}"${externo ? ' TargetMode="External"' : ''}/>`).join('')}</Relationships>`;
    }
  }

  // ───────── Colores, rellenos y líneas ─────────
  const color = (c, alfa) => {
    const h = hex6(c) || '000000';
    const a = typeof alfa === 'number' && alfa >= 0 && alfa < 1 ? `<a:alpha val="${entero(alfa * 100000, 0, 100000)}"/>` : '';
    return a ? `<a:srgbClr val="${h}">${a}</a:srgbClr>` : `<a:srgbClr val="${h}"/>`;
  };
  function degradado(g) {
    const paradas = (g.paradas || []).slice(0, 32).filter((p) => p && typeof p === 'object');
    if (paradas.length < 2) return paradas.length ? `<a:solidFill>${color(paradas[0].color, paradas[0].alfa)}</a:solidFill>` : '';
    const lista = paradas.map((p) => `<a:gs pos="${entero((Number(p.pos) || 0) * 1000, 0, 100000)}">${color(p.color, p.alfa)}</a:gs>`).join('');
    const forma = g.radial ? '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>' : `<a:lin ang="${entero(((Number(g.angulo) || 0) % 360 + 360) % 360 * 60000, 0, 21600000)}" scaled="0"/>`;
    return `<a:gradFill rotWithShape="1"><a:gsLst>${lista}</a:gsLst>${forma}</a:gradFill>`;
  }
  function relleno(r, rels) {
    if (r === null) return '<a:noFill/>';
    if (!r || typeof r !== 'object') return '';
    if (r.degradado) return degradado(r.degradado);
    if (r.imagen) { const id = rels.imagen(r.imagen); return id ? `<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${id}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>` : ''; }
    if (r.color) return `<a:solidFill>${color(r.color, r.alfa)}</a:solidFill>`;
    return '';
  }
  function linea(l, etiqueta = 'a:ln') {
    if (l === null) return `<${etiqueta}><a:noFill/></${etiqueta}>`;
    if (!l || typeof l !== 'object') return '';
    const ancho = entero(l.ancho, 0, 20116800);
    const guion = GUIONES.has(l.guion) && l.guion !== 'solid' ? `<a:prstDash val="${l.guion}"/>` : '';
    const union = l.union === 'round' || l.union === 'bevel' ? `<a:${l.union}/>` : l.union === 'miter' ? '<a:miter lim="800000"/>' : '';
    const punta = (n, v) => (FLECHAS.has(v) ? `<a:${n} type="${v}"/>` : '');
    const cap = l.cap === 'rnd' || l.cap === 'sq' || l.cap === 'flat' ? ` cap="${l.cap}"` : '';
    return `<${etiqueta} w="${ancho}"${cap}><a:solidFill>${color(l.color, l.alfa)}</a:solidFill>${guion}${union}${punta('headEnd', l.inicio)}${punta('tailEnd', l.fin)}</${etiqueta}>`;
  }
  function sombra(s) {
    if (!s || typeof s !== 'object') return '';
    const x = Number(s.x) || 0, y = Number(s.y) || 0;
    const dist = entero(Math.hypot(x, y), 0, 27273042), dir = entero(((Math.atan2(y, x) * 180 / Math.PI) % 360 + 360) % 360 * 60000, 0, 21600000);
    return `<a:effectLst><a:outerShdw blurRad="${entero(s.desenfoque, 0, 27273042)}" dist="${dist}" dir="${dir}" algn="tl" rotWithShape="0">${color(s.color, typeof s.alfa === 'number' ? s.alfa : 0.5)}</a:outerShdw></a:effectLst>`;
  }

  // ───────── Texto ─────────
  function enlaceXml(e, rels, ctx, etiqueta = 'a:hlinkClick') {
    if (!e) return '';
    if (typeof e === 'string') {
      if (!URL_PERMITIDA.test(e) || e.length > 2000) { ctx.aviso('enlace-descartado'); return ''; }
      return `<${etiqueta} r:id="${rels.agregar(`${REL}/hyperlink`, limpiar(e), true)}"/>`;
    }
    if (typeof e === 'object' && Number.isInteger(e.diapositiva) && e.diapositiva >= 1 && e.diapositiva <= (ctx.contrato.diapositivas || []).length) {
      // ruta desde la carpeta de la parte (slides, slideMasters o notesSlides): «slideN.xml» solo sirve desde slides
      return `<${etiqueta} r:id="${rels.agregar(`${REL}/slide`, `../slides/slide${e.diapositiva}.xml`)}" action="ppaction://hlinksldjump"/>`;
    }
    return '';
  }
  function propsR(r, rels, ctx, lang, enlace) {
    r = r || {};
    let atr = ` lang="${lang}"`;
    if (typeof r.tamano === 'number' && r.tamano > 0) atr += ` sz="${entero(r.tamano * 100, 100, 400000)}"`;
    if (r.negrita) atr += ' b="1"';
    if (r.cursiva) atr += ' i="1"';
    if (r.subrayado) atr += ' u="sng"';
    if (r.tachado) atr += ' strike="sngStrike"';
    if (r.sup) atr += ' baseline="30000"'; else if (r.sub) atr += ' baseline="-25000"';
    atr += ' dirty="0"';
    let hijos = '';
    if (hex6(r.color)) hijos += `<a:solidFill>${color(r.color, r.alfa)}</a:solidFill>`;
    if (r.fuente && typeof r.fuente === 'string') { const f = escAttr(r.fuente.slice(0, 64)); hijos += `<a:latin typeface="${f}"/><a:ea typeface="${f}"/><a:cs typeface="${f}"/>`; }
    hijos += enlaceXml(enlace, rels, ctx);
    return hijos ? `<a:rPr${atr}>${hijos}</a:rPr>` : `<a:rPr${atr}/>`;
  }
  function parrafo(p, rels, ctx, lang) {
    p = p || {};
    let atr = '';
    if (Number.isFinite(p.izq)) atr += ` marL="${entero(p.izq, 0, 51206400)}"`;
    if (Number.isFinite(p.primera)) atr += ` indent="${entero(p.primera, -51206400, 51206400)}"`;
    if (ALINEACION[p.alin]) atr += ` algn="${ALINEACION[p.alin]}"`;
    const nivel = entero(p.nivel, 0, 8);
    if (nivel) atr += ` lvl="${nivel}"`;
    let hijos = '';
    const interl = p.interlineado;
    if (interl && typeof interl.pct === 'number') hijos += `<a:lnSpc><a:spcPct val="${entero(interl.pct * 1000, 0, 13200000)}"/></a:lnSpc>`;
    else if (interl && typeof interl.pt === 'number') hijos += `<a:lnSpc><a:spcPts val="${entero(interl.pt * 100, 0, 1584000)}"/></a:lnSpc>`;
    const espacio = (v, e) => (typeof v === 'number' ? `<a:${e}><a:spcPts val="${entero(v * 100, 0, 1584000)}"/></a:${e}>` : v && typeof v.pct === 'number' ? `<a:${e}><a:spcPct val="${entero(v.pct * 1000, 0, 13200000)}"/></a:${e}>` : '');
    hijos += espacio(p.antes, 'spcBef') + espacio(p.despues, 'spcAft');
    const v = p.vineta;
    if (v === null) hijos += '<a:buNone/>';
    else if (v && v.num && NUMERACIONES.has(v.num)) {
      if (hex6(v.color)) hijos += `<a:buClr>${color(v.color)}</a:buClr>`;
      hijos += `<a:buFont typeface="+mj-lt"/><a:buAutoNum type="${v.num}"${Number.isInteger(v.inicio) && v.inicio > 1 ? ` startAt="${Math.min(v.inicio, 32767)}"` : ''}/>`;
    } else if (v && typeof v.car === 'string' && v.car) {
      if (hex6(v.color)) hijos += `<a:buClr>${color(v.color)}</a:buClr>`;
      if (typeof v.tamano === 'number') hijos += `<a:buSzPct val="${entero(v.tamano * 1000, 25000, 400000)}"/>`;
      hijos += `<a:buFont typeface="${escAttr((v.fuente || 'Arial').slice(0, 64))}"/><a:buChar char="${escAttr(Array.from(v.car)[0])}"/>`;
    }
    const salida = [];
    if ((ctx.unidades += 1 + (p.hijos || []).length) > MAX_UNIDADES) throw new Error('La presentación tiene demasiado texto y no se puede escribir.');
    for (const run of p.hijos || []) {
      if (!run || typeof run !== 'object') continue;
      if (run.salto) { salida.push(`<a:br>${propsR(run.r, rels, ctx, lang)}</a:br>`); continue; }
      if ((ctx.texto += String(run.texto || '').length) > MAX_TEXTO) throw new Error('La presentación tiene demasiado texto y no se puede escribir.');
      if (run.campo) {
        const tipo = run.campo === 'numero' ? 'slidenum' : 'datetime1';
        const id = run.campo === 'numero' ? '{B0A0BC7D-E01F-4586-9EED-2E3260ECE181}' : '{5E4B1A11-67A4-4E0C-8E5A-8D5D3F0A9C11}';
        salida.push(`<a:fld id="${id}" type="${tipo}">${propsR(run.r, rels, ctx, lang)}<a:t>${esc(run.texto || (run.campo === 'numero' ? '‹#›' : ''))}</a:t></a:fld>`);
        continue;
      }
      // '\n' y '\v' dentro del texto son saltos de línea; '\t' se queda (PowerPoint lo respeta)
      const partes = limpiar(run.texto).replace(/\r\n?|\u000B/g, '\n').split('\n');
      partes.forEach((parte, n) => {
        if (n > 0) salida.push(`<a:br>${propsR(run.r, rels, ctx, lang)}</a:br>`);
        if (parte !== '') salida.push(`<a:r>${propsR(run.r, rels, ctx, lang, run.enlace)}<a:t>${esc(parte)}</a:t></a:r>`);
      });
    }
    const fin = `<a:endParaRPr${p.r && typeof p.r.tamano === 'number' ? ` sz="${entero(p.r.tamano * 100, 100, 400000)}"` : ''} lang="${lang}" dirty="0"/>`;
    return `<a:p>${hijos || atr ? `<a:pPr${atr}>${hijos}</a:pPr>` : ''}${salida.join('')}${fin}</a:p>`;
  }
  function cuerpo(t, rels, ctx, lang, etiqueta = 'p:txBody', sencillo = false) {
    t = t || {};
    const m = t.margenes || {};
    const d = (v, def) => (Number.isFinite(v) ? entero(v, 0, 51206400) : def);
    let atr = ` wrap="${t.ajuste === false ? 'none' : 'square'}" lIns="${d(m.i, 91440)}" tIns="${d(m.s, 45720)}" rIns="${d(m.d, 91440)}" bIns="${d(m.a, 45720)}" rtlCol="0"`;
    if (ANCLAS[t.ancla]) atr += ` anchor="${ANCLAS[t.ancla]}"`;
    if (t.vertical) atr += ' vert="vert270"';
    const auto = t.autoajuste === 'forma' ? '<a:spAutoFit/>' : t.autoajuste === 'texto' ? `<a:normAutofit${typeof t.escala === 'number' && t.escala > 0 && t.escala < 1 ? ` fontScale="${entero(t.escala * 100000, 1000, 100000)}"` : ''}/>` : '<a:noAutofit/>';
    const parrafos = (t.parrafos || []).filter((p) => p && typeof p === 'object').map((p) => parrafo(p, rels, ctx, lang));
    return `<${etiqueta}>${sencillo ? '<a:bodyPr/>' : `<a:bodyPr${atr}>${auto}</a:bodyPr>`}<a:lstStyle/>${parrafos.join('') || `<a:p><a:endParaRPr lang="${lang}" dirty="0"/></a:p>`}</${etiqueta}>`;
  }

  // ───────── Formas ─────────
  const caja = (f) => {
    const w = entero(f.w, 0, 51206400), h = entero(f.h, 0, 51206400);
    const rot = Number.isFinite(f.rot) && f.rot ? ` rot="${entero(((f.rot % 360) + 360) % 360 * 60000, 0, 21599999)}"` : '';
    return `<a:xfrm${rot}${f.volteoH ? ' flipH="1"' : ''}${f.volteoV ? ' flipV="1"' : ''}><a:off x="${entero(f.x, -51206400, 51206400)}" y="${entero(f.y, -51206400, 51206400)}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>`;
  };
  const geometria = (g, ajustes, ctx) => {
    if (g && typeof g === 'object' && Array.isArray(g.rutas)) return geometriaLibre(g, ctx);
    const prst = GEOMETRIAS.has(g) ? g : 'rect';
    const av = Array.isArray(ajustes) ? ajustes.slice(0, 8).map((v, i) => (Number.isFinite(v) ? `<a:gd name="adj${ajustes.length === 1 ? '' : i + 1}" fmla="val ${entero(v, -2147483647, 2147483647)}"/>` : '')).join('') : '';
    return `<a:prstGeom prst="${prst}"><a:avLst>${av}</a:avLst></a:prstGeom>`;
  };
  // Forma libre: rutas de trazos en su propio sistema de coordenadas (w × h), que se estira a la caja de la forma
  function geometriaLibre(g, ctx) {
    const rutas = [];
    for (const r of g.rutas.slice(0, 64)) {
      if (!r || !Array.isArray(r.cmds)) continue;
      const w = entero(r.w, 1, 2147483647), h = entero(r.h, 1, 2147483647);
      const xy = (x, y) => `<a:pt x="${entero(x, -2147483647, 2147483647)}" y="${entero(y, -2147483647, 2147483647)}"/>`;
      const cmds = [];
      for (const c of r.cmds) {
        if ((ctx.puntos += 1) > MAX_PUNTOS) throw new Error('La presentación tiene demasiados trazos y no se puede escribir.');
        if (!Array.isArray(c)) continue;
        if (c[0] === 'M') cmds.push(`<a:moveTo>${xy(c[1], c[2])}</a:moveTo>`);
        else if (c[0] === 'L') cmds.push(`<a:lnTo>${xy(c[1], c[2])}</a:lnTo>`);
        else if (c[0] === 'C') cmds.push(`<a:cubicBezTo>${xy(c[1], c[2])}${xy(c[3], c[4])}${xy(c[5], c[6])}</a:cubicBezTo>`);
        else if (c[0] === 'Q') cmds.push(`<a:quadBezTo>${xy(c[1], c[2])}${xy(c[3], c[4])}</a:quadBezTo>`);
        else if (c[0] === 'Z') cmds.push('<a:close/>');
      }
      if (cmds.length) rutas.push(`<a:path w="${w}" h="${h}"${r.sinRelleno ? ' fill="none"' : ''}${r.sinTrazo ? ' stroke="0"' : ''}>${cmds.join('')}</a:path>`);
    }
    if (!rutas.length) return '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
    return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>${rutas.join('')}</a:pathLst></a:custGeom>`;
  }
  const nombreDe = (f, ctx, base) => escAttr(String(f.nombre || base).slice(0, 80));
  function cNvPr(f, id, base, rels, ctx) {
    const desc = f.descripcion ? ` descr="${escAttr(String(f.descripcion).slice(0, 500))}"` : '';
    const enl = enlaceXml(f.enlace, rels, ctx);
    return `<p:cNvPr id="${id}" name="${nombreDe(f, ctx, base + ' ' + (id - 1))}"${desc}${enl ? `>${enl}</p:cNvPr>` : '/>'}`;
  }

  function forma(f, estado, rels, ctx, lang, prof) {
    if (!f || typeof f !== 'object') return '';
    ctx.contar();
    const id = ++estado.id;
    switch (f.tipo) {
      case 'forma': {
        const efecto = sombra(f.sombra);
        const sp = `<p:spPr>${caja(f)}${geometria(f.geom, f.ajustes, ctx)}${relleno(f.relleno, rels)}${linea(f.linea)}${efecto}</p:spPr>`;
        const tx = f.texto ? cuerpo(f.texto, rels, ctx, lang) : '';
        return `<p:sp><p:nvSpPr>${cNvPr(f, id, f.texto ? 'Cuadro de texto' : 'Forma', rels, ctx)}<p:cNvSpPr${f.texto && f.geom === 'rect' && f.relleno === undefined && f.linea === undefined ? ' txBox="1"' : ''}/><p:nvPr/></p:nvSpPr>${sp}${tx}</p:sp>`;
      }
      case 'imagen': {
        const rid = rels.imagen(f.imagen);
        if (!rid) return '';
        const c = f.recorte || {};
        const src = ['izq', 'arr', 'der', 'aba'].map((k, i) => (Number(c[k]) ? ` ${'ltrb'[i]}="${entero(c[k] * 100000, -100000, 100000)}"` : '')).join('');
        return `<p:pic><p:nvPicPr>${cNvPr(f, id, 'Imagen', rels, ctx)}<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/><a:srcRect${src}/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${caja(f)}${geometria(f.geom || 'rect', f.ajustes, ctx)}${linea(f.linea)}</p:spPr></p:pic>`;
      }
      case 'linea':
        return `<p:cxnSp><p:nvCxnSpPr>${cNvPr(f, id, 'Línea', rels, ctx)}<p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>${caja(f)}<a:prstGeom prst="line"><a:avLst/></a:prstGeom>${linea(f.linea) || '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>'}${sombra(f.sombra)}</p:spPr></p:cxnSp>`;
      case 'grupo': {
        if (prof >= MAX_PROFUNDIDAD) { ctx.aviso('grupos-anidados'); return ''; }
        const hijos = (f.hijos || []).map((h) => forma(h, estado, rels, ctx, lang, prof + 1)).join('');
        const hx = Number.isFinite(f.hx) ? f.hx : f.x, hy = Number.isFinite(f.hy) ? f.hy : f.y, hw = Number.isFinite(f.hw) ? f.hw : f.w, hh = Number.isFinite(f.hh) ? f.hh : f.h;
        const x = caja(f).replace('</a:xfrm>', `<a:chOff x="${entero(hx, -51206400, 51206400)}" y="${entero(hy, -51206400, 51206400)}"/><a:chExt cx="${entero(hw, 0, 51206400)}" cy="${entero(hh, 0, 51206400)}"/></a:xfrm>`);
        return `<p:grpSp><p:nvGrpSpPr>${cNvPr(f, id, 'Grupo', rels, ctx)}<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${x}</p:grpSpPr>${hijos}</p:grpSp>`;
      }
      case 'tabla': return tabla(f, id, rels, ctx, lang);
      default: return '';
    }
  }

  function tabla(f, id, rels, ctx, lang) {
    const columnas = (f.columnas || []).slice(0, 63).map((w) => entero(w, 1, 51206400));
    const filas = (f.filas || []).slice(0, 5000);
    if (!columnas.length || !filas.length) return '';
    const celdasTotal = filas.reduce((s, fi) => s + ((fi && fi.celdas) || []).length, 0);
    if (celdasTotal > MAX_CELDAS) throw new Error('La tabla tiene demasiadas celdas y no se puede escribir.');
    ctx.contar(celdasTotal);
    const ancho = columnas.reduce((a, b) => a + b, 0), alto = filas.reduce((a, fi) => a + entero(fi && fi.alto, 0, 51206400), 0);
    const trs = filas.map((fi) => {
      const celdas = (fi && fi.celdas) || [];
      const tcs = [];
      for (let c = 0; c < columnas.length; c++) {
        const cel = celdas[c] || {};
        const m = cel.margenes || {};
        let atr = '';
        if (cel.columnas > 1) atr += ` gridSpan="${Math.min(entero(cel.columnas), 63)}"`;
        if (cel.filas > 1) atr += ` rowSpan="${Math.min(entero(cel.filas), 5000)}"`;
        if (cel.fusionada === 'h') atr += ' hMerge="1"'; else if (cel.fusionada === 'v') atr += ' vMerge="1"'; else if (cel.fusionada) atr += ' hMerge="1"';
        const b = cel.bordes || {};
        const bordes = [['izquierda', 'a:lnL'], ['derecha', 'a:lnR'], ['arriba', 'a:lnT'], ['abajo', 'a:lnB']].map(([k, e]) => (k in b ? linea(b[k], e) : '')).join('');
        const lado = (v, nombre) => (Number.isFinite(v) ? ` ${nombre}="${entero(v, 0, 51206400)}"` : '');
        const tcPr = `<a:tcPr${lado(m.i, 'marL')}${lado(m.d, 'marR')}${lado(m.s, 'marT')}${lado(m.a, 'marB')}${ANCLAS[cel.ancla] ? ` anchor="${ANCLAS[cel.ancla]}"` : ''}>${bordes}${relleno(cel.relleno, rels)}</a:tcPr>`;
        tcs.push(`<a:tc${atr}>${cuerpo(cel.texto, rels, ctx, lang, 'a:txBody', true)}${tcPr}</a:tc>`);
      }
      return `<a:tr h="${entero(fi && fi.alto, 0, 51206400)}">${tcs.join('')}</a:tr>`;
    }).join('');
    return `<p:graphicFrame><p:nvGraphicFramePr>${cNvPr(f, id, 'Tabla', rels, ctx)}<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>`
      + `<p:xfrm><a:off x="${entero(f.x, -51206400, 51206400)}" y="${entero(f.y, -51206400, 51206400)}"/><a:ext cx="${entero(ancho, 0, 51206400)}" cy="${entero(alto, 0, 51206400)}"/></p:xfrm>`
      + `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid>${columnas.map((w) => `<a:gridCol w="${w}"/>`).join('')}</a:tblGrid>${trs}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
  }

  const GRUPO = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
  function fondo(fd, rels) {
    if (!fd || typeof fd !== 'object') return '';
    const r = relleno(fd.degradado ? { degradado: fd.degradado } : fd.imagen ? { imagen: fd.imagen } : fd.color ? { color: fd.color } : undefined, rels);
    return r ? `<p:bg><p:bgPr>${r}<a:effectLst/></p:bgPr></p:bg>` : '';
  }
  function arbol(formas, ctx, rels, lang) {
    const estado = { id: 1 };
    return GRUPO + (formas || []).map((f) => forma(f, estado, rels, ctx, lang, 0)).join('');
  }

  // ───────── Partes fijas ─────────
  const estiloNivel = (n) => `<a:lvl${n}pPr marL="${457200 * (n - 1)}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${n}pPr>`;
  const niveles = Array.from({ length: 9 }, (_, i) => estiloNivel(i + 1)).join('');
  const CLR_MAP = '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
  const fuentesTema = (f) => `<a:fontScheme name="Office"><a:majorFont><a:latin typeface="${escAttr(f)}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="${escAttr(f)}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>`;
  const COLORES_TEMA = '<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
    + '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>'
    + '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>';

  function patron(m, ctx, rels, lang) {
    return `${CABECERA}<p:sldMaster ${NS}><p:cSld>${fondo(m.fondo, rels) || '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'}<p:spTree>${arbol(m.formas, ctx, rels, lang)}</p:spTree></p:cSld>${CLR_MAP}<p:sldLayoutIdLst><p:sldLayoutId id="__LAYOUT_ID__" r:id="rId1"/></p:sldLayoutIdLst>`
      + `<p:txStyles><p:titleStyle>${estiloNivel(1)}</p:titleStyle><p:bodyStyle>${niveles}</p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang="${lang}"/></a:defPPr>${niveles}</p:otherStyle></p:txStyles></p:sldMaster>`;
  }
  const diseno = () => `${CABECERA}<p:sldLayout ${NS} type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>${GRUPO}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

  function patronNotas() {
    const sp = (id, nombre, tipo, x, y, w, h, cuerpoXml) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${nombre}"/><p:cNvSpPr><a:spLocks noGrp="1"${tipo === 'sldImg' ? ' noRot="1" noChangeAspect="1"' : ''}/></p:cNvSpPr><p:nvPr><p:ph type="${tipo}"${tipo === 'body' ? ' sz="quarter" idx="1"' : ''}/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${tipo === 'sldImg' ? '<a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln>' : ''}</p:spPr>${cuerpoXml}</p:sp>`;
    return `${CABECERA}<p:notesMaster ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${GRUPO}`
      + sp(2, 'Slide Image Placeholder 1', 'sldImg', 685800, 1143000, 5486400, 3086100, '')
      + sp(3, 'Notes Placeholder 2', 'body', 685800, 4400550, 5486400, 3600450, '<p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>Notes</a:t></a:r></a:p></p:txBody>')
      + `</p:spTree></p:cSld>${CLR_MAP}<p:notesStyle>${niveles}</p:notesStyle></p:notesMaster>`;
  }

  async function generar(contrato) {
    if (!contrato || typeof contrato !== 'object') throw new Error('PptxDesdeContrato: falta el contrato');
    const ancho = entero(contrato.ancho, 0, 51206400) || 9144000, alto = entero(contrato.alto, 0, 51206400) || 6858000;
    const diapositivas = (contrato.diapositivas || []).slice(0, MAX_DIAPOSITIVAS);
    if (!diapositivas.length) throw new Error('PptxDesdeContrato: la presentación necesita al menos una diapositiva');
    const lang = 'en-US';
    const ctx = new Contexto(contrato);
    const fuenteBase = typeof contrato.fuente === 'string' && contrato.fuente ? contrato.fuente : 'Arial';
    const maestros = (contrato.maestros && contrato.maestros.length ? contrato.maestros : [{ formas: [] }]).slice(0, 256);
    const hayNotas = diapositivas.some((d) => d && Array.isArray(d.notas) && d.notas.length);

    const zip = new ZipPropio();
    const rels = (lista) => `${CABECERA}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${lista.map(([id, tipo, destino]) => `<Relationship Id="${id}" Type="${tipo}" Target="${destino}"/>`).join('')}</Relationships>`;
    const tipos = [];
    const sobre = (parte, ct) => tipos.push(`<Override PartName="${parte}" ContentType="${ct}"/>`);

    // Cada patrón: su tema, su único diseño en blanco
    const relsPres = [];
    maestros.forEach((m, i) => {
      const n = i + 1;
      const r = new Rels(ctx, [['rId1', `${REL}/slideLayout`, `../slideLayouts/slideLayout${n}.xml`], ['rId2', `${REL}/theme`, `../theme/theme${n}.xml`]]);
      zip.file(`ppt/slideMasters/slideMaster${n}.xml`, patron(m || {}, ctx, r, lang).replace('__LAYOUT_ID__', String(2147483648 + n)));
      zip.file(`ppt/slideMasters/_rels/slideMaster${n}.xml.rels`, r.xml());
      zip.file(`ppt/slideLayouts/slideLayout${n}.xml`, diseno());
      zip.file(`ppt/slideLayouts/_rels/slideLayout${n}.xml.rels`, rels([['rId1', `${REL}/slideMaster`, `../slideMasters/slideMaster${n}.xml`]]));
      zip.file(`ppt/theme/theme${n}.xml`, EscritorPptx.tema('Office Theme', COLORES_TEMA, fuentesTema(fuenteBase)));
      sobre(`/ppt/slideMasters/slideMaster${n}.xml`, `${CT}.slideMaster+xml`);
      sobre(`/ppt/slideLayouts/slideLayout${n}.xml`, `${CT}.slideLayout+xml`);
      sobre(`/ppt/theme/theme${n}.xml`, 'application/vnd.openxmlformats-officedocument.theme+xml');
      relsPres.push([`rId${n}`, `${REL}/slideMaster`, `slideMasters/slideMaster${n}.xml`]);
    });
    const nM = maestros.length;
    let relNotasMaestro = null;
    if (hayNotas) {
      zip.file('ppt/notesMasters/notesMaster1.xml', patronNotas());
      zip.file('ppt/notesMasters/_rels/notesMaster1.xml.rels', rels([['rId1', `${REL}/theme`, `../theme/theme${nM + 1}.xml`]]));
      zip.file(`ppt/theme/theme${nM + 1}.xml`, EscritorPptx.tema('Office Theme', COLORES_TEMA, fuentesTema(fuenteBase)));
      sobre('/ppt/notesMasters/notesMaster1.xml', `${CT}.notesMaster+xml`);
      sobre(`/ppt/theme/theme${nM + 1}.xml`, 'application/vnd.openxmlformats-officedocument.theme+xml');
      relNotasMaestro = [`rId${nM + 3}`, `${REL}/notesMaster`, 'notesMasters/notesMaster1.xml'];
    }

    const idsDiap = [];
    diapositivas.forEach((d, i) => {
      d = d && typeof d === 'object' ? d : {};
      const n = i + 1;
      const m = Number.isInteger(d.maestro) && d.maestro >= 0 && d.maestro < nM ? d.maestro + 1 : 1;
      const r = new Rels(ctx, [['rId1', `${REL}/slideLayout`, `../slideLayouts/slideLayout${m}.xml`]]);
      const atrib = (d.oculta ? ' show="0"' : '') + (d.ocultarPatron ? ' showMasterSp="0"' : '');
      const xml = `${CABECERA}<p:sld ${NS}${atrib}><p:cSld>${fondo(d.fondo, r)}<p:spTree>${arbol(d.formas, ctx, r, lang)}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
      if (hayNotas && Array.isArray(d.notas) && d.notas.length) {
        const nr = new Rels(ctx, [['rId1', `${REL}/notesMaster`, '../notesMasters/notesMaster1.xml'], ['rId2', `${REL}/slide`, `../slides/slide${n}.xml`]]);
        const cuerpoNotas = cuerpo({ parrafos: d.notas }, nr, ctx, lang, 'p:txBody', true);
        zip.file(`ppt/notesSlides/notesSlide${n}.xml`, `${CABECERA}<p:notes ${NS}><p:cSld><p:spTree>${GRUPO}`
          + '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>'
          + `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>${cuerpoNotas}</p:sp>`
          + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>');
        zip.file(`ppt/notesSlides/_rels/notesSlide${n}.xml.rels`, nr.xml());
        sobre(`/ppt/notesSlides/notesSlide${n}.xml`, `${CT}.notesSlide+xml`);
        r.lista.push([`rId${r.lista.length + 1 + r.base.length}`, `${REL}/notesSlide`, `../notesSlides/notesSlide${n}.xml`]);
      }
      zip.file(`ppt/slides/slide${n}.xml`, xml);
      zip.file(`ppt/slides/_rels/slide${n}.xml.rels`, r.xml());
      sobre(`/ppt/slides/slide${n}.xml`, `${CT}.slide+xml`);
      idsDiap.push(`<p:sldId id="${256 + i}" r:id="rId${nM + 10 + i}"/>`);
      relsPres.push([`rId${nM + 10 + i}`, `${REL}/slide`, `slides/slide${n}.xml`]);
    });

    relsPres.push([`rId${nM + 1}`, `${REL}/presProps`, 'presProps.xml'], [`rId${nM + 2}`, `${REL}/viewProps`, 'viewProps.xml']);
    if (relNotasMaestro) relsPres.push(relNotasMaestro);
    relsPres.push([`rId${nM + 4}`, `${REL}/tableStyles`, 'tableStyles.xml'], [`rId${nM + 5}`, `${REL}/theme`, 'theme/theme1.xml']);
    const tamBase = typeof contrato.tamano === 'number' && contrato.tamano > 0 ? entero(contrato.tamano * 100, 100, 400000) : 1800;
    const nivelesPred = Array.from({ length: 9 }, (_, i) => `<a:lvl${i + 1}pPr marL="${457200 * i}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="${tamBase}" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i + 1}pPr>`).join('');
    const presentacion = `${CABECERA}<p:presentation ${NS} saveSubsetFonts="1"><p:sldMasterIdLst>${maestros.map((_, i) => `<p:sldMasterId id="${2147484000 + i}" r:id="rId${i + 1}"/>`).join('')}</p:sldMasterIdLst>`
      + (relNotasMaestro ? `<p:notesMasterIdLst><p:notesMasterId r:id="${relNotasMaestro[0]}"/></p:notesMasterIdLst>` : '')
      + `<p:sldIdLst>${idsDiap.join('')}</p:sldIdLst><p:sldSz cx="${ancho}" cy="${alto}"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="${lang}"/></a:defPPr>${nivelesPred}</p:defaultTextStyle></p:presentation>`;
    // los id de diseños (2147483649…) y de patrones (2147484000…) comparten espacio y no se cruzan con hasta 256 patrones
    const prop = contrato.propiedades || {};
    const ahora = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    const core = `${CABECERA}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`
      + (prop.titulo ? `<dc:title>${esc(prop.titulo)}</dc:title>` : '') + (prop.autor ? `<dc:creator>${esc(prop.autor)}</dc:creator>` : '')
      + `<dcterms:created xsi:type="dcterms:W3CDTF">${ahora}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${ahora}</dcterms:modified></cp:coreProperties>`;

    const exts = new Set(ctx.medios.map((m) => m.tipo));
    const tiposXml = `${CABECERA}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`
      + [...exts].map((e) => `<Default Extension="${e}" ContentType="${MIME[e]}"/>`).join('')
      + `<Override PartName="/ppt/presentation.xml" ContentType="${CT}.presentation.main+xml"/>${tipos.join('')}`
      + `<Override PartName="/ppt/presProps.xml" ContentType="${CT}.presProps+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="${CT}.viewProps+xml"/><Override PartName="/ppt/tableStyles.xml" ContentType="${CT}.tableStyles+xml"/>`
      + '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>';
    zip.file('[Content_Types].xml', tiposXml);
    zip.file('_rels/.rels', rels([['rId1', `${REL}/officeDocument`, 'ppt/presentation.xml'], ['rId2', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml']]));
    zip.file('ppt/presentation.xml', presentacion);
    zip.file('ppt/_rels/presentation.xml.rels', rels(relsPres));
    zip.file('ppt/presProps.xml', `${CABECERA}<p:presentationPr ${NS}/>`);
    zip.file('ppt/viewProps.xml', `${CABECERA}<p:viewPr ${NS}><p:gridSpacing cx="72008" cy="72008"/></p:viewPr>`);
    zip.file('ppt/tableStyles.xml', `${CABECERA}<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`);
    for (const m of ctx.medios) zip.file(`ppt/media/${m.nombre}`, m.bytes);
    zip.file('docProps/core.xml', core);
    return { bytes: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), mensajes: [...ctx.mensajes] };
  }

  globalThis.PptxDesdeContrato = { generar };
})();

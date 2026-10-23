/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Modelo resuelto de un .pptx: lee el paquete y devuelve, por diapositiva, la
// lista de lo que hay que dibujar con TODA la herencia ya aplicada (tema →
// patrón → diseño → marcador → forma), los colores con sus modificadores ya
// calculados, y el relleno, la línea y la sombra que vienen del estilo del
// tema. Quien dibuja (pptx-dibujo.js) no vuelve a mirar el XML.
//
// Es la etapa 2 del motor propio (investigacion-motor-powerpoint\00-resumen.md,
// fuera del repo). La herencia sigue lo medido con comparar-pptx; donde la
// norma y PowerPoint no coinciden, manda PowerPoint.
//
// Requiere ZipPropio y pptx-geometria.js (y pptx-estilos-tabla.js para las tablas). Uso:
//   const pres = await PptxModelo.abrir(archivo);
//   pres.ancho, pres.alto (EMU), pres.cantidad
//   const d = await pres.diapositiva(i);  // { fondo, elementos, soportada, motivos }; motivos: lo que se omitió o se dibujó sin su formato
//   const img = await pres.imagen(ruta);  // { fuente, ancho, alto, dpiX, dpiY } | null

const PptxModelo = (() => {
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const URI_TABLA = 'http://schemas.openxmlformats.org/drawingml/2006/table';
  const URI_GRAFICO = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  const URI_DIAGRAMA = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
  const URI_OLE = 'http://schemas.openxmlformats.org/presentationml/2006/ole';
  const TIPO_DIBUJO_DIAGRAMA = 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing';
  const RELLENOS = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];
  const COLORES = ['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'hslClr', 'scrgbClr'];

  // --- XML ---------------------------------------------------------------------
  const { hijo, hijos, ruta, todos, parsear } = PptxXml;
  const primero = (el, nombres) => { if (!el) return null; for (const h of el.children) if (nombres.includes(h.localName)) return h; return null; };
  const ent = (el, a, def = null) => { if (!el || !el.hasAttribute(a)) return def; const n = parseInt(el.getAttribute(a), 10); return Number.isFinite(n) ? n : def; };
  const bool = (el, a, def = false) => { if (!el || !el.hasAttribute(a)) return def; const v = el.getAttribute(a); return v === '1' || v === 'true'; };

  function resolverRuta(desde, destino) {
    try { destino = decodeURIComponent(destino); } catch { /* tal cual */ }
    if (destino.startsWith('/')) return destino.slice(1);
    const partes = desde.split('/').slice(0, -1);
    for (const p of destino.split('/')) {
      if (p === '..') partes.pop();
      else if (p !== '.' && p !== '') partes.push(p);
    }
    return partes.join('/');
  }
  const rutaRels = (parte) => parte.replace(/([^/]+)$/, '_rels/$1.rels');

  // --- Colores -------------------------------------------------------------------
  // tint y shade en RGB lineal, lumMod/lumOff/satMod en HSL: así coincide con
  // los píxeles de PowerPoint (medido con comparar-pptx).
  const aLineal = (c) => { c /= 255; return c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4); };
  const aSrgb = (c) => { const v = c <= 0.0031308 ? c * 12.92 : 1.055 * Math.pow(c, 1 / 2.4) - 0.055; return Math.max(0, Math.min(255, Math.round(v * 255))); };
  function aHsl([r, g, b]) {
    r /= 255; g /= 255; b /= 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b), l = (max + min) / 2;
    if (max === min) return [0, 0, l];
    const d = max - min, s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    return [h * 60, s, l];
  }
  function deHsl([h, s, l]) {
    if (s === 0) return [l, l, l].map((v) => Math.round(v * 255));
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    const canal = (t) => { t = ((t % 1) + 1) % 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
    return [canal(h / 360 + 1 / 3), canal(h / 360), canal(h / 360 - 1 / 3)].map((v) => Math.round(v * 255));
  }
  const PRESETS = {
    black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', cyan: '00FFFF',
    magenta: 'FF00FF', gray: '808080', grey: '808080', silver: 'C0C0C0', maroon: '800000', navy: '000080', olive: '808000',
    purple: '800080', teal: '008080', lime: '00FF00', orange: 'FFA500', pink: 'FFC0CB', brown: 'A52A2A', gold: 'FFD700',
    dkBlue: '00008B', dkRed: '8B0000', dkGreen: '006400', dkGray: 'A9A9A9', ltGray: 'D3D3D3', ltBlue: 'ADD8E6',
    ltGreen: '90EE90', ltYellow: 'FFFFE0', medBlue: '0000CD', dkCyan: '008B8B', dkMagenta: '8B008B', dkOrange: 'FF8C00',
  };
  const deHex = (h) => (h && /^[0-9a-f]{6}$/i.test(h) ? [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16)) : null);

  // el: un elemento de color (srgbClr, schemeClr…) o uno que lo contiene
  // (solidFill, buClr…). cc: { tema(nombre) → [r,g,b], ph: {r,g,b,a} }.
  // Devuelve { r, g, b, a } (a de 0 a 1) o null.
  function color(el, cc) {
    if (!el) return null;
    const c = COLORES.includes(el.localName) ? el : primero(el, COLORES);
    if (!c) return null;
    let rgb = null, a = 1;
    switch (c.localName) {
      case 'srgbClr': rgb = deHex(c.getAttribute('val')); break;
      case 'sysClr': rgb = deHex(c.getAttribute('lastClr')) || (c.getAttribute('val') === 'window' ? [255, 255, 255] : [0, 0, 0]); break;
      case 'prstClr': rgb = deHex(PRESETS[c.getAttribute('val')]) || [0, 0, 0]; break;
      case 'scrgbClr': rgb = ['r', 'g', 'b'].map((k) => aSrgb((ent(c, k, 0) / 100000))); break;
      case 'hslClr': rgb = deHsl([ent(c, 'hue', 0) / 60000, ent(c, 'sat', 0) / 100000, ent(c, 'lum', 0) / 100000]); break;
      case 'schemeClr': {
        const v = c.getAttribute('val');
        if (v === 'phClr') { if (!cc.ph) return null; rgb = [cc.ph.r, cc.ph.g, cc.ph.b]; a = cc.ph.a; } else rgb = cc.tema(v);
        break;
      }
    }
    if (!rgb) return null;
    for (const m of c.children) {
      const v = ent(m, 'val', 0) / 100000;
      switch (m.localName) {
        case 'tint': rgb = rgb.map((x) => aSrgb(aLineal(x) * v + (1 - v))); break;
        case 'shade': rgb = rgb.map((x) => aSrgb(aLineal(x) * v)); break;
        case 'lumMod': { const h = aHsl(rgb); h[2] = Math.min(1, h[2] * v); rgb = deHsl(h); break; }
        case 'lumOff': { const h = aHsl(rgb); h[2] = Math.max(0, Math.min(1, h[2] + v)); rgb = deHsl(h); break; }
        case 'lum': { const h = aHsl(rgb); h[2] = Math.max(0, Math.min(1, v)); rgb = deHsl(h); break; }
        case 'satMod': { const h = aHsl(rgb); h[1] = Math.max(0, Math.min(1, h[1] * v)); rgb = deHsl(h); break; }
        case 'satOff': { const h = aHsl(rgb); h[1] = Math.max(0, Math.min(1, h[1] + v)); rgb = deHsl(h); break; }
        case 'sat': { const h = aHsl(rgb); h[1] = Math.max(0, Math.min(1, v)); rgb = deHsl(h); break; }
        case 'hueMod': { const h = aHsl(rgb); h[0] = (h[0] * v) % 360; rgb = deHsl(h); break; }
        case 'hueOff': { const h = aHsl(rgb); h[0] = ((h[0] + ent(m, 'val', 0) / 60000) % 360 + 360) % 360; rgb = deHsl(h); break; }
        case 'hue': { const h = aHsl(rgb); h[0] = (ent(m, 'val', 0) / 60000) % 360; rgb = deHsl(h); break; }
        case 'comp': { const h = aHsl(rgb); h[0] = (h[0] + 180) % 360; rgb = deHsl(h); break; }
        case 'inv': rgb = rgb.map((x) => 255 - x); break;
        case 'gray': { const y = Math.round(0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]); rgb = [y, y, y]; break; }
        case 'red': rgb[0] = aSrgb(v); break;
        case 'green': rgb[1] = aSrgb(v); break;
        case 'blue': rgb[2] = aSrgb(v); break;
        case 'redMod': rgb[0] = aSrgb(aLineal(rgb[0]) * v); break;
        case 'greenMod': rgb[1] = aSrgb(aLineal(rgb[1]) * v); break;
        case 'blueMod': rgb[2] = aSrgb(aLineal(rgb[2]) * v); break;
        case 'redOff': rgb[0] = aSrgb(aLineal(rgb[0]) + v); break;
        case 'greenOff': rgb[1] = aSrgb(aLineal(rgb[1]) + v); break;
        case 'blueOff': rgb[2] = aSrgb(aLineal(rgb[2]) + v); break;
        case 'gamma': rgb = rgb.map((x) => aSrgb(Math.pow(x / 255, 2.2))); break;
        case 'invGamma': rgb = rgb.map((x) => Math.round(255 * aLineal(x))); break;
        case 'alpha': a = v; break;
        case 'alphaMod': a *= v; break;
        case 'alphaOff': a += v; break;
      }
    }
    return { r: rgb[0], g: rgb[1], b: rgb[2], a: Math.max(0, Math.min(1, a)) };
  }

  // --- Rellenos ------------------------------------------------------------------
  // Devuelve null (sin relleno) o { tipo: 'solido'|'degradado'|'trama'|'imagen'|'grupo'|'fondo', … }.
  // parte: { ruta, rels } de donde salió el relleno (para las imágenes).
  function relleno(el, cc, parte) {
    if (!el) return null;
    switch (el.localName) {
      case 'noFill': return null;
      case 'solidFill': { const c = color(el, cc); return c ? { tipo: 'solido', color: c } : null; }
      case 'grpFill': return { tipo: 'grupo' };
      case 'gradFill': {
        const paradas = hijos(hijo(el, 'gsLst'), 'gs').map((gs) => ({ pos: ent(gs, 'pos', 0) / 100000, color: color(gs, cc) }))
          .filter((p) => p.color).sort((a, b) => a.pos - b.pos);
        if (!paradas.length) return null;
        const lin = hijo(el, 'lin'), path = hijo(el, 'path');
        const r = { tipo: 'degradado', paradas, rotarConForma: bool(el, 'rotWithShape', true) };
        if (path) {
          const f = hijo(path, 'fillToRect');
          r.ruta = path.getAttribute('path') || 'circle';
          r.foco = { l: ent(f, 'l', 0) / 100000, t: ent(f, 't', 0) / 100000, r: ent(f, 'r', 0) / 100000, b: ent(f, 'b', 0) / 100000 };
        } else {
          r.angulo = lin ? ent(lin, 'ang', 0) / 60000 : 0;
          r.escalado = lin ? bool(lin, 'scaled', false) : false;
        }
        return r;
      }
      case 'pattFill': {
        const fg = color(hijo(el, 'fgClr'), cc) || { r: 0, g: 0, b: 0, a: 1 };
        const bg = color(hijo(el, 'bgClr'), cc) || { r: 255, g: 255, b: 255, a: 1 };
        return { tipo: 'trama', prst: el.getAttribute('prst') || 'pct5', fg, bg };
      }
      case 'blipFill': return imagenDeBlip(el, cc, parte);
    }
    return null;
  }

  function imagenDeBlip(blipFill, cc, parte) {
    const blip = hijo(blipFill, 'blip');
    const id = blip && (blip.getAttributeNS(NS_R, 'embed') || blip.getAttributeNS(NS_R, 'link'));
    const rel = id && parte.rels.porId(id);
    if (!rel || !rel.destino) return null;
    const src = hijo(blipFill, 'srcRect');
    const recorte = src ? { l: ent(src, 'l', 0) / 100000, t: ent(src, 't', 0) / 100000, r: ent(src, 'r', 0) / 100000, b: ent(src, 'b', 0) / 100000 } : null;
    const r = { tipo: 'imagen', ruta: rel.destino, recorte, efectos: [], alfa: 1, rotarConForma: bool(blipFill, 'rotWithShape', true) };
    for (const e of blip.children) {
      switch (e.localName) {
        case 'alphaModFix': r.alfa = ent(e, 'amt', 100000) / 100000; break;
        case 'grayscl': r.efectos.push({ tipo: 'gris' }); break;
        case 'biLevel': r.efectos.push({ tipo: 'bitono', umbral: ent(e, 'thresh', 50000) / 100000 }); break;
        case 'lum': r.efectos.push({ tipo: 'luz', brillo: ent(e, 'bright', 0) / 100000, contraste: ent(e, 'contrast', 0) / 100000 }); break;
        case 'duotone': { const cs = [...e.children].map((c) => color(c, cc)).filter(Boolean); if (cs.length === 2) r.efectos.push({ tipo: 'duotono', colores: cs }); break; }
        case 'clrChange': { const de = color(hijo(e, 'clrFrom'), cc), a = color(hijo(e, 'clrTo'), cc); if (de && a) r.efectos.push({ tipo: 'cambioColor', de, a }); break; }
      }
    }
    const tile = hijo(blipFill, 'tile');
    if (tile) {
      r.mosaico = { tx: ent(tile, 'tx', 0), ty: ent(tile, 'ty', 0), sx: ent(tile, 'sx', 100000) / 100000, sy: ent(tile, 'sy', 100000) / 100000,
        flip: tile.getAttribute('flip') || 'none', algn: tile.getAttribute('algn') || 'tl' };
    } else {
      const fr = ruta(blipFill, 'stretch', 'fillRect');
      r.estirar = { l: ent(fr, 'l', 0) / 100000, t: ent(fr, 't', 0) / 100000, r: ent(fr, 'r', 0) / 100000, b: ent(fr, 'b', 0) / 100000 };
    }
    return r;
  }

  // --- Paquete --------------------------------------------------------------------
  // Las partes XML de un paquete OOXML (con sus relaciones) y sus imágenes, ya
  // leídas una sola vez. Sirve igual para un .pptx que para un .xlsx.
  function paquete(zip) {
    const cache = new Map();
    async function leerRels(parte) {
      const a = zip.file(rutaRels(parte));
      const doc = a ? parsear(await a.async('string')) : null;
      const lista = doc ? todos(doc, 'Relationship').map((r) => ({
        id: r.getAttribute('Id'), tipo: r.getAttribute('Type') || '',
        destino: r.getAttribute('TargetMode') === 'External' || !r.getAttribute('Target') ? null : resolverRuta(parte, r.getAttribute('Target')),
      })) : [];
      return { lista, porId: (id) => lista.find((r) => r.id === id), porTipo: (t) => lista.filter((r) => r.destino && r.tipo.endsWith('/' + t)) };
    }
    async function parte(rutaParte) {
      if (!cache.has(rutaParte)) {
        cache.set(rutaParte, (async () => {
          const a = zip.file(rutaParte);
          if (!a) return null;
          const doc = parsear(await a.async('string'));
          if (todos(doc, 'parsererror').length) return null;
          return { ruta: rutaParte, doc, raiz: doc.documentElement, rels: await leerRels(rutaParte) };
        })());
      }
      return cache.get(rutaParte);
    }
    const imagenes = new Map();
    function imagen(rutaImagen) {
      if (!imagenes.has(rutaImagen)) imagenes.set(rutaImagen, cargarImagen(zip, rutaImagen).catch(() => null));
      return imagenes.get(rutaImagen);
    }
    return { parte, imagen };
  }

  async function abrir(archivo) {
    // un .ppt o un .odp se pasa a .pptx antes (si formatos-powerpoint.js está en la página); un .pptx no se toca
    if (typeof FormatosPowerpoint !== 'undefined') archivo = await FormatosPowerpoint.comoPptx(archivo);
    const zip = await ZipPropio.loadAsync(archivo);
    const { parte, imagen } = paquete(zip);

    const pres = await parte('ppt/presentation.xml');
    if (!pres) throw new Error('No se pudo leer ppt/presentation.xml');
    const sldSz = hijo(pres.raiz, 'sldSz');
    const ancho = ent(sldSz, 'cx', 0), alto = ent(sldSz, 'cy', 0);
    if (!ancho || !alto) throw new Error('Tamaño de diapositiva inválido.');
    const diapositivas = hijos(hijo(pres.raiz, 'sldIdLst'), 'sldId')
      .map((s) => pres.rels.porId(s.getAttributeNS(NS_R, 'id'))).filter((r) => r && r.destino).map((r) => r.destino);
    const textoPorDefecto = hijo(pres.raiz, 'defaultTextStyle');
    // Estilos de tabla que trae el archivo (PowerPoint guarda solo los usados,
    // y a veces ni esos: entonces la tabla queda sin estilo propio, ver tabla()).
    const estilosTabla = new Map();
    const relEstilos = pres.rels.porTipo('tableStyles')[0];
    const partEstilos = relEstilos && await parte(relEstilos.destino);
    for (const e of partEstilos ? hijos(partEstilos.raiz, 'tblStyle') : []) estilosTabla.set((e.getAttribute('styleId') || '').toUpperCase(), e);

    async function diapositiva(i) {
      const sld = await parte(diapositivas[i]);
      if (!sld) throw new Error('Diapositiva ilegible: ' + diapositivas[i]);
      const relLayout = sld.rels.porTipo('slideLayout')[0];
      const layout = relLayout ? await parte(relLayout.destino) : null;
      const relMaster = layout && layout.rels.porTipo('slideMaster')[0];
      const master = relMaster ? await parte(relMaster.destino) : null;
      const relTema = master && master.rels.porTipo('theme')[0];
      const tema = relTema ? await parte(relTema.destino) : null;
      const ctx = crearContexto({ sld, layout, master, tema, textoPorDefecto, numero: i + 1, parte });
      ctx.estilosTabla = estilosTabla;
      const salida = { numero: i + 1, oculta: sld.raiz.getAttribute('show') === '0', elementos: [], motivos: [] };
      salida.fondo = fondo(ctx);
      const mostrarPatron = sld.raiz.getAttribute('showMasterSp') !== '0' && sld.raiz.getAttribute('showMasterSp') !== 'false';
      const capas = [];
      if (mostrarPatron && layout) {
        const layoutMuestra = layout.raiz.getAttribute('showMasterSp') !== '0' && layout.raiz.getAttribute('showMasterSp') !== 'false';
        if (master && layoutMuestra) capas.push({ p: master, cadena: [] });
        capas.push({ p: layout, cadena: master ? [master] : [] });
      }
      capas.push({ p: sld, cadena: [layout, master].filter(Boolean) });
      for (const capa of capas) {
        const arbol = ruta(capa.p.raiz, 'cSld', 'spTree');
        const esDiapositiva = capa.p === sld;
        salida.elementos.push(...await leerArbol(arbol, { ...ctx, parteActual: capa.p, cadena: capa.cadena, esDiapositiva }, salida.motivos));
      }
      salida.soportada = salida.motivos.length === 0;
      return salida;
    }

    return { ancho, alto, cantidad: diapositivas.length, diapositiva, imagen };
  }

  // --- Contexto de una diapositiva: tema, mapa de colores, fuentes -------------------
  function crearContexto({ sld, layout, master, tema, textoPorDefecto, numero, parte }) {
    const esquema = {};
    const clrScheme = tema && todos(tema.doc, 'clrScheme')[0];
    for (const el of clrScheme ? [...clrScheme.children] : []) {
      const c = el.firstElementChild;
      if (!c) continue;
      esquema[el.localName] = deHex(c.localName === 'sysClr' ? c.getAttribute('lastClr') : c.getAttribute('val'));
    }
    const mapa = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
    const clrMap = master && hijo(master.raiz, 'clrMap');
    for (const at of clrMap ? [...clrMap.attributes] : []) mapa[at.name] = at.value;
    for (const p of [layout, sld]) {
      const ovr = p && ruta(p.raiz, 'clrMapOvr', 'overrideClrMapping');
      for (const at of ovr ? [...ovr.attributes] : []) mapa[at.name] = at.value;
    }
    const temaColor = (nombre) => esquema[mapa[nombre] || nombre] || esquema[nombre] || [0, 0, 0];

    const fuentes = {};
    for (const [clave, nombre] of [['mj', 'majorFont'], ['mn', 'minorFont']]) {
      const f = tema && todos(tema.doc, nombre)[0];
      if (!f) continue;
      const latin = (hijo(f, 'latin') || { getAttribute: () => '' }).getAttribute('typeface') || '';
      for (const g of ['lt', 'ea', 'cs']) {
        const el = hijo(f, g === 'lt' ? 'latin' : g);
        fuentes[`+${clave}-${g}`] = (el && el.getAttribute('typeface')) || latin;
      }
    }
    const listaTema = (nombre) => { const l = tema && todos(tema.doc, nombre)[0]; return l ? [...l.children] : []; };
    return {
      sld, layout, master, tema, textoPorDefecto, numero, parte,
      temaColor, cc: { tema: temaColor, ph: null }, fuentes,
      estilosTema: { relleno: listaTema('fillStyleLst'), fondo: listaTema('bgFillStyleLst'), linea: listaTema('lnStyleLst'), efecto: listaTema('effectStyleLst') },
      txStyles: master && hijo(master.raiz, 'txStyles'),
    };
  }

  // El relleno de una referencia del tema (fillRef/bgRef idx) con su color.
  function rellenoDeReferencia(ref, ctx) {
    if (!ref) return undefined;
    const idx = ent(ref, 'idx', 0);
    if (!idx) return null;
    const plantilla = idx >= 1001 ? ctx.estilosTema.fondo[idx - 1001] : ctx.estilosTema.relleno[idx - 1];
    if (!plantilla) return null;
    const cc = { tema: ctx.temaColor, ph: color(ref, ctx.cc) };
    return relleno(plantilla, cc, ctx.tema);
  }

  function fondo(ctx) {
    for (const p of [ctx.sld, ctx.layout, ctx.master]) {
      const bg = p && ruta(p.raiz, 'cSld', 'bg');
      if (!bg) continue;
      const bgPr = hijo(bg, 'bgPr');
      if (bgPr) return relleno(primero(bgPr, RELLENOS), ctx.cc, p);
      const bgRef = hijo(bg, 'bgRef');
      if (bgRef) return rellenoDeReferencia(bgRef, ctx) || null;
    }
    return { tipo: 'solido', color: { r: 255, g: 255, b: 255, a: 1 } };
  }

  // --- Marcadores de posición -----------------------------------------------------------
  function datosMarcador(el) {
    const nvPr = primero(el, ['nvSpPr', 'nvPicPr', 'nvGraphicFramePr', 'nvCxnSpPr', 'nvGrpSpPr']);
    const ph = ruta(nvPr, 'nvPr', 'ph');
    if (!ph) return null;
    return { tipo: ph.getAttribute('type') || 'obj', idx: ph.getAttribute('idx') };
  }
  const tipoBase = (t) => (t === 'ctrTitle' ? 'title' : ['subTitle', 'obj', 'body'].includes(t) ? 'body' : t);
  function buscarMarcador(p, marcador, esPatron) {
    const arbol = ruta(p.raiz, 'cSld', 'spTree');
    const formas = todos(arbol, 'sp').filter((s) => datosMarcador(s));
    if (!esPatron && marcador.idx != null) {
      const porIdx = formas.find((s) => datosMarcador(s).idx === marcador.idx);
      if (porIdx) return porIdx;
    }
    const exacto = formas.find((s) => datosMarcador(s).tipo === marcador.tipo);
    return exacto || formas.find((s) => tipoBase(datosMarcador(s).tipo) === tipoBase(marcador.tipo)) || null;
  }
  // De la más cercana a la más lejana: el marcador del diseño y el del patrón.
  function herenciaDe(el, ctx) {
    const marcador = datosMarcador(el);
    if (!marcador) return [];
    return ctx.cadena.map((p) => buscarMarcador(p, marcador, p === ctx.master)).filter(Boolean);
  }

  function leerXfrm(x) {
    if (!x) return null;
    const off = hijo(x, 'off'), ext = hijo(x, 'ext');
    if (!off || !ext) return null;
    const r = {
      x: ent(off, 'x', 0), y: ent(off, 'y', 0), w: ent(ext, 'cx', 0), h: ent(ext, 'cy', 0),
      rot: ent(x, 'rot', 0) / 60000, flipH: bool(x, 'flipH'), flipV: bool(x, 'flipV'),
    };
    const chOff = hijo(x, 'chOff'), chExt = hijo(x, 'chExt');
    if (chOff && chExt) r.hijos = { x: ent(chOff, 'x', 0), y: ent(chOff, 'y', 0), w: ent(chExt, 'cx', 0), h: ent(chExt, 'cy', 0) };
    return r;
  }

  // --- Árbol de formas ------------------------------------------------------------------------
  async function leerArbol(arbol, ctx, motivos) {
    const salida = [];
    for (let el of arbol ? [...arbol.children] : []) {
      if (el.localName === 'AlternateContent') {
        // El respaldo (Fallback) es lo que dibuja una versión que no conoce la
        // novedad: casi siempre una imagen (ecuaciones, tinta) o una forma común.
        const opcion = hijo(el, 'Fallback') || hijo(el, 'Choice');
        for (const h of opcion ? [...opcion.children] : []) {
          const sub = await leerArbol({ children: [h] }, ctx, motivos);
          salida.push(...sub);
        }
        continue;
      }
      const cNvPr = ruta(primero(el, ['nvSpPr', 'nvPicPr', 'nvGraphicFramePr', 'nvCxnSpPr', 'nvGrpSpPr']), 'cNvPr');
      if (bool(cNvPr, 'hidden')) continue;
      // En diseños y patrones, los marcadores solo dan formato: no se dibujan.
      if (!ctx.esDiapositiva && datosMarcador(el)) continue;
      const nombre = cNvPr ? cNvPr.getAttribute('name') || '' : '';
      let elemento = null;
      switch (el.localName) {
        case 'sp': case 'cxnSp': elemento = forma(el, ctx); break;
        case 'pic': elemento = imagenForma(el, ctx); break;
        case 'grpSp': elemento = await grupo(el, ctx, motivos); break;
        case 'graphicFrame': elemento = await marco(el, ctx, motivos); break;
      }
      if (elemento) { elemento.nombre = nombre; salida.push(elemento); }
    }
    return salida;
  }

  async function grupo(el, ctx, motivos) {
    const grpSpPr = hijo(el, 'grpSpPr');
    const xfrm = leerXfrm(hijo(grpSpPr, 'xfrm'));
    if (!xfrm) return null;
    const hijosGrupo = await leerArbol(el, ctx, motivos);
    return { tipo: 'grupo', xfrm, hijos: hijosGrupo.filter((h) => !['nvGrpSpPr', 'grpSpPr'].includes(h.tipo)),
      relleno: relleno(primero(grpSpPr, RELLENOS), ctx.cc, ctx.parteActual) };
  }

  // Lo común de una forma o una imagen: posición, geometría, relleno, línea, sombra.
  function propiedadesForma(el, ctx, herencia) {
    const spPr = hijo(el, 'spPr');
    const spPrs = [spPr, ...herencia.map((h) => hijo(h, 'spPr'))];
    let xfrm = null;
    for (const s of spPrs) { xfrm = leerXfrm(hijo(s, 'xfrm')); if (xfrm) break; }
    let geometria = { def: PptxGeometria.predefinida('rect'), ajustes: {} };
    for (const s of spPrs) {
      const prst = hijo(s, 'prstGeom'), cust = hijo(s, 'custGeom');
      if (prst) { geometria = { def: PptxGeometria.predefinida(prst.getAttribute('prst')) || PptxGeometria.predefinida('rect'), ajustes: PptxGeometria.ajustesDeXml(prst), prst: prst.getAttribute('prst') }; break; }
      if (cust) { geometria = { def: PptxGeometria.desdeXml(cust), ajustes: {} }; break; }
    }
    const estilo = hijo(el, 'style');
    const estilos = [estilo, ...herencia.map((h) => hijo(h, 'style'))];

    // Relleno: el propio → el heredado → el del estilo de la forma.
    let rell;
    const usaFondo = bool(el, 'useBgFill');
    if (usaFondo) rell = { tipo: 'fondo' };
    else {
      let definido = false;
      for (let i = 0; i < spPrs.length && !definido; i++) {
        const r = primero(spPrs[i], RELLENOS);
        if (r) { rell = relleno(r, ctx.cc, i === 0 ? ctx.parteActual : ctx.cadena[i - 1] || ctx.parteActual); definido = true; }
      }
      if (!definido) {
        const ref = estilos.map((s) => hijo(s, 'fillRef')).find(Boolean);
        rell = ref ? rellenoDeReferencia(ref, ctx) : null;
      }
    }
    return { xfrm, geometria, relleno: rell, linea: linea(spPrs, estilos, ctx), sombra: sombra(spPrs, estilos, ctx) };
  }

  // nombreLn: el elemento de la línea dentro de cada spPr (lnL, lnT… en una celda).
  function linea(spPrs, estilos, ctx, nombreLn = 'ln') {
    const lns = spPrs.map((s) => hijo(s, nombreLn)).filter(Boolean);
    const ref = estilos.map((s) => hijo(s, 'lnRef')).find(Boolean);
    const idx = ref ? ent(ref, 'idx', 0) : 0;
    const delTema = idx > 0 ? ctx.estilosTema.linea[idx - 1] : null;
    if (!lns.length && !delTema) return null;
    const cadena = delTema ? [...lns, delTema] : lns;
    const cc = { tema: ctx.temaColor, ph: ref ? color(ref, ctx.cc) : null };
    let rell, hayRelleno = false;
    for (const ln of cadena) {
      const r = primero(ln, ['noFill', 'solidFill', 'gradFill', 'pattFill']);
      if (r) { rell = relleno(r, ln === delTema ? cc : ctx.cc, ctx.parteActual); hayRelleno = true; break; }
    }
    if (!hayRelleno || !rell) return null;
    const at = (a) => { for (const ln of cadena) if (ln.hasAttribute(a)) return ln.getAttribute(a); return null; };
    const sub = (nombres) => { for (const ln of cadena) { const e = primero(ln, nombres); if (e) return e; } return null; };
    const guion = sub(['prstDash', 'custDash']);
    const union = sub(['round', 'bevel', 'miter']);
    const punta = (n) => { const e = sub([n]); return e && e.getAttribute('type') && e.getAttribute('type') !== 'none' ? { tipo: e.getAttribute('type'), w: e.getAttribute('w') || 'med', len: e.getAttribute('len') || 'med' } : null; };
    return {
      ancho: parseInt(at('w') || '9525', 10), relleno: rell, tapa: at('cap') || 'flat', compuesto: at('cmpd') || 'sng',
      guion: !guion ? 'solid' : guion.localName === 'prstDash' ? guion.getAttribute('val') || 'solid'
        : hijos(guion, 'ds').flatMap((d) => [ent(d, 'd', 100000) / 100000, ent(d, 'sp', 100000) / 100000]),
      union: union ? union.localName : 'round', inicio: punta('headEnd'), fin: punta('tailEnd'),
    };
  }

  function sombraDeLista(lista, cc) {
    const s = hijo(lista, 'outerShdw');
    if (!s) return null;
    const c = color(s, cc);
    if (!c) return null;
    return { desenfoque: ent(s, 'blurRad', 0), distancia: ent(s, 'dist', 0), direccion: ent(s, 'dir', 0) / 60000,
      sx: ent(s, 'sx', 100000) / 100000, sy: ent(s, 'sy', 100000) / 100000, color: c };
  }
  function sombra(spPrs, estilos, ctx) {
    for (const s of spPrs) {
      const lista = hijo(s, 'effectLst');
      if (lista) return sombraDeLista(lista, ctx.cc);
    }
    const ref = estilos.map((s) => hijo(s, 'effectRef')).find(Boolean);
    const idx = ref ? ent(ref, 'idx', 0) : 0;
    const estilo = idx > 0 ? ctx.estilosTema.efecto[idx - 1] : null;
    return estilo ? sombraDeLista(hijo(estilo, 'effectLst'), { tema: ctx.temaColor, ph: color(ref, ctx.cc) }) : null;
  }

  function forma(el, ctx) {
    const herencia = herenciaDe(el, ctx);
    const props = propiedadesForma(el, ctx, herencia);
    if (!props.xfrm) return null;
    const f = { tipo: 'forma', ...props };
    // Una forma rellena de imagen en un marcador de imagen vacío, etc.: igual.
    const txBody = hijo(el, 'txBody');
    if (txBody) f.texto = texto(txBody, el, herencia, ctx);
    // Texto en un rectángulo propio (SmartArt: dsp:txXfrm).
    const txXfrm = leerXfrm(hijo(el, 'txXfrm'));
    if (txXfrm && f.texto) f.cajaTexto = txXfrm;
    return f;
  }

  // xfrm: la posición, cuando la da el contenedor (la vista previa de un
  // objeto OLE suele no traer la suya).
  function imagenForma(el, ctx, xfrm = null) {
    const herencia = herenciaDe(el, ctx);
    const props = propiedadesForma(el, ctx, herencia);
    if (xfrm) props.xfrm = xfrm;
    if (!props.xfrm) return null;
    const imagen = imagenDeBlip(hijo(el, 'blipFill') || document.createElement('x'), ctx.cc, ctx.parteActual);
    if (!imagen) return null;
    return { tipo: 'imagen', ...props, imagen };
  }

  async function marco(el, ctx, motivos) {
    const xfrm = leerXfrm(hijo(el, 'xfrm'));
    const datos = ruta(el, 'graphic', 'graphicData');
    if (!xfrm || !datos) return null;
    const uri = datos.getAttribute('uri') || '';
    if (uri === URI_DIAGRAMA) {
      const dibujo = await dibujoDeDiagrama(datos, ctx);
      if (!dibujo) { motivos.push('SmartArt sin dibujo'); return null; }
      const arbol = todos(dibujo.doc, 'spTree')[0];
      const hijosDiagrama = await leerArbol(arbol, { ...ctx, parteActual: dibujo, cadena: [], esDiapositiva: true }, motivos);
      return { tipo: 'grupo', xfrm: { ...xfrm, hijos: { x: 0, y: 0, w: xfrm.w, h: xfrm.h } }, hijos: hijosDiagrama, relleno: null };
    }
    if (uri === URI_OLE) {
      const pic = todos(datos, 'pic')[0];
      const img = pic && imagenForma(pic, { ...ctx, cadena: [] }, xfrm);
      if (img) return img;
      motivos.push('objeto OLE sin imagen');
      return null;
    }
    if (uri === URI_TABLA) return tabla(hijo(datos, 'tbl'), xfrm, ctx, motivos);
    if (uri === URI_GRAFICO) {
      const g = await grafico(datos, ctx);
      if (g.modelo) return { tipo: 'grafico', xfrm, grafico: g.modelo };
      motivos.push(g.motivo);
      return { tipo: 'grafico', xfrm };
    }
    motivos.push('marco desconocido ' + uri);
    return null;
  }

  // --- Gráficos ---------------------------------------------------------------------------
  // La parte del gráfico (c:chart r:id) leída por pptx-grafico.js con los
  // colores, rellenos, líneas y fuentes de esta diapositiva. Devuelve
  // { modelo } o { motivo } (lo que no sabe dibujar queda para el respaldo).
  async function grafico(datos, ctx) {
    if (typeof PptxGrafico === 'undefined') return { motivo: 'gráfico' };
    const ref = hijo(datos, 'chart');
    const rel = ref && ctx.parteActual.rels.porId(ref.getAttributeNS(NS_R, 'id'));
    const parteGrafico = rel && rel.destino ? await ctx.parte(rel.destino) : null;
    if (!parteGrafico) return { motivo: 'gráfico sin datos' };
    return modeloDeGrafico(parteGrafico, ctx);
  }

  // El modelo de una parte c:chartSpace con los resolvedores del contexto
  // (tema, mapa de colores, fuentes).
  function modeloDeGrafico(parteGrafico, ctx) {
    const r = {
      relleno: (spPr) => { if (!spPr) return undefined; const e = primero(spPr, RELLENOS); return e ? relleno(e, ctx.cc, parteGrafico) : undefined; },
      // Un a:ln sin relleno (solo el ancho) no apaga la línea: lleva el color automático.
      linea: (spPr) => {
        const ln = hijo(spPr, 'ln');
        if (!ln) return undefined;
        if (!primero(ln, ['noFill', 'solidFill', 'gradFill', 'pattFill'])) return { sinColor: true, ancho: ln.hasAttribute('w') ? parseInt(ln.getAttribute('w'), 10) : null };
        return linea([spPr], [], ctx);
      },
      color: (el) => color(el, ctx.cc),
      fuente: (nombre) => resolverFuente(nombre, ctx),
      acento: (i) => { const [rr, g, b] = ctx.temaColor('accent' + i); return { r: rr, g, b, a: 1 }; },
      excel: !!ctx.excel,
    };
    try {
      return { modelo: PptxGrafico.leer(parteGrafico.raiz, r) };
    } catch (e) {
      return { motivo: e && e.message ? e.message : 'gráfico' };
    }
  }

  // Un gráfico que no vive en una diapositiva —el de una hoja de Excel (con su
  // propio tema) o uno armado con datos— hecho una diapositiva de un solo
  // gráfico, para dibujarlo con PptxDibujo.dibujar igual que a los de un .pptx.
  //   zip: ZipPropio del paquete (opcional; da el tema, las imágenes y las relaciones)
  //   rutaGrafico: la parte c:chartSpace dentro del zip; o raiz: el c:chartSpace ya leído
  //   ancho, alto: EMU
  //   excel: los valores que Excel usa cuando el gráfico no dice nada (estilo clásico); sin él, los de PowerPoint
  // Devuelve { pres, diapositiva } o tira un Error con el motivo.
  async function graficoSuelto({ zip, rutaGrafico, raiz, ancho, alto, excel = true }) {
    const paq = zip ? paquete(zip) : null;
    const sinRels = { lista: [], porId: () => undefined, porTipo: () => [] };
    const parteGrafico = rutaGrafico ? paq && await paq.parte(rutaGrafico) : raiz && { ruta: '', doc: raiz.ownerDocument, raiz, rels: sinRels };
    if (!parteGrafico) throw new Error('gráfico sin datos');
    let tema = null;
    for (const rutaTema of paq ? ['xl/theme/theme1.xml', 'ppt/theme/theme1.xml'] : []) { tema = await paq.parte(rutaTema); if (tema) break; }
    if (!tema) tema = { doc: parsear(TEMA_OFFICE) };
    const ctx = crearContexto({ sld: null, layout: null, master: null, tema, textoPorDefecto: null, numero: 1, parte: paq ? paq.parte : async () => null });
    ctx.parteActual = parteGrafico; ctx.cadena = []; ctx.estilosTabla = new Map(); ctx.excel = excel;
    const g = modeloDeGrafico(parteGrafico, ctx);
    if (!g.modelo) throw new Error(g.motivo);
    return {
      pres: { ancho, alto, cantidad: 1, imagen: paq ? paq.imagen : async () => null },
      diapositiva: {
        fondo: { tipo: 'solido', color: { r: 255, g: 255, b: 255, a: 1 } }, motivos: [], soportada: true,
        elementos: [{ tipo: 'grafico', nombre: 'Gráfico', xfrm: { x: 0, y: 0, w: ancho, h: alto, rot: 0 }, grafico: g.modelo }],
      },
    };
  }

  // El tema de Office de siempre (Calibri y sus seis colores de acento), para
  // un gráfico que no trae el suyo.
  const TEMA_OFFICE = '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Office">'
    + '<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
    + '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>'
    + '<a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>'
    + '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>'
    + '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>'
    + '</a:themeElements></a:theme>';

  // --- Tablas -----------------------------------------------------------------------------
  // Partes de un estilo de tabla, de menor a mayor prioridad (cada una pisa a
  // las anteriores solo en lo que define), y qué bandera de tblPr la activa.
  const PARTES_TABLA = ['wholeTbl', 'band1H', 'band2H', 'band1V', 'band2V', 'lastCol', 'firstCol', 'lastRow', 'firstRow', 'seCell', 'swCell', 'neCell', 'nwCell'];
  const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const LADOS_CELDA = { l: 'lnL', r: 'lnR', t: 'lnT', b: 'lnB', tlbr: 'lnTlToBr', bltr: 'lnBlToTr' };

  // Devuelve { tipo: 'tabla', xfrm, columnas: [EMU], filas: [EMU mínimo],
  // celdas: [{ fila, col, filas, cols, relleno, bordes, texto }], fondo }.
  // Cada borde: { linea (null = sin línea), prioridad }; quien dibuja decide
  // entre los dos lados de un borde compartido por la prioridad.
  function tabla(tbl, xfrm, ctx, motivos) {
    const tblPr = hijo(tbl, 'tblPr');
    const id = ((hijo(tblPr, 'tableStyleId') || {}).textContent || '').trim().toUpperCase();
    // El estilo: el que viene escrito en la tabla, el del archivo o, si el
    // archivo trae solo el GUID, el integrado (pptx-estilos-tabla.js).
    const estilo = hijo(tblPr, 'tableStyle') ? hijo(hijo(tblPr, 'tableStyle'), 'tblStyle')
      : id ? ctx.estilosTabla.get(id) || (typeof PptxEstilosTabla !== 'undefined' ? PptxEstilosTabla.definicion(id) : null) : null;
    // Un estilo que no está en el archivo ni entre los integrados: la tabla se dibuja sin estilo (con lo que cada celda dice por su cuenta).
    if (id && !estilo) motivos.push('estilo de tabla desconocido');
    const bandera = (a) => bool(tblPr, a);
    const columnas = hijos(hijo(tbl, 'tblGrid'), 'gridCol').map((c) => ent(c, 'w', 0));
    const trs = hijos(tbl, 'tr');
    const nF = trs.length, nC = columnas.length;
    if (!nF || !nC) return null;
    const parteDe = (n) => estilo && hijo(estilo, n);

    // Qué partes del estilo tocan la celda (fila r, columna c, con su tamaño).
    function partesDe(r, c, rs, cs) {
      const lista = [];
      const filaDatos = r - (bandera('firstRow') ? 1 : 0), colDatos = c - (bandera('firstCol') ? 1 : 0);
      const enDatosF = filaDatos >= 0 && !(bandera('lastRow') && r + rs - 1 === nF - 1);
      const enDatosC = colDatos >= 0 && !(bandera('lastCol') && c + cs - 1 === nC - 1);
      lista.push({ n: 'wholeTbl', zona: 'tabla' });
      if (bandera('bandRow') && enDatosF) lista.push({ n: filaDatos % 2 ? 'band2H' : 'band1H', zona: 'fila' });
      if (bandera('bandCol') && enDatosC) lista.push({ n: colDatos % 2 ? 'band2V' : 'band1V', zona: 'col' });
      if (bandera('lastCol') && c + cs - 1 === nC - 1) lista.push({ n: 'lastCol', zona: 'col' });
      if (bandera('firstCol') && c === 0) lista.push({ n: 'firstCol', zona: 'col' });
      if (bandera('lastRow') && r + rs - 1 === nF - 1) lista.push({ n: 'lastRow', zona: 'fila' });
      if (bandera('firstRow') && r === 0) lista.push({ n: 'firstRow', zona: 'fila' });
      const arriba = r === 0 && bandera('firstRow'), abajo = r + rs - 1 === nF - 1 && bandera('lastRow');
      const izq = c === 0 && bandera('firstCol'), der = c + cs - 1 === nC - 1 && bandera('lastCol');
      if (abajo && der) lista.push({ n: 'seCell', zona: 'celda' });
      if (abajo && izq) lista.push({ n: 'swCell', zona: 'celda' });
      if (arriba && der) lista.push({ n: 'neCell', zona: 'celda' });
      if (arriba && izq) lista.push({ n: 'nwCell', zona: 'celda' });
      return lista.map((p) => ({ ...p, el: parteDe(p.n), prioridad: PARTES_TABLA.indexOf(p.n) + 1 })).filter((p) => p.el);
    }

    // El borde de un lado según la zona de la parte: los de afuera (left…) solo
    // en el borde de la zona; adentro, insideH / insideV.
    function bordeDeParte(p, lado, r, c, rs, cs) {
      const tcBdr = ruta(p.el, 'tcStyle', 'tcBdr');
      if (!tcBdr) return undefined;
      const primeraF = r === 0, ultimaF = r + rs - 1 === nF - 1, primeraC = c === 0, ultimaC = c + cs - 1 === nC - 1;
      let nombre;
      const enFila = p.zona === 'fila' || p.zona === 'celda', enCol = p.zona === 'col' || p.zona === 'celda';
      switch (lado) {
        case 'l': nombre = enCol || primeraC ? 'left' : 'insideV'; break;
        case 'r': nombre = enCol || ultimaC ? 'right' : 'insideV'; break;
        case 't': nombre = enFila || primeraF ? 'top' : 'insideH'; break;
        case 'b': nombre = enFila || ultimaF ? 'bottom' : 'insideH'; break;
        case 'tlbr': nombre = 'tl2br'; break;
        case 'bltr': nombre = 'tr2bl'; break;
      }
      const el = hijo(tcBdr, nombre);
      if (!el) return undefined;
      return linea([el], [el], ctx);
    }

    // tcTxStyle de las partes, juntado en una lista de estilos de texto (los
    // nueve niveles iguales) que entra debajo del lstStyle de la celda.
    function estiloTexto(partes) {
      let b = null, i = null, colorXml = null, fuente = null;
      for (const p of partes) {
        const t = hijo(p.el, 'tcTxStyle');
        if (!t) continue;
        const sino = (v) => (v === 'on' ? '1' : v === 'off' ? '0' : null);
        b = sino(t.getAttribute('b')) || b; i = sino(t.getAttribute('i')) || i;
        const c = primero(t, COLORES);
        const ref = hijo(t, 'fontRef');
        if (c) colorXml = new XMLSerializer().serializeToString(c);
        else if (ref && primero(ref, COLORES)) colorXml = new XMLSerializer().serializeToString(primero(ref, COLORES));
        if (ref) fuente = ref.getAttribute('idx') === 'major' ? '+mj-lt' : '+mn-lt';
        const f = ruta(t, 'font', 'latin');
        if (f && f.getAttribute('typeface')) fuente = f.getAttribute('typeface');
      }
      if (b == null && i == null && !colorXml && !fuente) return null;
      const def = `<a:defRPr${b ? ` b="${b}"` : ''}${i ? ` i="${i}"` : ''}>${colorXml ? `<a:solidFill>${colorXml}</a:solidFill>` : ''}${fuente ? `<a:latin typeface="${fuente.replace(/[<&"]/g, '')}"/>` : ''}</a:defRPr>`;
      let niveles = '';
      for (let n = 1; n <= 9; n++) niveles += `<a:lvl${n}pPr>${def}</a:lvl${n}pPr>`;
      return parsear(`<a:lstStyle xmlns:a="${NS_A}">${niveles}</a:lstStyle>`).documentElement;
    }

    function rellenoDeParte(p) {
      const tcStyle = hijo(p.el, 'tcStyle');
      const f = hijo(tcStyle, 'fill');
      if (f) { const r = primero(f, RELLENOS); if (r) return { valor: relleno(r, ctx.cc, ctx.parteActual) }; }
      const ref = hijo(tcStyle, 'fillRef');
      if (ref) return { valor: rellenoDeReferencia(ref, ctx) || null };
      return undefined;
    }

    const celdas = [];
    trs.forEach((tr, r) => {
      hijos(tr, 'tc').forEach((tc, c) => {
        if (c >= nC || bool(tc, 'hMerge') || bool(tc, 'vMerge')) return;
        const rs = Math.max(1, Math.min(nF - r, ent(tc, 'rowSpan', 1))), cs = Math.max(1, Math.min(nC - c, ent(tc, 'gridSpan', 1)));
        const tcPr = hijo(tc, 'tcPr');
        const partes = partesDe(r, c, rs, cs);
        let rell = null;
        for (const p of partes) { const v = rellenoDeParte(p); if (v !== undefined) rell = v.valor; }
        const directo = primero(tcPr, RELLENOS);
        if (directo) rell = relleno(directo, ctx.cc, ctx.parteActual);
        const bordes = {};
        for (const [lado, nombre] of Object.entries(LADOS_CELDA)) {
          let b;
          for (const p of partes) { const v = bordeDeParte(p, lado, r, c, rs, cs); if (v !== undefined) b = { linea: v, prioridad: p.prioridad }; }
          if (hijo(tcPr, nombre)) b = { linea: linea([tcPr], [], ctx, nombre), prioridad: 100 };
          if (b) bordes[lado] = b;
        }
        const txBody = hijo(tc, 'txBody');
        const listas = [hijo(txBody, 'lstStyle'), 'estilo', estiloTexto(partes), hijo(ctx.txStyles, 'otherStyle'), ctx.textoPorDefecto].filter(Boolean);
        const margen = (a, def) => ent(tcPr, a, def);
        const cuerpoCelda = {
          margenes: { l: margen('marL', 91440), t: margen('marT', 45720), r: margen('marR', 91440), b: margen('marB', 45720) },
          ancla: (tcPr && tcPr.getAttribute('anchor')) || 't', anclaCentro: bool(tcPr, 'anchorCtr'), ajuste: 'square',
          vertical: (tcPr && tcPr.getAttribute('vert')) || 'horz', giro: 0, derecho: false, columnas: 1, espacioColumnas: 0,
          autoajuste: { tipo: 'no' }, primeroYUltimo: false,
        };
        const parrafos = hijos(txBody, 'p').map((p) => parrafo(p, listas, null, ctx));
        celdas.push({ fila: r, col: c, filas: rs, cols: cs, relleno: rell, bordes, texto: parrafos.length ? { cuerpo: cuerpoCelda, parrafos } : null });
      });
    });

    // Fondo de toda la tabla, debajo de las celdas: el relleno escrito en
    // tblPr (Tabla.Background en PowerPoint) o, si no hay, el tblBg del
    // estilo. Su sombra (effectLst de tblPr o del tblBg) es la del rectángulo
    // de la tabla entera (23-tablas-avanzadas, diapositiva 4).
    const tblBg = parteDe('tblBg');
    let fondoTabla = null, sombraTabla = null;
    const propio = primero(tblPr, RELLENOS);
    if (propio) fondoTabla = relleno(propio, ctx.cc, ctx.parteActual);
    else if (tblBg) {
      const r = primero(tblBg, RELLENOS);
      fondoTabla = r ? relleno(r, ctx.cc, ctx.parteActual) : hijo(tblBg, 'fillRef') ? rellenoDeReferencia(hijo(tblBg, 'fillRef'), ctx) || null : null;
    }
    const efectos = hijo(tblPr, 'effectLst') || ruta(tblBg, 'effect', 'effectLst');
    if (efectos) sombraTabla = sombraDeLista(efectos, ctx.cc);
    return { tipo: 'tabla', xfrm, columnas, filas: trs.map((tr) => ent(tr, 'h', 0)), celdas, fondo: fondoTabla, sombra: sombraTabla, rtl: bandera('rtl') };
  }

  async function dibujoDeDiagrama(datos, ctx) {
    const relIds = hijo(datos, 'relIds');
    const idDatos = relIds && relIds.getAttributeNS(NS_R, 'dm');
    const relDatos = idDatos && ctx.parteActual.rels.porId(idDatos);
    if (relDatos && relDatos.destino) {
      const partDatos = await ctx.parte(relDatos.destino);
      const ext = partDatos && todos(partDatos.doc, 'dataModelExt')[0];
      const rel = ext && ctx.parteActual.rels.porId(ext.getAttribute('relId'));
      if (rel && rel.destino) return ctx.parte(rel.destino);
    }
    const dibujos = ctx.parteActual.rels.lista.filter((r) => r.tipo === TIPO_DIBUJO_DIAGRAMA && r.destino);
    return dibujos.length === 1 ? ctx.parte(dibujos[0].destino) : null;
  }

  // --- Texto ---------------------------------------------------------------------------------
  // Las listas de estilos que aplican, en orden de prioridad. 'estilo' marca
  // dónde entra el fontRef de la forma (color y fuente de la letra).
  function listasDeTexto(txBody, el, herencia, ctx) {
    const listas = [hijo(txBody, 'lstStyle'), 'estilo'];
    for (const h of herencia) listas.push(ruta(h, 'txBody', 'lstStyle'));
    const marcador = datosMarcador(el);
    if (marcador) {
      const base = tipoBase(marcador.tipo);
      listas.push(hijo(ctx.txStyles, base === 'title' ? 'titleStyle' : base === 'body' ? 'bodyStyle' : 'otherStyle'));
    } else {
      listas.push(hijo(ctx.txStyles, 'otherStyle'), ctx.textoPorDefecto);
    }
    return listas.filter(Boolean);
  }
  const nivelDe = (lista, n) => hijo(lista, `lvl${n + 1}pPr`);

  function cuerpo(txBody, herencia) {
    const cadena = [hijo(txBody, 'bodyPr'), ...herencia.map((h) => ruta(h, 'txBody', 'bodyPr'))].filter(Boolean);
    const at = (a, def) => { for (const b of cadena) if (b.hasAttribute(a)) return b.getAttribute(a); return def; };
    const num = (a, def) => parseInt(at(a, def), 10);
    let autoajuste = { tipo: 'no' };
    for (const b of cadena) {
      const a = primero(b, ['normAutofit', 'spAutoFit', 'noAutofit']);
      if (!a) continue;
      if (a.localName === 'normAutofit') autoajuste = { tipo: 'reducir', escala: ent(a, 'fontScale', 100000) / 100000, reduccion: ent(a, 'lnSpcReduction', 0) / 100000 };
      else if (a.localName === 'spAutoFit') autoajuste = { tipo: 'forma' };
      break;
    }
    return {
      margenes: { l: num('lIns', '91440'), t: num('tIns', '45720'), r: num('rIns', '91440'), b: num('bIns', '45720') },
      ancla: at('anchor', 't'), anclaCentro: ['1', 'true'].includes(at('anchorCtr', '0')),
      ajuste: at('wrap', 'square'), vertical: at('vert', 'horz'), giro: num('rot', '0') / 60000,
      derecho: ['1', 'true'].includes(at('upright', '0')), columnas: num('numCol', '1') || 1, espacioColumnas: num('spcCol', '0'),
      autoajuste, primeroYUltimo: ['1', 'true'].includes(at('spcFirstLastPara', '0')),
    };
  }

  function texto(txBody, el, herencia, ctx) {
    const listas = listasDeTexto(txBody, el, herencia, ctx);
    const fontRef = ruta(el, 'style', 'fontRef');
    const cuerpoTexto = cuerpo(txBody, herencia);
    const parrafos = hijos(txBody, 'p').map((p) => parrafo(p, listas, fontRef, ctx));
    if (!parrafos.some((p) => p.trozos.some((t) => t.texto))) {
      // Sin texto: nada que dibujar (un marcador vacío no muestra su aviso).
      return null;
    }
    return { cuerpo: cuerpoTexto, parrafos };
  }

  function parrafo(p, listas, fontRef, ctx) {
    const pPr = hijo(p, 'pPr');
    const nivel = Math.min(8, ent(pPr, 'lvl', 0));
    const cadenaP = [pPr, ...listas.filter((l) => l !== 'estilo').map((l) => nivelDe(l, nivel))].filter(Boolean);
    const at = (a, def) => { for (const e of cadenaP) if (e.hasAttribute(a)) return e.getAttribute(a); return def; };
    const sub = (nombres) => { for (const e of cadenaP) { const s = primero(e, nombres); if (s) return s; } return null; };
    const espacio = (nombre) => {
      const e = sub([nombre]);
      if (!e) return null;
      const pct = hijo(e, 'spcPct'), pts = hijo(e, 'spcPts');
      return pct ? { pct: ent(pct, 'val', 100000) / 100000 } : pts ? { pts: ent(pts, 'val', 0) / 100 } : null;
    };
    const vineta = sub(['buNone', 'buAutoNum', 'buChar', 'buBlip']);
    const tamVineta = sub(['buSzTx', 'buSzPct', 'buSzPts']);
    const colorVineta = sub(['buClrTx', 'buClr']);
    const fuenteVineta = sub(['buFontTx', 'buFont']);
    const tabs = sub(['tabLst']);
    const resultado = {
      nivel, alineacion: at('algn', 'l'), margen: parseInt(at('marL', '0'), 10), sangria: parseInt(at('indent', '0'), 10),
      tabulacion: parseInt(at('defTabSz', '914400'), 10), rtl: ['1', 'true'].includes(at('rtl', '0')),
      interlineado: espacio('lnSpc') || { pct: 1 }, antes: espacio('spcBef') || { pts: 0 }, despues: espacio('spcAft') || { pts: 0 },
      tabs: tabs ? hijos(tabs, 'tab').map((t) => ({ pos: ent(t, 'pos', 0), algn: t.getAttribute('algn') || 'l' })) : [],
      vineta: null, trozos: [],
    };
    const runs = [];
    for (const h of p.children) {
      if (h.localName === 'r' || h.localName === 'fld') {
        let t = (hijo(h, 't') || { textContent: '' }).textContent;
        if (h.localName === 'fld' && (h.getAttribute('type') || '') === 'slidenum') t = String(ctx.numero);
        runs.push({ rPr: hijo(h, 'rPr'), texto: t });
      } else if (h.localName === 'br') runs.push({ rPr: hijo(h, 'rPr'), texto: '\n', salto: true });
    }
    resultado.trozos = runs.map((r) => ({ ...letra(r.rPr, listas, fontRef, nivel, ctx), texto: r.texto, salto: !!r.salto }));
    resultado.fin = letra(hijo(p, 'endParaRPr'), listas, fontRef, nivel, ctx);
    if (vineta && vineta.localName !== 'buNone') {
      const base = resultado.trozos.find((t) => !t.salto) || resultado.fin;
      const v = { tipo: vineta.localName === 'buAutoNum' ? 'numero' : vineta.localName === 'buChar' ? 'caracter' : 'imagen' };
      if (v.tipo === 'numero') { v.esquema = vineta.getAttribute('type') || 'arabicPeriod'; v.inicio = ent(vineta, 'startAt', 1); }
      if (v.tipo === 'caracter') v.caracter = vineta.getAttribute('char') || '•';
      v.fuente = fuenteVineta && fuenteVineta.localName === 'buFont' ? resolverFuente(fuenteVineta.getAttribute('typeface'), ctx) : base.fuente;
      v.color = colorVineta && colorVineta.localName === 'buClr' ? color(colorVineta, ctx.cc) : null;
      v.escala = !tamVineta || tamVineta.localName === 'buSzTx' ? 1 : tamVineta.localName === 'buSzPct' ? ent(tamVineta, 'val', 100000) / 100000 : null;
      v.puntos = tamVineta && tamVineta.localName === 'buSzPts' ? ent(tamVineta, 'val', 1800) / 100 : null;
      resultado.vineta = v;
    }
    return resultado;
  }

  function resolverFuente(nombre, ctx) {
    if (!nombre) return null;
    return nombre.startsWith('+') ? ctx.fuentes[nombre] || null : nombre;
  }

  // Propiedades de un trozo de texto, con la herencia aplicada.
  function letra(rPr, listas, fontRef, nivel, ctx) {
    const cadena = [rPr];
    for (const l of listas) cadena.push(l === 'estilo' ? 'estilo' : hijo(nivelDe(l, nivel), 'defRPr'));
    const at = (a) => { for (const e of cadena) if (e && e !== 'estilo' && e.hasAttribute(a)) return e.getAttribute(a); return null; };
    let rell = undefined, fuente = null, fuenteEa = null;
    for (const e of cadena) {
      if (!e) continue;
      if (e === 'estilo') {
        if (fontRef) {
          if (rell === undefined && primero(fontRef, COLORES)) rell = { tipo: 'solido', color: color(fontRef, ctx.cc) };
          if (!fuente) fuente = fontRef.getAttribute('idx') === 'major' ? '+mj-lt' : fontRef.getAttribute('idx') === 'minor' ? '+mn-lt' : null;
        }
        continue;
      }
      if (rell === undefined) { const r = primero(e, ['solidFill', 'gradFill', 'noFill', 'pattFill', 'blipFill']); if (r) rell = relleno(r, ctx.cc, ctx.parteActual); }
      if (!fuente) { const l = hijo(e, 'latin'); if (l && l.getAttribute('typeface')) fuente = l.getAttribute('typeface'); }
      if (!fuenteEa) { const l = hijo(e, 'ea'); if (l && l.getAttribute('typeface')) fuenteEa = l.getAttribute('typeface'); }
    }
    const enlace = !!(rPr && hijo(rPr, 'hlinkClick'));
    if (enlace && !(rPr && primero(rPr, ['solidFill', 'gradFill', 'noFill']))) rell = { tipo: 'solido', color: color({ localName: 'schemeClr', getAttribute: () => 'hlink', children: [] }, ctx.cc) };
    const resaltado = rPr && hijo(rPr, 'highlight');
    const efecto = rPr ? hijo(rPr, 'effectLst') : null;
    const u = at('u');
    return {
      tamano: (parseInt(at('sz') || '1800', 10)) / 100,
      negrita: ['1', 'true'].includes(at('b')), cursiva: ['1', 'true'].includes(at('i')),
      subrayado: u && u !== 'none' ? u : enlace && u == null ? 'sng' : null,
      tachado: at('strike') && at('strike') !== 'noStrike' ? at('strike') : null,
      base: parseInt(at('baseline') || '0', 10) / 100000, mayusculas: at('cap') || 'none',
      espaciado: parseInt(at('spc') || '0', 10) / 100,
      // Kerning solo desde este tamaño (pt); sin el atributo, PowerPoint no lo aplica.
      kern: parseInt(at('kern') || '0', 10) / 100,
      fuente: resolverFuente(fuente || '+mn-lt', ctx) || 'Calibri', fuenteEa: resolverFuente(fuenteEa, ctx),
      relleno: rell === undefined ? { tipo: 'solido', color: color({ localName: 'schemeClr', getAttribute: () => 'tx1', children: [] }, ctx.cc) } : rell,
      resaltado: resaltado ? color(resaltado, ctx.cc) : null,
      sombra: efecto ? sombraDeLista(efecto, ctx.cc) : null,
    };
  }

  // --- Imágenes -------------------------------------------------------------------------------
  function tipoDeImagen(b) {
    if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return 'png';
    if (b[0] === 0xff && b[1] === 0xd8) return 'jpg';
    if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return 'gif';
    if (b[0] === 0x42 && b[1] === 0x4d) return 'bmp';
    if (b[0] === 0x01 && b[1] === 0x00 && b[2] === 0x00 && b[3] === 0x00 && b[40] === 0x20 && b[41] === 0x45 && b[42] === 0x4d && b[43] === 0x46) return 'emf';
    if ((b[0] === 0xd7 && b[1] === 0xcd && b[2] === 0xc6 && b[3] === 0x9a) || (b[0] === 0x01 && b[1] === 0x00 && b[2] === 0x09 && b[3] === 0x00)) return 'wmf';
    if (b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46) return 'webp';
    const inicio = new TextDecoder().decode(b.subarray(0, 300));
    if (/<svg|<\?xml/i.test(inicio)) return 'svg';
    return null;
  }
  // Puntos por pulgada guardados en la imagen (PNG pHYs, JPEG JFIF); si no, 96.
  function dpiDe(b, tipo) {
    if (tipo === 'png') {
      for (let k = 8; k + 8 < b.length && k < 65536;) {
        const largo = (b[k] << 24 | b[k + 1] << 16 | b[k + 2] << 8 | b[k + 3]) >>> 0;
        const nombre = String.fromCharCode(b[k + 4], b[k + 5], b[k + 6], b[k + 7]);
        if (nombre === 'pHYs' && b[k + 16] === 1) {
          const x = (b[k + 8] << 24 | b[k + 9] << 16 | b[k + 10] << 8 | b[k + 11]) >>> 0;
          const y = (b[k + 12] << 24 | b[k + 13] << 16 | b[k + 14] << 8 | b[k + 15]) >>> 0;
          if (x && y) return [x * 0.0254, y * 0.0254];
        }
        if (nombre === 'IDAT') break;
        k += 12 + largo;
      }
    }
    if (tipo === 'jpg' && b[6] === 0x4a && b[7] === 0x46 && b[8] === 0x49 && b[9] === 0x46) {
      const unidad = b[13], x = b[14] << 8 | b[15], y = b[16] << 8 | b[17];
      if (x && y && unidad === 1) return [x, y];
      if (x && y && unidad === 2) return [x * 2.54, y * 2.54];
    }
    return [96, 96];
  }
  async function cargarImagen(zip, rutaImagen) {
    const a = zip.file(rutaImagen);
    if (!a) return null;
    const bytes = await a.async('uint8array');
    const tipo = tipoDeImagen(bytes);
    if (!tipo || tipo === 'wmf') return null;
    if (tipo === 'emf') {
      if (typeof Emf === 'undefined') return null;
      const c = await Emf.dibujar(bytes, { escalaMaxima: 3, ladoMaximo: 4000 });
      return c ? { fuente: c, ancho: c.width, alto: c.height, dpiX: 96, dpiY: 96 } : null;
    }
    if (tipo === 'svg') {
      const url = URL.createObjectURL(new Blob([bytes], { type: 'image/svg+xml' }));
      try {
        const img = await new Promise((ok, mal) => { const i = new Image(); i.onload = () => ok(i); i.onerror = mal; i.src = url; });
        return { fuente: img, ancho: img.naturalWidth || 300, alto: img.naturalHeight || 150, dpiX: 96, dpiY: 96 };
      } finally { setTimeout(() => URL.revokeObjectURL(url), 60000); }
    }
    const bmp = await createImageBitmap(new Blob([bytes]));
    const [dpiX, dpiY] = dpiDe(bytes, tipo);
    // El JPEG original va tal cual al PDF (pdf-lienzo.js), sin volver a comprimir.
    if (tipo === 'jpg') bmp.bytesJpeg = bytes;
    return { fuente: bmp, ancho: bmp.width, alto: bmp.height, dpiX, dpiY };
  }

  // Los objetos de una hoja de Excel (xl/drawings/drawingN.xml: formas, imágenes, gráficos, grupos, cuadros de texto) resueltos con el mismo
  // motor: el dibujo de una hoja es DrawingML igual que el de una diapositiva, solo cambia el envoltorio (xdr:twoCellAnchor…) y el tema, que es el del libro.
  // Devuelve { anclajes: [elemento xdr:*Anchor], elementos(ancla, motivos) → elementos resueltos de esa ancla, imagen(ruta) } o null.
  async function dibujoExcel(zip, rutaDibujo, rutaTema) {
    const { parte, imagen } = paquete(zip);
    const dib = await parte(rutaDibujo);
    if (!dib) return null;
    const tema = rutaTema ? await parte(rutaTema) : null;
    const base = crearContexto({ sld: dib, layout: null, master: null, tema, textoPorDefecto: null, numero: 1, parte });
    const ctx = { ...base, estilosTabla: new Map(), parteActual: dib, cadena: [], esDiapositiva: true, excel: true }; // excel: los gráficos sin formato salen con el estilo clásico de Excel, no con el de PowerPoint
    // Las anclas hijas directas de la raíz y, de las que Excel envuelve en mc:AlternateContent (segmentaciones, controles de formulario), las de la rama
    // principal (Choice); el respaldo (Fallback) no se dibuja. xlsx-modelo.js (leerAnclas) cuenta las mismas, en el mismo orden.
    const anclajes = [];
    for (const e of dib.raiz.children) {
      if (/Anchor$/.test(e.localName)) anclajes.push(e);
      else if (e.localName === 'AlternateContent') { const ch = hijo(e, 'Choice'); for (const h of ch ? ch.children : []) if (/Anchor$/.test(h.localName)) anclajes.push(h); }
    }
    return { anclajes, elementos: (ancla, motivos) => leerArbol(ancla, ctx, motivos), imagen };
  }

  // El cargador de imágenes de un paquete (zip): imagen(ruta) → { fuente, ancho, alto, dpiX, dpiY } (las imágenes dentro de las celdas de Excel)
  const cargadorImagenes = (zip) => paquete(zip).imagen;

  return { abrir, graficoSuelto, color, dibujoExcel, cargadorImagenes };
})();

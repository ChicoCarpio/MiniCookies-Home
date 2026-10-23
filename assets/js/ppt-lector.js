/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Lector de .ppt propio (PowerPoint 97-2003, formato binario [MS-PPT]). Lo entrega como el «contrato de presentación» (ver pptx-desde-contrato.js) y, con aPptx,
// como un .pptx, que es lo que ya saben leer, mostrar e imprimir las herramientas de PowerPoint.
//
//   PptLector.parece(bytes)            → true si es un .ppt de PowerPoint 97 o posterior (contenedor CFB con el flujo «PowerPoint Document»)
//   await PptLector.leer(bytes)        → { contrato, mensajes: [clave…] }
//   await PptLector.aPptx(bytes)       → { bytes, mensajes }
//
// Lanza Error (con .codigo = 'cifrado' si tiene contraseña, 'antiguo' si es de PowerPoint 95 o anterior, 'danado' si está roto) con un mensaje listo para mostrar.
// Se llevan: diapositivas con su patrón (fondo y formas), formas predefinidas con relleno, línea, sombra, giro y volteo, texto con formato de carácter y de párrafo (fuente, tamaño,
// color, negrita, cursiva, subrayado, viñetas y numeración, sangrías, espacios, alineación), cuadros de texto y marcadores de posición (también el texto que viene del patrón),
// imágenes (PNG, JPEG; las DIB, EMF y WMF se llevan sin tocar, o se avisa si no se pueden), recorte de imágenes, grupos, líneas con flechas, fondos (color, degradado, imagen),
// diapositivas ocultas, enlaces, campos de número de diapositiva y fecha, notas del orador. Las tablas de .ppt son grupos de cuadros de texto y así se llevan.
// No se llevan: animaciones y transiciones, sonidos y vídeo, objetos OLE (gráficos, ecuaciones), WordArt, SmartArt, formas libres complejas.
// Requiere cfb-propio.js (y pptx-desde-contrato.js para aPptx). No depende del navegador.
(function () {
  'use strict';
  const MAX_REGISTROS = 4 * 1024 * 1024;     // registros que se recorren en total
  const MAX_DIAPOSITIVAS = 10000;
  const MAX_FORMAS = 150000;                 // formas de todo el archivo
  const MAX_PROFUNDIDAD = 24;
  const MAX_UNIDADES = 3000000;              // párrafos, tramos de texto y órdenes de trazo que produce el lector en total
  const MAX_TEXTO = 64 * 1024 * 1024;        // caracteres de todo el archivo
  const MAX_EDICIONES = 10000;               // guardados acumulados (cadena de UserEditAtom)
  const MAX_IMAGENES_BYTES = 256 * 1024 * 1024;
  const MAX_IMAGENES = 20000;
  const EMU_POR_MAESTRA = 1587.5;            // 1 unidad maestra = 1/576 de pulgada
  const sinProto = (o) => Object.assign(Object.create(null), o);
  const acotar = (x, min, max) => Math.min(Math.max(x, min), max);
  class Danado extends Error { constructor(m, codigo) { super(m); this.codigo = codigo || 'danado'; } }

  // ───────── Lectura de bytes sin salirse del arreglo ─────────
  const u8 = (b, p) => (p >= 0 && p < b.length ? b[p] : 0);
  const u16 = (b, p) => u8(b, p) | (u8(b, p + 1) << 8);
  const i16 = (b, p) => { const v = u16(b, p); return v > 32767 ? v - 65536 : v; };
  const u32 = (b, p) => (u8(b, p) | (u8(b, p + 1) << 8) | (u8(b, p + 2) << 16) | (u8(b, p + 3) << 24)) >>> 0;
  const i32 = (b, p) => u32(b, p) | 0;
  const emu = (u) => Math.round(u * EMU_POR_MAESTRA);

  const CP1252 = [0x20AC, 0x81, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x8D, 0x017D, 0x8F, 0x90, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x9D, 0x017E, 0x0178];
  function textoAnsi(b, desde, hasta) {
    let s = '';
    for (let i = desde; i < hasta; i += 8192) {
      const trozo = [];
      for (let k = i; k < Math.min(hasta, i + 8192); k++) { const c = b[k]; trozo.push(c >= 0x80 && c < 0xA0 ? CP1252[c - 0x80] : c); }
      s += String.fromCharCode.apply(null, trozo);
    }
    return s;
  }
  function textoUtf16(b, desde, hasta) {
    let s = '';
    for (let i = desde; i + 1 < hasta; i += 16384) {
      const trozo = [];
      for (let k = i; k + 1 < Math.min(hasta, i + 16384); k += 2) trozo.push(b[k] | (b[k + 1] << 8));
      s += String.fromCharCode.apply(null, trozo);
    }
    return s;
  }
  const hex2 = (n) => (n & 255).toString(16).padStart(2, '0').toUpperCase();
  const rgbHex = (r, g, b) => hex2(r) + hex2(g) + hex2(b);

  // Tipos de registro que usa el lector
  const RT = {
    Document: 0x03E8, DocumentAtom: 0x03E9, Slide: 0x03EE, SlideAtom: 0x03EF, Notes: 0x03F0, NotesAtom: 0x03F1, Environment: 0x03F2, SlidePersistAtom: 0x03F3, MainMaster: 0x03F8,
    SlideShowSlideInfoAtom: 0x03F9, PPDrawingGroup: 0x040B, PPDrawing: 0x040C, ColorSchemeAtom: 0x07F0, ExObjList: 0x0409, ExHyperlink: 0x0FD7, ExHyperlinkAtom: 0x0FD3, FontCollection: 0x07D5, FontEntityAtom: 0x0FB7,
    SlideListWithText: 0x0FF0, TextHeaderAtom: 0x0F9F, TextCharsAtom: 0x0FA0, StyleTextPropAtom: 0x0FA1, TextMasterStyleAtom: 0x0FA3, TextBytesAtom: 0x0FA8, TextRulerAtom: 0x0FA6, OutlineTextRefAtom: 0x0F9E,
    CString: 0x0FBA, SlideNumberMCAtom: 0x0FD8, DateTimeMCAtom: 0x0FF7, GenericDateMCAtom: 0x0FF8, HeaderMCAtom: 0x0FF9, FooterMCAtom: 0x0FDA, OEPlaceholderAtom: 0x0BC3, InteractiveInfo: 0x0FF2, InteractiveInfoAtom: 0x0FF3,
    TextInteractiveInfoAtom: 0x0FDF, UserEditAtom: 0x0FF5, PersistDirectoryAtom: 0x1772,
    DggContainer: 0xF000, BStoreContainer: 0xF001, DgContainer: 0xF002, SpgrContainer: 0xF003, SpContainer: 0xF004, Spgr: 0xF009, Fsp: 0xF00A, Fopt: 0xF00B, ClientTextbox: 0xF00D, ChildAnchor: 0xF00F,
    ClientAnchor: 0xF010, ClientData: 0xF011, FBSE: 0xF007, TertiaryFopt: 0xF122,
  };

  // Formas predefinidas de OfficeArt (msospt) → nombre de forma de ECMA-376. Las que no están se dibujan como rectángulo.
  const MSOSPT = sinProto({
    1: 'rect', 2: 'roundRect', 3: 'ellipse', 4: 'diamond', 5: 'triangle', 6: 'rtTriangle', 7: 'parallelogram', 8: 'trapezoid', 9: 'hexagon', 10: 'octagon', 11: 'plus', 12: 'star5', 13: 'rightArrow', 14: 'rightArrow', 15: 'homePlate',
    16: 'cube', 17: 'wedgeRoundRectCallout', 18: 'star16', 19: 'arc', 20: 'line', 21: 'plaque', 22: 'can', 23: 'donut', 32: 'straightConnector1', 33: 'bentConnector3', 34: 'bentConnector3', 35: 'bentConnector3', 36: 'bentConnector3',
    37: 'curvedConnector3', 38: 'curvedConnector3', 39: 'curvedConnector3', 40: 'curvedConnector3', 41: 'callout1', 44: 'accentCallout1', 47: 'borderCallout1', 50: 'accentBorderCallout1', 53: 'ribbon2', 54: 'ribbon', 55: 'chevron', 56: 'pentagon',
    57: 'noSmoking', 58: 'star8', 59: 'star16', 60: 'star32', 61: 'wedgeRectCallout', 62: 'wedgeRoundRectCallout', 63: 'wedgeEllipseCallout', 64: 'wave', 65: 'foldedCorner', 66: 'leftArrow', 67: 'downArrow', 68: 'upArrow', 69: 'leftRightArrow',
    70: 'upDownArrow', 71: 'star16', 72: 'star24', 73: 'lightningBolt', 74: 'heart', 75: 'frame', 76: 'quadArrow', 77: 'leftArrowCallout', 78: 'rightArrowCallout', 79: 'upArrowCallout', 80: 'downArrowCallout', 81: 'leftRightArrowCallout',
    82: 'upDownArrowCallout', 83: 'quadArrowCallout', 84: 'bevel', 85: 'leftBracket', 86: 'rightBracket', 87: 'leftBrace', 88: 'rightBrace', 89: 'leftUpArrow', 90: 'bentUpArrow', 91: 'bentArrow', 92: 'star24', 93: 'stripedRightArrow',
    94: 'notchedRightArrow', 95: 'blockArc', 96: 'smileyFace', 97: 'verticalScroll', 98: 'horizontalScroll', 99: 'circularArrow', 101: 'uturnArrow', 102: 'curvedRightArrow', 103: 'curvedLeftArrow', 104: 'curvedUpArrow', 105: 'curvedDownArrow',
    106: 'cloudCallout', 107: 'ellipseRibbon2', 108: 'ellipseRibbon', 109: 'flowChartProcess', 110: 'flowChartDecision', 111: 'flowChartInputOutput', 112: 'flowChartPredefinedProcess', 113: 'flowChartInternalStorage', 114: 'flowChartDocument',
    115: 'flowChartMultidocument', 116: 'flowChartTerminator', 117: 'flowChartPreparation', 118: 'flowChartManualInput', 119: 'flowChartManualOperation', 120: 'flowChartConnector', 121: 'flowChartPunchedCard', 122: 'flowChartPunchedTape',
    123: 'flowChartSummingJunction', 124: 'flowChartOr', 125: 'flowChartCollate', 126: 'flowChartSort', 127: 'flowChartExtract', 128: 'flowChartMerge', 130: 'flowChartOnlineStorage', 131: 'flowChartMagneticTape', 132: 'flowChartMagneticDisk',
    133: 'flowChartMagneticDrum', 134: 'flowChartDisplay', 135: 'flowChartDelay', 183: 'actionButtonBlank', 184: 'actionButtonHome', 185: 'actionButtonHelp', 186: 'actionButtonInformation', 187: 'actionButtonForwardNext', 188: 'actionButtonBackPrevious',
    189: 'actionButtonEnd', 190: 'actionButtonBeginning', 191: 'actionButtonReturn', 192: 'actionButtonDocument', 193: 'actionButtonSound', 194: 'actionButtonMovie', 202: 'rect',
  });
  const GUIONES = [null, 'dash', 'sysDot', 'dashDot', 'sysDashDot', 'sysDot', 'dash', 'lgDashDot', 'lgDashDotDot', 'dash'];
  const FLECHAS = [null, 'triangle', 'arrow', 'stealth', 'diamond', 'oval'];
  // msoLineDashing: 0 sólida, 1 guión largo corto, 2 punto, 3 guión punto, 4 guión punto punto, 5 punto, 6 guión, 7 guión punto, 8 guión punto punto
  const GUION_ECMA = [null, 'dash', 'sysDot', 'dashDot', 'lgDashDotDot', 'sysDot', 'dash', 'dashDot', 'lgDashDotDot', 'dash'];

  // Viñetas de las fuentes de símbolos de Office, a Unicode (el visor no tiene Wingdings)
  const WINGDINGS = sinProto({ 0xA7: '▪', 0x6E: '■', 0x6C: '●', 0x6F: '□', 0x71: '❑', 0x76: '❖', 0x77: '◆', 0xD8: '➢', 0xFC: '✓', 0xFB: '✗', 0xA8: '◻', 0xE8: '➔', 0x9F: '•', 0x9E: '·', 0xA4: '◉', 0x75: '◆', 0x4A: '☺' });

  // ───────── Ruta del lector ─────────
  class Lector {
    constructor(bytes) {
      this.mensajes = new Set();
      this.registros = 0;
      this.formas = 0;
      this.textoTotal = 0;
      this.bytesImagenes = 0;
      this.imagenesLeidas = 0;
      this.unidades = 0;
      let cfb;
      try { cfb = CfbPropio.leer(bytes, { danado: (d) => new Danado('El archivo está dañado y no se puede leer (' + d + ').') }); } catch (e) { throw e.codigo ? e : new Danado('El archivo está dañado y no se puede leer.'); }
      this.cfb = cfb;
      const flujo = (n) => { const e = cfb.entradas.find((x) => x.nombre === n && x.tipo === 2); return e ? cfb.flujo(e) : null; };
      this.pd = flujo('PowerPoint Document');
      this.usuario = flujo('Current User');
      this.imagenes = flujo('Pictures');
      if (!this.pd) {
        if (cfb.entradas.some((x) => /^EncryptedPackage$/i.test(x.nombre))) throw new Danado('La presentación tiene contraseña. Quitala en PowerPoint y volvé a intentarlo.', 'cifrado');
        throw new Danado('Este archivo no es una presentación de PowerPoint 97-2003.');
      }
      if (!this.usuario) throw new Danado('Esta presentación es de una versión de PowerPoint anterior a la 97 y no se puede leer.', 'antiguo');
    }
    aviso(c) { this.mensajes.add(c); }
    contarUnidad(n) { if ((this.unidades += n) > MAX_UNIDADES) throw new Danado('La presentación es demasiado grande y no se puede leer.'); }
    contarRegistro() { if (++this.registros > MAX_REGISTROS) throw new Danado('La presentación es demasiado grande o está dañada y no se puede leer.'); }

    // Registro en la posición p: { ver, inst, tipo, ini, fin } (fin recortado al flujo)
    reg(p, tope) {
      const b = this.pd;
      if (p < 0 || p + 8 > (tope == null ? b.length : Math.min(tope, b.length))) return null;
      const v = u16(b, p), tipo = u16(b, p + 2), len = u32(b, p + 4);
      const lim = tope == null ? b.length : Math.min(tope, b.length);
      if (len > lim - p - 8) return { ver: v & 15, inst: v >> 4, tipo, ini: p + 8, fin: lim, recortado: true };
      return { ver: v & 15, inst: v >> 4, tipo, ini: p + 8, fin: p + 8 + len };
    }
    hijos(r) {
      const lista = [];
      let p = r.ini;
      while (p + 8 <= r.fin) {
        this.contarRegistro();
        const h = this.reg(p, r.fin);
        if (!h) break;
        lista.push(h);
        p = h.fin;
      }
      return lista;
    }

    // ───── Historial de guardados: UserEditAtom → PersistDirectoryAtom ─────
    leerPersistencia() {
      const b = this.pd, cu = this.usuario;
      if (u32(cu, 12) === 0xF3D1C4DF) throw new Danado('La presentación tiene contraseña. Quitala en PowerPoint y volvé a intentarlo.', 'cifrado');
      if (u32(cu, 12) !== 0xE391C05F) throw new Danado('El archivo está dañado y no se puede leer (usuario actual).');
      let edicion = u32(cu, 16);
      const persist = new Map();
      const vistos = new Set();
      const dirsVistos = new Set();
      let docRef = 0;
      let n = 0;
      while (edicion && edicion + 8 <= b.length && !vistos.has(edicion) && n++ < MAX_EDICIONES) {
        vistos.add(edicion);
        const r = this.reg(edicion);
        if (!r || r.tipo !== RT.UserEditAtom || r.fin - r.ini < 28) break;
        const anterior = u32(b, r.ini + 8), dir = u32(b, r.ini + 12);
        if (!docRef) docRef = u32(b, r.ini + 16);
        if (r.fin - r.ini >= 32 && u32(b, r.ini + 28)) throw new Danado('La presentación tiene contraseña. Quitala en PowerPoint y volvé a intentarlo.', 'cifrado');
        const d = this.reg(dir);
        if (d && d.tipo === RT.PersistDirectoryAtom && !dirsVistos.has(dir)) {
          dirsVistos.add(dir);
          this.registros += (d.fin - d.ini) >> 2;
          let p = d.ini;
          while (p + 4 <= d.fin) {
            const cab = u32(b, p); p += 4;
            const id = cab & 0xFFFFF, cuantos = cab >>> 20;
            for (let k = 0; k < cuantos && p + 4 <= d.fin; k++, p += 4) { if (!persist.has(id + k)) persist.set(id + k, u32(b, p)); }
          }
        }
        edicion = anterior;
      }
      if (!persist.size || !docRef) throw new Danado('El archivo está dañado y no se puede leer (no tiene diapositivas).');
      this.persist = persist;
      this.docRef = docRef;
    }
    contenedor(id) {
      const p = this.persist.get(id);
      if (p == null) return null;
      return this.reg(p);
    }

    // ───── Documento ─────
    leerDocumento() {
      const doc = this.contenedor(this.docRef);
      if (!doc || doc.tipo !== RT.Document) throw new Danado('El archivo está dañado y no se puede leer (documento).');
      const b = this.pd;
      this.fuentes = [];
      this.maestras = [];      // [{ slideId, persist }]
      this.listaDiap = [];     // [{ persist, slideId, textos: [...] }]
      this.listaNotas = new Map();   // slideId de las notas → persist
      this.estiloEntorno = null;
      this.exEnlaces = new Map();
      this.fopt0 = null;
      this.blips = [];
      for (const h of this.hijos(doc)) {
        if (h.tipo === RT.DocumentAtom) {
          this.ancho = emu(u32(b, h.ini)); this.alto = emu(u32(b, h.ini + 4));
        } else if (h.tipo === RT.Environment) {
          for (const e of this.hijos(h)) {
            if (e.tipo === RT.FontCollection) for (const f of this.hijos(e)) if (f.tipo === RT.FontEntityAtom) this.fuentes.push(textoUtf16(b, f.ini, Math.min(f.fin, f.ini + 62)).replace(/\u0000.*$/s, ''));
            if (e.tipo === RT.TextMasterStyleAtom) (this.estiloEntorno = this.estiloEntorno || {})[e.inst] = this.leerEstiloMaestro(e);
          }
        } else if (h.tipo === RT.PPDrawingGroup) {
          this.leerGrupoDibujo(h);
        } else if (h.tipo === RT.SlideListWithText) {
          this.leerListaTextos(h);
        } else if (h.tipo === RT.ExObjList || h.tipo === 0x07D0) {
          this.leerEnlaces(h);
        }
      }
      if (!this.ancho || !this.alto || this.ancho > 51206400 || this.alto > 51206400) throw new Danado('El archivo está dañado y no se puede leer (tamaño de la diapositiva).');
    }
    leerListaTextos(h) {
      const b = this.pd;
      let actual = null;
      for (const r of this.hijos(h)) {
        if (r.tipo === RT.SlidePersistAtom) {
          const entrada = { persist: u32(b, r.ini), slideId: u32(b, r.ini + 12), textos: [] };
          actual = entrada;
          if (h.inst === 0) this.listaDiap.push(entrada);
          else if (h.inst === 1) this.maestras.push(entrada);
          else if (h.inst === 2) this.listaNotas.set(entrada.slideId, entrada);
        } else if (actual && r.tipo === RT.TextHeaderAtom) {
          actual.textos.push({ tipoTexto: u32(b, r.ini) });
        } else if (actual && actual.textos.length) {
          const t = actual.textos[actual.textos.length - 1];
          if (r.tipo === RT.TextCharsAtom) t.texto = textoUtf16(b, r.ini, r.fin);
          else if (r.tipo === RT.TextBytesAtom) t.texto = textoAnsi(b, r.ini, r.fin);
          else if (r.tipo === RT.StyleTextPropAtom) t.estilo = r;
        }
      }
    }
    leerEnlaces(h) {
      const b = this.pd;
      for (const e of this.hijos(h)) {
        if (e.tipo !== RT.ExHyperlink) continue;
        let id = null, destino = null, tipoDestino = null;
        for (const c of this.hijos(e)) {
          if (c.tipo === RT.ExHyperlinkAtom) id = u32(b, c.ini);
          else if (c.tipo === RT.CString) { if (c.inst === 1) destino = textoUtf16(b, c.ini, c.fin); else if (c.inst === 3 && destino == null) tipoDestino = textoUtf16(b, c.ini, c.fin); }
        }
        if (id != null && destino) this.exEnlaces.set(id, destino);
      }
    }
    // Opciones por defecto del dibujo del documento y la lista de imágenes (BStoreContainer)
    leerGrupoDibujo(h) {
      for (const dg of this.hijos(h)) {
        if (dg.tipo !== RT.DggContainer) continue;
        for (const x of this.hijos(dg)) {
          if (x.tipo === RT.Fopt) this.fopt0 = this.opciones(x);
          else if (x.tipo === RT.BStoreContainer) for (const f of this.hijos(x)) { if (this.blips.length < MAX_IMAGENES) this.blips.push(f.tipo === RT.FBSE ? f : null); }
        }
      }
    }

    // ───── Estilos de texto del patrón ─────
    leerEstiloMaestro(r) {
      const b = this.pd;
      const niveles = [];
      let p = r.ini;
      const cuantos = Math.min(u16(b, p), 5); p += 2;
      for (let i = 0; i < cuantos && p < r.fin; i++) {
        if (r.inst >= 5) p += 2;
        const pf = this.leerPF(p, r.fin); p = pf.sig;
        const cf = this.leerCF(p, r.fin); p = cf.sig;
        niveles.push({ pf: pf.v, cf: cf.v });
      }
      return niveles;
    }
    leerPF(p, fin) {
      const b = this.pd;
      const v = {};
      if (p + 4 > fin) return { v, sig: fin };
      const m = u32(b, p); p += 4;
      const leer2 = () => { const x = p + 2 <= fin ? u16(b, p) : 0; p += 2; return x; };
      if (m & 0x0F) { const f = leer2(); if (m & 1) v.vineta = !!(f & 1); if (m & 2) v.vinetaFuente = !!(f & 2); if (m & 4) v.vinetaColorPropio = !!(f & 4); if (m & 8) v.vinetaTamanoPropio = !!(f & 8); }
      if (m & 0x80) v.vinetaCar = leer2();
      if (m & 0x10) v.vinetaFuenteRef = leer2();
      if (m & 0x40) { const s = leer2(); v.vinetaTamano = s > 32767 ? s - 65536 : s; }
      if (m & 0x20) { v.vinetaColor = this.colorIndice(p); p += 4; }
      if (m & 0x800) v.alin = leer2();
      if (m & 0x1000) { const s = leer2(); v.interlineado = s > 32767 ? s - 65536 : s; }
      if (m & 0x2000) { const s = leer2(); v.antes = s > 32767 ? s - 65536 : s; }
      if (m & 0x4000) { const s = leer2(); v.despues = s > 32767 ? s - 65536 : s; }
      if (m & 0x100) v.izq = leer2();
      if (m & 0x400) { const s = leer2(); v.sangria = s > 32767 ? s - 65536 : s; }
      if (m & 0x8000) v.tabPred = leer2();
      if (m & 0x100000) { const n = Math.min(leer2(), 32); const tabs = []; for (let i = 0; i < n; i++) { tabs.push({ pos: leer2(), tipo: leer2() }); } v.tabs = tabs; }
      if (m & 0x10000) v.alinFuente = leer2();
      if (m & 0xE0000) v.ajusteLinea = leer2();
      if (m & 0x200000) v.direccion = leer2();
      return { v, sig: Math.min(p, fin) };
    }
    leerCF(p, fin) {
      const b = this.pd;
      const v = {};
      if (p + 4 > fin) return { v, sig: fin };
      const m = u32(b, p); p += 4;
      const leer2 = () => { const x = p + 2 <= fin ? u16(b, p) : 0; p += 2; return x; };
      if (m & 0xFFFF) { const f = leer2(); if (m & 1) v.negrita = !!(f & 1); if (m & 2) v.cursiva = !!(f & 2); if (m & 4) v.subrayado = !!(f & 4); if (m & 0x10) v.sombra = !!(f & 0x10); }
      if (m & 0x10000) v.fuente = leer2();
      if (m & 0x200000) v.fuenteEA = leer2();
      if (m & 0x400000) v.fuenteAnsi = leer2();
      if (m & 0x800000) v.fuenteSimbolo = leer2();
      if (m & 0x20000) v.tamano = leer2();
      if (m & 0x40000) { v.color = this.colorIndice(p); p += 4; }
      if (m & 0x80000) { const s = leer2(); v.posicion = s > 32767 ? s - 65536 : s; }
      return { v, sig: Math.min(p, fin) };
    }
    // ColorIndexStruct: rojo, verde, azul, índice (0xFE = el color es el RGB; 0..7 = color del esquema; 0xFF = ninguno)
    colorIndice(p) {
      const b = this.pd;
      const idx = u8(b, p + 3);
      if (idx === 0xFE) return { rgb: rgbHex(u8(b, p), u8(b, p + 1), u8(b, p + 2)) };
      if (idx <= 7) return { esquema: idx };
      return null;
    }
  }

  // ───────── OfficeArt: propiedades, imágenes, formas ─────────
  Object.assign(Lector.prototype, {
    // Tabla de propiedades de una forma (OfficeArtFOPT): id → valor, y los datos de las propiedades complejas
    opciones(r) {
      const b = this.pd;
      const p = new Map(), c = new Map();
      let pos = r.ini;
      const n = Math.min(r.inst, 1024);
      const complejos = [];
      for (let i = 0; i < n && pos + 6 <= r.fin; i++, pos += 6) {
        const opid = u16(b, pos), valor = u32(b, pos + 2), id = opid & 0x3FFF;
        p.set(id, valor);
        if (opid & 0x8000) complejos.push([id, valor]);
      }
      for (const [id, tam] of complejos) { if (pos + tam > r.fin) break; c.set(id, b.subarray(pos, pos + tam)); pos += tam; }
      return { p, c };
    },

    // Una imagen del BStoreContainer (índice desde 1) → { datos, tipo } | null
    imagen(indice) {
      const f = this.blips[indice - 1];
      if (!f) return null;
      this.cacheImagenes = this.cacheImagenes || new Map();
      if (this.cacheImagenes.has(indice)) return this.cacheImagenes.get(indice);
      const b = this.pd;
      let resultado = null;
      const foDelay = u32(b, f.ini + 28);
      const cbNombre = u8(b, f.ini + 33);
      let fuente, pos;
      if (foDelay !== 0xFFFFFFFF && this.imagenes && foDelay < this.imagenes.length) { fuente = this.imagenes; pos = foDelay; }
      else { fuente = b; pos = f.ini + 36 + cbNombre; }
      // varias entradas que señalan los mismos datos comparten una sola imagen
      const clave = (fuente === this.imagenes ? 'P' : 'D') + pos;
      this.porPosicion = this.porPosicion || new Map();
      if (this.porPosicion.has(clave)) resultado = this.porPosicion.get(clave);
      else { resultado = this.leerBlip(fuente, pos); this.porPosicion.set(clave, resultado); }
      this.cacheImagenes.set(indice, resultado);
      return resultado;
    },
    leerBlip(b, pos) {
      if (pos + 8 > b.length) { this.aviso('imagen-falta'); return null; }
      const v = u16(b, pos), tipo = u16(b, pos + 2), len = u32(b, pos + 4), inst = v >> 4;
      let fin = Math.min(b.length, pos + 8 + len);
      let o = pos + 8 + 16;
      let formato = null;
      if (tipo === 0xF01E) { formato = 'png'; if (inst === 0x6E1) o += 16; }
      else if (tipo === 0xF01D || tipo === 0xF02A) { formato = 'jpeg'; if (inst === 0x46B || inst === 0x6E3) o += 16; }
      else if (tipo === 0xF01F) { formato = 'dib'; if (inst === 0x7A9) o += 16; }
      else if (tipo === 0xF01A) { formato = 'emf'; if (inst === 0x3D5) o += 16; }
      else if (tipo === 0xF01B) { formato = 'wmf'; if (inst === 0x217) o += 16; }
      else if (tipo === 0xF029) { formato = 'tiff'; if (inst === 0x6E5) o += 16; }
      else { this.aviso('imagen-formato'); return null; }
      if (formato === 'png' || formato === 'jpeg') {
        const datos = b.subarray(o + 1, fin);
        if (!datos.length) { this.aviso('imagen-falta'); return null; }
        if ((this.bytesImagenes += datos.length) > MAX_IMAGENES_BYTES || ++this.imagenesLeidas > MAX_IMAGENES) throw new Danado('Las imágenes de la presentación ocupan demasiado y no se puede leer.');
        return { datos: new Uint8Array(datos), tipo: formato };
      }
      if (formato === 'emf' || formato === 'wmf') {
        // metarchivo: cabecera de 34 bytes (tamaño sin comprimir, límites, tamaño, tamaño guardado, compresión, filtro) y los datos (zlib si compresión = 0)
        const h = o;
        const sinComprimir = u32(b, h), guardado = u32(b, h + 28), compresion = u8(b, h + 32);
        const datos = b.subarray(h + 34, Math.min(fin, h + 34 + guardado));
        if (!datos.length || sinComprimir > 64 * 1024 * 1024) { this.aviso('imagen-falta'); return null; }
        if ((this.bytesImagenes += sinComprimir) > MAX_IMAGENES_BYTES || ++this.imagenesLeidas > MAX_IMAGENES) throw new Danado('Las imágenes de la presentación ocupan demasiado y no se puede leer.');
        return { metarchivo: true, datos: new Uint8Array(datos), tipo: formato, comprimido: compresion === 0, sinComprimir };
      }
      this.aviso('imagen-formato');
      return null;
    },

    // ───── Dibujo (PPDrawing) ─────
    leerDibujo(ppd) {
      const res = { formas: [], fondo: null };
      const dg = this.hijos(ppd).find((h) => h.tipo === RT.DgContainer);
      if (!dg) return res;
      for (const h of this.hijos(dg)) {
        if (h.tipo === RT.SpgrContainer) res.formas = this.leerGrupo(h, 0).hijos;
        else if (h.tipo === RT.SpContainer) { const sp = this.leerSp(h); if (sp && (sp.flags & 0x400)) res.fondo = sp; }
      }
      return res;
    },
    // El contenedor de un grupo: el primer SpContainer describe el grupo, los demás son sus hijos
    leerGrupo(c, prof) {
      if (prof > MAX_PROFUNDIDAD) { this.aviso('grupos-anidados'); return { hijos: [] }; }
      const hijos = this.hijos(c);
      let info = null;
      const lista = [];
      hijos.forEach((h, i) => {
        if (h.tipo === RT.SpContainer) {
          const sp = this.leerSp(h);
          if (i === 0) info = sp; else if (sp) lista.push(sp);
        } else if (h.tipo === RT.SpgrContainer) {
          const g = this.leerGrupo(h, prof + 1);
          if (g.info) { g.info.hijos = g.hijos; g.info.esGrupo = true; lista.push(g.info); } else for (const x of g.hijos) lista.push(x);
        }
      });
      return { info, hijos: lista };
    },
    leerSp(c) {
      if (++this.formas > MAX_FORMAS) throw new Danado('La presentación tiene demasiadas formas y no se puede leer.');
      const b = this.pd;
      const sp = { flags: 0, tipoSp: 0, props: new Map(), comp: new Map(), ancla: null, anclaHijo: false, rectHijos: null, colocador: null, texto: null, enlace: null, pf9: null, refTexto: null };
      for (const h of this.hijos(c)) {
        switch (h.tipo) {
          case RT.Fsp: sp.id = u32(b, h.ini); sp.flags = u32(b, h.ini + 4); sp.tipoSp = h.inst; break;
          case RT.Fopt: case RT.TertiaryFopt: {
            const o = this.opciones(h);
            for (const [k, v] of o.p) {
              // las propiedades booleanas (ids terminados en 0x3F) se combinan bit a bit: la segunda tabla solo cambia los bits que marca
              const viejo = sp.props.get(k);
              if (viejo !== undefined && (k & 0x3F) === 0x3F) {
                const marca = (v >>> 16) & 0xFFFF;
                sp.props.set(k, ((((viejo >>> 16) | marca) << 16) | (((viejo & 0xFFFF) & ~marca) | (v & 0xFFFF & marca))) >>> 0);
              } else sp.props.set(k, v);
            }
            for (const [k, v] of o.c) sp.comp.set(k, v);
            break;
          }
          case RT.Spgr: sp.rectHijos = { izq: i32(b, h.ini), arr: i32(b, h.ini + 4), der: i32(b, h.ini + 8), aba: i32(b, h.ini + 12) }; break;
          case RT.ChildAnchor: if (h.fin - h.ini >= 16) { sp.ancla = { izq: i32(b, h.ini), arr: i32(b, h.ini + 4), der: i32(b, h.ini + 8), aba: i32(b, h.ini + 12) }; sp.anclaHijo = true; } break;
          case RT.ClientAnchor:
            if (h.fin - h.ini >= 16) sp.ancla = { arr: i32(b, h.ini), izq: i32(b, h.ini + 4), der: i32(b, h.ini + 8), aba: i32(b, h.ini + 12) };
            else if (h.fin - h.ini >= 8) sp.ancla = { arr: i16(b, h.ini), izq: i16(b, h.ini + 2), der: i16(b, h.ini + 4), aba: i16(b, h.ini + 6) };
            break;
          case RT.ClientData: this.leerDatosCliente(h, sp); break;
          case RT.ClientTextbox: sp.texto = this.leerCuadroTexto(h); break;
          default: break;
        }
      }
      return sp;
    },
    leerDatosCliente(c, sp) {
      const b = this.pd;
      for (const h of this.hijos(c)) {
        if (h.tipo === RT.OEPlaceholderAtom) sp.colocador = { posicion: i32(b, h.ini), id: u8(b, h.ini + 4), tam: u8(b, h.ini + 5) };
        else if (h.tipo === RT.OutlineTextRefAtom) sp.refTexto = u32(b, h.ini);
        else if (h.tipo === RT.InteractiveInfo && h.inst === 0) sp.enlace = this.enlaceDe(h) || sp.enlace;     // al hacer clic (inst 1 = al pasar el ratón)
        else if (h.tipo === 0x1388) sp.pf9 = this.leerPP9(h) || sp.pf9;
      }
    },
    // InteractiveInfo → destino del enlace (URL, o { diapositiva })
    enlaceDe(h) {
      const b = this.pd;
      for (const a of this.hijos(h)) {
        if (a.tipo !== RT.InteractiveInfoAtom) continue;
        const id = u32(b, a.ini + 4), accion = u8(b, a.ini + 8), salto = u8(b, a.ini + 10), tipo = u8(b, a.ini + 12);
        if (tipo === 8 || tipo === 0) { const d = this.exEnlaces.get(id); if (d && /^(https?:|mailto:)/i.test(d)) return d; if (d && d.length < 320 && /^[^:\\/@]+@[^:\\/@]+$/.test(d)) return 'mailto:' + d; }
        if (accion === 3 && tipo === 7) { const d = this.exEnlaces.get(id); const m = d && /^(\d+),(\d+)/.exec(d); if (m) return { slideId: +m[1] }; }
        if (accion === 1 && salto === 1) return { relativa: 'siguiente' };
        if (accion === 1 && salto === 2) return { relativa: 'anterior' };
      }
      return null;
    },
    // ProgTags "___PPT9": las numeraciones automáticas por tramo de párrafos
    leerPP9(h) {
      const b = this.pd;
      for (const t of this.hijos(h)) {
        if (t.tipo !== 0x138A) continue;
        const partes = this.hijos(t);
        const nombre = partes.find((x) => x.tipo === RT.CString);
        if (!nombre || textoUtf16(b, nombre.ini, nombre.fin) !== '___PPT9') continue;
        const datos = partes.find((x) => x.tipo === 0x138B);
        if (!datos) continue;
        const sub = this.reg(datos.ini);
        if (!sub || sub.tipo !== 0x0FAC) continue;
        const lista = [];
        let p = sub.ini;
        while (p + 4 <= sub.fin && lista.length < 4096) {
          const m = u32(b, p); p += 4;
          const e = {};
          if (m & 0x00800000) p += 2;
          if (m & 0x01000000) { e.auto = u16(b, p); p += 2; }
          if (m & 0x02000000) { e.esquema = u16(b, p); e.inicio = u16(b, p + 2); p += 4; }
          if (p + 8 > sub.fin) { lista.push(e); break; }
          const cf = u32(b, p), si = u32(b, p + 4);
          p += 8;
          lista.push(e);
          if (cf || si) break;
        }
        return lista;
      }
      return null;
    },
    leerCuadroTexto(c) {
      const b = this.pd;
      const t = { tipoTexto: 4, texto: '', estilo: null, regla: null, campos: new Map(), enlaces: [] };
      let enlacePendiente = null;
      for (const h of this.hijos(c)) {
        switch (h.tipo) {
          case RT.TextHeaderAtom: t.tipoTexto = u32(b, h.ini); break;
          case RT.TextCharsAtom: t.texto = textoUtf16(b, h.ini, h.fin); break;
          case RT.TextBytesAtom: t.texto = textoAnsi(b, h.ini, h.fin); break;
          case RT.StyleTextPropAtom: t.estilo = h; break;
          case RT.TextRulerAtom: t.regla = h; break;
          case RT.SlideNumberMCAtom: t.campos.set(u32(b, h.ini), 'numero'); break;
          case RT.DateTimeMCAtom: case RT.GenericDateMCAtom: t.campos.set(u32(b, h.ini), 'fecha'); break;
          case RT.InteractiveInfo: enlacePendiente = this.enlaceDe(h); break;
          case RT.TextInteractiveInfoAtom: if (enlacePendiente) t.enlaces.push({ ini: u32(b, h.ini), fin: u32(b, h.ini + 4), destino: enlacePendiente }); break;
          default: break;
        }
      }
      return t;
    },
  });

  // ───────── Del dibujo al contrato ─────────
  const ALIN = ['l', 'ctr', 'r', 'just', 'just', 'l', 'l'];
  const ANCLA_TEXTO = ['t', 'ctr', 'b', 't', 'ctr', 'b', 't', 'b', 't', 'b'];
  const GUION = [null, 'sysDash', 'sysDot', 'sysDashDot', 'sysDashDotDot', 'dot', 'dash', 'lgDash', 'dashDot', 'lgDashDot', 'lgDashDotDot'];
  const PUNTA = [null, 'triangle', 'stealth', 'diamond', 'oval', 'arrow'];
  const NUMERACION = ['alphaLcPeriod', 'alphaUcPeriod', 'arabicParenR', 'arabicPeriod', 'romanLcParenBoth', 'romanLcParenR', 'romanLcPeriod', 'romanUcPeriod', 'alphaLcParenBoth', 'alphaLcParenR', 'alphaUcParenBoth', 'alphaUcParenR', 'arabicParenBoth', 'arabicPlain', 'romanUcParenBoth', 'romanUcParenR'];
  // marcador de posición de la diapositiva (OEPlaceholderAtom) → el del patrón que le da el formato
  const COLOCADOR_PATRON = sinProto({ 0x0D: 1, 0x0E: 2, 0x0F: 3, 0x10: 4, 0x11: 5, 0x12: 6, 0x13: 2, 0x14: 2, 0x15: 2, 0x16: 2, 0x17: 2, 0x18: 2, 0x19: 6, 0x1A: 2 });
  const PADRE_ESTILO = sinProto({ 5: 1, 6: 0, 7: 1, 8: 1 });
  const SIN_ESQUEMA = ['FFFFFF', '000000', '808080', '000000', 'E6E6E6', '4472C4', 'ED7D31', 'A5A5A5'];

  Object.assign(Lector.prototype, {
    // Valor de una propiedad buscando de la forma hacia atrás por los mapas de base (documento → patrón → forma)
    prop(mapas, id) { for (let i = mapas.length - 1; i >= 0; i--) { const v = mapas[i] ? mapas[i].get(id) : undefined; if (v !== undefined) return v; } return undefined; },
    // Propiedad booleana (los 16 bits de arriba dicen cuáles de los 16 de abajo están puestos): el bit de la posición `bit` (0..15)
    booleano(mapas, id, bit) {
      for (let i = mapas.length - 1; i >= 0; i--) {
        const v = mapas[i] ? mapas[i].get(id) : undefined;
        if (v !== undefined && ((v >>> (16 + bit)) & 1)) return !!((v >>> bit) & 1);
      }
      return undefined;
    },
    colorOA(v, ctx) {
      if (v === undefined || v === null) return null;
      const r = v & 255, g = (v >>> 8) & 255, b = (v >>> 16) & 255, f = (v >>> 24) & 255;
      if (f & 0x08) return ctx.esquema[r] || SIN_ESQUEMA[r] || '000000';
      if (f & 0x10) {
        if (r === 0xF0) return ctx.esquema[4]; if (r === 0xF1) return ctx.esquema[1]; if (r === 0xF2) return ctx.esquema[2];
        if (r === 0xF3) return ctx.esquema[1]; if (r === 0xF4 || r === 0xF5) return ctx.esquema[0];
        return '000000';
      }
      return rgbHex(r, g, b);
    },
    colorTexto(c, ctx) {
      if (!c) return null;
      if (c.rgb) return c.rgb;
      if (c.esquema != null) return ctx.esquema[c.esquema] || SIN_ESQUEMA[c.esquema];
      return null;
    },
    fijo(v) { return (v | 0) / 65536; },

    // fillShadeColors (0x197): lista de colores con su posición (fija 16.16) → paradas del degradado
    paradasComplejas(sp, ctx) {
      const d = sp && sp.comp && sp.comp.get(0x197);
      if (!d || d.length < 6) return null;
      const n = Math.min(u16(d, 0), 64), cb = u16(d, 4);
      if (cb !== 8 || 6 + n * 8 > d.length || n < 2) return null;
      const out = [];
      for (let i = 0; i < n; i++) out.push({ pos: acotar(((i32(d, 6 + i * 8 + 4)) / 65536) * 100, 0, 100), color: this.colorOA(u32(d, 6 + i * 8), ctx) || 'FFFFFF' });
      return out.sort((a, b) => a.pos - b.pos);
    },
    relleno(mapas, ctx, implicito, sp) {
      const lleno = this.booleano(mapas, 0x1BF, 4);
      if (lleno === false || (lleno === undefined && !implicito)) return null;
      const tipo = this.prop(mapas, 0x180) || 0;
      const opac = this.prop(mapas, 0x182);
      const alfa = opac === undefined ? undefined : acotar(opac / 65536, 0, 1);
      const color = this.colorOA(this.prop(mapas, 0x181), ctx) || 'FFFFFF';
      if (tipo === 3 || tipo === 2) {
        const bid = this.prop(mapas, 0x186);
        const im = bid ? this.imagen(bid) : null;
        if (im) return { imagen: im };
        return { color };
      }
      if (tipo >= 4 && tipo <= 8) {
        const fondo = this.colorOA(this.prop(mapas, 0x183), ctx) || 'FFFFFF';
        const angulo = this.fijo(this.prop(mapas, 0x18B) || 0);
        const foco = this.prop(mapas, 0x18C);
        const varias = this.paradasComplejas(sp, ctx);
        if (varias) return { degradado: { angulo: ((90 - angulo) % 360 + 360) % 360, radial: tipo === 5 || tipo === 6 ? true : undefined, paradas: varias } };
        const alReves = (foco === undefined ? 0 : foco | 0) < 50;      // con el foco en 0 el degradado empieza por el color de fondo
        const paradas = alReves ? [{ pos: 0, color: fondo, alfa }, { pos: 100, color, alfa }] : [{ pos: 0, color, alfa }, { pos: 100, color: fondo, alfa }];
        return { degradado: { angulo: ((90 - angulo) % 360 + 360) % 360, radial: tipo === 5 || tipo === 6 ? true : undefined, paradas } };
      }
      return alfa !== undefined && alfa < 1 ? { color, alfa } : { color };
    },
    linea(mapas, ctx) {
      const hay = this.booleano(mapas, 0x1FF, 3);
      if (!hay) return null;
      const l = { color: this.colorOA(this.prop(mapas, 0x1C0), ctx) || '000000', ancho: this.prop(mapas, 0x1CB) === undefined ? 9525 : this.prop(mapas, 0x1CB) };
      let guion = GUION[this.prop(mapas, 0x1CE) || 0];
      if (this.prop(mapas, 0x1CE) === 5 && this.prop(mapas, 0x1D7) === 0) guion = 'sysDot';     // punto redondo (con remate redondo); con remate cuadrado es el punto cuadrado
      if (guion) l.guion = guion;
      const cap = this.prop(mapas, 0x1D7); if (cap === 0) l.cap = 'rnd'; else if (cap === 1) l.cap = 'sq'; else if (cap === 2) l.cap = 'flat';
      const ini = PUNTA[this.prop(mapas, 0x1D0) || 0], fin = PUNTA[this.prop(mapas, 0x1D1) || 0];
      if (ini) l.inicio = ini; if (fin) l.fin = fin;
      const unir = this.prop(mapas, 0x1D6); if (unir === 0) l.union = 'bevel'; else if (unir === 2) l.union = 'round';
      const op = this.prop(mapas, 0x1C1); if (op !== undefined && op < 65536) l.alfa = op / 65536;
      return l;
    },
    sombra(mapas, ctx) {
      if (!this.booleano(mapas, 0x23F, 1)) return undefined;
      const s = { color: this.colorOA(this.prop(mapas, 0x201), ctx) || '000000', x: (this.prop(mapas, 0x205) | 0) || 25400, y: (this.prop(mapas, 0x206) | 0) || 25400 };
      const op = this.prop(mapas, 0x204); if (op !== undefined) s.alfa = acotar(op / 65536, 0, 1);
      return s;
    },

    // ───── Estilo del texto ─────
    estiloDe(ctx, tipoTexto, nivel) {
      const padre = PADRE_ESTILO[tipoTexto];
      if (padre !== undefined && nivel >= 0) {
        const clave = tipoTexto + ':' + nivel;
        ctx.cacheEstilo = ctx.cacheEstilo || new Map();
        if (!ctx.cacheEstilo.has(clave)) {
          const p = this.estiloDe(ctx, padre, nivel), propio = ctx.estilos && ctx.estilos[tipoTexto] && ctx.estilos[tipoTexto].length ? this.nivelPropio(ctx.estilos[tipoTexto], nivel) : { pf: {}, cf: {} };
          ctx.cacheEstilo.set(clave, { pf: Object.assign({}, p.pf, propio.pf), cf: Object.assign({}, p.cf, propio.cf) });
        }
        return ctx.cacheEstilo.get(clave);
      }
      return this.nivelPropio((ctx.estilos && ctx.estilos[tipoTexto]) || (ctx.estilos && ctx.estilos[4]) || (this.estiloEntorno && (this.estiloEntorno[tipoTexto] || this.estiloEntorno[4])) || [], nivel);
    },
    // el estilo de un nivel suma lo que cambian los niveles de arriba
    nivelPropio(lista, nivel) {
      // cada nivel solo trae lo que cambia respecto al anterior
      const n = acotar(nivel, 0, Math.max(0, lista.length - 1));
      if (!lista.acumulados) lista.acumulados = [];
      if (!lista.acumulados[n]) {
        const anterior = n > 0 ? this.nivelPropio(lista, n - 1) : { pf: {}, cf: {} };
        lista.acumulados[n] = lista[n] ? { pf: Object.assign({}, anterior.pf, lista[n].pf), cf: Object.assign({}, anterior.cf, lista[n].cf) } : anterior;
      }
      return lista.acumulados[n];
    },
    leerEstiloTexto(rec, largo) {
      const b = this.pd;
      const pfs = [], cfs = [];
      if (!rec) return { pfs, cfs };
      const total = largo + 1;
      let p = rec.ini, cubierto = 0;
      while (cubierto < total && p + 6 <= rec.fin && pfs.length < 200000) {
        const n = u32(b, p), nivel = u16(b, p + 4);
        const r = this.leerPF(p + 6, rec.fin);
        if (!n) break;
        pfs.push({ ini: cubierto, fin: cubierto + n, nivel: acotar(nivel, 0, 8), pf: r.v });
        cubierto += n; p = r.sig;
      }
      cubierto = 0;
      while (cubierto < total && p + 4 <= rec.fin && cfs.length < 200000) {
        const n = u32(b, p);
        const r = this.leerCF(p + 4, rec.fin);
        if (!n) break;
        cfs.push({ ini: cubierto, fin: cubierto + n, cf: r.v });
        cubierto += n; p = r.sig;
      }
      return { pfs, cfs };
    },
    nombreFuente(i) { return i == null ? null : this.fuentes[i] || null; },
    vinetaDe(pf, cf, ctx, colorTexto) {
      if (!pf.vineta) return undefined;
      let car = pf.vinetaCar === undefined ? 0x2022 : pf.vinetaCar;
      let fuente = pf.vinetaFuenteRef !== undefined ? this.nombreFuente(pf.vinetaFuenteRef) : null;
      if (/wingdings|symbol/i.test(fuente || '')) {
        const m = /wingdings/i.test(fuente) ? WINGDINGS[car & 0xFF] : car === 0xB7 || car === 0xF0B7 ? '•' : null;
        if (m) { car = m.codePointAt(0); fuente = 'Arial'; } else car &= 0xFF;
      }
      const v = { car: String.fromCodePoint(car || 0x2022), fuente: fuente || 'Arial' };
      const color = pf.vinetaColorPropio !== false && pf.vinetaColor ? this.colorTexto(pf.vinetaColor, ctx) : null;
      if (color) v.color = color;
      if (pf.vinetaTamano > 0 && pf.vinetaTamano !== 100) v.tamano = pf.vinetaTamano;
      return v;
    },
    // Texto de una forma → { parrafos }
    convertirTexto(t, ctx, pp9) {
      const b = this.pd;
      const texto = t.texto || '';
      if ((this.textoTotal += texto.length) > MAX_TEXTO) throw new Danado('La presentación tiene demasiado texto y no se puede leer.');
      const { pfs, cfs } = this.leerEstiloTexto(t.estilo, texto.length);
      const parrafos = [];
      let ip = 0, ic = 0;
      const enlaces = t.enlaces.slice().sort((a, c) => a.ini - c.ini);
      const posCampos = [...t.campos.keys()].sort((x, y) => x - y);
      let ie = 0, ipc = 0;
      let ini = 0;
      let numPf = -1;
      while (ini <= texto.length) {
        let fin = texto.indexOf('\r', ini);
        const ultimo = fin < 0;
        if (ultimo) fin = texto.length;
        while (ip < pfs.length - 1 && pfs[ip].fin <= ini) ip++;
        while (ic < cfs.length - 1 && cfs[ic].fin <= ini) ic++;
        while (ie < enlaces.length && enlaces[ie].fin <= ini) ie++;
        while (ipc < posCampos.length && posCampos[ipc] < ini) ipc++;
        const runPf = pfs[ip] && pfs[ip].ini <= ini ? pfs[ip] : null;
        if (runPf) numPf = ip;
        const nivel = runPf ? runPf.nivel : 0;
        const base = this.estiloDe(ctx, t.tipoTexto, nivel);
        const pf = Object.assign({}, base.pf, t.base ? t.base.pf : {}, runPf ? runPf.pf : {});
        const parrafo = { hijos: [] };
        parrafo.alin = ALIN[pf.alin || 0] || 'l';
        if (nivel) parrafo.nivel = nivel;
        if (pf.izq !== undefined) parrafo.izq = emu(pf.izq);
        if (pf.sangria !== undefined) parrafo.primera = emu(pf.sangria - (pf.izq || 0));
        if (pf.interlineado > 0) parrafo.interlineado = { pct: pf.interlineado }; else if (pf.interlineado < 0) parrafo.interlineado = { pt: -pf.interlineado / 8 };
        const esp = (v) => (v > 0 ? { pct: v } : v < 0 ? -v / 8 : 0);
        if (pf.antes) parrafo.antes = esp(pf.antes); if (pf.despues) parrafo.despues = esp(pf.despues);
        // segmentos del párrafo
        const cortes = new Set([ini, fin]);
        for (let k = ic; k < cfs.length; k++) { if (cfs[k].fin <= ini) continue; if (cfs[k].ini >= fin) break; cortes.add(Math.max(cfs[k].ini, ini)); cortes.add(Math.min(cfs[k].fin, fin)); }
        for (let q = ie; q < enlaces.length && enlaces[q].ini < fin; q++) { const e = enlaces[q]; if (e.fin <= ini) continue; cortes.add(Math.max(e.ini, ini)); cortes.add(Math.min(e.fin, fin)); }
        for (let q = ipc; q < posCampos.length && posCampos[q] < fin; q++) { cortes.add(posCampos[q]); cortes.add(posCampos[q] + 1); }
        for (let k = ini; k < fin; k++) if (texto.charCodeAt(k) === 0x0B) { cortes.add(k); cortes.add(k + 1); }
        const puntos = [...cortes].sort((x, y) => x - y);
        let primerRun = null;
        for (let k = 0; k + 1 < puntos.length; k++) {
          const a = puntos[k], z = puntos[k + 1];
          if (a >= z) continue;
          while (ic < cfs.length - 1 && cfs[ic].fin <= a) ic++;
          const cfRun = cfs[ic] && cfs[ic].ini <= a && cfs[ic].fin > a ? cfs[ic].cf : {};
          const cf = Object.assign({}, base.cf, t.base ? t.base.cf : {}, cfRun);
          const r = {};
          const fuente = this.nombreFuente(cf.fuente !== undefined ? cf.fuente : cf.fuenteAnsi);
          if (fuente) r.fuente = fuente;
          if (cf.tamano) r.tamano = cf.tamano;
          if (cf.negrita) r.negrita = true; if (cf.cursiva) r.cursiva = true; if (cf.subrayado) r.subrayado = true;
          const col = this.colorTexto(cf.color, ctx); if (col) r.color = col;
          if (cf.posicion > 0) r.sup = true; else if (cf.posicion < 0) r.sub = true;
          if (!primerRun) primerRun = { cf, r };
          const trozo = texto.slice(a, z);
          const hijo = {};
          if (t.campos.has(a) && z === a + 1) { hijo.campo = t.campos.get(a); hijo.texto = ''; }
          else if (trozo === '\u000B') { hijo.salto = true; }
          else hijo.texto = trozo.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
          if (Object.keys(r).length) hijo.r = r;
          let en = null;
          for (let q = ie; q < enlaces.length && enlaces[q].ini <= a; q++) if (enlaces[q].fin >= z) { en = enlaces[q]; break; }
          if (en && hijo.texto) { const d = this.destinoEnlace(en.destino, ctx); if (d) hijo.enlace = d; }
          if (hijo.texto === '' && !hijo.campo && !hijo.salto) continue;
          parrafo.hijos.push(hijo);
        }
        const v = this.vinetaDe(pf, primerRun ? primerRun.cf : base.cf, ctx);
        const num = pp9 && numPf >= 0 && pp9[numPf] && pp9[numPf].auto !== 0 && pp9[numPf].esquema !== undefined ? pp9[numPf] : null;
        if (num && num.esquema !== undefined) parrafo.vineta = { num: NUMERACION[num.esquema] || 'arabicPeriod', inicio: num.inicio > 1 ? num.inicio : undefined };
        else if (v) parrafo.vineta = v;
        if (primerRun && primerRun.r.tamano) parrafo.r = { tamano: primerRun.r.tamano };
        parrafos.push(parrafo);
        this.contarUnidad(1 + parrafo.hijos.length);
        if (ultimo || parrafos.length > 1000000) break;
        ini = fin + 1;
      }
      return { parrafos };
    },
    destinoEnlace(d, ctx) {
      if (typeof d === 'string') return d;
      if (d && d.slideId) { const i = this.listaDiap.findIndex((x) => x.slideId === d.slideId); return i >= 0 ? { diapositiva: i + 1 } : null; }
      if (d && d.relativa && ctx.numeroDiapositiva) return { diapositiva: Math.max(1, ctx.numeroDiapositiva + (d.relativa === 'siguiente' ? 1 : -1)) };
      return null;
    },

    // ───── Formas ─────
    cajaDe(sp, base, u) {
      const a = sp.ancla || (base && base.ancla);
      if (!a) return null;
      const w = a.der - a.izq, h = a.aba - a.arr;
      const e = (v) => Math.round(v * u);
      return { x: e(a.izq), y: e(a.arr), w: e(Math.max(0, w)), h: e(Math.max(0, h)) };
    },
    // u: EMU por unidad de las coordenadas de la forma (las maestras fuera de un grupo; dentro de un grupo se dejan tal cual: chOff y chExt las escalan)
    convertirForma(sp, ctx, u = EMU_POR_MAESTRA) {
      if (sp.flags & 0x4 || sp.flags & 0x8 || sp.flags & 0x400) return null;       // patriarca, borrada, fondo
      const ph = sp.colocador;
      const baseId = ph && COLOCADOR_PATRON[ph.id] !== undefined ? COLOCADOR_PATRON[ph.id] : ph && ph.id >= 1 && ph.id <= 10 ? ph.id : null;
      const baseSp = baseId !== null && ctx.colocadores ? ctx.colocadores.get(baseId) : null;
      const mapas = [baseSp && baseSp.props, sp.props];
      const caja = this.cajaDe(sp, baseSp, u);
      const noOcultar = this.booleano(mapas, 0x3BF, 1);
      if (noOcultar === true) return null;
      if (!caja) return null;
      const f = Object.assign({}, caja);
      const rot = this.fijo(this.prop(mapas, 0x04) || 0);
      if (rot) f.rot = ((rot % 360) + 360) % 360;
      if (sp.flags & 0x40) f.volteoH = true;
      if (sp.flags & 0x80) f.volteoV = true;

      if (sp.esGrupo) {
        const hijos = sp.hijos.map((h) => this.convertirForma(h, ctx, 1)).filter(Boolean);
        if (!hijos.length) return null;
        f.tipo = 'grupo'; f.hijos = hijos;
        if (sp.rectHijos) { f.hx = sp.rectHijos.izq; f.hy = sp.rectHijos.arr; f.hw = sp.rectHijos.der - sp.rectHijos.izq; f.hh = sp.rectHijos.aba - sp.rectHijos.arr; }
        return f;
      }
      const tipoSp = sp.tipoSp;
      const bid = this.prop(mapas, 0x104);
      const esImagen = bid && (tipoSp === 75 || tipoSp === 1 || tipoSp === 0) && !(this.booleano(mapas, 0x1BF, 4)) && !sp.texto && !sp.comp.has(0xC0);
      const linea = this.linea(mapas, ctx);
      // enlace de la forma
      if (sp.enlace) { const d = this.destinoEnlace(sp.enlace, ctx); if (d) f.enlace = d; }
      if (esImagen) {
        const im = this.imagen(bid);
        if (!im) return null;
        f.tipo = 'imagen'; f.imagen = im;
        const rc = { arr: this.fijo(this.prop(mapas, 0x100) || 0), aba: this.fijo(this.prop(mapas, 0x101) || 0), izq: this.fijo(this.prop(mapas, 0x102) || 0), der: this.fijo(this.prop(mapas, 0x103) || 0) };
        if (rc.arr || rc.aba || rc.izq || rc.der) f.recorte = rc;
        if (linea) f.linea = linea;
        return f;
      }
      if (tipoSp === 20 || tipoSp === 32 || (sp.flags & 0x100 && tipoSp >= 32 && tipoSp <= 40)) {
        f.tipo = 'linea';
        if (linea) f.linea = linea; else f.linea = { color: '000000', ancho: 9525 };
        const s = this.sombra(mapas, ctx); if (s) f.sombra = s;
        return f;
      }
      f.tipo = 'forma';
      f.geom = MSOSPT[tipoSp] || 'rect';
      if (tipoSp !== 0 && !MSOSPT[tipoSp]) this.aviso('formas');
      if (tipoSp === 0 && sp.comp.has(0x145)) { const libre = this.formaLibre(sp, mapas); if (libre) f.geom = libre; else this.aviso('formas'); }
      const aj = this.ajustes(sp, mapas, f.geom);
      if (aj) f.ajustes = aj;
      f.relleno = this.relleno(mapas, ctx, false, sp);
      f.linea = linea;
      const sm = this.sombra(mapas, ctx); if (sm) f.sombra = sm;
      // texto
      let tx = sp.texto;
      if (!tx && sp.refTexto != null && ctx.textosDiap && ctx.textosDiap[sp.refTexto]) {
        const e = ctx.textosDiap[sp.refTexto];
        tx = { tipoTexto: e.tipoTexto, texto: e.texto || '', estilo: e.estilo || null, regla: null, campos: new Map(), enlaces: [] };
      }
      // WordArt: el texto va en la propiedad gtextUNICODE y la forma lo dibuja con su relleno; se lleva como texto normal con el color del relleno
      let base = null;
      if (!tx && sp.comp.has(0xC0)) {
        const d = sp.comp.get(0xC0);
        const s = textoUtf16(d, 0, d.length).replace(/\u0000[\s\S]*$/, '');
        if (s.trim() !== '') {
          const lineas = Math.max(1, s.split('\r').length);
          const nombre = sp.comp.has(0xC5) ? textoUtf16(sp.comp.get(0xC5), 0, sp.comp.get(0xC5).length).replace(/\u0000[\s\S]*$/, '') : '';
          const cf = { tamano: acotar(Math.round(f.h / 12700 * 0.7 / lineas), 8, 200) };
          const col = f.relleno && f.relleno.color; if (col) cf.color = { rgb: col };
          if (nombre) { let i = this.fuentes.indexOf(nombre); if (i < 0 && this.fuentes.length < 4096) { this.fuentes.push(nombre); i = this.fuentes.length - 1; } if (i >= 0) cf.fuente = i; }
          tx = { tipoTexto: 4, texto: s, estilo: null, regla: null, campos: new Map(), enlaces: [] };
          base = { pf: { alin: 1 }, cf };
          f.relleno = null; f.linea = null;
          this.aviso('wordart');
        }
      }
      if (tx && base) tx.base = base;
      if (tx && tx.texto && tx.texto.replace(/[\r\u000B\s]/g, '') !== '') {
        const t = this.convertirTexto(tx, ctx, sp.pf9);
        const margen = (id, def) => { const v = this.prop(mapas, id); return v === undefined ? def : v; };
        const ancla = this.prop(mapas, 0x87);
        const texto = { parrafos: t.parrafos, margenes: { i: margen(0x81, 91440), s: margen(0x82, 45720), d: margen(0x83, 91440), a: margen(0x84, 45720) } };
        if (ancla !== undefined && ANCLA_TEXTO[ancla]) texto.ancla = ANCLA_TEXTO[ancla]; else if (ph) texto.ancla = 't';
        const ajuste = this.prop(mapas, 0x85); if (ajuste === 2) texto.ajuste = false;
        const flujo = this.prop(mapas, 0x88); if (flujo === 1 || flujo === 3) texto.vertical = true;
        if (this.booleano(mapas, 0xBF, 1)) texto.autoajuste = 'forma';
        f.texto = texto;
      } else if (!f.relleno && !f.linea) return null;   // marcador de posición o cuadro de texto vacío, sin relleno ni línea
      return f;
    },
    // Forma libre: vértices (pVertices) y tipo de cada trazo (pSegmentInfo) → rutas del contrato
    formaLibre(sp, mapas) {
      const v = sp.comp.get(0x145), s = sp.comp.get(0x146);
      if (!v || !s || v.length < 6 || s.length < 6) return null;
      const nv = Math.min(u16(v, 0), 20000), cb = u16(v, 4);
      const tam = cb === 0xFFF0 ? 4 : 8;
      if (nv * tam + 6 > v.length) return null;
      const pt = (k) => (tam === 4 ? [i16(v, 6 + k * 4), i16(v, 8 + k * 4)] : [i32(v, 6 + k * 8), i32(v, 10 + k * 8)]);
      const ns = Math.min(u16(s, 0), 20000);
      if (ns * 2 + 6 > s.length) return null;
      const izq = this.prop(mapas, 0x140) | 0, arr = this.prop(mapas, 0x141) | 0;
      const der = this.prop(mapas, 0x142) === undefined ? 21600 : this.prop(mapas, 0x142) | 0, aba = this.prop(mapas, 0x143) === undefined ? 21600 : this.prop(mapas, 0x143) | 0;
      const w = der - izq, h = aba - arr;
      if (w <= 0 || h <= 0) return null;
      const cmds = [];
      let k = 0;
      const p = () => { if (k >= nv) throw new Error('fin'); const q = pt(k++); return [q[0] - izq, q[1] - arr]; };
      try {
        for (let i = 0; i < ns; i++) {
          const x = u16(s, 6 + i * 2), tipo = x >> 13, cuenta = x & 0x1FFF;
          if (tipo === 2) { const a = p(); cmds.push(['M', a[0], a[1]]); }
          else if (tipo === 0) { for (let j = 0; j < cuenta; j++) { const a = p(); cmds.push(['L', a[0], a[1]]); } }
          else if (tipo === 1) { for (let j = 0; j < cuenta; j++) { const a = p(), b = p(), c = p(); cmds.push(['C', a[0], a[1], b[0], b[1], c[0], c[1]]); } }
          else if (tipo === 3) cmds.push(['Z']);
          else if (tipo === 4) break;
          else return null;     // escapes (arcos, rellenos especiales): no se llevan
        }
      } catch (e) { return null; }
      this.contarUnidad(cmds.length);
      return cmds.length ? { rutas: [{ w, h, cmds }] } : null;
    },
    // Los ajustes de OfficeArt (adjustValue) a los de ECMA para las formas más comunes
    ajustes(sp, mapas, geom) {
      const a1 = this.prop(mapas, 0x147);
      if (a1 === undefined) return null;
      if (geom === 'roundRect') return [Math.round(acotar(a1, 0, 10800) * 50000 / 10800)];
      return null;
    },
  });

  // ───────── Patrones, diapositivas y notas ─────────
  const LEN_ESQUEMA = 32;
  Object.assign(Lector.prototype, {
    leerEsquema(h) {
      const b = this.pd;
      if (h.fin - h.ini < LEN_ESQUEMA) return null;
      const e = [];
      for (let i = 0; i < 8; i++) e.push(rgbHex(u8(b, h.ini + i * 4), u8(b, h.ini + i * 4 + 1), u8(b, h.ini + i * 4 + 2)));
      return e;
    },
    fondoDe(sp, ctx) {
      if (!sp) return null;
      const r = this.relleno([sp.props], ctx, true);
      if (!r) return null;
      if (r.degradado) return { degradado: r.degradado };
      if (r.imagen) return { imagen: r.imagen };
      return r.color ? { color: r.color } : null;
    },
    leerMaestro(entrada) {
      const rec = this.contenedor(entrada.persist);
      const m = { esquema: SIN_ESQUEMA.slice(), estilos: {}, colocadores: new Map(), formas: [], fondo: null };
      if (!rec || (rec.tipo !== RT.MainMaster && rec.tipo !== RT.Slide)) return m;
      let dibujo = null;
      let esquemaPuesto = false;
      for (const h of this.hijos(rec)) {
        if (h.tipo === RT.ColorSchemeAtom && !esquemaPuesto) { const e = this.leerEsquema(h); if (e) { m.esquema = e; esquemaPuesto = true; } }
        else if (h.tipo === RT.TextMasterStyleAtom) m.estilos[h.inst] = this.leerEstiloMaestro(h);
        else if (h.tipo === RT.PPDrawing) dibujo = this.leerDibujo(h);
      }
      const ctx = { esquema: m.esquema, estilos: m.estilos, colocadores: null };
      if (dibujo) {
        m.fondo = this.fondoDe(dibujo.fondo, ctx);
        for (const sp of dibujo.formas) {
          if (sp.colocador) { m.colocadores.set(sp.colocador.id, sp); continue; }
          const f = this.convertirForma(sp, ctx);
          if (f) m.formas.push(f);
        }
      }
      return m;
    },
    leerDiapositiva(entrada, indice, maestros) {
      const rec = this.contenedor(entrada.persist);
      const d = { formas: [] };
      if (!rec || rec.tipo !== RT.Slide) return { d, maestro: 0 };
      const b = this.pd;
      let dibujo = null, esquemaPropio = null, flags = 7, idMaestro = 0, idNotas = 0, pies = null;
      for (const h of this.hijos(rec)) {
        if (h.tipo === RT.SlideAtom) { idMaestro = u32(b, h.ini + 12); idNotas = u32(b, h.ini + 16); flags = u16(b, h.ini + 20); }
        else if (h.tipo === RT.ColorSchemeAtom && h.inst === 1) esquemaPropio = this.leerEsquema(h);
        else if (h.tipo === RT.PPDrawing) dibujo = this.leerDibujo(h);
        else if (h.tipo === RT.SlideShowSlideInfoAtom && h.fin - h.ini >= 12 && (u16(b, h.ini + 10) & 4)) d.oculta = true;
        else if (h.tipo === 0x0FD9) pies = this.leerPies(h);
      }
      let im = this.maestras.findIndex((x) => x.slideId === idMaestro);
      if (im < 0) im = 0;
      const m = maestros[im] || { esquema: SIN_ESQUEMA.slice(), estilos: {}, colocadores: new Map() };
      const ctx = { esquema: esquemaPropio && !(flags & 2) ? esquemaPropio : m.esquema, estilos: m.estilos, colocadores: m.colocadores, textosDiap: entrada.textos, numeroDiapositiva: indice + 1 };
      if (!(flags & 1)) d.ocultarPatron = true;
      if (dibujo) {
        const fondo = flags & 4 ? null : this.fondoDe(dibujo.fondo, ctx);     // con el bit 2 puesto el fondo es el del patrón, aunque la diapositiva traiga uno
        if (fondo) d.fondo = fondo;
        for (const sp of dibujo.formas) { const f = this.convertirForma(sp, ctx); if (f) d.formas.push(f); }
        if (pies) this.agregarPies(pies, d, ctx, dibujo.formas);
      }
      const notas = idNotas ? this.leerNotas(idNotas) : null;
      if (notas) d.notas = notas;
      return { d, maestro: im };
    },
    // Encabezados y pies de la diapositiva (HeadersFooters): lo que PowerPoint dibuja con los marcadores de fecha, número y pie del patrón
    leerPies(h) {
      const b = this.pd;
      const p = { flags: 0, fecha: '', encabezado: '', pie: '' };
      for (const x of this.hijos(h)) {
        if (x.tipo === 0x0FDA) p.flags = u16(b, x.ini + 2);
        else if (x.tipo === RT.CString) { const s = textoUtf16(b, x.ini, x.fin).replace(/\u0000.*$/s, ''); if (x.inst === 0) p.fecha = s; else if (x.inst === 1) p.encabezado = s; else if (x.inst === 2) p.pie = s; }
      }
      return p;
    },
    agregarPies(p, d, ctx, formasSlide) {
      const hay = new Set(formasSlide.filter((s) => s.colocador).map((s) => s.colocador.id));
      const base = (tx) => { if (!tx || !tx.estilo) return null; const e = this.leerEstiloTexto(tx.estilo, (tx.texto || '').length); return { pf: e.pfs[0] ? e.pfs[0].pf : {}, cf: e.cfs[0] ? e.cfs[0].cf : {} }; };
      const poner = (id, activo, texto) => {
        const m = ctx.colocadores.get(id);
        if (!activo || !m || hay.has(id) || !m.texto) return;
        const sp = Object.assign({}, m, { colocador: null, pf9: null, refTexto: null, flags: m.flags & ~0x20 });
        sp.texto = texto === null ? m.texto : { tipoTexto: m.texto.tipoTexto, texto, estilo: null, regla: null, campos: new Map(), enlaces: [], base: base(m.texto) };
        const f = this.convertirForma(sp, ctx);
        if (f) d.formas.splice(nuevos++, 0, f);     // PowerPoint los pone primero (debajo del resto)
      };
      let nuevos = 0;
      poner(7, (p.flags & 0x7) && (p.flags & 0x2 || p.fecha), p.flags & 0x2 || !p.fecha ? null : p.fecha);
      poner(9, (p.flags & 0x20) && p.pie, p.pie);
      poner(8, p.flags & 0x8, null);
      poner(10, (p.flags & 0x10) && p.encabezado, p.encabezado);
    },
    leerNotas(idNotas) {
      const entrada = this.listaNotas.get(idNotas);
      if (!entrada) return null;
      const rec = this.contenedor(entrada.persist);
      if (!rec || rec.tipo !== RT.Notes) return null;
      let dibujo = null;
      for (const h of this.hijos(rec)) if (h.tipo === RT.PPDrawing) dibujo = this.leerDibujo(h);
      if (!dibujo) return null;
      const buscar = (lista) => {
        for (const sp of lista) {
          let tx = sp.texto;
          if (!tx && sp.refTexto != null && entrada.textos[sp.refTexto]) { const e = entrada.textos[sp.refTexto]; tx = { tipoTexto: e.tipoTexto, texto: e.texto || '' }; }
          if (tx && tx.tipoTexto === 2 && tx.texto) return tx.texto;
          if (sp.hijos) { const r = buscar(sp.hijos); if (r) return r; }
        }
        return null;
      };
      let texto = buscar(dibujo.formas);
      if (!texto) { const u = entrada.textos.find((e) => e.tipoTexto === 2 && e.texto); texto = u ? u.texto : null; }
      if (!texto) return null;
      if ((this.textoTotal += texto.length) > MAX_TEXTO) throw new Danado('La presentación tiene demasiado texto y no se puede leer.');
      const partes = texto.replace(/\u000B/g, '\n').split('\r').map((s) => s.replace(/[\u0000-\u0008\u000C\u000E-\u001F]/g, ''));
      while (partes.length && partes[partes.length - 1] === '') partes.pop();
      return partes.length ? partes.map((s) => ({ hijos: s ? [{ texto: s }] : [] })) : null;
    },
    contrato() {
      this.leerPersistencia();
      this.leerDocumento();
      if (!this.listaDiap.length) throw new Danado('La presentación no tiene diapositivas.');
      if (this.listaDiap.length > MAX_DIAPOSITIVAS) throw new Danado('La presentación tiene demasiadas diapositivas y no se puede leer.');
      // lo que varias entradas comparten (un mismo contenedor señalado muchas veces) se lee una sola vez
      const cacheM = new Map();
      const maestros = this.maestras.map((e) => { if (!cacheM.has(e.persist)) cacheM.set(e.persist, this.leerMaestro(e)); return cacheM.get(e.persist); });
      const cacheD = new Map();
      const usados = new Set();
      const diapositivas = this.listaDiap.map((e, i) => {
        const clave = e.textos.length ? null : e.persist;
        let r = clave !== null ? cacheD.get(clave) : null;
        if (!r) { r = this.leerDiapositiva(e, i, maestros); if (clave !== null) cacheD.set(clave, r); }
        usados.add(r.maestro);
        return Object.assign({}, r.d, { maestro: r.maestro });
      });
      // solo se escriben los patrones que alguna diapositiva usa
      const orden = [...usados].sort((x, y) => x - y);
      const nuevo = new Map(orden.map((m, k) => [m, k]));
      diapositivas.forEach((d) => { d.maestro = nuevo.get(d.maestro); });
      return { ancho: this.ancho, alto: this.alto, maestros: orden.map((m) => ({ fondo: maestros[m] ? maestros[m].fondo : null, formas: maestros[m] ? maestros[m].formas : [] })), diapositivas };
    },
  });

  // DIB (mapa de bits sin la cabecera de archivo) → BMP
  function dibABmp(d) {
    if (d.length < 40) return null;
    const tamCab = u32(d, 0), bits = u16(d, 14), usados = u32(d, 32);
    const paleta = bits <= 8 ? (usados || (1 << bits)) * 4 : 0;
    const desp = 14 + tamCab + paleta;
    const bmp = new Uint8Array(14 + d.length);
    bmp[0] = 0x42; bmp[1] = 0x4D;
    const total = bmp.length;
    bmp[2] = total & 255; bmp[3] = (total >>> 8) & 255; bmp[4] = (total >>> 16) & 255; bmp[5] = (total >>> 24) & 255;
    bmp[10] = desp & 255; bmp[11] = (desp >>> 8) & 255; bmp[12] = (desp >>> 16) & 255; bmp[13] = (desp >>> 24) & 255;
    bmp.set(d, 14);
    return bmp;
  }
  // zlib → bytes (la compresión de los metarchivos de OfficeArt), con tope
  async function inflarZlib(datos, tope) {
    const flujo = new Blob([datos]).stream().pipeThrough(new DecompressionStream('deflate')).getReader();
    const partes = [];
    let total = 0;
    for (;;) {
      const { done, value } = await flujo.read();
      if (done) break;
      total += value.length;
      if (total > tope) { try { await flujo.cancel(); } catch (e) { /* ya está */ } throw new Error('demasiado grande'); }
      partes.push(value);
    }
    const salida = new Uint8Array(total);
    let p = 0;
    for (const x of partes) { salida.set(x, p); p += x.length; }
    return salida;
  }
  // Recorre el contrato y deja las imágenes listas para el .pptx: metarchivos descomprimidos, DIB como BMP, formatos que no se pueden llevar fuera
  async function prepararImagenes(contrato, lector) {
    const vistas = new Map();
    const arreglar = async (im) => {
      if (!im) return null;
      if (vistas.has(im)) return vistas.get(im);
      let r = im;
      if (im.tipo === 'dib') { const bmp = dibABmp(im.datos); r = bmp ? { datos: bmp, tipo: 'bmp' } : null; }
      else if (im.tipo === 'tiff') r = null;
      else if (im.metarchivo) {
        try { r = { datos: im.comprimido ? await inflarZlib(im.datos, Math.min(im.sinComprimir + 1024, 64 * 1024 * 1024)) : im.datos, tipo: im.tipo }; } catch (e) { r = null; }
      }
      if (!r) lector.aviso('imagen-formato');
      vistas.set(im, r);
      return r;
    };
    const f = async (formas) => {
      for (let i = formas.length - 1; i >= 0; i--) {
        const x = formas[i];
        if (x.hijos) await f(x.hijos);
        if (x.tipo === 'imagen') { x.imagen = await arreglar(x.imagen); if (!x.imagen) formas.splice(i, 1); }
        if (x.relleno && x.relleno.imagen) { const r = await arreglar(x.relleno.imagen); if (r) x.relleno = { imagen: r }; else x.relleno = undefined; }
      }
    };
    const fondo = async (o) => { if (o.fondo && o.fondo.imagen) { const r = await arreglar(o.fondo.imagen); if (r) o.fondo = { imagen: r }; else delete o.fondo; } };
    for (const m of contrato.maestros) { await f(m.formas); await fondo(m); }
    for (const d of contrato.diapositivas) { await f(d.formas); await fondo(d); }
  }

  const esCfb = (b) => b.length >= 512 && [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1].every((x, i) => b[i] === x);
  function parece(bytes) {
    if (!esCfb(bytes)) return false;
    try { const cfb = CfbPropio.leer(bytes); return cfb.entradas.some((e) => e.nombre === 'PowerPoint Document' && e.tipo === 2); } catch (e) { return false; }
  }
  async function leer(bytes) {
    const l = new Lector(bytes);
    const contrato = l.contrato();
    await prepararImagenes(contrato, l);
    return { contrato, mensajes: [...l.mensajes] };
  }
  async function aPptx(bytes) {
    const { contrato, mensajes } = await leer(bytes);
    const r = await PptxDesdeContrato.generar(contrato);
    return { bytes: r.bytes, mensajes: [...new Set([...mensajes, ...r.mensajes])] };
  }

  globalThis.PptLector = { parece, leer, aPptx, _Lector: (b) => new Lector(b) };
})();

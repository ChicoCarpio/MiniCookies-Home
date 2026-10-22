/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Modelo de formato de un .docx para el maquetador de Word (docx-maquetador.js):
// valores por defecto del documento, estilos con herencia (basedOn), tema
// (fuentes y colores), numeración (listas) y ajustes. Convierte el XML de Word a
// objetos simples con unidades de pantalla: puntos (pt) para tamaños, sangrías y
// espacios; "sin especificar" = la propiedad no existe en el objeto.
//
// Requiere DocxBase (docx-base.js), ZipPropio y XmlLigero. No depende del
// navegador: corre igual en Node.
//
// Uso:
//   const paquete = await DocxBase.abrirPaquete(datos);
//   const m = await DocxEstilos.cargar(paquete);
//   const { p, r } = m.estilos.parrafo(idEstilo, medio?);  // propiedades efectivas del estilo (con docDefaults y, si hay, el formato intermedio de una tabla)
//   const r2 = m.mezclar(r, m.leerRPr(nodoRPr));       // el formato directo gana
//   m.color(r2.color) → '#rrggbb' | null (automático)   m.fuente(r2.fuentes) → 'Calibri'
//   const num = m.numeracion.numerador(); num.siguiente(numId, ilvl) → { texto, nivel }
// Límites conocidos: las propiedades de alternancia (negrita, cursiva) del estilo
// de carácter sobre el de párrafo no se invierten como en Word: la última gana.

(function () {
  'use strict';
  const { arbol, hijo, lista, buscar, porTipo } = DocxBase;

  // ───────── Utilidades ─────────
  const num = (v) => { const n = parseFloat(v); return Number.isFinite(n) ? n : undefined; };
  const tw = (v) => { const n = num(v); return n === undefined ? undefined : n / 20; };           // veinteavos de punto → pt
  const ep = (v) => { const n = num(v); return n === undefined ? undefined : n / 8; };            // octavos de punto → pt
  const val = (e) => (e ? e.a.val : undefined);
  const encendido = (e) => { // <w:b/> = sí; <w:b w:val="0|false|off"/> = no
    if (!e) return undefined;
    const v = e.a.val;
    return !(v === '0' || v === 'false' || v === 'off');
  };

  const esObjeto = (x) => x && typeof x === 'object' && !Array.isArray(x);
  // Combina dos conjuntos de propiedades: lo definido en `sobre` gana; los objetos
  // (sangría, espacios, bordes) se combinan campo por campo; las listas (tabs) se reemplazan salvo `tabs`.
  function mezclar(base, sobre) {
    if (!sobre) return base ? { ...base } : {};
    const sal = { ...(base || {}) };
    for (const k of Object.keys(sobre)) {
      const v = sobre[k];
      if (v === undefined) continue;
      if (k === 'tabs' && Array.isArray(v)) sal.tabs = mezclarTabs(sal.tabs, v);
      else if (esObjeto(v) && esObjeto(sal[k])) sal[k] = mezclar(sal[k], v);
      else sal[k] = v;
    }
    return sal;
  }
  function mezclarTabs(base, nuevos) {
    const sal = (base || []).slice();
    for (const t of nuevos) {
      const i = sal.findIndex((x) => x.pos === t.pos);
      if (t.tipo === 'clear') { if (i >= 0) sal.splice(i, 1); } else if (i >= 0) sal[i] = t; else sal.push(t);
    }
    return sal.sort((a, b) => a.pos - b.pos);
  }

  // ───────── Propiedades de carácter, párrafo, tabla ─────────
  function colorDe(e) {
    if (!e) return undefined;
    const a = e.a;
    return { val: a.val, tema: a.themeColor, tinte: a.themeTint, sombra: a.themeShade, relleno: a.fill };
  }
  function borde(e) {
    if (!e) return undefined;
    const a = e.a;
    return { tipo: a.val, ancho: ep(a.sz), espacio: num(a.space), color: colorDe({ a: { val: a.color, themeColor: a.themeColor, themeTint: a.themeTint, themeShade: a.themeShade } }) };
  }
  function bordes(e, nombres) {
    if (!e) return undefined;
    const o = {};
    for (const n of nombres) {
      const b = borde(hijo(e, n));
      if (b) o[n === 'start' ? 'left' : n === 'end' ? 'right' : n] = b;
    }
    return o;
  }
  function sombreado(e) {
    if (!e) return undefined;
    return { tipo: e.a.val, relleno: e.a.fill, color: e.a.color, tema: e.a.themeFill };
  }
  function limpiar(o) { // quita claves undefined
    for (const k of Object.keys(o)) if (o[k] === undefined) delete o[k];
    return o;
  }

  function leerRPr(e) {
    if (!e) return {};
    const f = hijo(e, 'rFonts');
    const r = {
      estilo: val(hijo(e, 'rStyle')),
      fuentes: f ? limpiar({
        ascii: f.a.ascii || (f.a.asciiTheme ? 'tema:' + f.a.asciiTheme : undefined),
        hAnsi: f.a.hAnsi || (f.a.hAnsiTheme ? 'tema:' + f.a.hAnsiTheme : undefined),
        eastAsia: f.a.eastAsia || (f.a.eastAsiaTheme ? 'tema:' + f.a.eastAsiaTheme : undefined),
        cs: f.a.cs || (f.a.cstheme ? 'tema:' + f.a.cstheme : undefined),
      }) : undefined,
      b: encendido(hijo(e, 'b')), i: encendido(hijo(e, 'i')),
      tam: tw1(hijo(e, 'sz')), tamCs: tw1(hijo(e, 'szCs')),
      u: val(hijo(e, 'u')),
      tachado: encendido(hijo(e, 'strike')), tachadoDoble: encendido(hijo(e, 'dstrike')),
      mayusculas: encendido(hijo(e, 'caps')), versalitas: encendido(hijo(e, 'smallCaps')),
      oculto: encendido(hijo(e, 'vanish')),
      alineacionVertical: val(hijo(e, 'vertAlign')),
      color: colorDe(hijo(e, 'color')),
      resaltado: val(hijo(e, 'highlight')),
      sombreado: sombreado(hijo(e, 'shd')),
      espaciado: tw(val(hijo(e, 'spacing'))),
      escala: num(val(hijo(e, 'w'))),
      posicion: num(val(hijo(e, 'position'))) === undefined ? undefined : num(val(hijo(e, 'position'))) / 2,
      borde: borde(hijo(e, 'bdr')),
      idioma: hijo(e, 'lang') ? hijo(e, 'lang').a.val : undefined,
    };
    return limpiar(r);
  }
  function tw1(e) { const n = num(val(e)); return n === undefined ? undefined : n / 2; } // medios puntos → pt

  function leerPPr(e) {
    if (!e) return {};
    const sp = hijo(e, 'spacing');
    const ind = hijo(e, 'ind');
    const np = hijo(e, 'numPr');
    const tabs = hijo(e, 'tabs');
    const marco = hijo(e, 'framePr');
    const p = {
      estilo: val(hijo(e, 'pStyle')),
      jc: val(hijo(e, 'jc')),
      sangria: ind ? limpiar({
        izq: tw(ind.a.left !== undefined ? ind.a.left : ind.a.start), der: tw(ind.a.right !== undefined ? ind.a.right : ind.a.end),
        primera: tw(ind.a.firstLine), colgante: tw(ind.a.hanging),
      }) : undefined,
      espacio: sp ? limpiar({
        antes: tw(sp.a.before), despues: tw(sp.a.after),
        antesAuto: sp.a.beforeAutospacing === undefined ? undefined : sp.a.beforeAutospacing === '1' || sp.a.beforeAutospacing === 'true',
        despuesAuto: sp.a.afterAutospacing === undefined ? undefined : sp.a.afterAutospacing === '1' || sp.a.afterAutospacing === 'true',
        linea: num(sp.a.line), regla: sp.a.lineRule || (sp.a.line !== undefined ? 'auto' : undefined),
      }) : undefined,
      bordes: bordes(hijo(e, 'pBdr'), ['top', 'left', 'bottom', 'right', 'between', 'bar', 'start', 'end']),
      sombreado: sombreado(hijo(e, 'shd')),
      tabs: tabs ? lista(tabs, 'tab').map((t) => ({ tipo: t.a.val, pos: tw(t.a.pos), guia: t.a.leader })) : undefined,
      mantenerConSiguiente: encendido(hijo(e, 'keepNext')),
      mantenerLineas: encendido(hijo(e, 'keepLines')),
      saltoAntes: encendido(hijo(e, 'pageBreakBefore')),
      controlViudas: encendido(hijo(e, 'widowControl')),
      sinEspacioEntreIguales: encendido(hijo(e, 'contextualSpacing')),
      derechaAIzquierda: encendido(hijo(e, 'bidi')),
      nivelEsquema: num(val(hijo(e, 'outlineLvl'))),
      numeracion: np ? limpiar({ numId: val(hijo(np, 'numId')), nivel: val(hijo(np, 'ilvl')) }) : undefined,
      marco: marco ? { ...marco.a } : undefined,
      marca: hijo(e, 'rPr') ? leerRPr(hijo(e, 'rPr')) : undefined,
    };
    return limpiar(p);
  }

  const LADOS_TABLA = ['top', 'left', 'bottom', 'right', 'start', 'end', 'insideH', 'insideV'];
  const LADOS_CELDA = ['top', 'left', 'bottom', 'right', 'start', 'end', 'insideH', 'insideV', 'tl2br', 'tr2bl'];
  function ancho(e) { return e ? { tipo: e.a.type || 'dxa', valor: num(e.a.w) } : undefined; }
  function margenes(e) {
    if (!e) return undefined;
    const o = {};
    for (const [n, k] of [['top', 'top'], ['left', 'left'], ['start', 'left'], ['bottom', 'bottom'], ['right', 'right'], ['end', 'right']]) {
      const x = hijo(e, n);
      if (x) o[k] = tw(x.a.w);
    }
    return o;
  }

  function leerTblPr(e) {
    if (!e) return {};
    const look = hijo(e, 'tblLook');
    return limpiar({
      estilo: val(hijo(e, 'tblStyle')),
      ancho: ancho(hijo(e, 'tblW')),
      jc: val(hijo(e, 'jc')),
      sangria: ancho(hijo(e, 'tblInd')),
      bordes: bordes(hijo(e, 'tblBorders'), LADOS_TABLA),
      disposicion: hijo(e, 'tblLayout') ? hijo(e, 'tblLayout').a.type : undefined,
      margenes: margenes(hijo(e, 'tblCellMar')),
      espaciadoCeldas: hijo(e, 'tblCellSpacing') ? tw(hijo(e, 'tblCellSpacing').a.w) : undefined,
      flotante: hijo(e, 'tblpPr') ? { ...hijo(e, 'tblpPr').a } : undefined,
      bandasFila: num(val(hijo(e, 'tblStyleRowBandSize'))), bandasColumna: num(val(hijo(e, 'tblStyleColBandSize'))),
      aspecto: look ? aspectoTabla(look) : undefined,
    });
  }
  // w:tblLook: o atributos sueltos (firstRow="1"…) o un valor hexadecimal (w:val="04A0").
  function aspectoTabla(look) {
    const a = look.a;
    if (a.val !== undefined && a.firstRow === undefined) {
      const v = parseInt(a.val, 16) || 0;
      return { primeraFila: !!(v & 0x20), ultimaFila: !!(v & 0x40), primeraColumna: !!(v & 0x80), ultimaColumna: !!(v & 0x100), sinBandasH: !!(v & 0x200), sinBandasV: !!(v & 0x400) };
    }
    const si = (x) => x === '1' || x === 'true';
    return { primeraFila: si(a.firstRow), ultimaFila: si(a.lastRow), primeraColumna: si(a.firstColumn), ultimaColumna: si(a.lastColumn), sinBandasH: si(a.noHBand), sinBandasV: si(a.noVBand) };
  }
  function leerTcPr(e) {
    if (!e) return {};
    const vm = hijo(e, 'vMerge');
    return limpiar({
      ancho: ancho(hijo(e, 'tcW')),
      columnas: num(val(hijo(e, 'gridSpan'))),
      combinadaVertical: vm ? (vm.a.val === 'restart' ? 'restart' : 'continue') : undefined,
      bordes: bordes(hijo(e, 'tcBorders'), LADOS_CELDA),
      sombreado: sombreado(hijo(e, 'shd')),
      alineacionVertical: val(hijo(e, 'vAlign')),
      margenes: margenes(hijo(e, 'tcMar')),
      sinAjuste: encendido(hijo(e, 'noWrap')),
      direccion: val(hijo(e, 'textDirection')),
    });
  }
  function leerTrPr(e) {
    if (!e) return {};
    const h = hijo(e, 'trHeight');
    return limpiar({
      alto: h ? tw(h.a.val) : undefined, regla: h ? h.a.hRule || 'atLeast' : undefined,
      encabezado: encendido(hijo(e, 'tblHeader')), noPartir: encendido(hijo(e, 'cantSplit')),
    });
  }

  // ───────── Tema ─────────
  const COLORES_TEMA = {
    text1: 'dk1', dark1: 'dk1', background1: 'lt1', light1: 'lt1', text2: 'dk2', dark2: 'dk2', background2: 'lt2', light2: 'lt2',
    accent1: 'accent1', accent2: 'accent2', accent3: 'accent3', accent4: 'accent4', accent5: 'accent5', accent6: 'accent6',
    hyperlink: 'hlink', followedHyperlink: 'folHlink',
  };
  function leerTema(raiz) {
    const t = { colores: {}, fuentes: { mayor: {}, menor: {} } };
    if (!raiz) return t;
    const cs = buscar(raiz, 'clrScheme');
    if (cs) {
      for (const c of cs.h) {
        const x = c.h[0];
        if (x) t.colores[c.n] = (x.a.val || x.a.lastClr || '').toLowerCase();
      }
    }
    for (const [nombre, clave] of [['majorFont', 'mayor'], ['minorFont', 'menor']]) {
      const f = buscar(raiz, nombre);
      if (!f) continue;
      for (const [hijoNombre, k] of [['latin', 'latin'], ['ea', 'eastAsia'], ['cs', 'cs']]) {
        const x = hijo(f, hijoNombre);
        if (x && x.a.typeface) t.fuentes[clave][k] = x.a.typeface;
      }
    }
    return t;
  }
  const hex2 = (n) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0');
  function aplicarTinte(rgb, tinte, sombra) {
    let c = [0, 2, 4].map((i) => parseInt(rgb.slice(i, i + 2), 16));
    if (sombra !== undefined) { const s = parseInt(sombra, 16) / 255; c = c.map((v) => v * s); }
    if (tinte !== undefined) { const t = parseInt(tinte, 16) / 255; c = c.map((v) => v * t + 255 * (1 - t)); }
    return c.map(hex2).join('');
  }
  function colorHex(tema, c) {
    if (!c) return null;
    if (c.tema) {
      const base = tema.colores[COLORES_TEMA[c.tema] || c.tema];
      if (base && /^[0-9a-f]{6}$/.test(base)) return '#' + aplicarTinte(base, c.tinte, c.sombra);
    }
    if (c.val && /^[0-9a-fA-F]{6}$/.test(c.val)) return '#' + c.val.toLowerCase();
    return null; // "auto" o desconocido: lo decide quien dibuja (negro sobre blanco)
  }
  function nombreFuente(tema, fuentes, tipo = 'ascii') {
    if (!fuentes) return null;
    let n = fuentes[tipo] || fuentes.ascii || fuentes.hAnsi;
    if (n && n.startsWith('tema:')) {
      const t = n.slice(5);
      const grupo = /^major/i.test(t) ? tema.fuentes.mayor : tema.fuentes.menor;
      n = /EastAsia/i.test(t) ? grupo.eastAsia : /Bidi/i.test(t) ? grupo.cs : grupo.latin;
    }
    return n || null;
  }

  // ───────── Estilos ─────────
  const TIPOS_CONDICIONALES = ['wholeTable', 'firstRow', 'lastRow', 'firstCol', 'lastCol', 'band1Vert', 'band2Vert', 'band1Horz', 'band2Horz', 'neCell', 'nwCell', 'seCell', 'swCell'];

  function leerEstilos(raiz) {
    const def = { p: {}, r: {} };
    const estilos = new Map();
    const porDefecto = {};
    if (!raiz) return { def, estilos, porDefecto };
    const dd = hijo(raiz, 'docDefaults');
    if (dd) {
      def.r = leerRPr(hijo(hijo(dd, 'rPrDefault'), 'rPr'));
      def.p = leerPPr(hijo(hijo(dd, 'pPrDefault'), 'pPr'));
    }
    for (const s of lista(raiz, 'style')) {
      const id = s.a.styleId;
      if (!id) continue;
      const condicionales = {};
      for (const c of lista(s, 'tblStylePr')) {
        if (TIPOS_CONDICIONALES.includes(c.a.type)) {
          condicionales[c.a.type] = { p: leerPPr(hijo(c, 'pPr')), r: leerRPr(hijo(c, 'rPr')), tbl: leerTblPr(hijo(c, 'tblPr')), tc: leerTcPr(hijo(c, 'tcPr')), tr: leerTrPr(hijo(c, 'trPr')) };
        }
      }
      const e = {
        id, tipo: s.a.type || 'paragraph', nombre: val(hijo(s, 'name')) || '', basadoEn: val(hijo(s, 'basedOn')) || null,
        enlace: val(hijo(s, 'link')) || null, siguiente: val(hijo(s, 'next')) || null,
        p: leerPPr(hijo(s, 'pPr')), r: leerRPr(hijo(s, 'rPr')),
        tbl: leerTblPr(hijo(s, 'tblPr')), tc: leerTcPr(hijo(s, 'tcPr')), tr: leerTrPr(hijo(s, 'trPr')),
        condicionales,
      };
      estilos.set(id, e);
      if (s.a.default === '1' || s.a.default === 'true') porDefecto[e.tipo] = id;
    }
    return { def, estilos, porDefecto };
  }

  function crearEstilos(datos) {
    const { def, estilos, porDefecto } = datos;
    const cache = new Map();
    // Cadena de herencia, del más lejano al propio (con tope, por si hay ciclos).
    function cadena(id) {
      const c = [];
      const visto = new Set();
      for (let e = estilos.get(id); e && !visto.has(e.id) && c.length < 30; e = estilos.get(e.basadoEn)) { visto.add(e.id); c.unshift(e); }
      return c;
    }
    function resolver(id) {
      if (cache.has(id)) return cache.get(id);
      const r = { p: {}, r: {}, tbl: {}, tc: {}, tr: {}, condicionales: {} };
      for (const e of cadena(id)) {
        r.p = mezclar(r.p, e.p); r.r = mezclar(r.r, e.r); r.tbl = mezclar(r.tbl, e.tbl); r.tc = mezclar(r.tc, e.tc); r.tr = mezclar(r.tr, e.tr);
        for (const k of Object.keys(e.condicionales)) {
          const c = e.condicionales[k], a = r.condicionales[k] || { p: {}, r: {}, tbl: {}, tc: {}, tr: {} };
          r.condicionales[k] = { p: mezclar(a.p, c.p), r: mezclar(a.r, c.r), tbl: mezclar(a.tbl, c.tbl), tc: mezclar(a.tc, c.tc), tr: mezclar(a.tr, c.tr) };
        }
      }
      cache.set(id, r);
      return r;
    }
    return {
      existe: (id) => estilos.has(id),
      nombre: (id) => (estilos.has(id) ? estilos.get(id).nombre : ''),
      siguiente: (id) => (estilos.has(id) ? estilos.get(id).siguiente : null),
      // Propiedades efectivas de un párrafo con ese estilo (sin id: el estilo por defecto), con docDefaults.
      // `medio`: formato intermedio { p, r } (el del estilo de tabla), entre los valores por defecto y el estilo.
      parrafo(id, medio) {
        const e = resolver(id && estilos.has(id) ? id : porDefecto.paragraph);
        return { p: mezclar(mezclar(def.p, medio && medio.p), e.p), r: mezclar(mezclar(def.r, medio && medio.r), e.r) };
      },
      // Solo lo que el estilo de carácter agrega sobre el texto.
      caracter(id) { return id && estilos.has(id) ? resolver(id).r : {}; },
      // Estilo de tabla ya resuelto (con la herencia y los formatos condicionales).
      tabla(id) { return resolver(id && estilos.has(id) ? id : porDefecto.table); },
      base: def,
    };
  }

  // ───────── Numeración ─────────
  const SIMBOLOS = { '': '•', '': '▪', '': '➢', '': '❖', '': '✓', '': '◦', '': '➔', 'o': '○', 'ø': '➢', 'Ø': '➢', '§': '▪', 'v': '❖', 'ü': '✓', '·': '•' };
  const simboloDeViñeta = (t) => (t && SIMBOLOS[t]) || (!t || /[-]/.test(t) ? '•' : t);

  function leerNivel(l) {
    return limpiar({
      fmt: val(hijo(l, 'numFmt')), texto: val(hijo(l, 'lvlText')), inicio: num(val(hijo(l, 'start'))),
      sufijo: val(hijo(l, 'suff')), estilo: val(hijo(l, 'pStyle')), reinicio: hijo(l, 'lvlRestart') ? num(val(hijo(l, 'lvlRestart'))) : undefined,
      legal: encendido(hijo(l, 'isLgl')), jc: val(hijo(l, 'lvlJc')),
      p: leerPPr(hijo(l, 'pPr')), r: leerRPr(hijo(l, 'rPr')),
    });
  }
  function leerNumeracion(raiz) {
    const abstractos = new Map();
    const nums = new Map();
    if (raiz) {
      for (const a of lista(raiz, 'abstractNum')) {
        const niveles = new Map();
        for (const l of lista(a, 'lvl')) niveles.set(l.a.ilvl || '0', leerNivel(l));
        abstractos.set(a.a.abstractNumId, { niveles, vinculo: val(hijo(a, 'numStyleLink')) });
      }
      for (const n of lista(raiz, 'num')) {
        const sobre = new Map();
        for (const o of lista(n, 'lvlOverride')) {
          const l = hijo(o, 'lvl');
          sobre.set(o.a.ilvl || '0', { inicio: num(val(hijo(o, 'startOverride'))), nivel: l ? leerNivel(l) : null });
        }
        nums.set(n.a.numId, { abstracto: val(hijo(n, 'abstractNumId')), sobre });
      }
    }
    return crearNumeracion(abstractos, nums);
  }

  function crearNumeracion(abstractos, nums) {
    function abstractoDe(numId) {
      const n = nums.get(numId);
      return n ? abstractos.get(n.abstracto) : null;
    }
    // Definición efectiva de un nivel: la del abstracto con lo que cambie el `num`.
    function nivel(numId, ilvl) {
      const n = nums.get(numId);
      const a = abstractoDe(numId);
      if (!n || !a) return null;
      const base = a.niveles.get(ilvl) || null;
      const o = n.sobre.get(ilvl);
      const r = mezclar(base, o && o.nivel);
      return Object.keys(r).length ? r : null;
    }
    // Lleva la cuenta de cada lista a medida que se recorren los párrafos.
    function numerador() {
      const cuentas = new Map(); // idAbstracto → { ilvl → valor actual }
      return {
        siguiente(numId, ilvl) {
          const n = nums.get(numId);
          const def = nivel(numId, ilvl);
          if (!n || !def) return null;
          const nivelN = parseInt(ilvl, 10) || 0;
          let c = cuentas.get(n.abstracto);
          if (!c) { c = new Map(); cuentas.set(n.abstracto, c); }
          // El reinicio con startOverride solo vale la primera vez que se usa ese num.
          const o = n.sobre.get(ilvl);
          const marca = 'o' + numId + ':' + ilvl;
          if (o && o.inicio !== undefined && !c.has(marca)) { c.set(marca, true); c.set(nivelN, o.inicio - 1); }
          const inicio = def.inicio === undefined ? 1 : def.inicio;
          c.set(nivelN, c.has(nivelN) ? c.get(nivelN) + 1 : inicio);
          // Los niveles más profundos empiezan de nuevo (salvo que lvlRestart diga otra cosa).
          for (const k of [...c.keys()]) {
            if (typeof k !== 'number' || k <= nivelN) continue;
            const dk = nivel(numId, String(k));
            const reinicio = dk && dk.reinicio !== undefined ? dk.reinicio : nivelN + 1;
            if (reinicio !== 0 && nivelN + 1 <= reinicio) c.delete(k);
          }
          return { texto: etiqueta(numId, def, c, nivelN), nivel: def };
        },
      };
    }
    function etiqueta(numId, def, c, nivelN) {
      if (def.fmt === 'none') return '';
      if (def.fmt === 'bullet') return simboloDeViñeta(def.texto);
      const plantilla = def.texto === undefined ? '%' + (nivelN + 1) + '.' : def.texto;
      return plantilla.replace(/%([1-9])/g, (todo, d) => {
        const k = parseInt(d, 10) - 1;
        const dk = nivel(numId, String(k));
        if (!dk) return todo;
        const inicio = dk.inicio === undefined ? 1 : dk.inicio;
        const v = c.has(k) ? c.get(k) : inicio;
        return formatear(v, k === nivelN || !def.legal ? dk.fmt : 'decimal');
      });
    }
    return { nivel, numerador, existe: (numId) => nums.has(numId) };
  }

  const ROMANOS = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
  function romano(n) {
    let s = '';
    for (const [v, r] of ROMANOS) while (n >= v) { s += r; n -= v; }
    return s;
  }
  function letras(n) { // 1→a … 26→z, 27→aa, 28→bb (así cuenta Word)
    const l = String.fromCharCode(97 + ((n - 1) % 26));
    return l.repeat(Math.floor((n - 1) / 26) + 1);
  }
  function formatear(n, fmt) {
    // Letras y romanos crecen con el número (con 2·10^9 eran 77 millones de letras por párrafo y la página se
    // colgaba): fuera de 1…32767 (el tope de Word) salen vacíos si no son positivos y en cifras si son enormes.
    if (/Letter|Roman/.test(fmt || '') && !(n >= 1 && n <= 32767)) return n >= 1 ? String(n) : '';
    switch (fmt) {
      case 'lowerLetter': return letras(n);
      case 'upperLetter': return letras(n).toUpperCase();
      case 'lowerRoman': return romano(n);
      case 'upperRoman': return romano(n).toUpperCase();
      case 'decimalZero': return n < 10 ? '0' + n : String(n);
      case 'ordinal': { const r = n % 100; return n + (r >= 11 && r <= 13 ? 'th' : ['th', 'st', 'nd', 'rd'][n % 10] || 'th'); }
      case 'bullet': return '•';
      case 'none': return '';
      default: return String(n);
    }
  }

  // ───────── Ajustes ─────────
  function leerAjustes(raiz) {
    const a = { tabPorDefecto: 36, modoCompat: 12, parImpar: false };
    if (!raiz) return a;
    const t = hijo(raiz, 'defaultTabStop');
    if (t && tw(t.a.val)) a.tabPorDefecto = tw(t.a.val);
    const cp = hijo(raiz, 'compat');
    if (cp) {
      for (const s of lista(cp, 'compatSetting')) if (s.a.name === 'compatibilityMode') a.modoCompat = num(s.a.val) || a.modoCompat;
    }
    a.parImpar = !!hijo(raiz, 'evenAndOddHeaders');
    return a;
  }

  // ───────── Entrada ─────────
  async function cargar(paquete) {
    const { zip, rels } = paquete;
    const parte = async (fin, defecto) => {
      const ruta = porTipo(rels, fin) || defecto;
      const f = ruta ? zip.file(ruta) : null;
      return f ? arbol(await f.async('string')) : null;
    };
    const estilosRaiz = await parte('/styles', 'word/styles.xml');
    const numRaiz = await parte('/numbering', 'word/numbering.xml');
    const temaRaiz = await parte('/theme', 'word/theme/theme1.xml');
    const ajustesRaiz = await parte('/settings', 'word/settings.xml');
    const tema = leerTema(temaRaiz);
    return {
      tema,
      ajustes: leerAjustes(ajustesRaiz),
      estilos: crearEstilos(leerEstilos(estilosRaiz)),
      numeracion: leerNumeracion(numRaiz),
      color: (c) => colorHex(tema, c),
      fuente: (f, tipo) => nombreFuente(tema, f, tipo),
      mezclar, leerRPr, leerPPr, leerTblPr, leerTcPr, leerTrPr,
    };
  }

  globalThis.DocxEstilos = { cargar, mezclar, leerRPr, leerPPr, leerTblPr, leerTcPr, leerTrPr, formatear, aplicarTinte };
})();

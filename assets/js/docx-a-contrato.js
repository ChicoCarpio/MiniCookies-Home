/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Un .docx leído como «contrato de documento» (el que describe docx-desde-contrato.js): es el tercer lector de Word, junto a DocLector y OdtLector, y lo que permite que los
// escritores de .doc y .odt reciban cualquier origen (.docx → .doc, .docx → .odt) sin un conversor por pareja.
//
//   const { contrato, mensajes } = await DocxAContrato.leer(bytes);     // bytes: Uint8Array de un .docx
//
// Lleva: estilos (con su herencia), valores predeterminados del documento, listas con numeración (se resuelven los numId y los estilos con lista), secciones (hoja, márgenes,
// columnas, tipo, encabezado y pie predeterminados; una sección sin los suyos hereda los de la anterior, como en Word), párrafos y su formato, tablas (anidadas, combinadas,
// bordes y relleno), imágenes en línea o flotantes (dibujadas en línea), enlaces, campos (PAGE, NUMPAGES…), marcadores y texto con cambios aceptados (lo borrado se quita).
// No lleva: notas al pie y comentarios, cuadros de texto, formas, ecuaciones, estilos condicionales de tabla, encabezado de primera página o de páginas pares.
// mensajes: 'notas', 'dibujos', 'imagen-formato' (lo avisa el escritor), 'imagen-falta', 'imagen-externa'.
// Requiere zip-propio.js y xml-ligero.js.
(function () {
  'use strict';
  const MAX_TEXTO_XML = 96 * 1024 * 1024;
  const MAX_IMAGENES = 5000;
  const MAX_BLOQUES = 4000000;
  const MAX_PROFUNDIDAD_TABLAS = 8;

  const hijos = (e, n) => (e ? e.h.filter((x) => typeof x !== 'string' && (n === undefined || x.n === n)) : []);
  const hijo = (e, n) => (e ? e.h.find((x) => typeof x !== 'string' && x.n === n) || null : null);
  const textoDe = (e) => (e ? e.h.map((x) => (typeof x === 'string' ? x : textoDe(x))).join('') : '');
  const num = (v) => { const n = Number(v); return Number.isFinite(n) ? n : null; };
  const entero = (v) => { const n = parseInt(v, 10); return Number.isFinite(n) ? n : null; };
  const activo = (e) => !!e && !/^(0|false|off|none)$/i.test(e.a.val == null ? '1' : e.a.val);
  const color6 = (c) => (/^[0-9a-fA-F]{6}$/.test(c || '') ? c.toUpperCase() : null);
  const definir = (o, k, v) => { if (v !== undefined && v !== null) o[k] = v; };
  const vacio = (o) => Object.keys(o).length === 0;

  const ALINEACION = { left: 'izquierda', start: 'izquierda', center: 'centro', right: 'derecha', end: 'derecha', both: 'justificado', distribute: 'justificado', justify: 'justificado' };
  const TAB = { left: 'izquierda', start: 'izquierda', center: 'centro', right: 'derecha', end: 'derecha', decimal: 'decimal' };
  const RELLENO_TAB = { dot: 'punto', middleDot: 'punto', hyphen: 'guion', underscore: 'linea', heavy: 'linea' };
  const BORDE = { single: 'simple', thick: 'simple', double: 'doble', dotted: 'punteado', dashed: 'guiones', dashSmallGap: 'guiones', nil: 'ninguno', none: 'ninguno' };
  const FORMATOS_LISTA = new Set(['decimal', 'bullet', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman', 'ordinal', 'cardinalText', 'ordinalText', 'decimalZero', 'none']);
  const CAMPOS = new Set(['PAGE', 'NUMPAGES', 'SECTIONPAGES', 'DATE', 'TIME', 'SAVEDATE', 'PRINTDATE', 'CREATEDATE', 'FILENAME', 'TITLE', 'AUTHOR', 'NUMWORDS', 'NUMCHARS']);

  // ───────── Propiedades ─────────
  function Lector(zip, mensajes) {
    this.zip = zip; this.mensajes = mensajes; this.tema = {}; this.estiloLista = new Map(); this.estilosTabla = new Map(); this.estilosParrafo = new Map(); this.estiloPredeterminado = 'Normal'; this.numeracion = null;
    this.imagenes = new Map(); this.totalImagenes = 0; this.bloques = 0;
  }
  const P = Lector.prototype;

  P.fuenteTema = function (nombre) {
    const t = String(nombre || '');
    const clase = /^major/.test(t) ? this.tema.mayor : this.tema.menor;
    if (!clase) return null;
    return /EastAsia/i.test(t) ? clase.ea || null : /Bidi|cs$/i.test(t) ? clase.cs || null : clase.latin || null;
  };

  P.propsR = function (rPr) {
    const r = {};
    if (!rPr) return r;
    const f = hijo(rPr, 'rFonts');
    if (f) {
      definir(r, 'fuente', f.a.ascii || f.a.hAnsi || (f.a.asciiTheme && this.fuenteTema(f.a.asciiTheme)) || (f.a.hAnsiTheme && this.fuenteTema(f.a.hAnsiTheme)) || null);
      definir(r, 'fuenteAsia', f.a.eastAsia || (f.a.eastAsiaTheme && this.fuenteTema(f.a.eastAsiaTheme)) || null);
      definir(r, 'fuenteCs', f.a.cs || (f.a.cstheme && this.fuenteTema(f.a.cstheme)) || null);
    }
    const b = hijo(rPr, 'b'); if (b) r.negrita = activo(b);
    const i = hijo(rPr, 'i'); if (i) r.cursiva = activo(i);
    const u = hijo(rPr, 'u'); if (u) r.subrayado = u.a.val === 'none' ? false : /^double/.test(u.a.val || '') ? 'doble' : true;
    const s = hijo(rPr, 'strike') || hijo(rPr, 'dstrike'); if (s) r.tachado = activo(s);
    const va = hijo(rPr, 'vertAlign');
    if (va) { if (va.a.val === 'superscript') r.sup = true; else if (va.a.val === 'subscript') r.sub = true; }
    const vs = hijo(rPr, 'smallCaps'); if (vs) r.versalitas = activo(vs);
    const m = hijo(rPr, 'caps'); if (m) r.mayusculas = activo(m);
    const o = hijo(rPr, 'vanish'); if (o && activo(o)) r.oculto = true;
    const c = hijo(rPr, 'color'); if (c && color6(c.a.val)) r.color = color6(c.a.val);
    const sz = hijo(rPr, 'sz'); if (sz && num(sz.a.val) != null) r.tamano = num(sz.a.val) / 2;
    const h = hijo(rPr, 'highlight'); if (h && h.a.val && h.a.val !== 'none') r.resaltado = h.a.val;
    return r;
  };

  P.bordes = function (e) {
    if (!e) return null;
    const salida = {};
    for (const [clave, nombres] of [['arriba', ['top']], ['abajo', ['bottom']], ['izquierda', ['left', 'start']], ['derecha', ['right', 'end']]]) {
      const b = nombres.map((n) => hijo(e, n)).find(Boolean);
      if (!b) continue;
      const estilo = BORDE[b.a.val] || 'simple';
      if (estilo === 'ninguno') continue;
      salida[clave] = { tam: num(b.a.sz) == null ? 4 : num(b.a.sz), color: color6(b.a.color) || undefined, estilo, espacio: num(b.a.space) || 0 };
    }
    return vacio(salida) ? null : salida;
  };

  P.propsP = function (pPr) {
    const p = {};
    if (!pPr) return p;
    if (activo(hijo(pPr, 'keepNext'))) p.mantener = true;
    if (activo(hijo(pPr, 'keepLines'))) p.mantenerJuntas = true;
    if (activo(hijo(pPr, 'pageBreakBefore'))) p.saltoAntes = true;
    const bd = this.bordes(hijo(pPr, 'pBdr')); if (bd) p.bordes = bd;
    const shd = hijo(pPr, 'shd'); if (shd && color6(shd.a.fill)) p.relleno = color6(shd.a.fill);
    const tabs = hijos(hijo(pPr, 'tabs'), 'tab').filter((t) => t.a.val !== 'clear' && t.a.val !== 'bar' && num(t.a.pos) != null);
    if (tabs.length) p.tabs = tabs.map((t) => { const x = { tipo: TAB[t.a.val] || 'izquierda', pos: num(t.a.pos) }; if (RELLENO_TAB[t.a.leader]) x.relleno = RELLENO_TAB[t.a.leader]; return x; });
    const sp = hijo(pPr, 'spacing');
    if (sp) {
      if (num(sp.a.before) != null) p.antes = num(sp.a.before);
      if (num(sp.a.after) != null) p.despues = num(sp.a.after);
      if (num(sp.a.line) != null) { p.linea = num(sp.a.line); p.regla = sp.a.lineRule === 'exact' ? 'exacto' : sp.a.lineRule === 'atLeast' ? 'minimo' : 'auto'; }
    }
    const ind = hijo(pPr, 'ind');
    if (ind) {
      const izq = ind.a.left != null ? ind.a.left : ind.a.start, der = ind.a.right != null ? ind.a.right : ind.a.end;
      if (num(izq) != null) p.izq = num(izq);
      if (num(der) != null) p.der = num(der);
      if (num(ind.a.hanging) != null && num(ind.a.hanging) > 0) p.primera = -num(ind.a.hanging);
      else if (num(ind.a.firstLine) != null) p.primera = num(ind.a.firstLine);
    }
    const jc = hijo(pPr, 'jc'); if (jc && ALINEACION[jc.a.val]) p.alineacion = ALINEACION[jc.a.val];
    const ol = hijo(pPr, 'outlineLvl'); if (ol && num(ol.a.val) != null && num(ol.a.val) >= 0 && num(ol.a.val) <= 8) p.nivelEsquema = num(ol.a.val);
    return p;
  };

  // ───────── Partes del paquete ─────────
  P.texto = async function (ruta) {
    const f = this.zip.file(ruta);
    if (!f) return null;
    const t = await f.async('string');
    if (t.length > MAX_TEXTO_XML) throw new Error('El documento es demasiado grande para abrirlo aquí.');
    return t;
  };
  P.arbol = async function (ruta) {
    const t = await this.texto(ruta);
    return t == null ? null : XmlLigero.arbol(t);
  };
  P.relaciones = async function (ruta) {
    const carpeta = ruta.replace(/[^/]*$/, ''), nombre = ruta.slice(carpeta.length);
    const mapa = new Map();
    const raiz = await this.arbol(`${carpeta}_rels/${nombre}.rels`);
    for (const r of hijos(raiz, 'Relationship')) {
      const externo = r.a.TargetMode === 'External';
      let destino = String(r.a.Target || '');
      if (!externo) {
        const partes = (destino.startsWith('/') ? destino.slice(1) : carpeta + destino).split('/');
        const sal = [];
        for (const x of partes) { if (x === '..') sal.pop(); else if (x && x !== '.') sal.push(x); }
        destino = sal.join('/');
      }
      mapa.set(r.a.Id, { tipo: String(r.a.Type || '').replace(/^.*\//, ''), destino, externo });
    }
    return mapa;
  };

  // ───────── Tema, estilos y numeración ─────────
  P.leerTema = async function (rels) {
    const t = [...rels.values()].find((r) => r.tipo === 'theme');
    const raiz = t ? await this.arbol(t.destino).catch(() => null) : null;
    const esquema = hijo(hijo(raiz, 'themeElements'), 'fontScheme');
    for (const [clave, nombre] of [['mayor', 'majorFont'], ['menor', 'minorFont']]) {
      const f = hijo(esquema, nombre);
      if (!f) continue;
      const g = (n) => { const x = hijo(f, n); return x && x.a.typeface ? x.a.typeface : null; };
      this.tema[clave] = { latin: g('latin'), ea: g('ea'), cs: g('cs') };
    }
  };

  P.leerEstilos = async function (rels) {
    const rel = [...rels.values()].find((r) => r.tipo === 'styles');
    const raiz = rel ? await this.arbol(rel.destino) : null;
    const salida = { predeterminado: null, estilos: [] };
    if (!raiz) return salida;
    const dd = hijo(raiz, 'docDefaults');
    if (dd) {
      const r = this.propsR(hijo(hijo(dd, 'rPrDefault'), 'rPr')), p = this.propsP(hijo(hijo(dd, 'pPrDefault'), 'pPr'));
      if (!vacio(r) || !vacio(p)) salida.predeterminado = { r, p };
    }
    const todos = hijos(raiz, 'style');
    for (const s of todos.slice(0, 4096)) {
      const id = s.a.styleId;
      if (!id) continue;
      const tipo = s.a.type === 'character' ? 'caracter' : s.a.type === 'table' ? 'tabla' : s.a.type === 'paragraph' ? 'parrafo' : null;
      if (!tipo) continue;
      const nombre = hijo(s, 'name'), basado = hijo(s, 'basedOn'), sig = hijo(s, 'next');
      const e = { id, nombre: nombre ? nombre.a.val : id, tipo };
      if (basado && basado.a.val !== id) e.basadoEn = basado.a.val;
      if (sig && sig.a.val) e.siguiente = sig.a.val;
      const pPr = hijo(s, 'pPr');
      const numPr = hijo(pPr, 'numPr');
      if (numPr && hijo(numPr, 'numId')) this.estiloLista.set(id, { numId: entero(hijo(numPr, 'numId').a.val), nivel: entero((hijo(numPr, 'ilvl') || { a: {} }).a.val) || 0 });
      if (tipo === 'parrafo') { const p = this.propsP(pPr); if (!vacio(p)) e.p = p; this.estilosParrafo.set(id, { basadoEn: e.basadoEn, p }); if (s.a.default === '1') this.estiloPredeterminado = id; }
      const r = this.propsR(hijo(s, 'rPr')); if (!vacio(r)) e.r = r;
      if (tipo === 'tabla') this.estilosTabla.set(id, { basadoEn: e.basadoEn, bordes: hijo(hijo(s, 'tblPr'), 'tblBorders'), p: this.propsP(pPr) });
      salida.estilos.push(e);
    }
    return salida;
  };

  P.leerNumeracion = async function (rels) {
    const rel = [...rels.values()].find((r) => r.tipo === 'numbering');
    const raiz = rel ? await this.arbol(rel.destino) : null;
    const abstractos = new Map(), nums = new Map();
    this.nivelPorEstilo = new Map();
    for (const a of hijos(raiz, 'abstractNum')) {
      abstractos.set(a.a.abstractNumId, a);
      for (const l of hijos(a, 'lvl')) { const ps = hijo(l, 'pStyle'); if (ps && ps.a.val && !this.nivelPorEstilo.has(ps.a.val)) this.nivelPorEstilo.set(ps.a.val, entero(l.a.ilvl) || 0); }
    }
    for (const n of hijos(raiz, 'num')) nums.set(n.a.numId, n);
    this.numeracion = { abstractos, nums, listas: [], hechas: new Map() };
  };

  P.nivelLista = function (lvl, inicioForzado) {
    const v = {};
    const fmt = hijo(lvl, 'numFmt');
    v.formato = fmt && FORMATOS_LISTA.has(fmt.a.val) ? fmt.a.val : 'decimal';
    const txt = hijo(lvl, 'lvlText');
    v.texto = txt && txt.a.val != null ? txt.a.val : '%1.';
    const ind = hijo(hijo(lvl, 'pPr'), 'ind');
    v.izquierda = Math.max(0, ind ? num(ind.a.left != null ? ind.a.left : ind.a.start) || 0 : 0);
    v.colgante = Math.max(0, ind ? num(ind.a.hanging) || 0 : 0);
    const st = hijo(lvl, 'start');
    v.inicio = inicioForzado != null ? inicioForzado : st && num(st.a.val) != null ? num(st.a.val) : 1;
    const r = this.propsR(hijo(lvl, 'rPr'));
    if (r.fuente) v.fuente = r.fuente;
    const { fuente, fuenteAsia, fuenteCs, ...resto } = r;
    void fuente; void fuenteAsia; void fuenteCs;
    if (!vacio(resto)) v.r = resto;
    return v;
  };

  // numId del documento → id de lista del contrato. Los numId sin cambios de nivel comparten lista con los de su mismo abstractNum (la numeración sigue); con «empezar en» van aparte.
  P.listaDe = function (numId) {
    const N = this.numeracion;
    if (!N || numId == null) return null;
    const clave = String(numId);
    if (N.hechas.has(clave)) return N.hechas.get(clave);
    const n = N.nums.get(clave);
    const abs = n ? N.abstractos.get((hijo(n, 'abstractNumId') || { a: {} }).a.val) : null;
    if (!abs) { N.hechas.set(clave, null); return null; }
    const sobres = new Map();
    for (const o of hijos(n, 'lvlOverride')) sobres.set(o.a.ilvl, o);
    const id = sobres.size ? 'n' + clave : 'a' + abs.a.abstractNumId;
    N.hechas.set(clave, id);
    if (!N.listas.some((l) => l.id === id)) {
      const base = new Map();
      for (const l of hijos(abs, 'lvl')) base.set(l.a.ilvl, l);
      const niveles = [];
      for (let k = 0; k < 9; k++) {
        const o = sobres.get(String(k));
        const lvl = (o && hijo(o, 'lvl')) || base.get(String(k));
        const so = o && hijo(o, 'startOverride');
        niveles.push(lvl ? this.nivelLista(lvl, so && num(so.a.val) != null ? num(so.a.val) : null) : { formato: 'decimal', texto: `%${k + 1}.`, izquierda: 720 * (k + 1), colgante: 360, inicio: 1 });
      }
      N.listas.push({ id, niveles });
    }
    return id;
  };

  // La lista de un párrafo: la suya, o la de su estilo (o la de los estilos de los que hereda)
  P.listaParrafo = function (pPr, estiloId) {
    const numPr = hijo(pPr, 'numPr');
    if (numPr) {
      const nid = hijo(numPr, 'numId');
      if (nid) {
        const n = entero(nid.a.val), nivel = entero((hijo(numPr, 'ilvl') || { a: {} }).a.val) || 0;
        if (n === 0) return { id: null, nivel: 0 };
        const id = this.listaDe(n);
        return id ? { id, nivel: Math.max(0, Math.min(8, nivel)) } : null;
      }
    }
    for (let e = estiloId, k = 0; e && k < 32; k++) {
      const s = this.estiloLista.get(e);
      if (s) {
        if (s.numId === 0) return { id: null, nivel: 0 };
        const id = this.listaDe(s.numId);
        const nivel = this.nivelPorEstilo.has(e) ? this.nivelPorEstilo.get(e) : s.nivel || 0;
        return id ? { id, nivel: Math.max(0, Math.min(8, nivel)) } : null;
      }
      e = this.basePorEstilo.get(e);
    }
    return null;
  };

  // ───────── Imágenes ─────────
  P.imagen = async function (dib, ctx) {
    const blip = this.buscar(dib, 'blip');
    if (!blip) return null;
    const rid = blip.a.embed;
    if (!rid) { if (blip.a.link) this.mensajes.add('imagen-externa'); return null; }
    const rel = ctx.rels.get(rid);
    const f = rel && !rel.externo ? this.zip.file(rel.destino) : null;
    if (!f) { this.mensajes.add('imagen-falta'); return null; }
    let datos = this.imagenes.get(rel.destino);
    if (!datos) {
      if (++this.totalImagenes > MAX_IMAGENES) throw new Error('El documento tiene demasiadas imágenes para abrirlo aquí.');
      datos = await f.async('uint8array');
      this.imagenes.set(rel.destino, datos);
    }
    const ext = this.buscar(dib, 'extent');
    const ancho = ext && num(ext.a.cx) ? num(ext.a.cx) / 12700 : 72, alto = ext && num(ext.a.cy) ? num(ext.a.cy) / 12700 : 72;
    const dp = this.buscar(dib, 'docPr');
    const img = { datos, ancho, alto };
    if (dp && dp.a.descr) img.texto = dp.a.descr;
    return img;
  };
  P.buscar = function (e, nombre, prof = 0) {
    if (!e || prof > 24) return null;
    for (const x of e.h) {
      if (typeof x === 'string') continue;
      if (x.n === nombre) return x;
      const r = this.buscar(x, nombre, prof + 1);
      if (r) return r;
    }
    return null;
  };

  // ───────── Texto: párrafos y sus runs ─────────
  P.parrafo = async function (p, ctx, pila) {
    if (++this.bloques > MAX_BLOQUES) throw new Error('El documento tiene demasiados párrafos para abrirlo aquí.');
    const pPr = hijo(p, 'pPr');
    const est = hijo(pPr, 'pStyle');
    const b = { tipo: 'parrafo', hijos: [] };
    if (est && est.a.val) b.estilo = est.a.val;
    const props = this.propsP(pPr);
    const lista = this.listaParrafo(pPr, est && est.a.val);
    if (lista) props.lista = lista;
    if (!vacio(props)) b.p = props;
    const salida = b.hijos;
    // la pila de campos viene del documento entero: un campo (un índice) puede abrirse en un párrafo y cerrarse en otro
    const destino = () => {
      const f = pila[pila.length - 1];
      return !f ? salida : f.fase === 'res' ? f.coleccion : null;   // dentro de la instrucción de un campo no hay texto que llevar
    };
    const emitir = (x) => { const d = destino(); if (d) d.push(x); };
    await this.recorrerRuns(p, ctx, pila, emitir);
    // los campos que siguen abiertos al terminar el párrafo (un índice) vuelcan lo recogido en él; lo que viene después va suelto
    for (let k = pila.length - 1; k >= 0; k--) {
      const c = pila[k];
      if (c.fase !== 'res' || !c.coleccion.length) continue;
      const padre = k === 0 ? salida : pila[k - 1].fase === 'res' ? pila[k - 1].coleccion : null;
      this.volcarCampo(c, (x) => { if (padre) padre.push(x); });
      c.coleccion = [];
    }
    return b;
  };

  P.volcarCampo = function (c, a) {
    const instr = c.instr.trim();
    const nombre = (instr.split(/\s+/)[0] || '').toUpperCase();
    const hijosC = c.coleccion || [];
    if (!hijosC.length) return;
    if (nombre === 'HYPERLINK') {
      const url = (/"([^"]*)"/.exec(instr) || [])[1], ancla = (/\\l\s+"([^"]*)"/.exec(instr) || [])[1];
      if (ancla) a({ enlace: { ancla }, hijos: hijosC });
      else if (url) a({ enlace: { url }, hijos: hijosC });
      else for (const x of hijosC) a(x);
    } else if (CAMPOS.has(nombre) && c.coleccion) {
      const texto = hijosC.map((x) => x.texto || '').join('');
      const primero = hijosC.find((x) => x.r || x.estilo);
      const x = { campo: instr, resultado: texto };
      if (primero && primero.r) x.r = primero.r;
      a(x);
    } else for (const x of hijosC) a(x);
  };

  P.recorrerRuns = async function (e, ctx, pila, emitir) {
    for (const c of e.h) {
      if (typeof c === 'string') continue;
      switch (c.n) {
        case 'r': await this.run(c, ctx, pila, emitir); break;
        case 'hyperlink': {
          const rid = c.a.id, ancla = c.a.anchor;
          const rel = rid ? ctx.rels.get(rid) : null;
          const lista = [];
          await this.recorrerRuns(c, ctx, pila, (x) => lista.push(x));
          if (ancla) emitir({ enlace: { ancla }, hijos: lista });
          else if (rel && rel.externo) emitir({ enlace: { url: rel.destino }, hijos: lista });
          else for (const x of lista) emitir(x);
          break;
        }
        case 'fldSimple': {
          const campo = { instr: c.a.instr || '', fase: 'res', coleccion: [], padre: null };
          await this.recorrerRuns(c, ctx, pila, (x) => campo.coleccion.push(x));
          this.volcarCampo(campo, emitir);
          break;
        }
        case 'ins': case 'moveTo': case 'smartTag': case 'customXml': case 'sdtContent': case 'sdt': case 'dir': case 'bdo':
          await this.recorrerRuns(c, ctx, pila, emitir); break;
        case 'AlternateContent': {
          const ch = hijo(c, 'Choice') || hijo(c, 'Fallback');
          if (ch) await this.recorrerRuns(ch, ctx, pila, emitir);
          break;
        }
        case 'bookmarkStart': if (c.a.name && c.a.name !== '_GoBack') emitir({ marcador: c.a.name }); break;
        default: break;   // del, moveFrom, proofErr, comentarios…
      }
    }
  };

  P.run = async function (r, ctx, pila, emitir) {
    const rPr = hijo(r, 'rPr');
    const props = this.propsR(rPr);
    const rs = hijo(rPr, 'rStyle');
    const base = {};
    if (rs && rs.a.val) base.estilo = rs.a.val;
    if (!vacio(props)) base.r = props;
    let buffer = '';
    const volcar = () => { if (buffer !== '') { emitir({ texto: buffer, ...base }); buffer = ''; } };
    for (const c of r.h) {
      if (typeof c === 'string') continue;
      switch (c.n) {
        case 't': buffer += textoDe(c); break;
        case 'tab': buffer += '\t'; break;
        case 'cr': buffer += '\n'; break;
        case 'noBreakHyphen': buffer += '‑'; break;
        case 'softHyphen': buffer += '­'; break;
        case 'br':
          volcar();
          if (c.a.type === 'page') emitir({ salto: 'pagina' }); else if (c.a.type === 'column') emitir({ salto: 'columna' }); else emitir({ salto: 'linea' });
          break;
        case 'fldChar': {
          volcar();
          const t = c.a.fldCharType;
          if (t === 'begin') pila.push({ instr: '', fase: 'instr', coleccion: null, padre: null });
          else if (t === 'separate' && pila.length) { const f = pila[pila.length - 1]; f.fase = 'res'; f.coleccion = []; }
          else if (t === 'end' && pila.length) { const f = pila.pop(); f.fase = 'fin'; if (f.coleccion) this.volcarCampo(f, emitir); }
          break;
        }
        case 'instrText': { const f = pila[pila.length - 1]; if (f && f.fase === 'instr') f.instr += textoDe(c); break; }
        case 'drawing': {
          volcar();
          const img = await this.imagen(c, ctx);
          if (img) emitir({ imagen: img }); else if (!this.buscar(c, 'blip')) this.mensajes.add('dibujos');
          break;
        }
        case 'pict': case 'object': volcar(); this.mensajes.add('dibujos'); break;
        case 'footnoteReference': case 'endnoteReference': this.mensajes.add('notas'); break;
        default: break;
      }
    }
    volcar();
  };

  // ───────── Bloques: párrafos y tablas ─────────
  P.bordesTabla = function (tblPr) {
    let b = hijo(tblPr, 'tblBorders');
    for (let id = (hijo(tblPr, 'tblStyle') || { a: {} }).a.val, k = 0; !b && id && k < 32; k++) {
      const s = this.estilosTabla.get(id);
      if (!s) break;
      b = s.bordes;
      id = s.basadoEn;
    }
    if (!b) return false;
    const lados = ['top', 'bottom', 'left', 'right', 'insideH', 'insideV', 'start', 'end'].map((n) => hijo(b, n)).filter(Boolean);
    const visible = lados.find((l) => BORDE[l.a.val] !== 'ninguno' && l.a.val);
    if (!visible) return false;
    return { tam: num(visible.a.sz) == null ? 4 : num(visible.a.sz), color: color6(visible.a.color) || undefined };
  };

  // Las propiedades de párrafo de un estilo de tabla (p. ej. «Tabla con cuadrícula»: espacio después 0) no existen en .doc ni en .odt: se llevan a los párrafos de las celdas,
  // salvo lo que el propio párrafo (o su estilo) ya define.
  P.cadena = function (mapa, id) {
    const lista = [];
    for (let e = id, k = 0; e && mapa.has(e) && k < 32; k++) { lista.unshift(mapa.get(e).p || {}); e = mapa.get(e).basadoEn; }
    return Object.assign({}, ...lista);
  };
  P.aplicarEstiloTabla = function (bloques, deTabla) {
    for (const b of bloques) {
      if (b.tipo !== 'parrafo') continue;
      const propia = this.cadena(this.estilosParrafo, b.estilo || this.estiloPredeterminado);
      for (const k of Object.keys(deTabla)) if (!(k in propia) && !(b.p && k in b.p)) (b.p = b.p || {})[k] = deTabla[k];
    }
  };

  P.tabla = async function (t, ctx, pila, prof) {
    if (prof > MAX_PROFUNDIDAD_TABLAS) return null;
    const tblPr = hijo(t, 'tblPr');
    const anchos = hijos(hijo(t, 'tblGrid'), 'gridCol').map((g) => num(g.a.w) || 0);
    const salida = { tipo: 'tabla', anchos, bordes: this.bordesTabla(tblPr), filas: [] };
    const est = hijo(tblPr, 'tblStyle');
    if (est && est.a.val && this.estilosTabla.has(est.a.val)) salida.estilo = est.a.val;
    const ind = hijo(tblPr, 'tblInd'); if (ind && num(ind.a.w)) salida.sangria = num(ind.a.w);
    const deTabla = salida.estilo ? this.cadena(this.estilosTabla, salida.estilo) : {};
    delete deTabla.lista;
    for (const tr of hijos(t, 'tr')) {
      const trPr = hijo(tr, 'trPr');
      const fila = { celdas: [] };
      const th = hijo(trPr, 'trHeight'); if (th && num(th.a.val)) fila.alto = num(th.a.val);
      if (activo(hijo(trPr, 'tblHeader'))) fila.encabezado = true;
      let col = 0;
      for (const tc of hijos(tr, 'tc')) {
        const tcPr = hijo(tc, 'tcPr');
        const span = Math.max(1, entero((hijo(tcPr, 'gridSpan') || { a: {} }).a.val) || 1);
        const celda = { hijos: [] };
        if (span > 1) celda.fusionH = span;
        const vm = hijo(tcPr, 'vMerge'); if (vm) celda.fusionV = vm.a.val === 'restart' ? 'inicio' : 'continua';
        const w = anchos.length ? anchos.slice(col, col + span).reduce((s, x) => s + x, 0) : 0;
        const tw = hijo(tcPr, 'tcW');
        celda.ancho = w || (tw && tw.a.type !== 'pct' && num(tw.a.w)) || 1440;
        const bd = this.bordes(hijo(tcPr, 'tcBorders')); if (bd) celda.bordes = bd;
        const shd = hijo(tcPr, 'shd'); if (shd && color6(shd.a.fill)) celda.relleno = color6(shd.a.fill);
        const va = hijo(tcPr, 'vAlign'); if (va && /^(center|bottom)$/.test(va.a.val)) celda.alineacionV = va.a.val === 'center' ? 'centro' : 'abajo';
        celda.hijos = await this.bloquesDe(tc, ctx, pila, prof + 1);
        if (!vacio(deTabla)) this.aplicarEstiloTabla(celda.hijos, deTabla);
        fila.celdas.push(celda);
        col += span;
      }
      if (fila.celdas.length) salida.filas.push(fila);
    }
    return salida;
  };

  P.bloquesDe = async function (e, ctx, pila, prof, onSectPr) {
    const salida = [];
    for (const c of e.h) {
      if (typeof c === 'string') continue;
      if (c.n === 'p') {
        const b = await this.parrafo(c, ctx, pila);
        salida.push(b);
        const sp = hijo(hijo(c, 'pPr'), 'sectPr');
        if (sp && onSectPr) onSectPr(sp, salida);
      } else if (c.n === 'tbl') {
        const t = await this.tabla(c, ctx, pila, prof);
        if (t) salida.push(t);
      } else if (c.n === 'sdt' || c.n === 'ins' || c.n === 'moveTo' || c.n === 'customXml') {
        const dentro = c.n === 'sdt' ? hijo(c, 'sdtContent') : c;
        if (dentro) for (const x of await this.bloquesDe(dentro, ctx, pila, prof, onSectPr)) salida.push(x);
      }
    }
    return salida;
  };

  // ───────── Secciones, encabezados y pies ─────────
  P.seccion = async function (sp, rels, previa) {
    const s = { pagina: {}, margenes: {}, hijos: [] };
    const pg = hijo(sp, 'pgSz');
    if (pg) { if (num(pg.a.w)) s.pagina.ancho = num(pg.a.w); if (num(pg.a.h)) s.pagina.alto = num(pg.a.h); if (pg.a.orient === 'landscape') s.pagina.horizontal = true; }
    const m = hijo(sp, 'pgMar');
    if (m) for (const [k, n] of [['arriba', 'top'], ['abajo', 'bottom'], ['izquierda', 'left'], ['derecha', 'right'], ['encabezado', 'header'], ['pie', 'footer']]) if (num(m.a[n]) != null) s.margenes[k] = num(m.a[n]);
    const ty = hijo(sp, 'type'); if (ty && ty.a.val === 'continuous') s.tipo = 'continua';
    const cols = hijo(sp, 'cols'); if (cols && entero(cols.a.num) > 1) s.columnas = { n: entero(cols.a.num), espacio: num(cols.a.space) == null ? 720 : num(cols.a.space) };
    for (const [ref, clave] of [['headerReference', 'encabezado'], ['footerReference', 'pie']]) {
      const r = hijos(sp, ref).find((x) => x.a.type === 'default');
      if (r) {
        const rel = rels.get(r.a.id);
        if (!rel) continue;
        if (!this.partes.has(rel.destino)) {
          const raiz = await this.arbol(rel.destino);
          const ctx = { rels: await this.relaciones(rel.destino) };
          this.partes.set(rel.destino, raiz ? await this.bloquesDe(raiz, ctx, []) : []);
        }
        s[clave] = this.partes.get(rel.destino);
      } else if (previa && previa[clave]) s[clave] = previa[clave];   // sin los suyos, hereda los de la sección anterior
    }
    return s;
  };

  P.convertir = async function () {
    const rels = await this.relaciones('word/document.xml');
    await this.leerTema(rels);
    const { predeterminado, estilos } = await this.leerEstilos(rels);
    this.basePorEstilo = new Map(estilos.filter((e) => e.basadoEn).map((e) => [e.id, e.basadoEn]));
    await this.leerNumeracion(rels);
    this.partes = new Map();
    const doc = await this.arbol('word/document.xml');
    const cuerpo = hijo(doc, 'body');
    if (!cuerpo) throw new Error('El archivo no tiene un documento de Word legible.');
    const ctx = { rels };
    const pila = [];
    const secciones = [];
    let actual = [];
    const cerrar = async (sp) => {
      const s = await this.seccion(sp, rels, secciones[secciones.length - 1]);
      s.hijos = actual;
      secciones.push(s);
      actual = [];
    };
    // los párrafos con sectPr cierran una sección; como las llamadas son asíncronas se recogen y se resuelven en orden
    const cierres = [];
    const bloques = await this.bloquesDe(cuerpo, ctx, pila, 0, (sp, salida) => cierres.push({ sp, hasta: salida.length }));
    let desde = 0;
    for (const c of cierres) { actual = bloques.slice(desde, c.hasta); desde = c.hasta; await cerrar(c.sp); }
    actual = bloques.slice(desde);
    await cerrar(hijo(cuerpo, 'sectPr'));
    const contrato = { secciones, estilos, listas: this.numeracion.listas };
    if (predeterminado) contrato.predeterminado = predeterminado;
    const props = await this.arbol('docProps/core.xml').catch(() => null);
    if (props) {
      const titulo = textoDe(hijo(props, 'title')), autor = textoDe(hijo(props, 'creator'));
      if (titulo || autor) contrato.propiedades = { titulo: titulo || undefined, autor: autor || undefined };
    }
    return contrato;
  };

  async function leer(bytes) {
    const datos = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const zip = await ZipPropio.loadAsync(datos);
    if (!zip.file('word/document.xml')) throw new Error('El archivo no es un documento de Word (.docx).');
    const mensajes = new Set();
    const l = new Lector(zip, mensajes);
    const contrato = await l.convertir();
    return { contrato, mensajes: [...mensajes] };
  }

  globalThis.DocxAContrato = { leer };
})();

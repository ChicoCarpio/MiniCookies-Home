/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Estilos de un .odt para el editor de Documentos (Fase 6b, etapa 6b.1). Sin DOM: corre igual en Node.
//
// Lee los estilos de content.xml y styles.xml (nombrados, automáticos, predeterminados, de lista, de página y maestras) y los entrega con el MISMO
// esquema de propiedades que DocxEstilos (docx-estilos.js), de modo que el maquetador, el pintor y la capa accesible no saben que el archivo es ODT:
//   r: fuentes{ascii,hAnsi,eastAsia,cs}, b, i, tam (pt), u, tachado, color{val}, resaltado, sombreado{relleno}, alineacionVertical, espaciado, escala, idioma…
//   p: jc, sangria{izq,der,primera,colgante}, espacio{antes,despues,linea,regla}, bordes, sombreado, tabs[{tipo,pos,guia}], mantenerConSiguiente, …
// Las medidas se convierten a puntos al leer. La herencia es la de ODF: estilo predeterminado de la familia → padres con nombre → estilo propio.
// Adaptado de odt-lector.js del sitio (clase Estilos: propiedades, medidas, bordes, niveles de lista); ver docs/motores-adaptados.md.
//
// Requiere xml-editar.js y docx-estilos.js (para mezclar y dar formato a los números de lista).
//
//   const est = OdtEstilos.crear({ contenido: raizContent, estilos: raizStyles|null, settings: textoSettings|null })
//   est.m                       → el adaptador que usa todo el motor (estilos.parrafo(id), caracter(id), mezclar, color, fuente, ajustes…)
//   est.celda(id) / est.columna(id) / est.fila(id) / est.tabla(id) / est.lista(nombre) / est.pagina(nombreMaestra) / est.maestras
(function () {
  'use strict';
  const X = () => globalThis.XmlEditar;
  const MAX_CADENA = 24;

  // ───────── Atributos por nombre local ─────────
  const cacheA = new WeakMap();
  function A(n) {
    if (!n) return Object.create(null);
    let a = cacheA.get(n);
    if (!a) { a = Object.create(null); for (const k of Object.keys(n.attrs)) a[k.slice(k.indexOf(':') + 1)] = n.attrs[k]; cacheA.set(n, a); }
    return a;
  }
  const hijos = (n, local) => (n ? n.hijos.filter((h) => h.local === local) : []);
  const hijo = (n, local) => (n ? n.hijos.find((h) => h.local === local) || null : null);

  // ───────── Medidas ─────────
  // "1.25cm", "0.5in", "12pt", "50%" → puntos (null si no es una medida; los % necesitan `base`)
  function pt(t, base) {
    const m = /^\s*(-?[\d.]+)\s*(cm|mm|in|pt|px|pc|%)?\s*$/.exec(t == null ? '' : String(t));
    if (!m) return null;
    const n = parseFloat(m[1]);
    if (!Number.isFinite(n)) return null;
    const u = m[2] || 'pt';
    if (u === '%') return base == null ? null : base * n / 100;
    return u === 'cm' ? n * 28.3465 : u === 'mm' ? n * 2.83465 : u === 'in' ? n * 72 : u === 'px' ? n * 0.75 : u === 'pc' ? n * 12 : n;
  }
  const acotar = (x, min, max) => Math.min(Math.max(x, min), max);
  const hex = (c) => (/^#[0-9a-f]{6}$/i.test(c || '') ? c.slice(1).toUpperCase() : undefined);

  const RESALTADOS = { FFFF00: 'yellow', '00FF00': 'green', '00FFFF': 'cyan', FF00FF: 'magenta', '0000FF': 'blue', FF0000: 'red', '000080': 'darkBlue', '008080': 'darkCyan', '008000': 'darkGreen', '800080': 'darkMagenta', '800000': 'darkRed', '808000': 'darkYellow', '808080': 'darkGray', C0C0C0: 'lightGray', '000000': 'black' };

  // "0.5pt solid #000000" → { tipo, ancho (pt), color:{val} } | null
  function borde(t) {
    if (!t || t === 'none') return null;
    const partes = String(t).trim().split(/\s+/);
    const ancho = partes.map((p) => pt(p)).find((v) => v !== null && v > 0);
    const color = partes.map(hex).find(Boolean);
    return { tipo: partes.includes('double') ? 'double' : partes.includes('dotted') ? 'dotted' : partes.includes('dashed') ? 'dashed' : 'single', ancho: acotar(ancho == null ? 0.5 : ancho, 0.25, 12), color: color ? { val: color } : undefined };
  }
  const LADOS = { top: 'top', bottom: 'bottom', left: 'left', right: 'right' };
  function bordes(a, claves) {
    let sal;
    const poner = (k, b) => { sal = sal || {}; sal[k] = b || { tipo: 'nil' }; };
    if (a.border !== undefined) for (const k of claves) poner(k, borde(a.border));
    for (const l of Object.keys(LADOS)) if (a['border-' + l] !== undefined) poner(l, borde(a['border-' + l]));
    return sal;
  }
  const limpiar = (o) => { for (const k of Object.keys(o)) if (o[k] === undefined) delete o[k]; return o; };

  // ───────── Propiedades ─────────
  function propsTexto(a, fuentes) {
    const fam = (n, f) => (n !== undefined ? fuentes.get(n) || n : f !== undefined ? String(f).replace(/^["']+|["']+$/g, '') : undefined);
    const r = {};
    const ascii = fam(a['font-name'], a['font-family']), asia = fam(a['font-name-asian'], a['font-family-asian']), cs = fam(a['font-name-complex'], a['font-family-complex']);
    if (ascii !== undefined || asia !== undefined || cs !== undefined) r.fuentes = limpiar({ ascii, hAnsi: ascii, eastAsia: asia, cs });
    if (a['font-size'] !== undefined) {
      if (/%$/.test(a['font-size'])) { const f = parseFloat(a['font-size']) / 100; if (f > 0) r.tamRel = f; }
      else { const v = pt(a['font-size']); if (v !== null && v > 0) r.tam = acotar(v, 1, 1638); }
    }
    if (a['font-weight'] !== undefined) r.b = a['font-weight'] === 'bold' || parseInt(a['font-weight'], 10) >= 600;
    if (a['font-style'] !== undefined) r.i = a['font-style'] === 'italic' || a['font-style'] === 'oblique';
    const us = a['text-underline-style'] !== undefined ? a['text-underline-style'] : a['text-underline-type'];
    if (us !== undefined) r.u = us === 'none' ? 'none' : (a['text-underline-type'] === 'double' ? 'double' : 'single');
    const ls = a['text-line-through-style'] !== undefined ? a['text-line-through-style'] : a['text-line-through-type'];
    if (ls !== undefined) { r.tachado = ls !== 'none'; if (r.tachado && a['text-line-through-type'] === 'double') r.tachadoDoble = true; }
    if (a.color !== undefined) { const c = hex(a.color); if (c) r.color = { val: c }; }
    else if (a['use-window-font-color'] === 'true') r.color = { val: 'auto' };
    if (a['background-color'] !== undefined) {
      const c = hex(a['background-color']);
      if (c) { if (RESALTADOS[c]) r.resaltado = RESALTADOS[c]; else r.sombreado = { relleno: c }; } else if (a['background-color'] === 'transparent') r.resaltado = 'none';
    }
    if (a['text-position'] !== undefined) {
      const m = /^(super|sub|-?[\d.]+%)/.exec(a['text-position']);
      if (m) { const v = m[1]; const sup = v === 'super' || (v.endsWith('%') && parseFloat(v) > 0), sub = v === 'sub' || (v.endsWith('%') && parseFloat(v) < 0); r.alineacionVertical = sup ? 'superscript' : sub ? 'subscript' : 'baseline'; }
    }
    if (a['font-variant'] !== undefined) r.versalitas = a['font-variant'] === 'small-caps';
    if (a['text-transform'] !== undefined) r.mayusculas = a['text-transform'] === 'uppercase';
    if (a.display !== undefined) r.oculto = a.display === 'none';
    if (a['letter-spacing'] !== undefined) { const v = pt(a['letter-spacing']); if (v !== null) r.espaciado = acotar(v, -10, 100); }
    if (a['text-scale'] !== undefined) { const v = parseFloat(a['text-scale']); if (v > 0) r.escala = acotar(v, 1, 600); }
    if (a.language !== undefined && a.language !== 'zxx') r.idioma = a.country && a.country !== 'none' ? a.language + '-' + a.country : a.language;
    return r;
  }

  const ALINEACION = { start: 'left', left: 'left', end: 'right', right: 'right', center: 'center', justify: 'both' };
  function propsParrafo(a, nodo, ctx) {
    const p = {}, ex = {};
    if (a['text-align'] !== undefined && ALINEACION[a['text-align']]) p.jc = ALINEACION[a['text-align']];
    const izq = pt(a['margin-left']), der = pt(a['margin-right']), ti = pt(a['text-indent']);
    if (izq !== null || der !== null || ti !== null) {
      p.sangria = limpiar({ izq: izq === null ? undefined : izq, der: der === null ? undefined : der, primera: ti !== null && ti > 0 ? ti : undefined, colgante: ti !== null && ti < 0 ? -ti : undefined });
      if (ti !== null && ti === 0) { p.sangria.primera = 0; p.sangria.colgante = 0; }
    }
    const antes = pt(a['margin-top']), despues = pt(a['margin-bottom']);
    const lh = a['line-height'], lm = a['line-height-at-least'];
    let linea, regla;
    if (lh !== undefined) { if (/%$/.test(lh)) { const v = parseFloat(lh); if (v > 0) { linea = Math.round(v * 2.4); regla = 'auto'; } } else { const v = pt(lh); if (v !== null && v > 0) { linea = Math.round(v * 20); regla = 'exact'; } } }
    else if (lm !== undefined) { const v = pt(lm); if (v !== null && v > 0) { linea = Math.round(v * 20); regla = 'atLeast'; } }
    if (antes !== null || despues !== null || linea !== undefined) p.espacio = limpiar({ antes: antes === null ? undefined : antes, despues: despues === null ? undefined : despues, linea, regla });
    if (a['break-before'] !== undefined) p.saltoAntes = a['break-before'] === 'page';
    if (a['break-after'] !== undefined && a['break-after'] === 'page') ex.saltoDespues = true;
    if (a['keep-with-next'] !== undefined) p.mantenerConSiguiente = a['keep-with-next'] === 'always';
    if (a['keep-together'] !== undefined) p.mantenerLineas = a['keep-together'] === 'always';
    if (a.widows !== undefined) p.controlViudas = parseInt(a.widows, 10) > 0;
    if (a['writing-mode'] !== undefined) p.derechaAIzquierda = /^rl/.test(a['writing-mode']);
    const fondo = hex(a['background-color']);
    if (fondo) p.sombreado = { relleno: fondo };
    const b = bordes(a, ['top', 'bottom', 'left', 'right']);
    if (b) {
      const pad = pt(a.padding);
      for (const l of Object.keys(b)) {
        const e = a['padding-' + l] !== undefined ? pt(a['padding-' + l]) : pad;
        if (b[l].tipo !== 'nil' && e !== null && e !== undefined) b[l].espacio = e;
      }
      p.bordes = b;
    }
    const tabs = hijo(nodo, 'tab-stops');
    if (tabs) p.tabs = hijos(tabs, 'tab-stop').slice(0, 64).map((t) => { const q = A(t); return { tipo: { left: 'left', center: 'center', right: 'right', char: 'decimal' }[q.type] || 'left', pos: pt(q.position), guia: q['leader-style'] && q['leader-style'] !== 'none' ? ({ dotted: 'dot', dash: 'hyphen', 'long-dash': 'hyphen', solid: 'underscore' }[q['leader-style']] || 'dot') : undefined }; }).filter((t) => t.pos !== null);
    if (ctx) Object.assign(ctx, ex);
    return p;
  }

  function propsCelda(a) {
    const tc = {};
    const f = hex(a['background-color']);
    if (f) tc.sombreado = { relleno: f };
    const b = bordes(a, ['top', 'bottom', 'left', 'right']);
    if (b) tc.bordes = b;
    if (a['vertical-align'] !== undefined) tc.alineacionVertical = { top: 'top', middle: 'center', bottom: 'bottom' }[a['vertical-align']];
    const pad = pt(a.padding);
    const m = {};
    for (const l of ['top', 'bottom', 'left', 'right']) { const v = a['padding-' + l] !== undefined ? pt(a['padding-' + l]) : pad; if (v !== null && v !== undefined) m[l] = v; }
    if (Object.keys(m).length) tc.margenes = m;
    return limpiar(tc);
  }

  // ───────── Entrada ─────────
  function crear({ contenido, estilos, settings }) {
    const mezclar = globalThis.DocxEstilos.mezclar;
    const fuentes = new Map();
    const nombrados = new Map(), automaticos = new Map(), predeterminados = new Map(), listas = new Map(), paginas = new Map(), maestras = new Map();

    const declaraciones = (raiz) => {
      for (const f of hijos(hijo(raiz, 'font-face-decls'), 'font-face')) {
        const a = A(f), fam = String(a['font-family'] || a.name || '').replace(/^["']+|["']+$/g, '');
        if (a.name && fam) fuentes.set(a.name, fam);
      }
    };
    declaraciones(estilos); declaraciones(contenido);

    function leerEstilo(e, destino, parte) {
      const a = A(e), familia = a.family || 'paragraph';
      const est = { nombre: a.name, nombreVisible: a['display-name'] || a.name, familia, padre: a['parent-style-name'] || null, siguiente: a['next-style-name'] || null,
        maestra: a['master-page-name'] || null, nivelEsquema: a['default-outline-level'] ? parseInt(a['default-outline-level'], 10) - 1 : undefined, lista: a['list-style-name'] || null, p: {}, r: {}, tc: {}, ex: {}, auto: destino === automaticos, nodo: e, parte: parte || 'creado' };
      for (const prop of e.hijos) {
        const q = A(prop);
        if (prop.local === 'text-properties') est.r = mezclar(est.r, propsTexto(q, fuentes));
        else if (prop.local === 'paragraph-properties') { est.p = mezclar(est.p, propsParrafo(q, prop, est.ex)); if (q['master-page-name']) est.maestra = q['master-page-name']; }
        else if (prop.local === 'table-cell-properties') est.tc = mezclar(est.tc, propsCelda(q));
        else if (prop.local === 'table-column-properties') { const v = pt(q['column-width']); if (v !== null) est.ancho = v; if (q['rel-column-width']) est.relAncho = parseFloat(q['rel-column-width']); }
        else if (prop.local === 'table-row-properties') { const v = pt(q['row-height']), mn = pt(q['min-row-height']); if (v !== null) { est.alto = v; est.reglaAlto = q['use-optimal-row-height'] === 'true' ? 'atLeast' : 'exact'; } else if (mn !== null) { est.alto = mn; est.reglaAlto = 'atLeast'; } }
        else if (prop.local === 'table-properties') {
          const w = pt(q.width); if (w !== null) est.anchoTabla = w;
          if (q['rel-width']) est.relAnchoTabla = parseFloat(q['rel-width']);
          if (q.align) est.alineaTabla = { left: 'left', center: 'center', right: 'right', margins: 'left' }[q.align];
          const ml = pt(q['margin-left']); if (ml !== null) est.sangriaTabla = ml;
          if (q.align === 'margins') est.sangriaTabla = 0;
        } else if (prop.local === 'section-properties') {
          const c = hijo(prop, 'columns');
          if (c) { const n = parseInt(A(c)['column-count'], 10); if (n > 1) est.columnas = { n: acotar(n, 2, 45), espacio: pt(A(c)['column-gap']) || 36 }; }
          if (q['dont-balance-text-columns']) est.sinEquilibrar = q['dont-balance-text-columns'] === 'true';
        }
      }
      destino.set(familia + '|' + a.name, est);
    }

    function listaDe(e) {
      const niveles = [];
      for (const x of e.hijos) {
        if (!/^list-level-style-/.test(x.local)) continue;
        const q = A(x), n = parseInt(q.level, 10);
        if (!(n >= 1 && n <= 10) || niveles[n - 1]) continue;
        const lp = hijo(x, 'list-level-properties'), la = A(lp), al = hijo(lp, 'list-level-label-alignment'), aa = A(al);
        const nivel = { tipo: x.local.slice('list-level-style-'.length), caracter: q['bullet-char'], formato: q['num-format'], prefijo: q['num-prefix'] || '', sufijo: q['num-suffix'] || '', inicio: parseInt(q['start-value'], 10) || 1,
          niveles: Math.max(1, parseInt(q['display-levels'], 10) || 1), sepa: 'tab', p: {}, r: {} };
        if (al) {
          nivel.sepa = { listtab: 'tab', space: 'space', nothing: 'nothing' }[aa['label-followed-by']] || 'tab';
          const izq = pt(aa['margin-left']), ti = pt(aa['text-indent']);
          nivel.p = { sangria: limpiar({ izq: izq === null ? 0 : izq, primera: ti !== null && ti > 0 ? ti : undefined, colgante: ti !== null && ti < 0 ? -ti : undefined }) };
          const tp = pt(aa['list-tab-stop-position']); if (tp !== null) nivel.p.tabs = [{ tipo: 'left', pos: tp }];
        } else if (lp) {
          const sb = pt(la['space-before']) || 0, mw = pt(la['min-label-width']) || 0;
          nivel.sepa = 'tab';
          nivel.p = { sangria: { izq: sb + mw, colgante: mw } };
        }
        const tp = hijo(x, 'text-properties');
        if (tp) nivel.r = propsTexto(A(tp), fuentes);
        niveles[n - 1] = nivel;
      }
      return niveles;
    }

    function leerContenedor(c, automatico, parte) {
      const destino = automatico ? automaticos : nombrados;
      for (const e of c.hijos) {
        const a = A(e);
        if (e.local === 'style' && a.name) leerEstilo(e, destino, parte);
        else if (e.local === 'default-style' && !automatico) {
          const est = { p: {}, r: {}, ex: {} };
          for (const prop of e.hijos) {
            if (prop.local === 'text-properties') est.r = propsTexto(A(prop), fuentes);
            else if (prop.local === 'paragraph-properties') { est.p = propsParrafo(A(prop), prop, est.ex); const td = pt(A(prop)['tab-stop-distance']); if (td) est.tabPorDefecto = td; }
          }
          predeterminados.set(a.family || 'paragraph', est);
        } else if (e.local === 'list-style' && a.name) listas.set(a.name, listaDe(e));
      }
    }
    if (estilos) {
      leerContenedor(hijo(estilos, 'styles') || { hijos: [] }, false, 'estilos');
      leerContenedor(hijo(estilos, 'automatic-styles') || { hijos: [] }, true, 'estilos');
    }
    leerContenedor(hijo(contenido, 'automatic-styles') || { hijos: [] }, true, 'contenido');
    // los estilos nombrados que algunos programas escriben en el contenido
    if (hijo(contenido, 'styles')) leerContenedor(hijo(contenido, 'styles'), false, 'contenido');

    // Páginas y maestras (styles.xml; en content.xml solo los estilos de página automáticos)
    const lePagina = (pl) => {
      const a = A(hijo(pl, 'page-layout-properties')), cab = A(hijo(hijo(pl, 'header-style'), 'header-footer-properties')), pie = A(hijo(hijo(pl, 'footer-style'), 'header-footer-properties'));
      paginas.set(A(pl).name, { a, cab: hijo(pl, 'header-style') ? cab : null, pie: hijo(pl, 'footer-style') ? pie : null });
    };
    const ordenMaestras = [];
    // (se vuelve a llamar cuando styles.xml cambia: un encabezado o un pie nuevo trae una página y una maestra nuevas)
    function leerPaginas(raizEstilos) {
      paginas.clear(); maestras.clear(); ordenMaestras.length = 0;
      for (const raiz of [raizEstilos, contenido]) for (const pl of hijos(hijo(raiz, 'automatic-styles'), 'page-layout')) lePagina(pl);
      if (raizEstilos) for (const mp of hijos(hijo(raizEstilos, 'master-styles'), 'master-page')) { maestras.set(A(mp).name, mp); ordenMaestras.push(A(mp).name); }
    }
    leerPaginas(estilos);

    // ───────── Herencia ─────────
    const buscar = (familia, nombre) => (nombre ? automaticos.get(familia + '|' + nombre) || nombrados.get(familia + '|' + nombre) || null : null);
    function cadena(familia, nombre) {
      const l = [], visto = new Set();
      for (let e = buscar(familia, nombre); e && !visto.has(e.nombre) && l.length < MAX_CADENA; e = e.padre ? nombrados.get(familia + '|' + e.padre) || null : null) { visto.add(e.nombre); l.unshift(e); }
      return l;
    }
    const pred = (f) => predeterminados.get(f) || { p: {}, r: {}, ex: {} };
    // el tamaño en % es relativo al del padre: se resuelve a lo largo de la cadena
    function sumarR(base, r) {
      const sal = mezclar(base, r);
      delete sal.tamRel;
      if (r.tamRel && base.tam) sal.tam = Math.round(base.tam * r.tamRel * 100) / 100;
      return sal;
    }
    const cacheP = new Map(), cacheC = new Map();
    function resolverParrafo(nombre) {
      if (cacheP.has(nombre || '')) return cacheP.get(nombre || '');
      const d = pred('paragraph'), dt = pred('text');
      let p = mezclar({}, d.p), r = sumarR(sumarR({ tam: 12 }, dt.r), d.r), ex = { ...d.ex }, lista = null, maestra = null, nivelEsquema;
      for (const e of cadena('paragraph', nombre)) {
        p = mezclar(p, e.p); r = sumarR(r, e.r); Object.assign(ex, e.ex);
        if (e.lista) lista = e.lista;
        if (e.maestra) maestra = e.maestra;
        if (e.nivelEsquema !== undefined) nivelEsquema = e.nivelEsquema;
      }
      if (nivelEsquema !== undefined) p.nivelEsquema = nivelEsquema;
      const sal = { p, r, ex, lista, maestra };
      cacheP.set(nombre || '', sal);
      return sal;
    }
    function resolverCaracter(nombre) {
      if (cacheC.has(nombre || '')) return cacheC.get(nombre || '');
      let r = {};
      for (const e of cadena('text', nombre)) r = mezclar(r, e.r);
      cacheC.set(nombre || '', r);
      return r;
    }
    function simple(familia, nombre) {
      const sal = {};
      for (const e of cadena(familia, nombre)) {
        for (const k of ['tc', 'ancho', 'relAncho', 'alto', 'reglaAlto', 'anchoTabla', 'relAnchoTabla', 'alineaTabla', 'sangriaTabla', 'columnas']) if (e[k] !== undefined) sal[k] = k === 'tc' ? mezclar(sal.tc, e.tc) : e[k];
      }
      return sal;
    }

    // ───────── Ajustes ─────────
    const configuracion = (nombre) => { const m = settings ? new RegExp('config-item[^>]*config:name="' + nombre + '"[^>]*>\\s*(true|false)', 'i').exec(settings) : null; return m ? m[1].toLowerCase() === 'true' : undefined; };
    const tabsRel = configuracion('TabsRelativeToIndent');
    const dp = pred('paragraph');
    const ajustes = { tabPorDefecto: dp.tabPorDefecto || 35.4, modoCompat: 15, parImpar: false, tabsRelativos: tabsRel === undefined ? true : tabsRel, perfil: 'odf' };
    for (const mp of maestras.values()) if (hijo(mp, 'header-left') || hijo(mp, 'footer-left')) ajustes.parImpar = true;

    const defecto = { p: mezclar({}, dp.p), r: sumarR(sumarR({ tam: 12 }, pred('text').r), dp.r) };
    const FUENTE0 = defecto.r.fuentes || { ascii: 'Liberation Serif', hAnsi: 'Liberation Serif', eastAsia: 'Liberation Serif', cs: 'Liberation Serif' };
    defecto.r.fuentes = FUENTE0;
    const m = {
      mezclar, ajustes, tema: null,
      color: (c) => (c && c.val && /^[0-9a-fA-F]{6}$/.test(c.val) ? '#' + c.val.toLowerCase() : null),
      fuente: (f, tipo = 'ascii') => (f ? f[tipo] || f.ascii || f.hAnsi || null : null),
      estilos: {
        existe: (id) => !!buscar('paragraph', id) || !!buscar('text', id),
        nombre: (id) => { for (const e of cadena('paragraph', id).reverse()) if (!e.auto) return e.nombreVisible; return ''; },
        siguiente: (id) => { const e = buscar('paragraph', id); return e ? e.siguiente : null; },
        parrafo(id, medio) {
          const e = resolverParrafo(id);
          return { p: mezclar(mezclar({}, medio && medio.p), e.p), r: mezclar(mezclar({ fuentes: FUENTE0 }, medio && medio.r), e.r) };
        },
        caracter: (id) => (id ? resolverCaracter(id) : {}),
        tabla: () => ({ p: {}, r: {}, tbl: {}, tc: {}, tr: {}, condicionales: {} }),
        base: defecto,
      },
    };
    return {
      m, ajustes, defecto, resolverParrafo, resolverCaracter, sumarR,
      celda: (id) => simple('table-cell', id), columna: (id) => simple('table-column', id), fila: (id) => simple('table-row', id), tabla: (id) => simple('table', id),
      seccion: (id) => simple('section', id),
      lista: (nombre) => listas.get(nombre) || null,
      pagina: (nombre) => paginas.get(nombre) || null,
      maestras, ordenMaestras, buscar,
      estiloAuto: (familia, nombre) => automaticos.get(familia + '|' + nombre) || null,
      get primeraMaestra() { return ordenMaestras[0] || null; },
      releerPaginas: leerPaginas,
      // estilos que se crean al editar (odt-formato.js): se leen igual que los del archivo; `xml` es un solo <style:style> o <text:list-style>
      registrar(xml, nombrado) {
        const r = X().analizar(xml);
        if (r.local === 'list-style') { listas.set(A(r).name, listaDe(r)); return; }
        leerEstilo(r, nombrado ? nombrados : automaticos);
        cacheP.clear(); cacheC.clear();
      },
      nombrados: (familia) => [...nombrados.values()].filter((e) => e.familia === familia),
      automaticosDe: (familia) => [...automaticos.values()].filter((e) => e.familia === familia && e.parte === 'contenido'),
      nombrado: (familia, nombre) => nombrados.get(familia + '|' + nombre) || null,
      predeterminado: (familia) => predeterminados.get(familia) || null,
    };
  }

  globalThis.OdtEstilos = { crear, A, pt, hex, borde };
})();

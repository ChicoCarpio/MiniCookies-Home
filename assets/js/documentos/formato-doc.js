/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Formato de Documentos (Fase 6, etapa 6.5): cambios de formato de caracteres y de párrafo sobre el XML de un párrafo. Sin DOM: corre igual en Node.
//
// Igual que edicion-doc.js, no reescribe nada desde el modelo: cambia solo el elemento tocado dentro del XML del párrafo (<w:rPr> de una
// corrida, <w:pPr> del párrafo), en el orden que pide el esquema (Word se queja de un <w:rPr> con los hijos desordenados), y parte las
// corridas solo donde la selección las corta. Lo demás del párrafo queda byte a byte.
//
//   FormatoDoc.elementosCaracter({ b:true, color:'FF0000', tam:14 }, 'w:') → Map local → XML del elemento | null (quitar)
//   FormatoDoc.edicionesCaracter(xml, raiz, piezas, desde, hasta, elementos) → ediciones para XmlEditar.aplicar
//   FormatoDoc.ponerPPr(xml, cambios) → XML del párrafo con esos cambios de párrafo
//   FormatoDoc.ponerMarca(xml, elementos) → lo mismo para el <w:rPr> de la marca de párrafo (un párrafo vacío escribe con ese formato)
//   FormatoDoc.estiloNuevo(id) → XML de un estilo de párrafo predefinido (Título 1, Título…) o null
//   FormatoDoc.numeracionNueva(tipo) → XML de un <w:abstractNum> de viñetas o de números
//
// Requiere xml-editar.js.
(function () {
  'use strict';
  const X = () => globalThis.XmlEditar;
  const c16 = String.fromCharCode;

  // Orden de los hijos de <w:rPr> y de <w:pPr> en el esquema de Word (CT_RPr y CT_PPr)
  const ORDEN_RPR = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof',
    'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText',
    'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'];
  const ORDEN_PPR = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs',
    'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid',
    'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl',
    'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
  const rango = (orden, local) => { const i = orden.indexOf(local); return i < 0 ? 1e9 : i; };
  const prefijoDe = (raiz) => { const i = raiz.nombre.indexOf(':'); return i < 0 ? '' : raiz.nombre.slice(0, i + 1); };
  const hijoDe = (n, local) => (n ? n.hijos.find((h) => h.local === local) || null : null);
  const acotar = (v, a, b) => Math.min(Math.max(v, a), b);
  const tw = (pt) => Math.round(pt * 20);
  // pares sustitutos (emoji y otras escrituras): un corte nunca cae en medio
  const ALTO = new RegExp('[' + c16(0xD800) + '-' + c16(0xDBFF) + ']'), BAJO = new RegExp('[' + c16(0xDC00) + '-' + c16(0xDFFF) + ']');

  // ───────── Propiedades ─────────
  // Cambia (o quita, con null) hijos de un <w:rPr>/<w:pPr> sin mover los demás. `xmlPr` es el elemento tal como está ('' si no hay).
  function conElementos(xmlPr, nombre, orden, elementos, pf) {
    let abre = `<${pf}${nombre}>`, cierra = `</${pf}${nombre}>`;
    const kids = [];
    if (xmlPr) {
      const n = X().analizar(xmlPr);
      if (n.iniCierre >= 0) abre = xmlPr.slice(n.ini, n.finApertura);
      for (const h of n.hijos) kids.push({ local: h.local, xml: xmlPr.slice(h.ini, h.fin) });
      cierra = `</${n.nombre}>`;
    }
    for (const [local, nuevo] of elementos) {
      const donde = kids.findIndex((k) => k.local === local);
      for (let i = kids.length - 1; i > donde && donde >= 0; i--) if (kids[i].local === local) kids.splice(i, 1);   // repetidos
      if (nuevo === null) { for (let i = kids.length - 1; i >= 0; i--) if (kids[i].local === local) kids.splice(i, 1); continue; }
      if (donde >= 0) { kids[donde].xml = nuevo; continue; }
      const r = rango(orden, local);
      let pos = kids.findIndex((k) => rango(orden, k.local) > r);
      if (pos < 0) pos = kids.length;
      kids.splice(pos, 0, { local, xml: nuevo });
    }
    // los elementos que se tocan salen en el orden del esquema aunque el original viniera desordenado (el orden estable deja igual lo que ya lo estaba)
    kids.forEach((k, i) => { k.i = i; });
    kids.sort((p, q) => rango(orden, p.local) - rango(orden, q.local) || p.i - q.i);
    return kids.length ? abre + kids.map((k) => k.xml).join('') + cierra : '';
  }
  // Un elemento con los atributos cambiados (null quita un atributo); `existente` es el elemento actual o ''
  function elementoConAtributos(existente, local, cambios, pf, quitar = []) {
    const attrs = [];
    if (existente) {
      const n = X().analizar(existente);
      for (const k of Object.keys(n.attrs)) attrs.push([k, n.attrs[k]]);
    }
    const sinPrefijo = (k) => k.slice(k.indexOf(':') + 1);
    for (const q of quitar) for (let i = attrs.length - 1; i >= 0; i--) if (sinPrefijo(attrs[i][0]) === q) attrs.splice(i, 1);
    for (const k of Object.keys(cambios)) {
      const i = attrs.findIndex((a) => sinPrefijo(a[0]) === k);
      if (cambios[k] === null) { if (i >= 0) attrs.splice(i, 1); continue; }
      if (i >= 0) attrs[i][1] = String(cambios[k]); else attrs.push([pf + k, String(cambios[k])]);
    }
    return `<${pf}${local}${attrs.map(([k, v]) => ` ${k}="${X().escAttr(v)}"`).join('')}/>`;
  }

  // Cambios de formato de caracteres → elementos de <w:rPr>
  //   b, i, tachado: true/false · u: 'single'|'none'|… · color: 'RRGGBB'|'auto' · tam: puntos · fuente: nombre · resaltado: 'yellow'…|'none'
  //   vertical: 'superscript'|'subscript'|'baseline' · limpiar: true (quita todo el formato directo menos el estilo de carácter)
  function elementosCaracter(c, pf) {
    const m = new Map();
    const interruptor = (local, v) => `<${pf}${local}${v ? '' : ` ${pf}val="0"`}/>`;
    if (c.limpiar) for (const l of ORDEN_RPR) if (l !== 'rStyle') m.set(l, null);
    if ('b' in c) { m.set('b', interruptor('b', c.b)); m.set('bCs', interruptor('bCs', c.b)); }
    if ('i' in c) { m.set('i', interruptor('i', c.i)); m.set('iCs', interruptor('iCs', c.i)); }
    if ('tachado' in c) m.set('strike', interruptor('strike', c.tachado));
    if ('u' in c) m.set('u', `<${pf}u ${pf}val="${X().escAttr(c.u === true ? 'single' : c.u || 'none')}"/>`);
    if ('color' in c) m.set('color', c.color === null ? null : `<${pf}color ${pf}val="${/^[0-9a-fA-F]{6}$/.test(c.color) ? c.color.toUpperCase() : 'auto'}"/>`);
    if ('tam' in c) { const v = acotar(Math.round(c.tam * 2), 2, 3276); m.set('sz', `<${pf}sz ${pf}val="${v}"/>`); m.set('szCs', `<${pf}szCs ${pf}val="${v}"/>`); }
    if ('fuente' in c) {
      const f = X().escAttr(c.fuente);
      m.set('rFonts', `<${pf}rFonts ${pf}ascii="${f}" ${pf}hAnsi="${f}" ${pf}eastAsia="${f}" ${pf}cs="${f}"/>`);
    }
    if ('resaltado' in c) m.set('highlight', !c.resaltado || c.resaltado === 'none' ? null : `<${pf}highlight ${pf}val="${X().escAttr(c.resaltado)}"/>`);
    if ('vertical' in c) m.set('vertAlign', !c.vertical || c.vertical === 'baseline' ? null : `<${pf}vertAlign ${pf}val="${c.vertical === 'subscript' ? 'subscript' : 'superscript'}"/>`);
    return m;
  }

  // ───────── Corridas ─────────
  function nodos(raiz) {
    const mapa = new Map();
    const rec = (n) => { for (const h of n.hijos) { mapa.set(h.ini, h); rec(h); } };
    rec(raiz);
    return mapa;
  }
  const textoT = (z, s) => `<${z.tn} xml:space="preserve">${X().esc(s)}</${z.tn}>`;

  // Ediciones que aplican `elementos` al formato de los caracteres [desde, hasta) de un párrafo. Una corrida que cae entera dentro cambia su
  // propio <w:rPr>; una que la selección corta se parte en dos o tres corridas con el mismo formato salvo en el trozo tocado.
  function edicionesCaracter(xml, raiz, piezas, desde, hasta, elementos) {
    const pf = prefijoDe(raiz), mapa = nodos(raiz), ed = [];
    const corridas = new Map();
    for (const z of piezas) {
      if (z.hasta <= z.desde || z.hasta <= desde || z.desde >= hasta) continue;
      const run = mapa.get(z.ini);
      if (run && run.local === 'r') corridas.set(run.ini, run);
    }
    for (const run of corridas.values()) {
      const delRun = piezas.filter((z) => z.ini === run.ini);
      const rPr = hijoDe(run, 'rPr'), rPrXml = rPr ? xml.slice(rPr.ini, rPr.fin) : '';
      const partes = [];
      for (const k of run.hijos) {
        if (k.local === 'rPr') continue;
        const z = delRun.find((q) => q.ei === k.ini);
        const original = xml.slice(k.ini, k.fin);
        if (!z || z.hasta <= z.desde) { partes.push({ xml: original, en: null }); continue; }
        const a = Math.max(desde, z.desde), b = Math.min(hasta, z.hasta);
        if (b <= a) partes.push({ xml: original, en: false });
        else if (a <= z.desde && b >= z.hasta) partes.push({ xml: original, en: true });
        else if (z.t === 'texto' && z.ti !== undefined && !z.nota) {
          // un corte nunca cae en medio de un par sustituto (un emoji): el trozo se agranda para llevarse el par entero
          const mitad = (k) => k > 0 && k < z.s.length && ALTO.test(z.s[k - 1]) && BAJO.test(z.s[k]);
          let k1 = a - z.desde, k2 = b - z.desde;
          if (mitad(k1)) k1--;
          if (mitad(k2)) k2++;
          if (k1 > 0) partes.push({ xml: textoT(z, z.s.slice(0, k1)), en: false });
          partes.push({ xml: textoT(z, z.s.slice(k1, k2)), en: true });
          if (k2 < z.s.length) partes.push({ xml: textoT(z, z.s.slice(k2)), en: false });
        } else partes.push({ xml: original, en: true });
      }
      // lo que no tiene texto (marcas, saltos sin ancho) se queda con lo que tiene al lado
      let previo = null;
      for (const p of partes) { if (p.en === null) p.en = previo; else previo = p.en; }
      for (let i = partes.length - 1, sig = null; i >= 0; i--) { if (partes[i].en === null) partes[i].en = sig; else sig = partes[i].en; }
      if (partes.every((p) => p.en === null)) continue;
      const nuevo = conElementos(rPrXml, 'rPr', ORDEN_RPR, elementos, pf);
      if (partes.every((p) => p.en)) {
        if (rPr) ed.push({ ini: rPr.ini, fin: rPr.fin, texto: nuevo });
        else if (nuevo) ed.push({ ini: run.finApertura, fin: run.finApertura, texto: nuevo });
        continue;
      }
      const abre = xml.slice(run.ini, run.finApertura), cierra = `</${run.nombre}>`;
      const trozos = [];
      for (const p of partes) {
        const ult = trozos[trozos.length - 1];
        if (ult && ult.en === !!p.en) ult.xml += p.xml; else trozos.push({ en: !!p.en, xml: p.xml });
      }
      ed.push({ ini: run.ini, fin: run.fin, texto: trozos.map((t) => abre + (t.en ? nuevo : rPrXml) + t.xml + cierra).join('') });
    }
    return ed;
  }

  // ───────── Párrafo ─────────
  // Cambios de formato de párrafo
  //   jc: 'left'|'center'|'right'|'both' (ya en el sentido del archivo) · sangria: { izq, der, primera, colgante } en pt (null en un campo lo quita)
  //   espacio: { antes, despues, linea, regla } (pt; linea en 240avos si regla es 'auto') · numeracion: { numId, nivel } | null · saltoAntes, mantenerConSiguiente
  function ponerPPr(xml, c) {
    const raiz = X().analizar(xml), pf = prefijoDe(raiz), pPr = hijoDe(raiz, 'pPr');
    const actual = (local) => { const h = hijoDe(pPr, local); return h ? xml.slice(h.ini, h.fin) : ''; };
    const m = new Map();
    if ('jc' in c) m.set('jc', c.jc ? `<${pf}jc ${pf}val="${X().escAttr(c.jc)}"/>` : null);
    if (c.sangria) {
      const s = c.sangria, a = {}, quitar = [];
      if ('izq' in s) { a.left = s.izq === null ? null : tw(Math.max(s.izq, 0)); quitar.push('start'); }
      if ('der' in s) { a.right = s.der === null ? null : tw(Math.max(s.der, 0)); quitar.push('end'); }
      if ('primera' in s) { a.firstLine = s.primera === null ? null : tw(s.primera); a.hanging = null; }
      if ('colgante' in s) { a.hanging = s.colgante === null ? null : tw(s.colgante); a.firstLine = null; }
      const nuevo = elementoConAtributos(actual('ind'), 'ind', a, pf, quitar);
      m.set('ind', /^<[\w:]+\s*\/>$/.test(nuevo) ? null : nuevo);
    }
    if (c.espacio) {
      const s = c.espacio, a = {}, quitar = [];
      if ('antes' in s) { a.before = s.antes === null ? null : tw(s.antes); quitar.push('beforeAutospacing', 'beforeLines'); }
      if ('despues' in s) { a.after = s.despues === null ? null : tw(s.despues); quitar.push('afterAutospacing', 'afterLines'); }
      if ('linea' in s) { a.line = s.linea === null ? null : Math.round(s.linea); a.lineRule = s.linea === null ? null : s.regla || 'auto'; }
      const nuevo = elementoConAtributos(actual('spacing'), 'spacing', a, pf, quitar);
      m.set('spacing', /^<[\w:]+\s*\/>$/.test(nuevo) ? null : nuevo);
    }
    if ('numeracion' in c) {
      m.set('numPr', c.numeracion ? `<${pf}numPr><${pf}ilvl ${pf}val="${c.numeracion.nivel || 0}"/><${pf}numId ${pf}val="${c.numeracion.numId}"/></${pf}numPr>` : null);
    }
    if ('saltoAntes' in c) m.set('pageBreakBefore', c.saltoAntes ? `<${pf}pageBreakBefore/>` : null);
    if ('mantenerConSiguiente' in c) m.set('keepNext', c.mantenerConSiguiente ? `<${pf}keepNext/>` : null);
    const nuevo = conElementos(pPr ? xml.slice(pPr.ini, pPr.fin) : '', 'pPr', ORDEN_PPR, m, pf);
    if (pPr) return X().aplicar(xml, [{ ini: pPr.ini, fin: pPr.fin, texto: nuevo }]);
    return nuevo ? X().aplicar(xml, [{ ini: raiz.finApertura, fin: raiz.finApertura, texto: nuevo }]) : xml;
  }
  // El formato de la marca de párrafo (<w:pPr><w:rPr>), que un párrafo vacío usa para lo que se escriba en él
  function ponerMarca(xml, elementos) {
    const raiz = X().analizar(xml), pf = prefijoDe(raiz), pPr = hijoDe(raiz, 'pPr'), rPr = hijoDe(pPr, 'rPr');
    const nuevo = conElementos(rPr ? xml.slice(rPr.ini, rPr.fin) : '', 'rPr', ORDEN_RPR, elementos, pf);
    if (!nuevo && !rPr) return xml;
    if (rPr) return X().aplicar(xml, [{ ini: rPr.ini, fin: rPr.fin, texto: nuevo }]);
    const p = conElementos(pPr ? xml.slice(pPr.ini, pPr.fin) : '', 'pPr', ORDEN_PPR, new Map([['rPr', nuevo]]), pf);
    return pPr ? X().aplicar(xml, [{ ini: pPr.ini, fin: pPr.fin, texto: p }]) : X().aplicar(xml, [{ ini: raiz.finApertura, fin: raiz.finApertura, texto: p }]);
  }

  // ───────── Secciones ─────────
  const ORDEN_SECTPR = ['headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType',
    'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange'];
  // Un <w:sectPr> con otro tipo de inicio de sección ('nextPage' | 'continuous')
  function sectPrConTipo(sectPrXml, tipo, pf) {
    return conElementos(sectPrXml, 'sectPr', ORDEN_SECTPR, new Map([['type', tipo === 'nextPage' ? null : `<${pf}type ${pf}val="${tipo}"/>`]]), pf);
  }
  // El <w:sectPr> de un párrafo (XML del párrafo) cambiado por la función dada, o puesto si no tenía
  function ponerSectPr(xml, sectPrXml) {
    const raiz = X().analizar(xml), pf = prefijoDe(raiz), pPr = hijoDe(raiz, 'pPr');
    const nuevo = conElementos(pPr ? xml.slice(pPr.ini, pPr.fin) : '', 'pPr', ORDEN_PPR, new Map([['sectPr', sectPrXml]]), pf);
    return pPr ? X().aplicar(xml, [{ ini: pPr.ini, fin: pPr.fin, texto: nuevo }]) : X().aplicar(xml, [{ ini: raiz.finApertura, fin: raiz.finApertura, texto: nuevo }]);
  }
  function sectPrDe(xml) {
    const raiz = X().analizar(xml), s = hijoDe(hijoDe(raiz, 'pPr'), 'sectPr');
    return s ? xml.slice(s.ini, s.fin) : null;
  }
  function ponerTipoSeccion(xml, tipo) {
    const raiz = X().analizar(xml), s = hijoDe(hijoDe(raiz, 'pPr'), 'sectPr');
    return s ? ponerSectPr(xml, sectPrConTipo(xml.slice(s.ini, s.fin), tipo, prefijoDe(raiz))) : xml;
  }

  // ───────── Estilos y listas predefinidos ─────────
  const FUENTE_TITULO = '<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>';
  const AZUL = '<w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>';
  const ESTILOS = {
    Heading1: { nombre: 'heading 1', es: 'Título 1', cuerpo: '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr>' + FUENTE_TITULO + AZUL + '<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>' },
    Heading2: { nombre: 'heading 2', es: 'Título 2', cuerpo: '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr>' + FUENTE_TITULO + AZUL + '<w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>' },
    Heading3: { nombre: 'heading 3', es: 'Título 3', cuerpo: '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr>' + FUENTE_TITULO + '<w:color w:val="1F3763" w:themeColor="accent1" w:themeShade="7F"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>' },
    Title: { nombre: 'Title', es: 'Título', cuerpo: '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr><w:rPr>' + FUENTE_TITULO + '<w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr>' },
    Subtitle: { nombre: 'Subtitle', es: 'Subtítulo', cuerpo: '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="11"/><w:qFormat/><w:pPr><w:spacing w:after="160"/></w:pPr><w:rPr><w:color w:val="5A5A5A" w:themeColor="text1" w:themeTint="A5"/><w:spacing w:val="15"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>' },
    ListParagraph: { nombre: 'List Paragraph', es: 'Párrafo de lista', cuerpo: '<w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr>' },
  };
  function estiloNuevo(id) {
    const e = ESTILOS[id];
    return e ? `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${e.nombre}"/>${e.cuerpo}</w:style>` : null;
  }
  const nombreEs = (id) => (ESTILOS[id] ? ESTILOS[id].es : null);

  // Un <w:abstractNum> de 9 niveles: viñetas (como las de Word: Symbol, «o» en Courier New, Wingdings) o números (1. a. i. …)
  function numeracionNueva(tipo, id) {
    let niveles = '';
    const REGLA = ['decimal', 'lowerLetter', 'lowerRoman'];
    const VINETA = [[c16(0xF0B7), 'Symbol'], ['o', 'Courier New'], [c16(0xF0A7), 'Wingdings']];
    for (let k = 0; k < 9; k++) {
      const sangria = `<w:pPr><w:ind w:left="${720 * (k + 1)}" w:hanging="360"/></w:pPr>`;
      if (tipo === 'vinetas') {
        const [car, fuente] = VINETA[k % 3];
        const txt = car === 'o' ? 'o' : `&#x${car.charCodeAt(0).toString(16).toUpperCase()};`;
        niveles += `<w:lvl w:ilvl="${k}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${txt}"/><w:lvlJc w:val="left"/>${sangria}<w:rPr><w:rFonts w:ascii="${fuente}" w:hAnsi="${fuente}" w:hint="default"/></w:rPr></w:lvl>`;
      } else {
        niveles += `<w:lvl w:ilvl="${k}"><w:start w:val="1"/><w:numFmt w:val="${REGLA[k % 3]}"/><w:lvlText w:val="%${k + 1}."/><w:lvlJc w:val="${k % 3 === 2 ? 'right' : 'left'}"/>${k % 3 === 2 ? `<w:pPr><w:ind w:left="${720 * (k + 1)}" w:hanging="180"/></w:pPr>` : sangria}</w:lvl>`;
      }
    }
    return `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${niveles}</w:abstractNum>`;
  }
  const NUMERACION_VACIA = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:numbering>';

  globalThis.FormatoDoc = {
    ORDEN_RPR, ORDEN_PPR, conElementos, elementoConAtributos, elementosCaracter, edicionesCaracter, ponerPPr, ponerMarca,
    sectPrConTipo, ponerSectPr, sectPrDe, ponerTipoSeccion, estiloNuevo, nombreEs, ESTILOS_PREDEFINIDOS: Object.keys(ESTILOS), numeracionNueva, NUMERACION_VACIA,
  };
})();

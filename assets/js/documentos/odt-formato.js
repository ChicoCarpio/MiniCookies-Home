/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Formato de un .odt para EdicionDoc (Fase 6b, etapa 6b.4): formato de caracteres y de párrafo, y estilos con nombre. Sin DOM.
//
// En ODF el formato directo vive en ESTILOS AUTOMÁTICOS (`text:style-name` de un `text:span` o de un párrafo). Reglas de este módulo:
//   · nunca se modifica un estilo que ya está en el archivo: se deriva uno nuevo (`MCT<n>` de texto, `MCP<n>` de párrafo) y se cambia la referencia;
//   · un estilo con las mismas propiedades se reutiliza (los del archivo y los ya creados): negrita y quitar la negrita devuelven el texto a como estaba;
//   · solo se escriben los estilos creados que el documento final usa (`finalizar`, `alGuardar`): deshacer todo y guardar da el mismo paquete;
//   · un `text:span` que cubre justo lo seleccionado cambia de estilo (o se quita si ya no tiene formato); si la selección es un trozo, se mete un span
//     ANIDADO con solo lo que cambia respecto de lo que ya tiene el texto (los spans anidados suman sus propiedades, también en LibreOffice).
// Los estilos con nombre que faltan (Título, Subtítulo, Título 1 a 3) se crean en styles.xml al guardar, solo si el documento los usa.
//
// Requiere xml-editar.js y odt-estilos.js. `crear(modelo, { codificar, Rechazo, prefijoDe })` lo llama OdtParrafo.
(function () {
  'use strict';
  const X = () => globalThis.XmlEditar;
  const E = () => globalThis.OdtEstilos;
  const A = (n) => E().A(n);

  const URI = { style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0', fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0', text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0' };
  const FO = new Set(['font-weight', 'font-style', 'font-size', 'color', 'font-family', 'background-color', 'text-align', 'margin-left', 'margin-right', 'margin-top', 'margin-bottom', 'text-indent', 'line-height', 'break-before', 'keep-with-next']);
  const nsDe = (l) => (FO.has(l) ? 'fo' : 'style');
  const local = (k) => k.slice(k.indexOf(':') + 1);
  const redondo = (v) => Math.round(v * 100) / 100;
  const pt = (v) => redondo(v) + 'pt';
  const NOMBRE_COLOR = { yellow: 'FFFF00', green: '00FF00', cyan: '00FFFF', magenta: 'FF00FF', blue: '0000FF', red: 'FF0000', darkBlue: '000080', darkCyan: '008080', darkGreen: '008000', darkMagenta: '800080', darkRed: '800000', darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0', black: '000000' };
  const ordenado = (m) => [...m.entries()].sort((p, q) => (p[0] < q[0] ? -1 : p[0] > q[0] ? 1 : 0));

  // ───────── Cambios de la suite → atributos de ODF ─────────
  // Cada clave de `cambios` (las mismas que entiende FormatoDoc.elementosCaracter) da pares [atributo, valor|null]; null quita el atributo.
  const PARES_TEXTO = {
    b: (v) => { const x = v ? 'bold' : 'normal'; return [['font-weight', x], ['font-weight-asian', x], ['font-weight-complex', x]]; },
    i: (v) => { const x = v ? 'italic' : 'normal'; return [['font-style', x], ['font-style-asian', x], ['font-style-complex', x]]; },
    tachado: (v) => (v ? [['text-line-through-style', 'solid'], ['text-line-through-type', 'single']] : [['text-line-through-style', 'none'], ['text-line-through-type', null]]),
    u: (v) => {
      const u = v === true ? 'single' : v || 'none';
      if (u === 'none') return [['text-underline-style', 'none'], ['text-underline-type', null], ['text-underline-width', null], ['text-underline-color', null]];
      return [['text-underline-style', 'solid'], ['text-underline-type', u === 'double' ? 'double' : 'single'], ['text-underline-width', 'auto'], ['text-underline-color', 'font-color']];
    },
    color: (v) => (/^[0-9a-f]{6}$/i.test(v || '') ? [['color', '#' + v.toUpperCase()], ['use-window-font-color', null]] : [['color', null], ['use-window-font-color', 'true']]),
    tam: (v) => { const x = pt(Math.min(Math.max(v, 1), 1638)); return [['font-size', x], ['font-size-asian', x], ['font-size-complex', x]]; },
    fuente: (v) => {
      const f = /^[\w-]+$/.test(v) ? v : "'" + String(v).replace(/'/g, '') + "'";
      return [['font-name', null], ['font-name-asian', null], ['font-name-complex', null], ['font-family', f], ['font-family-asian', f], ['font-family-complex', f]];
    },
    resaltado: (v) => [['background-color', !v || v === 'none' ? 'transparent' : '#' + (NOMBRE_COLOR[v] || 'FFFF00')]],
    vertical: (v) => [['text-position', v === 'superscript' ? 'super 58%' : v === 'subscript' ? 'sub 58%' : '0% 100%']],
  };
  // el valor «normalizado» de un cambio y el de un formato efectivo (para saber si un cambio ya está en el texto)
  const normalizar = {
    b: (v) => !!v, i: (v) => !!v, tachado: (v) => !!v,
    u: (v) => (v === true ? 'single' : !v || v === 'none' ? 'none' : v === 'double' ? 'double' : 'single'),
    color: (v) => (/^[0-9a-f]{6}$/i.test(v || '') ? v.toUpperCase() : 'auto'),
    tam: (v) => redondo(v), fuente: (v) => String(v), resaltado: (v) => (!v || v === 'none' ? 'none' : v), vertical: (v) => (!v || v === 'baseline' ? 'baseline' : v),
  };
  function efectivo(r) {
    return {
      b: !!r.b, i: !!r.i, tachado: !!r.tachado, u: r.u && r.u !== 'none' ? (r.u === 'double' ? 'double' : 'single') : 'none',
      color: r.color && r.color.val && /^[0-9a-f]{6}$/i.test(r.color.val) ? r.color.val.toUpperCase() : 'auto', tam: r.tam === undefined ? undefined : redondo(r.tam),
      fuente: r.fuentes && r.fuentes.ascii ? String(r.fuentes.ascii) : undefined, resaltado: r.resaltado || 'none', vertical: r.alineacionVertical || 'baseline',
    };
  }
  // El formato de un texto «sin formato directo»: lo que trae el estilo del párrafo
  function limpio(r) {
    const e = efectivo(r), sal = {};
    for (const k of Object.keys(e)) if (e[k] !== undefined) sal[k] = k === 'color' ? (e.color === 'auto' ? null : e.color) : e[k];
    return sal;
  }
  function paresTexto(c, ef) {
    const sal = [];
    for (const k of Object.keys(c)) {
      if (!PARES_TEXTO[k] || c[k] === undefined) continue;
      const v = k === 'color' && c[k] === null ? 'auto' : c[k];
      const igual = ef && ef[k] !== undefined && normalizar[k](v) === ef[k];
      for (const [l, x] of PARES_TEXTO[k](k === 'color' ? (v === 'auto' ? null : v) : v)) sal.push([l, igual ? null : x]);
    }
    return sal;
  }
  const ALINEAR = { left: 'start', right: 'end', center: 'center', both: 'justify', start: 'start', end: 'end' };
  function paresParrafo(c, ref) {
    const sal = [];
    if ('jc' in c) sal.push(['text-align', c.jc && !(ref && (ref.jc || 'left') === c.jc) ? ALINEAR[c.jc] || null : null]);
    if (c.sangria) {
      const s = c.sangria;
      if ('izq' in s) sal.push(['margin-left', s.izq === null ? null : pt(Math.max(s.izq, 0))]);
      if ('der' in s) sal.push(['margin-right', s.der === null ? null : pt(Math.max(s.der, 0))]);
      if ('primera' in s) sal.push(['text-indent', s.primera === null ? null : pt(s.primera)]);
      if ('colgante' in s) sal.push(['text-indent', s.colgante === null ? null : pt(-s.colgante)]);
    }
    if (c.espacio) {
      const s = c.espacio;
      if ('antes' in s) sal.push(['margin-top', s.antes === null ? null : pt(s.antes)]);
      if ('despues' in s) sal.push(['margin-bottom', s.despues === null ? null : pt(s.despues)]);
      if ('linea' in s) {
        const regla = s.regla || 'auto';
        sal.push(['line-height', s.linea === null || regla === 'atLeast' ? null : regla === 'auto' ? redondo(s.linea / 240 * 100) + '%' : pt(s.linea / 20)]);
        sal.push(['line-height-at-least', s.linea !== null && regla === 'atLeast' ? pt(s.linea / 20) : null]);
      }
    }
    if ('saltoAntes' in c) sal.push(['break-before', c.saltoAntes ? 'page' : null]);
    if ('mantenerConSiguiente' in c) sal.push(['keep-with-next', c.mantenerConSiguiente ? 'always' : null]);
    return sal;
  }

  // ───────── Estilos con nombre que se crean al usarlos ─────────
  const TEXTO = (extra) => ({ 'font-weight': 'bold', 'font-weight-asian': 'bold', 'font-weight-complex': 'bold', ...extra });
  const PREDEFINIDOS = {
    Title: { vis: 'Title', es: 'Título', p: { 'text-align': 'center', 'margin-top': '0pt', 'margin-bottom': '6pt' }, t: TEXTO({ 'font-size': '28pt', 'font-size-asian': '28pt', 'font-size-complex': '28pt' }) },
    Subtitle: { vis: 'Subtitle', es: 'Subtítulo', p: { 'text-align': 'center', 'margin-top': '0pt', 'margin-bottom': '12pt' }, t: { 'font-size': '14pt', 'font-size-asian': '14pt', 'font-size-complex': '14pt' } },
    Heading_20_1: { vis: 'Heading 1', es: 'Título 1', nivel: 1, p: { 'margin-top': '12pt', 'margin-bottom': '6pt', 'keep-with-next': 'always' }, t: TEXTO({ 'font-size': '18pt', 'font-size-asian': '18pt', 'font-size-complex': '18pt' }) },
    Heading_20_2: { vis: 'Heading 2', es: 'Título 2', nivel: 2, p: { 'margin-top': '10pt', 'margin-bottom': '4pt', 'keep-with-next': 'always' }, t: TEXTO({ 'font-size': '14pt', 'font-size-asian': '14pt', 'font-size-complex': '14pt' }) },
    Heading_20_3: { vis: 'Heading 3', es: 'Título 3', nivel: 3, p: { 'margin-top': '8pt', 'margin-bottom': '4pt', 'keep-with-next': 'always' }, t: TEXTO({ 'font-size': '12pt', 'font-size-asian': '12pt', 'font-size-complex': '12pt' }) },
  };
  const ES = { Standard: 'Normal', Text_20_body: 'Cuerpo de texto', Quotations: 'Cita', Preformatted_20_Text: 'Texto preformateado', Heading_20_4: 'Título 4', Heading_20_5: 'Título 5', Heading_20_6: 'Título 6', List_20_Paragraph: 'Párrafo de lista' };
  const ORDEN = ['Standard', 'Title', 'Subtitle', 'Heading_20_1', 'Heading_20_2', 'Heading_20_3', 'Text_20_body', 'Quotations', 'Preformatted_20_Text'];

  function crear(modelo, deps) {
    const { codificar, Rechazo, prefijoDe } = deps;
    const est = modelo.est;
    const fuentes = { contenido: modelo.xml, estilos: modelo.xmlEstilos || '' };
    const creados = new Map();        // nombre → { nombre, spec, xml, nombrado }
    const porClave = new Map();       // clave canónica → nombre
    const cuenta = { MCT: 0, MCP: 0, MCL: 0 };
    const usados = new Set();
    // (con cualquier prefijo: un archivo puede llamar `s:` al espacio de nombres de los estilos)
    for (const src of [fuentes.contenido, fuentes.estilos]) for (const m of src.matchAll(/[\w.-]+:name="([^"]*)"/g)) usados.add(m[1]);

    // ── prefijos de espacios de nombres de una parte ──
    function prefijosDe(raiz) {
      const mapa = {};
      for (const k of Object.keys(raiz.attrs)) if (k.startsWith('xmlns:')) for (const n of Object.keys(URI)) if (raiz.attrs[k] === URI[n]) mapa[n] = k.slice(6);
      return mapa;
    }
    const pref = { contenido: prefijosDe(modelo.raizContenido), estilos: {} };
    if (modelo.xmlEstilos) pref.estilos = prefijosDe(X().analizar(modelo.xmlEstilos));

    // ── especificación de un estilo: { familia, padre, extra, props, resto } ──
    const vacioSpec = (familia, padre) => ({ familia, padre: padre || null, extra: new Map(), props: new Map(), resto: '' });
    const copiar = (s) => ({ familia: s.familia, padre: s.padre, extra: new Map(s.extra), props: new Map([...s.props].map(([k, e]) => [k, { attrs: new Map(e.attrs), interior: e.interior }])), resto: s.resto });
    const esVacio = (s) => !s.props.size && !s.extra.size && !s.resto;
    function specDeNodo(nodo, parte) {
      const src = fuentes[parte] || '', a = nodo.attrs, spec = vacioSpec(A(nodo).family, A(nodo)['parent-style-name']);
      for (const k of Object.keys(a)) { const l = local(k); if (k.startsWith('xmlns') || ['name', 'family', 'parent-style-name', 'auto-update', 'display-name'].includes(l)) continue; spec.extra.set(l, { k, v: a[k] }); }
      for (const h of nodo.hijos) {
        if (h.local === 'text-properties' || h.local === 'paragraph-properties') {
          const m = new Map();
          for (const k of Object.keys(h.attrs)) if (!k.startsWith('xmlns')) m.set(local(k), { k, v: h.attrs[k] });
          spec.props.set(h.local, { attrs: m, interior: h.iniCierre >= 0 ? src.slice(h.finApertura, h.iniCierre) : '' });
        } else spec.resto += src.slice(h.ini, h.fin);
      }
      return spec;
    }
    function especificacion(familia, nombre) {
      if (!nombre) return vacioSpec(familia, null);
      const c = creados.get(nombre);
      if (c) return copiar(c.spec);
      const e = est.estiloAuto(familia, nombre);
      if (e && e.nodo && fuentes[e.parte] !== undefined) return specDeNodo(e.nodo, e.parte);
      return vacioSpec(familia, nombre);
    }
    function aplicarPares(spec, elemento, pares) {
      let e = spec.props.get(elemento);
      for (const [l, v] of pares) {
        if (v === null) { if (e) e.attrs.delete(l); continue; }
        if (!e) { e = { attrs: new Map(), interior: '' }; spec.props.set(elemento, e); }
        const previo = e.attrs.get(l);
        e.attrs.set(l, { k: previo ? previo.k : null, v });
      }
      if (e && !e.attrs.size && !e.interior) spec.props.delete(elemento);
    }
    function claveDe(s) {
      return JSON.stringify([s.familia, s.padre, ordenado(s.extra).map(([l, e]) => [l, e.v]),
        ['paragraph-properties', 'text-properties'].map((n) => { const e = s.props.get(n); return e ? [ordenado(e.attrs).map(([l, x]) => [l, x.v]), e.interior] : null; }), s.resto]);
    }
    // el XML de un <style:style> a partir de su especificación, con los prefijos de la parte donde irá
    function escribir(nombre, spec, p) {
      const decl = new Set();
      const P = (ns) => { if (p[ns]) return p[ns]; decl.add(ns); return ns; };
      const att = (e, l) => (e.k ? e.k : P(nsDe(l)) + ':' + l);
      const S = P('style');
      let cuerpo = '';
      for (const n of ['paragraph-properties', 'text-properties']) {
        const e = spec.props.get(n);
        if (!e) continue;
        const attrs = ordenado(e.attrs).map(([l, x]) => ` ${att(x, l)}="${X().escAttr(x.v)}"`).join('');
        cuerpo += e.interior ? `<${S}:${n}${attrs}>${e.interior}</${S}:${n}>` : `<${S}:${n}${attrs}/>`;
      }
      cuerpo += spec.resto;
      let a = ` ${S}:name="${X().escAttr(nombre)}" ${S}:family="${spec.familia}"`;
      if (spec.padre) a += ` ${S}:parent-style-name="${X().escAttr(spec.padre)}"`;
      for (const [l, e] of ordenado(spec.extra)) a += ` ${att(e, l)}="${X().escAttr(e.v)}"`;
      const xmlns = [...decl].map((n) => ` xmlns:${n}="${URI[n]}"`).join('');
      return `<${S}:style${xmlns}${a}>${cuerpo}</${S}:style>`;
    }

    // ── registro ──
    // Los estilos automáticos de content.xml no existen para styles.xml: mientras se edita un encabezado o un pie (`enEstilos`) solo se reutilizan los creados,
    // que `finalizar` escribe en la parte que los use; un T1 del cuerpo puesto en un encabezado quedaría sin definir y el formato se perdería. En el cuerpo
    // el del archivo va primero (aunque un encabezado ya haya creado uno igual): así quitar y volver a poner devuelve el archivo de antes.
    let existentes = null, enEncabezado = false;
    function enEstilos(fn) { const antes = enEncabezado; enEncabezado = true; try { return fn(); } finally { enEncabezado = antes; } }
    function reutilizable(clave) {
      if (!enEncabezado) {
        if (!existentes) {
          existentes = new Map();
          for (const familia of ['text', 'paragraph']) for (const e of est.automaticosDe(familia)) { const k = claveDe(specDeNodo(e.nodo, e.parte)); if (!existentes.has(k)) existentes.set(k, e.nombre); }
        }
        if (existentes.has(clave)) return existentes.get(clave);
      }
      return porClave.get(clave) || null;
    }
    // El nombre de un estilo automático con esa especificación: el que ya existe, o uno nuevo
    function registrar(spec, prefijo) {
      const clave = claveDe(spec), ya = reutilizable(clave);
      if (ya) return ya;
      let n;
      do n = prefijo + (++cuenta[prefijo]); while (usados.has(n) || est.buscar(spec.familia, n));
      usados.add(n);
      const xml = escribir(n, spec, pref.contenido);
      creados.set(n, { nombre: n, spec: copiar(spec), xml, nombrado: false });
      porClave.set(clave, n);
      est.registrar(xml, false);
      return n;
    }

    // Un estilo de otra clase (de lista) que se escribe tal cual lo construye `construir(nombre)`; se reutiliza si la clave ya está
    function registrarXml(prefijo, clave, construir) {
      const k = 'x|' + clave;
      if (porClave.has(k)) return porClave.get(k);
      let n;
      do { cuenta[prefijo] = (cuenta[prefijo] || 0) + 1; n = prefijo + cuenta[prefijo]; } while (usados.has(n));
      usados.add(n);
      const xml = construir(n);
      creados.set(n, { nombre: n, spec: null, xml, nombrado: false });
      porClave.set(k, n);
      est.registrar(xml, false);
      return n;
    }

    // ───────── Formato de caracteres ─────────
    const chain = (raiz, ini, fin) => {
      const sal = [];
      let n = raiz;
      for (;;) {
        const h = n.hijos.find((c) => c.iniCierre >= 0 && c.finApertura <= ini && fin <= c.iniCierre);
        if (!h) break;
        sal.push(h); n = h;
      }
      return sal;
    };
    const conEstilo = (n) => n.local === 'span' || n.local === 'a';
    const alto = /[\uD800-\uDBFF]/, bajo = /[\uDC00-\uDFFF]/;

    // Las ediciones que ponen `cambios` a los caracteres [desde, hasta) de un párrafo; F: el párrafo leído; rBase: su formato base (el de su estilo)
    function formatoCaracter(F, desde, hasta, cambios, rBase) {
      const { xml, raiz, piezas, texto } = F, pf = prefijoDe(raiz);
      const cub = piezas.filter((z) => z.hasta > z.desde && z.hasta > desde && z.desde < hasta);
      if (!cub.length) return xml;
      // las unidades: un trozo de texto, una pieza suelta, o el campo entero cuando el texto es el resultado de un campo
      const unidades = [];
      for (const z of cub) {
        const cad = chain(raiz, z.ini, z.fin);
        const campo = z.opaco ? cad.findIndex((n) => n.local === z.opaco) : -1;
        const padre = campo >= 0 ? cad.slice(0, campo) : cad;
        const previa = unidades[unidades.length - 1];
        if (campo >= 0) {
          const nodo = cad[campo];
          if (previa && previa.nodo === nodo) { previa.piezas.push(z); continue; }
          unidades.push({ ini: nodo.ini, fin: nodo.fin, nodo, padre, piezas: [z], completa: true });
          continue;
        }
        const editable = z.t === 'texto' && (z.hueco || z.espacio) && !z.rara && !z.nota && !z.opaco;
        const k1 = Math.max(desde, z.desde) - z.desde, k2 = Math.min(hasta, z.hasta) - z.desde;
        const parcial = editable && (k1 > 0 || k2 < z.hasta - z.desde);
        unidades.push({ ini: z.ini, fin: z.fin, z, padre, piezas: [z], completa: !parcial, k1, k2 });
      }
      const plena = new Set();
      for (const u of unidades) if (u.completa) for (const z of u.piezas) plena.add(z);
      const rDe = (cad) => cad.reduce((r, n) => (conEstilo(n) ? est.sumarR(r, est.resolverCaracter(A(n)['style-name'])) : r), rBase);
      // «Borrar formato»: queda lo del estilo del párrafo y los estilos de carácter con nombre (no el formato directo, que es lo que escriben los spans automáticos)
      const conservado = (cad) => cad.reduce((r, n) => {
        if (!conEstilo(n)) return r;
        const nombre = A(n)['style-name'], auto = nombre && est.estiloAuto('text', nombre), fijo = auto ? auto.padre : nombre;
        return fijo ? est.sumarR(r, est.resolverCaracter(fijo)) : r;
      }, rBase);
      const cambiosDe = (cad) => (cambios.limpiar ? limpio(conservado(cad)) : cambios);
      // runs: unidades seguidas con el mismo padre inmediato
      const runs = [];
      for (const u of unidades) {
        const ultimo = runs[runs.length - 1], P = u.padre[u.padre.length - 1] || raiz;
        if (ultimo && ultimo.P === P) { ultimo.unidades.push(u); continue; }
        runs.push({ P, cad: u.padre, unidades: [u] });
      }
      // un span cuyo contenido queda cubierto por entero cambia de estilo en vez de envolverse
      const ed = [], enSitio = new Map();
      for (const run of runs) {
        const P = run.P;
        if (P === raiz || P.local !== 'span' || enSitio.has(P)) continue;
        const dentro = piezas.filter((z) => z.hasta > z.desde && z.ini >= P.finApertura && z.fin <= P.iniCierre);
        enSitio.set(P, dentro.every((z) => plena.has(z)) ? run.cad : null);
      }
      for (const [P, cad] of enSitio) {
        if (!cad) continue;
        const nombreP = A(P)['style-name'] || null, spec = especificacion('text', nombreP);
        const fuera = rDe(cad.slice(0, -1)), base = spec.padre ? est.sumarR(fuera, est.resolverCaracter(spec.padre)) : fuera;
        aplicarPares(spec, 'text-properties', paresTexto(cambiosDe(cad), efectivo(base)));
        const nuevo = esVacio(spec) ? spec.padre : registrar(spec, 'MCT');
        if (nuevo === nombreP) continue;
        if (nuevo === null) { ed.push({ ini: P.ini, fin: P.finApertura, texto: '' }, { ini: P.iniCierre, fin: P.fin, texto: '' }); continue; }
        ed.push({ ini: P.ini, fin: P.finApertura, texto: X().ponerAtributo(xml.slice(P.ini, P.finApertura), P.attrs && Object.keys(P.attrs).find((k) => local(k) === 'style-name') || pf + 'style-name', nuevo) });
      }
      for (const run of runs) {
        if (enSitio.get(run.P)) continue;
        const r = rDe(run.cad), pares = paresTexto(cambiosDe(run.cad), efectivo(r));
        if (!pares.some(([, v]) => v !== null)) continue;
        const spec = vacioSpec('text', null);
        aplicarPares(spec, 'text-properties', pares);
        if (esVacio(spec)) continue;
        const nombre = registrar(spec, 'MCT');
        const abrir = `<${pf}span ${pf}style-name="${X().escAttr(nombre)}">`, cerrar = `</${pf}span>`;
        const u0 = run.unidades[0], uN = run.unidades[run.unidades.length - 1];
        const trozos = (u) => {
          const z = u.z;
          let k1 = u.completa ? 0 : u.k1, k2 = u.completa ? z.s.length : u.k2;
          if (!u.completa) { if (k1 > 0 && k1 < z.s.length && alto.test(z.s[k1 - 1]) && bajo.test(z.s[k1])) k1--; if (k2 > 0 && k2 < z.s.length && alto.test(z.s[k2 - 1]) && bajo.test(z.s[k2])) k2++; }
          const antes = texto[z.desde - 1] || '';
          const izq = z.s.slice(0, k1), med = z.s.slice(k1, k2), der = z.s.slice(k2);
          // lo que sigue al trozo: si la pieza de al lado empieza con un espacio suelto, un espacio al final de este trozo tendría que ir como <text:s/>
          const sig = piezas.find((q) => q.desde >= z.hasta && q.hasta > q.desde && q !== z);
          const finSeguro = !(sig && sig.t === 'texto' && sig.hueco && sig.s[0] === ' ');
          return {
            izq: izq ? codificar(izq, antes, pf, !!z.hueco) : '', med: codificar(med, izq ? izq[izq.length - 1] : antes, pf, der ? !!z.hueco : finSeguro),
            der: der ? codificar(der, med[med.length - 1], pf, finSeguro) : '',
          };
        };
        let cuerpo, izq = '', der = '', ini = u0.ini, fin = uN.fin;
        if (u0 === uN) {
          if (u0.completa) cuerpo = xml.slice(u0.ini, u0.fin);
          else { const t = trozos(u0); izq = t.izq; cuerpo = t.med; der = t.der; }
        } else {
          let cab = xml.slice(u0.ini, u0.fin), col = xml.slice(uN.ini, uN.fin);
          if (!u0.completa) { const t = trozos(u0); izq = t.izq; cab = t.med; }
          if (!uN.completa) { const t = trozos(uN); der = t.der; col = t.med; }
          cuerpo = cab + xml.slice(u0.fin, uN.ini) + col;
        }
        ed.push({ ini, fin, texto: izq + abrir + cuerpo + cerrar + der });
      }
      return ed.length ? X().aplicar(xml, ed) : xml;
    }

    // ───────── Formato de párrafo ─────────
    function nombreEstiloDe(raiz) { return A(raiz)['style-name'] || null; }
    function ponerNombreEstilo(xml, raiz, nombre) {
      const clave = Object.keys(raiz.attrs).find((k) => local(k) === 'style-name');
      const tag = X().ponerAtributo(xml.slice(raiz.ini, raiz.finApertura), clave || prefijoDe(raiz) + 'style-name', nombre || null);
      return tag + xml.slice(raiz.finApertura);
    }
    function formatoParrafo(xml, cambios) {
      const raiz = X().analizar(xml), actual = nombreEstiloDe(raiz);
      const spec = especificacion('paragraph', actual);
      const ref = spec.padre && est.resolverParrafo(spec.padre) ? est.resolverParrafo(spec.padre).p : null;
      aplicarPares(spec, 'paragraph-properties', paresParrafo(cambios, ref));
      const nuevo = esVacio(spec) ? spec.padre : registrar(spec, 'MCP');
      return nuevo === actual ? xml : ponerNombreEstilo(xml, raiz, nuevo);
    }

    // ───────── Estilos con nombre ─────────
    const hayStyles = !!modelo.xmlEstilos;
    const existe = (id) => !!est.nombrado('paragraph', id);
    function predeterminado() { return existe('Standard') ? 'Standard' : null; }
    function estilosDisponibles() {
      const sal = [], visto = new Set();
      const poner = (id, nuevo) => {
        if (visto.has(id)) return;
        visto.add(id);
        const e = est.nombrado('paragraph', id), def = PREDEFINIDOS[id];
        sal.push({ id, nombre: ES[id] || (def && def.es) || (e && e.nombreVisible) || id, predeterminado: id === 'Standard', ...(nuevo ? { nuevo: true } : {}) });
      };
      if (!existe('Standard')) poner('Standard');
      for (const id of ORDEN) { if (existe(id)) poner(id); else if (PREDEFINIDOS[id] && hayStyles) poner(id, true); }
      const enUso = new Set();
      for (const b of modelo.parrafos()) { let e = b.estilo && est.buscar('paragraph', b.estilo); while (e && e.auto) e = e.padre && est.nombrado('paragraph', e.padre); if (e && e.nombre) enUso.add(e.nombre); }
      for (const id of enUso) poner(id);
      return sal;
    }
    // Crea (en memoria; se escribe en styles.xml al guardar) un estilo predefinido que el archivo no tiene; devuelve su nombre o null
    // (la cinta de la aplicación usa los nombres de Word: Heading1…; aquí son los de LibreOffice)
    const ALIAS = { Heading1: 'Heading_20_1', Heading2: 'Heading_20_2', Heading3: 'Heading_20_3', Normal: 'Standard' };
    const INVERSO = { Heading_20_1: 'Heading1', Heading_20_2: 'Heading2', Heading_20_3: 'Heading3' };
    function estiloComun(nombre) {
      let e = nombre && est.buscar('paragraph', nombre);
      while (e && e.auto) e = e.padre && est.nombrado('paragraph', e.padre);
      const n = e ? e.nombre : null;
      return !n || n === 'Standard' ? null : INVERSO[n] || n;
    }
    function asegurarEstilo(id) {
      id = ALIAS[id] || id;
      if (!id || id === 'Standard') return predeterminado();
      if (existe(id)) return id;
      const d = PREDEFINIDOS[id];
      if (!d || !hayStyles) return null;
      const spec = vacioSpec('paragraph', null);
      spec.extra.set('display-name', { k: null, v: d.vis });
      if (predeterminado()) spec.extra.set('next-style-name', { k: null, v: 'Standard' });
      if (d.nivel) spec.extra.set('default-outline-level', { k: null, v: String(d.nivel) });
      spec.extra.set('class', { k: null, v: 'text' });
      aplicarPares(spec, 'paragraph-properties', Object.entries(d.p));
      aplicarPares(spec, 'text-properties', Object.entries(d.t));
      const xml = escribir(id, spec, pref.estilos);
      creados.set(id, { nombre: id, spec, xml, nombrado: true });
      usados.add(id);
      est.registrar(xml, true);
      return id;
    }
    // Pone un estilo con nombre a un párrafo conservando su formato directo (como Word y LibreOffice): el estilo automático pasa a colgar del nuevo
    function ponerEstilo(xml, id, dialecto) {
      const raiz = X().analizar(xml), actual = nombreEstiloDe(raiz);
      let nombre = id;
      if (actual && est.estiloAuto('paragraph', actual) && id !== actual) {
        const spec = especificacion('paragraph', actual);
        spec.padre = id;
        nombre = esVacio(spec) ? id : registrar(spec, 'MCP');
      }
      return dialecto.ponerEstilo(xml, nombre);
    }

    // ───────── Al guardar ─────────
    const referenciados = (xml) => new Set([...xml.matchAll(/style-name="([^"]+)"/g)].map((m) => m[1]));
    function cierre(xmlFinal) {
      const usadosAhora = referenciados(xmlFinal), automaticos = [], nombrados = new Set();
      for (const c of creados.values()) if (!c.nombrado && usadosAhora.has(c.nombre)) { automaticos.push(c); if (c.spec && c.spec.padre && creados.has(c.spec.padre)) nombrados.add(c.spec.padre); }
      for (const c of creados.values()) if (c.nombrado && usadosAhora.has(c.nombre)) nombrados.add(c.nombre);
      return { automaticos, nombrados: [...nombrados].map((n) => creados.get(n)) };
    }
    // content.xml con los estilos automáticos creados que usa
    // (los que ya están definidos no se repiten: un encabezado que se edita otra vez parte del styles.xml que ya los trae; y van con los prefijos de ESTA
    // parte, que en styles.xml pueden no ser los de content.xml)
    function finalizar(xml) {
      const raiz = X().analizar(xml), as = raiz.hijos.find((h) => h.local === 'automatic-styles');
      const ya = new Set(as ? as.hijos.map((h) => A(h).name) : []);
      const automaticos = cierre(xml).automaticos.filter((c) => !ya.has(c.nombre));
      if (!automaticos.length) return xml;
      const p = prefijosDe(raiz);
      const bloque = automaticos.map((c) => (c.spec ? escribir(c.nombre, c.spec, p) : c.xml)).join('');
      if (as && as.iniCierre >= 0) return X().aplicar(xml, [{ ini: as.iniCierre, fin: as.iniCierre, texto: bloque }]);
      if (as) return X().aplicar(xml, [{ ini: as.ini, fin: as.fin, texto: `<${as.nombre}>${bloque}</${as.nombre}>` }]);
      const cuerpo = raiz.hijos.find((h) => h.local === 'body'), o = raiz.nombre.slice(0, raiz.nombre.indexOf(':') + 1);
      const donde = cuerpo ? cuerpo.ini : raiz.finApertura;
      return X().aplicar(xml, [{ ini: donde, fin: donde, texto: `<${o}automatic-styles>${bloque}</${o}automatic-styles>` }]);
    }
    // styles.xml con los estilos con nombre creados que usa el documento final
    // styles.xml: los estilos con nombre que el documento usa y los estilos automáticos que usan sus encabezados y pies. Parte del styles.xml original mientras
    // nadie más lo haya tocado (así un estilo creado que ya no se usa deja de escribirse y deshacer todo devuelve el archivo de antes) y del de ahora si los
    // encabezados escribieron en él.
    let ultimoEscrito = null, baseUsada = null;
    async function alGuardar(paquete, xmlFinal) {
      if (!hayStyles) return;
      const actual = await paquete.leer('styles.xml');
      if (actual === null) return;
      let xml = actual === ultimoEscrito ? baseUsada : actual;
      baseUsada = xml;
      const { nombrados } = cierre(xmlFinal + '\n' + xml);
      const enEstilos = referenciados(xml);
      const tiene = (xml0, nombre) => new RegExp('style:name="' + nombre + '"').test(xml0);
      const faltanN = nombrados.filter((c) => !tiene(xml, c.nombre));
      const faltanA = [...creados.values()].filter((c) => !c.nombrado && enEstilos.has(c.nombre) && !tiene(xml, c.nombre));
      const poner = (nombreEl, bloque, crearAntesDe) => {
        if (!bloque) return;
        const raiz = X().analizar(xml), el = raiz.hijos.find((h) => h.local === nombreEl);
        const o = raiz.nombre.slice(0, raiz.nombre.indexOf(':') + 1);
        if (el && el.iniCierre >= 0) xml = X().aplicar(xml, [{ ini: el.iniCierre, fin: el.iniCierre, texto: bloque }]);
        else if (el) xml = X().aplicar(xml, [{ ini: el.ini, fin: el.fin, texto: `<${el.nombre}>${bloque}</${el.nombre}>` }]);
        else {
          const sig = raiz.hijos.find((h) => crearAntesDe.includes(h.local)), donde = sig ? sig.ini : raiz.iniCierre;
          xml = X().aplicar(xml, [{ ini: donde, fin: donde, texto: `<${o}${nombreEl}>${bloque}</${o}${nombreEl}>` }]);
        }
      };
      poner('styles', faltanN.map((c) => c.xml).join(''), ['automatic-styles', 'master-styles']);
      poner('automatic-styles', faltanA.map((c) => (c.spec ? escribir(c.nombre, c.spec, pref.estilos) : c.xml)).join(''), ['master-styles']);
      if (xml !== actual) paquete.escribir('styles.xml', xml);
      ultimoEscrito = xml;
    }

    return { enEstilos, registrarXml, estiloComun, formatoCaracter, formatoParrafo, ponerEstilo, estilosDisponibles, asegurarEstilo, predeterminado, finalizar, alGuardar, registrar, especificacion, vacioSpec, aplicarPares, esVacio, escribir, creados, pref, cuenta };
  }

  globalThis.OdtFormato = { crear };
})();

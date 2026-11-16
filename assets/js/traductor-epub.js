/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Traductor de documentos · EPUB. Edita el libro ORIGINAL: cambia el texto de cada capítulo (XHTML) y del índice (nav y toc.ncx), y el idioma
// del libro (dc:language y lang de cada capítulo). Todo lo demás (portada, imágenes, hojas de estilo, fuentes, orden de lectura) queda igual.
//
//   const { bytes, estadistica } = await TraductorEpub.traducir(datos, { traducir, destino, progreso })
//
// Cada BLOQUE (p, h1…h6, li, td, th, dt, dd, figcaption, caption, blockquote, div…) es una unidad. Sus elementos EN LÍNEA (em, strong,
// a, span, sup…) dan el formato: la cadena de etiquetas que envuelve cada texto. El contenido del bloque se rearma abriendo y cerrando las
// etiquetas cuando cambia el formato (si la traducción cambia el orden, un enlace puede quedar partido en dos con el mismo destino).
// Imágenes, saltos (br), fórmulas (math), svg, ruby y código quedan tal cual y parten el bloque. Un bloque dentro de otro (un p dentro
// de un li) es su propia unidad. No se traduce: pre, code, script, style.
// Los espacios y saltos de línea dentro de un bloque cuentan como un espacio (como en HTML).
//
// Requiere: zip-propio.js, xml-editar.js y traductor-documentos.js.
(function (raiz) {
  'use strict';

  const X = () => raiz.XmlEditar;
  const BLOQUES = new Set(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'dt', 'dd', 'td', 'th', 'caption', 'figcaption', 'blockquote', 'div',
    'section', 'article', 'aside', 'header', 'footer', 'nav', 'address', 'summary', 'legend', 'label', 'body', 'figure', 'main', 'hgroup']);
  const EN_LINEA = new Set(['a', 'abbr', 'b', 'bdi', 'bdo', 'cite', 'data', 'dfn', 'em', 'i', 'kbd', 'mark', 'q', 's', 'samp', 'small', 'span',
    'strong', 'sub', 'sup', 'time', 'u', 'var', 'ins', 'font', 'big', 'tt', 'strike']);
  const NO_TRADUCIR = new Set(['pre', 'code', 'script', 'style', 'svg', 'math', 'ruby', 'head', 'del']);
  // Entidades de HTML que un EPUB a veces trae sin declararlas (XML solo conoce las cinco básicas).
  const ENTIDADES = { nbsp: ' ', mdash: '—', ndash: '–', hellip: '…', laquo: '«', raquo: '»', ldquo: '“', rdquo: '”', lsquo: '‘',
    rsquo: '’', copy: '©', reg: '®', trade: '™', shy: '­', iexcl: '¡', iquest: '¿', middot: '·', deg: '°', euro: '€' };
  const decodificar = (s) => X().decodificar(s.replace(/&([a-z]+);/g, (t, e) => ENTIDADES[e] || t));

  // ¿Tiene texto adentro (en sus hijos en línea)? Un elemento en línea vacío (<a id="p12"></a>, un salto de página) queda tal cual.
  const conTexto = (xml, n) => !n.vacio && /\S/.test(xml.slice(n.finApertura, n.iniCierre).replace(/<[^>]*>/g, ''));
  // Contenedores (ul, ol, table, tr, dl…): ni bloque ni en línea; se rearman por dentro (sus hijos sí pueden ser bloques).
  const esContenedor = (n) => !n.vacio && !BLOQUES.has(n.local) && !EN_LINEA.has(n.local) && !NO_TRADUCIR.has(n.local) && n.hijos.length > 0;

  // Contenido de un bloque, aplanado: [{ texto, cadena } | { barrera: xml, cadena } | { bloque: nodo } | { contenedor: nodo }].
  // cadena = [{ apertura, nombre }], las etiquetas en línea que envuelven a cada cosa.
  function aplanar(xml, nodo, cadena, items) {
    let pos = nodo.finApertura;
    // Los espacios y saltos del HTML cuentan como uno; el espacio duro (&nbsp;) se conserva.
    const texto = (fin) => { if (fin > pos) items.push({ texto: decodificar(xml.slice(pos, fin)).replace(/[ \t\r\n]+/g, ' '), cadena }); };
    for (const h of nodo.hijos) {
      texto(h.ini);
      if (BLOQUES.has(h.local)) items.push({ bloque: h });
      else if (EN_LINEA.has(h.local) && conTexto(xml, h)) aplanar(xml, h, [...cadena, { apertura: xml.slice(h.ini, h.finApertura), nombre: h.nombre }], items);
      else if (esContenedor(h)) items.push({ contenedor: h, cadena });
      else items.push({ barrera: xml.slice(h.ini, h.fin), cadena });
      pos = h.fin;
    }
    if (!nodo.vacio) texto(nodo.iniCierre);
    return items;
  }

  const clave = (cadena) => cadena.map((c) => c.apertura).join('');

  // Emite tramos con su cadena de etiquetas: cierra lo que sobra y abre lo que falta respecto del anterior. Una etiqueta con id que
  // se abre por segunda vez (un enlace partido en dos) va sin el id: no puede haber dos iguales.
  function emitir(elementos) {
    let pila = [];
    let r = '';
    const abiertas = new Set();
    const abrir = (c) => {
      if (abiertas.has(c.apertura) && /\sid\s*=/.test(c.apertura)) return X().ponerAtributo(c.apertura, 'id', null);
      abiertas.add(c.apertura);
      return c.apertura;
    };
    const ir = (cadena) => {
      let comun = 0;
      while (comun < pila.length && comun < cadena.length && pila[comun].apertura === cadena[comun].apertura) comun++;
      for (let i = pila.length - 1; i >= comun; i--) r += `</${pila[i].nombre}>`;
      for (let i = comun; i < cadena.length; i++) r += abrir(cadena[i]);
      pila = cadena.slice();
    };
    for (const e of elementos) {
      ir(e.cadena || []);
      r += e.xml !== undefined ? e.xml : X().esc(e.texto);
    }
    ir([]);
    return r;
  }

  // Analiza un XHTML: los bloques de más afuera (los de adentro se rearman dentro de ellos) y las unidades de todos.
  function analizarXhtml(xml) {
    const raizXml = X().analizar(xml);
    const externos = [];
    const unidades = [];      // { nodo, trozos: [[índices de items]] , items }
    const porNodo = new Map();
    (function ir(n, dentro) {
      for (const h of n.hijos) {
        if (NO_TRADUCIR.has(h.local)) continue;
        if (BLOQUES.has(h.local)) {
          if (!dentro) externos.push(h);
          const items = aplanar(xml, h, [], []);
          // Trozos: textos entre barreras o bloques internos.
          const trozos = [];
          let actual = [];
          items.forEach((it, i) => {
            if (it.texto !== undefined) actual.push(i);
            else { if (actual.length) trozos.push(actual); actual = []; }
          });
          if (actual.length) trozos.push(actual);
          const u = { nodo: h, items, trozos: trozos.filter((t) => t.some((i) => /\p{L}/u.test(items[i].texto))) };
          porNodo.set(h, u);
          unidades.push(u);
          ir(h, true);
        } else ir(h, dentro);
      }
    })(raizXml.local === 'html' ? raizXml : raizXml, false);
    // <title> está en <head>, que se saltea arriba: se agrega aparte.
    const head = X().hijo(raizXml, 'head');
    const title = head && X().hijo(head, 'title');
    if (title && !title.vacio) {
      const items = aplanar(xml, title, [], []);
      const u = { nodo: title, items, trozos: items.some((it) => it.texto && /\p{L}/u.test(it.texto)) ? [items.map((_, i) => i)] : [] };
      porNodo.set(title, u); unidades.push(u); externos.push(title);
    }
    return { raizXml, externos, unidades, porNodo };
  }

  function rearmar(xml, nodo, porNodo, traducidos) {
    const u = porNodo.get(nodo);
    const elementos = [];
    const trozoDe = new Map();
    u.trozos.forEach((t, k) => t.forEach((i, j) => trozoDe.set(i, [k, j])));
    u.items.forEach((it, i) => {
      if (it.bloque) { elementos.push({ xml: rearmar(xml, it.bloque, porNodo, traducidos), cadena: [] }); return; }
      if (it.contenedor) { elementos.push({ xml: rearmarContenedor(xml, it.contenedor, porNodo, traducidos), cadena: it.cadena }); return; }
      if (it.barrera !== undefined) { elementos.push({ xml: it.barrera, cadena: it.cadena }); return; }
      const pos = trozoDe.get(i);
      if (!pos) { elementos.push({ texto: it.texto, cadena: it.cadena }); return; }
      if (pos[1] !== 0) return;              // el trozo traducido se emite entero en su primer texto
      const tr = traducidos.get(u)[pos[0]];
      const cadenaDe = new Map(u.trozos[pos[0]].map((ii) => [clave(u.items[ii].cadena), u.items[ii].cadena]));
      for (const t of tr) elementos.push({ texto: t.texto, cadena: cadenaDe.get(t.formato) || [] });
    });
    return xml.slice(nodo.ini, nodo.finApertura) + emitir(elementos) + `</${nodo.nombre}>`;
  }

  // Un contenedor (ul, table…): su apertura, sus hijos rearmados (los bloques traducidos; lo demás y los espacios entre ellos, tal cual).
  function rearmarContenedor(xml, nodo, porNodo, traducidos) {
    let r = xml.slice(nodo.ini, nodo.finApertura);
    let pos = nodo.finApertura;
    for (const h of nodo.hijos) {
      r += xml.slice(pos, h.ini);
      if (porNodo.has(h)) r += rearmar(xml, h, porNodo, traducidos);
      else if (esContenedor(h)) r += rearmarContenedor(xml, h, porNodo, traducidos);
      else r += xml.slice(h.ini, h.fin);
      pos = h.fin;
    }
    return r + xml.slice(pos, nodo.iniCierre) + xml.slice(nodo.iniCierre, nodo.fin);
  }

  const ponerLang = (apertura, cod) => {
    let e = apertura;
    if (/\slang\s*=/.test(e) || !/\sxml:lang\s*=/.test(e)) e = X().ponerAtributo(e, 'lang', cod);
    if (/\sxml:lang\s*=/.test(e)) e = X().ponerAtributo(e, 'xml:lang', cod);
    return e;
  };

  async function traducir(datos, { traducir: traductor, destino, progreso }) {
    const zip = await raiz.ZipPropio.loadAsync(datos);
    const contenedor = await zip.file('META-INF/container.xml')?.async('string');
    if (!contenedor) throw new Error('El EPUB no tiene META-INF/container.xml');
    const rutaOpf = (contenedor.match(/full-path\s*=\s*"([^"]+)"/) || [])[1];
    if (!rutaOpf || !zip.file(rutaOpf)) throw new Error('El EPUB no indica dónde está su paquete (OPF)');
    const base = rutaOpf.includes('/') ? rutaOpf.slice(0, rutaOpf.lastIndexOf('/') + 1) : '';
    let opf = await zip.file(rutaOpf).async('string');
    const manifiesto = [...opf.matchAll(/<(?:\w+:)?item\b[^>]*>/g)].map((m) => ({
      href: (m[0].match(/\shref\s*=\s*"([^"]+)"/) || [])[1], tipo: (m[0].match(/\smedia-type\s*=\s*"([^"]+)"/) || [])[1],
    })).filter((i) => i.href);
    const ruta = (href) => decodeURIComponent((base + href).replace(/[^/]+\/\.\.\//g, ''));
    const xhtmls = manifiesto.filter((i) => i.tipo === 'application/xhtml+xml').map((i) => ruta(i.href)).filter((r) => zip.file(r));
    const ncx = manifiesto.filter((i) => i.tipo === 'application/x-dtbncx+xml').map((i) => ruta(i.href)).filter((r) => zip.file(r));

    const analizados = [];
    const unidades = [];
    for (const r of xhtmls) {
      const xml = await zip.file(r).async('string');
      const a = { ruta: r, xml, ...analizarXhtml(xml) };
      for (const u of a.unidades) for (const t of u.trozos) unidades.push(t.map((i) => ({ texto: u.items[i].texto, formato: clave(u.items[i].cadena) })));
      analizados.push(a);
    }
    // toc.ncx (EPUB 2): cada <text> del índice.
    const ncxs = [];
    for (const r of ncx) {
      const xml = await zip.file(r).async('string');
      const textos = [];
      (function ir(n) { for (const h of n.hijos) { if (h.local === 'text' && !h.vacio) textos.push(h); ir(h); } })(X().analizar(xml));
      for (const t of textos) unidades.push([{ texto: decodificar(X().interior(xml, t)), formato: '' }]);
      ncxs.push({ ruta: r, xml, textos });
    }

    const r = await raiz.TraductorDocumentos.traducirUnidades(unidades, {
      traducir: (segs, d, avisar) => traductor(segs, d, { progreso: avisar }), destino, progreso,
    });
    let k = 0;
    for (const a of analizados) {
      const traducidos = new Map();
      for (const u of a.unidades) traducidos.set(u, u.trozos.map(() => r[k++]));
      const ediciones = a.externos.map((n) => ({ ini: n.ini, fin: n.fin, texto: rearmar(a.xml, n, a.porNodo, traducidos) }));
      // Idioma del capítulo.
      const html = a.raizXml;
      if (html.local === 'html') ediciones.push({ ini: html.ini, fin: html.finApertura, texto: ponerLang(a.xml.slice(html.ini, html.finApertura), destino) });
      const nuevo = X().aplicar(a.xml, ediciones);
      X().analizar(nuevo);
      zip.file(a.ruta, nuevo);
    }
    for (const n of ncxs) {
      const ediciones = n.textos.map((t) => ({ ini: t.finApertura, fin: t.iniCierre, texto: X().esc(r[k++].map((x) => x.texto).join('')) }));
      const nuevo = X().aplicar(n.xml, ediciones);
      X().analizar(nuevo);
      zip.file(n.ruta, nuevo);
    }
    // Idioma del libro.
    opf = opf.replace(/(<dc:language\b[^>]*>)[^<]*(<\/dc:language>)/, `$1${destino}$2`);
    zip.file(rutaOpf, opf);
    // El mimetype va primero y sin comprimir (lo pide el estándar): ZipPropio conserva el orden de las entradas.
    zip.file('mimetype', 'application/epub+zip', { compression: 'STORE' });
    const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/epub+zip' });
    return { bytes, estadistica: r.estadistica };
  }

  raiz.TraductorEpub = { traducir, _nucleo: { analizarXhtml, aplanar, emitir } };
})(typeof self !== 'undefined' ? self : globalThis);

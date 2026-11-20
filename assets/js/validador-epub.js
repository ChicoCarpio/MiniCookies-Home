/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Validador básico de EPUB (2026-11-19). Revisa lo que más hace que un libro no abra o se vea mal en un lector y
// devuelve hallazgos con un código y sus datos; quien lo usa pone el texto (es/en). No reemplaza a epubcheck: no
// valida contra los esquemas de EPUB ni lee el CSS por dentro.
//
//   const r = await ValidadorEpub.validar(archivoOBytes);
//   r = { version, titulo, idioma, capitulos, recursos, hallazgos: [{ nivel: 'error'|'aviso'|'info', codigo, ruta, p }],
//         conteo: { error, aviso, info } }
//
// Qué mira:
//  - Paquete: que sea un .zip; `mimetype` primero, sin comprimir, con el texto exacto y sin campo extra; nombres
//    repetidos sin contar mayúsculas o con caracteres prohibidos; cifrado (DRM) distinto de la ofuscación de fuentes.
//  - container.xml y el OPF: XML bien formado, versión, identificador único, título, idioma, fecha de modificación
//    (EPUB 3), manifest (ids repetidos, archivos que faltan, tipos que no coinciden con el contenido, archivos sin
//    listar), spine (referencias que no existen, orden de lectura vacío), portada.
//  - Índice: el nav de EPUB 3 y el NCX de EPUB 2 existen y sus enlaces llegan a un archivo del libro.
//  - Cada capítulo (XHTML): XML bien formado, espacio de nombres, ids repetidos, imágenes y hojas de estilo que faltan,
//    enlaces y anclas rotos, scripts, formularios y recursos de internet.
//  - Hojas de estilo: archivos que `url()` o `@import` piden y no están.
// Requiere: zip-propio.js. Usa detalleError (mensajes-error.js) si está cargado.

const ValidadorEpub = (function () {
  const NS_XHTML = 'http://www.w3.org/1999/xhtml';
  const NS_OPS = 'http://www.idpf.org/2007/ops';
  const TOPE_POR_CODIGO = 25; // un libro con mil imágenes sin descripción no llena el informe: se cuenta el resto
  const FUENTES_OFUSCADAS = new Set(['http://www.idpf.org/2008/embedding', 'http://ns.adobe.com/pdf/enc#RC']);
  const TIPOS_CAPITULO = new Set(['application/xhtml+xml', 'image/svg+xml']);

  // ---------- Utilidades ----------

  function leerXml(texto, tipo) {
    const doc = new DOMParser().parseFromString(texto, tipo || 'application/xml');
    const error = doc.getElementsByTagName('parsererror')[0];
    if (error) {
      // El navegador escribe «error on line 4 at column 35: Unexpected entity: nbsp» seguido de un trozo del documento.
      const bruto = (error.textContent || '').replace(/\s+/g, ' ').trim();
      const m = /error on line (\d+) at column (\d+): (.*?)(?:Below is a rendering|$)/.exec(bruto);
      return { error: m ? { linea: m[1], col: m[2], detalle: m[3].trim() } : { linea: '?', col: '?', detalle: bruto.slice(0, 160) || 'XML' } };
    }
    return { doc };
  }

  const hijos = (el, nombre) => [...el.children].filter((h) => h.localName === nombre);
  const todos = (el, nombre) => [...el.getElementsByTagNameNS('*', nombre)];
  const txt = (el) => (el && el.textContent ? el.textContent.trim() : '');
  const esRemota = (url) => /^[a-z][a-z0-9+.-]*:/i.test(url) && !/^(data|about):/i.test(url);

  // Ruta de un archivo del libro a partir de la carpeta del que lo menciona: «../img/a%20b.png#x» → «OEBPS/img/a b.png».
  function resolver(base, href) {
    const limpio = String(href).split('#')[0].split('?')[0];
    if (!limpio) return { ruta: base, fragmento: String(href).split('#')[1] || '' };
    let ruta;
    try {
      const url = new URL(limpio.split('/').map((s) => encodeURIComponent(decodeURIComponent(s))).join('/'), 'http://libro.invalido/' + base.split('/').map(encodeURIComponent).join('/'));
      ruta = decodeURIComponent(url.pathname.slice(1));
    } catch (e) { ruta = limpio; }
    return { ruta, fragmento: String(href).split('#')[1] || '' };
  }

  function tipoPorBytes(b) {
    if (b.length >= 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return 'image/png';
    if (b.length >= 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return 'image/jpeg';
    if (b.length >= 4 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x38) return 'image/gif';
    if (b.length >= 12 && b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 && b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50) return 'image/webp';
    const inicio = new TextDecoder().decode(b.subarray(0, 300)).replace(/^﻿/, '').trimStart();
    if (/^(<\?xml[^>]*>\s*)?(<!--[\s\S]*?-->\s*)?(<!DOCTYPE[^>]*>\s*)?<svg[\s>]/i.test(inicio)) return 'image/svg+xml';
    return null;
  }

  // ---------- Validación ----------

  async function validar(entrada) {
    const hallazgos = [];
    const cuenta = new Map();
    const sobran = new Map();
    const marca = (nivel, codigo, ruta, p) => {
      const n = (cuenta.get(codigo) || 0) + 1;
      cuenta.set(codigo, n);
      if (n > TOPE_POR_CODIGO) { sobran.set(codigo, (sobran.get(codigo) || 0) + 1); return; }
      hallazgos.push({ nivel, codigo, ruta: ruta || '', p: p || {} });
    };
    const resultado = { version: '', titulo: '', idioma: '', capitulos: 0, recursos: 0, hallazgos, conteo: { error: 0, aviso: 0, info: 0 } };
    const cerrar = () => {
      for (const [codigo, n] of sobran) hallazgos.push({ nivel: 'info', codigo: 'recortado', ruta: '', p: { codigo, n } });
      for (const h of hallazgos) resultado.conteo[h.nivel]++;
      const orden = { error: 0, aviso: 1, info: 2 };
      hallazgos.sort((a, b) => orden[a.nivel] - orden[b.nivel]);
      return resultado;
    };

    let zip;
    try { zip = await ZipPropio.loadAsync(entrada); } catch (e) {
      marca('error', 'zip-invalido', '', { detalle: typeof detalleError === 'function' ? detalleError(e) : String(e && e.message) });
      return cerrar();
    }
    const nombres = Object.keys(zip.files);
    const archivos = nombres.filter((n) => !zip.files[n].dir);
    const existe = (ruta) => !!zip.files[ruta] && !zip.files[ruta].dir;
    const leer = (ruta) => zip.files[ruta].async('string');

    // --- Paquete ---
    const mt = zip.files.mimetype;
    if (!mt) marca('error', 'sin-mimetype');
    else {
      if (nombres[0] !== 'mimetype') marca('error', 'mimetype-no-primero', 'mimetype', { primero: nombres[0] });
      const o = mt._origen;
      if (o && o.metodo !== 0) marca('error', 'mimetype-comprimido', 'mimetype');
      const valor = await mt.async('string');
      if (valor !== 'application/epub+zip') marca('error', 'mimetype-contenido', 'mimetype', { valor: valor.replace(/\s/g, ' ').slice(0, 40) });
      if (o && o.bytes && o.local + 30 <= o.bytes.length) {
        const v = new DataView(o.bytes.buffer, o.bytes.byteOffset, o.bytes.byteLength);
        if (v.getUint32(o.local, true) === 0x04034b50 && v.getUint16(o.local + 28, true) > 0) marca('aviso', 'mimetype-extra', 'mimetype');
      }
    }
    const minusculas = new Map();
    for (const n of archivos) {
      const k = n.toLowerCase();
      if (minusculas.has(k)) marca('error', 'nombre-repetido', n, { otro: minusculas.get(k) });
      else minusculas.set(k, n);
      if (/[\\:*?"<>|\u0000-\u001f]/.test(n) || /(^|\/)\.+$/.test(n) || /\.$/.test(n)) marca('aviso', 'nombre-prohibido', n);
    }

    let hayDrm = false;
    if (existe('META-INF/encryption.xml')) {
      const e = leerXml(await leer('META-INF/encryption.xml'));
      if (e.doc) {
        const algoritmos = todos(e.doc, 'EncryptionMethod').map((m) => m.getAttribute('Algorithm') || '');
        hayDrm = algoritmos.some((a) => !FUENTES_OFUSCADAS.has(a));
      }
      if (hayDrm) marca('aviso', 'drm', 'META-INF/encryption.xml');
    }

    // --- container.xml ---
    let rutaOpf = '';
    if (!existe('META-INF/container.xml')) marca('error', 'sin-container');
    else {
      const c = leerXml(await leer('META-INF/container.xml'));
      if (c.error) marca('error', 'xml-mal-formado', 'META-INF/container.xml', c.error);
      else {
        const rootfiles = todos(c.doc, 'rootfile');
        const raiz = rootfiles.find((r) => (r.getAttribute('media-type') || '') === 'application/oebps-package+xml') || rootfiles[0];
        if (!raiz) marca('error', 'container-sin-rootfile', 'META-INF/container.xml');
        else {
          rutaOpf = raiz.getAttribute('full-path') || '';
          if (!rutaOpf || !existe(rutaOpf)) { marca('error', 'opf-no-existe', 'META-INF/container.xml', { ruta: rutaOpf }); rutaOpf = ''; }
        }
      }
    }
    if (!rutaOpf) return cerrar();

    // --- OPF ---
    const x = leerXml(await leer(rutaOpf));
    if (x.error) { marca('error', 'xml-mal-formado', rutaOpf, x.error); return cerrar(); }
    const paquete = x.doc.documentElement;
    if (paquete.localName !== 'package') { marca('error', 'opf-sin-package', rutaOpf); return cerrar(); }
    const version = paquete.getAttribute('version') || '';
    resultado.version = version;
    const epub3 = /^3(\.|$)/.test(version);
    const epub2 = /^2(\.|$)/.test(version);
    if (!epub2 && !epub3) marca('aviso', 'version-desconocida', rutaOpf, { version: version || '—' });

    const metadatos = hijos(paquete, 'metadata')[0];
    if (!metadatos) marca('error', 'sin-metadata', rutaOpf);
    else {
      const ids = todos(metadatos, 'identifier');
      const titulo = todos(metadatos, 'title').map(txt).find(Boolean);
      const idioma = todos(metadatos, 'language').map(txt).find(Boolean);
      resultado.titulo = titulo || '';
      resultado.idioma = idioma || '';
      if (!titulo) marca('error', 'sin-titulo', rutaOpf);
      if (!idioma) marca('error', 'sin-idioma', rutaOpf);
      else if (!/^[a-z]{2,3}(-[a-z0-9]{2,8})*$/i.test(idioma)) marca('aviso', 'idioma-raro', rutaOpf, { idioma });
      if (!ids.length || !ids.some((i) => txt(i))) marca('error', 'sin-identificador', rutaOpf);
      const unico = paquete.getAttribute('unique-identifier');
      if (!unico) marca('error', 'sin-unique-identifier', rutaOpf);
      else if (!ids.some((i) => i.getAttribute('id') === unico)) marca('error', 'unique-identifier-roto', rutaOpf, { id: unico });
      if (epub3) {
        const modificado = todos(metadatos, 'meta').find((m) => m.getAttribute('property') === 'dcterms:modified');
        if (!modificado) marca('error', 'sin-modificado', rutaOpf);
        else if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/.test(txt(modificado))) marca('aviso', 'modificado-formato', rutaOpf, { valor: txt(modificado) });
      }
    }

    // --- manifest ---
    const carpetaOpf = rutaOpf.includes('/') ? rutaOpf.slice(0, rutaOpf.lastIndexOf('/') + 1) : '';
    const manifest = hijos(paquete, 'manifest')[0];
    const items = new Map(); // id → { id, href, ruta, tipo, props }
    const rutasListadas = new Set([rutaOpf]);
    if (!manifest) marca('error', 'falta-manifest', rutaOpf);
    else {
      for (const it of hijos(manifest, 'item')) {
        const id = it.getAttribute('id') || '';
        const href = it.getAttribute('href') || '';
        const tipo = (it.getAttribute('media-type') || '').trim().toLowerCase();
        const props = (it.getAttribute('properties') || '').split(/\s+/).filter(Boolean);
        if (!id) marca('error', 'item-sin-id', rutaOpf, { href });
        if (!href) { marca('error', 'item-sin-href', rutaOpf, { id }); continue; }
        if (!tipo) marca('error', 'item-sin-tipo', rutaOpf, { id });
        if (id && items.has(id)) marca('error', 'id-repetido-opf', rutaOpf, { id });
        const remoto = esRemota(href);
        const { ruta } = remoto ? { ruta: href } : resolver(carpetaOpf + 'x', href);
        const datos = { id, href, ruta, tipo, props, remoto, fallback: it.getAttribute('fallback') || '' };
        if (id && !items.has(id)) items.set(id, datos);
        if (remoto) {
          if (!props.includes('remote-resources')) marca('aviso', 'recurso-remoto-manifest', rutaOpf, { href });
          continue;
        }
        rutasListadas.add(ruta);
        if (!existe(ruta)) marca('error', 'manifest-falta', rutaOpf, { href });
      }
    }
    resultado.recursos = items.size;

    // Contenido de los archivos listados: vacíos y tipo declarado contra el real.
    for (const it of items.values()) {
      if (it.remoto || !existe(it.ruta)) continue;
      const o = zip.files[it.ruta]._origen;
      if (o && o.real === 0) { marca('aviso', 'archivo-vacio', it.ruta); continue; }
      if (/^image\//.test(it.tipo)) {
        const bytes = await zip.files[it.ruta].async('uint8array');
        const real = tipoPorBytes(bytes);
        if (!real) marca('aviso', 'imagen-ilegible', it.ruta, { tipo: it.tipo });
        else if (real !== it.tipo && !(it.tipo === 'image/jpg' && real === 'image/jpeg')) marca('aviso', 'tipo-no-coincide', it.ruta, { declarado: it.tipo, real });
      }
      if (it.tipo === 'text/html' || (it.tipo === 'application/xhtml+xml' && /\.html?$/i.test(it.ruta))) marca('aviso', 'extension-html', it.ruta, { tipo: it.tipo });
    }
    for (const n of archivos) {
      if (n === 'mimetype' || n.startsWith('META-INF/') || rutasListadas.has(n)) continue;
      marca('aviso', 'sin-manifest', n);
    }

    // --- spine ---
    const spine = hijos(paquete, 'spine')[0];
    const ordenLectura = [];
    if (!spine) marca('error', 'sin-spine', rutaOpf);
    else {
      const refs = hijos(spine, 'itemref');
      if (!refs.length) marca('error', 'spine-vacio', rutaOpf);
      for (const r of refs) {
        const idref = r.getAttribute('idref') || '';
        const it = items.get(idref);
        if (!it) { marca('error', 'spine-ref-rota', rutaOpf, { idref }); continue; }
        ordenLectura.push(it);
        if (!TIPOS_CAPITULO.has(it.tipo) && !it.fallback) marca('error', 'spine-tipo', rutaOpf, { id: idref, tipo: it.tipo });
      }
      const toc = spine.getAttribute('toc');
      if (toc && !items.has(toc)) marca('error', 'spine-toc-roto', rutaOpf, { id: toc });
      else if (epub2 && !toc) marca('error', 'epub2-sin-toc', rutaOpf);
    }
    resultado.capitulos = ordenLectura.length;

    // Portada
    const portada = [...items.values()].some((i) => i.props.includes('cover-image'))
      || (metadatos && todos(metadatos, 'meta').some((m) => m.getAttribute('name') === 'cover' && items.has(m.getAttribute('content'))));
    if (!portada) marca('aviso', 'sin-portada', rutaOpf);
    if (metadatos) {
      for (const m of todos(metadatos, 'meta')) {
        if (m.getAttribute('name') === 'cover' && !items.has(m.getAttribute('content') || '')) marca('aviso', 'portada-rota', rutaOpf, { id: m.getAttribute('content') || '' });
      }
    }

    // --- Capítulos: se leen todos antes de revisar enlaces entre ellos ---
    const capitulos = new Map(); // ruta → { doc, ids, item }
    const paginas = [...items.values()].filter((i) => i.tipo === 'application/xhtml+xml' && !i.remoto && existe(i.ruta));
    for (const it of paginas) {
      const r = leerXml(await leer(it.ruta), 'application/xhtml+xml');
      if (r.error) { marca('error', 'xml-mal-formado', it.ruta, r.error); continue; }
      const raiz = r.doc.documentElement;
      if (raiz.namespaceURI !== NS_XHTML) marca('error', 'xhtml-espacio', it.ruta);
      const ids = new Map();
      for (const el of r.doc.querySelectorAll('[id]')) {
        const id = el.getAttribute('id');
        if (ids.has(id)) marca('error', 'id-repetido', it.ruta, { id });
        ids.set(id, true);
      }
      capitulos.set(it.ruta, { doc: r.doc, ids, item: it });
    }

    // --- Índice ---
    const enlaceAlLibro = (desde, href, codigo, ruta) => {
      if (!href || esRemota(href)) return;
      const { ruta: destino, fragmento } = resolver(desde, href);
      if (!existe(destino)) { marca('aviso', codigo, ruta, { href }); return; }
      const cap = capitulos.get(destino);
      if (fragmento && cap && !cap.ids.has(decodeURIComponent(fragmento))) marca('aviso', 'indice-ancla-rota', ruta, { href });
    };
    if (epub3) {
      const navs = [...items.values()].filter((i) => i.props.includes('nav'));
      if (navs.length !== 1) marca('error', navs.length ? 'nav-varios' : 'sin-nav', rutaOpf);
      for (const nv of navs.slice(0, 1)) {
        const c = capitulos.get(nv.ruta);
        if (!c) { if (existe(nv.ruta) && nv.tipo !== 'application/xhtml+xml') marca('error', 'nav-tipo', nv.ruta, { tipo: nv.tipo }); continue; }
        const indices = todos(c.doc, 'nav').filter((n) => ((n.getAttributeNS(NS_OPS, 'type') || n.getAttribute('epub:type') || '').split(/\s+/)).includes('toc'));
        if (!indices.length) marca('error', 'nav-sin-toc', nv.ruta);
        else {
          if (!indices[0].getElementsByTagNameNS('*', 'a').length) marca('aviso', 'nav-vacio', nv.ruta);
          for (const a of todos(indices[0], 'a')) enlaceAlLibro(nv.ruta, a.getAttribute('href') || '', 'indice-enlace-roto', nv.ruta);
        }
      }
    }
    for (const nc of [...items.values()].filter((i) => i.tipo === 'application/x-dtbncx+xml' && !i.remoto && existe(i.ruta))) {
      const r = leerXml(await leer(nc.ruta));
      if (r.error) { marca('error', 'xml-mal-formado', nc.ruta, r.error); continue; }
      const puntos = todos(r.doc, 'navPoint');
      if (!puntos.length) marca('aviso', 'ncx-vacio', nc.ruta);
      for (const p of puntos) {
        const contenido = hijos(p, 'content')[0];
        if (contenido) enlaceAlLibro(nc.ruta, contenido.getAttribute('src') || '', 'indice-enlace-roto', nc.ruta);
      }
    }

    // --- Cada capítulo ---
    for (const [ruta, { doc, item }] of capitulos) {
      const cabeza = todos(doc, 'head')[0];
      const titulo = cabeza ? todos(cabeza, 'title')[0] : null;
      if (!titulo || !txt(titulo)) marca('aviso', 'capitulo-sin-titulo', ruta);
      let sinAlt = 0;
      for (const img of todos(doc, 'img')) {
        const src = img.getAttribute('src') || '';
        if (!img.hasAttribute('alt')) sinAlt++;
        if (!src) { marca('aviso', 'img-sin-src', ruta); continue; }
        if (/^data:/i.test(src)) continue;
        if (esRemota(src)) { if (!item.props.includes('remote-resources')) marca('aviso', 'recurso-remoto', ruta, { href: src }); continue; }
        if (!existe(resolver(ruta, src).ruta)) marca('error', 'recurso-falta', ruta, { href: src });
      }
      if (sinAlt) marca('aviso', 'img-sin-alt', ruta, { n: sinAlt });
      for (const enlace of todos(doc, 'link')) {
        const href = enlace.getAttribute('href') || '';
        if (!href || !/stylesheet/i.test(enlace.getAttribute('rel') || '')) continue;
        if (esRemota(href)) { if (!item.props.includes('remote-resources')) marca('aviso', 'recurso-remoto', ruta, { href }); continue; }
        if (!existe(resolver(ruta, href).ruta)) marca('error', 'recurso-falta', ruta, { href });
      }
      for (const a of todos(doc, 'a')) {
        const href = (a.getAttribute('href') || '').trim();
        if (!href) continue;
        if (/^javascript:/i.test(href)) { marca('aviso', 'enlace-javascript', ruta); continue; }
        if (esRemota(href)) continue;
        const { ruta: destino, fragmento } = resolver(ruta, href);
        if (!existe(destino)) { marca('aviso', 'enlace-roto', ruta, { href }); continue; }
        const cap = capitulos.get(destino);
        let ancla = fragmento;
        try { ancla = decodeURIComponent(fragmento); } catch (e) { /* se deja como viene */ }
        if (ancla && cap && !cap.ids.has(ancla)) marca('aviso', 'ancla-rota', ruta, { href });
      }
      if (todos(doc, 'script').length && !item.props.includes('scripted')) marca('aviso', 'script', ruta);
      if (todos(doc, 'iframe').length) marca('aviso', 'iframe', ruta);
      if (todos(doc, 'form').length) marca('aviso', 'formulario', ruta);
      // Estilos escritos dentro del capítulo
      for (const est of todos(doc, 'style')) revisarCss(est.textContent || '', ruta, ruta);
    }

    // --- Hojas de estilo ---
    function revisarCss(css, desde, ruta) {
      const sinComentarios = css.replace(/\/\*[\s\S]*?\*\//g, '');
      const pedidos = [...sinComentarios.matchAll(/url\(\s*(['"]?)([^'")]+)\1\s*\)|@import\s+(['"])([^'"]+)\3/gi)].map((m) => (m[2] || m[4] || '').trim());
      for (const p of pedidos) {
        if (!p || /^(data|about|#)/i.test(p)) continue;
        if (esRemota(p)) { marca('aviso', 'css-remoto', ruta, { href: p }); continue; }
        if (!existe(resolver(desde, p).ruta)) marca('aviso', 'css-recurso-falta', ruta, { href: p });
      }
    }
    for (const it of items.values()) {
      if (it.tipo === 'text/css' && !it.remoto && existe(it.ruta)) revisarCss(await leer(it.ruta), it.ruta, it.ruta);
    }

    return cerrar();
  }

  return { validar };
})();

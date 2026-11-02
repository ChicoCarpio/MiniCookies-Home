/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Helpers compartidos para leer y construir archivos EPUB. Un .epub es un .zip
// con estructura fija (mimetype sin comprimir, META-INF/container.xml, un OPF
// con manifest/spine, capítulos en XHTML) — no hace falta una librería de EPUB
// dedicada, alcanza con ZipPropio + DOMParser (mismo criterio que el resto del
// proyecto: no sumar una dependencia si ya se puede resolver con lo que hay).
// Requiere que la página ya haya cargado ZipPropio antes de usar estas funciones.

const ESPACIO_NOMBRES_XHTML = 'http://www.w3.org/1999/xhtml';

function generarUuidEpub() {
  return 'urn:uuid:' + 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
    const r = crypto.getRandomValues(new Uint8Array(1))[0] % 16;
    const v = c === 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

function escaparXmlEpub(texto) {
  return String(texto ?? '').replace(/[&<>"']/g, (c) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;',
  }[c]));
}

// --- Construir un .epub desde capítulos ya en HTML ---

// Saca las imágenes data:URI de un HTML y las deja como archivos aparte del
// paquete EPUB (algunos lectores de ebooks no soportan bien data:URI dentro
// del XHTML de un capítulo, así que se extraen a archivos reales referenciados
// por ruta relativa, que es lo que espera el estándar EPUB).
function extraerImagenesDataUriEpub(html, prefijo) {
  const imagenes = [];
  let contador = 0;
  const htmlNuevo = html.replace(/src="data:image\/(png|jpe?g|gif);base64,([^"]+)"/gi, (coincide, tipo, base64) => {
    tipo = tipo.toLowerCase();
    const extension = tipo === 'jpeg' ? 'jpg' : tipo;
    const nombre = `images/${prefijo}-img${++contador}.${extension}`;
    imagenes.push({ nombre, tipo: `image/${tipo === 'jpg' ? 'jpeg' : tipo}`, base64 });
    return `src="${nombre}"`;
  });
  return { html: htmlNuevo, imagenes };
}

// Un EPUB exige XHTML estricto: el HTML que sale de innerHTML, el conversor de Markdown o
// el lector de Word (docx-lector.js) trae <br>, <img>, <hr> sin cerrar y entidades como &nbsp;, que un
// lector estricto (Apple Books, epubcheck) rechaza o muestra roto. Se
// reserializa con XMLSerializer, que cierra todo y convierte las entidades.
function htmlAXhtmlEpub(html) {
  const doc = new DOMParser().parseFromString(`<!DOCTYPE html><html><body>${html}</body></html>`, 'text/html');
  const serializador = new XMLSerializer();
  return [...doc.body.childNodes].map((nodo) => serializador.serializeToString(nodo)).join('')
    // El espacio de nombres ya lo declara el <html> del capítulo.
    .replace(/ xmlns="http:\/\/www\.w3\.org\/1999\/xhtml"/g, '');
}

// contenidoAParrafos-style: recibe { titulo, autor, idioma, identificador, capitulos: [{titulo, html}] }
// y devuelve un Blob del .epub listo para descargar. "identificador": el del
// libro original cuando se lo modifica (ISBN, uuid...), para que un lector o
// una tienda lo siga reconociendo como el mismo libro; si falta, se genera uno.
async function construirEpub({ titulo, autor, idioma = 'es', identificador, capitulos }) {
  if (!capitulos || capitulos.length === 0) throw new Error('No hay contenido para armar el EPUB.');

  const zip = new ZipPropio();
  const uuid = identificador || generarUuidEpub();
  const fecha = new Date().toISOString().replace(/\.\d+Z$/, 'Z');

  zip.file('mimetype', 'application/epub+zip', { compression: 'STORE' });

  zip.file('META-INF/container.xml', `<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles>
    <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>
  </rootfiles>
</container>`);

  zip.file('OEBPS/styles.css', `body { font-family: serif; line-height: 1.5; margin: 1em; }
h1, h2, h3 { font-family: sans-serif; }
h1 { font-size: 1.5em; margin-top: 1.4em; page-break-before: always; }
img { max-width: 100%; }
table { border-collapse: collapse; }
table td, table th { border: 1px solid #999; padding: 4px 8px; }`);

  const manifestItems = ['<item id="css" href="styles.css" media-type="text/css"/>'];
  const spineItems = [];
  const navItems = [];
  const navPoints = [];

  capitulos.forEach((capitulo, indice) => {
    const numero = indice + 1;
    const idCapitulo = `cap${numero}`;
    const nombreArchivo = `chapter${numero}.xhtml`;
    const tituloCapitulo = capitulo.titulo || `Capítulo ${numero}`;

    const { html: htmlSinImagenes, imagenes } = extraerImagenesDataUriEpub(capitulo.html || '', idCapitulo);
    imagenes.forEach((imagen, indiceImagen) => {
      zip.file(`OEBPS/${imagen.nombre}`, imagen.base64, { base64: true });
      manifestItems.push(`<item id="${idCapitulo}-img${indiceImagen + 1}" href="${imagen.nombre}" media-type="${imagen.tipo}"/>`);
    });

    zip.file(`OEBPS/${nombreArchivo}`, `<?xml version="1.0" encoding="UTF-8"?>
<html xmlns="${ESPACIO_NOMBRES_XHTML}" lang="${escaparXmlEpub(idioma || 'es')}">
<head><title>${escaparXmlEpub(tituloCapitulo)}</title><link rel="stylesheet" type="text/css" href="styles.css"/></head>
<body>${htmlSinImagenes ? htmlAXhtmlEpub(htmlSinImagenes) : `<h1>${escaparXmlEpub(tituloCapitulo)}</h1>`}</body>
</html>`);

    manifestItems.push(`<item id="${idCapitulo}" href="${nombreArchivo}" media-type="application/xhtml+xml"/>`);
    spineItems.push(`<itemref idref="${idCapitulo}"/>`);
    navItems.push(`<li><a href="${nombreArchivo}">${escaparXmlEpub(tituloCapitulo)}</a></li>`);
    navPoints.push(`<navPoint id="np${numero}" playOrder="${numero}"><navLabel><text>${escaparXmlEpub(tituloCapitulo)}</text></navLabel><content src="${nombreArchivo}"/></navPoint>`);
  });

  zip.file('OEBPS/nav.xhtml', `<?xml version="1.0" encoding="UTF-8"?>
<html xmlns="${ESPACIO_NOMBRES_XHTML}" xmlns:epub="http://www.idpf.org/2007/ops">
<head><title>Índice</title></head>
<body><nav epub:type="toc"><h1>Índice</h1><ol>${navItems.join('')}</ol></nav></body>
</html>`);
  manifestItems.push('<item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>');

  zip.file('OEBPS/toc.ncx', `<?xml version="1.0" encoding="UTF-8"?>
<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">
  <head><meta name="dtb:uid" content="${escaparXmlEpub(uuid)}"/></head>
  <docTitle><text>${escaparXmlEpub(titulo)}</text></docTitle>
  <navMap>${navPoints.join('')}</navMap>
</ncx>`);
  manifestItems.push('<item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>');

  zip.file('OEBPS/content.opf', `<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="uid">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:identifier id="uid">${escaparXmlEpub(uuid)}</dc:identifier>
    <dc:title>${escaparXmlEpub(titulo || 'Documento')}</dc:title>
    <dc:language>${escaparXmlEpub(idioma || 'es')}</dc:language>
    ${autor ? `<dc:creator>${escaparXmlEpub(autor)}</dc:creator>` : ''}
    <meta property="dcterms:modified">${fecha}</meta>
  </metadata>
  <manifest>${manifestItems.join('')}</manifest>
  <spine toc="ncx">${spineItems.join('')}</spine>
</package>`);

  return zip.generateAsync({ type: 'blob', mimeType: 'application/epub+zip', compression: 'DEFLATE' });
}

// Corta un HTML (del lector de Word o del conversor de Markdown) en capítulos, uno por cada <h1> que
// encuentra — si no hay ningún <h1>, todo el contenido queda como un único
// capítulo. El contenido antes del primer <h1> (si lo hay) se agrega como un
// capítulo inicial sin título propio. El <h1> queda también al principio del
// capítulo (antes solo se usaba como nombre en el índice y el lector mostraba
// el capítulo sin su título).
function dividirHtmlEnCapitulosEpub(html, tituloPorDefecto) {
  const contenedor = document.createElement('div');
  contenedor.innerHTML = html;

  const tieneEncabezados = contenedor.querySelector('h1') !== null;
  if (!tieneEncabezados) {
    return [{ titulo: tituloPorDefecto, html: contenedor.innerHTML }];
  }

  const capitulos = [];
  let actual = null;
  [...contenedor.childNodes].forEach((nodo) => {
    if (nodo.nodeType === Node.ELEMENT_NODE && nodo.tagName === 'H1') {
      actual = { titulo: nodo.textContent.trim() || `Capítulo ${capitulos.length + 1}`, nodos: [nodo] };
      capitulos.push(actual);
    } else if (actual) {
      actual.nodos.push(nodo);
    } else {
      actual = { titulo: tituloPorDefecto, nodos: [nodo] };
      capitulos.push(actual);
    }
  });

  return capitulos.map((capitulo) => {
    const div = document.createElement('div');
    capitulo.nodos.forEach((nodo) => div.appendChild(nodo.cloneNode(true)));
    return { titulo: capitulo.titulo, html: div.innerHTML };
  });
}

// --- Leer un .epub existente ---

// DOMParser sobre XML conserva el nombre de la etiqueta tal cual viene escrito
// (con el prefijo, ej. "dc:title"), así que alcanza con getElementsByTagName
// para leer un OPF real sin tener que resolver namespaces XML de verdad.
function resolverRutaRelativaEpub(base, relativa) {
  if (/^(https?:|data:)/i.test(relativa)) return relativa;
  const url = new URL(relativa.replace(/^\.\//, ''), 'zip://raiz/' + base);
  return decodeURIComponent(url.pathname.replace(/^\//, ''));
}

async function leerEpub(archivo) {
  const zip = await ZipPropio.loadAsync(archivo);

  await fuentesOfuscadasEpub(zip); // corta si hay DRM de verdad

  const parser = new DOMParser();
  const entradaContainer = zip.file('META-INF/container.xml');
  if (!entradaContainer) throw new Error(window.t ? t('epub-sin-container') : 'El EPUB no tiene un container.xml válido.');
  const containerXml = await entradaContainer.async('string');
  const containerDoc = parser.parseFromString(containerXml, 'application/xml');
  const rutaOpf = containerDoc.getElementsByTagName('rootfile')[0]?.getAttribute('full-path');
  if (!rutaOpf) throw new Error(window.t ? t('epub-sin-container') : 'El EPUB no tiene un container.xml válido.');
  const carpetaOpf = rutaOpf.includes('/') ? rutaOpf.slice(0, rutaOpf.lastIndexOf('/') + 1) : '';

  const opfXml = await zip.file(rutaOpf).async('string');
  const opfDoc = parser.parseFromString(opfXml, 'application/xml');

  const titulo = opfDoc.getElementsByTagName('dc:title')[0]?.textContent?.trim() || (window.t ? t('epub-sin-titulo') : 'Sin título');
  const autor = opfDoc.getElementsByTagName('dc:creator')[0]?.textContent?.trim() || '';
  const idioma = opfDoc.getElementsByTagName('dc:language')[0]?.textContent?.trim() || '';
  // El identificador principal es el que nombra "unique-identifier" en <package>.
  const idUnico = opfDoc.documentElement.getAttribute('unique-identifier');
  const identificadores = [...opfDoc.getElementsByTagName('dc:identifier')];
  const identificador = ((idUnico && identificadores.find((e) => e.getAttribute('id') === idUnico)) || identificadores[0])?.textContent?.trim() || '';

  const manifest = {};
  [...opfDoc.getElementsByTagName('item')].forEach((item) => {
    manifest[item.getAttribute('id')] = {
      href: carpetaOpf + item.getAttribute('href'),
      tipo: item.getAttribute('media-type'),
    };
  });

  const idsSpine = [...opfDoc.getElementsByTagName('itemref')].map((ref) => ref.getAttribute('idref'));
  const capitulos = [];

  for (const idItem of idsSpine) {
    const item = manifest[idItem];
    if (!item || !/xhtml|html/i.test(item.tipo)) continue;

    const archivoZip = zip.file(item.href);
    if (!archivoZip) continue;

    const xhtml = await archivoZip.async('string');
    const doc = parser.parseFromString(xhtml, 'application/xhtml+xml');
    const cuerpo = doc.querySelector('body');
    if (!cuerpo) continue;

    const carpetaCapitulo = item.href.includes('/') ? item.href.slice(0, item.href.lastIndexOf('/') + 1) : '';

    // Incrusta cada imagen referenciada como data:URI, resolviendo la ruta
    // relativa al archivo del capítulo (no a la del OPF).
    // Las portadas suelen venir envueltas en SVG (<image xlink:href>), no en <img>:
    // se pasan a <img> para que la misma incrustación las alcance.
    for (const svgImagen of [...cuerpo.querySelectorAll('image')]) {
      const href = svgImagen.getAttributeNS('http://www.w3.org/1999/xlink', 'href') || svgImagen.getAttribute('href') || '';
      if (!href) continue;
      const img = cuerpo.ownerDocument.createElementNS('http://www.w3.org/1999/xhtml', 'img');
      img.setAttribute('src', href);
      img.setAttribute('alt', 'Portada');
      // Si el <svg> solo envuelve esta imagen, se reemplaza entero: vacío seguía
      // ocupando su alto (una hoja en blanco antes de la portada en EPUB → PDF).
      const svg = svgImagen.closest('svg');
      const soloImagen = svg && [...svg.querySelectorAll('*')].every((n) => n === svgImagen || /^(title|desc|defs)$/i.test(n.localName));
      (soloImagen ? svg : svgImagen).replaceWith(img);
    }
    const imagenes = [...cuerpo.querySelectorAll('img')];
    for (const img of imagenes) {
      const src = img.getAttribute('src') || '';
      if (!src || /^(https?:|data:)/i.test(src)) continue;
      const rutaImagen = resolverRutaRelativaEpub(carpetaCapitulo, src);
      const archivoImagen = zip.file(rutaImagen);
      if (!archivoImagen) continue;
      const base64 = await archivoImagen.async('base64');
      const extension = (rutaImagen.split('.').pop() || '').toLowerCase();
      const tipoMime = extension === 'svg' ? 'image/svg+xml' : `image/${extension === 'jpg' ? 'jpeg' : extension}`;
      img.setAttribute('src', `data:${tipoMime};base64,${base64}`);
    }

    const primerEncabezado = cuerpo.querySelector('h1, h2');
    const tituloCapitulo = primerEncabezado ? primerEncabezado.textContent.trim() : `Capítulo ${capitulos.length + 1}`;
    // El XHTML del libro viene de un archivo externo: se limpia acá (sobre el
    // documento inerte de DOMParser) antes de que cualquier herramienta lo
    // meta con innerHTML — ver assets/js/seguridad-html.js.
    sanitizarArbolHtml(cuerpo);
    capitulos.push({ titulo: tituloCapitulo, html: cuerpo.innerHTML });
  }

  if (capitulos.length === 0) throw new Error(window.t ? t('epub-sin-contenido') : 'No se encontró contenido legible dentro del EPUB.');

  return { titulo, autor, idioma, identificador, capitulos };
}

// Un libro → un PDF, sin imprimir: se arma en una hoja A4 fuera de pantalla
// (portada y cada capítulo en hoja nueva) y html-a-pdf.js lo escribe.
// Compartido por EPUB → PDF y Unir PDF; requiere además seguridad-html.js,
// pdf-lib y html-a-pdf.js. Devuelve los bytes del PDF.
async function epubAPdf(archivo, alAvanzar) {
  const { titulo, autor, capitulos } = await leerEpub(archivo);
  const lienzo = HtmlAPdf.crearLienzo('libro');
  const portada = `<div class="portada"><h1>${escaparHtml(titulo)}</h1>${autor ? `<p>${escaparHtml(autor)}</p>` : ''}</div>`;
  lienzo.innerHTML = portada + capitulos.map((cap) => `<section class="capitulo">${cap.html}</section>`).join('');
  try {
    await Promise.all([...lienzo.querySelectorAll('img')].map((img) => img.decode().catch(() => {})));
    return await HtmlAPdf.convertir([{ elemento: lienzo, anchoPx: 794, altoPx: 1123, margenArribaPx: 76, margenAbajoPx: 76 }], { titulo, alAvanzar });
  } finally {
    lienzo.remove();
  }
}

// ===========================================================================
// --- Editar un .epub en su lugar ---
// Administrar, Agregar texto, Dividir, Unir y Reparar leían el libro con
// leerEpub() (solo el texto de cada capítulo) y lo rearmaban con
// construirEpub(): se perdían la hoja de estilos, las fuentes, la portada y
// los metadatos extra. Esto trabaja sobre el paquete original —el mismo .zip—
// y toca solo lo que la herramienta cambia: el spine (orden de lectura), los
// índices (nav.xhtml y toc.ncx), los capítulos afectados y algunos metadatos.
// Mismo criterio que editar-metadatos-epub.html.
// ===========================================================================

const NS_OPF_EPUB = 'http://www.idpf.org/2007/opf';
const NS_DC_EPUB = 'http://purl.org/dc/elements/1.1/';
const NS_EPUB_OPS = 'http://www.idpf.org/2007/ops';
const NS_NCX_EPUB = 'http://www.daisy.org/z3986/2005/ncx/';

function textoEpub(clave, porDefecto) {
  return window.t ? t(clave) : porDefecto;
}

function parsearXmlEpub(texto, ruta) {
  const doc = new DOMParser().parseFromString(texto, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) {
    throw new Error(textoEpub('epub-xml-invalido', 'Un archivo interno del EPUB no es XML válido') + ': ' + ruta);
  }
  return doc;
}

function carpetaDeRutaEpub(ruta) {
  return ruta.includes('/') ? ruta.slice(0, ruta.lastIndexOf('/') + 1) : '';
}

// Ruta completa dentro del .zip a partir de un href (sin el #fragmento).
function rutaDesdeHrefEpub(carpeta, href) {
  if (!href || /^[a-z][a-z0-9+.-]*:/i.test(href)) return null; // http:, mailto:, data:...
  const sinFragmento = href.split('#')[0];
  return sinFragmento ? resolverRutaRelativaEpub(carpeta, sinFragmento) : null;
}

// href relativo desde una carpeta hasta una ruta del .zip (para escribir en
// el OPF o en un índice que vive en otra carpeta).
function hrefRelativoEpub(desdeCarpeta, hastaRuta) {
  const desde = desdeCarpeta.split('/').filter(Boolean);
  const hasta = hastaRuta.split('/');
  let comun = 0;
  while (comun < desde.length && comun < hasta.length - 1 && desde[comun] === hasta[comun]) comun++;
  const subir = desde.slice(comun).map(() => '..');
  return subir.concat(hasta.slice(comun)).map((parte) => (parte === '..' ? parte : encodeURIComponent(parte))).join('/');
}

// --- Fuentes ofuscadas vs. DRM ---
// META-INF/encryption.xml no siempre es DRM: muchos libros comerciales
// "ofuscan" sus fuentes incrustadas (un XOR de los primeros bytes con una
// clave sacada del identificador del libro) para que no se copien sueltas.
// El contenido se lee igual. Antes cualquier encryption.xml se rechazaba como
// DRM y esos libros no se podían abrir. Ahora solo se rechaza si cifra algo
// que no sea una fuente con uno de esos dos algoritmos.
const OFUSCACION_IDPF_EPUB = 'http://www.idpf.org/2008/embedding';
const OFUSCACION_ADOBE_EPUB = 'http://ns.adobe.com/pdf/enc#RC';

async function fuentesOfuscadasEpub(zip) {
  const entrada = zip.file('META-INF/encryption.xml');
  if (!entrada) return [];
  const doc = new DOMParser().parseFromString(await entrada.async('string'), 'application/xml');
  const ofuscadas = [];
  for (const dato of [...doc.getElementsByTagNameNS('*', 'EncryptedData')]) {
    const metodo = dato.getElementsByTagNameNS('*', 'EncryptionMethod')[0]?.getAttribute('Algorithm');
    const ruta = dato.getElementsByTagNameNS('*', 'CipherReference')[0]?.getAttribute('URI');
    const comprimido = dato.getElementsByTagNameNS('*', 'Compression').length > 0;
    if (!ruta || comprimido || (metodo !== OFUSCACION_IDPF_EPUB && metodo !== OFUSCACION_ADOBE_EPUB)) {
      throw new Error(textoEpub('epub-drm', 'Este EPUB está protegido con DRM (cifrado) — no se puede leer su contenido en el navegador.'));
    }
    ofuscadas.push({ ruta: decodeURIComponent(ruta), metodo });
  }
  return ofuscadas;
}

// Deshace la ofuscación (el XOR es su propio inverso). Hace falta antes de
// cambiar el identificador del libro (Dividir, Unir): la clave sale de él, y
// con otro identificador las fuentes quedarían ilegibles.
async function desofuscarFuentesEpub(paquete) {
  const ofuscadas = await fuentesOfuscadasEpub(paquete.zip);
  if (!ofuscadas.length) return;
  const idUnico = paquete.opf.documentElement.getAttribute('unique-identifier');
  const identificadores = [...paquete.opf.getElementsByTagNameNS(NS_DC_EPUB, 'identifier')];
  const principal = ((idUnico && identificadores.find((e) => e.getAttribute('id') === idUnico)) || identificadores[0])?.textContent || '';
  for (const { ruta, metodo } of ofuscadas) {
    const archivo = paquete.zip.file(ruta);
    if (!archivo) continue;
    let clave;
    let largo;
    if (metodo === OFUSCACION_IDPF_EPUB) {
      clave = new Uint8Array(await crypto.subtle.digest('SHA-1', new TextEncoder().encode(principal.replace(/[ \u0009\u000d\u000a]/g, ''))));
      largo = 1040;
    } else {
      const uuid = identificadores.map((e) => e.textContent.trim()).find((v) => /^urn:uuid:/i.test(v)) || principal;
      const hex = uuid.replace(/^urn:uuid:/i, '').replace(/[^0-9a-f]/gi, '');
      if (hex.length < 32) continue;
      clave = Uint8Array.from(hex.slice(0, 32).match(/../g), (par) => parseInt(par, 16));
      largo = 1024;
    }
    const bytes = await archivo.async('uint8array');
    for (let i = 0; i < Math.min(largo, bytes.length); i++) bytes[i] ^= clave[i % clave.length];
    paquete.zip.file(ruta, bytes);
  }
  paquete.zip.remove('META-INF/encryption.xml');
}

// --- Abrir y guardar ---

async function rutaOpfDeZipEpub(zip) {
  const contenedor = zip.file('META-INF/container.xml');
  if (!contenedor) throw new Error(textoEpub('epub-sin-container', 'El EPUB no tiene un container.xml válido.'));
  const raiz = parsearXmlEpub(await contenedor.async('string'), 'META-INF/container.xml').getElementsByTagNameNS('*', 'rootfile')[0];
  const ruta = raiz && decodeURIComponent(raiz.getAttribute('full-path') || '');
  if (!ruta || !zip.file(ruta)) throw new Error(textoEpub('epub-sin-container', 'El EPUB no tiene un container.xml válido.'));
  return ruta;
}

// Devuelve el paquete listo para editar. "archivo" puede ser un File, un Blob
// o bytes: cada herramienta vuelve a abrir el original en cada guardado, así
// una edición nunca arrastra los cambios de la anterior.
async function abrirPaqueteEpub(archivo) {
  const zip = await ZipPropio.loadAsync(archivo);
  await fuentesOfuscadasEpub(zip); // corta si hay DRM de verdad
  const rutaOpf = await rutaOpfDeZipEpub(zip);
  const opf = parsearXmlEpub(await zip.file(rutaOpf).async('string'), rutaOpf);
  return { zip, rutaOpf, carpetaOpf: carpetaDeRutaEpub(rutaOpf), opf, docs: new Map() };
}

// Un XML interno del libro (nav, ncx) como documento editable. Queda anotado
// para escribirse al guardar.
async function documentoDePaqueteEpub(paquete, ruta) {
  if (paquete.docs.has(ruta)) return paquete.docs.get(ruta);
  const archivo = paquete.zip.file(ruta);
  if (!archivo) return null;
  const doc = parsearXmlEpub(await archivo.async('string'), ruta);
  paquete.docs.set(ruta, doc);
  return doc;
}

function serializarXmlEpub(doc) {
  const texto = new XMLSerializer().serializeToString(doc);
  return texto.startsWith('<?xml') ? texto : '<?xml version="1.0" encoding="UTF-8"?>\n' + texto;
}

// Un .epub es un .zip con una regla propia: "mimetype" va primero y sin
// comprimir. Compartido con editar-metadatos-epub.html.
async function generarBlobEpub(zip) {
  const salida = new ZipPropio();
  salida.file('mimetype', 'application/epub+zip', { compression: 'STORE' });
  for (const ruta of Object.keys(zip.files)) {
    const entrada = zip.files[ruta];
    if (entrada.dir || ruta === 'mimetype') continue;
    salida.file(ruta, await entrada.async('uint8array'));
  }
  return salida.generateAsync({ type: 'blob', mimeType: 'application/epub+zip', compression: 'DEFLATE' });
}

async function guardarPaqueteEpub(paquete) {
  const modificado = [...paquete.opf.getElementsByTagNameNS(NS_OPF_EPUB, 'meta')].find((m) => m.getAttribute('property') === 'dcterms:modified');
  if (modificado) modificado.textContent = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  paquete.zip.file(paquete.rutaOpf, serializarXmlEpub(paquete.opf));
  for (const [ruta, doc] of paquete.docs) paquete.zip.file(ruta, serializarXmlEpub(doc));
  return generarBlobEpub(paquete.zip);
}

// --- Lo que hay adentro ---

function itemsManifestEpub(paquete) {
  return [...paquete.opf.getElementsByTagNameNS(NS_OPF_EPUB, 'item')].map((el) => ({
    el,
    id: el.getAttribute('id'),
    ruta: rutaDesdeHrefEpub(paquete.carpetaOpf, el.getAttribute('href') || ''),
    tipo: el.getAttribute('media-type') || '',
    propiedades: (el.getAttribute('properties') || '').split(/\s+/).filter(Boolean),
  }));
}

function spineEpub(paquete) {
  return paquete.opf.getElementsByTagNameNS(NS_OPF_EPUB, 'spine')[0] || null;
}

// Rutas del índice EPUB 3 (nav.xhtml) y del EPUB 2 (toc.ncx), si existen.
function rutasIndicesEpub(paquete) {
  const items = itemsManifestEpub(paquete);
  const nav = items.find((it) => it.propiedades.includes('nav'));
  const idNcx = spineEpub(paquete)?.getAttribute('toc');
  const ncx = items.find((it) => it.id === idNcx) || items.find((it) => it.tipo === 'application/x-dtbncx+xml');
  return { nav: nav ? nav.ruta : null, ncx: ncx ? ncx.ruta : null };
}

function navTocEpub(doc) {
  const navs = [...doc.getElementsByTagNameNS('*', 'nav')];
  return navs.find((n) => (n.getAttributeNS(NS_EPUB_OPS, 'type') || n.getAttribute('epub:type') || '').split(/\s+/).includes('toc')) || navs[0] || null;
}

function hijosDirectosEpub(nodo, nombreLocal) {
  return nodo ? [...nodo.children].filter((n) => n.localName === nombreLocal) : [];
}

// Nombre de cada archivo según el índice del libro (la primera entrada que lo
// nombra), para mostrar los capítulos con el título que les puso el autor.
async function etiquetasDeIndiceEpub(paquete) {
  const etiquetas = new Map();
  const { nav, ncx } = rutasIndicesEpub(paquete);
  const docNav = nav && await documentoDePaqueteEpub(paquete, nav);
  if (docNav) {
    for (const a of navTocEpub(docNav)?.getElementsByTagNameNS('*', 'a') || []) {
      const ruta = rutaDesdeHrefEpub(carpetaDeRutaEpub(nav), a.getAttribute('href') || '');
      if (ruta && !etiquetas.has(ruta) && a.textContent.trim()) etiquetas.set(ruta, a.textContent.trim().replace(/\s+/g, ' '));
    }
  }
  const docNcx = ncx && await documentoDePaqueteEpub(paquete, ncx);
  if (docNcx) {
    for (const punto of docNcx.getElementsByTagNameNS('*', 'navPoint')) {
      const contenido = hijosDirectosEpub(punto, 'content')[0];
      const texto = hijosDirectosEpub(hijosDirectosEpub(punto, 'navLabel')[0], 'text')[0];
      const ruta = contenido && rutaDesdeHrefEpub(carpetaDeRutaEpub(ncx), contenido.getAttribute('src') || '');
      if (ruta && !etiquetas.has(ruta) && texto && texto.textContent.trim()) etiquetas.set(ruta, texto.textContent.trim().replace(/\s+/g, ' '));
    }
  }
  return etiquetas;
}

// Los documentos del spine, en orden de lectura, con un título para mostrar.
async function capitulosDePaqueteEpub(paquete) {
  const porId = new Map(itemsManifestEpub(paquete).map((it) => [it.id, it]));
  const etiquetas = await etiquetasDeIndiceEpub(paquete);
  const capitulos = [];
  for (const ref of hijosDirectosEpub(spineEpub(paquete), 'itemref')) {
    const item = porId.get(ref.getAttribute('idref'));
    if (!item || !item.ruta || !/x?html/i.test(item.tipo) || !paquete.zip.file(item.ruta)) continue;
    let titulo = etiquetas.get(item.ruta);
    if (!titulo) {
      const xhtml = await paquete.zip.file(item.ruta).async('string');
      const encabezado = /<h[1-3][^>]*>([\s\S]*?)<\/h[1-3]>/i.exec(xhtml) || /<title[^>]*>([\s\S]*?)<\/title>/i.exec(xhtml);
      titulo = encabezado ? encabezado[1].replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim() : '';
    }
    capitulos.push({ idref: item.id, ruta: item.ruta, titulo: titulo || item.ruta.split('/').pop().replace(/\.x?html?$/i, '') });
  }
  return capitulos;
}

// --- Cambiar el orden de lectura ---

// Deja en el spine solo los capítulos de "idrefs", en ese orden. Los que
// salen del libro se borran (archivo e item del manifest) salvo que sean el
// índice nav, y los índices se ajustan: sin entradas a archivos que ya no
// están y en el orden nuevo.
async function fijarCapitulosEpub(paquete, idrefs) {
  const spine = spineEpub(paquete);
  const refs = hijosDirectosEpub(spine, 'itemref');
  const porId = new Map(refs.map((ref) => [ref.getAttribute('idref'), ref]));
  const quedan = new Set(idrefs);
  const items = itemsManifestEpub(paquete);
  for (const ref of refs) {
    const id = ref.getAttribute('idref');
    if (quedan.has(id)) continue;
    spine.removeChild(ref);
    const item = items.find((it) => it.id === id);
    if (item && !item.propiedades.includes('nav')) {
      item.el.parentNode.removeChild(item.el);
      if (item.ruta) paquete.zip.remove(item.ruta);
    }
  }
  for (const id of idrefs) if (porId.has(id)) spine.appendChild(porId.get(id));
  await ajustarIndicesEpub(paquete);
}

// Orden de cada archivo en el spine (para reordenar índices).
function ordenSpineEpub(paquete) {
  const porId = new Map(itemsManifestEpub(paquete).map((it) => [it.id, it.ruta]));
  const orden = new Map();
  hijosDirectosEpub(spineEpub(paquete), 'itemref').forEach((ref, i) => {
    const ruta = porId.get(ref.getAttribute('idref'));
    if (ruta && !orden.has(ruta)) orden.set(ruta, i);
  });
  return orden;
}

async function ajustarIndicesEpub(paquete) {
  const orden = ordenSpineEpub(paquete);
  const existe = (ruta) => !!ruta && !!paquete.zip.file(ruta);
  const { nav, ncx } = rutasIndicesEpub(paquete);

  const docNav = nav && await documentoDePaqueteEpub(paquete, nav);
  if (docNav) {
    const carpeta = carpetaDeRutaEpub(nav);
    const destinoDe = (li) => {
      const a = [...li.getElementsByTagNameNS('*', 'a')][0];
      return a ? rutaDesdeHrefEpub(carpeta, a.getAttribute('href') || '') : null;
    };
    // Fuera las entradas (de cualquier nav: índice, puntos de referencia,
    // páginas) que apuntan a un archivo que ya no está.
    for (const a of [...docNav.getElementsByTagNameNS('*', 'a')]) {
      const ruta = rutaDesdeHrefEpub(carpeta, a.getAttribute('href') || '');
      if (!ruta || existe(ruta)) continue;
      const li = a.closest ? a.closest('li') : a.parentNode;
      if (!li) continue;
      const anidada = hijosDirectosEpub(li, 'ol')[0];
      if (anidada && anidada.children.length) {
        // Conserva las sub-entradas: sube sus <li> al lugar de la que se va.
        for (const hijo of [...anidada.children]) li.parentNode.insertBefore(hijo, li);
      }
      li.parentNode.removeChild(li);
    }
    // Una lista no puede quedar vacía (epubcheck): fuera las sub-listas sin
    // entradas, y los nav secundarios (puntos de referencia, páginas) que se
    // quedaron sin ninguna. El índice principal siempre conserva algo.
    const navToc = navTocEpub(docNav);
    for (const lista of [...docNav.getElementsByTagNameNS('*', 'ol')].reverse()) {
      if (hijosDirectosEpub(lista, 'li').length) continue;
      const contenedor = lista.parentNode;
      if (contenedor === navToc) continue;
      contenedor.removeChild(lista);
      if (contenedor.localName === 'nav') contenedor.parentNode.removeChild(contenedor);
    }
    // El índice principal, en el orden nuevo del spine.
    const ol = hijosDirectosEpub(navToc, 'ol')[0];
    if (ol) {
      const entradas = hijosDirectosEpub(ol, 'li');
      entradas
        .map((li, i) => ({ li, i, clave: orden.has(destinoDe(li)) ? orden.get(destinoDe(li)) : Infinity }))
        .sort((x, y) => x.clave - y.clave || x.i - y.i)
        .forEach(({ li }) => ol.appendChild(li));
    }
  }

  const docNcx = ncx && await documentoDePaqueteEpub(paquete, ncx);
  if (docNcx) {
    const carpeta = carpetaDeRutaEpub(ncx);
    const destinoDe = (nodo) => {
      const contenido = hijosDirectosEpub(nodo, 'content')[0];
      return contenido ? rutaDesdeHrefEpub(carpeta, contenido.getAttribute('src') || '') : null;
    };
    for (const nodo of [...docNcx.getElementsByTagNameNS('*', 'navPoint'), ...docNcx.getElementsByTagNameNS('*', 'pageTarget')]) {
      if (!nodo.parentNode || existe(destinoDe(nodo))) continue;
      for (const hijo of hijosDirectosEpub(nodo, 'navPoint')) nodo.parentNode.insertBefore(hijo, nodo);
      nodo.parentNode.removeChild(nodo);
    }
    const mapa = docNcx.getElementsByTagNameNS('*', 'navMap')[0];
    if (mapa) {
      hijosDirectosEpub(mapa, 'navPoint')
        .map((p, i) => ({ p, i, clave: orden.has(destinoDe(p)) ? orden.get(destinoDe(p)) : Infinity }))
        .sort((x, y) => x.clave - y.clave || x.i - y.i)
        .forEach(({ p }) => mapa.appendChild(p));
    }
    [...docNcx.getElementsByTagNameNS('*', 'navPoint')].forEach((p, i) => p.setAttribute('playOrder', String(i + 1)));
  }

  // EPUB 2: <guide> del OPF (portada, inicio del texto...), que tampoco puede
  // quedar vacío. Lo mismo la lista de páginas del toc.ncx.
  for (const referencia of [...paquete.opf.getElementsByTagNameNS(NS_OPF_EPUB, 'reference')]) {
    if (!existe(rutaDesdeHrefEpub(paquete.carpetaOpf, referencia.getAttribute('href') || ''))) referencia.parentNode.removeChild(referencia);
  }
  for (const guia of [...paquete.opf.getElementsByTagNameNS(NS_OPF_EPUB, 'guide')]) {
    if (!hijosDirectosEpub(guia, 'reference').length) guia.parentNode.removeChild(guia);
  }
  for (const lista of docNcx ? [...docNcx.getElementsByTagNameNS('*', 'pageList')] : []) {
    if (!hijosDirectosEpub(lista, 'pageTarget').length) lista.parentNode.removeChild(lista);
  }
}

// --- Recursos que quedaron sin usar ---

// Rutas a las que apunta un documento de texto (XHTML, CSS, SVG): imágenes,
// hojas de estilo, fuentes, otros capítulos.
function referenciasDeTextoEpub(texto, carpeta) {
  const rutas = new Set();
  const patrones = [
    /\b(?:src|href|xlink:href|poster|data)\s*=\s*["']([^"']+)["']/gi,
    /url\(\s*["']?([^"')]+)["']?\s*\)/gi,
    /@import\s+["']([^"']+)["']/gi,
  ];
  for (const patron of patrones) {
    for (const m of texto.matchAll(patron)) {
      const ruta = rutaDesdeHrefEpub(carpeta, m[1].trim());
      if (ruta) rutas.add(ruta);
    }
  }
  return rutas;
}

// Quita del libro los recursos (imágenes, fuentes, estilos, documentos fuera
// del spine) que ya nada usa. Se usa al dividir: sin esto cada parte cargaba
// con todas las imágenes del libro entero.
async function podarRecursosEpub(paquete) {
  const items = itemsManifestEpub(paquete);
  const porRuta = new Map(items.filter((it) => it.ruta).map((it) => [it.ruta, it]));
  const usados = new Set();
  const pendientes = [];
  const usar = (ruta) => { if (ruta && porRuta.has(ruta) && !usados.has(ruta)) { usados.add(ruta); pendientes.push(ruta); } };

  const porId = new Map(items.map((it) => [it.id, it]));
  for (const ref of hijosDirectosEpub(spineEpub(paquete), 'itemref')) usar(porId.get(ref.getAttribute('idref'))?.ruta);
  const { nav, ncx } = rutasIndicesEpub(paquete);
  usar(nav); usar(ncx);
  for (const it of items) if (it.propiedades.includes('cover-image')) usar(it.ruta);
  const metaPortada = [...paquete.opf.getElementsByTagNameNS(NS_OPF_EPUB, 'meta')].find((m) => m.getAttribute('name') === 'cover');
  if (metaPortada) usar(porId.get(metaPortada.getAttribute('content'))?.ruta);
  for (const referencia of paquete.opf.getElementsByTagNameNS(NS_OPF_EPUB, 'reference')) usar(rutaDesdeHrefEpub(paquete.carpetaOpf, referencia.getAttribute('href') || ''));

  while (pendientes.length) {
    const ruta = pendientes.pop();
    const item = porRuta.get(ruta);
    if (item.el.getAttribute('fallback')) usar(porId.get(item.el.getAttribute('fallback'))?.ruta);
    if (!/html|xml|css|svg|ncx/i.test(item.tipo) || !paquete.zip.file(ruta)) continue;
    const texto = paquete.docs.has(ruta) ? serializarXmlEpub(paquete.docs.get(ruta)) : await paquete.zip.file(ruta).async('string');
    for (const destino of referenciasDeTextoEpub(texto, carpetaDeRutaEpub(ruta))) usar(destino);
  }

  for (const it of items) {
    if (!it.ruta || usados.has(it.ruta)) continue;
    it.el.parentNode.removeChild(it.el);
    paquete.zip.remove(it.ruta);
  }
}

// --- Metadatos ---

// Quita un metadato y los <meta refines="#su-id"> que lo describen (rol,
// nombre para ordenar, tipo de identificador): sueltos, epubcheck los marca.
function quitarMetadatoEpub(paquete, nodo) {
  const id = nodo.getAttribute('id');
  if (id) {
    for (const meta of [...paquete.opf.getElementsByTagNameNS(NS_OPF_EPUB, 'meta')]) {
      if (meta.getAttribute('refines') === '#' + id) meta.parentNode.removeChild(meta);
    }
  }
  nodo.parentNode.removeChild(nodo);
}

function fijarMetadatoDcEpub(paquete, nombre, valor) {
  const metadata = paquete.opf.getElementsByTagNameNS(NS_OPF_EPUB, 'metadata')[0];
  const existentes = [...paquete.opf.getElementsByTagNameNS(NS_DC_EPUB, nombre)];
  existentes.slice(1).forEach((n) => quitarMetadatoEpub(paquete, n));
  let nodo = existentes[0];
  if (!nodo) {
    nodo = paquete.opf.createElementNS(NS_DC_EPUB, 'dc:' + nombre);
    metadata.appendChild(nodo);
  }
  nodo.textContent = valor;
  // Un "file-as" o "title-type" refinando al anterior ya no corresponde.
  const id = nodo.getAttribute('id');
  if (id) {
    for (const meta of [...metadata.getElementsByTagNameNS(NS_OPF_EPUB, 'meta')]) {
      if (meta.getAttribute('refines') === '#' + id && /file-as|alternate-script/.test(meta.getAttribute('property') || '')) meta.parentNode.removeChild(meta);
    }
  }
}

function metadatoDcEpub(paquete, nombre) {
  return paquete.opf.getElementsByTagNameNS(NS_DC_EPUB, nombre)[0]?.textContent.trim() || '';
}

async function fijarTituloEpub(paquete, titulo) {
  fijarMetadatoDcEpub(paquete, 'title', titulo);
  const { ncx } = rutasIndicesEpub(paquete);
  const docNcx = ncx && await documentoDePaqueteEpub(paquete, ncx);
  const texto = docNcx && hijosDirectosEpub(docNcx.getElementsByTagNameNS('*', 'docTitle')[0], 'text')[0];
  if (texto) texto.textContent = titulo;
}

// Un libro nuevo (una parte, una unión) necesita otro identificador: si no, un
// lector o una tienda lo confunde con el original. Las fuentes ofuscadas se
// desofuscan antes, porque su clave sale del identificador viejo.
async function nuevoIdentificadorEpub(paquete) {
  await desofuscarFuentesEpub(paquete);
  const idUnico = paquete.opf.documentElement.getAttribute('unique-identifier');
  const identificadores = [...paquete.opf.getElementsByTagNameNS(NS_DC_EPUB, 'identifier')];
  const nodo = (idUnico && identificadores.find((e) => e.getAttribute('id') === idUnico)) || identificadores[0];
  const uuid = generarUuidEpub();
  if (nodo) nodo.textContent = uuid;
  else {
    const nuevo = paquete.opf.createElementNS(NS_DC_EPUB, 'dc:identifier');
    nuevo.setAttribute('id', 'uid');
    nuevo.textContent = uuid;
    paquete.opf.getElementsByTagNameNS(NS_OPF_EPUB, 'metadata')[0].appendChild(nuevo);
    paquete.opf.documentElement.setAttribute('unique-identifier', 'uid');
  }
  // El ISBN u otros identificadores del original ya no describen este libro.
  identificadores.filter((e) => e !== nodo).forEach((e) => quitarMetadatoEpub(paquete, e));
  const { ncx } = rutasIndicesEpub(paquete);
  const docNcx = ncx && await documentoDePaqueteEpub(paquete, ncx);
  const uid = docNcx && [...docNcx.getElementsByTagNameNS('*', 'meta')].find((m) => m.getAttribute('name') === 'dtb:uid');
  if (uid) uid.setAttribute('content', uuid);
}

// --- Texto agregado a cada capítulo ---

// Inserta un fragmento XHTML al principio o al final del <body> de cada
// documento del spine, sin tocar el resto del archivo.
async function insertarEnCapitulosEpub(paquete, fragmentoXhtml, posicion) {
  for (const { ruta } of await capitulosDePaqueteEpub(paquete)) {
    const xhtml = await paquete.zip.file(ruta).async('string');
    const nuevo = posicion === 'inicio'
      ? xhtml.replace(/<body\b[^>]*>/i, (apertura) => apertura + fragmentoXhtml)
      : xhtml.replace(/<\/body\s*>(?![\s\S]*<\/body\s*>)/i, fragmentoXhtml + '</body>');
    if (nuevo !== xhtml) paquete.zip.file(ruta, nuevo);
  }
}

// --- Unir varios libros ---

// El primero es la base (su portada, estilos y metadatos); cada uno de los
// demás se copia entero en una carpeta propia junto al OPF (libro2/, libro3/…),
// así sus rutas relativas —estilos, fuentes, imágenes— siguen funcionando. El
// índice queda con una entrada por libro y sus capítulos adentro.
async function unirPaquetesEpub(paquetes, { titulo, autor }) {
  const base = paquetes[0];
  const libros = [];
  for (const paquete of paquetes) {
    await desofuscarFuentesEpub(paquete);
    libros.push({ paquete, titulo: metadatoDcEpub(paquete, 'title'), autores: [...paquete.opf.getElementsByTagNameNS(NS_DC_EPUB, 'creator')].map((e) => e.textContent.trim()).filter(Boolean) });
  }

  // Entradas del índice de cada libro, ya con rutas completas del .zip final.
  const entradasDe = async (paquete, prefijo) => {
    const { nav, ncx } = rutasIndicesEpub(paquete);
    const docNav = nav && await documentoDePaqueteEpub(paquete, nav);
    const convertir = (carpeta, href) => {
      const ruta = rutaDesdeHrefEpub(carpeta, href);
      return ruta ? { ruta: prefijo + ruta, fragmento: href.includes('#') ? href.slice(href.indexOf('#')) : '' } : null;
    };
    if (docNav) {
      const leer = (ol) => hijosDirectosEpub(ol, 'li').map((li) => {
        const a = hijosDirectosEpub(li, 'a')[0] || hijosDirectosEpub(li, 'span')[0];
        return { texto: (a?.textContent || '').trim().replace(/\s+/g, ' '), destino: a && a.getAttribute('href') ? convertir(carpetaDeRutaEpub(nav), a.getAttribute('href')) : null, hijos: leer(hijosDirectosEpub(li, 'ol')[0]) };
      });
      return leer(hijosDirectosEpub(navTocEpub(docNav), 'ol')[0]);
    }
    const docNcx = ncx && await documentoDePaqueteEpub(paquete, ncx);
    if (docNcx) {
      const leer = (padre) => hijosDirectosEpub(padre, 'navPoint').map((p) => ({
        texto: (hijosDirectosEpub(hijosDirectosEpub(p, 'navLabel')[0], 'text')[0]?.textContent || '').trim(),
        destino: convertir(carpetaDeRutaEpub(ncx), hijosDirectosEpub(p, 'content')[0]?.getAttribute('src') || ''),
        hijos: leer(p),
      }));
      return leer(docNcx.getElementsByTagNameNS('*', 'navMap')[0]);
    }
    return [];
  };

  const manifest = base.opf.getElementsByTagNameNS(NS_OPF_EPUB, 'manifest')[0];
  const spine = spineEpub(base);
  const grupos = [];
  for (let i = 0; i < libros.length; i++) {
    const { paquete } = libros[i];
    const prefijo = i === 0 ? '' : base.carpetaOpf + `libro${i + 1}/`;
    const primerCapitulo = (await capitulosDePaqueteEpub(paquete))[0];
    grupos.push({ texto: libros[i].titulo, destino: primerCapitulo ? { ruta: prefijo + primerCapitulo.ruta, fragmento: '' } : null, hijos: await entradasDe(paquete, prefijo) });
    if (i === 0) continue;

    const { nav, ncx } = rutasIndicesEpub(paquete);
    const ids = new Map();
    for (const item of itemsManifestEpub(paquete)) {
      if (!item.ruta || item.ruta === nav || item.ruta === ncx || !paquete.zip.file(item.ruta)) continue;
      const nuevoId = `l${i + 1}-${item.id}`;
      ids.set(item.id, nuevoId);
      base.zip.file(prefijo + item.ruta, await paquete.zip.file(item.ruta).async('uint8array'));
      const nuevo = base.opf.createElementNS(NS_OPF_EPUB, 'item');
      nuevo.setAttribute('id', nuevoId);
      nuevo.setAttribute('href', hrefRelativoEpub(base.carpetaOpf, prefijo + item.ruta));
      nuevo.setAttribute('media-type', item.tipo);
      // Una sola portada y un solo índice: los del libro base.
      const propiedades = item.propiedades.filter((p) => p !== 'cover-image' && p !== 'nav');
      if (propiedades.length) nuevo.setAttribute('properties', propiedades.join(' '));
      manifest.appendChild(nuevo);
    }
    for (const nuevo of [...manifest.children]) {
      const fallback = nuevo.getAttribute('fallback');
      if (fallback && nuevo.getAttribute('id')?.startsWith(`l${i + 1}-`)) nuevo.setAttribute('fallback', ids.get(fallback) || fallback);
    }
    for (const ref of hijosDirectosEpub(spineEpub(paquete), 'itemref')) {
      const id = ids.get(ref.getAttribute('idref'));
      if (!id) continue;
      const nuevo = base.opf.createElementNS(NS_OPF_EPUB, 'itemref');
      nuevo.setAttribute('idref', id);
      if (ref.getAttribute('linear')) nuevo.setAttribute('linear', ref.getAttribute('linear'));
      spine.appendChild(nuevo);
    }
  }

  // Índices de la base, reescritos con un grupo por libro.
  const { nav, ncx } = rutasIndicesEpub(base);
  const docNav = nav && await documentoDePaqueteEpub(base, nav);
  if (docNav) {
    const navToc = navTocEpub(docNav);
    const ns = navToc.namespaceURI;
    const armar = (entradas) => {
      const ol = docNav.createElementNS(ns, 'ol');
      for (const entrada of entradas) {
        const li = docNav.createElementNS(ns, 'li');
        const etiqueta = docNav.createElementNS(ns, entrada.destino ? 'a' : 'span');
        if (entrada.destino) etiqueta.setAttribute('href', hrefRelativoEpub(carpetaDeRutaEpub(nav), entrada.destino.ruta) + entrada.destino.fragmento);
        etiqueta.textContent = entrada.texto || '—';
        li.appendChild(etiqueta);
        if (entrada.hijos.length) li.appendChild(armar(entrada.hijos));
        ol.appendChild(li);
      }
      return ol;
    };
    hijosDirectosEpub(navToc, 'ol').forEach((ol) => navToc.removeChild(ol));
    navToc.appendChild(armar(grupos));
  }
  const docNcx = ncx && await documentoDePaqueteEpub(base, ncx);
  const mapa = docNcx && docNcx.getElementsByTagNameNS('*', 'navMap')[0];
  if (mapa) {
    const ns = mapa.namespaceURI;
    let contador = 0;
    const armar = (padre, entradas) => {
      for (const entrada of entradas) {
        if (!entrada.destino) { armar(padre, entrada.hijos); continue; }
        const punto = docNcx.createElementNS(ns, 'navPoint');
        punto.setAttribute('id', `np-union-${++contador}`);
        const etiqueta = docNcx.createElementNS(ns, 'navLabel');
        const texto = docNcx.createElementNS(ns, 'text');
        texto.textContent = entrada.texto || '—';
        etiqueta.appendChild(texto);
        const contenido = docNcx.createElementNS(ns, 'content');
        contenido.setAttribute('src', hrefRelativoEpub(carpetaDeRutaEpub(ncx), entrada.destino.ruta) + entrada.destino.fragmento);
        punto.appendChild(etiqueta);
        punto.appendChild(contenido);
        armar(punto, entrada.hijos);
        padre.appendChild(punto);
      }
    };
    hijosDirectosEpub(mapa, 'navPoint').forEach((p) => mapa.removeChild(p));
    armar(mapa, grupos);
    [...docNcx.getElementsByTagNameNS('*', 'navPoint')].forEach((p, i) => p.setAttribute('playOrder', String(i + 1)));
  }

  await fijarTituloEpub(base, titulo);
  const autores = autor ? [autor] : [...new Set(libros.flatMap((l) => l.autores))];
  [...base.opf.getElementsByTagNameNS(NS_DC_EPUB, 'creator')].forEach((e) => quitarMetadatoEpub(base, e));
  for (const nombre of autores) {
    const nodo = base.opf.createElementNS(NS_DC_EPUB, 'dc:creator');
    nodo.textContent = nombre;
    base.opf.getElementsByTagNameNS(NS_OPF_EPUB, 'metadata')[0].appendChild(nodo);
  }
  // La descripción del primer libro no describe la unión.
  [...base.opf.getElementsByTagNameNS(NS_DC_EPUB, 'description')].forEach((e) => quitarMetadatoEpub(base, e));
  await nuevoIdentificadorEpub(base);
  return base;
}

// --- Reparar sin rearmar ---

// Arregla lo que se puede arreglar dejando el libro como está: container.xml
// que falta o apunta mal, items del manifest sin archivo, entradas del spine
// a items que no existen, índices con enlaces rotos, y el mimetype (que
// generarBlobEpub siempre escribe bien). Devuelve el paquete y la lista de
// arreglos hechos. Si el OPF no se puede leer, lanza: ahí corresponde rearmar.
async function repararPaqueteEpub(archivo) {
  const zip = await ZipPropio.loadAsync(archivo);
  await fuentesOfuscadasEpub(zip);
  const arreglos = [];
  let rutaOpf;
  try {
    rutaOpf = await rutaOpfDeZipEpub(zip);
  } catch {
    const candidatas = Object.keys(zip.files).filter((r) => !zip.files[r].dir && /\.opf$/i.test(r));
    if (!candidatas.length) throw new Error(textoEpub('epub-sin-container', 'El EPUB no tiene un container.xml válido.'));
    rutaOpf = candidatas[0];
    zip.file('META-INF/container.xml', `<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles>
    <rootfile full-path="${escaparXmlEpub(rutaOpf)}" media-type="application/oebps-package+xml"/>
  </rootfiles>
</container>`);
    arreglos.push('container');
  }
  const opf = parsearXmlEpub(await zip.file(rutaOpf).async('string'), rutaOpf);
  const paquete = { zip, rutaOpf, carpetaOpf: carpetaDeRutaEpub(rutaOpf), opf, docs: new Map() };

  const vistos = new Set();
  for (const item of itemsManifestEpub(paquete)) {
    if (!item.id || vistos.has(item.id) || !item.ruta || !zip.file(item.ruta)) {
      item.el.parentNode.removeChild(item.el);
      arreglos.push('manifest');
    } else vistos.add(item.id);
  }
  const spine = spineEpub(paquete);
  for (const ref of hijosDirectosEpub(spine, 'itemref')) {
    if (!vistos.has(ref.getAttribute('idref'))) { spine.removeChild(ref); arreglos.push('spine'); }
  }
  if (!(await capitulosDePaqueteEpub(paquete)).length) throw new Error(textoEpub('epub-sin-contenido', 'No se encontró contenido legible dentro del EPUB.'));
  try {
    await ajustarIndicesEpub(paquete);
  } catch (errorIndice) {
    // Un índice ilegible no impide leer el libro: queda como estaba.
    paquete.docs.clear();
    arreglos.push('indice');
  }
  return { paquete, arreglos: [...new Set(arreglos)] };
}

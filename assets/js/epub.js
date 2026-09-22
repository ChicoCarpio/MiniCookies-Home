// Helpers compartidos para leer y construir archivos EPUB. Un .epub es un .zip
// con estructura fija (mimetype sin comprimir, META-INF/container.xml, un OPF
// con manifest/spine, capítulos en XHTML) — no hace falta una librería de EPUB
// dedicada, alcanza con JSZip + DOMParser (mismo criterio que el resto del
// proyecto: no sumar una dependencia si ya se puede resolver con lo que hay).
// Requiere que la página ya haya cargado JSZip antes de usar estas funciones.

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

// contenidoAParrafos-style: recibe { titulo, autor, idioma, capitulos: [{titulo, html}] }
// y devuelve un Blob del .epub listo para descargar.
async function construirEpub({ titulo, autor, idioma = 'es', capitulos }) {
  if (!capitulos || capitulos.length === 0) throw new Error('No hay contenido para armar el EPUB.');

  const zip = new JSZip();
  const uuid = generarUuidEpub();
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
<html xmlns="${ESPACIO_NOMBRES_XHTML}" lang="${idioma}">
<head><title>${escaparXmlEpub(tituloCapitulo)}</title><link rel="stylesheet" type="text/css" href="styles.css"/></head>
<body>${htmlSinImagenes || `<h1>${escaparXmlEpub(tituloCapitulo)}</h1>`}</body>
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
  <head><meta name="dtb:uid" content="${uuid}"/></head>
  <docTitle><text>${escaparXmlEpub(titulo)}</text></docTitle>
  <navMap>${navPoints.join('')}</navMap>
</ncx>`);
  manifestItems.push('<item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>');

  zip.file('OEBPS/content.opf', `<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="uid">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:identifier id="uid">${uuid}</dc:identifier>
    <dc:title>${escaparXmlEpub(titulo || 'Documento')}</dc:title>
    <dc:language>${idioma}</dc:language>
    <dc:creator>${escaparXmlEpub(autor || 'Herramientas Ofimáticas')}</dc:creator>
    <meta property="dcterms:modified">${fecha}</meta>
  </metadata>
  <manifest>${manifestItems.join('')}</manifest>
  <spine toc="ncx">${spineItems.join('')}</spine>
</package>`);

  return zip.generateAsync({ type: 'blob', mimeType: 'application/epub+zip' });
}

// Corta un HTML (de mammoth o marked) en capítulos, uno por cada <h1> que
// encuentra — si no hay ningún <h1>, todo el contenido queda como un único
// capítulo. El contenido antes del primer <h1> (si lo hay) se agrega como un
// capítulo inicial sin título propio.
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
      actual = { titulo: nodo.textContent.trim() || `Capítulo ${capitulos.length + 1}`, nodos: [] };
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
  const zip = await JSZip.loadAsync(archivo);

  if (zip.file('META-INF/encryption.xml')) {
    throw new Error(window.t ? t('epub-drm') : 'Este EPUB está protegido con DRM (cifrado) — no se puede leer su contenido en el navegador.');
  }

  const parser = new DOMParser();
  const containerXml = await zip.file('META-INF/container.xml').async('string');
  const containerDoc = parser.parseFromString(containerXml, 'application/xml');
  const rutaOpf = containerDoc.getElementsByTagName('rootfile')[0]?.getAttribute('full-path');
  if (!rutaOpf) throw new Error(window.t ? t('epub-sin-container') : 'El EPUB no tiene un container.xml válido.');
  const carpetaOpf = rutaOpf.includes('/') ? rutaOpf.slice(0, rutaOpf.lastIndexOf('/') + 1) : '';

  const opfXml = await zip.file(rutaOpf).async('string');
  const opfDoc = parser.parseFromString(opfXml, 'application/xml');

  const titulo = opfDoc.getElementsByTagName('dc:title')[0]?.textContent?.trim() || (window.t ? t('epub-sin-titulo') : 'Sin título');
  const autor = opfDoc.getElementsByTagName('dc:creator')[0]?.textContent?.trim() || '';

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
    capitulos.push({ titulo: tituloCapitulo, html: cuerpo.innerHTML });
  }

  if (capitulos.length === 0) throw new Error(window.t ? t('epub-sin-contenido') : 'No se encontró contenido legible dentro del EPUB.');

  return { titulo, autor, capitulos };
}

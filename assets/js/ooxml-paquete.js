/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Edición de paquetes OOXML (.xlsx, .xlsm, .pptx: un .zip con partes XML
// unidas por relaciones) con ZipPropio, compartida por las herramientas que
// reordenan, quitan o separan hojas y diapositivas.
//
// Quitar una hoja o diapositiva de la lista del libro/presentación no alcanza:
// su XML, sus imágenes, sus notas y (en Excel) sus textos seguían dentro del
// .zip. Quien recibía el archivo "sin la hoja 3" podía recuperarla
// descomprimiéndolo, y cada archivo de "Dividir PowerPoint" era una copia
// completa de la presentación. Además quedaban referencias colgadas
// (secciones, nombres definidos, pestaña activa) por las que Office pedía
// "reparar" el archivo.

const NS_RELACIONES_OOXML = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function rutaRelsDeParte(rutaParte) {
  const barra = rutaParte.lastIndexOf('/');
  return rutaParte.slice(0, barra + 1) + '_rels/' + rutaParte.slice(barra + 1) + '.rels';
}

// Resuelve el Target de una relación (relativo a la parte, o absoluto con "/").
function resolverRutaOoxml(rutaParte, destino) {
  let limpio = destino;
  try { limpio = decodeURI(destino); } catch (e) { /* se usa tal cual */ }
  if (limpio.startsWith('/')) return limpio.slice(1);
  const partes = [];
  (rutaParte.slice(0, rutaParte.lastIndexOf('/') + 1) + limpio).split('/').forEach((segmento) => {
    if (segmento === '..') partes.pop();
    else if (segmento !== '.' && segmento !== '') partes.push(segmento);
  });
  return partes.join('/');
}

async function leerXmlDeZip(zip, ruta) {
  const archivo = zip.file(ruta);
  if (!archivo) return null;
  const doc = new DOMParser().parseFromString(await archivo.async('string'), 'application/xml');
  if (doc.getElementsByTagName('parsererror').length > 0) throw new Error((window.t ? window.t('xml-invalido') : 'XML inválido') + ': ' + ruta);
  return doc;
}

function guardarXmlEnZip(zip, ruta, doc) {
  zip.file(ruta, new XMLSerializer().serializeToString(doc));
}

// [{ id, tipo, destino (ruta ya resuelta), externa }] de una parte.
async function relacionesDeParte(zip, rutaParte) {
  const doc = await leerXmlDeZip(zip, rutaRelsDeParte(rutaParte));
  if (!doc) return [];
  return [...doc.getElementsByTagName('Relationship')].map((rel) => ({
    id: rel.getAttribute('Id'),
    tipo: rel.getAttribute('Type') || '',
    externa: rel.getAttribute('TargetMode') === 'External',
    destino: rel.getAttribute('TargetMode') === 'External' ? rel.getAttribute('Target') : resolverRutaOoxml(rutaParte, rel.getAttribute('Target') || ''),
  }));
}

async function quitarRelaciones(zip, rutaParte, ids) {
  const rutaRels = rutaRelsDeParte(rutaParte);
  const doc = await leerXmlDeZip(zip, rutaRels);
  if (!doc) return;
  [...doc.getElementsByTagName('Relationship')].forEach((rel) => {
    if (ids.has(rel.getAttribute('Id'))) rel.parentNode.removeChild(rel);
  });
  guardarXmlEnZip(zip, rutaRels, doc);
}

// Borra del paquete toda parte a la que no se llega siguiendo las relaciones
// desde _rels/.rels (así funciona el formato: lo que no está referenciado no
// forma parte del documento), junto con su entrada en [Content_Types].xml.
// Si algún .rels alcanzable no se puede leer, no borra nada: sin el mapa
// completo podría llevarse partes que sí se usan.
async function limpiarPartesHuerfanas(zip) {
  const alcanzables = new Set();
  const pendientes = [''];
  while (pendientes.length > 0) {
    const parte = pendientes.pop();
    const rutaRels = rutaRelsDeParte(parte);
    if (!zip.file(rutaRels)) continue;
    alcanzables.add(rutaRels);
    let relaciones;
    try {
      relaciones = await relacionesDeParte(zip, parte);
    } catch (error) {
      console.warn('No se limpian partes sin uso:', error);
      return [];
    }
    relaciones.forEach(({ externa, destino }) => {
      if (externa || alcanzables.has(destino)) return;
      alcanzables.add(destino);
      pendientes.push(destino);
    });
  }

  const borradas = [];
  zip.forEach((ruta, entrada) => {
    if (entrada.dir || ruta === '[Content_Types].xml' || alcanzables.has(ruta)) return;
    borradas.push(ruta);
  });
  borradas.forEach((ruta) => zip.remove(ruta));

  const tipos = await leerXmlDeZip(zip, '[Content_Types].xml');
  if (tipos) {
    const quitadas = new Set(borradas.map((ruta) => '/' + ruta.toLowerCase()));
    [...tipos.getElementsByTagName('Override')].forEach((nodo) => {
      if (quitadas.has((nodo.getAttribute('PartName') || '').toLowerCase())) nodo.parentNode.removeChild(nodo);
    });
    guardarXmlEnZip(zip, '[Content_Types].xml', tipos);
  }

  // Carpetas que quedaron vacías.
  const carpetas = [];
  zip.forEach((ruta, entrada) => { if (entrada.dir) carpetas.push(ruta); });
  carpetas.forEach((carpeta) => {
    let tieneArchivos = false;
    zip.folder(carpeta).forEach((ruta, entrada) => { if (!entrada.dir) tieneArchivos = true; });
    if (!tieneArchivos) zip.remove(carpeta);
  });
  return borradas;
}

// docProps/thumbnail.jpeg es una imagen de la primera hoja/diapositiva del
// original: después de quitar o reordenar mostraría contenido que ya no
// está (Office la regenera al guardar). Se quita su relación y
// limpiarPartesHuerfanas borra la imagen.
async function quitarMiniaturaDocumento(zip) {
  const relaciones = await relacionesDeParte(zip, '');
  const miniaturas = relaciones.filter((rel) => /\/metadata\/thumbnail$/.test(rel.tipo));
  if (miniaturas.length > 0) await quitarRelaciones(zip, '', new Set(miniaturas.map((rel) => rel.id)));
}

function hijosPorNombreLocal(nodo, nombreLocal) {
  return [...nodo.getElementsByTagName('*')].filter((el) => el.localName === nombreLocal);
}

// docProps/app.xml guarda la lista de nombres de hojas/diapositivas
// (TitlesOfParts) agrupada según HeadingPairs. "decidir(indiceGrupo,
// textos, totalGrupos)" devuelve la nueva lista de textos de cada grupo; los
// grupos que quedan vacíos se quitan.
async function actualizarTitulosApp(zip, decidir) {
  let doc;
  try {
    doc = await leerXmlDeZip(zip, 'docProps/app.xml');
  } catch (e) {
    return null;
  }
  if (!doc) return null;
  const pares = hijosPorNombreLocal(doc, 'HeadingPairs')[0];
  const titulos = hijosPorNombreLocal(doc, 'TitlesOfParts')[0];
  if (!pares || !titulos) return doc;
  const vectorPares = hijosPorNombreLocal(pares, 'vector')[0];
  const vectorTitulos = hijosPorNombreLocal(titulos, 'vector')[0];
  if (!vectorPares || !vectorTitulos) return doc;
  const variantes = hijosPorNombreLocal(vectorPares, 'variant');
  const entradas = [...vectorTitulos.children];
  let posicion = 0;
  const nuevas = [];
  const totalGrupos = Math.floor(variantes.length / 2);
  for (let i = 0; i + 1 < variantes.length; i += 2) {
    const nodoCantidad = hijosPorNombreLocal(variantes[i + 1], 'i4')[0];
    if (!nodoCantidad) return doc;
    const cantidad = parseInt(nodoCantidad.textContent, 10) || 0;
    const grupo = entradas.slice(posicion, posicion + cantidad);
    posicion += cantidad;
    const textos = decidir(i / 2, grupo.map((nodo) => nodo.textContent), totalGrupos);
    const plantilla = grupo[0] || entradas[0];
    textos.forEach((texto) => {
      const nodo = plantilla.cloneNode(false);
      nodo.textContent = texto;
      nuevas.push(nodo);
    });
    nodoCantidad.textContent = String(textos.length);
    if (textos.length === 0) {
      vectorPares.removeChild(variantes[i]);
      vectorPares.removeChild(variantes[i + 1]);
    }
  }
  vectorPares.setAttribute('size', String(hijosPorNombreLocal(vectorPares, 'variant').length));
  entradas.forEach((nodo) => vectorTitulos.removeChild(nodo));
  nuevas.forEach((nodo) => vectorTitulos.appendChild(nodo));
  vectorTitulos.setAttribute('size', String(nuevas.length));
  guardarXmlEnZip(zip, 'docProps/app.xml', doc);
  return doc;
}

// ---------------------------------------------------------------------------
// PowerPoint: deja la presentación con las diapositivas "idsEnOrden" (los
// r:id de sus <p:sldId>), en ese orden.
// ---------------------------------------------------------------------------
async function conservarDiapositivas(zip, idsEnOrden) {
  const rutaPresentacion = 'ppt/presentation.xml';
  const presentacion = await leerXmlDeZip(zip, rutaPresentacion);
  if (!presentacion) throw new Error(rutaPresentacion);
  const idRelacion = (nodo) => nodo.getAttributeNS(NS_RELACIONES_OOXML, 'id') || nodo.getAttribute('r:id');
  const lista = hijosPorNombreLocal(presentacion, 'sldIdLst')[0];
  const nodos = [...lista.children].filter((el) => el.localName === 'sldId');
  const nodoPorRid = new Map(nodos.map((nodo) => [idRelacion(nodo), nodo]));
  const ridsOriginales = nodos.map(idRelacion);
  const conservar = idsEnOrden.filter((rid) => nodoPorRid.has(rid));
  const ridsQuitados = new Set(ridsOriginales.filter((rid) => !conservar.includes(rid)));

  nodos.forEach((nodo) => lista.removeChild(nodo));
  conservar.forEach((rid) => lista.appendChild(nodoPorRid.get(rid)));

  // Secciones (p14:sectionLst, dentro de extLst): referencian las
  // diapositivas por su id numérico y tienen que recorrerlas en el mismo
  // orden que la lista principal. Se rearman siguiendo el orden nuevo: cada
  // diapositiva vuelve a su sección, salvo que esa sección ya haya quedado
  // atrás (se movió a otra parte): entonces se suma a la sección donde cayó.
  hijosPorNombreLocal(presentacion, 'sectionLst').forEach((secciones) => {
    const listasSeccion = hijosPorNombreLocal(secciones, 'section').map((seccion) => hijosPorNombreLocal(seccion, 'sldIdLst')[0]).filter(Boolean);
    const seccionDeId = new Map();
    const nodoSeccionDeId = new Map();
    listasSeccion.forEach((listaSeccion, indice) => {
      hijosPorNombreLocal(listaSeccion, 'sldId').forEach((nodo) => {
        seccionDeId.set(nodo.getAttribute('id'), indice);
        nodoSeccionDeId.set(nodo.getAttribute('id'), nodo);
        listaSeccion.removeChild(nodo);
      });
    });
    let actual = 0;
    conservar.forEach((rid) => {
      const idNumerico = nodoPorRid.get(rid).getAttribute('id');
      const nodo = nodoSeccionDeId.get(idNumerico);
      if (!nodo || listasSeccion.length === 0) return;
      actual = Math.max(actual, seccionDeId.get(idNumerico));
      listasSeccion[actual].appendChild(nodo);
    });
  });

  // Presentaciones personalizadas: referencian por r:id.
  hijosPorNombreLocal(presentacion, 'custShow').forEach((show) => {
    hijosPorNombreLocal(show, 'sld').forEach((nodo) => {
      if (ridsQuitados.has(idRelacion(nodo))) nodo.parentNode.removeChild(nodo);
    });
    if (hijosPorNombreLocal(show, 'sld').length === 0) show.parentNode.removeChild(show);
  });
  hijosPorNombreLocal(presentacion, 'custShowLst').forEach((listaShows) => {
    if (listaShows.children.length === 0) listaShows.parentNode.removeChild(listaShows);
  });
  guardarXmlEnZip(zip, rutaPresentacion, presentacion);

  const relaciones = await relacionesDeParte(zip, rutaPresentacion);
  const rutaPorRid = new Map(relaciones.map((rel) => [rel.id, rel.destino]));
  const rutasQuitadas = new Set([...ridsQuitados].map((rid) => rutaPorRid.get(rid)));
  await quitarRelaciones(zip, rutaPresentacion, ridsQuitados);

  // Hipervínculos "ir a la diapositiva X" hacia una diapositiva quitada:
  // mantendrían su XML dentro del paquete y apuntarían a algo que ya no
  // está en la presentación. Se quita el vínculo (el texto queda).
  let notas = 0;
  let ocultas = 0;
  for (const rid of conservar) {
    const rutaDiapositiva = rutaPorRid.get(rid);
    const relacionesDiapositiva = await relacionesDeParte(zip, rutaDiapositiva);
    if (relacionesDiapositiva.some((rel) => /\/notesSlide$/.test(rel.tipo))) notas++;
    const archivo = zip.file(rutaDiapositiva);
    if (archivo && /<p:sld\b[^>]*\sshow="(0|false)"/.test(await archivo.async('string'))) ocultas++;
    const idsRotos = new Set(relacionesDiapositiva.filter((rel) => !rel.externa && rutasQuitadas.has(rel.destino)).map((rel) => rel.id));
    if (idsRotos.size === 0) continue;
    const doc = await leerXmlDeZip(zip, rutaDiapositiva);
    [...doc.getElementsByTagName('*')].forEach((el) => {
      if (/^hlink/.test(el.localName) && idsRotos.has(idRelacion(el)) && el.parentNode) el.parentNode.removeChild(el);
    });
    guardarXmlEnZip(zip, rutaDiapositiva, doc);
    await quitarRelaciones(zip, rutaDiapositiva, idsRotos);
  }

  // docProps/app.xml: cantidad de diapositivas y sus títulos (el último
  // grupo de TitlesOfParts, uno por diapositiva en el orden original).
  const app = await actualizarTitulosApp(zip, (indice, textos, totalGrupos) => {
    if (indice !== totalGrupos - 1 || textos.length !== ridsOriginales.length) return textos;
    return conservar.map((rid) => textos[ridsOriginales.indexOf(rid)]);
  });
  if (app) {
    const fijar = (nombre, valor) => {
      const nodo = hijosPorNombreLocal(app, nombre)[0];
      if (nodo) nodo.textContent = String(valor);
    };
    fijar('Slides', conservar.length);
    fijar('Notes', notas);
    fijar('HiddenSlides', ocultas);
    guardarXmlEnZip(zip, 'docProps/app.xml', app);
  }

  await quitarMiniaturaDocumento(zip);
  await limpiarPartesHuerfanas(zip);
}

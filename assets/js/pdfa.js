/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Helper para agregar/quitar los marcadores de conformidad PDF/A de un PDF ya
// cargado con pdf-lib. Requiere que la página ya haya cargado pdf-lib.
//
// IMPORTANTE — esto es "mejor esfuerzo", no una conversión PDF/A certificada:
// agrega lo que un PDF/A real necesita (perfil de color embebido vía
// OutputIntent + metadatos XMP de identificación pdfaid) pero NO puede
// garantizar el resto de los requisitos del estándar ISO 19005 sobre un PDF
// que ya viene armado de otro lado — en particular, no puede forzar que TODAS
// las fuentes ya usadas en el PDF estén embebidas, ni quitar transparencias o
// JavaScript si el archivo original los trae. Se avisa esto explícitamente en
// la propia herramienta (ver pdf-a-pdfa.html) en vez de prometer conformidad
// real sin poder validarla (no hay un validador tipo veraPDF disponible para
// correr 100% en el navegador).

// Perfil sRGB de "Compact ICC Profiles" (CC0, ver creditos.html). Con
// rutaVendor, no con una ruta relativa: '../assets/icc/…' desde
// /en/herramientas/ apuntaba a /en/assets/, que no existe, y la versión en
// inglés no podía convertir (encontrado el 2026-10-02).
async function cargarPerfilIccSrgb() {
  const respuesta = await fetch(rutaVendor('compact-icc-profiles/8e22c54/sRGB-v2-nano.icc'));
  if (!respuesta.ok) throw new Error(t('error-perfil-icc'));
  return new Uint8Array(await respuesta.arrayBuffer());
}

function construirXmpPdfA(pdfDoc, { parte, nivel }) {
  const titulo = pdfDoc.getTitle() || '';
  const ahora = new Date().toISOString();
  return `<?xpacket begin="﻿" id="W5M0MpCehiHzreSzNTczkc9d"?>
<x:xmpmeta xmlns:x="adobe:ns:meta/">
 <rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
  <rdf:Description rdf:about=""
    xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:xmp="http://ns.adobe.com/xap/1.0/">
    <pdfaid:part>${parte}</pdfaid:part>
    <pdfaid:conformance>${nivel}</pdfaid:conformance>
    <dc:title><rdf:Alt><rdf:li xml:lang="x-default">${escaparXmlPdfa(titulo)}</rdf:li></rdf:Alt></dc:title>
    <xmp:ModifyDate>${ahora}</xmp:ModifyDate>
  </rdf:Description>
 </rdf:RDF>
</x:xmpmeta>
<?xpacket end="w"?>`;
}

function escaparXmlPdfa(texto) {
  return String(texto ?? '').replace(/[&<>"']/g, (c) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;',
  }[c]));
}

// parte: '1' | '2' | '3'  —  nivel: siempre 'B' (Basic). No se ofrece nivel
// "A" (Accesible) porque exige contenido etiquetado (tagged PDF) que esta
// herramienta no puede garantizar sobre un PDF que ya viene armado de otro lado.
async function agregarConformidadPdfA(pdfDoc, { parte = '2', nivel = 'B' } = {}) {
  const { PDFName, PDFString } = PDFLib;
  const contexto = pdfDoc.context;

  const bytesIcc = await cargarPerfilIccSrgb();
  const streamIcc = contexto.flateStream(bytesIcc, { N: 3, Alternate: PDFName.of('DeviceRGB') });
  const refIcc = contexto.register(streamIcc);

  const outputIntent = contexto.obj({
    Type: 'OutputIntent',
    S: 'GTS_PDFA1',
    OutputConditionIdentifier: PDFString.of('sRGB IEC61966-2.1'),
    Info: PDFString.of('sRGB IEC61966-2.1 (perfil compacto, ver aviso de la herramienta)'),
    DestOutputProfile: refIcc,
  });
  const refOutputIntent = contexto.register(outputIntent);
  pdfDoc.catalog.set(PDFName.of('OutputIntents'), contexto.obj([refOutputIntent]));

  const xmp = construirXmpPdfA(pdfDoc, { parte, nivel });
  const streamMetadata = contexto.stream(new TextEncoder().encode(xmp), {
    Type: 'Metadata',
    Subtype: 'XML',
  });
  const refMetadata = contexto.register(streamMetadata);
  pdfDoc.catalog.set(PDFName.of('Metadata'), refMetadata);
}

// Quita los marcadores de PDF/A (queda como un PDF normal — un PDF/A ya es un
// PDF válido de por sí, así que esto no "arregla" nada, solo saca el
// OutputIntent y los metadatos XMP de identificación pdfaid).
function quitarConformidadPdfA(pdfDoc) {
  const { PDFName } = PDFLib;
  pdfDoc.catalog.delete(PDFName.of('OutputIntents'));
  pdfDoc.catalog.delete(PDFName.of('Metadata'));
}

// Detecta si un PDF ya cargado con pdf-lib declara conformidad PDF/A, leyendo
// el XMP de forma simple (sin parsear XML completo, solo buscando la etiqueta).
async function detectarConformidadPdfA(pdfDoc) {
  const { PDFName } = PDFLib;
  const refMetadata = pdfDoc.catalog.get(PDFName.of('Metadata'));
  if (!refMetadata) return null;
  const streamMetadata = pdfDoc.context.lookup(refMetadata);
  const bytes = streamMetadata?.contents || streamMetadata?.getContents?.();
  if (!bytes) return null;
  const texto = new TextDecoder().decode(bytes);
  const parte = /<pdfaid:part>(\d)<\/pdfaid:part>/.exec(texto)?.[1];
  const nivel = /<pdfaid:conformance>(\w)<\/pdfaid:conformance>/.exec(texto)?.[1];
  return parte ? { parte, nivel } : null;
}

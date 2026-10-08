/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. */
// Pruebas de funcionamiento: abre el sitio compilado en Edge, SIN internet, y
// usa de verdad las herramientas — no solo mira cómo se ven (eso es
// revisar-interfaz.mjs).
//
// Por qué existe: los errores que más costaron el 2026-10-01 no se veían. El
// build salteaba las carpetas "dist" de las librerías y fontkit no se
// publicaba, pero Excel → PDF "andaba" igual porque caía en silencio a su
// respaldo; dictado y stickers en inglés estaban rotos en producción; el
// buscador en inglés buscaba en español. Todo eso se encontró con pruebas
// armadas a mano que se perdían al terminar la sesión. Esto las deja fijas.
//
// Reglas de estas pruebas (aprendidas a los golpes):
//  - Cada una verifica una SEÑAL POSITIVA (el texto que reconoció el OCR, que
//    no hubo caracteres sustituidos, cuántas miniaturas hay), no solo que no
//    hubo error: varias herramientas tienen respaldos silenciosos.
//  - Tiempo real, sin --virtual-time-budget: con ese flag performance.now()
//    miente y una espera de minutos parece de un segundo.
//  - Sin internet (--host-resolver-rules): si algo depende de una CDN, falla.
//    El servidor registra todo 404: un archivo que falta es una falla aunque
//    la página "parezca" andar.
//
// Uso: npm run probar     (compila y prueba; código 1 si algo falla)
// Igual que revisar: NO conectar a "npm run build" (Cloudflare no tiene Edge).

import { createServer } from 'node:http';
import { spawn } from 'node:child_process';
import { readFileSync, writeFileSync, rmSync, existsSync, readdirSync, statSync, mkdtempSync } from 'node:fs';
import { join, extname, relative, sep } from 'node:path';
import { tmpdir } from 'node:os';

const RAIZ = new URL('.', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1');
const DIST = join(RAIZ, 'dist');
const PUERTO = 8789;
const LIMITE_MS = 10 * 60 * 1000;
const EDGE = [process.env.EDGE,
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe'].filter(Boolean).find((r) => existsSync(r));

const TIPOS = {
  '.html': 'text/html; charset=utf-8', '.css': 'text/css; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8', '.json': 'application/json', '.png': 'image/png', '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml', '.ico': 'image/x-icon', '.wasm': 'application/wasm', '.ttf': 'font/ttf',
  '.woff2': 'font/woff2', '.gz': 'application/gzip', '.txt': 'text/plain; charset=utf-8', '.xml': 'application/xml',
};

function paginas(dir, acc = []) {
  for (const n of readdirSync(dir)) {
    const r = join(dir, n);
    if (statSync(r).isDirectory()) { if (n !== 'vendor') paginas(r, acc); } else if (n.endsWith('.html') && !n.startsWith('__')) {
      acc.push('/' + relative(DIST, r).split(sep).join('/'));
    }
  }
  return acc;
}

// --- Lo que corre en el navegador ------------------------------------------
const PRUEBAS = String.raw`
const marco = document.getElementById('marco');
const esperar = (ms) => new Promise((r) => setTimeout(r, ms));
const resultados = [];
const informar = (fin) => fetch('/__resultado', { method: 'POST', body: JSON.stringify({ resultados, fin }) });

function abrir(url) {
  return new Promise((r) => {
    let listo = false; const seguir = () => { if (!listo) { listo = true; r(); } };
    marco.onload = seguir; marco.src = url; setTimeout(seguir, 8000);
  });
}
async function hasta(cond, ms = 30000) {
  const t0 = Date.now();
  while (Date.now() - t0 < ms) { try { const v = cond(); if (v) return v; } catch (e) {} await esperar(150); }
  return null;
}
async function prueba(nombre, fn) {
  if (SOLO && !nombre.includes(SOLO)) return;
  const t0 = Date.now();
  try {
    const detalle = await fn();
    resultados.push({ nombre, ok: true, detalle, s: Math.round((Date.now() - t0) / 1000) });
  } catch (e) {
    resultados.push({ nombre, ok: false, detalle: e.message, s: Math.round((Date.now() - t0) / 1000) });
  }
  await informar(false);
}
const afirmar = (cond, msj) => { if (!cond) throw new Error(msj); };
const W = () => marco.contentWindow, D = () => marco.contentDocument;

// Espía los clics sobre <input type=file> para no abrir el diálogo real.
function espiarSelector() {
  const clics = [];
  const proto = W().HTMLInputElement.prototype;
  proto.click = function () { clics.push(this); };
  return clics;
}
function archivos(lista) {
  const dt = new (W().DataTransfer)();
  for (const f of lista) dt.items.add(f);
  return dt;
}
// La zona por piezas del Editor avanzado (assets/js/mascaras-capa.js).
function agregarPieza(clave) {
  D().getElementById('botonAgregarPieza').click();
  D().querySelector('#menuPiezas [data-pieza="' + clave + '"]').click();
  return W().piezaActiva();
}
function vaciarZona() {
  for (const b of [...D().querySelectorAll('#listaPiezas .fila-pieza .icono-capa:last-child')]) b.click();
}
function zonaSoloToda() { vaciarZona(); agregarPieza('toda'); }
function invertirPieza(valor) {
  const c = D().getElementById('invertirPieza');
  c.checked = valor; c.dispatchEvent(new (W().Event)('change'));
}
async function pdfDePrueba(paginas = 1) {
  const doc = await W().PDFLib.PDFDocument.create();
  for (let i = 0; i < paginas; i++) doc.addPage();
  return new (W().File)([await doc.save()], 'prueba.pdf', { type: 'application/pdf' });
}
async function imagenDePrueba(texto, ancho = 600, alto = 200) {
  const c = D().createElement('canvas'); c.width = ancho; c.height = alto;
  const cx = c.getContext('2d'); cx.fillStyle = '#fff'; cx.fillRect(0, 0, ancho, alto);
  cx.fillStyle = '#000'; cx.font = 'bold 64px Arial'; cx.fillText(texto, 30, alto / 2 + 22);
  const b = await new Promise((r) => c.toBlob(r, 'image/png'));
  return new (W().File)([b], 'prueba.png', { type: 'image/png' });
}

(async () => {
  // 1. Todas las páginas cargan sin archivos faltantes (el servidor anota los
  //    404) y toda zona de arrastre se puede usar con teclado.
  await prueba('todas las páginas: zonas de arrastre accesibles con teclado', async () => {
    let zonas = 0; const fallas = [];
    for (const [i, p] of PAGINAS.entries()) {
      const t0 = Date.now();
      await abrir(p); await esperar(120);
      fetch('/__avance', { method: 'POST', body: JSON.stringify({ i, total: PAGINAS.length, p, ms: Date.now() - t0 }) });
      const lista = [...D().querySelectorAll('.zona-arrastre')];
      for (const z of lista) {
        zonas++;
        const clics = espiarSelector();
        z.dispatchEvent(new (W().KeyboardEvent)('keydown', { key: 'Enter', bubbles: true }));
        if (z.getAttribute('role') !== 'button' || z.tabIndex !== 0) fallas.push(p + ' #' + z.id + ': sin role/tabindex');
        else if (!clics.some((i) => i.type === 'file')) fallas.push(p + ' #' + z.id + ': Enter no abre el selector');
      }
    }
    afirmar(!fallas.length, fallas.slice(0, 5).join(' | '));
    afirmar(zonas > 80, 'solo ' + zonas + ' zonas: ¿cambió el marcado?');
    return zonas + ' zonas en ' + PAGINAS.length + ' páginas';
  });

  // 2. Volver a elegir el mismo archivo funciona, y soltar también.
  await prueba('zona: mismo archivo dos veces y soltar (unir-pdf)', async () => {
    await abrir('/herramientas/unir-pdf.html'); await esperar(300);
    const entrada = D().getElementById('entradaArchivos'), zona = D().getElementById('zonaArrastre');
    const grid = D().getElementById('gridArchivos');
    const pdf = await pdfDePrueba();
    for (let i = 0; i < 2; i++) {
      entrada.files = archivos([pdf]).files;
      entrada.dispatchEvent(new (W().Event)('change'));
      afirmar(entrada.value === '', 'el selector no se limpió después de elegir');
      afirmar(await hasta(() => grid.children.length === i + 1, 15000), 'la elección ' + (i + 1) + ' del mismo archivo no se agregó');
    }
    zona.dispatchEvent(new (W().DragEvent)('drop', { dataTransfer: archivos([pdf]), bubbles: true, cancelable: true }));
    afirmar(await hasta(() => grid.children.length === 3, 15000), 'soltar un archivo no lo agregó');
    return '3 miniaturas (elegir ×2 + soltar)';
  });

  await prueba('zona: miniatura no reabre el selector (comparar-imagenes)', async () => {
    await abrir('/herramientas/comparar-imagenes.html'); await esperar(300);
    const entrada = D().getElementById('entradaArchivoA');
    entrada.files = archivos([await imagenDePrueba('A')]).files;
    entrada.dispatchEvent(new (W().Event)('change'));
    const mini = await hasta(() => { const m = D().getElementById('miniaturaA'); return m.style.display === 'block' && m; });
    afirmar(mini, 'la imagen no se cargó');
    const clics = espiarSelector();
    mini.click();
    afirmar(!clics.length, 'clic en la miniatura abrió el selector');
    D().getElementById('zonaArrastreA').click();
    afirmar(clics.length === 1, 'clic en la zona no abrió el selector');
    return 'ok';
  });

  await prueba('zona: filtra por tipo (comparar-word)', async () => {
    await abrir('/herramientas/comparar-word.html'); await esperar(300);
    const entrada = D().getElementById('entradaA'), nombre = D().getElementById('nombreA');
    entrada.files = archivos([new (W().File)(['x'], 'nota.txt', { type: 'text/plain' })]).files;
    entrada.dispatchEvent(new (W().Event)('change'));
    await esperar(200);
    afirmar(!nombre.textContent, 'aceptó un .txt');
    entrada.files = archivos([new (W().File)(['x'], 'informe.docx')]).files;
    entrada.dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => nombre.textContent === 'informe.docx', 3000), 'no tomó el .docx');
    return 'rechaza .txt, acepta .docx';
  });

  await prueba('selector de dirección: un panel por opción (herramientas bidireccionales)', async () => {
    let paginasConSelector = 0;
    for (const p of PAGINAS.filter((x) => x.startsWith('/herramientas/'))) {
      await abrir(p); await esperar(80);
      const paneles = [...D().querySelectorAll('[data-panel-direccion]')];
      if (!paneles.length) continue;
      paginasConSelector++;
      for (const radio of D().querySelectorAll('.selector-direccion input[type="radio"]')) {
        radio.checked = true;
        radio.dispatchEvent(new (W().Event)('change', { bubbles: true }));
        const visibles = paneles.filter((x) => W().getComputedStyle(x).display !== 'none');
        afirmar(visibles.length === 1 && visibles[0].dataset.panelDireccion === radio.value,
          p + ': con "' + radio.value + '" se ven ' + visibles.map((x) => x.dataset.panelDireccion).join(', '));
      }
    }
    afirmar(paginasConSelector >= 10, 'solo ' + paginasConSelector + ' herramientas con selector');
    return paginasConSelector + ' herramientas, todas las direcciones';
  });

  // Pantalla completa en celular (arquetipo C): con contenido cargado a 390px
  // tiene que aparecer la barra de categorías. El 2026-10-04 un moveBefore
  // sobre la tira de publicidad todavía suelta cortaba app-movil.js a la
  // mitad: la página quedaba en modo celular pero sin categorías ni botones,
  // y revisar-interfaz no lo veía porque mira las páginas sin imagen.
  await prueba('pantalla completa en celular: barra de categorías con contenido cargado', async () => {
    const detalle = [];
    marco.style.width = '390px'; marco.style.height = '844px';
    try {
      for (const pagina of ['editor-avanzado-imagen', 'Editor-de-color-de-imagen', 'collage-imagenes']) {
        await abrir('/herramientas/' + pagina + '.html'); await esperar(300);
        const errores = [];
        W().addEventListener('error', (ev) => errores.push(ev.message));
        const entrada = D().getElementById('entradaArchivo');
        entrada.files = archivos([await imagenDePrueba('A', 400, 300), await imagenDePrueba('B', 400, 300)]).files;
        entrada.dispatchEvent(new (W().Event)('change'));
        const chips = await hasta(() => D().querySelectorAll('.app-chip').length, 15000);
        afirmar(!errores.length, pagina + ': ' + errores[0]);
        afirmar(chips, pagina + ': sin barra de categorías (body: "' + D().body.className + '")');
        afirmar(D().querySelector('.app-anuncio'), pagina + ': sin tira de publicidad');
        detalle.push(pagina + ' ' + chips);
      }
    } finally {
      marco.style.width = ''; marco.style.height = '';
    }
    return detalle.join(', ');
  });

  await prueba('.columnas: dos columnas desde estilos.css', async () => {
    await abrir('/herramientas/comparador-texto.html'); await esperar(150);
    const c = D().querySelector('.columnas');
    const e = W().getComputedStyle(c);
    afirmar(e.display === 'grid' && e.gridTemplateColumns.split(' ').length === 2, 'display ' + e.display + ', columnas ' + e.gridTemplateColumns);
    return e.gridTemplateColumns;
  });

  // 3. Librerías de vendor/ cargadas de verdad (sin internet).
  for (const ruta of ['/herramientas/excel-csv-a-pdf.html', '/en/herramientas/excel-csv-a-pdf.html']) {
    await prueba('Excel → PDF con fuente Noto (' + (ruta.startsWith('/en/') ? 'en' : 'es') + ')', async () => {
      await abrir(ruta); await esperar(300);
      const X = W().XLSX;
      const libro = X.utils.book_new();
      X.utils.book_append_sheet(libro, X.utils.aoa_to_sheet([['Producto', 'Precio'], ['Café', '₡1 500'], ['Té', '₡900']]), 'Hoja');
      const bytes = X.write(libro, { type: 'array', bookType: 'xlsx' });
      await W().mostrarArchivo(new (W().File)([bytes], 'prueba.xlsx'));
      await hasta(() => D().getElementById('visorPdf').getAttribute('src'));
      const boton = D().getElementById('botonConvertir');
      boton.click(); await esperar(300);
      afirmar(await hasta(() => !boton.disabled, 60000), 'la conversión no terminó');
      const pdf = await (await fetch(D().getElementById('enlaceDescarga').href)).arrayBuffer();
      const aviso = D().getElementById('advertencia').textContent;
      // Si Noto no carga, el ₡ se sustituye y aparece este aviso: es la señal.
      afirmar(!/sustitu|replaced/i.test(aviso), 'caracteres sustituidos (¿no cargó Noto/fontkit?): ' + aviso);
      afirmar(pdf.byteLength > 2000, 'PDF demasiado chico');
      return 'PDF de ' + pdf.byteLength + ' bytes, sin sustituciones';
    });
  }

  // Negrita, cursiva, subrayado y color de letra: SheetJS gratuito no los
  // escribe ni los entrega por celda, así que el Excel de prueba se arma
  // reemplazando styles.xml y marcando cada celda con su formato (s="N").
  function armarExcelConEstilos() {
    return (async () => {
      const l = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(l, XLSX.utils.aoa_to_sheet([['A', 'B', 'C', 'D', 'E'], ['Normal', 'Negrita', 'Cursiva', 'Subrayado', 'Blanco']]), 'Hoja1');
      const zip = await JSZip.loadAsync(XLSX.write(l, { type: 'array', bookType: 'xlsx' }));
      const fuente = (extra) => '<font>' + extra + '<sz val="11"/><name val="Calibri"/></font>';
      zip.file('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + '<fonts count="5">' + fuente('<color rgb="FF000000"/>') + fuente('<b/>') + fuente('<i/>') + fuente('<u/>') + fuente('<color rgb="FFFFFFFF"/>') + '</fonts>'
        + '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill></fills>'
        + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        + '<cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="4" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>'
        + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>');
      let hoja = await zip.file('xl/worksheets/sheet1.xml').async('text');
      // Segunda fila: la primera es el encabezado, que siempre va en negrita.
      ['B', 'C', 'D', 'E'].forEach((col, i) => { hoja = hoja.replace(new RegExp('<c r="' + col + '2"( s="\\d+")?'), '<c r="' + col + '2" s="' + (i + 1) + '"'); });
      zip.file('xl/worksheets/sheet1.xml', hoja);
      return zip.generateAsync({ type: 'uint8array' });
    })();
  }

  await prueba('Excel → PDF: negrita, cursiva, subrayado y color de letra', async () => {
    await abrir('/herramientas/excel-csv-a-pdf.html'); await esperar(300);
    const bytes = new Uint8Array(await W().eval('(' + armarExcelConEstilos + ')')());
    await W().mostrarArchivo(new (W().File)([bytes], 'estilos.xlsx'));
    const leidas = W().eval('[0,1,2,3,4].map(c => leerCeldaConEstilo(libroActual.Sheets.Hoja1, 1, c))');
    const marca = (c) => (c.negrita ? 'N' : '') + (c.cursiva ? 'C' : '') + (c.subrayado ? 'S' : '') + (c.color ? '#' + [c.color.r, c.color.g, c.color.b].map((v) => Math.round(v * 255)).join(',') : '');
    const obtenido = leidas.map(marca).join(' | ');
    afirmar(obtenido === ' | N | C | S | #255,255,255', 'leído: "' + obtenido + '"');
    // Y que el PDF use de verdad la fuente de cada estilo.
    D().getElementById('botonConvertir').click();
    const enlace = await hasta(() => D().getElementById('resultado').classList.contains('visible') && D().getElementById('enlaceDescarga').href, 60000);
    afirmar(enlace, 'no generó el PDF');
    const pdf = new Uint8Array(await (await W().fetch(enlace)).arrayBuffer());
    await abrir('/herramientas/rotar-pdf.html'); await esperar(300);
    const doc = await W().abrirPdf(new (W().Uint8Array)(pdf));
    const pagina = await doc.getPage(1);
    await pagina.getOperatorList();
    const items = (await pagina.getTextContent()).items.filter((i) => i.str.trim());
    const fuenteDe = (texto) => { const it = items.find((i) => i.str === texto); return it ? pagina.commonObjs.get(it.fontName).name : 'no está'; };
    const esperado = { Normal: /^Helvetica$/, Negrita: /Helvetica-Bold$/, Cursiva: /Helvetica-Oblique$/, Subrayado: /^Helvetica$/ };
    const fuentes = Object.keys(esperado).map((t) => t + '=' + fuenteDe(t));
    afirmar(Object.entries(esperado).every(([t, re]) => re.test(fuenteDe(t))), 'fuentes en el PDF: ' + fuentes.join(', '));
    return obtenido + ' · ' + fuentes.join(', ');
  });

  // EPUB: las herramientas que modifican un libro lo editan en su lugar
  // (assets/js/epub.js) y conservan estilos, fuentes, portada y metadatos.
  // Antes lo rearmaban con un estilo genérico. El libro de prueba trae todo
  // eso; cada herramienta se usa desde su interfaz.
  function armarEpubDePrueba(fuenteOfuscada) {
    return (async () => {
      const uuid = 'urn:uuid:1b4e28ba-2fa1-41d2-883f-0016d3cca427';
      const zip = new JSZip();
      zip.file('mimetype', 'application/epub+zip', { compression: 'STORE' });
      zip.file('META-INF/container.xml', '<?xml version="1.0"?><container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles></container>');
      const fuente = new Uint8Array(await (await fetch('/vendor/google-fonts/d2761cb6ca3b.woff2')).arrayBuffer());
      if (fuenteOfuscada) {
        const clave = new Uint8Array(await crypto.subtle.digest('SHA-1', new TextEncoder().encode(uuid)));
        for (let i = 0; i < 1040 && i < fuente.length; i++) fuente[i] ^= clave[i % 20];
        zip.file('META-INF/encryption.xml', '<?xml version="1.0"?><encryption xmlns="urn:oasis:names:tc:opendocument:xmlns:container" xmlns:enc="http://www.w3.org/2001/04/xmlenc#"><enc:EncryptedData><enc:EncryptionMethod Algorithm="http://www.idpf.org/2008/embedding"/><enc:CipherData><enc:CipherReference URI="OEBPS/Fonts/fuente.woff2"/></enc:CipherData></enc:EncryptedData></encryption>');
      }
      zip.file('OEBPS/Fonts/fuente.woff2', fuente);
      zip.file('OEBPS/Styles/estilo.css', '@font-face { font-family: "Prueba"; src: url(../Fonts/fuente.woff2); } .destacado { color: #c0392b; }');
      const c = document.createElement('canvas'); c.width = c.height = 20;
      const png = new Uint8Array(await (await new Promise((r) => c.toBlob(r, 'image/png'))).arrayBuffer());
      const pagina = (cuerpo) => '<?xml version="1.0" encoding="UTF-8"?><html xmlns="http://www.w3.org/1999/xhtml"><head><title>x</title><link rel="stylesheet" type="text/css" href="../Styles/estilo.css"/></head><body>' + cuerpo + '</body></html>';
      zip.file('OEBPS/Images/portada.png', png);
      let items = ''; let refs = ''; let nav = ''; let ncx = '';
      for (let i = 1; i <= 3; i++) {
        zip.file('OEBPS/Images/img' + i + '.png', png);
        zip.file('OEBPS/Text/cap' + i + '.xhtml', pagina('<h1 class="destacado">Capítulo ' + i + '</h1><p><img src="../Images/img' + i + '.png" alt=""/></p>'));
        items += '<item id="cap' + i + '" href="Text/cap' + i + '.xhtml" media-type="application/xhtml+xml"/><item id="img' + i + '" href="Images/img' + i + '.png" media-type="image/png"/>';
        refs += '<itemref idref="cap' + i + '"/>';
        nav += '<li><a href="cap' + i + '.xhtml">Capítulo ' + i + '</a></li>';
        ncx += '<navPoint id="np' + i + '" playOrder="' + i + '"><navLabel><text>Capítulo ' + i + '</text></navLabel><content src="Text/cap' + i + '.xhtml"/></navPoint>';
      }
      zip.file('OEBPS/Text/nav.xhtml', '<?xml version="1.0" encoding="UTF-8"?><html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops"><head><title>Índice</title></head><body><nav epub:type="toc"><ol>' + nav + '</ol></nav></body></html>');
      zip.file('OEBPS/toc.ncx', '<?xml version="1.0"?><ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1"><head><meta name="dtb:uid" content="' + uuid + '"/></head><docTitle><text>Libro</text></docTitle><navMap>' + ncx + '</navMap></ncx>');
      zip.file('OEBPS/content.opf', '<?xml version="1.0"?><package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="bookid"><metadata xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:identifier id="bookid">' + uuid + '</dc:identifier><dc:title>Libro</dc:title><dc:creator>Autora</dc:creator><dc:language>es</dc:language><dc:publisher>Editorial Prueba</dc:publisher><meta property="dcterms:modified">2024-05-01T00:00:00Z</meta></metadata><manifest><item id="nav" href="Text/nav.xhtml" media-type="application/xhtml+xml" properties="nav"/><item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/><item id="css" href="Styles/estilo.css" media-type="text/css"/><item id="fuente" href="Fonts/fuente.woff2" media-type="font/woff2"/><item id="portada" href="Images/portada.png" media-type="image/png" properties="cover-image"/>' + items + '</manifest><spine toc="ncx">' + refs + '</spine></package>');
      return zip.generateAsync({ type: 'uint8array' });
    })();
  }
  // Lo que tiene que sobrevivir, y el orden de lectura resultante.
  async function revisarEpub(bytes) {
    const zip = await W().JSZip.loadAsync(new (W().Uint8Array)(bytes));
    const opf = await zip.file('OEBPS/content.opf').async('string');
    const faltan = ['OEBPS/Styles/estilo.css', 'OEBPS/Fonts/fuente.woff2', 'OEBPS/Images/portada.png'].filter((r) => !zip.file(r));
    if (!/Editorial Prueba/.test(opf)) faltan.push('editorial');
    if (!/properties="cover-image"/.test(opf)) faltan.push('portada marcada');
    if (Object.keys(zip.files)[0] !== 'mimetype') faltan.push('mimetype primero');
    const spine = [...opf.matchAll(/<itemref idref="([^"]+)"/g)].map((m) => m[1]).join(',');
    return { faltan, spine, zip };
  }
  async function epubDePrueba(ofuscada) {
    await abrir('/herramientas/administrar-epub.html'); await esperar(300);
    return new Uint8Array(await W().eval('(' + armarEpubDePrueba + ')')(!!ofuscada));
  }
  const descargarResultado = async (ms = 30000) => {
    const enlace = await hasta(() => D().getElementById('resultado').classList.contains('visible') && D().getElementById('enlaceDescarga').href, ms);
    afirmar(enlace, 'no terminó');
    return new Uint8Array(await (await W().fetch(enlace)).arrayBuffer());
  };

  await prueba('EPUB: administrar conserva diseño y reordena', async () => {
    const libro = await epubDePrueba();
    const alertas = []; W().alert = (m) => alertas.push(m);
    await W().cargarLibro(new (W().File)([libro], 'libro.epub'));
    const lista = D().getElementById('listaCapitulos');
    lista.querySelectorAll('li')[1].querySelector('.quitar').click(); // fuera el capítulo 2
    lista.insertBefore(lista.querySelectorAll('li')[1], lista.querySelectorAll('li')[0]); // el 3 primero
    lista.dispatchEvent(new (W().Event)('drop'));
    D().getElementById('botonGuardar').click();
    const r = await revisarEpub(await descargarResultado());
    afirmar(!alertas.length, 'error: ' + alertas.join(' | '));
    afirmar(!r.faltan.length, 'se perdió: ' + r.faltan.join(', '));
    afirmar(r.spine === 'cap3,cap1', 'spine: ' + r.spine);
    afirmar(!r.zip.file('OEBPS/Images/img2.png'), 'quedó la imagen del capítulo quitado');
    const nav = await r.zip.file('OEBPS/Text/nav.xhtml').async('string');
    afirmar(/cap3\.xhtml[\s\S]*cap1\.xhtml/.test(nav) && !/cap2\.xhtml/.test(nav), 'índice: ' + nav.replace(/^[\s\S]*<ol>/, '<ol>').slice(0, 200));
    return 'spine ' + r.spine + ', estilo, fuente, portada y editorial intactos';
  });

  await prueba('EPUB: agregar texto, dividir, unir y reparar conservan el diseño', async () => {
    const libro = await epubDePrueba();
    const detalles = [];
    // Agregar texto a cada capítulo.
    await abrir('/herramientas/marca-agua-epub.html'); await esperar(300);
    let alertas = []; W().alert = (m) => alertas.push(m);
    W().mostrarArchivo(new (W().File)([libro], 'libro.epub'));
    D().getElementById('textoAgregar').value = 'Copia de prueba & más';
    D().getElementById('botonAplicar').click();
    let r = await revisarEpub(await descargarResultado());
    afirmar(!alertas.length && !r.faltan.length, 'agregar texto: ' + alertas.join(' | ') + r.faltan.join(', '));
    const cap = await r.zip.file('OEBPS/Text/cap2.xhtml').async('string');
    afirmar(/Copia de prueba &amp; más/.test(cap) && /class="destacado"/.test(cap), 'capítulo: ' + cap.slice(0, 300));
    detalles.push('texto');
    // Dividir: una parte por capítulo, cada una con su imagen y nada más.
    await abrir('/herramientas/dividir-epub.html'); await esperar(300);
    alertas = []; W().alert = (m) => alertas.push(m);
    W().mostrarArchivo(new (W().File)([libro], 'libro.epub'));
    D().getElementById('botonDividir').click();
    const zipPartes = await W().JSZip.loadAsync(new (W().Uint8Array)(await descargarResultado()));
    const partes = Object.keys(zipPartes.files).sort();
    afirmar(!alertas.length && partes.length === 3, 'dividir: ' + alertas.join(' | ') + partes.join(', '));
    const parte2 = await revisarEpub(await zipPartes.file(partes[1]).async('uint8array'));
    afirmar(!parte2.faltan.length && parte2.spine === 'cap2', 'parte 2: ' + parte2.faltan.join(', ') + ' spine ' + parte2.spine);
    afirmar(parte2.zip.file('OEBPS/Images/img2.png') && !parte2.zip.file('OEBPS/Images/img1.png'), 'parte 2 con imágenes ajenas');
    detalles.push('dividir en ' + partes.length);
    // Unir dos copias: el segundo libro entra entero en su carpeta.
    await abrir('/herramientas/unir-epub.html'); await esperar(300);
    alertas = []; W().alert = (m) => alertas.push(m);
    W().agregarArchivos([new (W().File)([libro], 'a.epub'), new (W().File)([libro], 'b.epub')]);
    D().getElementById('botonUnir').click();
    r = await revisarEpub(await descargarResultado());
    afirmar(!alertas.length && !r.faltan.length, 'unir: ' + alertas.join(' | ') + r.faltan.join(', '));
    afirmar(r.spine.split(',').length === 6 && r.zip.file('OEBPS/libro2/OEBPS/Styles/estilo.css'), 'unir: spine ' + r.spine);
    detalles.push('unir (' + r.spine.split(',').length + ' capítulos)');
    // Reparar un libro sano: nivel 1, sin rearmar.
    await abrir('/herramientas/reparar-epub.html'); await esperar(300);
    alertas = []; W().alert = (m) => alertas.push(m);
    W().mostrarArchivo(new (W().File)([libro], 'libro.epub'));
    D().getElementById('botonReparar').click();
    r = await revisarEpub(await descargarResultado());
    afirmar(!alertas.length && !r.faltan.length, 'reparar: ' + alertas.join(' | ') + r.faltan.join(', '));
    afirmar(D().getElementById('mensajeNivel').textContent.includes('✅'), 'reparar: ' + D().getElementById('mensajeNivel').textContent);
    detalles.push('reparar');
    return detalles.join(', ');
  });

  await prueba('EPUB: editar metadatos (guardado compartido)', async () => {
    const libro = await epubDePrueba();
    await abrir('/herramientas/editar-metadatos-epub.html'); await esperar(300);
    const alertas = []; W().alert = (m) => alertas.push(m);
    await W().cargarLibro(new (W().File)([libro], 'libro.epub'));
    D().getElementById('campoTitulo').value = 'Título nuevo';
    D().getElementById('botonGuardar').click();
    const r = await revisarEpub(await descargarResultado());
    afirmar(!alertas.length && !r.faltan.length, 'editar: ' + alertas.join(' | ') + r.faltan.join(', '));
    const opf = await r.zip.file('OEBPS/content.opf').async('string');
    afirmar(/<dc:title[^>]*>Título nuevo</.test(opf), 'título: ' + (opf.match(/<dc:title[^>]*>[^<]*/) || [''])[0]);
    return 'título cambiado, el resto intacto';
  });

  await prueba('EPUB con fuentes ofuscadas: se abre y al dividir quedan legibles', async () => {
    const libro = await epubDePrueba(true);
    const original = new Uint8Array(await (await W().fetch('/vendor/google-fonts/d2761cb6ca3b.woff2')).arrayBuffer());
    await abrir('/herramientas/dividir-epub.html'); await esperar(300);
    const alertas = []; W().alert = (m) => alertas.push(m);
    W().mostrarArchivo(new (W().File)([libro], 'libro.epub'));
    D().getElementById('botonDividir').click();
    const zipPartes = await W().JSZip.loadAsync(new (W().Uint8Array)(await descargarResultado()));
    afirmar(!alertas.length, 'error: ' + alertas.join(' | '));
    const parte = await W().JSZip.loadAsync(await zipPartes.file(Object.keys(zipPartes.files)[0]).async('uint8array'));
    const fuente = await parte.file('OEBPS/Fonts/fuente.woff2').async('uint8array');
    afirmar(!parte.file('META-INF/encryption.xml') && fuente.length === original.length && fuente.every((b, i) => b === original[i]), 'la fuente de la parte no es la original');
    return 'fuente desofuscada, idéntica a la original';
  });

  // SheetJS < 0.20.2 tiene dos CVE al leer archivos maliciosos
  // (CVE-2023-30533, CVE-2024-22363): ninguna herramienta puede cargar una así.
  await prueba('SheetJS sin versiones vulnerables (herramientas de Excel)', async () => {
    const versiones = new Set();
    let paginas = 0;
    for (const p of PAGINAS.filter((x) => x.startsWith('/herramientas/'))) {
      await abrir(p); await esperar(60);
      const X = W().XLSX;
      if (!X || !X.utils) continue;
      paginas++;
      versiones.add(X.version);
      const [a, b, c] = X.version.split('.').map(Number);
      afirmar(a > 0 || b > 20 || (b === 20 && c >= 2), p + ' carga SheetJS ' + X.version);
    }
    afirmar(paginas >= 12, 'solo ' + paginas + ' páginas con SheetJS');
    return paginas + ' páginas, versión ' + [...versiones].join(', ');
  });

  // Unir arma los nombres de hoja con el nombre de cada archivo. SheetJS
  // rechaza los que Excel no admite (apóstrofo al borde, ":", "History") y
  // Excel trata "Datos" y "datos" como la misma hoja.
  await prueba('unir-excel: nombres de hoja válidos para Excel', async () => {
    await abrir('/herramientas/unir-excel.html'); await esperar(300);
    const X = W().XLSX;
    const alertas = []; W().alert = (m) => alertas.push(m);
    const libro = (texto) => {
      const l = X.utils.book_new();
      X.utils.book_append_sheet(l, X.utils.aoa_to_sheet([[texto]]), 'Hoja1');
      return X.write(l, { type: 'array', bookType: 'xlsx' });
    };
    const nombres = ["'notas'.xlsx", 'Datos.xlsx', 'datos.xlsx', 'a:b.xlsx', 'History.xlsx'];
    W().agregarArchivos(nombres.map((n) => new (W().File)([libro(n)], n)));
    D().getElementById('botonUnir').click();
    const enlace = await hasta(() => D().getElementById('resultado').classList.contains('visible') && D().getElementById('enlaceDescarga').href);
    afirmar(!alertas.length, 'error al unir: ' + alertas.join(' | '));
    afirmar(enlace, 'no terminó de unir');
    const unido = X.read(new Uint8Array(await (await fetch(enlace)).arrayBuffer()), { type: 'array' });
    const hojas = unido.SheetNames;
    afirmar(hojas.length === 5, 'hojas: ' + hojas.join(', '));
    afirmar(new Set(hojas.map((h) => h.toLowerCase())).size === 5, 'nombres repetidos para Excel: ' + hojas.join(', '));
    afirmar(hojas.every((h) => !/[:\\/?*\[\]]|^'|'$/.test(h) && h.toLowerCase() !== 'history'), 'nombre inválido: ' + hojas.join(', '));
    return hojas.join(' · ');
  });

  // Las herramientas trabajan con fechas como número de serie + formato de
  // Excel (sin cellDates): el texto tiene que salir igual que en Excel.
  await prueba('convertir-excel-csv: fecha con su formato', async () => {
    await abrir('/herramientas/convertir-excel-csv.html'); await esperar(300);
    const X = W().XLSX;
    const hoja = X.utils.aoa_to_sheet([['Fecha', 'Monto']]);
    hoja.A2 = { t: 'n', v: 46297, z: 'dd/mm/yyyy' }; // 2 de octubre de 2026
    hoja.B2 = { t: 'n', v: 1234.5, z: '#,##0.00' };
    hoja['!ref'] = 'A1:B2';
    const l = X.utils.book_new(); X.utils.book_append_sheet(l, hoja, 'Hoja1');
    await W().mostrarArchivo(new (W().File)([X.write(l, { type: 'array', bookType: 'xlsx' })], 'fechas.xlsx'));
    D().getElementById('selectorSeparador').value = ',';
    D().getElementById('botonConvertir').click();
    const enlace = await hasta(() => D().getElementById('enlaceDescarga').href.startsWith('blob:') && D().getElementById('enlaceDescarga').href);
    const csv = (await (await fetch(enlace)).text()).replace(/^﻿/, '');
    afirmar(csv.includes('02/10/2026'), 'la fecha salió así: ' + JSON.stringify(csv));
    afirmar(csv.includes('"1,234.50"'), 'el monto salió así: ' + JSON.stringify(csv));
    return JSON.stringify(csv);
  });

  // Si el worker no arranca (ruta o tipo MIME mal), PDF.js no falla: corre
  // todo en la página ("fake worker") y las demás pruebas pasarían igual.
  // Por eso se espía el constructor de Worker antes del primer documento.
  for (const ruta of ['/herramientas/rotar-pdf.html', '/en/herramientas/rotar-pdf.html']) {
    await prueba('PDF.js con su worker (' + (ruta.startsWith('/en/') ? 'en' : 'es') + ')', async () => {
      await abrir(ruta); await esperar(300);
      const creados = [];
      const Original = W().Worker;
      W().Worker = class extends Original { constructor(url, op) { creados.push(String(url)); super(url, op); } };
      const doc = await W().abrirPdf(await (await pdfDePrueba(3)).arrayBuffer());
      afirmar(doc.numPages === 3, 'leyó ' + doc.numPages + ' páginas');
      afirmar(creados.some((u) => /pdf\.worker\.min\.mjs$/.test(u)), 'sin worker propio (¿fake worker?): ' + JSON.stringify(creados));
      return '3 páginas leídas, worker ' + creados[0].replace(/^.*\/vendor\//, 'vendor/');
    });
  }

  // PDF → PDF/A incrusta un perfil de color de vendor/. Con una ruta relativa
  // la versión en inglés lo buscaba en /en/assets/ y no convertía (2026-10-02).
  for (const ruta of ['/herramientas/pdf-a-pdfa.html', '/en/herramientas/pdf-a-pdfa.html']) {
    await prueba('PDF → PDF/A con su perfil de color (' + (ruta.startsWith('/en/') ? 'en' : 'es') + ')', async () => {
      await abrir(ruta); await esperar(300);
      const alertas = []; W().alert = (m) => alertas.push(m);
      await W().mostrarArchivo(await pdfDePrueba(1));
      D().getElementById('botonConvertir').click();
      const enlace = await hasta(() => alertas.length || (D().getElementById('resultado').classList.contains('visible') && D().getElementById('enlaceDescarga').href));
      afirmar(!alertas.length, 'error: ' + alertas.join(' | '));
      // pdf-lib comprime los objetos: se lee el catálogo, no el texto crudo.
      const P = W().PDFLib;
      const doc = await P.PDFDocument.load(await (await W().fetch(enlace)).arrayBuffer());
      const intents = doc.catalog.lookup(P.PDFName.of('OutputIntents'));
      afirmar(intents && intents.size() > 0, 'el PDF no trae OutputIntents');
      const perfil = intents.lookup(0).lookup(P.PDFName.of('DestOutputProfile'));
      afirmar(perfil && perfil.getContents().length > 100, 'OutputIntent sin perfil ICC');
      return 'OutputIntent con perfil ICC de ' + perfil.getContents().length + ' bytes';
    });
  }

  // Cada herramienta que carga PDF.js lo hace con assets/js/pdfjs.js: que en
  // todas abra un PDF (orden de scripts, ruta del worker) y, si tiene
  // miniaturas, la dibuje.
  await prueba('PDF.js en todas las herramientas que lo usan', async () => {
    await abrir('/herramientas/rotar-pdf.html'); await esperar(300); // tiene pdf-lib
    const bytes = new Uint8Array(await (await pdfDePrueba(2)).arrayBuffer());
    let paginas = 0;
    for (const p of PAGINAS.filter((x) => x.startsWith('/herramientas/'))) {
      await abrir(p); await esperar(60);
      if (typeof W().abrirPdf !== 'function') continue;
      paginas++;
      const archivo = new (W().File)([bytes], 'prueba.pdf', { type: 'application/pdf' });
      const doc = await W().abrirPdf(await archivo.arrayBuffer());
      afirmar(doc.numPages === 2, p + ': leyó ' + doc.numPages + ' páginas');
      if (typeof W().generarMiniaturaPdf === 'function') {
        const { elemento } = await W().generarMiniaturaPdf(archivo, 0.2);
        afirmar(elemento.width > 0, p + ': miniatura vacía');
      }
    }
    afirmar(paginas >= 20, 'solo ' + paginas + ' herramientas con PDF.js');
    return paginas + ' herramientas';
  });

  // Texto (fuente estándar no incrustada), dibujo de página y una imagen
  // JPEG 2000: desde la 5.x esta última la decodifica el WebAssembly de
  // vendor/pdf.js/<versión>/wasm/, que antes venía dentro de la librería.
  // JPX de 16×16 px, rojo (200, 30, 30), generado con Pillow/OpenJPEG.
  const JPX_ROJO = '/0//UQAvAAAAAAAQAAAAEAAAAAAAAAAAAAAAEAAAABAAAAAAAAAAAAADBwEBBwEBBwEB/1IADAAAAAEABAQEAAH/XAAQQEBISFBISFBISFBISFD/ZAAlAAFDcmVhdGVkIGJ5IE9wZW5KUEVHIHZlcnNpb24gMi41LjT/kAAKAAAAAAApAAH/k8+0CARPz7QICJ/PtAgIn4CAgICAgICAgICAgP/Z';
  // Se arma DENTRO del marco (pdf-lib reconoce arreglos y objetos por
  // instanceof, y los de esta ventana son de otro "realm"). Necesita una
  // página con pdf-lib cargado; devuelve los bytes copiados a esta ventana.
  function armarPdfConTextoEImagenes(jpxBase64) {
    return (async () => {
      const P = PDFLib;
      const doc = await P.PDFDocument.create();
      const helvetica = await doc.embedFont(P.StandardFonts.Helvetica);
      doc.addPage([300, 100]).drawText('HOLA MUNDO', { x: 20, y: 40, size: 30, font: helvetica });
      // Página 2: el JPX ocupando toda la página.
      const jpx = Uint8Array.from(atob(jpxBase64), (c) => c.charCodeAt(0));
      const ref = doc.context.register(doc.context.stream(jpx, { Type: 'XObject', Subtype: 'Image', Width: 16, Height: 16, Filter: 'JPXDecode' }));
      const p2 = doc.addPage([100, 100]);
      p2.node.setXObject(P.PDFName.of('Jpx'), ref);
      p2.pushOperators(P.pushGraphicsState(), P.concatTransformationMatrix(100, 0, 0, 100, 0, 0), P.drawObject('Jpx'), P.popGraphicsState());
      // Página 3: un PNG azul.
      const c = document.createElement('canvas'); c.width = c.height = 20;
      const cx = c.getContext('2d'); cx.fillStyle = 'rgb(20, 40, 210)'; cx.fillRect(0, 0, 20, 20);
      const png = await doc.embedPng(await (await new Promise((r) => c.toBlob(r, 'image/png'))).arrayBuffer());
      doc.addPage([100, 100]).drawImage(png, { x: 0, y: 0, width: 100, height: 100 });
      return doc.save();
    })();
  }
  async function pdfConTextoEImagenes() {
    return new Uint8Array(await W().eval('(' + armarPdfConTextoEImagenes + ')')(JPX_ROJO));
  }
  async function dibujarPagina(doc, n) {
    const pagina = await doc.getPage(n);
    const viewport = pagina.getViewport({ scale: 1 });
    const canvas = D().createElement('canvas');
    canvas.width = viewport.width; canvas.height = viewport.height;
    await pagina.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
    return canvas.getContext('2d');
  }

  await prueba('PDF.js: texto, dibujo y JPEG 2000 (WebAssembly)', async () => {
    await abrir('/herramientas/rotar-pdf.html'); await esperar(300);
    const doc = await W().abrirPdf(new (W().Uint8Array)(await pdfConTextoEImagenes()));
    const texto = (await (await doc.getPage(1)).getTextContent()).items.map((i) => i.str).join('');
    afirmar(texto.includes('HOLA MUNDO'), 'texto leído: ' + JSON.stringify(texto));
    const pixeles = (await dibujarPagina(doc, 1)).getImageData(0, 0, 300, 100).data;
    let oscuros = 0;
    for (let i = 0; i < pixeles.length; i += 4) if (pixeles[i + 3] > 200 && pixeles[i] < 100) oscuros++;
    afirmar(oscuros > 300, 'la página 1 no dibujó el texto (' + oscuros + ' píxeles oscuros)');
    const [r, g, b] = (await dibujarPagina(doc, 2)).getImageData(50, 50, 1, 1).data;
    afirmar(Math.abs(r - 200) < 25 && g < 60 && b < 60, 'el JPEG 2000 se dibujó rgb(' + [r, g, b] + '), se esperaba ~rgb(200,30,30)');
    return '"' + texto + '", ' + oscuros + ' px de texto, JPX rgb(' + [r, g, b] + ')';
  });

  await prueba('extraer-imagenes: JPEG 2000 y PNG', async () => {
    await abrir('/herramientas/rotar-pdf.html'); await esperar(300); // tiene pdf-lib
    const bytes = await pdfConTextoEImagenes();
    await abrir('/herramientas/extraer-imagenes.html'); await esperar(300);
    const alertas = []; W().alert = (m) => alertas.push(m);
    await W().mostrarArchivo(new (W().File)([bytes], 'prueba.pdf', { type: 'application/pdf' }));
    D().getElementById('botonExtraer').click();
    const enlace = await hasta(() => D().getElementById('resultado').classList.contains('visible') && D().getElementById('enlaceDescarga').href);
    afirmar(!alertas.length, 'error: ' + alertas.join(' | '));
    afirmar(enlace, 'no terminó (¿sin imágenes?): ' + D().getElementById('sinImagenes').style.display);
    const zip = await W().JSZip.loadAsync(await (await W().fetch(enlace)).arrayBuffer());
    const nombres = Object.keys(zip.files);
    afirmar(nombres.length === 2, 'imágenes: ' + nombres.join(', '));
    return nombres.join(', ');
  });

  await prueba('OCR completo (motor y diccionario propios)', async () => {
    await abrir('/herramientas/ocr-pdf.html'); await esperar(300);
    await W().agregarArchivos(archivos([await imagenDePrueba('HOLA 2026', 700, 220)]).files);
    const boton = D().getElementById('botonOcr');
    await hasta(() => !boton.disabled, 10000);
    boton.click();
    afirmar(await hasta(() => D().getElementById('resultado').classList.contains('visible'), 180000), 'el OCR no terminó');
    const texto = await (await fetch(D().getElementById('enlaceDescargaTxt').href)).text();
    afirmar(/HOLA\s*2026/.test(texto), 'reconoció: ' + JSON.stringify(texto.trim().slice(0, 40)));
    return 'reconoció "HOLA 2026"';
  });

  await prueba('fuentes propias (formateador de fórmulas)', async () => {
    await abrir('/herramientas/formateador-formulas-excel.html');
    await Promise.all([D().fonts.load('16px "Public Sans"'), D().fonts.load('16px "JetBrains Mono"')]);
    const fam = new Set([...D().fonts].filter((f) => f.status === 'loaded').map((f) => f.family.replace(/"/g, '')));
    afirmar(fam.has('Public Sans') && fam.has('JetBrains Mono'), 'cargadas: ' + [...fam].join(', '));
    return 'Public Sans y JetBrains Mono';
  });

  // El formateador escapa solo & < > a propósito (busca comillas después de
  // escapar); con el escaparHtml completo se rompía el resaltado de textos.
  await prueba('formateador de fórmulas: resalta textos y hojas entre comillas', async () => {
    await abrir('/herramientas/formateador-formulas-excel.html'); await esperar(200);
    const entrada = D().getElementById('input-formula'), salida = D().getElementById('output-formula');
    entrada.value = '=SI(A1>0;"sí";\'Hoja 1\'!B2)';
    entrada.dispatchEvent(new (W().Event)('input', { bubbles: true }));
    // Mientras no hay entrada se ve una fórmula de ejemplo: esperar la propia.
    afirmar(await hasta(() => salida.textContent.includes('Hoja 1'), 8000), 'no formateó la fórmula: ' + salida.textContent.slice(0, 80));
    const strs = [...salida.querySelectorAll('.tok-str')].map((x) => x.textContent);
    afirmar(strs.includes('"sí"'), 'texto entre comillas no resaltado; resaltados: ' + JSON.stringify(strs));
    afirmar(!/&amp;#39;|&amp;quot;/.test(salida.innerHTML) && !salida.textContent.includes('&#'), 'entidades rotas en: ' + salida.textContent.slice(0, 80));
    afirmar(salida.textContent.includes("'Hoja 1'!B2"), 'la referencia a hoja se alteró: ' + salida.textContent.slice(0, 80));
    return 'resaltado "sí" y \'Hoja 1\'!B2 intacta';
  });

  await prueba('ejemplos de Markdown en el idioma de la página', async () => {
    const casos = [];
    for (const [idioma, titulo, tabla, csv] of [['es', '## Hola mundo', '| Nombre | Edad |', 'Nombre,Edad'], ['en', '## Hello world', '| Name | Age |', 'Name,Age']]) {
      const pre = idioma === 'en' ? '/en' : '';
      await abrir(pre + '/herramientas/markdown-word.html'); await esperar(200);
      const md = D().getElementById('entradaMarkdown').value;
      afirmar(md.startsWith(titulo), idioma + ' markdown-word: ' + JSON.stringify(md.slice(0, 30)));
      await abrir(pre + '/herramientas/markdown-tabla-csv.html'); await esperar(200);
      const entrada = D().getElementById('entrada');
      afirmar(entrada.value.startsWith(tabla), idioma + ' tabla: ' + JSON.stringify(entrada.value.slice(0, 30)));
      const otra = [...D().querySelectorAll('input[name="direccion"]')].find((r) => !r.checked);
      otra.checked = true; otra.dispatchEvent(new (W().Event)('change', { bubbles: true }));
      afirmar(entrada.value.startsWith(csv), idioma + ' CSV: ' + JSON.stringify(entrada.value.slice(0, 30)));
      casos.push(idioma);
    }
    return 'es y en, las dos direcciones';
  });

  // Markdown → HTML: la vista previa se limpia (un .md puede traer HTML con
  // código); el código que se copia queda tal cual lo genera marked.
  await prueba('markdown-html: la vista previa no ejecuta scripts ni atributos onload', async () => {
    await abrir('/herramientas/markdown-html.html'); await esperar(200);
    const entrada = D().getElementById('entrada'), vista = D().getElementById('vistaPrevia');
    entrada.value = '# Título de prueba\n\n<script>window.__pwn = 1</script>\n\n<svg onload="window.__pwn = 2"></svg>\n\nTexto normal.';
    entrada.dispatchEvent(new (W().Event)('input', { bubbles: true }));
    afirmar(await hasta(() => vista.textContent.includes('Texto normal'), 8000), 'no renderizó: ' + vista.textContent.slice(0, 80));
    afirmar(!vista.querySelector('script') && !vista.querySelector('[onload]'), 'quedó código en la vista previa: ' + vista.innerHTML.slice(0, 120));
    afirmar(!W().__pwn, 'el script de la vista previa se ejecutó');
    return 'script y onload fuera; texto y título intactos';
  });

  // Ruta de imagen relativa: se copia tal cual (la ayuda de Dividir Markdown
  // y las notas de la herramienta dependen de esto).
  await prueba('markdown-html: imagen con ruta relativa se conserva tal cual', async () => {
    await abrir('/herramientas/markdown-html.html'); await esperar(200);
    const entrada = D().getElementById('entrada');
    entrada.value = '![foto](fotos/vacaciones.png)';
    entrada.dispatchEvent(new (W().Event)('input', { bubbles: true }));
    await esperar(150);
    const html = D().getElementById('vistaPrevia').innerHTML;
    afirmar(html.includes('src="fotos/vacaciones.png"'), 'ruta alterada: ' + html.slice(0, 120));
    return 'src="fotos/vacaciones.png" intacto';
  });

  // Modos de fusión del editor avanzado: con una capa sin ajustes, Multiplicar
  // funde la imagen consigo misma (gris 128 -> 64) y Trama la aclara (-> 192).
  // Se mira el píxel en la vista previa Y en la imagen descargada, que se
  // componen por caminos distintos.
  for (const ruta of ['/herramientas/editor-avanzado-imagen.html', '/en/herramientas/editor-avanzado-imagen.html']) {
    await prueba('editor avanzado: modos de fusión (' + (ruta.startsWith('/en/') ? 'en' : 'es') + ')', async () => {
      await abrir(ruta); await esperar(300);
      const c = D().createElement('canvas'); c.width = 64; c.height = 64;
      const cx = c.getContext('2d'); cx.fillStyle = 'rgb(128,128,128)'; cx.fillRect(0, 0, 64, 64);
      const gris = new (W().File)([await new Promise((r) => c.toBlob(r, 'image/png'))], 'gris.png', { type: 'image/png' });
      const entrada = D().getElementById('entradaArchivo');
      entrada.files = archivos([gris]).files;
      entrada.dispatchEvent(new (W().Event)('change'));
      const selector = await hasta(() => D().querySelector('#listaCapas .fila-opacidad select'), 15000);
      afirmar(selector, 'la capa no tiene selector de modo de fusión');
      afirmar(selector.options.length === 16, selector.options.length + ' modos en vez de 16');
      const nombre = [...selector.options].find((o) => o.value === 'multiply').textContent;
      afirmar(nombre === (ruta.startsWith('/en/') ? 'Multiply' : 'Multiplicar'), 'nombre del modo: ' + nombre);
      const lienzo = D().getElementById('canvasTrabajo');
      const pixel = (cv) => cv.getContext('2d').getImageData(cv.width >> 1, cv.height >> 1, 1, 1).data[0];
      const valores = {};
      for (const modo of ['normal', 'multiply', 'screen']) {
        selector.value = modo; selector.dispatchEvent(new (W().Event)('change'));
        await esperar(250);
        valores[modo] = pixel(lienzo);
      }
      afirmar(Math.abs(valores.normal - 128) <= 2 && Math.abs(valores.multiply - 64) <= 2 && Math.abs(valores.screen - 192) <= 2,
        'vista previa: ' + JSON.stringify(valores));
      selector.value = 'multiply'; selector.dispatchEvent(new (W().Event)('change'));
      D().getElementById('botonDescargarAvanzado').click();
      const enlace = D().getElementById('enlaceDescarga');
      afirmar(await hasta(() => enlace.href.startsWith('blob:'), 15000), 'la descarga no se generó');
      const img = new (W().Image)(); img.src = enlace.href; await img.decode();
      const cd = D().createElement('canvas'); cd.width = img.width; cd.height = img.height;
      cd.getContext('2d').drawImage(img, 0, 0);
      afirmar(Math.abs(pixel(cd) - 64) <= 2, 'descarga en Multiplicar: ' + pixel(cd));
      return 'vista previa ' + JSON.stringify(valores) + ', descarga ' + pixel(cd);
    });
  }

  // Máscaras del editor avanzado. Imagen mitad oscura (40) y mitad clara
  // (220); la capa en Multiplicar y sin ajustes oscurece justo donde se
  // aplica (40 -> 6, 220 -> 190), así cada máscara deja una huella exacta.
  await prueba('editor avanzado: máscaras (luz, invertir, pincel, degradado, descarga, proyecto viejo)', async () => {
    await abrir('/herramientas/editor-avanzado-imagen.html'); await esperar(300);
    const base = D().createElement('canvas'); base.width = 100; base.height = 50;
    const cb = base.getContext('2d');
    cb.fillStyle = 'rgb(40,40,40)'; cb.fillRect(0, 0, 50, 50);
    cb.fillStyle = 'rgb(220,220,220)'; cb.fillRect(50, 0, 50, 50);
    const png = await new Promise((r) => base.toBlob(r, 'image/png'));
    const entrada = D().getElementById('entradaArchivo');
    entrada.files = archivos([new (W().File)([png], 'mitades.png', { type: 'image/png' })]).files;
    entrada.dispatchEvent(new (W().Event)('change'));
    const selector = await hasta(() => D().querySelector('#listaCapas .fila-opacidad select'), 15000);
    afirmar(selector, 'el editor no cargó la imagen');
    selector.value = 'multiply'; selector.dispatchEvent(new (W().Event)('change'));
    const lienzo = D().getElementById('canvasTrabajo');
    // A un cuarto del alto: la línea de ayuda del degradado pasa por la mitad.
    const px = (cv, x) => cv.getContext('2d').getImageData(Math.round(cv.width * x), cv.height >> 2, 1, 1).data[0];
    const leer = async () => { await esperar(250); return [px(lienzo, 0.25), px(lienzo, 0.75)]; };
    const cerca = (v, e) => Math.abs(v - e) <= 3;
    const clic = (sel) => D().querySelector(sel).click();
    const huellas = {};

    huellas.toda = await leer();
    afirmar(cerca(huellas.toda[0], 6) && cerca(huellas.toda[1], 190), 'toda: ' + huellas.toda);
    agregarPieza('luz'); clic('#rangosLuz [data-rango="sombras"]');
    afirmar(W().capaActiva().mascara.piezas.length === 1, 'la pieza de luz no reemplazó a "Toda la imagen"');
    huellas.sombras = await leer();
    afirmar(cerca(huellas.sombras[0], 6) && cerca(huellas.sombras[1], 220), 'luz sombras: ' + huellas.sombras);
    const invertir = D().getElementById('invertirMascara');
    invertir.checked = true; invertir.dispatchEvent(new (W().Event)('change'));
    huellas.invertida = await leer();
    afirmar(cerca(huellas.invertida[0], 40) && cerca(huellas.invertida[1], 190), 'invertida: ' + huellas.invertida);

    // La descarga tiene que salir igual que la vista previa.
    D().getElementById('botonDescargarAvanzado').click();
    const enlace = D().getElementById('enlaceDescarga');
    afirmar(await hasta(() => enlace.href.startsWith('blob:'), 15000), 'la descarga no se generó');
    const img = new (W().Image)(); img.src = enlace.href; await img.decode();
    const cd = D().createElement('canvas'); cd.width = img.width; cd.height = img.height;
    cd.getContext('2d').drawImage(img, 0, 0);
    huellas.descarga = [px(cd, 0.25), px(cd, 0.75)];
    afirmar(cerca(huellas.descarga[0], 40) && cerca(huellas.descarga[1], 190), 'descarga invertida: ' + huellas.descarga);
    invertir.checked = false; invertir.dispatchEvent(new (W().Event)('change'));

    // Sin piezas (lo que antes era la base "Nada") + un pincel que suma
    // sobre la mitad izquierda; después "Toda" y un pincel que quita la
    // derecha.
    vaciarZona();
    huellas.nada = await leer();
    afirmar(cerca(huellas.nada[0], 40) && cerca(huellas.nada[1], 220), 'nada: ' + huellas.nada);
    const estampar = (x) => {
      W().estamparPincel((ctx, modo) => W().pintarPunto(ctx, { x: lienzo.width * x, y: lienzo.height / 4 }, 12, 100, modo));
      W().marcarMascaraSucia(W().capaActiva()); W().programarRenderizado();
    };
    clic('#herramientasLienzo [data-herramienta="pincel"]');
    clic('#modosPincelAvanzado [data-modo-pincel="pintar"]'); estampar(0.25);
    huellas.sumar = await leer();
    afirmar(cerca(huellas.sumar[0], 6) && cerca(huellas.sumar[1], 220), 'nada + sumar: ' + huellas.sumar);
    agregarPieza('toda'); agregarPieza('pincel-restar'); estampar(0.75);
    huellas.quitar = await leer();
    afirmar(cerca(huellas.quitar[0], 6) && cerca(huellas.quitar[1], 220), 'toda - quitar: ' + huellas.quitar);
    afirmar(W().capaActiva().mascara.piezas.map((p) => p.tipo + ':' + p.op).join(' ') === 'pincel:sumar toda:sumar pincel:restar', 'piezas: ' + W().capaActiva().mascara.piezas.map((p) => p.tipo + ':' + p.op).join(' '));
    // El borrador quita lo pintado de la pieza elegida (el pincel que resta).
    clic('#modosPincelAvanzado [data-modo-pincel="borrador"]'); estampar(0.75);
    huellas.borrador = await leer();
    afirmar(cerca(huellas.borrador[1], 190), 'borrador sobre el pincel que resta: ' + huellas.borrador);
    clic('#modosPincelAvanzado [data-modo-pincel="pintar"]');
    clic('#botonBorrarMascaraCapa');
    afirmar(!W().piezaActiva().trazo, '"Borrar lo pintado" no vació la pieza');

    // Degradado lineal de izquierda (entera) a derecha (nada): a 1/4 se
    // aplica 3/4, a 3/4 se aplica 1/4.
    vaciarZona(); agregarPieza('degradado');
    Object.assign(W().piezaActiva().degradado, { tipo: 'lineal', a: { x: 0, y: 0.5 }, b: { x: 1, y: 0.5 } });
    W().programarRenderizado();
    huellas.degradado = await leer();
    afirmar(cerca(huellas.degradado[0], 15) && cerca(huellas.degradado[1], 213), 'degradado: ' + huellas.degradado);

    // Proyecto de la versión 1 (antes de las máscaras): "pincel" con la
    // mitad izquierda pintada tiene que abrir como base "Nada" + lo pintado.
    const pintado = D().createElement('canvas'); pintado.width = 100; pintado.height = 50;
    pintado.getContext('2d').fillStyle = '#fff'; pintado.getContext('2d').fillRect(0, 0, 50, 50);
    const proyecto = {
      version: 1, imagen: base.toDataURL('image/png'), nombreArchivo: 'mitades.png', capaActivaId: 'capa-1',
      capas: [{ id: 'capa-1', nombre: 'Capa 1', visible: true, opacidad: 100, modoFusion: 'multiply', seleccion: 'pincel',
        colorObjetivo: null, tolerancia: 30, ajustes: {}, mascaraPincel: pintado.toDataURL('image/png') }],
    };
    const entradaProyecto = D().getElementById('entradaProyecto');
    entradaProyecto.files = archivos([new (W().File)([JSON.stringify(proyecto)], 'viejo.json', { type: 'application/json' })]).files;
    entradaProyecto.dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => W().capaActiva().mascara.piezas.map((p) => p.tipo).join() === 'pincel', 10000), 'el proyecto viejo no abrió con solo lo pintado');
    huellas.proyectoViejo = await leer();
    afirmar(cerca(huellas.proyectoViejo[0], 6) && cerca(huellas.proyectoViejo[1], 220), 'proyecto v1: ' + huellas.proyectoViejo);
    return Object.entries(huellas).map(([k, v]) => k + ' ' + v).join(' · ');
  });

  // Grupos del editor avanzado, con la misma imagen de dos mitades: el grupo
  // aplica su visibilidad, opacidad, máscara y modo a todo lo de adentro, y
  // subir/bajar cambian el orden (que importa: Multiplicar y Trama no
  // conmutan). Desde los tipos de grupo (2026-10-07) un grupo de ajustes
  // nunca se aísla: su modo funde el resultado entero con lo de abajo.
  await prueba('editor avanzado: grupos (agrupar, ocultar, opacidad, máscara, modo, orden, proyecto)', async () => {
    await abrir('/herramientas/editor-avanzado-imagen.html'); await esperar(300);
    const base = D().createElement('canvas'); base.width = 100; base.height = 50;
    const cb = base.getContext('2d');
    cb.fillStyle = 'rgb(40,40,40)'; cb.fillRect(0, 0, 50, 50);
    cb.fillStyle = 'rgb(220,220,220)'; cb.fillRect(50, 0, 50, 50);
    const png = await new Promise((r) => base.toBlob(r, 'image/png'));
    const entrada = D().getElementById('entradaArchivo');
    entrada.files = archivos([new (W().File)([png], 'mitades.png', { type: 'image/png' })]).files;
    entrada.dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => D().querySelector('#listaCapas .fila-opacidad select'), 15000), 'el editor no cargó la imagen');
    const lienzo = D().getElementById('canvasTrabajo');
    const px = (x) => lienzo.getContext('2d').getImageData(Math.round(lienzo.width * x), lienzo.height >> 2, 1, 1).data[0];
    const leer = async () => { await esperar(250); return [px(0.25), px(0.75)]; };
    const cerca = (v, e) => Math.abs(v - e) <= 3;
    const clic = (sel) => D().querySelector(sel).click();
    // El selector de modo y la opacidad de la fila del elemento elegido.
    const filaActiva = () => D().querySelector('#listaCapas .fila-capa.activa').parentElement;
    const ponerModo = (modo) => { const s = filaActiva().querySelector('select'); s.value = modo; s.dispatchEvent(new (W().Event)('change')); };
    const ponerOpacidad = (v) => { const r = filaActiva().querySelector('input[type=range]'); r.value = v; r.dispatchEvent(new (W().Event)('input')); };
    const h = {};

    ponerModo('multiply');
    const idCapa1 = W().capaActiva().id;
    clic('#botonAgrupar');
    afirmar(W().esGrupo(W().itemActivo()), 'Agrupar no dejó elegido el grupo');
    afirmar(W().itemActivo().clase === 'ajustes', 'agrupar una capa de ajuste no armó un grupo de ajustes: ' + W().itemActivo().clase);
    afirmar(![...filaActiva().querySelector('select').options].some((o) => o.value === 'paso'), 'el grupo todavía ofrece "Pasar a través"');
    afirmar(D().querySelector('[data-app-categoria="Luz"], .panel-ajustes .acordeon-contenido').inert, 'con un grupo elegido, los ajustes de capa no quedaron inactivos');
    h.agrupada = await leer();
    afirmar(cerca(h.agrupada[0], 6) && cerca(h.agrupada[1], 190), 'grupo Normal: ' + h.agrupada);
    filaActiva().querySelector('.icono-capa').click(); // ojo del grupo
    h.oculto = await leer();
    afirmar(cerca(h.oculto[0], 40) && cerca(h.oculto[1], 220), 'grupo oculto: ' + h.oculto);
    filaActiva().querySelector('.icono-capa').click();
    ponerOpacidad(50);
    h.mitad = await leer();
    afirmar(cerca(h.mitad[0], 23) && cerca(h.mitad[1], 205), 'grupo al 50%: ' + h.mitad);
    ponerOpacidad(100);
    agregarPieza('luz'); clic('#rangosLuz [data-rango="sombras"]');
    h.mascara = await leer();
    afirmar(cerca(h.mascara[0], 6) && cerca(h.mascara[1], 220), 'máscara del grupo: ' + h.mascara);
    zonaSoloToda();
    // Multiplicar en el grupo no lo aísla: lo de adentro (40 -> 6) se
    // multiplica otra vez con lo de abajo (6 · 40 / 255 = 1; 190 · 220 / 255 = 164).
    ponerModo('multiply');
    h.grupoMultiplica = await leer();
    afirmar(cerca(h.grupoMultiplica[0], 1) && cerca(h.grupoMultiplica[1], 164), 'grupo de ajustes en Multiplicar: ' + h.grupoMultiplica);
    ponerModo('normal');

    // Proyecto: el árbol se guarda y vuelve igual.
    const raiz = W().ubicar(W().itemActivo().id).lista;
    const proyecto = { version: 7, imagen: base.toDataURL('image/png'), nombreArchivo: 'mitades.png', capaActivaId: idCapa1, capas: raiz.map(W().itemAProyecto) };
    afirmar(proyecto.capas.length === 1 && proyecto.capas[0].tipo === 'grupo' && proyecto.capas[0].hijos.length === 1, 'el proyecto no guardó el grupo: ' + JSON.stringify(proyecto.capas).slice(0, 120));
    const entradaProyecto = D().getElementById('entradaProyecto');
    entradaProyecto.files = archivos([new (W().File)([JSON.stringify(proyecto)], 'grupo.json', { type: 'application/json' })]).files;
    entradaProyecto.dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => W().capaActiva() && W().capaActiva().id === idCapa1 && W().ubicar(idCapa1).padre, 10000), 'el proyecto no volvió con la capa dentro del grupo');
    h.proyecto = await leer();
    afirmar(cerca(h.proyecto[0], 6) && cerca(h.proyecto[1], 190), 'proyecto con grupo: ' + h.proyecto);

    // Orden: una capa nueva en Trama, arriba de la de Multiplicar (dentro del
    // grupo): 40 -> 6 -> 12. Bajarla la pone debajo: 40 -> 74 -> 21. Bajarla
    // otra vez la saca del grupo, sin cambiar el resultado.
    clic('#botonNuevaCapa');
    ponerModo('screen');
    const idTrama = W().capaActiva().id;
    afirmar(W().ubicar(idTrama).padre, 'la capa nueva no quedó dentro del grupo');
    h.tramaArriba = await leer();
    afirmar(cerca(h.tramaArriba[0], 12), 'Trama sobre Multiplicar: ' + h.tramaArriba);
    clic('#botonBajarCapa');
    h.tramaAbajo = await leer();
    afirmar(cerca(h.tramaAbajo[0], 21), 'Trama debajo de Multiplicar: ' + h.tramaAbajo);
    clic('#botonBajarCapa');
    afirmar(!W().ubicar(idTrama).padre, 'bajar desde el borde no la sacó del grupo');
    h.fuera = await leer();
    afirmar(cerca(h.fuera[0], 21), 'Trama fuera del grupo, debajo: ' + h.fuera);

    // Desagrupar deja las capas sueltas, en el mismo orden.
    W().seleccionarItem(W().ubicar(idCapa1).padre.id);
    clic('#botonDesagrupar');
    afirmar(!W().ubicar(idCapa1).padre, 'Desagrupar no soltó la capa');
    h.desagrupado = await leer();
    afirmar(cerca(h.desagrupado[0], 21), 'desagrupado: ' + h.desagrupado);
    return Object.entries(h).map(([k, v]) => k + ' ' + v).join(' · ');
  });

  // Capas con contenido (texto, copia de la foto) del editor avanzado, con
  // la imagen de dos mitades: la copia de la mitad oscura movida a la
  // derecha la tapa (40 | 40); en Multiplicar la oscurece (40 | 35). Los
  // gestos sobre la imagen se hacen con eventos de puntero de verdad.
  await prueba('editor avanzado: texto e imágenes (copia, mover, escalar, girar, texto, descarga, proyecto)', async () => {
    await abrir('/herramientas/editor-avanzado-imagen.html'); await esperar(300);
    const base = D().createElement('canvas'); base.width = 100; base.height = 50;
    const cb = base.getContext('2d');
    cb.fillStyle = 'rgb(40,40,40)'; cb.fillRect(0, 0, 50, 50);
    cb.fillStyle = 'rgb(220,220,220)'; cb.fillRect(50, 0, 50, 50);
    const png = await new Promise((r) => base.toBlob(r, 'image/png'));
    const entrada = D().getElementById('entradaArchivo');
    entrada.files = archivos([new (W().File)([png], 'mitades.png', { type: 'image/png' })]).files;
    entrada.dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => D().querySelector('#listaCapas .fila-opacidad select'), 15000), 'el editor no cargó la imagen');
    const lienzo = D().getElementById('canvasTrabajo');
    lienzo.setPointerCapture = () => {}; // los punteros sintéticos no se pueden capturar
    const px = (cv, x, y = 0.25) => cv.getContext('2d').getImageData(Math.round(cv.width * x), Math.round(cv.height * y), 1, 1).data;
    const leer = async () => { await esperar(250); return [px(lienzo, 0.25)[0], px(lienzo, 0.75)[0]]; };
    const cerca = (v, e, tol = 3) => Math.abs(v - e) <= tol;
    const clic = (sel) => D().querySelector(sel).click();
    const deslizar = (id, v) => { const el = D().getElementById(id); el.value = v; el.dispatchEvent(new (W().Event)('input')); };
    // Punto del lienzo -> punto de la pantalla, para los eventos de puntero.
    const enPantalla = (p) => { const r = lienzo.getBoundingClientRect(); return { clientX: r.left + p.x * r.width / lienzo.width, clientY: r.top + p.y * r.height / lienzo.height }; };
    const arrastrar = async (desde, hasta_) => {
      const o = { bubbles: true, pointerId: 7, pointerType: 'mouse', isPrimary: true, button: 0, buttons: 1 };
      lienzo.dispatchEvent(new (W().PointerEvent)('pointerdown', { ...o, ...enPantalla(desde) }));
      lienzo.dispatchEvent(new (W().PointerEvent)('pointermove', { ...o, ...enPantalla(hasta_) }));
      lienzo.dispatchEvent(new (W().PointerEvent)('pointerup', { ...o, ...enPantalla(hasta_), buttons: 0 }));
      await esperar(200);
    };
    const h = {};

    // Copia de la mitad oscura (máscara de la capa 1 en Sombras), en su lugar.
    agregarPieza('luz'); clic('#rangosLuz [data-rango="sombras"]');
    clic('#botonCopiarFoto');
    const objetoCopia = W().itemActivo();
    afirmar(objetoCopia && objetoCopia.clase === 'objeto', 'Copia de la foto no creó un objeto');
    const copia = W().capaDeContenido(objetoCopia);
    afirmar(copia && copia.contenido && copia.contenido.tipo === 'imagen' && objetoCopia.hijos[0] === copia, 'el objeto no tiene la imagen abajo');
    afirmar(cerca(copia.contenido.lienzo.width, 50, 1) && cerca(copia.contenido.centro.x, 0.25, 0.01), 'la copia no es la mitad izquierda: ' + copia.contenido.lienzo.width + ' en ' + copia.contenido.centro.x);
    h.enSuLugar = await leer();
    afirmar(cerca(h.enSuLugar[0], 40) && cerca(h.enSuLugar[1], 220), 'copia en su lugar: ' + h.enSuLugar);
    deslizar('posX', 75);
    h.movida = await leer();
    afirmar(cerca(h.movida[0], 40) && cerca(h.movida[1], 40), 'copia movida a la derecha: ' + h.movida);
    const fila = D().querySelector('#listaCapas .fila-capa.activa').parentElement;
    const modo = fila.querySelector('select'); modo.value = 'multiply'; modo.dispatchEvent(new (W().Event)('change'));
    h.multiplicada = await leer();
    afirmar(cerca(h.multiplicada[0], 40) && cerca(h.multiplicada[1], 35), 'copia en Multiplicar: ' + h.multiplicada);

    // Gestos: arrastrar el centro la mueve; una esquina la escala; la manija la gira.
    const c = copia.contenido;
    let k = W().cajaDeContenido(copia);
    await arrastrar({ x: k.cx, y: k.cy }, { x: k.cx - 10, y: k.cy + 5 });
    h.gestoMover = [Math.round(c.centro.x * 100), Math.round(c.centro.y * 100)];
    afirmar(h.gestoMover[0] === 65 && h.gestoMover[1] === 60, 'mover con el puntero: ' + h.gestoMover);
    k = W().cajaDeContenido(copia);
    const f = W().pixelesPorPuntoDePantalla();
    let m = W().ContenidoCapa.manijas(k, 28 * f);
    const esquina = m.esquinas[2];
    await arrastrar(esquina, { x: k.cx + (esquina.x - k.cx) * 1.5, y: k.cy + (esquina.y - k.cy) * 1.5 });
    h.gestoEscalar = Math.round(c.escala * 100);
    afirmar(cerca(h.gestoEscalar, 150, 2), 'escalar desde la esquina: ' + h.gestoEscalar);
    k = W().cajaDeContenido(copia);
    m = W().ContenidoCapa.manijas(k, 28 * f);
    const r = Math.hypot(m.girar.x - k.cx, m.girar.y - k.cy);
    await arrastrar(m.girar, { x: k.cx + r, y: k.cy }); // de arriba a la derecha: +90°
    h.gestoGirar = Math.round(c.rotacion);
    afirmar(cerca(h.gestoGirar, 90, 1), 'girar con la manija: ' + h.gestoGirar);
    Object.assign(c, { centro: { x: 0.75, y: 0.5 }, escala: 1, rotacion: 0 });
    W().programarRenderizado();
    h.restaurada = await leer();
    afirmar(cerca(h.restaurada[1], 35), 'copia vuelta a su lugar: ' + h.restaurada);

    // Un ajuste adentro del objeto (+ Capa con el objeto elegido) en Trama
    // aclara solo la copia (derecha 40 -> 74), no la foto (izquierda sigue en
    // 40). Subido afuera del objeto, aclara las dos.
    const modoCopia = D().querySelector('#listaCapas .fila-capa.activa').parentElement.querySelector('select');
    modoCopia.value = 'normal'; modoCopia.dispatchEvent(new (W().Event)('change'));
    clic('#botonNuevaCapa');
    const idAjuste = W().itemActivo().id;
    afirmar(W().ubicar(idAjuste).padre === objetoCopia && W().ubicar(idAjuste).indice === 1, '+ Capa con un objeto elegido no la puso adentro, arriba de su imagen');
    const ajuste = D().querySelector('#listaCapas .fila-capa.activa').parentElement.querySelector('select');
    ajuste.value = 'screen'; ajuste.dispatchEvent(new (W().Event)('change'));
    h.ajusteDelObjeto = await leer();
    afirmar(cerca(h.ajusteDelObjeto[0], 40) && cerca(h.ajusteDelObjeto[1], 74), 'ajuste adentro del objeto: ' + h.ajusteDelObjeto);
    afirmar(D().getElementById('botonBajarCapa').disabled === false, 'no se puede bajar el ajuste');
    clic('#botonSubirCapa');
    afirmar(!W().ubicar(idAjuste).padre, 'subir desde arriba de todo no lo sacó del objeto');
    h.ajusteAfuera = await leer();
    afirmar(cerca(h.ajusteAfuera[0], 74) && cerca(h.ajusteAfuera[1], 74), 'el mismo ajuste afuera del objeto: ' + h.ajusteAfuera);
    // Se deshace para seguir con lo de abajo.
    W().eliminarCapa(idAjuste);
    W().seleccionarItem(copia.id);
    afirmar(W().itemActivo() === objetoCopia, 'elegir la imagen de un objeto no eligió el objeto');
    const filaCopia = D().querySelector('#listaCapas .fila-capa.activa').parentElement.querySelector('select');
    filaCopia.value = 'multiply'; filaCopia.dispatchEvent(new (W().Event)('change'));
    h.vuelta = await leer();
    afirmar(cerca(h.vuelta[1], 35), 'copia otra vez sola en Multiplicar: ' + h.vuelta);

    // Texto: aparece, cambia de color y crece con el tamaño.
    const contar = (cond) => { const d = lienzo.getContext('2d').getImageData(0, 0, lienzo.width, lienzo.height).data; let n = 0; for (let i = 0; i < d.length; i += 4) if (cond(d[i], d[i + 1], d[i + 2])) n++; return n; };
    const rojos = (r, g, b) => r > 180 && g < 90 && b < 90;
    clic('#botonNuevoTexto');
    const objetoTexto = W().itemActivo();
    const capaTexto = W().capaDeContenido(objetoTexto);
    afirmar(W().esObjeto(objetoTexto) && capaTexto && capaTexto.contenido && capaTexto.contenido.tipo === 'texto', '+ Texto no creó un objeto de texto');
    afirmar(!W().ubicar(objetoTexto.id).padre, 'el texto quedó adentro de otro objeto');
    const color = D().getElementById('colorTexto'); color.value = '#ff0000'; color.dispatchEvent(new (W().Event)('input'));
    D().getElementById('textoCapa').value = 'HOLA'; D().getElementById('textoCapa').dispatchEvent(new (W().Event)('input'));
    await esperar(250);
    h.textoRojo = contar(rojos);
    afirmar(h.textoRojo > 5, 'el texto rojo no aparece: ' + h.textoRojo + ' píxeles');
    afirmar(objetoTexto.nombre === 'HOLA', 'el nombre no siguió al texto: ' + objetoTexto.nombre);
    deslizar('tamanoContenido', 250);
    await esperar(250);
    h.textoGrande = contar(rojos);
    afirmar(h.textoGrande > h.textoRojo * 3, 'al 250% el texto no creció: ' + h.textoRojo + ' -> ' + h.textoGrande);

    // Descarga: la copia movida y en Multiplicar también está ahí.
    D().getElementById('botonDescargarAvanzado').click();
    const enlace = D().getElementById('enlaceDescarga');
    afirmar(await hasta(() => enlace.href.startsWith('blob:'), 15000), 'la descarga no se generó');
    const img = new (W().Image)(); img.src = enlace.href; await img.decode();
    const cd = D().createElement('canvas'); cd.width = img.width; cd.height = img.height;
    cd.getContext('2d').drawImage(img, 0, 0);
    h.descarga = [px(cd, 0.25, 0.1)[0], px(cd, 0.75, 0.1)[0]];
    afirmar(cerca(h.descarga[0], 40) && cerca(h.descarga[1], 35), 'descarga: ' + h.descarga);

    // Proyecto: guardado y vuelto a abrir, da lo mismo.
    const antes = [px(lienzo, 0.25, 0.1)[0], px(lienzo, 0.75, 0.1)[0], contar(rojos)];
    const raiz = W().eval('capas');
    const proyecto = { version: 7, imagen: base.toDataURL('image/png'), nombreArchivo: 'mitades.png', capaActivaId: objetoTexto.id, capas: raiz.map(W().itemAProyecto) };
    const entradaProyecto = D().getElementById('entradaProyecto');
    entradaProyecto.files = archivos([new (W().File)([JSON.stringify(proyecto)], 'contenido.json', { type: 'application/json' })]).files;
    entradaProyecto.dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => W().itemActivo() && W().itemActivo() !== objetoTexto && W().itemActivo().id === objetoTexto.id && W().capaDeContenido(W().itemActivo()).id === capaTexto.id, 10000), 'el proyecto no volvió a abrir');
    await esperar(300);
    h.proyecto = [px(lienzo, 0.25, 0.1)[0], px(lienzo, 0.75, 0.1)[0], contar(rojos)];
    afirmar(cerca(h.proyecto[0], antes[0]) && cerca(h.proyecto[1], antes[1]) && cerca(h.proyecto[2], antes[2], antes[2] * 0.05 + 2), 'proyecto: ' + antes + ' -> ' + h.proyecto);
    return Object.entries(h).map(([k2, v]) => k2 + ' ' + v).join(' · ');
  });

  // Tipos de grupo (2026-10-07): las reglas de qué va adentro de qué, el
  // motivo al soltar donde no se puede, y los proyectos viejos convertidos.
  // Foto de dos mitades (40 | 220) y una imagen gris 40 de 50x50 puesta sobre
  // la mitad derecha. La equivalencia exacta con la versión vieja se comprobó
  // aparte, píxel por píxel, contra la rama anterior (hallazgos.md).
  await prueba('editor avanzado: tipos de grupo (reglas, soltar, proyectos viejos)', async () => {
    await abrir('/herramientas/editor-avanzado-imagen.html'); await esperar(300);
    const base = D().createElement('canvas'); base.width = 100; base.height = 50;
    const cb = base.getContext('2d');
    cb.fillStyle = 'rgb(40,40,40)'; cb.fillRect(0, 0, 50, 50);
    cb.fillStyle = 'rgb(220,220,220)'; cb.fillRect(50, 0, 50, 50);
    const gris = D().createElement('canvas'); gris.width = 50; gris.height = 50;
    gris.getContext('2d').fillStyle = 'rgb(40,40,40)'; gris.getContext('2d').fillRect(0, 0, 50, 50);
    const lienzo = D().getElementById('canvasTrabajo');
    const leer = async () => { await esperar(300); return [0.25, 0.75].map((x) => lienzo.getContext('2d').getImageData(Math.round(lienzo.width * x), lienzo.height >> 1, 1, 1).data[0]); };
    const cerca = (v, e) => Math.abs(v - e) <= 3;
    const forma = () => JSON.stringify(W().eval('capas').map(function f(c) { return W().esGrupo(c) ? c.clase + (c.aisladoV6 ? '*' : '') + '[' + c.hijos.map(f).join(',') + ']' : (c.contenido ? 'imagen' : 'ajuste'); }));
    const abrirProyecto = async (version, capas) => {
      const p = { version, imagen: base.toDataURL('image/png'), nombreArchivo: 'mitades.png', capaActivaId: capas[0].id, capas };
      const e = D().getElementById('entradaProyecto'), antes = W().eval('capas');
      e.files = archivos([new (W().File)([JSON.stringify(p)], 'p.json', { type: 'application/json' })]).files;
      e.dispatchEvent(new (W().Event)('change'));
      afirmar(await hasta(() => W().eval('capas') !== antes && W().ubicar(capas[0].id), 10000), 'el proyecto no abrió');
    };
    const capa = (id, extra = {}) => ({ id, nombre: id, visible: true, opacidad: 100, modoFusion: 'normal', mascara: null, ajustes: {}, ...extra });
    const imagen = (id, extra = {}) => capa(id, { contenido: { tipo: 'imagen', lienzo: gris.toDataURL('image/png'), centro: { x: 0.75, y: 0.5 }, escala: 1, rotacion: 0, voltearH: false, voltearV: false }, ...extra });
    const h = {};

    // Una imagen con ajustes propios (invertir: 40 -> 215) pasa a ser un objeto con un ajuste adentro.
    await abrirProyecto(6, [capa('capa-1'), imagen('capa-2', { ajustes: { invertir: 100 } })]);
    h.formaImagen = forma();
    afirmar(h.formaImagen === '["ajuste","objeto[imagen,ajuste]"]', 'conversión de una imagen con ajustes: ' + h.formaImagen);
    h.imagen = await leer();
    afirmar(cerca(h.imagen[0], 40) && cerca(h.imagen[1], 215), 'imagen convertida: ' + h.imagen);

    // Un grupo aislado (v6) con la imagen y un ajuste en Trama: aclara solo la imagen (40 -> 74).
    await abrirProyecto(6, [capa('capa-1'), { ...capa('capa-3'), tipo: 'grupo', hijos: [imagen('capa-2'), capa('capa-4', { modoFusion: 'screen' })] }]);
    h.formaAislado = forma();
    afirmar(h.formaAislado === '["ajuste","carpeta*[objeto[imagen],ajuste]"]', 'conversión de un grupo aislado: ' + h.formaAislado);
    h.aislado = await leer();
    afirmar(cerca(h.aislado[0], 40) && cerca(h.aislado[1], 74), 'grupo aislado convertido: ' + h.aislado);

    // El mismo grupo en "Pasar a través" (v6): aclara las dos mitades.
    await abrirProyecto(6, [capa('capa-1'), { ...capa('capa-3'), tipo: 'grupo', modoFusion: 'paso', hijos: [imagen('capa-2'), capa('capa-4', { modoFusion: 'screen' })] }]);
    h.formaPaso = forma();
    afirmar(h.formaPaso === '["ajuste","carpeta[objeto[imagen],ajuste]"]', 'conversión de un grupo "Pasar a través" con un objeto: ' + h.formaPaso);
    h.paso = await leer();
    afirmar(cerca(h.paso[0], 74) && cerca(h.paso[1], 74), 'grupo "Pasar a través" convertido: ' + h.paso);

    // Reglas: lo que no admite cada tipo, con su motivo.
    const T = W().TiposGrupo, carpetaV = W().ubicar('capa-3').item, objeto = carpetaV.hijos[0], ajusteV = carpetaV.hijos[1];
    const grupoAjustes = W().crearGrupo(null, 'ajustes');
    h.reglas = [
      T.motivoNoPuede(objeto, grupoAjustes, 0), T.motivoNoPuede(W().crearGrupo(null, 'carpeta'), objeto, 1), T.motivoNoPuede(ajusteV, objeto, 0),
      T.motivoNoPuede(objeto.hijos[0], null, 0), T.motivoNoPuede(ajusteV, objeto, 1), T.motivoNoPuede(objeto, carpetaV, 0),
    ].join(',');
    afirmar(h.reglas === 'motivo-ajustes-solo-ajustes,motivo-objeto-solo-ajustes,motivo-debajo-del-contenido,motivo-contenido-fijo,,', 'reglas: ' + h.reglas);
    // Soltar el objeto adentro del ajuste no se puede (y se dice por qué); arriba de la imagen del objeto, sí.
    h.soltar = [W().motivoNoSoltar(objeto.id, ajusteV, 'arriba'), W().motivoNoSoltar(ajusteV.id, objeto.hijos[0], 'arriba'), W().motivoNoSoltar(carpetaV.id, objeto, 'dentro')].join(',');
    afirmar(h.soltar === ',,motivo-dentro-de-si' || h.soltar === ',,motivo-objeto-solo-ajustes', 'soltar: ' + h.soltar);
    // Agrupar un objeto arma una carpeta; desagrupar un objeto no se puede.
    W().seleccionarItem(objeto.id);
    D().getElementById('botonDesagrupar').click();
    afirmar(W().ubicar(objeto.id) && W().esObjeto(W().ubicar(objeto.id).item) && D().getElementById('avisoCapas').textContent === W().t('aviso-objeto-no-desagrupa'), 'desagrupar un objeto: ' + D().getElementById('avisoCapas').textContent);
    D().getElementById('botonAgrupar').click();
    afirmar(W().itemActivo().clase === 'carpeta' && W().itemActivo().hijos[0] === objeto, 'agrupar un objeto no armó una carpeta');
    return Object.entries(h).map(([k, v]) => k + ' ' + v).join(' · ');
  });

  // Deshacer y rehacer (2026-10-07, assets/js/historial.js) en los tres
  // editores: botones, teclado y toque con dos dedos; un deslizador es un paso
  // al soltarlo; los lienzos del pincel y del borrador vuelven como estaban.
  await prueba('deshacer y rehacer (editor avanzado, de color y collage)', async () => {
    const h = {};
    const mitades = async () => {
      const c = D().createElement('canvas'); c.width = 100; c.height = 50;
      const cx = c.getContext('2d'); cx.fillStyle = 'rgb(40,40,40)'; cx.fillRect(0, 0, 50, 50); cx.fillStyle = 'rgb(220,220,220)'; cx.fillRect(50, 0, 50, 50);
      return new (W().File)([await new Promise((r) => c.toBlob(r, 'image/png'))], 'mitades.png', { type: 'image/png' });
    };
    const el = (id) => D().getElementById(id);
    const deslizar = (id, v) => { const e = el(id); e.value = v; e.dispatchEvent(new (W().Event)('input', { bubbles: true })); e.dispatchEvent(new (W().Event)('change', { bubbles: true })); };
    const tecla = (shift) => D().body.dispatchEvent(new (W().KeyboardEvent)('keydown', { key: shift ? 'Z' : 'z', ctrlKey: true, shiftKey: !!shift, bubbles: true }));
    const cerca = (v, e) => Math.abs(v - e) <= 3;

    // --- Editor avanzado ---
    await abrir('/herramientas/editor-avanzado-imagen.html'); await esperar(300);
    el('entradaArchivo').files = archivos([await mitades()]).files;
    el('entradaArchivo').dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => W().eval('imagenActual') && W().eval('historial').cantidad() === 1, 15000), 'el editor no cargó la imagen o el historial no arrancó');
    const lienzo = el('canvasTrabajo');
    lienzo.setPointerCapture = () => {};
    const leer = async (y = 0.5) => { await esperar(300); const d = lienzo.getContext('2d').getImageData(0, 0, lienzo.width, lienzo.height).data; return [0.25, 0.75].map((x) => { const i = (Math.round(lienzo.height * y) * lienzo.width + Math.round(lienzo.width * x)) * 4; return d[i + 3] < 10 ? 'T' : d[i]; }); };
    const puntero = (tipo, x, id = 1, tipoPuntero = 'mouse') => { const r = lienzo.getBoundingClientRect(); lienzo.dispatchEvent(new (W().PointerEvent)(tipo, { bubbles: true, pointerId: id, pointerType: tipoPuntero, isPrimary: id === 1, buttons: tipo === 'pointerup' ? 0 : 1, clientX: r.left + x * r.width, clientY: r.top + r.height / 2 })); };
    const trazo = (a, b) => { puntero('pointerdown', a); puntero('pointermove', (a + b) / 2); puntero('pointermove', b); puntero('pointerup', b); };
    afirmar(el('botonDeshacer').disabled, 'Deshacer habilitado sin haber hecho nada');

    deslizar('exposicion', 60);
    h.claro = await leer();
    afirmar(h.claro[0] > 60 && !el('botonDeshacer').disabled, 'exposición: ' + h.claro + ', deshacer ' + (el('botonDeshacer').disabled ? 'apagado' : 'encendido'));
    el('botonDeshacer').click();
    h.deshecho = await leer();
    afirmar(cerca(h.deshecho[0], 40) && cerca(h.deshecho[1], 220) && Number(el('exposicion').value) === 0, 'deshacer la exposición: ' + h.deshecho + ' (deslizador en ' + el('exposicion').value + ')');
    el('botonRehacer').click();
    h.rehecho = await leer();
    afirmar(h.rehecho[0] === h.claro[0] && Number(el('exposicion').value) === 60, 'rehacer la exposición: ' + h.rehecho);

    // Agregar un texto y deshacerlo con el teclado.
    el('botonNuevoTexto').click();
    const objetos = () => W().todosLosItems().filter((i) => W().esObjeto(i)).length;
    afirmar(objetos() === 1, '+ Texto no agregó un objeto');
    await esperar(150);
    tecla(false);
    h.sinTexto = objetos();
    afirmar(h.sinTexto === 0, 'Ctrl+Z no sacó el texto: ' + h.sinTexto);
    tecla(true);
    h.conTexto = objetos();
    afirmar(h.conTexto === 1, 'Ctrl+Shift+Z no devolvió el texto');
    tecla(false);

    // Pincel de la zona: elegirlo con la zona de partida deja solo lo
    // pintado; un trazo a la izquierda aclara solo ahí.
    W().seleccionarItem(W().eval('capas')[0].id);
    D().querySelector('#herramientasLienzo [data-herramienta="pincel"]').click();
    h.nada = await leer();
    afirmar(cerca(h.nada[0], 40), 'con solo un pincel sin pintar: ' + h.nada);
    trazo(0.1, 0.4);
    h.pintado = await leer();
    afirmar(h.pintado[0] > 60 && cerca(h.pintado[1], 220), 'trazo de pincel: ' + h.pintado);
    el('botonDeshacer').click();
    h.trazoDeshecho = await leer();
    afirmar(cerca(h.trazoDeshecho[0], 40), 'deshacer el trazo: ' + h.trazoDeshecho);
    el('botonRehacer').click();
    h.trazoRehecho = await leer();
    afirmar(h.trazoRehecho[0] === h.pintado[0], 'rehacer el trazo: ' + h.trazoRehecho);

    // Un pellizco que empieza como trazo no deja un paso vacío.
    const pasos = W().eval('historial').cantidad();
    // (en píxeles de pantalla fijos: el pellizco cambia el zoom y el tamaño del lienzo)
    const r0 = lienzo.getBoundingClientRect(), y0 = r0.top + r0.height / 2;
    const dedo = (tipo, x, id) => lienzo.dispatchEvent(new (W().PointerEvent)(tipo, { bubbles: true, pointerId: id, pointerType: 'touch', isPrimary: id === 1, buttons: tipo === 'pointerup' ? 0 : 1, clientX: r0.left + x, clientY: y0 }));
    dedo('pointerdown', 120, 1); dedo('pointerdown', 160, 2);
    dedo('pointermove', 90, 1); dedo('pointermove', 200, 2);
    dedo('pointerup', 90, 1); dedo('pointerup', 200, 2);
    await esperar(200);
    h.pasosTrasPellizco = W().eval('historial').cantidad() - pasos;
    afirmar(h.pasosTrasPellizco === 0, 'un pellizco dejó ' + h.pasosTrasPellizco + ' paso(s)');
    // Tocar con dos dedos (sin moverlos) deshace el trazo.
    const diag = () => [W().eval('historial').cantidad(), W().eval('historial').puedeRehacer(), W().eval('punterosActivos').size, W().eval('pintando')].join('/');
    const d0 = diag();
    puntero('pointerdown', 0.6, 3, 'touch'); const d1 = diag(); puntero('pointerdown', 0.7, 4, 'touch'); const d2 = diag();
    puntero('pointerup', 0.6, 3, 'touch'); puntero('pointerup', 0.7, 4, 'touch');
    h.dosDedos = await leer();
    afirmar(cerca(h.dosDedos[0], 40), 'el toque con dos dedos no deshizo: ' + h.dosDedos + ' · ' + [d0, d1, d2, diag()].join(' → '));

    // Borrar la foto y deshacerlo.
    agregarPieza('toda');
    D().querySelector('#herramientasLienzo [data-herramienta="borrar"]').click();
    trazo(0.2, 0.3);
    h.borrada = await leer();
    afirmar(h.borrada[0] === 'T', 'el borrador no borró: ' + h.borrada);
    el('botonDeshacer').click();
    h.borradoDeshecho = await leer();
    afirmar(h.borradoDeshecho[0] !== 'T', 'deshacer el borrado: ' + h.borradoDeshecho);

    // --- Editor de color ---
    await abrir('/herramientas/Editor-de-color-de-imagen.html'); await esperar(300);
    el('entradaArchivo').files = archivos([await mitades()]).files;
    el('entradaArchivo').dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => W().eval('imagenActual') && W().eval('historial').cantidad() === 1, 15000), 'el editor de color no cargó la imagen');
    const previa = el('vistaPrevia');
    const leerColor = async () => { await esperar(300); return previa.getContext('2d').getImageData(Math.round(previa.width * 0.25), previa.height >> 1, 1, 1).data[0]; };
    deslizar('exposicion', 60);
    h.colorClaro = await leerColor();
    D().querySelector('.presets button[data-preset]:not([data-preset="ninguno"])').click();
    const filtro = D().querySelector('.preset-miniatura.activo').dataset.preset;
    await esperar(150);
    el('botonDeshacer').click();
    afirmar(!D().querySelector('.preset-miniatura.activo') || D().querySelector('.preset-miniatura.activo').dataset.preset !== filtro, 'deshacer el filtro no lo quitó');
    h.colorSinFiltro = await leerColor();
    afirmar(h.colorSinFiltro === h.colorClaro && Number(el('exposicion').value) === 60, 'deshacer el filtro: ' + h.colorSinFiltro + ' (antes ' + h.colorClaro + ')');
    el('botonDeshacer').click();
    h.colorOriginal = await leerColor();
    afirmar(cerca(h.colorOriginal, 40), 'deshacer la exposición del editor de color: ' + h.colorOriginal);

    // --- Collage ---
    await abrir('/herramientas/collage-imagenes.html'); await esperar(300);
    el('entradaArchivo').files = archivos([await mitades(), await mitades()]).files;
    el('entradaArchivo').dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => W().eval('fotos').length === 2 && W().eval('historial').cantidad() === 1, 15000), 'el collage no cargó las fotos');
    const disenoActivo = () => [...D().querySelectorAll('#listaDisenos .boton-diseno')].findIndex((b) => b.classList.contains('activo'));
    D().querySelectorAll('#listaDisenos .boton-diseno')[1].click();
    await esperar(150);
    deslizar('espacio', 40);
    await esperar(150);
    afirmar(disenoActivo() === 1, 'no se eligió el segundo diseño');
    el('botonDeshacer').click();
    h.collageBorde = el('espacio').value;
    afirmar(disenoActivo() === 1 && h.collageBorde !== '40', 'deshacer el borde del collage: ' + h.collageBorde);
    el('botonDeshacer').click();
    h.collageDiseno = disenoActivo();
    afirmar(h.collageDiseno === 0, 'deshacer el diseño del collage: ' + h.collageDiseno);
    el('botonRehacer').click();
    afirmar(disenoActivo() === 1, 'rehacer el diseño del collage');
    return Object.entries(h).map(([k, v]) => k + ' ' + v).join(' · ');
  });

  // Pantalla completa y atajos de teclado (2026-10-07, pantalla-completa.js y
  // atajos.js). Sin una persona de verdad el navegador no deja entrar a
  // pantalla completa: se simula su aviso (fullscreenchange) y se comprueba
  // lo que hace el editor: el anuncio entra a su franja y vuelve a su lugar,
  // la burbuja "?" se cuelga adentro y el botón cambia de estado.
  await prueba('pantalla completa y atajos (editor avanzado, de color y collage)', async () => {
    const h = {};
    const el = (id) => D().getElementById(id);
    const archivo = async (n = 1) => {
      const c = D().createElement('canvas'); c.width = 120; c.height = 80;
      const cx = c.getContext('2d'); cx.fillStyle = 'rgb(40,40,40)'; cx.fillRect(0, 0, 60, 80); cx.fillStyle = 'rgb(220,220,220)'; cx.fillRect(60, 0, 60, 80);
      const b = await new Promise((r) => c.toBlob(r, 'image/png'));
      return Array.from({ length: n }, (_, i) => new (W().File)([b], 'f' + i + '.png', { type: 'image/png' }));
    };
    const fingir = (elemento) => {
      Object.defineProperty(D(), 'fullscreenElement', { configurable: true, get: () => elemento });
      D().dispatchEvent(new (W().Event)('fullscreenchange'));
    };
    const tecla = (key, tipo = 'keydown', extra = {}) => (extra.en || D().body).dispatchEvent(new (W().KeyboardEvent)(tipo, { key, bubbles: true, ...extra }));
    const revisarPantallaCompleta = async (contenedor, nombre) => {
      const boton = el('botonPantallaCompleta');
      afirmar(boton && !boton.hidden, nombre + ': no aparece el botón de pantalla completa');
      const anuncio = D().querySelector('.anuncio'), padre = anuncio.parentNode, siguiente = anuncio.nextSibling;
      fingir(contenedor);
      await esperar(100);
      // En PC el esqueleto ya tiene el anuncio adentro, en su franja.
      const franja = contenedor.querySelector('.franja-anuncio-completa, .ed-anuncio-fila');
      afirmar(contenedor.classList.contains('pantalla-completa-activa'), nombre + ': el contenedor no pasó a pantalla completa');
      afirmar(franja && franja.contains(anuncio) && franja.parentNode === contenedor, nombre + ': el anuncio no entró a su franja');
      afirmar(boton.getAttribute('aria-pressed') === 'true' && boton.title === W().t('title-salir-pantalla-completa'), nombre + ': el botón no dice "salir": ' + boton.title);
      const alto = franja.querySelector('.anuncio').getBoundingClientRect().height;
      fingir(null);
      await esperar(100);
      afirmar(anuncio.parentNode === padre && anuncio.nextSibling === siguiente && !contenedor.querySelector('.franja-anuncio-completa'), nombre + ': el anuncio no volvió a su lugar');
      afirmar(!contenedor.classList.contains('pantalla-completa-activa') && boton.getAttribute('aria-pressed') === 'false', nombre + ': no salió de pantalla completa');
      return 'anuncio ' + Math.round(alto) + ' px';
    };

    // --- Editor avanzado ---
    await abrir('/herramientas/editor-avanzado-imagen.html'); await esperar(300);
    el('entradaArchivo').files = archivos(await archivo()).files;
    el('entradaArchivo').dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => W().eval('imagenActual'), 15000), 'el editor avanzado no cargó la imagen');
    await esperar(300);
    const layout = el('layoutEditor');
    h.avanzado = await revisarPantallaCompleta(layout, 'avanzado');
    // La burbuja "?" se cuelga del elemento en pantalla completa.
    fingir(layout);
    const tip = D().querySelector('.ayuda-tip'), burbuja = tip.querySelector('.ayuda-burbuja');
    tip.querySelector('.ayuda-boton').dispatchEvent(new (W().FocusEvent)('focus'));
    afirmar(layout.contains(burbuja), 'la burbuja "?" quedó fuera de la pantalla completa');
    tip.querySelector('.ayuda-boton').dispatchEvent(new (W().FocusEvent)('blur'));
    fingir(null);
    await esperar(100);

    // Atajos: herramientas, espacio mantenido, "\" mantenido, Ctrl+0.
    const herramienta = () => W().eval('herramientaLienzo');
    tecla('e'); h.e = herramienta();
    afirmar(h.e === 'borrar', 'E no eligió el borrador: ' + h.e);
    tecla('b'); h.b = herramienta();
    afirmar(h.b === 'pincel', 'B no eligió el pincel: ' + h.b);
    tecla(' '); h.espacio = herramienta();
    tecla(' ', 'keyup'); h.trasEspacio = herramienta();
    afirmar(h.espacio === 'mover' && h.trasEspacio === 'pincel', 'espacio mantenido: ' + h.espacio + ' -> ' + h.trasEspacio);
    tecla('g'); afirmar(herramienta() === 'pincel', 'G eligió el degradado sin una pieza de degradado');
    tecla('\\'); h.original = W().eval('mostrandoOriginal');
    tecla('\\', 'keyup'); h.trasOriginal = W().eval('mostrandoOriginal');
    afirmar(h.original === true && h.trasOriginal === false, '"\\" mantenido: ' + h.original + ' -> ' + h.trasOriginal);
    W().aplicarZoom(2.5); W().eval('zoomAutomatico = false');
    tecla('0', 'keydown', { ctrlKey: true });
    afirmar(W().eval('zoomAutomatico') === true, 'Ctrl+0 no encajó la imagen');
    // Escribiendo, las letras son del texto.
    el('botonNuevoTexto').click();
    tecla('e', 'keydown', { en: el('textoCapa') });
    afirmar(herramienta() !== 'borrar', 'escribir una "e" en el texto eligió el borrador');

    // --- Editor de color ---
    await abrir('/herramientas/Editor-de-color-de-imagen.html'); await esperar(300);
    el('entradaArchivo').files = archivos(await archivo()).files;
    el('entradaArchivo').dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => W().eval('imagenActual'), 15000), 'el editor de color no cargó la imagen');
    await esperar(300);
    h.color = await revisarPantallaCompleta(el('layoutEditor'), 'de color');
    tecla('\\'); const comparando = W().eval('comparando'); tecla('\\', 'keyup');
    afirmar(comparando && !W().eval('comparando'), 'el editor de color no mostró el original con "\\"');

    // --- Collage ---
    await abrir('/herramientas/collage-imagenes.html'); await esperar(300);
    el('entradaArchivo').files = archivos(await archivo(2)).files;
    el('entradaArchivo').dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => W().eval('fotos').length === 2, 15000), 'el collage no cargó las fotos');
    await esperar(300);
    h.collage = await revisarPantallaCompleta(el('editor'), 'collage');
    return Object.entries(h).map(([k, v]) => k + ' ' + v).join(' · ');
  });

  // Ajustes del motor agregados el 2026-10-04 (Temperatura, Matiz, Blancos,
  // Negros, Nitidez, Viñeta), en los dos editores. Imagen: mitad gris 40 y
  // mitad gris 220, así cada uno deja una huella: el balance separa los
  // canales, Blancos mueve solo lo claro, Negros solo lo oscuro, la nitidez
  // agranda el salto en el borde y la viñeta oscurece la esquina y no el centro.
  for (const [pagina, idLienzo] of [['editor-avanzado-imagen', 'canvasTrabajo'], ['Editor-de-color-de-imagen', 'vistaPrevia']]) {
    await prueba('ajustes nuevos del motor (' + pagina + ')', async () => {
      await abrir('/herramientas/' + pagina + '.html'); await esperar(300);
      const base = D().createElement('canvas'); base.width = 200; base.height = 100;
      const cb = base.getContext('2d');
      cb.fillStyle = 'rgb(40,40,40)'; cb.fillRect(0, 0, 100, 100);
      cb.fillStyle = 'rgb(220,220,220)'; cb.fillRect(100, 0, 100, 100);
      const png = await new Promise((r) => base.toBlob(r, 'image/png'));
      const entrada = D().getElementById('entradaArchivo');
      entrada.files = archivos([new (W().File)([png], 'mitades.png', { type: 'image/png' })]).files;
      entrada.dispatchEvent(new (W().Event)('change'));
      const lienzo = await hasta(() => { const c = D().getElementById(idLienzo); return c && c.width >= 100 && c; }, 15000);
      afirmar(lienzo, 'la imagen no se cargó');
      const px = (x, y = 0.5) => [...lienzo.getContext('2d').getImageData(Math.round((lienzo.width - 1) * x), Math.round((lienzo.height - 1) * y), 1, 1).data].slice(0, 3);
      const poner = async (id, v) => { const el = D().getElementById(id); el.value = v; el.dispatchEvent(new (W().Event)('input')); await esperar(250); };
      const h = {};
      afirmar(await hasta(() => px(0.75)[0] > 0, 10000), 'la vista previa no se dibujó');
      await esperar(300);
      h.neutro = [px(0.25)[0], px(0.75)[0]];
      afirmar(h.neutro[0] === 40 && h.neutro[1] === 220, 'sin ajustes no es la imagen: ' + h.neutro);

      await poner('temperatura', 100);
      const calido = px(0.25);
      afirmar(calido[0] > 45 && calido[2] < 35, 'Temperatura +100 no calentó (R sube, B baja): ' + calido);
      h.temperatura = calido.join('/');
      await poner('temperatura', 0);
      await poner('matizBlancos', 100);
      const magenta = px(0.25);
      afirmar(magenta[1] < 36 && magenta[0] > magenta[1], 'Matiz +100 no fue a magenta (G baja): ' + magenta);
      h.matiz = magenta.join('/');
      await poner('matizBlancos', 0);

      await poner('blancos', 100);
      h.blancos = [px(0.25)[0], px(0.75)[0]];
      afirmar(h.blancos[0] === 40 && h.blancos[1] > 235, 'Blancos +100 (solo lo claro sube): ' + h.blancos);
      await poner('blancos', 0);
      await poner('negros', -100);
      h.negros = [px(0.25)[0], px(0.75)[0]];
      afirmar(h.negros[0] < 25 && h.negros[1] === 220, 'Negros -100 (solo lo oscuro baja): ' + h.negros);
      await poner('negros', 0);

      // Nitidez: junto al borde, lo oscuro se oscurece y lo claro se aclara.
      const borde = 0.5 - 1.5 / lienzo.width, despues = 0.5 + 1.5 / lienzo.width;
      const antesNitidez = [px(borde)[0], px(despues)[0]];
      await poner('nitidez', 100);
      h.nitidez = antesNitidez + ' -> ' + [px(borde)[0], px(despues)[0]];
      afirmar(px(borde)[0] < antesNitidez[0] && px(despues)[0] > antesNitidez[1], 'Nitidez no agrandó el salto del borde: ' + h.nitidez);
      await poner('nitidez', 0);

      await poner('vineta', -100);
      h.vineta = [px(0.02, 0.02)[0], px(0.25)[0]];
      afirmar(h.vineta[0] < 15 && h.vineta[1] >= 36, 'Viñeta -100 (esquina oscura, centro de la mitad casi igual): ' + h.vineta);
      await poner('vineta', 0);
      h.vuelta = [px(0.25)[0], px(0.75)[0]];
      afirmar(h.vuelta[0] === 40 && h.vuelta[1] === 220, 'con todo en cero no volvió a la imagen: ' + h.vuelta);
      return Object.entries(h).map(([k, v]) => k + ' ' + v).join(' · ');
    });
  }

  // Motor de filtros por GPU (filtros-gpu.js, 2026-10-05) contra el bucle en
  // JavaScript, que es el respaldo y la referencia: tienen que dar los mismos
  // píxeles (±2 por redondeo). Un cambio al pipeline en uno solo de los dos
  // lo rompe acá. Imagen opaca con todos los tonos y luces. "Por color" y la
  // mezcla B/N deciden el color de cada píxel mirando su entorno desenfocado
  // (lo desenfoca el navegador), y en unos pocos píxeles al borde de esa
  // decisión pueden diferir unos niveles: ahí se mira que sean muy pocos.
  // Además comprueba que la GPU de verdad corrió (si no, compararía JS con JS).
  await prueba('motor de filtros: GPU igual que JavaScript', async () => {
    await abrir('/herramientas/editor-avanzado-imagen.html'); await esperar(300);
    const PF = W().ProcesarFiltros, GPU = W().FiltrosGPU;
    afirmar(GPU && GPU.disponible(10, 10), 'sin WebGL2: la página usaría solo el motor en JavaScript');
    const fuente = D().createElement('canvas'); fuente.width = 360; fuente.height = 120;
    const cf = fuente.getContext('2d');
    for (let x = 0; x < 360; x++) for (let y = 0; y < 120; y++) { cf.fillStyle = 'hsl(' + x + ',' + (20 + (y % 40) * 2) + '%,' + (4 + y * 0.77) + '%)'; cf.fillRect(x, y, 1, 1); }
    let pasadas = 0;
    const procesar = GPU.procesar;
    GPU.procesar = function () { const r = procesar.apply(this, arguments); if (r) pasadas += 1; return r; };
    const correr = (caso, porGpu) => {
      const c = PF.copiarCanvas(fuente);
      if (!porGpu) W().FiltrosGPU = undefined;
      try { PF.renderizarFiltro(c, { ...PF.AJUSTES_NEUTROS, ...caso.ajustes }, caso.capas || [], PF.SIN_TEXTURA, 1); } finally { W().FiltrosGPU = GPU; }
      return c.getContext('2d').getImageData(0, 0, c.width, c.height).data;
    };
    const porColor = { ...W().eval('AjustePorColor').neutro(), rojos: { tono: 40, saturacion: 50, luminancia: -30 }, azules: { tono: -60, saturacion: -80, luminancia: 40 } };
    const CASOS = [
      ['exposición', { exposicion: 60 }], ['contraste', { contraste: -40 }], ['luces', { sombras: 60, reflejos: -50, blancos: 40, negros: -40 }],
      ['balance', { temperatura: 60, matizBlancos: -40 }], ['saturación', { saturacion: 80, intensidad: 50 }], ['tono', { tono: 120 }],
      ['colores', { colorRojo: 40, colorAzul: -30 }], ['b/n', { bn: true, canalRojo: 150 }], ['sepia', { sepia: 70 }], ['invertir', { invertir: 40 }],
      ['claridad', { claridad: 60 }], ['nitidez', { nitidez: 70 }], ['viñeta', { vineta: -80 }],
      ['dispersión', { dispersionIntensidad: 70, dispersionMatiz: 60 }], ['aberración', { aberracionIntensidad: 80, aberracionAngulo: 30, aberracionMatiz: 40 }],
      ['intensidad del filtro', { exposicion: 40, intensidadFiltro: 40 }],
    ].map(([nombre, ajustes]) => ({ nombre, ajustes }));
    CASOS.push({ nombre: 'filtro vintage (capas extra)', ajustes: { ...PF.PRESETS.vintage, grano: 0 }, capas: PF.PRESETS.vintage.capasExtra });
    CASOS.push({ nombre: 'por color', ajustes: { hsl: porColor }, sensible: true });
    CASOS.push({ nombre: 'mezcla b/n', ajustes: { bn: true, mezclaAzules: -60, mezclaNaranjas: 50 }, sensible: true });
    const malos = [];
    let maxNormal = 0;
    for (const caso of CASOS) {
      const g = correr(caso, true), j = correr(caso, false);
      let max = 0, fuera = 0;
      for (let i = 0; i < g.length; i++) { const d = Math.abs(g[i] - j[i]); if (d > max) max = d; if (d > 2) fuera += 1; }
      if (!caso.sensible) maxNormal = Math.max(maxNormal, max);
      if (caso.sensible ? fuera > g.length * 0.001 : max > 2) malos.push(caso.nombre + ': máx ' + max + ', ' + fuera + ' valores fuera de ±2');
    }
    GPU.procesar = procesar;
    afirmar(malos.length === 0, 'la GPU y JavaScript no dan lo mismo — ' + malos.join(' · '));
    afirmar(pasadas >= CASOS.length, 'la GPU no corrió (pasadas: ' + pasadas + ')');
    return CASOS.length + ' casos · diferencia máxima ' + maxNormal + ' · ' + pasadas + ' pasadas por GPU';
  });

  // "Ajustar por color" (ajuste-por-color.js): hasta el 2026-10-05 el
  // escuchador de cambio de idioma quedaba adentro del "input" de cada
  // deslizador: el panel no se traducía hasta mover uno y después se sumaba
  // un escuchador por movimiento.
  await prueba('ajustar por color: idioma sin tocar nada y sin escuchadores de más', async () => {
    await abrir('/herramientas/editor-avanzado-imagen.html'); await esperar(300);
    const entrada = D().getElementById('entradaArchivo');
    entrada.files = archivos([await imagenDePrueba('A', 400, 300)]).files;
    entrada.dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => W().eval('imagenActual'), 15000), 'la imagen no se cargó');
    const panel = D().getElementById('panelAjustePorColor');
    const subtitulo = () => panel.querySelector('.ajuste-color-subtitulo').textContent;
    // Cambio de idioma en el lugar (el botón navega a /en/ cuando existe:
    // acá se dispara el mismo evento que aplicarIdioma de idioma.js).
    const idiomaReal = W().idiomaActual;
    const cambiarA = (idioma) => { W().idiomaActual = () => idioma; D().dispatchEvent(new (W().CustomEvent)('cambioIdioma', { detail: { idioma } })); };
    const inicial = subtitulo();
    cambiarA('en');
    const cambiado = subtitulo();
    cambiarA('es');
    W().idiomaActual = idiomaReal;
    afirmar(cambiado === 'Adjust by color' && subtitulo() === 'Ajustar por color', 'el panel no cambió de idioma sin tocar un deslizador: ' + inicial + ' -> ' + cambiado + ' -> ' + subtitulo());
    let nuevos = 0;
    const agregar = D().addEventListener;
    D().addEventListener = function (tipo) { if (tipo === 'cambioIdioma') nuevos += 1; return agregar.apply(this, arguments); };
    const tono = panel.querySelector('input[data-ajuste="tono"]');
    for (let i = 1; i <= 5; i++) { tono.value = i * 10; tono.dispatchEvent(new (W().Event)('input')); }
    D().addEventListener = agregar;
    afirmar(nuevos === 0, 'mover el deslizador sumó ' + nuevos + ' escuchadores de idioma');
    afirmar(W().capaActiva().ajustes.hsl.rojos.tono === 50, 'el deslizador no llegó a la capa: ' + JSON.stringify(W().capaActiva().ajustes.hsl.rojos));
    return inicial + ' ⇄ ' + cambiado + ' · 5 movimientos, 0 escuchadores nuevos';
  });

  // Desde el motor por GPU, "Original" también se puede elegir en celular,
  // si el dispositivo puede (canvas de ese tamaño y la GPU lo acepta).
  await prueba('resolución media y original en celular (con GPU)', async () => {
    marco.style.width = '390px'; marco.style.height = '844px';
    try {
      await abrir('/herramientas/editor-avanzado-imagen.html'); await esperar(300);
      const entrada = D().getElementById('entradaArchivo');
      entrada.files = archivos([await imagenDePrueba('A', 3000, 2000)]).files;
      entrada.dispatchEvent(new (W().Event)('change'));
      afirmar(await hasta(() => W().eval('imagenActual'), 15000), 'la imagen no se cargó');
      const selector = D().getElementById('resolucionVistaPrevia');
      const opcion = D().getElementById('opcionResolucionOriginal');
      afirmar(selector.value === 'rapida', 'en celular no arrancó en Rápida: ' + selector.value);
      afirmar(!opcion.disabled, 'Original sigue desactivada en celular: ' + D().getElementById('avisoResolucionOriginal').textContent);
      const lienzo = D().getElementById('canvasTrabajo');
      afirmar(!D().getElementById('opcionResolucionMedia').disabled, 'Media desactivada en celular');
      selector.value = 'media'; selector.dispatchEvent(new (W().Event)('change'));
      afirmar(await hasta(() => lienzo.width === 2000, 10000), 'Media no pasó a 2000 px: ' + lienzo.width);
      const anchoMedia = lienzo.width + '×' + lienzo.height;
      selector.value = 'original'; selector.dispatchEvent(new (W().Event)('change'));
      afirmar(await hasta(() => lienzo.width === 3000, 10000), 'no pasó a 3000 px: ' + lienzo.width);
      const anchoOriginal = lienzo.width + '×' + lienzo.height;
      const aviso = D().getElementById('avisoResolucionOriginal').textContent;
      afirmar(aviso === W().t('aviso-original-memoria'), 'aviso inesperado: ' + aviso);
      // Un celular cuya placa admite menos que la foto (2500 px): con Original
      // elegida, al abrir otra foto baja sola a Media y dice por qué.
      const GPU = W().FiltrosGPU, ladoReal = GPU.ladoMaximo;
      GPU.ladoMaximo = () => 2500;
      try {
        entrada.files = archivos([await imagenDePrueba('B', 3000, 2000)]).files;
        entrada.dispatchEvent(new (W().Event)('change'));
        afirmar(await hasta(() => lienzo.width === 2000, 10000), 'con placa de 2500 px no bajó a Media: ' + lienzo.width + ' (' + selector.value + ')');
        const motivo = D().getElementById('avisoResolucionOriginal').textContent;
        afirmar(opcion.disabled && selector.value === 'media' && motivo.includes('2500'), 'Original sin desactivar o sin motivo: ' + motivo);
        return 'Rápida → Media ' + anchoMedia + ' → Original ' + anchoOriginal + ' · "' + aviso + '" · placa de 2500 px: baja a Media y avisa';
      } finally {
        GPU.ladoMaximo = ladoReal;
      }
    } finally {
      marco.style.width = ''; marco.style.height = '';
    }
  });

  // Máscaras de primer plano / fondo. Sin internet no se puede bajar el
  // modelo, así que el recorte de la IA se reemplaza por uno falso: dos
  // cuadrados separados (dos "objetos") sobre una foto gris 128. La capa en
  // Multiplicar baja a 64 lo que la máscara marca. Se comprueba aparte que el
  // módulo real carga desde la página.
  await prueba('editor avanzado: primer plano, fondo y elegir cualquier objeto', async () => {
    await abrir('/herramientas/editor-avanzado-imagen.html'); await esperar(300);
    const mod = await W().eval("import(new URL('recorte-sujeto.js', document.querySelector('script[src*=mascaras-capa]').src).href)");
    afirmar(typeof mod.recortarSujeto === 'function', 'el módulo recorte-sujeto.js no cargó desde la página');
    const modSam = await W().eval("import(new URL('objeto-por-toque.js', document.querySelector('script[src*=mascaras-capa]').src).href)");
    afirmar(typeof modSam.analizarImagen === 'function' && typeof modSam.objetoEn === 'function', 'el módulo objeto-por-toque.js no cargó desde la página');
    const base = D().createElement('canvas'); base.width = 200; base.height = 100;
    base.getContext('2d').fillStyle = 'rgb(128,128,128)'; base.getContext('2d').fillRect(0, 0, 200, 100);
    let llamadas = 0;
    W().calcularSujeto = async () => {
      llamadas++;
      const s = D().createElement('canvas'); s.width = 200; s.height = 100;
      const cs = s.getContext('2d'); cs.fillStyle = '#fff';
      cs.fillRect(20, 30, 40, 40); cs.fillRect(140, 30, 40, 40);
      return s;
    };
    // SlimSAM falso: tocar dentro de un cuadrado da ese cuadrado; tocar en
    // cualquier otro lado da una "flor" de 20x20 alrededor del punto. Tarda
    // un poco a propósito, para ver el velo de "la IA está trabajando".
    let analisis = 0;
    W().analizarParaObjetos = async () => {
      analisis++;
      await esperar(400);
      return {
        objetoEn: async (x, y) => {
          await esperar(200);
          const s = D().createElement('canvas'); s.width = 200; s.height = 100;
          const cs = s.getContext('2d'); cs.fillStyle = '#fff';
          const px_ = x * 200;
          if (px_ >= 20 && px_ < 60) cs.fillRect(20, 30, 40, 40);
          else if (px_ >= 140 && px_ < 180) cs.fillRect(140, 30, 40, 40);
          else cs.fillRect(px_ - 10, 40, 20, 20);
          return s;
        },
      };
    };
    const png = await new Promise((r) => base.toBlob(r, 'image/png'));
    const entrada = D().getElementById('entradaArchivo');
    entrada.files = archivos([new (W().File)([png], 'gris.png', { type: 'image/png' })]).files;
    entrada.dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => D().querySelector('#listaCapas .fila-opacidad select'), 15000), 'el editor no cargó la imagen');
    const modo = D().querySelector('#listaCapas .fila-opacidad select'); modo.value = 'multiply'; modo.dispatchEvent(new (W().Event)('change'));
    const lienzo = D().getElementById('canvasTrabajo');
    const px = (cv, x) => cv.getContext('2d').getImageData(Math.round(x * cv.width / 200), Math.round(cv.height / 2), 1, 1).data[0];
    // izquierda (cuadrado 1), medio (nada), derecha (cuadrado 2)
    const leer = async () => { await esperar(300); return [px(lienzo, 40), px(lienzo, 100), px(lienzo, 160)]; };
    const cerca = (a, b) => a.every((v, i) => Math.abs(v - b[i]) <= 2);
    const clic = (sel) => D().querySelector(sel).click();
    const h = {};

    agregarPieza('sujeto');
    afirmar(D().querySelector('#herramientasLienzo [data-herramienta="objeto"]').classList.contains('activo'), 'elegir Primer plano no dejó elegida la herramienta "Elegir objeto"');
    afirmar(await hasta(() => W().sujetoListo(W().eval('imagenActual')), 5000), 'el primer plano no se calculó');
    h.primerPlano = await leer();
    afirmar(cerca(h.primerPlano, [64, 128, 64]), 'primer plano: ' + h.primerPlano);
    // Quitar el fondo: con el borrador, "Borrar lo que está fuera de la zona"
    // deja transparente todo menos el primer plano (2026-10-08).
    clic('#herramientasLienzo [data-herramienta="borrar"]');
    afirmar(D().getElementById('borrarConZona').style.display !== 'none', 'con el borrador no aparecen los botones de borrar con la zona');
    clic('#botonBorrarFueraZona');
    await esperar(300);
    const alfa = (x) => lienzo.getContext('2d').getImageData(Math.round(x * lienzo.width / 200), Math.round(lienzo.height / 2), 1, 1).data[3];
    h.sinFondo = [alfa(40), alfa(100), alfa(160)].join(',');
    afirmar(alfa(40) > 250 && alfa(100) < 5 && alfa(160) > 250, 'borrar fuera del primer plano (alfa izquierda, medio, derecha): ' + h.sinFondo);
    clic('#botonRestaurarBorrado');
    await esperar(300);
    afirmar(alfa(100) > 250, 'restaurar no devolvió el fondo');
    clic('#herramientasLienzo [data-herramienta="objeto"]'); // de vuelta a Elegir objeto, como lo dejó el primer plano
    invertirPieza(true);
    h.fondo = await leer();
    afirmar(cerca(h.fondo, [128, 64, 128]), 'fondo: ' + h.fondo);
    invertirPieza(false);

    // Elegir objeto: tocar algo que NO es el primer plano (la "flor" del
    // medio) también lo encuentra; mientras tanto, el velo sobre la imagen.
    const tocar = (x) => {
      const r = lienzo.getBoundingClientRect();
      lienzo.dispatchEvent(new (W().MouseEvent)('click', { bubbles: true, clientX: r.left + (x / 200) * r.width, clientY: r.top + r.height / 2 }));
    };
    const velo = D().getElementById('trabajandoIA');
    tocar(100);
    await esperar(100);
    afirmar(!velo.hidden && velo.textContent.trim().length > 0, 'no apareció el velo de "la IA está trabajando"');
    h.velo = velo.textContent.trim();
    afirmar(await hasta(() => velo.hidden, 5000), 'el velo no se fue al terminar');
    h.flor = await leer();
    afirmar(cerca(h.flor, [128, 64, 128]), 'un objeto que no es el primer plano: ' + h.flor);
    tocar(160);
    afirmar(await hasta(() => velo.hidden, 5000), 'el velo no se fue al terminar');
    h.objeto = await leer();
    afirmar(cerca(h.objeto, [128, 128, 64]), 'solo el objeto derecho: ' + h.objeto);
    afirmar(W().piezaActiva().tipo === 'objeto', 'tocar no pasó la pieza a "Objeto"');
    invertirPieza(true);
    h.fondoMenosObjeto = await leer();
    afirmar(cerca(h.fondoMenosObjeto, [64, 64, 128]), 'fondo menos el objeto elegido: ' + h.fondoMenosObjeto);
    invertirPieza(false);

    // La descarga también.
    D().getElementById('botonDescargarAvanzado').click();
    const enlace = D().getElementById('enlaceDescarga');
    afirmar(await hasta(() => enlace.href.startsWith('blob:'), 15000), 'la descarga no se generó');
    const img = new (W().Image)(); img.src = enlace.href; await img.decode();
    const cd = D().createElement('canvas'); cd.width = img.width; cd.height = img.height;
    cd.getContext('2d').drawImage(img, 0, 0);
    h.descarga = [px(cd, 40), px(cd, 100), px(cd, 160)];
    afirmar(cerca(h.descarga, [128, 128, 64]), 'descarga con el objeto elegido: ' + h.descarga);

    // Copia de la foto con el objeto elegido: un recorte del tamaño del cuadrado.
    clic('#botonCopiarFoto');
    const copia = W().capaDeContenido(W().itemActivo());
    afirmar(copia && copia.contenido && copia.contenido.lienzo.width === 40 && copia.contenido.lienzo.height === 40, 'la copia del objeto no mide 40x40: ' + (copia.contenido && copia.contenido.lienzo.width));

    clic('#botonTodoSujeto');
    afirmar(llamadas === 1, 'el primer plano se calculó ' + llamadas + ' veces (tiene que ser una por imagen)');
    afirmar(analisis === 1, 'la foto se analizó ' + analisis + ' veces para elegir objetos (tiene que ser una)');

    // El generador de stickers comparte las islas: tiene que cargar el módulo.
    await abrir('/herramientas/generador-stickers.html'); await esperar(300);
    afirmar(W().MascarasCapa && typeof W().MascarasCapa.islas === 'function', 'el generador de stickers no cargó mascaras-capa.js');
    const islas = W().MascarasCapa.islas(new Uint8ClampedArray([0, 0, 0, 255, 0, 0, 0, 0, 0, 0, 0, 255]), 3, 1, 20);
    afirmar(islas.tamanos.length === 2 && islas.etiquetas[1] === -1, 'islas: ' + JSON.stringify(islas.tamanos));
    return Object.entries(h).map(([k, v]) => k + ' ' + v).join(' · ');
  });

  // Una imagen agregada como capa mide su máscara sobre SÍ MISMA (no sobre la
  // foto): primer plano, elegir objeto y color, y la máscara la sigue al
  // moverla. Foto gris 128 de 200x100; la capa, de 100x50, roja a la
  // izquierda y azul a la derecha, queda centrada (x 50..150). La IA es falsa:
  // el "primer plano" es lo rojo de lo que reciba, y tocar un objeto da la
  // mitad tocada. Anota el tamaño de lo que analiza, para ver que es la capa.
  await prueba('editor avanzado: máscara medida sobre una capa de imagen (IA, color, la sigue al moverla)', async () => {
    await abrir('/herramientas/editor-avanzado-imagen.html'); await esperar(300);
    const tamanos = [];
    const leerUrl = async (url) => { const i = new (W().Image)(); i.src = url; await i.decode(); tamanos.push(i.width + 'x' + i.height); return i; };
    W().calcularSujeto = async (url) => {
      const i = await leerUrl(url);
      const s = D().createElement('canvas'); s.width = i.width; s.height = i.height;
      const cs = s.getContext('2d', { willReadFrequently: true }); cs.drawImage(i, 0, 0);
      const d = cs.getImageData(0, 0, s.width, s.height);
      for (let k = 0; k < d.data.length; k += 4) {
        const rojo = d.data[k] > 200 && d.data[k + 2] < 50;
        d.data[k] = d.data[k + 1] = d.data[k + 2] = 255; d.data[k + 3] = rojo ? 255 : 0;
      }
      cs.putImageData(d, 0, 0);
      return s;
    };
    W().analizarParaObjetos = async (url) => {
      const i = await leerUrl(url);
      return { objetoEn: async (x) => {
        const s = D().createElement('canvas'); s.width = i.width; s.height = i.height;
        const cs = s.getContext('2d'); cs.fillStyle = '#fff';
        cs.fillRect(x < 0.5 ? 0 : i.width / 2, 0, i.width / 2, i.height);
        return s;
      } };
    };
    const lienzoPng = async (ancho, alto, dibujar, nombre) => {
      const c = D().createElement('canvas'); c.width = ancho; c.height = alto; dibujar(c.getContext('2d'));
      const b = await new Promise((r) => c.toBlob(r, 'image/png'));
      return new (W().File)([b], nombre, { type: 'image/png' });
    };
    const entrada = D().getElementById('entradaArchivo');
    entrada.files = archivos([await lienzoPng(200, 100, (cx) => { cx.fillStyle = 'rgb(128,128,128)'; cx.fillRect(0, 0, 200, 100); }, 'gris.png')]).files;
    entrada.dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => W().eval('imagenActual'), 15000), 'la imagen no se cargó');
    const entradaCapa = D().getElementById('entradaImagenCapa');
    entradaCapa.files = archivos([await lienzoPng(100, 50, (cx) => { cx.fillStyle = '#f00'; cx.fillRect(0, 0, 50, 50); cx.fillStyle = '#00f'; cx.fillRect(50, 0, 50, 50); }, 'rojo-azul.png')]).files;
    entradaCapa.dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => W().esObjeto(W().itemActivo()), 5000), 'la imagen no entró como objeto');
    const capa = W().itemActivo(); // el objeto: su zona dice qué parte de la imagen se ve
    const cont = W().capaDeContenido(capa);

    const lienzo = D().getElementById('canvasTrabajo');
    const color = (cv, x) => { const d = cv.getContext('2d').getImageData(Math.round(x * cv.width / 200), Math.round(cv.height / 2), 1, 1).data; return d[0] > 200 && d[2] < 60 ? 'rojo' : (d[2] > 200 && d[0] < 60 ? 'azul' : (Math.abs(d[0] - 128) <= 3 && Math.abs(d[2] - 128) <= 3 ? 'gris' : d[0] + ',' + d[1] + ',' + d[2])); };
    const leer = async (xs) => { await esperar(300); return xs.map((x) => color(lienzo, x)).join(' '); };
    const clic = (sel) => D().querySelector(sel).click();
    const tocar = (x) => {
      const r = lienzo.getBoundingClientRect();
      lienzo.dispatchEvent(new (W().MouseEvent)('click', { bubbles: true, clientX: r.left + (x / 200) * r.width, clientY: r.top + r.height / 2 }));
    };
    const velo = D().getElementById('trabajandoIA');
    const h = {};
    h.sinMascara = await leer([75, 125]);
    afirmar(h.sinMascara === 'rojo azul', 'la capa no se ve entera: ' + h.sinMascara);

    const sujeto = agregarPieza('sujeto');
    afirmar(sujeto.fuente === 'capa', 'una pieza nueva de un objeto no mide sobre su imagen: ' + sujeto.fuente);
    afirmar(D().getElementById('controlesFuenteMascara').style.display !== 'none', 'no aparece "Medir sobre" en una capa de imagen');
    afirmar(await hasta(() => W().sujetoListo(W().fuenteDeContenido(cont)), 5000), 'el primer plano de la capa no se calculó');
    h.primerPlano = await leer([75, 125]);
    afirmar(h.primerPlano === 'rojo gris', 'primer plano de la capa: ' + h.primerPlano);
    afirmar(tamanos[0] === '100x50', 'la IA no analizó la capa sino ' + tamanos[0]);

    // Sobre la foto (gris, sin rojo): no queda nada de la capa.
    clic('#modosFuenteMascara [data-fuente="foto"]');
    afirmar(await hasta(() => W().sujetoListo(W().eval('imagenActual')), 5000), 'el primer plano de la foto no se calculó');
    h.sobreLaFoto = await leer([75, 125]);
    afirmar(h.sobreLaFoto === 'gris gris', 'primer plano de la foto: ' + h.sobreLaFoto);
    clic('#modosFuenteMascara [data-fuente="capa"]');
    h.deVuelta = await leer([75, 125]);
    afirmar(h.deVuelta === 'rojo gris', 'al volver a "Esta capa": ' + h.deVuelta + ' (el primer plano de la capa tiene que quedar guardado)');

    // La máscara acompaña a la capa: movida a la izquierda (x 0..100).
    cont.contenido.centro = { x: 0.25, y: 0.5 }; W().programarRenderizado();
    h.movida = await leer([25, 75, 125]);
    afirmar(h.movida === 'rojo gris gris', 'la máscara no siguió a la capa: ' + h.movida);
    // Y volteada: el rojo pasa a la derecha y la máscara con él.
    cont.contenido.voltearH = true; W().programarRenderizado();
    h.volteada = await leer([25, 75]);
    afirmar(h.volteada === 'gris rojo', 'volteada: ' + h.volteada);
    cont.contenido.voltearH = false; cont.contenido.centro = { x: 0.5, y: 0.5 }; W().programarRenderizado();

    // Elegir objeto: tocar la mitad azul de la capa (x 125 de la foto = 0,75 de la capa).
    afirmar(D().querySelector('#herramientasLienzo [data-herramienta="objeto"]').classList.contains('activo'), '"Elegir objeto" no quedó elegida');
    tocar(125);
    await esperar(100);
    afirmar(await hasta(() => velo.hidden, 5000), 'el velo no se fue al terminar');
    h.objeto = await leer([75, 125]);
    afirmar(h.objeto === 'gris azul', 'objeto elegido en la capa: ' + h.objeto);
    afirmar(tamanos.filter((t) => t === '100x50').length === 2, 'elegir objeto no analizó la capa: ' + tamanos.join(', '));
    // Un toque fuera de la capa no elige nada y lo dice.
    tocar(10);
    afirmar(D().getElementById('estadoSujeto').textContent === W().t('objeto-fuera-de-capa'), 'tocar fuera de la capa: ' + D().getElementById('estadoSujeto').textContent);

    // Color: el gotero toma el color de la capa (rojo), no el gris de la foto.
    vaciarZona();
    const piezaColor = agregarPieza('color'); // deja elegido el gotero
    tocar(75);
    afirmar(piezaColor.color && piezaColor.color.r > 200 && piezaColor.color.b < 50, 'el gotero no tomó el rojo de la capa: ' + JSON.stringify(piezaColor.color));
    h.color = await leer([75, 125]);
    afirmar(h.color === 'rojo gris', 'color de la capa: ' + h.color);

    // La descarga a resolución completa, y el proyecto guarda "Medir sobre".
    D().getElementById('botonDescargarAvanzado').click();
    const enlace = D().getElementById('enlaceDescarga');
    afirmar(await hasta(() => enlace.href.startsWith('blob:'), 15000), 'la descarga no se generó');
    const img = new (W().Image)(); img.src = enlace.href; await img.decode();
    const cd = D().createElement('canvas'); cd.width = img.width; cd.height = img.height;
    cd.getContext('2d').drawImage(img, 0, 0);
    h.descarga = [75, 125].map((x) => color(cd, x)).join(' ');
    afirmar(h.descarga === 'rojo gris', 'descarga: ' + h.descarga);
    afirmar(W().itemAProyecto(capa).mascara.piezas[0].fuente === 'capa', 'el proyecto no guarda "Medir sobre"');
    afirmar(W().MascarasCapa.desdeV7({ base: 'color' }).piezas[0].fuente === 'foto', 'un proyecto viejo no sigue midiendo sobre la foto');
    // Quitar el fondo de la imagen agregada: con su primer plano (la mitad
    // roja) como zona, "Borrar lo que está fuera de la zona" borra la mitad
    // azul de la imagen misma: con la zona de vuelta en "toda", sigue sin
    // verse, y lo borrado la acompaña al moverla.
    vaciarZona(); agregarPieza('sujeto');
    afirmar(await hasta(() => W().sujetoListo(W().fuenteDeContenido(cont)), 5000), 'el primer plano de la capa no está');
    clic('#herramientasLienzo [data-herramienta="borrar"]');
    clic('#botonBorrarFueraZona');
    zonaSoloToda();
    h.sinFondoCapa = await leer([75, 125]);
    afirmar(h.sinFondoCapa === 'rojo gris', 'borrar fuera de la zona en la imagen agregada: ' + h.sinFondoCapa);
    cont.contenido.centro = { x: 0.25, y: 0.5 }; W().programarRenderizado();
    h.sinFondoMovida = await leer([25, 75, 125]);
    afirmar(h.sinFondoMovida === 'rojo gris gris', 'lo borrado no acompañó a la imagen: ' + h.sinFondoMovida);
    return Object.entries(h).map(([k, v]) => k + ' ' + v).join(' · ') + ' · la IA analizó ' + tamanos.join(', ');
  });

  // Esqueleto de PC (2026-10-08, plan fase 6, assets/js/esqueleto-editor.js)
  // en los tres editores: barra de íconos, una sola categoría a la vista, los
  // botones reales en la barra superior, el anuncio en su franja, teclas 1–9,
  // lo contextual (agregar un texto, una herramienta, tocar una foto), y a
  // 390 px el de celular (fase 7, con app-movil.js): la misma barra, los
  // mismos íconos abajo y la hoja; al volver a PC, otra vez el esqueleto.
  await prueba('esqueleto de PC y de celular (editor de color, avanzado y collage)', async () => {
    const h = {};
    const el = (id) => D().getElementById(id);
    const archivo = async (n = 1) => {
      const c = D().createElement('canvas'); c.width = 300; c.height = 200;
      const cx = c.getContext('2d'); cx.fillStyle = 'rgb(40,40,40)'; cx.fillRect(0, 0, 150, 200); cx.fillStyle = 'rgb(220,220,220)'; cx.fillRect(150, 0, 150, 200);
      const b = await new Promise((r) => c.toBlob(r, 'image/png'));
      return Array.from({ length: n }, (_, i) => new (W().File)([b], 'f' + i + '.png', { type: 'image/png' }));
    };
    const cargar = async (n = 1) => {
      el('entradaArchivo').files = archivos(await archivo(n)).files;
      el('entradaArchivo').dispatchEvent(new (W().Event)('change'));
      afirmar(await hasta(() => D().body.classList.contains('ed-esqueleto-activo'), 15000), 'el esqueleto no se armó');
      await esperar(200);
    };
    const visibles = () => [...new Set([...D().querySelectorAll('.ed-cat-visible')].map((b) => b.getAttribute('data-categoria')))].join(',');
    const rail = () => [...D().querySelectorAll('.ed-rail button')].map((b) => b.getAttribute('data-ed-categoria')).join(',');
    const categoria = (c) => D().querySelector('.ed-rail [data-ed-categoria="' + c + '"]').click();
    const tecla = (key) => D().body.dispatchEvent(new (W().KeyboardEvent)('keydown', { key, bubbles: true }));
    const enBarra = (id) => !!D().querySelector('.ed-barra #' + id + ', .ed-menu #' + id);
    const comprobarPc = (nombre, ids) => {
      const fila = D().querySelector('.ed-anuncio-fila');
      const anuncio = D().querySelector('.anuncio');
      afirmar(fila && fila.contains(anuncio), nombre + ': el anuncio no está en su franja');
      const alto = Math.round(anuncio.getBoundingClientRect().height);
      afirmar(alto === 90, nombre + ': el anuncio mide ' + alto + ' px de alto');
      const r = anuncio.getBoundingClientRect();
      const encima = D().elementFromPoint(r.left + r.width / 2, r.top + r.height / 2);
      afirmar(encima && (anuncio === encima || anuncio.contains(encima)), nombre + ': algo tapa el anuncio: ' + (encima && encima.className));
      for (const id of ids) afirmar(enBarra(id), nombre + ': ' + id + ' no está en la barra');
    };
    // Celular (fase 7): la misma barra superior (con los botones reales, en
    // la barra o en "⋯"), la barra de abajo con los mismos íconos y el mismo
    // orden que la de la PC, el anuncio en su franja debajo de la imagen, y
    // una hoja que se abre y se cierra tocando la categoría.
    const chip = (c) => [...D().querySelectorAll('.app-chip')].find((b) => b.querySelector('.app-icono').textContent === W().EsqueletoEditor.CATEGORIAS[c].ico);
    const enHoja = () => [...D().querySelectorAll('.app-hoja > [data-categoria]')].map((b) => b.getAttribute('data-categoria'));
    const aCelularYVuelta = async (nombre, ids, enCelular) => {
      const ordenPc = rail();
      marco.style.width = '390px'; marco.style.height = '844px';
      afirmar(await hasta(() => D().body.classList.contains('app-activo') && D().body.classList.contains('ed-esqueleto-movil') && !D().body.classList.contains('ed-esqueleto-activo'), 5000), nombre + ': no pasó a celular');
      await esperar(150);
      const shell = D().querySelector('.app-con-esqueleto');
      afirmar(shell && shell.firstElementChild.classList.contains('ed-barra'), nombre + ': en celular la barra superior no está arriba');
      for (const id of ids) afirmar(shell.querySelector('.ed-barra #' + id + ', .ed-menu #' + id), nombre + ': ' + id + ' no está en la barra de celular ni en "⋯"');
      afirmar(!D().querySelector('.ed-rail') || !D().querySelector('.ed-rail').isConnected, nombre + ': quedó la barra de íconos de PC en celular');
      const ordenCelular = [...D().querySelectorAll('.app-chip')].map((b) => Object.keys(W().EsqueletoEditor.CATEGORIAS).find((c) => W().EsqueletoEditor.CATEGORIAS[c].ico === b.querySelector('.app-icono').textContent)).join(',');
      afirmar(ordenCelular === ordenPc, nombre + ': la barra de abajo (' + ordenCelular + ') no es la de la PC (' + ordenPc + ')');
      // El anuncio: debajo de la imagen, 50 px, sin nada encima.
      const tira = D().querySelector('.app-anuncio');
      const escenario = D().querySelector('.app-escenario');
      afirmar(tira && tira.contains(D().querySelector('.anuncio')), nombre + ': en celular el anuncio no está en su franja');
      afirmar(tira.getBoundingClientRect().top >= escenario.getBoundingClientRect().bottom - 1, nombre + ': el anuncio no está debajo de la imagen');
      const ra = D().querySelector('.anuncio').getBoundingClientRect();
      afirmar(Math.round(ra.height) === 50, nombre + ': el anuncio mide ' + Math.round(ra.height) + ' px en celular');
      const encima = D().elementFromPoint(ra.left + ra.width / 2, ra.top + ra.height / 2);
      afirmar(encima && D().querySelector('.anuncio').contains(encima), nombre + ': algo tapa el anuncio en celular: ' + (encima && encima.className));
      // La hoja arranca cerrada; tocar una categoría la abre y tocarla otra vez la cierra.
      afirmar(!enHoja().length, nombre + ': la hoja arrancó abierta: ' + enHoja());
      await enCelular();
      marco.style.width = ''; marco.style.height = '';
      afirmar(await hasta(() => D().body.classList.contains('ed-esqueleto-activo'), 5000), nombre + ': no volvió el esqueleto');
      await esperar(150);
      comprobarPc(nombre + ' (de vuelta)', ids);
    };

    // --- Editor de color ---
    await abrir('/herramientas/Editor-de-color-de-imagen.html'); await esperar(300);
    await cargar();
    h.colorRail = rail();
    afirmar(h.colorRail === 'luz,color,bn,efectos,filtros,textura', 'barra de íconos del de color: ' + h.colorRail);
    afirmar(visibles() === 'luz', 'al abrir no se ve solo Luz: ' + visibles());
    comprobarPc('color', ['botonDeshacer', 'botonRehacer', 'botonDescargar']);
    afirmar(D().querySelector('.ed-menu #botonCambiarImagen') && D().querySelector('.ed-menu #botonRestablecer'), 'el menú ⋯ no tiene cambiar imagen y restablecer');
    categoria('efectos');
    h.efectos = D().querySelectorAll('.ed-cat-visible').length;
    afirmar(visibles() === 'efectos' && h.efectos === 3, 'Efectos: ' + visibles() + ' en ' + h.efectos + ' grupos');
    tecla('2');
    afirmar(visibles() === 'color', 'la tecla 2 no abrió Color: ' + visibles());
    categoria('color');
    afirmar(D().querySelector('.esqueleto-pc').classList.contains('ed-sin-columna'), 'tocar la categoría abierta no cerró la columna');
    categoria('color');
    // El escenario ocupa al menos la mitad de la pantalla (1366x900 en esta prueba).
    const esc = D().querySelector('[data-app-escenario]').getBoundingClientRect();
    h.escenario = Math.round(esc.width * esc.height / (W().innerWidth * W().innerHeight) * 100) + '%';
    afirmar(esc.width * esc.height >= 0.5 * W().innerWidth * W().innerHeight, 'el escenario ocupa ' + h.escenario);
    await aCelularYVuelta('color', ['botonDeshacer', 'botonDescargar', 'botonCambiarImagen', 'botonRestablecer'], async () => {
      afirmar(!D().querySelector('.app-acciones'), 'color: en celular siguen los botones de abajo (ya están en la barra)');
      chip('efectos').click();
      afirmar(enHoja().join() === 'efectos,efectos,efectos', 'color: la hoja de Efectos tiene ' + enHoja());
      chip('efectos').click();
      afirmar(!enHoja().length, 'color: tocar la categoría abierta no cerró la hoja');
      chip('luz').click();
      afirmar(enHoja().join() === 'luz', 'color: la hoja de Luz tiene ' + enHoja());
    });

    // --- Editor avanzado ---
    await abrir('/herramientas/editor-avanzado-imagen.html'); await esperar(300);
    await cargar();
    h.avanzadoRail = rail();
    afirmar(h.avanzadoRail === 'capas,texto,zona,borrar,luz,color,bn,efectos,filtros,textura', 'barra de íconos del avanzado: ' + h.avanzadoRail);
    afirmar(visibles() === 'capas', 'al abrir no se ve solo Capas: ' + visibles());
    comprobarPc('avanzado', ['botonDeshacer', 'botonRehacer', 'botonDescargarAvanzado', 'resolucionVistaPrevia']);
    // La lista de capas usa toda la columna (antes tenía 320 px de alto máximo).
    afirmar(W().getComputedStyle(el('listaCapas')).maxHeight === 'none', 'la lista de capas tiene alto máximo: ' + W().getComputedStyle(el('listaCapas')).maxHeight);
    afirmar(/Capa 1/.test(D().querySelector('.ed-editando').textContent), '"Editando" no dice la capa: ' + D().querySelector('.ed-editando').textContent);
    categoria('zona');
    afirmar(W().getComputedStyle(el('menuPiezas')).display === 'none', 'el menú de "Agregar pieza" se ve sin abrirlo');
    tecla('e');
    afirmar(visibles() === 'borrar' && W().eval('herramientaLienzo') === 'borrar' && D().querySelector('#lugarBorrar #controlesPincelAvanzado'), 'E no abrió Borrar con sus controles: ' + visibles() + ' / ' + W().eval('herramientaLienzo'));
    categoria('luz');
    afirmar(W().eval('herramientaLienzo') === 'mover' && !D().querySelector('#lugarBorrar #controlesPincelAvanzado'), 'salir de Borrar no soltó la herramienta');
    el('botonNuevoTexto').click();
    afirmar(visibles() === 'texto', 'agregar un texto no abrió Texto e imagen: ' + visibles());
    await esperar(200); // el historial anota el paso después de la acción
    // Los botones de la barra son los de siempre: Deshacer saca el texto.
    D().querySelector('.ed-barra #botonDeshacer').click();
    afirmar(!W().todosLosItems().some((i) => W().esObjeto(i)), 'Deshacer desde la barra no sacó el texto');
    const zoomAntes = D().querySelector('.ed-zoom-texto').textContent;
    D().querySelector('.ed-barra .ed-zoom-texto').nextElementSibling.click(); // +
    h.zoom = zoomAntes + ' -> ' + D().querySelector('.ed-zoom-texto').textContent;
    afirmar(D().querySelector('.ed-zoom-texto').textContent !== zoomAntes, 'el + de la barra no cambió el zoom');
    categoria('borrar');
    await aCelularYVuelta('avanzado', ['botonDeshacer', 'botonDescargarAvanzado', 'resolucionVistaPrevia', 'botonGuardarProyecto'], async () => {
      const editando = D().querySelector('.app-con-esqueleto > .ed-editando');
      afirmar(editando && /Capa/.test(editando.textContent), 'avanzado: en celular no está "Editando": ' + (editando && editando.textContent));
      chip('borrar').click();
      afirmar(enHoja().join() === 'borrar' && D().querySelector('.app-hoja #lugarBorrar #controlesPincelAvanzado') && W().eval('herramientaLienzo') === 'borrar', 'avanzado: Borrar en celular: ' + enHoja() + ' / ' + W().eval('herramientaLienzo'));
      chip('zona').click();
      afirmar(enHoja().join() === 'zona' && W().eval('herramientaLienzo') === 'mover' && !D().querySelector('#lugarBorrar #controlesPincelAvanzado'), 'avanzado: salir de Borrar en celular no soltó la herramienta');
      afirmar(W().getComputedStyle(D().querySelector('#herramientasLienzo [data-herramienta="borrar"]')).display === 'none', 'avanzado: en celular "Borrar" sigue también en Zona');
      el('botonNuevoTexto').click();
      afirmar(enHoja().join() === 'texto', 'avanzado: agregar un texto en celular no abrió Texto e imagen: ' + enHoja());
      // El zoom, en "⋯": − y + no cierran el menú.
      D().querySelector('.ed-boton-menu').click();
      const zt = D().querySelector('.ed-menu .ed-zoom-texto');
      const antes = zt.textContent;
      zt.nextElementSibling.click();
      afirmar(zt.textContent !== antes && !D().querySelector('.ed-menu').hidden, 'avanzado: el + del menú en celular: ' + antes + ' -> ' + zt.textContent);
      D().querySelector('.ed-boton-menu').click();
    });

    // --- Collage ---
    await abrir('/herramientas/collage-imagenes.html'); await esperar(300);
    await cargar(3);
    h.collageRail = rail();
    afirmar(h.collageRail === 'diseno,formato,estilo,fotos', 'barra de íconos del collage: ' + h.collageRail);
    comprobarPc('collage', ['botonDeshacer', 'botonDescargar']);
    afirmar(D().querySelector('#lugarFotos #tiraOrden') && D().querySelector('.ed-menu #tamanoSalida'), 'la tira no está en Fotos o el tamaño no está en el menú');
    W().eval('seleccionar(1)');
    afirmar(visibles() === 'fotos', 'elegir una foto no abrió Fotos: ' + visibles());
    await aCelularYVuelta('collage', ['botonDeshacer', 'botonDescargar', 'tamanoSalida'], async () => {
      W().eval('seleccionar(0)');
      afirmar(enHoja().join() === 'fotos' && D().querySelector('.app-hoja #tiraOrden') && D().querySelector('.app-hoja #barraFoto'), 'collage: tocar una foto en celular no abrió Fotos: ' + enHoja());
    });
    return Object.entries(h).map(([k, v]) => k + ' ' + v).join(' · ');
  });

  // Croma (2026-10-08): una pieza que marca lo que no es el fondo liso (verde,
  // azul…), sin IA. Foto de 200x100: fondo verde, una sombra verde oscura a
  // la izquierda y un objeto rojo en el medio. El color del fondo se detecta
  // solo, la sombra también es fondo, y con "Borrar lo que está fuera de la
  // zona" queda solo el objeto.
  // Pendientes de las fases 6 y 7 (2026-10-08, aprobados por el usuario):
  // ordenar las piezas de la zona (↑ ↓, arrastre, deshacer) y "＋ Ajuste para
  // «…»" con un objeto elegido, más la lista de las capas de adentro.
  await prueba('editor avanzado: ordenar piezas y "＋ Ajuste para"', async () => {
    const h = {};
    const el = (id) => D().getElementById(id);
    await abrir('/herramientas/editor-avanzado-imagen.html'); await esperar(300);
    el('entradaArchivo').files = archivos([await imagenDePrueba('A', 300, 200)]).files;
    el('entradaArchivo').dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => D().body.classList.contains('ed-esqueleto-activo'), 15000), 'el editor no abrió');
    await esperar(200);
    const tipos = () => W().itemActivo().mascara.piezas.map((x) => x.tipo).join(',');
    W().agregarPieza('luz');
    W().agregarPieza('pincel-restar');
    afirmar(tipos() === 'luz,pincel', 'piezas de partida: ' + tipos());
    afirmar(el('botonSubirPieza').disabled && !el('botonBajarPieza').disabled, 'la pieza de arriba puede subir o no puede bajar');
    el('botonBajarPieza').click();
    h.bajar = tipos();
    afirmar(h.bajar === 'pincel,luz', '↓ Bajar: ' + h.bajar);
    const filas = () => [...D().querySelectorAll('#listaPiezas .fila-pieza')];
    afirmar(filas()[0].textContent.includes(W().t('pieza-luz')), 'la lista no muestra Luz arriba');
    // Arrastrar: se reordenan las filas y al soltar las piezas toman ese orden.
    await esperar(200);
    filas()[1].parentNode.insertBefore(filas()[1], filas()[0]);
    W().ordenarPiezasSegunLista();
    h.arrastre = tipos();
    afirmar(h.arrastre === 'luz,pincel', 'arrastrar: ' + h.arrastre);
    await esperar(300);
    el('botonDeshacer').click();
    h.deshacer = tipos();
    afirmar(h.deshacer === 'pincel,luz', 'deshacer el arrastre: ' + h.deshacer);

    // "＋ Ajuste para «…»": con un objeto elegido, en Luz.
    el('botonNuevoTexto').click();
    await esperar(200);
    const objeto = W().itemActivo();
    afirmar(W().esObjeto(objeto), 'agregar un texto no eligió el objeto');
    D().querySelector('.ed-rail [data-ed-categoria="luz"]').click();
    const bloqueLuz = D().querySelector('[data-categoria="luz"]');
    const boton = bloqueLuz.querySelector('.aviso-solo-capas .boton-ajuste-para');
    afirmar(boton && boton.getBoundingClientRect().height > 0, 'no se ve "＋ Ajuste para" en Luz');
    afirmar(boton.textContent.includes(objeto.nombre), 'el botón no nombra al objeto: ' + boton.textContent);
    afirmar(bloqueLuz.querySelector('.acordeon-contenido').inert, 'con el objeto elegido los deslizadores no están inactivos');
    boton.click();
    const nueva = W().itemActivo();
    afirmar(!W().esGrupo(nueva) && W().objetoQueContiene(nueva) === objeto, 'la capa nueva no quedó adentro del objeto');
    afirmar(!bloqueLuz.querySelector('.acordeon-contenido').inert && !bloqueLuz.classList.contains('solo-capas'), 'con la capa nueva los deslizadores siguen inactivos');
    // Con el objeto elegido otra vez, la lista ofrece la capa de adentro.
    W().seleccionarItem(objeto.id);
    const deAdentro = [...bloqueLuz.querySelectorAll('.aviso-solo-capas .lista-de-adentro button')];
    afirmar(deAdentro.length === 1 && deAdentro[0].textContent.includes(nueva.nombre), 'la lista de adentro: ' + deAdentro.map((b) => b.textContent));
    deAdentro[0].click();
    afirmar(W().itemActivo() === nueva, 'tocar la capa de adentro no la eligió');
    // En Efectos (tres bloques) la tarjeta sale una sola vez.
    W().seleccionarItem(objeto.id);
    D().querySelector('.ed-rail [data-ed-categoria="efectos"]').click();
    h.tarjetasEfectos = [...D().querySelectorAll('.aviso-solo-capas')].filter((c) => c.getBoundingClientRect().height > 0).length;
    afirmar(h.tarjetasEfectos === 1, 'en Efectos se ven ' + h.tarjetasEfectos + ' tarjetas');
    return Object.entries(h).map(([k, v]) => k + ' ' + v).join(' · ');
  });

  await prueba('editor avanzado: croma (detecta el fondo verde, sombras, quitar el fondo)', async () => {
    await abrir('/herramientas/editor-avanzado-imagen.html'); await esperar(300);
    const c = D().createElement('canvas'); c.width = 200; c.height = 100;
    const cx = c.getContext('2d');
    cx.fillStyle = 'rgb(40,180,70)'; cx.fillRect(0, 0, 200, 100);
    cx.fillStyle = 'rgb(20,95,35)'; cx.fillRect(0, 0, 50, 100); // sombra sobre el fondo
    // El objeto, con un contorno de 3 px de color mezclado (verde amarillento):
    // el hilo que deja el suavizado de una foto real sobre un croma.
    cx.fillStyle = 'rgb(130,150,55)'; cx.fillRect(77, 17, 46, 66);
    cx.fillStyle = 'rgb(210,40,40)'; cx.fillRect(80, 20, 40, 60); // el objeto
    const entrada = D().getElementById('entradaArchivo');
    entrada.files = archivos([new (W().File)([await new Promise((r) => c.toBlob(r, 'image/png'))], 'verde.png', { type: 'image/png' })]).files;
    entrada.dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => W().eval('imagenActual') && D().querySelector('#listaCapas .fila-capa'), 15000), 'el editor no cargó la imagen');
    await esperar(300);
    const h = {};
    const pieza = agregarPieza('croma');
    afirmar(pieza && pieza.tipo === 'croma' && pieza.color, 'no se agregó el croma o no detectó el fondo');
    h.fondo = [pieza.color.r, pieza.color.g, pieza.color.b].join(',');
    afirmar(pieza.color.g > 150 && pieza.color.r < 80 && pieza.color.b < 100, 'no detectó el verde: ' + h.fondo);
    afirmar(D().getElementById('controlesCroma').style.display !== 'none', 'no aparecen los controles del croma');
    D().querySelector('#herramientasLienzo [data-herramienta="borrar"]').click();
    D().getElementById('botonBorrarFueraZona').click();
    await esperar(300);
    const lienzo = D().getElementById('canvasTrabajo');
    const alfa = (x, y) => lienzo.getContext('2d').getImageData(Math.round(x * lienzo.width / 200), Math.round(y * lienzo.height / 100), 1, 1).data[3];
    h.alfa = ['sombra ' + alfa(20, 50), 'verde ' + alfa(160, 50), 'objeto ' + alfa(100, 50)].join(' ');
    afirmar(alfa(20, 50) < 5 && alfa(160, 50) < 5 && alfa(100, 50) > 250, 'quitar el fondo con el croma: ' + h.alfa);
    // Recortar el borde: sin recorte, el contorno mezclado queda como un hilo
    // alrededor del objeto; con 3 px (de la imagen original) desaparece y el
    // objeto sigue entero.
    const conRecorte = async (valor) => {
      D().getElementById('botonRestaurarBorrado').click();
      const deslizador = D().getElementById('recorteCroma');
      deslizador.value = valor; deslizador.dispatchEvent(new (W().Event)('input'));
      D().getElementById('botonBorrarFueraZona').click();
      await esperar(300);
      return [alfa(78.5, 50), alfa(100, 50)];
    };
    const sinRecorte = await conRecorte(0);
    h.hiloSinRecorte = sinRecorte.join(',');
    afirmar(sinRecorte[0] > 250, 'sin recorte el contorno mezclado no quedó (la prueba no reproduce el hilo): ' + h.hiloSinRecorte);
    const recortado = await conRecorte(3);
    h.hiloRecortado = recortado.join(',');
    afirmar(recortado[0] < 5 && recortado[1] > 250, 'recortar el borde 3 px no quitó el hilo o se comió el objeto: ' + h.hiloRecortado);
    afirmar(W().piezaActiva().recorte === 3, 'el deslizador no cambió el recorte de la pieza');
    // Y el proyecto guarda su configuración.
    afirmar(W().MascarasCapa.TIPOS.includes('croma') && W().MascarasCapa.normalizar({ piezas: [{ tipo: 'croma', color: { r: 1, g: 2, b: 3 }, suavidad: 7 }] }).piezas[0].suavidad === 7, 'el croma no se guarda bien en el proyecto');
    return Object.entries(h).map(([k, v]) => k + ' ' + v).join(' · ');
  });

  // Máscara por piezas (2026-10-07, fase 5 del plan): sumar, restar e
  // intersecar, invertir una pieza, apagarla, deshacer, y la conversión de
  // los proyectos de la versión 7 (base + pincel) a piezas. Imagen de dos
  // mitades (40 | 220) con la capa en Multiplicar: donde se aplica, 40 -> 6 y
  // 220 -> 190. La comparación píxel por píxel con la rama anterior se hizo
  // aparte (hallazgos.md).
  await prueba('editor avanzado: máscara por piezas (sumar, restar, intersecar, invertir, ojo, proyecto v7)', async () => {
    await abrir('/herramientas/editor-avanzado-imagen.html'); await esperar(300);
    const base = D().createElement('canvas'); base.width = 100; base.height = 50;
    const cb = base.getContext('2d');
    cb.fillStyle = 'rgb(40,40,40)'; cb.fillRect(0, 0, 50, 50);
    cb.fillStyle = 'rgb(220,220,220)'; cb.fillRect(50, 0, 50, 50);
    const entrada = D().getElementById('entradaArchivo');
    entrada.files = archivos([new (W().File)([await new Promise((r) => base.toBlob(r, 'image/png'))], 'mitades.png', { type: 'image/png' })]).files;
    entrada.dispatchEvent(new (W().Event)('change'));
    const selector = await hasta(() => D().querySelector('#listaCapas .fila-opacidad select'), 15000);
    afirmar(selector, 'el editor no cargó la imagen');
    selector.value = 'multiply'; selector.dispatchEvent(new (W().Event)('change'));
    const lienzo = D().getElementById('canvasTrabajo');
    const px = (x) => lienzo.getContext('2d').getImageData(Math.round(lienzo.width * x), lienzo.height >> 1, 1, 1).data[0];
    const leer = async () => { await esperar(250); return [px(0.25), px(0.75)]; };
    const es = (v, a, b) => Math.abs(v[0] - a) <= 3 && Math.abs(v[1] - b) <= 3;
    const clic = (sel) => D().querySelector(sel).click();
    const operacion = (op) => clic('#modosOperacion [data-op="' + op + '"]');
    const filas = () => [...D().querySelectorAll('#listaPiezas .fila-pieza')];
    const lista = () => filas().map((f) => f.querySelector('.nombre-pieza').textContent.trim()).join(' | ');
    const h = {};

    agregarPieza('luz'); clic('#rangosLuz [data-rango="sombras"]');
    h.sombras = await leer();
    afirmar(es(h.sombras, 6, 220), 'luz en sombras: ' + h.sombras);
    agregarPieza('toda');
    h.sumar = await leer();
    afirmar(es(h.sumar, 6, 190), 'sombras + toda: ' + h.sumar);
    afirmar(filas().length === 2 && filas()[0].classList.contains('activa'), 'la pieza nueva no quedó arriba y elegida: ' + lista());
    operacion('intersecar');
    h.intersecar = await leer();
    afirmar(es(h.intersecar, 6, 220), 'sombras ∩ toda: ' + h.intersecar);
    afirmar(filas()[0].querySelector('.signo-pieza').textContent === '∩', 'la lista no muestra la operación: ' + lista());
    operacion('restar');
    h.restar = await leer();
    afirmar(es(h.restar, 40, 220), 'sombras − toda: ' + h.restar);
    // El ojo apaga la pieza: queda solo la luz.
    filas()[0].querySelector('.icono-capa').click();
    h.apagada = await leer();
    afirmar(es(h.apagada, 6, 220), 'con "toda" apagada: ' + h.apagada);
    // Invertir solo la pieza de luz: pasa a todo menos las sombras.
    filas()[1].querySelector('.nombre-pieza').click();
    afirmar(W().piezaActiva().tipo === 'luz', 'tocar la fila no eligió la pieza');
    invertirPieza(true);
    h.luzInvertida = await leer();
    afirmar(es(h.luzInvertida, 40, 190), 'luz invertida: ' + h.luzInvertida);
    afirmar(/\(invertida\)/.test(lista()), 'la lista no dice que la pieza está invertida: ' + lista());
    // Deshacer vuelve a la luz sin invertir.
    D().getElementById('botonDeshacer').click();
    h.deshecho = await leer();
    afirmar(es(h.deshecho, 6, 220), 'deshacer invertir la pieza: ' + h.deshecho);
    // Borrar las piezas: sin ninguna, la capa no se aplica y la nota lo dice.
    vaciarZona();
    h.sinPiezas = await leer();
    afirmar(es(h.sinPiezas, 40, 220) && D().getElementById('notaSinPiezas').style.display !== 'none', 'sin piezas: ' + h.sinPiezas);

    // Proyecto de la versión 7: base "luz" (sombras) + pincel que suma la
    // mitad derecha + pincel que quita un cuadrado a la izquierda, sin
    // invertir -> (izquierda ∪ derecha) − cuadrado.
    const pintar = (x0, x1) => { const c = D().createElement('canvas'); c.width = 100; c.height = 50; const cx = c.getContext('2d'); cx.fillStyle = '#fff'; cx.fillRect(x0, 0, x1 - x0, 50); return c.toDataURL('image/png'); };
    const capaV7 = (mascara, sumar, quitar) => ({ id: 'capa-1', nombre: 'Capa 1', visible: true, opacidad: 100, modoFusion: 'multiply', ajustes: {}, mascara, pincelSumar: sumar, pincelQuitar: quitar });
    const abrirV7 = async (capa) => {
      const proyecto = { version: 7, imagen: base.toDataURL('image/png'), nombreArchivo: 'mitades.png', capaActivaId: 'capa-1', capas: [capa] };
      const entradaProyecto = D().getElementById('entradaProyecto');
      const antes = W().eval('capas');
      entradaProyecto.files = archivos([new (W().File)([JSON.stringify(proyecto)], 'v7.json', { type: 'application/json' })]).files;
      entradaProyecto.dispatchEvent(new (W().Event)('change'));
      afirmar(await hasta(() => W().eval('capas') !== antes, 10000), 'el proyecto v7 no abrió');
      return W().capaActiva().mascara.piezas.map((p) => p.tipo + ':' + p.op + (p.invertir ? ':inv' : '')).join(' ');
    };
    const luz = { base: 'luz', invertir: false, luz: { desde: 0, hasta: 35, suavidad: 15 }, degradado: { tipo: 'lineal', a: null, b: null }, color: null, tolerancia: 30, objeto: null, fuente: 'foto' };
    h.formaV7 = await abrirV7(capaV7(luz, pintar(50, 100), pintar(0, 25)));
    afirmar(h.formaV7 === 'luz:sumar pincel:sumar pincel:restar', 'conversión v7: ' + h.formaV7);
    const v7 = [px(0.1), px(0.4), px(0.75)];
    h.v7 = v7.join(',');
    afirmar(Math.abs(v7[0] - 40) <= 3 && Math.abs(v7[1] - 6) <= 3 && Math.abs(v7[2] - 190) <= 3, 'proyecto v7: ' + h.v7);
    // La misma, invertida: el cuadrado quitado es lo único que queda.
    h.formaV7Invertida = await abrirV7(capaV7({ ...luz, invertir: true }, pintar(50, 100), pintar(0, 25)));
    await esperar(250);
    const inv = [px(0.1), px(0.4), px(0.75)];
    h.v7Invertida = inv.join(',');
    afirmar(Math.abs(inv[0] - 6) <= 3 && Math.abs(inv[1] - 40) <= 3 && Math.abs(inv[2] - 220) <= 3 && W().capaActiva().mascara.invertir, 'proyecto v7 invertido: ' + h.v7Invertida);
    // Base "Nada" sin pintar: no se aplica en ningún lado.
    h.formaNada = await abrirV7(capaV7({ ...luz, base: 'nada' }, null, null));
    afirmar(h.formaNada === '' && es(await leer(), 40, 220), 'base "Nada": ' + h.formaNada);

    // Y la conversión de cada base, sin abrir nada.
    const M = W().MascarasCapa;
    const forma = (m) => m.piezas.map((p) => p.tipo + ':' + p.op + (p.invertir ? ':inv' : '') + (p.fuente ? ':' + p.fuente : '')).join(' ');
    h.bases = [
      forma(M.desdeV7({ base: 'fondo', fuente: 'capa' })),
      forma(M.desdeV7({ base: 'fondo', objeto: { x: 0.5, y: 0.5 } })),
      forma(M.desdeV7({ base: 'sujeto', objeto: { x: 0.2, y: 0.3 } })),
      forma(M.desdeV7({ base: 'color', color: { r: 1, g: 2, b: 3 }, tolerancia: 40 })),
      forma(M.desdeV7({ base: 'degradado', degradado: { tipo: 'radial', a: { x: 0, y: 0 }, b: { x: 1, y: 1 } } })),
      forma(M.desdeV7(null)),
    ].join(' / ');
    afirmar(h.bases === 'sujeto:sumar:inv:capa / objeto:sumar:inv:foto / objeto:sumar:foto / color:sumar:foto / degradado:sumar / toda:sumar', 'bases convertidas: ' + h.bases);
    afirmar(M.desdeV7({ base: 'sujeto', objeto: { x: 0.2, y: 0.3 } }).piezas[0].punto.x === 0.2
      && M.desdeV7({ base: 'color', color: { r: 1, g: 2, b: 3 }, tolerancia: 40 }).piezas[0].tolerancia === 40
      && M.desdeV7({ base: 'degradado', degradado: { tipo: 'radial', a: { x: 0, y: 0 }, b: { x: 1, y: 1 } } }).piezas[0].degradado.tipo === 'radial', 'la conversión perdió la configuración de la base');
    afirmar(M.esTodaLaImagen(M.neutra()) && !M.esTodaLaImagen({ piezas: [], invertir: false }) && M.esTodaLaImagen({ piezas: [M.pieza('luz'), M.pieza('toda')], invertir: false }), 'esTodaLaImagen');

    // El proyecto nuevo (versión 8) guarda las piezas y el lienzo del pincel.
    const guardada = W().itemAProyecto(W().capaActiva());
    afirmar(Array.isArray(guardada.mascara.piezas) && !('pincelSumar' in guardada), 'el proyecto no guarda la máscara como piezas');
    return Object.entries(h).map(([k, v]) => k + ' ' + v).join(' · ');
  });

  // Borrar a transparente: con la capa de la foto elegida borra la foto (queda
  // transparente, también bajo una capa de ajuste); con una capa de imagen,
  // borra esa imagen y lo borrado la acompaña al moverla. Foto JPG gris de
  // 200x100 -> la descarga pasa a PNG. Con eventos de puntero de verdad.
  await prueba('editor avanzado: borrar a transparente (foto, capa, mover, descarga PNG, proyecto, restaurar)', async () => {
    await abrir('/herramientas/editor-avanzado-imagen.html'); await esperar(300);
    const lienzoArchivo = async (ancho, alto, dibujar, nombre, tipo) => {
      const c = D().createElement('canvas'); c.width = ancho; c.height = alto; dibujar(c.getContext('2d'));
      const b = await new Promise((r) => c.toBlob(r, tipo, 0.95));
      return new (W().File)([b], nombre, { type: tipo });
    };
    const entrada = D().getElementById('entradaArchivo');
    entrada.files = archivos([await lienzoArchivo(200, 100, (cx) => { cx.fillStyle = 'rgb(128,128,128)'; cx.fillRect(0, 0, 200, 100); }, 'gris.jpg', 'image/jpeg')]).files;
    entrada.dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => W().eval('imagenActual'), 15000), 'la imagen no se cargó');
    const capaFoto = W().capaActiva();
    const lienzo = D().getElementById('canvasTrabajo');
    const clic = (sel) => D().querySelector(sel).click();
    const puntero = (tipo, x) => {
      const r = lienzo.getBoundingClientRect();
      lienzo.dispatchEvent(new (W().PointerEvent)(tipo, { bubbles: true, pointerId: 1, pointerType: 'mouse', isPrimary: true, buttons: tipo === 'pointerup' ? 0 : 1,
        clientX: r.left + (x / 200) * r.width, clientY: r.top + r.height / 2 }));
    };
    const trazo = (desde, hasta_) => { puntero('pointerdown', desde); puntero('pointermove', (desde + hasta_) / 2); puntero('pointermove', hasta_); puntero('pointerup', hasta_); };
    // "gris", "rojo", "transparente" o los valores, en x (de 0..200) a media altura.
    const ver = (cv, x) => {
      const d = cv.getContext('2d').getImageData(Math.round(x * cv.width / 200), Math.round(cv.height / 2), 1, 1).data;
      if (d[3] < 10) return 'transparente';
      if (d[3] < 245) return 'semi' + d[3];
      if (d[0] > 200 && d[2] < 60) return 'rojo';
      return Math.abs(d[0] - d[2]) < 8 ? 'gris' + (d[0] > 160 ? '+' : '') : d.join(',');
    };
    const leer = async (xs) => { await esperar(300); return xs.map((x) => ver(lienzo, x)).join(' '); };
    const h = {};
    const tamano = D().getElementById('tamanoPincelAvanzado');
    const factor = lienzo.width / lienzo.getBoundingClientRect().width;
    tamano.value = Math.max(6, Math.round(14 / factor)); tamano.dispatchEvent(new (W().Event)('input'));

    clic('#herramientasLienzo [data-herramienta="borrar"]');
    afirmar(D().getElementById('notaBorrar').textContent === W().t('nota-borrar-foto'), 'la nota no dice que borra la foto: ' + D().getElementById('notaBorrar').textContent);
    afirmar(D().getElementById('botonRestaurarBorrado').disabled, '"Restaurar" habilitado sin haber borrado nada');
    trazo(20, 30);
    h.foto = await leer([25, 100, 160]);
    afirmar(h.foto === 'transparente gris gris', 'borrar la foto: ' + h.foto);
    afirmar(D().getElementById('notaDescargaPng').style.display !== 'none', 'no avisa que la descarga sale en PNG');

    // Una capa de ajuste encima no rellena el agujero.
    capaFoto.ajustes.exposicion = 60; W().programarRenderizado();
    h.conAjuste = await leer([25, 100]);
    afirmar(h.conAjuste === 'transparente gris+', 'con exposición: ' + h.conAjuste);

    // Una imagen roja de 100x50 como capa (x 50..150): borrar en x 120..130 borra la imagen.
    const entradaCapa = D().getElementById('entradaImagenCapa');
    entradaCapa.files = archivos([await lienzoArchivo(100, 50, (cx) => { cx.fillStyle = '#f00'; cx.fillRect(0, 0, 100, 50); }, 'rojo.png', 'image/png')]).files;
    entradaCapa.dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => W().esObjeto(W().itemActivo()), 5000), 'la imagen no entró como objeto');
    const capa = W().capaDeContenido(W().itemActivo());
    clic('#herramientasLienzo [data-herramienta="borrar"]');
    afirmar(D().getElementById('notaBorrar').textContent === W().t('nota-borrar-capa'), 'la nota no dice que borra la capa');
    trazo(120, 130);
    h.capa = await leer([25, 75, 125]);
    afirmar(h.capa === 'transparente rojo gris+', 'borrar la capa: ' + h.capa);
    afirmar(capa.borrado && capa.borrado.width === 100, 'lo borrado de la capa no está a su tamaño: ' + (capa.borrado && capa.borrado.width));
    // Movida a la izquierda (x 0..100): el hueco va con ella (x 75) y tapa el de la foto.
    capa.contenido.centro = { x: 0.25, y: 0.5 }; W().programarRenderizado();
    h.movida = await leer([25, 75, 125]);
    afirmar(h.movida === 'rojo gris+ gris+', 'la capa movida: ' + h.movida);
    capa.contenido.centro = { x: 0.5, y: 0.5 }; W().programarRenderizado();

    // Descarga: JPG -> PNG con el agujero.
    D().getElementById('botonDescargarAvanzado').click();
    const enlace = D().getElementById('enlaceDescarga');
    afirmar(await hasta(() => enlace.href.startsWith('blob:'), 15000), 'la descarga no se generó');
    const blob = await (await W().fetch(enlace.href)).blob();
    afirmar(blob.type === 'image/png' && enlace.download.endsWith('.png'), 'la descarga no salió en PNG: ' + blob.type + ' ' + enlace.download);
    const img = new (W().Image)(); img.src = enlace.href; await img.decode();
    const cd = D().createElement('canvas'); cd.width = img.width; cd.height = img.height;
    cd.getContext('2d').drawImage(img, 0, 0);
    h.descarga = [25, 75, 125].map((x) => ver(cd, x)).join(' ');
    afirmar(h.descarga === 'transparente rojo gris+', 'descarga: ' + h.descarga);

    // Proyecto: guardar (lo mismo que arma el botón) y volver a abrir.
    const proyecto = {
      version: 7, imagen: W().crearCanvasConImagen(W().eval('imagenActual'), 200, 100).toDataURL('image/jpeg', 0.95), nombreArchivo: 'gris.jpg',
      capaActivaId: capaFoto.id, capas: W().eval('capas').map(W().itemAProyecto), borradoFoto: W().canvasADataUrl(W().eval('borradoFoto')),
    };
    const entradaProyecto = D().getElementById('entradaProyecto');
    entradaProyecto.files = archivos([new (W().File)([JSON.stringify(proyecto)], 'p.json', { type: 'application/json' })]).files;
    entradaProyecto.dispatchEvent(new (W().Event)('change'));
    afirmar(await hasta(() => W().eval('capas').length === 2 && W().eval('capas')[1].hijos[0].borrado, 10000), 'el proyecto no abrió con lo borrado de la capa');
    h.proyecto = await leer([25, 75, 125]);
    afirmar(h.proyecto === 'transparente rojo gris+', 'proyecto: ' + h.proyecto);

    // Restaurar lo borrado de la foto (la capa de la foto está elegida).
    clic('#herramientasLienzo [data-herramienta="borrar"]');
    clic('#botonRestaurarBorrado');
    h.restaurada = await leer([25, 125]);
    afirmar(h.restaurada === 'gris+ gris+', 'restaurar la foto: ' + h.restaurada);
    afirmar(D().getElementById('notaDescargaPng').style.display === 'none', 'sigue avisando PNG sin nada borrado de la foto');
    return Object.entries(h).map(([k, v]) => k + ' ' + v).join(' · ') + ' · ' + blob.type;
  });

  for (const [ruta, texto, espera] of [['/index.html', 'sticker', 'Generador de stickers'], ['/en/index.html', 'merge pdf', 'Merge PDFs and images']]) {
    await prueba('buscador "' + texto + '"', async () => {
      await abrir(ruta); await esperar(200);
      const e = D().getElementById('buscador'); e.value = texto; e.dispatchEvent(new (W().Event)('input'));
      await esperar(100);
      const t = [...D().querySelectorAll('#gridResultados h3')].map((h) => h.textContent);
      afirmar(t.includes(espera), 'encontró: ' + (t.join(', ') || 'nada'));
      return t.join(', ');
    });
  }

  await informar(true);
})();
`;

// El código de las pruebas va dentro de un <script>: un "</script>" escrito en
// una prueba (la de Markdown → HTML lo usa para probar la sanitización) cerraba
// el script de la página y ninguna prueba arrancaba. Dentro de un texto de JS,
// "<\/script" es lo mismo y el navegador ya no lo toma como cierre.
function paginaDeControl(lista) {
  return `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><title>pruebas</title>
<style>html,body{margin:0}iframe{width:1366px;height:900px;border:0}</style></head><body>
<iframe id="marco" allow="fullscreen"></iframe>
<script>const PAGINAS = ${JSON.stringify(lista)};\nconst SOLO = ${JSON.stringify(process.env.PROBAR_SOLO || '')};\n${PRUEBAS.replace(/<\/(script)/gi, '<\\/$1')}</script></body></html>`;
}

// --- Principal --------------------------------------------------------------
if (!existsSync(DIST)) { console.error('No existe dist/. Corré primero: npm run build'); process.exit(1); }
if (!EDGE) { console.error('No se encontró Edge. Indicá la ruta con la variable EDGE.'); process.exit(1); }

const lista = paginas(DIST).sort();
const control = join(DIST, '__pruebas.html');
writeFileSync(control, paginaDeControl(lista), 'utf8');

const faltantes = new Set();
// Rutas que una prueba pide sabiendo que no existen: "markdown-html: imagen con
// ruta relativa" escribe fotos/vacaciones.png para comprobar que la ruta no se
// altera, y la vista previa intenta cargarla.
const FALTAN_A_PROPOSITO = new Set(['/herramientas/fotos/vacaciones.png']);
let terminar;
const final = new Promise((r) => { terminar = r; });
let ultimo = null;

const servidor = createServer((q, r) => {
  if (q.method === 'POST' && q.url === '/__avance') {
    let cuerpo = '';
    q.on('data', (d) => { cuerpo += d; });
    q.on('end', () => { const a = JSON.parse(cuerpo); if (a.ms > 3000 || a.i % 40 === 0) console.log(`    página ${a.i + 1}/${a.total} ${a.p} (${a.ms} ms)`); r.end('ok'); });
    return;
  }
  if (q.method === 'POST' && q.url === '/__resultado') {
    let cuerpo = '';
    q.on('data', (d) => { cuerpo += d; });
    q.on('end', () => {
      ultimo = JSON.parse(cuerpo);
      const u = ultimo.resultados[ultimo.resultados.length - 1];
      if (u && !ultimo.fin) console.log(`  ${u.ok ? '✓' : '✗'} ${u.nombre} (${u.s}s)`);
      r.end('ok');
      if (ultimo.fin) terminar();
    });
    return;
  }
  const ruta = decodeURIComponent(q.url.split('?')[0]);
  let archivo = join(DIST, ruta);
  if (existsSync(archivo) && statSync(archivo).isDirectory()) archivo = join(archivo, 'index.html');
  if (!existsSync(archivo) && existsSync(archivo + '.html')) archivo += '.html';
  if (!archivo.startsWith(DIST) || !existsSync(archivo)) {
    if (ruta !== '/favicon.ico' && !FALTAN_A_PROPOSITO.has(ruta)) faltantes.add(ruta);
    r.writeHead(404).end(); return;
  }
  r.writeHead(200, { 'Content-Type': TIPOS[extname(archivo)] || 'application/octet-stream' });
  r.end(readFileSync(archivo));
});
await new Promise((r) => servidor.listen(PUERTO, r));

console.log(`Probando sin internet (${lista.length} páginas)...`);
const perfil = mkdtempSync(join(tmpdir(), 'mc-pruebas-'));
const edge = spawn(EDGE, [
  // headless=new y sin frenar temporizadores: con el modo viejo, una página en
  // segundo plano (el iframe) quedaba frenada y la primera prueba no terminaba.
  '--headless=new', '--disable-background-timer-throttling', '--disable-renderer-backgrounding',
  '--disable-backgrounding-occluded-windows', `--user-data-dir=${perfil}`, '--no-first-run',
  '--host-resolver-rules=MAP * ~NOTFOUND, EXCLUDE localhost',
  `http://localhost:${PUERTO}/__pruebas.html`,
], { windowsHide: true, stdio: 'ignore' });

const vencido = setTimeout(() => terminar('tiempo'), LIMITE_MS);
const motivo = await final;
clearTimeout(vencido);
edge.kill();
servidor.close();
rmSync(control, { force: true });
try { rmSync(perfil, { recursive: true, force: true }); } catch {}

const resultados = ultimo ? ultimo.resultados : [];
const fallas = resultados.filter((r) => !r.ok);
if (motivo === 'tiempo') fallas.push({ nombre: `las pruebas no terminaron en ${LIMITE_MS / 60000} minutos`, detalle: '' });
if (faltantes.size) fallas.push({ nombre: `${faltantes.size} archivos pedidos que no existen en dist/`, detalle: [...faltantes].slice(0, 8).join(', ') });

console.log('');
for (const r of resultados) console.log(`${r.ok ? '✓' : '✗'} ${r.nombre} — ${r.detalle}`);
if (fallas.length) {
  console.error(`\n${fallas.length} prueba(s) fallaron:`);
  for (const f of fallas) console.error(`  ✗ ${f.nombre}${f.detalle ? ': ' + f.detalle : ''}`);
  process.exit(1);
}
console.log(`\nTodo OK — ${resultados.length} pruebas.`);
process.exit(0);

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
<iframe id="marco"></iframe>
<script>const PAGINAS = ${JSON.stringify(lista)};\n${PRUEBAS.replace(/<\/(script)/gi, '<\\/$1')}</script></body></html>`;
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

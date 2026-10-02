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
    for (const p of PAGINAS) {
      await abrir(p); await esperar(120);
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

function paginaDeControl(lista) {
  return `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><title>pruebas</title>
<style>html,body{margin:0}iframe{width:1366px;height:900px;border:0}</style></head><body>
<iframe id="marco"></iframe>
<script>const PAGINAS = ${JSON.stringify(lista)};\n${PRUEBAS}</script></body></html>`;
}

// --- Principal --------------------------------------------------------------
if (!existsSync(DIST)) { console.error('No existe dist/. Corré primero: npm run build'); process.exit(1); }
if (!EDGE) { console.error('No se encontró Edge. Indicá la ruta con la variable EDGE.'); process.exit(1); }

const lista = paginas(DIST).sort();
const control = join(DIST, '__pruebas.html');
writeFileSync(control, paginaDeControl(lista), 'utf8');

const faltantes = new Set();
let terminar;
const final = new Promise((r) => { terminar = r; });
let ultimo = null;

const servidor = createServer((q, r) => {
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
    if (ruta !== '/favicon.ico') faltantes.add(ruta);
    r.writeHead(404).end(); return;
  }
  r.writeHead(200, { 'Content-Type': TIPOS[extname(archivo)] || 'application/octet-stream' });
  r.end(readFileSync(archivo));
});
await new Promise((r) => servidor.listen(PUERTO, r));

console.log(`Probando sin internet (${lista.length} páginas)...`);
const perfil = mkdtempSync(join(tmpdir(), 'mc-pruebas-'));
const edge = spawn(EDGE, [
  '--headless', `--user-data-dir=${perfil}`, '--no-first-run',
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

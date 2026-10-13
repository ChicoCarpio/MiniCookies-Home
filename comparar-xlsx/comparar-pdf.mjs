// Genera con el motor propio (xlsx-pdf.js, en Edge) el PDF de cada hoja de un libro del corpus y lo compara con el que exporta Excel
// (volcar-paginas.ps1 guarda esos PDF en medidas/tmp/pdf-excel/<libro>): paginas, tamano de hoja y donde empieza cada texto.
//
//   node comparar-xlsx/comparar-pdf.mjs 24-impresion            todas las hojas del libro (Carta; con A4 el PDF de Excel sale reducido)
//   node comparar-xlsx/comparar-pdf.mjs 27-impresion-extras Enc9 solo esas hojas
//
// Un texto "coincide" si en el PDF de Excel hay otro que empieza a menos de 0,7 pt en x y en y (Excel parte las cadenas en trozos
// por el kerning: se mira el primero de cada cadena del motor). Solo con archivos sinteticos. Necesita Windows con Excel y Edge.
import { createServer } from 'node:http';
import { spawn } from 'node:child_process';
import { readFileSync, writeFileSync, existsSync, mkdirSync, mkdtempSync, rmSync, statSync } from 'node:fs';
import { join, extname, resolve } from 'node:path';
import { tmpdir } from 'node:os';
import { leerPdf } from './leer-pdf-excel.mjs';

const RAIZ = resolve(import.meta.dirname, '..');
const CARPETA = join(RAIZ, 'comparar-xlsx');
const [libroNombre, ...soloHojas] = process.argv.slice(2).filter((a) => !a.startsWith('--'));
if (!libroNombre) { console.error('Uso: node comparar-xlsx/comparar-pdf.mjs <libro-del-corpus> [hojas]'); process.exit(1); }
const TOL = +(process.env.TOL || 0.7);
const XLSX = join(CARPETA, 'corpus', libroNombre + '.xlsx');
const VOLCADO = JSON.parse(readFileSync(join(CARPETA, 'medidas', 'paginas', libroNombre + '.json'), 'utf8').replace(/^﻿/, ''));
const PDF_EXCEL = join(CARPETA, 'medidas', 'tmp', 'pdf-excel', libroNombre);
const SALIDA = join(CARPETA, 'medidas', 'tmp', 'pdf-motor', libroNombre);
mkdirSync(SALIDA, { recursive: true });
const PUERTO = 8795;
const EDGE = [process.env.EDGE, 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe', 'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe'].filter(Boolean).find((r) => existsSync(r));
const TIPOS = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.mjs': 'text/javascript; charset=utf-8', '.json': 'application/json', '.ttf': 'font/ttf', '.woff2': 'font/woff2', '.wasm': 'application/wasm', '.txt': 'text/plain' };
if (!EDGE) { console.error('No se encontro Microsoft Edge.'); process.exit(1); }

// Hojas con filas mas bajas que la letra (6 y 12,75 pt con Calibri 11): Excel recorta el texto y su linea base ahi no sigue una regla simple.
const CONOCIDAS = new Set(['VariadoH', 'AjusteVariado', 'AjusteVariado2', 'HF']); // HF: el nombre de archivo del pie sale 1 pt más ancho
const hojas = VOLCADO.filter((h) => !h.error && (!soloHojas.length || soloHojas.includes(h.nombre)));

const CONTROL = String.raw`<!doctype html><meta charset="utf-8"><body><script>
const HOJAS = __HOJAS__;
const NOMBRE = __NOMBRE__;
const enviar = (ruta, cuerpo) => fetch(ruta, { method: 'POST', body: cuerpo });
const cargar = (src) => new Promise((ok, mal) => { const s = document.createElement('script'); s.src = src; s.onload = ok; s.onerror = () => mal(new Error('no cargo ' + src)); document.head.appendChild(s); });
(async () => {
  for (const s of ['vendor.js', 'zip-propio.js', 'xml-ligero.js', 'xlsx-estilos.js', 'xlsx-modelo.js', 'xlsx-formato.js', 'xlsx-geometria.js', 'pptx-texto.js', 'xlsx-paginas.js', 'xlsx-dibujo.js']) await cargar('/assets/js/' + s);
  await cargar(rutaVendor('pdf-lib/1.17.1/pdf-lib.min.js'));
  await cargar('/assets/js/pdf-lienzo.js');
  await cargar('/assets/js/xlsx-pdf.js');
  const libro = await LibroXlsx.abrir(new Uint8Array(await (await fetch('/__xlsx')).arrayBuffer()));
  for (const h of HOJAS) {
    const r = { nombre: h.nombre };
    try {
      const t0 = performance.now();
      const { bytes, paginas } = await XlsxPdf.generar(libro, [h.indice], { idioma: 'es', nombreArchivo: NOMBRE, numeracion: 'porHoja' });
      r.ms = Math.round(performance.now() - t0); r.paginas = paginas;
      await enviar('/__pdf/' + encodeURIComponent(h.nombre), bytes);
    } catch (e) { r.error = String(e && e.stack || e); }
    await enviar('/__resultado', JSON.stringify(r));
  }
  await enviar('/__fin', '1');
})().catch((e) => enviar('/__fin', 'ERROR ' + (e && e.stack || e)));
</script>`;

const resultados = new Map();
let terminar; const fin = new Promise((r) => { terminar = r; });
const servidor = createServer(async (q, s) => {
  const ruta = decodeURIComponent(q.url.split('?')[0]);
  if (q.method === 'POST') {
    const partes = []; for await (const p of q) partes.push(p);
    const cuerpo = Buffer.concat(partes);
    if (ruta.startsWith('/__pdf/')) writeFileSync(join(SALIDA, ruta.slice(7) + '.pdf'), cuerpo);
    else if (ruta === '/__resultado') { const r = JSON.parse(cuerpo.toString()); resultados.set(r.nombre, r); if (r.error) console.log(`  ${r.nombre}: ERROR ${r.error.split('\n').slice(0, 4).join(' | ')}`); }
    else if (ruta === '/__fin') { const t = cuerpo.toString(); if (t !== '1') console.error(t); terminar(); }
    s.end('ok'); return;
  }
  if (ruta === '/__control') { s.setHeader('content-type', 'text/html; charset=utf-8'); return s.end(CONTROL.replace('__NOMBRE__', JSON.stringify(libroNombre + '.xlsx')).replace('__HOJAS__', JSON.stringify(hojas.map((h) => ({ nombre: h.nombre, indice: h.indice }))))); }
  if (ruta === '/__xlsx') return s.end(readFileSync(XLSX));
  const p = join(RAIZ, ruta);
  if (!p.startsWith(RAIZ) || !existsSync(p) || statSync(p).isDirectory()) { s.statusCode = 404; return s.end('no'); }
  s.setHeader('content-type', TIPOS[extname(p)] || 'application/octet-stream');
  s.end(readFileSync(p));
}).listen(PUERTO);

const perfil = mkdtempSync(join(tmpdir(), 'comparar-pdf-'));
const edge = spawn(EDGE, ['--headless=new', '--disable-gpu', '--lang=es-CR', '--user-data-dir=' + perfil, '--window-size=1300,950', `http://localhost:${PUERTO}/__control`], { stdio: 'ignore' });
const reloj = setTimeout(() => { console.error('Se paso el tiempo limite.'); terminar(); }, 10 * 60 * 1000);
await fin;
clearTimeout(reloj); edge.kill(); servidor.close();
setTimeout(() => { try { rmSync(perfil, { recursive: true, force: true }); } catch { /* Edge todavia lo usa */ } }, 1500);

// --- comparar ---
let mal = 0;
for (const h of hojas) {
  const r = resultados.get(h.nombre);
  if (!r || r.error) { mal++; continue; }
  const archivoMotor = join(SALIDA, encodeURIComponent(h.nombre) + '.pdf');
  if (!existsSync(archivoMotor)) { console.log(`✗ ${h.nombre}: no hay PDF del motor`); mal++; continue; }
  const motor = await leerPdf(archivoMotor);
  const excel = await leerPdf(join(PDF_EXCEL, h.pdf));
  const igual = (a, b) => Math.abs(a.ancho - b.ancho) < 1.5 && Math.abs(a.alto - b.alto) < 1.5;
  const esCarta = (q) => (Math.abs(q.ancho - 612) < 1 && Math.abs(q.alto - 792) < 1) || (Math.abs(q.ancho - 792) < 1 && Math.abs(q.alto - 612) < 1);
  // la impresora de las mediciones exporta la hoja A4 en Carta (con el contenido reducido): ahi no se compara ni el papel ni las posiciones
  const a4EnCarta = esCarta(excel[0]) && !esCarta(motor[0]);
  const papelOk = igual(motor[0], excel[0]) || a4EnCarta;
  const carta = true;
  const tol = TOL;
  let encontrados = 0, total = 0, maxDx = 0, maxDy = 0;
  const ejemplos = [];
  const n = Math.min(motor.length, excel.length);
  for (let i = 0; i < n; i++) {
    const E = excel[i].textos;
    // el motor dibuja los numeros letra por letra (ancho fijo del digito): se junta lo que va seguido en la misma linea y se mira el primero
    const xMax = Math.max(...E.map((e) => e.x), 0) + 60; // lo recortado fuera de la pagina (vecinas que desbordan) no se ve
    const ordenados = motor[i].textos.filter((t) => t.x <= xMax).sort((a, b) => a.y - b.y || a.x - b.x);
    const grupos = [];
    for (const t of ordenados) { const u = grupos[grupos.length - 1]; if (u && Math.abs(u.y - t.y) < 0.2 && t.x - u.ultimo <= t.tam * 1.6) u.ultimo = t.x; else grupos.push({ x: t.x, y: t.y, ultimo: t.x, tam: t.tam }); }
    for (const t of grupos) {
      total++;
      let mejor = null, md = 1e9;
      for (const e of E) { const d = Math.hypot(e.x - t.x, e.y - t.y); if (d < md) { md = d; mejor = e; } }
      if (mejor && Math.abs(mejor.x - t.x) <= tol && Math.abs(mejor.y - t.y) <= tol) encontrados++;
      else if (ejemplos.length < 4) ejemplos.push(`pag ${i + 1}: motor (${t.x.toFixed(1)}, ${t.y.toFixed(1)}) Excel mas cercano ${mejor ? `(${mejor.x.toFixed(1)}, ${mejor.y.toFixed(1)})` : '-'}`);
      if (mejor) { maxDx = Math.max(maxDx, Math.abs(mejor.x - t.x)); maxDy = Math.max(maxDy, Math.abs(mejor.y - t.y)); }
    }
  }
  const pag = motor.length === excel.length;
  // la A4 sale de esa impresora reducida dentro de una Carta: en esas hojas solo se comparan las paginas (las posiciones, en las hojas Carta)
  const umbral = CONOCIDAS.has(h.nombre) ? 0.9 : 0.98;
  const ok = pag && papelOk && (a4EnCarta || total === 0 || encontrados / total > umbral);
  if (!ok) mal++;
  console.log(`${ok ? '✓' : '✗'} ${h.nombre.padEnd(18)} paginas ${motor.length}/${excel.length}${papelOk ? '' : ' PAPEL ' + motor[0].ancho + 'x' + motor[0].alto + ' vs ' + excel[0].ancho + 'x' + excel[0].alto}  textos ${encontrados}/${total} en ±${tol.toFixed(1)} pt${a4EnCarta ? ' (A4 en Carta: solo se comparan las paginas)' : ''}  ${r.ms} ms`);
  if (!ok) for (const e of ejemplos) console.log('      ' + e);
}
console.log(`\n${hojas.length - mal}/${hojas.length} hojas bien. PDF del motor: ${SALIDA}`);
process.exit(mal ? 1 : 0);

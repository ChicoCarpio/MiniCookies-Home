// Pone lado a lado, como imagen, una pagina del PDF del motor propio y la del PDF de Excel (para mirarlas a ojo).
//   node comparar-xlsx/ver-pdf.mjs 27-impresion-extras Bordes [pagina] [escala]     → medidas/tmp/pdf-motor/<libro>/<hoja>.png
// Usa PDF.js de vendor/ en Edge. Antes: comparar-pdf.mjs para generar el PDF del motor.
import { createServer } from 'node:http';
import { spawn } from 'node:child_process';
import { readFileSync, writeFileSync, existsSync, mkdtempSync, rmSync, statSync } from 'node:fs';
import { join, extname, resolve } from 'node:path';
import { tmpdir } from 'node:os';

const RAIZ = resolve(import.meta.dirname, '..');
const CARPETA = join(RAIZ, 'comparar-xlsx');
const [libro, hoja, pagina = '1', escala = '1.4'] = process.argv.slice(2);
if (!libro || !hoja) { console.error('Uso: node comparar-xlsx/ver-pdf.mjs <libro> <hoja> [pagina] [escala]'); process.exit(1); }
const volcado = JSON.parse(readFileSync(join(CARPETA, 'medidas', 'paginas', libro + '.json'), 'utf8').replace(/^﻿/, ''));
const h = volcado.find((x) => x.nombre === hoja);
const PDF_MOTOR = join(CARPETA, 'medidas', 'tmp', 'pdf-motor', libro, encodeURIComponent(hoja) + '.pdf');
const PDF_EXCEL = join(CARPETA, 'medidas', 'tmp', 'pdf-excel', libro, h.pdf);
const SALIDA = join(CARPETA, 'medidas', 'tmp', 'pdf-motor', libro, encodeURIComponent(hoja) + '.png');
const EDGE = ['C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe', 'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe'].find((r) => existsSync(r));
const PAGINA = String.raw`<!doctype html><meta charset="utf-8"><body><script type="module">
import * as pdfjs from '/vendor/pdf.js/6.3.289/pdf.min.mjs';
pdfjs.GlobalWorkerOptions.workerSrc = '/vendor/pdf.js/6.3.289/pdf.worker.min.mjs';
const dibujar = async (url) => {
  const doc = await pdfjs.getDocument({ url, wasmUrl: '/vendor/pdf.js/6.3.289/wasm/' }).promise;
  const p = await doc.getPage(__PAGINA__);
  const vp = p.getViewport({ scale: __ESCALA__ });
  const c = document.createElement('canvas'); c.width = Math.ceil(vp.width); c.height = Math.ceil(vp.height);
  const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height);
  await p.render({ canvasContext: x, viewport: vp }).promise;
  return c;
};
try {
  const a = await dibujar('/__excel'), b = await dibujar('/__motor');
  const lado = document.createElement('canvas'); lado.width = a.width + b.width + 10; lado.height = Math.max(a.height, b.height);
  const l = lado.getContext('2d'); l.fillStyle = '#888'; l.fillRect(0, 0, lado.width, lado.height); l.drawImage(a, 0, 0); l.drawImage(b, a.width + 10, 0);
  const blob = await new Promise((ok) => lado.toBlob(ok, 'image/png'));
  await fetch('/__png', { method: 'POST', body: blob });
} catch (e) { await fetch('/__error', { method: 'POST', body: String(e && e.stack || e) }); }
</script>`;
let terminar; const fin = new Promise((r) => { terminar = r; });
const TIPOS = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.mjs': 'text/javascript', '.wasm': 'application/wasm', '.json': 'application/json' };
const servidor = createServer(async (q, s) => {
  const ruta = decodeURIComponent(q.url.split('?')[0]);
  if (q.method === 'POST') {
    const partes = []; for await (const p of q) partes.push(p);
    if (ruta === '/__png') writeFileSync(SALIDA, Buffer.concat(partes)); else console.error(Buffer.concat(partes).toString());
    s.end('ok'); terminar(); return;
  }
  if (ruta === '/') { s.setHeader('content-type', TIPOS['.html']); return s.end(PAGINA.replace('__PAGINA__', pagina).replace('__ESCALA__', escala)); }
  if (ruta === '/__excel') return s.end(readFileSync(PDF_EXCEL));
  if (ruta === '/__motor') return s.end(readFileSync(PDF_MOTOR));
  const p = join(RAIZ, ruta);
  if (!p.startsWith(RAIZ) || !existsSync(p) || statSync(p).isDirectory()) { s.statusCode = 404; return s.end('no'); }
  s.setHeader('content-type', TIPOS[extname(p)] || 'application/octet-stream');
  s.end(readFileSync(p));
}).listen(8796);
const perfil = mkdtempSync(join(tmpdir(), 'ver-pdf-'));
const edge = spawn(EDGE, ['--headless=new', '--disable-gpu', '--user-data-dir=' + perfil, 'http://localhost:8796/'], { stdio: 'ignore' });
const reloj = setTimeout(() => { console.error('Se paso el tiempo limite.'); terminar(); }, 120000);
await fin; clearTimeout(reloj); edge.kill(); servidor.close();
setTimeout(() => { try { rmSync(perfil, { recursive: true, force: true }); } catch { /* en uso */ } }, 1500);
console.log(existsSync(SALIDA) ? 'Imagen: ' + SALIDA : 'No se pudo generar la imagen');

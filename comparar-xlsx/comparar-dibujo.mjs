// Compara el dibujo de una hoja del motor propio (xlsx-dibujo.js, en Edge) con la imagen que da Excel real (exportar-rango.ps1).
//
//   node comparar-xlsx/comparar-dibujo.mjs                  todos los casos de casos-dibujo.json
//   node comparar-xlsx/comparar-dibujo.mjs 3 5              solo esos casos (posicion en la lista, desde 1)
//   --solo-referencia   solo (re)hace las imagenes de Excel;  --rehacer   ignora las imagenes guardadas
//
// Por caso: imagen de Excel (se guarda por SHA-256 del libro + hoja + rango en medidas/tmp/dibujo/ref), imagen del motor, porcentaje de
// pixeles distintos (algun canal difiere mas de 48/255; ruido de suavizado) y la hoja de contraste lado a lado en
// medidas/tmp/dibujo/salida/NN-nombre.png: [Excel | motor | diferencias]. Solo con archivos sinteticos (corpus).
// Necesita Windows con Excel y Edge: no va en "npm run build".
import { createServer } from 'node:http';
import { spawn, spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync, existsSync, mkdirSync, mkdtempSync, rmSync, statSync } from 'node:fs';
import { join, extname, resolve } from 'node:path';
import { tmpdir } from 'node:os';

const RAIZ = resolve(import.meta.dirname, '..');
const CARPETA = join(RAIZ, 'comparar-xlsx');
const TMP = join(CARPETA, 'medidas', 'tmp', 'dibujo');
const REF = join(TMP, 'ref'), SALIDA = join(TMP, 'salida');
mkdirSync(REF, { recursive: true }); mkdirSync(SALIDA, { recursive: true });
const PUERTO = 8794;
const EDGE = [process.env.EDGE, 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe', 'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe'].filter(Boolean).find((r) => existsSync(r));
const TIPOS = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.json': 'application/json', '.ttf': 'font/ttf', '.png': 'image/png', '.txt': 'text/plain' };

const args = process.argv.slice(2);
const opcion = (n) => args.includes(n);
const elegidos = args.filter((a) => /^\d+$/.test(a)).map(Number);
const todos = JSON.parse(readFileSync(join(CARPETA, 'casos-dibujo.json'), 'utf8'));
const casos = todos.map((c, i) => ({ ...c, n: i + 1 })).filter((c) => !elegidos.length || elegidos.includes(c.n));
if (!EDGE) { console.error('No se encontro Microsoft Edge.'); process.exit(1); }

// --- 1. Imagenes de Excel ---
for (const c of casos) {
  c.ruta = join(CARPETA, 'corpus', c.archivo);
  const sha = createHash('sha256').update(readFileSync(c.ruta)).update(`|${c.hoja}|${c.rango}|${c.cuadricula !== false}|${c.zoom || 100}`).digest('hex').slice(0, 16);
  c.ref = join(REF, `${String(c.n).padStart(2, '0')}-${sha}.png`);
  if (opcion('--rehacer') || !existsSync(c.ref)) {
    const r = spawnSync('powershell', ['-ExecutionPolicy', 'Bypass', '-File', join(CARPETA, 'exportar-rango.ps1'), '-Libro', c.ruta, '-Hoja', String(c.hoja), '-Rango', c.rango, '-Salida', c.ref, '-Zoom', String(c.zoom || 100), ...(c.cuadricula === false ? [] : ['-ConCuadricula'])], { encoding: 'utf8' });
    if (r.status !== 0 || !existsSync(c.ref)) { console.error(`Caso ${c.n}: fallo exportar-rango.ps1\n${r.stdout}${r.stderr}`); process.exit(1); }
  }
}
if (opcion('--solo-referencia')) { console.log(`${casos.length} imagenes de Excel listas en ${REF}`); process.exit(0); }

// --- 2. El motor propio, en Edge ---
const CONTROL = String.raw`<!doctype html><meta charset="utf-8"><body><script>
const CASOS = __CASOS__;
const enviar = (ruta, cuerpo) => fetch(ruta, { method: 'POST', body: cuerpo });
const cargar = (src) => new Promise((ok, mal) => { const s = document.createElement('script'); s.src = src; s.onload = ok; s.onerror = () => mal(new Error('no cargo ' + src)); document.head.appendChild(s); });
const imagen = (src) => new Promise((ok, mal) => { const i = new Image(); i.onload = () => ok(i); i.onerror = () => mal(new Error('no cargo ' + src)); i.src = src; });
const aPng = (c) => new Promise((ok) => c.toBlob(ok, 'image/png'));
(async () => {
  for (const s of ['vendor.js', 'zip-propio.js', 'xml-ligero.js', 'xlsx-estilos.js', 'xlsx-modelo.js', 'xlsx-formato.js', 'xlsx-geometria.js', 'pptx-texto.js', 'xlsx-dibujo.js']) await cargar('/assets/js/' + s);
  XlsxDibujo.config.avance = +new URLSearchParams(location.search).get('avance') || 0;
  for (const caso of CASOS) {
    const r = { n: caso.n, nombre: caso.nombre };
    try {
      const libro = await LibroXlsx.abrir(new Uint8Array(await (await fetch('/__xlsx/' + caso.n)).arrayBuffer()));
      await XlsxDibujo.cargarFuentes(libro);
      const hoja = await libro.cargarHoja(caso.hoja - 1);
      const g = XlsxGeometria.crear(libro, hoja, { medirDigito: XlsxDibujo.medirDigito });
      const rg = LibroXlsx.decodificarRango(caso.rango);
      const ref = await imagen('/__ref/' + caso.n);
      const W = ref.naturalWidth, H = ref.naturalHeight;
      const mio = document.createElement('canvas'); mio.width = W; mio.height = H;
      const cm = mio.getContext('2d', { willReadFrequently: true });
      cm.fillStyle = '#fff'; cm.fillRect(0, 0, W, H);
      XlsxDibujo.dibujarRango(cm, libro, hoja, g, rg, { x: 0, y: 0, idioma: 'es', cuadricula: caso.cuadricula !== false });
      const cr = document.createElement('canvas'); cr.width = W; cr.height = H;
      const cx = cr.getContext('2d', { willReadFrequently: true });
      cx.fillStyle = '#fff'; cx.fillRect(0, 0, W, H); cx.drawImage(ref, 0, 0);
      const a = cr.getContext('2d').getImageData(0, 0, W, H).data, b = cm.getImageData(0, 0, W, H).data;
      const dif = document.createElement('canvas'); dif.width = W; dif.height = H;
      const cd = dif.getContext('2d'); const di = cd.createImageData(W, H);
      // Comparacion tolerante al suavizado del texto: luminancia con desenfoque de 3x3 (Excel pinta con ClearType, el Canvas en grises).
      const lum = (d) => { const l = new Float32Array(W * H); for (let i = 0, k = 0; i < d.length; i += 4, k++) l[k] = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2]; return l; };
      const borroso = (l) => { const o = new Float32Array(W * H); for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) { let t = 0, n = 0; for (let j = -1; j <= 1; j++) for (let i = -1; i <= 1; i++) { const xx = x + i, yy = y + j; if (xx >= 0 && yy >= 0 && xx < W && yy < H) { t += l[yy * W + xx]; n++; } } o[y * W + x] = t / n; } return o; };
      const la = borroso(lum(a)), lb = borroso(lum(b));
      let malos = 0, tinta = 0;
      for (let k = 0, i = 0; k < W * H; k++, i += 4) {
        const d = Math.max(Math.abs(a[i] - b[i]), Math.abs(a[i + 1] - b[i + 1]), Math.abs(a[i + 2] - b[i + 2]));
        const sinTexto = d > 48 && Math.abs(la[k] - lb[k]) <= 40 && Math.abs(a[i] - b[i]) < 255; // diferencia solo de suavizado
        const blancoA = a[i] > 240 && a[i + 1] > 240 && a[i + 2] > 240, blancoB = b[i] > 240 && b[i + 1] > 240 && b[i + 2] > 240;
        if (!blancoA || !blancoB) tinta++;
        const mal = d > 48 && !sinTexto;
        if (mal) { malos++; di.data[i] = 255; di.data[i + 1] = 0; di.data[i + 2] = 0; di.data[i + 3] = 255; }
        else { const gris = 255 - Math.min(80, d); di.data[i] = di.data[i + 1] = di.data[i + 2] = Math.max(gris, 200); di.data[i + 3] = 255; }
      }
      cd.putImageData(di, 0, 0);
      const lado = document.createElement('canvas'); lado.width = W * 3 + 20; lado.height = H;
      const cl = lado.getContext('2d'); cl.fillStyle = '#888'; cl.fillRect(0, 0, lado.width, H);
      cl.drawImage(cr, 0, 0); cl.drawImage(mio, W + 10, 0); cl.drawImage(dif, 2 * W + 20, 0);
      await enviar('/__img/' + caso.n, await aPng(lado));
      r.ancho = W; r.alto = H; r.pct = +(100 * malos / (W * H)).toFixed(2); r.pctTinta = tinta ? +(100 * malos / tinta).toFixed(1) : 0;
    } catch (e) { r.error = String(e && e.stack || e); }
    await enviar('/__resultado', JSON.stringify(r));
  }
  await enviar('/__fin', '1');
})().catch((e) => enviar('/__fin', 'ERROR ' + (e && e.stack || e)));
</script>`;

const resultados = [];
let terminar; const fin = new Promise((r) => { terminar = r; });
const servidor = createServer(async (q, s) => {
  const ruta = decodeURIComponent(q.url.split('?')[0]);
  if (q.method === 'POST') {
    const partes = []; for await (const p of q) partes.push(p);
    const cuerpo = Buffer.concat(partes);
    if (ruta.startsWith('/__img/')) {
      const c = casos.find((x) => x.n === +ruta.split('/')[2]);
      writeFileSync(join(SALIDA, `${String(c.n).padStart(2, '0')}-${c.nombre}.png`), cuerpo);
    } else if (ruta === '/__resultado') {
      const r = JSON.parse(cuerpo.toString()); resultados.push(r);
      console.log(r.error ? `  ${r.n} ${r.nombre}: ERROR ${r.error.split('\n').slice(0, 3).join(' | ')}` : `  ${String(r.n).padStart(2)} ${r.nombre.padEnd(28)} ${r.ancho}x${r.alto}  ${String(r.pct).padStart(5)} % de la imagen, ${String(r.pctTinta).padStart(5)} % de lo pintado`);
    } else if (ruta === '/__fin') { const t = cuerpo.toString(); if (t !== '1') console.error(t); terminar(); }
    s.end('ok'); return;
  }
  if (ruta === '/__control') {
    s.setHeader('content-type', 'text/html; charset=utf-8');
    return s.end(CONTROL.replace('__CASOS__', JSON.stringify(casos.map((c) => ({ n: c.n, nombre: c.nombre, hoja: c.hoja, rango: c.rango, cuadricula: c.cuadricula })))));
  }
  if (ruta.startsWith('/__xlsx/')) return s.end(readFileSync(casos.find((x) => x.n === +ruta.split('/')[2]).ruta));
  if (ruta.startsWith('/__ref/')) { s.setHeader('content-type', 'image/png'); return s.end(readFileSync(casos.find((x) => x.n === +ruta.split('/')[2]).ref)); }
  const p = join(RAIZ, ruta);
  if (!p.startsWith(RAIZ) || !existsSync(p) || statSync(p).isDirectory()) { s.statusCode = 404; return s.end('no'); }
  s.setHeader('content-type', TIPOS[extname(p)] || 'application/octet-stream');
  s.end(readFileSync(p));
}).listen(PUERTO);

const perfil = mkdtempSync(join(tmpdir(), 'comparar-xlsx-'));
const edge = spawn(EDGE, ['--headless=new', '--disable-gpu', '--lang=es-CR', '--user-data-dir=' + perfil, '--window-size=1300,950', `http://localhost:${PUERTO}/__control?avance=${process.env.AVANCE || 0}`], { stdio: 'ignore' });
const reloj = setTimeout(() => { console.error('Se paso el tiempo limite.'); terminar(); }, 10 * 60 * 1000);
await fin;
clearTimeout(reloj); edge.kill(); servidor.close();
setTimeout(() => { try { rmSync(perfil, { recursive: true, force: true }); } catch { /* Edge todavia lo usa */ } }, 1500);
const ok = resultados.filter((r) => !r.error);
console.log(`\n${ok.length}/${casos.length} casos; media ${ok.length ? (ok.reduce((a, r) => a + r.pct, 0) / ok.length).toFixed(2) : '-'} % de la imagen distinta. Hojas de contraste: ${SALIDA}`);
process.exit(ok.length === casos.length ? 0 : 1);

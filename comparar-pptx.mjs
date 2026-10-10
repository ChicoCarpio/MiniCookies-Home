/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. */
// Compara el motor de diapositivas del sitio con PowerPoint real.
//
// Por qué existe: la fidelidad de PowerPoint → PDF / imágenes no se puede
// juzgar a ojo ni con una prueba de "no hubo error". El 2026-10-09 se midió a
// mano (arnés en la carpeta temporal) que las 22 diapositivas de una charla
// real salían infieles; con este comando la medición queda fija y repetible,
// y cada etapa del motor propio se aprueba contra PowerPoint, no contra la
// impresión de quien la mira. Ver investigacion-motor-powerpoint\00-resumen.md
// (en la raíz del proyecto, fuera del repo).
//
// Qué hace, por cada .pptx:
//  1. Referencia: PowerPoint real por COM (comparar-pptx\powerpoint.ps1)
//     exporta cada diapositiva a PNG y escribe la caja de cada forma y de cada
//     línea de texto. Se guarda por SHA-256 del archivo: solo se rehace si el
//     archivo cambió.
//  2. Nuestro motor: el sitio compilado en Edge, sin internet, dibuja cada
//     diapositiva con renderizarDiapositivasComoImagenes (lo mismo que usan
//     PowerPoint → PDF, PowerPoint → imágenes y Unir PDF).
//  3. Mide por diapositiva el % de píxeles distintos (algún canal difiere más
//     de 48/255, a 960 px de ancho: el ruido de suavizado medido es ~1,3 %) y
//     el SSIM, arma comparar-pptx\salida\informe.html (lado a lado y mapa de
//     diferencias) y compara con comparar-pptx\linea-base.json.
//
// Uso:
//   npm run comparar-pptx                     corpus\*.pptx y locales\*.pptx
//   npm run comparar-pptx -- a.pptx b.pptx    esos archivos
// Opciones: --sin-preparar (el original sin pptx-preparar.js, para medir
// cuánto aporta), --motor-propio (el motor nuevo, etapa 2 en adelante), --base (guarda el resultado como nueva línea base),
// --sin-build, --pdf (exporta también el PDF de PowerPoint).
// Código de salida 1 si una diapositiva empeoró respecto de la línea base
// (más de 1 punto de % de píxeles o más de 0,01 de SSIM) o si faltó dibujar
// alguna. Necesita Windows con PowerPoint y Edge: NO conectar a "npm run
// build" (Cloudflare compila en Linux).

import { createServer } from 'node:http';
import { spawn, spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync, existsSync, readdirSync, mkdirSync, mkdtempSync, rmSync, statSync } from 'node:fs';
import { join, extname, basename, resolve } from 'node:path';
import { tmpdir } from 'node:os';

const RAIZ = decodeURIComponent(new URL('.', import.meta.url).pathname).replace(/^\/([A-Za-z]:)/, '$1');
const DIST = join(RAIZ, 'dist');
const CARPETA = join(RAIZ, 'comparar-pptx');
const SALIDA = join(CARPETA, 'salida');
const BASE = join(CARPETA, 'linea-base.json');
const PUERTO = 8793;
const LIMITE_MS = 60 * 60 * 1000;
const MARGEN_PCT = 1.0;
const MARGEN_SSIM = 0.01;
const EDGE = [process.env.EDGE,
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe'].filter(Boolean).find((r) => existsSync(r));
const TIPOS = {
  '.html': 'text/html; charset=utf-8', '.css': 'text/css; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8', '.json': 'application/json', '.png': 'image/png', '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml', '.wasm': 'application/wasm', '.ttf': 'font/ttf', '.woff2': 'font/woff2',
  '.gz': 'application/gzip', '.txt': 'text/plain; charset=utf-8', '.xml': 'application/xml',
};

const argumentos = process.argv.slice(2);
const opcion = (n) => argumentos.includes(n);
const sinPreparar = opcion('--sin-preparar');
const motorPropio = opcion('--motor-propio');
let archivos = argumentos.filter((a) => !a.startsWith('--')).map((a) => resolve(a));
// Sin archivos: el corpus sintético (en git) y comparar-pptx\locales (fuera de
// git: presentaciones reales que no se pueden publicar, como la charla del
// usuario).
if (!archivos.length) {
  for (const carpeta of ['corpus', 'locales'].map((c) => join(CARPETA, c))) {
    if (existsSync(carpeta)) archivos.push(...readdirSync(carpeta).filter((n) => /\.pptx$/i.test(n)).sort().map((n) => join(carpeta, n)));
  }
}
if (!archivos.length) { console.error('No hay .pptx para comparar (comparar-pptx\\corpus está vacío).'); process.exit(1); }
if (!EDGE) { console.error('No se encontró Microsoft Edge.'); process.exit(1); }

// --- 1. Referencias de PowerPoint ------------------------------------------
const casos = archivos.map((ruta) => {
  const sha = createHash('sha256').update(readFileSync(ruta)).digest('hex');
  return { ruta, nombre: basename(ruta), sha, ref: join(SALIDA, 'referencias', sha.slice(0, 16)) };
});
for (const c of casos) {
  const json = join(c.ref, 'referencia.json');
  if (existsSync(json) && (!opcion('--pdf') || existsSync(join(c.ref, 'powerpoint.pdf')))) continue;
  console.log(`PowerPoint: ${c.nombre}…`);
  const r = spawnSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', join(CARPETA, 'powerpoint.ps1'),
    '-Archivo', c.ruta, '-Salida', c.ref, ...(opcion('--pdf') ? ['-Pdf'] : [])], { encoding: 'utf8', timeout: 15 * 60 * 1000 });
  if (r.status !== 0 || !existsSync(json)) {
    console.error(`  falló:\n${r.stdout}\n${r.stderr}`);
    process.exit(1);
  }
  console.log('  ' + r.stdout.trim().split('\n').pop());
}
for (const c of casos) c.referencia = JSON.parse(readFileSync(join(c.ref, 'referencia.json'), 'utf8').replace(/^\uFEFF/, ''));

// --- 2. Nuestro motor, en Edge ---------------------------------------------
if (!opcion('--sin-build')) {
  const b = spawnSync(process.execPath, [join(RAIZ, 'build.mjs')], { stdio: 'inherit' });
  if (b.status !== 0) process.exit(1);
}
const corrida = sinPreparar ? 'sin-preparar' : motorPropio ? 'propio' : 'actual';
const destino = join(SALIDA, corrida);
rmSync(destino, { recursive: true, force: true });

// En el navegador: por cada archivo, dibuja, compara con la referencia y manda
// las imágenes y las medidas al servidor.
const CONTROL = String.raw`<!doctype html><meta charset="utf-8"><iframe id="m" style="width:1280px;height:900px"></iframe><script>
const CASOS = __CASOS__;
const SIN_PREPARAR = __SIN_PREPARAR__;
const MOTOR_PROPIO = __MOTOR_PROPIO__;
const enviar = (ruta, cuerpo) => fetch(ruta, { method: 'POST', body: cuerpo });
const imagen = (src) => new Promise((ok, mal) => { const i = new Image(); i.onload = () => ok(i); i.onerror = () => mal(new Error('no cargó ' + src)); i.src = src; });
function lienzo(fuente, w, h) {
  const c = document.createElement('canvas'); c.width = w; c.height = h;
  const x = c.getContext('2d', { willReadFrequently: true });
  x.fillStyle = '#fff'; x.fillRect(0, 0, w, h);
  x.imageSmoothingQuality = 'high'; x.drawImage(fuente, 0, 0, w, h);
  return c;
}
// SSIM sobre la luminancia, en ventanas de 8×8 sin solapar.
function ssim(a, b, w, h) {
  const L = (d, i) => 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
  const C1 = (0.01 * 255) ** 2, C2 = (0.03 * 255) ** 2;
  let suma = 0, n = 0;
  for (let y = 0; y + 8 <= h; y += 8) for (let x = 0; x + 8 <= w; x += 8) {
    let ma = 0, mb = 0, va = 0, vb = 0, cv = 0;
    for (let j = 0; j < 8; j++) for (let i = 0; i < 8; i++) { const k = ((y + j) * w + x + i) * 4; ma += L(a, k); mb += L(b, k); }
    ma /= 64; mb /= 64;
    for (let j = 0; j < 8; j++) for (let i = 0; i < 8; i++) {
      const k = ((y + j) * w + x + i) * 4; const da = L(a, k) - ma, db = L(b, k) - mb;
      va += da * da; vb += db * db; cv += da * db;
    }
    va /= 63; vb /= 63; cv /= 63;
    suma += ((2 * ma * mb + C1) * (2 * cv + C2)) / ((ma * ma + mb * mb + C1) * (va + vb + C2)); n++;
  }
  return n ? suma / n : 1;
}
(async () => {
  const m = document.getElementById('m');
  await new Promise((r) => { m.onload = r; m.src = '/herramientas/powerpoint-a-imagenes.html'; });
  const W = m.contentWindow;
  const avisos = [];
  for (const nivel of ['warn', 'error']) { const o = W.console[nivel]; W.console[nivel] = (...a) => { avisos.push(nivel + ': ' + a.map(String).join(' ').slice(0, 300)); o.apply(W.console, a); }; }
  for (const caso of CASOS) {
    const t0 = performance.now();
    const resultado = { nombre: caso.nombre, sha: caso.sha, diapositivas: [], avisos: [] };
    try {
      const bytes = await (await fetch('/__pptx/' + caso.indice)).arrayBuffer();
      const archivo = new W.File([bytes], caso.nombre);
      const lienzoRender = W.document.createElement('canvas');
      avisos.length = 0;
      let imagenes;
      if (SIN_PREPARAR) {
        imagenes = await W.dibujarDiapositivas(archivo, archivo, lienzoRender, null, { conImagenes: true });
      } else {
        resultado.cambios = (await W.eval('PptxPreparar').preparar(archivo)).cambios;
        imagenes = await W.renderizarDiapositivasComoImagenes(archivo, lienzoRender, null, { conImagenes: true, motor: MOTOR_PROPIO ? 'propio' : undefined });
      }
      resultado.ms = Math.round(performance.now() - t0);
      for (let i = 0; i < caso.total; i++) {
        const n = String(i + 1).padStart(3, '0');
        const ref = await imagen('/__ref/' + caso.indice + '/d' + n + '.png');
        const w = 960, h = Math.round(960 * ref.height / ref.width);
        const cRef = lienzo(ref, w, h);
        const fila = { n: i + 1 };
        if (!imagenes[i]) { fila.falta = true; resultado.diapositivas.push(fila); continue; }
        if (imagenes[i].motor) fila.motor = imagenes[i].motor;
        if (imagenes[i].motivos) fila.motivos = imagenes[i].motivos;
        const nuestra = await imagen(imagenes[i].url);
        const cNos = lienzo(nuestra, w, h);
        const a = cRef.getContext('2d').getImageData(0, 0, w, h).data;
        const b = cNos.getContext('2d').getImageData(0, 0, w, h).data;
        const dif = new ImageData(w, h);
        let distintos = 0;
        for (let k = 0; k < a.length; k += 4) {
          const d = Math.max(Math.abs(a[k] - b[k]), Math.abs(a[k + 1] - b[k + 1]), Math.abs(a[k + 2] - b[k + 2]));
          if (d > 48) { distintos++; dif.data[k] = 220; dif.data[k + 1] = 0; dif.data[k + 2] = 0; dif.data[k + 3] = 255; }
          else { const v = 255 - Math.min(80, d * 2); dif.data[k] = dif.data[k + 1] = dif.data[k + 2] = v; dif.data[k + 3] = 255; }
        }
        fila.pct = Math.round(10000 * distintos / (w * h)) / 100;
        fila.ssim = Math.round(10000 * ssim(a, b, w, h)) / 10000;
        const cDif = document.createElement('canvas'); cDif.width = w; cDif.height = h; cDif.getContext('2d').putImageData(dif, 0, 0);
        const png = (c) => new Promise((r) => c.toBlob(r, 'image/png'));
        await enviar('/__img/' + caso.indice + '/n' + n + '.png', await png(lienzo(nuestra, ref.width, ref.height)));
        await enviar('/__img/' + caso.indice + '/x' + n + '.png', await png(cDif));
        resultado.diapositivas.push(fila);
        W.URL.revokeObjectURL(imagenes[i].url);
      }
      if (imagenes.length !== caso.total) resultado.avisos.push('nuestro motor dibujó ' + imagenes.length + ' diapositivas y PowerPoint ' + caso.total);
    } catch (e) {
      resultado.error = String(e && e.stack || e);
    }
    resultado.avisos.push(...avisos);
    await enviar('/__resultado', JSON.stringify(resultado));
  }
  await enviar('/__fin', '1');
})().catch((e) => enviar('/__fin', 'ERROR ' + (e && e.stack || e)));
</script>`;

const resultados = [];
let terminar;
const fin = new Promise((r) => { terminar = r; });
const datosCasos = casos.map((c, indice) => ({ indice, nombre: c.nombre, sha: c.sha, total: c.referencia.diapositivas.length }));
const servidor = createServer(async (q, s) => {
  const ruta = decodeURIComponent(q.url.split('?')[0]);
  if (q.method === 'POST') {
    const partes = []; for await (const p of q) partes.push(p);
    const cuerpo = Buffer.concat(partes);
    if (ruta.startsWith('/__img/')) {
      const [, , indice, nombre] = ruta.split('/');
      const carpeta = join(destino, casos[+indice].nombre.replace(/\.pptx$/i, ''));
      mkdirSync(carpeta, { recursive: true });
      writeFileSync(join(carpeta, nombre), cuerpo);
    } else if (ruta === '/__resultado') {
      const r = JSON.parse(cuerpo.toString());
      resultados.push(r);
      const ps = r.diapositivas.filter((d) => d.pct != null);
      const media = ps.length ? (ps.reduce((a, d) => a + d.pct, 0) / ps.length).toFixed(2) : '—';
      const ajenas = r.diapositivas.filter((d) => d.motor === 'pptxviewjs').length;
      console.log(`  ${r.nombre}: ${ps.length} diapositivas${ajenas ? ` (${ajenas} con PptxViewJS)` : ''}, media ${media} %${r.error ? ' — ERROR ' + r.error.split('\n')[0] : ''}`);
    } else if (ruta === '/__fin') {
      const t = cuerpo.toString(); if (t !== '1') console.error(t); terminar();
    }
    s.end('ok'); return;
  }
  if (ruta === '/__control') {
    s.setHeader('content-type', 'text/html; charset=utf-8');
    return s.end(CONTROL.replace('__CASOS__', JSON.stringify(datosCasos)).replace('__SIN_PREPARAR__', String(sinPreparar)).replace('__MOTOR_PROPIO__', String(motorPropio)));
  }
  if (ruta.startsWith('/__pptx/')) return s.end(readFileSync(casos[+ruta.split('/')[2]].ruta));
  if (ruta.startsWith('/__ref/')) {
    const [, , indice, nombre] = ruta.split('/');
    s.setHeader('content-type', 'image/png');
    return s.end(readFileSync(join(casos[+indice].ref, nombre)));
  }
  let p = join(DIST, ruta);
  if (existsSync(p) && statSync(p).isDirectory()) p = join(p, 'index.html');
  if (!existsSync(p) && existsSync(p + '.html')) p += '.html';
  if (!existsSync(p)) { s.statusCode = 404; return s.end('no'); }
  s.setHeader('content-type', TIPOS[extname(p)] || 'application/octet-stream');
  s.end(readFileSync(p));
}).listen(PUERTO);

const perfil = mkdtempSync(join(tmpdir(), 'comparar-pptx-'));
const edge = spawn(EDGE, ['--headless=new', '--disable-gpu', '--user-data-dir=' + perfil,
  '--host-resolver-rules=MAP * ~NOTFOUND , EXCLUDE localhost', '--window-size=1300,950',
  `http://localhost:${PUERTO}/__control`], { stdio: 'ignore' });
console.log(`Nuestro motor (${corrida}): ${casos.length} archivo(s)…`);
const reloj = setTimeout(() => { console.error('Se pasó el tiempo límite.'); terminar(); }, LIMITE_MS);
await fin;
clearTimeout(reloj);
edge.kill();
servidor.close();
setTimeout(() => { try { rmSync(perfil, { recursive: true, force: true }); } catch { /* Edge todavía lo usa */ } }, 1500);

// --- 3. Informe y línea base ----------------------------------------------
const base = existsSync(BASE) ? JSON.parse(readFileSync(BASE, 'utf8')) : {};
const peores = [];
let fallas = 0;
for (const r of resultados) {
  const anterior = base[r.nombre];
  for (const d of r.diapositivas) {
    if (d.falta || d.pct == null) { fallas++; continue; }
    const a = anterior && anterior.sha === r.sha ? anterior.diapositivas[d.n - 1] : null;
    if (a) {
      d.antes = a;
      if (d.pct > a.pct + MARGEN_PCT || d.ssim < a.ssim - MARGEN_SSIM) peores.push(`${r.nombre} diapositiva ${d.n}: ${a.pct} → ${d.pct} %, SSIM ${a.ssim} → ${d.ssim}`);
    }
  }
  if (r.error) fallas++;
}
const todas = resultados.flatMap((r) => r.diapositivas.filter((d) => d.pct != null));
const media = (k) => todas.length ? (todas.reduce((a, d) => a + d[k], 0) / todas.length) : 0;
writeFileSync(join(destino, 'resultado.json'), JSON.stringify(resultados, null, 1));
escribirInforme();
if (opcion('--base') && !sinPreparar) {
  const nueva = { ...base };
  for (const r of resultados) {
    if (r.error) continue;
    nueva[r.nombre] = { sha: r.sha, fecha: new Date().toISOString().slice(0, 10), diapositivas: r.diapositivas.map((d) => ({ pct: d.pct, ssim: d.ssim })) };
  }
  writeFileSync(BASE, JSON.stringify(nueva, null, 1) + '\n');
  console.log('Línea base actualizada: comparar-pptx\\linea-base.json');
}

console.log(`\n${todas.length} diapositivas: media ${media('pct').toFixed(2)} % de píxeles distintos, SSIM ${media('ssim').toFixed(4)}`);
console.log('Informe: ' + join(destino, 'informe.html'));
if (peores.length) console.log('\nEmpeoraron respecto de la línea base:\n  ' + peores.join('\n  '));
if (fallas) console.log(`\n${fallas} diapositiva(s) o archivo(s) sin dibujar o con error (ver el informe).`);
process.exit(peores.length || fallas || resultados.length !== casos.length ? 1 : 0);

function escribirInforme() {
  const esc = (t) => String(t).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const relRef = (c, n) => `../referencias/${c.sha.slice(0, 16)}/d${n}.png`;
  let html = `<!doctype html><meta charset="utf-8"><title>Comparar PowerPoint — ${corrida}</title>
<style>body{font:14px system-ui,sans-serif;margin:16px;background:#f6f6f6;color:#222}table{border-collapse:collapse;margin:8px 0 24px}
td,th{border:1px solid #ccc;padding:4px 8px;text-align:right;background:#fff}th{background:#eee}td.t{text-align:left}
.fila{display:flex;gap:6px;margin:6px 0 18px}.fila figure{margin:0;flex:1}.fila img{width:100%;border:1px solid #bbb;background:#fff}
figcaption{font-size:12px;color:#555}.peor{color:#b00;font-weight:600}.mejor{color:#070;font-weight:600}h2{margin-top:32px}</style>
<h1>Motor de diapositivas vs. PowerPoint (${corrida})</h1>
<p>${todas.length} diapositivas · media ${media('pct').toFixed(2)} % de píxeles distintos · SSIM ${media('ssim').toFixed(4)} · ${new Date().toLocaleString('es')}</p>`;
  html += '<table><tr><th class="t">Archivo</th><th>Diap.</th><th>Media %</th><th>SSIM</th><th>ms</th><th class="t">Avisos</th></tr>';
  for (const r of resultados) {
    const ds = r.diapositivas.filter((d) => d.pct != null);
    const m = (k) => ds.length ? (ds.reduce((a, d) => a + d[k], 0) / ds.length) : 0;
    html += `<tr><td class="t"><a href="#${esc(r.nombre)}">${esc(r.nombre)}</a></td><td>${ds.length}</td><td>${m('pct').toFixed(2)}</td><td>${m('ssim').toFixed(4)}</td><td>${r.ms || ''}</td><td class="t">${esc([r.error, ...r.avisos].filter(Boolean).join(' · ')).slice(0, 400)}</td></tr>`;
  }
  html += '</table>';
  for (const r of resultados) {
    const c = casos.find((x) => x.nombre === r.nombre);
    const carpeta = encodeURIComponent(r.nombre.replace(/\.pptx$/i, ''));
    html += `<h2 id="${esc(r.nombre)}">${esc(r.nombre)}</h2>`;
    if (r.cambios) html += `<p>pptx-preparar: ${esc(JSON.stringify(r.cambios))}</p>`;
    for (const d of r.diapositivas) {
      const n = String(d.n).padStart(3, '0');
      const delta = d.antes ? d.pct - d.antes.pct : 0;
      const clase = delta > MARGEN_PCT ? 'peor' : delta < -MARGEN_PCT ? 'mejor' : '';
      html += `<p><b>Diapositiva ${d.n}</b> — ${d.falta ? 'NO SE DIBUJÓ' : `${d.pct} % · SSIM ${d.ssim}`}${d.motor === 'pptxviewjs' ? ` · <i>PptxViewJS (${esc((d.motivos || []).join(', '))})</i>` : ''}${d.antes ? ` <span class="${clase}">(antes ${d.antes.pct} %)</span>` : ''}</p>`;
      if (!d.falta) html += `<div class="fila"><figure><img loading="lazy" src="${relRef(c, n)}"><figcaption>PowerPoint</figcaption></figure><figure><img loading="lazy" src="${carpeta}/n${n}.png"><figcaption>Nuestro motor</figcaption></figure><figure><img loading="lazy" src="${carpeta}/x${n}.png"><figcaption>Diferencias (rojo)</figcaption></figure></div>`;
    }
  }
  writeFileSync(join(destino, 'informe.html'), html);
}

// Compara el lector propio (xlsx-modelo.js) con SheetJS 0.20.3 en cada libro: hojas, celdas, valores, fórmulas, combinadas.
// Uso: node comparar-xlsx/comparar-modelo.mjs [archivo-o-carpeta …]   (por defecto: corpus\)
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import vm from 'node:vm';
const raiz = path.resolve(import.meta.dirname, '..');
const req = createRequire(import.meta.url);
for (const f of ['zip-propio.js', 'xml-ligero.js', 'xlsx-estilos.js', 'xlsx-modelo.js']) vm.runInThisContext(fs.readFileSync(path.join(raiz, 'assets/js', f), 'utf8'), { filename: f });
const XLSX = req(path.join(raiz, 'vendor/xlsx/0.20.3/xlsx.full.min.js'));

const entradas = process.argv.slice(2);
if (!entradas.length) entradas.push(path.join(import.meta.dirname, 'corpus'));
const archivos = [];
for (const e of entradas) {
  if (fs.statSync(e).isDirectory()) fs.readdirSync(e).filter((n) => /\.(xlsx|xlsm)$/i.test(n)).forEach((n) => archivos.push(path.join(e, n)));
  else archivos.push(e);
}

let fallos = 0;
for (const archivo of archivos) {
  const nombre = path.basename(archivo);
  const bytes = fs.readFileSync(archivo);
  const problemas = [];
  let t0 = performance.now();
  let libro;
  try { libro = await LibroXlsx.abrir(new Uint8Array(bytes)); } catch (e) { console.log('✗', nombre, 'ABRIR:', e.message, e.detalle || ''); fallos++; continue; }
  const sj = XLSX.read(bytes, { cellDates: false, cellFormula: true, cellStyles: false, dense: false });
  if (sj.SheetNames.length !== libro.hojas.length) problemas.push('hojas: ' + sj.SheetNames.length + ' vs ' + libro.hojas.length);
  let celdas = 0;
  for (let i = 0; i < libro.hojas.length; i++) {
    const meta = libro.hojas[i];
    if (sj.SheetNames[i] !== meta.nombre) problemas.push('nombre hoja ' + i + ': ' + sj.SheetNames[i] + ' vs ' + meta.nombre);
    const hoja = await libro.cargarHoja(i);
    const h = sj.Sheets[sj.SheetNames[i]];
    if (!h) continue;
    const propias = new Map();
    hoja.celdas.paraCada((f, c, k) => {
      if (hoja.celdas.tipo[k] === 0 && !hoja.celdas.formulas.has(k)) return; // vacía con estilo: SheetJS la omite sin sheetStubs
      propias.set(LibroXlsx.codificarDireccion(f, c), k);
    });
    const ajenas = Object.keys(h).filter((k) => k[0] !== '!');
    celdas += ajenas.length;
    for (const dir of ajenas) {
      const sc = h[dir];
      if (sc.t === 'z') continue; // SheetJS crea una celda vacía donde hay un comentario
      const k = propias.get(dir);
      if (k === undefined) { if (problemas.length < 12) problemas.push(meta.nombre + '!' + dir + ' falta en el propio (SheetJS: ' + JSON.stringify(sc.v) + ')'); continue; }
      const tipo = hoja.celdas.tipo[k], valor = hoja.celdas.valor[k];
      let esperado = sc.v;
      if (sc.t === 'n' && (tipo !== LibroXlsx.TIPO.numero || valor !== esperado)) problemas.push(meta.nombre + '!' + dir + ' número ' + valor + ' vs ' + esperado);
      else if (sc.t === 's' && (tipo !== LibroXlsx.TIPO.texto || valor !== esperado)) problemas.push(meta.nombre + '!' + dir + ' texto ' + JSON.stringify(valor) + ' vs ' + JSON.stringify(esperado));
      else if (sc.t === 'b' && (tipo !== LibroXlsx.TIPO.logico || valor !== esperado)) problemas.push(meta.nombre + '!' + dir + ' lógico');
      else if (sc.t === 'e' && tipo !== LibroXlsx.TIPO.error) problemas.push(meta.nombre + '!' + dir + ' error ' + valor + ' vs ' + sc.w);
      const fx = (hoja.celdas.formulaDe(k) || '').replace(/_xlfn\.(_xlws\.)?|_xlws\./g, '') || null; // SheetJS quita los prefijos _xlfn.
      const fsj = sc.f === undefined ? undefined : sc.f.replace(/\r\n?/g, '\n');
      if (fsj !== undefined && fx !== null && fsj !== fx && !(hoja.celdas.formulas.get(k).tipo === 'array' || hoja.celdas.formulas.get(k).tipo === 'shared')) { let d = 0; while (fx[d] === fsj[d]) d++; problemas.push(meta.nombre + '!' + dir + ' fórmula distinta en ' + d + ': ' + JSON.stringify(fx.slice(Math.max(0, d - 15), d + 25)) + ' vs ' + JSON.stringify(fsj.slice(Math.max(0, d - 15), d + 25))); }
      else if (sc.f !== undefined && fx === null) problemas.push(meta.nombre + '!' + dir + ' sin fórmula (SheetJS: ' + sc.f + ')');
    }
    for (const dir of propias.keys()) if (!(dir in h) ) { if (hoja.celdas.tipo[propias.get(dir)] !== 0 && problemas.length < 12) problemas.push(meta.nombre + '!' + dir + ' sobra en el propio'); }
    const mg = (h['!merges'] || []).length;
    if (mg !== hoja.combinadas.length) problemas.push(meta.nombre + ' combinadas: ' + mg + ' vs ' + hoja.combinadas.length);
  }
  const ms = (performance.now() - t0).toFixed(0);
  if (problemas.length) { fallos++; console.log('✗', nombre, `(${celdas} celdas, ${ms} ms)`); problemas.slice(0, 12).forEach((p) => console.log('    ', p)); } else console.log('✓', nombre, `${libro.hojas.length} hojas, ${celdas} celdas, ${ms} ms`);
}
console.log(fallos ? `\n${fallos} libro(s) con diferencias` : '\nTodo igual que SheetJS');
process.exit(fallos ? 1 : 0);

// Compara el lector de CSV propio (csv-planilla.js) con lo que hacían las herramientas con SheetJS (lector-planillas.js).
// Uso: node comparar-xlsx/comparar-csv.mjs [archivo.csv …]   (sin argumentos: unos casos sintéticos difíciles)
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import vm from 'node:vm';
const raiz = path.resolve(import.meta.dirname, '..');
const req = createRequire(import.meta.url);
for (const f of ['zip-propio.js', 'xml-ligero.js', 'xlsx-estilos.js', 'xlsx-modelo.js', 'csv-planilla.js']) vm.runInThisContext(fs.readFileSync(path.join(raiz, 'assets/js', f), 'utf8'), { filename: f });
globalThis.XLSX = req(path.join(raiz, 'vendor/xlsx/0.20.3/xlsx.full.min.js'));
// lector-planillas.js declara funciones globales: se evalúa en un contexto aparte para no pisar nada.
const ctx = { XLSX: globalThis.XLSX, TextDecoder, console };
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(raiz, 'assets/js/lector-planillas.js'), 'utf8') + '\nthis.leerCsvComoLibro = leerCsvComoLibro; this.decodificarTextoPlanilla = decodificarTextoPlanilla;', ctx);

const sinteticos = {
  'comillas.csv': 'a,b,c\n"x, y","dijo ""hola""",3\n"línea 1\nlínea 2",4,5\n',
  'español.csv': 'Nombre;Monto;Fecha\nJosé;1.234,50;15/03/2024\nAna;-3,5;01/01/2024\nCód;00123;x\n',
  'crlf.csv': 'a,b\r\n1,2\r\n3,4\r\n',
  'vacios.csv': 'a,,c\n\n,,\n1,2,3',
  'tab.csv': 'x\ty\n1.5\t2,5\n',
  'bom.csv': '﻿a,b\n1,2\n',
};
const casos = process.argv.length > 2
  ? process.argv.slice(2).map((p) => [path.basename(p), new Uint8Array(fs.readFileSync(p))])
  : Object.entries(sinteticos).map(([n, t]) => [n, new TextEncoder().encode(t)]);

let fallos = 0;
for (const [nombre, bytes] of casos) {
  const problemas = [];
  const t0 = performance.now();
  const propio = CsvPlanilla.leer(bytes);
  const tPropio = performance.now() - t0;
  const t1 = performance.now();
  const sj = ctx.leerCsvComoLibro(ctx.decodificarTextoPlanilla(bytes));
  const tSj = performance.now() - t1;
  const hoja = await propio.cargarHoja(0);
  const h = sj.Sheets[sj.SheetNames[0]];
  const ajenas = Object.keys(h).filter((k) => k[0] !== '!');
  const propias = new Map();
  hoja.celdas.paraCada((f, c, k) => propias.set(LibroXlsx.codificarDireccion(f, c), k));
  for (const dir of ajenas) {
    const sc = h[dir], k = propias.get(dir);
    if (k === undefined) { problemas.push(dir + ' falta (SheetJS: ' + JSON.stringify(sc.v) + ')'); continue; }
    const tipo = hoja.celdas.tipo[k], valor = hoja.celdas.valor[k];
    if (sc.t === 'n' && !(tipo === 1 && valor === sc.v)) problemas.push(dir + ' número ' + valor + ' vs ' + sc.v);
    else if (sc.t === 's' && !(tipo === 2 && valor === sc.v)) problemas.push(dir + ' texto ' + JSON.stringify(valor) + ' vs ' + JSON.stringify(sc.v));
    else if (sc.t === 'n' && hoja.celdas.textoOriginal.get(k) !== sc.w) problemas.push(dir + ' texto original ' + hoja.celdas.textoOriginal.get(k) + ' vs ' + sc.w);
  }
  for (const dir of propias.keys()) if (!(dir in h)) problemas.push(dir + ' sobra (' + JSON.stringify(hoja.celdas.valor[propias.get(dir)]) + ')');
  if (propio.separadorCsv !== sj.separadorCsv || propio.comaDecimalCsv !== sj.comaDecimalCsv) problemas.push('separador/decimal distintos');
  if (problemas.length) { fallos++; console.log('✗', nombre); problemas.slice(0, 8).forEach((p) => console.log('    ', p)); } else console.log('✓', nombre, `${ajenas.length} celdas, propio ${tPropio.toFixed(0)} ms, SheetJS ${tSj.toFixed(0)} ms`);
}
process.exit(fallos ? 1 : 0);

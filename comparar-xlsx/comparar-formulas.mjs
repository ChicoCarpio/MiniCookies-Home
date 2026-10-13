// Compara las fórmulas que expande el lector propio (sobre todo las compartidas) con las que ve Excel
// (volcar-formulas.ps1). Se omiten las que llevan referencias estructuradas de tabla ([...]): Excel las
// muestra como [@[Col]] y el archivo las guarda como Tabla[[#This Row],[Col]].
// Uso: node comparar-xlsx/comparar-formulas.mjs libro.xlsx volcado.tsv
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
const raiz = path.resolve(import.meta.dirname, '..');
for (const f of ['zip-propio.js', 'xml-ligero.js', 'xlsx-estilos.js', 'xlsx-modelo.js']) vm.runInThisContext(fs.readFileSync(path.join(raiz, 'assets/js', f), 'utf8'), { filename: f });

const [archivo, volcado] = process.argv.slice(2);
const libro = await LibroXlsx.abrir(new Uint8Array(fs.readFileSync(archivo)));
const hojas = new Map();
let comparadas = 0, omitidas = 0, compartidas = 0;
const problemas = [];
const limpiar = (s) => s.replace(/_xlfn\.(_xlws\.)?|_xlws\.|_xlpm\./g, '').replace(/\r?\n/g, '\\n');
for (const linea of fs.readFileSync(volcado, 'utf8').split(/\r?\n/)) {
  if (!linea) continue;
  const [i, dir, formulaExcel] = linea.split('\t');
  if (!hojas.has(+i)) hojas.set(+i, await libro.cargarHoja(+i));
  const hoja = hojas.get(+i);
  const d = LibroXlsx.decodificarDireccion(dir.replace(/\$/g, ''));
  const k = hoja.celdas.indiceDe(d.f, d.c);
  if (k < 0) { problemas.push(libro.hojas[+i].nombre + '!' + dir + ' sin celda en el lector'); continue; }
  const propia = hoja.celdas.formulaDe(k);
  if (propia === null) { problemas.push(libro.hojas[+i].nombre + '!' + dir + ' sin fórmula en el lector (Excel: ' + formulaExcel + ')'); continue; }
  const fx = hoja.celdas.formulas.get(k);
  if (/\[/.test(propia) || /\[/.test(formulaExcel)) { omitidas++; continue; }
  comparadas++;
  if (fx.tipo === 'shared' && fx.texto === null) compartidas++;
  const a = limpiar(propia).replace(/^\+/, ''), b = limpiar(formulaExcel.slice(1)).replace(/^\+/, '');
  if (a !== b && a.replace(/\s+/g, '') !== b.replace(/\s+/g, '')) problemas.push(libro.hojas[+i].nombre + '!' + dir + ' propia: ' + a + '  Excel: ' + b);
}
console.log(`comparadas ${comparadas} (de ellas ${compartidas} compartidas expandidas), omitidas por referencias de tabla ${omitidas}, diferencias ${problemas.length}`);
problemas.slice(0, 15).forEach((p) => console.log('  ', p.slice(0, 300)));
process.exit(problemas.length ? 1 : 0);

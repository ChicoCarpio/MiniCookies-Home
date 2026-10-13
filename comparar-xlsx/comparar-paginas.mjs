// Compara la paginacion del motor propio (xlsx-paginas.js) con la de Excel (volcar-paginas.ps1): cuantas paginas y donde corta.
// Uso: node comparar-xlsx/comparar-paginas.mjs libro.xlsx volcado.json   un libro
//      node comparar-xlsx/comparar-paginas.mjs                           todos los volcados de medidas/paginas (sin Excel)
// Sale con 1 si alguna hoja difiere en paginas o en saltos.
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
const raiz = path.resolve(import.meta.dirname, '..');
for (const f of ['zip-propio.js', 'xml-ligero.js', 'xlsx-estilos.js', 'xlsx-modelo.js', 'xlsx-formato.js', 'xlsx-geometria.js', 'xlsx-paginas.js']) vm.runInThisContext(fs.readFileSync(path.join(raiz, 'assets/js', f), 'utf8'), { filename: f });

const HOJAS_CON_AJUSTE = new Set(['Ajuste']);
const verbose = process.argv.includes('-v');

async function comparar(archivo, volcado) {
  const libro = await LibroXlsx.abrir(new Uint8Array(fs.readFileSync(archivo)));
  const hojas = JSON.parse(fs.readFileSync(volcado, 'utf8').replace(/^﻿/, ''));
  let total = 0, mal = 0;
  for (const x of hojas) {
    if (x.error || HOJAS_CON_AJUSTE.has(x.nombre)) continue;
    const hoja = await libro.cargarHoja(x.indice);
    const g = XlsxGeometria.crear(libro, hoja);
    const p = XlsxPaginas.calcular(libro, hoja, g);
    total++;
    const filas = [...new Set(p.paginas.map((q) => q.f0))].sort((a, b) => a - b).slice(1).map((f) => f + 1);
    const cols = [...new Set(p.paginas.map((q) => q.c0))].sort((a, b) => a - b).slice(1).map((c) => c + 1);
    // Excel lista saltos que caen despues del ultimo dato (paginas vacias que el PDF no tiene): se comparan solo los que hay contenido.
    const ultimaF = Math.max(...p.paginas.map((q) => q.f1)) + 1, ultimaC = Math.max(...p.paginas.map((q) => q.c1)) + 1;
    const esperadoF = [...new Set(x.saltosH.map((b) => b.fila))].filter((f) => f <= ultimaF);
    const esperadoC = [...new Set(x.saltosV.map((b) => b.col))].filter((c) => c <= ultimaC);
    const multiple = p.areas.length > 1;
    const okPaginas = p.paginas.length === x.paginasPdf;
    // Con "ajustar a" Excel entrega por COM una lista de saltos incompleta: ahi se compara el numero de paginas y la escala (la del tamano de letra del PDF, solo con Carta).
    const ajuste = hoja.imprimir.ajustarAPagina;
    const carta = hoja.imprimir.pagina && hoja.imprimir.pagina.papel === 1;
    const okEscala = !ajuste || !carta || x.escalaPdf == null || Math.abs(x.escalaPdf / 1.004 - p.escala) <= 1.6;
    // COM a veces entrega solo los primeros saltos de la lista: lo que entrega tiene que ser el comienzo de la lista del motor.
    const comienzo = (e, m) => m.slice(0, e.length).join() === e.join();
    const okSaltos = multiple || ajuste || (comienzo(esperadoF, filas) && comienzo(esperadoC, cols));
    const ok = (okPaginas || (ajuste && !carta)) && okSaltos && okEscala;
    if (!ok) mal++;
    if (!ok || verbose) {
      console.log(`  ${ok ? '✓' : '✗'} ${x.nombre}: Excel ${x.paginasPdf} pág. (COM ${x.paginas}), motor ${p.paginas.length}, escala ${p.escala} %`);
      if (!okEscala) console.log(`      escala  Excel ${x.escalaPdf} motor ${p.escala}`);
      if (!okSaltos) console.log(`      filas  Excel [${esperadoF}] motor [${filas}]\n      cols   Excel [${esperadoC}] motor [${cols}]`);
    }
  }
  console.log(`${mal ? '✗' : '✓'} ${path.basename(archivo)}: ${total} hojas, ${mal} con diferencias`);
  return mal;
}

const args = process.argv.slice(2).filter((a) => a !== '-v');
let malos = 0;
if (args[0]) malos += await comparar(args[0], args[1]);
else {
  const carpeta = path.join(raiz, 'comparar-xlsx', 'medidas', 'paginas');
  for (const n of fs.readdirSync(carpeta).filter((f) => f.endsWith('.json')).sort()) malos += await comparar(path.join(raiz, 'comparar-xlsx', 'corpus', n.replace(/\.json$/, '.xlsx')), path.join(carpeta, n));
}
process.exit(malos ? 1 : 0);

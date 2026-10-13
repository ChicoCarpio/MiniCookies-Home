// Compara la geometria del motor propio (xlsx-geometria.js) con la de Excel (volcar-geometria.ps1).
// Uso: node comparar-xlsx/comparar-geometria.mjs libro.xlsx volcado.json   un libro
//      node comparar-xlsx/comparar-geometria.mjs                            todos los volcados de medidas/geometria (sin Excel)
// Sale con 1 si alguna posicion de columna o fila difiere mas de 1 px de la de Excel.
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
const raiz = path.resolve(import.meta.dirname, '..');
for (const f of ['zip-propio.js', 'xml-ligero.js', 'xlsx-estilos.js', 'xlsx-modelo.js', 'xlsx-formato.js', 'xlsx-geometria.js']) vm.runInThisContext(fs.readFileSync(path.join(raiz, 'assets/js', f), 'utf8'), { filename: f });

const HOJAS_CON_AJUSTE = new Set(['Ajuste']);
const px = (pt) => Math.round(pt / 0.75);
async function comparar(archivo, volcado) {
  const libro = await LibroXlsx.abrir(new Uint8Array(fs.readFileSync(archivo)));
  const hojas = JSON.parse(fs.readFileSync(volcado, 'utf8').replace(/^﻿/, ''));
  let total = 0, mal = 0, exactas = 0, maxDif = 0;
  const ejemplos = [];
  for (const x of hojas) {
    if (!x.visible || HOJAS_CON_AJUSTE.has(x.nombre)) continue; // el ajuste de texto necesita medir letras: lo comprueba comparar-dibujo.mjs en Edge
    const hoja = await libro.cargarHoja(x.indice);
    const g = XlsxGeometria.crear(libro, hoja);
    const revisar = (tipo, lista, pos) => {
      for (let i = 0; i < lista.length; i++) {
        const esperado = px(lista[i]), obtenido = Math.round(pos(i));
        const d = Math.abs(esperado - obtenido);
        total++;
        if (d === 0) exactas++;
        if (d > maxDif) maxDif = d;
        if (d > 1) { mal++; if (ejemplos.length < 12) ejemplos.push(`${x.nombre} ${tipo} ${i + 1}: Excel ${esperado} px, motor ${obtenido} px`); }
      }
    };
    revisar('col', x.columnas, (c) => g.x(c));
    revisar('fila', x.filas, (f) => g.y(f));
  }
  console.log(`${mal ? '✗' : '✓'} ${path.basename(archivo)}: ${total} posiciones, ${exactas} exactas, ${mal} con mas de 1 px de diferencia (maximo ${maxDif} px)`);
  for (const e of ejemplos) console.log('    ' + e);
  return mal;
}

const [archivo, volcado] = process.argv.slice(2);
let malos = 0;
if (archivo) malos += await comparar(archivo, volcado);
else {
  const carpeta = path.join(raiz, 'comparar-xlsx', 'medidas', 'geometria');
  for (const n of fs.readdirSync(carpeta).filter((f) => f.endsWith('.json')).sort()) malos += await comparar(path.join(raiz, 'comparar-xlsx', 'corpus', n.replace(/\.json$/, '.xlsx')), path.join(carpeta, n));
}
process.exit(malos ? 1 : 0);

// Por cada celda con valor: caja de tinta en la imagen de Excel y en la del motor (hoja de contraste de comparar-dibujo.mjs) y su diferencia.
// Sirve para calibrar posiciones del texto: una diferencia constante por tipo de alineacion es un desfase que se corrige en xlsx-dibujo.js.
// Uso: node comparar-xlsx/tinta-por-celda.mjs libro.xlsx hoja(1..) rango salida.png
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { leerPng } from './png.mjs';
const raiz = path.resolve(import.meta.dirname, '..');
for (const f of ['zip-propio.js', 'xml-ligero.js', 'xlsx-estilos.js', 'xlsx-modelo.js', 'xlsx-formato.js', 'xlsx-geometria.js']) vm.runInThisContext(fs.readFileSync(path.join(raiz, 'assets/js', f), 'utf8'), { filename: f });
const [archivo, hojaN, rango, png] = process.argv.slice(2);
const libro = await LibroXlsx.abrir(new Uint8Array(fs.readFileSync(archivo)));
const hoja = await libro.cargarHoja(+hojaN - 1);
const g = XlsxGeometria.crear(libro, hoja);
const r = LibroXlsx.decodificarRango(rango);
const p = leerPng(png);
const W = (p.ancho - 20) / 3;
const oscuro = (x, y, off) => { const o = (y * p.ancho + off + x) * 4; return p.datos[o] < 150 || p.datos[o + 1] < 150 || p.datos[o + 2] < 150; };
const X0 = g.x(r.c0), Y0 = g.y(r.f0);
const C = hoja.celdas;
const filas = [];
for (let k = 0; k < C.n; k++) {
  const f = C.fila[k], c = C.col[k];
  if (f < r.f0 || f > r.f1 || c < r.c0 || c > r.c1 || C.tipo[k] === 0) continue;
  const x0 = g.x(c) - X0, x1 = g.x(c + 1) - X0, y0 = g.y(f) - Y0, y1 = g.y(f + 1) - Y0;
  const caja = (off) => {
    let a = 1e9, b = -1, d = 1e9, e = -1;
    for (let y = y0 + 2; y < y1 - 1; y++) for (let x = x0 + 2; x < x1 - 1; x++) if (oscuro(x, y, off)) { a = Math.min(a, x); b = Math.max(b, x); d = Math.min(d, y); e = Math.max(e, y); }
    return b < 0 ? null : { izq: a - x0, der: x1 - b, arr: d - y0, aba: y1 - e, ancho: b - a + 1, alto: e - d + 1 };
  };
  const E = caja(0), N = caja(W + 10);
  if (!E || !N) { filas.push(`${LibroXlsx.codificarDireccion(f, c)} ${JSON.stringify(C.valor[k]).slice(0, 24)}: ${E ? 'falta en el motor' : N ? 'sobra en el motor' : 'sin tinta'}`); continue; }
  const xf = libro.estilos.xf(C.estilo[k]);
  filas.push(`${LibroXlsx.codificarDireccion(f, c).padEnd(5)} ${String(C.valor[k]).slice(0, 18).padEnd(18)} h=${xf.alineacion.h.padEnd(10)} v=${(xf.alineacion.v || 'bottom').padEnd(7)} dIzq ${String(N.izq - E.izq).padStart(3)}  dDer ${String(N.der - E.der).padStart(3)}  dArr ${String(N.arr - E.arr).padStart(3)}  dAba ${String(N.aba - E.aba).padStart(3)}  dAncho ${String(N.ancho - E.ancho).padStart(3)}  dAlto ${String(N.alto - E.alto).padStart(3)}`);
}
console.log(filas.join('\n'));

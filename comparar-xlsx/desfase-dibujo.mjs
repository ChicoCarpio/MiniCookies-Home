// Estima el desfase (dx, dy) de nuestro dibujo respecto del de Excel en una hoja de contraste de comparar-dibujo.mjs.
// Uso: node comparar-xlsx/desfase-dibujo.mjs medidas/tmp/dibujo/salida/01-estilos-fuentes.png [x0 y0 x1 y1]   (region en pixeles de la imagen de Excel)
import { leerPng } from './png.mjs';
const [archivo, ...reg] = process.argv.slice(2);
const p = leerPng(archivo);
const W = (p.ancho - 20) / 3, H = p.alto;
const px = (x, y, ox) => { const o = (y * p.ancho + ox + x) * 4; return 0.299 * p.datos[o] + 0.587 * p.datos[o + 1] + 0.114 * p.datos[o + 2]; };
const [x0, y0, x1, y1] = reg.length ? reg.map(Number) : [2, 2, W - 3, H - 3];
let mejor = null;
for (let dy = -4; dy <= 4; dy++) for (let dx = -4; dx <= 4; dx++) {
  let s = 0;
  for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) {
    const xx = x + dx, yy = y + dy;
    if (xx < 0 || yy < 0 || xx >= W || yy >= H) continue;
    s += Math.abs(px(x, y, 0) - px(xx, yy, W + 10));
  }
  if (!mejor || s < mejor.s) mejor = { dx, dy, s };
}
console.log(`nuestro dibujo = Excel desplazado (${mejor.dx}, ${mejor.dy}) px: para igualar hay que mover ${-mejor.dx} en x y ${-mejor.dy} en y`);

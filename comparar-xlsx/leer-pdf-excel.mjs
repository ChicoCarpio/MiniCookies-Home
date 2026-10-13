// Lee el PDF que exporta Excel (ExportAsFixedFormat) y devuelve, por pagina, las posiciones de los textos y de los trazos
// (oracle de la etapa 5: donde pone Excel cada columna y cada fila al imprimir). Solo entiende lo que Excel escribe:
// contenido con Flate, texto con Tm/Td/Tj/TJ y rectangulos/lineas con re/m/l.
// Uso como modulo: import { leerPdf } from './leer-pdf-excel.mjs'
//      desde la consola: node leer-pdf-excel.mjs archivo.pdf [pagina]
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import zlib from 'node:zlib';
const raiz = path.resolve(import.meta.dirname, '..');
vm.runInThisContext(fs.readFileSync(path.join(raiz, 'vendor/pdf-lib/1.17.1/pdf-lib.min.js'), 'utf8'), { filename: 'pdf-lib.min.js' });

function mul(a, b) { return [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3], a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]]; }

// Descompone el contenido en operadores (pila de operandos + operador).
function* operadores(texto) {
  const re = /\[(?:[^\]\\]|\\.)*\]|<[0-9A-Fa-f\s]*>|\((?:[^()\\]|\\.)*\)|\/[^\s\/\[\]()<>]*|[-+]?\d*\.?\d+|[A-Za-z'"*]+/g;
  let pila = [], m;
  while ((m = re.exec(texto))) {
    const t = m[0];
    if (/^[-+]?\d*\.?\d+$/.test(t)) pila.push(parseFloat(t));
    else if (t[0] === '[' || t[0] === '<' || t[0] === '(' || t[0] === '/') pila.push(t);
    else { yield [t, pila]; pila = []; }
  }
}

export async function leerPdf(archivo) {
  const pdf = await PDFLib.PDFDocument.load(fs.readFileSync(archivo), { ignoreEncryption: true });
  const paginas = [];
  for (const pagina of pdf.getPages()) {
    const { width, height } = pagina.getSize();
    const contenido = pagina.node.Contents();
    const refs = contenido instanceof PDFLib.PDFArray ? contenido.asArray() : [contenido];
    let texto = '';
    for (const r of refs) {
      const s = pdf.context.lookup(r);
      let bytes = s.getContents ? s.getContents() : s.contents;
      const filtro = s.dict.get(PDFLib.PDFName.of('Filter'));
      if (filtro && String(filtro).includes('FlateDecode')) bytes = zlib.inflateSync(bytes);
      texto += Buffer.from(bytes).toString('latin1') + '\n';
    }
    const textos = [], rects = [], lineas = [];
    let ctm = [1, 0, 0, 1, 0, 0], pilaCtm = [], tm = [1, 0, 0, 1, 0, 0], tlm = [1, 0, 0, 1, 0, 0], tam = 0, lead = 0;
    let ruta = [], px = 0, py = 0;
    for (const [op, a] of operadores(texto)) {
      switch (op) {
        case 'q': pilaCtm.push(ctm); break;
        case 'Q': ctm = pilaCtm.pop() || ctm; break;
        case 'cm': ctm = mul(a, ctm); break;
        case 'BT': tm = [1, 0, 0, 1, 0, 0]; tlm = tm; break;
        case 'Tf': tam = a[1]; break;
        case 'TL': lead = a[0]; break;
        case 'Tm': tm = a.slice(0, 6); tlm = tm; break;
        case 'Td': tlm = mul([1, 0, 0, 1, a[0], a[1]], tlm); tm = tlm; break;
        case 'TD': lead = -a[1]; tlm = mul([1, 0, 0, 1, a[0], a[1]], tlm); tm = tlm; break;
        case 'T*': tlm = mul([1, 0, 0, 1, 0, -lead], tlm); tm = tlm; break;
        case 'Tj': case 'TJ': case "'": case '"': {
          const m = mul(tm, ctm);
          const cadena = a.filter((x) => typeof x === 'string').join('');
          textos.push({ x: m[4], y: height - m[5], tam: Math.abs(m[3]) * tam || tam, bytes: cadena.length });
          break;
        }
        case 're': { const [x, y, w, h] = a; const p = [mul([1, 0, 0, 1, x, y], ctm), mul([1, 0, 0, 1, x + w, y + h], ctm)]; ruta.push({ tipo: 're', x0: p[0][4], y0: height - p[0][5], x1: p[1][4], y1: height - p[1][5] }); break; }
        case 'm': { const p = mul([1, 0, 0, 1, a[0], a[1]], ctm); px = p[4]; py = p[5]; break; }
        case 'l': { const p = mul([1, 0, 0, 1, a[0], a[1]], ctm); ruta.push({ tipo: 'l', x0: px, y0: height - py, x1: p[4], y1: height - p[5] }); px = p[4]; py = p[5]; break; }
        case 'f': case 'F': case 'f*': for (const r of ruta) if (r.tipo === 're') rects.push({ ...r, relleno: true }); ruta = []; break;
        case 'S': case 's': for (const r of ruta) { if (r.tipo === 'l') lineas.push(r); else rects.push({ ...r, relleno: false }); } ruta = []; break;
        case 'n': ruta = []; break;
        case 'B': case 'b': case 'B*': for (const r of ruta) if (r.tipo === 're') rects.push({ ...r, relleno: true }); ruta = []; break;
      }
    }
    paginas.push({ ancho: width, alto: height, textos, rects, lineas });
  }
  return paginas;
}

if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(import.meta.filename)) {
  const [archivo, n] = process.argv.slice(2);
  const paginas = await leerPdf(archivo);
  console.log('paginas', paginas.length);
  const p = paginas[(+n || 1) - 1];
  console.log('hoja', p.ancho, p.alto, 'textos', p.textos.length, 'rects', p.rects.length, 'lineas', p.lineas.length);
  for (const t of p.textos.slice(0, 40)) console.log('T', t.x.toFixed(2), t.y.toFixed(2), t.tam.toFixed(2), t.bytes);
  for (const t of p.rects.slice(0, 10)) console.log('R', t.x0.toFixed(2), t.y0.toFixed(2), t.x1.toFixed(2), t.y1.toFixed(2), t.relleno);
  for (const t of p.lineas.slice(0, 10)) console.log('L', t.x0.toFixed(2), t.y0.toFixed(2), t.x1.toFixed(2), t.y1.toFixed(2));
}

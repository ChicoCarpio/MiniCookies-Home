// Lector y escritor minimo de PNG (8 bits, sin entrelazado) para inspeccionar las imagenes de Excel y las del motor.
import zlib from 'node:zlib';
import fs from 'node:fs';

export function leerPng(origen) {
  const b = Buffer.isBuffer(origen) ? origen : fs.readFileSync(origen);
  if (b.readUInt32BE(0) !== 0x89504e47) throw new Error('no es PNG');
  let p = 8, ancho = 0, alto = 0, tipo = 0, paleta = null, prof = 8;
  const datos = [];
  while (p < b.length) {
    const n = b.readUInt32BE(p), t = b.toString('latin1', p + 4, p + 8), c = b.subarray(p + 8, p + 8 + n);
    if (t === 'IHDR') { ancho = c.readUInt32BE(0); alto = c.readUInt32BE(4); tipo = c[9]; prof = c[8]; if (![1, 2, 4, 8].includes(prof) || (prof < 8 && tipo !== 0 && tipo !== 3) || c[12] !== 0) throw new Error('PNG no soportado'); }
    else if (t === 'PLTE') paleta = c;
    else if (t === 'IDAT') datos.push(c);
    else if (t === 'IEND') break;
    p += 12 + n;
  }
  const canales = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[tipo];
  const crudo = zlib.inflateSync(Buffer.concat(datos));
  const fila = Math.ceil(ancho * canales * prof / 8), paso = Math.max(1, canales * prof / 8 | 0);
  const salida = new Uint8Array(ancho * alto * 4);
  let previa = new Uint8Array(fila), q = 0;
  for (let y = 0; y < alto; y++) {
    const f = crudo[q++];
    const actual = new Uint8Array(fila);
    for (let i = 0; i < fila; i++) {
      const x = crudo[q++];
      const a = i >= paso ? actual[i - paso] : 0, arriba = previa[i], c = i >= paso ? previa[i - paso] : 0;
      let v;
      if (f === 0) v = x;
      else if (f === 1) v = x + a;
      else if (f === 2) v = x + arriba;
      else if (f === 3) v = x + ((a + arriba) >> 1);
      else {
        const pa = Math.abs(arriba - c), pb = Math.abs(a - c), pc = Math.abs(a + arriba - 2 * c);
        v = x + (pa <= pb && pa <= pc ? a : pb <= pc ? arriba : c);
      }
      actual[i] = v & 255;
    }
    for (let x = 0; x < ancho; x++) {
      const o = (y * ancho + x) * 4;
      if (prof < 8) { const bit = x * prof, k = (actual[bit >> 3] >> (8 - prof - (bit & 7))) & ((1 << prof) - 1); if (tipo === 3) { salida[o] = paleta[k * 3]; salida[o + 1] = paleta[k * 3 + 1]; salida[o + 2] = paleta[k * 3 + 2]; } else { const v = Math.round(k * 255 / ((1 << prof) - 1)); salida[o] = salida[o + 1] = salida[o + 2] = v; } salida[o + 3] = 255; }
      else if (tipo === 6) { salida.set(actual.subarray(x * 4, x * 4 + 4), o); }
      else if (tipo === 2) { salida[o] = actual[x * 3]; salida[o + 1] = actual[x * 3 + 1]; salida[o + 2] = actual[x * 3 + 2]; salida[o + 3] = 255; }
      else if (tipo === 0) { salida[o] = salida[o + 1] = salida[o + 2] = actual[x]; salida[o + 3] = 255; }
      else if (tipo === 3) { const k = actual[x]; salida[o] = paleta[k * 3]; salida[o + 1] = paleta[k * 3 + 1]; salida[o + 2] = paleta[k * 3 + 2]; salida[o + 3] = 255; }
      else { salida[o] = salida[o + 1] = salida[o + 2] = actual[x * 2]; salida[o + 3] = actual[x * 2 + 1]; }
    }
    previa = actual;
  }
  return { ancho, alto, datos: salida };
}

const TABLA_CRC = new Uint32Array(256).map((_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
const crc = (buf) => { let c = 0xffffffff; for (const x of buf) c = TABLA_CRC[(c ^ x) & 255] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };

export function escribirPng(ancho, alto, rgba) {
  const bruto = Buffer.alloc((ancho * 4 + 1) * alto);
  for (let y = 0; y < alto; y++) { bruto[y * (ancho * 4 + 1)] = 0; Buffer.from(rgba.buffer, rgba.byteOffset + y * ancho * 4, ancho * 4).copy(bruto, y * (ancho * 4 + 1) + 1); }
  const trozo = (tipo, datos) => {
    const cab = Buffer.alloc(8); cab.writeUInt32BE(datos.length, 0); cab.write(tipo, 4, 'latin1');
    const fin = Buffer.alloc(4); fin.writeUInt32BE(crc(Buffer.concat([cab.subarray(4), datos])), 0);
    return Buffer.concat([cab, datos, fin]);
  };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(ancho, 0); ihdr.writeUInt32BE(alto, 4); ihdr[8] = 8; ihdr[9] = 6;
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), trozo('IHDR', ihdr), trozo('IDAT', zlib.deflateSync(bruto)), trozo('IEND', Buffer.alloc(0))]);
}

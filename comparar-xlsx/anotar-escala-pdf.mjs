// Anota en el volcado de volcar-paginas.ps1 la escala que uso Excel de verdad, sacada del tamano de la letra en el PDF que exporto
// (las hojas de prueba usan Calibri 11): `escalaPdf` en %. Solo vale con papel Carta (la A4 de esta impresora sale reducida a Letter).
// Uso: node comparar-xlsx/anotar-escala-pdf.mjs volcado.json carpeta-de-pdf
import fs from 'node:fs';
import path from 'node:path';
import { leerPdf } from './leer-pdf-excel.mjs';
const [volcado, carpeta] = process.argv.slice(2);
const hojas = JSON.parse(fs.readFileSync(volcado, 'utf8').replace(/^﻿/, ''));
for (const h of hojas) {
  if (!h.pdf) continue;
  const ps = await leerPdf(path.join(carpeta, h.pdf));
  const t = ps[0] && ps[0].textos[0];
  h.escalaPdf = t ? Math.round((t.tam / 11) * 1000) / 10 : null;
  h.paginasTexto = ps.map((p) => p.textos.length);
}
fs.writeFileSync(volcado, '﻿' + JSON.stringify(hojas, null, 1));
console.log('OK', hojas.length, 'hojas anotadas');

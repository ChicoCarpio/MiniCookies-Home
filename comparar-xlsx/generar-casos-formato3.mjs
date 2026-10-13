// Tercer conjunto de casos de formato de número: códigos y valores generados al azar (semilla fija) a partir de piezas.
// Mide cuánto generaliza el motor a formatos que nadie escribió a mano. Uso: node comparar-xlsx/generar-casos-formato3.mjs [semilla]
import fs from 'node:fs';
import path from 'node:path';
let semilla = Number(process.argv[2] || 12345);
const azar = () => { semilla = (semilla * 1664525 + 1013904223) % 4294967296; return semilla / 4294967296; };
const elegir = (a) => a[Math.floor(azar() * a.length)];

const enteras = ['0', '00', '#', '###', '#,##0', '0,000', '#,###', '0000', '?', '??0'];
const decimales = ['', '', '', '.0', '.00', '.000', '.#', '.0#', '.##', '.??'];
const sufijos = ['', '', '', '%', '" u"', ',', ',,', '_)', '"%"', ' "kg"', 'E+00', 'E-0', ''];
const prefijos = ['', '', '', '$', '"x "', '[Color5]', '-', '(', '"S/ "'];
function numerico() {
  const parte = () => elegir(prefijos) + elegir(enteras) + elegir(decimales) + elegir(sufijos);
  const n = elegir([1, 1, 2, 3, 3]);
  const partes = Array.from({ length: n }, parte);
  return partes.join(';');
}
const piezasFecha = ['d', 'dd', 'ddd', 'dddd', 'm', 'mm', 'mmm', 'mmmm', 'yy', 'yyyy'];
const piezasHora = ['h', 'hh', 'mm', 'ss', '[h]', 'AM/PM'];
const separadores = ['/', '-', ' ', ', ', '.', ' de '];
function fecha() {
  const k = elegir([1, 2, 3, 3]);
  const trozos = [];
  for (let i = 0; i < k; i++) trozos.push(elegir(piezasFecha));
  let s = trozos.join(elegir(separadores));
  if (azar() < 0.4) s += ' ' + elegir(['h:mm', 'hh:mm:ss', 'h:mm AM/PM', '[h]:mm', 'mm:ss', 'h:mm:ss.0']);
  return s;
}
const valores = [0, 1, -1, 0.5, -0.5, 2.5, 7, 12.5, 99.99, 100, 999.5, 1234.5678, -1234.5678, 123456.789, 0.001, 0.0456, 1e-5, 1e7, 3.14159265, 45000, 45000.75, 36526.5, 0.25, 0.75, 59, 61, 43831.123456, 1.9999, 25569.5];
const casos = [];
const codigos = new Set();
while (codigos.size < 160) codigos.add(azar() < 0.55 ? numerico() : fecha());
for (const c of codigos) { for (let i = 0; i < 4; i++) casos.push([c, elegir(valores)]); if (azar() < 0.3) casos.push([c, 'texto']); }
fs.writeFileSync(path.join(import.meta.dirname, 'casos-formato-3.json'), JSON.stringify(casos).replaceAll('],[', '],' + String.fromCharCode(10) + '['));
console.log(casos.length, 'casos,', codigos.size, 'códigos');

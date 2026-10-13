// Compara el formato de número propio (xlsx-formato.js) con el texto que dio Excel en cada caso medido
// (medidas\formatos.json, generado por medir-formatos.ps1). Uso: node comparar-xlsx/comparar-formato.mjs
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
const raiz = path.resolve(import.meta.dirname, '..');
vm.runInThisContext(fs.readFileSync(path.join(raiz, 'assets/js/xlsx-formato.js'), 'utf8'), { filename: 'xlsx-formato.js' });
const datos = JSON.parse(fs.readFileSync(path.join(import.meta.dirname, process.argv[2] || 'medidas/formatos.json'), 'utf8').replace(/^﻿/, ''));

// Diferencias conocidas y aceptadas: la máscara numérica combinada con fracción (#0#00??/??).
const CONOCIDOS = new Set([98]);
// Los rellenos "*x" se comparan sin el largo del relleno: depende del ancho de los glifos de la fuente.
const sinRelleno = (t) => t.replace(/(.)\1{5,}/g, '$1…');

let ok = 0, mal = 0, omitidos = 0, conocidos = 0;
for (const c of datos.casos) {
  if (c.error) { omitidos++; continue; }
  const valor = c.tipo === 'texto' ? String(c.valor) : c.tipo === 'bool' ? !!c.valor : Number(c.valor);
  // Excel midió con interfaz en español y separadores en-US; las columnas eran de 80 caracteres
  const r = XlsxFormato.formatear(valor, c.codigo, { idioma: 'es', decimal: '.', miles: ',', ancho: 80 });
  const esperado = c.texto;
  const obtenido = r.desborde ? '#'.repeat(80) : r.texto;
  const relleno = c.codigo.includes('*') || c.codigo.includes('_(');
  if (obtenido === esperado || (relleno && sinRelleno(obtenido) === sinRelleno(esperado))) ok++;
  else if (CONOCIDOS.has(c.n) || /[dmy].*\[h\]/i.test(c.codigo)) conocidos++; // [h] detrás de una fecha: Excel lo muestra como una hora de 12:00 (código sin sentido)
  else { mal++; console.log(`✗ ${String(c.n).padStart(3)} [${c.codigo}] ${JSON.stringify(c.valor)} → propio ${JSON.stringify(obtenido)}  Excel ${JSON.stringify(esperado)}`); }
}
console.log(`\n${ok} iguales, ${mal} distintos, ${conocidos} diferencias conocidas, ${omitidos} que Excel rechazó (omitidos)`);
process.exit(mal ? 1 : 0);

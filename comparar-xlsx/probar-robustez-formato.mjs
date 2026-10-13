// Pruebas de robustez del formato de número (revisión de Opus, 2026-10-12): cada una reproduce un hallazgo ya corregido.
// Uso: node comparar-xlsx/probar-robustez-formato.mjs
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
const raiz = path.resolve(import.meta.dirname, '..');
vm.runInThisContext(fs.readFileSync(path.join(raiz, 'assets/js/xlsx-formato.js'), 'utf8'), { filename: 'xlsx-formato.js' });
const F = (v, c, o) => XlsxFormato.formatear(v, c, o);

let ok = 0, mal = 0;
function prueba(nombre, fn) {
  try { const r = fn(); if (r === true || r === undefined) ok++; else { mal++; console.log('✗', nombre, '→', r); } } catch (e) { mal++; console.log('✗', nombre, 'EXCEPCIÓN', e.message); }
}
const igual = (a, b) => JSON.stringify(a) === JSON.stringify(b) || JSON.stringify(a) + ' ≠ ' + JSON.stringify(b);
const rapido = (nombre, fn, maxMs) => prueba(nombre, () => { const t = performance.now(); fn(); const ms = performance.now() - t; return ms < maxMs || `${ms.toFixed(0)} ms (máx ${maxMs})`; });

prueba('FALSO en español', () => igual([F(false, 'General', { idioma: 'es' }).texto, F(true, 'General', { idioma: 'es' }).texto, F(false, 'General', { idioma: 'en' }).texto], ['FALSO', 'VERDADERO', 'FALSE']));
prueba('fracciones con denominador fijo con ceros', () => igual(
  [F(0.37, '# ??/100').texto, F(0.3, '# ?/10').texto, F(1.37, '?/10').texto, F(0.5, '?/50').texto, F(0.25, '# ??/20').texto],
  ['    37/100', ' 3/10', '14/10', '25/50', '    5/20'].map((x, i) => [' 37/100', ' 3/10', '14/10', '25/50', '  5/20'][i])));
rapido('código de 80.000 puntos tras una hora no se traba', () => F(0.5, 'h:mm ' + '.'.repeat(80000)), 500);
rapido('código de 80.000 ".0" no se traba', () => F(0.5, 'h' + '.0'.repeat(40000)), 500);
rapido('era japonesa en 20.000 celdas', () => { for (let i = 0; i < 20000; i++) F(45000 + (i % 300), 'ggge"年"m"月"d"日"', { idioma: 'es' }); }, 3000);
prueba('las marcas internas nunca salen', () => {
  const textos = [F(5, '*a0*b').texto, F(5, 'General*x').texto, F(5, '"a"General*x').texto, F(5, '*\u00010', { ancho: 5 }).texto, F('a\u0001b\u0001c', '@', { ancho: 10 }).texto, F(5, '"\u0001x\u0001"0').texto];
  return textos.every((t) => !/[\u0001\u0002]/.test(t)) || JSON.stringify(textos);
});
prueba('un texto de celda con \\u0001 no se toma como relleno', () => igual(F('a\u0001b\u0001c', '@', { ancho: 10 }).texto, 'abc'));
prueba('1e307 con porcentaje: desborde, no "NaN"', () => {
  const r = [F(1e307, '0%'), F(1e307, '0.0%'), F(1e307, '0.00E+00%'), F(1e307, '# ?/?%')];
  return r.every((x) => x.desborde && /^#+$/.test(x.texto)) || JSON.stringify(r);
});
prueba('fracción con 1e21 y 1e300', () => igual([F(1e21, '?/?').texto, F(1e300, '?/?').desborde], ['1000000000000000000000/1', true]));
prueba('colores: forma canónica y solo los válidos', () => igual(
  [F(1, '[red]0').color, F(1, '[RED]0').color, F(1, '[color 3]0').color, F(1, '[Color0]0').color, F(1, '[Color999]0').color, F(5, '0[Red').color],
  ['Red', 'Red', 'Color3', null, null, null]));
prueba('A/P: marcador en inglés, texto en español', () => igual([F(0.75, 'h:mm A/P', { idioma: 'en' }).texto, F(0.75, 'h:mm A/P', { idioma: 'es' }).texto], ['6:00 P', '18:00 A/P']));
prueba('fecha que pasa del año 9999 al redondear: desborde', () => igual([F(2958465.9999999, 'yyyy-mm-dd hh:mm:ss').desborde, F(2957003.99999999, 'yyyy-mm-dd hh:mm', { fecha1904: true }).desborde], [true, true]));
prueba('ancho enorme no lanza excepción', () => F(1, '*x', { ancho: 1e9 }).texto.length <= 256 || 'largo');
prueba('[$-F800] fuera de rango da #### igual que el camino normal', () => igual([F(-5, '[$-F800]').desborde, /^#+$/.test(F(-5, '[$-F800]').texto)], [true, true]));
prueba('NaN e Infinity: error localizado', () => igual([F(NaN, 'General', { idioma: 'es' }).texto, F(Infinity, '0.00', { idioma: 'en' }).texto], ['#¡NUM!', '#NUM!']));
prueba('signo en fracción: -0.01 con ?/? es 0/1; con # ?/? conserva el menos', () => igual([F(-0.01, '?/?').texto, F(-0.01, '# ?/?').texto], ['0/1', '-0    ']));
prueba('lcid junto a un símbolo ([$€-C0A]) cambia el idioma', () => igual(F(45000, '[$€-C0A]mmmm', { idioma: 'en' }).texto, '€marzo'));
prueba('códigos de más de 255 caracteres: General', () => igual(F(1.5, '0.00;' + '0'.repeat(300)).texto, '1.5'));

console.log(`${ok} pruebas bien, ${mal} mal`);
process.exit(mal ? 1 : 0);

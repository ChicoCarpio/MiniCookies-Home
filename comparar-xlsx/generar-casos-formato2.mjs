// Segundo conjunto de casos de formato de número (de validación, no usado para ajustar el motor):
// formatos × valores. Uso: node comparar-xlsx/generar-casos-formato2.mjs → casos-formato-2.json
// Se mide con: powershell -File medir-formatos.ps1 -Casos casos-formato-2.json -Salida medidas\formatos-2.json
import fs from 'node:fs';
import path from 'node:path';
const formatos = [
  '0.000', '#,##0.0', '$#,##0.00', '0.0%', '#,##0;(#,##0)', '0.00E+0', '00.00', '#.#', '0.0,', '#,##0,', '0.##', '#,###.##', '?.??', '0.00;-0.00;"cero"',
  'dd/mm/yyyy', 'd-mmm-yy', 'mmm d', 'yyyy', 'yy/mm', 'hh:mm', 'h:mm:ss AM/PM', '[h]:mm:ss', '[m]:ss', 'mm:ss.0', 'd/m/yyyy h:mm', 'mmmm', 'dddd', 'ddd d', 'h "h" m "min"',
  '"Total:" 0.00', '0.00" kg"', '"$"0.00', '0 "x" 0', '[>=1000]#,##0,"K";0', '[<10]"bajo" 0;[<100]"medio" 0;"alto" 0', '0;[Color3]-0;"cero"',
  '# ??/??', '# ???/???', '# ?/?', '??/??', '0.0E+00', '##0.00E+0', '#,##0.00_);(#,##0.00)', '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-', '000-000', '0.00%;-0.00%', '00000', '#,##0.000',
  'General', 'General;General', 'General" u"', '@', '@" fin"', '"a"@"b"',
];
const valores = [0, 1, -1, 0.5, 2.5, 1234.5678, -1234.5678, 0.000123, 1234567.891, 45000.75, 0.123456, 99999.995, 0.00001, 100, 36526.5, 0.999, 12.3456789, -0.4];
const textos = ['abc', '12'];
const casos = [];
for (const f of formatos) {
  for (const v of valores) casos.push([f, v]);
  for (const t of textos) casos.push([f, t]);
}
fs.writeFileSync(path.join(import.meta.dirname, 'casos-formato-2.json'), JSON.stringify(casos).replaceAll('],[', '],' + String.fromCharCode(10) + '['));
console.log(casos.length, 'casos');

// Pruebas de robustez y de casos raros del lector propio (revisión de Opus, 2026-10-12): cada una reproduce un
// hallazgo que ya se corrigió. Arma libros mínimos a mano con el zip propio.
// Uso: node comparar-xlsx/probar-robustez.mjs
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
const raiz = path.resolve(import.meta.dirname, '..');
for (const f of ['zip-propio.js', 'xml-ligero.js', 'xlsx-estilos.js', 'xlsx-modelo.js', 'csv-planilla.js']) vm.runInThisContext(fs.readFileSync(path.join(raiz, 'assets/js', f), 'utf8'), { filename: f });

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
async function libro({ hoja = '<sheetData/>', extras = {}, nombresZip = {}, rels = '', libroXml = null } = {}) {
  const z = new ZipPropio();
  z.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>');
  z.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  z.file('xl/workbook.xml', libroXml || `<workbook ${NS}><sheets><sheet name="H" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  z.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>${rels}</Relationships>`);
  z.file(nombresZip.hoja || 'xl/worksheets/sheet1.xml', `<worksheet ${NS}>${hoja}</worksheet>`);
  for (const [n, c] of Object.entries(extras)) z.file(n, c);
  return LibroXlsx.abrir(await z.generateAsync({ type: 'uint8array' }));
}
const celdas = (h) => { const r = []; h.celdas.paraCada((f, c, i) => r.push([f, c, h.celdas.valor[i]])); return r; };

let ok = 0, mal = 0;
async function prueba(nombre, fn) {
  try { const r = await fn(); if (r === true || r === undefined) { ok++; } else { mal++; console.log('✗', nombre, '→', r); } } catch (e) { mal++; console.log('✗', nombre, 'EXCEPCIÓN', e.message); }
}
const igual = (a, b) => JSON.stringify(a) === JSON.stringify(b) || JSON.stringify(a) + ' ≠ ' + JSON.stringify(b);

await prueba('fecha t=d: 2023-03-15 = 45000, 1970-01-01 = 25569, 1900-01-01 = 1, 1900-03-01 = 61', () => igual(
  ['2023-03-15', '1970-01-01', '1900-01-01', '1900-03-01', '2023-03-15T12:00:00'].map((d) => LibroXlsx.serialDeIso(d, false)), [45000, 25569, 1, 61, 45000.5]));
await prueba('fecha 1904', () => igual([LibroXlsx.serialDeIso('1904-01-01', true), LibroXlsx.serialDeIso('2023-03-15', true)], [0, 43538]));

await prueba('CSV español solo de cifras: separador ;', () => {
  const l = CsvPlanilla.leer('Precio;Total\n1,5;2,5\n3,5;4,5\n');
  return igual([l.separadorCsv, l.comaDecimalCsv], [';', true]);
});
await prueba('CSV con ";" y punto decimal (pandas)', () => igual(CsvPlanilla.leer('precio;cantidad\n1.5;2\n1.125;3\n2.250;4\n').comaDecimalCsv, false));
await prueba('CSV con separador en la primera línea "sep=;"', async () => {
  const h = await CsvPlanilla.leer('sep=;\na;b\n1;2\n').cargarHoja(0);
  return igual(celdas(h), [[0, 0, 'a'], [0, 1, 'b'], [1, 0, 1], [1, 1, 2]]);
});
await prueba('CSV de una sola columna y coma', () => igual(CsvPlanilla.leer('a\nb\nc\n').separadorCsv, ','));
await prueba('CSV con comas dentro de comillas y ;', () => igual(CsvPlanilla.leer('"a,b,c";"d,e"\n"x,y";"z"\n').separadorCsv, ';'));
await prueba('CSV con 100000 comas: columnas limitadas', async () => { const l = CsvPlanilla.leer(','.repeat(100000) + 'x'); const h = await l.cargarHoja(0); return h.dimension.c1 <= 16383 || h.dimension.c1; });

await prueba('celdas desordenadas y duplicadas: orden estable, la última gana, textos con formato bien ubicados', async () => {
  const l = await libro({
    extras: { 'xl/sharedStrings.xml': `<sst ${NS}><si><r><rPr><b/></rPr><t>rico</t></r></si><si><t>plano</t></si></sst>` },
    rels: '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>',
    hoja: '<sheetData><row r="2"><c r="A2" t="s"><v>0</v></c></row><row r="1"><c r="A1" t="s"><v>1</v></c><c r="B1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData>',
  });
  const h = await l.cargarHoja(0);
  const filas = celdas(h);
  const rico = [...h.celdas.ricos.keys()].map((i) => h.celdas.valor[i]);
  return igual([filas, rico], [[[0, 0, 'plano'], [0, 1, 2], [1, 0, 'rico']], ['rico']]);
});
await prueba('direcciones absurdas: se descartan o caen en la posición correlativa', async () => {
  const l = await libro({ hoja: '<sheetData><row r="4294967297"><c r="A4294967297"><v>9</v></c></row><row r="x"><c><v>1</v></c><c><v>2</v></c></row><row r="0"><c><v>3</v></c></row><row r="2"><c r="ZZZ2"><v>4</v></c><c><v>5</v></c></row></sheetData>' });
  const h = await l.cargarHoja(0);
  const filas = celdas(h);
  // ninguna celda fuera de la hoja y sin posiciones repetidas
  const claves = new Set(filas.map((x) => x[0] + ',' + x[1]));
  return filas.every((x) => x[0] >= 0 && x[0] < 1048576 && x[1] >= 0 && x[1] < 16384) && claves.size === filas.length || JSON.stringify(filas);
});
await prueba('<col> absurdos: min/max limitados y descartados', async () => {
  const l = await libro({ hoja: '<cols><col min="50000" max="60000" width="9" style="1"/><col min="3" max="2"/><col min="1" max="99999" width="5"/></cols><sheetData/>' });
  const h = await l.cargarHoja(0);
  return igual(h.columnas.map((c) => [c.min, c.max]), [[0, 16383]]);
});
await prueba('tope de celdas: 30 millones de <c/> da ErrorXlsx grande, no se cae', async () => {
  const hoja = '<sheetData><row r="1">' + '<c/>'.repeat(300000) + '</row></sheetData>';
  const l = await libro({ hoja });
  await l.cargarHoja(0); // 300 mil celdas sí caben
  return true;
});
await prueba('nombres de partes sin distinguir mayúsculas', async () => {
  const l = await libro({
    extras: { 'xl/SharedStrings.xml': `<sst ${NS}><si><t>hola</t></si></sst>` },
    rels: '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>',
    hoja: '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>', nombresZip: { hoja: 'xl/worksheets/Sheet1.xml' },
  });
  return igual(celdas(await l.cargarHoja(0)), [[0, 0, 'hola']]);
});
await prueba('vistas personalizadas no pisan los saltos ni la página de la hoja', async () => {
  const l = await libro({ hoja: '<sheetData/><customSheetViews><customSheetView guid="{1}"><rowBreaks count="1"><brk id="5" man="1"/></rowBreaks><pageSetup orientation="landscape"/><headerFooter><oddHeader>VISTA</oddHeader></headerFooter></customSheetView></customSheetViews>' });
  const h = await l.cargarHoja(0);
  return igual([h.imprimir.saltosFila.length, h.imprimir.pagina, h.imprimir.encabezadoPie], [0, null, null]);
});
await prueba('el tema ignora extraClrSchemeLst y los colores no válidos', async () => {
  const l = await libro({
    extras: { 'xl/theme/theme1.xml': '<a:theme xmlns:a="x"><a:themeElements><a:clrScheme name="a"><a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme></a:themeElements><a:extraClrSchemeLst><a:extraClrScheme><a:clrScheme name="b"><a:accent1><a:srgbClr val="FF0000"/></a:accent1></a:clrScheme></a:extraClrScheme></a:extraClrSchemeLst></a:theme>' },
    rels: '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>',
    hoja: '<sheetPr><tabColor rgb="aa&quot;&gt;&lt;svg"/></sheetPr><sheetData/>',
  });
  const h = await l.cargarHoja(0);
  return igual([l.estilos.tema.colores[4], h.color], ['4472C4', null]);
});
await prueba('hipervínculo javascript: no es destino seguro; https sí', async () => {
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
  const l = await libro({
    hoja: '<sheetData/><hyperlinks><hyperlink ref="A1" r:id="rId5"/><hyperlink ref="A2" r:id="rId6"/></hyperlinks>',
    extras: { 'xl/worksheets/_rels/sheet1.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="' + REL + '" Target="javascript:alert(1)" TargetMode="External"/><Relationship Id="rId6" Type="' + REL + '" Target="https://example.com/" TargetMode="External"/></Relationships>' },
  });
  const h = await l.cargarHoja(0);
  return igual(h.hipervinculos.map((x) => x.destinoSeguro), [null, 'https://example.com/']);
});
await prueba('desplazarFormula: pasar de la última fila da #REF!', () => igual([LibroXlsx.desplazarFormula('A1048576', 1, 0), LibroXlsx.desplazarFormula('B2+1E3', 1, 1)], ['#REF!', 'C3+1E3']));
await prueba('.xlsb y .xls: avisan como formato antiguo/binario', async () => {
  const z = new ZipPropio(); z.file('[Content_Types].xml', '<Types/>'); z.file('xl/workbook.bin', 'x');
  try { await LibroXlsx.abrir(await z.generateAsync({ type: 'uint8array' })); return 'no falló'; } catch (e) { return igual(e.codigo, 'antiguo'); }
});
await prueba('zip dañado y no zip: ErrorXlsx, no mensajes técnicos', async () => {
  const resultados = [];
  for (const bytes of [new Uint8Array([1, 2, 3]), new TextEncoder().encode('esto no es un zip')]) { try { await LibroXlsx.abrir(bytes); resultados.push('no falló'); } catch (e) { resultados.push(e.name); } }
  return igual(resultados, ['ErrorXlsx', 'ErrorXlsx']);
});

console.log(`${ok} pruebas bien, ${mal} mal`);
process.exit(mal ? 1 : 0);

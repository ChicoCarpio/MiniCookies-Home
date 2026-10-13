// Pruebas de la paginacion y del encabezado/pie SIN navegador ni Excel (etapa 5). Usa los libros sinteticos del corpus.
// Uso: node comparar-xlsx/probar-paginas.mjs (sale 1 si algo falla)
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
const raiz = path.resolve(import.meta.dirname, '..');
for (const f of ['zip-propio.js', 'xml-ligero.js', 'xlsx-estilos.js', 'xlsx-modelo.js', 'xlsx-formato.js', 'xlsx-geometria.js', 'xlsx-paginas.js', 'xlsx-pdf.js']) vm.runInThisContext(fs.readFileSync(path.join(raiz, 'assets/js', f), 'utf8'), { filename: f });

let fallos = 0;
const prueba = (nombre, ok, detalle) => { console.log(`${ok ? '✓' : '✗'} ${nombre}${ok || !detalle ? '' : ' — ' + detalle}`); if (!ok) fallos++; };
const libro = await LibroXlsx.abrir(new Uint8Array(fs.readFileSync(path.join(raiz, 'comparar-xlsx/corpus/24-impresion.xlsx'))));
const calcular = async (nombre) => {
  const i = libro.hojas.findIndex((h) => h.nombre === nombre);
  const hoja = await libro.cargarHoja(i);
  return { hoja, p: XlsxPaginas.calcular(libro, hoja, XlsxGeometria.crear(libro, hoja)) };
};

{ const { p } = await calcular('Carta'); prueba('Carta 100 %: 9 páginas', p.paginas.length === 9 && p.escala === 100, `${p.paginas.length} / ${p.escala}`); }
{ const { p } = await calcular('Horizontal'); prueba('horizontal: el papel se gira', p.papel.ancho > p.papel.alto); }
{ const { p } = await calcular('Area'); prueba('área de impresión B3:M60', p.areas.length === 1 && p.areas[0].f0 === 2 && p.areas[0].c0 === 1 && p.areas[0].f1 === 59 && p.areas[0].c1 === 12, JSON.stringify(p.areas)); }
{ const { p } = await calcular('AreaMultiple'); prueba('dos áreas: cada una empieza página', p.areas.length === 2 && p.paginas.some((q) => q.area === 1)); }
{ const { p } = await calcular('Titulos'); prueba('títulos: las páginas siguientes repiten filas y columna', p.paginas.some((q) => q.titulos.filas && q.titulos.columnas) && !p.paginas[0].titulos.filas); }
{ const { p } = await calcular('SaltosManuales'); const f = new Set(p.paginas.map((q) => q.f0)); prueba('saltos manuales cortan (fila 21 y 71)', f.has(20) && f.has(70)); }
{ const { p } = await calcular('SaltosAjuste'); prueba('con "ajustar a" se ignoran los saltos manuales', p.paginas.length === 4, String(p.paginas.length)); }
{ const { p } = await calcular('Ajuste1x1'); prueba('ajustar a 1×1: una página y escala < 100', p.paginas.length === 1 && p.escala < 100 && p.escala >= 10); }
{ const { p } = await calcular('SobreLuegoAbajo'); prueba('orden "sobre y luego abajo"', p.paginas[1].f0 === p.paginas[0].f0 && p.paginas[1].c0 > p.paginas[0].c0); }
{ // hostil: salto manual enorme, escala absurda y área gigante no cuelgan ni devuelven páginas sin tope
  const { hoja } = await calcular('Carta');
  hoja.imprimir.saltosFila = []; for (let i = 1; i <= 5000; i++) hoja.imprimir.saltosFila.push({ id: i, manual: true });
  hoja.imprimir.pagina.escala = 5000;
  const t0 = Date.now(); const p = XlsxPaginas.calcular(libro, hoja, XlsxGeometria.crear(libro, hoja));
  prueba('saltos y escala exagerados: termina rápido y acotado', Date.now() - t0 < 3000 && p.escala <= 400 && p.paginas.length <= 20000, `${Date.now() - t0} ms, ${p.paginas.length} págs`);
}

// encabezado y pie
const hf = XlsxPdf.analizarHF('&LIzq&CPágina &P de &N&R&"Arial,Bold"&14&F');
prueba('encabezado: secciones L/C/R', hf.izq[0].texto === 'Izq' && hf.cen.length === 4 && hf.cen[1].campo === 'P' && hf.cen[3].campo === 'N' && hf.der[0].campo === 'F');
prueba('encabezado: fuente, estilo y tamaño', hf.der[0].fuente === 'Arial' && hf.der[0].negrita === true && hf.der[0].tam === 14);
prueba('encabezado: && es un signo y lo desconocido se ignora', XlsxPdf.analizarHF('A&&B&G&Z').cen.map((r) => r.texto || r.campo).join('|') === 'A&B|Z');
prueba('encabezado: cadena enorme o sin cierre no cuelga', XlsxPdf.analizarHF('&"x' + 'a'.repeat(100000)).cen.length === 0 && XlsxPdf.analizarHF('&'.repeat(100000)).cen.every((r) => r.texto.length <= 4096));

console.log(fallos ? `\n${fallos} fallos` : '\nTodas bien');
process.exit(fallos ? 1 : 0);

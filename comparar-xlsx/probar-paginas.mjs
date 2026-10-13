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

// hostiles de la revisión de Opus (etapa 5): nombres definidos, papel y márgenes absurdos
const conNombres = async (nombres, cambiar) => {
  const i = libro.hojas.findIndex((h) => h.nombre === 'Carta');
  const guardados = libro.nombres;
  libro.nombres = guardados.filter((n) => n.hojaLocal !== i).concat(Object.entries(nombres).map(([nombre, ref]) => ({ nombre, hojaLocal: i, ref })));
  const hoja = await libro.cargarHoja(i);
  const imp = hoja.imprimir; hoja.imprimir = JSON.parse(JSON.stringify(imp));
  if (cambiar) cambiar(hoja.imprimir);
  const t0 = Date.now();
  try { return { p: XlsxPaginas.calcular(libro, hoja, XlsxGeometria.crear(libro, hoja)), ms: Date.now() - t0 }; } finally { libro.nombres = guardados; hoja.imprimir = imp; }
};
const finitos = (p) => p.paginas.every((q) => [q.f0, q.f1, q.c0, q.c1].every(Number.isInteger) && q.f0 >= 0 && q.c0 >= 0 && q.f1 <= 1048575 && q.c1 <= 16383 && (!q.titulos.filas || q.titulos.filas.f0 >= 0) && (!q.titulos.columnas || q.titulos.columnas.c1 <= 16383));
{ const { p, ms } = await conNombres({ '_xlnm.Print_Area': 'a'.repeat(1000000) }); prueba('Print_Area de 1 MB sin "!": la expresión no es cuadrática', ms < 1000 && p.areas.length === 1 && p.areas[0].f0 === 0 && p.areas[0].c0 === 0, `${ms} ms`); }
{ const { p, ms } = await conNombres({ '_xlnm.Print_Area': "'".repeat(1000000) }); prueba('Print_Area de 1 MB de comillas: no cuelga', ms < 1000 && p.areas.length === 1 && p.areas[0].f0 === 0 && p.areas[0].c0 === 0, `${ms} ms`); }
{ const { p } = await conNombres({ '_xlnm.Print_Titles': 'Carta!$0:$0,Carta!$A:$ZZZ', '_xlnm.Print_Area': 'Carta!$A$0:$ZZZ$99999999' }); prueba('fila 0, columna ZZZ y fila 99.999.999: se acotan a la hoja (sin índices -1 ni NaN)', finitos(p) && p.areas[0].f0 === 0 && p.areas[0].c1 === 16383 && p.areas[0].f1 === 1048575 && !p.titulos.filas && p.titulos.columnas.c1 === 16383, JSON.stringify([p.areas, p.titulos])); }
{ const { p } = await conNombres({}, (imp) => { imp.pagina.anchoPapel = '9'.repeat(400) + 'in'; imp.pagina.altoPapel = '9'.repeat(400) + 'in'; imp.margenes = { izq: 1e300, der: 1e300, arr: 0.75, aba: 0.75, encabezado: 0.3, pie: 0.3 }; });
  prueba('papel "999…9in" y margen 1e300: medidas finitas (el PDF escribía Infinity y 1e+303)', [p.papel.ancho, p.papel.alto, p.margenes.izq, p.anchoUtil, p.altoUtil].every(Number.isFinite) && p.papel.ancho <= 14400 && p.margenes.izq <= 14400 && finitos(p), JSON.stringify([p.papel, p.margenes])); }
{ const area = Array.from({ length: 300 }, () => 'Carta!$A$1:$XFD$1048576').join(',');
  const { p, ms } = await conNombres({ '_xlnm.Print_Area': area }, (imp) => { imp.margenes = { izq: 5, der: 5, arr: 6, aba: 6, encabezado: 0.3, pie: 0.3 }; });
  prueba('256 áreas de toda la hoja con papel sin lugar: se corta en 20.000 páginas sin recorrer el resto', ms < 1500 && p.paginas.length === 20000, `${ms} ms`);
  const r2 = await conNombres({ '_xlnm.Print_Area': area }, (imp) => { imp.ajustarAPagina = true; imp.pagina.ajustarAncho = 3; imp.pagina.ajustarAlto = 0; });
  prueba('lo mismo con "ajustar a 3 de ancho": la búsqueda de escala no repagina todo', r2.ms < 1500 && r2.p.escala === 10, `${r2.ms} ms, escala ${r2.p.escala}`); }
{ const { p } = await conNombres({ '_xlnm.Print_Titles': 'Carta!$1:$100000', '_xlnm.Print_Area': 'Carta!$A$100001:$A$1048576' });
  const t = p.paginas[0].titulos.filas;
  prueba('títulos de 100.000 filas: cada página lleva solo los que caben en ella', p.paginas.length === 20000 && t && t.f0 === 0 && t.f1 < 200, JSON.stringify(t)); }
{ const { p } = await calcular('Titulos'); const q = p.paginas.find((x) => x.titulos.filas); prueba('títulos que caben: se repiten enteros', q && q.titulos.filas.f1 === p.titulos.filas.f1); }

// encabezado y pie
const hf = XlsxPdf.analizarHF('&LIzq&CPágina &P de &N&R&"Arial,Bold"&14&F');
prueba('encabezado: secciones L/C/R', hf.izq[0].texto === 'Izq' && hf.cen.length === 4 && hf.cen[1].campo === 'P' && hf.cen[3].campo === 'N' && hf.der[0].campo === 'F');
prueba('encabezado: fuente, estilo y tamaño', hf.der[0].fuente === 'Arial' && hf.der[0].negrita === true && hf.der[0].tam === 14);
prueba('encabezado: && es un signo y lo desconocido se ignora', XlsxPdf.analizarHF('A&&B&G&Z').cen.map((r) => r.texto || r.campo).join('|') === 'A&B|Z');
prueba('encabezado: cadena enorme o sin cierre no cuelga', XlsxPdf.analizarHF('&"x' + 'a'.repeat(100000)).cen.length === 0 && XlsxPdf.analizarHF('&'.repeat(100000)).cen.every((r) => r.texto.length <= 4096));

console.log(fallos ? `\n${fallos} fallos` : '\nTodas bien');
process.exit(fallos ? 1 : 0);

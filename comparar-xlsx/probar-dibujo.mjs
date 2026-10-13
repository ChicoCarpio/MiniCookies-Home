// Pruebas del dibujo y la geometria SIN navegador ni Excel: un Canvas falso que mide 7 px por letra y anota lo que se dibuja.
// Cubre lo que encontro la revision de Opus (2026-10-13): cuadricula dentro de combinadas, palabra larguisima con ajuste, reducir hasta ajustar
// con letra enorme, combinadas hostiles, tamanos invalidos, columna oculta con ajuste, desborde al desplazarse, relleno "*" de contabilidad,
// "mostrar ceros" y nombres de fuente con barras. Uso: node comparar-xlsx/probar-dibujo.mjs (sale 1 si algo falla)
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
const raiz = path.resolve(import.meta.dirname, '..');

function lienzoFalso() {
  const registro = { rects: [], textos: [], fuentes: [] };
  const ctx = {
    font: '10px x', fillStyle: '', strokeStyle: '', lineWidth: 1, textBaseline: '', textAlign: '',
    save() {}, restore() {}, beginPath() {}, rect() {}, clip() {}, translate() {}, rotate() {}, moveTo() {}, lineTo() {}, stroke() {},
    fillRect(x, y, w, h) { registro.rects.push({ x, y, w, h, c: this.fillStyle }); },
    fillText(t, x, y) { registro.textos.push({ t, x, y }); registro.fuentes.push(this.font); },
    measureText(t) { return { width: String(t).length * 7 }; },
    createPattern() { return {}; },
  };
  return { ctx, registro };
}
globalThis.OffscreenCanvas = class { getContext() { return lienzoFalso().ctx; } };
globalThis.DOMMatrix = class { translate() { return this; } };
globalThis.PptxTexto = { familia: (n) => `"${n}"`, cargarFuentes: async () => {} };
for (const f of ['zip-propio.js', 'xml-ligero.js', 'xlsx-estilos.js', 'xlsx-modelo.js', 'xlsx-formato.js', 'xlsx-geometria.js', 'xlsx-dibujo.js']) vm.runInThisContext(fs.readFileSync(path.join(raiz, 'assets/js', f), 'utf8'), { filename: f });

let fallos = 0;
const prueba = (nombre, ok, detalle) => { console.log(`${ok ? '✓' : '✗'} ${nombre}${ok || !detalle ? '' : ' — ' + detalle}`); if (!ok) fallos++; };
const T = LibroXlsx.TIPO;
// Un libro falso de una hoja: celdas [{f,c,tipo,valor,estilo}], estilos por indice (xf), y ajustes de la hoja.
function armar({ celdas = [], xfs = [], combinadas = [], filas = [], columnas = [], formato = {}, vista = {} }) {
  const hoja = LibroXlsx.hojaVacia();
  for (const x of celdas) hoja.celdas.agregar(x.f, x.c, x.tipo === undefined ? T.texto : x.tipo, x.valor, x.estilo || 0);
  hoja.celdas._indexar();
  hoja.combinadas.push(...combinadas);
  for (const [k, v] of filas) hoja.filas.set(k, Object.assign({ alto: null, oculta: false, personalizada: false, estilo: 0, formatoPersonalizado: false, nivel: 0, colapsada: false }, v));
  hoja.columnas.push(...columnas.map((c) => Object.assign({ ancho: null, oculta: false, personalizado: false, estilo: 0 }, c)));
  Object.assign(hoja.formato, formato); Object.assign(hoja.vista, vista);
  const fuenteBase = { nombre: 'Calibri', tam: 11, negrita: false, cursiva: false, subrayado: null, tachado: false, color: null };
  const xfBase = { numFmtId: 0, formato: 'General', fuente: fuenteBase, relleno: { tipo: 'ninguno', patron: 'none', fg: null, bg: null }, borde: null, alineacion: { h: 'general', v: 'bottom', sangria: 0, rotacion: 0, ajustar: false, reducir: false }, proteccion: {}, comillaPrefijo: false };
  const lista = [xfBase, ...xfs.map((x) => Object.assign({}, xfBase, x, { fuente: Object.assign({}, fuenteBase, x.fuente), alineacion: Object.assign({}, xfBase.alineacion, x.alineacion) }))];
  const libro = { fecha1904: false, textos: [], estilos: { xf: (i) => lista[i] || lista[0], color: () => null, fuentes: [fuenteBase], paletaPropia: null } };
  return { libro, hoja };
}
const tiempo = (fn) => { const t0 = performance.now(); const r = fn(); return [performance.now() - t0, r]; };

// 1. cuadricula dentro de una combinada sin relleno
{
  const { libro, hoja } = armar({ combinadas: [{ f0: 0, c0: 0, f1: 1, c1: 2 }] });
  const g = XlsxGeometria.crear(libro, hoja);
  const { ctx, registro } = lienzoFalso();
  XlsxDibujo.dibujarRango(ctx, libro, hoja, g, { f0: 0, c0: 0, f1: 3, c1: 4 }, { x: 0, y: 0, cuadricula: true });
  const gris = registro.rects.findIndex((q) => q.c === '#E0E0E0');
  const blancoSobre = registro.rects.slice(gris).filter((q) => q.c === '#FFFFFF' && q.x === 1 && q.y === 1 && q.w === 191 && q.h === 39);
  prueba('la cuadrícula no se ve dentro de una combinada', blancoSobre.length === 1, JSON.stringify(registro.rects.filter((q) => q.c === '#FFFFFF').slice(0, 3)));
}

// 2. palabra enorme con ajuste de texto: ni se cuelga ni tarda
{
  const enorme = 'x'.repeat(32767);
  const { libro, hoja } = armar({ celdas: [{ f: 0, c: 0, valor: enorme }], xfs: [{ alineacion: { ajustar: true } }] });
  hoja.celdas.estilo[0] = 1;
  const [ms, g] = tiempo(() => XlsxDibujo.geometria(libro, hoja));
  prueba(`palabra de 32.767 letras con ajuste: geometría en ${ms.toFixed(0)} ms`, ms < 2000, ms + ' ms');
  const { ctx } = lienzoFalso();
  const [ms2] = tiempo(() => XlsxDibujo.dibujarRango(ctx, libro, hoja, g, { f0: 0, c0: 0, f1: 5, c1: 3 }, { x: 0, y: 0 }));
  prueba(`...y dibujarla: ${ms2.toFixed(0)} ms`, ms2 < 2000);
}

// 3. reducir hasta ajustar con letra de 300.000 pt
{
  const { libro, hoja } = armar({ celdas: [{ f: 0, c: 0, valor: 'hola mundo' }], xfs: [{ fuente: { tam: 300000 }, alineacion: { reducir: true } }] });
  hoja.celdas.estilo[0] = 1;
  const g = XlsxGeometria.crear(libro, hoja);
  const { ctx } = lienzoFalso();
  let llamadas = 0; const mt = ctx.measureText; ctx.measureText = (t) => { llamadas++; return mt(t); };
  const [ms] = tiempo(() => XlsxDibujo.dibujarRango(ctx, libro, hoja, g, { f0: 0, c0: 0, f1: 2, c1: 2 }, { x: 0, y: 0 }));
  prueba(`reducir con letra de 300.000 pt: ${llamadas} mediciones, ${ms.toFixed(0)} ms`, llamadas < 200 && ms < 500);
}

// 4. combinadas hostiles (20.000 copias de la hoja entera) y tamaños inválidos
{
  const hostiles = Array.from({ length: 20000 }, () => ({ f0: 0, c0: 0, f1: 1048575, c1: 16383 }));
  const { libro, hoja } = armar({ combinadas: hostiles });
  const g = XlsxGeometria.crear(libro, hoja);
  const { ctx } = lienzoFalso();
  const [ms] = tiempo(() => XlsxDibujo.dibujarRango(ctx, libro, hoja, g, { f0: 0, c0: 0, f1: 40, c1: 25 }, { x: 0, y: 0 }));
  prueba(`20.000 combinadas repetidas: ${ms.toFixed(0)} ms por cuadro`, ms < 100);
}
{
  const { libro, hoja } = armar({ celdas: [{ f: 0, c: 0, valor: 'a' }, { f: 3, c: 3, valor: 'b' }], filas: [[0, { alto: -10 }], [1, { alto: 'x' ? NaN : 0 }], [2, { alto: 1e9 }]], columnas: [{ min: 0, max: 0, ancho: -5 }, { min: 1, max: 1, ancho: 1e9 }], formato: { altoPorDefecto: -3, anchoPorDefecto: NaN } });
  const g = XlsxGeometria.crear(libro, hoja);
  let ok = true;
  for (let i = 0; i < 6; i++) { if (!(g.y(i + 1) >= g.y(i)) || !(g.x(i + 1) >= g.x(i)) || g.altoFila(i) < 0 || g.anchoCol(i) < 0 || g.altoFila(i) > 546 || g.anchoCol(i) > 1790) ok = false; }
  prueba('tamaños negativos, NaN o enormes no rompen las posiciones', ok, [0, 1, 2, 3].map((i) => g.altoFila(i) + '/' + g.anchoCol(i)).join(' '));
  const g2 = XlsxGeometria.crear(armar({ celdas: [{ f: 0, c: 0, valor: 'a' }], xfs: [{ fuente: { tam: 2e9 } }] }).libro, (() => { const q = armar({ celdas: [{ f: 0, c: 0, valor: 'a' }], xfs: [{ fuente: { tam: 2e9 } }] }); q.hoja.celdas.estilo[0] = 1; return q.hoja; })());
  prueba('letra de 2.000.000.000 pt no desborda el alto de fila', g2.altoFila(0) > 0 && g2.altoFila(0) <= 546, String(g2.altoFila(0)));
  const vacia = armar({ columnas: [{ min: 0, max: 16383, oculta: true }] });
  const g3 = XlsxGeometria.crear(vacia.libro, vacia.hoja);
  prueba('todas las columnas ocultas: colEn no pasa de XFD', g3.colEn(0) <= 16383 && g3.filaEn(1e12) <= 1048575, g3.colEn(0) + ' ' + g3.filaEn(1e12));
}

// 5. columna oculta con ajuste de texto: la fila no se dispara
{
  const { libro, hoja } = armar({ celdas: [{ f: 0, c: 1, valor: 'hello world' }], xfs: [{ alineacion: { ajustar: true } }], columnas: [{ min: 1, max: 1, oculta: true }] });
  hoja.celdas.estilo[0] = 1;
  const g = XlsxDibujo.geometria(libro, hoja);
  prueba('columna oculta con ajuste: la fila mide una línea', g.altoFila(0) === 20, String(g.altoFila(0)));
}

// 6. desborde de un texto largo cuando la vista se desplaza una columna
{
  const { libro, hoja } = armar({ celdas: [{ f: 0, c: 0, valor: 'un texto muy largo que pasa por varias columnas vacías' }] });
  const g = XlsxGeometria.crear(libro, hoja);
  const { ctx, registro } = lienzoFalso();
  XlsxDibujo.dibujarVista(ctx, libro, hoja, g, { x: 70, y: 0, ancho: 400, alto: 100, encabezados: false });
  prueba('el texto que viene de una celda anterior se sigue dibujando al desplazarse', registro.textos.some((q) => q.t.startsWith('un texto')), JSON.stringify(registro.textos.slice(0, 2)));
}

// 7. combinada centrada: su texto no se mueve según el rango visible
{
  const { libro, hoja } = armar({ celdas: [{ f: 0, c: 0, valor: 'T' }], xfs: [{ alineacion: { h: 'center' } }], combinadas: [{ f0: 0, c0: 0, f1: 0, c1: 9 }] });
  hoja.celdas.estilo[0] = 1;
  const g = XlsxGeometria.crear(libro, hoja);
  const posiciones = [9, 5, 3].map((c1) => { const { ctx, registro } = lienzoFalso(); XlsxDibujo.dibujarRango(ctx, libro, hoja, g, { f0: 0, c0: 0, f1: 1, c1 }, { x: 0, y: 0 }); return registro.textos.find((q) => q.t === 'T').x; });
  prueba('texto de una combinada que pasa del rango: misma posición', posiciones[0] === posiciones[1] && posiciones[1] === posiciones[2], posiciones.join(' '));
}

// 8. combinada con la fila superior oculta: su relleno se dibuja
{
  const rojo = { tipo: 'patron', patron: 'solid', fg: { tipo: 'rgb', valor: 'FF0000', tint: 0 }, bg: null };
  const { libro, hoja } = armar({ celdas: [{ f: 0, c: 0, valor: 'x' }], xfs: [{ relleno: rojo }], combinadas: [{ f0: 0, c0: 0, f1: 2, c1: 2 }], filas: [[0, { oculta: true }]] });
  hoja.celdas.estilo[0] = 1;
  libro.estilos.color = (ref) => (ref ? '#' + ref.valor : null);
  const g = XlsxGeometria.crear(libro, hoja);
  const { ctx, registro } = lienzoFalso();
  XlsxDibujo.dibujarRango(ctx, libro, hoja, g, { f0: 0, c0: 0, f1: 3, c1: 3 }, { x: 0, y: 0 });
  prueba('combinada con la fila superior oculta: se rellena', registro.rects.some((q) => q.c === '#FF0000'));
}

// 9. contabilidad: el "$" queda a la izquierda y el número a la derecha
{
  const { libro, hoja } = armar({ celdas: [{ f: 0, c: 0, tipo: T.numero, valor: 1234 }], xfs: [{ numFmtId: 44, formato: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)' }], columnas: [{ min: 0, max: 0, ancho: 20 }] });
  hoja.celdas.estilo[0] = 1;
  const g = XlsxGeometria.crear(libro, hoja);
  const { ctx, registro } = lienzoFalso();
  XlsxDibujo.dibujarRango(ctx, libro, hoja, g, { f0: 0, c0: 0, f1: 1, c1: 1 }, { x: 0, y: 0, idioma: 'en' });
  const dolar = registro.textos.find((q) => q.t === '$'), ultimo = registro.textos[registro.textos.length - 1];
  prueba('contabilidad: "$" a la izquierda, número a la derecha', !!dolar && dolar.x <= 15 && ultimo.x > 100, JSON.stringify([dolar, ultimo]));
}

// 10. mostrar ceros desactivado
{
  const { libro, hoja } = armar({ celdas: [{ f: 0, c: 0, tipo: T.numero, valor: 0 }, { f: 1, c: 0, tipo: T.numero, valor: 5 }], vista: { ceros: false } });
  const g = XlsxGeometria.crear(libro, hoja);
  const { ctx, registro } = lienzoFalso();
  XlsxDibujo.dibujarRango(ctx, libro, hoja, g, { f0: 0, c0: 0, f1: 2, c1: 1 }, { x: 0, y: 0 });
  prueba('showZeros="0": no se dibuja el 0 pero sí el 5', registro.textos.length === 1 && registro.textos[0].t === '5', JSON.stringify(registro.textos));
}

// 11. nombre de fuente con barra invertida
{
  const { libro, hoja } = armar({ celdas: [{ f: 0, c: 0, valor: 'x' }], xfs: [{ fuente: { nombre: 'Arial\\' } }] });
  hoja.celdas.estilo[0] = 1;
  const g = XlsxGeometria.crear(libro, hoja);
  const { ctx, registro } = lienzoFalso();
  XlsxDibujo.dibujarRango(ctx, libro, hoja, g, { f0: 0, c0: 0, f1: 1, c1: 1 }, { x: 0, y: 0 });
  prueba('nombre de fuente con "\\": la cadena de la fuente queda limpia', registro.fuentes.length > 0 && registro.fuentes.every((f) => !f.includes('\\')), registro.fuentes[0]);
}

// 12. filas con zeroHeight: la vista no recorre un millón de filas
{
  const { libro, hoja } = armar({ celdas: [{ f: 0, c: 0, valor: 'a' }, { f: 1, c: 0, valor: 'b' }], formato: { filasOcultasPorDefecto: true } });
  const g = XlsxGeometria.crear(libro, hoja);
  const { ctx } = lienzoFalso();
  const [ms, v] = tiempo(() => XlsxDibujo.dibujarVista(ctx, libro, hoja, g, { x: 0, y: 0, ancho: 400, alto: 300 }));
  prueba(`filas ocultas por defecto: la vista termina enseguida (${ms.toFixed(1)} ms, última fila ${v.f1 + 1})`, ms < 50 && v.f1 < 100);
}

console.log(fallos ? `\n${fallos} prueba(s) fallaron` : '\nTodas bien');
process.exit(fallos ? 1 : 0);

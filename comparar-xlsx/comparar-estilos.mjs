// Compara el estilo que resuelve el lector propio con el que ve Excel (volcar-estilos.ps1).
// Uso: node comparar-xlsx/comparar-estilos.mjs libro.xlsx volcado.json
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
const raiz = path.resolve(import.meta.dirname, '..');
for (const f of ['zip-propio.js', 'xml-ligero.js', 'xlsx-estilos.js', 'xlsx-modelo.js', 'xlsx-formato.js']) vm.runInThisContext(fs.readFileSync(path.join(raiz, 'assets/js', f), 'utf8'), { filename: f });

const [archivo, volcado] = process.argv.slice(2);
const libro = await LibroXlsx.abrir(new Uint8Array(fs.readFileSync(archivo)));
const celdas = JSON.parse(fs.readFileSync(volcado, 'utf8').replace(/^﻿/, ''));
const hojas = new Map();
const bgrARgb = (n) => '#' + [(n & 255), ((n >> 8) & 255), ((n >> 16) & 255)].map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
const dif = (a, b) => { if (!a || !b) return a === b ? 0 : 999; return Math.max(...[1, 3, 5].map((i) => Math.abs(parseInt(a.substr(i, 2), 16) - parseInt(b.substr(i, 2), 16)))); };
const H = { general: 1, left: -4131, center: -4108, right: -4152, fill: 5, justify: -4130, centerContinuous: 7, distributed: -4117 };
const V = { bottom: -4107, top: -4160, center: -4108, justify: -4130, distributed: -4117 };
const cuentas = {}, ejemplos = {};
const falla = (cat, msg) => { cuentas[cat] = (cuentas[cat] || 0) + 1; (ejemplos[cat] = ejemplos[cat] || []).length < 4 && ejemplos[cat].push(msg); };
let comparadas = 0, maxTint = 0;

function idxDe(hoja, f, c) {
  const k = hoja.celdas.indiceDe(f, c);
  if (k >= 0) return hoja.celdas.estilo[k];
  const fila = hoja.filas.get(f);
  if (fila && fila.formatoPersonalizado) return fila.estilo;
  const col = hoja.columnas.find((q) => q.min <= c && c <= q.max);
  return col ? col.estilo : 0;
}

for (const x of celdas) {
  if (!hojas.has(x.hoja)) hojas.set(x.hoja, await libro.cargarHoja(x.hoja));
  const hoja = hojas.get(x.hoja);
  const d = LibroXlsx.decodificarDireccion(x.dir);
  const k = hoja.celdas.indiceDe(d.f, d.c);
  const e = libro.estilos.xf(idxDe(hoja, d.f, d.c));
  comparadas++;
  const ref = libro.hojas[x.hoja].nombre + '!' + x.dir;
  // fuente
  if (e.fuente.nombre !== x.fuente.nombre) falla('fuente-nombre', `${ref} ${e.fuente.nombre} vs ${x.fuente.nombre}`);
  if (Math.abs(e.fuente.tam - x.fuente.tam) > 0.01) falla('fuente-tam', `${ref} ${e.fuente.tam} vs ${x.fuente.tam}`);
  if (e.fuente.negrita !== x.fuente.negrita) falla('negrita', ref);
  if (e.fuente.cursiva !== x.fuente.cursiva) falla('cursiva', ref);
  if (e.fuente.tachado !== x.fuente.tachado) falla('tachado', ref);
  if (!!(e.fuente.subrayado && e.fuente.subrayado !== 'none') !== (x.fuente.subrayado !== -4142)) falla('subrayado', `${ref} ${e.fuente.subrayado} vs ${x.fuente.subrayado}`);
  const colorFuente = libro.estilos.color(e.fuente.color) || '#000000';
  const df = dif(colorFuente, bgrARgb(x.fuente.color)); maxTint = Math.max(maxTint, df < 999 ? df : 0);
  if (df > 3) falla('fuente-color', `${ref} ${colorFuente} vs ${bgrARgb(x.fuente.color)}`);
  // relleno
  const hayRelleno = e.relleno.tipo !== 'ninguno';
  if (hayRelleno !== (x.relleno.patron !== -4142)) falla('relleno-hay', `${ref} ${e.relleno.patron} vs ${x.relleno.patron}`);
  else if (hayRelleno && e.relleno.patron === 'solid') {
    const c = libro.estilos.color(e.relleno.fg) || '#FFFFFF';
    const dr = dif(c, bgrARgb(x.relleno.color)); maxTint = Math.max(maxTint, dr < 999 ? dr : 0);
    if (dr > 3) falla('relleno-color', `${ref} ${c} vs ${bgrARgb(x.relleno.color)}`);
  }
  // formato de número (Excel escribe los nombres de color en su idioma y a veces reordena)
  const norm = (s) => s.replace(/\[(Red|Rojo)\]/gi, '[Red]').replace(/\[(Blue|Azul)\]/gi, '[Blue]').replace(/\\/g, '');
  if (norm(e.formato) !== norm(x.formato)) falla('formato', `${ref} ${e.formato} vs ${x.formato}`);
  // texto mostrado (Range.Text): lo que Excel dibuja en la celda. Las celdas en "####" dependen del ancho y se omiten.
  if (k >= 0 && !/^#+$/.test(x.texto)) {
    const mio = LibroXlsx.textoCelda(libro, hoja, k, { idioma: 'es', decimal: '.', miles: ',', ancho: 0 }).texto;
    // el relleno "*x" ocupa el ancho de la columna: se compara sin el largo de las tiras de espacios
    const sinTiras = (t) => t.replace(/ {2,}/g, ' ').trim();
    if (sinTiras(mio) !== sinTiras(x.texto)) falla('texto', `${ref} propio ${JSON.stringify(mio)} vs Excel ${JSON.stringify(x.texto)} (formato ${e.formato})`);
  }
  // alineación
  const a = e.alineacion;
  if ((H[a.h] ?? 1) !== x.alineacion.h) falla('alin-h', `${ref} ${a.h} vs ${x.alineacion.h}`);
  if ((V[a.v] ?? -4107) !== x.alineacion.v) falla('alin-v', `${ref} ${a.v} vs ${x.alineacion.v}`);
  if (a.ajustar !== x.alineacion.ajustar) falla('ajustar', ref);
  if (a.reducir !== x.alineacion.reducir) falla('reducir', ref);
  if (a.sangria !== x.alineacion.sangria) falla('sangria', `${ref} ${a.sangria} vs ${x.alineacion.sangria}`);
  const rot = a.rotacion === 255 ? -4166 : a.rotacion > 90 ? 90 - a.rotacion : a.rotacion;
  if (rot !== x.alineacion.rotacion && !(a.rotacion === 0 && x.alineacion.rotacion === -4128)) falla('rotacion', `${ref} ${a.rotacion} vs ${x.alineacion.rotacion}`);
  // bordes: presencia por lado
  const vecino = (df, dc, lado) => { const q = libro.estilos.xf(idxDe(hoja, d.f + df, d.c + dc)); return q.borde && q.borde[lado]; };
  const lados = { izq: ['izq', 0, -1, 'der'], der: ['der', 0, 1, 'izq'], arr: ['arr', -1, 0, 'aba'], aba: ['aba', 1, 0, 'arr'] };
  for (const [lado, [clave, df, dc, opuesto]] of Object.entries(lados)) {
    const nuestro = (e.borde && e.borde[clave]) || vecino(df, dc, opuesto), suyo = x.bordes[lado];
    if (!!nuestro !== !!suyo) falla('borde-' + lado, `${ref} ${nuestro && nuestro.estilo} vs ${suyo && suyo.estilo}`);
  }
}
console.log(`celdas comparadas ${comparadas}; diferencia máxima de color por canal ${maxTint}`);
const cats = Object.keys(cuentas);
if (!cats.length) console.log('Sin diferencias');
for (const c of cats) console.log(`  ${c}: ${cuentas[c]}  ej: ${ejemplos[c].join(' | ')}`);
process.exit(cats.length ? 1 : 0);

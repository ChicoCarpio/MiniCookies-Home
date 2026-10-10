/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. */
// Convierte las definiciones de las formas predefinidas de DrawingML
// (comparar-pptx\ecma\presetShapeDefinitions.xml, del ZIP de ECMA-376 Parte 1,
// 5.ª edición, 2016: OfficeOpenXML-DrawingMLGeometries.zip) a un .js que el
// motor de diapositivas carga (vendor/ecma-376-geometrias/5ed-2016/).
//
// Solo cambia el formato (XML → JSON), como permite el aviso de Ecma: las
// fórmulas, los nombres y los números van tal cual. Se omiten las partes que
// el motor no usa (manijas de ajuste ahLst y puntos de conexión cxnLst).
//
// Uso: node comparar-pptx/convertir-geometrias.mjs

import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';

const AQUI = decodeURIComponent(new URL('.', import.meta.url).pathname).replace(/^\/([A-Za-z]:)/, '$1');
const ORIGEN = join(AQUI, 'ecma', 'presetShapeDefinitions.xml');
const DESTINO = join(AQUI, '..', 'vendor', 'ecma-376-geometrias', '5ed-2016');

// Lector mínimo de XML: el archivo de Ecma es regular (sin CDATA ni entidades
// raras), así que alcanza con etiquetas y atributos.
function parsear(texto) {
  const raiz = { nombre: '#raiz', at: {}, hijos: [] };
  const pila = [raiz];
  const re = /<(\/?)([A-Za-z_][\w:.-]*)((?:\s+[\w:.-]+\s*=\s*"[^"]*")*)\s*(\/?)>|<\?[^>]*\?>|<!--[\s\S]*?-->/g;
  let m;
  while ((m = re.exec(texto))) {
    if (!m[2]) continue;
    if (m[1]) { pila.pop(); continue; }
    const at = {};
    for (const a of m[3].matchAll(/([\w:.-]+)\s*=\s*"([^"]*)"/g)) at[a[1]] = a[2];
    const el = { nombre: m[2].replace(/^.*:/, ''), at, hijos: [] };
    pila[pila.length - 1].hijos.push(el);
    if (!m[4]) pila.push(el);
  }
  return raiz;
}
const hijo = (el, n) => el.hijos.find((h) => h.nombre === n);

const COMANDOS = { moveTo: 'M', lnTo: 'L', arcTo: 'A', quadBezTo: 'Q', cubicBezTo: 'C', close: 'Z' };
const raiz = parsear(readFileSync(ORIGEN, 'utf8')).hijos[0];
const formas = {};
for (const forma of raiz.hijos) {
  const def = {};
  const guias = (n) => (hijo(forma, n) ? hijo(forma, n).hijos.map((g) => [g.at.name, g.at.fmla]) : []);
  const av = guias('avLst'), gd = guias('gdLst');
  if (av.length) def.av = av;
  if (gd.length) def.gd = gd;
  const rect = hijo(forma, 'rect');
  if (rect) def.rect = [rect.at.l, rect.at.t, rect.at.r, rect.at.b];
  def.p = (hijo(forma, 'pathLst') ? hijo(forma, 'pathLst').hijos : []).map((p) => {
    const ruta = {};
    for (const a of ['w', 'h', 'fill', 'stroke', 'extrusionOk']) if (p.at[a] != null) ruta[a] = /^\d+$/.test(p.at[a]) ? +p.at[a] : p.at[a];
    ruta.c = p.hijos.map((c) => {
      const letra = COMANDOS[c.nombre];
      if (!letra) throw new Error(`Comando desconocido ${c.nombre} en ${forma.nombre}`);
      if (letra === 'A') return [letra, c.at.wR, c.at.hR, c.at.stAng, c.at.swAng];
      return [letra, ...c.hijos.flatMap((pt) => [pt.at.x, pt.at.y])];
    });
    return ruta;
  });
  formas[forma.nombre] = def;
}

const AVISO = `/*! Definiciones de las formas predefinidas de DrawingML (presetShapeDefinitions.xml),
 * ECMA-376 Parte 1, 5.ª edición (2016), © Ecma International. Convertidas de XML
 * a JSON sin cambiar su contenido; se omiten las manijas y los puntos de
 * conexión. Uso bajo el aviso de copyright y licencia de Ecma: ver
 * AVISO-ECMA.txt en esta misma carpeta. */\n`;
mkdirSync(DESTINO, { recursive: true });
writeFileSync(join(DESTINO, 'geometrias.js'), AVISO + 'window.GeometriasEcma = ' + JSON.stringify(formas) + ';\n');
console.log(`${Object.keys(formas).length} formas → ${join(DESTINO, 'geometrias.js')}`);

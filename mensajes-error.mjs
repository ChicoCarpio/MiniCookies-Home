// Comprobación del build para los errores con código (assets/js/mensajes-error.js; plan en
// Investigaciones/investigacion-mensajes-error/00-plan.md). Falla el build si:
//   - un errorMC('código', …, 'texto') usa un código sin su texto en «es» o en «en» del catálogo, o con un texto en español distinto del
//     catálogo (los dos tienen que decir lo mismo: el de la llamada es el que se ve si el catálogo no cargó);
//   - el catálogo tiene un código que nadie usa, o un idioma tiene claves que el otro no tiene;
//   - una página carga un script que usa errorMC (o flujos.js o lector-hojas.js, que cargan motores por dentro) sin cargar antes mensajes-error.js.
import { readdirSync, readFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import vm from 'node:vm';

// El final de una llamada que empieza en «(»: salta las cadenas y los comentarios.
function finDeLlamada(src, abre) {
  let prof = 0;
  for (let i = abre; i < src.length; i++) {
    const c = src[i];
    if (c === '"' || c === "'" || c === '`') { i++; while (i < src.length && src[i] !== c) { if (src[i] === '\\') i++; i++; } continue; }
    if (c === '/' && src[i + 1] === '/') { while (i < src.length && src[i] !== '\n') i++; continue; }
    if (c === '(') prof++;
    else if (c === ')' && --prof === 0) return i;
  }
  return -1;
}

export function revisarMensajesError(raiz) {
  const problemas = [];
  const dirJs = join(raiz, 'assets', 'js');
  const scripts = [];
  (function recorrer(dir) {
    for (const e of readdirSync(dir, { withFileTypes: true })) {
      const ruta = join(dir, e.name);
      if (e.isDirectory()) recorrer(ruta); else if (e.name.endsWith('.js')) scripts.push(ruta);
    }
  })(dirJs);

  const contexto = { self: {} };
  vm.runInNewContext(readFileSync(join(dirJs, 'mensajes-error.js'), 'utf8'), contexto);
  const catalogo = contexto.self.MensajesError;

  const usados = new Map();       // código → texto en español de la llamada
  const conErrorMC = new Set();   // nombres de archivo que lanzan errorMC
  for (const ruta of scripts) {
    if (ruta.endsWith('mensajes-error.js')) continue;
    const src = readFileSync(ruta, 'utf8');
    for (const m of src.matchAll(/\berrorMC\(\s*/g)) {
      const fin = finDeLlamada(src, src.indexOf('(', m.index));
      const llamada = src.slice(src.indexOf('(', m.index) + 1, fin);
      const codigo = (llamada.match(/^\s*'([\w-]+)'/) || [])[1];
      const texto = (llamada.match(/'((?:[^'\\]|\\.)*)'\s*$/) || [])[1];
      const donde = `${relative(raiz, ruta).replace(/\\/g, '/')}:${src.slice(0, m.index).split('\n').length}`;
      if (!codigo || texto === undefined || codigo === texto) { problemas.push(`${donde}: errorMC sin código o sin texto en español literal`); continue; }
      conErrorMC.add(ruta.slice(dirJs.length + 1).replace(/\\/g, '/'));
      if (usados.has(codigo) && usados.get(codigo) !== texto) problemas.push(`${donde}: el código ${codigo} se usa con dos textos distintos`);
      usados.set(codigo, texto);
      const clave = 'error-' + codigo;
      if (!catalogo.es[clave]) problemas.push(`${donde}: falta «${clave}» en español en mensajes-error.js`);
      else if (catalogo.es[clave] !== texto.replace(/\\'/g, "'")) problemas.push(`${donde}: el texto de «${codigo}» no es el mismo que el del catálogo`);
      if (!catalogo.en[clave]) problemas.push(`${donde}: falta «${clave}» en inglés en mensajes-error.js`);
    }
  }
  for (const idioma of ['es', 'en']) {
    const otro = idioma === 'es' ? 'en' : 'es';
    for (const clave of Object.keys(catalogo[idioma])) {
      if (!(clave in catalogo[otro])) problemas.push(`mensajes-error.js: «${clave}» está en ${idioma} y no en ${otro}`);
      if (idioma === 'es' && !usados.has(clave.replace(/^error-/, ''))) problemas.push(`mensajes-error.js: «${clave}» no lo usa ningún errorMC`);
    }
  }

  // Las páginas: lo que usa errorMC (y flujos.js y lector-hojas.js, que cargan por dentro motores que lo usan) necesita mensajes-error.js antes.
  const necesitan = new Set([...conErrorMC].map((f) => f.split('/').pop()).concat('flujos.js', 'lector-hojas.js'));
  for (const carpeta of ['herramientas', '.']) {
    for (const archivo of readdirSync(join(raiz, carpeta)).filter((f) => f.endsWith('.html'))) {
      const orden = [...readFileSync(join(raiz, carpeta, archivo), 'utf8').matchAll(/<script[^>]*\bsrc="[^"]*?([\w.-]+\.js)(?:\?[^"]*)?"/g)].map((m) => m[1]);
      const catalogoEn = orden.indexOf('mensajes-error.js');
      for (const s of orden) {
        if (necesitan.has(s) && !(catalogoEn >= 0 && catalogoEn < orden.indexOf(s))) problemas.push(`${carpeta}/${archivo}: ${s} lanza errores con código y necesita cargar antes mensajes-error.js`);
      }
    }
  }
  if (problemas.length) throw new Error('Mensajes de error:\n  - ' + problemas.join('\n  - '));
  return usados.size;
}

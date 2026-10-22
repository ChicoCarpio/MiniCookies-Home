// Orden de carga de los módulos propios del motor de Excel: una página que carga uno tiene que cargar antes lo que él usa.
// Sin esto el fallo no se ve al abrir la página sino al usar la herramienta ("XmlLigero is not defined").
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';

const NECESITA = {
  'docx-base.js': ['zip-propio.js', 'xml-ligero.js'],
  'docx-estilos.js': ['docx-base.js', 'zip-propio.js', 'xml-ligero.js'],
  'docx-lector.js': ['docx-base.js', 'zip-propio.js', 'xml-ligero.js'],
  'xlsx-paquete.js': ['zip-propio.js', 'xml-ligero.js', 'xml-editar.js'],
  'xlsx-modelo.js': ['zip-propio.js', 'xml-ligero.js', 'xlsx-formato.js', 'xlsx-estilos.js'],
  'lector-hojas.js': ['zip-propio.js', 'xml-ligero.js', 'xlsx-formato.js', 'xlsx-estilos.js', 'xlsx-modelo.js', 'csv-planilla.js'],
  'ods-lector.js': ['zip-propio.js', 'xml-ligero.js', 'xlsx-modelo.js'],
  'xlsb-lector.js': ['zip-propio.js', 'xml-ligero.js', 'xlsx-estilos.js', 'xlsx-modelo.js'],
  'xls-lector.js': ['xlsx-estilos.js', 'xlsx-modelo.js'],
  'motor-hojas.js': ['vendor.js'],
  'rejilla-hoja.js': ['vendor.js', 'lector-hojas.js', 'motor-hojas.js', 'rejilla-orden-filtro.js'],
  'excel-a-pdf-motor.js': ['vendor.js', 'motor-hojas.js', 'lector-hojas.js'],
  'xml2003-lector.js': ['xml-ligero.js', 'xlsx-modelo.js', 'xlsx-formato.js'],
  'html-tabla-lector.js': ['csv-planilla.js', 'xlsx-modelo.js'],
  'formatos-antiguos.js': ['zip-propio.js', 'xlsx-formato.js', 'xlsx-estilos.js', 'xlsx-escritor.js', 'ods-lector.js', 'xlsb-lector.js', 'xls-lector.js', 'xml2003-lector.js'],
  'csv-planilla.js': ['xlsx-estilos.js', 'xlsx-modelo.js'],
  'xlsx-escritor.js': ['zip-propio.js', 'xml-editar.js', 'xlsx-paquete.js', 'xlsx-modelo.js'],
};

export function revisarDependenciasScripts(raiz) {
  const problemas = [];
  const dir = join(raiz, 'herramientas');
  for (const archivo of readdirSync(dir).filter((f) => f.endsWith('.html'))) {
    const orden = [...readFileSync(join(dir, archivo), 'utf8').matchAll(/<script[^>]*\bsrc="[^"]*?([\w.-]+\.js)"/g)].map((m) => m[1]);
    for (const [modulo, previos] of Object.entries(NECESITA)) {
      const pos = orden.indexOf(modulo);
      if (pos < 0) continue;
      for (const p of previos) if (!(orden.indexOf(p) >= 0 && orden.indexOf(p) < pos)) problemas.push(`${archivo}: ${modulo} necesita cargar antes ${p}`);
    }
  }
  if (problemas.length) throw new Error('Scripts sin sus dependencias:\n  - ' + problemas.join('\n  - '));
}

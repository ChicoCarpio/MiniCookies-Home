/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. */
// Las páginas que traen el motor propio de PowerPoint en <script> lo traen
// entero y en el orden de la lista única (SCRIPTS_MOTOR_PPTX, en
// assets/js/pptx-visual.js). Esto lo comprueba en cada build.
//
// Por qué existe: cada página tenía su copia de la lista; cuando el motor
// propio pasó a ser el predeterminado (2026-10-10), Unir PDF se quedó con la
// lista vieja y siguió dibujando con PptxViewJS sin que nada avisara.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';

const PAGINAS = ['herramientas/powerpoint-a-pdf.html', 'herramientas/powerpoint-a-imagenes.html'];

export function revisarMotorPptx(raiz) {
  const visual = readFileSync(join(raiz, 'assets/js/pptx-visual.js'), 'utf8');
  const m = visual.match(/const SCRIPTS_MOTOR_PPTX = \[([^\]]*)\]/);
  if (!m) throw new Error('No se encontró SCRIPTS_MOTOR_PPTX en assets/js/pptx-visual.js');
  const lista = [...m[1].matchAll(/'([^']+)'/g)].map((x) => (x[1].startsWith('vendor/') ? '../' + x[1] : '../assets/js/' + x[1]));
  for (const pagina of PAGINAS) {
    const html = readFileSync(join(raiz, pagina), 'utf8');
    const scripts = [...html.matchAll(/<script[^>]*\ssrc="([^"]+)"/g)].map((x) => x[1].replace(/\?.*$/, ''));
    let desde = 0;
    for (const src of lista) {
      const i = scripts.indexOf(src, desde);
      if (i < 0) throw new Error(`${pagina}: falta ${src} del motor de PowerPoint, o no está en el orden de SCRIPTS_MOTOR_PPTX (assets/js/pptx-visual.js)`);
      desde = i + 1;
    }
  }
  return PAGINAS.length;
}

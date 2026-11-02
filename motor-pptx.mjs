/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Las páginas que traen el motor propio de PowerPoint en <script> lo traen
// entero y en el orden de la lista única (SCRIPTS_MOTOR_PPTX, en
// assets/js/pptx-visual.js). Esto lo comprueba en cada build.
//
// Por qué existe: cada página tenía su copia de la lista; cuando el motor
// propio pasó a ser el predeterminado (2026-10-10), Unir PDF se quedó con la
// lista vieja y siguió dibujando con la librería de antes sin que nada avisara.
//
// Además: grafico-suelto.js (los gráficos de Excel y de "Gráfico desde Excel /
// CSV") trae su propia lista, un subconjunto del motor; tiene que seguir el
// mismo orden que la lista única.

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
  // La lista de grafico-suelto.js: los mismos archivos, en el mismo orden.
  const suelto = readFileSync(join(raiz, 'assets/js/grafico-suelto.js'), 'utf8').match(/const SCRIPTS_MOTOR = \[([^\]]*)\]/);
  if (!suelto) throw new Error('No se encontró SCRIPTS_MOTOR en assets/js/grafico-suelto.js');
  let desdeLista = 0;
  for (const x of suelto[1].matchAll(/'([^']+)'/g)) {
    const i = lista.indexOf(x[1].startsWith('vendor/') ? '../' + x[1] : '../assets/js/' + x[1], desdeLista);
    if (i < 0) throw new Error(`assets/js/grafico-suelto.js: ${x[1]} no está en SCRIPTS_MOTOR_PPTX (assets/js/pptx-visual.js), o está en otro orden`);
    desdeLista = i + 1;
  }
  return PAGINAS.length;
}

/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Galería de formas de Mini Cookies Presentaciones (Insertar → Formas): un diálogo con las formas predefinidas de la norma, agrupadas como en PowerPoint.
//
// ECMA-376: las definiciones de las formas (vendor/ecma-376-geometrias/5ed-2016/geometrias.js, window.GeometriasEcma) están bajo el aviso de
// copyright de Ecma (AVISO-ECMA.txt, al lado del archivo). Este módulo NO copia ni cambia esas definiciones: las lee tal cual a través de
// PptxGeometria (pptx-geometria.js) solo para dibujar la miniatura de cada forma, y lo que se inserta en un archivo es el nombre de la forma
// (<a:prstGeom prst="…">) — es decir, implementar la norma en un producto conforme (punto iv del aviso), no republicarla ni usarla para otra cosa.
// Los nombres que se ven (español e inglés) son etiquetas de interfaz propias; la lista de formas es la de la norma.
//
//   const r = await GaleriaFormas.abrir({ idioma: 'es' | 'en' })   → { prst, nombre, flecha?, ancho, alto (cm) } | null
(function () {
  // [clave de la forma, español, inglés]. Las que no estén en GeometriasEcma no se muestran.
  const L = (id, es, en, extra) => Object.assign({ id, prst: id, es, en }, extra || {});
  const CATEGORIAS = [
    ['gal-lineas', [
      L('line', 'Línea', 'Line', { ancho: 4, alto: 0 }), L('flecha', 'Flecha', 'Arrow', { prst: 'straightConnector1', flecha: true, ancho: 4, alto: 0 }),
      L('bentConnector3', 'Conector: angular', 'Connector: Elbow', { ancho: 4, alto: 1.5 }), L('curvedConnector3', 'Conector: curvado', 'Connector: Curved', { ancho: 4, alto: 1.5 }),
    ]],
    ['gal-rectangulos', [
      L('rect', 'Rectángulo', 'Rectangle'), L('roundRect', 'Rectángulo: esquinas redondeadas', 'Rectangle: Rounded Corners'),
      L('snip1Rect', 'Rectángulo: una esquina cortada', 'Rectangle: Single Corner Snipped'), L('snip2SameRect', 'Rectángulo: esquinas superiores cortadas', 'Rectangle: Top Corners Snipped'),
      L('snip2DiagRect', 'Rectángulo: esquinas opuestas cortadas', 'Rectangle: Diagonal Corners Snipped'), L('snipRoundRect', 'Rectángulo: una esquina cortada y otra redondeada', 'Rectangle: One Corner Snipped, One Rounded'),
      L('round1Rect', 'Rectángulo: una esquina redondeada', 'Rectangle: Single Corner Rounded'), L('round2SameRect', 'Rectángulo: esquinas superiores redondeadas', 'Rectangle: Top Corners Rounded'),
      L('round2DiagRect', 'Rectángulo: esquinas opuestas redondeadas', 'Rectangle: Diagonal Corners Rounded'),
    ]],
    ['gal-basicas', [
      L('ellipse', 'Elipse', 'Oval'), L('triangle', 'Triángulo isósceles', 'Isosceles Triangle'), L('rtTriangle', 'Triángulo rectángulo', 'Right Triangle'), L('parallelogram', 'Paralelogramo', 'Parallelogram'),
      L('trapezoid', 'Trapecio', 'Trapezoid'), L('diamond', 'Rombo', 'Diamond'), L('pentagon', 'Pentágono regular', 'Regular Pentagon'), L('hexagon', 'Hexágono', 'Hexagon'),
      L('heptagon', 'Heptágono', 'Heptagon'), L('octagon', 'Octágono', 'Octagon'), L('decagon', 'Decágono', 'Decagon'), L('dodecagon', 'Dodecágono', 'Dodecagon'),
      L('pie', 'Sector circular', 'Partial Circle'), L('chord', 'Cuerda', 'Chord'), L('teardrop', 'Lágrima', 'Teardrop'), L('frame', 'Marco', 'Frame'), L('halfFrame', 'Medio marco', 'Half Frame'),
      L('corner', 'Forma de L', 'L-Shape'), L('diagStripe', 'Franja diagonal', 'Diagonal Stripe'), L('plus', 'Cruz', 'Cross'), L('plaque', 'Placa', 'Plaque'), L('can', 'Cilindro', 'Cylinder'),
      L('cube', 'Cubo', 'Cube'), L('bevel', 'Bisel', 'Bevel'), L('donut', 'Rosquilla', 'Donut'), L('noSmoking', 'Símbolo de prohibido', '"No" Symbol'), L('blockArc', 'Arco de bloque', 'Block Arc'),
      L('foldedCorner', 'Esquina doblada', 'Folded Corner'), L('smileyFace', 'Cara sonriente', 'Smiley Face'), L('heart', 'Corazón', 'Heart'), L('lightningBolt', 'Rayo', 'Lightning Bolt'),
      L('sun', 'Sol', 'Sun'), L('moon', 'Luna', 'Moon'), L('cloud', 'Nube', 'Cloud'), L('arc', 'Arco', 'Arc'), L('bracketPair', 'Corchetes dobles', 'Double Bracket'), L('bracePair', 'Llaves dobles', 'Double Brace'),
      L('leftBracket', 'Corchete izquierdo', 'Left Bracket'), L('rightBracket', 'Corchete derecho', 'Right Bracket'), L('leftBrace', 'Llave izquierda', 'Left Brace'), L('rightBrace', 'Llave derecha', 'Right Brace'),
    ]],
    ['gal-flechas', [
      L('rightArrow', 'Flecha: a la derecha', 'Arrow: Right'), L('leftArrow', 'Flecha: a la izquierda', 'Arrow: Left'), L('upArrow', 'Flecha: hacia arriba', 'Arrow: Up'), L('downArrow', 'Flecha: hacia abajo', 'Arrow: Down'),
      L('leftRightArrow', 'Flecha: izquierda y derecha', 'Arrow: Left-Right'), L('upDownArrow', 'Flecha: arriba y abajo', 'Arrow: Up-Down'), L('quadArrow', 'Flecha: cuádruple', 'Arrow: Quad'),
      L('leftRightUpArrow', 'Flecha: izquierda, derecha y arriba', 'Arrow: Left-Right-Up'), L('bentArrow', 'Flecha: doblada', 'Arrow: Bent'), L('uturnArrow', 'Flecha: en U', 'Arrow: U-Turn'),
      L('leftUpArrow', 'Flecha: izquierda y arriba', 'Arrow: Left-Up'), L('bentUpArrow', 'Flecha: doblada hacia arriba', 'Arrow: Bent-Up'), L('curvedRightArrow', 'Flecha: curvada a la derecha', 'Arrow: Curved Right'),
      L('curvedLeftArrow', 'Flecha: curvada a la izquierda', 'Arrow: Curved Left'), L('curvedUpArrow', 'Flecha: curvada hacia arriba', 'Arrow: Curved Up'), L('curvedDownArrow', 'Flecha: curvada hacia abajo', 'Arrow: Curved Down'),
      L('stripedRightArrow', 'Flecha: a la derecha con rayas', 'Arrow: Striped Right'), L('notchedRightArrow', 'Flecha: a la derecha con muesca', 'Arrow: Notched Right'), L('homePlate', 'Flecha: pentágono', 'Arrow: Pentagon'),
      L('chevron', 'Flecha: divisa', 'Arrow: Chevron'), L('rightArrowCallout', 'Llamada: flecha a la derecha', 'Callout: Right Arrow'), L('downArrowCallout', 'Llamada: flecha hacia abajo', 'Callout: Down Arrow'),
      L('leftArrowCallout', 'Llamada: flecha a la izquierda', 'Callout: Left Arrow'), L('upArrowCallout', 'Llamada: flecha hacia arriba', 'Callout: Up Arrow'), L('leftRightArrowCallout', 'Llamada: flecha izquierda y derecha', 'Callout: Left-Right Arrow'),
      L('quadArrowCallout', 'Llamada: flecha cuádruple', 'Callout: Quad Arrow'), L('circularArrow', 'Flecha: circular', 'Arrow: Circular'),
    ]],
    ['gal-ecuaciones', [
      L('mathPlus', 'Signo más', 'Math Plus'), L('mathMinus', 'Signo menos', 'Math Minus'), L('mathMultiply', 'Signo de multiplicación', 'Math Multiply'),
      L('mathDivide', 'Signo de división', 'Math Divide'), L('mathEqual', 'Signo igual', 'Math Equal'), L('mathNotEqual', 'Signo distinto', 'Math Not Equal'),
    ]],
    ['gal-flujo', [
      L('flowChartProcess', 'Diagrama de flujo: proceso', 'Flowchart: Process'), L('flowChartAlternateProcess', 'Diagrama de flujo: proceso alternativo', 'Flowchart: Alternate Process'),
      L('flowChartDecision', 'Diagrama de flujo: decisión', 'Flowchart: Decision'), L('flowChartInputOutput', 'Diagrama de flujo: datos', 'Flowchart: Data'),
      L('flowChartPredefinedProcess', 'Diagrama de flujo: proceso predefinido', 'Flowchart: Predefined Process'), L('flowChartInternalStorage', 'Diagrama de flujo: almacenamiento interno', 'Flowchart: Internal Storage'),
      L('flowChartDocument', 'Diagrama de flujo: documento', 'Flowchart: Document'), L('flowChartMultidocument', 'Diagrama de flujo: varios documentos', 'Flowchart: Multidocument'),
      L('flowChartTerminator', 'Diagrama de flujo: terminador', 'Flowchart: Terminator'), L('flowChartPreparation', 'Diagrama de flujo: preparación', 'Flowchart: Preparation'),
      L('flowChartManualInput', 'Diagrama de flujo: entrada manual', 'Flowchart: Manual Input'), L('flowChartManualOperation', 'Diagrama de flujo: operación manual', 'Flowchart: Manual Operation'),
      L('flowChartConnector', 'Diagrama de flujo: conector', 'Flowchart: Connector'), L('flowChartOffpageConnector', 'Diagrama de flujo: conector fuera de página', 'Flowchart: Off-page Connector'),
      L('flowChartPunchedCard', 'Diagrama de flujo: tarjeta', 'Flowchart: Card'), L('flowChartPunchedTape', 'Diagrama de flujo: cinta perforada', 'Flowchart: Punched Tape'),
      L('flowChartSummingJunction', 'Diagrama de flujo: unión de suma', 'Flowchart: Summing Junction'), L('flowChartOr', 'Diagrama de flujo: o', 'Flowchart: Or'),
      L('flowChartCollate', 'Diagrama de flujo: intercalar', 'Flowchart: Collate'), L('flowChartSort', 'Diagrama de flujo: ordenar', 'Flowchart: Sort'),
      L('flowChartExtract', 'Diagrama de flujo: extracción', 'Flowchart: Extract'), L('flowChartMerge', 'Diagrama de flujo: combinar', 'Flowchart: Merge'),
      L('flowChartOnlineStorage', 'Diagrama de flujo: datos almacenados', 'Flowchart: Stored Data'), L('flowChartOfflineStorage', 'Diagrama de flujo: almacenamiento fuera de línea', 'Flowchart: Offline Storage'),
      L('flowChartMagneticTape', 'Diagrama de flujo: almacenamiento secuencial', 'Flowchart: Sequential Access Storage'), L('flowChartMagneticDisk', 'Diagrama de flujo: disco magnético', 'Flowchart: Magnetic Disk'),
      L('flowChartMagneticDrum', 'Diagrama de flujo: almacenamiento de acceso directo', 'Flowchart: Direct Access Storage'), L('flowChartDisplay', 'Diagrama de flujo: visualizar', 'Flowchart: Display'),
      L('flowChartDelay', 'Diagrama de flujo: retraso', 'Flowchart: Delay'),
    ]],
    ['gal-estrellas', [
      L('irregularSeal1', 'Explosión 1', 'Explosion 1'), L('irregularSeal2', 'Explosión 2', 'Explosion 2'),
      ...[4, 5, 6, 7, 8, 10, 12, 16, 24, 32].map((n) => L('star' + n, 'Estrella: de ' + n + ' puntas', 'Star: ' + n + ' Points')),
      L('ribbon2', 'Cinta: inclinada hacia arriba', 'Ribbon: Tilted Up'), L('ribbon', 'Cinta: inclinada hacia abajo', 'Ribbon: Tilted Down'),
      L('ellipseRibbon2', 'Cinta: curvada hacia arriba', 'Ribbon: Curved and Tilted Up'), L('ellipseRibbon', 'Cinta: curvada hacia abajo', 'Ribbon: Curved and Tilted Down'),
      L('verticalScroll', 'Pergamino: vertical', 'Scroll: Vertical'), L('horizontalScroll', 'Pergamino: horizontal', 'Scroll: Horizontal'), L('wave', 'Onda', 'Wave'), L('doubleWave', 'Onda doble', 'Double Wave'),
    ]],
    ['gal-llamadas', [
      L('wedgeRectCallout', 'Bocadillo: rectángulo', 'Speech Bubble: Rectangle'), L('wedgeRoundRectCallout', 'Bocadillo: rectángulo con esquinas redondeadas', 'Speech Bubble: Rectangle with Corners Rounded'),
      L('wedgeEllipseCallout', 'Bocadillo: ovalado', 'Speech Bubble: Oval'), L('cloudCallout', 'Bocadillo: nube', 'Thought Bubble: Cloud'),
      ...[1, 2, 3].map((n) => L('borderCallout' + n, 'Llamada: ' + ['línea', 'línea doblada', 'línea doblada dos veces'][n - 1], 'Callout: ' + ['Line', 'Bent Line', 'Double Bent Line'][n - 1])),
      ...[1, 2, 3].map((n) => L('accentCallout' + n, 'Llamada: ' + ['línea', 'línea doblada', 'línea doblada dos veces'][n - 1] + ' con barra de énfasis', 'Callout: ' + ['Line', 'Bent Line', 'Double Bent Line'][n - 1] + ' with Accent Bar')),
      ...[1, 2, 3].map((n) => L('callout' + n, 'Llamada: ' + ['línea', 'línea doblada', 'línea doblada dos veces'][n - 1] + ' sin borde', 'Callout: ' + ['Line', 'Bent Line', 'Double Bent Line'][n - 1] + ' with No Border')),
      ...[1, 2, 3].map((n) => L('accentBorderCallout' + n, 'Llamada: ' + ['línea', 'línea doblada', 'línea doblada dos veces'][n - 1] + ' con borde y barra de énfasis', 'Callout: ' + ['Line', 'Bent Line', 'Double Bent Line'][n - 1] + ' with Border and Accent Bar')),
    ]],
  ];
  const TEXTOS = {
    es: { 'gal-titulo': 'Formas', 'gal-buscar': 'Buscar una forma', 'gal-cerrar': 'Cerrar', 'gal-lineas': 'Líneas', 'gal-rectangulos': 'Rectángulos', 'gal-basicas': 'Formas básicas', 'gal-flechas': 'Flechas de bloque', 'gal-ecuaciones': 'Formas de ecuación', 'gal-flujo': 'Diagrama de flujo', 'gal-estrellas': 'Estrellas y cintas', 'gal-llamadas': 'Llamadas', 'gal-nada': 'Ninguna forma se llama así' },
    en: { 'gal-titulo': 'Shapes', 'gal-buscar': 'Search for a shape', 'gal-cerrar': 'Close', 'gal-lineas': 'Lines', 'gal-rectangulos': 'Rectangles', 'gal-basicas': 'Basic Shapes', 'gal-flechas': 'Block Arrows', 'gal-ecuaciones': 'Equation Shapes', 'gal-flujo': 'Flowchart', 'gal-estrellas': 'Stars and Banners', 'gal-llamadas': 'Callouts', 'gal-nada': 'No shape has that name' },
  };
  const pagina = window.i18nPagina || (window.i18nPagina = {});
  pagina.es = Object.assign(pagina.es || {}, TEXTOS.es);
  pagina.en = Object.assign(pagina.en || {}, TEXTOS.en);
  const CONECTORES = ['line', 'straightConnector1', 'bentConnector3', 'curvedConnector3'];
  const sinTildes = (s) => String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
  const hay = (f) => CONECTORES.includes(f.prst) || (typeof GeometriasEcma !== 'undefined' && !!GeometriasEcma[f.prst]);

  // La miniatura: el trazado de la forma (pptx-geometria.js, definiciones de la norma sin cambios) en un cuadro de 40 px
  function miniatura(canvas, prst) {
    const ctx = canvas.getContext('2d'), N = canvas.width, margen = N * 0.08, lado = N - 2 * margen;
    ctx.clearRect(0, 0, N, N);
    const def = PptxGeometria.predefinida(prst);
    if (!def) return;
    const g = PptxGeometria.trazar(def, lado, lado);
    ctx.save(); ctx.translate(margen, margen);
    ctx.lineWidth = Math.max(1, N / 28); ctx.lineJoin = 'round'; ctx.fillStyle = '#9db7e0'; ctx.strokeStyle = '#2f528f';
    for (const r of g.rutas) {
      const p = new Path2D();
      for (const c of r.cmds) {
        if (c[0] === 'M') p.moveTo(c[1], c[2]); else if (c[0] === 'L') p.lineTo(c[1], c[2]);
        else if (c[0] === 'C') p.bezierCurveTo(c[1], c[2], c[3], c[4], c[5], c[6]); else if (c[0] === 'Z') p.closePath();
      }
      if (r.relleno !== 'none') ctx.fill(p);
      if (r.trazo) ctx.stroke(p);
    }
    ctx.restore();
  }

  function abrir(opciones) {
    const idioma = opciones && opciones.idioma === 'en' ? 'en' : 'es';
    const t = (k) => TEXTOS[idioma][k];
    const el = (tag, props, hijos) => { const e = document.createElement(tag); for (const k in props || {}) { if (k === 'texto') e.textContent = props[k]; else e.setAttribute(k, props[k]); } for (const h of hijos || []) e.append(h); return e; };
    const dlg = el('dialog', { class: 'of-dialogo pr-galeria', 'aria-labelledby': 'galTitulo' });
    const buscar = el('input', { type: 'search', class: 'pr-gal-buscar', 'aria-label': t('gal-buscar'), placeholder: t('gal-buscar'), autocomplete: 'off', spellcheck: 'false' });
    const cerrar = el('button', { type: 'button', class: 'pr-gal-cerrar', 'aria-label': t('gal-cerrar') }, ['✕']);
    const cuerpo = el('div', { class: 'pr-gal-cuerpo' });
    const vacio = el('p', { class: 'pr-gal-vacio', texto: t('gal-nada'), hidden: '' });
    dlg.append(el('div', { class: 'pr-gal-cabecera' }, [el('h2', { id: 'galTitulo', texto: t('gal-titulo') }), buscar, cerrar]), cuerpo, vacio);
    const botones = [];
    for (const [clave, formas] of CATEGORIAS) {
      const visibles = formas.filter(hay);
      if (!visibles.length) continue;
      const rejilla = el('div', { class: 'pr-gal-rejilla', role: 'group', 'aria-label': t(clave) });
      for (const f of visibles) {
        const nombre = idioma === 'en' ? f.en : f.es;
        const c = el('canvas', { width: '80', height: '80', 'aria-hidden': 'true' });
        const b = el('button', { type: 'button', class: 'pr-gal-forma', title: nombre, 'aria-label': nombre, 'data-prst': f.prst }, [c]);
        b._forma = f; b._nombre = sinTildes(nombre);
        try { miniatura(c, f.prst); } catch (e) { console.warn('miniatura', f.prst, e); }
        rejilla.append(b); botones.push(b);
      }
      const seccion = el('section', { class: 'pr-gal-seccion' }, [el('h3', { texto: t(clave) }), rejilla]);
      cuerpo.append(seccion);
    }
    buscar.addEventListener('input', () => {
      const q = sinTildes(buscar.value.trim());
      let alguna = false;
      for (const b of botones) { const si = !q || b._nombre.includes(q); b.hidden = !si; if (si) alguna = true; }
      for (const s of cuerpo.children) s.hidden = ![...s.querySelectorAll('.pr-gal-forma')].some((b) => !b.hidden);
      vacio.hidden = alguna;
    });
    return new Promise((resolver) => {
      let resultado = null;
      cuerpo.addEventListener('click', (ev) => {
        const b = ev.target.closest('.pr-gal-forma');
        if (!b) return;
        const f = b._forma;
        resultado = { prst: f.prst, nombre: idioma === 'en' ? f.en : f.es, flecha: !!f.flecha, ancho: f.ancho !== undefined ? f.ancho : 3, alto: f.alto !== undefined ? f.alto : 3 };
        dlg.close();
      });
      cerrar.addEventListener('click', () => dlg.close());
      dlg.addEventListener('close', () => { dlg.remove(); resolver(resultado); });
      document.body.append(dlg);
      dlg.showModal();
      buscar.focus();
    });
  }
  globalThis.GaleriaFormas = { abrir, categorias: () => CATEGORIAS.map(([clave, formas]) => [clave, formas.filter(hay).map((f) => f.prst)]) };
})();

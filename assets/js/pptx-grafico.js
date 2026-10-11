/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Gráficos de PowerPoint (c:chartSpace, ppt/charts/chartN.xml): los lee, les
// calcula la escala de los ejes y el diseño automático como PowerPoint, y los
// dibuja en un canvas. Etapa 7 del motor propio.
//
// De dónde salen las reglas (investigacion-motor-powerpoint\06-graficos.md,
// fuera del repo): PowerPoint deja vacío el diseño (c:layout) y lo calcula al
// abrir; ninguna norma lo describe. Todo lo de acá se midió por COM contra
// PowerPoint y Excel de verdad (comparar-pptx/powerpoint.ps1 anota la caja de
// cada parte del gráfico): la escala automática del eje (720 de 720 casos), el
// recuadro del contenido, los espacios entre título, leyenda, etiquetas y
// trazado, y el ancho de las barras. Lo que todavía es aproximado está
// marcado "[calibrar]". Todas las medidas van en puntos, relativas a la
// esquina de la forma del gráfico.
//
// El núcleo (escala, diseño y dibujo) trabaja sobre un modelo ya resuelto,
// sin XML: así lo puede usar también un gráfico de Excel.
//
// Uso (desde pptx-modelo.js y pptx-dibujo.js):
//   const g = PptxGrafico.leer(raizChartSpace, r)  // r: resolvedores del modelo (ver leer)
//   const geo = PptxGrafico.dibujar(ctx2d, g, anchoPx, altoPx, pxPorPt, pintar)  (pintar: ver dibujar)
//   PptxGrafico.letras(g) → [{ fuente, negrita, cursiva }] para cargar las fuentes antes.

const PptxGrafico = (() => {
  // --- XML -------------------------------------------------------------------------
  const hijos = (el, nombre) => el ? [...el.children].filter((h) => h.localName === nombre) : [];
  const hijo = (el, nombre) => { if (!el) return null; for (const h of el.children) if (h.localName === nombre) return h; return null; };
  const ruta = (el, ...nombres) => nombres.reduce((e, n) => hijo(e, n), el);
  const val = (el, def = null) => (el && el.hasAttribute('val') ? el.getAttribute('val') : def);
  const num = (el, def = null) => { const v = val(el); const n = v == null ? NaN : parseFloat(v); return Number.isFinite(n) ? n : def; };
  const si = (el, def = false) => { if (!el) return def; const v = val(el, '1'); return v === '1' || v === 'true'; };

  // --- Lectura -------------------------------------------------------------------------
  const TIPOS = {
    barChart: 'barra', bar3DChart: 'barra', lineChart: 'linea', line3DChart: 'linea', areaChart: 'area', area3DChart: 'area',
    pieChart: 'circular', pie3DChart: 'circular', ofPieChart: 'circular', doughnutChart: 'anillo', scatterChart: 'dispersion',
    radarChart: 'radar', bubbleChart: 'burbuja', stockChart: 'linea',
  };
  // Lo que cambia mucho el dibujo y todavía no sabe hacer: con eso, el
  // gráfico va al respaldo en vez de salir distinto sin aviso. Los 3D sí se
  // dibujan, planos: la librería de respaldo también los dibuja planos, y peor.
  const NO_SOPORTADOS = ['ofPieChart', 'upDownBars'];
  // Lo que PowerPoint escribe cuando el archivo no trae el texto: en el idioma
  // de quien mira (como Office, que lo pone en el de su instalación).
  const TEXTOS = /^es/i.test((typeof navigator !== 'undefined' && navigator.language) || '') ? {
    serie: 'Serie', titulo: 'Título del gráfico', tituloEje: 'Título del eje',
    tendencia: { linear: 'Lineal', exp: 'Exponencial', log: 'Logarítmica', poly: 'Polinómica', power: 'Potencial', movingAvg: 'Media móvil' },
    unidades: { hundreds: 'Cientos', thousands: 'Millares', tenThousands: 'Decenas de millar', hundredThousands: 'Cientos de miles', millions: 'Millones', tenMillions: 'Decenas de millones', hundredMillions: 'Cientos de millones', billions: 'Miles de millones', trillions: 'Billones' },
  } : {
    serie: 'Series', titulo: 'Chart Title', tituloEje: 'Axis Title',
    tendencia: { linear: 'Linear', exp: 'Expon.', log: 'Log.', poly: 'Poly.', power: 'Power', movingAvg: 'Mov. Avg.' },
    unidades: { hundreds: 'Hundreds', thousands: 'Thousands', tenThousands: 'Ten Thousands', hundredThousands: 'Hundred Thousands', millions: 'Millions', tenMillions: 'Ten Millions', hundredMillions: 'Hundred Millions', billions: 'Billions', trillions: 'Trillions' },
  };
  const DIVISORES = { hundreds: 1e2, thousands: 1e3, tenThousands: 1e4, hundredThousands: 1e5, millions: 1e6, tenMillions: 1e7, hundredMillions: 1e8, billions: 1e9, trillions: 1e12 };
  // Los que sabe dibujar; con otro, el gráfico queda sin soporte (lo dibuja
  // la librería de respaldo).
  const SOPORTADOS = new Set(['barra', 'linea', 'area', 'circular', 'anillo', 'dispersion', 'radar', 'burbuja']);

  // raiz: <c:chartSpace>. r (de pptx-modelo.js):
  //   r.relleno(spPr) → relleno | null (sin relleno) | undefined (no dice)
  //   r.linea(spPr)   → línea | null | undefined | { sinColor, ancho } (a:ln sin relleno)
  //   r.color(el)     → { r, g, b, a } | null
  //   r.fuente(nombre) → nombre resuelto (+mn-lt → la del tema)
  //   r.acento(i)     → color del acento i (1..6) del tema
  // Devuelve el modelo, o tira un Error con el motivo si no lo puede dibujar.
  function leer(raiz, r) {
    const chart = hijo(raiz, 'chart');
    const plot = hijo(chart, 'plotArea');
    if (!plot) throw new Error('gráfico sin área de trazado');
    for (const n of NO_SOPORTADOS) if (plot.getElementsByTagNameNS('*', n).length) throw new Error('gráfico con ' + n);
    // Una línea que dice el ancho pero no el color (python-pptx) lleva el
    // color automático: aquí cuenta como "no dice".
    const linea = (spPr) => { const l = r.linea(spPr); return l && l.sinColor ? undefined : l; };
    const letraBase = letraDe(hijo(raiz, 'txPr'), r, { tamano: 10, negrita: false, cursiva: false, fuente: r.fuente('+mn-lt') || 'Calibri', color: { r: 89, g: 89, b: 89, a: 1 }, kern: 0 });
    const ejes = new Map();
    for (const el of plot.children) {
      if (!['catAx', 'valAx', 'dateAx', 'serAx'].includes(el.localName)) continue;
      const id = val(hijo(el, 'axId'));
      const escala = hijo(el, 'scaling');
      const fmt = hijo(el, 'numFmt');
      const ln = ruta(el, 'spPr', 'ln');
      ejes.set(id, {
        id, clase: el.localName, pos: val(hijo(el, 'axPos'), 'l'), borrado: si(hijo(el, 'delete')),
        invertido: val(hijo(escala, 'orientation')) === 'maxMin', log: num(hijo(escala, 'logBase')),
        minFijo: num(hijo(escala, 'min')), maxFijo: num(hijo(escala, 'max')), unidadFija: num(hijo(el, 'majorUnit')),
        formato: fmt ? { codigo: fmt.getAttribute('formatCode') || 'General', enlazado: ['1', 'true'].includes(fmt.getAttribute('sourceLinked')) } : { codigo: 'General', enlazado: true },
        cuadricula: hijo(el, 'majorGridlines') ? conDefecto(linea(hijo(hijo(el, 'majorGridlines'), 'spPr')), lineaFina({ r: 217, g: 217, b: 217, a: 1 })) : null,
        linea: ln ? linea(hijo(el, 'spPr')) ?? null : null,
        marca: val(hijo(el, 'majorTickMark'), 'cross'), posEtiquetas: val(hijo(el, 'tickLblPos'), 'nextTo'),
        letra: letraDe(hijo(el, 'txPr'), r, letraBase), giro: giroDe(hijo(el, 'txPr')),
        cruza: hijo(el, 'crossesAt') ? num(hijo(el, 'crossesAt')) : val(hijo(el, 'crosses'), 'autoZero'),
        cruzaCon: val(hijo(el, 'crossAx')), entre: val(hijo(el, 'crossBetween'), 'between'),
        salto: num(hijo(el, 'tickLblSkip')),
        titulo: tituloDeEje(hijo(el, 'title'), r, letraBase, ['l', 'r'].includes(val(hijo(el, 'axPos'), 'l')), linea),
        unidades: unidadesDe(hijo(el, 'dispUnits'), r, letraBase, ['l', 'r'].includes(val(hijo(el, 'axPos'), 'l')), linea),
      });
    }

    const grupos = [];
    let indiceSerie = 0;
    for (const el of plot.children) {
      const tipo = TIPOS[el.localName];
      if (!tipo) continue;
      if (!SOPORTADOS.has(tipo)) throw new Error('gráfico de tipo ' + el.localName);
      const g = {
        tipo, es3D: /3DChart$/.test(el.localName), dir: val(hijo(el, 'barDir'), 'col'), agrupacion: val(hijo(el, 'grouping'), tipo === 'barra' ? 'clustered' : 'standard'),
        separacion: num(hijo(el, 'gapWidth'), 150) / 100, solapamiento: num(hijo(el, 'overlap'), 0) / 100,
        hueco: num(hijo(el, 'holeSize'), 10) / 100, anguloInicial: num(hijo(el, 'firstSliceAng'), 0),
        variarColores: si(hijo(el, 'varyColors')), estiloDispersion: val(hijo(el, 'scatterStyle'), 'marker'), estiloRadar: val(hijo(el, 'radarStyle'), 'marker'),
        separacionProf: num(hijo(el, 'gapDepth'), 150) / 100,
        escalaBurbuja: num(hijo(el, 'bubbleScale'), 100), burbujaAncho: val(hijo(el, 'sizeRepresents'), 'area') === 'w', burbujasNegativas: si(hijo(el, 'showNegBubbles')),
        ejes: hijos(el, 'axId').map((a) => val(a)),
        etiquetas: etiquetasDe(hijo(el, 'dLbls'), r, letraBase, null),
        maxMin: hijo(el, 'hiLowLines') ? conDefecto(linea(hijo(hijo(el, 'hiLowLines'), 'spPr')), lineaFina({ r: 0, g: 0, b: 0, a: 1 })) : null,
        series: [],
      };
      for (const s of hijos(el, 'ser')) {
        const idx = num(hijo(s, 'idx'), indiceSerie);
        indiceSerie++;
        const spPr = hijo(s, 'spPr');
        const datos = (nombre) => valores(hijo(s, nombre));
        const cat = datos('cat') || datos('xVal');
        const serie = {
          idx, orden: num(hijo(s, 'order'), idx), nombre: textoDe(hijo(s, 'tx')) || TEXTOS.serie + (idx + 1),
          categorias: cat ? cat.valores : null, formatoCategorias: cat ? cat.formato : null,
          valores: (datos('val') || datos('yVal') || { valores: [] }).valores,
          formato: (datos('val') || datos('yVal') || {}).formato || 'General',
          tamanos: (datos('bubbleSize') || {}).valores || null,
          relleno: r.relleno(spPr), linea: r.linea(spPr), // sinColor: lo completa automatico
          suave: si(hijo(s, 'smooth')), explosion: num(hijo(s, 'explosion'), 0) / 100,
          // Sin c:marker, la norma dice automático: en la dispersión (salvo
          // estilo de solo línea) y en las líneas con <c:marker val="1"/> en el grupo.
          marcador: hijo(s, 'marker') ? marcadorDe(hijo(s, 'marker'), r)
            : (tipo === 'dispersion' && !['line', 'smooth'].includes(g.estiloDispersion)) || (tipo === 'linea' && si(hijo(el, 'marker'))) || (tipo === 'radar' && g.estiloRadar === 'marker') ? { simbolo: 'auto', tamano: 5, relleno: undefined, linea: undefined } : null,
          puntos: new Map(),
          etiquetas: etiquetasDe(hijo(s, 'dLbls'), r, letraBase, g.etiquetas),
          tendencias: hijos(s, 'trendline').map((t) => ({
            tipo: val(hijo(t, 'trendlineType'), 'linear'), orden: num(hijo(t, 'order'), 2), periodo: num(hijo(t, 'period'), 2),
            nombre: (hijo(t, 'name') || {}).textContent || null, linea: linea(hijo(t, 'spPr')),
            adelante: num(hijo(t, 'forward'), 0), atras: num(hijo(t, 'backward'), 0), corte: num(hijo(t, 'intercept')),
          })),
          errores: hijos(s, 'errBars').map((e) => ({
            dir: val(hijo(e, 'errDir'), 'y'), tipo: val(hijo(e, 'errBarType'), 'both'), valor: val(hijo(e, 'errValType'), 'fixedVal'),
            cantidad: num(hijo(e, 'val'), 1), sinTope: si(hijo(e, 'noEndCap')), linea: linea(hijo(e, 'spPr')),
            mas: (valores(hijo(e, 'plus')) || {}).valores || null, menos: (valores(hijo(e, 'minus')) || {}).valores || null,
          })),
        };
        for (const p of hijos(s, 'dPt')) {
          const i = num(hijo(p, 'idx'), 0);
          serie.puntos.set(i, { relleno: r.relleno(hijo(p, 'spPr')), linea: linea(hijo(p, 'spPr')), explosion: hijo(p, 'explosion') ? num(hijo(p, 'explosion'), 0) / 100 : null, marcador: hijo(p, 'marker') ? marcadorDe(hijo(p, 'marker'), r) : null });
        }
        g.series.push(serie);
      }
      g.series.sort((a, b) => a.orden - b.orden);
      grupos.push(g);
    }
    if (!grupos.length || grupos.some((g) => !g.series.length)) throw new Error('gráfico sin series');
    for (const g of grupos) if (g.tipo !== 'circular' && g.tipo !== 'anillo' && g.ejes.filter((id) => ejes.has(id)).length < 2) throw new Error('gráfico sin ejes');
    if (grupos.some((g) => g.tipo === 'circular' || g.tipo === 'anillo') && grupos.length > 1) throw new Error('gráfico circular combinado');

    // Colores que el archivo no trae: los acentos del tema en ciclo (cs:colorStyle
    // meth="cycle"; ver 06-graficos.md, 6.6). [calibrar] las variaciones de la 7.ª en adelante.
    const automatico = (i) => {
      const base = r.acento((i % 6) + 1);
      const vuelta = Math.floor(i / 6);
      if (!vuelta) return base;
      const k = [1, 0.6, 1.4, 0.8, 1.2][vuelta % 5];
      const m = (c) => Math.max(0, Math.min(255, Math.round(k < 1 ? c * k : c + (255 - c) * (k - 1))));
      return { r: m(base.r), g: m(base.g), b: m(base.b), a: base.a };
    };
    for (const g of grupos) {
      for (const s of g.series) {
        const c = automatico(s.idx);
        const deLinea = g.tipo === 'linea' || g.tipo === 'dispersion' || (g.tipo === 'radar' && g.estiloRadar !== 'filled');
        if (s.relleno === undefined) s.relleno = deLinea ? null : { tipo: 'solido', color: c };
        if (s.linea === undefined || (s.linea && s.linea.sinColor)) {
          const ancho = s.linea && s.linea.ancho;
          s.linea = deLinea ? lineaFina(c, ancho || 28575) : g.tipo === 'circular' || g.tipo === 'anillo' ? lineaFina({ r: 255, g: 255, b: 255, a: 1 }, ancho || 19050) : null;
        }
        s.colorAuto = c;
        // Tendencia sin línea propia: punteada del color de la serie; barras de error: gris.
        for (const t of s.tendencias || []) if (t.linea === undefined) t.linea = { ...lineaFina(c, 19050), guion: 'sysDot', tapa: 'rnd' };
        for (const e of s.errores || []) if (e.linea === undefined) e.linea = lineaFina({ r: 89, g: 89, b: 89, a: 1 });
      }
    }

    const tituloEl = hijo(chart, 'title');
    const borrado = si(hijo(chart, 'autoTitleDeleted'));
    let titulo = null;
    if (tituloEl) {
      const letra = letraDe(ruta(tituloEl, 'tx', 'rich') || hijo(tituloEl, 'txPr'), r, { ...letraBase, tamano: letraBase.tamano * 1.4 });
      let texto = textoRico(ruta(tituloEl, 'tx', 'rich')) || textoDe(hijo(tituloEl, 'tx'));
      // Sin texto propio: el nombre de la serie si hay una sola (como PowerPoint).
      if (!texto) { const todas = grupos.flatMap((g) => g.series); texto = todas.length === 1 ? todas[0].nombre : TEXTOS.titulo; }
      titulo = { texto, letra, diseno: disenoDe(tituloEl), relleno: r.relleno(hijo(tituloEl, 'spPr')) ?? null, linea: linea(hijo(tituloEl, 'spPr')) ?? null };
    }
    void borrado; // [calibrar] autoTitleDeleted="0" sin c:title y una sola serie: ¿PowerPoint muestra su nombre?

    const leyendaEl = hijo(chart, 'legend');
    const leyenda = leyendaEl ? {
      pos: val(hijo(leyendaEl, 'legendPos'), 'r'), diseno: disenoDe(leyendaEl), letra: letraDe(hijo(leyendaEl, 'txPr'), r, letraBase),
      borradas: new Set(hijos(leyendaEl, 'legendEntry').filter((e) => si(hijo(e, 'delete'))).map((e) => num(hijo(e, 'idx'), -1))),
      relleno: r.relleno(hijo(leyendaEl, 'spPr')) ?? null, linea: linea(hijo(leyendaEl, 'spPr')) ?? null,
    } : null;

    return {
      grupos, ejes, titulo, leyenda, letraBase, automatico,
      // <a:noFill/> (null) es transparente; sin decir nada (undefined), blanco.
      fondo: { relleno: conDefecto(r.relleno(hijo(raiz, 'spPr')), { tipo: 'solido', color: { r: 255, g: 255, b: 255, a: 1 } }), linea: linea(hijo(raiz, 'spPr')) ?? null },
      fondoTrazado: { relleno: r.relleno(hijo(plot, 'spPr')) ?? null, linea: linea(hijo(plot, 'spPr')) ?? null },
      vacios: val(hijo(chart, 'dispBlanksAs'), 'gap'),
      disenoTrazado: disenoDe(plot),
      tabla: tablaDe(hijo(plot, 'dTable'), r, letraBase, linea),
      rotX: num(ruta(chart, 'view3D', 'rotX'), 30),
      rotY: num(ruta(chart, 'view3D', 'rotY'), 20),
      perspectiva: num(ruta(chart, 'view3D', 'perspective'), 30),
      profundidad: num(ruta(chart, 'view3D', 'depthPercent'), 100),
      ejesRectos: si(ruta(chart, 'view3D', 'rAngAx'), true),
    };
  }

  const conDefecto = (v, def) => (v === undefined ? def : v);
  const lineaFina = (color, ancho = 9525) => ({ ancho, relleno: { tipo: 'solido', color }, tapa: 'flat', compuesto: 'sng', guion: 'solid', union: 'round', inicio: null, fin: null });

  // c:layout/c:manualLayout (una pieza movida o agrandada a mano): fracciones
  // del alto y el ancho de la forma; en modo "edge" son la posición, en
  // "factor" un corrimiento desde donde iría sola. layoutTarget del trazado:
  // "inner" es el área de trazado sin las etiquetas de los ejes.
  function disenoDe(el) {
    const m = ruta(el, 'layout', 'manualLayout');
    if (!m || !m.children.length) return null;
    return {
      objetivo: val(hijo(m, 'layoutTarget'), 'outer'), xModo: val(hijo(m, 'xMode'), 'factor'), yModo: val(hijo(m, 'yMode'), 'factor'),
      wModo: val(hijo(m, 'wMode'), 'factor'), hModo: val(hijo(m, 'hMode'), 'factor'),
      x: num(hijo(m, 'x')), y: num(hijo(m, 'y')), w: num(hijo(m, 'w')), h: num(hijo(m, 'h')),
    };
  }
  function aplicarDiseno(caja, d, W, H) {
    if (!d) return caja;
    const c = { ...caja };
    if (d.x != null) c.x = d.xModo === 'edge' ? d.x * W : caja.x + d.x * W;
    if (d.y != null) c.y = d.yModo === 'edge' ? d.y * H : caja.y + d.y * H;
    if (d.w != null) c.w = d.wModo === 'edge' ? d.w * W - c.x : d.w * W;
    if (d.h != null) c.h = d.hModo === 'edge' ? d.h * H - c.y : d.h * H;
    return c;
  }

  // Unidades de visualización (c:dispUnits): los números del eje divididos
  // (miles, millones…) y, con c:dispUnitsLbl, su rótulo, que va como un título de eje.
  function unidadesDe(el, r, letraBase, dePie, linea) {
    if (!el) return null;
    const nombre = val(hijo(el, 'builtInUnit'));
    const divisor = nombre ? DIVISORES[nombre] : num(hijo(el, 'custUnit'));
    if (!divisor) return null;
    const lbl = hijo(el, 'dispUnitsLbl');
    let rotulo = null;
    if (lbl) {
      rotulo = tituloDeEje(lbl, r, { ...letraBase, tamano: letraBase.tamano * 1197 / 1330 }, dePie, linea);
      if (rotulo.texto === TEXTOS.tituloEje) rotulo.texto = (nombre && TEXTOS.unidades[nombre]) || String(divisor);
    }
    return { divisor, rotulo };
  }

  // Tabla de datos (c:dTable): bordes, muestras y su letra.
  function tablaDe(el, r, letraBase, linea) {
    if (!el) return null;
    return {
      horizontales: si(hijo(el, 'showHorzBorder')), verticales: si(hijo(el, 'showVertBorder')),
      contorno: si(hijo(el, 'showOutline')), muestras: si(hijo(el, 'showKeys')),
      letra: letraDe(hijo(el, 'txPr'), r, letraBase),
      linea: conDefecto(linea(hijo(el, 'spPr')), lineaFina({ r: 217, g: 217, b: 217, a: 1 })),
    };
  }

  // Título de un eje: texto rico (o el de por omisión), letra de 13,3 pt si no
  // dice (10/9 de la del gráfico) y, en un eje de pie, girado -90°.
  function tituloDeEje(el, r, letraBase, dePie, linea) {
    if (!el) return null;
    const rich = ruta(el, 'tx', 'rich');
    const letra = letraDe(rich || hijo(el, 'txPr'), r, { ...letraBase, tamano: letraBase.tamano * 1330 / 1197 });
    const giro = giroDe(rich || hijo(el, 'txPr'));
    return {
      texto: textoRico(rich) || textoDe(hijo(el, 'tx')) || TEXTOS.tituloEje, letra,
      giro: giro != null ? giro : dePie ? -90 : 0,
      relleno: r.relleno(hijo(el, 'spPr')) ?? null, linea: linea(hijo(el, 'spPr')) ?? null,
    };
  }

  // Valores de un c:cat / c:val / c:xVal… desde sus caches (los datos del libro
  // incrustado solo hacen falta para editar). Un punto que falta es un hueco.
  function valores(el) {
    if (!el) return null;
    const ref = el.firstElementChild;
    if (!ref) return null;
    const cache = hijo(ref, 'numCache') || hijo(ref, 'strCache') || (['numLit', 'strLit'].includes(ref.localName) ? ref : null);
    if (ref.localName === 'multiLvlStrRef') {
      // Categorías de varios niveles: el nivel de más adentro. [calibrar]
      const lvl = ruta(ref, 'multiLvlStrCache', 'lvl');
      return lvl ? leerPuntos(lvl, ref.parentElement, false) : null;
    }
    if (!cache) return null;
    const numerico = cache.localName === 'numCache' || cache.localName === 'numLit';
    return leerPuntos(cache, null, numerico);
  }
  function leerPuntos(cache, _padre, numerico) {
    const n = num(hijo(cache, 'ptCount'), 0);
    const lista = new Array(n).fill(null);
    for (const p of hijos(cache, 'pt')) {
      const i = parseInt(p.getAttribute('idx'), 10);
      const t = (hijo(p, 'v') || { textContent: '' }).textContent;
      // #N/A y otros errores de Excel: hueco. Un idx fuera de ptCount no cuenta
      // (un archivo con idx enorme colgaría la pestaña).
      if (i >= 0 && i < n) { const x = parseFloat(t); lista[i] = numerico ? (Number.isFinite(x) ? x : null) : t; }
    }
    const fc = hijo(cache, 'formatCode');
    return { valores: lista, formato: fc ? fc.textContent : 'General', numerico };
  }
  function textoDe(tx) {
    if (!tx) return '';
    const v = tx.getElementsByTagNameNS('*', 'v')[0];
    if (v) return v.textContent;
    return textoRico(hijo(tx, 'rich'));
  }
  function textoRico(rich) {
    if (!rich) return '';
    return hijos(rich, 'p').map((p) => [...p.getElementsByTagNameNS('*', 't')].map((t) => t.textContent).join('')).join('\n').trim();
  }
  // La letra de un txPr / rich (defRPr del primer párrafo, o el rPr del primer trozo).
  function letraDe(txPr, r, base) {
    if (!txPr) return { ...base };
    const p = hijo(txPr, 'p');
    const def = ruta(p, 'pPr', 'defRPr');
    const rpr = (p && [...p.getElementsByTagNameNS('*', 'rPr')][0]) || null;
    const letra = { ...base };
    for (const e of [def, rpr]) {
      if (!e) continue;
      if (e.hasAttribute('sz')) letra.tamano = parseInt(e.getAttribute('sz'), 10) / 100;
      if (e.hasAttribute('b')) letra.negrita = ['1', 'true'].includes(e.getAttribute('b'));
      if (e.hasAttribute('i')) letra.cursiva = ['1', 'true'].includes(e.getAttribute('i'));
      if (e.hasAttribute('kern')) letra.kern = parseInt(e.getAttribute('kern'), 10) / 100;
      const relleno = r.relleno(e);
      if (relleno && relleno.tipo === 'solido') letra.color = relleno.color;
      const latin = hijo(e, 'latin');
      if (latin && latin.getAttribute('typeface')) letra.fuente = r.fuente(latin.getAttribute('typeface')) || letra.fuente;
    }
    return letra;
  }
  function giroDe(txPr) {
    const b = hijo(txPr, 'bodyPr');
    if (!b || !b.hasAttribute('rot')) return null;
    const g = parseInt(b.getAttribute('rot'), 10);
    return g === -60000000 ? null : g / 60000; // -60000000: automático
  }
  function marcadorDe(el, r) {
    if (!el) return null;
    const simbolo = val(hijo(el, 'symbol'), 'auto');
    const l = r.linea(hijo(el, 'spPr'));
    return { simbolo, tamano: num(hijo(el, 'size'), 5), relleno: r.relleno(hijo(el, 'spPr')), linea: l && l.sinColor ? undefined : l };
  }
  function etiquetasDe(el, r, letraBase, herencia) {
    if (!el) return herencia;
    const h = herencia || {};
    const mostrar = (n) => (hijo(el, n) ? si(hijo(el, n)) : !!h[n]);
    const fmt = hijo(el, 'numFmt');
    return {
      showVal: mostrar('showVal'), showCatName: mostrar('showCatName'), showSerName: mostrar('showSerName'),
      showPercent: mostrar('showPercent'), showLegendKey: mostrar('showLegendKey'),
      pos: val(hijo(el, 'dLblPos'), h.pos || null), separador: (hijo(el, 'separator') || {}).textContent || h.separador || null,
      formato: fmt ? fmt.getAttribute('formatCode') : h.formato || null,
      letra: letraDe(hijo(el, 'txPr'), r, h.letra || letraBase),
      borrar: si(hijo(el, 'delete')),
    };
  }

  // --- Escala del eje de valores (medida; 06-graficos.md, 6.3) ------------------------------
  // a, b: mínimo y máximo de los datos; n: divisiones que caben. Devuelve
  // { min, max, unidad }.
  function escala(a, b, n, fijo = {}) {
    if (fijo.exacto && Number.isFinite(a) && Number.isFinite(b) && b > a) {
      // Rango ya agrandado (burbujas): solo redondear a la unidad.
      const lo = fijo.min != null ? fijo.min : a, hi = fijo.max != null ? fijo.max : b;
      const unidad = fijo.unidad || agradable((hi - lo) / Math.max(1, n));
      const min = fijo.min != null ? fijo.min : parseFloat((Math.floor(lo / unidad + 1e-9) * unidad).toFixed(10));
      const max = fijo.max != null ? fijo.max : parseFloat((Math.ceil(hi / unidad - 1e-9) * unidad).toFixed(10));
      return { min, max: max > min ? max : min + unidad, unidad };
    }
    if (fijo.log) {
      // Logarítmica: de una potencia de la base a otra, una división por potencia.
      const B = fijo.log, lg = (x) => Math.log(x) / Math.log(B);
      const pa = a > 0 && Number.isFinite(a) ? a : 1, pb = b > 0 && Number.isFinite(b) ? b : pa;
      const min = fijo.min > 0 ? fijo.min : Math.pow(B, Math.floor(lg(pa) + 1e-9));
      let max = fijo.max > 0 ? fijo.max : Math.pow(B, Math.ceil(lg(pb) - 1e-9));
      if (!(max > min)) max = min * B;
      return { min: parseFloat(min.toPrecision(12)), max: parseFloat(max.toPrecision(12)), unidad: B, log: B };
    }
    if (!Number.isFinite(a) || !Number.isFinite(b)) { a = 0; b = 1; }
    if (fijo.min != null) a = Math.min(a, fijo.min);
    if (fijo.max != null) b = Math.max(b, fijo.max);
    if (a === b) { if (a === 0) b = 1; else if (a > 0) a = 0; else b = 0; }
    const ceroAbajo = fijo.min == null && a >= 0 && a < (5 / 6) * b;
    const ceroArriba = fijo.max == null && b <= 0 && b > (5 / 6) * a;
    let lo, hi;
    // fijo.sinAire: el radar no deja el 5 % de aire después del dato (medido).
    const aire = fijo.sinAire ? 0 : 1 / 20;
    if (ceroAbajo) { lo = 0; hi = b + b * aire; }
    else if (ceroArriba) { hi = 0; lo = a + a * aire; }
    else if (a >= 0) { lo = a - (b - a) / 2; hi = b + (b - lo) / 20; }
    else if (b <= 0) { hi = b + (b - a) / 2; lo = a - (hi - a) / 20; }
    else { lo = a - (b - a) / 20; hi = b + (b - a) / 20; }
    if (fijo.min != null) lo = fijo.min;
    if (fijo.max != null) hi = fijo.max;
    const unidad = fijo.unidad || agradable((hi - lo) / Math.max(1, n));
    const r = (x) => Math.round(x / unidad) * unidad;
    const min = fijo.min != null ? fijo.min : ceroAbajo ? 0 : r(Math.floor(lo / unidad + 1e-9) * unidad);
    const max = fijo.max != null ? fijo.max : ceroArriba ? 0 : r(Math.ceil(hi / unidad - 1e-9) * unidad);
    return { min, max: max > min ? max : min + unidad, unidad };
  }
  // Diámetro de la burbuja mayor como fracción del lado menor del trazado,
  // según bubbleScale (medido: 50 → 0,139, 100 → 0,2475, 200 → 0,40). [calibrar] el resto.
  function fraccionBurbuja(escala) {
    const puntos = [[0, 0], [50, 0.139], [100, 0.2475], [200, 0.4], [300, 0.52]];
    const e = Math.max(0, Math.min(300, escala));
    for (let i = 1; i < puntos.length; i++) if (e <= puntos[i][0]) { const [x0, y0] = puntos[i - 1], [x1, y1] = puntos[i]; return y0 + (y1 - y0) * (e - x0) / (x1 - x0); }
    return 0.52;
  }
  // Valores de las marcas de una escala, y dónde cae un valor (0 a 1).
  function marcas(esc) {
    const lista = [];
    if (esc.log) for (let k = 0, v = esc.min; v <= esc.max * (1 + 1e-9) && k < 200; v = esc.min * Math.pow(esc.log, ++k)) lista.push(parseFloat(v.toPrecision(12)));
    else for (let k = 0, v = esc.min; v <= esc.max + esc.unidad * 1e-6 && k < 200; v = esc.min + (++k) * esc.unidad) lista.push(parseFloat(v.toFixed(12)));
    return lista;
  }
  function fraccion(esc, v) {
    if (esc.log) return (Math.log(Math.max(v, esc.min * 1e-6)) - Math.log(esc.min)) / (Math.log(esc.max) - Math.log(esc.min));
    return (v - esc.min) / (esc.max - esc.min);
  }
  function agradable(x) {
    if (!(x > 0)) return 1;
    const e = Math.floor(Math.log10(x));
    for (const m of [1, 2, 5, 10]) { const u = m * Math.pow(10, e); if (u >= x * (1 - 1e-12)) return parseFloat(u.toPrecision(12)); }
    return Math.pow(10, e + 1);
  }
  // Cuántas divisiones caben (medido en Excel; 6.3): de pie, cada una cuesta
  // una altura de línea de las etiquetas; acostado, unas 3,8 veces la letra.
  function divisiones(largo, letra, lineaPt, acostado) {
    // De pie, PowerPoint da una división más que Excel cerca del umbral
    // (144,9 pt con 11,97: Excel 8, PowerPoint 9): -1,45 en vez del -1,55 de Excel.
    const n = acostado ? Math.floor((largo - 5.1 * letra) / (3.81 * letra)) + 2 : Math.floor((largo + 8) / lineaPt - 1.45);
    return Math.max(1, Math.min(10, n));
  }

  // --- Formato de números ------------------------------------------------------------------------
  // "General" y los códigos comunes de Excel (0, #,##0, 0.0%, "texto", secciones
  // pos;neg;cero), con los separadores del idioma del navegador, como los
  // muestra PowerPoint en la PC de quien mira.
  // Decimales de un número tal como se escribe (2,5 → 1; 0,25 → 2).
  function decimalesDe(x) {
    const t = String(parseFloat(Number(x).toPrecision(12)));
    const m = t.match(/e-(\d+)$/);
    if (m) return parseInt(m[1], 10) + ((t.split('e')[0].split('.')[1] || '').length);
    return (t.split('.')[1] || '').length;
  }
  let separadores = null;
  function sep() {
    if (!separadores) {
      const partes = new Intl.NumberFormat(undefined, { useGrouping: true }).formatToParts(12345.6);
      separadores = { decimal: (partes.find((p) => p.type === 'decimal') || { value: '.' }).value, miles: (partes.find((p) => p.type === 'group') || { value: ',' }).value };
      if (/\s/.test(separadores.miles)) separadores.miles = '.';
    }
    return separadores;
  }
  function formatear(v, codigo, decimalesGeneral = null) {
    if (v == null || !Number.isFinite(v)) return '';
    const s = sep();
    codigo = codigo || 'General';
    if (/^general$/i.test(codigo.trim())) return general(v, decimalesGeneral, s);
    // Secciones pos;neg;cero, o con condiciones ([>=1000]…): la primera que se cumple.
    const secciones = codigo.split(';');
    let fmt = secciones[0], valor = v, conSigno = secciones.length < 2;
    const condicion = (sec) => { const m = sec.match(/\[(<=|>=|<>|<|>|=)(-?[\d.]+)\]/); return m ? { op: m[1], n: parseFloat(m[2]) } : null; };
    if (secciones.some(condicion)) {
      const cumple = (c) => ({ '<': v < c.n, '>': v > c.n, '=': v === c.n, '<=': v <= c.n, '>=': v >= c.n, '<>': v !== c.n })[c.op];
      fmt = secciones.find((x) => { const c = condicion(x); return c ? cumple(c) : true; }) || secciones[secciones.length - 1];
      conSigno = true;
    } else if (v < 0 && secciones[1] != null) { fmt = secciones[1]; valor = -v; }
    else if (v === 0 && secciones[2] != null) fmt = secciones[2];
    fmt = fmt.replace(/\[[^\]]*\]/g, '');
    const literales = [];
    fmt = fmt.replace(/"([^"]*)"|\\(.)/g, (_, a, b) => { literales.push(a != null ? a : b); return '\u0001'; });
    const conLiterales = (t) => { let k = 0; return t.replace(/\u0001/g, () => literales[k++]).replace(/[_*]./g, ' ').trim(); };
    // Fechas y horas (número de serie de Excel, sistema 1900).
    if (!/[#0?]/.test(fmt) && /[dmyhs]/i.test(fmt)) return conLiterales(fecha(v, fmt));
    // Lo que no entiende (texto @, fracciones, letras sueltas): General, nunca el código tal cual.
    if (/[@\/]/.test(fmt) || /[a-df-zA-DF-Z]/.test(fmt.replace(/E[+-]/g, ''))) return general(v, null, s);
    if (/%/.test(fmt)) valor *= 100;
    // Científico: 0.00E+00.
    const cientifico = fmt.match(/([#0?.,]+)E([+-])([0?#]+)/i);
    if (cientifico) {
      const dec = ((cientifico[1].split('.')[1] || '').match(/[0#?]/g) || []).length;
      let [m, e] = Math.abs(valor).toExponential(dec).split('e');
      e = parseInt(e, 10);
      const exp = (e < 0 ? '-' : cientifico[2] === '+' ? '+' : '') + String(Math.abs(e)).padStart(cientifico[3].length, '0');
      return conLiterales(fmt.replace(cientifico[0], (valor < 0 && conSigno ? '-' : '') + m.replace('.', s.decimal) + 'E' + exp));
    }
    const patron = /[#0?]/.test(fmt) ? fmt.match(/[#0?,]+(\.[#0?]+)?,*|\.[#0?]+,*/) : null;
    let numero = '';
    if (patron) {
      let p = patron[0];
      // Comas al final del número: cada una divide por mil (#,##0,"K").
      const escala = (p.match(/,+$/) || [''])[0].length;
      p = p.slice(0, p.length - escala);
      valor /= Math.pow(1000, escala);
      const dec = (p.split('.')[1] || '').replace(/[^0#?]/g, '');
      const minDec = (dec.match(/0/g) || []).length;
      let t = Math.abs(valor).toFixed(dec.length);
      if (dec.length > minDec) t = t.replace(/0+$/, (z) => z.slice(0, Math.max(0, z.length - (dec.length - minDec)))).replace(/\.$/, '');
      let [ent, fr] = t.split('.');
      const minEnt = ((p.split('.')[0] || '').match(/0/g) || []).length;
      if (ent === '0' && minEnt === 0) ent = '';
      ent = ent.padStart(minEnt, '0');
      if (/,/.test(p.split('.')[0] || '')) ent = ent.replace(/\B(?=(\d{3})+(?!\d))/g, s.miles);
      numero = (valor < 0 && conSigno && /[1-9]/.test(ent + (fr || '')) ? '-' : '') + ent + (fr ? s.decimal + fr : '');
      fmt = fmt.replace(patron[0], '\u0002');
    }
    return conLiterales(fmt.replace(/\u0002/, numero));
  }
  // "General": hasta 10 cifras; muy grande o muy chico, científico (1,5E+20).
  function general(v, decimales, s) {
    let t;
    const abs = Math.abs(v);
    if (decimales != null) t = v.toFixed(Math.max(0, Math.min(10, decimales)));
    else if (abs !== 0 && (abs >= 1e11 || abs < 1e-9)) t = v.toExponential(5);
    else t = String(parseFloat(v.toPrecision(10)));
    const m = t.match(/^(-?[\d.]+)e([+-])(\d+)$/);
    if (m) {
      const mant = m[1].includes('.') ? m[1].replace(/0+$/, '').replace(/\.$/, '') : m[1];
      return mant.replace('.', s.decimal) + 'E' + m[2] + m[3].padStart(2, '0');
    }
    if (t.includes('.')) t = t.replace(/0+$/, '').replace(/\.$/, '');
    if (/^-0(\.0*)?$/.test(t)) t = t.slice(1);
    return t.replace('.', s.decimal);
  }
  const MESES = Array.from({ length: 12 }, (_, i) => new Date(Date.UTC(2001, i, 15)));
  const nombreMes = (i, estilo) => new Intl.DateTimeFormat(undefined, { month: estilo, timeZone: 'UTC' }).format(MESES[i]).replace(/\.$/, '');
  const nombreDia = (d, estilo) => new Intl.DateTimeFormat(undefined, { weekday: estilo, timeZone: 'UTC' }).format(d).replace(/\.$/, '');
  function fecha(v, fmt) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v * 86400000));
    const dos = (n) => String(n).padStart(2, '0');
    const ampm = /AM\/PM|A\/P/i.test(fmt);
    let hora = d.getUTCHours();
    if (ampm) hora = hora % 12 || 12;
    // m es minuto si viene después de h o antes de s; si no, mes.
    const fichas = fmt.match(/AM\/PM|A\/P|y+|m+|d+|h+|s+|[^ymdhsAa]+|./gi) || [];
    return fichas.map((f, i) => {
      const c = f[0].toLowerCase();
      if (/^(AM\/PM|A\/P)$/i.test(f)) return d.getUTCHours() < 12 ? 'AM' : 'PM';
      if (c === 'y') return f.length <= 2 ? dos(d.getUTCFullYear() % 100) : String(d.getUTCFullYear());
      if (c === 'd') return f.length === 1 ? String(d.getUTCDate()) : f.length === 2 ? dos(d.getUTCDate()) : nombreDia(d, f.length === 3 ? 'short' : 'long');
      if (c === 'h') return f.length === 1 ? String(hora) : dos(hora);
      if (c === 's') return f.length === 1 ? String(d.getUTCSeconds()) : dos(d.getUTCSeconds());
      if (c === 'm') {
        const antes = fichas.slice(0, i).reverse().find((x) => /^[ymdhs]/i.test(x));
        const despues = fichas.slice(i + 1).find((x) => /^[ymdhs]/i.test(x));
        if (f.length <= 2 && ((antes && /^h/i.test(antes)) || (despues && /^s/i.test(despues)))) return f.length === 1 ? String(d.getUTCMinutes()) : dos(d.getUTCMinutes());
        const mes = d.getUTCMonth();
        return f.length === 1 ? String(mes + 1) : f.length === 2 ? dos(mes + 1) : f.length === 3 ? nombreMes(mes, 'short') : f.length === 5 ? nombreMes(mes, 'narrow') : nombreMes(mes, 'long');
      }
      return f;
    }).join('');
  }

  // --- Medir texto ---------------------------------------------------------------------------------
  function cssLetra(l, px) { return `${l.cursiva ? 'italic ' : ''}${l.negrita ? 'bold ' : ''}${px.toFixed(3)}px ${PptxTexto.familia(l.fuente)}`; }
  function medidor(ctx) {
    const cache = new Map();
    return {
      ancho(texto, l) {
        const clave = l.fuente + l.tamano + l.negrita + l.cursiva + '|' + texto;
        if (!cache.has(clave)) {
          ctx.save();
          ctx.font = cssLetra(l, 100);
          ctx.fontKerning = l.kern > 0 && l.tamano >= l.kern ? 'normal' : 'none';
          ctx.letterSpacing = '0px';
          cache.set(clave, ctx.measureText(texto).width / 100 * l.tamano);
          ctx.restore();
        }
        return cache.get(clave);
      },
      // Altura de línea (alto de la fuente, ascendente + descendente) en pt.
      linea(l) {
        const clave = 'L' + l.fuente + l.negrita + l.cursiva;
        if (!cache.has(clave)) {
          ctx.save(); ctx.font = cssLetra(l, 100);
          const m = ctx.measureText('Hg');
          cache.set(clave, ((m.fontBoundingBoxAscent || 80) + (m.fontBoundingBoxDescent || 20)) / 100);
          ctx.restore();
        }
        return cache.get(clave) * l.tamano;
      },
      ascendente(l) {
        const clave = 'A' + l.fuente + l.negrita + l.cursiva;
        if (!cache.has(clave)) { ctx.save(); ctx.font = cssLetra(l, 100); cache.set(clave, (ctx.measureText('Hg').fontBoundingBoxAscent || 80) / 100); ctx.restore(); }
        return cache.get(clave) * l.tamano;
      },
    };
  }
  // Corta un texto en líneas de un ancho máximo (por palabras).
  function cortar(texto, l, ancho, med) {
    const salida = [];
    for (const parrafo of String(texto).split('\n')) {
      let actual = '';
      for (const palabra of parrafo.split(/ +/)) {
        const prueba = actual ? actual + ' ' + palabra : palabra;
        if (actual && med.ancho(prueba, l) > ancho) { salida.push(actual); actual = palabra; } else actual = prueba;
      }
      salida.push(actual);
    }
    return salida;
  }

  // --- Diseño -------------------------------------------------------------------------------------------
  // Constantes medidas (06-graficos.md, 6.4), en pt.
  const MARGEN = 6;            // del borde de la forma al contenido
  const SEPARACION = 6;        // entre título, leyenda y lo demás
  const AIRE_TRAZADO = 5;      // del contenido al trazado si ese lado no tiene etiquetas
  const PAD_TEXTO = { x: 3, y: 1.5 }; // márgenes del cuadro de título y etiquetas
  const AIRE_VALORES = 11.56;  // de las etiquetas de un eje de pie (valores o categorías) al trazado
  const AIRE_CATEGORIAS_IZQ = AIRE_VALORES;
  const AIRE_BORDE = 5.27;     // de la mitad de una etiqueta centrada en el borde del trazado al contenido
  const AIRE_ABAJO = 8;        // alto extra de la fila de etiquetas de abajo
  const BAJA_ETIQUETAS = 7.5;  // del borde del trazado a la línea de las etiquetas de abajo (medido en píxeles)
  const CORRE_VALORES = 0.7;   // los números de pie quedan 0,7 pt más cerca del trazado que su reserva (medido)
  // Leyenda horizontal (medida en píxeles y con LegendEntries por COM, letra de
  // 11,97 pt; se escala con la letra): cada entrada ocupa su texto más 19,3 pt
  // (muestra cuadrada) o 31,8 (de línea), y la fila empieza 6,1 pt adentro.
  const PASO_CAJA = 19.3 / 11.97, PASO_LINEA = 31.83 / 11.97, INICIO_LEYENDA = 6.1;
  // Con una línea de tendencia, las muestras de línea son más largas (15,75 pt
  // desde 7,87 adentro), cada entrada ocupa su texto más 43,2 pt y la fila
  // empieza 11,71 adentro (medido por COM con letra de 11,97).
  const MUESTRAS = {
    comun: { paso: PASO_LINEA, inicio: INICIO_LEYENDA, largo: 1, desde: 6 },
    tendencia: { paso: 43.2 / 11.97, inicio: 11.71, largo: 15.75 / 11.97, desde: 7.87 },
  };
  const muestrasDe = (g) => (g.grupos.some((x) => x.series.some((s) => s.tendencias && s.tendencias.length)) ? MUESTRAS.tendencia : MUESTRAS.comun);

  // Etiquetas de categorías abajo (medido en PowerPoint, 06-graficos.md 6.4 ter):
  // se parten por palabras con el ancho de la categoría (+0,5 de tolerancia);
  // si una palabra no cabe, se giran 45° si la letra cabe en la diagonal
  // (paso·sen 45° ≥ tamaño), 90° si cabe una línea (paso ≥ alto de línea), y
  // si nada alcanza se salta de a 2, 3… Girada, cada etiqueta es una línea
  // anclada por su punta derecha al centro de la categoría.
  const TOLERANCIA_CORTE = 0.5;
  // Tabla de datos (medido por COM, letra de 11,97): filas de 17,99 pt; la
  // columna de nombres mide el texto más 16,13 y empieza 11,02 adentro del
  // margen; la tabla termina 0,68 pt por debajo del límite de abajo.
  const FILA_TABLA = 17.99 / 11.97, COLUMNA_TABLA = 16.13 / 11.97, INICIO_TABLA = 11.02, BAJA_TABLA = 0.68;
  const AIRE_RADAR = 3.5;      // de la punta del radar a sus etiquetas de categorías
  const RADAR_NUMEROS = 2;     // [calibrar] de los números del eje al rayo de arriba
  const LIMITE_IZQ = 4.3;      // hasta dónde puede llegar la primera etiqueta girada (pt de la forma)
  const LARGO_GIRADAS = 0.53;  // [calibrar] largo máximo de una etiqueta girada, en altos del gráfico

  function claseEje(e) { return e.clase === 'valAx' ? 'valor' : e.clase === 'dateAx' ? 'fecha' : 'categoria'; }

  function maquetar(g, W, H, med) {
    const geo = { W, H, ejes: [], series: [] };
    let arriba = MARGEN, abajo = H - MARGEN, izq = MARGEN, der = W - MARGEN;

    if (g.titulo) {
      const l = g.titulo.letra, lh = med.linea(l);
      const lineas = cortar(g.titulo.texto, l, W * 0.8, med);
      const w = Math.max(...lineas.map((t) => med.ancho(t, l))) + 2 * PAD_TEXTO.x, h = lineas.length * lh + 2 * PAD_TEXTO.y;
      // Movido a mano, el resto se diseña igual que si estuviera en su lugar (medido).
      geo.titulo = { ...aplicarDiseno({ x: W / 2 - w / 2, y: arriba, w, h }, g.titulo.diseno, W, H), lineas, letra: l };
      arriba += h + SEPARACION;
    }

    const circular = g.grupos[0].tipo === 'circular' || g.grupos[0].tipo === 'anillo';
    if (g.leyenda) {
      let ley = entradasLeyenda(g, circular);
      const l = g.leyenda.letra, lh = med.linea(l);
      const pos = g.leyenda.pos;
      let anchos = ley.map((e) => med.ancho(e.texto, l));
      // [calibrar] la muestra: 20,4 pt por entrada con letra de 11,97 (medido).
      const M = muestrasDe(g);
      const porEntrada = (e) => (e.larga ? M.paso : PASO_CAJA) * l.tamano;
      const padV = (5.85 / 14.61) * lh;
      let w, h;
      let grilla = null, fila = null, lineas = null;
      if (pos === 'b' || pos === 't') {
        w = M.inicio + ley.reduce((a, e, i) => a + porEntrada(e) + anchos[i], 0);
        h = lh + padV;
        // Una fila (medido por COM, sondeo-leyenda-fila2.ps1): cada entrada
        // mide su texto + 13,14 (muestra cuadrada) o + 24 (línea); la leyenda,
        // 1,2 veces la suma (hasta 0,9 del ancho de la forma), y lo que sobra
        // va repartido parejo antes, entre y después de las entradas.
        if (M === MUESTRAS.comun) {
          const extraFila = (ley.some((e) => e.larga) ? 24 : 13.14) / 11.97 * l.tamano;
          const suma = anchos.reduce((a, x) => a + x + extraFila, 0);
          w = Math.min(1.2 * suma, 0.9 * W);
          fila = { extra: extraFila, hueco: Math.max(0, (w - suma) / (ley.length + 1)) };
        }
        // Si las entradas no caben en una fila (0,9 del ancho de la forma),
        // PowerPoint arma una grilla de columnas iguales: la entrada más
        // ancha (texto + 13,14 con muestra cuadrada, + 24 con línea, letra
        // de 11,97), tantas filas como hagan falta y las columnas repartidas
        // parejo (10 entradas que caben de a 7 → 2 filas de 5). Medido por
        // COM (sondeo-leyenda-filas.ps1).
        const limite = 0.9 * W, extra = (ley.some((e) => e.larga) ? 24 : 13.14) / 11.97 * l.tamano;
        const columna = Math.max(0, ...anchos) + extra;
        if (anchos.reduce((a, x) => a + x + extra, 0) > limite && ley.length > 1) {
          const maximo = Math.max(1, Math.floor(limite / columna + 1e-6));
          let filas = Math.ceil(ley.length / maximo);
          const columnas = Math.ceil(ley.length / filas);
          // Las filas que no caben a lo alto no se muestran.
          const caben = Math.max(1, Math.floor((abajo - arriba) / 2 / (lh + padV) + 1e-6));
          if (filas > caben) { filas = caben; ley = ley.slice(0, filas * columnas); anchos = anchos.slice(0, ley.length); }
          grilla = { columnas, columna }; fila = null;
          w = columnas * columna;
          h = filas * (lh + padV) - (filas > 1 ? 1.49 / 14.61 * lh : 0);
        }
      } else {
        // Las que no caben a lo alto no se muestran (medido: 9 de 12 en un circular).
        const caben = Math.max(1, Math.floor((abajo - arriba) / (lh + padV) + 1e-6));
        if (ley.length > caben) { ley = ley.slice(0, caben); anchos = anchos.slice(0, caben); }
        // Medido por COM con Aptos: texto + 19,15 (cuadrada) o + 30 (línea).
        // Lo que pasa de un tercio del ancho de la forma se parte en líneas.
        const extraV = (ley.some((e) => e.larga) ? 30 : 19.15) / 11.97 * l.tamano;
        lineas = ley.map((e) => cortar(e.texto, l, Math.max(1, W / 3 - extraV), med));
        anchos = lineas.map((ls) => Math.max(0, ...ls.map((t) => med.ancho(t, l))));
        w = Math.max(0, ...anchos) + extraV;
        h = lineas.reduce((a, ls) => a + ls.length * lh + padV, 0);
      }
      w = Math.min(w, der - izq);
      let x, y;
      if (pos === 'b') { x = W / 2 - w / 2; y = abajo - h; abajo = y - SEPARACION; }
      else if (pos === 't' ) { x = W / 2 - w / 2; y = arriba; arriba = y + h + SEPARACION; }
      else if (pos === 'l') { x = izq; y = (arriba + abajo) / 2 - h / 2; izq = x + w + SEPARACION; }
      else { x = der - w; y = (pos === 'tr' ? arriba : (arriba + abajo) / 2 - h / 2); der = x - SEPARACION; }
      ({ x, y, w, h } = aplicarDiseno({ x, y, w, h }, g.leyenda.diseno, W, H));
      geo.leyenda = { x, y, w, h, entradas: ley, anchos, letra: l, horizontal: pos === 'b' || pos === 't', padV, grilla, fila, lineas };
    }

    // Trazado movido a mano: "outer" incluye las etiquetas de los ejes (lo
    // que queda se reparte igual); "inner" es el trazado mismo.
    const dt = g.disenoTrazado;
    const manual = dt ? aplicarDiseno({ x: izq, y: arriba, w: der - izq, h: abajo - arriba }, dt, W, H) : null;
    if (manual && (dt.objetivo !== 'inner' || circular)) { izq = manual.x; der = manual.x + manual.w; arriba = manual.y; abajo = manual.y + manual.h; }

    if (circular) {
      // El círculo, cuadrado y centrado en lo que queda. Con etiquetas de datos,
      // un 14 % más chico (medido: 140,58 de 163,81 y 163,74 de 190,27). [calibrar]
      const zona = { x: izq + AIRE_TRAZADO, y: arriba + AIRE_TRAZADO, w: der - izq - 2 * AIRE_TRAZADO, h: abajo - arriba - 2 * AIRE_TRAZADO };
      let lado = Math.max(0, Math.min(zona.w, zona.h));
      const s0 = g.grupos[0].series[0];
      const conEtiquetas = s0 && s0.etiquetas && !s0.etiquetas.borrar && (s0.etiquetas.showVal || s0.etiquetas.showPercent || s0.etiquetas.showCatName || s0.etiquetas.showSerName);
      if (conEtiquetas) lado *= 0.86;
      geo.trazado = { x: zona.x + zona.w / 2 - lado / 2, y: zona.y + zona.h / 2 - lado / 2, w: lado, h: lado };
      if (g.grupos[0].es3D && g.grupos[0].tipo === 'circular') {
        // Circular 3D: el contorno (lejos + cerca + espesor de alto, el ancho
        // mayor a lo ancho) entra en el trazado menos 11,5 de ancho y 7 de
        // alto, centrado (medido; ver forma3D).
        const f = forma3D(g.rotX, g.perspectiva);
        const lejos = f.b / (1 + f.u), cerca = f.b / (1 - f.u), medioAncho = 1 / Math.sqrt(1 - f.u * f.u);
        const dw = Math.max(1, zona.w - 11.5), dh = Math.max(1, zona.h - 7);
        const a = Math.max(1, Math.min(dw / (2 * medioAncho), dh / (lejos + cerca + f.d)));
        const arribaC = zona.y + 2.5 + (dh - a * (lejos + cerca + f.d)) / 2;
        geo.trazado = { ...zona };
        geo.circular3D = { cx: zona.x + zona.w / 2, cy: arribaC + lejos * a, a, b: f.b * a, u: f.u, d: f.d * a };
      }
      geo.circular = true;
      return geo;
    }

    // Ejes visibles y su lado.
    const ejes = [];
    const titulosEje = [];
    for (const grupo of g.grupos) {
      for (const id of grupo.ejes) {
        const e = g.ejes.get(id);
        if (!e || ejes.some((x) => x.id === id)) continue;
        ejes.push(e);
      }
    }
    const ejeDe = (grupo, cual) => g.ejes.get(grupo.ejes[cual]);
    // Valores de cada eje de valores (para la escala), según los grupos que lo usan.
    const rangos = new Map();
    for (const grupo of g.grupos) {
      const ex = ejeDe(grupo, 0), ey = ejeDe(grupo, 1);
      const agregar = (e, v) => { if (!e || v == null || !Number.isFinite(v)) return; const r = rangos.get(e.id) || { a: Infinity, b: -Infinity }; r.a = Math.min(r.a, v); r.b = Math.max(r.b, v); rangos.set(e.id, r); };
      if (grupo.tipo === 'burbuja') {
        // Burbujas (medido, 06-graficos.md 6.4 ter): cada eje se agranda un
        // radio de burbuja (f/2 del tramo final, f = diámetro de la mayor) a
        // cada lado; el Y incluye siempre el 0; sin el 5 % de aire.
        const f = fraccionBurbuja(grupo.escalaBurbuja);
        const xs = [], ys = [];
        for (const s of grupo.series) s.valores.forEach((v, i) => { if (v == null) return; ys.push(v); xs.push(s.categorias && typeof s.categorias[i] === 'number' ? s.categorias[i] : i + 1); });
        const extender = (e, lista, conCero) => {
          if (!e || !lista.length) return;
          let a = Math.min(...lista), b = Math.max(...lista);
          if (conCero) { a = Math.min(a, 0); b = Math.max(b, 0); }
          const S = (b - a || Math.abs(b) || 1) / (1 - f), r = f / 2 * S;
          rangos.set(e.id, { a: a - r, b: b + r, exacto: true });
        };
        extender(ex, xs, false); extender(ey, ys, true);
        continue;
      }
      if (grupo.tipo === 'dispersion') {
        for (const s of grupo.series) { (s.categorias || []).forEach((x, i) => agregar(ex, typeof x === 'number' ? x : i + 1)); s.valores.forEach((v) => agregar(ey, v)); }
        continue;
      }
      const apilado = grupo.agrupacion === 'stacked' || grupo.agrupacion === 'percentStacked';
      if (grupo.agrupacion === 'percentStacked') { agregar(ey, 0); agregar(ey, 1); const r = rangos.get(ey.id); r.porcentaje = true; continue; }
      const n = Math.max(0, ...grupo.series.map((s) => s.valores.length));
      if (apilado) {
        for (let i = 0; i < n; i++) {
          let pos = 0, neg = 0;
          for (const s of grupo.series) { const v = s.valores[i]; if (v == null) continue; if (v >= 0) pos += v; else neg += v; }
          agregar(ey, pos); agregar(ey, neg);
        }
      } else for (const s of grupo.series) s.valores.forEach((v) => agregar(ey, v));
    }

    const horizontal = g.grupos.some((x) => x.tipo === 'barra' && x.dir === 'bar');
    // Radar (medido): un cuadrado centrado; arriba y abajo, una línea de
    // etiquetas de categorías más 3,5 pt. La escala, con el lado como alto.
    if (g.grupos[0].tipo === 'radar') {
      const grupo = g.grupos[0], ec = ejeDe(grupo, 0), ev = ejeDe(grupo, 1);
      const conCat = ec && !ec.borrado && ec.posEtiquetas !== 'none';
      const cats = ec ? categoriasDe(g, ec) : [];
      const lc = ec ? ec.letra : g.letraBase;
      const aireV = conCat ? med.linea(lc) + AIRE_RADAR : AIRE_TRAZADO;
      const aireH = conCat ? Math.max(0, ...cats.map((t) => med.ancho(t, lc))) + AIRE_RADAR : AIRE_TRAZADO;
      const lado = Math.max(1, Math.min(abajo - arriba - 2 * aireV, der - izq - 2 * aireH));
      const cx = (izq + der) / 2, cy = (arriba + abajo) / 2;
      const escalas = new Map();
      if (ev) {
        const r = rangos.get(ev.id) || { a: 0, b: 1 };
        escalas.set(ev.id, escala(r.a, r.b, divisiones(lado, ev.letra.tamano, med.linea(ev.letra), false), { min: ev.minFijo, max: ev.maxFijo, unidad: ev.unidadFija, sinAire: true }));
      }
      geo.trazado = { x: cx - lado / 2, y: cy - lado / 2, w: lado, h: lado };
      geo.radar = { cx, cy, R: lado / 2, cats, ec, ev, conCat };
      geo.escalas = escalas;
      geo.todosLosEjes = ejes;
      return geo;
    }
    // Títulos de eje (medido): su cuadro es la letra más 1,5 pt arriba y abajo
    // (3 a los lados), a 5 pt de lo de afuera y a 1 pt de las etiquetas.
    for (const e of ejes) {
      const ti = e.titulo || (e.unidades && e.unidades.rotulo);
      if (!ti || e.borrado) continue;
      const lh = med.linea(ti.letra), largo = med.ancho(ti.texto, ti.letra) + 2 * PAD_TEXTO.x;
      const girado = Math.abs(ti.giro) === 90;
      const w = girado ? lh + 2 * PAD_TEXTO.y : largo, h = girado ? largo : lh + 2 * PAD_TEXTO.y;
      const caja = { w, h, girado, eje: e, titulo: ti };
      if (e.pos === 'l') { caja.x = izq + 5; izq = caja.x + w + 1; }
      else if (e.pos === 'r') { caja.x = der - 5 - w; der = caja.x - 1; }
      else if (e.pos === 'b') { caja.y = abajo - 5 - h; abajo = caja.y - 1; }
      else { caja.y = arriba + 5; arriba = caja.y + h + 1; }
      titulosEje.push(caja);
    }
    // Lado de cada eje: el de categorías de un gráfico de barras horizontales va
    // a la izquierda; los de valores, según axPos.
    const lado = (e) => e.pos;
    const visibles = ejes.filter((e) => !e.borrado && e.posEtiquetas !== 'none');
    const etiquetasDeEje = (e, esc) => {
      if (claseEje(e) === 'valor' && esc) {
        const lista = [];
        const formato = e.formato.enlazado ? null : e.formato.codigo;
        const div = e.unidades ? e.unidades.divisor : 1;
        const dec = Math.max(decimalesDe(esc.unidad / div), decimalesDe(esc.min / div));
        for (const r of marcas(esc)) {
          const x = parseFloat((r / div).toPrecision(12));
          const d = esc.log ? decimalesDe(x) : dec;
          lista.push({ v: r, texto: formato ? formatear(x, formato) : (rangos.get(e.id) || {}).porcentaje ? formatear(x, '0%') : formatoDeSerie(g, e, x, d) });
        }
        return lista;
      }
      return categoriasDe(g, e).map((t, i) => ({ i, texto: t }));
    };

    // Primera pasada: escala con el alto/ancho aproximado, para medir las
    // etiquetas; después se recalcula con el trazado final.
    const escalas = new Map();
    const calcularEscalas = (tw, th) => {
      for (const e of ejes) {
        if (claseEje(e) !== 'valor') continue;
        const r = rangos.get(e.id) || { a: 0, b: 1 };
        const acostado = e.pos === 'b' || e.pos === 't';
        const lh = med.linea(e.letra);
        const n = divisiones(acostado ? tw : th, e.letra.tamano, lh, acostado);
        // Los 3D, como el radar, no dejan aire después del dato (medido).
        const fijo = { min: e.minFijo, max: e.maxFijo, unidad: e.unidadFija, log: e.log, exacto: r.exacto, sinAire: g.grupos.some((x) => x.es3D) };
        if (r.porcentaje) { fijo.min = fijo.min ?? 0; fijo.max = fijo.max ?? 1; }
        escalas.set(e.id, escala(r.a, r.b, n, fijo));
      }
    };
    calcularEscalas(der - izq, abajo - arriba);

    // Las etiquetas de un eje van junto a su línea ("nextTo"): si el eje cruza
    // al otro en el medio (hay negativos), quedan adentro del trazado y no
    // ocupan lugar afuera (medido con valores negativos).
    const adentro = (e) => {
      if (e.posEtiquetas !== 'nextTo') return false;
      const otro = g.ejes.get(e.cruzaCon);
      const s = otro && escalas.get(otro.id);
      if (!s || e.cruza !== 'autoZero') return false;
      return s.min < 0 && s.max > 0;
    };
    // Con crossBetween="midCat" las categorías van en los bordes del trazado:
    // la mitad de la primera y de la última sobresale (medido en áreas).
    const enBordes = (e) => claseEje(e) !== 'valor' && g.grupos.some((x) => x.ejes.includes(e.id) && (g.ejes.get(x.ejes[1]) || {}).entre === 'midCat');

    // Espacio de las etiquetas de cada lado.
    let interior, sobraAnterior = 0;
    const disposiciones = new Map();
    for (let vuelta = 0; vuelta < 5; vuelta++) {
      let sobraVuelta = 0;
      let l = izq, rr = der, t = arriba, b = abajo;
      const reserva = { l: 0, r: 0, t: 0, b: 0 };
      // Lo que las etiquetas partidas o giradas ocupan de más sobre una línea:
      // la escala de valores se calcula como si no estuviera (medido).
      let deMas = 0;
      // Lo que sobresale de las etiquetas del extremo hacia un lado sin etiquetas.
      const extra = { t: 0, b: 0, l: 0, r: 0 };
      for (const e of visibles) {
        const lista = etiquetasDeEje(e, escalas.get(e.id));
        const lh = med.linea(e.letra);
        const ladoE = lado(e);
        const fuera = !adentro(e);
        if (ladoE === 'l' || ladoE === 'r') {
          const w = Math.max(0, ...lista.map((x) => med.ancho(x.texto, e.letra)));
          if (fuera) reserva[ladoE] = Math.max(reserva[ladoE], w + (claseEje(e) === 'valor' ? AIRE_VALORES : AIRE_CATEGORIAS_IZQ));
          // Los números de los extremos sobresalen la mitad de su alto, menos
          // el margen (medido: +1,3 pt con 11,97 pt, arriba y abajo).
          if (claseEje(e) === 'valor') { extra.t = Math.max(extra.t, lh / 2 - MARGEN); extra.b = Math.max(extra.b, lh / 2 - MARGEN); }
        } else {
          // El ancho de antes del recorte de la derecha por la última
          // etiqueta (si no, girar o no puede ir y venir entre vueltas).
          const ancho = interior ? interior.w + sobraAnterior : rr - l;
          const n = Math.max(1, lista.length);
          let lineas = 1;
          let alto = null;
          if (g.tabla && claseEje(e) !== 'valor' && ladoE === 'b') {
            const tl = g.tabla.letra, filas = 1 + g.grupos.reduce((a, x) => a + x.series.length, 0);
            alto = filas * FILA_TABLA * tl.tamano - BAJA_TABLA;
            const nombres = Math.max(0, ...g.grupos.flatMap((x) => x.series.map((sr) => med.ancho(sr.nombre, tl))));
            extra.l = Math.max(extra.l, MARGEN + INICIO_TABLA + nombres + COLUMNA_TABLA * tl.tamano - izq);
          } else if (claseEje(e) !== 'valor') {
            // PowerPoint recorta primero la derecha como si la última
            // etiqueta fuera sin girar y con ese ancho decide si gira o salta.
            let sobraRecta = 0;
            if (!enBordes(e) && lista.length) {
              const ultima = cortar(lista[lista.length - 1].texto, e.letra, ancho / n + TOLERANCIA_CORTE, med);
              sobraRecta = Math.max(0, (0.95 * Math.max(0, ...ultima.map((x) => med.ancho(x, e.letra))) - ancho / n) / (2 - 1 / n));
            }
            const d = disposicion(e, lista, ancho - sobraRecta, H, med, enBordes(e));
            disposiciones.set(e.id, d);
            alto = d.alto;
            if (d.giro) extra.l = Math.max(extra.l, LIMITE_IZQ + d.sobraIzq - izq);
            // El trazado se achica a la derecha lo que la mitad de la última
            // etiqueta (aunque el salto la oculte) pasa de la mitad de su
            // categoría, con el 95 % de su ancho; girada, cuenta el ancho de
            // su renglón (alto de línea / sen del giro), sea cual sea el texto
            // (medido, barrido-salto.ps1 y barrido-etiquetas.ps1).
            if (!enBordes(e) && lista.length) {
              // Se resuelve junto con el ancho final: s = (0,95·w − (ancho − s)/n) / 2.
              const sobra = d.giro ? (0.95 * lh / Math.sin(Math.abs(d.giro) * Math.PI / 180) - ancho / n) / (2 - 1 / n) : sobraRecta;
              if (sobra > 0) { extra.r = Math.max(extra.r, AIRE_TRAZADO + sobra); sobraVuelta = Math.max(sobraVuelta, sobra); }
            }
            if (enBordes(e) && lista.length && !d.giro) {
              extra.l = Math.max(extra.l, med.ancho(lista[0].texto, e.letra) / 2 + AIRE_BORDE);
              extra.r = Math.max(extra.r, med.ancho(lista[lista.length - 1].texto, e.letra) / 2 + AIRE_BORDE);
            }
          } else {
            // El último número sobresale la mitad de su ancho a la derecha.
            const ult = lista[lista.length - 1];
            if (ult) extra.r = Math.max(extra.r, med.ancho(ult.texto, e.letra) / 2 + AIRE_TRAZADO);
          }
          if (fuera) reserva[ladoE] = Math.max(reserva[ladoE], alto != null ? alto : lineas * lh + AIRE_ABAJO);
          if (fuera && alto != null) deMas = Math.max(deMas, alto - lh - AIRE_ABAJO);
        }
      }
      l += Math.max(reserva.l || AIRE_TRAZADO, extra.l);
      rr -= Math.max(reserva.r || AIRE_TRAZADO, extra.r);
      t += reserva.t || AIRE_TRAZADO + Math.max(0, extra.t);
      b -= reserva.b || AIRE_TRAZADO + Math.max(0, extra.b);
      const nuevo = { x: l, y: t, w: Math.max(1, rr - l), h: Math.max(1, b - t) };
      const igual = interior && Math.abs(interior.w - nuevo.w) < 0.01 && Math.abs(interior.h - nuevo.h) < 0.01;
      interior = nuevo;
      sobraAnterior = sobraVuelta;
      calcularEscalas(interior.w, interior.h + deMas);
      if (igual) break;
    }
    if (manual && dt.objetivo === 'inner') { interior = { x: manual.x, y: manual.y, w: Math.max(1, manual.w), h: Math.max(1, manual.h) }; calcularEscalas(interior.w, interior.h); }
    geo.trazado = interior;
    const grupo3D = g.grupos.find((x) => x.es3D && x.tipo === 'barra');
    if (grupo3D && g.ejesRectos) geo.caja3D = caja3D(g, grupo3D, interior);
    // Centrados en el trazado (sin pasar de su borde de arriba o izquierdo).
    for (const c of titulosEje) {
      if (c.x == null) c.x = Math.max(interior.x, interior.x + interior.w / 2 - c.w / 2);
      else c.y = Math.max(interior.y, interior.y + interior.h / 2 - c.h / 2);
    }
    geo.titulosEje = titulosEje;
    geo.horizontal = horizontal;
    geo.escalas = escalas;
    for (const e of visibles) { const lista = etiquetasDeEje(e, escalas.get(e.id)); const acostado = e.pos === 'b' || e.pos === 't'; geo.ejes.push({ eje: e, etiquetas: lista, adentro: adentro(e), enBordes: enBordes(e), disp: claseEje(e) !== 'valor' && acostado ? (disposiciones.get(e.id) || disposicion(e, lista, interior.w, H, med, enBordes(e))) : null }); }
    geo.todosLosEjes = ejes;
    return geo;
  }

  // Columnas y barras 3D con ejes en ángulo recto (medido por COM y en las
  // imágenes de PowerPoint, sondeo-3d.ps1; 06-graficos.md 6.4 quater): una
  // proyección oblicua en la que la profundidad z corre z·sen(rotY) a la
  // derecha y z·sen(rotX) hacia arriba. Cada barra tiene de profundidad su
  // ancho × depthPercent; el piso, eso × (1 + gapDepth), con las barras en
  // el medio. Devuelve el frente (z = 0), la profundidad
  // del piso (D), la de la barra (d) y dónde empieza la barra (a).
  function caja3D(g, grupo, P) {
    const rad = (x) => Math.max(0, Math.min(90, x)) * Math.PI / 180;
    const sx = Math.sin(rad(g.rotY)), sy = Math.sin(rad(g.rotX));
    const n = Math.max(1, ...grupo.series.map((s) => s.valores.length));
    const apilado = grupo.agrupacion === 'stacked' || grupo.agrupacion === 'percentStacked';
    const ns = apilado ? 1 : grupo.series.length, ov = apilado ? 1 : grupo.solapamiento;
    const q = 1 / (n * (ns - (ns - 1) * ov + grupo.separacion)); // ancho de barra por pt del eje de categorías
    const k = q * g.profundidad / 100 * (1 + grupo.separacionProf); // profundidad del piso, ídem
    // El frente tiene la proporción del trazado (alto / ancho igual, en los 8
    // casos medidos) y la caja entra en el trazado menos 5,5 de ancho y 4,3
    // de alto, centrada.
    const dw = Math.max(1, P.w - 5.5), dh = Math.max(1, P.h - 4.3), r = P.h / P.w;
    let A, B, D;
    if (grupo.dir === 'bar') {
      B = Math.max(1, Math.min(dh / (1 + k * sy), dw / (1 / r + k * sx))); D = B * k; A = B / r;
    } else {
      A = Math.max(1, Math.min(dw / (1 + k * sx), dh / (r + k * sy))); D = A * k; B = A * r;
    }
    const anchoCaja = A + D * sx, altoCaja = B + D * sy;
    const x0 = P.x + (P.w - anchoCaja) / 2, y0 = P.y + (P.h - altoCaja) / 2;
    const d = D / (1 + grupo.separacionProf);
    return { F: { x: x0, y: y0 + D * sy, w: A, h: B }, D, d, a: (D - d) / 2, sx, sy };
  }

  // Cómo van las etiquetas de un eje de categorías horizontal en un ancho dado:
  // { giro (grados, 0 o negativo), salto, lineas: [[...]] por etiqueta, alto, sobraIzq }.
  function disposicion(e, lista, ancho, H, med, enBordes) {
    const l = e.letra, lh = med.linea(l);
    const n = Math.max(1, lista.length);
    const paso = ancho / (enBordes ? Math.max(1, n - 1) : n);
    const fijo = e.giro;
    const palabra = Math.max(0, ...lista.flatMap((x) => String(x.texto).split(/\s+/).map((w) => med.ancho(w, l))));
    const probar = (s) => {
      const p = paso * s;
      if (fijo != null) return fijo === 0 || p * Math.abs(Math.sin(fijo * Math.PI / 180)) >= l.tamano ? fijo : undefined;
      if (palabra <= p + TOLERANCIA_CORTE) return 0;
      if (p * Math.SQRT1_2 >= l.tamano) return -45;
      if (p >= lh) return -90;
      return undefined;
    };
    let salto = e.salto || 1, giro = probar(salto);
    if (!e.salto) while (giro === undefined && salto < n) giro = probar(++salto);
    if (giro === undefined) giro = fijo != null ? fijo : -90;
    if (!giro) {
      const lineas = lista.map((x, i) => (i % salto ? [] : cortar(x.texto, l, paso * salto + TOLERANCIA_CORTE, med)));
      return { giro: 0, salto, paso, lineas, alto: Math.max(1, ...lineas.map((x) => x.length)) * lh + AIRE_ABAJO, sobraIzq: 0 };
    }
    const th = Math.abs(giro) * Math.PI / 180, sen = Math.sin(th), cos = Math.cos(th);
    const lineas = lista.map((x, i) => (i % salto ? [] : [acortar(x.texto, l, LARGO_GIRADAS * H, med)]));
    const ancho1 = (i) => (lineas[i] && lineas[i].length ? med.ancho(lineas[i][0], l) : 0);
    const wMax = Math.max(0, ...lineas.map((_, i) => ancho1(i)));
    // Lo que la primera etiqueta se mete a la izquierda de su categoría.
    const sobraIzq = ancho1(0) * cos + lh / 2 * sen - (enBordes ? 0 : paso / 2);
    return { giro, salto, paso, lineas, alto: wMax * sen + lh * cos + AIRE_ABAJO, sobraIzq };
  }
  // Un texto que no cabe en un largo: recortado con "…".
  function acortar(texto, l, largo, med) {
    texto = String(texto);
    if (med.ancho(texto, l) <= largo) return texto;
    let a = 0, b = texto.length;
    while (a < b) { const m = Math.ceil((a + b) / 2); if (med.ancho(texto.slice(0, m) + '…', l) <= largo) a = m; else b = m - 1; }
    return texto.slice(0, a) + '…';
  }

  // Texto de las categorías del eje (de la primera serie que las tenga).
  function categoriasDe(g, e) {
    for (const grupo of g.grupos) {
      if (!grupo.ejes.includes(e.id)) continue;
      const s = grupo.series.find((x) => x.categorias && x.categorias.length);
      const n = Math.max(0, ...grupo.series.map((x) => x.valores.length));
      if (s) return Array.from({ length: Math.max(n, s.categorias.length) }, (_, i) => {
        const c = s.categorias[i];
        // Con un formato propio del eje (no enlazado a los datos), ese.
        return c == null ? '' : typeof c === 'number' ? formatear(c, e.formato && !e.formato.enlazado ? e.formato.codigo : s.formatoCategorias) : String(c);
      });
      return Array.from({ length: n }, (_, i) => String(i + 1));
    }
    return [];
  }
  // El formato de los números del eje: el de la primera serie si está
  // enlazado ("General": con los decimales de la unidad).
  function formatoDeSerie(g, e, v, dec) {
    for (const grupo of g.grupos) {
      if (!grupo.ejes.includes(e.id)) continue;
      const s0 = grupo.series[0];
      // En la dispersión, el eje horizontal lleva el formato de los X.
      const f = s0 && ((grupo.tipo === 'dispersion' || grupo.tipo === 'burbuja') && grupo.ejes[0] === e.id ? s0.formatoCategorias : s0.formato);
      if (f && !/^general$/i.test(f)) return formatear(v, f);
    }
    return formatear(v, 'General', dec);
  }

  function entradasLeyenda(g, circular) {
    const lista = [];
    if (circular) {
      const s = g.grupos[0].series[0];
      const cats = s.categorias || s.valores.map((_, i) => String(i + 1));
      cats.forEach((c, i) => {
        if (g.leyenda.borradas.has(i)) return;
        const p = s.puntos.get(i) || {};
        lista.push({ texto: c == null ? '' : typeof c === 'number' ? formatear(c, s.formatoCategorias) : String(c), relleno: p.relleno !== undefined && p.relleno !== null ? p.relleno : colorPunto(g, s, i), linea: p.linea !== undefined ? p.linea : s.linea, tipoMuestra: 'caja' });
      });
      return lista;
    }
    for (const grupo of g.grupos) {
      // En las barras horizontales la leyenda va al revés, como se ven de
      // arriba abajo (medido en el corpus 16 y 17).
      const series = grupo.tipo === 'barra' && grupo.dir === 'bar' ? [...grupo.series].reverse() : grupo.series;
      for (const s of series) {
        if (g.leyenda.borradas.has(s.idx)) continue;
        const conLinea = grupo.tipo === 'linea' || (grupo.tipo === 'dispersion' && grupo.estiloDispersion !== 'marker') || (grupo.tipo === 'radar' && grupo.estiloRadar !== 'filled');
        // La dispersión de solo puntos ocupa en la leyenda lo mismo que una línea (medido).
        lista.push({ texto: s.nombre, relleno: s.relleno, linea: s.linea, marcador: s.marcador, tipoMuestra: conLinea ? 'linea' : grupo.tipo === 'dispersion' ? 'marcador' : 'caja', serie: s, grupo, larga: (conLinea || grupo.tipo === 'dispersion') && !!s.linea });
      }
    }
    // Las tendencias van al final, con su nombre o "Lineal (serie)".
    for (const grupo of g.grupos) for (const s of grupo.series) (s.tendencias || []).forEach((t) => {
      lista.push({ texto: t.nombre || (TEXTOS.tendencia[t.tipo] || t.tipo) + ' (' + s.nombre + ')', relleno: null, linea: t.linea, tipoMuestra: 'linea', larga: true });
    });
    // En un gráfico combinado con alguna línea, todas las muestras ocupan lo de una línea (medido).
    if (lista.some((e) => e.larga)) for (const e of lista) e.larga = true;
    return lista;
  }
  // Color de un sector sin relleno propio: con varyColors, un acento por punto.
  function colorPunto(g, s, i) {
    return g.grupos[0].variarColores ? { tipo: 'solido', color: g.automatico(i) } : s.relleno;
  }

  // --- Dibujo --------------------------------------------------------------------------------------------
  // pintar: { relleno(ctx, rell, path, w, h), linea(ctx, linea, path) } de
  // pptx-dibujo.js (las mismas rutinas que las formas). Devuelve la geometría
  // en pt (para comparar con PowerPoint).
  function dibujar(ctx, g, anchoPx, altoPx, pxPorPt, pintar) {
    const W = anchoPx / pxPorPt, H = altoPx / pxPorPt;
    const med = medidor(ctx);
    const geo = maquetar(g, W, H, med);
    const k = pxPorPt;
    ctx.save();
    try {
      ctx.scale(k, k);
      // Nada sale del marco del gráfico (un máximo fijo menor que los datos).
      ctx.beginPath(); ctx.rect(0, 0, W, H); ctx.clip();
      const rect = (x, y, w, h) => { const p = new Path2D(); p.rect(x, y, w, h); return p; };
      const rellenar = (rell, path, w, h) => { if (rell) { ctx.save(); ctx.scale(1 / k, 1 / k); const m = new DOMMatrix().scale(k, k); const p2 = new Path2D(); p2.addPath(path, m); pintar.relleno(ctx, rell, p2, w * k, h * k); ctx.restore(); } };
      const trazar = (linea, path) => { if (linea) { ctx.save(); ctx.scale(1 / k, 1 / k); const p2 = new Path2D(); p2.addPath(path, new DOMMatrix().scale(k, k)); pintar.linea(ctx, linea, p2); ctx.restore(); } };

      // Fondo del gráfico y del trazado.
      rellenar(g.fondo.relleno, rect(0, 0, W, H), W, H);
      trazar(g.fondo.linea, rect(0, 0, W, H));
      const t = geo.trazado;
      if (!geo.circular) { rellenar(g.fondoTrazado.relleno, rect(t.x, t.y, t.w, t.h), t.w, t.h); trazar(g.fondoTrazado.linea, rect(t.x, t.y, t.w, t.h)); }

      const salida = { trazado: { ...t }, ejes: [], series: [] };
      if (geo.titulo) salida.titulo = { x: geo.titulo.x, y: geo.titulo.y, w: geo.titulo.w, h: geo.titulo.h };
      if (geo.leyenda) salida.leyenda = { x: geo.leyenda.x, y: geo.leyenda.y, w: geo.leyenda.w, h: geo.leyenda.h };

      if (geo.circular3D) dibujarCircular3D(ctx, g, geo, rellenar, trazar, salida);
      else if (geo.circular) dibujarCircular(ctx, g, geo, med, rellenar, trazar, salida);
      else if (geo.radar) dibujarRadar(ctx, g, geo, med, rellenar, trazar, salida);
      else dibujarEjesYSeries(ctx, g, geo, med, rellenar, trazar, salida);

      if (geo.titulo) {
        const ti = geo.titulo, l = ti.letra, lh = med.linea(l);
        rellenar(g.titulo.relleno, rect(ti.x, ti.y, ti.w, ti.h), ti.w, ti.h);
        trazar(g.titulo.linea, rect(ti.x, ti.y, ti.w, ti.h));
        ti.lineas.forEach((texto, i) => escribir(ctx, texto, l, ti.x + ti.w / 2, ti.y + PAD_TEXTO.y + i * lh + med.ascendente(l), 'center'));
      }
      if (geo.leyenda) dibujarLeyenda(ctx, g, geo, med, rellenar, trazar);
      for (const c of geo.titulosEje || []) {
        const ti = c.titulo, l = ti.letra, lh = med.linea(l);
        rellenar(ti.relleno, rect(c.x, c.y, c.w, c.h), c.w, c.h);
        trazar(ti.linea, rect(c.x, c.y, c.w, c.h));
        const cx = c.x + c.w / 2, cy = c.y + c.h / 2, d = med.ascendente(l) - lh / 2, a = ti.giro * Math.PI / 180;
        escribir(ctx, ti.texto, l, cx - d * Math.sin(a), cy + d * Math.cos(a), 'center', ti.giro);
        (salida.titulosEje = salida.titulosEje || []).push({ x: c.x, y: c.y, w: c.w, h: c.h });
      }
      return salida;
    } finally { ctx.restore(); }
  }

  function escribir(ctx, texto, l, x, yBase, alinear = 'left', giro = 0) {
    ctx.save();
    ctx.font = cssLetra(l, l.tamano);
    ctx.fontKerning = l.kern > 0 && l.tamano >= l.kern ? 'normal' : 'none';
    ctx.letterSpacing = '0px';
    ctx.fillStyle = `rgba(${l.color.r},${l.color.g},${l.color.b},${l.color.a})`;
    ctx.textAlign = alinear; ctx.textBaseline = 'alphabetic';
    if (giro) { ctx.translate(x, yBase); ctx.rotate(giro * Math.PI / 180); ctx.fillText(texto, 0, 0); }
    else ctx.fillText(texto, x, yBase);
    ctx.restore();
  }

  function dibujarEjesYSeries(ctx, g, geo, med, rellenar, trazar, salida) {
    const c3 = geo.caja3D;
    const t = c3 ? c3.F : geo.trazado;
    const escalas = geo.escalas;
    // Posición en el trazado de un valor de un eje de valores.
    const aPos = (e, v) => {
      const s = escalas.get(e.id);
      let u = fraccion(s, v);
      if (e.invertido) u = 1 - u;
      return e.pos === 'b' || e.pos === 't' ? t.x + u * t.w : t.y + t.h - u * t.h;
    };
    // Cuadrícula (debajo de todo).
    for (const e of geo.todosLosEjes) {
      if (!e.cuadricula || e.borrado && !e.cuadricula) continue;
      const vertical = e.pos === 'b' || e.pos === 't';
      if (claseEje(e) === 'valor') {
        const s = escalas.get(e.id);
        if (!s) continue;
        for (const v of marcas(s)) {
          const p = new Path2D(); const q = aPos(e, v);
          if (c3) {
            // 3D: en la pared del fondo y en el piso o la pared de la izquierda.
            const dx = c3.D * c3.sx, dy = c3.D * c3.sy;
            if (vertical) { p.moveTo(q, t.y + t.h); p.lineTo(q + dx, t.y + t.h - dy); p.lineTo(q + dx, t.y - dy); }
            else { p.moveTo(t.x, q); p.lineTo(t.x + dx, q - dy); p.lineTo(t.x + t.w + dx, q - dy); }
          } else if (vertical) { p.moveTo(q, t.y); p.lineTo(q, t.y + t.h); } else { p.moveTo(t.x, q); p.lineTo(t.x + t.w, q); }
          trazar(e.cuadricula, p);
        }
      } else if (!c3) {
        const n = Math.max(1, categoriasDe(g, e).length);
        for (let i = 0; i <= n; i++) {
          const p = new Path2D();
          if (vertical) { const x = t.x + t.w * i / n; p.moveTo(x, t.y); p.lineTo(x, t.y + t.h); } else { const y = t.y + t.h * i / n; p.moveTo(t.x, y); p.lineTo(t.x + t.w, y); }
          trazar(e.cuadricula, p);
        }
      }
    }

    // Series.
    for (const grupo of g.grupos) {
      const ex = g.ejes.get(grupo.ejes[0]), ey = g.ejes.get(grupo.ejes[1]);
      if (!ex || !ey) continue;
      const geoGrupo = grupo.tipo === 'barra' ? barras(g, grupo, ex, ey, t, aPos, escalas)
        : grupo.tipo === 'dispersion' || grupo.tipo === 'burbuja' ? dispersion(grupo, ex, ey, aPos)
          : lineasOAreas(g, grupo, ex, ey, t, aPos, escalas);
      // Áreas primero, después barras, después líneas (como PowerPoint).
      // Las barras y áreas se cortan en el borde del trazado (un máximo fijo
      // menor que los datos); las líneas y marcadores, con un poco de aire.
      const tc = c3 ? geo.trazado : t;
      const recortar = (aire) => { ctx.save(); ctx.beginPath(); ctx.rect(tc.x - aire, tc.y - aire, tc.w + 2 * aire, tc.h + 2 * aire); ctx.clip(); };
      recortar(0);
      for (const sg of geoGrupo) {
        const s = sg.serie;
        if (grupo.tipo === 'burbuja') {
          // Diámetro: el de la mayor es f × el lado menor del trazado; el resto
          // según el área (o el ancho, con sizeRepresents="w").
          const tam = s.tamanos || [];
          const mayor = Math.max(0, ...grupo.series.flatMap((x) => (x.tamanos || []).map((v) => Math.abs(v || 0))));
          const dMax = fraccionBurbuja(grupo.escalaBurbuja) * Math.min(t.w, t.h);
          sg.puntos.forEach((p, i) => {
            const v = tam[i];
            if (!p || v == null || !mayor || (v < 0 && !grupo.burbujasNegativas)) { if (p) p.w = p.h = 0; return; }
            const k = Math.abs(v) / mayor, d = dMax * (grupo.burbujaAncho ? k : Math.sqrt(k));
            const pt = s.puntos.get(i) || {};
            const rell = pt.relleno !== undefined && pt.relleno !== null ? pt.relleno : grupo.variarColores ? { tipo: 'solido', color: g.automatico(i) } : s.relleno;
            const c = new Path2D(); c.ellipse(p.x, p.y, d / 2, d / 2, 0, 0, Math.PI * 2);
            ctx.save(); ctx.translate(p.x - d / 2, p.y - d / 2);
            const local = new Path2D(); local.ellipse(d / 2, d / 2, d / 2, d / 2, 0, 0, Math.PI * 2);
            rellenar(rell, local, d, d);
            ctx.restore();
            trazar(pt.linea !== undefined ? pt.linea : s.linea, c);
            p.x -= d / 2; p.y -= d / 2; p.w = p.h = d;
          });
        } else if (grupo.tipo === 'barra' && c3 && grupo.es3D) {
          // Prisma: frente, arriba (× 0,75) y costado derecho (× 0,64), medido
          // en las imágenes. Se dibujan de izquierda a derecha y de abajo
          // hacia arriba, así cada barra tapa el costado de la anterior.
          const ox = c3.a * c3.sx, oy = c3.a * c3.sy, ex2 = c3.d * c3.sx, ey2 = c3.d * c3.sy;
          sg.puntos.forEach((p, i) => {
            if (!p) return;
            const pt = s.puntos.get(i) || {};
            const rell = pt.relleno !== undefined && pt.relleno !== null ? pt.relleno : s.relleno;
            const lin = pt.linea !== undefined ? pt.linea : s.linea;
            const x = p.x + ox, y = p.y - oy;
            const cara = (puntos, r) => {
              const path = new Path2D(); puntos.forEach(([a, b], j) => (j ? path.lineTo(a, b) : path.moveTo(a, b))); path.closePath();
              const xs = puntos.map((z) => z[0]), ys = puntos.map((z) => z[1]), mx = Math.min(...xs), my = Math.min(...ys);
              ctx.save(); ctx.translate(mx, my);
              const local = new Path2D(); local.addPath(path, new DOMMatrix().translate(-mx, -my));
              rellenar(r, local, Math.max(...xs) - mx, Math.max(...ys) - my);
              ctx.restore();
              trazar(lin, path);
            };
            cara([[x + p.w, y], [x + p.w + ex2, y - ey2], [x + p.w + ex2, y + p.h - ey2], [x + p.w, y + p.h]], sombra(rell, 0.64));
            cara([[x, y], [x + ex2, y - ey2], [x + p.w + ex2, y - ey2], [x + p.w, y]], sombra(rell, 0.75));
            cara([[x, y], [x + p.w, y], [x + p.w, y + p.h], [x, y + p.h]], rell);
            // Como COM: la caja de todo el prisma.
            p.x = x; p.y = y - ey2; p.w += ex2; p.h += ey2;
          });
        } else if (grupo.tipo === 'barra') {
          sg.puntos.forEach((p, i) => {
            if (!p) return;
            const pt = s.puntos.get(i) || {};
            const rell = pt.relleno !== undefined && pt.relleno !== null ? pt.relleno : s.relleno;
            const lin = pt.linea !== undefined ? pt.linea : s.linea;
            const path = new Path2D(); path.rect(p.x, p.y, p.w, p.h);
            ctx.save(); ctx.translate(p.x, p.y);
            const local = new Path2D(); local.rect(0, 0, p.w, p.h);
            rellenar(rell, local, p.w, p.h);
            ctx.restore();
            trazar(lin, path);
          });
        } else if (grupo.tipo === 'area') {
          const path = new Path2D();
          const pts = sg.puntos.filter(Boolean);
          if (pts.length) {
            path.moveTo(pts[0].x, sg.base[0]);
            for (const p of pts) path.lineTo(p.x, p.y);
            for (let i = pts.length - 1; i >= 0; i--) path.lineTo(pts[i].x, sg.base[i]);
            path.closePath();
            ctx.save(); ctx.translate(t.x, t.y);
            const local = new Path2D(); local.addPath(path, new DOMMatrix().translate(-t.x, -t.y));
            rellenar(s.relleno, local, t.w, t.h);
            ctx.restore();
            trazar(s.linea, path);
          }
        }
        // Como COM: en las áreas, cada punto es una caja de alto cero de ancho, del valor a su base.
        salida.series.push({ nombre: s.nombre, puntos: sg.puntos.map((p, i) => (p ? { caja: { x: p.x, y: p.y, w: p.w || 0, h: grupo.tipo === 'area' ? sg.base[i] - p.y : p.h || 0 } } : {})) });
      }
      ctx.restore();
      recortar(12);
      for (const sg of geoGrupo) {
        const s = sg.serie;
        if (grupo.tipo === 'linea' || grupo.tipo === 'dispersion') {
          const pts = sg.puntos;
          const conLinea = grupo.tipo === 'linea' || grupo.estiloDispersion !== 'marker' && s.linea;
          if (conLinea && s.linea) {
            const path = new Path2D();
            const suave = s.suave || grupo.estiloDispersion === 'smoothMarker';
            const tramo = [];
            const cerrarTramo = () => {
              if (tramo.length > 1) {
                path.moveTo(tramo[0].x, tramo[0].y);
                if (suave && tramo.length > 2) curvaSuave(path, tramo); else for (const p of tramo.slice(1)) path.lineTo(p.x, p.y);
              }
              tramo.length = 0;
            };
            for (const p of pts) { if (p) tramo.push(p); else if (g.vacios !== 'span') cerrarTramo(); }
            cerrarTramo();
            trazar(s.linea, path);
          }
          pts.forEach((p, i) => {
            if (!p) return;
            const m = (s.puntos.get(i) || {}).marcador || s.marcador;
            dibujarMarcador(ctx, m, s, p.x, p.y, rellenar, trazar, grupo);
          });
        }
      }
      agregados(ctx, g, grupo, geoGrupo, ex, ey, t, aPos, escalas, trazar);
      ctx.restore();
      // Etiquetas de datos.
      for (const sg of geoGrupo) etiquetasDeSerie(ctx, g, grupo, sg, med, t, salida);
    }

    // Ejes: línea y etiquetas.
    if (g.tabla) dibujarTabla(ctx, g, geo, med, rellenar, trazar);
    for (const { eje: e, etiquetas, adentro, enBordes, disp } of geo.ejes) {
      const lh = med.linea(e.letra), asc = med.ascendente(e.letra);
      // Dónde está la línea del eje: donde cruza al otro (en el cero si hay negativos).
      const otro = g.ejes.get(e.cruzaCon);
      let donde;
      if (otro && claseEje(otro) === 'valor' && escalas.get(otro.id)) {
        const so = escalas.get(otro.id);
        const v = e.cruza === 'max' ? so.max : e.cruza === 'min' ? so.min : typeof e.cruza === 'number' ? e.cruza : Math.min(so.max, Math.max(so.min, 0));
        donde = aPos(otro, v);
      }
      if (claseEje(e) === 'valor') {
        // De pie: alineados contra el trazado, a AIRE_VALORES; acostados, centrados en su marca.
        // Si el otro eje cruza en el medio, junto a la línea del eje (adentro).
        const borde = (lado) => (adentro && donde != null ? donde : lado);
        for (const x of etiquetas) {
          const q = aPos(e, x.v);
          if (e.pos === 'l') escribir(ctx, x.texto, e.letra, borde(t.x) - AIRE_VALORES + CORRE_VALORES, q - lh / 2 + asc, 'right');
          else if (e.pos === 'r') escribir(ctx, x.texto, e.letra, borde(t.x + t.w) + AIRE_VALORES - CORRE_VALORES, q - lh / 2 + asc, 'left');
          else if (e.pos === 'b') escribir(ctx, x.texto, e.letra, q, borde(t.y + t.h) + BAJA_ETIQUETAS + asc, 'center');
          else escribir(ctx, x.texto, e.letra, q, borde(t.y) - BAJA_ETIQUETAS - lh + asc, 'center');
        }
      } else if (!(g.tabla && e.pos === 'b')) {
        const n = Math.max(1, etiquetas.length);
        const vertical = e.pos === 'l' || e.pos === 'r';
        const paso = (vertical ? t.h : t.w) / (enBordes ? Math.max(1, n - 1) : n);
        const desplazamiento = enBordes ? 0 : 0.5;
        const salto = disp ? disp.salto : e.salto || 1;
        // Junto a la línea del eje si cruza en el medio; si no, afuera del trazado.
        const bordeY = adentro && donde != null ? donde : t.y + t.h;
        const bordeX = adentro && donde != null ? donde : e.pos === 'l' ? t.x : t.x + t.w;
        etiquetas.forEach((x, i) => {
          if (i % salto) return;
          const u = paso * (i + desplazamiento);
          // En 3D, centradas en el frente de las barras (medido).
          const centro = (vertical ? (e.invertido ? t.y + u : t.y + t.h - u) : (e.invertido ? t.x + t.w - u : t.x + u)) + (c3 ? (vertical ? -c3.a * c3.sy : c3.a * c3.sx) : 0);
          if (vertical) {
            const xDer = e.pos === 'l' ? bordeX - AIRE_CATEGORIAS_IZQ : bordeX + AIRE_CATEGORIAS_IZQ;
            const lineas = [x.texto];
            lineas.forEach((tx, k) => escribir(ctx, tx, e.letra, xDer, centro - lineas.length * lh / 2 + k * lh + asc, e.pos === 'l' ? 'right' : 'left'));
          } else if (disp && disp.giro) {
            // Girada: la punta derecha (a media altura) en el centro de la categoría.
            const th = -disp.giro * Math.PI / 180;
            const py = bordeY + BAJA_ETIQUETAS + lh / 2 * Math.cos(th);
            const d = asc - lh / 2;
            escribir(ctx, disp.lineas[i][0], e.letra, centro + d * Math.sin(th), py + d * Math.cos(th), 'right', disp.giro);
          } else {
            const lineas = disp ? disp.lineas[i] : cortar(x.texto, e.letra, paso + TOLERANCIA_CORTE, med);
            lineas.forEach((tx, k) => escribir(ctx, tx, e.letra, centro, bordeY + BAJA_ETIQUETAS + k * lh + asc, 'center'));
          }
        });
      }
      // La línea del eje, donde cruza el otro.
      if (e.linea) {
        const p = new Path2D();
        if (e.pos === 'b' || e.pos === 't') { const y = donde ?? (e.pos === 'b' ? t.y + t.h : t.y); p.moveTo(t.x, y); p.lineTo(t.x + t.w, y); }
        else { const x = donde ?? (e.pos === 'l' ? t.x : t.x + t.w); p.moveTo(x, t.y); p.lineTo(x, t.y + t.h); }
        trazar(e.linea, p);
      }
    }
    for (const e of geo.todosLosEjes) {
      const s = escalas.get(e.id);
      const grupoEje = g.grupos.findIndex((x) => x.ejes.includes(e.id));
      const tipoCom = (claseEje(e) === 'valor' && !(g.grupos[grupoEje] && ['dispersion', 'burbuja'].includes(g.grupos[grupoEje].tipo) && (e.pos === 'b' || e.pos === 't'))) ? 2 : 1;
      const secundario = grupoEje > 0 && g.grupos[grupoEje].ejes.some((id) => !g.grupos[0].ejes.includes(id));
      // Eje de fechas: de la primera a la última fecha, de a una unidad base (como lo informa COM).
      let fechas = null;
      if (claseEje(e) === 'fecha') {
        const grupo = g.grupos[grupoEje], serie = grupo && grupo.series.find((x) => x.categorias && typeof x.categorias[0] === 'number');
        const nums = serie ? serie.categorias.filter((x) => typeof x === 'number') : [];
        if (nums.length) fechas = { min: Math.min(...nums), max: Math.max(...nums), unidad: 1 };
      }
      salida.ejes.push({ tipo: tipoCom, grupo: secundario ? 2 : 1, borrado: e.borrado, ...(s ? { min: s.min, max: s.max, unidad: s.unidad } : fechas || {}) });
    }
  }

  // Líneas de tendencia, barras de error y líneas de máximo a mínimo.
  function agregados(ctx, g, grupo, geoGrupo, ex, ey, t, aPos, escalas, trazar) {
    const dispersion = grupo.tipo === 'dispersion' || grupo.tipo === 'burbuja', horizontal = grupo.tipo === 'barra' && grupo.dir === 'bar';
    const n = Math.max(1, ...grupo.series.map((s) => s.valores.length));
    // Dónde cae una x (1, 2, 3… en categorías; el valor en dispersión) a lo largo del eje de categorías.
    const midCat = ey.entre === 'midCat';
    const posX = (x) => {
      if (dispersion) return aPos(ex, x);
      const u = midCat ? (n > 1 ? (x - 1) / (n - 1) : 0.5) : (x - 0.5) / n;
      const largo = horizontal ? t.h : t.w;
      const inv = horizontal ? !ex.invertido : ex.invertido;
      return (horizontal ? t.y : t.x) + largo * (inv ? 1 - u : u);
    };
    const punto = (x, y) => (horizontal ? { x: aPos(ey, y), y: posX(x) } : { x: posX(x), y: aPos(ey, y) });
    for (const sg of geoGrupo) {
      const s = sg.serie;
      const xs = s.valores.map((_, i) => (dispersion ? (s.categorias && typeof s.categorias[i] === 'number' ? s.categorias[i] : i + 1) : i + 1));
      const pares = s.valores.map((v, i) => [xs[i], v]).filter(([x, v]) => v != null && Number.isFinite(v) && Number.isFinite(x));
      for (const tr of s.tendencias || []) {
        const curva = tendencia(tr, pares);
        if (!curva || !tr.linea) continue;
        const p = new Path2D();
        if (curva.puntos) curva.puntos.forEach(([x, y], k) => { const q = punto(x, y); if (k) p.lineTo(q.x, q.y); else p.moveTo(q.x, q.y); });
        else {
          const x0 = Math.min(...pares.map((q) => q[0])) - tr.atras, x1 = Math.max(...pares.map((q) => q[0])) + tr.adelante;
          for (let k = 0; k <= 100; k++) { const x = x0 + (x1 - x0) * k / 100, y = curva.f(x); if (!Number.isFinite(y)) continue; const q = punto(x, y); if (k) p.lineTo(q.x, q.y); else p.moveTo(q.x, q.y); }
        }
        trazar(tr.linea, p);
      }
      for (const eb of s.errores || []) {
        if (!eb.linea) continue;
        const media = pares.reduce((a, q) => a + q[1], 0) / Math.max(1, pares.length);
        const desv = Math.sqrt(pares.reduce((a, q) => a + (q[1] - media) ** 2, 0) / Math.max(1, pares.length - 1));
        const p = new Path2D();
        sg.puntos.forEach((pt, i) => {
          const v = s.valores[i];
          if (!pt || v == null) return;
          let mas, menos, centro = v;
          if (eb.valor === 'percentage') mas = menos = Math.abs(v) * eb.cantidad / 100;
          else if (eb.valor === 'stdDev') { centro = media; mas = menos = desv * eb.cantidad; }
          else if (eb.valor === 'stdErr') mas = menos = desv / Math.sqrt(Math.max(1, pares.length));
          else if (eb.valor === 'cust') { mas = (eb.mas && eb.mas[i]) || 0; menos = (eb.menos && eb.menos[i]) || 0; }
          else mas = menos = eb.cantidad;
          const enX = eb.dir === 'x' && dispersion;
          const desde = eb.tipo === 'plus' ? centro : centro - menos, hasta = eb.tipo === 'minus' ? centro : centro + mas;
          // El extremo de una barra, o el punto de una línea.
          const c = grupo.tipo === 'barra' ? (horizontal ? { x: aPos(ey, v), y: pt.y + pt.h / 2 } : { x: pt.x + pt.w / 2, y: aPos(ey, v) }) : { x: pt.x, y: pt.y };
          const tope = TOPE_ERROR;
          if (enX) {
            const a = aPos(ex, xs[i] - (eb.tipo === 'plus' ? 0 : menos)), b = aPos(ex, xs[i] + (eb.tipo === 'minus' ? 0 : mas));
            p.moveTo(a, c.y); p.lineTo(b, c.y);
            if (!eb.sinTope) for (const q of [a, b]) { p.moveTo(q, c.y - tope); p.lineTo(q, c.y + tope); }
          } else if (horizontal) {
            const a = aPos(ey, desde), b = aPos(ey, hasta);
            p.moveTo(a, c.y); p.lineTo(b, c.y);
            if (!eb.sinTope) for (const q of [a, b]) { p.moveTo(q, c.y - tope); p.lineTo(q, c.y + tope); }
          } else {
            const a = aPos(ey, desde), b = aPos(ey, hasta);
            p.moveTo(c.x, a); p.lineTo(c.x, b);
            if (!eb.sinTope) for (const q of [a, b]) { p.moveTo(c.x - tope, q); p.lineTo(c.x + tope, q); }
          }
        });
        trazar(eb.linea, p);
      }
    }
    if (grupo.maxMin && geoGrupo.length) {
      const p = new Path2D();
      for (let i = 0; i < n; i++) {
        const ys = geoGrupo.map((sg) => sg.puntos[i]).filter(Boolean).map((q) => q.y);
        const x = geoGrupo.map((sg) => sg.puntos[i]).find(Boolean);
        if (ys.length < 2 || !x) continue;
        p.moveTo(x.x, Math.min(...ys)); p.lineTo(x.x, Math.max(...ys));
      }
      trazar(grupo.maxMin, p);
    }
  }
  const TOPE_ERROR = 3.75; // [calibrar] media longitud del tope de una barra de error, en pt

  // Ajuste de una línea de tendencia (cuadrados mínimos, como Excel): devuelve
  // { f(x) } o, la media móvil, { puntos }.
  function tendencia(tr, pares) {
    if (pares.length < 2) return null;
    if (tr.tipo === 'movingAvg') {
      const k = Math.max(2, Math.round(tr.periodo));
      const puntos = [];
      for (let i = k - 1; i < pares.length; i++) puntos.push([pares[i][0], pares.slice(i - k + 1, i + 1).reduce((a, q) => a + q[1], 0) / k]);
      return { puntos };
    }
    const recta = (ps, corte) => {
      const m = ps.length, sx = ps.reduce((a, q) => a + q[0], 0), sy = ps.reduce((a, q) => a + q[1], 0);
      const sxx = ps.reduce((a, q) => a + q[0] * q[0], 0), sxy = ps.reduce((a, q) => a + q[0] * q[1], 0);
      if (corte != null) { const b = (sxy - corte * sx) / (sxx || 1); return [corte, b]; }
      const b = (m * sxy - sx * sy) / ((m * sxx - sx * sx) || 1);
      return [(sy - b * sx) / m, b];
    };
    if (tr.tipo === 'linear') { const [a, b] = recta(pares, tr.corte); return { f: (x) => a + b * x }; }
    if (tr.tipo === 'exp') { const ps = pares.filter((q) => q[1] > 0).map((q) => [q[0], Math.log(q[1])]); const [a, b] = recta(ps, tr.corte != null ? Math.log(tr.corte) : null); return { f: (x) => Math.exp(a + b * x) }; }
    if (tr.tipo === 'log') { const ps = pares.filter((q) => q[0] > 0).map((q) => [Math.log(q[0]), q[1]]); const [a, b] = recta(ps); return { f: (x) => (x > 0 ? a + b * Math.log(x) : NaN) }; }
    if (tr.tipo === 'power') { const ps = pares.filter((q) => q[0] > 0 && q[1] > 0).map((q) => [Math.log(q[0]), Math.log(q[1])]); const [a, b] = recta(ps); return { f: (x) => (x > 0 ? Math.exp(a) * Math.pow(x, b) : NaN) }; }
    if (tr.tipo === 'poly') {
      const k = Math.max(2, Math.min(6, Math.round(tr.orden)));
      // Ecuaciones normales con eliminación de Gauss.
      const A = Array.from({ length: k + 1 }, (_, i) => Array.from({ length: k + 2 }, (_, j) => (j <= k ? pares.reduce((a, q) => a + Math.pow(q[0], i + j), 0) : pares.reduce((a, q) => a + q[1] * Math.pow(q[0], i), 0))));
      for (let c = 0; c <= k; c++) {
        let piv = c; for (let r = c + 1; r <= k; r++) if (Math.abs(A[r][c]) > Math.abs(A[piv][c])) piv = r;
        [A[c], A[piv]] = [A[piv], A[c]];
        if (!A[c][c]) return null;
        for (let r = 0; r <= k; r++) if (r !== c) { const f = A[r][c] / A[c][c]; for (let j = c; j <= k + 1; j++) A[r][j] -= f * A[c][j]; }
      }
      const coef = A.map((fila, i) => fila[k + 1] / fila[i]);
      return { f: (x) => coef.reduce((a, c, i) => a + c * Math.pow(x, i), 0) };
    }
    return null;
  }

  // Barras: ancho = categoría / (n − (n−1)·solapamiento + separación) (medido:
  // 18,37 pt exactos en el corpus).
  function barras(g, grupo, ex, ey, t, aPos, escalas) {
    const horizontal = grupo.dir === 'bar';
    const n = Math.max(1, ...grupo.series.map((s) => s.valores.length));
    const largo = horizontal ? t.h : t.w;
    const cw = largo / n;
    const apilado = grupo.agrupacion === 'stacked' || grupo.agrupacion === 'percentStacked';
    const ns = apilado ? 1 : grupo.series.length;
    const ov = apilado ? 1 : grupo.solapamiento;
    const bw = cw / (ns - (ns - 1) * ov + grupo.separacion);
    const sy = escalas.get(ey.id);
    const cero = Math.min(sy.max, Math.max(sy.min, 0));
    const totales = [];
    if (grupo.agrupacion === 'percentStacked') for (let i = 0; i < n; i++) totales[i] = grupo.series.reduce((a, s) => a + Math.abs(s.valores[i] || 0), 0) || 1;
    const acumPos = new Array(n).fill(0), acumNeg = new Array(n).fill(0);
    return grupo.series.map((s, k) => {
      const puntos = s.valores.map((v, i) => {
        if (v == null) return null;
        let desde = cero, hasta = v;
        if (apilado) {
          const vv = grupo.agrupacion === 'percentStacked' ? v / totales[i] : v;
          if (vv >= 0) { desde = acumPos[i]; hasta = desde + vv; acumPos[i] = hasta; } else { desde = acumNeg[i]; hasta = desde + vv; acumNeg[i] = hasta; }
        }
        const inicioCat = cw * i + bw * grupo.separacion / 2 + (apilado ? 0 : k * bw * (1 - ov));
        const p0 = aPos(ey, desde), p1 = aPos(ey, hasta);
        if (horizontal) {
          const y = (ex.invertido ? t.y + inicioCat : t.y + t.h - inicioCat - bw);
          return { x: Math.min(p0, p1), y, w: Math.abs(p1 - p0), h: bw, v, positivo: hasta >= desde, horizontal };
        }
        const x = ex.invertido ? t.x + t.w - inicioCat - bw : t.x + inicioCat;
        return { x, y: Math.min(p0, p1), w: bw, h: Math.abs(p1 - p0), v, positivo: hasta >= desde };
      });
      return { serie: s, puntos };
    });
  }
  function lineasOAreas(g, grupo, ex, ey, t, aPos, escalas) {
    const n = Math.max(1, ...grupo.series.map((s) => s.valores.length));
    // crossBetween (en el eje de valores): "between" pone cada punto en el
    // centro de su categoría; "midCat", en los bordes.
    const entre = ey.entre !== 'midCat';
    const xDe = (i) => {
      const u = entre ? (i + 0.5) / n : n > 1 ? i / (n - 1) : 0.5;
      return ex.invertido ? t.x + t.w * (1 - u) : t.x + t.w * u;
    };
    const apilado = grupo.agrupacion === 'stacked' || grupo.agrupacion === 'percentStacked';
    const totales = [];
    if (grupo.agrupacion === 'percentStacked') for (let i = 0; i < n; i++) totales[i] = grupo.series.reduce((a, s) => a + Math.abs(s.valores[i] || 0), 0) || 1;
    const acum = new Array(n).fill(0);
    const sy = escalas.get(ey.id);
    const cero = Math.min(sy.max, Math.max(sy.min, 0));
    return grupo.series.map((s) => {
      const base = [];
      const puntos = Array.from({ length: n }, (_, i) => {
        let v = s.valores[i];
        if (v == null && (grupo.tipo === 'area' || g.vacios === 'zero')) v = 0;
        if (v == null) { base.push(aPos(ey, cero)); return null; }
        if (grupo.agrupacion === 'percentStacked') v = v / totales[i];
        const desde = apilado ? acum[i] : cero;
        const hasta = apilado ? acum[i] + v : v;
        if (apilado) acum[i] = hasta;
        base.push(aPos(ey, grupo.tipo === 'area' ? (apilado ? desde : cero) : desde));
        return { x: xDe(i), y: aPos(ey, hasta), v: s.valores[i] };
      });
      return { serie: s, puntos, base };
    });
  }
  function dispersion(grupo, ex, ey, aPos) {
    return grupo.series.map((s) => ({
      serie: s,
      puntos: s.valores.map((v, i) => {
        if (v == null) return null;
        const xv = s.categorias && typeof s.categorias[i] === 'number' ? s.categorias[i] : i + 1;
        return { x: aPos(ex, xv), y: aPos(ey, v), v };
      }),
    }));
  }
  // Curva suave por los puntos (Catmull-Rom con tensión de Office ≈ 1/6). [calibrar]
  function curvaSuave(path, p) {
    for (let i = 0; i < p.length - 1; i++) {
      const p0 = p[Math.max(0, i - 1)], p1 = p[i], p2 = p[i + 1], p3 = p[Math.min(p.length - 1, i + 2)];
      path.bezierCurveTo(p1.x + (p2.x - p0.x) / 6, p1.y + (p2.y - p0.y) / 6, p2.x - (p3.x - p1.x) / 6, p2.y - (p3.y - p1.y) / 6, p2.x, p2.y);
    }
  }
  const SIMBOLOS = ['diamond', 'square', 'triangle', 'x', 'star', 'circle', 'plus', 'dash', 'dot'];
  function dibujarMarcador(ctx, m, s, x, y, rellenar, trazar, grupo) {
    if (!m || m.simbolo === 'none') return;
    if (grupo && grupo.tipo === 'linea' && !m) return;
    const simbolo = m.simbolo === 'auto' ? SIMBOLOS[s.idx % SIMBOLOS.length] : m.simbolo;
    const r = (m.tamano || 5) / 2;
    const p = new Path2D();
    switch (simbolo) {
      case 'circle': case 'dot': p.ellipse(x, y, simbolo === 'dot' ? r / 2 : r, simbolo === 'dot' ? r / 2 : r, 0, 0, Math.PI * 2); break;
      case 'square': p.rect(x - r, y - r, 2 * r, 2 * r); break;
      case 'diamond': p.moveTo(x, y - r); p.lineTo(x + r, y); p.lineTo(x, y + r); p.lineTo(x - r, y); p.closePath(); break;
      case 'triangle': p.moveTo(x, y - r); p.lineTo(x + r, y + r); p.lineTo(x - r, y + r); p.closePath(); break;
      case 'x': p.moveTo(x - r, y - r); p.lineTo(x + r, y + r); p.moveTo(x + r, y - r); p.lineTo(x - r, y + r); break;
      case 'plus': p.moveTo(x - r, y); p.lineTo(x + r, y); p.moveTo(x, y - r); p.lineTo(x, y + r); break;
      case 'dash': p.moveTo(x - r, y); p.lineTo(x + r, y); break;
      case 'star': p.moveTo(x - r, y - r); p.lineTo(x + r, y + r); p.moveTo(x + r, y - r); p.lineTo(x - r, y + r); p.moveTo(x, y - r); p.lineTo(x, y + r); break;
      default: p.ellipse(x, y, r, r, 0, 0, Math.PI * 2);
    }
    const rell = m.relleno !== undefined ? m.relleno : { tipo: 'solido', color: s.colorAuto };
    const lin = m.linea !== undefined ? m.linea : lineaFina(s.colorAuto);
    if (rell && !['x', 'plus', 'dash', 'star'].includes(simbolo)) {
      ctx.save(); ctx.translate(x - r, y - r);
      const local = new Path2D(); local.addPath(p, new DOMMatrix().translate(-(x - r), -(y - r)));
      rellenar(rell, local, 2 * r, 2 * r);
      ctx.restore();
    }
    trazar(lin, p);
  }

  // Etiquetas de datos de una serie, en su posición (MS-OI29500: la posición
  // por defecto depende del tipo; 06-graficos.md, 6.5).
  function etiquetasDeSerie(ctx, g, grupo, sg, med, t, salida) {
    const s = sg.serie;
    const e = s.etiquetas;
    if (!e || e.borrar || !(e.showVal || e.showCatName || e.showSerName || e.showPercent)) return;
    const l = e.letra, lh = med.linea(l), asc = med.ascendente(l);
    const apilado = grupo.agrupacion === 'stacked' || grupo.agrupacion === 'percentStacked';
    const pos = e.pos || (grupo.tipo === 'barra' ? (apilado ? 'ctr' : 'outEnd') : grupo.tipo === 'area' ? 'ctr' : 'r');
    const registro = salida.series.find((x) => x.nombre === s.nombre && x.puntos.length === sg.puntos.length && !x._etiquetas);
    if (registro) registro._etiquetas = true;
    sg.puntos.forEach((p, i) => {
      if (!p) return;
      const partes = [];
      if (e.showSerName) partes.push(s.nombre);
      if (e.showCatName) partes.push(s.categorias ? String(s.categorias[i] ?? '') : String(i + 1));
      if (e.showVal) partes.push(formatear(p.v, e.formato || s.formato));
      const texto = partes.join(e.separador || (e.showCatName && e.showVal ? '; ' : ', '));
      const w = med.ancho(texto, l) + 2 * PAD_TEXTO.x, h = lh + 2 * PAD_TEXTO.y;
      let cx, cy;
      if (grupo.tipo === 'barra') {
        const horizontal = p.horizontal;
        const extremo = horizontal ? (p.positivo ? p.x + p.w : p.x) : (p.positivo ? p.y : p.y + p.h);
        if (!horizontal) {
          cx = p.x + p.w / 2;
          cy = pos === 'ctr' ? p.y + p.h / 2 : pos === 'inEnd' ? (p.positivo ? p.y + h / 2 : p.y + p.h - h / 2) : pos === 'inBase' ? (p.positivo ? p.y + p.h - h / 2 : p.y + h / 2) : (p.positivo ? extremo - h / 2 : extremo + h / 2);
        } else {
          cy = p.y + p.h / 2;
          cx = pos === 'ctr' ? p.x + p.w / 2 : pos === 'inEnd' ? (p.positivo ? p.x + p.w - w / 2 : p.x + w / 2) : pos === 'inBase' ? (p.positivo ? p.x + w / 2 : p.x + p.w - w / 2) : (p.positivo ? extremo + w / 2 : extremo - w / 2);
        }
      } else {
        // Medido: 5,5 pt del punto al borde de la etiqueta (línea y dispersión, "r"). En
        // las áreas, "ctr" es el medio entre el valor y su base.
        const sep = 5.5;
        cx = pos === 'l' ? p.x - sep - w / 2 : pos === 'ctr' ? p.x : pos === 't' || pos === 'b' ? p.x : p.x + sep + w / 2;
        cy = pos === 't' ? p.y - sep - h / 2 : pos === 'b' ? p.y + sep + h / 2 : grupo.tipo === 'area' && sg.base ? (p.y + sg.base[i]) / 2 : p.y;
      }
      escribir(ctx, texto, l, cx, cy - h / 2 + PAD_TEXTO.y + asc, 'center');
      if (registro && registro.puntos[i]) registro.puntos[i].etiqueta = { texto, x: cx - w / 2, y: cy - h / 2, w, h };
    });
  }

  // Radar: una corona de polígonos (cuadrícula del eje de valores), los
  // rayos (línea del eje de categorías) y cada serie como polígono cerrado,
  // empezando arriba y en el sentido de las agujas del reloj.
  function dibujarRadar(ctx, g, geo, med, rellenar, trazar, salida) {
    const { cx, cy, R, cats, ec, ev, conCat } = geo.radar;
    const grupo = g.grupos[0];
    const esc = ev && geo.escalas.get(ev.id);
    const n = Math.max(1, cats.length, ...grupo.series.map((s) => s.valores.length));
    const angulo = (i) => -Math.PI / 2 + i * 2 * Math.PI / n;
    const radio = (v) => (esc ? Math.max(0, (v - esc.min) / (esc.max - esc.min)) * R : 0);
    const punto = (i, r) => ({ x: cx + r * Math.cos(angulo(i)), y: cy + r * Math.sin(angulo(i)) });
    const poligono = (r) => { const p = new Path2D(); for (let i = 0; i < n; i++) { const q = punto(i, r); if (i) p.lineTo(q.x, q.y); else p.moveTo(q.x, q.y); } p.closePath(); return p; };
    if (esc && ev.cuadricula) for (let k = 1, v = esc.min + esc.unidad; v <= esc.max + esc.unidad * 1e-6 && k < 200; v = esc.min + (++k) * esc.unidad) trazar(ev.cuadricula, poligono(radio(v)));
    for (const l of [ec && ec.cuadricula, ec && !ec.borrado && ec.linea]) {
      if (!l) continue;
      const p = new Path2D();
      for (let i = 0; i < n; i++) { const q = punto(i, R); p.moveTo(cx, cy); p.lineTo(q.x, q.y); }
      trazar(l, p);
    }
    for (const s of grupo.series) {
      const pts = Array.from({ length: n }, (_, i) => (s.valores[i] == null ? null : punto(i, radio(s.valores[i]))));
      const p = new Path2D();
      let primero = true;
      for (const q of pts) { if (!q) continue; if (primero) p.moveTo(q.x, q.y); else p.lineTo(q.x, q.y); primero = false; }
      p.closePath();
      if (s.relleno) {
        const t = geo.trazado;
        ctx.save(); ctx.translate(t.x, t.y);
        const local = new Path2D(); local.addPath(p, new DOMMatrix().translate(-t.x, -t.y));
        rellenar(s.relleno, local, t.w, t.h);
        ctx.restore();
      }
      trazar(s.linea, p);
      pts.forEach((q, i) => { if (q) dibujarMarcador(ctx, (s.puntos.get(i) || {}).marcador || s.marcador, s, q.x, q.y, rellenar, trazar, grupo); });
      salida.series.push({ nombre: s.nombre, puntos: pts.map((q) => (q ? { caja: { x: q.x, y: q.y, w: 0, h: 0 } } : {})) });
    }
    if (ev) salida.ejes.push({ tipo: 2, grupo: 1, borrado: ev.borrado, ...(esc ? { min: esc.min, max: esc.max, unidad: esc.unidad } : {}) });
    // Números del eje: sobre el rayo de arriba, a su izquierda.
    if (esc && ev && !ev.borrado && ev.posEtiquetas !== 'none') {
      const l = ev.letra, lh = med.linea(l), asc = med.ascendente(l);
      for (let k = 0, v = esc.min; v <= esc.max + esc.unidad * 1e-6 && k < 200; v = esc.min + (++k) * esc.unidad) {
        const r = parseFloat(v.toFixed(12));
        const texto = ev.formato.enlazado ? formatear(r, 'General', decimalesDe(esc.unidad)) : formatear(r, ev.formato.codigo);
        escribir(ctx, texto, l, cx - RADAR_NUMEROS, cy - radio(r) - lh / 2 + asc, 'right');
      }
    }
    // Categorías: afuera de cada punta, a AIRE_RADAR.
    if (conCat) {
      const l = ec.letra, lh = med.linea(l), asc = med.ascendente(l);
      cats.forEach((t, i) => {
        const a = angulo(i), sx = Math.cos(a), sy = Math.sin(a);
        const q = punto(i, R + AIRE_RADAR);
        const alinear = Math.abs(sx) < 0.01 ? 'center' : sx > 0 ? 'left' : 'right';
        const y = Math.abs(sy) < 0.01 ? q.y - lh / 2 : sy < 0 ? q.y - lh : q.y;
        escribir(ctx, t, l, q.x, y + asc, alinear);
      });
    }
  }

  // Tabla de datos debajo del trazado: una fila de categorías y una por serie,
  // con la muestra de la serie y sus valores con su formato.
  function dibujarTabla(ctx, g, geo, med, rellenar, trazar) {
    const T = g.tabla, l = T.letra, lh = med.linea(l), asc = med.ascendente(l), t = geo.trazado;
    const fila = FILA_TABLA * l.tamano;
    const series = g.grupos.flatMap((x) => x.series);
    const grupo = g.grupos[0], ec = g.ejes.get(grupo.ejes[0]);
    const cats = ec ? categoriasDe(g, ec) : [];
    const n = Math.max(1, cats.length, ...series.map((s) => s.valores.length));
    const cw = t.w / n, y0 = t.y + t.h;
    const nombres = Math.max(0, ...series.map((s) => med.ancho(s.nombre, l)));
    const x0 = t.x - (nombres + COLUMNA_TABLA * l.tamano);
    const lado = 0.549 * l.tamano;
    const p = new Path2D();
    const linea = (xa, ya, xb, yb) => { p.moveTo(xa, ya); p.lineTo(xb, yb); };
    const abajo = y0 + (1 + series.length) * fila;
    // Contorno: el de las categorías y el de los nombres (la esquina de arriba a la izquierda queda vacía).
    if (T.contorno) { linea(t.x, y0, t.x + t.w, y0); linea(t.x + t.w, y0, t.x + t.w, abajo); linea(x0, abajo, t.x + t.w, abajo); linea(x0, y0 + fila, x0, abajo); linea(x0, y0 + fila, t.x, y0 + fila); linea(t.x, y0, t.x, abajo); }
    if (T.horizontales) for (let k = 1; k <= series.length; k++) linea(k === 1 ? x0 : x0, y0 + k * fila, t.x + t.w, y0 + k * fila);
    if (T.verticales) for (let i = 0; i <= n; i++) linea(t.x + i * cw, y0, t.x + i * cw, abajo);
    trazar(T.linea, p);
    const centroFila = (k) => y0 + k * fila + fila / 2 - lh / 2 + asc;
    cats.forEach((c, i) => escribir(ctx, c, l, t.x + (i + 0.5) * cw, centroFila(0), 'center'));
    series.forEach((s, k) => {
      const yc = y0 + (k + 1) * fila + fila / 2;
      let x = x0 + 3.29;
      if (T.muestras) {
        ctx.save(); ctx.translate(x, yc - lado / 2);
        const r = new Path2D(); r.rect(0, 0, lado, lado);
        rellenar(s.relleno || (s.linea && s.linea.relleno) || { tipo: 'solido', color: s.colorAuto }, r, lado, lado);
        ctx.restore();
        x += lado + 3.5;
      }
      escribir(ctx, s.nombre, l, x, centroFila(k + 1), 'left');
      s.valores.forEach((v, i) => { if (v != null) escribir(ctx, formatear(v, s.formato), l, t.x + (i + 0.5) * cw, centroFila(k + 1), 'center'); });
    });
  }

  // Un relleno liso oscurecido (las caras de los 3D); los demás, igual.
  function sombra(r, f) {
    return r && r.tipo === 'solido' ? { tipo: 'solido', color: { r: Math.round(r.color.r * f), g: Math.round(r.color.g * f), b: Math.round(r.color.b * f), a: r.color.a } } : r;
  }

  // Circular 3D (medido en imágenes de PowerPoint, sondeo-circular-3d.ps1;
  // 06-graficos.md 6.4 quater). La cara de arriba es el círculo en
  // perspectiva: el punto de ángulo θ va a (a·cos θ·k, b·sen θ·k) desde el
  // centro, con k = 1 / (1 − u·sen θ) (lo de adelante, más grande); el
  // costado baja d·k. u, b/a y d/a salen de tablas por elevación (rotX) con
  // la perspectiva 30 de siempre, corregidas por perspectiva. Los costados
  // tienen un degradado de izquierda a derecha: el color × 0,89 en el borde
  // izquierdo, que baja parejo hasta × 0,395 al 73 % del ancho y sigue así.
  const CIRCULAR_3D = {
    u: [[0, 0.27], [15, 0.248], [30, 0.19], [60, 0.074], [80, 0.028], [90, 0]],
    b: [[0, 0], [15, 0.202], [30, 0.463], [60, 0.874], [80, 0.999], [90, 1]],
    d: [[0, 0.33], [15, 0.303], [30, 0.244], [60, 0.104], [80, 0.019], [90, 0]],
  };
  // Por perspectiva (medido solo con rotX 30; [calibrar] más allá de 100).
  const PERSPECTIVA_3D = {
    u: [[0, 0], [30, 1], [60, 2.09], [100, 3.26], [240, 4.5]],
    b: [[0, 1.09], [30, 1], [60, 0.811], [100, 0.447], [240, 0.3]],
    d: [[0, 0.877], [30, 1], [60, 1.217], [100, 1.8], [240, 2.2]],
  };
  function interpolar(tabla, x) {
    if (x <= tabla[0][0]) return tabla[0][1];
    for (let i = 1; i < tabla.length; i++) if (x <= tabla[i][0]) { const [x0, y0] = tabla[i - 1], [x1, y1] = tabla[i]; return y0 + (y1 - y0) * (x - x0) / (x1 - x0); }
    return tabla[tabla.length - 1][1];
  }
  function forma3D(rotX, perspectiva) {
    const e = Math.max(0, Math.min(90, rotX)), p = Math.max(0, Math.min(240, perspectiva));
    return {
      u: Math.min(0.9, interpolar(CIRCULAR_3D.u, e) * interpolar(PERSPECTIVA_3D.u, p)),
      b: interpolar(CIRCULAR_3D.b, e) * interpolar(PERSPECTIVA_3D.b, p),
      d: interpolar(CIRCULAR_3D.d, e) * interpolar(PERSPECTIVA_3D.d, p),
    };
  }

  // Primero los costados de la mitad de adelante y encima las caras de arriba.
  function dibujarCircular3D(ctx, g, geo, rellenar, trazar, salida) {
    const grupo = g.grupos[0], s = grupo.series[0];
    const { cx, cy, a, b, u, d } = geo.circular3D;
    const k = (th) => 1 / (1 - u * Math.sin(th)), kFrente = k(Math.PI / 2);
    const punto = (ox, oy, th, abajo) => { const f = k(th); return [ox + a * Math.cos(th) * f, oy + b * Math.sin(th) * f + (abajo ? d * f / kFrente : 0)]; };
    const arco = (p, ox, oy, desde, hasta, abajo, mover) => {
      const n = Math.max(2, Math.ceil(Math.abs(hasta - desde) / (Math.PI / 90)));
      for (let j = 0; j <= n; j++) { const [x, y] = punto(ox, oy, desde + (hasta - desde) * j / n, abajo); if (j === 0 && mover) p.moveTo(x, y); else p.lineTo(x, y); }
    };
    const vals = s.valores.map((v) => (v != null && v > 0 ? v : 0));
    const total = vals.reduce((x, y) => x + y, 0) || 1;
    const rebanadas = [];
    let ang = (grupo.anguloInicial - 90) * Math.PI / 180;
    vals.forEach((v, i) => {
      const pt = s.puntos.get(i) || {};
      const expl = (pt.explosion != null ? pt.explosion : s.explosion) || 0;
      const da = v / total * Math.PI * 2, medio = ang + da / 2;
      const rell = pt.relleno !== undefined && pt.relleno !== null ? pt.relleno : grupo.variarColores ? { tipo: 'solido', color: g.automatico(i) } : s.relleno;
      rebanadas.push({ i, desde: ang, hasta: ang + da, dx: Math.cos(medio) * expl * a, dy: Math.sin(medio) * expl * b, rell, linea: pt.linea !== undefined ? pt.linea : s.linea });
      ang += da;
    });
    const pintar = (path, rell) => {
      const t = geo.trazado;
      ctx.save(); ctx.translate(t.x, t.y);
      const local = new Path2D(); local.addPath(path, new DOMMatrix().translate(-t.x, -t.y));
      rellenar(rell, local, t.w, t.h);
      ctx.restore();
    };
    const medioAncho = a / Math.sqrt(1 - u * u);
    const degradado = (r) => {
      if (!r || r.tipo !== 'solido') return null;
      const c = r.color, tono = (f) => `rgba(${Math.round(c.r * f)},${Math.round(c.g * f)},${Math.round(c.b * f)},${c.a})`;
      const gr = ctx.createLinearGradient(cx - medioAncho, 0, cx + medioAncho, 0);
      gr.addColorStop(0, tono(0.89)); gr.addColorStop(0.728, tono(0.395)); gr.addColorStop(1, tono(0.395));
      return gr;
    };
    // Costados: la parte de cada rebanada que cae en la mitad de adelante (seno > 0).
    for (const r of rebanadas) {
      for (let vuelta = -1; vuelta <= 1; vuelta++) {
        const ini = Math.max(r.desde, vuelta * 2 * Math.PI), fin = Math.min(r.hasta, vuelta * 2 * Math.PI + Math.PI);
        if (fin <= ini) continue;
        const ox = cx + r.dx, oy = cy + r.dy;
        const p = new Path2D();
        arco(p, ox, oy, ini, fin, false, true);
        arco(p, ox, oy, fin, ini, true, false);
        p.closePath();
        const gr = degradado(r.rell);
        if (gr) { ctx.save(); ctx.fillStyle = gr; ctx.fill(p); ctx.restore(); } else pintar(p, sombra(r.rell, 0.6));
      }
    }
    const registro = { nombre: s.nombre, puntos: [] };
    for (const r of rebanadas) {
      const ox = cx + r.dx, oy = cy + r.dy;
      const p = new Path2D();
      p.moveTo(ox, oy); arco(p, ox, oy, r.desde, r.hasta, false, false); p.closePath();
      pintar(p, r.rell);
      trazar(r.linea, p);
      registro.puntos.push({ caja: { x: ox - medioAncho, y: oy - b / (1 + u), w: 2 * medioAncho, h: b / (1 + u) + b / (1 - u) + d } });
    }
    salida.series.push(registro);
  }

  // Circular: solo la primera serie. Anillo: cada serie es una corona, de
  // adentro hacia afuera, repartidas en partes iguales entre el hueco
  // (holeSize) y el borde (medido por COM: tres series, coronas de 5,87 pt).
  function dibujarCircular(ctx, g, geo, med, rellenar, trazar, salida) {
    const grupo = g.grupos[0];
    const anillo = grupo.tipo === 'anillo';
    const series = anillo ? grupo.series : grupo.series.slice(0, 1);
    const t = geo.trazado;
    const cx = t.x + t.w / 2, cy = t.y + t.h / 2;
    const s0 = series[0];
    const maxExpl = Math.max(s0.explosion || 0, ...[...s0.puntos.values()].map((p) => p.explosion || 0));
    const R = (t.w / 2) / (1 + maxExpl);
    const hueco = anillo ? R * Math.max(0.1, Math.min(0.9, grupo.hueco || 0.5)) : 0;
    const grosor = anillo ? (R - hueco) / series.length : R;
    series.forEach((s, k) => {
      const rIn = anillo ? hueco + k * grosor : 0, rOut = anillo ? rIn + grosor : R;
      const vals = s.valores.map((v) => (v != null && v > 0 ? v : 0));
      const total = vals.reduce((a, b) => a + b, 0) || 1;
      let ang = (grupo.anguloInicial - 90) * Math.PI / 180;
      const registro = { nombre: s.nombre, puntos: [] };
      vals.forEach((v, i) => {
        const a = v / total * Math.PI * 2;
        const pt = s.puntos.get(i) || {};
        const expl = (pt.explosion != null ? pt.explosion : s.explosion) || 0;
        const medio = ang + a / 2;
        const dx = Math.cos(medio) * expl * R, dy = Math.sin(medio) * expl * R;
        const p = new Path2D();
        if (rIn > 0) { p.arc(cx + dx, cy + dy, rOut, ang, ang + a); p.arc(cx + dx, cy + dy, rIn, ang + a, ang, true); p.closePath(); }
        else { p.moveTo(cx + dx, cy + dy); p.arc(cx + dx, cy + dy, rOut, ang, ang + a); p.closePath(); }
        const rell = pt.relleno !== undefined && pt.relleno !== null ? pt.relleno : grupo.variarColores ? { tipo: 'solido', color: g.automatico(i) } : s.relleno;
        ctx.save(); ctx.translate(t.x, t.y);
        const local = new Path2D(); local.addPath(p, new DOMMatrix().translate(-t.x, -t.y));
        rellenar(rell, local, t.w, t.h);
        ctx.restore();
        trazar(pt.linea !== undefined ? pt.linea : s.linea, p);
        registro.puntos.push({ caja: cajaSector(cx + dx, cy + dy, rOut, rIn, ang, ang + a), _medio: medio, _expl: expl, _v: s.valores[i], _pct: vals[i] / total });
        ang += a;
      });
      salida.series.push(registro);
      const e = s.etiquetas;
      if (e && !e.borrar && (e.showVal || e.showCatName || e.showPercent || e.showSerName)) {
        const l = e.letra, lh = med.linea(l), asc = med.ascendente(l);
        registro.puntos.forEach((p, i) => {
          const partes = [];
          if (e.showSerName) partes.push(s.nombre);
          if (e.showCatName) partes.push(s.categorias ? String(s.categorias[i] ?? '') : String(i + 1));
          if (e.showVal) partes.push(formatear(p._v, e.formato || s.formato));
          if (e.showPercent) partes.push(formatear(p._pct, '0%'));
          const sepTexto = e.separador || (e.showCatName && (e.showVal || e.showPercent) ? '\n' : '; ');
          const lineas = partes.join(sepTexto).split('\n');
          const w = Math.max(...lineas.map((x) => med.ancho(x, l))) + 2 * PAD_TEXTO.x, h = lineas.length * lh + 2 * PAD_TEXTO.y;
          const pos = e.pos || (anillo ? 'ctr' : 'bestFit');
          // Afuera (outEnd y, en lo medido, también bestFit): a 6 pt del borde, más
          // la mitad de la etiqueta en esa dirección. [calibrar] cuándo bestFit la deja adentro.
          const fuera = Math.hypot(w / 2 * Math.cos(p._medio), h / 2 * Math.sin(p._medio));
          const radio = anillo || pos === 'ctr' ? (rOut + rIn) / 2 : pos === 'inEnd' ? rOut - Math.max(w, h) / 2 : rOut + 6 + fuera;
          const lx = cx + Math.cos(p._medio) * (radio + p._expl * R), ly = cy + Math.sin(p._medio) * (radio + p._expl * R);
          lineas.forEach((x, k2) => escribir(ctx, x, l, lx, ly - h / 2 + PAD_TEXTO.y + k2 * lh + asc, 'center'));
          p.etiqueta = { texto: lineas.join('\n'), x: lx - w / 2, y: ly - h / 2, w, h };
        });
      }
      for (const p of registro.puntos) { delete p._medio; delete p._expl; delete p._v; delete p._pct; }
    });
  }

  // Caja de un sector (de a0 a a1, en radianes, en el sentido del canvas): los
  // extremos de los arcos, los puntos cardinales que caen adentro y el centro.
  function cajaSector(cx, cy, R, r0, a0, a1) {
    const pts = [];
    const en = (r, a) => pts.push([cx + Math.cos(a) * r, cy + Math.sin(a) * r]);
    en(R, a0); en(R, a1);
    if (r0 > 0) { en(r0, a0); en(r0, a1); } else pts.push([cx, cy]);
    for (let k = Math.ceil(a0 / (Math.PI / 2)); k * Math.PI / 2 <= a1; k++) { en(R, k * Math.PI / 2); if (r0 > 0) en(r0, k * Math.PI / 2); }
    const xs = pts.map((p) => p[0]), ys = pts.map((p) => p[1]);
    const x = Math.min(...xs), y = Math.min(...ys);
    return { x, y, w: Math.max(...xs) - x, h: Math.max(...ys) - y };
  }

  function dibujarLeyenda(ctx, g, geo, med, rellenar, trazar) {
    const L = geo.leyenda, l = L.letra, lh = med.linea(l), asc = med.ascendente(l);
    const rect = (x, y, w, h) => { const p = new Path2D(); p.rect(x, y, w, h); return p; };
    rellenar(g.leyenda.relleno, rect(L.x, L.y, L.w, L.h), L.w, L.h);
    trazar(g.leyenda.linea, rect(L.x, L.y, L.w, L.h));
    // Medido por COM (LegendKey): la muestra cuadrada mide 0,549 × la letra y
    // empieza 3,29 pt adentro de su entrada; la de línea, 12 pt (1 × la letra)
    // desde 6 pt adentro. Vertical: las entradas empiezan 1,5 pt adentro.
    const lado = 0.549 * l.tamano;
    const M = muestrasDe(g);
    const pasos = L.entradas.map((en, i) => (en.larga ? M.paso : PASO_CAJA) * l.tamano + L.anchos[i]);
    let x = L.horizontal ? L.x + (L.w - M.inicio - pasos.reduce((a, b) => a + b, 0)) / 2 + M.inicio : L.x + 1.5, y = L.y + L.padV / 2;
    L.entradas.forEach((en, i) => {
      if (L.grilla) { x = L.x + (i % L.grilla.columnas) * L.grilla.columna; y = L.y + L.padV / 2 + Math.floor(i / L.grilla.columnas) * (lh + L.padV); }
      else if (L.fila) x = L.x + L.fila.hueco * (i + 1) + L.entradas.slice(0, i).reduce((a, _, j) => a + L.anchos[j] + L.fila.extra, 0);
      const lineasEntrada = L.lineas ? L.lineas[i] : [en.texto];
      const largoMuestra = en.larga ? M.largo * l.tamano : lado;
      const centroY = y + lineasEntrada.length * lh / 2;
      const mx = x + (en.larga ? M.desde : 3.29);
      if (en.tipoMuestra === 'linea') {
        const p = new Path2D(); p.moveTo(mx, centroY); p.lineTo(mx + largoMuestra, centroY);
        trazar(en.linea, p);
        if (en.marcador) dibujarMarcador(ctx, en.marcador, en.serie, mx + largoMuestra / 2, centroY, rellenar, trazar, en.grupo);
      } else if (en.tipoMuestra === 'marcador') {
        dibujarMarcador(ctx, en.marcador || { simbolo: 'auto', tamano: 5 }, en.serie, mx + largoMuestra / 2, centroY, rellenar, trazar, en.grupo);
      } else {
        // La muestra de una serie de barras: el relleno; el borde blanco de los
        // sectores de un circular no se dibuja en la leyenda.
        ctx.save(); ctx.translate(mx, centroY - lado / 2);
        rellenar(en.relleno, rect(0, 0, largoMuestra, lado), largoMuestra, lado);
        ctx.restore();
        const blanca = en.linea && en.linea.relleno && en.linea.relleno.tipo === 'solido' && en.linea.relleno.color.r === 255 && en.linea.relleno.color.g === 255 && en.linea.relleno.color.b === 255;
        trazar(blanca ? null : en.linea, rect(mx, centroY - lado / 2, largoMuestra, lado));
      }
      const tx = mx + largoMuestra + 3.5; // medido en píxeles
      lineasEntrada.forEach((t, k) => escribir(ctx, t, l, tx, y + asc + k * lh, 'left'));
      if (L.horizontal) x += pasos[i];
      else y += lineasEntrada.length * lh + L.padV;
    });
  }

  // Fuentes que hay que cargar antes de medir.
  function letras(g) {
    const lista = [g.letraBase];
    if (g.titulo) lista.push(g.titulo.letra);
    if (g.leyenda) lista.push(g.leyenda.letra);
    for (const e of g.ejes.values()) lista.push(e.letra);
    for (const grupo of g.grupos) for (const s of grupo.series) if (s.etiquetas) lista.push(s.etiquetas.letra);
    return lista.map((l) => ({ fuente: l.fuente, negrita: l.negrita, cursiva: l.cursiva }));
  }

  return { leer, dibujar, letras, escala, divisiones, formatear };
})();

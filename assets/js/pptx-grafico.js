/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
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
  const { hijo, hijos, ruta } = PptxXml;
  const val = (el, def = null) => (el && el.hasAttribute('val') ? el.getAttribute('val') : def);
  const num = (el, def = null) => { const v = val(el); const n = v == null ? NaN : parseFloat(v); return Number.isFinite(n) ? n : def; };
  const si = (el, def = false) => { if (!el) return def; const v = val(el, '1'); return v === '1' || v === 'true'; };

  // --- Lectura -------------------------------------------------------------------------
  const TIPOS = {
    barChart: 'barra', bar3DChart: 'barra', lineChart: 'linea', line3DChart: 'linea', areaChart: 'area', area3DChart: 'area',
    pieChart: 'circular', pie3DChart: 'circular', ofPieChart: 'circular', doughnutChart: 'anillo', scatterChart: 'dispersion',
    radarChart: 'radar', bubbleChart: 'burbuja', stockChart: 'linea', surfaceChart: 'superficie', surface3DChart: 'superficie',
  };
  // Lo que cambia mucho el dibujo y todavía no sabe hacer: con eso, el
  // gráfico va al respaldo en vez de salir distinto sin aviso.
  const NO_SOPORTADOS = [];
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
  const SOPORTADOS = new Set(['barra', 'linea', 'area', 'circular', 'anillo', 'dispersion', 'radar', 'burbuja', 'superficie']);

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
    // r.excel: el gráfico es de Excel y trae poco formato: sin decir nada, Excel usa
    // su estilo clásico (texto negro, líneas grises, marcador de 7, título de 18 en
    // negrita…), no el de Office 2013 que usa PowerPoint (medido contra Excel por COM).
    const clasico = !!r.excel;
    const GRIS_CLASICO = { r: 137, g: 137, b: 137, a: 1 };
    const letraBase = letraDe(hijo(raiz, 'txPr'), r, { tamano: 10, negrita: false, cursiva: false, fuente: r.fuente('+mn-lt') || 'Calibri', color: clasico ? { r: 0, g: 0, b: 0, a: 1 } : { r: 89, g: 89, b: 89, a: 1 }, kern: 0 });
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
        cuadricula: hijo(el, 'majorGridlines') ? conDefecto(linea(hijo(hijo(el, 'majorGridlines'), 'spPr')), lineaFina(clasico ? GRIS_CLASICO : { r: 217, g: 217, b: 217, a: 1 })) : null,
        linea: ln ? linea(hijo(el, 'spPr')) ?? null : clasico ? lineaFina(GRIS_CLASICO) : null,
        marca: val(hijo(el, 'majorTickMark'), 'cross'), posEtiquetas: val(hijo(el, 'tickLblPos'), 'nextTo'),
        letra: letraDe(hijo(el, 'txPr'), r, letraBase), giro: giroDe(hijo(el, 'txPr')),
        cruza: hijo(el, 'crossesAt') ? num(hijo(el, 'crossesAt')) : val(hijo(el, 'crosses'), 'autoZero'),
        cruzaCon: val(hijo(el, 'crossAx')), entre: val(hijo(el, 'crossBetween'), 'between'),
        salto: num(hijo(el, 'tickLblSkip')),
        unidadBase: val(hijo(el, 'baseTimeUnit')), unidadPrincipal: val(hijo(el, 'majorTimeUnit')),
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
        tipo, es3D: /3DChart$/.test(el.localName) || tipo === 'superficie', dir: val(hijo(el, 'barDir'), 'col'), agrupacion: val(hijo(el, 'grouping'), tipo === 'barra' ? 'clustered' : 'standard'),
        separacion: num(hijo(el, 'gapWidth'), 150) / 100, solapamiento: num(hijo(el, 'overlap'), 0) / 100,
        hueco: num(hijo(el, 'holeSize'), 10) / 100, anguloInicial: num(hijo(el, 'firstSliceAng'), 0),
        variarColores: si(hijo(el, 'varyColors')), estiloDispersion: val(hijo(el, 'scatterStyle'), 'marker'), estiloRadar: val(hijo(el, 'radarStyle'), 'marker'),
        separacionProf: num(hijo(el, 'gapDepth'), 150) / 100, forma: val(hijo(el, 'shape'), 'box'),
        escalaBurbuja: num(hijo(el, 'bubbleScale'), 100), burbujaAncho: val(hijo(el, 'sizeRepresents'), 'area') === 'w', burbujasNegativas: si(hijo(el, 'showNegBubbles')),
        ejes: hijos(el, 'axId').map((a) => val(a)),
        // Superficie: de alambre, y el formato de cada banda del eje de valores.
        alambre: si(hijo(el, 'wireframe')),
        bandas: hijos(hijo(el, 'bandFmts'), 'bandFmt').map((b) => ({ idx: num(hijo(b, 'idx'), 0), relleno: r.relleno(hijo(b, 'spPr')), linea: linea(hijo(b, 'spPr')) })),
        etiquetas: etiquetasDe(hijo(el, 'dLbls'), r, letraBase, null),
        maxMin: hijo(el, 'hiLowLines') ? conDefecto(linea(hijo(hijo(el, 'hiLowLines'), 'spPr')), lineaFina({ r: 0, g: 0, b: 0, a: 1 })) : null,
        arribaAbajo: barrasArribaAbajo(hijo(el, 'upDownBars'), r, linea),
        de: el.localName === 'ofPieChart' ? {
          barra: val(hijo(el, 'ofPieType'), 'pie') === 'bar', corte: val(hijo(el, 'splitType'), 'auto'), pos: num(hijo(el, 'splitPos')),
          propios: hijos(hijo(el, 'custSplit'), 'secondPiePt').map((x) => num(x)), segundo: num(hijo(el, 'secondPieSize'), 75) / 100,
          lineas: hijo(el, 'serLines') ? conDefecto(linea(hijo(hijo(el, 'serLines'), 'spPr')), lineaFina({ r: 0, g: 0, b: 0, a: 1 })) : null,
        } : null,
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
          // Relleno de imagen: estirada (stretch, lo de siempre) o apilada.
          // [calibrar] stackScale (apilada cada pictureStackUnit unidades del eje) va como stack.
          imagen: val(ruta(s, 'pictureOptions', 'pictureFormat'), 'stretch'),
          // Barras 3D: caja, cilindro, cono o pirámide (la de la serie, si no la del grupo).
          forma: val(hijo(s, 'shape')) || g.forma,
          suave: si(hijo(s, 'smooth')), explosion: num(hijo(s, 'explosion'), 0) / 100,
          // Sin c:marker, la norma dice automático: en la dispersión (salvo
          // estilo de solo línea) y en las líneas con <c:marker val="1"/> en el grupo.
          marcador: hijo(s, 'marker') ? marcadorDe(hijo(s, 'marker'), r, clasico ? 7 : 5)
            : (tipo === 'dispersion' && !['line', 'smooth'].includes(g.estiloDispersion)) || (tipo === 'linea' && si(hijo(el, 'marker'))) || (tipo === 'radar' && g.estiloRadar === 'marker') ? { simbolo: 'auto', tamano: clasico ? (tipo === 'linea' ? 9 : 7) : 5, relleno: undefined, linea: undefined } : null,
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
          serie.puntos.set(i, { relleno: r.relleno(hijo(p, 'spPr')), linea: linea(hijo(p, 'spPr')), imagen: val(ruta(p, 'pictureOptions', 'pictureFormat')), explosion: hijo(p, 'explosion') ? num(hijo(p, 'explosion'), 0) / 100 : null, marcador: hijo(p, 'marker') ? marcadorDe(hijo(p, 'marker'), r, clasico ? 7 : 5) : null });
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
    // En Excel, los sectores de un circular o un anillo salen un 12 % más oscuros que el acento (medido).
    const sombrear = clasico && grupos.some((g) => g.tipo === 'circular' || g.tipo === 'anillo');
    const automatico = (i) => {
      const acento = r.acento((i % 6) + 1);
      // Los sectores de un circular de Excel van oscurecidos como un `shade` de 76 % en luz lineal (exacto en los 6 colores del tema medidos en el PDF de Excel)
      const sombra = (c) => { const l = c / 255 <= 0.04045 ? c / 255 / 12.92 : Math.pow((c / 255 + 0.055) / 1.055, 2.4), v = 0.76 * l; return Math.round(255 * (v <= 0.0031308 ? v * 12.92 : 1.055 * Math.pow(v, 1 / 2.4) - 0.055)); };
      const base = sombrear ? { r: sombra(acento.r), g: sombra(acento.g), b: sombra(acento.b), a: acento.a } : acento;
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
          s.linea = deLinea ? lineaFina(c, ancho || (clasico ? 38100 : 28575)) : (g.tipo === 'circular' || g.tipo === 'anillo') && !clasico ? lineaFina({ r: 255, g: 255, b: 255, a: 1 }, ancho || 19050) : null;
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
      const letra = letraDe(ruta(tituloEl, 'tx', 'rich') || hijo(tituloEl, 'txPr'), r, clasico ? { ...letraBase, tamano: 18, negrita: true } : { ...letraBase, tamano: letraBase.tamano * 1.4 });
      let texto = textoRico(ruta(tituloEl, 'tx', 'rich')) || textoDe(hijo(tituloEl, 'tx'));
      // Sin texto propio: el nombre de la serie si hay una sola (como PowerPoint).
      if (!texto) { const todas = grupos.flatMap((g) => g.series); texto = todas.length === 1 ? todas[0].nombre : TEXTOS.titulo; }
      titulo = { texto, letra, diseno: disenoDe(tituloEl), relleno: r.relleno(hijo(tituloEl, 'spPr')) ?? null, linea: linea(hijo(tituloEl, 'spPr')) ?? null };
    }
    // Excel, con autoTitleDeleted="0" y sin c:title, pone de título el nombre de
    // la serie si hay una sola (medido por COM). [calibrar] en PowerPoint no se midió.
    if (!tituloEl && clasico && !borrado) {
      const todas = grupos.flatMap((g) => g.series);
      if (todas.length === 1) titulo = { texto: todas[0].nombre, letra: { ...letraBase, tamano: 18, negrita: true }, diseno: null, relleno: null, linea: null };
    }

    const leyendaEl = hijo(chart, 'legend');
    const leyenda = leyendaEl ? {
      pos: val(hijo(leyendaEl, 'legendPos'), 'r'), diseno: disenoDe(leyendaEl), letra: letraDe(hijo(leyendaEl, 'txPr'), r, letraBase),
      borradas: new Set(hijos(leyendaEl, 'legendEntry').filter((e) => si(hijo(e, 'delete'))).map((e) => num(hijo(e, 'idx'), -1))),
      relleno: r.relleno(hijo(leyendaEl, 'spPr')) ?? null, linea: linea(hijo(leyendaEl, 'spPr')) ?? null,
    } : null;

    return {
      grupos, ejes, titulo, leyenda, letraBase, automatico, clasico,
      // <a:noFill/> (null) es transparente; sin decir nada (undefined), blanco.
      fondo: { relleno: conDefecto(r.relleno(hijo(raiz, 'spPr')), { tipo: 'solido', color: { r: 255, g: 255, b: 255, a: 1 } }), linea: linea(hijo(raiz, 'spPr')) ?? (clasico && !hijo(raiz, 'spPr') ? lineaFina(GRIS_CLASICO) : null) },
      fondoTrazado: { relleno: r.relleno(hijo(plot, 'spPr')) ?? null, linea: linea(hijo(plot, 'spPr')) ?? null },
      vacios: val(hijo(chart, 'dispBlanksAs'), 'gap'),
      disenoTrazado: disenoDe(plot),
      tabla: tablaDe(hijo(plot, 'dTable'), r, letraBase, linea),
      rotX: num(ruta(chart, 'view3D', 'rotX'), 30),
      rotY: num(ruta(chart, 'view3D', 'rotY'), 20),
      perspectiva: num(ruta(chart, 'view3D', 'perspective'), 30),
      profundidad: num(ruta(chart, 'view3D', 'depthPercent'), 100),
      ejesRectos: si(ruta(chart, 'view3D', 'rAngAx'), true),
      // Paredes y piso de un 3D: su relleno y su línea (sin spPr, nada) y su grosor.
      paredes: Object.fromEntries([['fondo', 'backWall'], ['lado', 'sideWall'], ['piso', 'floor']].map(([k, n]) => {
        const spPr = ruta(chart, n, 'spPr');
        return [k, { relleno: r.relleno(spPr) ?? null, linea: linea(spPr) ?? null, grosor: parseFloat(val(ruta(chart, n, 'thickness'), '0')) || 0 }];
      })),
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

  // Barras arriba-abajo (c:upDownBars) de un gráfico de líneas: sin spPr,
  // blancas las de subida y negras las de bajada, con borde negro.
  function barrasArribaAbajo(el, r, linea) {
    if (!el) return null;
    const negro = { r: 0, g: 0, b: 0, a: 1 };
    const de = (nombre, color) => {
      const spPr = hijo(hijo(el, nombre), 'spPr');
      return { relleno: conDefecto(r.relleno(spPr), { tipo: 'solido', color }), linea: conDefecto(linea(spPr), lineaFina(negro)) };
    };
    return { separacion: num(hijo(el, 'gapWidth'), 150) / 100, arriba: de('upBars', { r: 255, g: 255, b: 255, a: 1 }), abajo: de('downBars', negro) };
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
  function marcadorDe(el, r, tamanoAuto = 5) {
    if (!el) return null;
    const simbolo = val(hijo(el, 'symbol'), 'auto');
    const l = r.linea(hijo(el, 'spPr'));
    return { simbolo, tamano: num(hijo(el, 'size'), tamanoAuto), relleno: r.relleno(hijo(el, 'spPr')), linea: l && l.sinColor ? undefined : l };
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
      // Líneas guía (circular): de la etiqueta corrida a su sector.
      guias: hijo(el, 'showLeaderLines') ? si(hijo(el, 'showLeaderLines')) : !!h.guias,
      lineaGuia: hijo(el, 'leaderLines') ? r.linea(ruta(el, 'leaderLines', 'spPr')) : h.lineaGuia,
      // Cada etiqueta con algo propio (c:dLbl): borrada o movida a mano.
      propias: new Map(hijos(el, 'dLbl').map((d) => [num(hijo(d, 'idx'), -1), { borrar: si(hijo(d, 'delete')), diseno: disenoDe(d) }])),
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
    // Fechas y horas (número de serie de Excel, sistema 1900). "m/d/yyyy" es
    // la fecha corta del sistema (formato 14 de Excel): Office la muestra en
    // el idioma de quien mira (2/3/2026 en español, 3/2/2026 en inglés).
    if (/^m\/d\/yyyy$/i.test(fmt.trim())) return new Intl.DateTimeFormat(undefined, { day: 'numeric', month: 'numeric', year: 'numeric', timeZone: 'UTC' }).format(new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000));
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
  const cssLetra = (l, px) => PptxTexto.fuenteCss(l, px);
  function medidor(ctx) {
    const cache = new Map();
    return {
      ancho(texto, l) {
        const clave = l.fuente + l.tamano + l.negrita + l.cursiva + '|' + texto;
        if (!cache.has(clave)) {
          ctx.save();
          ctx.font = cssLetra(l, 100);
          ctx.fontKerning = PptxTexto.kerningDe(l, l.tamano);
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
  const AIRE_VALORES_EXCEL = 10.0; // en un gráfico de Excel las etiquetas de los valores quedan 1,5 pt más cerca del trazado (medido en pantalla y en el PDF)
  // Eje de series de un 3D (medido en 25-graficos-3d): del borde del piso al
  // nombre y lo que el trazado se achica a la derecha además del nombre.
  const AIRE_SERIES_3D = 14.7, RESERVA_SERIES_3D = 7.8;
  const AIRE_SERIES_CONTORNO = 7.5, MARCA_3D = 4.5, SALTO_SERIES_3D = 1.05; // [calibrar] los tres (superficies)
  const AIRE_BORDE = 5.27;     // de la mitad de una etiqueta centrada en el borde del trazado al contenido
  const AIRE_ABAJO = 8;        // alto extra de la fila de etiquetas de abajo
  const MARCA_EJE = 4.5;       // largo de las marcas de un eje 2D (medido en pantalla en Excel)
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
  // Un 3D "estándar": cada serie en su propia fila de profundidad, con el eje
  // de series (c:serAx) a lo largo del piso.
  function enFilas3D(grupo) { return !!grupo && grupo.es3D && grupo.agrupacion === 'standard' && grupo.tipo !== 'circular'; }

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
      const padV = (g.clasico ? 5.875 / 12.195 : 5.85 / 14.61) * lh; // en un gráfico de Excel la leyenda vertical tiene más aire (paso de 18,07 pt con letra de 10; medido en el PDF y en pantalla)
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
        const extraV = (ley.some((e) => e.larga) ? 30 + (g.clasico ? 6.35 : 0) : 19.15) / 11.97 * l.tamano; // con muestra de línea, en Excel la leyenda es 5,3 pt más ancha (la muestra es más larga)
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
      // El círculo, cuadrado y centrado en lo que queda. Con etiquetas de
      // datos afuera o en "mejor ajuste" (aunque queden adentro), el diámetro
      // es 0,875 × el lado − 2,75 (medido: 163,81 → 140,58, 190,27 → 163,74,
      // 222 → 191,5 y 428 → 371,75, sin importar el tamaño de las
      // etiquetas); al centro o en el extremo interno no se achica.
      const zona = { x: izq + AIRE_TRAZADO, y: arriba + AIRE_TRAZADO, w: der - izq - 2 * AIRE_TRAZADO, h: abajo - arriba - 2 * AIRE_TRAZADO };
      let lado = Math.max(0, Math.min(zona.w, zona.h));
      const s0 = g.grupos[0].series[0], e0 = s0 && s0.etiquetas;
      const conEtiquetas = e0 && !e0.borrar && (e0.showVal || e0.showPercent || e0.showCatName || e0.showSerName);
      const anilloG = g.grupos[0].tipo === 'anillo';
      if (conEtiquetas && (anilloG || ['bestFit', 'outEnd'].includes(e0.pos || 'bestFit'))) lado = anilloG ? lado * 0.86 : Math.max(0, 0.875 * lado - 2.75);
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
      if (g.grupos[0].de) {
        // Circular o barra de circular (medido en imágenes, sondeo-de-circular.ps1;
        // 06-graficos.md 6.4 quinquies): el radio R del primero sale de que
        // todo entre a lo ancho (2R + hueco + el segundo) y a lo alto; el
        // hueco es gapWidth × R (la mitad con barra); el segundo circular
        // tiene radio secondPieSize × R; la barra, ese ancho y el doble de
        // alto. Todo centrado en lo que queda.
        const de = g.grupos[0].de, sep = g.grupos[0].separacion, s2 = de.segundo;
        const hueco = de.barra ? sep / 2 : sep, ancho = 2 + hueco + (de.barra ? s2 : 2 * s2);
        const R = Math.max(0, Math.min(zona.w / ancho, zona.h / (2 * Math.max(1, s2))));
        const x0 = zona.x + zona.w / 2 - ancho * R / 2;
        geo.trazado = { ...zona };
        geo.deCircular = { cx: x0 + R, cy: zona.y + zona.h / 2, R, x2: x0 + (2 + hueco) * R, R2: s2 * R };
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
      const lista = categoriasDe(g, e).map((t, i) => ({ i, texto: t }));
      const F = fechasDe(g, e);
      if (F) lista.pasos = PASOS_FECHA[F.base];
      return lista;
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
    const enBordes = (e) => claseEje(e) !== 'valor' && g.grupos.some((x) => x.ejes.includes(e.id) && (x.tipo === 'superficie' || (g.ejes.get(x.ejes[1]) || {}).entre === 'midCat'));

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
        // El eje de series de un 3D va a la derecha, a lo largo del piso.
        // [calibrar] la reserva (medido un caso: 38,5 pt con "Centro").
        if (e.clase === 'serAx') {
          const w = Math.max(0, ...lista.map((x) => med.ancho(x.texto, e.letra)));
          reserva.r = Math.max(reserva.r, w + RESERVA_SERIES_3D);
          continue;
        }
        const ladoE = lado(e);
        const fuera = !adentro(e);
        if (ladoE === 'l' || ladoE === 'r') {
          const w = Math.max(0, ...lista.map((x) => med.ancho(x.texto, e.letra)));
          if (fuera) reserva[ladoE] = Math.max(reserva[ladoE], w + (g.clasico ? AIRE_VALORES_EXCEL : claseEje(e) === 'valor' ? AIRE_VALORES : AIRE_CATEGORIAS_IZQ));
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
    if (grupo3D && g.ejesRectos) {
      geo.caja3D = caja3D(g, grupo3D, interior);
      // La escala de valores, con el largo del frente de la caja (medido: con
      // las series en filas el frente es mucho más bajo que el trazado y
      // PowerPoint pasa de 0-5 de a 1 a 0-6 de a 2).
      calcularEscalas(geo.caja3D.F.w, geo.caja3D.F.h);
    } else {
      // En perspectiva (rAngAx = 0), y las líneas y áreas 3D: una escena de
      // PptxGrafico3D con su cámara.
      const g3 = g.grupos.find((x) => x.es3D && ['barra', 'linea', 'area', 'superficie'].includes(x.tipo));
      if (g3) {
        geo.escena3D = escena3D(g, g3, interior, enBordes);
        // La escala de valores, con el largo en pantalla del eje (medido: en
        // 25-graficos-3d, con 67 pt PowerPoint va de a 5).
        const e3 = geo.escena3D;
        calcularEscalas(e3.largoValores, e3.largoValores);
        geo.escalaSuperficie = g3.tipo === 'superficie' ? escalas.get(g3.ejes[1]) : null;
      }
    }
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
    const acostada = grupo.dir === 'bar';
    const filas = enFilas3D(grupo) && !acostada ? Math.max(1, grupo.series.length) : 1;
    const ns = apilado || filas > 1 ? 1 : grupo.series.length, ov = apilado || filas > 1 ? 1 : grupo.solapamiento;
    const q = 1 / (n * (ns - (ns - 1) * ov + grupo.separacion)); // ancho de barra por pt del eje de categorías
    const r = P.h / P.w;
    // k: profundidad del piso por pt del eje de categorías; rr: alto / ancho
    // del frente; dw, dh: dónde entra la caja en el trazado.
    let k, rr, dw, dh;
    if (filas > 1) {
      // En filas (medido en 40 casos de sondeo-filas.ps1): cada fila es un
      // paso de categoría × depthPercent, con la barra de fila / (1 +
      // gapDepth) en el medio; el frente es de r × piso((categorías + series)
      // / 2) pasos de alto (cuenta entera de PowerPoint); la caja entra en el
      // trazado menos 4,5 de ancho y 3,6 de alto, centrada (25-graficos-3d;
      // en el sondeo, sin título ni leyenda, quedaba corrida unos 4 pt).
      // [calibrar] gapDepth (el sondeo no lo varió: en COM es de Chart).
      k = filas * g.profundidad / 100 / n;
      rr = r * Math.floor((n + filas) / 2) / n;
      dw = Math.max(1, P.w - 4.5); dh = Math.max(1, P.h - 3.6);
    } else {
      // El frente tiene la proporción del trazado (alto / ancho igual, en los
      // 8 casos medidos) y la caja entra en el trazado menos 5,5 de ancho y
      // 4,3 de alto, centrada.
      k = q * g.profundidad / 100 * (1 + grupo.separacionProf);
      rr = r;
      dw = Math.max(1, P.w - 5.5); dh = Math.max(1, P.h - 4.3);
    }
    // Paredes y piso con grosor (c:thickness): placas por fuera de la caja,
    // de grosor × el alto del frente / 100 (medido en 25-graficos-3d: piso de
    // 15 → 16,5 pt con el frente de 110 pt). En las barras el piso es la
    // pared de la izquierda y la pared lateral, la de abajo. La de la
    // izquierda cuenta aunque no tenga relleno. [calibrar] queda ~1 pt grande.
    const gr = (x) => Math.max(0, (x && x.grosor) || 0) / 100;
    const pa = g.paredes || {};
    const gIzq = gr(acostada ? pa.piso : pa.lado), gAbajo = gr(acostada ? pa.lado : pa.piso), gFondo = gr(pa.fondo);
    let A, B, D;
    if (acostada) {
      B = Math.max(1, Math.min(dh / (1 + gAbajo + (k + gFondo) * sy), dw / (gIzq + 1 / rr + (k + gFondo) * sx))); D = B * k; A = B / rr;
    } else {
      // D = A·k = B·k/rr.
      B = Math.max(1, Math.min(dh / (gAbajo + 1 + (k / rr + gFondo) * sy), dw / (gIzq + 1 / rr + (k / rr + gFondo) * sx))); A = B / rr; D = A * k;
    }
    const grosor = { izq: gIzq * B, abajo: gAbajo * B, fondo: gFondo * B };
    const atras = D + grosor.fondo;
    const anchoCaja = grosor.izq + A + atras * sx, altoCaja = grosor.abajo + B + atras * sy;
    const x0 = P.x + (P.w - anchoCaja) / 2, y0 = P.y + (P.h - altoCaja) / 2;
    const fila = D / filas, d = fila / (1 + grupo.separacionProf);
    return { F: { x: x0 + grosor.izq, y: y0 + atras * sy, w: A, h: B }, D, d, a: (fila - d) / 2, fila, filas, sx, sy, grosor, acostada };
  }

  // La escena 3D de un gráfico en perspectiva (o de líneas y áreas 3D):
  // el mundo en unidades de categoría (x las categorías, y los valores, z la
  // profundidad; en las barras acostadas, x los valores e y las categorías),
  // la cámara de PptxGrafico3D y de qué lado quedan las paredes.
  //  - La caja tiene la proporción del trazado (medido, sondeo-perspectiva).
  //  - En filas (cada serie en su fila: columnas estándar, líneas y áreas):
  //    cada fila es depthPercent de una categoría; la barra o la cinta, la
  //    fila / (1 + gapDepth). Agrupadas: la barra × depthPercent de fondo y
  //    el piso, eso × (1 + gapDepth).
  //  - Con ejes en ángulo recto (líneas y áreas 3D), la misma caja que las
  //    columnas (caja3D) con su proyección oblicua.
  function escena3D(g, grupo, P, enBordes) {
    if (grupo.tipo === 'superficie') return escenaSuperficie(g, grupo, P, enBordes);
    const horizontal = grupo.tipo === 'barra' && grupo.dir === 'bar';
    const n = Math.max(1, ...grupo.series.map((s) => s.valores.length));
    const r = P.h / P.w;
    const apilado = /stacked/i.test(grupo.agrupacion);
    const filas = (enFilas3D(grupo) || grupo.tipo === 'linea') && !horizontal && !apilado ? Math.max(1, grupo.series.length) : 1;
    const dp = g.profundidad / 100, gd = grupo.separacionProf;
    const ns = apilado || filas > 1 ? 1 : grupo.series.length, ov = apilado || filas > 1 ? 1 : grupo.solapamiento;
    const w = 1 / (ns - (ns - 1) * ov + grupo.separacion);
    let fila, d;
    if (filas > 1) { fila = dp; d = fila / (1 + gd); } else { d = w * dp; fila = d * (1 + gd); }
    const D = filas * fila;
    let W, H, cam;
    if (horizontal) { H = n; W = H / r; } else { W = n; H = W * r; }
    if (g.ejesRectos) {
      const c3 = caja3D(g, grupo, P);
      const s = horizontal ? c3.F.h / n : c3.F.w / n; // pt por unidad del mundo
      if (horizontal) W = c3.F.w / s; else H = c3.F.h / s;
      const base = PptxGrafico3D.oblicua(c3.sx, c3.sy);
      cam = {
        p: ([x, y, z]) => { const [X, Y] = base.p([x * s, y * s, z * s]); return [c3.F.x + X, c3.F.y + c3.F.h + Y]; },
        ojo: base.ojo, lejania: (eje, v) => base.lejania(eje, v * s),
      };
    } else cam = PptxGrafico3D.perspectiva(g.rotX, g.rotY, g.perspectiva, { W, H, D }, P);
    // Las paredes van del lado lejano: la del fondo atrás del ojo, la de
    // costado del otro lado del ojo.
    const v = cam.ojo([W / 2, H / 2, D / 2]);
    // Con el ojo de frente (rotY 0, 180…), la de costado queda a la izquierda, como en caja3D.
    const zF = v[2] < 1e-9 ? D : 0, zN = D - zF, xL = v[0] > -1e-9 ? 0 : W;
    const dist = (a, b) => Math.hypot(a[0] - b[0], a[1] - b[1]);
    const largoValores = horizontal ? dist(cam.p([0, 0, zN]), cam.p([W, 0, zN])) : dist(cam.p([xL, 0, zN]), cam.p([xL, H, zN]));
    const ex = g.ejes.get(grupo.ejes[0]);
    const bordes = grupo.tipo !== 'barra' && ex && enBordes(ex);
    return { cam, W, H, D, n, filas, fila, d, horizontal, zF, zN, xL, largoValores: Math.max(1, largoValores), grupo, bordes };
  }

  // Superficies (medido en 26-graficos-superficie, sondeo-superficie.ps1 y
  // las imágenes de PowerPoint; 06-graficos.md 6.4 septies). Los datos van
  // del borde al borde de la caja (c categorías, s series). La caja, medida
  // en pasos de categoría (c pasos de ancho): c de ancho, s de fondo y de
  // alto r' × piso((c + s − 1) / 2), con r' = (alto − 5,06) / (ancho + 4,7)
  // del trazado; entra en ese ancho y alto (anisótropo), y la cámara es la de
  // las columnas en perspectiva. La escala de valores se calcula con 0,52 ×
  // el alto del trazado (con rotX ≥ 80, el alto entero). [calibrar] el
  // centro (−4, +4) y el 0,52.
  const SUP_PARAMS = '{"ha":-3,"da":0.82,"db":0.15,"lw":10.5,"lh":0,"dx":12,"dy":-4.5}';
  function escenaSuperficie(g, grupo, P, enBordes) {
    const c = Math.max(2, ...grupo.series.map((s) => s.valores.length)), s = Math.max(2, grupo.series.length);
    // Los ajustes de la vista en perspectiva valen cos(rotX): de arriba (rotX 90) no hay ninguno.
    const K = JSON.parse(SUP_PARAMS), cx = Math.cos(g.rotX * Math.PI / 180);
    const dw = P.w + 4.7, dh = P.h - 5.06;
    const paso = (c - 1) / c, m = Math.floor((c + s - 1) / 2);
    const W = c - 1, H = ((P.h - 5.06 - K.ha * cx) / dw) * m * paso, D = ((1 + (K.da - 1) * cx) * s + K.db * cx) * paso;
    const cam = PptxGrafico3D.perspectiva(g.rotX, g.rotY, g.perspectiva, { W, H, D }, P, { w: dw - K.lw * cx, h: dh + K.lh * cx, dx: (-4 + K.dx) * cx, dy: (4 + K.dy) * cx });
    const v = cam.ojo([W / 2, H / 2, D / 2]);
    const zF = v[2] < 1e-9 ? D : 0, zN = D - zF, xL = v[0] > -1e-9 ? 0 : W;
    const ex = g.ejes.get(grupo.ejes[0]), se = g.ejes.get(grupo.ejes[2]);
    const largo = (g.rotX >= 80 ? 0.7 : 0.52) * P.h;
    return { cam, W, H, D, n: c, filas: s, fila: D / (s - 1), d: 0, horizontal: false, zF, zN, xL, largoValores: Math.max(1, largo), grupo, bordes: !!(ex && enBordes(ex)), zSerie: (k) => (se && se.invertido ? s - 1 - k : k) * D / (s - 1) };
  }

  // Cómo van las etiquetas de un eje de categorías horizontal en un ancho dado:
  // { giro (grados, 0 o negativo), salto, lineas: [[...]] por etiqueta, alto, sobraIzq }.
  function disposicion(e, lista, ancho, H, med, enBordes) {
    const l = e.letra, lh = med.linea(l);
    const n = Math.max(1, lista.length);
    const paso = ancho / (enBordes ? Math.max(1, n - 1) : n);
    const fijo = e.giro;
    const palabra = Math.max(0, ...lista.flatMap((x) => String(x.texto).split(/\s+/).map((w) => med.ancho(w, l))));
    // En un eje de fechas, girada necesita más lugar: el paso por el seno
    // del giro ≥ 1,5 × la letra (medido en 13 casos de sondeo-fechas.ps1 y
    // el corpus 21).
    const girada = lista.pasos ? 1.5 * l.tamano : null;
    // Y derecha, una fecha no se parte: entra entera y con aire. [calibrar]
    // el aire: con 5,5 pt PowerPoint gira (36 meses "mmm" en 415 pt, de a
    // 2) y con 16,5 pt no (12 meses en 410 pt), en 24-graficos-detalles.
    const derecha = lista.pasos ? Math.max(0, ...lista.map((x) => med.ancho(String(x.texto), l))) + l.tamano : null;
    const probar = (s) => {
      const p = paso * s;
      if (fijo != null) return fijo === 0 || p * Math.abs(Math.sin(fijo * Math.PI / 180)) >= l.tamano ? fijo : undefined;
      if (derecha != null ? derecha <= p : palabra <= p + TOLERANCIA_CORTE) return 0;
      if (p * Math.SQRT1_2 >= (girada || l.tamano)) return -45;
      if (p >= (girada || lh)) return -90;
      return undefined;
    };
    let salto = e.salto || 1, giro = probar(salto);
    // En un eje de fechas, solo los pasos de fecha (y después de a varios del último).
    const siguiente = (x) => { if (!lista.pasos) return x + 1; const p = lista.pasos.find((y) => y > x); return p || x + lista.pasos[lista.pasos.length - 1]; };
    if (!e.salto) while (giro === undefined && salto < n) giro = probar(salto = siguiente(salto));
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

  // Eje de fechas (c:dateAx con fechas en las categorías), medido por COM e
  // imágenes (sondeo-fechas.ps1; 06-graficos.md 6.4 quinquies): del primer
  // al último dato hay un lugar por unidad base (días, meses o años) y cada
  // punto va en el de su fecha, así que las fechas que faltan dejan hueco
  // (los fines de semana de una serie de días hábiles). Sin baseTimeUnit, la
  // unidad sale de las fechas: años si todas caen el mismo día del año,
  // meses si el mismo día del mes, si no días. Las etiquetas son las fechas
  // de cada lugar (también las sin dato) y se saltean de a 1, 2, 7 o 14
  // días, o 1, 2, 3, 6 o 12 meses. [calibrar] PowerPoint a veces prefiere
  // pasar a meses sin girar (60 días en 260 pt).
  const DIA_CERO = Date.UTC(1899, 11, 30);
  const PASOS_FECHA = { days: [1, 2, 7, 14], months: [1, 2, 3, 6, 12], years: [1, 2, 5, 10, 20, 50] };
  function fechasDe(g, e) {
    if (!e || e.clase !== 'dateAx') return null;
    for (const grupo of g.grupos) {
      if (!grupo.ejes.includes(e.id) || grupo.tipo === 'dispersion' || grupo.tipo === 'burbuja') continue;
      const s = grupo.series.find((x) => x.categorias && x.categorias.length);
      if (!s || !s.categorias.every((c) => c == null || typeof c === 'number')) return null;
      const nums = s.categorias.filter((c) => c != null).map((c) => Math.floor(c));
      if (!nums.length) return null;
      const dia = (v) => { const d = new Date(DIA_CERO + v * 86400000); return { y: d.getUTCFullYear(), m: d.getUTCMonth(), d: d.getUTCDate() }; };
      const ds = nums.map(dia);
      let base = e.unidadBase;
      if (!base) base = ds.every((x) => x.d === ds[0].d && x.m === ds[0].m) && nums.length > 1 ? 'years' : ds.every((x) => x.d === ds[0].d) && nums.length > 1 ? 'months' : 'days';
      const min = Math.min(...nums), d0 = dia(min);
      const lugar = (v) => { const x = dia(v); return base === 'days' ? v - min : base === 'months' ? (x.y - d0.y) * 12 + x.m - d0.m : x.y - d0.y; };
      const lugares = s.categorias.map((c) => (c == null ? null : lugar(Math.floor(c))));
      const n = Math.max(...lugares.filter((x) => x != null)) + 1;
      if (n > 5000) return null;
      const fechaDe = (k) => (base === 'days' ? min + k : (Date.UTC(d0.y + (base === 'years' ? k : 0), d0.m + (base === 'months' ? k : 0), d0.d) - DIA_CERO) / 86400000);
      return { base, n, lugares, fechaDe, formato: s.formatoCategorias };
    }
    return null;
  }

  // Texto de las categorías del eje (de la primera serie que las tenga).
  function categoriasDe(g, e) {
    // El eje de series de un 3D: los nombres de las series, de adelante hacia atrás.
    if (e.clase === 'serAx') {
      const grupo = g.grupos.find((x) => x.ejes.includes(e.id));
      return grupo ? grupo.series.map((s) => s.nombre) : [];
    }
    const F = fechasDe(g, e);
    if (F) return Array.from({ length: F.n }, (_, k) => formatear(F.fechaDe(k), e.formato && !e.formato.enlazado ? e.formato.codigo : F.formato));
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

  // Superficies: la escala del eje de valores (la de la pasada anterior del
  // diseño si ya hay, si no una con 4 divisiones) y sus bandas, una por
  // unidad principal: [{ lo, hi, texto, relleno, linea }]. De alambre solo
  // tienen línea; si no, relleno (el de bandFmt o el acento del ciclo).
  function bandasDe(g, grupo) {
    const ev = g.ejes.get(grupo.ejes[1]);
    let esc = g.escalaSuperficie;
    if (!esc && ev) {
      let a = Infinity, b = -Infinity;
      for (const s of grupo.series) for (const v of s.valores) if (v != null && Number.isFinite(v)) { a = Math.min(a, v); b = Math.max(b, v); }
      esc = escala(a, b, 4, { min: ev.minFijo, max: ev.maxFijo, unidad: ev.unidadFija, log: ev.log, sinAire: true });
    }
    if (!esc) return [];
    const niveles = marcas(esc);
    // Con el formato del eje (si no está enlazado) y su unidad de visualización, como las etiquetas.
    const formato = ev && !ev.formato.enlazado ? ev.formato.codigo : null, div = ev && ev.unidades ? ev.unidades.divisor : 1;
    const texto = (v) => { const x = parseFloat((v / div).toPrecision(12)); return formato ? formatear(x, formato) : formatoDeSerie(g, ev, x, Math.max(decimalesDe(esc.unidad / div), decimalesDe(esc.min / div))); };
    const bandas = [];
    for (let k = 0; k + 1 < niveles.length; k++) {
      const f = grupo.bandas.find((x) => x.idx === k) || {};
      const c = g.automatico(k);
      bandas.push({
        lo: niveles[k], hi: niveles[k + 1], texto: texto(niveles[k]) + '-' + texto(niveles[k + 1]),
        relleno: grupo.alambre ? null : f.relleno !== undefined ? f.relleno : { tipo: 'solido', color: c },
        linea: grupo.alambre ? f.linea !== undefined ? f.linea : lineaFina(c) : null,
      });
    }
    return bandas;
  }

  function entradasLeyenda(g, circular) {
    const lista = [];
    if (g.grupos[0].tipo === 'superficie') {
      // Una entrada por banda, de la menor a la mayor.
      bandasDe(g, g.grupos[0]).forEach((b, k) => { if (!g.leyenda.borradas.has(k)) lista.push({ texto: b.texto, relleno: b.relleno, linea: b.linea, tipoMuestra: 'caja' }); });
      return lista;
    }
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
      // arriba abajo (medido en el corpus 16 y 17); apiladas, en su orden,
      // como se ven de izquierda a derecha (25-graficos-3d). Las columnas, las
      // áreas y las líneas apiladas, con la leyenda a un lado, van al revés: como
      // se apilan (medido en Excel por COM; con la leyenda abajo no se midió).
      const apilado = /stacked/i.test(grupo.agrupacion);
      const horizontal = grupo.tipo === 'barra' && grupo.dir === 'bar';
      const alReves = horizontal ? !apilado : apilado && ['r', 'l', 'tr'].includes(g.leyenda.pos);
      const series = alReves ? [...grupo.series].reverse() : grupo.series;
      for (const s of series) {
        if (g.leyenda.borradas.has(s.idx)) continue;
        // Las líneas 3D son cintas: en la leyenda, un cuadro del color de la línea (25-graficos-3d).
        const cinta = grupo.tipo === 'linea' && grupo.es3D;
        const conLinea = (grupo.tipo === 'linea' && !cinta) || (grupo.tipo === 'dispersion' && grupo.estiloDispersion !== 'marker') || (grupo.tipo === 'radar' && grupo.estiloRadar !== 'filled');
        // La dispersión de solo puntos ocupa en la leyenda lo mismo que una línea (medido).
        lista.push({ texto: s.nombre, relleno: cinta ? (s.linea && s.linea.relleno) || s.relleno : s.relleno, linea: cinta ? null : s.linea, marcador: s.marcador, tipoMuestra: conLinea ? 'linea' : grupo.tipo === 'dispersion' ? 'marcador' : 'caja', serie: s, grupo, larga: (conLinea || grupo.tipo === 'dispersion') && !!s.linea });
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
    g.escalaSuperficie = null;
    let geo = maquetar(g, W, H, med);
    // Superficies: la leyenda lleva las bandas de la escala final, y la escala
    // depende del diseño: una segunda pasada con la de la primera.
    // (la leyenda cambia de alto con el número de bandas: hasta que no cambie, 4 vueltas como mucho).
    for (let v = 0; geo.escalaSuperficie && v < 4; v++) {
      const e0 = g.escalaSuperficie, e1 = geo.escalaSuperficie;
      if (e0 && e0.min === e1.min && e0.max === e1.max && e0.unidad === e1.unidad) break;
      g.escalaSuperficie = e1; geo = maquetar(g, W, H, med);
    }
    if (geo.escalaSuperficie) g.escalaSuperficie = geo.escalaSuperficie;
    const k = pxPorPt;
    ctx.save();
    try {
      ctx.scale(k, k);
      // Nada sale del marco del gráfico (un máximo fijo menor que los datos).
      ctx.beginPath(); ctx.rect(0, 0, W, H); ctx.clip();
      const rect = (x, y, w, h) => { const p = new Path2D(); p.rect(x, y, w, h); return p; };
      // (x, y): la esquina de la caja del relleno. Un degradado o una imagen
      // se acomodan a la caja que empieza ahí (sin ella, en la esquina del gráfico).
      const rellenar = (rell, path, w, h, x = 0, y = 0) => {
        if (!rell) return;
        ctx.save(); ctx.translate(x, y); ctx.scale(1 / k, 1 / k);
        const p2 = new Path2D(); p2.addPath(path, new DOMMatrix().scale(k, k).translate(-x, -y));
        pintar.relleno(ctx, rell, p2, w * k, h * k);
        ctx.restore();
      };
      const trazar = (linea, path) => { if (linea) { ctx.save(); ctx.scale(1 / k, 1 / k); const p2 = new Path2D(); p2.addPath(path, new DOMMatrix().scale(k, k)); pintar.linea(ctx, linea, p2); ctx.restore(); } };

      // Fondo del gráfico y del trazado.
      rellenar(g.fondo.relleno, rect(0, 0, W, H), W, H);
      // El borde de un gráfico de Excel va por dentro del marco; el de PowerPoint, centrado en él.
      const a = g.clasico && g.fondo.linea ? (g.fondo.linea.ancho || 9525) / 12700 / 2 : 0;
      trazar(g.fondo.linea, rect(a, a, W - 2 * a, H - 2 * a));
      const t = geo.trazado;
      if (!geo.circular) { rellenar(g.fondoTrazado.relleno, rect(t.x, t.y, t.w, t.h), t.w, t.h, t.x, t.y); trazar(g.fondoTrazado.linea, rect(t.x, t.y, t.w, t.h)); }

      const salida = { trazado: { ...t }, ejes: [], series: [] };
      if (geo.titulo) salida.titulo = { x: geo.titulo.x, y: geo.titulo.y, w: geo.titulo.w, h: geo.titulo.h };
      if (geo.leyenda) salida.leyenda = { x: geo.leyenda.x, y: geo.leyenda.y, w: geo.leyenda.w, h: geo.leyenda.h };

      if (geo.circular3D) dibujarCircular3D(ctx, g, geo, rellenar, trazar, salida);
      else if (geo.deCircular) dibujarDeCircular(ctx, g, geo, rellenar, trazar, salida);
      else if (geo.circular) dibujarCircular(ctx, g, geo, med, rellenar, trazar, salida);
      else if (geo.radar) dibujarRadar(ctx, g, geo, med, rellenar, trazar, salida);
      else if (geo.escena3D) dibujarEscena3D(ctx, g, geo, med, rellenar, trazar, salida);
      else dibujarEjesYSeries(ctx, g, geo, med, rellenar, trazar, salida);

      if (geo.titulo) {
        const ti = geo.titulo, l = ti.letra, lh = med.linea(l);
        rellenar(g.titulo.relleno, rect(ti.x, ti.y, ti.w, ti.h), ti.w, ti.h, ti.x, ti.y);
        trazar(g.titulo.linea, rect(ti.x, ti.y, ti.w, ti.h));
        ti.lineas.forEach((texto, i) => escribir(ctx, texto, l, ti.x + ti.w / 2, ti.y + PAD_TEXTO.y + i * lh + med.ascendente(l), 'center'));
      }
      if (geo.leyenda) dibujarLeyenda(ctx, g, geo, med, rellenar, trazar);
      for (const c of geo.titulosEje || []) {
        const ti = c.titulo, l = ti.letra, lh = med.linea(l);
        rellenar(ti.relleno, rect(c.x, c.y, c.w, c.h), c.w, c.h, c.x, c.y);
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
    ctx.fontKerning = PptxTexto.kerningDe(l, l.tamano);
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
    // Paredes y piso de un 3D (debajo de todo): la pared del fondo, la de la
    // izquierda y la de abajo (en las columnas, el piso; en las barras, el
    // piso es la de la izquierda). Cada una es una placa (una caja de
    // PptxGrafico3D) con su grosor por fuera; con grosor 0, solo la cara de
    // adentro. La luz de PowerPoint (medido, corpus 22 y 25): de frente tal
    // cual, la de arriba × 0,75 y la de costado × 0,64. [calibrar] el piso
    // cuando el eje cruza en el medio.
    if (c3 && g.paredes) {
      const cam = PptxGrafico3D.oblicua(c3.sx, c3.sy);
      const { izq, abajo, fondo } = c3.grosor;
      const x0 = t.x, x1 = t.x + t.w, yb = -(t.y + t.h), yt = -t.y, D = c3.D;
      const pIzq = c3.acostada ? g.paredes.piso : g.paredes.lado, pAbajo = c3.acostada ? g.paredes.lado : g.paredes.piso;
      // [pared, x, y, z (de / a)]
      const placas = [
        [g.paredes.fondo, x0 - izq, x1, yb, yt, D, D + fondo],
        [pIzq, x0 - izq, x0, yb, yt, 0, D + fondo],
        [pAbajo, x0 - izq, x1, yb - abajo, yb, 0, D + fondo],
      ];
      for (const [cara, xa, xb, ya, yb2, za, zb] of placas) {
        if (!cara || (!cara.relleno && !cara.linea)) continue;
        const centro = (xa + xb) / 2, zc = (za + zb) / 2;
        const sol = PptxGrafico3D.solido({ forma: 'box', eje: 'v', c0: [centro, ya, zc], c1: [centro, yb2, zc], r0: (xb - xa) / 2, r1: (xb - xa) / 2, q0: (zb - za) / 2, q1: (zb - za) / 2 }, cam);
        for (const { puntos: pts, luz } of sol.caras) {
          const xs = pts.map((q) => q[0]), ys = pts.map((q) => q[1]), mx = Math.min(...xs), my = Math.min(...ys);
          // Las caras de canto de una placa sin grosor no se ven.
          if (Math.max(...xs) - mx < 1e-6 || Math.max(...ys) - my < 1e-6) continue;
          const path = new Path2D(); pts.forEach(([x, y], j) => (j ? path.lineTo(x, y) : path.moveTo(x, y))); path.closePath();
          ctx.save(); ctx.translate(mx, my);
          const local = new Path2D(); local.addPath(path, new DOMMatrix().translate(-mx, -my));
          rellenar(luz === 1 ? cara.relleno : sombra(cara.relleno, luz), local, Math.max(...xs) - mx, Math.max(...ys) - my);
          ctx.restore();
          trazar(cara.linea, path);
        }
      }
    }
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
      if (grupo.tipo === 'barra' && c3 && grupo.es3D) dibujarBarras3D(ctx, grupo, geoGrupo, PptxGrafico3D.oblicua(c3.sx, c3.sy), (k) => c3.a + (c3.filas > 1 ? k * c3.fila : 0) + c3.d / 2, c3.d / 2, rellenar, trazar);
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
          // Ya pintadas antes, todas juntas (dibujarBarras3D).
        } else if (grupo.tipo === 'barra') {
          sg.puntos.forEach((p, i) => {
            if (!p) return;
            const pt = s.puntos.get(i) || {};
            let rell = pt.relleno !== undefined && pt.relleno !== null ? pt.relleno : s.relleno;
            const lin = pt.linea !== undefined ? pt.linea : s.linea;
            if (rell && rell.tipo === 'imagen' && /^stack/.test(pt.imagen || s.imagen)) {
              // Apilada desde la base de la barra (medido en 24-graficos-detalles).
              const alReves = ((s.valores[i] ?? 0) < 0) !== !!ey.invertido;
              rell = { ...rell, apilar: grupo.dir === 'bar' ? { dir: 'h', desde: alReves ? 'derecha' : 'izquierda' } : { dir: 'v', desde: alReves ? 'arriba' : 'abajo' } };
            }
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
      // Barras arriba-abajo: de la primera serie a la última en cada
      // categoría, del ancho de la categoría / (1 + gapWidth), encima de las
      // líneas y los marcadores; con los dos valores iguales, una de bajada
      // de alto cero (medido en imágenes).
      const aa = grupo.arribaAbajo;
      if (aa && geoGrupo.length > 1) {
        const primera = geoGrupo[0], ultima = geoGrupo[geoGrupo.length - 1];
        const xs = primera.puntos.filter(Boolean).map((p) => p.x);
        const paso = xs.length > 1 ? Math.abs(xs[1] - xs[0]) : t.w;
        const ancho = paso / (1 + aa.separacion);
        primera.puntos.forEach((p, i) => {
          const q = ultima.puntos[i];
          if (!p || !q) return;
          const sube = (ultima.serie.valores[i] ?? 0) > (primera.serie.valores[i] ?? 0);
          const b = sube ? aa.arriba : aa.abajo;
          const y = Math.min(p.y, q.y), h = Math.abs(p.y - q.y);
          const path = new Path2D(); path.rect(p.x - ancho / 2, y, ancho, h);
          ctx.save(); ctx.translate(p.x - ancho / 2, y);
          const local = new Path2D(); local.rect(0, 0, ancho, h);
          rellenar(b.relleno, local, ancho, h);
          ctx.restore();
          trazar(b.linea, path);
        });
      }
      ctx.restore();
      // Etiquetas de datos.
      for (const sg of geoGrupo) etiquetasDeSerie(ctx, g, grupo, sg, med, t, salida);
    }

    // Ejes: línea y etiquetas.
    if (g.tabla) dibujarTabla(ctx, g, geo, med, rellenar, trazar);
    // En 3D, por fuera del grosor de la pared de la izquierda y de la de abajo.
    const gIzq = c3 ? c3.grosor.izq : 0, gAbajo = c3 ? c3.grosor.abajo : 0;
    for (const { eje: e, etiquetas, adentro, enBordes, disp } of geo.ejes) {
      const lh = med.linea(e.letra), asc = med.ascendente(e.letra);
      // Eje de series de un 3D: cada nombre a la derecha del piso, a la
      // altura del medio de su fila. [calibrar] el aire (un caso medido).
      if (e.clase === 'serAx') {
        if (c3 && c3.filas > 1) etiquetas.forEach((x, k) => {
          const zc = (k + 0.5) * c3.fila;
          escribir(ctx, x.texto, e.letra, t.x + t.w + zc * c3.sx + AIRE_SERIES_3D, t.y + t.h + gAbajo - zc * c3.sy - lh / 2 + asc, 'left');
        });
        continue;
      }
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
          if (e.pos === 'l') escribir(ctx, x.texto, e.letra, borde(t.x - gIzq) - (g.clasico ? AIRE_VALORES_EXCEL : AIRE_VALORES) + CORRE_VALORES, q - lh / 2 + asc, 'right');
          else if (e.pos === 'r') escribir(ctx, x.texto, e.letra, borde(t.x + t.w) + (g.clasico ? AIRE_VALORES_EXCEL : AIRE_VALORES) - CORRE_VALORES, q - lh / 2 + asc, 'left');
          else if (e.pos === 'b') escribir(ctx, x.texto, e.letra, q, borde(t.y + t.h + gAbajo) + BAJA_ETIQUETAS + asc, 'center');
          else escribir(ctx, x.texto, e.letra, q, borde(t.y) - BAJA_ETIQUETAS - lh + asc, 'center');
        }
      } else if (!(g.tabla && e.pos === 'b')) {
        const n = Math.max(1, etiquetas.length);
        const vertical = e.pos === 'l' || e.pos === 'r';
        const paso = (vertical ? t.h : t.w) / (enBordes ? Math.max(1, n - 1) : n);
        const desplazamiento = enBordes ? 0 : 0.5;
        const salto = disp ? disp.salto : e.salto || 1;
        // Junto a la línea del eje si cruza en el medio; si no, afuera del trazado.
        const bordeY = adentro && donde != null ? donde : t.y + t.h + gAbajo;
        const bordeX = adentro && donde != null ? donde : e.pos === 'l' ? t.x - gIzq : t.x + t.w;
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
        // Marcas principales del eje (medido en pantalla en Excel: 4,5 px, con el trazo del eje; hacia afuera, hacia adentro o las dos).
        if (!c3 && (e.marca === 'out' || e.marca === 'in' || e.marca === 'cross')) {
          const horizontal = e.pos === 'b' || e.pos === 't';
          const lineaEje = horizontal ? (donde ?? (e.pos === 'b' ? t.y + t.h : t.y)) : (donde ?? (e.pos === 'l' ? t.x : t.x + t.w));
          const afuera = e.pos === 'b' || e.pos === 'r' ? 1 : -1;
          const lados = e.marca === 'out' ? [afuera] : e.marca === 'in' ? [-afuera] : [afuera, -afuera];
          let lugares = [];
          if (claseEje(e) === 'valor') { const sc = escalas.get(e.id); if (sc) lugares = marcas(sc).map((v) => aPos(e, v)); }
          else {
            const n = Math.max(1, etiquetas.length), nm = enBordes ? Math.max(1, n - 1) : n, total = enBordes ? nm + 1 : n + 1;
            for (let i = 0; i < total; i++) { const u = (horizontal ? t.w : t.h) * i / nm; lugares.push(horizontal ? (e.invertido ? t.x + t.w - u : t.x + u) : (e.invertido ? t.y + u : t.y + t.h - u)); }
          }
          const pm = new Path2D();
          for (const u of lugares) for (const l of lados) { if (horizontal) { pm.moveTo(u, lineaEje); pm.lineTo(u, lineaEje + l * MARCA_EJE); } else { pm.moveTo(lineaEje, u); pm.lineTo(lineaEje + l * MARCA_EJE, u); } }
          trazar(e.linea, pm);
        }
      }
    }
    for (const e of geo.todosLosEjes) {
      const s = escalas.get(e.id);
      const grupoEje = g.grupos.findIndex((x) => x.ejes.includes(e.id));
      const tipoCom = (claseEje(e) === 'valor' && !(g.grupos[grupoEje] && ['dispersion', 'burbuja'].includes(g.grupos[grupoEje].tipo) && (e.pos === 'b' || e.pos === 't'))) ? 2 : 1;
      const secundario = grupoEje > 0 && g.grupos[grupoEje].ejes.some((id) => !g.grupos[0].ejes.includes(id));
      // Eje de fechas: de la primera a la última fecha y, como unidad
      // principal, el paso de las etiquetas en unidades base (como COM).
      let fechas = null;
      if (claseEje(e) === 'fecha') {
        const grupo = g.grupos[grupoEje], serie = grupo && grupo.series.find((x) => x.categorias && typeof x.categorias[0] === 'number');
        const nums = serie ? serie.categorias.filter((x) => typeof x === 'number') : [];
        const ge = geo.ejes.find((x) => x.eje === e);
        if (nums.length) fechas = { min: Math.min(...nums), max: Math.max(...nums), unidad: (ge && ge.disp && ge.disp.salto) || 1 };
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
    const F = fechasDe(g, ex);
    const n = F ? F.n : Math.max(1, ...grupo.series.map((s) => s.valores.length));
    const largo = horizontal ? t.h : t.w;
    const cw = largo / n;
    const apilado = grupo.agrupacion === 'stacked' || grupo.agrupacion === 'percentStacked';
    // 3D con las series en filas de profundidad: una barra por categoría.
    // (Las acostadas no van en filas: igual que caja3D y escena3D.)
    const enFilas = enFilas3D(grupo) && grupo.dir !== 'bar';
    const ns = apilado || enFilas ? 1 : grupo.series.length;
    const ov = apilado || enFilas ? 1 : grupo.solapamiento;
    const bw = cw / (ns - (ns - 1) * ov + grupo.separacion);
    const sy = escalas.get(ey.id);
    const cero = Math.min(sy.max, Math.max(sy.min, 0));
    const nv = Math.max(1, ...grupo.series.map((s) => s.valores.length));
    const totales = [];
    if (grupo.agrupacion === 'percentStacked') for (let i = 0; i < nv; i++) totales[i] = grupo.series.reduce((a, s) => a + Math.abs(s.valores[i] || 0), 0) || 1;
    const acumPos = new Array(nv).fill(0), acumNeg = new Array(nv).fill(0);
    // 3D: dónde termina la pila de cada categoría (la punta de un cono apilado)
    // y el máximo y el mínimo del eje (la de los "hasta el máximo").
    let cimaPos = null, cimaNeg = null;
    if (grupo.es3D) {
      cimaPos = new Array(nv).fill(0); cimaNeg = new Array(nv).fill(0);
      for (const s of grupo.series) s.valores.forEach((v, i) => {
        if (v == null) return;
        const vv = grupo.agrupacion === 'percentStacked' ? v / totales[i] : v;
        if (vv >= 0) cimaPos[i] += vv; else cimaNeg[i] += vv;
      });
    }
    return grupo.series.map((s, k) => {
      const puntos = s.valores.map((v, i) => {
        if (v == null) return null;
        let desde = cero, hasta = v;
        if (apilado) {
          const vv = grupo.agrupacion === 'percentStacked' ? v / totales[i] : v;
          if (vv >= 0) { desde = acumPos[i]; hasta = desde + vv; acumPos[i] = hasta; } else { desde = acumNeg[i]; hasta = desde + vv; acumNeg[i] = hasta; }
        }
        const j = F ? F.lugares[i] : i;
        if (j == null) return null;
        const inicioCat = cw * j + bw * grupo.separacion / 2 + (apilado || enFilas ? 0 : k * bw * (1 - ov));
        const p0 = aPos(ey, desde), p1 = aPos(ey, hasta);
        // base y tope: dónde empieza y termina sobre el eje de valores; cima y
        // cimaMax: la punta de un cono apilado y la de uno hasta el máximo.
        const extremos = { base: p0, tope: p1 };
        if (grupo.es3D) {
          const arriba = hasta >= desde;
          extremos.cero = aPos(ey, cero);
          extremos.cima = apilado ? aPos(ey, arriba ? cimaPos[i] : cimaNeg[i]) : p1;
          extremos.cimaMax = arriba ? aPos(ey, Math.max(sy.max, hasta)) : aPos(ey, Math.min(sy.min, hasta));
        }
        if (horizontal) {
          const y = (ex.invertido ? t.y + inicioCat : t.y + t.h - inicioCat - bw);
          return { x: Math.min(p0, p1), y, w: Math.abs(p1 - p0), h: bw, v, positivo: hasta >= desde, horizontal, k, i: j, ...extremos };
        }
        const x = ex.invertido ? t.x + t.w - inicioCat - bw : t.x + inicioCat;
        return { x, y: Math.min(p0, p1), w: bw, h: Math.abs(p1 - p0), v, positivo: hasta >= desde, k, i: j, ...extremos };
      });
      return { serie: s, puntos };
    });
  }
  function lineasOAreas(g, grupo, ex, ey, t, aPos, escalas) {
    const n = Math.max(1, ...grupo.series.map((s) => s.valores.length));
    const F = fechasDe(g, ex), nx = F ? F.n : n;
    // crossBetween (en el eje de valores): "between" pone cada punto en el
    // centro de su categoría; "midCat", en los bordes.
    const entre = ey.entre !== 'midCat';
    const xDe = (i) => {
      const j = F ? F.lugares[i] ?? i : i;
      const u = entre ? (j + 0.5) / nx : nx > 1 ? j / (nx - 1) : 0.5;
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

  // Un relleno liso oscurecido o aclarado (las caras de los 3D); los demás, igual.
  function sombra(r, f) {
    const k = (x) => Math.min(255, Math.round(x * f));
    return r && r.tipo === 'solido' ? { tipo: 'solido', color: { r: k(r.color.r), g: k(r.color.g), b: k(r.color.b), a: r.color.a } } : r;
  }

  // Barras 3D: cada punto, un sólido de PptxGrafico3D (caja, cilindro, cono o
  // pirámide), y todos los del grupo de atrás hacia adelante. zDe(serie) y q:
  // centro y medio ancho en profundidad. Las puntas (medido en
  // 25-graficos-3d): el cono y la pirámide terminan en el valor; apilados, la
  // pila entera es un solo cono que termina en el total; "ToMax", uno que
  // termina en el máximo del eje y se corta en el valor.
  // Un gráfico 3D en escena (perspectiva, o líneas y áreas 3D): paredes y
  // piso del lado lejano, cuadrícula de valores sobre ellas, las series de
  // atrás hacia adelante (barras: sólidos; líneas: cintas; áreas: placas con
  // el contorno de los datos) y las etiquetas de los ejes junto a los bordes
  // de adelante. [calibrar] el lugar de las etiquetas y el grosor de las
  // paredes en perspectiva; las etiquetas de datos no se dibujan.
  function dibujarEscena3D(ctx, g, geo, med, rellenar, trazar, salida) {
    const e3 = geo.escena3D, { cam, W, H, D, zF, zN, xL, horizontal, grupo } = e3;
    const escalas = geo.escalas;
    const ex = g.ejes.get(grupo.ejes[0]), ey = g.ejes.get(grupo.ejes[1]);
    const sy = ey && escalas.get(ey.id);
    // Valor → coordenada del mundo (y en las de pie, x en las acostadas).
    const aMundo = (v) => { if (!sy) return 0; let u = Math.max(0, Math.min(1, fraccion(sy, v))); if (ey.invertido) u = 1 - u; return (horizontal ? W : H) * u; };
    const poligono = (pts3, rell, linea, luz = 1) => {
      const pts = pts3.map(cam.p);
      const xs = pts.map((q) => q[0]), ys = pts.map((q) => q[1]), mx = Math.min(...xs), my = Math.min(...ys);
      const path = new Path2D(); pts.forEach(([x, y], j) => (j ? path.lineTo(x, y) : path.moveTo(x, y))); path.closePath();
      if (rell) {
        ctx.save(); ctx.translate(mx, my);
        const local = new Path2D(); local.addPath(path, new DOMMatrix().translate(-mx, -my));
        rellenar(luz === 1 ? rell : sombra(rell, luz), local, Math.max(...xs) - mx || 1, Math.max(...ys) - my || 1);
        ctx.restore();
      }
      trazar(linea, path);
    };
    const linea3 = (pts3, ln) => { if (!ln) return; const p = new Path2D(); pts3.map(cam.p).forEach(([x, y], j) => (j ? p.lineTo(x, y) : p.moveTo(x, y))); trazar(ln, p); };
    const xO = W - xL; // el borde de adelante del lado contrario a la pared
    // Paredes: la del fondo, la de costado y la de abajo (en las barras
    // acostadas, el piso es la pared de los valores en cero).
    const pa = g.paredes || {};
    const caras = [
      [pa.fondo, [[0, 0, zF], [W, 0, zF], [W, H, zF], [0, H, zF]], [0, 0, zF ? -1 : 1]],
      [horizontal ? pa.piso : pa.lado, [[xL, 0, 0], [xL, 0, D], [xL, H, D], [xL, H, 0]], [xL ? -1 : 1, 0, 0]],
      [horizontal ? pa.lado : pa.piso, [[0, 0, 0], [W, 0, 0], [W, 0, D], [0, 0, D]], [0, 1, 0]],
    ];
    for (const [cara, pts3, n] of caras) if (cara && (cara.relleno || cara.linea)) poligono(pts3, cara.relleno, cara.linea, PptxGrafico3D.luz(n));
    // Cuadrícula de valores: por la pared de costado (o el piso) y la del fondo.
    if (ey && ey.cuadricula && sy) for (const v of marcas(sy)) {
      const q = aMundo(v);
      if (horizontal) linea3([[q, 0, zN], [q, 0, zF], [q, H, zF]], ey.cuadricula);
      else linea3([[xL, q, zN], [xL, q, zF], [xO, q, zF]], ey.cuadricula);
    }

    // Series.
    const ox0 = (i) => (e3.bordes ? (e3.n > 1 ? i * W / (e3.n - 1) : W / 2) : i + 0.5);
    const ox = (i) => (ex && ex.invertido ? W - ox0(i) : ox0(i)); // x de la categoría i
    const zFila = (k) => (e3.zSerie ? e3.zSerie(k) : e3.filas > 1 ? (k + 0.5) * e3.fila : D / 2);
    const series = grupo.series;
    salida.series.push(...series.map((s) => ({ nombre: s.nombre, puntos: s.valores.map(() => ({})) })));
    if (grupo.tipo === 'superficie') {
      // La malla en el mundo; las caras de lejos a cerca (por el centro), o el alambre.
      const z = series.map((s) => s.valores);
      const niveles = sy ? marcas(sy) : [];
      const bandas = bandasDe(g, grupo);
      const mundo = { X: (i) => ox(i), Z: (j) => zFila(j), Y: aMundo };
      const lejos = (pts) => cam.prof(pts.reduce((a, q) => [a[0] + q[0] / pts.length, a[1] + q[1] / pts.length, a[2] + q[2] / pts.length], [0, 0, 0]));
      if (grupo.alambre) {
        const segs = PptxSuperficie.alambre(z, niveles, mundo).sort((a, b) => lejos([b.a, b.b]) - lejos([a.a, a.b]));
        for (const sg of segs) if (bandas[sg.banda]) linea3([sg.a, sg.b], bandas[sg.banda].linea);
      } else {
        const caras = PptxSuperficie.facetas(z, niveles, mundo).sort((a, b) => lejos(b.pts) - lejos(a.pts));
        for (const cara of caras) {
          const b = bandas[cara.banda];
          if (!b || !b.relleno) continue;
          // Un trazo del mismo color tapa las costuras entre triángulos.
          const rs = sombra(b.relleno, cara.luz);
          poligono(cara.pts, b.relleno, rs && rs.tipo === 'solido' ? lineaFina(rs.color, 6350) : null, cara.luz);
        }
      }
    } else if (grupo.tipo === 'barra') {
      const tw = { x: 0, y: 0, w: W, h: H };
      // Sin pasar de la caja (un máximo fijo menor que el dato, o ≤ 0 en logarítmica).
      const aPosW = (e, v) => { const u0 = Math.max(0, Math.min(1, fraccion(escalas.get(e.id), v))), u = e.invertido ? 1 - u0 : u0; return e.pos === 'b' || e.pos === 't' ? u * W : H - u * H; };
      const geoGrupo = barras(g, grupo, ex, ey, tw, aPosW, escalas);
      dibujarBarras3D(ctx, grupo, geoGrupo, cam, (k) => zFila(k), e3.d / 2, rellenar, trazar, H);
      geoGrupo.forEach((sg, k) => { salida.series[salida.series.length - series.length + k].puntos = sg.puntos.map((p) => (p ? { caja: { x: p.x, y: p.y, w: p.w, h: p.h } } : {})); });
    } else {
      // Líneas (cintas) y áreas (placas), una fila por serie, de atrás hacia adelante.
      const orden = series.map((s, k) => k).sort((a, b) => cam.lejania(2, zFila(b)) - cam.lejania(2, zFila(a)));
      const apilada = /stacked/i.test(grupo.agrupacion), pct = grupo.agrupacion === 'percentStacked';
      const base = new Array(e3.n).fill(0), arriba = [];
      if (grupo.tipo === 'area' && apilada) {
        const tot = Array.from({ length: e3.n }, (_, i) => series.reduce((a, s) => a + Math.abs(s.valores[i] || 0), 0) || 1);
        for (const s of series) { arriba.push(s.valores.map((v, i) => { const b = base[i]; base[i] += pct ? (v || 0) / tot[i] : v || 0; return [b, base[i]]; })); }
      }
      for (const k of orden) {
        const s = series[k];
        const z0 = zFila(k) - e3.d / 2, z1 = zFila(k) + e3.d / 2;
        const tramos = s.valores.map((v, i) => (v == null ? null : arriba.length ? arriba[k][i] : [0, v]));
        const pts = tramos.map((t, i) => (t ? { x: ox(i), y0: aMundo(t[0]), y1: aMundo(t[1]) } : null));
        // Los tramos entre dos puntos, de lejos a cerca a lo largo de x.
        const segs = [];
        for (let i = 0; i + 1 < pts.length; i++) if (pts[i] && pts[i + 1]) segs.push([pts[i], pts[i + 1]]);
        segs.sort((a, b) => cam.lejania(0, (b[0].x + b[1].x) / 2) - cam.lejania(0, (a[0].x + a[1].x) / 2));
        if (grupo.tipo === 'linea') {
          const color = s.linea && s.linea.relleno ? s.linea.relleno : s.relleno;
          for (const [a, b] of segs) {
            const sup = [[a.x, a.y1, z0], [b.x, b.y1, z0], [b.x, b.y1, z1], [a.x, a.y1, z1]];
            let n = [-(b.y1 - a.y1), b.x - a.x, 0];
            const o = cam.ojo([(a.x + b.x) / 2, (a.y1 + b.y1) / 2, (z0 + z1) / 2]);
            if (n[0] * o[0] + n[1] * o[1] + n[2] * o[2] < 0) n = n.map((x) => -x);
            poligono(sup, color, null, PptxGrafico3D.luz(n));
          }
        } else {
          const vis = (n, c) => { const o = cam.ojo(c); return n[0] * o[0] + n[1] * o[1] + n[2] * o[2] > 1e-9; };
          const cont = (z) => { const lista = pts.filter(Boolean); return [...lista.map((p) => [p.x, p.y1, z]), ...lista.slice().reverse().map((p) => [p.x, p.y0, z])]; };
          const lista = pts.filter(Boolean);
          if (!lista.length) continue;
          const zl = cam.lejania(2, z0) > cam.lejania(2, z1) ? z0 : z1, zc = zl === z0 ? z1 : z0;
          poligono(cont(zl), s.relleno, s.linea, PptxGrafico3D.luz([0, 0, zl === z0 ? -1 : 1]));
          const tapas = [];
          for (const [a, b] of segs) {
            // Hacia afuera de la placa: con valores negativos el borde del dato queda abajo.
            const sg = Math.sign((a.y1 + b.y1) - (a.y0 + b.y0)) || 1, sx = Math.sign(b.x - a.x) || 1;
            tapas.push({ x: (a.x + b.x) / 2, pts3: [[a.x, a.y1, z0], [b.x, b.y1, z0], [b.x, b.y1, z1], [a.x, a.y1, z1]], n: [-(b.y1 - a.y1) * sg * sx, (b.x - a.x) * sg * sx, 0] });
            tapas.push({ x: (a.x + b.x) / 2, pts3: [[a.x, a.y0, z0], [b.x, b.y0, z0], [b.x, b.y0, z1], [a.x, a.y0, z1]], n: [(b.y0 - a.y0) * sg * sx, -(b.x - a.x) * sg * sx, 0] });
          }
          const pi = lista[0], pu = lista[lista.length - 1], sx = Math.sign(pu.x - pi.x) || 1;
          tapas.push({ x: pi.x, pts3: [[pi.x, pi.y0, z0], [pi.x, pi.y1, z0], [pi.x, pi.y1, z1], [pi.x, pi.y0, z1]], n: [-sx, 0, 0] });
          tapas.push({ x: pu.x, pts3: [[pu.x, pu.y0, z0], [pu.x, pu.y1, z0], [pu.x, pu.y1, z1], [pu.x, pu.y0, z1]], n: [sx, 0, 0] });
          tapas.sort((a, b) => cam.lejania(0, b.x) - cam.lejania(0, a.x));
          for (const t of tapas) { const c = t.pts3.reduce((a, q) => [a[0] + q[0] / 4, a[1] + q[1] / 4, a[2] + q[2] / 4], [0, 0, 0]); if (vis(t.n, c)) poligono(t.pts3, s.relleno, s.linea, PptxGrafico3D.luz(t.n)); }
          poligono(cont(zc), s.relleno, s.linea, PptxGrafico3D.luz([0, 0, zc === z0 ? -1 : 1]));
        }
      }
    }

    // Superficies: la línea de los ejes de categorías y de series (en el borde
    // de adelante) con sus marcas hacia afuera. [calibrar] el largo de las marcas.
    if (grupo.tipo === 'superficie') {
      const marca = (P, dx, dy, ln) => { const p = new Path2D(); p.moveTo(P[0], P[1]); p.lineTo(P[0] + dx, P[1] + dy); trazar(ln, p); };
      for (const e of geo.todosLosEjes) {
        if (e.borrado || !e.linea) continue;
        const conMarca = e.marca === 'out' || e.marca === 'cross';
        if (e.clase === 'catAx' || e.clase === 'dateAx') {
          linea3([[0, 0, zN], [W, 0, zN]], e.linea);
          if (conMarca) for (let i = 0; i < e3.n; i++) marca(cam.p([ox0(i), 0, zN]), 0, MARCA_3D, e.linea);
        } else if (e.clase === 'serAx') {
          linea3([[xO, 0, 0], [xO, 0, D]], e.linea);
          if (conMarca) for (let k = 0; k < e3.filas; k++) marca(cam.p([xO, 0, zFila(k)]), (xO ? 1 : -1) * MARCA_3D, 0, e.linea);
        }
      }
    }

    // Ejes: las etiquetas junto a los bordes de adelante.
    for (const { eje: e, etiquetas } of geo.ejes) {
      const lh = med.linea(e.letra), asc = med.ascendente(e.letra);
      const izquierda = (P, texto) => escribir(ctx, texto, e.letra, P[0] - AIRE_VALORES + CORRE_VALORES, P[1] - lh / 2 + asc, 'right');
      const derecha = (P, texto) => escribir(ctx, texto, e.letra, P[0] + AIRE_VALORES - CORRE_VALORES, P[1] - lh / 2 + asc, 'left');
      const abajo = (P, texto) => escribir(ctx, texto, e.letra, P[0], P[1] + BAJA_ETIQUETAS + asc, 'center');
      const lado = xL ? derecha : izquierda;
      if (e.clase === 'serAx') {
        // En las superficies, el eje de series muestra a lo sumo piso(largo /
        // (1,05 × alto de letra)) etiquetas y salta de a las que haga falta
        // (medido en 26-graficos-superficie: 5 series en 65 pt, de a 2; 10 en
        // 83 pt, de a 2; 5 en 142 pt, todas). [calibrar] el 1,05.
        let salto = 1;
        if (grupo.tipo === 'superficie' && e3.filas > 1) {
          const a = cam.p([xO, 0, zFila(0)]), b = cam.p([xO, 0, zFila(e3.filas - 1)]);
          const caben = Math.max(1, Math.floor(Math.hypot(a[0] - b[0], a[1] - b[1]) / (SALTO_SERIES_3D * lh)));
          while (Math.ceil(e3.filas / salto) > caben && salto < e3.filas) salto++;
        }
        const aire = grupo.tipo === 'superficie' && g.rotX >= 80 ? AIRE_SERIES_CONTORNO : AIRE_SERIES_3D;
        if (e3.filas > 1) etiquetas.forEach((x, k) => { if (k % salto) return; const P = cam.p([xO, 0, zFila(k)]); escribir(ctx, x.texto, e.letra, P[0] + (xO ? 1 : -1) * aire, P[1] - lh / 2 + asc, xO ? 'left' : 'right'); });
      } else if (claseEje(e) === 'valor') {
        for (const x of etiquetas) {
          const q = aMundo(x.v);
          if (horizontal) abajo(cam.p([q, 0, zN]), x.texto); else lado(cam.p([xL, q, zN]), x.texto);
        }
      } else {
        // Superficies: las etiquetas del eje de categorías se saltean igual que
        // las de series si el eje, en pantalla, es corto (rotY de 60: de a 2).
        let saltoCat = 1;
        if (grupo.tipo === 'superficie' && !horizontal) {
          const a = cam.p([0, 0, zN]), b = cam.p([W, 0, zN]);
          const caben = Math.max(1, Math.floor(Math.hypot(a[0] - b[0], a[1] - b[1]) / (SALTO_SERIES_3D * lh)));
          // Y que las etiquetas no se pisen: el paso en pantalla, a lo ancho
          // del texto más ancho o a lo alto de una línea.
          const dx = Math.abs(a[0] - b[0]) / Math.max(1, e3.n - 1), dy = Math.abs(a[1] - b[1]) / Math.max(1, e3.n - 1);
          const ancho = Math.max(0, ...etiquetas.map((x) => med.ancho(x.texto, e.letra))) + 4;
          while ((Math.ceil(e3.n / saltoCat) > caben || (saltoCat * dx < ancho && saltoCat * dy < SALTO_SERIES_3D * lh)) && saltoCat < e3.n) saltoCat++;
        }
        etiquetas.forEach((x, i) => {
          if (i >= e3.n || i % saltoCat) return;
          const c = grupo.tipo === 'barra' ? i + 0.5 : ox0(i);
          if (horizontal) lado(cam.p([xL, e.invertido ? H - c : c, zN]), x.texto);
          else abajo(cam.p([e.invertido ? W - c : c, 0, zN]), x.texto);
        });
      }
    }
    // Como COM: la escala de cada eje (sin el de series, que COM no da).
    for (const e of geo.todosLosEjes) {
      if (e.clase === 'serAx') continue;
      const s = escalas.get(e.id);
      salida.ejes.push({ tipo: claseEje(e) === 'valor' ? 2 : 1, grupo: 1, borrado: e.borrado, ...(s ? { min: s.min, max: s.max, unidad: s.unidad } : {}) });
    }
  }

  // Los puntos vienen como en pantalla (y hacia abajo); el mundo tiene y
  // hacia arriba: y del mundo = arribaY − y (0 con la cámara oblicua, el
  // alto de la caja en la escena).
  function dibujarBarras3D(ctx, grupo, geoGrupo, cam, zDe, q, rellenar, trazar, arribaY = 0) {
    const acostada = grupo.dir === 'bar';
    const lista = [];
    for (const sg of geoGrupo) {
      const s = sg.serie;
      sg.puntos.forEach((p, i) => {
        if (!p) return;
        const z = zDe(p.k);
        const pt = s.puntos.get(i) || {};
        const forma = String(s.forma || 'box');
        const punta = /^(cone|pyramid)ToMax$/.test(forma) ? p.cimaMax : /^(cone|pyramid)$/.test(forma) ? p.cima : null;
        const ancho = (acostada ? p.h : p.w) / 2;
        // Medio ancho en un lugar del eje de valores (en pantalla).
        const k = (t) => (punta == null || Math.abs(punta - p.cero) < 1e-9 ? 1 : Math.max(0, (punta - t) / (punta - p.cero)));
        const centro = acostada ? arribaY - (p.y + p.h / 2) : p.x + p.w / 2;
        const en = (t) => (acostada ? [t, centro, z] : [centro, arribaY - t, z]);
        const def = {
          forma: forma.replace(/ToMax$/, ''), eje: acostada ? 'h' : 'v',
          c0: en(p.base), c1: en(p.tope), r0: ancho * k(p.base), r1: ancho * k(p.tope), q0: q * k(p.base), q1: q * k(p.tope),
        };
        const medioValor = (p.base + p.tope) / 2;
        lista.push({
          p, solido: PptxGrafico3D.solido(def, cam),
          rell: pt.relleno !== undefined && pt.relleno !== null ? pt.relleno : s.relleno,
          lin: pt.linea !== undefined ? pt.linea : s.linea,
          celda: acostada ? { z, a: centro, s: medioValor } : { z, a: centro, s: arribaY - medioValor },
        });
      });
    }
    for (const b of PptxGrafico3D.ordenar(lista, cam, acostada ? { a: 1, s: 0 } : { a: 0, s: 1 })) {
      for (const cara of b.solido.caras) {
        const pts = cara.puntos;
        if (pts.length < 3) continue;
        const path = new Path2D(); pts.forEach(([a, c], j) => (j ? path.lineTo(a, c) : path.moveTo(a, c))); path.closePath();
        const xs = pts.map((v) => v[0]), ys = pts.map((v) => v[1]), mx = Math.min(...xs), my = Math.min(...ys);
        const rell = cara.perfil ? PptxGrafico3D.degradadoDe(b.rell, cara.perfil, cara.angulo) || b.rell : sombra(b.rell, cara.luz);
        ctx.save(); ctx.translate(mx, my);
        const local = new Path2D(); local.addPath(path, new DOMMatrix().translate(-mx, -my));
        rellenar(rell, local, Math.max(...xs) - mx, Math.max(...ys) - my);
        ctx.restore();
        trazar(b.lin, path);
      }
      // Como COM: la caja de todo el sólido.
      Object.assign(b.p, b.solido.caja);
    }
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
    const vals = s.valores.map((v) => (v != null ? Math.abs(v) : 0)); // un negativo, como su valor absoluto (medido)
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

  // Circular o barra de circular (ofPieChart; medido en imágenes,
  // sondeo-de-circular.ps1). Los puntos que van al segundo (los últimos
  // splitPos, o sin splitType el último tercio redondeado arriba (COM: 6 → 2,
  // 8 → 3, 10 → 4), por valor, por porcentaje o elegidos) se juntan en el primero
  // en un sector "otros", centrado a la derecha y del color del punto n (el
  // que sigue al último). Los dos empiezan en el borde de abajo de "otros" y
  // siguen en el orden de los datos; en la barra, el primero va arriba. Las
  // líneas de serie van de los bordes de "otros" a las tangentes del segundo
  // circular o a las esquinas de la barra.
  function puntosDelSegundo(de, vals) {
    const n = vals.length, total = vals.reduce((a, b) => a + b, 0) || 1;
    const lista = [];
    vals.forEach((v, i) => {
      if (de.corte === 'val' ? v < (de.pos ?? 0) : de.corte === 'percent' ? v / total * 100 < (de.pos ?? 0) : de.corte === 'cust' ? de.propios.includes(i) : i >= n - Math.min(n - 1, de.pos ?? Math.ceil(n / 3))) lista.push(i);
    });
    return lista;
  }
  function dibujarDeCircular(ctx, g, geo, rellenar, trazar, salida) {
    const grupo = g.grupos[0], s = grupo.series[0], de = grupo.de;
    const { cx, cy, R, x2, R2 } = geo.deCircular;
    const vals = s.valores.map((v) => (v != null ? Math.abs(v) : 0)); // un negativo, como su valor absoluto (medido)
    const total = vals.reduce((a, b) => a + b, 0) || 1;
    const segundo = puntosDelSegundo(de, vals);
    const otros = segundo.reduce((a, i) => a + vals[i], 0);
    const relleno = (i) => { const pt = s.puntos.get(i) || {}; return pt.relleno !== undefined && pt.relleno !== null ? pt.relleno : colorPunto(g, s, i); };
    const lineaDe = (i) => { const pt = s.puntos.get(i) || {}; return pt.linea !== undefined ? pt.linea : s.linea; };
    const pintar = (path, rell, caja) => {
      ctx.save(); ctx.translate(caja.x, caja.y);
      const local = new Path2D(); local.addPath(path, new DOMMatrix().translate(-caja.x, -caja.y));
      rellenar(rell, local, caja.w, caja.h);
      ctx.restore();
    };
    const cajas = [];
    const sector = (ox, oy, r, a0, a1, i) => {
      const p = new Path2D(); p.moveTo(ox, oy); p.arc(ox, oy, r, a0, a1); p.closePath();
      pintar(p, relleno(i), { x: ox - r, y: oy - r, w: 2 * r, h: 2 * r });
      trazar(lineaDe(i), p);
      cajas[i] = cajaSector(ox, oy, r, 0, a0, a1);
    };
    const alfa = otros / total * Math.PI;
    let ang = alfa;
    vals.forEach((v, i) => { if (segundo.includes(i)) return; const da = v / total * Math.PI * 2; sector(cx, cy, R, ang, ang + da, i); ang += da; });
    if (otros > 0) sector(cx, cy, R, ang, alfa + Math.PI * 2, vals.length);
    const arriba = [cx + R * Math.cos(alfa), cy - R * Math.sin(alfa)], abajo = [cx + R * Math.cos(alfa), cy + R * Math.sin(alfa)];
    const lineas = new Path2D();
    if (de.barra) {
      const bw = R2, bh = 2 * R2, by = cy - R2;
      let y = by;
      for (const i of segundo) {
        const h = otros > 0 ? vals[i] / otros * bh : 0;
        const p = new Path2D(); p.rect(x2, y, bw, h);
        pintar(p, relleno(i), { x: x2, y, w: bw, h });
        trazar(lineaDe(i), p);
        cajas[i] = { x: x2, y, w: bw, h };
        y += h;
      }
      lineas.moveTo(...arriba); lineas.lineTo(x2, by); lineas.moveTo(...abajo); lineas.lineTo(x2, by + bh);
    } else {
      const c2 = x2 + R2;
      ang = alfa;
      for (const i of segundo) { const da = otros > 0 ? vals[i] / otros * Math.PI * 2 : 0; sector(c2, cy, R2, ang, ang + da, i); ang += da; }
      // Tangente al segundo desde cada borde de "otros" (la de arriba y la de abajo).
      for (const [px, py, lado] of [[...arriba, -1], [...abajo, 1]]) {
        const dx = c2 - px, dy = cy - py, d = Math.hypot(dx, dy);
        if (d <= R2) continue;
        const th = Math.atan2(dy, dx) + lado * Math.asin(R2 / d), l = Math.sqrt(d * d - R2 * R2);
        lineas.moveTo(px, py); lineas.lineTo(px + l * Math.cos(th), py + l * Math.sin(th));
      }
    }
    if (de.lineas && segundo.length) trazar(de.lineas, lineas);
    salida.series.push({ nombre: s.nombre, puntos: vals.map((_, i) => (cajas[i] ? { caja: cajas[i] } : {})) });
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
    // La separación de la serie (c:ser/c:explosion) achica el círculo para
    // que entre; la de un sector suelto, no (medido: con etiquetas, todos al
    // 20 % → radio / 1,2; uno solo al 12 % → el mismo radio).
    const R = (t.w / 2) / (1 + (s0.explosion || 0));
    const hueco = anillo ? R * Math.max(0.1, Math.min(0.9, grupo.hueco || 0.5)) : 0;
    const grosor = anillo ? (R - hueco) / series.length : R;
    series.forEach((s, k) => {
      const rIn = anillo ? hueco + k * grosor : 0, rOut = anillo ? rIn + grosor : R;
      const vals = s.valores.map((v) => (v != null ? Math.abs(v) : 0)); // un negativo, como su valor absoluto (medido)
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
        registro.puntos.push({ caja: cajaSector(cx + dx, cy + dy, rOut, rIn, ang, ang + a), _medio: medio, _abre: a, _expl: expl, _v: s.valores[i], _pct: vals[i] / total });
        ang += a;
      });
      salida.series.push(registro);
      const e = s.etiquetas;
      if (e && !e.borrar && (e.showVal || e.showCatName || e.showPercent || e.showSerName)) {
        const l = e.letra, lh = med.linea(l), asc = med.ascendente(l);
        const adentro = [];
        const cajas = [];
        registro.puntos.forEach((p, i) => {
          const propia = e.propias && e.propias.get(i);
          if (propia && propia.borrar) return;
          const partes = [];
          if (e.showSerName) partes.push(s.nombre);
          if (e.showCatName) partes.push(s.categorias ? String(s.categorias[i] ?? '') : String(i + 1));
          if (e.showVal) partes.push(formatear(p._v, e.formato || s.formato));
          if (e.showPercent) partes.push(formatear(p._pct, '0%'));
          // Sin separador propio: nombre y porcentaje en dos líneas; lo demás, con "; " (medido).
          const sepTexto = e.separador || (e.showCatName && e.showPercent && !e.showVal ? '\n' : '; ');
          const lineas = partes.join(sepTexto).split('\n');
          const w = Math.max(...lineas.map((x) => med.ancho(x, l))) + 2 * PAD_TEXTO.x, h = lineas.length * lh + 2 * PAD_TEXTO.y;
          const pos = e.pos || (anillo ? 'ctr' : 'bestFit');
          const c = Math.cos(p._medio), sn = Math.sin(p._medio), ex = c * p._expl * R, ey = sn * p._expl * R;
          // Medido en imágenes y por COM (sondeo-etiquetas-circular.ps1;
          // 06-graficos.md 6.4 quinquies). Lo que la etiqueta ocupa hacia el centro:
          const apoyo = w / 2 * Math.abs(c) + h / 2 * Math.abs(sn);
          let lx, ly, fuera = false;
          // Extremo externo, en el ángulo th: un punto de apoyo a 1,044 R −
          // 0,32; hasta 66° de la horizontal la etiqueta se apoya por su
          // costado (a 1 pt) y se corre en vertical con el ángulo; más arriba
          // o abajo, por su borde de arriba o de abajo. Nunca sale de la forma
          // del gráfico.
          const enAngulo = (th) => {
            const c1 = Math.cos(th), s1 = Math.sin(th);
            const ta = rOut * 1.0438 - 0.32, ax = cx + ex + c1 * ta, ay = cy + ey + s1 * ta;
            const dh = Math.atan2(Math.abs(s1), Math.abs(c1)) * 180 / Math.PI, sx = c1 >= 0 ? 1 : -1, sy = s1 >= 0 ? 1 : -1;
            let x, y;
            if (dh < 65.7) { x = ax + sx * (w / 2 + 1.02); y = ay + sy * Math.min(0.5, dh / 98.9) * h; }
            else { x = ax + sx * Math.min(0.5, (90 - dh) / 35.24) * w; y = ay + sy * h / 2; }
            return { x: Math.max(w / 2, Math.min(geo.W - w / 2, x)), y: Math.max(h / 2, Math.min(geo.H - h / 2, y)) };
          };
          const afuera = () => { fuera = true; ({ x: lx, y: ly } = enAngulo(p._medio)); };
          const radial = (r) => { lx = cx + ex + c * r; ly = cy + ey + sn * r; };
          if (anillo || pos === 'ctr') radial((rOut + rIn) / 2);
          else if (pos === 'inEnd') radial(rOut - 3.06 - apoyo);
          else if (pos === 'bestFit') {
            // Adentro, cerca del borde, si entra entera en su sector y no
            // choca con una que ya quedó adentro (se reparten en orden); si
            // no, afuera (24-graficos-detalles: los sectores chicos).
            radial(rOut - 1.5 - apoyo);
            const caja = { x: lx - w / 2, y: ly - h / 2, w, h };
            const ox = cx + ex, oy = cy + ey;
            const entra = [[caja.x, caja.y], [caja.x + w, caja.y], [caja.x, caja.y + h], [caja.x + w, caja.y + h]].every(([qx, qy]) => {
              const d = Math.atan2(qy - oy, qx - ox) - p._medio;
              // Las esquinas de afuera pueden pasar el borde curvo un poco (medio alto).
              return Math.hypot(qx - ox, qy - oy) <= rOut + h / 2 && Math.abs(Math.atan2(Math.sin(d), Math.cos(d))) <= p._abre / 2;
            });
            if (!entra || adentro.some((o) => caja.x < o.x + o.w && o.x < caja.x + caja.w && caja.y < o.y + o.h && o.y < caja.y + caja.h)) afuera();
            else adentro.push(caja);
          } else afuera();
          // Movida a mano: corrida desde donde iría sola, en fracciones del gráfico.
          const movida = !!(propia && propia.diseno);
          if (movida) {
            const c2 = aplicarDiseno({ x: lx - w / 2, y: ly - h / 2, w, h }, { ...propia.diseno, w: null, h: null }, geo.W, geo.H);
            lx = c2.x + w / 2; ly = c2.y + h / 2;
          }
          cajas.push({ p, lineas, w, h, lx, ly, fuera: fuera && !movida, movida, ancla: { x: cx + ex + c * rOut, y: cy + ey + sn * rOut }, angulo: p._medio, enAngulo });
        });
        separarEtiquetas(cajas, adentro);
        // Línea guía: del borde del sector, a la mitad, al punto más cercano de
        // la etiqueta, si quedó lejos (corrida para no pisar a otra, o a mano)
        // y afuera del círculo. [calibrar] desde qué distancia la dibuja PowerPoint.
        // Con <a:ln><a:noFill/> la línea es null: apagada (sin dato, undefined: gris).
        if (e.guias && e.lineaGuia !== null) {
          const lg = e.lineaGuia;
          const guia = lg && !lg.sinColor ? lg : lineaFina({ r: 166, g: 166, b: 166, a: 1 }, (lg && lg.ancho) || 9525);
          for (const k of cajas) {
            if (!k.movida && !k.corrida) continue;
            const x0 = k.lx - k.w / 2, y0 = k.ly - k.h / 2;
            const qx = Math.max(x0, Math.min(x0 + k.w, k.ancla.x)), qy = Math.max(y0, Math.min(y0 + k.h, k.ancla.y));
            if (Math.hypot(qx - k.ancla.x, qy - k.ancla.y) < lh * 0.75 || Math.hypot(k.lx - cx, k.ly - cy) < rOut) continue;
            const path = new Path2D(); path.moveTo(k.ancla.x, k.ancla.y); path.lineTo(qx, qy);
            trazar(guia, path);
          }
        }
        for (const k of cajas) {
          k.lineas.forEach((x, k2) => escribir(ctx, x, l, k.lx, k.ly - k.h / 2 + PAD_TEXTO.y + k2 * lh + asc, 'center'));
          k.p.etiqueta = { texto: k.lineas.join('\n'), x: k.lx - k.w / 2, y: k.ly - k.h / 2, w: k.w, h: k.h };
        }
      }
      for (const p of registro.puntos) { delete p._medio; delete p._abre; delete p._expl; delete p._v; delete p._pct; }
    });
  }

  // Etiquetas de afuera de un circular que se pisan (entre ellas o con una de
  // adentro): se corren alrededor del círculo, de a medio grado, alejándose
  // una de la otra, hasta que ninguna pise a otra (o 400 vueltas). Las que se
  // corrieron más de 2° quedan marcadas para la línea guía. [calibrar] Es
  // parecido a PowerPoint (24-graficos-detalles, diapositiva 1), no igual:
  // PowerPoint además parte el texto de alguna.
  function separarEtiquetas(cajas, adentro) {
    const fuera = cajas.filter((k) => k.fuera);
    if (!fuera.length) return;
    const caja = (k) => ({ x: k.lx - k.w / 2, y: k.ly - k.h / 2, w: k.w, h: k.h });
    const pisa = (a, b) => a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h;
    const PASO = Math.PI / 360;
    for (let vuelta = 0; vuelta < 400; vuelta++) {
      let hubo = false;
      for (let i = 0; i < fuera.length; i++) {
        const a = fuera[i];
        for (let j = i + 1; j < fuera.length; j++) {
          const b = fuera[j];
          if (!pisa(caja(a), caja(b))) continue;
          hubo = true;
          // La de ángulo menor retrocede y la otra avanza (en el círculo).
          const d = Math.atan2(Math.sin(b.angulo - a.angulo), Math.cos(b.angulo - a.angulo)) >= 0 ? 1 : -1;
          a.angulo -= d * PASO; b.angulo += d * PASO;
          ({ x: a.lx, y: a.ly } = a.enAngulo(a.angulo));
          ({ x: b.lx, y: b.ly } = b.enAngulo(b.angulo));
        }
        for (const o of adentro) {
          if (!pisa(caja(a), o)) continue;
          hubo = true;
          const d = Math.atan2(Math.sin(a.angulo - a.p._medio), Math.cos(a.angulo - a.p._medio)) >= 0 ? 1 : -1;
          a.angulo += d * PASO;
          ({ x: a.lx, y: a.ly } = a.enAngulo(a.angulo));
        }
      }
      if (!hubo) break;
    }
    for (const k of fuera) if (Math.abs(Math.atan2(Math.sin(k.angulo - k.p._medio), Math.cos(k.angulo - k.p._medio))) > Math.PI / 90) k.corrida = true;
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
    rellenar(g.leyenda.relleno, rect(L.x, L.y, L.w, L.h), L.w, L.h, L.x, L.y);
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
      const mx = x + (en.larga ? M.desde : g.clasico ? 2.47 : 3.29);
      if (en.tipoMuestra === 'linea') {
        let a0 = mx, a1 = mx + largoMuestra;
        if (M === MUESTRAS.comun) {
          const k = l.tamano / 11.97, tapa = en.linea && en.linea.tapa && en.linea.tapa !== 'flat' ? (en.linea.ancho || 9525) / 12700 / 2 : 0;
          a0 = x + 2.04 * k + tapa; a1 = x + 23.54 * k - tapa + (g.clasico ? 4.55 : 0); // en Excel la muestra de la leyenda es 4,5 pt más larga (la leyenda, 5,3 pt más ancha)
        }
        const p = new Path2D(); p.moveTo(a0, centroY); p.lineTo(a1, centroY);
        trazar(en.linea, p);
        if (en.marcador) dibujarMarcador(ctx, g.clasico && en.marcador.tamano > 5 ? { ...en.marcador, tamano: en.marcador.tamano * 2 / 3 } : en.marcador, en.serie, (a0 + a1) / 2, centroY, rellenar, trazar, en.grupo); // y su marcador, a 2/3 (6 pt con el de 9; medido en el PDF)
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
      // La de línea (muestras comunes), medida en las imágenes de 18
      // gráficos: COM da una caja de 12 pt, pero se ve de 2 a 23,5 pt de la
      // entrada (con la tapa) y el texto empieza a 24,8 (letra de 11,97).
      const tx = en.tipoMuestra === 'linea' && M === MUESTRAS.comun ? x + 24.8 / 11.97 * l.tamano + (g.clasico ? 4.0 : 0) : mx + largoMuestra + (g.clasico ? 2.4 : 3.5); // medido en píxeles
      lineasEntrada.forEach((t, k) => escribir(ctx, t, l, tx, y + asc + k * lh - (g.clasico && !L.horizontal ? 0.54 : 0), 'left')); // en Excel el texto de una leyenda vertical queda 0,54 pt más arriba
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

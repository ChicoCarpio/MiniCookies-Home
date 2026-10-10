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
    radarChart: 'radar', bubbleChart: 'burbuja', stockChart: 'cotizaciones',
  };
  // Lo que cambia mucho el dibujo y todavía no sabe hacer: con eso, el
  // gráfico va al respaldo en vez de salir distinto sin aviso.
  const NO_SOPORTADOS = ['ofPieChart', 'bar3DChart', 'line3DChart', 'area3DChart', 'pie3DChart', 'trendline', 'errBars', 'dTable', 'dispUnits', 'hiLowLines', 'upDownBars', 'logBase'];
  // Lo que PowerPoint escribe cuando el archivo no trae el texto: en el idioma
  // de quien mira (como Office, que lo pone en el de su instalación).
  const TEXTOS = /^es/i.test((typeof navigator !== 'undefined' && navigator.language) || '') ? { serie: 'Serie', titulo: 'Título del gráfico' } : { serie: 'Series', titulo: 'Chart Title' };
  // Los que sabe dibujar; con otro, el gráfico queda sin soporte (lo dibuja
  // la librería de respaldo).
  const SOPORTADOS = new Set(['barra', 'linea', 'area', 'circular', 'anillo', 'dispersion']);

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
    for (const ax of plot.children) if (/Ax$/.test(ax.localName) && hijo(ax, 'title')) throw new Error('gráfico con título de eje');
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
      });
    }

    const grupos = [];
    let indiceSerie = 0;
    for (const el of plot.children) {
      const tipo = TIPOS[el.localName];
      if (!tipo) continue;
      if (!SOPORTADOS.has(tipo)) throw new Error('gráfico de tipo ' + el.localName);
      const g = {
        tipo, dir: val(hijo(el, 'barDir'), 'col'), agrupacion: val(hijo(el, 'grouping'), tipo === 'barra' ? 'clustered' : 'standard'),
        separacion: num(hijo(el, 'gapWidth'), 150) / 100, solapamiento: num(hijo(el, 'overlap'), 0) / 100,
        hueco: num(hijo(el, 'holeSize'), 10) / 100, anguloInicial: num(hijo(el, 'firstSliceAng'), 0),
        variarColores: si(hijo(el, 'varyColors')), estiloDispersion: val(hijo(el, 'scatterStyle'), 'marker'),
        ejes: hijos(el, 'axId').map((a) => val(a)),
        etiquetas: etiquetasDe(hijo(el, 'dLbls'), r, letraBase, null),
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
          relleno: r.relleno(spPr), linea: r.linea(spPr), // sinColor: lo completa automatico
          suave: si(hijo(s, 'smooth')), explosion: num(hijo(s, 'explosion'), 0) / 100,
          // Sin c:marker, la norma dice automático: en la dispersión (salvo
          // estilo de solo línea) y en las líneas con <c:marker val="1"/> en el grupo.
          marcador: hijo(s, 'marker') ? marcadorDe(hijo(s, 'marker'), r)
            : (tipo === 'dispersion' && !['line', 'smooth'].includes(g.estiloDispersion)) || (tipo === 'linea' && si(hijo(el, 'marker'))) ? { simbolo: 'auto', tamano: 5, relleno: undefined, linea: undefined } : null,
          puntos: new Map(),
          etiquetas: etiquetasDe(hijo(s, 'dLbls'), r, letraBase, g.etiquetas),
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
        if (s.relleno === undefined) s.relleno = g.tipo === 'linea' || g.tipo === 'dispersion' ? null : { tipo: 'solido', color: c };
        if (s.linea === undefined || (s.linea && s.linea.sinColor)) {
          const ancho = s.linea && s.linea.ancho;
          s.linea = g.tipo === 'linea' || g.tipo === 'dispersion' ? lineaFina(c, ancho || 28575) : g.tipo === 'circular' || g.tipo === 'anillo' ? lineaFina({ r: 255, g: 255, b: 255, a: 1 }, ancho || 19050) : null;
        }
        s.colorAuto = c;
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
      titulo = { texto, letra, relleno: r.relleno(hijo(tituloEl, 'spPr')) ?? null, linea: linea(hijo(tituloEl, 'spPr')) ?? null };
    }
    void borrado; // [calibrar] autoTitleDeleted="0" sin c:title y una sola serie: ¿PowerPoint muestra su nombre?

    const leyendaEl = hijo(chart, 'legend');
    const leyenda = leyendaEl ? {
      pos: val(hijo(leyendaEl, 'legendPos'), 'r'), letra: letraDe(hijo(leyendaEl, 'txPr'), r, letraBase),
      borradas: new Set(hijos(leyendaEl, 'legendEntry').filter((e) => si(hijo(e, 'delete'))).map((e) => num(hijo(e, 'idx'), -1))),
      relleno: r.relleno(hijo(leyendaEl, 'spPr')) ?? null, linea: linea(hijo(leyendaEl, 'spPr')) ?? null,
    } : null;

    return {
      grupos, ejes, titulo, leyenda, letraBase, automatico,
      // <a:noFill/> (null) es transparente; sin decir nada (undefined), blanco.
      fondo: { relleno: conDefecto(r.relleno(hijo(raiz, 'spPr')), { tipo: 'solido', color: { r: 255, g: 255, b: 255, a: 1 } }), linea: linea(hijo(raiz, 'spPr')) ?? null },
      fondoTrazado: { relleno: r.relleno(hijo(plot, 'spPr')) ?? null, linea: linea(hijo(plot, 'spPr')) ?? null },
      vacios: val(hijo(chart, 'dispBlanksAs'), 'gap'),
    };
  }

  const conDefecto = (v, def) => (v === undefined ? def : v);
  const lineaFina = (color, ancho = 9525) => ({ ancho, relleno: { tipo: 'solido', color }, tapa: 'flat', compuesto: 'sng', guion: 'solid', union: 'round', inicio: null, fin: null });

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
    if (!Number.isFinite(a) || !Number.isFinite(b)) { a = 0; b = 1; }
    if (fijo.min != null) a = Math.min(a, fijo.min);
    if (fijo.max != null) b = Math.max(b, fijo.max);
    if (a === b) { if (a === 0) b = 1; else if (a > 0) a = 0; else b = 0; }
    const ceroAbajo = fijo.min == null && a >= 0 && a < (5 / 6) * b;
    const ceroArriba = fijo.max == null && b <= 0 && b > (5 / 6) * a;
    let lo, hi;
    if (ceroAbajo) { lo = 0; hi = b + b / 20; }
    else if (ceroArriba) { hi = 0; lo = a + a / 20; }
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

  function claseEje(e) { return e.clase === 'valAx' ? 'valor' : e.clase === 'dateAx' ? 'fecha' : 'categoria'; }

  function maquetar(g, W, H, med) {
    const geo = { W, H, ejes: [], series: [] };
    let arriba = MARGEN, abajo = H - MARGEN, izq = MARGEN, der = W - MARGEN;

    if (g.titulo) {
      const l = g.titulo.letra, lh = med.linea(l);
      const lineas = cortar(g.titulo.texto, l, W * 0.8, med);
      const w = Math.max(...lineas.map((t) => med.ancho(t, l))) + 2 * PAD_TEXTO.x, h = lineas.length * lh + 2 * PAD_TEXTO.y;
      geo.titulo = { x: W / 2 - w / 2, y: arriba, w, h, lineas, letra: l };
      arriba += h + SEPARACION;
    }

    const circular = g.grupos[0].tipo === 'circular' || g.grupos[0].tipo === 'anillo';
    if (g.leyenda) {
      const ley = entradasLeyenda(g, circular);
      const l = g.leyenda.letra, lh = med.linea(l);
      const pos = g.leyenda.pos;
      const anchos = ley.map((e) => med.ancho(e.texto, l));
      // [calibrar] la muestra: 20,4 pt por entrada con letra de 11,97 (medido).
      const porEntrada = (e) => (e.larga ? PASO_LINEA : PASO_CAJA) * l.tamano;
      const padV = (5.85 / 14.61) * lh;
      let w, h;
      if (pos === 'b' || pos === 't') {
        w = INICIO_LEYENDA + ley.reduce((a, e, i) => a + porEntrada(e) + anchos[i], 0);
        h = lh + padV;
      } else {
        w = Math.max(0, ...anchos) + (ley.some((e) => e.larga) ? 32.1 : 19.31) / 11.97 * l.tamano; // medido: 19,31 con muestra cuadrada
        h = ley.length * (lh + padV);
      }
      w = Math.min(w, der - izq);
      let x, y;
      if (pos === 'b') { x = W / 2 - w / 2; y = abajo - h; abajo = y - SEPARACION; }
      else if (pos === 't' ) { x = W / 2 - w / 2; y = arriba; arriba = y + h + SEPARACION; }
      else if (pos === 'l') { x = izq; y = (arriba + abajo) / 2 - h / 2; izq = x + w + SEPARACION; }
      else { x = der - w; y = (pos === 'tr' ? arriba : (arriba + abajo) / 2 - h / 2); der = x - SEPARACION; }
      geo.leyenda = { x, y, w, h, entradas: ley, anchos, letra: l, horizontal: pos === 'b' || pos === 't', padV };
    }

    if (circular) {
      // El círculo, cuadrado y centrado en lo que queda. Con etiquetas de datos,
      // un 14 % más chico (medido: 140,58 de 163,81 y 163,74 de 190,27). [calibrar]
      const zona = { x: izq + AIRE_TRAZADO, y: arriba + AIRE_TRAZADO, w: der - izq - 2 * AIRE_TRAZADO, h: abajo - arriba - 2 * AIRE_TRAZADO };
      let lado = Math.max(0, Math.min(zona.w, zona.h));
      const s0 = g.grupos[0].series[0];
      const conEtiquetas = s0 && s0.etiquetas && !s0.etiquetas.borrar && (s0.etiquetas.showVal || s0.etiquetas.showPercent || s0.etiquetas.showCatName || s0.etiquetas.showSerName);
      if (conEtiquetas) lado *= 0.86;
      geo.trazado = { x: zona.x + zona.w / 2 - lado / 2, y: zona.y + zona.h / 2 - lado / 2, w: lado, h: lado };
      geo.circular = true;
      return geo;
    }

    // Ejes visibles y su lado.
    const ejes = [];
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
    // Lado de cada eje: el de categorías de un gráfico de barras horizontales va
    // a la izquierda; los de valores, según axPos.
    const lado = (e) => e.pos;
    const visibles = ejes.filter((e) => !e.borrado && e.posEtiquetas !== 'none');
    const etiquetasDeEje = (e, esc) => {
      if (claseEje(e) === 'valor' && esc) {
        const lista = [];
        const formato = e.formato.enlazado ? null : e.formato.codigo;
        const dec = Math.max(decimalesDe(esc.unidad), decimalesDe(esc.min));
        for (let v = esc.min, k = 0; v <= esc.max + esc.unidad * 1e-6 && k < 200; v = esc.min + (++k) * esc.unidad) {
          const r = parseFloat(v.toFixed(12));
          lista.push({ v: r, texto: formato ? formatear(r, formato) : (rangos.get(e.id) || {}).porcentaje ? formatear(r, '0%') : formatoDeSerie(g, e, r, dec) });
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
        const fijo = { min: e.minFijo, max: e.maxFijo, unidad: e.unidadFija };
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
    let interior;
    for (let vuelta = 0; vuelta < 3; vuelta++) {
      let l = izq, rr = der, t = arriba, b = abajo;
      const reserva = { l: 0, r: 0, t: 0, b: 0 };
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
          const ancho = interior ? interior.w : rr - l;
          const n = Math.max(1, lista.length);
          let lineas = 1;
          if (claseEje(e) !== 'valor') {
            const porCategoria = ancho / (enBordes(e) ? Math.max(1, n - 1) : n);
            for (const x of lista) lineas = Math.max(lineas, cortar(x.texto, e.letra, Math.max(1, porCategoria - 2 * PAD_TEXTO.x), med).length);
            if (enBordes(e) && lista.length) {
              extra.l = Math.max(extra.l, med.ancho(lista[0].texto, e.letra) / 2 + AIRE_BORDE);
              extra.r = Math.max(extra.r, med.ancho(lista[lista.length - 1].texto, e.letra) / 2 + AIRE_BORDE);
            }
          } else {
            // El último número sobresale la mitad de su ancho a la derecha.
            const ult = lista[lista.length - 1];
            if (ult) extra.r = Math.max(extra.r, med.ancho(ult.texto, e.letra) / 2 + AIRE_TRAZADO);
          }
          if (fuera) reserva[ladoE] = Math.max(reserva[ladoE], lineas * lh + AIRE_ABAJO);
        }
      }
      l += Math.max(reserva.l || AIRE_TRAZADO, extra.l);
      rr -= Math.max(reserva.r || AIRE_TRAZADO, extra.r);
      t += reserva.t || AIRE_TRAZADO + Math.max(0, extra.t);
      b -= reserva.b || AIRE_TRAZADO + Math.max(0, extra.b);
      const nuevo = { x: l, y: t, w: Math.max(1, rr - l), h: Math.max(1, b - t) };
      const igual = interior && Math.abs(interior.w - nuevo.w) < 0.01 && Math.abs(interior.h - nuevo.h) < 0.01;
      interior = nuevo;
      calcularEscalas(interior.w, interior.h);
      if (igual) break;
    }
    geo.trazado = interior;
    geo.horizontal = horizontal;
    geo.escalas = escalas;
    for (const e of visibles) geo.ejes.push({ eje: e, etiquetas: etiquetasDeEje(e, escalas.get(e.id)), adentro: adentro(e), enBordes: enBordes(e) });
    geo.todosLosEjes = ejes;
    return geo;
  }

  // Texto de las categorías del eje (de la primera serie que las tenga).
  function categoriasDe(g, e) {
    for (const grupo of g.grupos) {
      if (!grupo.ejes.includes(e.id)) continue;
      const s = grupo.series.find((x) => x.categorias && x.categorias.length);
      const n = Math.max(0, ...grupo.series.map((x) => x.valores.length));
      if (s) return Array.from({ length: Math.max(n, s.categorias.length) }, (_, i) => {
        const c = s.categorias[i];
        return c == null ? '' : typeof c === 'number' ? formatear(c, s.formatoCategorias) : String(c);
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
      const f = s0 && (grupo.tipo === 'dispersion' && grupo.ejes[0] === e.id ? s0.formatoCategorias : s0.formato);
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
      for (const s of grupo.series) {
        if (g.leyenda.borradas.has(s.idx)) continue;
        const conLinea = grupo.tipo === 'linea' || (grupo.tipo === 'dispersion' && grupo.estiloDispersion !== 'marker');
        // La dispersión de solo puntos ocupa en la leyenda lo mismo que una línea (medido).
        lista.push({ texto: s.nombre, relleno: s.relleno, linea: s.linea, marcador: s.marcador, tipoMuestra: conLinea ? 'linea' : grupo.tipo === 'dispersion' ? 'marcador' : 'caja', serie: s, grupo, larga: grupo.tipo === 'linea' || (grupo.tipo === 'dispersion' && !!s.linea) });
      }
    }
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

      if (geo.circular) dibujarCircular(ctx, g, geo, med, rellenar, trazar, salida);
      else dibujarEjesYSeries(ctx, g, geo, med, rellenar, trazar, salida);

      if (geo.titulo) {
        const ti = geo.titulo, l = ti.letra, lh = med.linea(l);
        rellenar(g.titulo.relleno, rect(ti.x, ti.y, ti.w, ti.h), ti.w, ti.h);
        trazar(g.titulo.linea, rect(ti.x, ti.y, ti.w, ti.h));
        ti.lineas.forEach((texto, i) => escribir(ctx, texto, l, ti.x + ti.w / 2, ti.y + PAD_TEXTO.y + i * lh + med.ascendente(l), 'center'));
      }
      if (geo.leyenda) dibujarLeyenda(ctx, g, geo, med, rellenar, trazar);
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
    const t = geo.trazado;
    const escalas = geo.escalas;
    // Posición en el trazado de un valor de un eje de valores.
    const aPos = (e, v) => {
      const s = escalas.get(e.id);
      let u = (v - s.min) / (s.max - s.min);
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
        for (let k = 0, v = s.min; v <= s.max + s.unidad * 1e-6 && k < 200; v = s.min + (++k) * s.unidad) {
          const p = new Path2D(); const q = aPos(e, v);
          if (vertical) { p.moveTo(q, t.y); p.lineTo(q, t.y + t.h); } else { p.moveTo(t.x, q); p.lineTo(t.x + t.w, q); }
          trazar(e.cuadricula, p);
        }
      } else {
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
        : grupo.tipo === 'dispersion' ? dispersion(grupo, ex, ey, aPos)
          : lineasOAreas(g, grupo, ex, ey, t, aPos, escalas);
      // Áreas primero, después barras, después líneas (como PowerPoint).
      // Las barras y áreas se cortan en el borde del trazado (un máximo fijo
      // menor que los datos); las líneas y marcadores, con un poco de aire.
      const recortar = (aire) => { ctx.save(); ctx.beginPath(); ctx.rect(t.x - aire, t.y - aire, t.w + 2 * aire, t.h + 2 * aire); ctx.clip(); };
      recortar(0);
      for (const sg of geoGrupo) {
        const s = sg.serie;
        if (grupo.tipo === 'barra') {
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
      ctx.restore();
      // Etiquetas de datos.
      for (const sg of geoGrupo) etiquetasDeSerie(ctx, g, grupo, sg, med, t, salida);
    }

    // Ejes: línea y etiquetas.
    for (const { eje: e, etiquetas, adentro, enBordes } of geo.ejes) {
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
        for (const x of etiquetas) {
          const q = aPos(e, x.v);
          if (e.pos === 'l') escribir(ctx, x.texto, e.letra, t.x - AIRE_VALORES + CORRE_VALORES, q - lh / 2 + asc, 'right');
          else if (e.pos === 'r') escribir(ctx, x.texto, e.letra, t.x + t.w + AIRE_VALORES - CORRE_VALORES, q - lh / 2 + asc, 'left');
          else if (e.pos === 'b') escribir(ctx, x.texto, e.letra, q, t.y + t.h + BAJA_ETIQUETAS + asc, 'center');
          else escribir(ctx, x.texto, e.letra, q, t.y - BAJA_ETIQUETAS - lh + asc, 'center');
        }
      } else {
        const n = Math.max(1, etiquetas.length);
        const vertical = e.pos === 'l' || e.pos === 'r';
        const paso = (vertical ? t.h : t.w) / (enBordes ? Math.max(1, n - 1) : n);
        const desplazamiento = enBordes ? 0 : 0.5;
        const salto = e.salto || 1;
        // Junto a la línea del eje si cruza en el medio; si no, afuera del trazado.
        const bordeY = adentro && donde != null ? donde : t.y + t.h;
        const bordeX = adentro && donde != null ? donde : e.pos === 'l' ? t.x : t.x + t.w;
        etiquetas.forEach((x, i) => {
          if (i % salto) return;
          const u = paso * (i + desplazamiento);
          const centro = vertical ? (e.invertido ? t.y + u : t.y + t.h - u) : (e.invertido ? t.x + t.w - u : t.x + u);
          if (vertical) {
            const xDer = e.pos === 'l' ? bordeX - AIRE_CATEGORIAS_IZQ : bordeX + AIRE_CATEGORIAS_IZQ;
            const lineas = [x.texto];
            lineas.forEach((tx, k) => escribir(ctx, tx, e.letra, xDer, centro - lineas.length * lh / 2 + k * lh + asc, e.pos === 'l' ? 'right' : 'left'));
          } else {
            const lineas = cortar(x.texto, e.letra, Math.max(1, paso - 2 * PAD_TEXTO.x), med);
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
      const tipoCom = (claseEje(e) === 'valor' && !(g.grupos[grupoEje] && g.grupos[grupoEje].tipo === 'dispersion' && (e.pos === 'b' || e.pos === 't'))) ? 2 : 1;
      const secundario = grupoEje > 0 && g.grupos[grupoEje].ejes.some((id) => !g.grupos[0].ejes.includes(id));
      salida.ejes.push({ tipo: tipoCom, grupo: secundario ? 2 : 1, borrado: e.borrado, ...(s ? { min: s.min, max: s.max, unidad: s.unidad } : {}) });
    }
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
    const pasos = L.entradas.map((en, i) => (en.larga ? PASO_LINEA : PASO_CAJA) * l.tamano + L.anchos[i]);
    let x = L.horizontal ? L.x + (L.w - INICIO_LEYENDA - pasos.reduce((a, b) => a + b, 0)) / 2 + INICIO_LEYENDA : L.x + 1.5, y = L.y + L.padV / 2;
    L.entradas.forEach((en, i) => {
      const largoMuestra = en.larga ? l.tamano : lado;
      const centroY = y + lh / 2;
      const mx = x + (en.larga ? 6 : 3.29);
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
      escribir(ctx, en.texto, l, tx, y + asc, 'left');
      if (L.horizontal) x += pasos[i];
      else y += lh + L.padV;
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

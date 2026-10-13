/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Estilos de tabla de Excel del motor propio (etapa 6): qué relleno, letra y bordes le da a cada celda el estilo de la tabla que la contiene.
// Los estilos integrados (60) vienen de xlsx-tablas-estilos.js (medidos en Excel real); los propios del libro, de styles.xml (tableStyles + dxfs).
// Cada estilo define elementos (toda la tabla, encabezado, totales, primera/última columna, filas y columnas en bandas, celdas de las esquinas
// del encabezado y de los totales); sobre una celda pueden valer varios y gana el de mayor precedencia en cada propiedad. El formato propio de
// la celda gana al de la tabla, y el formato condicional gana a los dos. Reglas medidas contra Range.DisplayFormat
// (Investigaciones\investigacion-motor-excel\03-mediciones.md, etapa 6). Sin DOM. Requiere xlsx-modelo.js y xlsx-estilos.js.
//
// Uso:
//   const tablas = XlsxTablas.preparar(libro, hoja);        // null si la hoja no tiene tablas con estilo
//   const e = tablas.para(fila, col);                       // null o { relleno, fuente, borde } (colores como referencias de estilo)
//   const xf2 = XlsxTablas.aplicar(libro, xf, e);           // el estilo de la celda con el de la tabla debajo de su formato propio

(function () {
  const TIPOS = ['wholeTable', 'headerRow', 'totalRow', 'firstColumn', 'lastColumn', 'firstRowStripe', 'secondRowStripe', 'firstColumnStripe', 'secondColumnStripe',
    'firstHeaderCell', 'lastHeaderCell', 'firstTotalCell', 'lastTotalCell'];
  // De menor a mayor precedencia (índices de elemento).
  const ORDEN = [0, 7, 8, 5, 6, 3, 4, 1, 2, 9, 10, 11, 12];
  const BORDES = { t: 'thin', m: 'medium', k: 'thick', d: 'double', h: 'hair', a: 'dashed', p: 'dotted' };
  const MAX_TABLAS = 5000;

  const refTema = (r) => (r[0] === -1 ? { tipo: 'rgb', valor: [r[2] & 255, (r[2] >> 8) & 255, (r[2] >> 16) & 255].map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase(), tint: 0 } : { tipo: 'tema', valor: r[0], tint: r[1] || 0 });

  // Un estilo → 13 elementos normalizados { f, n, i, c, b: { izq, arr, aba, der, vert, horiz }, tam }
  function integrado(nombre) {
    const d = typeof XlsxTablasEstilos !== 'undefined' && XlsxTablasEstilos.ESTILOS[nombre];
    if (!d) return null;
    const els = new Array(13).fill(null);
    for (const k of Object.keys(d)) {
      const e = d[k], o = { f: null, n: undefined, i: undefined, c: null, b: null, tam: 1 };
      if (e.f) o.f = refTema(e.f);
      if (e.n) o.n = true;
      if (e.i) o.i = true;
      if (e.c) o.c = refTema(e.c);
      if (e.b) {
        o.b = {};
        for (const [lado, nom] of [['l', 'izq'], ['a', 'arr'], ['b', 'aba'], ['r', 'der'], ['v', 'vert'], ['h', 'horiz']]) if (e.b[lado]) o.b[nom] = { estilo: BORDES[e.b[lado][0]] || 'thin', color: refTema(e.b[lado].slice(1)) };
      }
      els[+k] = o;
    }
    return els;
  }
  function propio(libro, def) {
    const els = new Array(13).fill(null);
    TIPOS.forEach((tipo, k) => {
      const e = def.elementos[tipo];
      if (!e) return;
      const dx = e.dxf !== null && e.dxf !== undefined ? libro.estilos.dxf(e.dxf) : null;
      const o = { f: null, n: undefined, i: undefined, c: null, b: null, tam: e.tam > 0 ? e.tam : 1 };
      if (dx) {
        if (dx.relleno && dx.relleno.tipo !== 'ninguno') o.f = dx.relleno.patron === 'solid' ? (dx.relleno.bg || dx.relleno.fg) : dx.relleno.fg;
        const fu = dx.fuente;
        if (fu) {
          const dado = fu.dado || new Set();
          if (dado.has('b')) o.n = fu.negrita;
          if (dado.has('i')) o.i = fu.cursiva;
          if (fu.color) o.c = fu.color;
        }
        if (dx.borde) { o.b = {}; for (const l of ['izq', 'arr', 'aba', 'der', 'vert', 'horiz']) if (dx.borde[l]) o.b[l] = dx.borde[l]; }
      }
      els[k] = o;
    });
    return els;
  }

  function preparar(libro, hoja) {
    const lista = (hoja.tablas || []).filter((t) => t.estilo).slice(0, MAX_TABLAS);
    if (!lista.length) return null;
    const estilos = new Map();
    const tablas = [];
    for (const t of lista) {
      if (!estilos.has(t.estilo)) {
        const p = libro.estilos.estiloTabla ? libro.estilos.estiloTabla(t.estilo) : null;
        estilos.set(t.estilo, p ? propio(libro, p) : integrado(t.estilo));
      }
      const els = estilos.get(t.estilo);
      if (!els) continue;
      tablas.push({ t, els, f0: t.ref.f0, c0: t.ref.c0, f1: t.ref.f1, c1: t.ref.c1, fb0: t.ref.f0 + t.encabezados, fb1: t.ref.f1 - t.totales });
    }
    if (!tablas.length) return null;

    // La región (rectángulo) de cada elemento que cubre la celda, o null.
    function region(T, k, f, c) {
      const { t } = T;
      switch (k) {
        case 0: return { f0: T.f0, f1: T.f1, c0: T.c0, c1: T.c1 };
        case 1: return t.encabezados && f === T.f0 ? { f0: T.f0, f1: T.f0, c0: T.c0, c1: T.c1 } : null;
        case 2: return t.totales && f === T.f1 ? { f0: T.f1, f1: T.f1, c0: T.c0, c1: T.c1 } : null;
        case 3: return t.primeraCol && c === T.c0 ? { f0: T.f0, f1: T.f1, c0: T.c0, c1: T.c0 } : null;
        case 4: return t.ultimaCol && c === T.c1 ? { f0: T.f0, f1: T.f1, c0: T.c1, c1: T.c1 } : null;
        case 5: case 6: {
          if (!t.bandasFilas || f < T.fb0 || f > T.fb1) return null;
          const tam = (T.els[5] || T.els[6] || { tam: 1 }).tam || 1;
          const banda = Math.floor((f - T.fb0) / tam);
          if ((banda % 2 === 0) !== (k === 5)) return null;
          return { f0: T.fb0 + banda * tam, f1: Math.min(T.fb1, T.fb0 + (banda + 1) * tam - 1), c0: T.c0, c1: T.c1 };
        }
        case 7: case 8: {
          if (!t.bandasCols || f < T.fb0 || f > T.fb1) return null;
          const tam = (T.els[7] || T.els[8] || { tam: 1 }).tam || 1;
          const banda = Math.floor((c - T.c0) / tam);
          if ((banda % 2 === 0) !== (k === 7)) return null;
          return { f0: T.fb0, f1: T.fb1, c0: T.c0 + banda * tam, c1: Math.min(T.c1, T.c0 + (banda + 1) * tam - 1) };
        }
        case 9: return t.encabezados && f === T.f0 && c === T.c0 ? { f0: f, f1: f, c0: c, c1: c } : null;
        case 10: return t.encabezados && f === T.f0 && c === T.c1 ? { f0: f, f1: f, c0: c, c1: c } : null;
        case 11: return t.totales && f === T.f1 && c === T.c0 ? { f0: f, f1: f, c0: c, c1: c } : null;
        case 12: return t.totales && f === T.f1 && c === T.c1 ? { f0: f, f1: f, c0: c, c1: c } : null;
        default: return null;
      }
    }

    function calcular(f, c) {
      const T = tablas.find((q) => f >= q.f0 && f <= q.f1 && c >= q.c0 && c <= q.c1);
      if (!T) return null;
      const aplican = [];
      for (const k of ORDEN) { const e = T.els[k]; if (!e) continue; const r = region(T, k, f, c); if (r) aplican.push({ e, r }); }
      if (!aplican.length) return null;
      const res = { relleno: null, fuente: null, borde: null };
      for (let i = aplican.length - 1; i >= 0; i--) { // de mayor a menor precedencia: manda el primero que define cada cosa
        const { e, r } = aplican[i];
        if (e.f && !res.relleno) res.relleno = { patron: 'solid', fg: e.f, bg: null };
        if (e.n !== undefined || e.i !== undefined || e.c) {
          res.fuente = res.fuente || {};
          if (e.n !== undefined && res.fuente.negrita === undefined) res.fuente.negrita = e.n;
          if (e.i !== undefined && res.fuente.cursiva === undefined) res.fuente.cursiva = e.i;
          if (e.c && res.fuente.color === undefined) res.fuente.color = e.c;
        }
        if (e.b) {
          const lado = (nombre, interior, hayVecino) => { const b = hayVecino ? e.b[interior] : e.b[nombre]; if (b && !(res.borde && res.borde[nombre])) { res.borde = res.borde || {}; res.borde[nombre] = Object.assign({ deTabla: true }, b); } };
          lado('izq', 'vert', c > r.c0); lado('der', 'vert', c < r.c1);
          lado('arr', 'horiz', f > r.f0); lado('aba', 'horiz', f < r.f1);
        }
      }
      return res.relleno || res.fuente || res.borde ? res : null;
    }

    const memo = new Map();
    return {
      para(f, c) {
        const k = f * 16384 + c;
        if (memo.has(k)) return memo.get(k);
        const r = calcular(f, c);
        if (memo.size < 200000) memo.set(k, r);
        return r;
      },
      reiniciar() { memo.clear(); },
      tablas: tablas.length,
    };
  }

  // El estilo de la celda con el de su tabla debajo del formato propio: solo se pone lo que la celda no trae por su cuenta.
  function aplicar(libro, xf, efecto) {
    if (!efecto) return xf;
    const r = Object.assign({}, xf);
    if (efecto.relleno && (!xf.relleno || xf.relleno.tipo === 'ninguno' || xf.relleno.patron === 'none')) r.relleno = { tipo: 'patron', patron: 'solid', fg: efecto.relleno.fg, bg: null, degradado: null };
    if (efecto.fuente) {
      const fu = Object.assign({}, xf.fuente), e = efecto.fuente;
      const porDefecto = libro.estilos.fuentes[0] || {};
      const hex = (x) => (x ? libro.estilos.color(x) : null);
      if (e.color && hex(xf.fuente.color) === hex(porDefecto.color)) fu.color = e.color;
      if (e.negrita !== undefined && !xf.fuente.negrita) fu.negrita = e.negrita;
      if (e.cursiva !== undefined && !xf.fuente.cursiva) fu.cursiva = e.cursiva;
      r.fuente = fu;
    }
    if (efecto.borde) {
      const b = Object.assign({ izq: null, der: null, arr: null, aba: null, vert: null, horiz: null, diagonal: null, diagArriba: false, diagAbajo: false }, xf.borde || {});
      for (const l of ['izq', 'arr', 'aba', 'der']) if (efecto.borde[l]) b[l] = XlsxEstilos.mejorBorde(xf.borde && xf.borde[l], efecto.borde[l], (x) => libro.estilos.color(x)); // el de la celda, salvo que el de la tabla gane (más pesado o más oscuro)
      r.borde = b;
    }
    return r;
  }

  globalThis.XlsxTablas = { preparar, aplicar };
})();

/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Minigráficos de Excel 2010 (sparklines, extensión x14) del motor propio (etapa 6): línea, columnas y ganancia/pérdida dentro de una celda.
// hoja.minigraficos (xlsx-modelo.js) trae los grupos con sus colores y opciones y, por minigráfico, el rango de datos y la celda donde va.
// Medido en el PDF de Excel (36-x14): el minigráfico ocupa la celda con 5 u a los lados y 3 u arriba y abajo (u = un píxel de la impresora,
// 0,96 pt); n columnas miden el ancho / (n + 0,4·(n − 1)) con un hueco de 0,4 de su ancho, y los puntos de la línea caen en el centro de cada
// columna; la línea de 0,75 pt va del mínimo al máximo de la región vertical; las columnas salen de la línea del cero; en ganancia/pérdida
// cada barra ocupa media región. Sin DOM: solo la API de contexto 2D. Requiere xlsx-modelo.js y xlsx-estilos.js.
//
// Uso:
//   const mini = XlsxMinigraficos.preparar(libro, hoja);       // null si la hoja no tiene minigráficos
//   const s = mini.para(fila, col);                             // null o { grupo, valores } de la celda
//   XlsxMinigraficos.dibujar(ctx, libro, s, x0, y0, x1, y1, u); // u: el píxel de la impresora en unidades del lienzo (1 en pantalla)

(function () {
  const MAX_VALORES = 5000;
  const PT = 0.96; // pt por u

  // 'Hoja1'!A1:F1 → { hoja: nombre | null, f0, c0, f1, c1 }
  function leerReferencia(ref) {
    const m = /^(?:(?:'((?:[^']|'')+)'|([^!']+))!)?(.+)$/.exec((ref || '').trim());
    if (!m) return null;
    const r = LibroXlsx.decodificarRango(m[3].replace(/\$/g, ''));
    if (!r) return null;
    return { hoja: m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2] !== undefined ? m[2] : null, f0: r.f0, c0: r.c0, f1: r.f1, c1: r.c1 };
  }

  function preparar(libro, hoja) {
    const grupos = hoja.minigraficos || [];
    if (!grupos.length) return null;
    const C = hoja.celdas, TIPO = LibroXlsx.TIPO;
    const nombreHoja = libro.hojas[hoja.indice] && libro.hojas[hoja.indice].nombre;
    const porCelda = new Map();
    function valoresDe(r) { // los números del rango en orden de lectura (null: vacía o que no es número)
      const out = [];
      for (let f = r.f0; f <= r.f1 && out.length < MAX_VALORES; f++) {
        for (let c = r.c0; c <= r.c1 && out.length < MAX_VALORES; c++) {
          const i = C.indiceDe(f, c);
          out.push(i >= 0 && C.tipo[i] === TIPO.numero && Number.isFinite(C.valor[i]) ? C.valor[i] : null);
        }
      }
      return out;
    }
    for (const g of grupos) {
      g.hechos = [];
      for (const it of g.items) {
        const r = leerReferencia(it.f), d = LibroXlsx.decodificarRango((it.sqref || '').replace(/\$/g, ''));
        if (!r || !d || (r.hoja !== null && r.hoja !== nombreHoja)) continue; // los datos de otra hoja no están cargados aquí
        const s = { grupo: g, valores: valoresDe(r) };
        g.hechos.push(s);
        if (porCelda.size < MAX_VALORES) porCelda.set(d.f0 * 16384 + d.c0, s);
      }
    }
    for (const g of grupos) { // escala común del grupo (minAxisType / maxAxisType = 'group')
      const todos = [].concat(...(g.hechos || []).map((s) => s.valores.filter((v) => v !== null)));
      g.comun = todos.length ? { min: Math.min(...todos), max: Math.max(...todos) } : null;
    }
    return { para: (f, c) => porCelda.get(f * 16384 + c) || null };
  }

  const hex = (libro, r, defecto) => (r ? libro.estilos.color(r) || defecto : defecto);

  function dibujar(ctx, libro, s, x0, y0, x1, y1, u) {
    const g = s.grupo;
    u = u > 0 ? u : 1;
    const mx = 5 * u, my = 3 * u;
    const W = x1 - x0 - 2 * mx, H = y1 - y0 - 2 * my;
    if (!(W > 1) || !(H > 1)) return;
    let vals = s.valores.slice();
    if (g.vacias === 'zero') vals = vals.map((v) => (v === null ? 0 : v));
    const n = vals.length;
    if (!n) return;
    const finitos = vals.filter((v) => v !== null);
    if (!finitos.length) return;
    let lo = Math.min(...finitos), hi = Math.max(...finitos);
    if (g.minTipo === 'group' && g.comun) lo = g.comun.min; else if (g.minTipo === 'custom' && g.minManual !== null) lo = g.minManual;
    if (g.maxTipo === 'group' && g.comun) hi = g.comun.max; else if (g.maxTipo === 'custom' && g.maxManual !== null) hi = g.maxManual;
    const col = (nombre, defecto) => hex(libro, g.colores[nombre], defecto);
    const serie = col('series', '#376092'), negativo = col('negative', '#D00000'), colorEje = col('axis', '#000000');
    const marcador = col('markers', '#D00000');
    const i0 = vals.findIndex((v) => v !== null);
    let i1 = -1; for (let k = n - 1; k >= 0; k--) if (vals[k] !== null) { i1 = k; break; }
    const iMax = finitos.length ? vals.indexOf(Math.max(...finitos)) : -1, iMin = finitos.length ? vals.indexOf(Math.min(...finitos)) : -1;
    // color propio de cada punto: primero, último, más alto, más bajo y negativos
    const colorDe = (k, base) => {
      if (g.ultimo && k === i1) return col('last', '#D00000');
      if (g.primero && k === i0) return col('first', '#D00000');
      if (g.alto && k === iMax) return col('high', '#D00000');
      if (g.bajo && k === iMin) return col('low', '#D00000');
      if (g.negativos && vals[k] < 0) return negativo;
      return base;
    };
    const b = n > 1 ? W / (n + 0.4 * (n - 1)) : W, paso = n > 1 ? 1.4 * b : 0;
    const izq = (k) => x0 + mx + k * paso + (g.derechaAIzquierda ? 0 : 0);
    const posK = (k) => (g.derechaAIzquierda ? n - 1 - k : k);
    const yDe = (v, a, z) => (z > a ? y0 + my + (z - v) / (z - a) * H : y0 + my + H / 2);

    ctx.save();
    if (g.tipo === 'column') {
      const a = Math.min(lo, 0), z = Math.max(hi, 0);
      const yCero = yDe(0, a, z);
      vals.forEach((v, k) => {
        if (v === null) return;
        const yv = yDe(v, a, z);
        ctx.fillStyle = colorDe(k, serie);
        ctx.fillRect(izq(posK(k)), Math.min(yv, yCero), b, Math.max(Math.abs(yv - yCero), 0));
      });
      if (g.ejeX) { ctx.fillStyle = colorEje; ctx.fillRect(x0 + mx, yCero - u / 2, W, u); }
    } else if (g.tipo === 'stacked') {
      const mitad = H / 2, yC = y0 + my + mitad;
      vals.forEach((v, k) => {
        if (v === null || v === 0) return;
        ctx.fillStyle = colorDe(k, serie);
        ctx.fillRect(izq(posK(k)), v > 0 ? yC - mitad : yC, b, mitad);
      });
      if (g.ejeX) { ctx.fillStyle = colorEje; ctx.fillRect(x0 + mx, yC - u / 2, W, u); }
    } else {
      const xs = (k) => izq(posK(k)) + b / 2;
      const ancho = Math.max(u > 1 ? 0 : 1, g.grosor * u / PT);
      if (g.ejeX && lo < 0 && hi > 0) { const yc = yDe(0, lo, hi); ctx.fillStyle = colorEje; ctx.fillRect(x0 + mx, yc - u / 2, W, u); }
      ctx.strokeStyle = serie; ctx.lineWidth = Math.max(ancho, u > 1 ? 0.5 : 1); ctx.lineJoin = 'round'; ctx.lineCap = 'round';
      ctx.beginPath();
      let abierto = false;
      vals.forEach((v, k) => {
        if (v === null) { if (g.vacias === 'gap') abierto = false; return; }
        const x = xs(k), y = yDe(v, lo, hi);
        if (!abierto) { ctx.moveTo(x, y); abierto = true; } else ctx.lineTo(x, y);
      });
      ctx.stroke();
      const r = Math.max(u > 1 ? 0 : 1.5, 1.15 * u / PT * (g.grosor / 0.75));
      vals.forEach((v, k) => {
        if (v === null) return;
        const especial = colorDe(k, null);
        if (!g.marcadores && !especial) return;
        ctx.fillStyle = especial || marcador;
        ctx.beginPath(); ctx.arc(xs(k), yDe(v, lo, hi), r, 0, Math.PI * 2); ctx.closePath(); ctx.fill();
      });
    }
    ctx.restore();
  }

  globalThis.XlsxMinigraficos = { preparar, dibujar };
})();

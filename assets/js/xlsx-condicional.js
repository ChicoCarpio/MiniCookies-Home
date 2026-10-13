/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Formato condicional de Excel del motor propio (etapa 6): qué formato efectivo le toca a cada celda según las reglas de su hoja.
// Reglas medidas contra Excel real (Range.DisplayFormat; Investigaciones\investigacion-motor-excel\03-mediciones.md, etapa 6):
// comparaciones, texto, vacías y errores, superiores/inferiores, sobre/bajo el promedio, duplicados y únicos, expresiones, escalas de
// color, barras de datos y conjuntos de iconos. Las reglas van por prioridad (número menor = primero); cada propiedad (relleno, color de
// letra, negrita…) la fija la primera regla que la trae y "detener si es verdadero" corta las siguientes. Sin DOM.
// Requiere xlsx-modelo.js, xlsx-estilos.js y xlsx-expresion.js.
//
// Uso:
//   await XlsxCondicional.cargar(libro, hoja);                                  // una vez: carga las hojas que citan las fórmulas (Hoja2!A1)
//   const cond = XlsxCondicional.preparar(libro, hoja, { hoy: serialDeHoy });   // null si la hoja no tiene reglas
//   const r = cond.para(fila, col);   // null o { fuente, relleno, borde, formato, barra, icono } (colores como referencias de estilo)
//   cond.reiniciar();                 // al empezar un cuadro: reinicia el presupuesto de trabajo

(function () {
  const MAX_REGLAS = 2000;          // reglas que se tienen en cuenta por hoja
  const PRESUPUESTO = 400000;       // celdas leídas por cuadro al evaluar expresiones
  const ref = (hex) => ({ tipo: 'rgb', valor: hex.replace('#', '').toUpperCase(), tint: 0 });
  const hexDe = (libro, r) => (r ? libro.estilos.color(r) : null);

  function preparar(libro, hoja, opc) {
    opc = opc || {};
    const bloques = hoja.formatosCondicionales || [];
    if (!bloques.length) return null;
    const C = hoja.celdas;
    const TIPO = LibroXlsx.TIPO;
    const hoy = opc.hoy !== undefined ? opc.hoy : Math.floor((Date.now() - Date.UTC(1899, 11, 30)) / 86400000);
    const reglas = [];
    let orden = 0;
    for (const b of bloques) {
      if (!b.rangos.length) continue;
      for (const r of b.reglas) reglas.push({ r, rangos: b.rangos, origen: { f: b.rangos[0].f0, c: b.rangos[0].c0 }, orden: orden++, datos: null, comp: null });
    }
    if (!reglas.length) return null;
    reglas.sort((a, b) => (a.r.prioridad || 1e9) - (b.r.prioridad || 1e9) || a.orden - b.orden);
    if (reglas.length > MAX_REGLAS) reglas.length = MAX_REGLAS;
    let gasto = 0;

    const enRangos = (rg, f, c) => { for (const q of rg) if (f >= q.f0 && f <= q.f1 && c >= q.c0 && c <= q.c1) return true; return false; };
    function valorDe(f, c) { // valor tal como lo ve una expresión
      const i = C.indiceDe(f, c);
      if (i < 0) return null;
      switch (C.tipo[i]) {
        case TIPO.numero: return C.valor[i];
        case TIPO.texto: return C.valor[i];
        case TIPO.logico: return !!C.valor[i];
        case TIPO.error: return new XlsxExpresion.ErrorExcel(String(C.valor[i]));
        default: return null;
      }
    }
    function valoresRango(f0, c0, f1, c1) { // plano, sin las vacías
      const out = [];
      const area = (f1 - f0 + 1) * (c1 - c0 + 1);
      gasto += Math.min(area, C.n);
      if (gasto > PRESUPUESTO) return out;
      if (area <= 4096) { for (let f = f0; f <= f1; f++) for (let c = c0; c <= c1; c++) { const v = valorDe(f, c); if (v !== null) out.push(v); } return out; }
      for (let i = 0; i < C.n; i++) if (C.fila[i] >= f0 && C.fila[i] <= f1 && C.col[i] >= c0 && C.col[i] <= c1) { const v = valorDe(C.fila[i], C.col[i]); if (v !== null) out.push(v); }
      return out;
    }
    // las demás hojas que citan las fórmulas (XlsxCondicional.cargar las dejó en hoja._condHojas)
    const accesos = new Map();
    function otra(nombre) {
      const k = nombre.toLowerCase();
      if (accesos.has(k)) return accesos.get(k);
      const h2 = hoja._condHojas && hoja._condHojas.get(k);
      let acc = null;
      if (h2) {
        const C2 = h2.celdas;
        const valor = (f, c) => {
          const i = C2.indiceDe(f, c);
          if (i < 0) return null;
          switch (C2.tipo[i]) {
            case TIPO.numero: case TIPO.texto: return C2.valor[i];
            case TIPO.logico: return !!C2.valor[i];
            case TIPO.error: return new XlsxExpresion.ErrorExcel(String(C2.valor[i]));
            default: return null;
          }
        };
        const rango = (f0, c0, f1, c1) => {
          const out = [];
          const area = (f1 - f0 + 1) * (c1 - c0 + 1);
          gasto += Math.min(area, C2.n);
          if (gasto > PRESUPUESTO) return out;
          if (area <= 4096) { for (let f = f0; f <= f1; f++) for (let c = c0; c <= c1; c++) { const v = valor(f, c); if (v !== null) out.push(v); } return out; }
          for (let i = 0; i < C2.n; i++) if (C2.fila[i] >= f0 && C2.fila[i] <= f1 && C2.col[i] >= c0 && C2.col[i] <= c1) { const v = valor(C2.fila[i], C2.col[i]); if (v !== null) out.push(v); }
          return out;
        };
        acc = { valor, rango };
      }
      accesos.set(k, acc);
      return acc;
    }
    const expr = (txt) => { const e = XlsxExpresion.compilar(txt); return e; };
    function evaluarFormula(reg, k, f, c) {
      reg.comp = reg.comp || [];
      if (reg.comp[k] === undefined) reg.comp[k] = reg.r.formulas[k] !== undefined ? expr(reg.r.formulas[k]) : null;
      if (!reg.comp[k]) return XlsxExpresion.ERR.nombre;
      return XlsxExpresion.evaluar(reg.comp[k], { fila: f, col: c, origen: reg.origen, valor: valorDe, rango: valoresRango, otra, hoy, hojaActual: libro.hojas[hoja.indice] && libro.hojas[hoja.indice].nombre, gasto: 0 });
    }
    const verdadero = (v) => v === true || (typeof v === 'number' && v !== 0);

    // Estadísticas de los valores numéricos de los rangos de una regla (una vez por hoja).
    function datosDe(reg) {
      if (reg.datos) return reg.datos;
      const nums = [];
      const cuenta = new Map();
      for (let i = 0; i < C.n; i++) {
        if (C.tipo[i] === TIPO.vacio || !enRangos(reg.rangos, C.fila[i], C.col[i])) continue;
        const t = C.tipo[i], v = C.valor[i];
        if (t === TIPO.numero) nums.push(v);
        if (t === TIPO.numero || t === TIPO.texto || t === TIPO.logico) { const k = t === TIPO.texto ? 's' + String(v).toLowerCase() : t === TIPO.logico ? 'b' + (v ? 1 : 0) : 'n' + v; cuenta.set(k, (cuenta.get(k) || 0) + 1); }
      }
      nums.sort((a, b) => a - b);
      const n = nums.length;
      let suma = 0;
      for (const v of nums) suma += v;
      const media = n ? suma / n : 0;
      let ss = 0;
      for (const v of nums) ss += (v - media) * (v - media);
      reg.datos = { nums, n, min: n ? nums[0] : 0, max: n ? nums[n - 1] : 0, media, desv: n > 1 ? Math.sqrt(ss / (n - 1)) : 0, cuenta };
      return reg.datos;
    }
    function percentil(d, p) { // PERCENTILE.INC
      if (!d.n) return 0;
      const x = (d.n - 1) * Math.min(1, Math.max(0, p));
      const i = Math.floor(x), fr = x - i;
      return i + 1 < d.n ? d.nums[i] + (d.nums[i + 1] - d.nums[i]) * fr : d.nums[i];
    }
    function umbral(reg, d, cfvo, f, c, porDefecto) { // valor de un punto de corte (escalas, barras, iconos)
      const v = cfvo && cfvo.valor !== null && cfvo.valor !== undefined ? parseFloat(cfvo.valor) : NaN;
      switch (cfvo ? cfvo.tipo : porDefecto) {
        case 'min': case 'autoMin': return d.min;
        case 'max': case 'autoMax': return d.max;
        case 'num': return Number.isFinite(v) ? v : d.min;
        case 'percent': return Number.isFinite(v) ? d.min + (d.max - d.min) * v / 100 : d.min;
        case 'percentile': return Number.isFinite(v) ? percentil(d, v / 100) : d.min;
        case 'formula': { const x = evaluarFormula({ r: { formulas: [cfvo.valor] }, origen: reg.origen, comp: null }, 0, f, c); return typeof x === 'number' ? x : d.min; }
        default: return d.min;
      }
    }

    function mezclarColor(a, b, t) { // por canal en RGB: inicio + (fin − inicio)·t con el incremento truncado hacia cero (medido en Excel)
      const p = (h) => [parseInt(h.slice(1, 3), 16), parseInt(h.slice(3, 5), 16), parseInt(h.slice(5, 7), 16)];
      const x = p(a), y = p(b);
      return '#' + x.map((q, i) => (q + Math.trunc((y[i] - q) * t)).toString(16).padStart(2, '0')).join('').toUpperCase();
    }

    // ¿Se cumple la regla para esta celda? → true / false (las escalas, barras e iconos devuelven su efecto aparte).
    function cumple(reg, f, c, i) {
      const r = reg.r;
      const t = i >= 0 ? C.tipo[i] : TIPO.vacio;
      const v = i >= 0 && t === TIPO.numero ? C.valor[i] : null;
      switch (r.tipo) {
        case 'cellIs': {
          const a = evaluarFormula(reg, 0, f, c);
          if (XlsxExpresion.esError(a)) return false;
          const x = valorDe(f, c);
          if (XlsxExpresion.esError(x)) return false;
          const cmp = (y) => XlsxExpresion.comparaValores(x, y);
          switch (r.operador) {
            case 'lessThan': return cmp(a) < 0;
            case 'lessThanOrEqual': return cmp(a) <= 0;
            case 'equal': return cmp(a) === 0;
            case 'notEqual': return cmp(a) !== 0;
            case 'greaterThan': return cmp(a) > 0;
            case 'greaterThanOrEqual': return cmp(a) >= 0;
            case 'between': case 'notBetween': {
              const b = evaluarFormula(reg, 1, f, c);
              if (XlsxExpresion.esError(b)) return false;
              const dentro = cmp(a) >= 0 && cmp(b) <= 0;
              return r.operador === 'between' ? dentro : !dentro;
            }
            default: return false;
          }
        }
        case 'top10': {
          if (t !== TIPO.numero) return false;
          const d = datosDe(reg);
          if (!d.n) return false;
          let k = r.rango !== null && r.rango !== undefined ? r.rango : 10;
          if (r.porcentaje) k = Math.max(1, Math.floor(d.n * k / 100));
          k = Math.min(Math.max(1, k), d.n);
          return r.abajo ? v <= d.nums[k - 1] : v >= d.nums[d.n - k];
        }
        case 'aboveAverage': {
          if (t !== TIPO.numero) return false;
          const d = datosDe(reg);
          if (!d.n) return false;
          const s = r.desvEst ? r.desvEst * d.desv : 0;
          if (r.encima) return r.igual ? v >= d.media + s : v > d.media + s;
          return r.igual ? v <= d.media - s : v < d.media - s;
        }
        case 'duplicateValues': case 'uniqueValues': {
          if (t !== TIPO.numero && t !== TIPO.texto && t !== TIPO.logico) return false;
          const k = t === TIPO.texto ? 's' + String(C.valor[i]).toLowerCase() : t === TIPO.logico ? 'b' + (C.valor[i] ? 1 : 0) : 'n' + C.valor[i];
          const n = datosDe(reg).cuenta.get(k) || 0;
          return r.tipo === 'duplicateValues' ? n > 1 : n === 1;
        }
        case 'colorScale': case 'dataBar': case 'iconSet': return t === TIPO.numero;
        default: { // expression, containsText, containsBlanks, timePeriod… : Excel guarda la fórmula
          if (r.formulas && r.formulas.length) return verdadero(evaluarFormula(reg, 0, f, c));
          return false;
        }
      }
    }

    function color(libro2, rf) { return hexDe(libro2, rf); }

    function efectoVisual(reg, f, c, i) { // escala / barra / icono para un valor numérico
      const r = reg.r, v = C.valor[i], d = datosDe(reg);
      if (r.escala) {
        const cv = r.escala.cfvo, cs = r.escala.colores.map((x) => hexDe(libro, x));
        if (cv.length < 2 || cs.length < cv.length || cs.some((x) => !x)) return null;
        const pts = cv.map((q, k) => umbral(reg, d, q, f, c, k === 0 ? 'min' : 'max'));
        if (v <= pts[0]) return { escala: cs[0] };
        for (let k = 1; k < pts.length; k++) {
          if (v <= pts[k] || k === pts.length - 1) {
            if (v >= pts[k]) return { escala: cs[k] };
            const ancho = pts[k] - pts[k - 1];
            return { escala: ancho > 0 ? mezclarColor(cs[k - 1], cs[k], (v - pts[k - 1]) / ancho) : cs[k] };
          }
        }
        return null;
      }
      if (r.barras) {
        const b = r.barras, cv = b.cfvo;
        const lo = umbral(reg, d, cv[0], f, c, 'min'), hi = umbral(reg, d, cv[1], f, c, 'max');
        const fr = hi > lo ? Math.min(1, Math.max(0, (v - lo) / (hi - lo))) : 0;
        const color = hexDe(libro, b.color) || '#638EC6';
        const barra = { fraccion: fr, largoMin: b.min, largoMax: b.max, color, mostrarValor: b.mostrarValor, valor: v, lo, hi };
        const x = b.x14; // extensión de Excel 2010: eje, negativos, relleno liso, borde y dirección
        if (x) {
          const borde = hexDe(libro, x.colorBorde) || color;
          barra.x14 = {
            largoMin: x.min, largoMax: x.max, eje: x.eje, direccion: x.direccion, gradiente: x.gradiente, borde: x.borde ? borde : null,
            colorNeg: x.negIgual ? color : hexDe(libro, x.colorNeg) || '#FF0000',
            bordeNeg: x.borde ? (x.negBordeIgual ? borde : hexDe(libro, x.colorBordeNeg) || '#FF0000') : null,
            colorEje: hexDe(libro, x.colorEje) || '#000000',
          };
          barra.largoMin = x.min; barra.largoMax = x.max; barra.mostrarValor = x.mostrarValor !== false && b.mostrarValor;
        }
        return { barra };
      }
      if (r.iconos) {
        const ic = r.iconos, cv = ic.cfvo;
        const pts = cv.map((q, k) => (k === 0 ? -Infinity : umbral(reg, d, q, f, c, 'num')));
        let k = 0;
        for (let j = 1; j < pts.length; j++) { const gte = cv[j].gte !== false; if (gte ? v >= pts[j] : v > pts[j]) k = j; }
        const nivel = ic.invertir ? cv.length - 1 - k : k;
        const propio = ic.propios && ic.propios.length ? ic.propios[nivel] : null; // iconos personalizados (x14): cada nivel con su conjunto y su número, o sin icono ('NoIcons')
        if (propio) return { icono: { conjunto: propio.conjunto, indice: propio.id, de: cv.length, mostrarValor: ic.mostrarValor } };
        return { icono: { conjunto: ic.conjunto, indice: nivel, de: cv.length, mostrarValor: ic.mostrarValor } };
      }
      return null;
    }

    const memo = new Map();
    function calcular(f, c) {
      const i = C.indiceDe(f, c);
      let res = null;
      const fijo = { relleno: false, color: false, negrita: false, cursiva: false, subrayado: false, tachado: false, formato: false, izq: false, der: false, arr: false, aba: false };
      for (const reg of reglas) {
        if (!enRangos(reg.rangos, f, c)) continue;
        const visual = !!(reg.r.escala || reg.r.barras || reg.r.iconos);
        if (!cumple(reg, f, c, i)) continue;
        res = res || { fuente: null, relleno: null, borde: null, formato: null, barra: null, icono: null };
        if (visual) {
          const e = efectoVisual(reg, f, c, i);
          if (e && e.escala && !fijo.relleno) { res.relleno = { patron: 'solid', fg: ref(e.escala), bg: null }; fijo.relleno = true; }
          if (e && e.barra && !res.barra) res.barra = e.barra;
          if (e && e.icono && !res.icono) res.icono = e.icono;
        } else if (reg.r.dxfXml || (reg.r.dxf !== null && reg.r.dxf !== undefined)) {
          const dx = reg.r.dxfXml ? libro.estilos.dxfDeXml(reg.r.dxfXml) : libro.estilos.dxf(reg.r.dxf);
          if (dx) {
            const fu = dx.fuente;
            if (fu) {
              res.fuente = res.fuente || {};
              const dado = fu.dado || new Set();
              if (fu.color && !fijo.color) { res.fuente.color = fu.color; fijo.color = true; }
              if (dado.has('b') && !fijo.negrita) { res.fuente.negrita = fu.negrita; fijo.negrita = true; }
              if (dado.has('i') && !fijo.cursiva) { res.fuente.cursiva = fu.cursiva; fijo.cursiva = true; }
              if (dado.has('u') && !fijo.subrayado) { res.fuente.subrayado = fu.subrayado === 'none' ? null : fu.subrayado; fijo.subrayado = true; }
              if (dado.has('strike') && !fijo.tachado) { res.fuente.tachado = fu.tachado; fijo.tachado = true; }
            }
            if (dx.relleno && dx.relleno.tipo !== 'ninguno' && !fijo.relleno) {
              const rl = dx.relleno;
              // en un formato diferencial el color de un relleno liso es el de fondo (bgColor); en una trama, el de primer plano es la tinta
              res.relleno = rl.patron === 'solid' ? { patron: 'solid', fg: rl.bg || rl.fg, bg: null } : { patron: rl.patron, fg: rl.fg, bg: rl.bg };
              fijo.relleno = true;
            }
            if (dx.borde) for (const [lado, ya] of [['izq', 'izq'], ['der', 'der'], ['arr', 'arr'], ['aba', 'aba']]) {
              if (dx.borde[lado] && !fijo[ya]) { res.borde = res.borde || {}; res.borde[lado] = dx.borde[lado]; fijo[ya] = true; }
            }
            if (dx.formato && !fijo.formato) { res.formato = dx.formato.codigo !== undefined ? dx.formato.codigo : libro.estilos.formatoDe(dx.formato.id); fijo.formato = true; }
          }
        }
        if (reg.r.detener) break;
      }
      return res;
    }

    return {
      reiniciar() { gasto = 0; memo.clear(); },
      para(f, c) {
        const k = f * 16384 + c;
        if (memo.has(k)) return memo.get(k);
        const r = calcular(f, c);
        if (memo.size < 200000) memo.set(k, r);
        return r;
      },
      color,
      reglas: reglas.length,
    };
  }

  // Aplica el efecto de una celda a su estilo: devuelve un xf nuevo (memoizado por el llamador) con letra, relleno y bordes cambiados.
  function aplicar(xf, efecto) {
    if (!efecto) return xf;
    const r = Object.assign({}, xf);
    if (efecto.fuente) {
      const fu = Object.assign({}, xf.fuente);
      const e = efecto.fuente;
      if (e.color !== undefined) fu.color = e.color;
      if (e.negrita !== undefined) fu.negrita = e.negrita;
      if (e.cursiva !== undefined) fu.cursiva = e.cursiva;
      if (e.subrayado !== undefined) fu.subrayado = e.subrayado;
      if (e.tachado !== undefined) fu.tachado = e.tachado;
      r.fuente = fu;
    }
    if (efecto.relleno) r.relleno = { tipo: 'patron', patron: efecto.relleno.patron, fg: efecto.relleno.fg, bg: efecto.relleno.bg, degradado: null };
    if (efecto.borde) r.borde = Object.assign({ izq: null, der: null, arr: null, aba: null, diagonal: null, diagArriba: false, diagAbajo: false }, xf.borde || {}, efecto.borde);
    if (efecto.formato) r.formato = efecto.formato;
    return r;
  }

  // Carga las hojas que citan las fórmulas de las reglas (Hoja2!$A$1: el dibujo es síncrono): hoja._condHojas = Map(nombre en minúsculas → hoja).
  async function cargar(libro, hoja) {
    if (hoja._condHojas) return;
    hoja._condHojas = new Map();
    const propia = libro.hojas[hoja.indice] && libro.hojas[hoja.indice].nombre;
    const pedidas = new Set();
    const buscar = (txt) => {
      if (typeof txt !== 'string' || txt.indexOf('!') < 0) return;
      for (const m of txt.matchAll(/(?:'((?:[^']|'')+)'|([A-Za-z_\u00C0-\uFFFF][\w.\u00C0-\uFFFF]*))!/g)) {
        const nom = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
        if (nom && nom !== propia) pedidas.add(nom.toLowerCase());
      }
    };
    for (const b of hoja.formatosCondicionales || []) {
      for (const r of b.reglas) {
        for (const f of r.formulas || []) buscar(f);
        for (const k of [r.escala, r.barras, r.iconos]) if (k) for (const v of k.cfvo || []) buscar(v.valor);
      }
    }
    let n = 0;
    for (const nombre of pedidas) {
      if (++n > 50) break;
      const i = libro.hojas.findIndex((h) => h.nombre.toLowerCase() === nombre);
      if (i < 0) continue;
      try { hoja._condHojas.set(nombre, await libro.cargarHoja(i)); } catch (e) { /* una hoja que no se lee: las fórmulas que la citan dan #REF! */ }
    }
  }

  globalThis.XlsxCondicional = { cargar, preparar, aplicar };
})();

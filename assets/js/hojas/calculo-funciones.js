/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Más funciones para el motor de Mini Cookies Hojas (fase 3.1 del PLAN): trigonometría hiperbólica, combinatoria, bases numéricas y bits,
// estadística (correlación, regresión, asimetría…), distribuciones (normal, t, chi², F, gamma, beta, binomial, Poisson…), finanzas (TASA,
// TIR, VNA.NO.PER, PAGOINT, amortizaciones…), texto (MONEDA, DIVIDIRTEXTO, REGEX…), referencias (INDIRECTO, DIRECCION, AREAS, HOJA) y
// bases de datos (BDSUMA…). Se suma al motor con CalculoHojas.ampliar; va después de calculo.js. Sin DOM.
// Cada función recibe (args sin evaluar, ctx, A) como las de calculo.js; A es la API del motor (num, txt, matrizDe, E, lanzar…).
(function () {
  CalculoHojas.ampliar((U) => {
    const { numeros, sumar, criterio, varianza, desviacion, redondear, truncar, esError, textoFormato } = U;
    const mat = (filas) => ({ tipo: 'mat', filas });
    const ent = (a, c, A) => Math.trunc(A.num(a, c));
    const opc = (a, c, A, d) => A.opcional(a, c, d, A.num);
    const esNum = (x) => typeof x === 'number';
    const NUMERO = (x, A) => (Number.isFinite(x) ? x : A.E.num);

    // ---- funciones especiales: gamma, beta y error -------------------------------------------------------------------
    const LANCZOS = [0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059, 12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7];
    function lgamma(x) {
      if (x < 0.5) return Math.log(Math.PI / Math.abs(Math.sin(Math.PI * x))) - lgamma(1 - x);
      x -= 1;
      let s = LANCZOS[0];
      for (let i = 1; i < 9; i++) s += LANCZOS[i] / (x + i);
      const t = x + 7.5;
      return 0.5 * Math.log(2 * Math.PI) + (x + 0.5) * Math.log(t) - t + Math.log(s);
    }
    function gamma(x) {
      if (Number.isInteger(x) && x > 0 && x < 171) { let r = 1; for (let i = 2; i < x; i++) r *= i; return r; }
      if (x < 0.5) return Math.PI / (Math.sin(Math.PI * x) * gamma(1 - x));
      return Math.exp(lgamma(x));
    }
    // gamma incompleta regularizada P(a, x) (serie o fracción continua de Lentz)
    function gammaP(a, x) {
      if (x <= 0) return 0;
      if (x < a + 1) {
        let suma = 1 / a, del = suma, ap = a;
        for (let n = 0; n < 1000; n++) { ap++; del *= x / ap; suma += del; if (Math.abs(del) < Math.abs(suma) * 1e-16) break; }
        return suma * Math.exp(-x + a * Math.log(x) - lgamma(a));
      }
      return 1 - gammaQcf(a, x);
    }
    function gammaQcf(a, x) {
      const MIN = 1e-300;
      let b = x + 1 - a, c = 1 / MIN, d = 1 / b, h = d;
      for (let i = 1; i < 1000; i++) {
        const an = -i * (i - a); b += 2;
        d = an * d + b; if (Math.abs(d) < MIN) d = MIN;
        c = b + an / c; if (Math.abs(c) < MIN) c = MIN;
        d = 1 / d; const del = d * c; h *= del;
        if (Math.abs(del - 1) < 1e-16) break;
      }
      return Math.exp(-x + a * Math.log(x) - lgamma(a)) * h;
    }
    const gammaQ = (a, x) => (x <= 0 ? 1 : x < a + 1 ? 1 - gammaP(a, x) : gammaQcf(a, x));
    // beta incompleta regularizada I_x(a, b)
    function betacf(a, b, x) {
      const MIN = 1e-300;
      let qab = a + b, qap = a + 1, qam = a - 1, c = 1, d = 1 - qab * x / qap;
      if (Math.abs(d) < MIN) d = MIN;
      d = 1 / d; let h = d;
      for (let m = 1; m <= 1000; m++) {
        const m2 = 2 * m;
        let aa = m * (b - m) * x / ((qam + m2) * (a + m2));
        d = 1 + aa * d; if (Math.abs(d) < MIN) d = MIN; c = 1 + aa / c; if (Math.abs(c) < MIN) c = MIN; d = 1 / d; h *= d * c;
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
        d = 1 + aa * d; if (Math.abs(d) < MIN) d = MIN; c = 1 + aa / c; if (Math.abs(c) < MIN) c = MIN; d = 1 / d;
        const del = d * c; h *= del;
        if (Math.abs(del - 1) < 1e-16) break;
      }
      return h;
    }
    function betaI(x, a, b) {
      if (x <= 0) return 0;
      if (x >= 1) return 1;
      const bt = Math.exp(lgamma(a + b) - lgamma(a) - lgamma(b) + a * Math.log(x) + b * Math.log(1 - x));
      return x < (a + 1) / (a + b + 2) ? bt * betacf(a, b, x) / a : 1 - bt * betacf(b, a, 1 - x) / b;
    }
    // erf y erfc con precisión doble: serie de Taylor cerca de 0, fracción continua lejos
    function erf(x) {
      if (Math.abs(x) >= 3) return Math.sign(x) * (1 - erfcGrande(Math.abs(x)));
      let suma = x, t = x, n = 0;
      const x2 = x * x;
      while (Math.abs(t) > 1e-17 * Math.abs(suma) && n < 200) { n++; t *= -x2 / n; suma += t / (2 * n + 1); }
      return 2 / Math.sqrt(Math.PI) * suma;
    }
    function erfcGrande(x) {   // x ≥ 3: fracción continua de Laplace (Lentz)
      const MIN = 1e-300;
      let f = x, c = x, d = 0;
      for (let i = 1; i < 500; i++) {
        const an = i / 2;
        d = x + an * d; if (Math.abs(d) < MIN) d = MIN; d = 1 / d;
        c = x + an / c; if (Math.abs(c) < MIN) c = MIN;
        const del = c * d; f *= del;
        if (Math.abs(del - 1) < 1e-16) break;
      }
      return Math.exp(-x * x) / (f * Math.sqrt(Math.PI));
    }
    const erfc = (x) => (x >= 3 ? erfcGrande(x) : 1 - erf(x));
    const normalCdf = (z) => (z < 0 ? 0.5 * erfc(-z / Math.SQRT2) : 1 - 0.5 * erfc(z / Math.SQRT2));
    const normalPdf = (z) => Math.exp(-z * z / 2) / Math.sqrt(2 * Math.PI);
    // inversa de la normal estándar: Acklam y un paso de Halley
    function normalInv(p) {
      const a = [-3.969683028665376e+01, 2.209460984245205e+02, -2.759285104469687e+02, 1.383577518672690e+02, -3.066479806614716e+01, 2.506628277459239e+00];
      const b = [-5.447609879822406e+01, 1.615858368580409e+02, -1.556989798598866e+02, 6.680131188771972e+01, -1.328068155288572e+01];
      const c = [-7.784894002430293e-03, -3.223964580411365e-01, -2.400758277161838e+00, -2.549732539343734e+00, 4.374664141464968e+00, 2.938163982698783e+00];
      const d = [7.784695709041462e-03, 3.224671290700398e-01, 2.445134137142996e+00, 3.754408661907416e+00];
      let x;
      if (p < 0.02425) { const q = Math.sqrt(-2 * Math.log(p)); x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1); }
      else if (p > 1 - 0.02425) { const q = Math.sqrt(-2 * Math.log(1 - p)); x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1); }
      else { const q = p - 0.5, r = q * q; x = (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1); }
      for (let k = 0; k < 2; k++) { const e = normalCdf(x) - p, u = e * Math.sqrt(2 * Math.PI) * Math.exp(x * x / 2); x -= u / (1 + x * u / 2); }
      return x;
    }
    // inversa de una distribución acumulada creciente por bisección (lo, hi: intervalo que la contiene)
    function inversa(cdf, p, lo, hi) {
      while (cdf(hi) < p && hi < 1e300) { lo = hi; hi *= 2; }
      for (let i = 0; i < 2000; i++) {
        const m = (lo + hi) / 2;
        if (m === lo || m === hi) break;
        if (cdf(m) < p) lo = m; else hi = m;
      }
      return (lo + hi) / 2;
    }
    const tCdf = (x, n) => { const ib = betaI(n / (n + x * x), n / 2, 0.5); return x >= 0 ? 1 - ib / 2 : ib / 2; };
    const tPdf = (x, n) => Math.exp(lgamma((n + 1) / 2) - lgamma(n / 2) - (n + 1) / 2 * Math.log(1 + x * x / n)) / Math.sqrt(n * Math.PI);
    const tInv = (p, n) => { if (p === 0.5) return 0; const x = inversa((y) => tCdf(y, n), p > 0.5 ? p : 1 - p, 0, 1); return p > 0.5 ? x : -x; };
    const chiCdf = (x, k) => gammaP(k / 2, x / 2);
    const fCdf = (x, d1, d2) => (x <= 0 ? 0 : betaI(d1 * x / (d1 * x + d2), d1 / 2, d2 / 2));
    const combin = (n, k) => { if (k < 0 || k > n) return 0; k = Math.min(k, n - k); let r = 1; for (let i = 1; i <= k; i++) r = r * (n - k + i) / i; return Math.round(r) === r || r > 1e15 ? r : Math.round(r); };
    const lcombin = (n, k) => lgamma(n + 1) - lgamma(k + 1) - lgamma(n - k + 1);
    const binomPmf = (k, n, p) => (p === 0 ? (k === 0 ? 1 : 0) : p === 1 ? (k === n ? 1 : 0) : Math.exp(lcombin(n, k) + k * Math.log(p) + (n - k) * Math.log(1 - p)));

    // parámetros de las distribuciones: los grados de libertad se truncan; < 1 da #NUM!
    const gl = (a, c, A, max) => { const n = Math.trunc(A.num(a, c)); if (n < 1 || (max && n > max)) A.lanzar(A.E.num); return n; };
    const prob = (a, c, A, abierta) => { const p = A.num(a, c); if (p < 0 || p > 1 || (abierta && (p === 0 || p === 1))) A.lanzar(A.E.num); return p; };

    // ---- pares de rangos (correlación, regresión, SUMX2MY2…) -----------------------------------------------------------
    function pares(ay, ax, c, A, mismoTamano) {
      const Y = A.matrizDe(A.evaluarNodo(ay, c)), X = A.matrizDe(A.evaluarNodo(ax, c));
      const ly = Y.flat(), lx = X.flat();
      if (ly.length !== lx.length) A.lanzar(A.E.na);
      if (mismoTamano && (Y.length !== X.length)) A.lanzar(A.E.na);
      const xs = [], ys = [];
      for (let i = 0; i < ly.length; i++) {
        if (esError(ly[i])) A.lanzar(ly[i]);
        if (esError(lx[i])) A.lanzar(lx[i]);
        if (esNum(ly[i]) && esNum(lx[i])) { ys.push(ly[i]); xs.push(lx[i]); }
      }
      return { xs, ys };
    }
    function momentos(xs, ys) {
      const n = xs.length, mx = sumar(xs) / n, my = sumar(ys) / n;
      let sxx = 0, syy = 0, sxy = 0;
      for (let i = 0; i < n; i++) { const dx = xs[i] - mx, dy = ys[i] - my; sxx += dx * dx; syy += dy * dy; sxy += dx * dy; }
      return { n, mx, my, sxx, syy, sxy };
    }
    const regresion = (a, c, A) => { const { xs, ys } = pares(a[0], a[1], c, A); if (xs.length < 1) A.lanzar(A.E.div); const m = momentos(xs, ys); if (m.sxx === 0) A.lanzar(A.E.div); return m; };
    const correl = (a, c, A) => { const { xs, ys } = pares(a[0], a[1], c, A); const m = momentos(xs, ys); if (!m.n || m.sxx === 0 || m.syy === 0) return A.E.div; return m.sxy / Math.sqrt(m.sxx * m.syy); };
    const sumaPares = (fn) => (a, c, A) => { const { xs, ys } = pares(a[0], a[1], c, A, true); let s = 0; for (let i = 0; i < xs.length; i++) s += fn(ys[i], xs[i]); return s; };

    // ---- bases numéricas -----------------------------------------------------------------------------------------------
    // DEC2BIN y compañía: 10 dígitos como máximo; los negativos en complemento a dos de 10 dígitos
    const BITS = { 2: 10, 8: 30, 16: 40 };
    function decA(base) {
      const lim = Math.pow(2, BITS[base] - 1);
      return (a, c, A) => {
        const n = Math.trunc(A.num(a[0], c));
        if (n < -lim || n >= lim) return A.E.num;
        const lugares = a[1] === undefined || a[1] === null ? null : Math.trunc(A.num(a[1], c));
        if (n < 0) return (Math.pow(2, BITS[base]) + n).toString(base).toUpperCase();
        let s = n.toString(base).toUpperCase();
        if (lugares !== null) { if (lugares < 0 || lugares < s.length || lugares > 10) return A.E.num; s = s.padStart(lugares, '0'); }
        return s;
      };
    }
    function aDec(base, s, A) {
      s = String(s).trim().toUpperCase();
      if (s.length > 10) A.lanzar(A.E.num);
      if (s === '') return 0;
      const re = base === 2 ? /^[01]+$/ : base === 8 ? /^[0-7]+$/ : /^[0-9A-F]+$/;
      if (!re.test(s)) A.lanzar(A.E.num);
      let n = parseInt(s, base);
      if (s.length === 10 && n >= Math.pow(2, BITS[base] - 1)) n -= Math.pow(2, BITS[base]);
      return n;
    }
    const textoDeBase = (a, c, A) => { const v = A.escalar(A.evaluarNodo(a, c), c); if (esError(v)) A.lanzar(v); if (typeof v === 'boolean') A.lanzar(A.E.valor); return typeof v === 'number' ? String(v) : A.aTexto(v); };
    const deBase = (base) => (a, c, A) => aDec(base, textoDeBase(a[0], c, A), A);
    const entreBases = (desde, hasta) => (a, c, A) => { const n = aDec(desde, textoDeBase(a[0], c, A), A); return decA(hasta)([{ t: 'num', v: n }, a[1]], c, A); };
    const bit = (a, c, A) => { const n = A.num(a, c); if (n < 0 || n >= 281474976710656 || n !== Math.trunc(n)) A.lanzar(A.E.num); return n; };
    const bitOp = (op) => (a, c, A) => {
      const x = bit(a[0], c, A), y = bit(a[1], c, A);
      const xh = Math.floor(x / 16777216), xl = x % 16777216, yh = Math.floor(y / 16777216), yl = y % 16777216;
      return op(xh, yh) * 16777216 + op(xl, yl);
    };

    // ---- romanos -------------------------------------------------------------------------------------------------------
    const ROMANOS = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
    const VALOR_ROMANO = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };

    // ---- finanzas ------------------------------------------------------------------------------------------------------
    const pago = (r, n, pv, fv, t) => { if (r === 0) return -(pv + fv) / n; const p = Math.pow(1 + r, n); return -(r * (pv * p + fv)) / ((1 + r * t) * (p - 1)); };
    const vf = (r, n, pmt, pv, t) => { if (r === 0) return -(pv + pmt * n); const p = Math.pow(1 + r, n); return -(pv * p + pmt * (1 + r * t) * (p - 1) / r); };
    function interes(r, per, n, pv, fv, t) {
      const p = pago(r, n, pv, fv, t);
      let i;
      if (per === 1) i = t === 1 ? 0 : -pv;
      else i = t === 1 ? vf(r, per - 2, p, pv, 1) - p : vf(r, per - 1, p, pv, 0);
      return i * r;
    }
    // raíz de f cerca de x0 por Newton (con derivada numérica), como TASA, TIR y TIR.NO.PER
    function newton(f, x0, A, escala) {
      let x = x0;
      const tol = 1e-9 * Math.max(1, escala || 0);
      for (let i = 0; i < 200; i++) {
        const y = f(x);
        if (!Number.isFinite(y)) break;
        const h = Math.max(1e-7, Math.abs(x) * 1e-7), d = (f(x + h) - f(x - h)) / (2 * h);
        if (!d || !Number.isFinite(d)) break;
        const nx = x - y / d;
        if (!Number.isFinite(nx)) break;
        if (Math.abs(nx - x) < 1e-12 * Math.max(1, Math.abs(x)) || Math.abs(f(nx)) < tol * 1e-6) return Math.abs(f(nx)) < tol ? nx : A.E.num;
        x = nx;
      }
      return A.E.num;
    }
    const flujos = (a, c, A) => { const l = []; A.recorrerValor(A.evaluarNodo(a, c), (v, r) => { if (esError(v)) A.lanzar(v); if (esNum(v)) l.push(v); else if (!r && v !== null) l.push(A.aNumero(v)); }); return l; };

    // ---- bases de datos: BDSUMA(base; campo; criterios) ------------------------------------------------------------------
    function registros(a, c, A) {
      const B = A.matrizDe(A.evaluarNodo(a[0], c)), K = A.matrizDe(A.evaluarNodo(a[2], c));
      if (B.length < 1 || K.length < 1) A.lanzar(A.E.valor);
      const enc = B[0].map((x) => (x === null ? '' : String(x).toUpperCase()));
      let col = null;
      if (a[1] !== undefined && a[1] !== null) {
        const v = A.escalar(A.evaluarNodo(a[1], c), c);
        if (esError(v)) A.lanzar(v);
        if (typeof v === 'number') { col = Math.trunc(v) - 1; if (col < 0 || col >= enc.length) A.lanzar(A.E.valor); }
        else { col = enc.indexOf(String(v).toUpperCase()); if (col < 0) A.lanzar(A.E.valor); }
      }
      // cada fila de criterios es un O; dentro de la fila, Y. Un texto sin operador vale como «empieza por»
      const encK = K[0].map((x) => (x === null ? null : String(x).toUpperCase()));
      const filasK = [];
      for (let i = 1; i < K.length; i++) {
        const conds = [];
        for (let j = 0; j < encK.length; j++) {
          const v = K[i][j];
          if (v === null || v === '') continue;
          const k = enc.indexOf(encK[j]);
          if (k < 0) throw new A.NoSoportada('criterio calculado de base de datos');
          const cr = typeof v === 'string' && !/^(<=|>=|<>|<|>|=)/.test(v) && A.textoANumero(v) === null ? criterio(v + '*', A) : typeof v === 'string' && /^=/.test(v) && v.length > 1 && A.textoANumero(v.slice(1)) === null ? criterio(v.slice(1), A) : criterio(v, A);
          conds.push([k, cr]);
        }
        filasK.push(conds);
      }
      const salida = [];
      for (let i = 1; i < B.length; i++) if (!filasK.length || filasK.some((conds) => conds.every(([k, cr]) => cr(B[i][k])))) salida.push(col === null ? B[i] : B[i][col]);
      return { valores: salida, col };
    }
    const bd = (fn) => (a, c, A) => { const { valores } = registros(a, c, A); return fn(valores.filter(esNum), valores, A); };

    // ---- texto ---------------------------------------------------------------------------------------------------------
    function horaDeTexto(t, A) {
      const s = String(t).trim();
      const m = /(\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:[.,]\d+)?))?\s*([AaPp]\.?\s*[Mm]\.?)?\s*$/.exec(s);
      if (!m) {
        const n = A.textoANumero(s);
        if (n !== null || /\d{1,4}[/-]\d{1,2}[/-]\d{1,4}/.test(s)) return 0;
        A.lanzar(A.E.valor);
      }
      let h = +m[1]; const mi = +m[2], se = m[3] ? parseFloat(m[3].replace(',', '.')) : 0;
      if (m[4]) { if (h > 12) A.lanzar(A.E.valor); const pm = /^[Pp]/.test(m[4]); if (h === 12) h = 0; if (pm) h += 12; }
      if (mi > 59 || se >= 60) A.lanzar(A.E.valor);
      return ((h * 3600 + mi * 60 + se) / 86400) % 1;
    }
    function expresion(patron, sinMay, A, global) {
      try { return new RegExp(patron, (global ? 'g' : '') + (sinMay ? 'i' : '') + 'u'); } catch { A.lanzar(A.E.valor); }
    }
    const textoDeValor = (v, estricto, A) => {
      if (v === null) return '';
      if (esError(v)) return v.codigo;
      if (typeof v === 'string') return estricto ? '"' + v.replace(/"/g, '""') + '"' : v;
      if (typeof v === 'boolean' && A.fuente.idioma !== 'en') return v ? 'VERDADERO' : 'FALSO';   // el nombre local del booleano (es-CR), como Excel
      return A.aTexto(v);
    };

    // ---- matrices ------------------------------------------------------------------------------------------------------
    const numerica = (a, c, A) => { const M = A.matrizDe(A.evaluarNodo(a, c)); for (const f of M) for (const x of f) { if (esError(x)) A.lanzar(x); if (!esNum(x)) A.lanzar(A.E.valor); } return M; };
    function determinante(M) {
      const n = M.length, B = M.map((f) => f.slice());
      let d = 1;
      for (let i = 0; i < n; i++) {
        let p = i; for (let k = i + 1; k < n; k++) if (Math.abs(B[k][i]) > Math.abs(B[p][i])) p = k;
        if (B[p][i] === 0) return 0;
        if (p !== i) { [B[p], B[i]] = [B[i], B[p]]; d = -d; }
        d *= B[i][i];
        for (let k = i + 1; k < n; k++) { const f = B[k][i] / B[i][i]; for (let j = i; j < n; j++) B[k][j] -= f * B[i][j]; }
      }
      return d;
    }
    function inversaMatriz(M) {
      const n = M.length, B = M.map((f, i) => f.concat(Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))));
      for (let i = 0; i < n; i++) {
        let p = i; for (let k = i + 1; k < n; k++) if (Math.abs(B[k][i]) > Math.abs(B[p][i])) p = k;
        if (Math.abs(B[p][i]) < 1e-300) return null;
        [B[p], B[i]] = [B[i], B[p]];
        const piv = B[i][i]; for (let j = 0; j < 2 * n; j++) B[i][j] /= piv;
        for (let k = 0; k < n; k++) if (k !== i) { const f = B[k][i]; if (f) for (let j = 0; j < 2 * n; j++) B[k][j] -= f * B[i][j]; }
      }
      return B.map((f) => f.slice(n));
    }

    // ---- referencias ---------------------------------------------------------------------------------------------------
    const citarHoja = (n) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) && !/^R\d*C\d*$/i.test(n) ? n : "'" + n.replace(/'/g, "''") + "'");
    const letras = (c) => { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - 1 - m) / 26; } return s; };
    // R1C1 (F1C1 en español, como Excel: ahí R2C1 da #REF!) → A1, relativo a la celda de la fórmula
    function r1c1aA1(t, ctx, letraFila) {
      const parte = (s) => {
        const m = new RegExp('^' + letraFila + '(\\[-?\\d+\\]|\\d+)?C(\\[-?\\d+\\]|\\d+)?$', 'i').exec(s);
        if (!m) return null;
        const eje = (x, base) => (x === undefined ? base : x[0] === '[' ? base + parseInt(x.slice(1), 10) : parseInt(x, 10) - 1);
        const f = eje(m[1], ctx.f), c = eje(m[2], ctx.c);
        if (f < 0 || c < 0 || f > 1048575 || c > 16383) return null;
        return letras(c) + (f + 1);
      };
      const k = t.lastIndexOf('!');
      const hoja = k >= 0 ? t.slice(0, k + 1) : '', resto = k >= 0 ? t.slice(k + 1) : t;
      const ps = resto.split(':').map(parte);
      if (ps.some((x) => x === null) || ps.length > 2) return null;
      return hoja + ps.join(':');
    }

    // ---- matrices dinámicas -------------------------------------------------------------------------------------------
    const matrizArg = (nodo, c, A) => A.matrizDe(A.evaluarNodo(nodo, Object.assign({}, c, { matriz: true })));
    // ---- regresión (ESTIMACION.LINEAL y compañía), medida en Excel: pruebas/oraculo/regresion-casos.mjs -----------------------------------------
    // y en columna: cada columna de x es una variable; y en fila: cada fila; y de dos dimensiones: se aplana por filas y hay una sola variable.
    // Sin x, 1, 2, 3… Un tamaño que no coincide da #REF!; una celda vacía, un texto o un error, #VALUE!.
    function datosRegresion(a, c, A, logaritmo) {
      const Y0 = matrizArg(a[0], c, A), fY = Y0.length, cY = Y0[0].length, n = fY * cY;
      const plana = fY > 1 && cY > 1, enFilas = fY === 1 && cY > 1;
      let xs, k = 1;
      if (a[1] === undefined || a[1] === null) xs = Array.from({ length: n }, (_, i) => [i + 1]);
      else {
        const X0 = matrizArg(a[1], c, A);
        if (plana) {
          if (X0.length !== fY || X0[0].length !== cY) return A.E.ref;
          xs = X0.flat().map((x) => [x]);
        } else if (enFilas) {
          if (X0[0].length !== n) return A.E.ref;
          k = X0.length;
          xs = Array.from({ length: n }, (_, i) => X0.map((f) => f[i]));
        } else {
          if (X0.length !== n) return A.E.ref;
          k = X0[0].length;
          xs = X0;
        }
      }
      const y = Y0.flat();
      if (!y.every(esNum) || !xs.every((f) => f.every(esNum))) return A.E.valor;
      if (logaritmo && y.some((v) => v <= 0)) return A.E.num;
      return { y: logaritmo ? y.map(Math.log) : y, xs, k, n, enFilas, plana, forma: [fY, cY] };
    }
    // Mínimos cuadrados por QR con pivoteo: cada paso toma la columna de mayor norma que queda, y la que ya no aporta (colineal) vale 0, como en Excel.
    function ajustar(y, xs, k, conConstante) {
      const n = y.length;
      const media = (v) => v.reduce((s, x) => s + x, 0) / n;
      const my = conConstante ? media(y) : 0;
      const mx = Array.from({ length: k }, (_, j) => (conConstante ? media(xs.map((f) => f[j])) : 0));
      const Y = conConstante && y.every((v) => v === y[0]) ? y.map(() => 0) : y.map((v) => v - my);   // y constante: el redondeo de la media no cuenta como variación
      const cols = mx.map((m, j) => xs.map((f) => f[j] - m));
      const norma = (v) => Math.sqrt(v.reduce((s, x) => s + x * x, 0));
      const puntoCon = (u, v) => { let s = 0; for (let i = 0; i < n; i++) s += u[i] * v[i]; return s; };
      const original = cols.map(norma);
      const quedan = cols.map((_, j) => j), orden = [], Q = [], Rq = [];
      const maxRango = n - (conConstante ? 1 : 0);
      while (quedan.length && orden.length < maxRango) {
        let mejor = -1, valor = 0;
        for (const j of quedan) { const v = norma(cols[j]); if (v > valor && v > 1e-11 * original[j]) { valor = v; mejor = j; } }
        if (mejor < 0) break;
        quedan.splice(quedan.indexOf(mejor), 1);
        const q = cols[mejor].map((x) => x / valor), fila = new Map([[mejor, valor]]);
        for (const j of quedan) { const r = puntoCon(q, cols[j]); fila.set(j, r); for (let i = 0; i < n; i++) cols[j][i] -= r * q[i]; }
        Q.push(q); Rq.push(fila); orden.push(mejor);
      }
      const r = orden.length, qy = Q.map((q) => puntoCon(q, Y));
      // de R a los coeficientes, y su inversa para los errores
      const coefOrden = new Array(r).fill(0), Rinv = Array.from({ length: r }, () => new Array(r).fill(0));
      for (let s = r - 1; s >= 0; s--) {
        let v = qy[s];
        for (let t = s + 1; t < r; t++) v -= Rq[s].get(orden[t]) * coefOrden[t];
        coefOrden[s] = v / Rq[s].get(orden[s]);
        Rinv[s][s] = 1 / Rq[s].get(orden[s]);
        for (let t = s + 1; t < r; t++) { let w = 0; for (let u = s + 1; u <= t; u++) w += Rq[s].get(orden[u]) * Rinv[u][t]; Rinv[s][t] = -w / Rq[s].get(orden[s]); }
      }
      const coef = new Array(k).fill(0), se = new Array(k).fill(0);
      orden.forEach((j, s) => { coef[j] = coefOrden[s]; });
      const sst = Y.reduce((s, v) => s + v * v, 0), ssreg = sst === 0 ? 0 : qy.reduce((s, v) => s + v * v, 0);
      let ssres = 0;
      if (sst !== 0) for (let i = 0; i < n; i++) { let e = Y[i]; for (let s = 0; s < r; s++) e -= qy[s] * Q[s][i]; ssres += e * e; }
      const df = n - r - (conConstante ? 1 : 0), s2 = df > 0 ? ssres / df : 0;
      const cov = (s, t) => { let w = 0; for (let u = Math.max(s, t); u < r; u++) w += Rinv[s][u] * Rinv[t][u]; return w; };
      orden.forEach((j, s) => { se[j] = Math.sqrt(s2 * cov(s, s)); });
      let varB = s2 / n;
      for (let s = 0; s < r; s++) for (let t = 0; t < r; t++) varB += s2 * mx[orden[s]] * mx[orden[t]] * cov(s, t);
      const b = conConstante ? my - coef.reduce((s, m, j) => s + m * mx[j], 0) : 0;
      return { coef, b, se, seB: Math.sqrt(varB), r2: ssreg + ssres === 0 ? 1 : ssreg / (ssreg + ssres), sey: Math.sqrt(s2), df, ssreg, ssres, F: df > 0 && ssres > 0 && r > 0 ? (ssreg / r) / (ssres / df) : null };
    }
    function estimacion(a, c, A, logaritmo) {
      const conConstante = A.opcional(a[2], c, true, A.log), conEstadisticas = A.opcional(a[3], c, false, A.log);
      const d = datosRegresion(a, c, A, logaritmo);
      if (!d.y) return d;
      const f = ajustar(d.y, d.xs, d.k, conConstante);
      const salida = (v) => (logaritmo ? Math.exp(v) : v);
      const filas = [[...f.coef.map(salida).reverse(), salida(f.b)]];
      if (conEstadisticas) {
        const relleno = new Array(d.k - 1).fill(A.E.na);
        filas.push([...[...f.se].reverse(), conConstante ? f.seB : A.E.na]);
        filas.push([f.r2, f.sey, ...relleno], [f.F === null ? A.E.num : f.F, f.df, ...relleno], [f.ssreg, f.ssres, ...relleno]);
      }
      return mat(filas);
    }
    // TENDENCIA y CRECIMIENTO: sin x nuevas, las de x con la forma de y. Con una variable, cada celda de x nuevas da un valor; con varias, cada fila (o cada columna, si y está en fila).
    function prediccion(a, c, A, logaritmo) {
      const conConstante = A.opcional(a[3], c, true, A.log);
      const d = datosRegresion(a, c, A, logaritmo);
      if (!d.y) return d;
      const f = ajustar(d.y, d.xs, d.k, conConstante);
      const valor = (x) => { const v = x.reduce((s, xi, j) => s + f.coef[j] * xi, f.b); return logaritmo ? Math.exp(v) : v; };
      if (a[2] === undefined || a[2] === null) {
        const v = d.xs.map(valor);
        if (d.enFilas) return mat([v]);
        return mat(d.plana ? Array.from({ length: d.forma[0] }, (_, i) => v.slice(i * d.forma[1], (i + 1) * d.forma[1])) : v.map((x) => [x]));
      }
      const N = matrizArg(a[2], c, A);
      if (!N.every((l) => l.every(esNum))) return A.E.valor;
      if (d.k === 1) return mat(N.map((l) => l.map((x) => valor([x]))));
      if (d.enFilas) { if (N.length !== d.k) return A.E.ref; return mat([N[0].map((_, j) => valor(N.map((l) => l[j])))]); }
      if (N[0].length !== d.k) return A.E.ref;
      return mat(N.map((l) => [valor(l)]));
    }

    // ---- CONVERTIR: los factores son los que da Excel, respecto de la primera unidad de referencia de cada grupo -----------------------------------
    const UNIDADES_CONVERT = [
      { g: 1, sg: 14593.902937206363, lbm: 453.59237, u: 1.660538782e-24, ozm: 28.349523125, grain: 0.06479891, cwt: 45359.237, shweight: 45359.237, uk_cwt: 50802.34544, lcwt: 50802.34544, hweight: 50802.34544, stone: 6350.29318, ton: 907184.74, uk_ton: 1016046.9088, LTON: 1016046.9088, brton: 1016046.9088 },
      { m: 1, mi: 1609.344, Nmi: 1852, in: 0.0254, ft: 0.3048, yd: 0.9144, ang: 1e-10, ell: 1.143, ly: 9460730472580800, parsec: 30856775812815532, pc: 30856775812815532, Picapt: 0.0003527777777777778, Pica: 0.0003527777777777778, pica: 0.004233333333333334, survey_mi: 1609.3472186944373 },
      { yr: 31557600, day: 86400, d: 86400, hr: 3600, mn: 60, min: 60, sec: 1, s: 1 },
      { Pa: 1, p: 1, atm: 101325, at: 101325, mmHg: 133.322, psi: 6894.757293168362, Torr: 133.32236842105263 },
      { N: 1, dyn: 0.00001, dy: 0.00001, lbf: 4.4482216152605005, pond: 0.00980665 },
      { J: 1, e: 1e-7, c: 4.184, cal: 4.1868, eV: 1.602176487e-19, ev: 1.602176487e-19, HPh: 2684519.537696173, hh: 2684519.537696173, Wh: 3600, wh: 3600, flb: 1.3558179483314003, BTU: 1055.05585262, btu: 1055.05585262 },
      { HP: 745.6998715822702, h: 745.6998715822702, PS: 735.49875, W: 1, w: 1 },
      { T: 1, ga: 0.0001 },
      { tsp: 0.00492892159375, tspm: 0.005, tbs: 0.01478676478125, oz: 0.0295735295625, cup: 0.2365882365, pt: 0.473176473, us_pt: 0.473176473, uk_pt: 0.56826125, qt: 0.946352946, uk_qt: 1.1365225, gal: 3.785411784, uk_gal: 4.54609, l: 1, L: 1, lt: 1, ang3: 1e-27, 'ang^3': 1e-27, barrel: 158.987294928, bushel: 35.23907016688, ft3: 28.316846592, 'ft^3': 28.316846592, in3: 0.016387064, 'in^3': 0.016387064, ly3: 8.467866646237152e+50, 'ly^3': 8.467866646237152e+50, m3: 1000, 'm^3': 1000, mi3: 4168181825440.5796, 'mi^3': 4168181825440.5796, yd3: 764.554857984, 'yd^3': 764.554857984, Nmi3: 6352182208000, 'Nmi^3': 6352182208000, Picapt3: 4.39039566186557e-8, 'Picapt^3': 4.39039566186557e-8, Pica3: 4.39039566186557e-8, 'Pica^3': 4.39039566186557e-8, GRT: 2831.6846592, regton: 2831.6846592, MTON: 1132.67386368 },
      { uk_acre: 4046.8564224, us_acre: 4046.8726098742522, ang2: 1e-20, 'ang^2': 1e-20, ar: 100, ft2: 0.09290304, 'ft^2': 0.09290304, ha: 10000, in2: 0.00064516, 'in^2': 0.00064516, ly2: 8.95054210748189e+31, 'ly^2': 8.95054210748189e+31, m2: 1, 'm^2': 1, Morgen: 2500, mi2: 2589988.110336, 'mi^2': 2589988.110336, Nmi2: 3429904, 'Nmi^2': 3429904, Picapt2: 1.2445216049382715e-7, Pica2: 1.2445216049382715e-7, 'Pica^2': 1.2445216049382715e-7, 'Picapt^2': 1.2445216049382715e-7, yd2: 0.83612736, 'yd^2': 0.83612736 },
      { bit: 1, byte: 8 },
      { admkn: 0.5147733333333333, kn: 0.5144444444444445, 'm/h': 0.0002777777777777778, 'm/hr': 0.0002777777777777778, 'm/s': 1, 'm/sec': 1, mph: 0.44704 },
    ];
    // las que admiten prefijo, y la potencia a la que se eleva: km2 son 1000² m²
    const PREFIJABLES = new Map([..."g u m ang ly parsec pc sec s Pa p atm at mmHg N dyn dy pond J e c cal eV ev Wh wh W w T ga uk_pt l L lt ar bit byte m/h m/hr m/s m/sec mph".split(' ').map((u) => [u, 1]), ..."ang2 ang^2 m2 m^2".split(' ').map((u) => [u, 2]), ..."ang3 ang^3 m3 m^3".split(' ').map((u) => [u, 3])]);
    const PREFIJOS_SI = { Y: 1e24, Z: 1e21, E: 1e18, P: 1e15, T: 1e12, G: 1e9, M: 1e6, k: 1e3, h: 100, da: 10, e: 10, d: 0.1, c: 0.01, m: 1e-3, u: 1e-6, n: 1e-9, p: 1e-12, f: 1e-15, a: 1e-18, z: 1e-21, y: 1e-24 };
    const PREFIJOS_BIN = { Yi: 2 ** 80, Zi: 2 ** 70, Ei: 2 ** 60, Pi: 2 ** 50, Ti: 2 ** 40, Gi: 2 ** 30, Mi: 2 ** 20, ki: 2 ** 10 };
    const A_KELVIN = { C: (x) => x + 273.15, cel: (x) => x + 273.15, F: (x) => (x - 32) * 5 / 9 + 273.15, fah: (x) => (x - 32) * 5 / 9 + 273.15, K: (x) => x, kel: (x) => x, Rank: (x) => x * 5 / 9, Reau: (x) => x * 1.25 + 273.15 };
    const DE_KELVIN = { C: (x) => x - 273.15, cel: (x) => x - 273.15, F: (x) => (x - 273.15) * 9 / 5 + 32, fah: (x) => (x - 273.15) * 9 / 5 + 32, K: (x) => x, kel: (x) => x, Rank: (x) => x * 9 / 5, Reau: (x) => (x - 273.15) / 1.25 };
    const propia = (o, k) => Object.prototype.hasOwnProperty.call(o, k);
    // Las lecturas de un nombre de unidad: { grupo, factor }, la exacta primero. La temperatura es el grupo -1: nombre y factor del prefijo (mK).
    // Puede haber más de una: «pc» es un pársec y también una picocaloría (Excel elige la que corresponde a la otra unidad).
    function lecturasConvert(t) {
      const directa = (u, p) => {
        if (propia(A_KELVIN, u)) return p === 1 || u === 'K' || u === 'kel' ? { grupo: -1, nombre: u, factor: p } : null;
        const g = UNIDADES_CONVERT.findIndex((x) => propia(x, u));
        return g < 0 ? null : { grupo: g, factor: UNIDADES_CONVERT[g][u] * p };
      };
      const salida = [directa(t, 1)];
      for (const n of [2, 1]) {
        const p = t.slice(0, n), u = t.slice(n);
        if (!u) continue;
        if (propia(PREFIJOS_BIN, p) && (u === 'bit' || u === 'byte')) salida.push(directa(u, PREFIJOS_BIN[p]));
        if (!propia(PREFIJOS_SI, p)) continue;
        if (u === 'K' || u === 'kel') salida.push(directa(u, PREFIJOS_SI[p]));
        else if (PREFIJABLES.has(u)) salida.push(directa(u, PREFIJOS_SI[p] ** PREFIJABLES.get(u)));
      }
      return salida.filter(Boolean);
    }
    function convertir(a, c, A) {
      const v = A.escalar(A.evaluarNodo(a[0], c), c);
      if (typeof v === 'boolean') return A.E.valor;
      const x = A.aNumero(v), desde = lecturasConvert(A.txt(a[1], c)), hasta = lecturasConvert(A.txt(a[2], c));
      const de = desde.find((d) => hasta.some((h) => h.grupo === d.grupo));
      if (!de) return A.E.na;
      const a2 = hasta.find((h) => h.grupo === de.grupo);
      if (de.grupo === -1) return DE_KELVIN[a2.nombre](A_KELVIN[de.nombre](x * de.factor)) / a2.factor;
      return x * de.factor / a2.factor;
    }

    const formaDe = (M) => [M.length, M[0].length];
    const vacioA0 = (x) => (x === null || x === undefined ? 0 : x);
    const conCalc = (filas, A) => (filas.length && filas[0].length ? mat(filas) : A.error('#CALC!'));
    // comparación de SORT/UNIQUE: números < textos < lógicos; vacías al final; errores como valores
    const ordenSort = (x, y, A) => {
      if (x === null && y === null) return 0;
      if (x === null) return 1;
      if (y === null) return -1;
      if (esError(x) || esError(y)) return esError(x) && esError(y) ? 0 : esError(x) ? 1 : -1;
      return A.comparar(x, y);
    };
    const claveUnica = (fila) => JSON.stringify(fila.map((x) => (typeof x === 'string' ? 's' + x.toUpperCase() : esError(x) ? 'e' + x.codigo : x)));
    return {
      // trigonometría e hiperbólicas
      SINH: (a, c, A) => Math.sinh(A.num(a[0], c)), COSH: (a, c, A) => NUMERO(Math.cosh(A.num(a[0], c)), A), TANH: (a, c, A) => Math.tanh(A.num(a[0], c)),
      ASINH: (a, c, A) => Math.asinh(A.num(a[0], c)),
      ACOSH: (a, c, A) => { const x = A.num(a[0], c); return x < 1 ? A.E.num : Math.acosh(x); },
      ATANH: (a, c, A) => { const x = A.num(a[0], c); return Math.abs(x) >= 1 ? A.E.num : Math.atanh(x); },
      COT: (a, c, A) => { const x = A.num(a[0], c); return x === 0 ? A.E.div : 1 / Math.tan(x); },
      COTH: (a, c, A) => { const x = A.num(a[0], c); return x === 0 ? A.E.div : 1 / Math.tanh(x); },
      CSC: (a, c, A) => { const x = A.num(a[0], c); return x === 0 ? A.E.div : 1 / Math.sin(x); },
      CSCH: (a, c, A) => { const x = A.num(a[0], c); return x === 0 ? A.E.div : 1 / Math.sinh(x); },
      SEC: (a, c, A) => 1 / Math.cos(A.num(a[0], c)),
      SECH: (a, c, A) => 1 / Math.cosh(A.num(a[0], c)),
      ACOT: (a, c, A) => { const x = A.num(a[0], c); return Math.PI / 2 - Math.atan(x); },
      ACOTH: (a, c, A) => { const x = A.num(a[0], c); return Math.abs(x) <= 1 ? A.E.num : 0.5 * Math.log((x + 1) / (x - 1)); },
      SQRTPI: (a, c, A) => { const x = A.num(a[0], c); return x < 0 ? A.E.num : Math.sqrt(x * Math.PI); },

      // combinatoria y redondeos
      FACTDOUBLE: (a, c, A) => { const n = Math.trunc(A.num(a[0], c)); if (n < -1) return A.E.num; let r = 1; for (let k = n; k > 1; k -= 2) r *= k; return NUMERO(r, A); },
      COMBIN: (a, c, A) => { const n = ent(a[0], c, A), k = ent(a[1], c, A); if (n < 0 || k < 0 || k > n) return A.E.num; return NUMERO(combin(n, k), A); },
      COMBINA: (a, c, A) => { const n = ent(a[0], c, A), k = ent(a[1], c, A); if (n < 0 || k < 0 || (n === 0 && k > 0)) return A.E.num; return k === 0 ? 1 : NUMERO(combin(n + k - 1, k), A); },
      PERMUT: (a, c, A) => { const n = ent(a[0], c, A), k = ent(a[1], c, A); if (n < 0 || k < 0 || k > n) return A.E.num; let r = 1; for (let i = n - k + 1; i <= n; i++) r *= i; return NUMERO(r, A); },
      PERMUTATIONA: (a, c, A) => { const n = ent(a[0], c, A), k = ent(a[1], c, A); if (n < 0 || k < 0) return A.E.num; return NUMERO(Math.pow(n, k), A); },
      MULTINOMIAL: (a, c, A) => { const l = numeros(a, c, A).map(Math.trunc); if (l.some((x) => x < 0)) return A.E.num; let s = 0, r = 1; for (const x of l) { for (let i = 1; i <= x; i++) { s++; r = r * s / i; } } return NUMERO(Math.round(r), A); },
      'CEILING.PRECISE': (a, c, A) => { const x = A.num(a[0], c), m = Math.abs(opc(a[1], c, A, 1)); return m === 0 ? 0 : Math.ceil(Number((x / m).toPrecision(15))) * m; },
      'ISO.CEILING': (a, c, A) => { const x = A.num(a[0], c), m = Math.abs(opc(a[1], c, A, 1)); return m === 0 ? 0 : Math.ceil(Number((x / m).toPrecision(15))) * m; },
      'FLOOR.PRECISE': (a, c, A) => { const x = A.num(a[0], c), m = Math.abs(opc(a[1], c, A, 1)); return m === 0 ? 0 : Math.floor(Number((x / m).toPrecision(15))) * m; },
      ROMAN: (a, c, A) => {
        const n = Math.trunc(A.num(a[0], c)), forma = a[1] === undefined || a[1] === null ? 0 : A.escalar(A.evaluarNodo(a[1], c), c);
        if (n < 0 || n > 3999) return A.E.valor;
        if (forma !== 0 && forma !== true) throw new A.NoSoportada('ROMAN con forma abreviada');
        let s = '', x = n; for (const [v, t] of ROMANOS) while (x >= v) { s += t; x -= v; }
        return s;
      },
      ARABIC: (a, c, A) => {
        let t = A.txt(a[0], c).trim().toUpperCase(), signo = 1;
        if (t.length > 255) return A.E.valor;
        if (t[0] === '-') { signo = -1; t = t.slice(1); }
        let n = 0;
        for (let i = 0; i < t.length; i++) { const v = VALOR_ROMANO[t[i]], s = VALOR_ROMANO[t[i + 1]]; if (!v) return A.E.valor; n += s > v ? -v : v; }
        return signo * n;
      },
      BASE: (a, c, A) => {
        const n = A.num(a[0], c), b = ent(a[1], c, A), largo = Math.trunc(opc(a[2], c, A, 0));
        if (n < 0 || n >= 9007199254740992 || b < 2 || b > 36 || largo < 0 || largo > 255) return A.E.num;
        return Math.trunc(n).toString(b).toUpperCase().padStart(largo, '0');
      },
      DECIMAL: (a, c, A) => {
        const t = A.txt(a[0], c).trim().toUpperCase(), b = ent(a[1], c, A);
        if (b < 2 || b > 36 || t.length > 255) return A.E.num;
        let n = 0;
        for (const ch of t) { const d = parseInt(ch, 36); if (Number.isNaN(d) || d >= b) return A.E.num; n = n * b + d; }
        return n;
      },
      DEC2BIN: decA(2), DEC2OCT: decA(8), DEC2HEX: decA(16),
      BIN2DEC: deBase(2), OCT2DEC: deBase(8), HEX2DEC: deBase(16),
      BIN2OCT: entreBases(2, 8), BIN2HEX: entreBases(2, 16), OCT2BIN: entreBases(8, 2), OCT2HEX: entreBases(8, 16), HEX2BIN: entreBases(16, 2), HEX2OCT: entreBases(16, 8),
      BITAND: bitOp((x, y) => (x & y) >>> 0), BITOR: bitOp((x, y) => (x | y) >>> 0), BITXOR: bitOp((x, y) => (x ^ y) >>> 0),
      BITLSHIFT: (a, c, A) => { const n = bit(a[0], c, A), s = ent(a[1], c, A); if (Math.abs(s) > 53) return A.E.num; const r = s >= 0 ? n * Math.pow(2, s) : Math.floor(n / Math.pow(2, -s)); return r >= 281474976710656 ? A.E.num : r; },
      BITRSHIFT: (a, c, A) => { const n = bit(a[0], c, A), s = ent(a[1], c, A); if (Math.abs(s) > 53) return A.E.num; const r = s >= 0 ? Math.floor(n / Math.pow(2, s)) : n * Math.pow(2, -s); return r >= 281474976710656 ? A.E.num : r; },
      DELTA: (a, c, A) => (A.num(a[0], c) === opc(a[1], c, A, 0) ? 1 : 0),
      GESTEP: (a, c, A) => (A.num(a[0], c) >= opc(a[1], c, A, 0) ? 1 : 0),
      SUMX2MY2: sumaPares((y, x) => y * y - x * x), SUMX2PY2: sumaPares((y, x) => y * y + x * x), SUMXMY2: sumaPares((y, x) => (y - x) * (y - x)),
      SERIESSUM: (a, c, A) => { const x = A.num(a[0], c), n = A.num(a[1], c), m = A.num(a[2], c); const k = numeros([a[3]], c, A); let s = 0; k.forEach((ci, i) => { s += ci * Math.pow(x, n + i * m); }); return NUMERO(s, A); },

      // matrices
      MMULT: (a, c, A) => {
        const X = numerica(a[0], c, A), Y = numerica(a[1], c, A);
        if (X[0].length !== Y.length) return A.E.valor;
        return mat(X.map((f) => Y[0].map((_, j) => f.reduce((s, x, k) => s + x * Y[k][j], 0))));
      },
      MDETERM: (a, c, A) => { const M = numerica(a[0], c, A); if (M.length !== M[0].length) return A.E.valor; return determinante(M); },
      MINVERSE: (a, c, A) => { const M = numerica(a[0], c, A); if (M.length !== M[0].length) return A.E.valor; const R = inversaMatriz(M); return R ? mat(R) : A.E.num; },
      MUNIT: (a, c, A) => { const n = ent(a[0], c, A); if (n < 1) return A.E.valor; return mat(Array.from({ length: n }, (_, i) => Array.from({ length: n }, (_, j) => (i === j ? 1 : 0)))); },
      LINEST: (a, c, A) => estimacion(a, c, A, false),
      LOGEST: (a, c, A) => estimacion(a, c, A, true),
      TREND: (a, c, A) => prediccion(a, c, A, false),
      GROWTH: (a, c, A) => prediccion(a, c, A, true),
      CONVERT: convertir,
      TRANSPOSE: (a, c, A) => { let v = A.evaluarNodo(a[0], c); if (!c.matriz && A.esRef(v)) { v = A.escalar(v, c); if (esError(v)) return v; } const M = A.matrizDe(v); return mat(M[0].map((_, j) => M.map((f) => (f[j] === null ? 0 : f[j])))); },

      // estadística
      MAXA: (a, c, A) => { const l = numeros(a, c, A, { conTodo: true }); return l.length ? Math.max(...l) : 0; },
      MINA: (a, c, A) => { const l = numeros(a, c, A, { conTodo: true }); return l.length ? Math.min(...l) : 0; },
      STDEVA: (a, c, A) => desviacion(numeros(a, c, A, { conTodo: true }), 1, A), STDEVPA: (a, c, A) => desviacion(numeros(a, c, A, { conTodo: true }), 0, A),
      VARA: (a, c, A) => varianza(numeros(a, c, A, { conTodo: true }), 1, A), VARPA: (a, c, A) => varianza(numeros(a, c, A, { conTodo: true }), 0, A),
      AVEDEV: (a, c, A) => { const l = numeros(a, c, A); if (!l.length) return A.E.num; const m = sumar(l) / l.length; return sumar(l.map((x) => Math.abs(x - m))) / l.length; },
      DEVSQ: (a, c, A) => { const l = numeros(a, c, A); if (!l.length) return A.E.num; const m = sumar(l) / l.length; return sumar(l.map((x) => (x - m) * (x - m))); },
      GEOMEAN: (a, c, A) => { const l = numeros(a, c, A); if (!l.length || l.some((x) => x <= 0)) return A.E.num; return Math.exp(sumar(l.map(Math.log)) / l.length); },
      HARMEAN: (a, c, A) => { const l = numeros(a, c, A); if (!l.length || l.some((x) => x <= 0)) return A.E.num; return l.length / sumar(l.map((x) => 1 / x)); },
      KURT: (a, c, A) => {
        const l = numeros(a, c, A), n = l.length; if (n < 4) return A.E.div;
        const m = sumar(l) / n, s = desviacion(l, 1, A); if (esError(s) || s === 0) return A.E.div;
        const k = sumar(l.map((x) => Math.pow((x - m) / s, 4)));
        return n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)) * k - 3 * (n - 1) * (n - 1) / ((n - 2) * (n - 3));
      },
      SKEW: (a, c, A) => { const l = numeros(a, c, A), n = l.length; if (n < 3) return A.E.div; const m = sumar(l) / n, s = desviacion(l, 1, A); if (esError(s) || s === 0) return A.E.div; return n / ((n - 1) * (n - 2)) * sumar(l.map((x) => Math.pow((x - m) / s, 3))); },
      'SKEW.P': (a, c, A) => { const l = numeros(a, c, A), n = l.length; if (n < 1) return A.E.div; const m = sumar(l) / n, s = desviacion(l, 0, A); if (esError(s) || s === 0) return A.E.div; return sumar(l.map((x) => Math.pow((x - m) / s, 3))) / n; },
      TRIMMEAN: (a, c, A) => {
        const l = numeros([a[0]], c, A).sort((x, y) => x - y), p = A.num(a[1], c);
        if (!l.length || p < 0 || p >= 1) return A.E.num;
        const k = Math.floor(l.length * p / 2), r = l.slice(k, l.length - k);
        return sumar(r) / r.length;
      },
      CORREL: correl, PEARSON: correl,
      RSQ: (a, c, A) => { const r = correl(a, c, A); return esError(r) ? r : r * r; },
      COVAR: (a, c, A) => { const { xs, ys } = pares(a[0], a[1], c, A); if (!xs.length) return A.E.div; return momentos(xs, ys).sxy / xs.length; },
      'COVARIANCE.P': (a, c, A) => { const { xs, ys } = pares(a[0], a[1], c, A); if (!xs.length) return A.E.div; return momentos(xs, ys).sxy / xs.length; },
      'COVARIANCE.S': (a, c, A) => { const { xs, ys } = pares(a[0], a[1], c, A); if (xs.length < 2) return A.E.div; return momentos(xs, ys).sxy / (xs.length - 1); },
      SLOPE: (a, c, A) => { const m = regresion(a, c, A); return m.sxy / m.sxx; },
      INTERCEPT: (a, c, A) => { const m = regresion(a, c, A); return m.my - m.sxy / m.sxx * m.mx; },
      FORECAST: (a, c, A) => { const x = A.num(a[0], c); const m = regresion(a.slice(1), c, A); return m.my + m.sxy / m.sxx * (x - m.mx); },
      'FORECAST.LINEAR': (a, c, A) => { const x = A.num(a[0], c); const m = regresion(a.slice(1), c, A); return m.my + m.sxy / m.sxx * (x - m.mx); },
      STEYX: (a, c, A) => { const { xs, ys } = pares(a[0], a[1], c, A); if (xs.length < 3) return A.E.div; const m = momentos(xs, ys); if (m.sxx === 0) return A.E.div; return Math.sqrt((m.syy - m.sxy * m.sxy / m.sxx) / (m.n - 2)); },
      STANDARDIZE: (a, c, A) => { const x = A.num(a[0], c), m = A.num(a[1], c), s = A.num(a[2], c); return s <= 0 ? A.E.num : (x - m) / s; },
      'MODE.SNGL': (a, c, A) => { const l = numeros(a, c, A); const k = new Map(); let mejor = null, n = 1; for (const x of l) { const v = (k.get(x) || 0) + 1; k.set(x, v); if (v > n) { n = v; mejor = x; } } return mejor === null ? A.E.na : mejor; },
      'MODE.MULT': (a, c, A) => {
        const l = numeros(a, c, A); const k = new Map(); for (const x of l) k.set(x, (k.get(x) || 0) + 1);
        const max = Math.max(1, ...k.values()); if (max < 2) return A.E.na;
        return mat([...k.keys()].filter((x) => k.get(x) === max).map((x) => [x]));
      },
      'RANK.AVG': (a, c, A) => {
        const x = A.num(a[0], c), l = numeros([a[1]], c, A), asc = opc(a[2], c, A, 0) !== 0;
        let menores = 0, iguales = 0; for (const y of l) { if (y === x) iguales++; else if (asc ? y < x : y > x) menores++; }
        return iguales ? menores + (iguales + 1) / 2 : A.E.na;
      },
      PERCENTRANK: (a, c, A) => rangoPercentil(a, c, A, true), 'PERCENTRANK.INC': (a, c, A) => rangoPercentil(a, c, A, true), 'PERCENTRANK.EXC': (a, c, A) => rangoPercentil(a, c, A, false),
      FREQUENCY: (a, c, A) => {
        const datos = numeros([a[0]], c, A), cortes = numeros([a[1]], c, A);
        const orden = cortes.map((v, i) => [v, i]).sort((p, q) => p[0] - q[0]);
        const cuenta = new Array(cortes.length + 1).fill(0);
        for (const x of datos) { let k = orden.findIndex(([v]) => x <= v); cuenta[k < 0 ? cortes.length : orden[k][1]]++; }
        return mat(cuenta.map((n) => [n]));
      },

      // distribuciones
      'NORM.DIST': (a, c, A) => { const x = A.num(a[0], c), m = A.num(a[1], c), s = A.num(a[2], c), acum = A.log(a[3], c); if (s <= 0) return A.E.num; const z = (x - m) / s; return acum ? normalCdf(z) : normalPdf(z) / s; },
      NORMDIST: (a, c, A) => { const x = A.num(a[0], c), m = A.num(a[1], c), s = A.num(a[2], c), acum = A.log(a[3], c); if (s <= 0) return A.E.num; const z = (x - m) / s; return acum ? normalCdf(z) : normalPdf(z) / s; },
      'NORM.S.DIST': (a, c, A) => { const z = A.num(a[0], c); return A.log(a[1], c) ? normalCdf(z) : normalPdf(z); },
      NORMSDIST: (a, c, A) => normalCdf(A.num(a[0], c)),
      'NORM.INV': (a, c, A) => { const p = prob(a[0], c, A, true), m = A.num(a[1], c), s = A.num(a[2], c); return s <= 0 ? A.E.num : m + s * normalInv(p); },
      NORMINV: (a, c, A) => { const p = prob(a[0], c, A, true), m = A.num(a[1], c), s = A.num(a[2], c); return s <= 0 ? A.E.num : m + s * normalInv(p); },
      'NORM.S.INV': (a, c, A) => normalInv(prob(a[0], c, A, true)), NORMSINV: (a, c, A) => normalInv(prob(a[0], c, A, true)),
      PHI: (a, c, A) => normalPdf(A.num(a[0], c)),
      GAUSS: (a, c, A) => normalCdf(A.num(a[0], c)) - 0.5,
      'LOGNORM.DIST': (a, c, A) => { const x = A.num(a[0], c), m = A.num(a[1], c), s = A.num(a[2], c), acum = A.log(a[3], c); if (x <= 0 || s <= 0) return A.E.num; const z = (Math.log(x) - m) / s; return acum ? normalCdf(z) : normalPdf(z) / (x * s); },
      LOGNORMDIST: (a, c, A) => { const x = A.num(a[0], c), m = A.num(a[1], c), s = A.num(a[2], c); if (x <= 0 || s <= 0) return A.E.num; return normalCdf((Math.log(x) - m) / s); },
      'LOGNORM.INV': (a, c, A) => { const p = prob(a[0], c, A, true), m = A.num(a[1], c), s = A.num(a[2], c); return s <= 0 ? A.E.num : Math.exp(m + s * normalInv(p)); },
      LOGINV: (a, c, A) => { const p = prob(a[0], c, A, true), m = A.num(a[1], c), s = A.num(a[2], c); return s <= 0 ? A.E.num : Math.exp(m + s * normalInv(p)); },
      'EXPON.DIST': (a, c, A) => { const x = A.num(a[0], c), l = A.num(a[1], c), acum = A.log(a[2], c); if (x < 0 || l <= 0) return A.E.num; return acum ? 1 - Math.exp(-l * x) : l * Math.exp(-l * x); },
      EXPONDIST: (a, c, A) => { const x = A.num(a[0], c), l = A.num(a[1], c), acum = A.log(a[2], c); if (x < 0 || l <= 0) return A.E.num; return acum ? 1 - Math.exp(-l * x) : l * Math.exp(-l * x); },
      'POISSON.DIST': (a, c, A) => poisson(a, c, A), POISSON: (a, c, A) => poisson(a, c, A),
      'BINOM.DIST': (a, c, A) => binomial(a, c, A), BINOMDIST: (a, c, A) => binomial(a, c, A),
      'BINOM.INV': (a, c, A) => binomInv(a, c, A), CRITBINOM: (a, c, A) => binomInv(a, c, A),
      'NEGBINOM.DIST': (a, c, A) => {
        const f = ent(a[0], c, A), s = ent(a[1], c, A), p = A.num(a[2], c), acum = A.log(a[3], c);
        if (f < 0 || s < 1 || p < 0 || p > 1) return A.E.num;
        const pmf = (k) => Math.exp(lcombin(k + s - 1, s - 1) + s * Math.log(p) + k * Math.log(1 - p));
        if (!acum) return pmf(f);
        return betaI(p, s, f + 1);
      },
      NEGBINOMDIST: (a, c, A) => { const f = ent(a[0], c, A), s = ent(a[1], c, A), p = A.num(a[2], c); if (f < 0 || s < 1 || p < 0 || p > 1) return A.E.num; return Math.exp(lcombin(f + s - 1, s - 1) + s * Math.log(p) + f * Math.log(1 - p)); },
      'HYPGEOM.DIST': (a, c, A) => hipergeometrica(a, c, A, A.log(a[4], c)), HYPGEOMDIST: (a, c, A) => hipergeometrica(a, c, A, false),
      'WEIBULL.DIST': (a, c, A) => weibull(a, c, A), WEIBULL: (a, c, A) => weibull(a, c, A),
      'GAMMA.DIST': (a, c, A) => gammaDist(a, c, A), GAMMADIST: (a, c, A) => gammaDist(a, c, A),
      'GAMMA.INV': (a, c, A) => gammaInv(a, c, A), GAMMAINV: (a, c, A) => gammaInv(a, c, A),
      GAMMA: (a, c, A) => { const x = A.num(a[0], c); if (x <= 0 && Number.isInteger(x)) return A.E.num; return NUMERO(gamma(x), A); },
      GAMMALN: (a, c, A) => { const x = A.num(a[0], c); return x <= 0 ? A.E.num : lgamma(x); }, 'GAMMALN.PRECISE': (a, c, A) => { const x = A.num(a[0], c); return x <= 0 ? A.E.num : lgamma(x); },
      'BETA.DIST': (a, c, A) => betaDist(a.slice(0, 3), A.log(a[3], c), a[4], a[5], c, A), BETADIST: (a, c, A) => betaDist(a.slice(0, 3), true, a[3], a[4], c, A),
      'BETA.INV': (a, c, A) => betaInv(a, c, A), BETAINV: (a, c, A) => betaInv(a, c, A),
      'T.DIST': (a, c, A) => { const x = A.num(a[0], c), n = gl(a[1], c, A); return A.log(a[2], c) ? tCdf(x, n) : tPdf(x, n); },
      'T.DIST.2T': (a, c, A) => { const x = A.num(a[0], c), n = gl(a[1], c, A); return x < 0 ? A.E.num : betaI(n / (n + x * x), n / 2, 0.5); },
      'T.DIST.RT': (a, c, A) => { const x = A.num(a[0], c), n = gl(a[1], c, A); return 1 - tCdf(x, n); },
      TDIST: (a, c, A) => { const x = A.num(a[0], c), n = gl(a[1], c, A), colas = ent(a[2], c, A); if (x < 0 || (colas !== 1 && colas !== 2)) return A.E.num; const ib = betaI(n / (n + x * x), n / 2, 0.5); return colas === 2 ? ib : ib / 2; },
      'T.INV': (a, c, A) => { const p = prob(a[0], c, A, true), n = gl(a[1], c, A); return tInv(p, n); },
      'T.INV.2T': (a, c, A) => { const p = A.num(a[0], c), n = gl(a[1], c, A); if (p <= 0 || p > 1) return A.E.num; return Math.abs(tInv(p / 2, n)); },
      TINV: (a, c, A) => { const p = A.num(a[0], c), n = gl(a[1], c, A); if (p <= 0 || p > 1) return A.E.num; return Math.abs(tInv(p / 2, n)); },
      'CHISQ.DIST': (a, c, A) => { const x = A.num(a[0], c), k = gl(a[1], c, A, 1e10); if (x < 0) return A.E.num; if (A.log(a[2], c)) return chiCdf(x, k); return x === 0 ? (k === 2 ? 0.5 : k < 2 ? A.E.num : 0) : Math.exp((k / 2 - 1) * Math.log(x) - x / 2 - (k / 2) * Math.LN2 - lgamma(k / 2)); },
      'CHISQ.DIST.RT': (a, c, A) => { const x = A.num(a[0], c), k = gl(a[1], c, A, 1e10); return x < 0 ? A.E.num : gammaQ(k / 2, x / 2); },
      CHIDIST: (a, c, A) => { const x = A.num(a[0], c), k = gl(a[1], c, A, 1e10); return x < 0 ? A.E.num : gammaQ(k / 2, x / 2); },
      'CHISQ.INV': (a, c, A) => { const p = prob(a[0], c, A), k = gl(a[1], c, A, 1e10); if (p === 1) return A.E.num; return p === 0 ? 0 : inversa((x) => chiCdf(x, k), p, 0, Math.max(1, k)); },
      'CHISQ.INV.RT': (a, c, A) => { const p = prob(a[0], c, A), k = gl(a[1], c, A, 1e10); if (p === 0) return A.E.num; return p === 1 ? 0 : inversa((x) => 1 - gammaQ(k / 2, x / 2), 1 - p, 0, Math.max(1, k)); },
      CHIINV: (a, c, A) => { const p = prob(a[0], c, A), k = gl(a[1], c, A, 1e10); if (p === 0) return A.E.num; return p === 1 ? 0 : inversa((x) => 1 - gammaQ(k / 2, x / 2), 1 - p, 0, Math.max(1, k)); },
      'F.DIST': (a, c, A) => {
        const x = A.num(a[0], c), d1 = gl(a[1], c, A, 1e10), d2 = gl(a[2], c, A, 1e10);
        if (x < 0) return A.E.num;
        if (A.log(a[3], c)) return fCdf(x, d1, d2);
        if (x === 0) return d1 === 2 ? 1 : d1 < 2 ? A.E.num : 0;
        return Math.exp(0.5 * (d1 * Math.log(d1 * x) + d2 * Math.log(d2) - (d1 + d2) * Math.log(d1 * x + d2)) - Math.log(x) - (lgamma(d1 / 2) + lgamma(d2 / 2) - lgamma((d1 + d2) / 2)));
      },
      'F.DIST.RT': (a, c, A) => { const x = A.num(a[0], c), d1 = gl(a[1], c, A, 1e10), d2 = gl(a[2], c, A, 1e10); return x < 0 ? A.E.num : 1 - fCdf(x, d1, d2); },
      FDIST: (a, c, A) => { const x = A.num(a[0], c), d1 = gl(a[1], c, A, 1e10), d2 = gl(a[2], c, A, 1e10); return x < 0 ? A.E.num : 1 - fCdf(x, d1, d2); },
      'F.INV': (a, c, A) => { const p = prob(a[0], c, A), d1 = gl(a[1], c, A, 1e10), d2 = gl(a[2], c, A, 1e10); if (p === 1) return A.E.num; return p === 0 ? 0 : inversa((x) => fCdf(x, d1, d2), p, 0, 1); },
      'F.INV.RT': (a, c, A) => { const p = prob(a[0], c, A), d1 = gl(a[1], c, A, 1e10), d2 = gl(a[2], c, A, 1e10); if (p === 0) return A.E.num; return p === 1 ? 0 : inversa((x) => fCdf(x, d1, d2), 1 - p, 0, 1); },
      FINV: (a, c, A) => { const p = prob(a[0], c, A), d1 = gl(a[1], c, A, 1e10), d2 = gl(a[2], c, A, 1e10); if (p === 0) return A.E.num; return p === 1 ? 0 : inversa((x) => fCdf(x, d1, d2), 1 - p, 0, 1); },
      FISHER: (a, c, A) => { const x = A.num(a[0], c); return Math.abs(x) >= 1 ? A.E.num : 0.5 * Math.log((1 + x) / (1 - x)); },
      FISHERINV: (a, c, A) => Math.tanh(A.num(a[0], c)),
      CONFIDENCE: (a, c, A) => confianzaNormal(a, c, A), 'CONFIDENCE.NORM': (a, c, A) => confianzaNormal(a, c, A),
      'CONFIDENCE.T': (a, c, A) => { const al = A.num(a[0], c), s = A.num(a[1], c), n = Math.trunc(A.num(a[2], c)); if (al <= 0 || al >= 1 || s <= 0 || n < 1) return A.E.num; if (n === 1) return A.E.div; return Math.abs(tInv(al / 2, n - 1)) * s / Math.sqrt(n); },
      ERF: (a, c, A) => { const x = A.num(a[0], c); if (a[1] === undefined || a[1] === null) return erf(x); return erf(A.num(a[1], c)) - erf(x); },
      'ERF.PRECISE': (a, c, A) => erf(A.num(a[0], c)),
      ERFC: (a, c, A) => erfc(A.num(a[0], c)), 'ERFC.PRECISE': (a, c, A) => erfc(A.num(a[0], c)),
      'Z.TEST': (a, c, A) => pruebaZ(a, c, A), ZTEST: (a, c, A) => pruebaZ(a, c, A),
      'T.TEST': (a, c, A) => pruebaT(a, c, A), TTEST: (a, c, A) => pruebaT(a, c, A),
      'F.TEST': (a, c, A) => pruebaF(a, c, A), FTEST: (a, c, A) => pruebaF(a, c, A),
      'CHISQ.TEST': (a, c, A) => pruebaChi(a, c, A), CHITEST: (a, c, A) => pruebaChi(a, c, A),

      // finanzas
      NPER: (a, c, A) => {
        const r = A.num(a[0], c), p = A.num(a[1], c), pv = A.num(a[2], c), fv = opc(a[3], c, A, 0), t = opc(a[4], c, A, 0) ? 1 : 0;
        if (r === 0) return p === 0 ? A.E.num : -(pv + fv) / p;
        const q = p * (1 + r * t), x = (q - fv * r) / (q + pv * r);
        return x <= 0 ? A.E.num : Math.log(x) / Math.log(1 + r);
      },
      RATE: (a, c, A) => {
        const n = A.num(a[0], c), p = A.num(a[1], c), pv = A.num(a[2], c), fv = opc(a[3], c, A, 0), t = opc(a[4], c, A, 0) ? 1 : 0, g = opc(a[5], c, A, 0.1);
        if (n <= 0) return A.E.num;
        // (1 + r)^n − 1 con expm1/log1p: sin perder cifras cuando la tasa es casi 0
        const f = (r) => { if (r === 0) return pv + p * n + fv; const e = Math.expm1(n * Math.log1p(r)); return pv * (e + 1) + p * (1 + r * t) * e / r + fv; };
        return newton(f, g, A, Math.abs(pv) + Math.abs(fv) + Math.abs(p * n));
      },
      IRR: (a, c, A) => {
        const l = flujos(a[0], c, A), g = opc(a[1], c, A, 0.1);
        if (!l.some((x) => x > 0) || !l.some((x) => x < 0)) return A.E.num;
        return newton((r) => l.reduce((s, x, i) => s + x / Math.pow(1 + r, i), 0), g, A, l.reduce((s, x) => s + Math.abs(x), 0));
      },
      MIRR: (a, c, A) => {
        const l = flujos(a[0], c, A), fr = A.num(a[1], c), rr = A.num(a[2], c), n = l.length;
        let vpn = 0, vfp = 0;
        l.forEach((x, i) => { if (x < 0) vpn += x / Math.pow(1 + fr, i); else vfp += x * Math.pow(1 + rr, n - 1 - i); });
        if (vpn === 0 || vfp === 0) return A.E.div;
        return Math.pow(-vfp / vpn, 1 / (n - 1)) - 1;
      },
      XNPV: (a, c, A) => {
        const r = A.num(a[0], c), v = flujos(a[1], c, A), d = flujos(a[2], c, A).map(Math.trunc);
        if (v.length !== d.length) return A.E.num;
        if (d.some((x) => x < d[0])) return A.E.num;
        return v.reduce((s, x, i) => s + x / Math.pow(1 + r, (d[i] - d[0]) / 365), 0);
      },
      XIRR: (a, c, A) => {
        const v = flujos(a[0], c, A), d = flujos(a[1], c, A).map(Math.trunc), g = opc(a[2], c, A, 0.1);
        if (v.length !== d.length || !v.some((x) => x > 0) || !v.some((x) => x < 0) || d.some((x) => x < d[0])) return A.E.num;
        return newton((r) => v.reduce((s, x, i) => s + x / Math.pow(1 + r, (d[i] - d[0]) / 365), 0), g, A, v.reduce((s, x) => s + Math.abs(x), 0));
      },
      IPMT: (a, c, A) => { const r = A.num(a[0], c), per = A.num(a[1], c), n = A.num(a[2], c), pv = A.num(a[3], c), fv = opc(a[4], c, A, 0), t = opc(a[5], c, A, 0) ? 1 : 0; if (per < 1 || per > n) return A.E.num; return interes(r, per, n, pv, fv, t); },
      PPMT: (a, c, A) => { const r = A.num(a[0], c), per = A.num(a[1], c), n = A.num(a[2], c), pv = A.num(a[3], c), fv = opc(a[4], c, A, 0), t = opc(a[5], c, A, 0) ? 1 : 0; if (per < 1 || per > n) return A.E.num; return pago(r, n, pv, fv, t) - interes(r, per, n, pv, fv, t); },
      CUMIPMT: (a, c, A) => acumulado(a, c, A, true), CUMPRINC: (a, c, A) => acumulado(a, c, A, false),
      ISPMT: (a, c, A) => { const r = A.num(a[0], c), per = A.num(a[1], c), n = A.num(a[2], c), pv = A.num(a[3], c); if (n === 0) return A.E.div; return pv * r * (per / n - 1); },
      SLN: (a, c, A) => { const v = A.num(a[0], c), s = A.num(a[1], c), n = A.num(a[2], c); return n === 0 ? A.E.div : (v - s) / n; },
      SYD: (a, c, A) => { const v = A.num(a[0], c), s = A.num(a[1], c), n = A.num(a[2], c), p = A.num(a[3], c); if (n <= 0 || p <= 0 || p > n) return A.E.num; return (v - s) * (n - p + 1) * 2 / (n * (n + 1)); },
      DDB: (a, c, A) => {
        const v = A.num(a[0], c), s = A.num(a[1], c), n = A.num(a[2], c), p = A.num(a[3], c), f = opc(a[4], c, A, 2);
        if (v < 0 || s < 0 || n <= 0 || p <= 0 || p > n || f <= 0) return A.E.num;
        let total = 0, actual = 0;
        for (let i = 1; i <= p; i++) { actual = Math.max(0, Math.min((v - total) * (f / n), v - s - total)); total += actual; }
        return actual;
      },
      DB: (a, c, A) => {
        const v = A.num(a[0], c), s = A.num(a[1], c), n = Math.trunc(A.num(a[2], c)), p = Math.trunc(A.num(a[3], c)), m = Math.trunc(opc(a[4], c, A, 12));
        if (v < 0 || s < 0 || n <= 0 || p <= 0 || m < 1 || m > 12 || p > n + (m < 12 ? 1 : 0)) return A.E.num;
        if (v === 0) return 0;
        const tasa = redondear(1 - Math.pow(s / v, 1 / n), 3);
        let total = v * tasa * m / 12, actual = total;
        for (let i = 2; i <= p; i++) { actual = i === n + 1 ? (v - total) * tasa * (12 - m) / 12 : (v - total) * tasa; total += actual; }
        return p === 1 ? v * tasa * m / 12 : actual;
      },
      EFFECT: (a, c, A) => { const r = A.num(a[0], c), n = Math.trunc(A.num(a[1], c)); if (r <= 0 || n < 1) return A.E.num; return Math.pow(1 + r / n, n) - 1; },
      NOMINAL: (a, c, A) => { const r = A.num(a[0], c), n = Math.trunc(A.num(a[1], c)); if (r <= 0 || n < 1) return A.E.num; return n * (Math.pow(1 + r, 1 / n) - 1); },
      PDURATION: (a, c, A) => { const r = A.num(a[0], c), pv = A.num(a[1], c), fv = A.num(a[2], c); if (r <= 0 || pv <= 0 || fv <= 0) return A.E.num; return (Math.log(fv) - Math.log(pv)) / Math.log(1 + r); },
      RRI: (a, c, A) => { const n = A.num(a[0], c), pv = A.num(a[1], c), fv = A.num(a[2], c); if (n <= 0 || pv === 0) return A.E.num; return NUMERO(Math.pow(fv / pv, 1 / n) - 1, A); },
      FVSCHEDULE: (a, c, A) => { let p = A.num(a[0], c); A.recorrerValor(A.evaluarNodo(a[1], c), (v) => { if (esError(v)) A.lanzar(v); p *= 1 + (v === null ? 0 : A.aNumero(v)); }); return p; },
      DOLLARDE: (a, c, A) => { const x = A.num(a[0], c), f = Math.trunc(A.num(a[1], c)); if (f < 0) return A.E.num; if (f === 0) return A.E.div; const e = Math.trunc(x), d = Math.ceil(Math.log10(f)); return e + (x - e) * Math.pow(10, d) / f; },
      DOLLARFR: (a, c, A) => { const x = A.num(a[0], c), f = Math.trunc(A.num(a[1], c)); if (f < 0) return A.E.num; if (f === 0) return A.E.div; const e = Math.trunc(x), d = Math.ceil(Math.log10(f)); return e + (x - e) * f / Math.pow(10, d); },

      // texto
      DOLLAR: (a, c, A) => {
        const x = A.num(a[0], c), d = Math.trunc(opc(a[1], c, A, 2));
        const n = redondear(x, d), dec = d > 0 ? '.' + '0'.repeat(d) : '';
        const en = A.fuente.idioma === 'en';
        // la moneda de la configuración regional (en: dólar con paréntesis; es-CR: colón con signo menos), como Excel
        const codigo = en ? '"$"#,##0' + dec + '_);("$"#,##0' + dec + ')' : '"₡"#,##0' + dec + ';-"₡"#,##0' + dec;
        return textoFormato(n, en ? codigo : codigo.replace(/\./g, A.fuente.separadores ? A.fuente.separadores.decimal : ',').replace(/#,##/g, '#' + (A.fuente.separadores ? A.fuente.separadores.miles : ' ') + '##'), A);
      },
      TIMEVALUE: (a, c, A) => horaDeTexto(A.txt(a[0], c), A),
      // las «B» cuentan bytes solo con un idioma de doble byte (japonés, chino, coreano) como predeterminado; si no, son las mismas
      LENB: (a, c, A) => A.txt(a[0], c).length,
      LEFTB: (a, c, A) => { const t = A.txt(a[0], c), n = Math.trunc(opc(a[1], c, A, 1)); return n < 0 ? A.E.valor : t.slice(0, n); },
      RIGHTB: (a, c, A) => { const t = A.txt(a[0], c), n = Math.trunc(opc(a[1], c, A, 1)); return n < 0 ? A.E.valor : n === 0 ? '' : t.slice(-n); },
      MIDB: (a, c, A) => { const t = A.txt(a[0], c), i = Math.trunc(A.num(a[1], c)), n = Math.trunc(A.num(a[2], c)); return i < 1 || n < 0 ? A.E.valor : t.substr(i - 1, n); },
      TEXTSPLIT: (a, c, A) => {
        const t = A.txt(a[0], c);
        const lista = (nodo) => { if (nodo === undefined || nodo === null) return []; const l = []; A.recorrerValor(A.evaluarNodo(nodo, c), (v) => { if (esError(v)) A.lanzar(v); l.push(A.aTexto(v)); }); return l.filter((x) => x !== ''); };
        const dc = lista(a[1]), df = lista(a[2]), sinVacias = A.opcional(a[3], c, false, A.log), sinMay = opc(a[4], c, A, 0) === 1;
        const relleno = a[5] !== undefined && a[5] !== null ? A.val(a[5], c) : A.E.na;
        if (!dc.length && !df.length) return A.E.valor;
        const cortar = (s, ds) => {
          if (!ds.length) return [s];
          const esc = ds.map((d) => d.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).sort((p, q) => q.length - p.length);
          return s.split(new RegExp(esc.join('|'), sinMay ? 'i' : ''));
        };
        let filas = cortar(t, df).map((f) => cortar(f, dc));
        if (sinVacias) filas = filas.map((f) => f.filter((x) => x !== '')).filter((f) => f.length);
        if (!filas.length) return A.error('#CALC!');
        const ancho = Math.max(...filas.map((f) => f.length));
        return mat(filas.map((f) => f.concat(new Array(ancho - f.length).fill(relleno))));
      },
      // así guarda Excel los textos de más de 255 caracteres: los parte en trozos, _xlfn._LONGTEXT(trozo, trozo, …)
      _LONGTEXT: (a, c, A) => a.map((x) => A.txt(x, c)).join(''),
      VALUETOTEXT: (a, c, A) => { const v = A.escalar(A.evaluarNodo(a[0], c), c); const f = Math.trunc(opc(a[1], c, A, 0)); if (f !== 0 && f !== 1) return A.E.valor; return textoDeValor(v, f === 1, A); },
      FORMULATEXT: (a, c, A) => {
        const r = A.evaluarNodo(a[0], c);
        if (!A.esRef(r)) return A.E.na;
        const f = A.fuente.formulaDe ? A.fuente.formulaDe(r.h, r.f0, r.c0) : null;
        if (!f) return A.E.na;
        const visible = typeof EntradaCelda !== 'undefined' && A.fuente.idioma === 'es' ? EntradaCelda.formulaVisible(f, 'es') : f.replace(/_xlfn\.(?:_xlws\.)?|_xlpm\./gi, '');
        return '=' + visible;
      },
      REGEXTEST: (a, c, A) => expresion(A.txt(a[1], c), opc(a[2], c, A, 0) === 1, A).test(A.txt(a[0], c)),
      REGEXEXTRACT: (a, c, A) => {
        const t = A.txt(a[0], c), modo = Math.trunc(opc(a[2], c, A, 0)), sinMay = opc(a[3], c, A, 0) === 1;
        if (modo < 0 || modo > 2) return A.E.valor;
        if (modo === 1) { const m = t.match(expresion(A.txt(a[1], c), sinMay, A, true)); return m ? mat(m.map((x) => [x])) : A.E.na; }
        const m = expresion(A.txt(a[1], c), sinMay, A).exec(t);
        if (!m) return A.E.na;
        if (modo === 0) return m[0];
        return m.length > 1 ? mat([m.slice(1).map((x) => (x === undefined ? '' : x))]) : A.E.na;
      },
      REGEXREPLACE: (a, c, A) => {
        const t = A.txt(a[0], c), re = expresion(A.txt(a[1], c), opc(a[4], c, A, 0) === 1, A, true), r = A.txt(a[2], c), n = Math.trunc(opc(a[3], c, A, 0));
        if (n === 0) return t.replace(re, r);
        const todos = [...t.matchAll(re)];
        const k = n > 0 ? n - 1 : todos.length + n;
        if (k < 0 || k >= todos.length) return t;
        const m = todos[k];
        const una = new RegExp(re.source, re.flags.replace('g', ''));
        return t.slice(0, m.index) + m[0].replace(una, r) + t.slice(m.index + m[0].length);
      },

      // fechas
      DAYS360: (a, c, A) => {
        const p = A.partes(Math.trunc(A.num(a[0], c))), q = A.partes(Math.trunc(A.num(a[1], c))), europeo = A.opcional(a[2], c, false, A.log);
        let d1 = p.d, d2 = q.d, m2 = q.m, y2 = q.a;
        const ultimo = (x) => new Date(Date.UTC(x.a, x.m, 0)).getUTCDate() === x.d;
        if (europeo) { if (d1 === 31) d1 = 30; if (d2 === 31) d2 = 30; }
        else {
          if (ultimo(p)) d1 = 30;
          if (ultimo(q) && (q.m !== 2 || ultimo(p) && p.m === 2)) { if (d1 < 30) { d2 = 1; m2++; if (m2 > 12) { m2 = 1; y2++; } } else d2 = 30; }
        }
        return (y2 - p.a) * 360 + (m2 - p.m) * 30 + (d2 - d1);
      },
      ISOWEEKNUM: (a, c, A) => {
        const x = Math.trunc(A.num(a[0], c)); if (x < 0) return A.E.num;
        const p = A.partes(x);
        const d = new Date(Date.UTC(p.a, p.m - 1, p.d));
        const ds = (d.getUTCDay() + 6) % 7;
        d.setUTCDate(d.getUTCDate() - ds + 3);
        const primero = new Date(Date.UTC(d.getUTCFullYear(), 0, 4));
        return 1 + Math.round(((d - primero) / 86400000 - 3 + ((primero.getUTCDay() + 6) % 7)) / 7);
      },

      // referencias
      INDIRECT: (a, c, A) => {
        let t = A.txt(a[0], c).trim();
        const a1 = A.opcional(a[1], c, true, A.log);
        if (!a1) { t = r1c1aA1(t, c, A.fuente.idioma === 'es' ? 'F' : 'R'); if (!t) return A.E.ref; }
        if (!t) return A.E.ref;
        const ast = A.analizar(t, c.h, c.f, c.c);
        if (!ast || !(ast.t === 'ref' || ast.t === 'nombre' || (ast.t === 'bin' && ast.op === ':'))) return A.E.ref;
        if (ast.t === 'ref' && ast.hoja && /^\[\d+\]/.test(ast.hoja.a)) return A.E.ref;
        const v = A.evaluarNodo(ast, c);
        return A.esRef(v) || A.esUnion(v) ? v : A.E.ref;
      },
      ADDRESS: (a, c, A) => {
        const f = ent(a[0], c, A), col = ent(a[1], c, A), abs = Math.trunc(opc(a[2], c, A, 1)), a1 = A.opcional(a[3], c, true, A.log);
        if (f < 1 || f > 1048576 || col < 1 || col > 16384 || abs < 1 || abs > 4) return A.E.valor;
        let s;
        if (a1) s = (abs === 1 || abs === 3 ? '$' : '') + letras(col - 1) + (abs === 1 || abs === 2 ? '$' : '') + f;
        else { const R = A.fuente.idioma === 'es' ? 'F' : 'R'; s = R + (abs === 1 || abs === 2 ? f : '[' + f + ']') + 'C' + (abs === 1 || abs === 3 ? col : '[' + col + ']'); }
        if (a[4] !== undefined && a[4] !== null) { const h = A.txt(a[4], c); if (h !== '') s = citarHoja(h) + '!' + s; }
        return s;
      },
      AREAS: (a, c, A) => { const v = A.evaluarNodo(a[0], c); if (A.esUnion(v)) return v.refs.length; if (A.esRef(v)) return 1; return A.E.valor; },
      SHEET: (a, c, A) => {
        if (a[0] === undefined || a[0] === null) return c.h + 1;
        const v = A.evaluarNodo(a[0], c);
        if (A.esRef(v)) return v.h + 1;
        if (A.esUnion(v)) return v.refs[0].h + 1;
        const t = A.escalar(v, c);
        if (typeof t !== 'string') return A.E.na;
        const i = A.indiceHoja(t);
        return i < 0 ? A.E.na : i + 1;
      },
      SHEETS: (a, c, A) => {
        if (a[0] === undefined || a[0] === null) return A.fuente.hojas().length;
        const v = A.evaluarNodo(a[0], c);
        if (A.esUnion(v)) return new Set(v.refs.map((r) => r.h)).size;
        return A.esRef(v) ? 1 : A.E.na;
      },

      // matrices dinámicas (fase 3.1)
      FILTER: (a, c, A) => {
        const M = matrizArg(a[0], c, A), K = matrizArg(a[1], c, A), [nf, nc] = formaDe(M);
        let filas;
        if (K.length === nf && K[0].length === 1) filas = M.filter((_, i) => { const x = K[i][0]; if (esError(x)) A.lanzar(x); return A.aLogico(x === null ? false : x); });
        else if (K.length === 1 && K[0].length === nc) { const cols = K[0].map((x) => { if (esError(x)) A.lanzar(x); return A.aLogico(x === null ? false : x); }); filas = M.map((f) => f.filter((_, j) => cols[j])); if (!filas[0].length) filas = []; }
        else return A.E.valor;
        if (!filas.length) return a[2] !== undefined && a[2] !== null ? A.evaluarNodo(a[2], c) : A.error('#CALC!');
        return mat(filas.map((f) => f.map(vacioA0)));
      },
      SORT: (a, c, A) => {
        const M = matrizArg(a[0], c, A), idx = Math.trunc(opc(a[1], c, A, 1)), orden = Math.trunc(opc(a[2], c, A, 1)), porCol = A.opcional(a[3], c, false, A.log);
        if (orden !== 1 && orden !== -1) return A.E.valor;
        const T = porCol ? M[0].map((_, j) => M.map((f) => f[j])) : M;
        if (idx < 1 || idx > T[0].length) return A.E.valor;
        const o = T.map((f, i) => [f, i]).sort((p, q) => orden * ordenSort(p[0][idx - 1], q[0][idx - 1], A) || p[1] - q[1]).map((x) => x[0]);
        const R = porCol ? o[0].map((_, j) => o.map((f) => f[j])) : o;
        return mat(R.map((f) => f.map(vacioA0)));
      },
      SORTBY: (a, c, A) => {
        const M = matrizArg(a[0], c, A), claves = [];
        for (let k = 1; k < a.length; k += 2) {
          const K = matrizArg(a[k], c, A), o = a[k + 1] === undefined || a[k + 1] === null ? 1 : Math.trunc(A.num(a[k + 1], c));
          if (o !== 1 && o !== -1) return A.E.valor;
          if (K[0].length === 1 && K.length === M.length) claves.push({ v: K.map((f) => f[0]), o, filas: true });
          else if (K.length === 1 && K[0].length === M[0].length) claves.push({ v: K[0], o, filas: false });
          else return A.E.valor;
        }
        const porFilas = !claves.length || claves[0].filas;
        if (claves.some((x) => x.filas !== porFilas)) return A.E.valor;
        const n = porFilas ? M.length : M[0].length;
        const ix = [...Array(n).keys()].sort((p, q) => { for (const k of claves) { const d = k.o * ordenSort(k.v[p], k.v[q], A); if (d) return d; } return p - q; });
        return mat((porFilas ? ix.map((i) => M[i]) : M.map((f) => ix.map((j) => f[j]))).map((f) => f.map(vacioA0)));
      },
      UNIQUE: (a, c, A) => {
        const M = matrizArg(a[0], c, A), porCol = A.opcional(a[1], c, false, A.log), solo = A.opcional(a[2], c, false, A.log);
        const T = porCol ? M[0].map((_, j) => M.map((f) => f[j])) : M;
        const cuenta = new Map();
        for (const f of T) { const k = claveUnica(f); cuenta.set(k, (cuenta.get(k) || 0) + 1); }
        const vistos = new Set(), R = [];
        for (const f of T) { const k = claveUnica(f); if (vistos.has(k)) continue; vistos.add(k); if (!solo || cuenta.get(k) === 1) R.push(f); }
        if (!R.length) return A.error('#CALC!');
        const S = porCol ? R[0].map((_, i) => R.map((f) => f[i])) : R;
        return mat(S.map((f) => f.map(vacioA0)));
      },
      SEQUENCE: (a, c, A) => {
        const nf = Math.trunc(A.num(a[0], c)), nc = Math.trunc(opc(a[1], c, A, 1)), ini = opc(a[2], c, A, 1), paso = opc(a[3], c, A, 1);
        if (nf < 1 || nc < 1) return nf === 0 || nc === 0 ? A.error('#CALC!') : A.E.valor;
        if (nf * nc > 1048576 * 16) return A.E.num;
        return mat(Array.from({ length: nf }, (_, i) => Array.from({ length: nc }, (_, j) => ini + (i * nc + j) * paso)));
      },
      RANDARRAY: (a, c, A) => {
        const nf = Math.trunc(opc(a[0], c, A, 1)), nc = Math.trunc(opc(a[1], c, A, 1)), lo = opc(a[2], c, A, 0), hi = opc(a[3], c, A, 1), entero = A.opcional(a[4], c, false, A.log);
        if (nf < 1 || nc < 1 || lo > hi) return A.E.valor;
        return mat(Array.from({ length: nf }, () => Array.from({ length: nc }, () => (entero ? Math.ceil(lo) + Math.floor(Math.random() * (Math.floor(hi) - Math.ceil(lo) + 1)) : lo + Math.random() * (hi - lo)))));
      },
      VSTACK: (a, c, A) => {
        const Ms = a.filter((x) => x !== null).map((x) => matrizArg(x, c, A)), ancho = Math.max(...Ms.map((M) => M[0].length));
        return mat([].concat(...Ms.map((M) => M.map((f) => f.map(vacioA0).concat(new Array(ancho - f.length).fill(A.E.na))))));
      },
      HSTACK: (a, c, A) => {
        const Ms = a.filter((x) => x !== null).map((x) => matrizArg(x, c, A)), alto = Math.max(...Ms.map((M) => M.length));
        return mat(Array.from({ length: alto }, (_, i) => [].concat(...Ms.map((M) => (M[i] ? M[i].map(vacioA0) : new Array(M[0].length).fill(A.E.na))))));
      },
      TAKE: (a, c, A) => {
        const M = matrizArg(a[0], c, A);
        const nf = a[1] === undefined || a[1] === null ? M.length : Math.trunc(A.num(a[1], c)), nc = a[2] === undefined || a[2] === null ? M[0].length : Math.trunc(A.num(a[2], c));
        if (nf === 0 || nc === 0) return A.error('#CALC!');
        const corta = (l, n) => (n >= 0 ? l.slice(0, n) : l.slice(Math.max(0, l.length + n)));
        return conCalc(corta(M, nf).map((f) => corta(f, nc).map(vacioA0)), A);
      },
      DROP: (a, c, A) => {
        const M = matrizArg(a[0], c, A);
        const nf = a[1] === undefined || a[1] === null ? 0 : Math.trunc(A.num(a[1], c)), nc = a[2] === undefined || a[2] === null ? 0 : Math.trunc(A.num(a[2], c));
        const quita = (l, n) => (n >= 0 ? l.slice(n) : l.slice(0, Math.max(0, l.length + n)));
        return conCalc(quita(M, nf).map((f) => quita(f, nc).map(vacioA0)), A);
      },
      CHOOSECOLS: (a, c, A) => {
        const M = matrizArg(a[0], c, A), n = M[0].length, cols = [];
        for (const x of a.slice(1)) A.recorrerValor(A.evaluarNodo(x, c), (v) => { const k = Math.trunc(A.aNumero(v)); if (k === 0 || Math.abs(k) > n) A.lanzar(A.E.valor); cols.push(k > 0 ? k - 1 : n + k); });
        return mat(M.map((f) => cols.map((j) => vacioA0(f[j]))));
      },
      CHOOSEROWS: (a, c, A) => {
        const M = matrizArg(a[0], c, A), n = M.length, filas = [];
        for (const x of a.slice(1)) A.recorrerValor(A.evaluarNodo(x, c), (v) => { const k = Math.trunc(A.aNumero(v)); if (k === 0 || Math.abs(k) > n) A.lanzar(A.E.valor); filas.push(k > 0 ? k - 1 : n + k); });
        return mat(filas.map((i) => M[i].map(vacioA0)));
      },
      TOCOL: (a, c, A) => aplanar(a, c, A, true), TOROW: (a, c, A) => aplanar(a, c, A, false),
      WRAPROWS: (a, c, A) => envolver(a, c, A, true), WRAPCOLS: (a, c, A) => envolver(a, c, A, false),
      EXPAND: (a, c, A) => {
        const M = matrizArg(a[0], c, A);
        const nf = a[1] === undefined || a[1] === null ? M.length : Math.trunc(A.num(a[1], c)), nc = a[2] === undefined || a[2] === null ? M[0].length : Math.trunc(A.num(a[2], c));
        if (nf < M.length || nc < M[0].length) return A.E.valor;
        const rel = a[3] !== undefined && a[3] !== null ? A.val(a[3], c) : A.E.na;
        return mat(Array.from({ length: nf }, (_, i) => Array.from({ length: nc }, (_, j) => (M[i] && j < M[i].length ? vacioA0(M[i][j]) : rel))));
      },

      // bases de datos
      DSUM: bd((l) => sumar(l)),
      DCOUNT: (a, c, A) => { const { valores, col } = registros(a, c, A); return col === null ? valores.length : valores.filter(esNum).length; },
      DCOUNTA: (a, c, A) => { const { valores, col } = registros(a, c, A); return col === null ? valores.length : valores.filter((v) => v !== null && v !== '').length; },
      DAVERAGE: bd((l, _, A) => (l.length ? sumar(l) / l.length : A.E.div)),
      DMAX: bd((l) => (l.length ? Math.max(...l) : 0)),
      DMIN: bd((l) => (l.length ? Math.min(...l) : 0)),
      DPRODUCT: bd((l) => l.reduce((p, x) => p * x, 1)),
      DSTDEV: bd((l, _, A) => desviacion(l, 1, A)), DSTDEVP: bd((l, _, A) => desviacion(l, 0, A)),
      DVAR: bd((l, _, A) => varianza(l, 1, A)), DVARP: bd((l, _, A) => varianza(l, 0, A)),
      DGET: (a, c, A) => { const { valores } = registros(a, c, A); if (!valores.length) return A.E.valor; if (valores.length > 1) return A.E.num; return valores[0] === null ? 0 : valores[0]; },
    };

    // TOCOL/TOROW(matriz; ignorar: 0 nada, 1 vacías, 2 errores, 3 ambas; por columna)
    function aplanar(a, c, A, columna) {
      const M = matrizArg(a[0], c, A), ign = Math.trunc(opc(a[1], c, A, 0)), porCol = A.opcional(a[2], c, false, A.log);
      if (ign < 0 || ign > 3) return A.E.valor;
      const l = [];
      const tomar = (x) => { if ((ign & 1) && x === null) return; if ((ign & 2) && esError(x)) return; l.push(vacioA0(x)); };
      if (porCol) { for (let j = 0; j < M[0].length; j++) for (const f of M) tomar(f[j]); } else for (const f of M) for (const x of f) tomar(x);
      if (!l.length) return A.error('#CALC!');
      return mat(columna ? l.map((x) => [x]) : [l]);
    }
    function envolver(a, c, A, porFilas) {
      const M = matrizArg(a[0], c, A), n = Math.trunc(A.num(a[1], c));
      if (M.length > 1 && M[0].length > 1) return A.E.valor;
      if (n < 1) return A.E.num;
      const l = M.length > 1 ? M.map((f) => f[0]) : M[0];
      const rel = a[2] !== undefined && a[2] !== null ? A.val(a[2], c) : A.E.na;
      const grupos = [];
      for (let i = 0; i < l.length; i += n) { const g = l.slice(i, i + n).map(vacioA0); while (g.length < n) g.push(rel); grupos.push(g); }
      return mat(porFilas ? grupos : grupos[0].map((_, j) => grupos.map((g) => g[j])));
    }
    function rangoPercentil(a, c, A, inclusivo) {
      const l = numeros([a[0]], c, A).sort((p, q) => p - q), x = A.num(a[1], c), sig = Math.trunc(opc(a[2], c, A, 3));
      const n = l.length;
      if (!n || sig < 1) return A.E.num;
      if (x < l[0] || x > l[n - 1]) return A.E.na;
      let i = l.findIndex((v) => v >= x), pos;
      if (l[i] === x) pos = i;
      else pos = i - 1 + (x - l[i - 1]) / (l[i] - l[i - 1]);
      const r = inclusivo ? (n === 1 ? 1 : pos / (n - 1)) : (pos + 1) / (n + 1);
      return truncar(r, sig);
    }
    function poisson(a, c, A) {
      const x = ent(a[0], c, A), m = A.num(a[1], c), acum = A.log(a[2], c);
      if (x < 0 || m < 0) return A.E.num;
      if (acum) return gammaQ(x + 1, m);
      return Math.exp(x * Math.log(m) - m - lgamma(x + 1));
    }
    function binomial(a, c, A) {
      const k = ent(a[0], c, A), n = ent(a[1], c, A), p = A.num(a[2], c), acum = A.log(a[3], c);
      if (k < 0 || k > n || p < 0 || p > 1) return A.E.num;
      if (!acum) return binomPmf(k, n, p);
      if (k === n) return 1;
      return 1 - betaI(p, k + 1, n - k);
    }
    function binomInv(a, c, A) {
      const n = ent(a[0], c, A), p = A.num(a[1], c), al = A.num(a[2], c);
      if (n < 0 || p < 0 || p > 1 || al < 0 || al > 1) return A.E.num;
      let s = 0;
      for (let k = 0; k <= n; k++) { s += binomPmf(k, n, p); if (s >= al * (1 - 1e-15)) return k; }
      return n;
    }
    function hipergeometrica(a, c, A, acum) {
      const k = ent(a[0], c, A), n = ent(a[1], c, A), K = ent(a[2], c, A), N = ent(a[3], c, A);
      if (k < 0 || n <= 0 || K <= 0 || N <= 0 || n > N || K > N || k > n || k > K || n - k > N - K) return A.E.num;
      const pmf = (x) => Math.exp(lcombin(K, x) + lcombin(N - K, n - x) - lcombin(N, n));
      if (!acum) return pmf(k);
      let s = 0; for (let x = Math.max(0, n - (N - K)); x <= k; x++) s += pmf(x);
      return Math.min(1, s);
    }
    function weibull(a, c, A) {
      const x = A.num(a[0], c), al = A.num(a[1], c), be = A.num(a[2], c), acum = A.log(a[3], c);
      if (x < 0 || al <= 0 || be <= 0) return A.E.num;
      const z = Math.pow(x / be, al);
      return acum ? 1 - Math.exp(-z) : al / be * Math.pow(x / be, al - 1) * Math.exp(-z);
    }
    function gammaDist(a, c, A) {
      const x = A.num(a[0], c), al = A.num(a[1], c), be = A.num(a[2], c), acum = A.log(a[3], c);
      if (x < 0 || al <= 0 || be <= 0) return A.E.num;
      if (acum) return gammaP(al, x / be);
      if (x === 0) return al < 1 ? A.E.num : al === 1 ? 1 / be : 0;
      return Math.exp((al - 1) * Math.log(x) - x / be - al * Math.log(be) - lgamma(al));
    }
    function gammaInv(a, c, A) {
      const p = prob(a[0], c, A), al = A.num(a[1], c), be = A.num(a[2], c);
      if (al <= 0 || be <= 0 || p === 1) return A.E.num;
      if (p === 0) return 0;
      return be * inversa((x) => gammaP(al, x), p, 0, Math.max(1, al));
    }
    function betaDist(a, acum, nodoA, nodoB, c, A) {
      const x = A.num(a[0], c), al = A.num(a[1], c), be = A.num(a[2], c);
      const lo = nodoA === undefined || nodoA === null ? 0 : A.num(nodoA, c), hi = nodoB === undefined || nodoB === null ? 1 : A.num(nodoB, c);
      if (al <= 0 || be <= 0 || x < lo || x > hi || lo === hi) return A.E.num;
      const z = (x - lo) / (hi - lo);
      if (acum) return betaI(z, al, be);
      return Math.exp((al - 1) * Math.log(z) + (be - 1) * Math.log(1 - z) - (lgamma(al) + lgamma(be) - lgamma(al + be))) / (hi - lo);
    }
    function betaInv(a, c, A) {
      const p = prob(a[0], c, A), al = A.num(a[1], c), be = A.num(a[2], c);
      const lo = opc(a[3], c, A, 0), hi = opc(a[4], c, A, 1);
      if (al <= 0 || be <= 0 || p === 0 || p === 1 || lo >= hi) return A.E.num;
      let x0 = 0, x1 = 1;
      for (let i = 0; i < 200; i++) { const m = (x0 + x1) / 2; if (betaI(m, al, be) < p) x0 = m; else x1 = m; }
      return lo + (hi - lo) * (x0 + x1) / 2;
    }
    function confianzaNormal(a, c, A) {
      const al = A.num(a[0], c), s = A.num(a[1], c), n = Math.trunc(A.num(a[2], c));
      if (al <= 0 || al >= 1 || s <= 0 || n < 1) return A.E.num;
      return normalInv(1 - al / 2) * s / Math.sqrt(n);
    }
    function acumulado(a, c, A, deInteres) {
      const r = A.num(a[0], c), n = Math.trunc(A.num(a[1], c)), pv = A.num(a[2], c), i0 = Math.trunc(A.num(a[3], c)), i1 = Math.trunc(A.num(a[4], c)), t = A.num(a[5], c);
      if (r <= 0 || n <= 0 || pv <= 0 || i0 < 1 || i1 < i0 || i1 > n || (t !== 0 && t !== 1)) return A.E.num;
      let s = 0;
      for (let k = i0; k <= i1; k++) { const i = interes(r, k, n, pv, 0, t); s += deInteres ? i : pago(r, n, pv, 0, t) - i; }
      return s;
    }
    function pruebaZ(a, c, A) {
      const l = numeros([a[0]], c, A), x = A.num(a[1], c);
      if (!l.length) return A.E.na;
      const s = a[2] === undefined || a[2] === null ? desviacion(l, 1, A) : A.num(a[2], c);
      if (esError(s)) return s;
      if (s === 0) return A.E.div;
      return 1 - normalCdf((sumar(l) / l.length - x) / (s / Math.sqrt(l.length)));
    }
    function pruebaT(a, c, A) {
      const colas = ent(a[2], c, A), tipo = ent(a[3], c, A);
      if ((colas !== 1 && colas !== 2) || tipo < 1 || tipo > 3) return A.E.num;
      let t, df;
      if (tipo === 1) {
        const { xs, ys } = pares(a[0], a[1], c, A);
        const d = ys.map((y, i) => y - xs[i]); const n = d.length;
        if (n < 2) return A.E.div;
        const m = sumar(d) / n, s = desviacion(d, 1, A); if (s === 0) return A.E.div;
        t = m / (s / Math.sqrt(n)); df = n - 1;
      } else {
        const x = numeros([a[0]], c, A), y = numeros([a[1]], c, A), n1 = x.length, n2 = y.length;
        if (n1 < 2 || n2 < 2) return A.E.div;
        const m1 = sumar(x) / n1, m2 = sumar(y) / n2, v1 = varianza(x, 1, A), v2 = varianza(y, 1, A);
        if (tipo === 2) { df = n1 + n2 - 2; const sp = ((n1 - 1) * v1 + (n2 - 1) * v2) / df; t = (m1 - m2) / Math.sqrt(sp * (1 / n1 + 1 / n2)); }
        else { const q1 = v1 / n1, q2 = v2 / n2; t = (m1 - m2) / Math.sqrt(q1 + q2); df = (q1 + q2) * (q1 + q2) / (q1 * q1 / (n1 - 1) + q2 * q2 / (n2 - 1)); }
        if (!Number.isFinite(t)) return A.E.div;
      }
      const p = betaI(df / (df + t * t), df / 2, 0.5);
      return colas === 2 ? p : p / 2;
    }
    function pruebaF(a, c, A) {
      const x = numeros([a[0]], c, A), y = numeros([a[1]], c, A);
      if (x.length < 2 || y.length < 2) return A.E.div;
      const v1 = varianza(x, 1, A), v2 = varianza(y, 1, A);
      if (v1 === 0 || v2 === 0) return A.E.div;
      const p = fCdf(v1 / v2, x.length - 1, y.length - 1);
      return 2 * Math.min(p, 1 - p);
    }
    function pruebaChi(a, c, A) {
      const R = A.matrizDe(A.evaluarNodo(a[0], c)), E = A.matrizDe(A.evaluarNodo(a[1], c));
      if (R.length !== E.length || R[0].length !== E[0].length) return A.E.na;
      let chi = 0;
      for (let i = 0; i < R.length; i++) for (let j = 0; j < R[0].length; j++) {
        const r = R[i][j], e = E[i][j];
        if (!esNum(r) || !esNum(e)) continue;
        if (e === 0) return A.E.div;
        chi += (r - e) * (r - e) / e;
      }
      const df = R.length > 1 && R[0].length > 1 ? (R.length - 1) * (R[0].length - 1) : R.length * R[0].length - 1;
      if (df < 1) return A.E.na;
      return gammaQ(df / 2, chi / 2);
    }
  });
})();

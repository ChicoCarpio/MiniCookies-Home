/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Evaluador de expresiones de Excel del motor propio, el mínimo que necesita el formato condicional (etapa 6): las reglas de expresión,
// las de texto, de vacíos y de fechas se guardan como fórmula y se evalúan celda por celda con referencias relativas. La etapa 8
// (fórmulas) lo amplía; por eso vive en su propio archivo. Sin DOM.
//
// Valores: número, texto, booleano, null (vacía) o un ErrorExcel ('#DIV/0!', '#N/A', '#VALUE!', '#REF!', '#NAME?', '#NUM!').
// Uso:
//   const e = XlsxExpresion.compilar('AND(A1>2,MOD(ROW(),2)=0)');        // null si no se entiende
//   XlsxExpresion.evaluar(e, { fila, col, origen: { f, c }, valor: (f, c) => …, rango: (f0, c0, f1, c1) => [valores], hoy })
//   (origen = celda a la que se refieren las direcciones relativas; fila/col = la celda que se evalúa)

(function () {
  const MAX_LARGO = 8192, MAX_PROFUNDIDAD = 64, MAX_ARBOL = 1500, MAX_CELDAS_RANGO = 2000000; // anidamiento al compilar / profundidad del árbol al evaluar (una suma de 1.000 términos es un árbol de 1.000 niveles)

  class ErrorExcel { constructor(codigo) { this.codigo = codigo; } toString() { return this.codigo; } }
  const ERR = { div: new ErrorExcel('#DIV/0!'), na: new ErrorExcel('#N/A'), valor: new ErrorExcel('#VALUE!'), ref: new ErrorExcel('#REF!'), nombre: new ErrorExcel('#NAME?'), num: new ErrorExcel('#NUM!') };
  const esError = (v) => v instanceof ErrorExcel;

  // ---- Análisis ---------------------------------------------------------------------------------------
  const colANum = (s) => { let n = 0; for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) & 31); return n - 1; };

  function tokenizar(t) {
    const out = [];
    let i = 0;
    while (i < t.length) {
      const ch = t[i];
      if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') { i++; continue; }
      let m;
      if (ch === '"') { // texto con "" como comilla
        let j = i + 1, s = '';
        for (;;) { if (j >= t.length) return null; if (t[j] === '"') { if (t[j + 1] === '"') { s += '"'; j += 2; continue; } break; } s += t[j++]; }
        out.push({ t: 'texto', v: s }); i = j + 1; continue;
      }
      if ((m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(t.slice(i, i + 40)))) { out.push({ t: 'num', v: parseFloat(m[0]) }); i += m[0].length; continue; }
      // dirección con o sin hoja: se ignora el prefijo si es esta hoja (el llamador decide); $A$1, A1:B2, $A:$A, 1:3
      if ((m = /^(?:(?:'(?:[^']|'')*'|[A-Za-z_][\w.]*)!)?(\$?[A-Za-z]{1,3})(\$?)(\d+)(?![\w(])(?::(\$?[A-Za-z]{1,3})(\$?)(\d+))?/.exec(t.slice(i, i + 80)))) {
        const pre = m[0].includes('!') ? m[0].slice(0, m[0].lastIndexOf('!')) : null;
        const a = (c, af, f) => ({ c: colANum(c.replace('$', '')), ac: c[0] === '$', f: parseInt(f, 10) - 1, af: af === '$' });
        const r1 = a(m[1], m[2], m[3]);
        out.push({ t: 'ref', hoja: pre, a: r1, b: m[4] ? a(m[4], m[5], m[6]) : null }); i += m[0].length; continue;
      }
      if ((m = /^(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})(?![\w(])/.exec(t.slice(i, i + 20)))) { // columnas completas
        out.push({ t: 'ref', hoja: null, a: { c: colANum(m[1].replace('$', '')), ac: m[1][0] === '$', f: 0, af: true }, b: { c: colANum(m[2].replace('$', '')), ac: m[2][0] === '$', f: 1048575, af: true } }); i += m[0].length; continue;
      }
      if ((m = /^[A-Za-zÁÉÍÓÚÑáéíóúñ_][\w.]*/.exec(t.slice(i, i + 80)))) { out.push({ t: 'id', v: m[0] }); i += m[0].length; continue; }
      if ((m = /^(<=|>=|<>|[-+*/^&=<>%(),;:])/.exec(t.slice(i, i + 2)))) { out.push({ t: 'op', v: m[0] }); i += m[0].length; continue; }
      if ((m = /^#(DIV\/0!|N\/A|VALUE!|REF!|NAME\?|NUM!)/.exec(t.slice(i, i + 10)))) { out.push({ t: 'error', v: m[0] }); i += m[0].length; continue; }
      return null;
    }
    return out;
  }

  // Precedencia de Excel: comparación < & < + - < * / < ^ < % < unario (el menos unario liga más que ^ en Excel: -2^2 = 4).
  function compilar(texto) {
    let t = String(texto == null ? '' : texto).trim();
    if (t[0] === '=') t = t.slice(1);
    if (!t || t.length > MAX_LARGO) return null;
    const tk = tokenizar(t);
    if (!tk) return null;
    let p = 0, prof = 0;
    const sig = () => tk[p], es = (v) => tk[p] && tk[p].t === 'op' && tk[p].v === v;
    function comparar() {
      let a = concatenar();
      while (a && tk[p] && tk[p].t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(tk[p].v)) { const o = tk[p++].v; const b = concatenar(); if (!b) return null; a = { k: 'bin', o, a, b }; }
      return a;
    }
    function concatenar() {
      let a = suma();
      while (a && es('&')) { p++; const b = suma(); if (!b) return null; a = { k: 'bin', o: '&', a, b }; }
      return a;
    }
    function suma() {
      let a = producto();
      while (a && (es('+') || es('-'))) { const o = tk[p++].v; const b = producto(); if (!b) return null; a = { k: 'bin', o, a, b }; }
      return a;
    }
    function producto() {
      let a = potencia();
      while (a && (es('*') || es('/'))) { const o = tk[p++].v; const b = potencia(); if (!b) return null; a = { k: 'bin', o, a, b }; }
      return a;
    }
    function potencia() {
      let a = unario();
      while (a && es('^')) { p++; const b = unario(); if (!b) return null; a = { k: 'bin', o: '^', a, b }; }
      return a;
    }
    function unario() {
      if (es('-') || es('+')) { const o = tk[p++].v; const a = unario(); return a && { k: 'un', o, a }; }
      let a = primario();
      while (a && es('%')) { p++; a = { k: 'un', o: '%', a }; }
      return a;
    }
    function primario() {
      if (++prof > MAX_PROFUNDIDAD) return null;
      try {
        const x = tk[p];
        if (!x) return null;
        if (x.t === 'num') { p++; return { k: 'lit', v: x.v }; }
        if (x.t === 'texto') { p++; return { k: 'lit', v: x.v }; }
        if (x.t === 'error') { p++; return { k: 'lit', v: new ErrorExcel(x.v) }; }
        if (x.t === 'ref') { p++; return { k: 'ref', a: x.a, b: x.b, hoja: x.hoja }; }
        if (x.t === 'op' && x.v === '(') { p++; const a = comparar(); if (!a || !es(')')) return null; p++; return a; }
        if (x.t === 'id') {
          p++;
          if (es('(')) {
            p++;
            const args = [];
            if (es(')')) { p++; return { k: 'fn', n: x.v.toUpperCase().replace(/^_XLFN\./, ''), args }; }
            for (;;) {
              if (es(',') || es(';') || es(')')) args.push({ k: 'vacio' }); else { const a = comparar(); if (!a) return null; args.push(a); }
              if (es(',') || es(';')) { p++; continue; }
              if (es(')')) { p++; break; }
              return null;
            }
            return { k: 'fn', n: x.v.toUpperCase().replace(/^_XLFN\./, ''), args };
          }
          const u = x.v.toUpperCase();
          if (u === 'TRUE') return { k: 'lit', v: true };
          if (u === 'FALSE') return { k: 'lit', v: false };
          return { k: 'nombre', n: x.v };
        }
        return null;
      } finally { prof--; }
    }
    const raiz = comparar();
    return raiz && p === tk.length ? raiz : null;
  }

  // ---- Conversión y comparación -------------------------------------------------------------------------
  const aNumero = (v) => {
    if (v === null || v === undefined) return 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (typeof v === 'string') { const s = v.trim(); if (s === '') return ERR.valor; const n = Number(s.replace(/,/g, '')); return Number.isFinite(n) && /^[-+]?(\d|\.\d)/.test(s) ? n : ERR.valor; }
    return v;
  };
  const aTexto = (v) => (v === null || v === undefined ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : typeof v === 'number' ? String(+v.toPrecision(15)) : String(v));
  const aLogico = (v) => (typeof v === 'boolean' ? v : typeof v === 'number' ? v !== 0 : typeof v === 'string' ? (v.toUpperCase() === 'TRUE' ? true : v.toUpperCase() === 'FALSE' ? false : ERR.valor) : v === null || v === undefined ? false : v);
  // número < texto < lógico; el texto sin distinguir mayúsculas; vacía vale 0 contra número, "" contra texto y FALSE contra lógico
  const rango = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3);
  function comparaValores(a, b) {
    if (a === null || a === undefined) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
    if (b === null || b === undefined) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
    const ra = rango(a), rb = rango(b);
    if (ra !== rb) return ra < rb ? -1 : 1;
    if (ra === 1) { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
    return a < b ? -1 : a > b ? 1 : 0;
  }

  // serie de Excel (días desde 1899-12-30, válida desde el 1-mar-1900) ↔ fecha UTC
  const deSerie = (n) => new Date(Date.UTC(1899, 11, 30) + n * 86400000);
  const aSerie = (ms) => Math.round((ms - Date.UTC(1899, 11, 30)) / 86400000);

  // SEARCH con comodines (? = un carácter, * = cualquier cantidad, ~ escapa ? * ~) sin expresiones regulares: una expresión con muchos * colgaba
  // el navegador con texto largo (retroceso exponencial) y el . de una expresión no toma los saltos de línea. Los trozos entre * se buscan de
  // izquierda a derecha, cada uno en la primera posición posible: si el primero aparece en i y los demás no caben después, tampoco caben más
  // adelante. Devuelve la posición (desde 0), -1 si no está o null si la búsqueda pasa del tope de trabajo.
  const MAX_COMPARACIONES = 2000000;
  function buscarComodin(patron, texto, desde) {
    const trozos = [[]];
    for (let k = 0; k < patron.length; k++) {
      const ch = patron[k];
      if (ch === '~' && k + 1 < patron.length && '?*~'.includes(patron[k + 1])) trozos[trozos.length - 1].push(patron[++k]);
      else if (ch === '*') trozos.push([]);
      else trozos[trozos.length - 1].push(ch === '?' ? null : ch);
    }
    let trabajo = 0;
    const buscar = (s, p) => {
      for (let q = p; q + s.length <= texto.length; q++) {
        let k = 0;
        while (k < s.length && (s[k] === null || s[k] === texto[q + k])) k++;
        if ((trabajo += k + 1) > MAX_COMPARACIONES) return null;
        if (k === s.length) return q;
      }
      return -1;
    };
    const ini = buscar(trozos[0], desde);
    if (ini === null || ini < 0) return ini;
    let p = ini + trozos[0].length;
    for (let k = 1; k < trozos.length; k++) {
      const q = buscar(trozos[k], p);
      if (q === null || q < 0) return q;
      p = q + trozos[k].length;
    }
    return ini;
  }

  // ---- Evaluación ---------------------------------------------------------------------------------------
  function evaluar(nodo, ctx) {
    if (!nodo) return ERR.valor;
    ctx.gasto = ctx.gasto || 0;
    return ev(nodo, ctx, 0);
  }
  function celda(r, ctx) { // posición absoluta de una referencia relativa a ctx.origen
    const df = ctx.fila - ctx.origen.f, dc = ctx.col - ctx.origen.c;
    const f = r.af ? r.f : r.f + df, c = r.ac ? r.c : r.c + dc;
    return f < 0 || c < 0 || f > 1048575 || c > 16383 ? null : { f, c };
  }
  function rangoDe(n, ctx) {
    const a = celda(n.a, ctx), b = n.b ? celda(n.b, ctx) : a;
    if (!a || !b) return null;
    return { f0: Math.min(a.f, b.f), c0: Math.min(a.c, b.c), f1: Math.max(a.f, b.f), c1: Math.max(a.c, b.c) };
  }
  // De dónde se leen los valores de una referencia: de la propia hoja (ctx) o, si la referencia trae el nombre de otra, de lo que da ctx.otra(nombre)
  // ({ valor, rango }); null si esa hoja no está disponible.
  function acceso(n, ctx) {
    if (!n.hoja || !ctx.hojaActual) return ctx;
    const nombre = n.hoja.replace(/^'|'$/g, '').replace(/''/g, "'");
    if (nombre.toLowerCase() === ctx.hojaActual.toLowerCase()) return ctx;
    return (ctx.otra && ctx.otra(nombre)) || null;
  }
  function valoresDe(n, ctx) { // lista plana de los valores de un rango (ErrorExcel incluido)
    const r = rangoDe(n, ctx);
    if (!r) return ERR.ref;
    const acc = acceso(n, ctx);
    if (!acc) return ERR.ref;
    if ((r.f1 - r.f0 + 1) * (r.c1 - r.c0 + 1) > MAX_CELDAS_RANGO && !acc.rango) return ERR.valor;
    ctx.gasto += 1;
    if (ctx.gasto > 200000) return ERR.valor;
    return acc.rango(r.f0, r.c0, r.f1, r.c1);
  }
  const planos = (args, ctx, prof) => { // argumentos de una función de agregado: números de rangos (texto y vacías se ignoran) y valores sueltos
    const out = [];
    for (const a of args) {
      if (a.k === 'ref') { const v = valoresDe(a, ctx); if (esError(v)) return v; for (const x of v) { if (esError(x)) return x; if (typeof x === 'number') out.push(x); } continue; }
      const v = ev(a, ctx, prof + 1);
      if (esError(v)) return v;
      if (a.k === 'vacio') continue;
      const n = aNumero(v);
      if (esError(n)) return n;
      out.push(n);
    }
    return out;
  };

  function ev(n, ctx, prof) {
    if (prof > MAX_ARBOL) return ERR.valor;
    switch (n.k) {
      case 'lit': return n.v;
      case 'vacio': return null;
      case 'nombre': return ERR.nombre;
      case 'ref': {
        const acc = acceso(n, ctx);
        if (!acc) return ERR.ref;
        if (n.b) { const v = valoresDe(n, ctx); return esError(v) ? v : v[0] === undefined ? null : v[0]; } // un rango suelto: su primera celda (aprox.)
        const a = celda(n.a, ctx);
        return a ? acc.valor(a.f, a.c) : ERR.ref;
      }
      case 'un': {
        const v = ev(n.a, ctx, prof + 1);
        if (esError(v)) return v;
        const x = aNumero(v);
        if (esError(x)) return x;
        return n.o === '-' ? -x : n.o === '%' ? x / 100 : x;
      }
      case 'bin': {
        const a = ev(n.a, ctx, prof + 1), b = ev(n.b, ctx, prof + 1);
        if (esError(a)) return a;
        if (esError(b)) return b;
        switch (n.o) {
          case '=': return comparaValores(a, b) === 0;
          case '<>': return comparaValores(a, b) !== 0;
          case '<': return comparaValores(a, b) < 0;
          case '>': return comparaValores(a, b) > 0;
          case '<=': return comparaValores(a, b) <= 0;
          case '>=': return comparaValores(a, b) >= 0;
          case '&': { const s = aTexto(a) + aTexto(b); return s.length > 32767 ? ERR.valor : s; } // el texto más largo de Excel: 32.767 caracteres
        }
        const x = aNumero(a), y = aNumero(b);
        if (esError(x)) return x;
        if (esError(y)) return y;
        switch (n.o) {
          case '+': return x + y;
          case '-': return x - y;
          case '*': return x * y;
          case '/': return y === 0 ? ERR.div : x / y;
          case '^': { const r = Math.pow(x, y); return Number.isFinite(r) ? r : ERR.num; }
        }
        return ERR.valor;
      }
      case 'fn': return funcion(n, ctx, prof);
    }
    return ERR.valor;
  }

  function funcion(n, ctx, prof) {
    const A = n.args;
    const val = (i) => ev(A[i], ctx, prof + 1);
    const num = (i) => { const v = val(i); return esError(v) ? v : aNumero(v); };
    const txt = (i) => { const v = val(i); return esError(v) ? v : aTexto(v); };
    const nargs = (min, max) => A.length >= min && A.length <= (max === undefined ? min : max);
    switch (n.n) {
      case 'AND': case 'OR': {
        let r = n.n === 'AND', visto = false;
        for (const a of A) {
          const vs = a.k === 'ref' ? valoresDe(a, ctx) : [ev(a, ctx, prof + 1)];
          if (esError(vs)) return vs;
          for (const v of vs) {
            if (esError(v)) return v;
            if (a.k === 'ref' && (typeof v === 'string' || v === null)) continue; // en rangos el texto y las vacías se ignoran
            const l = aLogico(v);
            if (esError(l)) return l;
            visto = true;
            if (n.n === 'AND') r = r && l; else r = r || l;
          }
        }
        return visto ? r : ERR.valor;
      }
      case 'NOT': { if (!nargs(1)) return ERR.valor; const v = val(0); if (esError(v)) return v; const l = aLogico(v); return esError(l) ? l : !l; }
      case 'IF': { if (!nargs(2, 3)) return ERR.valor; const c = val(0); if (esError(c)) return c; const l = aLogico(c); if (esError(l)) return l; return l ? val(1) : A.length > 2 ? val(2) : false; }
      case 'IFERROR': { if (!nargs(2)) return ERR.valor; const v = val(0); return esError(v) ? val(1) : v; }
      case 'TRUE': return true;
      case 'FALSE': return false;
      case 'ROW': { if (A.length && A[0].k === 'ref') { const a = celda(A[0].a, ctx); return a ? a.f + 1 : ERR.ref; } return ctx.fila + 1; }
      case 'COLUMN': { if (A.length && A[0].k === 'ref') { const a = celda(A[0].a, ctx); return a ? a.c + 1 : ERR.ref; } return ctx.col + 1; }
      case 'ISBLANK': { if (!nargs(1)) return ERR.valor; const v = val(0); return v === null; }
      case 'ISNUMBER': { if (!nargs(1)) return ERR.valor; return typeof val(0) === 'number'; }
      case 'ISTEXT': { if (!nargs(1)) return ERR.valor; return typeof val(0) === 'string'; }
      case 'ISNONTEXT': { if (!nargs(1)) return ERR.valor; return typeof val(0) !== 'string'; }
      case 'ISLOGICAL': { if (!nargs(1)) return ERR.valor; return typeof val(0) === 'boolean'; }
      case 'ISERROR': { if (!nargs(1)) return ERR.valor; return esError(val(0)); }
      case 'ISERR': { if (!nargs(1)) return ERR.valor; const v = val(0); return esError(v) && v.codigo !== '#N/A'; }
      case 'ISNA': { if (!nargs(1)) return ERR.valor; const v = val(0); return esError(v) && v.codigo === '#N/A'; }
      case 'ISEVEN': case 'ISODD': { const x = num(0); if (esError(x)) return x; const par = Math.floor(Math.abs(x)) % 2 === 0; return n.n === 'ISEVEN' ? par : !par; }
      case 'MOD': { if (!nargs(2)) return ERR.valor; const a = num(0), b = num(1); if (esError(a)) return a; if (esError(b)) return b; return b === 0 ? ERR.div : a - b * Math.floor(a / b); }
      case 'ABS': { const x = num(0); return esError(x) ? x : Math.abs(x); }
      case 'INT': { const x = num(0); return esError(x) ? x : Math.floor(x); }
      case 'ROUND': { const x = num(0), d = A.length > 1 ? num(1) : 0; if (esError(x)) return x; if (esError(d)) return d; const k = Math.pow(10, Math.trunc(d)); return Math.sign(x) * Math.round(Math.abs(x) * k + 1e-9) / k; }
      case 'FLOOR': { const x = num(0), s = A.length > 1 ? num(1) : 1; if (esError(x)) return x; if (esError(s)) return s; return s === 0 ? 0 : Math.floor(x / s) * s; }
      case 'LEN': { const t = txt(0); return esError(t) ? t : t.length; }
      case 'TRIM': { const t = txt(0); return esError(t) ? t : t.replace(/^ +| +$/g, '').replace(/ {2,}/g, ' '); }
      case 'UPPER': { const t = txt(0); return esError(t) ? t : t.toUpperCase(); }
      case 'LOWER': { const t = txt(0); return esError(t) ? t : t.toLowerCase(); }
      case 'LEFT': { const t = txt(0), k = A.length > 1 ? num(1) : 1; if (esError(t)) return t; if (esError(k)) return k; return k < 0 ? ERR.valor : t.slice(0, Math.floor(k)); }
      case 'RIGHT': { const t = txt(0), k = A.length > 1 ? num(1) : 1; if (esError(t)) return t; if (esError(k)) return k; return k < 0 ? ERR.valor : Math.floor(k) === 0 ? '' : t.slice(-Math.floor(k)); }
      case 'MID': { const t = txt(0), s = num(1), k = num(2); if (esError(t)) return t; if (esError(s)) return s; if (esError(k)) return k; return s < 1 || k < 0 ? ERR.valor : t.substr(Math.floor(s) - 1, Math.floor(k)); }
      case 'SEARCH': case 'FIND': {
        const b = txt(0), t = txt(1), ini = A.length > 2 ? num(2) : 1;
        if (esError(b)) return b; if (esError(t)) return t; if (esError(ini)) return ini;
        if (ini < 1 || ini > t.length + 1) return ERR.valor;
        let i;
        if (n.n === 'FIND') i = t.indexOf(b, ini - 1);
        else { i = buscarComodin(b.toLowerCase(), t.toLowerCase(), ini - 1); if (i === null) return ERR.valor; } // SEARCH: sin mayúsculas y con comodines ? y *
        return i < 0 ? ERR.valor : i + 1;
      }
      case 'EXACT': { const a = txt(0), b = txt(1); return esError(a) ? a : esError(b) ? b : a === b; }
      case 'VALUE': { const v = val(0); if (esError(v)) return v; return aNumero(v); }
      case 'TODAY': return Math.floor(ctx.hoy !== undefined ? ctx.hoy : 0);
      case 'NOW': return ctx.hoy !== undefined ? ctx.hoy : 0;
      case 'YEAR': case 'MONTH': case 'DAY': { const x = num(0); if (esError(x)) return x; if (x < 0) return ERR.num; const d = deSerie(Math.floor(x)); return n.n === 'YEAR' ? d.getUTCFullYear() : n.n === 'MONTH' ? d.getUTCMonth() + 1 : d.getUTCDate(); }
      case 'EDATE': { const x = num(0), m = num(1); if (esError(x)) return x; if (esError(m)) return m; const d = deSerie(Math.floor(x)); const t = d.getUTCMonth() + Math.trunc(m); const a = d.getUTCFullYear() + Math.floor(t / 12), mes = ((t % 12) + 12) % 12; const ultimo = new Date(Date.UTC(a, mes + 1, 0)).getUTCDate(); return aSerie(Date.UTC(a, mes, Math.min(d.getUTCDate(), ultimo))); }
      case 'DATE': { const a = num(0), m = num(1), d = num(2); if (esError(a)) return a; if (esError(m)) return m; if (esError(d)) return d; return aSerie(Date.UTC(a < 1900 ? a + 1900 : a, Math.trunc(m) - 1, Math.trunc(d))); }
      case 'ROUNDDOWN': { const x = num(0), d = A.length > 1 ? num(1) : 0; if (esError(x)) return x; if (esError(d)) return d; const k = Math.pow(10, Math.trunc(d)); return Math.trunc(x * k) / k; }
      case 'WEEKDAY': { const s = num(0), t = A.length > 1 ? num(1) : 1; if (esError(s)) return s; if (esError(t)) return t; const d = ((Math.floor(s) - 1) % 7 + 7) % 7; const dom = (d + 0) % 7; return t === 2 ? ((dom + 6) % 7) + 1 : t === 3 ? (dom + 6) % 7 : dom + 1; }
      case 'SUM': case 'AVERAGE': case 'MIN': case 'MAX': case 'COUNT': {
        const xs = planos(A, ctx, prof);
        if (esError(xs)) return xs;
        if (n.n === 'COUNT') return xs.length;
        if (n.n === 'SUM') return xs.reduce((u, v) => u + v, 0);
        if (!xs.length) return n.n === 'AVERAGE' ? ERR.div : 0;
        if (n.n === 'AVERAGE') return xs.reduce((u, v) => u + v, 0) / xs.length;
        return n.n === 'MIN' ? xs.reduce((u, v) => (v < u ? v : u), Infinity) : xs.reduce((u, v) => (v > u ? v : u), -Infinity);
      }
      case 'COUNTIF': {
        if (!nargs(2) || A[0].k !== 'ref') return ERR.valor;
        const vs = valoresDe(A[0], ctx), crit = val(1);
        if (esError(vs)) return vs; if (esError(crit)) return crit;
        let k = 0;
        for (const v of vs) if (typeof crit === 'string' ? (typeof v === 'string' && v.toLowerCase() === crit.toLowerCase()) : (typeof v === 'number' && v === crit)) k++;
        return k;
      }
    }
    return ERR.nombre;
  }

  globalThis.XlsxExpresion = { compilar, evaluar, ErrorExcel, ERR, esError, comparaValores, aNumero, aTexto };
})();

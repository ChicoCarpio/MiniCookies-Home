/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// La condición de una validación de datos de ODF (`table:condition="of:cell-content-is-between(1;10)"`) ↔ la regla del modelo de Excel
// ({ tipo, operador, formula1, formula2 }). Sin DOM. Requiere ods-formulas.
//
//   OdsValidacion.leer(condicion) → { tipo, operador, formula1, formula2 } | null     (las fórmulas en sintaxis de Excel; los textos de una lista como "a,b,c")
//   OdsValidacion.escribir(regla) → condicion | null                                     (sin escapar para XML)
(function () {
  const OF = () => globalThis.OdsFormulas;
  const TIPOS = { 'whole-number': 'whole', 'decimal-number': 'decimal', date: 'date', time: 'time' };
  const TIPOS_ODF = { whole: 'whole-number', decimal: 'decimal-number', date: 'date', time: 'time' };
  const OPERADORES = { '=': 'equal', '!=': 'notEqual', '>': 'greaterThan', '<': 'lessThan', '>=': 'greaterThanOrEqual', '<=': 'lessThanOrEqual' };
  const OPERADORES_ODF = { equal: '=', notEqual: '!=', greaterThan: '>', lessThan: '<', greaterThanOrEqual: '>=', lessThanOrEqual: '<=' };

  // separa los argumentos de una llamada por «;» fuera de comillas, corchetes y paréntesis
  function argumentos(t) {
    const salida = []; let actual = '', nivel = 0, comillas = false, corchetes = 0;
    for (let i = 0; i < t.length; i++) {
      const ch = t[i];
      if (ch === '"') { comillas = !comillas; actual += ch; continue; }
      if (!comillas) {
        if (ch === '[') corchetes++; else if (ch === ']') corchetes--;
        else if (ch === '(') nivel++; else if (ch === ')') nivel--;
        else if ((ch === ';' || ch === ',') && !nivel && !corchetes) { salida.push(actual); actual = ''; continue; }
      }
      actual += ch;
    }
    if (actual !== '' || salida.length) salida.push(actual);
    return salida;
  }
  const aExcel = (expr) => { const x = OF().aExcel('of:=' + expr.trim()); return x === null || x === undefined ? expr.trim() : x; };
  const aOds = (texto) => { const x = OF().aOpenFormula(String(texto).replace(/^=/, '')); return x === null || x === undefined ? null : x; };

  function leer(condicion) {
    const c = String(condicion || '').replace(/\bof:/g, '').trim();
    let m;
    if ((m = /^cell-content-is-in-list\(([\s\S]*)\)$/.exec(c))) {
      const args = argumentos(m[1]);
      if (args.length && args.every((a) => /^"(?:[^"]|"")*"$/.test(a.trim()))) return { tipo: 'list', operador: 'between', formula1: '"' + args.map((a) => a.trim().slice(1, -1).replace(/""/g, '"')).join(',') + '"', formula2: null };
      return { tipo: 'list', operador: 'between', formula1: aExcel(m[1]), formula2: null };
    }
    if ((m = /^is-true-formula\(([\s\S]*)\)$/.exec(c))) return { tipo: 'custom', operador: 'between', formula1: aExcel(m[1]), formula2: null };
    let tipo = 'decimal', resto = c;
    if ((m = /^cell-content-is-(whole-number|decimal-number|date|time)\(\)\s+and\s+([\s\S]*)$/.exec(c))) { tipo = TIPOS[m[1]]; resto = m[2]; }
    else if (/^cell-content-text-length/.test(c)) tipo = 'textLength';
    else if (/^cell-content-is-(whole-number|decimal-number|date|time)\(\)$/.test(c)) return { tipo: TIPOS[/is-([a-z-]+)\(\)/.exec(c)[1]], operador: 'between', formula1: null, formula2: null };
    if ((m = /^cell-content(?:-text-length)?-is-(not-)?between\(([\s\S]*)\)$/.exec(resto))) {
      const a = argumentos(m[2]);
      if (a.length !== 2) return null;
      return { tipo, operador: m[1] ? 'notBetween' : 'between', formula1: aExcel(a[0]), formula2: aExcel(a[1]) };
    }
    if ((m = /^cell-content(?:-text-length)?\(\)\s*(>=|<=|!=|=|>|<)\s*([\s\S]+)$/.exec(resto))) return { tipo, operador: OPERADORES[m[1]], formula1: aExcel(m[2]), formula2: null };
    return null;
  }

  function escribir(regla) {
    if (!regla || !regla.tipo || regla.tipo === 'none') return null;
    const f1 = regla.formula1 === null || regla.formula1 === undefined ? '' : String(regla.formula1).replace(/^=/, ''), f2 = regla.formula2 === null || regla.formula2 === undefined ? '' : String(regla.formula2).replace(/^=/, '');
    const odf = (t) => { const x = aOds(t); if (x === null) throw new Error('la fórmula de la validación no se puede escribir en .ods: ' + t); return x; };
    if (regla.tipo === 'list') {
      if (/^".*"$/.test(f1.trim())) return 'of:cell-content-is-in-list(' + f1.trim().slice(1, -1).replace(/""/g, '"').split(',').map((x) => '"' + x.trim().replace(/"/g, '""') + '"').join(';') + ')';
      return 'of:cell-content-is-in-list(' + odf(f1) + ')';
    }
    if (regla.tipo === 'custom') return 'of:is-true-formula(' + odf(f1) + ')';
    const texto = regla.tipo === 'textLength';
    const base = texto ? 'cell-content-text-length' : 'cell-content';
    let cuerpo;
    if (regla.operador === 'between' || regla.operador === 'notBetween') cuerpo = base + '-is-' + (regla.operador === 'notBetween' ? 'not-' : '') + 'between(' + odf(f1) + ',' + odf(f2) + ')';   // con «,»: es como lo escribe y lo lee LibreOffice
    else cuerpo = base + '()' + (OPERADORES_ODF[regla.operador] || '=') + odf(f1);
    return texto ? 'of:' + cuerpo : 'of:cell-content-is-' + TIPOS_ODF[regla.tipo] + '() and ' + cuerpo;   // «of:» solo en la primera parte, como LibreOffice
  }

  globalThis.OdsValidacion = { leer, escribir };
})();

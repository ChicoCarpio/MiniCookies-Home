/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Formato condicional de ODF (las extensiones `calcext:` de LibreOffice) ↔ las reglas del modelo y del diálogo de Mini Cookies Hojas (las mismas de Excel). Sin DOM.
// La condición va en `calcext:value`: `>5`, `between(1,10)`, `formula-is(expr)`, `contains-text("x")`, `duplicate`, `above-average`, `top-elements(3)`…; las escalas de color,
// barras y conjuntos de iconos son elementos aparte (`calcext:color-scale`, `data-bar`, `icon-set`). Requiere ods-formulas.
//
//   OdsCondicional.leerValor(valor) → { tipo, operador?, formulas?, texto?, rango?, porcentaje?, abajo?, encima?, igual? } | null        (fórmulas en sintaxis de Excel)
//   OdsCondicional.escribirValor(regla) → texto de calcext:value | null
//   OdsCondicional.TIPOS_CFVO / aCfvo(tipoOds) / deCfvo(tipoXlsx)    (escalas, barras e iconos)
(function () {
  const OF = () => globalThis.OdsFormulas;
  const OPERADORES = { '>': 'greaterThan', '<': 'lessThan', '>=': 'greaterThanOrEqual', '<=': 'lessThanOrEqual', '=': 'equal', '!=': 'notEqual' };
  const OPERADORES_ODF = { greaterThan: '>', lessThan: '<', greaterThanOrEqual: '>=', lessThanOrEqual: '<=', equal: '=', notEqual: '!=' };

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
  const aExcel = (expr) => { const x = OF().aExcel('of:=' + String(expr).trim()); return x === null || x === undefined ? String(expr).trim() : x; };
  const aOds = (texto) => { const x = OF().aOpenFormula(String(texto).replace(/^=/, '')); if (x === null || x === undefined) throw new Error('la fórmula no se puede escribir en .ods: ' + texto); return x; };
  const cadena = (a) => { const m = /^\s*"((?:[^"]|"")*)"\s*$/.exec(a); return m ? m[1].replace(/""/g, '"') : null; };

  function leerValor(valor) {
    const v = String(valor || '').trim();
    let m;
    if ((m = /^(>=|<=|!=|=|>|<)\s*([\s\S]+)$/.exec(v))) return { tipo: 'cellIs', operador: OPERADORES[m[1]], formulas: [aExcel(m[2])] };
    if ((m = /^(not-)?between\(([\s\S]*)\)$/.exec(v))) { const a = argumentos(m[2]); return a.length === 2 ? { tipo: 'cellIs', operador: m[1] ? 'notBetween' : 'between', formulas: [aExcel(a[0]), aExcel(a[1])] } : null; }
    if ((m = /^formula-is\(([\s\S]*)\)$/.exec(v))) return { tipo: 'expression', formulas: [aExcel(m[1])] };
    if ((m = /^(not-contains-text|contains-text|begins-with|ends-with)\(([\s\S]*)\)$/.exec(v))) {
      const t = cadena(m[2]);
      return { tipo: { 'contains-text': 'containsText', 'not-contains-text': 'notContainsText', 'begins-with': 'beginsWith', 'ends-with': 'endsWith' }[m[1]], texto: t === null ? aExcel(m[2]) : t, formulas: [] };
    }
    if (v === 'duplicate') return { tipo: 'duplicateValues', formulas: [] };
    if (v === 'unique') return { tipo: 'uniqueValues', formulas: [] };
    if ((m = /^(above|below)-(equal-)?average(?:\(\))?$/.exec(v))) return { tipo: 'aboveAverage', encima: m[1] === 'above', igual: !!m[2], formulas: [] };
    if ((m = /^(top|bottom)-(elements|percent)\(\s*(\d+)\s*\)$/.exec(v))) return { tipo: 'top10', abajo: m[1] === 'bottom', porcentaje: m[2] === 'percent', rango: +m[3], formulas: [] };
    if (v === 'is-error') return { tipo: 'containsErrors', formulas: [] };
    if (v === 'is-no-error') return { tipo: 'notContainsErrors', formulas: [] };
    return null;
  }

  function escribirValor(r) {
    if (!r || !r.tipo) return null;
    const f = (k) => aOds((r.formulas || [])[k] === undefined ? '' : r.formulas[k]);
    switch (r.tipo) {
      case 'cellIs':
        if (r.operador === 'between' || r.operador === 'notBetween') return (r.operador === 'notBetween' ? 'not-' : '') + 'between(' + f(0) + ',' + f(1) + ')';
        return (OPERADORES_ODF[r.operador] || '=') + f(0);
      case 'expression': return 'formula-is(' + f(0) + ')';
      case 'containsText': case 'notContainsText': case 'beginsWith': case 'endsWith': {
        const nombre = { containsText: 'contains-text', notContainsText: 'not-contains-text', beginsWith: 'begins-with', endsWith: 'ends-with' }[r.tipo];
        return nombre + '("' + String(r.texto === undefined || r.texto === null ? '' : r.texto).replace(/"/g, '""') + '")';
      }
      case 'duplicateValues': return 'duplicate';
      case 'uniqueValues': return 'unique';
      case 'aboveAverage': return (r.encima === false ? 'below' : 'above') + (r.igual ? '-equal' : '') + '-average';
      case 'top10': return (r.abajo ? 'bottom' : 'top') + '-' + (r.porcentaje ? 'percent' : 'elements') + '(' + (r.rango || 10) + ')';
      case 'containsErrors': return 'is-error';
      case 'notContainsErrors': return 'is-no-error';
      case 'containsBlanks': return 'formula-is(LEN(TRIM(' + '[.#BASE#]' + '))=0)';   // el escritor lo completa con la celda base
      case 'notContainsBlanks': return 'formula-is(LEN(TRIM(' + '[.#BASE#]' + '))>0)';
      default: return null;
    }
  }

  // tipos de los puntos de una escala, barra o conjunto de iconos
  const A_XLSX = { minimum: 'min', maximum: 'max', percentile: 'percentile', percent: 'percent', number: 'num', formula: 'formula', 'auto-minimum': 'min', 'auto-maximum': 'max' };
  const A_ODS = { min: 'minimum', max: 'maximum', percentile: 'percentile', percent: 'percent', num: 'number', formula: 'formula' };
  const aCfvo = (t) => A_XLSX[t] || 'num';
  const deCfvo = (t) => A_ODS[t] || 'number';

  globalThis.OdsCondicional = { leerValor, escribirValor, aCfvo, deCfvo, argumentos };
})();

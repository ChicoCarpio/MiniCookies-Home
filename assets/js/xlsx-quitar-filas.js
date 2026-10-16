/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Quitar filas de una hoja EN EL LUGAR (fase C del motor de Excel), como "Eliminar filas" de Excel: las filas de abajo suben, y todo lo que apunta a
// la hoja se acomoda: fórmulas de todas las hojas (una referencia a una fila quitada pasa a #REF!, un rango se achica), celdas combinadas, formato
// condicional, validación de datos, hipervínculos, autofiltro, tablas, nombres definidos, gráficos, saltos de página y anclas de los dibujos.
// El resto del paquete (estilos, imágenes, macros, otras hojas) queda tal cual. No depende del navegador.
//
//   const r = await XlsxEscritor.quitarFilas(libro, hoja, [filas desde 0]);   // libro: LibroXlsx.abrir; hoja: índice o nombre
//   // r = { quitadas: n }  · después se guarda con libro.zip.generateAsync(...)
// Rechaza (ErrorEscritor, sin tocar nada) lo que no sabe mover sin dejar un archivo que Excel repare: comentarios y notas, tablas dinámicas y
// segmentaciones de la hoja o con ella de origen, fórmulas matriciales y tablas de datos, controles y objetos OLE, consultas, extensiones x14 de
// formato condicional, validación o minigráficos, el encabezado o la fila de totales de una tabla, una tabla que quedaría sin filas, hojas
// escritas con prefijo de espacio de nombres (<x:row>) y formato condicional con fórmulas relativas que obligaría a partirlo en demasiadas reglas.
// Las fórmulas compartidas de las hojas afectadas se escriben una a una (al quitar filas el desplazamiento entre la maestra y sus copias cambia).
// Requiere xlsx-modelo.js, xml-editar.js, xlsx-paquete.js y xlsx-escritor.js.
(function () {
  const X = XmlEditar;
  const { ErrorEscritor } = XlsxEscritor;
  const colANumero = (s) => LibroXlsx.colANumero(s);
  const numeroACol = (n) => LibroXlsx.numeroACol(n);
  const IDENT = /[A-Za-z0-9_.À-￿]/;
  const MAX_PARTES = 5000;
  // elementos de la hoja que Excel nunca escribe con prefijo (sí escribe <xm:f>, <x14:dataValidation>…)
  const PREFIJADOS = /<[A-Za-z_][\w.-]*:(?:sheetData|row|c|mergeCells|hyperlinks)[\s/>]/;

  // ---- qué fila va a dónde ---------------------------------------------------------------------------------------
  function mapeo(filas) {
    const q = filas;   // ordenadas y sin repetir
    const hasta = (f, incluye) => { let lo = 0, hi = q.length; while (lo < hi) { const m = (lo + hi) >> 1; if (incluye ? q[m] <= f : q[m] < f) lo = m + 1; else hi = m; } return lo; };
    const finRacha = new Int32Array(q.length);   // índice de la última fila de cada tramo de filas quitadas seguidas
    for (let k = q.length - 1; k >= 0; k--) finRacha[k] = k + 1 < q.length && q[k + 1] === q[k] + 1 ? finRacha[k + 1] : k;
    return {
      lista: q,
      // la primera fila desde f (incluida) que no se quita
      siguiente(f) { const k = hasta(f, false); return k < q.length && q[k] === f ? q[finRacha[k]] + 1 : f; },
      // fila nueva, o -1 si se quitó
      fila(f) { const k = hasta(f, false); return k < q.length && q[k] === f ? -1 : f - k; },
      // rango nuevo [a, b], o null si todas sus filas se quitaron: el inicio sube a la primera fila que queda, el final baja a la última
      rango(f0, f1) { const a = f0 - hasta(f0, false), b = f1 - hasta(f1, true); return a > b ? null : [a, b]; },
      quita: (f0, f1) => hasta(f1, true) - hasta(f0, false),
      antes: (f) => hasta(f, false),   // filas quitadas por encima de f
      hasta: (f) => hasta(f, true),    // filas quitadas hasta f, incluida
    };
  }

  // ---- referencias dentro de una fórmula --------------------------------------------------------------------------
  // `enObjetivo`: la fórmula vive en la hoja de la que se quitan filas (sus referencias sin hoja apuntan a ella). `objetivo`: el nombre de esa hoja.
  const REF_CELDA = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+))?(?![A-Za-z0-9_(.!])/;
  const REF_FILAS = /^(\$?)(\d+):(\$?)(\d+)(?![A-Za-z0-9_(.!:])/;
  const PREFIJO = /^('(?:[^']|'')*'|[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*)!/;
  function reescribirFormula(texto, enObjetivo, objetivo, M) {
    if (!texto || !/\d/.test(texto)) return texto;
    const destino = objetivo.toLowerCase();
    let salida = '', i = 0, externo = false;
    const n = texto.length;
    while (i < n) {
      const ch = texto[i];
      if (ch === '"') { let j = i + 1; while (j < n) { if (texto[j] === '"') { if (texto[j + 1] === '"') j += 2; else break; } else j++; } salida += texto.slice(i, j + 1); i = j + 1; continue; }
      if (ch === '[') {
        let nivel = 0, j = i;
        for (; j < n; j++) { if (texto[j] === '[') nivel++; else if (texto[j] === ']') { nivel--; if (nivel === 0) break; } }
        const grupo = texto.slice(i, j + 1);
        salida += grupo; i = j + 1;
        if (/^\[\d+\]$/.test(grupo)) externo = true;   // [1]Hoja!A1: una hoja de otro libro
        continue;
      }
      const previo = i > 0 ? texto[i - 1] : '';
      if ((/[A-Za-z_$'\dÀ-￿]/.test(ch)) && !(previo && (IDENT.test(previo) || previo === '!'))) {
        let prefijo = '', hoja = null, pos = i;
        const mp = PREFIJO.exec(texto.slice(i, i + 80));
        if (mp) { prefijo = mp[0]; hoja = mp[1][0] === "'" ? mp[1].slice(1, -1).replace(/''/g, "'") : mp[1]; pos = i + prefijo.length; }
        const resto = texto.slice(pos, pos + 60);
        let m = REF_CELDA.exec(resto);
        if (m && (colANumero(m[2].toUpperCase()) > 16383 || (m[6] && colANumero(m[6].toUpperCase()) > 16383))) m = null;
        const mf = m ? null : REF_FILAS.exec(resto);
        if (!m && !mf) {
          externo = false;
          if (prefijo) { salida += prefijo; i = pos; continue; }
          salida += ch; i++; continue;
        }
        const todo = (m || mf)[0];
        // Hoja1:Hoja3!A1 (referencia 3D): Excel no la toca al quitar filas de una de sus hojas (medido)
        const aplica = !externo && !(prefijo && previo === ':') && (prefijo ? hoja.toLowerCase() === destino : enObjetivo);
        externo = false;
        if (!aplica) { salida += prefijo + todo; i = pos + todo.length; continue; }
        // M.fila y M.rango reciben además qué extremos son relativos (sin $ en la fila): lo usa partirPorFilas
        let nuevo;
        if (m) {
          const f0 = parseInt(m[4], 10) - 1;
          if (!m[8]) { const f = M.fila(f0, !m[3]); nuevo = f < 0 ? '#REF!' : m[1] + m[2] + m[3] + (f + 1); }
          else {
            const f1 = parseInt(m[8], 10) - 1, r = M.rango(Math.min(f0, f1), Math.max(f0, f1), [[f0, !m[3]], [f1, !m[7]]]);
            nuevo = r ? m[1] + m[2] + m[3] + (r[0] + 1) + ':' + m[5] + m[6] + m[7] + (r[1] + 1) : '#REF!';
          }
        } else {
          const a = parseInt(mf[2], 10) - 1, b = parseInt(mf[4], 10) - 1, r = M.rango(Math.min(a, b), Math.max(a, b), [[a, !mf[1]], [b, !mf[3]]]);
          nuevo = r ? mf[1] + (r[0] + 1) + ':' + mf[3] + (r[1] + 1) : '#REF!';
        }
        salida += prefijo + nuevo; i = pos + todo.length; continue;
      }
      if (ch !== '!') externo = false;   // [1]!Nombre: el libro externo vale hasta el final del nombre
      salida += ch; i++;
    }
    return salida;
  }

  // Una lista de rangos de celdas ("A1:B3 D5", sqref de formato condicional o validación) → la lista nueva ('' si no queda ninguno).
  // soloVacios: true descarta un rango que queda en una sola celda si era un combinado (mergeCell).
  function reescribirSqref(lista, M) {
    const salida = [];
    for (const parte of String(lista).split(/\s+/).filter(Boolean)) {
      const r = LibroXlsx.decodificarRango(parte.replace(/\$/g, ''));
      if (!r) { salida.push(parte); continue; }
      const nf = M.rango(r.f0, r.f1);
      if (!nf) continue;
      const a = numeroACol(r.c0) + (nf[0] + 1), b = numeroACol(r.c1) + (nf[1] + 1);
      salida.push(a === b ? a : a + ':' + b);
    }
    return salida.join(' ');
  }

  // ---- formato condicional y validación: fórmulas relativas a la esquina del sqref -------------------------------
  // Lo relativo de estas fórmulas cuenta desde la esquina superior izquierda de todo el sqref. Al quitar filas Excel parte la regla en varias
  // (medido, corpus 45-quitar-filas-cf, con un solo "Eliminar" de varias filas): recorre las filas que quedan en tramos seguidos (sin filas
  // quitadas en medio) y, dentro de cada tramo, en rachas donde cada extremo de referencia sigue igual de "quitado o no"; la primera celda de
  // cada racha recibe su fórmula exacta (como una fórmula de celda) y las demás de la racha la copian relativa. Las rachas con la misma fórmula
  // relativa vuelven a ser una sola regla. Una racha solo puede cambiar cerca de una fila quitada (a la distancia de cada referencia relativa),
  // así que se evalúa una vez por tramo y no una vez por fila.
  // → [{ sqref, formulas }] (vacío si no queda ninguna celda). `cuenta` = { n } lleva las evaluaciones de toda la hoja (tope MAX_EVALUACIONES).
  // enObjetivo = false: la regla es de OTRA hoja y solo sus referencias con el nombre de la hoja editada se mueven; sus celdas no se quitan. Excel la
  // parte igual (medido, corpus 49: A1:A10 con Datos!A1 al borrar las filas 3 y 4 de Datos → A1:A2 Datos!A1, A3:A4 Datos!#REF!, A5:A10 Datos!A3).
  const MAX_EVALUACIONES = 1000000;
  const IDENTIDAD = { fila: (f) => f, siguiente: (f) => f, rango: (a, b) => [a, b] };
  function partirPorFilas(sqref, formulas, M, objetivo, cuenta, enObjetivo = true) {
    const P = enObjetivo ? M : IDENTIDAD;   // qué pasa con las filas de la propia regla
    const partes = String(sqref).split(/\s+/).filter(Boolean).map((p) => LibroXlsx.decodificarRango(p.replace(/\$/g, '')));
    if (!partes.length || partes.some((r) => !r)) {   // un sqref que no se entiende: se acomoda como un rango cualquiera
      const sq = enObjetivo ? reescribirSqref(sqref, M) : sqref;
      return sq ? [{ sqref: sq, formulas: formulas.map((f) => reescribirFormula(f, enObjetivo, objetivo, M)) }] : [];
    }
    let af = Infinity, ac = Infinity;
    for (const r of partes) { af = Math.min(af, r.f0); ac = Math.min(ac, r.c0); }
    // desplazamientos de las filas relativas (desde la esquina) y filas absolutas que citan las fórmulas
    const desp = new Set(), abs = new Set();
    const anota = (f, rel) => { if (rel) desp.add(f - af); else abs.add(f); };
    const espia = { fila: (f, rel) => { anota(f, rel); return f; }, rango: (a, b, ext) => { for (const [f, rel] of ext) anota(f, rel); return [a, b]; } };
    for (const f of formulas) reescribirFormula(f, enObjetivo, objetivo, espia);
    if (enObjetivo) desp.delete(0);   // la fila de la propia celda nunca se quita: no cambia nada (en otra hoja, la misma fila de la hoja editada sí)
    // la fila r (que queda): su racha (qué extremos de referencia se quitan) y su fórmula exacta en su fila nueva (`fila`), en la columna de la esquina.
    // (No se escribe "desde la esquina vieja": en otra hoja las referencias suben y la celda no, y desde la esquina la fila quedaría antes de la 1.)
    let fija = null;
    // el tope es de trabajo, no de veces: una fórmula larga cuenta por varias evaluaciones (una de 8.192 caracteres, el máximo de Excel, por 513)
    const largo = formulas.reduce((s, f) => s + f.length, 0);
    if (formulas.some((f) => f.length > 8192)) throw new ErrorEscritor('una fórmula de formato condicional o validación más larga que el máximo de Excel');
    const peso = 1 + (largo >> 4);
    const evaluar = (r) => {
      if (!desp.size && fija) return fija;
      if ((cuenta.n += peso) > MAX_EVALUACIONES) throw new ErrorEscritor('demasiado formato condicional o validación con fórmulas relativas para acomodar');
      let racha = '';
      const Mr = {
        fila: (f) => { const n = M.fila(f); racha += n < 0 ? 'q' : '-'; return n; },
        rango: (a, b, ext) => { const n = M.rango(a, b); for (const [f] of ext) racha += M.fila(f) < 0 ? 'q' : '-'; racha += n ? '-' : 'x'; return n; },
      };
      const fx = formulas.map((f) => reescribirFormula(LibroXlsx.desplazarFormula(f, r - af, 0), enObjetivo, objetivo, Mr));
      fija = { racha, fx, fila: P.fila(r) };
      return fija;
    };
    // dos tramos con la misma fórmula relativa van en la misma regla: se comparan llevadas a una fila de referencia (si alguna referencia se saldría
    // de la hoja al llevarla, el tramo queda en su propia regla)
    const BASE = 524288;
    const claveDe = (fx, fila) => {
      const k = fx.map((f) => LibroXlsx.desplazarFormula(f, BASE - fila, 0));
      return k.every((f, i) => LibroXlsx.desplazarFormula(f, fila - BASE, 0) === fx[i]) ? k.join('\u0000') : '\u0001' + fila + '\u0000' + fx.join('\u0000');
    };
    // filas donde la fórmula puede cambiar: cerca de cada fila quitada, a la distancia de cada desplazamiento
    let cortes = [];
    if (desp.size) {
      let lo = Infinity, hi = -Infinity, dmax = 0;
      for (const r of partes) { lo = Math.min(lo, r.f0); hi = Math.max(hi, r.f1); }
      for (const d of desp) dmax = Math.max(dmax, Math.abs(d));
      const q = M.lista, k0 = M.antes(lo - dmax - 1), k1 = M.hasta(hi + dmax + 1);
      if ((k1 - k0 + abs.size) * (desp.size + 1) * 2 > MAX_EVALUACIONES * 4) throw new ErrorEscritor('demasiado formato condicional o validación con fórmulas relativas para acomodar');
      for (let k = k0; k < k1; k++) { cortes.push(q[k], q[k] + 1); for (const d of desp) cortes.push(q[k] - d, q[k] - d + 1); }
      for (const a of abs) for (const d of desp) cortes.push(a - d, a - d + 1);
      cortes = [...new Set(cortes)].sort((a, b) => a - b);
    }
    const grupos = new Map();   // fórmulas → { formulas, tramos: [{ f0, f1, c0, c1 }] }
    const porIntervalo = new Map();   // índice del intervalo entre cortes → evaluación
    for (const r of partes) {
      let k = 0, actual = null;   // actual: la racha en curso { racha, fx y fila (los de su primera celda) }
      for (let hi = cortes.length; k < hi;) { const m = (k + hi) >> 1; if (cortes[m] <= r.f0) k = m + 1; else hi = m; }   // el primer corte después de f0
      for (let s = r.f0; s <= r.f1;) {
        const entre = k;   // las filas de este tramo están entre los mismos cortes: misma racha y misma fórmula relativa
        const e = k < cortes.length && cortes[k] <= r.f1 ? cortes[k++] - 1 : r.f1;
        const x = P.siguiente(s);
        if (x > e) { actual = null; s = e + 1; continue; }   // un tramo de filas quitadas corta la racha
        // una evaluación por intervalo entre cortes, no por rango del sqref: miles de rangos con una fórmula larga eran minutos
        let ev = porIntervalo.get(entre);
        if (!ev) porIntervalo.set(entre, ev = evaluar(x));
        if (!actual || actual.racha !== ev.racha) actual = ev;
        // las celdas de la racha copian relativa la fórmula de su primera celda: todos sus tramos tienen la misma fórmula relativa (la clave se calcula
        // una vez por evaluación: una regla de miles de rangos sin referencias a filas quitadas no la recalcula por rango)
        if (actual.clave === undefined) actual.clave = claveDe(actual.fx, actual.fila);
        let g = grupos.get(actual.clave);
        if (!g) grupos.set(actual.clave, g = { formulas: actual.fx, fila: actual.fila, tramos: [] });
        const u = g.tramos[g.tramos.length - 1];
        if (u && u.c0 === r.c0 && u.c1 === r.c1 && u.f1 === s - 1) u.f1 = e; else g.tramos.push({ f0: s, f1: e, c0: r.c0, c1: r.c1 });
        s = e + 1;
      }
    }
    const salida = [];
    for (const g of grupos.values()) {
      let gf = Infinity, gc = Infinity;
      const sq = g.tramos.map((t) => {
        const nf = P.rango(t.f0, t.f1);
        gf = Math.min(gf, nf[0]); gc = Math.min(gc, t.c0);
        const a = numeroACol(t.c0) + (nf[0] + 1), b = numeroACol(t.c1) + (nf[1] + 1);
        return a === b ? a : a + ':' + b;
      });
      salida.push({ sqref: sq.join(' '), formulas: g.formulas.map((f) => LibroXlsx.desplazarFormula(f, gf - g.fila, gc - ac)) });
    }
    return salida;
  }

  // ---- fórmulas compartidas → una por celda -----------------------------------------------------------------------
  const RE_CELDA = /<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g;
  function descompartir(texto) {
    if (!/\bt\s*=\s*["']shared["']/.test(texto)) return texto;
    const maestras = new Map();
    const refCelda = (attrs) => { const m = /\br\s*=\s*"([A-Za-z]{1,3})(\d+)"/.exec(attrs); if (!m) throw new ErrorEscritor('una celda con fórmula compartida no trae su dirección'); return { f: parseInt(m[2], 10) - 1, c: colANumero(m[1].toUpperCase()) }; };
    const RE_F = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/;
    for (const m of texto.matchAll(RE_CELDA)) {
      if (!m[3] || m[3].indexOf('<f') < 0) continue;
      const f = RE_F.exec(m[3]);
      if (!f || !/\bt\s*=\s*["']shared["']/.test(f[1]) || f[2] === undefined || f[2] === '') continue;
      const si = /\bsi\s*=\s*"(\d+)"/.exec(f[1]);
      if (si && !maestras.has(si[1])) maestras.set(si[1], Object.assign({ texto: X.decodificar(f[2]) }, refCelda(m[1])));
    }
    return texto.replace(RE_CELDA, (todo, attrs, resto, interior) => {
      if (!interior || interior.indexOf('<f') < 0) return todo;
      const f = RE_F.exec(interior);
      if (!f || !/\bt\s*=\s*["']shared["']/.test(f[1])) return todo;
      const si = /\bsi\s*=\s*"(\d+)"/.exec(f[1]);
      const mae = si && maestras.get(si[1]);
      if (!mae) throw new ErrorEscritor('una fórmula compartida sin su celda maestra');
      const pos = refCelda(attrs);
      const formula = pos.f === mae.f && pos.c === mae.c ? mae.texto : LibroXlsx.desplazarFormula(mae.texto, pos.f - mae.f, pos.c - mae.c);
      const atributos = f[1].replace(/\s+(?:t|si|ref)\s*=\s*"[^"]*"/g, '');
      return todo.replace(f[0], `<f${atributos}>${X.esc(formula)}</f>`);
    });
  }

  // ---- la hoja de la que se quitan filas -------------------------------------------------------------------------
  // <f>…</f> y <formula>, <formula1>, <formula2>; sin tomar un <f/> vacío por la apertura de otra fórmula
  const RE_F_CELDA = /(<f\b[^>]*(?<!\/)>)([\s\S]*?)(<\/f>)/g;
  const RE_FORMULA = /(<formula[12]?\b[^>]*(?<!\/)>)([\s\S]*?)(<\/formula[12]?>)/g;
  function procesarFormulasDeTexto(texto, enObjetivo, objetivo, M) {
    // <f>…</f> de celdas y <formula>, <formula1>, <formula2> de formato condicional y validación
    return texto
      .replace(RE_F_CELDA, (t, a, f, c) => a + X.esc(reescribirFormula(X.decodificar(f), enObjetivo, objetivo, M)) + c)
      .replace(RE_FORMULA, (t, a, f, c) => a + X.esc(reescribirFormula(X.decodificar(f), enObjetivo, objetivo, M)) + c)
      // <xm:f> de las extensiones x14: así guarda Excel una lista desplegable o un formato condicional que cita otra hoja (medido, corpus 46)
      .replace(/(<xm:f>)([\s\S]*?)(<\/xm:f>)/g, (t, a, f, c) => a + X.esc(reescribirFormula(X.decodificar(f), enObjetivo, objetivo, M)) + c);
  }

  // Umbrales de escalas de color, barras e iconos (<cfvo val="$B$5"/>): una fórmula en un atributo. Excel la acomoda como las demás (medido, corpus 49:
  // $B$5 → $B$3 al borrar las filas 3 y 4); solo admite referencias absolutas, así que basta la fórmula tal cual.
  const RE_CFVO = /(<cfvo\b[^>]*?\bval\s*=\s*")([^"]*)(")/g;
  const umbrales = (t, enObjetivo, objetivo, M) => t.replace(RE_CFVO, (x, a, v, c) => a + X.escAttr(reescribirFormula(X.decodificar(v), enObjetivo, objetivo, M)) + c);
  // Columnas calculadas y fórmulas de totales de una tabla (<calculatedColumnFormula>, <totalsRowFormula>): Excel las acomoda igual que las celdas
  // (medido, corpus 49: [@Monto]*$B$8 → [@Monto]*$B$6, también Datos!$B$8 desde una tabla de otra hoja).
  const RE_F_TABLA = /(<(calculatedColumnFormula|totalsRowFormula)\b[^>]*(?<!\/)>)([\s\S]*?)(<\/\2>)/g;
  const formulasDeTabla = (t, enObjetivo, objetivo, M) => t.replace(RE_F_TABLA, (x, a, n, f, c) => a + X.esc(reescribirFormula(X.decodificar(f), enObjetivo, objetivo, M)) + c);

  // Otra hoja que cita a la editada: sus fórmulas de celda, validación y minigráficos se acomodan como una sola fórmula; su formato condicional (clásico
  // y x14) que cita a la hoja editada se parte como lo hace Excel (partirPorFilas con enObjetivo = false) y, si se partió, se renumeran las prioridades.
  function otraHoja(u, objetivo, M, menciona) {
    const guardados = [], cuenta = { n: 0 };
    let partida = false;
    const guardar = (s) => { guardados.push(s); return '\u0000' + (guardados.length - 1) + '\u0000'; };
    const partir = (sq, formulas) => { const g = partirPorFilas(sq, formulas, M, objetivo, cuenta, false); if (g.length > 1) partida = true; return g; };
    u = u.replace(/<conditionalFormatting\b([^>]*?)\bsqref\s*=\s*"([^"]*)"([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g, (todo, a, sq, b, hijos) => {
      if (!menciona(hijos)) return todo;
      const g = partir(sq, [...hijos.matchAll(RE_FORMULA)].map((m) => X.decodificar(m[2])));
      return guardar(g.map((p) => { let k = 0; return `<conditionalFormatting${a}sqref="${p.sqref}"${b}>${umbrales(hijos.replace(RE_FORMULA, (x, ab, f, c) => ab + X.esc(p.formulas[k++]) + c), false, objetivo, M)}</conditionalFormatting>`; }).join(''));
    });
    u = u.replace(/<x14:conditionalFormatting\b([^>]*)>([\s\S]*?)<\/x14:conditionalFormatting>/g, (todo, a, hijos) => {
      const sq = /<xm:sqref>([^<]*)<\/xm:sqref>/.exec(hijos);
      if (!sq || !menciona(hijos)) return todo;
      const g = partir(X.decodificar(sq[1]), [...hijos.matchAll(/<xm:f>([\s\S]*?)<\/xm:f>/g)].map((m) => X.decodificar(m[1])));
      return guardar(g.map((p) => { let k = 0; return `<x14:conditionalFormatting${a}>${hijos.replace(/(<xm:f>)[\s\S]*?(<\/xm:f>)/g, (x, ab, c) => ab + X.esc(p.formulas[k++]) + c).replace(/(<xm:sqref>)[^<]*(<\/xm:sqref>)/, (x, ab, c) => ab + p.sqref + c)}</x14:conditionalFormatting>`; }).join(''));
    });
    u = procesarFormulasDeTexto(u, false, objetivo, M).replace(/\u0000(\d+)\u0000/g, (x, k) => guardados[+k]);
    if (partida) {
      const RE_P = /(<(?:x14:)?cfRule\b[^>]*?\bpriority\s*=\s*")(\d+)(")/g;
      const ps = [...u.matchAll(RE_P)].map((m, i) => [parseInt(m[2], 10), i]).sort((x, y) => x[0] - y[0] || x[1] - y[1]);
      const nueva = []; ps.forEach(([, i], k) => { nueva[i] = k + 1; });
      let i = 0;
      u = u.replace(RE_P, (x, a, p, c) => a + nueva[i++] + c);
    }
    return u;
  }

  function filasDeLaHoja(xml, M, objetivo) {
    const RE_FILA = /<row\b([^>]*?)(\/>|>([\s\S]*?)<\/row>)/g;
    let n = 0, quitadas = 0;
    xml = xml.replace(RE_FILA, (todo, attrs) => {
      const rm = /\br\s*=\s*"(\d+)"/.exec(attrs);
      n = rm ? parseInt(rm[1], 10) : n + 1;
      const nueva = M.fila(n - 1);
      if (nueva < 0) { quitadas++; return ''; }
      let t = todo;
      if (nueva !== n - 1 || !rm) {
        t = t.replace(/^<row\b[^>]*?(\/?>)/, (ap) => X.ponerAtributo(ap.replace(/\s*\/?>$/, '') + (/\/>$/.test(ap) ? '/>' : '>'), 'r', String(nueva + 1)));
        t = t.replace(/(<c\b[^>]*?\br\s*=\s*")([A-Za-z]{1,3})\d+(")/g, (x, a, col, c) => a + col + (nueva + 1) + c);
      }
      return t;
    });
    return { xml, quitadas };
  }

  // Todo lo que está después de sheetData y apunta a celdas de la hoja.
  function colaDeLaHoja(cola, M, objetivo) {
    let t = cola;
    // celdas combinadas: se achican; una que queda en una sola celda se descarta
    t = t.replace(/<mergeCell\b[^>]*?\bref\s*=\s*"([^"]+)"[^>]*?\/>/g, (todo, ref) => {
      const nuevo = reescribirSqref(ref, M);
      if (!nuevo || !nuevo.includes(':')) return '';
      return todo.replace(ref, nuevo);
    });
    t = t.replace(/<mergeCells\b([^>]*)>([\s\S]*?)<\/mergeCells>/g, (todo, attrs, hijos) => {
      const k = (hijos.match(/<mergeCell\b/g) || []).length;
      return k ? `<mergeCells${attrs.replace(/\bcount\s*=\s*"\d+"/, `count="${k}"`)}>${hijos}</mergeCells>` : '';
    });
    t = t.replace(/<mergeCells\b[^>]*\/>/g, '');
    // formato condicional y validación de datos: una regla puede partirse en varias (partirPorFilas)
    const cuenta = { n: 0 };
    let partida = false;
    const conFormulas = (hijos, formulas) => { let k = 0; return hijos.replace(RE_FORMULA, (x, a, f, c) => a + X.esc(formulas[k++]) + c); };
    const partir = (sq, hijos) => {
      const g = partirPorFilas(sq, hijos ? [...hijos.matchAll(RE_FORMULA)].map((m) => X.decodificar(m[2])) : [], M, objetivo, cuenta);
      if (g.length > 1) partida = true;
      return g;
    };
    t = t.replace(/<conditionalFormatting\b([^>]*?)\bsqref\s*=\s*"([^"]*)"([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g, (todo, a, sq, b, hijos) =>
      partir(sq, hijos).map((g) => `<conditionalFormatting${a}sqref="${g.sqref}"${b}>${conFormulas(hijos, g.formulas)}</conditionalFormatting>`).join(''));
    t = umbrales(t, true, objetivo, M);
    t = t.replace(/<dataValidation\b([^>]*?)\bsqref\s*=\s*"([^"]*)"([^>]*?)(\/>|>([\s\S]*?)<\/dataValidation>)/g, (todo, a, sq, b, cierre, hijos) =>
      partir(sq, hijos).map((g, k) => {
        const b2 = k ? b.replace(/\s+xr:uid\s*=\s*"[^"]*"/, '') : b, a2 = k ? a.replace(/\s+xr:uid\s*=\s*"[^"]*"/, '') : a;   // el identificador, solo en la primera
        return cierre === '/>' ? `<dataValidation${a2}sqref="${g.sqref}"${b2}/>` : `<dataValidation${a2}sqref="${g.sqref}"${b2}>${conFormulas(hijos, g.formulas)}</dataValidation>`;
      }).join(''));
    // si una regla se partió, las prioridades se vuelven a numerar 1, 2, 3… en el mismo orden (las partes de una regla, seguidas)
    if (partida) {
      const RE_P = /(<cfRule\b[^>]*?\bpriority\s*=\s*")(\d+)(")/g;
      const ps = [...t.matchAll(RE_P)].map((m, i) => [parseInt(m[2], 10), i]).sort((x, y) => x[0] - y[0] || x[1] - y[1]);
      const nueva = []; ps.forEach(([, i], k) => { nueva[i] = k + 1; });
      let i = 0;
      t = t.replace(RE_P, (x, a, p, c) => a + nueva[i++] + c);
    }
    t = t.replace(/<dataValidations\b([^>]*)>([\s\S]*?)<\/dataValidations>/g, (todo, attrs, hijos) => {
      const k = (hijos.match(/<dataValidation\b/g) || []).length;
      return k ? `<dataValidations${attrs.replace(/\bcount\s*=\s*"\d+"/, `count="${k}"`)}>${hijos}</dataValidations>` : '';
    });
    // hipervínculos: la celda (ref) sube con su fila
    t = t.replace(/<hyperlink\b[^>]*?\/>/g, (todo) => {
      const ref = /\bref\s*=\s*"([^"]*)"/.exec(todo);
      if (!ref) return todo;
      const nuevo = reescribirSqref(ref[1], M);
      if (!nuevo) return '';
      const u = todo.replace(ref[0], `ref="${nuevo}"`);
      // el destino interno (location) no se toca: Excel lo deja como está (medido)
      return u;
    });
    t = t.replace(/<hyperlinks\b[^>]*>\s*<\/hyperlinks>/g, '');
    // autofiltro y su orden
    t = t.replace(/<(autoFilter|sortState)\b([^>]*?)\bref\s*=\s*"([^"]*)"([^>]*?)(\/?>)/g, (todo, nombre, a, ref, b, cierre) => {
      const nuevo = reescribirSqref(ref, M);
      return nuevo ? `<${nombre}${a}ref="${nuevo}"${b}${cierre}` : todo;
    });
    t = t.replace(/<sortCondition\b([^>]*?)\bref\s*=\s*"([^"]*)"/g, (todo, a, ref) => { const nuevo = reescribirSqref(ref, M); return nuevo ? `<sortCondition${a}ref="${nuevo}"` : todo; });
    // rangos protegidos y errores ignorados
    t = t.replace(/<(protectedRange|ignoredError)\b([^>]*?)\bsqref\s*=\s*"([^"]*)"/g, (todo, nombre, a, sq) => { const nuevo = reescribirSqref(sq, M); return `<${nombre}${a}sqref="${nuevo || sq}"`; });
    // saltos de página: el salto va después de la fila `id` (desde 1)
    t = t.replace(/<rowBreaks\b([^>]*)>([\s\S]*?)<\/rowBreaks>/g, (todo, attrs, hijos) => {
      let total = 0, manuales = 0, ultimo = 0;
      const nuevos = hijos.replace(/<brk\b([^>]*?)\bid\s*=\s*"(\d+)"([^>]*?)\/>/g, (x, a, id, b) => {
        const nuevo = parseInt(id, 10) - M.hasta(parseInt(id, 10) - 1);   // sigue siendo "después de la fila `nuevo`"; si la fila quitada era la del salto, queda donde quedó el hueco
        if (nuevo <= 0 || nuevo === ultimo) return '';
        ultimo = nuevo; total++; if (!/\bman\s*=\s*"(?:0|false)"/.test(a + b)) manuales++;
        return `<brk${a}id="${nuevo}"${b}/>`;
      });
      return total ? `<rowBreaks${attrs.replace(/\bcount\s*=\s*"\d+"/, `count="${total}"`).replace(/\bmanualBreakCount\s*=\s*"\d+"/, `manualBreakCount="${manuales}"`)}>${nuevos}</rowBreaks>` : '';
    });
    return t;
  }

  // ---- el paquete ------------------------------------------------------------------------------------------------
  async function leerTexto(zip, ruta) { const f = zip.file(ruta); return f ? f.async('string') : null; }

  async function quitarFilas(libro, hoja, filas) {
    const zip = libro.zip;
    const i = typeof hoja === 'number' ? hoja : libro.hojas.findIndex((h) => h.nombre === hoja);
    const entrada = libro.hojas[i];
    if (!entrada) throw new ErrorEscritor('hoja inexistente');
    if (entrada.tipo !== 'hoja' || !entrada.ruta) throw new ErrorEscritor('la hoja no es de celdas');
    if (!Array.isArray(filas)) throw new ErrorEscritor('las filas tienen que ser una lista');
    const lista = [...new Set(filas)].sort((a, b) => a - b);
    for (const f of lista) if (!Number.isInteger(f) || f < 0 || f >= 1048576) throw new ErrorEscritor('fila fuera de los límites de Excel');
    if (!lista.length) return { quitadas: 0 };
    const M = mapeo(lista);
    const objetivo = entrada.nombre;
    const nuevos = new Map();   // ruta → texto (se escribe todo al final, o nada)

    const xmlObjetivo = await leerTexto(zip, entrada.ruta);
    if (xmlObjetivo === null) throw new ErrorEscritor('falta la parte ' + entrada.ruta);

    // 1. lo que no se sabe mover
    const rels = await XlsxPaquete.leerRels(zip, entrada.ruta);
    const prohibidos = rels.find((r) => /^(comments|threadedComment|pivotTable|slicer|timeline|queryTable|ctrlProp|oleObject|vmlDrawing)$/.test(r.tipo) || /^(comments|threadedComments)/.test(r.tipo));
    if (prohibidos) throw new ErrorEscritor('la hoja tiene ' + prohibidos.tipo + ' (notas, tablas dinámicas, segmentaciones o controles): no se pueden mover con las filas');
    if (/<legacyDrawing\b|<controls\b|<oleObjects\b/.test(xmlObjetivo)) throw new ErrorEscritor('la hoja tiene notas, controles u objetos OLE');
    if (/\bt\s*=\s*["'](?:array|dataTable)["']/.test(xmlObjetivo)) throw new ErrorEscritor('la hoja tiene fórmulas matriciales o tablas de datos');
    if (/<x14:(?:conditionalFormatting|dataValidation|sparklineGroup)|<xm:sqref/.test(xmlObjetivo)) throw new ErrorEscritor('la hoja tiene formato condicional, validación o minigráficos de Excel 2010 o posterior');
    // <x:row>, <x:c>…: otros programas escriben la hoja con prefijo; las expresiones de aquí buscan los nombres sin prefijo
    if (PREFIJADOS.test(xmlObjetivo)) throw new ErrorEscritor('la hoja está escrita con prefijos de espacio de nombres');
    // el nombre de la hoja tal como puede aparecer en el XML de otra parte (con '' en las fórmulas y con &amp; &lt; &apos;…)
    const formas = new Set();
    for (const base of [objetivo, objetivo.replace(/'/g, "''")]) { const e = X.esc(base); formas.add(base); formas.add(e); formas.add(e.replace(/'/g, '&apos;')); formas.add(e.replace(/'/g, '&#39;')); formas.add(X.escAttr(base)); }
    const marcas = [...formas].map((s) => s.toLowerCase());
    const menciona = (t) => { const b = t.toLowerCase(); return marcas.some((m) => b.includes(m)); };
    // una tabla dinámica de otra hoja que toma sus datos de esta
    for (const h of libro.hojas) {
      if (h.tipo !== 'hoja' || !h.ruta) continue;
      for (const r of await XlsxPaquete.leerRels(zip, h.ruta)) if (r.tipo === 'pivotTable' && h.ruta !== entrada.ruta) {
        const cache = (await XlsxPaquete.leerRels(zip, r.destino)).find((c) => c.tipo === 'pivotCacheDefinition');
        const t = cache && await leerTexto(zip, cache.destino);
        const fuente = t && /<worksheetSource\b[^>]*\bsheet\s*=\s*"([^"]*)"/.exec(t);
        if (fuente && X.decodificar(fuente[1]).toLowerCase() === objetivo.toLowerCase()) throw new ErrorEscritor('una tabla dinámica toma sus datos de esta hoja');
      }
    }

    // 2. tablas de la hoja
    for (const r of rels.filter((x) => x.tipo === 'table')) {
      const t = await leerTexto(zip, r.destino);
      if (t === null) continue;
      const ref = /<table\b[^>]*?\bref\s*=\s*"([^"]+)"/.exec(t);
      const rg = ref && LibroXlsx.decodificarRango(ref[1]);
      if (!rg) continue;
      const encabezado = !/\bheaderRowCount\s*=\s*"0"/.test(t), totales = /\btotalsRowCount\s*=\s*"[1-9]/.test(t) || /\btotalsRowShown\s*=\s*"1"/.test(t);
      if (encabezado && M.quita(rg.f0, rg.f0)) throw new ErrorEscritor('se quita el encabezado de una tabla');
      if (totales && M.quita(rg.f1, rg.f1)) throw new ErrorEscritor('se quita la fila de totales de una tabla');
      const nf = M.rango(rg.f0, rg.f1);
      if (!nf || nf[1] - nf[0] + 1 <= (encabezado ? 1 : 0) + (totales ? 1 : 0)) throw new ErrorEscritor('se quitan todas las filas de datos de una tabla');
      const cambiar = (s) => formulasDeTabla(s.replace(/(<(?:table|autoFilter)\b[^>]*?\bref\s*=\s*")([^"]+)(")/g, (x, a, rf, c) => a + reescribirSqref(rf, M) + c), true, objetivo, M);
      nuevos.set(r.destino, cambiar(t));
    }

    // 3. la hoja: filas, celdas, fórmulas y todo lo que apunta a ellas
    let xml = descompartir(xmlObjetivo);
    const sd = /<sheetData\b[^>]*?(\/>|>[\s\S]*<\/sheetData>)/.exec(xml);
    if (!sd) throw new ErrorEscritor('la hoja no tiene sheetData');
    const ini = sd.index, fin = ini + sd[0].length;
    let datos = sd[0];
    if (sd[1] !== '/>') {
      const r = filasDeLaHoja(datos, M, objetivo);
      datos = procesarFormulasDeTexto(r.xml, true, objetivo, M);
    }
    let antes = xml.slice(0, ini), cola = xml.slice(fin);
    cola = colaDeLaHoja(cola, M, objetivo);
    antes = antes.replace(/(<dimension\b[^>]*?\bref\s*=\s*")([^"]+)(")/, (x, a, ref, c) => {
      const rg = LibroXlsx.decodificarRango(ref);
      if (!rg) return x;
      const nf = M.rango(rg.f0, rg.f1) || [0, 0];
      return a + numeroACol(rg.c0) + (nf[0] + 1) + ':' + numeroACol(rg.c1) + (nf[1] + 1) + c;
    });
    // paneles inmovilizados: las filas fijas que se quitan dejan de contar (medido, corpus 47: ySplit 3 → 2 y topLeftCell A4 → A3)
    antes = antes.replace(/<pane\b[^>]*\/?>/, (p) => {
      const y = /\bySplit\s*=\s*"(\d+)"/.exec(p);
      if (!y || !/\bstate\s*=\s*"frozen(?:Split)?"/.test(p)) return p;
      const n = parseInt(y[1], 10), ny = n - M.quita(0, n - 1);
      if (ny === n || ny <= 0) return p;   // si se quitan todas las filas fijas, el panel queda como estaba
      return p.replace(y[0], `ySplit="${ny}"`).replace(/(\btopLeftCell\s*=\s*"[A-Za-z]{1,3})(\d+)(")/, (x, a, f, c) => a + (M.fila(M.siguiente(parseInt(f, 10) - 1)) + 1) + c);
    });
    nuevos.set(entrada.ruta, antes + datos + cola);

    // 4. dibujos de la hoja: las anclas suben con sus filas según cómo están colocadas (medido, corpus 47): "mover y cambiar tamaño" (twoCell, lo
    // normal) sube el inicio y el final; "mover sin cambiar tamaño" (editAs="oneCell") sube el inicio y conserva su alto en filas; "no mover"
    // (editAs="absolute") se queda donde está
    const filaDe = (s, cual) => { const m = new RegExp('<(?:\\w+:)?' + cual + '>[\\s\\S]*?<(?:\\w+:)?row>(\\d+)<\\/').exec(s); return m ? parseInt(m[1], 10) : null; };
    const ponerFila = (s, cual, f) => s.replace(new RegExp('(<(?:\\w+:)?' + cual + '>[\\s\\S]*?<(?:\\w+:)?row>)(\\d+)(<\\/)'), (x, a, v, c) => a + f + c);
    for (const r of rels.filter((x) => x.tipo === 'drawing')) {
      const t = await leerTexto(zip, r.destino);
      if (t === null) continue;
      const u = t.replace(/<((?:\w+:)?)(twoCellAnchor|oneCellAnchor)\b([^>]*)>([\s\S]*?)<\/\1\2>/g, (todo, p, tipo, attrs, cuerpo) => {
        const como = tipo === 'oneCellAnchor' ? 'oneCell' : (/\beditAs\s*=\s*"(\w+)"/.exec(attrs) || [])[1] || 'twoCell';
        if (como === 'absolute') return todo;
        const de = filaDe(cuerpo, 'from'), a = filaDe(cuerpo, 'to');
        if (de === null) return todo;
        const nde = de - M.antes(de);
        let c = ponerFila(cuerpo, 'from', nde);
        if (a !== null) c = ponerFila(c, 'to', como === 'oneCell' ? nde + Math.max(0, a - de) : Math.max(0, a - M.hasta(a)));
        return `<${p}${tipo}${attrs}>${c}</${p}${tipo}>`;
      });
      nuevos.set(r.destino, u);
    }

    // 5. otras hojas, nombres definidos y gráficos: solo las referencias que llevan el nombre de esta hoja
    const rutaLibro = await XlsxPaquete.rutaLibroDe(zip);
    const wb = await leerTexto(zip, rutaLibro);
    let partes = 0;
    for (const h of libro.hojas) {
      if (h.tipo !== 'hoja' || !h.ruta || h.ruta === entrada.ruta) continue;
      // sus tablas: la columna calculada o el total que citan a esta hoja
      for (const r of (await XlsxPaquete.leerRels(zip, h.ruta)).filter((x) => x.tipo === 'table')) {
        const tt = await leerTexto(zip, r.destino);
        if (tt === null || !menciona(tt)) continue;
        if (++partes > MAX_PARTES) throw new ErrorEscritor('demasiadas partes');
        const ut = formulasDeTabla(tt, false, objetivo, M);
        if (ut !== tt) nuevos.set(r.destino, ut);
      }
      const t = await leerTexto(zip, h.ruta);
      if (t === null || !menciona(t)) continue;
      if (++partes > MAX_PARTES) throw new ErrorEscritor('demasiadas partes');
      if (PREFIJADOS.test(t)) throw new ErrorEscritor('una hoja que cita a esta está escrita con prefijos de espacio de nombres');
      nuevos.set(h.ruta, otraHoja(descompartir(t), objetivo, M, menciona));
    }
    if (wb !== null) {
      const u = wb.replace(/(<definedName\b[^>]*>)([\s\S]*?)(<\/definedName>)/g, (x, a, f, c) => a + X.esc(reescribirFormula(X.decodificar(f), false, objetivo, M)) + c);
      if (u !== wb) nuevos.set(rutaLibro, u);
    }
    const graficos = [];
    zip.forEach((nombre) => { if (/^xl\/charts\/[^/]+\.xml$/i.test(nombre)) graficos.push(nombre); });
    for (const nombre of graficos) {
      const t = await leerTexto(zip, nombre);
      if (t === null || !menciona(t)) continue;
      nuevos.set(nombre, t.replace(/(<(?:\w+:)?f>)([\s\S]*?)(<\/(?:\w+:)?f>)/g, (x, a, f, c) => a + X.esc(reescribirFormula(X.decodificar(f), false, objetivo, M)) + c));
    }

    // 6. todo bien: se escribe
    for (const [ruta, t] of nuevos) zip.file(ruta, t);
    if (libro._hojasCargadas) libro._hojasCargadas.clear();   // los modelos leídos antes ya no valen
    await XlsxPaquete.quitarCalcChain(zip);
    await XlsxPaquete.marcarRecalculo(zip);
    return { quitadas: lista.length };
  }

  XlsxEscritor.quitarFilas = quitarFilas;
  XlsxEscritor._quitarFilasInterno = { reescribirFormula, reescribirSqref, mapeo, partirPorFilas };
})();

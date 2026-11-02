/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// COPIA ADAPTADA para la suite (Office-Temporal, 2026-10-25) de codigo/MiniCookies-Home/assets/js/xlsx-quitar-filas.js. Cambio: además de quitar,
// INSERTA filas con el mismo recorrido (fórmulas de todas las hojas, combinadas, formato condicional, validación, tablas, nombres, gráficos, anclas,
// notas): solo cambia el "mapeo" de filas (mapeoInsercion). Excel expande un rango si se inserta dentro de él y lo corre si se inserta encima.
//   await XlsxEscritor.insertarFilas(libro, hoja, fila, cuantas)
// Ver docs/motores-adaptados.md.
// Quitar filas de una hoja EN EL LUGAR (fase C del motor de Excel), como "Eliminar filas" de Excel: las filas de abajo suben, y todo lo que apunta a
// la hoja se acomoda: fórmulas de todas las hojas (una referencia a una fila quitada pasa a #REF!, un rango se achica), celdas combinadas, formato
// condicional, validación de datos, hipervínculos, autofiltro, tablas, nombres definidos, gráficos, saltos de página y anclas de los dibujos.
// El resto del paquete (estilos, imágenes, macros, otras hojas) queda tal cual. No depende del navegador.
//
//   const r = await XlsxEscritor.quitarFilas(libro, hoja, [filas desde 0]);   // libro: LibroXlsx.abrir; hoja: índice o nombre
//   // r = { quitadas: n }  · después se guarda con libro.zip.generateAsync(...)
// También mueve: notas (comentarios clásicos con su dibujo VML) y comentarios en hilo (una nota o un hilo en una fila quitada desaparece con ella,
// con su forma del VML y su entrada del hilo; las demás suben con su fila), fórmulas matriciales y tablas de datos (las que quedan arriba o abajo
// de las filas quitadas suben con su rango; las que quedan enteras dentro de ellas desaparecen) y las extensiones x14 de formato condicional,
// validación de datos y minigráficos (un minigráfico en una fila quitada desaparece; si lo que pierde son sus datos, queda sin ellos, como en Excel).
// Rechaza (ErrorEscritor, sin tocar nada) lo que no sabe mover sin dejar un archivo que Excel repare: una fila quitada que parte una fórmula
// matricial o una tabla de datos (Excel tampoco deja borrar parte de una) o que quita la celda de entrada de una tabla de datos, tablas dinámicas
// y segmentaciones de la hoja o con ella de origen, controles y objetos OLE (también las formas del dibujo de notas que no son notas), consultas,
// el encabezado o la fila de totales de una tabla, una tabla que quedaría sin filas, hojas escritas con prefijo de espacio de nombres (<x:row>),
// extensiones x14 con otro prefijo y formato condicional con fórmulas relativas que obligaría a partirlo en demasiadas reglas.
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
  // las extensiones x14 las escribe Excel siempre con estos prefijos; con otro, las expresiones de aquí no las verían
  const NS_X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main', NS_XM = 'http://schemas.microsoft.com/office/excel/2006/main';

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

  // Insertar `n` filas antes de la fila `r` (desde 0): la fila f pasa a f + n si f >= r. Misma interfaz que mapeo().
  function mapeoInsercion(r, n) {
    const fila = (f) => (f < r ? f : f + n);
    const corrido = (f) => (f < r ? 0 : -n);   // "filas quitadas encima": negativo, la fila baja
    return {
      lista: [], insercion: { r, n },
      siguiente: (f) => f,
      fila,
      rango: (f0, f1) => { const a = fila(f0), b = fila(f1); return b > 1048575 ? (a > 1048575 ? null : [a, 1048575]) : [a, b]; },
      quita: () => 0,
      antes: corrido,
      hasta: corrido,
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

  // Un minigráfico (<x14:sparkline>: <xm:f> con sus datos y <xm:sqref> con su celda). enObjetivo: el minigráfico está en la hoja editada, así que su
  // celda sube con su fila y, si esa fila se quita, desaparece (→ ''); en otra hoja solo se acomoda su rango de datos. Si lo que se quita son todos sus
  // datos, el minigráfico queda sin <xm:f> (medido: es lo que guarda Excel; una referencia #REF! en <xm:f> no la sabe leer).
  const RE_MINI = /<x14:sparkline\s*>([\s\S]*?)<\/x14:sparkline>/g;
  function minigrafico(hijos, enObjetivo, objetivo, M) {
    let t = hijos;
    if (enObjetivo) {
      const sq = /<xm:sqref>([^<]*)<\/xm:sqref>/.exec(t);
      if (!sq) throw new ErrorEscritor('un minigráfico sin su celda');
      const nuevo = reescribirSqref(X.decodificar(sq[1]), M);
      if (!nuevo) return '';
      t = t.replace(sq[0], () => '<xm:sqref>' + nuevo + '</xm:sqref>');
    }
    t = t.replace(/<xm:f>([\s\S]*?)<\/xm:f>/, (x, f) => {
      const viejo = X.decodificar(f), n = reescribirFormula(viejo, enObjetivo, objetivo, M);
      return /#REF!/.test(n) && !/#REF!/.test(viejo) ? '' : '<xm:f>' + X.esc(n) + '</xm:f>';
    });
    return '<x14:sparkline>' + t + '</x14:sparkline>';
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
    u = u.replace(RE_MINI, (todo, hijos) => guardar(minigrafico(hijos, false, objetivo, M)));
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

  // Fórmulas matriciales (<f t="array" ref="F10:F12">) y tablas de datos (<f t="dataTable" ref="Q28:Q30" r1="P1">). Excel no deja borrar parte de una: o las
  // filas quitadas no la tocan (su rango sube con ellas) o la contienen entera (desaparece con sus celdas); si no, se rechaza. De una tabla de datos
  // cuenta también la fila de arriba de su rango (donde están los valores o la fórmula) y sus celdas de entrada r1 y r2 suben con su fila (si se quita
  // la fila de una, Excel guarda la tabla con del1/del2 y aquí se rechaza).
  const RE_F_ESPECIAL = /<f\b([^>]*?\bt\s*=\s*["'](?:array|dataTable)["'][^>]*?)(\/?)>/g;
  const dirCelda = (s) => {
    const m = /^\s*\$?([A-Za-z]{1,3})\$?(\d+)\s*$/.exec(String(s)), f = m ? parseInt(m[2], 10) - 1 : -1;
    return m && f >= 0 && f < 1048576 && colANumero(m[1].toUpperCase()) <= 16383 ? { col: m[1], f } : null;
  };
  function ajustarMatrices(fila, M) {
    return fila.replace(RE_F_ESPECIAL, (todo, attrs, cierre) => {
      const esTabla = /\bt\s*=\s*["']dataTable["']/.test(attrs);
      const mr = /(\bref\s*=\s*")([^"]*)(")/.exec(attrs), rg = mr && LibroXlsx.decodificarRango(mr[2]);
      if (!rg || rg.f1 >= 1048576) throw new ErrorEscritor('una fórmula matricial o tabla de datos sin su rango');
      const alto = rg.f1 - rg.f0 + 1, q = M.quita(rg.f0, rg.f1);
      if (q === alto) return todo;
      if (q) throw new ErrorEscritor('se quita parte de una fórmula matricial o de una tabla de datos');
      if (esTabla && rg.f0 > 0 && M.fila(rg.f0 - 1) < 0) throw new ErrorEscritor('se quita parte de una tabla de datos');
      let a = attrs;
      const d = M.antes(rg.f0);
      if (d) {
        const solo = rg.f0 === rg.f1 && rg.c0 === rg.c1;
        const nuevo = numeroACol(rg.c0) + (rg.f0 - d + 1) + (solo ? '' : ':' + numeroACol(rg.c1) + (rg.f1 - d + 1));
        a = a.replace(mr[0], () => mr[1] + nuevo + mr[3]);
      }
      if (esTabla) {
        for (const nombre of ['r1', 'r2']) {
          const mi = new RegExp('(\\b' + nombre + '\\s*=\\s*")([^"]*)(")').exec(a);
          if (!mi) continue;
          const c = dirCelda(mi[2]);
          if (!c) throw new ErrorEscritor('una tabla de datos con una celda de entrada que no se entiende');
          const nf = M.fila(c.f);
          if (nf < 0) throw new ErrorEscritor('se quita la celda de entrada de una tabla de datos');
          if (nf !== c.f) a = a.replace(mi[0], () => mi[1] + mi[2].replace(/\d+$/, String(nf + 1)) + mi[3]);
        }
      }
      return '<f' + a + cierre + '>';
    });
  }

  function filasDeLaHoja(xml, M, objetivo) {
    const RE_FILA = /<row\b([^>]*?)(\/>|>([\s\S]*?)<\/row>)/g;
    let n = 0, quitadas = 0;
    const altos = [];   // alto de las filas quitadas que lo traen escrito (para mover las formas de las notas)
    xml = xml.replace(RE_FILA, (todo, attrs) => {
      const rm = /\br\s*=\s*"(\d+)"/.exec(attrs);
      n = rm ? parseInt(rm[1], 10) : n + 1;
      const nueva = M.fila(n - 1);
      if (todo.indexOf('array') > 0 || todo.indexOf('dataTable') > 0) todo = ajustarMatrices(todo, M);
      if (nueva < 0) {
        quitadas++;
        const ht = /\bht\s*=\s*"([\d.]+)"/.exec(attrs);
        if (/\bhidden\s*=\s*"(?:1|true)"/.test(attrs)) altos.push({ f: n - 1, h: 0 }); else if (ht) altos.push({ f: n - 1, h: Math.round(parseFloat(ht[1]) * 4 / 3) * 0.75 });   // Excel cuenta la fila en píxeles: 39,95 pt son 53 px = 39,75 pt (medido)
        return '';
      }
      let t = todo;
      if (nueva !== n - 1 || !rm) {
        t = t.replace(/^<row\b[^>]*?(\/?>)/, (ap) => X.ponerAtributo(ap.replace(/\s*\/?>$/, '') + (/\/>$/.test(ap) ? '/>' : '>'), 'r', String(nueva + 1)));
        t = t.replace(/(<c\b[^>]*?\br\s*=\s*")([A-Za-z]{1,3})\d+(")/g, (x, a, col, c) => a + col + (nueva + 1) + c);
      }
      return t;
    });
    return { xml, quitadas, altos };
  }

  // Extensiones de Excel 2010 de la hoja editada (<extLst> al final): formato condicional (<x14:conditionalFormatting>, regla y <xm:sqref>), validación
  // de datos (<x14:dataValidation>) y minigráficos. Se acomodan como sus equivalentes clásicos (partirPorFilas); lo que queda vacío se quita, también la
  // <ext> y el <extLst> que lo contenían (medido, corpus 52: Excel hace lo mismo). Una regla partida recibe identificadores nuevos (id de la regla).
  const RE_XMF = /(<xm:f>)([\s\S]*?)(<\/xm:f>)/g, RE_XMSQ = /(<xm:sqref>)([^<]*)(<\/xm:sqref>)/;
  const idDerivado = (id, k) => {
    const h = id.replace(/-/g, '').toUpperCase(), n = ((parseInt(h.slice(-8), 16) + k * 2654435761) >>> 0).toString(16).toUpperCase().padStart(8, '0'), t = h.slice(0, -8) + n;
    return t.slice(0, 8) + '-' + t.slice(8, 12) + '-' + t.slice(12, 16) + '-' + t.slice(16, 20) + '-' + t.slice(20);
  };
  function x14Objetivo(t, M, objetivo, cuenta, marcarPartida) {
    if (t.indexOf('x14:') < 0) return t;
    for (const n of ['x14:conditionalFormattings', 'x14:conditionalFormatting', 'x14:dataValidations', 'x14:dataValidation', 'x14:sparklineGroups', 'x14:sparklineGroup', 'x14:sparklines', 'x14:sparkline']) equilibrado(t, n, 'las extensiones de Excel 2010');
    const guardados = [];
    const guardar = (s) => { if (!s) return ''; guardados.push(s); return '\u0000' + (guardados.length - 1) + '\u0000'; };
    const formulas = (h) => [...h.matchAll(RE_XMF)].map((m) => X.decodificar(m[2]));
    const conFormulas = (h, fx) => { let k = 0; return h.replace(RE_XMF, (x, a, f, c) => a + X.esc(fx[k++]) + c); };
    const partes = (h, que) => {
      const sq = RE_XMSQ.exec(h);
      if (!sq) throw new ErrorEscritor('una extensión de Excel 2010 (' + que + ') sin su rango');
      const g = partirPorFilas(X.decodificar(sq[2]), formulas(h), M, objetivo, cuenta);
      if (g.length > 1) marcarPartida();
      return g;
    };
    const conRango = (h, p) => conFormulas(h, p.formulas).replace(RE_XMSQ, (x, a, s, c) => a + p.sqref + c);
    t = t.replace(/<x14:conditionalFormattings\b([^>]*)>([\s\S]*?)<\/x14:conditionalFormattings>/g, (todo, a, hijos) => {
      let n = 0;
      const u = hijos.replace(/<x14:conditionalFormatting\b([^>]*)>([\s\S]*?)<\/x14:conditionalFormatting>/g, (x, b, h) => partes(h, 'formato condicional').map((p, k) => {
        n++;
        const cuerpo = conRango(h, p);
        return '<x14:conditionalFormatting' + b + '>' + (k ? cuerpo.replace(/(<x14:cfRule\b[^>]*?\bid\s*=\s*"\{)([0-9A-Fa-f-]{36})(\})/, (y, c1, id, c3) => c1 + idDerivado(id, k) + c3) : cuerpo) + '</x14:conditionalFormatting>';
      }).join(''));
      return n ? guardar('<x14:conditionalFormattings' + a + '>' + u + '</x14:conditionalFormattings>') : '';
    });
    t = t.replace(/<x14:dataValidations\b([^>]*)>([\s\S]*?)<\/x14:dataValidations>/g, (todo, a, hijos) => {
      let n = 0;
      const u = hijos.replace(/<x14:dataValidation\b([^>]*)>([\s\S]*?)<\/x14:dataValidation>/g, (x, b, h) => partes(h, 'validación').map((p, k) => {
        n++;
        return '<x14:dataValidation' + (k ? b.replace(/\s+xr:uid\s*=\s*"[^"]*"/, '') : b) + '>' + conRango(h, p) + '</x14:dataValidation>';
      }).join(''));
      return n ? guardar('<x14:dataValidations' + a.replace(/\bcount\s*=\s*"\d+"/, 'count="' + n + '"') + '>' + u + '</x14:dataValidations>') : '';
    });
    t = t.replace(/<x14:sparklineGroups\b([^>]*)>([\s\S]*?)<\/x14:sparklineGroups>/g, (todo, a, hijos) => {
      let grupos = 0;
      const u = hijos.replace(/<x14:sparklineGroup\b([^>]*)>([\s\S]*?)<\/x14:sparklineGroup>/g, (x, b, h) => {
        let lista = '';
        const molde = h.replace(/<x14:sparklines>[\s\S]*?<\/x14:sparklines>/, (y) => { lista = y; return '\u0001'; });
        // el eje de fechas del grupo (un <xm:f> fuera de la lista) también cita celdas
        const fuera = molde.replace(RE_XMF, (y, c1, f, c3) => {
          const viejo = X.decodificar(f), n = reescribirFormula(viejo, true, objetivo, M);
          if (/#REF!/.test(n) && !/#REF!/.test(viejo)) throw new ErrorEscritor('el eje de fechas de un minigráfico perdería sus datos');
          return c1 + X.esc(n) + c3;
        });
        let k = 0;
        const dentro = lista.replace(RE_MINI, (y, hs) => { const r = minigrafico(hs, true, objetivo, M); if (r) k++; return r; });
        if (!k) return '';
        grupos++;
        return '<x14:sparklineGroup' + b + '>' + fuera.replace('\u0001', () => dentro) + '</x14:sparklineGroup>';
      });
      return grupos ? guardar('<x14:sparklineGroups' + a + '>' + u + '</x14:sparklineGroups>') : '';
    });
    // lo que quede con <xm:…> es de una extensión que no se sabe acomodar
    if (/<xm:/.test(t)) throw new ErrorEscritor('la hoja tiene una extensión de Excel 2010 o posterior que no se sabe acomodar');
    t = t.replace(/\u0000(\d+)\u0000/g, (x, i) => guardados[+i]);
    return t.replace(/<ext\b[^>]*>\s*<\/ext>/g, '').replace(/<extLst>\s*<\/extLst>/g, '');
  }

  // ---- notas y comentarios en hilo --------------------------------------------------------------------------------
  // Tres partes ligadas por la posición de la celda: xl/comments*.xml (<comment ref>), el dibujo VML de la hoja (una <v:shape> por nota, con su fila en
  // <x:Row> y la posición del cuadro en <x:Anchor>) y, en los comentarios en hilo, xl/threadedComments/*.xml (<threadedComment ref>, con sus respuestas
  // en la misma celda); cada hilo deja además una nota "[Comentario encadenado]" en comments*.xml. Una celda quitada se lleva las tres. Medido (corpus
  // 51-52): el cuadro de una nota que queda sube entero, con su tamaño, tantas filas como se quitaron por encima de SU CELDA (aunque el cuadro esté
  // arrastrado lejos de ella o cruce filas quitadas; si no cabe más arriba, solo sube hasta la fila 1); Excel además pone al día margin-top y deja solo los
  // autores que se siguen usando, en el orden en que aparecen por primera vez entre las notas.
  // Cada apertura con su cierre: con etiquetas sin cerrar las expresiones de abajo recorrerían todo el texto una vez por apertura (cuadrático)
  function equilibrado(t, nombre, que, raiz) {
    if (raiz && !t.trimEnd().endsWith('</' + raiz + '>')) throw new ErrorEscritor('hay XML mal formado en ' + que);   // cortado antes de su cierre
    let a = 0;
    for (const x of t.match(new RegExp('<' + nombre + '\\b[^>]*>', 'g')) || []) if (x[x.length - 2] !== '/') a++;
    if (a !== (t.match(new RegExp('</' + nombre + '>', 'g')) || []).length) throw new ErrorEscritor('hay XML mal formado en ' + que);
  }
  const reRef = (attrs) => { const mr = /(\bref\s*=\s*")([^"]*)(")/.exec(attrs), c = mr && dirCelda(mr[2]); if (!c) throw new ErrorEscritor('una nota o comentario sin su celda'); return { mr, c }; };
  function comentariosClasicos(texto, M) {
    equilibrado(texto, 'comment', 'las notas', 'comments');
    let quitados = 0, quedan = 0;
    const usados = [];   // autores en el orden en que se usan por primera vez
    let t = texto.replace(/<comment\b([^>]*?)(?:\/>|>[\s\S]*?<\/comment>)/g, (todo, attrs) => {
      const { mr, c } = reRef(attrs), nf = M.fila(c.f);
      if (nf < 0) { quitados++; return ''; }
      quedan++;
      const au = /\bauthorId\s*=\s*"(\d+)"/.exec(attrs);
      if (au && !usados.includes(+au[1])) usados.push(+au[1]);
      return nf === c.f ? todo : todo.replace(mr[0], () => mr[1] + c.col + (nf + 1) + mr[3]);
    });
    if (quitados) {
      const nuevo = new Map(), vivos = [];
      t = t.replace(/<authors>([\s\S]*?)<\/authors>/, (todo, h) => {
        const lista = h.match(/<author\b[^>]*?(?:\/>|>[\s\S]*?<\/author>)/g) || [];
        for (const i of usados) if (lista[i] !== undefined) { nuevo.set(i, vivos.length); vivos.push(lista[i]); }
        return vivos.length ? '<authors>' + vivos.join('') + '</authors>' : todo;
      });
      if (nuevo.size) t = t.replace(/(<comment\b[^>]*?\bauthorId\s*=\s*")(\d+)(")/g, (x, a, i, c) => a + (nuevo.has(+i) ? nuevo.get(+i) : i) + c);
    }
    return { texto: t, quitados, quedan };
  }
  function hilosDeComentarios(texto, M) {
    equilibrado(texto, 'threadedComment', 'los comentarios en hilo', 'ThreadedComments');
    let quitados = 0, quedan = 0;
    const t = texto.replace(/<threadedComment\b([^>]*?)(?:\/>|>[\s\S]*?<\/threadedComment>)/g, (todo, attrs) => {
      const { mr, c } = reRef(attrs), nf = M.fila(c.f);
      if (nf < 0) { quitados++; return ''; }
      quedan++;
      return nf === c.f ? todo : todo.replace(mr[0], () => mr[1] + c.col + (nf + 1) + mr[3]);
    });
    return { texto: t, quitados, quedan };
  }
  // quitadaAntes(f): puntos de alto de las filas quitadas por encima de la fila f (desde 0)
  function dibujoDeNotas(texto, M, quitadaAntes) {
    equilibrado(texto, '(?:\\w+:)?shape', 'el dibujo de las notas', 'xml');
    let quitados = 0, notas = 0, otras = 0;
    const t = texto.replace(/<((?:\w+:)?)shape\b[^>]*>[\s\S]*?<\/\1shape>/g, (todo) => {
      const ob = /<(?:\w+:)?ClientData\b[^>]*?\bObjectType\s*=\s*["']([^"']*)["']/.exec(todo);
      if (!ob) { otras++; return todo; }
      if (ob[1] !== 'Note') throw new ErrorEscritor('el dibujo de la hoja tiene controles u objetos que no son notas');
      const mf = /<(?:\w+:)?Row>\s*(\d+)\s*<\/(?:\w+:)?Row>/.exec(todo);
      if (!mf) throw new ErrorEscritor('una nota del dibujo sin su fila');
      const f = parseInt(mf[1], 10), nf = M.fila(f);
      if (nf < 0) { quitados++; return ''; }
      notas++;
      const d = M.antes(f);
      if (!d) return todo;
      let u = todo.replace(mf[0], (x) => x.replace(/\d+/, String(nf)));
      u = u.replace(/(<(?:\w+:)?Anchor>)([^<]*)(<\/)/, (x, a, v, c) => {
        const p = v.split(',');
        if (p.length !== 8) return x;
        const sube = Math.min(d, parseInt(p[2], 10) || 0);   // el cuadro no sale por arriba de la hoja: conserva su alto
        for (const i of [2, 6]) p[i] = p[i].replace(/\d+/, (nn) => String(Math.max(0, parseInt(nn, 10) - sube)));
        return a + p.join(',') + c;
      });
      const pt = quitadaAntes(f);
      if (pt) u = u.replace(/(margin-top:\s*)(-?[\d.]+)(pt)?/, (x, a, v) => { const n = Math.max(0, Math.round((parseFloat(v) - pt) * 100) / 100); return a + (n ? n + 'pt' : '0'); });   // Excel escribe "0" sin unidad
      return u;
    });
    return { texto: t, quitados, quedan: notas, otras };
  }
  function alturaQuitada(M, altos, defecto) {
    altos.sort((a, b) => a.f - b.f);
    const pref = [0];
    for (const a of altos) pref.push(pref[pref.length - 1] + a.h - defecto);
    return (f) => { let lo = 0, hi = altos.length; while (lo < hi) { const m = (lo + hi) >> 1; if (altos[m].f < f) lo = m + 1; else hi = m; } return M.antes(f) * defecto + pref[lo]; };
  }
  // → { nuevos: Map ruta→texto, borrar: [rutas], rels: Set de ids de la hoja que se quitan, legacy: se quita <legacyDrawing> de la hoja }
  async function notasDeLaHoja(zip, rels, xmlHoja, M, quitadaAntes) {
    const plan = { nuevos: new Map(), borrar: [], rels: new Set(), legacy: false };
    const comentarios = rels.filter((r) => r.tipo === 'comments'), hilos = rels.filter((r) => r.tipo === 'threadedComment');
    const ld = /<legacyDrawing\b[^>]*?\bid\s*=\s*"([^"]*)"/.exec(xmlHoja);
    if (!comentarios.length && !hilos.length && !ld) return plan;
    if (comentarios.length > 1 || hilos.length > 1) throw new ErrorEscritor('la hoja tiene más de una parte de notas');
    const vml = ld && rels.find((r) => r.id === ld[1]);
    if (ld && (!vml || vml.tipo !== 'vmlDrawing' || vml.externa)) throw new ErrorEscritor('las notas de la hoja apuntan a un dibujo que no se encuentra');
    const leer = async (r) => { const t = r.externa ? null : await leerTexto(zip, r.destino); if (t === null) throw new ErrorEscritor('falta la parte ' + r.destino); return t; };
    const orig = {};
    const com = comentarios.length ? comentariosClasicos(orig.com = await leer(comentarios[0]), M) : null;
    const hil = hilos.length ? hilosDeComentarios(orig.hil = await leer(hilos[0]), M) : null;
    const dib = vml ? dibujoDeNotas(orig.dib = await leer(vml), M, quitadaAntes) : null;
    if (com && com.quedan === 0 && com.quitados > 0) {
      // ya no queda ninguna nota: Excel quita las tres partes y sus relaciones
      if (dib && dib.otras) throw new ErrorEscritor('el dibujo de las notas trae otras formas');
      for (const r of [comentarios[0], hilos[0], vml]) if (r) { plan.borrar.push(r.destino); plan.rels.add(r.id); }
      plan.legacy = !!vml;
      return plan;
    }
    if (hil && hil.quedan === 0 && hil.quitados > 0) { plan.borrar.push(hilos[0].destino); plan.rels.add(hilos[0].id); }
    else if (hil && hil.texto !== orig.hil) plan.nuevos.set(hilos[0].destino, hil.texto);
    if (com && com.texto !== orig.com) plan.nuevos.set(comentarios[0].destino, com.texto);
    if (dib && dib.texto !== orig.dib) plan.nuevos.set(vml.destino, dib.texto);
    return plan;
  }
  // quita partes del paquete (con su Override) y las relaciones de la hoja que las citaban (el .rels sin relaciones se borra, como hace Excel)
  async function borrarPartes(zip, rutaHoja, plan) {
    const buscar = (ruta) => { if (zip.file(ruta)) return ruta; let r = null; zip.forEach((n) => { if (!r && n.toLowerCase() === ruta.toLowerCase()) r = n; }); return r; };
    for (const ruta of plan.borrar) zip.remove(ruta);
    const ct = await leerTexto(zip, '[Content_Types].xml');
    if (ct !== null) {
      const idx = new Set(plan.borrar.map((r) => '/' + r.toLowerCase()));
      const u = ct.replace(/<Override\b[^>]*?\/>/g, (o) => { const p = /\bPartName\s*=\s*"([^"]*)"/.exec(o); return p && idx.has(p[1].toLowerCase()) ? '' : o; });
      if (u !== ct) zip.file('[Content_Types].xml', u);
    }
    const b = rutaHoja.lastIndexOf('/'), relsRuta = buscar(rutaHoja.slice(0, b + 1) + '_rels/' + rutaHoja.slice(b + 1) + '.rels');
    if (relsRuta) {
      const t = await leerTexto(zip, relsRuta);
      const u = t.replace(/<Relationship\b[^>]*?\/>/g, (r) => { const id = /\bId\s*=\s*"([^"]*)"/.exec(r); return id && plan.rels.has(id[1]) ? '' : r; });
      if (!/<Relationship\b/.test(u)) zip.remove(relsRuta); else if (u !== t) zip.file(relsRuta, u);
    }
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
    // las mismas tres cosas en las extensiones de Excel 2010 (listas de otra hoja, reglas que citan otra hoja, minigráficos)
    t = x14Objetivo(t, M, objetivo, cuenta, () => { partida = true; });
    // si una regla se partió, las prioridades se vuelven a numerar 1, 2, 3… en el mismo orden (las partes de una regla, seguidas; las de x14 comparten la cuenta)
    if (partida) {
      const RE_P = /(<(?:x14:)?cfRule\b[^>]*?\bpriority\s*=\s*")(\d+)(")/g;
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

  async function quitarFilas(libro, hoja, filas, mapeoDado) {
    const zip = libro.zip;
    const i = typeof hoja === 'number' ? hoja : libro.hojas.findIndex((h) => h.nombre === hoja);
    const entrada = libro.hojas[i];
    if (!entrada) throw new ErrorEscritor('hoja inexistente');
    if (entrada.tipo !== 'hoja' || !entrada.ruta) throw new ErrorEscritor('la hoja no es de celdas');
    if (!Array.isArray(filas)) throw new ErrorEscritor('las filas tienen que ser una lista');
    const lista = [...new Set(filas)].sort((a, b) => a - b);
    for (const f of lista) if (!Number.isInteger(f) || f < 0 || f >= 1048576) throw new ErrorEscritor('fila fuera de los límites de Excel');
    if (!lista.length && !mapeoDado) return { quitadas: 0 };
    const M = mapeoDado || mapeo(lista);
    const objetivo = entrada.nombre;
    const nuevos = new Map();   // ruta → texto (se escribe todo al final, o nada)

    const xmlObjetivo = await leerTexto(zip, entrada.ruta);
    if (xmlObjetivo === null) throw new ErrorEscritor('falta la parte ' + entrada.ruta);

    // 1. lo que no se sabe mover
    const rels = await XlsxPaquete.leerRels(zip, entrada.ruta);
    // (las notas y los comentarios en hilo sí se mueven: notasDeLaHoja; el dibujo de notas con controles se rechaza ahí)
    const prohibidos = rels.find((r) => /^(pivotTable|slicer|timeline|queryTable|ctrlProp|oleObject)$/.test(r.tipo));
    if (prohibidos) throw new ErrorEscritor('la hoja tiene ' + prohibidos.tipo + ' (tablas dinámicas, segmentaciones, controles o consultas): no se pueden mover con las filas');
    if (/<controls\b|<oleObjects\b/.test(xmlObjetivo)) throw new ErrorEscritor('la hoja tiene controles u objetos OLE');
    if (M.insercion) {
      // como Excel: no se inserta dentro de una fórmula matricial ni de una tabla de datos, ni si lo de abajo se saldría de la hoja
      const { r, n } = M.insercion;
      for (const m of xmlObjetivo.matchAll(/<(?:[\w.-]+:)?f\s([^>]*)>/g)) {
        if (!/\bt\s*=\s*["'](?:array|dataTable)["']/.test(m[1])) continue;
        const ref = /\bref\s*=\s*["']([^"']+)["']/.exec(m[1]), rg = ref && LibroXlsx.decodificarRango(ref[1]);
        if (rg && rg.f0 < r && rg.f1 >= r) throw new ErrorEscritor('no se pueden insertar filas dentro de una fórmula matricial o una tabla de datos');
      }
      let ultima = -1;
      for (const m of xmlObjetivo.matchAll(/<row\b[^>]*?\br\s*=\s*"(\d+)"/g)) ultima = Math.max(ultima, +m[1] - 1);
      if (ultima >= 0 && ultima + n > 1048575) throw new ErrorEscritor('lo de abajo se saldría de la hoja');
    }
    // las extensiones x14 se buscan con sus prefijos de siempre
    for (const m of xmlObjetivo.matchAll(/\bxmlns(?::([\w.-]+))?\s*=\s*"([^"]*)"/g)) {
      if ((m[2] === NS_X14 && m[1] !== 'x14') || (m[2] === NS_XM && m[1] !== 'xm')) throw new ErrorEscritor('la hoja usa otro prefijo para las extensiones de Excel 2010');
    }
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
    let datos = sd[0], altos = [];
    if (sd[1] !== '/>') {
      const r = filasDeLaHoja(datos, M, objetivo);
      altos = r.altos;
      datos = procesarFormulasDeTexto(r.xml, true, objetivo, M);
    }
    let antes = xml.slice(0, ini), cola = xml.slice(fin);
    cola = colaDeLaHoja(cola, M, objetivo);
    const defecto = parseFloat((/<sheetFormatPr\b[^>]*?\bdefaultRowHeight\s*=\s*"([\d.]+)"/.exec(antes) || [])[1]) || 15;
    const notas = await notasDeLaHoja(zip, rels, xml, M, alturaQuitada(M, altos, defecto));
    if (notas.legacy) cola = cola.replace(/<legacyDrawing\b[^>]*?\/>/, '');
    for (const [ruta, t] of notas.nuevos) nuevos.set(ruta, t);
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
      const n = parseInt(y[1], 10), ny = M.insercion ? (M.insercion.r < n ? n + M.insercion.n : n) : n - M.quita(0, n - 1);
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
    if (notas.borrar.length) await borrarPartes(zip, entrada.ruta, notas);
    if (libro._hojasCargadas) libro._hojasCargadas.clear();   // los modelos leídos antes ya no valen
    await XlsxPaquete.quitarCalcChain(zip);
    await XlsxPaquete.marcarRecalculo(zip);
    return { quitadas: lista.length, insertadas: M.insercion ? M.insercion.n : 0 };
  }
  async function insertarFilas(libro, hoja, fila, cuantas) {
    if (!Number.isInteger(fila) || fila < 0 || fila > 1048575 || !Number.isInteger(cuantas) || cuantas < 1 || cuantas > 1048576) throw new ErrorEscritor('fila fuera de los límites de Excel');
    return quitarFilas(libro, hoja, [], mapeoInsercion(fila, cuantas));
  }

  XlsxEscritor.insertarFilas = insertarFilas;
  XlsxEscritor.quitarFilas = quitarFilas;
  XlsxEscritor._quitarFilasInterno = { reescribirFormula, reescribirSqref, mapeo, partirPorFilas };
})();

/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Insertar y quitar COLUMNAS de una hoja en el lugar (suite, fase 2 del PLAN), como "Insertar / Eliminar columnas" de Excel: las columnas de la
// derecha se corren y todo lo que apunta a la hoja se acomoda: celdas y sus fórmulas, fórmulas de las otras hojas, nombres definidos, gráficos,
// anchos de columna, combinadas, formato condicional, validación de datos, hipervínculos, autofiltro, selección y paneles, notas y anclas de
// los dibujos. Es el gemelo de xlsx-filas.js para el otro eje; lo demás del paquete queda tal cual. No depende del navegador.
//
//   await XlsxEscritor.insertarColumnas(libro, hoja, columna, cuantas, { nombreColumna: 'Column' })   (nombre de las columnas nuevas de una tabla)
//   await XlsxEscritor.quitarColumnas(libro, hoja, [columnas desde 0])
// Dentro de una tabla, insertar le agrega columnas (como Excel) y quitar le saca las suyas (fase 3.2): sus tableColumn, los filtros de esas
// columnas, y las referencias estructuradas que las nombraban pasan a #¡REF! (Tabla1[Monto], [@Monto]…), como en Excel.
// Rechaza (ErrorEscritor, sin tocar nada), como hace el módulo de filas, lo que no sabe mover sin dejar un archivo que Excel repare: quitar
// todas las columnas de una tabla, fórmulas matriciales y tablas de datos partidas,
// tablas dinámicas, segmentaciones, controles y OLE, extensiones x14 de formato condicional o validación, y hojas con prefijo de espacio de nombres.
// Requiere xlsx-modelo.js, xml-editar.js, xlsx-paquete.js y xlsx-escritor.js.
(function () {
  const X = XmlEditar;
  const { ErrorEscritor } = XlsxEscritor;
  const MAX_COL = 16383;
  const colANum = (s) => LibroXlsx.colANumero(s.toUpperCase());
  const numACol = (n) => LibroXlsx.numeroACol(n);
  const texto = async (zip, ruta) => { const e = zip.file(ruta); return e ? e.async('string') : null; };

  // ---- qué columna va a dónde ------------------------------------------------------------------------------------------
  function mapeoInsercion(c, n) {
    const col = (x) => (x < c ? x : x + n);
    return { col, rango: (a, b) => { const x = col(a), y = col(b); return x > MAX_COL ? null : [x, Math.min(y, MAX_COL)]; }, quita: () => 0, insercion: { c, n } };
  }
  function mapeoQuitar(lista) {
    const q = lista;
    const hasta = (x, incluye) => { let lo = 0, hi = q.length; while (lo < hi) { const m = (lo + hi) >> 1; if (incluye ? q[m] <= x : q[m] < x) lo = m + 1; else hi = m; } return lo; };
    const set = new Set(q);
    return {
      col: (x) => (set.has(x) ? -1 : x - hasta(x, false)),
      rango: (a, b) => { const x = a - hasta(a, false), y = b - hasta(b, true); return x > y ? null : [x, y]; },
      quita: (a, b) => hasta(b, true) - hasta(a, false),
      lista: q,
    };
  }

  // ---- referencias en una fórmula ----------------------------------------------------------------------------------------
  // `propia`: la fórmula vive en la hoja que cambia (sus referencias sin hoja la apuntan). `objetivo`: el nombre de esa hoja.
  const RE_CELDA = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7}))?(?![A-Za-z0-9_(.!:])/;
  const RE_COLS = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(.!:])/;
  const RE_HOJA = /^('(?:[^']|'')*'|[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*)!/;
  const IDENT = /[A-Za-z0-9_.À-￿]/;
  function reescribirFormula(f, propia, objetivo, M) {
    if (!f) return f;
    const destino = objetivo.toLowerCase();
    let salida = '', i = 0;
    const n = f.length;
    while (i < n) {
      const ch = f[i];
      if (ch === '"') { let j = i + 1; while (j < n) { if (f[j] === '"') { if (f[j + 1] === '"') j += 2; else break; } else j++; } salida += f.slice(i, j + 1); i = j + 1; continue; }
      if (ch === '[') { let nivel = 0, j = i; for (; j < n; j++) { if (f[j] === '[') nivel++; else if (f[j] === ']') { nivel--; if (nivel === 0) break; } } salida += f.slice(i, j + 1); i = j + 1; continue; }
      const previo = i > 0 ? f[i - 1] : '';
      if (/[A-Za-z_$'À-￿]/.test(ch) && !(previo && (IDENT.test(previo) || previo === '!' || previo === ']'))) {
        let pos = i, deEsta = propia, prefijo = '';
        const mh = RE_HOJA.exec(f.slice(i, i + 300));
        if (mh) {
          prefijo = mh[0];
          const nombre = (mh[1][0] === "'" ? mh[1].slice(1, -1).replace(/''/g, "'") : mh[1]).toLowerCase();
          deEsta = nombre === destino;
          pos = i + mh[0].length;
        }
        const resto = f.slice(pos, pos + 40);
        let m;
        if ((m = RE_COLS.exec(resto)) && !/^\d/.test(resto)) {
          salida += prefijo;
          if (deEsta) {
            const a = colANum(m[2]), b = colANum(m[4]), r = M.rango(Math.min(a, b), Math.max(a, b));
            salida += r ? m[1] + numACol(r[0]) + ':' + m[3] + numACol(r[1]) : '#REF!';
          } else salida += m[0];
          i = pos + m[0].length; continue;
        }
        if ((m = RE_CELDA.exec(resto))) {
          salida += prefijo;
          if (deEsta) {
            const a = colANum(m[2]);
            if (m[5] === undefined) { const x = M.col(a); salida += x < 0 || x > MAX_COL ? '#REF!' : m[1] + numACol(x) + m[3] + m[4]; }
            else {
              const b = colANum(m[6]), r = M.rango(Math.min(a, b), Math.max(a, b));
              if (!r) salida += '#REF!';
              else { const [x, y] = a <= b ? r : [r[1], r[0]]; salida += m[1] + numACol(x) + m[3] + m[4] + ':' + m[5] + numACol(y) + m[7] + m[8]; }
            }
          } else salida += m[0];
          i = pos + m[0].length; continue;
        }
        // un identificador cualquiera (función, nombre): se copia entero para no cortarlo
        let j = pos;
        while (j < n && IDENT.test(f[j])) j++;
        salida += f.slice(i, Math.max(j, i + 1)); i = Math.max(j, i + 1); continue;
      }
      salida += ch; i++;
    }
    return salida;
  }
  // lista de rangos ("A1:B2 D5") de la hoja que cambia
  function reescribirSqref(s, M) {
    const r = [];
    for (const parte of String(s).trim().split(/\s+/)) {
      const g = LibroXlsx.decodificarRango(parte);
      if (!g) { r.push(parte); continue; }
      const nc = M.rango(g.c0, g.c1);
      if (!nc) continue;
      r.push(numACol(nc[0]) + (g.f0 + 1) + (g.f0 === g.f1 && nc[0] === nc[1] ? '' : ':' + numACol(nc[1]) + (g.f1 + 1)));
    }
    return r.join(' ');
  }

  // las fórmulas compartidas se escriben una a una: al correr columnas el desplazamiento entre la maestra y sus copias cambia
  function descompartir(xml) {
    // cada celda: <c …/> (vacía) o <c …>…</c>; la alternativa con "/>" primero evita tragarse la celda siguiente
    const re = /<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g;
    const maestras = new Map();
    for (const m of xml.matchAll(re)) {
      if (m[2] === '/>') continue;
      const r = /\br="([A-Za-z]{1,3})(\d+)"/.exec(m[1]);
      const f = /<f\b([^>]*)>([\s\S]*?)<\/f>/.exec(m[3]);
      if (r && f && /\bt="shared"/.test(f[1]) && /\bref="/.test(f[1])) { const si = (/\bsi="(\d+)"/.exec(f[1]) || [])[1]; if (si !== undefined) maestras.set(si, { texto: X.decodificar(f[2]), f: +r[2] - 1, c: colANum(r[1]) }); }
    }
    if (!maestras.size) return xml;
    return xml.replace(re, (todo, attrs, cierre, cuerpo) => {
      if (cierre === '/>') return todo;
      const r = /\br="([A-Za-z]{1,3})(\d+)"/.exec(attrs);
      if (!r) return todo;
      const nuevo = cuerpo.replace(/<f\b([^>]*?)(\/>|>([\s\S]*?)<\/f>)/, (fx, a) => {
        if (!/\bt="shared"/.test(a)) return fx;
        const si = (/\bsi="(\d+)"/.exec(a) || [])[1], ma = maestras.get(si);
        if (!ma) return fx;
        const t = LibroXlsx.desplazarFormula(ma.texto, +r[2] - 1 - ma.f, colANum(r[1]) - ma.c);
        const resto = a.replace(/\s*\bt="shared"/, '').replace(/\s*\bref="[^"]*"/, '').replace(/\s*\bsi="\d+"/, '');
        return '<f' + resto + '>' + X.esc(t) + '</f>';
      });
      return '<c' + attrs + '>' + nuevo + '</c>';
    });
  }

  // ---- la hoja que cambia ----------------------------------------------------------------------------------------------
  function hojaPropia(xml, M, objetivo) {
    const FX = (s) => reescribirFormula(s, true, objetivo, M);
    // celdas: r="B5" corre; las de columnas quitadas desaparecen; sus fórmulas se reescriben
    xml = xml.replace(/<row\b([^>]*?)(\/>|>([\s\S]*?)<\/row>)/g, (todo, attrs, cierre, cuerpo) => {
      const a = attrs.replace(/\s*\bspans="[^"]*"/, '');
      if (cierre === '/>') return '<row' + a + '/>';
      const celdas = cuerpo.replace(/<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g, (c, ca, cc, ci) => {
        const mr = /\br="([A-Za-z]{1,3})(\d+)"/.exec(ca);
        if (!mr) throw new ErrorEscritor('una celda sin dirección (r) no se puede correr');
        const x = M.col(colANum(mr[1]));
        if (x < 0) return '';
        if (x > MAX_COL) throw new ErrorEscritor('lo de la derecha se saldría de la hoja');
        const nca = ca.replace(mr[0], 'r="' + numACol(x) + mr[2] + '"');
        if (cc === '/>') return '<c' + nca + '/>';
        const ni = ci.replace(/(<f\b[^>]*>)([\s\S]*?)(<\/f>)/, (fx, a1, t, a2) => {
          const mref = /\bref="([^"]+)"/.exec(a1);
          let ap = a1;
          if (mref) { const g = reescribirSqref(mref[1], M); if (!g) throw new ErrorEscritor('se quita parte de una fórmula matricial'); ap = a1.replace(mref[0], 'ref="' + g + '"'); }
          return ap + X.esc(FX(X.decodificar(t))) + a2;
        });
        return '<c' + nca + '>' + ni + '</c>';
      });
      return '<row' + a + '>' + celdas + '</row>';
    });
    // anchos de columna: <col min max> corridos o recortados
    xml = xml.replace(/<cols\b[^>]*>([\s\S]*?)<\/cols>/, (todo, cuerpo) => {
      const nuevos = [];
      for (const m of cuerpo.matchAll(/<col\b[^>]*\/>/g)) {
        const et = m[0], a = +(/\bmin="(\d+)"/.exec(et) || [])[1] - 1, b = +(/\bmax="(\d+)"/.exec(et) || [])[1] - 1;
        if (!(a >= 0 && b >= a)) { nuevos.push(et); continue; }
        if (M.insercion && a < M.insercion.c && b >= M.insercion.c) {   // el tramo se parte: lo insertado hereda el formato de la columna de la izquierda (como Excel)
          nuevos.push(X.ponerAtributo(X.ponerAtributo(et, 'min', String(a + 1)), 'max', String(Math.min(MAX_COL, b + M.insercion.n) + 1)));
          continue;
        }
        const r = M.rango(a, b);
        if (!r) continue;
        nuevos.push(X.ponerAtributo(X.ponerAtributo(et, 'min', String(r[0] + 1)), 'max', String(r[1] + 1)));
      }
      return nuevos.length ? '<cols>' + nuevos.join('') + '</cols>' : '';
    });
    const sq = (s) => reescribirSqref(s, M);
    xml = xml.replace(/(<dimension\b[^>]*?\bref=")([^"]+)(")/, (t, a, r, c) => a + (sq(r) || 'A1') + c);
    xml = xml.replace(/(<mergeCells\b[^>]*>)([\s\S]*?)(<\/mergeCells>)/, (t, a, cuerpo, c) => {
      let n = 0;
      const u = cuerpo.replace(/<mergeCell\b[^>]*?\bref="([^"]+)"[^>]*\/>/g, (m, r) => { const g = LibroXlsx.decodificarRango(r); const nc = g && M.rango(g.c0, g.c1); if (!nc || (nc[0] === nc[1] && g.f0 === g.f1)) return ''; n++; return m.replace(r, sq(r)); });
      return n ? X.ponerAtributo(a, 'count', String(n)) + u + c : '';
    });
    // lo que queda sin rango (todas sus columnas se quitaron) desaparece; lo demás corre
    xml = xml.replace(/<conditionalFormatting\b([^>]*?)\bsqref="([^"]+)"([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g, (t, a, r, b, cuerpo) => { const n = sq(r); return n ? '<conditionalFormatting' + a + 'sqref="' + n + '"' + b + '>' + cuerpo + '</conditionalFormatting>' : ''; });
    xml = xml.replace(/<dataValidation\b([^>]*?)\bsqref="([^"]+)"([^>]*?)(\/>|>([\s\S]*?)<\/dataValidation>)/g, (t, a, r, b, cierre, cuerpo) => { const n = sq(r); return n ? '<dataValidation' + a + 'sqref="' + n + '"' + b + (cierre === '/>' ? '/>' : '>' + cuerpo + '</dataValidation>') : ''; });
    xml = xml.replace(/<hyperlink\b([^>]*?)\bref="([^"]+)"([^>]*?)\/>/g, (t, a, r, b) => { const n = sq(r); return n ? '<hyperlink' + a + 'ref="' + n + '"' + b + '/>' : ''; });
    xml = xml.replace(/(<(?:selection|ignoredError|protectedRange)\b[^>]*?\bsqref=")([^"]+)(")/g, (t, a, r, c) => a + (sq(r) || 'A1') + c);
    xml = xml.replace(/<dataValidations\b([^>]*)>([\s\S]*?)<\/dataValidations>/, (t, a, cuerpo) => { const n = (cuerpo.match(/<dataValidation\b/g) || []).length; return n ? '<dataValidations' + a.replace(/\bcount="\d+"/, 'count="' + n + '"') + '>' + cuerpo + '</dataValidations>' : ''; });
    xml = xml.replace(/<hyperlinks\b[^>]*>\s*<\/hyperlinks>/, '');
    xml = xml.replace(/(<(?:formula|formula1|formula2)>)([\s\S]*?)(<\/(?:formula|formula1|formula2)>)/g, (t, a, f, c) => a + X.esc(FX(X.decodificar(f))) + c);
    xml = xml.replace(/(<autoFilter\b[^>]*?\bref=")([^"]+)(")/, (t, a, r, c) => a + (sq(r) || 'A1') + c);
    xml = xml.replace(/(<selection\b[^>]*?\bactiveCell=")([A-Za-z]{1,3})(\d+)(")/g, (t, a, l, f, c) => { const x = M.col(colANum(l)); return a + numACol(Math.max(0, x < 0 ? 0 : Math.min(x, MAX_COL))) + f + c; });
    xml = xml.replace(/<pane\b[^>]*\/?>/, (p) => {
      const xs = /\bxSplit="(\d+)"/.exec(p);
      if (!xs || !/\bstate="frozen/.test(p)) return p;
      const n = +xs[1], nx = M.insercion ? (M.insercion.c < n ? n + M.insercion.n : n) : n - M.quita(0, n - 1);
      if (nx <= 0) return p;
      return p.replace(xs[0], 'xSplit="' + nx + '"').replace(/(\btopLeftCell=")([A-Za-z]{1,3})(\d+)(")/, (t, a, l, f, c) => a + numACol(Math.min(MAX_COL, nx)) + f + c);
    });
    xml = xml.replace(/(<colBreaks\b[^>]*>)([\s\S]*?)(<\/colBreaks>)/, (t, a, cuerpo, c) => {
      let n = 0;
      const u = cuerpo.replace(/<brk\b[^>]*?\bid="(\d+)"[^>]*\/>/g, (m, id) => { const x = M.col(+id); if (x <= 0) return ''; n++; return m.replace('id="' + id + '"', 'id="' + x + '"'); });
      return n ? X.ponerAtributo(X.ponerAtributo(a, 'count', String(n)), 'manualBreakCount', String(n)) + u + c : '';
    });
    return xml;
  }

  // ---- la operación ----------------------------------------------------------------------------------------------------
  // Agrega n columnas a la tabla (su XML `t`, rango g) en la posición de la inserción: <tableColumn> nuevos con id y nombre únicos, count,
  // y el colId de los filtros que quedan a la derecha. Anota en `encabezados` la celda de encabezado que lleva cada una (con el estilo de la
  // celda de la izquierda, como Excel). Las referencias (ref de la tabla y del autofiltro) las corre después reescribirSqref.
  function columnasEnTabla(t, g, ins, opciones, encabezados, hojaXml) {
    const raiz = X.analizar(t);
    const cols = X.hijo(raiz, 'tableColumns');
    if (!cols || cols.iniCierre < 0) throw new ErrorEscritor('una tabla sin tableColumns');
    const lista = X.hijos(cols, 'tableColumn');
    if (lista.length !== g.c1 - g.c0 + 1) throw new ErrorEscritor('una tabla con columnas que no coinciden con su rango');
    const k = ins.c - g.c0;
    const usados = new Set(lista.map((x) => String(x.attrs.name || '').toLowerCase()));
    let id = Math.max(0, ...lista.map((x) => +x.attrs.id || 0));
    const base = (opciones && opciones.nombreColumna) || 'Column';
    const conEncabezado = raiz.attrs.headerRowCount !== '0';
    const estiloIzq = conEncabezado ? (new RegExp('<c\\b[^>]*\\br="' + numACol(ins.c - 1) + (g.f0 + 1) + '"[^>]*>').exec(hojaXml) || [''])[0] : '';
    const s = (/\bs="(\d+)"/.exec(estiloIzq) || [])[1];
    let nuevas = '', j = 1;
    for (let i = 0; i < ins.n; i++) {
      while (usados.has((base + j).toLowerCase())) j++;
      const nombre = base + j; usados.add(nombre.toLowerCase());
      nuevas += '<tableColumn id="' + (++id) + '" name="' + X.escAttr(nombre) + '"/>';
      if (conEncabezado) encabezados.push({ f: g.f0, c: ins.c + i, texto: nombre, s });
    }
    const ediciones = [{ ini: lista[k].ini, fin: lista[k].ini, texto: nuevas }];
    ediciones.push({ ini: cols.ini, fin: cols.finApertura, texto: X.ponerAtributo(t.slice(cols.ini, cols.finApertura), 'count', String(lista.length + ins.n)) });
    t = X.aplicar(t, ediciones);
    // los filtros de las columnas de la derecha se corren
    return t.replace(/(<filterColumn\b[^>]*?\bcolId=")(\d+)(")/g, (m, a, n, c) => a + (+n >= k ? +n + ins.n : +n) + c);
  }
  // Quita de la tabla (su XML `t`, rango g) las columnas que se van: sus <tableColumn>, count, los filtros y ordenaciones de esas columnas, y
  // corre el colId de los demás filtros. Devuelve { t, quitadas: [nombres en minúsculas] }.
  function quitarDeTabla(t, g, M) {
    const raiz = X.analizar(t, { profundidad: 4 });
    const cols = X.hijo(raiz, 'tableColumns');
    if (!cols || cols.iniCierre < 0) throw new ErrorEscritor('una tabla sin tableColumns');
    const lista = X.hijos(cols, 'tableColumn');
    if (lista.length !== g.c1 - g.c0 + 1) throw new ErrorEscritor('una tabla con columnas que no coinciden con su rango');
    const se = lista.map((x, k) => M.col(g.c0 + k) < 0);
    const quitadas = lista.filter((x, k) => se[k]).map((x) => X.decodificar(x.attrs.name || '').toLowerCase());
    const ediciones = lista.filter((x, k) => se[k]).map((x) => ({ ini: x.ini, fin: x.fin, texto: '' }));
    ediciones.push({ ini: cols.ini, fin: cols.finApertura, texto: X.ponerAtributo(t.slice(cols.ini, cols.finApertura), 'count', String(lista.length - quitadas.length)) });
    // filtros: el de una columna quitada se va; los de la derecha corren su colId
    const nuevoId = (k) => { let n = 0; for (let i = 0; i < k; i++) if (!se[i]) n++; return n; };
    const af = X.hijo(raiz, 'autoFilter');
    if (af && af.iniCierre >= 0) for (const fc of X.hijos(af, 'filterColumn')) {
      const k = +fc.attrs.colId;
      if (se[k]) ediciones.push({ ini: fc.ini, fin: fc.fin, texto: '' });
      else if (nuevoId(k) !== k) ediciones.push({ ini: fc.ini, fin: fc.finApertura, texto: X.ponerAtributo(t.slice(fc.ini, fc.finApertura), 'colId', String(nuevoId(k))) });
    }
    t = X.aplicar(t, ediciones);
    // ordenaciones de una columna quitada (su ref es solo esa columna): se van
    t = t.replace(/<sortCondition\b[^>]*?\bref="([^"]+)"[^>]*\/>/g, (m, rf) => { const r = LibroXlsx.decodificarRango(rf); return r && M.quita(r.c0, r.c1) === r.c1 - r.c0 + 1 ? '' : m; });
    return { t, quitadas };
  }
  // Referencias estructuradas a columnas quitadas → #REF! (Excel hace lo mismo). `tablas`: [{ nombre (minúsculas), cols: Set }]; las
  // referencias sin nombre de tabla ([@Monto]) se buscan en `propias` (las tablas de la hoja de la fórmula).
  function sinColumnasQuitadas(f, tablas, propias) {
    if (!f || f.indexOf('[') < 0) return f;
    let salida = '', i = 0;
    const n = f.length;
    while (i < n) {
      const ch = f[i];
      if (ch === '"') { let j = i + 1; while (j < n) { if (f[j] === '"') { if (f[j + 1] === '"') j += 2; else break; } else j++; } salida += f.slice(i, j + 1); i = j + 1; continue; }
      if (ch !== '[') { salida += ch; i++; continue; }
      // el nombre de la tabla va justo antes del corchete
      const m = /([A-Za-z_\\À-￿][A-Za-z0-9_.\\À-￿]*)$/.exec(salida);
      let nivel = 0, j = i;
      for (; j < n; j++) { if (f[j] === "'" ) { j++; continue; } if (f[j] === '[') nivel++; else if (f[j] === ']') { nivel--; if (nivel === 0) break; } }
      const cuerpo = f.slice(i, j + 1);
      const candidatas = m ? tablas.filter((x) => x.nombre === m[1].toLowerCase()) : propias;
      const nombres = [];
      for (const x of cuerpo.matchAll(/\[((?:[^\[\]']|'.)+)\]/g)) nombres.push(x[1].replace(/^@/, '').replace(/'(.)/g, '$1').trim().toLowerCase());
      if (!nombres.length) { const x = cuerpo.slice(1, -1).replace(/^@/, '').replace(/'(.)/g, '$1').trim().toLowerCase(); if (x) nombres.push(x); }
      if (candidatas.some((x) => nombres.some((nm) => x.cols.has(nm)))) { if (m) salida = salida.slice(0, salida.length - m[1].length); salida += '#REF!'; }
      else salida += cuerpo;
      i = j + 1;
    }
    return salida;
  }
  // escribe las celdas de encabezado (texto en línea) en su fila, en orden de columna; la fila del encabezado siempre existe (tiene los otros)
  function ponerEncabezados(xml, encabezados) {
    if (!encabezados.length) return xml;
    const porFila = new Map();
    for (const e of encabezados) { if (!porFila.has(e.f)) porFila.set(e.f, []); porFila.get(e.f).push(e); }
    const raiz = X.analizar(xml, { profundidad: 4 });
    const datos = X.hijo(raiz, 'sheetData');
    const ediciones = [];
    for (const [f, lista] of porFila) {
      const fila = datos && X.hijos(datos, 'row').find((r) => +r.attrs.r === f + 1);
      if (!fila || fila.iniCierre < 0) throw new ErrorEscritor('falta la fila del encabezado de la tabla');
      const celdas = X.hijos(fila, 'c');
      for (const e of lista) {
        const ref = numACol(e.c) + (f + 1);
        if (celdas.some((x) => x.attrs.r === ref)) {   // la columna insertada trae la celda vacía (con formato): se reemplaza
          const x = celdas.find((y) => y.attrs.r === ref);
          ediciones.push({ ini: x.ini, fin: x.fin, texto: '<c r="' + ref + '"' + (x.attrs.s ? ' s="' + x.attrs.s + '"' : e.s ? ' s="' + e.s + '"' : '') + ' t="inlineStr"><is><t>' + X.esc(e.texto) + '</t></is></c>' });
          continue;
        }
        const sig = celdas.find((x) => colANum(String(x.attrs.r || '').replace(/\d+$/, '')) > e.c);
        const pos = sig ? sig.ini : fila.iniCierre;
        ediciones.push({ ini: pos, fin: pos, texto: '<c r="' + ref + '"' + (e.s ? ' s="' + e.s + '"' : '') + ' t="inlineStr"><is><t>' + X.esc(e.texto) + '</t></is></c>' });
      }
    }
    return X.aplicar(xml, ediciones);
  }

  async function moverColumnas(libro, hoja, M, opciones) {
    const zip = libro.zip;
    const i = typeof hoja === 'number' ? hoja : libro.hojas.findIndex((h) => h.nombre === hoja);
    const entrada = libro.hojas[i];
    if (!entrada || entrada.tipo !== 'hoja' || !entrada.ruta) throw new ErrorEscritor('la hoja no es de celdas');
    const objetivo = entrada.nombre;
    const nuevos = new Map();
    const xml0 = await texto(zip, entrada.ruta);
    if (xml0 === null) throw new ErrorEscritor('falta la parte ' + entrada.ruta);
    const rels = await XlsxPaquete.leerRels(zip, entrada.ruta);
    const prohibido = rels.find((r) => /^(pivotTable|slicer|timeline|queryTable|ctrlProp|oleObject)$/.test(r.tipo));
    if (prohibido) throw new ErrorEscritor('la hoja tiene ' + prohibido.tipo + ': no se pueden mover sus columnas aquí');
    if (/<controls\b|<oleObjects\b/.test(xml0)) throw new ErrorEscritor('la hoja tiene controles u objetos OLE');
    if (/<(?:[\w.-]+:)(?:sheetData|row|c)[\s/>]/.test(xml0)) throw new ErrorEscritor('la hoja está escrita con prefijos de espacio de nombres');
    if (/<x14:conditionalFormatting\b|<x14:dataValidation\b|<x14:sparklineGroup\b/.test(xml0)) throw new ErrorEscritor('la hoja tiene extensiones de Excel 2010 (formato condicional, validación o minigráficos nuevos): mueve sus columnas en Excel');
    // fórmulas matriciales y tablas de datos: no se insertan adentro ni se quitan en parte (como Excel)
    for (const m of xml0.matchAll(/<f\s([^>]*)>/g)) {
      if (!/\bt="(?:array|dataTable)"/.test(m[1])) continue;
      const g = LibroXlsx.decodificarRango((/\bref="([^"]+)"/.exec(m[1]) || [])[1] || '');
      if (!g) continue;
      if (M.insercion && g.c0 < M.insercion.c && g.c1 >= M.insercion.c) throw new ErrorEscritor('no se pueden insertar columnas dentro de una fórmula matricial');
      const q = M.quita(g.c0, g.c1);
      if (q && q < g.c1 - g.c0 + 1) throw new ErrorEscritor('se quita parte de una fórmula matricial');
    }
    // tablas: se corren enteras. Insertar adentro le agrega columnas a la tabla (como Excel: «Column1»… con el encabezado escrito en la hoja);
    // quitar parte de una tabla se rechaza (las referencias estructuradas a la columna quitada tendrían que pasar a #¡REF!)
    const encabezados = [];   // [{ f, c, texto, s }] celdas de encabezado de las columnas nuevas
    const tablasQuitadas = [];   // [{ nombre, cols }] las tablas que pierden columnas (para sus referencias estructuradas)
    for (const r of rels.filter((x) => x.tipo === 'table')) {
      let t = await texto(zip, r.destino);
      if (t === null) continue;
      const g = LibroXlsx.decodificarRango((/<table\b[^>]*?\bref="([^"]+)"/.exec(t) || [])[1] || '');
      if (!g) continue;
      if (!M.insercion && M.quita(g.c0, g.c1)) {
        if (M.quita(g.c0, g.c1) === g.c1 - g.c0 + 1) throw new ErrorEscritor('la operación quita todas las columnas de una tabla: conviértela en rango o quítala en Excel');
        const q = quitarDeTabla(t, g, M);
        t = q.t;
        const nombre = X.decodificar((/<table\b[^>]*?\bdisplayName="([^"]+)"/.exec(t) || /<table\b[^>]*?\bname="([^"]+)"/.exec(t) || [])[1] || '').toLowerCase();
        tablasQuitadas.push({ nombre, cols: new Set(q.quitadas) });
      }
      if (M.insercion && g.c0 < M.insercion.c && g.c1 >= M.insercion.c) t = columnasEnTabla(t, g, M.insercion, opciones, encabezados, xml0);
      const u = t.replace(/(<(?:table|autoFilter)\b[^>]*?\bref=")([^"]+)(")/g, (x, a, rf, c) => a + reescribirSqref(rf, M) + c)
        .replace(/(<(?:calculatedColumnFormula|totalsRowFormula)\b[^>]*>)([\s\S]*?)(<\/(?:calculatedColumnFormula|totalsRowFormula)>)/g, (x, a, f, c) => a + X.esc(reescribirFormula(X.decodificar(f), true, objetivo, M)) + c);
      if (u !== t || encabezados.length) nuevos.set(r.destino, u);
    }
    nuevos.set(entrada.ruta, ponerEncabezados(hojaPropia(descompartir(xml0), M, objetivo), encabezados));

    // notas: <comment ref> y la columna de su cuadro en el VML; comentarios en hilo
    for (const r of rels.filter((x) => x.tipo === 'comments' || x.tipo === 'threadedComment')) {
      const t = await texto(zip, r.destino);
      if (t === null) continue;
      nuevos.set(r.destino, t.replace(/(<(?:comment|threadedComment)\b[^>]*?\bref=")([A-Za-z]{1,3})(\d+)(")/g, (m, a, l, f, c) => { const x = M.col(colANum(l)); return x < 0 ? a + 'A' + f + c : a + numACol(x) + f + c; }));
    }
    for (const r of rels.filter((x) => x.tipo === 'vmlDrawing')) {
      const t = await texto(zip, r.destino);
      if (t === null) continue;
      nuevos.set(r.destino, t.replace(/(<x:Column>)(\d+)(<\/x:Column>)/g, (m, a, n, c) => { const x = M.col(+n); return a + Math.max(0, x) + c; }));
    }
    // dibujos: la columna de las anclas
    for (const r of rels.filter((x) => x.tipo === 'drawing')) {
      const t = await texto(zip, r.destino);
      if (t === null) continue;
      nuevos.set(r.destino, t.replace(/<((?:\w+:)?)(twoCellAnchor|oneCellAnchor)\b([^>]*)>([\s\S]*?)<\/\1\2>/g, (todo, p, tipo, attrs, cuerpo) => {
        if ((/\beditAs="(\w+)"/.exec(attrs) || [])[1] === 'absolute') return todo;
        const c = cuerpo.replace(/(<(?:\w+:)?(?:from|to)>[\s\S]*?<(?:\w+:)?col>)(\d+)(<\/)/g, (m, a, n, z) => { const x = M.col(+n); return a + (x < 0 ? Math.max(0, +n - M.quita(0, +n)) : x) + z; });
        return `<${p}${tipo}${attrs}>${c}</${p}${tipo}>`;
      }));
    }
    // otras hojas, nombres definidos y gráficos que citan esta hoja
    const marca = objetivo.toLowerCase();
    const cita = (t) => t.toLowerCase().includes(marca) || t.toLowerCase().includes(X.esc(marca));
    for (const h of libro.hojas) {
      if (h.tipo !== 'hoja' || !h.ruta || h.ruta === entrada.ruta) continue;
      const t = await texto(zip, h.ruta);
      if (t === null || !cita(t)) continue;
      const u = descompartir(t).replace(/(<f\b[^>]*>)([\s\S]*?)(<\/f>)/g, (m, a, f, c) => a + X.esc(reescribirFormula(X.decodificar(f), false, objetivo, M)) + c)
        .replace(/(<(?:formula|formula1|formula2)>)([\s\S]*?)(<\/(?:formula|formula1|formula2)>)/g, (m, a, f, c) => a + X.esc(reescribirFormula(X.decodificar(f), false, objetivo, M)) + c);
      if (u !== t) nuevos.set(h.ruta, u);
    }
    const wb = await texto(zip, libro.rutaLibro);
    if (wb !== null) {
      const u = wb.replace(/(<definedName\b([^>]*)>)([\s\S]*?)(<\/definedName>)/g, (m, a, attrs, f, c) => {
        return a + X.esc(reescribirFormula(X.decodificar(f), false, objetivo, M)) + c;   // los nombres siempre citan su hoja
      });
      if (u !== wb) nuevos.set(libro.rutaLibro, u);
    }
    const graficos = [];
    zip.forEach((r) => { if (/^xl\/charts\/[^/]+\.xml$/i.test(r)) graficos.push(r); });
    for (const r of graficos) {
      const t = await texto(zip, r);
      if (t === null || !cita(t)) continue;
      nuevos.set(r, t.replace(/(<(?:\w+:)?f>)([\s\S]*?)(<\/(?:\w+:)?f>)/g, (m, a, f, c) => a + X.esc(reescribirFormula(X.decodificar(f), false, objetivo, M)) + c));
    }
    // referencias estructuradas a las columnas quitadas: en todas las hojas, los nombres definidos y las fórmulas de las tablas
    if (tablasQuitadas.length) {
      const enF = (t, propias) => t.replace(/(<(?:f|formula|formula1|formula2|calculatedColumnFormula|totalsRowFormula)\b[^>]*>)([\s\S]*?)(<\/(?:f|formula|formula1|formula2|calculatedColumnFormula|totalsRowFormula)>)/g, (m, a, f, c) => { const d = X.decodificar(f), u = sinColumnasQuitadas(d, tablasQuitadas, propias); return u === d ? m : a + X.esc(u) + c; });
      for (const h of libro.hojas) {
        if (h.tipo !== 'hoja' || !h.ruta) continue;
        const t = nuevos.has(h.ruta) ? nuevos.get(h.ruta) : await texto(zip, h.ruta);
        if (t === null || t.indexOf('[') < 0) continue;
        const u = enF(t, h.ruta === entrada.ruta ? tablasQuitadas : []);
        if (u !== t) nuevos.set(h.ruta, u);
      }
      for (const r of rels.filter((x) => x.tipo === 'table')) { const t = nuevos.has(r.destino) ? nuevos.get(r.destino) : await texto(zip, r.destino); if (t) { const u = enF(t, tablasQuitadas); if (u !== t) nuevos.set(r.destino, u); } }
      const w = nuevos.has(libro.rutaLibro) ? nuevos.get(libro.rutaLibro) : await texto(zip, libro.rutaLibro);
      if (w) { const u = w.replace(/(<definedName\b[^>]*>)([\s\S]*?)(<\/definedName>)/g, (m, a, f, c) => { const d = X.decodificar(f), v = sinColumnasQuitadas(d, tablasQuitadas, []); return v === d ? m : a + X.esc(v) + c; }); if (u !== w) nuevos.set(libro.rutaLibro, u); }
    }
    for (const [ruta, t] of nuevos) zip.file(ruta, t);
    if (libro._hojasCargadas) libro._hojasCargadas.clear();
    await XlsxPaquete.quitarCalcChain(zip);
    await XlsxPaquete.marcarRecalculo(zip);
  }

  async function insertarColumnas(libro, hoja, columna, cuantas, opciones) {
    if (!Number.isInteger(columna) || columna < 0 || columna > MAX_COL || !Number.isInteger(cuantas) || cuantas < 1 || cuantas > MAX_COL + 1) throw new ErrorEscritor('columna fuera de los límites de Excel');
    return moverColumnas(libro, hoja, mapeoInsercion(columna, cuantas), opciones);
  }
  async function quitarColumnas(libro, hoja, columnas) {
    const lista = [...new Set(columnas)].sort((a, b) => a - b);
    for (const c of lista) if (!Number.isInteger(c) || c < 0 || c > MAX_COL) throw new ErrorEscritor('columna fuera de los límites de Excel');
    if (!lista.length) return;
    return moverColumnas(libro, hoja, mapeoQuitar(lista));
  }

  XlsxEscritor.insertarColumnas = insertarColumnas;
  XlsxEscritor.quitarColumnas = quitarColumnas;
  XlsxEscritor._columnasInterno = { reescribirFormula, reescribirSqref, mapeoInsercion, mapeoQuitar, sinColumnasQuitadas };
})();

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
// formato condicional, validación o minigráficos, el encabezado o la fila de totales de una tabla, y una tabla que quedaría sin filas.
// Las fórmulas compartidas de las hojas afectadas se escriben una a una (al quitar filas el desplazamiento entre la maestra y sus copias cambia).
// Requiere xlsx-modelo.js, xml-editar.js, xlsx-paquete.js y xlsx-escritor.js.
(function () {
  const X = XmlEditar;
  const { ErrorEscritor } = XlsxEscritor;
  const colANumero = (s) => LibroXlsx.colANumero(s);
  const numeroACol = (n) => LibroXlsx.numeroACol(n);
  const IDENT = /[A-Za-z0-9_.À-￿]/;
  const MAX_PARTES = 5000;

  // ---- qué fila va a dónde ---------------------------------------------------------------------------------------
  function mapeo(filas) {
    const q = filas;   // ordenadas y sin repetir
    const hasta = (f, incluye) => { let lo = 0, hi = q.length; while (lo < hi) { const m = (lo + hi) >> 1; if (incluye ? q[m] <= f : q[m] < f) lo = m + 1; else hi = m; } return lo; };
    return {
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
          if (prefijo) { salida += prefijo; i = pos; continue; }
          salida += ch; i++; continue;
        }
        const todo = (m || mf)[0];
        const aplica = !externo && (prefijo ? hoja.toLowerCase() === destino : enObjetivo);
        externo = false;
        if (!aplica) { salida += prefijo + todo; i = pos + todo.length; continue; }
        let nuevo;
        if (m) {
          const f0 = parseInt(m[4], 10) - 1;
          if (!m[8]) { const f = M.fila(f0); nuevo = f < 0 ? '#REF!' : m[1] + m[2] + m[3] + (f + 1); }
          else {
            const f1 = parseInt(m[8], 10) - 1, r = M.rango(Math.min(f0, f1), Math.max(f0, f1));
            nuevo = r ? m[1] + m[2] + m[3] + (r[0] + 1) + ':' + m[5] + m[6] + m[7] + (r[1] + 1) : '#REF!';
          }
        } else {
          const a = parseInt(mf[2], 10) - 1, b = parseInt(mf[4], 10) - 1, r = M.rango(Math.min(a, b), Math.max(a, b));
          nuevo = r ? mf[1] + (r[0] + 1) + ':' + mf[3] + (r[1] + 1) : '#REF!';
        }
        salida += prefijo + nuevo; i = pos + todo.length; continue;
      }
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
  function procesarFormulasDeTexto(texto, enObjetivo, objetivo, M) {
    // <f>…</f> de celdas y <formula>, <formula1>, <formula2> de formato condicional y validación
    return texto
      .replace(/(<f\b[^>]*>)([\s\S]*?)(<\/f>)/g, (t, a, f, c) => a + X.esc(reescribirFormula(X.decodificar(f), enObjetivo, objetivo, M)) + c)
      .replace(/(<formula[12]?\b[^>]*>)([\s\S]*?)(<\/formula[12]?>)/g, (t, a, f, c) => a + X.esc(reescribirFormula(X.decodificar(f), enObjetivo, objetivo, M)) + c);
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
    // formato condicional
    t = t.replace(/<conditionalFormatting\b([^>]*?)\bsqref\s*=\s*"([^"]*)"([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g, (todo, a, sq, b, hijos) => {
      const nuevo = reescribirSqref(sq, M);
      return nuevo ? `<conditionalFormatting${a}sqref="${nuevo}"${b}>${procesarFormulasDeTexto(hijos, true, objetivo, M)}</conditionalFormatting>` : '';
    });
    // validación de datos
    t = t.replace(/<dataValidation\b([^>]*?)\bsqref\s*=\s*"([^"]*)"([^>]*?)(\/>|>([\s\S]*?)<\/dataValidation>)/g, (todo, a, sq, b, cierre, hijos) => {
      const nuevo = reescribirSqref(sq, M);
      if (!nuevo) return '';
      return cierre === '/>' ? `<dataValidation${a}sqref="${nuevo}"${b}/>` : `<dataValidation${a}sqref="${nuevo}"${b}>${procesarFormulasDeTexto(hijos, true, objetivo, M)}</dataValidation>`;
    });
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
    // una tabla dinámica de otra hoja que toma sus datos de esta
    for (const h of libro.hojas) {
      if (h.tipo !== 'hoja' || !h.ruta) continue;
      for (const r of await XlsxPaquete.leerRels(zip, h.ruta)) if (r.tipo === 'pivotTable' && h.ruta !== entrada.ruta) {
        const cache = (await XlsxPaquete.leerRels(zip, r.destino)).find((c) => c.tipo === 'pivotCacheDefinition');
        const t = cache && await leerTexto(zip, cache.destino);
        if (t && new RegExp('<worksheetSource\\b[^>]*\\bsheet\\s*=\\s*"' + objetivo.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"', 'i').test(t)) throw new ErrorEscritor('una tabla dinámica toma sus datos de esta hoja');
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
      const cambiar = (s) => s.replace(/(<(?:table|autoFilter)\b[^>]*?\bref\s*=\s*")([^"]+)(")/g, (x, a, rf, c) => a + reescribirSqref(rf, M) + c);
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
    nuevos.set(entrada.ruta, antes + datos + cola);

    // 4. dibujos de la hoja: las anclas suben con sus filas
    for (const r of rels.filter((x) => x.tipo === 'drawing')) {
      const t = await leerTexto(zip, r.destino);
      if (t === null) continue;
      const u = t
        .replace(/(<(?:\w+:)?from>[\s\S]*?<(?:\w+:)?row>)(\d+)(<\/)/g, (x, a, fila, c) => a + (parseInt(fila, 10) - M.antes(parseInt(fila, 10))) + c)
        .replace(/(<(?:\w+:)?to>[\s\S]*?<(?:\w+:)?row>)(\d+)(<\/)/g, (x, a, fila, c) => a + Math.max(0, parseInt(fila, 10) - M.hasta(parseInt(fila, 10))) + c);
      nuevos.set(r.destino, u);
    }

    // 5. otras hojas, nombres definidos y gráficos: solo las referencias que llevan el nombre de esta hoja
    const marca = objetivo.toLowerCase();
    const rutaLibro = await XlsxPaquete.rutaLibroDe(zip);
    const wb = await leerTexto(zip, rutaLibro);
    let partes = 0;
    for (const h of libro.hojas) {
      if (h.tipo !== 'hoja' || !h.ruta || h.ruta === entrada.ruta) continue;
      const t = await leerTexto(zip, h.ruta);
      if (t === null || !t.toLowerCase().includes(marca)) continue;
      if (++partes > MAX_PARTES) throw new ErrorEscritor('demasiadas partes');
      let u = descompartir(t);
      u = procesarFormulasDeTexto(u, false, objetivo, M);
      nuevos.set(h.ruta, u);
    }
    if (wb !== null) {
      const u = wb.replace(/(<definedName\b[^>]*>)([\s\S]*?)(<\/definedName>)/g, (x, a, f, c) => a + X.esc(reescribirFormula(X.decodificar(f), false, objetivo, M)) + c);
      if (u !== wb) nuevos.set(rutaLibro, u);
    }
    const graficos = [];
    zip.forEach((nombre) => { if (/^xl\/charts\/[^/]+\.xml$/i.test(nombre)) graficos.push(nombre); });
    for (const nombre of graficos) {
      const t = await leerTexto(zip, nombre);
      if (t === null || !t.toLowerCase().includes(marca)) continue;
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
  XlsxEscritor._quitarFilasInterno = { reescribirFormula, reescribirSqref, mapeo };
})();

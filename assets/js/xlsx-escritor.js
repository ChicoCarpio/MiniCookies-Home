/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Escritor de libros de Excel (etapa 7), sin DOM, para Node y el navegador. Dos operaciones:
//
//   await XlsxEscritor.editarCeldas(libro, hoja, cambios)
//     Cambia el valor de celdas de un libro ABIERTO con LibroXlsx.abrir, en el lugar: el resto del paquete (estilos, dibujos, tablas, nombres, macros,
//     extensiones) y de la hoja queda tal cual. `hoja`: índice o nombre. `cambios`: [{ f, c, valor, estilo? }] (f y c desde 0); valor: número, texto,
//     lógico, { error: '#N/A' } o null (vaciar la celda conservando su formato). `estilo`: índice de cellXfs (si no, la celda conserva el suyo).
//     Una celda con fórmula pasa a ser un valor fijo. Se borra calcChain y se pide a Excel recalcular al abrir. Después se guarda con
//     libro.zip.generateAsync(...). Rechaza (ErrorEscritor): fuera de los límites de Excel, texto de más de 32.767 caracteres, hoja que no es de celdas,
//     celdas de una fórmula matricial y encabezados de tabla (cambiarlos sin cambiar la tabla hace que Excel repare el archivo).
//
//   await XlsxEscritor.crearLibro(hojas, opciones) → Uint8Array
//     Libro nuevo desde matrices. hojas: [{ nombre, filas: [[valor…]…], anchos?: [caracteres…], combinadas?: ['A1:B2'…], inmovilizarFilas?: n }]; valor: número,
//     texto, lógico, Date, null o { error }. opciones: { negritaPrimeraFila: bool, ajustarTexto: bool }. Los nombres de hoja se corrigen
//     (XlsxEscritor.nombreHoja) para que Excel los acepte.
(function () {
  const X = XmlEditar;
  const MAX_FILA = 1048576, MAX_COL = 16384, MAX_TEXTO = 32767, MAX_CAMBIOS = 5000000;

  class ErrorEscritor extends Error {}

  const colANumero = (s) => LibroXlsx.colANumero(s);
  const numeroACol = (n) => LibroXlsx.numeroACol(n);

  // ---- valores a XML de celda ------------------------------------------------------------------------------------
  // Excel lee _xHHHH_ como el carácter HHHH ("_x0041_" sería "A") y el XML convierte un \r suelto en \n: se escriben como Excel (_x005F_ y _x000D_).
  const escTexto = (t) => X.esc(t.replace(/_(x[0-9A-Fa-f]{4}_)/g, '_x005F_$1').replace(/\r/g, '_x000D_'));
  const ERRORES = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA']);
  function celdaXml(ref, estilo, valor) {
    const s = estilo ? ` s="${estilo}"` : '';
    if (valor === null || valor === undefined) return `<c r="${ref}"${s}/>`;
    if (typeof valor === 'number') {
      if (!Number.isFinite(valor)) throw new ErrorEscritor('número no finito en ' + ref);
      return `<c r="${ref}"${s}><v>${String(valor)}</v></c>`;
    }
    if (typeof valor === 'boolean') return `<c r="${ref}"${s} t="b"><v>${valor ? 1 : 0}</v></c>`;
    if (typeof valor === 'object' && valor && typeof valor.error === 'string') {
      if (!ERRORES.has(valor.error)) throw new ErrorEscritor('error desconocido en ' + ref + ': ' + valor.error);
      return `<c r="${ref}"${s} t="e"><v>${valor.error}</v></c>`;
    }
    const t = String(valor);
    if (t.length > MAX_TEXTO) throw new ErrorEscritor('texto de más de ' + MAX_TEXTO + ' caracteres en ' + ref);
    return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${escTexto(t)}</t></is></c>`;
  }

  async function texto(zip, ruta) { const f = zip.file(ruta); return f ? f.async('string') : null; }

  // La columna (desde 0) de una celda de una fila: su r="B7" o, si falta, la siguiente a la anterior.
  function columnaDe(celda, anterior) {
    const r = celda.attrs.r;
    if (r) { const m = /^([A-Za-z]{1,3})\d*$/.exec(r); if (m) return colANumero(m[1].toUpperCase()); }
    return anterior + 1;
  }

  // ---- edición en el lugar -----------------------------------------------------------------------------------------
  async function editarCeldas(libro, hoja, cambios) {
    const zip = libro.zip;
    const i = typeof hoja === 'number' ? hoja : libro.hojas.findIndex((h) => h.nombre === hoja);
    const entrada = libro.hojas[i];
    if (!entrada) throw new ErrorEscritor('hoja inexistente');
    if (entrada.tipo !== 'hoja' || !entrada.ruta) throw new ErrorEscritor('la hoja no es de celdas');
    if (!Array.isArray(cambios) || cambios.length > MAX_CAMBIOS) throw new ErrorEscritor('demasiados cambios');
    const mapa = new Map();   // fila → Map(columna → cambio)  (el último cambio de una celda gana)
    for (const c of cambios) {
      if (!Number.isInteger(c.f) || !Number.isInteger(c.c) || c.f < 0 || c.c < 0 || c.f >= MAX_FILA || c.c >= MAX_COL) throw new ErrorEscritor('celda fuera de los límites de Excel');
      if (c.estilo !== undefined && !(Number.isInteger(c.estilo) && c.estilo >= 0)) throw new ErrorEscritor('estilo que no es un índice de cellXfs');
      if (!mapa.has(c.f)) mapa.set(c.f, new Map());
      mapa.get(c.f).set(c.c, c);
    }
    if (!mapa.size) return;
    let xml = await texto(zip, entrada.ruta);
    if (xml === null) throw new ErrorEscritor('falta la parte ' + entrada.ruta);

    // encabezados de tabla: no se pueden cambiar sin tocar la tabla
    const modelo = await libro.cargarHoja(i);
    for (const t of modelo.tablas || []) {
      if (!t.encabezados) continue;
      const fil = mapa.get(t.ref.f0);
      if (fil) for (const c of fil.keys()) if (c >= t.ref.c0 && c <= t.ref.c1) throw new ErrorEscritor('la celda ' + numeroACol(c) + (t.ref.f0 + 1) + ' es un encabezado de la tabla ' + t.nombre);
    }

    // celdas dentro del rango de una fórmula matricial (no solo la maestra): cambiarlas deja la matriz rota
    {
      // los atributos en cualquier orden (ref antes que t también vale); se recorren las filas cambiadas, no las del rango (que puede ser A1:XFD1048576)
      const re = /<(?:[\w.-]+:)?f\s([^>]*)>/g, filasCambiadas = [...mapa.keys()];
      let m;
      while ((m = re.exec(xml))) {
        if (!/\bt\s*=\s*["'](?:array|dataTable)["']/.test(m[1])) continue;
        const ref = /\bref\s*=\s*["']([^"']+)["']/.exec(m[1]);
        const r = ref && LibroXlsx.decodificarRango(ref[1]);
        if (!r) continue;
        for (const f of filasCambiadas) {
          if (f < r.f0 || f > r.f1) continue;
          for (const c of mapa.get(f).keys()) if (c >= r.c0 && c <= r.c1) throw new ErrorEscritor('la celda ' + numeroACol(c) + (f + 1) + ' está dentro de una fórmula matricial o de una tabla de datos (' + ref[1] + ')');
        }
      }
    }

    // 1.ª pasada: fórmulas compartidas cuya celda maestra se va a pisar → la maestra pasa a la siguiente celda del grupo (el resto sigue valiendo)
    const promociones = [];
    {
      const raiz = X.analizar(xml, { profundidad: 3 });
      const sd = X.hijo(raiz, 'sheetData');
      if (!sd) throw new ErrorEscritor('la hoja no tiene sheetData');
      let n = 0;
      for (const fila of X.hijos(sd, 'row')) {
        n = fila.attrs.r ? parseInt(fila.attrs.r, 10) : n + 1;
        const cm = mapa.get(n - 1);
        if (!cm || fila.iniCierre < 0) continue;
        const texFila = xml.slice(fila.ini, fila.fin), rf = X.analizar(texFila, { profundidad: 3 });
        let col = -1;
        for (const celda of X.hijos(rf, 'c')) {
          col = columnaDe(celda, col);
          if (!cm.has(col)) continue;
          const f = X.hijo(celda, 'f');
          if (!f) continue;
          if (f.attrs.t === 'array' || f.attrs.t === 'dataTable') throw new ErrorEscritor('la celda ' + numeroACol(col) + n + ' es de una fórmula matricial o de una tabla de datos');
          // si va dentro de una expresión regular y de un atributo: solo dígitos (un libro hostil podría traer cualquier cosa)
          if (f.attrs.t === 'shared' && f.attrs.ref && f.fin > f.finApertura && /^\d+$/.test(f.attrs.si || '')) promociones.push({ si: f.attrs.si, ref: f.attrs.ref, formula: X.decodificar(X.interior(texFila, f)), f: n - 1, c: col, desde: fila.ini });
        }
      }
    }
    if (promociones.length) {
      const ed = [];
      for (const p of promociones) {
        // la primera otra celda del grupo que no se vaya a pisar
        const re = new RegExp('<c\\b[^>]*\\br="([A-Za-z]{1,3})(\\d+)"[^>]*>\\s*<f\\b(?=[^>]*\\bsi="' + p.si + '")(?=[^>]*\\bt="shared")[^>]*?/>', 'g');
        let m, destino = null;
        re.lastIndex = p.desde;   // las demás celdas del grupo van después de la maestra: no se recorre la hoja desde el principio por cada una
        while ((m = re.exec(xml))) {
          const f = parseInt(m[2], 10) - 1, c = colANumero(m[1].toUpperCase());
          const pisada = mapa.has(f) && mapa.get(f).has(c);
          if (!pisada) { destino = { m, f, c }; break; }
        }
        if (!destino) continue;   // todo el grupo se pisa: nadie lo necesita
        const todo = destino.m[0], pos = destino.m.index + todo.lastIndexOf('<f');
        const nueva = LibroXlsx.desplazarFormula(p.formula, destino.f - p.f, destino.c - p.c);
        ed.push({ ini: pos, fin: destino.m.index + todo.length, texto: `<f t="shared" ref="${X.escAttr(p.ref)}" si="${p.si}">${X.esc(nueva)}</f>` });
      }
      if (ed.length) xml = X.aplicar(xml, ed);
    }

    // 2.ª pasada: las filas afectadas
    const raiz = X.analizar(xml, { profundidad: 3 });
    const sd = X.hijo(raiz, 'sheetData');
    const filas = [];
    {
      let n = 0;
      for (const fila of X.hijos(sd, 'row')) { n = fila.attrs.r ? parseInt(fila.attrs.r, 10) : n + 1; filas.push({ n, nodo: fila }); }
    }
    const ed = [], nuevasFilas = new Map();   // posición de inserción → texto
    const porNumero = new Map(filas.map((x) => [x.n, x.nodo]));
    const ordenadas = filas.every((x, k) => !k || filas[k - 1].n < x.n);
    let sigFila = 0;   // con las filas en orden, la siguiente mayor avanza con f (antes: una búsqueda lineal por fila nueva, cuadrático)
    let f0 = Infinity, f1 = -1, c0 = Infinity, c1 = -1;
    const marcar = (f, c) => { if (f < f0) f0 = f; if (f > f1) f1 = f; if (c < c0) c0 = c; if (c > c1) c1 = c; };
    const refDe = (f, c) => numeroACol(c) + (f + 1);

    for (const f of [...mapa.keys()].sort((a, b) => a - b)) {
      const cm = mapa.get(f), cols = [...cm.keys()].sort((a, b) => a - b);
      cols.forEach((c) => marcar(f, c));
      const nodo = porNumero.get(f + 1);
      if (!nodo) { // fila nueva: va antes de la primera fila mayor, o al final
        let sig;
        if (ordenadas) { while (sigFila < filas.length && filas[sigFila].n <= f + 1) sigFila++; sig = filas[sigFila]; } else sig = filas.find((x) => x.n > f + 1);
        const pos = sig ? sig.nodo.ini : sd.iniCierre >= 0 ? sd.iniCierre : -1;
        const textoFila = `<row r="${f + 1}">` + cols.map((c) => celdaXml(refDe(f, c), cm.get(c).estilo, cm.get(c).valor)).join('') + '</row>';
        const clave = pos < 0 ? 'vacio' : pos;
        nuevasFilas.set(clave, (nuevasFilas.get(clave) || '') + textoFila);
        continue;
      }
      // fila existente: se reconstruye entera, copiando tal cual las celdas que no cambian
      const abrir = xml.slice(nodo.ini, nodo.finApertura);
      const interior = nodo.iniCierre < 0 ? '' : xml.slice(nodo.finApertura, nodo.iniCierre);
      const envuelto = '<r>' + interior + '</r>';
      const rf = interior ? X.analizar(envuelto, { profundidad: 3 }) : { hijos: [] };
      const piezas = [], pendientes = cols.slice();
      const nueva = (c, previa) => {
        const cam = cm.get(c);
        const s = previa && /^\d+$/.test(previa.attrs.s || '') ? previa.attrs.s : undefined;   // el de la celda, si es un índice
        return celdaXml(refDe(f, c), cam.estilo !== undefined ? cam.estilo : s, cam.valor);
      };
      const soltar = (hasta) => { while (pendientes.length && pendientes[0] < hasta) piezas.push(nueva(pendientes.shift(), null)); };
      let pos = 3, col = -1;
      for (const h of rf.hijos) {
        piezas.push(envuelto.slice(pos, h.ini));
        if (h.local !== 'c') { soltar(Infinity); piezas.push(envuelto.slice(h.ini, h.fin)); } // lo que no es celda (extLst) va después de todas
        else {
          col = columnaDe(h, col);
          soltar(col);
          if (pendientes.length && pendientes[0] === col) { pendientes.shift(); piezas.push(nueva(col, h)); } else piezas.push(envuelto.slice(h.ini, h.fin));
        }
        pos = h.fin;
      }
      soltar(Infinity);
      piezas.push(envuelto.slice(pos, envuelto.length - 4));
      const ab = X.ponerAtributo(abrir.replace(/\s*\/>$/, '>'), 'spans', null);
      ed.push({ ini: nodo.ini, fin: nodo.fin, texto: ab + piezas.join('') + '</row>' });
    }
    for (const [pos, t] of nuevasFilas) {
      if (pos === 'vacio') ed.push({ ini: sd.ini, fin: sd.fin, texto: xml.slice(sd.ini, sd.finApertura).replace(/\s*\/>$/, '>') + t + '</sheetData>' });
      else ed.push({ ini: pos, fin: pos, texto: t });
    }
    // dimension: tiene que cubrir las celdas nuevas
    const dim = X.hijo(raiz, 'dimension');
    if (dim && dim.attrs.ref) {
      const r = LibroXlsx.decodificarRango(dim.attrs.ref);
      if (r) {
        const nf0 = Math.min(r.f0, f0), nf1 = Math.max(r.f1, f1), nc0 = Math.min(r.c0, c0), nc1 = Math.max(r.c1, c1);
        if (nf0 !== r.f0 || nf1 !== r.f1 || nc0 !== r.c0 || nc1 !== r.c1) {
          ed.push({ ini: dim.ini, fin: dim.finApertura, texto: X.ponerAtributo(xml.slice(dim.ini, dim.finApertura), 'ref', numeroACol(nc0) + (nf0 + 1) + ':' + numeroACol(nc1) + (nf1 + 1)) });
        }
      }
    }
    zip.file(entrada.ruta, X.aplicar(xml, ed));
    if (libro._hojasCargadas) libro._hojasCargadas.delete(i);   // el modelo leído antes ya no vale: cargarHoja(i) vuelve a leer la hoja editada
    await XlsxPaquete.quitarCalcChain(zip);
    await XlsxPaquete.marcarRecalculo(zip);
  }

  // ---- libro nuevo --------------------------------------------------------------------------------------------------
  // Nombre de hoja que Excel acepta: sin : \ / ? * [ ], sin apóstrofo al principio ni al final, 31 caracteres como máximo, único sin distinguir mayúsculas.
  // Los 31 caracteres de Excel son unidades UTF-16 (un emoji cuenta 2; con 20 emojis Excel no abre el libro): se corta sin partir un par.
  const corte = (s, max) => (s.length <= max ? s : s.slice(0, /[\uD800-\uDBFF]/.test(s[max - 1]) ? max - 1 : max));
  function nombreHoja(nombre, usados) {
    let n = String(nombre === undefined || nombre === null ? '' : nombre).replace(/[:\\\/?*\[\]\u0000-\u001F￾￿]/g, ' ').replace(/^['\s]+|['\s]+$/g, '');
    if (!n) n = 'Hoja';
    n = corte(n, 31).replace(/^'+|'+$/g, '');
    if (!n) n = 'Hoja';
    const clave = (x) => x.toLowerCase();
    if (usados) {
      let base = n, k = 2;
      while ([...usados].some((u) => clave(u) === clave(n)) || clave(n) === 'history') {
        const suf = ' (' + k++ + ')';
        n = corte(base, 31 - suf.length) + suf;
      }
      usados.add(n);
    }
    return n;
  }

  const ESTILO = { normal: 0, negrita: 1, fecha: 2, fechaHora: 3, ajustar: 4, negritaAjustar: 5 };
  const ESTILOS_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

  const serialDeFecha = (d) => d.getTime() / 86400000 + 25569;   // base 1900 (las fechas anteriores al 1-mar-1900 no se usan)

  async function crearLibro(hojas, opciones) {
    opciones = opciones || {};
    if (!Array.isArray(hojas) || !hojas.length) throw new ErrorEscritor('un libro necesita al menos una hoja');
    if (hojas.length > 255) throw new ErrorEscritor('demasiadas hojas');
    const zip = new ZipPropio();
    const usados = new Set(), nombres = [];
    const textos = [], indiceTexto = new Map();
    let totalTextos = 0;
    const idTexto = (t) => { totalTextos++; if (!indiceTexto.has(t)) { indiceTexto.set(t, textos.length); textos.push(t); } return indiceTexto.get(t); };
    const hojasXml = hojas.map((h, k) => {
      const nombre = nombreHoja(h.nombre, usados);
      nombres.push(nombre);
      const filas = h.filas || [];
      if (filas.length > MAX_FILA) throw new ErrorEscritor('demasiadas filas');
      let maxCol = 0;
      const filasXml = [];
      filas.forEach((fila, f) => {
        if (!Array.isArray(fila) || !fila.length) return;
        if (fila.length > MAX_COL) throw new ErrorEscritor('demasiadas columnas');
        const celdas = [];
        fila.forEach((v, c) => {
          if (v === null || v === undefined || v === '') return;
          maxCol = Math.max(maxCol, c + 1);
          const ref = numeroACol(c) + (f + 1);
          const cab = f === 0 && opciones.negritaPrimeraFila;
          if (v instanceof Date) { if (Number.isNaN(v.getTime())) return; const s = serialDeFecha(v); celdas.push(celdaXml(ref, Number.isInteger(s) ? ESTILO.fecha : ESTILO.fechaHora, s)); return; }
          if (typeof v === 'string') {
            if (v.length > MAX_TEXTO) throw new ErrorEscritor('texto de más de ' + MAX_TEXTO + ' caracteres en ' + ref);
            const est = cab ? (opciones.ajustarTexto ? ESTILO.negritaAjustar : ESTILO.negrita) : (opciones.ajustarTexto && /\n/.test(v) ? ESTILO.ajustar : 0);
            celdas.push(`<c r="${ref}"${est ? ` s="${est}"` : ''} t="s"><v>${idTexto(v)}</v></c>`);
            return;
          }
          celdas.push(celdaXml(ref, cab ? ESTILO.negrita : 0, v));
        });
        if (celdas.length) filasXml.push(`<row r="${f + 1}">${celdas.join('')}</row>`);
      });
      const anchos = (h.anchos || []).slice(0, MAX_COL).map((w, c) => (w > 0 ? `<col min="${c + 1}" max="${c + 1}" width="${Math.min(255, w)}" customWidth="1"/>` : '')).join('');
      // combinadas: solo rangos A1:B2 válidos y que no se pisen con una anterior (dos que se solapan: Excel no abre el libro)
      const comb = [], rangos = [];
      for (const r of h.combinadas || []) {
        const d = typeof r === 'string' && /^\$?[A-Za-z]{1,3}\$?\d+:\$?[A-Za-z]{1,3}\$?\d+$/.test(r) ? LibroXlsx.decodificarRango(r) : null;
        if (!d || d.f1 >= MAX_FILA || d.c1 >= MAX_COL || rangos.some((o) => d.f0 <= o.f1 && o.f0 <= d.f1 && d.c0 <= o.c1 && o.c0 <= d.c1)) continue;
        rangos.push(d); comb.push(numeroACol(d.c0) + (d.f0 + 1) + ':' + numeroACol(d.c1) + (d.f1 + 1));
      }
      const inm = Math.min(MAX_FILA - 1, Math.floor(Number(h.inmovilizarFilas) || 0));   // entero y dentro de la hoja: 2.5 o 1048576 y Excel no abre el libro
      const congelar = inm > 0 ? `<sheetViews><sheetView workbookViewId="0"${k === 0 ? ' tabSelected="1"' : ''}><pane ySplit="${inm}" topLeftCell="A${inm + 1}" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` : `<sheetViews><sheetView workbookViewId="0"${k === 0 ? ' tabSelected="1"' : ''}/></sheetViews>`;
      return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + `<dimension ref="A1${maxCol && filas.length ? ':' + numeroACol(maxCol - 1) + filas.length : ''}"/>` + congelar + '<sheetFormatPr defaultRowHeight="15"/>'
        + (anchos ? `<cols>${anchos}</cols>` : '') + `<sheetData>${filasXml.join('')}</sheetData>`
        + (comb.length ? `<mergeCells count="${comb.length}">${comb.map((r) => `<mergeCell ref="${r}"/>`).join('')}</mergeCells>` : '') + '</worksheet>';
    });
    const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + hojasXml.map((_, k) => `<Override PartName="/xl/worksheets/sheet${k + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')
      + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>');
    zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
    zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><bookViews><workbookView/></bookViews><sheets>`
      + nombres.map((n, k) => `<sheet name="${X.escAttr(n)}" sheetId="${k + 1}" r:id="rId${k + 1}"/>`).join('') + '</sheets></workbook>');
    zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + hojasXml.map((_, k) => `<Relationship Id="rId${k + 1}" Type="${R}/worksheet" Target="worksheets/sheet${k + 1}.xml"/>`).join('')
      + `<Relationship Id="rId${hojasXml.length + 1}" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rId${hojasXml.length + 2}" Type="${R}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`);
    zip.file('xl/styles.xml', ESTILOS_XML);
    zip.file('xl/sharedStrings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${totalTextos}" uniqueCount="${textos.length}">`
      + textos.map((t) => `<si><t xml:space="preserve">${escTexto(t)}</t></si>`).join('') + '</sst>');
    hojasXml.forEach((x, k) => zip.file(`xl/worksheets/sheet${k + 1}.xml`, x));
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  globalThis.XlsxEscritor = { editarCeldas, crearLibro, nombreHoja, ErrorEscritor };
})();

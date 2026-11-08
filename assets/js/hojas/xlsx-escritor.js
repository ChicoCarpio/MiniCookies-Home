/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// COPIA ADAPTADA para la suite (Office-Temporal, 2026-10-25) de codigo/MiniCookies-Home/assets/js/xlsx-escritor.js. Cambio: editarCeldas también escribe
// fórmulas ({ formula, valor } → <f> con su último resultado en <v>). Ver docs/motores-adaptados.md.
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
//     Libro nuevo desde matrices. hojas: [{ nombre, filas: [[valor…]…], anchos?: [caracteres…], combinadas?: ['A1:B2'…], inmovilizarFilas?: n, autofiltro?: bool (botones de filtro sobre todo el rango),
//     estado?: 'visible' | 'oculta' | 'muyOculta', columnasOcultas?: [columna desde 0…], filasOcultas?: [fila desde 0…] }]; valor: número,
//     texto, lógico, Date, null, { error } o { valor, formato } (valor: número o Date; formato: el código de formato de número de Excel, p. ej. '0.00%', 'dd/mm/yyyy', o el número de uno integrado, p. ej. 14 (fecha corta del idioma de quien abre el libro); hasta 200 distintos por libro). opciones: { negritaPrimeraFila: bool, ajustarTexto: bool, fechasUtc: bool }. Un Date se guarda con su hora local (fechasUtc: con la UTC). Los nombres de hoja se corrigen
//     (XlsxEscritor.nombreHoja) para que Excel los acepte.
//     Estilo de celda (lo que llevan los .xls, .xlsb y .ods al pasarlos a .xlsx): { valor, formato?, estilo } con estilo = { fuente?: { nombre, tam, negrita, cursiva, tachado, subrayado: true | 'double', color: 'RRGGBB' },
//     relleno?: 'RRGGBB' | { patron: 'gray125' | 'lightGrid' | … (los de ECMA-376 18.18.55), fg?: 'RRGGBB', bg?: 'RRGGBB' }, bordes?: { izq, der, arr, aba, diagA, diagB: { e: 'thin' | 'medium' | …, c: 'RRGGBB' } }, alinea?: { h, v, ajustar, sangria, rotacion } }; una celda vacía con estilo (valor null) se escribe.
//     Una hoja puede traer altosFila: { fila desde 0: puntos } (alto propio de cada fila).
//     Fórmulas: un valor { formula: 'SUM(A1:A3)', valor?: …, formato? } escribe la fórmula (sin '=') con su último resultado guardado (valor: lo mismo que una celda
//     normal, incluido { valor, formato }, o nada) y pide a Excel recalcular al abrir. Para pasar fórmulas de otro libro sin romperlas: XlsxEscritor.formulaConservable(formula,
//     nombresDeHojas) dice si se puede llevar tal cual (sin nombres definidos, tablas ni libros externos, y solo con hojas que existen) y renombrarHojas(formula, mapa) cambia
//     los nombres de hoja que la fórmula cita.
(function () {
  const X = XmlEditar;
  const MAX_FILA = 1048576, MAX_COL = 16384, MAX_TEXTO = 32767, MAX_CAMBIOS = 5000000;

  class ErrorEscritor extends Error {}

  const colANumero = (s) => LibroXlsx.colANumero(s);
  const numeroACol = (n) => LibroXlsx.numeroACol(n);

  // ---- valores a XML de celda ------------------------------------------------------------------------------------
  // Excel lee _xHHHH_ como el carácter HHHH ("_x0041_" sería "A") y el XML convierte un \r suelto en \n: se escriben como Excel (_x005F_ y _x000D_).
  const escTexto = (t) => X.esc(t.replace(/_(x[0-9A-Fa-f]{4}_)/g, '_x005F_$1').replace(/\r/g, '_x000D_'));
  const ERRORES = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!']);
  // Los únicos errores que Excel acepta dentro de <v t="e">: con #SPILL! o #CALC! no abre el libro (probado con Excel, oráculo de 2026-10-25).
  // Una fórmula con ese error se guarda sin resultado (Excel la calcula al abrir y vuelve a dar #SPILL!); un valor suelto se guarda como #VALUE!.
  const ERRORES_EN_ARCHIVO = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA']);
  function celdaXml(ref, estilo, valor) {
    let s = estilo ? ` s="${estilo}"` : '';
    if (valor === null || valor === undefined) return `<c r="${ref}"${s}/>`;
    // fórmula con su último resultado (suite): <c t="str|b|e"><f>…</f><v>…</v></c>; sin resultado conocido, solo la fórmula (Excel la calcula al abrir)
    if (typeof valor === 'object' && valor && typeof valor.formula === 'string') {
      // matriz (dinámica si trae cm, de Excel 365: el ancla lleva el rango que desborda; las otras celdas, solo su valor)
      const enMatriz = typeof valor.matriz === 'string' && /^[A-Za-z]{1,3}\d+(:[A-Za-z]{1,3}\d+)?$/.test(valor.matriz);
      const f = (enMatriz ? `<f t="array" ref="${valor.matriz.toUpperCase()}">` : '<f>') + X.esc(valor.formula.replace(/^=/, '')) + '</f>';
      if (enMatriz && Number.isInteger(valor.cm) && valor.cm > 0) s += ` cm="${valor.cm}"`;
      const r = valor.valor;
      if (r === null || r === undefined) return `<c r="${ref}"${s}>${f}</c>`;
      if (typeof r === 'number') return Number.isFinite(r) ? `<c r="${ref}"${s}>${f}<v>${String(r)}</v></c>` : `<c r="${ref}"${s} t="e">${f}<v>#NUM!</v></c>`;
      if (typeof r === 'boolean') return `<c r="${ref}"${s} t="b">${f}<v>${r ? 1 : 0}</v></c>`;
      if (typeof r === 'object' && typeof r.error === 'string') return ERRORES_EN_ARCHIVO.has(r.error) ? `<c r="${ref}"${s} t="e">${f}<v>${r.error}</v></c>` : `<c r="${ref}"${s}>${f}</c>`;
      const t = String(r);
      return `<c r="${ref}"${s} t="str">${f}<v>${escTexto(t.slice(0, MAX_TEXTO))}</v></c>`;
    }
    if (typeof valor === 'number') {
      if (!Number.isFinite(valor)) throw new ErrorEscritor('número no finito en ' + ref);
      return `<c r="${ref}"${s}><v>${String(valor)}</v></c>`;
    }
    if (typeof valor === 'boolean') return `<c r="${ref}"${s} t="b"><v>${valor ? 1 : 0}</v></c>`;
    if (typeof valor === 'object' && valor && typeof valor.error === 'string') {
      if (!ERRORES.has(valor.error)) throw new ErrorEscritor('error desconocido en ' + ref + ': ' + valor.error);
      return `<c r="${ref}"${s} t="e"><v>${ERRORES_EN_ARCHIVO.has(valor.error) ? valor.error : '#VALUE!'}</v></c>`;
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
  // Predicado (f, c) → true para las celdas que editarCeldas rechazaría: encabezados de tabla y celdas de fórmulas matriciales o tablas de datos.
  // `hoja` es la hoja ya leída (await libro.cargarHoja(i)). Sirve a las herramientas que quieren saltarlas en vez de fallar.
  function celdasProtegidas(hoja) {
    const rangos = [];
    for (const t of hoja.tablas || []) {
      if (t.encabezados) rangos.push({ f0: t.ref.f0, f1: t.ref.f0, c0: t.ref.c0, c1: t.ref.c1 });
      if (t.totales) rangos.push({ f0: t.ref.f1, f1: t.ref.f1, c0: t.ref.c0, c1: t.ref.c1 });   // la fila de totales: su etiqueta también está en table.xml
    }
    for (const fx of hoja.celdas.formulas.values()) {
      if ((fx.tipo === 'array' || fx.tipo === 'dataTable') && fx.ref && !fx.dinamica) { const r = LibroXlsx.decodificarRango(fx.ref); if (r) rangos.push(r); }   // las dinámicas (Excel 365) se reescriben enteras
    }
    // por fila: con miles de fórmulas matriciales de una celda, recorrer todos los rangos por cada celda consultada era cuadrático
    const porFila = new Map(), altos = [];
    for (const r of rangos) {
      if (r.f1 - r.f0 >= 64) { altos.push(r); continue; }
      for (let f = r.f0; f <= r.f1; f++) { if (!porFila.has(f)) porFila.set(f, []); porFila.get(f).push(r); }
    }
    const dentro = (f, c) => (r) => f >= r.f0 && f <= r.f1 && c >= r.c0 && c <= r.c1;
    return (f, c) => { const l = porFila.get(f); return (!!l && l.some(dentro(f, c))) || altos.some(dentro(f, c)); };
  }

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
      if (t.encabezados) {
        const fil = mapa.get(t.ref.f0);
        if (fil) for (const c of fil.keys()) if (c >= t.ref.c0 && c <= t.ref.c1) throw new ErrorEscritor('la celda ' + numeroACol(c) + (t.ref.f0 + 1) + ' es un encabezado de la tabla ' + t.nombre);
      }
      if (t.totales) {
        const fil = mapa.get(t.ref.f1);
        if (fil) for (const c of fil.keys()) if (c >= t.ref.c0 && c <= t.ref.c1) throw new ErrorEscritor('la celda ' + numeroACol(c) + (t.ref.f1 + 1) + ' es de la fila de totales de la tabla ' + t.nombre);
      }
    }

    // celdas dentro del rango de una fórmula matricial (no solo la maestra): cambiarlas deja la matriz rota
    {
      // los atributos en cualquier orden (ref antes que t también vale); se recorren las filas cambiadas, no las del rango (que puede ser A1:XFD1048576)
      const re = /<(?:[\w.-]+:)?f\s([^>]*)>/g, filasCambiadas = [...mapa.keys()];
      let m;
      while ((m = re.exec(xml))) {
        if (!/\bt\s*=\s*["'](?:array|dataTable)["']/.test(m[1])) continue;
        if (/\bcm\s*=\s*["']\d+["']/.test(xml.slice(xml.lastIndexOf('<c', m.index), m.index))) continue;   // matriz dinámica: sus celdas son solo valores
        const ref = /\bref\s*=\s*["']([^"']+)["']/.exec(m[1]);
        const r = ref && LibroXlsx.decodificarRango(ref[1]);
        if (!r) continue;
        // mover una matricial: se escribe entera, o solo cambia el texto de su maestra (lo que apunta a celdas que se movieron)
        if (/\bt\s*=\s*["']array["']/.test(m[1])) {
          const maestra = mapa.get(r.f0) && mapa.get(r.f0).get(r.c0);
          if (maestra && maestra.valor && maestra.valor.fija) continue;
          let entera = (r.f1 - r.f0 + 1) * (r.c1 - r.c0 + 1) <= 100000;
          for (let f = r.f0; entera && f <= r.f1; f++) for (let c = r.c0; entera && c <= r.c1; c++) entera = !!(mapa.get(f) && mapa.get(f).has(c));
          if (entera) continue;
        }
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
          if ((f.attrs.t === 'array' && !celda.attrs.cm && !f.attrs.ref) || f.attrs.t === 'dataTable') throw new ErrorEscritor('la celda ' + numeroACol(col) + n + ' es de una fórmula matricial o de una tabla de datos');
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

  const MAX_FORMATOS = 200;
  const MAX_XFS = 60000;       // Excel guarda hasta 64.000 formatos de celda
  const MAX_ESTILOS_PIEZA = 20000;   // fuentes, rellenos o bordes distintos
  const ESTILOS_BORDE = new Set(['thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']);
  const ALIN_H = new Set(['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed']);
  const ALIN_V = new Set(['top', 'center', 'bottom', 'justify', 'distributed']);
  const PATRONES = new Set(['solid', 'darkDown', 'darkGray', 'darkGrid', 'darkHorizontal', 'darkTrellis', 'darkUp', 'darkVertical', 'gray0625', 'gray125', 'lightDown', 'lightGray', 'lightGrid', 'lightHorizontal', 'lightTrellis', 'lightUp', 'lightVertical', 'mediumGray']);
  const esColor = (c) => typeof c === 'string' && /^[0-9A-Fa-f]{6}$/.test(c);
  // Deja un estilo pedido en su forma canónica (para repartir los iguales) y descarta lo que Excel no aceptaría; null si no queda nada.
  // Los lectores comparten un objeto por estilo de celda: un millón de celdas no normaliza ni serializa un millón de veces.
  const cacheNormal = new WeakMap(), clavesEstilo = new WeakMap();
  const claveEstilo = (e) => { let k = clavesEstilo.get(e); if (k === undefined) { k = JSON.stringify(e); clavesEstilo.set(e, k); } return k; };
  function normalizarEstilo(e) {
    if (!e || typeof e !== 'object') return null;
    if (cacheNormal.has(e)) return cacheNormal.get(e);
    const r = normalizarEstiloSinCache(e);
    cacheNormal.set(e, r);
    return r;
  }
  function normalizarEstiloSinCache(e) {
    const r = {};
    const f = e.fuente;
    if (f && typeof f === 'object') {
      const o = {};
      if (typeof f.nombre === 'string' && f.nombre && f.nombre.length <= 31 && !/[\u0000-\u001F]/.test(f.nombre)) o.nombre = f.nombre;
      if (Number.isFinite(f.tam) && f.tam >= 1 && f.tam <= 409) o.tam = Math.round(f.tam * 100) / 100;
      if (f.negrita === true) o.negrita = true;
      if (f.cursiva === true) o.cursiva = true;
      if (f.tachado === true) o.tachado = true;
      if (f.subrayado === true || f.subrayado === 'double') o.subrayado = f.subrayado;
      if (esColor(f.color)) o.color = f.color.toUpperCase();
      if (Object.keys(o).length) r.fuente = o;
    }
    if (esColor(e.relleno)) r.relleno = e.relleno.toUpperCase();
    else if (e.relleno && typeof e.relleno === 'object' && PATRONES.has(e.relleno.patron) && (esColor(e.relleno.fg) || esColor(e.relleno.bg))) {   // trama: { patron, fg, bg }
      r.relleno = e.relleno.patron === 'solid' && esColor(e.relleno.fg) ? e.relleno.fg.toUpperCase()
        : { patron: e.relleno.patron, ...(esColor(e.relleno.fg) ? { fg: e.relleno.fg.toUpperCase() } : {}), ...(esColor(e.relleno.bg) ? { bg: e.relleno.bg.toUpperCase() } : {}) };
    }
    if (e.bordes && typeof e.bordes === 'object') {
      const o = {};
      for (const lado of ['izq', 'der', 'arr', 'aba', 'diagA', 'diagB']) {
        const b = e.bordes[lado];
        if (b && typeof b === 'object' && ESTILOS_BORDE.has(b.e)) o[lado] = esColor(b.c) ? { e: b.e, c: b.c.toUpperCase() } : { e: b.e };
      }
      if (Object.keys(o).length) r.bordes = o;
    }
    const a = e.alinea;
    if (a && typeof a === 'object') {
      const o = {};
      if (ALIN_H.has(a.h) && a.h !== 'general') o.h = a.h;
      if (ALIN_V.has(a.v) && a.v !== 'bottom') o.v = a.v;
      if (a.ajustar === true) o.ajustar = true;
      if (Number.isInteger(a.sangria) && a.sangria > 0) o.sangria = Math.min(250, a.sangria);
      if (Number.isInteger(a.rotacion) && ((a.rotacion > 0 && a.rotacion <= 180) || a.rotacion === 255)) o.rotacion = a.rotacion;
      if (Object.keys(o).length) r.alinea = o;
    }
    return Object.keys(r).length ? r : null;
  }
  const colorXml = (c, etiqueta = 'color') => (c ? `<${etiqueta} rgb="FF${c}"/>` : `<${etiqueta} indexed="64"/>`);
  // Los estilos base más un cellXfs por cada combinación de formato de número y estilo pedida (ids de formato propios desde 164).
  // combos: [{ codigo, estilo }] en el orden de sus índices (el primero es el 6).
  function estilosConFormatos(combos) {
    if (!combos.length) return ESTILOS_XML;
    const cadenas = [...new Set(combos.map((c) => c.codigo).filter((c) => typeof c === 'string'))];
    const idDe = (c) => (c === undefined ? 0 : typeof c === 'number' ? c : 164 + cadenas.indexOf(c));   // un número es un formato integrado de Excel (14, 22…): se usa su id
    const numFmts = cadenas.length ? `<numFmts count="${cadenas.length}">${cadenas.map((c, k) => `<numFmt numFmtId="${164 + k}" formatCode="${X.escAttr(c)}"/>`).join('')}</numFmts>` : '';
    const fuentes = ['<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>', '<font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>'];
    const rellenos = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    const bordes = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    // por un índice y no con indexOf: 60.000 xfs con 20.000 fuentes distintas eran 1.200 millones de comparaciones (35 s)
    const indices = new Map([fuentes, rellenos, bordes].map((l) => [l, new Map(l.map((x, i) => [x, i]))]));
    const idPieza = (lista, xml) => { const m = indices.get(lista); let i = m.get(xml); if (i === undefined) { if (lista.length >= MAX_ESTILOS_PIEZA) return 0; i = lista.push(xml) - 1; m.set(xml, i); } return i; };
    const xfs = combos.map(({ codigo, estilo }) => {
      let fontId = 0, fillId = 0, borderId = 0, alin = '';
      const e = estilo || {};
      if (e.fuente) {
        const f = e.fuente;
        fontId = idPieza(fuentes, '<font>' + (f.negrita ? '<b/>' : '') + (f.cursiva ? '<i/>' : '') + (f.tachado ? '<strike/>' : '') + (f.subrayado ? (f.subrayado === 'double' ? '<u val="double"/>' : '<u/>') : '')
          + `<sz val="${f.tam || 11}"/>` + (f.color ? colorXml(f.color) : '') + `<name val="${X.escAttr(f.nombre || 'Calibri')}"/><family val="2"/></font>`);
      }
      if (e.relleno) {
        const r = typeof e.relleno === 'string' ? { patron: 'solid', fg: e.relleno } : e.relleno;
        fillId = idPieza(rellenos, `<fill><patternFill patternType="${r.patron}">` + (r.fg ? `<fgColor rgb="FF${r.fg}"/>` : '') + (r.bg ? `<bgColor rgb="FF${r.bg}"/>` : r.patron === 'solid' ? '<bgColor indexed="64"/>' : '') + '</patternFill></fill>');
      }
      if (e.bordes) {
        const lado = (nombre, k) => { const b = e.bordes[k]; return b ? `<${nombre} style="${b.e}">${colorXml(b.c)}</${nombre}>` : `<${nombre}/>`; };
        const diag = e.bordes.diagA || e.bordes.diagB;
        borderId = idPieza(bordes, `<border${e.bordes.diagA ? ' diagonalDown="1"' : ''}${e.bordes.diagB ? ' diagonalUp="1"' : ''}>` + lado('left', 'izq') + lado('right', 'der') + lado('top', 'arr') + lado('bottom', 'aba')
          + (diag ? `<diagonal style="${diag.e}">${colorXml(diag.c)}</diagonal>` : '<diagonal/>') + '</border>');
      }
      if (e.alinea) {
        const a = e.alinea;
        alin = '<alignment' + (a.h ? ` horizontal="${a.h}"` : '') + (a.v ? ` vertical="${a.v}"` : '') + (a.ajustar ? ' wrapText="1"' : '') + (a.sangria ? ` indent="${a.sangria}"` : '') + (a.rotacion ? ` textRotation="${a.rotacion}"` : '') + '/>';
      }
      const atr = `numFmtId="${idDe(codigo)}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"` + (codigo !== undefined ? ' applyNumberFormat="1"' : '') + (fontId ? ' applyFont="1"' : '') + (fillId ? ' applyFill="1"' : '')
        + (borderId ? ' applyBorder="1"' : '') + (alin ? ' applyAlignment="1"' : '');
      return alin ? `<xf ${atr}>${alin}</xf>` : `<xf ${atr}/>`;
    }).join('');
    // con funciones: un formato como "$"#,##0 lleva $& o $$, que replace con una cadena interpretaría
    return ESTILOS_XML.replace(/<fonts [\s\S]*<\/borders>/, () => numFmts + `<fonts count="${fuentes.length}">${fuentes.join('')}</fonts><fills count="${rellenos.length}">${rellenos.join('')}</fills><borders count="${bordes.length}">${bordes.join('')}</borders>`)
      .replace('<cellXfs count="6">', () => `<cellXfs count="${6 + combos.length}">`).replace('</cellXfs>', () => xfs + '</cellXfs>');
  }

  // Un Date de JavaScript es un instante; Excel guarda la hora "de pared". Por defecto se toma la del reloj local del Date (new Date(2024, 2, 15) → 15/3/2024
  // sin hora, en cualquier zona); con opciones.fechasUtc se toma la hora UTC (new Date(Date.UTC(2024, 2, 15))).
  // Base 1900 (las fechas anteriores al 1-mar-1900 no se usan). Se redondea al milisegundo para que 15/3/2024 00:00 dé un entero exacto.
  const serialDeFecha = (d, utc) => {
    const ms = utc ? d.getTime() : Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds(), d.getMilliseconds());
    return Math.round(ms / 86400000 * 1e9) / 1e9 + 25569;
  };
  const MAX_CARACTERES = 400000000;   // el límite de una cadena en V8 es ~536 millones de caracteres

  // ---- fórmulas ---------------------------------------------------------------------------------------------------
  const MAX_FORMULA = 8192;
  // <c> de una celda con fórmula y su último resultado; null si la fórmula no se puede escribir.
  function formulaXml(ref, v, estiloDe) {
    const f = v.formula.replace(/^=/, '');
    if (!f || f.length > MAX_FORMULA || /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(f)) return null;
    // el resultado puede traer su formato (y estilo) dentro (valor: { valor, formato }) o la celda al lado de la fórmula ({ formula, valor, formato, estilo })
    let x = v.valor, formato = v.formato, estilo = v.estilo;
    if (x && typeof x === 'object' && !(x instanceof Date) && 'formato' in x && 'valor' in x) { if (formato === undefined) formato = x.formato; if (estilo === undefined) estilo = x.estilo; x = x.valor; }
    const est = formato !== undefined || estilo ? estiloDe(formato, normalizarEstilo(estilo)) : 0;
    const s = est ? ` s="${est}"` : '';
    // fórmula de matriz (CSE): el rango va en la primera celda; las demás guardan solo su valor
    const enMatriz = typeof v.matriz === 'string' && /^[A-Za-z]{1,3}\d+(:[A-Za-z]{1,3}\d+)?$/.test(v.matriz);
    const ft = enMatriz ? `<f t="array" ref="${v.matriz.toUpperCase()}">${X.esc(f)}</f>` : `<f>${X.esc(f)}</f>`;
    // matriz dinámica de Excel 365: cm apunta a la metadata XLDAPR de xl/metadata.xml
    const cm = enMatriz && Number.isInteger(v.cm) && v.cm > 0 ? ` cm="${v.cm}"` : '';
    if (x instanceof Date) { if (Number.isNaN(x.getTime())) x = null; else { x = serialDeFecha(x, false); if (!est) return `<c r="${ref}"${cm} s="${Number.isInteger(x) ? ESTILO.fecha : ESTILO.fechaHora}">${ft}<v>${x}</v></c>`; } }
    if (typeof x === 'number') return Number.isFinite(x) ? `<c r="${ref}"${cm}${s}>${ft}<v>${String(x)}</v></c>` : `<c r="${ref}"${cm}${s}>${ft}</c>`;
    if (typeof x === 'boolean') return `<c r="${ref}"${cm}${s} t="b">${ft}<v>${x ? 1 : 0}</v></c>`;
    if (x && typeof x === 'object' && typeof x.error === 'string' && ERRORES_EN_ARCHIVO.has(x.error)) return `<c r="${ref}"${cm}${s} t="e">${ft}<v>${x.error}</v></c>`;
    if (typeof x === 'string' && x.length <= MAX_TEXTO) return `<c r="${ref}"${cm}${s} t="str">${ft}<v>${escTexto(x)}</v></c>`;
    return `<c r="${ref}"${cm}${s}>${ft}</c>`;
  }

  // Nombres definidos: [{ nombre, formula, hoja? }] (hoja: índice de la hoja si el nombre es de una sola hoja). Los de nombre o fórmula raros se omiten.
  function nombresDefinidos(lista, cantidadHojas) {
    if (!Array.isArray(lista) || !lista.length) return '';
    let xml = '';
    for (const n of lista.slice(0, 5000)) {
      if (!n || typeof n.nombre !== 'string' || typeof n.formula !== 'string') continue;
      if (!/^[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*$/.test(n.nombre) || n.nombre.length > 255 || /^[A-Za-z]{1,3}\d+$/.test(n.nombre) || /^[Rr]\d*[Cc]\d*$/.test(n.nombre)) continue;
      const f = n.formula.replace(/^=/, '');
      if (!f || f.length > MAX_FORMULA || /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(f)) continue;
      const local = Number.isInteger(n.hoja) && n.hoja >= 0 && n.hoja < cantidadHojas ? ` localSheetId="${n.hoja}"` : '';
      xml += `<definedName name="${X.escAttr(n.nombre)}"${local}>${X.esc(f)}</definedName>`;
    }
    return xml ? `<definedNames>${xml}</definedNames>` : '';
  }

  // ¿Se puede llevar esta fórmula a otro libro sin que cambie de sentido? Solo si cita hojas que existen en el libro nuevo (`nombresHojas`: Set en minúsculas) y no usa
  // nombres definidos, referencias de tabla ni otros libros: esos no viajan con las celdas.
  const PREFIJO_HOJA = /('(?:[^']|'')*'|[A-Za-z_\u00C0-\uFFFF][A-Za-z0-9_.\u00C0-\uFFFF]*)!/g;
  function formulaConservable(formula, nombresHojas) {
    if (typeof formula !== 'string') return false;
    let t = formula.replace(/^=/, '');
    if (!t || t.length > MAX_FORMULA) return false;
    t = t.replace(/"(?:[^"]|"")*"/g, '""').replace(/#(?:REF!|N\/A|DIV\/0!|NAME\?|NULL!|NUM!|VALUE!)/g, ' ');   // los errores literales no son hojas
    if (t.includes('[')) return false;
    let ok = true;
    t = t.replace(PREFIJO_HOJA, (m, p) => {
      const nombre = (p[0] === "'" ? p.slice(1, -1).replace(/''/g, "'") : p).toLowerCase();
      if (!nombresHojas.has(nombre)) ok = false;
      return ' ';
    });
    if (!ok) return false;
    const re = /[A-Za-z_\u00C0-\uFFFF][A-Za-z0-9_.\u00C0-\uFFFF]*/g;
    let m;
    while ((m = re.exec(t))) {
      const sig = t[m.index + m[0].length];
      if (sig === '(') continue;   // función
      if (/^(?:TRUE|FALSE)$/i.test(m[0]) || /[\d.]/.test(t[m.index - 1] || '')) continue;   // lógico, o el exponente de un número (1E+5)
      // celda (A1, A$1) o columna entera (A:C, $A:$C) hasta XFD. Unas letras solas o una "celda" más allá de XFD son un nombre definido: Excel
      // acepta IVA, tax, x o XFE1 como nombres (medido)
      const col = /^([A-Za-z]{1,3})(\d*)$/.exec(m[0]);
      if (col && colANumero(col[1].toUpperCase()) <= 16383) {
        const fin = m.index + m[0].length, ant = t[m.index - 1] === '$' ? t[m.index - 2] : t[m.index - 1];
        if (col[2] || sig === ':' || ant === ':' || /^\$\d/.test(t.slice(fin, fin + 2))) continue;
      }
      return false;                // un nombre definido
    }
    return true;
  }

  // Cambia los nombres de hoja que cita una fórmula. mapa: Map (nombre en minúsculas → nombre nuevo).
  // también las dos puntas de una referencia 3D: Hoja1:Hoja3!B1 y 'Hoja 1:Hoja 3'!B1
  function renombrarHojas(formula, mapa) {
    const nuevo = (n) => { const k = n.toLowerCase(); return mapa.has(k) ? mapa.get(k) : null; };
    const cita = (n) => "'" + n.replace(/'/g, "''") + "'!";
    return formula.replace(/"(?:[^"]|"")*"|#REF!|('(?:[^']|'')*'|[A-Za-z_\u00C0-\uFFFF][A-Za-z0-9_.\u00C0-\uFFFF]*)(?::([A-Za-z_\u00C0-\uFFFF][A-Za-z0-9_.\u00C0-\uFFFF]*))?!/g, (m, p, q) => {
      if (!p) return m;
      let a = p[0] === "'" ? p.slice(1, -1).replace(/''/g, "'") : p, b = q || null;
      if (b === null && p[0] === "'" && a.indexOf(':') > 0) { b = a.slice(a.indexOf(':') + 1); a = a.slice(0, a.indexOf(':')); }   // un nombre de hoja no lleva «:»
      if (b === null) { const n = nuevo(a); return n === null ? m : cita(n); }
      const na = nuevo(a), nb = nuevo(b);
      return na === null && nb === null ? m : cita((na === null ? a : na) + ':' + (nb === null ? b : nb));
    });
  }

  // Rechaza con ErrorEscritor lo que no cabe en memoria en vez de dejar escapar un RangeError ("Invalid string length").
  async function crearLibro(hojas, opciones) {
    if (Array.isArray(hojas)) {
      let est = 0;
      for (const h of hojas) for (const fila of (h && h.filas) || []) { if (!Array.isArray(fila)) continue; est += fila.length * 48; for (const v of fila) if (typeof v === 'string') est += v.length * 2; }
      if (est > MAX_CARACTERES) throw new ErrorEscritor('el libro es demasiado grande para armarlo en memoria');
    }
    try { return await crearLibroInterno(hojas, opciones); } catch (e) {
      if (e instanceof RangeError) throw new ErrorEscritor('el libro es demasiado grande para armarlo en memoria');
      throw e;
    }
  }

  async function crearLibroInterno(hojas, opciones) {
    opciones = opciones || {};
    if (!Array.isArray(hojas) || !hojas.length) throw new ErrorEscritor('un libro necesita al menos una hoja');
    if (hojas.length > 1000) throw new ErrorEscritor('demasiadas hojas');   // Excel abre 1000 sin repararlo (medido 2026-10-15); no tiene un tope propio
    const zip = new ZipPropio();
    const usados = new Set(), nombres = [];
    const textos = [], indiceTexto = new Map();
    let totalTextos = 0, hayFormulas = false;
    const combos = [], indiceCombo = new Map(), codigosUsados = new Set();   // combinaciones de formato de número y estilo → índice de cellXfs (se añaden después de los de ESTILO)
    // codigo: formato de número (código de Excel o número integrado) o undefined; estilo: ya normalizado o null. Devuelve el índice de cellXfs (0 si no se puede).
    const estiloDe = (codigo, estilo) => {
      if (typeof codigo === 'number') { if (!Number.isInteger(codigo) || codigo < 1 || codigo > 163) codigo = undefined; }
      else if (typeof codigo !== 'string' || !codigo || codigo.length > 255 || /[\u0000-\u001F]/.test(codigo)) codigo = undefined;
      if (codigo === undefined && !estilo) return 0;
      const clave = (codigo === undefined ? '' : typeof codigo === 'number' ? '#' + codigo : '"' + codigo) + '\u0001' + (estilo ? claveEstilo(estilo) : '');
      if (indiceCombo.has(clave)) return indiceCombo.get(clave);
      if (typeof codigo === 'string' && !codigosUsados.has(codigo)) { if (codigosUsados.size >= MAX_FORMATOS) codigo = undefined; else codigosUsados.add(codigo); }
      if (combos.length >= MAX_XFS || (codigo === undefined && !estilo)) return 0;
      combos.push({ codigo, estilo }); indiceCombo.set(clave, 5 + combos.length);
      return 5 + combos.length;
    };
    const estiloDeFormato = (codigo) => estiloDe(codigo, null);
    // hojas ocultas: Excel necesita al menos una visible (y la pestaña activa tiene que ser una visible)
    const estados = hojas.map((h) => (h && (h.estado === 'oculta' || h.estado === 'muyOculta') ? h.estado : 'visible'));
    const primeraVisible = Math.max(0, estados.indexOf('visible'));
    if (!estados.includes('visible')) estados[0] = 'visible';
    const idTexto = (t) => { totalTextos++; if (!indiceTexto.has(t)) { indiceTexto.set(t, textos.length); textos.push(t); } return indiceTexto.get(t); };
    const hojasXml = hojas.map((h, k) => {
      const nombre = nombreHoja(h.nombre, usados);
      nombres.push(nombre);
      const filas = h.filas || [];
      if (filas.length > MAX_FILA) throw new ErrorEscritor('demasiadas filas');
      let maxCol = 0;
      const filasXml = [];   // [fila, xml]
      const filasOcultas = new Set((Array.isArray(h.filasOcultas) ? h.filasOcultas : []).filter((f) => Number.isInteger(f) && f >= 0 && f < MAX_FILA));
      // alto propio de cada fila (puntos; Excel admite hasta 409,5)
      const altosFila = new Map();
      const altosPedidos = h.altosFila && typeof h.altosFila === 'object' ? h.altosFila : {};
      for (const k of Object.keys(altosPedidos).slice(0, MAX_FILA)) { const fila = Number(k), pt = Number(altosPedidos[k]); if (Number.isInteger(fila) && fila >= 0 && fila < MAX_FILA && pt > 0) altosFila.set(fila, Math.min(409.5, Math.round(pt * 100) / 100)); }
      const altoFila = (f) => (altosFila.has(f) ? ` ht="${altosFila.get(f)}" customHeight="1"` : '');
      filas.forEach((fila, f) => {
        if (!Array.isArray(fila) || !fila.length) return;
        if (fila.length > MAX_COL) throw new ErrorEscritor('demasiadas columnas');
        const celdas = [];
        fila.forEach((v, c) => {
          if (v === null || v === undefined || v === '') return;
          maxCol = Math.max(maxCol, c + 1);
          const ref = numeroACol(c) + (f + 1);
          const cab = f === 0 && opciones.negritaPrimeraFila;
          if (v && typeof v === 'object' && !(v instanceof Date) && typeof v.formula === 'string') {
            const x = formulaXml(ref, v, estiloDe);
            if (x) { celdas.push(x); hayFormulas = true; return; }
            v = v.valor;   // fórmula que no se puede escribir (vacía, enorme o con caracteres de control): queda su valor
            if (v === null || v === undefined || v === '') return;
          }
          if (v && typeof v === 'object' && !(v instanceof Date) && 'valor' in v && ('formato' in v || 'estilo' in v)) {
            const x = v.valor, norm = normalizarEstilo(v.estilo);
            if (x instanceof Date) {
              if (Number.isNaN(x.getTime())) return;
              const n = serialDeFecha(x, opciones.fechasUtc);
              celdas.push(celdaXml(ref, estiloDe(v.formato !== undefined && v.formato !== null ? v.formato : norm ? (Number.isInteger(n) ? 14 : 22) : undefined, norm), n));
              return;
            }
            const est = estiloDe(v.formato, norm);
            if (typeof x === 'number') { celdas.push(celdaXml(ref, est, x)); return; }
            if (norm && est) {   // un texto, un lógico, un error o una celda vacía con estilo
              if (typeof x === 'string' && x !== '') {
                if (x.length > MAX_TEXTO) throw new ErrorEscritor('texto de más de ' + MAX_TEXTO + ' caracteres en ' + ref);
                celdas.push(`<c r="${ref}" s="${est}" t="s"><v>${idTexto(x)}</v></c>`);
              } else celdas.push(celdaXml(ref, est, x === '' || x === undefined ? null : x));
              return;
            }
            v = x;   // un formato sobre un texto, un lógico o nada no cambia nada
            if (v === null || v === undefined || v === '') return;
          }
          if (v instanceof Date) { if (Number.isNaN(v.getTime())) return; const s = serialDeFecha(v, opciones.fechasUtc); celdas.push(celdaXml(ref, Number.isInteger(s) ? ESTILO.fecha : ESTILO.fechaHora, s)); return; }
          if (typeof v === 'string') {
            if (v.length > MAX_TEXTO) throw new ErrorEscritor('texto de más de ' + MAX_TEXTO + ' caracteres en ' + ref);
            const est = cab ? (opciones.ajustarTexto ? ESTILO.negritaAjustar : ESTILO.negrita) : (opciones.ajustarTexto && /\n/.test(v) ? ESTILO.ajustar : 0);
            celdas.push(`<c r="${ref}"${est ? ` s="${est}"` : ''} t="s"><v>${idTexto(v)}</v></c>`);
            return;
          }
          celdas.push(celdaXml(ref, cab ? ESTILO.negrita : 0, v));
        });
        if (celdas.length) filasXml.push([f, `<row r="${f + 1}"${altoFila(f)}${filasOcultas.has(f) ? ' hidden="1"' : ''}>${celdas.join('')}</row>`]);
      });
      // una fila oculta sin celdas también se escribe (si no, reaparecería): se inserta en su lugar
      if (filasOcultas.size || altosFila.size) {
        const hay = new Set(filasXml.map((x) => x[0]));
        for (const f of new Set([...filasOcultas, ...altosFila.keys()])) if (!hay.has(f)) filasXml.push([f, `<row r="${f + 1}"${altoFila(f)}${filasOcultas.has(f) ? ' hidden="1"' : ''}/>`]);
        filasXml.sort((a, b) => a[0] - b[0]);
      }
      const colsOcultas = new Set((Array.isArray(h.columnasOcultas) ? h.columnasOcultas : []).filter((c) => Number.isInteger(c) && c >= 0 && c < MAX_COL));
      const anchosHoja = (h.anchos || []).slice(0, MAX_COL);
      let anchos = '';
      for (let c = 0, n = Math.max(anchosHoja.length, colsOcultas.size ? Math.max(...colsOcultas) + 1 : 0); c < n; c++) {
        const w = anchosHoja[c], oculta = colsOcultas.has(c);
        if (w > 0 || oculta) anchos += `<col min="${c + 1}" max="${c + 1}" width="${Math.min(255, w > 0 ? w : 8.43)}"${oculta ? ' hidden="1"' : ''} customWidth="1"/>`;
      }
      // combinadas: solo rangos A1:B2 válidos y que no se pisen con una anterior (dos que se solapan: Excel no abre el libro). Se buscan por fila, como en
      // celdasProtegidas (contra todas las anteriores era cuadrático: 80.000 combinadas de un .xlsb, segundos), y con un tope de comparaciones para un
      // archivo hostil: pasado el tope, el resto no se combina.
      const comb = [], porFila = new Map(), altos = [];
      let comparaciones = 0;
      const pisa = (d) => (o) => d.f0 <= o.f1 && o.f0 <= d.f1 && d.c0 <= o.c1 && o.c0 <= d.c1;
      for (const r of h.combinadas || []) {
        const d = typeof r === 'string' && /^\$?[A-Za-z]{1,3}\$?\d+:\$?[A-Za-z]{1,3}\$?\d+$/.test(r) ? LibroXlsx.decodificarRango(r) : null;
        if (!d || d.f1 >= MAX_FILA || d.c1 >= MAX_COL) continue;
        const alto = d.f1 - d.f0 >= 64;
        let choca = altos.some(pisa(d));
        comparaciones += altos.length;
        const enFila = (l) => { comparaciones += l.length; return l.some(pisa(d)); };
        if (!choca && (!alto || d.f1 - d.f0 < porFila.size)) for (let f = d.f0; f <= d.f1 && !choca; f++) { const l = porFila.get(f); if (l) choca = enFila(l); }
        else if (!choca) for (const [f, l] of porFila) { comparaciones++; if (f >= d.f0 && f <= d.f1 && enFila(l)) { choca = true; break; } }
        if (comparaciones > 2e8) break;
        if (choca) continue;
        if (alto) altos.push(d);
        else for (let f = d.f0; f <= d.f1; f++) { const l = porFila.get(f); if (l) l.push(d); else porFila.set(f, [d]); }
        comb.push(numeroACol(d.c0) + (d.f0 + 1) + ':' + numeroACol(d.c1) + (d.f1 + 1));
      }
      const inm = Math.min(MAX_FILA - 1, Math.floor(Number(h.inmovilizarFilas) || 0));   // entero y dentro de la hoja: 2.5 o 1048576 y Excel no abre el libro
      const congelar = inm > 0 ? `<sheetViews><sheetView workbookViewId="0"${k === primeraVisible ? ' tabSelected="1"' : ''}><pane ySplit="${inm}" topLeftCell="A${inm + 1}" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` : `<sheetViews><sheetView workbookViewId="0"${k === primeraVisible ? ' tabSelected="1"' : ''}/></sheetViews>`;
      return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + `<dimension ref="A1${maxCol && filas.length ? ':' + numeroACol(maxCol - 1) + filas.length : ''}"/>` + congelar + '<sheetFormatPr defaultRowHeight="15"/>'
        + (anchos ? `<cols>${anchos}</cols>` : '') + `<sheetData>${filasXml.map((x) => x[1]).join('')}</sheetData>`
        + (h.autofiltro && maxCol && filas.length ? `<autoFilter ref="A1:${numeroACol(maxCol - 1)}${filas.length}"/>` : '')
        + (comb.length ? `<mergeCells count="${comb.length}">${comb.map((r) => `<mergeCell ref="${r}"/>`).join('')}</mergeCells>` : '') + '</worksheet>';
    });
    const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + hojasXml.map((_, k) => `<Override PartName="/xl/worksheets/sheet${k + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')
      + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>');
    zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
    zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><bookViews><workbookView${primeraVisible ? ` activeTab="${primeraVisible}"` : ''}/></bookViews><sheets>`
      + nombres.map((n, k) => `<sheet name="${X.escAttr(n)}" sheetId="${k + 1}"${estados[k] === 'oculta' ? ' state="hidden"' : estados[k] === 'muyOculta' ? ' state="veryHidden"' : ''} r:id="rId${k + 1}"/>`).join('') + '</sheets>' + nombresDefinidos(opciones.nombres, hojas.length) + (hayFormulas ? '<calcPr fullCalcOnLoad="1"/>' : '') + '</workbook>');
    zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + hojasXml.map((_, k) => `<Relationship Id="rId${k + 1}" Type="${R}/worksheet" Target="worksheets/sheet${k + 1}.xml"/>`).join('')
      + `<Relationship Id="rId${hojasXml.length + 1}" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rId${hojasXml.length + 2}" Type="${R}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`);
    zip.file('xl/styles.xml', estilosConFormatos(combos));
    zip.file('xl/sharedStrings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${totalTextos}" uniqueCount="${textos.length}">`
      + textos.map((t) => `<si><t xml:space="preserve">${escTexto(t)}</t></si>`).join('') + '</sst>');
    hojasXml.forEach((x, k) => zip.file(`xl/worksheets/sheet${k + 1}.xml`, x));
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  globalThis.XlsxEscritor = { editarCeldas, celdasProtegidas, crearLibro, nombreHoja, formulaConservable, renombrarHojas, ErrorEscritor };
})();

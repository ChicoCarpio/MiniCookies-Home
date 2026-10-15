/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Lectura de hojas para las herramientas de Excel/CSV (fase C del motor de Excel): un solo punto de entrada que lee .xlsx, .xlsm y .csv con el motor
// propio (LibroXlsx / CsvPlanilla) y deja SheetJS solo para lo que el motor aún no lee (.xls, .xlsb, .ods, .tsv, .txt). Las dos vías devuelven lo mismo,
// así que una herramienta no sabe quién leyó el archivo. Las filas se arman bajo demanda: un libro grande no genera una matriz entera de textos.
//
// Requiere: lector-planillas.js y SheetJS (para el respaldo), zip-propio.js, xml-ligero.js, xlsx-formato.js, xlsx-estilos.js, xlsx-modelo.js, csv-planilla.js.
//
// Uso:
//   const libro = await LectorHojas.abrir(archivo);          // lanza Error si no se puede leer (mensaje listo para mostrar)
//   libro.comaDecimalCsv                                     // solo CSV: true si el archivo usa coma decimal
//   libro.motor                                              // true si lo leyó el motor propio; false si SheetJS
//   libro.libroMotor / libro.sheetjs                         // el libro del motor (xlsx, xlsm, csv) o el de SheetJS (el resto): para imprimir
//   libro.hojas                                              // [{ nombre, indice, estado: 'visible'|'oculta'|'muyOculta' }]  (sin hojas de gráfico)
//   const hoja = await libro.cargar(i);                      // i = hojas[k].indice
//   hoja.nfilas, hoja.ncols                                  // el rango ocupado (empieza en su primera celda con contenido)
//   hoja.f0, hoja.c0                                         // dónde empieza ese rango en la hoja (desde 0); hoja.anchos / hoja.combinadas: en la hoja ('A1:B2'), no en el rango
//   hoja.texto(r, c) / hoja.fila(r)                          // lo que Excel muestra (formato de número aplicado); r y c desde 0 dentro del rango
//   hoja.valor(r, c)                                         // el valor para XlsxEscritor.crearLibro: número | texto | lógico | { error } | { valor, formato } | undefined
//   hoja.formula(r, c)                                       // texto de la fórmula sin "=" (o null)
//   hoja.matrizTextos() / hoja.matrizValores()               // todo el rango (para hojas chicas)
// Los separadores del texto siguen el idioma de la página (es: 1.234,50; en: 1,234.50); libro.opciones los cambia.
(function () {
  const esMotor = (archivo) => /\.(xlsx|xlsm|csv)$/i.test(archivo.name);

  function opcionesDeIdioma() {
    const en = typeof globalThis.idiomaActual === 'function' && globalThis.idiomaActual() === 'en';
    return { idioma: en ? 'en' : 'es', decimal: en ? '.' : ',', miles: en ? ',' : '.' };
  }

  // Los errores que se pueden escribir en una celda (los de XlsxEscritor). Uno desconocido (#SPILL! de otro programa, o un libro hostil) queda
  // como texto: crearLibro lo rechazaría y la herramienta entera fallaría por una celda.
  const ERRORES = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA']);

  // Un número con formato de fecha de un libro con base 1904 se pasa a la base 1900 (la de crearLibro), salvo que el formato sea solo de horas
  // ("h:mm", "[h]:mm"): una duración de 1,5 días es 1,5 en las dos bases (sumarle 1462 daría 35124:00 en vez de 36:00).
  function tieneCalendario(codigo) {
    const s = codigo.replace(/"[^"]*"|\\./g, '').replace(/\[(h+|m+|s+)\]/gi, '$1').replace(/\[[^\]]*\]/g, '').replace(/AM\/PM|A\/P/gi, '');
    return /[ydeg]/i.test(s) || (/m/i.test(s) && !/[hs]/i.test(s));
  }

  const VACIA = { f0: 0, c0: 0, anchos: [], combinadas: [], nfilas: 0, ncols: 0, texto: () => '', fila: () => [], valor: () => undefined, formula: () => null, matrizTextos: () => [], matrizValores: () => [] };

  function completar(h) {
    h.fila = h.fila || ((r) => { const a = new Array(h.ncols); for (let c = 0; c < h.ncols; c++) a[c] = h.texto(r, c); return a; });
    h.matrizTextos = () => { const m = []; for (let r = 0; r < h.nfilas; r++) m.push(h.fila(r)); return m; };
    // desdeA1: deja las filas y columnas vacías de antes del rango, para que cada dato quede en la misma celda que en el original.
    h.matrizValores = (desdeA1) => {
      const m = [], f0 = desdeA1 ? h.f0 : 0, c0 = desdeA1 ? h.c0 : 0;
      for (let f = 0; f < f0; f++) m.push([]);
      for (let r = 0; r < h.nfilas; r++) { const a = new Array(c0 + h.ncols).fill(undefined); for (let c = 0; c < h.ncols; c++) a[c0 + c] = h.valor(r, c); m.push(a); }
      return m;
    };
    return h;
  }

  // ---- motor propio ------------------------------------------------------------------------------------------------
  function hojaDelMotor(libro, hoja, libroLector) {
    const C = hoja.celdas, T = LibroXlsx.TIPO;
    let f0 = Infinity, f1 = -1, c0 = Infinity, c1 = -1;
    C.paraCada((f, c, i) => {
      if (C.tipo[i] === T.vacio) return;
      if (f < f0) f0 = f; if (f > f1) f1 = f; if (c < c0) c0 = c; if (c > c1) c1 = c;
    });
    if (f1 < 0) return completar(Object.assign({}, VACIA));
    const indice = (r, c) => C.indiceDe(f0 + r, c0 + c);
    const errores = (XlsxFormato.NOMBRES.en || {}).errores || {};
    const anchos = [];
    for (const col of hoja.columnas) if (col.ancho > 0 && !col.oculta) for (let q = col.min; q <= Math.min(col.max, 255); q++) anchos[q] = col.ancho;
    const combinadas = hoja.combinadas.map((m) => LibroXlsx.numeroACol(m.c0) + (m.f0 + 1) + ':' + LibroXlsx.numeroACol(m.c1) + (m.f1 + 1));
    return completar({
      f0, c0, anchos, combinadas,
      nfilas: f1 - f0 + 1, ncols: c1 - c0 + 1,
      texto(r, c) {
        const i = indice(r, c);
        return i < 0 ? '' : LibroXlsx.textoCelda(libro, hoja, i, libroLector.opciones).texto;
      },
      valor(r, c) {
        const i = indice(r, c);
        if (i < 0) return undefined;
        const tipo = C.tipo[i], v = C.valor[i];
        if (tipo === T.vacio) return undefined;
        if (tipo === T.texto || tipo === T.logico) return v;
        if (tipo === T.error) { const e = errores[v] || String(v); return ERRORES.has(e) ? { error: e } : e; }
        // número: un CSV conserva el texto original ("1.234,50") pero el valor es el número; el formato viene del estilo
        const codigo = libro.estilos.xf(C.estilo[i]).formato;
        const fecha1904 = libro.fecha1904 && codigo && XlsxFormato.esFecha(codigo) && tieneCalendario(codigo);
        const n = fecha1904 ? v + 1462 : v;
        return codigo && codigo !== 'General' ? { valor: n, formato: codigo } : n;
      },
      formula(r, c) { const i = indice(r, c); return i < 0 ? null : C.formulaDe(i); },
    });
  }

  // ---- SheetJS (respaldo) ------------------------------------------------------------------------------------------
  function hojaDeSheetJs(hojaSheetJs, libroLector, fecha1904) {
    const ref = hojaSheetJs['!ref'];
    if (!ref) return completar(Object.assign({}, VACIA));
    const rango = XLSX.utils.decode_range(ref);
    const celda = (r, c) => hojaSheetJs[XLSX.utils.encode_cell({ r: rango.s.r + r, c: rango.s.c + c })];
    const anchos = (hojaSheetJs['!cols'] || []).map((col) => (col && col.wch > 0 && !col.hidden ? col.wch : 0));
    const combinadas = (hojaSheetJs['!merges'] || []).map((m) => XLSX.utils.encode_range(m));
    return completar({
      f0: rango.s.r, c0: rango.s.c, anchos, combinadas,
      nfilas: rango.e.r - rango.s.r + 1, ncols: rango.e.c - rango.s.c + 1,
      texto(r, c) {
        const x = celda(r, c);
        if (!x) return '';
        const t = x.w !== undefined ? String(x.w) : (x.v === undefined || x.v === null ? '' : String(x.v));
        // SheetJS formatea los números al estilo inglés: se pasan a los separadores pedidos
        const o = libroLector.opciones;
        return x.t === 'n' && (o.decimal !== '.' || o.miles !== ',') ? t.replace(/[.,]/g, (ch) => (ch === '.' ? o.decimal : o.miles)) : t;
      },
      valor(r, c) {
        const x = celda(r, c);
        if (!x || x.v === undefined || x.v === null || x.t === 'z') return undefined;
        if (x.t === 'e') return { error: x.w && ERRORES.has(x.w) ? x.w : '#VALUE!' };
        if (x.t === 'n') {
          if (!x.z || x.z === 'General') return x.v;
          const z = String(x.z);   // SheetJS deja el número de serie en la base del libro: un .xls de base 1904 se pasa a 1900, como en el motor
          return { valor: fecha1904 && XlsxFormato.esFecha(z) && tieneCalendario(z) ? x.v + 1462 : x.v, formato: z };
        }
        if (x.t === 'd') return x.v;
        return x.t === 'b' ? !!x.v : String(x.v);
      },
      formula(r, c) { const x = celda(r, c); return x && x.f ? String(x.f) : null; },
    });
  }

  async function abrir(archivo) {
    const libroLector = { motor: false, hojas: [], opciones: opcionesDeIdioma(), cargar: null };
    if (esMotor(archivo)) {
      try {
        const bytes = new Uint8Array(await archivo.arrayBuffer());
        const csv = /\.csv$/i.test(archivo.name);
        const libro = csv ? CsvPlanilla.leer(bytes, archivo.name.replace(/\.csv$/i, '')) : await LibroXlsx.abrir(bytes);
        libroLector.comaDecimalCsv = libro.comaDecimalCsv;   // CSV: si el archivo usa coma decimal (undefined en un libro de Excel)
        libro.hojas.forEach((h, k) => { if (h.tipo === 'hoja') libroLector.hojas.push({ nombre: h.nombre, indice: k, estado: h.estado }); });
        if (libroLector.hojas.length) {
          libroLector.motor = true;
          if (!csv) libroLector.libroXlsx = libro;   // el libro del motor (solo lectura aquí: para editar se abre otro con LibroXlsx.abrir)
          libroLector.libroMotor = libro;            // el mismo, también en CSV: para imprimir (excel-a-pdf-motor.js)
          libroLector.cargar = async (k) => hojaDelMotor(libro, await libro.cargarHoja(k), libroLector);
          return libroLector;
        }
      } catch (error) {
        if (error && error.codigo === 'cifrado') throw error;
        // Un ".xlsx" que no es zip ni formato antiguo no es un libro: SheetJS lo leería como texto plano y mostraría basura sin avisar.
        if (error && error.codigo === 'noZip' && !(await empiezaConPK(archivo))) throw error;
        console.warn('El motor propio no pudo leer el archivo, se usa SheetJS:', error);
      }
    }
    const sheetjs = await leerLibroDesdeArchivo(archivo, { cellNF: true, cellStyles: true });   // cellNF: formato de número; cellStyles: anchos
    libroLector.comaDecimalCsv = sheetjs.comaDecimalCsv;
    libroLector.sheetjs = sheetjs;   // el libro de SheetJS: excel-a-pdf-motor.js lo pasa a .xlsx para imprimir
    libroLector.hojas = sheetjs.SheetNames.map((nombre, k) => ({ nombre, indice: k, estado: 'visible' }));
    const fecha1904 = !!(sheetjs.Workbook && sheetjs.Workbook.WBProps && sheetjs.Workbook.WBProps.date1904);
    libroLector.cargar = async (k) => hojaDeSheetJs(sheetjs.Sheets[sheetjs.SheetNames[k]], libroLector, fecha1904);
    return libroLector;
  }

  async function empiezaConPK(archivo) {
    const b = new Uint8Array(await archivo.slice(0, 2).arrayBuffer());
    return b[0] === 0x50 && b[1] === 0x4B;
  }

  globalThis.LectorHojas = { abrir };
})();

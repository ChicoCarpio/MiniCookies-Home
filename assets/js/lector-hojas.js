/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Lectura de hojas para las herramientas de Excel/CSV (fase C del motor de Excel): un solo punto de entrada que lee .xlsx, .xlsm y .csv con el motor
// propio (LibroXlsx / CsvPlanilla) y deja SheetJS solo para lo que el motor aún no lee (.xls, .xlsb, .ods, .tsv, .txt). Las dos vías devuelven lo mismo,
// así que una herramienta no sabe quién leyó el archivo. Las filas se arman bajo demanda: un libro grande no genera una matriz entera de textos.
//
// Requiere: lector-planillas.js y SheetJS (para el respaldo), zip-propio.js, xml-ligero.js, xlsx-formato.js, xlsx-estilos.js, xlsx-modelo.js, csv-planilla.js.
//
// Uso:
//   const libro = await LectorHojas.abrir(archivo);          // lanza Error si no se puede leer (mensaje listo para mostrar)
//   libro.motor                                              // true si lo leyó el motor propio; false si SheetJS
//   libro.hojas                                              // [{ nombre, indice, estado: 'visible'|'oculta'|'muyOculta' }]  (sin hojas de gráfico)
//   const hoja = await libro.cargar(i);                      // i = hojas[k].indice
//   hoja.nfilas, hoja.ncols                                  // el rango ocupado (empieza en su primera celda con contenido)
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

  const VACIA = { nfilas: 0, ncols: 0, texto: () => '', fila: () => [], valor: () => undefined, formula: () => null, matrizTextos: () => [], matrizValores: () => [] };

  function completar(h) {
    h.fila = h.fila || ((r) => { const a = new Array(h.ncols); for (let c = 0; c < h.ncols; c++) a[c] = h.texto(r, c); return a; });
    h.matrizTextos = () => { const m = []; for (let r = 0; r < h.nfilas; r++) m.push(h.fila(r)); return m; };
    h.matrizValores = () => { const m = []; for (let r = 0; r < h.nfilas; r++) { const a = new Array(h.ncols); for (let c = 0; c < h.ncols; c++) a[c] = h.valor(r, c); m.push(a); } return m; };
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
    return completar({
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
        if (tipo === T.error) { const e = errores[v] || String(v); return /^#/.test(e) ? { error: e } : e; }
        // número: un CSV conserva el texto original ("1.234,50") pero el valor es el número; el formato viene del estilo
        const codigo = libro.estilos.xf(C.estilo[i]).formato;
        const fecha1904 = libro.fecha1904 && codigo && XlsxFormato.esFecha(codigo);
        const n = fecha1904 ? v + 1462 : v;
        return codigo && codigo !== 'General' ? { valor: n, formato: codigo } : n;
      },
      formula(r, c) { const i = indice(r, c); return i < 0 ? null : C.formulaDe(i); },
    });
  }

  // ---- SheetJS (respaldo) ------------------------------------------------------------------------------------------
  function hojaDeSheetJs(hojaSheetJs) {
    const ref = hojaSheetJs['!ref'];
    if (!ref) return completar(Object.assign({}, VACIA));
    const rango = XLSX.utils.decode_range(ref);
    const celda = (r, c) => hojaSheetJs[XLSX.utils.encode_cell({ r: rango.s.r + r, c: rango.s.c + c })];
    return completar({
      nfilas: rango.e.r - rango.s.r + 1, ncols: rango.e.c - rango.s.c + 1,
      texto(r, c) { const x = celda(r, c); return x ? (x.w !== undefined ? String(x.w) : (x.v === undefined || x.v === null ? '' : String(x.v))) : ''; },
      valor(r, c) {
        const x = celda(r, c);
        if (!x || x.v === undefined || x.v === null || x.t === 'z') return undefined;
        if (x.t === 'e') return { error: x.w && /^#/.test(x.w) ? x.w : '#VALUE!' };
        if (x.t === 'n') return x.z && x.z !== 'General' ? { valor: x.v, formato: x.z } : x.v;
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
        libro.hojas.forEach((h, k) => { if (h.tipo === 'hoja') libroLector.hojas.push({ nombre: h.nombre, indice: k, estado: h.estado }); });
        if (libroLector.hojas.length) {
          libroLector.motor = true;
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
    const sheetjs = await leerLibroDesdeArchivo(archivo);
    libroLector.hojas = sheetjs.SheetNames.map((nombre, k) => ({ nombre, indice: k, estado: 'visible' }));
    libroLector.cargar = async (k) => hojaDeSheetJs(sheetjs.Sheets[sheetjs.SheetNames[k]]);
    return libroLector;
  }

  async function empiezaConPK(archivo) {
    const b = new Uint8Array(await archivo.slice(0, 2).arrayBuffer());
    return b[0] === 0x50 && b[1] === 0x4B;
  }

  globalThis.LectorHojas = { abrir };
})();

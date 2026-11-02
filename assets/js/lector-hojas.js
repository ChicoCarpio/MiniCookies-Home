/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Lectura de hojas para las herramientas de Excel/CSV (fase C del motor de Excel): un solo punto de entrada que lee .xlsx y .xlsm (LibroXlsx), .csv, .tsv y .txt
// (CsvPlanilla) y, desde la etapa 9, .xls, .xlsb y .ods (formatos-antiguos.js los pasa a un .xlsx en memoria, sin fórmulas ni estilos de celda). Las filas
// se arman bajo demanda: un libro grande no genera una matriz entera de textos.
//
// Requiere: zip-propio.js, xml-ligero.js, xml-editar.js, xlsx-formato.js, xlsx-estilos.js, xlsx-modelo.js, csv-planilla.js y, para los formatos antiguos,
// xlsx-escritor.js, formatos-antiguos.js, ods-lector.js, xlsb-lector.js, cfb-propio.js, xls-lector.js, xml2003-lector.js y html-tabla-lector.js.
//
// Uso:
//   const libro = await LectorHojas.abrir(archivo);          // lanza Error si no se puede leer (mensaje listo para mostrar)
//   libro.comaDecimalCsv                                     // solo CSV: true si el archivo usa coma decimal
//   libro.motor                                              // siempre true (se conserva para las herramientas que lo preguntaban)
//   libro.libroMotor                                         // el libro del motor (también el de un CSV o un .xls pasado a .xlsx): para imprimir
//   libro.libroXlsx                                          // el libro del motor si el archivo no es de texto (solo lectura aquí: para editar se abre otro con LibroXlsx.abrir)
//   libro.bytesXlsx                                          // solo si venía en .xls, .xlsb u .ods: los bytes del .xlsx en que se convirtió (valores y formato de número). Con
//                                                            // ellos se edita, se divide o se repara igual que un .xlsx
//   libro.hojas                                              // [{ nombre, indice, estado: 'visible'|'oculta'|'muyOculta' }]  (sin hojas de gráfico)
//   const hoja = await libro.cargar(i);                      // i = hojas[k].indice
//   hoja.nfilas, hoja.ncols                                  // el rango ocupado (empieza en su primera celda con contenido)
//   hoja.columnasOcultas / hoja.filasOcultas                 // columnas y filas ocultas de la hoja (desde 0); hoja.opcionesLibro() → { anchos, combinadas, columnasOcultas, filasOcultas } para crearLibro
//   hoja.f0, hoja.c0                                         // dónde empieza ese rango en la hoja (desde 0); hoja.anchos / hoja.combinadas: en la hoja ('A1:B2'), no en el rango
//   hoja.texto(r, c) / hoja.fila(r)                          // lo que Excel muestra (formato de número aplicado); r y c desde 0 dentro del rango
//   hoja.valor(r, c)                                         // el valor para XlsxEscritor.crearLibro: número | texto | lógico | { error } | { valor, formato } | undefined
//   hoja.formula(r, c)                                       // texto de la fórmula sin "=" (o null)
//   hoja.formulaSimple(r, c)                                 // lo mismo, pero null si es matricial o tabla de datos
//   hoja.matrizValores(true, { formulas: { hojas, mapa } })  // con las fórmulas que se puedan llevar al libro de destino (ver abajo)
//   hoja.matrizTextos() / hoja.matrizValores()               // todo el rango (para hojas chicas)
// Los separadores del texto siguen el idioma de la página (es: 1.234,50; en: 1,234.50); libro.opciones los cambia.
(function () {
  // Los lectores de .xls, .xlsb y .ods están al lado de este archivo y se piden recién cuando llega uno: la mayoría de los archivos son .xlsx o .csv.
  // La ruta sale de este script y no de la página (que existe también bajo /en/).
  const base = typeof document !== 'undefined' && document.currentScript ? new URL('./', document.currentScript.src).href : null;
  const ANTIGUOS = [['XmlEditar', 'xml-editar.js'], ['XlsxEscritor', 'xlsx-escritor.js'], ['OdsLector', 'ods-lector.js'], ['XlsbLector', 'xlsb-lector.js'], ['CfbPropio', 'cfb-propio.js'], ['XlsLector', 'xls-lector.js'], ['Xml2003Lector', 'xml2003-lector.js'], ['HtmlTablaLector', 'html-tabla-lector.js'], ['FormatosAntiguos', 'formatos-antiguos.js']];
  async function cargarAntiguos() {
    for (const [global, archivo] of ANTIGUOS) {
      if (globalThis[global] || !base) continue;
      await new Promise((resolver, rechazar) => {
        const s = document.createElement('script');
        s.src = base + archivo;
        s.onload = resolver;
        s.onerror = () => rechazar(new Error((typeof globalThis.idiomaActual === 'function' && globalThis.idiomaActual() === 'en' ? 'Couldn’t load the file reader (offline?): ' : 'No se pudo cargar el lector del archivo (¿sin conexión?): ') + archivo));
        document.head.appendChild(s);
      });
    }
  }

  function opcionesDeIdioma() {
    const en = typeof globalThis.idiomaActual === 'function' && globalThis.idiomaActual() === 'en';
    return { idioma: en ? 'en' : 'es', decimal: en ? '.' : ',', miles: en ? ',' : '.' };
  }

  // Los errores que se pueden escribir en una celda (los de XlsxEscritor). Uno desconocido (#SPILL! de otro programa, o un libro hostil) queda
  // como texto: crearLibro lo rechazaría y la herramienta entera fallaría por una celda.
  const ERRORES = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA']);

  const VACIA = { f0: 0, c0: 0, anchos: [], combinadas: [], columnasOcultas: [], filasOcultas: [], nfilas: 0, ncols: 0, texto: () => '', fila: () => [], valor: () => undefined, formula: () => null, matrizTextos: () => [], matrizValores: () => [] };

  function completar(h) {
    // lo que XlsxEscritor.crearLibro acepta de una hoja además de sus filas: { anchos, combinadas, columnasOcultas, filasOcultas } (todo en la hoja, no en el rango)
    h.opcionesLibro = () => ({ anchos: h.anchos, combinadas: h.combinadas, columnasOcultas: h.columnasOcultas, filasOcultas: h.filasOcultas });
    h.fila = h.fila || ((r) => { const a = new Array(h.ncols); for (let c = 0; c < h.ncols; c++) a[c] = h.texto(r, c); return a; });
    h.matrizTextos = () => { const m = []; for (let r = 0; r < h.nfilas; r++) m.push(h.fila(r)); return m; };
    // desdeA1: deja las filas y columnas vacías de antes del rango, para que cada dato quede en la misma celda que en el original.
    // opc.formulas = { hojas: Set de nombres (minúsculas) del libro de destino, mapa?: Map nombre viejo (minúsculas) → nombre nuevo }: las celdas con fórmula salen como
    // { formula, valor } cuando la fórmula se puede llevar al libro nuevo (XlsxEscritor.formulaConservable); si no, solo su valor. Requiere xlsx-escritor.js.
    h.matrizValores = (desdeA1, opc) => {
      const m = [], f0 = desdeA1 ? h.f0 : 0, c0 = desdeA1 ? h.c0 : 0;
      const pf = opc && opc.formulas;
      const celda = !pf || !h.formulaSimple ? (r, c) => h.valor(r, c) : (r, c) => {
        const v = h.valor(r, c), fx = h.formulaSimple(r, c);
        if (!fx) return v;
        const f = pf.mapa ? XlsxEscritor.renombrarHojas(fx, pf.mapa) : fx;
        return XlsxEscritor.formulaConservable(f, pf.hojas) ? { formula: f, valor: v } : v;
      };
      for (let f = 0; f < f0; f++) m.push([]);
      for (let r = 0; r < h.nfilas; r++) { const a = new Array(c0 + h.ncols).fill(undefined); for (let c = 0; c < h.ncols; c++) a[c0 + c] = celda(r, c); m.push(a); }
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
    // se marcan y después se listan: 16.384 <col> "A:XFD ocultas" (3 KB comprimidos) eran 268 millones de entradas repetidas
    const columnasOcultas = [], filasOcultas = [], marcas = new Uint8Array(16384);
    for (const col of hoja.columnas) if (col.oculta) marcas.fill(1, col.min, Math.min(col.max, 16383) + 1);
    for (let q = 0; q < 16384; q++) if (marcas[q]) columnasOcultas.push(q);
    hoja.filas.forEach((info, fila) => { if (info.oculta) filasOcultas.push(fila); });
    const combinadas = hoja.combinadas.map((m) => LibroXlsx.numeroACol(m.c0) + (m.f0 + 1) + ':' + LibroXlsx.numeroACol(m.c1) + (m.f1 + 1));
    return completar({
      f0, c0, anchos, combinadas, columnasOcultas, filasOcultas,
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
        const fecha1904 = libro.fecha1904 && codigo && XlsxFormato.esFecha(codigo) && XlsxFormato.tieneCalendario(codigo);
        const n = fecha1904 ? v + 1462 : v;
        return codigo && codigo !== 'General' ? { valor: n, formato: codigo } : n;
      },
      formula(r, c) { const i = indice(r, c); return i < 0 ? null : C.formulaDe(i); },
      // la fórmula de una celda que se puede escribir sola: sin las matriciales ni las tablas de datos (esas viven en un rango)
      formulaSimple(r, c) { const i = indice(r, c); if (i < 0) return null; const fx = C.formulas.get(i); return !fx || fx.tipo === 'array' || fx.tipo === 'dataTable' ? null : C.formulaDe(i); },
    });
  }

  // Un lector de formato antiguo o crearLibro que fallan con un archivo raro dejan un error técnico ("Invalid array length", "demasiadas hojas", el de un
  // zip): sale como ErrorXlsx, con su mensaje en el idioma de la página, igual que lo que lanza LibroXlsx.abrir.
  function comoErrorXlsx(error) {
    if (error instanceof LibroXlsx.ErrorXlsx) return error;
    const detalle = error && error.message ? error.message : String(error);
    if (error && error.name === 'ErrorZip') return new LibroXlsx.ErrorXlsx(error.codigo === 'grande' ? 'grande' : error.codigo === 'cifrado' ? 'cifrado' : 'danado', detalle);
    const grande = error instanceof RangeError || (typeof XlsxEscritor !== 'undefined' && error instanceof XlsxEscritor.ErrorEscritor);   // crearLibro: muy grande, demasiadas hojas o filas
    return new LibroXlsx.ErrorXlsx(grande ? 'grande' : 'danado', detalle);
  }

  // Sin ningún byte nulo en el principio: texto, no un binario dañado. Con la marca de UTF-16 también ("Texto Unicode" de Excel: UTF-16 con tabuladores)
  const esTextoPlano = (bytes) => (bytes.length >= 2 && ((bytes[0] === 0xFF && bytes[1] === 0xFE) || (bytes[0] === 0xFE && bytes[1] === 0xFF))) || (bytes.length > 0 && !bytes.subarray(0, 4096).includes(0));

  // CSV, TSV y texto delimitado: CsvPlanilla (decodifica, detecta el separador y distingue números de códigos)
  const esTexto = (archivo) => /\.(csv|tsv|txt)$/i.test(archivo.name) || archivo.type === 'text/csv';

  async function abrir(archivo) {
    const libroLector = { motor: true, hojas: [], opciones: opcionesDeIdioma(), cargar: null };
    let bytes = new Uint8Array(await archivo.arrayBuffer());
    let texto = esTexto(archivo);
    let libro;
    if (texto) libro = CsvPlanilla.leer(bytes, archivo.name.replace(/\.(csv|tsv|txt)$/i, ''));
    else {
      try { libro = await LibroXlsx.abrir(bytes); } catch (error) {
        // .xls, .xlsb y .ods: los lectores de formatos-antiguos.js los pasan a un .xlsx en memoria, que el motor ya sabe leer e imprimir
        if (!error || error.codigo === 'cifrado' || !['antiguo', 'sinLibro', 'noZip'].includes(error.codigo)) throw error;
        await cargarAntiguos();
        if (await FormatosAntiguos.detectar(bytes)) {
          bytes = await FormatosAntiguos.aXlsx(bytes).catch((e) => { throw comoErrorXlsx(e); });
          libro = await LibroXlsx.abrir(bytes);
          libroLector.bytesXlsx = bytes;
        } else if (/\.xls$/i.test(archivo.name) && esTextoPlano(bytes)) {
          // un ".xls" que es texto separado por tabuladores o comas (también lo hacen algunos sistemas); un ".xlsx" de texto sí es un archivo dañado.
          // Se pasa a .xlsx como un .xls de verdad: las herramientas que editan (buscar y reemplazar, dividir…) trabajan sobre bytesXlsx.
          const csv = CsvPlanilla.leer(bytes, archivo.name.replace(/\.[^.]+$/, ''));
          const hoja = hojaDelMotor(csv, await csv.cargarHoja(0), libroLector);
          bytes = await XlsxEscritor.crearLibro([{ nombre: csv.hojas[0].nombre, filas: hoja.matrizValores(true) }]).catch((e) => { throw comoErrorXlsx(e); });
          libro = await LibroXlsx.abrir(bytes);
          libroLector.bytesXlsx = bytes;
        } else throw error;
      }
    }
    libroLector.comaDecimalCsv = libro.comaDecimalCsv;   // texto: si el archivo usa coma decimal (undefined en un libro de Excel)
    libro.hojas.forEach((h, k) => { if (h.tipo === 'hoja') libroLector.hojas.push({ nombre: h.nombre, indice: k, estado: h.estado }); });
    if (!libroLector.hojas.length) throw new LibroXlsx.ErrorXlsx('sinLibro');
    if (!texto) libroLector.libroXlsx = libro;   // el libro del motor (solo lectura aquí: para editar se abre otro con LibroXlsx.abrir)
    libroLector.libroMotor = libro;              // el mismo, también en texto: para imprimir (excel-a-pdf-motor.js)
    libroLector.cargar = async (k) => hojaDelMotor(libro, await libro.cargarHoja(k), libroLector);
    return libroLector;
  }

  globalThis.LectorHojas = { abrir };
})();

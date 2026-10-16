/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Excel / CSV → PDF con el motor propio (fase C del motor de Excel): imprime como Excel (mismas páginas, saltos, área de impresión,
// títulos, encabezado y pie, bordes, formato condicional, objetos y gráficos) con xlsx-paginas.js + xlsx-pdf.js, y deja que la
// pantalla cambie encima lo que el archivo dice: papel, orientación, "ajustar a N páginas de ancho", cuadrícula, rangos a imprimir
// y columnas/filas ocultas. Nada se escribe en el archivo: las sobreposiciones se aplican a la hoja en memoria y se deshacen al terminar.
//
// Uso:
//   await ExcelAPdfMotor.cargarMotor();                              // trae los módulos pesados la primera vez
//   const libro = await ExcelAPdfMotor.libroDeMotor(libroLector);    // libroLector = lo que devuelve LectorHojas.abrir
//   const { bytes, paginas, truncado } = await ExcelAPdfMotor.generar(libro, {
//     hojas: [0, 2],                  // índices del libro
//     papel: 'archivo' | 'a4' | … | [anchoPt, altoPt],  orientacion: 'archivo' | 'vertical' | 'horizontal',
//     ajustar: 'archivo' | 'no' | 1 | 2 | 3 | 'hoja' | 'columnas' | 'filas' | { ancho, alto } | { porcentaje },   // páginas de ancho (alto libre) o una escala de Excel
//     margenes: 'archivo' | 'normal' | 'ancho' | 'estrecho' | { izq, der, arr, aba, encabezado, pie } (pulgadas),
//     centrarH, centrarV: true | false | null (null: lo del archivo),  encabezadoPie: 'archivo' | 'ninguno' | 'paginas' | 'hoja',
//     cuadricula: 'archivo' | 'si' | 'no',
//     areas: [{ f0, c0, f1, c1 }] | null,                // desde 0; solo con una hoja
//     incluirOcultas: false,                             // imprimir también columnas y filas ocultas
//     maxPaginas, progreso(n, total), idioma, nombreArchivo });
// Requiere vendor.js (rutaVendor), motor-hojas.js, lector-hojas.js y los módulos de xlsx-*.js que lector-hojas.js ya necesita.

// Se suman a los textos de la página (los de la página mandan).
(function () {
  const pagina = window.i18nPagina || (window.i18nPagina = {});
  pagina.es = Object.assign({ 'excel-pdf-sin-datos': 'El archivo no tiene nada que imprimir.', 'excel-pdf-margenes': 'Los márgenes no caben en el papel: achicalos o elegí un papel más grande.' }, pagina.es);
  pagina.en = Object.assign({ 'excel-pdf-sin-datos': 'The file has nothing to print.', 'excel-pdf-margenes': 'The margins don’t fit the paper: make them smaller or choose a larger paper.' }, pagina.en);
})();

const ExcelAPdfMotor = (() => {
  // Papel en puntos (vertical). Los A son los de la norma ISO redondeados como los usa Excel; el resto, los de la lista de la herramienta.
  const mm = (a, b) => [a * 72 / 25.4, b * 72 / 25.4];
  const PAPELES = {
    '4a0': mm(1682, 2378), '2a0': mm(1189, 1682), 'arch-e': mm(914.4, 1219.2), a0: mm(841, 1189), 'ansi-e': mm(863.6, 1117.6), a1: mm(594, 841), a2: mm(420, 594),
    a3: [841.92, 1190.52], a4: [595.32, 841.92], a5: [419.52, 595.32], a6: [297.6, 419.52], b4: mm(257, 364), b5: mm(182, 257),
    carta: [612, 792], legal: [612, 1008], oficio: [612, 936], tabloide: [792, 1224], ejecutivo: [522, 756], 'media-carta': [396, 612], folio: [612, 936],
  };

  // Los módulos pesados los trae MotorHojas (motor-hojas.js), que comparte con el visor
  const cargarMotor = () => MotorHojas.cargarPdf();

  // El libro del motor para imprimir: el que ya abrió LectorHojas. Un .xls, .xlsb u .ods ya viene pasado a .xlsx (valores, formatos de número,
  // anchos y combinadas; sin estilos de celda).
  async function libroDeMotor(libroLector) {
    if (!libroLector.libroMotor) throw new Error('No se pudo preparar el libro para imprimir.');
    return libroLector.libroMotor;
  }

  // Márgenes predefinidos de Excel, en pulgadas (izquierdo, derecho, superior, inferior, encabezado, pie)
  const MARGENES = {
    normal: { izq: 0.7, der: 0.7, arr: 0.75, aba: 0.75, encabezado: 0.3, pie: 0.3 },
    ancho: { izq: 1, der: 1, arr: 1, aba: 1, encabezado: 0.5, pie: 0.5 },
    estrecho: { izq: 0.25, der: 0.25, arr: 0.75, aba: 0.75, encabezado: 0.3, pie: 0.3 },
  };

  const PAGINA_BASE = { papel: 9, anchoPapel: null, altoPapel: null, orientacion: 'default', escala: 100, ajustarAncho: 1, ajustarAlto: 1, orden: 'downThenOver', primeraPagina: null, usarPrimeraPagina: false, blancoYNegro: false, borrador: false, comentarios: 'none', errores: 'displayed' };

  const celdaA1 = (f, c) => LibroXlsx.numeroACol(c) + (f + 1);
  const comillas = (nombre) => "'" + String(nombre).replace(/'/g, "''") + "'";

  // Aplica las sobreposiciones a una hoja cargada y devuelve cómo deshacerlas (si algo falla a mitad, deshace lo que alcanzó a poner).
  function sobreponer(libro, hoja, indice, o) {
    const deshacer = [];
    const todo = () => { for (let i = deshacer.length - 1; i >= 0; i--) deshacer[i](); };
    try { sobreponerEn(libro, hoja, indice, o, deshacer); } catch (error) { todo(); throw error; }
    return todo;
  }
  function sobreponerEn(libro, hoja, indice, o, deshacer) {
    const imp = hoja.imprimir || (hoja.imprimir = { margenes: null, pagina: null, opciones: null, encabezadoPie: null, saltosFila: [], saltosCol: [], ajustarAPagina: false, autoSaltos: true });
    const antes = { ajustarAPagina: imp.ajustarAPagina, pagina: imp.pagina, opciones: imp.opciones, margenes: imp.margenes, encabezadoPie: imp.encabezadoPie };
    deshacer.push(() => { imp.ajustarAPagina = antes.ajustarAPagina; imp.pagina = antes.pagina; imp.opciones = antes.opciones; imp.margenes = antes.margenes; imp.encabezadoPie = antes.encabezadoPie; });
    imp.pagina = Object.assign({}, imp.pagina || (libro.esCsv ? PAGINA_BASE : Object.assign({}, PAGINA_BASE, { papel: 1 })));

    if (Array.isArray(o.papel)) { imp.pagina.anchoPapel = Math.min(o.papel[0], o.papel[1]) + 'pt'; imp.pagina.altoPapel = Math.max(o.papel[0], o.papel[1]) + 'pt'; }
    else if (o.papel && o.papel !== 'archivo' && Object.prototype.hasOwnProperty.call(PAPELES, o.papel)) { imp.pagina.anchoPapel = PAPELES[o.papel][0] + 'pt'; imp.pagina.altoPapel = PAPELES[o.papel][1] + 'pt'; }
    if (o.orientacion === 'vertical') imp.pagina.orientacion = 'portrait';
    else if (o.orientacion === 'horizontal') imp.pagina.orientacion = 'landscape';

    // Escala: 'no' (100 %), N páginas de ancho, 'hoja' (1 × 1), 'columnas' (1 de ancho), 'filas' (1 de alto), { ancho, alto } (0 = libre) o { porcentaje }
    const ajustar = o.ajustar, fija = (n) => Math.min(99, Math.max(0, Math.floor(Number(n)) || 0));
    if (ajustar === 'no') { imp.ajustarAPagina = false; imp.pagina.escala = 100; }
    else if (ajustar > 0) { imp.ajustarAPagina = true; imp.pagina.ajustarAncho = ajustar; imp.pagina.ajustarAlto = 0; }
    else if (ajustar === 'hoja') { imp.ajustarAPagina = true; imp.pagina.ajustarAncho = 1; imp.pagina.ajustarAlto = 1; }
    else if (ajustar === 'columnas') { imp.ajustarAPagina = true; imp.pagina.ajustarAncho = 1; imp.pagina.ajustarAlto = 0; }
    else if (ajustar === 'filas') { imp.ajustarAPagina = true; imp.pagina.ajustarAncho = 0; imp.pagina.ajustarAlto = 1; }
    else if (ajustar && typeof ajustar === 'object' && ajustar.porcentaje > 0) { imp.ajustarAPagina = false; imp.pagina.escala = Math.min(400, Math.max(10, Math.round(ajustar.porcentaje))); }
    else if (ajustar && typeof ajustar === 'object' && (ajustar.ancho > 0 || ajustar.alto > 0)) { imp.ajustarAPagina = true; imp.pagina.ajustarAncho = fija(ajustar.ancho); imp.pagina.ajustarAlto = fija(ajustar.alto); }

    // Márgenes (pulgadas, como los guarda el archivo): 'normal', 'ancho', 'estrecho' son los de Excel; también { izq, der, arr, aba, encabezado, pie }
    const preajuste = typeof o.margenes === 'string' ? MARGENES[o.margenes] : (o.margenes && typeof o.margenes === 'object' ? o.margenes : null);
    if (preajuste) {
      imp.margenes = Object.assign({}, MARGENES.normal, preajuste);
      // como el diálogo de Excel: unos márgenes que no dejan lugar en el papel no se aceptan (darían una celda por página)
      const papel = XlsxPaginas.papelDe(imp), mg = imp.margenes;
      if (!((+mg.izq + +mg.der) * 72 < papel.ancho && (+mg.arr + +mg.aba) * 72 < papel.alto)) throw new Error(typeof t === 'function' ? t('excel-pdf-margenes') : 'Los márgenes no caben en el papel.');
    }

    imp.opciones = Object.assign({ cuadricula: false, encabezados: false, centrarH: false, centrarV: false }, imp.opciones);
    if (o.cuadricula === 'si') imp.opciones.cuadricula = true; else if (o.cuadricula === 'no') imp.opciones.cuadricula = false;
    if (o.centrarH === true || o.centrarH === false) imp.opciones.centrarH = o.centrarH;
    if (o.centrarV === true || o.centrarV === false) imp.opciones.centrarV = o.centrarV;

    // Encabezado y pie: 'ninguno', 'paginas' (pie «Página N de M») o 'hoja' (nombre de la hoja arriba); 'archivo' deja lo que dice el archivo
    if (o.encabezadoPie === 'ninguno') imp.encabezadoPie = null;
    else if (o.encabezadoPie === 'paginas' || o.encabezadoPie === 'hoja') {
      const pagina = o.idioma === 'en' ? 'Page &P of &N' : 'Página &P de &N';
      const grupo = o.encabezadoPie === 'paginas' ? { encabezado: '', pie: '&C' + pagina } : { encabezado: '&C&A', pie: '' };
      imp.encabezadoPie = { distintoPrimera: false, distintoParImpar: false, escalar: true, alinear: true, impar: grupo, par: {}, primera: {} };
    }

    if (o.areas && o.areas.length) {
      const nombreHoja = libro.hojas[indice].nombre;
      const ref = o.areas.map((a) => comillas(nombreHoja) + '!$' + LibroXlsx.numeroACol(a.c0) + '$' + (a.f0 + 1) + ':$' + LibroXlsx.numeroACol(a.c1) + '$' + (a.f1 + 1)).join(',');
      const previo = libro.nombres.findIndex((n) => n.nombre === '_xlnm.Print_Area' && n.hojaLocal === hoja.indice);
      const original = previo >= 0 ? libro.nombres[previo] : null;
      const nuevo = { nombre: '_xlnm.Print_Area', hojaLocal: hoja.indice, ref };
      if (previo >= 0) libro.nombres[previo] = nuevo; else libro.nombres.push(nuevo);
      deshacer.push(() => { const k = libro.nombres.indexOf(nuevo); if (k < 0) return; if (original) libro.nombres[k] = original; else libro.nombres.splice(k, 1); });
    }

    if (o.incluirOcultas) {
      // `ocultaOriginal` deja constancia de lo que decía el archivo mientras dura la impresión: quien lea la hoja entretanto (la cuadrícula de selección de la página)
      // usa `oculta || ocultaOriginal`
      for (const col of hoja.columnas) if (col.oculta) { col.oculta = false; col.ocultaOriginal = true; deshacer.push(() => { col.oculta = true; delete col.ocultaOriginal; }); }
      for (const info of hoja.filas.values()) if (info.oculta) { info.oculta = false; info.ocultaOriginal = true; deshacer.push(() => { info.oculta = true; delete info.ocultaOriginal; }); }
    }
  }

  // Las sobreposiciones se ponen en la hoja en memoria mientras se calcula o se dibuja y se quitan al terminar; como dos trabajos a la vez se
  // pisarían, todo lo que las usa (el PDF y la vista previa) entra en una cola.
  let cola = Promise.resolve();
  const enCola = (tarea) => { const r = cola.then(tarea); cola = r.catch(() => {}); return r; };

  // Si una hoja falla (no carga, un valor raro), se deshace lo ya puesto en las anteriores: la hoja en memoria no puede quedar cambiada.
  async function ponerSobreposiciones(libro, o) {
    const deshacer = [];
    const todo = () => { for (let k = deshacer.length - 1; k >= 0; k--) deshacer[k](); deshacer.length = 0; };
    try {
      for (const i of o.hojas) {
        const hoja = await libro.cargarHoja(i);
        deshacer.push(sobreponer(libro, hoja, i, o.hojas.length === 1 ? o : Object.assign({}, o, { areas: null })));
      }
    } catch (error) { todo(); throw error; }
    return todo;
  }

  function generar(libro, o) {
    return enCola(async () => {
      await cargarMotor();
      const deshacer = await ponerSobreposiciones(libro, o);
      try {
        return await XlsxPdf.generar(libro, o.hojas, { idioma: o.idioma || 'es', nombreArchivo: o.nombreArchivo, maxPaginas: o.maxPaginas, progreso: o.progreso, numeracion: 'porHoja' });
      } finally { deshacer(); }
    });
  }

  // La vista previa de las páginas, sin armar el PDF: calcula las páginas con las mismas opciones que `generar` y deja dibujar cada una
  // en un <canvas> con el mismo dibujo que el PDF. No tiene tope de páginas (el cálculo no dibuja nada).
  //   const v = await ExcelAPdfMotor.vistaPrevia(libro, opciones);   // v.total, v.truncado, v.medida(n) → { ancho, alto, hoja } en pt
  //   await v.dibujar(n, canvas, { ancho: 600, dpr: 2, margenes: true });
  function vistaPrevia(libro, o) {
    return enCola(async () => {
      await cargarMotor();
      const deshacer = await ponerSobreposiciones(libro, o);
      let plan;
      try {
        plan = await XlsxPdf.planificar(libro, o.hojas, { idioma: o.idioma || 'es', nombreArchivo: o.nombreArchivo, numeracion: 'porHoja', paraPantalla: true });
      } finally { deshacer(); }
      return {
        total: plan.total, truncado: plan.truncado,
        medida: (n) => XlsxPdf.medidaPagina(plan, n),
        dibujar: (n, canvas, op) => enCola(async () => {
          // solo la hoja de esa página (con «Todas las hojas», no todas en cada miniatura); las áreas valen solo con una hoja
          const m = XlsxPdf.medidaPagina(plan, n);
          const quitar = await ponerSobreposiciones(libro, m && o.hojas.length > 1 ? Object.assign({}, o, { hojas: [o.hojas[m.indiceHoja]], areas: null }) : o);
          try { await XlsxPdf.dibujarEnCanvas(plan, n, canvas, op); } finally { quitar(); }
        }),
      };
    });
  }

  // Un libro entero → los bytes de un PDF con las páginas que imprime Excel (lo que usa "Unir PDF"): todas las hojas visibles, con la
  // configuración del archivo; un CSV sale con cuadrícula y a una página de ancho. Falla con `t('excel-pdf-sin-datos')` si no hay nada que imprimir.
  // alAvanzar(fracción 0–1).
  async function libroAPdf(archivo, alAvanzar) {
    const lector = await LectorHojas.abrir(archivo);
    const libro = await libroDeMotor(lector);
    const csv = /\.csv$/i.test(archivo.name);
    const hojas = lector.hojas.filter((h) => h.estado === 'visible').map((h) => h.indice);
    const idioma = typeof idiomaActual === 'function' && idiomaActual() === 'en' ? 'en' : 'es';
    const { bytes, paginas } = await generar(libro, {
      hojas, idioma, nombreArchivo: archivo.name,
      ajustar: csv ? 1 : 'archivo', cuadricula: csv ? 'si' : 'archivo',
      progreso: alAvanzar ? (n, total) => alAvanzar(n / total) : null,
    });
    if (!paginas) throw new Error(t('excel-pdf-sin-datos'));
    return bytes;
  }

  return { cargarMotor, libroDeMotor, generar, vistaPrevia, libroAPdf, PAPELES, MARGENES };
})();

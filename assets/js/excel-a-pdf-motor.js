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
//     ajustar: 'archivo' | 'no' | 1 | 2 | 3,             // páginas de ancho (alto libre)
//     cuadricula: 'archivo' | 'si' | 'no',
//     areas: [{ f0, c0, f1, c1 }] | null,                // desde 0; solo con una hoja
//     incluirOcultas: false,                             // imprimir también columnas y filas ocultas
//     maxPaginas, progreso(n, total), idioma, nombreArchivo });
// Requiere vendor.js (rutaVendor), lector-hojas.js y los módulos de xlsx-*.js que lector-hojas.js ya necesita.

const ExcelAPdfMotor = (() => {
  const base = new URL('./', document.currentScript.src).href;
  const MODULOS = [
    'pptx-xml.js', 'vendor/ecma-376-geometrias/5ed-2016/geometrias.js', 'emf.js', 'pptx-geometria.js', 'pptx-estilos-tabla.js', 'pptx-modelo.js', 'pptx-texto.js',
    'pptx-grafico-3d.js', 'pptx-grafico-superficie.js', 'pptx-grafico.js', 'pptx-dibujo.js', 'xlsx-objetos.js',
    'xlsx-geometria.js', 'xlsx-fuentes-impresion.js', 'xlsx-paginas.js', 'xlsx-expresion.js', 'xlsx-condicional.js', 'xlsx-iconos.js', 'xlsx-minigraficos.js',
    'xlsx-notas.js', 'xlsx-filtros.js', 'xlsx-tablas-estilos.js', 'xlsx-tablas.js', 'xlsx-dibujo.js', 'vendor/pdf-lib/1.17.1/pdf-lib.min.js', 'pdf-lienzo.js', 'xlsx-pdf.js',
  ];

  // Papel en puntos (vertical). Los A son los de la norma ISO redondeados como los usa Excel; el resto, los de la lista de la herramienta.
  const mm = (a, b) => [a * 72 / 25.4, b * 72 / 25.4];
  const PAPELES = {
    '4a0': mm(1682, 2378), '2a0': mm(1189, 1682), 'arch-e': mm(914.4, 1219.2), a0: mm(841, 1189), 'ansi-e': mm(863.6, 1117.6), a1: mm(594, 841), a2: mm(420, 594),
    a3: [841.92, 1190.52], a4: [595.32, 841.92], a5: [419.52, 595.32], a6: [297.6, 419.52], b4: mm(257, 364), b5: mm(182, 257),
    carta: [612, 792], legal: [612, 1008], oficio: [612, 936], tabloide: [792, 1224], ejecutivo: [522, 756], 'media-carta': [396, 612], folio: [612, 936],
  };

  const cargas = new Map();
  function cargarScript(src) {
    if (!cargas.has(src) && document.querySelector('script[src="' + src + '"]')) cargas.set(src, Promise.resolve());
    if (!cargas.has(src)) {
      cargas.set(src, new Promise((resolver, rechazar) => {
        const s = document.createElement('script');
        s.src = src;
        s.onload = resolver;
        s.onerror = () => { cargas.delete(src); rechazar(new Error('No se pudo cargar ' + src.split('/').pop().replace(/\?.*$/, ''))); };
        document.head.appendChild(s);
      }));
    }
    return cargas.get(src);
  }
  async function cargarMotor() {
    if (typeof XlsxPdf !== 'undefined') return;
    for (const x of MODULOS) await cargarScript(x.startsWith('vendor/') ? rutaVendor(x.slice(7)) : base + x);
  }

  // El libro del motor para imprimir: el que ya abrió LectorHojas (xlsx, xlsm y csv) o, si el archivo es de otro formato (xls, xlsb, ods),
  // el que sale de pasarlo a .xlsx con SheetJS: conserva valores, formatos de número, anchos, combinadas y columnas ocultas.
  async function libroDeMotor(libroLector) {
    if (libroLector.libroMotor) return libroLector.libroMotor;
    if (!libroLector.sheetjs) throw new Error('No se pudo preparar el libro para imprimir.');
    const bytes = XLSX.write(libroLector.sheetjs, { type: 'array', bookType: 'xlsx', cellStyles: true });
    libroLector.libroMotor = await LibroXlsx.abrir(new Uint8Array(bytes));
    return libroLector.libroMotor;
  }

  const PAGINA_BASE = { papel: 9, anchoPapel: null, altoPapel: null, orientacion: 'default', escala: 100, ajustarAncho: 1, ajustarAlto: 1, orden: 'downThenOver', primeraPagina: null, usarPrimeraPagina: false, blancoYNegro: false, borrador: false, comentarios: 'none', errores: 'displayed' };

  const celdaA1 = (f, c) => LibroXlsx.numeroACol(c) + (f + 1);
  const comillas = (nombre) => "'" + String(nombre).replace(/'/g, "''") + "'";

  // Aplica las sobreposiciones a una hoja cargada y devuelve cómo deshacerlas.
  function sobreponer(libro, hoja, indice, o) {
    const deshacer = [];
    const imp = hoja.imprimir || (hoja.imprimir = { margenes: null, pagina: null, opciones: null, encabezadoPie: null, saltosFila: [], saltosCol: [], ajustarAPagina: false, autoSaltos: true });
    const antes = { ajustarAPagina: imp.ajustarAPagina, pagina: imp.pagina, opciones: imp.opciones };
    deshacer.push(() => { imp.ajustarAPagina = antes.ajustarAPagina; imp.pagina = antes.pagina; imp.opciones = antes.opciones; });
    imp.pagina = Object.assign({}, imp.pagina || (libro.esCsv ? PAGINA_BASE : Object.assign({}, PAGINA_BASE, { papel: 1 })));

    if (Array.isArray(o.papel)) { imp.pagina.anchoPapel = Math.min(o.papel[0], o.papel[1]) + 'pt'; imp.pagina.altoPapel = Math.max(o.papel[0], o.papel[1]) + 'pt'; }
    else if (o.papel && o.papel !== 'archivo' && Object.prototype.hasOwnProperty.call(PAPELES, o.papel)) { imp.pagina.anchoPapel = PAPELES[o.papel][0] + 'pt'; imp.pagina.altoPapel = PAPELES[o.papel][1] + 'pt'; }
    if (o.orientacion === 'vertical') imp.pagina.orientacion = 'portrait';
    else if (o.orientacion === 'horizontal') imp.pagina.orientacion = 'landscape';

    if (o.ajustar === 'no') { imp.ajustarAPagina = false; imp.pagina.escala = 100; }
    else if (o.ajustar > 0) { imp.ajustarAPagina = true; imp.pagina.ajustarAncho = o.ajustar; imp.pagina.ajustarAlto = 0; }

    imp.opciones = Object.assign({ cuadricula: false, encabezados: false, centrarH: false, centrarV: false }, imp.opciones);
    if (o.cuadricula === 'si') imp.opciones.cuadricula = true; else if (o.cuadricula === 'no') imp.opciones.cuadricula = false;

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
      for (const col of hoja.columnas) if (col.oculta) { col.oculta = false; deshacer.push(() => { col.oculta = true; }); }
      for (const info of hoja.filas.values()) if (info.oculta) { info.oculta = false; deshacer.push(() => { info.oculta = true; }); }
    }
    return () => { for (let i = deshacer.length - 1; i >= 0; i--) deshacer[i](); };
  }

  async function generar(libro, o) {
    await cargarMotor();
    const deshacer = [];
    try {
      for (const i of o.hojas) {
        const hoja = await libro.cargarHoja(i);
        deshacer.push(sobreponer(libro, hoja, i, o.hojas.length === 1 ? o : Object.assign({}, o, { areas: null })));
      }
      return await XlsxPdf.generar(libro, o.hojas, { idioma: o.idioma || 'es', nombreArchivo: o.nombreArchivo, maxPaginas: o.maxPaginas, progreso: o.progreso, numeracion: 'porHoja' });
    } finally {
      for (const d of deshacer) d();
    }
  }

  return { cargarMotor, libroDeMotor, generar, PAPELES };
})();

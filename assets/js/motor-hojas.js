/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Carga bajo demanda de los módulos pesados del motor de Excel (interfaz de hojas): el visor solo necesita DIBUJAR la hoja (sin pdf-lib, que es lo más
// grande), y Excel / CSV → PDF necesita además el PDF. Antes la lista vivía en excel-a-pdf-motor.js; aquí se comparte.
//
//   await MotorHojas.cargarDibujo();   // PowerPoint (objetos y gráficos), geometría, formato condicional, tablas, minigráficos y xlsx-dibujo.js
//   await MotorHojas.cargarPdf();      // lo anterior + pdf-lib, pdf-lienzo.js y xlsx-pdf.js
//   await MotorHojas.cargarScript(url) // un script cualquiera, una sola vez
//
// Requiere vendor.js (rutaVendor) y que los módulos livianos (lector-hojas.js y sus dependencias) ya estén cargados en la página.
const MotorHojas = (() => {
  const base = new URL('./', document.currentScript.src).href;
  const DIBUJO = [
    'pptx-xml.js', 'vendor/ecma-376-geometrias/5ed-2016/geometrias.js', 'emf.js', 'pptx-geometria.js', 'pptx-estilos-tabla.js', 'pptx-modelo.js', 'pptx-texto.js',
    'pptx-grafico-3d.js', 'pptx-grafico-superficie.js', 'pptx-grafico.js', 'pptx-dibujo.js', 'xlsx-objetos.js',
    'xlsx-geometria.js', 'xlsx-fuentes-impresion.js', 'xlsx-paginas.js', 'xlsx-expresion.js', 'xlsx-condicional.js', 'xlsx-iconos.js', 'xlsx-minigraficos.js',
    'xlsx-notas.js', 'xlsx-filtros.js', 'xlsx-tablas-estilos.js', 'xlsx-tablas.js', 'xlsx-dibujo.js',
  ];
  const PDF = ['vendor/pdf-lib/1.17.1/pdf-lib.min.js', 'pdf-lienzo.js', 'xlsx-pdf.js'];

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
  const url = (x) => (x.startsWith('vendor/') ? rutaVendor(x.slice(7)) : base + x);

  async function cargarLista(lista) {
    for (const x of lista) if (!(typeof PDFLib !== 'undefined' && /pdf-lib/.test(x))) await cargarScript(url(x));
  }
  async function cargarDibujo() {
    if (typeof XlsxDibujo !== 'undefined' && typeof XlsxDibujo.dibujarVista === 'function') return;
    await cargarLista(DIBUJO);
  }
  async function cargarPdf() {
    if (typeof XlsxPdf !== 'undefined') return;
    await cargarDibujo();
    await cargarLista(PDF);
  }

  return { cargarDibujo, cargarPdf, cargarScript };
})();

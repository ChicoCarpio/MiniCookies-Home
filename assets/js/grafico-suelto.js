/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Gráficos que no viven en una diapositiva, dibujados con el mismo motor que
// los de PowerPoint (pptx-grafico.js): el de una hoja de Excel (la parte
// xl/charts/chartN.xml de un .xlsx, que es el mismo formato) y el que arma
// "Gráfico desde Excel / CSV" con los datos de una tabla. Reemplaza a Chart.js:
// un gráfico sale igual en la pantalla, en el PNG y en el PDF.
//
// Uso:
//   await GraficoSuelto.cargarMotor();                       // trae el motor la primera vez
//   await GraficoSuelto.dibujarEn(canvas, { zip, rutaGrafico, anchoPt, altoPt, escala });
//   const raiz = GraficoSuelto.xmlDeDatos({ tipo, titulo, categorias, series });
//   await GraficoSuelto.dibujarEn(canvas, { raiz, anchoPt, altoPt, escala });
// Requiere vendor.js (rutaVendor) y, para un .xlsx, ZipPropio.

const GraficoSuelto = (() => {
  // Los módulos propios están al lado de este archivo, en assets/js/.
  const base = new URL('./', document.currentScript.src).href;
  // El motor sin lo propio de las diapositivas, en el orden de
  // SCRIPTS_MOTOR_PPTX (pptx-visual.js; el build comprueba que sigan igual).
  const SCRIPTS_MOTOR = ['pptx-xml.js', 'vendor/ecma-376-geometrias/5ed-2016/geometrias.js', 'pptx-geometria.js', 'pptx-modelo.js', 'pptx-texto.js', 'pptx-grafico-3d.js', 'pptx-grafico-superficie.js', 'pptx-grafico.js', 'pptx-dibujo.js'];
  const EMU_POR_PT = 12700;
  const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

  const cargas = new Map(); // src → promesa: cada script se pide una sola vez
  function cargarScript(src) {
    if (!cargas.has(src) && document.querySelector('script[src="' + src + '"]')) cargas.set(src, Promise.resolve()); // ya está en la página
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
    if (typeof PptxDibujo !== 'undefined' && typeof PptxModelo !== 'undefined') return;
    for (const x of SCRIPTS_MOTOR) await cargarScript(x.startsWith('vendor/') ? rutaVendor(x.slice(7)) : base + x);
  }

  // --- Un gráfico armado con datos -----------------------------------------------------
  // tipo: 'bar' | 'line' | 'pie'; series: [{ nombre, valores }]; categorias: [texto].
  // Devuelve el c:chartSpace, tal como lo escribiría Excel.
  const esc = (t) => String(t == null ? '' : t).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const limpio = (t) => esc(t).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/g, '');
  function cacheTexto(lista) {
    return `<c:strCache><c:ptCount val="${lista.length}"/>${lista.map((v, i) => `<c:pt idx="${i}"><c:v>${limpio(v)}</c:v></c:pt>`).join('')}</c:strCache>`;
  }
  function cacheNumeros(lista) {
    return `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${lista.length}"/>${lista.map((v, i) => (Number.isFinite(v) ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : '')).join('')}</c:numCache>`;
  }
  function xmlDeDatos({ tipo, titulo, categorias, series }) {
    const circular = tipo === 'pie';
    const serie = (s, i) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:f>Hoja1!$${String.fromCharCode(66 + i)}$1</c:f>${cacheTexto([s.nombre])}</c:strRef></c:tx>`
      + (tipo === 'bar' ? '<c:invertIfNegative val="0"/>' : tipo === 'line' ? '<c:marker><c:symbol val="' + (categorias.length > 40 ? 'none' : 'circle') + '"/><c:size val="5"/></c:marker>' : '')
      + `<c:cat><c:strRef><c:f>Hoja1!$A$2:$A$${categorias.length + 1}</c:f>${cacheTexto(categorias)}</c:strRef></c:cat>`
      + `<c:val><c:numRef><c:f>Hoja1!$${String.fromCharCode(66 + i)}$2:$${String.fromCharCode(66 + i)}$${categorias.length + 1}</c:f>${cacheNumeros(s.valores)}</c:numRef></c:val>`
      + (tipo === 'line' ? '<c:smooth val="0"/>' : '') + '</c:ser>';
    const ejes = '<c:axId val="111111"/><c:axId val="222222"/>';
    const grupo = tipo === 'bar' ? `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series.map(serie).join('')}<c:gapWidth val="100"/>${ejes}</c:barChart>`
      : tipo === 'line' ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series.map(serie).join('')}<c:marker val="1"/>${ejes}</c:lineChart>` // con muchos puntos, sin marcadores
      : `<c:pieChart><c:varyColors val="1"/>${series.slice(0, 1).map(serie).join('')}<c:firstSliceAng val="0"/></c:pieChart>`;
    const ejesXml = circular ? '' : '<c:catAx><c:axId val="111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/>'
      + '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="222222"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
      + '<c:valAx><c:axId val="222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/>'
      + '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="111111"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>';
    const tituloXml = titulo ? `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr lang="es"/><a:t>${limpio(titulo)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
    const leyenda = circular || series.length > 1 ? '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>' : '';
    const xml = `<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}"><c:chart>${tituloXml}<c:plotArea><c:layout/>${grupo}${ejesXml}</c:plotArea>${leyenda}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
    return new DOMParser().parseFromString(xml, 'application/xml').documentElement;
  }

  // --- Dibujar ---------------------------------------------------------------------------
  // Pone el gráfico en el canvas, del tamaño que tendría en anchoPt × altoPt
  // puntos, con `escala` píxeles por punto.
  async function dibujarEn(canvas, { zip, rutaGrafico, raiz, anchoPt, altoPt, escala = 2, excel = true }) {
    await cargarMotor();
    const { pres, diapositiva } = await PptxModelo.graficoSuelto({ zip, rutaGrafico, raiz, ancho: Math.round(anchoPt * EMU_POR_PT), alto: Math.round(altoPt * EMU_POR_PT), excel });
    canvas.width = Math.round(anchoPt * escala);
    canvas.height = Math.round(altoPt * escala);
    await PptxDibujo.dibujar(pres, diapositiva, canvas);
  }

  return { cargarMotor, xmlDeDatos, dibujarEn };
})();

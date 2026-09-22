// Helpers compartidos de "texto plano -> PDF" y "extraer texto de un .pptx".
// Requiere que la página ya haya cargado pdf-lib (global PDFLib) y, para
// extraerTextoPowerPoint, JSZip (global JSZip).

async function generarPdfDeTexto(texto) {
  const { PDFDocument, StandardFonts, rgb } = PDFLib;
  const pdf = await PDFDocument.create();
  const fuente = await pdf.embedFont(StandardFonts.Helvetica);
  const tamanoFuente = 11;
  const margen = 50;
  const anchoPagina = 595, altoPagina = 842; // A4
  const anchoUtil = anchoPagina - margen * 2;
  const interlineado = tamanoFuente * 1.4;

  function envolverLinea(linea) {
    const palabras = linea.split(' ');
    const lineas = [];
    let actual = '';
    palabras.forEach((palabra) => {
      const prueba = actual ? actual + ' ' + palabra : palabra;
      if (actual && fuente.widthOfTextAtSize(prueba, tamanoFuente) > anchoUtil) {
        lineas.push(actual);
        actual = palabra;
      } else {
        actual = prueba;
      }
    });
    lineas.push(actual);
    return lineas;
  }

  const lineasFinales = [];
  texto.split('\n').forEach((linea) => envolverLinea(linea).forEach(l => lineasFinales.push(l)));

  let pagina = pdf.addPage([anchoPagina, altoPagina]);
  let y = altoPagina - margen;

  lineasFinales.forEach((linea) => {
    if (y < margen) {
      pagina = pdf.addPage([anchoPagina, altoPagina]);
      y = altoPagina - margen;
    }
    pagina.drawText(linea, { x: margen, y, size: tamanoFuente, font: fuente, color: rgb(0.15, 0.15, 0.15) });
    y -= interlineado;
  });

  return pdf;
}

async function extraerTextoPowerPoint(archivo) {
  const zip = await JSZip.loadAsync(archivo);
  const nombresDiapositivas = Object.keys(zip.files)
    .filter(nombre => /^ppt\/slides\/slide\d+\.xml$/.test(nombre))
    .sort((a, b) => parseInt(a.match(/slide(\d+)/)[1], 10) - parseInt(b.match(/slide(\d+)/)[1], 10));

  if (nombresDiapositivas.length === 0) {
    throw new Error(window.t ? t('ppt-sin-diapositivas') : 'No se encontraron diapositivas legibles en este archivo.');
  }

  const bloques = [];
  for (let i = 0; i < nombresDiapositivas.length; i++) {
    const xmlTexto = await zip.files[nombresDiapositivas[i]].async('string');
    const doc = new DOMParser().parseFromString(xmlTexto, 'application/xml');
    const parrafos = [...doc.getElementsByTagName('a:p')];
    const lineas = parrafos
      .map(p => [...p.getElementsByTagName('a:t')].map(t => t.textContent).join(''))
      .filter(l => l.trim() !== '');
    const tituloDiapositiva = (window.t ? t('ppt-diapositiva-titulo') : '--- Diapositiva {n} ---').replace('{n}', i + 1);
    const sinTexto = window.t ? t('ppt-sin-texto') : '(sin texto)';
    bloques.push(`${tituloDiapositiva}\n` + (lineas.length > 0 ? lineas.join('\n') : sinTexto));
  }
  return bloques.join('\n\n');
}

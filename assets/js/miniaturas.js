/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Helper compartido para mostrar miniaturas reales de PDFs/imágenes en vez de solo el
// nombre del archivo, en una grilla de tarjetas (mismo estilo visual que "Administrar
// páginas"). Requiere que la página ya haya cargado pdf.js (pdfjsLib) para PDFs.

// gradosExtra: rotación adicional a sumarle a la que ya trae la página (para
// previsualizar herramientas como "Rotar PDF" antes de aplicar el cambio de verdad).
async function generarMiniaturaPdf(archivo, escala, gradosExtra) {
  escala = escala || 0.28;
  const bytes = await archivo.arrayBuffer();
  // isEvalSupported: false mitiga CVE-2024-4367 (un PDF armado a propósito podía
  // ejecutar JavaScript al abrirse) mientras se use PDF.js 3.x.
  const pdf = await pdfjsLib.getDocument({ data: bytes, isEvalSupported: false }).promise;
  const pagina = await pdf.getPage(1);
  const rotacionTotal = (pagina.rotate + (gradosExtra || 0)) % 360;
  const viewport = pagina.getViewport({ scale: escala, rotation: rotacionTotal });
  const canvas = document.createElement('canvas');
  canvas.width = viewport.width;
  canvas.height = viewport.height;
  await pagina.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
  return { elemento: canvas, totalPaginas: pdf.numPages };
}

// --- Imágenes → páginas de PDF (compartido por unir-pdf.html y pdf-imagenes.html) ---

// Orientación EXIF de un JPEG (1 = normal, 2-8 = girada/espejada). Casi todas
// las fotos de celular se guardan "acostadas" con esta etiqueta indicando
// cómo girarlas: el navegador la aplica al mostrarlas (por eso la miniatura
// se ve bien), pero pdf-lib incrusta el JPEG tal cual y en el PDF la foto
// quedaba girada 90°.
function leerOrientacionJpeg(bytes) {
  const vista = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (vista.byteLength < 4 || vista.getUint16(0) !== 0xFFD8) return 1;
  let offset = 2;
  while (offset + 4 <= vista.byteLength) {
    const marcador = vista.getUint16(offset);
    if (marcador === 0xFFDA || (marcador & 0xFF00) !== 0xFF00) return 1;
    const largo = vista.getUint16(offset + 2);
    if (marcador === 0xFFE1 && offset + 18 <= vista.byteLength && vista.getUint32(offset + 4) === 0x45786966) {
      const tiff = offset + 10;
      const little = vista.getUint16(tiff) === 0x4949;
      const ifd0 = tiff + vista.getUint32(tiff + 4, little);
      if (ifd0 + 2 > vista.byteLength) return 1;
      const cantidad = vista.getUint16(ifd0, little);
      for (let i = 0; i < cantidad; i++) {
        const entrada = ifd0 + 2 + i * 12;
        if (entrada + 10 > vista.byteLength) return 1;
        if (vista.getUint16(entrada, little) === 0x0112) return vista.getUint16(entrada + 8, little) || 1;
      }
      return 1;
    }
    offset += 2 + largo;
  }
  return 1;
}

function cargarImagenElemento(archivo) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = reject;
    img.src = URL.createObjectURL(archivo);
  });
}

// Incrusta una imagen en un PDFDocument de pdf-lib. JPEG sin rotación y PNG
// van directo (sin recomprimir). El resto —JPEG con orientación EXIF, WebP,
// GIF, BMP, etc., que pdf-lib no sabe incrustar— se redibuja en un canvas (el
// navegador ya aplica la orientación al dibujar) y se reincrusta: las fotos
// como JPEG de alta calidad, para no inflar el PDF, y lo demás como PNG para
// conservar la transparencia.
async function incrustarImagenEnPdf(pdf, archivo) {
  const bytes = new Uint8Array(await archivo.arrayBuffer());
  const esJpeg = bytes[0] === 0xFF && bytes[1] === 0xD8;
  const esPng = bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4E && bytes[3] === 0x47;
  if (esJpeg && leerOrientacionJpeg(bytes) === 1) return pdf.embedJpg(bytes);
  if (esPng) return pdf.embedPng(bytes);
  const img = await cargarImagenElemento(archivo);
  const canvas = document.createElement('canvas');
  canvas.width = img.naturalWidth;
  canvas.height = img.naturalHeight;
  canvas.getContext('2d').drawImage(img, 0, 0);
  URL.revokeObjectURL(img.src);
  const formato = esJpeg ? 'image/jpeg' : 'image/png';
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, formato, 0.92));
  const bytesNuevos = new Uint8Array(await blob.arrayBuffer());
  return esJpeg ? pdf.embedJpg(bytesNuevos) : pdf.embedPng(bytesNuevos);
}

// Tamaños de página estándar en puntos: A4 = 210x297 mm, Carta = 8.5x11 pulgadas.
const TAMANOS_PAGINA_PDF = {
  a4: [595.276, 841.89],
  carta: [612, 792],
};

// Agrega una página con la imagen ya incrustada, según el tamaño elegido:
// "original" (la página mide lo que la imagen), "ajustar-a4"/"ajustar-carta"
// (ancho estándar y alto según la proporción: sin recorte ni márgenes) o
// "a4"/"carta" exactos (imagen completa centrada, con márgenes blancos).
function agregarPaginaConImagen(pdf, imagenEmbebida, tamanoElegido) {
  const { width, height } = imagenEmbebida.size();
  if (tamanoElegido === 'original') {
    pdf.addPage([width, height]).drawImage(imagenEmbebida, { x: 0, y: 0, width, height });
    return;
  }
  if (tamanoElegido.startsWith('ajustar-')) {
    // El ancho es siempre el del lado corto (vertical), también para fotos
    // apaisadas: así todas las páginas quedan del mismo ancho y el documento
    // se lee parejo al desplazarse (lo mismo que hace PDF24).
    const [anchoPagina] = TAMANOS_PAGINA_PDF[tamanoElegido.slice('ajustar-'.length)];
    const altoPagina = anchoPagina * (height / width);
    pdf.addPage([anchoPagina, altoPagina]).drawImage(imagenEmbebida, { x: 0, y: 0, width: anchoPagina, height: altoPagina });
    return;
  }
  // Página estándar exacta, en la orientación que mejor se ajuste, con la
  // imagen escalada para entrar COMPLETA (Math.min: nunca se recorta nada).
  const [anchoBase, altoBase] = TAMANOS_PAGINA_PDF[tamanoElegido];
  const [anchoPagina, altoPagina] = width > height ? [altoBase, anchoBase] : [anchoBase, altoBase];
  const escala = Math.min(anchoPagina / width, altoPagina / height);
  const anchoFinal = width * escala;
  const altoFinal = height * escala;
  pdf.addPage([anchoPagina, altoPagina]).drawImage(imagenEmbebida, {
    x: (anchoPagina - anchoFinal) / 2,
    y: (altoPagina - altoFinal) / 2,
    width: anchoFinal,
    height: altoFinal,
  });
}

function generarMiniaturaImagen(archivo) {
  return new Promise((resolve, reject) => {
    const img = document.createElement('img');
    img.onload = () => resolve({ elemento: img, totalPaginas: null });
    img.onerror = reject;
    img.src = URL.createObjectURL(archivo);
  });
}

// Crea una tarjeta de miniatura estilo grilla (misma clase .miniatura-pagina que ya
// usa administrar-paginas.html). options: { elementoVisual, etiqueta, conQuitar,
// alQuitar, arrastrable, indiceOrden }
function crearTarjetaMiniatura({ elementoVisual, etiqueta, conQuitar, alQuitar, arrastrable, indiceOrden }) {
  const tarjeta = document.createElement('div');
  tarjeta.className = 'miniatura-pagina';
  if (arrastrable) tarjeta.draggable = true;
  if (indiceOrden !== undefined) tarjeta.dataset.indiceOrden = indiceOrden;

  if (elementoVisual) {
    tarjeta.appendChild(elementoVisual);
  } else {
    const marcador = document.createElement('div');
    marcador.textContent = '📄';
    marcador.style.fontSize = '2rem';
    marcador.style.padding = '20px 0';
    tarjeta.appendChild(marcador);
  }

  const etiquetaEl = document.createElement('div');
  etiquetaEl.className = 'numero-pagina';
  etiquetaEl.textContent = etiqueta;
  tarjeta.appendChild(etiquetaEl);

  if (conQuitar) {
    const botonQuitar = document.createElement('button');
    botonQuitar.className = 'quitar-pagina';
    botonQuitar.type = 'button';
    botonQuitar.textContent = '✕';
    botonQuitar.title = window.t ? t('quitar') : 'Quitar';
    botonQuitar.addEventListener('click', () => alQuitar && alQuitar());
    tarjeta.appendChild(botonQuitar);
  }

  return tarjeta;
}

// Crea el par de botones "Lista"/"Cuadrícula" que alternan la clase vista-lista /
// vista-grid en el contenedor. Recuerda la preferencia del usuario en localStorage
// (por nombre de página) para que no tenga que elegir cada vez.
function crearSelectorVista(contenedor, claveAlmacenamiento) {
  const selector = document.createElement('div');
  selector.className = 'selector-vista';

  const botonGrid = document.createElement('button');
  botonGrid.type = 'button';
  botonGrid.textContent = window.t ? t('vista-cuadricula') : '▦ Cuadrícula';

  const botonLista = document.createElement('button');
  botonLista.type = 'button';
  botonLista.textContent = window.t ? t('vista-lista') : '☰ Lista';

  document.addEventListener('cambioIdioma', function () {
    botonGrid.textContent = t('vista-cuadricula');
    botonLista.textContent = t('vista-lista');
  });

  function aplicarVista(vista) {
    contenedor.classList.remove('vista-grid', 'vista-lista');
    contenedor.classList.add(vista === 'lista' ? 'vista-lista' : 'vista-grid');
    botonGrid.classList.toggle('activo', vista !== 'lista');
    botonLista.classList.toggle('activo', vista === 'lista');
    if (claveAlmacenamiento) {
      try { localStorage.setItem(claveAlmacenamiento, vista); } catch (error) { /* almacenamiento no disponible, no es crítico */ }
    }
  }

  botonGrid.addEventListener('click', () => aplicarVista('grid'));
  botonLista.addEventListener('click', () => aplicarVista('lista'));

  selector.appendChild(botonGrid);
  selector.appendChild(botonLista);

  let vistaGuardada = 'grid';
  if (claveAlmacenamiento) {
    try { vistaGuardada = localStorage.getItem(claveAlmacenamiento) || 'grid'; } catch (error) { /* usa el default */ }
  }
  aplicarVista(vistaGuardada);

  return selector;
}

// Habilita arrastrar-y-soltar para reordenar tarjetas dentro de un contenedor grid,
// llamando a alReordenar() cada vez que cambia el orden en el DOM.
function habilitarArrastreReordenable(contenedor, alReordenar) {
  contenedor.addEventListener('dragstart', (e) => {
    if (e.target.classList.contains('miniatura-pagina')) e.target.classList.add('arrastrando');
  });
  contenedor.addEventListener('dragend', (e) => {
    if (e.target.classList.contains('miniatura-pagina')) e.target.classList.remove('arrastrando');
  });
  contenedor.addEventListener('dragover', (e) => {
    e.preventDefault();
    const tarjeta = e.target.closest('.miniatura-pagina');
    const arrastrando = contenedor.querySelector('.arrastrando');
    if (!tarjeta || !arrastrando || arrastrando === tarjeta) return;
    // insertarSegunPosicion (assets/js/reordenar-tactil.js) compara la X en
    // la vista de cuadrícula y la Y en la vista de lista (antes siempre la
    // X, y en la vista de lista el orden quedaba mal).
    if (insertarSegunPosicion(arrastrando, tarjeta, e.clientX, e.clientY)) alReordenar();
  });
  habilitarReordenTactil(contenedor, '.miniatura-pagina', alReordenar);
}

// --- Coordenadas "tal como se ve" en una página de pdf-lib ---
// Las páginas escaneadas (o pasadas por "Rotar PDF") tienen /Rotate: el
// contenido se muestra girado, pero drawText trabaja en las coordenadas sin
// girar, así que un número "abajo a la derecha" quedaba de costado y en otra
// esquina. Esta función da el ancho/alto visibles (según la CropBox, que es
// lo que muestran los visores) y convierte un punto visible (origen abajo a
// la izquierda de lo que se ve) a coordenadas de la página. "rotacion" es
// lo que hay que sumarle al ángulo del texto para que se lea derecho.
function marcoVisualPagina(pagina) {
  const caja = pagina.getCropBox();
  const rotacion = (((pagina.getRotation().angle || 0) % 360) + 360) % 360;
  const girada = rotacion === 90 || rotacion === 270;
  const ancho = girada ? caja.height : caja.width;
  const alto = girada ? caja.width : caja.height;
  function aPagina(vx, vy) {
    let x = vx;
    let y = vy;
    if (rotacion === 90) { x = caja.width - vy; y = vx; }
    else if (rotacion === 180) { x = caja.width - vx; y = caja.height - vy; }
    else if (rotacion === 270) { x = vy; y = caja.height - vx; }
    return { x: caja.x + x, y: caja.y + y };
  }
  return { ancho, alto, rotacion, aPagina };
}

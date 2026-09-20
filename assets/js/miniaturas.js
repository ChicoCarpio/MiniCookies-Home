// Helper compartido para mostrar miniaturas reales de PDFs/imágenes en vez de solo el
// nombre del archivo, en una grilla de tarjetas (mismo estilo visual que "Administrar
// páginas"). Requiere que la página ya haya cargado pdf.js (pdfjsLib) para PDFs.

// gradosExtra: rotación adicional a sumarle a la que ya trae la página (para
// previsualizar herramientas como "Rotar PDF" antes de aplicar el cambio de verdad).
async function generarMiniaturaPdf(archivo, escala, gradosExtra) {
  escala = escala || 0.28;
  const bytes = await archivo.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;
  const pagina = await pdf.getPage(1);
  const rotacionTotal = (pagina.rotate + (gradosExtra || 0)) % 360;
  const viewport = pagina.getViewport({ scale: escala, rotation: rotacionTotal });
  const canvas = document.createElement('canvas');
  canvas.width = viewport.width;
  canvas.height = viewport.height;
  await pagina.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
  return { elemento: canvas, totalPaginas: pdf.numPages };
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
    botonQuitar.title = 'Quitar';
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
  botonGrid.textContent = '▦ Cuadrícula';

  const botonLista = document.createElement('button');
  botonLista.type = 'button';
  botonLista.textContent = '☰ Lista';

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
    const rectangulo = tarjeta.getBoundingClientRect();
    const despuesDelMedio = e.clientX > rectangulo.left + rectangulo.width / 2;
    tarjeta.parentNode.insertBefore(arrastrando, despuesDelMedio ? tarjeta.nextSibling : tarjeta);
    alReordenar();
  });
}

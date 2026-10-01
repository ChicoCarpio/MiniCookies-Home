/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Generador de stickers: recorta el sujeto de una foto con un modelo de
// segmentación (ISNet, MIT) que corre 100% en el navegador con transformers.js,
// deja elegir uno de los filtros del Editor de color (assets/js/filtros-color.js
// + assets/js/procesar-filtros.js, mismo motor que los editores de imagen) y
// arma un sticker con borde de color alrededor del recorte. La foto nunca
// sale del navegador: solo el modelo de recorte (un archivo estático, ~42 MB)
// se descarga una vez desde la CDN de Hugging Face y queda cacheado.
// Requiere, cargados antes como <script>: ajuste-por-color.js, filtros-color.js,
// efectos-textura.js, procesar-filtros.js. El panel lateral usa el mismo
// esqueleto de 2 columnas que Editor-de-color-de-imagen.html (clase
// .editor-color-layout, sin el panel plegable de ese editor: acá el panel
// izquierdo queda siempre visible), pero con tres acordeones independientes
// ("Filtros", "Color del borde", "Acciones") en vez de un panel único: son
// herramientas distintas, y así cada una tiene su propia categoría en la
// barra inferior de celular (assets/js/app-movil.js), en vez de quedar
// escondidas dentro de "Filtros". El grosor del borde sigue debajo del
// canvas, en el panel del lienzo, porque se ajusta mirando el resultado.

import { pipeline, env } from 'https://cdn.jsdelivr.net/npm/@huggingface/transformers@4.3.0/dist/transformers.min.js';

const tt = (clave, porDefecto) => (window.t ? window.t(clave) : porDefecto);

env.allowLocalModels = false;

const MODELO_ID = 'xrds/isnet-general-onnx-int8';
const LADO_MAXIMO = 512; // tamaño estándar de un sticker (WhatsApp/Telegram)

export function iniciarGeneradorStickers() {
  const zonaArrastre = document.getElementById('zonaArrastre');
  const entradaArchivo = document.getElementById('entradaArchivo');
  const panelProgreso = document.getElementById('panelProgreso');
  const barraProgreso = document.getElementById('barraProgresoRelleno');
  const textoProgreso = document.getElementById('textoProgreso');
  const mensajeError = document.getElementById('mensajeError');
  const editorCompleto = document.getElementById('editorCompleto');
  const entradaGrosor = document.getElementById('entradaGrosor');
  const etiquetaGrosor = document.getElementById('etiquetaGrosor');
  const entradaColorBorde = document.getElementById('entradaColorBorde');
  const entradaIntensidad = document.getElementById('intensidadFiltro');
  const valorIntensidad = document.getElementById('valorIntensidadFiltro');
  const controlesTextura = document.getElementById('controlesTextura');
  const entradaTextura = document.getElementById('texturaIntensidad');
  const valorTextura = document.getElementById('valorTexturaIntensidad');
  const contenedorPresets = document.getElementById('presets');
  const vistaResultado = document.getElementById('vistaResultado');
  const contextoResultado = vistaResultado.getContext('2d');
  const botonDescargar = document.getElementById('botonDescargar');

  // La grilla usa columnas "minmax(72px, 1fr)": en un panel angosto con pocas
  // miniaturas por fila, cada columna crece bastante más allá de esos 72px
  // mínimos. Dibujar la miniatura a 64px y dejar que el CSS la estire para
  // llenar la columna la hacía ver borrosa/pixelada — se dibuja a un tamaño
  // base más grande (160px) que cubre el ancho real de columna más común,
  // multiplicado además por devicePixelRatio para pantallas de alta densidad.
  const TAMANO_MINIATURA = 160;
  const TAMANO_MINIATURA_FISICO = Math.round(TAMANO_MINIATURA * (window.devicePixelRatio || 1));
  FiltrosColor.crearGrilla(contenedorPresets);

  let segmentador = null;
  let cargaPromesa = null;
  let canvasBase = null; // recorte limpio, sin filtro, ya escalado a LADO_MAXIMO
  let canvasFiltrado = null; // canvasBase con el filtro elegido ya aplicado
  let presetActual = 'ninguno';

  function mostrarError(texto) {
    mensajeError.textContent = '❌ ' + texto;
    mensajeError.style.display = 'block';
  }
  function ocultarError() {
    mensajeError.style.display = 'none';
  }

  // Carga (o reusa) el segmentador. Igual que en dictado-voz.js: si el
  // navegador tiene WebGPU se intenta ahí primero (más rápido, más memoria
  // disponible) y si falla se reintenta por WASM (CPU).
  function cargarSegmentador() {
    if (segmentador) return Promise.resolve(segmentador);
    if (cargaPromesa) return cargaPromesa;

    panelProgreso.style.display = 'block';
    textoProgreso.textContent = tt('estado-preparando', 'Preparando el modelo...');
    barraProgreso.style.width = '0%';

    const archivos = {};
    const progress_callback = (info) => {
      if (info.status === 'progress' && info.file) {
        archivos[info.file] = info;
        const valores = Object.values(archivos);
        const cargado = valores.reduce((a, f) => a + (f.loaded || 0), 0);
        const total = valores.reduce((a, f) => a + (f.total || 0), 0);
        const porcentaje = total ? Math.round((cargado / total) * 100) : 0;
        barraProgreso.style.width = porcentaje + '%';
        textoProgreso.textContent = tt('estado-descargando',
          'Descargando el modelo... {p}% (solo esta vez, después queda guardado)').replace('{p}', porcentaje);
      }
    };

    const intentarCargar = (device) => pipeline('background-removal', MODELO_ID, {
      dtype: 'q8',
      ...(device ? { device } : {}),
      progress_callback,
    });

    const promesa = (typeof navigator !== 'undefined' && navigator.gpu)
      ? intentarCargar('webgpu').catch(() => intentarCargar())
      : intentarCargar();

    cargaPromesa = promesa.then((nuevo) => {
      segmentador = nuevo;
      panelProgreso.style.display = 'none';
      return segmentador;
    }).catch((error) => {
      cargaPromesa = null;
      throw error;
    });

    return cargaPromesa;
  }

  // Reduce el recorte a un máximo de LADO_MAXIMO px de lado (tamaño típico
  // de sticker) para que el grosor de borde elegido se vea proporcionado —
  // una foto de 4000px con un borde de 12px se vería como si no tuviera nada.
  function escalarRecorte(canvasOrigen) {
    const factor = Math.min(1, LADO_MAXIMO / Math.max(canvasOrigen.width, canvasOrigen.height));
    const ancho = Math.max(1, Math.round(canvasOrigen.width * factor));
    const alto = Math.max(1, Math.round(canvasOrigen.height * factor));
    const destino = document.createElement('canvas');
    destino.width = ancho;
    destino.height = alto;
    destino.getContext('2d').drawImage(canvasOrigen, 0, 0, ancho, alto);
    return destino;
  }

  // Segunda pasada de limpieza: el modelo a veces deja "manchas" — zonas
  // chicas con opacidad parcial, sueltas del sujeto principal (más notorio
  // con la variante cuantizada, más liviana pero más ruidosa). Se agrupan
  // los píxeles no transparentes en islas conectadas (relleno por
  // inundación, 4 vecinos) y se borran las islas chicas comparadas con la
  // isla más grande (el sujeto). Se conservan islas sueltas razonablemente
  // grandes (ej. un brazo separado del cuerpo) para no comerse partes
  // legítimas del recorte.
  function limpiarMascara(canvas) {
    const ctx = canvas.getContext('2d');
    const { width, height } = canvas;
    const total = width * height;
    const imageData = ctx.getImageData(0, 0, width, height);
    const datos = imageData.data;

    const UMBRAL_ALFA = 20;
    const visitado = new Uint8Array(total);
    const etiquetas = new Int32Array(total).fill(-1);
    const tamanos = [];
    const pila = new Int32Array(total);

    for (let inicio = 0; inicio < total; inicio++) {
      if (visitado[inicio] || datos[inicio * 4 + 3] < UMBRAL_ALFA) { visitado[inicio] = 1; continue; }
      const etiqueta = tamanos.length;
      let tope = 0;
      pila[tope++] = inicio;
      visitado[inicio] = 1;
      let tamano = 0;
      while (tope > 0) {
        const p = pila[--tope];
        etiquetas[p] = etiqueta;
        tamano++;
        const x = p % width;
        const y = (p / width) | 0;
        if (x > 0 && !visitado[p - 1] && datos[(p - 1) * 4 + 3] >= UMBRAL_ALFA) { visitado[p - 1] = 1; pila[tope++] = p - 1; }
        if (x < width - 1 && !visitado[p + 1] && datos[(p + 1) * 4 + 3] >= UMBRAL_ALFA) { visitado[p + 1] = 1; pila[tope++] = p + 1; }
        if (y > 0 && !visitado[p - width] && datos[(p - width) * 4 + 3] >= UMBRAL_ALFA) { visitado[p - width] = 1; pila[tope++] = p - width; }
        if (y < height - 1 && !visitado[p + width] && datos[(p + width) * 4 + 3] >= UMBRAL_ALFA) { visitado[p + width] = 1; pila[tope++] = p + width; }
      }
      tamanos.push(tamano);
    }

    if (tamanos.length === 0) return;
    const tamanoMaximo = Math.max(...tamanos);
    const UMBRAL_RELATIVO = 0.08;
    const etiquetasValidas = new Set();
    tamanos.forEach((tamano, etiqueta) => { if (tamano >= tamanoMaximo * UMBRAL_RELATIVO) etiquetasValidas.add(etiqueta); });

    let seModificoAlgo = false;
    for (let i = 0; i < total; i++) {
      if (datos[i * 4 + 3] > 0 && !etiquetasValidas.has(etiquetas[i])) {
        datos[i * 4 + 3] = 0;
        seModificoAlgo = true;
      }
    }
    if (seModificoAlgo) ctx.putImageData(imageData, 0, 0);
  }

  // Encaja canvasBase dentro de un cuadrado de "lado" px sin recortarlo (a
  // diferencia de crearBaseMiniatura() del Editor de color, que recorta al
  // centro): un sticker ya viene ajustado al sujeto, y recortarlo de nuevo
  // para la miniatura podría comerse una parte.
  function crearMiniaturaSticker(canvas, lado) {
    const factor = Math.min(lado / canvas.width, lado / canvas.height);
    const ancho = Math.max(1, Math.round(canvas.width * factor));
    const alto = Math.max(1, Math.round(canvas.height * factor));
    const destino = document.createElement('canvas');
    destino.width = lado;
    destino.height = lado;
    destino.getContext('2d').drawImage(canvas, (lado - ancho) / 2, (lado - alto) / 2, ancho, alto);
    return destino;
  }

  // Algunas texturas (ej. los puntos de "Pop art", pixelar-circulos en
  // efectos-textura.js) dibujan círculos con ctx.fill() sin fijarse en la
  // transparencia de lo que pisan — ahí aparecían puntos sobre el fondo
  // vacío del sticker. Volver a poner el canal alfa ORIGINAL después de
  // cualquier filtro garantiza que el fondo se quede transparente pase lo
  // que pase adentro del filtro, sin tener que auditar cada textura.
  function restaurarAlfa(canvas, alfaOriginal) {
    const ctx = canvas.getContext('2d');
    const datos = ctx.getImageData(0, 0, canvas.width, canvas.height);
    for (let i = 3; i < datos.data.length; i += 4) datos.data[i] = alfaOriginal.data[i];
    ctx.putImageData(datos, 0, 0);
  }

  // Dibuja, sobre el <canvas> de cada botón de la grilla, cómo quedaría el
  // sticker con ese filtro — se calcula una sola vez por recorte nuevo (los
  // presets no cambian). Son ~50 filtros: procesarlos todos de un tirón
  // congela la pantalla justo después de "Recortando el fondo...", así que
  // se hacen de a lotes en el tiempo libre del navegador (mismo mecanismo que
  // editor-avanzado-imagen.html usa para su propia grilla de filtros).
  let generacionMiniaturasPresets = 0;

  function actualizarMiniaturasPresets() {
    if (!canvasBase) return;
    const generacion = ++generacionMiniaturasPresets;
    const origen = canvasBase;
    const botones = [...contenedorPresets.querySelectorAll('.preset-miniatura')];
    const escala = TAMANO_MINIATURA_FISICO / Math.max(origen.width, origen.height);
    let indice = 0;

    function procesarLote(plazo) {
      if (generacion !== generacionMiniaturasPresets || origen !== canvasBase) return;
      const inicio = performance.now();
      let procesadas = 0;
      while (indice < botones.length && (!procesadas || performance.now() - inicio < 8)) {
        if (plazo && !plazo.didTimeout && procesadas && plazo.timeRemaining() < 2) break;
        const boton = botones[indice++];
        const preset = ProcesarFiltros.PRESETS[boton.dataset.preset];
        procesadas++;
        if (!preset) continue;
        const miniatura = crearMiniaturaSticker(origen, TAMANO_MINIATURA_FISICO);
        const alfaOriginal = miniatura.getContext('2d').getImageData(0, 0, miniatura.width, miniatura.height);
        ProcesarFiltros.renderizarFiltro(miniatura, preset, preset.capasExtra, preset.textura || ProcesarFiltros.SIN_TEXTURA, escala);
        restaurarAlfa(miniatura, alfaOriginal);
        const canvasDestino = boton.querySelector('canvas');
        canvasDestino.width = TAMANO_MINIATURA_FISICO;
        canvasDestino.height = TAMANO_MINIATURA_FISICO;
        canvasDestino.getContext('2d').drawImage(miniatura, 0, 0);
      }

      if (indice >= botones.length) return;
      if (window.requestIdleCallback) window.requestIdleCallback(procesarLote, { timeout: 150 });
      else window.requestAnimationFrame(() => procesarLote(null));
    }

    if (window.requestIdleCallback) window.requestIdleCallback(procesarLote, { timeout: 150 });
    else window.requestAnimationFrame(() => procesarLote(null));
  }

  // Aplica el filtro elegido (assets/js/filtros-color.js, vía
  // assets/js/procesar-filtros.js) sobre una COPIA de canvasBase — nunca
  // sobre el original — para poder cambiar de filtro sin volver a correr el
  // modelo de recorte ni perder la limpieza de manchas ya hecha. La
  // intensidad mezcla el resultado con la foto sin filtrar (100% = filtro
  // completo), igual que en el Editor de color — no afecta a las miniaturas,
  // que siempre muestran cada filtro a pleno para poder compararlos.
  function actualizarFiltro() {
    if (!canvasBase) return;
    const preset = ProcesarFiltros.PRESETS[presetActual];
    if (!preset || presetActual === 'ninguno') {
      canvasFiltrado = canvasBase;
    } else {
      const copia = ProcesarFiltros.copiarCanvas(canvasBase);
      const alfaOriginal = copia.getContext('2d').getImageData(0, 0, copia.width, copia.height);
      const ajustes = { ...preset, intensidadFiltro: Number(entradaIntensidad.value) };
      // El tamaño de bloque de una textura (pixelar/puntos) viene calibrado
      // para fotos de resolución normal — en un sticker (máx. 512px) el mismo
      // número de píxeles se ve enorme en proporción, por eso queda ajustable
      // en vez de usar siempre el valor fijo del preset.
      const textura = preset.textura
        ? { ...preset.textura, intensidad: Number(entradaTextura.value) }
        : ProcesarFiltros.SIN_TEXTURA;
      ProcesarFiltros.renderizarFiltro(copia, ajustes, preset.capasExtra, textura, 1);
      restaurarAlfa(copia, alfaOriginal);
      canvasFiltrado = copia;
    }
    dibujarSticker();
  }

  // Agranda una silueta hacia afuera "estampándola" desplazada en varios
  // ángulos alrededor de un círculo — truco de canvas clásico para lograr un
  // contorno grueso sin depender de ninguna librería nueva. Con suficientes
  // pasos, la zona pegada al recorte queda cubierta por muchas copias
  // superpuestas (borde sólido) y el límite exterior sale con un
  // antialiasing suave solo.
  function agrandarSilueta(canvasOrigen, radio) {
    const salida = document.createElement('canvas');
    salida.width = canvasOrigen.width;
    salida.height = canvasOrigen.height;
    const ctx = salida.getContext('2d');
    const pasos = Math.min(48, Math.max(24, Math.round(radio * 2)));
    for (let i = 0; i < pasos; i++) {
      const angulo = (i / pasos) * Math.PI * 2;
      ctx.drawImage(canvasOrigen, Math.round(Math.cos(angulo) * radio), Math.round(Math.sin(angulo) * radio));
    }
    return salida;
  }

  function dibujarSticker() {
    if (!canvasFiltrado) return;
    const grosor = parseInt(entradaGrosor.value, 10) || 0;
    const margen = grosor + 4;
    const ancho = canvasFiltrado.width + margen * 2;
    const alto = canvasFiltrado.height + margen * 2;

    vistaResultado.width = ancho;
    vistaResultado.height = alto;
    contextoResultado.clearRect(0, 0, ancho, alto);

    if (grosor > 0) {
      const silueta = document.createElement('canvas');
      silueta.width = canvasFiltrado.width;
      silueta.height = canvasFiltrado.height;
      const ctxSil = silueta.getContext('2d');
      ctxSil.fillStyle = entradaColorBorde.value;
      ctxSil.fillRect(0, 0, silueta.width, silueta.height);
      ctxSil.globalCompositeOperation = 'destination-in';
      ctxSil.drawImage(canvasFiltrado, 0, 0);

      const borde = agrandarSilueta(silueta, grosor);
      contextoResultado.drawImage(borde, margen, margen);
    }

    contextoResultado.drawImage(canvasFiltrado, margen, margen);
    vistaResultado.style.display = 'block';
  }

  async function procesarArchivo(archivo) {
    ocultarError();
    vistaResultado.style.display = 'none';
    editorCompleto.style.display = 'none';
    canvasBase = null;
    canvasFiltrado = null;
    const url = URL.createObjectURL(archivo);
    try {
      const modelo = await cargarSegmentador();
      panelProgreso.style.display = 'block';
      barraProgreso.style.width = '100%';
      textoProgreso.textContent = tt('estado-recortando', 'Recortando el fondo...');

      const recorte = await modelo(url);
      canvasBase = escalarRecorte(recorte.toCanvas());
      limpiarMascara(canvasBase);
      editorCompleto.style.display = 'block';
      presetActual = 'ninguno';
      entradaIntensidad.value = 100;
      valorIntensidad.textContent = '100%';
      controlesTextura.style.display = 'none';
      contenedorPresets.querySelectorAll('.preset-miniatura.activo').forEach((b) => b.classList.remove('activo'));
      contenedorPresets.querySelector('[data-preset="ninguno"]')?.classList.add('activo');
      actualizarMiniaturasPresets();
      actualizarFiltro();
    } catch (error) {
      const sinMemoria = /bad_alloc|memory|out of memory|OrtRun/i.test(error.message);
      mostrarError(sinMemoria
        ? tt('error-sin-memoria', 'Tu navegador se quedó sin memoria para correr este modelo.')
        : tt('error-recortar', 'No se pudo recortar la imagen: ') + error.message);
    } finally {
      URL.revokeObjectURL(url);
      panelProgreso.style.display = 'none';
    }
  }

  function elegirArchivo(listaFileList) {
    const archivo = [...listaFileList].find((f) => f.type.startsWith('image/'));
    if (archivo) procesarArchivo(archivo);
  }

  zonaArrastre.addEventListener('click', () => entradaArchivo.click());
  document.getElementById('botonCambiarImagen').addEventListener('click', () => entradaArchivo.click());
  entradaArchivo.addEventListener('change', (e) => { elegirArchivo(e.target.files); entradaArchivo.value = ''; });
  zonaArrastre.addEventListener('dragover', (e) => { e.preventDefault(); zonaArrastre.classList.add('activa'); });
  zonaArrastre.addEventListener('dragleave', () => zonaArrastre.classList.remove('activa'));
  zonaArrastre.addEventListener('drop', (e) => {
    e.preventDefault();
    zonaArrastre.classList.remove('activa');
    elegirArchivo(e.dataTransfer.files);
  });

  entradaGrosor.addEventListener('input', () => {
    etiquetaGrosor.textContent = entradaGrosor.value;
    dibujarSticker();
  });
  entradaColorBorde.addEventListener('input', dibujarSticker);

  entradaIntensidad.addEventListener('input', () => {
    valorIntensidad.textContent = entradaIntensidad.value + '%';
    actualizarFiltro();
  });

  entradaTextura.addEventListener('input', () => {
    valorTextura.textContent = entradaTextura.value;
    actualizarFiltro();
  });

  contenedorPresets.addEventListener('click', (e) => {
    const boton = e.target.closest('.preset-miniatura');
    if (!boton) return;
    contenedorPresets.querySelectorAll('.preset-miniatura.activo').forEach((b) => b.classList.remove('activo'));
    boton.classList.add('activo');
    presetActual = boton.dataset.preset;
    entradaIntensidad.value = 100;
    valorIntensidad.textContent = '100%';
    const preset = ProcesarFiltros.PRESETS[presetActual];
    if (preset && preset.textura) {
      entradaTextura.value = preset.textura.intensidad;
      valorTextura.textContent = preset.textura.intensidad;
      controlesTextura.style.display = 'block';
    } else {
      controlesTextura.style.display = 'none';
    }
    actualizarFiltro();
  });

  botonDescargar.addEventListener('click', async () => {
    const blob = await new Promise((resolve) => vistaResultado.toBlob(resolve, 'image/png'));
    const enlace = document.createElement('a');
    enlace.href = URL.createObjectURL(blob);
    enlace.download = 'sticker.png';
    enlace.click();
  });
}

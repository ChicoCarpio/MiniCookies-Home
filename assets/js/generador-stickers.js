/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Generador de stickers: recorta el sujeto de una foto con un modelo de
// segmentación (ISNet, MIT) que corre 100% en el navegador con transformers.js,
// deja elegir uno de los filtros del Editor de color (assets/js/filtros-color.js
// + assets/js/procesar-filtros.js, mismo motor que los editores de imagen) y
// arma un sticker con borde de color alrededor del recorte. La foto nunca
// sale del navegador: solo el modelo de recorte (un archivo estático, ~42 MB)
// se descarga una vez desde la CDN de Hugging Face y queda cacheado.
// Requiere, cargados antes como <script>: ajuste-por-color.js, filtros-color.js,
// efectos-textura.js, paleta-color.js, vendor/imagetracerjs, efectos-mosaico.js,
// procesar-filtros.js, mascaras-capa.js, borde-sticker.js. El panel lateral usa el mismo
// esqueleto de 2 columnas que Editor-de-color-de-imagen.html (clase
// .editor-color-layout, sin el panel plegable de ese editor: acá el panel
// izquierdo queda siempre visible), pero con cuatro acordeones independientes
// ("Filtros", "Texturas", "Color del borde", "Acciones") en vez de un panel único: son
// herramientas distintas, y así cada una tiene su propia categoría en la
// barra inferior de celular (assets/js/app-movil.js), en vez de quedar
// escondidas dentro de "Filtros". El grosor del borde sigue debajo del
// canvas, en el panel del lienzo, porque se ajusta mirando el resultado.

// El modelo de recorte y su carga son compartidos con el Editor avanzado
// (recorte-sujeto.js); las islas, con sus máscaras (mascaras-capa.js).
import { cargarRecortador } from './recorte-sujeto.js';

const tt = (clave, porDefecto) => (window.t ? window.t(clave) : porDefecto);
// Resoluciones (2026-10-08). Antes todo se hacía a 512 px —el recorte, el
// borde y la descarga— y una foto de 4000 px perdía casi toda su calidad.
// Ahora el recorte se guarda a tamaño completo (con tope, por la memoria del
// celular), la pantalla trabaja sobre una copia de LADO_VISTA y la descarga
// rehace filtro, textura y borde a tamaño completo.
const LADO_VISTA = 1024;
const LADO_DESCARGA = 4096;
// El grosor y la suavidad del borde, y el desenfoque de los filtros, están
// pensados en píxeles de un sticker de este lado: se escalan al tamaño real
// para que se vean igual en la vista, las miniaturas y la descarga.
const LADO_REFERENCIA = 512;
const escalaDe = (canvas) => Math.max(canvas.width, canvas.height) / LADO_REFERENCIA;

export function iniciarGeneradorStickers() {
  const zonaArrastre = document.getElementById('zonaArrastre');
  const entradaArchivo = document.getElementById('entradaArchivo');
  // Mientras la IA trabaja, la foto se ve con el velo animado encima
  // (assets/js/trabajando-ia.js), con el texto del paso y el avance.
  const veloIA = TrabajandoIA.crear(document.getElementById('trabajandoIA'));
  const mensajeError = document.getElementById('mensajeError');
  const editorCompleto = document.getElementById('editorCompleto');
  const entradaGrosor = document.getElementById('entradaGrosor');
  const etiquetaGrosor = document.getElementById('etiquetaGrosor');
  const entradaColorBorde = document.getElementById('entradaColorBorde');
  const entradaSuavidad = document.getElementById('entradaSuavidad');
  const etiquetaSuavidad = document.getElementById('etiquetaSuavidad');
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

  // Texturas: las mismas de los editores (galería de efectos-textura.js, con
  // los mosaicos y el cómic de efectos-mosaico.js). Se aplican después del
  // filtro y sus miniaturas se dibujan sobre el propio sticker. "Tamaño" es
  // la intensidad de las texturas de siempre; los mosaicos y el cómic lo
  // ocultan y muestran sus propios controles.
  const galeriaTexturas = document.getElementById('modosTextura');
  EfectosTextura.crearGaleria(galeriaTexturas);
  const controlesTexturaSticker = document.getElementById('controlesTexturaSticker');
  const entradaTamanoTextura = document.getElementById('tamanoTextura');
  const valorTamanoTextura = document.getElementById('valorTamanoTextura');
  const panelTextura = EfectosMosaico.crearPanel(controlesTexturaSticker, {
    alCambiar: () => actualizarFiltro(),
    campoIntensidad: entradaTamanoTextura.closest('label'),
  });
  let texturaActual = 'ninguno';

  let segmentador = null;
  let cargaPromesa = null;
  let canvasAlta = null; // recorte limpio, sin filtro, a tamaño completo (≤ LADO_DESCARGA)
  let canvasBase = null; // el mismo, reducido a LADO_VISTA para la pantalla
  let canvasFiltrado = null; // canvasBase con el filtro elegido ya aplicado
  let presetActual = 'ninguno';
  // Canvases de repuesto reutilizados en cada actualizarFiltro() (la vista
  // previa en vivo): evita crear canvases nuevos por cada ajuste. Las
  // miniaturas de presets no los usan (se generan una sola vez por imagen,
  // en lotes de fondo, no en un control que el usuario arrastra).
  const scratchFiltro = ProcesarFiltros.crearScratch();

  // A diferencia de los deslizadores de los editores de imagen,
  // "Intensidad del filtro" y "Tamaño" de textura llamaban a actualizarFiltro()
  // directo en cada evento 'input', sin agruparlos por fotograma: arrastrar
  // el control podía correr el pipeline completo (varios canvases nuevos +
  // getImageData/putImageData) muchas más veces de las que el navegador
  // llega a pintar, sintiéndose pegado. Mismo límite de "una vez por
  // requestAnimationFrame" que ya usan los editores.
  let actualizarFiltroPendiente = null;
  function programarActualizarFiltro() {
    if (actualizarFiltroPendiente) return;
    actualizarFiltroPendiente = requestAnimationFrame(() => {
      actualizarFiltroPendiente = null;
      actualizarFiltro();
    });
  }

  // Mismo motivo que programarActualizarFiltro: el grosor, la suavidad y el
  // color del borde redibujan el sticker una vez por fotograma.
  let dibujarStickerPendiente = null;
  function programarDibujarSticker() {
    if (dibujarStickerPendiente) return;
    dibujarStickerPendiente = requestAnimationFrame(() => {
      dibujarStickerPendiente = null;
      dibujarSticker();
    });
  }

  function mostrarError(texto) {
    mensajeError.textContent = '❌ ' + texto;
    mensajeError.style.display = 'block';
  }
  function ocultarError() {
    mensajeError.style.display = 'none';
  }

  // Carga (o reusa) el segmentador (recorte-sujeto.js: placa de video si hay
  // una de verdad, si no CPU), mostrando el avance de la descarga en el velo.
  function cargarSegmentador(trabajo) {
    if (segmentador) return Promise.resolve(segmentador);
    if (cargaPromesa) return cargaPromesa;

    trabajo.texto(tt('estado-preparando', 'Preparando el modelo...'));
    const promesa = cargarRecortador((porcentaje) => {
      // El modelo son varios archivos: uno chico llega al 100 % antes que el
      // grande, y "Recortando…" aparecía antes de la descarga de verdad.
      trabajo.texto(porcentaje >= 100
        ? tt('estado-preparando', 'Preparando el modelo...')
        : tt('estado-descargando', 'Descargando el modelo... {p}% (solo esta vez, después queda guardado)').replace('{p}', porcentaje));
    });

    cargaPromesa = promesa.then((nuevo) => {
      segmentador = nuevo;
      return segmentador;
    }).catch((error) => {
      cargaPromesa = null;
      throw error;
    });

    return cargaPromesa;
  }

  // Reduce el recorte a un máximo de "lado" px (nunca lo agranda).
  function escalarRecorte(canvasOrigen, lado) {
    const factor = Math.min(1, lado / Math.max(canvasOrigen.width, canvasOrigen.height));
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

    const { etiquetas, tamanos } = MascarasCapa.islas(datos, width, height, 20);

    if (tamanos.length === 0) return;
    // Con un bucle: a 4096 px puede haber tantas islas que Math.max(...tamanos)
    // supera el límite de argumentos de una función.
    let tamanoMaximo = 0;
    for (const tamano of tamanos) if (tamano > tamanoMaximo) tamanoMaximo = tamano;
    const UMBRAL_RELATIVO = 0.08;
    const etiquetasValidas = new Set();
    tamanos.forEach((tamano, etiqueta) => { if (tamano >= tamanoMaximo * UMBRAL_RELATIVO) etiquetasValidas.add(etiqueta); });

    // Alfa bajo = restos del fondo que el segmentador dejó a medias (un libro
    // oscuro detrás de la taza, por ejemplo). Si se conservan, el borde blanco
    // los pinta como un velo translúcido. Se quitan aunque formen un bloque grande.
    const UMBRAL_VELO = 96;
    let seModificoAlgo = false;
    for (let i = 0; i < total; i++) {
      const alfa = datos[i * 4 + 3];
      if (alfa > 0 && (!etiquetasValidas.has(etiquetas[i]) || alfa < UMBRAL_VELO)) {
        datos[i * 4 + 3] = 0;
        seModificoAlgo = true;
      }
    }
    if (seModificoAlgo) ctx.putImageData(imageData, 0, 0);
  }

  // Recorta el lienzo al rectángulo del sujeto: la IA devuelve la foto entera
  // con el fondo transparente, y un sujeto chico en una foto grande quedaba
  // en una esquina de un PNG enorme y casi vacío.
  function ajustarAlSujeto(canvas) {
    const { width: w, height: h } = canvas;
    const datos = canvas.getContext('2d').getImageData(0, 0, w, h).data;
    let x0 = w, y0 = h, x1 = -1, y1 = -1;
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        if (datos[(y * w + x) * 4 + 3] === 0) continue;
        if (x < x0) x0 = x;
        if (x > x1) x1 = x;
        if (y < y0) y0 = y;
        y1 = y;
      }
    }
    if (x1 < 0 || (x0 === 0 && y0 === 0 && x1 === w - 1 && y1 === h - 1)) return canvas;
    const recorte = document.createElement('canvas');
    recorte.width = x1 - x0 + 1;
    recorte.height = y1 - y0 + 1;
    recorte.getContext('2d').drawImage(canvas, -x0, -y0);
    return recorte;
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

  // Algunas texturas (ej. los puntos, pixelar-circulos en
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

  // Dibuja, sobre el <canvas> de cada botón de una grilla (filtros o
  // texturas), cómo quedaría el sticker con esa opción — se calcula una sola
  // vez por recorte nuevo. Son ~80 filtros: procesarlos todos de un tirón
  // congela la pantalla justo después de "Recortando el fondo...", así que
  // se hacen de a lotes en el tiempo libre del navegador (mismo mecanismo que
  // editor-avanzado-imagen.html usa para su propia grilla de filtros).
  // dibujar(miniatura, boton) aplica la opción sobre la miniatura; devuelve
  // false si ese botón no tiene nada que dibujar.
  const generacionMiniaturas = { presets: 0, texturas: 0 };

  function dibujarMiniaturasEnLotes(grilla, botones, dibujar) {
    if (!canvasBase) return;
    const generacion = ++generacionMiniaturas[grilla];
    const origen = canvasBase;
    let indice = 0;

    function procesarLote(plazo) {
      if (generacion !== generacionMiniaturas[grilla] || origen !== canvasBase) return;
      const inicio = performance.now();
      let procesadas = 0;
      while (indice < botones.length && (!procesadas || performance.now() - inicio < 8)) {
        if (plazo && !plazo.didTimeout && procesadas && plazo.timeRemaining() < 2) break;
        const boton = botones[indice++];
        procesadas++;
        const miniatura = crearMiniaturaSticker(origen, TAMANO_MINIATURA_FISICO);
        if (dibujar(miniatura, boton) === false) continue;
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

  function actualizarMiniaturasPresets() {
    if (!canvasBase) return;
    dibujarMiniaturasEnLotes('presets', [...contenedorPresets.querySelectorAll('.preset-miniatura')], (miniatura, boton) => {
      const preset = ProcesarFiltros.PRESETS[boton.dataset.preset];
      if (!preset) return false;
      const alfaOriginal = miniatura.getContext('2d').getImageData(0, 0, miniatura.width, miniatura.height);
      ProcesarFiltros.renderizarFiltro(miniatura, preset, preset.capasExtra, preset.textura || ProcesarFiltros.SIN_TEXTURA, escalaDe(miniatura));
      restaurarAlfa(miniatura, alfaOriginal);
    });
  }

  // Cada textura con sus opciones de fábrica, sobre el sticker sin filtro.
  // Las de siempre, al máximo del deslizador (como en los editores), para
  // que se distingan entre sí.
  function actualizarMiniaturasTexturas() {
    if (!canvasBase) return;
    dibujarMiniaturasEnLotes('texturas', [...galeriaTexturas.querySelectorAll('.miniatura-textura')], (miniatura, boton) => {
      aplicarTextura(miniatura, boton.dataset.textura, Number(entradaTamanoTextura.max));
    });
  }

  // Aplica una textura sobre "canvas". Las de siempre pintan también sobre el
  // fondo vacío, así que se les devuelve la transparencia de antes; los
  // mosaicos y el cómic no la necesitan (y les recortaría las celdas del borde).
  // "intensidad" es el valor crudo del deslizador "Tamaño" (un porcentaje del
  // ancho de ESTE canvas, ver EfectosTextura.REFERENCIA_TEXTURA) — por eso
  // esta función recibe siempre el canvas real a dibujar (el sticker o su
  // miniatura, cada uno con su propio ancho) y no un número de píxeles ya
  // calculado por el que llama: antes ese cálculo se hacía afuera contra el
  // ancho del sticker (o de la miniatura) por separado, con su propia
  // fórmula, y un sticker recortado más chico que el tope de 512px terminaba
  // con la textura más débil que uno al tope, para el mismo deslizador.
  function aplicarTextura(canvas, tipo, intensidad, opciones) {
    if (tipo === 'ninguno') return;
    const ctx = canvas.getContext('2d');
    const intensidadPx = Math.max(1, intensidad * (canvas.width / EfectosTextura.REFERENCIA_TEXTURA));
    const alfaAntes = EfectosTextura.esMosaico(tipo) ? null : ctx.getImageData(0, 0, canvas.width, canvas.height);
    EfectosTextura.aplicarCompleto(tipo, ctx, undefined, canvas.width, canvas.height, intensidadPx, opciones);
    if (alfaAntes) restaurarAlfa(canvas, alfaAntes);
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
    canvasFiltrado = filtrar(canvasBase, scratchFiltro);
    dibujarSticker();
  }

  // El filtro y la textura elegidos sobre "origen" (la copia de pantalla o el
  // recorte a tamaño completo, para la descarga). Devuelve un canvas nuevo, o
  // el mismo origen si no hay nada que aplicar.
  function filtrar(origen, scratch) {
    let resultado = origen;
    const preset = ProcesarFiltros.PRESETS[presetActual];
    if (preset && presetActual !== 'ninguno') {
      const copia = ProcesarFiltros.copiarCanvas(origen);
      const alfaOriginal = copia.getContext('2d').getImageData(0, 0, copia.width, copia.height);
      const ajustes = { ...preset, intensidadFiltro: Number(entradaIntensidad.value) };
      // El tamaño de bloque de una textura (pixelar/puntos) viene calibrado
      // para fotos de resolución normal — en un sticker el mismo
      // número de píxeles se ve enorme en proporción, por eso queda ajustable
      // en vez de usar siempre el valor fijo del preset.
      const textura = preset.textura
        ? { ...preset.textura, intensidad: Number(entradaTextura.value) }
        : ProcesarFiltros.SIN_TEXTURA;
      ProcesarFiltros.renderizarFiltro(copia, ajustes, preset.capasExtra, textura, escalaDe(copia), scratch);
      restaurarAlfa(copia, alfaOriginal);
      resultado = copia;
    }
    // La textura va después del filtro
    if (texturaActual !== 'ninguno') {
      if (resultado === origen) resultado = ProcesarFiltros.copiarCanvas(origen);
      aplicarTextura(resultado, texturaActual, Number(entradaTamanoTextura.value), panelTextura.leer());
    }
    return resultado;
  }

  // El sticker completo (borde + recorte filtrado) en "destino". El borde es
  // una figura (assets/js/borde-sticker.js): contorno del recorte agrandado,
  // suavizado y unido con curvas. Se guarda la última figura de la vista para
  // no recalcularla si solo cambia el color.
  let figuraEnCache = { origen: null, radio: -1, suavidad: -1, figura: null };
  // El área máxima de un canvas en iPhone (4096²): más grande, no se dibuja.
  const AREA_MAXIMA = 4096 * 4096;
  function margenDe(filtrado) {
    const escala = escalaDe(filtrado);
    return Math.ceil((parseInt(entradaGrosor.value, 10) || 0) * escala) + Math.ceil(4 * escala);
  }
  function componerSticker(destino, filtrado, usarCache) {
    const escala = escalaDe(filtrado);
    const radio = (parseInt(entradaGrosor.value, 10) || 0) * escala;
    const suavidad = (parseInt(entradaSuavidad.value, 10) || 0) * escala;
    const margen = margenDe(filtrado);
    destino.width = filtrado.width + margen * 2;
    destino.height = filtrado.height + margen * 2;
    const ctx = destino.getContext('2d');
    ctx.clearRect(0, 0, destino.width, destino.height);
    if (radio > 0) {
      let figura;
      // La forma sale del alfa. Los filtros y las texturas de siempre lo
      // devuelven intacto (restaurarAlfa), así que se mide sobre el recorte
      // sin filtrar y la caché sirve aunque se muevan sus deslizadores; los
      // mosaicos y el cómic sí cambian el alfa.
      const fuente = EfectosTextura.esMosaico(texturaActual) || !canvasBase || filtrado.width !== canvasBase.width ? filtrado : canvasBase;
      const c = figuraEnCache;
      if (usarCache && c.origen === fuente && c.radio === radio && c.suavidad === suavidad) figura = c.figura;
      else {
        figura = BordeSticker.figura(fuente, radio, suavidad);
        if (usarCache) figuraEnCache = { origen: fuente, radio, suavidad, figura };
      }
      ctx.save();
      ctx.translate(margen, margen);
      ctx.fillStyle = entradaColorBorde.value;
      ctx.fill(figura, 'evenodd');
      ctx.restore();
    }
    ctx.drawImage(filtrado, margen, margen);
  }

  function dibujarSticker() {
    if (marcando) { dibujarMarcas(); return; }
    if (!canvasFiltrado) return;
    componerSticker(vistaResultado, canvasFiltrado, true);
    vistaResultado.style.display = 'block';
  }

  // ===== Qué recortar: el primer plano o los objetos que se marcan =====
  // (2026-10-08, idea del usuario.) Con "Los objetos que marco", la vista
  // muestra la foto original: cada toque deja una marca numerada (tocar una
  // marca la quita) y "Recortar los objetos marcados" le pide al modelo de
  // objetos (SlimSAM, objeto-por-toque.js) el objeto de cada marca y los junta
  // en un solo recorte. El análisis de la foto, lo lento, se hace una vez.
  const radiosModo = [...document.querySelectorAll('input[name="modoRecorte"]')];
  const ayudaMarcas = document.getElementById('ayudaMarcas');
  const botonRecortarObjetos = document.getElementById('botonRecortarObjetos');
  const botonBorrarMarcas = document.getElementById('botonBorrarMarcas');
  const botonCambiarMarcas = document.getElementById('botonCambiarMarcas');
  // El modelo de objetos trabaja a 1024 px: se le da una copia de a lo sumo
  // este lado (una foto de 24 MP a tamaño completo eran cientos de MB por
  // marca) y la máscara se agranda después a la foto.
  const LADO_OBJETOS = 2048;
  let fotoActual = null; // { url, imagen } de la foto cargada
  let altaSujeto = null; // el recorte del primer plano ya listo, para volver a él sin recalcular
  let analisisObjetos = null; // el análisis de SlimSAM de fotoActual
  let marcas = []; // { x, y } en 0..1 de la foto
  let marcando = false;
  let trabajando = false; // SlimSAM buscando: no se cambia el modo ni las marcas
  // Cada foto cargada es una "generación": lo que termina tarde de una foto
  // anterior (ISNet o SlimSAM, que tardan) se descarta en vez de mezclarse.
  let generacion = 0;
  const modoRecorte = () => radiosModo.find((r) => r.checked).value;

  function mostrarControlesModo() {
    const objetos = modoRecorte() === 'objetos';
    ayudaMarcas.hidden = !(objetos && marcando);
    botonRecortarObjetos.hidden = !(objetos && marcando);
    botonBorrarMarcas.hidden = !(objetos && marcando && marcas.length);
    botonCambiarMarcas.hidden = !(objetos && !marcando);
    botonRecortarObjetos.disabled = trabajando || !marcas.length;
    botonBorrarMarcas.disabled = trabajando;
    radiosModo.forEach((r) => { r.disabled = trabajando; });
    // Mientras se marca, la vista es la foto: no hay sticker que descargar.
    botonDescargar.disabled = trabajando || marcando;
    botonRecortarObjetos.textContent = tt('boton-recortar-objetos', 'Recortar los objetos marcados ({n})').replace('{n}', marcas.length);
  }

  // Lo que mide el radio de una marca en la vista (también para quitarla).
  const radioDeMarca = () => Math.max(10, Math.max(vistaResultado.width, vistaResultado.height) * 0.022);

  // La foto (reducida a LADO_VISTA) con las marcas numeradas encima.
  function dibujarMarcas() {
    if (!fotoActual) return;
    const { imagen } = fotoActual;
    const escala = Math.min(1, LADO_VISTA / Math.max(imagen.naturalWidth, imagen.naturalHeight));
    vistaResultado.width = Math.round(imagen.naturalWidth * escala);
    vistaResultado.height = Math.round(imagen.naturalHeight * escala);
    contextoResultado.drawImage(imagen, 0, 0, vistaResultado.width, vistaResultado.height);
    const enVista = marcas.map((m) => ({ x: m.x * vistaResultado.width, y: m.y * vistaResultado.height }));
    MarcasObjetos.dibujar(contextoResultado, enVista, radioDeMarca());
    vistaResultado.style.display = 'block';
  }

  function empezarAMarcar() {
    if (trabajando) return;
    marcando = true;
    vistaResultado.style.cursor = 'crosshair';
    mostrarControlesModo();
    dibujarMarcas();
  }

  radiosModo.forEach((r) => r.addEventListener('change', () => {
    if (modoRecorte() === 'objetos') { empezarAMarcar(); return; }
    marcando = false;
    vistaResultado.style.cursor = '';
    mostrarControlesModo();
    if (altaSujeto) usarRecorteListo(altaSujeto);
  }));
  botonCambiarMarcas.addEventListener('click', empezarAMarcar);
  // El texto de "Recortar los objetos marcados (n)" lo arma el JavaScript.
  document.addEventListener('cambioIdioma', mostrarControlesModo);
  botonBorrarMarcas.addEventListener('click', () => {
    if (trabajando) return;
    marcas = [];
    mostrarControlesModo();
    dibujarMarcas();
  });

  vistaResultado.addEventListener('click', (e) => {
    if (!marcando || trabajando) return;
    // Sin el borde del lienzo (clientLeft/clientTop), que corría el punto.
    const caja = vistaResultado.getBoundingClientRect();
    const x = (e.clientX - caja.left - vistaResultado.clientLeft) / vistaResultado.clientWidth;
    const y = (e.clientY - caja.top - vistaResultado.clientTop) / vistaResultado.clientHeight;
    if (x < 0 || x > 1 || y < 0 || y > 1) return;
    // Tocar sobre una marca (o casi) la quita; la distancia se mide en
    // píxeles de la vista, con el mismo radio con que se dibuja.
    const ancho = vistaResultado.width, alto = vistaResultado.height;
    const enVista = marcas.map((m) => ({ x: m.x * ancho, y: m.y * alto }));
    const cerca = MarcasObjetos.cercana(enVista, x * ancho, y * alto, radioDeMarca());
    if (cerca >= 0) marcas.splice(cerca, 1);
    else marcas.push({ x, y });
    mostrarControlesModo();
    dibujarMarcas();
  });

  // Una copia de la foto de a lo sumo LADO_OBJETOS px, como URL, para el modelo.
  async function urlParaObjetos(imagen) {
    const escala = Math.min(1, LADO_OBJETOS / Math.max(imagen.naturalWidth, imagen.naturalHeight));
    const copia = document.createElement('canvas');
    copia.width = Math.max(1, Math.round(imagen.naturalWidth * escala));
    copia.height = Math.max(1, Math.round(imagen.naturalHeight * escala));
    copia.getContext('2d').drawImage(imagen, 0, 0, copia.width, copia.height);
    const blob = await new Promise((r) => copia.toBlob(r, 'image/png'));
    return URL.createObjectURL(blob);
  }

  // Qué parte de la máscara tiene objeto (0 a 1), mirando una grilla de puntos.
  function proporcionCubierta(mascara) {
    const datos = mascara.getContext('2d').getImageData(0, 0, mascara.width, mascara.height).data;
    const paso = Math.max(1, Math.floor(Math.sqrt(mascara.width * mascara.height / 40000)));
    let total = 0, cubiertos = 0;
    for (let y = 0; y < mascara.height; y += paso) {
      for (let x = 0; x < mascara.width; x += paso) {
        total++;
        if (datos[(y * mascara.width + x) * 4 + 3]) cubiertos++;
      }
    }
    return total ? cubiertos / total : 0;
  }

  botonRecortarObjetos.addEventListener('click', async () => {
    if (!marcas.length || !fotoActual || trabajando) return;
    const gen = generacion;
    const foto = fotoActual;
    const puntos = marcas.slice(); // las marcas de este momento, aunque cambien
    ocultarError();
    trabajando = true;
    mostrarControlesModo();
    // La primera vez en la página, lo que tarda es preparar el modelo (bajarlo
    // o leerlo de la caché del navegador), no analizar la foto.
    const trabajo = veloIA.empezar(analisisObjetos ? tt('ia-buscando-objetos', 'Buscando los objetos… {i} de {n}').replace('{i}', 1).replace('{n}', puntos.length)
      : tt('estado-preparando', 'Preparando el modelo...'));
    try {
      const { analizarImagen, objetosEn } = await import('./objeto-por-toque.js');
      if (!analisisObjetos) {
        const url = await urlParaObjetos(foto.imagen);
        try {
          const analisis = await analizarImagen(url, (porcentaje) => {
            if (gen !== generacion) return;
            // Al 100 % ya bajó: lo que sigue es analizar la foto.
            trabajo.texto(porcentaje >= 100
              ? tt('ia-analizando-objetos', 'Analizando la foto para encontrar objetos…')
              : tt('ia-descargando-objetos', 'Descargando el modelo para elegir objetos… {p}% (solo la primera vez)').replace('{p}', porcentaje));
          });
          if (gen !== generacion) return;
          analisisObjetos = analisis;
        } finally {
          URL.revokeObjectURL(url);
        }
      }
      const mascara = await objetosEn(analisisObjetos, puntos, (i, n) => {
        if (gen !== generacion) return;
        trabajo.texto(tt('ia-buscando-objetos', 'Buscando los objetos… {i} de {n}').replace('{i}', i + 1).replace('{n}', n));
      });
      if (gen !== generacion) return;
      const cubierta = proporcionCubierta(mascara);
      if (cubierta === 0) throw new Error(tt('error-sin-objetos', 'no se encontró ningún objeto en las marcas.'));
      // La foto a tamaño completo (con tope), con la máscara agrandada como
      // transparencia: el suavizado al agrandarla deja el contorno parejo.
      const escala = Math.min(1, LADO_DESCARGA / Math.max(foto.imagen.naturalWidth, foto.imagen.naturalHeight));
      const recorte = document.createElement('canvas');
      recorte.width = Math.max(1, Math.round(foto.imagen.naturalWidth * escala));
      recorte.height = Math.max(1, Math.round(foto.imagen.naturalHeight * escala));
      const ctx = recorte.getContext('2d');
      ctx.drawImage(foto.imagen, 0, 0, recorte.width, recorte.height);
      ctx.globalCompositeOperation = 'destination-in';
      ctx.imageSmoothingQuality = 'high';
      ctx.drawImage(mascara, 0, 0, recorte.width, recorte.height);
      // Sin limpiarMascara: borra las islas chicas comparadas con la mayor, y
      // acá un objeto chico marcado a propósito es tan válido como uno grande.
      usarRecorteListo(ajustarAlSujeto(recorte));
      marcando = false;
      vistaResultado.style.cursor = '';
      if (cubierta > 0.9) mostrarError(tt('aviso-casi-toda', 'El recorte ocupa casi toda la foto: puede que una marca haya caído sobre el fondo.'));
    } catch (error) {
      if (gen !== generacion) return;
      const sinMemoria = /bad_alloc|memory|out of memory|OrtRun|allocation failed|RangeError/i.test(error.message);
      mostrarError(sinMemoria
        ? tt('error-sin-memoria', 'Tu navegador se quedó sin memoria para correr este modelo.')
        : tt('error-objetos', 'No se pudieron buscar los objetos: ') + error.message);
    } finally {
      trabajo.terminar();
      if (gen === generacion) {
        trabajando = false;
        mostrarControlesModo();
        dibujarSticker();
      }
    }
  });

  // Prepara un recorte crudo (la foto con el fondo transparente): a tamaño
  // completo con tope, sin manchas si "limpiar", y ajustado al sujeto.
  function prepararRecorte(recorte, limpiar) {
    let alta = escalarRecorte(recorte, LADO_DESCARGA);
    if (limpiar) limpiarMascara(alta);
    return ajustarAlSujeto(alta);
  }

  // Usa un recorte ya preparado: su copia para la pantalla, y el filtro, la
  // textura y el borde que ya estaban elegidos.
  function usarRecorteListo(alta) {
    canvasAlta = alta;
    canvasBase = escalarRecorte(canvasAlta, LADO_VISTA);
    actualizarMiniaturasPresets();
    actualizarMiniaturasTexturas();
    actualizarFiltro();
  }

  async function procesarArchivo(archivo) {
    const gen = ++generacion;
    ocultarError();
    vistaResultado.style.display = 'none';
    editorCompleto.style.display = 'none';
    canvasAlta = null;
    canvasBase = null;
    canvasFiltrado = null;
    // La foto queda guardada (y su URL viva) para poder marcar objetos en ella.
    if (fotoActual) URL.revokeObjectURL(fotoActual.url);
    fotoActual = null;
    altaSujeto = null;
    analisisObjetos = null;
    marcas = [];
    marcando = false;
    trabajando = false;
    vistaResultado.style.cursor = '';
    radiosModo.forEach((r) => { r.checked = r.value === 'sujeto'; });
    mostrarControlesModo();
    const url = URL.createObjectURL(archivo);
    let trabajo = null;
    // Si mientras tanto se cargó otra foto, esta URL ya no le sirve a nadie.
    const vigente = () => {
      if (gen === generacion) return true;
      if (!fotoActual || fotoActual.url !== url) URL.revokeObjectURL(url);
      return false;
    };
    try {
      const imagen = new Image();
      imagen.src = url;
      await imagen.decode();
      if (!vigente()) return;
      fotoActual = { url, imagen };
      // La foto se ve enseguida, con el velo, mientras la IA la recorta.
      editorCompleto.style.display = 'block';
      dibujarMarcas();
      trabajo = veloIA.empezar(tt('estado-recortando', 'Recortando el fondo...'));
      const modelo = await cargarSegmentador(trabajo);
      if (!vigente()) return;
      trabajo.texto(tt('estado-recortando', 'Recortando el fondo...'));

      const recorte = await modelo(url);
      if (!vigente()) return;
      altaSujeto = prepararRecorte(recorte.toCanvas(), true);
      editorCompleto.style.display = 'block';
      presetActual = 'ninguno';
      entradaIntensidad.value = 100;
      valorIntensidad.textContent = '100%';
      controlesTextura.style.display = 'none';
      contenedorPresets.querySelectorAll('.preset-miniatura.activo').forEach((b) => b.classList.remove('activo'));
      contenedorPresets.querySelector('[data-preset="ninguno"]')?.classList.add('activo');
      elegirTextura(galeriaTexturas.querySelector('[data-textura="ninguno"]'));
      usarRecorteListo(altaSujeto);
      mostrarControlesModo();
    } catch (error) {
      if (!vigente()) return;
      const sinMemoria = /bad_alloc|memory|out of memory|OrtRun|allocation failed|RangeError/i.test(error.message);
      mostrarError(sinMemoria
        ? tt('error-sin-memoria', 'Tu navegador se quedó sin memoria para correr este modelo.')
        : tt('error-recortar', 'No se pudo recortar la imagen: ') + error.message);
      // Sin recorte no hay editor: la foto no se va a usar.
      editorCompleto.style.display = 'none';
      vistaResultado.style.display = 'none';
      if (fotoActual) { URL.revokeObjectURL(fotoActual.url); fotoActual = null; } else URL.revokeObjectURL(url);
    } finally {
      if (trabajo) trabajo.terminar();
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
    programarDibujarSticker();
  });
  entradaColorBorde.addEventListener('input', programarDibujarSticker);
  entradaSuavidad.addEventListener('input', () => {
    etiquetaSuavidad.textContent = entradaSuavidad.value;
    programarDibujarSticker();
  });

  entradaIntensidad.addEventListener('input', () => {
    valorIntensidad.textContent = entradaIntensidad.value + '%';
    programarActualizarFiltro();
  });

  entradaTextura.addEventListener('input', () => {
    valorTextura.textContent = entradaTextura.value;
    programarActualizarFiltro();
  });

  function elegirTextura(boton) {
    galeriaTexturas.querySelectorAll('.miniatura-textura').forEach((b) => b.classList.toggle('activo', b === boton));
    texturaActual = boton.dataset.textura;
    controlesTexturaSticker.hidden = texturaActual === 'ninguno';
    panelTextura.mostrar(texturaActual);
  }

  entradaTamanoTextura.addEventListener('input', () => {
    valorTamanoTextura.textContent = entradaTamanoTextura.value;
    programarActualizarFiltro();
  });

  galeriaTexturas.addEventListener('click', (e) => {
    const boton = e.target.closest('.miniatura-textura');
    if (!boton) return;
    elegirTextura(boton);
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

  // La descarga rehace filtro, textura y borde sobre el recorte a tamaño
  // completo (puede tardar un par de segundos en una foto grande). Si el
  // sticker no entra en el área máxima de un canvas, o el navegador se queda
  // sin memoria, se reintenta más chico: 2048 y, último recurso, la vista.
  async function generarDescarga(lado) {
    let origen = escalarRecorte(canvasAlta, lado);
    const margen = margenDe(origen);
    const area = (origen.width + margen * 2) * (origen.height + margen * 2);
    if (area > AREA_MAXIMA) origen = escalarRecorte(origen, Math.floor(Math.max(origen.width, origen.height) * Math.sqrt(AREA_MAXIMA / area) * 0.98));
    const final = document.createElement('canvas');
    componerSticker(final, filtrar(origen, null), false);
    return new Promise((resolve) => final.toBlob(resolve, 'image/png'));
  }

  botonDescargar.addEventListener('click', async () => {
    if (!canvasAlta) return;
    const textoBoton = botonDescargar.textContent;
    botonDescargar.disabled = true;
    botonDescargar.textContent = tt('boton-descargar-preparando', 'Preparando el sticker…');
    await new Promise((r) => requestAnimationFrame(() => setTimeout(r, 0)));
    try {
      let blob = null;
      for (const lado of [LADO_DESCARGA, 2048, LADO_VISTA]) {
        try { blob = await generarDescarga(lado); } catch (error) { blob = null; }
        if (blob) break;
      }
      if (!blob) throw new Error('sin memoria');
      const enlace = document.createElement('a');
      enlace.href = URL.createObjectURL(blob);
      enlace.download = 'sticker.png';
      enlace.click();
      // Un PNG de varios MB: se libera cuando el navegador ya lo tomó.
      setTimeout(() => URL.revokeObjectURL(enlace.href), 60000);
    } catch (error) {
      mostrarError(tt('error-descargar', 'No se pudo armar el sticker para descargar: tu navegador se quedó sin memoria.'));
    } finally {
      botonDescargar.disabled = false;
      botonDescargar.textContent = textoBoton;
    }
  });
}

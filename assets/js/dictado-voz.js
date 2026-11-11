/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Dictado de voz a texto 100% local con Whisper (transformers.js) — sin backend propio.
// El audio nunca sale del navegador; solo el modelo (un archivo estático) se descarga
// una vez desde la CDN de Hugging Face y queda cacheado. Reusado por dictado-voz.html,
// que deja elegir entre modelo rápido (whisper-tiny), preciso (whisper-base) o muy
// preciso (whisper-small).

import { pipeline, env } from '../../vendor/@huggingface/transformers/4.3.0/dist/transformers.min.js?v=2';
import { cargarModeloIA } from './ia-dispositivo.js';

// Textos de usuario: los toma del diccionario de la pagina (idioma.js).
// El respaldo en espanol cubre el caso de que idioma.js no llegue a cargar.
const tt = (clave, porDefecto) => (window.t ? window.t(clave) : porDefecto);

env.allowLocalModels = false;

export function iniciarDictado() {
  const selectorModelo = document.getElementById('selectorModelo');
  const botonGrabar = document.getElementById('botonGrabar');
  const estadoGrabacion = document.getElementById('estadoGrabacion');
  const tiempoGrabando = document.getElementById('tiempoGrabando');
  const panelProgreso = document.getElementById('panelProgreso');
  const barraProgreso = document.getElementById('barraProgresoRelleno');
  const textoProgreso = document.getElementById('textoProgreso');
  const areaTexto = document.getElementById('areaTexto');
  const botonCopiar = document.getElementById('botonCopiar');
  const botonDescargar = document.getElementById('botonDescargar');
  const selectorIdioma = document.getElementById('selectorIdioma');
  const mensajeError = document.getElementById('mensajeError');
  const resultado = document.getElementById('resultado');

  let mediaRecorder = null;
  let chunks = [];
  let grabando = false;
  let inicioGrabacion = 0;
  let intervaloTiempo = null;
  let transcriptor = null;
  let modeloCargadoId = null;
  let cargaPromesa = null;
  let modeloIdEnCarga = null;

  function mostrarError(texto) {
    mensajeError.textContent = '❌ ' + texto;
    mensajeError.style.display = 'block';
  }
  function ocultarError() {
    mensajeError.style.display = 'none';
  }

  function formatearTiempo(ms) {
    const s = Math.floor(ms / 1000);
    return String(Math.floor(s / 60)).padStart(2, '0') + ':' + String(s % 60).padStart(2, '0');
  }

  // Carga (o reusa) el modelo elegido. Si dos llamadas piden el mismo modelo a la vez
  // (ej. la precarga automática y una transcripción), comparten la misma descarga.
  function cargarModelo(modeloId) {
    if (transcriptor && modeloCargadoId === modeloId) return Promise.resolve(transcriptor);
    if (cargaPromesa && modeloIdEnCarga === modeloId) return cargaPromesa;

    modeloIdEnCarga = modeloId;
    panelProgreso.style.display = 'block';
    textoProgreso.textContent = tt('estado-preparando', 'Preparando el modelo elegido...');
    barraProgreso.style.width = '0%';

    const archivos = {};
    const progressCallback = (info) => {
      if (info.status === 'progress' && info.file) {
        archivos[info.file] = info;
        const valores = Object.values(archivos);
        const cargado = valores.reduce((a, f) => a + (f.loaded || 0), 0);
        const total = valores.reduce((a, f) => a + (f.total || 0), 0);
        const porcentaje = total ? Math.round((cargado / total) * 100) : 0;
        barraProgreso.style.width = porcentaje + '%';
        textoProgreso.textContent = tt('estado-descargando',
          'Descargando modelo... {p}% (solo esta vez, después queda guardado)').replace('{p}', porcentaje);
      }
    };

    // Con placa de video, la memoria del modelo la maneja la GPU en vez del
    // límite de WebAssembly (más margen y más rápido); sin ella, CPU. La
    // elección y el respaldo están en ia-dispositivo.js.
    const promesaPipeline = cargarModeloIA((device) => pipeline('automatic-speech-recognition', modeloId, {
      dtype: 'q8',
      device,
      progress_callback: progressCallback,
    }));

    cargaPromesa = promesaPipeline.then((nuevoTranscriptor) => {
      transcriptor = nuevoTranscriptor;
      modeloCargadoId = modeloId;
      if (modeloIdEnCarga === modeloId) panelProgreso.style.display = 'none';
      return transcriptor;
    }).catch((error) => {
      cargaPromesa = null;
      modeloIdEnCarga = null;
      throw error;
    });

    return cargaPromesa;
  }

  // Precarga en segundo plano el modelo elegido, sin bloquear nada — así cuando
  // el usuario termine de grabar, lo más probable es que ya esté listo. Se
  // llama al empezar a grabar (o al cambiar de modelo después de haber
  // grabado), NO al abrir la página: el modelo pesa entre 41 y 249 MB y no
  // tiene sentido gastarle esos datos a quien solo entra a mirar.
  let yaSeUsoElDictado = false;
  function precargarModeloElegido() {
    ocultarError();
    cargarModelo(selectorModelo.value).catch((error) => {
      mostrarError(tt('error-cargar-modelo', 'No se pudo cargar el modelo: ') + error.message);
    });
  }

  async function resamplearA16kHz(blob) {
    const arrayBuffer = await blob.arrayBuffer();
    const audioCtx = new AudioContext();
    const bufferDecodificado = await audioCtx.decodeAudioData(arrayBuffer);
    await audioCtx.close();
    const offlineCtx = new OfflineAudioContext(1, Math.ceil(bufferDecodificado.duration * 16000), 16000);
    const fuente = offlineCtx.createBufferSource();
    fuente.buffer = bufferDecodificado;
    fuente.connect(offlineCtx.destination);
    fuente.start(0);
    const buffer16k = await offlineCtx.startRendering();
    return recortarSilencio(buffer16k.getChannelData(0));
  }

  // Whisper suele "alucinar" texto genérico (relleno, saludos repetidos) cuando
  // el audio arranca o termina en silencio/ruido de fondo bajo. Recortar esos
  // tramos antes de mandarlo al modelo reduce mucho ese efecto.
  function recortarSilencio(muestras) {
    const UMBRAL = 0.01;
    const VENTANA = 400; // ~25ms a 16kHz
    let inicio = 0;
    let fin = muestras.length;

    for (let i = 0; i < muestras.length - VENTANA; i += VENTANA) {
      let suma = 0;
      for (let j = i; j < i + VENTANA; j++) suma += Math.abs(muestras[j]);
      if (suma / VENTANA > UMBRAL) { inicio = Math.max(0, i - VENTANA); break; }
    }
    for (let i = muestras.length - VENTANA; i > inicio; i -= VENTANA) {
      let suma = 0;
      for (let j = i; j < i + VENTANA; j++) suma += Math.abs(muestras[j]);
      if (suma / VENTANA > UMBRAL) { fin = Math.min(muestras.length, i + VENTANA * 2); break; }
    }
    return inicio < fin ? muestras.subarray(inicio, fin) : muestras;
  }

  async function transcribir(blob) {
    ocultarError();
    resultado.classList.remove('visible');
    try {
      const modelo = await cargarModelo(selectorModelo.value);
      panelProgreso.style.display = 'block';
      barraProgreso.style.width = '100%';
      textoProgreso.textContent = tt('estado-transcribiendo', 'Transcribiendo...');

      const audio = await resamplearA16kHz(blob);
      const idioma = selectorIdioma.value;
      // no_repeat_ngram_size corta bucles de frases repetidas (una alucinación
      // típica de Whisper en audio con ruido o silencios).
      const opciones = { chunk_length_s: 30, stride_length_s: 5, no_repeat_ngram_size: 3 };
      if (idioma !== 'auto') opciones.language = idioma;

      const salida = await modelo(audio, opciones);
      // Cada grabación nueva se agrega al final, en vez de borrar lo que ya se
      // había dictado (o corregido a mano) antes.
      const textoNuevo = (salida.text || '').trim();
      const textoPrevio = areaTexto.value.trim();
      areaTexto.value = textoPrevio && textoNuevo ? textoPrevio + '\n' + textoNuevo : (textoPrevio || textoNuevo);
      resultado.classList.add('visible');
    } catch (error) {
      const sinMemoria = /bad_alloc|memory|out of memory|OrtRun/i.test(error.message);
      mostrarError(sinMemoria
        ? tt('error-sin-memoria', 'Tu navegador se quedó sin memoria para correr este modelo. Probá con uno más liviano.')
        : tt('error-transcribir', 'No se pudo transcribir: ') + error.message);
    } finally {
      panelProgreso.style.display = 'none';
    }
  }

  async function iniciarGrabacion() {
    ocultarError();
    yaSeUsoElDictado = true;
    // Ojo con el orden: la descarga del modelo se dispara DESPUÉS de que la
    // grabación ya arrancó. Aunque no se la espere, pipeline() de
    // transformers.js trabaja en el hilo principal, y lanzarla antes dejaba
    // el pedido del micrófono haciendo cola detrás de ella: el usuario
    // apretaba "grabar" y tenía que esperar a que bajara el modelo (41 a 249
    // MB) para que empezara a grabarse su voz. Primero el micrófono; el
    // modelo se baja mientras habla.
    try {
      // echoCancellation apagado y bitrate alto: evitan que el navegador
      // comprima/distorsione la voz antes de que la escuche el modelo.
      const stream = await navigator.mediaDevices.getUserMedia({
        audio: { echoCancellation: false, noiseSuppression: true, autoGainControl: true },
      });
      chunks = [];
      mediaRecorder = new MediaRecorder(stream, { audioBitsPerSecond: 128000 });
      mediaRecorder.ondataavailable = (e) => { if (e.data.size > 0) chunks.push(e.data); };
      mediaRecorder.onstop = () => {
        stream.getTracks().forEach((t) => t.stop());
        const blob = new Blob(chunks, { type: mediaRecorder.mimeType });
        transcribir(blob);
      };
      mediaRecorder.start();
      grabando = true;
      inicioGrabacion = Date.now();
      botonGrabar.textContent = tt('boton-detener', '⏹️ Detener grabación');
      botonGrabar.style.background = '#dc2626';
      selectorModelo.disabled = true;
      estadoGrabacion.style.display = 'flex';
      tiempoGrabando.textContent = '00:00';
      intervaloTiempo = setInterval(() => {
        tiempoGrabando.textContent = formatearTiempo(Date.now() - inicioGrabacion);
      }, 200);

      // Ya se está grabando: ahora sí, a bajar el modelo. Lo más probable es
      // que para cuando el usuario termine de hablar ya esté listo, y la
      // transcripción salga sin espera. El audio no depende de esto: lo
      // captura MediaRecorder por su cuenta, no el hilo principal.
      precargarModeloElegido();
    } catch (error) {
      mostrarError(tt('error-microfono', 'No se pudo acceder al micrófono: ') + error.message);
    }
  }

  function detenerGrabacion() {
    if (mediaRecorder && grabando) mediaRecorder.stop();
    grabando = false;
    clearInterval(intervaloTiempo);
    botonGrabar.textContent = tt('boton-grabar', '🎙️ Empezar a grabar');
    botonGrabar.style.background = '';
    selectorModelo.disabled = false;
    estadoGrabacion.style.display = 'none';
  }

  botonGrabar.addEventListener('click', () => {
    if (grabando) detenerGrabacion();
    else iniciarGrabacion();
  });

  selectorModelo.addEventListener('change', () => { if (yaSeUsoElDictado) precargarModeloElegido(); });

  botonCopiar.addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(areaTexto.value);
      botonCopiar.textContent = tt('boton-copiado', '✅ Copiado');
      setTimeout(() => (botonCopiar.textContent = tt('boton-copiar', 'Copiar')), 1500);
    } catch (error) {
      mostrarError(tt('error-copiar', 'No se pudo copiar: ') + error.message);
    }
  });

  botonDescargar.addEventListener('click', () => {
    const blob = new Blob([areaTexto.value], { type: 'text/plain' });
    const enlace = document.createElement('a');
    enlace.href = URL.createObjectURL(blob);
    enlace.download = (window.idiomaActual && idiomaActual()) === 'en' ? 'dictation.txt' : 'dictado.txt';
    enlace.click();
  });

  window.addEventListener('beforeunload', () => { if (grabando) detenerGrabacion(); });
}

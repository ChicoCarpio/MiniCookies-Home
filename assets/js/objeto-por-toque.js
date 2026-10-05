/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Encontrar el objeto que se tocó en una foto, sea el sujeto principal o
// cualquier otra cosa (una flor del fondo, una persona entre varias, aunque
// se toquen), con SlimSAM: una versión chica de Segment Anything (Apache-2.0,
// ~14 MB cuantizado) que corre en el navegador con transformers.js. Lo usa
// "Elegir objeto" del Editor avanzado.
//
// Dos pasos: analizarImagen() es lo caro (el codificador mira la foto
// entera, una vez por foto); objetoEn() es barato (el decodificador, uno por
// toque, sobre ese análisis). Lo único que se descarga, una vez, es el
// modelo (huggingface.co) y su motor (cdn.jsdelivr.net), como ISNet en
// recorte-sujeto.js.

import { SamModel, AutoProcessor, RawImage, Tensor, env } from '../../vendor/@huggingface/transformers/4.3.0/dist/transformers.min.js?v=2';
import { cargarModeloIA } from './ia-dispositivo.js';

env.allowLocalModels = false;

const MODELO_ID = 'Xenova/slimsam-77-uniform';

let carga = null;
const oyentesProgreso = new Set();

function cargarModelo(alProgresar) {
  if (alProgresar) oyentesProgreso.add(alProgresar);
  if (carga) return carga;
  const archivos = {};
  const progress_callback = (info) => {
    if (info.status !== 'progress' || !info.file) return;
    archivos[info.file] = info;
    const valores = Object.values(archivos);
    const cargado = valores.reduce((a, f) => a + (f.loaded || 0), 0);
    const total = valores.reduce((a, f) => a + (f.total || 0), 0);
    const porcentaje = total ? Math.round((cargado / total) * 100) : 0;
    oyentesProgreso.forEach((fn) => fn(porcentaje));
  };
  carga = cargarModeloIA((device) => SamModel.from_pretrained(MODELO_ID, { dtype: 'q8', device, progress_callback }))
    .then(async (modelo) => ({ modelo, procesador: await AutoProcessor.from_pretrained(MODELO_ID, { progress_callback }) }))
    .then((listo) => { oyentesProgreso.clear(); return listo; })
    .catch((error) => { carga = null; throw error; });
  return carga;
}

// El análisis de una foto (URL), para pedirle después objetos con objetoEn().
export async function analizarImagen(url, alProgresar = null) {
  const { modelo, procesador } = await cargarModelo(alProgresar);
  const entradas = await procesador(await RawImage.read(url));
  const embeddings = await modelo.get_image_embeddings(entradas);
  return { modelo, procesador, entradas, embeddings };
}

// La máscara del objeto que está en (x, y), en 0..1 de la foto: un canvas del
// tamaño de la foto, blanco donde está el objeto y transparente fuera. De las
// tres propuestas del modelo se queda con la de mejor puntaje.
export async function objetoEn(analisis, x, y) {
  const { modelo, procesador, entradas, embeddings } = analisis;
  const [altoModelo, anchoModelo] = entradas.reshaped_input_sizes[0];
  const input_points = new Tensor('float32', [x * anchoModelo, y * altoModelo], [1, 1, 1, 2]);
  const input_labels = new Tensor('int64', [1n], [1, 1, 1]);
  const salida = await modelo({ ...embeddings, input_points, input_labels });
  const mascaras = await procesador.post_process_masks(salida.pred_masks, entradas.original_sizes, entradas.reshaped_input_sizes);
  const mascara = mascaras[0]; // [1, 3, alto, ancho]
  const puntajes = salida.iou_scores.data;
  let mejor = 0;
  for (let i = 1; i < puntajes.length; i++) if (puntajes[i] > puntajes[mejor]) mejor = i;
  const [, , alto, ancho] = mascara.dims;
  const datos = mascara.data;
  const desde = mejor * alto * ancho;
  const lienzo = document.createElement('canvas');
  lienzo.width = ancho; lienzo.height = alto;
  const ctx = lienzo.getContext('2d');
  const imagen = ctx.createImageData(ancho, alto);
  const d = imagen.data;
  for (let p = 0; p < alto * ancho; p++) {
    if (datos[desde + p]) { const i = p * 4; d[i] = d[i + 1] = d[i + 2] = 255; d[i + 3] = 255; }
  }
  ctx.putImageData(imagen, 0, 0);
  return lienzo;
}

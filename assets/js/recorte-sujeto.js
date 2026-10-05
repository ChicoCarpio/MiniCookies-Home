/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Separar el sujeto del fondo de una foto con IA (ISNet, MIT), en el
// navegador con transformers.js. Lo usan el generador de stickers (recorta el
// sujeto) y el Editor avanzado (máscara "Primer plano" / "Fondo"). Hasta el
// 2026-10-04 la carga vivía dentro de generador-stickers.js.
//
// La foto nunca sale del navegador: lo único que se descarga, una vez y
// después queda guardado, es el modelo (~42 MB, de huggingface.co) y el motor
// que lo corre (onnxruntime-web, de cdn.jsdelivr.net), declarados en
// terceros.mjs y en la política de privacidad.
//
// Uso:
//   const recorte = await recortarSujeto(urlOArchivo, (porcentaje) => …);
//   // un canvas del tamaño de la foto: la foto con el fondo transparente
//   // (su canal alfa ES la máscara del sujeto).

import { pipeline, env } from '../../vendor/@huggingface/transformers/4.3.0/dist/transformers.min.js?v=2';
import { cargarModeloIA } from './ia-dispositivo.js';

env.allowLocalModels = false;

const MODELO_ID = 'xrds/isnet-general-onnx-int8';

let segmentador = null;
let cargaPromesa = null;
// Quienes esperan la descarga, para avisarles el avance (la carga se comparte).
const oyentesProgreso = new Set();

// Carga (o reusa) el modelo. Placa de video si hay una de verdad, si no CPU
// (ver ia-dispositivo.js). alProgresar(porcentaje) mientras se descarga.
export function cargarRecortador(alProgresar = null) {
  if (segmentador) return Promise.resolve(segmentador);
  if (alProgresar) oyentesProgreso.add(alProgresar);
  if (cargaPromesa) return cargaPromesa;

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

  cargaPromesa = cargarModeloIA((device) => pipeline('background-removal', MODELO_ID, {
    dtype: 'q8',
    device,
    progress_callback,
  })).then((nuevo) => {
    segmentador = nuevo;
    oyentesProgreso.clear();
    return segmentador;
  }).catch((error) => {
    cargaPromesa = null;
    throw error;
  });
  return cargaPromesa;
}

// La foto (URL, por ejemplo de URL.createObjectURL) con el fondo
// transparente, en un canvas común del tamaño de la foto.
export async function recortarSujeto(url, alProgresar = null) {
  const modelo = await cargarRecortador(alProgresar);
  const resultado = (await modelo(url)).toCanvas();
  const lienzo = document.createElement('canvas');
  lienzo.width = resultado.width; lienzo.height = resultado.height;
  lienzo.getContext('2d').drawImage(resultado, 0, 0);
  return lienzo;
}

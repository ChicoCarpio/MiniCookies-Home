/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
/*
 * Trabajador (Web Worker) del traductor de documentos: tiene su propio motor (traductor-motor.js) con las dos sesiones del modelo y traduce
 * de a un segmento por pedido (lote 1: con int8 dinámico, la traducción de un párrafo no depende de con cuáles se junta).
 * Dirección: traductor-hilo.js?ort=/ruta/de/ort/&piezas=/ruta/traductor-piezas.js&motor=/ruta/traductor-motor.js
 * Mensajes: { tipo: 'cargar', vocabulario, codificador, decodificador }  (bytes del modelo, ya bajados por la página)
 *           { tipo: 'traducir', segmento: { texto, contexto? }, destino, haz, alfa }  -> { texto, ms }
 * onnxruntime-web con UN hilo (el sitio no tiene aislamiento de origen); el paralelismo sale de tener varios trabajadores.
 */
'use strict';
const parametros = new URLSearchParams(self.location.search);
const rutaOrt = parametros.get('ort');
importScripts(rutaOrt + 'ort.wasm.min.js', parametros.get('piezas'), parametros.get('motor'));
ort.env.wasm.wasmPaths = rutaOrt;
ort.env.wasm.numThreads = 1;

let motor = null;

self.onmessage = async (e) => {
  const m = e.data;
  try {
    if (m.tipo === 'cargar') {
      const tokenizador = self.TraductorPiezas.crear(m.vocabulario);
      motor = await self.TraductorMotor.crear({
        ort, tokenizador, codificador: m.codificador, decodificador: m.decodificador,
        opcionesSesion: { executionProviders: ['wasm'], graphOptimizationLevel: 'all' },
      });
      self.postMessage({ id: m.id });
    } else if (m.tipo === 'traducir') {
      const t0 = performance.now();
      const [texto] = await motor.traducir([m.segmento], m.destino, { haz: m.haz, alfa: m.alfa });
      self.postMessage({ id: m.id, texto, ms: performance.now() - t0 });
    } else if (m.tipo === 'cerrar') {
      if (motor) await motor.cerrar();
      motor = null;
      self.postMessage({ id: m.id });
    }
  } catch (err) {
    self.postMessage({ id: m.id, error: String((err && err.stack) || err) });
  }
};

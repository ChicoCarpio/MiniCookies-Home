/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
/*
 * Trabajador (Web Worker) de Librarius página. Dos roles, según el parámetro `rol` de su dirección:
 *   detector: tiene la sesión de Versus y la página; detecta las líneas, las recorta enderezadas y gira la página.
 *   lector:   tiene una sesión de Librarius; lee recortes.
 * Dirección: librarius-hilo.js?rol=detector|lector&ort=/ruta/de/ort/&pag=/ruta/librarius-pagina.js&modelo=/ruta/modelo.onnx
 * Cada trabajador corre onnxruntime-web con UN hilo (el sitio no tiene aislamiento de origen); el paralelismo sale de tener varios lectores.
 */
'use strict';
const parametros = new URLSearchParams(self.location.search);
const rol = parametros.get('rol'), rutaOrt = parametros.get('ort'), rutaPagina = parametros.get('pag'), rutaModelo = parametros.get('modelo');
importScripts(rutaOrt + 'ort.wasm.min.js', rutaPagina);
ort.env.wasm.wasmPaths = rutaOrt;
ort.env.wasm.numThreads = 1;
const N = self.LibrariusPagina._nucleo;

let ctx = null, pagina = null, lineas = null;

self.onmessage = async (e) => {
  const m = e.data;
  try {
    if (m.tipo === 'cargar') {
      const datos = new Uint8Array(await (await fetch(rutaModelo)).arrayBuffer());
      const sesion = await ort.InferenceSession.create(datos, { executionProviders: ['wasm'], graphOptimizationLevel: 'all' });
      ctx = N.contextos(Object.assign({ ort, alfabeto: m.alfabeto }, m.opciones, rol === 'detector' ? { versus: sesion } : { librarius: sesion }));
      self.postMessage({ id: m.id, mb: datos.length / 1e6 });
    } else if (m.tipo === 'detectar') {                         // detector: con la página nueva, o con la que ya tiene (después de girarla)
      if (m.gris) pagina = { gris: m.gris, w: m.w, h: m.h };
      pagina.relleno = Math.trunc(N.percentil(pagina.gris, 80));
      const tiempos = {};
      lineas = await N.detectarLineas(ctx.det, pagina.gris, pagina.w, pagina.h, tiempos);
      N.alinearEje(lineas, m.eje || null);
      self.postMessage({ id: m.id, lineas, w: pagina.w, h: pagina.h, tiempos });
    } else if (m.tipo === 'recortes') {                          // detector: recortes enderezados de las líneas pedidas
      const t0 = performance.now();
      const crops = m.indices.map((i) => N.recortarLinea(ctx.rec, pagina.gris, pagina.w, pagina.h, lineas[i], m.invertir, pagina.relleno));
      self.postMessage({ id: m.id, crops, ms: performance.now() - t0 }, crops.map((c) => c.px.buffer));
    } else if (m.tipo === 'girar') {                             // detector: gira la página que tiene
      const g = N.girarPagina(pagina.gris, pagina.w, pagina.h, m.k);
      pagina = { gris: g.gris, w: g.w, h: g.h };
      self.postMessage({ id: m.id, w: g.w, h: g.h });
    } else if (m.tipo === 'leer') {                              // lector: un recorte
      const r = await N.leerRecorte(ctx.lec, { px: m.px, ancho: m.ancho });
      self.postMessage({ id: m.id, texto: r.texto, seguridad: r.seguridad });
    }
  } catch (err) {
    self.postMessage({ id: m.id, error: String(err && err.stack || err) });
  }
};

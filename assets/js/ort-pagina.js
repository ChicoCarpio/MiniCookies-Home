/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Carga onnxruntime-web (WebAssembly, un hilo) una sola vez, para los módulos con IA propia que corren en la página o
// en un hilo de Flujos (Custos, Vestigator).
//
//   cargarOrt() -> Promise   (rechaza si no se puede; el siguiente pedido lo vuelve a intentar)
//
// Necesita assets/js/vendor.js antes.
(function () {
  let promesa = null;

  function cargarScript(src) {
    return new Promise((ok, mal) => {
      const s = document.createElement('script');
      s.src = src; s.onload = ok; s.onerror = () => mal(errorMC('carga-no-cargo', { archivo: src }, 'No cargó {archivo}'));
      document.head.appendChild(s);
    });
  }

  window.cargarOrt = function () {
    if (!promesa) promesa = (async () => {
      if (typeof WebAssembly !== 'object') throw new Error('WebAssembly no disponible');
      if (!window.ort) await cargarScript(rutaVendor('onnxruntime-web/1.20.1/ort.wasm.min.js'));
      ort.env.wasm.numThreads = 1;   // el sitio no tiene aislamiento de origen (sin SharedArrayBuffer)
      ort.env.wasm.wasmPaths = {
        mjs: rutaVendor('onnxruntime-web/1.20.1/ort-wasm-simd-threaded.mjs'),
        wasm: rutaVendor('onnxruntime-web/1.20.1/ort-wasm-simd-threaded.wasm'),
      };
    })().catch((error) => { promesa = null; throw error; });
    return promesa;
  };
})();

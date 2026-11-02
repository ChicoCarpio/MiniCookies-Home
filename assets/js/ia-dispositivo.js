/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Elegir dónde corre un modelo de IA (placa de video o CPU) y cargarlo con
// respaldo. Lo usan dictado por voz y el generador de stickers.
//
// Por qué existe: las dos herramientas probaban WebGPU si `navigator.gpu`
// existía y, si fallaba, reintentaban SIN indicar dispositivo. Pero que exista
// `navigator.gpu` no quiere decir que haya una placa usable (pasa en PCs sin
// placa dedicada, con drivers viejos o en escritorios remotos), y sin
// dispositivo explícito transformers.js volvía a elegir WebGPU: el reintento
// fallaba igual ("no available backend found … [webgpu]") y la herramienta
// quedaba inutilizable en esas computadoras. Ahora se pregunta por un
// adaptador real antes de intentar, y el respaldo fuerza 'wasm' (CPU).
//
// Uso:
//   const modelo = await cargarModeloIA((device) => pipeline(tarea, id, { device, … }));

async function hayPlacaDeVideo() {
  try {
    return !!(navigator.gpu && await navigator.gpu.requestAdapter());
  } catch {
    return false;
  }
}

export async function cargarModeloIA(crear) {
  if (await hayPlacaDeVideo()) {
    try {
      return await crear('webgpu');
    } catch (error) {
      console.warn('WebGPU no pudo cargar el modelo; se usa la CPU:', error);
    }
  }
  return crear('wasm');
}

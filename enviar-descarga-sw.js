/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Service worker de "Enviar archivos con un QR": convierte lo que llega por la
// conexión en una descarga normal del navegador, sin juntar el archivo en
// memoria ni copiarlo antes al disco. Corre en el dispositivo de quien recibe;
// no habla con ningún servidor.
//
// Cómo: la página le manda un mensaje con los datos del archivo y un canal
// (MessagePort); después pide la dirección /enviar-descarga/<id>/<nombre> en
// un iframe oculto. Acá se responde a esa petición con un flujo que se llena
// con lo que la página va enviando por el canal. El flujo avisa "pull" cuando
// el navegador consumió lo anterior: la página manda el trozo siguiente solo
// entonces, así no se acumula nada pendiente.
const pendientes = new Map();   // id → { nombre, tam, tipo, flujo }
const esperando = new Map();    // id → función que atiende la petición cuando llega el mensaje

self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (e) => e.waitUntil(self.clients.claim()));

self.addEventListener('message', (e) => {
  const d = e.data;
  if (!d || d.tipo !== 'preparar' || !e.ports[0]) return;
  const puerto = e.ports[0];
  let controlador;
  const flujo = new ReadableStream({
    start(c) { controlador = c; },
    pull() { puerto.postMessage('pull'); },
    cancel() { puerto.postMessage('cancelado'); },
  }, { highWaterMark: 1 });
  puerto.onmessage = (m) => {
    try {
      if (m.data === 'fin') controlador.close();
      else if (m.data === 'abortar') controlador.error(new Error('abortado'));
      else controlador.enqueue(new Uint8Array(m.data));
    } catch (error) { /* el navegador ya canceló la descarga */ }
  };
  const info = { nombre: d.nombre, tam: d.tam, tipo: d.mime, flujo };
  if (esperando.has(d.id)) { esperando.get(d.id)(info); esperando.delete(d.id); }
  else pendientes.set(d.id, info);
});

function respuesta(info) {
  const cabeceras = {
    'Content-Type': info.tipo || 'application/octet-stream',
    'Content-Disposition': "attachment; filename*=UTF-8''" + encodeURIComponent(info.nombre),
    'Cache-Control': 'no-store',
  };
  if (Number.isFinite(info.tam)) cabeceras['Content-Length'] = String(info.tam);
  return new Response(info.flujo, { headers: cabeceras });
}

self.addEventListener('fetch', (e) => {
  const m = /^\/enviar-descarga\/([A-Za-z0-9-]+)\//.exec(new URL(e.request.url).pathname);
  if (!m) return;
  const id = m[1];
  if (pendientes.has(id)) {
    const info = pendientes.get(id);
    pendientes.delete(id);
    e.respondWith(respuesta(info));
    return;
  }
  // La petición puede llegar un instante antes que el mensaje: se espera hasta 3 s.
  e.respondWith(new Promise((listo) => {
    const reloj = setTimeout(() => { esperando.delete(id); listo(new Response('No encontrada', { status: 404 })); }, 3000);
    esperando.set(id, (info) => { clearTimeout(reloj); listo(respuesta(info)); });
  }));
});

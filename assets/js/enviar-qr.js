/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Enviar archivos de un dispositivo a otro, sin servidor de por medio y
// cifrados de extremo a extremo (herramienta "Enviar archivos con un QR").
//
// Cómo se conectan sin servidor: WebRTC necesita intercambiar una oferta y una
// respuesta antes de abrir el canal. Acá viajan por QR (o por un código de
// texto): el que recibe muestra un QR con su oferta y una clave; el que envía
// lo lee, muestra el QR de su respuesta y el que recibe lo lee. No se usa
// ningún servidor propio: solo el STUN de Cloudflare, que dice cuál es la
// dirección pública de cada equipo. Sin TURN (reenvío), entre redes distintas
// funciona solo si los NAT lo permiten; en la misma red siempre.
//
// Qué protege qué:
// - WebRTC cifra el canal (DTLS) y la huella del certificado viaja en el QR,
//   que se lee de pantalla a cámara: nadie puede meterse en el medio.
// - Encima, cada mensaje se cifra con AES-GCM de 256 bits con una clave
//   aleatoria que viaja solo en el QR (nunca por la red). El número de cada
//   mensaje va en el IV y se exige consecutivo: un mensaje repetido, perdido
//   o reordenado corta la conexión.
// - Un código de verificación de 6 dígitos, calculado con la clave y las dos
//   huellas, se muestra en las dos pantallas: si alguien se coló con el QR
//   visto de lejos, no coinciden.
//
// Piezas: `anfitrion()` (el que recibe) e `invitado(codigo)` (el que envía)
// devuelven una sesión con un canal cifrado; `enviarArchivos` y
// `recibirArchivos` hablan por ese canal.
const EnviarQR = (() => {
  const PREFIJO = 'MC1.';
  const TROZO = 48 * 1024;            // bytes de archivo por mensaje (bajo los 64 KiB que aceptan todos los navegadores)
  const BLOQUE = 1024 * 1024;         // cuánto se lee del archivo de una vez (leer en pedacitos es lo lento, sobre todo en celular)
  const BUFFER_ALTO = 8 * 1024 * 1024; // pausa el envío si el canal acumula más que esto
  const BUFFER_BAJO = 1024 * 1024;
  const ESPERA_ICE_MS = 4000;
  const ESPERA_CONEXION_MS = 20000;
  const TIPO_CONTROL = 1;
  const TIPO_TROZO = 2;
  const DIR_ANFITRION = 1;
  const DIR_INVITADO = 2;

  // ---------- Códigos: JSON comprimido en base64 seguro para direcciones ----------

  function aBase64Url(bytes) {
    let s = '';
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
    return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }
  function deBase64Url(texto) {
    const b = atob(texto.replace(/-/g, '+').replace(/_/g, '/'));
    const bytes = new Uint8Array(b.length);
    for (let i = 0; i < b.length; i++) bytes[i] = b.charCodeAt(i);
    return bytes;
  }
  async function pasar(bytes, flujo) {
    const salida = new Response(new Blob([bytes]).stream().pipeThrough(flujo));
    return new Uint8Array(await salida.arrayBuffer());
  }
  async function empaquetar(objeto) {
    const bytes = new TextEncoder().encode(JSON.stringify(objeto));
    return PREFIJO + aBase64Url(await pasar(bytes, new CompressionStream('deflate-raw')));
  }
  async function desempaquetar(codigo) {
    const bytes = await pasar(deBase64Url(codigo.slice(PREFIJO.length)), new DecompressionStream('deflate-raw'));
    return JSON.parse(new TextDecoder().decode(bytes));
  }
  // Acepta la dirección completa del QR ("…#MC1.xxxx"), el código solo o texto
  // con espacios y saltos de línea de más (al pegar). Devuelve el contenido
  // ({ t: 'o' | 'r', s, k? }) o null si no es un código de esta herramienta.
  async function leerCodigo(texto) {
    const m = /MC1\.[A-Za-z0-9_-]+/.exec(String(texto || '').replace(/\s+/g, ''));
    if (!m) return null;
    try {
      const datos = await desempaquetar(m[0]);
      return datos && (datos.t === 'o' || datos.t === 'r') && typeof datos.s === 'string' ? datos : null;
    } catch (e) {
      return null;
    }
  }

  // ---------- Conexión ----------

  function esperarIce(pc) {
    return new Promise((listo) => {
      if (pc.iceGatheringState === 'complete') return listo();
      const fin = () => { clearTimeout(reloj); pc.removeEventListener('icegatheringstatechange', cambio); listo(); };
      const cambio = () => { if (pc.iceGatheringState === 'complete') fin(); };
      const reloj = setTimeout(fin, ESPERA_ICE_MS);
      pc.addEventListener('icegatheringstatechange', cambio);
    });
  }
  function nuevaConexion() {
    // El STUN de Cloudflare solo descubre la dirección pública (no ve los
    // archivos): con él se puede conectar entre redes distintas cuando el NAT
    // no es estricto. Sin TURN, con datos móviles todavía puede fallar.
    return new RTCPeerConnection({ iceServers: [{ urls: 'stun:stun.cloudflare.com:3478' }] });
  }
  function huella(sdp) {
    const m = /a=fingerprint:\S+\s+([0-9A-Fa-f:]+)/.exec(sdp);
    return m ? m[1].toUpperCase() : '';
  }
  async function codigoDeVerificacion(clave, sdpOferta, sdpRespuesta) {
    const datos = new TextEncoder().encode(aBase64Url(clave) + '|' + huella(sdpOferta) + '|' + huella(sdpRespuesta));
    const h = new DataView(await crypto.subtle.digest('SHA-256', datos));
    const n = String(h.getUint32(0) % 1000000).padStart(6, '0');
    return n.slice(0, 3) + ' ' + n.slice(3);
  }
  function importarClave(bytes) {
    return crypto.subtle.importKey('raw', bytes, 'AES-GCM', false, ['encrypt', 'decrypt']);
  }
  function esperarApertura(canal, ms) {
    return new Promise((listo, error) => {
      if (canal.readyState === 'open') return listo();
      const reloj = setTimeout(() => error(new Error('tiempo')), ms);
      canal.addEventListener('open', () => { clearTimeout(reloj); listo(); }, { once: true });
      canal.addEventListener('close', () => { clearTimeout(reloj); error(new Error('cerrado')); }, { once: true });
    });
  }

  // ---------- Canal cifrado ----------

  // Cada mensaje: IV (12 bytes: dirección + contador) + AES-GCM(tipo + datos).
  class Canal {
    constructor(dc, clave, soyAnfitrion) {
      this.dc = dc; this.clave = clave;
      this.dirPropia = soyAnfitrion ? DIR_ANFITRION : DIR_INVITADO;
      this.dirAjena = soyAnfitrion ? DIR_INVITADO : DIR_ANFITRION;
      this.enviados = 0; this.recibidos = 0;
      this.alMensaje = null; this.alCerrar = null;
      this.cerrado = false;
      this.cola = Promise.resolve();
      dc.binaryType = 'arraybuffer';
      dc.bufferedAmountLowThreshold = BUFFER_BAJO;
      dc.onmessage = (e) => { this.cola = this.cola.then(() => this.recibir(e.data)).catch(() => this.cerrar()); };
      dc.onclose = () => this.cerrar();
    }
    // Asigna el número de mensaje en el acto (el orden queda fijo) y empieza a
    // cifrar; devuelve la promesa del mensaje listo. Así se pueden cifrar
    // varios a la vez y enviarlos en orden con `despachar`.
    preparar(tipo, datos) {
      if (this.cerrado) return Promise.reject(new Error('cerrado'));
      const iv = new Uint8Array(12);
      iv[0] = this.dirPropia;
      new DataView(iv.buffer).setBigUint64(4, BigInt(this.enviados++));
      const plano = new Uint8Array(1 + datos.length);
      plano[0] = tipo; plano.set(datos, 1);
      return crypto.subtle.encrypt({ name: 'AES-GCM', iv }, this.clave, plano).then((c) => {
        const mensaje = new Uint8Array(12 + c.byteLength);
        mensaje.set(iv); mensaje.set(new Uint8Array(c), 12);
        return mensaje;
      });
    }
    async despachar(preparado) {
      const mensaje = await preparado;
      if (this.dc.bufferedAmount > BUFFER_ALTO) {
        await new Promise((listo) => {
          const seguir = () => { this.dc.removeEventListener('bufferedamountlow', seguir); this.dc.removeEventListener('close', seguir); listo(); };
          this.dc.addEventListener('bufferedamountlow', seguir);
          this.dc.addEventListener('close', seguir);
        });
      }
      if (this.dc.readyState !== 'open') throw new Error('cerrado');
      this.dc.send(mensaje);
    }
    enviar(tipo, datos) { return this.despachar(this.preparar(tipo, datos)); }
    enviarControl(objeto) {
      return this.enviar(TIPO_CONTROL, new TextEncoder().encode(JSON.stringify(objeto)));
    }
    async recibir(buffer) {
      const m = new Uint8Array(buffer);
      if (m.length < 12 + 16 + 1) throw new Error('mensaje corto');
      const iv = m.subarray(0, 12);
      // Dirección contraria y contador consecutivo: lo repetido, perdido o
      // reordenado no se acepta.
      if (iv[0] !== this.dirAjena || new DataView(iv.buffer, iv.byteOffset).getBigUint64(4) !== BigInt(this.recibidos)) {
        throw new Error('mensaje fuera de orden');
      }
      const plano = new Uint8Array(await crypto.subtle.decrypt({ name: 'AES-GCM', iv }, this.clave, m.subarray(12)));
      this.recibidos++;
      if (this.alMensaje) await this.alMensaje(plano[0], plano.subarray(1));
    }
    cerrar() {
      if (this.cerrado) return;
      this.cerrado = true;
      try { this.dc.close(); } catch (e) {}
      if (this.alCerrar) this.alCerrar();
    }
  }

  function crearSesion(pc, dc, clave, soyAnfitrion, sdpOferta, sdpRespuesta, rawClave) {
    const canal = new Canal(dc, clave, soyAnfitrion);
    const sesion = {
      canal,
      verificacion: '',
      cerrar() { canal.cerrar(); try { pc.close(); } catch (e) {} },
    };
    pc.addEventListener('connectionstatechange', () => {
      if (pc.connectionState === 'failed' || pc.connectionState === 'closed') canal.cerrar();
    });
    return codigoDeVerificacion(rawClave, sdpOferta, sdpRespuesta).then((v) => { sesion.verificacion = v; return sesion; });
  }

  // El que recibe: arma la oferta. `codigo` es lo que va dentro del QR.
  async function anfitrion() {
    const pc = nuevaConexion();
    const dc = pc.createDataChannel('archivos', { ordered: true });
    const rawClave = crypto.getRandomValues(new Uint8Array(32));
    await pc.setLocalDescription(await pc.createOffer());
    await esperarIce(pc);
    const sdp = pc.localDescription.sdp;
    const codigo = await empaquetar({ t: 'o', s: sdp, k: aBase64Url(rawClave) });
    let usado = false;
    return {
      codigo,
      // Recibe el código de la respuesta (leído del QR del otro equipo) y
      // espera a que el canal abra. Una oferta sirve para una sola respuesta.
      async aceptarRespuesta(texto) {
        if (usado) throw new Error('usado');
        const datos = await leerCodigo(texto);
        if (!datos || datos.t !== 'r') throw new Error('codigo');
        usado = true;
        await pc.setRemoteDescription({ type: 'answer', sdp: datos.s });
        await esperarApertura(dc, ESPERA_CONEXION_MS);
        return crearSesion(pc, dc, await importarClave(rawClave), true, sdp, datos.s, rawClave);
      },
      cerrar() { try { dc.close(); } catch (e) {} try { pc.close(); } catch (e) {} },
    };
  }

  // El que envía: lee la oferta del QR y arma la respuesta. `conexion` se
  // cumple cuando el canal abre (con la sesión) o falla (por tiempo).
  async function invitado(texto) {
    const datos = await leerCodigo(texto);
    if (!datos || datos.t !== 'o' || !datos.k) throw new Error('codigo');
    const pc = nuevaConexion();
    const rawClave = deBase64Url(datos.k);
    const canalAbierto = new Promise((listo) => { pc.ondatachannel = (e) => listo(e.channel); });
    await pc.setRemoteDescription({ type: 'offer', sdp: datos.s });
    await pc.setLocalDescription(await pc.createAnswer());
    await esperarIce(pc);
    const sdp = pc.localDescription.sdp;
    const codigo = await empaquetar({ t: 'r', s: sdp });
    const conexion = (async () => {
      const dc = await Promise.race([
        canalAbierto,
        new Promise((_, error) => setTimeout(() => error(new Error('tiempo')), ESPERA_CONEXION_MS * 6)),
      ]);
      await esperarApertura(dc, ESPERA_CONEXION_MS);
      return crearSesion(pc, dc, await importarClave(rawClave), false, datos.s, sdp, rawClave);
    })();
    return { codigo, conexion, cerrar() { try { pc.close(); } catch (e) {} } };
  }

  // ---------- Transferencia de archivos ----------

  // Envía `archivos` (File[]). Espera a que el otro equipo acepte; devuelve
  // 'listo' si recibió todo o 'rechazado'. `alProgreso({ nombre, hecho, tam,
  // totalHecho, totalTam })` y `cancelado()` (devuelve true para parar).
  async function enviarArchivos(sesion, archivos, { alProgreso, cancelado } = {}) {
    const canal = sesion.canal;
    const espera = {};
    let cancelaOtro = false;
    const respuesta = new Promise((listo, error) => { espera.listo = listo; espera.error = error; });
    canal.alMensaje = (tipo, datos) => {
      if (tipo !== TIPO_CONTROL) return;
      const m = JSON.parse(new TextDecoder().decode(datos));
      if (m.t === 'cancela') cancelaOtro = true;
      if (m.t === 'acepta' || m.t === 'rechaza' || m.t === 'ok' || m.t === 'cancela') espera.listo(m.t);
    };
    canal.alCerrar = () => espera.error(new Error('cerrado'));
    const totalTam = archivos.reduce((s, f) => s + f.size, 0);
    await canal.enviarControl({ t: 'lista', archivos: archivos.map((f, id) => ({ id, nombre: f.name, tam: f.size, tipo: f.type })) });
    if ((await respuesta) !== 'acepta') return 'rechazado';
    let totalHecho = 0;
    for (let id = 0; id < archivos.length; id++) {
      const f = archivos[id];
      await canal.enviarControl({ t: 'inicio', id });
      for (let bloque = 0; bloque < f.size; bloque += BLOQUE) {
        if (cancelaOtro) return 'cancelado';
        if (cancelado && cancelado()) { await canal.enviarControl({ t: 'cancela' }); return 'cancelado'; }
        const datos = new Uint8Array(await f.slice(bloque, bloque + BLOQUE).arrayBuffer());
        const preparados = [];
        for (let o = 0; o < datos.length; o += TROZO) {
          const parte = datos.subarray(o, o + TROZO);
          const cuerpo = new Uint8Array(4 + parte.length);
          new DataView(cuerpo.buffer).setUint32(0, id);
          cuerpo.set(parte, 4);
          preparados.push([parte.length, canal.preparar(TIPO_TROZO, cuerpo)]);
        }
        let hecho = bloque;
        for (const [largo, preparado] of preparados) {
          await canal.despachar(preparado);
          hecho += largo; totalHecho += largo;
          if (alProgreso) alProgreso({ nombre: f.name, hecho, tam: f.size, totalHecho, totalTam });
        }
      }
      await canal.enviarControl({ t: 'fin', id });
    }
    const fin = new Promise((listo, error) => { espera.listo = listo; espera.error = error; });
    await canal.enviarControl({ t: 'listo' });
    return (await fin) === 'ok' ? 'listo' : 'cancelado';
  }

  // Recibe lo que manda el otro equipo. `alOferta(lista)` devuelve (o promete)
  // true para aceptar; `alArchivo(File)` por cada archivo terminado; el resto
  // son avisos. Devuelve una promesa con 'listo', 'rechazado' o 'cancelado'.
  function recibirArchivos(sesion, { alOferta, alProgreso, alArchivo }) {
    const canal = sesion.canal;
    return new Promise((fin, error) => {
      let lista = [], actual = null, totalHecho = 0, totalTam = 0;
      canal.alCerrar = () => error(new Error('cerrado'));
      canal.alMensaje = async (tipo, datos) => {
        if (tipo === TIPO_TROZO) {
          const id = new DataView(datos.buffer, datos.byteOffset).getUint32(0);
          if (!actual || actual.id !== id) throw new Error('trozo inesperado');
          const parte = datos.slice(4);
          actual.partes.push(parte); actual.hecho += parte.length; totalHecho += parte.length;
          if (alProgreso) alProgreso({ nombre: actual.nombre, hecho: actual.hecho, tam: actual.tam, totalHecho, totalTam });
          return;
        }
        const m = JSON.parse(new TextDecoder().decode(datos));
        if (m.t === 'lista') {
          lista = m.archivos;
          totalTam = lista.reduce((s, f) => s + f.tam, 0);
          const ok = await alOferta(lista);
          await canal.enviarControl({ t: ok ? 'acepta' : 'rechaza' });
          if (!ok) fin('rechazado');
        } else if (m.t === 'inicio') {
          const f = lista.find((x) => x.id === m.id);
          if (!f) throw new Error('archivo desconocido');
          actual = { id: f.id, nombre: f.nombre, tam: f.tam, tipo: f.tipo, partes: [], hecho: 0 };
        } else if (m.t === 'fin') {
          if (!actual || actual.id !== m.id || actual.hecho !== actual.tam) throw new Error('archivo incompleto');
          alArchivo(new File(actual.partes, actual.nombre, { type: actual.tipo }));
          actual = null;
        } else if (m.t === 'listo') {
          await canal.enviarControl({ t: 'ok' });
          fin('listo');
        } else if (m.t === 'cancela') {
          fin('cancelado');
        }
      };
    });
  }

  return { anfitrion, invitado, leerCodigo, enviarArchivos, recibirArchivos, TROZO };
})();

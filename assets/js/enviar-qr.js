/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
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
// - Un código de verificación de 8 dígitos, calculado con la clave y las dos
//   huellas: uno lo ve y se lo dice al otro, que lo escribe. Si alguien se coló
//   con el QR visto de lejos o se puso en el medio, no coincide. Hasta que los
//   dos confirman, el canal no acepta ni envía archivos.
//
// Los archivos grandes (más de 64 MB) no se juntan en memoria: el receptor los
// entrega a una descarga directa del navegador (service worker) o, si no se
// puede, los escribe en el disco del sitio. Los chicos se arman en memoria.
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
  // La conexión se corta sola si pasan 3 minutos sin que ninguno de los dos envíe
  // ni reciba nada (no cuenta mientras hay una transferencia en curso).
  let inactividadMs = 3 * 60 * 1000, revisionMs = 5000;
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
  // Un código real ocupa ~630 caracteres y su contenido unos pocos KB. Un código
  // pegado puede venir de cualquiera: con estos topes, uno armado para inflarse
  // (deflate multiplica hasta ~1000 veces) no llena la memoria ni cuelga la pestaña.
  const LARGO_MAXIMO_CODIGO = 16 * 1024;
  const MAXIMO_DESCOMPRIMIDO = 64 * 1024;
  async function pasar(bytes, flujo, maximo = Infinity) {
    const lector = new Blob([bytes]).stream().pipeThrough(flujo).getReader();
    const partes = [];
    let total = 0;
    for (;;) {
      const { done, value } = await lector.read();
      if (done) break;
      total += value.length;
      if (total > maximo) { lector.cancel().catch(() => {}); throw new Error('codigo'); }
      partes.push(value);
    }
    const salida = new Uint8Array(total);
    let o = 0;
    for (const p of partes) { salida.set(p, o); o += p.length; }
    return salida;
  }
  async function empaquetar(objeto) {
    const bytes = new TextEncoder().encode(JSON.stringify(objeto));
    return PREFIJO + aBase64Url(await pasar(bytes, new CompressionStream('deflate-raw')));
  }
  async function desempaquetar(codigo) {
    const bytes = await pasar(deBase64Url(codigo.slice(PREFIJO.length)), new DecompressionStream('deflate-raw'), MAXIMO_DESCOMPRIMIDO);
    return JSON.parse(new TextDecoder().decode(bytes));
  }
  // Acepta la dirección completa del QR ("…#MC1.xxxx"), el código solo o texto
  // con espacios y saltos de línea de más (al pegar). Devuelve el contenido
  // ({ t: 'o' | 'r', s, k? }) o null si no es un código de esta herramienta.
  //
  // Un código protegido con PIN ("MC2.…") se abre con `pin`: sin PIN lanza
  // 'pin-requerido' y con un PIN equivocado, 'pin'.
  async function leerCodigo(texto, pin) {
    texto = String(texto || '').replace(/\s+/g, '');
    if (texto.length > LARGO_MAXIMO_CODIGO) return null;
    const protegido = /MC2\.[A-Za-z0-9_-]+/.exec(texto);
    if (protegido) {
      if (!pin) throw new Error('pin-requerido');
      texto = await abrirProtegido(protegido[0], pin);
    }
    const m = /MC1\.[A-Za-z0-9_-]+/.exec(texto);
    if (!m) return null;
    try {
      const datos = await desempaquetar(m[0]);
      if (!datos || (datos.t !== 'o' && datos.t !== 'r') || typeof datos.s !== 'string') return null;
      // Un par "ice-lite" (solo candidato de servidor) no es de esta herramienta: ningún navegador lo arma así.
      if (/a=ice-lite/i.test(datos.s)) return null;
      return datos;
    } catch (e) {
      return null;
    }
  }

  // ---------- PIN opcional para el código de texto ----------

  // Si el código va por un chat o un correo, lleva la clave de cifrado a la
  // vista de quien lo vea. Con un PIN de 8 dígitos que se dice aparte, el
  // código viaja cifrado (AES-GCM con una clave sacada del PIN con PBKDF2) y
  // sin el PIN no sirve. Límite conocido: quien capture el código y tenga mucho
  // cómputo (GPU) puede probar los 100 millones de PIN fuera de línea; PBKDF2
  // lo hace caro (horas, no segundos como con 6 dígitos) y el código vence en
  // 2 minutos. Por eso el PIN se dice por otro medio, nunca junto al código.
  // El QR no lleva PIN: se lee de pantalla a cámara.
  const PREFIJO_PIN = 'MC2.';
  const ITERACIONES_PIN = 200000;
  async function claveDePin(pin, sal) {
    const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(String(pin).replace(/\D/g, '')), 'PBKDF2', false, ['deriveKey']);
    return crypto.subtle.deriveKey({ name: 'PBKDF2', salt: sal, iterations: ITERACIONES_PIN, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
  }
  function nuevoPin() {
    return String(crypto.getRandomValues(new Uint32Array(1))[0] % 100000000).padStart(8, '0');
  }
  async function proteger(codigo, pin) {
    const sal = crypto.getRandomValues(new Uint8Array(16));
    const iv = crypto.getRandomValues(new Uint8Array(12));
    const cifrado = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, await claveDePin(pin, sal), new TextEncoder().encode(codigo)));
    const salida = new Uint8Array(28 + cifrado.length);
    salida.set(sal); salida.set(iv, 16); salida.set(cifrado, 28);
    return PREFIJO_PIN + aBase64Url(salida);
  }
  async function abrirProtegido(codigo, pin) {
    try {
      const b = deBase64Url(codigo.slice(PREFIJO_PIN.length));
      const plano = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: b.subarray(16, 28) }, await claveDePin(pin, b.subarray(0, 16)), b.subarray(28));
      return new TextDecoder().decode(plano);
    } catch (e) {
      throw new Error('pin');
    }
  }
  // ¿Este texto es un código protegido con PIN?
  function esProtegido(texto) {
    return /MC2\.[A-Za-z0-9_-]+/.test(String(texto || '').replace(/\s+/g, ''));
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
  // 8 dígitos: un intermediario que quisiera fabricar certificados hasta que los
  // códigos de sus dos tramos coincidan tendría que probar ~100 millones.
  const DIGITOS_VERIFICACION = 8;
  async function codigoDeVerificacion(clave, sdpOferta, sdpRespuesta) {
    const datos = new TextEncoder().encode(aBase64Url(clave) + '|' + huella(sdpOferta) + '|' + huella(sdpRespuesta));
    const h = new DataView(await crypto.subtle.digest('SHA-256', datos));
    const n = String(h.getUint32(0) % 10 ** DIGITOS_VERIFICACION).padStart(DIGITOS_VERIFICACION, '0');
    return n.slice(0, 4) + ' ' + n.slice(4);
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
      // Un mismo canal sirve para enviar y recibir (de a una transferencia por vez):
      // lo que contesta el otro lado va al `emisor` y lo que inicia una transferencia
      // ajena, al `receptor`. Los dos se registran en enviarArchivos / recibirArchivos.
      this.emisor = null; this.receptor = null; this.cierreEmisor = null;
      this.cierres = new Set();
      this.ultima = Date.now();   // último mensaje enviado o recibido
      this.recibiendo = false;    // hay una transferencia entrante en curso (o esperando respuesta de la persona)
      this.motivo = null;         // por qué se cerró: 'inactividad', 'codigo' o null
      this.esAnfitrion = soyAnfitrion;
      this.verificado = false;    // las dos personas confirmaron el código de verificación
      // En el invitado hacen falta las dos cosas: que el anfitrión avise que escribió
      // bien el código y que la persona de este lado toque "ya lo confirmó". El aviso
      // solo no alcanza: un anfitrión malicioso podría mandarlo sin que nadie escriba nada.
      this.anfitrionConfirmo = false;
      this.yoConfirme = false;
      this.alVerificado = null;
      this.reloj = null;
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
      this.ultima = Date.now();
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
      this.ultima = Date.now();
      await this.repartir(plano[0], plano.subarray(1));
    }
    // Reparte un mensaje que llegó: al emisor (respuestas) o al receptor (transferencia ajena).
    repartir(tipo, datos) {
      // Sin verificar, el canal no entrega nada: ni trozos ni ofertas. Así la regla
      // vive acá y no depende de que la página espere a la verificación.
      if (tipo === TIPO_TROZO) return this.verificado && this.receptor ? this.receptor(tipo, datos) : undefined;
      const t = JSON.parse(new TextDecoder().decode(datos)).t;
      if (t === 'verificado') {
        // La manda el anfitrión al aceptar el código. Solo el invitado la acepta: si el
        // anfitrión también la aceptara, quien se colara con el QR se daría por verificado solo.
        if (!this.esAnfitrion) { this.anfitrionConfirmo = true; this.revisarVerificado(); }
        return;
      }
      if (!this.verificado) {
        // Si el otro ya verificó y quiere enviar antes de que esta persona confirme, se le avisa.
        if (t === 'lista') return this.enviarControl({ t: 'no-verificado' });
        return;
      }
      const paraEmisor = t === 'acepta' || t === 'rechaza' || t === 'sigue' || t === 'ok' || t === 'ocupado' || t === 'no-verificado' || (t === 'cancela' && this.emisor);
      const manejador = paraEmisor ? this.emisor : this.receptor;
      return manejador ? manejador(tipo, datos) : undefined;
    }
    revisarVerificado() {
      if (this.verificado || !this.anfitrionConfirmo || !this.yoConfirme) return;
      this.verificado = true;
      if (this.alVerificado) this.alVerificado();
    }
    cerrar() {
      if (this.cerrado) return;
      this.cerrado = true;
      clearInterval(this.reloj);
      try { this.dc.close(); } catch (e) {}
      for (const f of [...this.cierres]) f();
    }
  }

  function crearSesion(pc, dc, clave, soyAnfitrion, sdpOferta, sdpRespuesta, rawClave) {
    const canal = new Canal(dc, clave, soyAnfitrion);
    // Verificación: el código de 8 dígitos (clave + huellas de los dos) lo muestra el
    // invitado y el anfitrión, que mostró el QR, lo escribe. Quien se coló con un QR
    // visto de lejos no tiene a nadie al lado que le muestre un código, y un
    // intermediario que reenvía la conexión tiene huellas distintas en cada tramo.
    // El invitado, además, toca "ya lo confirmó" cuando el otro le dice que le
    // funcionó: el aviso del anfitrión solo no basta (ver Canal.anfitrionConfirmo).
    // Hasta que se confirme, no se puede enviar ni recibir.
    const sesion = {
      canal,
      verificacion: '',
      muestraCodigo: !soyAnfitrion,
      escribeCodigo: soyAnfitrion,
      intentos: 0,
      get verificado() { return canal.verificado; },
      // Solo el anfitrión. Tres intentos fallidos cierran la conexión.
      async confirmarCodigo(texto) {
        if (!this.escribeCodigo) return false;
        if (canal.verificado) return true;
        const limpio = (x) => String(x || '').replace(/\D/g, '');
        if (limpio(texto).length === DIGITOS_VERIFICACION && limpio(texto) === limpio(this.verificacion)) {
          canal.verificado = true;
          await canal.enviarControl({ t: 'verificado' });
          return true;
        }
        if (++this.intentos >= 3) { canal.motivo = 'codigo'; this.cerrar(); }
        return false;
      },
      // Solo el invitado: la persona confirma que el otro escribió el código y le funcionó.
      // Queda verificado cuando además llegó el aviso del anfitrión (en cualquier orden).
      confirmarMostrado() {
        if (!this.muestraCodigo || canal.cerrado) return false;
        canal.yoConfirme = true;
        canal.revisarVerificado();
        return true;
      },
      get anfitrionConfirmo() { return canal.anfitrionConfirmo; },
      cerrar() { canal.cerrar(); try { pc.close(); } catch (e) {} },
    };
    // Cierre por inactividad.
    canal.reloj = setInterval(() => {
      if (canal.cerrado || canal.emisor || canal.recibiendo) return;
      if (Date.now() - canal.ultima > inactividadMs) { canal.motivo = 'inactividad'; sesion.cerrar(); }
    }, revisionMs);
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
    let usado = false, pin = null;
    return {
      codigo,
      // El mismo código cifrado con un PIN nuevo (el PIN se crea la primera vez).
      async protegido() {
        if (!pin) pin = nuevoPin();
        return { pin, codigo: await proteger(codigo, pin) };
      },
      // Recibe el código de la respuesta (leído del QR del otro equipo) y
      // espera a que el canal abra. Una oferta sirve para una sola respuesta.
      async aceptarRespuesta(texto) {
        if (usado) throw new Error('usado');
        const datos = await leerCodigo(texto, pin);
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
  async function invitado(texto, pin) {
    const datos = await leerCodigo(texto, pin);
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
    // Si la oferta vino protegida, la respuesta también puede ir protegida con el mismo PIN.
    const protegida = pin && esProtegido(texto) ? await proteger(codigo, pin) : null;
    return { codigo, codigoProtegido: protegida, conexion, cerrar() { try { pc.close(); } catch (e) {} } };
  }

  // ---------- Transferencia de archivos ----------

  // Envía `archivos` (File[]). Espera a que el otro equipo acepte; devuelve
  // 'listo' si recibió todo o 'rechazado'. `alProgreso({ nombre, hecho, tam,
  // totalHecho, totalTam })` y `cancelado()` (devuelve true para parar).
  // Devuelve 'listo', 'rechazado', 'cancelado', 'ocupado' (el otro equipo está
  // recibiendo o enviando otra cosa en ese momento) o 'no-verificado' (la otra
  // persona todavía no confirmó el código de verificación).
  async function enviarArchivos(sesion, archivos, opciones = {}) {
    const canal = sesion.canal;
    if (!sesion.verificado) throw new Error('sin-verificar');
    if (canal.emisor) throw new Error('ocupado');
    try {
      return await enviarInterno(sesion, archivos, opciones);
    } finally {
      canal.emisor = null;
      if (canal.cierreEmisor) { canal.cierres.delete(canal.cierreEmisor); canal.cierreEmisor = null; }
    }
  }
  async function enviarInterno(sesion, archivos, { alProgreso, cancelado }) {
    const canal = sesion.canal;
    const espera = {};
    let cancelaOtro = false, sigue = null, rechazarCierre = null;
    const cierre = new Promise((_, rechazar) => { rechazarCierre = rechazar; });
    cierre.catch(() => {});
    const respuesta = new Promise((listo, error) => { espera.listo = listo; espera.error = error; });
    canal.emisor = (tipo, datos) => {
      if (tipo !== TIPO_CONTROL) return;
      const m = JSON.parse(new TextDecoder().decode(datos));
      if (m.t === 'cancela') { cancelaOtro = true; if (sigue) sigue(); }
      if (m.t === 'sigue' && sigue) sigue();
      if (m.t === 'acepta' || m.t === 'rechaza' || m.t === 'ocupado' || m.t === 'no-verificado' || m.t === 'ok' || m.t === 'cancela') espera.listo(m.t);
    };
    canal.cierreEmisor = () => { espera.error(new Error('cerrado')); rechazarCierre(new Error('cerrado')); };
    canal.cierres.add(canal.cierreEmisor);
    const totalTam = archivos.reduce((s, f) => s + f.size, 0);
    await canal.enviarControl({ t: 'lista', archivos: archivos.map((f, id) => ({ id, nombre: f.name, tam: f.size, tipo: f.type })) });
    const contesto = await respuesta;
    if (contesto !== 'acepta') return contesto === 'ocupado' || contesto === 'no-verificado' ? contesto : 'rechazado';
    let totalHecho = 0;
    for (let id = 0; id < archivos.length; id++) {
      const f = archivos[id];
      // El receptor contesta 'sigue' cuando está listo para recibir este archivo
      // (en un archivo grande, cuando la persona tocó "Empezar descarga").
      const listoElOtro = new Promise((listo) => { sigue = listo; });
      await canal.enviarControl({ t: 'inicio', id });
      if (alProgreso) alProgreso({ nombre: f.name, hecho: 0, tam: f.size, totalHecho, totalTam });
      await Promise.race([listoElOtro, cierre]);
      if (cancelaOtro) return 'cancelado';
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

  // ---------- Almacenamiento en disco del receptor ----------

  // Un archivo grande no se junta en memoria: se va escribiendo en el
  // almacenamiento privado del sitio (OPFS, que vive en el disco) y de ahí sale
  // el File que se guarda. Se decide por el tamaño, que se conoce de antemano.
  // Orden para los grandes: descarga directa, disco del sitio, memoria.
  const CARPETA_TEMPORAL = 'mc-recibidos';
  const UMBRAL_DISCO = 64 * 1024 * 1024; // hasta acá el archivo se arma en memoria (más rápido); más grande, directo a la descarga o al disco
  async function limpiarTemporales() {
    try { await (await navigator.storage.getDirectory()).removeEntry(CARPETA_TEMPORAL, { recursive: true }); } catch (e) {}
  }
  async function abrirTemporal() {
    try {
      const raiz = await navigator.storage.getDirectory();
      const carpeta = await raiz.getDirectoryHandle(CARPETA_TEMPORAL, { create: true });
      const manija = await carpeta.getFileHandle(crypto.randomUUID(), { create: true });
      const escritor = await manija.createWritable();
      return { manija, escribir: (b) => escritor.write(b), cerrar: () => escritor.close(), abortar: () => escritor.abort() };
    } catch (e) {
      return null;
    }
  }
  // Descarga directa: lo que llega se entrega, trozo a trozo, a un service
  // worker (enviar-descarga-sw.js, en la raíz del sitio) que lo convierte en una
  // descarga normal del navegador. No pasa por la memoria ni por el disco del
  // sitio: solo hace falta el espacio final en la carpeta de descargas. Devuelve
  // null si el navegador no puede (sin service worker, sin HTTPS, o si la
  // descarga no arranca): entonces se usa el disco del sitio o la memoria.
  //
  // Un navegador de celular solo deja bajar un archivo si la descarga nace de un
  // toque de la persona, así que `descargaDirecta` se llama desde el botón
  // "Empezar descarga" y el service worker se deja listo antes con `prepararSW`.
  let swListo = null;
  function prepararSW() {
    if (!swListo) {
      swListo = (async () => {
        if (!window.isSecureContext || !('serviceWorker' in navigator)) return null;
        try {
          await navigator.serviceWorker.register('/enviar-descarga-sw.js');
          const activo = (await navigator.serviceWorker.ready).active;
          if (!activo) return null;
          // Comprobar que atiende estas direcciones (si no, el navegador mostraría una descarga fallida).
          for (let i = 0; i < 20 && !navigator.serviceWorker.controller; i++) await new Promise((r) => setTimeout(r, 100));
          const sondeo = navigator.serviceWorker.controller ? await fetch('/enviar-descarga/sondeo/x').catch(() => null) : null;
          return sondeo && sondeo.headers.get('X-MC-SW') === '1' ? activo : null;
        } catch (e) {
          return null;
        }
      })();
    }
    return swListo;
  }
  async function descargaDirecta(meta) {
    let marco = null;
    try {
      const activo = await prepararSW();
      if (!activo) return null;
      const id = crypto.randomUUID();
      const canal = new MessageChannel();
      let puede = false, leyendo = false, cancelado = false, despertar = null;
      const avisar = () => { if (despertar) { const d = despertar; despertar = null; d(); } };
      canal.port1.onmessage = (e) => {
        if (e.data === 'pull') puede = true;
        else if (e.data === 'leyendo') leyendo = true;
        else if (e.data === 'cancelado') cancelado = true;
        avisar();
      };
      activo.postMessage({ tipo: 'preparar', id, nombre: meta.nombre, tam: meta.tam, mime: meta.tipo }, [canal.port2]);
      marco = document.createElement('iframe');
      marco.hidden = true;
      marco.src = '/enviar-descarga/' + id + '/' + encodeURIComponent(meta.nombre);
      document.body.append(marco);
      // Si el navegador no pide la descarga en 4 s, no sirve.
      const limite = Date.now() + 4000;
      while (!leyendo && Date.now() < limite) await new Promise((listo) => { despertar = listo; setTimeout(avisar, 100); });
      if (!leyendo) { marco.remove(); canal.port1.close(); return null; }
      const latido = setInterval(() => activo.postMessage({ tipo: 'ping' }), 10000);
      const quitar = () => { clearInterval(latido); setTimeout(() => marco.remove(), 60000); };
      return {
        async escribir(bytes) {
          while (!puede && !cancelado) await new Promise((listo) => { despertar = listo; });
          if (cancelado) throw new Error('cancelado');
          puede = false;
          canal.port1.postMessage(bytes.buffer, [bytes.buffer]);
        },
        async cerrar() { canal.port1.postMessage('fin'); quitar(); },
        async abortar() { canal.port1.postMessage('abortar'); quitar(); },
      };
    } catch (e) {
      if (marco) marco.remove();
      return null;
    }
  }
  // Espacio libre del navegador para guardar (o null si no lo dice).
  async function espacioLibre() {
    try { const e = await navigator.storage.estimate(); return e.quota - e.usage; } catch (e) { return null; }
  }

  // Recibe lo que manda el otro equipo. `alOferta(lista)` devuelve (o promete)
  // true para aceptar; `alArchivo(File)` por cada archivo terminado; el resto
  // son avisos. Devuelve una promesa con 'listo', 'rechazado' o 'cancelado'.
  function recibirArchivos(sesion, { alOferta, alProgreso, alArchivo, alGrande, sinDescargaDirecta }) {
    const canal = sesion.canal;
    let cierreFn = null, manejador = null;
    const quitar = () => { canal.recibiendo = false; if (cierreFn) canal.cierres.delete(cierreFn); if (canal.receptor === manejador) canal.receptor = null; };
    return new Promise((fin, error) => {
      let lista = [], actual = null, totalHecho = 0, totalTam = 0;
      // Tras una cancelación siguen llegando los trozos que el emisor ya había mandado: se ignoran.
      let anulado = false;
      const soltar = () => { anulado = true; if (actual && actual.destino) { actual.destino.abortar().catch(() => {}); } actual = null; };
      cierreFn = () => { soltar(); error(new Error('cerrado')); };
      canal.cierres.add(cierreFn);
      manejador = async (tipo, datos) => {
        if (anulado) return;
        if (tipo === TIPO_TROZO) {
          const id = new DataView(datos.buffer, datos.byteOffset).getUint32(0);
          if (!actual || actual.id !== id) throw new Error('trozo inesperado');
          const parte = datos.slice(4);
          // El largo se toma antes: la descarga directa transfiere el buffer y `parte` queda vacía.
          const largo = parte.length;
          if (actual.destino) {
            try { await actual.destino.escribir(parte); } catch (e) {
              // Si la persona canceló la descarga en el navegador se avisa al emisor; si no, es el almacenamiento.
              const cancelo = e.message === 'cancelado';
              soltar();
              if (cancelo) { canal.enviarControl({ t: 'cancela' }).catch(() => {}); fin('cancelado'); } else { error(new Error('almacenamiento')); canal.cerrar(); }
              return;
            }
          } else actual.partes.push(parte);
          actual.hecho += largo; totalHecho += largo;
          if (alProgreso) alProgreso({ nombre: actual.nombre, hecho: actual.hecho, tam: actual.tam, totalHecho, totalTam });
          return;
        }
        const m = JSON.parse(new TextDecoder().decode(datos));
        if (m.t === 'lista') {
          // Si este equipo está enviando, no puede recibir a la vez.
          if (canal.emisor) { await canal.enviarControl({ t: 'ocupado' }); return; }
          canal.recibiendo = true;
          lista = m.archivos;
          totalTam = lista.reduce((s, f) => s + f.tam, 0);
          const ok = await alOferta(lista);
          await canal.enviarControl({ t: ok ? 'acepta' : 'rechaza' });
          if (!ok) fin('rechazado');
        } else if (m.t === 'inicio') {
          const f = lista.find((x) => x.id === m.id);
          if (!f) throw new Error('archivo desconocido');
          actual = { id: f.id, nombre: f.nombre, tam: f.tam, tipo: f.tipo, partes: [], hecho: 0, destino: null };
          if (f.tam > UMBRAL_DISCO) {
            // `alGrande(meta, iniciar)`: la página pide el toque de la persona y llama a `iniciar()` desde ese toque.
            const iniciar = () => descargaDirecta(actual);
            const directo = sinDescargaDirecta ? null : await (alGrande ? alGrande(actual, iniciar) : iniciar());
            actual.destino = directo || (await abrirTemporal());
          }
          actual.directa = !!actual.destino && !actual.destino.manija;
          // Avisa que empezó el archivo: la barra aparece ya, sin esperar el primer trozo.
          if (alProgreso) alProgreso({ nombre: actual.nombre, hecho: 0, tam: actual.tam, totalHecho, totalTam });
          await canal.enviarControl({ t: 'sigue', id: f.id }); // el emisor espera esto para mandar el archivo
        } else if (m.t === 'fin') {
          if (!actual || actual.id !== m.id || actual.hecho !== actual.tam) throw new Error('archivo incompleto');
          if (actual.destino) await actual.destino.cerrar();
          // `destino`: 'descarga' (ya bajó a la carpeta de descargas), 'disco' (File
          // respaldado en el disco del sitio) o 'memoria'.
          if (actual.directa) alArchivo({ name: actual.nombre, size: actual.tam, type: actual.tipo }, { destino: 'descarga' });
          else if (actual.destino) alArchivo(new File([await actual.destino.manija.getFile()], actual.nombre, { type: actual.tipo }), { destino: 'disco' });
          else alArchivo(new File(actual.partes, actual.nombre, { type: actual.tipo }), { destino: 'memoria' });
          actual = null;
        } else if (m.t === 'listo') {
          await canal.enviarControl({ t: 'ok' });
          fin('listo');
        } else if (m.t === 'cancela') {
          soltar();
          fin('cancelado');
        }
      };
      canal.receptor = manejador;
    }).finally(quitar);
  }

  // Queda recibiendo todo lo que el otro equipo quiera mandar, una transferencia
  // tras otra. `alTerminar(fin)` se llama al terminar cada una y `alError(e)` si
  // el canal se cierra o falla el almacenamiento.
  function escucharEnvios(sesion, opciones) {
    const ciclo = () => recibirArchivos(sesion, opciones).then(
      (fin) => { if (opciones.alTerminar) opciones.alTerminar(fin); if (!sesion.canal.cerrado) ciclo(); },
      (e) => { if (opciones.alError) opciones.alError(e); });
    ciclo();
  }

  return { anfitrion, invitado, leerCodigo, esProtegido, enviarArchivos, recibirArchivos, escucharEnvios, configurar(o) { if (o.inactividadMs) inactividadMs = o.inactividadMs; if (o.revisionMs) revisionMs = o.revisionMs; }, limpiarTemporales, espacioLibre, prepararSW, UMBRAL_DISCO, TROZO };
})();

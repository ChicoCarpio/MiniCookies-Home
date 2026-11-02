/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Motor del administrador de contraseñas: la caja fuerte cifrada, su
// almacenamiento local y la importación y exportación de contraseñas.
// Sin servidor ni librerías de terceros: todo con WebCrypto.
//
// Cómo se protege:
// - La contraseña maestra nunca se guarda ni sale del equipo. Se normaliza
//   (NFKC, para que una tilde escrita en otro teclado dé la misma clave) y se
//   convierte en una clave AES de 256 bits con PBKDF2-SHA256 (2.000.000
//   vueltas) y una sal aleatoria de 32 bytes. La clave no es extraíble.
// - El contenido (todas las entradas) se cifra de una vez con AES-GCM de 256
//   bits y un IV nuevo de 12 bytes en cada guardado. Los datos de cabecera
//   (formato, versión, vueltas, sal, nombre) van como datos asociados: si
//   alguien los cambia, el descifrado falla en vez de aceptarse en silencio.
// - Si se olvida la contraseña maestra no hay forma de recuperar nada: no
//   existe una copia de la clave en ningún lado.
//
// Qué NO protege: un equipo con un programa espía, una contraseña maestra
// débil, ni lo que ya esté copiado al portapapeles o visible en pantalla.
//
// Piezas:
//   BovedaContrasenas.crear(maestra, nombre) → sesión
//   BovedaContrasenas.abrir(archivo, maestra) → sesión (archivo: objeto o texto JSON)
//   sesión: nombre, revision, entradas(), obtener(id), agregar(datos),
//     actualizar(id, cambios), borrar(id), renombrar(nombre), cambiarMaestra(nueva),
//     cifrar() → objeto de archivo, cerrar()
//   BovedaContrasenas.Almacen: listar / leer / guardar / borrar (IndexedDB; solo
//     guarda el contenido ya cifrado)
//   BovedaContrasenas.leerCsv(texto) / aCsv(entradas, formato) /
//     compararImportadas(existentes, importadas) / auditar(entradas)
// Los errores llevan `codigo`: 'maestra-corta', 'maestra-incorrecta', 'formato',
// 'version', 'danado', 'sin-almacen'.
var BovedaContrasenas = (function () {
  const FORMATO = 'minicookies-boveda';
  const VERSION = 1;
  const KDF = 'PBKDF2-SHA256';
  const VUELTAS = 2000000;
  const VUELTAS_MIN = 600000;
  const VUELTAS_MAX = 10000000;
  const MAESTRA_MIN = 12;
  const SAL_BYTES = 32;
  const MAX_FILAS_IMPORTAR = 20000;
  const BLOQUE_RELLENO = 2048;
  const CAMPOS = ['sitio', 'usuario', 'contrasena', 'url', 'notas'];

  const codificar = new TextEncoder();
  const decodificar = new TextDecoder();

  function error(codigo, mensaje) {
    return Object.assign(new Error(mensaje), { codigo });
  }

  function aBase64(bytes) {
    let s = '';
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  }

  function deBase64(texto) {
    const limpio = String(texto).replace(/-/g, '+').replace(/_/g, '/');
    const bin = atob(limpio + '='.repeat((4 - limpio.length % 4) % 4));
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }

  function aleatorios(n) {
    return crypto.getRandomValues(new Uint8Array(n));
  }

  function normalizarMaestra(maestra) {
    return String(maestra).normalize('NFKC');
  }

  // Con clave secreta, lo que entra a PBKDF2 es la maestra + un byte 0 + los 16
  // bytes del secreto: sin el secreto no se puede ni intentar adivinar la
  // maestra con el archivo robado.
  async function material(maestra, secreto) {
    const m = codificar.encode(normalizarMaestra(maestra));
    let bytes = m;
    if (secreto) { bytes = new Uint8Array(m.length + 1 + secreto.length); bytes.set(m); bytes.set(secreto, m.length + 1); }
    try {
      return await crypto.subtle.importKey('raw', bytes, 'PBKDF2', false, ['deriveKey', 'deriveBits']);
    } finally { bytes.fill(0); m.fill(0); }
  }

  async function derivarClave(maestra, sal, vueltas, secreto) {
    return crypto.subtle.deriveKey(
      { name: 'PBKDF2', hash: 'SHA-256', salt: sal, iterations: vueltas },
      await material(maestra, secreto), { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
  }

  function datosAsociados(nombre, vueltas, salB64, conSecreto) {
    return codificar.encode([FORMATO, VERSION, KDF, vueltas, salB64, conSecreto ? 1 : 0, nombre].join('|'));
  }

  // Clave secreta: 16 bytes aleatorios escritos en 26 letras y números sin
  // ambigüedades (sin I, O, 0 ni 1), en grupos de 4 para copiarla a mano.
  const ALFABETO = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';

  function secretoATexto(bytes) {
    let bits = 0, acum = 0, salida = '';
    for (const b of bytes) {
      acum = ((acum << 8) | b) & 0xffff; bits += 8;
      while (bits >= 5) { bits -= 5; salida += ALFABETO[(acum >> bits) & 31]; }
    }
    if (bits > 0) salida += ALFABETO[(acum << (5 - bits)) & 31];
    return salida.match(/.{1,4}/g).join('-');
  }

  function textoASecreto(texto) {
    const limpio = String(texto || '').toUpperCase().replace(/[\s-]/g, '');
    if (limpio.length !== 26) throw error('clave-secreta', 'La clave secreta no es válida.');
    const bytes = new Uint8Array(16);
    let bits = 0, acum = 0, n = 0;
    for (const c of limpio) {
      const v = ALFABETO.indexOf(c);
      if (v < 0) throw error('clave-secreta', 'La clave secreta no es válida.');
      acum = ((acum << 5) | v) & 0xffff; bits += 5;
      if (bits >= 8 && n < 16) { bits -= 8; bytes[n++] = (acum >> bits) & 255; }
    }
    return bytes;
  }

  function nuevaEntrada(datos) {
    const ahora = new Date().toISOString();
    const entrada = { id: aBase64(aleatorios(12)).replace(/[+/=]/g, 'x'), creado: ahora, modificado: ahora };
    for (const campo of CAMPOS) entrada[campo] = String(datos && datos[campo] != null ? datos[campo] : '');
    return entrada;
  }

  function normalizarEntrada(crudo) {
    const e = nuevaEntrada(crudo);
    if (typeof crudo.id === 'string' && crudo.id) e.id = crudo.id;
    if (typeof crudo.creado === 'string') e.creado = crudo.creado;
    if (typeof crudo.modificado === 'string') e.modificado = crudo.modificado;
    return e;
  }

  // ─── Sesión ───
  function crearSesion(estado) {
    let { nombre, clave, sal, vueltas, lista, secreto } = estado;
    let revision = 0;
    const abierta = () => { if (!lista) throw error('cerrada', 'La caja fuerte está cerrada.'); };
    const buscar = (id) => {
      abierta();
      const e = lista.find((x) => x.id === id);
      if (!e) throw error('no-existe', 'No existe esa entrada.');
      return e;
    };
    return {
      get nombre() { return nombre; },
      get revision() { return revision; },
      get conClaveSecreta() { return !!secreto; },
      // La clave secreta en texto, para mostrarla o imprimirla (quien la pide
      // tiene que haber verificado la maestra antes).
      get claveSecreta() { abierta(); return secreto ? secretoATexto(secreto) : null; },
      entradas() { abierta(); return lista.map((e) => Object.assign({}, e)); },
      obtener(id) { return Object.assign({}, buscar(id)); },
      agregar(datos) { abierta(); const e = nuevaEntrada(datos); lista.push(e); revision++; return Object.assign({}, e); },
      actualizar(id, cambios) {
        const e = buscar(id);
        for (const campo of CAMPOS) if (cambios && cambios[campo] != null) e[campo] = String(cambios[campo]);
        e.modificado = new Date().toISOString();
        revision++;
        return Object.assign({}, e);
      },
      borrar(id) { const e = buscar(id); lista.splice(lista.indexOf(e), 1); revision++; },
      renombrar(nuevo) { abierta(); nombre = String(nuevo).trim().slice(0, 80) || nombre; revision++; },
      async cambiarMaestra(nueva) {
        abierta();
        if (String(nueva).length < MAESTRA_MIN) throw error('maestra-corta', 'La contraseña maestra es muy corta.');
        const salNueva = aleatorios(SAL_BYTES);
        const claveNueva = await derivarClave(nueva, salNueva, VUELTAS, secreto);
        abierta(); // si la caja se cerró mientras se derivaba, no se deja viva la clave nueva
        clave = claveNueva;
        sal = salNueva;
        vueltas = VUELTAS;
        revision++;
      },
      // Devuelve el archivo cifrado (objeto listo para JSON.stringify). Cada
      // llamada usa un IV nuevo.
      async cifrar() {
        abierta();
        const salB64 = aBase64(sal);
        const iv = aleatorios(12);
        // Se rellena hasta la potencia de 2 siguiente (mínimo 2 KB): el tamaño del
        // archivo solo da una idea aproximada de cuántas contraseñas hay.
        const largo = codificar.encode(JSON.stringify({ v: 1, entradas: lista, r: '' })).length;
        let objetivo = BLOQUE_RELLENO;
        while (objetivo < largo) objetivo *= 2;
        const texto = codificar.encode(JSON.stringify({ v: 1, entradas: lista, r: 'x'.repeat(objetivo - largo) }));
        const cifrado = await crypto.subtle.encrypt(
          { name: 'AES-GCM', iv, additionalData: datosAsociados(nombre, vueltas, salB64, !!secreto) }, clave, texto);
        texto.fill(0);
        return {
          formato: FORMATO, version: VERSION, nombre, kdf: KDF, iteraciones: vueltas, sal: salB64, claveSecreta: !!secreto,
          cifrado: 'AES-256-GCM', iv: aBase64(iv), datos: aBase64(new Uint8Array(cifrado)),
          actualizado: new Date().toISOString(),
        };
      },
      cerrar() { lista = null; clave = null; sal = null; if (secreto) secreto.fill(0); secreto = null; },
    };
  }

  // opciones.claveSecreta: true agrega una clave secreta aleatoria además de la
  // maestra (la sesión la devuelve en claveSecreta para mostrarla una vez).
  async function crear(maestra, nombre, opciones) {
    if (String(maestra).length < MAESTRA_MIN) throw error('maestra-corta', 'La contraseña maestra es muy corta.');
    const sal = aleatorios(SAL_BYTES);
    const secreto = opciones && opciones.claveSecreta ? aleatorios(16) : null;
    const clave = await derivarClave(maestra, sal, VUELTAS, secreto);
    return crearSesion({ nombre: String(nombre || '').trim().slice(0, 80) || 'Mi caja fuerte', clave, sal, vueltas: VUELTAS, lista: [], secreto });
  }

  function leerArchivo(archivo) {
    let a = archivo;
    if (typeof a === 'string') {
      try { a = JSON.parse(a); } catch (e) { throw error('formato', 'El archivo no es una caja fuerte de Mini Cookies.'); }
    }
    if (!a || typeof a !== 'object' || a.formato !== FORMATO) throw error('formato', 'El archivo no es una caja fuerte de Mini Cookies.');
    if (a.version !== VERSION) throw error('version', 'La caja fuerte es de una versión que esta página no conoce.');
    const ok = a.kdf === KDF && a.cifrado === 'AES-256-GCM' && typeof a.nombre === 'string' &&
      Number.isInteger(a.iteraciones) && a.iteraciones >= VUELTAS_MIN && a.iteraciones <= VUELTAS_MAX &&
      typeof a.sal === 'string' && typeof a.iv === 'string' && typeof a.datos === 'string' &&
      (a.claveSecreta === undefined || typeof a.claveSecreta === 'boolean');
    if (!ok) throw error('danado', 'El archivo de la caja fuerte está dañado.');
    return a;
  }

  // Si la caja tiene clave secreta y no se da (o está mal escrita) lanza
  // 'clave-secreta' antes de gastar las vueltas de PBKDF2.
  async function abrir(archivo, maestra, claveSecreta) {
    const a = leerArchivo(archivo);
    const conSecreto = a.claveSecreta === true;
    const secreto = conSecreto ? (claveSecreta instanceof Uint8Array ? claveSecreta.slice() : textoASecreto(claveSecreta)) : null;
    let sal, iv, datos;
    try { sal = deBase64(a.sal); iv = deBase64(a.iv); datos = deBase64(a.datos); } catch (e) { throw error('danado', 'El archivo de la caja fuerte está dañado.'); }
    if (sal.length < 16 || sal.length > 64 || iv.length !== 12) throw error('danado', 'El archivo de la caja fuerte está dañado.');
    const clave = await derivarClave(maestra, sal, a.iteraciones, secreto);
    let plano;
    try {
      plano = await crypto.subtle.decrypt({ name: 'AES-GCM', iv, additionalData: datosAsociados(a.nombre, a.iteraciones, a.sal, conSecreto) }, clave, datos);
    } catch (e) {
      throw error('maestra-incorrecta', 'La contraseña maestra es incorrecta (o el archivo fue modificado).');
    }
    let lista;
    try {
      const contenido = JSON.parse(decodificar.decode(plano));
      lista = contenido.entradas.map(normalizarEntrada);
    } catch (e) { throw error('danado', 'El archivo de la caja fuerte está dañado.'); }
    return crearSesion({ nombre: a.nombre, clave, sal, vueltas: a.iteraciones, lista, secreto });
  }

  // Pasa la clave secreta de texto a bytes (para recordarla en este dispositivo).
  function claveSecretaABytes(texto) { return textoASecreto(texto); }

  // ─── Almacén local (IndexedDB): solo guarda contenido ya cifrado ───
  const Almacen = (function () {
    let abierta = null;
    function db() {
      if (!abierta) {
        abierta = new Promise((resolver, rechazar) => {
          if (typeof indexedDB === 'undefined') return rechazar(error('sin-almacen', 'Este navegador no permite guardar datos locales.'));
          let pedido;
          try { pedido = indexedDB.open('minicookies-boveda', 1); } catch (e) { return rechazar(error('sin-almacen', 'Este navegador no permite guardar datos locales.')); }
          pedido.onupgradeneeded = () => pedido.result.createObjectStore('cajas', { keyPath: 'id' });
          pedido.onsuccess = () => resolver(pedido.result);
          pedido.onerror = () => rechazar(error('sin-almacen', 'Este navegador no permite guardar datos locales.'));
        });
        abierta.catch(() => { abierta = null; });
      }
      return abierta;
    }
    async function pedir(modo, accion) {
      const base = await db();
      return new Promise((resolver, rechazar) => {
        const tx = base.transaction('cajas', modo);
        const r = accion(tx.objectStore('cajas'));
        tx.oncomplete = () => resolver(r && r.result);
        tx.onerror = tx.onabort = () => rechazar(error('sin-almacen', 'No se pudo guardar en este navegador.'));
      });
    }
    return {
      async listar() {
        const todas = (await pedir('readonly', (s) => s.getAll())) || [];
        return todas.map((c) => ({ id: c.id, nombre: c.archivo.nombre, actualizado: c.archivo.actualizado }))
          .sort((a, b) => String(b.actualizado).localeCompare(String(a.actualizado)));
      },
      async leer(id) { const c = await pedir('readonly', (s) => s.get(id)); return c ? c.archivo : null; },
      async guardar(id, archivo) { await pedir('readwrite', (s) => s.put({ id, archivo })); },
      async borrar(id) { await pedir('readwrite', (s) => s.delete(id)); },
      nuevoId() { return aBase64(aleatorios(9)).replace(/[+/=]/g, 'x'); },
    };
  })();

  // ─── CSV de navegadores (Chrome, Edge y Firefox) ───
  function analizarCsv(texto) {
    const filas = [];
    let fila = [], campo = '', entreComillas = false;
    const s = String(texto).replace(/^﻿/, '');
    for (let i = 0; i < s.length; i++) {
      const c = s[i];
      if (entreComillas) {
        if (c === '"') { if (s[i + 1] === '"') { campo += '"'; i++; } else entreComillas = false; } else campo += c;
      } else if (c === '"') entreComillas = true;
      else if (c === ',') { fila.push(campo); campo = ''; }
      else if (c === '\n' || c === '\r') {
        if (c === '\r' && s[i + 1] === '\n') i++;
        fila.push(campo); campo = '';
        if (fila.length > 1 || fila[0] !== '') filas.push(fila);
        fila = [];
        if (filas.length > MAX_FILAS_IMPORTAR + 1) throw error('demasiadas', 'El archivo tiene demasiadas filas.');
      } else campo += c;
    }
    if (campo !== '' || fila.length) { fila.push(campo); if (fila.length > 1 || fila[0] !== '') filas.push(fila); }
    return filas;
  }

  function nombreDeSitio(url) {
    try { return new URL(url).hostname.replace(/^www\./, '') || url; } catch (e) { return url; }
  }

  // Devuelve entradas parciales { sitio, usuario, contrasena, url, notas }.
  function leerCsv(texto) {
    const filas = analizarCsv(texto);
    if (filas.length > MAX_FILAS_IMPORTAR + 1) throw error('demasiadas', 'El archivo tiene demasiadas filas.');
    if (!filas.length) throw error('formato', 'El archivo no tiene el formato de Chrome, Edge o Firefox.');
    const cab = filas[0].map((c) => c.trim().toLowerCase());
    const col = (n) => cab.indexOf(n);
    if (col('url') < 0 || col('username') < 0 || col('password') < 0) throw error('formato', 'El archivo no tiene el formato de Chrome, Edge o Firefox.');
    const dato = (fila, n) => (col(n) >= 0 ? (fila[col(n)] || '') : '');
    const salida = [];
    for (const fila of filas.slice(1)) {
      const url = dato(fila, 'url').trim();
      const usuario = dato(fila, 'username').trim();
      const contrasena = dato(fila, 'password');
      const sitio = dato(fila, 'name').trim() || nombreDeSitio(url);
      if (!sitio && !usuario && !contrasena) continue;
      salida.push({ sitio, usuario, contrasena, url, notas: dato(fila, 'note') });
    }
    return salida;
  }

  // Excel y LibreOffice convierten en fórmula lo que empieza con = + - @ (o con
  // tabulación o retorno): en los campos de texto se antepone una comilla. La
  // contraseña NO se toca (el navegador la importaría con la comilla puesta).
  function campoCsv(valor, esTexto) {
    let s = String(valor == null ? '' : valor);
    if (esTexto && /^[=+\-@\t\r]/.test(s)) s = "'" + s;
    return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }

  // formato: 'chrome' (también Edge) o 'firefox'. El resultado es texto plano:
  // quien lo pide tiene que avisar a la persona.
  function aCsv(entradas, formato) {
    const filas = [];
    if (formato === 'firefox') {
      filas.push(['url', 'username', 'password', 'httpRealm', 'formActionOrigin', 'guid', 'timeCreated', 'timeLastUsed', 'timePasswordChanged']);
      for (const e of entradas) filas.push([e.url, e.usuario, e.contrasena, '', '', '', '', '', '']);
      filas.textos = [0, 1];
    } else {
      filas.push(['name', 'url', 'username', 'password', 'note']);
      for (const e of entradas) filas.push([e.sitio, e.url, e.usuario, e.contrasena, e.notas]);
      filas.textos = [0, 1, 2, 4];
    }
    return filas.map((f, i) => f.map((c, j) => campoCsv(c, i > 0 && filas.textos.includes(j))).join(',')).join('\r\n') + '\r\n';
  }

  // Separa lo importado en nuevas, iguales (mismo sitio, usuario y contraseña:
  // se omiten) y conflictos (mismo sitio y usuario con otra contraseña, u otros
  // datos: la persona elige). No modifica nada.
  function compararImportadas(existentes, importadas) {
    const clave = (e) => (e.sitio || '').trim().toLowerCase() + '\n' + (e.usuario || '').trim().toLowerCase();
    const porClave = new Map(existentes.map((e) => [clave(e), e]));
    const nuevas = [], iguales = [], conflictos = [];
    for (const importada of importadas) {
      const existente = porClave.get(clave(importada));
      if (!existente) nuevas.push(importada);
      else if (existente.contrasena === importada.contrasena && (existente.url || '') === (importada.url || '')) iguales.push(importada);
      else conflictos.push({ existente, importada });
    }
    return { nuevas, iguales, conflictos };
  }

  // Revisión de la caja: contraseñas repetidas en varias entradas y débiles.
  function auditar(entradas) {
    const usos = new Map();
    for (const e of entradas) {
      if (!e.contrasena) continue;
      const ids = usos.get(e.contrasena);
      if (ids) ids.push(e.id); else usos.set(e.contrasena, [e.id]);
    }
    const reutilizadas = [];
    for (const ids of usos.values()) if (ids.length > 1) for (const id of ids) reutilizadas.push(id);
    const debiles = entradas.filter((e) => e.contrasena && GeneradorContrasenas.fortaleza(e.contrasena) === 'debil').map((e) => e.id);
    return { reutilizadas, debiles };
  }

  return { crear, abrir, claveSecretaABytes, MAX_FILAS_IMPORTAR, Almacen, leerCsv, aCsv, compararImportadas, auditar, MAESTRA_MIN, VUELTAS };
})();

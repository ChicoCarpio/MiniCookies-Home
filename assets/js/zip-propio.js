/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. */
// Lector y escritor de ZIP propio (reemplaza a JSZip). Un .docx, .xlsx, .pptx
// o .epub es un zip por dentro; todas las herramientas que los abren o los
// guardan pasan por acá.
//
// La compresión no se programa: la hace el navegador con DecompressionStream /
// CompressionStream('deflate-raw'). Acá solo está el formato del contenedor
// (directorio central, cabeceras locales, CRC32) y las protecciones.
//
// La API es el subconjunto de JSZip que el sitio usa, con los mismos nombres,
// para que migrar una página sea cambiar el nombre de la clase:
//   const zip = await ZipPropio.loadAsync(archivoOBytes);
//   const entrada = zip.file('ruta');            // null si no existe o es carpeta
//   await entrada.async('string' | 'text' | 'uint8array' | 'arraybuffer' | 'base64' | 'blob' | 'binarystring');
//   zip.files                                    // { ruta: entrada } (entrada.dir, entrada.name)
//   zip.forEach((rutaRelativa, entrada) => …);  zip.remove(ruta);  zip.folder('x')
//   zip.file('ruta', datos, { compression: 'STORE', base64: true });
//   await zip.generateAsync({ type: 'blob' | 'uint8array' | 'arraybuffer' | 'base64', compression: 'DEFLATE', mimeType });
//
// Lo que no hace (se rechaza con un ErrorZip claro): zip cifrado, ZIP64 y
// métodos de compresión distintos de STORE y DEFLATE. En Office y EPUB reales
// no aparecen; si aparece un caso, se agrega.
//
// Protecciones: el tamaño que declara el archivo se verifica contra los bytes
// reales al descomprimir (una "zip bomba" que miente en el tamaño se corta), y
// hay topes de entradas y de tamaño.
(function () {
  'use strict';

  const LIMITE_ENTRADAS = 100000;
  const LIMITE_POR_ENTRADA = 1024 * 1024 * 1024;
  const LIMITE_TOTAL = 4 * 1024 * 1024 * 1024;
  const MAXIMO_ZIP = 0xFFFFFFFF;

  // codigo: 'invalido' | 'danado' | 'cifrado' | 'zip64' | 'metodo' | 'grande'
  class ErrorZip extends Error {
    constructor(codigo, mensaje) {
      super(mensaje);
      this.name = 'ErrorZip';
      this.codigo = codigo;
    }
  }

  // ---------- CRC32 ----------
  const TABLA_CRC = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
    TABLA_CRC[n] = c >>> 0;
  }
  function crc32(bytes) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) c = TABLA_CRC[(c ^ bytes[i]) & 255] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  // ---------- Bytes y texto ----------
  const CODIFICADOR = new TextEncoder();
  // ignoreBOM: se conserva el BOM como JSZip, para no cambiar lo que ven las herramientas.
  const DECODIFICADOR = new TextDecoder('utf-8', { ignoreBOM: true });
  const DECODIFICADOR_ESTRICTO = new TextDecoder('utf-8', { fatal: true, ignoreBOM: true });

  // Mitad alta de la página de códigos 437 (nombres de zips viejos sin el bit UTF-8).
  const CP437 = 'ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜ¢£¥₧ƒáíóúñÑªº¿⌐¬½¼¡«»░▒▓│┤╡╢╖╕╣║╗╝╜╛┐└┴┬├─┼╞╟╚╔╩╦╠═╬╧╨╤╥╙╘╒╓╫╪┘┌█▄▌▐▀αßΓπΣσµτΦΘΩδ∞φε∩≡±≥≤⌠⌡÷≈°∙·√ⁿ²■ ';

  function nombreDe(bytes, utf8) {
    if (utf8) return DECODIFICADOR.decode(bytes);
    try { return DECODIFICADOR_ESTRICTO.decode(bytes); } catch (e) { /* no es UTF-8: CP437 */ }
    let s = '';
    for (const b of bytes) s += b < 128 ? String.fromCharCode(b) : CP437[b - 128];
    return s;
  }

  function unir(partes, total) {
    if (total === undefined) total = partes.reduce((suma, p) => suma + p.length, 0);
    const salida = new Uint8Array(total);
    let pos = 0;
    for (const p of partes) { salida.set(p, pos); pos += p.length; }
    return salida;
  }

  function aBase64(bytes) {
    let s = '';
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  }
  function deBase64(texto) {
    const s = atob(String(texto).replace(/^data:[^,]*,/, ''));
    const bytes = new Uint8Array(s.length);
    for (let i = 0; i < s.length; i++) bytes[i] = s.charCodeAt(i);
    return bytes;
  }
  function deBinaria(texto) {
    const bytes = new Uint8Array(texto.length);
    for (let i = 0; i < texto.length; i++) bytes[i] = texto.charCodeAt(i) & 255;
    return bytes;
  }

  // ---------- Compresión nativa ----------
  function flujoDe(bytes) {
    return new ReadableStream({ start(c) { if (bytes.length) c.enqueue(bytes); c.close(); } });
  }

  async function inflar(bytes, esperado) {
    const lector = flujoDe(bytes).pipeThrough(new DecompressionStream('deflate-raw')).getReader();
    const partes = [];
    let total = 0;
    try {
      for (;;) {
        const { done, value } = await lector.read();
        if (done) break;
        total += value.length;
        // Más bytes de los que declara la entrada: tamaño falso (zip bomba) o dañado.
        if (total > esperado) throw new ErrorZip('danado', 'Corrupted zip: an entry is larger than it declares');
        partes.push(value);
      }
    } catch (error) {
      if (error instanceof ErrorZip) { lector.cancel().catch(() => {}); throw error; }
      throw new ErrorZip('danado', 'Corrupted zip: an entry cannot be decompressed');
    }
    if (total !== esperado) throw new ErrorZip('danado', 'Corrupted zip: an entry has a different size than it declares');
    return partes.length === 1 ? partes[0] : unir(partes, total);
  }

  async function deflar(bytes) {
    const lector = flujoDe(bytes).pipeThrough(new CompressionStream('deflate-raw')).getReader();
    const partes = [];
    for (;;) {
      const { done, value } = await lector.read();
      if (done) break;
      partes.push(value);
    }
    return partes.length === 1 ? partes[0] : unir(partes);
  }

  // ---------- Lectura del directorio ----------
  function leerDirectorio(bytes) {
    const vista = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let fin = -1;
    for (let i = bytes.length - 22; i >= 0 && i >= bytes.length - 22 - 65535; i--) {
      if (vista.getUint32(i, true) === 0x06054b50) { fin = i; break; }
    }
    if (fin < 0) throw new ErrorZip('invalido', 'Not a zip file (end of central directory not found)');
    const total = vista.getUint16(fin + 10, true);
    const tamanoDir = vista.getUint32(fin + 12, true);
    const desdeDir = vista.getUint32(fin + 16, true);
    if (total === 0xFFFF || tamanoDir === MAXIMO_ZIP || desdeDir === MAXIMO_ZIP) {
      throw new ErrorZip('zip64', 'Unsupported zip: ZIP64 archives are not supported');
    }
    if (total > LIMITE_ENTRADAS) throw new ErrorZip('grande', 'Zip with too many entries');
    // Bytes extra antes del zip (instaladores, adjuntos): el directorio está más
    // adelante de lo que dice su propio desplazamiento.
    const desplazamiento = fin - tamanoDir - desdeDir;
    if (desplazamiento < 0) throw new ErrorZip('danado', 'Corrupted zip: the central directory is out of place');

    const entradas = [];
    let pos = desdeDir + desplazamiento;
    let declarado = 0;
    for (let i = 0; i < total; i++) {
      if (pos + 46 > bytes.length || vista.getUint32(pos, true) !== 0x02014b50) {
        throw new ErrorZip('danado', 'Corrupted zip: bad central directory');
      }
      const banderas = vista.getUint16(pos + 8, true);
      const metodo = vista.getUint16(pos + 10, true);
      const hora = vista.getUint16(pos + 12, true);
      const fecha = vista.getUint16(pos + 14, true);
      const crc = vista.getUint32(pos + 16, true);
      const comprimido = vista.getUint32(pos + 20, true);
      const real = vista.getUint32(pos + 24, true);
      const largoNombre = vista.getUint16(pos + 28, true);
      const largoExtra = vista.getUint16(pos + 30, true);
      const largoComentario = vista.getUint16(pos + 32, true);
      const local = vista.getUint32(pos + 42, true);
      if (pos + 46 + largoNombre > bytes.length) throw new ErrorZip('danado', 'Corrupted zip: bad central directory');
      const nombre = nombreDe(bytes.subarray(pos + 46, pos + 46 + largoNombre), (banderas & 0x800) !== 0);
      pos += 46 + largoNombre + largoExtra + largoComentario;

      if ((banderas & 1) || metodo === 99) throw new ErrorZip('cifrado', 'Encrypted zip is not supported');
      if (comprimido === MAXIMO_ZIP || real === MAXIMO_ZIP || local === MAXIMO_ZIP) {
        throw new ErrorZip('zip64', 'Unsupported zip: ZIP64 archives are not supported');
      }
      const carpeta = nombre.endsWith('/');
      if (!carpeta && metodo !== 0 && metodo !== 8) {
        throw new ErrorZip('metodo', 'Unsupported zip: compression method ' + metodo + ' is not supported');
      }
      if (real > LIMITE_POR_ENTRADA) throw new ErrorZip('grande', 'Zip entry too large');
      declarado += real;
      if (declarado > LIMITE_TOTAL) throw new ErrorZip('grande', 'Zip too large once decompressed');
      entradas.push({ nombre, carpeta, metodo, hora, fecha, crc, comprimido, real, local: local + desplazamiento });
    }
    return entradas;
  }

  // Los bytes comprimidos de una entrada, cortados con su tamaño comprimido
  // (el del directorio central: sirve también si la cabecera local usa data descriptor).
  function bytesCrudos(origen) {
    const { bytes, local, comprimido } = origen;
    if (local + 30 > bytes.length) throw new ErrorZip('danado', 'Corrupted zip: bad local header');
    const vista = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    if (vista.getUint32(local, true) !== 0x04034b50) throw new ErrorZip('danado', 'Corrupted zip: bad local header');
    const inicio = local + 30 + vista.getUint16(local + 26, true) + vista.getUint16(local + 28, true);
    if (inicio + comprimido > bytes.length) throw new ErrorZip('danado', 'Corrupted zip: truncated entry');
    return bytes.subarray(inicio, inicio + comprimido);
  }

  // ---------- Entrada ----------
  class Entrada {
    constructor(nombre, dir) {
      this.name = nombre;
      this.dir = dir;
      this.date = new Date();
      this._datos = undefined;   // contenido nuevo: Uint8Array, Blob o Promise
      this._origen = null;       // contenido leído de un zip: ver leerDirectorio
      this._compresion = null;   // 'STORE' | 'DEFLATE' pedido para esta entrada
    }

    async _bytes() {
      if (this.dir) return new Uint8Array(0);
      if (this._origen) {
        const o = this._origen;
        const crudos = bytesCrudos(o);
        if (o.metodo === 0) {
          if (crudos.length !== o.real) throw new ErrorZip('danado', 'Corrupted zip: an entry has a different size than it declares');
          return crudos;
        }
        return inflar(crudos, o.real);
      }
      let d = await this._datos;
      if (d instanceof Blob) d = new Uint8Array(await d.arrayBuffer());
      else if (d instanceof ArrayBuffer) d = new Uint8Array(d);
      return d;
    }

    async async(tipo) {
      const bytes = await this._bytes();
      switch (String(tipo).toLowerCase()) {
        case 'string': case 'text': return DECODIFICADOR.decode(bytes);
        case 'uint8array': case 'nodebuffer': return bytes;
        case 'arraybuffer': return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
        case 'base64': return aBase64(bytes);
        case 'blob': return new Blob([bytes]);
        case 'binarystring': {
          let s = '';
          for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
          return s;
        }
        case 'array': return Array.from(bytes);
        default: throw new Error('Unsupported type: ' + tipo);
      }
    }
  }

  function datosDe(datos, opciones) {
    if (typeof datos === 'string') {
      if (opciones.base64) return deBase64(datos);
      if (opciones.binary) return deBinaria(datos);
      return CODIFICADOR.encode(datos);
    }
    if (datos instanceof ArrayBuffer) return new Uint8Array(datos);
    if (ArrayBuffer.isView(datos)) return new Uint8Array(datos.buffer, datos.byteOffset, datos.byteLength);
    if (Array.isArray(datos)) return Uint8Array.from(datos);
    return datos; // Blob o Promise: se resuelven al leer o generar
  }

  // ---------- Zip ----------
  class ZipPropio {
    constructor() {
      this.files = Object.create(null);
      this._prefijo = '';
    }

    static async loadAsync(datos, opciones) {
      return new ZipPropio().loadAsync(datos, opciones);
    }

    async loadAsync(datos) {
      datos = await datos;
      let bytes;
      if (datos instanceof Blob) bytes = new Uint8Array(await datos.arrayBuffer());
      else if (datos instanceof ArrayBuffer) bytes = new Uint8Array(datos);
      else if (ArrayBuffer.isView(datos)) bytes = new Uint8Array(datos.buffer, datos.byteOffset, datos.byteLength);
      else throw new ErrorZip('invalido', 'Not a zip file (unsupported input)');
      for (const e of leerDirectorio(bytes)) {
        const entrada = new Entrada(e.nombre, e.carpeta);
        entrada.date = fechaDeDos(e.fecha, e.hora);
        if (!e.carpeta) entrada._origen = { bytes, ...e };
        this.files[e.nombre] = entrada;
      }
      return this;
    }

    file(nombre, datos, opciones) {
      if (arguments.length === 1) {
        const e = this.files[this._prefijo + nombre];
        return e && !e.dir ? e : null;
      }
      opciones = opciones || {};
      const ruta = this._prefijo + nombre;
      const carpeta = opciones.dir || ruta.endsWith('/');
      const entrada = new Entrada(carpeta && !ruta.endsWith('/') ? ruta + '/' : ruta, !!carpeta);
      if (opciones.date) entrada.date = opciones.date;
      if (!carpeta) {
        entrada._datos = datosDe(datos, opciones);
        if (opciones.compression) entrada._compresion = String(opciones.compression).toUpperCase() === 'DEFLATE' ? 'DEFLATE' : 'STORE';
      }
      this.files[entrada.name] = entrada;
      return this;
    }

    folder(nombre) {
      const ruta = this._prefijo + nombre.replace(/\/?$/, '/');
      if (!this.files[ruta]) this.files[ruta] = new Entrada(ruta, true);
      const vista = Object.create(this);
      vista._prefijo = ruta;
      return vista;
    }

    forEach(devolucion) {
      for (const ruta of Object.keys(this.files)) {
        if (!ruta.startsWith(this._prefijo) || ruta === this._prefijo) continue;
        devolucion(ruta.slice(this._prefijo.length), this.files[ruta]);
      }
    }

    remove(nombre) {
      const ruta = this._prefijo + nombre.replace(/\/$/, '');
      delete this.files[ruta];
      delete this.files[ruta + '/'];
      for (const clave of Object.keys(this.files)) if (clave.startsWith(ruta + '/')) delete this.files[clave];
      return this;
    }

    async generateAsync(opciones) {
      opciones = opciones || {};
      const tipo = String(opciones.type || 'uint8array').toLowerCase();
      const global = String(opciones.compression || 'STORE').toUpperCase() === 'DEFLATE' ? 'DEFLATE' : 'STORE';
      const lista = Object.values(this.files);
      if (lista.length > 0xFFFE) throw new ErrorZip('zip64', 'Too many entries for a zip without ZIP64');

      // Cada entrada se prepara (CRC y compresión) en tandas, para no abrir miles de flujos a la vez.
      const preparadas = [];
      for (let i = 0; i < lista.length; i += 16) {
        preparadas.push(...await Promise.all(lista.slice(i, i + 16).map((e) => preparar(e, global))));
      }

      const partes = [];
      const centrales = [];
      let posicion = 0;
      for (const p of preparadas) {
        const nombre = CODIFICADOR.encode(p.nombre);
        const banderas = /^[\x20-\x7e]*$/.test(p.nombre) ? 0 : 0x800;
        const cabecera = new Uint8Array(30 + nombre.length);
        const v = new DataView(cabecera.buffer);
        v.setUint32(0, 0x04034b50, true);
        v.setUint16(4, 20, true);
        v.setUint16(6, banderas, true);
        v.setUint16(8, p.metodo, true);
        v.setUint16(10, p.hora, true);
        v.setUint16(12, p.fecha, true);
        v.setUint32(14, p.crc, true);
        v.setUint32(18, p.datos.length, true);
        v.setUint32(22, p.real, true);
        v.setUint16(26, nombre.length, true);
        cabecera.set(nombre, 30);
        partes.push(cabecera, p.datos);
        centrales.push({ nombre, banderas, p, local: posicion });
        posicion += cabecera.length + p.datos.length;
        if (posicion >= MAXIMO_ZIP) throw new ErrorZip('zip64', 'Zip too large for a zip without ZIP64');
      }

      const inicioDir = posicion;
      let tamanoDir = 0;
      for (const { nombre, banderas, p, local } of centrales) {
        const registro = new Uint8Array(46 + nombre.length);
        const v = new DataView(registro.buffer);
        v.setUint32(0, 0x02014b50, true);
        v.setUint16(4, 20, true);
        v.setUint16(6, 20, true);
        v.setUint16(8, banderas, true);
        v.setUint16(10, p.metodo, true);
        v.setUint16(12, p.hora, true);
        v.setUint16(14, p.fecha, true);
        v.setUint32(16, p.crc, true);
        v.setUint32(20, p.datos.length, true);
        v.setUint32(24, p.real, true);
        v.setUint16(28, nombre.length, true);
        v.setUint32(38, p.dir ? 0x10 : 0, true);
        v.setUint32(42, local, true);
        registro.set(nombre, 46);
        partes.push(registro);
        tamanoDir += registro.length;
      }
      const fin = new Uint8Array(22);
      const vf = new DataView(fin.buffer);
      vf.setUint32(0, 0x06054b50, true);
      vf.setUint16(8, centrales.length, true);
      vf.setUint16(10, centrales.length, true);
      vf.setUint32(12, tamanoDir, true);
      vf.setUint32(16, inicioDir, true);
      partes.push(fin);

      switch (tipo) {
        case 'blob': return new Blob(partes, { type: opciones.mimeType || '' });
        case 'uint8array': case 'nodebuffer': return unir(partes);
        case 'arraybuffer': return unir(partes).buffer;
        case 'base64': return aBase64(unir(partes));
        case 'binarystring': case 'string': {
          const bytes = unir(partes);
          let s = '';
          for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
          return s;
        }
        default: throw new Error('Unsupported type: ' + tipo);
      }
    }
  }

  // Una entrada lista para escribir: { nombre, dir, metodo, crc, datos (ya comprimidos), real, hora, fecha }.
  async function preparar(entrada, global) {
    const { hora, fecha } = entrada._origen ? { hora: entrada._origen.hora, fecha: entrada._origen.fecha } : aDos(entrada.date);
    if (entrada.dir) return { nombre: entrada.name, dir: true, metodo: 0, crc: 0, datos: new Uint8Array(0), real: 0, hora, fecha };
    // Una entrada que vino de un zip y no se tocó se copia tal cual, sin descomprimir ni recomprimir.
    if (entrada._origen) {
      const o = entrada._origen;
      return { nombre: entrada.name, dir: false, metodo: o.metodo, crc: o.crc, datos: bytesCrudos(o), real: o.real, hora, fecha };
    }
    const bytes = await entrada._bytes();
    if (bytes.length >= MAXIMO_ZIP) throw new ErrorZip('zip64', 'Entry too large for a zip without ZIP64');
    const crc = crc32(bytes);
    const pedido = entrada._compresion || global;
    if (pedido === 'DEFLATE' && bytes.length > 0) {
      const comprimido = await deflar(bytes);
      if (comprimido.length < bytes.length) return { nombre: entrada.name, dir: false, metodo: 8, crc, datos: comprimido, real: bytes.length, hora, fecha };
    }
    return { nombre: entrada.name, dir: false, metodo: 0, crc, datos: bytes, real: bytes.length, hora, fecha };
  }

  function aDos(fecha) {
    const f = fecha instanceof Date && !isNaN(fecha) ? fecha : new Date();
    const anio = Math.min(Math.max(f.getFullYear(), 1980), 2107);
    return {
      hora: (f.getHours() << 11) | (f.getMinutes() << 5) | (f.getSeconds() >> 1),
      fecha: ((anio - 1980) << 9) | ((f.getMonth() + 1) << 5) | f.getDate(),
    };
  }
  function fechaDeDos(fecha, hora) {
    return new Date(1980 + (fecha >> 9), Math.max(((fecha >> 5) & 15) - 1, 0), Math.max(fecha & 31, 1), hora >> 11, (hora >> 5) & 63, (hora & 31) * 2);
  }

  ZipPropio.ErrorZip = ErrorZip;
  ZipPropio.crc32 = crc32;
  globalThis.ZipPropio = ZipPropio;
})();

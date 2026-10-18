/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Codificador de códigos QR propio (reemplaza a qrcode.js). Lo usan el
// generador de QR y Enviar archivos con un QR (que muestra un QR con una
// clave), así que no hay red ni dependencias: texto → matriz de módulos.
//
// Qué hace: modo byte (UTF-8), versiones 1 a 40, corrección de errores
// L/M/Q/H con Reed-Solomon sobre GF(256), las 8 máscaras con elección por
// puntaje de penalización, información de formato y de versión.
//
// Dos decisiones para que la salida siga siendo la que daba la librería:
//  - La penalización de máscara es la de qrcode.js (no la de la tabla de la
//    norma), así que se elige la misma máscara que antes.
//  - Si el texto trae caracteres fuera de ASCII se antepone la marca EF BB BF
//    (BOM de UTF-8), igual que hacía la librería: ayuda a los lectores que
//    suponen Latin-1 a reconocer el UTF-8. Se puede apagar con { bom: false }.
//
// API: QRPropio.crear(texto, { nivel: 'L'|'M'|'Q'|'H', version: mínima, bom })
//      devuelve { version, tamano, mascara, esOscuro(fila, col) } o lanza
//      Error si el texto no cabe.
//      QRPropio.pintar(canvas, qr, lado, colores) lo dibuja con la zona de
//      silencio (4 módulos) y deja el canvas de lado + 2 × margen.
//      QRPropio.posicionesAlineacion(versión) da los centros de los patrones
//      de alineación (los usa qr-estilos.js para no taparlos con un logo).
//      QRPropio.estructura(versión, nivel) da { ec, nb, total, largos } (bloques).
(function () {
  'use strict';

  // Bits del nivel de corrección en la información de formato.
  const BITS_NIVEL = { L: 1, M: 0, Q: 3, H: 2 };
  const INDICE_NIVEL = { L: 0, M: 1, Q: 2, H: 3 };

  // Palabras de corrección por bloque y cantidad de bloques, por nivel
  // (L, M, Q, H) y versión (1 a 40, empezando en el índice 0).
  const EC_POR_BLOQUE = [
    [7, 10, 15, 20, 26, 18, 20, 24, 30, 18, 20, 24, 26, 30, 22, 24, 28, 30, 28, 28, 28, 28, 30, 30, 26, 28, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30],
    [10, 16, 26, 18, 24, 16, 18, 22, 22, 26, 30, 22, 22, 24, 24, 28, 28, 26, 26, 26, 26, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28],
    [13, 22, 18, 26, 18, 24, 18, 22, 20, 24, 28, 26, 24, 20, 30, 24, 28, 28, 26, 30, 28, 30, 30, 30, 30, 28, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30],
    [17, 28, 22, 16, 22, 28, 26, 26, 24, 28, 24, 28, 22, 24, 24, 30, 28, 28, 26, 28, 30, 24, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30]
  ];
  const NUM_BLOQUES = [
    [1, 1, 1, 1, 1, 2, 2, 2, 2, 4, 4, 4, 4, 4, 6, 6, 6, 6, 7, 8, 8, 9, 9, 10, 12, 12, 12, 13, 14, 15, 16, 17, 18, 19, 19, 20, 21, 22, 24, 25],
    [1, 1, 1, 2, 2, 4, 4, 4, 5, 5, 5, 8, 9, 9, 10, 10, 11, 13, 14, 16, 17, 17, 18, 20, 21, 23, 25, 26, 28, 29, 31, 33, 35, 37, 38, 40, 43, 45, 47, 49],
    [1, 1, 2, 2, 4, 4, 6, 6, 8, 8, 8, 10, 12, 16, 12, 17, 16, 18, 21, 20, 23, 23, 25, 27, 29, 34, 34, 35, 38, 40, 43, 45, 48, 51, 53, 56, 59, 62, 65, 68],
    [1, 1, 2, 4, 4, 4, 5, 6, 8, 8, 11, 11, 16, 16, 18, 16, 19, 21, 25, 25, 25, 34, 30, 32, 35, 37, 40, 42, 45, 48, 51, 54, 57, 60, 63, 66, 70, 74, 77, 81]
  ];

  // GF(256) con el polinomio 0x11D.
  const EXP = new Uint8Array(512);
  const LOG = new Uint8Array(256);
  (function () {
    let x = 1;
    for (let i = 0; i < 255; i++) {
      EXP[i] = x;
      LOG[x] = i;
      x <<= 1;
      if (x & 0x100) x ^= 0x11D;
    }
    for (let i = 255; i < 512; i++) EXP[i] = EXP[i - 255];
  })();

  function modulosDeDatos(v) {
    let r = (16 * v + 128) * v + 64;
    if (v >= 2) {
      const a = Math.floor(v / 7) + 2;
      r -= (25 * a - 10) * a - 55;
      if (v >= 7) r -= 36;
    }
    return r;
  }
  // Palabras de código de datos que caben en la versión v con el nivel dado.
  function capacidadDatos(v, nivel) {
    const i = INDICE_NIVEL[nivel];
    return Math.floor(modulosDeDatos(v) / 8) - EC_POR_BLOQUE[i][v - 1] * NUM_BLOQUES[i][v - 1];
  }

  // Texto → bytes UTF-8 (con la marca BOM si hay algo fuera de ASCII).
  function aBytes(texto, bom) {
    const b = new TextEncoder().encode(texto);
    if (bom && b.some((x) => x > 127)) {
      const c = new Uint8Array(b.length + 3);
      c.set([0xEF, 0xBB, 0xBF]);
      c.set(b, 3);
      return c;
    }
    return b;
  }

  // Flujo de bits de datos: modo, largo, bytes, terminador y relleno.
  function armarDatos(bytes, v, nivel) {
    const cap = capacidadDatos(v, nivel) * 8;
    const bits = [];
    const poner = (valor, n) => { for (let i = n - 1; i >= 0; i--) bits.push((valor >>> i) & 1); };
    poner(4, 4); // modo byte
    poner(bytes.length, v < 10 ? 8 : 16);
    for (const x of bytes) poner(x, 8);
    if (bits.length > cap) throw new Error('Too long data');
    if (bits.length + 4 <= cap) poner(0, 4);
    while (bits.length % 8) bits.push(0);
    const palabras = [];
    for (let i = 0; i < bits.length; i += 8) {
      let w = 0;
      for (let j = 0; j < 8; j++) w = (w << 1) | bits[i + j];
      palabras.push(w);
    }
    for (let relleno = 0xEC; palabras.length < cap / 8; relleno ^= 0xEC ^ 0x11) palabras.push(relleno);
    return palabras;
  }

  // Resto de dividir los datos por el generador de Reed-Solomon de grado n.
  function restoRS(datos, n) {
    let gen = [1];
    for (let i = 0; i < n; i++) {
      const sig = new Array(gen.length + 1).fill(0);
      for (let j = 0; j < gen.length; j++) {
        sig[j] ^= gen[j];
        sig[j + 1] ^= gen[j] ? EXP[LOG[gen[j]] + i] : 0;
      }
      gen = sig;
    }
    const resto = new Array(n).fill(0);
    for (const d of datos) {
      const f = d ^ resto.shift();
      resto.push(0);
      if (f) for (let j = 0; j < n; j++) if (gen[j + 1]) resto[j] ^= EXP[LOG[gen[j + 1]] + LOG[f]];
    }
    return resto;
  }

  // Parte los datos en bloques, calcula la corrección de cada uno y los intercala.
  function intercalar(datos, v, nivel) {
    const i = INDICE_NIVEL[nivel];
    const ec = EC_POR_BLOQUE[i][v - 1];
    const nb = NUM_BLOQUES[i][v - 1];
    const total = Math.floor(modulosDeDatos(v) / 8);
    const cortos = nb - (total % nb);
    const largoCorto = Math.floor(total / nb) - ec;
    const bloques = [];
    const correccion = [];
    let pos = 0;
    for (let b = 0; b < nb; b++) {
      const n = largoCorto + (b < cortos ? 0 : 1);
      const bloque = datos.slice(pos, pos + n);
      pos += n;
      bloques.push(bloque);
      correccion.push(restoRS(bloque, ec));
    }
    const salida = [];
    for (let k = 0; k <= largoCorto; k++) for (const bl of bloques) if (k < bl.length) salida.push(bl[k]);
    for (let k = 0; k < ec; k++) for (const c of correccion) salida.push(c[k]);
    return salida;
  }

  function digitos(x) { let n = 0; while (x) { n++; x >>>= 1; } return n; }
  // Código BCH: resto de valor·2^grado dividido por el generador g.
  function bch(valor, g, grado) {
    let r = valor << grado;
    while (digitos(r) - digitos(g) >= 0) r ^= g << (digitos(r) - digitos(g));
    return r;
  }

  function posicionesAlineacion(v) {
    if (v === 1) return [];
    const n = 4 * v + 17;
    const cuenta = Math.floor(v / 7) + 2;
    const paso = v === 32 ? 26 : Math.ceil((n - 13) / (cuenta * 2 - 2)) * 2;
    const p = [6];
    for (let k = 0; k < cuenta - 1; k++) p.splice(1, 0, n - 7 - k * paso);
    return p;
  }

  const MASCARAS = [
    (f, c) => (f + c) % 2 === 0,
    (f) => f % 2 === 0,
    (f, c) => c % 3 === 0,
    (f, c) => (f + c) % 3 === 0,
    (f, c) => (Math.floor(f / 2) + Math.floor(c / 3)) % 2 === 0,
    (f, c) => (f * c) % 2 + (f * c) % 3 === 0,
    (f, c) => ((f * c) % 2 + (f * c) % 3) % 2 === 0,
    (f, c) => ((f * c) % 3 + (f + c) % 2) % 2 === 0
  ];

  // Arma la matriz con una máscara. prueba = true deja en claro la información
  // de formato y de versión (así se puntúa la máscara, como hacía la librería).
  function armarMatriz(v, nivel, datos, mascara, prueba) {
    const n = 4 * v + 17;
    const m = [];
    for (let f = 0; f < n; f++) m.push(new Array(n).fill(null));

    // Cuadros de posición (con su borde claro) en tres esquinas.
    for (const [f0, c0] of [[0, 0], [n - 7, 0], [0, n - 7]]) {
      for (let f = -1; f <= 7; f++) {
        if (f0 + f <= -1 || f0 + f >= n) continue;
        for (let c = -1; c <= 7; c++) {
          if (c0 + c <= -1 || c0 + c >= n) continue;
          m[f0 + f][c0 + c] = (f >= 0 && f <= 6 && (c === 0 || c === 6)) ||
            (c >= 0 && c <= 6 && (f === 0 || f === 6)) ||
            (f >= 2 && f <= 4 && c >= 2 && c <= 4);
        }
      }
    }
    // Patrones de alineación (no pisan los cuadros de posición).
    const pos = posicionesAlineacion(v);
    for (const f of pos) {
      for (const c of pos) {
        if (m[f][c] !== null) continue;
        for (let a = -2; a <= 2; a++) {
          for (let b = -2; b <= 2; b++) {
            m[f + a][c + b] = a === -2 || a === 2 || b === -2 || b === 2 || (a === 0 && b === 0);
          }
        }
      }
    }
    // Líneas de temporización.
    for (let i = 8; i < n - 8; i++) {
      if (m[i][6] === null) m[i][6] = i % 2 === 0;
      if (m[6][i] === null) m[6][i] = i % 2 === 0;
    }
    // Información de formato (nivel + máscara) y módulo oscuro fijo.
    const base = (BITS_NIVEL[nivel] << 3) | mascara;
    const formato = ((base << 10) | bch(base, 0x537, 10)) ^ 0x5412;
    for (let i = 0; i < 15; i++) {
      const bit = !prueba && ((formato >> i) & 1) === 1;
      if (i < 6) m[i][8] = bit; else if (i < 8) m[i + 1][8] = bit; else m[n - 15 + i][8] = bit;
      if (i < 8) m[8][n - i - 1] = bit; else if (i < 9) m[8][15 - i] = bit; else m[8][15 - i - 1] = bit;
    }
    m[n - 8][8] = !prueba;
    // Información de versión (desde la 7).
    if (v >= 7) {
      const ver = (v << 12) | bch(v, 0x1F25, 12);
      for (let i = 0; i < 18; i++) {
        const bit = !prueba && ((ver >> i) & 1) === 1;
        m[Math.floor(i / 3)][i % 3 + n - 11] = bit;
        m[i % 3 + n - 11][Math.floor(i / 3)] = bit;
      }
    }
    // Datos en zigzag, de abajo hacia arriba, de dos columnas en dos.
    let dir = -1, f = n - 1, bit = 7, byte = 0;
    for (let c = n - 1; c > 0; c -= 2) {
      if (c === 6) c--;
      for (;;) {
        for (let k = 0; k < 2; k++) {
          if (m[f][c - k] === null) {
            let oscuro = byte < datos.length && ((datos[byte] >>> bit) & 1) === 1;
            if (MASCARAS[mascara](f, c - k)) oscuro = !oscuro;
            m[f][c - k] = oscuro;
            if (--bit === -1) { byte++; bit = 7; }
          }
        }
        f += dir;
        if (f < 0 || f >= n) { f -= dir; dir = -dir; break; }
      }
    }
    return m;
  }

  // Puntaje de penalización de qrcode.js (menor es mejor).
  function penalizacion(m) {
    const n = m.length;
    let p = 0;
    for (let f = 0; f < n; f++) {
      for (let c = 0; c < n; c++) {
        let iguales = 0;
        const o = m[f][c];
        for (let a = -1; a <= 1; a++) {
          if (f + a < 0 || f + a >= n) continue;
          for (let b = -1; b <= 1; b++) {
            if (c + b < 0 || c + b >= n) continue;
            if ((a !== 0 || b !== 0) && o === m[f + a][c + b]) iguales++;
          }
        }
        if (iguales > 5) p += 3 + iguales - 5;
      }
    }
    for (let f = 0; f < n - 1; f++) {
      for (let c = 0; c < n - 1; c++) {
        const s = (m[f][c] ? 1 : 0) + (m[f + 1][c] ? 1 : 0) + (m[f][c + 1] ? 1 : 0) + (m[f + 1][c + 1] ? 1 : 0);
        if (s === 0 || s === 4) p += 3;
      }
    }
    for (let f = 0; f < n; f++) {
      for (let c = 0; c < n - 6; c++) {
        if (m[f][c] && !m[f][c + 1] && m[f][c + 2] && m[f][c + 3] && m[f][c + 4] && !m[f][c + 5] && m[f][c + 6]) p += 40;
      }
    }
    for (let c = 0; c < n; c++) {
      for (let f = 0; f < n - 6; f++) {
        if (m[f][c] && !m[f + 1][c] && m[f + 2][c] && m[f + 3][c] && m[f + 4][c] && !m[f + 5][c] && m[f + 6][c]) p += 40;
      }
    }
    let oscuros = 0;
    for (let f = 0; f < n; f++) for (let c = 0; c < n; c++) if (m[f][c]) oscuros++;
    p += 10 * (Math.abs(100 * oscuros / n / n - 50) / 5);
    return p;
  }

  function crear(texto, opciones) {
    const op = opciones || {};
    const nivel = op.nivel || 'M';
    if (!(nivel in BITS_NIVEL)) throw new Error('Nivel de corrección desconocido: ' + nivel);
    const bytes = aBytes(String(texto), op.bom !== false);
    // Versión más chica donde caben el encabezado y los bytes.
    let v = Math.max(1, op.version || 1);
    for (; v <= 40; v++) {
      if (4 + (v < 10 ? 8 : 16) + bytes.length * 8 <= capacidadDatos(v, nivel) * 8) break;
    }
    if (v > 40) throw new Error('Too long data');
    const datos = intercalar(armarDatos(bytes, v, nivel), v, nivel);
    let mejor = 0, menor = 0;
    for (let k = 0; k < 8; k++) {
      const p = penalizacion(armarMatriz(v, nivel, datos, k, true));
      if (k === 0 || p < menor) { menor = p; mejor = k; }
    }
    const m = armarMatriz(v, nivel, datos, mejor, false);
    return {
      version: v,
      tamano: m.length,
      mascara: mejor,
      esOscuro: (f, c) => m[f][c] === true
    };
  }

  // Dibuja el QR en un canvas, con 4 módulos de zona de silencio. "lado" es el
  // tamaño del código sin margen; el canvas queda de lado + 2 × margen.
  function pintar(canvas, qr, lado, colores) {
    const col = colores || {};
    const margen = Math.ceil(4 * lado / qr.tamano);
    canvas.width = canvas.height = lado + margen * 2;
    const cx = canvas.getContext('2d');
    cx.fillStyle = col.claro || '#ffffff';
    cx.fillRect(0, 0, canvas.width, canvas.height);
    cx.fillStyle = col.oscuro || '#000000';
    const celda = lado / qr.tamano;
    // Los bordes se redondean para que los módulos queden nítidos, sin huecos.
    for (let f = 0; f < qr.tamano; f++) {
      const y0 = Math.round(f * celda), y1 = Math.round((f + 1) * celda);
      for (let c = 0; c < qr.tamano; c++) {
        if (!qr.esOscuro(f, c)) continue;
        const x0 = Math.round(c * celda), x1 = Math.round((c + 1) * celda);
        cx.fillRect(margen + x0, margen + y0, x1 - x0, y1 - y0);
      }
    }
    return canvas;
  }

  // Cómo reparte la norma las palabras de código de la versión v con el nivel dado: cuántas de
  // corrección por bloque y el largo de datos de cada bloque (solo lectura; lo usa qr-estilos.js
  // para saber cuánto daño aguanta cada bloque).
  function estructura(v, nivel) {
    const i = INDICE_NIVEL[nivel];
    const ec = EC_POR_BLOQUE[i][v - 1];
    const nb = NUM_BLOQUES[i][v - 1];
    const total = Math.floor(modulosDeDatos(v) / 8);
    const cortos = nb - (total % nb);
    const largoCorto = Math.floor(total / nb) - ec;
    const largos = [];
    for (let b = 0; b < nb; b++) largos.push(largoCorto + (b < cortos ? 0 : 1));
    return { ec, nb, total, largos };
  }

  globalThis.QRPropio = { crear, pintar, posicionesAlineacion, estructura };
})();

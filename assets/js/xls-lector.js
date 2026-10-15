/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Lector de libros de Excel 97-2003 (.xls) del motor propio de Excel (etapa 9). Un .xls es un contenedor CFB ("Compound File", el de Office antiguo)
// con un flujo "Workbook" de registros BIFF8 ([MS-CFB], [MS-XLS]); también se leen los de Excel 5/95 (BIFF5, flujo "Book"). Requiere xlsx-modelo.js
// (errores) y xlsx-estilos.js (formatos integrados); no depende del navegador.
//
//   const libro = XlsLector.leer(bytes);
//   libro = { hojas: [{ nombre, filas: [[valor…]…], anchos: [caracteres…], combinadas: ['A1:B2'…] }], fecha1904 }   (valor: ver OdsLector)
//
// Las fórmulas no se leen (queda su último valor, que el archivo guarda). Con contraseña (FILEPASS) se rechaza. Se omiten las hojas de gráfico y de macros.
(function () {
  const MAX_FILAS = 65536, MAX_COLS = 256, MAX_TEXTO = 32767;   // los límites de un .xls
  const MAX_CELDAS = 10000000;   // en todo el libro: crearLibro no arma un .xlsx de más de ~8 M celdas, y antes de llegar ahí la memoria ya se fue
  const ERRORES = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0F: '#VALUE!', 0x17: '#REF!', 0x1D: '#NAME?', 0x24: '#NUM!', 0x2A: '#N/A' };
  const danado = (detalle) => new LibroXlsx.ErrorXlsx('danado', detalle);
  const u16 = (b, p) => b[p] | (b[p + 1] << 8);
  const u32 = (b, p) => (b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24)) >>> 0;
  const FIN = 0xFFFFFFFE, LIBRE = 0xFFFFFFFF;

  // ---- contenedor CFB -----------------------------------------------------------------------------------------------
  function leerCfb(b) {
    if (b.length < 512) throw danado('contenedor muy corto');
    const desp = u16(b, 0x1E), despMini = u16(b, 0x20);
    if (desp < 7 || desp > 16 || despMini > desp) throw danado('tamaño de sector');
    const tam = 1 << desp, tamMini = 1 << despMini;
    const nSect = Math.floor((b.length - 512) / tam) + 1;
    const sector = (id) => {
      const ini = (id + 1) * tam;
      if (id >= nSect + 1 || ini >= b.length) throw danado('sector fuera del archivo');
      return b.subarray(ini, Math.min(ini + tam, b.length));
    };
    // tabla de asignación (FAT): sus sectores los dice el encabezado y, si son más de 109, la cadena DIFAT. No puede haber más sectores de FAT que
    // sectores en el archivo: una DIFAT hostil que repite el mismo id armaba una FAT de gigas con un archivo de 10 MB.
    const idsFat = [];
    for (let k = 0; k < 109; k++) { const id = u32(b, 0x4C + k * 4); if (id < 0xFFFFFFFA) idsFat.push(id); }
    let dif = u32(b, 0x44), vueltas = 0;
    while (dif < 0xFFFFFFFA && vueltas++ < nSect && idsFat.length <= nSect) {
      const s = sector(dif);
      for (let k = 0; k < tam / 4 - 1; k++) { const id = u32(s, k * 4); if (id < 0xFFFFFFFA) idsFat.push(id); }
      dif = u32(s, tam - 4);
    }
    if (idsFat.length > nSect) throw danado('FAT más grande que el archivo');
    const fat = new Uint32Array(idsFat.length * (tam / 4));
    idsFat.forEach((id, k) => { const s = sector(id); for (let j = 0; j < tam / 4; j++) fat[k * (tam / 4) + j] = u32(s, j * 4); });
    // una cadena no pasa por más sectores que los que tiene el archivo (con un ciclo, la tabla entera: un flujo de cientos de megas)
    const cadena = (inicio, tabla, max) => {
      const ids = [];
      let id = inicio;
      while (id < 0xFFFFFFFA) {
        if (ids.length >= max || id >= tabla.length) throw danado('cadena de sectores');
        ids.push(id); id = tabla[id];
      }
      return ids;
    };
    const leerCadena = (inicio, tamano) => {
      const ids = cadena(inicio, fat, nSect), r = new Uint8Array(ids.length * tam);
      ids.forEach((id, k) => r.set(sector(id), k * tam));
      return tamano === undefined ? r : r.subarray(0, tamano);
    };
    // directorio
    const dir = leerCadena(u32(b, 0x30));
    const entradas = [];
    for (let p = 0; p + 128 <= dir.length; p += 128) {
      const nc = u16(dir, p + 64);
      if (!dir[p + 66]) continue;
      entradas.push({ nombre: new TextDecoder('utf-16le').decode(dir.subarray(p, p + Math.max(0, nc - 2))), tipo: dir[p + 66], inicio: u32(dir, p + 116), tam: u32(dir, p + 120) });
    }
    const raiz = entradas.find((e) => e.tipo === 5);
    const corte = u32(b, 0x38);
    let miniFat = null, miniFlujo = null;
    const flujo = (e) => {
      if (e.tam < corte && raiz) {
        if (!miniFat) {
          const bytes = leerCadena(u32(b, 0x3C));
          miniFat = new Uint32Array(bytes.length / 4);
          for (let k = 0; k < miniFat.length; k++) miniFat[k] = u32(bytes, k * 4);
          miniFlujo = leerCadena(raiz.inicio, raiz.tam);
        }
        const ids = cadena(e.inicio, miniFat, Math.ceil(miniFlujo.length / tamMini)), r = new Uint8Array(ids.length * tamMini);
        ids.forEach((id, k) => { const ini = id * tamMini; if (ini + tamMini > miniFlujo.length + tamMini) throw danado('minisector'); r.set(miniFlujo.subarray(ini, ini + tamMini), k * tamMini); });
        return r.subarray(0, e.tam);
      }
      return leerCadena(e.inicio, e.tam);
    };
    return { entradas, flujo };
  }

  // ---- registros BIFF -----------------------------------------------------------------------------------------------
  // Un registro lleva hasta 8224 bytes; si hay más, siguen registros CONTINUE (0x3C). Devuelve { id, d (datos juntos), partes (cada pedazo) }.
  function* registros(b, desde) {
    let p = desde || 0;
    const n = b.length;
    let actual = null;
    while (p + 4 <= n) {
      const id = u16(b, p), tam = u16(b, p + 2);
      if (p + 4 + tam > n) break;
      const datos = b.subarray(p + 4, p + 4 + tam);
      p += 4 + tam;
      if (id === 0x3C && actual) { actual.partes.push(datos); continue; }
      if (actual) { yield juntar(actual); }
      actual = { id, partes: [datos] };
    }
    if (actual) yield juntar(actual);
  }
  function juntar(r) {
    if (r.partes.length === 1) r.d = r.partes[0];
    else { const t = r.partes.reduce((s, x) => s + x.length, 0); r.d = new Uint8Array(t); let q = 0; for (const x of r.partes) { r.d.set(x, q); q += x.length; } }
    return r;
  }

  const utf16 = new TextDecoder('utf-16le');
  // BIFF5 (Excel 5/95) guarda los textos en la página de códigos del libro (registro CODEPAGE; casi siempre 1252: "€", comillas tipográficas).
  // BIFF8 no: sus textos "comprimidos" son UTF-16 sin el byte alto (latin1).
  const PAGINAS = { 874: 'windows-874', 932: 'shift_jis', 936: 'gbk', 949: 'euc-kr', 950: 'big5', 1250: 'windows-1250', 1251: 'windows-1251', 1252: 'windows-1252', 1253: 'windows-1253', 1254: 'windows-1254', 1255: 'windows-1255', 1256: 'windows-1256', 1257: 'windows-1257', 1258: 'windows-1258', 10000: 'macintosh', 32768: 'macintosh' };
  const decodificador = (pagina) => { try { return new TextDecoder(PAGINAS[pagina] || 'windows-1252'); } catch (e) { return new TextDecoder('windows-1252'); } };
  const vista = new DataView(new ArrayBuffer(8));
  function doble(b, p) { vista.setUint32(0, u32(b, p), true); vista.setUint32(4, u32(b, p + 4), true); return vista.getFloat64(0, true); }
  function numeroRk(v) {
    let n;
    if (v & 2) n = v >> 2;
    else { vista.setUint32(0, 0, true); vista.setUint32(4, (v & 0xFFFFFFFC) >>> 0, true); n = vista.getFloat64(0, true); }
    return v & 1 ? n / 100 : n;
  }

  // XLUnicodeString dentro de un solo pedazo: cch (2 bytes, o 1 en las cadenas cortas), banderas, [ejecuciones], [extensión], caracteres
  function cadenaBiff8(b, p, cchBytes) {
    const cch = cchBytes === 1 ? b[p] : u16(b, p);
    const banderas = b[p + cchBytes];
    let q = p + cchBytes + 1;
    let runs = 0, ext = 0;
    if (banderas & 8) { runs = u16(b, q); q += 2; }
    if (banderas & 4) { ext = u32(b, q); q += 4; }
    const bytes = (banderas & 1) ? cch * 2 : cch;
    const crudo = b.subarray(q, q + bytes);
    return { texto: (banderas & 1) ? utf16.decode(crudo) : latin1(crudo), fin: q + bytes + runs * 4 + ext };
  }
  const latin1 = (crudo) => { let s = ''; for (let k = 0; k < crudo.length; k += 8192) s += String.fromCharCode.apply(null, crudo.subarray(k, k + 8192)); return s; };

  // Tabla de textos compartidos (SST): una cadena puede cortarse en un CONTINUE, y entonces el pedazo siguiente empieza con una bandera nueva
  // (¿los caracteres que siguen son de 1 o de 2 bytes?).
  function leerSst(partes) {
    let ci = 0, d = partes[0], p = 8;   // 8: cstTotal y cstUnique
    const total = u32(d, 4);
    const textos = [];
    const siguiente = () => { ci++; d = partes[ci]; p = 0; return d !== undefined; };
    const byte = () => { if (p >= d.length && !siguiente()) throw danado('SST cortada'); return d[p++]; };
    const salta = (n) => { while (n > 0) { if (p >= d.length && !siguiente()) return; const t = Math.min(n, d.length - p); p += t; n -= t; } };
    while (textos.length < total && (p < d.length || ci + 1 < partes.length)) {
      if (p >= d.length && !siguiente()) break;
      const cch = byte() | (byte() << 8);
      const banderas = byte();
      let runs = 0, ext = 0;
      if (banderas & 8) runs = byte() | (byte() << 8);
      if (banderas & 4) ext = (byte() | (byte() << 8) | (byte() << 16) | (byte() << 24)) >>> 0;
      let alto = (banderas & 1) === 1, restantes = cch, s = '';
      while (restantes > 0) {
        if (p >= d.length) { if (!siguiente()) break; alto = (d[p++] & 1) === 1; }   // el pedazo nuevo trae su bandera
        const caben = Math.min(restantes, Math.floor((d.length - p) / (alto ? 2 : 1)));
        if (caben <= 0) { if (!siguiente()) break; alto = (d[p++] & 1) === 1; continue; }
        const crudo = d.subarray(p, p + caben * (alto ? 2 : 1));
        s += alto ? utf16.decode(crudo) : latin1(crudo);
        p += crudo.length; restantes -= caben;
      }
      salta(runs * 4 + ext);
      textos.push(s.length > MAX_TEXTO ? s.slice(0, MAX_TEXTO) : s);
    }
    return textos;
  }

  function leer(bytes) {
    const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const cfb = leerCfb(b);
    const entrada = cfb.entradas.find((e) => e.tipo === 2 && (e.nombre === 'Workbook' || e.nombre === 'Book')) || cfb.entradas.find((e) => e.tipo === 2 && /^(Workbook|Book)$/i.test(e.nombre));
    if (!entrada) {
      if (cfb.entradas.some((e) => e.nombre === 'EncryptedPackage')) throw new LibroXlsx.ErrorXlsx('cifrado');
      throw new LibroXlsx.ErrorXlsx('sinLibro');
    }
    const wb = cfb.flujo(entrada);
    // ---- globales ----
    let version = 0x0600, fecha1904 = false, fin = false;
    const hojasDef = [], textos = [], formatos = new Map(), xfFormato = [];
    let primero = true, profundidad = 0, pagina = decodificador(1252);
    const texto5 = (crudo) => pagina.decode(crudo);   // un texto de BIFF5
    for (const r of registros(wb)) {
      const d = r.d;
      if (r.id === 0x0809) {
        if (primero) { version = u16(d, 0); primero = false; }
        profundidad++;
      } else if (r.id === 0x000A) { if (--profundidad <= 0) { fin = true; break; } }
      else if (profundidad !== 1) continue;
      else if (r.id === 0x002F) throw new LibroXlsx.ErrorXlsx('cifrado');                       // FILEPASS
      else if (r.id === 0x0022) fecha1904 = u16(d, 0) === 1;                                    // DATEMODE
      else if (r.id === 0x0042) pagina = decodificador(u16(d, 0));                              // CODEPAGE
      else if (r.id === 0x0085) {                                                              // BOUNDSHEET: posición, estado, tipo, nombre
        const nombre = version >= 0x0600 ? cadenaBiff8(d, 6, 1).texto : texto5(d.subarray(7, 7 + d[6]));
        hojasDef.push({ pos: u32(d, 0), tipo: d[5], nombre, estado: d[4] === 1 ? 'oculta' : d[4] === 2 ? 'muyOculta' : 'visible' });
      } else if (r.id === 0x00FC) { for (const t of leerSst(r.partes)) textos.push(t); }       // SST
      else if (r.id === 0x041E) {                                                              // FORMAT: id y código
        formatos.set(u16(d, 0), version >= 0x0600 ? cadenaBiff8(d, 2, 2).texto : texto5(d.subarray(3, 3 + d[2])));
      } else if (r.id === 0x00E0) xfFormato.push(u16(d, 2));                                   // XF: fuente, formato…
    }
    if (!fin && !hojasDef.length) throw danado('libro sin hojas');
    // un formato integrado sale como su número (crearLibro lo escribe tal cual y Excel lo muestra según el idioma de quien abre el libro)
    const codigoDe = (xf) => {
      const id = xfFormato[xf];
      if (id === undefined || id === 0) return null;
      if (formatos.has(id)) return formatos.get(id) === 'General' ? null : formatos.get(id);
      return XlsxEstilos.FORMATOS_INTEGRADOS[id] !== undefined && id < 164 ? id : null;
    };

    const hojas = [];
    // Los subflujos de las hojas no se pisan: entre todas no recorren más que el flujo entero. Un libro hostil con cientos de hojas que apuntan al mismo
    // subflujo (o a uno sin EOF) lo leería una vez por hoja: 400 KB armaban gigas. Se deja margen para un EOF que falte en un libro real.
    let presupuesto = 4 * wb.length + 65536, celdas = 0;
    for (const def of hojasDef) {
      if (def.tipo !== 0) continue;   // 0: hoja de cálculo; 2: gráfico; 1 y 6: macros
      const filas = [], anchos = [], combinadas = [], columnasOcultas = new Set(), filasOcultas = [];
      let profund = 0, ultimaFormula = null;
      const poner = (f, c, v) => {
        if (f >= MAX_FILAS || c >= MAX_COLS) return;
        if (++celdas > MAX_CELDAS) throw new LibroXlsx.ErrorXlsx('grande', 'más de ' + MAX_CELDAS + ' celdas');
        (filas[f] || (filas[f] = []))[c] = v;
      };
      const numero = (f, c, xf, n) => { if (!Number.isFinite(n)) return; const cod = codigoDe(xf); poner(f, c, cod ? { valor: n, formato: cod } : n); };
      for (const r of registros(wb, def.pos)) {
        const d = r.d;
        if ((presupuesto -= 4 + d.length) < 0) throw danado('hojas que comparten su subflujo');
        if (r.id === 0x0809) { profund++; continue; }
        if (r.id === 0x000A) { if (--profund <= 0) break; continue; }
        if (profund !== 1) continue;
        switch (r.id) {
          case 0x0208:                                                                                 // ROW: la bandera 0x20 de grbit (fDyZero) es "oculta"
            if (d.length >= 14 && (u16(d, 12) & 0x20)) filasOcultas.push(u16(d, 0));
            break;
          case 0x0203: numero(u16(d, 0), u16(d, 2), u16(d, 4), doble(d, 6)); break;                    // NUMBER
          case 0x027E: numero(u16(d, 0), u16(d, 2), u16(d, 4), numeroRk(u32(d, 6))); break;             // RK
          case 0x00BD: {                                                                               // MULRK
            const f = u16(d, 0), c0 = u16(d, 2), c1 = u16(d, d.length - 2);
            for (let c = c0, q = 4; c <= c1 && q + 6 <= d.length - 2; c++, q += 6) numero(f, c, u16(d, q), numeroRk(u32(d, q + 2)));
            break;
          }
          case 0x00FD: { const t = textos[u32(d, 6)]; if (t) poner(u16(d, 0), u16(d, 2), t); break; }    // LABELSST
          case 0x0204: case 0x00D6: {                                                                  // LABEL y RSTRING (texto con formato) de BIFF5, y algún BIFF8 antiguo
            const t = version >= 0x0600 ? cadenaBiff8(d, 6, 2).texto : texto5(d.subarray(8, 8 + u16(d, 6)));
            if (t) poner(u16(d, 0), u16(d, 2), t.length > MAX_TEXTO ? t.slice(0, MAX_TEXTO) : t);
            break;
          }
          case 0x0205:                                                                                 // BOOLERR
            poner(u16(d, 0), u16(d, 2), d[7] ? { error: ERRORES[d[6]] || '#VALUE!' } : d[6] === 1);
            break;
          case 0x0006: {                                                                               // FORMULA: se guarda el último resultado
            const f = u16(d, 0), c = u16(d, 2);
            if (d[12] === 0xFF && d[13] === 0xFF) {
              if (d[6] === 0) ultimaFormula = { f, c };                                                // el texto viene en el STRING que sigue
              else if (d[6] === 1) poner(f, c, d[8] === 1);
              else if (d[6] === 2) poner(f, c, { error: ERRORES[d[8]] || '#VALUE!' });
            } else numero(f, c, u16(d, 4), doble(d, 6));
            break;
          }
          case 0x0207:                                                                                 // STRING (resultado de una fórmula)
            if (ultimaFormula) {
              const t = version >= 0x0600 ? cadenaBiff8(d, 0, 2).texto : texto5(d.subarray(2, 2 + u16(d, 0)));
              if (t) poner(ultimaFormula.f, ultimaFormula.c, t.length > MAX_TEXTO ? t.slice(0, MAX_TEXTO) : t);
              ultimaFormula = null;
            }
            break;
          case 0x00E5: {                                                                               // MERGEDCELLS
            const n = u16(d, 0);
            for (let k = 0, q = 2; k < n && q + 8 <= d.length; k++, q += 8) {
              const f0 = u16(d, q), f1 = u16(d, q + 2), c0 = u16(d, q + 4), c1 = u16(d, q + 6);
              if ((f1 > f0 || c1 > c0) && f1 >= f0 && c1 >= c0) combinadas.push(LibroXlsx.numeroACol(c0) + (f0 + 1) + ':' + LibroXlsx.numeroACol(c1) + (f1 + 1));
            }
            break;
          }
          case 0x007D: {                                                                               // COLINFO: ancho en 1/256 de carácter, oculta en el bit 0
            const c0 = u16(d, 0), c1 = Math.min(u16(d, 2), MAX_COLS - 1), w = u16(d, 4) / 256, oculta = (u16(d, 8) & 1) === 1;
            for (let c = c0; c <= c1; c++) { anchos[c] = w; if (oculta) columnasOcultas.add(c); }
            break;
          }
          default: break;
        }
      }
      let ancho = 0;
      for (let f = 0; f < filas.length; f++) { if (!filas[f]) filas[f] = []; if (filas[f].length > ancho) ancho = filas[f].length; }
      hojas.push({ nombre: def.nombre, filas, anchos: Array.from(anchos.slice(0, ancho), (w) => w || 0), combinadas, estado: def.estado, columnasOcultas: [...columnasOcultas].filter((c) => c < ancho).sort((a, b) => a - b), filasOcultas: filasOcultas.filter((r) => r < filas.length) });
    }
    if (!hojas.length) throw new LibroXlsx.ErrorXlsx('sinLibro', 'sin hojas');
    return { hojas, fecha1904 };
  }

  globalThis.XlsLector = { leer, _leerCfb: leerCfb, _registros: registros, _leerSst: leerSst };
})();

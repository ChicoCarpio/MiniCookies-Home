/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Lector de libros de Excel 97-2003 (.xls) del motor propio de Excel (etapa 9). Un .xls es un contenedor CFB ("Compound File", el de Office antiguo)
// con un flujo "Workbook" de registros BIFF8 ([MS-CFB], [MS-XLS]); también se leen los de Excel 5/95 (BIFF5, flujo "Book"). Requiere cfb-propio.js
// (el contenedor), xlsx-modelo.js (errores) y xlsx-estilos.js (formatos integrados); no depende del navegador.
//
//   const libro = XlsLector.leer(bytes);
//   libro = { hojas: [{ nombre, filas: [[valor…]…], anchos: [caracteres…], combinadas: ['A1:B2'…], altosFila }], fecha1904 }   (valor: ver OdsLector)
//
// Estilos de celda (solo BIFF8): fuente, color, relleno, bordes y alineación salen como { valor, formato?, estilo } (el contrato de XlsxEscritor.crearLibro;
// un objeto de estilo por XF, compartido por todas sus celdas), una celda vacía con relleno, borde o alineación como { valor: null, estilo }, y el alto de
// las filas con alto propio en altosFila ({ fila desde 0: puntos }). BIFF5 queda solo con valores.
// Las fórmulas no se leen (queda su último valor, que el archivo guarda). Con contraseña (FILEPASS) se rechaza. Se omiten las hojas de gráfico y de macros.
(function () {
  const MAX_FILAS = 65536, MAX_COLS = 256, MAX_TEXTO = 32767;   // los límites de un .xls
  const MAX_CELDAS = 10000000;   // en todo el libro: crearLibro no arma un .xlsx de más de ~8 M celdas, y antes de llegar ahí la memoria ya se fue
  const ERRORES = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0F: '#VALUE!', 0x17: '#REF!', 0x1D: '#NAME?', 0x24: '#NUM!', 0x2A: '#N/A' };
  const danado = (detalle) => new LibroXlsx.ErrorXlsx('danado', detalle);
  const u16 = (b, p) => b[p] | (b[p + 1] << 8);
  const u32 = (b, p) => (b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24)) >>> 0;

  // ---- contenedor CFB (cfb-propio.js, compartido con los lectores de .doc y .ppt) ----------------------------------------
  const leerCfb = (b) => CfbPropio.leer(b, { danado });

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

  // ---- estilos (BIFF8) ----------------------------------------------------------------------------------------------
  const MAX_FUENTES = 1024, MAX_XF = 65536, MAX_VACIAS = 2000000;   // lo que se acepta de un archivo hostil
  const BORDES = [null, 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot'];
  const TRAMAS = [null, 'solid', 'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis', 'gray125', 'gray0625'];
  const ALIN_H = [null, 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'], ALIN_V = ['top', 'center', null, 'justify', 'distributed'];
  // tipos de ExtProp de un XFEXT que son un color: relleno, trama, borde de arriba, abajo, izquierda, derecha, diagonal y texto
  const TIPOS_EXT = { 4: 'fg', 5: 'bg', 7: 'arr', 8: 'aba', 9: 'izq', 10: 'der', 11: 'diag', 13: 'texto' };
  // FONT (0x0031): alto en 1/20 de punto, banderas (cursiva, tachado), color, peso, subrayado y nombre
  function leerFuente(d) {
    if (d.length < 16) return null;
    const alto = u16(d, 0), banderas = u16(d, 2), icv = u16(d, 4), peso = u16(d, 6), sub = d[10];
    const cch = d[14], alta = (d[15] & 1) === 1;
    const nombre = d.length >= 16 + cch * (alta ? 2 : 1) ? (alta ? utf16.decode(d.subarray(16, 16 + cch * 2)) : latin1(d.subarray(16, 16 + cch))) : '';
    return { alto: alto / 20, cursiva: (banderas & 2) !== 0, tachado: (banderas & 8) !== 0, icv, negrita: peso >= 700, sub, nombre };
  }
  // XF (0x00E0) de BIFF8: formato de número, alineación, bordes y relleno
  function leerXf(d) {
    if (d.length < 20) return null;
    const bor = u32(d, 10), col = u32(d, 14);
    return { fuente: u16(d, 0), estiloXf: (u16(d, 4) & 4) !== 0, h: d[6] & 7, ajustar: (d[6] & 8) !== 0, v: (d[6] >> 4) & 7, rot: d[7], sangria: d[8] & 15,
      izq: bor & 15, der: (bor >> 4) & 15, arr: (bor >> 8) & 15, aba: (bor >> 12) & 15, cIzq: u16(d, 12) & 127, cDer: (u16(d, 12) >> 7) & 127, diag: (u16(d, 12) >> 14) & 3,
      cArr: col & 127, cAba: (col >> 7) & 127, cDiag: (col >> 14) & 127, eDiag: (col >> 21) & 15, trama: (col >>> 26) & 63, fg: u16(d, 18) & 127, bg: (u16(d, 18) >> 7) & 127 };
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
    const hojasDef = [], textos = [], formatos = new Map(), xfFormato = [], xfs = [], fuentes = [], paletaPropia = [], extensiones = new Map();
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
      } else if (r.id === 0x00E0) {                                                            // XF: fuente, formato…
        if (xfFormato.length < MAX_XF) { xfFormato.push(u16(d, 2)); xfs.push(version >= 0x0600 ? leerXf(d) : null); }
      } else if (r.id === 0x0031) { if (fuentes.length < MAX_FUENTES) fuentes.push(version >= 0x0600 ? leerFuente(d) : null); }   // FONT
      else if (r.id === 0x087D && version >= 0x0600) {                                         // XFEXT: colores RGB exactos de un XF (la paleta del .xls solo guarda 56)
        const xf = u16(d, 14), n = Math.min(u16(d, 18), 16);
        if (xf < MAX_XF && extensiones.size < MAX_XF) {
          const ext = extensiones.get(xf) || {};
          for (let k = 0, q = 20; k < n && q + 4 <= d.length; k++) {
            const tipo = u16(d, q), cb = u16(d, q + 2);
            if (cb < 4 || q + cb > d.length) break;
            if (tipo in TIPOS_EXT && cb >= 12 && u16(d, q + 4) === 2) {   // FullColorExt de tipo RGB
              const hex = (v) => ('0' + v.toString(16)).slice(-2);
              ext[TIPOS_EXT[tipo]] = hex(d[q + 8]) + hex(d[q + 9]) + hex(d[q + 10]);
            }
            q += cb;
          }
          extensiones.set(xf, ext);
        }
      }
      else if (r.id === 0x0092 && version >= 0x0600) {                                         // PALETTE: reemplaza los colores 8..63
        const hex = (n) => ('0' + n.toString(16)).slice(-2);
        for (let k = 0, q = 2; k < Math.min(u16(d, 0), 56) && q + 4 <= d.length; k++, q += 4) paletaPropia[k] = hex(d[q]) + hex(d[q + 1]) + hex(d[q + 2]);
      }
    }
    if (!fin && !hojasDef.length) throw danado('libro sin hojas');
    // un formato integrado sale como su número (crearLibro lo escribe tal cual y Excel lo muestra según el idioma de quien abre el libro)
    const codigoDe = (xf) => {
      const id = xfFormato[xf];
      if (id === undefined || id === 0) return null;
      if (formatos.has(id)) return formatos.get(id) === 'General' ? null : formatos.get(id);
      return XlsxEstilos.FORMATOS_INTEGRADOS[id] !== undefined && id < 164 ? id : null;
    };

    // color de un índice de la paleta: 0-7 y 8-63 de la paleta (la propia, si el libro trae PALETTE), 64 y 65 del sistema; null si es "automático"
    const colorDe = (i) => {
      if (i === 64) return '000000';
      if (i === 65) return 'FFFFFF';
      if (i >= 64) return null;
      return (i >= 8 && paletaPropia[i - 8]) || XlsxEstilos.PALETA_INDEXADA[i];
    };
    // El estilo de cada XF se arma una vez y lo comparten todas sus celdas. Sin estilo (null): el XF por defecto, de fuente Calibri 11 y nada más.
    const estilos = new Map(), visibles = new Set();   // visibles: XF con relleno, borde o alineación (una celda vacía solo se guarda si se nota)
    const estiloDe = (xf) => {
      if (estilos.has(xf)) return estilos.get(xf);
      let est = null;
      const x = xfs[xf];
      if (x && !x.estiloXf) {
        est = {};
        const cx = extensiones.get(xf) || {};   // los colores RGB exactos, si el archivo los trae
        const fx = fuentes[x.fuente >= 4 ? x.fuente - 1 : x.fuente];   // el índice 4 no existe nunca
        if (fx) {
          const f = { nombre: fx.nombre, tam: fx.alto };
          if (fx.negrita) f.negrita = true;
          if (fx.cursiva) f.cursiva = true;
          if (fx.tachado) f.tachado = true;
          if (fx.sub === 1 || fx.sub === 0x21) f.subrayado = true; else if (fx.sub === 2 || fx.sub === 0x22) f.subrayado = 'double';
          const c = cx.texto || colorDe(fx.icv);
          if (c && c !== '000000') f.color = c;   // negro es el color automático
          est.fuente = f;
        }
        const tr = TRAMAS[x.trama];
        if (tr) {
          const fg = cx.fg || colorDe(x.fg), bg = cx.bg || colorDe(x.bg);
          if (tr === 'solid') { if (fg) est.relleno = fg; }
          else est.relleno = { patron: tr, ...(fg ? { fg } : {}), ...(bg ? { bg } : {}) };
        }
        const bordes = {}, lado = (k, e, ci, rgb) => { const c = rgb || colorDe(ci); if (BORDES[e]) bordes[k] = { e: BORDES[e], ...(c ? { c } : {}) }; };
        lado('izq', x.izq, x.cIzq, cx.izq); lado('der', x.der, x.cDer, cx.der); lado('arr', x.arr, x.cArr, cx.arr); lado('aba', x.aba, x.cAba, cx.aba);
        if (x.diag & 1) lado('diagA', x.eDiag, x.cDiag, cx.diag);
        if (x.diag & 2) lado('diagB', x.eDiag, x.cDiag, cx.diag);
        if (Object.keys(bordes).length) est.bordes = bordes;
        const a = {};
        if (ALIN_H[x.h]) a.h = ALIN_H[x.h];
        if (ALIN_V[x.v]) a.v = ALIN_V[x.v];
        if (x.ajustar) a.ajustar = true;
        if (x.sangria) a.sangria = x.sangria;
        if (x.rot && (x.rot <= 180 || x.rot === 255)) a.rotacion = x.rot;
        if (Object.keys(a).length) est.alinea = a;
        const f0 = est.fuente;
        if (!est.relleno && !est.bordes && !est.alinea && (!f0 || (f0.nombre === 'Calibri' && f0.tam === 11 && !f0.negrita && !f0.cursiva && !f0.tachado && !f0.subrayado && !f0.color))) est = null;
        else if (est.relleno || est.bordes || est.alinea) visibles.add(xf);
      }
      estilos.set(xf, est);
      return est;
    };

    const hojas = [];
    // Los subflujos de las hojas no se pisan: entre todas no recorren más que el flujo entero. Un libro hostil con cientos de hojas que apuntan al mismo
    // subflujo (o a uno sin EOF) lo leería una vez por hoja: 400 KB armaban gigas. Se deja margen para un EOF que falte en un libro real.
    let presupuesto = 4 * wb.length + 65536, celdas = 0;
    for (const def of hojasDef) {
      if (def.tipo !== 0) continue;   // 0: hoja de cálculo; 2: gráfico; 1 y 6: macros
      const filas = [], anchos = [], combinadas = [], columnasOcultas = new Set(), filasOcultas = [], altosFila = {};
      let vacias = 0;
      let profund = 0, ultimaFormula = null;
      const poner = (f, c, v) => {
        if (f >= MAX_FILAS || c >= MAX_COLS) return;
        if (++celdas > MAX_CELDAS) throw new LibroXlsx.ErrorXlsx('grande', 'más de ' + MAX_CELDAS + ' celdas');
        (filas[f] || (filas[f] = []))[c] = v;
      };
      // un valor con el estilo de su XF (si lo tiene)
      const con = (xf, v) => { const e = estiloDe(xf); return e ? { valor: v, estilo: e } : v; };
      const numero = (f, c, xf, n) => {
        if (!Number.isFinite(n)) return;
        const cod = codigoDe(xf), e = estiloDe(xf);
        poner(f, c, cod ? (e ? { valor: n, formato: cod, estilo: e } : { valor: n, formato: cod }) : e ? { valor: n, estilo: e } : n);
      };
      // una celda sin valor solo cuenta si se nota (relleno, borde o alineación); las que pasan del tope se descartan sin fallar
      const vacia = (f, c, xf) => {
        estiloDe(xf);
        if (!visibles.has(xf)) return;
        if (f >= MAX_FILAS || c >= MAX_COLS || ++vacias > MAX_VACIAS) return;
        if (filas[f] && filas[f][c] !== undefined) return;
        poner(f, c, { valor: null, estilo: estiloDe(xf) });
      };
      for (const r of registros(wb, def.pos)) {
        const d = r.d;
        if ((presupuesto -= 4 + d.length) < 0) throw danado('hojas que comparten su subflujo');
        if (r.id === 0x0809) { profund++; continue; }
        if (r.id === 0x000A) { if (--profund <= 0) break; continue; }
        if (profund !== 1) continue;
        switch (r.id) {
          case 0x0208:                                                                                 // ROW: la bandera 0x20 de grbit (fDyZero) es "oculta"
            if (d.length >= 14 && (u16(d, 12) & 0x20)) filasOcultas.push(u16(d, 0));
            if (d.length >= 14 && (u16(d, 12) & 0x40) && !(u16(d, 6) & 0x8000) && u16(d, 6) > 0) altosFila[u16(d, 0)] = u16(d, 6) / 20;   // alto propio (fUnsynced), en puntos
            break;
          case 0x0201: vacia(u16(d, 0), u16(d, 2), u16(d, 4)); break;                                  // BLANK
          case 0x00BE: {                                                                               // MULBLANK
            const f = u16(d, 0), c0 = u16(d, 2), c1 = u16(d, d.length - 2);
            for (let c = c0, q = 4; c <= c1 && q + 2 <= d.length - 2; c++, q += 2) vacia(f, c, u16(d, q));
            break;
          }
          case 0x0203: numero(u16(d, 0), u16(d, 2), u16(d, 4), doble(d, 6)); break;                    // NUMBER
          case 0x027E: numero(u16(d, 0), u16(d, 2), u16(d, 4), numeroRk(u32(d, 6))); break;             // RK
          case 0x00BD: {                                                                               // MULRK
            const f = u16(d, 0), c0 = u16(d, 2), c1 = u16(d, d.length - 2);
            for (let c = c0, q = 4; c <= c1 && q + 6 <= d.length - 2; c++, q += 6) numero(f, c, u16(d, q), numeroRk(u32(d, q + 2)));
            break;
          }
          case 0x00FD: { const t = textos[u32(d, 6)]; if (t) poner(u16(d, 0), u16(d, 2), con(u16(d, 4), t)); break; }    // LABELSST
          case 0x0204: case 0x00D6: {                                                                  // LABEL y RSTRING (texto con formato) de BIFF5, y algún BIFF8 antiguo
            const t = version >= 0x0600 ? cadenaBiff8(d, 6, 2).texto : texto5(d.subarray(8, 8 + u16(d, 6)));
            if (t) poner(u16(d, 0), u16(d, 2), con(u16(d, 4), t.length > MAX_TEXTO ? t.slice(0, MAX_TEXTO) : t));
            break;
          }
          case 0x0205:                                                                                 // BOOLERR
            poner(u16(d, 0), u16(d, 2), con(u16(d, 4), d[7] ? { error: ERRORES[d[6]] || '#VALUE!' } : d[6] === 1));
            break;
          case 0x0006: {                                                                               // FORMULA: se guarda el último resultado
            const f = u16(d, 0), c = u16(d, 2), xf = u16(d, 4);
            if (d[12] === 0xFF && d[13] === 0xFF) {
              if (d[6] === 0) ultimaFormula = { f, c, xf };                                            // el texto viene en el STRING que sigue
              else if (d[6] === 1) poner(f, c, con(xf, d[8] === 1));
              else if (d[6] === 2) poner(f, c, con(xf, { error: ERRORES[d[8]] || '#VALUE!' }));
            } else numero(f, c, u16(d, 4), doble(d, 6));
            break;
          }
          case 0x0207:                                                                                 // STRING (resultado de una fórmula)
            if (ultimaFormula) {
              const t = version >= 0x0600 ? cadenaBiff8(d, 0, 2).texto : texto5(d.subarray(2, 2 + u16(d, 0)));
              if (t) poner(ultimaFormula.f, ultimaFormula.c, con(ultimaFormula.xf, t.length > MAX_TEXTO ? t.slice(0, MAX_TEXTO) : t));
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
      hojas.push({ nombre: def.nombre, filas, anchos: Array.from(anchos.slice(0, ancho), (w) => w || 0), combinadas, estado: def.estado, columnasOcultas: [...columnasOcultas].filter((c) => c < ancho).sort((a, b) => a - b), filasOcultas: filasOcultas.filter((r) => r < filas.length), altosFila });
    }
    if (!hojas.length) throw new LibroXlsx.ErrorXlsx('sinLibro', 'sin hojas');
    return { hojas, fecha1904 };
  }

  globalThis.XlsLector = { leer, _leerCfb: leerCfb, _registros: registros, _leerSst: leerSst };
})();

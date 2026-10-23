/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Lector de .doc propio (Word 97-2003, formato binario [MS-DOC]). Lo entrega como el «contrato de documento» (ver docx-desde-contrato.js) y, con aDocx, como un .docx,
// que es lo que ya saben leer, mostrar e imprimir las herramientas de Word.
//
//   DocLector.parece(bytes)            → true si es un .doc de Word 97 o posterior (contenedor CFB con el flujo WordDocument)
//   await DocLector.leer(bytes)        → { contrato, mensajes: [clave…] }
//   await DocLector.aDocx(bytes)       → { bytes, mensajes }
//
// Lanza Error (con .codigo = 'cifrado' si tiene contraseña, 'antiguo' si es de Word 95 o anterior, 'danado' si está roto) con un mensaje listo para mostrar.
// Se llevan: texto (también con guardado rápido: tabla de piezas), estilos con nombre, formato de carácter (negrita, cursiva, subrayado, tachado, superíndice, subíndice,
// versalitas, mayúsculas, fuente, tamaño, color, resaltado), formato de párrafo (alineación, sangrías, espaciado, interlineado, tabulaciones, salto antes), listas con
// viñetas y numeradas, tablas (celdas combinadas, relleno, bordes, anidadas), imágenes en línea (PNG y JPEG; las DIB, EMF y WMF se omiten con un aviso), enlaces y marcadores, números de página,
// secciones con su hoja, encabezado y pie. No se llevan: notas, comentarios, cuadros de texto, formas flotantes, ecuaciones, texto borrado con control de cambios.
// Requiere cfb-propio.js (y docx-desde-contrato.js para aDocx). No depende del navegador.
(function () {
  'use strict';
  const MAX_TEXTO = 64 * 1024 * 1024;     // caracteres de todo el documento
  const MAX_ENTRADAS = 4 * 1024 * 1024;   // elementos de las tablas internas (piezas, formatos, estilos…)
  const MAX_PROFUNDIDAD_TABLAS = 8;
  const MAX_CAMPOS = 64;                  // campos anidados
  const MAX_MARCADORES = 100000;
  const MAX_PARRAFOS = 2000000;           // párrafos y celdas de todo el documento (la tabla de piezas puede repetir un mismo trozo de texto)
  const MAX_IMAGENES_BYTES = 200 * 1024 * 1024;
  const sinProto = (o) => Object.assign(Object.create(null), o);
  const acotar = (x, min, max) => Math.min(Math.max(x, min), max);

  class Danado extends Error { constructor(m, codigo) { super(m); this.codigo = codigo || 'danado'; } }

  // ───────── Lectura de bytes sin salirse del arreglo ─────────
  const u8 = (b, p) => (p >= 0 && p < b.length ? b[p] : 0);
  const u16 = (b, p) => u8(b, p) | (u8(b, p + 1) << 8);
  const i16 = (b, p) => { const v = u16(b, p); return v > 32767 ? v - 65536 : v; };
  const u32 = (b, p) => (u8(b, p) | (u8(b, p + 1) << 8) | (u8(b, p + 2) << 16) | (u8(b, p + 3) << 24)) >>> 0;
  const i32 = (b, p) => u32(b, p) | 0;

  const CP1252 = [0x20AC, 0x81, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x8D, 0x017D, 0x8F, 0x90, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x9D, 0x017E, 0x0178];
  function decodificarAnsi(b, desde, hasta) {
    let s = '';
    for (let i = desde; i < hasta; i += 8192) {
      const trozo = [];
      for (let k = i; k < Math.min(hasta, i + 8192); k++) { const c = b[k]; trozo.push(c >= 0x80 && c < 0xA0 ? CP1252[c - 0x80] : c); }
      s += String.fromCharCode.apply(null, trozo);
    }
    return s;
  }
  function decodificarUtf16(b, desde, hasta) {
    let s = '';
    for (let i = desde; i + 1 < hasta; i += 16384) {
      const trozo = [];
      for (let k = i; k + 1 < Math.min(hasta, i + 16384); k += 2) trozo.push(b[k] | (b[k + 1] << 8));
      s += String.fromCharCode.apply(null, trozo);
    }
    return s;
  }

  // ───────── Sprms ─────────
  // Recorre una lista de sprms (grpprl) y llama a f(opcode, bytes, inicioOperando, largoOperando).
  function sprms(g, f) {
    let p = 0;
    const n = g.length;
    for (let cuenta = 0; p + 2 <= n && cuenta < 4096; cuenta++) {
      const op = u16(g, p); p += 2;
      const spra = op >> 13;
      let largo, inicio = p;
      if (spra === 0 || spra === 1) largo = 1;
      else if (spra === 2 || spra === 4 || spra === 5) largo = 2;
      else if (spra === 3) largo = 4;
      else if (spra === 7) largo = 3;
      else if (op === 0xD608) { largo = u16(g, p) - 1; inicio = p + 2; if (largo < 0) return; }   // sprmTDefTable: cb (2 bytes) cuenta todo menos 1
      else if (op === 0xC615 && u8(g, p) === 255) return;                                             // sprmPChgTabs con tamaño especial: no se sigue
      else { largo = u8(g, p); inicio = p + 1; }
      if (inicio > n) return;
      if (inicio + largo > n) largo = n - inicio;
      f(op, g, inicio, largo);
      p = inicio + largo;
    }
  }

  const ICO = [null, '000000', '0000FF', '00FFFF', '00FF00', 'FF00FF', 'FF0000', 'FFFF00', 'FFFFFF', '000080', '008080', '008000', '800080', '800000', '808000', '808080', 'C0C0C0'];
  const RESALTADO = [null, 'black', 'blue', 'cyan', 'green', 'magenta', 'red', 'yellow', 'white', 'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray'];
  const hex2 = (n) => (n < 16 ? '0' : '') + n.toString(16).toUpperCase();
  // COLORREF de Word (r, g, b, bandera): 0xFF en la bandera es «automático»
  const colorRef = (b, p) => (u8(b, p + 3) === 0xFF ? null : hex2(u8(b, p)) + hex2(u8(b, p + 1)) + hex2(u8(b, p + 2)));

  // Un borde BRC80 (4 bytes): ancho en octavos de punto, tipo, color
  function brc80(b, p) {
    const ancho = u8(b, p), tipo = u8(b, p + 1);
    if (tipo === 0 || tipo === 255 || ancho === 255 || ancho === 0) return null;
    const ico = u8(b, p + 2) & 0x1F;
    return { tam: acotar(ancho, 2, 96), color: ICO[ico] || undefined, estilo: tipo === 3 ? 'doble' : tipo === 6 ? 'punteado' : tipo === 7 || tipo === 8 ? 'guiones' : 'simple' };
  }

  // Un borde BRC (8 bytes, el de Word 2000 en adelante): color, ancho en octavos de punto y tipo
  function brc8(b, p) {
    const ancho = u8(b, p + 4), tipo = u8(b, p + 5);
    if (tipo === 0 || tipo === 255 || ancho === 255 || ancho === 0) return null;
    return { tam: acotar(ancho, 2, 96), color: colorRef(b, p) || undefined, estilo: tipo === 3 ? 'doble' : tipo === 6 ? 'punteado' : tipo === 7 || tipo === 8 ? 'guiones' : 'simple' };
  }

  // Propiedades de carácter (CHP) de una lista de sprms. Los interruptores 0x80 («como el estilo») se ignoran y 0x81 («al revés del estilo») se resuelven con `base`.
  function leerChp(g, base) {
    const r = {}, extra = {};
    const interruptor = (clave, v) => {
      if (v === 0) r[clave] = false;
      else if (v === 1) r[clave] = true;
      else if (v === 0x81) r[clave] = !(base && base[clave]);
    };
    sprms(g, (op, b, p, n) => {
      const v = u8(b, p);
      switch (op) {
        case 0x0835: interruptor('negrita', v); break;
        case 0x0836: interruptor('cursiva', v); break;
        case 0x0837: interruptor('tachado', v); break;
        case 0x2A53: if (v === 1) r.tachado = true; break;
        case 0x083A: interruptor('versalitas', v); break;
        case 0x083B: interruptor('mayusculas', v); break;
        case 0x083C: interruptor('oculto', v); break;
        case 0x2A3E: r.subrayado = v === 0 ? false : (v === 3 ? 'doble' : true); break;
        case 0x2A42: { const c = ICO[v]; if (v === 0) r.color = undefined; else if (c) r.color = c; break; }
        case 0x6870: { const c = colorRef(b, p); if (c) r.color = c; break; }
        case 0x4A43: r.tamano = u16(b, p) / 2; break;
        case 0x2A48: r.sup = v === 1; r.sub = v === 2; break;
        case 0x4A4F: extra.ftc = u16(b, p); break;
        case 0x4A50: extra.ftcFE = u16(b, p); break;
        case 0x4A5E: extra.ftcCs = u16(b, p); break;   // sprmCFtcBi: fuente de escritura compleja (0x4A51 es la de «otros caracteres» y se ignora)
        case 0x2A0C: { const h = RESALTADO[v]; r.resaltado = h || null; break; }
        case 0x0855: extra.spec = v === 1; break;
        case 0x6A03: extra.imagen = u32(b, p); break;
        case 0x0806: extra.datosCampo = v === 1; break;   // sprmCFData: el  guarda datos de un campo (formulario, PAGEREF), no una imagen
        case 0x4A30: extra.estilo = u16(b, p); break;
        case 0x0800: extra.borrado = v === 1 || v === 0x81; break;   // Word escribe 0x81 («al revés del estilo»)
        case 0x6A09: if (n >= 4) extra.simbolo = { ftc: u16(b, p), c: u16(b, p + 2) }; break;
        default: break;
      }
    });
    return { r, extra };
  }

  const JC = [ 'izquierda', 'centro', 'derecha', 'justificado', 'justificado' ];
  // Propiedades de párrafo (PAP). `tap` recoge las de tabla (TAP) que viajan en el mismo grpprl.
  function leerPap(g, base) {
    const p = {}, extra = { tap: {} };
    const interruptor = (clave, v) => { if (v === 0) p[clave] = false; else if (v === 1) p[clave] = true; else if (v === 0x81) p[clave] = !(base && base[clave]); };
    sprms(g, (op, b, q, n) => {
      const v = u8(b, q);
      switch (op) {
        case 0x4600: extra.istd = u16(b, q); break;
        case 0x2403: case 0x2461: p.alineacion = JC[v] || 'izquierda'; break;
        case 0x2405: interruptor('mantenerJuntas', v); break;
        case 0x2406: interruptor('mantener', v); break;
        case 0x2407: interruptor('saltoAntes', v); break;
        case 0x260A: extra.ilvl = v; break;
        case 0x460B: extra.ilfo = i16(b, q); break;
        case 0x2416: extra.enTabla = v === 1; break;
        case 0x2417: extra.ttp = v === 1; break;
        case 0x244B: extra.celdaInterna = v === 1; break;   // en una tabla anidada la marca de celda es un \r con esta bandera
        case 0x244C: extra.ttpInterna = v === 1; break;
        case 0x6649: extra.itap = u32(b, q); break;
        case 0x840F: case 0x845E: p.izq = i16(b, q); break;
        case 0x8411: case 0x8460: p.primera = i16(b, q); break;
        case 0x840E: case 0x845D: p.der = i16(b, q); break;
        case 0x6412: { const d = i16(b, q), mult = u16(b, q + 2); if (d !== 0) { if (mult === 1) { p.linea = Math.abs(d); p.regla = 'auto'; } else if (d > 0) { p.linea = d; p.regla = 'minimo'; } else { p.linea = -d; p.regla = 'exacto'; } } break; }
        case 0xA413: p.antes = u16(b, q); break;
        case 0xA414: p.despues = u16(b, q); break;
        case 0x2640: if (v <= 8) p.nivelEsquema = v; break;
        case 0xC60D: case 0xC615: {
          // cb, itbdDelMax, rgdxaDel[n], rgdxaClose[n], itbdAddMax, rgdxaAdd[m], rgtbdAdd[m]
          let o = q + 1;
          const borrar = u8(b, o); o += 1 + 4 * borrar;
          const anadir = u8(b, o); o += 1;
          const lista = [];
          for (let k = 0; k < anadir && k < 64; k++) {
            const t = u8(b, o + 2 * anadir + k);
            lista.push({ pos: i16(b, o + 2 * k), tipo: ['izquierda', 'centro', 'derecha', 'decimal', 'izquierda'][t & 7] || 'izquierda', relleno: [null, 'punto', 'guion', 'linea', 'linea'][(t >> 3) & 7] || null });
          }
          const quitar = [];
          for (let k = 0; k < borrar && k < 64; k++) quitar.push(i16(b, q + 2 + 2 * k));
          extra.tabs = { lista, quitar };
          break;
        }
        // tablas (TAP)
        case 0xD608: extra.tap.def = { desde: q, largo: n }; extra.tapBytes = b; break;
        case 0xD609: extra.tap.shd80 = { b, q, n }; break;
        case 0xD612: extra.tap.shd1 = { b, q, n }; break;
        case 0xD613: extra.tap.bordes = { b, q, n, ancho: 8 }; break;     // sprmTTableBorders: seis BRC de 8 bytes
        case 0xD605: if (!extra.tap.bordes) extra.tap.bordes = { b, q, n, ancho: 4 }; break;   // sprmTTableBorders80: seis BRC de 4 bytes
        case 0xD616: extra.tap.shd2 = { b, q, n }; break;   // sprmTDefTableShd2nd: celdas 22 a 43
        case 0xD60C: extra.tap.shd3 = { b, q, n }; break;   // sprmTDefTableShd3rd: celdas 44 a 62
        case 0x9407: extra.tap.alto = i16(b, q); break;
        case 0x3404: extra.tap.encabezado = (v & 1) === 1; break;
        case 0x9601: extra.tap.izq = i16(b, q); break;
        default: break;
      }
    });
    return { p, extra };
  }

  const NFC = { 0: 'decimal', 1: 'upperRoman', 2: 'lowerRoman', 3: 'upperLetter', 4: 'lowerLetter', 5: 'ordinal', 6: 'cardinalText', 7: 'ordinalText', 22: 'decimalZero', 23: 'bullet', 255: 'none' };

  // ───────── El documento ─────────
  class Documento {
    constructor(bytes) {
      this.mensajes = new Set();
      this.abrir(bytes);
    }
    aviso(c) { this.mensajes.add(c); }

    abrir(bytes) {
      let cfb;
      try { cfb = CfbPropio.leer(bytes, { danado: (d) => new Danado('El archivo está dañado y no se puede abrir (' + d + ').') }); } catch (e) {
        if (e instanceof Danado) throw e;
        throw new Danado('El archivo no es un documento de Word 97-2003 válido.');
      }
      const entrada = (n) => cfb.entradas.find((e) => e.tipo === 2 && e.nombre === n);
      const eWd = entrada('WordDocument');
      // Un .docx con contraseña también es un contenedor CFB (con EncryptedPackage en vez de WordDocument)
      if (!eWd && entrada('EncryptedPackage')) throw new Danado('El documento está protegido con contraseña. Sacale la contraseña en Word y volvé a intentarlo.', 'cifrado');
      if (!eWd) throw new Danado('El archivo no es un documento de Word (.doc).');
      this.wd = cfb.flujo(eWd);
      const wd = this.wd;
      if (wd.length < 0x9A || u16(wd, 0) !== 0xA5EC) throw new Danado('El archivo no es un documento de Word 97-2003 válido.');
      const nFib = u16(wd, 2);
      if (nFib < 0xC0) throw new Danado('Este documento es de una versión de Word anterior a Word 97 y no se puede leer. Abrilo en Word y guardalo como .docx.', 'antiguo');
      const flags = u16(wd, 0x0A);
      if (flags & 0x100) throw new Danado('El documento está protegido con contraseña. Sacale la contraseña en Word y volvé a intentarlo.', 'cifrado');
      const eTabla = entrada((flags & 0x200) ? '1Table' : '0Table');
      if (!eTabla) throw new Danado('El documento está dañado: falta la tabla de estructura.');
      this.tabla = cfb.flujo(eTabla);
      const eData = entrada('Data');
      this.datos = eData ? cfb.flujo(eData) : new Uint8Array(0);
      this.pares = u16(wd, 0x98);
      this.ccp = { texto: u32(wd, 0x4C), ftn: u32(wd, 0x50), hdd: u32(wd, 0x54), atn: u32(wd, 0x5C), edn: u32(wd, 0x60), txbx: u32(wd, 0x64), hdrTxbx: u32(wd, 0x68) };
      this.leerTexto();
      this.leerFormatos();
      this.leerFuentes();
      this.leerEstilos();
      this.leerListas();
      this.leerMarcadores();
    }

    // (fc, lcb) del índice i del bloque FibRgFcLcb97
    par(i) { return i < this.pares ? [u32(this.wd, 0x9A + 8 * i), u32(this.wd, 0x9A + 8 * i + 4)] : [0, 0]; }
    enTabla(fc, lcb) { return fc + lcb <= this.tabla.length ? this.tabla.subarray(fc, fc + lcb) : this.tabla.subarray(Math.min(fc, this.tabla.length), this.tabla.length); }

    // ───────── Texto: la tabla de piezas (CLX) ─────────
    leerTexto() {
      const [fc, lcb] = this.par(33);
      const t = this.enTabla(fc, lcb);
      let p = 0;
      while (p < t.length && t[p] === 1) p += 3 + u16(t, p + 1);   // los Prc (propiedades de piezas) se saltan
      if (t[p] !== 2) throw new Danado('El documento está dañado: no tiene tabla de piezas de texto.');
      const tam = u32(t, p + 1);
      const n = Math.floor((tam - 4) / 12);
      if (n < 0 || n > MAX_ENTRADAS || p + 5 + tam > t.length + 8) throw new Danado('El documento está dañado: tabla de piezas inválida.');
      const base = p + 5;
      const piezas = [];
      const partes = [];
      let total = 0;
      for (let i = 0; i < n; i++) {
        const cp0 = u32(t, base + 4 * i), cp1 = u32(t, base + 4 * (i + 1));
        const fcRaw = u32(t, base + 4 * (n + 1) + 8 * i + 2);
        const comp = (fcRaw & 0x40000000) !== 0;
        const fcp = comp ? (fcRaw & 0x3FFFFFFF) >>> 1 : fcRaw & 0x3FFFFFFF;
        if (cp1 < cp0) throw new Danado('El documento está dañado: piezas de texto desordenadas.');
        if (cp0 !== total) throw new Danado('El documento está dañado: piezas de texto con huecos.');
        let largo = cp1 - cp0;
        const bpc = comp ? 1 : 2;
        if (fcp >= this.wd.length) largo = 0;
        else if (fcp + largo * bpc > this.wd.length) largo = Math.floor((this.wd.length - fcp) / bpc);
        if (total + (cp1 - cp0) > MAX_TEXTO) throw new Danado('El documento es demasiado grande para abrirlo aquí.');
        let s = largo ? (comp ? decodificarAnsi(this.wd, fcp, fcp + largo) : decodificarUtf16(this.wd, fcp, fcp + 2 * largo)) : '';
        if (s.length < cp1 - cp0) s += ' '.repeat(cp1 - cp0 - s.length);
        partes.push(s);
        piezas.push({ cp0, cp1, fc: fcp, comp });
        total = cp1;
      }
      this.piezas = piezas;
      this.T = partes.join('');
    }

    // fc del carácter en la posición cp (en el flujo WordDocument)
    fcDe(cp) {
      const ps = this.piezas;
      let lo = 0, hi = ps.length - 1;
      while (lo < hi) { const m = (lo + hi + 1) >> 1; if (ps[m].cp0 <= cp) lo = m; else hi = m - 1; }
      const q = ps[lo];
      if (!q) return -1;
      return q.fc + (cp - q.cp0) * (q.comp ? 1 : 2);
    }
    // Número de bytes por carácter en cp
    bytesPorCar(cp) {
      const ps = this.piezas;
      let lo = 0, hi = ps.length - 1;
      while (lo < hi) { const m = (lo + hi + 1) >> 1; if (ps[m].cp0 <= cp) lo = m; else hi = m - 1; }
      return ps[lo] && ps[lo].comp ? 1 : 2;
    }

    // ───────── Formatos de carácter y de párrafo (FKP) ─────────
    leerFormatos() {
      const wd = this.wd;
      const leerPlc = (indice, tipo) => {
        const [fc, lcb] = this.par(indice);
        const t = this.enTabla(fc, lcb);
        const n = Math.floor((t.length - 4) / 8);
        const lista = [];
        if (n <= 0 || n > MAX_ENTRADAS) return lista;
        for (let i = 0; i < n; i++) {
          const pn = u32(t, 4 * (n + 1) + 4 * i) & 0x3FFFFF;
          const ini = pn * 512;
          if (ini + 512 > wd.length) continue;
          const fkp = wd.subarray(ini, ini + 512);
          const crun = fkp[511];
          for (let k = 0; k < crun; k++) {
            const fc0 = u32(fkp, 4 * k), fc1 = u32(fkp, 4 * (k + 1));
            if (fc1 <= fc0) continue;
            if (tipo === 'chp') {
              const o = fkp[4 * (crun + 1) + k] * 2;
              lista.push({ fc0, fc1, g: o ? fkp.subarray(o + 1, o + 1 + fkp[o]) : null });
            } else {
              const o = fkp[4 * (crun + 1) + 13 * k] * 2;
              let istd = 0, g = null;
              if (o) {
                let cb = fkp[o], ini2 = o + 1, tam;
                if (cb === 0) { cb = fkp[o + 1]; ini2 = o + 2; tam = cb * 2; } else tam = cb * 2 - 1;
                if (tam >= 2) { istd = u16(fkp, ini2); g = fkp.subarray(ini2 + 2, ini2 + tam); }
              }
              if (g) g = this.papGrande(g);
              lista.push({ fc0, fc1, istd, g });
            }
            if (lista.length > MAX_ENTRADAS) throw new Danado('El documento está dañado: demasiados formatos.');
          }
        }
        lista.sort((a, b) => a.fc0 - b.fc0);
        return lista;
      };
      this.chp = leerPlc(12, 'chp');
      this.pap = leerPlc(13, 'pap');
    }
    // sprmPHugePapx (0x6646): las propiedades que no caben en la página de formatos (p. ej. una fila de tabla con muchas columnas) están en el flujo Data
    // (PrcData: cbGrpprl de 2 bytes y la lista de sprms). La norma dice que ese sprm va solo en el FKP: se usa la lista del flujo Data (una vista, sin copiar).
    papGrande(g) {
      let desp = -1;
      sprms(g, (op, b, p) => { if (op === 0x6646) desp = u32(b, p); });
      if (desp < 0) return g;
      const dt = this.datos;
      const cb = u16(dt, desp);
      if (desp + 2 >= dt.length || cb === 0 || cb > 0x3FA2) return g;
      return dt.subarray(desp + 2, Math.min(dt.length, desp + 2 + cb));
    }
    buscar(lista, fc) {
      let lo = 0, hi = lista.length - 1;
      while (lo <= hi) {
        const m = (lo + hi) >> 1;
        const e = lista[m];
        if (fc < e.fc0) hi = m - 1; else if (fc >= e.fc1) lo = m + 1; else return e;
      }
      return null;
    }

    // ───────── Fuentes ─────────
    leerFuentes() {
      const [fc, lcb] = this.par(15);
      const t = this.enTabla(fc, lcb);
      this.fuentes = [];
      const n = u16(t, 0);
      let p = 4;
      for (let i = 0; i < n && i < 4096 && p < t.length; i++) {
        const cb = t[p];
        let nombre = '';
        for (let q = p + 40; q + 1 < p + 1 + cb && q + 1 < t.length; q += 2) { const c = t[q] | (t[q + 1] << 8); if (c === 0) break; nombre += String.fromCharCode(c); }
        this.fuentes.push(nombre);
        p += cb + 1;
      }
    }
    fuente(i) { return i != null && this.fuentes[i] ? this.fuentes[i] : null; }

    // ───────── Estilos ─────────
    leerEstilos() {
      const [fc, lcb] = this.par(1);
      const t = this.enTabla(fc, lcb);
      this.estilos = [];
      this.stshi = { ftcAsci: 0 };
      if (t.length < 20) return;
      const cbStshi = u16(t, 0);
      const cstd = u16(t, 2), cbBase = u16(t, 4);
      this.stshi.ftcAsci = u16(t, 14);
      this.stshi.ftcFE = u16(t, 16);
      this.stshi.ftcBi = cbStshi >= 20 ? u16(t, 20) : u16(t, 18);
      let p = 2 + cbStshi;
      for (let i = 0; i < cstd && i < 8192 && p + 2 <= t.length; i++) {
        const cb = u16(t, p);
        const ini = p + 2;
        p = ini + cb;
        if (cb === 0 || cb < cbBase) { this.estilos.push(null); continue; }
        const sti = u16(t, ini) & 0xFFF;
        const w1 = u16(t, ini + 2), w2 = u16(t, ini + 4);
        const stk = w1 & 0xF, istdBase = w1 >> 4, cupx = w2 & 0xF, istdNext = w2 >> 4;
        let q = ini + cbBase;
        const cch = u16(t, q);
        let nombre = '';
        for (let k = 0; k < cch && k < 255; k++) nombre += String.fromCharCode(u16(t, q + 2 + 2 * k));
        q += 4 + 2 * cch;
        const upx = [];
        for (let k = 0; k < cupx && q + 2 <= ini + cb; k++) {
          const cbu = u16(t, q);
          upx.push(t.subarray(q + 2, Math.min(q + 2 + cbu, ini + cb)));
          q += 2 + cbu + (cbu & 1);
        }
        let papx = null, chpx = null, istdPapx = 0;
        if (stk === 1) { if (upx[0]) { istdPapx = u16(upx[0], 0); papx = upx[0].subarray(2); } if (upx[1]) chpx = upx[1]; }
        else if (stk === 2) { if (upx[0]) chpx = upx[0]; }
        this.estilos.push({ sti, stk, base: istdBase === 0xFFF ? null : istdBase, siguiente: istdNext === 0xFFF ? null : istdNext, nombre, papx, chpx, istdPapx });
      }
    }
    nombreEstilo(i, e) {
      if (e.nombre) return e.nombre;
      if (e.sti >= 1 && e.sti <= 9) return 'heading ' + e.sti;
      if (e.sti === 0) return 'Normal';
      return 'Estilo' + i;
    }
    // Propiedades de carácter heredadas del estilo de párrafo (para los interruptores «al revés»)
    cacheR(i) {
      this._cacheR = this._cacheR || new Map();
      if (this._cacheR.has(i)) return this._cacheR.get(i);
      const cadena = [];
      for (let k = 0, e = this.estilos[i], n = i; e && k < 32; k++, n = e.base, e = n == null ? null : this.estilos[n]) cadena.unshift(e);
      const r = {};
      for (const e of cadena) if (e.chpx) Object.assign(r, leerChp(e.chpx, null).r);
      this._cacheR.set(i, r);
      return r;
    }
    cadenaPap(i) {
      const cad = [];
      for (let k = 0, e = this.estilos[i], n = i; e && k < 32; k++, n = e.base, e = n == null ? null : this.estilos[n]) cad.unshift(e);
      return cad;
    }

    // ───────── Marcadores ─────────
    // Marcadores (SttbfBkmk con los nombres y PlcfBkf con el CP donde empieza cada uno): para que los enlaces internos (\l, el índice) tengan destino
    leerMarcadores() {
      this.marcadores = new Map();   // cp → [nombre…]
      const [fn, ln] = this.par(21), [fi, li] = this.par(22);
      const nombres = this.enTabla(fn, ln), inicios = this.enTabla(fi, li);
      if (nombres.length < 6 || inicios.length < 8) return;
      const extendido = u16(nombres, 0) === 0xFFFF;
      const cuantos = u16(nombres, extendido ? 2 : 0);
      let p = extendido ? 6 : 4;
      const n = Math.min(cuantos, Math.floor((inicios.length - 4) / 8), MAX_MARCADORES);
      for (let i = 0; i < n && p + 2 <= nombres.length; i++) {
        const cch = u16(nombres, p);
        let nombre = '';
        for (let k = 0; k < cch && k < 255; k++) nombre += String.fromCharCode(u16(nombres, p + 2 + 2 * k));
        p += 2 + 2 * cch;
        const cp = u32(inicios, 4 * i);
        if (!nombre) continue;
        if (!this.marcadores.has(cp)) this.marcadores.set(cp, []);
        this.marcadores.get(cp).push(nombre);
      }
    }

    // ───────── Listas ─────────
    leerListas() {
      this.lstf = new Map();   // lsid → niveles
      const [fc] = this.par(73);
      // Los niveles (LVL) van justo después de las cabeceras pero fuera de lo que cuenta lcbPlfLst: se lee el resto del flujo
      const t = this.tabla.subarray(Math.min(fc, this.tabla.length));
      const cLst = u16(t, 0);
      if (cLst > 0 && cLst < 30000) {
        const cabeceras = [];
        let p = 2;
        for (let i = 0; i < cLst && p + 28 <= t.length; i++, p += 28) cabeceras.push({ lsid: i32(t, p), simple: (t[p + 26] & 1) === 1 });
        for (const h of cabeceras) {
          const nivs = [];
          const n = h.simple ? 1 : 9;
          for (let k = 0; k < n && p + 28 <= t.length; k++) { const r = this.leerLvl(t, p); nivs.push(r.nivel); p = r.fin; }
          this.lstf.set(h.lsid, nivs);
        }
      }
      // PlfLfo: cada lista en uso, con sus cambios de nivel
      this.lfo = [];
      const [fc2, lcb2] = this.par(74);
      const l = this.enTabla(fc2, lcb2);
      const n = u32(l, 0);
      if (n > 0 && n < 100000) {
        let p = 4;
        for (let i = 0; i < n && p + 16 <= l.length; i++, p += 16) this.lfo.push({ lsid: i32(l, p), clfolvl: l[p + 12], niveles: new Map() });
        // rgLfoData: cada LFO tiene un LFOData (cp de 4 bytes y luego sus clfolvl cambios de nivel), también los que no cambian nada
        for (const f of this.lfo) {
          p += 4;
          for (let k = 0; k < f.clfolvl && p + 8 <= l.length; k++) {
            const inicio = i32(l, p), fl = u32(l, p + 4);
            p += 8;
            const nivel = fl & 0xF, tieneInicio = (fl & 0x10) !== 0, formato = (fl & 0x20) !== 0;
            const o = { inicio: tieneInicio ? inicio : null, nivel: null };
            if (formato && p + 28 <= l.length) { const r = this.leerLvl(l, p); o.nivel = r.nivel; p = r.fin; }
            f.niveles.set(nivel, o);
          }
        }
      }
    }
    leerLvl(t, p) {
      const inicio = i32(t, p), nfc = t[p + 4];
      const cbChpx = t[p + 24], cbPapx = t[p + 25];
      let q = p + 28;
      const papx = t.subarray(q, q + cbPapx); q += cbPapx;
      const chpx = t.subarray(q, q + cbChpx); q += cbChpx;
      const cch = u16(t, q);
      let texto = '';
      for (let k = 0; k < cch && k < 255; k++) { const c = u16(t, q + 2 + 2 * k); texto += c < 9 ? '%' + (c + 1) : String.fromCharCode(c); }
      q += 2 + 2 * cch;
      const nivel = { formato: NFC[nfc] || 'decimal', inicio: inicio < 0 ? 1 : inicio, texto, izquierda: 0, colgante: 0 };
      if (nivel.formato === 'none') nivel.texto = '';
      const pp = leerPap(papx, null).p;
      nivel.izquierda = acotar(pp.izq || 0, 0, 31680);
      nivel.colgante = acotar(-(pp.primera || 0), 0, 31680);
      const cc = leerChp(chpx, null);
      const f = this.fuente(cc.extra.ftc);
      if (f && nivel.formato === 'bullet') nivel.fuente = f;
      return { nivel, fin: q };
    }
  }

  // ───────── Conversión al contrato ─────────
  class Conversor {
    constructor(doc) {
      this.d = doc;
      this.contrato = { propiedades: {}, predeterminado: {}, estilos: [], listas: [], secciones: [] };
      this.bytesImagenes = 0;
      this.listasUsadas = new Map();   // ilfo → id de lista del contrato
      this.campos = [];                // pila de campos abiertos (pueden cruzar párrafos)
      this.camposDeMas = 0;
      this.parrafosLeidos = 0;
      this._deEstilo = new Map();      // istd → lo que el estilo de párrafo aporta (lista, nivel, tabla)
    }

    idEstilo(i) {
      const e = this.d.estilos[i];
      if (!e || e.sti === 65 || e.sti === 105 || e.stk > 2) return null;
      return i === 0 ? 'Normal' : 'S' + i;
    }

    // ───────── Estilos con nombre ─────────
    estilosDelContrato() {
      const d = this.d;
      const out = [];
      d.estilos.forEach((e, i) => {
        const id = e && this.idEstilo(i);
        if (!id) return;
        const pap = e.papx ? leerPap(e.papx, null) : { p: {}, extra: {} };
        const chp = e.chpx ? leerChp(e.chpx, null) : { r: {}, extra: {} };
        const r = Object.assign({}, chp.r);
        if (chp.extra.ftc != null) r.fuente = d.fuente(chp.extra.ftc) || undefined;
        if (chp.extra.ftcFE != null) r.fuenteAsia = d.fuente(chp.extra.ftcFE) || undefined;
        if (chp.extra.ftcCs != null) r.fuenteCs = d.fuente(chp.extra.ftcCs) || undefined;
        const p = pap.p;
        if (pap.extra.tabs) p.tabs = pap.extra.tabs.lista.map((q) => ({ tipo: q.tipo, pos: q.pos, relleno: q.relleno }));
        const base = e.base != null && this.idEstilo(e.base) && e.base !== i ? this.idEstilo(e.base) : undefined;
        const sig = e.siguiente != null && this.idEstilo(e.siguiente) ? this.idEstilo(e.siguiente) : undefined;
        out.push({ id, nombre: d.nombreEstilo(i, e), tipo: e.stk === 2 ? 'caracter' : 'parrafo', basadoEn: base, siguiente: sig, p: e.stk === 1 ? p : undefined, r });
      });
      this.contrato.estilos = out;
      const normal = d.estilos[0];
      const rn = normal && normal.chpx ? leerChp(normal.chpx, null) : { r: {}, extra: {} };
      // La fuente asiática o compleja que es la misma que la latina no se escribe: es el valor de relleno de Word, y sin ella Word elige la que corresponde al idioma
      const distinta = (i) => (i !== d.stshi.ftcAsci ? d.fuente(i) || undefined : undefined);
      this.contrato.predeterminado = { r: { fuente: d.fuente(d.stshi.ftcAsci) || 'Times New Roman', fuenteAsia: distinta(d.stshi.ftcFE), fuenteCs: distinta(d.stshi.ftcBi), tamano: 10 }, p: {} };
      void rn;
    }

    // ───────── Listas del contrato ─────────
    listaDe(ilfo) {
      if (this.listasUsadas.has(ilfo)) return this.listasUsadas.get(ilfo);
      const f = this.d.lfo[ilfo - 1];
      const base = f && this.d.lstf.get(f.lsid);
      if (!f || !base) { this.listasUsadas.set(ilfo, null); return null; }
      // Varios LFO de la misma lista (LST) sin cambios de nivel comparten el contador: la numeración sigue donde iba. Con cambios (empezar en…) cada uno cuenta aparte.
      const clave = f.niveles.size === 0 ? 'lsid' + f.lsid : null;
      if (clave && this.listasUsadas.has(clave)) { const previa = this.listasUsadas.get(clave); this.listasUsadas.set(ilfo, previa); return previa; }
      const niveles = [];
      for (let n = 0; n < 9; n++) {
        const o = f.niveles.get(n);
        const nivel = Object.assign({}, (o && o.nivel) || base[n] || { formato: 'decimal', texto: '%' + (n + 1) + '.', izquierda: 720 * (n + 1), colgante: 360 });
        if (o && o.inicio != null) nivel.inicio = o.inicio;
        niveles.push(nivel);
      }
      const id = 'F' + ilfo;
      this.contrato.listas.push({ id, niveles });
      this.listasUsadas.set(ilfo, id);
      if (clave) this.listasUsadas.set(clave, id);
      return id;
    }

    // ───────── Párrafos del documento ─────────
    // Los párrafos de un tramo de CP: [{ cp0, cp1 (sin la marca), marca, pap, extra… }]
    parrafosDe(cp0, cp1, finSecciones) {
      const d = this.d, T = d.T;
      const lista = [];
      let ini = cp0;
      const fin = Math.min(cp1, T.length);
      for (let cp = cp0; cp < fin; cp++) {
        const c = T.charCodeAt(cp);
        let termina = false;
        if (c === 13) termina = true;
        else if (c === 7 || c === 12) {
          const fc = d.fcDe(cp);
          const pap = fc >= 0 ? d.buscar(d.pap, fc) : null;
          if (c === 7) { const g = pap && pap.g ? leerPap(pap.g, null).extra : {}; termina = !!(g.enTabla || g.ttp || (pap && this.estiloEnTabla(pap.istd))); }
          else termina = finSecciones && finSecciones.has(cp);
        }
        if (!termina) continue;
        const fc = d.fcDe(cp);
        const pap = fc >= 0 ? d.buscar(d.pap, fc) : null;
        lista.push({ cp0: ini, cp1: cp, marca: c, pap });
        if (++this.parrafosLeidos > MAX_PARRAFOS) throw new Danado('El documento tiene demasiados párrafos y no se puede leer.');
        ini = cp + 1;
      }
      if (ini < fin) lista.push({ cp0: ini, cp1: fin, marca: 13, pap: null });
      return lista;
    }
    estiloEnTabla(istd) {
      this._enTabla = this._enTabla || new Map();
      if (!this._enTabla.has(istd)) this._enTabla.set(istd, this.d.cadenaPap(istd).some((e) => e.papx && leerPap(e.papx, null).extra.enTabla));
      return this._enTabla.get(istd);
    }

    // Propiedades de un párrafo: las del estilo para saber la lista y los interruptores, y las directas
    propiedadesParrafo(rec) {
      const d = this.d;
      const istd = rec.pap ? rec.pap.istd : 0;
      const dir = rec.pap && rec.pap.g ? leerPap(rec.pap.g, null) : { p: {}, extra: { tap: {} } };
      let deEstilo = this._deEstilo.get(istd);
      if (!deEstilo) {
        deEstilo = { p: {}, extra: {} };
        for (const e of d.cadenaPap(istd)) if (e.papx) { const r = leerPap(e.papx, null); Object.assign(deEstilo.p, r.p); Object.assign(deEstilo.extra, { ilfo: r.extra.ilfo != null ? r.extra.ilfo : deEstilo.extra.ilfo, ilvl: r.extra.ilvl != null ? r.extra.ilvl : deEstilo.extra.ilvl, enTabla: r.extra.enTabla || deEstilo.extra.enTabla }); }
        this._deEstilo.set(istd, deEstilo);
      }
      const ex = dir.extra;
      const ilfo = ex.ilfo != null ? ex.ilfo : deEstilo.extra.ilfo;
      const ilvl = ex.ilvl != null ? ex.ilvl : deEstilo.extra.ilvl;
      return { istd, p: dir.p, extra: ex, ilfo: ilfo || 0, ilvl: ilvl || 0, enTabla: !!(ex.enTabla || deEstilo.extra.enTabla), deEstilo };
    }

    // ───────── Texto en línea ─────────
    // Los trozos de [cp0, cp1) con un mismo formato de carácter y una misma pieza
    trozos(cp0, cp1) {
      const d = this.d;
      const salida = [];
      let cp = cp0;
      while (cp < cp1) {
        const fc = d.fcDe(cp);
        if (fc < 0) break;
        const bpc = d.bytesPorCar(cp);
        const ps = d.piezas;
        if (!ps.length || cp >= ps[ps.length - 1].cp1) break;
        // fin de la pieza
        let lo = 0, hi = ps.length - 1;
        while (lo < hi) { const m = (lo + hi + 1) >> 1; if (ps[m].cp0 <= cp) lo = m; else hi = m - 1; }
        let fin = Math.min(cp1, ps[lo].cp1);
        const r = d.buscar(d.chp, fc);
        if (r) fin = Math.min(fin, cp + Math.max(1, Math.ceil((r.fc1 - fc) / bpc)));
        else {
          // hasta el siguiente formato (para no ir carácter por carácter)
          let l = 0, h = d.chp.length - 1, sig = -1;
          while (l <= h) { const m = (l + h) >> 1; if (d.chp[m].fc0 > fc) { sig = m; h = m - 1; } else l = m + 1; }
          if (sig >= 0) fin = Math.min(fin, cp + Math.max(1, Math.ceil((d.chp[sig].fc0 - fc) / bpc)));
        }
        salida.push({ cp0: cp, cp1: fin, g: r ? r.g : null });
        cp = fin;
      }
      return salida;
    }

    formatoCaracter(g, istdParrafo) {
      const base = this.d.cacheR(istdParrafo);
      const c = g ? leerChp(g, base) : { r: {}, extra: {} };
      const r = c.r;
      if (c.extra.ftc != null) { const f = this.d.fuente(c.extra.ftc); if (f) r.fuente = f; }
      if (c.extra.ftcFE != null) { const f = this.d.fuente(c.extra.ftcFE); if (f) r.fuenteAsia = f; }
      if (c.extra.ftcCs != null) { const f = this.d.fuente(c.extra.ftcCs); if (f) r.fuenteCs = f; }
      const estilo = c.extra.estilo != null ? this.idEstilo(c.extra.estilo) : null;
      return { r, extra: c.extra, estilo };
    }

    agregarTexto(salida, texto, estilo, r) {
      if (!texto) return;
      const u = salida[salida.length - 1];
      const clave = JSON.stringify(r || {});
      if (u && u.texto !== undefined && u.estilo === estilo && u._k === clave) { u.texto += texto; return; }
      const run = { texto, _k: clave };
      if (estilo) run.estilo = estilo;
      if (r && Object.keys(r).length) run.r = r;
      salida.push(run);
    }

    // Adónde va el texto que está dentro de los primeros k campos abiertos: al párrafo, al resultado de un enlace o de un campo que se reemplaza, o a ninguna parte
    // (instrucción de un campo). El resultado de los demás campos (TOC, PAGEREF, REF…) se lleva tal cual, aunque ocupe varios párrafos.
    destinoEn(k) {
      for (let i = k - 1; i >= 0; i--) {
        const c = this.campos[i];
        if (c.fase !== 'res') return null;
        if (c.tipo !== 'pasar') return c.salida;
      }
      return this._raiz;
    }
    tipoCampoNombre(c) { const palabra = (/^\s*([A-Za-z]+)/.exec(c.instr) || [])[1]; return palabra ? palabra.toUpperCase() : ''; }
    tipoCampo(c) {
      const nombre = this.tipoCampoNombre(c);
      return nombre === 'HYPERLINK' ? 'enlace' : nombre === 'PAGE' || nombre === 'NUMPAGES' ? 'campo' : 'pasar';
    }
    enlaceDe(c, salida) {
      const instr = c.instr.trim();
      const m = /HYPERLINK\s+(?:\\l\s+)?(?:"([^"]*)"|(\S+))/i.exec(instr);
      const ancla = /\\l\s+"([^"]*)"/i.exec(instr);
      const url = m && (m[1] || m[2]);
      const hijos = salida.map((x) => { const y = Object.assign({}, x); delete y._k; return y; });
      if (ancla) return [{ enlace: { ancla: ancla[1] }, hijos }];
      if (url) return [{ enlace: { url }, hijos }];
      return hijos;
    }

    // Convierte el texto de un párrafo [cp0, cp1) en runs
    runsDe(cp0, cp1, istd) {
      const d = this.d, T = d.T;
      const raiz = [];
      this._raiz = raiz;
      const destino = () => this.destinoEn(this.campos.length);
      const marcas = d.marcadores;
      for (const tz of this.trozos(cp0, cp1)) {
        const f = this.formatoCaracter(tz.g, istd);
        if (f.extra.borrado) continue;   // texto borrado con control de cambios
        const textoBruto = T.slice(tz.cp0, tz.cp1);
        let acumulado = '';
        const vaciar = () => { if (acumulado) { const dst = destino(); if (dst) this.agregarTexto(dst, acumulado, f.estilo, f.r); acumulado = ''; } };
        for (let i = 0; i < textoBruto.length; i++) {
          const ch = textoBruto.charCodeAt(i);
          if (marcas.size && marcas.has(tz.cp0 + i)) { vaciar(); const dst = destino(); if (dst) for (const m of marcas.get(tz.cp0 + i)) dst.push({ marcador: m }); }
          switch (ch) {
            case 9: acumulado += '\t'; break;
            case 11: case 10: acumulado += '\n'; break;
            case 30: acumulado += '‑'; break;
            case 31: acumulado += '­'; break;
            case 12: { vaciar(); const dst = destino(); if (dst) dst.push({ salto: 'pagina' }); break; }
            case 19: {
              vaciar();
              // más de MAX_CAMPOS campos anidados (solo en un archivo hostil): los de más se cuentan y se ignoran, para no llenar la memoria
              if (this.campos.length >= MAX_CAMPOS) this.camposDeMas++;
              else this.campos.push({ instr: '', fase: 'instr', salida: [], f, tipo: null });
              break;
            }
            case 20: {
              vaciar();
              const c = this.campos[this.campos.length - 1];
              if (this.camposDeMas === 0 && c && c.fase === 'instr') { c.fase = 'res'; c.tipo = this.tipoCampo(c); }
              break;
            }
            case 21: { vaciar(); if (this.camposDeMas > 0) this.camposDeMas--; else this.cerrarCampo(); break; }
            case 1: {
              vaciar();
              const dst = destino();
              if (dst && f.extra.spec && f.extra.imagen != null && !f.extra.datosCampo) {
                const im = this.imagen(f.extra.imagen);
                if (im) dst.push({ imagen: im });
              }
              break;
            }
            case 8: vaciar(); this.d.aviso('dibujos'); break;
            case 2: vaciar(); break;
            case 5: vaciar(); break;
            case 40: if (f.extra.spec && f.extra.simbolo) { acumulado += String.fromCharCode(f.extra.simbolo.c); } else acumulado += '('; break;
            default:
              if (ch < 32) break;
              acumulado += textoBruto[i];
          }
          // El texto de la instrucción de un campo se guarda aparte
          const c = this.campos[this.campos.length - 1];
          if (c && c.fase === 'instr' && acumulado) { if (c.instr.length < 4096) c.instr += acumulado; acumulado = ''; }
        }
        vaciar();
      }
      // Un enlace que sigue en el párrafo siguiente: lo que tiene hasta aquí queda en este párrafo (como enlace al mismo destino)
      for (let i = this.campos.length - 1; i >= 0; i--) {
        const c = this.campos[i];
        if (c.fase !== 'res' || c.tipo !== 'enlace' || !c.salida.length) continue;
        const dst = this.destinoEn(i);
        if (dst) for (const x of this.enlaceDe(c, c.salida)) dst.push(x);
        c.salida = [];
      }
      return raiz;
    }

    cerrarCampo() {
      const c = this.campos.pop();
      if (!c) return;
      const tipo = c.tipo || this.tipoCampo(c);
      const padre = this.destinoEn(this.campos.length);
      if (!padre) return;
      if (tipo === 'enlace') { for (const x of this.enlaceDe(c, c.salida)) padre.push(x); }
      else if (tipo === 'campo') padre.push({ campo: this.tipoCampoNombre(c), resultado: '1', r: c.f.r, estilo: c.f.estilo });
      // los demás campos ya dejaron su resultado en el párrafo
    }

    // ───────── Imágenes en línea ─────────
    imagen(desp) {
      const dt = this.d.datos;
      if (desp + 68 > dt.length) { this.d.aviso('imagen-falta'); return null; }
      const lcb = u32(dt, desp), cbHeader = u16(dt, desp + 4);
      if (cbHeader < 0x44 || lcb < cbHeader || desp + lcb > dt.length + 16) { this.d.aviso('imagen-falta'); return null; }
      const dxaGoal = i16(dt, desp + 28), dyaGoal = i16(dt, desp + 30), mx = u16(dt, desp + 32), my = u16(dt, desp + 34);
      let p = desp + cbHeader;
      const fin = Math.min(dt.length, desp + lcb);
      // mm 0x66 (MM_SHAPEFILE) = imagen vinculada a un archivo (viene un nombre antes de los datos)
      if (u16(dt, desp + 6) === 0x66) p += 1 + u8(dt, p);
      let intentos = 0;
      while (p + 8 <= fin && intentos++ < 64) {
        const ver = u16(dt, p), tipo = u16(dt, p + 2), largo = u32(dt, p + 4);
        const recVer = ver & 0xF, inst = ver >> 4;
        if (tipo >= 0xF000 && tipo <= 0xF00F && recVer === 0xF) { p += 8; continue; }   // contenedores: se entra
        if (tipo === 0xF007) { p += 8 + 36; continue; }                                  // FBSE: 36 bytes de cabecera y luego el blip
        if (tipo >= 0xF018 && tipo <= 0xF117) {
          const b = dt.subarray(p + 8, Math.min(fin, p + 8 + largo));
          let o = 16;
          const dosUid = tipo === 0xF01E ? inst === 0x6E1 : tipo === 0xF01D ? inst === 0x46B || inst === 0x6E3 : tipo === 0xF01F ? inst === 0x7A9 : false;
          if (dosUid) o += 16;
          let datos = null, tipoImg = null;
          if (tipo === 0xF01E) { datos = b.subarray(o + 1); tipoImg = 'png'; }
          else if (tipo === 0xF01D || tipo === 0xF02A) { datos = b.subarray(o + 1); tipoImg = 'jpeg'; }
          else if (tipo === 0xF01F) { datos = null; this.d.aviso('imagen-formato'); }
          else { this.d.aviso('imagen-formato'); }
          if (!datos || !datos.length) return null;
          this.bytesImagenes += datos.length;
          if (this.bytesImagenes > MAX_IMAGENES_BYTES) throw new Danado('Las imágenes del documento ocupan demasiado y no se puede leer.');
          return { datos: new Uint8Array(datos), tipo: tipoImg, ancho: acotar((dxaGoal * mx / 1000) / 20, 1, 1584) || 72, alto: acotar((dyaGoal * my / 1000) / 20, 1, 1584) || 72 };
        }
        p += 8 + largo;
      }
      this.d.aviso('imagen-falta');
      return null;
    }

    // ───────── Párrafos a bloques ─────────
    bloqueParrafo(rec, info, tabla) {
      const d = this.d;
      const bloque = { tipo: 'parrafo', hijos: [] };
      const id = this.idEstilo(info.istd);
      if (id && id !== 'Normal') bloque.estilo = id;
      else if (id === 'Normal') bloque.estilo = 'Normal';
      const p = Object.assign({}, info.p);
      if (info.extra.tabs) p.tabs = info.extra.tabs.lista;
      delete p.estilo;
      if (info.ilfo && info.ilfo > 0) { const lid = this.listaDe(info.ilfo); if (lid) p.lista = { id: lid, nivel: Math.min(info.ilvl, 8) }; }
      else if (info.ilfo === 0 && info.extra.ilfo === 0) p.lista = { id: null };
      if (Object.keys(p).length) bloque.p = p;
      this._raiz = bloque.hijos;
      const hijos = this.runsDe(rec.cp0, rec.cp1, info.istd);
      for (const h of hijos) delete h._k;
      bloque.hijos = hijos;
      void tabla; void d;
      return bloque;
    }

    // Convierte un tramo de texto en bloques (párrafos y tablas)
    bloquesDe(cp0, cp1, finSecciones) {
      const recs = this.parrafosDe(cp0, cp1, finSecciones);
      const raiz = [];
      const pila = [];   // tablas abiertas, de la más externa a la más interna
      const nuevaTabla = () => ({ filas: [], celdas: [], bloques: [] });
      const cerrar = (hasta) => {
        while (pila.length > hasta) {
          const t = pila.pop();
          // una fila sin marca de fin (archivo raro o dañado): sus celdas se llevan igual, como una fila más, para no perder el texto
          if (t.bloques.length) t.celdas.push({ hijos: t.bloques });
          if (t.celdas.length) t.filas.push({ celdas: t.celdas, tap: {}, bytes: null });
          const tabla = this.tablaFinal(t);
          const padre = pila.length ? pila[pila.length - 1].bloques : raiz;
          if (tabla) padre.push(tabla);
        }
      };
      for (const rec of recs) {
        const info = this.propiedadesParrafo(rec);
        const itap = info.enTabla ? Math.max(1, Math.min(info.extra.itap || 1, MAX_PROFUNDIDAD_TABLAS)) : 0;
        if (itap < pila.length) cerrar(itap);
        if (itap === 0) {
          const b = this.bloqueParrafo(rec, info);
          if (rec.marca === 12 && rec.cp1 >= rec.cp0) { /* marca de fin de sección: no es un salto de página extra */ }
          raiz.push(b);
          continue;
        }
        while (pila.length < itap) pila.push(nuevaTabla());
        const t = pila[pila.length - 1];
        if (info.extra.ttp || info.extra.ttpInterna) {
          // fin de fila: trae las propiedades de la tabla
          t.filas.push({ celdas: t.celdas, tap: info.extra.tap, bytes: info.extra.tapBytes });
          t.celdas = []; t.bloques = [];
          continue;
        }
        const b = this.bloqueParrafo(rec, info);
        t.bloques.push(b);
        if (rec.marca === 7 || info.extra.celdaInterna) { t.celdas.push({ hijos: t.bloques }); t.bloques = []; }
      }
      cerrar(0);
      return raiz;
    }

    tablaFinal(t) {
      const filas = [];
      let anchos = null, ultimoBt = null;
      for (const f of t.filas) {
        const tap = f.tap || {};
        let def = null;
        if (tap.def) {
          const b = f.bytes, q = tap.def.desde;
          const n = u8(b, q);
          if (n > 0 && n <= 63 && q + 1 + 2 * (n + 1) <= b.length) {
            const centros = [];
            for (let k = 0; k <= n; k++) centros.push(i16(b, q + 1 + 2 * k));
            const tcs = [];
            const base = q + 1 + 2 * (n + 1);
            for (let k = 0; k < n; k++) {
              const o = base + 20 * k;
              if (o + 20 > q + tap.def.largo || o + 20 > b.length) { tcs.push(null); continue; }
              const fl = u16(b, o);
              tcs.push({ horz: fl & 3, vert: (fl >> 5) & 3, alin: (fl >> 7) & 3, bordes: { arriba: brc80(b, o + 4), izquierda: brc80(b, o + 8), abajo: brc80(b, o + 12), derecha: brc80(b, o + 16) } });
            }
            def = { n, centros, tcs };
          }
        }
        // relleno por celda
        const rellenos = [];
        const shd = (s, desde) => {
          if (!s) return;
          for (let k = 0; k * 10 + 10 <= s.n && k < 64; k++) {
            const cvFore = colorRef(s.b, s.q + 10 * k), cvBack = colorRef(s.b, s.q + 10 * k + 4), ipat = u16(s.b, s.q + 10 * k + 8);
            rellenos[desde + k] = ipat === 1 ? cvFore : (cvBack || (ipat > 1 ? cvFore : null));
          }
        };
        shd(tap.shd1, 0); shd(tap.shd2, 22); shd(tap.shd3, 44);
        if (tap.shd80 && !tap.shd1) {
          // sombreado de Word 97: dos bytes por celda (color de primer plano y de fondo en el índice de 16 colores, y la trama)
          for (let k = 0; k * 2 + 2 <= tap.shd80.n && k < 64; k++) {
            const w = u16(tap.shd80.b, tap.shd80.q + 2 * k), fore = w & 0x1F, back = (w >> 5) & 0x1F, ipat = (w >> 10) & 0x3F;
            rellenos[k] = ipat === 0 ? (back ? ICO[back] : null) : ipat === 1 ? (fore ? ICO[fore] : null) : (back ? ICO[back] : null);
          }
        }
        // bordes de toda la tabla: arriba, izquierda, abajo, derecha, horizontales y verticales de dentro
        let bt = null;
        if (tap.bordes) {
          const w = tap.bordes.ancho;
          const leer = (k) => (w === 8 ? brc8(tap.bordes.b, tap.bordes.q + 8 * k) : brc80(tap.bordes.b, tap.bordes.q + 4 * k));
          if (tap.bordes.n >= 6 * w) bt = { arriba: leer(0), izquierda: leer(1), abajo: leer(2), derecha: leer(3), h: leer(4), v: leer(5) };
        }
        if (bt) ultimoBt = bt;
        const celdas = [];
        if (def) {
          let col = 0;   // las celdas combinadas con la anterior no tienen marca de celda propia: col solo avanza con las que se ven
          for (let k = 0; k < def.n; k++) {
            const tc = def.tcs[k];
            const ancho = Math.max(0, def.centros[k + 1] - def.centros[k]);
            if (tc && tc.horz === 2 && celdas.length) {
              const prev = celdas[celdas.length - 1];
              prev.fusionH = (prev.fusionH || 1) + 1;
              prev.ancho += ancho;
              prev._fin = def.centros[k + 1];
              continue;
            }
            if (!f.celdas[col]) break;
            const celda = { hijos: f.celdas[col].hijos, ancho, _ini: def.centros[k], _fin: def.centros[k + 1] };
            // vertMerge (2 bits): 3 = primera de las combinadas (fVertMerge y fVertRestart), 1 = sigue a la de arriba
            if (tc && (tc.vert === 3 || tc.vert === 2)) celda.fusionV = 'inicio'; else if (tc && tc.vert === 1) celda.fusionV = 'continua';
            if (rellenos[k]) celda.relleno = rellenos[k];
            celda._k = k;
            if (tc) {
              const bd = {};
              for (const l of ['arriba', 'izquierda', 'abajo', 'derecha']) if (tc.bordes[l]) bd[l] = tc.bordes[l];
              if (Object.keys(bd).length) celda.bordes = bd;
              if (tc.alin === 1) celda.alineacionV = 'centro'; else if (tc.alin === 2) celda.alineacionV = 'abajo';
            }
            celdas.push(celda);
            col++;
          }
          if (!anchos) anchos = def.centros.slice(1).map((c, k) => Math.max(0, c - def.centros[k]));
          if (!celdas.length) continue;
        } else {
          for (const c of f.celdas) celdas.push({ hijos: c.hijos, ancho: 1440 });
          if (!celdas.length) continue;
        }
        for (const c of celdas) if (!c.hijos.length) c.hijos.push({ tipo: 'parrafo', hijos: [] });
        const fila = { celdas, _bt: bt || ultimoBt };
        if (tap.alto && tap.alto > 0) fila.alto = tap.alto;
        if (tap.encabezado) fila.encabezado = true;
        filas.push(fila);
      }
      if (!filas.length) return null;
      // Filas con distintos límites de celda (una celda dividida, otra más ancha): la rejilla es la unión de todos los límites y cada celda abarca las columnas
      // que cubre, como hace Word al guardar .docx. Si las filas no empiezan en el mismo sitio, tienen huecos o pasan de 63 columnas, queda la de la primera fila.
      const limites = new Set();
      let rejillaValida = true;
      for (const fila of filas) {
        let fin = null;
        for (const c of fila.celdas) {
          if (c._ini == null || c._fin <= c._ini || (fin != null && c._ini !== fin)) { rejillaValida = false; break; }
          limites.add(c._ini); limites.add(c._fin); fin = c._fin;
        }
        if (!rejillaValida || limites.size > 64) { rejillaValida = false; break; }
      }
      if (rejillaValida && limites.size >= 2) {
        const rejilla = [...limites].sort((a, b) => a - b);
        if (filas.every((fila) => fila.celdas[0]._ini === rejilla[0])) {
          const indice = new Map(rejilla.map((x, i) => [x, i]));
          for (const fila of filas) for (const c of fila.celdas) { const n = indice.get(c._fin) - indice.get(c._ini); if (n > 1) c.fusionH = n; else delete c.fusionH; }
          anchos = rejilla.slice(1).map((x, k) => x - rejilla[k]);
        }
      }
      // Los bordes de la tabla (arriba, izquierda, abajo, derecha, y los de dentro) pasan a cada celda que no trae los suyos
      filas.forEach((fila, i) => {
        const bt = fila._bt;
        delete fila._bt;
        fila.celdas.forEach((c, j) => {
          delete c._k; delete c._ini; delete c._fin;
          if (!bt) return;
          const propio = c.bordes || {};
          const bd = {
            arriba: propio.arriba || (i === 0 ? bt.arriba : bt.h),
            abajo: propio.abajo || (i === filas.length - 1 ? bt.abajo : bt.h),
            izquierda: propio.izquierda || (j === 0 ? bt.izquierda : bt.v),
            derecha: propio.derecha || (j === fila.celdas.length - 1 ? bt.derecha : bt.v),
          };
          for (const l of Object.keys(bd)) if (!bd[l]) delete bd[l];
          if (Object.keys(bd).length) c.bordes = bd; else delete c.bordes;
        });
      });
      return { tipo: 'tabla', anchos: anchos || filas[0].celdas.map((c) => c.ancho), bordes: false, filas };
    }

    // ───────── Secciones, encabezados y pies ─────────
    secciones() {
      const d = this.d;
      const [fc, lcb] = d.par(6);
      const t = d.enTabla(fc, lcb);
      const n = Math.floor((t.length - 4) / 16);
      const cps = [], sepx = [];
      for (let i = 0; i < n && i < 4096; i++) { cps.push(u32(t, 4 * i)); sepx.push(i32(t, 4 * (n + 1) + 12 * i + 2)); }
      cps.push(n > 0 ? u32(t, 4 * n) : 0);
      const fines = new Set();
      for (let i = 1; i < cps.length; i++) if (cps[i] - 1 >= 0) fines.add(cps[i] - 1);
      const fin = d.ccp.texto;
      if (!n) { cps.length = 0; cps.push(0, fin); sepx.push(-1); }
      // headers
      const [fh, lh] = d.par(11);
      const th = d.enTabla(fh, lh);
      const nh = Math.floor(th.length / 4);
      const baseHdd = d.ccp.texto + d.ccp.ftn;
      const historia = (indice) => {
        if (indice + 1 >= nh) return null;
        const a = u32(th, 4 * indice), b = u32(th, 4 * (indice + 1));
        if (b <= a + 1) return null;
        return [baseHdd + a, baseHdd + b];
      };
      const secs = [];
      let ultimoEnc = null, ultimoPie = null;
      for (let i = 0; i < cps.length - 1; i++) {
        const sep = { ancho: 12240, alto: 15840, izq: 1800, der: 1800, arriba: 1440, abajo: 1440, enc: 720, pie: 720, tipo: 2, horizontal: false, cols: 1, colEsp: 720 };
        if (sepx[i] >= 0 && sepx[i] + 2 <= d.wd.length) {
          const cb = u16(d.wd, sepx[i]);
          const g = d.wd.subarray(sepx[i] + 2, Math.min(d.wd.length, sepx[i] + 2 + cb));
          sprms(g, (op, b, p) => {
            switch (op) {
              case 0x3009: sep.tipo = u8(b, p); break;
              case 0xB01F: sep.ancho = u16(b, p); break;
              case 0xB020: sep.alto = u16(b, p); break;
              case 0xB021: sep.izq = u16(b, p); break;
              case 0xB022: sep.der = u16(b, p); break;
              case 0x9023: sep.arriba = i16(b, p); break;
              case 0x9024: sep.abajo = i16(b, p); break;
              case 0xB017: sep.enc = u16(b, p); break;
              case 0xB018: sep.pie = u16(b, p); break;
              case 0x301D: sep.horizontal = u8(b, p) === 2; break;
              case 0x500B: sep.cols = u16(b, p) + 1; break;
              case 0x900C: sep.colEsp = u16(b, p); break;
              default: break;
            }
          });
        }
        const sec = {
          pagina: { ancho: acotar(sep.ancho, 720, 31680), alto: acotar(sep.alto, 720, 31680), horizontal: sep.horizontal },
          margenes: { arriba: Math.abs(sep.arriba), abajo: Math.abs(sep.abajo), izquierda: sep.izq, derecha: sep.der, encabezado: sep.enc, pie: sep.pie },
          tipo: sep.tipo === 0 ? 'continua' : 'siguiente',
          hijos: [],
        };
        if (sep.cols > 1) sec.columnas = { n: Math.min(sep.cols, 45), espacio: sep.colEsp };
        const cp0 = cps[i], cp1 = i === cps.length - 2 ? fin : cps[i + 1];
        sec.hijos = this.bloquesDe(cp0, Math.min(cp1, fin), fines);
        const base = 6 + 6 * i;
        const he = historia(base + 1), hp = historia(base + 3);
        // los campos abiertos del texto principal no siguen en el encabezado ni en el pie (ni al revés)
        const abiertos = this.campos, deMas = this.camposDeMas;
        this.campos = []; this.camposDeMas = 0;
        if (he) ultimoEnc = this.bloquesDe(he[0], he[1], null);
        if (hp) { this.campos = []; this.camposDeMas = 0; ultimoPie = this.bloquesDe(hp[0], hp[1], null); }
        this.campos = abiertos; this.camposDeMas = deMas;
        if (ultimoEnc && ultimoEnc.length) sec.encabezado = ultimoEnc;
        if (ultimoPie && ultimoPie.length) sec.pie = ultimoPie;
        secs.push(sec);
      }
      return secs;
    }

    convertir() {
      this.estilosDelContrato();
      this.contrato.secciones = this.secciones();
      if (this.d.ccp.ftn > 1) this.d.aviso('notas');
      if (this.d.ccp.txbx > 1 || this.d.ccp.hdrTxbx > 1) this.d.aviso('dibujos');
      return this.contrato;
    }
  }

  function parece(bytes) {
    try {
      const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
      if (b.length < 512 || b[0] !== 0xD0 || b[1] !== 0xCF || b[2] !== 0x11 || b[3] !== 0xE0) return false;
      const cfb = CfbPropio.leer(b);
      const e = cfb.entradas.find((x) => x.tipo === 2 && x.nombre === 'WordDocument');
      if (!e) return false;
      const wd = cfb.flujo(e);
      return wd.length > 0x20 && u16(wd, 0) === 0xA5EC;
    } catch (e) { return false; }
  }

  async function leer(bytes) {
    const datos = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const doc = new Documento(datos);
    const conv = new Conversor(doc);
    const contrato = conv.convertir();
    return { contrato, mensajes: [...doc.mensajes] };
  }

  async function aDocx(bytes) {
    const { contrato, mensajes } = await leer(bytes);
    const r = await DocxDesdeContrato.generar(contrato);
    return { bytes: r.bytes, mensajes: [...new Set([...mensajes, ...r.mensajes])] };
  }

  globalThis.DocLector = { parece, leer, aDocx, _interno: { Documento, Conversor, leerPap, leerChp } };
})();

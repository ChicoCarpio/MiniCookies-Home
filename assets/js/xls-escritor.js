/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Escritor de libros de Excel 97-2003 (.xls) del motor propio de formatos antiguos (fase 3). Arma el flujo "Workbook" de registros BIFF8 ([MS-XLS]) y lo mete en un
// contenedor CFB ([MS-CFB]). Requiere cfb-propio.js, xlsx-estilos.js (formatos integrados y paleta) y, para llevar fórmulas, xls-formulas.js. No depende del navegador.
//
//   const { bytes, avisos } = XlsEscritor.crearLibro(hojas, opciones);
//
// `hojas` es lo mismo que recibe XlsxEscritor.crearLibro (y lo que devuelven los lectores): [{ nombre, filas: [[valor…]…], anchos?: [caracteres…], combinadas?: ['A1:B2'…],
// estado?: 'visible' | 'oculta' | 'muyOculta', columnasOcultas?: [columna…], filasOcultas?: [fila…], altosFila?: { fila: puntos } }]. valor: número, texto, lógico, Date, null,
// { error: '#N/A' } o { valor, formato?, estilo?, formula?, matriz? } (estilo: ver XlsxEscritor.crearLibro; formula: texto sin '=' como lo escribe un .xlsx; matriz: 'B3:B5',
// el rango de una fórmula de matriz, en su primera celda). opciones: { formulas: true (false: solo valores), fecha1904: bool, idiomaComplemento: 'en' | 'es' (nombre de las
// funciones del Analysis ToolPak, por defecto 'es'), nombres: [{ nombre, formula, hoja? }] (nombres definidos; hoja: índice de la hoja si es local) }.
// `avisos`: lo que el formato .xls no puede llevar: [{ tipo, cantidad, ejemplo? }] con tipo 'fuera' (celdas pasadas de 65.536 × 256), 'formula-sin-llevar' (quedó el valor),
// 'futura' (función posterior a Excel 97), 'estilos' (más de 4.000 estilos distintos), 'texto-largo' (más de 32.767 caracteres), 'hojas' (más de 255 hojas).
(function () {
  const MAX_FILAS = 65536, MAX_COLS = 256, MAX_TEXTO = 32767, MAX_HOJAS = 255, MAX_XF = 4000, MAX_REGISTRO = 8224;
  const CODIGOS_ERROR = { '#NULL!': 0x00, '#DIV/0!': 0x07, '#VALUE!': 0x0F, '#REF!': 0x17, '#NAME?': 0x1D, '#NUM!': 0x24, '#N/A': 0x2A, '#GETTING_DATA': 0x2B };
  const BORDES = { thin: 1, medium: 2, dashed: 3, dotted: 4, thick: 5, double: 6, hair: 7, mediumDashed: 8, dashDot: 9, mediumDashDot: 10, dashDotDot: 11, mediumDashDotDot: 12, slantDashDot: 13 };
  const TRAMAS = { solid: 1, mediumGray: 2, darkGray: 3, lightGray: 4, darkHorizontal: 5, darkVertical: 6, darkDown: 7, darkUp: 8, darkGrid: 9, darkTrellis: 10, lightHorizontal: 11, lightVertical: 12, lightDown: 13, lightUp: 14, lightGrid: 15, lightTrellis: 16, gray125: 17, gray0625: 18 };
  const ALIN_H = { general: 0, left: 1, center: 2, right: 3, fill: 4, justify: 5, centerContinuous: 6, distributed: 7 };
  const ALIN_V = { top: 0, center: 1, bottom: 2, justify: 3, distributed: 4 };

  // ---- bytes -----------------------------------------------------------------------------------------------------
  class Buf {
    constructor() { this.b = new Uint8Array(1024); this.n = 0; this.dv = new DataView(this.b.buffer); }
    _a(k) {
      if (this.n + k <= this.b.length) return;
      const m = new Uint8Array(Math.max(this.b.length * 2, this.n + k));
      m.set(this.b.subarray(0, this.n)); this.b = m; this.dv = new DataView(m.buffer);
    }
    u8(v) { this._a(1); this.b[this.n++] = v & 255; return this; }
    u16(v) { this._a(2); this.dv.setUint16(this.n, v & 0xFFFF, true); this.n += 2; return this; }
    u32(v) { this._a(4); this.dv.setUint32(this.n, v >>> 0, true); this.n += 4; return this; }
    f64(v) { this._a(8); this.dv.setFloat64(this.n, v, true); this.n += 8; return this; }
    bytes(a) { this._a(a.length); this.b.set(a, this.n); this.n += a.length; return this; }
    fin() { return this.b.subarray(0, this.n); }
  }
  // Un registro: id, longitud y datos; si pasa de 8224 bytes sigue en registros CONTINUE (solo lo usan los que lo admiten).
  function registro(out, id, datos) {
    const d = datos instanceof Uint8Array ? datos : Uint8Array.from(datos);
    if (d.length > MAX_REGISTRO) throw new Error('registro 0x' + id.toString(16) + ' de ' + d.length + ' bytes');
    out.u16(id).u16(d.length).bytes(d);
  }
  const reg = (id, ...v) => { const b = new Buf(); for (const x of v) { if (typeof x === 'number') b.u16(x); else b.bytes(x); } return { id, d: b.fin() }; };

  // Cadena Unicode de BIFF8: cch (1 o 2 bytes), 0 = Latin-1 comprimida o 1 = UTF-16, caracteres.
  function cadena(b, s, cchBytes) {
    let ancho = false;
    for (let i = 0; i < s.length; i++) if (s.charCodeAt(i) > 0xFF) { ancho = true; break; }
    if (cchBytes === 1) b.u8(s.length); else b.u16(s.length);
    b.u8(ancho ? 1 : 0);
    for (let i = 0; i < s.length; i++) { if (ancho) b.u16(s.charCodeAt(i)); else b.u8(s.charCodeAt(i)); }
  }

  // ---- utilidades ------------------------------------------------------------------------------------------------
  const colANumero = (s) => { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
  function parseRango(t) {   // 'A1:B2' o 'A1' → { f0, c0, f1, c1 } desde 0, o null
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(String(t).trim());
    if (!m) return null;
    const c0 = colANumero(m[1]), f0 = +m[2] - 1, c1 = m[3] ? colANumero(m[3]) : c0, f1 = m[4] ? +m[4] - 1 : f0;
    return f0 < 0 || f1 < f0 || c1 < c0 ? null : { f0, c0, f1, c1 };
  }
  function nombreHojaValido(n, usados, k) {
    let t = String(n == null ? '' : n).replace(/[\[\]:*?\/\\]/g, '_').replace(/^'+|'+$/g, '').slice(0, 31);
    if (!t) t = 'Hoja' + (k + 1);
    let base = t, i = 2;
    while (usados.has(t.toLowerCase())) { const suf = '(' + i++ + ')'; t = base.slice(0, 31 - suf.length) + suf; }
    usados.add(t.toLowerCase());
    return t;
  }
  // Las columnas que tienen algo en una fila, en orden. Una fila ancha y rala (una celda en XFD) es un arreglo disperso: se recorren sus índices, no sus 16.384 lugares.
  function presentes(fila) {
    const r = [];
    if (fila.length <= 1024) { for (let c = 0; c < fila.length; c++) if (fila[c] !== undefined) r.push(c); return r; }
    for (const k in fila) { const c = +k; if (Number.isInteger(c) && c >= 0 && fila[c] !== undefined) r.push(c); }
    return r;
  }
  const MAX_EXPS = 1000000;   // celdas de fórmulas de matriz (cada una lleva su registro): una matriz A1:IV65536 son 16 millones
  const hexRgb = (h) => [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
  const colorValido = (c) => typeof c === 'string' && /^[0-9A-Fa-f]{6}$/.test(c) ? c.toUpperCase() : null;

  // ---- estilos ---------------------------------------------------------------------------------------------------
  // Se arma primero una descripción neutra de cada estilo (con los colores como 'RRGGBB'); los registros salen al final, cuando ya se sabe la paleta.
  function nuevoRegistroEstilos(avisar) {
    const fuentes = [], mapaFuentes = new Map(), xfs = [], mapaXf = new Map(), cacheObj = new WeakMap(), formatos = [], mapaFormatos = new Map();
    const usoColor = new Map();
    const color = (c) => { const h = colorValido(c); if (h) usoColor.set(h, (usoColor.get(h) || 0) + 1); return h; };
    const inversoIntegrados = new Map();
    for (const [id, cod] of Object.entries(XlsxEstilos.FORMATOS_INTEGRADOS)) if (!inversoIntegrados.has(cod)) inversoIntegrados.set(cod, +id);

    function fuenteIdx(f) {
      const d = {
        nombre: (f && typeof f.nombre === 'string' && f.nombre ? f.nombre : 'Calibri').slice(0, 255),
        tam: f && Number.isFinite(f.tam) && f.tam > 0 ? Math.min(409, f.tam) : 11,
        negrita: !!(f && f.negrita), cursiva: !!(f && f.cursiva), tachado: !!(f && f.tachado),
        sub: f && f.subrayado === 'double' ? 2 : f && f.subrayado ? 1 : 0, color: f ? color(f.color) : null,
      };
      const k = JSON.stringify(d);
      if (mapaFuentes.has(k)) return mapaFuentes.get(k);
      fuentes.push(d); mapaFuentes.set(k, fuentes.length - 1);
      return fuentes.length - 1;
    }
    fuenteIdx(null);   // la 0: Calibri 11 (las cuatro primeras son la misma)

    function formatoId(fmt) {
      if (fmt === undefined || fmt === null) return 0;
      if (typeof fmt === 'number') return Number.isInteger(fmt) && fmt >= 0 && fmt < 50 && XlsxEstilos.FORMATOS_INTEGRADOS[fmt] !== undefined ? fmt : 0;
      const cod = String(fmt);
      if (cod === 'General') return 0;
      if (inversoIntegrados.has(cod)) return inversoIntegrados.get(cod);
      if (mapaFormatos.has(cod)) return mapaFormatos.get(cod);
      if (formatos.length >= 200 || cod.length > 250) return 0;
      formatos.push(cod); mapaFormatos.set(cod, 164 + formatos.length - 1);
      return 164 + formatos.length - 1;
    }

    function xfDe(formato, estilo) {
      let porFormato = estilo && typeof estilo === 'object' ? cacheObj.get(estilo) : null;
      if (estilo && typeof estilo === 'object' && !porFormato) { porFormato = new Map(); cacheObj.set(estilo, porFormato); }
      if (porFormato && porFormato.has(formato)) return porFormato.get(formato);
      const est = estilo && typeof estilo === 'object' ? estilo : {};
      const r = typeof est.relleno === 'string' ? { patron: 'solid', fg: est.relleno } : est.relleno && typeof est.relleno === 'object' ? est.relleno : null;
      const b = est.bordes || {}, a = est.alinea || {};
      const lado = (x) => x && BORDES[x.e] ? { e: BORDES[x.e], c: color(x.c) } : { e: 0, c: null };
      const d = {
        fuente: est.fuente ? fuenteIdx(est.fuente) : 0, fmt: formatoId(formato),
        fls: r && TRAMAS[r.patron] ? TRAMAS[r.patron] : 0, fg: r ? color(r.fg) : null, bg: r ? color(r.bg) : null,
        izq: lado(b.izq), der: lado(b.der), arr: lado(b.arr), aba: lado(b.aba),
        diag: lado(b.diagA && BORDES[b.diagA.e] ? b.diagA : b.diagB), diagBits: (b.diagA && BORDES[b.diagA.e] ? 1 : 0) | (b.diagB && BORDES[b.diagB.e] ? 2 : 0),
        h: ALIN_H[a.h] || 0, v: a.v in ALIN_V ? ALIN_V[a.v] : 2, ajustar: !!a.ajustar, sangria: Number.isInteger(a.sangria) ? Math.max(0, Math.min(15, a.sangria)) : 0,
        rot: Number.isInteger(a.rotacion) && ((a.rotacion > 0 && a.rotacion <= 180) || a.rotacion === 255) ? a.rotacion : 0,
      };
      const k = JSON.stringify(d);
      let idx;
      if (mapaXf.has(k)) idx = mapaXf.get(k);
      else if (xfs.length >= MAX_XF) { avisar('estilos'); idx = 15; }
      else { xfs.push(d); idx = 16 + xfs.length - 1; mapaXf.set(k, idx); }
      if (porFormato) porFormato.set(formato, idx);
      return idx;
    }
    return { fuentes, xfs, formatos, usoColor, xfDe };
  }

  // La paleta: los colores exactos que tiene la estándar usan su índice; el resto ocupa los lugares que ningún color usado necesita. Si faltan lugares, el más cercano.
  function armarPaleta(usoColor) {
    const base = XlsxEstilos.PALETA_INDEXADA.map((c) => c.toUpperCase());
    const paleta = base.slice(0, 64), indice = new Map();
    const usados = [...usoColor.entries()].sort((x, y) => y[1] - x[1]).map((e) => e[0]);
    const ocupados = new Set();
    for (const c of usados) { const i = paleta.indexOf(c, 8); if (i >= 8 && i < 64) { indice.set(c, i); ocupados.add(i); } }
    const libres = []; for (let i = 63; i >= 8; i--) if (!ocupados.has(i)) libres.push(i);
    let cambio = false;
    const distancia = (x, y) => { const p = hexRgb(x), q = hexRgb(y); return (p[0] - q[0]) ** 2 + (p[1] - q[1]) ** 2 + (p[2] - q[2]) ** 2; };
    for (const c of usados) {
      if (indice.has(c)) continue;
      if (libres.length) { const i = libres.pop(); paleta[i] = c; indice.set(c, i); cambio = true; continue; }
      let mejor = 8, md = Infinity;
      for (let i = 8; i < 64; i++) { const dd = distancia(c, paleta[i]); if (dd < md) { md = dd; mejor = i; } }
      indice.set(c, mejor);
    }
    return { paleta, indice, cambio };
  }

  function fuenteBytes(f, idxColor) {
    const b = new Buf();
    b.u16(Math.round(f.tam * 20)).u16((f.cursiva ? 2 : 0) | (f.tachado ? 8 : 0)).u16(f.color ? idxColor(f.color) : 0x7FFF).u16(f.negrita ? 700 : 400).u16(0).u8(f.sub).u8(2).u8(0).u8(0);
    cadena(b, f.nombre, 1);
    return b.fin();
  }
  function xfBytes(x, idxColor) {
    const c = (v) => v ? idxColor(v) : 64;
    const b = new Buf();
    const usado = (x.fmt ? 0x04 : 0) | (x.fuente ? 0x08 : 0) | (x.h || x.v !== 2 || x.ajustar || x.sangria || x.rot ? 0x10 : 0) |
      (x.izq.e || x.der.e || x.arr.e || x.aba.e || x.diagBits ? 0x20 : 0) | (x.fls ? 0x40 : 0);
    b.u16(x.fuente === 0 ? 0 : x.fuente + 4).u16(x.fmt).u16(0x0001);
    b.u8(x.h | (x.ajustar ? 8 : 0) | (x.v << 4)).u8(x.rot).u8(x.sangria).u8(usado);
    b.u16(x.izq.e | (x.der.e << 4) | (x.arr.e << 8) | (x.aba.e << 12));
    b.u16(c(x.izq.c) | (c(x.der.c) << 7) | (x.diagBits << 14));
    b.u32(c(x.arr.c) | (c(x.aba.c) << 7) | ((x.diagBits ? c(x.diag.c) : 0) << 14) | ((x.diagBits ? x.diag.e : 0) << 21) | ((x.fls << 26) >>> 0));
    b.u16((x.fls ? (x.fg ? idxColor(x.fg) : 64) : 64) | ((x.fls ? (x.bg ? idxColor(x.bg) : 65) : 65) << 7));
    return b.fin();
  }

  // Los colores RGB exactos de un XF (XFEXT): la paleta de un .xls solo tiene 56, y esto es lo que Excel usa si el color no cabe. null si el XF no lleva colores.
  function xfExtBytes(ixfe, x, colorFuente) {
    const props = [];
    const color = (tipo, hex) => { if (hex) props.push([tipo, hexRgb(hex)]); };
    if (x.fls) { color(4, x.fg); color(5, x.bg); }
    color(7, x.arr.c); color(8, x.aba.c); color(9, x.izq.c); color(10, x.der.c);
    if (x.diagBits) color(11, x.diag.c);
    color(13, colorFuente);
    if (!props.length) return null;
    const b = new Buf().u16(0x087D).u16(0).u32(0).u32(0).u16(0).u16(ixfe).u16(0).u16(props.length);
    for (const [tipo, rgb] of props) b.u16(tipo).u16(20).u16(2).u16(0).u8(rgb[0]).u8(rgb[1]).u8(rgb[2]).u8(0xFF).u32(0).u32(0);
    return b.fin();
  }

  // ---- cadenas compartidas (SST) -----------------------------------------------------------------------------------
  // Los caracteres de una cadena pueden seguir en un registro CONTINUE (que empieza con su propio byte de ancho); la cabecera de la cadena no se parte.
  function escribirSst(out, lista, total) {
    if (!lista.length) return;
    let b = new Buf(), lleno = 0, primero = true;
    const cerrar = () => { out.u16(primero ? 0x00FC : 0x003C).u16(b.n).bytes(b.fin()); primero = false; b = new Buf(); lleno = 0; };
    b.u32(total).u32(lista.length); lleno = 8;
    for (const s of lista) {
      let ancho = false;
      for (let i = 0; i < s.length; i++) if (s.charCodeAt(i) > 0xFF) { ancho = true; break; }
      const unidad = ancho ? 2 : 1;
      if (MAX_REGISTRO - lleno < 3 + (s.length ? unidad : 0)) cerrar();
      b.u16(s.length).u8(ancho ? 1 : 0); lleno += 3;
      let i = 0;
      while (i < s.length) {
        let cabe = Math.floor((MAX_REGISTRO - lleno) / unidad);
        if (cabe <= 0) { cerrar(); b.u8(ancho ? 1 : 0); lleno = 1; cabe = Math.floor((MAX_REGISTRO - lleno) / unidad); }
        const n = Math.min(cabe, s.length - i);
        for (let k = 0; k < n; k++) { const ch = s.charCodeAt(i + k); if (ancho) b.u16(ch); else b.u8(ch); }
        lleno += n * unidad; i += n;
      }
    }
    cerrar();
  }

  // ---- el libro ----------------------------------------------------------------------------------------------------
  function crearLibro(hojasEntrada, opciones) {
    opciones = opciones || {};
    const conFormulas = opciones.formulas !== false && typeof XlsFormulas !== 'undefined';
    const idioma = opciones.idiomaComplemento === 'en' ? 'en' : 'es';
    const avisosMapa = new Map();
    const avisar = (tipo, ejemplo) => { const a = avisosMapa.get(tipo); if (a) a.cantidad++; else avisosMapa.set(tipo, { tipo, cantidad: 1, ...(ejemplo ? { ejemplo } : {}) }); };
    if (!Array.isArray(hojasEntrada) || !hojasEntrada.length) throw new Error('el libro necesita al menos una hoja');
    if (hojasEntrada.length > MAX_HOJAS) { avisar('hojas'); hojasEntrada = hojasEntrada.slice(0, MAX_HOJAS); }

    const usados = new Set();
    const hojas = hojasEntrada.map((h, k) => ({ ...h, nombreOriginal: h.nombre, nombre: nombreHojaValido(h.nombre, usados, k) }));
    const indiceHoja = new Map();
    hojas.forEach((h, k) => { indiceHoja.set(h.nombre.toLowerCase(), k); if (h.nombreOriginal != null && !indiceHoja.has(String(h.nombreOriginal).toLowerCase())) indiceHoja.set(String(h.nombreOriginal).toLowerCase(), k); });

    const estilos = nuevoRegistroEstilos(avisar);
    const cadenas = [], mapaCadenas = new Map();
    let totalCadenas = 0;
    const idCadena = (s) => {
      totalCadenas++;
      if (s.length > MAX_TEXTO) { s = s.slice(0, MAX_TEXTO); avisar('texto-largo'); }
      let i = mapaCadenas.get(s);
      if (i === undefined) { i = cadenas.length; cadenas.push(s); mapaCadenas.set(s, i); }
      return i;
    };

    // tablas de referencias: EXTERNSHEET (hojas), libro de complemento (Analysis ToolPak) y nombres
    const xtis = [], mapaXti = new Map();
    const nombresDef = [];   // { nombre, hoja (itab), formula, funcion }
    const mapaNombres = new Map();
    for (const n of opciones.nombres || []) {
      if (!n || typeof n.nombre !== 'string' || !n.nombre || typeof n.formula !== 'string' || !n.formula) continue;   // sin fórmula no se puede llevar: lo que lo use queda en solo valor
      if (n.nombre.length > 255) continue;   // el NAME guarda el largo en un byte (y Excel no admite más): las fórmulas que lo usan quedan en su valor, con su aviso
      // hoja: el índice (desde 0) o el nombre de la hoja de un nombre local; sin ella, el nombre es del libro
      const ih = typeof n.hoja === 'string' ? (indiceHoja.has(n.hoja.toLowerCase()) ? indiceHoja.get(n.hoja.toLowerCase()) : -2) : Number.isInteger(n.hoja) && n.hoja >= 0 && n.hoja < hojas.length ? n.hoja : -1;
      if (ih === -2) continue;
      const clave = n.nombre.toLowerCase() + '|' + (ih >= 0 ? ih : '');
      if (mapaNombres.has(clave)) continue;
      nombresDef.push({ nombre: n.nombre, hoja: ih + 1, formula: n.formula, funcion: false });
      mapaNombres.set(clave, nombresDef.length);
    }
    const complemento = { nombres: [], xti: -1 };
    let hojaActual = -1;
    const ctxBase = {
      hoja: (n) => { const k = indiceHoja.get(String(n).toLowerCase()); return k === undefined ? -1 : k; },
      ixti: (a, b) => {
        const k = a + ':' + b;
        if (!mapaXti.has(k)) { xtis.push({ libro: 0, primera: a, ultima: b }); mapaXti.set(k, xtis.length - 1); }
        return mapaXti.get(k);
      },
      nombre: (t, esFuncion) => {
        const i = (hojaActual >= 0 && mapaNombres.get(t.toLowerCase() + '|' + hojaActual)) || mapaNombres.get(t.toLowerCase() + '|');
        if (i) return i;
        if (!/^_xl(fn|ws)\./i.test(t)) return null;
        nombresDef.push({ nombre: t, hoja: 0, formula: null, funcion: true });
        mapaNombres.set(t.toLowerCase() + '|', nombresDef.length);
        return nombresDef.length;
      },
      complemento: (d) => {
        const nombre = idioma === 'es' && d.local ? d.local : d.nombre;
        let k = complemento.nombres.indexOf(nombre);
        if (k < 0) { complemento.nombres.push(nombre); k = complemento.nombres.length - 1; }
        if (complemento.xti < 0) { xtis.push({ libro: 1, primera: -2, ultima: -2 }); complemento.xti = xtis.length - 1; }
        return { ixti: complemento.xti, idx: k + 1 };
      },
    };
    // el libro interno es el SUPBOOK 0; el de complemento (si hace falta) el 1. Si el de complemento se crea antes de que haya otro EXTERNSHEET, el orden no importa.
    let ultimoMotivo = '';
    const codificarFormula = (texto, f, c, clase) => {
      const r = XlsFormulas.codificar(texto, ctxBase, { clase });
      if (r.error) { ultimoMotivo = r.error; return null; }
      for (const a of r.avisos) { if (a.startsWith('futura:')) avisar('futura', a.slice(7)); else avisar(a); }
      return r;
    };

    // ---- una hoja → flujo de bytes ----
    const cuerposHoja = hojas.map((h, ih) => construirHoja(h, ih));
    function construirHoja(h, ih) {
      hojaActual = ih;
      const out = new Buf();
      let filas = Array.isArray(h.filas) ? h.filas : [];
      // Una fórmula de matriz de varias celdas: cada celda del rango (menos la primera) lleva su propio registro FORMULA con un PtgExp que apunta a la primera.
      // Una matriz que no entra en el .xls (o en el tope de MAX_EXPS) queda en su valor: va en sinMatriz, por su celda de origen.
      const exps = new Map(), sinMatriz = new Set();
      if (conFormulas) {
        for (let f = 0; f < filas.length && f < MAX_FILAS; f++) {
          const fila = filas[f];
          if (!fila) continue;
          for (const c of presentes(fila)) {
            if (c >= MAX_COLS) break;
            const x = fila[c];
            if (!x || typeof x !== 'object' || !x.formula || !x.matriz) continue;
            const m = parseRango(x.matriz);
            if (!m || (m.f1 === m.f0 && m.c1 === m.c0) || m.f0 !== f || m.c0 !== c) continue;
            if (m.f1 >= MAX_FILAS || m.c1 >= MAX_COLS || exps.size + (m.f1 - m.f0 + 1) * (m.c1 - m.c0 + 1) > MAX_EXPS) { sinMatriz.add(f * MAX_COLS + c); continue; }
            if (filas === h.filas) filas = filas.slice();
            for (let r = m.f0; r <= m.f1; r++) {
              if (!filas[r] || filas[r] === h.filas[r]) filas[r] = Array.from(filas[r] || []);
              for (let k = m.c0; k <= m.c1; k++) if (r !== f || k !== c) exps.set(r * MAX_COLS + k, { f, c });
            }
          }
        }
      }
      const columnasOcultas = new Set(h.columnasOcultas || []), filasOcultas = new Set(h.filasOcultas || []);
      const altos = h.altosFila || {};
      let ultimaFila = -1, ultimaCol = -1;
      const cuerpoCeldas = new Buf();
      const infoFilas = [];   // { f, c0, c1 }
      const nombreCelda = (f, c) => { let s = '', n = c + 1; while (n > 0) { s = String.fromCharCode(65 + (n - 1) % 26) + s; n = Math.floor((n - 1) / 26); } return s + (f + 1); };

      // normaliza una celda a { valor, formato, estilo, formula, matriz }
      const normal = (v) => {
        if (v === undefined || v === null) return { valor: null };
        if (v instanceof Date) { const t = v.getTime(); return Number.isFinite(t) ? { valor: (t - v.getTimezoneOffset() * 60000) / 86400000 + 25569, formato: 14 } : { valor: null }; }
        if (typeof v === 'object' && typeof v.error !== 'string') {
          let val = v.valor;
          if (val instanceof Date) { const t = val.getTime(); val = Number.isFinite(t) ? (t - val.getTimezoneOffset() * 60000) / 86400000 + 25569 : null; }
          let formato = v.formato, estilo = v.estilo;
          if (val && typeof val === 'object' && !(typeof val.error === 'string') && 'valor' in val) { if (formato === undefined) formato = val.formato; if (estilo === undefined) estilo = val.estilo; val = val.valor; }
          return { valor: val === undefined ? null : val, formato, estilo, formula: typeof v.formula === 'string' && v.formula ? v.formula : null, matriz: v.matriz, sinValor: val === undefined };
        }
        return { valor: v };
      };
      // los 8 bytes del resultado de una fórmula, y el texto que va en un registro STRING detrás
      const resultado = (val) => {
        if (typeof val === 'number' && Number.isFinite(val)) return { b: null, num: val };
        if (typeof val === 'boolean') return { b: [1, 0, val ? 1 : 0, 0, 0, 0, 0xFF, 0xFF] };
        // el texto va en un solo registro STRING (8.224 bytes: 8.221 caracteres de 1 byte o 4.110 de 2); es el último resultado, que Excel vuelve a calcular al abrir
        if (typeof val === 'string') { let ancho = false; for (let i = 0; i < val.length && !ancho; i++) if (val.charCodeAt(i) > 0xFF) ancho = true; return { b: [0, 0, 0, 0, 0, 0, 0xFF, 0xFF], texto: val.slice(0, ancho ? (MAX_REGISTRO - 3) >> 1 : MAX_REGISTRO - 3) }; }
        if (val && typeof val === 'object' && typeof val.error === 'string') return { b: [2, 0, CODIGOS_ERROR[val.error] === undefined ? 0x0F : CODIGOS_ERROR[val.error], 0, 0, 0, 0xFF, 0xFF] };
        return { b: null, num: 0, vacio: true };
      };

      for (let f = 0; f < filas.length; f++) {
        const fila = filas[f];
        if (!Array.isArray(fila) && !(fila && fila.length >= 0)) continue;
        if (f >= MAX_FILAS) { if (presentes(fila).some((c) => fila[c] !== null)) avisar('fuera'); continue; }
        let c0 = -1, c1 = -1;
        if (fila.length > MAX_COLS) for (const c of presentes(fila)) if (c >= MAX_COLS) { const n = normal(fila[c]); if (n.valor !== null || n.formula) avisar('fuera'); }
        for (let c = 0; c < fila.length && c < MAX_COLS; c++) {
          const n = normal(fila[c]);
          const origen = exps.get(f * MAX_COLS + c);
          const tieneEstilo = n.estilo !== undefined && n.estilo !== null || (n.formato !== undefined && n.formato !== null && n.formato !== 'General' && n.formato !== 0);
          const xf = tieneEstilo ? estilos.xfDe(n.formato, n.estilo) : 15;
          const v = n.valor;
          if (origen) {
            const res = resultado(n.sinValor ? undefined : v);
            const b = new Buf().u16(f).u16(c).u16(xf);
            if (res.b) b.bytes(res.b); else b.f64(res.num);
            b.u16(res.vacio ? 0x0001 : 0).u32(0).u16(5).u8(0x01).u16(origen.f).u16(origen.c);
            registro(cuerpoCeldas, 0x0006, b.fin());
            if (res.texto !== undefined) { const s = new Buf(); cadena(s, res.texto, 2); registro(cuerpoCeldas, 0x0207, s.fin()); }
          } else if (v === null && !n.formula) {
            if (!tieneEstilo) continue;
            registro(cuerpoCeldas, 0x0201, new Buf().u16(f).u16(c).u16(xf).fin());
          } else {
            let formulaHecha = false;
            if (n.formula && /^#(NULL!|DIV[/]0!|VALUE!|REF!|NAME[?]|NUM!|N[/]A)$/.test(n.formula)) avisar('formula-sin-llevar', nombreCelda(f, c) + ': fórmula que da error en Excel 97-2003');
            else if (n.formula && conFormulas && sinMatriz.has(f * MAX_COLS + c)) avisar('formula-sin-llevar', nombreCelda(f, c) + ': fórmula de matriz demasiado grande para un .xls');
            else if (n.formula && conFormulas) {
              const matriz = n.matriz ? parseRango(n.matriz) : null;
              const r = codificarFormula(n.formula, f, c, matriz ? 'A' : 'V');
              if (r && (r.rgce.length + r.rgcb.length) < 8000) {
                const res = resultado(n.sinValor ? undefined : v);
                const b = new Buf();
                b.u16(f).u16(c).u16(xf);
                if (res.b) b.bytes(res.b); else b.f64(res.num);
                const cuerpo = matriz ? Uint8Array.from([0x01, f & 255, f >> 8, c & 255, c >> 8]) : r.rgce;
                b.u16(res.vacio ? 0x0001 : 0x0000).u32(0).u16(cuerpo.length).bytes(cuerpo);
                if (!matriz && r.rgcb.length) b.bytes(r.rgcb);
                registro(cuerpoCeldas, 0x0006, b.fin());
                if (matriz) {
                  const a = new Buf();
                  a.u16(matriz.f0).u16(Math.min(matriz.f1, MAX_FILAS - 1)).u8(matriz.c0).u8(Math.min(matriz.c1, MAX_COLS - 1)).u16(res.vacio ? 1 : 0).u32(0).u16(r.rgce.length).bytes(r.rgce).bytes(r.rgcb);
                  registro(cuerpoCeldas, 0x0221, a.fin());
                }
                if (res.texto !== undefined) { const s = new Buf(); cadena(s, res.texto, 2); registro(cuerpoCeldas, 0x0207, s.fin()); }   // el STRING va después del ARRAY
                formulaHecha = true;
              } else {
                avisar('formula-sin-llevar', nombreCelda(f, c) + (ultimoMotivo ? ': ' + ultimoMotivo : ''));
                // las celdas de su matriz quedan con su valor (solo se mira el rango de esta fórmula, no todas las matrices)
                if (matriz && matriz.f0 === f && matriz.c0 === c) for (let r = matriz.f0; r <= matriz.f1 && r < MAX_FILAS; r++) for (let k = matriz.c0; k <= matriz.c1 && k < MAX_COLS; k++) { const o = exps.get(r * MAX_COLS + k); if (o && o.f === f && o.c === c) exps.delete(r * MAX_COLS + k); }
              }
            } else if (n.formula) avisar('formula-sin-llevar', nombreCelda(f, c) + (ultimoMotivo ? ': ' + ultimoMotivo : ''));
            if (!formulaHecha) {
              if (typeof v === 'number') {
                if (!Number.isFinite(v)) registro(cuerpoCeldas, 0x0205, new Buf().u16(f).u16(c).u16(xf).u8(0x24).u8(1).fin());
                else if (Number.isInteger(v) && Math.abs(v) < 0x20000000 && !(v === 0 && 1 / v < 0)) registro(cuerpoCeldas, 0x027E, new Buf().u16(f).u16(c).u16(xf).u32(((v << 2) | 2) >>> 0).fin());
                else registro(cuerpoCeldas, 0x0203, new Buf().u16(f).u16(c).u16(xf).f64(v).fin());
              } else if (typeof v === 'boolean') registro(cuerpoCeldas, 0x0205, new Buf().u16(f).u16(c).u16(xf).u8(v ? 1 : 0).u8(0).fin());
              else if (v && typeof v === 'object' && typeof v.error === 'string') registro(cuerpoCeldas, 0x0205, new Buf().u16(f).u16(c).u16(xf).u8(CODIGOS_ERROR[v.error] === undefined ? 0x0F : CODIGOS_ERROR[v.error]).u8(1).fin());
              else if (v === null) registro(cuerpoCeldas, 0x0201, new Buf().u16(f).u16(c).u16(xf).fin());
              else registro(cuerpoCeldas, 0x00FD, new Buf().u16(f).u16(c).u16(xf).u32(idCadena(String(v))).fin());
            }
          }
          if (c0 < 0) c0 = c;
          c1 = c;
        }
        if (c0 >= 0 || altos[f] !== undefined || filasOcultas.has(f)) infoFilas.push({ f, c0: Math.max(c0, 0), c1: Math.max(c1, 0), vacia: c0 < 0 });
        if (c0 >= 0) { ultimaFila = Math.max(ultimaFila, f); ultimaCol = Math.max(ultimaCol, c1); }
      }
      // filas con alto u oculta pero sin celdas
      const conCeldas = new Set(infoFilas.map((i) => i.f));
      for (const k of Object.keys(altos)) { const f = +k; if (Number.isInteger(f) && f >= 0 && f < MAX_FILAS && !conCeldas.has(f)) { infoFilas.push({ f, c0: 0, c1: 0, vacia: true }); conCeldas.add(f); } }
      for (const f of filasOcultas) if (Number.isInteger(f) && f >= 0 && f < MAX_FILAS && !conCeldas.has(f)) { infoFilas.push({ f, c0: 0, c1: 0, vacia: true }); conCeldas.add(f); }

      // ---- registros de la hoja, en el orden de Excel ----
      registro(out, 0x0809, new Buf().u16(0x0600).u16(0x0010).u16(0x4F5A).u16(0x07CD).u32(0x000200C1).u32(0x0806).fin());
      registro(out, 0x000D, [1, 0]);                 // CALCMODE automático
      registro(out, 0x000C, [100, 0]);               // CALCCOUNT
      registro(out, 0x000F, [1, 0]);                 // REFMODE A1
      registro(out, 0x0011, [0, 0]);                 // ITERATION
      registro(out, 0x0010, new Buf().f64(0.001).fin());   // DELTA
      registro(out, 0x005F, [1, 0]);                 // SAVERECALC
      registro(out, 0x002A, [0, 0]);                 // PRINTHEADERS
      registro(out, 0x002B, [0, 0]);                 // PRINTGRIDLINES
      registro(out, 0x0082, [1, 0]);                 // GRIDSET
      registro(out, 0x0080, [0, 0, 0, 0, 0, 0, 0, 0]);   // GUTS
      registro(out, 0x0225, [0, 0, 0x2C, 0x01]);     // DEFAULTROWHEIGHT 15 pt
      registro(out, 0x0081, [0xC1, 0x04]);           // WSBOOL
      registro(out, 0x0055, [8, 0]);                 // DEFCOLWIDTH
      // anchos y columnas ocultas, en tramos de columnas iguales
      const anchos = Array.isArray(h.anchos) ? h.anchos : [];
      const tamCol = (c) => { const w = anchos[c]; return Number.isFinite(w) && w > 0 ? Math.min(255 * 256, Math.round(w * 256)) : 0; };
      for (let c = 0; c < MAX_COLS;) {
        const w = tamCol(c), oc = columnasOcultas.has(c);
        if (!w && !oc) { c++; continue; }
        let d = c; while (d + 1 < MAX_COLS && tamCol(d + 1) === w && columnasOcultas.has(d + 1) === oc) d++;
        registro(out, 0x007D, new Buf().u16(c).u16(d).u16(w || 2194).u16(15).u16(oc ? 1 : 0).u16(0).fin());
        c = d + 1;
      }
      registro(out, 0x0200, new Buf().u32(0).u32(ultimaFila + 1).u16(0).u16(ultimaCol + 1).u16(0).fin());   // DIMENSIONS
      infoFilas.sort((a, b) => a.f - b.f);
      for (const i of infoFilas) {
        const alto = altos[i.f], propio = Number.isFinite(alto) && alto >= 0 && alto <= 409.5;
        registro(out, 0x0208, new Buf().u16(i.f).u16(i.c0).u16(i.vacia ? i.c0 : i.c1 + 1).u16(propio ? Math.round(alto * 20) : 300).u16(0).u16(0)
          .u16(0x0100 | (propio ? 0x40 : 0) | (filasOcultas.has(i.f) ? 0x20 : 0)).u16(15).fin());
      }
      out.bytes(cuerpoCeldas.fin());
      registro(out, 0x023E, new Buf().u16(ih === 0 ? 0x06B6 : 0x04B6).u16(0).u16(0).u32(0x40).u16(0).u16(0).u32(0).fin());   // WINDOW2
      // celdas combinadas (hasta 1026 por registro)
      const combinadas = (h.combinadas || []).map(parseRango).filter((r) => r && r.f0 < MAX_FILAS && r.c0 < MAX_COLS && (r.f1 > r.f0 || r.c1 > r.c0)).map((r) => ({ ...r, f1: Math.min(r.f1, MAX_FILAS - 1), c1: Math.min(r.c1, MAX_COLS - 1) }));
      for (let i = 0; i < combinadas.length; i += 1026) {
        const lote = combinadas.slice(i, i + 1026), b = new Buf().u16(lote.length);
        for (const r of lote) b.u16(r.f0).u16(r.f1).u16(r.c0).u16(r.c1);
        registro(out, 0x00E5, b.fin());
      }
      registro(out, 0x000A, []);
      return out.fin();
    }

    // ---- nombres definidos por el usuario (sus fórmulas pueden agregar más nombres o hojas) ----
    const cuerposNombre = new Map();
    for (let i = 0; i < nombresDef.length; i++) {
      const n = nombresDef[i];
      if (n.funcion || !n.formula || !conFormulas) continue;
      hojaActual = n.hoja - 1;
      const r = codificarFormula(n.formula, 0, 0, 'R');
      if (r && 15 + n.nombre.length * 2 + r.rgce.length + r.rgcb.length <= MAX_REGISTRO) cuerposNombre.set(i, r);   // más largo no cabe en el NAME: queda #REF!
    }

    // ---- el flujo global ----
    const paleta = armarPaleta(estilos.usoColor);
    const idx = (c) => paleta.indice.get(c) === undefined ? 0 : paleta.indice.get(c);
    const g = new Buf();
    registro(g, 0x0809, new Buf().u16(0x0600).u16(0x0005).u16(0x4F5A).u16(0x07CD).u32(0x000200C1).u32(0x0806).fin());
    registro(g, 0x00E1, [0xB0, 0x04]); registro(g, 0x00C1, [0, 0]); registro(g, 0x00E2, []);
    registro(g, 0x0042, [0xB0, 0x04]);                      // CODEPAGE UTF-16
    registro(g, 0x0161, [0, 0]); registro(g, 0x01C0, []);   // DSF, EXCEL9FILE
    const tabid = new Buf(); for (let i = 1; i <= hojas.length; i++) tabid.u16(i);
    registro(g, 0x013D, tabid.fin());
    registro(g, 0x009C, [0x11, 0]);                         // FNGROUPCOUNT
    registro(g, 0x0019, [0, 0]); registro(g, 0x0012, [0, 0]); registro(g, 0x0013, [0, 0]); registro(g, 0x01AF, [0, 0]); registro(g, 0x01BC, [0, 0]);
    registro(g, 0x003D, new Buf().u16(0x00E1).u16(0x005A).u16(0x3FCF).u16(0x2A4E).u16(0x0038).u16(0).u16(0).u16(1).u16(600).fin());   // WINDOW1
    registro(g, 0x0040, [0, 0]); registro(g, 0x008D, [0, 0]);
    registro(g, 0x0022, [opciones.fecha1904 ? 1 : 0, 0]);   // DATE1904
    registro(g, 0x000E, [1, 0]); registro(g, 0x01B7, [0, 0]); registro(g, 0x00DA, [0, 0]);
    // fuentes: las cuatro primeras iguales, la quinta (índice 4) no existe nunca, y las propias desde el índice 5
    const fuentesRegistro = [estilos.fuentes[0], estilos.fuentes[0], estilos.fuentes[0], estilos.fuentes[0], ...(estilos.fuentes.length > 1 ? estilos.fuentes.slice(1) : [estilos.fuentes[0]])];   // siempre hay una propia, aunque sea de relleno
    for (const f of fuentesRegistro) registro(g, 0x0031, fuenteBytes(f, idx));
    estilos.formatos.forEach((cod, k) => { const b = new Buf().u16(164 + k); cadena(b, cod, 2); registro(g, 0x041E, b.fin()); });
    // 16 XF de base: 15 de estilo y 1 de celda por defecto; después los propios
    const xfBase = (estilo) => new Buf().u16(0).u16(0).u16(estilo ? 0xFFF5 : 0x0001).u8(0x20).u8(0).u8(0).u8(0).u16(0).u16(0).u32(0).u16(0x20C0).fin();
    for (let i = 0; i < 15; i++) registro(g, 0x00E0, xfBase(true));
    registro(g, 0x00E0, xfBase(false));
    for (const x of estilos.xfs) registro(g, 0x00E0, xfBytes(x, idx));
    // colores exactos de los XF: XFCRC y un XFEXT por cada XF con colores
    const ext = [];
    estilos.xfs.forEach((x, i) => { const e = xfExtBytes(16 + i, x, x.fuente ? estilos.fuentes[x.fuente].color : null); if (e) ext.push(e); });
    if (ext.length) {
      registro(g, 0x087C, new Buf().u16(0x087C).u16(0).u32(0).u32(0).u16(0).u16(16 + estilos.xfs.length).u32(0).fin());
      for (const e of ext) registro(g, 0x087D, e);
    }
    registro(g, 0x0293, [0x00, 0x80, 0x00, 0xFF]);           // STYLE: «Normal»
    if (paleta.cambio) { const p = new Buf().u16(56); for (let i = 8; i < 64; i++) { const rgb = hexRgb(paleta.paleta[i]); p.u8(rgb[0]).u8(rgb[1]).u8(rgb[2]).u8(0); } registro(g, 0x0092, p.fin()); }
    registro(g, 0x0160, [1, 0]);                             // USESELFS

    // BOUNDSHEET: la posición de cada hoja se conoce al final del flujo global; se arma todo con un hueco y se corrige
    const huecosBoundsheet = [];
    hojas.forEach((h) => {
      huecosBoundsheet.push(g.n + 4);
      const b = new Buf().u32(0).u8(h.estado === 'oculta' ? 1 : h.estado === 'muyOculta' ? 2 : 0).u8(0);
      cadena(b, h.nombre, 1);
      registro(g, 0x0085, b.fin());
    });
    registro(g, 0x008C, [0x01, 0x00, 0x01, 0x00]);           // COUNTRY
    // libros de referencia: 0 interno, 1 complemento
    const hayRefs = xtis.length || nombresDef.length;
    if (hayRefs || complemento.nombres.length) {
      registro(g, 0x01AE, new Buf().u16(hojas.length).u16(0x0401).fin());
      if (complemento.nombres.length) {
        registro(g, 0x01AE, new Buf().u16(1).u16(0x3A01).fin());
        for (const nom of complemento.nombres) { const b = new Buf().u16(0).u32(0); cadena(b, nom, 1); b.u16(2).u8(0x1C).u8(0x17); registro(g, 0x0023, b.fin()); }
      }
      // más de 1.370 entradas no caben en un registro: siguen en CONTINUE, cortadas entre entradas de 6 bytes (como LibreOffice)
      const e = new Buf().u16(xtis.length);
      let id = 0x0017, lleno = 2;
      for (const x of xtis) {
        if (lleno + 6 > MAX_REGISTRO) { registro(g, id, e.fin()); e.n = 0; id = 0x003C; lleno = 0; }
        e.u16(x.libro).u16(x.primera & 0xFFFF).u16(x.ultima & 0xFFFF); lleno += 6;
      }
      registro(g, id, e.fin());
    }
    nombresDef.forEach((n, i) => {
      const cuerpo = n.funcion ? { rgce: Uint8Array.from([0x1C, 0x1D]), rgcb: new Uint8Array(0) } : cuerposNombre.get(i) || { rgce: Uint8Array.from([0x1C, 0x17]), rgcb: new Uint8Array(0) };
      const b = new Buf().u16(n.funcion ? 0x000B : 0).u8(0).u8(n.nombre.length).u16(cuerpo.rgce.length).u16(0).u16(n.hoja).u32(0);
      let ancho = false; for (let k = 0; k < n.nombre.length; k++) if (n.nombre.charCodeAt(k) > 0xFF) ancho = true;
      b.u8(ancho ? 1 : 0);
      for (let k = 0; k < n.nombre.length; k++) { if (ancho) b.u16(n.nombre.charCodeAt(k)); else b.u8(n.nombre.charCodeAt(k)); }
      b.bytes(cuerpo.rgce).bytes(cuerpo.rgcb);
      registro(g, 0x0018, b.fin());
    });
    escribirSst(g, cadenas, totalCadenas);
    registro(g, 0x000A, []);

    // posiciones de las hojas y flujo final
    const globales = g.fin();
    let pos = globales.length;
    const dv = new DataView(globales.buffer, globales.byteOffset, globales.length);
    cuerposHoja.forEach((c, i) => { dv.setUint32(huecosBoundsheet[i], pos, true); pos += c.length; });
    const flujo = new Uint8Array(pos);
    flujo.set(globales, 0);
    let q = globales.length; for (const c of cuerposHoja) { flujo.set(c, q); q += c.length; }
    const bytes = CfbPropio.escribir([{ nombre: 'Workbook', bytes: flujo }]);
    return { bytes, avisos: [...avisosMapa.values()] };
  }

  globalThis.XlsEscritor = { crearLibro };
})();

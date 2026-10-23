/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Escritor de .doc propio (Word 97-2003, formato binario [MS-DOC]) desde el «contrato de documento» que describe docx-desde-contrato.js. Es el gemelo de DocLector: lo que ese
// lector entiende, este lo escribe, así que cualquier lector (DocxAContrato, DocLector, OdtLector) → .doc sin un conversor por pareja.
//
//   const { bytes, mensajes } = await DocEscritor.generar(contrato);     // Uint8Array con el .doc (contenedor CFB con WordDocument, 1Table y Data)
//
// Lleva: texto en una sola pieza Unicode, estilos con nombre (párrafo y carácter), formato de carácter y de párrafo, tabulaciones, listas (numeradas y de viñetas, con niveles),
// secciones (hoja, márgenes, columnas, tipo) con su encabezado y pie, tablas (anidadas, combinadas, bordes, relleno, alto de fila), imágenes en línea (PNG y JPEG), enlaces,
// campos (número de página, total de páginas, fecha), marcadores, saltos de página y de columna.
// mensajes: 'imagen-formato' (una imagen que no es png ni jpeg se omite), 'enlace-descartado' (un destino que no es una dirección web).
// Requiere cfb-propio.js. No depende del navegador.
(function () {
  'use strict';
  const MAX_TEXTO = 64 * 1024 * 1024;
  const MAX_BLOQUES = 4000000;          // párrafos (y marcas de fila) que se escriben en total: Word guarda las historias de cada sección, así que un encabezado repetido se escribe cada vez
  const MAX_DATA = 256 * 1024 * 1024;   // bytes del flujo Data (imágenes)
  const ERROR_GRANDE = 'El documento es demasiado grande para guardarlo como .doc.';
  const MAX_ESTILOS = 4000;
  const MAX_SECCIONES = 4096;
  const MAX_FUENTES = 30000;            // Word abre un .doc con 31.000 fuentes y no con 31.500 (medido el 2026-10-23)
  const MAX_PROFUNDIDAD_TABLAS = 8;
  const FC_TEXTO = 0x800;                 // dónde empieza el texto en el flujo WordDocument (detrás del FIB)
  const sinProto = (o) => Object.assign(Object.create(null), o);
  const entero = (v, min, max, def = 0) => { const n = Math.round(Number(v)); return Number.isFinite(n) ? Math.max(min, Math.min(max, n)) : def; };

  // ───────── Bytes ─────────
  class Bytes {
    constructor(n = 1024) { this.a = new Uint8Array(n); this.n = 0; }
    asegura(k) { if (this.n + k > this.a.length) { let m = this.a.length * 2; while (m < this.n + k) m *= 2; const b = new Uint8Array(m); b.set(this.a.subarray(0, this.n)); this.a = b; } }
    u8(v) { this.asegura(1); this.a[this.n++] = v & 255; return this; }
    u16(v) { this.asegura(2); this.a[this.n++] = v & 255; this.a[this.n++] = (v >> 8) & 255; return this; }
    u32(v) { this.asegura(4); this.a[this.n++] = v & 255; this.a[this.n++] = (v >>> 8) & 255; this.a[this.n++] = (v >>> 16) & 255; this.a[this.n++] = (v >>> 24) & 255; return this; }
    bytes(b) { this.asegura(b.length); this.a.set(b, this.n); this.n += b.length; return this; }
    ceros(k) { this.asegura(k); this.n += k; return this; }
    alinear(m) { while (this.n % m) this.u8(0); return this; }
    poner32(p, v) { this.a[p] = v & 255; this.a[p + 1] = (v >>> 8) & 255; this.a[p + 2] = (v >>> 16) & 255; this.a[p + 3] = (v >>> 24) & 255; }
    poner16(p, v) { this.a[p] = v & 255; this.a[p + 1] = (v >> 8) & 255; }
    salida() { return this.a.slice(0, this.n); }
  }
  const utf16 = (s) => { const b = new Bytes(s.length * 2 + 2); for (let i = 0; i < s.length; i++) b.u16(s.charCodeAt(i)); b.u16(0); return b.salida(); };   // con el terminador nulo

  // ───────── Sprms ─────────
  // Una lista de sprms (grpprl) se arma con operandos del tamaño que pide su spra (bits 13-15 del código) y va ordenada por grupo y por número, como la escribe Word.
  class Sprms {
    constructor() { this.l = []; }
    _poner(op, operando) { this.l.push({ op, operando }); return this; }
    s1(op, v) { return this._poner(op, [v & 255]); }
    s2(op, v) { return this._poner(op, [v & 255, (v >> 8) & 255]); }
    s3(op, v) { return this._poner(op, [v & 255, (v >> 8) & 255, (v >> 16) & 255]); }
    s4(op, v) { return this._poner(op, [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255]); }
    sv(op, bytes) { return this._poner(op, [bytes.length, ...bytes]); }                       // operando variable, tamaño en 1 byte
    sw(op, bytes) { const cb = bytes.length + 1; return this._poner(op, [cb & 255, (cb >> 8) & 255, ...bytes]); }   // sprmTDefTable: cb de 2 bytes que cuenta todo menos 1
    vacio() { return this.l.length === 0; }
    bytes() {
      // las marcas de tabla anidada (0x244B y 0x244C) van al final, como las deja Word
      const tarde = (op) => (op === 0x244B || op === 0x244C ? 1 : 0);
      const orden = this.l.map((x, i) => ({ x, i })).sort((a, b) => (tarde(a.x.op) - tarde(b.x.op)) || (((a.x.op >> 10) & 7) - ((b.x.op >> 10) & 7)) || ((a.x.op & 0x1FF) - (b.x.op & 0x1FF)) || (a.i - b.i));
      const b = new Bytes(64);
      for (const { x } of orden) { b.u16(x.op); for (const v of x.operando) b.u8(v); }
      return b.salida();
    }
  }

  const ICO = [null, [0, 0, 0], [0, 0, 255], [0, 255, 255], [0, 255, 0], [255, 0, 255], [255, 0, 0], [255, 255, 0], [255, 255, 255], [0, 0, 128], [0, 128, 128], [0, 128, 0], [128, 0, 128], [128, 0, 0], [128, 128, 0], [128, 128, 128], [192, 192, 192]];
  const RESALTADO = sinProto({ black: 1, blue: 2, cyan: 3, green: 4, magenta: 5, red: 6, yellow: 7, white: 8, darkBlue: 9, darkCyan: 10, darkGreen: 11, darkMagenta: 12, darkRed: 13, darkYellow: 14, darkGray: 15, lightGray: 16 });
  const rgb = (hex) => (/^[0-9a-fA-F]{6}$/.test(hex || '') ? [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)] : null);
  // el índice de los 16 colores de Word más cercano a un color
  function icoCercano(hex) {
    const c = rgb(hex);
    if (!c) return 0;
    let mejor = 1, dist = Infinity;
    for (let i = 1; i <= 16; i++) { const d = (ICO[i][0] - c[0]) ** 2 + (ICO[i][1] - c[1]) ** 2 + (ICO[i][2] - c[2]) ** 2; if (d < dist) { dist = d; mejor = i; } }
    return mejor;
  }
  const ALINEACION = sinProto({ izquierda: 0, centro: 1, derecha: 2, justificado: 3 });
  const TAB_TIPO = sinProto({ izquierda: 0, centro: 1, derecha: 2, decimal: 3 });
  const TAB_RELLENO = sinProto({ punto: 1, guion: 2, linea: 3 });
  const NFC = sinProto({ decimal: 0, upperRoman: 1, lowerRoman: 2, upperLetter: 3, lowerLetter: 4, ordinal: 5, cardinalText: 6, ordinalText: 7, decimalZero: 22, bullet: 23, none: 255 });
  const BORDE_TIPO = sinProto({ simple: 1, doble: 3, punteado: 6, guiones: 7 });
  // sti (identificador de estilo integrado) de los estilos que Word reconoce por nombre; los demás son estilos del usuario (4094)
  const STI = sinProto({ normal: 0, 'heading 1': 1, 'heading 2': 2, 'heading 3': 3, 'heading 4': 4, 'heading 5': 5, 'heading 6': 6, 'heading 7': 7, 'heading 8': 8, 'heading 9': 9,
    'toc 1': 19, 'toc 2': 20, 'toc 3': 21, 'toc 4': 22, 'toc 5': 23, 'toc 6': 24, 'toc 7': 25, 'toc 8': 26, 'toc 9': 27, header: 31, footer: 32, caption: 35, 'list': 47, 'list bullet': 48, 'list number': 49,
    title: 62, 'body text': 66, subtitle: 74, 'plain text': 78, hyperlink: 85, strong: 87, emphasis: 88, 'no spacing': 157, 'list paragraph': 179 });
  const STI_CARACTER = new Set([85, 87, 88]);

  function tipoDeImagen(b) {
    if (b.length > 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4E && b[3] === 0x47) return 'png';
    if (b.length > 3 && b[0] === 0xFF && b[1] === 0xD8 && b[2] === 0xFF) return 'jpeg';
    return null;
  }

  // El contenedor de datos del documento mientras se escribe: el texto de todas las historias (principal y encabezados) y los formatos por posición
  class Documento {
    constructor() {
      this.T = new Uint16Array(65536); this.n = 0;
      this.pars = [];       // { fin, istd, g }  párrafos: terminan en fin (posición del siguiente carácter)
      this.runs = [];       // { fin, g, clave } formato de carácter por tramos consecutivos
      this.campos = [];     // { cp, tipo, ch, flags } caracteres de campo, por historia
      this.marcadores = []; // { nombre, cp }
    }
    car(c) {
      if (this.n >= MAX_TEXTO) throw new Error('El documento es demasiado grande para guardarlo como .doc.');
      if (this.n >= this.T.length) { const t = new Uint16Array(this.T.length * 2); t.set(this.T); this.T = t; }
      this.T[this.n++] = c;
    }
    texto(s, g, clave) {
      if (!s) return;
      for (let i = 0; i < s.length; i++) this.car(s.charCodeAt(i));
      this.tramo(s.length, g, clave);
    }
    tramo(largo, g, clave) {
      const k = clave || '';
      const u = this.runs[this.runs.length - 1];
      if (u && u.clave === k && u.fin === this.n - largo) u.fin = this.n; else this.runs.push({ fin: this.n, g: g || null, clave: k });
    }
  }

  // El grupo de dibujos del documento (OfficeArt DggInfo) que Word escribe cuando hay imágenes: la forma de fondo de la página y los datos de la numeración de formas.
  const DGG_INFO = Uint8Array.from('0f 00 00 f0 4c 00 00 00 00 00 06 f0 18 00 00 00 02 04 00 00 02 00 00 00 01 00 00 00 01 00 00 00 01 00 00 00 02 00 00 00 23 00 0b f0 0c 00 00 00 86 c1 00 00 00 00 c5 c1 00 00 00 00 40 00 1e f1 10 00 00 00 ff ff 00 00 00 00 ff 00 80 80 80 00 f7 00 00 10 00 0f 00 02 f0 92 00 00 00 10 00 08 f0 08 00 00 00 01 00 00 00 01 04 00 00 0f 00 03 f0 30 00 00 00 0f 00 04 f0 28 00 00 00 01 00 09 f0 10 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 02 00 0a f0 08 00 00 00 00 04 00 00 05 00 00 00 0f 00 04 f0 42 00 00 00 12 00 0a f0 08 00 00 00 01 04 00 00 00 0e 00 00 53 00 0b f0 1e 00 00 00 bf 01 00 00 10 00 cb 01 00 00 00 00 ff 01 00 00 08 00 04 03 09 00 00 00 3f 03 01 00 01 00 00 00 11 f0 04 00 00 00 01 00 00 00'.split(' ').map((h) => parseInt(h, 16)));

  async function generar(contrato) {
    const mensajes = [];
    const aviso = (m) => { if (!mensajes.includes(m)) mensajes.push(m); };
    const D = new Documento();
    const data = new Bytes(1024);          // flujo Data: imágenes y propiedades de párrafo que no caben en la página de formatos
    let bloquesEscritos = 0;
    // El texto de un enlace o de un campo va dentro de la instrucción o del resultado, entre las marcas de campo: sin caracteres de control (una marca de campo, de párrafo, de celda o de sección falsa lo rompería)
    const limpio = (s) => String(s).replace(/[\u0000-\u0008\u000A-\u001F\u007F]/g, '');

    // ───────── Fuentes ─────────
    const fuentes = [], fuenteIdx = new Map();
    const ftc = (nombre) => {
      const n = String(nombre || 'Times New Roman').slice(0, 31);
      if (!fuenteIdx.has(n)) { if (fuentes.length >= MAX_FUENTES) return 0; fuenteIdx.set(n, fuentes.length); fuentes.push(n); }   // las de más quedan con la predeterminada
      return fuenteIdx.get(n);
    };
    const pd = contrato.predeterminado || {};
    const fuenteBase = ftc((pd.r && pd.r.fuente) || 'Times New Roman');
    const fuenteAsiaBase = ftc((pd.r && pd.r.fuenteAsia) || (pd.r && pd.r.fuente) || 'Times New Roman');
    const fuenteCsBase = ftc((pd.r && pd.r.fuenteCs) || (pd.r && pd.r.fuente) || 'Times New Roman');

    // ───────── Propiedades de carácter ─────────
    function sprmsChp(r, istdCar) {
      const s = new Sprms();
      if (istdCar != null) s.s2(0x4A30, istdCar);
      if (!r) return s;
      if (r.negrita != null) s.s1(0x0835, r.negrita ? 1 : 0);
      if (r.cursiva != null) s.s1(0x0836, r.cursiva ? 1 : 0);
      if (r.tachado != null) s.s1(0x0837, r.tachado ? 1 : 0);
      if (r.versalitas != null) s.s1(0x083A, r.versalitas ? 1 : 0);
      if (r.mayusculas != null) s.s1(0x083B, r.mayusculas ? 1 : 0);
      if (r.oculto) s.s1(0x083C, 1);
      if (r.subrayado != null) s.s1(0x2A3E, r.subrayado === 'doble' ? 3 : r.subrayado ? 1 : 0);
      const c = rgb(r.color);
      if (c) s.s4(0x6870, c[0] | (c[1] << 8) | (c[2] << 16));
      if (r.tamano != null) { const h = entero(r.tamano * 2, 2, 3276, 22); s.s2(0x4A43, h); s.s2(0x4A61, h); }
      if (r.sup) s.s1(0x2A48, 1); else if (r.sub) s.s1(0x2A48, 2);
      if (r.resaltado && RESALTADO[r.resaltado]) s.s1(0x2A0C, RESALTADO[r.resaltado]);
      if (r.fuente) { const i = ftc(r.fuente); s.s2(0x4A4F, i); s.s2(0x4A51, i); }
      if (r.fuenteAsia) s.s2(0x4A50, ftc(r.fuenteAsia));
      if (r.fuenteCs) s.s2(0x4A5E, ftc(r.fuenteCs));
      return s;
    }

    // ───────── Estilos ─────────
    const estilosContrato = (contrato.estilos || []).filter((e) => e && e.tipo !== 'tabla').slice(0, MAX_ESTILOS);
    const istdDe = new Map();      // id del contrato → istd
    const filasStd = [];           // por istd: el estilo del contrato o null (hueco)
    {
      const porNombre = (e) => String(e.nombre || '').toLowerCase();
      filasStd.length = 15; filasStd.fill(null);
      const normal = estilosContrato.find((e) => e.id === 'Normal') || estilosContrato.find((e) => e.tipo !== 'caracter' && porNombre(e) === 'normal');
      filasStd[0] = normal || { id: 'Normal', nombre: 'Normal', tipo: 'parrafo' };
      istdDe.set(filasStd[0].id, 0);
      for (const e of estilosContrato) {
        if (e === normal || istdDe.has(e.id)) continue;
        const m = e.tipo !== 'caracter' ? /^heading ([1-9])$/.exec(porNombre(e)) : null;
        if (m && !filasStd[+m[1]]) { filasStd[+m[1]] = e; istdDe.set(e.id, +m[1]); continue; }
        filasStd.push(e);
        istdDe.set(e.id, filasStd.length - 1);
      }
    }
    const deEstilo = (id) => (id != null && istdDe.has(id) ? istdDe.get(id) : null);

    // ───────── Propiedades de párrafo ─────────
    function sprmsPap(p, opciones) {
      const s = new Sprms();
      if (!p) p = {};
      if (p.alineacion && ALINEACION[p.alineacion] != null) s.s1(0x2461, ALINEACION[p.alineacion]);
      if (p.mantenerJuntas) s.s1(0x2405, 1);
      if (p.mantener) s.s1(0x2406, 1);
      if (p.saltoAntes) s.s1(0x2407, 1);
      if (p.lista) {
        const ilfo = p.lista.id == null ? 0 : opciones.ilfoDe(p.lista.id);
        if (ilfo != null) { s.s2(0x460B, ilfo); if (ilfo > 0) s.s1(0x260A, entero(p.lista.nivel, 0, 8)); }
      }
      if (p.izq != null) s.s2(0x845E, entero(p.izq, -31680, 31680));
      if (p.der != null) s.s2(0x845D, entero(p.der, -31680, 31680));
      if (p.primera != null) s.s2(0x8460, entero(p.primera, -31680, 31680));
      if (p.linea != null) {
        const l = entero(p.linea, 1, 31680, 240);
        s.s4(0x6412, p.regla === 'exacto' ? ((-l) & 0xFFFF) : p.regla === 'minimo' ? l : (l | (1 << 16)));
      }
      if (p.antes != null) s.s2(0xA413, entero(p.antes, 0, 31680));
      if (p.despues != null) s.s2(0xA414, entero(p.despues, 0, 31680));
      if (p.nivelEsquema != null) s.s1(0x2640, entero(p.nivelEsquema, 0, 9, 9));
      if (p.tabs && p.tabs.length) {
        const tabs = p.tabs.slice(0, 64).map((t) => ({ pos: entero(t.pos, -31680, 31680), tipo: (TAB_TIPO[t.tipo] || 0) | ((TAB_RELLENO[t.relleno] || 0) << 3) })).sort((a, b) => a.pos - b.pos);
        const b = new Bytes(16);
        b.u8(0).u8(tabs.length);
        for (const t of tabs) b.u16(t.pos);
        for (const t of tabs) b.u8(t.tipo);
        s.sv(0xC60D, b.salida());
      }
      if (p.relleno && rgb(p.relleno)) {
        const c = rgb(p.relleno), sb = new Bytes(10);
        sb.u32(0xFF000000).u32(c[0] | (c[1] << 8) | (c[2] << 16)).u16(0);   // SHD: primer plano automático, fondo del color, trama clara
        s.sv(0xC64D, sb.salida());
      }
      if (p.bordes) {
        for (const [clave, op] of [['arriba', 0xC64E], ['izquierda', 0xC64F], ['abajo', 0xC650], ['derecha', 0xC651]]) {
          const l = p.bordes[clave];
          if (!l) continue;
          s.sv(op, brc8(l, l.espacio));
        }
      }
      return s;
    }
    // BRC de 8 bytes: color (COLORREF), ancho (octavos de punto), tipo, espacio (puntos) y banderas
    function brc8(l, espacio) {
      const c = rgb(l.color), b = new Bytes(8);
      if (c) b.u8(c[0]).u8(c[1]).u8(c[2]).u8(0); else b.u32(0xFF000000);
      b.u8(entero(l.tam == null ? 4 : l.tam, 2, 96, 4)).u8(BORDE_TIPO[l.estilo || 'simple'] || 1).u16(entero(espacio || 0, 0, 31) & 0x1F);
      return b.salida();
    }
    function brc80(l) {
      const b = new Bytes(4);
      if (!l) return b.u32(0).salida();
      b.u8(entero(l.tam == null ? 4 : l.tam, 2, 96, 4)).u8(BORDE_TIPO[l.estilo || 'simple'] || 1).u8(icoCercano(l.color)).u8(0);
      return b.salida();
    }

    // ───────── Listas ─────────
    const listas = (contrato.listas || []).slice(0, 2000);
    const ilfoPorId = new Map(listas.map((l, i) => [l.id, i + 1]));
    const ilfoDe = (id) => (ilfoPorId.has(id) ? ilfoPorId.get(id) : null);

    // ───────── Imágenes ─────────
    const imagenesEscritas = new Map();   // Uint8Array → posición en Data
    let hayImagenes = false, numImagenes = 0;
    function escribirImagen(img) {
      const bytes = img.datos instanceof Uint8Array ? img.datos : new Uint8Array(img.datos || 0);
      const tipo = tipoDeImagen(bytes);
      if (!tipo) { aviso('imagen-formato'); return null; }
      const ancho = entero((Number(img.ancho) || 72) * 20, 20, 31680, 1440), alto = entero((Number(img.alto) || 72) * 20, 20, 31680, 1440);
      const clave = `${tipo}|${ancho}|${alto}`;
      let porImagen = imagenesEscritas.get(bytes);
      if (!porImagen) { porImagen = new Map(); imagenesEscritas.set(bytes, porImagen); }
      if (porImagen.has(clave)) return porImagen.get(clave);
      // Word guarda la imagen dentro de cada dibujo en línea, así que cada tamaño distinto de la misma imagen la lleva otra vez: el flujo Data tiene tope
      if (data.n + bytes.length + 1024 > MAX_DATA) throw new Error(ERROR_GRANDE);
      hayImagenes = true;
      data.alinear(4);
      const pos = data.n;
      const numero = ++numImagenes;
      // un identificador de 16 bytes distinto por imagen (Word no los repite)
      const uid = new Uint8Array(16);
      let sumaA = 0x9E3779B1 ^ bytes.length, sumaB = numero * 0x85EBCA6B;
      for (let i = 0; i < bytes.length; i++) { sumaA = (Math.imul(sumaA, 31) + bytes[i]) | 0; sumaB = (sumaB ^ (bytes[i] << ((i & 3) * 8))) | 0; }
      for (let i = 0; i < 4; i++) { uid[i] = (sumaA >>> (8 * i)) & 255; uid[4 + i] = (sumaB >>> (8 * i)) & 255; uid[8 + i] = (numero >>> (8 * i)) & 255; uid[12 + i] = (bytes.length >>> (8 * i)) & 255; }
      // OfficeArtInlineSpContainer: contenedor de la forma (FSP + FOPT con la propiedad pib) y el FBSE con la imagen
      // la imagen se copia una sola vez, directamente a Data (antes pasaba por ocho copias intermedias): aquí solo se arman las cabeceras, que son chicas
      const largoBlip = 16 + 1 + bytes.length;       // uid, marca 0xFF e imagen
      const largoBlipRec = 8 + largoBlip;
      const cabBlipRec = new Bytes(8);
      cabBlipRec.u16((tipo === 'png' ? 0x06E0 : 0x046A) << 4).u16(tipo === 'png' ? 0xF01E : 0xF01D).u32(largoBlip);
      // la forma, con los mismos registros que escribe Word para una imagen en línea
      const fsp = new Bytes(16); fsp.u16(0x0002 | (75 << 4)).u16(0xF00A).u32(8).u32(0x0400 + numero).u32(0x0A00);
      const nombre = utf16('Picture ' + numero);
      const fopt = new Bytes(64 + nombre.length);
      fopt.u16(0x0003 | (7 << 4)).u16(0xF00B).u32(42 + nombre.length);
      fopt.u16(0x007F).u32(0x00E10080).u16(0x4104).u32(numero).u16(0x013F).u32(0x00060000).u16(0x01BF).u32(0x00100000).u16(0x01FF).u32(0x00080000).u16(0xC380).u32(nombre.length).u16(0x03BF).u32(0x00020000);
      fopt.bytes(nombre);
      const fopt2 = new Bytes(20); fopt2.u16(0x0003 | (1 << 4)).u16(0xF122).u32(6).u16(0x03AA).u32(0x0F000000);
      const ancla = new Bytes(12); ancla.u16(0x0000).u16(0xF010).u32(4).u32((numero - 1) | 0x80000000);
      const interior = fsp.n + fopt.n + fopt2.n + ancla.n;
      const contenedor = new Bytes(interior + 8);
      contenedor.u16(0x000F).u16(0xF004).u32(interior).bytes(fsp.salida()).bytes(fopt.salida()).bytes(fopt2.salida()).bytes(ancla.salida());
      const inicioArte = pos + 68;
      const cabFbse = new Bytes(44);
      cabFbse.u16(0x0002 | ((tipo === 'png' ? 6 : 5) << 4)).u16(0xF007).u32(36 + largoBlipRec);
      cabFbse.u8(tipo === 'png' ? 6 : 5).u8(tipo === 'png' ? 6 : 5).bytes(uid).u16(0xFF).u32(largoBlipRec).u32(1).u32(inicioArte).u8(0).u8(0).u8(0).u8(0);
      const largoCuerpo = contenedor.n + cabFbse.n + largoBlipRec;
      const picf = new Bytes(68);
      picf.u32(68 + largoCuerpo).u16(68).u16(0x64).u16(0).u16(0).u16(0).u8(0x08).ceros(13);
      picf.u16(ancho).u16(alto).u16(1000).u16(1000);   // dxaGoal, dyaGoal, mx, my
      picf.ceros(68 - picf.n);
      data.asegura(68 + largoCuerpo);
      data.bytes(picf.salida()).bytes(contenedor.salida()).bytes(cabFbse.salida()).bytes(cabBlipRec.salida()).bytes(uid).u8(0xFF).bytes(bytes);
      porImagen.set(clave, pos);
      return pos;
    }

    // ───────── Texto: párrafos y runs ─────────
    const camposPorHistoria = { principal: [], encabezados: [] };
    let historiaActual = 'principal', inicioHistoria = 0;
    const marcarCampo = (tipo, ch, flags) => camposPorHistoria[historiaActual].push({ cp: D.n - inicioHistoria, ch, flags: flags | 0, tipo });
    const claveR = (s) => s.l.map((x) => x.op + ':' + x.operando.join(',')).join(';');
    const cacheRuns = new Map();
    function formato(r, estilo) {
      const istd = estilo ? deEstilo(estilo) : null;
      const s = sprmsChp(r, istd);
      if (s.vacio()) return { g: null, clave: '' };
      const clave = claveR(s);
      let g = cacheRuns.get(clave);
      if (!g) { g = s.bytes(); cacheRuns.set(clave, g); }
      return { g, clave };
    }

    function runs(lista, ctx, prof) {
      for (const r of lista || []) {
        if (r.salto) {
          if (r.salto === 'pagina') { D.car(12); D.tramo(1, null, ''); }
          else if (r.salto === 'columna') { D.car(14); D.tramo(1, null, ''); }
          else { D.car(11); D.tramo(1, null, ''); }
          continue;
        }
        if (r.marcador != null) { D.marcadores.push({ nombre: String(r.marcador).slice(0, 40), cp: D.n }); continue; }
        if (r.imagen) {
          const pos = escribirImagen(r.imagen);
          if (pos == null) continue;
          const s = new Sprms(); s.s1(0x0855, 1); s.s4(0x6A03, pos);
          D.car(1); D.tramo(1, s.bytes(), 'img:' + pos);
          continue;
        }
        if (r.enlace) {
          if (prof > 3) { runs(r.hijos, ctx, prof + 1); continue; }   // más de 3 enlaces anidados: se escribe su contenido sin el enlace
          let instr = null;
          if (r.enlace.ancla) instr = ` HYPERLINK \\l "${limpio(String(r.enlace.ancla).slice(0, 40)).replace(/"/g, '')}" `;
          else if (/^(https?:|mailto:|tel:|ftp:)/i.test(r.enlace.url || '')) instr = ` HYPERLINK "${limpio(r.enlace.url).replace(/"/g, '%22')}" `;
          if (!instr) { aviso('enlace-descartado'); runs(r.hijos, ctx, prof + 1); continue; }
          campo(0x58, instr, () => runs(r.hijos, ctx, prof + 1), null);
          continue;
        }
        if (r.campo) {
          const nombre = String(r.campo).trim().split(/\s+/)[0].toUpperCase();
          const tipo = { PAGE: 0x21, NUMPAGES: 0x1A, DATE: 0x1F, TIME: 0x20, FILENAME: 0x1D, AUTHOR: 0x11, TITLE: 0x0F, SECTIONPAGES: 0x1B, SAVEDATE: 0x16, PRINTDATE: 0x17, NUMWORDS: 0x1C, NUMCHARS: 0x1E }[nombre];
          if (tipo == null) { const f = formato(r.r, r.estilo); D.texto(limpio(r.resultado == null ? '' : r.resultado), f.g, f.clave); continue; }
          const f = formato(r.r, r.estilo);
          // la instrucción es el nombre del campo y, como mucho, unos pocos modificadores de formato (\@ fecha, \* mayúsculas, \# número) con su texto corto entre comillas
          const mods = (String(r.campo).trim().slice(nombre.length).match(/\\[@*#]\s*(?:"[^"\\\u0000-\u001F\u007F]{0,60}"|[^\s"\\\u0000-\u001F\u007F]{1,30})/g) || []).slice(0, 4);
          campo(tipo, ` ${nombre}${mods.length ? ' ' + mods.join(' ') : ''} `, () => D.texto(limpio(r.resultado == null ? '1' : r.resultado), f.g, f.clave), f);
          continue;
        }
        const f = formato(r.r, r.estilo);
        const t = String(r.texto == null ? '' : r.texto).replace(/\r\n?/g, '\n');
        let acumulado = '';
        const volcar = () => { if (acumulado) { D.texto(acumulado, f.g, f.clave); acumulado = ''; } };
        for (let i = 0; i < t.length; i++) {
          const c = t.charCodeAt(i);
          if (c === 10) { volcar(); D.car(11); D.tramo(1, f.g, f.clave); }
          else if (c === 9) acumulado += '\t';
          else if (c < 32 || c === 127) continue;   // los demás caracteres de control no son texto
          else if (c === 0x2011) { volcar(); D.car(30); D.tramo(1, f.g, f.clave); }
          else if (c === 0x00AD) { volcar(); D.car(31); D.tramo(1, f.g, f.clave); }
          else acumulado += t[i];
        }
        volcar();
      }
    }

    // Un campo: 0x13 instrucción 0x14 resultado 0x15; las marcas llevan la bandera de «código de campo oculto»
    function campo(tipo, instr, resultado, f) {
      // las tres marcas del campo llevan la bandera «carácter especial» (con posición de imagen 0), como las escribe Word
      const marca = new Sprms(); marca.s4(0x6A03, 0); marca.s1(0x0855, 1);
      const mg = marca.bytes();
      marcarCampo('inicio', 0x13, tipo); D.car(0x13); D.tramo(1, mg, 'fldmarca');
      D.texto(instr, null, '');
      marcarCampo('sep', 0x14, 0xFF); D.car(0x14); D.tramo(1, mg, 'fldmarca');
      resultado();
      marcarCampo('fin', 0x15, tipo === 0x58 ? 0x8C : 0x80); D.car(0x15); D.tramo(1, mg, 'fldmarca');
      void f;
    }

    // Termina un párrafo: la marca, su estilo y sus propiedades
    // Las propiedades iguales comparten un solo arreglo de bytes (hasta cierto número de distintas: con claves todas diferentes la caché no sirve de nada); no se guarda la lista de sprms
    const cachePars = new Map();
    function finParrafo(marca, istd, sprms) {
      if (++bloquesEscritos > MAX_BLOQUES) throw new Error(ERROR_GRANDE);
      D.car(marca); D.tramo(1, null, '');
      let g = null;
      if (sprms && !sprms.vacio()) {
        const clave = claveR(sprms);
        g = cachePars.get(clave);
        if (!g) { g = sprms.bytes(); if (cachePars.size < 50000) cachePars.set(clave, g); }
      }
      D.pars.push({ fin: D.n, istd, g });
    }

    function parrafo(b, tabla, marcaFinal) {
      const estilo = b.estilo != null ? deEstilo(b.estilo) : null;
      const istd = estilo == null ? 0 : estilo;
      const s = sprmsPap(b.p, { ilfoDe });
      if (tabla) {
        s.s1(0x2416, 1);
        if (tabla.itap > 1) s.s4(0x6649, tabla.itap);
        else s.s4(0x6649, 1);
        if (tabla.finCelda && tabla.itap > 1) s.s1(0x244B, 1);
        s.sv(0xC66C, [0x0B, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);   // sprmPTableProps: propiedades de tabla del párrafo (ancho automático)
      }
      runs(b.hijos, null, 0);
      finParrafo(marcaFinal || (tabla && tabla.finCelda && tabla.itap === 1 ? 7 : 13), istd, s);
    }

    // ───────── Tablas ─────────
    // Word no pasa de 63 celdas por fila: las de más se juntan en la última (con su texto y su ancho); si no, sus marcas de celda quedan sueltas dentro de la celda 63
    function juntarCeldas(celdas) {
      const resto = celdas.slice(62);
      const span = Math.min(63, resto.reduce((s, c) => s + Math.max(1, Math.min(63, c.fusionH || 1)), 0));
      return celdas.slice(0, 62).concat([Object.assign({}, resto[0], { fusionH: span, hijos: resto.flatMap((c) => (c.hijos && c.hijos.length ? c.hijos : [])) })]);
    }
    function tabla(t, itap) {
      const filas = (t.filas || []).filter((f) => f.celdas && f.celdas.length).map((f) => (f.celdas.length > 63 ? Object.assign({}, f, { celdas: juntarCeldas(f.celdas) }) : f));
      if (!filas.length || itap > MAX_PROFUNDIDAD_TABLAS) return false;
      const posiciones = filas.map((f) => { let c = 0; return f.celdas.map((celda) => { const x = c; c += Math.max(1, Math.min(63, celda.fusionH || 1)); return x; }); });
      const nCols = posiciones.reduce((m, ps, i) => Math.max(m, ps.length ? ps[ps.length - 1] + Math.max(1, filas[i].celdas[ps.length - 1].fusionH || 1) : 0), Math.max(t.anchos ? t.anchos.length : 0, 1));
      const anchos = Array.from({ length: Math.min(nCols, 126) }, (_, i) => entero(t.anchos && t.anchos[i], 0, 31680, 1440));
      const izq = entero(t.sangria || 0, -31680, 31680);
      const bordeTabla = t.bordes === false ? null : { tam: (t.bordes && t.bordes.tam) || 4, color: t.bordes && t.bordes.color };
      for (let fi = 0; fi < filas.length; fi++) {
        const f = filas[fi];
        f.celdas.forEach((c, ci) => {
          const hijos = c.hijos && c.hijos.length ? c.hijos : [{ tipo: 'parrafo', hijos: [] }];
          // la celda termina siempre en un párrafo (que lleva la marca de celda)
          let ultimoEsParrafo = hijos[hijos.length - 1].tipo !== 'tabla';
          hijos.forEach((h, k) => {
            const esUltimo = k === hijos.length - 1;
            if (h.tipo === 'tabla') { if (!tabla(h, itap + 1)) parrafo({ tipo: 'parrafo', hijos: [] }, { itap, finCelda: false }); }
            else parrafo(h, { itap, finCelda: esUltimo });
          });
          if (!ultimoEsParrafo) parrafo({ tipo: 'parrafo', hijos: [] }, { itap, finCelda: true });
        });
        // fin de fila: una marca con las propiedades de la tabla. Una celda combinada en horizontal es una sola celda ancha (así la escribe Word)
        const s = new Sprms();
        s.s1(0x2416, 1);
        if (itap === 1) s.s1(0x2417, 1);
        s.s4(0x6649, itap);
        if (itap > 1) { s.s1(0x244B, 1); s.s1(0x244C, 1); }
        s.s2(0x9602, 108);   // sprmTDxaGapHalf: la mitad del espacio entre columnas
        const celdasTotal = Math.min(f.celdas.length, 63);
        const centros = [izq];
        let col = 0;
        const anchoDe = (desde, span) => { let w = 0; for (let k = desde; k < desde + span; k++) w += k < anchos.length ? anchos[k] : 1440; return w; };
        f.celdas.slice(0, celdasTotal).forEach((c) => { const span = Math.max(1, Math.min(63, c.fusionH || 1)); centros.push(centros[centros.length - 1] + anchoDe(col, span)); col += span; });
        const def = new Bytes(3 + 22 * celdasTotal);
        def.u8(celdasTotal);
        centros.forEach((v) => def.u16(entero(v, -32768, 32767) & 0xFFFF));
        const rellenos = [];
        f.celdas.slice(0, celdasTotal).forEach((c, k) => {
          let fl = 0;
          if (c.fusionV === 'inicio') fl |= (3 << 5); else if (c.fusionV === 'continua') fl |= (1 << 5);
          fl |= ((c.alineacionV === 'centro' ? 1 : c.alineacionV === 'abajo' ? 2 : 0) << 7) | (3 << 9);
          def.u16(fl).u16(entero(centros[k + 1] - centros[k], 0, 31680));
          const bd = c.bordes || {};
          for (const lado of ['arriba', 'izquierda', 'abajo', 'derecha']) def.bytes(brc80(bd[lado]));
          rellenos[k] = c.relleno && rgb(c.relleno) ? c.relleno : null;
        });
        s.sw(0xD608, def.salida());
        if (bordeTabla) {
          const b = new Bytes(48);
          for (let k = 0; k < 6; k++) b.bytes(brc8({ tam: bordeTabla.tam, color: bordeTabla.color }, 0));
          s.sv(0xD613, b.salida());
        }
        if (rellenos.some(Boolean)) {
          for (const [desde, op, maxc] of [[0, 0xD612, 22], [22, 0xD616, 22], [44, 0xD60C, 19]]) {
            if (celdasTotal <= desde) break;
            const hasta = Math.min(celdasTotal, desde + maxc);
            const sb = new Bytes(10 * (hasta - desde));
            for (let k = desde; k < hasta; k++) {
              const c = rellenos[k] ? rgb(rellenos[k]) : null;
              if (c) sb.u32(0xFF000000).u32(c[0] | (c[1] << 8) | (c[2] << 16)).u16(0); else sb.u32(0xFFFFFFFF).u32(0xFFFFFFFF).u16(0xFFFF);   // trama «clara» con el color en el fondo, como Word
            }
            s.sv(op, sb.salida());
          }
        }
        if (f.alto) s.s2(0x9407, entero(f.alto, 1, 31680));
        if (f.encabezado) s.s1(0x3404, 1);
        finParrafo(itap === 1 ? 7 : 13, 0, s);
      }
      return true;
    }

    function bloques(lista) {
      let ultimoEsTabla = false;
      for (const b of lista || []) {
        if (b.tipo === 'tabla') { ultimoEsTabla = tabla(b, 1); }
        else { parrafo(b, null); ultimoEsTabla = false; }
      }
      return ultimoEsTabla;
    }

    // ───────── Secciones ─────────
    const secciones = seccionesHasta(contrato.secciones, MAX_SECCIONES);
    const cpsSeccion = [0];
    secciones.forEach((s, i) => {
      const hijos = s.hijos && s.hijos.length ? s.hijos : [{ tipo: 'parrafo', hijos: [] }];
      const ultimo = hijos[hijos.length - 1];
      // la marca de sección (0x0C) cierra el último párrafo de la sección, como en Word; si la sección termina en tabla, va en un párrafo propio
      const cierraParrafo = i < secciones.length - 1 && ultimo && ultimo.tipo !== 'tabla';
      const terminaEnTabla = bloques(cierraParrafo ? hijos.slice(0, -1) : hijos);
      if (terminaEnTabla) parrafo({ tipo: 'parrafo', hijos: [] }, null);
      if (cierraParrafo) parrafo(ultimo, null, 12);
      else if (i < secciones.length - 1) finParrafo(12, 0, null);
      cpsSeccion.push(D.n);
    });
    const ccpText = D.n;

    // Encabezados y pies: cada sección tiene 6 historias (par, impar/predeterminado, pie par, pie impar/predeterminado, primera página); solo se llenan las predeterminadas
    // y solo cuando la sección trae las suyas (la que hereda las de la anterior deja las suyas vacías)
    const hayCabeceras = secciones.some((s) => (s.encabezado && s.encabezado.length) || (s.pie && s.pie.length));
    const hdd = [];
    let ccpHdd = 0;
    historiaActual = 'encabezados';
    if (hayCabeceras) {
      inicioHistoria = D.n;
      // las seis historias de separadores de notas (vacías de contenido, pero Word las espera con su marca)
      const separador = (c) => { D.car(c); D.tramo(1, null, ''); D.car(13); D.tramo(1, null, ''); D.pars.push({ fin: D.n, istd: 0, g: null }); D.car(13); D.tramo(1, null, ''); D.pars.push({ fin: D.n, istd: 0, g: null }); };
      for (const c of [3, 4, 0, 3, 4, 0]) { hdd.push(D.n - inicioHistoria); if (c) separador(c); }
      let previoEnc = null, previoPie = null;
      secciones.forEach((s) => {
        const historia = (lista, previo) => {
          hdd.push(D.n - inicioHistoria);
          if (!lista || !lista.length || lista === previo) return;
          const terminaEnTabla = bloques(lista);
          if (terminaEnTabla) parrafo({ tipo: 'parrafo', hijos: [] }, null);
          D.car(13); D.tramo(1, null, ''); D.pars.push({ fin: D.n, istd: 0, g: null });   // marca final de la historia
        };
        hdd.push(D.n - inicioHistoria);                                   // encabezado de páginas pares
        historia(s.encabezado, previoEnc);                                // predeterminado
        historia(null);                                                   // pie de páginas pares
        historia(s.pie, previoPie);                                       // predeterminado
        hdd.push(D.n - inicioHistoria); hdd.push(D.n - inicioHistoria);  // primera página
        if (s.encabezado && s.encabezado.length) previoEnc = s.encabezado;
        if (s.pie && s.pie.length) previoPie = s.pie;
      });
      hdd.push(D.n - inicioHistoria);
      // dos marcas más después de todas las historias, como las deja Word; ccpHdd no cuenta la última
      for (let k = 0; k < 2; k++) { D.car(13); D.tramo(1, null, ''); D.pars.push({ fin: D.n, istd: 0, g: null }); }
      ccpHdd = D.n - ccpText - 1;
      hdd.push(D.n - inicioHistoria + 1);
    }
    const totalCp = D.n;

    // ───────── FKP ─────────
    const fcDe = (cp) => FC_TEXTO + 2 * cp;
    const wd = new Bytes(65536);
    wd.ceros(FC_TEXTO);
    for (let i = 0; i < D.n; i++) wd.u16(D.T[i]);
    wd.alinear(512);

    // Una página de formatos de carácter (CHPX) con todos los tramos que quepan
    function paginasChp() {
      const paginas = [];
      let i = 0;
      const lista = D.runs;
      let ini = 0;
      while (i < lista.length) {
        const rgfc = [], bs = [], contenido = [];
        let pos = 510, usado = new Map(), n = 0;
        const fcIni = fcDe(ini);
        rgfc.push(fcIni);
        while (i < lista.length) {
          const r = lista[i];
          // espacio que ocupa añadir este tramo: 4 (rgfc) + 1 (rgb) y, si trae propiedades nuevas, 1 + largo
          let nuevoPos = pos, off = 0;
          if (r.g) {
            if (usado.has(r.clave)) off = usado.get(r.clave);
            else { nuevoPos = (pos - (1 + r.g.length)) & ~1; off = nuevoPos; }
          }
          const cabecera = 4 * (n + 2) + (n + 1) + 1;
          if (cabecera > nuevoPos || (r.g && 1 + r.g.length > 255)) { if (n === 0) { r.g = null; r.clave = ''; continue; } break; }
          if (r.g && !usado.has(r.clave)) { usado.set(r.clave, nuevoPos); contenido.push({ off: nuevoPos, g: r.g }); pos = nuevoPos; off = nuevoPos; }
          bs.push(r.g ? off / 2 : 0);
          rgfc.push(fcDe(r.fin));
          ini = r.fin; n++; i++;
        }
        const pg = new Uint8Array(512);
        const dv = new DataView(pg.buffer);
        rgfc.forEach((v, k) => dv.setUint32(4 * k, v, true));
        bs.forEach((v, k) => { pg[4 * (n + 1) + k] = v; });
        for (const c of contenido) { pg[c.off] = c.g.length; pg.set(c.g, c.off + 1); }
        pg[511] = n;
        paginas.push({ bytes: pg, fcIni, fcFin: rgfc[rgfc.length - 1] });
      }
      return paginas;
    }

    // Páginas de formatos de párrafo (PAPX): un registro por párrafo
    function paginasPap() {
      const paginas = [];
      const lista = D.pars;
      let i = 0, ini = 0;
      while (i < lista.length) {
        const rgfc = [], bxs = [], contenido = [];
        let pos = 510, n = 0;
        const fcIni = fcDe(ini);
        rgfc.push(fcIni);
        while (i < lista.length) {
          const p = lista[i];
          let g = p.g || new Uint8Array(0);
          // las propiedades que no caben en la página (una fila con muchas celdas) van al flujo Data (sprmPHugePapx)
          if (2 + g.length > 480) {
            data.alinear(4);
            const o = data.n;
            data.u16(g.length).bytes(g);
            const s = new Sprms(); s.s4(0x6646, o);
            g = s.bytes();
          }
          const L = 2 + g.length;
          const total = L % 2 === 1 ? 1 + L : 2 + L;   // con cb de una palabra (L impar) o con el byte de relleno y cb (L par)
          const nuevoPos = (pos - total) & ~1;
          const cabecera = 4 * (n + 2) + 13 * (n + 1) + 1;
          if (cabecera > nuevoPos) { if (n === 0) throw new Error('Un párrafo tiene demasiadas propiedades para guardarlo como .doc.'); break; }
          pos = nuevoPos;
          const c = new Bytes(total);
          if (L % 2 === 1) c.u8((L + 1) / 2); else c.u8(0).u8(L / 2);
          c.u16(p.istd).bytes(g);
          contenido.push({ off: nuevoPos, b: c.salida() });
          bxs.push(nuevoPos / 2);
          rgfc.push(fcDe(p.fin));
          ini = p.fin; n++; i++;
        }
        const pg = new Uint8Array(512);
        const dv = new DataView(pg.buffer);
        rgfc.forEach((v, k) => dv.setUint32(4 * k, v, true));
        bxs.forEach((v, k) => { pg[4 * (n + 1) + 13 * k] = v; });
        for (const c of contenido) pg.set(c.b, c.off);
        pg[511] = n;
        paginas.push({ bytes: pg, fcIni, fcFin: rgfc[rgfc.length - 1] });
      }
      return paginas;
    }

    const pagChp = paginasChp(), pagPap = paginasPap();
    const pnChp = [], pnPap = [];
    for (const p of pagChp) { pnChp.push(wd.n / 512); wd.bytes(p.bytes); }
    for (const p of pagPap) { pnPap.push(wd.n / 512); wd.bytes(p.bytes); }

    // ───────── Secciones: propiedades (SEPX) ─────────
    const sepx = secciones.map((s, i) => {
      const g = new Sprms();
      const pg = s.pagina || {}, m = s.margenes || {};
      g.s1(0x3009, i > 0 && s.tipo === 'continua' ? 0 : 2);
      g.s2(0xB01F, entero(pg.ancho || 12240, 720, 31680)); g.s2(0xB020, entero(pg.alto || 15840, 720, 31680));
      if (pg.horizontal) g.s1(0x301D, 2);
      g.s2(0xB021, entero(m.izquierda == null ? 1440 : m.izquierda, 0, 31680)); g.s2(0xB022, entero(m.derecha == null ? 1440 : m.derecha, 0, 31680));
      g.s2(0x9023, entero(m.arriba == null ? 1440 : m.arriba, 0, 31680)); g.s2(0x9024, entero(m.abajo == null ? 1440 : m.abajo, 0, 31680));
      g.s2(0xB017, entero(m.encabezado == null ? 708 : m.encabezado, 0, 31680)); g.s2(0xB018, entero(m.pie == null ? 708 : m.pie, 0, 31680));
      if (s.columnas && s.columnas.n > 1) { g.s2(0x500B, entero(s.columnas.n, 2, 45) - 1); g.s2(0x900C, entero(s.columnas.espacio == null ? 720 : s.columnas.espacio, 0, 31680)); g.s1(0x3005, 1); }
      const b = g.bytes();
      wd.alinear(2);
      const fc = wd.n;
      wd.u16(b.length).bytes(b);
      return fc;
    });

    // ───────── Flujo de tablas (1Table) ─────────
    const tb = new Bytes(65536);
    const fcLcb = new Map();    // índice del FibRgFcLcb97 → [fc, lcb]
    const poner = (indice, fn) => { tb.alinear(4); const ini = tb.n; fn(); fcLcb.set(indice, [ini, tb.n - ini]); };

    // Hoja de estilos (STSH)
    poner(1, () => {
      // STSHI de 20 bytes: cstd, cbSTDBaseInFile, banderas, sti máximo, último istd fijo, versión de nombres y las fuentes predeterminadas (ASCII, asiática, otras, compleja)
      tb.u16(20).u16(filasStd.length).u16(10).u16(1).u16(0x0178).u16(15).u16(7).u16(fuenteBase).u16(fuenteAsiaBase).u16(fuenteBase).u16(fuenteCsBase);
      filasStd.forEach((e, istd) => {
        if (!e) { tb.u16(0); return; }
        const esCar = e.tipo === 'caracter';
        const nombre = istd === 0 ? 'Normal' : String(e.nombre || e.id).slice(0, 254);
        const nombreSti = String(nombre).toLowerCase();
        let sti = istd < 10 ? istd : (STI[nombreSti] != null && (STI_CARACTER.has(STI[nombreSti]) === esCar) ? STI[nombreSti] : 4094);
        // propiedades del estilo; Normal lleva además las predeterminadas del documento
        const rEst = Object.assign({}, istd === 0 ? pd.r : null, e.r);
        if (istd === 0) { rEst.fuente = rEst.fuente || 'Times New Roman'; if (rEst.tamano == null) rEst.tamano = 10; }
        const pEst = Object.assign({}, istd === 0 ? pd.p : null, e.p);
        delete pEst.lista;
        const papx = esCar ? null : sprmsPap(pEst, { ilfoDe: () => null }).bytes();
        const chpx = sprmsChp(rEst, null).bytes();
        const base = istd === 0 ? 0x0FFF : (e.basadoEn != null && istdDe.has(e.basadoEn) && istdDe.get(e.basadoEn) !== istd ? istdDe.get(e.basadoEn) : 0);
        const siguiente = e.siguiente != null && istdDe.has(e.siguiente) ? istdDe.get(e.siguiente) : istd;
        const std = new Bytes(64);
        std.u16(sti | 0x6000).u16((esCar ? 2 : 1) | (base << 4)).u16((esCar ? 1 : 2) | (siguiente << 4)).u16(0).u16(0x100C);
        // nombre: cch + caracteres + terminador
        std.u16(nombre.length); for (let k = 0; k < nombre.length; k++) std.u16(nombre.charCodeAt(k)); std.u16(0);
        const upx = (b) => { std.u16(b.length).bytes(b); if (b.length % 2) std.u8(0); };
        if (esCar) upx(chpx);
        else { const p = new Bytes(papx.length + 2); p.u16(istd).bytes(papx); upx(p.salida()); upx(chpx); }
        std.poner16(6, std.n);   // bchUpe: el tamaño del STD
        tb.u16(std.n).bytes(std.salida());
      });
    });

    // PlcfSed: posiciones de las secciones y dónde están sus propiedades
    poner(6, () => {
      const finTotal = totalCp;
      cpsSeccion.forEach((c, k) => tb.u32(k === cpsSeccion.length - 1 ? finTotal : c));
      sepx.forEach((fc) => { tb.u16(0).u32(fc).u16(0).u32(0xFFFFFFFF); });
    });

    // PlcfHdd: dónde empieza cada historia de encabezados y pies
    if (hdd.length) poner(11, () => hdd.forEach((c) => tb.u32(c)));

    // Referencias de formatos
    poner(12, () => { pagChp.forEach((p) => tb.u32(p.fcIni)); tb.u32(pagChp[pagChp.length - 1].fcFin); pnChp.forEach((n) => tb.u32(n)); });
    poner(13, () => { pagPap.forEach((p) => tb.u32(p.fcIni)); tb.u32(pagPap[pagPap.length - 1].fcFin); pnPap.forEach((n) => tb.u32(n)); });

    // Campos (PlcfFld): posición de cada marca y su tipo
    const escribirCampos = (indice, lista, fin) => {
      if (!lista.length) return;
      poner(indice, () => {
        lista.forEach((c) => tb.u32(c.cp));
        tb.u32(fin);
        lista.forEach((c) => { tb.u8(c.ch); tb.u8(c.flags); });
      });
    };
    escribirCampos(16, camposPorHistoria.principal, ccpText);
    escribirCampos(17, camposPorHistoria.encabezados, D.n - ccpText);

    // Marcadores
    if (D.marcadores.length) {
      // sin nombre vacío ni repetido (Word no los admite; vale el primero)
      const vistos = new Set();
      const ms = D.marcadores.filter((m) => m.cp < ccpText && m.nombre && !vistos.has(m.nombre) && vistos.add(m.nombre)).slice(0, 4000).sort((a, b) => a.cp - b.cp);
      if (ms.length) {
        poner(21, () => { tb.u16(0xFFFF).u16(ms.length).u16(0); ms.forEach((m) => { tb.u16(m.nombre.length); for (let k = 0; k < m.nombre.length; k++) tb.u16(m.nombre.charCodeAt(k)); }); });
        poner(22, () => { ms.forEach((m) => tb.u32(m.cp)); tb.u32(totalCp + 2); ms.forEach((m, k) => { tb.u16(k); tb.u16(0); }); });
        poner(23, () => { ms.forEach((m) => tb.u32(m.cp)); tb.u32(totalCp + 2); });
      }
    }

    // Dibujos (solo con imágenes)
    if (hayImagenes) poner(50, () => tb.bytes(DGG_INFO));

    // Fuentes
    poner(15, () => {
      tb.u16(fuentes.length).u16(0);
      for (const n of fuentes) {
        const nombre = utf16(n);
        const cb = 39 + nombre.length;
        tb.u8(cb).u8(0x06).u16(400).u8(0).u8(0).ceros(10).ceros(24).bytes(nombre);
      }
    });

    // Propiedades del documento (Dop)
    poner(31, () => {
      const dop = new Uint8Array(500);
      dop.set([0x22, 0x00, 0x04, 0x00, 0x31, 0x08, 0x80, 0x18, 0x00, 0xF0, 0xD0, 0x02, 0x00, 0x00, 0xA9, 0x01, 0x00, 0x00, 0x00, 0x00, 0x8C, 0xB5, 0xEA, 0x87, 0x8C, 0xB5, 0xEA, 0x87, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x3E, 0x00, 0x00, 0x00, 0x59, 0x01, 0x00, 0x00, 0x02, 0x00, 0x01, 0x00, 0x00, 0x00, 0x04, 0x00, 0x03, 0x90, 0x02, 0x00, 0x00, 0x00, 0x3E, 0x00, 0x00, 0x00, 0x59, 0x01, 0x00, 0x00, 0x02, 0x00, 0x01, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]);   // DopBase de Word, con las opciones habituales del documento
      tb.bytes(dop);
    });

    // Tabla de piezas (Clx): una sola pieza Unicode con todo el texto
    poner(33, () => {
      tb.u8(2).u32(4 * 2 + 8);
      tb.u32(0).u32(totalCp);
      tb.u16(0).u32(FC_TEXTO).u16(0);
    });

    // Listas: PlfLst (las definiciones, seguidas de sus niveles) y PlfLfo (cada lista en uso)
    if (listas.length) {
      tb.alinear(4);
      const ini = tb.n;
      tb.u16(listas.length);
      listas.forEach((l, i) => {
        tb.u32(1000 + i).u32(0x00000000);
        for (let k = 0; k < 9; k++) tb.u16(0x0FFF);
        tb.u8(0x00).u8(0);
      });
      fcLcb.set(73, [ini, tb.n - ini]);
      listas.forEach((l) => { for (let n = 0; n < 9; n++) nivelLista(tb, (l.niveles || [])[n], n); });
      poner(74, () => {
        tb.u32(listas.length);
        listas.forEach((l, i) => { tb.u32(1000 + i).ceros(8).u8(0).u8(0).u8(0).u8(0); });
        listas.forEach(() => { tb.u32(0xFFFFFFFF); });
      });
    }
    function nivelLista(b, v, n) {
      v = v || { formato: 'decimal', texto: `%${n + 1}.`, izquierda: 720 * (n + 1), colgante: 360 };
      const nfc = NFC[v.formato] != null ? NFC[v.formato] : 0;
      // texto: los %1..%9 son los números de cada nivel; en el flujo van como el carácter 0..8 y rgbxchNums dice dónde está cada uno
      let texto = String(v.texto == null ? '' : v.texto);
      if (nfc === 23) texto = texto.slice(0, 1) || '•';
      const xst = [];
      const pos = new Array(9).fill(0);
      for (let i = 0; i < texto.length && xst.length < 255; i++) {
        const m = nfc !== 23 && texto[i] === '%' && /[1-9]/.test(texto[i + 1] || '') ? +texto[i + 1] - 1 : -1;
        if (m >= 0) { if (!pos[m]) pos[m] = xst.length + 1; xst.push(m); i++; } else xst.push(texto.charCodeAt(i));
      }
      const ps = new Sprms();
      ps.s2(0x845E, entero(v.izquierda, 0, 31680)); ps.s2(0x8460, -entero(v.colgante, 0, 31680));
      const cs = sprmsChp(Object.assign({}, v.r, v.fuente ? { fuente: v.fuente } : null), null);
      const papx = ps.bytes(), chpx = cs.bytes();
      b.u32(entero(v.inicio == null ? 1 : v.inicio, 0, 32767, 1)).u8(nfc).u8(0);
      for (let k = 0; k < 9; k++) b.u8(pos[k]);
      b.u8(0).u32(0).u32(0).u8(chpx.length).u8(papx.length).u8(0).u8(0);
      b.bytes(papx).bytes(chpx);
      b.u16(xst.length); xst.forEach((c) => b.u16(c));
    }

    // ───────── Armado del FIB ─────────
    const tablaBytes = tb.salida();
    const fib = new Bytes(1700);
    fib.u16(0xA5EC).u16(0x00C1).u16(0x0019).u16(0x0409).u16(0).u16(0x12F0 | (hayImagenes ? 0x08 : 0)).u16(0x00BF).u32(0).u8(0).u8(0x10).u16(0).u16(0).u32(FC_TEXTO).u32(FC_TEXTO + 2 * totalCp);
    fib.u16(14).u16(0x6A62).u16(0x6A62).u16(0xD7D6).u16(0xD7D6);
    for (let k = 0; k < 9; k++) fib.u16(0);
    fib.u16(0x0409);
    fib.u16(22).u32(0).u32(0).u32(0);                                   // cslw, cbMac (luego), reserved1, reserved2
    fib.u32(ccpText).u32(0).u32(ccpHdd).u32(0).u32(0).u32(0).u32(0).u32(0);   // ccpText, ccpFtn, ccpHdd, reserved3, ccpAtn, ccpEdn, ccpTxbx, ccpHdrTxbx
    while (fib.n < 0x98) fib.u32(0);
    // 183 pares (la tabla de Word 2007) y, detrás, la versión real del programa (nFibNew 0x0112): sin ella Word 2000+ trata el archivo como Word 97 y no entiende las tablas anidadas
    fib.u16(183);
    for (let i = 0; i < 183; i++) { const par = fcLcb.get(i) || [0, 0]; fib.u32(par[0]).u32(par[1]); }
    fib.u16(5).u16(0x0112).u16(0).u16(0x0409).u16(0x0409).u16(0x0409);
    wd.alinear(512);
    while (wd.n < 4096) wd.u8(0);
    const salidaWd = wd.salida();
    salidaWd.set(fib.salida(), 0);
    new DataView(salidaWd.buffer).setUint32(0x40, salidaWd.length, true);

    const flujos = [{ nombre: 'WordDocument', bytes: salidaWd }, { nombre: '1Table', bytes: relleno4096(tablaBytes) }];
    if (data.n) flujos.push({ nombre: 'Data', bytes: relleno4096(data.salida()) });
    return { bytes: CfbPropio.escribir(flujos), mensajes };
  }

  // Más de max secciones (solo un archivo raro): el contenido de las de más pasa a la última, para no perder texto
  function seccionesHasta(lista, max) {
    const s = lista && lista.length ? lista : [{ hijos: [] }];
    if (s.length <= max) return s;
    const ultima = Object.assign({}, s[max - 1], { hijos: s.slice(max - 1).flatMap((x) => x.hijos || []) });
    return s.slice(0, max - 1).concat([ultima]);
  }

  function relleno4096(b) {
    if (b.length >= 4096) return b;
    const r = new Uint8Array(4096); r.set(b); return r;
  }

  globalThis.DocEscritor = { generar };
})();

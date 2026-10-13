/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Lector de libros de Excel (.xlsx y .xlsm) del motor propio: el paquete OOXML
// (zip propio + relaciones) → un modelo resuelto que el dibujo, el formato de
// número y las herramientas usan sin saber de XML.
//
// Requiere: zip-propio.js, xml-ligero.js, xlsx-estilos.js. No depende del
// navegador (corre igual en Node, que es como se prueba contra SheetJS y Excel).
//
// Uso:
//   const libro = await LibroXlsx.abrir(archivoOBytes);        // lee el paquete, los textos, los estilos y la lista de hojas
//   libro.hojas[i]  → { nombre, id, estado: 'visible'|'oculta'|'muyOculta', tipo: 'hoja'|'grafico'|…, ruta }
//   const hoja = await libro.cargarHoja(i);                    // lee las celdas de esa hoja (una vez)
//   hoja.celdas.obtener(fila, col)  → { tipo, valor, estilo, formula } | null   (fila y columna desde 0)
//   hoja.celdas.paraCada((fila, col, i) => …)
//   libro.estilos.xf(hoja.celdas.estilo[i]) → estilo resuelto (ver xlsx-estilos.js)
//
// Los textos que lleva Excel con caracteres de control (_x000D_) vuelven resueltos. Las fechas se
// guardan como número de serie (con libro.fecha1904 para saber la base). Lo que este lector no
// interpreta (gráficos, dibujos, tablas dinámicas…) se anota en hoja.partes para las etapas siguientes
// y se conserva intacto en el zip.

(function () {
  const TIPO = { vacio: 0, numero: 1, texto: 2, logico: 3, error: 4 };

  const TEXTOS_ERROR = {
    noZip: { es: 'el archivo no es un libro de Excel (.xlsx) válido', en: 'the file is not a valid Excel (.xlsx) workbook' },
    cifrado: { es: 'el libro está protegido con contraseña; quítale la contraseña en Excel y vuelve a intentarlo', en: 'the workbook is password-protected; remove the password in Excel and try again' },
    antiguo: { es: 'es un libro en formato antiguo (.xls) o binario', en: 'this is a workbook in an old (.xls) or binary format' },
    sinLibro: { es: 'el archivo no contiene un libro de Excel', en: 'the file does not contain an Excel workbook' },
    danado: { es: 'el libro está dañado', en: 'the workbook is damaged' },
    grande: { es: 'el libro es demasiado grande para abrirlo en el navegador', en: 'the workbook is too large to open in the browser' },
  };

  // Límites de Excel (filas y columnas) y topes propios contra archivos que, chicos comprimidos, agotarían la memoria.
  const MAX_FILAS = 1048576, MAX_COLUMNAS = 16384;
  const MAX_CELDAS = 25000000;      // ~600 MB de memoria; un libro real de 2,5 M celdas ocupa ~80 MB
  const MAX_ELEMENTOS = 500000;     // combinadas, hipervínculos, reglas y validaciones por hoja

  class ErrorXlsx extends Error {
    constructor(codigo, detalle) {
      const t = TEXTOS_ERROR[codigo] || TEXTOS_ERROR.danado;
      const en = typeof globalThis.idiomaActual === 'function' && globalThis.idiomaActual() === 'en';
      super(en ? t.en : t.es);
      this.name = 'ErrorXlsx';
      this.codigo = codigo;
      this.detalle = detalle || '';
    }
  }

  // ---- Direcciones ----------------------------------------------------------------------------------

  function colANumero(letras) {
    let n = 0;
    for (let i = 0; i < letras.length; i++) n = n * 26 + (letras.charCodeAt(i) & 31);
    return n - 1;
  }
  function numeroACol(n) {
    let s = '';
    for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    return s;
  }
  // 'B3' → { f: 2, c: 1 }; '$B$3' también. Devuelve null si no es una dirección.
  function decodificarDireccion(ref) {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref);
    return m ? { f: parseInt(m[2], 10) - 1, c: colANumero(m[1].toUpperCase()) } : null;
  }
  function codificarDireccion(f, c) { return numeroACol(c) + (f + 1); }
  // 'A1:C3' → { f0, c0, f1, c1 }; 'A1' → rango de una celda.
  function decodificarRango(ref) {
    const partes = String(ref).split(':');
    const a = decodificarDireccion(partes[0]);
    const b = partes.length > 1 ? decodificarDireccion(partes[1]) : a;
    return a && b ? { f0: Math.min(a.f, b.f), c0: Math.min(a.c, b.c), f1: Math.max(a.f, b.f), c1: Math.max(a.c, b.c) } : null;
  }

  // ---- Fechas ---------------------------------------------------------------------------------------

  // '2023-03-15T12:30:00' → número de serie de Excel (con el falso 29-feb-1900 si es base 1900).
  function serialDeIso(texto, fecha1904) {
    const m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?)?/.exec(texto.trim());
    if (!m) return null;
    const dias = Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000;
    const fraccion = ((+m[4] || 0) * 3600 + (+m[5] || 0) * 60 + (+m[6] || 0) + (m[7] ? +('0.' + m[7]) : 0)) / 86400;
    if (fecha1904) return dias + 24107 + fraccion; // 1904-01-01 = serial 0
    // 1970-01-01 = serial 25569 (ya cuenta el 29-feb-1900 que no existió); antes del 1-mar-1900 ese día no está.
    let serial = dias + 25569;
    if (serial < 61) serial -= 1;
    return serial + fraccion;
  }

  // ---- Rutas y relaciones ---------------------------------------------------------------------------

  function resolverRuta(rutaParte, destino) {
    let limpio = destino;
    try { limpio = decodeURI(destino); } catch (e) { /* se usa tal cual */ }
    if (limpio.startsWith('/')) return limpio.slice(1);
    const partes = [];
    (rutaParte.slice(0, rutaParte.lastIndexOf('/') + 1) + limpio).split('/').forEach((s) => {
      if (s === '..') partes.pop(); else if (s !== '.' && s !== '') partes.push(s);
    });
    return partes.join('/');
  }
  const rutaRels = (ruta) => ruta.slice(0, ruta.lastIndexOf('/') + 1) + '_rels/' + ruta.slice(ruta.lastIndexOf('/') + 1) + '.rels';

  // Los nombres de las partes no distinguen mayúsculas (OPC): un índice en minúsculas por zip, armado una vez.
  const indicesZip = new WeakMap();
  function nombreReal(zip, ruta) {
    if (zip.file(ruta)) return ruta;
    let indice = indicesZip.get(zip);
    if (!indice) {
      indice = new Map();
      zip.forEach((nombre) => { const k = nombre.toLowerCase(); if (!indice.has(k)) indice.set(k, nombre); });
      indicesZip.set(zip, indice);
    }
    return indice.get(ruta.toLowerCase()) || null;
  }
  async function textoDe(zip, ruta) {
    const real = nombreReal(zip, ruta);
    return real ? zip.file(real).async('string') : null;
  }
  async function buscarEntrada(zip, ruta) { return nombreReal(zip, ruta); }

  // [{ id, tipo (último segmento del Type), destino (ruta resuelta o URL si es externa), externa }]
  async function leerRelaciones(zip, rutaParte) {
    const texto = await textoDe(zip, rutaRels(rutaParte));
    const lista = [];
    if (!texto) return lista;
    XmlLigero.recorrer(texto, {
      abrir(nombre, a) {
        if (nombre !== 'Relationship' || !a.Id) return;
        const externa = a.TargetMode === 'External';
        lista.push({ id: a.Id, tipo: (a.Type || '').split('/').pop(), tipoCompleto: a.Type || '', externa, destino: externa ? (a.Target || '') : resolverRuta(rutaParte, a.Target || '') });
      },
    });
    return lista;
  }

  // ---- Textos compartidos ---------------------------------------------------------------------------

  // [{ texto, ricos: null | [{ texto, fuente }] }]
  function leerTextosCompartidos(texto) {
    const lista = [];
    if (!texto) return lista;
    let actual = null, enT = false, enRun = false, enFonetico = false, run = null, enRpr = false;
    XmlLigero.recorrer(texto, {
      abrir(nombre, a) {
        switch (nombre) {
          case 'si': actual = { texto: '', ricos: null }; break;
          case 'r': if (actual) { enRun = true; run = { texto: '', fuente: {} }; if (!actual.ricos) actual.ricos = []; } break;
          case 'rPh': enFonetico = true; break;
          case 'rPr': enRpr = true; break;
          case 't': enT = !enFonetico; break;
          default:
            if (enRpr && run) {
              if (nombre === 'b') run.fuente.negrita = a.val === undefined || a.val === '1' || a.val === 'true';
              else if (nombre === 'i') run.fuente.cursiva = a.val === undefined || a.val === '1' || a.val === 'true';
              else if (nombre === 'strike') run.fuente.tachado = a.val === undefined || a.val === '1' || a.val === 'true';
              else if (nombre === 'u') run.fuente.subrayado = a.val || 'single';
              else if (nombre === 'sz') run.fuente.tam = parseFloat(a.val);
              else if (nombre === 'rFont') run.fuente.nombre = a.val;
              else if (nombre === 'vertAlign') run.fuente.alinVertical = a.val;
              else if (nombre === 'color') run.fuente.color = XlsxEstilos.colorRef(a);
            }
        }
      },
      texto(t) {
        if (!enT || !actual) return;
        const limpio = XmlLigero.sinEscapes(t);
        if (enRun && run) run.texto += limpio; else actual.texto += limpio;
      },
      cerrar(nombre) {
        if (nombre === 't') enT = false;
        else if (nombre === 'rPh') enFonetico = false;
        else if (nombre === 'rPr') enRpr = false;
        else if (nombre === 'r' && actual && run) { actual.ricos.push(run); actual.texto += run.texto; run = null; enRun = false; }
        else if (nombre === 'si' && actual) { lista.push(actual); actual = null; }
      },
    });
    return lista;
  }

  // ---- Celdas de una hoja (columnas paralelas: poca memoria con millones de celdas) ----------------

  class CeldasHoja {
    constructor() {
      this.n = 0;
      this.fila = new Int32Array(1024);
      this.col = new Int32Array(1024);
      this.estilo = new Int32Array(1024);
      this.tipo = new Uint8Array(1024);
      this.valor = [];          // número | texto | boolean según tipo
      this.ricos = new Map();    // índice → [{ texto, fuente }] (texto con formato por tramos)
      this.textoOriginal = new Map(); // índice → texto tal como venía (CSV: "1.234,50")
      this.formulas = new Map(); // índice → { texto, tipo, ref, si, maestra:bool }
      this.compartidas = new Map(); // si → { texto, f, c } de la celda maestra
      this._inicioFila = null;
    }
    _crecer() {
      const m = this.fila.length * 2;
      const copiar = (a, T) => { const b = new T(m); b.set(a); return b; };
      this.fila = copiar(this.fila, Int32Array); this.col = copiar(this.col, Int32Array);
      this.estilo = copiar(this.estilo, Int32Array); this.tipo = copiar(this.tipo, Uint8Array);
    }
    agregar(f, c, tipo, valor, estilo) {
      if (this.n === this.fila.length) this._crecer();
      const i = this.n++;
      this.fila[i] = f; this.col[i] = c; this.tipo[i] = tipo; this.estilo[i] = estilo;
      this.valor[i] = valor;
      this._inicioFila = null;
      return i;
    }
    // Las hojas bien formadas traen las celdas en orden; si no, se ordenan una vez.
    _indexar() {
      if (this._inicioFila) return;
      let ordenado = true;
      for (let i = 1; i < this.n; i++) {
        if (this.fila[i] < this.fila[i - 1] || (this.fila[i] === this.fila[i - 1] && this.col[i] <= this.col[i - 1])) { ordenado = false; break; }
      }
      if (!ordenado) this._ordenar();
      const inicio = new Map();
      for (let i = 0; i < this.n; i++) if (!inicio.has(this.fila[i])) inicio.set(this.fila[i], i);
      this._inicioFila = inicio;
    }
    // Ordena por (fila, columna) y, si una celda aparece dos veces, deja la última del archivo (como Excel).
    // Todo lo que está indexado por posición (fórmulas, textos con formato, texto original del CSV) se reubica.
    _ordenar() {
      const idx = Array.from({ length: this.n }, (_, i) => i).sort((a, b) => this.fila[a] - this.fila[b] || this.col[a] - this.col[b] || a - b);
      const f = new Int32Array(this.fila.length), c = new Int32Array(this.col.length), e = new Int32Array(this.estilo.length), t = new Uint8Array(this.tipo.length);
      const v = [];
      const mapas = ['formulas', 'ricos', 'textoOriginal'];
      const nuevos = mapas.map(() => new Map());
      let nueva = 0;
      for (let k = 0; k < idx.length; k++) {
        const vieja = idx[k];
        const sig = idx[k + 1];
        if (sig !== undefined && this.fila[sig] === this.fila[vieja] && this.col[sig] === this.col[vieja]) continue; // duplicada: manda la última
        f[nueva] = this.fila[vieja]; c[nueva] = this.col[vieja]; e[nueva] = this.estilo[vieja]; t[nueva] = this.tipo[vieja]; v[nueva] = this.valor[vieja];
        mapas.forEach((nombre, m) => { if (this[nombre].has(vieja)) nuevos[m].set(nueva, this[nombre].get(vieja)); });
        nueva++;
      }
      mapas.forEach((nombre, m) => { this[nombre] = nuevos[m]; });
      this.fila = f; this.col = c; this.estilo = e; this.tipo = t; this.valor = v; this.n = nueva;
    }
    indiceDe(f, c) {
      this._indexar();
      const i = this._inicioFila.get(f);
      if (i === undefined) return -1;
      // búsqueda binaria dentro de la fila (las celdas están ordenadas por fila y columna): primera con (fila, col) >= (f, c)
      let lo = i, hi = this.n;
      while (lo < hi) {
        const m = (lo + hi) >> 1;
        if (this.fila[m] < f || (this.fila[m] === f && this.col[m] < c)) lo = m + 1; else hi = m;
      }
      return lo < this.n && this.fila[lo] === f && this.col[lo] === c ? lo : -1;
    }
    obtener(f, c) {
      const i = this.indiceDe(f, c);
      return i < 0 ? null : { tipo: this.tipo[i], valor: this.valor[i], estilo: this.estilo[i], formula: this.formulaDe(i) };
    }
    // Texto de la fórmula (sin "="); las compartidas se expanden desplazando las referencias relativas.
    formulaDe(i) {
      const fx = this.formulas.get(i);
      if (!fx) return null;
      if (fx.texto !== null) return fx.texto;
      const maestra = this.compartidas.get(fx.si);
      if (!maestra) return null;
      if (fx._memo === undefined) fx._memo = desplazarFormula(maestra.texto, this.fila[i] - maestra.f, this.col[i] - maestra.c);
      return fx._memo;
    }
    paraCada(fn) { for (let i = 0; i < this.n; i++) fn(this.fila[i], this.col[i], i); }
    get filaMax() { let m = -1; for (let i = 0; i < this.n; i++) if (this.fila[i] > m) m = this.fila[i]; return m; }
    get colMax() { let m = -1; for (let i = 0; i < this.n; i++) if (this.col[i] > m) m = this.col[i]; return m; }
  }

  // Desplaza las referencias relativas de una fórmula (para expandir las compartidas). Respeta los
  // textos entre comillas, los nombres de hoja entre apóstrofos y las referencias con $.
  function desplazarFormula(texto, df, dc) {
    if (!df && !dc) return texto;
    let salida = '';
    const n = texto.length;
    let i = 0;
    const esIdent = (ch) => /[A-Za-z0-9_.À-￿]/.test(ch);
    while (i < n) {
      const ch = texto[i];
      if (ch === '"') { // texto literal ("" es una comilla dentro)
        let j = i + 1;
        while (j < n) { if (texto[j] === '"') { if (texto[j + 1] === '"') j += 2; else break; } else j++; }
        salida += texto.slice(i, j + 1); i = j + 1; continue;
      }
      if (ch === "'") { // 'Nombre de hoja'!
        let j = i + 1;
        while (j < n) { if (texto[j] === "'") { if (texto[j + 1] === "'") j += 2; else break; } else j++; }
        salida += texto.slice(i, j + 1); i = j + 1; continue;
      }
      if (ch === '[') { // referencia estructurada de tabla: Tabla[[#This Row],[Col]]
        let nivel = 0, j = i;
        for (; j < n; j++) { if (texto[j] === '[') nivel++; else if (texto[j] === ']') { nivel--; if (nivel === 0) break; } }
        salida += texto.slice(i, j + 1); i = j + 1; continue;
      }
      if (/[$A-Za-z]/.test(ch) && (i === 0 || !esIdent(texto[i - 1]))) {
        const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(.!])/.exec(texto.slice(i, i + 16));
        if (m && colANumero(m[2].toUpperCase()) <= 16383) {
          let fila = parseInt(m[4], 10), col = colANumero(m[2].toUpperCase());
          if (!m[3]) fila += df;
          if (!m[1]) col += dc;
          if (fila < 1 || fila > MAX_FILAS || col < 0 || col > 16383) salida += '#REF!'; // se salió de la hoja
          else salida += m[1] + numeroACol(col) + m[3] + fila;
          i += m[0].length; continue;
        }
        // columnas enteras ($A:$C) y filas enteras ($1:$3)
        const mc = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(.!])/.exec(texto.slice(i, i + 16));
        if (mc) {
          const a = colANumero(mc[2].toUpperCase()) + (mc[1] ? 0 : dc), b = colANumero(mc[4].toUpperCase()) + (mc[3] ? 0 : dc);
          salida += (a < 0 || b < 0 || a > 16383 || b > 16383) ? '#REF!' : mc[1] + numeroACol(a) + ':' + mc[3] + numeroACol(b);
          i += mc[0].length; continue;
        }
      }
      if (/\d|\$/.test(ch) && (i === 0 || !esIdent(texto[i - 1]))) {
        const mf = /^(\$?)(\d+):(\$?)(\d+)(?![A-Za-z0-9_(.!])/.exec(texto.slice(i, i + 16));
        if (mf) {
          const a = parseInt(mf[2], 10) + (mf[1] ? 0 : df), b = parseInt(mf[4], 10) + (mf[3] ? 0 : df);
          salida += (a < 1 || b < 1 || a > MAX_FILAS || b > MAX_FILAS) ? '#REF!' : mf[1] + a + ':' + mf[3] + b;
          i += mf[0].length; continue;
        }
        // un número cualquiera (1.5E+3, 100): se copia entero para no tomar su "E3" por una referencia
        const mn = /^\d+(\.\d+)?([eE][+-]?\d+)?/.exec(texto.slice(i, i + 32));
        if (mn) { salida += mn[0]; i += mn[0].length; continue; }
      }
      salida += ch; i++;
    }
    return salida;
  }

  // ---- Una hoja --------------------------------------------------------------------------------------

  function leerHoja(texto, ctx) {
    const hoja = {
      celdas: new CeldasHoja(),
      dimension: null,             // la que declara el archivo (no siempre es cierta)
      combinadas: [],              // { f0, c0, f1, c1 }
      columnas: [],                // { min, max (desde 0), ancho, oculta, personalizado, estilo, nivel, colapsada, ajusteAuto }
      filas: new Map(),            // fila (desde 0) → { alto, oculta, personalizada, estilo, formatoPersonalizado, nivel, colapsada }
      formato: { anchoBase: 8, anchoPorDefecto: null, altoPorDefecto: null, altoPersonalizado: false, filasOcultasPorDefecto: false, nivelFilaMax: 0, nivelColMax: 0 },
      vista: { cuadricula: true, encabezados: true, zoom: 100, derechaAIzquierda: false, tipo: 'normal', paneles: null, seleccion: null, mostrarFormulas: false, ceros: true, esquema: true },
      color: null,                 // color de pestaña
      hipervinculos: [],           // { f0, c0, f1, c1, destino, ubicacion, visible, ayuda, externo }
      imprimir: { margenes: null, pagina: null, opciones: null, encabezadoPie: null, saltosFila: [], saltosCol: [], ajustarAPagina: false, autoSaltos: true },
      filtroAuto: null,
      proteccion: null,
      formatosCondicionales: [],   // { rangos: [rango], reglas: [{ tipo, prioridad, dxf, … }] } — lectura completa en la etapa 6
      validaciones: [],            // { rangos, tipo, … }
      partes: { dibujo: null, dibujoVml: null, tablas: [], comentarios: null, otras: [] },
      relaciones: [],
    };
    const C = hoja.celdas;
    let fila = -1, colSiguiente = 0;
    let celda = null;              // { f, c, tipo, estilo, v, formula, enIs }
    let enV = false, enF = false, enIs = false, enT = false, enFonetico = false;
    let textoCelda = '';
    let seccion = '';
    let sheetView = true;
    let enVistasPersonalizadas = false;
    let cf = null, regla = null, enFormulaCf = false, validacion = null, enFormulaVal = null, hf = null, enHf = null;
    const cerrarCelda = () => {
      if (!celda) return;
      // Fuera de la hoja (o sin sentido): se descarta. Un tope total evita que un archivo chico agote la memoria.
      if (celda.f < 0 || celda.f >= MAX_FILAS || celda.c < 0 || celda.c >= MAX_COLUMNAS) { celda = null; return; }
      if (C.n >= MAX_CELDAS) throw new ErrorXlsx('grande', 'más de ' + MAX_CELDAS + ' celdas');
      let tipo = TIPO.vacio, valor = null;
      const t = celda.t;
      if (t === 'inlineStr') { tipo = TIPO.texto; valor = XmlLigero.sinEscapes(celda.is); }
      else if (celda.v !== null) {
        const v = celda.v;
        if (t === 's') { const s = ctx.textos[parseInt(v, 10)]; tipo = s ? TIPO.texto : TIPO.vacio; valor = s ? s.texto : null; if (s && s.ricos) celda.rico = s.ricos; }
        else if (t === 'str') { tipo = TIPO.texto; valor = XmlLigero.sinEscapes(v); }
        else if (t === 'b') { tipo = TIPO.logico; valor = v.trim() === '1' || v.trim() === 'true'; }
        else if (t === 'e') { tipo = TIPO.error; valor = v.trim(); }
        else if (t === 'd') { const n = serialDeIso(v, ctx.fecha1904); tipo = n === null ? TIPO.texto : TIPO.numero; valor = n === null ? v : n; }
        else { // 'n' o sin tipo
          if (v.trim() === '') { tipo = TIPO.vacio; valor = null; } else { const n = Number(v); tipo = Number.isNaN(n) ? TIPO.texto : TIPO.numero; valor = Number.isNaN(n) ? v : n; }
        }
      }
      const i = C.agregar(celda.f, celda.c, tipo, valor, celda.estilo);
      if (celda.rico) C.ricos.set(i, celda.rico);
      if (celda.formula) {
        const fx = celda.formula;
        C.formulas.set(i, fx);
        if (fx.tipo === 'shared' && fx.texto !== null && fx.si !== null) { C.compartidas.set(fx.si, { texto: fx.texto, f: celda.f, c: celda.c }); }
      }
      celda = null;
    };

    XmlLigero.recorrer(texto, {
      abrir(nombre, a, vacio) {
        if (nombre === 'customSheetViews') enVistasPersonalizadas = true;
        if (enVistasPersonalizadas) return;
        switch (nombre) {
          case 'sheetPr': seccion = 'sheetPr'; break;
          case 'tabColor': hoja.color = XlsxEstilos.colorRef(a); break;
          case 'pageSetUpPr': hoja.imprimir.ajustarAPagina = a.fitToPage === '1' || a.fitToPage === 'true'; if (a.autoPageBreaks === '0' || a.autoPageBreaks === 'false') hoja.imprimir.autoSaltos = false; break;
          case 'dimension': hoja.dimension = a.ref ? decodificarRango(a.ref) : null; break;
          case 'sheetView':
            if (sheetView) { // solo la primera vista
              sheetView = false; seccion = 'sheetView';
              hoja.vista.cuadricula = a.showGridLines !== '0' && a.showGridLines !== 'false';
              hoja.vista.encabezados = a.showRowColHeaders !== '0' && a.showRowColHeaders !== 'false';
              hoja.vista.zoom = a.zoomScale ? parseFloat(a.zoomScale) : 100;
              hoja.vista.derechaAIzquierda = a.rightToLeft === '1' || a.rightToLeft === 'true';
              hoja.vista.tipo = a.view || 'normal';
              hoja.vista.mostrarFormulas = a.showFormulas === '1' || a.showFormulas === 'true';
              hoja.vista.ceros = a.showZeros !== '0' && a.showZeros !== 'false';
              hoja.vista.esquema = a.showOutlineSymbols !== '0' && a.showOutlineSymbols !== 'false';
              hoja.vista.seleccionada = a.tabSelected === '1' || a.tabSelected === 'true';
              hoja.vista.celdaSuperior = a.topLeftCell || null;
            }
            break;
          case 'pane': if (seccion === 'sheetView') hoja.vista.paneles = { x: parseFloat(a.xSplit) || 0, y: parseFloat(a.ySplit) || 0, celdaSuperior: a.topLeftCell || null, activo: a.activePane || 'topLeft', estado: a.state || 'split' }; break;
          case 'selection': if (seccion === 'sheetView' && !hoja.vista.seleccion) hoja.vista.seleccion = { celda: a.activeCell || 'A1', rangos: a.sqref || a.activeCell || 'A1' }; break;
          case 'sheetFormatPr':
            hoja.formato.anchoBase = a.baseColWidth ? parseFloat(a.baseColWidth) : 8;
            hoja.formato.anchoPorDefecto = a.defaultColWidth ? parseFloat(a.defaultColWidth) : null;
            hoja.formato.altoPorDefecto = a.defaultRowHeight ? parseFloat(a.defaultRowHeight) : null;
            hoja.formato.altoPersonalizado = a.customHeight === '1' || a.customHeight === 'true';
            hoja.formato.filasOcultasPorDefecto = a.zeroHeight === '1' || a.zeroHeight === 'true';
            hoja.formato.nivelFilaMax = parseInt(a.outlineLevelRow, 10) || 0;
            hoja.formato.nivelColMax = parseInt(a.outlineLevelCol, 10) || 0;
            break;
          case 'col': {
            const min = parseInt(a.min, 10), max = parseInt(a.max, 10);
            if (min >= 1 && min <= MAX_COLUMNAS && max >= min && hoja.columnas.length < MAX_COLUMNAS) hoja.columnas.push({ min: min - 1, max: Math.min(max, MAX_COLUMNAS) - 1, ancho: a.width !== undefined ? parseFloat(a.width) : null,
              oculta: a.hidden === '1' || a.hidden === 'true', personalizado: a.customWidth === '1' || a.customWidth === 'true', estilo: parseInt(a.style, 10) || 0,
              nivel: parseInt(a.outlineLevel, 10) || 0, colapsada: a.collapsed === '1' || a.collapsed === 'true', ajusteAuto: a.bestFit === '1' || a.bestFit === 'true' });
            break;
          }
          case 'sheetData': seccion = 'sheetData'; break;
          case 'row': {
            const rf = a.r !== undefined ? parseInt(a.r, 10) : NaN;
            fila = rf >= 1 && rf <= MAX_FILAS ? rf - 1 : fila + 1; // r ausente o absurdo: la fila siguiente
            colSiguiente = 0;
            const info = { alto: a.ht !== undefined ? parseFloat(a.ht) : null, oculta: a.hidden === '1' || a.hidden === 'true', personalizada: a.customHeight === '1' || a.customHeight === 'true',
              estilo: parseInt(a.s, 10) || 0, formatoPersonalizado: a.customFormat === '1' || a.customFormat === 'true', nivel: parseInt(a.outlineLevel, 10) || 0, colapsada: a.collapsed === '1' || a.collapsed === 'true' };
            if (fila < MAX_FILAS && (info.alto !== null || info.oculta || info.estilo || info.nivel || info.colapsada || info.personalizada)) hoja.filas.set(fila, info);
            break;
          }
          case 'c': {
            let c = colSiguiente;
            if (a.r !== undefined) { const d = decodificarDireccion(a.r); if (d && d.f >= 0 && d.f < MAX_FILAS && d.c < MAX_COLUMNAS) { c = d.c; if (d.f !== fila) fila = d.f; } }
            colSiguiente = c + 1;
            celda = { f: fila, c, t: a.t || 'n', estilo: parseInt(a.s, 10) || 0, v: null, is: '', formula: null, rico: null };
            if (vacio) cerrarCelda();
            break;
          }
          case 'v': if (celda) { enV = true; textoCelda = ''; } break;
          case 'f': if (celda) {
            enF = true; textoCelda = '';
            celda.formula = { texto: '', tipo: a.t || 'normal', ref: a.ref || null, si: a.si !== undefined ? a.si : null, ca: a.ca === '1' || a.ca === 'true', aca: a.aca === '1' || a.aca === 'true', dt2D: a.dt2D === '1', r1: a.r1 || null, r2: a.r2 || null, del1: a.del1 === '1', del2: a.del2 === '1', dtr: a.dtr === '1' };
            if (vacio) { enF = false; celda.formula.texto = null; }
          } break;
          case 'is': if (celda) enIs = true; break;
          case 'rPh': enFonetico = true; break;
          case 't': if (enIs && !enFonetico) enT = true; break;
          case 'mergeCell': { const r = a.ref ? decodificarRango(a.ref) : null; if (r && hoja.combinadas.length < MAX_ELEMENTOS) hoja.combinadas.push(r); break; }
          case 'hyperlink': {
            const r = a.ref ? decodificarRango(a.ref) : null;
            if (r && hoja.hipervinculos.length < MAX_ELEMENTOS) hoja.hipervinculos.push({ f0: r.f0, c0: r.c0, f1: r.f1, c1: r.c1, rid: a.id || null, destino: null, ubicacion: a.location || null, visible: a.display || null, ayuda: a.tooltip || null, externo: false });
            break;
          }
          case 'pageMargins': hoja.imprimir.margenes = { izq: parseFloat(a.left), der: parseFloat(a.right), arr: parseFloat(a.top), aba: parseFloat(a.bottom), encabezado: parseFloat(a.header), pie: parseFloat(a.footer) }; break;
          case 'pageSetup': hoja.imprimir.pagina = { papel: a.paperSize !== undefined ? parseInt(a.paperSize, 10) : 1, anchoPapel: a.paperWidth || null, altoPapel: a.paperHeight || null, orientacion: a.orientation || 'default', escala: a.scale !== undefined ? parseInt(a.scale, 10) : 100,
            ajustarAncho: a.fitToWidth !== undefined ? parseInt(a.fitToWidth, 10) : 1, ajustarAlto: a.fitToHeight !== undefined ? parseInt(a.fitToHeight, 10) : 1, orden: a.pageOrder || 'downThenOver', primeraPagina: a.firstPageNumber !== undefined ? parseInt(a.firstPageNumber, 10) : null,
            usarPrimeraPagina: a.useFirstPageNumber === '1' || a.useFirstPageNumber === 'true', blancoYNegro: a.blackAndWhite === '1' || a.blackAndWhite === 'true', borrador: a.draft === '1' || a.draft === 'true', comentarios: a.cellComments || 'none', errores: a.errors || 'displayed' }; break;
          case 'printOptions': hoja.imprimir.opciones = { cuadricula: (a.gridLines === '1' || a.gridLines === 'true') && (a.gridLinesSet === undefined || a.gridLinesSet === '1' || a.gridLinesSet === 'true'), encabezados: a.headings === '1' || a.headings === 'true', centrarH: a.horizontalCentered === '1' || a.horizontalCentered === 'true', centrarV: a.verticalCentered === '1' || a.verticalCentered === 'true' }; break;
          case 'rowBreaks': seccion = 'rowBreaks'; break;
          case 'colBreaks': seccion = 'colBreaks'; break;
          case 'brk': {
            const salto = { id: parseInt(a.id, 10), min: a.min !== undefined ? parseInt(a.min, 10) : 0, max: a.max !== undefined ? parseInt(a.max, 10) : 0, manual: a.man === '1' || a.man === 'true' };
            if (seccion === 'rowBreaks') hoja.imprimir.saltosFila.push(salto); else if (seccion === 'colBreaks') hoja.imprimir.saltosCol.push(salto);
            break;
          }
          case 'headerFooter': hf = { distintoPrimera: a.differentFirst === '1' || a.differentFirst === 'true', distintoParImpar: a.differentOddEven === '1' || a.differentOddEven === 'true', escalar: a.scaleWithDoc !== '0' && a.scaleWithDoc !== 'false', alinear: a.alignWithMargins !== '0' && a.alignWithMargins !== 'false', impar: {}, par: {}, primera: {} }; hoja.imprimir.encabezadoPie = hf; break;
          case 'oddHeader': case 'oddFooter': case 'evenHeader': case 'evenFooter': case 'firstHeader': case 'firstFooter': enHf = nombre; textoCelda = ''; break;
          case 'autoFilter': if (!hoja.filtroAuto) hoja.filtroAuto = { ref: a.ref || null }; break;
          case 'sheetProtection': hoja.proteccion = { hoja: a.sheet === '1' || a.sheet === 'true', objetos: a.objects === '1', escenarios: a.scenarios === '1', conClave: !!(a.password || a.hashValue || a.algorithmName) }; break;
          case 'drawing': hoja.partes.dibujoRid = a.id || null; break;
          case 'legacyDrawing': hoja.partes.dibujoVmlRid = a.id || null; break;
          case 'tablePart': if (a.id) hoja.partes.otras.push({ tipo: 'tabla', rid: a.id }); break;
          case 'conditionalFormatting': cf = { rangos: (a.sqref || '').split(/\s+/).filter(Boolean).map(decodificarRango).filter(Boolean), reglas: [] }; if (hoja.formatosCondicionales.length < MAX_ELEMENTOS) hoja.formatosCondicionales.push(cf); break;
          case 'cfRule': if (cf) { regla = { tipo: a.type || 'expression', prioridad: parseInt(a.priority, 10) || 0, dxf: a.dxfId !== undefined ? parseInt(a.dxfId, 10) : null, operador: a.operator || null, texto: a.text || null, detener: a.stopIfTrue === '1', encima: a.aboveAverage !== '0', igual: a.equalAverage === '1', rango: a.rank !== undefined ? parseInt(a.rank, 10) : null, porcentaje: a.percent === '1', abajo: a.bottom === '1', periodo: a.timePeriod || null, desvEst: a.stdDev !== undefined ? parseInt(a.stdDev, 10) : null, formulas: [], escala: null, barras: null, iconos: null }; cf.reglas.push(regla); } break;
          case 'formula': if (regla) { enFormulaCf = true; textoCelda = ''; } else if (validacion) { enFormulaVal = 'formula'; textoCelda = ''; } break;
          case 'formula1': case 'formula2': if (validacion) { enFormulaVal = nombre; textoCelda = ''; } break;
          case 'colorScale': if (regla) regla.escala = { cfvo: [], colores: [] }; break;
          case 'dataBar': if (regla) regla.barras = { cfvo: [], color: null, mostrarValor: a.showValue !== '0', min: a.minLength !== undefined ? parseInt(a.minLength, 10) : 10, max: a.maxLength !== undefined ? parseInt(a.maxLength, 10) : 90 }; break;
          case 'iconSet': if (regla) regla.iconos = { conjunto: a.iconSet || '3TrafficLights1', cfvo: [], mostrarValor: a.showValue !== '0', invertir: a.reverse === '1', soloIcono: false }; break;
          case 'cfvo': if (regla) { const v = { tipo: a.type || 'num', valor: a.val !== undefined ? a.val : null, gte: a.gte !== '0' }; (regla.escala || regla.barras || regla.iconos || { cfvo: [] }).cfvo.push(v); } break;
          case 'color': if (regla && regla.escala) regla.escala.colores.push(XlsxEstilos.colorRef(a)); else if (regla && regla.barras) regla.barras.color = XlsxEstilos.colorRef(a); break;
          case 'dataValidation': validacion = { rangos: (a.sqref || '').split(/\s+/).filter(Boolean).map(decodificarRango).filter(Boolean), tipo: a.type || 'none', operador: a.operator || 'between', permitirVacio: a.allowBlank === '1', ocultarFlecha: a.showDropDown === '1', mostrarEntrada: a.showInputMessage === '1', mostrarError: a.showErrorMessage === '1', estiloError: a.errorStyle || 'stop', tituloError: a.errorTitle || '', error: a.error || '', tituloEntrada: a.promptTitle || '', entrada: a.prompt || '', formula1: null, formula2: null }; if (hoja.validaciones.length < MAX_ELEMENTOS) hoja.validaciones.push(validacion); break;
          default: break;
        }
      },
      texto(t) {
        if (enV || enF || enFormulaCf || enFormulaVal || enHf) textoCelda += t;
        else if (enT && celda) celda.is += t;
      },
      cerrar(nombre) {
        if (nombre === 'customSheetViews') { enVistasPersonalizadas = false; return; }
        if (enVistasPersonalizadas) return;
        switch (nombre) {
          case 'v': if (enV && celda) { celda.v = textoCelda; } enV = false; break;
          case 'f': if (enF && celda && celda.formula) { celda.formula.texto = textoCelda === '' && celda.formula.tipo === 'shared' && celda.formula.ref === null ? null : XmlLigero.sinEscapes(textoCelda); } enF = false; break;
          case 'is': enIs = false; break;
          case 'rPh': enFonetico = false; break;
          case 't': enT = false; break;
          case 'c': cerrarCelda(); break;
          case 'sheetView': seccion = ''; break;
          case 'sheetData': seccion = ''; break;
          case 'rowBreaks': case 'colBreaks': seccion = ''; break;
          case 'conditionalFormatting': cf = null; regla = null; break;
          case 'cfRule': regla = null; break;
          case 'formula': if (enFormulaCf && regla) regla.formulas.push(textoCelda); else if (enFormulaVal && validacion) validacion.formula1 = textoCelda; enFormulaCf = false; enFormulaVal = null; break;
          case 'formula1': if (validacion) validacion.formula1 = textoCelda; enFormulaVal = null; break;
          case 'formula2': if (validacion) validacion.formula2 = textoCelda; enFormulaVal = null; break;
          case 'dataValidation': validacion = null; break;
          case 'oddHeader': case 'oddFooter': case 'evenHeader': case 'evenFooter': case 'firstHeader': case 'firstFooter':
            if (hf) { const grupo = nombre.startsWith('odd') ? hf.impar : nombre.startsWith('even') ? hf.par : hf.primera; grupo[nombre.endsWith('Header') ? 'encabezado' : 'pie'] = textoCelda; }
            enHf = null; break;
          default: break;
        }
      },
    });
    // Estilo (índice de xf) de una celda: el suyo; si no tiene <c>, el de su fila (si es de formato propio) y luego el de su columna.
    let estiloColumna = null;
    hoja.estiloDe = (f, c) => {
      const k = C.indiceDe(f, c);
      if (k >= 0) return C.estilo[k];
      const fila = hoja.filas.get(f);
      if (fila && fila.formatoPersonalizado) return fila.estilo;
      if (!estiloColumna) {
        estiloColumna = new Int32Array(16384);
        for (const col of hoja.columnas) for (let q = col.min; q <= col.max; q++) estiloColumna[q] = col.estilo;
      }
      return estiloColumna[c] || 0;
    };
    C._indexar(); // deja las celdas en orden (fila, columna) y sin duplicadas: los índices de paraCada no cambian después
    return hoja;
  }

  // ---- Libro ------------------------------------------------------------------------------------------

  const TIPOS_HOJA = { worksheet: 'hoja', chartsheet: 'grafico', dialogsheet: 'dialogo', macrosheet: 'macro', intlmacrosheet: 'macro' };

  async function aBytes(origen) {
    if (origen instanceof Uint8Array) return origen;
    if (origen instanceof ArrayBuffer) return new Uint8Array(origen);
    if (origen && typeof origen.arrayBuffer === 'function') return new Uint8Array(await origen.arrayBuffer());
    throw new ErrorXlsx('noZip', 'origen desconocido');
  }

  // 'D0CF11E0A1B11AE1': contenedor binario de Office (.xls antiguo, o un .xlsx cifrado con contraseña).
  function esContenedorBinario(bytes) {
    return bytes.length > 8 && bytes[0] === 0xD0 && bytes[1] === 0xCF && bytes[2] === 0x11 && bytes[3] === 0xE0 && bytes[4] === 0xA1 && bytes[5] === 0xB1 && bytes[6] === 0x1A && bytes[7] === 0xE1;
  }
  function contieneUtf16(bytes, cadena) {
    const objetivo = [];
    for (let i = 0; i < cadena.length; i++) { objetivo.push(cadena.charCodeAt(i) & 255, 0); }
    const limite = bytes.length - objetivo.length;
    outer: for (let i = 0; i <= limite; i++) {
      if (bytes[i] !== objetivo[0]) continue;
      for (let j = 1; j < objetivo.length; j++) if (bytes[i + j] !== objetivo[j]) continue outer;
      return true;
    }
    return false;
  }

  // Cualquier fallo al leer sale como ErrorXlsx: un zip dañado o enorme (ErrorZip) o una hoja que agota la memoria
  // (RangeError) no deben llegar a la interfaz como mensajes técnicos.
  async function envolver(trabajo) {
    try { return await trabajo(); } catch (error) {
      if (error instanceof ErrorXlsx) throw error;
      if (error && error.name === 'ErrorZip') throw new ErrorXlsx(error.codigo === 'grande' ? 'grande' : error.codigo === 'cifrado' ? 'cifrado' : 'danado', error.detalle || '');
      if (error instanceof RangeError) throw new ErrorXlsx('grande', String(error.message));
      throw new ErrorXlsx('danado', error && error.message ? error.message : String(error));
    }
  }
  const abrir = (origen) => envolver(() => abrirInterno(origen));

  async function abrirInterno(origen) {
    const bytes = await aBytes(origen);
    if (esContenedorBinario(bytes)) {
      throw new ErrorXlsx(contieneUtf16(bytes, 'EncryptedPackage') ? 'cifrado' : 'antiguo');
    }
    let zip;
    try { zip = await ZipPropio.loadAsync(bytes); } catch (error) {
      throw new ErrorXlsx('noZip', error && error.detalle ? error.detalle : String(error));
    }

    // Parte principal: la relación officeDocument de la raíz (o xl/workbook.xml).
    let rutaLibro = null;
    const textoRaiz = await textoDe(zip, '_rels/.rels');
    if (textoRaiz) {
      XmlLigero.recorrer(textoRaiz, { abrir(nombre, a) { if (nombre === 'Relationship' && /officeDocument$/.test(a.Type || '') && !rutaLibro) rutaLibro = resolverRuta('.rels', a.Target || ''); } });
    }
    if (!rutaLibro || !(await buscarEntrada(zip, rutaLibro))) rutaLibro = await buscarEntrada(zip, 'xl/workbook.xml');
    if (!rutaLibro) throw new ErrorXlsx(await buscarEntrada(zip, 'xl/workbook.bin') ? 'antiguo' : 'sinLibro'); // .xlsb: binario
    rutaLibro = (await buscarEntrada(zip, rutaLibro)) || rutaLibro;

    const textoLibro = await textoDe(zip, rutaLibro);
    if (!textoLibro) throw new ErrorXlsx('sinLibro');
    const rels = await leerRelaciones(zip, rutaLibro);
    const relPorId = new Map(rels.map((r) => [r.id, r]));

    const libro = {
      zip, rutaLibro, relaciones: rels, hojas: [], nombres: [], fecha1904: false, estilos: null, textos: [], vistaLibro: { hojaActiva: 0, primeraHoja: 0 },
      calculo: null, propiedades: {}, _hojasCargadas: new Map(),
    };
    // workbook.xml
    let enNombre = null, textoNombre = '';
    XmlLigero.recorrer(textoLibro, {
      abrir(nombre, a) {
        switch (nombre) {
          case 'workbookPr': libro.fecha1904 = a.date1904 === '1' || a.date1904 === 'true'; libro.propiedades.codeName = a.codeName || null; break;
          case 'workbookView': libro.vistaLibro = { hojaActiva: parseInt(a.activeTab, 10) || 0, primeraHoja: parseInt(a.firstSheet, 10) || 0 }; break;
          case 'sheet': {
            const rel = relPorId.get(a.id);
            libro.hojas.push({ nombre: a.name || ('Hoja' + (libro.hojas.length + 1)), id: parseInt(a.sheetId, 10), estado: a.state === 'hidden' ? 'oculta' : a.state === 'veryHidden' ? 'muyOculta' : 'visible',
              tipo: rel ? (Object.prototype.hasOwnProperty.call(TIPOS_HOJA, rel.tipo) ? TIPOS_HOJA[rel.tipo] : 'otra') : 'hoja', ruta: rel && !rel.externa ? rel.destino : null, rid: a.id || null });
            break;
          }
          case 'definedName': enNombre = { nombre: a.name, hojaLocal: a.localSheetId !== undefined ? parseInt(a.localSheetId, 10) : null, oculto: a.hidden === '1' || a.hidden === 'true', comentario: a.comment || '', ref: '' }; textoNombre = ''; break;
          case 'calcPr': libro.calculo = { modo: a.calcMode || 'auto', completoAlCargar: a.fullCalcOnLoad === '1' || a.fullCalcOnLoad === 'true', id: a.calcId || null, precisionCompleta: a.fullPrecision !== '0' }; break;
          default: break;
        }
      },
      texto(t) { if (enNombre) textoNombre += t; },
      cerrar(nombre) { if (nombre === 'definedName' && enNombre) { enNombre.ref = textoNombre.trim(); libro.nombres.push(enNombre); enNombre = null; } },
    });
    if (!libro.hojas.length) throw new ErrorXlsx('sinLibro', 'sin hojas');

    // Textos compartidos, estilos y tema (los que existan).
    const relDe = (tipo) => rels.find((r) => r.tipo === tipo && !r.externa);
    const relTextos = relDe('sharedStrings'), relEstilos = relDe('styles'), relTema = relDe('theme');
    libro.textos = leerTextosCompartidos(relTextos ? await textoDe(zip, relTextos.destino) : null);
    libro.estilos = XlsxEstilos.leer(relEstilos ? await textoDe(zip, relEstilos.destino) : null, relTema ? await textoDe(zip, relTema.destino) : null);
    libro.tieneMacros = !!rels.find((r) => r.tipo === 'vbaProject');

    libro.cargarHoja = (indice) => envolver(() => cargarHoja(libro, indice));
    return libro;
  }

  async function cargarHoja(libro, indice) {
    if (libro._hojasCargadas.has(indice)) return libro._hojasCargadas.get(indice);
    const meta = libro.hojas[indice];
    if (!meta) throw new RangeError('hoja ' + indice);
    let hoja;
    if (meta.tipo !== 'hoja' || !meta.ruta) {
      hoja = Object.assign(leerHoja('<worksheet/>', { textos: [], fecha1904: libro.fecha1904 }), { sinCeldas: true });
    } else {
      const texto = await textoDe(libro.zip, meta.ruta);
      if (texto === null) throw new ErrorXlsx('danado', 'falta ' + meta.ruta);
      hoja = leerHoja(texto, { textos: libro.textos, fecha1904: libro.fecha1904 });
      hoja.relaciones = await leerRelaciones(libro.zip, meta.ruta);
      const porId = new Map(hoja.relaciones.map((r) => [r.id, r]));
      for (const h of hoja.hipervinculos) {
        const rel = h.rid ? porId.get(h.rid) : null;
        if (rel) { h.destino = rel.destino; h.externo = rel.externa; }
        // Solo estos esquemas se pueden usar como enlace de página (un libro hostil puede traer javascript:, data:…).
        h.destinoSeguro = h.externo && /^(https?:|mailto:|tel:|ftp:)/i.test(h.destino || '') ? h.destino : null;
      }
      const rd = hoja.partes.dibujoRid ? porId.get(hoja.partes.dibujoRid) : null;
      if (rd) hoja.partes.dibujo = rd.destino;
      const rv = hoja.partes.dibujoVmlRid ? porId.get(hoja.partes.dibujoVmlRid) : null;
      if (rv) hoja.partes.dibujoVml = rv.destino;
      const rc = hoja.relaciones.find((r) => r.tipo === 'comments');
      if (rc) hoja.partes.comentarios = rc.destino;
      hoja.partes.tablas = hoja.relaciones.filter((r) => r.tipo === 'table').map((r) => r.destino);
    }
    hoja.nombre = meta.nombre;
    hoja.indice = indice;
    libro._hojasCargadas.set(indice, hoja);
    return hoja;
  }

  const hojaVacia = () => leerHoja('<worksheet/>', { textos: [], fecha1904: false });

  // El texto que Excel muestra en la celda i de la hoja: el valor con el formato de número de su estilo
  // (requiere xlsx-formato.js). Para un CSV, el número sale con el texto original ("1.234,50").
  // opciones: las de XlsxFormato.formatear (idioma, decimal, miles, ancho…). Devuelve { texto, color, desborde }.
  function textoCelda(libro, hoja, i, opciones) {
    const C = hoja.celdas;
    const tipo = C.tipo[i], valor = C.valor[i];
    const opc = Object.assign({}, opciones || {}, { fecha1904: libro.fecha1904 });
    if (tipo === TIPO.vacio) return { texto: '', color: null };
    if (C.textoOriginal.has(i)) return { texto: C.textoOriginal.get(i), color: null };
    if (tipo === TIPO.error) {
      const nombres = XlsxFormato.NOMBRES[opc.idioma] || XlsxFormato.NOMBRES.en;
      return { texto: nombres.errores[valor] || valor, color: null };
    }
    const xf = libro.estilos.xf(C.estilo[i]);
    opc.idFormato = xf.numFmtId;
    if (opc.formatoCf) { opc.idFormato = -1; return XlsxFormato.formatear(valor, opc.formatoCf, opc); } // el formato de número de una regla condicional
    return XlsxFormato.formatear(valor, xf.formato, opc);
  }

  globalThis.LibroXlsx = { abrir, hojaVacia, textoCelda, ErrorXlsx, TIPO, colANumero, numeroACol, decodificarDireccion, codificarDireccion, decodificarRango, serialDeIso, desplazarFormula };
})();

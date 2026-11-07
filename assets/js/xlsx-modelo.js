/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
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
    // un rango con una punta fuera de la hoja es todo #REF!, como en Excel
    if (salida.includes('#REF!') && !texto.includes('#REF!')) salida = salida.replace(/(?:\$?[A-Za-z]{1,3}\$?\d+|#REF!):#REF!|#REF!:\$?[A-Za-z]{1,3}\$?\d+/g, '#REF!');
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
      celdasVm: new Map(),         // celdas con valor enriquecido (imagen en la celda): fila × 16384 + columna → índice vm de metadata.xml
      imagenesCelda: new Map(),    // las mismas celdas → ruta de la imagen en el paquete (se resuelve al cargar la hoja)
      hilos: [],                   // comentarios en hilo: { ref, f, c, entradas: [{ autor, fecha, texto, respuesta }] }
      comentarios: [],             // notas de celda: { f, c, autor, texto, runs, visible, ancla, ancho, alto, relleno } (comments + VML)
      minigraficos: [],            // grupos de minigráficos x14: { tipo, colores…, items: [{ f, sqref }] }
      formatosCondicionales: [],   // { rangos: [rango], reglas: [{ tipo, prioridad, dxf, … }] } — lectura completa en la etapa 6
      validaciones: [],            // { rangos, tipo, … }
      anclas: [],                  // dónde están los objetos del dibujo de la hoja (ver leerAnclas); su contenido lo lee xlsx-objetos.js
      tablas: [],                  // { nombre, ref: { f0, c0, f1, c1 }, encabezados (0|1), totales (0|1), estilo, primeraCol, ultimaCol, bandasFilas, bandasCols } (se llenan al cargar la hoja)
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
    let enFiltro = null, colFiltro = null;   // autoFilter y filterColumn abiertos
    let cf = null, regla = null, enFormulaCf = false, validacion = null, enFormulaVal = null, hf = null, enHf = null;
    // Extensiones x14 de Excel 2010 (extLst): barras de datos con eje y negativos, conjuntos de iconos propios y minigráficos.
    // Las reglas con su par clásico se unen por el id (x14:id); las que solo existen en x14 se agregan como reglas nuevas.
    const URI_CF14 = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}', URI_SPARK = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}';
    let extHondo = 0;   // <ext> anidados dentro de un bloque x14 abierto
    let ext14 = null, x14cf = null, x14r = null, x14v = null, x14dxf = null, enX14 = null, enId14 = false, spark = null, sparkItem = null;
    const ids14 = new Map();
    const rangosX14 = (t) => t.split(/\s+/).filter(Boolean).map(decodificarRango).filter(Boolean);
    const escX = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    const x14Abrir = (nombre, a) => {
      if (x14dxf) { x14dxf.push('<' + nombre + Object.keys(a).map((k) => ` ${k}="${escX(a[k])}"`).join('') + '>'); return; } // el dxf en línea se vuelve a escribir tal cual para el lector de estilos
      switch (nombre) {
        case 'dxf': if (x14r && x14dxf === null) x14dxf = []; break;
        case 'conditionalFormatting': x14cf = { rangos: [], reglas: [] }; break;
        case 'cfRule': if (x14cf) { x14r = { id: a.id || null, tipo: a.type || '', prioridad: parseInt(a.priority, 10) || 0, barras: null, iconos: null, attr: a, formulas: [], dxfXml: null }; x14cf.reglas.push(x14r); } break;
        case 'dataBar': if (x14r) x14r.barras = {
          min: a.minLength !== undefined ? parseFloat(a.minLength) : 0, max: a.maxLength !== undefined ? parseFloat(a.maxLength) : 100,
          gradiente: a.gradient !== '0' && a.gradient !== 'false', borde: a.border === '1' || a.border === 'true',
          negIgual: a.negativeBarColorSameAsPositive === '1' || a.negativeBarColorSameAsPositive === 'true', // sin el atributo: color propio de los negativos
          negBordeIgual: a.negativeBarBorderColorSameAsPositive !== '0' && a.negativeBarBorderColorSameAsPositive !== 'false',
          eje: a.axisPosition || 'automatic', direccion: a.direction || 'context', mostrarValor: a.showValue !== '0' && a.showValue !== 'false',
          cfvo: [], colorBorde: null, colorNeg: null, colorBordeNeg: null, colorEje: null }; break;
        case 'iconSet': if (x14r) x14r.iconos = { conjunto: a.iconSet || '3TrafficLights1', cfvo: [], mostrarValor: a.showValue !== '0' && a.showValue !== 'false', invertir: a.reverse === '1' || a.reverse === 'true', soloIcono: false, propios: [], personalizado: a.custom === '1' || a.custom === 'true' }; break;
        case 'cfIcon': if (x14r && x14r.iconos) x14r.iconos.propios.push({ conjunto: a.iconSet || null, id: parseInt(a.iconId, 10) }); break;
        case 'cfvo': { const d = x14r && (x14r.barras || x14r.iconos); if (d) { x14v = { tipo: a.type || 'num', valor: a.val !== undefined ? a.val : null, gte: a.gte !== '0' }; d.cfvo.push(x14v); } break; }
        case 'borderColor': if (x14r && x14r.barras) x14r.barras.colorBorde = XlsxEstilos.colorRef(a); break;
        case 'negativeFillColor': if (x14r && x14r.barras) x14r.barras.colorNeg = XlsxEstilos.colorRef(a); break;
        case 'negativeBorderColor': if (x14r && x14r.barras) x14r.barras.colorBordeNeg = XlsxEstilos.colorRef(a); break;
        case 'axisColor': if (x14r && x14r.barras) x14r.barras.colorEje = XlsxEstilos.colorRef(a); break;
        case 'f': if (x14v || sparkItem || x14r) { enX14 = 'f'; textoCelda = ''; } break;
        case 'sqref': enX14 = 'sqref'; textoCelda = ''; break;
        case 'sparklineGroup':
          spark = {
            tipo: a.type || 'line', grosor: a.lineWeight !== undefined ? parseFloat(a.lineWeight) : 0.75, vacias: a.displayEmptyCellsAs || 'zero',
            marcadores: a.markers === '1', alto: a.high === '1', bajo: a.low === '1', primero: a.first === '1', ultimo: a.last === '1', negativos: a.negative === '1',
            ejeX: a.displayXAxis === '1', ocultas: a.displayHidden === '1', derechaAIzquierda: a.rightToLeft === '1', ejeFechas: a.dateAxis === '1',
            minTipo: a.minAxisType || 'individual', maxTipo: a.maxAxisType || 'individual', minManual: a.manualMin !== undefined ? parseFloat(a.manualMin) : null, maxManual: a.manualMax !== undefined ? parseFloat(a.manualMax) : null,
            colores: {}, items: [] };
          if (hoja.minigraficos.length < MAX_ELEMENTOS) hoja.minigraficos.push(spark);
          break;
        case 'colorSeries': case 'colorNegative': case 'colorAxis': case 'colorMarkers': case 'colorFirst': case 'colorLast': case 'colorHigh': case 'colorLow':
          if (spark) spark.colores[nombre.slice(5).toLowerCase()] = XlsxEstilos.colorRef(a); break;
        case 'sparkline': if (spark) { sparkItem = { f: '', sqref: '' }; spark.items.push(sparkItem); } break;
        default: break;
      }
    };
    const x14Cerrar = (nombre) => {
      if (x14dxf) {
        if (nombre === 'dxf') { if (x14r) x14r.dxfXml = '<dxf>' + x14dxf.join('') + '</dxf>'; x14dxf = null; } else x14dxf.push('</' + nombre + '>');
        return;
      }
      switch (nombre) {
        case 'f': if (enX14 === 'f') { const t = textoCelda.trim(); if (sparkItem) sparkItem.f = t; else if (x14v) x14v.valor = t; else if (x14r) x14r.formulas.push(t); } enX14 = null; break;
        case 'cfvo': x14v = null; break;
        case 'sqref': if (enX14 === 'sqref') { const t = textoCelda.trim(); if (sparkItem) sparkItem.sqref = t; else if (x14cf) x14cf.rangos = rangosX14(t); } enX14 = null; break;
        case 'cfRule': x14r = null; break;
        case 'sparkline': sparkItem = null; break;
        case 'sparklineGroup': spark = null; break;
        case 'conditionalFormatting':
          if (x14cf) {
            for (const r of x14cf.reglas) {
              const par = r.id ? ids14.get(r.id) : null;
              if (par && par.barras && r.barras) par.barras.x14 = r.barras;
              else if (!par && !r.barras && !r.iconos && r.tipo && hoja.formatosCondicionales.length < MAX_ELEMENTOS) { // regla de fórmula o de valor (apunta a otra hoja) con su formato en línea
                const q = r.attr;
                hoja.formatosCondicionales.push({ rangos: x14cf.rangos, reglas: [{ tipo: r.tipo, prioridad: r.prioridad, dxf: null, dxfXml: r.dxfXml, operador: q.operator || null, texto: q.text || null, detener: q.stopIfTrue === '1', encima: q.aboveAverage !== '0', igual: q.equalAverage === '1',
                  rango: q.rank !== undefined ? parseInt(q.rank, 10) : null, porcentaje: q.percent === '1', abajo: q.bottom === '1', periodo: q.timePeriod || null, desvEst: q.stdDev !== undefined ? parseInt(q.stdDev, 10) : null, formulas: r.formulas, escala: null, barras: null, iconos: null }] });
              } else if (!par && (r.barras || r.iconos) && hoja.formatosCondicionales.length < MAX_ELEMENTOS) {
                hoja.formatosCondicionales.push({ rangos: x14cf.rangos, reglas: [{ tipo: r.iconos ? 'iconSet' : 'dataBar', prioridad: r.prioridad, dxf: null, operador: null, texto: null, detener: false, encima: true, igual: false, rango: null, porcentaje: false, abajo: false, periodo: null, desvEst: null, formulas: [], escala: null,
                  barras: r.barras ? { cfvo: r.barras.cfvo, color: null, mostrarValor: r.barras.mostrarValor, min: r.barras.min, max: r.barras.max, x14: r.barras } : null, iconos: r.iconos }] });
              }
            }
          }
          x14cf = null; break;
        default: break;
      }
    };
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
      if (celda.vm > 0 && hoja.celdasVm.size < MAX_ELEMENTOS) hoja.celdasVm.set(celda.f * 16384 + celda.c, celda.vm);
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
        if (nombre === 'ext') { if (ext14) extHondo++; else ext14 = a.uri === URI_CF14 ? 'cf' : a.uri === URI_SPARK ? 'spark' : null; } // un <ext> dentro del bloque x14 no lo cierra
        if (ext14) { x14Abrir(nombre, a); return; }
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
            celda = { f: fila, c, t: a.t || 'n', estilo: parseInt(a.s, 10) || 0, v: null, is: '', formula: null, rico: null, vm: a.vm !== undefined ? parseInt(a.vm, 10) || 0 : 0 };
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
          case 'autoFilter': if (!hoja.filtroAuto) hoja.filtroAuto = { ref: a.ref || null, columnas: new Map() }; enFiltro = hoja.filtroAuto; break;
          case 'filterColumn': if (enFiltro && enFiltro.columnas.size < 16384) { colFiltro = { col: parseInt(a.colId, 10) || 0, oculto: a.hiddenButton === '1' || a.hiddenButton === 'true', filtrada: false }; enFiltro.columnas.set(colFiltro.col, colFiltro); } break;
          case 'filters': case 'customFilters': case 'top10': case 'colorFilter': case 'dynamicFilter': case 'iconFilter': if (colFiltro) colFiltro.filtrada = true; break;
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
          case 'id': if (regla) { enId14 = true; textoCelda = ''; } break;
          case 'cfvo': if (regla) { const v = { tipo: a.type || 'num', valor: a.val !== undefined ? a.val : null, gte: a.gte !== '0' }; (regla.escala || regla.barras || regla.iconos || { cfvo: [] }).cfvo.push(v); } break;
          case 'color': if (regla && regla.escala) regla.escala.colores.push(XlsxEstilos.colorRef(a)); else if (regla && regla.barras) regla.barras.color = XlsxEstilos.colorRef(a); break;
          case 'dataValidation': validacion = { rangos: (a.sqref || '').split(/\s+/).filter(Boolean).map(decodificarRango).filter(Boolean), tipo: a.type || 'none', operador: a.operator || 'between', permitirVacio: a.allowBlank === '1', ocultarFlecha: a.showDropDown === '1', mostrarEntrada: a.showInputMessage === '1', mostrarError: a.showErrorMessage === '1', estiloError: a.errorStyle || 'stop', tituloError: a.errorTitle || '', error: a.error || '', tituloEntrada: a.promptTitle || '', entrada: a.prompt || '', formula1: null, formula2: null }; if (hoja.validaciones.length < MAX_ELEMENTOS) hoja.validaciones.push(validacion); break;
          default: break;
        }
      },
      texto(t) {
        if (x14dxf) x14dxf.push(escX(t));
        else if (enV || enF || enFormulaCf || enFormulaVal || enHf || enId14 || enX14) textoCelda += t;
        else if (enT && celda) celda.is += t;
      },
      cerrar(nombre) {
        if (nombre === 'customSheetViews') { enVistasPersonalizadas = false; return; }
        if (enVistasPersonalizadas) return;
        if (ext14) { if (nombre === 'ext' && extHondo === 0) ext14 = null; else { if (nombre === 'ext') extHondo--; x14Cerrar(nombre); } return; }
        switch (nombre) {
          case 'id': if (enId14 && regla) ids14.set(textoCelda.trim(), regla); enId14 = false; break;
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
          case 'autoFilter': enFiltro = null; colFiltro = null; break;
          case 'filterColumn': colFiltro = null; break;
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
          case 'workbookPr':
            // solo el primero: Excel 2013+ escribe otro <x15:workbookPr> dentro de extLst (sin date1904) que no debe pisar la base de fechas
            if (libro._workbookPr) break;
            libro._workbookPr = true;
            libro.fecha1904 = a.date1904 === '1' || a.date1904 === 'true'; libro.propiedades.codeName = a.codeName || null; break;
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
    libro.rutaTema = relTema ? relTema.destino : null;
    libro.relaciones = rels;
    const relPersonas = relDe('person');
    libro.rutaPersonas = relPersonas ? relPersonas.destino : null;   // quiénes escriben los comentarios en hilo
    libro.tieneMacros = !!rels.find((r) => r.tipo === 'vbaProject');

    libro.cargarHoja = (indice) => envolver(() => cargarHoja(libro, indice));
    return libro;
  }

  const MAX_TABLAS = 5000;
  const MAX_ANCLAS = 5000;
  // Las anclas del dibujo de una hoja (xl/drawings/drawingN.xml), sin DOM: dónde está cada objeto. Las posiciones en EMU.
  // { tipo: 'dos' | 'uno' | 'abs', de: { col, fila, dx, dy }, a: {…}, ext: { w, h }, pos: { x, y }, imprimible, oculto }, en el orden del archivo.
  // Notas de celda (xl/commentsN.xml) con su cuadro (xl/drawings/vmlDrawingN.vml: dónde va, si se muestra siempre y de qué tamaño).
  // El VML no es XML estricto pero el lector ligero lo traga; solo se miran las formas con ObjectType="Note".
  function leerComentarios(xml, vml) {
    const lista = [];
    const autores = [];
    let enAutor = false, c = null, enR = false, enText = false, enT = false, rpr = null, run = null;
    if (xml) XmlLigero.recorrer(xml, {
      abrir(nombre, a) {
        if (c && enText) {
          if (nombre === 'r') { enR = true; run = { texto: '', negrita: false, cursiva: false, subrayado: false, tachado: false, tam: null, fuente: null, color: null }; }
          else if (nombre === 'rPr' && enR) rpr = run;
          else if (rpr) {
            if (nombre === 'b') rpr.negrita = a.val !== '0' && a.val !== 'false';
            else if (nombre === 'i') rpr.cursiva = a.val !== '0' && a.val !== 'false';
            else if (nombre === 'u') rpr.subrayado = a.val !== 'none';
            else if (nombre === 'strike') rpr.tachado = a.val !== '0' && a.val !== 'false';
            else if (nombre === 'sz') rpr.tam = parseFloat(a.val) || null;
            else if (nombre === 'rFont') rpr.fuente = a.val || null;
            else if (nombre === 'color') rpr.color = XlsxEstilos.colorRef(a);
          }
          if (nombre === 't') enT = true;
          return;
        }
        if (nombre === 'author') { enAutor = true; autores.push(''); }
        else if (nombre === 'comment' && lista.length < MAX_ELEMENTOS) {
          const d = a.ref ? decodificarDireccion(a.ref) : null;
          c = d ? { f: d.f, c: d.c, autor: autores[parseInt(a.authorId, 10)] || '', texto: '', runs: [], visible: false, ancla: null, ancho: null, alto: null, relleno: null, borde: null } : null;
          if (c) lista.push(c);
        } else if (nombre === 'text' && c) { enText = true; }
      },
      texto(t) { if (enAutor) autores[autores.length - 1] += t; else if (enT && run) run.texto += t; else if (enT && c && !enR) c.runs.push({ texto: t, negrita: false, cursiva: false, subrayado: false, tachado: false, tam: null, fuente: null, color: null }); },
      cerrar(nombre) {
        if (nombre === 'author') { enAutor = false; }
        else if (nombre === 'rPr') rpr = null;
        else if (nombre === 't') enT = false;
        else if (nombre === 'r' && enR) { if (c) c.runs.push(run); run = null; enR = false; rpr = null; }
        else if (nombre === 'text') { enText = false; if (c) c.texto = c.runs.map((r) => r.texto).join(''); }
        else if (nombre === 'comment') c = null;
      },
    });
    // el elemento author llega vacío si no hay texto: se rellena al abrir
    // (autores se arma con un push por <author>)
    if (vml && lista.length) {
      const porCelda = new Map(lista.map((x) => [x.f * 16384 + x.c, x]));
      let forma = null, enTxt = null;
      XmlLigero.recorrer(vml, {
        abrir(nombre, a) {
          if (nombre === 'shape') { forma = { estilo: a.style || '', relleno: a.fillcolor || null, borde: a.strokecolor || null, esNota: false, visible: false, f: -1, c: -1, anclaje: null }; }
          else if (!forma) return;
          else if (nombre === 'ClientData') forma.esNota = a.ObjectType === 'Note';
          else if (nombre === 'Visible') forma.visible = true;
          else if (nombre === 'Anchor' || nombre === 'Row' || nombre === 'Column') enTxt = nombre;
        },
        texto(t) {
          if (!forma || !enTxt) return;
          if (enTxt === 'Anchor') forma.anclaje = (forma.anclaje || '') + t; else if (enTxt === 'Row') forma.f = parseInt(t, 10); else forma.c = parseInt(t, 10);
        },
        cerrar(nombre) {
          if (nombre === 'Anchor' || nombre === 'Row' || nombre === 'Column') enTxt = null;
          if (nombre !== 'shape' || !forma) return;
          const x = forma.esNota ? porCelda.get(forma.f * 16384 + forma.c) : null;
          if (x) {
            x.visible = forma.visible || /visibility:\s*visible/i.test(forma.estilo);
            const n = (forma.anclaje || '').split(',').map((v) => parseInt(v, 10));
            if (n.length >= 8 && n.every(Number.isFinite)) {
              const PX = 9525; // 1 px de pantalla en EMU
              x.ancla = { tipo: 'dos', de: { col: n[0], dx: n[1] * PX, fila: n[2], dy: n[3] * PX }, a: { col: n[4], dx: n[5] * PX, fila: n[6], dy: n[7] * PX }, ext: null, pos: null, imprimible: true, oculto: false };
            }
            const w = /width:\s*([\d.]+)pt/i.exec(forma.estilo), h = /height:\s*([\d.]+)pt/i.exec(forma.estilo);
            if (w) x.ancho = parseFloat(w[1]);
            if (h) x.alto = parseFloat(h[1]);
            const col = (v) => { const m = /^#([0-9a-f]{6})$/i.exec(v || ''); return m ? '#' + m[1].toUpperCase() : null; };
            x.relleno = col(forma.relleno); x.borde = col(forma.borde);
          }
          forma = null;
        },
      });
    }
    return lista;
  }

  // Imágenes dentro de las celdas (Excel 365, "colocar en la celda"): la celda trae vm="k" (1 en adelante) → xl/metadata.xml (valueMetadata: el k-ésimo
  // bk → rc v = índice en futureMetadata; su xlrd:rvb i = índice del valor en rdrichvalue.xml) → ese rv, cuya estructura (rdrichvaluestructure.xml)
  // dice qué <v> es el _rvRel:LocalImageIdentifier → richValueRel.xml (la rel r:id número ese) → su .rels → la imagen. Devuelve k → ruta.
  async function leerImagenesEnCeldas(libro) {
    const salida = new Map();
    const rels = libro.relaciones || [];
    const destino = (tipo) => { const r = rels.find((x) => x.tipo === tipo && !x.externa); return r ? r.destino : null; };
    const rutaMeta = destino('sheetMetadata'), rutaValores = destino('rdRichValue'), rutaEstructura = destino('rdRichValueStructure'), rutaRel = destino('richValueRel');
    if (!rutaMeta || !rutaValores || !rutaEstructura || !rutaRel) return salida;
    const [meta, valores, estructura, rel] = await Promise.all([textoDe(libro.zip, rutaMeta), textoDe(libro.zip, rutaValores), textoDe(libro.zip, rutaEstructura), textoDe(libro.zip, rutaRel)]);
    if (meta === null || valores === null || estructura === null || rel === null) return salida;
    // 1. metadata: vm → índice del bloque de futuro, y de este → índice del valor
    const valueMeta = [], futuro = [];
    let enValueMeta = false, enFuturo = false, tipoRico = 0, nTipos = 0, enBk = false;
    XmlLigero.recorrer(meta, {
      abrir(nombre, a) {
        if (nombre === 'metadataType') { nTipos++; if (a.name === 'XLRICHVALUE') tipoRico = nTipos; }
        else if (nombre === 'valueMetadata') enValueMeta = true;
        else if (nombre === 'futureMetadata' && a.name === 'XLRICHVALUE') enFuturo = true;
        else if (enValueMeta && nombre === 'bk') { enBk = true; valueMeta.push(NaN); } // un bk por vm, aunque traiga más de un rc: vale el del tipo XLRICHVALUE
        else if (enValueMeta && enBk && nombre === 'rc' && (!tipoRico || parseInt(a.t, 10) === tipoRico) && Number.isNaN(valueMeta[valueMeta.length - 1])) valueMeta[valueMeta.length - 1] = parseInt(a.v, 10);
        else if (enFuturo && nombre === 'rvb') futuro.push(parseInt(a.i, 10));
      },
      cerrar(nombre) { if (nombre === 'valueMetadata') enValueMeta = false; else if (nombre === 'futureMetadata') enFuturo = false; else if (nombre === 'bk') enBk = false; },
    });
    // 2. estructuras: para cada una, la posición de la clave _rvRel:LocalImageIdentifier
    const estructuras = [];
    let est = null;
    XmlLigero.recorrer(estructura, {
      abrir(nombre, a) {
        if (nombre === 's') { est = { tipo: a.t || '', claves: [] }; estructuras.push(est); }
        else if (nombre === 'k' && est) est.claves.push(a.n || '');
      },
      cerrar(nombre) { if (nombre === 's') est = null; },
    });
    // 3. valores
    const rvs = [];
    let rv = null, enV = false;
    XmlLigero.recorrer(valores, {
      abrir(nombre, a) {
        if (nombre === 'rv' && rvs.length < MAX_ELEMENTOS) { rv = { s: parseInt(a.s, 10), v: [] }; rvs.push(rv); }
        else if (nombre === 'v' && rv) { enV = true; rv.v.push(''); }
      },
      texto(t) { if (enV && rv) rv.v[rv.v.length - 1] += t; },
      cerrar(nombre) { if (nombre === 'v') enV = false; else if (nombre === 'rv') rv = null; },
    });
    // 4. richValueRel: orden de las rel → su r:id → destino de la imagen
    const ids = [];
    XmlLigero.recorrer(rel, { abrir(nombre, a) { if (nombre === 'rel') ids.push(a.id || a['r:id'] || ''); } });
    const relsRel = await leerRelaciones(libro.zip, rutaRel);
    const porId = new Map(relsRel.map((r) => [r.id, r]));
    valueMeta.forEach((rc, k) => {
      const iv = futuro[rc];
      const valor = rvs[iv];
      const e = valor && estructuras[valor.s];
      if (!e || e.tipo !== '_localImage') return;
      const pos = e.claves.indexOf('_rvRel:LocalImageIdentifier');
      const idx = pos >= 0 ? parseInt(valor.v[pos], 10) : NaN;
      const r = Number.isFinite(idx) ? porId.get(ids[idx]) : null;
      if (r && !r.externa && r.tipo === 'image') salida.set(k + 1, r.destino);
    });
    return salida;
  }

  // Comentarios en hilo (Excel 365): xl/threadedComments/threadedCommentN.xml, con los nombres de xl/persons/person.xml.
  // Devuelve [{ ref, f, c, entradas: [{ autor, fecha (ISO UTC), texto, respuesta }] }] en el orden del archivo (cada respuesta tras su comentario).
  function leerHilos(xml, personas) {
    const por = new Map();
    let t = null, enText = false;
    XmlLigero.recorrer(xml, {
      abrir(nombre, a) {
        if (nombre === 'threadedComment' && por.size < MAX_ELEMENTOS && a.ref) {
          const d = decodificarDireccion(a.ref);
          if (!d) { t = null; return; }
          if (!por.has(a.ref)) por.set(a.ref, { ref: a.ref, f: d.f, c: d.c, entradas: [] });
          t = { autor: personas.get(a.personId) || '', fecha: a.dT || null, texto: '', respuesta: !!a.parentId };
          por.get(a.ref).entradas.push(t);
        } else if (nombre === 'text' && t) enText = true;
      },
      texto(x) { if (enText && t) t.texto += x; },
      cerrar(nombre) { if (nombre === 'text') enText = false; else if (nombre === 'threadedComment') t = null; },
    });
    return [...por.values()];
  }
  function leerPersonas(xml) {
    const m = new Map();
    if (xml) XmlLigero.recorrer(xml, { abrir(nombre, a) { if (nombre === 'person' && a.id) m.set(a.id, a.displayName || ''); } });
    return m;
  }

  function leerAnclas(texto) {
    const anclas = [];
    let profundidad = 0, a = null, nivelAncla = -1, modo = null, campo = '', primeraCNvPr = true, enChoice = false;
    XmlLigero.recorrer(texto, {
      abrir(nombre, at) {
        profundidad++;
        if (!a) {
          // solo las hijas directas de xdr:wsDr, como las que lee xlsx-objetos.js (PptxModelo.dibujoExcel): las dos listas van por índice. Las que Excel
          // envuelve en mc:AlternateContent (segmentaciones, controles de formulario) no cuentan en ninguna; si contaran aquí, todos los objetos
          // siguientes quedarían con el ancla de otro. Las de la rama principal (mc:Choice) de esos envoltorios sí cuentan, igual que allá.
          if (nombre === 'Choice' && profundidad === 3) enChoice = true;
          if ((profundidad === 2 || (profundidad === 4 && enChoice)) && (nombre === 'twoCellAnchor' || nombre === 'oneCellAnchor' || nombre === 'absoluteAnchor') && anclas.length < MAX_ANCLAS) {
            a = { tipo: nombre === 'twoCellAnchor' ? 'dos' : nombre === 'oneCellAnchor' ? 'uno' : 'abs', de: null, a: null, ext: null, pos: null, imprimible: true, oculto: false };
            anclas.push(a); nivelAncla = profundidad; primeraCNvPr = true;
          }
          return;
        }
        const hijoDirecto = profundidad === nivelAncla + 1;
        if (hijoDirecto && (nombre === 'from' || nombre === 'to')) { modo = nombre === 'from' ? 'de' : 'a'; a[modo] = { col: 0, fila: 0, dx: 0, dy: 0 }; }
        else if (hijoDirecto && nombre === 'pos') a.pos = { x: parseInt(at.x, 10) || 0, y: parseInt(at.y, 10) || 0 };
        else if (hijoDirecto && nombre === 'ext') a.ext = { w: parseInt(at.cx, 10) || 0, h: parseInt(at.cy, 10) || 0 };
        else if (hijoDirecto && nombre === 'clientData') a.imprimible = !(at.fPrintsWithSheet === '0' || at.fPrintsWithSheet === 'false');
        else if (modo && profundidad === nivelAncla + 2 && (nombre === 'col' || nombre === 'colOff' || nombre === 'row' || nombre === 'rowOff')) campo = nombre;
        else if (nombre === 'cNvPr' && primeraCNvPr) { primeraCNvPr = false; a.oculto = at.hidden === '1' || at.hidden === 'true'; }
      },
      texto(t) {
        if (a && modo && campo) { const v = parseInt(t, 10) || 0; const d = a[modo]; if (campo === 'col') d.col = v; else if (campo === 'colOff') d.dx = v; else if (campo === 'row') d.fila = v; else d.dy = v; }
      },
      cerrar(nombre) {
        if (nombre === 'Choice' && profundidad === 3) enChoice = false;
        if (a && profundidad === nivelAncla) { a = null; modo = null; }
        else if (nombre === 'from' || nombre === 'to') modo = null;
        campo = '';
        profundidad--;
      },
    });
    return anclas;
  }
  // Una tabla (xl/tables/tableN.xml): su rango, si lleva encabezado y totales, y las opciones de estilo.
  function leerTabla(texto) {
    let tabla = null, enFiltroT = null, colFiltroT = null;
    XmlLigero.recorrer(texto, {
      abrir(nombre, a) {
        if (nombre === 'table' && !tabla) {
          const ref = decodificarRango(a.ref || '');
          if (!ref) return;
          const encabezados = a.headerRowCount === undefined ? 1 : (parseInt(a.headerRowCount, 10) > 0 ? 1 : 0);
          const totales = parseInt(a.totalsRowCount, 10) > 0 ? 1 : 0;
          tabla = { nombre: a.name || a.displayName || '', ref, encabezados, totales, estilo: null, primeraCol: false, ultimaCol: false, bandasFilas: false, bandasCols: false, filtro: null };
        } else if (nombre === 'autoFilter' && tabla && !tabla.filtro) { tabla.filtro = { columnas: new Map() }; enFiltroT = tabla.filtro;
        } else if (nombre === 'filterColumn' && enFiltroT && enFiltroT.columnas.size < 16384) { colFiltroT = { col: parseInt(a.colId, 10) || 0, oculto: a.hiddenButton === '1' || a.hiddenButton === 'true', filtrada: false }; enFiltroT.columnas.set(colFiltroT.col, colFiltroT);
        } else if (colFiltroT && (nombre === 'filters' || nombre === 'customFilters' || nombre === 'top10' || nombre === 'colorFilter' || nombre === 'dynamicFilter' || nombre === 'iconFilter')) { colFiltroT.filtrada = true;
        } else if (nombre === 'tableStyleInfo' && tabla) {
          tabla.estilo = a.name || null;
          tabla.primeraCol = a.showFirstColumn === '1' || a.showFirstColumn === 'true';
          tabla.ultimaCol = a.showLastColumn === '1' || a.showLastColumn === 'true';
          tabla.bandasFilas = a.showRowStripes === '1' || a.showRowStripes === 'true';
          tabla.bandasCols = a.showColumnStripes === '1' || a.showColumnStripes === 'true';
        }
      },
    });
    return tabla;
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
      if (hoja.partes.comentarios) { // las notas: sus textos y, del VML, su cuadro
        const tc = await textoDe(libro.zip, hoja.partes.comentarios);
        const tv = hoja.partes.dibujoVml ? await textoDe(libro.zip, hoja.partes.dibujoVml) : null;
        if (tc !== null) hoja.comentarios = leerComentarios(tc, tv);
      }
      if (hoja.celdasVm.size) { // imágenes en celdas
        if (!libro._imagenesEnCeldas) libro._imagenesEnCeldas = leerImagenesEnCeldas(libro).catch(() => new Map());
        const porVm = await libro._imagenesEnCeldas;
        for (const [clave, vm] of hoja.celdasVm) { const ruta = porVm.get(vm); if (ruta) hoja.imagenesCelda.set(clave, ruta); }
      }
      const rh = hoja.relaciones.find((r) => r.tipo === 'threadedComment');
      if (rh) { // los hilos reemplazan a la nota de relleno ("[Comentario encadenado]…") que Excel deja para las versiones viejas
        const th = await textoDe(libro.zip, rh.destino);
        if (th !== null) {
          if (!libro._personas) libro._personas = leerPersonas(libro.rutaPersonas ? await textoDe(libro.zip, libro.rutaPersonas) : null);
          hoja.hilos = leerHilos(th, libro._personas);
          const notaDe = new Map(); // la primera nota de cada celda (con find por hilo eran N × M: minutos con 100.000 de cada uno)
          for (const x of hoja.comentarios) if (!notaDe.has(x.f * 16384 + x.c)) notaDe.set(x.f * 16384 + x.c, x);
          for (const h of hoja.hilos) { const n = notaDe.get(h.f * 16384 + h.c); if (n) { n.hilo = h; n.visible = false; } }
        }
      }
      if (hoja.partes.dibujo) { const t = await textoDe(libro.zip, hoja.partes.dibujo); if (t !== null) hoja.anclas = leerAnclas(t); }
      hoja.partes.tablas = hoja.relaciones.filter((r) => r.tipo === 'table').map((r) => r.destino);
      for (const ruta of hoja.partes.tablas.slice(0, MAX_TABLAS)) {
        const t = await textoDe(libro.zip, ruta);
        const tabla = t === null ? null : leerTabla(t);
        if (tabla) hoja.tablas.push(tabla);
      }
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

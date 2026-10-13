/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Estilos de un libro de Excel (.xlsx/.xlsm): styles.xml y theme1.xml resueltos
// a objetos simples que el dibujo y el formato de número pueden usar sin saber
// nada de OOXML. Requiere xml-ligero.js. No depende del navegador.
//
// Uso:
//   const estilos = XlsxEstilos.leer(textoStylesXml, textoThemeXml);   // textos nulos → valores por defecto
//   estilos.xf(indice)            → { numFmtId, formato (código), fuente, relleno, borde, alineacion, proteccion, estiloCelda }
//   estilos.color(ref)            → '#RRGGBB' | null (resuelve tema + tint, indexado y rgb)
//   estilos.dxf(indice)           → formato diferencial (formato condicional, tablas)
//   XlsxEstilos.FORMATOS_INTEGRADOS[id] → código de formato de número integrado (0-49)
//
// La norma (ECMA-376 18.8) dice que los atributos applyX del xf indican si ese
// grupo (fuente, relleno…) anula al del estilo de celda. Aquí se usan siempre
// los ids del propio xf, como hacen los lectores comunes; ver
// Investigaciones\investigacion-motor-excel\03-mediciones.md.

(function () {
  // 18.8.30 numFmt: formatos integrados que no llevan <numFmt> en el archivo.
  const FORMATOS_INTEGRADOS = {
    0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
    14: 'mm-dd-yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yy h:mm',
    37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
    45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
    // Los que la norma deja sin definir (5-8, 41-44) los escribe Excel con su moneda y su idioma; se usa el de en-US.
    5: '"$"#,##0_);("$"#,##0)', 6: '"$"#,##0_);[Red]("$"#,##0)', 7: '"$"#,##0.00_);("$"#,##0.00)', 8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
    41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)', 42: '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)', 44: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
  };

  // 18.8.27 colores indexados por defecto (índices 0-63; 64 y 65 son del sistema).
  const PALETA_INDEXADA = [
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
    '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
    '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
    '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
  ];

  // Tema de Office 2013-2022 (el que usan los libros que no traen el suyo).
  const TEMA_POR_DEFECTO = {
    colores: ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'],
    fuenteTitulos: 'Calibri Light', fuenteCuerpo: 'Calibri',
  };

  const NOMBRES_COLOR_TEMA = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

  const bool = (v, porDefecto) => (v === undefined ? porDefecto : v === '1' || v === 'true');
  const num = (v, porDefecto) => { const n = parseFloat(v); return Number.isFinite(n) ? n : porDefecto; };
  const entero = (v, porDefecto) => { const n = parseInt(v, 10); return Number.isFinite(n) ? n : porDefecto; };

  // Atributos de un <color> → { tipo, valor, tint }. El rgb solo se acepta si son 6 u 8 dígitos hexadecimales:
  // el valor acaba en estilos de página (CSS), así que no puede pasar nada más del archivo.
  function colorRef(a) {
    if (a.rgb !== undefined) {
      if (!/^[0-9A-Fa-f]{6}$|^[0-9A-Fa-f]{8}$/.test(a.rgb)) return null;
      return { tipo: 'rgb', valor: (a.rgb.length === 8 ? a.rgb.slice(2) : a.rgb).toUpperCase(), tint: num(a.tint, 0) };
    }
    if (a.theme !== undefined) return { tipo: 'tema', valor: entero(a.theme, 0), tint: num(a.tint, 0) };
    if (a.indexed !== undefined) return { tipo: 'indexado', valor: entero(a.indexed, 0), tint: num(a.tint, 0) };
    if (a.auto !== undefined) return { tipo: 'auto', valor: 0, tint: 0 };
    return null;
  }

  // ---- Color ----------------------------------------------------------------------------------------

  // El tinte lo aplica Excel como Windows (ColorRGBToHLS / ColorHLSToRGB, con H, L y S enteros de 0 a 240): se cambia solo la luminosidad,
  // con el tinte a 4 decimales (los archivos de Excel traen 0,39997558… por 0,4), la luminosidad truncada y los grises redondeados.
  // Medido contra 43 colores resueltos por Excel (estilos de tabla integrados): exacto en los 43 (el cálculo en coma flotante falla hasta en 3/255).
  function aHls240(R, G, B) {
    const cMax = Math.max(R, G, B), cMin = Math.min(R, G, B), suma = cMax + cMin, d = cMax - cMin;
    const L = Math.floor((suma * 240 + 255) / 510);
    if (!d) return [160, L, 0];
    const S = L <= 120 ? Math.floor((d * 240 + Math.floor(suma / 2)) / suma) : Math.floor((d * 240 + Math.floor((510 - suma) / 2)) / (510 - suma));
    const delta = (c) => Math.floor(((cMax - c) * 40 + Math.floor(d / 2)) / d);
    let H = R === cMax ? delta(B) - delta(G) : G === cMax ? 80 + delta(R) - delta(B) : 160 + delta(G) - delta(R);
    if (H < 0) H += 240;
    if (H > 240) H -= 240;
    return [H, L, S];
  }
  function deHls240(H, L, S) {
    const m2 = L <= 120 ? Math.floor((L * (240 + S) + 120) / 240) : L + S - Math.floor((L * S + 120) / 240);
    const m1 = 2 * L - m2;
    const tono = (h) => {
      if (h < 0) h += 240;
      if (h > 240) h -= 240;
      if (h < 40) return m1 + Math.floor(((m2 - m1) * h + 20) / 40);
      if (h < 120) return m2;
      if (h < 160) return m1 + Math.floor(((m2 - m1) * (160 - h) + 20) / 40);
      return m1;
    };
    return [tono(H + 80), tono(H), tono(H - 80)].map((v) => Math.floor((v * 255 + 120) / 240));
  }
  const hex2 = (v) => v.toString(16).padStart(2, '0');

  // Qué borde gana cuando dos celdas vecinas piden uno distinto para la línea que comparten (medido en Excel): el más pesado; a igual peso, el de color
  // con menor valor BGR; y a igual color, el propio de la celda y no el de un estilo de tabla.
  const PESO_BORDE = { hair: 1, dotted: 2, dashed: 3, thin: 4, dashDotDot: 5, dashDot: 6, mediumDashDotDot: 7, slantDashDot: 8, mediumDashDot: 9, mediumDashed: 10, medium: 11, double: 12, thick: 13 };
  // a y b son bordes { estilo, color, deTabla? } o null; `hex(color)` da '#RRGGBB' (sin él no se compara el color); los empates los gana `a`
  function mejorBorde(a, b, hex) {
    if (!a) return b || null;
    if (!b) return a;
    const wa = PESO_BORDE[a.estilo] || 0, wb = PESO_BORDE[b.estilo] || 0;
    if (wa !== wb) return wb > wa ? b : a;
    if (hex) { // el color con el valor BGR (el entero COLORREF de Windows) menor: negro < rojo < el resto de los de tema (medido: Excel deja el negro de una tabla sobre un rojo propio, y el rojo sobre los demás)
      const bgr = (x) => { const h = hex(x.color); return h ? parseInt(h.slice(5, 7) + h.slice(3, 5) + h.slice(1, 3), 16) : 0; };
      const ca = bgr(a), cb = bgr(b);
      if (ca !== cb) return cb < ca ? b : a;
    }
    return !!a.deTabla && !b.deTabla ? b : a;
  }

  // 18.8.3 color/@tint: se cambia solo la luminosidad del color en HLS.
  function aplicarTint(rgb6, tint) {
    tint = Math.round(tint * 10000) / 10000;
    if (!tint) return rgb6;
    const R = parseInt(rgb6.slice(0, 2), 16), G = parseInt(rgb6.slice(2, 4), 16), B = parseInt(rgb6.slice(4, 6), 16);
    const [H, L, S] = aHls240(R, G, B);
    const L2 = Math.max(0, Math.min(240, Math.floor((tint < 0 ? L * (1 + tint) : L * (1 - tint) + 240 * tint) + 1e-9)));
    const [r, g, b] = S ? deHls240(H, L2, S) : [0, 0, 0].map(() => Math.round(L2 * 255 / 240));
    return (hex2(r) + hex2(g) + hex2(b)).toUpperCase();
  }

  // ---- Tema -----------------------------------------------------------------------------------------

  function leerTema(texto) {
    if (!texto) return TEMA_POR_DEFECTO;
    const porNombre = {};
    let enEsquema = false, esquemaLeido = false, enExtra = false, actual = null, fuenteActual = null;
    const tema = { colores: TEMA_POR_DEFECTO.colores.slice(), fuenteTitulos: TEMA_POR_DEFECTO.fuenteTitulos, fuenteCuerpo: TEMA_POR_DEFECTO.fuenteCuerpo };
    const hex6 = (v) => (typeof v === 'string' && /^[0-9A-Fa-f]{6}$/.test(v) ? v.toUpperCase() : null);
    XmlLigero.recorrer(texto, {
      abrir(nombre, a) {
        if (nombre === 'extraClrSchemeLst') enExtra = true; // esquemas alternativos: no son los del libro
        if (enExtra) return;
        if (nombre === 'clrScheme' && !esquemaLeido) enEsquema = true;
        else if (enEsquema && NOMBRES_COLOR_TEMA.includes(nombre)) actual = nombre;
        else if (actual && nombre === 'srgbClr' && hex6(a.val)) porNombre[actual] = hex6(a.val);
        else if (actual && nombre === 'sysClr') porNombre[actual] = hex6(a.lastClr) || (a.val === 'window' ? 'FFFFFF' : '000000');
        else if (nombre === 'majorFont') fuenteActual = 'fuenteTitulos';
        else if (nombre === 'minorFont') fuenteActual = 'fuenteCuerpo';
        else if (fuenteActual && nombre === 'latin' && a.typeface) { tema[fuenteActual] = a.typeface; fuenteActual = null; }
      },
      cerrar(nombre) {
        if (nombre === 'extraClrSchemeLst') enExtra = false;
        if (nombre === 'clrScheme' && enEsquema) { enEsquema = false; esquemaLeido = true; }
        if (NOMBRES_COLOR_TEMA.includes(nombre)) actual = null;
      },
    });
    // En el archivo, theme="0" es lt1 y "1" es dk1 (y "2" lt2, "3" dk2): van cambiados respecto del orden del esquema.
    const orden = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
    tema.colores = orden.map((nom, i) => porNombre[nom] || TEMA_POR_DEFECTO.colores[i]);
    return tema;
  }

  // ---- styles.xml -----------------------------------------------------------------------------------

  function fuenteVacia() { return { nombre: null, tam: null, negrita: false, cursiva: false, subrayado: null, tachado: false, color: null, alinVertical: null, familia: null, esquema: null }; }

  function leerFormatoFuente(a, destino, nombre) {
    (destino.dado || (destino.dado = new Set())).add(nombre); // qué propiedades trae el archivo (un formato diferencial solo cambia esas)
    switch (nombre) {
      case 'b': destino.negrita = bool(a.val, true); break;
      case 'i': destino.cursiva = bool(a.val, true); break;
      case 'strike': destino.tachado = bool(a.val, true); break;
      case 'u': destino.subrayado = a.val || 'single'; break;
      case 'sz': destino.tam = num(a.val, null); break;
      case 'name': case 'rFont': destino.nombre = a.val; break;
      case 'color': destino.color = colorRef(a); break;
      case 'vertAlign': destino.alinVertical = a.val || null; break;
      case 'family': destino.familia = entero(a.val, null); break;
      case 'scheme': destino.esquema = a.val || null; break;
      default: break;
    }
  }

  function leerStyles(texto) {
    const r = { numFmts: {}, fuentes: [], rellenos: [], bordes: [], xfsEstilo: [], xfs: [], dxfs: [], estilosCelda: [], estilosTabla: new Map(), paletaPropia: null };
    if (!texto) return r;
    let seccion = '';      // numFmts | fonts | fills | borders | cellStyleXfs | cellXfs | cellStyles | dxfs | colors | …
    let estiloTabla = null, fuente = null, relleno = null, borde = null, ladoBorde = null, xf = null, dxf = null, enIndexed = false, enDegradado = false;
    XmlLigero.recorrer(texto, {
      abrir(nombre, a, vacio) {
        switch (nombre) {
          case 'numFmts': case 'fonts': case 'fills': case 'borders': case 'cellStyleXfs': case 'cellXfs': case 'cellStyles': case 'dxfs': case 'colors': case 'tableStyles': case 'extLst':
            seccion = nombre; return;
          case 'numFmt':
            if (dxf) dxf.formato = { id: entero(a.numFmtId, 0), codigo: a.formatCode };
            else if (seccion === 'numFmts') r.numFmts[entero(a.numFmtId, 0)] = a.formatCode;
            return;
          case 'dxf': dxf = { fuente: null, relleno: null, borde: null, formato: null, alineacion: null }; if (vacio) cerrarDxf(); return;
          case 'font': fuente = fuenteVacia(); if (vacio) cerrarFuente(); return;
          case 'fill': relleno = { tipo: 'ninguno', patron: 'none', fg: null, bg: null, degradado: null }; if (vacio) cerrarRelleno(); return;
          case 'patternFill':
            if (relleno) { relleno.patron = a.patternType || (dxf ? 'solid' : 'none'); relleno.tipo = relleno.patron === 'none' ? 'ninguno' : 'patron'; }
            return;
          case 'gradientFill':
            if (relleno) { relleno.tipo = 'degradado'; enDegradado = true; relleno.degradado = { tipo: a.type || 'linear', grados: num(a.degree, 0), izq: num(a.left, 0), der: num(a.right, 0), arr: num(a.top, 0), aba: num(a.bottom, 0), paradas: [] }; }
            return;
          case 'stop': if (relleno && relleno.degradado) relleno.degradado.paradas.push({ pos: num(a.position, 0), color: null }); return;
          case 'fgColor': if (relleno) relleno.fg = colorRef(a); return;
          case 'bgColor': if (relleno) relleno.bg = colorRef(a); return;
          case 'border': borde = { izq: null, der: null, arr: null, aba: null, vert: null, horiz: null, diagonal: null, diagArriba: bool(a.diagonalUp, false), diagAbajo: bool(a.diagonalDown, false) }; if (vacio) cerrarBorde(); return;
          case 'tableStyle': if (seccion === 'tableStyles' && a.name && r.estilosTabla.size < 5000) { estiloTabla = { nombre: a.name, elementos: {} }; r.estilosTabla.set(a.name, estiloTabla); } return;
          case 'tableStyleElement': if (estiloTabla && a.type) estiloTabla.elementos[a.type] = { dxf: a.dxfId !== undefined ? entero(a.dxfId, null) : null, tam: entero(a.size, 1) }; return;
          case 'left': case 'start': case 'right': case 'end': case 'top': case 'bottom': case 'diagonal': case 'vertical': case 'horizontal':
            if (borde) { ladoBorde = { left: 'izq', start: 'izq', right: 'der', end: 'der', top: 'arr', bottom: 'aba', diagonal: 'diagonal', vertical: 'vert', horizontal: 'horiz' }[nombre]; borde[ladoBorde] = a.style && a.style !== 'none' ? { estilo: a.style, color: null } : null; if (vacio) ladoBorde = null; }
            return;
          case 'color':
            if (enIndexed) return;
            if (enDegradado && relleno && relleno.degradado.paradas.length) relleno.degradado.paradas[relleno.degradado.paradas.length - 1].color = colorRef(a);
            else if (borde && ladoBorde && borde[ladoBorde]) borde[ladoBorde].color = colorRef(a);
            else if (fuente) fuente.color = colorRef(a);
            return;
          case 'indexedColors': enIndexed = true; r.paletaPropia = []; return;
          case 'rgbColor': if (enIndexed) { const c = colorRef({ rgb: a.rgb || '' }); r.paletaPropia.push(c ? c.valor : '000000'); } return;
          case 'xf': {
            xf = { numFmtId: entero(a.numFmtId, 0), fuenteId: entero(a.fontId, 0), rellenoId: entero(a.fillId, 0), bordeId: entero(a.borderId, 0), xfId: entero(a.xfId, 0),
              alineacion: null, proteccion: null, comillaPrefijo: bool(a.quotePrefix, false), botonDinamica: bool(a.pivotButton, false) };
            if (seccion === 'cellStyleXfs') r.xfsEstilo.push(xf); else if (seccion === 'cellXfs') r.xfs.push(xf);
            if (vacio) xf = null;
            return;
          }
          case 'alignment': {
            const al = { h: a.horizontal || 'general', v: a.vertical || 'bottom', sangria: entero(a.indent, 0), rotacion: entero(a.textRotation, 0), ajustar: bool(a.wrapText, false), reducir: bool(a.shrinkToFit, false), ordenLectura: entero(a.readingOrder, 0), justificarUltima: bool(a.justifyLastLine, false), sangriaRelativa: entero(a.relativeIndent, 0) };
            if (xf) xf.alineacion = al; else if (dxf) dxf.alineacion = al;
            return;
          }
          case 'protection': if (xf) xf.proteccion = { bloqueada: bool(a.locked, true), oculta: bool(a.hidden, false) }; return;
          case 'cellStyle': if (seccion === 'cellStyles') r.estilosCelda.push({ nombre: a.name, xfId: entero(a.xfId, 0), integrado: a.builtinId !== undefined ? entero(a.builtinId, null) : null }); return;
          default: if (fuente) leerFormatoFuente(a, fuente, nombre);
        }
      },
      cerrar(nombre) {
        switch (nombre) {
          case 'font': cerrarFuente(); break;
          case 'fill': cerrarRelleno(); break;
          case 'border': cerrarBorde(); break;
          case 'gradientFill': enDegradado = false; break;
          case 'left': case 'start': case 'right': case 'end': case 'top': case 'bottom': case 'diagonal': case 'vertical': case 'horizontal': ladoBorde = null; break;
          case 'tableStyle': estiloTabla = null; break;
          case 'xf': xf = null; break;
          case 'dxf': cerrarDxf(); break;
          case 'indexedColors': enIndexed = false; break;
          default: break;
        }
      },
    });

    function cerrarFuente() { if (fuente) { if (dxf) dxf.fuente = fuente; else if (seccion === 'fonts') r.fuentes.push(fuente); } fuente = null; }
    function cerrarRelleno() { if (relleno) { if (dxf) dxf.relleno = relleno; else if (seccion === 'fills') r.rellenos.push(relleno); } relleno = null; enDegradado = false; }
    function cerrarBorde() { if (borde) { if (dxf) dxf.borde = borde; else if (seccion === 'borders') r.bordes.push(borde); } borde = null; ladoBorde = null; }
    function cerrarDxf() { if (dxf) r.dxfs.push(dxf); dxf = null; }
    return r;
  }

  // ---- API ------------------------------------------------------------------------------------------

  function leer(textoStyles, textoTema) {
    const crudo = leerStyles(textoStyles);
    const tema = leerTema(textoTema);
    const paleta = crudo.paletaPropia && crudo.paletaPropia.length ? crudo.paletaPropia : PALETA_INDEXADA;
    const memoColor = new Map();
    const memoXf = [], dxfsEnLinea = new Map();

    // Un libro sin styles.xml (o con listas vacías) usa el estilo 0: Calibri 11, sin relleno ni bordes.
    const fuentePorDefecto = crudo.fuentes[0] || Object.assign(fuenteVacia(), { nombre: tema.fuenteCuerpo, tam: 11 });

    function color(ref) {
      if (!ref) return null;
      const clave = ref.tipo + ref.valor + ':' + ref.tint;
      if (memoColor.has(clave)) return memoColor.get(clave);
      let rgb = null;
      if (ref.tipo === 'rgb') rgb = ref.valor;
      else if (ref.tipo === 'tema') rgb = tema.colores[ref.valor] || null;
      else if (ref.tipo === 'indexado') rgb = ref.valor === 64 ? '000000' : ref.valor === 65 ? 'FFFFFF' : (paleta[ref.valor] || null);
      else if (ref.tipo === 'auto') rgb = null;
      const final = rgb ? '#' + aplicarTint(rgb.toUpperCase(), ref.tint) : null;
      memoColor.set(clave, final);
      return final;
    }

    function formatoDe(id) {
      if (crudo.numFmts[id] !== undefined) return crudo.numFmts[id];
      return FORMATOS_INTEGRADOS[id] !== undefined ? FORMATOS_INTEGRADOS[id] : 'General';
    }

    function xf(indice) {
      if (memoXf[indice]) return memoXf[indice];
      const x = crudo.xfs[indice] || crudo.xfs[0] || { numFmtId: 0, fuenteId: 0, rellenoId: 0, bordeId: 0, xfId: 0, alineacion: null, proteccion: null, comillaPrefijo: false };
      const estiloCelda = crudo.estilosCelda.find((s) => s.xfId === x.xfId) || null;
      const res = {
        numFmtId: x.numFmtId, formato: formatoDe(x.numFmtId),
        fuente: crudo.fuentes[x.fuenteId] || fuentePorDefecto,
        relleno: crudo.rellenos[x.rellenoId] || { tipo: 'ninguno', patron: 'none', fg: null, bg: null, degradado: null },
        borde: crudo.bordes[x.bordeId] || null,
        alineacion: x.alineacion || { h: 'general', v: 'bottom', sangria: 0, rotacion: 0, ajustar: false, reducir: false, ordenLectura: 0, justificarUltima: false, sangriaRelativa: 0 },
        proteccion: x.proteccion || { bloqueada: true, oculta: false },
        comillaPrefijo: x.comillaPrefijo,
        estiloCelda: estiloCelda ? estiloCelda.nombre : null,
      };
      memoXf[indice] = res;
      return res;
    }

    return {
      xf, color, tema, formatoDe,
      cantidadXfs: crudo.xfs.length,
      fuentes: crudo.fuentes, rellenos: crudo.rellenos, bordes: crudo.bordes,
      dxf: (i) => crudo.dxfs[i] || null,
      dxfDeXml: (xml) => { // un <dxf> escrito en línea (extensión x14 de las reglas condicionales), leído con el mismo lector de estilos
        if (!dxfsEnLinea.has(xml)) dxfsEnLinea.set(xml, leerStyles('<styleSheet><dxfs>' + xml + '</dxfs></styleSheet>').dxfs[0] || null);
        return dxfsEnLinea.get(xml);
      },
      dxfs: crudo.dxfs,
      estiloTabla: (nombre) => crudo.estilosTabla.get(nombre) || null,   // { nombre, elementos: { wholeTable: { dxf, tam }, … } } de los estilos propios del libro
      formatosPropios: crudo.numFmts,
      paletaPropia: crudo.paletaPropia,
    };
  }

  globalThis.XlsxEstilos = { leer, colorRef, aplicarTint, PESO_BORDE, mejorBorde, FORMATOS_INTEGRADOS, PALETA_INDEXADA, TEMA_POR_DEFECTO };
})();

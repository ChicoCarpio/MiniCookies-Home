/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Lector de libros de Excel binarios (.xlsb) del motor propio de Excel (etapa 9). Un .xlsb es un zip como el .xlsx, pero sus partes (workbook.bin,
// sheetN.bin, sharedStrings.bin, styles.bin) son listas de registros binarios ([MS-XLSB]). Requiere zip-propio.js, xml-ligero.js, xlsx-modelo.js
// (errores) y xlsx-estilos.js (formatos de número integrados); no depende del navegador.
//
//   const libro = await XlsbLector.leer(bytes);
//   libro = { hojas: [{ nombre, filas: [[valor…]…], anchos: [caracteres…], combinadas: ['A1:B2'…], altosFila: { fila: puntos } }], fecha1904 }   (valor: ver OdsLector)
//
// Estilos de celda: styles.bin (fuentes, tramas, bordes y xfs) y theme1.xml se leen y cada celda con estilo propio sale como { valor, formato?, estilo }
// (estilo: ver XlsxEscritor.crearLibro), con UN objeto por xf compartido por todas sus celdas; una celda vacía con estilo, como { valor: null, estilo }.
// Los colores se resuelven aquí (indexados, del tema con su tinte, RGB). El alto de una fila con alto propio (BrtRowHdr, fUnsynced) va en altosFila.
// Las fórmulas no se leen (queda su último valor, que el archivo guarda). Las hojas de gráfico y de macros se omiten.
(function () {
  const MAX_FILAS = 1048576, MAX_COLS = 16384, MAX_TEXTO = 32767;
  const MAX_CELDAS = 10000000;   // en todo el libro: crearLibro no arma un .xlsx de más de ~8 M celdas, y antes de llegar ahí la memoria ya se fue
  const MAX_COMBINADAS = 500000; // por hoja, como LibroXlsx
  const ERRORES = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0F: '#VALUE!', 0x17: '#REF!', 0x1D: '#NAME?', 0x24: '#NUM!', 0x2A: '#N/A', 0x2B: '#GETTING_DATA' };
  const danado = (detalle) => new LibroXlsx.ErrorXlsx('danado', detalle);

  // Registros: tipo (1-2 bytes, 7 bits cada uno) y tamaño (1-4 bytes, 7 bits cada uno).
  function* registros(b) {
    let p = 0;
    const n = b.length;
    while (p < n) {
      let id = b[p++];
      if (id & 0x80) { if (p >= n) return; id = (id & 0x7F) | (b[p++] << 7); }
      let tam = 0, mult = 1, k = 0, byte;
      do {
        if (p >= n || k++ > 3) return;
        byte = b[p++];
        tam += (byte & 0x7F) * mult; mult *= 128;
      } while (byte & 0x80);
      if (p + tam > n) return;   // registro cortado: se ignora el resto
      yield { id, p, tam };
      p += tam;
    }
  }

  const leerU32 = (b, p) => (b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24)) >>> 0;
  const leerU16 = (b, p) => b[p] | (b[p + 1] << 8);
  const texto = new TextDecoder('utf-16le');
  // XLWideString: 4 bytes de cantidad de caracteres + UTF-16; devuelve [cadena, siguiente posición]
  function cadena(b, p, fin) {
    const cch = leerU32(b, p);
    if (cch === 0xFFFFFFFF) return ['', p + 4];
    if (p + 4 + cch * 2 > fin) throw danado('cadena cortada');
    return [texto.decode(b.subarray(p + 4, p + 4 + cch * 2)), p + 4 + cch * 2];
  }

  const vistaDouble = new DataView(new ArrayBuffer(8));
  function numeroRk(v) {
    let n;
    if (v & 2) n = v >> 2;   // entero de 30 bits con signo
    else { vistaDouble.setUint32(0, 0, true); vistaDouble.setUint32(4, (v & 0xFFFFFFFC) >>> 0, true); n = vistaDouble.getFloat64(0, true); }
    return v & 1 ? n / 100 : n;
  }

  const ruta = (base, destino) => {
    if (destino.startsWith('/')) return destino.slice(1);
    const partes = base.split('/'); partes.pop();
    for (const s of destino.split('/')) { if (s === '..') partes.pop(); else if (s !== '.') partes.push(s); }
    return partes.join('/');
  };

  async function relaciones(zip, parte) {
    const dir = parte.slice(0, parte.lastIndexOf('/') + 1), nombre = parte.slice(parte.lastIndexOf('/') + 1);
    const e = zip.file(dir + '_rels/' + nombre + '.rels');
    const m = new Map();
    if (!e) return m;
    XmlLigero.recorrer(await e.async('string'), { abrir(n, a) { if (n === 'Relationship' && a.Id) m.set(a.Id, { tipo: a.Type || '', destino: ruta(parte, a.Target || '') }); } });
    return m;
  }

  async function leer(bytes) {
    let zip;
    try { zip = await ZipPropio.loadAsync(bytes); } catch (e) { throw new LibroXlsx.ErrorXlsx('noZip', 'no es un .xlsb'); }
    const libroBin = zip.file('xl/workbook.bin');
    if (!libroBin) throw new LibroXlsx.ErrorXlsx('sinLibro');
    const wb = new Uint8Array(await libroBin.async('uint8array'));
    const rels = await relaciones(zip, 'xl/workbook.bin');
    let fecha1904 = false;
    const hojasDef = [];
    for (const r of registros(wb)) {
      if (r.id === 153) fecha1904 = (leerU32(wb, r.p) & 1) === 1;   // BrtWbProp
      else if (r.id === 156) {                                      // BrtBundleSh: estado, id, relación, nombre
        const fin = r.p + r.tam;
        let q = r.p + 8;
        const [rid, q2] = cadena(wb, q, fin);
        const [nombre] = cadena(wb, q2, fin);
        hojasDef.push({ nombre, rid, estado: leerU32(wb, r.p) === 1 ? 'oculta' : leerU32(wb, r.p) === 2 ? 'muyOculta' : 'visible' });
      }
    }
    // textos compartidos y formatos
    let rutaTextos = 'xl/sharedStrings.bin', rutaEstilos = 'xl/styles.bin';
    for (const rel of rels.values()) { if (/\/sharedStrings$/.test(rel.tipo)) rutaTextos = rel.destino; else if (/\/styles$/.test(rel.tipo)) rutaEstilos = rel.destino; }
    const textos = [];
    if (zip.file(rutaTextos)) {
      const b = new Uint8Array(await zip.file(rutaTextos).async('uint8array'));
      for (const r of registros(b)) if (r.id === 19) textos.push(cadena(b, r.p + 1, r.p + r.tam)[0].slice(0, MAX_TEXTO));   // BrtSSTItem
    }
    const codigos = new Map();    // id de formato personalizado → código
    const xfs = [];               // cellXfs: { fmt, fuente, relleno, borde, trot, sangria, h, v, ajustar }
    const fuentes = [], rellenos = [], bordes = [];
    let paletaPropia = null;
    const MAX_ESTILOS = 65536;    // Excel admite 64.000 xfs: de más no se leen, para que un archivo hostil no agrande la memoria
    // Tema: theme="0" es lt1 y "1" es dk1 (y "2" lt2, "3" dk2): van cambiados respecto del orden del esquema
    const ordenTema = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
    const porNombre = {};
    const relTema = [...rels.values()].find((rel) => /\/theme$/.test(rel.tipo));
    const entradaTema = zip.file(relTema ? relTema.destino : 'xl/theme/theme1.xml');
    if (entradaTema) {
      let actual = null, enExtra = false;
      XmlLigero.recorrer(await entradaTema.async('string'), {
        abrir(n, a) {
          if (n === 'extraClrSchemeLst') enExtra = true;
          if (enExtra) return;
          if (ordenTema.includes(n)) actual = n;
          else if (actual && n === 'srgbClr' && /^[0-9A-Fa-f]{6}$/.test(a.val || '')) porNombre[actual] = a.val.toUpperCase();
          else if (actual && n === 'sysClr') porNombre[actual] = /^[0-9A-Fa-f]{6}$/.test(a.lastClr || '') ? a.lastClr.toUpperCase() : a.val === 'window' ? 'FFFFFF' : '000000';
        },
        cerrar(n) { if (n === 'extraClrSchemeLst') enExtra = false; if (ordenTema.includes(n)) actual = null; },
      });
    }
    const tema = ordenTema.map((n, i) => porNombre[n] || XlsxEstilos.TEMA_POR_DEFECTO.colores[i]);
    let paleta = XlsxEstilos.PALETA_INDEXADA;
    // BrtColor (8 bytes): bandera y tipo (0 automático, 1 indexado, 2 RGB, 3 tema), índice, tinte (int16/32767), RGBA. Devuelve 'RRGGBB' o null (automático / del sistema).
    const hex6 = (r, g, bl) => ((r << 16) | (g << 8) | bl).toString(16).toUpperCase().padStart(6, '0');
    const color = (b, p) => {
      const tipo = b[p] >> 1, idx = b[p + 1];
      let rgb = null;
      if (tipo === 1) rgb = idx < 64 ? paleta[idx] : idx === 65 ? 'FFFFFF' : null;
      else if (tipo === 2) rgb = hex6(b[p + 4], b[p + 5], b[p + 6]);
      else if (tipo === 3) rgb = idx < tema.length ? tema[idx] : null;
      if (!rgb) return null;
      const tint = ((b[p + 2] | (b[p + 3] << 8)) << 16) >> 16;
      return tint ? XlsxEstilos.aplicarTint(rgb, tint / 32767) : rgb;
    };
    if (zip.file(rutaEstilos)) {
      const b = new Uint8Array(await zip.file(rutaEstilos).async('uint8array'));
      let enXfs = false, enPaleta = false;
      for (const r of registros(b)) {
        const p = r.p;
        if (r.id === 44) codigos.set(leerU16(b, p), cadena(b, p + 2, p + r.tam)[0]);          // BrtFmt
        else if (r.id === 563) { enPaleta = true; paletaPropia = []; }                         // BrtBeginIndexedColors
        else if (r.id === 565) { enPaleta = false; if (paletaPropia && paletaPropia.length) paleta = XlsxEstilos.PALETA_INDEXADA.map((c, i) => paletaPropia[i] || c); }
        else if (r.id === 564) { if (enPaleta && r.tam >= 4 && paletaPropia.length < 64) paletaPropia.push(hex6(b[p], b[p + 1], b[p + 2])); }   // BrtIndexedColor: RGBA
        else if (r.id === 43 && r.tam >= 21 && fuentes.length < MAX_ESTILOS) {                // BrtFont: alto (twips), banderas, peso, supra/subíndice, subrayado, …, color (en 12), esquema, nombre (en 21)
          const graf = leerU16(b, p + 2), uls = b[p + 8];
          fuentes.push({ tam: leerU16(b, p) / 20, negrita: leerU16(b, p + 4) >= 600 || (graf & 1) === 1, cursiva: (graf & 2) !== 0, tachado: (graf & 8) !== 0, subrayado: uls === 2 || uls === 34 ? 'double' : uls ? true : false, p, fin: p + r.tam, b });
        } else if (r.id === 45 && r.tam >= 20 && rellenos.length < MAX_ESTILOS) rellenos.push({ fls: leerU32(b, p), p, b });   // BrtFill: trama, color de frente (en 4) y de fondo (en 12)
        else if (r.id === 46 && r.tam >= 51 && bordes.length < MAX_ESTILOS) bordes.push({ p, b });                          // BrtBorder
        else if (r.id === 617) enXfs = true;                                                   // BrtBeginCellXFs
        else if (r.id === 618) enXfs = false;
        else if (r.id === 47 && enXfs && r.tam >= 16 && xfs.length < MAX_ESTILOS) {           // BrtXF: padre, formato, fuente, trama, borde, rotación, sangría, banderas
          const graf = leerU16(b, p + 12);
          xfs.push({ fmt: leerU16(b, p + 2), fuente: leerU16(b, p + 4), relleno: leerU16(b, p + 6), borde: leerU16(b, p + 8), trot: b[p + 10], sangria: b[p + 11], h: graf & 7, v: (graf >> 3) & 7, ajustar: (graf & 0x40) !== 0 });
        }
      }
    }
    const xfFormato = xfs.map((x) => x.fmt);
    const H = [null, 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
    const V = ['top', 'center', 'bottom', 'justify', 'distributed'];
    const BORDES = [null, 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot'];
    const PATRONES = [null, 'solid', 'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis', 'gray125', 'gray0625'];
    // El estilo de un xf: UN solo objeto que comparten todas sus celdas (crearLibro reparte los estilos por identidad). null si es el de fábrica (xf 0).
    const estilos = new Map();
    function estiloDe(xf) {
      if (xf === 0) return null;
      if (estilos.has(xf)) return estilos.get(xf);
      const x = xfs[xf];
      let e = null;
      if (x) {
        e = {};
        const f = fuentes[x.fuente];
        if (f) {
          const o = {};
          let nombre = '';
          try { if (f.p + 25 <= f.fin) nombre = cadena(f.b, f.p + 21, f.fin)[0]; } catch (e) { /* nombre cortado: queda sin nombre */ }
          if (nombre) o.nombre = nombre.slice(0, 31);
          if (f.tam >= 1 && f.tam <= 409) o.tam = f.tam;
          if (f.negrita) o.negrita = true;
          if (f.cursiva) o.cursiva = true;
          if (f.tachado) o.tachado = true;
          if (f.subrayado) o.subrayado = f.subrayado;
          const c = color(f.b, f.p + 12);
          if (c) o.color = c;
          e.fuente = o;
        }
        const rl = rellenos[x.relleno];
        if (rl && rl.fls > 0 && rl.fls < PATRONES.length) {
          const fg = color(rl.b, rl.p + 4), bg = color(rl.b, rl.p + 12);
          if (rl.fls === 1) e.relleno = fg || '000000';   // color automático = negro, como lo pinta Excel
          else if (fg || bg) e.relleno = { patron: PATRONES[rl.fls], ...(fg ? { fg } : {}), ...(bg ? { bg } : {}) };
        }
        const bd = bordes[x.borde];
        if (bd) {
          const o = {};
          // arriba, abajo, izquierda, derecha y diagonal, en ese orden: cada uno son el estilo (2 bytes) y el color (8)
          [['arr', 1], ['aba', 11], ['izq', 21], ['der', 31], ['diag', 41]].forEach(([lado, d]) => {
            const dg = leerU16(bd.b, bd.p + d);
            if (!BORDES[dg]) return;
            const c = color(bd.b, bd.p + d + 2);
            const linea = c ? { e: BORDES[dg], c } : { e: BORDES[dg] };
            if (lado !== 'diag') o[lado] = linea;
            else { if (bd.b[bd.p] & 1) o.diagA = linea; if (bd.b[bd.p] & 2) o.diagB = linea; }
          });
          if (Object.keys(o).length) e.bordes = o;
        }
        const a = {};
        if (H[x.h]) a.h = H[x.h];
        if (x.v !== 2 && V[x.v]) a.v = V[x.v];
        if (x.ajustar) a.ajustar = true;
        if (x.sangria > 0) a.sangria = x.sangria;
        if ((x.trot > 0 && x.trot <= 180) || x.trot === 255) a.rotacion = x.trot;
        if (Object.keys(a).length) e.alinea = a;
        if (!Object.keys(e).length) e = null;
      }
      estilos.set(xf, e);
      return e;
    }
    // un formato integrado sale como su número (crearLibro lo escribe tal cual y Excel lo muestra según el idioma de quien abre el libro)
    const codigoDe = (xf) => {
      const id = xfFormato[xf];
      if (id === undefined || id === 0) return null;
      if (codigos.has(id)) return codigos.get(id) === 'General' ? null : codigos.get(id);
      return XlsxEstilos.FORMATOS_INTEGRADOS[id] !== undefined && id < 164 ? id : null;
    };

    const hojas = [];
    // una parte se lee una sola vez: un libro hostil con mil hojas sobre la misma parte de celdas multiplicaba la memoria por mil
    const partesLeidas = new Set();
    let celdas = 0;
    for (const def of hojasDef) {
      const rel = rels.get(def.rid);
      if (!rel || !/\/worksheet$/.test(rel.tipo) || partesLeidas.has(rel.destino)) continue;
      const entrada = zip.file(rel.destino);
      if (!entrada) continue;
      partesLeidas.add(rel.destino);
      const b = new Uint8Array(await entrada.async('uint8array'));
      const filas = [], anchos = [], combinadas = [], columnasOcultas = new Set(), filasOcultas = [], altosFila = {};
      let fila = 0;
      const poner = (c, valor) => {
        if (fila >= MAX_FILAS || c >= MAX_COLS) return;
        if (++celdas > MAX_CELDAS) throw new LibroXlsx.ErrorXlsx('grande', 'más de ' + MAX_CELDAS + ' celdas');
        const f = filas[fila] || (filas[fila] = []);
        f[c] = valor;
      };
      // con estilo propio la celda es { valor, formato?, estilo }; sin él, el valor suelto (o con su formato de número)
      const celda = (c, xf, v, esNumero) => {
        const est = estiloDe(xf), cod = esNumero ? codigoDe(xf) : null;
        if (est) poner(c, cod ? { valor: v, formato: cod, estilo: est } : { valor: v, estilo: est });
        else poner(c, cod ? { valor: v, formato: cod } : v);
      };
      const numero = (c, xf, n) => celda(c, xf, n, true);
      // BrtColInfo: en un libro real los rangos no se pisan (entre todos, 16.384 columnas); uno hostil repetía "A:XFD" un millón de veces en 1 KB comprimido
      let presupuestoCols = 4 * MAX_COLS;
      let sigCol = 0;   // la columna que sigue a la última celda: la usan los registros "cortos" (BrtShort*), que no la traen
      for (const r of registros(b)) {
        const p = r.p;
        // celdas: 1-11 traen su columna y su estilo (8 bytes) antes del valor; 12-18 (BrtShortBlank … BrtShortIsst, las del mismo tipo que 1-7) solo el estilo (4 bytes)
        let clave = 1000 + r.id, col = 0, xf = 0, v = 0;   // clave: 1-11 una celda (con su tipo de valor); 1000 + id, cualquier otro registro
        if (r.id >= 1 && r.id <= 11) { clave = r.id; col = leerU32(b, p); xf = leerU32(b, p + 4) & 0xFFFFFF; v = p + 8; sigCol = col + 1; }
        else if (r.id >= 12 && r.id <= 18) { clave = r.id - 11; col = sigCol++; xf = leerU32(b, p) & 0xFFFFFF; v = p + 4; }
        switch (clave) {
          case 1000:                                                            // BrtRowHdr: fila, estilo, alto, banderas (fDyZero, el bit 0x10 del byte 11, es "oculta")
            fila = leerU32(b, p); sigCol = 0;
            if (r.tam >= 12 && (b[p + 11] & 0x10) && fila < MAX_FILAS) filasOcultas.push(fila);
            // alto propio: fUnsynced (bit 5 del byte 11), con miyRw en veinteavos de punto; las demás filas tienen el alto automático
            if (r.tam >= 12 && (b[p + 11] & 0x20) && fila < MAX_FILAS && leerU16(b, p + 8) > 0) altosFila[fila] = leerU16(b, p + 8) / 20;
            break;
          case 1:                                                                // BrtCellBlank: solo cuenta si tiene estilo propio
            if (estiloDe(xf)) poner(col, { valor: null, estilo: estiloDe(xf) });
            break;
          case 2: numero(col, xf, numeroRk(leerU32(b, v))); break;               // BrtCellRk
          case 5: case 9: {                                                     // BrtCellReal, BrtFmlaNum
            vistaDouble.setUint32(0, leerU32(b, v), true); vistaDouble.setUint32(4, leerU32(b, v + 4), true);
            const n = vistaDouble.getFloat64(0, true);
            if (Number.isFinite(n)) numero(col, xf, n);
            break;
          }
          case 3: case 11: celda(col, xf, { error: ERRORES[b[v]] || '#VALUE!' }, false); break;   // BrtCellError, BrtFmlaError
          case 4: case 10: celda(col, xf, b[v] === 1, false); break;                        // BrtCellBool, BrtFmlaBool
          case 6: case 8: {                                                     // BrtCellSt, BrtFmlaString
            const t = cadena(b, v, p + r.tam)[0].slice(0, MAX_TEXTO);
            if (t !== '') celda(col, xf, t, false);
            break;
          }
          case 7: { const t = textos[leerU32(b, v)]; if (t) celda(col, xf, t, false); break; }   // BrtCellIsst
          case 1176: {                                                           // BrtMergeCell: filas y columnas, de la primera a la última
            const f0 = leerU32(b, p), f1 = leerU32(b, p + 4), c0 = leerU32(b, p + 8), c1 = leerU32(b, p + 12);
            if (f1 >= f0 && c1 >= c0 && f1 < MAX_FILAS && c1 < MAX_COLS && (f1 > f0 || c1 > c0) && combinadas.length < MAX_COMBINADAS) combinadas.push(LibroXlsx.numeroACol(c0) + (f0 + 1) + ':' + LibroXlsx.numeroACol(c1) + (f1 + 1));
            break;
          }
          case 1060: {                                                            // BrtColInfo: ancho en 1/256 de carácter, oculta en el bit 0 de las banderas
            const c0 = leerU32(b, p), c1 = Math.min(leerU32(b, p + 4), MAX_COLS - 1), w = leerU32(b, p + 8) / 256, oculta = (leerU16(b, p + 16) & 1) === 1;
            if ((presupuestoCols -= Math.max(0, c1 - c0 + 1)) < 0) break;
            for (let c = c0; c <= c1 && c < MAX_COLS; c++) { anchos[c] = w; if (oculta) columnasOcultas.add(c); }
            break;
          }
          default: break;
        }
      }
      let ancho = 0;
      for (let r = 0; r < filas.length; r++) { if (!filas[r]) filas[r] = []; if (filas[r].length > ancho) ancho = filas[r].length; }
      hojas.push({ nombre: def.nombre, filas, anchos: Array.from(anchos.slice(0, ancho), (w) => w || 0), combinadas, estado: def.estado, columnasOcultas: [...columnasOcultas].filter((c) => c < ancho).sort((a, b) => a - b), filasOcultas: filasOcultas.filter((r) => r < filas.length), altosFila });
    }
    if (!hojas.length) throw new LibroXlsx.ErrorXlsx('sinLibro', 'sin hojas');
    return { hojas, fecha1904 };
  }

  globalThis.XlsbLector = { leer, _registros: registros };
})();

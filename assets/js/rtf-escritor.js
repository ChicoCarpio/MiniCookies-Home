/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Escritor de .rtf (Rich Text Format) desde el «contrato de documento» que describe docx-desde-contrato.js. RTF es el formato de texto con formato que casi todo programa abre
// (Word, LibreOffice, WordPad, Google Documentos, Apple Pages): sirve para entregar un documento a quien no sabe qué versión tiene.
//
//   const { bytes, mensajes } = await RtfEscritor.generar(contrato);     // Uint8Array con el .rtf (ASCII puro: los demás caracteres van como \uN)
//
// Lleva: estilos con nombre (párrafo y carácter), formato de carácter y de párrafo, tabulaciones, listas con niveles, secciones (hoja, márgenes, columnas, tipo) con su encabezado y
// pie, tablas (anidadas, combinadas, bordes, relleno), imágenes en línea (PNG y JPEG), enlaces, campos (número de página, total de páginas, fecha), marcadores y saltos.
// mensajes: 'imagen-formato' (una imagen que no es png ni jpeg se omite), 'enlace-descartado' (un destino que no es una dirección web).
// No depende del navegador.
(function () {
  'use strict';
  const MAX_PROFUNDIDAD_TABLAS = 8;
  const MAX_SECCIONES = 4096;
  const sinProto = (o) => Object.assign(Object.create(null), o);
  const entero = (v, min, max, def = 0) => { const n = Math.round(Number(v)); return Number.isFinite(n) ? Math.max(min, Math.min(max, n)) : def; };
  const rgb = (hex) => (/^[0-9a-fA-F]{6}$/.test(hex || '') ? [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)] : null);

  const ALINEACION = sinProto({ izquierda: '\\ql', centro: '\\qc', derecha: '\\qr', justificado: '\\qj' });
  const TAB_TIPO = sinProto({ izquierda: '', centro: '\\tqc', derecha: '\\tqr', decimal: '\\tqdec' });
  const TAB_RELLENO = sinProto({ punto: '\\tldot', guion: '\\tlhyph', linea: '\\tlul' });
  const NFC = sinProto({ decimal: 0, upperRoman: 1, lowerRoman: 2, upperLetter: 3, lowerLetter: 4, ordinal: 5, cardinalText: 6, ordinalText: 7, decimalZero: 22, bullet: 23, none: 255 });
  const BORDE = sinProto({ simple: '\\brdrs', doble: '\\brdrdb', punteado: '\\brdrdot', guiones: '\\brdrdash' });
  const RESALTADO = sinProto({ black: '000000', blue: '0000FF', cyan: '00FFFF', green: '00FF00', magenta: 'FF00FF', red: 'FF0000', yellow: 'FFFF00', white: 'FFFFFF', darkBlue: '000080', darkCyan: '008080', darkGreen: '008000', darkMagenta: '800080', darkRed: '800000', darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0' });
  const VINETAS = sinProto({ '\uF0B7': '\u2022', '\uF0A7': '\u25AA', '\uF0D8': '\u27A2', '\uF0FC': '\u2713', '\uF076': '\u2756', '\uF0A8': '\u2610', o: '\u25CB', '\u00A7': '\u25AA' });

  function tipoDeImagen(b) {
    if (b.length > 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4E && b[3] === 0x47) return 'png';
    if (b.length > 3 && b[0] === 0xFF && b[1] === 0xD8 && b[2] === 0xFF) return 'jpeg';
    return null;
  }

  // Texto → RTF: \ { } escapados y todo lo que no es ASCII como \uN? (con N de 16 bits con signo)
  function textoRtf(s) {
    let x = '';
    for (let i = 0; i < s.length; i++) {
      const c = s.charCodeAt(i);
      if (c === 92) x += '\\\\';
      else if (c === 123) x += '\\{';
      else if (c === 125) x += '\\}';
      else if (c === 9) x += '\\tab ';
      else if (c === 10) x += '\\line ';
      else if (c === 0xA0) x += '\\~';
      else if (c === 0x2011) x += '\\_';
      else if (c === 0xAD) x += '\\-';
      else if (c < 32 || c === 127) continue;
      else if (c < 128) x += s[i];
      else x += `\\u${c > 32767 ? c - 65536 : c}?`;
    }
    return x;
  }
  // Nombre de marcador (y ancla de enlace, que tiene que coincidir): cortado a 40 como Word, sin lo que rompe la instrucción del campo, y con lo que no es ASCII como \uN
  const nombreMarca = (n) => textoRtf(String(n).slice(0, 40).replace(/[\\"{}]/g, ''));
  const hexa = (bytes) => {
    let s = '';
    for (let i = 0; i < bytes.length; i += 64) {
      let linea = '';
      for (let k = i; k < Math.min(bytes.length, i + 64); k++) linea += (bytes[k] < 16 ? '0' : '') + bytes[k].toString(16);
      s += linea + '\n';
    }
    return s;
  };

  async function generar(contrato) {
    const mensajes = [];
    const aviso = (m) => { if (!mensajes.includes(m)) mensajes.push(m); };

    // ───────── Fuentes y colores ─────────
    const fuentes = [], fuenteIdx = new Map();
    const ftc = (nombre) => {
      const n = String(nombre || 'Times New Roman').replace(/[;{}\\]/g, '').slice(0, 60) || 'Times New Roman';
      if (!fuenteIdx.has(n)) { fuenteIdx.set(n, fuentes.length); fuentes.push(n); }
      return fuenteIdx.get(n);
    };
    const colores = [], colorIdx = new Map();   // el índice 0 es «automático»
    const color = (hex) => {
      const c = rgb(hex);
      if (!c) return 0;
      const k = hex.toUpperCase();
      if (!colorIdx.has(k)) { colores.push(c); colorIdx.set(k, colores.length); }
      return colorIdx.get(k);
    };
    const pd = contrato.predeterminado || {};
    ftc((pd.r && pd.r.fuente) || 'Times New Roman');

    // ───────── Propiedades ─────────
    function propsR(r) {
      if (!r) return '';
      let x = '';
      if (r.fuente) x += `\\f${ftc(r.fuente)}`;
      if (r.tamano != null) x += `\\fs${entero(r.tamano * 2, 2, 3276, 22)}`;
      if (r.negrita != null) x += r.negrita ? '\\b' : '\\b0';
      if (r.cursiva != null) x += r.cursiva ? '\\i' : '\\i0';
      if (r.subrayado != null) x += r.subrayado === 'doble' ? '\\uldb' : r.subrayado ? '\\ul' : '\\ulnone';
      if (r.tachado != null) x += r.tachado ? '\\strike' : '\\strike0';
      if (r.versalitas != null) x += r.versalitas ? '\\scaps' : '\\scaps0';
      if (r.mayusculas != null) x += r.mayusculas ? '\\caps' : '\\caps0';
      if (r.oculto) x += '\\v';
      if (r.sup) x += '\\super'; else if (r.sub) x += '\\sub';
      const c = color(r.color); if (c) x += `\\cf${c}`;
      if (r.resaltado && RESALTADO[r.resaltado]) x += `\\highlight${color(RESALTADO[r.resaltado])}`;
      return x;
    }
    const brd = (l) => `${BORDE[l.estilo || 'simple'] || '\\brdrs'}\\brdrw${entero((l.tam == null ? 4 : l.tam) * 2.5, 5, 240, 10)}${color(l.color) ? `\\brdrcf${color(l.color)}` : ''}`;

    const idEstilo = new Map();      // id del contrato → número de estilo
    let siguienteEstilo = 1;
    const estilos = (contrato.estilos || []).filter((e) => e && e.tipo !== 'tabla').slice(0, 4000);
    for (const e of estilos) { if (e.id === 'Normal') idEstilo.set(e.id, 0); else idEstilo.set(e.id, siguienteEstilo++); }
    if (!idEstilo.has('Normal')) idEstilo.set('Normal', 0);

    const listas = (contrato.listas || []).slice(0, 2000);
    const numLista = new Map(listas.map((l, i) => [l.id, i + 1]));

    function propsP(p, nivelLista) {
      let x = '';
      if (!p) return x;
      if (p.alineacion && ALINEACION[p.alineacion]) x += ALINEACION[p.alineacion];
      const izq = p.izq != null ? p.izq : (nivelLista ? nivelLista.izquierda : null), primera = p.primera != null ? p.primera : (nivelLista ? -nivelLista.colgante : null);
      if (izq != null) x += `\\li${entero(izq, -31680, 31680)}`;
      if (p.der != null) x += `\\ri${entero(p.der, -31680, 31680)}`;
      if (primera != null) x += `\\fi${entero(primera, -31680, 31680)}`;
      if (p.antes != null) x += `\\sb${entero(p.antes, 0, 31680)}`;
      if (p.despues != null) x += `\\sa${entero(p.despues, 0, 31680)}`;
      if (p.linea != null) x += `\\sl${p.regla === 'exacto' ? -entero(p.linea, 1, 31680, 240) : entero(p.linea, 1, 31680, 240)}\\slmult${p.regla === 'auto' || !p.regla ? 1 : 0}`;
      if (p.mantener) x += '\\keepn';
      if (p.mantenerJuntas) x += '\\keep';
      if (p.saltoAntes) x += '\\pagebb';
      if (p.nivelEsquema != null && p.nivelEsquema >= 0 && p.nivelEsquema <= 8) x += `\\outlinelevel${p.nivelEsquema}`;
      if (p.bordes) for (const [k, t] of [['arriba', 't'], ['izquierda', 'l'], ['abajo', 'b'], ['derecha', 'r']]) if (p.bordes[k]) x += `\\brdr${t}${brd(p.bordes[k])}\\brsp${entero(p.bordes[k].espacio || 1, 0, 31) * 20}`;
      if (p.relleno && color(p.relleno)) x += `\\cbpat${color(p.relleno)}`;
      if (p.tabs) for (const t of p.tabs.slice(0, 64)) x += `${TAB_TIPO[t.tipo] || ''}${TAB_RELLENO[t.relleno] || ''}\\tx${entero(t.pos, 0, 31680)}`;
      return x;
    }

    // Word no aplica las propiedades de un estilo a un párrafo por el solo hecho de nombrarlo en RTF: las escribe también en el párrafo, ya resueltas con lo que hereda
    const porId = new Map(estilos.map((e) => [e.id, e]));
    const normalE = porId.get('Normal');
    const baseNormal = { r: Object.assign({ tamano: 10 }, pd.r, normalE && normalE.r), p: Object.assign({}, pd.p, normalE && normalE.p) };
    const cacheEstilo = new Map();
    function resuelto(id, soloR) {
      const clave = (soloR ? 'r|' : 'p|') + id;
      if (cacheEstilo.has(clave)) return cacheEstilo.get(clave);
      const cadena = [];
      for (let e = porId.get(id), k = 0; e && k < 32; k++) { if (e.id === 'Normal') break; cadena.unshift(e); e = e.basadoEn != null ? porId.get(e.basadoEn) : null; }
      // los interruptores van siempre explícitos: si no, el estilo integrado del mismo nombre que trae Word (un título en negrita) manda sobre lo que se quiso
      const r = soloR ? {} : Object.assign({ negrita: false, cursiva: false, subrayado: false, tachado: false }, baseNormal.r), pp = soloR ? {} : Object.assign({}, baseNormal.p);
      for (const e of cadena) { Object.assign(r, e.r); if (e.tipo !== 'caracter') Object.assign(pp, e.p); }
      delete pp.lista;
      const res = { r, p: pp };
      cacheEstilo.set(clave, res);
      return res;
    }

    // ───────── Texto ─────────
    let imagenes = 0;
    // el formato de un tramo: su estilo de carácter (con lo que trae resuelto) y lo que lleva encima
    const fmDe = (r) => (r.estilo && idEstilo.has(r.estilo) ? `\\cs${idEstilo.get(r.estilo)}${propsR(Object.assign({}, resuelto(r.estilo, true).r, r.r))}` : propsR(r.r));
    function runs(lista, prof) {
      let x = '';
      for (const r of lista || []) {
        if (r.salto) { x += r.salto === 'pagina' ? '\\page ' : r.salto === 'columna' ? '\\column ' : '\\line '; continue; }
        if (r.marcador != null) { const n = nombreMarca(r.marcador); x += `{\\*\\bkmkstart ${n}}{\\*\\bkmkend ${n}}`; continue; }
        if (r.imagen) { x += imagen(r.imagen); continue; }
        if (r.enlace) {
          if (prof > 3) continue;
          const interior = runs(r.hijos, prof + 1);
          let instr = null;
          if (r.enlace.ancla) instr = `HYPERLINK \\\\l "${nombreMarca(r.enlace.ancla)}"`;
          else if (/^(https?:|mailto:|tel:|ftp:)/i.test(r.enlace.url || '')) instr = `HYPERLINK "${textoRtf(String(r.enlace.url).replace(/[\\"{}]/g, ''))}"`;
          if (!instr) { aviso('enlace-descartado'); x += interior; continue; }
          x += `{\\field{\\*\\fldinst ${instr}}{\\fldrslt ${interior}}}`;
          continue;
        }
        if (r.campo) {
          const nombre = String(r.campo).trim().split(/\s+/)[0].toUpperCase();
          const fm = fmDe(r);
          if (!/^[A-Z]+$/.test(nombre)) { x += `{${fm} ${textoRtf(String(r.resultado == null ? '' : r.resultado))}}`; continue; }
          x += `{\\field{\\*\\fldinst ${nombre}}{\\fldrslt {${fm ? fm + ' ' : ''}${textoRtf(String(r.resultado == null ? '1' : r.resultado))}}}}`;
          continue;
        }
        const t = textoRtf(String(r.texto == null ? '' : r.texto).replace(/\r\n?/g, '\n'));
        if (!t) continue;
        const fm = fmDe(r);
        x += fm ? `{${fm} ${t}}` : t;
      }
      return x;
    }

    function imagen(img) {
      const bytes = img.datos instanceof Uint8Array ? img.datos : new Uint8Array(img.datos || 0);
      const tipo = tipoDeImagen(bytes);
      if (!tipo) { aviso('imagen-formato'); return ''; }
      imagenes++;
      const wTw = entero((Number(img.ancho) || 72) * 20, 20, 31680, 1440), hTw = entero((Number(img.alto) || 72) * 20, 20, 31680, 1440);
      return `{\\pict\\${tipo === 'png' ? 'pngblip' : 'jpegblip'}\\picw${Math.max(1, Math.round(wTw / 15))}\\pich${Math.max(1, Math.round(hTw / 15))}\\picwgoal${wTw}\\pichgoal${hTw}\n${hexa(bytes)}}`;
    }

    function parrafo(b, ctx) {
      const e = b.estilo != null && idEstilo.has(b.estilo) ? idEstilo.get(b.estilo) : 0;
      let nivel = null, lista = '';
      if (b.p && b.p.lista && b.p.lista.id != null && numLista.has(b.p.lista.id)) {
        const l = listas[numLista.get(b.p.lista.id) - 1];
        const n = entero(b.p.lista.nivel, 0, 8);
        nivel = (l.niveles || [])[n] || null;
        lista = `\\ls${numLista.get(b.p.lista.id)}\\ilvl${n}`;
      } else if (b.p && b.p.lista && b.p.lista.id == null) lista = '\\ls0';
      const rs = resuelto(b.estilo != null && porId.has(b.estilo) ? b.estilo : 'Normal');
      const p = Object.assign({}, rs.p, b.p);
      delete p.lista;
      const tabla = ctx && ctx.itap ? `\\intbl${ctx.itap > 1 ? `\\itap${ctx.itap}` : ''}` : '';
      return `\\pard${tabla}\\s${e}${lista}${propsP(p, nivel)}${propsR(rs.r)} ${runs(b.hijos, 0)}`;
    }

    // ───────── Tablas ─────────
    function filaDef(t, f, anchos, izq, bordeTabla, etiqueta) {
      let x = `\\trowd\\trgaph108\\trleft${izq}`;
      if (f.encabezado) x += '\\trhdr';
      if (f.alto) x += `\\trrh${entero(f.alto, 1, 31680)}`;
      let pos = izq, col = 0;
      for (const c of f.celdas) {
        const span = Math.max(1, Math.min(63, c.fusionH || 1));
        for (let k = col; k < col + span; k++) pos += k < anchos.length ? anchos[k] : 1440;
        col += span;
        if (c.fusionV === 'inicio') x += '\\clvmgf'; else if (c.fusionV === 'continua') x += '\\clvmrg';
        if (c.alineacionV === 'centro') x += '\\clvertalc'; else if (c.alineacionV === 'abajo') x += '\\clvertalb';
        const lados = { arriba: bordeTabla, abajo: bordeTabla, izquierda: bordeTabla, derecha: bordeTabla };
        for (const k of Object.keys(lados)) if (c.bordes && c.bordes[k]) lados[k] = c.bordes[k];
        for (const [k, t2] of [['arriba', 't'], ['izquierda', 'l'], ['abajo', 'b'], ['derecha', 'r']]) if (lados[k]) x += `\\clbrdr${t2}${brd(lados[k])}`;
        if (c.relleno && color(c.relleno)) x += `\\clcbpat${color(c.relleno)}`;
        x += `\\cellx${entero(pos, -32768, 32767)}`;
      }
      void t; void etiqueta;
      return x;
    }
    function tabla(t, itap) {
      const filas = (t.filas || []).filter((f) => f.celdas && f.celdas.length);
      if (!filas.length || itap > MAX_PROFUNDIDAD_TABLAS) return '';
      const nCols = filas.reduce((m, f) => Math.max(m, f.celdas.reduce((s, c) => s + Math.max(1, c.fusionH || 1), 0)), Math.max(t.anchos ? t.anchos.length : 0, 1));
      const anchos = Array.from({ length: Math.min(nCols, 63) }, (_, i) => entero(t.anchos && t.anchos[i], 0, 31680, 1440));
      const izq = entero(t.sangria || 0, -31680, 31680);
      const bordeTabla = t.bordes === false ? null : { tam: (t.bordes && t.bordes.tam) || 4, color: t.bordes && t.bordes.color };
      let x = '';
      for (const f of filas) {
        const def = filaDef(t, f, anchos, izq, bordeTabla);
        let celdas = '';
        for (const c of f.celdas) {
          const hijos = c.hijos && c.hijos.length ? c.hijos : [{ tipo: 'parrafo', hijos: [] }];
          const ultimoEsParrafo = hijos[hijos.length - 1].tipo !== 'tabla';
          let contenido = '';
          hijos.forEach((h, k) => {
            if (h.tipo === 'tabla') { contenido += tabla(h, itap + 1); return; }
            const esUltimo = k === hijos.length - 1;
            contenido += parrafo(h, { itap }) + (esUltimo ? (itap > 1 ? '\\nestcell ' : '\\cell ') : '\\par\n');
          });
          if (!ultimoEsParrafo) contenido += `\\pard\\plain\\intbl${itap > 1 ? `\\itap${itap}` : ''} ${itap > 1 ? '\\nestcell ' : '\\cell '}`;
          celdas += contenido;
        }
        if (itap > 1) x += `${celdas}{\\*\\nesttableprops${def}\\nestrow}{\\nonesttables\\par}\n`;
        else x += `${def}\n${celdas}\\row\n`;
      }
      return x;
    }
    // sinFinal: el último párrafo no lleva par porque lo termina el sect de la sección (así Word no suma un párrafo vacío)
    function bloques(lista, sinFinal) {
      let x = '';
      let ultimoEsTabla = false;
      const n = (lista || []).length;
      (lista || []).forEach((b, k) => {
        if (b.tipo === 'tabla') { x += tabla(b, 1); ultimoEsTabla = true; }
        else { x += parrafo(b, null) + (sinFinal && k === n - 1 ? '\n' : '\\par\n'); ultimoEsTabla = false; }
      });
      if (ultimoEsTabla) x += '\\pard\\par\n';
      return x;
    }

    // ───────── Secciones ─────────
    let secciones = contrato.secciones && contrato.secciones.length ? contrato.secciones : [{ hijos: [] }];
    // Más de MAX_SECCIONES (solo un archivo raro): el contenido de las de más pasa a la última, para no perder texto
    if (secciones.length > MAX_SECCIONES) secciones = secciones.slice(0, MAX_SECCIONES - 1).concat([Object.assign({}, secciones[MAX_SECCIONES - 1], { hijos: secciones.slice(MAX_SECCIONES - 1).flatMap((x) => x.hijos || []) })]);
    let cuerpo = '';
    secciones.forEach((s, i) => {
      const pg = s.pagina || {}, m = s.margenes || {};
      cuerpo += i > 0 ? '\\sect' : '';
      cuerpo += `\\sectd${i > 0 && s.tipo === 'continua' ? '\\sbknone' : '\\sbkpage'}\\pgwsxn${entero(pg.ancho || 12240, 720, 31680)}\\pghsxn${entero(pg.alto || 15840, 720, 31680)}${pg.horizontal ? '\\lndscpsxn' : ''}`
        + `\\marglsxn${entero(m.izquierda == null ? 1440 : m.izquierda, 0, 31680)}\\margrsxn${entero(m.derecha == null ? 1440 : m.derecha, 0, 31680)}\\margtsxn${entero(m.arriba == null ? 1440 : m.arriba, 0, 31680)}\\margbsxn${entero(m.abajo == null ? 1440 : m.abajo, 0, 31680)}`
        + `\\headery${entero(m.encabezado == null ? 708 : m.encabezado, 0, 31680)}\\footery${entero(m.pie == null ? 708 : m.pie, 0, 31680)}`;
      if (s.columnas && s.columnas.n > 1) cuerpo += `\\cols${entero(s.columnas.n, 2, 45, 2)}\\colsx${entero(s.columnas.espacio == null ? 720 : s.columnas.espacio, 0, 31680)}`;
      cuerpo += '\n';
      if (s.encabezado && s.encabezado.length) cuerpo += `{\\header ${bloques(s.encabezado)}}\n`;
      if (s.pie && s.pie.length) cuerpo += `{\\footer ${bloques(s.pie)}}\n`;
      const hijos = s.hijos && s.hijos.length ? s.hijos : [{ tipo: 'parrafo', hijos: [] }];
      cuerpo += bloques(hijos, i < secciones.length - 1);
    });

    // ───────── Cabecera: fuentes, colores, estilos, listas ─────────
    let estilosRtf = '';
    const normal = estilos.find((e) => e.id === 'Normal');
    const rNormal = Object.assign({}, pd.r, normal && normal.r), pNormal = Object.assign({}, pd.p, normal && normal.p);
    if (rNormal.tamano == null) rNormal.tamano = 10;
    delete pNormal.lista;
    estilosRtf += `{\\s0\\snext0${propsP(pNormal)}${propsR(rNormal)} Normal;}`;
    for (const e of estilos) {
      if (e.id === 'Normal') continue;
      const n = idEstilo.get(e.id);
      const base = e.basadoEn != null && idEstilo.has(e.basadoEn) ? idEstilo.get(e.basadoEn) : 0;
      const sig = e.siguiente != null && idEstilo.has(e.siguiente) ? idEstilo.get(e.siguiente) : n;
      const nombre = textoRtf(String(e.nombre || e.id));
      if (e.tipo === 'caracter') estilosRtf += `{\\*\\cs${n}\\additive\\sbasedon${base}${propsR(e.r)} ${nombre};}`;
      else { const p = e.p ? Object.assign({}, e.p) : null; if (p) delete p.lista; estilosRtf += `{\\s${n}\\sbasedon${base}\\snext${sig}${propsP(p)}${propsR(e.r)} ${nombre};}`; }
    }

    let listasRtf = '', sobreRtf = '';
    listas.forEach((l, i) => {
      let niveles = '';
      for (let n = 0; n < 9; n++) {
        const v = (l.niveles || [])[n] || { formato: 'decimal', texto: `%${n + 1}.`, izquierda: 720 * (n + 1), colgante: 360 };
        const nfc = NFC[v.formato] != null ? NFC[v.formato] : 0;
        let texto = String(v.texto == null ? '' : v.texto);
        let leveltext = '', numeros = '';
        if (nfc === 23) { const c = VINETAS[texto.slice(0, 1)] || texto.slice(0, 1) || '\u2022'; leveltext = `\\'01${textoRtf(c.replace(/[\uE000-\uF8FF]/g, '\u2022'))}`; }
        else {
          const cuerpoTxt = [];
          const pos = [];
          for (let k = 0; k < texto.length && cuerpoTxt.length < 250; k++) {
            if (texto[k] === '%' && /[1-9]/.test(texto[k + 1] || '')) { pos.push(cuerpoTxt.length + 1); cuerpoTxt.push(`\\'${(+texto[k + 1] - 1).toString(16).padStart(2, '0')}`); k++; }
            else cuerpoTxt.push(textoRtf(texto[k]));
          }
          leveltext = `\\'${cuerpoTxt.length.toString(16).padStart(2, '0')}${cuerpoTxt.join('')}`;
          numeros = pos.map((p) => `\\'${p.toString(16).padStart(2, '0')}`).join('');
        }
        niveles += `{\\listlevel\\levelnfc${nfc}\\leveljc0\\levelfollow0\\levelstartat${entero(v.inicio == null ? 1 : v.inicio, 0, 32767, 1)}\\levelspace0\\levelindent0`
          + `{\\leveltext${leveltext};}{\\levelnumbers${numeros};}${v.fuente ? `\\f${ftc(v.fuente)}` : ''}\\fi-${entero(v.colgante, 0, 31680)}\\li${entero(v.izquierda, 0, 31680)}}`;
      }
      listasRtf += `{\\list\\listtemplateid${1000 + i}\\listhybrid${niveles}\\listid${1000 + i}}`;
      sobreRtf += `{\\listoverride\\listid${1000 + i}\\listoverridecount0\\ls${i + 1}}`;
    });

    const pr = contrato.propiedades || {};
    const cab = `{\\rtf1\\ansi\\ansicpg1252\\deff0\\nouicompat\\deflang1033`
      + `{\\fonttbl${fuentes.map((f, i) => `{\\f${i}\\fnil\\fcharset0 ${textoRtf(f)};}`).join('')}}`
      + `{\\colortbl ;${colores.map((c) => `\\red${c[0]}\\green${c[1]}\\blue${c[2]};`).join('')}}`
      + `{\\stylesheet${estilosRtf}}`
      + (listas.length ? `{\\*\\listtable${listasRtf}}{\\*\\listoverridetable${sobreRtf}}` : '')
      + `{\\info${pr.titulo ? `{\\title ${textoRtf(pr.titulo)}}` : ''}${pr.autor ? `{\\author ${textoRtf(pr.autor)}}` : ''}}\n`;
    // la cabecera se arma al final porque las fuentes y los colores se descubren al recorrer el contenido
    const texto = cab.replace(/^/, '') + cuerpo + '}';
    return { bytes: new TextEncoder().encode(texto), mensajes };
  }

  globalThis.RtfEscritor = { generar };
})();

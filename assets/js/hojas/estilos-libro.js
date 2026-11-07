/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Estilos nuevos de celda para Mini Cookies Hojas (fase 2 del PLAN). Poner negrita, un color o un formato de número a una celda es darle
// otro índice de cellXfs: este módulo arma ese xf a partir del que la celda ya tenía, lo agrega al final de styles.xml (con su fuente,
// relleno, borde o formato de número nuevos si hacen falta) y deja todo lo demás del archivo como estaba. Los xf que se crean no se borran
// al deshacer: un xf sin uso es inofensivo y Excel lo acepta. Sin DOM.
//
// Requiere xml-editar.js, xlsx-estilos.js y paquete-editable.js.
//
//   const est = await EstilosLibro.crear(libro, paquete)   // libro: el de LibroXlsx.abrir; paquete: PaqueteEditable.desdeZip(libro.zip)
//   const i = est.derivar(indiceBase, cambios)            // → índice de cellXfs (el mismo si no cambia nada); deja libro.estilos al día
//   cambios: { negrita, cursiva, subrayado (true|'double'|false), tachado, colorFuente: 'RRGGBB'|null, tamano, fuente,
//              relleno: 'RRGGBB'|null, h: 'left'|'center'|'right'|'general', v: 'top'|'center'|'bottom', ajustar: bool,
//              formato: código de formato de número ('0.00', '0%', …), bordes: 'todos'|'ninguno'|'exterior' (exterior lo resuelve quien llama por lado),
//              borde: { izq, der, arr, aba: 'thin'|'medium'|'thick'|null } }
//   est.propiedades(indice) → { negrita, cursiva, subrayado, tachado, colorFuente, tamano, fuente, relleno, h, v, ajustar, formato } (para la cinta)
(function () {
  const X = () => globalThis.XmlEditar;
  const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const TIPO_ESTILOS = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
  const REL_ESTILOS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
  const ESTILOS_MINIMOS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="' + NS_MAIN + '"><fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  const ORDEN_STYLESHEET = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'];
  // Orden de los hijos de <font> en CT_Font (Excel rechaza otro orden)
  const ORDEN_FUENTE = ['b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow', 'u', 'vertAlign', 'sz', 'color', 'name', 'family', 'charset', 'scheme'];
  const INTEGRADOS = { 'General': 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10, '0.00E+00': 11, '# ?/?': 12, '# ??/??': 13, 'mm-dd-yy': 14, 'd-mmm-yy': 15, 'd-mmm': 16, 'mmm-yy': 17, 'h:mm AM/PM': 18, 'h:mm:ss AM/PM': 19, 'h:mm': 20, 'h:mm:ss': 21, 'm/d/yy h:mm': 22, '@': 49 };

  async function crear(libro, paquete) {
    let rel = (libro.relaciones || []).find((r) => (r.tipo === 'styles' || /\/styles$/.test(r.tipo)) && !r.externa);
    let ruta = rel ? rel.destino : null;
    if (!ruta || !(await paquete.leer(ruta))) {   // libro sin styles.xml: se crea uno mínimo (Excel lo crea igual al guardar)
      ruta = 'xl/styles.xml';
      paquete.escribir(ruta, ESTILOS_MINIMOS);
      await paquete.asegurarTipo('/' + ruta, TIPO_ESTILOS);
      if (!rel) await paquete.agregarRelacion(libro.rutaLibro, REL_ESTILOS, 'styles.xml');
    }
    const tema = libro.rutaTema ? await paquete.leer(libro.rutaTema) : null;
    let texto = await paquete.leer(ruta);
    const memo = new Map();

    // la sección (fonts, fills…) y sus hijos, analizada de nuevo cuando el texto cambia
    function seccion(nombre) {
      const raiz = X().analizar(texto, { profundidad: 3 });
      return { raiz, nodo: X().hijo(raiz, nombre) };
    }
    function hijosDe(nombre, hijo) {
      const { nodo } = seccion(nombre);
      return nodo ? X().hijos(nodo, hijo).map((h) => texto.slice(h.ini, h.fin)) : [];
    }
    // agrega un elemento al final de la sección (la crea en su lugar si falta) y actualiza count; devuelve su índice
    function agregar(nombre, xml) {
      let { raiz, nodo } = seccion(nombre);
      if (!nodo) {
        const k = ORDEN_STYLESHEET.indexOf(nombre);
        const siguiente = raiz.hijos.find((h) => ORDEN_STYLESHEET.indexOf(h.local) > k);
        const pos = siguiente ? siguiente.ini : raiz.iniCierre;
        texto = X().aplicar(texto, [{ ini: pos, fin: pos, texto: `<${nombre} count="0"></${nombre}>` }]);
        ({ raiz, nodo } = seccion(nombre));
      }
      const n = nodo.hijos.length;
      const apertura = X().ponerAtributo(texto.slice(nodo.ini, nodo.finApertura).replace(/\s*\/>$/, '>'), 'count', String(n + 1));
      const ed = nodo.iniCierre < 0
        ? [{ ini: nodo.ini, fin: nodo.fin, texto: apertura + xml + `</${nodo.nombre}>` }]
        : [{ ini: nodo.ini, fin: nodo.finApertura, texto: apertura }, { ini: nodo.iniCierre, fin: nodo.iniCierre, texto: xml }];
      texto = X().aplicar(texto, ed);
      return n;
    }
    function buscarOAgregar(nombre, hijo, xml) {
      const lista = hijosDe(nombre, hijo);
      const k = lista.indexOf(xml);
      return k >= 0 ? k : agregar(nombre, xml);
    }
    const atributo = (xml, nombre) => { const m = new RegExp('\\b' + nombre + '="([^"]*)"').exec(xml); return m ? m[1] : null; };

    // <font> base con los cambios pedidos, con los hijos en el orden de la norma
    function fuenteNueva(base, c) {
      const raiz = X().analizar(base || '<font/>', { profundidad: 2 });
      const hijos = new Map();
      for (const h of raiz.hijos) hijos.set(h.local, (base || '').slice(h.ini, h.fin));
      const poner = (k, v) => { if (v === null) hijos.delete(k); else hijos.set(k, v); };
      if ('negrita' in c) poner('b', c.negrita ? '<b/>' : null);
      if ('cursiva' in c) poner('i', c.cursiva ? '<i/>' : null);
      if ('tachado' in c) poner('strike', c.tachado ? '<strike/>' : null);
      if ('subrayado' in c) poner('u', c.subrayado ? (c.subrayado === 'double' ? '<u val="double"/>' : '<u/>') : null);
      if ('colorFuente' in c) poner('color', c.colorFuente ? `<color rgb="FF${c.colorFuente.toUpperCase()}"/>` : null);
      if ('tamano' in c && c.tamano) poner('sz', `<sz val="${+c.tamano}"/>`);
      if ('fuente' in c && c.fuente) { poner('name', `<name val="${X().escAttr(c.fuente)}"/>`); hijos.delete('scheme'); hijos.delete('family'); hijos.delete('charset'); }
      const orden = [...hijos.keys()].sort((a, b) => (ORDEN_FUENTE.indexOf(a) < 0 ? 99 : ORDEN_FUENTE.indexOf(a)) - (ORDEN_FUENTE.indexOf(b) < 0 ? 99 : ORDEN_FUENTE.indexOf(b)));
      return '<font>' + orden.map((k) => hijos.get(k)).join('') + '</font>';
    }
    function bordeNuevo(base, c) {
      const lados = { left: null, right: null, top: null, bottom: null, diagonal: null };
      const raiz = X().analizar(base || '<border/>', { profundidad: 2 });
      for (const h of raiz.hijos) if (h.local in lados) lados[h.local] = (base || '').slice(h.ini, h.fin);
      const lado = (e) => { const estilo = e && e.estilo ? e.estilo : e, col = e && e.color; return estilo ? `<${'$'} style="${estilo}">${col ? `<color rgb="FF${col}"/>` : '<color indexed="64"/>'}</${'$'}>` : '<$/>'; };
      const poner = (k, estilo) => { lados[k] = lado(estilo).replace(/\$/g, k); };
      if (c.bordes === 'todos') ['left', 'right', 'top', 'bottom'].forEach((k) => poner(k, 'thin'));
      if (c.bordes === 'ninguno') ['left', 'right', 'top', 'bottom'].forEach((k) => poner(k, null));
      if (c.borde) { const m = { izq: 'left', der: 'right', arr: 'top', aba: 'bottom' }; for (const k in c.borde) if (m[k]) poner(m[k], c.borde[k]); }
      const aperturaBase = base ? base.slice(0, X().analizar(base, { profundidad: 1 }).finApertura).replace(/\s*\/>$/, '>') : '<border>';
      return aperturaBase + ['left', 'right', 'top', 'bottom', 'diagonal'].map((k) => lados[k] || `<${k}/>`).join('') + '</border>';
    }
    function idFormato(codigo) {
      if (codigo in INTEGRADOS) return INTEGRADOS[codigo];
      const lista = hijosDe('numFmts', 'numFmt');
      let mayor = 163;
      for (const x of lista) {
        const id = +atributo(x, 'numFmtId'), cod = X().decodificar(atributo(x, 'formatCode') || '');
        if (cod === codigo) return id;
        if (id > mayor) mayor = id;
      }
      agregar('numFmts', `<numFmt numFmtId="${mayor + 1}" formatCode="${X().escAttr(codigo)}"/>`);
      return mayor + 1;
    }

    function derivar(indiceBase, cambios) {
      const k = indiceBase + '|' + JSON.stringify(cambios);
      if (memo.has(k)) return memo.get(k);
      const xfs = hijosDe('cellXfs', 'xf');
      let xf = xfs[indiceBase] || xfs[0] || '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
      const raizXf = X().analizar(xf, { profundidad: 2 });
      let apertura = xf.slice(0, raizXf.finApertura).replace(/\s*\/>$/, '>');
      const hijos = raizXf.hijos.map((h) => ({ local: h.local, xml: xf.slice(h.ini, h.fin) }));
      const quiereFuente = ['negrita', 'cursiva', 'subrayado', 'tachado', 'colorFuente', 'tamano', 'fuente'].some((p) => p in cambios);
      if (quiereFuente) {
        const fuentes = hijosDe('fonts', 'font');
        const id = buscarOAgregar('fonts', 'font', fuenteNueva(fuentes[+atributo(apertura, 'fontId') || 0] || fuentes[0], cambios));
        apertura = X().ponerAtributo(X().ponerAtributo(apertura, 'fontId', String(id)), 'applyFont', '1');
      }
      if ('relleno' in cambios) {
        const xml = cambios.relleno ? `<fill><patternFill patternType="solid"><fgColor rgb="FF${cambios.relleno.toUpperCase()}"/><bgColor indexed="64"/></patternFill></fill>` : '<fill><patternFill patternType="none"/></fill>';
        const id = cambios.relleno ? buscarOAgregar('fills', 'fill', xml) : 0;
        apertura = X().ponerAtributo(X().ponerAtributo(apertura, 'fillId', String(id)), 'applyFill', '1');
      }
      if ('bordes' in cambios || 'borde' in cambios) {
        const bordes = hijosDe('borders', 'border');
        const nuevo = cambios.bordes === 'ninguno' && !cambios.borde ? null : bordeNuevo(bordes[+atributo(apertura, 'borderId') || 0], cambios);
        const id = nuevo === null ? 0 : buscarOAgregar('borders', 'border', nuevo);
        apertura = X().ponerAtributo(X().ponerAtributo(apertura, 'borderId', String(id)), 'applyBorder', '1');
      }
      if ('bordeDe' in cambios) { const b = atributo(xfs[cambios.bordeDe] || xfs[0], 'borderId') || '0'; apertura = X().ponerAtributo(X().ponerAtributo(apertura, 'borderId', b), 'applyBorder', '1'); }   // los bordes de otro xf
      if ('formato' in cambios) apertura = X().ponerAtributo(X().ponerAtributo(apertura, 'numFmtId', String(idFormato(cambios.formato || 'General'))), 'applyNumberFormat', '1');
      if ('h' in cambios || 'v' in cambios || 'ajustar' in cambios) {
        let al = hijos.find((h) => h.local === 'alignment');
        if (!al) { al = { local: 'alignment', xml: '<alignment/>' }; hijos.unshift(al); }
        let a = al.xml.replace(/\s*\/>$/, '/>').replace(/<alignment([^>]*)>[\s\S]*$/, '<alignment$1/>').replace(/\/\/>$/, '/>');
        if ('h' in cambios) a = X().ponerAtributo(a, 'horizontal', cambios.h && cambios.h !== 'general' ? cambios.h : null);
        if ('v' in cambios) a = X().ponerAtributo(a, 'vertical', cambios.v && cambios.v !== 'bottom' ? cambios.v : null);
        if ('ajustar' in cambios) a = X().ponerAtributo(a, 'wrapText', cambios.ajustar ? '1' : null);
        al.xml = a;
        apertura = X().ponerAtributo(apertura, 'applyAlignment', '1');
      }
      const nuevo = apertura + hijos.filter((h) => h.local !== 'alignment' || /\s\w+="/.test(h.xml)).map((h) => h.xml).join('') + '</xf>';
      let indice = xfs.indexOf(nuevo);
      if (indice < 0) indice = agregar('cellXfs', nuevo);
      paquete.escribir(ruta, texto);
      libro.estilos = XlsxEstilos.leer(texto, tema);
      memo.set(k, indice);
      return indice;
    }

    function propiedades(indice) {
      const x = libro.estilos.xf(indice);
      const color = (ref) => { const c = ref ? libro.estilos.color(ref) : null; return c ? c.slice(1).toUpperCase() : null; };
      const f = x.fuente || {};
      return {
        negrita: !!f.negrita, cursiva: !!f.cursiva, subrayado: f.subrayado || false, tachado: !!f.tachado, colorFuente: color(f.color), tamano: f.tam || 11, fuente: f.nombre || 'Calibri',
        relleno: x.relleno && x.relleno.tipo === 'patron' && x.relleno.patron === 'solid' ? color(x.relleno.fg) : null,
        h: (x.alineacion && x.alineacion.h) || 'general', v: (x.alineacion && x.alineacion.v) || 'bottom', ajustar: !!(x.alineacion && x.alineacion.ajustar), formato: x.formato || 'General',
      };
    }

    return { derivar, propiedades, get ruta() { return ruta; } };
  }

  // Aumentar o disminuir decimales de un código de formato de número, sección por sección, como los botones de Excel:
  // "#,##0.00" +1 → "#,##0.000"; "0%" +1 → "0.0%"; "0.0E+00" -1 → "0E+00"; "$#,##0.00;[Red]-$#,##0.00" cambia las dos secciones.
  // No se tocan los textos entre comillas, los escapes (\x, _x, *x), los corchetes ([Red], [$€-es]), las fechas, las fracciones ni "@".
  function cambiarDecimales(codigo, delta) {
    if (!codigo || /^general$/i.test(codigo)) return delta > 0 ? '0.0' : '0';
    const secciones = [];
    let actual = '';
    for (let i = 0; i < codigo.length; i++) {   // separa por ";" fuera de comillas y escapes
      const ch = codigo[i];
      if (ch === '"') { const k = codigo.indexOf('"', i + 1); const fin = k < 0 ? codigo.length : k + 1; actual += codigo.slice(i, fin); i = fin - 1; }
      else if (ch === '\\' || ch === '_' || ch === '*') { actual += codigo.slice(i, i + 2); i++; }
      else if (ch === ';') { secciones.push(actual); actual = ''; }
      else actual += ch;
    }
    secciones.push(actual);
    return secciones.map((sec) => decimalesSeccion(sec, delta)).join(';');
  }
  function decimalesSeccion(sec, delta) {
    // marca qué posiciones son código (no literales)
    const codigoEn = new Array(sec.length).fill(false);
    for (let i = 0; i < sec.length; i++) {
      const ch = sec[i];
      if (ch === '"') { const k = sec.indexOf('"', i + 1); i = k < 0 ? sec.length : k; }
      else if (ch === '\\' || ch === '_' || ch === '*') i++;
      else if (ch === '[') { const k = sec.indexOf(']', i); i = k < 0 ? sec.length : k; }
      else codigoEn[i] = true;
    }
    const codigo = [...sec].map((ch, i) => (codigoEn[i] ? ch : ' ')).join('');
    if (/general/i.test(codigo)) return sec.replace(/general/i, delta > 0 ? '0.0' : '0');
    if (/[ymdhs@\/]/i.test(codigo.replace(/E[+-]/gi, ''))) return sec;   // fecha, hora, texto o fracción
    if (!/[0#]/.test(codigo) && /\?/.test(codigo)) {   // sección del cero de los formatos contables ("-"??): Excel suma o quita un "?"
      const ultimo = codigo.lastIndexOf('?');
      if (delta > 0) return sec.slice(0, ultimo + 1) + '?'.repeat(delta) + sec.slice(ultimo + 1);
      let quitar = -delta, i = ultimo;
      let r = sec;
      while (quitar > 0 && i >= 0 && codigo[i] === '?') { r = r.slice(0, i) + r.slice(i + 1); i--; quitar--; }
      return r;
    }
    const exp = codigo.search(/E[+-]/i), fin = exp < 0 ? codigo.length : exp;   // la mantisa (en científico)
    const punto = codigo.slice(0, fin).indexOf('.');
    if (punto >= 0) {
      let k = punto + 1;
      while (k < fin && /[0#?]/.test(codigo[k])) k++;
      const n = Math.max(0, k - punto - 1 + delta);
      return sec.slice(0, punto) + (n ? '.' + '0'.repeat(n) : '') + sec.slice(k);
    }
    if (delta <= 0) return sec;
    let ultimo = -1;
    for (let i = 0; i < fin; i++) if (/[0#?]/.test(codigo[i])) ultimo = i;
    if (ultimo < 0) return sec;
    return sec.slice(0, ultimo + 1) + '.' + '0'.repeat(delta) + sec.slice(ultimo + 1);
  }

  globalThis.EstilosLibro = { crear, cambiarDecimales };
})();

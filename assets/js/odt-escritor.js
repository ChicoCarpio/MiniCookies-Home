/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Escritor de .odt (OpenDocument de LibreOffice / OpenOffice) desde el «contrato de documento» que describe docx-desde-contrato.js. Es el gemelo de OdtLector: lo que ese lector
// entiende, este lo escribe, así que cualquier lector (DocxAContrato, DocLector, OdtLector) → .odt sin un conversor por pareja.
//
//   const { bytes, mensajes } = await OdtEscritor.generar(contrato);     // Uint8Array con el .odt
//
// Lleva: estilos con nombre (párrafo y carácter), valores predeterminados, listas (cada contrato de lista es un estilo de lista; si se interrumpe sigue numerando),
// secciones (una página maestra por hoja distinta con su encabezado y pie; las continuas con columnas son secciones de texto), tablas (anidadas, combinadas, bordes y relleno),
// imágenes en línea, enlaces, campos (número de página, total de páginas, fecha), marcadores, saltos de página y de columna, tabulaciones y bordes de párrafo.
// mensajes: 'imagen-formato' (una imagen que no es png, jpeg, gif ni bmp se omite), 'enlace-descartado' (un destino que no es una dirección web).
// Requiere zip-propio.js.
(function () {
  'use strict';
  const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
    + 'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" '
    + 'xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" '
    + 'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xml="http://www.w3.org/XML/1998/namespace" office:version="1.2"';
  const CABECERA = '<?xml version="1.0" encoding="UTF-8"?>';
  const INVALIDOS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;
  const esc = (s) => String(s).replace(INVALIDOS, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const escAttr = (s) => esc(s).replace(/"/g, '&quot;');
  const sinProto = (o) => Object.assign(Object.create(null), o);
  const hex6 = (c) => (/^[0-9a-fA-F]{6}$/.test(c || '') ? '#' + c.toUpperCase() : null);
  const num = (v, min, max, def = 0) => { const n = Number(v); return Number.isFinite(n) ? Math.max(min, Math.min(max, n)) : def; };
  const pulgadas = (tw) => `${(num(tw, -31680, 31680) / 1440).toFixed(4)}in`;
  const puntos = (pt) => `${num(pt, 0, 1584, 1).toFixed(2)}pt`;
  const MAX_SECCIONES_MAESTRAS = 2000;
  const MAX_PROFUNDIDAD_TABLAS = 8;

  const RESALTADO = sinProto({ yellow: 'FFFF00', green: '00FF00', cyan: '00FFFF', magenta: 'FF00FF', blue: '0000FF', red: 'FF0000', darkBlue: '000080', darkCyan: '008080', darkGreen: '008000', darkMagenta: '800080', darkRed: '800000', darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0', black: '000000', white: 'FFFFFF' });
  const ALINEACION = sinProto({ izquierda: 'start', centro: 'center', derecha: 'end', justificado: 'justify' });
  const TAB = sinProto({ izquierda: 'left', centro: 'center', derecha: 'right', decimal: 'char' });
  const RELLENO_TAB = sinProto({ punto: 'dotted', guion: 'dashed', linea: 'solid' });
  const BORDE = sinProto({ simple: 'solid', doble: 'double', punteado: 'dotted', guiones: 'dash' });
  const FORMATO_NUM = sinProto({ decimal: '1', lowerLetter: 'a', upperLetter: 'A', lowerRoman: 'i', upperRoman: 'I', ordinal: '1', decimalZero: '1', cardinalText: '1', ordinalText: '1', none: '' });
  const MIME = sinProto({ png: 'image/png', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp' });
  // Viñetas que Word guarda como un carácter de las fuentes Symbol / Wingdings (zona privada): se pasan al carácter Unicode que dibujan
  const VINETAS = sinProto({ '': '•', '': '▪', '': '➢', '': '✓', '': '❖', '': '☐', 'o': '○', '§': '▪' });

  function tipoDeImagen(b) {
    if (b.length > 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4E && b[3] === 0x47) return 'png';
    if (b.length > 3 && b[0] === 0xFF && b[1] === 0xD8 && b[2] === 0xFF) return 'jpeg';
    if (b.length > 6 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return 'gif';
    if (b.length > 2 && b[0] === 0x42 && b[1] === 0x4D) return 'bmp';
    return null;
  }

  async function generar(contrato) {
    const mensajes = [];
    const aviso = (m) => { if (!mensajes.includes(m)) mensajes.push(m); };
    const fuentes = new Set();
    const medios = [];
    const mediosPorDatos = new Map();
    const estilosAuto = [], estilosAutoCabeceras = [];   // XML de los estilos automáticos de content.xml y de styles.xml (encabezados y pies viven en styles.xml)
    let destinoAuto = { lista: estilosAuto, mapa: new Map() };
    const mapaCabeceras = new Map();
    let contador = 0;
    const nombreEstilo = (id) => (id === 'Normal' ? 'Standard' : String(id));
    const estilosPorId = new Map((contrato.estilos || []).map((e) => [e.id, e]));

    // Un estilo automático por combinación distinta de propiedades
    function automatico(prefijo, familia, cuerpo, extra = '') {
      const clave = `${prefijo}|${familia}|${extra}|${cuerpo}`;
      if (!destinoAuto.mapa.has(clave)) {
        const nombre = `${prefijo}${++contador}`;
        destinoAuto.mapa.set(clave, nombre);
        destinoAuto.lista.push(`<style:style style:name="${nombre}" style:family="${familia}"${extra}>${cuerpo}</style:style>`);
      }
      return destinoAuto.mapa.get(clave);
    }

    // ───────── Propiedades ─────────
    function textProps(r, aMenos) {
      if (!r) return '';
      let a = '';
      if (r.fuente) { fuentes.add(r.fuente); a += ` style:font-name="${escAttr(r.fuente)}"`; }
      if (r.fuenteAsia) { fuentes.add(r.fuenteAsia); a += ` style:font-name-asian="${escAttr(r.fuenteAsia)}"`; }
      if (r.fuenteCs) { fuentes.add(r.fuenteCs); a += ` style:font-name-complex="${escAttr(r.fuenteCs)}"`; }
      if (r.tamano != null) { const t = `${num(r.tamano, 1, 1638, 11)}pt`; a += ` fo:font-size="${t}" style:font-size-asian="${t}" style:font-size-complex="${t}"`; }
      if (r.negrita != null) { const v = r.negrita ? 'bold' : 'normal'; a += ` fo:font-weight="${v}" style:font-weight-asian="${v}" style:font-weight-complex="${v}"`; }
      if (r.cursiva != null) { const v = r.cursiva ? 'italic' : 'normal'; a += ` fo:font-style="${v}" style:font-style-asian="${v}" style:font-style-complex="${v}"`; }
      if (r.subrayado != null) a += r.subrayado ? ` style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"${r.subrayado === 'doble' ? ' style:text-underline-type="double"' : ''}` : ' style:text-underline-style="none"';
      if (r.tachado != null) a += r.tachado ? ' style:text-line-through-style="solid"' : ' style:text-line-through-style="none"';
      if (r.sup) a += ' style:text-position="super 58%"'; else if (r.sub) a += ' style:text-position="sub 58%"';
      if (r.versalitas != null) a += ` fo:font-variant="${r.versalitas ? 'small-caps' : 'normal'}"`;
      if (r.mayusculas != null) a += ` fo:text-transform="${r.mayusculas ? 'uppercase' : 'none'}"`;
      if (r.oculto) a += ' text:display="none"';
      const color = hex6(r.color); if (color) a += ` fo:color="${color}"`;
      if (r.resaltado && RESALTADO[r.resaltado]) a += ` fo:background-color="#${RESALTADO[r.resaltado]}"`;
      void aMenos;
      return a ? `<style:text-properties${a}/>` : '';
    }

    function borde(l) {
      const estilo = BORDE[l.estilo || 'simple'] || 'solid';
      return `${(num(l.tam == null ? 4 : l.tam, 1, 96, 4) / 8).toFixed(2)}pt ${estilo} ${hex6(l.color) || '#000000'}`;
    }
    function bordesAttr(b, ladosPad) {
      let a = '';
      if (!b) return a;
      for (const [clave, lado] of [['arriba', 'top'], ['abajo', 'bottom'], ['izquierda', 'left'], ['derecha', 'right']]) {
        if (!b[clave]) continue;
        a += ` fo:border-${lado}="${borde(b[clave])}"`;
        if (ladosPad) a += ` fo:padding-${lado}="${puntos(b[clave].espacio == null ? 1 : b[clave].espacio)}"`;
      }
      return a;
    }

    // Propiedades de párrafo → { attrs, hijos (tabulaciones) }
    function parProps(p, extra) {
      let a = '', tabs = '';
      if (p) {
        if (p.alineacion && ALINEACION[p.alineacion]) a += ` fo:text-align="${ALINEACION[p.alineacion]}"`;
        if (p.izq != null) a += ` fo:margin-left="${pulgadas(p.izq)}"`;
        if (p.der != null) a += ` fo:margin-right="${pulgadas(p.der)}"`;
        if (p.primera != null) a += ` fo:text-indent="${pulgadas(p.primera)}"`;
        if (p.antes != null) a += ` fo:margin-top="${pulgadas(Math.max(0, p.antes))}"`;
        if (p.despues != null) a += ` fo:margin-bottom="${pulgadas(Math.max(0, p.despues))}"`;
        if (p.linea != null) {
          if (p.regla === 'exacto') a += ` fo:line-height="${pulgadas(p.linea)}"`;
          else if (p.regla === 'minimo') a += ` style:line-height-at-least="${pulgadas(p.linea)}"`;
          else a += ` fo:line-height="${Math.round(num(p.linea, 1, 31680, 240) / 240 * 100)}%"`;
        }
        if (p.saltoAntes) a += ' fo:break-before="page"';
        if (p.mantener) a += ' fo:keep-with-next="always"';
        if (p.mantenerJuntas) a += ' fo:keep-together="always"';
        const fondo = hex6(p.relleno); if (fondo) a += ` fo:background-color="${fondo}"`;
        a += bordesAttr(p.bordes, true);
        if (p.tabs && p.tabs.length) {
          tabs = '<style:tab-stops>' + p.tabs.slice(0, 64).map((t) => `<style:tab-stop style:position="${pulgadas(t.pos)}" style:type="${TAB[t.tipo] || 'left'}"${RELLENO_TAB[t.relleno] ? ` style:leader-style="${RELLENO_TAB[t.relleno]}" style:leader-text="${t.relleno === 'punto' ? '.' : t.relleno === 'guion' ? '-' : '_'}"` : ''}${t.tipo === 'decimal' ? ' style:char="."' : ''}/>`).join('') + '</style:tab-stops>';
        }
      }
      a += extra || '';
      return a || tabs ? `<style:paragraph-properties${a}${tabs ? `>${tabs}</style:paragraph-properties>` : '/>'}` : '';
    }

    // ───────── Listas ─────────
    const listas = new Map();     // id del contrato → nombre del estilo de lista
    const listaXml = [];
    (contrato.listas || []).forEach((l, i) => {
      const nombre = `LS${i + 1}`;
      listas.set(l.id, nombre);
      let niveles = '';
      (l.niveles || []).slice(0, 9).forEach((v, n) => {
        const posicion = `<style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="${pulgadas(num(v.izquierda, 0, 31680))}" fo:text-indent="${pulgadas(-num(v.colgante, 0, 31680))}" fo:margin-left="${pulgadas(num(v.izquierda, 0, 31680))}"/></style:list-level-properties>`;
        const letra = v.r ? textProps(Object.assign({}, v.r, { fuente: undefined })) : '';
        if (v.formato === 'bullet') {
          const texto = String(v.texto == null ? '' : v.texto);
          const car = VINETAS[texto] || (texto && !/[-]/.test(texto) ? texto.slice(0, 2) : '•');
          const usaFuente = v.fuente && !/^(Symbol|Wingdings|Wingdings \d)$/i.test(v.fuente) && !VINETAS[texto];
          if (usaFuente) fuentes.add(v.fuente);
          niveles += `<text:list-level-style-bullet text:level="${n + 1}" text:bullet-char="${escAttr(car)}">${posicion}`
            + (usaFuente ? `<style:text-properties style:font-name="${escAttr(v.fuente)}"/>` : '') + '</text:list-level-style-bullet>';
        } else {
          const texto = String(v.texto == null ? '%' + (n + 1) + '.' : v.texto);
          const marcas = [...texto.matchAll(/%([1-9])/g)];
          const primero = marcas.length ? marcas[0].index : texto.length, ultimo = marcas.length ? marcas[marcas.length - 1].index + 2 : texto.length;
          const prefijo = texto.slice(0, primero), sufijo = texto.slice(ultimo);
          const mostrados = Math.max(1, new Set(marcas.map((m) => m[1])).size);
          const formato = FORMATO_NUM[v.formato] != null ? FORMATO_NUM[v.formato] : '1';
          niveles += `<text:list-level-style-number text:level="${n + 1}" style:num-format="${formato}"${prefijo ? ` style:num-prefix="${escAttr(prefijo)}"` : ''}${sufijo ? ` style:num-suffix="${escAttr(sufijo)}"` : ''}`
            + `${mostrados > 1 ? ` text:display-levels="${mostrados}"` : ''}${v.inicio != null && v.inicio !== 1 ? ` text:start-value="${num(v.inicio, 0, 32767, 1)}"` : ''}>${posicion}`
            + (letra ? letra : '') + '</text:list-level-style-number>';
        }
      });
      listaXml.push(`<text:list-style style:name="${nombre}">${niveles}</text:list-style>`);
    });

    // ───────── Texto ─────────
    function textoXml(t, estado) {
      let x = '';
      const s = String(t).replace(INVALIDOS, '').replace(/\r\n?/g, '\n');
      for (let i = 0; i < s.length;) {
        const c = s[i];
        if (c === ' ') {
          let j = i;
          while (s[j] === ' ') j++;
          const n = j - i;
          if (n === 1 && !estado.espacio) x += ' ';
          else if (estado.espacio) x += `<text:s text:c="${n}"/>`;
          else x += ` <text:s text:c="${n - 1}"/>`;
          estado.espacio = true;
          i = j;
          continue;
        }
        estado.espacio = false;
        if (c === '\t') x += '<text:tab/>';
        else if (c === '\n') { x += '<text:line-break/>'; estado.espacio = true; }
        else if (c === '&') x += '&amp;';
        else if (c === '<') x += '&lt;';
        else if (c === '>') x += '&gt;';
        else x += c;
        i++;
      }
      return x;
    }

    function dibujo(img, estado) {
      const bytes = img.datos instanceof Uint8Array ? img.datos : new Uint8Array(img.datos || 0);
      const tipo = tipoDeImagen(bytes);
      if (!tipo) { aviso('imagen-formato'); return ''; }
      let medio = mediosPorDatos.get(bytes);
      if (!medio) { medio = { nombre: `Pictures/image${medios.length + 1}.${tipo === 'jpeg' ? 'jpg' : tipo}`, bytes, tipo }; medios.push(medio); mediosPorDatos.set(bytes, medio); }
      estado.espacio = true;   // un espacio justo después de la imagen se escribe como <text:s/> para que no se pierda
      const id = ++contador;
      return `<draw:frame draw:name="Imagen${id}" text:anchor-type="as-char" svg:width="${puntos(img.ancho)}" svg:height="${puntos(img.alto)}"><draw:image xlink:href="${medio.nombre}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>`
        + `${img.texto ? `<svg:desc>${esc(img.texto)}</svg:desc>` : ''}</draw:frame>`;
    }

    function campo(c, r, estado) {
      const nombre = String(c.campo || '').trim().split(/\s+/)[0].toUpperCase();
      const resultado = c.resultado == null ? '1' : c.resultado;
      estado.espacio = false;
      if (nombre === 'PAGE') return `<text:page-number text:select-page="current">${esc(resultado)}</text:page-number>`;
      if (nombre === 'NUMPAGES') return `<text:page-count>${esc(resultado)}</text:page-count>`;
      return textoXml(resultado, estado);
    }

    // Un run con formato: un <text:span> con un estilo de carácter (el del contrato) y propiedades automáticas
    function conFormato(r, estiloCaracter, interior) {
      if (interior === '') return '';
      const props = textProps(r);
      if (!props && !estiloCaracter) return interior;
      const nombre = props ? automatico('T', 'text', props, estiloCaracter ? ` style:parent-style-name="${escAttr(nombreEstilo(estiloCaracter))}"` : '') : nombreEstilo(estiloCaracter);
      return `<text:span text:style-name="${escAttr(nombre)}">${interior}</text:span>`;
    }

    function runs(lista, estado, prof) {
      let x = '';
      for (const r of lista || []) {
        if (r.salto) { if (r.salto === 'linea') { x += '<text:line-break/>'; estado.espacio = true; } continue; }   // los saltos de página y de columna los trata parrafos()
        if (r.imagen) { x += dibujo(r.imagen, estado); continue; }
        if (r.marcador != null) { x += `<text:bookmark text:name="${escAttr(String(r.marcador).slice(0, 255))}"/>`; continue; }
        if (r.enlace) {
          if (prof > 3) continue;
          const interior = runs(r.hijos, estado, prof + 1);
          if (r.enlace.ancla) x += `<text:a xlink:type="simple" xlink:href="#${escAttr(r.enlace.ancla)}">${interior}</text:a>`;
          else if (/^(https?:|mailto:|tel:|ftp:)/i.test(r.enlace.url || '')) x += `<text:a xlink:type="simple" xlink:href="${escAttr(r.enlace.url)}">${interior}</text:a>`;
          else { aviso('enlace-descartado'); x += interior; }
          continue;
        }
        if (r.campo) { x += conFormato(r.r, r.estilo, campo(r, r.r, estado)); continue; }
        x += conFormato(r.r, r.estilo, textoXml(r.texto == null ? '' : r.texto, estado));
      }
      return x;
    }

    // Un salto de página o de columna dentro de un párrafo: al principio es «salto antes» del párrafo, al final «salto después», en medio parte el párrafo en dos
    function partirPorSaltos(hijos) {
      const trozos = [{ hijos: [], antes: null, despues: null }];
      for (const r of hijos || []) {
        if (r.salto === 'pagina' || r.salto === 'columna') {
          const t = trozos[trozos.length - 1];
          if (!t.hijos.some((x) => !x.marcador)) t.antes = r.salto;
          else { t.despues = r.salto; trozos.push({ hijos: [], antes: null, despues: null }); }
        } else trozos[trozos.length - 1].hijos.push(r);
      }
      // el trozo que queda vacío tras un salto final se une a «salto después» del anterior
      if (trozos.length > 1 && !trozos[trozos.length - 1].hijos.some((x) => !x.marcador) && !trozos[trozos.length - 1].antes && !trozos[trozos.length - 1].despues) {
        const ult = trozos.pop();
        trozos[trozos.length - 1].hijos.push(...ult.hijos);
      }
      return trozos;
    }

    function nivelEsquemaDe(b) {
      if (b.p && b.p.nivelEsquema != null) return b.p.nivelEsquema;
      for (let id = b.estilo, k = 0; id && k < 32; k++) {
        const e = estilosPorId.get(id);
        if (!e) break;
        if (e.p && e.p.nivelEsquema != null) return e.p.nivelEsquema;
        id = e.basadoEn;
      }
      return null;
    }

    let maestraPendiente = null;   // nombre de la página maestra con la que empieza el siguiente bloque
    function parrafos(b) {
      const p = b.p ? Object.assign({}, b.p) : null;
      if (p) { delete p.lista; delete p.nivelEsquema; }
      const nivel = nivelEsquemaDe(b);
      return partirPorSaltos(b.hijos).map((t, i) => {
        let extra = '';
        const q = p ? Object.assign({}, p) : {};
        if (i > 0) delete q.saltoAntes;   // el «salto antes» del párrafo vale para su primer trozo
        if (t.antes === 'pagina') q.saltoAntes = true;
        if (t.antes === 'columna') extra += ' fo:break-before="column"';
        if (t.despues === 'pagina') extra += ' fo:break-after="page"';
        if (t.despues === 'columna') extra += ' fo:break-after="column"';
        const maestra = i === 0 ? maestraPendiente : null;
        if (maestra) { maestraPendiente = null; delete q.saltoAntes; }   // cambiar de página maestra ya empieza hoja nueva
        const propiedades = parProps(Object.keys(q).length ? q : null, extra);
        const padre = ` style:parent-style-name="${escAttr(nombreEstilo(b.estilo || 'Normal'))}"`;
        const estiloAuto = propiedades || maestra
          ? automatico('P', 'paragraph', propiedades, `${padre}${maestra ? ` style:master-page-name="${escAttr(maestra)}"` : ''}`)
          : nombreEstilo(b.estilo || 'Normal');
        const estado = { espacio: true };
        const interior = runs(t.hijos, estado, 0);
        const etiqueta = nivel != null && nivel >= 0 && nivel <= 8 ? `text:h text:outline-level="${nivel + 1}"` : 'text:p';
        return `<${etiqueta}${estiloAuto ? ` text:style-name="${escAttr(estiloAuto)}"` : ''}>${interior}</${etiqueta.split(' ')[0]}>`;
      }).join('');
    }

    // ───────── Tablas ─────────
    function tabla(t, prof) {
      const filas = (t.filas || []).filter((f) => f.celdas && f.celdas.length);
      if (!filas.length || prof > MAX_PROFUNDIDAD_TABLAS) return '';
      // posición de cada celda en la rejilla de columnas
      const posiciones = filas.map((f) => { let c = 0; return f.celdas.map((celda) => { const x = c; c += Math.max(1, celda.fusionH || 1); return x; }); });
      const nCols = posiciones.reduce((m, ps, i) => Math.max(m, ps.length ? ps[ps.length - 1] + Math.max(1, filas[i].celdas[ps.length - 1].fusionH || 1) : 0), Math.max(t.anchos ? t.anchos.length : 0, 1));
      const anchos = Array.from({ length: nCols }, (_, i) => num(t.anchos && t.anchos[i], 0, 31680, 1440));
      const total = anchos.reduce((s, x) => s + x, 0);
      const maestra = maestraPendiente; maestraPendiente = null;
      const estiloTabla = automatico('Tbl', 'table', `<style:table-properties style:width="${pulgadas(total)}" table:align="left" fo:margin-left="${pulgadas(t.sangria || 0)}" style:border-model="collapsing"/>`, maestra ? ` style:master-page-name="${escAttr(maestra)}"` : '');
      let x = `<table:table table:name="Tabla${++contador}" table:style-name="${estiloTabla}">`;
      for (const w of anchos) x += `<table:table-column table:style-name="${automatico('Col', 'table-column', `<style:table-column-properties style:column-width="${pulgadas(w)}"/>`)}"/>`;
      const bordeTabla = t.bordes === false ? null : { tam: (t.bordes && t.bordes.tam) || 4, color: t.bordes && t.bordes.color };
      let abiertaCabecera = false;
      filas.forEach((f, fi) => {
        if (f.encabezado && !abiertaCabecera && fi === 0) { x += '<table:table-header-rows>'; abiertaCabecera = true; }
        if (abiertaCabecera && !f.encabezado) { x += '</table:table-header-rows>'; abiertaCabecera = false; }
        const estiloFila = f.alto ? automatico('Fila', 'table-row', `<style:table-row-properties style:min-row-height="${pulgadas(f.alto)}"/>`) : null;
        x += `<table:table-row${estiloFila ? ` table:style-name="${estiloFila}"` : ''}>`;
        let usadas = 0;
        f.celdas.forEach((c, ci) => {
          const col = posiciones[fi][ci], span = Math.max(1, c.fusionH || 1);
          usadas = col + span;
          if (c.fusionV === 'continua') { x += '<table:covered-table-cell/>'.repeat(span); return; }
          let filasSpan = 1;
          if (c.fusionV === 'inicio') {
            for (let k = fi + 1; k < filas.length; k++) {
              const idx = posiciones[k].indexOf(col);
              if (idx >= 0 && filas[k].celdas[idx].fusionV === 'continua') filasSpan++; else break;
            }
          }
          const lados = { arriba: null, abajo: null, izquierda: null, derecha: null };
          if (bordeTabla) for (const k of Object.keys(lados)) lados[k] = { tam: bordeTabla.tam, color: bordeTabla.color };
          for (const k of Object.keys(lados)) if (c.bordes && c.bordes[k]) lados[k] = c.bordes[k];
          let a = ' fo:padding-left="0.075in" fo:padding-right="0.075in" fo:padding-top="0in" fo:padding-bottom="0in"';
          a += Object.keys(lados).every((k) => !lados[k]) ? ' fo:border="none"' : Object.keys(lados).map((k) => (lados[k] ? bordesAttr({ [k]: lados[k] }) : ` fo:border-${{ arriba: 'top', abajo: 'bottom', izquierda: 'left', derecha: 'right' }[k]}="none"`)).join('');
          const fondo = hex6(c.relleno); if (fondo) a += ` fo:background-color="${fondo}"`;
          if (c.alineacionV) a += ` style:vertical-align="${c.alineacionV === 'centro' ? 'middle' : c.alineacionV === 'abajo' ? 'bottom' : 'top'}"`;
          const estiloCelda = automatico('Cel', 'table-cell', `<style:table-cell-properties${a}/>`);
          const contenido = bloques(c.hijos, prof + 1) || '<text:p/>';
          x += `<table:table-cell table:style-name="${estiloCelda}" office:value-type="string"${span > 1 ? ` table:number-columns-spanned="${span}"` : ''}${filasSpan > 1 ? ` table:number-rows-spanned="${filasSpan}"` : ''}>${contenido}</table:table-cell>`;
          if (span > 1) x += '<table:covered-table-cell/>'.repeat(span - 1);
        });
        for (let k = usadas; k < nCols; k++) x += '<table:table-cell office:value-type="string"><text:p/></table:table-cell>';
        x += '</table:table-row>';
      });
      if (abiertaCabecera) x += '</table:table-header-rows>';
      return x + '</table:table>';
    }

    // ───────── Bloques, con las listas que se abren y se cierran ─────────
    let idLista = 0;
    const ultimaPorLista = new Map();   // id de lista del contrato → xml:id de su último <text:list>
    function bloques(lista, prof) {
      let x = '';
      let abierta = null;   // { id, profundidad (niveles abiertos) }
      const cerrarLista = () => { if (abierta) { x += '</text:list-item></text:list>'.repeat(abierta.profundidad); abierta = null; } };
      for (const b of lista || []) {
        if (b.tipo === 'tabla') { cerrarLista(); x += tabla(b, prof); continue; }
        const l = b.p && b.p.lista && b.p.lista.id != null && listas.has(b.p.lista.id) ? b.p.lista : null;
        if (!l) { cerrarLista(); x += parrafos(b); continue; }
        const nivel = Math.max(0, Math.min(8, Math.round(Number(l.nivel) || 0)));
        if (abierta && abierta.id !== l.id) cerrarLista();
        if (!abierta) {
          const xmlId = `lst${++idLista}`;
          const previa = ultimaPorLista.get(l.id);
          x += `<text:list text:style-name="${listas.get(l.id)}" xml:id="${xmlId}"${previa ? ` text:continue-list="${previa}"` : ''}><text:list-item>`;
          ultimaPorLista.set(l.id, xmlId);
          abierta = { id: l.id, profundidad: 1 };
        } else if (nivel + 1 === abierta.profundidad) x += '</text:list-item><text:list-item>';
        else if (nivel + 1 < abierta.profundidad) {
          while (abierta.profundidad > nivel + 1) { x += '</text:list-item></text:list>'; abierta.profundidad--; }
          x += '</text:list-item><text:list-item>';
        }
        while (abierta.profundidad < nivel + 1) { x += '<text:list><text:list-item>'; abierta.profundidad++; }
        x += parrafos(b);
      }
      cerrarLista();
      return x;
    }

    // ───────── Páginas maestras ─────────
    const maestras = [];       // { nombre, layout, encabezado, pie }
    const maestraPorClave = new Map();
    const layouts = [];
    function maestraDe(s, indice) {
      const pg = s.pagina || {}, m = s.margenes || {};
      const ancho = num(pg.ancho || 12240, 720, 31680), alto = num(pg.alto || 15840, 720, 31680);
      const arriba = num(m.arriba == null ? 1440 : m.arriba, 0, 31680), abajo = num(m.abajo == null ? 1440 : m.abajo, 0, 31680);
      const tieneEnc = s.encabezado && s.encabezado.length, tienePie = s.pie && s.pie.length;
      const dEnc = num(m.encabezado == null ? 708 : m.encabezado, 0, 31680), dPie = num(m.pie == null ? 708 : m.pie, 0, 31680);
      const cols = s.columnas && s.columnas.n > 1 && s.tipo !== 'continua' ? s.columnas : null;
      const clave = JSON.stringify([ancho, alto, !!pg.horizontal, arriba, abajo, m.izquierda, m.derecha, dEnc, dPie, cols]) + '|' + (tieneEnc ? idDeLista(s.encabezado) : '') + '|' + (tienePie ? idDeLista(s.pie) : '');
      if (maestraPorClave.has(clave)) return maestraPorClave.get(clave);
      const n = maestras.length + 1;
      const nombre = n === 1 ? 'Standard' : `MP${n}`;
      const layout = `pm${n}`;
      let a = `fo:page-width="${pulgadas(ancho)}" fo:page-height="${pulgadas(alto)}" style:print-orientation="${pg.horizontal ? 'landscape' : 'portrait'}"`
        + ` fo:margin-top="${pulgadas(tieneEnc ? dEnc : arriba)}" fo:margin-bottom="${pulgadas(tienePie ? dPie : abajo)}" fo:margin-left="${pulgadas(m.izquierda == null ? 1440 : m.izquierda)}" fo:margin-right="${pulgadas(m.derecha == null ? 1440 : m.derecha)}" style:writing-mode="lr-tb"`;
      const columnas = cols ? `<style:columns fo:column-count="${num(cols.n, 2, 45, 2)}" fo:column-gap="${pulgadas(cols.espacio == null ? 720 : cols.espacio)}"/>` : '';
      layouts.push(`<style:page-layout style:name="${layout}"><style:page-layout-properties ${a}>${columnas}</style:page-layout-properties>`
        + (tieneEnc ? `<style:header-style><style:header-footer-properties fo:min-height="${pulgadas(Math.max(0, arriba - dEnc))}" fo:margin-bottom="0in"/></style:header-style>` : '')
        + (tienePie ? `<style:footer-style><style:header-footer-properties fo:min-height="${pulgadas(Math.max(0, abajo - dPie))}" fo:margin-top="0in"/></style:footer-style>` : '') + '</style:page-layout>');
      const registro = { nombre, layout, encabezado: tieneEnc ? s.encabezado : null, pie: tienePie ? s.pie : null };
      maestras.push(registro);
      maestraPorClave.set(clave, registro);
      void indice;
      return registro;
    }
    const idsListas = new WeakMap(); let siguienteIdLista = 0;
    const idDeLista = (l) => { if (!idsListas.has(l)) idsListas.set(l, ++siguienteIdLista); return idsListas.get(l); };

    // ───────── Cuerpo ─────────
    const secciones = contrato.secciones && contrato.secciones.length ? contrato.secciones : [{ hijos: [] }];
    let cuerpo = '';
    let primeraMaestra = null, seccionesTexto = 0;
    secciones.forEach((s, i) => {
      const esContinua = i > 0 && s.tipo === 'continua';
      if (i === 0 || (!esContinua && maestras.length < MAX_SECCIONES_MAESTRAS)) {
        const m = maestraDe(s, i);
        if (i === 0) primeraMaestra = m; else maestraPendiente = m.nombre;
        // una sección que empieza una hoja nueva con la misma maestra también necesita el salto: maestraPendiente lo da al primer párrafo
      }
      let interior = bloques(s.hijos, 0);
      if (maestraPendiente) { interior = `<text:p text:style-name="${automatico('P', 'paragraph', '', ` style:parent-style-name="Standard" style:master-page-name="${escAttr(maestraPendiente)}"`)}"/>` + interior; maestraPendiente = null; }
      if (esContinua && s.columnas && s.columnas.n > 1) {
        const est = automatico('Sec', 'section', `<style:section-properties text:dont-balance-text-columns="false" style:editable="false"><style:columns fo:column-count="${num(s.columnas.n, 2, 45, 2)}" fo:column-gap="${pulgadas(s.columnas.espacio == null ? 720 : s.columnas.espacio)}"/></style:section-properties>`);
        cuerpo += `<text:section text:style-name="${est}" text:name="Seccion${++seccionesTexto}">${interior}</text:section>`;
      } else cuerpo += interior;
    });
    void primeraMaestra;
    if (!cuerpo) cuerpo = '<text:p/>';

    // ───────── Estilos con nombre ─────────
    const pd = contrato.predeterminado || {};
    let nombrados = '';
    for (const e of (contrato.estilos || []).slice(0, 4096)) {
      if (e.tipo === 'tabla') continue;
      const familia = e.tipo === 'caracter' ? 'text' : 'paragraph';
      const nombre = nombreEstilo(e.id);
      const p = e.p ? Object.assign({}, e.p) : null;
      const nivel = p && p.nivelEsquema != null ? p.nivelEsquema : null;
      if (p) { delete p.lista; delete p.nivelEsquema; }
      const props = (familia === 'paragraph' ? parProps(p && Object.keys(p).length ? p : null) : '') + textProps(e.r);
      nombrados += `<style:style style:name="${escAttr(nombre)}" style:display-name="${escAttr(e.nombre || e.id)}" style:family="${familia}"`
        + (e.basadoEn && e.basadoEn !== e.id ? ` style:parent-style-name="${escAttr(nombreEstilo(e.basadoEn))}"` : '')
        + (familia === 'paragraph' && e.siguiente ? ` style:next-style-name="${escAttr(nombreEstilo(e.siguiente))}"` : '')
        + (familia === 'paragraph' && nivel != null && nivel >= 0 && nivel <= 8 ? ` style:default-outline-level="${nivel + 1}"` : '')
        + (familia === 'paragraph' ? ' style:class="text"' : '') + `>${props}</style:style>`;
    }
    if (!(contrato.estilos || []).some((e) => nombreEstilo(e.id) === 'Standard' && e.tipo !== 'caracter' && e.tipo !== 'tabla')) nombrados = '<style:style style:name="Standard" style:display-name="Normal" style:family="paragraph" style:class="text"/>' + nombrados;
    const pdP = pd.p ? Object.assign({}, pd.p) : null;
    if (pdP) { delete pdP.lista; delete pdP.nivelEsquema; }
    const textoPredeterminado = textProps(Object.assign({ tamano: 11 }, pd.r));
    const predeterminado = `<style:default-style style:family="paragraph">${parProps(pdP && Object.keys(pdP).length ? pdP : null, ' style:tab-stop-distance="0.5in"') || '<style:paragraph-properties style:tab-stop-distance="0.5in"/>'}${textoPredeterminado}</style:default-style>`;

    // Encabezados y pies: se escriben después del cuerpo porque dependen de estilos automáticos
    destinoAuto = { lista: estilosAutoCabeceras, mapa: mapaCabeceras };
    const maestrasXml = maestras.map((m) => {
      let h = '';
      const hf = (etiqueta, lista) => (lista ? `<style:${etiqueta}>${bloques(lista, 0) || '<text:p/>'}</style:${etiqueta}>` : '');
      h += hf('header', m.encabezado) + hf('footer', m.pie);
      return `<style:master-page style:name="${m.nombre}" style:page-layout-name="${m.layout}">${h}</style:master-page>`;
    }).join('');
    if (!maestras.length) { layouts.push('<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="8.5in" fo:page-height="11in" fo:margin-top="1in" fo:margin-bottom="1in" fo:margin-left="1in" fo:margin-right="1in"/></style:page-layout>'); }

    const fontFaces = '<office:font-face-decls>' + [...fuentes].map((f) => `<style:font-face style:name="${escAttr(f)}" svg:font-family="'${escAttr(f).replace(/'/g, '')}'" style:font-pitch="variable"/>`).join('') + '</office:font-face-decls>';
    const content = `${CABECERA}<office:document-content ${NS}>${fontFaces}<office:automatic-styles>${estilosAuto.join('')}${listaXml.join('')}</office:automatic-styles><office:body><office:text>${cuerpo}</office:text></office:body></office:document-content>`;
    const styles = `${CABECERA}<office:document-styles ${NS}>${fontFaces}<office:styles>${predeterminado}${nombrados}</office:styles><office:automatic-styles>${layouts.join('')}${estilosAutoCabeceras.join('')}${listaXml.join('')}</office:automatic-styles><office:master-styles>${maestrasXml || '<style:master-page style:name="Standard" style:page-layout-name="pm1"/>'}</office:master-styles></office:document-styles>`;
    const pr = contrato.propiedades || {};
    const ahora = new Date().toISOString().replace(/\.\d+Z$/, '');
    const meta = `${CABECERA}<office:document-meta ${NS}><office:meta>${pr.titulo ? `<dc:title>${esc(pr.titulo)}</dc:title>` : ''}${pr.autor ? `<meta:initial-creator>${esc(pr.autor)}</meta:initial-creator>` : ''}<meta:creation-date>${ahora}</meta:creation-date></office:meta></office:document-meta>`;
    const manifiesto = `${CABECERA}<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>`
      + ['content.xml', 'styles.xml', 'meta.xml'].map((n) => `<manifest:file-entry manifest:full-path="${n}" manifest:media-type="text/xml"/>`).join('')
      + medios.map((m) => `<manifest:file-entry manifest:full-path="${m.nombre}" manifest:media-type="${MIME[m.tipo]}"/>`).join('') + '</manifest:manifest>';

    const zip = new ZipPropio();
    zip.file('mimetype', 'application/vnd.oasis.opendocument.text', { compression: 'STORE' });
    zip.file('META-INF/manifest.xml', manifiesto);
    zip.file('content.xml', content);
    zip.file('styles.xml', styles);
    zip.file('meta.xml', meta);
    for (const m of medios) zip.file(m.nombre, m.bytes);
    return { bytes: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), mensajes };
  }

  globalThis.OdtEscritor = { generar };
})();

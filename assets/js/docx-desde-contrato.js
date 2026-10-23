/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Un documento de Word descrito con el «contrato de documento» (el que entregan DocLector y OdtLector, y el que recibirán los escritores de .doc y .odt)
// escrito como .docx. Es el gemelo de XlsxContrato/XlsxEscritor para Word: cualquier lector → cualquier escritor, sin un conversor por pareja, y todas las
// herramientas de Word (que leen .docx) aceptan los formatos viejos sin cambios.
//
//   const { bytes, mensajes } = await DocxDesdeContrato.generar(contrato);   // Uint8Array con el .docx
//
//   contrato = { propiedades?: { titulo, autor },
//                predeterminado?: { r: propsR, p: propsP },                       // docDefaults (fuente y tamaño normales del documento)
//                estilos?: [{ id, nombre, tipo: 'parrafo'|'caracter'|'tabla', basadoEn?, siguiente?, p?: propsP, r?: propsR }],   // id y basadoEn: ids de esta lista
//                listas?: [{ id, niveles: [{ formato: 'decimal'|'bullet'|'lowerLetter'|'upperLetter'|'lowerRoman'|'upperRoman'|…, texto: '%1.', izquierda, colgante,
//                                            inicio?, fuente?, r? }] }],             // texto: %1…%9 son los números de cada nivel
//                secciones: [{ pagina: { ancho, alto, horizontal? }, margenes: { arriba, abajo, izquierda, derecha, encabezado, pie }, tipo?: 'siguiente'|'continua',
//                              columnas?: { n, espacio }, encabezado?: [bloque], pie?: [bloque], hijos: [bloque] }] }
//   Las medidas de página y de párrafo van en twips (1/1440 de pulgada); los tamaños de letra en puntos.
//   propsP = { alineacion?: 'izquierda'|'centro'|'derecha'|'justificado', izq?, der?, primera? (negativo = sangría francesa), antes?, despues?, linea?, regla?: 'auto'|'minimo'|'exacto',
//              saltoAntes?, mantener?, mantenerJuntas?, nivelEsquema? (0-8), tabs?: [{ tipo: 'izquierda'|'centro'|'derecha'|'decimal', pos, relleno?: 'punto'|'guion'|'linea' }],
//              lista?: { id, nivel }, relleno?: 'RRGGBB', bordes?: { arriba|abajo|izquierda|derecha: { tam (1/8 pt), color, estilo } } }
//   propsR = { fuente?, tamano?, negrita?, cursiva?, subrayado?: true|'doble', tachado?, sup?, sub?, versalitas?, mayusculas?, oculto?, color?: 'RRGGBB', resaltado?: nombre de Word ('yellow'…) }
//   bloque = { tipo: 'parrafo', estilo?, p?, hijos: [run] } | { tipo: 'tabla', estilo?, anchos: [twips por columna], bordes?: false|{ tam, color }, sangria?,
//              filas: [{ alto?, encabezado?, celdas: [{ hijos: [bloque], ancho?, fusionH? (columnas), fusionV?: 'inicio'|'continua', relleno?, bordes?, alineacionV?: 'arriba'|'centro'|'abajo' }] }] }
//   run    = { texto, estilo?, r? }                         // '\t' = tabulación, '\n' = salto de línea
//          | { salto: 'pagina'|'columna'|'linea' }
//          | { imagen: { datos: Uint8Array, ancho, alto, texto? } }   // ancho/alto en puntos; solo en línea; png, jpeg, gif y bmp (otro formato se omite con un aviso)
//          | { enlace: { url? | ancla? }, hijos: [run] }
//          | { campo: 'PAGE'|'NUMPAGES'|…, resultado?: texto, r? }
//          | { marcador: nombre }
//
// Requiere zip-propio.js.
(function () {
  'use strict';
  const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const CABECERA = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const EMU_POR_PT = 12700;
  const INVALIDOS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;
  const limpiar = (s) => String(s).replace(INVALIDOS, '');
  const esc = (s) => limpiar(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const escAttr = (s) => esc(s).replace(/"/g, '&quot;');
  const entero = (n, min = -2147483647, max = 2147483647) => { const v = Math.round(Number(n)); return Number.isFinite(v) ? Math.max(min, Math.min(max, v)) : 0; };
  const hex6 = (c) => (/^[0-9a-fA-F]{6}$/.test(c || '') ? c.toUpperCase() : null);
  const sinProto = (o) => Object.assign(Object.create(null), o);
  // Word corta los nombres de marcador a 40 caracteres: el ancla de un enlace se corta igual para que siga apuntando al marcador
  const nombreMarcador = (n) => String(n).slice(0, 40);

  const ALINEACION = sinProto({ izquierda: 'left', centro: 'center', derecha: 'right', justificado: 'both' });
  const REGLA = sinProto({ minimo: 'atLeast', exacto: 'exact', auto: 'auto' });
  const TAB = sinProto({ izquierda: 'left', centro: 'center', derecha: 'right', decimal: 'decimal' });
  const RELLENO_TAB = sinProto({ punto: 'dot', guion: 'hyphen', linea: 'underscore' });
  const RESALTADOS = new Set(['yellow', 'green', 'cyan', 'magenta', 'blue', 'red', 'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray', 'black', 'white']);
  const FORMATOS_LISTA = new Set(['decimal', 'bullet', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman', 'ordinal', 'cardinalText', 'ordinalText', 'decimalZero', 'none']);
  const BORDE_ESTILO = sinProto({ simple: 'single', doble: 'double', punteado: 'dotted', guiones: 'dashed', ninguno: 'nil' });
  const TIPOS_IMAGEN = sinProto({ png: 'png', jpeg: 'jpeg', jpg: 'jpeg', gif: 'gif', bmp: 'bmp' });
  const MIME_IMAGEN = sinProto({ png: 'image/png', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp' });
  const MAX_PT = 1584;
  const MAX_SECCIONES = 10000;

  function tipoDeImagen(b, indicado) {
    if (b.length > 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4E && b[3] === 0x47) return 'png';
    if (b.length > 3 && b[0] === 0xFF && b[1] === 0xD8 && b[2] === 0xFF) return 'jpeg';
    if (b.length > 6 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return 'gif';
    if (b.length > 2 && b[0] === 0x42 && b[1] === 0x4D) return 'bmp';
    return TIPOS_IMAGEN[indicado] || null;
  }

  function propiedadesRun(r) {
    if (!r) return '';
    let x = '';
    if (r.fuente || r.fuenteAsia || r.fuenteCs) {
      const f = r.fuente ? escAttr(r.fuente) : '', a = r.fuenteAsia ? escAttr(r.fuenteAsia) : '', c = r.fuenteCs ? escAttr(r.fuenteCs) : '';
      x += `<w:rFonts${f ? ` w:ascii="${f}" w:hAnsi="${f}"` : ''}${a ? ` w:eastAsia="${a}"` : ''}${c ? ` w:cs="${c}"` : ''}/>`;
    }
    if (r.negrita != null) x += r.negrita ? '<w:b/><w:bCs/>' : '<w:b w:val="0"/><w:bCs w:val="0"/>';
    if (r.cursiva != null) x += r.cursiva ? '<w:i/><w:iCs/>' : '<w:i w:val="0"/><w:iCs w:val="0"/>';
    if (r.versalitas != null) x += r.versalitas ? '<w:smallCaps/>' : '<w:smallCaps w:val="0"/>';
    if (r.mayusculas != null) x += r.mayusculas ? '<w:caps/>' : '<w:caps w:val="0"/>';
    if (r.tachado != null) x += r.tachado ? '<w:strike/>' : '<w:strike w:val="0"/>';
    if (r.oculto) x += '<w:vanish/>';
    const color = hex6(r.color);
    if (color) x += `<w:color w:val="${color}"/>`;
    if (r.tamano != null) { const t = entero(r.tamano * 2, 2, 3276); x += `<w:sz w:val="${t}"/><w:szCs w:val="${t}"/>`; }
    if (r.resaltado && RESALTADOS.has(r.resaltado)) x += `<w:highlight w:val="${r.resaltado}"/>`;
    if (r.subrayado != null) x += r.subrayado ? `<w:u w:val="${r.subrayado === 'doble' ? 'double' : 'single'}"/>` : '<w:u w:val="none"/>';
    if (r.sup) x += '<w:vertAlign w:val="superscript"/>';
    else if (r.sub) x += '<w:vertAlign w:val="subscript"/>';
    return x;
  }

  function bordes(b, etiqueta = 'pBdr') {
    if (!b) return '';
    let x = '';
    for (const [clave, tag] of [['arriba', 'top'], ['izquierda', 'left'], ['abajo', 'bottom'], ['derecha', 'right']]) {
      const l = b[clave];
      if (!l) continue;
      const estilo = BORDE_ESTILO[l.estilo || 'simple'] || 'single';
      x += `<w:${tag} w:val="${estilo}" w:sz="${entero(l.tam == null ? 4 : l.tam, 2, 96)}" w:space="${entero(l.espacio || 0, 0, 31)}" w:color="${hex6(l.color) || 'auto'}"/>`;
    }
    return x ? `<w:${etiqueta}>${x}</w:${etiqueta}>` : '';
  }

  function propiedadesParrafo(p, estilo) {
    let x = '';
    if (estilo) x += `<w:pStyle w:val="${escAttr(estilo)}"/>`;
    if (p) {
      if (p.mantener) x += '<w:keepNext/>';
      if (p.mantenerJuntas) x += '<w:keepLines/>';
      if (p.saltoAntes) x += '<w:pageBreakBefore/>';
      if (p.lista) {
        if (p.lista.id == null) x += '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>';
        else x += `<w:numPr><w:ilvl w:val="${entero(p.lista.nivel || 0, 0, 8)}"/><w:numId w:val="${entero(p.lista.numId, 0, 65535)}"/></w:numPr>`;
      }
      x += bordes(p.bordes);
      const color = hex6(p.relleno);
      if (color) x += `<w:shd w:val="clear" w:color="auto" w:fill="${color}"/>`;
      if (p.tabs && p.tabs.length) {
        x += `<w:tabs>${p.tabs.slice(0, 64).map((t) => `<w:tab w:val="${TAB[t.tipo] || 'left'}"${RELLENO_TAB[t.relleno] ? ` w:leader="${RELLENO_TAB[t.relleno]}"` : ''} w:pos="${entero(t.pos, 0, 31680)}"/>`).join('')}</w:tabs>`;
      }
      let a = '';
      if (p.antes != null) a += ` w:before="${entero(p.antes, 0, 31680)}"`;
      if (p.despues != null) a += ` w:after="${entero(p.despues, 0, 31680)}"`;
      if (p.linea != null) a += ` w:line="${entero(p.linea, 0, 31680)}" w:lineRule="${REGLA[p.regla] || 'auto'}"`;
      if (a) x += `<w:spacing${a}/>`;
      let s = '';
      if (p.izq != null) s += ` w:left="${entero(p.izq, -31680, 31680)}"`;
      if (p.der != null) s += ` w:right="${entero(p.der, -31680, 31680)}"`;
      if (p.primera != null) s += p.primera < 0 ? ` w:hanging="${entero(-p.primera, 0, 31680)}"` : ` w:firstLine="${entero(p.primera, 0, 31680)}"`;
      if (s) x += `<w:ind${s}/>`;
      if (p.alineacion && ALINEACION[p.alineacion]) x += `<w:jc w:val="${ALINEACION[p.alineacion]}"/>`;
      if (p.nivelEsquema != null && p.nivelEsquema >= 0 && p.nivelEsquema <= 8) x += `<w:outlineLvl w:val="${entero(p.nivelEsquema)}"/>`;
    }
    return x;
  }

  async function generar(contrato) {
    const mensajes = [];
    const aviso = (m) => { if (!mensajes.includes(m)) mensajes.push(m); };
    const medios = [];            // { nombre, bytes, tipo }
    const mediosPorDatos = new Map();   // Uint8Array → medio: la misma imagen repetida (mismo arreglo) se guarda una sola vez
    const listas = new Map();     // id de contrato → { abstractId, numId }
    const numeraciones = [];      // en el orden de numId
    let idMarcador = 0, idDibujo = 0;
    const idsEstilo = new Map();  // id del contrato → id válido para XML
    // Solo letras ASCII, cifras, _ y -; dos ids que quedan iguales (p. ej. «Título_20_á» y «Título_20_é» de un .odt, o dos nombres en cirílico) reciben uno distinto cada uno
    const idsUsados = new Set(['Normal', 'DefaultParagraphFont', 'TableNormal']);
    idsEstilo.set('Normal', 'Normal');
    const idDe = (id) => {
      if (!idsEstilo.has(id)) {
        const base = String(id).replace(/[^A-Za-z0-9_-]/g, '').slice(0, 200) || 'E';
        let v = base;
        for (let k = 2; idsUsados.has(v); k++) v = base + '_' + k;
        idsUsados.add(v);
        idsEstilo.set(id, v);
      }
      return idsEstilo.get(id);
    };

    // Las listas del contrato reciben su numId (1, 2, 3…) antes de armar los párrafos
    (contrato.listas || []).forEach((l, i) => { listas.set(l.id, { abstractId: i, numId: i + 1 }); numeraciones.push(l); });
    const numIdDe = (id) => (listas.has(id) ? listas.get(id).numId : 0);

    // Se usa un solo conjunto de partes de relación por parte del paquete: document.xml, y uno por encabezado/pie
    function nuevoContexto() { return { rels: [], contador: 0, imagenes: new Map() }; }
    function relacion(ctx, tipo, destino, externo) { const id = `rIdX${++ctx.contador}`; ctx.rels.push({ id, tipo, destino, externo }); return id; }

    function dibujo(img, ctx) {
      const bytes = img.datos instanceof Uint8Array ? img.datos : new Uint8Array(img.datos || 0);
      const tipo = tipoDeImagen(bytes, img.tipo);
      if (!tipo) { aviso('imagen-formato'); return ''; }
      let medio = mediosPorDatos.get(bytes);
      if (!medio) {
        medio = { nombre: `image${medios.length + 1}.${tipo === 'jpeg' ? 'jpg' : tipo}`, bytes, tipo };
        medios.push(medio);
        mediosPorDatos.set(bytes, medio);
      }
      // una sola relación por imagen en cada parte (document.xml, cada encabezado y pie)
      let rid = ctx.imagenes.get(bytes);
      if (!rid) { rid = relacion(ctx, 'image', `media/${medio.nombre}`); ctx.imagenes.set(bytes, rid); }
      const cx = entero(Math.min(Math.max(Number(img.ancho) || 0, 1), MAX_PT) * EMU_POR_PT), cy = entero(Math.min(Math.max(Number(img.alto) || 0, 1), MAX_PT) * EMU_POR_PT);
      const id = ++idDibujo;
      const alt = img.texto ? ` descr="${escAttr(img.texto)}"` : '';
      return '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        + `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="Imagen ${id}"${alt}/>`
        + '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        + '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        + `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${id}" name="Imagen ${id}"/><pic:cNvPicPr/></pic:nvPicPr>`
        + `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
        + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>`
        + '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    }

    function rpr(r, estilo) {
      const x = (estilo ? `<w:rStyle w:val="${escAttr(idDe(estilo))}"/>` : '') + propiedadesRun(r);
      return x ? `<w:rPr>${x}</w:rPr>` : '';
    }

    function run(r, ctx, prof) {
      if (r.salto) return `<w:r>${r.salto === 'pagina' ? '<w:br w:type="page"/>' : r.salto === 'columna' ? '<w:br w:type="column"/>' : '<w:br/>'}</w:r>`;
      if (r.imagen) return dibujo(r.imagen, ctx);
      if (r.marcador != null) { const id = ++idMarcador; return `<w:bookmarkStart w:id="${id}" w:name="${escAttr(nombreMarcador(r.marcador))}"/><w:bookmarkEnd w:id="${id}"/>`; }
      if (r.enlace) {
        if (prof > 3) return '';
        const hijos = (r.hijos || []).map((h) => run(h, ctx, prof + 1)).join('');
        if (r.enlace.ancla) return `<w:hyperlink w:anchor="${escAttr(nombreMarcador(r.enlace.ancla))}">${hijos}</w:hyperlink>`;
        if (/^(https?:|mailto:|tel:|ftp:)/i.test(r.enlace.url || '')) return `<w:hyperlink r:id="${relacion(ctx, 'hyperlink', r.enlace.url, true)}">${hijos}</w:hyperlink>`;
        return hijos;   // un destino que no es una dirección web (javascript:, file:…) no se lleva
      }
      if (r.campo) {
        const p = rpr(r.r, r.estilo);
        return `<w:r>${p}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${p}<w:instrText xml:space="preserve"> ${esc(r.campo)} </w:instrText></w:r>`
          + `<w:r>${p}<w:fldChar w:fldCharType="separate"/></w:r><w:r>${p}<w:t xml:space="preserve">${esc(r.resultado == null ? '1' : r.resultado)}</w:t></w:r><w:r>${p}<w:fldChar w:fldCharType="end"/></w:r>`;
      }
      const p = rpr(r.r, r.estilo);
      let x = '';
      limpiar(String(r.texto == null ? '' : r.texto).replace(/\r\n?/g, '\n')).split('\n').forEach((linea, n) => {
        if (n > 0) x += `<w:r>${p}<w:br/></w:r>`;
        linea.split('\t').forEach((trozo, k) => {
          if (k > 0) x += `<w:r>${p}<w:tab/></w:r>`;
          if (trozo !== '') x += `<w:r>${p}<w:t xml:space="preserve">${esc(trozo)}</w:t></w:r>`;
        });
      });
      return x;
    }

    function parrafo(b, ctx, sectPr) {
      const p = b.p ? { ...b.p } : null;
      if (p && p.lista) p.lista = { ...p.lista, numId: numIdDe(p.lista.id) };
      const ppr = propiedadesParrafo(p, b.estilo ? idDe(b.estilo) : '') + (sectPr || '');
      return `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${(b.hijos || []).map((r) => run(r, ctx, 0)).join('')}</w:p>`;
    }

    function tabla(t, ctx, prof) {
      const filas = (t.filas || []).filter((f) => f.celdas && f.celdas.length);
      if (!filas.length || prof > 8) return '';
      const columnas = t.anchos && t.anchos.length ? t.anchos : Array.from({ length: filas.reduce((m, f) => Math.max(m, f.celdas.reduce((s, c) => s + (c.fusionH || 1), 0)), 1) }, () => 1440);
      const borde = t.bordes === false ? '' : `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((l) => `<w:${l} w:val="single" w:sz="${entero((t.bordes && t.bordes.tam) || 4, 2, 96)}" w:space="0" w:color="${hex6(t.bordes && t.bordes.color) || 'auto'}"/>`).join('')}</w:tblBorders>`;
      const filasXml = filas.map((f) => {
        const celdas = f.celdas.map((c) => {
          let tcPr = `<w:tcW w:w="${entero(c.ancho == null ? 1440 : c.ancho, 0, 31680)}" w:type="dxa"/>`;
          if (c.fusionH > 1) tcPr += `<w:gridSpan w:val="${entero(c.fusionH, 1, 63)}"/>`;
          if (c.fusionV) tcPr += c.fusionV === 'inicio' ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>';
          tcPr += bordes(c.bordes, 'tcBorders');
          const color = hex6(c.relleno);
          if (color) tcPr += `<w:shd w:val="clear" w:color="auto" w:fill="${color}"/>`;
          if (c.alineacionV) tcPr += `<w:vAlign w:val="${c.alineacionV === 'centro' ? 'center' : c.alineacionV === 'abajo' ? 'bottom' : 'top'}"/>`;
          const hijos = (c.hijos || []).map((h) => (h.tipo === 'tabla' ? tabla(h, ctx, prof + 1) : parrafo(h, ctx))).join('');
          return `<w:tc><w:tcPr>${tcPr}</w:tcPr>${hijos}${/<\/w:p>$/.test(hijos) ? '' : '<w:p/>'}</w:tc>`;   // una celda termina siempre en un párrafo
        }).join('');
        const trPr = (f.alto ? `<w:trHeight w:val="${entero(f.alto, 0, 31680)}" w:hRule="atLeast"/>` : '') + (f.encabezado ? '<w:tblHeader/>' : '');
        return `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}${celdas}</w:tr>`;
      }).join('');
      const sangria = t.sangria ? `<w:tblInd w:w="${entero(t.sangria, -31680, 31680)}" w:type="dxa"/>` : '';
      return `<w:tbl><w:tblPr>${t.estilo ? `<w:tblStyle w:val="${escAttr(idDe(t.estilo))}"/>` : ''}<w:tblW w:w="0" w:type="auto"/>${sangria}${borde}<w:tblLayout w:type="fixed"/>`
        + '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr>'
        + `<w:tblGrid>${columnas.map((w) => `<w:gridCol w:w="${entero(w, 0, 31680)}"/>`).join('')}</w:tblGrid>${filasXml}</w:tbl>`;
    }

    const bloques = (lista, ctx, ultimoSectPr) => {
      const xml = lista.map((b, i) => (b.tipo === 'tabla' ? tabla(b, ctx, 0) : parrafo(b, ctx, i === lista.length - 1 ? ultimoSectPr : ''))).join('');
      return xml;
    };

    // ───────── Secciones, encabezados y pies ─────────
    const partes = [];   // { ruta, tipo, rid, tipoRel, xml, ctx }
    // Más de MAX_SECCIONES (solo un archivo raro): el contenido de las de más pasa a la última, para no perder texto
    let secciones = contrato.secciones && contrato.secciones.length ? contrato.secciones : [{ hijos: [] }];
    if (secciones.length > MAX_SECCIONES) {
      const ultima = Object.assign({}, secciones[MAX_SECCIONES - 1], { hijos: [] });
      for (const s of secciones.slice(MAX_SECCIONES - 1)) for (const b of s.hijos || []) ultima.hijos.push(b);
      secciones = secciones.slice(0, MAX_SECCIONES - 1).concat([ultima]);
    }
    const ctxDoc = nuevoContexto();
    const partesPorLista = new Map();   // un mismo encabezado o pie (el mismo arreglo) que repiten varias secciones se escribe una sola vez
    let cuerpo = '';
    secciones.forEach((s, i) => {
      const refs = [];
      for (const [lista, raiz, etiqueta, nombre, tipoCt] of [[s.encabezado, 'w:hdr', 'header', 'header', 'header'], [s.pie, 'w:ftr', 'footer', 'footer', 'footer']]) {
        if (!lista || !lista.length) continue;
        if (partesPorLista.has(lista)) { refs.push(`<w:${etiqueta}Reference w:type="default" r:id="${partesPorLista.get(lista)}"/>`); continue; }
        const ctx = nuevoContexto();
        const n = partes.length + 1;
        const ruta = `word/${nombre}${n}.xml`;
        const rid = relacion(ctxDoc, etiqueta, `${nombre}${n}.xml`);
        const contenido = bloques(lista, ctx) || '<w:p/>';
        const xml = `${CABECERA}<${raiz} xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">${/<\/w:p>$/.test(contenido) ? contenido : contenido + '<w:p/>'}</${raiz}>`;
        partes.push({ ruta, tipo: `application/vnd.openxmlformats-officedocument.wordprocessingml.${tipoCt}+xml`, ctx, xml });
        refs.push(`<w:${etiqueta}Reference w:type="default" r:id="${rid}"/>`);
        partesPorLista.set(lista, rid);
      }
      const pg = s.pagina || {}, m = s.margenes || {};
      const ancho = entero(pg.ancho || 12240, 720, 31680), alto = entero(pg.alto || 15840, 720, 31680);
      const sectPr = `<w:sectPr>${refs.join('')}${i > 0 && s.tipo === 'continua' ? '<w:type w:val="continuous"/>' : i > 0 ? '<w:type w:val="nextPage"/>' : ''}`
        + `<w:pgSz w:w="${ancho}" w:h="${alto}"${pg.horizontal ? ' w:orient="landscape"' : ''}/>`
        + `<w:pgMar w:top="${entero(m.arriba == null ? 1440 : m.arriba, 0, 31680)}" w:right="${entero(m.derecha == null ? 1440 : m.derecha, 0, 31680)}" w:bottom="${entero(m.abajo == null ? 1440 : m.abajo, 0, 31680)}" w:left="${entero(m.izquierda == null ? 1440 : m.izquierda, 0, 31680)}" w:header="${entero(m.encabezado == null ? 708 : m.encabezado, 0, 31680)}" w:footer="${entero(m.pie == null ? 708 : m.pie, 0, 31680)}" w:gutter="0"/>`
        + (s.columnas && s.columnas.n > 1 ? `<w:cols w:num="${entero(s.columnas.n, 1, 45)}" w:space="${entero(s.columnas.espacio == null ? 720 : s.columnas.espacio, 0, 31680)}"/>` : '<w:cols w:space="708"/>') + '</w:sectPr>';
      const hijos = s.hijos || [];
      const ultimo = hijos[hijos.length - 1];
      // Las propiedades de una sección que no es la última van en su último párrafo; si termina en tabla (o no tiene nada), en un párrafo vacío
      if (i < secciones.length - 1) {
        if (!ultimo || ultimo.tipo === 'tabla') cuerpo += bloques(hijos, ctxDoc) + `<w:p><w:pPr>${sectPr}</w:pPr></w:p>`;
        else cuerpo += bloques(hijos, ctxDoc, sectPr);
      } else {
        cuerpo += bloques(hijos, ctxDoc);
        if (!ultimo || ultimo.tipo === 'tabla') cuerpo += '<w:p/>';
        cuerpo += sectPr;
      }
    });
    const documento = `${CABECERA}<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>${cuerpo}</w:body></w:document>`;

    // ───────── Estilos ─────────
    const pd = contrato.predeterminado || {};
    let estilosXml = `${CABECERA}<w:styles xmlns:w="${NS_W}"><w:docDefaults><w:rPrDefault><w:rPr>${propiedadesRun(pd.r) || '<w:sz w:val="22"/>'}</w:rPr></w:rPrDefault>`
      + `<w:pPrDefault><w:pPr>${propiedadesParrafo(pd.p, '') || ''}</w:pPr></w:pPrDefault></w:docDefaults>`;
    const estilos = contrato.estilos || [];
    if (!estilos.some((e) => e.id === 'Normal')) estilosXml += '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>';
    estilosXml += '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
      + '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>';
    for (const e of estilos.slice(0, 4096)) {
      const tipo = e.tipo === 'caracter' ? 'character' : e.tipo === 'tabla' ? 'table' : 'paragraph';
      const normal = e.id === 'Normal';
      if (tipo === 'table') { estilosXml += `<w:style w:type="table" w:styleId="${escAttr(idDe(e.id))}"><w:name w:val="${escAttr(e.nombre || e.id)}"/><w:basedOn w:val="TableNormal"/></w:style>`; continue; }
      estilosXml += `<w:style w:type="${tipo}"${normal ? ' w:default="1"' : ''} w:styleId="${escAttr(idDe(e.id))}"><w:name w:val="${escAttr(e.nombre || e.id)}"/>`
        + (e.basadoEn && e.basadoEn !== e.id ? `<w:basedOn w:val="${escAttr(idDe(e.basadoEn))}"/>` : '')
        + (e.siguiente ? `<w:next w:val="${escAttr(idDe(e.siguiente))}"/>` : '') + '<w:qFormat/>'
        + (tipo === 'paragraph' && e.p && propiedadesParrafo(e.p, '') ? `<w:pPr>${propiedadesParrafo(e.p, '')}</w:pPr>` : '')
        + (e.r && propiedadesRun(e.r) ? `<w:rPr>${propiedadesRun(e.r)}</w:rPr>` : '') + '</w:style>';
    }
    estilosXml += '</w:styles>';

    // ───────── Numeración ─────────
    let numeracion = '';
    if (numeraciones.length) {
      numeracion = `${CABECERA}<w:numbering xmlns:w="${NS_W}">`
        + numeraciones.map((l, i) => `<w:abstractNum w:abstractNumId="${i}"><w:multiLevelType w:val="hybridMultilevel"/>`
          + (l.niveles || []).slice(0, 9).map((v, n) => {
            const formato = FORMATOS_LISTA.has(v.formato) ? v.formato : 'decimal';
            return `<w:lvl w:ilvl="${n}"><w:start w:val="${entero(v.inicio == null ? 1 : v.inicio, 0, 32767)}"/><w:numFmt w:val="${formato}"/><w:lvlText w:val="${escAttr(v.texto == null ? '' : v.texto)}"/><w:lvlJc w:val="left"/>`
              + `<w:pPr><w:ind w:left="${entero(v.izquierda, 0, 31680)}" w:hanging="${entero(v.colgante, 0, 31680)}"/></w:pPr>`
              + (v.fuente || v.r ? `<w:rPr>${v.fuente ? `<w:rFonts w:ascii="${escAttr(v.fuente)}" w:hAnsi="${escAttr(v.fuente)}" w:hint="default"/>` : ''}${propiedadesRun(Object.assign({}, v.r, { fuente: undefined }))}</w:rPr>` : '') + '</w:lvl>';
          }).join('') + '</w:abstractNum>').join('')
        + numeraciones.map((l, i) => `<w:num w:numId="${i + 1}"><w:abstractNumId w:val="${i}"/></w:num>`).join('')
        + '</w:numbering>';
    }

    // ───────── Paquete ─────────
    const tiposMedios = [...new Set(medios.map((x) => x.tipo))];
    const tipos = `${CABECERA}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`
      + tiposMedios.map((t) => `<Default Extension="${t === 'jpeg' ? 'jpg' : t}" ContentType="${MIME_IMAGEN[t]}"/>`).join('')
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      + (numeracion ? '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' : '')
      + partes.map((p) => `<Override PartName="/${p.ruta}" ContentType="${p.tipo}"/>`).join('')
      + '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>';
    const relsRaiz = `${CABECERA}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>`
      + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>';
    const relXml = (ctx) => ctx.rels.map((r) => `<Relationship Id="${r.id}" Type="${REL}/${r.tipo}" Target="${escAttr(r.destino)}"${r.externo ? ' TargetMode="External"' : ''}/>`).join('');
    const relsDoc = `${CABECERA}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdEst" Type="${REL}/styles" Target="styles.xml"/>`
      + (numeracion ? `<Relationship Id="rIdNum" Type="${REL}/numbering" Target="numbering.xml"/>` : '') + relXml(ctxDoc) + '</Relationships>';
    const ahora = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    const pr = contrato.propiedades || {};
    const core = `${CABECERA}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`
      + (pr.titulo ? `<dc:title>${esc(pr.titulo)}</dc:title>` : '') + (pr.autor ? `<dc:creator>${esc(pr.autor)}</dc:creator>` : '')
      + `<dcterms:created xsi:type="dcterms:W3CDTF">${ahora}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${ahora}</dcterms:modified></cp:coreProperties>`;

    const zip = new ZipPropio();
    zip.file('[Content_Types].xml', tipos);
    zip.file('_rels/.rels', relsRaiz);
    zip.file('word/document.xml', documento);
    zip.file('word/_rels/document.xml.rels', relsDoc);
    zip.file('word/styles.xml', estilosXml);
    if (numeracion) zip.file('word/numbering.xml', numeracion);
    for (const p of partes) {
      zip.file(p.ruta, p.xml);
      if (p.ctx.rels.length) zip.file(p.ruta.replace('word/', 'word/_rels/') + '.rels', `${CABECERA}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relXml(p.ctx)}</Relationships>`);
    }
    for (const x of medios) zip.file(`word/media/${x.nombre}`, x.bytes);
    zip.file('docProps/core.xml', core);
    return { bytes: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), mensajes };
  }

  globalThis.DocxDesdeContrato = { generar };
})();

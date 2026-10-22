/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Escritor de .docx propio (reemplaza a la librería docx). Recibe un documento descrito con objetos planos
// y devuelve el .docx; el aspecto es el de un documento nuevo de Word (Aptos 12 pt, títulos azules, viñetas de Word).
//
//   const bytes = await EscritorDocx.generar(doc);   // Uint8Array
//   const blob  = await EscritorDocx.aBlob(doc);
//
//   doc = { numeraciones?: [{ id, niveles: [{ formato: 'decimal'|'lowerLetter'|'lowerRoman'|'bullet', texto, izquierda, colgante }] }],
//           pagina?: { ancho, alto, margen: { arriba, abajo, izquierda, derecha } },      // twips
//           encabezado?: [parrafo], pie?: [parrafo], hijos: [parrafo | tabla] }
//   parrafo = { tipo: 'parrafo', titulo?: 1..6, alineacion?, sangria?: { izquierda, primeraLinea, colgante },
//               espaciado?: { antes, despues, linea, regla: 'minimo'|'exacto'|'auto' }, tabs?: [{ tipo, pos }],
//               bordeAbajo?: { tam, color, espacio }, saltoAntes?, lista?: { id, nivel, instancia? } | { vineta: true, nivel }, hijos: [run] }
//   run     = { texto, negrita?, cursiva?, subrayado?, tamano?, fuente?, color? }       // '\t' = tabulación, '\n' = salto de línea
//           | { salto: 'pagina' } | { campo: 'pagina'|'paginas', …formato }
//           | { imagen: { datos: Uint8Array, tipo?, ancho, alto, flotante?: { x, y } } } // ancho/alto en px; x, y en EMU desde la hoja
//   tabla   = { tipo: 'tabla', anchoPct, filas: [{ celdas: [{ hijos: [parrafo] }] }] }
// La especificación completa y el porqué: Investigaciones\reemplazos-librerias\docx-pptx-escritor\00-plan.md (secciones 8 y 10).
//
// Requiere: zip-propio.js.

(function () {
  const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const CABECERA = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const EMU_POR_PX = 9525;

  // Página y márgenes como los crea Word (medidos con Word 16 real).
  const PAGINA = { ancho: 12240, alto: 15840, margen: { arriba: 1417, abajo: 1417, izquierda: 1701, derecha: 1701 } };

  // Los caracteres que XML 1.0 no admite (controles, sustitutos sueltos) hacen que Word rechace el archivo: se quitan.
  const INVALIDOS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;
  const limpiar = (s) => String(s).replace(INVALIDOS, '');
  const esc = (s) => limpiar(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const escAttr = (s) => esc(s).replace(/"/g, '&quot;');
  // Entero de 32 bits: «Infinity», «NaN» o «1e+21» en un atributo hacen que Word rechace el archivo.
  const entero = (n) => { const v = Math.round(Number(n)); return Number.isFinite(v) ? Math.max(-2147483647, Math.min(2147483647, v)) : 0; };

  // ---------- Partes fijas ----------

  const ESTILOS = (() => {
    const titulo = (n, antes, despues, rpr) => `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>`
      + `<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${antes}" w:after="${despues}"/><w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr>${rpr}</w:rPr></w:style>`;
    const mayor = '<w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display" w:eastAsia="Aptos Display" w:cs="Aptos Display"/>';
    const azul = '<w:color w:val="0F4761"/>';
    return `${CABECERA}<w:styles xmlns:w="${NS_W}">`
      + '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:eastAsia="Aptos" w:cs="Aptos"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault>'
      + '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="278" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
      + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
      + '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
      + '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
      + '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
      + titulo(1, 360, 80, `${mayor}${azul}<w:sz w:val="40"/><w:szCs w:val="40"/>`)
      + titulo(2, 160, 80, `${mayor}${azul}<w:sz w:val="32"/><w:szCs w:val="32"/>`)
      + titulo(3, 160, 80, `${azul}<w:sz w:val="28"/><w:szCs w:val="28"/>`)
      + titulo(4, 80, 40, `<w:i/><w:iCs/>${azul}`)
      + titulo(5, 80, 40, azul)
      + titulo(6, 40, 0, '<w:i/><w:iCs/><w:color w:val="595959"/>')
      + '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>'
      + '<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4252"/><w:tab w:val="right" w:pos="8504"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>'
      + '<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4252"/><w:tab w:val="right" w:pos="8504"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>'
      + '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
      + '<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>'
      + '</w:styles>';
  })();

  // compatibilityMode 15: sin él Word abre el archivo en «Modo de compatibilidad».
  const AJUSTES = `${CABECERA}<w:settings xmlns:w="${NS_W}"><w:zoom w:percent="100"/><w:defaultTabStop w:val="708"/><w:characterSpacingControl w:val="doNotCompress"/>`
    + '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>';

  // ---------- Imágenes ----------

  function tipoDeImagen(b, indicado) {
    if (b[0] === 0x89 && b[1] === 0x50) return 'png';
    if (b[0] === 0xFF && b[1] === 0xD8) return 'jpg';
    if (b[0] === 0x47 && b[1] === 0x49) return 'gif';
    if (b[0] === 0x42 && b[1] === 0x4D) return 'bmp';
    const t = String(indicado || '').toLowerCase().replace('jpeg', 'jpg');
    if (['png', 'jpg', 'gif', 'bmp'].includes(t)) return t;
    throw new Error('EscritorDocx: formato de imagen no admitido (solo png, jpg, gif y bmp)');
  }

  function aBytes(datos) {
    if (datos instanceof Uint8Array) return datos;
    if (datos instanceof ArrayBuffer) return new Uint8Array(datos);
    if (typeof datos === 'string') { const bin = atob(datos); const b = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) b[i] = bin.charCodeAt(i); return b; }
    throw new Error('EscritorDocx: los datos de la imagen tienen que ser Uint8Array, ArrayBuffer o base64');
  }

  // ---------- Generación ----------

  async function generar(doc) {
    const medios = []; // { nombre, bytes, tipo }
    let idDibujo = 0;
    let alturaRel = 251658240;

    // numeración: abstractNum 0 = viñetas de Word; luego una por cada numeraciones[]
    const numeraciones = doc.numeraciones || [];
    const abstractDe = new Map(numeraciones.map((n, i) => [n.id, i + 1]));
    const numIds = new Map(); // 'abs|instancia' → numId
    const nums = []; // { abs, reinicia }
    let usaViñetas = false;
    function numId(abs, instancia) {
      const clave = `${abs}|${instancia || 0}`;
      if (!numIds.has(clave)) { nums.push({ abs, reinicia: !!instancia }); numIds.set(clave, nums.length); }
      return numIds.get(clave);
    }

    function propiedadesRun(r) {
      let x = '';
      if (r.fuente) { const f = escAttr(r.fuente); x += `<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}"/>`; }
      if (r.negrita) x += '<w:b/><w:bCs/>';
      if (r.cursiva) x += '<w:i/><w:iCs/>';
      if (r.color) x += `<w:color w:val="${escAttr(r.color)}"/>`;
      if (r.tamano) x += `<w:sz w:val="${entero(r.tamano)}"/><w:szCs w:val="${entero(r.tamano)}"/>`;
      if (r.subrayado) x += '<w:u w:val="single"/>';
      return x ? `<w:rPr>${x}</w:rPr>` : '';
    }

    function dibujo(img) {
      const bytes = aBytes(img.datos);
      const tipo = tipoDeImagen(bytes, img.tipo);
      const n = medios.length + 1;
      medios.push({ nombre: `image${n}.${tipo}`, bytes, tipo });
      const rid = `rIdImg${n}`;
      const cx = Math.max(0, entero(img.ancho * EMU_POR_PX)), cy = Math.max(0, entero(img.alto * EMU_POR_PX)); // negativo: Word no abre
      const id = ++idDibujo;
      const grafico = `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">`
        + `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${id}" name="Imagen ${id}"/><pic:cNvPicPr/></pic:nvPicPr>`
        + `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
        + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
      const comun = `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="Imagen ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>`;
      const f = img.flotante;
      const cuerpo = f
        ? `<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${alturaRel += 1024}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>`
          + `<wp:positionH relativeFrom="page"><wp:posOffset>${entero(f.x)}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>${entero(f.y)}</wp:posOffset></wp:positionV>`
          + `${comun.replace('<wp:docPr', '<wp:wrapNone/><wp:docPr')}${grafico}</wp:anchor>`
        : `<wp:inline distT="0" distB="0" distL="0" distR="0">${comun}${grafico}</wp:inline>`;
      return `<w:r><w:drawing>${cuerpo}</w:drawing></w:r>`;
    }

    function campo(codigo, r) {
      const rpr = propiedadesRun(r);
      return `<w:r>${rpr}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${rpr}<w:instrText xml:space="preserve"> ${codigo} </w:instrText></w:r>`
        + `<w:r>${rpr}<w:fldChar w:fldCharType="separate"/></w:r><w:r>${rpr}<w:t>1</w:t></w:r><w:r>${rpr}<w:fldChar w:fldCharType="end"/></w:r>`;
    }

    function run(r) {
      if (r.salto === 'pagina') return '<w:r><w:br w:type="page"/></w:r>';
      if (r.campo) return campo(r.campo === 'pagina' ? 'PAGE' : 'NUMPAGES', r);
      if (r.imagen) return dibujo(r.imagen);
      const rpr = propiedadesRun(r);
      const lineas = limpiar(String(r.texto == null ? '' : r.texto).replace(/\r\n?/g, '\n')).split('\n');
      let x = '';
      lineas.forEach((linea, n) => {
        if (n > 0) x += `<w:r>${rpr}<w:br/></w:r>`;
        linea.split('\t').forEach((trozo, k) => {
          if (k > 0) x += `<w:r>${rpr}<w:tab/></w:r>`;
          if (trozo !== '') x += `<w:r>${rpr}<w:t xml:space="preserve">${esc(trozo)}</w:t></w:r>`;
        });
      });
      return x;
    }

    const ALINEACION = { izquierda: 'left', centro: 'center', derecha: 'right', justificado: 'both' };
    const REGLA = { minimo: 'atLeast', exacto: 'exact', auto: 'auto' };

    function parrafo(p) {
      let ppr = '';
      const l = p.lista;
      if (p.titulo) ppr += `<w:pStyle w:val="Heading${Math.min(6, Math.max(1, entero(p.titulo)))}"/>`;
      else if (l) ppr += '<w:pStyle w:val="ListParagraph"/>';
      if (p.saltoAntes) ppr += '<w:pageBreakBefore/>';
      if (l) {
        let id;
        if (l.vineta) { usaViñetas = true; id = numId(0, 0); }
        else if (abstractDe.has(l.id)) id = numId(abstractDe.get(l.id), l.instancia);
        else throw new Error(`EscritorDocx: la lista «${l.id}» no está en numeraciones`);
        ppr += `<w:numPr><w:ilvl w:val="${Math.min(8, Math.max(0, entero(l.nivel)))}"/><w:numId w:val="${id}"/></w:numPr>`;
      }
      if (p.bordeAbajo) { const b = p.bordeAbajo; ppr += `<w:pBdr><w:bottom w:val="single" w:sz="${entero(b.tam)}" w:space="${entero(b.espacio)}" w:color="${escAttr(b.color || '000000')}"/></w:pBdr>`; }
      if (p.tabs && p.tabs.length) ppr += `<w:tabs>${p.tabs.map((t) => `<w:tab w:val="${t.tipo === 'centro' ? 'center' : t.tipo === 'derecha' ? 'right' : 'left'}" w:pos="${entero(t.pos)}"/>`).join('')}</w:tabs>`;
      if (p.espaciado) {
        const e = p.espaciado;
        let a = '';
        if (e.antes != null) a += ` w:before="${entero(e.antes)}"`;
        if (e.despues != null) a += ` w:after="${entero(e.despues)}"`;
        if (e.linea != null) a += ` w:line="${entero(e.linea)}" w:lineRule="${REGLA[e.regla] || 'auto'}"`;
        if (a) ppr += `<w:spacing${a}/>`;
      }
      if (p.sangria) {
        const s = p.sangria;
        let a = '';
        if (s.izquierda != null) a += ` w:left="${entero(s.izquierda)}"`;
        if (s.colgante != null) a += ` w:hanging="${entero(s.colgante)}"`;
        else if (s.primeraLinea != null) a += ` w:firstLine="${entero(s.primeraLinea)}"`;
        if (a) ppr += `<w:ind${a}/>`;
      }
      if (p.alineacion && ALINEACION[p.alineacion]) ppr += `<w:jc w:val="${ALINEACION[p.alineacion]}"/>`;
      return `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${(p.hijos || []).map(run).join('')}</w:p>`;
    }

    const pagina = { ...PAGINA, ...(doc.pagina || {}) };
    pagina.margen = { ...PAGINA.margen, ...((doc.pagina && doc.pagina.margen) || {}) };
    const anchoUtil = pagina.ancho - pagina.margen.izquierda - pagina.margen.derecha;

    function tabla(t) {
      // Una fila sin celdas (p. ej. la de un Word con todas sus celdas combinadas en vertical) hace que Word no abra el archivo.
      const filasConCeldas = (t.filas || []).filter((f) => f.celdas && f.celdas.length);
      if (!filasConCeldas.length) return '';
      t = { ...t, filas: filasConCeldas };
      const columnas = t.filas.reduce((max, f) => Math.max(max, f.celdas.length), 1); // sin «...»: con 100 000 filas desborda la pila
      const ancho = Math.floor(anchoUtil / columnas);
      const pct = entero((t.anchoPct == null ? 100 : t.anchoPct) * 50);
      const filas = t.filas.map((f) => `<w:tr>${f.celdas.map((c) => {
        const hijos = c.hijos && c.hijos.length ? c.hijos.map(parrafo).join('') : '<w:p/>';
        return `<w:tc><w:tcPr><w:tcW w:w="${ancho}" w:type="dxa"/></w:tcPr>${hijos}</w:tc>`;
      }).join('')}</w:tr>`).join('');
      return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="${pct}" w:type="pct"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>`
        + `<w:tblGrid>${Array.from({ length: columnas }, () => `<w:gridCol w:w="${ancho}"/>`).join('')}</w:tblGrid>${filas}</w:tbl>`;
    }

    const hijos = doc.hijos || [];
    let cuerpo = hijos.map((b) => (b.tipo === 'tabla' ? tabla(b) : parrafo(b))).join('');
    if (!hijos.length || hijos[hijos.length - 1].tipo === 'tabla') cuerpo += '<w:p/>'; // Word exige un párrafo después de una tabla final

    // Encabezado y pie: sus propias partes
    const partes = []; // { ruta, tipo, xml, rid }
    const refs = [];
    const bloque = (lista, raiz, ruta, tipoRel, tipoContenido, etiqueta, rid) => {
      const desde = medios.length;
      const contenido = lista.map(parrafo).join('') || '<w:p/>'; // Word exige al menos un párrafo
      const imagenes = medios.map((x, i) => [x, i + 1]).slice(desde); // sus imágenes van en las relaciones de la parte, no en las de document.xml
      partes.push({ ruta, tipo: tipoContenido, rid, tipoRel, imagenes, xml: `${CABECERA}<${raiz} xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">${contenido}</${raiz}>` });
      refs.push(`<w:${etiqueta}Reference w:type="default" r:id="${rid}"/>`);
    };
    if (doc.encabezado) bloque(doc.encabezado, 'w:hdr', 'word/header1.xml', 'header', 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml', 'header', 'rIdEnc');
    if (doc.pie) bloque(doc.pie, 'w:ftr', 'word/footer1.xml', 'footer', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml', 'footer', 'rIdPie');

    const m = pagina.margen;
    const seccion = `<w:sectPr>${refs.join('')}<w:pgSz w:w="${entero(pagina.ancho)}" w:h="${entero(pagina.alto)}"/>`
      + `<w:pgMar w:top="${entero(m.arriba)}" w:right="${entero(m.derecha)}" w:bottom="${entero(m.abajo)}" w:left="${entero(m.izquierda)}" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>`;
    const documento = `${CABECERA}<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>${cuerpo}${seccion}</w:body></w:document>`;

    // numbering.xml
    let numeracion = '';
    if (nums.length) {
      const VINETAS = [['', 'Symbol'], ['o', 'Courier New'], ['', 'Wingdings']];
      const niveles = (defs) => defs.slice(0, 9).map((v, n) => { // Word admite 9 niveles (0–8); un décimo no abre
        const bullet = v.formato === 'bullet';
        return `<w:lvl w:ilvl="${n}"><w:start w:val="1"/><w:numFmt w:val="${bullet ? 'bullet' : escAttr(v.formato)}"/><w:lvlText w:val="${escAttr(v.texto)}"/><w:lvlJc w:val="left"/>`
          + `<w:pPr><w:ind w:left="${entero(v.izquierda)}" w:hanging="${entero(v.colgante)}"/></w:pPr>${v.fuente ? `<w:rPr><w:rFonts w:ascii="${escAttr(v.fuente)}" w:hAnsi="${escAttr(v.fuente)}" w:hint="default"/></w:rPr>` : ''}</w:lvl>`;
      }).join('');
      const abstractos = [
        { id: 0, defs: Array.from({ length: 9 }, (_, n) => ({ formato: 'bullet', texto: VINETAS[n % 3][0], fuente: VINETAS[n % 3][1], izquierda: 720 * (n + 1), colgante: 360 })) },
        ...numeraciones.map((c, i) => ({ id: i + 1, defs: c.niveles })),
      ].filter((a) => a.id > 0 || usaViñetas);
      numeracion = `${CABECERA}<w:numbering xmlns:w="${NS_W}">`
        + abstractos.map((a) => `<w:abstractNum w:abstractNumId="${a.id}"><w:multiLevelType w:val="hybridMultilevel"/>${niveles(a.defs)}</w:abstractNum>`).join('')
        + nums.map((n, i) => `<w:num w:numId="${i + 1}"><w:abstractNumId w:val="${n.abs}"/>${n.reinicia ? Array.from({ length: 9 }, (_, k) => `<w:lvlOverride w:ilvl="${k}"><w:startOverride w:val="1"/></w:lvlOverride>`).join('') : ''}</w:num>`).join('')
        + '</w:numbering>';
    }

    // Paquete
    const tiposMedios = [...new Set(medios.map((x) => x.tipo))];
    const MIME = { png: 'image/png', jpg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp' };
    const tipos = `${CABECERA}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`
      + tiposMedios.map((t) => `<Default Extension="${t}" ContentType="${MIME[t]}"/>`).join('')
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      + (numeracion ? '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' : '')
      + partes.map((p) => `<Override PartName="/${p.ruta}" ContentType="${p.tipo}"/>`).join('')
      + '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>';
    const relsRaiz = `${CABECERA}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>`
      + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>';
    const relsDoc = `${CABECERA}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdEst" Type="${REL}/styles" Target="styles.xml"/>`
      + `<Relationship Id="rIdAju" Type="${REL}/settings" Target="settings.xml"/>`
      + (numeracion ? `<Relationship Id="rIdNum" Type="${REL}/numbering" Target="numbering.xml"/>` : '')
      + partes.map((p) => `<Relationship Id="${p.rid}" Type="${REL}/${p.tipoRel}" Target="${p.ruta.replace('word/', '')}"/>`).join('')
      + medios.map((x, i) => `<Relationship Id="rIdImg${i + 1}" Type="${REL}/image" Target="media/${x.nombre}"/>`).join('')
      + '</Relationships>';
    const ahora = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    const core = `${CABECERA}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`
      + `<dcterms:created xsi:type="dcterms:W3CDTF">${ahora}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${ahora}</dcterms:modified></cp:coreProperties>`;

    const zip = new ZipPropio();
    zip.file('[Content_Types].xml', tipos);
    zip.file('_rels/.rels', relsRaiz);
    zip.file('word/document.xml', documento);
    zip.file('word/_rels/document.xml.rels', relsDoc);
    zip.file('word/styles.xml', ESTILOS);
    zip.file('word/settings.xml', AJUSTES);
    if (numeracion) zip.file('word/numbering.xml', numeracion);
    for (const p of partes) {
      zip.file(p.ruta, p.xml);
      if (p.imagenes.length) {
        zip.file(p.ruta.replace('word/', 'word/_rels/') + '.rels', `${CABECERA}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
          + p.imagenes.map(([x, n]) => `<Relationship Id="rIdImg${n}" Type="${REL}/image" Target="media/${x.nombre}"/>`).join('') + '</Relationships>');
      }
    }
    for (const x of medios) zip.file(`word/media/${x.nombre}`, x.bytes);
    zip.file('docProps/core.xml', core);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  async function aBlob(doc) {
    return new Blob([await generar(doc)], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  }

  globalThis.EscritorDocx = { generar, aBlob };
})();

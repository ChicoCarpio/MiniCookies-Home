/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Escritor de .pptx propio (reemplaza a PptxGenJS). Escribe una presentación como la de un PowerPoint nuevo:
// hoja 16:9 de 13,33 × 7,5 pulgadas, tema de Office (Aptos), título de 44 pt y texto de 28 pt con viñetas, y el título y el
// contenido de cada diapositiva son marcadores de posición de verdad (se pueden restablecer, cambiar de diseño, etc.).
//
//   const bytes = await EscritorPptx.generar(pres);   // Uint8Array
//   const blob  = await EscritorPptx.aBlob(pres);
//
//   pres = { titulo?, idiomas?: ['es-CR'], diapositivas: [{ titulo?, cuerpo?: [parrafo] }] }
//   parrafo = { hijos: [{ texto, negrita?, cursiva?, subrayado? }], lista?: 'vineta'|'numerada', nivel?: 0..8 }
//     sin `lista` el párrafo va sin viñeta. '\n' dentro del texto es un salto de línea.
// Si el texto no cabe en el cuadro, se guarda el ajuste automático (normAutofit) que PowerPoint aplicaría.
// Los valores de hoja, patrón, diseños y tema se midieron en un PowerPoint 16 real; están escritos a mano.
// Detalle y porqué: Investigaciones\reemplazos-librerias\docx-pptx-escritor\00-plan.md (sección 15).
//
// Requiere: zip-propio.js.

(function () {
  const CABECERA = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const CT = 'application/vnd.openxmlformats-officedocument.presentationml';

  // Los caracteres que XML 1.0 no admite se quitan (sin esto PowerPoint pide reparar el archivo).
  const INVALIDOS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;
  const limpiar = (s) => String(s == null ? '' : s).replace(INVALIDOS, '');
  const esc = (s) => limpiar(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const escAttr = (s) => esc(s).replace(/"/g, '&quot;');

  // Geometría del patrón (EMU): la de la plantilla de PowerPoint.
  const ANCHO = 12192000, ALTO = 6858000;
  const CUADRO_TITULO = { x: 838200, y: 365125, w: 10515600, h: 1325563 };
  const CUADRO_CUERPO = { x: 838200, y: 1825625, w: 10515600, h: 4351338 };
  const TAMANO_NIVEL = [28, 24, 20, 18, 18, 18, 18, 18, 18]; // pt, por nivel de viñeta
  const ANTES_NIVEL = [10, 5, 5, 5, 5, 5, 5, 5, 5];          // pt de espacio antes

  const TEXTOS = {
    es: { tema: 'Tema de Office', titulo: 'Haga clic para modificar el estilo de título del patrón', texto: 'Haga clic para modificar los estilos de texto del patrón',
      niveles: ['Segundo nivel', 'Tercer nivel', 'Cuarto nivel', 'Quinto nivel'], objetos: 'Título y objetos', soloTitulo: 'Solo el título', blanco: 'En blanco',
      nTitulo: 'Título', nContenido: 'Marcador de contenido', nFecha: 'Marcador de fecha', nPie: 'Marcador de pie de página', nNumero: 'Marcador de número de diapositiva', numero: '‹Nº›' },
    en: { tema: 'Office Theme', titulo: 'Click to edit Master title style', texto: 'Click to edit Master text styles',
      niveles: ['Second level', 'Third level', 'Fourth level', 'Fifth level'], objetos: 'Title and Content', soloTitulo: 'Title Only', blanco: 'Blank',
      nTitulo: 'Title', nContenido: 'Content Placeholder', nFecha: 'Date Placeholder', nPie: 'Footer Placeholder', nNumero: 'Slide Number Placeholder', numero: '‹#›' },
  };

  function idiomaDe(idiomas) {
    if (!idiomas) {
      const n = typeof navigator !== 'undefined' ? navigator : null;
      idiomas = n ? (n.languages && n.languages.length ? n.languages : [n.language]) : [];
    }
    const primero = [].concat(idiomas).filter(Boolean).map(String)[0] || '';
    return { lang: /^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*$/.test(primero) ? primero : '', textos: /^es(-|$)/i.test(primero) ? TEXTOS.es : TEXTOS.en };
  }

  // ---------- Partes fijas ----------

  const xfrm = (c) => `<a:xfrm><a:off x="${c.x}" y="${c.y}"/><a:ext cx="${c.w}" cy="${c.h}"/></a:xfrm>`;
  const GRUPO = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
  const estiloNivel = (n, extra) => `<a:lvl${n}pPr marL="${457200 * (n - 1)}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">${extra || ''}</a:lvl${n}pPr>`;

  function tema(t) {
    return `${CABECERA}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${escAttr(t.tema)}"><a:themeElements>`
      + '<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="0E2841"/></a:dk2><a:lt2><a:srgbClr val="E8E8E8"/></a:lt2>'
      + '<a:accent1><a:srgbClr val="156082"/></a:accent1><a:accent2><a:srgbClr val="E97132"/></a:accent2><a:accent3><a:srgbClr val="196B24"/></a:accent3><a:accent4><a:srgbClr val="0F9ED5"/></a:accent4><a:accent5><a:srgbClr val="A02B93"/></a:accent5><a:accent6><a:srgbClr val="4EA72E"/></a:accent6>'
      + '<a:hlink><a:srgbClr val="467886"/></a:hlink><a:folHlink><a:srgbClr val="96607D"/></a:folHlink></a:clrScheme>'
      + '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
      + '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
      + '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
      + '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst>'
      + [12700, 19050, 25400].map((w, i) => (i === 0 ? '<a:lnStyleLst>' : '') + `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>` + (i === 2 ? '</a:lnStyleLst>' : '')).join('')
      + '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>'
      + '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>'
      + '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme>'
      + '</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>';
  }

  // Marcador de posición del patrón o de un diseño
  const sp = (id, nombre, ph, spPr, cuerpo) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escAttr(nombre)} ${id - 1}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${ph}/></p:nvPr></p:nvSpPr><p:spPr>${spPr}</p:spPr>${cuerpo}</p:sp>`;
  const rect = (c) => `${xfrm(c)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`;
  const pie = (algn) => `<a:lstStyle><a:lvl1pPr algn="${algn}"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="82000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>`;

  function patron(t, lang) {
    const l = lang ? ` lang="${escAttr(lang)}"` : '';
    const parrafosCuerpo = [t.texto, ...t.niveles].map((x, n) => `<a:p><a:pPr lvl="${n}"/><a:r><a:rPr${l}/><a:t>${esc(x)}</a:t></a:r></a:p>`).join('');
    const bodyPr = (extra, hijo) => `<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"${extra || ''}>${hijo || ''}</a:bodyPr>`;
    const dt = sp(4, t.nFecha, ' type="dt" sz="half" idx="2"', rect({ x: 838200, y: 6356350, w: 2743200, h: 365125 }),
      `<p:txBody>${bodyPr(' anchor="ctr"')}${pie('l')}<a:p><a:endParaRPr${l}/></a:p></p:txBody>`);
    const ftr = sp(5, t.nPie, ' type="ftr" sz="quarter" idx="3"', rect({ x: 4038600, y: 6356350, w: 4114800, h: 365125 }),
      `<p:txBody>${bodyPr(' anchor="ctr"')}${pie('ctr')}<a:p><a:endParaRPr${l}/></a:p></p:txBody>`);
    const num = sp(6, t.nNumero, ' type="sldNum" sz="quarter" idx="4"', rect({ x: 8610600, y: 6356350, w: 2743200, h: 365125 }),
      `<p:txBody>${bodyPr(' anchor="ctr"')}${pie('r')}<a:p><a:fld id="{B0A0BC7D-E01F-4586-9EED-2E3260ECE181}" type="slidenum"><a:rPr${l}/><a:t>${esc(t.numero)}</a:t></a:fld><a:endParaRPr${l}/></a:p></p:txBody>`);
    const titulo = sp(2, t.nTitulo, ' type="title"', rect(CUADRO_TITULO),
      `<p:txBody>${bodyPr(' anchor="ctr"', '<a:normAutofit/>')}<a:lstStyle/><a:p><a:r><a:rPr${l}/><a:t>${esc(t.titulo)}</a:t></a:r></a:p></p:txBody>`);
    const cuerpo = sp(3, t.nContenido, ' type="body" idx="1"', rect(CUADRO_CUERPO),
      `<p:txBody>${bodyPr('', '<a:normAutofit/>')}<a:lstStyle/>${parrafosCuerpo}</p:txBody>`);
    const rel = (n) => `<p:sldLayoutId id="${2147483648 + n}" r:id="rId${n}"/>`;
    const ln = (n, antes) => `<a:lvl${n}pPr marL="${228600 + 457200 * (n - 1)}" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="${antes * 100}"/></a:spcBef>`
      + `<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="${TAMANO_NIVEL[n - 1] * 100}" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${n}pPr>`;
    const otro = (n) => estiloNivel(n, '<a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>');
    const nueve = (f) => Array.from({ length: 9 }, (_, i) => f(i + 1)).join('');
    return `${CABECERA}<p:sldMaster ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${GRUPO}${titulo}${cuerpo}${dt}${ftr}${num}</p:spTree></p:cSld>`
      + '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
      + `<p:sldLayoutIdLst>${rel(1)}${rel(2)}${rel(3)}</p:sldLayoutIdLst>`
      + '<p:txStyles><p:titleStyle><a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr></p:titleStyle>'
      + `<p:bodyStyle>${nueve((n) => ln(n, ANTES_NIVEL[n - 1]))}</p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>${nueve(otro)}</p:otherStyle></p:txStyles></p:sldMaster>`;
  }

  // Diseños: 1 = título y contenido, 2 = solo el título, 3 = en blanco
  function diseno(t, lang, tipo) {
    const l = lang ? ` lang="${escAttr(lang)}"` : '';
    const dt = sp(4, t.nFecha, ' type="dt" sz="half" idx="10"', '', `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr${l}/></a:p></p:txBody>`);
    const ftr = sp(5, t.nPie, ' type="ftr" sz="quarter" idx="11"', '', `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr${l}/></a:p></p:txBody>`);
    const num = sp(6, t.nNumero, ' type="sldNum" sz="quarter" idx="12"', '', `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{B0A0BC7D-E01F-4586-9EED-2E3260ECE181}" type="slidenum"><a:rPr${l}/><a:t>${esc(t.numero)}</a:t></a:fld><a:endParaRPr${l}/></a:p></p:txBody>`);
    const titulo = sp(2, t.nTitulo, ' type="title"', '', `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr${l}/><a:t>${esc(t.titulo)}</a:t></a:r></a:p></p:txBody>`);
    const cuerpo = sp(3, t.nContenido, ' idx="1"', '', `<p:txBody><a:bodyPr/><a:lstStyle/>${[t.texto, ...t.niveles].map((x, n) => `<a:p><a:pPr lvl="${n}"/><a:r><a:rPr${l}/><a:t>${esc(x)}</a:t></a:r></a:p>`).join('')}</p:txBody>`);
    const [nombre, atributo, formas] = {
      obj: [t.objetos, 'obj', titulo + cuerpo], titleOnly: [t.soloTitulo, 'titleOnly', titulo], blank: [t.blanco, 'blank', ''],
    }[tipo];
    return `${CABECERA}<p:sldLayout ${NS} type="${atributo}" preserve="1"><p:cSld name="${escAttr(nombre)}"><p:spTree>${GRUPO}${formas}${dt}${ftr}${num}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
  }

  // ---------- Texto de las diapositivas ----------

  const tamanoNivel = (n) => TAMANO_NIVEL[Math.min(8, Math.max(0, n))];

  // Estima la altura (pt) del texto del cuadro y devuelve el ajuste que PowerPoint guardaría (fontScale y lnSpcReduction).
  function ajusteAutomatico(parrafos, caja) {
    const ancho = caja.w / 12700 - 14.4, alto = caja.h / 12700 - 7.2; // pt, sin los márgenes internos
    const altura = (escala, reduccion) => parrafos.reduce((suma, p) => {
      const nivel = Math.min(8, Math.max(0, p.nivel || 0));
      const sangria = (228600 + 457200 * nivel) / 12700;
      const sz = tamanoNivel(nivel) * escala;
      const texto = (p.hijos || []).map((r) => (r.texto == null ? '' : String(r.texto))).join('');
      const lineasDeTexto = texto.split('\n').reduce((n, linea) => n + Math.max(1, Math.ceil((linea.length * sz * 0.52) / Math.max(sz, ancho - sangria))), 0);
      return suma + lineasDeTexto * sz * 1.2 * (0.9 - reduccion) + ANTES_NIVEL[nivel] * escala;
    }, 0);
    if (altura(1, 0) <= alto) return null;
    for (const escala of [0.925, 0.85, 0.775, 0.7, 0.625, 0.55, 0.475, 0.4, 0.325, 0.25]) {
      const reduccion = escala > 0.7 ? 0.1 : 0.2;
      if (altura(escala, reduccion) <= alto || escala === 0.25) return { escala: Math.round(escala * 100000), reduccion: Math.round(reduccion * 100000) };
    }
    return null;
  }

  function corrida(r, lang) {
    const atributos = (lang ? ` lang="${escAttr(lang)}"` : '') + (r.negrita ? ' b="1"' : '') + (r.cursiva ? ' i="1"' : '') + (r.subrayado ? ' u="sng"' : '') + ' dirty="0"';
    const partes = limpiar(r.texto).replace(/\r\n?/g, '\n').split('\n');
    return partes.map((parte, n) => (n > 0 ? `<a:br><a:rPr${atributos}/></a:br>` : '') + (parte === '' ? '' : `<a:r><a:rPr${atributos}/><a:t>${esc(parte)}</a:t></a:r>`)).join('');
  }

  function parrafo(p, lang) {
    const nivel = Math.min(8, Math.max(0, Math.round(Number(p.nivel) || 0)));
    let ppr;
    // La sangría de los números es la que pone PowerPoint (514350 EMU): con la de las viñetas el número queda pegado al texto.
    if (p.lista === 'numerada') ppr = `<a:pPr marL="${514350 + 457200 * nivel}" indent="-514350"${nivel ? ` lvl="${nivel}"` : ''}><a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/></a:pPr>`;
    else if (p.lista === 'vineta') ppr = nivel ? `<a:pPr lvl="${nivel}"/>` : '';
    else ppr = '<a:pPr marL="0" indent="0"><a:buNone/></a:pPr>';
    const fin = `<a:endParaRPr${lang ? ` lang="${escAttr(lang)}"` : ''} dirty="0"/>`;
    return `<a:p>${ppr}${(p.hijos || []).map((r) => corrida(r, lang)).join('')}${fin}</a:p>`;
  }

  function diapositiva(d, lang, t) {
    const l = lang ? ` lang="${escAttr(lang)}"` : '';
    const titulo = limpiar(d.titulo).replace(/[\r\n]+/g, ' ').trim();
    const cuerpo = (d.cuerpo || []).filter(Boolean);
    const hayCuerpo = cuerpo.some((p) => (p.hijos || []).some((r) => limpiar(r.texto).trim() !== ''));
    const formas = [];
    if (titulo) formas.push(sp(2, t.nTitulo, ' type="title"', '', `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr${l} dirty="0"/><a:t>${esc(titulo)}</a:t></a:r></a:p></p:txBody>`));
    if (hayCuerpo) {
      const ajuste = ajusteAutomatico(cuerpo, CUADRO_CUERPO);
      const auto = ajuste ? `<a:bodyPr><a:normAutofit fontScale="${ajuste.escala}" lnSpcReduction="${ajuste.reduccion}"/></a:bodyPr>` : '<a:bodyPr><a:normAutofit/></a:bodyPr>';
      formas.push(sp(3, t.nContenido, ' idx="1"', '', `<p:txBody>${auto}<a:lstStyle/>${cuerpo.map((p) => parrafo(p, lang)).join('')}</p:txBody>`));
    }
    const diseno = hayCuerpo ? 1 : titulo ? 2 : 3;
    return { diseno, xml: `${CABECERA}<p:sld ${NS}><p:cSld><p:spTree>${GRUPO}${formas.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>` };
  }

  // ---------- Paquete ----------

  async function generar(pres) {
    const diapositivas = pres.diapositivas || [];
    if (!diapositivas.length) throw new Error('EscritorPptx: la presentación necesita al menos una diapositiva');
    const { lang, textos: t } = idiomaDe(pres.idiomas);
    const hechas = diapositivas.map((d) => diapositiva(d, lang, t));
    const n = hechas.length;

    const tipos = `${CABECERA}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`
      + `<Override PartName="/ppt/presentation.xml" ContentType="${CT}.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${CT}.slideMaster+xml"/>`
      + [1, 2, 3].map((i) => `<Override PartName="/ppt/slideLayouts/slideLayout${i}.xml" ContentType="${CT}.slideLayout+xml"/>`).join('')
      + hechas.map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="${CT}.slide+xml"/>`).join('')
      + `<Override PartName="/ppt/presProps.xml" ContentType="${CT}.presProps+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="${CT}.viewProps+xml"/><Override PartName="/ppt/tableStyles.xml" ContentType="${CT}.tableStyles+xml"/>`
      + '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>';
    const rels = (lista) => `${CABECERA}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${lista.map(([id, tipo, destino]) => `<Relationship Id="${id}" Type="${tipo}" Target="${destino}"/>`).join('')}</Relationships>`;
    const relRaiz = rels([['rId1', `${REL}/officeDocument`, 'ppt/presentation.xml'], ['rId2', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml']]);
    const relPres = rels([
      ['rId1', `${REL}/slideMaster`, 'slideMasters/slideMaster1.xml'], ['rId2', `${REL}/presProps`, 'presProps.xml'], ['rId3', `${REL}/viewProps`, 'viewProps.xml'],
      ['rId4', `${REL}/theme`, 'theme/theme1.xml'], ['rId5', `${REL}/tableStyles`, 'tableStyles.xml'],
      ...hechas.map((_, i) => [`rId${10 + i}`, `${REL}/slide`, `slides/slide${i + 1}.xml`]),
    ]);
    const nivelesPred = Array.from({ length: 9 }, (_, i) => estiloNivel(i + 1, '<a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>')).join('');
    const presentacion = `${CABECERA}<p:presentation ${NS} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>`
      + `<p:sldIdLst>${hechas.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${10 + i}"/>`).join('')}</p:sldIdLst><p:sldSz cx="${ANCHO}" cy="${ALTO}"/><p:notesSz cx="6858000" cy="9144000"/>`
      + `<p:defaultTextStyle><a:defPPr><a:defRPr${lang ? ` lang="${escAttr(lang)}"` : ''}/></a:defPPr>${nivelesPred}</p:defaultTextStyle></p:presentation>`;
    const ahora = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    const core = `${CABECERA}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`
      + (pres.titulo ? `<dc:title>${esc(pres.titulo)}</dc:title>` : '')
      + `<dcterms:created xsi:type="dcterms:W3CDTF">${ahora}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${ahora}</dcterms:modified></cp:coreProperties>`;

    const zip = new ZipPropio();
    zip.file('[Content_Types].xml', tipos);
    zip.file('_rels/.rels', relRaiz);
    zip.file('ppt/presentation.xml', presentacion);
    zip.file('ppt/_rels/presentation.xml.rels', relPres);
    zip.file('ppt/presProps.xml', `${CABECERA}<p:presentationPr ${NS}/>`);
    zip.file('ppt/viewProps.xml', `${CABECERA}<p:viewPr ${NS}><p:gridSpacing cx="72008" cy="72008"/></p:viewPr>`);
    zip.file('ppt/tableStyles.xml', `${CABECERA}<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`);
    zip.file('ppt/theme/theme1.xml', tema(t));
    zip.file('ppt/slideMasters/slideMaster1.xml', patron(t, lang));
    zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', rels([['rId1', `${REL}/slideLayout`, '../slideLayouts/slideLayout1.xml'], ['rId2', `${REL}/slideLayout`, '../slideLayouts/slideLayout2.xml'], ['rId3', `${REL}/slideLayout`, '../slideLayouts/slideLayout3.xml'], ['rId4', `${REL}/theme`, '../theme/theme1.xml']]));
    ['obj', 'titleOnly', 'blank'].forEach((tipo, i) => {
      zip.file(`ppt/slideLayouts/slideLayout${i + 1}.xml`, diseno(t, lang, tipo));
      zip.file(`ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`, rels([['rId1', `${REL}/slideMaster`, '../slideMasters/slideMaster1.xml']]));
    });
    hechas.forEach((h, i) => {
      zip.file(`ppt/slides/slide${i + 1}.xml`, h.xml);
      zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, rels([['rId1', `${REL}/slideLayout`, `../slideLayouts/slideLayout${h.diseno}.xml`]]));
    });
    zip.file('docProps/core.xml', core);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  async function aBlob(pres) {
    return new Blob([await generar(pres)], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
  }

  globalThis.EscritorPptx = { generar, aBlob };
})();

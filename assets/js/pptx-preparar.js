/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Prepara un .pptx antes de dárselo a PptxViewJS: reescribe, sobre una copia,
// lo que esa librería dibuja mal, en términos que sí entiende. No cambia el
// archivo del usuario. Medido el 2026-10-09 contra PowerPoint real (ver
// hallazgos.md, "Motor de PowerPoint: diagnóstico de fidelidad"):
//   - El rectángulo "useBgFill" (relleno = fondo de la diapositiva) salía del
//     color accent1 del tema y tapaba el fondo.
//   - Los cuadros de texto sin relleno ni estilo salían con accent1.
//   - El SmartArt salía como un rectángulo gris: acá se reemplaza por las
//     formas que PowerPoint ya dejó dibujadas en ppt/diagrams/drawingN.xml.
//   - El autoajuste (normAutofit fontScale/lnSpcReduction) se ignoraba.
//   - El interlineado en % se aplicaba sobre el tamaño de la letra y no sobre
//     la altura de su línea (~1,2 veces): las líneas se pisaban.
//   - Las fuentes del tema (+mn-lt, +mj-lt) no se resolvían.
//   - Las viñetas de Wingdings/Symbol salían como la letra ("v" en vez de ❖).
// Requiere JSZip. Uso: const buffer = await PptxPreparar.preparar(archivo).

const PptxPreparar = (() => {
  const NS = {
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  };
  const TIPO_DIBUJO_DIAGRAMA = 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing';
  const TIPO_DATOS_DIAGRAMA = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData';
  const RELLENOS = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];
  // Altura de una línea con interlineado sencillo, en veces el tamaño de la
  // letra. Es la que usa PptxViewJS cuando no hay interlineado (120 %).
  const ALTO_LINEA = 1.2;
  // Wingdings y Symbol: el carácter que guarda PowerPoint es la posición en
  // esa fuente, no el símbolo. Las viñetas más comunes.
  const VINETAS = {
    wingdings: { v: '❖', 'Ø': '➢', '§': '▪', 'ü': '✔', q: '❑', n: '■', l: '●', 'Ÿ': '•', 'à': '➔', 'è': '➜', 'ê': '➢', 'ð': '⇒', 'o': '□', 'p': '◻', 'ª': '✦', 'w': '⬥', 'u': '◆', 's': '⬧', 't': '⧫' },
    symbol: { '·': '•', 'Þ': '⇒', '®': '→', 'à': '◊', '¨': '♣', '©': '♦', 'ª': '♥', '«': '♠' },
  };

  const { hijo, hijos, ruta, todos, parsear } = PptxXml;
  const crear = (doc, ns, nombre) => doc.createElementNS(NS[ns], `${ns}:${nombre}`);

  // "../media/x.png" relativo a "ppt/slides/slide1.xml" → "ppt/media/x.png"
  function resolverRuta(desde, destino) {
    try { destino = decodeURIComponent(destino); } catch { /* se deja como está */ }
    if (destino.startsWith('/')) return destino.slice(1);
    const partes = desde.split('/').slice(0, -1);
    for (const p of destino.split('/')) {
      if (p === '..') partes.pop();
      else if (p !== '.') partes.push(p);
    }
    return partes.join('/');
  }
  function rutaRelativa(desde, destino) {
    const a = desde.split('/').slice(0, -1);
    const b = destino.split('/');
    let i = 0;
    while (i < a.length && a[i] === b[i]) i++;
    return '../'.repeat(a.length - i) + b.slice(i).join('/');
  }
  const rutaRels = (parte) => parte.replace(/([^/]+)$/, '_rels/$1.rels');

  async function leerRels(zip, parte) {
    const archivo = zip.file(rutaRels(parte));
    const doc = archivo ? parsear(await archivo.async('string')) : null;
    const lista = doc ? todos(doc, 'Relationship').map((r) => ({
      id: r.getAttribute('Id'), tipo: r.getAttribute('Type'),
      destino: r.getAttribute('TargetMode') === 'External' || !r.getAttribute('Target') ? null : resolverRuta(parte, r.getAttribute('Target')),
    })) : [];
    return { doc, lista, porId: (id) => lista.find((r) => r.id === id), porTipo: (t) => lista.filter((r) => r.destino && (r.tipo || '').endsWith('/' + t)) };
  }

  // --- Contexto: diseño, patrón, tema y estilo de texto por defecto ---------
  async function crearContexto(zip) {
    const cache = new Map();
    async function doc(parte) {
      if (!cache.has(parte)) {
        const archivo = zip.file(parte);
        cache.set(parte, archivo ? { doc: parsear(await archivo.async('string')), rels: await leerRels(zip, parte) } : null);
      }
      return cache.get(parte);
    }
    const presentacion = await doc('ppt/presentation.xml');
    return { doc, textoPorDefecto: presentacion && ruta(presentacion.doc.documentElement, 'defaultTextStyle') };
  }

  async function cadenaDeParte(ctx, parte) {
    const propia = await ctx.doc(parte);
    const cadena = [propia];
    let actual = propia;
    // Una diapositiva → su diseño → su patrón; un diseño → su patrón.
    for (const tipo of ['slideLayout', 'slideMaster']) {
      const rel = actual.rels.porTipo(tipo)[0];
      if (!rel) continue;
      actual = await ctx.doc(rel.destino);
      if (!actual) break;
      cadena.push(actual);
    }
    const patron = cadena.find((c) => c.doc.documentElement.localName === 'sldMaster');
    const relTema = patron && patron.rels.porTipo('theme')[0];
    const tema = relTema ? await ctx.doc(relTema.destino) : null;
    return { cadena, patron, tema };
  }

  // --- Marcadores de posición: el del diseño y el del patrón que le tocan ---
  function datosMarcador(sp) {
    const ph = ruta(sp, 'nvSpPr', 'nvPr', 'ph');
    if (!ph) return null;
    return { tipo: ph.getAttribute('type') || 'obj', idx: ph.getAttribute('idx') };
  }
  const tipoBase = (t) => (t === 'ctrTitle' ? 'title' : ['subTitle', 'obj', 'body'].includes(t) ? 'body' : t);
  function buscarMarcador(docParte, marcador, esPatron) {
    const formas = todos(docParte, 'sp').filter((s) => datosMarcador(s));
    if (!esPatron && marcador.idx != null) {
      const porIdx = formas.find((s) => datosMarcador(s).idx === marcador.idx);
      if (porIdx) return porIdx;
    }
    return formas.find((s) => tipoBase(datosMarcador(s).tipo) === tipoBase(marcador.tipo)) || null;
  }
  // Las formas de las que hereda una forma, de la más cercana a la más lejana.
  function herencia(sp, contexto) {
    const marcador = datosMarcador(sp);
    if (!marcador) return [];
    return contexto.cadena.slice(1).map((c) => buscarMarcador(c.doc, marcador, c === contexto.patron)).filter(Boolean);
  }

  // Listas de estilos de texto que aplican a una forma, en orden de prioridad.
  function estilosDeTexto(sp, contexto, ctx) {
    const lista = [ruta(sp, 'txBody', 'lstStyle')];
    for (const f of herencia(sp, contexto)) lista.push(ruta(f, 'txBody', 'lstStyle'));
    const txStyles = contexto.patron && ruta(contexto.patron.doc.documentElement, 'txStyles');
    const marcador = datosMarcador(sp);
    if (marcador) {
      const base = tipoBase(marcador.tipo);
      lista.push(hijo(txStyles, base === 'title' ? 'titleStyle' : base === 'body' ? 'bodyStyle' : 'otherStyle'));
    } else {
      lista.push(hijo(txStyles, 'otherStyle'), ctx.textoPorDefecto);
    }
    return lista.filter(Boolean);
  }
  const nivelDe = (estilo, nivel) => hijo(estilo, `lvl${nivel + 1}pPr`);

  function tamanoEfectivo(rPr, estilos, nivel) {
    if (rPr && rPr.hasAttribute('sz')) return parseInt(rPr.getAttribute('sz'), 10);
    for (const e of estilos) {
      const def = hijo(nivelDe(e, nivel), 'defRPr');
      if (def && def.hasAttribute('sz')) return parseInt(def.getAttribute('sz'), 10);
    }
    return 1800;
  }
  function interlineadoEfectivo(pPr, estilos, nivel) {
    for (const fuente of [pPr, ...estilos.map((e) => nivelDe(e, nivel))]) {
      const ln = hijo(fuente, 'lnSpc');
      if (!ln) continue;
      const pct = hijo(ln, 'spcPct');
      if (pct) return { pct: parseInt(pct.getAttribute('val'), 10) };
      const pts = hijo(ln, 'spcPts');
      if (pts) return { pts: parseInt(pts.getAttribute('val'), 10) };
    }
    return null;
  }

  // --- Texto: autoajuste, interlineado y viñetas ----------------------------
  function prepararTexto(sp, contexto, ctx, cambios) {
    const txBody = hijo(sp, 'txBody');
    if (!txBody) return;
    const doc = sp.ownerDocument;
    const estilos = estilosDeTexto(sp, contexto, ctx);
    const fuentes = fuentesDelTema(contexto.tema);
    const auto = ruta(txBody, 'bodyPr', 'normAutofit');
    const escala = auto && auto.hasAttribute('fontScale') ? parseInt(auto.getAttribute('fontScale'), 10) / 100000 : 1;
    const reduccion = auto && auto.hasAttribute('lnSpcReduction') ? parseInt(auto.getAttribute('lnSpcReduction'), 10) : 0;
    if (auto && (escala !== 1 || reduccion)) cambios.autoajuste++;
    // Estilo de la forma (fontRef): color y fuente de la letra. La librería no lo usa.
    const colorEstilo = ruta(sp, 'style', 'fontRef');
    const tieneColor = (el) => !!el && ['solidFill', 'gradFill', 'noFill', 'blipFill', 'pattFill', 'grpFill'].some((n) => hijo(el, n));
    // spcFirstLastPara ausente o 0: PowerPoint no aplica el espacio antes del
    // primer párrafo ni el después del último (la librería sí, y corría el
    // texto centrado hacia arriba).
    const bodyPr = hijo(txBody, 'bodyPr');
    const sinEspacioBorde = !bodyPr || !['1', 'true'].includes(bodyPr.getAttribute('spcFirstLastPara'));
    const parrafos = hijos(txBody, 'p');

    // La librería no sigue bien la herencia por niveles (lvl): tamaño, fuente,
    // color e interlineado se escriben en cada trozo y párrafo, ya resueltos.
    // Orden para la letra: el trozo → la lista de la forma → el estilo de la
    // forma (fontRef) → los marcadores de los que hereda → el patrón.
    const listaPropia = ruta(txBody, 'lstStyle');
    const fuentesDeLista = [listaPropia, 'estilo', ...estilos.filter((e) => e !== listaPropia)];
    function deLaLista(nivel, buscar) {
      for (const e of fuentesDeLista) {
        if (e === 'estilo') {
          const v = buscar(null, colorEstilo);
          if (v) return v;
          continue;
        }
        const v = e && buscar(hijo(nivelDe(e, nivel), 'defRPr'), null);
        if (v) return v;
      }
      return null;
    }

    for (const p of parrafos) {
      let pPr = hijo(p, 'pPr');
      const nivel = pPr && pPr.hasAttribute('lvl') ? parseInt(pPr.getAttribute('lvl'), 10) : 0;
      const trozos = [...hijos(p, 'r'), ...hijos(p, 'fld')].map((r) => hijo(r, 'rPr') || r.insertBefore(crear(doc, 'a', 'rPr'), r.firstChild));
      const fin = hijo(p, 'endParaRPr');
      const todosLosTrozos = [...trozos, ...(fin ? [fin] : [])];
      let mayor = 0;
      for (const rPr of todosLosTrozos) {
        // Tamaño, con el autoajuste ya aplicado.
        const tam = Math.round(tamanoEfectivo(rPr, estilos, nivel) * escala);
        rPr.setAttribute('sz', String(tam));
        if (rPr !== fin || trozos.length === 0) mayor = Math.max(mayor, tam);

        // Fuente, ya sin "+mn-lt" (la librería usa Arial si no la encuentra).
        if (!hijo(rPr, 'latin')) {
          let fuente = deLaLista(nivel, (def, ref) => {
            if (ref) return null;
            const lat = hijo(def, 'latin');
            return lat && lat.getAttribute('typeface');
          }) || (colorEstilo && colorEstilo.getAttribute('idx') === 'major' ? '+mj-lt' : '+mn-lt');
          fuente = fuentes[fuente] || fuente;
          if (!fuente.startsWith('+')) {
            const lat = crear(doc, 'a', 'latin');
            lat.setAttribute('typeface', fuente);
            const despues = ['ea', 'cs', 'sym', 'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst'].map((n) => hijo(rPr, n)).find(Boolean);
            rPr.insertBefore(lat, despues || null);
          }
        }

        // Un enlace sin color propio va con el color de enlace del tema y
        // subrayado, como en PowerPoint (la librería no lo hace).
        if (!tieneColor(rPr) && hijo(rPr, 'hlinkClick')) {
          const relleno = crear(doc, 'a', 'solidFill');
          relleno.appendChild(crear(doc, 'a', 'schemeClr')).setAttribute('val', 'hlink');
          const ln = hijo(rPr, 'ln');
          rPr.insertBefore(relleno, ln ? ln.nextSibling : rPr.firstChild);
          if (!rPr.hasAttribute('u')) rPr.setAttribute('u', 'sng');
        }
        // Color.
        if (!tieneColor(rPr)) {
          const relleno = deLaLista(nivel, (def, ref) => {
            if (ref) {
              const c = ref.firstElementChild;
              if (!c) return null;
              const s = crear(doc, 'a', 'solidFill');
              s.appendChild(c.cloneNode(true));
              return s;
            }
            const r = def && ['solidFill', 'gradFill', 'noFill'].map((n) => hijo(def, n)).find(Boolean);
            return r ? r.cloneNode(true) : null;
          });
          if (relleno) {
            // Los rellenos van después del contorno (ln) y antes de las fuentes.
            const ln = hijo(rPr, 'ln');
            rPr.insertBefore(doc.importNode(relleno, true), ln ? ln.nextSibling : rPr.firstChild);
            cambios.colorLetra++;
          }
        }
      }
      if (!mayor) mayor = Math.round(tamanoEfectivo(null, estilos, nivel) * escala);

      if (!pPr) pPr = p.insertBefore(crear(doc, 'a', 'pPr'), p.firstChild);
      // Espacio antes del primero y después del último párrafo.
      if (sinEspacioBorde && (p === parrafos[0] || p === parrafos[parrafos.length - 1])) {
        for (const [cual, es] of [['spcBef', p === parrafos[0]], ['spcAft', p === parrafos[parrafos.length - 1]]]) {
          if (!es) continue;
          for (const viejo of hijos(pPr, cual)) viejo.remove();
          const nuevo = crear(doc, 'a', cual);
          nuevo.appendChild(crear(doc, 'a', 'spcPts')).setAttribute('val', '0');
          // Orden del esquema: lnSpc, spcBef, spcAft, viñetas…
          const antes = cual === 'spcBef' ? hijo(pPr, 'lnSpc') : (hijo(pPr, 'spcBef') || hijo(pPr, 'lnSpc'));
          pPr.insertBefore(nuevo, antes ? antes.nextSibling : pPr.firstChild);
        }
      }

      // Interlineado: siempre en puntos, sobre la altura de la línea.
      const ln = interlineadoEfectivo(pPr, estilos, nivel);
      if (!ln || ln.pct != null) {
        const pct = (ln && ln.pct != null ? ln.pct : 100000) - reduccion;
        for (const viejo of hijos(pPr, 'lnSpc')) viejo.remove();
        const nuevo = crear(doc, 'a', 'lnSpc');
        const pts = crear(doc, 'a', 'spcPts');
        pts.setAttribute('val', String(Math.round(mayor * ALTO_LINEA * pct / 100000)));
        nuevo.appendChild(pts);
        pPr.insertBefore(nuevo, pPr.firstChild);
        cambios.interlineado++;
      }

    }
  }

  // Viñetas de Wingdings o Symbol: el símbolo de verdad. Se recorre la parte
  // entera porque la viñeta suele venir del patrón (txStyles) o de la lista de
  // estilos de la forma, no del párrafo.
  function convertirVinetas(doc) {
    let n = 0;
    for (const buChar of todos(doc, 'buChar')) {
      const buFont = hijo(buChar.parentNode, 'buFont');
      const tabla = buFont && VINETAS[(buFont.getAttribute('typeface') || '').toLowerCase()];
      const simbolo = tabla && tabla[buChar.getAttribute('char')];
      if (!simbolo) continue;
      buChar.setAttribute('char', simbolo);
      buFont.setAttribute('typeface', 'Segoe UI Symbol');
      n++;
    }
    return n;
  }

  // --- Rellenos --------------------------------------------------------------
  const tieneRelleno = (spPr) => !!spPr && RELLENOS.some((n) => hijo(spPr, n));
  function insertarRelleno(spPr, relleno) {
    for (const n of RELLENOS) for (const viejo of hijos(spPr, n)) viejo.remove();
    const despues = hijo(spPr, 'prstGeom') || hijo(spPr, 'custGeom') || hijo(spPr, 'xfrm');
    spPr.insertBefore(relleno, despues ? despues.nextSibling : spPr.firstChild);
  }

  // El fondo efectivo como relleno de forma: sólido o degradado; si es una
  // imagen u otra cosa, sin relleno (deja ver el fondo, que es lo mismo).
  function rellenoDelFondo(contexto, doc) {
    for (const c of contexto.cadena) {
      const bg = ruta(c.doc.documentElement, 'cSld', 'bg');
      if (!bg) continue;
      let relleno = null;
      const bgPr = hijo(bg, 'bgPr');
      if (bgPr) relleno = hijo(bgPr, 'solidFill') || hijo(bgPr, 'gradFill');
      const bgRef = hijo(bg, 'bgRef');
      if (bgRef && contexto.tema) {
        const idx = parseInt(bgRef.getAttribute('idx'), 10);
        const lista = idx >= 1001 ? todos(contexto.tema.doc, 'bgFillStyleLst')[0] : todos(contexto.tema.doc, 'fillStyleLst')[0];
        const estilo = lista && [...lista.children][idx >= 1001 ? idx - 1001 : idx - 1];
        const color = [...bgRef.children][0];
        if (estilo && estilo.localName === 'solidFill' && color) {
          relleno = crear(doc, 'a', 'solidFill');
          relleno.appendChild(doc.importNode(color, true));
          return relleno;
        }
      }
      return relleno ? doc.importNode(relleno, true) : crear(doc, 'a', 'noFill');
    }
    return crear(doc, 'a', 'noFill');
  }

  function prepararRelleno(sp, contexto, cambios) {
    const doc = sp.ownerDocument;
    const spPr = hijo(sp, 'spPr');
    if (!spPr) return;
    if (sp.getAttribute('useBgFill') === '1' || sp.getAttribute('useBgFill') === 'true') {
      insertarRelleno(spPr, rellenoDelFondo(contexto, doc));
      sp.removeAttribute('useBgFill');
      const estilo = hijo(sp, 'style');
      // Su estilo trae fillRef accent1, que es lo que pintaba la librería.
      if (estilo && !(hijo(sp, 'txBody') && todos(hijo(sp, 'txBody'), 't').some((t) => t.textContent))) estilo.remove();
      cambios.fondo++;
      return;
    }
    if (tieneRelleno(spPr) || hijo(sp, 'style')) return;
    // Un marcador puede heredar el relleno de su diseño o patrón.
    if (herencia(sp, contexto).some((f) => tieneRelleno(hijo(f, 'spPr')) || hijo(f, 'style'))) return;
    insertarRelleno(spPr, crear(doc, 'a', 'noFill'));
    cambios.sinRelleno++;
  }

  // El estilo de la forma (p:style: fillRef, lnRef) como relleno y línea
  // explícitos: la librería no lo aplica y la forma quedaba sin relleno ni
  // borde. El estilo del tema trae "phClr", que es el color de la referencia.
  function conColorDeReferencia(doc, plantilla, ref) {
    const copia = doc.importNode(plantilla, true);
    const color = ref.firstElementChild;
    for (const ph of todos(copia, 'schemeClr').filter((c) => c.getAttribute('val') === 'phClr')) {
      if (!color) { ph.setAttribute('val', 'tx1'); continue; }
      const nuevo = color.cloneNode(true);
      for (const m of [...ph.children]) nuevo.appendChild(m);
      ph.parentNode.replaceChild(nuevo, ph);
    }
    return copia;
  }
  function estiloDelTema(contexto, lista, idx) {
    const el = contexto.tema && todos(contexto.tema.doc, lista)[0];
    return el ? [...el.children][idx - 1] || null : null;
  }
  function prepararEstilo(sp, contexto, cambios) {
    const estilo = hijo(sp, 'style');
    const spPr = hijo(sp, 'spPr');
    if (!estilo || !spPr) return;
    const doc = sp.ownerDocument;
    const fillRef = hijo(estilo, 'fillRef');
    if (fillRef && !tieneRelleno(spPr)) {
      const idx = parseInt(fillRef.getAttribute('idx'), 10) || 0;
      const plantilla = idx >= 1001 ? estiloDelTema(contexto, 'bgFillStyleLst', idx - 1000) : idx > 0 ? estiloDelTema(contexto, 'fillStyleLst', idx) : null;
      insertarRelleno(spPr, plantilla ? conColorDeReferencia(doc, plantilla, fillRef) : crear(doc, 'a', 'noFill'));
      cambios.estilos++;
    }
    const lnRef = hijo(estilo, 'lnRef');
    const idxLn = lnRef ? parseInt(lnRef.getAttribute('idx'), 10) || 0 : 0;
    let ln = hijo(spPr, 'ln');
    const lineaDelTema = idxLn > 0 && estiloDelTema(contexto, 'lnStyleLst', idxLn);
    if (lnRef && !ln) {
      ln = lineaDelTema ? conColorDeReferencia(doc, lineaDelTema, lnRef) : crear(doc, 'a', 'ln');
      if (!lineaDelTema) ln.appendChild(crear(doc, 'a', 'noFill'));
      const antes = ['effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst'].map((n) => hijo(spPr, n)).find(Boolean);
      spPr.insertBefore(ln, antes || null);
      cambios.estilos++;
    } else if (ln && lineaDelTema && !RELLENOS.some((n) => hijo(ln, n))) {
      // Tiene línea propia (grosor, guiones) pero el color viene del estilo.
      const relleno = RELLENOS.map((n) => hijo(lineaDelTema, n)).find(Boolean);
      if (relleno) {
        const copia = conColorDeReferencia(doc, relleno, lnRef);
        ln.insertBefore(copia, ln.firstChild);
        if (!ln.hasAttribute('w') && lineaDelTema.hasAttribute('w')) ln.setAttribute('w', lineaDelTema.getAttribute('w'));
        cambios.estilos++;
      }
    }
  }

  // Grupos: la librería no transforma las coordenadas de lo que hay adentro
  // (la primera forma de un grupo salía en la esquina). Se "desagrupan":
  // cada forma pasa al nivel de arriba con su posición ya calculada. Un grupo
  // girado o volteado se deja como está.
  function desagrupar(doc, cambios) {
    for (const grupo of todos(doc, 'grpSp')) {
      if (grupo.namespaceURI !== NS.p || !grupo.parentNode) continue;
      const xfrm = ruta(grupo, 'grpSpPr', 'xfrm');
      if (!xfrm) continue;
      if (xfrm.getAttribute('rot') && xfrm.getAttribute('rot') !== '0') continue;
      if (['1', 'true'].includes(xfrm.getAttribute('flipH')) || ['1', 'true'].includes(xfrm.getAttribute('flipV'))) continue;
      const n = (nombre, a) => { const e = hijo(xfrm, nombre); return e ? parseInt(e.getAttribute(a), 10) || 0 : 0; };
      const ex = n('chExt', 'cx') ? n('ext', 'cx') / n('chExt', 'cx') : 1;
      const ey = n('chExt', 'cy') ? n('ext', 'cy') / n('chExt', 'cy') : 1;
      const t = { x: n('off', 'x') - n('chOff', 'x') * ex, y: n('off', 'y') - n('chOff', 'y') * ey, ex, ey };
      const contenido = [...grupo.children].filter((el) => !['nvGrpSpPr', 'grpSpPr', 'extLst'].includes(el.localName))
        .map((el) => ({ el, x: el.localName === 'grpSp' ? ruta(el, 'grpSpPr', 'xfrm')
          : el.localName === 'graphicFrame' ? hijo(el, 'xfrm') : ruta(el, 'spPr', 'xfrm') }));
      // Algo sin posición propia (mc:AlternateContent, contentPart…) quedaría
      // en la esquina: ese grupo se deja entero.
      if (contenido.some((c) => !c.x)) continue;
      for (const { el, x } of contenido) {
        desplazar(x, t);
        grupo.parentNode.insertBefore(el, grupo);
      }
      grupo.remove();
      cambios.grupos++;
    }
  }

  // Una forma rellena con una imagen y sin texto (los íconos del SmartArt) la
  // librería la pinta de negro; como imagen (p:pic) la dibuja bien.
  function formaComoImagen(sp) {
    const spPr = hijo(sp, 'spPr');
    const blip = spPr && hijo(spPr, 'blipFill');
    const txBody = hijo(sp, 'txBody');
    if (!blip || (txBody && todos(txBody, 't').some((t) => t.textContent))) return null;
    const doc = sp.ownerDocument;
    const pic = crear(doc, 'p', 'pic');
    const nv = pic.appendChild(crear(doc, 'p', 'nvPicPr'));
    const cNvPr = ruta(sp, 'nvSpPr', 'cNvPr');
    nv.appendChild(cNvPr ? doc.importNode(cNvPr, true) : crear(doc, 'p', 'cNvPr'));
    nv.appendChild(crear(doc, 'p', 'cNvPicPr'));
    nv.appendChild(crear(doc, 'p', 'nvPr'));
    pic.appendChild(copiarComo(doc, blip, 'p', 'blipFill'));
    blip.remove();
    pic.appendChild(copiarComo(doc, spPr, 'p', 'spPr'));
    sp.parentNode.replaceChild(pic, sp);
    return pic;
  }

  // --- SmartArt: las formas que PowerPoint ya dibujó --------------------------
  function desplazar(xfrm, t) {
    if (!xfrm) return;
    const off = hijo(xfrm, 'off');
    const ext = hijo(xfrm, 'ext');
    if (off) {
      off.setAttribute('x', String(Math.round(t.x + parseInt(off.getAttribute('x'), 10) * t.ex)));
      off.setAttribute('y', String(Math.round(t.y + parseInt(off.getAttribute('y'), 10) * t.ey)));
    }
    if (ext) {
      ext.setAttribute('cx', String(Math.round(parseInt(ext.getAttribute('cx'), 10) * t.ex)));
      ext.setAttribute('cy', String(Math.round(parseInt(ext.getAttribute('cy'), 10) * t.ey)));
    }
  }
  function nuevaForma(doc, id, nombre) {
    const sp = crear(doc, 'p', 'sp');
    const nv = sp.appendChild(crear(doc, 'p', 'nvSpPr'));
    const cNvPr = nv.appendChild(crear(doc, 'p', 'cNvPr'));
    cNvPr.setAttribute('id', String(id));
    cNvPr.setAttribute('name', nombre);
    nv.appendChild(crear(doc, 'p', 'cNvSpPr'));
    nv.appendChild(crear(doc, 'p', 'nvPr'));
    return sp;
  }
  function copiarComo(doc, origen, ns, nombre) {
    const nuevo = crear(doc, ns, nombre);
    for (const h of [...origen.childNodes]) nuevo.appendChild(doc.importNode(h, true));
    for (const at of [...origen.attributes]) nuevo.setAttribute(at.name, at.value);
    return nuevo;
  }
  // Convierte las formas del dibujo del diagrama en formas de la diapositiva,
  // ya ubicadas (sin grupos: la librería no transforma bien los grupos).
  function formasDelDiagrama(doc, contenedor, t, sigId) {
    const salida = [];
    for (const el of [...contenedor.children]) {
      if (el.localName === 'grpSp') {
        const x = ruta(el, 'grpSpPr', 'xfrm');
        let tt = t;
        if (x && hijo(x, 'chExt')) {
          const off = hijo(x, 'off'), ext = hijo(x, 'ext'), chOff = hijo(x, 'chOff'), chExt = hijo(x, 'chExt');
          const n = (e, a) => parseInt(e.getAttribute(a), 10) || 0;
          const ex = n(chExt, 'cx') ? n(ext, 'cx') / n(chExt, 'cx') : 1;
          const ey = n(chExt, 'cy') ? n(ext, 'cy') / n(chExt, 'cy') : 1;
          tt = { x: t.x + (n(off, 'x') - n(chOff, 'x') * ex) * t.ex, y: t.y + (n(off, 'y') - n(chOff, 'y') * ey) * t.ey, ex: t.ex * ex, ey: t.ey * ey };
        }
        salida.push(...formasDelDiagrama(doc, el, tt, sigId));
        continue;
      }
      if (el.localName !== 'sp') continue;
      const spPr = hijo(el, 'spPr'), estilo = hijo(el, 'style'), txBody = hijo(el, 'txBody'), txXfrm = hijo(el, 'txXfrm');
      const forma = nuevaForma(doc, sigId(), 'SmartArt');
      const spPrNuevo = forma.appendChild(spPr ? copiarComo(doc, spPr, 'p', 'spPr') : crear(doc, 'p', 'spPr'));
      desplazar(hijo(spPrNuevo, 'xfrm'), t);
      if (estilo) forma.appendChild(copiarComo(doc, estilo, 'p', 'style'));
      const conTexto = txBody && todos(txBody, 't').some((x) => x.textContent);
      if (conTexto && !txXfrm) forma.appendChild(copiarComo(doc, txBody, 'p', 'txBody'));
      salida.push(forma);
      if (conTexto && txXfrm) {
        // El texto va en su propio rectángulo (txXfrm), no en el de la forma.
        const texto = nuevaForma(doc, sigId(), 'SmartArt texto');
        const pr = texto.appendChild(crear(doc, 'p', 'spPr'));
        const xfrm = pr.appendChild(copiarComo(doc, txXfrm, 'a', 'xfrm'));
        desplazar(xfrm, t);
        const geom = pr.appendChild(crear(doc, 'a', 'prstGeom'));
        geom.setAttribute('prst', 'rect');
        geom.appendChild(crear(doc, 'a', 'avLst'));
        pr.appendChild(crear(doc, 'a', 'noFill'));
        pr.appendChild(crear(doc, 'a', 'ln')).appendChild(crear(doc, 'a', 'noFill'));
        if (estilo) {
          // Solo la letra (fontRef); sin relleno ni línea.
          const e = texto.appendChild(copiarComo(doc, estilo, 'p', 'style'));
          for (const ref of ['lnRef', 'fillRef', 'effectRef']) { const r = hijo(e, ref); if (r) r.setAttribute('idx', '0'); }
        }
        texto.appendChild(copiarComo(doc, txBody, 'p', 'txBody'));
        salida.push(texto);
      }
    }
    return salida;
  }

  async function prepararSmartArt(zip, parte, docDiap, rels, cambios) {
    const marcos = todos(docDiap, 'graphicFrame').filter((g) => {
      const datos = ruta(g, 'graphic', 'graphicData');
      return datos && /\/diagram$/.test(datos.getAttribute('uri') || '');
    });
    if (!marcos.length) return;
    let maxId = Math.max(0, ...todos(docDiap, 'cNvPr').map((c) => parseInt(c.getAttribute('id'), 10) || 0));
    const sigId = () => ++maxId;
    let nRel = 0;
    const dibujos = rels.lista.filter((r) => r.tipo === TIPO_DIBUJO_DIAGRAMA);
    for (const marco of marcos) {
      const relIds = todos(marco, 'relIds')[0];
      const idDatos = relIds && relIds.getAttributeNS(NS.r, 'dm');
      let dibujo = null;
      // El dibujo se enlaza desde la parte de datos (dsp:dataModelExt relId).
      const relDatos = idDatos && rels.porId(idDatos);
      if (relDatos && relDatos.tipo === TIPO_DATOS_DIAGRAMA && zip.file(relDatos.destino)) {
        const datos = parsear(await zip.file(relDatos.destino).async('string'));
        const ext = todos(datos, 'dataModelExt')[0];
        const rel = ext && rels.porId(ext.getAttribute('relId'));
        if (rel) dibujo = rel.destino;
      }
      if (!dibujo && dibujos.length === 1 && marcos.length === 1) dibujo = dibujos[0].destino;
      if (!dibujo || !zip.file(dibujo)) continue;

      const docDibujo = parsear(await zip.file(dibujo).async('string'));
      const arbol = todos(docDibujo, 'spTree')[0];
      const off = ruta(marco, 'xfrm', 'off');
      if (!arbol || !off) continue;
      const t = { x: parseInt(off.getAttribute('x'), 10) || 0, y: parseInt(off.getAttribute('y'), 10) || 0, ex: 1, ey: 1 };
      const formas = formasDelDiagrama(docDiap, arbol, t, sigId);

      // Las imágenes del dibujo apuntan a sus propias relaciones: se copian a
      // las de la diapositiva con un id nuevo.
      const relsDibujo = await leerRels(zip, dibujo);
      for (const forma of formas) {
        for (const el of todos(forma, '*')) {
          for (const at of ['embed', 'link']) {
            const id = el.getAttributeNS(NS.r, at);
            const rel = id && relsDibujo.porId(id);
            if (!rel || !rel.destino) continue;
            const nuevoId = `rIdMiniCookies${++nRel}`;
            const r = rels.doc.createElementNS(NS.rel, 'Relationship');
            r.setAttribute('Id', nuevoId);
            r.setAttribute('Type', rel.tipo);
            r.setAttribute('Target', rutaRelativa(parte, rel.destino));
            rels.doc.documentElement.appendChild(r);
            rels.lista.push({ id: nuevoId, tipo: rel.tipo, destino: rel.destino });
            el.setAttributeNS(NS.r, `r:${at}`, nuevoId);
          }
        }
      }
      for (const forma of formas) marco.parentNode.insertBefore(forma, marco);
      marco.remove();
      cambios.smartArt++;
    }
  }

  // --- Colores: los modificadores (tint, shade, lumMod…) ya calculados ---------
  // PptxViewJS no aplica tint ni shade: una barra en accent1 al 40 % salía
  // del color pleno. Se resuelve el color final y se escribe como srgbClr
  // (con su alpha, que la librería sí entiende). tint y shade se aplican en
  // RGB lineal y lumMod/lumOff/satMod en HSL, como PowerPoint (comparado con
  // sus píxeles: accent1 #156082 con tint 40 % da #CCD2D8 en PowerPoint y
  // #CCD4DB acá).
  const aLineal = (c) => { c /= 255; return c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4); };
  const aSrgb = (c) => { const v = c <= 0.0031308 ? c * 12.92 : 1.055 * Math.pow(c, 1 / 2.4) - 0.055; return Math.max(0, Math.min(255, Math.round(v * 255))); };
  function aHsl([r, g, b]) {
    r /= 255; g /= 255; b /= 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b), l = (max + min) / 2;
    if (max === min) return [0, 0, l];
    const d = max - min, s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    return [h * 60, s, l];
  }
  function deHsl([h, s, l]) {
    if (s === 0) return [l, l, l].map((v) => Math.round(v * 255));
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    const canal = (t) => { t = ((t % 1) + 1) % 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
    return [canal(h / 360 + 1 / 3), canal(h / 360), canal(h / 360 - 1 / 3)].map((v) => Math.round(v * 255));
  }
  const MODIFICADORES = new Set(['tint', 'shade', 'lumMod', 'lumOff', 'satMod', 'satOff', 'hueMod', 'hueOff', 'alpha', 'alphaMod', 'alphaOff', 'comp', 'inv', 'gray']);
  const hex = (rgb) => rgb.map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();

  function mapaDeColores(contexto) {
    const tema = {};
    const esquema = contexto.tema && todos(contexto.tema.doc, 'clrScheme')[0];
    for (const el of esquema ? [...esquema.children] : []) {
      const c = el.firstElementChild;
      if (!c) continue;
      const v = c.localName === 'sysClr' ? c.getAttribute('lastClr') : c.getAttribute('val');
      if (v && /^[0-9a-f]{6}$/i.test(v)) tema[el.localName] = v;
    }
    const mapa = {};
    const clrMap = contexto.patron && ruta(contexto.patron.doc.documentElement, 'clrMap');
    for (const at of clrMap ? [...clrMap.attributes] : []) mapa[at.name] = at.value;
    // Un diseño o una diapositiva pueden cambiar la correspondencia.
    for (const c of contexto.cadena.slice(0, -1).reverse()) {
      const ovr = ruta(c.doc.documentElement, 'clrMapOvr', 'overrideClrMapping');
      for (const at of ovr ? [...ovr.attributes] : []) mapa[at.name] = at.value;
    }
    return (nombre) => tema[mapa[nombre] || nombre] || tema[nombre] || null;
  }

  function resolverColores(doc, contexto) {
    const colorDelTema = mapaDeColores(contexto);
    let n = 0;
    for (const el of [...todos(doc, 'schemeClr'), ...todos(doc, 'srgbClr'), ...todos(doc, 'sysClr')]) {
      if (el.namespaceURI !== NS.a) continue;
      const mods = [...el.children];
      if (!mods.length || mods.some((m) => !MODIFICADORES.has(m.localName))) continue;
      if (mods.every((m) => m.localName === 'alpha' || (/(Off)$/.test(m.localName) && m.getAttribute('val') === '0'))) continue;
      const base = el.localName === 'schemeClr' ? colorDelTema(el.getAttribute('val'))
        : el.localName === 'sysClr' ? el.getAttribute('lastClr') : el.getAttribute('val');
      if (!base || !/^[0-9a-f]{6}$/i.test(base)) continue;
      let rgb = [0, 2, 4].map((i) => parseInt(base.slice(i, i + 2), 16));
      let alfa = 100000;
      for (const m of mods) {
        const v = parseInt(m.getAttribute('val'), 10) / 100000;
        switch (m.localName) {
          case 'tint': rgb = rgb.map((c) => aSrgb(aLineal(c) * v + (1 - v))); break;
          case 'shade': rgb = rgb.map((c) => aSrgb(aLineal(c) * v)); break;
          case 'lumMod': { const h = aHsl(rgb); h[2] = Math.min(1, h[2] * v); rgb = deHsl(h); break; }
          case 'lumOff': { const h = aHsl(rgb); h[2] = Math.max(0, Math.min(1, h[2] + v)); rgb = deHsl(h); break; }
          case 'satMod': { const h = aHsl(rgb); h[1] = Math.max(0, Math.min(1, h[1] * v)); rgb = deHsl(h); break; }
          case 'satOff': { const h = aHsl(rgb); h[1] = Math.max(0, Math.min(1, h[1] + v)); rgb = deHsl(h); break; }
          case 'hueMod': { const h = aHsl(rgb); h[0] = (h[0] * v) % 360; rgb = deHsl(h); break; }
          case 'hueOff': { const h = aHsl(rgb); h[0] = (h[0] + parseInt(m.getAttribute('val'), 10) / 60000 + 360) % 360; rgb = deHsl(h); break; }
          case 'comp': { const h = aHsl(rgb); h[0] = (h[0] + 180) % 360; rgb = deHsl(h); break; }
          case 'inv': rgb = rgb.map((c) => 255 - c); break;
          case 'gray': { const y = Math.round(0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]); rgb = [y, y, y]; break; }
          case 'alpha': alfa = v * 100000; break;
          case 'alphaMod': alfa *= v; break;
          case 'alphaOff': alfa += v * 100000; break;
        }
      }
      const nuevo = crear(doc, 'a', 'srgbClr');
      nuevo.setAttribute('val', hex(rgb));
      if (alfa < 100000) {
        const a = nuevo.appendChild(crear(doc, 'a', 'alpha'));
        a.setAttribute('val', String(Math.max(0, Math.round(alfa))));
      }
      el.parentNode.replaceChild(nuevo, el);
      n++;
    }
    return n;
  }

  // --- Fuentes del tema --------------------------------------------------------
  function fuentesDelTema(tema) {
    const mapa = {};
    if (!tema) return mapa;
    for (const [clave, nombre] of [['mj', 'majorFont'], ['mn', 'minorFont']]) {
      const f = todos(tema.doc, nombre)[0];
      if (!f) continue;
      const latin = (hijo(f, 'latin') || { getAttribute: () => '' }).getAttribute('typeface') || '';
      for (const guion of ['lt', 'ea', 'cs']) {
        const el = hijo(f, guion === 'lt' ? 'latin' : guion);
        mapa[`+${clave}-${guion}`] = (el && el.getAttribute('typeface')) || latin;
      }
    }
    return mapa;
  }
  function resolverFuentes(doc, mapa) {
    let n = 0;
    for (const el of [...todos(doc, 'latin'), ...todos(doc, 'ea'), ...todos(doc, 'cs')]) {
      const tf = el.getAttribute('typeface');
      if (tf && tf.startsWith('+') && mapa[tf]) { el.setAttribute('typeface', mapa[tf]); n++; }
    }
    return n;
  }

  // --- Todo junto ----------------------------------------------------------------
  async function preparar(archivo) {
    const zip = await JSZip.loadAsync(archivo);
    const ctx = await crearContexto(zip);
    const cambios = { fondo: 0, sinRelleno: 0, smartArt: 0, autoajuste: 0, interlineado: 0, vinetas: 0, fuentes: 0, colores: 0, colorLetra: 0, imagenes: 0, estilos: 0, grupos: 0 };
    const partes = Object.keys(zip.files).filter((n) => /^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(n));
    for (const parte of partes) {
      const contexto = await cadenaDeParte(ctx, parte);
      const { doc, rels } = contexto.cadena[0];
      // Una parte con XML roto se deja como está (reescribirla la dañaría más).
      if (todos(doc, 'parsererror').length) continue;
      const esDiapositiva = parte.startsWith('ppt/slides/');
      if (esDiapositiva) await prepararSmartArt(zip, parte, doc, rels, cambios);
      desagrupar(doc, cambios);
      for (const sp of [...todos(doc, 'sp'), ...todos(doc, 'cxnSp')]) {
        if (sp.namespaceURI !== NS.p) continue;
        // En diseños y patrones, los marcadores solo dan formato: no se dibujan.
        if (!esDiapositiva && datosMarcador(sp)) continue;
        if (formaComoImagen(sp)) { cambios.imagenes++; continue; }
        prepararRelleno(sp, contexto, cambios);
        prepararEstilo(sp, contexto, cambios);
        prepararTexto(sp, contexto, ctx, cambios);
      }
      cambios.fuentes += resolverFuentes(doc, fuentesDelTema(contexto.tema));
      cambios.colores += resolverColores(doc, contexto);
      cambios.vinetas += convertirVinetas(doc);
    }
    // Se escribe después: los diseños y patrones se leen ya modificados en
    // memoria, pero la cadena de herencia necesita los originales intactos
    // hasta haber pasado por todas las diapositivas. (Los marcadores de los
    // diseños y patrones no se tocan, así que no cambia nada heredado.)
    for (const parte of partes) {
      const { doc, rels } = await ctx.doc(parte);
      if (todos(doc, 'parsererror').length) continue;
      zip.file(parte, new XMLSerializer().serializeToString(doc));
      if (rels.doc) zip.file(rutaRels(parte), new XMLSerializer().serializeToString(rels.doc));
    }
    const buffer = await zip.generateAsync({ type: 'arraybuffer' });
    return { buffer, cambios };
  }

  return { preparar };
})();

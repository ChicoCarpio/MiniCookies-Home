/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Traductor de documentos · Word (.docx). Edita el documento ORIGINAL en el lugar: solo cambia el texto de los tramos (w:t) y deja todo lo
// demás exactamente como estaba (estilos, numeración, tablas, imágenes, campos, marcadores, secciones). No usa el escritor de Word del sitio,
// que arma documentos nuevos y perdería el formato.
//
//   const { bytes, estadistica } = await TraductorDocx.traducir(datos, { traducir, destino, progreso })
//        datos: Uint8Array | ArrayBuffer | Blob del .docx; traducir/destino/progreso como en TraductorDocumentos
//
// Qué se traduce: el cuerpo, los encabezados y pies, las notas al pie y finales, y los cuadros de texto (sus párrafos, uno por uno).
// Qué no: las fórmulas (m:oMath), los códigos de campo (w:instrText), el texto borrado con control de cambios y los comentarios.
// Cada párrafo (w:p) es una unidad; sus tramos (w:r) dan el formato (sus propiedades w:rPr y el enlace w:hyperlink que los envuelve).
// Tabulaciones, saltos, imágenes, notas y campos dentro del párrafo son BARRERAS: parten el párrafo en trozos que se traducen por separado
// y quedan en su lugar. Al volver, el texto de cada formato va al siguiente tramo de ese formato, en orden; si la traducción cambió el orden
// de los formatos, se agrega un tramo nuevo copiado del que tiene ese formato. Los tramos que quedan sin texto quedan vacíos.
// La propiedad de idioma (w:lang) de los tramos traducidos pasa al idioma de destino (corrector ortográfico de Word).
//
// Requiere: zip-propio.js, xml-editar.js y traductor-documentos.js.
(function (raiz) {
  'use strict';

  const X = () => raiz.XmlEditar;
  const IDIOMA_WORD = { es: 'es-ES', en: 'en-US', pt: 'pt-BR', fr: 'fr-FR', de: 'de-DE', it: 'it-IT', zh: 'zh-CN', ja: 'ja-JP', ko: 'ko-KR' };
  const ASIATICOS = new Set(['zh', 'ja', 'ko']);
  // Dentro de un tramo: lo que no es texto ni propiedades parte el párrafo (queda en su lugar).
  const IGNORAR_EN_TRAMO = new Set(['rPr', 'lastRenderedPageBreak', 'softHyphen']);
  // Ramas que no se traducen.
  const NO_ENTRAR = new Set(['del', 'moveFrom', 'oMath', 'oMathPara', 'instrText', 'delInstrText', 'rPrChange', 'pPrChange']);

  const partes = (zip) => Object.keys(zip.files).filter((r) => /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(r)).sort();

  // Clave de formato: las propiedades sin lo que no cambia el aspecto (idioma, revisión ortográfica), más el enlace que lo envuelve.
  const claveRPr = (rpr) => rpr.replace(/<(\w+:)?(lang|noProof|proofErr)\b[^>]*\/>/g, '').replace(/\s+/g, ' ');

  // Recorre un párrafo: sus tramos propios en orden (no los de los párrafos anidados de un cuadro de texto, que son otra unidad).
  function tramosDelParrafo(p) {
    const r = [];
    (function ir(n, enlace) {
      for (const h of n.hijos) {
        if (NO_ENTRAR.has(h.local) || h.local === 'p' || h.local === 'txbxContent') continue;
        if (h.local === 'r') { r.push({ nodo: h, enlace }); continue; }
        ir(h, h.local === 'hyperlink' ? h : enlace);
      }
    })(p, null);
    return r;
  }

  function todosLosParrafos(raizXml) {
    const r = [];
    (function ir(n) {
      for (const h of n.hijos) {
        if (NO_ENTRAR.has(h.local)) continue;
        if (h.local === 'p') r.push(h);
        ir(h);
      }
    })(raizXml);
    return r;
  }

  // Analiza una parte: devuelve sus trozos traducibles. trozo = { piezas: [{ tramo, k }] } donde tramo.hijos[k] es un texto.
  function analizarParte(xml) {
    const ed = X();
    const raizXml = ed.analizar(xml);
    const tramos = [];       // { nodo, enlace, rpr, formato, hijos: [{ tipo: 't'|'barrera'|'nada', texto?, nodo }] }
    const trozos = [];
    for (const p of todosLosParrafos(raizXml)) {
      let trozo = { piezas: [] };
      const cerrar = () => { if (trozo.piezas.some((x) => x.tramo.hijos[x.k].texto)) trozos.push(trozo); trozo = { piezas: [] }; };
      for (const { nodo, enlace } of tramosDelParrafo(p)) {
        const rprNodo = ed.hijo(nodo, 'rPr');
        const rpr = rprNodo ? xml.slice(rprNodo.ini, rprNodo.fin) : '';
        const t = { nodo, enlace, rpr, formato: claveRPr(rpr) + '|' + (enlace ? xml.slice(enlace.ini, enlace.finApertura) : ''), hijos: [] };
        for (const h of nodo.hijos) {
          if (h.local === 't') {
            t.hijos.push({ tipo: 't', texto: ed.decodificar(ed.interior(xml, h)), nodo: h });
            trozo.piezas.push({ tramo: t, k: t.hijos.length - 1 });
          } else if (h.local === 'noBreakHyphen') {
            t.hijos.push({ tipo: 't', texto: '‑', nodo: h });
            trozo.piezas.push({ tramo: t, k: t.hijos.length - 1 });
          } else if (IGNORAR_EN_TRAMO.has(h.local)) {
            t.hijos.push({ tipo: h.local === 'rPr' ? 'rpr' : 'nada', nodo: h });
          } else {
            t.hijos.push({ tipo: 'barrera', nodo: h });
            cerrar();
          }
        }
        tramos.push(t);
      }
      cerrar();
    }
    return { raizXml, tramos, trozos };
  }

  // Reparte los tramos traducidos de un trozo entre sus piezas de texto. Devuelve las piezas nuevas que hay que agregar: [{ despues, formato, texto }].
  function repartir(trozo, traducidos) {
    const piezas = trozo.piezas;
    for (const p of piezas) p.tramo.hijos[p.k].nuevo = '';
    const nuevas = [];
    let pos = 0;              // la pieza donde se está escribiendo
    let desde = 0;            // la primera pieza que todavía se puede usar (después de un tramo nuevo, solo las que siguen)
    for (const t of traducidos) {
      if (!t.texto) continue;
      let j = desde;
      while (j < piezas.length && piezas[j].tramo.formato !== t.formato) j++;
      if (j < piezas.length) {
        piezas[j].tramo.hijos[piezas[j].k].nuevo += t.texto;
        pos = desde = j;      // el mismo tramo puede recibir más texto de su formato si viene seguido
        continue;
      }
      // No queda un tramo de ese formato más adelante: uno nuevo, copia del primero con ese formato, después de la pieza actual.
      // Lo que venga después ya no puede volver a esa pieza (quedaría antes del tramo nuevo).
      const modelo = piezas.find((p) => p.tramo.formato === t.formato) || piezas[0];
      const ancla = piezas[pos];
      const u = nuevas[nuevas.length - 1];
      if (u && u.ancla === ancla && u.modelo.tramo.formato === t.formato) u.texto += t.texto;
      else nuevas.push({ ancla, modelo, texto: t.texto });
      desde = pos + 1;
    }
    return nuevas;
  }

  const prefijo = (nombre) => (nombre.includes(':') ? nombre.slice(0, nombre.indexOf(':') + 1) : '');

  function cambiarIdioma(rpr, destino) {
    if (!rpr || !destino) return rpr;
    const cod = IDIOMA_WORD[destino];
    return rpr.replace(/<(\w+:)?lang\b[^>]*\/>/, (etiqueta, pre = '') => {
      let e = X().ponerAtributo(etiqueta, `${pre}val`, cod);
      if (ASIATICOS.has(destino)) e = X().ponerAtributo(e, `${pre}eastAsia`, cod);
      return e;
    });
  }

  const textoXml = (pre, s) => (s ? `<${pre}t xml:space="preserve">${X().esc(s)}</${pre}t>` : '');

  // Un tramo nuevo (cuando la traducción cambió el orden de los formatos): las propiedades del modelo, con el idioma nuevo, y el texto.
  const tramoNuevo = (t, destino, texto) => `<${t.nodo.nombre}>${cambiarIdioma(t.rpr, destino)}${textoXml(prefijo(t.nodo.nombre), texto)}</${t.nodo.nombre}>`;

  // Ediciones quirúrgicas de un tramo existente: cada w:t por su texto nuevo (o nada si quedó vacío) y la etiqueta de idioma.
  // No se reescribe el tramo entero: puede tener adentro un cuadro de texto cuyos párrafos tienen sus propias ediciones.
  function edicionesDelTramo(xml, t, destino) {
    const ed = [];
    const pre = prefijo(t.nodo.nombre);
    for (const h of t.hijos) {
      if (h.tipo === 't' && h.nuevo !== undefined) ed.push({ ini: h.nodo.ini, fin: h.nodo.fin, texto: textoXml(pre, h.nuevo) });
      if (h.tipo === 'rpr') {
        const lang = X().hijo(h.nodo, 'lang');
        if (lang) {
          const etiqueta = xml.slice(lang.ini, lang.fin);
          const nueva = cambiarIdioma(etiqueta, destino);
          if (nueva !== etiqueta) ed.push({ ini: lang.ini, fin: lang.fin, texto: nueva });
        }
      }
    }
    return ed;
  }

  async function traducir(datos, { traducir: traductor, destino, progreso }) {
    const zip = await raiz.ZipPropio.loadAsync(datos);
    const analizadas = [];
    for (const ruta of partes(zip)) {
      const xml = await zip.file(ruta).async('string');
      analizadas.push({ ruta, xml, ...analizarParte(xml) });
    }
    // Todas las unidades de todas las partes en un solo pedido (el progreso es del documento entero).
    const unidades = [];
    for (const a of analizadas) for (const tr of a.trozos) unidades.push(tr.piezas.map((p) => ({ texto: p.tramo.hijos[p.k].texto, formato: p.tramo.formato })));
    const r = await raiz.TraductorDocumentos.traducirUnidades(unidades, {
      traducir: (segs, d, avisar) => traductor(segs, d, { progreso: avisar }), destino, progreso,
    });
    let u = 0;
    for (const a of analizadas) {
      const ediciones = [];
      const agregadosTrasEnlace = new Map();     // enlace -> xml a poner después de su cierre
      const agregadosTrasTramo = new Map();      // tramo -> xml a poner después
      const tocados = new Set();
      for (const tr of a.trozos) {
        const nuevas = repartir(tr, r[u++]);
        for (const p of tr.piezas) tocados.add(p.tramo);
        for (const n of nuevas) {
          const nuevoXml = tramoNuevo(n.modelo.tramo, destino, n.texto);
          const anclaT = n.ancla.tramo;
          const enlaceModelo = n.modelo.tramo.enlace;
          if (anclaT.enlace && anclaT.enlace !== enlaceModelo) {
            // El ancla está dentro de un enlace y el tramo nuevo no es de ese enlace: va después del enlace.
            const envuelto = enlaceModelo ? `${a.xml.slice(enlaceModelo.ini, enlaceModelo.finApertura)}${nuevoXml}</${enlaceModelo.nombre}>` : nuevoXml;
            agregadosTrasEnlace.set(anclaT.enlace, (agregadosTrasEnlace.get(anclaT.enlace) || '') + envuelto);
          } else {
            const envuelto = enlaceModelo && enlaceModelo !== anclaT.enlace
              ? `${a.xml.slice(enlaceModelo.ini, enlaceModelo.finApertura)}${nuevoXml}</${enlaceModelo.nombre}>` : nuevoXml;
            agregadosTrasTramo.set(anclaT, (agregadosTrasTramo.get(anclaT) || '') + envuelto);
          }
        }
      }
      for (const t of tocados) {
        ediciones.push(...edicionesDelTramo(a.xml, t, destino));
        if (agregadosTrasTramo.has(t)) ediciones.push({ ini: t.nodo.fin, fin: t.nodo.fin, texto: agregadosTrasTramo.get(t) });
      }
      for (const [enlace, x] of agregadosTrasEnlace) ediciones.push({ ini: enlace.fin, fin: enlace.fin, texto: x });
      const nuevo = X().aplicar(a.xml, ediciones);
      X().analizar(nuevo);                        // el XML tiene que seguir bien formado; si no, se corta acá con un error claro
      zip.file(a.ruta, nuevo);
    }
    const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    return { bytes, estadistica: r.estadistica };
  }

  raiz.TraductorDocx = { traducir, _nucleo: { analizarParte, repartir, claveRPr } };
})(typeof self !== 'undefined' ? self : globalThis);

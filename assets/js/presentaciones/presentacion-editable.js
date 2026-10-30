/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Presentación editable de Mini Cookies Presentaciones (fase 5 del PLAN). Cada edición es un parche del XML de las partes que toca (xml-editar):
// mover o cambiar el tamaño de una forma cambia su <a:off>/<a:ext>, editar el texto cambia solo los párrafos que cambiaron, agregar una
// diapositiva agrega su parte y su relación. Todo lo demás del paquete sale byte a byte como vino.
//
// Una edición se describe como { t: 'partes', cambios: [{ ruta, datos }] } (datos: texto, Uint8Array o null = la parte se borra). aplicar()
// la escribe y devuelve la inversa (con lo que había antes), así el historial la deshace sin volver a leer nada.
//
// Requiere zip-propio, xml-editar (del sitio) y paquete-editable (de la suite). Para dibujar, motores/pptx-modelo.js.
//
//   const pe = await PresentacionEditable.abrir(bytes)
//   pe.rutas() → rutas de las diapositivas en orden · pe.zip · pe.paquete
//   await pe.opMover(i, camino, { x, y, cx, cy, rot? })          (EMU)
//   await pe.opTexto(i, camino, ['párrafo 1', 'párrafo 2'])       (solo cambian los párrafos distintos)
//   await pe.opFormatoTexto(i, camino, { negrita, cursiva, subrayado, tamano (pt), color: 'RRGGBB' })
//   await pe.opRelleno(i, camino, 'RRGGBB' | null) · opLinea(i, camino, 'RRGGBB' | null) · opBorrar(i, camino) · opOrden(i, camino, 'frente'|'fondo')
//   await pe.opCuadroTexto(i, { x, y, cx, cy }, texto) → { op, id } · opImagen(i, bytes, tipo, { x, y, cx, cy }) → { op, id }
//   await pe.opDuplicarDiapositiva(i) · opNuevaDiapositiva(i) · opBorrarDiapositiva(i) · opMoverDiapositiva(i, j)
//   pe.aplicar(op) → inversa · await pe.generar() → Uint8Array del .pptx
//   await pe.textoDe(i, camino) → ['párrafo', …] · await pe.editable(i, camino) → motivo | null
(function () {
  const X = () => globalThis.XmlEditar;
  const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const REL_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
  const REL_LAYOUT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';
  const REL_IMAGEN = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
  const TIPO_SLIDE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
  const NV = ['nvSpPr', 'nvPicPr', 'nvGraphicFramePr', 'nvCxnSpPr', 'nvGrpSpPr'];
  const FORMAS = ['sp', 'pic', 'grpSp', 'graphicFrame', 'cxnSp'];
  const RELLENOS = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];
  const ORDEN_SPPR = ['xfrm', 'custGeom', 'prstGeom', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill', 'ln', 'effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst'];
  const MAX_EMU = 51206400 * 4;

  class ErrorPresentacion extends Error {}

  // prefijo con que el documento declara un espacio de nombres (a:, p:, r:…)
  function prefijo(xml, ns, porDefecto) {
    const m = new RegExp('xmlns:([\\w.-]+)="' + ns.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&') + '"').exec(xml);
    return m ? m[1] : porDefecto;
  }
  const idDe = (nodo) => { const nv = nodo.hijos.find((h) => NV.includes(h.local)); const c = nv && X().hijo(nv, 'cNvPr'); return c ? c.attrs.id : null; };
  const arbolDe = (raiz) => { const cSld = X().hijo(raiz, 'cSld'); return cSld && X().hijo(cSld, 'spTree'); };
  // la forma con ese camino de ids (los de los grupos que la contienen y el suyo). Si dos formas del mismo nivel comparten el id (archivos
  // de otros programas), no se sabe cuál es: null, y la forma queda sin editar en vez de cambiar la otra
  function buscar(raiz, camino) {
    let nivel = arbolDe(raiz);
    let nodo = null;
    for (const id of camino) {
      if (!nivel) return null;
      const iguales = nivel.hijos.filter((h) => FORMAS.includes(h.local) && idDe(h) === String(id));
      if (iguales.length !== 1) return null;
      nodo = iguales[0];
      nivel = nodo.local === 'grpSp' ? nodo : null;
    }
    return nodo;
  }
  const esc = (t) => X().esc(t);
  const nodoTexto = (xml, n) => xml.slice(n.ini, n.fin);
  // hijo nuevo en un elemento según un orden de nombres locales (o al final)
  function insertarHijo(xml, padre, local, nuevo, orden) {
    if (padre.iniCierre < 0) {   // <p:spPr/> → <p:spPr>nuevo</p:spPr>
      const ap = xml.slice(padre.ini, padre.fin).replace(/\s*\/>$/, '>');
      return X().aplicar(xml, [{ ini: padre.ini, fin: padre.fin, texto: ap + nuevo + '</' + padre.nombre + '>' }]);
    }
    const k = orden ? orden.indexOf(local) : -1;
    const sig = k < 0 ? null : padre.hijos.find((h) => orden.indexOf(h.local) > k);
    const pos = sig ? sig.ini : padre.iniCierre;
    return X().aplicar(xml, [{ ini: pos, fin: pos, texto: nuevo }]);
  }

  async function abrir(bytes) {
    const zip = await ZipPropio.loadAsync(bytes);
    const paquete = PaqueteEditable.desdeZip(zip);
    const leer = (r) => paquete.leer(r);
    const presXml = async () => { const t = await leer('ppt/presentation.xml'); if (t === null) throw new ErrorPresentacion('no es una presentación'); return t; };

    async function rutas() {
      const t = await presXml();
      const raiz = X().analizar(t, { profundidad: 3 });
      const lista = X().hijo(raiz, 'sldIdLst');
      const rels = await paquete.relaciones('ppt/presentation.xml');
      const porId = new Map(rels.map((r) => [r.id, r.destino]));
      return (lista ? X().hijos(lista, 'sldId') : []).map((s) => porId.get(s.attrs['r:id'] || Object.entries(s.attrs).find(([k]) => /:id$/.test(k))?.[1])).filter(Boolean);
    }
    const rutaDe = async (i) => { const r = (await rutas())[i]; if (!r) throw new ErrorPresentacion('diapositiva ' + (i + 1)); return r; };
    // una edición de una sola diapositiva: fn(xml, raiz, nodo, pref) → xml nuevo
    async function editarForma(i, camino, fn) {
      const ruta = await rutaDe(i);
      const xml = await leer(ruta);
      const raiz = X().analizar(xml);
      const nodo = buscar(raiz, camino);
      if (!nodo) throw new ErrorPresentacion('la forma no está en la diapositiva');
      const pref = { a: prefijo(xml, NS_A, 'a'), p: prefijo(xml, NS_P, 'p'), r: prefijo(xml, NS_R, 'r') };
      const nuevo = fn(xml, raiz, nodo, pref);
      if (nuevo === xml) return null;
      return { t: 'partes', cambios: [{ ruta, datos: nuevo, antes: xml }] };
    }
    async function editable(i, camino) {
      const ruta = await rutaDe(i);
      const xml = await leer(ruta);
      const raiz = X().analizar(xml);
      if (!buscar(raiz, camino)) return 'fuera';   // p. ej. dentro de mc:AlternateContent o con el id repetido: se ve, pero se conserva sin editar
      return null;
    }

    // ---- posición y tamaño ----------------------------------------------------------------------------------------
    function opMover(i, camino, caja) {
      for (const k of ['x', 'y', 'cx', 'cy']) if (!Number.isFinite(caja[k]) || Math.abs(caja[k]) > MAX_EMU) throw new ErrorPresentacion('posición fuera de rango');
      return editarForma(i, camino, (xml, raiz, nodo, pref) => {
        const x = Math.round(caja.x), y = Math.round(caja.y), cx = Math.max(0, Math.round(caja.cx)), cy = Math.max(0, Math.round(caja.cy));
        let contenedor, prefXfrm = pref.a;
        if (nodo.local === 'graphicFrame') { contenedor = nodo; prefXfrm = pref.p; }
        else contenedor = X().hijo(nodo, nodo.local === 'grpSp' ? 'grpSpPr' : 'spPr');
        if (!contenedor) throw new ErrorPresentacion('la forma no tiene propiedades');
        const xfrm = X().hijo(contenedor, 'xfrm');
        if (!xfrm) {
          if (nodo.local === 'graphicFrame' || nodo.local === 'grpSp') throw new ErrorPresentacion('forma sin posición');
          const nuevo = `<${pref.a}:xfrm${caja.rot ? ` rot="${Math.round(caja.rot)}"` : ''}><${pref.a}:off x="${x}" y="${y}"/><${pref.a}:ext cx="${cx}" cy="${cy}"/></${pref.a}:xfrm>`;
          return insertarHijo(xml, contenedor, 'xfrm', nuevo, ORDEN_SPPR);
        }
        const ed = [];
        const off = X().hijo(xfrm, 'off'), ext = X().hijo(xfrm, 'ext');
        if (off) ed.push({ ini: off.ini, fin: off.finApertura, texto: X().ponerAtributo(X().ponerAtributo(xml.slice(off.ini, off.finApertura), 'x', String(x)), 'y', String(y)) });
        if (ext) ed.push({ ini: ext.ini, fin: ext.finApertura, texto: X().ponerAtributo(X().ponerAtributo(xml.slice(ext.ini, ext.finApertura), 'cx', String(cx)), 'cy', String(cy)) });
        if (caja.rot !== undefined) ed.push({ ini: xfrm.ini, fin: xfrm.finApertura, texto: X().ponerAtributo(xml.slice(xfrm.ini, xfrm.finApertura), 'rot', caja.rot ? String(Math.round(caja.rot)) : null) });
        if (!off || !ext) throw new ErrorPresentacion('posición incompleta');
        return X().aplicar(xml, ed);
      });
    }

    // ---- texto -----------------------------------------------------------------------------------------------------
    const textoParrafo = (xml, p, salto = '\n') => {
      let s = '';
      const recorrer = (n) => { for (const h of n.hijos) { if (h.local === 't') s += X().decodificar(X().interior(xml, h)); else if (h.local === 'br') s += salto; else if (h.local === 'r' || h.local === 'fld') recorrer(h); } };
      recorrer(p);
      return s;
    };
    async function textoDe(i, camino) {
      const ruta = await rutaDe(i);
      const xml = await leer(ruta);
      const nodo = buscar(X().analizar(xml), camino);
      const cuerpo = nodo && X().hijo(nodo, 'txBody');
      if (!cuerpo) return [];
      return X().hijos(cuerpo, 'p').map((p) => textoParrafo(xml, p));
    }
    // como textoDe, pero un salto de línea dentro del párrafo (<a:br/>) sale como U+2028: así no se confunde con el cambio de párrafo
    async function textoEditable(i, camino) {
      const ruta = await rutaDe(i);
      const xml = await leer(ruta);
      const nodo = buscar(X().analizar(xml), camino);
      const cuerpo = nodo && X().hijo(nodo, 'txBody');
      return cuerpo ? X().hijos(cuerpo, 'p').map((p) => textoParrafo(xml, p, '\u2028')) : [];
    }
    // opTexto: cada elemento es un párrafo; U+2028 dentro de uno es un salto de línea (<a:br/>); un \n parte el párrafo en dos
    function opTexto(i, camino, parrafos) {
      return editarForma(i, camino, (xml, raiz, nodo, pref) => {
        if (nodo.local !== 'sp') throw new ErrorPresentacion('esta forma no lleva texto');
        const a = pref.a;
        const lineas = parrafos.flatMap((t) => String(t).split('\n'));
        let cuerpo = X().hijo(nodo, 'txBody');
        if (!cuerpo) {
          const nuevo = `<${pref.p}:txBody><${a}:bodyPr/><${a}:lstStyle/>` + lineas.map((l) => `<${a}:p>${l ? `<${a}:r><${a}:rPr lang="es-ES" dirty="0"/><${a}:t>${esc(l)}</${a}:t></${a}:r>` : ''}<${a}:endParaRPr lang="es-ES" dirty="0"/></${a}:p>`).join('') + `</${pref.p}:txBody>`;
          return insertarHijo(xml, nodo, 'txBody', nuevo, ['nvSpPr', 'spPr', 'style', 'txBody', 'extLst']);
        }
        const viejos = X().hijos(cuerpo, 'p');
        const textosViejos = viejos.map((p) => textoParrafo(xml, p, '\u2028'));
        const corridas = (rPr, l) => l.split('\u2028').map((s) => (s ? `<${a}:r>${rPr}<${a}:t>${esc(s)}</${a}:t></${a}:r>` : '')).join(`<${a}:br/>`);
        // los párrafos iguales del principio y del final quedan tal cual, con todo su formato (corridas, campos, saltos): así agregar o quitar
        // un párrafo no rehace los demás. Un párrafo que cambió se escribe de nuevo con el formato de su primera corrida
        // Como PowerPoint: al editar un párrafo solo cambia lo que cambió. Se comparan el texto viejo y el nuevo (lo igual del principio y del final),
        // las corridas, campos y saltos fuera del cambio quedan tal cual y el texto que se escribe toma el formato del carácter anterior
        // (o del siguiente, si se escribe al principio).
        function editarParrafo(pv, nuevoTexto) {
          const items = [];
          for (const h of pv.hijos) if (h.local === 'r' || h.local === 'br' || h.local === 'fld') items.push({ h, tipo: h.local, texto: h.local === 'br' ? '\u2028' : textoParrafo(xml, h, '\u2028') });
          const viejo = items.map((x) => x.texto).join('');
          let cp = 0; while (cp < viejo.length && cp < nuevoTexto.length && viejo[cp] === nuevoTexto[cp]) cp++;
          let cs = 0; while (cs < viejo.length - cp && cs < nuevoTexto.length - cp && viejo[viejo.length - 1 - cs] === nuevoTexto[nuevoTexto.length - 1 - cs]) cs++;
          const finBorrado = viejo.length - cs, insertado = nuevoTexto.slice(cp, nuevoTexto.length - cs);
          const rPrDe = (x) => { const rp = x && x.tipo !== 'br' && X().hijo(x.h, 'rPr'); return rp ? nodoTexto(xml, rp) : ''; };
          const ep = X().hijo(pv, 'endParaRPr');
          const deFin = ep ? nodoTexto(xml, ep).replace(/^<([\w.-]+:)?endParaRPr/, (mm, q) => '<' + (q || '') + 'rPr').replace(/<\/([\w.-]+:)?endParaRPr>$/, (mm, q) => '</' + (q || '') + 'rPr>') : '';
          // formato de lo escrito (como PowerPoint): si reemplaza algo, el del primer carácter reemplazado; si solo inserta, el del carácter
          // anterior; al principio del párrafo, el de la primera corrida
          let base = null, off = 0;
          for (const x of items) { const s = off, e = off + x.texto.length; off = e; if (x.tipo === 'r' && (finBorrado > cp ? s <= cp && cp < e : s < cp && cp <= e)) base = x; }
          if (!base) base = items.find((x) => x.tipo === 'r') || null;
          const rPrNuevo = base ? rPrDe(base) : deFin;
          const corrida = (rp, s) => (s ? `<${a}:r>${rp}<${a}:t>${esc(s)}</${a}:t></${a}:r>` : '');
          const piezasP = [];
          let colocado = false, pos = 0;
          const ponerInsertado = () => { if (!colocado) { colocado = true; if (insertado) piezasP.push(corridas(rPrNuevo, insertado)); } };
          for (const x of items) {
            const s = pos, e = pos + x.texto.length; pos = e;
            if (e === cp && x === base && finBorrado === cp && !insertado.includes('\u2028')) { piezasP.push(corrida(rPrDe(x), x.texto + insertado)); colocado = true; continue; }   // se escribe al final de la corrida: sigue siendo una sola
            if (e <= cp) { piezasP.push(nodoTexto(xml, x.h)); continue; }
            if (s >= finBorrado) { ponerInsertado(); piezasP.push(nodoTexto(xml, x.h)); continue; }
            // el elemento toca la parte cambiada
            if (x.tipo === 'r') {
              const antes = s < cp ? x.texto.slice(0, cp - s) : '', despues = e > finBorrado ? x.texto.slice(finBorrado - s) : '';
              const mismo = x === base && !insertado.includes('\u2028');
              if (mismo) { piezasP.push(corrida(rPrDe(x), antes + insertado + despues)); colocado = true; continue; }
              if (antes) piezasP.push(corrida(rPrDe(x), antes));
              ponerInsertado();
              if (despues) { ponerInsertado(); piezasP.push(corrida(rPrDe(x), despues)); }
            } else if (s < cp && x.tipo === 'fld') { piezasP.push(nodoTexto(xml, x.h)); ponerInsertado(); }
            else ponerInsertado();   // un campo es una sola pieza: si el cambio empieza dentro de él, se conserva entero
          }
          ponerInsertado();
          return armarParrafo(xml, pv, piezasP);
        }
        const n = lineas.length, m = viejos.length;
        let pre = 0; while (pre < n && pre < m && textosViejos[pre] === lineas[pre]) pre++;
        let suf = 0; while (suf < n - pre && suf < m - pre && textosViejos[m - 1 - suf] === lineas[n - 1 - suf]) suf++;
        const piezas = lineas.map((l, k) => {
          if (k < pre) return nodoTexto(xml, viejos[k]);
          if (k >= n - suf) return nodoTexto(xml, viejos[m - n + k]);
          if (k - pre < m - suf - pre) return editarParrafo(viejos[k], l);   // el párrafo que ya estaba en ese lugar, editado en el sitio
          // el modelo: el párrafo viejo en ese lugar; si solo se agregaron párrafos, el de antes (o el de después, si se agregó al principio)
          const modelo = m - suf > pre ? viejos[Math.min(k, m - suf - 1)] : viejos[pre > 0 ? pre - 1 : Math.min(pre, m - 1)];
          let pPr = '', rPr = '', fin = '';
          if (modelo) {
            const pp = X().hijo(modelo, 'pPr'); if (pp) pPr = nodoTexto(xml, pp);
            const r = X().hijo(modelo, 'r'), rp = r && X().hijo(r, 'rPr'); if (rp) rPr = nodoTexto(xml, rp);
            const ep = X().hijo(modelo, 'endParaRPr'); if (ep) fin = nodoTexto(xml, ep);
            if (!rPr && ep) rPr = fin.replace(/^<([\w.-]+:)?endParaRPr/, (m, p) => '<' + (p || '') + 'rPr').replace(/<\/([\w.-]+:)?endParaRPr>$/, (m, p) => '</' + (p || '') + 'rPr>');
          }
          return `<${a}:p>${pPr}${l ? corridas(rPr, l) : ''}${fin}</${a}:p>`;
        });
        if (!piezas.length) piezas.push(`<${a}:p/>`);
        const primero = viejos[0], ultimo = viejos[viejos.length - 1];
        if (!primero) return insertarHijo(xml, cuerpo, 'p', piezas.join(''), null);
        return X().aplicar(xml, [{ ini: primero.ini, fin: ultimo.fin, texto: piezas.join('') }]);
      });
    }
    // un párrafo con otras piezas de texto: su apertura, sus propiedades, las piezas y lo que viene después (endParaRPr, extLst)
    function armarParrafo(xml, pv, piezas) {
      const abre = xml.slice(pv.ini, pv.finApertura).replace(/\s*\/>$/, '>');
      const pPr = X().hijo(pv, 'pPr'), resto = pv.hijos.filter((h) => h !== pPr && !['r', 'br', 'fld'].includes(h.local));
      return abre + (pPr ? nodoTexto(xml, pPr) : '') + piezas.join('') + resto.map((h) => nodoTexto(xml, h)).join('') + '</' + pv.nombre + '>';
    }
    // devuelve la función que cambia un <a:rPr> (o <a:endParaRPr>) según f: { negrita, cursiva, subrayado, tamano, color }
    const crearAjuste = (f, a) => {
        const ajustar = (etiqueta, xmlEl) => {
          let ap = xmlEl.slice(0, etiqueta.finApertura - etiqueta.ini);
          if ('negrita' in f) ap = X().ponerAtributo(ap, 'b', f.negrita ? '1' : '0');
          if ('cursiva' in f) ap = X().ponerAtributo(ap, 'i', f.cursiva ? '1' : '0');
          if ('subrayado' in f) ap = X().ponerAtributo(ap, 'u', f.subrayado ? 'sng' : 'none');
          if ('tamano' in f && f.tamano) ap = X().ponerAtributo(ap, 'sz', String(Math.round(f.tamano * 100)));
          let interior = etiqueta.iniCierre < 0 ? '' : xmlEl.slice(etiqueta.finApertura - etiqueta.ini, etiqueta.iniCierre - etiqueta.ini);
          if ('color' in f) {
            interior = interior.replace(new RegExp('<([\\w.-]+:)?(solidFill|gradFill|noFill|pattFill)\\b[\\s\\S]*?(?:/>|</\\1?\\2>)'), '');
            if (f.color) interior = `<${a}:solidFill><${a}:srgbClr val="${f.color.toUpperCase()}"/></${a}:solidFill>` + interior;
          }
          const vacio = /\/>$/.test(ap) && !interior;
          if (vacio) return ap;
          return ap.replace(/\s*\/>$/, '>') + interior + '</' + etiqueta.nombre + '>';
        };
      return ajustar;
    };
    // las partes de un párrafo que llevan texto: corridas, saltos de línea y campos, con su texto (el salto es U+2028)
    const itemsDe = (xml, pv) => pv.hijos.filter((h) => h.local === 'r' || h.local === 'br' || h.local === 'fld').map((h) => ({ h, tipo: h.local, texto: h.local === 'br' ? '\u2028' : textoParrafo(xml, h, '\u2028') }));
    // el mismo formato en una parte del texto: ini y fin cuentan caracteres del texto de la forma (párrafos unidos con \n, saltos de línea como U+2028).
    // Como PowerPoint, las corridas que el rango corta se parten en dos o tres con el mismo formato y solo la del medio cambia.
    function opFormatoRango(i, camino, ini, fin, f) {
      return editarForma(i, camino, (xml, raiz, nodo, pref) => {
        const cuerpo = X().hijo(nodo, 'txBody');
        if (!cuerpo) return xml;
        const a = pref.a, ajustar = crearAjuste(f, a), ed = [];
        const rprNuevo = (rp) => (rp ? ajustar(rp, nodoTexto(xml, rp)) : ajustar({ ini: 0, finApertura: `<${a}:rPr lang="es-ES"/>`.length, iniCierre: -1, nombre: a + ':rPr' }, `<${a}:rPr lang="es-ES"/>`));
        const corrida = (rp, s) => (s ? `<${a}:r>${rp}<${a}:t>${esc(s)}</${a}:t></${a}:r>` : '');
        let base = 0;
        for (const pv of X().hijos(cuerpo, 'p')) {
          const items = itemsDe(xml, pv), largo = items.reduce((n, x) => n + x.texto.length, 0);
          const lo = Math.max(ini, base) - base, hi = Math.min(fin, base + largo) - base;
          base += largo + 1;
          if (hi <= lo) continue;
          const piezas = [];
          let pos = 0;
          for (const x of items) {
            const s = pos, e = pos + x.texto.length; pos = e;
            if (e <= lo || s >= hi || x.tipo === 'br') { piezas.push(nodoTexto(xml, x.h)); continue; }
            const rp = X().hijo(x.h, 'rPr');
            if (x.tipo === 'fld') {   // un campo es una sola pieza: si el rango lo toca, se formatea entero
              const txt = nodoTexto(xml, x.h);
              piezas.push(rp ? txt.slice(0, rp.ini - x.h.ini) + rprNuevo(rp) + txt.slice(rp.fin - x.h.ini) : txt.slice(0, x.h.finApertura - x.h.ini) + rprNuevo(null) + txt.slice(x.h.finApertura - x.h.ini));
              continue;
            }
            const viejo = rp ? nodoTexto(xml, rp) : '';
            const antes = x.texto.slice(0, Math.max(0, lo - s)), medio = x.texto.slice(Math.max(0, lo - s), Math.min(x.texto.length, hi - s)), despues = x.texto.slice(Math.min(x.texto.length, hi - s));
            piezas.push(corrida(viejo, antes), corrida(rprNuevo(rp), medio), corrida(viejo, despues));
          }
          ed.push({ ini: pv.ini, fin: pv.fin, texto: armarParrafo(xml, pv, piezas) });
        }
        return ed.length ? X().aplicar(xml, ed) : xml;
      });
    }
    // el mismo formato en todo el texto de la forma (todas las corridas y el final de párrafo)
    function opFormatoTexto(i, camino, f) {
      return editarForma(i, camino, (xml, raiz, nodo, pref) => {
        const cuerpo = X().hijo(nodo, 'txBody');
        if (!cuerpo) return xml;
        const a = pref.a, ed = [];
        const ajustar = crearAjuste(f, a);
        const recorrer = (n) => {
          for (const h of n.hijos) {
            if (h.local === 'p') recorrer(h);
            else if (h.local === 'r' || h.local === 'fld') {
              const rp = X().hijo(h, 'rPr');
              if (rp) ed.push({ ini: rp.ini, fin: rp.fin, texto: ajustar(rp, nodoTexto(xml, rp)) });
              else ed.push({ ini: h.finApertura, fin: h.finApertura, texto: ajustar({ ini: 0, finApertura: `<${a}:rPr lang="es-ES"/>`.length, iniCierre: -1, nombre: a + ':rPr' }, `<${a}:rPr lang="es-ES"/>`) });
            } else if (h.local === 'endParaRPr') ed.push({ ini: h.ini, fin: h.fin, texto: ajustar(h, nodoTexto(xml, h)) });
          }
        };
        recorrer(cuerpo);
        return ed.length ? X().aplicar(xml, ed) : xml;
      });
    }

    // ---- relleno, línea, borrar, orden ------------------------------------------------------------------------------------
    const colorXml = (a, c) => `<${a}:solidFill><${a}:srgbClr val="${c.toUpperCase()}"/></${a}:solidFill>`;
    function opRelleno(i, camino, color) {
      return editarForma(i, camino, (xml, raiz, nodo, pref) => {
        const spPr = X().hijo(nodo, nodo.local === 'grpSp' ? 'grpSpPr' : 'spPr');
        if (!spPr || nodo.local === 'graphicFrame') throw new ErrorPresentacion('esta forma no tiene relleno');
        const viejo = spPr.hijos.find((h) => RELLENOS.includes(h.local));
        const nuevo = color ? colorXml(pref.a, color) : `<${pref.a}:noFill/>`;
        if (viejo) return X().aplicar(xml, [{ ini: viejo.ini, fin: viejo.fin, texto: nuevo }]);
        return insertarHijo(xml, spPr, 'solidFill', nuevo, ORDEN_SPPR);
      });
    }
    function opLinea(i, camino, color) {
      return editarForma(i, camino, (xml, raiz, nodo, pref) => {
        const spPr = X().hijo(nodo, 'spPr');
        if (!spPr) throw new ErrorPresentacion('esta forma no tiene línea');
        const a = pref.a, ln = X().hijo(spPr, 'ln');
        const relleno = color ? colorXml(a, color) : `<${a}:noFill/>`;
        if (!ln) return insertarHijo(xml, spPr, 'ln', `<${a}:ln>${relleno}</${a}:ln>`, ORDEN_SPPR);
        const viejo = ln.hijos.find((h) => ['noFill', 'solidFill', 'gradFill', 'pattFill'].includes(h.local));
        if (viejo) return X().aplicar(xml, [{ ini: viejo.ini, fin: viejo.fin, texto: relleno }]);
        return insertarHijo(xml, ln, 'solidFill', relleno, ['noFill', 'solidFill', 'gradFill', 'pattFill', 'prstDash', 'custDash', 'round', 'bevel', 'miter', 'headEnd', 'tailEnd', 'extLst']);
      });
    }
    const opBorrar = (i, camino) => editarForma(i, camino, (xml, raiz, nodo) => X().aplicar(xml, [{ ini: nodo.ini, fin: nodo.fin, texto: '' }]));
    function opOrden(i, camino, hacia) {
      return editarForma(i, camino, (xml, raiz, nodo) => {
        const padre = camino.length > 1 ? buscar(raiz, camino.slice(0, -1)) : arbolDe(raiz);
        const formas = padre.hijos.filter((h) => FORMAS.includes(h.local) || h.local === 'AlternateContent');
        const destino = hacia === 'frente' ? formas[formas.length - 1] : formas[0];
        if (!destino || destino === nodo) return xml;
        const t = nodoTexto(xml, nodo);
        const sin = X().aplicar(xml, [{ ini: nodo.ini, fin: nodo.fin, texto: '' }]);
        const pos = hacia === 'frente' ? destino.fin - (nodo.fin - nodo.ini) : destino.ini;
        return X().aplicar(sin, [{ ini: pos, fin: pos, texto: t }]);
      });
    }

    // ---- formas nuevas ------------------------------------------------------------------------------------------------------
    const mayorId = (xml) => { let m = 1; for (const x of xml.matchAll(/<(?:[\w.-]+:)?cNvPr\b[^>]*?\bid="(\d+)"/g)) m = Math.max(m, +x[1]); return m; };
    async function insertarEnArbol(i, crear) {
      const ruta = await rutaDe(i);
      const xml = await leer(ruta);
      const raiz = X().analizar(xml);
      const arbol = arbolDe(raiz);
      if (!arbol) throw new ErrorPresentacion('diapositiva sin árbol de formas');
      const pref = { a: prefijo(xml, NS_A, 'a'), p: prefijo(xml, NS_P, 'p'), r: prefijo(xml, NS_R, 'r') };
      const id = mayorId(xml) + 1;
      const extra = await crear(id, pref, ruta);
      let nuevo = X().aplicar(xml, [{ ini: arbol.iniCierre, fin: arbol.iniCierre, texto: extra.xml }]);
      if (!new RegExp('xmlns:' + pref.r + '=').test(nuevo)) nuevo = nuevo.replace(/^(<\?xml[^>]*>\s*)?(<[\w.-]+:sld\b)/, (m, d, et) => (d || '') + et + ` xmlns:${pref.r}="${NS_R}"`);
      return { op: { t: 'partes', cambios: [{ ruta, datos: nuevo, antes: xml }].concat(extra.cambios || []) }, id };
    }
    const xfrmXml = (a, c) => `<${a}:xfrm><${a}:off x="${Math.round(c.x)}" y="${Math.round(c.y)}"/><${a}:ext cx="${Math.round(c.cx)}" cy="${Math.round(c.cy)}"/></${a}:xfrm>`;
    function opCuadroTexto(i, caja, texto) {
      return insertarEnArbol(i, (id, { a, p }) => ({
        xml: `<${p}:sp><${p}:nvSpPr><${p}:cNvPr id="${id}" name="CuadroTexto ${id}"/><${p}:cNvSpPr txBox="1"/><${p}:nvPr/></${p}:nvSpPr><${p}:spPr>${xfrmXml(a, caja)}<${a}:prstGeom prst="rect"><${a}:avLst/></${a}:prstGeom><${a}:noFill/></${p}:spPr>` +
          `<${p}:txBody><${a}:bodyPr wrap="square" rtlCol="0"><${a}:spAutoFit/></${a}:bodyPr><${a}:lstStyle/>` +
          String(texto || '').split('\n').map((l) => `<${a}:p>${l ? `<${a}:r><${a}:rPr lang="es-ES" dirty="0"/><${a}:t>${esc(l)}</${a}:t></${a}:r>` : `<${a}:endParaRPr lang="es-ES" dirty="0"/>`}</${a}:p>`).join('') + `</${p}:txBody></${p}:sp>`,
      }));
    }
    // Una forma predefinida de la norma (ECMA-376, prstGeom): se escribe solo su nombre; el dibujo lo hace pptx-geometria.js con las definiciones de
    // vendor/ecma-376-geometrias (sin modificarlas: ver AVISO-ECMA.txt). Como PowerPoint, sin color explícito la forma sigue el tema (p:style: relleno y
    // contorno del color de énfasis 1, texto claro); las líneas y conectores van en un p:cxnSp. Con `color` (hex) se escribe ese relleno y sin contorno.
    //   opciones: { nombre, flecha } — nombre: el que se le pone (PowerPoint usa «Rectángulo 3»); flecha: punta de flecha al final de un conector
    const CONECTORES = ['line', 'straightConnector1', 'bentConnector3', 'curvedConnector3'];
    function opForma(i, caja, prst, color, opciones) {
      const o = opciones || {};
      if (!/^[A-Za-z][A-Za-z0-9]*$/.test(String(prst))) throw new ErrorPresentacion('forma desconocida');
      return insertarEnArbol(i, (id, { a, p }) => {
        const nombre = esc((o.nombre || (CONECTORES.includes(prst) ? 'Conector' : 'Forma')) + ' ' + id).replace(/"/g, '&quot;');
        const geom = `<${a}:prstGeom prst="${prst}"><${a}:avLst/></${a}:prstGeom>`;
        const acento = (extra) => `<${a}:schemeClr val="accent1">${extra || ''}</${a}:schemeClr>`;
        if (CONECTORES.includes(prst)) {
          const punta = o.flecha ? `<${a}:ln><${a}:tailEnd type="triangle"/></${a}:ln>` : '';
          return { xml: `<${p}:cxnSp><${p}:nvCxnSpPr><${p}:cNvPr id="${id}" name="${nombre}"/><${p}:cNvCxnSpPr/><${p}:nvPr/></${p}:nvCxnSpPr><${p}:spPr>${xfrmXml(a, caja)}${geom}${punta}</${p}:spPr>`
            + `<${p}:style><${a}:lnRef idx="1">${acento()}</${a}:lnRef><${a}:fillRef idx="0">${acento()}</${a}:fillRef><${a}:effectRef idx="0">${acento()}</${a}:effectRef><${a}:fontRef idx="minor"><${a}:schemeClr val="tx1"/></${a}:fontRef></${p}:style></${p}:cxnSp>` };
        }
        const cuerpo = `<${p}:txBody><${a}:bodyPr rtlCol="0" anchor="ctr"/><${a}:lstStyle/><${a}:p><${a}:pPr algn="ctr"/><${a}:endParaRPr lang="es-ES" dirty="0"/></${a}:p></${p}:txBody>`;
        if (color) return { xml: `<${p}:sp><${p}:nvSpPr><${p}:cNvPr id="${id}" name="${nombre}"/><${p}:cNvSpPr/><${p}:nvPr/></${p}:nvSpPr><${p}:spPr>${xfrmXml(a, caja)}${geom}${colorXml(a, color)}<${a}:ln><${a}:noFill/></${a}:ln></${p}:spPr>${cuerpo}</${p}:sp>` };
        return { xml: `<${p}:sp><${p}:nvSpPr><${p}:cNvPr id="${id}" name="${nombre}"/><${p}:cNvSpPr/><${p}:nvPr/></${p}:nvSpPr><${p}:spPr>${xfrmXml(a, caja)}${geom}</${p}:spPr>`
          + `<${p}:style><${a}:lnRef idx="2">${acento('<' + a + ':shade val="50000"/>')}</${a}:lnRef><${a}:fillRef idx="1">${acento()}</${a}:fillRef><${a}:effectRef idx="0">${acento()}</${a}:effectRef><${a}:fontRef idx="minor"><${a}:schemeClr val="lt1"/></${a}:fontRef></${p}:style>${cuerpo}</${p}:sp>` };
      });
    }
    async function opImagen(i, bytes, tipo, caja) {
      const ext = tipo === 'image/png' ? 'png' : tipo === 'image/jpeg' ? 'jpeg' : tipo === 'image/gif' ? 'gif' : null;
      if (!ext) throw new ErrorPresentacion('formato de imagen no admitido (PNG, JPEG o GIF)');
      let k = 1; while (zip.file('ppt/media/imagen-mc' + k + '.' + ext)) k++;
      const media = 'ppt/media/imagen-mc' + k + '.' + ext;
      const ruta = await rutaDe(i);
      const rutaRels = PaqueteEditable.rutaRels(ruta);
      const relsAntes = await leer(rutaRels);
      const ctAntes = await leer('[Content_Types].xml');
      // la relación y el tipo de contenido se calculan sobre copias de texto (el op guarda lo de antes para deshacer)
      const relsBase = relsAntes || `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
      const usados = new Set([...relsBase.matchAll(/\bId="([^"]+)"/g)].map((m) => m[1]));
      let n = usados.size + 1; while (usados.has('rId' + n)) n++;
      const rid = 'rId' + n;
      const relsNuevo = relsBase.replace(/<\/Relationships>\s*$/, `<Relationship Id="${rid}" Type="${REL_IMAGEN}" Target="../media/imagen-mc${k}.${ext}"/></Relationships>`).replace(/<Relationships([^>]*)\/>\s*$/, `<Relationships$1><Relationship Id="${rid}" Type="${REL_IMAGEN}" Target="../media/imagen-mc${k}.${ext}"/></Relationships>`);
      const tieneDefault = new RegExp('<Default\\b[^>]*Extension="' + ext + '"', 'i').test(ctAntes);
      const ctNuevo = tieneDefault ? ctAntes : ctAntes.replace(/<\/Types>\s*$/, `<Default Extension="${ext}" ContentType="${tipo}"/></Types>`);
      return insertarEnArbol(i, (id, { a, p, r }) => ({
        xml: `<${p}:pic><${p}:nvPicPr><${p}:cNvPr id="${id}" name="Imagen ${id}"/><${p}:cNvPicPr><${a}:picLocks noChangeAspect="1"/></${p}:cNvPicPr><${p}:nvPr/></${p}:nvPicPr><${p}:blipFill><${a}:blip ${r}:embed="${rid}"/><${a}:stretch><${a}:fillRect/></${a}:stretch></${p}:blipFill><${p}:spPr>${xfrmXml(a, caja)}<${a}:prstGeom prst="rect"><${a}:avLst/></${a}:prstGeom></${p}:spPr></${p}:pic>`,
        cambios: [{ ruta: media, datos: bytes, antes: null }, { ruta: rutaRels, datos: relsNuevo, antes: relsAntes }].concat(tieneDefault ? [] : [{ ruta: '[Content_Types].xml', datos: ctNuevo, antes: ctAntes }]),
      }));
    }

    // ---- diapositivas -------------------------------------------------------------------------------------------------------
    async function contexto() {
      const pres = await presXml();
      const relsRuta = 'ppt/_rels/presentation.xml.rels';
      const rels = await leer(relsRuta);
      const ct = await leer('[Content_Types].xml');
      return { pres, rels, relsRuta, ct, pp: prefijo(pres, NS_P, 'p'), pr: prefijo(pres, NS_R, 'r') };
    }
    const siguienteParte = () => { let k = 1; while (zip.file('ppt/slides/slide' + k + '.xml')) k++; return k; };
    function nuevoRid(rels) { const usados = new Set([...rels.matchAll(/\bId="([^"]+)"/g)].map((m) => m[1])); let n = usados.size + 1; while (usados.has('rId' + n)) n++; return 'rId' + n; }
    function nuevoSldId(pres) { let m = 255; for (const x of pres.matchAll(/<(?:[\w.-]+:)?sldId\b[^>]*?\bid="(\d+)"/g)) m = Math.max(m, +x[1]); return m + 1; }
    // agrega la parte de una diapositiva después de la posición `despues` (el resto: relaciones, tipos, lista y secciones)
    async function opAgregarDiapositiva(despues, xmlDiapositiva, relsDiapositiva) {
      const c = await contexto();
      const k = siguienteParte(), ruta = 'ppt/slides/slide' + k + '.xml';
      const rid = nuevoRid(c.rels), sid = nuevoSldId(c.pres);
      const relsNuevo = c.rels.replace(/<\/Relationships>\s*$/, `<Relationship Id="${rid}" Type="${REL_SLIDE}" Target="slides/slide${k}.xml"/></Relationships>`);
      const ctNuevo = c.ct.replace(/<\/Types>\s*$/, `<Override PartName="/${ruta}" ContentType="${TIPO_SLIDE}"/></Types>`);
      const raiz = X().analizar(c.pres, { profundidad: 3 }), lista = X().hijo(raiz, 'sldIdLst');
      const ids = lista ? X().hijos(lista, 'sldId') : [];
      const nodo = `<${c.pp}:sldId id="${sid}" ${c.pr}:id="${rid}"/>`;
      let pres;
      if (!lista) pres = X().aplicar(c.pres, [{ ini: X().hijo(raiz, 'sldMasterIdLst').fin, fin: X().hijo(raiz, 'sldMasterIdLst').fin, texto: `<${c.pp}:sldIdLst>${nodo}</${c.pp}:sldIdLst>` }]);
      else {
        const tras = ids[Math.min(despues, ids.length - 1)];
        const pos = despues < 0 || !tras ? (ids[0] ? ids[0].ini : lista.iniCierre) : tras.fin;
        pres = X().aplicar(c.pres, [{ ini: pos, fin: pos, texto: nodo }]);
        // secciones (p14:sectionLst): la nueva va en la sección de la diapositiva de antes
        if (tras) pres = agregarASeccion(pres, tras.attrs.id, sid);
      }
      return { t: 'partes', cambios: [
        { ruta, datos: xmlDiapositiva, antes: null }, { ruta: PaqueteEditable.rutaRels(ruta), datos: relsDiapositiva, antes: null },
        { ruta: c.relsRuta, datos: relsNuevo, antes: c.rels }, { ruta: '[Content_Types].xml', datos: ctNuevo, antes: c.ct }, { ruta: 'ppt/presentation.xml', datos: pres, antes: c.pres },
      ] };
    }
    // en <p14:sectionLst>, el id nuevo justo después del de referencia (las secciones de PowerPoint 2010 listan los ids de sus diapositivas)
    function agregarASeccion(pres, idRef, idNuevo) {
      return pres.replace(new RegExp('(<([\\w.-]+:)sldId id="' + idRef + '"\\s*/>)(?![\\s\\S]*<([\\w.-]+:)?sldIdLst>)', 'g'), (m, et, pfx) => (/^<p\d*:|^<p14:/.test(et) && pfx !== 'p:' ? et + `<${pfx}sldId id="${idNuevo}"/>` : m));
    }
    async function opDuplicarDiapositiva(i) {
      const ruta = await rutaDe(i);
      const xml = await leer(ruta);
      let rels = (await leer(PaqueteEditable.rutaRels(ruta))) || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
      // Como PowerPoint: la copia lleva sus propias notas y sus propios comentarios (compartir una parte entre dos diapositivas haría que PowerPoint
      // repare el archivo). Se copian las partes con nombre nuevo; los comentarios modernos (con identificadores propios) no se copian.
      const k = siguienteParte(), extra = [], overrides = [];
      const libre = (patron) => { let n = 1; while (zip.file(patron(n)) || extra.some((e) => e.ruta === patron(n))) n++; return patron(n); };
      for (const r of await paquete.relaciones(ruta)) {
        const m = new RegExp('<Relationship\\b[^>]*\\bId="' + r.id + '"[^>]*/>').exec(rels);
        if (!m) continue;
        if (/\/notesSlide$/.test(r.tipo) && zip.file(r.destino)) {
          const nueva = libre((n) => 'ppt/notesSlides/notesSlide' + n + '.xml');
          extra.push({ ruta: nueva, datos: await leer(r.destino), antes: null });
          const relsNota = await leer(PaqueteEditable.rutaRels(r.destino));
          if (relsNota) extra.push({ ruta: PaqueteEditable.rutaRels(nueva), datos: relsNota.replace(/(<Relationship\b[^>]*Type="[^"]*\/slide"[^>]*Target=")[^"]*(")/, '$1../slides/slide' + k + '.xml$2'), antes: null });
          overrides.push(['/' + nueva, 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml']);
          rels = rels.replace(m[0], m[0].replace(/Target="[^"]*"/, 'Target="../notesSlides/' + nueva.split('/').pop() + '"'));
        } else if (/\/relationships\/comments$/.test(r.tipo) && zip.file(r.destino)) {
          const nueva = libre((n) => 'ppt/comments/comment' + n + '.xml');
          extra.push({ ruta: nueva, datos: await leer(r.destino), antes: null });
          overrides.push(['/' + nueva, 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml']);
          rels = rels.replace(m[0], m[0].replace(/Target="[^"]*"/, 'Target="../comments/' + nueva.split('/').pop() + '"'));
        } else if (/2018\/10\/relationships\/comments$/.test(r.tipo)) {
          rels = rels.replace(m[0], '');
        }
      }
      const op = await opAgregarDiapositiva(i, xml, rels);
      if (overrides.length) {
        const ctc = op.cambios.find((c) => c.ruta === '[Content_Types].xml');
        ctc.datos = ctc.datos.replace(/<\/Types>\s*$/, overrides.map(([pn, ty]) => '<Override PartName="' + pn + '" ContentType="' + ty + '"/>').join('') + '</Types>');
      }
      op.cambios.push(...extra);
      return op;
    }
    async function opNuevaDiapositiva(i) {
      const ruta = await rutaDe(i);
      const rels = await paquete.relaciones(ruta);
      const layout = rels.find((r) => /slideLayout$/.test(r.tipo));
      if (!layout) throw new ErrorPresentacion('la diapositiva no tiene diseño');
      const xl = (await leer(layout.destino)) || '';
      const pA = prefijo(xl, NS_A, 'a'), pP = prefijo(xl, NS_P, 'p');
      // los marcadores del diseño (título, cuerpo…) sin su texto de ejemplo; fecha, pie y número no
      let id = 1, formas = '';
      for (const m of xl.matchAll(/<([\w.-]+:)?sp\b[\s\S]*?<\/\1?sp>/g)) {
        const ph = /<(?:[\w.-]+:)?ph\b([^>]*)\/?>/.exec(m[0]);
        if (!ph || /type="(dt|ftr|sldNum)"/.test(ph[1])) continue;
        const nombre = (/<(?:[\w.-]+:)?cNvPr\b[^>]*?\bname="([^"]*)"/.exec(m[0]) || [, 'Marcador'])[1];
        id++;
        formas += `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${nombre}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${ph[1].replace(/\s*\/$/, '')}/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="es-ES" dirty="0"/></a:p></p:txBody></p:sp>`;
      }
      void pA; void pP;
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sld xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${formas}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
      const destino = layout.destino.replace(/^ppt\//, '../');
      const relsNuevo = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_LAYOUT}" Target="${destino}"/></Relationships>`;
      return opAgregarDiapositiva(i, xml, relsNuevo);
    }
    async function opBorrarDiapositiva(i) {
      const lista = await rutas();
      if (lista.length <= 1) throw new ErrorPresentacion('una presentación necesita al menos una diapositiva');
      const ruta = lista[i];
      const c = await contexto();
      const relsPres = await paquete.relaciones('ppt/presentation.xml');
      const rel = relsPres.find((r) => r.destino === ruta);
      const raiz = X().analizar(c.pres, { profundidad: 3 }), ids = X().hijos(X().hijo(raiz, 'sldIdLst'), 'sldId');
      const nodo = ids.find((s) => Object.entries(s.attrs).some(([k, v]) => /(^|:)id$/.test(k) && k !== 'id' && v === rel.id));
      let pres = X().aplicar(c.pres, [{ ini: nodo.ini, fin: nodo.fin, texto: '' }]);
      pres = pres.replace(new RegExp('<([\\w.-]+:)sldId id="' + nodo.attrs.id + '"\\s*/>', 'g'), (m, pfx) => (pfx === c.pp + ':' ? m : ''));   // también de las secciones
      const relsNuevo = c.rels.replace(new RegExp('<Relationship\\b[^>]*\\bId="' + rel.id + '"[^>]*/>'), '');
      const cambios = [{ ruta: 'ppt/presentation.xml', datos: pres, antes: c.pres }, { ruta: c.relsRuta, datos: relsNuevo, antes: c.rels }];
      let ct = c.ct;
      const quitarParte = async (r) => {
        cambios.push({ ruta: r, datos: null, antes: await paquete.leerBytes(r) });
        const rr = PaqueteEditable.rutaRels(r);
        if (zip.file(rr)) cambios.push({ ruta: rr, datos: null, antes: await paquete.leerBytes(rr) });
        ct = ct.replace(new RegExp('<Override\\b[^>]*PartName="/' + r.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"[^>]*/>'), '');
      };
      await quitarParte(ruta);
      // sus notas: solo las usa ella
      for (const n of (await paquete.relaciones(ruta)).filter((r) => /notesSlide$/.test(r.tipo))) if (zip.file(n.destino)) await quitarParte(n.destino);
      // otras diapositivas que enlazaban a esta: se quita la relación y el enlace (el texto o la forma quedan), como si no hubiera destino
      for (const otra of lista) {
        if (otra === ruta) continue;
        const rr = PaqueteEditable.rutaRels(otra);
        let relsOtra = await leer(rr);
        if (!relsOtra) continue;
        const colgantes = (await paquete.relaciones(otra)).filter((r) => /\/slide$/.test(r.tipo) && r.destino === ruta);
        if (!colgantes.length) continue;
        let xmlOtra = await leer(otra);
        const antesXml = xmlOtra, antesRels = relsOtra;
        for (const r of colgantes) {
          relsOtra = relsOtra.replace(new RegExp('<Relationship\\b[^>]*\\bId="' + r.id + '"[^>]*/>'), '');
          xmlOtra = xmlOtra.replace(new RegExp('<([\\w.-]+:)?(hlinkClick|hlinkHover)\\b[^>]*\\bid="' + r.id + '"[^>]*?(?:/>|>[\\s\\S]*?</\\1?\\2>)', 'g'), '');
        }
        cambios.push({ ruta: rr, datos: relsOtra, antes: antesRels }, { ruta: otra, datos: xmlOtra, antes: antesXml });
      }
      cambios.push({ ruta: '[Content_Types].xml', datos: ct, antes: c.ct });
      return { t: 'partes', cambios };
    }
    // p14:sectionLst reparte las diapositivas en bloques seguidos. Al mover una, sale de su sección y entra en la de la diapositiva que
    // le queda antes (o, si pasa a ser la primera, en la de la que le queda después); después cada sección toma sus ids del orden nuevo.
    function reacomodarSecciones(pres, orden, idMovido) {
      const lista = /<([\w.-]+:)sectionLst\b[^>]*>[\s\S]*?<\/\1sectionLst>/.exec(pres);
      if (!lista) return pres;
      const piezas = [];
      const re = /<([\w.-]+:)section\b[^>]*>[\s\S]*?<\/\1section>/g;
      let m;
      while ((m = re.exec(lista[0]))) {
        const sec = m[0], pf = m[1], inner = new RegExp('<' + pf + 'sldIdLst\\b[^>]*?(?:/>|>([\\s\\S]*?)</' + pf + 'sldIdLst>)').exec(sec);
        if (!inner) return pres;
        piezas.push({ ini: m.index, fin: m.index + sec.length, sec, pf, inner, ids: [...(inner[1] || '').matchAll(/\bid="(\d+)"/g)].map((x) => x[1]) });
      }
      if (!piezas.length) return pres;
      const de = piezas.findIndex((s) => s.ids.includes(idMovido));
      if (de < 0) return pres;
      piezas[de].ids = piezas[de].ids.filter((x) => x !== idMovido);
      const k = orden.indexOf(idMovido), vecino = orden[k > 0 ? k - 1 : k + 1];
      let a = piezas.findIndex((s) => s.ids.includes(vecino));
      if (a < 0) a = de;
      const tam = piezas.map((s) => s.ids.length); tam[a]++;
      let desde = 0;
      const nuevas = piezas.map((s, q) => {
        const ids = orden.slice(desde, desde + tam[q]); desde += tam[q];
        const lst = ids.length ? '<' + s.pf + 'sldIdLst>' + ids.map((x) => '<' + s.pf + 'sldId id="' + x + '"/>').join('') + '</' + s.pf + 'sldIdLst>' : '<' + s.pf + 'sldIdLst/>';
        return s.sec.slice(0, s.inner.index) + lst + s.sec.slice(s.inner.index + s.inner[0].length);
      });
      let bloque = '', cursor = 0;
      piezas.forEach((s, q) => { bloque += lista[0].slice(cursor, s.ini) + nuevas[q]; cursor = s.fin; });
      bloque += lista[0].slice(cursor);
      return pres.slice(0, lista.index) + bloque + pres.slice(lista.index + lista[0].length);
    }
    async function opMoverDiapositiva(i, j) {
      const c = await contexto();
      const raiz = X().analizar(c.pres, { profundidad: 3 }), lista = X().hijo(raiz, 'sldIdLst'), ids = X().hijos(lista, 'sldId');
      if (j < 0 || j >= ids.length || i === j) return null;
      const textos = ids.map((n) => nodoTexto(c.pres, n));
      const [m] = textos.splice(i, 1);
      textos.splice(j, 0, m);
      let pres = X().aplicar(c.pres, [{ ini: ids[0].ini, fin: ids[ids.length - 1].fin, texto: textos.join('') }]);
      const orden = ids.map((n) => n.attrs.id), movido = orden.splice(i, 1)[0];
      orden.splice(j, 0, movido);
      pres = reacomodarSecciones(pres, orden, movido);
      return { t: 'partes', cambios: [{ ruta: 'ppt/presentation.xml', datos: pres, antes: c.pres }] };
    }

    // ---- aplicar y guardar --------------------------------------------------------------------------------------------------
    function aplicar(op) {
      if (!op || op.t !== 'partes') throw new Error('operación desconocida');
      const inversa = { t: 'partes', cambios: [] };
      for (const c of op.cambios) {
        inversa.cambios.unshift({ ruta: c.ruta, datos: c.antes === undefined ? null : c.antes, antes: c.datos });
        if (c.datos === null || c.datos === undefined) paquete.quitar(c.ruta); else paquete.escribir(c.ruta, c.datos);
      }
      return inversa;
    }

    return {
      zip, paquete, rutas, textoDe, textoEditable, editable,
      opMover, opTexto, opFormatoTexto, opFormatoRango, opRelleno, opLinea, opBorrar, opOrden, opCuadroTexto, opForma, opImagen,
      opDuplicarDiapositiva, opNuevaDiapositiva, opBorrarDiapositiva, opMoverDiapositiva,
      aplicar, generar: () => paquete.generar(),
    };
  }

  // Presentación nueva en blanco (16:9): un patrón, un diseño «Diapositiva de título», un tema de Office y una diapositiva vacía.
  // Es el mínimo que PowerPoint abre sin reparar; se abre con abrir() como cualquier otra.
  async function nueva() {
    const D = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    const A = 'http://schemas.openxmlformats.org/drawingml/2006/main', P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
    const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', PK = 'http://schemas.openxmlformats.org/package/2006/relationships';
    const ns = `xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"`;
    const rel = (id, tipo, destino) => `<Relationship Id="${id}" Type="${R}/${tipo}" Target="${destino}"/>`;
    const rels = (...r) => D + `<Relationships xmlns="${PK}">${r.join('')}</Relationships>`;
    const arbol = (formas) => `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${formas || ''}</p:spTree></p:cSld>`;
    const mapa = '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
    const colores = [['dk1', '<a:sysClr val="windowText" lastClr="000000"/>'], ['lt1', '<a:sysClr val="window" lastClr="FFFFFF"/>'], ['dk2', '<a:srgbClr val="0E2841"/>'], ['lt2', '<a:srgbClr val="E8E8E8"/>'],
      ['accent1', '<a:srgbClr val="156082"/>'], ['accent2', '<a:srgbClr val="E97132"/>'], ['accent3', '<a:srgbClr val="196B24"/>'], ['accent4', '<a:srgbClr val="0F9ED5"/>'], ['accent5', '<a:srgbClr val="A02B93"/>'], ['accent6', '<a:srgbClr val="4EA72E"/>'],
      ['hlink', '<a:srgbClr val="467886"/>'], ['folHlink', '<a:srgbClr val="96607D"/>']].map(([k, v]) => `<a:${k}>${v}</a:${k}>`).join('');
    const fuente = (f) => `<a:latin typeface="${f}"/><a:ea typeface=""/><a:cs typeface=""/>`;
    const relleno = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
    const linea = `<a:ln w="12700">${relleno}</a:ln>`;
    const tema = D + `<a:theme xmlns:a="${A}" name="Office Theme"><a:themeElements><a:clrScheme name="Office">${colores}</a:clrScheme>`
      + `<a:fontScheme name="Office"><a:majorFont>${fuente('Aptos Display')}</a:majorFont><a:minorFont>${fuente('Aptos')}</a:minorFont></a:fontScheme>`
      + `<a:fmtScheme name="Office"><a:fillStyleLst>${relleno.repeat(3)}</a:fillStyleLst><a:lnStyleLst>${linea.repeat(3)}</a:lnStyleLst>`
      + `<a:effectStyleLst>${'<a:effectStyle><a:effectLst/></a:effectStyle>'.repeat(3)}</a:effectStyleLst><a:bgFillStyleLst>${relleno.repeat(3)}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
    const marcador = (id, nombre, tipo, x, y, cx, cy, tam) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${nombre}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${tipo}"/></p:nvPr></p:nvSpPr>`
      + `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="es-CR" sz="${tam}"/></a:p></p:txBody></p:sp>`;
    const nivel = (tam) => `<a:lvl1pPr><a:defRPr sz="${tam}"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr>`;
    const patron = D + `<p:sldMaster ${ns}>` + arbol(marcador(2, 'Título 1', 'title', 838200, 365125, 10515600, 1325563, 4400)) + mapa
      + `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle>${nivel(4400)}</p:titleStyle><p:bodyStyle>${nivel(2800)}</p:bodyStyle><p:otherStyle>${nivel(1800)}</p:otherStyle></p:txStyles></p:sldMaster>`;
    const diseno = D + `<p:sldLayout ${ns} type="title" preserve="1"><p:cSld name="Diapositiva de título"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>`
      + marcador(2, 'Título 1', 'ctrTitle', 1524000, 1122363, 9144000, 2387600, 6000) + `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
    const diapositiva = D + `<p:sld ${ns}>` + arbol() + `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
    const presentacion = D + `<p:presentation ${ns} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>`
      + `<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;
    const ct = 'application/vnd.openxmlformats-officedocument.';
    const tipos = D + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      + [['/ppt/presentation.xml', 'presentationml.presentation.main+xml'], ['/ppt/slideMasters/slideMaster1.xml', 'presentationml.slideMaster+xml'], ['/ppt/slideLayouts/slideLayout1.xml', 'presentationml.slideLayout+xml'],
        ['/ppt/slides/slide1.xml', 'presentationml.slide+xml'], ['/ppt/theme/theme1.xml', 'theme+xml']].map(([r, t]) => `<Override PartName="${r}" ContentType="${ct}${t}"/>`).join('') + '</Types>';
    const zip = new ZipPropio();
    const partes = {
      '[Content_Types].xml': tipos,
      '_rels/.rels': rels(rel('rId1', 'officeDocument', 'ppt/presentation.xml')),
      'ppt/presentation.xml': presentacion,
      'ppt/_rels/presentation.xml.rels': rels(rel('rId1', 'slideMaster', 'slideMasters/slideMaster1.xml'), rel('rId2', 'slide', 'slides/slide1.xml'), rel('rId3', 'theme', 'theme/theme1.xml')),
      'ppt/slideMasters/slideMaster1.xml': patron,
      'ppt/slideMasters/_rels/slideMaster1.xml.rels': rels(rel('rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml'), rel('rId2', 'theme', '../theme/theme1.xml')),
      'ppt/slideLayouts/slideLayout1.xml': diseno,
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels': rels(rel('rId1', 'slideMaster', '../slideMasters/slideMaster1.xml')),
      'ppt/slides/slide1.xml': diapositiva,
      'ppt/slides/_rels/slide1.xml.rels': rels(rel('rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml')),
      'ppt/theme/theme1.xml': tema,
    };
    for (const [r, t] of Object.entries(partes)) zip.file(r, t);
    return abrir(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  }

  globalThis.PresentacionEditable = { abrir, nueva, ErrorPresentacion, buscar, prefijo };
})();

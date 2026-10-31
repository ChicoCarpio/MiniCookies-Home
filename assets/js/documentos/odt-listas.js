/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Listas de un .odt para EdicionDoc (Fase 6b, etapa 6b.4). Sin DOM: corre igual en Node.
//
// En ODF las listas ENVUELVEN a sus párrafos (`text:list` > `text:list-item` > `text:p`, con listas dentro de los elementos para los niveles); el modelo las aplana
// (cada párrafo trae `lista { nivel, item, unidad }`). Cambiar la estructura (Enter, poner o quitar la lista, subir o bajar de nivel) no se puede hacer
// cortando el XML de un párrafo: se reescribe la REGIÓN completa, y solo ella: las listas de nivel superior que toca y los párrafos sueltos de en medio.
//   1. `segmentar` vuelve a armar, desde los bloques, la región como párrafos sueltos y unidades (listas) hechas de entradas { item, nivel, párrafos };
//   2. una operación (`poner`, `quitar`, `nivel`, `entrar`) cambia esas entradas y parte las unidades donde hace falta (la parte de después sigue la numeración
//      de la de antes con `text:continue-numbering`, y sin los identificadores, que no se pueden repetir);
//   3. `serializar` vuelve a escribir el XML de cada unidad con las etiquetas de apertura originales de cada elemento (`start-value`, estilos de los niveles…).
// EdicionDoc lee el XML resultante con OdtModelo.leerRegion y cambia los bloques viejos por los nuevos (una sola operación que se deshace).
// Lo que no se regenera: listas dentro de listas de otra clase de contenido (tablas, secciones dentro de un elemento), listas que continúan a otra
// (`continue-list` / `continue-numbering` al principio de la región), y la numeración de las listas de más abajo que continúan a la que cambia.
//
// Requiere xml-editar.js. `crear(modelo, { formato, prefijoDe, Rechazo })` lo llama OdtParrafo.
(function () {
  'use strict';
  const X = () => globalThis.XmlEditar;
  const URI_TEXT = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0';
  const MAX_NIVEL = 8;
  const nombreDe = (tag) => /^<([^\s>/]+)/.exec(tag)[1];
  const prefijoDeTag = (tag) => { const n = nombreDe(tag); const i = n.indexOf(':'); return i < 0 ? '' : n.slice(0, i + 1); };
  const sinIds = (tag) => ['xml:id', 'text:id', 'text:continue-list', 'text:continue-numbering'].reduce((t, a) => X().ponerAtributo(t, a.replace(/^text:/, prefijoDeTag(tag)), null), tag).replace(/\sxml:id="[^"]*"/, '');
  const conEstilo = (tag, estilo) => X().ponerAtributo(tag, prefijoDeTag(tag) + 'style-name', estilo);
  const clamp = (n) => Math.min(Math.max(n, 0), MAX_NIVEL);

  function crear(modelo, deps) {
    const { formato, Rechazo } = deps;

    // ───────── Estilos de lista nuevos ─────────
    const BULLETS = ['•', '◦', '▪'], NUMEROS = ['1', 'a', 'i'];
    function estiloLista(tipo) {
      return formato.registrarXml('MCL', 'lista|' + tipo, (nombre) => {
        const pref = formato.pref.contenido, P = (ns) => pref[ns] || ns, decl = ['style', 'fo', 'text'].filter((ns) => !pref[ns]);
        const T = P('text'), S = P('style'), F = P('fo');
        let niveles = '';
        for (let n = 1; n <= 10; n++) {
          const cm = (n * 1.27).toFixed(2) + 'cm';
          const props = `<${S}:list-level-properties ${T}:list-level-position-and-space-mode="label-alignment"><${S}:list-level-label-alignment ${T}:label-followed-by="listtab" ${T}:list-tab-stop-position="${cm}" ${F}:text-indent="-0.635cm" ${F}:margin-left="${cm}"/></${S}:list-level-properties>`;
          niveles += tipo === 'vinetas'
            ? `<${T}:list-level-style-bullet ${T}:level="${n}" ${T}:bullet-char="${BULLETS[(n - 1) % 3]}">${props}</${T}:list-level-style-bullet>`
            : `<${T}:list-level-style-number ${T}:level="${n}" ${S}:num-suffix="." ${S}:num-format="${NUMEROS[(n - 1) % 3]}">${props}</${T}:list-level-style-number>`;
        }
        const xmlns = decl.map((ns) => ` xmlns:${ns}="${{ style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0', fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0', text: URI_TEXT }[ns]}"`).join('');
        return `<${T}:list-style${xmlns} ${S}:name="${nombre}">${niveles}</${T}:list-style>`;
      });
    }
    const tagNuevo = (local, pf) => `<${pf}${local}>`;
    const etiquetaNueva = (estilo, pf) => `<${pf}list ${pf}style-name="${X().escAttr(estilo)}">`;
    const itemNuevo = (pf) => ({ tag: tagNuevo('list-item', pf), local: 'list-item', tagLista: tagNuevo('list', pf) });

    // ───────── Región → segmentos ─────────
    // Un segmento es un párrafo suelto { plano, b, xml } o una unidad { unidad, entradas: [{ item, nivel, paras, bloques }] }
    function segmentar(bloques, frag) {
      const segs = [];
      let u = null, entrada = null;
      for (const b of bloques) {
        if (!b.lista) { segs.push({ plano: true, b, xml: frag(b) }); u = null; entrada = null; continue; }
        if (!u || u.unidad !== b.lista.unidad) { u = { unidad: b.lista.unidad, tag: b.lista.unidad.tag, entradas: [] }; segs.push(u); entrada = null; }
        if (!entrada || entrada.item !== b.lista.item) { entrada = { item: b.lista.item, nivel: b.lista.nivel, paras: [], bloques: [] }; u.entradas.push(entrada); }
        entrada.paras.push(frag(b)); entrada.bloques.push(b);
      }
      return segs;
    }
    // Los niveles de una unidad: el primero a 0 y ninguno más de un escalón por debajo del anterior (una lista no puede empezar por un nivel interior)
    function normalizar(entradas) {
      const f = entradas.length ? entradas[0].nivel : 0;
      let prev = -1;
      for (const e of entradas) { e.nivel = clamp(Math.min(Math.max(e.nivel - f, 0), prev + 1)); prev = e.nivel; }
    }

    // ───────── Segmentos → XML ─────────
    function serializarUnidad(u) {
      normalizar(u.entradas);
      const raiz = { tag: u.tag, items: [] }, pila = [raiz];
      for (const e of u.entradas) {
        while (pila.length > e.nivel + 1) pila.pop();
        if (e.nivel === pila.length) {
          const padre = pila[pila.length - 1], itemPadre = padre.items[padre.items.length - 1];
          const nl = { tag: u.limpiarAnidadas ? X().ponerAtributo(e.item.tagLista, prefijoDeTag(e.item.tagLista) + 'style-name', null) : e.item.tagLista, items: [] };
          itemPadre.hijos.push(nl); pila.push(nl);
        }
        pila[e.nivel].items.push({ tag: e.item.tag, hijos: e.paras.slice() });
      }
      const ser = (l) => l.tag + l.items.map((it) => it.tag + it.hijos.map((h) => (typeof h === 'string' ? h : ser(h))).join('') + `</${nombreDe(it.tag)}>`).join('') + `</${nombreDe(l.tag)}>`;
      return ser(raiz);
    }
    function serializar(segs) { return segs.map((s) => (s.plano ? s.xml : serializarUnidad(s))).join(''); }

    // Una unidad hecha de las entradas dadas, que sigue a otra de la misma lista (continúa su numeración) o empieza de nuevo
    const unidadDe = (original, entradas, { primera = true, estilo = null, ids = true } = {}) => {
      let tag = original.tag;
      if (!primera || !ids) tag = sinIds(tag);
      if (!primera) tag = X().ponerAtributo(tag, prefijoDeTag(tag) + 'continue-numbering', 'true');
      if (estilo) tag = conEstilo(tag, estilo);
      return { unidad: original, tag, entradas, limpiarAnidadas: !!estilo };
    };
    const seleccionada = (e, objetivo) => e.bloques.some((b) => objetivo.has(b));
    // Parte una unidad en rachas seguidas de entradas elegidas / no elegidas
    function rachas(seg, objetivo) {
      const sal = [];
      for (const e of seg.entradas) {
        const sel = seleccionada(e, objetivo), ult = sal[sal.length - 1];
        if (ult && ult.sel === sel) ult.entradas.push(e); else sal.push({ sel, entradas: [e] });
      }
      return sal;
    }
    const tipoDe = (e) => { const b = e.bloques[0]; return b && b.etiqueta ? (b.etiqueta.fmt === 'bullet' ? 'vinetas' : 'numeros') : null; };

    // Comprueba que la región se puede regenerar; devuelve el motivo o null
    function revisar(segs) {
      const primeraUnidad = segs.find((s) => !s.plano);
      for (const s of segs) {
        if (s.plano) { if (s.b.protegido) return 'protegido'; continue; }
        if (s.unidad.compleja) return 'lista-compleja';
        if (s.entradas.some((e) => e.bloques.some((b) => b.protegido))) return 'protegido';
      }
      if (primeraUnidad && primeraUnidad.unidad.continua && !primeraUnidad.unidad.regenerada) return 'lista-continua';
      return null;
    }

    // ───────── Operaciones ─────────
    // Todas devuelven el XML nuevo de la región o { rechazo }
    function quitar(segs, objetivo) {
      const sal = [];
      for (const s of segs) {
        if (s.plano) { sal.push(s); continue; }
        let primera = true;
        for (const r of rachas(s, objetivo)) {
          if (r.sel) { for (const e of r.entradas) for (let k = 0; k < e.paras.length; k++) sal.push({ plano: true, b: e.bloques[k], xml: e.paras[k] }); continue; }
          sal.push(unidadDe(s.unidad, r.entradas, { primera: primera && !s.unidad.regenerada ? true : primera }));
          primera = false;
        }
      }
      return sal;
    }
    function poner(segs, objetivo, tipo) {
      const pf = prefijoDeTag(segs.find((s) => !s.plano) ? segs.find((s) => !s.plano).unidad.tag : '<text:list>');
      const estilo = estiloLista(tipo);
      const nodos = [];   // { sel, entrada | plano, origen }
      for (const s of segs) {
        if (s.plano) { nodos.push({ sel: objetivo.has(s.b), plano: s }); continue; }
        for (const e of s.entradas) nodos.push({ sel: seleccionada(e, objetivo) && tipoDe(e) !== tipo, entrada: e, origen: s });
      }
      const sal = [];
      let i = 0;
      while (i < nodos.length) {
        const n = nodos[i];
        if (n.plano && !n.sel) { sal.push(n.plano); i++; continue; }
        if (n.entrada && !n.sel) {
          // una racha de entradas de la misma unidad que se quedan como están
          const origen = n.origen, es = [];
          while (i < nodos.length && nodos[i].entrada && !nodos[i].sel && nodos[i].origen === origen) es.push(nodos[i++].entrada);
          const antes = sal.filter((x) => !x.plano && x.unidad === origen).length;
          sal.push(unidadDe(origen, es, { primera: antes === 0 }));
          continue;
        }
        // una racha de elegidas: una unidad nueva (o la anterior, si es de la misma clase y está justo antes)
        const es = [];
        while (i < nodos.length && nodos[i].sel) {
          const q = nodos[i++];
          es.push(q.plano ? { item: itemNuevo(pf), nivel: 0, paras: [q.plano.xml], bloques: [q.plano.b] } : q.entrada);
        }
        const previo = sal[sal.length - 1];
        const previoTipo = previo && !previo.plano && previo.entradas.length ? tipoDe(previo.entradas[previo.entradas.length - 1]) : null;
        if (previo && !previo.plano && previoTipo === tipo) { previo.entradas.push(...es); previo.limpiarAnidadas = true; }
        else sal.push({ unidad: { tag: etiquetaNueva(estilo, pf), regenerada: true, continua: false }, tag: etiquetaNueva(estilo, pf), entradas: es, limpiarAnidadas: true });
      }
      return sal;
    }
    function nivel(segs, objetivo, delta) {
      let cambio = false;
      for (const s of segs) {
        if (s.plano) continue;
        const antes = s.entradas.map((e) => e.nivel);
        for (const e of s.entradas) if (seleccionada(e, objetivo)) e.nivel = clamp(e.nivel + delta);
        // un nivel no puede quedar más de un escalón por debajo del anterior
        let prev = -1;
        for (const e of s.entradas) { e.nivel = Math.min(e.nivel, prev + 1); prev = e.nivel; }
        if (s.entradas.some((e, k) => e.nivel !== antes[k])) cambio = true;
      }
      return cambio ? segs : null;
    }
    // Enter: el párrafo `b` se parte en `partes` (XML de cada trozo); los trozos nuevos son elementos nuevos del mismo nivel
    function entrar(segs, b, partes) {
      const pf = prefijoDeTag((segs.find((s) => !s.plano) || { unidad: { tag: '<text:list>' } }).unidad.tag);
      for (const s of segs) {
        if (s.plano) continue;
        const k = s.entradas.findIndex((e) => e.bloques.includes(b));
        if (k < 0) continue;
        const e = s.entradas[k], j = e.bloques.indexOf(b);
        if (j > 0) { e.paras.splice(j, 1, ...partes); e.bloques.splice(j, 1, b, ...partes.slice(1).map(() => b)); return segs; }
        const resto = e.paras.slice(1), restoB = e.bloques.slice(1);
        const nuevos = partes.slice(1).map((x, q) => ({ item: itemNuevo(pf), nivel: e.nivel, paras: q === partes.length - 2 ? [x, ...resto] : [x], bloques: q === partes.length - 2 ? [b, ...restoB] : [b] }));
        e.paras = [partes[0]]; e.bloques = [b];
        s.entradas.splice(k + 1, 0, ...nuevos);
        return segs;
      }
      return null;
    }

    // Aplica una operación a los bloques de la región (en orden) y devuelve el XML nuevo de la región, o { rechazo }
    //   op: 'poner' (extra.tipo) · 'quitar' · 'nivel' (extra.delta) · 'entrar' (extra.bloque, extra.partes)
    function aplicar(op, bloques, frag, objetivo, extra = {}) {
      const segs = segmentar(bloques, frag);
      const motivo = revisar(segs);
      if (motivo) return { rechazo: motivo };
      let r;
      try {
        if (op === 'quitar') r = quitar(segs, objetivo);
        else if (op === 'poner') r = poner(segs, objetivo, extra.tipo);
        else if (op === 'nivel') r = nivel(segs, objetivo, extra.delta);
        else if (op === 'entrar') r = entrar(segs, extra.bloque, extra.partes);
      } catch (e) { if (e instanceof Rechazo) return { rechazo: e.rechazo }; throw e; }
      if (!r) return { sinCambio: true };
      return { xml: serializar(r), segs: r };
    }
    return { aplicar, segmentar, serializar, estiloLista, tipoDe };
  }

  globalThis.OdtListas = { crear };
})();

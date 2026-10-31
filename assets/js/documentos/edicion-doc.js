/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Edición de texto de Documentos (Fase 6, etapa 6.4). Sin DOM: corre igual en Node.
//
// Escribir, borrar, partir y unir párrafos sobre el modelo de docx-modelo.js, y devolver un .docx donde SOLO cambian los párrafos editados.
// La regla de la suite es la ida y vuelta sin pérdida, así que no se reescribe el párrafo desde el modelo: se corta el XML del párrafo
// (<w:p>…</w:p>) y se cambia justo lo que se tocó (el texto de un <w:t>, una corrida, el reparto en dos párrafos). Lo que el editor no entiende
// (marcadores, campos, controles, revisiones, comentarios) queda donde estaba. Después el párrafo se lee de nuevo de su XML nuevo, así el
// modelo y el archivo no pueden separarse.
//
// Una posición es { b, o }: el párrafo (el objeto del modelo) y el desplazamiento en su texto (UTF-16, el de `b.texto`).
// Cada cambio reemplaza uno o varios párrafos seguidos de una lista por otros, y esa es la operación que se deshace (HistorialDoc).
//
//   const ed = EdicionDoc.crear(modelo, { alCambiar, invalidar });
//   const fin = ed.reemplazar(desde, hasta, 'texto\notro párrafo')   // → posición donde queda el cursor
//   ed.borrarAtras(pos) · ed.borrarAdelante(pos) · ed.entrar(pos) · ed.texto(desde, hasta)
//   ed.deshacer() · ed.rehacer() → posición del cursor o null
//   await ed.guardar() → Uint8Array del .docx
//
// Requiere docx-modelo.js, xml-editar.js e historial-doc.js.
(function () {
  'use strict';
  const X = () => globalThis.XmlEditar;
  const SIN_PARTIR = /^(?:sdt|sdtContent|fldSimple|AlternateContent|Choice|Fallback|txbxContent)$/;
  const LETRA = /[\p{L}\p{N}_]/u;

  // ───────── Texto ─────────
  const esc = (s) => X().esc(s);
  // `salto`: deja pasar el salto de página (\f) que escribe la orden de insertarlo; en un texto pegado se descarta
  const limpiar = (s, salto) => String(s).replace(/\r\n?/g, '\n').replace(salto ? /[\uFFFC\u0000-\u0008\u000B\u000E-\u001F]/g : /[\f\uFFFC\u0000-\u0008\u000B\u000E-\u001F]/g, '');
  const prefijoDe = (raiz) => { const i = raiz.nombre.indexOf(':'); return i < 0 ? '' : raiz.nombre.slice(0, i + 1); };
  // El contenido de una corrida para un texto: tabuladores y saltos de línea son elementos propios.
  function contenidoDe(s, pf) {
    let sal = '';
    for (const parte of s.split(/(\t|\u2028|\f)/)) {
      if (parte === '\t') sal += `<${pf}tab/>`;
      else if (parte === '\f') sal += `<${pf}br ${pf}type="page"/>`;
      else if (parte === '\u2028') sal += `<${pf}br/>`;
      else if (parte) sal += `<${pf}t xml:space="preserve">${esc(parte)}</${pf}t>`;
    }
    return sal;
  }
  // Posición en el XML (dentro del interior de un <w:t>) del carácter número k del texto ya decodificado.
  function posRaw(raw, k) {
    let i = 0, n = 0;
    while (i < raw.length && n < k) {
      if (raw[i] === '&') { const j = raw.indexOf(';', i); if (j > 0 && j - i <= 10) { n += X().decodificar(raw.slice(i, j + 1)).length; i = j + 1; continue; } }
      i++; n++;
    }
    return i;
  }
  // Un salto de carácter respeta los pares sustitutos (emoji, escrituras raras)
  const pasoAtras = (t, o) => (o > 1 && /[\uDC00-\uDFFF]/.test(t[o - 1]) && /[\uD800-\uDBFF]/.test(t[o - 2]) ? o - 2 : o - 1);
  const pasoAdelante = (t, o) => (o + 1 < t.length && /[\uD800-\uDBFF]/.test(t[o]) && /[\uDC00-\uDFFF]/.test(t[o + 1]) ? o + 2 : o + 1);
  function palabraAntes(t, o) {
    let i = o;
    while (i > 0 && !LETRA.test(t[i - 1])) i--;
    while (i > 0 && LETRA.test(t[i - 1])) i--;
    return i;
  }
  function palabraDespues(t, o) {
    let i = o;
    while (i < t.length && LETRA.test(t[i])) i++;
    while (i < t.length && !LETRA.test(t[i])) i++;
    return i;
  }
  // La palabra bajo el cursor con los espacios que le siguen (doble clic)
  function palabraEn(t, o) {
    let a = o, f = o;
    if (o < t.length && !LETRA.test(t[o]) && o > 0 && LETRA.test(t[o - 1])) { a = o; }
    while (a > 0 && LETRA.test(t[a - 1])) a--;
    while (f < t.length && LETRA.test(t[f])) f++;
    if (f === a) { while (a > 0 && !LETRA.test(t[a - 1]) && t[a - 1] !== '\n') a--; while (f < t.length && !LETRA.test(t[f])) f++; return [a, f]; }
    while (f < t.length && /[ \t]/.test(t[f])) f++;
    return [a, f];
  }

  // ───────── Cirugía del XML de un párrafo ─────────
  // Todos los nodos del párrafo por su posición
  function nodos(raiz) {
    const mapa = new Map();
    const rec = (n) => { for (const h of n.hijos) { mapa.set(h.ini, h); rec(h); } };
    rec(raiz);
    return mapa;
  }
  // Un <w:p/> sin cierre pasa a <w:p></w:p>: así todo lo demás cuenta con iniCierre
  function normalizar(xml) {
    const r = X().analizar(xml);
    if (r.iniCierre >= 0) return xml;
    return xml.slice(r.ini, r.fin).replace(/\s*\/>$/, '>') + `</${r.nombre}>`;
  }
  const hijoDe = (n, local) => (n ? n.hijos.find((h) => h.local === local) || null : null);
  const rPrMarca = (xml, raiz) => { const r = hijoDe(hijoDe(raiz, 'pPr'), 'rPr'); return r ? xml.slice(r.ini, r.fin) : ''; };
  const textoT = (z, s) => `<${z.tn} xml:space="preserve">${esc(s)}</${z.tn}>`;
  const editable = (z) => z.t === 'texto' && z.ti !== undefined && !z.nota;

  // Ediciones que borran los caracteres [desde, hasta) del texto del párrafo
  function edicionesQuitar(raiz, piezas, desde, hasta) {
    const ed = [], porCorrida = new Map(), mapa = nodos(raiz);
    for (const z of piezas) {
      if (z.hasta <= z.desde || z.hasta <= desde || z.desde >= hasta || z.t === 'campo') continue;
      const a = Math.max(desde, z.desde) - z.desde, b = Math.min(hasta, z.hasta) - z.desde;
      if (editable(z)) {
        const nuevo = z.s.slice(0, a) + z.s.slice(b);
        if (nuevo) { ed.push({ ini: z.ei, fin: z.ef, texto: textoT(z, nuevo) }); continue; }
      }
      if (!porCorrida.has(z.ini)) porCorrida.set(z.ini, []);
      porCorrida.get(z.ini).push(z);
    }
    for (const [ini, zs] of porCorrida) {
      const run = mapa.get(ini);
      const contenido = run ? run.hijos.filter((h) => h.local !== 'rPr') : [];
      // si la corrida se queda sin nada, se va entera
      if (run && run.local === 'r' && contenido.every((h) => zs.some((z) => z.ei === h.ini))) ed.push({ ini: run.ini, fin: run.fin, texto: '' });
      else for (const z of zs) ed.push({ ini: z.ei, fin: z.ef, texto: '' });
    }
    return ed;
  }

  // Ediciones que meten `s` en el desplazamiento `pos` del texto del párrafo
  function edicionesInsertar(xml, raiz, piezas, pos, s) {
    const pf = prefijoDe(raiz), cont = contenidoDe(s, pf);
    if (!cont) return [];
    let prev = null, sig = null;
    for (const z of piezas) {
      if (z.hasta <= z.desde) continue;
      if (z.desde < pos && z.hasta >= pos) prev = z;
      if (z.desde === pos && !sig) sig = z;
      if (z.desde > pos) break;
    }
    const alvo = prev || sig;
    if (alvo && editable(alvo)) {
      const k = pos - alvo.desde, A = alvo.s.slice(0, k), B = alvo.s.slice(k);
      if (!/[\t\u2028\f]/.test(s)) return [{ ini: alvo.ei, fin: alvo.ef, texto: textoT(alvo, A + s + B) }];
      return [{ ini: alvo.ei, fin: alvo.ef, texto: (A ? textoT(alvo, A) : '') + cont + (B ? textoT(alvo, B) : '') }];
    }
    // detrás de una referencia de nota (que va en superíndice) o de un campo, el texto va en una corrida nueva y sin formato
    if (alvo && (alvo.nota || alvo.t === 'campo')) {
      const donde = prev ? alvo.fin : alvo.ini;
      return [{ ini: donde, fin: donde, texto: `<${pf}r>${cont}</${pf}r>` }];
    }
    if (prev) return [{ ini: prev.ef, fin: prev.ef, texto: cont }];
    if (sig) return [{ ini: sig.ei, fin: sig.ei, texto: cont }];
    // párrafo sin texto: corrida nueva con el formato de la marca de párrafo, justo antes del cierre
    return [{ ini: raiz.iniCierre, fin: raiz.iniCierre, texto: `<${pf}r>${rPrMarca(xml, raiz)}${cont}</${pf}r>` }];
  }

  // Los elementos abiertos (de fuera adentro) en una posición del XML, sin contar el propio párrafo
  function ancestros(raiz, corte) {
    const sal = [];
    let n = raiz;
    for (;;) {
      const h = n.hijos.find((c) => c.ini < corte && corte < c.fin);
      if (!h) break;
      sal.push(h); n = h;
    }
    return sal;
  }
  // Parte un párrafo en `pos`: devuelve los dos XML. El segundo repite las propiedades del párrafo (menos el salto de sección, que se queda
  // con el último, y el «salto de página antes»); los elementos que estaban abiertos en el corte (enlace, corrida con su formato) se cierran
  // en el primero y se vuelven a abrir en el segundo.
  function partirXml(xml, raiz, piezas, pos) {
    const util = piezas.filter((z) => z.hasta > z.desde);
    let prev = null;
    for (const z of util) if (z.desde < pos && z.hasta >= pos) prev = z;
    const sig = util.find((z) => z.desde >= pos);
    const mapa = nodos(raiz), pPr = hijoDe(raiz, 'pPr');
    let corte;
    if (prev) {
      if (editable(prev) && pos < prev.hasta) corte = prev.ti + posRaw(xml.slice(prev.ti, prev.tf), pos - prev.desde);
      else {
        const run = mapa.get(prev.ini);
        const resto = run ? run.hijos.some((h) => h.local !== 'rPr' && h.ini >= prev.ef) : false;
        corte = resto || !run ? prev.ef : run.fin;
      }
    } else if (sig) {
      const run = mapa.get(sig.ini);
      const antes = run ? run.hijos.some((h) => h.local !== 'rPr' && h.fin <= sig.ei) : false;
      corte = antes || !run ? sig.ei : run.ini;
    } else corte = pPr ? pPr.fin : raiz.finApertura;
    let pila = ancestros(raiz, corte);
    for (let vuelta = 0; vuelta < 8; vuelta++) {
      const mala = pila.find((n) => SIN_PARTIR.test(n.local));
      if (!mala) break;
      corte = mala.fin; pila = ancestros(raiz, corte);
    }
    const cierres = pila.slice().reverse().map((n) => `</${n.nombre}>`).join('');
    const abrir = (n) => {
      let ab = xml.slice(n.ini, n.finApertura);
      if (n.local === 't') ab = X().ponerAtributo(ab, 'xml:space', 'preserve');
      if (n.local === 'r') { const rPr = hijoDe(n, 'rPr'); if (rPr) ab += xml.slice(rPr.ini, rPr.fin); }
      return ab;
    };
    const sect = hijoDe(pPr, 'sectPr');
    const pf = prefijoDe(raiz);
    let antes = sect ? xml.slice(0, sect.ini) + xml.slice(sect.fin, corte) : xml.slice(0, corte);
    antes += cierres + `</${raiz.nombre}>`;
    let tag = xml.slice(raiz.ini, raiz.finApertura);
    for (const a of ['w14:paraId', 'w14:textId']) tag = X().ponerAtributo(tag, a, null);
    const pPr2 = pPr ? xml.slice(pPr.ini, pPr.fin).replace(new RegExp('<' + pf + 'pageBreakBefore\\b[^>]*?/>'), '') : '';
    const despues = tag + pPr2 + pila.map(abrir).join('') + xml.slice(corte, raiz.iniCierre) + `</${raiz.nombre}>`;
    return { antes, despues };
  }
  // Une dos párrafos: el resultado lleva las propiedades del primero, salvo que esté vacío o que el segundo cierre una sección
  function unirXml(xa, ra, xb, rb, vacioA, vacioB) {
    const pa = hijoDe(ra, 'pPr'), pb = hijoDe(rb, 'pPr');
    const sa = hijoDe(pa, 'sectPr'), sb = hijoDe(pb, 'sectPr');
    if (sa) return null;
    const usoB = !!sb || (vacioA && !vacioB);
    const cuerpo = (x, r, p) => x.slice(p ? p.fin : r.finApertura, r.iniCierre);
    const base = usoB ? [xb, rb, pb] : [xa, ra, pa];
    const inicio = base[0].slice(base[1].ini, base[2] ? base[2].fin : base[1].finApertura);
    return { xml: inicio + cuerpo(xa, ra, pa) + cuerpo(xb, rb, pb) + `</${base[1].nombre}>`, usoB };
  }
  // Cambia (o quita, con null) el estilo de párrafo de un <w:p> y devuelve el XML nuevo
  function ponerEstiloXml(xml, id) {
    const raiz = X().analizar(xml), pf = prefijoDe(raiz), pPr = hijoDe(raiz, 'pPr'), ps = hijoDe(pPr, 'pStyle');
    const el = id ? `<${pf}pStyle ${pf}val="${X().escAttr(id)}"/>` : '';
    let ed;
    if (ps) ed = [{ ini: ps.ini, fin: ps.fin, texto: el }];
    else if (!el) return xml;
    else if (pPr && pPr.iniCierre >= 0) ed = [{ ini: pPr.finApertura, fin: pPr.finApertura, texto: el }];
    else if (pPr) ed = [{ ini: pPr.ini, fin: pPr.fin, texto: `<${pf}pPr>${el}</${pf}pPr>` }];
    else ed = [{ ini: raiz.finApertura, fin: raiz.finApertura, texto: `<${pf}pPr>${el}</${pf}pPr>` }];
    return X().aplicar(xml, ed);
  }
  // Quita la numeración (<w:numPr>) de un <w:p>
  function quitarNumeracionXml(xml) {
    const raiz = X().analizar(xml), n = hijoDe(hijoDe(raiz, 'pPr'), 'numPr');
    return n ? X().aplicar(xml, [{ ini: n.ini, fin: n.fin, texto: '' }]) : xml;
  }

  const DIALECTO_DOCX = { edicionesQuitar, edicionesInsertar, partirXml, unirXml, ponerEstilo: ponerEstiloXml, quitarLista: quitarNumeracionXml };

  // ───────── El editor ─────────
  function crear(modelo, opciones = {}) {
    const alCambiar = opciones.alCambiar || (() => {});
    const invalidar = opciones.invalidar || (() => {});
    // El dialecto del formato: cómo se corta el XML de un párrafo. Word es el de este archivo; otros formatos (.odt: OdtParrafo) lo traen en modelo.dialecto.
    const dx = () => modelo.dialecto || DIALECTO_DOCX;
    const soporta = (capacidad) => !modelo.dialecto || modelo.dialecto.puede(capacidad);
    let ultimoRechazo = null;
    let lista = null, indice = null, padres = null;
    const origenes = new Map();      // lista → copia de sus bloques tal como vinieron del archivo
    let ultimoCursor = null;
    // Las tablas guardan lo que midieron: al cambiar un párrafo de una celda hay que avisar a la tabla y a las que la contienen
    const duenoDeLista = new Map(), listaDeTabla = new Map();
    function registrarTabla(b, l) {
      if (b.tipo !== 'tabla') return;
      listaDeTabla.set(b, l);
      for (const f of b.filas) for (const c of f.celdas) { duenoDeLista.set(c.bloques, b); for (const x of c.bloques) registrarTabla(x, c.bloques); }
    }
    for (const b of modelo.bloques) registrarTabla(b, modelo.bloques);

    function reconstruir() {
      lista = []; indice = new Map(); padres = new Map();
      const rec = (l) => {
        for (const b of l) {
          padres.set(b, l);
          if (b.tipo === 'p') { indice.set(b, lista.length); lista.push(b); }
          else for (const f of b.filas) for (const c of f.celdas) rec(c.bloques);
        }
      };
      rec(modelo.bloques);
    }
    const parrafos = () => { if (!lista) reconstruir(); return lista; };
    const indiceDe = (b) => { parrafos(); return indice.has(b) ? indice.get(b) : -1; };
    const contenedorDe = (b) => { parrafos(); return padres.get(b) || null; };
    const comparar = (p, q) => (p.b === q.b ? p.o - q.o : indiceDe(p.b) - indiceDe(q.b));
    const ordenar = (p, q) => (comparar(p, q) <= 0 ? [p, q] : [q, p]);
    const fragmento = (b) => normalizar((b.fuente || b.base || modelo.xml).slice(b.ini, b.fin));
    const rangoOrigen = (b) => b.orig || (b.fuente ? null : { ini: b.ini, fin: b.fin });
    // (.odt) los bloques que salen de reescribir una lista se escriben juntos: la región (ver odt-listas.js) manda en ellos
    const conRegion = (nuevo, b) => { if (b && b.region) nuevo.region = b.region; return nuevo; };

    // ── la operación que se deshace: cambia `n` bloques de una lista desde `idx` por `poner` ──
    function aplicar(op) {
      if (op.finalTipo) {
        // el tipo de inicio de la última sección (la del w:sectPr del cuerpo, que no es un bloque)
        const antes = { tipo: modelo.finalTipo === undefined ? null : modelo.finalTipo };
        if (modelo.finalContinuaOriginal === undefined) modelo.finalContinuaOriginal = modelo.seccionFinal.continua;
        modelo.finalTipo = op.finalTipo.tipo;
        modelo.seccionFinal = { ...modelo.seccionFinal, continua: op.finalTipo.tipo === null ? modelo.finalContinuaOriginal : op.finalTipo.tipo === 'continuous' };
        modelo.recalcularSecciones();
        ultimoCursor = op.cDespues || null;
        return { finalTipo: antes, cAntes: op.cDespues, cDespues: op.cAntes };
      }
      const l = op.lista;
      if (!origenes.has(l)) origenes.set(l, l.slice());
      const quitados = l.splice(op.idx, op.n, ...op.poner);
      if (l === modelo.bloques) modelo.recalcularSecciones();
      for (const b of op.poner) registrarTabla(b, l);
      lista = indice = padres = null;
      for (let t = duenoDeLista.get(l); t; t = duenoDeLista.get(listaDeTabla.get(t))) invalidar(t);
      ultimoCursor = op.cDespues || null;
      return { lista: l, idx: op.idx, n: op.poner.length, poner: quitados, cAntes: op.cDespues, cDespues: op.cAntes };
    }
    const historial = HistorialDoc.crear({
      aplicar,
      alCambiar: () => { for (const b of modelo.renumerar()) invalidar(b); alCambiar(); },
    });

    // ── leer y cambiar fragmentos ──
    const leer = (xml, base) => { xml = normalizar(xml); return { xml, ...modelo.releer(xml, base) }; };
    const conEdiciones = (F, ed, base) => (ed.length ? leer(X().aplicar(F.xml, ed), base) : F);
    // Con otro dialecto, tras cada cambio el párrafo se lee de nuevo y su texto tiene que ser el esperado; si no, la edición se rechaza
    function verificar(F, esperado) {
      if (modelo.dialecto && F.texto !== esperado) throw new modelo.dialecto.Rechazo('el texto habría salido distinto');
      return F;
    }
    const quitarDe = (F, d, h, base) => (h > d ? verificar(conEdiciones(F, dx().edicionesQuitar(F.raiz, F.piezas, d, h, F.texto), base), F.texto.slice(0, d) + F.texto.slice(h)) : F);
    const insertarEn = (F, o, s, base) => {
      if (!s) return F;
      const metido = modelo.dialecto ? s.replace(/\f/g, '') : s;
      return verificar(conEdiciones(F, dx().edicionesInsertar(F.xml, F.raiz, F.piezas, o, s, F.texto), base), F.texto.slice(0, o) + metido + F.texto.slice(o));
    };
    function partirEn(F, o, base) {
      const { antes, despues } = dx().partirXml(F.xml, F.raiz, F.piezas, o, F.texto);
      const A = leer(antes, base), B = leer(despues, base);
      verificar({ texto: A.texto + B.texto }, F.texto);
      return [A, B];
    }
    const vacio = (F) => F.texto.length === 0;
    function formatearEn(F, d, h, cambios, base) {
      const el = FormatoDoc.elementosCaracter(cambios, prefijoDe(F.raiz));
      return conEdiciones(F, FormatoDoc.edicionesCaracter(F.xml, F.raiz, F.piezas, d, h, el), base);
    }

    function bloqueDe(modeloB, F, orig) {
      return {
        ...modeloB, fuente: F.xml, ini: 0, fin: F.xml.length, orig, piezas: F.piezas, texto: F.texto, seccion: F.seccion,
        etiqueta: modeloB.etiqueta ? { ...modeloB.etiqueta } : null,
      };
    }

    // Un grupo: párrafos seguidos de una misma lista; se cambia el texto desde `s` (en el primero) hasta `e` (en el último), y se mete `lineas`.
    // Cada parte es un párrafo nuevo { F (su XML leído), mb (el bloque del que toma el formato) }; la primera ocupa el lugar del primer párrafo.
    function procesarGrupo(g, s, e, lineas, opc) {
      const g0 = g[0], gN = g[g.length - 1], base = g0.r;
      let A = leer(fragmento(g0), base), partes;
      if (g.length === 1) partes = [{ F: quitarDe(A, s, e, base), mb: g0 }];
      else {
        A = quitarDe(A, s, A.texto.length, base);
        let B = leer(fragmento(gN), gN.r);
        B = quitarDe(B, 0, e, gN.r);
        const u = dx().unirXml(A.xml, A.raiz, B.xml, B.raiz, vacio(A), vacio(B));
        partes = u ? [{ F: verificar(leer(u.xml, base), A.texto + B.texto), mb: u.usoB ? gN : g0 }] : [{ F: A, mb: g0 }, { F: B, mb: gN }];
      }
      let caret = null;
      if (lineas) {
        const mb = partes[0].mb;
        let cur = partes[0].F, o = s;
        const hechos = [{ F: cur, mb }];
        lineas.forEach((linea, i) => {
          if (i > 0) {
            if (cur.texto.length < o) o = cur.texto.length;
            const [x, y] = partirEn(cur, o, base);
            hechos[hechos.length - 1].F = x;
            cur = y; o = 0;
            hechos.push({ F: y, mb, siguiente: !!opc.estiloSiguiente && !y.texto });
          }
          if (linea) { cur = insertarEn(cur, o, linea, base); if (opc.formato) cur = formatearEn(cur, o, o + linea.length, opc.formato, base); hechos[hechos.length - 1].F = cur; hechos[hechos.length - 1].siguiente = false; o += linea.length; }
        });
        partes.splice(0, 1, ...hechos);
        caret = { parte: hechos.length - 1, o };
      }
      const bloques = partes.map((p, i) => {
        const orig = i === 0 ? (g0.orig || (g0.fuente ? null : { ini: g0.ini, fin: g0.fin })) : null;
        const sig = p.siguiente ? modelo.m.estilos.siguiente(g0.estilo) : null;
        if (sig && sig !== g0.estilo) {
          const xml = dx().ponerEstilo(p.F.xml, sig);
          return conRegion({ ...modelo.leerCompleto(xml, g0.medio, g0), fuente: xml, ini: 0, fin: xml.length, orig }, g0);
        }
        return bloqueDe(p.mb, p.F, orig);
      });
      return { bloques, caret };
    }

    // ── reemplazar un rango por texto ──
    function reemplazar(desde, hasta, texto, opc = {}) {
      let [p0, p1] = ordenar(desde, hasta);
      texto = limpiar(texto || '', opc.salto);
      if (p0.b === p1.b && p0.o === p1.o && !texto) return p0;
      if (indiceDe(p0.b) < 0 || indiceDe(p1.b) < 0) return p0;
      const todos = parrafos(), i0 = indice.get(p0.b), i1 = indice.get(p1.b);
      ultimoRechazo = null;
      if (modelo.dialecto) {
        const rango = todos.slice(i0, i1 + 1), nLineas = texto.split('\n').length;
        const motivo = modelo.dialecto.rechazar(rango, i1 > i0 ? 'unir' : nLineas > 1 ? 'partir' : null);
        if (motivo === 'lista' && i0 === i1 && nLineas > 1 && soporta('listas')) return partirEnLista(p0, p1, texto, opc);
        if (motivo) { ultimoRechazo = motivo; return p0; }
      }
      // grupos de párrafos seguidos que comparten lista
      // las tablas que el rango cubre enteras se quitan (como en Word); los párrafos de fuera que quedan juntos se unen
      const quitar = tablasCubiertas(p0.b, p1.b, i0, i1);
      const grupos = [];
      for (let i = i0; i <= i1; i++) {
        const b = todos[i];
        if (quitar.size && bajoTabla(b, quitar)) continue;
        const l = padres.get(b), ult = grupos[grupos.length - 1];
        if (ult && ult.l === l) ult.g.push(b); else grupos.push({ l, g: [b] });
      }
      const ops = [], lineas = texto.split('\n'), opsGrupos = [];
      let caret = null;
      try {
      grupos.forEach((G, k) => {
        const primero = k === 0, ultimo = k === grupos.length - 1;
        const s = primero ? p0.o : 0, e = ultimo ? p1.o : G.g[G.g.length - 1].texto.length;
        const r = procesarGrupo(G.g, s, e, primero ? lineas : null, opc);
        // el índice cuenta la lista sin las tablas que se van a quitar
        const idx = G.l.filter((x) => !quitar.has(x)).indexOf(G.g[0]);
        if (primero) caret = { b: r.bloques[r.caret.parte], o: r.caret.o };
        opsGrupos.push({ lista: G.l, idx, n: G.g.length, poner: r.bloques });
      });
      } catch (e) { if (e && e.rechazo) { ultimoRechazo = e.rechazo; return p0; } throw e; }
      opsGrupos.reverse();   // de abajo hacia arriba: los índices de arriba siguen valiendo
      const quitadas = [...quitar].map((t) => ({ t, l: listaDeTabla.get(t) })).sort((a, b) => b.l.indexOf(b.t) - a.l.indexOf(a.t));
      for (const q of quitadas) ops.push({ lista: q.l, idx: q.l.indexOf(q.t), n: 1, poner: [] });
      ops.push(...opsGrupos);
      const cAntes = { b: p0.b, o: p0.o };
      ops.forEach((op) => { op.cAntes = cAntes; op.cDespues = caret; });
      historial.hacer(opc.nombre || 'Escribir', ops, opc.fusionar ? { fusionar: opc.fusionar } : undefined);
      return caret;
    }

    // ── órdenes de edición ──
    function textoEntre(desde, hasta) {
      const [p0, p1] = ordenar(desde, hasta);
      const todos = parrafos(), i0 = indiceDe(p0.b), i1 = indiceDe(p1.b);
      if (i0 < 0 || i1 < 0) return '';
      const trozo = (b, a, f) => b.texto.slice(a, f).replace(/\uFFFC|\f/g, '').replace(/\u2028/g, '\n');
      if (i0 === i1) return trozo(p0.b, p0.o, p1.o);
      const sal = [trozo(p0.b, p0.o, p0.b.texto.length)];
      for (let i = i0 + 1; i < i1; i++) sal.push(trozo(todos[i], 0, todos[i].texto.length));
      sal.push(trozo(p1.b, 0, p1.o));
      return sal.join('\n');
    }
    function anterior(p) {
      if (p.o > 0) return { b: p.b, o: pasoAtras(p.b.texto, p.o) };
      const i = indiceDe(p.b);
      return i > 0 ? { b: parrafos()[i - 1], o: parrafos()[i - 1].texto.length } : p;
    }
    function siguiente(p) {
      if (p.o < p.b.texto.length) return { b: p.b, o: pasoAdelante(p.b.texto, p.o) };
      const i = indiceDe(p.b);
      return i >= 0 && i < parrafos().length - 1 ? { b: parrafos()[i + 1], o: 0 } : p;
    }
    // Retroceso: borra el carácter anterior; al principio de un párrafo de lista, primero le quita la lista; si no, lo une con el anterior
    function borrarAtras(p, porPalabra) {
      if (p.o === 0) {
        if (p.b.numPedida) return quitarLista(p);
        const i = indiceDe(p.b);
        if (i <= 0) return p;
        const prev = parrafos()[i - 1];
        if (padres.get(prev) !== padres.get(p.b)) return p;
        if (prev.seccion) return p;
        return reemplazar({ b: prev, o: prev.texto.length }, p, '', { nombre: 'Borrar', fusionar: 'borrar' });
      }
      const o = porPalabra ? palabraAntes(p.b.texto, p.o) : pasoAtras(p.b.texto, p.o);
      return reemplazar({ b: p.b, o }, p, '', { nombre: 'Borrar', fusionar: 'borrar' });
    }
    function borrarAdelante(p, porPalabra) {
      const t = p.b.texto;
      if (p.o >= t.length) {
        const i = indiceDe(p.b), sig = parrafos()[i + 1];
        if (!sig || padres.get(sig) !== padres.get(p.b) || p.b.seccion) return p;
        return reemplazar(p, { b: sig, o: 0 }, '', { nombre: 'Borrar', fusionar: 'borrar' });
      }
      const o = porPalabra ? palabraDespues(t, p.o) : pasoAdelante(t, p.o);
      return reemplazar(p, { b: p.b, o }, '', { nombre: 'Borrar', fusionar: 'borrar' });
    }
    // Quita la viñeta o el número de un párrafo sin tocar su texto
    function quitarLista(p) {
      if (!soporta('listas')) { ultimoRechazo = 'lista'; return p; }
      if (modelo.dialecto) { const r = operarLista('quitar', p, p, {}, 'Quitar lista'); return r ? r.cursor : p; }
      const b = p.b, xml = quitarNumeracionXml(fragmento(b));
      const completo = modelo.leerCompleto(xml, b.medio);
      const nuevo = { ...completo, fuente: xml, ini: 0, fin: xml.length, orig: b.orig || (b.fuente ? null : { ini: b.ini, fin: b.fin }) };
      const cp = { b: nuevo, o: p.o };
      historial.hacer('Quitar lista', [{ lista: contenedorDe(b), idx: contenedorDe(b).indexOf(b), n: 1, poner: [nuevo], cAntes: p, cDespues: cp }]);
      return cp;
    }
    // Enter: parte el párrafo; en un elemento de lista vacío, sale de la lista; al final de un título, el siguiente toma el estilo «siguiente»
    function entrar(desde, hasta) {
      const [p0, p1] = ordenar(desde, hasta || desde);
      if (p0.b === p1.b && p0.o === p1.o && !p0.b.texto && p0.b.numPedida) return quitarLista(p0);
      const alFinal = p0.b === p1.b && p0.o === p1.o && p0.o === p0.b.texto.length;
      return reemplazar(p0, p1, '\n', { nombre: 'Escribir', estiloSiguiente: alFinal });
    }
    // ───────── Formato (6.5) ─────────
    const D = () => globalThis.DocxBase;
    // Un bloque nuevo a partir del XML cambiado de un párrafo: se lee entero otra vez (formato efectivo, etiqueta de lista, sección)
    function bloqueDesdeXml(b, xml) {
      return conRegion({ ...modelo.leerCompleto(xml, b.medio, b), fuente: xml, ini: 0, fin: xml.length, orig: rangoOrigen(b) }, b);
    }
    // ── listas de un .odt (odt-listas.js): se reescribe la región de listas y párrafos sueltos que toca, de una vez y con un solo deshacer ──
    // op: 'poner' ({ tipo }) · 'quitar' · 'nivel' ({ delta }) · 'entrar' ({ bloque, partes }). Devuelve { cursor, nuevos, bloquesViejos } o null (y `rechazo()` dice por qué).
    function operarLista(op, desde, hasta, extra, nombre, opciones = {}) {
      ultimoRechazo = null;
      const [p0, p1] = ordenar(desde, hasta);
      const todos = parrafos(), i0 = indiceDe(p0.b), i1 = indiceDe(p1.b);
      if (i0 < 0 || i1 < 0) return null;
      const rango = todos.slice(i0, i1 + 1), l = padres.get(p0.b);
      if (rango.some((b) => padres.get(b) !== l)) { ultimoRechazo = 'lista'; return null; }
      const clave = (b) => b.region || (b.lista && b.lista.unidad) || null;
      let i = l.indexOf(p0.b), j = l.indexOf(p1.b);
      const crecer = () => {
        for (let cambio = true; cambio;) {
          cambio = false;
          while (i > 0 && l[i - 1].tipo === 'p' && clave(l[i]) && clave(l[i - 1]) === clave(l[i])) { i--; cambio = true; }
          while (j < l.length - 1 && l[j + 1].tipo === 'p' && clave(l[j]) && clave(l[j + 1]) === clave(l[j])) { j++; cambio = true; }
        }
      };
      crecer();
      // una lista nueva justo debajo de otra de la misma clase la continúa
      if (op === 'poner' && i > 0 && l[i - 1].tipo === 'p' && l[i - 1].lista) {
        let k = i - 1;
        while (k > 0 && !l[k].etiqueta && l[k - 1].tipo === 'p' && clave(l[k - 1]) === clave(l[k])) k--;
        const e = l[k].etiqueta;
        if (e && (e.fmt === 'bullet' ? 'vinetas' : 'numeros') === extra.tipo) { i--; crecer(); }
      }
      const bloques = l.slice(i, j + 1);
      if (bloques.some((b) => b.tipo !== 'p')) { ultimoRechazo = 'lista'; return null; }
      const r = modelo.dialecto.operacionLista(op, bloques, fragmento, new Set(rango), extra);
      if (r.rechazo) { ultimoRechazo = r.rechazo; return null; }
      if (r.sinCambio) return null;
      const lect = modelo.leerRegion(r.xml, bloques[0].medio), nuevos = lect.bloques;
      const esperados = opciones.esperadosDe ? opciones.esperadosDe(bloques) : bloques.map((b) => b.texto);
      if (nuevos.length !== esperados.length || nuevos.some((b, k) => b.texto !== esperados[k])) { ultimoRechazo = 'el texto habría salido distinto'; return null; }
      // los trozos del archivo que esta región reemplaza: las listas enteras, los párrafos sueltos y lo que ya cubría otra región
      const rangos = [], vistos = new Set();
      for (const b of bloques) {
        if (b.region) { if (!vistos.has(b.region)) { vistos.add(b.region); rangos.push(...b.region.rangos); } continue; }
        if (b.lista) { const u = b.lista.unidad; if (u && !vistos.has(u)) { vistos.add(u); rangos.push(u.rango); } continue; }
        const rg = rangoOrigen(b);
        if (rg) rangos.push(rg);
      }
      rangos.sort((a, b) => a.ini - b.ini);
      const R = { xml: lect.base, iniciales: nuevos.slice(), rangos };
      const col = bloques[0].columnasSeccion, mismaCol = col && bloques.every((q) => q.columnasSeccion === col);
      for (const b of nuevos) { b.region = R; if (mismaCol) b.columnasSeccion = col; }
      const c = opciones.caretDe ? opciones.caretDe(bloques) : { idx: bloques.indexOf(p0.b), o: p0.o };
      const cAntes = { b: p0.b, o: p0.o }, cDespues = { b: nuevos[c.idx], o: Math.min(c.o, nuevos[c.idx].texto.length) };
      historial.hacer(nombre, [{ lista: l, idx: i, n: j - i + 1, poner: nuevos, cAntes, cDespues }]);
      return { cursor: cDespues, nuevos, bloquesViejos: bloques };
    }
    // La misma posición, en los bloques nuevos de una región reescrita (los párrafos siguen en el mismo orden)
    const reubicar = (r) => (p) => ({ b: r.nuevos[r.bloquesViejos.indexOf(p.b)] || p.b, o: p.o });
    // Enter (o texto con saltos de línea) dentro de un elemento de lista: cada línea nueva es un elemento nuevo del mismo nivel
    function partirEnLista(p0, p1, texto, opc) {
      const b = p0.b;
      let r;
      try { r = procesarGrupo([b], p0.o, p1.o, texto.split('\n'), { ...opc, estiloSiguiente: false }); } catch (e) { if (e && e.rechazo) { ultimoRechazo = e.rechazo; return p0; } throw e; }
      const res = operarLista('entrar', p0, p0, { bloque: b, partes: r.bloques.map((x) => x.fuente) }, 'Escribir', {
        esperadosDe: (bs) => bs.flatMap((q) => (q === b ? r.bloques.map((x) => x.texto) : [q.texto])),
        caretDe: (bs) => ({ idx: bs.indexOf(b) + r.caret.parte, o: r.caret.o }),
      });
      return res ? res.cursor : p0;
    }

    // El formato de párrafo en el XML del párrafo: Word en <w:pPr>, ODF con un estilo automático derivado
    const ponerPPr = (b, xml, cambios) => (modelo.dialecto ? modelo.dialecto.formatoParrafo(xml, cambios) : FormatoDoc.ponerPPr(xml, cambios));
    // Cambia el XML de cada párrafo del rango con `fn(b, xml, s, e)` (s y e: la parte del texto del párrafo que cae en el rango); null = sin cambio.
    // Cada párrafo cambiado es una operación que se deshace; devuelve el rango con los bloques nuevos, o null si nada cambió.
    function cambiarParrafos(desde, hasta, nombre, fn) {
      if (!soporta('formato')) return null;
      const [p0, p1] = ordenar(desde, hasta);
      const todos = parrafos(), i0 = indiceDe(p0.b), i1 = indiceDe(p1.b);
      if (i0 < 0 || i1 < 0) return null;
      const ops = [], nuevos = new Map();
      ultimoRechazo = null;
      try {
        for (let i = i0; i <= i1; i++) {
          const b = todos[i], original = fragmento(b);
          if (modelo.dialecto && b.protegido) { ultimoRechazo = 'protegido'; return null; }
          const xml = fn(b, original, b === p0.b ? p0.o : 0, b === p1.b ? p1.o : b.texto.length);
          if (!xml || xml === original) continue;
          const nuevo = bloqueDesdeXml(b, xml), l = padres.get(b);
          if (modelo.dialecto) verificar(nuevo, b.texto);
          nuevos.set(b, nuevo);
          ops.push({ lista: l, idx: l.indexOf(b), n: 1, poner: [nuevo] });
        }
      } catch (e) { if (e && e.rechazo) { ultimoRechazo = e.rechazo; return null; } throw e; }
      if (!ops.length) return null;
      const re = (p) => ({ b: nuevos.get(p.b) || p.b, o: p.o });
      const cAntes = { b: p0.b, o: p0.o }, cDespues = re(p0);
      ops.forEach((op) => { op.cAntes = cAntes; op.cDespues = cDespues; });
      historial.hacer(nombre, ops);
      return { desde: re(desde), hasta: re(hasta) };
    }
    // Formato de caracteres del rango (ver FormatoDoc.elementosCaracter). Un párrafo vacío lo guarda en su marca, para lo que se escriba después.
    function formatoCaracter(desde, hasta, cambios, nombre) {
      return cambiarParrafos(desde, hasta, nombre || 'Formato', (b, xml, s, e) => {
        const F = leer(xml, b.r);
        if (modelo.dialecto) return e > s ? modelo.dialecto.formatoCaracter(F, s, e, cambios, b.r) : null;
        const el = FormatoDoc.elementosCaracter(cambios, prefijoDe(F.raiz));
        if (e > s) { const eds = FormatoDoc.edicionesCaracter(F.xml, F.raiz, F.piezas, s, e, el); return eds.length ? X().aplicar(F.xml, eds) : null; }
        return !b.texto.length ? FormatoDoc.ponerMarca(xml, el) : null;
      });
    }
    // Lo que valen en común los caracteres del rango (o el que está junto al cursor): true/false o el valor; null si cambia dentro del rango
    function estadoFormato(desde, hasta) {
      const [p0, p1] = ordenar(desde, hasta);
      const todos = parrafos(), i0 = indiceDe(p0.b), i1 = indiceDe(p1.b);
      if (i0 < 0 || i1 < 0) return null;
      const rs = [], bs = [];
      const colapsado = p0.b === p1.b && p0.o === p1.o;
      for (let i = i0; i <= i1; i++) {
        const b = todos[i], s = b === p0.b ? p0.o : 0, e = b === p1.b ? p1.o : b.texto.length;
        bs.push(b);
        if (colapsado) {
          const z = b.piezas.find((q) => q.hasta > q.desde && q.desde < s && q.hasta >= s && q.t === 'texto') || b.piezas.find((q) => q.hasta > q.desde && q.desde >= s && q.t === 'texto');
          rs.push(z ? z.r : b.marca || b.r);
        } else if (e > s) { for (const z of b.piezas) if (z.t === 'texto' && z.hasta > z.desde && z.desde < e && z.hasta > s) rs.push(z.r); }
        else if (!b.texto.length) rs.push(b.marca || b.r);
      }
      const m = modelo.m;
      const comun = (arr, f) => { if (!arr.length) return undefined; const v = arr.map(f); return v.every((x) => x === v[0]) ? v[0] : null; };
      const visual = (b) => {
        const j = b.p.jc, der = !!b.p.derechaAIzquierda;
        const v = j === 'both' || j === 'distribute' || j === 'justify' ? 'both' : j === 'center' ? 'center' : (j === 'right' || j === 'end') === !der ? 'right' : 'left';
        return v;
      };
      const lista = (b) => (b.etiqueta && b.numPedida ? (b.etiqueta.fmt === 'bullet' ? 'vinetas' : 'numeros') : null);
      return {
        b: comun(rs, (r) => !!r.b), i: comun(rs, (r) => !!r.i), u: comun(rs, (r) => !!(r.u && r.u !== 'none')), tachado: comun(rs, (r) => !!r.tachado),
        color: comun(rs, (r) => m.color(r.color)), tam: comun(rs, (r) => r.tam), fuente: comun(rs, (r) => m.fuente(r.fuentes)),
        resaltado: comun(rs, (r) => r.resaltado || null), vertical: comun(rs, (r) => r.alineacionVertical || null),
        jc: comun(bs, visual), estilo: comun(bs, (b) => (modelo.dialecto ? modelo.dialecto.estiloComun(b.estilo) : b.estilo || null)), lista: comun(bs, lista), nivel: comun(bs, (b) => (b.numPedida ? parseInt(b.numPedida.nivel || '0', 10) : null)),
        interlineado: comun(bs, (b) => (b.p.espacio && b.p.espacio.linea !== undefined && (b.p.espacio.regla || 'auto') === 'auto' ? b.p.espacio.linea / 240 : 1)),
        izq: comun(bs, (b) => (b.p.sangria && b.p.sangria.izq) || 0),
      };
    }
    // Pone o quita negrita, cursiva, subrayado o tachado: si todo el rango ya lo tiene, lo quita; si no, lo pone en todo (como Word)
    function alternarCaracter(desde, hasta, prop) {
      const est = estadoFormato(desde, hasta);
      if (!est) return null;
      const poner = est[prop] !== true;
      return formatoCaracter(desde, hasta, { [prop]: prop === 'u' ? (poner ? 'single' : 'none') : poner }, { b: 'Negrita', i: 'Cursiva', u: 'Subrayado', tachado: 'Tachado' }[prop]);
    }
    // Formato de párrafo (ver FormatoDoc.ponerPPr)
    function formatoParrafo(desde, hasta, cambios, nombre) {
      return cambiarParrafos(desde, hasta, nombre || 'Formato de párrafo', (b, xml) => ponerPPr(b, xml, cambios));
    }
    // Alineación como se ve: en un párrafo de derecha a izquierda «jc=left» es el borde derecho, así que se invierte
    function alinear(desde, hasta, visual) {
      return cambiarParrafos(desde, hasta, 'Alinear', (b, xml) => {
        const der = !!b.p.derechaAIzquierda;
        const jc = visual === 'left' ? (der ? 'right' : 'left') : visual === 'right' ? (der ? 'left' : 'right') : visual;
        return ponerPPr(b, xml, { jc });
      });
    }
    // Sangría: en una lista cambia el nivel (como Tab); fuera de una lista mueve la sangría izquierda media pulgada
    function sangrar(desde, hasta, delta) {
      if (modelo.dialecto) {
        const [p0, p1] = ordenar(desde, hasta), todos = parrafos();
        if (todos.slice(indiceDe(p0.b), indiceDe(p1.b) + 1).some((b) => b.numPedida)) {
          const r = operarLista('nivel', desde, hasta, { delta }, delta > 0 ? 'Aumentar sangría' : 'Disminuir sangría');
          return r ? { desde: reubicar(r)(desde), hasta: reubicar(r)(hasta) } : null;
        }
      }
      return cambiarParrafos(desde, hasta, delta > 0 ? 'Aumentar sangría' : 'Disminuir sangría', (b, xml) => {
        if (b.numPedida) {
          const nivel = acotarNivel(parseInt(b.numPedida.nivel || '0', 10) + delta);
          return String(nivel) === String(b.numPedida.nivel || '0') ? null : FormatoDoc.ponerPPr(xml, { numeracion: { numId: b.numPedida.numId, nivel } });
        }
        const izq = Math.max(0, ((b.p.sangria && b.p.sangria.izq) || 0) + 36 * delta);
        return ponerPPr(b, xml, { sangria: { izq } });
      });
    }
    const acotarNivel = (n) => Math.min(Math.max(n, 0), 8);
    // Tab al principio de un elemento de lista baja un nivel; Mayús+Tab sube uno
    function nivelDeLista(desde, hasta, delta) {
      const [p0] = ordenar(desde, hasta);
      if (!p0.b.numPedida) return null;
      return sangrar(desde, hasta, delta);
    }
    function interlineado(desde, hasta, factor) {
      return formatoParrafo(desde, hasta, { espacio: { linea: Math.round(factor * 240), regla: 'auto' } }, 'Interlineado');
    }
    function espaciado(desde, hasta, antes, despues) {
      const e = {};
      if (antes !== undefined) e.antes = antes;
      if (despues !== undefined) e.despues = despues;
      return formatoParrafo(desde, hasta, { espacio: e }, 'Espaciado');
    }

    // ── partes del paquete que crecen: estilos y numeración ──
    const rutaEstilos = () => (D() ? D().porTipo(modelo.rels, '/styles') : null) || 'word/styles.xml';
    const rutaNumeracion = () => (D() ? D().porTipo(modelo.rels, '/numbering') : null) || 'word/numbering.xml';
    const sinCierre = (xml) => { const r = X().analizar(xml); return r.iniCierre >= 0 ? xml : xml.slice(r.ini, r.fin).replace(/\s*\/>$/, '>') + `</${r.nombre}>`; };
    // Los estilos de párrafo del documento: { id, nombre (el de Word), predeterminado }
    async function leerEstilos() {
      const xml = await modelo.paquete.leer(rutaEstilos());
      if (xml === null) return [];
      return X().analizar(xml).hijos.filter((h) => h.local === 'style' && X().attr(h, 'type') === 'paragraph').map((h) => {
        const n = h.hijos.find((q) => q.local === 'name');
        return {
          id: X().attr(h, 'styleId'), nombre: n ? X().attr(n, 'val') || '' : '', predeterminado: X().attr(h, 'default') === '1' || X().attr(h, 'default') === 'true',
          oculto: h.hijos.some((q) => q.local === 'semiHidden' || q.local === 'hidden') && !h.hijos.some((q) => q.local === 'unhideWhenUsed' || q.local === 'qFormat'),
          rapido: h.hijos.some((q) => q.local === 'qFormat'),
        };
      }).filter((s) => s.id);
    }
    const NOMBRES_ES = { normal: 'Normal', 'heading 1': 'Título 1', 'heading 2': 'Título 2', 'heading 3': 'Título 3', 'heading 4': 'Título 4', title: 'Título', subtitle: 'Subtítulo', 'list paragraph': 'Párrafo de lista', quote: 'Cita', 'no spacing': 'Sin espaciado' };
    // Los estilos que ofrece el selector: los del documento y los predefinidos que aún no tiene (se crean al usarlos)
    async function estilosDisponibles() {
      if (modelo.dialecto) return modelo.dialecto.estilosDisponibles();
      const del = await leerEstilos(), sal = [];
      const porNombre = (n) => del.find((s) => s.nombre.toLowerCase() === n);
      for (const s of del) if (!s.oculto && (s.rapido || s.predeterminado)) sal.push({ id: s.id, nombre: NOMBRES_ES[s.nombre.toLowerCase()] || s.nombre, predeterminado: s.predeterminado });
      for (const id of FormatoDoc.ESTILOS_PREDEFINIDOS) {
        const e = FormatoDoc.ESTILOS_PREDEFINIDOS && FormatoDoc.nombreEs(id);
        const ingles = { Heading1: 'heading 1', Heading2: 'heading 2', Heading3: 'heading 3', Title: 'title', Subtitle: 'subtitle' }[id];
        if (ingles && !porNombre(ingles) && id !== 'ListParagraph') sal.push({ id, nombre: e, nuevo: true });
      }
      const orden = ['normal', 'title', 'subtitle', 'heading 1', 'heading 2', 'heading 3'];
      const clave = (s) => { const i = orden.indexOf((del.find((q) => q.id === s.id) || { nombre: { Title: 'title', Subtitle: 'subtitle', Heading1: 'heading 1', Heading2: 'heading 2', Heading3: 'heading 3' }[s.id] || '' }).nombre.toLowerCase()); return i < 0 ? 99 : i; };
      return sal.sort((a, b) => clave(a) - clave(b));
    }
    // El identificador real de un estilo predefinido (Word en español llama «Ttulo1» al «heading 1»); lo crea si el documento no lo tiene
    async function asegurarEstilo(id) {
      const ingles = FormatoDoc.estiloNuevo(id) ? /w:name w:val="([^"]+)"/.exec(FormatoDoc.estiloNuevo(id))[1].toLowerCase() : null;
      const del = await leerEstilos();
      if (del.some((s) => s.id === id)) return id;
      const igual = ingles && del.find((s) => s.nombre.toLowerCase() === ingles);
      if (igual) return igual.id;
      const nuevo = FormatoDoc.estiloNuevo(id);
      if (!nuevo) return null;
      const ruta = rutaEstilos();
      let xml = await modelo.paquete.leer(ruta);
      if (xml === null) return null;
      xml = sinCierre(xml);
      const raiz = X().analizar(xml);
      xml = X().aplicar(xml, [{ ini: raiz.iniCierre, fin: raiz.iniCierre, texto: nuevo }]);
      modelo.paquete.escribir(ruta, xml);
      await modelo.recargar();
      return id;
    }
    async function ponerEstilo(desde, hasta, id) {
      if (!soporta('estilos')) return null;
      if (modelo.dialecto) {
        const nombre = id ? modelo.dialecto.asegurarEstilo(id) : modelo.dialecto.estiloPredeterminado();
        if (id && !nombre && id !== 'Standard') return null;
        return cambiarParrafos(desde, hasta, 'Estilo', (b, xml) => (b.abreSeccion ? null : modelo.dialecto.estiloConservando(xml, nombre)));
      }
      const real = id ? await asegurarEstilo(id) : null;
      if (id && !real) return null;
      const def = (await leerEstilos()).find((s) => s.predeterminado);
      const final = real && def && real === def.id ? null : real;
      return cambiarParrafos(desde, hasta, 'Estilo', (b, xml) => ponerEstiloXml(xml, final));
    }
    // Una lista nueva de ese tipo: devuelve su numId. Las viñetas reutilizan siempre la misma; los números empiezan en 1.
    let numViñetas = null;
    async function nuevaLista(tipo) {
      if (tipo === 'vinetas' && numViñetas) return numViñetas;
      const ruta = rutaNumeracion();
      let xml = await modelo.paquete.leer(ruta), parteNueva = false;
      if (xml === null) { xml = FormatoDoc.NUMERACION_VACIA; parteNueva = true; }
      xml = sinCierre(xml);
      let raiz = X().analizar(xml);
      const maximo = (local, attr) => raiz.hijos.filter((h) => h.local === local).reduce((m, h) => Math.max(m, parseInt(X().attr(h, attr), 10) || 0), 0);
      const abId = maximo('abstractNum', 'abstractNumId') + 1, numId = maximo('num', 'numId') + 1;
      const primeroNum = raiz.hijos.find((h) => h.local === 'num' || h.local === 'numIdMacAtCleanup');
      const pos = primeroNum ? primeroNum.ini : raiz.iniCierre;
      xml = X().aplicar(xml, [{ ini: pos, fin: pos, texto: FormatoDoc.numeracionNueva(tipo, abId) }]);
      raiz = X().analizar(xml);
      const cierre = raiz.hijos.find((h) => h.local === 'numIdMacAtCleanup'), donde = cierre ? cierre.ini : raiz.iniCierre;
      const reinicio = tipo === 'numeros' ? '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>' : '';
      xml = X().aplicar(xml, [{ ini: donde, fin: donde, texto: `<w:num w:numId="${numId}"><w:abstractNumId w:val="${abId}"/>${reinicio}</w:num>` }]);
      modelo.paquete.escribir(ruta, xml);
      if (parteNueva) modelo.numeracionNueva = ruta;
      await modelo.recargar();
      if (tipo === 'vinetas') numViñetas = String(numId);
      return String(numId);
    }
    // Viñetas o números: si todos los párrafos del rango ya son de ese tipo, quita la lista; si no, los pone
    async function ponerLista(desde, hasta, tipo) {
      if (!soporta('listas')) return null;
      const [p0, p1] = ordenar(desde, hasta);
      const todos = parrafos(), i0 = indiceDe(p0.b), i1 = indiceDe(p1.b);
      if (i0 < 0 || i1 < 0) return null;
      const rango = todos.slice(i0, i1 + 1);
      const esTipo = (b) => !!(b.numPedida && b.etiqueta && (b.etiqueta.fmt === 'bullet') === (tipo === 'vinetas'));
      if (modelo.dialecto) {
        const r = rango.every(esTipo) ? operarLista('quitar', desde, hasta, {}, 'Quitar lista') : operarLista('poner', desde, hasta, { tipo }, tipo === 'vinetas' ? 'Viñetas' : 'Numeración');
        return r ? { desde: reubicar(r)(desde), hasta: reubicar(r)(hasta) } : null;
      }
      const parrafoLista = await asegurarEstilo('ListParagraph');
      if (rango.every(esTipo)) {
        return cambiarParrafos(desde, hasta, 'Quitar lista', (b, xml) => {
          let x = quitarNumeracionXml(xml);
          if (modelo.leerCompleto(x, b.medio).numPedida) x = FormatoDoc.ponerPPr(x, { numeracion: { numId: '0', nivel: 0 } });
          return parrafoLista && b.estilo === parrafoLista ? ponerEstiloXml(x, null) : x;
        });
      }
      const i = indiceDe(p0.b), prev = i > 0 && padres.get(todos[i - 1]) === padres.get(p0.b) ? todos[i - 1] : null;
      const numId = prev && esTipo(prev) ? String(prev.numPedida.numId) : await nuevaLista(tipo);
      return cambiarParrafos(desde, hasta, tipo === 'vinetas' ? 'Viñetas' : 'Numeración', (b, xml) => {
        if (esTipo(b)) return null;
        let x = FormatoDoc.ponerPPr(xml, { numeracion: { numId, nivel: b.numPedida ? parseInt(b.numPedida.nivel || '0', 10) : 0 } });
        if (parrafoLista && !b.estilo) x = ponerEstiloXml(x, parrafoLista);
        return x;
      });
    }

    // ───────── Tablas y saltos (6.6) ─────────
    function tablasDe(b) {
      const sal = [];
      for (let l = padres.get(b); l; ) { const t = duenoDeLista.get(l); if (!t) break; sal.push(t); l = listaDeTabla.get(t); }
      return sal;
    }
    const bajoTabla = (b, conjunto) => tablasDe(b).some((t) => conjunto.has(t));
    // Las tablas que quedan enteras dentro del rango (ni el principio ni el final están dentro de ellas); solo las más externas
    function tablasCubiertas(b0, b1, i0, i1) {
      const sal = new Set(), todos = parrafos(), fuera0 = new Set(tablasDe(b0)), fuera1 = new Set(tablasDe(b1));
      for (let i = i0; i <= i1; i++) for (const t of tablasDe(todos[i])) if (!fuera0.has(t) && !fuera1.has(t)) sal.add(t);
      for (const t of [...sal]) {
        // una celda no puede quedarse sin párrafos: si la tabla es lo único de su lista, no se quita
        const l = listaDeTabla.get(t);
        if (!l || !l.some((x) => x.tipo === 'p')) sal.delete(t);
      }
      // una tabla dentro de otra que también se quita no hace falta
      for (const t of [...sal]) { const dueno = duenoDeLista.get(listaDeTabla.get(t)); if (dueno && sal.has(dueno)) sal.delete(t); }
      return sal;
    }
    const vacioXml = '<w:p/>';
    const vacioP = () => (modelo.dialecto ? modelo.dialecto.parrafoVacio() : vacioXml);
    const TD = () => (modelo.dialecto ? modelo.dialecto.tabla : TablaDoc);
    function bloqueNuevo(xml) { const b = modelo.leerBloque(xml, null); return Object.assign(b, { fuente: xml, ini: 0, fin: xml.length, orig: null }); }

    // La celda donde está una posición: { T (la tabla), fila, celda, col, span, lista (donde está la tabla) } o null
    function celdaDe(p) {
      if (!p || indiceDe(p.b) < 0) return null;
      const l = padres.get(p.b), T = duenoDeLista.get(l);
      if (!T) return null;
      for (let fi = 0; fi < T.filas.length; fi++) for (let ci = 0; ci < T.filas[fi].celdas.length; ci++) {
        const c = T.filas[fi].celdas[ci];
        if (c.bloques === l) return { T, fila: fi, celda: ci, col: c.col, span: c.span, lista: listaDeTabla.get(T), filas: T.filas.length, columnas: T.nCols };
      }
      return null;
    }
    // El primer párrafo de la celda que cubre una columna de una fila de la tabla
    function parrafoEnCelda(T, fila, col) {
      const f = T.filas[Math.min(Math.max(fila, 0), T.filas.length - 1)];
      const c = f.celdas.find((x) => x.col <= col && col < x.col + x.span) || f.celdas[f.celdas.length - 1];
      const p = c.bloques.find((x) => x.tipo === 'p');
      return p ? { b: p, o: 0 } : null;
    }
    // Cambia la estructura de una tabla: \`xml\` es la tabla nueva ('' = quitarla). Devuelve la posición del cursor.
    function cambiarTabla(T, xml, nombre, fila, col, pos) {
      const L = listaDeTabla.get(T), idx = L.indexOf(T);
      if (xml === null || idx < 0) return null;
      let cDespues, poner;
      if (xml === '') {
        poner = [];
        if (!L.some((x) => x !== T && x.tipo === 'p')) poner = [bloqueNuevo(vacioP())];
        const sig = L[idx + 1] && L[idx + 1].tipo === 'p' ? L[idx + 1] : L[idx - 1] && L[idx - 1].tipo === 'p' ? L[idx - 1] : poner[0];
        cDespues = { b: sig, o: sig === L[idx - 1] ? sig.texto.length : 0 };
      } else {
        const M = bloqueNuevo(xml);
        M.orig = rangoOrigen(T);
        poner = [M];
        cDespues = parrafoEnCelda(M, fila, col);
      }
      historial.hacer(nombre, [{ lista: L, idx, n: 1, poner, cAntes: pos, cDespues }]);
      return cDespues;
    }
    const estructura = (pos, nombre, fn, dFila, dCol) => {
      if (!soporta('tablas')) return null;
      const c = celdaDe(pos);
      if (!c) return null;
      const nuevo = fn(xmlActual(c.T), c);
      return nuevo === null ? null : cambiarTabla(c.T, nuevo, nombre, c.fila + dFila(c), c.col + dCol(c), pos);
    };
    const filaAntes = (pos) => estructura(pos, 'Insertar fila', (x, c) => TD().insertarFila(x, c.fila, false), () => 0, () => 0);
    const filaDespues = (pos) => estructura(pos, 'Insertar fila', (x, c) => TD().insertarFila(x, c.fila, true), () => 1, () => 0);
    const columnaAntes = (pos) => estructura(pos, 'Insertar columna', (x, c) => TD().insertarColumna(x, c.col, false), () => 0, () => 0);
    const columnaDespues = (pos) => estructura(pos, 'Insertar columna', (x, c) => TD().insertarColumna(x, c.col, true), () => 0, (c) => c.span);
    const quitarFila = (pos) => estructura(pos, 'Eliminar fila', (x, c) => TD().eliminarFila(x, c.fila), (c) => (c.fila >= c.filas - 1 ? -1 : 0), () => 0);
    const quitarColumna = (pos) => estructura(pos, 'Eliminar columna', (x, c) => TD().eliminarColumna(x, c.col), () => 0, (c) => (c.col >= c.columnas - 1 ? -1 : 0));
    const quitarTabla = (pos) => estructura(pos, 'Eliminar tabla', () => '', () => 0, () => 0);
    // Tab: a la celda siguiente (en la última agrega una fila); Mayús+Tab: a la anterior
    function celdaVecina(pos, dir) {
      const c = celdaDe(pos);
      if (!c) return null;
      const celdas = [];
      c.T.filas.forEach((f) => f.celdas.forEach((x) => { if (!x.omitir) celdas.push(x); }));
      const k = celdas.findIndex((q) => q.bloques === padres.get(pos.b)), j = k + dir;
      if (j >= 0 && j < celdas.length) { const p = celdas[j].bloques.find((x) => x.tipo === 'p'); return p ? { b: p, o: 0 } : null; }
      if (dir > 0) return estructura(pos, 'Insertar fila', (x, cc) => TD().insertarFila(x, cc.filas - 1, true), (cc) => cc.filas - cc.fila, (cc) => -cc.col);
      return null;
    }
    // Una tabla nueva en la posición del cursor (solo en el cuerpo del documento)
    function insertarTabla(pos, filas, columnas) {
      if (!soporta('tablas') || indiceDe(pos.b) < 0) return null;
      if (modelo.dialecto && (pos.b.lista || pos.b.region || pos.b.protegido)) { ultimoRechazo = 'lista'; return null; }
      const L = padres.get(pos.b);
      if (L !== modelo.bloques) return null;
      const idx = L.indexOf(pos.b), sec = modelo.secciones.find((s) => s.hasta >= idx) || modelo.secciones[modelo.secciones.length - 1];
      const M = bloqueNuevo(TD().nueva(filas, columnas, sec.ancho - sec.izquierda - sec.derecha));
      const b = pos.b, o = pos.o, largo = b.texto.length;
      let op;
      if (o > 0 && o < largo) {
        const [x, y] = partirEn(leer(fragmento(b), b.r), o, b.r);
        op = { lista: L, idx, n: 1, poner: [bloqueDe(b, x, rangoOrigen(b)), M, bloqueDe(b, y, null)] };
      } else if (o === 0) {
        op = { lista: L, idx, n: 0, poner: L[idx - 1] && L[idx - 1].tipo === 'tabla' ? [bloqueNuevo(vacioP()), M] : [M] };
      } else {
        const sig = L[idx + 1];
        op = { lista: L, idx: idx + 1, n: 0, poner: sig && sig.tipo === 'p' ? [M] : [M, bloqueNuevo(vacioP())] };
      }
      const caret = parrafoEnCelda(M, 0, 0);
      op.cAntes = { b, o }; op.cDespues = caret;
      historial.hacer('Insertar tabla', [op]);
      return caret;
    }
    // Salto de sección ('nextPage' | 'continuous') en la posición del cursor (solo en el cuerpo del documento). El párrafo se parte en dos: el primero
    // lleva una copia del w:sectPr de la sección actual y cierra la primera; la segunda sección sigue con el sectPr de siempre, que ahora empieza como se pidió.
    function saltoDeSeccion(pos, tipo) {
      if (!soporta('secciones') || indiceDe(pos.b) < 0 || (tipo !== 'nextPage' && tipo !== 'continuous')) return null;
      const L = padres.get(pos.b);
      if (L !== modelo.bloques) return null;
      const idx = L.indexOf(pos.b), seccion = modelo.secciones.find((q) => q.hasta >= idx);
      if (!seccion) return null;
      const dueno = L[seccion.hasta], b = pos.b, o = Math.min(pos.o, b.texto.length);
      const esFinal = seccion === modelo.secciones[modelo.secciones.length - 1];
      // el w:sectPr que se copia al primer párrafo: el del párrafo que cierra la sección o el del cuerpo
      let copia = dueno && dueno.tipo === 'p' && dueno.seccion ? FormatoDoc.sectPrDe(fragmento(dueno)) : null;
      if (!copia) {
        const raiz = X().analizar(modelo.xml), cuerpo = raiz.hijos.find((h) => h.local === 'body'), sp = cuerpo && cuerpo.hijos.find((h) => h.local === 'sectPr');
        copia = sp ? modelo.xml.slice(sp.ini, sp.fin) : '<w:sectPr><w:pgSz w:w="' + Math.round(seccion.ancho * 20) + '" w:h="' + Math.round(seccion.alto * 20) + '"/><w:pgMar w:top="' + Math.round(seccion.arriba * 20) + '" w:right="' + Math.round(seccion.derecha * 20) + '" w:bottom="' + Math.round(seccion.abajo * 20) + '" w:left="' + Math.round(seccion.izquierda * 20) + '" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
        if (modelo.finalTipo !== undefined && modelo.finalTipo !== null) copia = FormatoDoc.sectPrConTipo(copia, modelo.finalTipo, 'w:');
      }
      const [x, y] = partirEn(leer(fragmento(b), b.r), o, b.r);
      const xmlA = FormatoDoc.ponerSectPr(x.xml, copia);
      let xmlB = y.xml;
      const ops = [];
      if (esFinal) ops.push({ finalTipo: { tipo } });
      else if (dueno === b) xmlB = FormatoDoc.ponerTipoSeccion(xmlB, tipo);
      else {
        const xd = FormatoDoc.ponerTipoSeccion(fragmento(dueno), tipo);
        ops.push({ lista: L, idx: seccion.hasta, n: 1, poner: [bloqueDesdeXml(dueno, xd)] });
      }
      const A = bloqueDesdeXml(b, xmlA), B = bloqueDesdeXml(b, xmlB);
      A.orig = rangoOrigen(b); B.orig = null;
      const caret = { b: B, o: 0 };
      ops.push({ lista: L, idx, n: 1, poner: [A, B] });
      ops.forEach((op) => { op.cAntes = { b, o }; op.cDespues = caret; });
      historial.hacer('Salto de sección', ops);
      return caret;
    }
    // Salto de página: el texto que sigue empieza en la hoja de abajo (en un párrafo nuevo, como Word)
    function saltoDePagina(desde, hasta) {
      if (!soporta('saltos')) return null;
      const [a, f] = ordenar(desde, hasta || desde);
      if (modelo.dialecto) {
        // ODF: el párrafo que sigue lleva un estilo con «salto antes»; al principio de un párrafo, es el mismo párrafo
        if (a.b.lista || a.b.protegido) { ultimoRechazo = a.b.protegido ? 'protegido' : 'lista'; return null; }
        return historial.agrupar('Salto de página', () => {
          let c = a;
          if (!(a.o === 0 && a.b === f.b && f.o === 0)) { c = reemplazar(a, f, '\n', { nombre: 'Salto de página' }); if (ultimoRechazo) return null; }
          const r = formatoParrafo(c, c, { saltoAntes: true }, 'Salto de página');
          return r ? r.desde : c;
        });
      }
      return reemplazar(a, f, '\f\n', { nombre: 'Salto de página', salto: true });
    }

    // Una imagen en línea en la posición del cursor. `bytes`: el archivo (png, jpeg, gif o bmp); `ancho` y `alto` en puntos.
    // Guarda la imagen en word/media, la relaciona con el documento y pone el dibujo en una corrida nueva.
    const TIPOS_IMAGEN = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp' };
    let contadorImagenes = 0;
    async function insertarImagen(pos, bytes, extension, ancho, alto, descripcion) {
      const ext = String(extension || '').toLowerCase().replace(/^\./, '');
      if (!soporta('imagenes') || !TIPOS_IMAGEN[ext] || indiceDe(pos.b) < 0 || !(ancho > 0) || !(alto > 0)) return null;
      if (modelo.dialecto && (pos.b.protegido || pos.b.region)) { ultimoRechazo = pos.b.protegido ? 'protegido' : 'lista'; return null; }
      let run = null;
      if (!modelo.dialecto) {
      const paquete = modelo.paquete;
      let n = 1;
      while (paquete.existe('word/media/imagen' + n + '.' + ext)) n++;
      const ruta = 'word/media/imagen' + n + '.' + ext;
      paquete.escribir(ruta, bytes);
      await paquete.asegurarTipo('/' + ruta, TIPOS_IMAGEN[ext]);
      const rid = await paquete.agregarRelacion(modelo.ruta, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', 'media/imagen' + n + '.' + ext);
      modelo.rels.set(rid, { tipo: 'image', destino: ruta, externo: false });
      const ids = [...modelo.xml.matchAll(/docPr\b[^>]*?\bid="(\d+)"/g)].map((m) => parseInt(m[1], 10));
      const id = Math.max(0, ...ids) + 1 + contadorImagenes++;
      const cx = Math.round(Math.min(ancho, 1584) * 12700), cy = Math.round(Math.min(alto, 1584) * 12700);
      const descr = X().escAttr(descripcion || '');
      run = '<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">'
        + `<wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="Imagen ${id}" descr="${descr}"/>`
        + '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        + '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        + `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="imagen${n}.${ext}"/><pic:cNvPicPr/></pic:nvPicPr>`
        + `<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
        + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>`
        + '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
      }
      const b = pos.b, base = b.r, F = leer(fragmento(b), base);
      const [A, B] = partirEn(F, Math.min(pos.o, F.texto.length), base);
      const solo = normalizar(modelo.dialecto ? await modelo.dialecto.imagenInline(bytes, ext, ancho, alto, descripcion, prefijoDe(F.raiz)) : '<w:p>' + run + '</w:p>'), rs = X().analizar(solo);
      const u1 = dx().unirXml(A.xml, A.raiz, solo, rs, false, false);
      const x1 = normalizar(u1.xml), u2 = dx().unirXml(x1, X().analizar(x1), B.xml, B.raiz, false, false);
      const nuevo = bloqueDesdeXml(b, u2.xml), L = contenedorDe(b);
      const cDespues = { b: nuevo, o: Math.min(pos.o, F.texto.length) + 1 };
      historial.hacer('Insertar imagen', [{ lista: L, idx: L.indexOf(b), n: 1, poner: [nuevo], cAntes: pos, cDespues }]);
      return cDespues;
    }

    // ───────── Buscar y reemplazar (6.7) ─────────
    // Una coincidencia que toca un campo (número de página…) o la marca de una nota no se puede cambiar: no cuenta
    const tocable = (m) => !m.b.piezas.some((z) => z.hasta > m.o && z.desde < m.o + m.n && (z.t === 'campo' || z.nota || (z.t === 'texto' && z.ti === undefined)));
    const buscar = (consulta, opc) => globalThis.BusquedaDoc.buscar(parrafos(), consulta, opc).filter(tocable);
    // Cambia una coincidencia por `nuevo`: lo nuevo se mete al final de lo viejo (así toma su formato) y después se quita lo viejo.
    // Devuelve el bloque nuevo y el desplazamiento donde termina lo nuevo.
    function cambiarCoincidencia(b, o, n, nuevo) {
      let blq = b;
      if (nuevo) blq = reemplazar({ b: blq, o: o + n }, { b: blq, o: o + n }, nuevo, { nombre: 'Reemplazar' }).b;
      blq = reemplazar({ b: blq, o }, { b: blq, o: o + n }, '', { nombre: 'Reemplazar' }).b;
      return { b: blq, o: o + nuevo.length };
    }
    const sinSaltos = (s) => String(s == null ? '' : s).replace(/\r?\n/g, String.fromCharCode(0x2028));
    function reemplazarUno(m, nuevo) {
      if (indiceDe(m.b) < 0) return null;
      return historial.agrupar('Reemplazar', () => cambiarCoincidencia(m.b, m.o, m.n, sinSaltos(nuevo)));
    }
    // Todas de una vez, como una sola cosa que se deshace; de abajo hacia arriba, para que los desplazamientos de arriba sigan valiendo
    function reemplazarTodo(consulta, nuevo, opc) {
      const ms = buscar(consulta, opc);
      if (!ms.length) return { n: 0, cursor: null };
      const rep = sinSaltos(nuevo), actual = new Map();
      let cursor = null;
      historial.agrupar('Reemplazar todo', () => {
        for (let i = ms.length - 1; i >= 0; i--) {
          const m = ms[i];
          cursor = cambiarCoincidencia(actual.get(m.b) || m.b, m.o, m.n, rep);
          actual.set(m.b, cursor.b);
        }
      });
      return { n: ms.length, cursor };
    }

    // ───────── Portapapeles con formato (6.7) ─────────
    // Lo que sale al copiar: un elemento por párrafo con su texto, sus tramos de formato (solo lo que se aparta del formato base del párrafo, para no
    // llenar el HTML de tamaños y fuentes que ya son los del documento) y las propiedades del párrafo. Es lo que lee y escribe HtmlDoc.
    function contenidoRico(desde, hasta) {
      const [p0, p1] = ordenar(desde, hasta), todos = parrafos(), i0 = indiceDe(p0.b), i1 = indiceDe(p1.b);
      if (i0 < 0 || i1 < 0) return [];
      const m = modelo.m;
      const col = (c) => { try { return c ? m.color(c) : null; } catch (e) { return null; } };
      const fue = (x) => { try { return x ? m.fuente(x) : null; } catch (e) { return null; } };
      const subr = (r) => !!(r && r.u && r.u !== 'none');
      const dif = (r, base) => {
        const f = {};
        if (r.b && !base.b) f.b = true;
        if (r.i && !base.i) f.i = true;
        if (subr(r) && !subr(base)) f.u = 'single';
        if (r.tachado && !base.tachado) f.tachado = true;
        const c = col(r.color);
        if (c && c !== col(base.color) && /^#[0-9a-f]{6}$/i.test(c)) f.color = c.slice(1).toUpperCase();
        if (r.tam && r.tam !== base.tam) f.tam = r.tam;
        const fu = fue(r.fuentes);
        if (fu && fu !== fue(base.fuentes)) f.fuente = fu;
        if (r.resaltado && r.resaltado !== base.resaltado) f.resaltado = r.resaltado;
        if (r.alineacionVertical && r.alineacionVertical !== 'baseline' && r.alineacionVertical !== base.alineacionVertical) f.vertical = r.alineacionVertical;
        return f;
      };
      const sal = [];
      for (let i = i0; i <= i1; i++) {
        const b = todos[i], s = b === p0.b ? p0.o : 0, e = b === p1.b ? p1.o : b.texto.length;
        let texto = '';
        const tramos = [];
        for (const z of b.piezas) {
          if (z.hasta <= z.desde || z.hasta <= s || z.desde >= e) continue;
          const t = b.texto.slice(Math.max(s, z.desde), Math.min(e, z.hasta)).replace(/[￼\f]/g, '');
          if (!t) continue;
          const f = dif(z.r || {}, b.r || {}), d = texto.length;
          texto += t;
          const u = tramos[tramos.length - 1];
          if (u && JSON.stringify(u.f) === JSON.stringify(f)) u.h = texto.length; else tramos.push({ d, h: texto.length, f });
        }
        const j = b.p.jc, der = !!b.p.derechaAIzquierda;
        const visual = j === 'both' || j === 'distribute' || j === 'justify' ? 'both' : j === 'center' ? 'center' : (j === 'right' || j === 'end') === !der ? 'right' : null;
        const h = /^(?:Heading|Ttulo)([1-3])$/.exec(b.estilo || '');
        sal.push({
          texto, tramos, jc: visual, estilo: h ? 'Heading' + h[1] : null,
          lista: b.etiqueta && b.numPedida ? { tipo: b.etiqueta.fmt === 'bullet' ? 'vinetas' : 'numeros', nivel: parseInt(b.numPedida.nivel || '0', 10) } : null,
        });
      }
      return sal;
    }
    // Pega párrafos con formato (lo que lee HtmlDoc) en lugar del rango, como un solo paso. Los que no traen formato se quedan sin él aunque el
    // texto de al lado lo tenga; las propiedades de párrafo (alineación, título, lista) solo se aplican cuando se pegan varios párrafos.
    async function pegarRico(desde, hasta, rico) {
      rico = (rico || []).filter((p) => p && typeof p.texto === 'string');
      if (!rico.length) return null;
      const [p0, p1] = ordenar(desde, hasta);
      if (indiceDe(p0.b) < 0 || indiceDe(p1.b) < 0) return null;
      const N = rico.length, SU = String.fromCharCode(0x2028);
      const linea = (p) => p.texto.replace(/\r?\n/g, SU);
      // lo que hace falta del paquete (estilos, listas) se prepara antes: esas partes no se deshacen, pero tampoco molestan
      // (.odt) estilos, alineación y listas van por el dialecto: nada de w:pPr ni de word/numbering.xml en un paquete ODF
      const estilos = new Map(), nums = [], odt = modelo.dialecto;
      if (N > 1 && odt) for (const p of rico) if (p.estilo && !estilos.has(p.estilo)) estilos.set(p.estilo, odt.asegurarEstilo(p.estilo));
      if (N > 1 && !odt) {
        for (const p of rico) if (p.estilo && !estilos.has(p.estilo)) estilos.set(p.estilo, await asegurarEstilo(p.estilo));
        estilos.set('ListParagraph', rico.some((p) => p.lista) ? await asegurarEstilo('ListParagraph') : null);
        let previa = null;
        for (const p of rico) {
          if (!p.lista) { previa = null; continue; }
          if (p.lista.tipo === 'vinetas') { nums.push(await nuevaLista('vinetas')); previa = null; } else { if (!previa) previa = await nuevaLista('numeros'); nums.push(previa); }
        }
      }
      const e = estadoFormato(p0, p0) || {}, neutro = {};
      if (e.b === true) neutro.b = false;
      if (e.i === true) neutro.i = false;
      if (e.u === true) neutro.u = 'none';
      if (e.tachado === true) neutro.tachado = false;
      if (e.vertical) neutro.vertical = 'baseline';
      if (e.resaltado) neutro.resaltado = 'none';
      if (e.color && e.color !== '#000000') neutro.color = 'auto';
      return historial.agrupar('Pegar', () => {
        const caret = reemplazar(p0, p1, rico.map(linea).join('\n'), { nombre: 'Pegar' });
        const iFin = indiceDe(caret.b), iIni = iFin - (N - 1);
        const bloque = (k) => parrafos()[iIni + k];
        rico.forEach((p, k) => {
          const base = k === 0 ? p0.o : 0, largo = linea(p).length;
          for (const x of p.tramos && p.tramos.length ? p.tramos : [{ d: 0, h: largo, f: {} }]) {
            const f = { ...neutro, ...(x.f || {}) };
            if (!Object.keys(f).length || x.h <= x.d) continue;
            const b = bloque(k);
            formatoCaracter({ b, o: base + x.d }, { b, o: base + x.h }, f, 'Pegar');
          }
        });
        if (N > 1) {
          rico.forEach((p, k) => {
            if ((k === 0 && p0.o > 0) || (!p.jc && !p.estilo && !p.lista)) return;
            const pos = { b: bloque(k), o: 0 };
            cambiarParrafos(pos, pos, 'Pegar', (bl, xml) => {
              let x = xml;
              if (odt) {
                if (p.jc) x = odt.formatoParrafo(x, { jc: p.jc });
                if (p.estilo && estilos.get(p.estilo)) x = odt.estiloConservando(x, estilos.get(p.estilo));
                return x === xml ? null : x;
              }
              if (p.jc) x = FormatoDoc.ponerPPr(x, { jc: p.jc });
              if (p.estilo && estilos.get(p.estilo)) x = ponerEstiloXml(x, estilos.get(p.estilo));
              if (p.lista) {
                const i = rico.slice(0, k).filter((q) => q.lista).length;
                x = FormatoDoc.ponerPPr(x, { numeracion: { numId: nums[i], nivel: Math.min(8, Math.max(0, p.lista.nivel || 0)) } });
                if (estilos.get('ListParagraph') && !bl.estilo) x = ponerEstiloXml(x, estilos.get('ListParagraph'));
              }
              return x === xml ? null : x;
            });
          });
          // (.odt) cada racha de párrafos seguidos con el mismo tipo de lista se vuelve una lista (región reescrita) y después cada uno toma su nivel
          for (let k = odt ? 0 : N; k < N;) {
            const t = rico[k].lista && !(k === 0 && p0.o > 0) ? rico[k].lista.tipo : null;
            let z = k;
            if (t) while (z + 1 < N && rico[z + 1].lista && rico[z + 1].lista.tipo === t) z++;
            if (t && operarLista('poner', { b: bloque(k), o: 0 }, { b: bloque(z), o: 0 }, { tipo: t }, 'Pegar')) {
              for (let q = k; q <= z; q++) {
                const n = Math.min(8, Math.max(0, rico[q].lista.nivel || 0)), pos = { b: bloque(q), o: 0 };
                if (n) operarLista('nivel', pos, pos, { delta: n }, 'Pegar');
              }
            }
            k = z + 1;
          }
        }
        return { b: parrafos()[iFin], o: caret.o };
      });
    }

    // ───────── Encabezados y pies (6.7) ─────────
    // La sección (su posición en modelo.secciones) a la que pertenece un párrafo; el de una tabla es el de la tabla
    function seccionDe(p) {
      if (indiceDe(p.b) < 0) return 0;
      const ts = tablasDe(p.b), raiz = ts.length ? ts[ts.length - 1] : p.b, i = modelo.bloques.indexOf(raiz);
      const k = modelo.secciones.findIndex((s) => s.hasta >= i);
      return k < 0 ? modelo.secciones.length - 1 : k;
    }
    // Un campo de página (PAGE) o de total de páginas (NUMPAGES) en la posición dada; vale un carácter en el texto. Devuelve el cursor detrás de él.
    function insertarCampo(pos, nombre) {
      if (!soporta('campos') || !/^(PAGE|NUMPAGES)$/.test(nombre) || indiceDe(pos.b) < 0) return null;
      const b = pos.b, F = leer(fragmento(b), b.r), o = Math.min(pos.o, F.texto.length), pf = prefijoDe(F.raiz);
      const [A, B] = partirEn(F, o, b.r);
      if (modelo.dialecto && (b.protegido || b.region)) { ultimoRechazo = b.protegido ? 'protegido' : 'lista'; return null; }
      const campo = normalizar(modelo.dialecto ? modelo.dialecto.campoXml(nombre, pf) : `<${pf}p><${pf}fldSimple ${pf}instr=" ${nombre} "><${pf}r><${pf}t>1</${pf}t></${pf}r></${pf}fldSimple></${pf}p>`);
      const u1 = dx().unirXml(A.xml, A.raiz, campo, X().analizar(campo), false, false);
      const x1 = normalizar(u1.xml), u2 = dx().unirXml(x1, X().analizar(x1), B.xml, B.raiz, false, false);
      const nuevo = bloqueDesdeXml(b, u2.xml), L = contenedorDe(b);
      const cDespues = { b: nuevo, o: o + 1 };
      historial.hacer('Insertar campo', [{ lista: L, idx: L.indexOf(b), n: 1, poner: [nuevo], cAntes: pos, cDespues }]);
      return cDespues;
    }

    function todo() {
      const t = parrafos();
      return t.length ? { desde: { b: t[0], o: 0 }, hasta: { b: t[t.length - 1], o: t[t.length - 1].texto.length } } : null;
    }
    const volver = (fn) => () => { ultimoCursor = null; return fn() ? ultimoCursor : null; };

    // ── guardar: el document.xml original con solo los párrafos editados cambiados ──
    // El texto de un bloque que se escribe en el archivo: un párrafo editado lleva su propio XML; una tabla reemplazada por una estructura nueva
    // (fila, columna…) lleva el suyo con las ediciones de sus celdas ya puestas
    const textoDe = (b) => (b.tipo === 'tabla' && b.fuente ? xmlActual(b) : (b.fuente || b.base || modelo.xml).slice(b.ini, b.fin));
    // Las ediciones de una lista de bloques (relativas al texto del que vienen sus bloques originales)
    function edicionesDeLista(l, orig, sinRegiones) {
      const ed = [];
      // (.odt) cada región se escribe de una vez: un solo «bloque» con el XML de la región como está ahora, en el sitio del primer original que cubría
      const cubiertos = [];
      if (!sinRegiones && l.some((b) => b.region)) {
        const hechas = new Set(), l2 = [];
        for (const b of l) {
          if (!b.region) { l2.push(b); continue; }
          if (hechas.has(b.region)) continue;
          hechas.add(b.region);
          const R = b.region, xml = armarRegion(R, l.filter((q) => q.region === R));
          l2.push({ tipo: 'p', fuente: xml, ini: 0, fin: xml.length, orig: R.rangos[0] || null, esRegion: true });
          for (const r of R.rangos) cubiertos.push(r);
          for (const r of R.rangos.slice(1)) ed.push({ ini: r.ini, fin: r.fin, texto: '' });
        }
        l = l2;
      }
      const dentro = (r) => cubiertos.some((c) => r.ini >= c.ini && r.fin <= c.fin);
      const vivos = new Set(l), heredados = new Set(l.map((b) => (rangoOrigen(b) || {}).ini));
      // un bloque original que ya no está se borra, salvo que otro bloque nuevo ocupe su lugar (hereda su rango)
      for (const b of orig) if (!vivos.has(b)) { const r = rangoOrigen(b); if (r && !heredados.has(r.ini) && !dentro(r)) ed.push({ ini: r.ini, fin: r.fin, texto: '' }); }
      const tras = new Map();        // bloque original → fragmentos nuevos que van detrás
      const delante = [];            // fragmentos nuevos que van antes del primer bloque original vivo
      let ancla = null;
      for (const b of l) {
        const r = rangoOrigen(b);
        if (r) {
          ancla = b;
          if (b.fuente) ed.push({ ini: r.ini, fin: r.fin, texto: textoDe(b) });
        } else {
          const f = textoDe(b);
          if (ancla) { if (!tras.has(ancla)) tras.set(ancla, []); tras.get(ancla).push(f); } else delante.push(f);
        }
      }
      for (const [b, fs] of tras) { const r = rangoOrigen(b); ed.push({ ini: r.fin, fin: r.fin, texto: fs.join('') }); }
      if (delante.length) {
        const primero = l.find((b) => rangoOrigen(b));
        const r = primero ? rangoOrigen(primero) : (orig.map(rangoOrigen).find(Boolean) || null);
        if (r) ed.push({ ini: r.ini, fin: r.ini, texto: delante.join('') });
      }
      return ed;
    }
    // El XML de una región como está ahora (con lo escrito en sus párrafos), sin el envoltorio <w>…</w>
    function armarRegion(R, vivos) {
      const ed = edicionesDeLista(vivos, R.iniciales, true);
      const xml = ed.length ? X().aplicar(R.xml, ed) : R.xml;
      return xml.slice(3, xml.length - 4);
    }
    // ¿De qué texto cuelgan los bloques de esta lista? null: del documento; una tabla con XML propio: de esa; undefined: de una tabla que ya no está
    function baseDeLista(l) {
      for (let x = l; ;) {
        if (x === modelo.bloques) return null;
        const t = duenoDeLista.get(x);
        if (!t) return undefined;
        const arriba = listaDeTabla.get(t);
        if (!arriba || !arriba.includes(t)) return undefined;
        if (t.fuente) return t;
        x = arriba;
      }
    }
    // El XML de una tabla como está ahora (con lo escrito en sus celdas)
    function xmlActual(T) {
      const base = T.fuente || T.base || modelo.xml, ed = [];
      const visitar = (t) => {
        for (const f of t.filas) for (const c of f.celdas) {
          if (origenes.has(c.bloques)) ed.push(...edicionesDeLista(c.bloques, origenes.get(c.bloques)));
          for (const b of c.bloques) if (b.tipo === 'tabla' && !b.fuente) visitar(b);
        }
      };
      visitar(T);
      const sub = base.slice(T.ini, T.fin);
      return ed.length ? X().aplicar(sub, ed.map((e) => ({ ini: e.ini - T.ini, fin: e.fin - T.ini, texto: e.texto }))) : sub;
    }
    async function xmlGuardado() {
      const ed = [];
      for (const [l, orig] of origenes) if (baseDeLista(l) === null) ed.push(...edicionesDeLista(l, orig));
      const nuevas = modelo.referenciasNuevas || [], hayTipo = modelo.finalTipo !== undefined && modelo.finalTipo !== null;
      if (hayTipo || nuevas.length) {
        const raiz = X().analizar(modelo.xml), cuerpo = raiz.hijos.find((h) => h.local === 'body'), sp = cuerpo && cuerpo.hijos.find((h) => h.local === 'sectPr');
        const pf = prefijoDe(raiz);
        // un cuerpo sin <w:sectPr> (Word usa entonces los valores por defecto) recibe uno, o la referencia o el tipo nuevos se perderían al guardar
        const sinSectPr = !sp && cuerpo && cuerpo.iniCierre >= 0;
        if (sp || sinSectPr) {
          let t = sp ? modelo.xml.slice(sp.ini, sp.fin) : `<${pf}sectPr></${pf}sectPr>`;
          if (hayTipo) t = FormatoDoc.sectPrConTipo(t, modelo.finalTipo, pf);
          // encabezados y pies creados en esta sesión: sus referencias van al principio del <w:sectPr>, con su propio espacio de nombres para r:id
          if (nuevas.length) {
            const r = X().analizar(sinCierre(t));
            const refs = nuevas.map((n) => `<${pf}${n.tipo === 'encabezado' ? 'header' : 'footer'}Reference ${pf}type="${X().escAttr(n.cual)}" r:id="${X().escAttr(n.id)}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`).join('');
            const u = sinCierre(t);
            t = u.slice(0, r.finApertura) + refs + u.slice(r.finApertura);
          }
          ed.push(sp ? { ini: sp.ini, fin: sp.fin, texto: t } : { ini: cuerpo.iniCierre, fin: cuerpo.iniCierre, texto: t });
        }
      }
      const final = ed.length ? X().aplicar(modelo.xml, ed) : modelo.xml;
      return modelo.dialecto && modelo.dialecto.finalizar ? modelo.dialecto.finalizar(final) : final;
    }
    async function guardar() {
      const xmlFinal = await xmlGuardado();
      modelo.paquete.escribir(modelo.ruta, xmlFinal);
      if (modelo.dialecto && modelo.dialecto.alGuardar) await modelo.dialecto.alGuardar(modelo.paquete, xmlFinal);
      // una numeración creada desde cero es una parte nueva: necesita su tipo de contenido y su relación
      // (una promesa y no un sí/no: dos guardados a la vez, el autoguardado y la descarga, la relacionarían dos veces)
      if (modelo.numeracionNueva && !modelo.numeracionRelacionada) {
        const ruta = modelo.numeracionNueva;
        modelo.numeracionRelacionada = (async () => {
          await modelo.paquete.asegurarTipo('/' + ruta, 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml');
          await modelo.paquete.agregarRelacion(modelo.ruta, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering', ruta.replace(/^word\//, ''));
        })();
        modelo.numeracionRelacionada.catch(() => { modelo.numeracionRelacionada = null; });
      }
      if (modelo.numeracionRelacionada) await modelo.numeracionRelacionada;
      return modelo.paquete.generar();
    }

    return {
      modelo, historial, parrafos, indiceDe, contenedorDe, comparar, ordenar, anterior, siguiente, todo,
      buscar, reemplazarUno, reemplazarTodo, contenidoRico, pegarRico, seccionDe, insertarCampo,
      reemplazar, borrarAtras, borrarAdelante, entrar, quitarLista, texto: textoEntre,
      saltoDeSeccion, insertarImagen, celdaDe, filaAntes, filaDespues, columnaAntes, columnaDespues, quitarFila, quitarColumna, quitarTabla, celdaVecina, insertarTabla, saltoDePagina,
      formatoCaracter, alternarCaracter, estadoFormato, formatoParrafo, alinear, sangrar, nivelDeLista, interlineado, espaciado, estilosDisponibles, ponerEstilo, ponerLista,
      deshacer: volver(() => historial.deshacer()), rehacer: volver(() => historial.rehacer()),
      guardar, xmlGuardado,
      puedeDeshacer: () => historial.puedeDeshacer(), puedeRehacer: () => historial.puedeRehacer(), sucio: () => historial.sucio(),
      cortarFusion: () => historial.cortarFusion(),
      puede: soporta, rechazo: () => ultimoRechazo,
    };
  }

  globalThis.EdicionDoc = { crear, palabraAntes, palabraDespues, palabraEn, pasoAtras, pasoAdelante, limpiar, contenidoDe, ponerEstiloXml, quitarNumeracionXml };
})();

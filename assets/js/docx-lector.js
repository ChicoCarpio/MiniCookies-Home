/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Lector de .docx propio (reemplaza a mammoth). Lee un Word y lo entrega como
// HTML simple (párrafos, títulos, listas, tablas, enlaces, imágenes, notas al
// pie) o como texto plano. Es lo que usan las herramientas de Word para leer.
//
// Requiere DocxBase (docx-base.js), ZipPropio y XmlLigero. No depende del navegador: corre igual en Node.
//
// API (la forma de mammoth, para migrar una página sin rehacerla):
//   const { value, messages } = await DocxLector.convertirAHtml(datos, {
//     imagen: async ({ bytes, tipo, alt }) => ({ src })   // opcional; sin ella, data: URI
//   });
//   const { value, messages } = await DocxLector.extraerTexto(datos);
//   `datos` = ArrayBuffer, Uint8Array o Blob/File.
//   `messages` = [{ type: 'warning', message }]: solo problemas reales (una
//   imagen que no se pudo leer, una imagen vinculada). No avisa de estilos que
//   no conoce, como hacía mammoth.
//
// Lo que lee: estilos de título, listas con numeración (anidadas), negrita,
// cursiva, tachado, superíndice y subíndice, tablas (celdas combinadas), enlaces
// (también los campos HYPERLINK, simples y complejos), marcadores, saltos, tabulaciones,
// imágenes, cuadros de texto, notas al pie y notas finales. Ignora el texto
// borrado con control de cambios y los comentarios.

(function () {
  'use strict';

  const { arbol, hijo, lista, activo, buscar, porTipo, base64, hrefSeguro, abrirPaquete } = DocxBase;

  // ───────── HTML ─────────
  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const escAttr = (s) => esc(s).replace(/"/g, '&quot;');
  const VACIOS = new Set(['br', 'img', 'hr']);

  function serializar(n) {
    if (typeof n === 'string') return esc(n);
    const attrs = n.attrs ? Object.keys(n.attrs).map((k) => ` ${k}="${escAttr(String(n.attrs[k]))}"`).join('') : '';
    if (VACIOS.has(n.e)) return `<${n.e}${attrs} />`;
    return `<${n.e}${attrs}>${n.h.map(serializar).join('')}</${n.e}>`;
  }
  const el = (e, attrs, h) => ({ e, attrs, h: h || [] });

  // Agrega `contenido` dentro de la cadena de etiquetas `etiquetas` (de afuera
  // hacia adentro), juntando con el hermano anterior si tiene la misma: así dos
  // trozos en negrita seguidos salen como un solo <strong>.
  function agregar(padre, etiquetas, contenido) {
    let nodo = padre;
    for (const etiqueta of etiquetas) {
      const ultimo = nodo.h[nodo.h.length - 1];
      if (ultimo && typeof ultimo !== 'string' && ultimo.e === etiqueta && !ultimo.attrs) nodo = ultimo;
      else { const nuevo = el(etiqueta, null); nodo.h.push(nuevo); nodo = nuevo; }
    }
    if (typeof contenido === 'string') {
      const ultimo = nodo.h[nodo.h.length - 1];
      if (typeof ultimo === 'string') nodo.h[nodo.h.length - 1] = ultimo + contenido;
      else nodo.h.push(contenido);
    } else nodo.h.push(contenido);
  }

  const TIPOS_IMAGEN = {
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp',
    tif: 'image/tiff', tiff: 'image/tiff', svg: 'image/svg+xml', emf: 'image/x-emf', wmf: 'image/x-wmf', webp: 'image/webp',
  };

  function leerEstilos(raiz) {
    const estilos = new Map();
    for (const s of lista(raiz, 'style')) {
      const pPr = hijo(s, 'pPr');
      const num = hijo(pPr, 'numPr');
      estilos.set(s.a.styleId, {
        nombre: (hijo(s, 'name') || { a: {} }).a.val || '',
        basadoEn: (hijo(s, 'basedOn') || { a: {} }).a.val || null,
        numId: num && hijo(num, 'numId') ? hijo(num, 'numId').a.val : null,
        ilvl: num && hijo(num, 'ilvl') ? hijo(num, 'ilvl').a.val : null,
      });
    }
    return estilos;
  }

  function leerNumeracion(raiz) {
    const abstractos = new Map(); // id → Map(ilvl → formato)
    const numeros = new Map();    // numId → { abstracto, niveles: Map(ilvl → formato) }
    for (const a of lista(raiz, 'abstractNum')) {
      const niveles = new Map();
      for (const l of lista(a, 'lvl')) niveles.set(l.a.ilvl || '0', (hijo(l, 'numFmt') || { a: {} }).a.val || 'decimal');
      abstractos.set(a.a.abstractNumId, niveles);
    }
    for (const n of lista(raiz, 'num')) {
      const niveles = new Map();
      for (const o of lista(n, 'lvlOverride')) {
        const l = hijo(o, 'lvl');
        if (l && hijo(l, 'numFmt')) niveles.set(o.a.ilvl || '0', hijo(l, 'numFmt').a.val);
      }
      numeros.set(n.a.numId, { abstracto: (hijo(n, 'abstractNumId') || { a: {} }).a.val, niveles });
    }
    return { abstractos, numeros };
  }

  const PROFUNDIDAD_MAXIMA = 100; // un Word real no pasa de unos pocos niveles; más es un archivo hecho para romper

  // ───────── Lectura del documento ─────────
  class Lector {
    constructor(zip, partes, opciones) {
      Object.assign(this, partes);
      this.zip = zip;
      this.opciones = opciones || {};
      this.mensajes = [];
      this.notas = [];        // referencias vistas: { tipo, id }
      this.imagenes = new Map();
      this.prof = 0;
    }

    entrar() { if (++this.prof > PROFUNDIDAD_MAXIMA) throw new Error('El documento tiene demasiados niveles anidados y no se puede leer.'); }

    aviso(m) { if (!this.mensajes.some((x) => x.message === m)) this.mensajes.push({ type: 'warning', message: m }); }

    nombreEstilo(id) { const e = this.estilos.get(id); return e ? e.nombre : ''; }

    numeracionDeEstilo(id) {
      for (let i = 0, e = this.estilos.get(id); e && i < 12; i++, e = this.estilos.get(e.basadoEn)) {
        if (e.numId != null) return { numId: e.numId, ilvl: e.ilvl || '0' };
        if (!e.basadoEn) break;
      }
      return null;
    }

    esOrdenada(numId, ilvl) {
      const n = this.numeracion.numeros.get(numId);
      if (!n) return false;
      const formato = n.niveles.get(ilvl) || (this.numeracion.abstractos.get(n.abstracto) || new Map()).get(ilvl);
      return !!formato && formato !== 'bullet' && formato !== 'none';
    }

    // Bloques de un contenedor (cuerpo, celda, nota, cuadro de texto).
    async bloques(contenedor, destino) {
      destino = destino || [];
      this.entrar();
      for (const e of contenedor.h) await this.bloque(e, destino);
      this.prof--;
      return destino;
    }

    async bloque(e, destino) {
      switch (e.n) {
        case 'p': {
          const p = await this.parrafo(e);
          destino.push(p.bloque);
          for (const extra of p.extra) destino.push(extra);
          break;
        }
        case 'tbl': destino.push(await this.tabla(e)); break;
        case 'sdt': { const c = hijo(e, 'sdtContent'); if (c) await this.bloques(c, destino); break; }
        case 'ins': case 'moveTo': case 'customXml': case 'smartTag': case 'sdtContent':
          await this.bloques(e, destino); break;
        case 'AlternateContent': { const c = hijo(e, 'Fallback') || hijo(e, 'Choice'); if (c) await this.bloques(c, destino); break; }
        default: break; // sectPr, bookmarkStart, del, etc.
      }
    }

    async parrafo(e) {
      const pPr = hijo(e, 'pPr');
      const idEstilo = pPr && hijo(pPr, 'pStyle') ? hijo(pPr, 'pStyle').a.val : null;
      const nombre = idEstilo ? this.nombreEstilo(idEstilo) : '';
      let etiqueta = 'p';
      const t = /^heading\s*([1-6])?$/i.exec(nombre.trim()) || /^Heading([1-6])$/i.exec(idEstilo || '');
      if (t) etiqueta = 'h' + (t[1] || '1');

      let num = null;
      if (etiqueta === 'p') {
        const np = pPr && hijo(pPr, 'numPr');
        let numId = np && hijo(np, 'numId') ? hijo(np, 'numId').a.val : null;
        let ilvl = np && hijo(np, 'ilvl') ? hijo(np, 'ilvl').a.val : null;
        if (numId == null && idEstilo) {
          const ne = this.numeracionDeEstilo(idEstilo);
          if (ne) { numId = ne.numId; if (ilvl == null) ilvl = ne.ilvl; }
        }
        if (numId != null && numId !== '0') {
          const nivel = parseInt(ilvl || '0', 10) || 0;
          num = { nivel, ordenada: this.esOrdenada(numId, String(nivel)) };
        }
      }
      const ctx = { campos: [], salida: [[]], extra: [] };
      await this.inline(e, ctx);
      while (ctx.salida.length > 1) { const x = ctx.salida.pop(); ctx.salida[ctx.salida.length - 1].push(...x); }
      return { bloque: { k: 'p', etiqueta, num, hijos: ctx.salida[0] }, extra: ctx.extra };
    }

    emitir(ctx, nodo) { ctx.salida[ctx.salida.length - 1].push(nodo); }

    // Hijos de un párrafo (o de un enlace, o de un fragmento) → nodos en ctx.salida.
    async inline(e, ctx) {
      this.entrar();
      for (const x of e.h) {
        switch (x.n) {
          case 'r': await this.corrida(x, ctx); break;
          case 'hyperlink': {
            let href = '';
            const rid = x.a.id;
            if (rid && this.rels.has(rid)) href = this.rels.get(rid).destino;
            if (x.a.anchor) href = (href || '') + '#' + x.a.anchor;
            ctx.salida.push([]);
            await this.inline(x, ctx);
            const hijos = ctx.salida.pop();
            this.emitir(ctx, { k: 'a', href: hrefSeguro(href), hijos });
            break;
          }
          case 'fldSimple': {
            const h = enlaceDeInstruccion(x.a.instr || '');
            ctx.salida.push([]);
            await this.inline(x, ctx);
            const hijos = ctx.salida.pop();
            if (h) this.emitir(ctx, { k: 'a', href: hrefSeguro(h), hijos });
            else for (const n of hijos) this.emitir(ctx, n);
            break;
          }
          case 'ins': case 'moveTo': case 'smartTag': case 'customXml': case 'sdtContent': case 'dir': case 'bdo':
            await this.inline(x, ctx); break;
          case 'sdt': { const c = hijo(x, 'sdtContent'); if (c) await this.inline(c, ctx); break; }
          case 'AlternateContent': { const c = hijo(x, 'Fallback') || hijo(x, 'Choice'); if (c) await this.inline(c, ctx); break; }
          case 'bookmarkStart': {
            const nombre = x.a.name;
            if (nombre && nombre !== '_GoBack') this.emitir(ctx, { k: 'ancla', id: nombre });
            break;
          }
          default: break; // pPr, proofErr, del, commentRangeStart…
        }
      }
      this.prof--;
    }

    async corrida(r, ctx) {
      const rPr = hijo(r, 'rPr');
      const f = { b: false, i: false, s: false, sup: false, sub: false };
      if (rPr) {
        f.b = activo(hijo(rPr, 'b'));
        f.i = activo(hijo(rPr, 'i'));
        f.s = activo(hijo(rPr, 'strike')) || activo(hijo(rPr, 'dstrike'));
        const va = hijo(rPr, 'vertAlign');
        if (va) { f.sup = va.a.val === 'superscript'; f.sub = va.a.val === 'subscript'; }
        const rs = hijo(rPr, 'rStyle');
        if (rs && /^strong$/i.test(this.nombreEstilo(rs.a.val))) f.b = true;
      }
      const texto = (s) => { if (s) this.emitir(ctx, { k: 't', s, f }); };
      const campo = ctx.campos[ctx.campos.length - 1];
      for (const x of r.h) {
        switch (x.n) {
          case 'fldChar': {
            const tipo = x.a.fldCharType;
            if (tipo === 'begin') ctx.campos.push({ fase: 'instr', instr: '', href: null });
            else if (tipo === 'separate') {
              const c = ctx.campos[ctx.campos.length - 1];
              if (c) {
                c.fase = 'res';
                c.href = enlaceDeInstruccion(c.instr);
                if (c.href !== null) ctx.salida.push([]);
              }
            } else if (tipo === 'end') {
              const c = ctx.campos.pop();
              if (c && c.href !== null && c.fase === 'res') {
                const hijos = ctx.salida.pop();
                this.emitir(ctx, { k: 'a', href: hrefSeguro(c.href), hijos });
              }
            }
            break;
          }
          case 'instrText': if (campo && campo.fase === 'instr') campo.instr += x.t; break;
          case 't': if (!campo || campo.fase === 'res') texto(x.t); break;
          case 'tab': case 'ptab': texto('\t'); break;
          case 'noBreakHyphen': texto('‑'); break;
          case 'softHyphen': texto('­'); break;
          case 'sym': if (x.a.char) texto(String.fromCodePoint(parseInt(x.a.char, 16) & 0xFFFF)); break;
          case 'br': if (!x.a.type || x.a.type === 'textWrapping') this.emitir(ctx, { k: 'br' }); break;
          case 'cr': this.emitir(ctx, { k: 'br' }); break;
          case 'footnoteReference': case 'endnoteReference': {
            const tipo = x.n === 'footnoteReference' ? 'footnote' : 'endnote';
            this.notas.push({ tipo, id: x.a.id });
            this.emitir(ctx, { k: 'nota', tipo, id: x.a.id, n: this.notas.filter((m) => m.tipo === tipo).length });
            break;
          }
          case 'drawing': case 'pict': case 'object': await this.imagen(x, ctx); await this.cuadros(x, ctx); break;
          case 'AlternateContent': {
            const c = hijo(x, 'Fallback') || hijo(x, 'Choice');
            if (c) { await this.imagen(c, ctx); await this.cuadros(c, ctx); }
            break;
          }
          default: break; // rPr, lastRenderedPageBreak, annotationRef…
        }
      }
    }

    // Texto de los cuadros de texto: sus párrafos salen después del párrafo que los contiene.
    async cuadros(e, ctx) {
      const caja = buscar(e, 'txbxContent');
      if (caja) await this.bloques(caja, ctx.extra);
    }

    async imagen(e, ctx) {
      const blip = buscar(e, 'blip');
      const vml = buscar(e, 'imagedata');
      let rid = null;
      if (blip) {
        rid = blip.a.embed;
        if (!rid && blip.a.link) { this.aviso('Hay una imagen vinculada a un archivo externo; no se puede leer.'); return; }
      } else if (vml) rid = vml.a.id;
      if (!rid) return;
      const rel = this.rels.get(rid);
      if (!rel || rel.externo) { this.aviso('Hay una imagen que no se pudo encontrar dentro del documento.'); return; }
      const doc = buscar(e, 'docPr');
      const alt = doc ? doc.a.descr || doc.a.title || '' : (vml && vml.a.title) || '';
      let src = this.imagenes.get(rel.destino + '|' + alt);
      if (src === undefined) {
        try {
          const f = this.zip.file(rel.destino);
          if (!f) throw new Error('falta');
          const bytes = await f.async('uint8array');
          const ext = rel.destino.split('.').pop().toLowerCase();
          const tipo = TIPOS_IMAGEN[ext] || 'application/octet-stream';
          src = this.opciones.imagen
            ? (await this.opciones.imagen({ bytes, tipo, alt })).src
            : `data:${tipo};base64,${base64(bytes)}`;
        } catch (error) {
          this.aviso('No se pudo leer una imagen del documento.');
          src = '';
        }
        this.imagenes.set(rel.destino + '|' + alt, src);
      }
      if (src) this.emitir(ctx, { k: 'img', src, alt });
    }

    async tabla(e) {
      const filas = [];
      const abiertas = []; // por columna: la celda de arriba que sigue combinada
      const todas = hijosDeTabla(e, 'tr');
      for (const tr of todas) {
        const trPr = hijo(tr, 'trPr');
        const fila = { encabezado: !!(trPr && hijo(trPr, 'tblHeader') && activo(hijo(trPr, 'tblHeader'))), celdas: [] };
        let col = 0;
        const tcs = hijosDeTabla(tr, 'tc');
        for (const tc of tcs) {
          const tcPr = hijo(tc, 'tcPr');
          const span = tcPr && hijo(tcPr, 'gridSpan') ? Math.max(parseInt(hijo(tcPr, 'gridSpan').a.val, 10) || 1, 1) : 1;
          const vm = tcPr && hijo(tcPr, 'vMerge');
          const continua = !!vm && (vm.a.val === undefined || vm.a.val === '' || vm.a.val === 'continue');
          if (continua && abiertas[col]) {
            abiertas[col].filasExtra++;
          } else {
            const celda = { span, filasExtra: 0, bloques: await this.bloques(tc), encabezado: fila.encabezado };
            fila.celdas.push(celda);
            if (vm) for (let k = 0; k < span; k++) abiertas[col + k] = celda;
            else for (let k = 0; k < span; k++) abiertas[col + k] = null;
          }
          col += span;
        }
        filas.push(fila);
      }
      return { k: 'tabla', filas };
    }

    // ───────── Bloques → HTML ─────────
    inlineAHtml(nodos, padre, ctxNotas) {
      for (const n of nodos) {
        switch (n.k) {
          case 't': {
            const etiquetas = [];
            if (n.f.b) etiquetas.push('strong');
            if (n.f.i) etiquetas.push('em');
            if (n.f.sup) etiquetas.push('sup');
            if (n.f.sub) etiquetas.push('sub');
            if (n.f.s) etiquetas.push('s');
            agregar(padre, etiquetas, n.s);
            break;
          }
          case 'br': padre.h.push(el('br', null)); break;
          case 'img': padre.h.push(el('img', n.alt ? { alt: n.alt, src: n.src } : { src: n.src })); break;
          case 'ancla': padre.h.push(el('a', { id: n.id })); break;
          case 'nota': {
            const a = el('a', { href: `#${n.tipo}-${n.id}`, id: `${n.tipo}-ref-${n.id}` }, [`[${n.n}]`]);
            padre.h.push(el('sup', null, [a]));
            break;
          }
          case 'a': {
            const ultimo = padre.h[padre.h.length - 1];
            if (n.href && ultimo && typeof ultimo !== 'string' && ultimo.e === 'a' && ultimo.attrs && ultimo.attrs.href === n.href) {
              this.inlineAHtml(n.hijos, ultimo);
            } else {
              const a = el('a', n.href ? { href: n.href } : null);
              this.inlineAHtml(n.hijos, a);
              if (a.h.length) padre.h.push(a);
            }
            break;
          }
          default: break;
        }
      }
    }

    vacio(n) { return typeof n !== 'string' ? n.h.length === 0 && n.e !== 'br' && n.e !== 'img' && !(n.e === 'a' && n.attrs && n.attrs.id) : n === ''; }

    bloquesAHtml(bloques) {
      const salida = [];
      let pila = []; // listas abiertas: { nivel, ordenada, nodo, ultimoLi }
      const cerrarTodo = () => { pila = []; };
      for (const b of bloques) {
        if (b.k === 'tabla') { cerrarTodo(); salida.push(this.tablaAHtml(b)); continue; }
        const nodo = el(b.etiqueta, null);
        this.inlineAHtml(b.hijos, nodo);
        const sinContenido = nodo.h.every((x) => this.vacio(x)) ;
        if (b.num) {
          if (sinContenido) continue;
          const { nivel, ordenada } = b.num;
          // La pila tiene una lista por nivel (0, 1, 2…). Si el documento salta
          // niveles, los que faltan se abren vacíos, como hace Word al mostrarlos.
          while (pila.length > nivel + 1) pila.pop();
          if (pila.length === nivel + 1 && pila[nivel].ordenada !== ordenada) pila.pop();
          while (pila.length < nivel + 1) {
            const lis = el(ordenada ? 'ol' : 'ul', null);
            const padre = pila[pila.length - 1];
            if (!padre) salida.push(lis);
            else {
              if (!padre.ultimoLi) { padre.ultimoLi = el('li', null); padre.nodo.h.push(padre.ultimoLi); }
              padre.ultimoLi.h.push(lis);
            }
            pila.push({ ordenada, nodo: lis, ultimoLi: null });
          }
          const tope = pila[nivel];
          const li = el('li', null, nodo.h);
          tope.nodo.h.push(li);
          tope.ultimoLi = li;
          continue;
        }
        cerrarTodo();
        if (sinContenido) continue;
        salida.push(nodo);
      }
      return salida;
    }

    tablaAHtml(t) {
      const tabla = el('table', null);
      const enc = el('thead', null);
      const cuerpo = el('tbody', null);
      let hayEnc = false;
      for (const f of t.filas) {
        const tr = el('tr', null);
        for (const c of f.celdas) {
          const attrs = {};
          if (c.span > 1) attrs.colspan = String(c.span);
          if (c.filasExtra > 0) attrs.rowspan = String(c.filasExtra + 1);
          const td = el(f.encabezado ? 'th' : 'td', Object.keys(attrs).length ? attrs : null, this.bloquesAHtml(c.bloques));
          tr.h.push(td);
        }
        if (f.encabezado) { hayEnc = true; enc.h.push(tr); } else cuerpo.h.push(tr);
      }
      if (hayEnc) { tabla.h.push(enc); tabla.h.push(cuerpo); } else tabla.h.push(...cuerpo.h);
      return tabla;
    }

    async notasAHtml() {
      const items = [];
      const vistas = { footnote: new Set(), endnote: new Set() };
      for (const tipo of ['footnote', 'endnote']) {
        const raiz = tipo === 'footnote' ? this.pieNotas : this.finNotas;
        if (!raiz) continue;
        const lis = [];
        for (const ref of this.notas.filter((r) => r.tipo === tipo)) {
          if (vistas[tipo].has(ref.id)) continue;
          vistas[tipo].add(ref.id);
          const nota = lista(raiz, tipo).find((x) => x.a.id === ref.id);
          if (!nota) continue;
          const bl = await this.bloques(nota);
          const ultimo = [...bl].reverse().find((b) => b.k === 'p');
          const volver = { k: 'a', href: `#${tipo}-ref-${ref.id}`, hijos: [{ k: 't', s: '↑', f: {} }] };
          if (ultimo) { ultimo.hijos.push({ k: 't', s: ' ', f: {} }, volver); } else bl.push({ k: 'p', etiqueta: 'p', num: null, hijos: [volver] });
          const li = el('li', { id: `${tipo}-${ref.id}` }, this.bloquesAHtml(bl));
          lis.push(li);
        }
        if (lis.length) items.push(el('ol', null, lis));
      }
      return items;
    }
  }

  // Filas o celdas de una tabla, también si un control de contenido (sdt) las envuelve, a cualquier profundidad.
  function hijosDeTabla(e, nombre, prof = 0) {
    const salida = [];
    for (const x of e.h) {
      if (x.n === nombre) salida.push(x);
      else if (prof < 20 && (x.n === 'sdt' || x.n === 'sdtContent' || x.n === 'ins' || x.n === 'customXml')) salida.push(...hijosDeTabla(x, nombre, prof + 1));
    }
    return salida;
  }

  // "HYPERLINK "url"" o "HYPERLINK \l "ancla""; devuelve null si el campo no es un enlace.
  function enlaceDeInstruccion(instr) {
    const m = /^\s*HYPERLINK\s+(.*)$/is.exec(instr);
    if (!m) return null;
    const resto = m[1];
    const ancla = /\\l\s+"([^"]*)"/i.exec(resto);
    const url = /^\s*"([^"]*)"/.exec(resto) || /^\s*([^\s\\"]+)/.exec(resto);
    let href = url ? url[1] : '';
    if (ancla) href += '#' + ancla[1];
    return href;
  }

  // ───────── Entrada ─────────
  async function abrir(datos, opciones) {
    const { zip, rels, cuerpo } = await abrirPaquete(datos);
    const parte = async (tipo, defecto) => {
      const ruta = porTipo(rels, tipo) || defecto;
      const g = zip.file(ruta);
      return g ? arbol(await g.async('string')) : null;
    };
    const estilosRaiz = await parte('/styles', 'word/styles.xml');
    const numRaiz = await parte('/numbering', 'word/numbering.xml');
    const partes = {
      rels,
      estilos: estilosRaiz ? leerEstilos(estilosRaiz) : new Map(),
      numeracion: numRaiz ? leerNumeracion(numRaiz) : { abstractos: new Map(), numeros: new Map() },
      pieNotas: await parte('/footnotes', 'word/footnotes.xml'),
      finNotas: await parte('/endnotes', 'word/endnotes.xml'),
    };
    return { lector: new Lector(zip, partes, opciones), cuerpo };
  }

  async function convertirAHtml(datos, opciones) {
    const { lector, cuerpo } = await abrir(datos, opciones);
    const bloques = await lector.bloques(cuerpo);
    const nodos = lector.bloquesAHtml(bloques);
    const notas = await lector.notasAHtml();
    return { value: nodos.concat(notas).map(serializar).join(''), messages: lector.mensajes };
  }

  function textoDeBloque(b, salida) {
    if (b.k === 'tabla') {
      for (const f of b.filas) for (const c of f.celdas) for (const x of c.bloques) textoDeBloque(x, salida);
      return;
    }
    let s = '';
    for (const n of b.hijos) s += textoDeNodo(n);
    salida.push(s);
  }
  function textoDeNodo(n) {
    if (n.k === 't') return n.s;
    if (n.k === 'br') return '\n';
    if (n.k === 'a') return n.hijos.map(textoDeNodo).join('');
    return '';
  }

  // Texto plano: cada párrafo seguido de una línea en blanco (como mammoth).
  // A diferencia de mammoth, un salto de línea (Shift+Enter) sale como "\n".
  async function extraerTexto(datos) {
    const { lector, cuerpo } = await abrir(datos, {});
    const bloques = await lector.bloques(cuerpo);
    const partes = [];
    for (const b of bloques) textoDeBloque(b, partes);
    return { value: partes.map((s) => s + '\n\n').join(''), messages: lector.mensajes };
  }

  globalThis.DocxLector = { convertirAHtml, extraerTexto };
})();

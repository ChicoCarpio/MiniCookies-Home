/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Lector de .odt propio (OpenDocument Text: LibreOffice, OpenOffice, Google Documentos, Word al guardar como OpenDocument). Lo entrega como el «contrato de documento»
// (ver docx-desde-contrato.js) y, con aDocx, como un .docx, que es lo que ya saben leer, mostrar e imprimir las herramientas de Word.
//
//   await OdtLector.parece(bytes)            → true si es un .odt (zip con content.xml y mimetype application/vnd.oasis.opendocument.text)
//   await OdtLector.leer(bytes, opciones?)   → { contrato, mensajes: [clave…] }
//   await OdtLector.aDocx(bytes)             → { bytes, mensajes }     (mensajes: claves de aviso, ver MENSAJES)
//
// Lanza Error (con .codigo = 'cifrado' si el .odt tiene contraseña) con un mensaje listo para mostrar.
// Se llevan: párrafos y títulos con su estilo (también los estilos con nombre), negrita, cursiva, subrayado, tachado, superíndice, subíndice, versalitas, mayúsculas, fuente, tamaño,
// color, resaltado, alineación, sangrías, espaciado, interlineado, tabulaciones, saltos de página, listas (numeradas y con viñetas, anidadas), tablas (celdas combinadas, relleno,
// bordes, anchos), imágenes en línea, enlaces, marcadores, números de página, secciones con su hoja (tamaño, márgenes, encabezado y pie). No se llevan: notas, comentarios, texto
// borrado con control de cambios, dibujos y fórmulas.
// Requiere zip-propio.js y xml-ligero.js (y docx-desde-contrato.js para aDocx).
(function () {
  'use strict';
  const MAX_PROFUNDIDAD = 64;
  const MAX_ELEMENTOS = 3000000;          // elementos del XML (un .odt real de cientos de páginas tiene unas decenas de miles)
  const MAX_TEXTO_XML = 256 * 1024 * 1024; // caracteres de content.xml
  const MAX_IMAGENES_BYTES = 200 * 1024 * 1024;
  const MAX_BLOQUES = 2000000;
  const MAX_FILAS_TABLA = 100000;
  const MAX_LISTAS = 20000;
  const MAX_IMAGENES = 5000;                 // referencias a imágenes (cada una es un elemento del .docx aunque apunte al mismo archivo)
  const MAX_CARACTERES = 64 * 1024 * 1024;   // texto de todo el documento (<text:s text:c> y las filas y celdas repetidas lo multiplican)
  const sinProto = (o) => Object.assign(Object.create(null), o);

  // ───────── XML → árbol con el texto en su lugar ─────────
  function arbol(xml) {
    const raiz = { n: '#', a: {}, h: [] };
    const pila = [raiz];
    let cuenta = 0;
    XmlLigero.recorrer(xml, {
      abrir(n, a) {
        if (++cuenta > MAX_ELEMENTOS) throw new Error('El documento tiene demasiados elementos y no se puede leer.');
        if (pila.length > MAX_PROFUNDIDAD * 4) throw new Error('El documento tiene demasiados niveles anidados y no se puede leer.');
        const e = { n, a, h: [] };
        pila[pila.length - 1].h.push(e);
        pila.push(e);
      },
      cerrar() { if (pila.length > 1) pila.pop(); },
      texto(t) { pila[pila.length - 1].h.push(t); },
    });
    return raiz.h.find((x) => typeof x !== 'string') || null;
  }
  const hijos = (e, n) => (e ? e.h.filter((x) => typeof x !== 'string' && (n === undefined || x.n === n)) : []);
  const hijo = (e, n) => (e ? e.h.find((x) => typeof x !== 'string' && x.n === n) || null : null);

  // ───────── Medidas ─────────
  // Una medida de OpenDocument ("0.5in", "1.25cm", "12pt") → twips (null si no es una medida).
  function twips(t, base) {
    const m = /^\s*(-?[\d.]+)\s*(cm|mm|in|pt|px|pc|%)?\s*$/.exec(t || '');
    if (!m) return null;
    const n = parseFloat(m[1]);
    if (!Number.isFinite(n)) return null;
    const u = m[2] || 'pt';
    if (u === '%') return base == null ? null : Math.round(base * n / 100);
    return Math.round(u === 'cm' ? n * 567 : u === 'mm' ? n * 56.6929 : u === 'in' ? n * 1440 : u === 'px' ? n * 15 : u === 'pc' ? n * 240 : n * 20);
  }
  const puntos = (t, base) => { const v = twips(t, base == null ? null : base * 20); return v == null ? null : v / 20; };
  const acotar = (x, min, max) => Math.min(Math.max(x, min), max);

  const RESALTADOS_HEX = sinProto({ '#ffff00': 'yellow', '#00ff00': 'green', '#00ffff': 'cyan', '#ff00ff': 'magenta', '#0000ff': 'blue', '#ff0000': 'red', '#000080': 'darkBlue', '#008080': 'darkCyan', '#008000': 'darkGreen', '#800080': 'darkMagenta', '#800000': 'darkRed', '#808000': 'darkYellow', '#808080': 'darkGray', '#c0c0c0': 'lightGray', '#000000': 'black', '#ffffff': 'white' });
  const FORMATO_NUMERO = sinProto({ '1': 'decimal', a: 'lowerLetter', A: 'upperLetter', i: 'lowerRoman', I: 'upperRoman' });

  // "0.5pt solid #000000" → { tam (1/8 pt), color, estilo } | null
  function bordeDe(t) {
    if (!t || t === 'none') return null;
    const partes = t.trim().split(/\s+/);
    const ancho = partes.map((p) => puntos(p)).find((v) => v != null && v > 0);
    const color = partes.find((p) => /^#[0-9a-f]{6}$/i.test(p));
    return { tam: acotar(Math.round((ancho == null ? 0.5 : ancho) * 8), 2, 96), color: color ? color.slice(1).toUpperCase() : undefined, estilo: partes.includes('double') ? 'doble' : partes.includes('dotted') ? 'punteado' : partes.includes('dashed') ? 'guiones' : 'simple' };
  }

  // ───────── Estilos ─────────
  class Estilos {
    constructor() {
      this.nombrados = new Map();    // 'familia|nombre' → estilo
      this.automaticos = new Map();
      this.fuentes = new Map();      // nombre de font-face → familia
      this.listas = new Map();       // nombre de list-style → niveles
      this.paginas = new Map();      // nombre de page-layout → propiedades
      this.maestras = new Map();     // nombre de master-page → { pagina, encabezado, pie }
      this.predeterminado = { paragraph: null, text: null };
    }

    leerDeclaraciones(raiz) {
      for (const f of hijos(hijo(raiz, 'font-face-decls'), 'font-face')) {
        const fam = (f.a['font-family'] || f.a.name || '').replace(/^["']+|["']+$/g, '');
        if (f.a.name && fam) this.fuentes.set(f.a.name, fam);
      }
    }

    leerEstilosDe(contenedor, automatico) {
      const destino = automatico ? this.automaticos : this.nombrados;
      for (const e of hijos(contenedor)) {
        if (e.n === 'style' && e.a.name) {
          const familia = e.a.family || 'paragraph';
          const est = { nombre: e.a.name, familia, padre: e.a['parent-style-name'] || null, siguiente: e.a['next-style-name'] || null, maestra: e.a['master-page-name'] || null, nivel: e.a['default-outline-level'] || null, lista: e.a['list-style-name'] || null, p: {}, r: {}, tabs: null, automatico };
          for (const prop of hijos(e)) this.propiedades(est, prop);
          destino.set(familia + '|' + e.a.name, est);
        } else if (e.n === 'default-style' && !automatico) {
          const est = { p: {}, r: {}, tabs: null };
          for (const prop of hijos(e)) this.propiedades(est, prop);
          this.predeterminado[e.a.family || 'paragraph'] = est;
        } else if (e.n === 'list-style' && e.a.name) {
          // por nivel (1 a 9; se queda el primero de cada uno): así nivelesDeLista no recorre los hijos una vez por nivel y por lista
          const niveles = [];
          for (const x of hijos(e)) {
            if (!/^list-level-style-/.test(x.n)) continue;
            const n = parseInt(x.a.level, 10);
            if (n >= 1 && n <= 9 && !niveles[n - 1]) niveles[n - 1] = { tipo: x.n.slice('list-level-style-'.length), a: x.a, props: hijo(x, 'list-level-properties'), texto: hijo(x, 'text-properties') };
          }
          this.listas.set(e.a.name, niveles);
        }
      }
    }

    propiedades(est, prop) {
      const a = prop.a;
      if (prop.n === 'text-properties') {
        const r = est.r;
        if (a['font-name'] !== undefined) r.fuente = this.fuentes.get(a['font-name']) || a['font-name'];
        if (a['font-family'] !== undefined) r.fuente = a['font-family'].replace(/^["']+|["']+$/g, '');
        if (a['font-name-asian'] !== undefined) r.fuenteAsia = this.fuentes.get(a['font-name-asian']) || a['font-name-asian'];
        else if (a['font-family-asian'] !== undefined) r.fuenteAsia = a['font-family-asian'].replace(/^["']+|["']+$/g, '');
        if (a['font-name-complex'] !== undefined) r.fuenteCs = this.fuentes.get(a['font-name-complex']) || a['font-name-complex'];
        else if (a['font-family-complex'] !== undefined) r.fuenteCs = a['font-family-complex'].replace(/^["']+|["']+$/g, '');
        if (a['font-size'] !== undefined) r.tamano = a['font-size'];
        if (a['font-weight'] !== undefined) r.negrita = a['font-weight'] === 'bold' || parseInt(a['font-weight'], 10) >= 600;
        if (a['font-style'] !== undefined) r.cursiva = a['font-style'] === 'italic' || a['font-style'] === 'oblique';
        if (a['text-underline-style'] !== undefined) r.subrayado = a['text-underline-style'] === 'none' ? false : (a['text-underline-type'] === 'double' ? 'doble' : true);
        else if (a['text-underline-type'] !== undefined) r.subrayado = a['text-underline-type'] === 'none' ? false : (a['text-underline-type'] === 'double' ? 'doble' : true);
        if (a['text-line-through-style'] !== undefined) r.tachado = a['text-line-through-style'] !== 'none';
        else if (a['text-line-through-type'] !== undefined) r.tachado = a['text-line-through-type'] !== 'none';
        if (a.color !== undefined) r.color = /^#[0-9a-f]{6}$/i.test(a.color) ? a.color.slice(1).toUpperCase() : undefined;
        if (a['background-color'] !== undefined) r.resaltado = RESALTADOS_HEX[String(a['background-color']).toLowerCase()] || null;
        if (a['text-position'] !== undefined) {
          const m = /^(super|sub|-?[\d.]+%)/.exec(a['text-position']);
          r.sup = !!m && (m[1] === 'super' || (m[1].endsWith('%') && parseFloat(m[1]) > 0));
          r.sub = !!m && (m[1] === 'sub' || (m[1].endsWith('%') && parseFloat(m[1]) < 0));
        }
        if (a['font-variant'] !== undefined) r.versalitas = a['font-variant'] === 'small-caps';
        if (a['text-transform'] !== undefined) r.mayusculas = a['text-transform'] === 'uppercase';
        if (a.display !== undefined) r.oculto = a.display === 'none';
      } else if (prop.n === 'paragraph-properties') {
        const p = est.p;
        if (a['text-align'] !== undefined) p.alineacion = { start: 'izquierda', left: 'izquierda', end: 'derecha', right: 'derecha', center: 'centro', justify: 'justificado' }[a['text-align']];
        if (a['margin-left'] !== undefined) p.izq = a['margin-left'];
        if (a['margin-right'] !== undefined) p.der = a['margin-right'];
        if (a['text-indent'] !== undefined) p.primera = a['text-indent'];
        if (a['margin-top'] !== undefined) p.antes = a['margin-top'];
        if (a['margin-bottom'] !== undefined) p.despues = a['margin-bottom'];
        if (a['line-height'] !== undefined) p.linea = a['line-height'];
        if (a['line-height-at-least'] !== undefined) p.lineaMinima = a['line-height-at-least'];
        if (a['break-before'] !== undefined) p.saltoAntes = a['break-before'] === 'page';
        if (a['keep-with-next'] !== undefined) p.mantener = a['keep-with-next'] === 'always';
        if (a['keep-together'] !== undefined) p.mantenerJuntas = a['keep-together'] === 'always';
        if (a['background-color'] !== undefined) p.relleno = /^#[0-9a-f]{6}$/i.test(a['background-color']) ? a['background-color'].slice(1).toUpperCase() : null;
        for (const lado of ['top', 'bottom', 'left', 'right']) if (a['border-' + lado] !== undefined) { (p.bordes = p.bordes || {})[{ top: 'arriba', bottom: 'abajo', left: 'izquierda', right: 'derecha' }[lado]] = bordeDe(a['border-' + lado]); }
        if (a.border !== undefined) { const b = bordeDe(a.border); p.bordes = { arriba: b, abajo: b, izquierda: b, derecha: b }; }
        const tabs = hijo(prop, 'tab-stops');
        if (tabs) est.tabs = hijos(tabs, 'tab-stop').slice(0, 64).map((t) => ({ pos: t.a.position, tipo: t.a.type, relleno: t.a['leader-style'] ? (t.a['leader-style'] === 'dotted' ? 'punto' : t.a['leader-style'] === 'solid' ? 'linea' : 'guion') : null }));
      } else if (prop.n === 'table-cell-properties') {
        const c = est.c = est.c || {};
        if (a['background-color'] !== undefined) c.relleno = /^#[0-9a-f]{6}$/i.test(a['background-color']) ? a['background-color'].slice(1).toUpperCase() : null;
        if (a.border !== undefined) { const b = bordeDe(a.border); c.bordes = { arriba: b, abajo: b, izquierda: b, derecha: b }; }
        for (const lado of ['top', 'bottom', 'left', 'right']) if (a['border-' + lado] !== undefined) { c.bordes = c.bordes || {}; c.bordes[{ top: 'arriba', bottom: 'abajo', left: 'izquierda', right: 'derecha' }[lado]] = bordeDe(a['border-' + lado]); }
        if (a['vertical-align'] !== undefined) c.alineacionV = { top: 'arriba', middle: 'centro', bottom: 'abajo' }[a['vertical-align']];
      } else if (prop.n === 'table-column-properties') {
        if (a['column-width'] !== undefined) est.ancho = a['column-width'];
      } else if (prop.n === 'section-properties') {
        const c = hijo(prop, 'columns');
        if (c) est.columnas = { n: acotar(parseInt(c.a['column-count'], 10) || 1, 1, 45), espacio: twips(c.a['column-gap']) == null ? 720 : twips(c.a['column-gap']) };
      } else if (prop.n === 'table-properties') {
        if (a.width !== undefined) est.anchoTabla = a.width;
        if (a['margin-left'] !== undefined) est.sangriaTabla = a['margin-left'];
      } else if (prop.n === 'table-row-properties') {
        if (a['row-height'] !== undefined) est.alto = a['row-height'];
        else if (a['min-row-height'] !== undefined) est.alto = a['min-row-height'];
      }
    }

    // El estilo y todos los que hereda, del más lejano al propio
    cadena(familia, nombre) {
      const lista = [];
      for (let i = 0, e = this.buscar(familia, nombre); e && i < 24; i++, e = e.padre ? this.buscar(familia, e.padre) : null) lista.unshift(e);
      return lista;
    }
    buscar(familia, nombre) { return this.automaticos.get(familia + '|' + nombre) || this.nombrados.get(familia + '|' + nombre) || null; }

    leerPaginas(raiz) {
      for (const pl of hijos(raiz, 'page-layout')) {
        const pr = hijo(pl, 'page-layout-properties');
        const a = pr ? pr.a : {};
        const cab = hijo(hijo(pl, 'header-style'), 'header-footer-properties'), pie = hijo(hijo(pl, 'footer-style'), 'header-footer-properties');
        this.paginas.set(pl.a.name, { a, cabecera: cab ? cab.a : null, pie: pie ? pie.a : null });
      }
    }
    leerMaestras(raiz) {
      for (const m of hijos(raiz, 'master-page')) {
        this.maestras.set(m.a.name, { pagina: m.a['page-layout-name'], encabezado: hijo(m, 'header'), pie: hijo(m, 'footer') });
        if (!this.primeraMaestra) this.primeraMaestra = m.a.name;
      }
    }
  }

  const decodificarNombre = (n) => String(n).replace(/_20_/g, ' ').replace(/_([0-9a-f]{2,4})_/gi, (t, h) => String.fromCharCode(parseInt(h, 16)));

  // ───────── Conversión ─────────
  class Conversor {
    constructor(estilos, zip) {
      this.est = estilos;
      this.zip = zip;
      this.mensajes = new Set();
      this.contrato = { propiedades: {}, predeterminado: {}, estilos: [], listas: [], secciones: [] };
      this.estilosUsados = new Set();
      this.bytesImagenes = 0;
      this.imagenes = 0;
      this.bloquesLeidos = 0;
      this.caracteres = 0;
      this.marcas = 0;
      this.maestraActual = null;
      this.sec = null;
      this.hayContenido = false;
      this.contadorListas = 0;
      this.ultimaLista = new Map(); // nombre de estilo de lista → id del contrato (para continuar la numeración)
      this.idsListas = new Map();   // xml:id de una lista → id del contrato (text:continue-list)
    }
    aviso(c) { this.mensajes.add(c); }

    // Propiedades de texto de un estilo → r del contrato
    resolverR(familia, nombre) {
      const r = {};
      let nombrado = null;
      for (const e of this.est.cadena(familia, nombre)) {
        if (e.automatico) Object.assign(r, e.r); else nombrado = e;   // los del padre con nombre ya viajan por el estilo; se parte del primer automático
      }
      // el estilo con nombre más cercano se referencia (no se copia)
      const propio = this.est.buscar(familia, nombre);
      return { estilo: propio && !propio.automatico ? propio.nombre : (nombrado ? nombrado.nombre : null), r: this.rFinal(r, null) };
    }

    rFinal(r, base) {
      const x = {};
      for (const k of Object.keys(r)) {
        const v = r[k];
        if (v === undefined || v === null) continue;
        if (k === 'tamano') { const pt = puntos(v, base); if (pt != null && pt > 0) x.tamano = acotar(pt, 1, 1638); }
        else x[k] = v;
      }
      return x;
    }

    pFinal(p, tabs, base) {
      const x = {};
      const t = (v) => twips(v, null);
      if (p.alineacion) x.alineacion = p.alineacion;
      if (p.izq !== undefined) { const v = t(p.izq); if (v != null) x.izq = v; }
      if (p.der !== undefined) { const v = t(p.der); if (v != null) x.der = v; }
      if (p.primera !== undefined) { const v = t(p.primera); if (v != null) x.primera = v; }
      if (p.antes !== undefined) { const v = t(p.antes); if (v != null) x.antes = v; }
      if (p.despues !== undefined) { const v = t(p.despues); if (v != null) x.despues = v; }
      if (p.linea !== undefined) {
        if (/%$/.test(p.linea)) { const v = parseFloat(p.linea); if (v > 0) { x.linea = Math.round(240 * v / 100); x.regla = 'auto'; } }
        else { const v = t(p.linea); if (v != null && v > 0) { x.linea = v; x.regla = 'exacto'; } }
      } else if (p.lineaMinima !== undefined) { const v = t(p.lineaMinima); if (v != null && v > 0) { x.linea = v; x.regla = 'minimo'; } }
      if (p.saltoAntes !== undefined) x.saltoAntes = p.saltoAntes;
      if (p.mantener) x.mantener = true;
      if (p.mantenerJuntas) x.mantenerJuntas = true;
      if (p.relleno) x.relleno = p.relleno;
      if (p.bordes) { const b = {}; for (const k of Object.keys(p.bordes)) if (p.bordes[k]) b[k] = p.bordes[k]; if (Object.keys(b).length) x.bordes = b; }
      if (tabs) { const l = tabs.map((q) => ({ tipo: { left: 'izquierda', center: 'centro', right: 'derecha', char: 'decimal' }[q.tipo] || 'izquierda', pos: t(q.pos), relleno: q.relleno })).filter((q) => q.pos != null && q.pos >= 0); if (l.length) x.tabs = l; }
      return x;
    }

    resolverP(nombre) {
      const p = {};
      let tabs = null, maestra = null, nivel = null;
      for (const e of this.est.cadena('paragraph', nombre)) {
        Object.assign(p, e.p);
        if (e.tabs) tabs = e.tabs;
        if (e.maestra) maestra = e.maestra;
        if (e.nivel) nivel = e.nivel;
      }
      return { p, tabs, maestra, nivel };
    }

    // ───────── Estilos con nombre al contrato ─────────
    estilosDelContrato() {
      const base = this.est.predeterminado;
      const r0 = Object.assign({}, base.text && base.text.r, base.paragraph && base.paragraph.r);
      this.contrato.predeterminado = { r: Object.assign({ tamano: 12 }, this.rFinal(r0, null)), p: this.pFinal(Object.assign({}, base.paragraph && base.paragraph.p), base.paragraph && base.paragraph.tabs) };
      const lista = [];
      for (const e of this.est.nombrados.values()) {
        if (e.familia !== 'paragraph' && e.familia !== 'text') continue;
        if (!this.estilosUsados.has(e.familia + '|' + e.nombre)) continue;
        const padre = e.padre && this.est.nombrados.get(e.familia + '|' + e.padre) ? decodificarNombre(e.padre) : null;
        const nombre = decodificarNombre(e.nombre);
        lista.push({ id: e.nombre, nombre, tipo: e.familia === 'text' ? 'caracter' : 'parrafo', basadoEn: padre ? e.padre : undefined, siguiente: e.siguiente || undefined,
          p: e.familia === 'paragraph' ? Object.assign(this.pFinal(e.p, e.tabs), e.nivel ? { nivelEsquema: parseInt(e.nivel, 10) - 1 } : {}) : undefined, r: this.rFinal(e.r, null) });
      }
      // los padres usados indirectamente también tienen que estar
      const ids = new Set(lista.map((e) => e.id));
      for (let pasada = 0; pasada < 24; pasada++) {
        let agregado = false;
        for (const e of lista.slice()) if (e.basadoEn && !ids.has(e.basadoEn)) {
          const fam = e.tipo === 'caracter' ? 'text' : 'paragraph';
          const padre = this.est.nombrados.get(fam + '|' + e.basadoEn);
          if (!padre) { e.basadoEn = undefined; continue; }
          const pn = padre.padre && this.est.nombrados.get(fam + '|' + padre.padre) ? padre.padre : null;
          lista.push({ id: padre.nombre, nombre: decodificarNombre(padre.nombre), tipo: e.tipo, basadoEn: pn || undefined, p: fam === 'paragraph' ? this.pFinal(padre.p, padre.tabs) : undefined, r: this.rFinal(padre.r, null) });
          ids.add(padre.nombre); agregado = true;
        }
        if (!agregado) break;
      }
      this.contrato.estilos = lista;
    }

    // ───────── Texto en línea ─────────
    nuevoParrafoEstado() { return { espacio: true }; }

    agregarTexto(salida, texto, estilo, r) {
      if (!texto) return;
      if ((this.caracteres += texto.length) > MAX_CARACTERES) throw new Error('El documento es demasiado grande para abrirlo aquí.');
      const u = salida[salida.length - 1];
      if (u && u.texto !== undefined && u.estilo === estilo && JSON.stringify(u.r || {}) === JSON.stringify(r || {})) { u.texto += texto; return; }
      const run = { texto };
      if (estilo) run.estilo = estilo;
      if (r && Object.keys(r).length) run.r = r;
      salida.push(run);
    }

    inline(nodo, fmt, salida, st, prof) {
      if (prof > MAX_PROFUNDIDAD) throw new Error('El documento tiene demasiados niveles anidados y no se puede leer.');
      for (const h of nodo.h) {
        if (typeof h === 'string') {
          let t = h.replace(/[ \t\r\n]+/g, ' ');
          if (st.espacio && t[0] === ' ') t = t.slice(1);
          if (!t) continue;
          st.espacio = t[t.length - 1] === ' ';
          this.agregarTexto(salida, t, fmt.estilo, fmt.r);
          continue;
        }
        const a = h.a;
        switch (h.n) {
          case 'span': {
            const f = a['style-name'] ? this.resolverR('text', a['style-name']) : null;
            if (f && f.estilo) this.estilosUsados.add('text|' + f.estilo);
            this.inline(h, f ? { estilo: f.estilo || fmt.estilo, r: Object.assign({}, fmt.r, f.r) } : fmt, salida, st, prof + 1);
            break;
          }
          case 'a': {
            const href = a.href || '';
            const destino = href[0] === '#' ? { ancla: href.slice(1) } : { url: href };
            const f = a['style-name'] ? this.resolverR('text', a['style-name']) : null;
            if (f && f.estilo) this.estilosUsados.add('text|' + f.estilo);
            const hijosRun = [];
            this.inline(h, f ? { estilo: f.estilo || fmt.estilo, r: Object.assign({}, fmt.r, f.r) } : fmt, hijosRun, st, prof + 1);
            if (hijosRun.length) salida.push({ enlace: destino, hijos: hijosRun });
            break;
          }
          case 's': { const n = acotar(parseInt(a.c, 10) || 1, 1, 10000); this.agregarTexto(salida, ' '.repeat(n), fmt.estilo, fmt.r); st.espacio = false; break; }
          case 'tab': this.agregarTexto(salida, '\t', fmt.estilo, fmt.r); st.espacio = true; break;
          case 'line-break': this.agregarTexto(salida, '\n', fmt.estilo, fmt.r); st.espacio = true; break;
          case 'bookmark': case 'bookmark-start': if (a.name && ++this.marcas < 100000) salida.push({ marcador: a.name }); break;
          case 'page-number': salida.push({ campo: 'PAGE', resultado: '1', r: fmt.r, estilo: fmt.estilo }); st.espacio = false; break;
          case 'page-count': salida.push({ campo: 'NUMPAGES', resultado: '1', r: fmt.r, estilo: fmt.estilo }); st.espacio = false; break;
          case 'frame': { const im = this.imagen(h); if (im) salida.push({ imagen: im }); break; }
          case 'note': this.aviso('notas'); break;
          case 'annotation': case 'annotation-end': case 'tracked-changes': case 'soft-page-break': case 'sequence-decls': case 'variable-decls': break;
          case 'change': case 'change-start': case 'change-end': break;
          case 'ruby': case 'meta': case 'meta-field': case 'date': case 'time': case 'title': case 'author-name': case 'file-name': case 'sender-firstname':
          case 'reference-ref': case 'bookmark-ref': case 'sequence': case 'page-continuation': case 'text-input': case 'user-defined': case 'chapter':
            this.inline(h, fmt, salida, st, prof + 1); break;
          default: break;
        }
      }
    }

    imagen(frame) {
      const img = hijo(frame, 'image');
      if (!img) { if (hijo(frame, 'text-box') || hijo(frame, 'object')) this.aviso('dibujos'); return null; }
      let ruta = img.a.href || '';
      if (!ruta || /^[a-z]+:/i.test(ruta) || ruta.includes('..')) { this.aviso('imagen-externa'); return null; }
      ruta = ruta.replace(/^\.\//, '');
      const entrada = this.zip.file(ruta);
      if (!entrada) { this.aviso('imagen-falta'); return null; }
      if (++this.imagenes > MAX_IMAGENES) throw new Error('El documento tiene demasiadas imágenes y no se puede leer.');
      return { ruta, ancho: puntos(frame.a.width) || 72, alto: puntos(frame.a.height) || 72, texto: (hijo(frame, 'desc') && hijo(frame, 'desc').h.filter((x) => typeof x === 'string').join('')) || undefined, entrada };
    }

    // ───────── Párrafos, listas, tablas ─────────
    parrafo(nodo, lista) {
      const nombreEstilo = nodo.a['style-name'] || '';
      const { p, tabs, maestra, nivel } = this.resolverP(nombreEstilo);
      const propio = this.est.buscar('paragraph', nombreEstilo);
      const bloque = { tipo: 'parrafo', hijos: [] };
      const nombrado = this.est.cadena('paragraph', nombreEstilo).slice().reverse().find((e) => !e.automatico);
      if (nombrado) { bloque.estilo = nombrado.nombre; this.estilosUsados.add('paragraph|' + nombrado.nombre); }
      // Solo se copian al párrafo las propiedades que fija un estilo automático (las del estilo con nombre viajan por el estilo)
      const propias = {};
      let tabsPropias = null;
      for (const e of this.est.cadena('paragraph', nombreEstilo)) if (e.automatico) { Object.assign(propias, e.p); if (e.tabs) tabsPropias = e.tabs; }
      const pc = this.pFinal(propias, tabsPropias);
      if (nodo.n === 'h') {
        const n = acotar(parseInt(nodo.a['outline-level'], 10) || (nivel ? parseInt(nivel, 10) : 1), 1, 10);
        pc.nivelEsquema = Math.min(n - 1, 8);
      }
      if (lista) pc.lista = { id: lista.id, nivel: lista.nivel };
      if (maestra && maestra !== this.maestraActual && !this.enCabecera) this.nuevaSeccion(maestra);
      if (!this.hayContenido && pc.saltoAntes) delete pc.saltoAntes;   // el salto de página «antes» del primer párrafo solo marca la hoja con la que empieza el documento
      if (Object.keys(pc).length) bloque.p = pc;
      const fmt = { estilo: null, r: {} };
      this.inline(nodo, fmt, bloque.hijos, this.nuevoParrafoEstado(), 0);
      // una imagen se resuelve al final para no tener la entrada del zip en el contrato
      void propio;
      return bloque;
    }

    lista(nodo, nivel, salida, idLista, prof) {
      if (prof > MAX_PROFUNDIDAD) throw new Error('El documento tiene demasiados niveles anidados y no se puede leer.');
      const nombre = nodo.a['style-name'] || nodo.a['list-style-name'] || '';
      let nueva = false;
      if (idLista === null) {
        const previa = (nodo.a['continue-list'] && this.idsListas.get(nodo.a['continue-list'])) || (nodo.a['continue-numbering'] === 'true' && this.ultimaLista.get(nombre));
        if (previa) idLista = previa;
        else {
          idLista = this.nuevaLista(nombre);
          nueva = true;
        }
        if (nodo.a.id) this.idsListas.set(nodo.a.id, idLista);
      }
      let primero = true;
      for (const item of hijos(nodo)) {
        if (item.n !== 'list-item' && item.n !== 'list-header') continue;
        // text:start-value en un elemento: la numeración sigue desde ese número (una lista nueva en el contrato, salvo que sea el primero de una lista recién creada)
        const inicio = parseInt(item.a['start-value'], 10);
        if (Number.isFinite(inicio) && inicio >= 0) {
          const valor = Math.min(inicio, 32767);
          if (!(nueva && primero)) {
            idLista = this.nuevaLista(nombre);
            if (nodo.a.id) this.idsListas.set(nodo.a.id, idLista);
          }
          if (this.contadorListas < MAX_LISTAS || (nueva && primero)) this.contrato.listas[this.contrato.listas.length - 1].niveles[Math.min(nivel, 8)].inicio = valor;
        }
        primero = false;
        for (const h of hijos(item)) {
          if (h.n === 'p' || h.n === 'h') {
            const b = this.parrafoDeLista(h, idLista, nivel);
            salida.push(b);
          } else if (h.n === 'list') this.lista(h, nivel + 1, salida, idLista, prof + 1);
          else this.bloque(h, salida, prof + 1);
        }
      }
    }
    // Una lista nueva en el contrato (con sus 9 niveles). Pasadas MAX_LISTAS (solo un archivo hostil) se sigue con la última: la numeración no reinicia.
    nuevaLista(nombre) {
      if (this.contadorListas >= MAX_LISTAS) return this.contrato.listas[this.contrato.listas.length - 1].id;
      const id = 'L' + (++this.contadorListas);
      this.contrato.listas.push({ id, niveles: this.nivelesDeLista(nombre) });
      this.ultimaLista.set(nombre, id);
      return id;
    }
    parrafoDeLista(nodo, id, nivel) {
      const b = this.parrafo(nodo, { id, nivel: Math.min(nivel, 8) });
      this.hayContenido = true;
      return b;
    }

    nivelesDeLista(nombre) {
      const def = this.est.listas.get(nombre) || [];   // por nivel: def[0] es el nivel 1
      const niveles = [];
      for (let n = 0; n < 9; n++) {
        const d = def[n];
        if (!d) { niveles.push({ formato: 'decimal', texto: '%' + (n + 1) + '.', izquierda: 720 * (n + 1), colgante: 360 }); continue; }
        const al = d.props ? hijo(d.props, 'list-level-label-alignment') : null;
        let izq, col;
        if (al) { izq = twips(al.a['margin-left']); const ti = twips(al.a['text-indent']); col = ti == null ? 360 : -ti; }
        else if (d.props) { const antes = twips(d.props.a['space-before']) || 0, ancho = twips(d.props.a['min-label-width']) || 360; izq = antes + ancho; col = ancho; }
        const nivelOut = { formato: 'decimal', texto: '', izquierda: acotar(izq == null ? 720 * (n + 1) : izq, 0, 31680), colgante: acotar(col == null ? 360 : col, 0, 31680) };
        if (d.tipo === 'bullet') {
          nivelOut.formato = 'bullet';
          nivelOut.texto = d.a['bullet-char'] || '•';
          const fn = d.texto && d.texto.a['font-name'];
          if (fn) nivelOut.fuente = this.est.fuentes.get(fn) || fn;
        } else if (d.tipo === 'number') {
          nivelOut.formato = FORMATO_NUMERO[d.a['num-format']] || (d.a['num-format'] === '' ? 'none' : 'decimal');
          const mostrar = acotar(parseInt(d.a['display-levels'], 10) || 1, 1, n + 1);
          let t = d.a['num-prefix'] || '';
          for (let k = n + 2 - mostrar; k <= n + 1; k++) t += '%' + k + (k < n + 1 ? '.' : '');
          nivelOut.texto = t + (d.a['num-suffix'] || '');
          if (d.a['start-value']) nivelOut.inicio = parseInt(d.a['start-value'], 10) || 1;
        } else { nivelOut.formato = 'bullet'; nivelOut.texto = '•'; }
        niveles.push(nivelOut);
      }
      return niveles;
    }

    tabla(nodo, salida, prof) {
      if (prof > 8) return;
      const est = this.est.buscar('table', nodo.a['style-name'] || '');
      const anchos = [];
      const columnas = [];
      const recolectar = (c) => {
        for (const x of hijos(c)) {
          if (x.n === 'table-column') {
            const e = this.est.buscar('table-column', x.a['style-name'] || '');
            const w = e && e.ancho ? twips(e.ancho) : null;
            const rep = acotar(parseInt(x.a['number-columns-repeated'], 10) || 1, 1, 1024);
            for (let i = 0; i < rep && columnas.length < 1024; i++) columnas.push(w == null ? 1440 : w);
          } else if (x.n === 'table-columns' || x.n === 'table-header-columns') recolectar(x);
        }
      };
      recolectar(nodo);
      void anchos;
      const filas = [];
      const pendiente = [];   // por columna: { resto, span, ancho }
      let totalFilas = 0;
      const procesarFila = (fila, encabezado) => {
        const rep = acotar(parseInt(fila.a['number-rows-repeated'], 10) || 1, 1, 1000);
        // cuentan las filas ya repetidas: una fila vacía con number-rows-repeated="1000" son 1000 filas del contrato
        totalFilas += rep;
        if (totalFilas > MAX_FILAS_TABLA) throw new Error('La tabla tiene demasiadas filas y no se puede leer.');
        if ((this.bloquesLeidos += rep) > MAX_BLOQUES) throw new Error('El documento tiene demasiados párrafos y no se puede leer.');
        const estFila = this.est.buscar('table-row', fila.a['style-name'] || '');
        for (let r = 0; r < rep; r++) {
          const celdas = [];
          let col = 0, saltarH = 0;
          const ancho = (desde, n) => { let w = 0; for (let i = 0; i < n; i++) w += columnas[desde + i] == null ? 1440 : columnas[desde + i]; return w; };
          for (const c of hijos(fila)) {
            if (c.n !== 'table-cell' && c.n !== 'covered-table-cell') continue;
            const repC = acotar(parseInt(c.a['number-columns-repeated'], 10) || 1, 1, 1024);
            for (let k = 0; k < repC; k++) {
              if (col >= 1024) break;
              if (++this.bloquesLeidos > MAX_BLOQUES) throw new Error('El documento tiene demasiados párrafos y no se puede leer.');
              if (saltarH > 0) { saltarH--; col++; continue; }
              const p = pendiente[col];
              if (c.n === 'covered-table-cell' || (p && p.resto > 0)) {
                if (p && p.resto > 0) {
                  p.resto--;
                  celdas.push({ hijos: [{ tipo: 'parrafo', hijos: [] }], ancho: ancho(col, p.span), fusionH: p.span > 1 ? p.span : undefined, fusionV: 'continua' });
                  saltarH = p.span - 1;
                  col += 1;
                } else col++;
                continue;
              }
              const spanH = acotar(parseInt(c.a['number-columns-spanned'], 10) || 1, 1, 63), spanV = acotar(parseInt(c.a['number-rows-spanned'], 10) || 1, 1, 1000);
              const ec = this.est.buscar('table-cell', c.a['style-name'] || '');
              const celda = { hijos: [], ancho: ancho(col, spanH) };
              if (spanH > 1) celda.fusionH = spanH;
              if (spanV > 1) { celda.fusionV = 'inicio'; pendiente[col] = { resto: spanV - 1, span: spanH }; }
              else if (pendiente[col]) pendiente[col] = null;
              if (ec && ec.c) {
                if (ec.c.relleno) celda.relleno = ec.c.relleno;
                if (ec.c.bordes) { const b = {}; for (const lado of Object.keys(ec.c.bordes)) if (ec.c.bordes[lado]) b[lado] = ec.c.bordes[lado]; if (Object.keys(b).length) celda.bordes = b; }
                if (ec.c.alineacionV) celda.alineacionV = ec.c.alineacionV;
              }
              for (const h of hijos(c)) this.bloque(h, celda.hijos, prof + 1);
              if (!celda.hijos.length) celda.hijos.push({ tipo: 'parrafo', hijos: [] });
              celdas.push(celda);
              saltarH = spanH - 1;
              col += 1;
            }
          }
          const f = { celdas };
          if (encabezado) f.encabezado = true;
          const alto = estFila && estFila.alto ? twips(estFila.alto) : null;
          if (alto) f.alto = alto;
          filas.push(f);
        }
      };
      const filasDe = (c, encabezado) => {
        for (const x of hijos(c)) {
          if (x.n === 'table-row') procesarFila(x, encabezado);
          else if (x.n === 'table-header-rows') filasDe(x, true);
          else if (x.n === 'table-row-group') filasDe(x, encabezado);
        }
      };
      filasDe(nodo, false);
      if (!filas.length) return;
      const t = { tipo: 'tabla', anchos: columnas, bordes: false, filas };
      if (est && est.sangriaTabla) t.sangria = twips(est.sangriaTabla) || undefined;
      salida.push(t);
    }

    bloque(nodo, salida, prof) {
      if (prof > MAX_PROFUNDIDAD) throw new Error('El documento tiene demasiados niveles anidados y no se puede leer.');
      if (++this.bloquesLeidos > MAX_BLOQUES) throw new Error('El documento tiene demasiados párrafos y no se puede leer.');
      switch (nodo.n) {
        case 'p': case 'h': {
          const b = this.parrafo(nodo, null);   // un párrafo con otra hoja (master-page-name) abre una sección: va en la nueva, no en la que estaba
          this.hayContenido = true;
          (prof === 0 && !this.enCabecera ? this.sec.hijos : salida).push(b);
          break;
        }
        case 'list': this.lista(nodo, 0, salida, null, prof + 1); break;
        case 'table': this.tabla(nodo, salida, prof + 1); this.hayContenido = true; break;
        case 'section': {
          const est = this.est.buscar('section', nodo.a['style-name'] || '');
          const cols = est && est.columnas && est.columnas.n > 1 ? est.columnas : null;
          if (cols && prof === 0 && !this.enCabecera) {
            // una sección con columnas es un tramo continuo de la misma hoja con ese número de columnas
            this.abrirSeccion({ tipo: 'continua', columnas: cols });
            for (const h of hijos(nodo)) this.bloque(h, this.sec.hijos, 0);
            this.abrirSeccion({ tipo: 'continua', columnas: null });
          } else for (const h of hijos(nodo)) this.bloque(h, salida, prof + 1);
          break;
        }
        case 'index': case 'table-of-content': case 'illustration-index': case 'alphabetical-index': case 'bibliography': case 'index-body': case 'text-box':
          for (const h of hijos(nodo)) { if (h.n === 'index-title' || h.n === 'index-body') { for (const g of hijos(h)) this.bloque(g, salida, prof + 1); } else this.bloque(h, salida, prof + 1); }
          break;
        default: break;
      }
    }

    // ───────── Hojas ─────────
    abrirSeccion(extra) {
      if (this.sec && !this.sec.hijos.length) { Object.assign(this.sec, extra); return; }
      this.sec = Object.assign({ maestra: this.maestraActual, hijos: [] }, extra);
      this.secs.push(this.sec);
    }

    nuevaSeccion(maestra) {
      if (this.sec && !this.sec.hijos.length) { this.sec.maestra = maestra; this.maestraActual = maestra; return; }
      this.maestraActual = maestra;
      this.sec = { maestra, hijos: [] };
      this.secs.push(this.sec);
    }

    seccionAlContrato(s) {
      const m = this.est.maestras.get(s.maestra) || (this.est.primeraMaestra && this.est.maestras.get(this.est.primeraMaestra));
      const pag = m && this.est.paginas.get(m.pagina);
      const a = pag ? pag.a : {};
      const ancho = twips(a['page-width']) || 11906, alto = twips(a['page-height']) || 16838;
      const margenes = { arriba: twips(a['margin-top']), abajo: twips(a['margin-bottom']), izquierda: twips(a['margin-left']), derecha: twips(a['margin-right']) };
      for (const k of Object.keys(margenes)) if (margenes[k] == null) margenes[k] = 1134;
      const out = { pagina: { ancho: acotar(ancho, 720, 31680), alto: acotar(alto, 720, 31680), horizontal: a['print-orientation'] === 'landscape' || ancho > alto }, margenes, tipo: s.tipo || 'siguiente', hijos: s.hijos };
      if (s.columnas) out.columnas = s.columnas;
      const cuerpoDe = (nodo) => { const l = []; this.enCabecera = true; if (nodo) for (const h of hijos(nodo)) this.bloque(h, l, 0); this.enCabecera = false; return l; };
      if (m && m.encabezado) {
        const l = cuerpoDe(m.encabezado);
        if (l.length) {
          out.encabezado = l;
          const alt = pag && pag.cabecera && (twips(pag.cabecera['min-height']) || 0) + (twips(pag.cabecera['margin-bottom']) || 0);
          out.margenes.encabezado = margenes.arriba;
          out.margenes.arriba = margenes.arriba + (alt || 567);
        }
      }
      if (m && m.pie) {
        const l = cuerpoDe(m.pie);
        if (l.length) {
          out.pie = l;
          const alt = pag && pag.pie && (twips(pag.pie['min-height']) || 0) + (twips(pag.pie['margin-top']) || 0);
          out.margenes.pie = margenes.abajo;
          out.margenes.abajo = margenes.abajo + (alt || 567);
        }
      }
      return out;
    }

    convertir(contenido, estilosRaiz) {
      this.secs = [];
      const cuerpo = hijo(hijo(contenido, 'body'), 'text');
      if (!cuerpo) throw new Error('El archivo no es un documento de texto de OpenDocument (.odt).');
      this.nuevaSeccion(null);
      this.maestraActual = null;
      for (const h of hijos(cuerpo)) this.bloque(h, this.sec.hijos, 0);
      this.contrato.secciones = this.secs.map((s) => this.seccionAlContrato(s));
      this.estilosDelContrato();
      void estilosRaiz;
      return this.contrato;
    }
  }

  // Tamaño que declara el zip para una entrada ya descomprimida (ZipPropio lo verifica al inflar: si el real es mayor, falla); 0 si no se sabe
  const tamanoDe = (entrada) => (entrada && entrada._origen && entrada._origen.real) || 0;

  async function abrir(bytes) {
    const datos = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    let zip;
    try { zip = await ZipPropio.loadAsync(datos); } catch (e) { throw new Error('El archivo no es un .odt válido (no se puede abrir como un paquete).'); }
    if (zip.file('encrypted-package') || (zip.file('META-INF/manifest.xml') && /encryption-data/.test(await zip.file('META-INF/manifest.xml').async('string')))) {
      const err = new Error('El documento está protegido con contraseña. Sacale la contraseña en LibreOffice y volvé a intentarlo.');
      err.codigo = 'cifrado';
      throw err;
    }
    if (!zip.file('content.xml')) throw new Error('El archivo no es un documento de OpenDocument (.odt).');
    return zip;
  }

  async function parece(bytes) {
    try {
      const datos = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
      if (datos.length < 4 || datos[0] !== 0x50 || datos[1] !== 0x4B) return false;
      const zip = await ZipPropio.loadAsync(datos);
      if (!zip.file('content.xml')) return false;
      const mt = zip.file('mimetype') ? (await zip.file('mimetype').async('string')).trim() : '';
      return mt === 'application/vnd.oasis.opendocument.text' || mt === 'application/vnd.oasis.opendocument.text-template' || (!mt && /<office:text|:text[ >]/.test((await zip.file('content.xml').async('string')).slice(0, 20000)));
    } catch (e) { return false; }
  }

  async function leer(bytes) {
    const zip = await abrir(bytes);
    const texto = async (ruta) => {
      const f = zip.file(ruta);
      if (!f) return null;
      if (tamanoDe(f) > MAX_TEXTO_XML) throw new Error('El documento es demasiado grande para abrirlo aquí.');   // antes de descomprimir
      const t = await f.async('string');
      if (t.length > MAX_TEXTO_XML) throw new Error('El documento es demasiado grande para abrirlo aquí.');
      return t;
    };
    const estilos = new Estilos();
    const contenidoXml = await texto('content.xml');
    const contenido = arbol(contenidoXml);
    const estilosXml = await texto('styles.xml');
    const estilosRaiz = estilosXml ? arbol(estilosXml) : null;
    if (!contenido) throw new Error('El archivo no es un documento de OpenDocument (.odt).');
    if (estilosRaiz) { estilos.leerDeclaraciones(estilosRaiz); }
    estilos.leerDeclaraciones(contenido);
    if (estilosRaiz) {
      estilos.leerEstilosDe(hijo(estilosRaiz, 'styles'), false);
      estilos.leerEstilosDe(hijo(estilosRaiz, 'automatic-styles'), true);
      estilos.leerPaginas(hijo(estilosRaiz, 'automatic-styles'));
      estilos.leerMaestras(hijo(estilosRaiz, 'master-styles'));
    }
    estilos.leerEstilosDe(hijo(contenido, 'automatic-styles'), true);
    const conv = new Conversor(estilos, zip);
    const contrato = conv.convertir(contenido, estilosRaiz);
    // Las imágenes se leen una por una (con tope de bytes) y se enlazan donde estaban
    const pendientes = [];
    const recorrer = (bloques) => {
      for (const b of bloques) {
        if (b.tipo === 'tabla') { for (const f of b.filas) for (const c of f.celdas) recorrer(c.hijos); continue; }
        const visitar = (runs) => { for (const r of runs) { if (r.imagen && r.imagen.entrada) pendientes.push(r); if (r.hijos) visitar(r.hijos); } };
        visitar(b.hijos || []);
      }
    };
    for (const s of contrato.secciones) { recorrer(s.hijos); if (s.encabezado) recorrer(s.encabezado); if (s.pie) recorrer(s.pie); }
    // Una imagen que se repite se descomprime una vez y todas sus referencias comparten el mismo arreglo (el escritor también la guarda una sola vez)
    const leidas = new Map();   // ruta en el zip → datos
    for (const r of pendientes) {
      const e = r.imagen.entrada;
      delete r.imagen.entrada;
      let datos = leidas.get(r.imagen.ruta);
      if (!datos) {
        if (conv.bytesImagenes + tamanoDe(e) > MAX_IMAGENES_BYTES) throw new Error('Las imágenes del documento ocupan demasiado y no se puede leer.');   // antes de descomprimir
        datos = await e.async('uint8array');
        conv.bytesImagenes += datos.length;
        if (conv.bytesImagenes > MAX_IMAGENES_BYTES) throw new Error('Las imágenes del documento ocupan demasiado y no se puede leer.');
        leidas.set(r.imagen.ruta, datos);
      }
      r.imagen.datos = datos;
    }
    const meta = await texto('meta.xml');
    if (meta) {
      const m = arbol(meta);
      const md = hijo(hijo(m, 'meta'), 'title');
      const t = md ? md.h.filter((x) => typeof x === 'string').join('') : '';
      const au = hijo(hijo(m, 'meta'), 'initial-creator');
      if (t) contrato.propiedades.titulo = t.slice(0, 255);
      if (au) contrato.propiedades.autor = au.h.filter((x) => typeof x === 'string').join('').slice(0, 255);
    }
    return { contrato, mensajes: [...conv.mensajes] };
  }

  async function aDocx(bytes) {
    const { contrato, mensajes } = await leer(bytes);
    const r = await DocxDesdeContrato.generar(contrato);
    return { bytes: r.bytes, mensajes: [...new Set([...mensajes, ...r.mensajes])] };
  }

  globalThis.OdtLector = { parece, leer, aDocx };
})();

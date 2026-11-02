/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Markdown → HTML propio (reemplaza a marked). Sigue CommonMark con las
// extensiones de GitHub que usa el sitio: tablas, tachado (~~x~~), listas de
// tareas y enlaces automáticos (https://…, www.…, correos). Genera el HTML con
// el mismo formato que marked (saltos de línea entre bloques, `<br>` sin
// cerrar, etc.) para que lo que se hacía con él siga igual.
//
// Uso:  MarkdownPropio.aHtml(texto)  →  cadena HTML
//
// El HTML del texto puede traer etiquetas propias (Markdown permite HTML
// escrito a mano) y se entrega tal cual: quien lo inserta en la página lo pasa
// antes por `sanitizarHtml` (seguridad-html.js), como siempre.
//
// Estructura: primero el árbol de bloques (citas, listas, código, títulos,
// tablas…, línea por línea), luego el texto de cada bloque (énfasis, enlaces,
// código en línea…) con una pila de delimitadores. No depende del navegador.

(function () {
  'use strict';

  // ───────── Nodos ─────────
  class Nodo {
    constructor(tipo) {
      this.tipo = tipo;
      this.padre = this.primero = this.ultimo = this.prev = this.sig = null;
      this.literal = null;
      this.abierto = true;
      this.ultimaLineaVacia = false;
    }
    agregar(h) {
      h.quitar();
      h.padre = this;
      if (this.ultimo) { this.ultimo.sig = h; h.prev = this.ultimo; this.ultimo = h; } else { this.primero = this.ultimo = h; }
    }
    insertarDespues(h) {
      h.quitar();
      h.prev = this; h.sig = this.sig; h.padre = this.padre;
      if (this.sig) this.sig.prev = h; else if (this.padre) this.padre.ultimo = h;
      this.sig = h;
    }
    quitar() {
      if (this.prev) this.prev.sig = this.sig; else if (this.padre) this.padre.primero = this.sig;
      if (this.sig) this.sig.prev = this.prev; else if (this.padre) this.padre.ultimo = this.prev;
      this.padre = this.sig = this.prev = null;
    }
  }

  // ───────── Expresiones ─────────
  const RE_ESPECIAL = /^[#`~*+_=<>0-9|-]/;
  const RE_SALTO = /^(?:\*[ \t]*){3,}$|^(?:_[ \t]*){3,}$|^(?:-[ \t]*){3,}$/;
  const RE_SETEXT = /^(?:=+|-+)[ \t]*$/;
  const RE_FILA_TABLA = /^\|?[ \t]*:?-+:?[ \t]*(?:\|[ \t]*:?-+:?[ \t]*)*\|?[ \t]*$/;
  const NOMBRE_ETIQUETA = '[A-Za-z][A-Za-z0-9-]*';
  const NOMBRE_ATRIBUTO = '[a-zA-Z_:][a-zA-Z0-9:._-]*';
  const VALOR = '(?:[^"\'=<>`\\x00-\\x20]+|\'[^\']*\'|"[^"]*")';
  const ATRIBUTO = '(?:\\s+' + NOMBRE_ATRIBUTO + '(?:\\s*=\\s*' + VALOR + ')?)';
  const ETIQUETA_ABRE = '<' + NOMBRE_ETIQUETA + ATRIBUTO + '*\\s*/?>';
  const ETIQUETA_CIERRA = '</' + NOMBRE_ETIQUETA + '\\s*>';
  const HTML_EN_LINEA = new RegExp('^(?:' + ETIQUETA_ABRE + '|' + ETIQUETA_CIERRA + '|<!---->|<!--(?:-?[^>-])(?:-?[^-])*-->|<[?][\\s\\S]*?[?]>|<![A-Za-z]+[^>]*>|<!\\[CDATA\\[[\\s\\S]*?\\]\\]>)');
  const HTML_BLOQUE_ABRE = [
    /./,
    /^<(?:script|pre|textarea|style)(?:\s|>|$)/i,
    /^<!--/,
    /^<[?]/,
    /^<![A-Za-z]/,
    /^<!\[CDATA\[/,
    /^<[/]?(?:address|article|aside|base|basefont|blockquote|body|caption|center|col|colgroup|dd|details|dialog|dir|div|dl|dt|fieldset|figcaption|figure|footer|form|frame|frameset|h[123456]|head|header|hr|html|iframe|legend|li|link|main|menu|menuitem|nav|noframes|ol|optgroup|option|p|param|section|source|summary|table|tbody|td|tfoot|th|thead|title|tr|track|ul)(?:\s|[/]?[>]|$)/i,
    new RegExp('^(?:' + ETIQUETA_ABRE + '|' + ETIQUETA_CIERRA + ')\\s*$', 'i'),
  ];
  const HTML_BLOQUE_CIERRA = [/./, /<\/(?:script|pre|textarea|style)>/i, /-->/, /\?>/, />/, /\]\]>/];
  const RE_PUNTUACION = /[\p{P}\p{S}]/u;
  const RE_ESPACIO_UNICODE = /^[\s  -   　]/;

  const esEspTab = (c) => c === ' ' || c === '\t';
  const PUNTUACION_ASCII = '!"#$%&\'()*+,-./:;<=>?@[\\]^_`{|}~';

  // ───────── Bloques ─────────
  class AnalizadorBloques {
    constructor() {
      this.doc = new Nodo('documento');
      this.tip = this.doc;
      this.refs = Object.create(null);
    }

    analizar(texto) {
      const lineas = texto.replace(/\r\n?/g, '\n').replace(/\u0000/g, '�').split('\n');
      if (lineas.length && lineas[lineas.length - 1] === '') lineas.pop();
      this.numLinea = 0;
      for (const l of lineas) { this.numLinea++; this.linea(l); }
      while (this.tip) this.cerrar(this.tip);
      return this.doc;
    }

    // ── utilidades de la línea actual ──
    buscarNoEspacio() {
      const l = this.l;
      let i = this.offset, cols = this.column, c;
      while ((c = l[i]) !== undefined) {
        if (c === ' ') { i++; cols++; } else if (c === '\t') { i++; cols += 4 - (cols % 4); } else break;
      }
      this.vacia = c === undefined;
      this.sigNoEsp = i;
      this.sigNoEspCol = cols;
      this.sangria = cols - this.column;
      this.sangrado = this.sangria >= 4;
    }
    avanzar(cuenta, columnas) {
      const l = this.l;
      let c;
      while (cuenta > 0 && (c = l[this.offset]) !== undefined) {
        if (c === '\t') {
          const hastaTab = 4 - (this.column % 4);
          if (columnas) this.tabParcial = hastaTab > cuenta;
          const avance = Math.min(cuenta, hastaTab);
          this.column += avance;
          this.offset += this.tabParcial ? 0 : 1;
          cuenta -= avance;
        } else {
          this.tabParcial = false;
          this.offset++; this.column++; cuenta--;
        }
      }
    }
    avanzarNoEsp() { this.offset = this.sigNoEsp; this.column = this.sigNoEspCol; this.tabParcial = false; }
    agregarLinea() {
      if (this.tabParcial) {
        this.offset += 1;
        const hastaTab = 4 - (this.column % 4);
        this.tip.contenido += ' '.repeat(hastaTab);
      }
      this.tip.contenido += this.l.slice(this.offset) + '\n';
    }
    aceptaLineas(t) { return t === 'parrafo' || t === 'codigo' || t === 'html'; }
    puedeContener(padre, hijo) {
      const t = padre.tipo;
      if (t === 'documento' || t === 'cita' || t === 'item') return hijo !== 'item';
      if (t === 'lista') return hijo === 'item';
      return false;
    }
    agregarHijo(tipo) {
      while (!this.puedeContener(this.tip, tipo)) this.cerrar(this.tip);
      const n = new Nodo(tipo);
      n.contenido = '';
      this.tip.agregar(n);
      this.tip = n;
      return n;
    }
    cerrarSinCoincidir() {
      if (!this.todoCerrado) {
        while (this.viejoTip !== this.ultimoCoincidente) {
          const padre = this.viejoTip.padre;
          this.cerrar(this.viejoTip);
          this.viejoTip = padre;
        }
        this.todoCerrado = true;
      }
    }
    cerrar(b) {
      const arriba = b.padre;
      b.abierto = false;
      this.alCerrar(b);
      this.tip = arriba;
    }

    alCerrar(b) {
      if (b.tipo === 'parrafo') {
        let tenia = false;
        let pos;
        while (b.contenido.charCodeAt(0) === 91 && (pos = this.leerReferencia(b.contenido))) { b.contenido = b.contenido.slice(pos); tenia = true; }
        if (tenia && /^[ \t\n]*$/.test(b.contenido)) b.quitar();
      } else if (b.tipo === 'codigo') {
        if (b.cercado) {
          const nl = b.contenido.indexOf('\n');
          b.info = unescapar(b.contenido.slice(0, nl).trim());
          b.literal = b.contenido.slice(nl + 1);
        } else {
          b.literal = b.contenido.replace(/(\n *)+$/, '\n');
        }
        b.contenido = '';
      } else if (b.tipo === 'html') {
        b.literal = b.contenido.replace(/(\n *)+$/, '');
        b.contenido = '';
      } else if (b.tipo === 'lista') {
        b.apretada = true;
        for (let item = b.primero; item; item = item.sig) {
          if (terminaVacia(item) && item.sig) { b.apretada = false; break; }
          for (let sub = item.primero; sub; sub = sub.sig) {
            if (terminaVacia(sub) && (item.sig || sub.sig)) { b.apretada = false; break; }
          }
          if (!b.apretada) break;
        }
      } else if (b.tipo === 'tabla') {
        b.contenido = '';
      }
    }

    // Continuación de un contenedor abierto: 0 = sigue, 1 = no coincide, 2 = la línea se consumió.
    continuar(c) {
      switch (c.tipo) {
        case 'cita':
          if (!this.sangrado && this.l[this.sigNoEsp] === '>') {
            this.avanzarNoEsp(); this.avanzar(1, false);
            if (esEspTab(this.l[this.offset])) this.avanzar(1, true);
            return 0;
          }
          return 1;
        case 'item':
          if (this.vacia) {
            if (c.primero === null) return 1;
            this.avanzarNoEsp();
          } else if (this.sangria >= c.datos.margen + c.datos.relleno) {
            this.avanzar(c.datos.margen + c.datos.relleno, true);
          } else return 1;
          return 0;
        case 'encabezado': case 'salto': return 1;
        case 'codigo':
          if (c.cercado) {
            const m = this.sangria <= 3 && this.l[this.sigNoEsp] === c.cerca && /^(?:`{3,}|~{3,})(?=[ \t]*$)/.exec(this.l.slice(this.sigNoEsp));
            if (m && m[0].length >= c.largoCerca) { this.cerrar(c); return 2; }
            let i = c.sangriaCerca;
            while (i > 0 && esEspTab(this.l[this.offset])) { this.avanzar(1, true); i--; }
          } else if (this.sangria >= 4) this.avanzar(4, true);
          else if (this.vacia) this.avanzarNoEsp();
          else return 1;
          return 0;
        case 'html': return this.vacia && (c.tipoHtml === 6 || c.tipoHtml === 7) ? 1 : 0;
        case 'parrafo': return this.vacia ? 1 : 0;
        case 'tabla': return this.vacia ? 1 : 0;
        default: return 0;
      }
    }

    linea(texto) {
      this.l = texto;
      this.offset = 0; this.column = 0; this.vacia = false; this.tabParcial = false;
      let cont = this.doc;
      this.viejoTip = this.tip;
      let coincide = true;
      let h;
      while ((h = cont.ultimo) && h.abierto) {
        cont = h;
        this.buscarNoEspacio();
        const r = this.continuar(cont);
        if (r === 2) return;
        if (r === 1) { coincide = false; }
        if (!coincide) { cont = cont.padre; break; }
      }
      this.todoCerrado = cont === this.viejoTip;
      this.ultimoCoincidente = cont;
      let hoja = cont.tipo !== 'parrafo' && this.aceptaLineas(cont.tipo);

      while (!hoja) {
        this.buscarNoEspacio();
        if (!this.sangrado && !RE_ESPECIAL.test(this.l.slice(this.sigNoEsp))) { this.avanzarNoEsp(); break; }
        let i = 0, res = 0;
        for (; i < 8; i++) {
          res = this.inicios(i, cont);
          if (res === 1) { cont = this.tip; break; }
          if (res === 2) { cont = this.tip; hoja = true; break; }
        }
        if (i === 8) { this.avanzarNoEsp(); break; }
      }

      if (!this.todoCerrado && !this.vacia && this.tip.tipo === 'tabla') { // filas de una tabla que quedó dentro de una lista o cita
        if (this.offset < this.l.length) this.tip.filas.push(this.l.slice(this.offset));
        return;
      }
      if (!this.todoCerrado && !this.vacia && this.tip.tipo === 'parrafo') {
        // Una tabla pegada a un párrafo de lista o de cita se queda dentro de él (como marked).
        if (this.sigNoEsp === this.offset && this.inicios(5, this.tip) === 2) { this.todoCerrado = true; return; }
        this.agregarLinea(); // continuación perezosa
        return;
      }
      this.cerrarSinCoincidir();
      if (this.vacia && cont.ultimo) cont.ultimo.ultimaLineaVacia = true;
      const t = cont.tipo;
      const lineaVacia = this.vacia && !(t === 'cita' || (t === 'codigo' && cont.cercado) || (t === 'item' && !cont.primero && cont.lineaInicio === this.numLinea));
      for (let n = cont; n; n = n.padre) n.ultimaLineaVacia = lineaVacia;
      if (t === 'tabla' && !this.vacia) { // otra fila de la tabla
        if (this.offset < this.l.length) cont.filas.push(this.l.slice(this.offset));
      } else if (this.aceptaLineas(t)) {
        this.agregarLinea();
        if (t === 'html' && cont.tipoHtml >= 1 && cont.tipoHtml <= 5 && HTML_BLOQUE_CIERRA[cont.tipoHtml].test(this.l.slice(this.offset))) this.cerrar(cont);
      } else if (this.offset < this.l.length && !this.vacia) {
        this.agregarHijo('parrafo');
        this.avanzarNoEsp();
        this.agregarLinea();
      }
    }

    inicios(i, cont) {
      const l = this.l;
      switch (i) {
        case 0: // cita
          if (!this.sangrado && l[this.sigNoEsp] === '>') {
            this.avanzarNoEsp(); this.avanzar(1, false);
            if (esEspTab(l[this.offset])) this.avanzar(1, true);
            this.cerrarSinCoincidir();
            this.agregarHijo('cita');
            return 1;
          }
          return 0;
        case 1: { // título ATX
          let m;
          if (!this.sangrado && (m = /^#{1,6}(?:[ \t]+|$)/.exec(l.slice(this.sigNoEsp)))) {
            this.avanzarNoEsp(); this.avanzar(m[0].length, false);
            this.cerrarSinCoincidir();
            const c = this.agregarHijo('encabezado');
            c.nivel = m[0].trim().length;
            c.contenido = l.slice(this.offset).replace(/^[ \t]*#+[ \t]*$/, '').replace(/[ \t]+#+[ \t]*$/, '');
            this.avanzar(l.length - this.offset, false);
            return 2;
          }
          return 0;
        }
        case 2: { // código cercado
          let m;
          if (!this.sangrado && (m = /^`{3,}(?!.*`)|^~{3,}/.exec(l.slice(this.sigNoEsp)))) {
            const largo = m[0].length;
            this.cerrarSinCoincidir();
            const c = this.agregarHijo('codigo');
            c.cercado = true; c.largoCerca = largo; c.cerca = m[0][0]; c.sangriaCerca = this.sangria;
            this.avanzarNoEsp(); this.avanzar(largo, false);
            return 2;
          }
          return 0;
        }
        case 3: // HTML
          if (!this.sangrado && l[this.sigNoEsp] === '<') {
            const s = l.slice(this.sigNoEsp);
            for (let bt = 1; bt <= 7; bt++) {
              if (HTML_BLOQUE_ABRE[bt].test(s) && (bt < 7 || (cont.tipo !== 'parrafo' && !(!this.todoCerrado && !this.vacia && this.tip.tipo === 'parrafo')))) {
                this.cerrarSinCoincidir();
                const b = this.agregarHijo('html');
                b.tipoHtml = bt;
                return 2;
              }
            }
          }
          return 0;
        case 4: { // título setext
          let m;
          if (!this.sangrado && cont.tipo === 'parrafo' && (m = RE_SETEXT.exec(l.slice(this.sigNoEsp)))) {
            this.cerrarSinCoincidir();
            let pos;
            while (cont.contenido.charCodeAt(0) === 91 && (pos = this.leerReferencia(cont.contenido))) cont.contenido = cont.contenido.slice(pos);
            if (cont.contenido.length > 0) {
              const h = new Nodo('encabezado');
              h.nivel = m[0][0] === '=' ? 1 : 2;
              h.contenido = cont.contenido;
              cont.insertarDespues(h);
              cont.quitar();
              this.tip = h;
              this.avanzar(l.length - this.offset, false);
              return 2;
            }
            return 0;
          }
          return 0;
        }
        case 5: { // tabla (GitHub): línea de guiones tras una cabecera
          if (cont.tipo === 'parrafo' && !this.sangrado && l.indexOf('-', this.sigNoEsp) >= 0 && RE_FILA_TABLA.test(l.slice(this.sigNoEsp)) && /[|]/.test(l.slice(this.sigNoEsp) + cont.contenido.slice(cont.contenido.lastIndexOf('\n', cont.contenido.length - 2) + 1))) {
            const guiones = dividirFila(l.slice(this.sigNoEsp));
            const lineasParrafo = cont.contenido.replace(/\n$/, '').split('\n');
            const cabecera = dividirFila(lineasParrafo[lineasParrafo.length - 1]);
            if (guiones.length === cabecera.length && guiones.every((g) => /^:?-+:?$/.test(g.trim()))) {
              this.cerrarSinCoincidir();
              lineasParrafo.pop();
              const t = new Nodo('tabla');
              t.contenido = '';
              t.cabecera = cabecera;
              t.alineaciones = guiones.map((g) => { g = g.trim(); const a = g[0] === ':', b = g[g.length - 1] === ':'; return a && b ? 'center' : a ? 'left' : b ? 'right' : null; });
              t.filas = [];
              if (lineasParrafo.length) {
                cont.contenido = lineasParrafo.join('\n') + '\n';
                cont.insertarDespues(t);
              } else {
                cont.insertarDespues(t);
                cont.quitar();
              }
              this.tip = t;
              this.avanzar(l.length - this.offset, false);
              return 2;
            }
          }
          return 0;
        }
        case 6: // línea horizontal
          if (!this.sangrado && RE_SALTO.test(l.slice(this.sigNoEsp))) {
            this.cerrarSinCoincidir();
            this.agregarHijo('salto');
            this.avanzar(l.length - this.offset, false);
            return 2;
          }
          return 0;
        case 7: { // ítem de lista o código con sangría
          let datos;
          if ((!this.sangrado || cont.tipo === 'lista') && (datos = this.marcadorLista(cont))) {
            this.cerrarSinCoincidir();
            if (this.tip.tipo !== 'lista' || !mismaLista(cont.datos, datos)) {
              const lista = this.agregarHijo('lista');
              lista.datos = datos;
            }
            const item = this.agregarHijo('item');
            item.datos = datos;
            item.lineaInicio = this.numLinea;
            return 1;
          }
          if (this.sangrado && this.tip.tipo !== 'parrafo' && !this.vacia) {
            this.avanzar(4, true);
            this.cerrarSinCoincidir();
            const c = this.agregarHijo('codigo');
            c.cercado = false;
            return 2;
          }
          return 0;
        }
        default: return 0;
      }
    }

    marcadorLista(cont) {
      const l = this.l;
      const resto = l.slice(this.sigNoEsp);
      const datos = { tipo: null, bala: null, inicio: null, delim: null, relleno: 0, margen: this.sangria };
      let m;
      if (this.sangria >= 4) return null;
      if ((m = /^[*+-]/.exec(resto))) { datos.tipo = 'bala'; datos.bala = m[0]; }
      else if ((m = /^(\d{1,9})([.)])/.exec(resto)) && (cont.tipo !== 'parrafo' || m[1] === '1')) { datos.tipo = 'ordenada'; datos.inicio = parseInt(m[1], 10); datos.delim = m[2]; }
      else return null;
      const sig = l[this.sigNoEsp + m[0].length];
      if (!(sig === undefined || sig === ' ' || sig === '\t')) return null;
      if (cont.tipo === 'parrafo' && !/[^ \t\n]/.test(l.slice(this.sigNoEsp + m[0].length))) return null;
      this.avanzarNoEsp();
      this.avanzar(m[0].length, true);
      const colEsp = this.column, offEsp = this.offset;
      do { this.avanzar(1, true); } while (this.column - colEsp < 5 && esEspTab(this.l[this.offset]));
      const itemVacio = this.l[this.offset] === undefined;
      const espacios = this.column - colEsp;
      if (espacios >= 5 || espacios < 1 || itemVacio) {
        datos.relleno = m[0].length + 1;
        this.column = colEsp; this.offset = offEsp;
        if (esEspTab(this.l[this.offset])) this.avanzar(1, true);
      } else datos.relleno = m[0].length + espacios;
      return datos;
    }

    // Definición de enlace `[ref]: destino "título"` al inicio de un párrafo. Devuelve cuánto texto consumió (0 si no hay).
    leerReferencia(s) {
      const a = new AnalizadorInline(this.refs);
      a.sujeto = s; a.pos = 0;
      const inicio = a.pos;
      const n = a.leerEtiqueta();
      if (n === 0) return 0;
      const etiqueta = a.sujeto.slice(0, n);
      if (a.sujeto[a.pos] !== ':') { return 0; }
      a.pos++;
      a.saltarEspaciosYLinea();
      const dest = a.leerDestino();
      if (dest === null || (dest.length === 0 && !a.destinoEntreAngulos)) return 0;
      const antesTitulo = a.pos;
      a.saltarEspaciosYLinea();
      let titulo = null;
      if (a.pos !== antesTitulo) titulo = a.leerTitulo();
      if (titulo === null) { titulo = ''; a.pos = antesTitulo; }
      let finLinea = true;
      if (a.coincide(/^[ \t]*(?:\n|$)/) === null) {
        if (titulo === '') finLinea = false;
        else { // el título no cierra la línea: se descarta y se vuelve a comprobar
          titulo = ''; a.pos = antesTitulo;
          finLinea = a.coincide(/^[ \t]*(?:\n|$)/) !== null;
        }
      }
      if (!finLinea) return 0;
      const clave = normalizarRef(etiqueta);
      if (clave === '') return 0;
      if (!this.refs[clave]) this.refs[clave] = { destino: dest, titulo };
      return a.pos - inicio;
    }
  }

  function terminaVacia(b) {
    while (b) {
      if (b.ultimaLineaVacia) return true;
      const t = b.tipo;
      if (!b.revisada && (t === 'lista' || t === 'item')) { b.revisada = true; b = b.ultimo; } else { b.revisada = true; break; }
    }
    return false;
  }
  const mismaLista = (a, b) => a && b && a.tipo === b.tipo && a.delim === b.delim && a.bala === b.bala;
  const normalizarRef = (s) => s.slice(1, -1).trim().replace(/[ \t\r\n]+/g, ' ').toLowerCase().toUpperCase();

  // Celdas de una fila de tabla: se parte por `|` sin escapar y se quitan los de los bordes.
  function dividirFila(linea) {
    let s = linea.trim();
    const celdas = [];
    let actual = '';
    for (let i = 0; i < s.length; i++) {
      if (s[i] === '\\' && s[i + 1] === '|') { actual += '\\|'; i++; }
      else if (s[i] === '|') { celdas.push(actual); actual = ''; }
      else actual += s[i];
    }
    celdas.push(actual);
    if (s[0] === '|') celdas.shift();
    if (celdas.length && s[s.length - 1] === '|' && s[s.length - 2] !== '\\') celdas.pop();
    return celdas;
  }

  function unescapar(s) {
    if (s.indexOf('\\') < 0) return s;
    return s.replace(/\\([!-/:-@[-`{-~])/g, '$1');
  }

  // ───────── Texto en línea ─────────
  class AnalizadorInline {
    constructor(refs) { this.refs = refs; }

    coincide(re) {
      const m = re.exec(this.sujeto.slice(this.pos));
      if (m === null) return null;
      this.pos += m[0].length;
      return m[0];
    }
    saltarEspaciosYLinea() {
      let m = this.coincide(/^[ \t]*(?:\n[ \t]*)?/);
      return m !== null;
    }
    leerEtiqueta() {
      const m = this.coincide(/^\[(?:[^\\\[\]]|\\.){0,1000}\]/s);
      if (m === null || m.length > 1001) return 0;
      return m.length;
    }
    leerDestino() {
      this.destinoEntreAngulos = false;
      let m = this.coincide(/^<(?:[^<>\n\\]|\\.)*>/);
      if (m !== null) { this.destinoEntreAngulos = true; return unescapar(m.slice(1, -1)); }
      const s = this.sujeto;
      let prof = 0, i = this.pos;
      for (; i < s.length; i++) {
        const c = s[i];
        if (c === '\\' && i + 1 < s.length && PUNTUACION_ASCII.indexOf(s[i + 1]) >= 0) { i++; continue; }
        if (c === '(') prof++;
        else if (c === ')') { if (prof === 0) break; prof--; }
        else if (c <= ' ' || c === '\x7f') break;
      }
      if (prof !== 0) return null;
      if (i === this.pos && s[i] !== ')') return null;
      const res = unescapar(s.slice(this.pos, i));
      this.pos = i;
      return res;
    }
    leerTitulo() {
      const m = this.coincide(/^(?:"(?:\\.|[^"\\\x00])*"|'(?:\\.|[^'\\\x00])*'|\((?:\\.|[^()\\\x00])*\))/s);
      return m === null ? null : unescapar(m.slice(1, -1));
    }

    analizar(padre, texto) {
      this.sujeto = texto;
      this.pos = 0;
      this.delimitadores = null;
      this.corchetes = null;
      while (this.pos < this.sujeto.length && this.uno(padre)) { /* sigue */ }
      this.procesarEnfasis(null);
      autoenlazar(padre);
    }

    uno(padre) {
      const s = this.sujeto;
      const c = s[this.pos];
      let res = false;
      switch (c) {
        case '\n': res = this.salto(padre); break;
        case '\\': res = this.barra(padre); break;
        case '`': res = this.acentos(padre); break;
        case '*': case '_': case '~': res = this.delimitadoresDe(c, padre); break;
        case '[': res = this.abreCorchete(padre); break;
        case '!': res = this.admiracion(padre); break;
        case ']': res = this.cierraCorchete(padre); break;
        case '<': res = this.autoenlace(padre) || this.htmlEnLinea(padre); break;
        case '&': res = this.entidad(padre); break;
        default: res = false;
      }
      if (!res) {
        // texto hasta el próximo carácter especial
        const ini = this.pos;
        this.pos++;
        while (this.pos < s.length && '\n\\`*_~[]!<&'.indexOf(s[this.pos]) < 0) this.pos++;
        padre.agregar(texto(s.slice(ini, this.pos)));
      }
      return true;
    }

    salto(padre) {
      this.pos++;
      const ult = padre.ultimo;
      if (ult && ult.tipo === 'texto' && ult.literal.endsWith('  ')) {
        ult.literal = ult.literal.replace(/ +$/, '');
        padre.agregar(new Nodo('quiebre'));
        this.coincide(/^[ \t]*/);
      } else padre.agregar(new Nodo('blando'));
      return true;
    }
    barra(padre) {
      this.pos++;
      const c = this.sujeto[this.pos];
      if (c === '\n') { this.pos++; padre.agregar(new Nodo('quiebre')); this.coincide(/^[ \t]*/); }
      else if (c !== undefined && PUNTUACION_ASCII.indexOf(c) >= 0) { this.pos++; padre.agregar(texto(c)); }
      else padre.agregar(texto('\\'));
      return true;
    }
    acentos(padre) {
      const m = /^`+/.exec(this.sujeto.slice(this.pos));
      const abre = m[0];
      const inicio = this.pos + abre.length;
      const re = /`+/g;
      re.lastIndex = inicio;
      let c;
      while ((c = re.exec(this.sujeto)) !== null) {
        if (c[0].length === abre.length) {
          let cont = this.sujeto.slice(inicio, c.index).replace(/\n/g, ' ');
          if (cont.length > 0 && /[^ ]/.test(cont) && cont[0] === ' ' && cont[cont.length - 1] === ' ') cont = cont.slice(1, -1);
          const n = new Nodo('codigoEnLinea');
          n.literal = cont;
          padre.agregar(n);
          this.pos = c.index + c[0].length;
          return true;
        }
      }
      this.pos += abre.length;
      padre.agregar(texto(abre));
      return true;
    }
    entidad(padre) {
      const m = /^&(?:#x[a-f0-9]{1,6}|#[0-9]{1,7}|[a-z][a-z0-9]{1,31});/i.exec(this.sujeto.slice(this.pos));
      if (!m) return false;
      this.pos += m[0].length;
      const n = new Nodo('crudo');
      n.literal = m[0];
      padre.agregar(n);
      return true;
    }
    autoenlace(padre) {
      let m;
      const resto = this.sujeto.slice(this.pos);
      if ((m = /^<([A-Za-z][A-Za-z0-9.+-]{1,31}:[^<>\x00-\x20]*)>/.exec(resto))) {
        const n = new Nodo('enlace');
        n.destino = m[1]; n.titulo = '';
        n.agregar(texto(m[1]));
        padre.agregar(n);
        this.pos += m[0].length;
        return true;
      }
      if ((m = /^<([a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*)>/.exec(resto))) {
        const n = new Nodo('enlace');
        n.destino = 'mailto:' + m[1]; n.titulo = '';
        n.agregar(texto(m[1]));
        padre.agregar(n);
        this.pos += m[0].length;
        return true;
      }
      return false;
    }
    htmlEnLinea(padre) {
      const m = HTML_EN_LINEA.exec(this.sujeto.slice(this.pos));
      if (!m) return false;
      this.pos += m[0].length;
      const n = new Nodo('htmlEnLinea');
      n.literal = m[0];
      padre.agregar(n);
      return true;
    }

    delimitadoresDe(c, padre) {
      const s = this.sujeto;
      const ini = this.pos;
      let cuenta = 0;
      while (s[this.pos] === c) { cuenta++; this.pos++; }
      const antes = ini === 0 ? '\n' : s.slice(Math.max(0, ini - 2), ini).replace(/^[\uD800-\uDBFF](?![\uDC00-\uDFFF])/, '');
      const antesCh = Array.from(antes).pop();
      const despues = this.pos >= s.length ? '\n' : String.fromCodePoint(s.codePointAt(this.pos));
      const antesEsp = RE_ESPACIO_UNICODE.test(antesCh), despuesEsp = RE_ESPACIO_UNICODE.test(despues);
      const antesPunt = RE_PUNTUACION.test(antesCh), despuesPunt = RE_PUNTUACION.test(despues);
      const izq = !despuesEsp && (!despuesPunt || antesEsp || antesPunt);
      const der = !antesEsp && (!antesPunt || despuesEsp || despuesPunt);
      let puedeAbrir, puedeCerrar;
      if (c === '_') { puedeAbrir = izq && (!der || antesPunt); puedeCerrar = der && (!izq || despuesPunt); }
      else { puedeAbrir = izq; puedeCerrar = der; }
      if (c === '~' && cuenta > 2) { padre.agregar(texto(s.slice(ini, this.pos))); return true; }
      const nodo = texto(s.slice(ini, this.pos));
      padre.agregar(nodo);
      if (puedeAbrir || puedeCerrar) {
        this.delimitadores = { cc: c, cuenta, original: cuenta, nodo, prev: this.delimitadores, sig: null, puedeAbrir, puedeCerrar };
        if (this.delimitadores.prev) this.delimitadores.prev.sig = this.delimitadores;
      }
      return true;
    }
    quitarDelimitador(d) {
      if (d.prev) d.prev.sig = d.sig;
      if (d.sig) d.sig.prev = d.prev; else this.delimitadores = d.prev;
    }
    quitarEntre(abre, cierra) {
      if (abre.sig !== cierra) { cierra.prev = abre; abre.sig = cierra; }
    }
    procesarEnfasis(fondo) {
      const topes = {};
      let cierra = this.delimitadores;
      while (cierra !== null && cierra.prev !== fondo) cierra = cierra.prev;
      while (cierra !== null) {
        if (!cierra.puedeCerrar) { cierra = cierra.sig; continue; }
        const clave = cierra.cc + (cierra.puedeAbrir ? 1 : 0) + (cierra.original % 3);
        let abre = cierra.prev;
        let hallado = false;
        const tope = topes[clave] === undefined ? fondo : topes[clave];
        while (abre !== null && abre !== fondo && abre !== tope) {
          if (cierra.cc === '~') {
            if (abre.cc === '~' && abre.puedeAbrir && abre.cuenta === cierra.cuenta) { hallado = true; break; }
          } else {
            const impar = (cierra.puedeAbrir || abre.puedeCerrar) && cierra.original % 3 !== 0 && (abre.original + cierra.original) % 3 === 0;
            if (abre.cc === cierra.cc && abre.puedeAbrir && !impar) { hallado = true; break; }
          }
          abre = abre.prev;
        }
        const viejo = cierra;
        if (hallado) {
          const usa = cierra.cc === '~' ? cierra.cuenta : (cierra.cuenta >= 2 && abre.cuenta >= 2 ? 2 : 1);
          const nAbre = abre.nodo, nCierra = cierra.nodo;
          abre.cuenta -= usa; cierra.cuenta -= usa;
          nAbre.literal = nAbre.literal.slice(0, nAbre.literal.length - usa);
          nCierra.literal = nCierra.literal.slice(0, nCierra.literal.length - usa);
          const enf = new Nodo(cierra.cc === '~' ? 'tachado' : usa === 1 ? 'enfasis' : 'fuerte');
          let tmp = nAbre.sig;
          while (tmp && tmp !== nCierra) { const sig = tmp.sig; tmp.quitar(); enf.agregar(tmp); tmp = sig; }
          nAbre.insertarDespues(enf);
          this.quitarEntre(abre, cierra);
          if (abre.cuenta === 0) { nAbre.quitar(); this.quitarDelimitador(abre); }
          if (cierra.cuenta === 0) { nCierra.quitar(); const sig = cierra.sig; this.quitarDelimitador(cierra); cierra = sig; }
        } else {
          cierra = cierra.sig;
          topes[clave] = viejo.prev;
          if (!viejo.puedeAbrir) this.quitarDelimitador(viejo);
        }
      }
      while (this.delimitadores !== null && this.delimitadores !== fondo) this.quitarDelimitador(this.delimitadores);
    }

    abreCorchete(padre) {
      const ini = this.pos;
      this.pos++;
      const n = texto('[');
      padre.agregar(n);
      this.agregarCorchete(n, ini, false);
      return true;
    }
    admiracion(padre) {
      const ini = this.pos;
      this.pos++;
      if (this.sujeto[this.pos] === '[') {
        this.pos++;
        const n = texto('![');
        padre.agregar(n);
        this.agregarCorchete(n, ini + 1, true);
      } else padre.agregar(texto('!'));
      return true;
    }
    agregarCorchete(nodo, indice, imagen) {
      if (this.corchetes !== null) this.corchetes.despues = true;
      this.corchetes = { nodo, prev: this.corchetes, delimPrev: this.delimitadores, indice, imagen, activo: true, despues: false };
    }
    cierraCorchete(padre) {
      this.pos++;
      const ini = this.pos;
      const abre = this.corchetes;
      if (abre === null) { padre.agregar(texto(']')); return true; }
      if (!abre.activo) { padre.agregar(texto(']')); this.corchetes = abre.prev; return true; }
      const esImagen = abre.imagen;
      const guardado = this.pos;
      let dest, titulo, hecho = false;
      if (this.sujeto[this.pos] === '(') {
        this.pos++;
        if (this.saltarEspaciosYLinea() && (dest = this.leerDestino()) !== null && this.saltarEspaciosYLinea()) {
          if (/\s/.test(this.sujeto[this.pos - 1] || '')) titulo = this.leerTitulo();
          this.saltarEspaciosYLinea();
          if (this.sujeto[this.pos] === ')') { this.pos++; hecho = true; } else this.pos = guardado;
        } else this.pos = guardado;
      }
      if (!hecho) {
        const antesEtiqueta = this.pos;
        const n = this.leerEtiqueta();
        let etiqueta;
        if (n > 2) etiqueta = this.sujeto.slice(antesEtiqueta, antesEtiqueta + n);
        else if (!abre.despues) etiqueta = this.sujeto.slice(abre.indice, ini);
        if (n === 0) this.pos = guardado;
        if (etiqueta) {
          const ref = this.refs[normalizarRef(etiqueta)];
          if (ref) { dest = ref.destino; titulo = ref.titulo; hecho = true; }
        }
      }
      if (hecho) {
        const nodo = new Nodo(esImagen ? 'imagen' : 'enlace');
        nodo.destino = dest; nodo.titulo = titulo || '';
        let tmp = abre.nodo.sig;
        while (tmp) { const sig = tmp.sig; tmp.quitar(); nodo.agregar(tmp); tmp = sig; }
        padre.agregar(nodo);
        this.procesarEnfasis(abre.delimPrev);
        this.corchetes = abre.prev;
        abre.nodo.quitar();
        if (!esImagen) for (let o = this.corchetes; o !== null; o = o.prev) if (!o.imagen) o.activo = false;
        return true;
      }
      this.corchetes = abre.prev;
      this.pos = ini;
      padre.agregar(texto(']'));
      return true;
    }
  }

  const texto = (s) => { const n = new Nodo('texto'); n.literal = s; return n; };

  // Enlaces automáticos al estilo de GitHub (https://…, www.…, correos), en los textos que no están dentro de un enlace.
  function autoenlazar(raiz) {
    const recorrer = (n, dentro) => {
      for (let h = n.primero; h; ) {
        while (h.tipo === 'texto' && h.sig && h.sig.tipo === 'texto') { h.literal += h.sig.literal; h.sig.quitar(); }
        const siguiente = h.sig; // trocear() puede reemplazar a h por varios nodos
        if (h.tipo === 'texto') { if (!dentro) trocear(h); }
        else if (h.primero) recorrer(h, dentro || h.tipo === 'enlace' || h.tipo === 'imagen');
        h = siguiente;
      }
    };
    recorrer(raiz, false);
  }
  const RE_URL = /(?:^|[^\p{L}\p{N}_])((?:(?:https?|ftp):\/\/|www\.)(?:[a-zA-Z0-9-]+\.?)+[^\s<]*)/gu;
  const RE_CORREO = /([A-Za-z0-9._+-]+@[a-zA-Z0-9_-]+(?:\.[a-zA-Z0-9_-]*[a-zA-Z0-9])+)(?![-_])/g;
  function recortarUrl(u) {
    for (;;) {
      const ant = u;
      u = u.replace(/[?!.,:;"'*_~]+$/, '');
      if (u.endsWith(')')) {
        const abre = (u.match(/\(/g) || []).length, cierra = (u.match(/\)/g) || []).length;
        if (cierra > abre) u = u.slice(0, -1);
      }
      if (/&[a-zA-Z0-9]+;$/.test(u)) u = u.replace(/&[a-zA-Z0-9]+;$/, '');
      if (u === ant) return u;
    }
  }
  function trocear(nodo) {
    const s = nodo.literal;
    if (s.indexOf('://') < 0 && s.indexOf('www.') < 0 && s.indexOf('@') < 0) return;
    const marcas = [];
    let m;
    RE_URL.lastIndex = 0;
    while ((m = RE_URL.exec(s)) !== null) {
      const inicio = m.index + m[0].length - m[1].length;
      const u = recortarUrl(m[1]);
      if (u.length > 4) marcas.push({ ini: inicio, fin: inicio + u.length, texto: u, destino: /^www\./i.test(u) ? 'http://' + u : u });
    }
    RE_CORREO.lastIndex = 0;
    while ((m = RE_CORREO.exec(s)) !== null) {
      const ini = m.index, fin = ini + m[0].length;
      if (marcas.some((x) => ini < x.fin && fin > x.ini)) continue;
      marcas.push({ ini, fin, texto: m[0], destino: 'mailto:' + m[0] });
    }
    if (!marcas.length) return;
    marcas.sort((a, b) => a.ini - b.ini);
    let pos = 0;
    let ancla = nodo;
    const agregar = (n) => { ancla.insertarDespues(n); ancla = n; };
    for (const k of marcas) {
      if (k.ini > pos) agregar(texto(s.slice(pos, k.ini)));
      const e = new Nodo('enlace');
      e.destino = k.destino; e.titulo = '';
      e.agregar(texto(k.texto));
      agregar(e);
      pos = k.fin;
    }
    if (pos < s.length) agregar(texto(s.slice(pos)));
    nodo.quitar();
  }

  // ───────── HTML ─────────
  const RE_ESC = /[&<>"']/g;
  const RE_ESC_SIN_ENTIDAD = /[<>"']|&(?!(?:#\d{1,7}|#[Xx][a-fA-F0-9]{1,6}|\w+);)/g;
  const MAPA_ESC = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
  const esc = (s, codificar) => s.replace(codificar ? RE_ESC : RE_ESC_SIN_ENTIDAD, (c) => MAPA_ESC[c]);
  function limpiarUrl(h) {
    try { return encodeURI(h).replace(/%25/g, '%'); } catch (e) { return null; }
  }

  function inlineAHtml(n) {
    let s = '';
    for (let h = n.primero; h; h = h.sig) {
      switch (h.tipo) {
        case 'texto': s += esc(h.literal, false); break;
        case 'crudo': s += h.literal; break;
        case 'htmlEnLinea': s += h.literal; break;
        case 'blando': s += '\n'; break;
        case 'quiebre': s += '<br>'; break;
        case 'codigoEnLinea': s += '<code>' + esc(h.literal, true) + '</code>'; break;
        case 'enfasis': s += '<em>' + inlineAHtml(h) + '</em>'; break;
        case 'fuerte': s += '<strong>' + inlineAHtml(h) + '</strong>'; break;
        case 'tachado': s += '<del>' + inlineAHtml(h) + '</del>'; break;
        case 'enlace': {
          const href = limpiarUrl(h.destino) === null ? null : esc(limpiarUrl(h.destino), false);
          const dentro = inlineAHtml(h);
          if (href === null) { s += dentro; break; }
          s += '<a href="' + href + '"' + (h.titulo ? ' title="' + esc(h.titulo, true) + '"' : '') + '>' + dentro + '</a>';
          break;
        }
        case 'imagen': {
          const href = limpiarUrl(h.destino) === null ? null : esc(limpiarUrl(h.destino), false);
          const alt = textoPlano(h);
          if (href === null) { s += esc(alt, true); break; }
          s += '<img src="' + href + '" alt="' + esc(alt, true) + '"' + (h.titulo ? ' title="' + esc(h.titulo, true) + '"' : '') + '>';
          break;
        }
        default: break;
      }
    }
    return s;
  }
  function textoPlano(n) {
    let s = '';
    for (let h = n.primero; h; h = h.sig) {
      if (h.tipo === 'texto' || h.tipo === 'codigoEnLinea' || h.tipo === 'crudo') s += h.literal;
      else if (h.tipo === 'blando' || h.tipo === 'quiebre') s += '\n';
      else if (h.primero) s += textoPlano(h);
    }
    return s;
  }

  function parsearInline(bloque, refs) {
    const raiz = new Nodo('raiz');
    new AnalizadorInline(refs).analizar(raiz, bloque.tipo === 'parrafo' ? bloque.contenido.replace(/\n+$/, '') : bloque.contenido.replace(/^[ \t\n]+|[ \t\n]+$/g, ''));
    return raiz;
  }

  function bloqueAHtml(n, refs, apretada) {
    switch (n.tipo) {
      case 'parrafo': {
        const dentro = inlineAHtml(parsearInline(n, refs));
        return apretada ? dentro : '<p>' + dentro + '</p>\n';
      }
      case 'encabezado': return '<h' + n.nivel + '>' + inlineAHtml(parsearInline(n, refs)) + '</h' + n.nivel + '>\n';
      case 'salto': return '<hr>\n';
      case 'cita': return '<blockquote>\n' + hijosAHtml(n, refs, false) + '</blockquote>\n';
      case 'codigo': {
        const lang = (n.info || '').split(/\s+/)[0];
        return '<pre><code' + (lang ? ' class="language-' + esc(lang, true) + '"' : '') + '>' + esc(n.literal, true) + '</code></pre>\n';
      }
      case 'html': return n.literal;
      case 'lista': {
        const ord = n.datos.tipo === 'ordenada';
        const etiqueta = ord ? 'ol' : 'ul';
        const inicio = ord && n.datos.inicio !== 1 ? ' start="' + n.datos.inicio + '"' : '';
        let s = '<' + etiqueta + inicio + '>\n';
        for (let i = n.primero; i; i = i.sig) s += itemAHtml(i, refs, n.apretada);
        return s + '</' + etiqueta + '>\n';
      }
      case 'tabla': return tablaAHtml(n, refs);
      default: return '';
    }
  }
  function hijosAHtml(n, refs, apretada) {
    let s = '';
    for (let h = n.primero; h; h = h.sig) s += bloqueAHtml(h, refs, apretada);
    return s;
  }
  function itemAHtml(item, refs, apretada) {
    // Lista de tareas: `[ ] ` o `[x] ` al inicio del primer párrafo.
    let casilla = '';
    const primero = item.primero;
    if (primero && primero.tipo === 'parrafo') {
      const m = /^\[([ xX])\][ \t]+(?=\S)/.exec(primero.contenido);
      if (m) {
        primero.contenido = primero.contenido.slice(m[0].length);
        casilla = '<input ' + (m[1] === ' ' ? '' : 'checked="" ') + 'disabled="" type="checkbox"> ';
      }
    }
    let s = '';
    for (let h = item.primero; h; h = h.sig) {
      if (h.tipo === 'parrafo') {
        const dentro = (h === primero ? casilla : '') + inlineAHtml(parsearInline(h, refs));
        s += apretada ? dentro : '<p>' + dentro + '</p>\n';
      } else s += bloqueAHtml(h, refs, apretada);
    }
    return '<li>' + s + '</li>\n';
  }
  function tablaAHtml(t, refs) {
    const celda = (txt, etiqueta, i) => {
      const raiz = new Nodo('raiz');
      new AnalizadorInline(refs).analizar(raiz, txt.trim().replace(/\\\|/g, '|'));
      const al = t.alineaciones[i];
      return '<' + etiqueta + (al ? ' align="' + al + '"' : '') + '>' + inlineAHtml(raiz) + '</' + etiqueta + '>\n';
    };
    const fila = (celdas, etiqueta) => {
      let s = '<tr>\n';
      for (let i = 0; i < t.alineaciones.length; i++) s += celda(celdas[i] || '', etiqueta, i);
      return s + '</tr>\n';
    };
    let s = '<table>\n<thead>\n' + fila(t.cabecera, 'th') + '</thead>\n';
    if (t.filas.length) {
      s += '<tbody>';
      for (const f of t.filas) s += fila(dividirFila(f), 'td');
      s += '</tbody>';
    }
    return s + '</table>\n';
  }

  function aHtml(textoMd) {
    const entrada = String(textoMd == null ? '' : textoMd);
    const a = new AnalizadorBloques();
    const doc = a.analizar(entrada);
    // marked conserva el salto de línea final cuando el documento termina en un bloque HTML sin línea en blanco detrás.
    if (doc.ultimo && doc.ultimo.tipo === 'html' && /[^\n]\r?\n$/.test(entrada)) doc.ultimo.literal += '\n';
    return hijosAHtml(doc, a.refs, false);
  }

  globalThis.MarkdownPropio = { aHtml };
})();

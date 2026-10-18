/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Un "lienzo" que se usa como el contexto 2D de un <canvas> pero, en vez de
// pintar píxeles, escribe los operadores de una página de PDF: trazados,
// rellenos, líneas, recortes, degradados (Shading), imágenes y texto con la
// fuente incrustada. Etapa 4b del motor de PowerPoint (pptx-dibujo.js y
// pptx-grafico.js dibujan igual sobre un canvas que sobre este lienzo).
//
// Lo que un PDF no puede expresar —sombras con desenfoque, tramas y mosaicos
// (createPattern), degradados con transparencia, texto con una fuente de la que
// no se tienen los bytes o con letras que le faltan— se pinta, solo ese trazo,
// en un canvas del tamaño justo y va a la página como imagen con transparencia.
// Si algo falla de verdad, `fallo` queda con el error y quien lo usa dibuja esa
// diapositiva como imagen entera.
//
// Uso:
//   const lienzo = new LienzoPdf.Lienzo({ ancho, alto, anchoPt, altoPt });
//   await LienzoPdf.precargarFuentes(usos);          // [{ fuente, negrita, cursiva }]
//   await PptxDibujo.dibujar(pres, diapositiva, lienzo);  // el lienzo hace de canvas
//   if (lienzo.fallo) …; await lienzo.aPagina(pdfDoc); // agrega la página al PDFDocument
// Requiere pdf-lib, assets/js/vendor.js y pptx-texto.js (PptxTexto.archivoDeFuente).

const LienzoPdf = (() => {
  const Path2DReal = Path2D;
  const URL_FONTKIT = rutaVendor('@pdf-lib/fontkit/1.1.1/dist/fontkit.umd.min.js?v=2');
  const TAU = Math.PI * 2;

  // --- Números y matrices ([a, b, c, d, e, f]: x' = a·x + c·y + e) --------------------------
  const num = (x) => {
    if (!isFinite(x)) return '0';
    const s = x.toFixed(4);
    if (s.indexOf('.') < 0) return s;
    const r = s.replace(/0+$/, '').replace(/\.$/, '');
    return r === '-0' || r === '' ? '0' : r;
  };
  // A∘B: primero B, después A (igual que transform() de un canvas y que cm de PDF).
  const mul = (A, B) => [
    A[0] * B[0] + A[2] * B[1], A[1] * B[0] + A[3] * B[1],
    A[0] * B[2] + A[2] * B[3], A[1] * B[2] + A[3] * B[3],
    A[0] * B[4] + A[2] * B[5] + A[4], A[1] * B[4] + A[3] * B[5] + A[5],
  ];
  function inv(M) {
    const det = M[0] * M[3] - M[1] * M[2];
    if (!det || !isFinite(det)) return null;
    const a = M[3] / det, b = -M[1] / det, c = -M[2] / det, d = M[0] / det;
    return [a, b, c, d, -(a * M[4] + c * M[5]), -(b * M[4] + d * M[5])];
  }
  const aplicar = (M, x, y) => [M[0] * x + M[2] * y + M[4], M[1] * x + M[3] * y + M[5]];
  const cm = (M) => `${M.map(num).join(' ')} cm`;
  const desdeMatriz = (m) => (Array.isArray(m) ? m : [m.a, m.b, m.c, m.d, m.e, m.f]);

  // --- Colores -----------------------------------------------------------------------------
  // El navegador normaliza cualquier color CSS a #rrggbb o rgba(r, g, b, a).
  const medidor = document.createElement('canvas').getContext('2d');
  const colores = new Map();
  function color(texto) {
    let c = colores.get(texto);
    if (c) return c;
    medidor.fillStyle = '#000000';
    medidor.fillStyle = texto;
    const v = medidor.fillStyle;
    let m;
    if ((m = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(v))) c = { r: parseInt(m[1], 16), g: parseInt(m[2], 16), b: parseInt(m[3], 16), a: 1 };
    else if ((m = /^rgba?\(([^)]+)\)/.exec(v))) { const p = m[1].split(',').map(Number); c = { r: p[0], g: p[1], b: p[2], a: p.length > 3 ? p[3] : 1 }; }
    else c = { r: 0, g: 0, b: 0, a: 1 };
    colores.set(texto, c);
    return c;
  }
  const esColor = (v) => typeof v === 'string';

  // --- Trazados -----------------------------------------------------------------------------
  // Hace de Path2D (pptx-dibujo.js lo toma de global mientras dibuja sobre un
  // lienzo): guarda las órdenes para escribirlas como operadores de PDF.
  class Ruta {
    constructor(otra) {
      this.cmds = otra && otra.cmds ? otra.cmds.map((c) => c.slice()) : [];
      this.ini = null; this.act = null;
    }
    moveTo(x, y) { if (!isFinite(x + y)) return; this.cmds.push(['M', x, y]); this.ini = this.act = [x, y]; }
    lineTo(x, y) { if (!isFinite(x + y)) return; if (!this.act) { this.moveTo(x, y); return; } this.cmds.push(['L', x, y]); this.act = [x, y]; }
    bezierCurveTo(a, b, c, d, x, y) { if (!isFinite(a + b + c + d + x + y)) return; if (!this.act) this.moveTo(a, b); this.cmds.push(['C', a, b, c, d, x, y]); this.act = [x, y]; }
    quadraticCurveTo(cx, cy, x, y) {
      if (!this.act) this.moveTo(cx, cy);
      const [x0, y0] = this.act;
      this.bezierCurveTo(x0 + 2 / 3 * (cx - x0), y0 + 2 / 3 * (cy - y0), x + 2 / 3 * (cx - x), y + 2 / 3 * (cy - y), x, y);
    }
    closePath() { if (!this.act) return; this.cmds.push(['Z']); this.act = this.ini; }
    rect(x, y, w, h) { this.moveTo(x, y); this.lineTo(x + w, y); this.lineTo(x + w, y + h); this.lineTo(x, y + h); this.closePath(); }
    ellipse(cx, cy, rx, ry, rot, ini, fin, antihorario) {
      let barrido = fin - ini;
      if (!antihorario) { if (barrido >= TAU || barrido <= -TAU) barrido = TAU; else if (barrido < 0) barrido = ((barrido % TAU) + TAU) % TAU; }
      else if (barrido <= -TAU) barrido = -TAU; else if (barrido > 0) barrido = (barrido % TAU) - TAU;
      const cr = Math.cos(rot), sr = Math.sin(rot);
      const punto = (t) => [cx + rx * Math.cos(t) * cr - ry * Math.sin(t) * sr, cy + rx * Math.cos(t) * sr + ry * Math.sin(t) * cr];
      const deriv = (t) => [-rx * Math.sin(t) * cr - ry * Math.cos(t) * sr, -rx * Math.sin(t) * sr + ry * Math.cos(t) * cr];
      const p0 = punto(ini);
      if (this.act) this.lineTo(p0[0], p0[1]); else this.moveTo(p0[0], p0[1]);
      const n = Math.max(1, Math.ceil(Math.abs(barrido) / (Math.PI / 2) - 1e-9));
      const paso = barrido / n, k = 4 / 3 * Math.tan(paso / 4);
      for (let i = 0; i < n; i++) {
        const t0 = ini + paso * i, t1 = t0 + paso;
        const a = punto(t0), b = punto(t1), da = deriv(t0), db = deriv(t1);
        this.bezierCurveTo(a[0] + k * da[0], a[1] + k * da[1], b[0] - k * db[0], b[1] - k * db[1], b[0], b[1]);
      }
    }
    arc(x, y, r, ini, fin, antihorario) { this.ellipse(x, y, r, r, 0, ini, fin, antihorario); }
    addPath(otra, m) {
      const M = m ? desdeMatriz(m) : null;
      const p = (x, y) => (M ? aplicar(M, x, y) : [x, y]);
      for (const c of otra.cmds) {
        if (c[0] === 'M') { const q = p(c[1], c[2]); this.moveTo(q[0], q[1]); }
        else if (c[0] === 'L') { const q = p(c[1], c[2]); this.lineTo(q[0], q[1]); }
        else if (c[0] === 'C') { const a = p(c[1], c[2]), b = p(c[3], c[4]), q = p(c[5], c[6]); this.bezierCurveTo(a[0], a[1], b[0], b[1], q[0], q[1]); }
        else this.closePath();
      }
    }
    // Para el respaldo por imagen: la misma ruta como Path2D del navegador.
    aReal() {
      const p = new Path2DReal();
      for (const c of this.cmds) {
        if (c[0] === 'M') p.moveTo(c[1], c[2]); else if (c[0] === 'L') p.lineTo(c[1], c[2]);
        else if (c[0] === 'C') p.bezierCurveTo(c[1], c[2], c[3], c[4], c[5], c[6]); else p.closePath();
      }
      return p;
    }
    caja() {
      let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
      for (const c of this.cmds) for (let i = 1; i < c.length; i += 2) { x0 = Math.min(x0, c[i]); x1 = Math.max(x1, c[i]); y0 = Math.min(y0, c[i + 1]); y1 = Math.max(y1, c[i + 1]); }
      return x0 === Infinity ? null : { x0, y0, x1, y1 };
    }
    operadores() {
      return this.cmds.map((c) => (c[0] === 'M' ? `${num(c[1])} ${num(c[2])} m` : c[0] === 'L' ? `${num(c[1])} ${num(c[2])} l`
        : c[0] === 'C' ? `${c.slice(1).map(num).join(' ')} c` : 'h'));
    }
  }

  class Degradado {
    constructor(tipo, coords) { this.tipo = tipo; this.coords = coords; this.paradas = []; }
    addColorStop(pos, c) { this.paradas.push({ pos: Math.min(1, Math.max(0, pos)), color: color(c) }); }
    translucido() { return this.paradas.some((p) => p.color.a < 0.999); }
    aReal(ctx) {
      const g = this.tipo === 'lineal' ? ctx.createLinearGradient(...this.coords) : ctx.createRadialGradient(...this.coords);
      for (const p of this.paradas) g.addColorStop(p.pos, `rgba(${p.color.r},${p.color.g},${p.color.b},${p.color.a})`);
      return g;
    }
  }
  class Patron {
    constructor(fuente, repeticion) { this.fuente = fuente; this.repeticion = repeticion; this.matriz = null; }
    setTransform(m) { this.matriz = m; }
    aReal(ctx) {
      const p = ctx.createPattern(this.fuente, this.repeticion || 'repeat');
      if (this.matriz) p.setTransform(this.matriz);
      return p;
    }
  }

  // --- Fuentes -------------------------------------------------------------------------------
  // Los bytes y el fontkit de cada fuente que el motor sabe (PptxTexto.archivoDeFuente),
  // cargados antes de dibujar para poder decidir, al dibujar, si una letra existe
  // en la fuente que va al PDF.
  const fuentesCargadas = new Map();
  let fontkitListo = null;
  function cargarFontkit() {
    if (window.fontkit) return Promise.resolve();
    return (fontkitListo ||= new Promise((listo, falla) => {
      const s = document.createElement('script');
      s.src = URL_FONTKIT; s.onload = listo;
      s.onerror = () => { fontkitListo = null; falla(new Error('No se pudo cargar fontkit.')); };
      document.head.appendChild(s);
    }));
  }
  // fontkit arma el subconjunto de una fuente TrueType con la tabla loca corta
  // (posiciones ÷ 2) pero deja cada glifo con la longitud que traía; Carlito
  // los tiene de longitud impar y el subconjunto sale corrido: casi sin
  // letras en el PDF. Se rellena cada glifo hasta una longitud par.
  function fontkitSeguro() {
    const fk = window.fontkit;
    return {
      create: (...args) => {
        const fuente = fk.create(...args);
        const crear = fuente.createSubset;
        if (crear) {
          fuente.createSubset = function (...a) {
            const sub = crear.apply(this, a);
            const agregar = sub._addGlyph;
            if (agregar) {
              sub._addGlyph = function (...b) {
                const r = agregar.apply(this, b);
                const g = this.glyf[this.glyf.length - 1];
                if (g && g.length % 2) { const B = g.constructor; this.glyf[this.glyf.length - 1] = B.concat([g, B.alloc(1)]); this.offset += 1; }
                return r;
              };
            }
            return sub;
          };
        }
        return fuente;
      },
    };
  }
  async function precargarFuentes(usos) {
    if (typeof PptxTexto === 'undefined' || !PptxTexto.archivoDeFuente) return;
    const urls = new Set();
    for (const u of usos || []) { const a = PptxTexto.archivoDeFuente(u, true); if (a) urls.add(a.url); }
    if (!urls.size) return;
    await cargarFontkit();
    await Promise.all([...urls].filter((u) => !fuentesCargadas.has(u)).map(async (url) => {
      const promesa = (async () => {
        const r = await fetch(url);
        if (!r.ok) throw new Error('No se pudo cargar ' + url);
        const bytes = new Uint8Array(await r.arrayBuffer());
        return { bytes, fk: window.fontkit.create(bytes) };
      })();
      fuentesCargadas.set(url, promesa);
      try { await promesa; } catch (e) { fuentesCargadas.delete(url); console.warn('Fuente sin cargar para el PDF:', e); }
    }));
  }
  // Las ya cargadas: lo que el dibujo necesita es síncrono.
  const fuenteLista = new Map();
  async function sincronizar() { for (const [u, p] of fuentesCargadas) if (!fuenteLista.has(u)) { try { fuenteLista.set(u, await p); } catch (e) { /* sin ella */ } } }

  // --- El lienzo -------------------------------------------------------------------------------
  const PROPIEDADES = ['fillStyle', 'strokeStyle', 'lineWidth', 'lineCap', 'lineJoin', 'miterLimit', 'globalAlpha', 'textBaseline', 'textAlign',
    'shadowColor', 'shadowBlur', 'shadowOffsetX', 'shadowOffsetY', 'imageSmoothingEnabled', 'imageSmoothingQuality', 'globalCompositeOperation'];
  const nuevoEstado = () => ({
    m: [1, 0, 0, 1, 0, 0], fillStyle: '#000000', strokeStyle: '#000000', lineWidth: 1, lineCap: 'butt', lineJoin: 'miter', miterLimit: 10,
    dash: [], globalAlpha: 1, font: '10px sans-serif', letterSpacing: '0px', fontKerning: 'auto', textBaseline: 'alphabetic', textAlign: 'start',
    shadowColor: 'rgba(0, 0, 0, 0)', shadowBlur: 0, shadowOffsetX: 0, shadowOffsetY: 0, imageSmoothingEnabled: true, imageSmoothingQuality: 'low',
    globalCompositeOperation: 'source-over',
    clips: [],
    pdf: { fill: null, stroke: null, lw: null, cap: null, join: null, miter: null, dash: null, alfa: 1 },
  });
  const copiarEstado = (e) => ({ ...e, m: e.m.slice(), dash: e.dash.slice(), clips: e.clips, pdf: { ...e.pdf } });
  // Tope de la imagen de respaldo (lado mayor, en píxeles).
  const LADO_MAXIMO_RESPALDO = 4096;

  class Lienzo {
    constructor({ ancho, alto, anchoPt, altoPt }) {
      this.width = ancho; this.height = alto;
      this.anchoPt = anchoPt; this.altoPt = altoPt;
      this.sal = [];
      this.e = nuevoEstado();
      this.pila = [];
      this.ruta = new Ruta();
      this.fallo = null;
      // Recursos que usa la página; los nombres se resuelven al escribirla.
      this.imagenes = []; this.indiceImagen = new Map();
      this.fuentes = []; this.indiceFuente = new Map();
      this.degradados = []; this.alfas = []; this.indiceAlfa = new Map();
      this.patrones = []; this.mascaras = [];
      this.respaldos = 0;
      this.motivos = [];
    }
    getContext() { return this; }
    get precargarFuentes() { return async (usos) => { await precargarFuentes(usos); await sincronizar(); }; }
    static get Ruta() { return Ruta; }

    get font() { return this.e.font; }
    set font(v) { this.e.font = v; medidor.font = v; }
    get letterSpacing() { return this.e.letterSpacing; }
    set letterSpacing(v) { this.e.letterSpacing = v; medidor.letterSpacing = v; }
    get fontKerning() { return this.e.fontKerning; }
    set fontKerning(v) { this.e.fontKerning = v; medidor.fontKerning = v; }
    measureText(t) { medidor.font = this.e.font; medidor.letterSpacing = this.e.letterSpacing; medidor.fontKerning = this.e.fontKerning; return medidor.measureText(t); }

    // --- estado y transformación
    save() { this.pila.push(copiarEstado(this.e)); this.sal.push('q'); }
    restore() { if (!this.pila.length) return; this.e = this.pila.pop(); this.sal.push('Q'); }
    transform(a, b, c, d, e, f) { const M = [a, b, c, d, e, f]; this.e.m = mul(this.e.m, M); this.sal.push(cm(M)); }
    translate(x, y) { this.transform(1, 0, 0, 1, x, y); }
    scale(x, y) { this.transform(x, 0, 0, y, 0, 0); }
    rotate(a) { const c = Math.cos(a), s = Math.sin(a); this.transform(c, s, -s, c, 0, 0); }
    setTransform(a, b, c, d, e, f) {
      const M = typeof a === 'object' ? desdeMatriz(a) : [a, b, c, d, e, f];
      const i = inv(this.e.m);
      if (!i) throw new Error('Transformación no invertible');
      this.sal.push(cm(mul(i, M)));
      this.e.m = M;
    }
    resetTransform() { this.setTransform(1, 0, 0, 1, 0, 0); }
    getTransform() { const m = this.e.m; return new DOMMatrix([m[0], m[1], m[2], m[3], m[4], m[5]]); }
    setLineDash(d) { this.e.dash = d.slice(); }
    getLineDash() { return this.e.dash.slice(); }

    // --- trazado actual
    beginPath() { this.ruta = new Ruta(); }
    moveTo(x, y) { this.ruta.moveTo(x, y); }
    lineTo(x, y) { this.ruta.lineTo(x, y); }
    bezierCurveTo(a, b, c, d, x, y) { this.ruta.bezierCurveTo(a, b, c, d, x, y); }
    quadraticCurveTo(a, b, x, y) { this.ruta.quadraticCurveTo(a, b, x, y); }
    closePath() { this.ruta.closePath(); }
    rect(x, y, w, h) { this.ruta.rect(x, y, w, h); }
    ellipse(...a) { this.ruta.ellipse(...a); }
    arc(...a) { this.ruta.arc(...a); }

    createLinearGradient(x0, y0, x1, y1) { if (![x0, y0, x1, y1].every(isFinite)) throw new TypeError('Degradado con valores no finitos'); return new Degradado('lineal', [x0, y0, x1, y1]); }
    createRadialGradient(x0, y0, r0, x1, y1, r1) { if (![x0, y0, r0, x1, y1, r1].every(isFinite)) throw new TypeError('Degradado con valores no finitos'); return new Degradado('radial', [x0, y0, r0, x1, y1, r1]); }
    createPattern(fuente, repeticion) { return new Patron(fuente, repeticion); }

    // --- pintar
    _args(path, regla) {
      if (typeof path === 'string') return { ruta: this.ruta, regla: path };
      return { ruta: path || this.ruta, regla: regla || 'nonzero' };
    }
    _comprobar(ruta) { if (!(ruta instanceof Ruta)) throw new Error('El lienzo de PDF solo acepta trazados propios (Path2D no está reemplazado).'); }
    _conSombra() {
      const e = this.e;
      if (!(e.shadowBlur > 0 || e.shadowOffsetX || e.shadowOffsetY)) return false;
      return color(e.shadowColor).a > 0;
    }
    _fallar(error) { if (!this.fallo) this.fallo = error; throw error; }

    fill(path, regla) {
      try {
        const a = this._args(path, regla); this._comprobar(a.ruta);
        if (!a.ruta.cmds.length) return;
        this._rellenar(a.ruta, a.regla);
      } catch (error) { this._fallar(error); }
    }
    fillRect(x, y, w, h) {
      try { const r = new Ruta(); r.rect(x, y, w, h); this._rellenar(r, 'nonzero'); } catch (error) { this._fallar(error); }
    }
    clearRect() { /* el lienzo de PDF empieza vacío */ }
    stroke(path) {
      try {
        const r = path || this.ruta; this._comprobar(r);
        if (!r.cmds.length || !(this.e.lineWidth > 0)) return;
        this._trazar(r);
      } catch (error) { this._fallar(error); }
    }
    strokeRect(x, y, w, h) {
      try { const r = new Ruta(); r.rect(x, y, w, h); this._trazar(r); } catch (error) { this._fallar(error); }
    }
    clip(path, regla) {
      try {
        const a = this._args(path, regla); this._comprobar(a.ruta);
        const copia = new Ruta(a.ruta);
        // Un recorte sin trazado deja todo afuera (no "sin recorte").
        this.sal.push(...(a.ruta.cmds.length ? a.ruta.operadores() : ['0 0 0 0 re']), a.regla === 'evenodd' ? 'W* n' : 'W n');
        this.e.clips = [...this.e.clips, { ruta: copia, m: this.e.m.slice(), regla: a.regla }];
      } catch (error) { this._fallar(error); }
    }

    _rellenar(ruta, regla) {
      const rell = this.e.fillStyle;
      if (this._conSombra()) {
        this._sombra(this._cajaRuta(ruta, 0), (c) => { c.fillStyle = this._real(rell, c); c.fill(ruta.aReal(), regla); });
        this._sinSombra(() => this._rellenar(ruta, regla));
        return;
      }
      if (rell instanceof Degradado && rell.paradas.length && rell.translucido()) {
        // Colores por un lado y la transparencia por otro (máscara de luminosidad).
        const d = this._registrarDegradado(rell), i = inv(this.e.m);
        if (i) {
          const pts = [aplicar(i, 0, 0), aplicar(i, this.width, 0), aplicar(i, 0, this.height), aplicar(i, this.width, this.height)];
          const caja = [Math.min(...pts.map((p) => p[0])), Math.min(...pts.map((p) => p[1])), Math.max(...pts.map((p) => p[0])), Math.max(...pts.map((p) => p[1]))];
          this.mascaras.push({ degradado: rell, caja });
          this._fijarAlfa(this.e.globalAlpha);
          this.sal.push('q', ...ruta.operadores(), regla === 'evenodd' ? 'W* n' : 'W n', `/M${this.mascaras.length - 1} gs`, `/S${d} sh`, 'Q');
          return;
        }
      }
      if ((rell instanceof Patron && !this._patronVectorial(rell)) || (rell instanceof Degradado && !rell.paradas.length)) {
        this._respaldo(this._cajaRuta(ruta, 0), (c) => { c.fillStyle = this._real(rell, c); c.fill(ruta.aReal(), regla); },
          false, rell instanceof Degradado ? 'degradado con transparencia' : 'trama o mosaico');
        return;
      }
      if (rell instanceof Patron && this._patronVectorial(rell)) {
        const fuente = rell.fuente, ancho = fuente.naturalWidth || fuente.width, alto = fuente.naturalHeight || fuente.height;
        const M = mul(this.e.m, rell.matriz ? desdeMatriz(rell.matriz) : [1, 0, 0, 1, 0, 0]);
        const f = this.anchoPt / this.width;
        const i = this.patrones.length;
        this.patrones.push({ imagen: this._registrarImagen(fuente, ancho, alto), ancho, alto, matriz: mul([f, 0, 0, -f, 0, this.altoPt], M) });
        this._fijarAlfa(this.e.globalAlpha);
        this.sal.push('/Pattern cs', `/P${i} scn`, ...ruta.operadores(), regla === 'evenodd' ? 'f*' : 'f');
        this.e.pdf.fill = null;
        return;
      }
      if (rell instanceof Degradado) {
        const d = this._registrarDegradado(rell);
        this._fijarAlfa(this.e.globalAlpha);
        this.sal.push('q', ...ruta.operadores(), regla === 'evenodd' ? 'W* n' : 'W n', `/S${d} sh`, 'Q');
        return;
      }
      const c = color(rell);
      this._fijarRelleno(c);
      this.sal.push(...ruta.operadores(), regla === 'evenodd' ? 'f*' : 'f');
    }

    // Un patrón de canvas que se repite en las dos direcciones es un patrón de
    // mosaico de PDF: la imagen una sola vez, en vez de una imagen del tamaño de
    // toda el área.
    _patronVectorial(p) {
      const f = p.fuente;
      const ancho = f && (f.naturalWidth || f.width), alto = f && (f.naturalHeight || f.height);
      return (p.repeticion || 'repeat') === 'repeat' && ancho > 0 && alto > 0;
    }

    _trazar(ruta) {
      const trazo = this.e.strokeStyle;
      if (this._conSombra()) {
        this._sombra(this._cajaRuta(ruta, this.e.lineWidth * 1.5 + 2), (c) => { this._copiarLinea(c); c.strokeStyle = this._real(trazo, c); c.stroke(ruta.aReal()); });
        this._sinSombra(() => this._trazar(ruta));
        return;
      }
      if (!esColor(trazo)) {
        this._respaldo(this._cajaRuta(ruta, this.e.lineWidth * 1.5 + 2), (c) => {
          this._copiarLinea(c); c.strokeStyle = this._real(trazo, c); c.stroke(ruta.aReal());
        });
        return;
      }
      const c = color(trazo), p = this.e.pdf, e = this.e;
      const clave = `${c.r} ${c.g} ${c.b}`;
      if (p.stroke !== clave) { this.sal.push(`${num(c.r / 255)} ${num(c.g / 255)} ${num(c.b / 255)} RG`); p.stroke = clave; }
      this._fijarAlfa(c.a * e.globalAlpha);
      if (p.lw !== e.lineWidth) { this.sal.push(`${num(e.lineWidth)} w`); p.lw = e.lineWidth; }
      const cap = e.lineCap === 'round' ? 1 : e.lineCap === 'square' ? 2 : 0;
      if (p.cap !== cap) { this.sal.push(`${cap} J`); p.cap = cap; }
      const union = e.lineJoin === 'round' ? 1 : e.lineJoin === 'bevel' ? 2 : 0;
      if (p.join !== union) { this.sal.push(`${union} j`); p.join = union; }
      if (p.miter !== e.miterLimit) { this.sal.push(`${num(e.miterLimit)} M`); p.miter = e.miterLimit; }
      const dash = e.dash.length ? `[${e.dash.map(num).join(' ')}] 0 d` : '[] 0 d';
      if (p.dash !== dash) { this.sal.push(dash); p.dash = dash; }
      this.sal.push(...ruta.operadores(), 'S');
    }

    _fijarRelleno(c) {
      const p = this.e.pdf, clave = `${c.r} ${c.g} ${c.b}`;
      if (p.fill !== clave) { this.sal.push(`${num(c.r / 255)} ${num(c.g / 255)} ${num(c.b / 255)} rg`); p.fill = clave; }
      this._fijarAlfa(c.a * this.e.globalAlpha);
    }
    _fijarAlfa(a) {
      a = Math.round(Math.min(1, Math.max(0, a)) * 1000) / 1000;
      const p = this.e.pdf;
      if (p.alfa === a) return;
      this.sal.push(`/G${this._alfa(a)} gs`);
      p.alfa = a;
    }
    _alfa(a) {
      let i = this.indiceAlfa.get(a);
      if (i === undefined) { i = this.alfas.length; this.alfas.push(a); this.indiceAlfa.set(a, i); }
      return i;
    }
    _registrarDegradado(g) { this.degradados.push(g); return this.degradados.length - 1; }

    // --- imágenes
    drawImage(fuente, ...a) {
      try {
        const ancho = fuente.naturalWidth || fuente.videoWidth || fuente.width, alto = fuente.naturalHeight || fuente.videoHeight || fuente.height;
        let sx = 0, sy = 0, sw = ancho, sh = alto, dx, dy, dw, dh;
        if (a.length === 2) [dx, dy] = a, dw = ancho, dh = alto;
        else if (a.length === 4) [dx, dy, dw, dh] = a;
        else [sx, sy, sw, sh, dx, dy, dw, dh] = a;
        if (!(sw > 0 && sh > 0 && dw !== 0 && dh !== 0) || !ancho || !alto) return;
        // Como el canvas: un destino con ancho o alto negativo es el mismo rectángulo, sin voltear.
        if (dw < 0) { dx += dw; dw = -dw; }
        if (dh < 0) { dy += dh; dh = -dh; }
        if (this._conSombra()) {
          this._sombra(this._cajaRuta(rutaRect(dx, dy, dw, dh), 0), (c) => { c.drawImage(fuente, sx, sy, sw, sh, dx, dy, dw, dh); });
          this._sinSombra(() => this.drawImage(fuente, ...a));
          return;
        }
        const i = this._registrarImagen(fuente, ancho, alto);
        // La imagen entera, estirada para que el pedazo (sx…) caiga en (dx…); si es
        // un pedazo, se recorta al destino.
        const kx = dw / sw, ky = dh / sh;
        const X = dx - sx * kx, Y = dy - sy * ky, W = ancho * kx, H = alto * ky;
        const recorta = sx !== 0 || sy !== 0 || sw !== ancho || sh !== alto;
        this._fijarAlfa(this.e.globalAlpha);
        this.sal.push('q');
        if (recorta) this.sal.push(...rutaRect(dx, dy, dw, dh).operadores(), 'W n');
        this.sal.push(cm([W, 0, 0, -H, X, Y + H]), `/I${i} Do`, 'Q');
      } catch (error) { this._fallar(error); }
    }
    _registrarImagen(fuente, ancho, alto) {
      let i = this.indiceImagen.get(fuente);
      if (i === undefined) { i = this.imagenes.length; this.imagenes.push({ fuente, ancho, alto }); this.indiceImagen.set(fuente, i); }
      return i;
    }

    // --- texto
    fillText(texto, x, y) {
      try {
        if (!texto) return;
        const e = this.e;
        const rell = e.fillStyle;
        const px = parseFloat(/([\d.]+)px/.exec(e.font) && /([\d.]+)px/.exec(e.font)[1]);
        const arch = typeof PptxTexto !== 'undefined' && PptxTexto.archivoDeFuente ? PptxTexto.archivoDeFuente(e.font, true) : null;
        const cargada = arch && fuenteLista.get(arch.url);
        const w = this.measureText(texto).width;
        let x0 = x;
        if (e.textAlign === 'center') x0 = x - w / 2; else if (e.textAlign === 'right' || e.textAlign === 'end') x0 = x - w;
        const ls = parseFloat(e.letterSpacing) || 0;
        const completa = cargada && [...texto].every((ch) => cargada.fk.hasGlyphForCodePoint(ch.codePointAt(0)));
        const vector = completa && esColor(rell) && e.textBaseline === 'alphabetic' && px > 0 && w > 0;
        const h = (px || 16);
        const pintarTexto = (c) => {
          c.font = e.font; c.letterSpacing = e.letterSpacing; c.fontKerning = e.fontKerning;
          c.textAlign = e.textAlign; c.textBaseline = e.textBaseline;
          c.fillStyle = this._real(rell, c); c.fillText(texto, x, y);
        };
        const cajaTexto = (extra) => this._cajaUsuario(x0 - h * 0.3, y - h * 1.4, w + h * 0.6, h * 2, extra);
        if (this._fueraDeRecortes(x0 - h * 0.3, y - h * 1.4, w + h * 0.6, h * 2)) return; // todo el texto cae fuera de un recorte activo: no se ve, y tampoco debe poder seleccionarse ni copiarse
        if (this._conSombra()) {
          this._sombra(cajaTexto(e.shadowBlur * 1.5 + Math.hypot(e.shadowOffsetX, e.shadowOffsetY) + 4), pintarTexto);
          this._sinSombra(() => this.fillText(texto, x, y));
          return;
        }
        if (!vector) { this._respaldo(cajaTexto(2), pintarTexto, false, !cargada ? 'texto con fuente sin archivo' : !completa ? 'texto con letras que la fuente no tiene' : 'texto en otra forma'); return; }
        const c = color(rell);
        this._fijarRelleno(c);
        let f = this.indiceFuente.get(arch.url);
        if (f === undefined) { f = this.fuentes.length; this.fuentes.push(arch.url); this.indiceFuente.set(arch.url, f); }
        const n = [...texto].length;
        this.sal.push((rec) => {
          const fuente = rec.fuentes[f];
          let natural = fuente.widthOfTextAtSize(texto, 1000) / 1000 * px + ls * n;
          const tz = natural > 0 ? Math.min(400, Math.max(25, 100 * w / natural)) : 100;
          return ['q', `${num(tz)} Tz`, ls ? `${num(ls / px)} Tc` : '', 'BT', `/F${f} 1 Tf`,
            `${num(px)} 0 0 ${num(-px)} ${num(x0)} ${num(y)} Tm`, `${fuente.encodeText(texto).toString()} Tj`, 'ET', 'Q'].filter(Boolean).join('\n');
        });
      } catch (error) { this._fallar(error); }
    }

    // ¿La caja (en coordenadas del usuario) queda toda fuera de alguno de los recortes activos? Compara con la caja de cada recorte (que contiene su trazado).
    _fueraDeRecortes(x, y, w, h) {
      const M = this.e.m;
      const pts = [aplicar(M, x, y), aplicar(M, x + w, y), aplicar(M, x, y + h), aplicar(M, x + w, y + h)];
      const bx0 = Math.min(...pts.map((p) => p[0])), bx1 = Math.max(...pts.map((p) => p[0])), by0 = Math.min(...pts.map((p) => p[1])), by1 = Math.max(...pts.map((p) => p[1]));
      for (const cl of this.e.clips) {
        const c = cl.ruta.caja();
        if (!c) return true; // recorte sin trazado: todo queda afuera
        const q = [aplicar(cl.m, c.x0, c.y0), aplicar(cl.m, c.x1, c.y0), aplicar(cl.m, c.x0, c.y1), aplicar(cl.m, c.x1, c.y1)];
        if (bx1 < Math.min(...q.map((p) => p[0])) || bx0 > Math.max(...q.map((p) => p[0])) || by1 < Math.min(...q.map((p) => p[1])) || by0 > Math.max(...q.map((p) => p[1]))) return true;
      }
      return false;
    }

    // --- respaldo: ese trazo como imagen
    _cajaUsuario(x, y, w, h, extra) {
      const M = this.e.m;
      const pts = [aplicar(M, x, y), aplicar(M, x + w, y), aplicar(M, x, y + h), aplicar(M, x + w, y + h)];
      return this._cajaPagina(Math.min(...pts.map((p) => p[0])) - extra, Math.min(...pts.map((p) => p[1])) - extra,
        Math.max(...pts.map((p) => p[0])) + extra, Math.max(...pts.map((p) => p[1])) + extra);
    }
    _cajaRuta(ruta, extra) {
      const c = ruta.caja();
      if (!c) return null;
      const sombra = this._conSombra() ? this.e.shadowBlur * 1.5 + Math.hypot(this.e.shadowOffsetX, this.e.shadowOffsetY) + 2 : 0;
      return this._cajaUsuario(c.x0, c.y0, c.x1 - c.x0, c.y1 - c.y0, extra + sombra);
    }
    _cajaPagina(x0, y0, x1, y1) {
      x0 = Math.max(0, Math.floor(x0)); y0 = Math.max(0, Math.floor(y0));
      x1 = Math.min(this.width, Math.ceil(x1)); y1 = Math.min(this.height, Math.ceil(y1));
      return x1 > x0 && y1 > y0 ? { x: x0, y: y0, w: x1 - x0, h: y1 - y0 } : null;
    }
    _copiarLinea(c) {
      const e = this.e;
      c.lineWidth = e.lineWidth; c.lineCap = e.lineCap; c.lineJoin = e.lineJoin; c.miterLimit = e.miterLimit; c.setLineDash(e.dash);
    }
    _real(valor, ctxReal) { return valor && typeof valor.aReal === 'function' ? valor.aReal(ctxReal) : valor; }
    // Pinta una operación en un canvas de la caja (en píxeles de la página) y
    // pone el resultado en su lugar. El recorte que esté activo en la página
    // sigue valiendo para la imagen.
    _respaldo(caja, pintar, soloSombra = false, motivo = 'otro') {
      if (!caja) return;
      this.motivos.push(soloSombra ? 'sombra' : motivo);
      const e = this.e, i = inv(e.m);
      if (!i) return;
      let r = Math.min(1, LADO_MAXIMO_RESPALDO / Math.max(caja.w, caja.h));
      // Solo la sombra (la figura se dibuja aparte, como vectores): es borrosa, no
      // hace falta toda la resolución. La figura se dibuja corrida D píxeles fuera
      // del canvas y la sombra se corre D de vuelta a su lugar.
      let D = 0;
      if (soloSombra) { r = Math.min(r, Math.max(0.04, 5 / Math.max(e.shadowBlur, 1))); D = caja.w + caja.h + e.shadowBlur * 3 + 50; }
      const lienzo = document.createElement('canvas');
      lienzo.width = Math.max(1, Math.ceil(caja.w * r)); lienzo.height = Math.max(1, Math.ceil(caja.h * r));
      const c = lienzo.getContext('2d');
      c.setTransform(r, 0, 0, r, -(caja.x + D) * r, -caja.y * r);
      c.transform(...e.m);
      c.globalAlpha = e.globalAlpha;
      c.imageSmoothingEnabled = e.imageSmoothingEnabled;
      c.shadowColor = e.shadowColor;
      c.shadowBlur = e.shadowBlur * r;
      c.shadowOffsetX = (e.shadowOffsetX + D) * r; c.shadowOffsetY = e.shadowOffsetY * r;
      pintar(c);
      const idx = this._registrarImagen(lienzo, lienzo.width, lienzo.height);
      this.respaldos++;
      const p = e.pdf;
      this.sal.push('q', cm(i));
      if (p.alfa !== 1) this.sal.push(`/G${this._alfa(1)} gs`);
      this.sal.push(cm([caja.w, 0, 0, -caja.h, caja.x, caja.y + caja.h]), `/I${idx} Do`, 'Q');
    }

    _sombra(caja, pintar) { this._respaldo(caja, pintar, true); }
    // Lo que se dibuja dentro va sin sombra (la sombra ya se puso aparte).
    _sinSombra(fn) {
      const e = this.e, guardada = e.shadowColor;
      e.shadowColor = 'rgba(0, 0, 0, 0)';
      try { fn(); } finally { e.shadowColor = guardada; }
    }

    // --- escribir la página
    async aPagina(pdf) {
      if (this.fallo) throw this.fallo;
      const ctx = pdf.context, { PDFName } = PDFLib;
      const compartido = cacheDe(pdf);
      if (this.fuentes.length) {
        await cargarFontkit();
        pdf.registerFontkit(fontkitSeguro());
      }
      const fuentes = [];
      for (const url of this.fuentes) {
        let f = compartido.fuentes.get(url);
        if (!f) {
          const datos = await fuentesCargadas.get(url);
          f = await pdf.embedFont(datos.bytes, { subset: true });
          compartido.fuentes.set(url, f);
        }
        fuentes.push(f);
      }
      const xobjetos = {};
      for (let i = 0; i < this.imagenes.length; i++) {
        const { fuente, ancho, alto } = this.imagenes[i];
        let ref = compartido.imagenes.get(fuente);
        if (!ref) { ref = await incrustarImagen(pdf, fuente, ancho, alto); compartido.imagenes.set(fuente, ref); }
        xobjetos['I' + i] = ref;
        if ((i & 15) === 15) await new Promise((listo) => setTimeout(listo, 0));
      }
      const estados = {};
      this.alfas.forEach((a, i) => { estados['G' + i] = ctx.register(ctx.obj({ Type: 'ExtGState', ca: a, CA: a })); });
      const sombreados = {};
      this.degradados.forEach((g, i) => { sombreados['S' + i] = ctx.register(sombreado(ctx, g)); });
      this.mascaras.forEach((m, j) => {
        const forma = ctx.register(ctx.flateStream('/SA sh', {
          Type: 'XObject', Subtype: 'Form', FormType: 1, BBox: m.caja, Group: { S: 'Transparency', CS: 'DeviceGray' },
          Resources: ctx.obj({ Shading: { SA: ctx.register(sombreado(ctx, m.degradado, true)) } }),
        }));
        estados['M' + j] = ctx.register(ctx.obj({ Type: 'ExtGState', SMask: { Type: 'Mask', S: 'Luminosity', G: forma } }));
      });
      const patrones = {};
      this.patrones.forEach((p, i) => {
        const contenido = `q ${num(p.ancho)} 0 0 ${num(-p.alto)} 0 ${num(p.alto)} cm /Im Do Q`;
        patrones['P' + i] = ctx.register(ctx.flateStream(contenido, {
          Type: 'Pattern', PatternType: 1, PaintType: 1, TilingType: 1, BBox: [0, 0, p.ancho, p.alto], XStep: p.ancho, YStep: p.alto,
          Matrix: p.matriz, Resources: ctx.obj({ XObject: { Im: xobjetos['I' + p.imagen] } }),
        }));
      });
      const fontDict = {};
      fuentes.forEach((f, i) => { fontDict['F' + i] = f.ref; });
      const recursos = ctx.obj({ ProcSet: ['PDF', 'Text', 'ImageC', 'ImageB'], Font: fontDict, XObject: xobjetos, ExtGState: estados, Shading: sombreados, Pattern: patrones });
      const resolver = { fuentes };
      const f = this.anchoPt / this.width;
      const lineas = ['q', cm([f, 0, 0, -f, 0, this.altoPt]), ...this.sal.map((s) => (typeof s === 'function' ? s(resolver) : s))];
      for (let i = 0; i < this.pila.length; i++) lineas.push('Q');
      lineas.push('Q');
      const pagina = pdf.addPage([this.anchoPt, this.altoPt]);
      pagina.node.set(PDFName.of('Resources'), recursos);
      pagina.node.set(PDFName.of('Contents'), ctx.register(ctx.flateStream(lineas.join('\n'))));
    }
  }

  // Cada método que pinta o recorta avisa en `fallo` si algo sale mal.
  for (const nombre of ['save', 'restore', 'transform', 'setTransform']) {
    const original = Lienzo.prototype[nombre];
    Lienzo.prototype[nombre] = function (...a) { try { return original.apply(this, a); } catch (error) { this._fallar(error); } };
  }
  for (const p of PROPIEDADES) {
    Object.defineProperty(Lienzo.prototype, p, { get() { return this.e[p]; }, set(v) { this.e[p] = v; }, configurable: true });
  }

  function rutaRect(x, y, w, h) { const r = new Ruta(); r.rect(x, y, w, h); return r; }

  // Lo que se comparte entre las páginas de un mismo PDF: fuentes e imágenes.
  const caches = new WeakMap();
  function cacheDe(pdf) {
    let c = caches.get(pdf);
    if (!c) { c = { fuentes: new Map(), imagenes: new Map() }; caches.set(pdf, c); }
    return c;
  }

  // Un JPEG de 1 o 3 componentes, de 8 bits y sin giro EXIF, se puede poner en el
  // PDF tal cual (DCTDecode). Si es CMYK (4 componentes, con tabla invertida de
  // Adobe), tiene giro EXIF o no se entiende, devuelve null y va por el canvas.
  function infoJpeg(b) {
    if (b[0] !== 0xff || b[1] !== 0xd8) return null;
    let i = 2, orientacion = 1, info = null;
    while (i + 4 < b.length) {
      if (b[i] !== 0xff) { i++; continue; }
      const m = b[i + 1];
      if (m === 0xff) { i++; continue; }
      if (m === 0xd8 || m === 0x01 || (m >= 0xd0 && m <= 0xd7)) { i += 2; continue; }
      if (m === 0xd9 || m === 0xda) break;
      const largo = (b[i + 2] << 8) | b[i + 3];
      if (m === 0xe1 && b[i + 4] === 0x45 && b[i + 5] === 0x78 && b[i + 6] === 0x69 && b[i + 7] === 0x66) {
        const t = i + 10, le = b[t] === 0x49;
        const u16 = (o) => (le ? b[t + o] | (b[t + o + 1] << 8) : (b[t + o] << 8) | b[t + o + 1]);
        const u32 = (o) => (le ? (b[t + o] | (b[t + o + 1] << 8) | (b[t + o + 2] << 16) | (b[t + o + 3] << 24)) >>> 0 : ((b[t + o] << 24) | (b[t + o + 1] << 16) | (b[t + o + 2] << 8) | b[t + o + 3]) >>> 0);
        const ifd = u32(4), n = u16(ifd);
        for (let k = 0; k < n && k < 64; k++) if (u16(ifd + 2 + k * 12) === 0x0112) orientacion = u16(ifd + 2 + k * 12 + 8);
      }
      if (m === 0xe2 && b[i + 4] === 0x49 && b[i + 5] === 0x43 && b[i + 6] === 0x43) {
        // Perfil ICC: solo sRGB va tal cual (el canvas convierte los demás a sRGB).
        let srgb = false;
        for (let k = i + 18; k < Math.min(i + 2 + largo - 3, i + 400); k++) if (b[k] === 0x73 && b[k + 1] === 0x52 && b[k + 2] === 0x47 && b[k + 3] === 0x42) { srgb = true; break; }
        if (!srgb) return null;
      }
      if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) {
        info = { precision: b[i + 4], alto: (b[i + 5] << 8) | b[i + 6], ancho: (b[i + 7] << 8) | b[i + 8], componentes: b[i + 9] };
      }
      i += 2 + largo;
    }
    if (!info || info.precision !== 8 || (info.componentes !== 1 && info.componentes !== 3) || orientacion > 1 || !info.ancho || !info.alto) return null;
    return info;
  }

  // Una imagen (canvas, ImageBitmap…) como XObject: colores y transparencia por
  // separado, comprimidos con Flate (sin pasar por PNG).
  async function incrustarImagen(pdf, fuente, ancho, alto) {
    const ctxPdf = pdf.context;
    const jpeg = fuente.bytesJpeg ? infoJpeg(fuente.bytesJpeg) : null;
    if (jpeg) {
      return ctxPdf.register(ctxPdf.stream(fuente.bytesJpeg, {
        Type: 'XObject', Subtype: 'Image', Width: jpeg.ancho, Height: jpeg.alto, Filter: 'DCTDecode',
        ColorSpace: jpeg.componentes === 1 ? 'DeviceGray' : 'DeviceRGB', BitsPerComponent: 8, Interpolate: true,
      }));
    }
    let lienzo = fuente;
    if (!(fuente instanceof HTMLCanvasElement)) {
      lienzo = document.createElement('canvas'); lienzo.width = ancho; lienzo.height = alto;
      lienzo.getContext('2d').drawImage(fuente, 0, 0, ancho, alto);
    }
    const datos = lienzo.getContext('2d', { willReadFrequently: true }).getImageData(0, 0, lienzo.width, lienzo.height).data;
    const total = lienzo.width * lienzo.height;
    const rgb = new Uint8Array(total * 3), alfa = new Uint8Array(total);
    let translucida = false;
    for (let p = 0, k = 0, r = 0; p < total; p++, k += 4, r += 3) {
      rgb[r] = datos[k]; rgb[r + 1] = datos[k + 1]; rgb[r + 2] = datos[k + 2];
      const a = datos[k + 3]; alfa[p] = a; if (a !== 255) translucida = true;
    }
    // Los mosaicos chicos (tramas de 8×8) tienen que quedar nítidos, no suavizados.
    const suave = lienzo.width > 64 || lienzo.height > 64;
    const dict = { Type: 'XObject', Subtype: 'Image', Width: lienzo.width, Height: lienzo.height, ColorSpace: 'DeviceRGB', BitsPerComponent: 8, Interpolate: suave };
    if (translucida) {
      const mascara = ctxPdf.flateStream(alfa, { Type: 'XObject', Subtype: 'Image', Width: lienzo.width, Height: lienzo.height, ColorSpace: 'DeviceGray', BitsPerComponent: 8, Interpolate: suave });
      dict.SMask = ctxPdf.register(mascara);
    }
    return ctxPdf.register(ctxPdf.flateStream(rgb, dict));
  }

  // Degradado de canvas → Shading de PDF (tipo 2 axial, tipo 3 radial) con una
  // función que une un tramo por cada par de paradas.
  function sombreado(ctx, g, soloAlfa = false) {
    let paradas = g.paradas.map((p) => ({ pos: p.pos, color: p.color })).sort((a, b) => a.pos - b.pos);
    if (paradas[0].pos > 0) paradas.unshift({ pos: 0, color: paradas[0].color });
    if (paradas[paradas.length - 1].pos < 1) paradas.push({ pos: 1, color: paradas[paradas.length - 1].color });
    const rgb = (c) => (soloAlfa ? [c.a] : [c.r / 255, c.g / 255, c.b / 255]);
    const tramos = [], limites = [];
    for (let i = 0; i < paradas.length - 1; i++) {
      tramos.push(ctx.obj({ FunctionType: 2, Domain: [0, 1], C0: rgb(paradas[i].color), C1: rgb(paradas[i + 1].color), N: 1 }));
      if (i > 0) limites.push(Math.min(1 - 1e-6, Math.max(paradas[i].pos, (limites.length ? limites[limites.length - 1] : 0) + 1e-6)));
    }
    const funcion = tramos.length === 1 ? tramos[0]
      : ctx.obj({ FunctionType: 3, Domain: [0, 1], Functions: tramos, Bounds: limites, Encode: tramos.flatMap(() => [0, 1]) });
    return ctx.obj({
      ShadingType: g.tipo === 'lineal' ? 2 : 3, ColorSpace: soloAlfa ? 'DeviceGray' : 'DeviceRGB', Coords: g.coords, Function: funcion, Extend: [true, true],
    });
  }

  return { Lienzo, Ruta, precargarFuentes, cargarFontkit };
})();

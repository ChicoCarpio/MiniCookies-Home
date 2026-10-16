/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Rejilla de hoja de cálculo (interfaz de hojas, fase A del plan 06): muestra una hoja de Excel, de un CSV o de un formato antiguo como se ve al abrirla
// en Excel o en Google Sheets: letras y números en los bordes, celda activa, selección de rango, cuadro de nombre, barra con el contenido de la celda,
// pestañas de hojas, barra de estado (Recuento, Suma, Promedio) y zoom. La hoja la dibuja el motor (XlsxDibujo.dibujarVista) en un Canvas del tamaño de la
// pantalla; el desplazamiento es el del navegador (un espaciador del tamaño de la hoja), así que en el celular la inercia y las barras son las del sistema.
// Solo lectura: nada se escribe en el archivo.
//
// Uso:
//   await MotorHojas.cargarDibujo();
//   const lector = await LectorHojas.abrir(archivo);
//   const rejilla = await RejillaHoja.crear(contenedor, { libro: lector });   // contenedor: un elemento con alto propio (la rejilla lo llena)
//   rejilla.mostrarHoja(indice)  ·  rejilla.seleccion() → { f0, c0, f1, c1, nombre }  ·  rejilla.zoom(100)  ·  rejilla.alCambiar(fn)  ·  rejilla.destruir()
// Requiere motor-hojas.js ya cargado (cargarDibujo), lector-hojas.js, idioma.js (t) y rejilla-hoja.css. Los textos están en window.i18nPagina ('rh-…').
(function () {
  const MAX_FILAS = 1048576, MAX_COLS = 16384;
  const ZOOM_MIN = 50, ZOOM_MAX = 200;
  const AREA_DIRECTA = 20000;        // hasta esta cantidad de celdas la barra de estado recorre el rango; más, recorre las celdas con contenido
  const MAX_COPIA = 2000000;         // celdas que se copian como mucho
  const MAX_COMBINADAS_INDICE = 5000;

  // Textos de la rejilla: se suman a los de la página (los de la página mandan)
  (function () {
    const pagina = window.i18nPagina || (window.i18nPagina = {});
    const es = {
      'rh-rejilla': 'Hoja de cálculo', 'rh-hojas': 'Hojas', 'rh-nombre': 'Cuadro de nombre', 'rh-contenido': 'Contenido de la celda', 'rh-listo': 'Listo',
      'rh-promedio': 'Promedio', 'rh-recuento': 'Recuento', 'rh-suma': 'Suma', 'rh-zoom': 'Zoom', 'rh-alejar': 'Alejar', 'rh-acercar': 'Acercar', 'rh-vacia': 'Hoja vacía',
      'rh-seleccion': 'Selección',
    };
    const en = {
      'rh-rejilla': 'Spreadsheet', 'rh-hojas': 'Sheets', 'rh-nombre': 'Name box', 'rh-contenido': 'Cell contents', 'rh-listo': 'Ready',
      'rh-promedio': 'Average', 'rh-recuento': 'Count', 'rh-suma': 'Sum', 'rh-zoom': 'Zoom', 'rh-alejar': 'Zoom out', 'rh-acercar': 'Zoom in', 'rh-vacia': 'Empty sheet',
      'rh-seleccion': 'Selection',
    };
    pagina.es = Object.assign(es, pagina.es);
    pagina.en = Object.assign(en, pagina.en);
  })();
  const tx = (clave) => (typeof t === 'function' ? t(clave) : clave);
  const idiomaPagina = () => (typeof idiomaActual === 'function' && idiomaActual() === 'en' ? 'en' : 'es');

  const colNombre = (c) => LibroXlsx.numeroACol(c);
  const aA1 = (f, c) => colNombre(c) + (f + 1);
  const limitar = (v, a, b) => Math.max(a, Math.min(b, v));

  async function crear(contenedor, opciones) {
    const lector = opciones.libro, libro = lector.libroMotor;
    if (!libro) throw new Error('rejilla-hoja: el lector no trae libroMotor');
    const idioma = idiomaPagina();
    const formatoNumero = new Intl.NumberFormat(idioma === 'en' ? 'en-US' : 'es-CR', { maximumFractionDigits: 6 });
    const formatoBarra = new Intl.NumberFormat(idioma === 'en' ? 'en-US' : 'es-CR', { useGrouping: false, maximumFractionDigits: 15 });   // la barra de contenido muestra el número sin separador de miles, como Excel
    const escuchas = [];
    const escuchar = (el, tipo, fn, op) => { el.addEventListener(tipo, fn, op); escuchas.push(() => el.removeEventListener(tipo, fn, op)); };
    const cambios = [];
    let destruida = false;

    await XlsxDibujo.cargarFuentes(libro);

    // ---- DOM -----------------------------------------------------------------------------------------------------
    contenedor.textContent = '';
    const raiz = document.createElement('div');
    raiz.className = 'rh';
    raiz.innerHTML =
      '<div class="rh-barra"><output class="rh-nombre" aria-label="' + tx('rh-nombre') + '"></output><span class="rh-fx" aria-hidden="true">fx</span>' +
      '<div class="rh-formula" role="textbox" aria-readonly="true" aria-label="' + tx('rh-contenido') + '"></div></div>' +
      '<div class="rh-cuerpo" tabindex="0" role="grid" aria-label="' + tx('rh-rejilla') + '">' +
      '<div class="rh-vista"><canvas class="rh-lienzo"></canvas><div class="rh-capa"><div class="rh-enc-col"></div><div class="rh-enc-fila"></div>' +
      '<div class="rh-recorte"><div class="rh-rango"></div><div class="rh-activa"></div></div></div></div><div class="rh-espaciador"></div></div>' +
      '<div class="rh-pestanas" role="tablist" aria-label="' + tx('rh-hojas') + '"></div>' +
      '<div class="rh-estado"><span class="rh-info" role="status" aria-live="polite"></span><span class="rh-zoom">' +
      '<button type="button" class="rh-menos" aria-label="' + tx('rh-alejar') + '">−</button>' +
      '<input type="range" class="rh-control-zoom" min="' + ZOOM_MIN + '" max="' + ZOOM_MAX + '" step="10" value="100" aria-label="' + tx('rh-zoom') + '">' +
      '<button type="button" class="rh-mas" aria-label="' + tx('rh-acercar') + '">+</button><output class="rh-valor-zoom">100%</output></span></div>';
    contenedor.appendChild(raiz);
    const q = (s) => raiz.querySelector(s);
    const cajaNombre = q('.rh-nombre'), cajaFormula = q('.rh-formula'), cuerpo = q('.rh-cuerpo'), vista = q('.rh-vista'), lienzo = q('.rh-lienzo');
    const capa = q('.rh-capa'), encCol = q('.rh-enc-col'), encFila = q('.rh-enc-fila'), recorte = q('.rh-recorte'), rangoDiv = q('.rh-rango'), activaDiv = q('.rh-activa');
    const espaciador = q('.rh-espaciador'), pestanas = q('.rh-pestanas'), info = q('.rh-info');
    const controlZoom = q('.rh-control-zoom'), valorZoom = q('.rh-valor-zoom');
    const ctx = lienzo.getContext('2d');

    // ---- La hoja abierta -------------------------------------------------------------------------------------------
    let est = null;          // { indice, hoja, g, nFilas, nCols, fm, cm, sel, zoom, areaX, areaY, combinadas }
    const estados = new Map();   // por hoja: selección, desplazamiento y zoom que tenía
    let W = 0, H = 0, dpr = 1, pendiente = false, vistaReal = { areaX: 26, areaY: 20 };

    function indiceCombinadas(hoja) {
      const lista = (hoja.combinadas || []).slice();
      if (lista.length > MAX_COMBINADAS_INDICE) return { lista: [], mapa: new Map(lista.map((m) => [m.f0 * MAX_COLS + m.c0, m])) };
      return { lista, mapa: null };
    }
    function combinadaEn(f, c) {
      const ic = est.combinadas;
      if (ic.mapa) return ic.mapa.get(f * MAX_COLS + c) || null;
      for (const m of ic.lista) if (f >= m.f0 && f <= m.f1 && c >= m.c0 && c <= m.c1) return m;
      return null;
    }
    // la selección crece hasta cubrir enteras las celdas combinadas que toca (como Excel)
    function expandir(r) {
      if (!est.combinadas.lista.length) return r;
      let { f0, c0, f1, c1 } = r, cambio = true, vueltas = 0;
      while (cambio && vueltas++ < 10) {
        cambio = false;
        for (const m of est.combinadas.lista) {
          if (m.f1 < f0 || m.f0 > f1 || m.c1 < c0 || m.c0 > c1) continue;
          if (m.f0 < f0) { f0 = m.f0; cambio = true; } if (m.f1 > f1) { f1 = m.f1; cambio = true; }
          if (m.c0 < c0) { c0 = m.c0; cambio = true; } if (m.c1 > c1) { c1 = m.c1; cambio = true; }
        }
      }
      return { f0, c0, f1, c1 };
    }
    const rango = () => {
      const s = est.sel;
      return expandir({ f0: Math.min(s.f, s.f2), c0: Math.min(s.c, s.c2), f1: Math.max(s.f, s.f2), c1: Math.max(s.c, s.c2) });
    };
    function nombreRango(r) {
      const entera = r.f0 === 0 && r.f1 >= MAX_FILAS - 1, enteraFila = r.c0 === 0 && r.c1 >= MAX_COLS - 1;
      if (entera && enteraFila) return 'A1:' + aA1(MAX_FILAS - 1, MAX_COLS - 1);
      if (entera) return colNombre(r.c0) + ':' + colNombre(r.c1);
      if (enteraFila) return (r.f0 + 1) + ':' + (r.f1 + 1);
      return r.f0 === r.f1 && r.c0 === r.c1 ? aA1(r.f0, r.c0) : aA1(r.f0, r.c0) + ':' + aA1(r.f1, r.c1);
    }

    async function abrirHoja(indice) {
      const hoja = await libro.cargarHoja(indice);
      await XlsxDibujo.prepararObjetos(libro, hoja, null);
      const C = hoja.celdas;
      const fm = C.filaMax, cm = C.colMax;
      const guardado = estados.get(indice);
      const e = {
        indice, hoja, g: XlsxDibujo.geometria(libro, hoja), fm, cm,
        nFilas: limitar(fm + 41, 60, MAX_FILAS), nCols: limitar(cm + 11, 26, MAX_COLS),
        sel: guardado ? guardado.sel : { f: 0, c: 0, f2: 0, c2: 0 }, zoom: guardado ? guardado.zoom : 100, scrollX: guardado ? guardado.scrollX : 0, scrollY: guardado ? guardado.scrollY : 0,
        combinadas: indiceCombinadas(hoja),
      };
      if (est) estados.set(est.indice, { sel: est.sel, zoom: est.zoom, scrollX: cuerpo.scrollLeft, scrollY: cuerpo.scrollTop });
      est = e;
      colocarTamano();
      cuerpo.scrollLeft = e.scrollX; cuerpo.scrollTop = e.scrollY;
      aplicarZoomControles();
      pintar();
      actualizarBarras();
      marcarPestana();
    }

    // ---- Tamaños y pintado ---------------------------------------------------------------------------------------------
    function colocarTamano() {
      if (!est) return;
      const z = est.zoom / 100;
      const encMax = Math.max(26, String(est.nFilas).length * 8 + 12);
      espaciador.style.width = Math.ceil((est.g.x(est.nCols) + encMax) * z) + 'px';
      espaciador.style.height = Math.ceil((est.g.y(est.nFilas) + 20) * z) + 'px';
    }
    function medirVista() {
      W = cuerpo.clientWidth; H = cuerpo.clientHeight; dpr = window.devicePixelRatio || 1;
      vista.style.width = W + 'px'; vista.style.height = H + 'px';
      lienzo.width = Math.max(1, Math.round(W * dpr)); lienzo.height = Math.max(1, Math.round(H * dpr));
      lienzo.style.width = W + 'px'; lienzo.style.height = H + 'px';
    }
    function pedirPintado() { if (pendiente || destruida) return; pendiente = true; requestAnimationFrame(() => { pendiente = false; if (!destruida) pintar(); }); }
    function pintar() {
      if (!est || !W || !H) return;
      const z = est.zoom / 100;
      ctx.setTransform(dpr * z, 0, 0, dpr * z, 0, 0);
      vistaReal = XlsxDibujo.dibujarVista(ctx, libro, est.hoja, est.g, { x: cuerpo.scrollLeft / z, y: cuerpo.scrollTop / z, ancho: W / z, alto: H / z }, { idioma });
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      pintarSeleccion();
    }

    // posición en la pantalla (px de la vista) de una columna / fila, con el zoom y el desplazamiento de ahora
    const pantallaX = (c) => (vistaReal.areaX + est.g.x(c) - cuerpo.scrollLeft / (est.zoom / 100)) * (est.zoom / 100);
    const pantallaY = (f) => (vistaReal.areaY + est.g.y(f) - cuerpo.scrollTop / (est.zoom / 100)) * (est.zoom / 100);
    function colocar(div, x0, y0, x1, y1) {
      const a = limitar(x0, -20, W + 20), b = limitar(x1, -20, W + 20), c = limitar(y0, -20, H + 20), d = limitar(y1, -20, H + 20);
      div.style.display = b <= a || d <= c ? 'none' : 'block';
      div.style.left = a + 'px'; div.style.top = c + 'px'; div.style.width = (b - a) + 'px'; div.style.height = (d - c) + 'px';
    }
    function pintarSeleccion() {
      const z = est.zoom / 100, ax = vistaReal.areaX * z, ay = vistaReal.areaY * z, r = rango();
      const x0 = pantallaX(r.c0), x1 = pantallaX(r.c1 + 1), y0 = pantallaY(r.f0), y1 = pantallaY(r.f1 + 1);
      colocar(encCol, Math.max(x0, ax), 0, x1, ay);
      colocar(encFila, 0, Math.max(y0, ay), ax, y1);
      recorte.style.left = ax + 'px'; recorte.style.top = ay + 'px'; recorte.style.width = Math.max(0, W - ax) + 'px'; recorte.style.height = Math.max(0, H - ay) + 'px';
      const unica = r.f0 === r.f1 && r.c0 === r.c1;
      const m = combinadaEn(est.sel.f, est.sel.c);
      const ac = m || { f0: est.sel.f, c0: est.sel.c, f1: est.sel.f, c1: est.sel.c };
      colocar(activaDiv, pantallaX(ac.c0) - ax, pantallaY(ac.f0) - ay, pantallaX(ac.c1 + 1) - ax, pantallaY(ac.f1 + 1) - ay);
      const mismo = unica || (m && r.f0 === m.f0 && r.c0 === m.c0 && r.f1 === m.f1 && r.c1 === m.c1);
      colocar(rangoDiv, x0 - ax, y0 - ay, x1 - ax, y1 - ay);
      rangoDiv.hidden = !!mismo;
    }

    // ---- Barras ----------------------------------------------------------------------------------------------------
    function textoDeCelda(f, c) {
      const i = est.hoja.celdas.indiceDe(f, c);
      return i < 0 ? '' : LibroXlsx.textoCelda(libro, est.hoja, i, lector.opciones).texto;
    }
    function contenidoDeCelda(f, c) {
      const C = est.hoja.celdas, i = C.indiceDe(f, c);
      if (i < 0) return '';
      const fx = C.formulaDe(i);
      if (fx) return '=' + fx;
      const T = LibroXlsx.TIPO, tipo = C.tipo[i];
      if (tipo === T.vacio) return '';
      if (tipo === T.numero) {
        const codigo = libro.estilos.xf(C.estilo[i]).formato;
        if (codigo && XlsxFormato.esFecha(codigo)) return textoDeCelda(f, c);
        return formatoBarra.format(C.valor[i]);
      }
      return textoDeCelda(f, c);
    }
    let temporizadorEstado = 0;
    function estadistica() {
      const r = rango(), C = est.hoja.celdas, T = LibroXlsx.TIPO;
      const f1 = Math.min(r.f1, est.fm), c1 = Math.min(r.c1, est.cm);
      let n = 0, nums = 0, suma = 0;
      if (f1 >= r.f0 && c1 >= r.c0) {
        const area = (f1 - r.f0 + 1) * (c1 - r.c0 + 1);
        const cuenta = (i) => {
          const tipo = C.tipo[i];
          if (tipo === T.vacio) return;
          n++;
          if (tipo === T.numero) { nums++; suma += C.valor[i]; }
        };
        if (area <= AREA_DIRECTA) { for (let f = r.f0; f <= f1; f++) for (let c = r.c0; c <= c1; c++) { const i = C.indiceDe(f, c); if (i >= 0) cuenta(i); } }
        else C.paraCada((f, c, i) => { if (f >= r.f0 && f <= f1 && c >= r.c0 && c <= c1) cuenta(i); });
      }
      info.textContent = '';
      if (n <= 1 && nums <= 1) { info.textContent = tx('rh-listo'); return; }
      const parte = (clave, valor) => { const s = document.createElement('span'); s.textContent = tx(clave) + ': '; const b = document.createElement('b'); b.textContent = valor; s.appendChild(b); info.appendChild(s); };
      if (nums > 1) parte('rh-promedio', formatoNumero.format(suma / nums));
      parte('rh-recuento', formatoNumero.format(n));
      if (nums > 1) parte('rh-suma', formatoNumero.format(suma));
    }
    function actualizarBarras() {
      const r = rango();
      cajaNombre.textContent = nombreRango(r);
      const m = combinadaEn(est.sel.f, est.sel.c);
      const texto = contenidoDeCelda(m ? m.f0 : est.sel.f, m ? m.c0 : est.sel.c);
      cajaFormula.textContent = texto;
      cajaFormula.classList.toggle('es-formula', texto.charAt(0) === '=');
      clearTimeout(temporizadorEstado);
      const area = (r.f1 - r.f0 + 1) * (r.c1 - r.c0 + 1);
      if (area > AREA_DIRECTA) temporizadorEstado = setTimeout(() => { if (!destruida && est) estadistica(); }, 120); else estadistica();
      for (const fn of cambios) fn(seleccion());
    }
    function seleccion() { const r = rango(); return { f0: r.f0, c0: r.c0, f1: r.f1, c1: r.c1, nombre: nombreRango(r), hoja: est ? est.indice : -1 }; }

    // ---- Selección y movimiento ---------------------------------------------------------------------------------------------
    function ponerSeleccion(f, c, f2, c2) {
      est.sel = { f: limitar(f, 0, MAX_FILAS - 1), c: limitar(c, 0, MAX_COLS - 1), f2: limitar(f2, 0, MAX_FILAS - 1), c2: limitar(c2, 0, MAX_COLS - 1) };
      pintarSeleccion(); actualizarBarras();
    }
    function mostrarCelda(f, c) {
      const z = est.zoom / 100, g = est.g;
      const ax = vistaReal.areaX * z, ay = vistaReal.areaY * z;
      const x0 = g.x(c) * z, x1 = g.x(c + 1) * z, y0 = g.y(f) * z, y1 = g.y(f + 1) * z;
      if (x0 < cuerpo.scrollLeft) cuerpo.scrollLeft = x0; else if (x1 > cuerpo.scrollLeft + W - ax) cuerpo.scrollLeft = x1 - (W - ax);
      if (y0 < cuerpo.scrollTop) cuerpo.scrollTop = y0; else if (y1 > cuerpo.scrollTop + H - ay) cuerpo.scrollTop = y1 - (H - ay);
    }
    // la siguiente fila o columna que se ve (las ocultas miden 0)
    function salto(desde, paso, mide, tope) {
      let k = desde + paso;
      while (k > 0 && k < tope - 1 && mide(k) === 0) k += paso;
      return limitar(k, 0, tope - 1);
    }
    function mover(df, dc, extender, ctrl) {
      const g = est.g, s = est.sel;
      let f = extender ? s.f2 : s.f, c = extender ? s.c2 : s.c;
      if (ctrl) {   // Ctrl + flecha: hasta el borde del contenido
        if (df) f = df < 0 ? 0 : Math.max(est.fm, 0); if (dc) c = dc < 0 ? 0 : Math.max(est.cm, 0);
      } else {
        if (df) f = salto(f, df, g.altoFila, est.nFilas);
        if (dc) c = salto(c, dc, g.anchoCol, est.nCols);
      }
      if (extender) ponerSeleccion(s.f, s.c, f, c); else ponerSeleccion(f, c, f, c);
      mostrarCelda(f, c);
      pedirPintado();
    }
    function enTecla(e) {
      if (!est) return;
      const ctrl = e.ctrlKey || e.metaKey, ext = e.shiftKey, s = est.sel;
      const pagina = Math.max(1, Math.floor((H / (est.zoom / 100) - vistaReal.areaY) / Math.max(1, est.g.altoPorDefecto)) - 1);
      let hecho = true;
      switch (e.key) {
        case 'ArrowUp': mover(-1, 0, ext, ctrl); break;
        case 'ArrowDown': mover(1, 0, ext, ctrl); break;
        case 'ArrowLeft': mover(0, -1, ext, ctrl); break;
        case 'ArrowRight': mover(0, 1, ext, ctrl); break;
        case 'PageUp': { const f = Math.max(0, (ext ? s.f2 : s.f) - pagina); ext ? ponerSeleccion(s.f, s.c, f, s.c2) : ponerSeleccion(f, s.c, f, s.c); mostrarCelda(f, ext ? s.c2 : s.c); pedirPintado(); break; }
        case 'PageDown': { const f = Math.min(MAX_FILAS - 1, (ext ? s.f2 : s.f) + pagina); ext ? ponerSeleccion(s.f, s.c, f, s.c2) : ponerSeleccion(f, s.c, f, s.c); mostrarCelda(f, ext ? s.c2 : s.c); pedirPintado(); break; }
        case 'Home': if (ctrl) { ponerSeleccion(0, 0, 0, 0); mostrarCelda(0, 0); } else { const f = ext ? s.f2 : s.f; ext ? ponerSeleccion(s.f, s.c, f, 0) : ponerSeleccion(f, 0, f, 0); mostrarCelda(f, 0); } pedirPintado(); break;
        case 'End': if (ctrl) { const f = Math.max(0, est.fm), c = Math.max(0, est.cm); ponerSeleccion(f, c, f, c); mostrarCelda(f, c); pedirPintado(); } else hecho = false; break;
        case 'a': case 'A': if (ctrl) { ponerSeleccion(0, 0, MAX_FILAS - 1, MAX_COLS - 1); } else hecho = false; break;
        case 'Escape': ponerSeleccion(s.f, s.c, s.f, s.c); break;
        default: hecho = false;
      }
      if (hecho) e.preventDefault();
    }
    escuchar(cuerpo, 'keydown', enTecla);

    // Ratón y tacto. El ratón selecciona al apretar y al arrastrar; el tacto, solo con un toque corto sin moverse (si se mueve es un desplazamiento).
    function punto(e) {
      const r = vista.getBoundingClientRect(), z = est.zoom / 100;
      return { x: (e.clientX - r.left) / z, y: (e.clientY - r.top) / z };
    }
    function celdaEn(p) {
      const z = est.zoom / 100, vx = cuerpo.scrollLeft / z, vy = cuerpo.scrollTop / z, g = est.g;
      const dentroX = p.x >= vistaReal.areaX, dentroY = p.y >= vistaReal.areaY;
      return {
        enc: !dentroX || !dentroY, esquina: !dentroX && !dentroY,
        c: dentroX ? limitar(g.colEn(p.x - vistaReal.areaX + vx), 0, MAX_COLS - 1) : -1,
        f: dentroY ? limitar(g.filaEn(p.y - vistaReal.areaY + vy), 0, MAX_FILAS - 1) : -1,
      };
    }
    let arrastre = null, toque = null;
    function aplicarPunto(h, extender) {
      const s = est.sel;
      if (h.esquina) { ponerSeleccion(0, 0, MAX_FILAS - 1, MAX_COLS - 1); return; }
      if (h.f < 0) { extender ? ponerSeleccion(s.f, s.c, s.f2, h.c) : ponerSeleccion(0, h.c, MAX_FILAS - 1, h.c); return; }       // encabezado de columna
      if (h.c < 0) { extender ? ponerSeleccion(s.f, s.c, h.f, s.c2) : ponerSeleccion(h.f, 0, h.f, MAX_COLS - 1); return; }       // encabezado de fila
      extender ? ponerSeleccion(s.f, s.c, h.f, h.c) : ponerSeleccion(h.f, h.c, h.f, h.c);
    }
    escuchar(vista, 'pointerdown', (e) => {
      if (!est || e.button !== 0) return;
      cuerpo.focus({ preventScroll: true });
      const h = celdaEn(punto(e));
      if (e.pointerType === 'mouse') {
        aplicarPunto(h, e.shiftKey);
        arrastre = { id: e.pointerId, columnas: h.f < 0 && !h.esquina, filas: h.c < 0 && !h.esquina };
        try { vista.setPointerCapture(e.pointerId); } catch (x) { /* sin captura */ }
      } else toque = { x: e.clientX, y: e.clientY, t: Date.now(), h, id: e.pointerId };
    });
    escuchar(vista, 'pointermove', (e) => {
      if (!est || !arrastre || e.pointerId !== arrastre.id || !(e.buttons & 1)) return;
      const h = celdaEn(punto(e)), s = est.sel;
      if (arrastre.columnas) ponerSeleccion(s.f, s.c, s.f2, Math.max(0, h.c));
      else if (arrastre.filas) ponerSeleccion(s.f, s.c, Math.max(0, h.f), s.c2);
      else if (h.f >= 0 && h.c >= 0) ponerSeleccion(s.f, s.c, h.f, h.c);
      // junto al borde, la hoja se desplaza sola
      const r = vista.getBoundingClientRect();
      cuerpo.scrollLeft += e.clientX > r.right - 12 ? 24 : e.clientX < r.left + vistaReal.areaX * (est.zoom / 100) + 12 ? -24 : 0;
      cuerpo.scrollTop += e.clientY > r.bottom - 12 ? 24 : e.clientY < r.top + vistaReal.areaY * (est.zoom / 100) + 12 ? -24 : 0;
    });
    const soltar = (e) => { if (arrastre && e.pointerId === arrastre.id) arrastre = null; };
    escuchar(vista, 'pointerup', (e) => {
      soltar(e);
      if (toque && e.pointerId === toque.id) {
        const t = toque; toque = null;
        if (Math.abs(e.clientX - t.x) < 8 && Math.abs(e.clientY - t.y) < 8 && Date.now() - t.t < 600 && est) aplicarPunto(t.h, false);
      }
    });
    escuchar(vista, 'pointercancel', (e) => { soltar(e); toque = null; });

    // Copiar: el texto que se ve, separado por tabuladores (como Excel); solo la parte con contenido
    escuchar(cuerpo, 'copy', (e) => {
      if (!est) return;
      const texto = textoCopia();
      if (e.clipboardData) { e.clipboardData.setData('text/plain', texto); e.preventDefault(); }
    });
    function textoCopia() {
      const r = rango(), f1 = Math.min(r.f1, est.fm), c1 = Math.min(r.c1, est.cm);
      if (f1 < r.f0 || c1 < r.c0) return '';
      const filas = [], columnas = c1 - r.c0 + 1;
      let celdas = 0;
      for (let f = r.f0; f <= f1; f++) {
        if (est.g.altoFila(f) === 0) continue;
        celdas += columnas; if (celdas > MAX_COPIA) break;
        const fila = [];
        for (let c = r.c0; c <= c1; c++) if (est.g.anchoCol(c) !== 0) fila.push(textoDeCelda(f, c).replace(/[\t\r\n]+/g, ' '));
        filas.push(fila.join('\t'));
      }
      return filas.join('\r\n');
    }
    async function copiar() {
      const texto = textoCopia();
      try { await navigator.clipboard.writeText(texto); return true; } catch (e) {
        const a = document.createElement('textarea'); a.value = texto; a.style.position = 'fixed'; a.style.opacity = '0'; document.body.appendChild(a); a.select();
        let ok = false; try { ok = document.execCommand('copy'); } catch (x) { ok = false; }
        a.remove(); cuerpo.focus({ preventScroll: true }); return ok;
      }
    }

    // ---- Zoom ------------------------------------------------------------------------------------------------------------
    function aplicarZoomControles() { controlZoom.value = est.zoom; valorZoom.textContent = est.zoom + '%'; }
    function ponerZoom(z, ancla) {
      if (!est) return;
      z = limitar(Math.round(z / 10) * 10, ZOOM_MIN, ZOOM_MAX);
      if (z === est.zoom) return;
      const antes = est.zoom / 100;
      // se mantiene fijo el punto de la hoja que está bajo el ancla (la esquina de las celdas, como Excel; con Ctrl + rueda o pellizco, el puntero)
      const ax = ancla ? ancla.x : vistaReal.areaX * antes, ay = ancla ? ancla.y : vistaReal.areaY * antes;
      const hojaX = (cuerpo.scrollLeft + ax) / antes, hojaY = (cuerpo.scrollTop + ay) / antes;
      est.zoom = z;
      colocarTamano();
      cuerpo.scrollLeft = hojaX * (z / 100) - ax; cuerpo.scrollTop = hojaY * (z / 100) - ay;
      aplicarZoomControles();
      pintar();
    }
    escuchar(controlZoom, 'input', () => ponerZoom(+controlZoom.value));
    escuchar(q('.rh-menos'), 'click', () => ponerZoom(est.zoom - 10));
    escuchar(q('.rh-mas'), 'click', () => ponerZoom(est.zoom + 10));
    escuchar(cuerpo, 'wheel', (e) => {
      if (!e.ctrlKey || !est) return;
      e.preventDefault();
      const r = vista.getBoundingClientRect();
      ponerZoom(est.zoom + (e.deltaY < 0 ? 10 : -10), { x: e.clientX - r.left, y: e.clientY - r.top });
    }, { passive: false });
    // pellizco en pantallas táctiles
    let distanciaInicial = 0, zoomInicial = 100;
    escuchar(vista, 'touchstart', (e) => { if (e.touches.length === 2) { distanciaInicial = Math.hypot(e.touches[0].clientX - e.touches[1].clientX, e.touches[0].clientY - e.touches[1].clientY); zoomInicial = est ? est.zoom : 100; toque = null; } }, { passive: true });
    escuchar(vista, 'touchmove', (e) => {
      if (e.touches.length !== 2 || !est || !distanciaInicial) return;
      e.preventDefault();
      const d = Math.hypot(e.touches[0].clientX - e.touches[1].clientX, e.touches[0].clientY - e.touches[1].clientY);
      const r = vista.getBoundingClientRect();
      ponerZoom(zoomInicial * d / distanciaInicial, { x: (e.touches[0].clientX + e.touches[1].clientX) / 2 - r.left, y: (e.touches[0].clientY + e.touches[1].clientY) / 2 - r.top });
    }, { passive: false });
    escuchar(vista, 'touchend', () => { distanciaInicial = 0; }, { passive: true });

    // ---- Pestañas ------------------------------------------------------------------------------------------------------------
    function dibujarPestanas() {
      pestanas.textContent = '';
      for (const h of lector.hojas) {
        if (h.estado !== 'visible') continue;
        const b = document.createElement('button');
        b.type = 'button'; b.className = 'rh-pestana'; b.setAttribute('role', 'tab'); b.dataset.indice = h.indice; b.textContent = h.nombre;
        b.setAttribute('aria-selected', 'false');
        pestanas.appendChild(b);
      }
    }
    const colorPestana = new Map();   // color de pestaña de las hojas ya abiertas (el de una hoja sin abrir no se conoce sin leerla entera)
    function marcarPestana() {
      if (est && est.hoja.color && !colorPestana.has(est.indice)) colorPestana.set(est.indice, libro.estilos.color(est.hoja.color));
      for (const b of pestanas.querySelectorAll('.rh-pestana')) {
        const indice = +b.dataset.indice, activa = !!est && indice === est.indice, color = colorPestana.get(indice);
        b.setAttribute('aria-selected', String(activa));
        if (color) b.style.setProperty('--rh-color', color); else b.style.removeProperty('--rh-color');
        if (activa) { if (b.offsetLeft < pestanas.scrollLeft) pestanas.scrollLeft = b.offsetLeft; else if (b.offsetLeft + b.offsetWidth > pestanas.scrollLeft + pestanas.clientWidth) pestanas.scrollLeft = b.offsetLeft + b.offsetWidth - pestanas.clientWidth; }   // solo la franja: scrollIntoView movería también la página
      }
    }
    escuchar(pestanas, 'click', async (e) => {
      const b = e.target.closest('.rh-pestana'); if (!b || (est && +b.dataset.indice === est.indice)) return;
      b.disabled = true;
      try { await abrirHoja(+b.dataset.indice); } finally { b.disabled = false; }
      cuerpo.focus({ preventScroll: true });
    });

    // ---- Arranque ---------------------------------------------------------------------------------------------------------
    escuchar(cuerpo, 'scroll', pedirPintado);
    // el cambio de tamaño se atiende fuera de la entrega del observador: cambiar el tamaño de la vista desde adentro provoca "ResizeObserver loop completed with undelivered notifications"
    let medidaPendiente = false;
    const observador = new ResizeObserver(() => {
      if (destruida || medidaPendiente) return;
      medidaPendiente = true;
      setTimeout(() => { medidaPendiente = false; if (destruida) return; medirVista(); pintar(); }, 0);
    });
    observador.observe(cuerpo);
    medirVista();
    dibujarPestanas();
    const primera = lector.hojas.find((h) => h.estado === 'visible') || lector.hojas[0];
    if (primera) await abrirHoja(primera.indice);

    return {
      mostrarHoja: abrirHoja,
      seleccion,
      copiar,
      // dónde está una celda en la vista (px, con el zoom y el desplazamiento de ahora): para las pruebas y para quien dibuje encima
      rectCelda(f, c) { const x = pantallaX(c), y = pantallaY(f); return { x, y, ancho: pantallaX(c + 1) - x, alto: pantallaY(f + 1) - y }; },
      zoom(z) { if (z === undefined) return est ? est.zoom : 100; ponerZoom(z); return est.zoom; },
      alCambiar(fn) { cambios.push(fn); },
      enfocar() { cuerpo.focus({ preventScroll: true }); },
      destruir() { destruida = true; observador.disconnect(); for (const quitar of escuchas) quitar(); clearTimeout(temporizadorEstado); contenedor.textContent = ''; },
    };
  }

  globalThis.RejillaHoja = { crear };
})();

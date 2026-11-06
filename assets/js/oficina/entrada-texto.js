/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Entrada de texto sobre el motor de texto (Fase 4, 4.3 a 4.5): edita el texto de una
// forma de Presentaciones mostrándolo con su formato y su corte de línea reales.
//
// El teclado, el IME y el portapapeles los sigue recibiendo un <textarea> nativo
// (transparente, junto al cursor): así el teclado del celular, la composición de
// japonés o chino y el pegado funcionan solos. Lo que se ve es un lienzo encima que
// dibuja el texto con MotorTexto, la selección y el cursor; el ratón y el dedo
// actúan sobre ese lienzo (posición ↔ punto) y las flechas arriba/abajo, Inicio y
// Fin siguen las líneas del motor, no las del cuadro nativo.
//
// El textarea sigue siendo la fuente de la selección y del texto (value,
// selectionStart/End): lo demás de la app (formato, guardar) lo lee de ahí.
//
// Necesita cargados antes: pptx-texto.js (PptxTexto), pptx-geometria.js (PptxGeometria),
// motor-texto.js y edicion-texto.js.
//
// Uso: const ent = EntradaTexto.crear({ marco, area, elemento, escala });
//      ent.pintar(); ent.destruir();
//      EntradaTexto.admite(elemento)  → false con texto vertical, con el texto girado dentro de la forma, espejado o de derecha a izquierda
//      (una forma girada sí se admite: el lienzo gira con ella y el clic se gira al revés)

const EntradaTexto = (() => {
  const LETRA = /[\p{L}\p{N}_]/u;
  const RTL = /[֐-ࣿיִ-﷿ﹰ-﻿]|\uD802[\uDC00-\uDFFF]|\uD83A[\uDC00-\uDFFF]/;

  function admite(e) {
    if (!e || !e.texto || !e.xfrm) return false;
    const c = e.texto.cuerpo;
    if (c.vertical && c.vertical !== 'horz') return false;
    if (c.giro || e.cajaTexto || e.xfrm.flipH || e.xfrm.flipV) return false;
    // de derecha a izquierda (hebreo, árabe…): el motor pone las piezas de izquierda a derecha y mide prefijos, así que el cursor y el clic
    // quedarían espejados dentro de cada palabra; el cuadro nativo sí sabe de bidi
    if (e.texto.parrafos.some((p) => p.rtl || p.trozos.some((t) => RTL.test(t.texto || '')))) return false;
    return true;
  }
  const luminancia = (css) => {
    const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(css || '');
    return m ? (0.2126 * m[1] + 0.7152 * m[2] + 0.0722 * m[3]) / 255 : 0;
  };

  function crear({ marco, area, elemento, escala }) {
    const s = escala;
    const r = elemento.xfrm;
    const cuerpo = elemento.texto.cuerpo;
    const w = Math.max(30, r.w * s), h = Math.max(20, r.h * s);
    const dpr = Math.max(1, Math.min(2, window.devicePixelRatio || 1));
    let vivos = elemento.texto.parrafos;
    if (area.value !== EdicionTexto.plano(vivos)) vivos = EdicionTexto.aplicar(vivos, area.value);
    const lienzo = document.createElement('canvas');
    lienzo.className = 'pr-editor-lienzo';
    lienzo.setAttribute('aria-hidden', 'true');
    const ctx = lienzo.getContext('2d');
    const motor = MotorTexto.crear(ctx);
    marco.append(lienzo);

    // el textarea: sigue en el DOM para recibir teclas e IME, pero no se ve
    Object.assign(area.style, {
      background: 'transparent', border: '0', padding: '0', color: 'transparent', caretColor: 'transparent', pointerEvents: 'none',
      whiteSpace: 'pre', overflow: 'hidden', transform: 'none', outline: 'none', resize: 'none', opacity: '1', width: '4px', height: '16px', zIndex: '6',
    });
    area.classList.add('pr-editor-oculto');

    // caja del texto: la del texto de la geometría menos los márgenes (igual que el dibujo)
    let rect = { l: 0, t: 0, r: w, b: h };
    try {
      const g = elemento.geometria && elemento.geometria.def ? PptxGeometria.trazar(elemento.geometria.def, w, h, elemento.geometria.ajustes) : null;
      if (g && g.texto) rect = g.texto;
    } catch (e) { /* sin geometría: la caja entera */ }
    const mg = cuerpo.margenes;
    const areaTexto = { x: rect.l + mg.l * s, y: rect.t + mg.t * s, w: Math.max(0, rect.r - rect.l - (mg.l + mg.r) * s), h: Math.max(0, rect.b - rect.t - (mg.t + mg.b) * s) };

    let m = null, y0 = 0, visible = true, objetivoX = null, lineaFija = null, composicion = null, vivo = true;
    let temporizador = null, altoLienzo = h;
    // El lienzo gira con la forma (rotate alrededor de su centro): un punto del lienzo sin girar ↔ el mismo punto girado
    const angulo = ((r.rot || 0) * Math.PI) / 180;
    const girar = (x, y, a) => { const dx = x - w / 2, dy = y - altoLienzo / 2; return { x: w / 2 + dx * Math.cos(a) - dy * Math.sin(a), y: altoLienzo / 2 + dx * Math.sin(a) + dy * Math.cos(a) }; };
    const centroEnPantalla = () => { const b = lienzo.getBoundingClientRect(); return { x: b.left + b.width / 2 - w / 2, y: b.top + b.height / 2 - altoLienzo / 2 }; };

    // posiciones globales (en el texto de la forma) ↔ (párrafo, desplazamiento)
    const inicios = () => { const a = []; let o = 0; for (const p of vivos) { a.push(o); o += EdicionTexto.textoParrafo(p).length + 1; } return a; };
    const aGlobal = (ip, d) => inicios()[ip] + d;
    const aParrafo = (g) => { const a = inicios(); let ip = a.length - 1; while (ip > 0 && a[ip] > g) ip--; return { ip, d: Math.min(g - a[ip], EdicionTexto.textoParrafo(vivos[ip]).length) }; };
    const foco = () => (area.selectionDirection === 'backward' ? area.selectionStart : area.selectionEnd);
    const ancla = () => (area.selectionDirection === 'backward' ? area.selectionEnd : area.selectionStart);
    function poner(anclaG, focoG) {
      const a = Math.max(0, Math.min(anclaG, area.value.length)), f = Math.max(0, Math.min(focoG, area.value.length));
      area.setSelectionRange(Math.min(a, f), Math.max(a, f), f < a ? 'backward' : 'forward');
    }

    function pintar() {
      if (!vivo) return;
      m = motor.maquetar({ cuerpo, parrafos: vivos }, areaTexto.w, areaTexto.h, s);
      const alto = Math.max(h, areaTexto.y + m.alto + mg.b * s);
      altoLienzo = alto;
      y0 = areaTexto.y;
      if (cuerpo.ancla === 'ctr') y0 += (areaTexto.h - m.alto) / 2; else if (cuerpo.ancla === 'b') y0 += areaTexto.h - m.alto;
      Object.assign(lienzo.style, { left: r.x * s + 'px', top: r.y * s + 'px', width: w + 'px', height: alto + 'px', transform: 'rotate(' + (r.rot || 0) + 'deg)' });
      lienzo.width = Math.round(w * dpr); lienzo.height = Math.round(alto * dpr);
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      // fondo: claro, u oscuro si el texto es claro
      const lum = []; for (const l of m.lineas) for (const t of l.trozos) if (t.color) lum.push(luminancia(t.color));
      const claro = lum.length && lum.reduce((a, b) => a + b, 0) / lum.length > 0.7;
      ctx.fillStyle = claro ? 'rgba(40,40,40,0.96)' : 'rgba(255,255,255,0.96)';
      ctx.fillRect(0, 0, w, alto);
      ctx.strokeStyle = '#1a73e8'; ctx.lineWidth = 2; ctx.strokeRect(1, 1, w - 2, alto - 2);
      const porDefecto = claro ? '#fff' : '#000';

      // selección y composición del IME
      const a = Math.min(area.selectionStart, area.selectionEnd), f = Math.max(area.selectionStart, area.selectionEnd);
      const ini = inicios();
      if (f > a) {
        ctx.fillStyle = 'rgba(26,115,232,0.35)';
        m.lineas.forEach((l, i) => {
          const base0 = ini[l.parrafoIndice], g0 = base0 + l.desde, g1 = base0 + l.hasta;
          const desde = Math.max(a, g0), hasta = Math.min(f, g1);
          // índice global del carácter que cierra la línea (fin de párrafo o salto suave), si lo hay
          const corte = l.saltoHasta != null ? base0 + l.saltoHasta - 1 : (l.final ? g1 : -1);
          const cubreCorte = corte >= 0 && corte < area.value.length + (l.saltoHasta != null ? 1 : 0) && a <= corte && f > corte;
          if (hasta < desde && !cubreCorte) return;
          const d0 = Math.min(Math.max(desde, g0), g1);
          const x1 = motor.xEnLinea(m, i, d0 - base0);
          let x2 = hasta > desde ? motor.xEnLinea(m, i, hasta - base0) : x1;
          if (cubreCorte) x2 += 5;
          if (x2 > x1) ctx.fillRect(areaTexto.x + x1, y0 + l.arriba, x2 - x1, l.alto);
        });
      }
      // texto
      ctx.textBaseline = 'alphabetic'; ctx.textAlign = 'left';
      for (const l of m.lineas) {
        const base = y0 + l.base;
        if (l.vineta && l.vineta.color) { ctx.font = l.vineta.css; ctx.letterSpacing = '0px'; ctx.fillStyle = l.vineta.color; ctx.fillText(l.vineta.texto, areaTexto.x + l.vineta.x, base); }
        for (const t of l.trozos) {
          const x = areaTexto.x + t.x, y = base + t.desplazamiento;
          if (t.resaltado) { ctx.fillStyle = t.resaltado; ctx.fillRect(x, y - t.px * 0.95, t.ancho, t.px * 1.2); }
          ctx.font = t.css; ctx.letterSpacing = t.espaciado ? `${t.espaciado}px` : '0px'; ctx.fontKerning = t.kerning || 'normal';
          ctx.fillStyle = t.color || porDefecto;
          ctx.fillText(t.texto, x, y);
          const grosor = Math.max(1, t.px * 0.06);
          if (t.subrayado) ctx.fillRect(x, y + t.px * 0.12, t.ancho, grosor);
          if (t.tachado) ctx.fillRect(x, y - t.px * 0.3, t.ancho, grosor);
        }
      }
      ctx.letterSpacing = '0px';
      // la parte que se está componiendo con el IME, subrayada
      if (composicion && composicion.fin > composicion.ini) {
        ctx.strokeStyle = porDefecto; ctx.lineWidth = 1.5;
        m.lineas.forEach((l, i) => {
          const g0 = ini[l.parrafoIndice] + l.desde, g1 = ini[l.parrafoIndice] + l.hasta;
          const desde = Math.max(composicion.ini, g0), hasta = Math.min(composicion.fin, g1);
          if (hasta <= desde) return;
          const x1 = motor.xEnLinea(m, i, desde - ini[l.parrafoIndice]), x2 = motor.xEnLinea(m, i, hasta - ini[l.parrafoIndice]);
          ctx.beginPath(); ctx.moveTo(areaTexto.x + x1, y0 + l.base + 3); ctx.lineTo(areaTexto.x + x2, y0 + l.base + 3); ctx.stroke();
        });
      }
      // cursor
      const c = puntoDeFoco();
      if (c) {
        if (f === a && visible && document.activeElement === area) { ctx.fillStyle = porDefecto; ctx.fillRect(areaTexto.x + c.x - 0.75, y0 + c.arriba, 1.5, c.alto); }
        // el cuadro nativo junto al cursor: ahí aparece la ventana del IME
        const q = girar(areaTexto.x + c.x, y0 + c.arriba, angulo);
        Object.assign(area.style, { left: r.x * s + q.x + 'px', top: r.y * s + q.y + 'px', height: Math.max(8, c.alto) + 'px', fontSize: Math.max(8, c.alto / 1.2) + 'px' });
      }
    }
    function puntoDeFoco() {
      if (!m) return null;
      const { ip, d } = aParrafo(foco());
      if (lineaFija != null && lineaFija.g === foco() && m.lineas[lineaFija.linea]) {
        const l = m.lineas[lineaFija.linea];
        return { x: motor.xEnLinea(m, lineaFija.linea, d), arriba: l.arriba, base: l.base, alto: l.alto, linea: lineaFija.linea };
      }
      return motor.puntoDe(m, ip, d);
    }
    const repintar = () => { visible = true; pintar(); };

    // ---- el texto cambió (teclas, pegado, IME): el modelo vivo sigue al cuadro nativo ----------------
    const alEscribir = () => {
      if (area.value !== EdicionTexto.plano(vivos)) { vivos = EdicionTexto.aplicar(vivos, area.value); }
      objetivoX = null; lineaFija = null;
      repintar();
    };
    const alSeleccionar = () => { repintar(); };
    area.addEventListener('input', alEscribir);
    for (const n of ['select', 'keyup', 'focus', 'blur']) area.addEventListener(n, alSeleccionar);
    const alSelDoc = () => { if (document.activeElement === area) pintar(); };
    document.addEventListener('selectionchange', alSelDoc);
    area.addEventListener('compositionstart', () => { composicion = { ini: area.selectionStart, fin: area.selectionStart }; });
    // lo compuesto es ev.data entero: el cursor del IME puede estar en medio (al elegir una parte), no vale selectionEnd
    area.addEventListener('compositionupdate', (ev) => { if (composicion) { composicion.fin = composicion.ini + (ev.data || '').length; repintar(); } });
    area.addEventListener('compositionend', () => { composicion = null; repintar(); });

    // ---- teclas que siguen las líneas del motor ----------------------------------------------------------
    area.addEventListener('keydown', (ev) => {
      const tecla = ev.key;
      // cualquier otra tecla que mueve o escribe (← →, Ctrl+…) olvida la columna de las flechas y el final de línea fijado
      if (!['ArrowUp', 'ArrowDown', 'Shift', 'Control', 'Alt', 'Meta'].includes(tecla)) { objetivoX = null; if (!['Home', 'End'].includes(tecla)) lineaFija = null; }
      if (ev.isComposing || ev.altKey || ev.ctrlKey || ev.metaKey) return;
      if (!['ArrowUp', 'ArrowDown', 'Home', 'End'].includes(tecla) || !m) return;
      const c = puntoDeFoco();
      if (!c) return;
      let destino = null;
      if (tecla === 'ArrowUp' || tecla === 'ArrowDown') {
        // la columna de las flechas se guarda relativa a la columna de texto (con columnas, la línea siguiente puede estar en otra)
        const xCol = (i) => (m.lineas[i].columna || 0) * (m.anchoCol + m.espacioCol);
        if (objetivoX == null) objetivoX = c.x - xCol(c.linea);
        const k = c.linea + (tecla === 'ArrowUp' ? -1 : 1);
        if (k < 0) destino = 0;
        else if (k >= m.lineas.length) destino = area.value.length;
        else {
          const l = m.lineas[k];
          const p = motor.posicionDe(m, objetivoX + xCol(k), l.arriba + l.alto / 2, k);
          destino = aGlobal(p.parrafo, p.desplazamiento);
          lineaFija = { g: destino, linea: k };
        }
      } else {
        const l = m.lineas[c.linea], ini = inicios()[l.parrafoIndice];
        if (tecla === 'Home') destino = ini + l.desde;
        else {
          let fin = l.hasta;
          if (l.saltoHasta != null) fin = l.saltoHasta - 1;
          else if (!l.final) { for (let i = l.piezas.length - 1; i >= 0; i--) if (!l.piezas[i].espacio) { fin = l.piezas[i].hasta; break; } }
          destino = ini + fin;
          if (!l.final || l.saltoHasta != null) lineaFija = { g: destino, linea: c.linea };
        }
        objetivoX = null;
      }
      ev.preventDefault();
      const guardada = lineaFija;
      poner(ev.shiftKey ? ancla() : destino, destino);
      lineaFija = guardada && guardada.g === destino ? guardada : null;
      area.dispatchEvent(new Event('select'));
      repintar();
    });

    // ---- ratón y dedo sobre el lienzo ------------------------------------------------------------------------
    let arrastre = null;
    const globalDe = (ev) => {
      // del punto de la pantalla al lienzo sin girar (en una forma girada, el clic se gira al revés)
      const o = centroEnPantalla(), q = girar(ev.clientX - o.x, ev.clientY - o.y, -angulo);
      const p = motor.posicionDe(m, q.x - areaTexto.x, q.y - y0);
      return p ? aGlobal(p.parrafo, p.desplazamiento) : 0;
    };
    const palabraEn = (g) => {
      const v = area.value; let a = g, f = g;
      while (a > 0 && LETRA.test(v[a - 1])) a--;
      while (f < v.length && LETRA.test(v[f])) f++;
      return [a, f];
    };
    // pointerdown no trae la cuenta de clics (detail vale 0): se cuenta aquí, también para el dedo
    let ultimo = { t: 0, x: 0, y: 0, n: 0 };
    lienzo.addEventListener('pointerdown', (ev) => {
      if (ev.button > 0 || !m) return;
      const cerca = ev.timeStamp - ultimo.t < 450 && Math.abs(ev.clientX - ultimo.x) < 6 && Math.abs(ev.clientY - ultimo.y) < 6;
      ultimo = { t: ev.timeStamp, x: ev.clientX, y: ev.clientY, n: cerca ? ultimo.n + 1 : 1 };
      ev.preventDefault(); ev.stopPropagation();
      area.focus({ preventScroll: true });
      const g = globalDe(ev), v = area.value;
      objetivoX = null; lineaFija = null;
      if (ultimo.n >= 3) { const a = g > 0 ? v.lastIndexOf('\n', g - 1) + 1 : 0; let f = v.indexOf('\n', g); if (f < 0) f = v.length; poner(a, f); arrastre = null; }
      else if (ultimo.n === 2) { const [a, f] = palabraEn(g); poner(a, f); arrastre = null; }
      else { poner(ev.shiftKey ? ancla() : g, g); arrastre = { ancla: ev.shiftKey ? ancla() : g }; try { lienzo.setPointerCapture(ev.pointerId); } catch (e) { /* sin captura */ } }
      area.dispatchEvent(new Event('select'));
      repintar();
    });
    lienzo.addEventListener('pointermove', (ev) => { if (!arrastre || !m) return; const g = globalDe(ev); poner(arrastre.ancla, g); repintar(); });
    const soltar = () => { if (arrastre) { arrastre = null; area.dispatchEvent(new Event('select')); } };
    lienzo.addEventListener('pointerup', soltar); lienzo.addEventListener('pointercancel', soltar);
    lienzo.addEventListener('dblclick', (ev) => { ev.stopPropagation(); });

    // cursor parpadeante
    temporizador = setInterval(() => { if (document.activeElement === area && area.selectionStart === area.selectionEnd) { visible = !visible; pintar(); } }, 530);

    function destruir() {
      vivo = false;
      clearInterval(temporizador);
      document.removeEventListener('selectionchange', alSelDoc);
      lienzo.remove();
    }
    pintar();
    // dónde cae en la pantalla (coordenadas del navegador) el desplazamiento global g: para las pruebas y para colocar menús junto al cursor
    function punto(g) {
      const { ip, d } = aParrafo(g);
      const pt = motor.puntoDe(m, ip, d);
      if (!pt) return null;
      const o = centroEnPantalla(), q = girar(areaTexto.x + pt.x, y0 + pt.arriba + pt.alto / 2, angulo);
      return { x: o.x + q.x, y: o.y + q.y, linea: pt.linea };
    }
    return { pintar: repintar, destruir, motor, modelo: () => vivos, maquetado: () => m, punto, lienzo, composicion: () => composicion && { ...composicion } };
  }

  return { crear, admite };
})();

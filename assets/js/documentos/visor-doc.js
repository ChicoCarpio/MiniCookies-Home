/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Visor de hojas de Documentos (Fase 6, etapa 6.3): apila las hojas del diseño de MaquetadorDoc en un contenedor con desplazamiento, con zoom,
// y solo dibuja (en Canvas, con la densidad de la pantalla) las hojas que se ven más una pantalla de margen: las demás son un recuadro blanco
// del tamaño de la hoja, sin canvas, así que un documento de cientos de hojas gasta la memoria de unas pocas.
//
//   const visor = VisorDoc.crear(contenedor, { diseno, pintor, zoom: 1 });
//   visor.zoom(1.5) · visor.irAHoja(3) · visor.hojaActual() · visor.actualizarDiseno(nuevo) · visor.repintar() · visor.destruir()
//   visor.hojas → los elementos de cada hoja (div.vd-hoja, con su canvas si está a la vista)
(function () {
  'use strict';
  const PX_POR_PT = 96 / 72;
  const MAX_PIXELES = 24e6;      // un canvas más grande lo rechazan algunos navegadores (y el celular se queda sin memoria)

  function crear(contenedor, opciones) {
    const doc = contenedor.ownerDocument;
    let diseno = opciones.diseno, zoom = opciones.zoom || 1;
    const pintor = opciones.pintor;
    const hueco = opciones.hueco === undefined ? 16 : opciones.hueco;
    const pila = doc.createElement('div');
    pila.className = 'vd-pila';
    pila.style.cssText = `display:flex;flex-direction:column;align-items:center;gap:${hueco}px;padding:${hueco}px;min-width:max-content;box-sizing:border-box`;
    contenedor.appendChild(pila);
    let hojas = [], tops = [], pendiente = false, vivo = true, sobre = opciones.sobre || null;   // sobre(ctx, hoja, n): lo que se dibuja encima (selección…)
    const conCanvas = new Set();
    const tiempos = [];   // ms de cada pintado, para medir

    function construir() {
      pila.textContent = '';
      conCanvas.clear();
      hojas = diseno.paginas.map((h, i) => {
        const d = doc.createElement('div');
        d.className = 'vd-hoja';
        d.dataset.hoja = String(i + 1);
        d.style.cssText = 'position:relative;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.35);flex:none';
        pila.appendChild(d);
        return d;
      });
      medir();
    }
    function medir() {
      let y = hueco;
      tops = [];
      diseno.paginas.forEach((h, i) => {
        const d = hojas[i];
        d.style.width = h.ancho * PX_POR_PT * zoom + 'px';
        d.style.height = h.alto * PX_POR_PT * zoom + 'px';
        tops.push(y);
        y += h.alto * PX_POR_PT * zoom + hueco;
      });
      for (const n of [...conCanvas]) { quitarCanvas(n); }
      programar();
    }

    function pintarHoja(n) {
      const h = diseno.paginas[n], d = hojas[n];
      let canvas = d.querySelector('canvas');
      const dpr = Math.min(opciones.dpr || (typeof devicePixelRatio === 'number' ? devicePixelRatio : 1), 3);
      const wCss = h.ancho * PX_POR_PT * zoom, hCss = h.alto * PX_POR_PT * zoom;
      const dprEf = Math.min(dpr, Math.sqrt(MAX_PIXELES / (wCss * hCss)));
      if (!canvas) { canvas = doc.createElement('canvas'); canvas.style.cssText = 'display:block;width:100%;height:100%'; d.appendChild(canvas); }
      const w = Math.max(Math.round(wCss * dprEf), 1), alto = Math.max(Math.round(hCss * dprEf), 1);
      if (canvas.width !== w || canvas.height !== alto) { canvas.width = w; canvas.height = alto; }
      const ctx = canvas.getContext('2d');
      const t0 = performance.now();
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      ctx.clearRect(0, 0, w, alto);
      const k = PX_POR_PT * zoom * dprEf;
      ctx.setTransform(k, 0, 0, k, 0, 0);
      pintor.pintarHoja(ctx, h);
      if (sobre) sobre(ctx, h, n);
      tiempos.push(performance.now() - t0);
      if (tiempos.length > 500) tiempos.shift();
      conCanvas.add(n);
    }
    function quitarCanvas(n) {
      const c = hojas[n] && hojas[n].querySelector('canvas');
      if (c) { c.width = 0; c.height = 0; c.remove(); }
      conCanvas.delete(n);
    }

    function rango() {
      const vista = contenedor.clientHeight || 800, y0 = contenedor.scrollTop - pila.offsetTop, margen = vista;
      let desde = 0, hasta = hojas.length - 1;
      const sup = y0 - margen, inf = y0 + vista + margen;
      // búsqueda binaria de la primera hoja cuyo borde de abajo pasa de `sup`
      let lo = 0, hi = hojas.length - 1;
      while (lo < hi) { const mid = (lo + hi) >> 1; if (tops[mid] + diseno.paginas[mid].alto * PX_POR_PT * zoom < sup) lo = mid + 1; else hi = mid; }
      desde = lo;
      hasta = desde;
      while (hasta + 1 < hojas.length && tops[hasta + 1] <= inf) hasta++;
      return [desde, hasta];
    }
    function actualizarVisibles() {
      pendiente = false;
      if (!vivo || !hojas.length) return;
      const [desde, hasta] = rango();
      for (const n of [...conCanvas]) if (n < desde || n > hasta) quitarCanvas(n);
      for (let n = desde; n <= hasta; n++) if (!conCanvas.has(n)) pintarHoja(n);
      if (opciones.alCambiarHoja) opciones.alCambiarHoja(hojaActual(), hojas.length);
    }
    function programar() {
      if (pendiente || !vivo) return;
      pendiente = true;
      (typeof requestAnimationFrame === 'function' ? requestAnimationFrame : (f) => setTimeout(f, 16))(actualizarVisibles);
    }
    function hojaActual() {
      const y = contenedor.scrollTop - pila.offsetTop + (contenedor.clientHeight || 800) / 3;
      let n = 0;
      for (let i = 0; i < tops.length; i++) if (tops[i] <= y) n = i;
      return n + 1;
    }
    contenedor.addEventListener('scroll', programar, { passive: true });
    if (typeof ResizeObserver !== 'undefined') { const ro = new ResizeObserver(programar); ro.observe(contenedor); }
    if (pintor) pintor.alCargarImagen = () => repintar();

    function repintar() { for (const n of [...conCanvas]) pintarHoja(n); }
    // ¿Dibuja lo mismo la hoja nueva que la vieja? Las líneas sin cambios son los mismos objetos (el motor los guarda), así que basta compararlos.
    const mismaLista = (a, b) => {
      if (a === b) return true;
      if (!a || !b || a.length !== b.length) return false;
      for (let i = 0; i < a.length; i++) {
        const p = a[i], q = b[i];
        if (p.t !== q.t || p.b !== q.b || p.x !== q.x || p.y !== q.y) return false;
        if (p.t === 'linea') { if (p.l !== q.l) return false; }
        else if (p.ancho !== q.ancho || p.alto !== q.alto || p.filas.length !== q.filas.length || p.filas.some((f, k) => f.y !== q.filas[k].y || f.alto !== q.filas[k].alto)) return false;
      }
      return true;
    };
    const mismaHoja = (a, b) => a.ancho === b.ancho && a.alto === b.alto && a.numero === b.numero && mismaLista(a.items, b.items)
      && mismaLista(a.encabezado && a.encabezado.items, b.encabezado && b.encabezado.items) && mismaLista(a.pie && a.pie.items, b.pie && b.pie.items)
      && a.flot.length === b.flot.length && a.flot.every((f, k) => f.x === b.flot[k].x && f.y === b.flot[k].y && f.pieza === b.flot[k].pieza);
    construir();
    actualizarVisibles();
    return {
      get hojas() { return hojas; },
      get tiempos() { return tiempos; },
      get canvasVivos() { return conCanvas.size; },
      zoom(z) {
        if (z === undefined) return zoom;
        const hoja = hojaActual(), fraccion = (contenedor.scrollTop - pila.offsetTop - (tops[hoja - 1] || 0)) / Math.max(diseno.paginas[hoja - 1].alto * PX_POR_PT * zoom, 1);
        zoom = Math.min(Math.max(z, 0.25), 4);
        medir();
        contenedor.scrollTop = pila.offsetTop + (tops[hoja - 1] || 0) + fraccion * diseno.paginas[hoja - 1].alto * PX_POR_PT * zoom;
        actualizarVisibles();
        return zoom;
      },
      irAHoja(n) { const i = Math.min(Math.max(n, 1), hojas.length) - 1; contenedor.scrollTop = pila.offsetTop + tops[i] - hueco; actualizarVisibles(); },
      hojaActual,
      repintar,
      // Un diseño nuevo (tras editar). Con las mismas hojas del mismo tamaño solo se vuelven a dibujar las que se ven y cambiaron.
      actualizarDiseno(nuevo) {
        const viejo = diseno;
        diseno = nuevo;
        const igualTam = viejo.paginas.length === nuevo.paginas.length && nuevo.paginas.every((h, i) => h.ancho === viejo.paginas[i].ancho && h.alto === viejo.paginas[i].alto);
        if (!igualTam) { construir(); actualizarVisibles(); return { reconstruido: true, repintadas: conCanvas.size }; }
        let repintadas = 0;
        for (const n of [...conCanvas]) if (!mismaHoja(viejo.paginas[n], nuevo.paginas[n])) { pintarHoja(n); repintadas++; }
        if (opciones.alCambiarHoja) opciones.alCambiarHoja(hojaActual(), hojas.length);
        return { reconstruido: false, repintadas };
      },
      get escala() { return PX_POR_PT * zoom; },
      ponerSobre(fn) { sobre = fn; },
      // Vuelve a dibujar la hoja n (0 a …) si está a la vista
      repintarHoja(n) { if (conCanvas.has(n)) pintarHoja(n); },
      estaViva: (n) => conCanvas.has(n),
      get pila() { return pila; },
      destruir() { vivo = false; pila.remove(); },
    };
  }

  globalThis.VisorDoc = { crear };
})();

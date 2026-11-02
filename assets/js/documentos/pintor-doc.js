/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Pintor de hojas de Word en Canvas (Fase 6, etapa 6.3). Dibuja una hoja del resultado de MaquetadorDoc (maquetador-doc.js): fondo y bordes de
// párrafo, sombreado y bordes de tablas, texto (fuente, color, negrita, cursiva, subrayado, tachado, resaltado, superíndices, mayúsculas),
// viñetas y números de lista, guías de tabulación, imágenes en línea y flotantes, encabezado, pie y notas al pie. Sin DOM salvo el canvas:
// el llamador pone la escala en el contexto (ctx.setTransform) para que todo se dibuje en puntos (pt), con el origen en la esquina de la hoja.
//
//   const pintor = PintorDoc.crear(modelo);
//   await pintor.precargar(diseno);                 // decodifica las imágenes del documento (una vez)
//   pintor.pintarHoja(ctx, diseno.paginas[0]);      // ctx con setTransform(pxPorPt·zoom·dpr, 0, 0, …, 0, 0)
//   pintor.imagenesPendientes → cuántas faltan por decodificar (el visor repinta cuando llegan)
// Lo que no dibuja todavía: texto dentro de cuadros de texto, formas, imágenes EMF/WMF (un recuadro gris), recortes y giros de imagen.
(function () {
  'use strict';
  const RESALTADOS = {
    yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000', darkBlue: '#000080',
    darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080',
    lightGray: '#c0c0c0', black: '#000000', white: '#ffffff',
  };
  const GUIAS = { dot: '.', hyphen: '-', underscore: '_', heavy: '_', middleDot: '·' };
  const TIPOS_IMAGEN = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp' };
  const hex6 = (c) => (typeof c === 'string' && /^[0-9a-fA-F]{6}$/.test(c) ? '#' + c.toLowerCase() : null);

  function crear(modelo, opciones = {}) {
    const m = modelo.m;
    const imagenes = new Map();   // destino → ImageBitmap | null (no se pudo) | 'pendiente'
    const api = { imagenesPendientes: 0, alCargarImagen: null };

    // ───────── Imágenes ─────────
    function cargarImagen(destino, zip) {
      if (!destino || imagenes.has(destino)) return imagenes.get(destino);
      const ext = (destino.split('.').pop() || '').toLowerCase();
      const f = zip.file(destino);
      if (!f || !TIPOS_IMAGEN[ext] || typeof createImageBitmap === 'undefined') { imagenes.set(destino, null); return null; }
      imagenes.set(destino, 'pendiente');
      api.imagenesPendientes++;
      f.async('uint8array').then(async (bytes) => {
        const bmp = await createImageBitmap(new Blob([bytes], { type: TIPOS_IMAGEN[ext] }));
        // el PDF copia los bytes del JPEG tal cual en vez de recodificar la foto como píxeles (13 veces más grande); el lienzo del PDF decide si sirven (sin giro EXIF, sRGB…)
        if (TIPOS_IMAGEN[ext] === 'image/jpeg') bmp.bytesJpeg = bytes;
        return bmp;
      })
        .then((bmp) => { imagenes.set(destino, bmp); }, () => { imagenes.set(destino, null); })
        .then(() => { api.imagenesPendientes--; if (api.alCargarImagen) api.alCargarImagen(destino); });
      return 'pendiente';
    }
    // Decodifica todas las imágenes que aparecen en el diseño (cuerpo, encabezados, pies y notas).
    async function precargar(diseno) {
      const pendientes = [];
      const ver = (items) => { for (const it of items) if (it.t === 'linea') for (const a of it.l.atomos) if (a.pieza && a.pieza.destino) pendientes.push(a.pieza.destino); };
      for (const h of diseno.paginas) {
        ver(h.items);
        if (h.encabezado) ver(h.encabezado.items);
        if (h.pie) ver(h.pie.items);
        for (const q of h.flot || []) if (q.pieza && q.pieza.destino) pendientes.push(q.pieza.destino);
      }
      for (const d of new Set(pendientes)) cargarImagen(d, modelo.zip);
      while (api.imagenesPendientes > 0) await new Promise((r) => setTimeout(r, 5));
    }

    // ───────── Texto ─────────
    const colorDe = (r) => m.color(r.color) || '#000000';
    function fuente(a) { return a.f.css.replace(/[0-9.]+px/, a.f.tam.toFixed(3) + 'px'); }   // el contexto trabaja en pt: el tamaño de la letra es el tamaño en pt

    function subrayado(ctx, a, x, y, ancho) {
      const r = a.r, u = r.u;
      const t = Math.max(a.f.tam * 0.05, 0.4);
      ctx.save();
      ctx.strokeStyle = colorDe(r); ctx.lineWidth = /thick|heavy/i.test(u || '') ? t * 2 : t;
      if (/dotted/i.test(u)) ctx.setLineDash([t, t * 1.5]); else if (/dash/i.test(u)) ctx.setLineDash([t * 4, t * 2]);
      const yy = y + a.f.tam * 0.12;
      ctx.beginPath();
      if (/wave/i.test(u)) { const p = t * 3; ctx.moveTo(x, yy); for (let k = 0; k * p < ancho; k++) ctx.lineTo(Math.min(x + (k + 1) * p, x + ancho), yy + (k % 2 ? 0 : t * 1.5)); }
      else { ctx.moveTo(x, yy); ctx.lineTo(x + ancho, yy); if (/double/i.test(u)) { ctx.moveTo(x, yy + t * 2); ctx.lineTo(x + ancho, yy + t * 2); } }
      ctx.stroke(); ctx.restore();
    }

    function lineaDeTexto(ctx, it) {
      const l = it.l;
      for (const a of l.atomos) {
        if (a.k === 'f' || a.k === 'b' || a.k === 'pg') continue;
        const x = it.x + a.x, base = it.y + l.base - (a.f ? a.f.subida : 0);
        const r = a.r || {};
        // resaltado y sombreado de carácter detrás del texto
        if (a.k !== 'i') {
          const fondo = (r.resaltado && RESALTADOS[r.resaltado]) || (r.sombreado && hex6(r.sombreado.relleno));
          if (fondo && a.ancho > 0) { ctx.fillStyle = fondo; ctx.fillRect(x, it.y, a.ancho, l.alto); }
        }
        if (a.k === 'i') {
          const img = a.pieza && a.pieza.destino ? cargarImagen(a.pieza.destino, modelo.zip) : null;
          const y = base - a.alto;
          if (img && img !== 'pendiente') ctx.drawImage(img, x, y, a.ancho, a.alto);
          else { ctx.save(); ctx.fillStyle = '#eef0f2'; ctx.fillRect(x, y, a.ancho, a.alto); ctx.strokeStyle = '#b8bec4'; ctx.lineWidth = 0.5; ctx.strokeRect(x, y, a.ancho, a.alto); ctx.restore(); }
          continue;
        }
        if (a.k === 't') {
          const car = a.guia && GUIAS[a.guia];
          if (car && a.ancho > 1) {
            ctx.font = fuente(a); ctx.fillStyle = colorDe(r); ctx.letterSpacing = '0px';
            const w = ctx.measureText(car).width || 1;
            const n = Math.floor(a.ancho / w);
            if (n > 0) ctx.fillText(car.repeat(n), x + (a.ancho - n * w), base);
          }
          continue;
        }
        if (a.k === 'e' || !a.texto) {
          if (r.u && r.u !== 'none' && a.ancho > 0 && a.k === 'e') subrayado(ctx, a, x, base, a.ancho);
          continue;
        }
        ctx.font = fuente(a); ctx.direction = (a.nivel & 1) ? 'rtl' : 'ltr';
        ctx.fontKerning = 'auto';
        ctx.letterSpacing = a.f.esp ? a.f.esp + 'px' : '0px';
        ctx.fillStyle = colorDe(r);
        if (a.f.escala !== 1) { ctx.save(); ctx.translate(x, base); ctx.scale(a.f.escala, 1); ctx.fillText(a.texto, 0, 0); ctx.restore(); }
        else ctx.fillText(a.texto, x, base);
        if (r.u && r.u !== 'none') subrayado(ctx, a, x, base, a.ancho);
        if (r.tachado || r.tachadoDoble) {
          ctx.save(); ctx.strokeStyle = colorDe(r); ctx.lineWidth = Math.max(a.f.tam * 0.05, 0.4);
          const yy = base - a.f.tam * 0.3;
          ctx.beginPath(); ctx.moveTo(x, yy); ctx.lineTo(x + a.ancho, yy); if (r.tachadoDoble) { ctx.moveTo(x, yy + a.f.tam * 0.1); ctx.lineTo(x + a.ancho, yy + a.f.tam * 0.1); } ctx.stroke(); ctx.restore();
        }
      }
    }

    // ───────── Bordes y sombreado ─────────
    function trazoBorde(ctx, b, x1, y1, x2, y2) {
      if (!b || !b.tipo || b.tipo === 'nil' || b.tipo === 'none') return;
      const w = Math.max(b.ancho === undefined ? 0.5 : b.ancho, 0.25);
      ctx.save();
      ctx.strokeStyle = m.color(b.color) || '#000000'; ctx.lineWidth = w;
      if (/dotted/i.test(b.tipo)) ctx.setLineDash([w, w * 1.5]); else if (/dash/i.test(b.tipo)) ctx.setLineDash([w * 4, w * 2]);
      const horizontal = y1 === y2;
      const trazar = (d) => { ctx.beginPath(); if (horizontal) { ctx.moveTo(x1, y1 + d); ctx.lineTo(x2, y2 + d); } else { ctx.moveTo(x1 + d, y1); ctx.lineTo(x2 + d, y2); } ctx.stroke(); };
      if (/double/i.test(b.tipo)) { trazar(-w); trazar(w); } else trazar(0);
      ctx.restore();
    }

    function fondosYBordesDeParrafos(ctx, items) {
      let i = 0;
      while (i < items.length) {
        const it = items[i];
        if (it.t !== 'linea') { i++; continue; }
        let j = i;
        while (j + 1 < items.length && items[j + 1].t === 'linea' && items[j + 1].b === it.b) j++;
        const p = it.b.p, grupo = items.slice(i, j + 1);
        const relleno = p.sombreado && hex6(p.sombreado.relleno), bordes = p.bordes;
        if (relleno || bordes) {
          const sang = p.sangria || {};
          const x1 = it.x + (sang.izq || 0), x2 = it.x + it.r.anchoDisp - (sang.der || 0);
          const y1 = grupo[0].y, y2 = grupo[grupo.length - 1].y + grupo[grupo.length - 1].l.alto;
          const e = (lado) => (bordes && bordes[lado] && bordes[lado].espacio) || 0;
          const ev = (lado) => (lado === 'top' ? it.r.bordeArriba : it.r.bordeAbajo) || e(lado);
          if (relleno) { ctx.fillStyle = relleno; ctx.fillRect(x1 - e('left'), y1 - ev('top'), x2 - x1 + e('left') + e('right'), y2 - y1 + ev('top') + ev('bottom')); }
          if (bordes) {
            const xa = x1 - e('left'), xb = x2 + e('right'), ya = y1 - ev('top'), yb = y2 + ev('bottom');
            trazoBorde(ctx, bordes.top, xa, ya, xb, ya); trazoBorde(ctx, bordes.bottom, xa, yb, xb, yb);
            trazoBorde(ctx, bordes.left, xa, ya, xa, yb); trazoBorde(ctx, bordes.right, xb, ya, xb, yb);
          }
        }
        i = j + 1;
      }
    }

    function fondoDeTabla(ctx, t) {
      const bt = t.b.tbl || {};
      const rellenoT = bt.sombreado && hex6(bt.sombreado.relleno);
      for (const f of t.filas) for (const c of f.celdas) {
        const relleno = (c.tc.sombreado && hex6(c.tc.sombreado.relleno)) || rellenoT;
        if (relleno) { ctx.fillStyle = relleno; ctx.fillRect(t.x + c.x, t.y + f.y, c.ancho, c.alto); }
      }
    }
    function bordesDeTabla(ctx, t) {
      const bT = (t.b.tbl && t.b.tbl.bordes) || {};
      const nFilas = t.b.filas.length, nCols = t.b.nCols;
      for (const f of t.filas) for (const c of f.celdas) {
        const bc = c.tc.bordes || {};
        const x1 = t.x + c.x, x2 = x1 + c.ancho, y1 = t.y + f.y, y2 = y1 + c.alto;
        const iF = f.idx === undefined ? 0 : f.idx;
        const ultimaFila = iF + (c.filas || 1) >= nFilas, ultimaCol = c.col + c.span >= nCols;
        trazoBorde(ctx, bc.top || (iF === 0 ? bT.top : bT.insideH), x1, y1, x2, y1);
        trazoBorde(ctx, bc.bottom || (ultimaFila ? bT.bottom : bT.insideH), x1, y2, x2, y2);
        trazoBorde(ctx, bc.left || (c.col === 0 ? bT.left : bT.insideV), x1, y1, x1, y2);
        trazoBorde(ctx, bc.right || (ultimaCol ? bT.right : bT.insideV), x2, y1, x2, y2);
      }
    }

    // ───────── Hoja ─────────
    function flotante(ctx, q) {
      const z = q.pieza;
      if (!z) return;
      const img = z.destino ? cargarImagen(z.destino, modelo.zip) : null;
      if (img && img !== 'pendiente') ctx.drawImage(img, q.x, q.y, q.ancho, q.alto);
      else if (z.t === 'imagen' || z.t === 'cuadro' || z.t === 'dibujo') { ctx.save(); ctx.fillStyle = '#eef0f2'; ctx.fillRect(q.x, q.y, q.ancho, q.alto); ctx.strokeStyle = '#b8bec4'; ctx.lineWidth = 0.5; ctx.strokeRect(q.x, q.y, q.ancho, q.alto); ctx.restore(); }
    }
    function pintarHoja(ctx, hoja) {
      ctx.save();
      ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, hoja.ancho, hoja.alto);
      ctx.textBaseline = 'alphabetic'; ctx.textAlign = 'left';
      const grupos = [hoja.items];
      if (hoja.encabezado) grupos.push(hoja.encabezado.items);
      if (hoja.pie) grupos.push(hoja.pie.items);
      for (const q of hoja.flot || []) if (q.pieza && q.pieza.ancla && q.pieza.ancla.detras) flotante(ctx, q);
      for (const items of grupos) {
        fondosYBordesDeParrafos(ctx, items);
        for (const it of items) if (it.t === 'tabla') fondoDeTabla(ctx, it);
        for (const it of items) if (it.t === 'linea') lineaDeTexto(ctx, it);
        for (const it of items) if (it.t === 'tabla') bordesDeTabla(ctx, it);
      }
      for (const q of hoja.flot || []) if (q.pieza && !(q.pieza.ancla && q.pieza.ancla.detras)) flotante(ctx, q);
      // separador de las notas al pie: una línea corta sobre la primera nota
      if (hoja.notas && hoja.notas.length && hoja.separadorY !== undefined) {
        ctx.strokeStyle = '#000000'; ctx.lineWidth = 0.5; ctx.beginPath();
        ctx.moveTo(hoja.seccion.izquierda, hoja.separadorY); ctx.lineTo(hoja.seccion.izquierda + 144, hoja.separadorY); ctx.stroke();
      }
      ctx.restore();
    }

    return Object.assign(api, { pintarHoja, precargar, imagen: (d) => cargarImagen(d, modelo.zip) });
  }

  globalThis.PintorDoc = { crear };
})();

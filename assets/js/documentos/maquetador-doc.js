/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Maquetador paginado de Word (Fase 6, etapa 6.2). Reparte los bloques de un documento (docx-modelo.js) en páginas: corta cada párrafo en
// líneas (motor-parrafo.js), decide dónde cae cada salto de página (control de viudas y huérfanas, «conservar con el siguiente»,
// «conservar líneas juntas», «salto de página antes», saltos de página escritos, secciones), calcula las tablas (anchos de columna, alto de
// cada fila, celdas combinadas), y coloca encabezados y pies de cada hoja (que empujan el texto si son más altos que su margen).
// Unidad: puntos (pt), con el origen en la esquina superior izquierda de cada hoja.
//
// Resultado: { paginas: [{ numero, seccion, ancho, alto, items, encabezado, pie }], total }
//   item de texto:  { t: 'linea', b (bloque), r (resultado de romper), l (línea), x, y }   x, y: esquina de la línea en la hoja; la posición de
//                   cada pieza es x + átomo.x y la línea base es y + l.base
//   item de tabla:  { t: 'tabla', b, x, y, ancho, filas: [{ y, alto, celdas: [{ x, ancho, alto, tc, bordes… }] }] }  (coordenadas relativas al item)
// Lo que no se maqueta todavía: columnas, imágenes y cuadros flotantes (no ocupan lugar), notas al pie, filas de tabla partidas entre hojas.
//
// Uso: const mq = MaquetadorDoc.crear(modelo, ctx2d); const dis = await mq.maquetar();
(function () {
  'use strict';
  const TOL = 0.75;   // pt de tolerancia al decidir si una línea cabe (el maquetador del sitio usaba 1 px)
  const MAX_PAGINAS = 5000;

  function crear(modelo, ctx, opciones = {}) {
    const m = modelo.m;
    const motor = MotorParrafo.crear(ctx, m, opciones);
    const suprimirAntes = opciones.suprimirAntes !== false;   // Word no deja el espacio anterior de un párrafo al principio de una hoja que empezó sola

    const desplazar = (items, dx, dy) => items.map((it) => ({ ...it, x: it.x + dx, y: it.y + dy }));
    const anchoPt = (w, disponible) => {
      if (!w || w.valor === undefined || Number.isNaN(w.valor)) return undefined;
      if (w.tipo === 'dxa') return w.valor / 20;
      if (w.tipo === 'pct') return disponible * (w.porcentaje !== undefined ? w.porcentaje : w.valor / 50) / 100;
      return undefined;
    };
    const anchoBorde = (b) => (b && b.tipo && b.tipo !== 'nil' && b.tipo !== 'none' ? Math.max(b.ancho === undefined ? 0.5 : b.ancho, 0.25) : 0);
    const sumaCols = (cols, a, n) => { let s = 0; for (let k = 0; k < n; k++) s += cols[a + k] || 0; return s; };
    const separacion = (previo, rPrev, b, r) => {
      if (!previo) return r.antes;
      if (b.p.sinEspacioEntreIguales && previo.estilo === b.estilo) return 0;
      return rPrev.despues + r.antes;
    };

    // ───────── Flujo sin páginas (celdas de tabla, encabezados y pies) ─────────
    function flujo(bloques, ancho, pag) {
      const items = [];
      let y = 0, previo = null, rPrev = null;
      for (const b of bloques) {
        if (b.tipo === 'p') {
          const r = motor.romper(b, ancho, pag);
          y += separacion(previo, rPrev, b, r);
          for (const l of r.lineas) items.push({ t: 'linea', b, r, l, x: 0, y: y + l.y });
          y += r.alto;
          previo = b; rPrev = r;
        } else {
          if (rPrev) y += rPrev.despues;
          const t = medirTabla(b, ancho, pag);
          items.push(...t.itemsPlanos(0, y));
          y += t.alto;
          previo = null; rPrev = null;
        }
      }
      return { alto: y + (rPrev ? rPrev.despues : 0), items };
    }

    // ───────── Tablas ─────────
    function medirTabla(b, disponible, pag) {
      const clave = disponible + '|' + (pag && pag.pagina) + '|' + (pag && pag.paginas);
      const previo = tablas.get(b);
      if (previo && previo.clave === clave) return previo.t;
      const t = calcularTabla(b, disponible, pag);
      tablas.set(b, { clave, t });
      return t;
    }
    const tablas = new WeakMap();

    function calcularTabla(b, disponible, pag) {
      const tbl = b.tbl || {};
      const margenesT = { top: 0, bottom: 0, left: 0, right: 0, ...(tbl.margenes || {}) };
      const nCols = b.nCols;
      const grid = b.grid.slice();
      while (grid.length < nCols) grid.push(0);
      const wTabla = anchoPt(b.anchoDecl || (tbl.ancho && { tipo: tbl.ancho.tipo, valor: tbl.ancho.valor }), disponible);
      let cols = grid.slice();
      const sumaGrid = cols.reduce((a, c) => a + c, 0);
      const primera = b.filas[0] ? b.filas[0].celdas : [];
      if (primera.length) {
        const pref = new Array(nCols).fill(0);
        let completas = true;
        for (const c of primera) {
          const w = anchoPt(c.anchoDecl, wTabla || disponible);
          if (w === undefined) { completas = false; continue; }
          for (let k = 0; k < c.span; k++) pref[c.col + k] = w / c.span;
        }
        if (completas && pref.every((x) => x > 0) && (!wTabla || sumaGrid < 1 || Math.abs(sumaGrid - pref.reduce((a, c) => a + c, 0)) > 1)) cols = pref;
      }
      let total = cols.reduce((a, c) => a + c, 0);
      if (total <= 0) { cols = new Array(nCols).fill((wTabla || disponible) / nCols); total = cols.reduce((a, c) => a + c, 0); }
      if (wTabla !== undefined && Math.abs(wTabla - total) > 0.5) { const f = wTabla / total; cols = cols.map((x) => x * f); total = wTabla; }
      else if (wTabla === undefined && total > disponible && disponible > 0) { const f = disponible / total; cols = cols.map((x) => x * f); total = disponible; }

      const todas = b.filas.flatMap((f) => f.celdas);
      const sinAnchos = todas.every((c) => !c.anchoDecl || !(c.anchoDecl.valor > 0));
      const rejillaFalsa = sinAnchos && sumaGrid / nCols < 30;
      const autoajuste = tbl.disposicion !== 'fixed';
      const margenesDe = (c) => ({ ...margenesT, ...(c.tc.margenes || {}) });
      const bT = tbl.bordes || {};
      // borde de cada lado de una celda: el suyo; si no, el de la tabla (en el interior insideH / insideV)
      const bordesDe = (c, i, ultimaFila) => {
        const bc = c.tc.bordes || {};
        const ultimaCol = c.col + c.span >= nCols;
        return {
          top: anchoBorde(bc.top || (i === 0 ? bT.top : bT.insideH)), bottom: anchoBorde(bc.bottom || (ultimaFila ? bT.bottom : bT.insideH)),
          left: anchoBorde(bc.left || (c.col === 0 ? bT.left : bT.insideV)), right: anchoBorde(bc.right || (ultimaCol ? bT.right : bT.insideV)),
        };
      };
      // Contenido mínimo y máximo de cada columna (solo celdas que no abarcan varias columnas)
      if (autoajuste) {
        const mins = new Array(nCols).fill(0), maxs = new Array(nCols).fill(0);
        for (const c of todas) {
          if (c.omitir || c.span !== 1) continue;
          const bd = bordesDe(c, 1, false);
          const mg = { ...margenesDe(c) }; mg.left += bd.left; mg.right += bd.right;
          let mn = 0, mx = 0;
          for (const bl of c.bloques) {
            if (bl.tipo !== 'p') continue;
            const k = motor.contenido(bl, pag);
            mn = Math.max(mn, k.min); mx = Math.max(mx, k.max);
          }
          mins[c.col] = Math.max(mins[c.col], mn + mg.left + mg.right);
          maxs[c.col] = Math.max(maxs[c.col], mx + mg.left + mg.right);
        }
        if (rejillaFalsa) {
          // la rejilla no dice nada: Word reparte según el contenido (como el ajuste automático de CSS)
          const objetivo = wTabla !== undefined ? wTabla : disponible;
          const sMax = maxs.reduce((a, c) => a + c, 0), sMin = mins.reduce((a, c) => a + c, 0);
          if (sMax <= objetivo) cols = wTabla !== undefined && sMax > 0 ? maxs.map((x) => x * wTabla / sMax) : maxs.map((x) => Math.max(x, 1));
          else if (sMin >= objetivo) cols = mins.slice();
          else { const f = (objetivo - sMin) / (sMax - sMin); cols = mins.map((mn, k) => mn + (maxs[k] - mn) * f); }
          total = cols.reduce((a, c) => a + c, 0);
        } else {
          // una palabra que no cabe ensancha su columna; lo que falta se quita de las columnas con holgura
          for (let j = 0; j < nCols; j++) {
            if (cols[j] + 0.01 >= mins[j]) continue;
            let falta = mins[j] - cols[j];
            cols[j] = mins[j];
            const holgura = cols.map((w, k) => (k === j ? 0 : Math.max(w - mins[k], 0)));
            const sH = holgura.reduce((a, c) => a + c, 0);
            const quitar = Math.min(falta, sH);
            if (sH > 0) for (let k = 0; k < nCols; k++) cols[k] -= quitar * holgura[k] / sH;
          }
          total = cols.reduce((a, c) => a + c, 0);
        }
      }
      const sangria = (tbl.sangria && tbl.sangria.tipo !== 'pct' ? (tbl.sangria.valor || 0) / 20 : 0) - ((m.ajustes && m.ajustes.modoCompat) < 15 ? margenesT.left : 0);
      const x = tbl.jc === 'center' ? Math.max((disponible - total) / 2, 0) : tbl.jc === 'right' || tbl.jc === 'end' ? Math.max(disponible - total, 0) : sangria;

      // Filas
      const xs = [0]; for (let j = 0; j < nCols; j++) xs.push(xs[j] + cols[j]);
      const filas = b.filas.map((f, i) => {
        const celdas = f.celdas.map((c) => {
          if (c.omitir) return { ...c, contenido: null, alto: 0 };
          const bd = bordesDe(c, i, i === b.filas.length - 1);
          const mg = { ...margenesDe(c) }; mg.left += bd.left; mg.top += bd.top; mg.bottom += i === b.filas.length - 1 ? bd.bottom : 0; mg.right += bd.right;
          const w = sumaCols(cols, c.col, c.span);
          const fl = flujo(c.bloques, Math.max(w - mg.left - mg.right, 1), pag);
          return { ...c, x: xs[c.col], ancho: w, mg, contenido: fl, alto: fl.alto + mg.top + mg.bottom };
        });
        let alto = 0;
        for (const c of celdas) if (!c.omitir && c.filas === 1) alto = Math.max(alto, c.alto);
        if (f.tr.alto) alto = f.tr.regla === 'exact' ? f.tr.alto : Math.max(alto, f.tr.alto);
        return { f, celdas, alto };
      });
      // celdas combinadas verticalmente: si su contenido pide más que las filas que abarca, crece la última
      filas.forEach((fila, i) => fila.celdas.forEach((c) => {
        if (c.omitir || c.filas <= 1) return;
        let suma = 0; for (let k = 0; k < c.filas && i + k < filas.length; k++) suma += filas[i + k].alto;
        if (c.alto > suma) filas[Math.min(i + c.filas - 1, filas.length - 1)].alto += c.alto - suma;
      }));
      let y = 0;
      for (const fila of filas) { fila.y = y; y += fila.alto; }
      const altoTotal = y;

      // Salida plana: las filas con sus celdas y las líneas de texto de cada celda, en coordenadas de la tabla
      function itemsPlanos(dx, dy, desde = 0, hasta = filas.length, yBase = 0) {
        const sal = [];
        const rf = [];
        for (let i = desde; i < hasta; i++) {
          const fila = filas[i];
          rf.push({ idx: i, y: fila.y - yBase, alto: fila.alto, tr: fila.f.tr, celdas: fila.celdas.filter((c) => !c.omitir).map((c) => {
            let alto = fila.alto; for (let k = 1; k < c.filas && i + k < filas.length; k++) alto += filas[i + k].alto;
            return { x: c.x, ancho: c.ancho, alto, tc: c.tc, mg: c.mg, col: c.col, span: c.span, filas: c.filas };
          }) });
        }
        sal.push({ t: 'tabla', b, x: x + dx, y: dy, ancho: total, alto: (hasta > desde ? filas[hasta - 1].y + filas[hasta - 1].alto - filas[desde].y : 0), filas: rf, cols });
        for (let i = desde; i < hasta; i++) {
          const fila = filas[i];
          for (const c of fila.celdas) {
            if (c.omitir || !c.contenido) continue;
            let alto = fila.alto; for (let k = 1; k < c.filas && i + k < filas.length; k++) alto += filas[i + k].alto;
            const libre = alto - c.alto;
            const va = c.tc.alineacionVertical;
            const ajuste = va === 'center' ? libre / 2 : va === 'bottom' ? libre : 0;
            sal.push(...desplazar(c.contenido.items, x + dx + c.x + c.mg.left, dy + fila.y - yBase + c.mg.top + Math.max(ajuste, 0)));
          }
        }
        return sal;
      }
      return { alto: altoTotal, ancho: total, x, filas, cols, itemsPlanos };
    }

    // Una fila más alta que la hoja se parte por líneas de texto, como hace Word: cada celda reparte sus líneas en trozos de a lo más `H` pt (cada celda corta donde le toca,
    // no tienen que coincidir). Solo filas simples: sin celdas combinadas ni cubiertas, sin tablas anidadas dentro de las celdas y sin altura exacta; si no, devuelve null.
    //   → [{ alto, itemsPlanos(dx, dy) }] con un trozo por hoja, o null
    function partirFila(t, b, i, H) {
      const fila = t.filas[i];
      if (fila.celdas.some((c) => c.omitir) || fila.f.tr.regla === 'exact') return null;
      const estado = fila.celdas.map((c) => ({ c, lineas: c.contenido ? c.contenido.items.slice().sort((p, q) => p.y - q.y) : [], k: 0 }));
      if (estado.some((e) => e.c.filas > 1 || e.lineas.some((it) => it.t !== 'linea'))) return null;
      const partes = [];
      for (let guardia = 0; guardia < 5000; guardia++) {
        const primera = partes.length === 0;
        const trozos = [];
        let alto = 1;
        for (const e of estado) {
          const { c } = e, arriba = primera ? c.mg.top : 0;
          const base = primera || e.k >= e.lineas.length ? 0 : e.lineas[e.k].y;   // arriba de la primera línea que queda (en las coordenadas de la celda entera)
          const lim = H - arriba - c.mg.bottom;
          const sel = [];
          while (e.k < e.lineas.length) {
            const it = e.lineas[e.k];
            if (sel.length && it.y + it.l.alto - base > lim + TOL) break;
            sel.push(it); e.k++;
          }
          const contenido = sel.length ? sel[sel.length - 1].y + sel[sel.length - 1].l.alto - base : 0;
          const h = arriba + contenido + (e.k >= e.lineas.length ? c.mg.bottom : 0);
          alto = Math.max(alto, h);
          trozos.push({ c, sel, base, arriba });
        }
        partes.push({
          alto,
          itemsPlanos(dx, dy) {
            const celdas = trozos.map((z) => ({ x: z.c.x, ancho: z.c.ancho, alto, tc: z.c.tc, mg: z.c.mg, col: z.c.col, span: z.c.span, filas: 1 }));
            const sal = [{ t: 'tabla', b, x: t.x + dx, y: dy, ancho: t.ancho, alto, filas: [{ idx: i, y: 0, alto, tr: fila.f.tr, celdas }], cols: t.cols }];
            for (const z of trozos) sal.push(...desplazar(z.sel, t.x + dx + z.c.x + z.c.mg.left, dy + z.arriba - z.base));
            return sal;
          },
        });
        if (estado.every((e) => e.k >= e.lineas.length)) return partes;
      }
      return null;
    }

    // ───────── Encabezados y pies ─────────
    async function partesDeSeccion(cfg, previas) {
      const actuales = { ...previas };
      for (const r of cfg.referencias) {
        const rel = modelo.rels.get(r.id);
        const parte = rel && !rel.externo ? await modelo.parte(rel.destino) : null;
        actuales[r.tipo + ':' + r.cual] = parte;
      }
      return actuales;
    }

    // ───────── Paginación ─────────
    // NUMPAGES en el cuerpo: se necesita el total antes de maquetar. Se maqueta una vez con 1 y, si el documento usa el campo, otra con el total que salió
    // (Word hace lo mismo al repaginar). Los encabezados y pies ya reciben el total real.
    let totalBase = 1;
    async function maquetar() {
      totalBase = 1;
      const r = await maquetarUnaVez();
      const usaTotal = modelo.parrafos().some((b) => (b.piezas || []).some((z) => z.t === 'campo' && z.nombre === 'NUMPAGES'));
      if (!usaTotal || r.total <= 1) return r;
      totalBase = r.total;
      return maquetarUnaVez();
    }
    async function maquetarUnaVez() {
      // Partes de encabezado y pie, y fuentes
      const partes = [];
      let previas = {};
      for (const cfg of modelo.secciones) { previas = await partesDeSeccion(cfg, previas); partes.push(previas); }
      const todos = modelo.parrafos();
      for (const pr of partes) for (const k of Object.keys(pr)) if (pr[k]) for (const b of pr[k].bloques) todos.push(...(b.tipo === 'p' ? [b] : b.filas.flatMap((f) => f.celdas.flatMap((c) => c.bloques.filter((x) => x.tipo === 'p')))));
      if (typeof PptxTexto !== 'undefined' && PptxTexto.cargarFuentes) await PptxTexto.cargarFuentes(motor.fuentesDe(todos));
      if (typeof document !== 'undefined' && document.fonts && document.fonts.ready) await document.fonts.ready;

      const paginas = [];
      let nCol = 1, col = 0, anchoCol = 0, huecoCol = 0, yTope = 0;
      let pag = null, y = 0, yMax = 0, pendiente = null, previo = null, rPrev = null, numero = 0, blando = false, secIdx = 0, enSeccion = 0;
      let cfg = modelo.secciones[0];
      const anchoTexto = () => cfg.ancho - cfg.izquierda - cfg.derecha;

      // alto de encabezado y pie de la hoja que viene (según portada y pares/impares)
      const tipoDe = (c, primeraDeSeccion, n) => (c.portada && primeraDeSeccion ? 'first' : m.ajustes.parImpar && n % 2 === 0 ? 'even' : 'default');
      // Rectángulo (en la hoja) de un dibujo flotante, con su distancia al texto
      const rectanguloFlotante = (z, yParrafo) => {
        const a = z.ancla, c = cfg, ancho = z.ancho, alto = z.alto;
        const textoW = c.ancho - c.izquierda - c.derecha;
        let x;
        const baseH = a.h && a.h.rel === 'page' ? 0 : c.izquierda;
        const anchoH = a.h && a.h.rel === 'page' ? c.ancho : textoW;
        if (a.h && a.h.align === 'right') x = baseH + anchoH - ancho; else if (a.h && a.h.align === 'center') x = baseH + (anchoH - ancho) / 2; else x = baseH + (a.h ? a.h.offset : 0);
        let yy;
        const rv = a.v && a.v.rel;
        const baseV = rv === 'page' ? 0 : rv === 'margin' || rv === 'topMargin' ? c.arriba : yParrafo;
        if (a.v && a.v.align === 'bottom') yy = (rv === 'page' ? c.alto : c.alto - c.abajo) - alto; else if (a.v && a.v.align === 'center') yy = baseV + ((rv === 'page' ? c.alto : c.alto - c.arriba - c.abajo) - alto) / 2; else yy = baseV + (a.v ? a.v.offset : 0);
        const d = a.dist;
        return { pieza: z, x0: x - d.l, x1: x + ancho + d.r, y0: yy - d.t, y1: yy + alto + d.b, x, y: yy, ancho, alto, lado: a.envoltura === 'wrapTopAndBottom' ? 'arribaAbajo' : a.envoltura === 'wrapNone' ? 'ninguno' : 'lados' };
      };
      // Un párrafo en un marco (w:framePr con ancho) flota como un dibujo: se maqueta aparte y el texto lo rodea. Se ancla al párrafo que sigue.
      const marcosPendientes = [];
      const colocarMarco = (b, yParrafo) => {
        const fr = b.p.marco, t = (v) => (fr[v] !== undefined ? parseFloat(fr[v]) / 20 : 0);
        const w = Math.max(t('w'), 20);
        const r = motor.romper({ ...b, p: { ...b.p, sangria: undefined } }, w, { pagina: numero, paginas: totalBase });
        const h = fr.h && fr.hRule === 'exact' ? t('h') : Math.max(r.alto, fr.h ? t('h') : 0);
        const base = fr.hAnchor === 'page' ? 0 : cfg.izquierda;
        const xx = fr.xAlign === 'right' ? base + (fr.hAnchor === 'page' ? cfg.ancho : anchoTexto()) - w : fr.xAlign === 'center' ? base + ((fr.hAnchor === 'page' ? cfg.ancho : anchoTexto()) - w) / 2 : base + t('x');
        const yy = fr.vAnchor === 'page' ? t('y') : fr.vAnchor === 'margin' ? cfg.arriba + t('y') : yParrafo + t('y');
        const hs = t('hSpace'), vs = t('vSpace');
        for (const l of r.lineas) pag.items.push({ t: 'linea', b, r, l, x: xx, y: yy + l.y });
        pag.flot.push({ pieza: null, marco: b, x: xx, y: yy, ancho: w, alto: h, x0: xx - hs, x1: xx + w + hs, y0: yy - vs, y1: yy + h + vs, lado: fr.wrap === 'none' || fr.wrap === 'notBeside' ? 'arribaAbajo' : 'lados' });
      };
      const parteDe = (idxSec, tipo, quien) => partes[idxSec][quien + ':' + tipo] || (tipo === 'even' ? partes[idxSec][quien + ':default'] : null);
      const altoParte = (idxSec, tipo, quien) => {
        const p = parteDe(idxSec, tipo, quien);
        if (!p) return 0;
        return flujo(p.bloques, modelo.secciones[idxSec].ancho - modelo.secciones[idxSec].izquierda - modelo.secciones[idxSec].derecha, { pagina: 1, paginas: totalBase }).alto;
      };
      function nuevaPagina(esBlando) {
        if (paginas.length >= MAX_PAGINAS) throw Object.assign(new Error('El documento tiene demasiadas hojas.'), { codigo: 'hojas' });   // la aplicación lo traduce
        numero = enSeccion === 0 && cfg.inicioPaginas !== undefined ? cfg.inicioPaginas : numero + 1;
        const tipo = tipoDe(cfg, enSeccion === 0, numero);
        const hEnc = altoParte(secIdx, tipo, 'encabezado'), hPie = altoParte(secIdx, tipo, 'pie');
        const arriba = Math.max(cfg.arriba, hEnc ? cfg.encabezado + hEnc : 0);
        const abajo = Math.max(cfg.abajo, hPie ? cfg.pie + hPie : 0);
        pag = { numero, seccion: cfg, secIdx, ancho: cfg.ancho, alto: cfg.alto, items: [], tipo, hEnc, hPie, vacia: true, notas: [], reserva: 0, flot: [] };
        paginas.push(pag);
        y = arriba; yMax = cfg.alto - abajo; pag.base = yMax; pag.tope = arriba;
        yTope = arriba; col = 0;
        pendiente = null; previo = null; rPrev = null;
        blando = !!esBlando;
        enSeccion++;
      }
      // Notas al pie: cada hoja reserva abajo el lugar de las notas citadas en ella (más el separador, la primera vez)
      const altoNotas = new Map();
      const anchoNotas = () => cfg.ancho - cfg.izquierda - cfg.derecha;
      const altoDe = (n) => { const k = n.id + '|' + anchoNotas(); if (!altoNotas.has(k)) altoNotas.set(k, flujo(n.bloques, anchoNotas(), { pagina: 1, paginas: totalBase }).alto); return altoNotas.get(k); };
      const altoSeparador = () => { const k = 'sep|' + anchoNotas(); if (!altoNotas.has(k)) altoNotas.set(k, modelo.notas.separador ? flujo(modelo.notas.separador, anchoNotas(), { pagina: 1, paginas: totalBase }).alto : 0); return altoNotas.get(k); };
      const notasDe = (l, ya) => {
        const sal = [];
        for (const a of l.atomos) {
          const n = a.nota && modelo.notas.get(a.nota);
          if (n && !pag.notas.includes(n) && !ya.includes(n) && !sal.includes(n)) sal.push(n);
        }
        return sal;
      };
      const disponible = () => yMax - pag.reserva - y;
      const siguienteBloque = (lista, i) => lista[i + 1];

      // Necesidad de altura de un párrafo con «conservar con el siguiente»: él y el principio de lo que sigue deben ir en la misma hoja.
      function necesidad(lista, i, ancho) {
        let need = 0;
        for (let k = i; k < lista.length; k++) {
          const b = lista[k];
          if (b.tipo === 'p') {
            const r = motor.romper(b, ancho, { pagina: numero, paginas: totalBase });
            const encadena = b.p.mantenerConSiguiente && k + 1 < lista.length;
            const nl = r.lineas.length;
            let h;
            if (encadena || b.p.mantenerLineas) h = r.alto;
            else { const n = Math.min(nl, b.p.controlViudas === false ? 1 : 2); h = r.lineas.slice(0, n).reduce((a, l) => a + l.alto, 0); }
            need += (k === i ? 0 : r.antes) + h + (encadena ? r.despues : 0);
            if (need > pag.base - pag.tope) return 0;   // una cadena más alta que la hoja no cabe nunca: Word ignora «conservar con el siguiente» (y no se sigue sumando, que era cuadrático)
            if (!encadena) break;
          } else {
            const t = medirTabla(b, ancho, { pagina: numero, paginas: totalBase });
            need += t.filas[0] ? t.filas[0].alto : 0;
            break;
          }
        }
        return need;
      }

      // Columnas: una sección con varias columnas llena la primera hasta el pie de la hoja y sigue arriba en la segunda
      const iniciarColumnas = () => {
        nCol = cfg.columnas ? cfg.columnas.n : 1;
        huecoCol = cfg.columnas ? cfg.columnas.espacio : 0;
        anchoCol = (anchoTexto() - huecoCol * (nCol - 1)) / nCol;
        col = 0; yTope = y + (rPrev && pag && !pag.vacia ? rPrev.despues : 0);
      };
      // lo que no cabe pasa a la columna siguiente o, en la última, a una hoja nueva
      const desbordar = () => {
        if (col < nCol - 1) { col++; y = yTope; pag.vacia = true; blando = true; pendiente = null; previo = null; rPrev = null; }
        else { nuevaPagina(true); iniciarColumnas(); yTope = y; }
      };
      const hojaNueva = (blandoNuevo) => { nuevaPagina(blandoNuevo); iniciarColumnas(); };
      for (secIdx = 0; secIdx < modelo.secciones.length; secIdx++) {
        cfg = modelo.secciones[secIdx];
        const inicio = secIdx === 0 ? 0 : modelo.secciones[secIdx - 1].hasta + 1;
        const lista = modelo.bloques.slice(inicio, cfg.hasta + 1);
        if (!pag || !cfg.continua || cfg.ancho !== pag.ancho || cfg.alto !== pag.alto) { enSeccion = 0; hojaNueva(false); } else iniciarColumnas();
        for (let i = 0; i < lista.length; i++) {
          const b = lista[i];
          const ancho = anchoCol;
          const izqPagina = cfg.izquierda + col * (anchoCol + huecoCol);
          if (b.tipo === 'p') {
            if (b.p.marco && b.p.marco.w) { marcosPendientes.push(b); continue; }
            // el párrafo vacío que lleva un salto de sección (la marca del salto) no ocupa lugar
            if (b.seccion && !b.piezas.some((z) => z.t !== 'ancla')) continue;
            // dibujos flotantes que ancla este párrafo: reservan su rectángulo en la hoja
            const antesB = b.p.espacio ? (b.p.espacio.antesAuto ? 14 : b.p.espacio.antes || 0) : 0;
            const yPar = y + (pag.vacia && blando && suprimirAntes ? 0 : (rPrev ? Math.max(rPrev.despues, 0) : 0) + antesB);
            while (marcosPendientes.length) colocarMarco(marcosPendientes.shift(), yPar);
            for (const z of b.piezas) if (z.ancla && !z.enLinea && !pag.flot.some((q) => q.pieza === z)) pag.flot.push(rectanguloFlotante(z, yPar));
            const r = motor.romper(b, ancho, { pagina: numero, paginas: totalBase, flotantes: pag.flot.some((q) => q.lado !== 'ninguno') ? pag.flot : null, x0: izqPagina, y0: yPar });
            if (b.p.saltoAntes && !pag.vacia) { hojaNueva(false); }
            if (b.p.mantenerConSiguiente && !pag.vacia) {
              const sep0 = separacion(previo, rPrev, b, r);
              if (sep0 + necesidad(lista, i, ancho) > disponible() + TOL) desbordar();
            }
            let k = 0;
            const n = r.lineas.length;
            while (k < n) {
              let sep = 0;
              if (k === 0) {
                sep = separacion(previo, rPrev, b, r);
                if (pag.vacia && blando && suprimirAntes) sep = 0;
              }
              // un salto de página escrito corta el párrafo: lo de antes (con la línea del salto) y lo de después van aparte
              let tope = n;
              for (let q = k; q < n; q++) if (r.lineas[q].saltoPagina || r.lineas[q].saltoColumna) { tope = q + 1; break; }
              let altura = 0, caben = 0, extra = 0;
              const nuevas = [];
              for (let q = k; q < tope; q++) {
                const nn = notasDe(r.lineas[q], nuevas);
                const h = nn.reduce((acc, n) => acc + altoDe(n), 0) + (nn.length && !pag.notas.length && !nuevas.length ? altoSeparador() : 0);
                const oc = r.lineas[q].alto + (r.lineas[q].gapAntes || 0);
                if (altura + oc + extra + h > disponible() - sep + TOL) break;
                altura += oc; extra += h; nuevas.push(...nn); caben++;
              }
              if (caben < tope - k) {
                // no caben todas: viudas y huérfanas, y «conservar líneas juntas»
                if (k === 0 && b.p.mantenerLineas && !pag.vacia) caben = 0;
                else if (b.p.controlViudas !== false) {
                  if (k === 0 && caben < 2) caben = 0;                                     // huérfana: una sola línea al pie de la hoja
                  else if (tope - k - caben === 1) caben = (k === 0 && caben <= 2) ? 0 : caben - 1;   // viuda: una sola línea al principio de la siguiente
                }
                if (caben === 0 && pag.vacia) caben = 1;                                   // nada cabe ni en una hoja vacía: se coloca igual
              }
              let salto = false, saltoCol = false;
              if (caben > 0) {
                y += sep;
                for (let q = k; q < k + caben; q++) {
                  const l = r.lineas[q];
                  pag.items.push({ t: 'linea', b, r, l, x: cfg.izquierda + col * (anchoCol + huecoCol), y: y + l.y - r.lineas[k].yIni });
                }
                y += r.lineas.slice(k, k + caben).reduce((acc, l) => acc + l.alto + (l.gapAntes || 0), 0);
                pag.vacia = false;
                for (let q = k; q < k + caben; q++) {
                  const nn = notasDe(r.lineas[q], []);
                  if (nn.length && !pag.notas.length) pag.reserva += altoSeparador();
                  for (const n of nn) { pag.notas.push(n); pag.reserva += altoDe(n); }
                }
                salto = !!r.lineas[k + caben - 1].saltoPagina;
                saltoCol = !!r.lineas[k + caben - 1].saltoColumna;
                k += caben;
                previo = b; rPrev = r;
              }
              if (k < n) { if (salto) hojaNueva(false); else if (saltoCol) { if (nCol > 1) { desbordar(); blando = false; } else hojaNueva(false); } else desbordar(); }
            }
          } else {
            const t = medirTabla(b, ancho, { pagina: numero, paginas: totalBase });
            // filas de encabezado repetidas al principio de la hoja que sigue
            const repetirEncabezados = (tabla, desde) => {
              let h = 0;
              const enc = [];
              for (let q = 0; q < tabla.filas.length && tabla.filas[q].f.tr && tabla.filas[q].f.tr.encabezado; q++) { enc.push(q); h += tabla.filas[q].alto; }
              if (enc.length && enc.length < desde + 1 && h < disponible() / 2) {
                pag.items.push(...tabla.itemsPlanos(cfg.izquierda + col * (anchoCol + huecoCol), y, 0, enc.length, 0));
                y += h;
              }
            };
            let i0 = 0;
            const sep = rPrev ? rPrev.despues : 0;
            y += sep;
            while (i0 < t.filas.length) {
              let i1 = i0, altura = 0;
              while (i1 < t.filas.length && altura + t.filas[i1].alto <= disponible() + TOL) { altura += t.filas[i1].alto; i1++; }
              // una celda combinada en vertical no se corta: si el corte cae dentro de una, se sube hasta la fila donde empieza (si empieza en esta misma hoja)
              if (i1 > i0 && i1 < t.filas.length) {
                let corte = i1;
                for (let j = i0; j < i1; j++) for (const c of t.filas[j].celdas) if (!c.omitir && j + c.filas > i1) corte = Math.min(corte, j);
                if (corte > i0) i1 = corte;
              }
              if (i1 === i0) {
                if (!pag.vacia) { desbordar(); continue; }
                // una fila más alta que la hoja: se parte por líneas; si no se puede (celdas combinadas, tablas dentro), se coloca entera aunque se salga
                const partes = partirFila(t, b, i0, disponible());
                if (partes && partes.length > 1) {
                  for (let k = 0; k < partes.length; k++) {
                    pag.items.push(...partes[k].itemsPlanos(cfg.izquierda + col * (anchoCol + huecoCol), y));
                    y += partes[k].alto; pag.vacia = false;
                    if (k < partes.length - 1) desbordar();
                  }
                  i0++;
                  if (i0 < t.filas.length) { desbordar(); repetirEncabezados(t, i0); }
                  continue;
                }
                i1 = i0 + 1;
              }
              const yBase = t.filas[i0].y;
              const items = t.itemsPlanos(cfg.izquierda + col * (anchoCol + huecoCol), y, i0, i1, yBase);
              pag.items.push(...items);
              y += t.filas[i1 - 1].y + t.filas[i1 - 1].alto - yBase;
              pag.vacia = false;
              i0 = i1;
              if (i0 < t.filas.length) {
                desbordar();
                repetirEncabezados(t, i0);
              }
            }
            previo = null; rPrev = null;
          }
        }
      }

      // Notas al pie de cada hoja, pegadas al margen de abajo
      for (const p of paginas) {
        if (!p.notas.length) continue;
        const c = p.seccion, ar = c.ancho - c.izquierda - c.derecha;
        let yy = p.base - p.reserva;
        if (modelo.notas.separador) { const sp = flujo(modelo.notas.separador, ar, { pagina: 1, paginas: totalBase }); p.separadorY = yy + sp.alto * 0.7; yy += sp.alto; }
        for (const n of p.notas) { const fl = flujo(n.bloques, ar, { pagina: 1, paginas: totalBase }); p.items.push(...desplazar(fl.items, c.izquierda, yy)); yy += fl.alto; }
      }
      // Encabezados y pies de cada hoja, con el número de hoja y el total
      const total = paginas.length;
      for (const p of paginas) {
        const c = p.seccion, ar = c.ancho - c.izquierda - c.derecha;
        const pagInfo = { pagina: p.numero, paginas: total };
        const enc = parteDe(p.secIdx, p.tipo, 'encabezado'), pie = parteDe(p.secIdx, p.tipo, 'pie');
        p.encabezado = enc ? flujo(enc.bloques, ar, pagInfo) : null;
        p.pie = pie ? flujo(pie.bloques, ar, pagInfo) : null;
        if (p.encabezado) p.encabezado.items = desplazar(p.encabezado.items, c.izquierda, c.encabezado);
        if (p.pie) p.pie.items = desplazar(p.pie.items, c.izquierda, c.alto - c.pie - p.pie.alto);
      }
      return { paginas, total };
    }

    // Una edición cambió este bloque (o una tabla que lo contiene): lo que se midió de él ya no vale
    const invalidar = (b) => { if (b) { tablas.delete(b); motor.invalidar(b); } else motor.invalidar(); };
    return { maquetar, motor, flujo, invalidar };
  }

  globalThis.MaquetadorDoc = { crear };
})();

/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Botones del autofiltro de Excel (hoja con autoFilter y tablas con filtro) del motor propio (etapa 6). Solo se ven en pantalla: Excel no los imprime
// (comprobado en el PDF del corpus 41-autofiltro). Cada celda de encabezado lleva un botón cuadrado pegado a su esquina de abajo a la derecha con una
// flecha hacia abajo; si la columna tiene un filtro aplicado la flecha lleva un embudo. El tamaño y los colores del botón NO están medidos contra
// Excel (necesitan CopyPicture, que no funcionó: PENDIENTES.md): son los de un botón de Excel 365 recordado a 100 % de zoom.
// hoja.filtroAuto { ref, columnas: Map(colId → { oculto, filtrada }) } y hoja.tablas[].filtro lo lee xlsx-modelo.js. Sin DOM.
//
// Uso: const lista = XlsxFiltros.botones(hoja);                     // [{ f, c, filtrada }] celdas de encabezado con botón
//      XlsxFiltros.dibujar(ctx, lista, x0, y0, x1, y1, op)           // un botón en la celda (x0, y0)–(x1, y1)

(function () {
  const MAX_BOTONES = 20000;
  const LADO = 17; // medido en pantalla (Excel 365, 100 %): caja de 17 × 17 px pegada a la esquina de abajo a la derecha de la celda (sobre el borde de la derecha, 1 px sobre el de abajo)

  function botones(hoja) {
    const lista = [];
    const poner = (f, c0, c1, filtro) => {
      for (let c = c0; c <= c1 && lista.length < MAX_BOTONES; c++) {
        const col = filtro && filtro.columnas ? filtro.columnas.get(c - c0) : null;
        if (col && col.oculto) continue;
        lista.push({ f, c, filtrada: !!(col && col.filtrada) });
      }
    };
    const fa = hoja.filtroAuto;
    if (fa && fa.ref) { const r = LibroXlsx.decodificarRango(fa.ref); if (r) poner(r.f0, r.c0, r.c1, fa); }
    for (const t of hoja.tablas || []) if (t.filtro && t.encabezados) poner(t.ref.f0, t.ref.c0, t.ref.c1, t.filtro);
    return lista;
  }

  function dibujar(ctx, b, x0, y0, x1, y1, op) {
    op = op || {};
    const w = x1 - x0, h = y1 - y0;
    if (!(w > LADO + 4) || !(h > 8)) return;
    const lado = Math.min(LADO, h - 2);
    const x = x1 - lado, y = y1 - lado - 1;
    ctx.save();
    ctx.translate(x, y); ctx.scale(lado / LADO, lado / LADO); // todo lo siguiente en píxeles de una caja de 17 × 17
    const relleno = ctx.createLinearGradient(0, 1, 0, 16); relleno.addColorStop(0, '#FFFFFF'); relleno.addColorStop(1, '#EFF2F7');
    ctx.fillStyle = relleno; ctx.fillRect(0, 0, LADO, LADO);
    ctx.strokeStyle = '#A6ACB3'; ctx.lineWidth = 1; ctx.strokeRect(0.5, 0.5, LADO - 1, LADO - 1);
    const grisar = (y0, y1) => { const g = ctx.createLinearGradient(0, y0, 0, y1); g.addColorStop(0, '#868A8D'); g.addColorStop(1, '#4E5052'); return g; };
    if (b.filtrada) { // columna filtrada: un embudo y, abajo a la izquierda, una flecha pequeña (sin la flecha grande de la lista)
      ctx.fillStyle = grisar(4, 12);
      ctx.beginPath(); ctx.moveTo(5, 4); ctx.lineTo(14, 4); ctx.lineTo(11, 7); ctx.lineTo(11, 12); ctx.lineTo(8, 12); ctx.lineTo(8, 7); ctx.closePath(); ctx.fill();
      ctx.fillStyle = grisar(10, 13);
      ctx.beginPath(); ctx.moveTo(2, 10); ctx.lineTo(7, 10); ctx.lineTo(5, 13); ctx.closePath(); ctx.fill();
    } else { // flecha de la lista: triángulo de 7 px de ancho y 4 de alto, centrado
      ctx.fillStyle = grisar(7, 11);
      ctx.beginPath(); ctx.moveTo(5, 7); ctx.lineTo(12, 7); ctx.lineTo(8.5, 11); ctx.closePath(); ctx.fill();
    }
    ctx.restore();
  }

  globalThis.XlsxFiltros = { botones, dibujar };
})();

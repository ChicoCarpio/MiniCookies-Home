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
  const LADO = 16;

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
    const x = x1 - lado - 1, y = y1 - lado - 1;
    ctx.save();
    ctx.fillStyle = '#F3F3F3'; ctx.fillRect(x, y, lado, lado);
    ctx.strokeStyle = '#A6A6A6'; ctx.lineWidth = 1; ctx.strokeRect(x + 0.5, y + 0.5, lado - 1, lado - 1);
    ctx.fillStyle = '#404040';
    const cx = x + lado / 2, cy = y + lado / 2;
    ctx.beginPath(); ctx.moveTo(cx - 3.5, cy - 1.5); ctx.lineTo(cx + 3.5, cy - 1.5); ctx.lineTo(cx, cy + 2.5); ctx.closePath(); ctx.fill();
    if (b.filtrada) { // embudo azul sobre la flecha
      ctx.fillStyle = '#1F6FB5';
      ctx.beginPath(); ctx.moveTo(x + 2, y + 2); ctx.lineTo(x + lado - 2, y + 2); ctx.lineTo(cx + 1, cy + 1); ctx.lineTo(cx + 1, y + lado - 3); ctx.lineTo(cx - 1, y + lado - 4); ctx.lineTo(cx - 1, cy + 1); ctx.closePath(); ctx.fill();
    }
    ctx.restore();
  }

  globalThis.XlsxFiltros = { botones, dibujar };
})();

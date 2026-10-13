/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Notas de celda de Excel (comentarios clásicos) del motor propio (etapa 6): el triángulo rojo de la celda, el cuadro de las notas que se muestran
// siempre (en pantalla y al imprimir "como se muestran en la hoja") con su flecha, y la lista de "al final de la hoja".
// hoja.comentarios lo lee xlsx-modelo.js (comments + VML). Medido en el PDF de Excel (38-comentarios): el cuadro lleva el relleno de la nota
// (por defecto #FFFFE1) y un borde negro de 0,75 pt; el texto (Tahoma 9 pt, línea de 11,76 pt) empieza 1,85 pt adentro y su primera línea base
// cae 9,36 pt bajo el borde de arriba; la flecha es un triángulo negro de 6,6 × 6 pt cuya punta queda 0,87 pt a la derecha y 1,08 pt bajo la
// esquina de arriba a la derecha de la celda, unido al cuadro por un hilo de 0,14 pt. El cuadro se coloca por su ancla como un objeto (escalado
// de píxeles de pantalla a la página). Sin DOM: usa la API de contexto 2D. Requiere xlsx-modelo.js, xlsx-geometria.js y xlsx-dibujo.js (fuenteCss).
//
// Uso:
//   XlsxNotas.dibujar(ctx, libro, hoja, g, { f0, c0, f1, c1 }, { x, y, fino, grosorRejilla, notas })   // las que se muestran + los triángulos (pantalla)
//   XlsxNotas.lista(hoja)           // [{ ref, tipo, lineas | entradas }] por filas, para "al final de la hoja"
//   XlsxNotas.filas(hoja, idioma)   // las filas de texto de esas páginas: [{ et, txt }]

(function () {
  const PUNTOS_LINEA = 11.76 / 9;   // alto de línea por punto de letra (Tahoma)
  // Las etiquetas del idioma de Excel (es medido; en es la lectura del nombre de Excel en inglés: sin contrastar)
  const ETIQUETAS = { es: { celda: 'Celda:', nota: 'Nota:', comentario: 'Comentario:', responder: 'Responder:' }, en: { celda: 'Cell:', nota: 'Note:', comentario: 'Comment:', responder: 'Reply:' } };

  const codigoCelda = (f, c) => LibroXlsx.numeroACol(c) + (f + 1);

  // El texto de una nota partido en líneas: los saltos propios y, dentro de cada uno, el ajuste a `ancho` (en la misma unidad que `medir`).
  function partir(runs, ancho, medir) {
    const lineas = [];
    let actual = [], w = 0;
    const cerrar = () => { lineas.push(actual); actual = []; w = 0; };
    for (const r of runs) {
      const partes = String(r.texto).replace(/\r\n?/g, '\n').split('\n');
      partes.forEach((trozo, i) => {
        if (i > 0) cerrar();
        for (const palabra of trozo.split(/(?<=\s)/)) {
          if (!palabra) continue;
          const pw = medir(palabra, r);
          if (w + pw > ancho && actual.length) cerrar();
          actual.push({ texto: palabra, run: r, w: pw }); w += pw;
        }
      });
    }
    if (actual.length || !lineas.length) cerrar();
    return lineas;
  }

  function dibujar(ctx, libro, hoja, g, r, op) {
    const notas = hoja.comentarios;
    if (!notas || !notas.length) return;
    op = op || {};
    const UPT = op.fino ? (op.grosorRejilla || 1.28) / 0.96 : 96 / 72;   // unidades del lienzo por punto
    const P = g.pantalla || g, ex = g.escalaX || 1, ey = g.escalaY || 1;
    const px = op.x || 0, py = op.y || 0;
    const sx0 = P.x(r.c0), sy0 = P.y(r.f0);
    const ancho = g.x(r.c1 + 1) - g.x(r.c0), alto = g.y(r.f1 + 1) - g.y(r.f0);
    const celdaX = (c) => px + g.x(c) - g.x(r.c0), celdaY = (f) => py + g.y(f) - g.y(r.f0);
    // triángulo rojo de las celdas con nota (solo en pantalla: al imprimir no sale)
    if (!op.fino && op.indicadores !== false) {
      ctx.save();
      ctx.fillStyle = '#FF0000';
      for (const n of notas) {
        if (n.f < r.f0 || n.f > r.f1 || n.c < r.c0 || n.c > r.c1 || g.altoFila(n.f) <= 0 || g.anchoCol(n.c) <= 0) continue;
        const x1 = celdaX(n.c + 1), y0 = celdaY(n.f), t = 6;
        ctx.beginPath(); ctx.moveTo(x1 - t, y0); ctx.lineTo(x1, y0); ctx.lineTo(x1, y0 + t); ctx.closePath(); ctx.fill();
      }
      ctx.restore();
    }
    if (op.notas === false) return;
    for (const n of notas) {
      if (!n.visible || !n.ancla) continue;
      const q = XlsxGeometria.rectanguloAncla(n.ancla, P);
      if (!q) continue;
      const x0 = px + (q.x0 - sx0) * ex, y0 = py + (q.y0 - sy0) * ey, w = (q.x1 - q.x0) * ex, h = (q.y1 - q.y0) * ey;
      if (!(w > 0) || !(h > 0) || x0 + w < px - 1 || x0 > px + ancho + 1 || y0 + h < py - 1 || y0 > py + alto + 1) continue;
      ctx.save();
      // flecha y hilo (la punta, junto a la esquina de arriba a la derecha de la celda)
      const tx = celdaX(n.c + 1) + 0.87 * UPT, ty = celdaY(n.f) + 1.08 * UPT;
      ctx.strokeStyle = '#000000'; ctx.lineWidth = 0.14 * UPT; ctx.lineCap = 'butt';
      ctx.beginPath(); ctx.moveTo(x0 + 0.11 * UPT, y0); ctx.lineTo(tx + 3.30 * UPT, ty - 2.46 * UPT); ctx.stroke();
      ctx.fillStyle = '#000000';
      ctx.beginPath(); ctx.moveTo(tx, ty); ctx.lineTo(tx + 3.12 * UPT, ty - 6.0 * UPT); ctx.lineTo(tx + 6.6 * UPT, ty - 1.9 * UPT); ctx.closePath(); ctx.fill();
      // cuadro
      ctx.fillStyle = n.relleno || '#FFFFE1'; ctx.fillRect(x0, y0, w, h);
      ctx.strokeStyle = n.borde || '#000000'; ctx.lineWidth = 0.75 * UPT; ctx.strokeRect(x0, y0, w, h);
      // texto
      ctx.beginPath(); ctx.rect(x0 + 0.34 * UPT, y0 + 0.34 * UPT, w - 0.68 * UPT, h - 0.68 * UPT); ctx.clip();
      const tam = (rn) => (rn.tam || 9) * UPT;
      const fuente = (rn) => XlsxDibujo.fuenteCss({ nombre: rn.fuente || 'Tahoma', negrita: rn.negrita, cursiva: rn.cursiva }, tam(rn));
      const medir = (t, rn) => { ctx.font = fuente(rn); return ctx.measureText(t).width; };
      // el ajuste del texto se guarda en la nota mientras no cambien el ancho ni la escala (se dibuja en cada cuadro)
      const clave = w.toFixed(3) + '|' + UPT.toFixed(5);
      if (!n._lineas || n._lineas.clave !== clave) n._lineas = { clave, lineas: partir(n.runs.length ? n.runs : [{ texto: n.texto, negrita: false, cursiva: false, tam: 9, fuente: 'Tahoma' }], w - 3.58 * UPT, medir) };
      const lineas = n._lineas.lineas;
      let by = y0 + 9.36 * UPT;
      ctx.textBaseline = 'alphabetic'; ctx.textAlign = 'left';
      for (const l of lineas) {
        let x = x0 + 1.85 * UPT;
        const mayor = l.reduce((a, t) => Math.max(a, t.run.tam || 9), 9);
        for (const t of l) {
          ctx.font = fuente(t.run); ctx.fillStyle = t.run.color ? (libro.estilos.color(t.run.color) || '#000000') : '#000000';
          ctx.fillText(t.texto, x, by);
          x += t.w;
        }
        by += mayor * PUNTOS_LINEA * UPT;
        if (by > y0 + h + 12 * UPT) break;
      }
      ctx.restore();
    }
  }

  // Lo que Excel imprime "al final de la hoja": por filas (y por columnas dentro de la fila), "Celda: B2" y la nota o el hilo, con una línea en blanco
  // entre una y otra. Devuelve [{ ref, tipo: 'nota' | 'hilo', lineas | entradas }].
  function lista(hoja) {
    const notas = (hoja.comentarios || []).slice().sort((x, y) => x.f - y.f || x.c - y.c);
    return notas.map((n) => (n.hilo
      ? { ref: codigoCelda(n.f, n.c), tipo: 'hilo', entradas: n.hilo.entradas.map((e) => ({ autor: e.autor, fecha: e.fecha, respuesta: e.respuesta, lineas: String(e.texto).replace(/\r\n?/g, '\n').split('\n') })) }
      : { ref: codigoCelda(n.f, n.c), tipo: 'nota', lineas: String(n.texto).replace(/\r\n?/g, '\n').split('\n') }));
  }

  // Las filas de texto que se imprimen para esa lista: { et: 'celda' | 'nota' | 'comentario' | 'responder' | null, txt } (et: la etiqueta de esa fila).
  function filas(hoja, idioma) {
    const out = [];
    const fecha = (iso) => {
      if (!iso) return '';
      const d = new Date(/[zZ]|[+-]\d\d:?\d\d$/.test(iso) ? iso : iso + 'Z');   // Excel guarda la hora en UTC y la muestra en la hora local
      if (!Number.isFinite(d.getTime())) return '';
      const dos = (v) => String(v).padStart(2, '0');
      return idioma === 'en'
        ? `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()} ${d.getHours() % 12 || 12}:${dos(d.getMinutes())} ${d.getHours() < 12 ? 'AM' : 'PM'}`
        : `${dos(d.getDate())}/${dos(d.getMonth() + 1)}/${d.getFullYear()} ${dos(d.getHours())}:${dos(d.getMinutes())}`;
    };
    for (const n of lista(hoja)) {
      out.push({ et: 'celda', txt: n.ref });
      if (n.tipo === 'nota') n.lineas.forEach((l, i) => out.push({ et: i === 0 ? 'nota' : null, txt: l }));
      else for (const e of n.entradas) {
        out.push({ et: e.respuesta ? 'responder' : 'comentario', txt: e.autor });
        for (const l of e.lineas) out.push({ et: null, txt: l });
        out.push({ et: null, txt: fecha(e.fecha) });
      }
      out.push({ et: null, txt: '' });
    }
    return out;
  }

  globalThis.XlsxNotas = { dibujar, lista, filas, etiquetas: (idioma) => ETIQUETAS[idioma === 'en' ? 'en' : 'es'] };
})();

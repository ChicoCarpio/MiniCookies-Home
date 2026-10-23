/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Un libro .xlsx/.xlsm pasado al contrato de celdas que entienden los escritores de formatos antiguos (XlsEscritor, y después el de .ods): el mismo que devuelven
// XlsLector, OdsLector y XlsbLector, para que cualquier lector → cualquier escritor se haga sin más adaptadores. Requiere xlsx-modelo.js (y xlsx-estilos.js).
//
//   const libro = await XlsxContrato.leer(bytes);
//   libro = { hojas: [{ nombre, filas: [[valor…]…], anchos, combinadas: ['A1:B2'…], estado, columnasOcultas, filasOcultas, altosFila }], fecha1904, nombres: [{ nombre, hoja?, formula }] }
//   valor: número | texto | lógico | { error } | { valor, formato?, estilo?, formula?, matriz? }   (ver XlsEscritor)
//
// Se llevan valores, fórmulas (las compartidas ya expandidas), formato de número, fuente, relleno, bordes, alineación, combinadas, anchos, altos, filas y columnas ocultas,
// hojas ocultas y nombres definidos. Las hojas de gráfico se omiten. Lo demás (dibujos, tablas, formato condicional…) no tiene destino en un .xls ni en un .ods de este motor.
(function () {
  const TIPO_VACIO = 0, TIPO_NUMERO = 1, TIPO_TEXTO = 2, TIPO_LOGICO = 3, TIPO_ERROR = 4;
  const MAX_COLS_ANCHO = 256;

  async function leer(bytes) {
    const libro = await LibroXlsx.abrir(bytes);
    const quitarAlmohadilla = (c) => (typeof c === 'string' && c[0] === '#' ? c.slice(1) : c);
    const colorDe = (ref) => { const c = libro.estilos.color(ref); return c ? quitarAlmohadilla(c).toUpperCase() : null; };

    // El estilo de cada xf se arma una vez y lo comparten todas sus celdas (los escritores reparten los iguales por identidad).
    const memo = new Map();
    function estiloDe(indice) {
      if (memo.has(indice)) return memo.get(indice);
      const x = libro.estilos.xf(indice);
      const r = { formato: x.formato && x.formato !== 'General' ? x.formato : null, estilo: null };
      const est = {};
      const f = x.fuente;
      if (f && (f.nombre || f.tam)) {
        const ft = { nombre: f.nombre || 'Calibri', tam: f.tam || 11 };
        if (f.negrita) ft.negrita = true;
        if (f.cursiva) ft.cursiva = true;
        if (f.tachado) ft.tachado = true;
        if (f.subrayado) ft.subrayado = /double/i.test(f.subrayado) ? 'double' : true;
        const c = colorDe(f.color); if (c && c !== '000000') ft.color = c;
        est.fuente = ft;
      }
      const rl = x.relleno;
      if (rl && rl.tipo === 'patron' && rl.patron !== 'none') {
        const fg = colorDe(rl.fg), bg = colorDe(rl.bg);
        if (rl.patron === 'solid') { if (fg) est.relleno = fg; }
        else est.relleno = { patron: rl.patron, ...(fg ? { fg } : {}), ...(bg ? { bg } : {}) };
      } else if (rl && rl.tipo === 'degradado' && rl.degradado && rl.degradado.paradas && rl.degradado.paradas[0]) {
        const c = colorDe(rl.degradado.paradas[0].color); if (c) est.relleno = c;
      }
      if (x.borde) {
        const bordes = {};
        for (const [lado, clave] of [['izq', 'izq'], ['der', 'der'], ['arr', 'arr'], ['aba', 'aba']]) {
          const b = x.borde[lado];
          if (b && b.estilo && b.estilo !== 'none') { const c = colorDe(b.color); bordes[clave] = { e: b.estilo, ...(c && c !== '000000' ? { c } : {}) }; }
        }
        if (Object.keys(bordes).length) est.bordes = bordes;
      }
      const a = x.alineacion;
      if (a) {
        const al = {};
        if (a.h && a.h !== 'general') al.h = a.h;
        if (a.v && a.v !== 'bottom') al.v = a.v;
        if (a.ajustar) al.ajustar = true;
        if (a.sangria) al.sangria = a.sangria;
        if (a.rotacion) al.rotacion = a.rotacion;
        if (Object.keys(al).length) est.alinea = al;
      }
      if (Object.keys(est).length) r.estilo = est;
      memo.set(indice, r);
      return r;
    }

    const hojas = [], indicePorHoja = new Map();
    for (let i = 0; i < libro.hojas.length; i++) {
      const info = libro.hojas[i];
      if (info.tipo && info.tipo !== 'hoja') continue;   // hojas de gráfico, de macros…
      const hoja = await libro.cargarHoja(i);
      const celdas = hoja.celdas;
      const filas = [];
      celdas.paraCada((f, c, k) => {
        const tipo = celdas.tipo[k];
        const e = celdas.estilo[k] ? estiloDe(celdas.estilo[k]) : null;
        const formula = celdas.formulaDe(k);
        let valor;
        if (tipo === TIPO_NUMERO || tipo === TIPO_TEXTO) valor = celdas.valor[k];
        else if (tipo === TIPO_LOGICO) valor = !!celdas.valor[k];
        else if (tipo === TIPO_ERROR) valor = { error: celdas.valor[k] };
        else valor = null;
        if (tipo === TIPO_VACIO && !formula && !(e && (e.estilo))) return;
        let celda = valor;
        if (formula || (e && (e.estilo || e.formato))) {
          celda = { valor };
          if (e && e.formato) celda.formato = e.formato;
          if (e && e.estilo) celda.estilo = e.estilo;
          if (formula) {
            celda.formula = formula;
            const fx = celdas.formulas.get(k);
            if (fx && fx.tipo === 'array' && fx.ref) celda.matriz = fx.ref;
          }
        }
        (filas[f] || (filas[f] = []))[c] = celda;
      });
      const anchos = [], columnasOcultas = [];
      for (const col of hoja.columnas) for (let c = col.min; c <= col.max && c < MAX_COLS_ANCHO; c++) { if (col.ancho) anchos[c] = col.ancho; if (col.oculta) columnasOcultas.push(c); }
      const filasOcultas = [], altosFila = {};
      for (const [f, fi] of hoja.filas) { if (fi.oculta) filasOcultas.push(f); if (fi.personalizada && fi.alto) altosFila[f] = fi.alto; }
      const combinadas = hoja.combinadas.map((r) => LibroXlsx.numeroACol(r.c0) + (r.f0 + 1) + ':' + LibroXlsx.numeroACol(r.c1) + (r.f1 + 1));
      indicePorHoja.set(i, hojas.length);
      hojas.push({ nombre: info.nombre, filas: Array.from(filas, (f) => f || []), anchos, combinadas, estado: info.estado, columnasOcultas, filasOcultas, altosFila });
    }
    const nombres = [];
    for (const n of libro.nombres || []) {
      if (!n.nombre || !n.ref || /^_xl/i.test(n.nombre)) continue;   // los integrados (_xlnm.Print_Area…) no se llevan
      const o = { nombre: n.nombre, formula: n.ref.replace(/^=/, '') };
      if (n.hojaLocal !== null && n.hojaLocal !== undefined) { if (!indicePorHoja.has(n.hojaLocal)) continue; o.hoja = indicePorHoja.get(n.hojaLocal); }
      nombres.push(o);
    }
    return { hojas, fecha1904: !!libro.fecha1904, nombres };
  }

  globalThis.XlsxContrato = { leer };
})();

/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Copiar y pegar de Hojas: qué llega a cada celda con cada tipo de pegado de Excel (Todo, Fórmulas, Valores, Formatos, Notas, Validación, Todo excepto bordes, Anchos, Fórmulas y Valores con formato de
// números, Combinar formato condicional), las operaciones, omitir blancos, transponer, pegar vínculo y la repetición sobre una selección mayor. Medido en Excel: pruebas/oraculo/excel-pegado-especial.ps1.
// Lo que no es una celda (reglas, validación, notas, enlaces, combinadas, anchos, tablas) lo escribe pegado-objetos.js sobre el archivo, y solo si hay algo que escribir.
//   const P = PegadoEspecial.capturar(ed, h, rango)                     lo copiado: { h, r, filas, celdas, ed, colEntera, filaEntera }
//   const plan = PegadoEspecial.planear(ed, P, h, seleccion, opciones)  → { h, celdas, destino, objetos } | { error }
//      opciones = { tipo, op: 'sumar'|'restar'|'multiplicar'|'dividir'|null, blancos, trans, vinculo, brocha }
//   await PegadoEspecial.ejecutar(ed, plan, { idioma }) → bytes del archivo con todo aplicado
(function () {
  // contenido: 'formula' | 'valor'; formato: 'todo' | 'sinBordes' | 'numero'; cf: 'reemplaza' | 'suma'; dv, notas, enlaces, comb, anchos: sí o no
  const TIPOS = {
    todo: { contenido: 'formula', formato: 'todo', cf: 'reemplaza', dv: 1, notas: 1, enlaces: 1, comb: 1, tabla: 1 },
    formulas: { contenido: 'formula' },
    valores: { contenido: 'valor' },
    formatos: { formato: 'todo', cf: 'reemplaza', comb: 1 },
    notas: { notas: 1 },
    validacion: { dv: 1 },
    sinBordes: { contenido: 'formula', formato: 'sinBordes', cf: 'reemplaza', dv: 1, notas: 1, enlaces: 1, comb: 1, tabla: 1 },
    anchos: { anchos: 1 },
    formulasNum: { contenido: 'formula', formato: 'numero' },
    valoresNum: { contenido: 'valor', formato: 'numero' },
    fusionCF: { contenido: 'formula', formato: 'todo', cf: 'suma', dv: 1, notas: 1, enlaces: 1, comb: 1, tabla: 1 },
  };
  const FILAS = 1048576, COLS = 16384, MAX_COPIA = 200000;
  const estaVacia = (x) => !x || (x.tipo === 'vacio' && !x.formula);
  const toca = (a, b) => a.f0 <= b.f1 && b.f0 <= a.f1 && a.c0 <= b.c1 && b.c0 <= a.c1;
  const dentro = (a, b) => a.f0 >= b.f0 && a.f1 <= b.f1 && a.c0 >= b.c0 && a.c1 <= b.c1;

  // Con un filtro puesto solo se copian las filas que se ven; una fila oculta a mano se copia igual (medido)
  function filasDeCopia(hoja, r, f1) {
    const rangos = [];
    if (hoja.filtroAuto && hoja.filtroAuto.columnas && hoja.filtroAuto.columnas.size) { const g = LibroXlsx.decodificarRango(hoja.filtroAuto.ref || ''); if (g) rangos.push(g); }
    for (const t of hoja.tablas || []) if (t.filtro && t.filtro.columnas && t.filtro.columnas.size) rangos.push(t.ref);
    const l = [];
    for (let f = r.f0; f <= f1; f++) {
      const x = hoja.filas && hoja.filas.get(f);
      if (x && x.oculta && rangos.some((g) => f >= g.f0 && f <= g.f1)) continue;
      l.push(f);
    }
    return l.length ? l : [r.f0];
  }
  function capturar(ed, h, r) {
    const lim = ed.limite(h), hoja = ed.hojas[h];
    const colEntera = r.f0 === 0 && r.f1 >= FILAS - 1, filaEntera = r.c0 === 0 && r.c1 >= COLS - 1;
    // lo seleccionado cuenta entero aunque esté vacío, salvo que sea enorme: entonces hasta donde hay datos
    const chico = (r.f1 - r.f0 + 1) * (r.c1 - r.c0 + 1) <= MAX_COPIA;
    const f1 = chico ? r.f1 : Math.min(r.f1, Math.max(r.f0, lim.filas - 1)), c1 = chico ? r.c1 : Math.min(r.c1, Math.max(r.c0, lim.cols - 1));
    const filas = filasDeCopia(hoja, r, f1), celdas = [];
    for (const f of filas) { const fila = []; for (let c = r.c0; c <= c1; c++) fila.push(ed.estadoDe(h, f, c)); celdas.push(fila); }
    return { h, r: { f0: r.f0, c0: r.c0, f1, c1 }, filas, celdas, ed, colEntera, filaEntera };
  }

  // ---- transponer una fórmula ------------------------------------------------------------------------------------------------------------------------------------------------------------
  // Una referencia a lo copiado va a donde quedó esa celda (también si es absoluta). Fuera de lo copiado, la relativa cambia filas por columnas y la que tiene algo absoluto se queda (medido).
  function transponerFormula(texto, X) {
    const { sf, sc, df, dc, S, f0d, c0d } = X;
    const mapa = (fila, col) => ({ f: f0d + (col - S.c0), c: c0d + (fila - S.f0) });
    const adentro = (fila, col) => fila >= S.f0 && fila <= S.f1 && col >= S.c0 && col <= S.c1;
    const punto = (a, cc, b, ff, citada) => {   // $ col $ fila → texto | null
      const fila = ff - 1, col = LibroXlsx.colANumero(cc.toUpperCase());
      if (!citada && adentro(fila, col)) { const m = mapa(fila, col); return a + LibroXlsx.numeroACol(m.c) + b + (m.f + 1); }
      if (a || b) return a + cc.toUpperCase() + b + ff;
      const nf = df + (col - sc), nc = dc + (fila - sf);
      return nf < 0 || nf >= FILAS || nc < 0 || nc >= COLS ? null : LibroXlsx.numeroACol(nc) + (nf + 1);
    };
    let salida = '', i = 0;
    const n = texto.length, esIdent = (ch) => /[A-Za-z0-9_.À-￿]/.test(ch);
    while (i < n) {
      const ch = texto[i];
      if (ch === '"' || ch === "'") {
        let j = i + 1;
        while (j < n) { if (texto[j] === ch) { if (texto[j + 1] === ch) j += 2; else break; } else j++; }
        salida += texto.slice(i, j + 1); i = j + 1; continue;
      }
      if (ch === '[') {
        let nivel = 0, j = i;
        for (; j < n; j++) { if (texto[j] === '[') nivel++; else if (texto[j] === ']') { nivel--; if (nivel === 0) break; } }
        salida += texto.slice(i, j + 1); i = j + 1; continue;
      }
      if (/[$A-Za-z0-9]/.test(ch) && (i === 0 || !esIdent(texto[i - 1])) && texto[i - 1] !== '$') {
        const citada = texto[i - 1] === '!';
        const resto = texto.slice(i, i + 40);
        let m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+):(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(.!])/.exec(resto);
        if (m) {
          const a = punto(m[1], m[2], m[3], m[4], citada), b = punto(m[5], m[6], m[7], m[8], citada);
          salida += a && b ? a + ':' + b : '#REF!';
          i += m[0].length; continue;
        }
        m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(.!:])/.exec(resto);
        if (m && LibroXlsx.colANumero(m[2].toUpperCase()) < COLS) { salida += punto(m[1], m[2], m[3], m[4], citada) || '#REF!'; i += m[0].length; continue; }
        // columnas y filas enteras: sus dos puntas, con lo que se sale de la hoja hasta el borde
        m = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})(?![A-Za-z0-9_(.!])/.exec(resto);
        if (m) {
          const fila = df + (LibroXlsx.colANumero(m[1].toUpperCase()) - sc);
          if (fila < 0 || fila >= FILAS) { salida += '#REF!'; i += m[0].length; continue; }
          const ini = Math.max(0, dc - sf);
          salida += '$' + LibroXlsx.numeroACol(ini) + (fila + 1) + ':$XFD' + (fila + 1);
          i += m[0].length; continue;
        }
        m = /^(\d+):(\d+)(?![A-Za-z0-9_(.!])/.exec(resto);
        if (m) {
          const col = dc + (+m[1] - 1 - sf);
          if (col < 0 || col >= COLS) { salida += '#REF!'; i += m[0].length; continue; }
          const ini = Math.max(0, df - sc), t = LibroXlsx.numeroACol(col);
          salida += t + '$' + (ini + 1) + ':' + t + '$1048576';
          i += m[0].length; continue;
        }
        if (/[A-Za-z$]/.test(ch)) { let j = i + 1; while (j < n && (esIdent(texto[j]) || texto[j] === '$')) j++; salida += texto.slice(i, j); i = j; continue; }
      }
      salida += ch; i++;
    }
    return salida;
  }

  // ---- operaciones -----------------------------------------------------------------------------------------------------------------------------------------------------------------------
  const SIGNO = { sumar: '+', restar: '-', multiplicar: '*', dividir: '/' };
  function calcular(op, a, b) {
    if (op === 'dividir' && b === 0) return { tipo: 'error', valor: '#DIV/0!', formula: null };
    const v = op === 'sumar' ? a + b : op === 'restar' ? a - b : op === 'multiplicar' ? a * b : a / b;
    return { tipo: 'numero', valor: v, formula: null };
  }
  // Lo que queda en una celda del destino (d) con la operación. Solo opera con números y con lo vacío, que vale 0; texto, lógicos y errores se quedan. → null si no cambia
  function operar(op, d, x, formulaOrigen) {
    const s = SIGNO[op];
    const dVacia = estaVacia(d);
    if (!dVacia && !d.formula && d.tipo !== 'numero') return null;
    if (formulaOrigen) {
      const izq = dVacia ? '0' : d.formula ? '(' + d.formula + ')' : String(d.valor);
      return { tipo: 'numero', valor: 0, formula: izq + s + '(' + formulaOrigen + ')' };
    }
    const xVacia = estaVacia(x);
    if (!xVacia && x.tipo !== 'numero') return null;
    if (dVacia && xVacia) return null;
    const n = xVacia ? 0 : x.valor;
    if (!dVacia && d.formula) return { tipo: 'numero', valor: 0, formula: '(' + d.formula + ')' + s + n };
    return calcular(op, dVacia ? 0 : d.valor, n);
  }

  // ---- el plan -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  const arrayDe = (hoja) => { const l = []; for (const fx of hoja.celdas.formulas.values()) if (fx.tipo === 'array' && (!fx.dinamica || fx.matrizOds) && fx.ref) { const g = LibroXlsx.decodificarRango(fx.ref); if (g) l.push(g); } return l; };
  function geometria(P, s, o, T) {
    const alto = P.celdas.length, ancho = P.celdas[0].length;
    const A = o.trans ? ancho : alto, B = o.trans ? alto : ancho;
    const sa = s.f1 - s.f0 + 1, sb = s.c1 - s.c0 + 1;
    if (P.colEntera && (s.f0 !== 0 || o.trans)) return { error: 'limite' };
    if (P.filaEntera && (s.c0 !== 0 || o.trans)) return { error: 'limite' };
    const altoCopia = P.colEntera ? FILAS : A, anchoCopia = P.filaEntera ? COLS : B;
    const una = sa === 1 && sb === 1;
    const multiplo = sa % altoCopia === 0 && sb % anchoCopia === 0;
    let dAlto = altoCopia, dAncho = anchoCopia, ciclo = false, rellena = false;
    if (o.vinculo) { if (!una) { dAlto = sa; dAncho = sb; rellena = true; } }
    else if (multiplo) { dAlto = sa; dAncho = sb; }
    else if (o.brocha && !una) { dAlto = sa; dAncho = sb; ciclo = true; }
    else if (o.tipo === 'formatos' && sa >= altoCopia && sb >= anchoCopia) { dAlto = sa; dAncho = sb; ciclo = true; }
    const D = { f0: s.f0, c0: s.c0, f1: s.f0 + dAlto - 1, c1: s.c0 + dAncho - 1 };
    if (D.f1 >= FILAS || D.c1 >= COLS) return { error: 'limite' };
    void T; void ciclo;
    return { D, A, B, rellena, repite: dAlto > altoCopia || dAncho > anchoCopia };
  }
  // los pares origen → destino: (i, j) en P.celdas y (f, c) en la hoja. Las filas o columnas enteras solo hasta donde hay datos.
  function paresDe(P, G, o) {
    const pares = [], { D, A, B } = G, alto = P.celdas.length, ancho = P.celdas[0].length;
    const f1 = P.colEntera ? D.f0 + alto - 1 : D.f1, c1 = P.filaEntera ? D.c0 + ancho - 1 : D.c1;
    for (let f = D.f0; f <= f1; f++) for (let c = D.c0; c <= c1; c++) {
      const a = (f - D.f0) % A, b = (c - D.c0) % B;
      pares.push({ i: o.trans ? b : a, j: o.trans ? a : b, f, c });
    }
    return pares;
  }
  function estiloOrigen(P, x, i, j) { return x ? x.estilo || 0 : (P.ed.hojas[P.h].estiloDe ? P.ed.hojas[P.h].estiloDe(P.filas[i], P.r.c0 + j) : 0); }
  // de otro libro abierto antes: los números de estilo no valen aquí, se trae lo que se ve
  function memoEstilos(ed, P) {
    const memo = new Map();
    return (e) => {
      if (P.ed === ed || !e) return P.ed === ed ? e : 0;
      if (!memo.has(e)) { const p = P.ed.estilos.propiedades(e), c = {}; for (const k of ['negrita', 'cursiva', 'subrayado', 'tachado', 'colorFuente', 'tamano', 'fuente', 'relleno', 'h', 'v', 'ajustar', 'formato']) c[k] = p[k]; memo.set(e, ed.estilos.derivar(0, c)); }
      return memo.get(e);
    };
  }

  function citaOrigen(ed, P, h) { return P.h === h ? '' : XlsFormulas._citarHoja(ed.libro.hojas[P.h].nombre) + '!'; }

  function planear(ed, P, h, seleccion, opciones) {
    const o = Object.assign({ tipo: 'todo' }, opciones || {}), T = TIPOS[o.tipo] || TIPOS.todo, hoja = ed.hojas[h], mismoLibro = P.ed === ed;
    const G = geometria(P, seleccion, o, T);
    if (G.error) return G;
    let D = G.D;
    if (o.trans && P.h === h && mismoLibro && toca(P.r, D)) return { error: 'solape' };
    const conContenido = !!T.contenido || o.vinculo;
    const prot = hoja.proteccion && hoja.proteccion.hoja;
    const bloqueada = (f, c) => { if (!prot) return false; const x = ed.estadoDe(h, f, c), e = x ? x.estilo : (hoja.estiloDe ? hoja.estiloDe(f, c) : 0), xf = ed.libro.estilos.xf(e); return !(xf && xf.proteccion && xf.proteccion.bloqueada === false); };
    // combinadas en el destino
    const combinadas = hoja.combinadas || [];
    const unaCelda = P.celdas.length === 1 && P.celdas[0].length === 1;
    const tocadas = combinadas.filter((m) => toca(m, D));
    let sobreCombinada = false;   // una sola celda sobre una combinada: llega su contenido, el formato de la combinada se queda
    if (tocadas.length && o.tipo !== 'anchos' && o.tipo !== 'notas' && o.tipo !== 'validacion') {
      if (T.contenido === 'valor') return { error: 'combinada' };
      if (unaCelda && !o.vinculo && tocadas.length === 1 && dentro(D, tocadas[0])) { D = { f0: tocadas[0].f0, c0: tocadas[0].c0, f1: tocadas[0].f1, c1: tocadas[0].c1 }; sobreCombinada = T.contenido === 'formula' && T.formato === 'todo'; }
      else if (tocadas.some((m) => !dentro(m, D))) return { error: 'combinada' };
    }
    // una matricial no se parte
    const fijas = arrayDe(hoja);
    if (conContenido && fijas.some((g) => toca(g, D) && !dentro(g, D) && !(o.vinculo && false))) return { error: o.vinculo ? 'vinculo' : 'matriz' };
    // los encabezados y totales de una tabla, y lo que queda a su lado (donde crece), los pone al día MoverTablas al guardar; salvo una tabla que queda entera debajo de otra que llega.
    // Pegar un vínculo ahí no se hace
    const tablaEntera = mismoLibro && T.tabla && (ed.hojas[P.h].tablas || []).some((t) => dentro(t.ref, P.r));
    if (o.vinculo) for (const t of hoja.tablas || []) {
      if (t.ref.c0 > D.c1 || t.ref.c1 < D.c0) continue;
      if ((t.encabezados && t.ref.f0 >= D.f0 && t.ref.f0 <= D.f1) || (t.totales && t.ref.f1 >= D.f0 && t.ref.f1 <= D.f1)) return { error: 'protegida' };
    }
    const conTablas = conContenido && !o.vinculo && !tablaEntera && MoverTablas.toca(ed, h, D, unaCelda && T.contenido === 'formula' && !!(P.celdas[0][0] && P.celdas[0][0].formula));
    if (conContenido && prot) for (let f = D.f0; f <= Math.min(D.f1, D.f0 + 2000); f++) for (let c = D.c0; c <= Math.min(D.c1, D.c0 + 200); c++) if (bloqueada(f, c)) return { error: 'protegida' };
    const G2 = Object.assign({}, G, { D });
    const pares = o.vinculo ? [] : paresDe(P, G2, o);
    if (pares.length > 500000) return { error: 'grande' };
    const celdas = [];
    const traerEstilo = memoEstilos(ed, P);
    const estiloDe = (f, c, d) => (d ? d.estilo || 0 : (hoja.estiloDe ? hoja.estiloDe(f, c) : 0));
    const S = { f0: P.filas[0], f1: P.filas[P.filas.length - 1], c0: P.r.c0, c1: P.r.c1 };

    if (o.vinculo) {
      const cita = citaOrigen(ed, P, h), corner = P.celdas[0][0], unaSola = unaCelda;
      const refUna = '$' + LibroXlsx.numeroACol(P.r.c0) + '$' + (P.filas[0] + 1);
      const fNum = corner ? ed.libro.estilos.xf(corner.estilo || 0).formato : 'General';
      for (let f = D.f0; f <= (P.colEntera ? D.f0 + P.celdas.length - 1 : D.f1); f++) for (let c = D.c0; c <= (P.filaEntera ? D.c0 + P.celdas[0].length - 1 : D.c1); c++) {
        const a = f - D.f0, b = c - D.c0;
        const fo = a < P.filas.length ? P.filas[a] : P.filas[P.filas.length - 1] + (a - P.filas.length + 1), co = P.r.c0 + b;
        const d = ed.estadoDe(h, f, c);
        let est = estiloDe(f, c, d);
        if (fNum && fNum !== 'General' && (ed.libro.estilos.xf(est).formato || 'General') === 'General' && mismoLibro) est = ed.estilos.derivar(est, { formato: fNum });
        const ref = unaSola ? refUna : LibroXlsx.numeroACol(co) + (fo + 1);
        celdas.push({ f, c, v: { tipo: 'numero', valor: 0, formula: cita + ref, estilo: est } });
      }
      return { h, celdas, destino: D, objetos: null, tipo: o };
    }

    const memoSinBordes = new Map(), memoNumero = new Map();
    const conBordesDe = (e, d) => { const k = e + '|' + d; if (!memoSinBordes.has(k)) memoSinBordes.set(k, ed.estilos.derivar(e, { bordeDe: d })); return memoSinBordes.get(k); };
    const conNumero = (d, cod) => { const k = d + '|' + cod; if (!memoNumero.has(k)) memoNumero.set(k, ed.estilos.derivar(d, { formato: cod })); return memoNumero.get(k); };
    const tabla = (T.tabla && mismoLibro && !o.blancos) ? tablasCopiadas(ed, P, D, G, o) : null;
    if (tabla && tabla.error) return tabla;

    const paresObj = [];
    const combOrigen = P.ed.hojas[P.h].combinadas || [];
    const horneado = T.formato ? estilosDeTabla(ed, P) : null;
    for (const p of pares) {
      if (sobreCombinada && (p.f !== D.f0 || p.c !== D.c0)) continue;
      const x = P.celdas[p.i][p.j], sf0 = P.filas[p.i], sc0 = P.r.c0 + p.j;
      const blanca = (estaVacia(x) || !!(x && x.derrame && !T.contenido)) && !combOrigen.some((m) => sf0 >= m.f0 && sf0 <= m.f1 && sc0 >= m.c0 && sc0 <= m.c1);
      paresObj.push({ sf: P.filas[p.i], sc: P.r.c0 + p.j, df: p.f, dc: p.c, blanca });
      if (o.blancos && blanca) continue;
      if (!T.contenido && !T.formato) continue;
      const d = ed.estadoDe(h, p.f, p.c), deProt = prot;
      const estDest = estiloDe(p.f, p.c, d);
      let estOrig = traerEstilo(estiloOrigen(P, x, p.i, p.j));
      if (horneado && !(tabla && tabla.lista.some((l) => dentro({ f0: sf0, c0: sc0, f1: sf0, c1: sc0 }, l.t.ref)))) estOrig = horneado(sf0, sc0, estOrig);
      let est = estDest;
      if (T.formato === 'todo' && !deProt) est = sobreCombinada ? estDest : estOrig;
      else if (T.formato === 'sinBordes' && !deProt) est = conBordesDe(estOrig, estDest);
      else if (T.formato === 'numero' && !deProt) est = conNumero(estDest, (ed.libro.estilos.xf(estOrig) || {}).formato || 'General');
      let v;
      if (T.contenido) {
        const sf = P.filas[p.i], sc = P.r.c0 + p.j;
        // de un desborde: la fórmula solo en el ancla (vuelve a desbordar); sus otras celdas quedan vacías, con su formato
        const hija = !!(x && x.derrame);
        let fx = null;
        if (T.contenido === 'formula' && x && x.formula && !hija) {
          fx = o.trans ? transponerFormula(x.formula, { sf, sc, df: p.f, dc: p.c, S, f0d: D.f0, c0d: D.c0 }) : LibroXlsx.desplazarFormula(x.formula, p.f - sf, p.c - sc);
          if (tabla && tabla.renombrar) fx = tabla.renombrar(fx, p.f, p.c);
        }
        if (o.op) {
          const r = operar(o.op, d, T.contenido === 'valor' || !fx ? (hija && T.contenido === 'formula' ? null : x) : x, fx);
          if (r) v = Object.assign(r, { estilo: est });
          else if (est === estDest) continue;
          else v = d ? Object.assign({}, d, { estilo: est }) : { tipo: 'vacio', valor: null, formula: null, estilo: est };
        } else if (T.contenido === 'formula') {
          if (!x || (x.tipo === 'vacio' && !x.formula) || (hija && !fx)) v = est ? { tipo: 'vacio', valor: null, formula: null, estilo: est } : null;
          else {
            v = Object.assign({}, x, { estilo: est });
            delete v.matriz; delete v.derrame;
            if (fx !== null) v.formula = fx;
            if (x.fija) {
              const g = LibroXlsx.decodificarRango(x.fija), entera = g && S.f0 <= g.f0 && S.f1 >= g.f1 && S.c0 <= g.c0 && S.c1 >= g.c1 && !o.trans;
              delete v.fija;
              if (g && entera) v.fija = LibroXlsx.numeroACol(g.c0 + p.c - sc) + (g.f0 + p.f - sf + 1) + (g.f0 === g.f1 && g.c0 === g.c1 ? '' : ':' + LibroXlsx.numeroACol(g.c1 + p.c - sc) + (g.f1 + p.f - sf + 1));
              else if (g && !o.trans) v.fija = LibroXlsx.numeroACol(p.c) + (p.f + 1);
            }
          }
        } else {
          v = x && x.tipo !== 'vacio' ? { tipo: x.tipo, valor: x.valor, formula: null, estilo: est } : est ? { tipo: 'vacio', valor: null, formula: null, estilo: est } : null;
        }
      } else {
        v = d ? Object.assign({}, d, { estilo: est }) : est ? { tipo: 'vacio', valor: null, formula: null, estilo: est } : null;
        if (v) delete v.matriz;
      }
      celdas.push({ f: p.f, c: p.c, v });
    }

    const objetos = mismoLibro && !prot ? PegadoObjetos.planear(ed, P, h, D, T, o, paresObj, G2, tabla) : null;
    return { h, celdas, destino: D, objetos, tipo: o, tablas: conTablas };
  }

  // Pegar el formato de una celda de una tabla trae lo que se ve: lo que le da el estilo de la tabla pasa a ser formato propio (medido)
  function estilosDeTabla(ed, P) {
    if (P.ed !== ed || ed.formato === 'ods') return null;
    const t = globalThis.XlsxTablas ? XlsxTablas.preparar(ed.libro, ed.hojas[P.h]) : null;
    if (!t) return null;
    const memo = new Map(), E = ed.libro.estilos;
    const hex = (ref) => { const c = ref ? E.color(ref) : null; return c ? c.slice(1).toUpperCase() : null; };
    const lado = (b) => (b && b.estilo && b.estilo !== 'none' ? { estilo: b.estilo, color: hex(b.color) } : null);
    return (f, c, est) => {
      const e = t.para(f, c);
      if (!e) return est;
      const k = est + '|' + JSON.stringify(e);
      if (memo.has(k)) return memo.get(k);
      const x = XlsxTablas.aplicar(ed.libro, E.xf(est), e), q = {};
      if (x.fuente) { if (x.fuente.negrita) q.negrita = true; if (x.fuente.cursiva) q.cursiva = true; const co = hex(x.fuente.color); if (co) q.colorFuente = co; }
      if (x.relleno && x.relleno.patron === 'solid') { const r = hex(x.relleno.fg); if (r) q.relleno = r; }
      if (x.borde) { const b = {}; for (const n of ['izq', 'der', 'arr', 'aba']) { const l = lado(x.borde[n]); if (l) b[n] = l; } if (Object.keys(b).length) q.borde = b; }
      const i = ed.estilos.derivar(est, q);
      memo.set(k, i);
      return i;
    };
  }

  // Una tabla copiada entera (con su encabezado) nace otra vez en el destino; transpuesta o repetida, Excel no lo deja (medido)
  function tablasCopiadas(ed, P, D, G, o) {
    const de = (ed.hojas[P.h].tablas || []).filter((t) => dentro(t.ref, P.r) && t.encabezados);
    if (!de.length) return null;
    if (o.trans || G.repite) return { error: 'limite' };
    const df = D.f0 - P.r.f0, dc = D.c0 - P.r.c0, usados = new Set(), lista = [];
    let total = 0;
    ed.hojas.forEach((x) => { for (const t of (x && x.tablas) || []) { usados.add(String(t.nombreVisible || t.nombre).toLowerCase()); total++; } });
    for (const n of ed.libro.nombres) usados.add(n.nombre.toLowerCase());
    for (const t of de) {
      const base = t.nombreVisible || t.nombre, sep = base.includes('_') ? '' : '_';
      let k = total + 2;
      while (usados.has((base + sep + k).toLowerCase())) k++;
      usados.add((base + sep + k).toLowerCase());
      lista.push({ t, nombre: base + sep + k, ref: { f0: t.ref.f0 + df, c0: t.ref.c0 + dc, f1: t.ref.f1 + df, c1: t.ref.c1 + dc } });
      total++;
    }
    const renombrar = (fx, f, c) => {
      let u = fx;
      for (const l of lista) if (f >= l.ref.f0 && f <= l.ref.f1 && c >= l.ref.c0 && c <= l.ref.c1) u = u.replace(new RegExp('(^|[^A-Za-z0-9_.])' + (l.t.nombreVisible || l.t.nombre).replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\[', 'g'), (m, a) => a + l.nombre + '[');
      return u;
    };
    return { lista, renombrar };
  }

  async function ejecutar(ed, plan, opciones) {
    if (plan.celdas.length) ed.aplicar({ t: 'celdas', h: plan.h, celdas: plan.celdas });
    ed.terminar();
    if (plan.objetos && plan.objetos.tapadas && plan.objetos.tapadas.length) ed.hojas[plan.h].tablas = (ed.hojas[plan.h].tablas || []).filter((t) => !plan.objetos.tapadas.some((m) => m.ref.f0 === t.ref.f0 && m.ref.c0 === t.ref.c0));
    if (plan.objetos && plan.objetos.formato === 'ods' && ed.libro.extrasDeCelda) plan.objetos.enlacesPrevios = [...MoverOds._interno.extrasDe(await ed.paquete.leer('content.xml'), plan.h, plan.objetos.D)];
    const b = plan.tablas ? await MoverTablas.guardar(ed, plan.h, plan.destino, 'pegar', opciones && opciones.idioma) : await ed.guardar();
    return plan.objetos ? PegadoObjetos.aplicar(b, plan.objetos, opciones) : b;
  }

  globalThis.PegadoEspecial = { TIPOS, capturar, planear, ejecutar, _transponerFormula: transponerFormula, _operar: operar };
})();

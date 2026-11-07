/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Libro editable de Mini Cookies Hojas (fase 2 del PLAN): el libro que lee el motor de Excel del sitio (LibroXlsx) más lo que hace falta para
// editarlo en vivo: operaciones que se deshacen, recálculo de fórmulas y guardado SIN PÉRDIDA (solo cambian las celdas editadas; el resto del
// archivo sale byte a byte como vino). Sin DOM: se prueba en Node.
//
// Requiere zip-propio, xml-ligero, xml-editar, xlsx-formato, xlsx-estilos, xlsx-modelo, xlsx-paquete, xls-formulas (del sitio) y, de la suite,
// paquete-editable, motores/xlsx-escritor, hojas/calculo, hojas/estilos-libro y hojas/entrada.
//
//   const ed = await LibroEditable.abrir(bytes, { idioma: 'es' })
//   ed.libro, ed.hojas (los modelos de LibroXlsx, por índice del libro), ed.nombres() → [{ indice, nombre, estado }]
//   ed.celda(h, f, c) → { tipo: 'numero'|'texto'|'logico'|'error'|'vacio', valor, formula, estilo } | null
//   ed.opEscribir(h, f, c, texto) → op | { error }     (interpreta lo que se escribió, como Excel; no aplica)
//   ed.opEstado(h, celdas: [{ f, c, v }]) → op            (v = estado completo de la celda, null = vacía sin formato)
//   ed.opEstilo(h, rango, cambios) → op                   (deriva estilos nuevos para cada celda del rango)
//   ed.aplicar(op) → inversa                              (para HistorialDoc)
//   ed.terminar() → { calculadas, noSoportadas, circulares, tocadas }   (recalcula lo que cambió desde la última vez)
//   ed.protegida(h, f, c) → motivo | null                 (encabezado de tabla, fórmula matricial: se rechaza la edición)
//   await ed.guardar() → Uint8Array del .xlsx
//   ed.recalcularTodo() → { calculadas, noSoportadas, circulares }   (después de transformar la estructura con EstructuraLibro)
//   ed.pendientes() → cantidad de celdas cambiadas sin guardar en el paquete
(function () {
  const T = () => LibroXlsx.TIPO;
  const NOMBRE_TIPO = ['vacio', 'numero', 'texto', 'logico', 'error'];
  const MAX_FILAS = 1048576, MAX_COLS = 16384;
  const clave = (f, c) => f * MAX_COLS + c;

  // Los backends, del más específico al más general: el primero que reconoce los bytes abre el archivo (.ods si OdsBackend está cargado; si no, solo .xlsx/.xlsm)
  const backends = [globalThis.OdsBackend, XlsxBackend].filter(Boolean);

  async function abrir(bytes, opciones) {
    const o = opciones || {};
    let backend = null;
    for (const b of backends) if (await b.reconoce(bytes)) { backend = b; break; }
    const modelo = await backend.abrir(bytes, o);
    const { libro, paquete, hojas, estilos, externos } = modelo;
    const limites = hojas.map((h) => ({ filas: h.celdas.filaMax + 1, cols: h.celdas.colMax + 1 }));
    // celdas cambiadas desde el último guardado: hoja → Map(clave → estado original serializado)
    let cambios = new Map();
    let pendientesCalculo = [];

    // ---- lectura y escritura de una celda del modelo --------------------------------------------------------------------
    function celda(h, f, c) {
      const hoja = hojas[h];
      if (!hoja || hoja.sinCeldas) return null;
      const C = hoja.celdas, i = C.indiceDe(f, c);
      if (i < 0) return null;
      const fx = C.formulas.get(i);
      const x = { tipo: NOMBRE_TIPO[C.tipo[i]] || 'vacio', valor: C.valor[i] === undefined ? null : C.valor[i], formula: C.formulaDe(i), estilo: C.estilo[i], formulaTipo: fx ? fx.tipo : null };
      if (fx && fx.dinamica) x.dinamica = true;
      if (hoja.derramadas && hoja.derramadas.has(clave(f, c))) x.derrame = true;
      return x;
    }
    // el rango que desborda un ancla, como texto ("B2:D9"; solo el ancla si da #SPILL!)
    function rangoDerrame(h, f, c) {
      const d = construido ? motor.derrame(h, f, c) : null;
      if (!d) {   // un solo valor: las que Excel 365 guardó como matriz de una celda (cm, ref de sí misma) siguen así, y las que Excel guardaría así (SUM(LEN(A1:A3)))
        const C = hojas[h].celdas, i = C.indiceDe(f, c), fx = i >= 0 && C.formulas.get(i);
        return fx && (fx.tipo === 'array' || (construido && motor.necesitaMatriz(h, f, c))) ? LibroXlsx.numeroACol(c) + (f + 1) : null;
      }
      return d.ok ? LibroXlsx.numeroACol(d.c0) + (d.f0 + 1) + ':' + LibroXlsx.numeroACol(d.c1) + (d.f1 + 1) : LibroXlsx.numeroACol(c) + (f + 1);
    }
    const estadoDe = (h, f, c) => {
      const x = celda(h, f, c);
      if (!x || (x.tipo === 'vacio' && !x.formula && !x.estilo)) return null;   // una celda vacía sin formato es lo mismo que no tenerla
      const e = { tipo: x.tipo, valor: x.valor, formula: x.formula, estilo: x.estilo };
      if (x.derrame) e.derrame = true;
      if (x.dinamica) { const m = rangoDerrame(h, f, c); if (m) e.matriz = m; }
      else if (x.formula && x.formulaTipo !== 'array' && x.formulaTipo !== 'dataTable') e.legado = true;   // del archivo sin cm: deshacer, pegar u ordenar no la vuelven dinámica
      return e;
    };
    const igualEstado = (a, b) => JSON.stringify(a) === JSON.stringify(b);
    function registrarCambio(h, f, c, antes) {
      let m = cambios.get(h); if (!m) cambios.set(h, m = new Map());
      const k = clave(f, c);
      if (!m.has(k)) m.set(k, JSON.stringify(antes));
    }
    // deja la celda con ese estado (null = sin celda); no toca el historial. Con i = -1 (se sabe que la celda no existe) la agrega al final
    // del almacén sin reordenarlo: el almacén se ordena una sola vez en la próxima búsqueda (pegar miles de celdas nuevas no reordena por cada una)
    function ponerEstado(h, f, c, v, indice) {
      const hoja = hojas[h], C = hoja.celdas;
      let i = indice === undefined ? C.indiceDe(f, c) : indice;
      const tipo = v ? T()[v.tipo] : T().vacio;
      if (i < 0) {
        if (!v) return;
        i = C.agregar(f, c, tipo, v.valor, v.estilo || 0);   // las fórmulas y textos se guardan por este índice; el orden posterior los reubica
      } else {
        C.tipo[i] = tipo; C.valor[i] = v ? v.valor : null; C.estilo[i] = v ? v.estilo || 0 : 0;
      }
      C.ricos.delete(i); C.textoOriginal.delete(i);
      const tenia = C.formulas.get(i);
      // las fórmulas que se escriben en el editor son de Excel 365 (dinámicas: si dan una matriz, desbordan)
      if (v && v.formula) {
        if (!tenia || C.formulaDe(i) !== v.formula || (v.legado && tenia.dinamica)) C.formulas.set(i, { texto: v.formula, tipo: 'normal', ref: null, si: null, dinamica: !v.legado });
      } else if (tenia) C.formulas.delete(i);
      if (v && v.derrame) { if (!hoja.derramadas) hoja.derramadas = new Set(); hoja.derramadas.add(clave(f, c)); }
      else if (hoja.derramadas) hoja.derramadas.delete(clave(f, c));
      if (f + 1 > limites[h].filas) limites[h].filas = f + 1;
      if (c + 1 > limites[h].cols) limites[h].cols = c + 1;
    }

    // ---- motor de cálculo ----------------------------------------------------------------------------------------------
    const aValorCalculo = (x) => {
      if (!x || x.tipo === 'vacio') return null;
      if (x.tipo === 'error') return CalculoHojas.error(x.valor);
      return x.valor;
    };
    const motor = CalculoHojas.crear({
      hojas: () => libro.hojas.map((m) => m.nombre),
      valor: (h, f, c) => (hojas[h] ? aValorCalculo(celda(h, f, c)) : CalculoHojas.error('#REF!')),
      poner: (h, f, c, v, soloVista) => {
        const antes = estadoDe(h, f, c);
        if (!antes) return;
        let tipo, valor;
        if (CalculoHojas.esError(v)) { tipo = 'error'; valor = v.codigo; }
        else if (typeof v === 'number') { tipo = 'numero'; valor = v; }
        else if (typeof v === 'boolean') { tipo = 'logico'; valor = v; }
        else if (typeof v === 'string') { tipo = 'texto'; valor = v; }
        else { tipo = 'numero'; valor = 0; }
        if (antes.tipo === tipo && antes.valor === valor) return;
        if (!soloVista) registrarCambio(h, f, c, antes);   // lo volátil se ve al día; en el archivo lo recalcula Excel al abrir
        ponerEstado(h, f, c, { tipo, valor, formula: antes.formula, estilo: antes.estilo });
      },
      formulas: function* () {
        for (let h = 0; h < hojas.length; h++) {
          const hoja = hojas[h];
          if (!hoja || hoja.sinCeldas) continue;
          const C = hoja.celdas;
          for (const [i, fx] of C.formulas) {
            if ((fx.tipo === 'array' && !fx.dinamica) || fx.tipo === 'dataTable') continue;   // las matriciales de antes (Ctrl+Mayús+Intro) conservan sus valores
            const texto = C.formulaDe(i);
            if (texto === null) continue;
            const valor = C.tipo[i] === T().vacio ? null : C.tipo[i] === T().error ? CalculoHojas.error(C.valor[i]) : C.valor[i];
            yield { h, f: C.fila[i], c: C.col[i], texto, valor, dinamica: !!fx.dinamica, ref: fx.dinamica ? fx.ref : null };
          }
        }
      },
      limite: (h) => limites[h] || { filas: 0, cols: 0 },
      // matrices dinámicas: una celda con algo propio (no lo que desbordó otra fórmula) no se pisa
      ocupada: (h, f, c) => { const x = estadoDe(h, f, c); return !!x && !x.derrame && (x.tipo !== 'vacio' || !!x.formula); },
      derramar: (h, f, c, v, soloVista) => {
        const antes = estadoDe(h, f, c);
        let despues;
        if (v === undefined) despues = antes && antes.estilo ? { tipo: 'vacio', valor: null, formula: null, estilo: antes.estilo } : null;
        else {
          let tipo, valor;
          if (CalculoHojas.esError(v)) { tipo = 'error'; valor = v.codigo; }
          else if (typeof v === 'number') { tipo = 'numero'; valor = v; }
          else if (typeof v === 'boolean') { tipo = 'logico'; valor = v; }
          else { tipo = 'texto'; valor = String(v); }
          despues = { tipo, valor, formula: null, estilo: antes ? antes.estilo : (hojas[h].estiloDe ? hojas[h].estiloDe(f, c) : 0), derrame: true };
        }
        if (igualEstado(antes, despues)) return;
        if (!soloVista) registrarCambio(h, f, c, antes);
        ponerEstado(h, f, c, despues);
      },
      // el ancla cambia de rango: se anota antes, para que al guardar se reescriba con el rango nuevo
      antesDeDerrame: (h, f, c, soloVista) => { if (!soloVista) registrarCambio(h, f, c, estadoDe(h, f, c)); },
      nombre: (n, h) => {
        const u = String(n).toUpperCase();
        const local = libro.nombres.find((x) => x.nombre.toUpperCase() === u && x.hojaLocal === h);
        const global = libro.nombres.find((x) => x.nombre.toUpperCase() === u && x.hojaLocal === null);
        const x = local || global;
        return x ? x.ref : null;
      },
      esFormula: (h, f, c) => { const x = celda(h, f, c); return !!(x && x.formula); },
      formulaDe: (h, f, c) => { const x = hojas[h] && !hojas[h].sinCeldas ? celda(h, f, c) : null; return x ? x.formula : null; },
      filaOculta: (h, f) => { const x = hojas[h] && hojas[h].filas && hojas[h].filas.get(f); return !!(x && x.oculta); },
      // las tablas del libro (para Tabla[Columna]); los nombres de columna son los encabezados de la tabla
      tablas: () => {
        const lista = [];
        hojas.forEach((hoja, h) => {
          if (!hoja || hoja.sinCeldas) return;
          for (const t of hoja.tablas || []) {
            const columnas = [];
            for (let c = t.ref.c0; c <= t.ref.c1; c++) {
              const k = c - t.ref.c0;
              const x = t.encabezados ? celda(h, t.ref.f0, c) : null;
              columnas.push(t.nombresColumnas && t.nombresColumnas[k] ? t.nombresColumnas[k] : x && x.valor !== null ? String(x.valor) : 'Column' + (k + 1));
            }
            lista.push({ nombre: t.nombreVisible || t.nombre, h, ref: t.ref, encabezados: t.encabezados, totales: t.totales, columnas, nombreHoja: libro.hojas[h].nombre });
          }
        });
        return lista;
      },
      // [n]Hoja!A1: los valores que Excel guardó del otro libro (no se abre el otro archivo)
      externo: (n, nombreHoja) => {
        const x = externos[n - 1];
        if (!x) return null;
        const k = x.hojas.findIndex((nh) => nh.toLowerCase() === String(nombreHoja).toLowerCase());
        return k < 0 ? null : x.datos.get(k) || { celdas: new Map(), errorAlActualizar: false };
      },
      fecha1904: libro.fecha1904,
      idioma: o.idioma === 'en' ? 'en' : 'es',
      separadores: o.separadores || null,
    });
    let construido = false;
    function asegurarMotor() { if (!construido) { construido = true; motor.construir(); } }
    // un archivo con fórmulas sin valor guardado (lo escribió otro programa sin calcular): se calculan ya para verlas, sin tocar el archivo
    if (hojas.some((hoja) => hoja && !hoja.sinCeldas && [...hoja.celdas.formulas].some(([i, fx]) => fx.tipo !== 'array' && fx.tipo !== 'dataTable' && hoja.celdas.tipo[i] === T().vacio))) {
      construido = true; motor.construir({ soloVista: true });
    }

    // ---- operaciones -----------------------------------------------------------------------------------------------------
    function protegida(h, f, c) {
      const hoja = hojas[h];
      if (!hoja || hoja.sinCeldas) return 'hoja';
      if (modelo.protegida(h, f, c)) return 'protegida';
      return null;
    }
    function aplicar(op) {
      if (op.t !== 'celdas') throw new Error('operación desconocida ' + op.t);
      const inversa = { t: 'celdas', h: op.h, celdas: [] };
      const C = hojas[op.h].celdas;
      // 1) se lee todo antes de cambiar nada (si una celda se repite en la operación, cuenta la última)
      const ultima = new Map();
      op.celdas.forEach((x, k) => ultima.set(clave(x.f, x.c), k));
      const cambiosOp = [];
      op.celdas.forEach((x, k) => {
        if (ultima.get(clave(x.f, x.c)) !== k) return;
        const i = C.indiceDe(x.f, x.c);
        const antes = estadoDe(op.h, x.f, x.c);
        const despues = x.v ? Object.assign({ estilo: antes ? antes.estilo : 0 }, x.v) : null;
        if (igualEstado(antes, despues) || (!antes && despues && despues.tipo === 'vacio' && !despues.estilo && !despues.formula)) return;
        cambiosOp.push({ x, i, antes, despues });
      });
      // 2) primero las celdas que ya existen (sus índices siguen valiendo), después las nuevas, que se agregan sin reordenar
      cambiosOp.sort((a, b) => (a.i < 0) - (b.i < 0));
      for (const { x, i, antes, despues } of cambiosOp) {
        registrarCambio(op.h, x.f, x.c, antes);
        ponerEstado(op.h, x.f, x.c, despues, i);
      }
      // 3) el motor se entera al final, con el almacén ya ordenado
      for (const { x, antes, despues } of cambiosOp) {
        const f2 = despues && despues.formula ? despues.formula : null;
        const legado = !!(despues && despues.legado);
        if (((antes && antes.formula) !== f2 || (legado && !antes.legado)) && construido) motor.ponerFormula(op.h, x.f, x.c, f2, { dinamica: !legado });
        pendientesCalculo.push({ h: op.h, f: x.f, c: x.c });
        inversa.celdas.push({ f: x.f, c: x.c, v: antes });
      }
      return inversa.celdas.length ? inversa : null;
    }
    // después de un cambio de estructura (filas, columnas, hojas) los valores que trae el archivo son los de antes: se recalcula todo y
    // lo que cambia queda como editado (se guarda), como hace Excel al quitar filas
    function recalcularTodo() {
      const tenia = construido;
      asegurarMotor();
      const r = motor.recalcularTodo();
      pendientesCalculo = [];
      return Object.assign(r, { construido: !tenia });
    }
    function terminar() {
      const tocadas = pendientesCalculo;
      pendientesCalculo = [];
      if (!tocadas.length) return { calculadas: 0, noSoportadas: [], circulares: 0, tocadas };
      const tenia = construido;
      asegurarMotor();
      if (!tenia) { /* construir ya leyó las fórmulas nuevas del modelo */ }
      const r = motor.recalcular(tocadas);
      r.tocadas = tocadas;
      return r;
    }

    const opcionesEntrada = { idioma: o.idioma === "en" ? "en" : "es", fecha1904: libro.fecha1904 };
    function opEscribir(h, f, c, texto) {
      if (protegida(h, f, c)) return { error: 'protegida' };
      const x = EntradaCelda.interpretar(texto, opcionesEntrada);
      const antes = estadoDe(h, f, c);
      const estilo = antes ? antes.estilo : (hojas[h].estiloDe ? hojas[h].estiloDe(f, c) : 0);
      let v;
      switch (x.tipo) {
        case 'vacio': v = estilo ? { tipo: 'vacio', valor: null, formula: null, estilo } : null; break;
        case 'formula': {
          if (!XlsFormulasValida(x.formula)) return { error: 'formula' };
          v = { tipo: 'numero', valor: 0, formula: x.formula, estilo };   // el valor lo pone el recálculo
          break;
        }
        case 'error': v = { tipo: 'error', valor: x.valor, formula: null, estilo }; break;
        default: {
          let est = estilo;
          if (x.formato && x.tipo === 'numero') {
            const actual = libro.estilos.xf(estilo).formato || 'General';
            if (actual === 'General') est = estilos.derivar(estilo, { formato: x.formato });   // como Excel: "15%" en una celda General la deja en porcentaje
          }
          v = { tipo: x.tipo, valor: x.valor, formula: null, estilo: est };
        }
      }
      return { t: 'celdas', h, celdas: [{ f, c, v }] };
    }
    // ¿Excel aceptaría la fórmula? Que se pueda analizar y que cada función conocida tenga una cantidad de argumentos válida
    const aridad = new Map();
    for (const d of XlsFormulas.FUNCIONES || []) if (d && d.nombre) aridad.set(String(d.nombre).toUpperCase().replace(/^_XLFN\.(_XLWS\.)?/, ''), d);
    function XlsFormulasValida(texto) {
      let ast;
      // las referencias estructuradas (Tabla1[Precio], [@Precio]) las resuelve el motor: para el analizador valen como una celda
      const sinTablas = String(texto).replace(/"(?:[^"]|"")*"|((?:[A-Za-z_\\À-￿][A-Za-z0-9_.\\À-￿]*)?\[(?:[^\[\]']|'.|\[(?:[^\]']|'.)*\])*\])/g, (m, ref) => (ref ? 'A1' : m));
      try { ast = XlsFormulas._analizar(XlsFormulas._tokenizar(sinTablas)); } catch (e) { return false; }
      let ok = true;
      const ver = (n, prof) => {
        if (!n || !ok || prof > 300) return;
        if (n.t === 'fn') {
          const d = aridad.get(String(n.nombre).toUpperCase().replace(/^_XLFN\.(_XLWS\.)?/, ''));
          if (d && Number.isFinite(d.min) && Number.isFinite(d.max) && (n.args.length < d.min || n.args.length > d.max) && !(n.args.length === 1 && n.args[0] === null && d.min === 0)) ok = false;
          for (const a of n.args) ver(a, prof + 1);
        } else if (n.t === 'bin') { ver(n.a, prof + 1); ver(n.b, prof + 1); } else if (n.t === 'un' || n.t === 'paren') ver(n.a, prof + 1);
      };
      ver(ast, 0);
      return ok;
    }

    // ¿el rango toca un desborde? Excel no deja cambiar una parte de una matriz: ni ordenarla
    function hayDesborde(h, r) {
      asegurarMotor();
      const hoja = hojas[h];
      if (hoja.derramadas) for (const k of hoja.derramadas) { const f = Math.floor(k / MAX_COLS), c = k % MAX_COLS; if (f >= r.f0 && f <= r.f1 && c >= r.c0 && c <= r.c1) return true; }
      for (const d of motor.derrames.values()) if (d.ok && d.h === h && d.f0 >= r.f0 && d.f0 <= r.f1 && d.c0 >= r.c0 && d.c0 <= r.c1) return true;
      return false;
    }
    function opEstado(h, celdas) { return { t: 'celdas', h, celdas: celdas.filter((x) => !protegida(h, x.f, x.c)) }; }

    // para cada celda del rango (las que existen, o todas si el rango es chico) un estilo derivado del suyo
    function opEstilo(h, r, cambiosEstilo) {
      const hoja = hojas[h];
      const celdas = [];
      const area = (r.f1 - r.f0 + 1) * (r.c1 - r.c0 + 1);
      const porEstilo = new Map();
      const derivar = (base) => { if (!porEstilo.has(base)) porEstilo.set(base, estilos.derivar(base, typeof cambiosEstilo === 'function' ? cambiosEstilo(base) : cambiosEstilo)); return porEstilo.get(base); };
      const tratar = (f, c) => {
        const antes = estadoDe(h, f, c);
        const base = antes ? antes.estilo : hoja.estiloDe(f, c);
        const nuevo = derivar(base);
        if (nuevo === base && antes) return;
        celdas.push({ f, c, v: antes ? Object.assign({}, antes, { estilo: nuevo }) : { tipo: 'vacio', valor: null, formula: null, estilo: nuevo } });
      };
      if (area <= 20000) { for (let f = r.f0; f <= r.f1; f++) for (let c = r.c0; c <= r.c1; c++) tratar(f, c); }
      else { const C = hoja.celdas; C.paraCada((f, c) => { if (f >= r.f0 && f <= r.f1 && c >= r.c0 && c <= r.c1) tratar(f, c); }); }   // filas o columnas enteras: solo las celdas que existen
      return { t: 'celdas', h, celdas };
    }

    // ---- guardar --------------------------------------------------------------------------------------------------------
    function valorParaEscritor(x) {
      if (!x || x.tipo === 'vacio') return x && x.formula ? { formula: x.formula, valor: null } : null;
      let v = x.valor;
      if (x.tipo === 'error') v = { error: x.valor };
      if (x.formula && x.matriz) return { formula: x.formula, valor: v, matriz: x.matriz };   // el backend le pone el cm
      if (x.formula) return { formula: x.formula, valor: v };
      return v;
    }
    // Se arma la lista de cambios de una vez (sin esperas) y recién después se escribe: lo que se edite mientras tanto queda para el próximo guardado.
    let enCurso = Promise.resolve();
    function guardar() {
      const porHoja = [];
      for (const [h, m] of cambios) {
        const lista = [];
        for (const [k, original] of m) {
          const c = k % MAX_COLS, f = (k - c) / MAX_COLS;
          const ahora = estadoDe(h, f, c);
          if (JSON.stringify(ahora) === original) continue;   // volvió a como estaba
          lista.push({ f, c, valor: valorParaEscritor(ahora), estilo: ahora ? ahora.estilo : 0 });
        }
        if (lista.length) porHoja.push([h, lista]);
      }
      cambios = new Map();
      const tarea = enCurso.then(() => modelo.guardarCeldas(porHoja));
      // si el guardado falla, esas celdas vuelven a la lista (con un original imposible: se escriben sí o sí la próxima vez)
      enCurso = tarea.catch(() => { for (const [h, lista] of porHoja) for (const x of lista) { let m = cambios.get(h); if (!m) cambios.set(h, m = new Map()); if (!m.has(clave(x.f, x.c))) m.set(clave(x.f, x.c), '\u0000'); } });
      return tarea;
    }
    const pendientes = () => { let n = 0; for (const m of cambios.values()) n += m.size; return n; };

    return {
      libro, paquete, hojas, estilos, motor, backend, formato: backend.formato, destino: () => modelo.destino(),
      nombres: () => libro.hojas.map((m, i) => ({ indice: i, nombre: m.nombre, estado: m.estado, tipo: m.tipo })),
      celda, estadoDe, aplicar, terminar, recalcularTodo, protegida, opEscribir, opEstado, opEstilo, hayDesborde, guardar, pendientes,
      asegurarMotor, limite: (h) => limites[h], opciones: opcionesEntrada,
      // la celda ancla de la matriz dinámica que desbordó en (h, f, c), o null
      anclaDe: (h, f, c) => {
        const x = celda(h, f, c);
        if (!x || !x.derrame) return null;
        asegurarMotor();
        for (const d of motor.derrames.values()) if (d.ok && d.h === h && f >= d.f0 && f <= d.f1 && c >= d.c0 && c <= d.c1) return { f: d.f0, c: d.c0 };
        return null;
      },
    };
  }

  // Libro nuevo en blanco (una hoja), con la cara de un libro nuevo de Excel
  async function nuevo(opciones) {
    const nombre = (opciones && opciones.idioma) === 'en' ? 'Sheet1' : 'Hoja1';
    const bytes = await XlsxEscritor.crearLibro([{ nombre, filas: [] }]);
    return abrir(bytes, opciones);
  }

  globalThis.LibroEditable = { abrir, nuevo };
})();

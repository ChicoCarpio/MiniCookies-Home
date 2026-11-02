/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Ordenar y filtrar «solo para mirar» (fase A completa del plan 06): a partir de la hoja que se leyó arma una VISTA de esa hoja con las filas de un rango
// reordenadas y/o ocultas, sin tocar la hoja ni el archivo. La vista es un objeto que hereda de la hoja (Object.create) con otro almacén de celdas y otro
// mapa de filas, así que el motor de dibujo (XlsxDibujo.dibujarVista), la búsqueda y las estadísticas de la rejilla la usan como cualquier hoja.
// Sin DOM. Requiere xlsx-modelo.js (LibroXlsx).
//
// Uso: const rango = RejillaOrdenFiltro.regionActual(hoja, fila, columna);              // { f0, c0, f1, c1 }: la fila f0 es el encabezado
//      const valores = RejillaOrdenFiltro.valoresColumna(libro, hoja, rango, col, opc); // [{ clave, texto, cuenta }]
//      const v = RejillaOrdenFiltro.construirVista(libro, hoja, rango, { orden: { col, descendente }, filtros: Map(col → Set(clave excluida)), idioma }, opc);
//        (un Set con .solo = true lista las claves que QUEDAN: así se filtra cuando la lista de valores está cortada en MAX_VALORES)
//      v.hoja (la vista), v.visibles, v.total, v.orden ({ ok:false, motivo:'combinadas' } si no se pudo ordenar)

(function () {
  const MAX_ITERACIONES_REGION = 30000000;   // celdas que mira regionActual como mucho (una tabla de 1.000.000 × 25 entra)
  const MAX_FILA = 1048575, MAX_COL = 16383;
  const MAX_VALORES = 5000;          // valores distintos que ofrece la lista de un filtro
  const CLAVE_VACIAS = '\u0000vacias';

  // Región actual de una celda (como Ctrl+* de Excel): el rectángulo de celdas con contenido que la rodea, tocándose por los lados o las esquinas.
  function regionActual(hoja, f, c) {
    const C = hoja.celdas, T = LibroXlsx.TIPO;
    const lleno = (ff, cc) => { if (ff < 0 || cc < 0) return false; const i = C.indiceDe(ff, cc); return i >= 0 && C.tipo[i] !== T.vacio; };
    let r = { f0: f, c0: c, f1: f, c1: c };
    let gasto = 0;
    const hayEn = (f0, c0, f1, c1) => {   // alguna celda con contenido en el segmento (una fila o una columna de largo)
      for (let ff = f0; ff <= f1; ff++) for (let cc = c0; cc <= c1; cc++) { if (++gasto > MAX_ITERACIONES_REGION) return false; if (lleno(ff, cc)) return true; }
      return false;
    };
    // cada lado crece mientras pueda antes de mirar los otros: crecer una fila y volver a recorrer las columnas enteras en cada vuelta era cuadrático
    // (con 2.000 filas se gastaba el tope y la región quedaba cortada sin avisar)
    for (let cambio = true; cambio && gasto <= MAX_ITERACIONES_REGION;) {
      cambio = false;
      while (r.f0 > 0 && hayEn(r.f0 - 1, Math.max(0, r.c0 - 1), r.f0 - 1, r.c1 + 1)) { r.f0--; cambio = true; }
      while (r.f1 < MAX_FILA && hayEn(r.f1 + 1, Math.max(0, r.c0 - 1), r.f1 + 1, r.c1 + 1)) { r.f1++; cambio = true; }
      while (r.c0 > 0 && hayEn(Math.max(0, r.f0 - 1), r.c0 - 1, r.f1 + 1, r.c0 - 1)) { r.c0--; cambio = true; }
      while (r.c1 < MAX_COL && hayEn(Math.max(0, r.f0 - 1), r.c1 + 1, r.f1 + 1, r.c1 + 1)) { r.c1++; cambio = true; }
    }
    return r;
  }

  const dentro = (r, f, c) => f >= r.f0 && f <= r.f1 && c >= r.c0 && c <= r.c1;

  // El rango con el que se trabaja: el del autofiltro o de la tabla que contiene la celda; si no hay, su región actual
  function rangoDeTrabajo(hoja, f, c) {
    const candidatos = [];
    if (hoja.filtroAuto && hoja.filtroAuto.ref) candidatos.push(LibroXlsx.decodificarRango(hoja.filtroAuto.ref));
    for (const t of hoja.tablas || []) if (t.ref && t.encabezados !== 0 && t.encabezados !== false) candidatos.push(Object.assign({}, t.ref, { totales: t.totales ? 1 : 0 }));
    // la fila de totales de una tabla no se ordena ni se filtra (como en Excel)
    for (const r of candidatos) if (r && dentro(r, f, c)) return { f0: r.f0, c0: r.c0, f1: r.f1 - (r.totales || 0), c1: r.c1 };
    const g = regionActual(hoja, f, c);
    // una sola celda aislada no es una tabla
    return g;
  }

  function textoVisible(libro, hoja, i, opc) { return LibroXlsx.textoCelda(libro, hoja, i, opc).texto; }

  // Los valores distintos (por el texto que se ve, sin distinguir mayúsculas) de una columna del rango, sin el encabezado
  function valoresColumna(libro, hoja, rango, col, opc) {
    const C = hoja.celdas, mapa = new Map();
    let vacias = 0, truncado = false;
    for (let f = rango.f0 + 1; f <= rango.f1; f++) {
      const i = C.indiceDe(f, col);
      const texto = i < 0 || C.tipo[i] === LibroXlsx.TIPO.vacio ? '' : textoVisible(libro, hoja, i, opc);
      if (texto === '') { vacias++; continue; }
      const clave = texto.toLowerCase();
      const e = mapa.get(clave);
      if (e) e.cuenta++;
      else if (mapa.size < MAX_VALORES) mapa.set(clave, { clave, texto, cuenta: 1 });
      else truncado = true;
    }
    const lista = [...mapa.values()];
    if (vacias) lista.push({ clave: CLAVE_VACIAS, texto: '', cuenta: vacias, vacias: true });
    return { lista, truncado };
  }

  // Orden de Excel: números, textos, lógicos, errores; las vacías siempre al final
  function llaveOrden(C, i, texto) {
    const T = LibroXlsx.TIPO;
    if (i < 0 || C.tipo[i] === T.vacio) return { clase: 9 };
    switch (C.tipo[i]) {
      case T.numero: return { clase: 0, n: C.valor[i] };
      case T.texto: return { clase: 1, s: String(C.valor[i]) };
      case T.logico: return { clase: 2, n: C.valor[i] ? 1 : 0 };
      case T.error: return { clase: 3, s: String(C.valor[i]) };
      default: return { clase: 1, s: texto };
    }
  }

  function construirVista(libro, base, rango, estado, opc) {
    const idioma = estado && estado.idioma === 'en' ? 'en' : 'es';
    const C = base.celdas, T = LibroXlsx.TIPO;
    const fD0 = rango.f0 + 1, fD1 = rango.f1;
    const nDatos = Math.max(0, fD1 - fD0 + 1);
    const orden = estado && estado.orden ? estado.orden : null;
    const filtros = estado && estado.filtros ? estado.filtros : new Map();

    // ordenar con celdas combinadas dentro de las filas de datos y las columnas del rango: Excel tampoco lo permite
    let resultadoOrden = { ok: true };
    const intersecta = (m) => m.f1 >= fD0 && m.f0 <= fD1 && m.c1 >= rango.c0 && m.c0 <= rango.c1 && (m.f1 > m.f0 || m.c1 > m.c0);
    const hayCombinadas = (base.combinadas || []).some(intersecta);

    // perm[k] = fila de origen que queda en la posición k de las filas de datos
    const perm = new Int32Array(nDatos);
    for (let k = 0; k < nDatos; k++) perm[k] = fD0 + k;
    if (orden && nDatos > 1) {
      if (hayCombinadas) resultadoOrden = { ok: false, motivo: 'combinadas' };
      else {
        const col = orden.col, signo = orden.descendente ? -1 : 1;
        const comparador = new Intl.Collator(idioma === 'en' ? 'en-US' : 'es-CR', { sensitivity: 'accent' });
        const llaves = new Array(nDatos);
        for (let k = 0; k < nDatos; k++) {
          const i = C.indiceDe(fD0 + k, col);
          llaves[k] = llaveOrden(C, i, i >= 0 ? textoVisible(libro, base, i, opc) : '');
        }
        const idx = Array.from({ length: nDatos }, (_, k) => k);
        idx.sort((a, b) => {
          const A = llaves[a], B = llaves[b];
          if (A.clase === 9 || B.clase === 9) return A.clase === B.clase ? a - b : (A.clase === 9 ? 1 : -1);   // vacías al final, en los dos sentidos
          if (A.clase !== B.clase) return signo * (A.clase - B.clase);
          const d = A.s !== undefined ? comparador.compare(A.s, B.s) : A.n - B.n;
          return d !== 0 ? signo * d : a - b;
        });
        for (let k = 0; k < nDatos; k++) perm[k] = fD0 + idx[k];
      }
    }

    // la vista es una hoja con las filas de datos permutadas
    const cambia = resultadoOrden.ok && orden && nDatos > 1;
    const vista = Object.create(base);
    const nuevaFilaDe = new Map();   // fila de origen → fila destino (solo las de datos)
    for (let k = 0; k < nDatos; k++) nuevaFilaDe.set(perm[k], fD0 + k);
    const mueve = (f, c) => cambia && f >= fD0 && f <= fD1 && c >= rango.c0 && c <= rango.c1;
    if (cambia) {
      const almacen = new C.constructor();
      almacen.compartidas = C.compartidas;
      for (let i = 0; i < C.n; i++) {
        const f = C.fila[i], c = C.col[i], m = mueve(f, c);
        const j = almacen.agregar(m ? nuevaFilaDe.get(f) : f, c, C.tipo[i], C.valor[i], C.estilo[i]);
        if (C.ricos.has(i)) almacen.ricos.set(j, C.ricos.get(i));
        if (C.textoOriginal.has(i)) almacen.textoOriginal.set(j, C.textoOriginal.get(i));
        const fx = C.formulas.get(i);
        if (fx) almacen.formulas.set(j, m ? { texto: C.formulaDe(i), tipo: fx.tipo, ref: null, si: null, maestra: false } : fx);
      }
      vista.celdas = almacen;
      // lo que el archivo guarda por celda fuera del almacén va con su fila: las notas y las imágenes dentro de las celdas
      if (base.comentarios && base.comentarios.length) vista.comentarios = base.comentarios.map((n) => (mueve(n.f, n.c) ? Object.assign({}, n, { f: nuevaFilaDe.get(n.f) }) : n));
      if (base.imagenesCelda && base.imagenesCelda.size) {
        vista.imagenesCelda = new Map();
        for (const [clave, ruta] of base.imagenesCelda) { const f = Math.floor(clave / 16384), c = clave % 16384; vista.imagenesCelda.set(mueve(f, c) ? nuevaFilaDe.get(f) * 16384 + c : clave, ruta); }
      }
    }   // sin orden la vista usa el mismo almacén de celdas (solo cambian las filas ocultas)

    // filas: las propiedades (alto, oculta) siguen a los datos; y se ocultan las que no pasan los filtros
    const filas = new Map(base.filas);
    if (cambia) for (let k = 0; k < nDatos; k++) {
      const origen = perm[k], destino = fD0 + k;
      if (origen === destino) continue;
      if (base.filas.has(origen)) filas.set(destino, base.filas.get(origen)); else filas.delete(destino);
    }
    let visibles = 0;
    const activos = [...filtros].filter(([, ex]) => ex && ex.size);
    for (let k = 0; k < nDatos; k++) {
      const origen = perm[k], destino = fD0 + k;
      let pasa = true;
      for (const [col, excluidas] of activos) {
        const i = C.indiceDe(origen, col);
        const t = i < 0 || C.tipo[i] === T.vacio ? '' : textoVisible(libro, base, i, opc);
        const clave = t === '' ? CLAVE_VACIAS : t.toLowerCase();
        if (excluidas.solo ? !excluidas.has(clave) : excluidas.has(clave)) { pasa = false; break; }
      }
      const info = filas.get(destino);
      if (!pasa) filas.set(destino, Object.assign({}, info || {}, { oculta: true }));
      if (pasa && !(info && info.oculta)) visibles++;
    }
    vista.filas = filas;
    vista._botonesFiltro = undefined;
    return { hoja: vista, visibles, total: nDatos, orden: resultadoOrden, filtrada: activos.length > 0 };
  }

  globalThis.RejillaOrdenFiltro = { regionActual, rangoDeTrabajo, valoresColumna, construirVista, CLAVE_VACIAS };
})();

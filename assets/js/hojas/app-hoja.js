/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Mini Cookies Hojas, fase 3.3 del PLAN: la interfaz de filtrar, ordenar, los gráficos y la página. Autofiltro con el menú de cada columna
// (ordenar, valores con casillas, condiciones, superiores), ordenar por varios niveles, elegir un gráfico (o imagen) de la hoja para moverlo o
// cambiarle el tamaño, editar el rango de sus series, insertar un gráfico desde la selección, configurar la página e imprimir. Extensión de
// app.js (window.HojasExtensiones); los cambios del archivo los hacen FiltrosLibro, GraficosLibro y PaginaLibro y se deshacen como los de estructura.
(function () {
  'use strict';
  const ES = {
    'g-filtrar': 'Ordenar y filtrar', 'g-graficos': 'Gráficos', 'g-pagina': 'Página', 'cmd-filtro': 'Filtro', 'cmd-ordenar-niveles': 'Ordenar…', 'cmd-insertar-grafico': 'Gráfico', 'cmd-editar-grafico': 'Datos del gráfico',
    'cmd-pagina': 'Configurar página', 'cmd-imprimir': 'Imprimir', 'cmd-borrar-filtros': 'Borrar los filtros',
    'f-asc': 'Ordenar de A a Z', 'f-desc': 'Ordenar de Z a A', 'f-condicion': 'Filtrar por condición…', 'f-superiores': 'Los superiores…', 'f-buscar': 'Buscar', 'f-todo': '(Seleccionar todo)', 'f-vacias': '(Vacías)',
    'f-borrar': 'Borrar el filtro de la columna', 'f-aceptar': 'Aceptar', 'f-cancelar': 'Cancelar', 'f-muchos': 'Se muestran los primeros {0} valores', 'f-titulo-condicion': 'Filtrar por condición', 'f-y': 'Y', 'f-o': 'O',
    'f-mostrar': 'Mostrar las filas en las que', 'f-op1': 'Condición', 'f-val1': 'Valor', 'f-union': 'Y / O', 'f-op2': 'Segunda condición', 'f-val2': 'Valor', 'f-ninguna': '(ninguna)', 'f-tabla': 'Las tablas ya traen su filtro',
    'fop-equal': 'es igual a', 'fop-notEqual': 'no es igual a', 'fop-greaterThan': 'es mayor que', 'fop-greaterThanOrEqual': 'es mayor o igual que', 'fop-lessThan': 'es menor que', 'fop-lessThanOrEqual': 'es menor o igual que',
    'f-cuantos': 'Cuántos', 'f-porcentaje': 'En porcentaje', 'f-inferiores': 'Los inferiores', 'f-boton': 'Filtro de la columna {0}',
    'o-titulo': 'Ordenar', 'o-rango': 'Rango', 'o-col': 'Ordenar por', 'o-luego': 'Luego por', 'o-orden': 'Orden', 'o-asc': 'De A a Z (menor a mayor)', 'o-desc': 'De Z a A (mayor a menor)', 'o-encabezados': 'Los datos tienen encabezados',
    'o-mayus': 'Distinguir mayúsculas', 'o-nada': '(no)', 'o-columna': 'Columna {0}',
    'gr-titulo': 'Insertar gráfico', 'gr-tipo': 'Tipo', 'gr-col': 'Columnas', 'gr-bar': 'Barras', 'gr-line': 'Líneas', 'gr-pie': 'Circular', 'gr-rango': 'Datos', 'gr-nombre': 'Gráfico', 'gr-hecho': 'Gráfico insertado',
    'gr-datos': 'Datos del gráfico', 'gr-serie': 'Serie {0}', 'gr-s-nombre': 'Nombre', 'gr-s-cat': 'Categorías (eje X)', 'gr-s-val': 'Valores', 'gr-elige': 'Elige un gráfico (haz clic sobre él)', 'gr-objeto': 'Objeto de la hoja: arrástralo para moverlo; las esquinas cambian el tamaño',
    'p-titulo': 'Configurar página', 'p-orientacion': 'Orientación', 'p-vertical': 'Vertical', 'p-horizontal': 'Horizontal', 'p-papel': 'Tamaño del papel', 'p-ajuste': 'Ajuste', 'p-sin': 'Sin ajuste', 'p-ancho': 'Todas las columnas en una página',
    'p-una': 'Toda la hoja en una página', 'p-escala': 'Escala', 'p-escala-pct': 'Escala (%)', 'p-margenes': 'Márgenes', 'p-m-normal': 'Normales', 'p-m-estrecho': 'Estrechos', 'p-m-ancho': 'Anchos', 'p-m-propio': 'Los del archivo',
    'p-area': 'Área de impresión (vacía = todo)', 'p-titulos': 'Filas que se repiten arriba (por ejemplo 1:1)', 'p-encabezado': 'Encabezado (centro)', 'p-pie': 'Pie de página (centro; &P = página, &N = total)', 'p-cuadricula': 'Imprimir las líneas de la cuadrícula',
    'p-centrar': 'Centrar en la página', 'p-imprimiendo': 'Preparando la impresión…', 'p-bloqueada': 'El navegador bloqueó la ventana de impresión: se descarga el PDF',
  };
  const EN = {
    'g-filtrar': 'Sort & filter', 'g-graficos': 'Charts', 'g-pagina': 'Page', 'cmd-filtro': 'Filter', 'cmd-ordenar-niveles': 'Sort…', 'cmd-insertar-grafico': 'Chart', 'cmd-editar-grafico': 'Chart data',
    'cmd-pagina': 'Page setup', 'cmd-imprimir': 'Print', 'cmd-borrar-filtros': 'Clear filters',
    'f-asc': 'Sort A to Z', 'f-desc': 'Sort Z to A', 'f-condicion': 'Filter by condition…', 'f-superiores': 'Top items…', 'f-buscar': 'Search', 'f-todo': '(Select all)', 'f-vacias': '(Blanks)',
    'f-borrar': 'Clear filter from column', 'f-aceptar': 'OK', 'f-cancelar': 'Cancel', 'f-muchos': 'Showing the first {0} values', 'f-titulo-condicion': 'Filter by condition', 'f-y': 'And', 'f-o': 'Or',
    'f-mostrar': 'Show rows where', 'f-op1': 'Condition', 'f-val1': 'Value', 'f-union': 'And / Or', 'f-op2': 'Second condition', 'f-val2': 'Value', 'f-ninguna': '(none)', 'f-tabla': 'Tables already have their filter',
    'fop-equal': 'equals', 'fop-notEqual': 'does not equal', 'fop-greaterThan': 'is greater than', 'fop-greaterThanOrEqual': 'is greater than or equal to', 'fop-lessThan': 'is less than', 'fop-lessThanOrEqual': 'is less than or equal to',
    'f-cuantos': 'How many', 'f-porcentaje': 'As a percentage', 'f-inferiores': 'Bottom items', 'f-boton': 'Filter for column {0}',
    'o-titulo': 'Sort', 'o-rango': 'Range', 'o-col': 'Sort by', 'o-luego': 'Then by', 'o-orden': 'Order', 'o-asc': 'A to Z (smallest to largest)', 'o-desc': 'Z to A (largest to smallest)', 'o-encabezados': 'My data has headers',
    'o-mayus': 'Case sensitive', 'o-nada': '(none)', 'o-columna': 'Column {0}',
    'gr-titulo': 'Insert chart', 'gr-tipo': 'Type', 'gr-col': 'Column', 'gr-bar': 'Bar', 'gr-line': 'Line', 'gr-pie': 'Pie', 'gr-rango': 'Data', 'gr-nombre': 'Chart', 'gr-hecho': 'Chart inserted',
    'gr-datos': 'Chart data', 'gr-serie': 'Series {0}', 'gr-s-nombre': 'Name', 'gr-s-cat': 'Categories (X axis)', 'gr-s-val': 'Values', 'gr-elige': 'Pick a chart (click on it)', 'gr-objeto': 'Sheet object: drag it to move it; the corners resize it',
    'p-titulo': 'Page setup', 'p-orientacion': 'Orientation', 'p-vertical': 'Portrait', 'p-horizontal': 'Landscape', 'p-papel': 'Paper size', 'p-ajuste': 'Scaling', 'p-sin': 'No scaling', 'p-ancho': 'Fit all columns on one page',
    'p-una': 'Fit sheet on one page', 'p-escala': 'Custom scale', 'p-escala-pct': 'Scale (%)', 'p-margenes': 'Margins', 'p-m-normal': 'Normal', 'p-m-estrecho': 'Narrow', 'p-m-ancho': 'Wide', 'p-m-propio': 'The file’s',
    'p-area': 'Print area (empty = all)', 'p-titulos': 'Rows to repeat at top (for example 1:1)', 'p-encabezado': 'Header (center)', 'p-pie': 'Footer (center; &P = page, &N = total)', 'p-cuadricula': 'Print gridlines',
    'p-centrar': 'Center on page', 'p-imprimiendo': 'Preparing to print…', 'p-bloqueada': 'The browser blocked the print window: the PDF is downloaded instead',
  };
  const pagina = window.i18nPagina || (window.i18nPagina = {});
  pagina.es = Object.assign(pagina.es || {}, ES);
  pagina.en = Object.assign(pagina.en || {}, EN);

  (window.HojasExtensiones = window.HojasExtensiones || []).push((A) => {
    const tx = A.tx, ed = () => A.ed;
    const comun = { aceptar: tx('aceptar'), cancelar: tx('cancelar'), alCerrar: () => A.rejilla && A.rejilla.enfocar() };
    const colLetra = (c) => LibroXlsx.numeroACol(c);
    const relativa = (r) => colLetra(r.c0) + (r.f0 + 1) + ':' + colLetra(r.c1) + (r.f1 + 1);
    const crear = (tag, props) => Object.assign(document.createElement(tag), props || {});
    // lo que se ve en la celda (como en la pantalla: con su formato de número)
    function textoCelda(h, f, c) {
      const x = ed().celda(h, f, c);
      if (!x || x.tipo === 'vacio' || x.valor === null || x.valor === undefined) return { texto: '', valor: null };
      if (x.tipo === 'numero') {
        const fmt = (ed().libro.estilos.xf(x.estilo) || {}).formato || 'General';
        let t = String(x.valor);
        try { t = XlsxFormato.formatear(x.valor, fmt, { idioma: A.idioma(), fecha1904: ed().libro.fecha1904 }).texto; } catch (e) { /* el número tal cual */ }
        return { texto: t.trim(), valor: x.valor };
      }
      if (x.tipo === 'logico') return { texto: A.idioma() === 'en' ? (x.valor ? 'TRUE' : 'FALSE') : (x.valor ? 'VERDADERO' : 'FALSO'), valor: x.valor };
      return { texto: String(x.valor), valor: x.valor };
    }
    function regionActual(h, f, c) {
      const vacia = (ff, cc) => { const x = ed().celda(h, ff, cc); return !x || x.tipo === 'vacio' || x.valor === null || x.valor === ''; };
      const r = { f0: f, c0: c, f1: f, c1: c };
      for (let vuelta = 0; vuelta < 2000; vuelta++) {
        let creció = false;
        const borde = (ff, cc) => ff >= 0 && cc >= 0 && ff <= 1048575 && cc <= 16383 && !vacia(ff, cc);
        for (let cc = r.c0 - 1; cc <= r.c1 + 1 && !creció; cc++) { if (borde(r.f0 - 1, cc)) { r.f0--; creció = true; } else if (borde(r.f1 + 1, cc)) { r.f1++; creció = true; } }
        for (let ff = r.f0 - 1; ff <= r.f1 + 1 && !creció; ff++) { if (borde(ff, r.c0 - 1)) { r.c0--; creció = true; } else if (borde(ff, r.c1 + 1)) { r.c1++; creció = true; } }
        if (!creció) break;
      }
      return r;
    }
    const rangoDeTrabajo = () => { const s = A.filasSel(), a = A.rejilla.activa(); return s.f0 === s.f1 && s.c0 === s.c1 ? regionActual(A.hojaActiva(), a.f, a.c) : s; };

    // ---- ordenar ----------------------------------------------------------------------------------------------------------
    // Como Excel: números < textos < lógicos < errores; las vacías siempre al final (también de Z a A). Las fórmulas se corren con su fila.
    function ordenarRango(h, r, claves, encabezado, mayus, nombre) {
      const f0 = encabezado ? r.f0 + 1 : r.f0;
      if (r.f1 <= f0 - 1) return;
      if ((ed().hojas[h].combinadas || []).some((m) => !(m.f1 < r.f0 || m.f0 > r.f1 || m.c1 < r.c0 || m.c0 > r.c1))) { A.avisar(tx('aviso-ordenar-combinadas')); return; }
      if (ed().hayDesborde(h, r)) { A.avisar(tx('aviso-ordenar-matriz')); return; }
      const cmp = new Intl.Collator(A.idioma() === 'en' ? 'en-US' : 'es-CR', { numeric: true, sensitivity: mayus ? 'variant' : 'base', caseFirst: mayus ? 'lower' : 'false' }).compare;
      const clase = (x) => (!x || x.tipo === 'vacio' || x.valor === null || x.valor === '' ? 4 : x.tipo === 'numero' ? 0 : x.tipo === 'texto' ? 1 : x.tipo === 'logico' ? 2 : 3);
      const filas = [];
      for (let f = f0; f <= r.f1; f++) { const fila = []; for (let c = r.c0; c <= r.c1; c++) fila.push(ed().estadoDe(h, f, c)); filas.push({ f, fila, claves: claves.map((k) => ed().celda(h, f, k.c)) }); }
      filas.sort((p, q) => {
        for (let i = 0; i < claves.length; i++) {
          const a = p.claves[i], b = q.claves[i], ra = clase(a), rb = clase(b);
          if (ra === 4 || rb === 4) { if (ra !== rb) return ra - rb; continue; }
          let d = ra - rb;
          if (!d) d = ra === 0 ? a.valor - b.valor : ra === 1 ? cmp(String(a.valor), String(b.valor)) : ra === 2 ? (a.valor ? 1 : 0) - (b.valor ? 1 : 0) : 0;
          if (d) return claves[i].desc ? -d : d;
        }
        return p.f - q.f;
      });
      const celdas = [];
      filas.forEach((x, k) => {
        const destino = f0 + k;
        x.fila.forEach((v, j) => {
          const nv = v ? Object.assign({}, v) : null;
          if (nv) { delete nv.derrame; delete nv.matriz; delete nv.fija; }
          if (nv && nv.formula && destino !== x.f) nv.formula = LibroXlsx.desplazarFormula(nv.formula, destino - x.f, 0);
          celdas.push({ f: destino, c: r.c0 + j, v: nv });
        });
      });
      A.hacer(nombre || tx('o-titulo'), ed().opEstado(h, celdas));
    }
    async function ordenarNiveles() {
      const h = A.hojaActiva(), r = rangoDeTrabajo(), a = A.rejilla.activa();
      const tipo = (f, c) => { const x = ed().celda(h, f, c); return x ? x.tipo : 'vacio'; };
      let encabezado = true;
      for (let c = r.c0; c <= r.c1; c++) if (tipo(r.f0, c) !== 'texto') encabezado = false;
      const opciones = (conEnc) => { const l = []; for (let c = r.c0; c <= Math.min(r.c1, r.c0 + 200); c++) { const t = conEnc ? textoCelda(h, r.f0, c).texto : ''; l.push([String(c), t ? t + ' (' + colLetra(c) + ')' : tx('o-columna', colLetra(c))]); } return l; };
      const ordenes = [['asc', tx('o-asc')], ['desc', tx('o-desc')]];
      const v = await Dialogo.formulario(Object.assign({}, comun, {
        titulo: tx('o-titulo') + ' · ' + relativa(r),
        campos: [
          { nombre: 'encabezados', etiqueta: tx('o-encabezados'), tipo: 'check', valor: encabezado },
          { nombre: 'c1', etiqueta: tx('o-col'), tipo: 'select', valor: String(Math.max(r.c0, Math.min(r.c1, a.c))), opciones: opciones(encabezado) },
          { nombre: 'o1', etiqueta: tx('o-orden'), tipo: 'select', valor: 'asc', opciones: ordenes },
          { nombre: 'c2', etiqueta: tx('o-luego'), tipo: 'select', valor: '', opciones: [['', tx('o-nada')], ...opciones(encabezado)] },
          { nombre: 'o2', etiqueta: tx('o-orden'), tipo: 'select', valor: 'asc', opciones: ordenes, visible: (x) => x.c2 !== '' },
          { nombre: 'c3', etiqueta: tx('o-luego'), tipo: 'select', valor: '', opciones: [['', tx('o-nada')], ...opciones(encabezado)], visible: (x) => x.c2 !== '' },
          { nombre: 'o3', etiqueta: tx('o-orden'), tipo: 'select', valor: 'asc', opciones: ordenes, visible: (x) => x.c2 !== '' && x.c3 !== '' },
          { nombre: 'mayus', etiqueta: tx('o-mayus'), tipo: 'check', valor: false },
        ],
      }));
      if (!v) return;
      const claves = [[v.c1, v.o1], [v.c2, v.o2], [v.c3, v.o3]].filter(([c]) => c !== '').map(([c, o]) => ({ c: +c, desc: o === 'desc' }));
      ordenarRango(h, r, claves, v.encabezados, v.mayus);
    }

    // ---- autofiltro -------------------------------------------------------------------------------------------------------
    const filtrosHoja = (h) => {
      const hoja = ed().hojas[h], l = [];
      const fa = hoja && hoja.filtroAuto && LibroXlsx.decodificarRango(hoja.filtroAuto.ref || '');
      if (fa) l.push({ tipo: 'hoja', ref: fa, columnas: hoja.filtroAuto.columnas });
      for (const t of (hoja && hoja.tablas) || []) if (t.filtro && t.encabezados) l.push({ tipo: 'tabla', ref: t.ref, tabla: t, columnas: t.filtro.columnas });
      return l;
    };
    async function alternarFiltro() {
      const h = A.hojaActiva(), a = A.rejilla.activa();
      const en = A.FIL().filtroEn(ed().hojas[h], a.f, a.c);
      if (en && en.tipo === 'tabla') { A.avisar(tx('f-tabla')); return; }
      const hoja = ed().hojas[h];
      if (hoja.filtroAuto) { await A.estructural(tx('cmd-filtro'), (b) => A.FIL().ponerAutofiltro(b, h, null), { sel: A.seleccion() }); return; }
      const r = rangoDeTrabajo();
      await A.estructural(tx('cmd-filtro'), (b) => A.FIL().ponerAutofiltro(b, h, r), { sel: A.seleccion() });
    }
    async function aplicarFiltro(h, objetivo, colId, criterio) {
      const textoDe = (f, c) => textoCelda(h, f, c);
      await A.estructural(tx('cmd-filtro'), (b) => A.FIL().filtrar(b, h, objetivo, colId, criterio, textoDe), { sel: A.seleccion() });
    }
    // los botones del filtro: transparentes sobre los que dibuja la hoja (para el clic y el teclado)
    const capaBotones = crear('div', { className: 'hc-capa-filtros' });
    document.body.append(capaBotones);
    const menu = crear('div', { className: 'hc-menu-filtro', hidden: true });
    menu.setAttribute('role', 'dialog');
    document.body.append(menu);
    function ponerBotones() {
      capaBotones.textContent = '';
      if (!ed() || !A.rejilla || A.lectura) return;
      const h = A.hojaActiva();
      const z = A.rejilla.zoom() / 100;
      for (const o of filtrosHoja(h)) {
        for (let c = o.ref.c0; c <= o.ref.c1 && c - o.ref.c0 < 300; c++) {
          const col = o.columnas ? o.columnas.get(c - o.ref.c0) : null;
          if (col && col.oculto) continue;
          const r = A.rectEnPantalla(o.ref.f0, c);
          if (!r || r.w < 4 || r.x + r.w < r.vista.left || r.x > r.vista.right || r.y + r.h < r.vista.top || r.y > r.vista.bottom) continue;
          const lado = Math.min(17 * z, r.h - 2, r.w - 2);
          const b = crear('button', { type: 'button', className: 'hc-boton-filtro' });
          b.setAttribute('aria-label', tx('f-boton', colLetra(c)));
          b.style.left = (r.x + r.w - lado - 1) + 'px'; b.style.top = (r.y + r.h - lado - 1) + 'px'; b.style.width = lado + 'px'; b.style.height = lado + 'px';
          b.addEventListener('mousedown', (e) => e.preventDefault());
          b.addEventListener('click', () => abrirMenu(h, o, c - o.ref.c0, r));
          capaBotones.append(b);
        }
      }
    }
    let pendiente = 0;
    const reponer = () => { if (pendiente) return; pendiente = requestAnimationFrame(() => { pendiente = 0; ponerBotones(); colocarMarco(); }); };
    document.addEventListener('scroll', reponer, true); window.addEventListener('resize', reponer); document.addEventListener('wheel', reponer, { passive: true, capture: true });
    A.ganchos.alSeleccionar.push(reponer);
    A.ganchos.alCambiar.push(reponer);
    function cerrarMenu() { menu.hidden = true; menu.textContent = ''; }
    document.addEventListener('mousedown', (e) => { if (!menu.hidden && !menu.contains(e.target) && !e.target.classList.contains('hc-boton-filtro')) cerrarMenu(); });
    async function abrirMenu(h, objetivo, colId, rect) {
      cerrarMenu();
      const r = objetivo.ref, c = r.c0 + colId;
      const criterios = await A.FIL().criterios(await ed().guardar(), h, objetivo);
      const actual = criterios.get(colId) || null;
      // los valores distintos de la columna (como se ven), en orden
      const vistos = new Map();
      let vacias = false;
      const ultima = objetivo.tipo === 'tabla' && objetivo.tabla.totales ? r.f1 - 1 : r.f1;
      for (let f = r.f0 + 1; f <= ultima && vistos.size < 10000; f++) { const x = textoCelda(h, f, c); if (x.texto === '') vacias = true; else if (!vistos.has(x.texto.toLowerCase())) vistos.set(x.texto.toLowerCase(), x); }
      const cmp = new Intl.Collator(A.idioma() === 'en' ? 'en-US' : 'es-CR', { numeric: true, sensitivity: 'base' }).compare;
      const valores = [...vistos.values()].sort((a, b) => (typeof a.valor === 'number' && typeof b.valor === 'number' ? a.valor - b.valor : typeof a.valor === 'number' ? -1 : typeof b.valor === 'number' ? 1 : cmp(a.texto, b.texto)));
      const marcado = (t) => !actual || actual.tipo !== 'valores' || actual.valores.some((v) => v.toLowerCase() === t.toLowerCase());
      const boton = (texto, fn, clase) => { const b = crear('button', { type: 'button', textContent: texto, className: clase || 'hc-mf-item' }); b.addEventListener('click', fn); return b; };
      menu.append(
        boton(tx('f-asc'), () => { cerrarMenu(); ordenarRango(h, objetivo.tipo === 'tabla' && objetivo.tabla.totales ? { ...r, f1: r.f1 - 1 } : r, [{ c, desc: false }], true, false, tx('f-asc')); }),
        boton(tx('f-desc'), () => { cerrarMenu(); ordenarRango(h, objetivo.tipo === 'tabla' && objetivo.tabla.totales ? { ...r, f1: r.f1 - 1 } : r, [{ c, desc: true }], true, false, tx('f-desc')); }),
        crear('hr'),
        boton(tx('f-condicion'), () => { cerrarMenu(); dialogoCondicion(h, objetivo, colId, actual); }),
        boton(tx('f-superiores'), () => { cerrarMenu(); dialogoSuperiores(h, objetivo, colId, actual); }),
        boton(tx('f-borrar'), () => { cerrarMenu(); aplicarFiltro(h, objetivo, colId, null); }),
        crear('hr'),
      );
      const busca = crear('input', { type: 'search', className: 'hc-mf-buscar', placeholder: tx('f-buscar') });
      busca.setAttribute('aria-label', tx('f-buscar'));
      const lista = crear('div', { className: 'hc-mf-lista' });
      const casilla = (texto, valor, si) => { const l = crear('label'); const i = crear('input', { type: 'checkbox', checked: si }); i.dataset.valor = valor; l.append(i, ' ', crear('span', { textContent: texto })); return l; };
      const todo = casilla(tx('f-todo'), '\u0000todo', true);
      lista.append(todo);
      for (const x of valores) lista.append(casilla(x.texto, x.texto, marcado(x.texto)));
      if (vacias) lista.append(casilla(tx('f-vacias'), '\u0000vacias', !actual || actual.tipo !== 'valores' || actual.vacias));
      const items = () => [...lista.querySelectorAll('input')].filter((i) => i.dataset.valor !== '\u0000todo');
      const sincronizarTodo = () => { const l = items().filter((i) => !i.parentNode.hidden); todo.querySelector('input').checked = l.every((i) => i.checked); };
      sincronizarTodo();
      todo.querySelector('input').addEventListener('change', (e) => { for (const i of items()) if (!i.parentNode.hidden) i.checked = e.target.checked; });
      lista.addEventListener('change', (e) => { if (e.target !== todo.querySelector('input')) sincronizarTodo(); });
      busca.addEventListener('input', () => { const q = busca.value.trim().toLowerCase(); for (const i of items()) i.parentNode.hidden = !!q && !i.parentNode.textContent.toLowerCase().includes(q); sincronizarTodo(); });
      menu.append(busca, lista);
      if (vistos.size >= 10000) menu.append(crear('small', { textContent: tx('f-muchos', '10000') }));
      const pie = crear('div', { className: 'hc-mf-pie' });
      pie.append(boton(tx('f-aceptar'), () => {
        const sel = items().filter((i) => i.checked).map((i) => i.dataset.valor);
        const conVacias = sel.includes('\u0000vacias'), textos = sel.filter((x) => x !== '\u0000vacias');
        cerrarMenu();
        const todos = textos.length === valores.length && (conVacias || !vacias);
        aplicarFiltro(h, objetivo, colId, todos ? null : { tipo: 'valores', valores: textos, vacias: conVacias });
      }, 'boton'), boton(tx('f-cancelar'), () => { cerrarMenu(); A.rejilla.enfocar(); }, 'boton boton-secundario'));
      menu.append(pie);
      menu.hidden = false;
      const ancho = Math.min(300, window.innerWidth - 16);
      menu.style.width = ancho + 'px';
      menu.style.left = Math.max(8, Math.min(rect.x + rect.w - ancho, window.innerWidth - ancho - 8)) + 'px';
      menu.style.top = Math.min(rect.y + rect.h + 2, Math.max(8, window.innerHeight - menu.offsetHeight - 8)) + 'px';
      busca.focus();
    }
    menu.addEventListener('keydown', (e) => { if (e.key === 'Escape') { e.preventDefault(); cerrarMenu(); A.rejilla.enfocar(); } });
    const OPS_F = ['equal', 'notEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual'];
    async function dialogoCondicion(h, objetivo, colId, actual) {
      const c0 = actual && actual.tipo === 'condiciones' ? actual.condiciones : [];
      const v = await Dialogo.formulario(Object.assign({}, comun, {
        titulo: tx('f-titulo-condicion') + ' · ' + colLetra(objetivo.ref.c0 + colId),
        campos: [
          { nombre: 'op1', etiqueta: tx('f-mostrar'), tipo: 'select', valor: c0[0] ? c0[0].operador : 'equal', opciones: OPS_F.map((k) => [k, tx('fop-' + k)]) },
          { nombre: 'val1', etiqueta: tx('f-val1'), tipo: 'texto', valor: c0[0] ? c0[0].valor : '', requerido: true },
          { nombre: 'union', etiqueta: tx('f-union'), tipo: 'select', valor: actual && actual.y ? 'y' : 'o', opciones: [['y', tx('f-y')], ['o', tx('f-o')]] },
          { nombre: 'op2', etiqueta: tx('f-op2'), tipo: 'select', valor: c0[1] ? c0[1].operador : '', opciones: [['', tx('f-ninguna')], ...OPS_F.map((k) => [k, tx('fop-' + k)])] },
          { nombre: 'val2', etiqueta: tx('f-val2'), tipo: 'texto', valor: c0[1] ? c0[1].valor : '', visible: (x) => x.op2 !== '' },
        ],
      }));
      if (!v) return;
      // los números se escriben como en la página (coma decimal en español) y se guardan como Excel (punto)
      const valor = (t) => { const x = EntradaCelda.interpretar(String(t), ed().opciones); return x.tipo === 'numero' ? String(x.valor) : String(t); };
      const condiciones = [{ operador: v.op1, valor: valor(v.val1) }];
      if (v.op2 && String(v.val2).trim()) condiciones.push({ operador: v.op2, valor: valor(v.val2) });
      aplicarFiltro(h, objetivo, colId, { tipo: 'condiciones', y: v.union === 'y', condiciones });
    }
    async function dialogoSuperiores(h, objetivo, colId, actual) {
      const a0 = actual && actual.tipo === 'top10' ? actual : { n: 10, porcentaje: false, abajo: false };
      const v = await Dialogo.formulario(Object.assign({}, comun, {
        titulo: tx('f-superiores').replace('…', '') + ' · ' + colLetra(objetivo.ref.c0 + colId),
        campos: [{ nombre: 'n', etiqueta: tx('f-cuantos'), tipo: 'numero', valor: a0.n, requerido: true }, { nombre: 'porcentaje', etiqueta: tx('f-porcentaje'), tipo: 'check', valor: a0.porcentaje }, { nombre: 'abajo', etiqueta: tx('f-inferiores'), tipo: 'check', valor: a0.abajo }],
      }));
      if (!v) return;
      aplicarFiltro(h, objetivo, colId, { tipo: 'top10', n: v.n, porcentaje: v.porcentaje, abajo: v.abajo });
    }
    async function borrarFiltros() {
      const h = A.hojaActiva();
      for (const o of filtrosHoja(h)) {
        const crit = await A.FIL().criterios(await ed().guardar(), h, o);
        if (!crit.size) continue;
        const textoDe = (f, c) => textoCelda(h, f, c);
        await A.estructural(tx('cmd-borrar-filtros'), async (b) => { let x = b; for (const id of crit.keys()) x = await A.FIL().filtrar(x, h, o, id, null, textoDe); return x; }, { sel: A.seleccion() });
      }
    }

    // ---- gráficos (y demás objetos de la hoja) ---------------------------------------------------------------------------
    // Un clic sobre un objeto lo elige: un marco con asas encima; arrastrarlo lo mueve y las esquinas le cambian el tamaño. Doble clic en un
    // gráfico: los datos de sus series.
    const marco = crear('div', { className: 'hc-marco-objeto', hidden: true, tabIndex: 0 });
    marco.setAttribute('role', 'img');
    for (const q of ['nw', 'ne', 'sw', 'se']) marco.append(crear('div', { className: 'hc-asa-objeto hc-asa-' + q }));
    document.body.append(marco);
    let elegido = null;   // { indice, esGrafico }
    function colocarMarco() {
      if (!elegido || !A.rejilla) { marco.hidden = true; return; }
      const o = A.rejilla.objetos().find((x) => x.indice === elegido.indice);
      if (!o) { elegido = null; marco.hidden = true; return; }
      const v = A.rejilla.vista.getBoundingClientRect();
      marco.style.left = (v.left + o.x) + 'px'; marco.style.top = (v.top + o.y) + 'px'; marco.style.width = o.ancho + 'px'; marco.style.height = o.alto + 'px';
      marco.hidden = false;
    }
    function soltarObjeto() { elegido = null; marco.hidden = true; }
    async function graficosHoja() { return A.GRA().listar(await ed().guardar(), A.hojaActiva()); }
    document.addEventListener('pointerdown', async (e) => {
      if (!ed() || !A.rejilla || A.lectura || e.button !== 0) return;
      const v = A.rejilla.vista, rv = v.getBoundingClientRect();
      if (marco.contains(e.target)) return;
      const dentro = e.clientX >= rv.left && e.clientX <= rv.right && e.clientY >= rv.top && e.clientY <= rv.bottom && A.rejilla.cuerpo.contains(e.target);
      if (!dentro) { if (!menu.contains(e.target) && !e.target.closest('dialog')) soltarObjeto(); return; }
      const x = e.clientX - rv.left, y = e.clientY - rv.top;
      const objs = A.rejilla.objetos();
      const o = [...objs].reverse().find((q) => x >= q.x && x <= q.x + q.ancho && y >= q.y && y <= q.y + q.alto);
      if (!o) { soltarObjeto(); return; }
      e.preventDefault(); e.stopPropagation();
      elegido = { indice: o.indice };
      colocarMarco();
      marco.setAttribute('aria-label', tx('gr-objeto'));
      marco.focus({ preventScroll: true });
      arrastre(e, 'mover');
    }, true);
    marco.addEventListener('pointerdown', (e) => { if (e.button !== 0) return; e.preventDefault(); e.stopPropagation(); const asa = e.target.classList.contains('hc-asa-objeto') ? e.target.className.replace(/.*hc-asa-(\w+).*/, '$1') : 'mover'; arrastre(e, asa); });
    marco.addEventListener('dblclick', () => editarGrafico());
    marco.addEventListener('keydown', (e) => { if (e.key === 'Escape') { soltarObjeto(); A.rejilla.enfocar(); } else if (e.key === 'Enter') editarGrafico(); });
    function arrastre(e0, modo) {
      const ini = { x: e0.clientX, y: e0.clientY, l: parseFloat(marco.style.left), t: parseFloat(marco.style.top), w: parseFloat(marco.style.width), h: parseFloat(marco.style.height) };
      let movido = false;
      const mover = (e) => {
        const dx = e.clientX - ini.x, dy = e.clientY - ini.y;
        if (Math.abs(dx) + Math.abs(dy) > 3) movido = true;
        let { l, t, w, h } = ini;
        if (modo === 'mover') { l += dx; t += dy; }
        else { if (modo.includes('w')) { l += dx; w -= dx; } if (modo.includes('e')) w += dx; if (modo.includes('n')) { t += dy; h -= dy; } if (modo.includes('s')) h += dy; }
        if (w < 20 || h < 20) return;
        marco.style.left = l + 'px'; marco.style.top = t + 'px'; marco.style.width = w + 'px'; marco.style.height = h + 'px';
      };
      const fin = async () => {
        document.removeEventListener('pointermove', mover); document.removeEventListener('pointerup', fin);
        if (!movido || !elegido) return;
        const rv = A.rejilla.vista.getBoundingClientRect();
        const l = parseFloat(marco.style.left) - rv.left, t = parseFloat(marco.style.top) - rv.top, w = parseFloat(marco.style.width), h = parseFloat(marco.style.height);
        const de = A.rejilla.anclaEn(l, t), a = A.rejilla.anclaEn(l + w, t + h);
        const z = A.rejilla.zoom() / 100, k = elegido.indice, hh = A.hojaActiva();
        await A.estructural(tx('cmd-editar-grafico'), (b) => A.GRA().moverAncla(b, hh, k, { de, a, ext: { w: w / z * 9525, h: h / z * 9525 } }), { sel: A.seleccion() });
        elegido = { indice: k };
        setTimeout(colocarMarco, 50);
      };
      document.addEventListener('pointermove', mover); document.addEventListener('pointerup', fin);
    }
    const visible = (f) => (f ? '=' + EntradaCelda.formulaVisible(f, A.idioma()) : '');
    const guardada = (t) => { const s = String(t || '').trim(); return s ? EntradaCelda.formulaAlmacenada(s.replace(/^=/, ''), A.idioma()) : null; };
    async function editarGrafico() {
      if (!elegido) { A.avisar(tx('gr-elige')); return; }
      const h = A.hojaActiva(), k = elegido.indice;
      const g = (await graficosHoja()).find((x) => x.indice === k);
      if (!g) return;
      const campos = [];
      g.series.forEach((s, i) => {
        campos.push({ nombre: 'n' + i, etiqueta: tx('gr-serie', i + 1) + ' · ' + tx('gr-s-nombre'), tipo: 'texto', valor: visible(s.nombre) });
        campos.push({ nombre: 'c' + i, etiqueta: tx('gr-serie', i + 1) + ' · ' + tx('gr-s-cat'), tipo: 'texto', valor: visible(s.categorias) });
        campos.push({ nombre: 'v' + i, etiqueta: tx('gr-serie', i + 1) + ' · ' + tx('gr-s-val'), tipo: 'texto', valor: visible(s.valores), requerido: true });
      });
      const v = await Dialogo.formulario(Object.assign({}, comun, { titulo: tx('gr-datos'), campos }));
      if (!v) return;
      const series = g.series.map((s, i) => ({ nombre: guardada(v['n' + i]), categorias: guardada(v['c' + i]), valores: guardada(v['v' + i]) }));
      await A.estructural(tx('cmd-editar-grafico'), (b) => A.GRA().cambiarSeries(b, h, k, series, A.GRA().lectorDe(ed())), { sel: A.seleccion() });
    }
    async function insertarGrafico() {
      const h = A.hojaActiva(), r = rangoDeTrabajo();
      const v = await Dialogo.formulario(Object.assign({}, comun, {
        titulo: tx('gr-titulo'),
        campos: [{ nombre: 'tipo', etiqueta: tx('gr-tipo'), tipo: 'select', valor: 'col', opciones: ['col', 'bar', 'line', 'pie'].map((k) => [k, tx('gr-' + k)]) }, { nombre: 'rango', etiqueta: tx('gr-rango'), tipo: 'texto', valor: relativa(r), requerido: true }],
        validar: (x) => (LibroXlsx.decodificarRango(String(x.rango).replace(/\$/g, '').toUpperCase()) ? null : tx('t-rango-malo')),
      }));
      if (!v) return;
      const rr = LibroXlsx.decodificarRango(String(v.rango).replace(/\$/g, '').toUpperCase());
      const esTexto = (f, c) => { const x = ed().celda(h, f, c); return !!x && x.tipo === 'texto'; };
      const pos = { de: { col: rr.c1 + 2, dx: 0, fila: rr.f0, dy: 0 }, a: { col: rr.c1 + 9, dx: 0, fila: rr.f0 + 15, dy: 0 } };
      let indice = -1;
      await A.estructural(tx('cmd-insertar-grafico'), async (b) => { const x = await A.GRA().insertar(b, h, rr, v.tipo, A.GRA().lectorDe(ed()), pos, { grafico: tx('gr-nombre') }, esTexto); indice = x.indice; return x.bytes; });
      if (indice >= 0) { A.avisar(tx('gr-hecho')); elegido = { indice }; setTimeout(colocarMarco, 300); }
    }
    // los datos cambiaron: los gráficos de la hoja se ponen al día (sus cachés en el paquete) y se vuelven a dibujar
    let temporizador = 0;
    A.ganchos.alCambiar.push(() => {
      clearTimeout(temporizador);
      temporizador = setTimeout(async () => {
        const hoja = ed() && ed().hojas[A.hojaActiva()];
        if (!hoja || !(hoja.anclas && hoja.anclas.length)) return;
        try { if (await A.GRA().refrescarCaches(ed().libro.zip, A.GRA().lectorDe(ed()))) await A.rejilla.recargarObjetos(); } catch (e) { console.warn(e); }
      }, 250);
    });

    // ---- página e impresión ---------------------------------------------------------------------------------------------
    const MARGENES = { normal: { izq: 0.7, der: 0.7, sup: 0.75, inf: 0.75, enc: 0.3, pie: 0.3 }, estrecho: { izq: 0.25, der: 0.25, sup: 0.75, inf: 0.75, enc: 0.3, pie: 0.3 }, ancho: { izq: 1, der: 1, sup: 1, inf: 1, enc: 0.5, pie: 0.5 } };
    async function configurarPagina() {
      const h = A.hojaActiva();
      const c = await A.PAG().leer(await ed().guardar(), h);
      const margen = Object.keys(MARGENES).find((k) => JSON.stringify(MARGENES[k]) === JSON.stringify(c.margenes)) || 'propio';
      const ajuste = c.ajuste ? (c.ajuste.ancho === 1 && c.ajuste.alto === 0 ? 'ancho' : c.ajuste.ancho === 1 && c.ajuste.alto === 1 ? 'una' : 'ancho') : c.escala !== 100 ? 'escala' : 'sin';
      const papeles = Object.entries(A.PAG().PAPELES).map(([k, v]) => [k, v]);
      const v = await Dialogo.formulario(Object.assign({}, comun, {
        titulo: tx('p-titulo'),
        campos: [
          { nombre: 'orientacion', etiqueta: tx('p-orientacion'), tipo: 'select', valor: c.orientacion, opciones: [['portrait', tx('p-vertical')], ['landscape', tx('p-horizontal')]] },
          { nombre: 'papel', etiqueta: tx('p-papel'), tipo: 'select', valor: String(c.papel), opciones: papeles.some(([k]) => +k === c.papel) ? papeles : [[String(c.papel), String(c.papel)], ...papeles] },
          { nombre: 'ajuste', etiqueta: tx('p-ajuste'), tipo: 'select', valor: ajuste, opciones: [['sin', tx('p-sin')], ['ancho', tx('p-ancho')], ['una', tx('p-una')], ['escala', tx('p-escala')]] },
          { nombre: 'escala', etiqueta: tx('p-escala-pct'), tipo: 'numero', valor: c.escala, visible: (x) => x.ajuste === 'escala' },
          { nombre: 'margenes', etiqueta: tx('p-margenes'), tipo: 'select', valor: margen, opciones: [...(margen === 'propio' ? [['propio', tx('p-m-propio')]] : []), ['normal', tx('p-m-normal')], ['estrecho', tx('p-m-estrecho')], ['ancho', tx('p-m-ancho')]] },
          { nombre: 'area', etiqueta: tx('p-area'), tipo: 'texto', valor: c.area || '' },
          { nombre: 'titulos', etiqueta: tx('p-titulos'), tipo: 'texto', valor: c.titulosFilas || '' },
          { nombre: 'encabezado', etiqueta: tx('p-encabezado'), tipo: 'texto', valor: c.encabezado },
          { nombre: 'pie', etiqueta: tx('p-pie'), tipo: 'texto', valor: c.pie },
          { nombre: 'cuadricula', etiqueta: tx('p-cuadricula'), tipo: 'check', valor: c.cuadricula },
          { nombre: 'centrar', etiqueta: tx('p-centrar'), tipo: 'check', valor: c.centrarH },
        ],
        validar: (x) => {
          if (x.area && !LibroXlsx.decodificarRango(String(x.area).replace(/\$/g, '').toUpperCase())) return tx('t-rango-malo');
          if (x.titulos && !/^\$?\d+(:\$?\d+)?$/.test(String(x.titulos).trim())) return tx('p-titulos');
          if (x.ajuste === 'escala' && !(x.escala >= 10 && x.escala <= 400)) return tx('p-escala-pct') + ': 10–400';
          return null;
        },
      }));
      if (!v) return;
      const conf = {
        orientacion: v.orientacion, papel: +v.papel, ajuste: v.ajuste === 'ancho' ? { ancho: 1, alto: 0 } : v.ajuste === 'una' ? { ancho: 1, alto: 1 } : null, escala: v.ajuste === 'escala' ? v.escala : 100,
        area: v.area ? String(v.area).replace(/\$/g, '').toUpperCase() : null, titulosFilas: v.titulos ? String(v.titulos).trim() : null, encabezado: v.encabezado, pie: v.pie, cuadricula: v.cuadricula, centrarH: v.centrar,
      };
      if (v.margenes !== 'propio') conf.margenes = MARGENES[v.margenes];
      await A.estructural(tx('cmd-pagina'), (b) => A.PAG().configurar(b, h, conf), { sel: A.seleccion() });
    }
    // imprimir: el PDF del motor (con la configuración de la página) en una ventana nueva, que abre el diálogo de impresión del navegador
    async function imprimir() {
      A.confirmarEdicion();
      A.avisar(tx('p-imprimiendo'));
      const ventana = window.open('', '_blank');
      try {
        const r = await A.pdf(false);
        const url = URL.createObjectURL(new Blob([r.bytes], { type: 'application/pdf' }));
        if (ventana) { ventana.location.href = url; setTimeout(() => { try { ventana.focus(); ventana.print(); } catch (e) { /* el visor de PDF del navegador imprime con su botón */ } }, 800); }
        else { A.avisar(tx('p-bloqueada')); const a = crear('a', { href: url, download: 'hoja.pdf' }); a.click(); }
        setTimeout(() => URL.revokeObjectURL(url), 60000);
      } catch (e) { if (ventana) ventana.close(); A.avisar(tx('aviso-estructura') + ': ' + (globalThis.MensajesHojas ? MensajesHojas.traducir(String(e && e.message ? e.message : e), A.idioma()) : (e && e.message ? e.message : e))); }
    }

    A.reg({ id: 'filtro', texto: 'cmd-filtro', atajo: 'mod+shift+l', ejecutar: alternarFiltro, activo: () => !!(ed() && ed().hojas[A.hojaActiva()] && ed().hojas[A.hojaActiva()].filtroAuto) });
    A.reg({ id: 'borrarFiltros', texto: 'cmd-borrar-filtros', ejecutar: borrarFiltros });
    A.reg({ id: 'ordenarNiveles', texto: 'cmd-ordenar-niveles', ejecutar: ordenarNiveles });
    A.reg({ id: 'insertarGrafico', texto: 'cmd-insertar-grafico', atajo: 'alt+f1', ejecutar: insertarGrafico });
    A.reg({ id: 'editarGrafico', texto: 'cmd-editar-grafico', ejecutar: editarGrafico });
    A.reg({ id: 'pagina', texto: 'cmd-pagina', ejecutar: configurarPagina });
    A.reg({ id: 'imprimir', texto: 'cmd-imprimir', ejecutar: imprimir, habilitado: () => !!ed() });
    A.cinta.push(['datos', { texto: 'g-filtrar', items: [{ comando: 'filtro', icono: '⛛', conTexto: true }, { comando: 'borrarFiltros', icono: '⛛̸' }, { comando: 'ordenarNiveles', icono: '⇅', conTexto: true }] }]);
    A.cinta.push(['insertar', { texto: 'g-graficos', items: [{ comando: 'insertarGrafico', icono: '📊', conTexto: true }, { comando: 'editarGrafico', icono: '✎' }] }]);
    A.cinta.push(['vista', { texto: 'g-pagina', items: [{ comando: 'pagina', icono: '📄', conTexto: true }, { comando: 'imprimir', icono: '⎙', conTexto: true }] }]);
    A.menu.push({ comando: 'imprimir' }, { comando: 'pagina' });
  });
})();

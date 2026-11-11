/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// La interfaz común de Flujos (2026-11-10): la usan el editor (herramientas/flujos.html) y las páginas de los flujos
// listos (herramientas/flujo-*.html), para que las dos muestren las opciones, el diagrama, la ejecución y el resultado
// con el mismo código. Necesita flujos.js, flujos-modulos.js, zona-arrastre.js, descargar.js, zip-propio.js,
// lote-archivos.js e idioma.js (t, idiomaActual, detalleError; idioma.js va al final de la página: llamar a esto
// después de DOMContentLoaded).
//
//   FlujosPagina.CLAVE_ACTUAL               clave de localStorage del flujo abierto en el editor
//   FlujosPagina.texto(obj) / rellenar(clave, valores) / nombreTipos(tipos) / textoError(error)
//   FlujosPagina.campoOpcion(paso, op, alCambiar)   label.campo con el control de una opción (cambia paso.opciones)
//   FlujosPagina.pintarDiagrama(contenedor, flujo)  el diagrama de cajas y flechas (.diagrama-flujo), solo para mirar
//   FlujosPagina.aceptaDe(flujo)            el accept del <input type="file"> según lo que reciben las entradas
//   FlujosPagina.ejecutar({ flujo, archivos, elementos, ocupar })   corre el flujo y pinta el resultado
//       elementos: { estado, avance, resultado, resultadoTexto, descargas, codigos, fallos }; ocupar(si) bloquea la página
//   FlujosPagina.limpiarResultado(elementos)
//   FlujosPagina.montarListo(flujo)         arma la página de un flujo listo (marcado fijo: ver flujo-*.html)
//
// Los textos comunes vienen acá en es y en; los de la página mandan (pueden pisar cualquiera).
// Los textos comunes, en window.i18nPieza: el build (ingles.mjs) los lee de acá para armar la versión /en/ de las
// páginas que cargan esta pieza, y en el navegador se suman a los de la página (los de la página mandan).
window.i18nPieza = {
  es: {
    'titulo-diagrama': 'Así va tu flujo', 'boton-ampliar': '⤢ Ver en grande', 'boton-cerrar': 'Cerrar',
    'diagrama-vacio': 'Agregá pasos para ver el diagrama.', 'diagrama-entrada': '📂 Entrada {n}: {tipos}',
    'diagrama-juntar': '🔀 Se juntan las entradas', 'diagrama-resultado': '📦 Resultado: {tipos}', 'diagrama-sin-pasos': '(sin pasos)',
    'tipo-imagen': 'imágenes', 'tipo-pdf': 'PDF', 'tipo-documento': 'documentos', 'tipo-texto': 'texto', 'tipo-o': ' o ', 'tipo-nada': 'nada',
    'forma-juntar': 'junta todo en uno', 'forma-abrir': 'abre cada archivo en varios', 'paso-n': 'Paso {n}',
    'error-tipo': 'El paso {n} de la entrada {e} no recibe lo que entrega el anterior: cambiá el orden o quitalo.',
    'error-tipo-comun': 'El paso {n} de «Después de juntar» no recibe todo lo que llega de las entradas.',
    'error-juntar': 'El paso {n} de la entrada {e} junta todo en uno: con varias entradas, ponelo en «Después de juntar las entradas».',
    'error-entrada-vacia': 'La entrada {e} no tiene pasos y no hay pasos después de juntar.',
    'archivos-n': '{n} archivos listos.', 'archivos-ignorados': ' {n} no los recibe ninguna entrada y se van a saltar.',
    'boton-ejecutar': 'Ejecutar el flujo', 'procesando': 'Procesando {archivo} ({donde}, paso {paso})…',
    'donde-entrada': 'entrada {n}', 'donde-comun': 'después de juntar', 'preparando': 'Preparando los pasos…', 'falta-archivos': 'Elegí los archivos.',
    'resultado-n': '✅ Listo: {ok} resultados.', 'resultado-cero': 'No salió ningún resultado.',
    'boton-descargar': 'Descargar', 'boton-descargar-zip': 'Descargar todo (.zip)',
    'fallos-titulo': 'Lo que no pasó:', 'fallo-tipo': 'ninguna entrada recibe este tipo de archivo',
    'fallo-quitado': 'quitado en el paso {n}', 'fallo-paso': 'paso {n}: {error}',
    'titulo-codigos': 'Códigos encontrados', 'sin-codigos': 'sin código',
    // Páginas de flujos listos
    'listo-zona': '📂 Arrastrá tus archivos acá o hacé clic para elegirlos',
    'listo-titulo-pasos': 'Los pasos de este flujo',
    'listo-nota-opciones': 'Podés cambiar las opciones de cada paso antes de ejecutar.',
    'boton-personalizar': '✏️ Personalizar este flujo',
    'listo-recibe': 'Recibe {tipos}.',
  },
  en: {
    'titulo-diagrama': 'Your flow at a glance', 'boton-ampliar': '⤢ See it large', 'boton-cerrar': 'Close',
    'diagrama-vacio': 'Add steps to see the diagram.', 'diagrama-entrada': '📂 Input {n}: {tipos}',
    'diagrama-juntar': '🔀 The inputs join', 'diagrama-resultado': '📦 Result: {tipos}', 'diagrama-sin-pasos': '(no steps)',
    'tipo-imagen': 'images', 'tipo-pdf': 'PDF', 'tipo-documento': 'documents', 'tipo-texto': 'text', 'tipo-o': ' or ', 'tipo-nada': 'nothing',
    'forma-juntar': 'joins everything into one', 'forma-abrir': 'opens each file into several', 'paso-n': 'Step {n}',
    'error-tipo': 'Step {n} of input {e} does not take what the previous one produces: change the order or remove it.',
    'error-tipo-comun': 'Step {n} of “After the inputs join” does not take everything that comes from the inputs.',
    'error-juntar': 'Step {n} of input {e} joins everything into one: with several inputs, put it in “After the inputs join”.',
    'error-entrada-vacia': 'Input {e} has no steps and there are no steps after joining.',
    'archivos-n': '{n} files ready.', 'archivos-ignorados': ' {n} are not taken by any input and will be skipped.',
    'boton-ejecutar': 'Run the flow', 'procesando': 'Processing {archivo} ({donde}, step {paso})…',
    'donde-entrada': 'input {n}', 'donde-comun': 'after joining', 'preparando': 'Getting the steps ready…', 'falta-archivos': 'Choose the files.',
    'resultado-n': '✅ Done: {ok} results.', 'resultado-cero': 'No result came out.',
    'boton-descargar': 'Download', 'boton-descargar-zip': 'Download all (.zip)',
    'fallos-titulo': 'What did not go through:', 'fallo-tipo': 'no input takes this type of file',
    'fallo-quitado': 'removed at step {n}', 'fallo-paso': 'step {n}: {error}',
    'titulo-codigos': 'Codes found', 'sin-codigos': 'no code',
    'listo-zona': '📂 Drag your files here or click to choose them',
    'listo-titulo-pasos': 'The steps of this flow',
    'listo-nota-opciones': 'You can change the options of each step before running it.',
    'boton-personalizar': '✏️ Customize this flow',
    'listo-recibe': 'Takes {tipos}.',
  },
};

(function () {
  const F = window.Flujos;

  (function () {
    const pagina = window.i18nPagina || (window.i18nPagina = {});
    pagina.es = Object.assign({}, window.i18nPieza.es, pagina.es);
    pagina.en = Object.assign({}, window.i18nPieza.en, pagina.en);
  })();

  const CLAVE_ACTUAL = 'flujos-actual';
  const idioma = () => (window.idiomaActual ? idiomaActual() : 'es');
  const texto = (obj) => obj[idioma()] || obj.es;
  const rellenar = (clave, valores) => Object.entries(valores).reduce((s, [k, v]) => s.split('{' + k + '}').join(v), t(clave));
  const el = (etiqueta, clase, contenido) => { const e = document.createElement(etiqueta); if (clase) e.className = clase; if (contenido !== undefined) e.textContent = contenido; return e; };

  // «imágenes», «imágenes o PDF», «imágenes, PDF o documentos»… en un orden fijo.
  const ORDEN_TIPOS = ['imagen', 'pdf', 'documento', 'texto'];
  function nombreTipos(tipos) {
    if (!tipos || !tipos.length) return t('tipo-nada');
    const nombres = ORDEN_TIPOS.filter((x) => tipos.includes(x)).map((x) => t('tipo-' + x));
    return nombres.length === 1 ? nombres[0] : nombres.slice(0, -1).join(', ') + t('tipo-o') + nombres[nombres.length - 1];
  }

  function textoError(e) {
    if (e.motivo === 'juntar') return rellenar('error-juntar', { n: e.paso + 1, e: e.entrada + 1 });
    if (e.motivo === 'vacia') return rellenar('error-entrada-vacia', { e: e.entrada + 1 });
    if (e.entrada === 'comun') return rellenar('error-tipo-comun', { n: e.paso + 1 });
    return rellenar('error-tipo', { n: e.paso + 1, e: e.entrada + 1 });
  }

  // ---------- opciones de un paso ----------
  function campoOpcion(paso, op, alCambiar) {
    const avisar = alCambiar || (() => {});
    const label = el('label', 'campo');
    const etiqueta = el('span', '', texto(op.etiqueta));
    const limpio = (valor) => F.normalizar([{ id: paso.id, opciones: { ...paso.opciones, [op.id]: valor } }]).entradas[0][0].opciones;
    let control;
    if (op.tipo === 'select') {
      control = document.createElement('select');
      op.valores.forEach((v) => { const o = el('option', '', texto(v)); o.value = v.valor; control.appendChild(o); });
      control.value = paso.opciones[op.id];
      control.addEventListener('change', () => { paso.opciones[op.id] = control.value; avisar(); });
    } else if (op.tipo === 'texto' || op.tipo === 'color' || op.tipo === 'numero') {
      control = document.createElement('input');
      control.type = op.tipo === 'color' ? 'color' : op.tipo === 'numero' ? 'number' : 'text';
      if (op.tipo === 'texto') control.maxLength = op.max || 200;
      if (op.tipo === 'numero') Object.assign(control, { min: op.min, max: op.max, step: op.paso || 1 });
      control.value = paso.opciones[op.id];
      control.addEventListener('change', () => { paso.opciones = limpio(control.value); control.value = paso.opciones[op.id]; avisar(); });
    } else {
      control = document.createElement('input');
      control.type = 'checkbox';
      control.checked = !!paso.opciones[op.id];
      control.addEventListener('change', () => { paso.opciones[op.id] = control.checked; avisar(); });
      // Mismo marcado que las casillas de Rotar imagen: <label class="campo"><span><input> <span>texto</span></span></label>
      const fila = el('span');
      fila.append(control, ' ', etiqueta);
      label.appendChild(fila);
      return label;
    }
    label.append(etiqueta, control);
    return label;
  }

  // ---------- diagrama (solo para mirar) ----------
  function pintarDiagrama(cont, flujo) {
    cont.innerHTML = '';
    if (!flujo.entradas.some((p) => p.length) && !flujo.comun.length) { cont.appendChild(el('p', 'nota-suave', t('diagrama-vacio'))); return; }
    const error = F.validar(flujo);
    const nodo = (txt, clase) => el('li', 'diagrama-nodo' + (clase ? ' ' + clase : ''), txt);
    const flecha = () => { const f = el('li', 'diagrama-flecha'); f.setAttribute('aria-hidden', 'true'); return f; };
    const rama = (pasos, donde, cabeza) => {
      const ul = el('ul', 'diagrama-rama');
      if (cabeza) ul.appendChild(nodo(cabeza, 'punta'));
      if (!pasos.length && cabeza) { ul.append(flecha(), nodo(t('diagrama-sin-pasos'))); }
      pasos.forEach((p, k) => {
        const m = F.modulo(p.id);
        if (cabeza || k) ul.appendChild(flecha());
        const malo = error && error.entrada === donde && error.paso === k;
        ul.appendChild(nodo(`${malo ? '⚠️ ' : ''}${m.icono} ${texto(m.nombre)}`, malo ? 'con-error' : ''));
      });
      return ul;
    };
    const filas = el('div', 'diagrama-entradas');
    flujo.entradas.forEach((pasos, i) => filas.appendChild(rama(pasos, i, rellenar('diagrama-entrada', { n: i + 1, tipos: nombreTipos(F.tiposDeEntrada(flujo, i)) }))));
    cont.appendChild(filas);
    const final = el('ul', 'diagrama-rama');
    if (flujo.entradas.length > 1) {
      const union = el('div', 'diagrama-union'); union.setAttribute('aria-hidden', 'true');
      cont.appendChild(union);
      final.append(nodo(t('diagrama-juntar'), 'punta'));
      flujo.comun.forEach((p, k) => {
        const m = F.modulo(p.id);
        const malo = error && error.entrada === 'comun' && error.paso === k;
        final.append(flecha(), nodo(`${malo ? '⚠️ ' : ''}${m.icono} ${texto(m.nombre)}`, malo ? 'con-error' : ''));
      });
    }
    final.append(flecha(), nodo(rellenar('diagrama-resultado', { tipos: nombreTipos(error ? [] : F.salidaDe(flujo)) }), 'punta'));
    cont.appendChild(final);
  }

  // El accept del selector según lo que reciben las entradas (los documentos, con la lista de archivo-a-pdf.js).
  function aceptaDe(flujo) {
    const tipos = new Set(flujo.entradas.flatMap((_, i) => F.tiposDeEntrada(flujo, i)));
    const partes = [];
    if (tipos.has('imagen')) partes.push('image/*');
    if (tipos.has('pdf')) partes.push('.pdf,application/pdf');
    if ((tipos.has('documento') || tipos.has('texto')) && window.ArchivoAPdf) partes.push(ArchivoAPdf.EXTENSIONES);
    return partes.join(',');
  }

  // ---------- ejecutar y mostrar el resultado ----------
  function limpiarResultado(e) {
    (e._urls || []).forEach((u) => URL.revokeObjectURL(u));
    e._urls = [];
    e.resultado.classList.remove('visible');
  }

  function enlaceDescarga(e, blob, nombre, clave) {
    const a = el('a', 'boton', clave === 'boton-descargar' ? `${t(clave)}: ${nombre}` : t(clave));
    a.href = URL.createObjectURL(blob);
    (e._urls || (e._urls = [])).push(a.href);
    a.download = nombre;
    return a;
  }

  const donde = (entrada) => (entrada === 'comun' ? t('donde-comun') : rellenar('donde-entrada', { n: entrada + 1 }));

  function mostrarEstado(e, msg, error) {
    e.estado.textContent = msg; e.estado.hidden = !msg;
    e.estado.classList.toggle('error', !!error);
  }

  async function pintarResultado(e, flujo, resultados, fallos) {
    const varias = flujo.entradas.length > 1;
    const textoFallo = (f) => {
      if (f.codigo === 'tipo') return t('fallo-tipo');
      const en = varias ? donde(f.entrada) + ', ' : '';
      if (f.codigo === 'quitado') return en + rellenar('fallo-quitado', { n: f.paso + 1 });
      return en + rellenar('fallo-paso', { n: f.paso + 1, error: f.error });
    };
    e.resultadoTexto.textContent = resultados.length ? rellenar('resultado-n', { ok: resultados.length }) : t('resultado-cero');
    e.descargas.innerHTML = '';
    if (resultados.length === 1) e.descargas.appendChild(enlaceDescarga(e, resultados[0].blob, resultados[0].nombre, 'boton-descargar'));
    else if (resultados.length > 1) {
      const zip = await LoteArchivos.zipDeResultados(resultados.map((r) => ({ ok: true, nombre: r.nombre, blob: r.blob })));
      e.descargas.appendChild(enlaceDescarga(e, zip, nombreConFecha('flujo', 'zip'), 'boton-descargar-zip'));
    }
    // Lo que leyó el paso de códigos, archivo por archivo (texto de afuera: siempre con textContent)
    e.codigos.innerHTML = '';
    const conLector = [...flujo.entradas.flat(), ...flujo.comun].some((p) => p.id === 'leer-codigos');
    if (conLector && resultados.length) {
      const ul = el('ul');
      resultados.forEach((r) => {
        const lista = (r.datos.codigos || []).map((c) => c.texto);
        ul.appendChild(el('li', '', `${r.nombre}: ${lista.length ? lista.join(' · ') : t('sin-codigos')}`));
      });
      e.codigos.append(el('h3', '', t('titulo-codigos')), ul);
    }
    e.fallos.innerHTML = '';
    if (fallos.length) {
      const ul = el('ul');
      fallos.forEach((f) => ul.appendChild(el('li', '', `${f.nombre}: ${textoFallo(f)}`)));
      e.fallos.append(el('p', '', t('fallos-titulo')), ul);
    }
    e.resultado.classList.add('visible');
  }

  async function ejecutar({ flujo, archivos, elementos: e, ocupar }) {
    if (!archivos.length) { mostrarEstado(e, t('falta-archivos')); return; }
    limpiarResultado(e);
    ocupar(true);
    e.avance.hidden = false; e.avance.removeAttribute('value');
    mostrarEstado(e, t('preparando'));
    try {
      const { resultados, fallos } = await F.ejecutar(flujo, archivos, {
        alAvanzar: ({ hechos, total, archivo, entrada, paso }) => {
          if (total) { e.avance.max = total; e.avance.value = hechos; }
          mostrarEstado(e, rellenar('procesando', { archivo: archivo || '…', donde: donde(entrada), paso: paso + 1 }));
        },
      });
      await pintarResultado(e, flujo, resultados, fallos);
      mostrarEstado(e, '');
    } catch (error) {
      mostrarEstado(e, error.codigo ? textoError({ motivo: error.codigo, entrada: error.entrada, paso: error.paso }) : detalleError(error), true);
    } finally {
      e.avance.hidden = true;
      ocupar(false);
    }
  }

  // ---------- página de un flujo listo ----------
  // Marcado que espera (ids): zonaArrastre + entradaArchivos, avisoArchivos, recibe, listaPasos, diagramaMini,
  // botonEjecutar, botonPersonalizar, estado, avance, resultado (con resultadoTexto, descargas, codigos, fallos).
  function montarListo(base) {
    const $ = (id) => document.getElementById(id);
    let flujo = F.normalizar(JSON.parse(JSON.stringify(base)));
    let archivos = [];
    let ocupado = false;
    const elementos = { estado: $('estado'), avance: $('avance'), resultado: $('resultado'), resultadoTexto: $('resultadoTexto'), descargas: $('descargas'), codigos: $('codigos'), fallos: $('fallos') };
    $('entradaArchivos').accept = aceptaDe(flujo);

    function pintar() {
      const lista = $('listaPasos');
      lista.innerHTML = '';
      const pasos = [...flujo.entradas[0], ...flujo.comun];
      pasos.forEach((paso, k) => {
        const m = F.modulo(paso.id);
        const li = el('li', 'paso-flujo');
        const cabecera = el('div', 'paso-cabecera');
        const extra = m.forma === 'juntar' ? ` (${t('forma-juntar')})` : m.forma === 'abrir' ? ` (${t('forma-abrir')})` : '';
        cabecera.appendChild(el('strong', '', `${rellenar('paso-n', { n: k + 1 })} · ${m.icono} ${texto(m.nombre)}${extra}`));
        li.append(cabecera, el('p', 'nota-suave', texto(m.descripcion)));
        if ((m.opciones || []).length) {
          const controles = el('div', 'controles');
          m.opciones.forEach((op) => {
            const campo = campoOpcion(paso, op, () => limpiarResultado(elementos));
            campo.querySelectorAll('select, input').forEach((c) => { c.disabled = ocupado; });
            controles.appendChild(campo);
          });
          li.appendChild(controles);
        }
        lista.appendChild(li);
      });
      $('recibe').textContent = rellenar('listo-recibe', { tipos: nombreTipos(F.tiposDeEntrada(flujo, 0)) });
      const reparto = F.repartir(flujo, archivos);
      const aviso = $('avisoArchivos');
      aviso.textContent = archivos.length ? rellenar('archivos-n', { n: archivos.length - reparto.sinEntrada.length }) + (reparto.sinEntrada.length ? rellenar('archivos-ignorados', { n: reparto.sinEntrada.length }) : '') : '';
      aviso.hidden = !archivos.length;
      $('botonEjecutar').disabled = ocupado || !archivos.length || !!F.validar(flujo);
      $('botonPersonalizar').disabled = ocupado;
      pintarDiagrama($('diagramaMini'), flujo);
    }

    conectarZonaArrastre($('zonaArrastre'), $('entradaArchivos'), (lista) => {
      if (ocupado) return;
      archivos = lista;
      limpiarResultado(elementos);
      pintar();
    });
    $('botonEjecutar').addEventListener('click', () => ejecutar({
      flujo, archivos, elementos,
      ocupar: (si) => { ocupado = si; pintar(); },
    }));
    // Abre el editor con estos pasos (y las opciones elegidas): el editor lee el flujo abierto de localStorage.
    $('botonPersonalizar').addEventListener('click', () => {
      try { localStorage.setItem(CLAVE_ACTUAL, JSON.stringify(flujo)); } catch (e) { /* sin almacenamiento: abre el editor como esté */ }
      location.href = 'flujos.html';
    });
    document.addEventListener('cambioIdioma', pintar);
    pintar();
  }

  window.FlujosPagina = { CLAVE_ACTUAL, texto, rellenar, nombreTipos, textoError, campoOpcion, pintarDiagrama, aceptaDe, ejecutar, limpiarResultado, montarListo };
})();

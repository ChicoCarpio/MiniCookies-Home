/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Motor de Flujos: encadenar herramientas del sitio. No dibuja nada: la página arma la interfaz.
//
// Un flujo tiene una o varias ENTRADAS (cada una con sus archivos y su lista de pasos) que, si son varias, se juntan
// en un tramo COMÚN (p. ej. «Unir en un solo PDF»):
//   { entradas: [[paso, …], [paso, …]], comun: [paso, …] }      paso = { id, opciones }
//   imágenes > QR > recortar > PDF ─┐
//                                   ├─> unir PDF
//   PDF > a imágenes > QR > recortar > PDF ─┘
// Con una sola entrada el tramo común no hace falta (sus pasos pueden ir en la entrada).
//
//   Flujos.registrar(modulo)                 da de alta un módulo (ver flujos-modulos.js)
//   Flujos.modulos() / Flujos.modulo(id)     los módulos registrados
//   Flujos.compatibles(tipos)                módulos que reciben TODOS esos tipos; sin tipos: todos
//   Flujos.tiposDeEntrada(flujo, i)          tipos de archivo que pueden recorrer la entrada i (y el tramo común)
//   Flujos.tiposAlFinal(pasos, tipos)        tipos que salen de una lista de pasos si le entran esos tipos
//   Flujos.tiposTrasPasos(pasos)             tipos que siguen al final de los pasos de una entrada (para ofrecer el siguiente)
//   Flujos.tiposAlComun(flujo)             tipos que llegan al tramo común (lo que sale de todas las entradas)
//   Flujos.salidaDe(flujo)                   tipos que entrega el flujo entero
//   Flujos.validar(flujo)                    null si se puede correr, o { entrada (número o 'comun'), paso, motivo }
//                                            motivo: 'vacia' | 'modulo' | 'tipo' | 'juntar' (un paso que junta, dentro
//                                            de una entrada de un flujo con varias: tiene que ir en el tramo común)
//   Flujos.repartir(flujo, archivos)         { porEntrada, sinEntrada }: cada archivo va a las entradas que reciben su tipo
//   Flujos.ejecutar(flujo, archivos, { alAvanzar }) -> Promise<{ resultados, fallos }>   (los reparte con repartir)
//   Flujos.tipoDeArchivo(archivo)            'imagen' | 'pdf' | 'documento' | 'texto' | null
//                                            documento: Word, Excel/CSV, PowerPoint, EPUB o Markdown (los formatos de
//                                            ArchivoAPdf, que la página carga); texto: .txt
//   Flujos.normalizar(flujo)                 flujo limpio; acepta también la forma vieja (una lista de pasos)
//   Flujos.guardadas() / guardar(nombre, flujo) / borrar(nombre)       en este navegador (localStorage)
//   Flujos.exportar(nombre, flujo) -> Blob   archivo .json para guardar o pasar a otro equipo
//   Flujos.importar(texto) -> { nombre, flujo, omitidos }   lanza un Error con `codigo` si el archivo no sirve
//
// Lo que viaja de un paso al siguiente es una ficha de archivo:
//   { nombre, blob, tipo: 'imagen'|'pdf'|'documento'|'texto', datos: { codigos: [...], ... }, origen }
// `datos` se conserva de paso en paso: así el QR que se leyó en el paso 1 llega al PDF del paso 3.
//
// Un módulo: { id, icono, nombre:{es,en}, descripcion:{es,en}, entra:[tipos], sale:tipo|'igual',
//   forma:'uno'|'abrir'|'juntar', opciones:[{ id, tipo:'select'|'numero'|'casilla'|'texto'|'color', etiqueta:{es,en},
//   valores:[{ valor, es, en }], defecto, min, max, paso }], preparar() -> Promise (carga lo que necesita),
//   ejecutar(ficha | fichas, opciones, contexto) -> Promise<ficha | fichas | null>,
//   terminar() -> Promise (opcional: libera lo que preparar o ejecutar dejaron abierto, p. ej. el motor de OCR),
//   hilo: true (opcional: puede correr en un hilo de trabajo, flujos-hilo.js; solo si su código y el de sus piezas no
//   necesitan la página: nada de DOMParser ni elementos de verdad; canvas e imágenes sí),
//   simultaneos: n | () => n (opcional, para los que corren en la página: cuántos archivos a la vez admite) }.
//   'uno' recibe una ficha y devuelve una (o null para quitarla); 'abrir' recibe una y devuelve varias (un PDF ->
//   sus páginas); 'juntar' recibe todas las que llegan y devuelve una.
// Todo corre en este navegador. Varios archivos avanzan a la vez, repartidos entre hilos de trabajo (uno por núcleo)
// y solo unos pocos más que hilos, así un lote grande no ocupa la memoria de todos juntos.
(function () {
  const registro = new Map();
  const CLAVE = 'flujos-guardados';
  const FORMATO = 'minicookies-flujo';
  const TIPOS = ['imagen', 'pdf', 'documento', 'texto'];
  const MAX_ENTRADAS = 4, MAX_PASOS = 20;

  function registrar(m) {
    if (!m || !m.id || registro.has(m.id)) throw new Error('Módulo inválido o repetido: ' + (m && m.id));
    registro.set(m.id, m);
  }
  const modulos = () => [...registro.values()];
  const modulo = (id) => registro.get(id) || null;

  function tipoDeArchivo(archivo) {
    const tipo = (archivo.type || '').toLowerCase(), nombre = (archivo.name || '').toLowerCase();
    if (tipo === 'application/pdf' || nombre.endsWith('.pdf')) return 'pdf';
    if (tipo.startsWith('image/') || /\.(jpe?g|png|webp|gif|bmp)$/.test(nombre)) return 'imagen';
    const formato = window.ArchivoAPdf && archivo.name ? ArchivoAPdf.formatoDe(archivo) : null;
    if (formato) return formato === 'texto' ? 'texto' : 'documento';
    return null;
  }

  // ---------- tipos ----------
  const salida = (m, tipo) => (m.sale === 'igual' ? tipo : m.sale);

  /** Módulos que reciben todos los tipos dados (lo que llega puede ser una mezcla, p. ej. al tramo común). */
  function compatibles(tipos) {
    const t = Array.isArray(tipos) ? tipos : tipos ? [tipos] : [];
    return modulos().filter((m) => t.every((x) => m.entra.includes(x)));
  }

  /** Tipos que salen de `pasos` si le entran `tipos`; null si algún paso no recibe lo que le llega. */
  function tiposAlFinal(pasos, tipos) {
    let t = [...new Set(tipos)];
    for (const p of pasos) {
      const m = modulo(p.id);
      if (!m || !t.length || !t.every((x) => m.entra.includes(x))) return null;
      t = [...new Set(t.map((x) => salida(m, x)))];
    }
    return t;
  }

  /**
   * Tipos que siguen al final de los pasos de una entrada, descartando en cada paso lo que no recibe (lo que
   * no pase, se avisa al cargar los archivos). Sin pasos: todos los tipos. Sirve para ofrecer el paso siguiente.
   */
  function tiposTrasPasos(pasos) {
    let t = pasos.length && modulo(pasos[0].id) ? modulo(pasos[0].id).entra.slice() : TIPOS.slice();
    for (const p of pasos) {
      const m = modulo(p.id);
      if (!m) return [];
      t = t.filter((x) => m.entra.includes(x));
      if (!t.length) return [];
      t = [...new Set(t.map((x) => salida(m, x)))];
    }
    return t;
  }

  /** Tipos de archivo de la entrada i que pueden llegar hasta el final (pasando también por el tramo común). */
  function tiposDeEntrada(flujo, i) {
    if (Array.isArray(flujo)) flujo = normalizar(flujo);
    const pasos = flujo.entradas[i] || [];
    return TIPOS.filter((tipo) => {
      const fin = tiposAlFinal(pasos, [tipo]);
      return fin && (!flujo.comun.length || tiposAlFinal(flujo.comun, fin));
    });
  }

  /**
   * Reparte los archivos de la única zona de la página entre las entradas, según el tipo que recibe cada una.
   * Si dos entradas reciben el mismo tipo, el archivo va a las dos. Devuelve { porEntrada: [[archivos]…], sinEntrada }.
   */
  function repartir(flujo, archivos) {
    if (Array.isArray(flujo)) flujo = normalizar(flujo);
    const tipos = flujo.entradas.map((_, i) => tiposDeEntrada(flujo, i));
    const porEntrada = flujo.entradas.map(() => []), sinEntrada = [];
    for (const a of archivos) {
      const tipo = tipoDeArchivo(a);
      let alguna = false;
      tipos.forEach((t, i) => { if (tipo && t.includes(tipo)) { porEntrada[i].push(a); alguna = true; } });
      if (!alguna) sinEntrada.push(a);
    }
    return { porEntrada, sinEntrada };
  }

  function tiposAlComun(flujo) {
    if (Array.isArray(flujo)) flujo = normalizar(flujo);
    const t = new Set();
    flujo.entradas.forEach((pasos, i) => tiposDeEntrada(flujo, i).forEach((tipo) => (tiposAlFinal(pasos, [tipo]) || []).forEach((x) => t.add(x))));
    return [...t];
  }

  function salidaDe(flujo) {
    if (Array.isArray(flujo)) flujo = normalizar(flujo);
    const alComun = tiposAlComun(flujo);
    return flujo.comun.length ? tiposAlFinal(flujo.comun, alComun) || [] : alComun;
  }

  function validar(flujo) {
    if (Array.isArray(flujo)) flujo = normalizar(flujo);
    const varias = flujo.entradas.length > 1;
    if (flujo.entradas.every((p) => !p.length) && !flujo.comun.length) return { entrada: 0, paso: -1, motivo: 'vacia' };
    for (let i = 0; i < flujo.entradas.length; i++) {
      const pasos = flujo.entradas[i];
      if (!pasos.length && !flujo.comun.length) return { entrada: i, paso: -1, motivo: 'vacia' };
      for (let k = 0; k < pasos.length; k++) {
        const m = modulo(pasos[k].id);
        if (!m) return { entrada: i, paso: k, motivo: 'modulo' };
        if (varias && m.forma === 'juntar') return { entrada: i, paso: k, motivo: 'juntar' };
      }
      // El primer paso que no recibe lo que le llega (con cualquiera de los tipos que admite el primero).
      let t = pasos.length ? modulo(pasos[0].id).entra.slice() : TIPOS.slice();
      for (let k = 0; k < pasos.length; k++) {
        const m = modulo(pasos[k].id);
        const sirven = t.filter((x) => m.entra.includes(x));
        if (!sirven.length) return { entrada: i, paso: k, motivo: 'tipo' };
        t = [...new Set(sirven.map((x) => salida(m, x)))];
      }
      if (!tiposDeEntrada(flujo, i).length) return { entrada: 'comun', paso: 0, motivo: 'tipo' };
    }
    let t = tiposAlComun(flujo);
    for (let k = 0; k < flujo.comun.length; k++) {
      const m = modulo(flujo.comun[k].id);
      if (!m) return { entrada: 'comun', paso: k, motivo: 'modulo' };
      if (!t.every((x) => m.entra.includes(x))) return { entrada: 'comun', paso: k, motivo: 'tipo' };
      t = [...new Set(t.map((x) => salida(m, x)))];
    }
    return null;
  }

  // ---------- normalizar ----------
  // Opciones dentro de lo que el módulo permite; lo que falte o no sirva vuelve a su valor por defecto.
  function opcionesValidas(m, dadas) {
    const o = {};
    for (const op of m.opciones || []) {
      let v = dadas && Object.prototype.hasOwnProperty.call(dadas, op.id) ? dadas[op.id] : op.defecto;
      if (op.tipo === 'select' && !op.valores.some((x) => x.valor === v)) v = op.defecto;
      if (op.tipo === 'casilla') v = v === true || v === 'true';
      if (op.tipo === 'color' && !/^#[0-9a-f]{6}$/i.test(String(v))) v = op.defecto;
      if (op.tipo === 'texto') {
        v = typeof v === 'string' ? v.slice(0, op.max || 200) : op.defecto;
        if (!v.trim()) v = op.defecto;   // vacío: vuelve al texto por defecto
      }
      if (op.tipo === 'numero') {
        v = Number(v);
        if (!Number.isFinite(v)) v = op.defecto;
        v = Math.min(op.max ?? v, Math.max(op.min ?? v, v));
      }
      o[op.id] = v;
    }
    return o;
  }

  function normalizarPasos(pasos) {
    return (Array.isArray(pasos) ? pasos : [])
      .filter((p) => p && typeof p.id === 'string' && registro.has(p.id))
      .slice(0, MAX_PASOS)
      .map((p) => ({ id: p.id, opciones: opcionesValidas(registro.get(p.id), p.opciones) }));
  }

  function normalizar(flujo) {
    if (Array.isArray(flujo)) return { entradas: [normalizarPasos(flujo)], comun: [] };   // forma vieja: una lista
    const entradas = (flujo && Array.isArray(flujo.entradas) ? flujo.entradas : [[]]).slice(0, MAX_ENTRADAS).map(normalizarPasos);
    return { entradas: entradas.length ? entradas : [[]], comun: normalizarPasos(flujo && flujo.comun) };
  }

  // ---------- ejecutar ----------
  function base(nombre) { const i = nombre.lastIndexOf('.'); return i > 0 ? nombre.slice(0, i) : nombre; }

  function nombreUnico(nombre, usados) {
    let n = nombre, k = 2;
    while (usados.has(n)) { const i = nombre.lastIndexOf('.'); n = i > 0 ? `${nombre.slice(0, i)}-${k}${nombre.slice(i)}` : `${nombre}-${k}`; k++; }
    usados.add(n);
    return n;
  }

  // ---------- hilos (2026-11-10) ----------
  // Los módulos con `hilo: true` corren en hilos de trabajo (flujos-hilo.js), uno por núcleo menos uno (la página
  // sigue respondiendo). Los demás corren en la página: los que declaran `simultaneos` (p. ej. el OCR, cuyo trabajo
  // pesado ya va en los hilos de Tesseract) de a varios; el resto, de a uno y nunca dos a la vez aunque sean de
  // módulos distintos (el mismo cuidado que antes, cuando todo iba de a un archivo).
  // En la PC: con 120 fotos (leer QR > recortar > comprimir, Ryzen 9 9900X de 12 núcleos) 4 hilos dieron 2,8x, 8 hilos
  // 4,3x y 12 hilos 5,7x; 16 igual que 12 y con más memoria (Investigaciones\investigacion-flujos\probar-multihilo.mjs).
  const TOPE_HILOS = 12;
  function cuantosHilos() {
    const nucleos = navigator.hardwareConcurrency || 2;
    // El iPad con Safari dice «Macintosh» y no tiene userAgentData: se lo reconoce por la pantalla táctil.
    const movil = navigator.userAgentData ? navigator.userAgentData.mobile
      : /Mobi|Android|iPhone|iPad/i.test(navigator.userAgent) || (/Macintosh/.test(navigator.userAgent) && navigator.maxTouchPoints > 1);
    // En el celular la memoria manda: cada hilo carga sus motores (Custos, PDF.js) y una foto grande ocupa ~64 MB.
    const tope = movil ? Math.max(1, Math.min(3, Math.floor((navigator.deviceMemory || 4) / 2))) : TOPE_HILOS;
    return Math.max(1, Math.min(nucleos - 1, tope));
  }

  // Errores que dicen «esto no anda en el hilo» (el entorno), no «este archivo está mal»: el paso se repite en la página.
  const ERRORES_DE_ENTORNO = ['TypeError', 'ReferenceError', 'SecurityError', 'NotSupportedError', 'InvalidStateError', 'DataCloneError'];
  const FALLAS_SEGUIDAS_MAX = 3;   // un módulo que falla así 3 veces seguidas deja de ir a los hilos en esta ejecución
  // Un paso que no contesta en este tiempo da el hilo por trabado: se cierra y el paso se repite en la página.
  const ESPERA_MAXIMA_PASO = 10 * 60 * 1000;

  function semaforo(cupos) {
    let libres = cupos;
    const cola = [];
    return async function (fn) {
      if (libres > 0) libres--;
      else await new Promise((ok) => cola.push(ok));
      try { return await fn(); } finally {
        if (cola.length) cola.shift()();
        else libres++;
      }
    };
  }

  async function abrirHilos(cuantos) {
    if (typeof Worker !== 'function' || typeof OffscreenCanvas !== 'function' || location.protocol === 'file:') return null;
    const scripts = ['vendor.js', 'flujos.js', 'flujos-modulos.js'].map(rutaJs);
    const idioma = window.idiomaActual && idiomaActual() === 'en' ? 'en' : 'es';
    const nuevos = [];
    for (let i = 0; i < cuantos; i++) { try { nuevos.push(new Worker(rutaJs('flujos-hilo.js'))); } catch (e) { break; } }
    const listos = await Promise.all(nuevos.map((w) => new Promise((ok) => {
      const espera = setTimeout(() => ok(false), 20000);
      const fin = (bien) => { clearTimeout(espera); ok(bien); };
      w.onmessage = (e) => fin(!!e.data && e.data.tipo === 'listo');
      w.onerror = (e) => { if (e.preventDefault) e.preventDefault(); fin(false); };
      w.postMessage({ tipo: 'iniciar', scripts, idioma });
    })));
    const vivos = nuevos.filter((w, i) => listos[i] || (w.terminate(), false));
    return vivos.length ? grupoDeHilos(vivos) : null;
  }

  function grupoDeHilos(trabajadores) {
    const libres = trabajadores.slice(), cola = [], enCurso = new Map();
    const pasos = new Map(trabajadores.map((w) => [w, 0]));
    const sinHilo = new Set();   // módulos que no pudieron prepararse en un hilo: siguen en la página
    let siguiente = 1;
    function despachar() {
      while (libres.length && cola.length) {
        const w = libres.shift(), tarea = cola.shift();
        enCurso.set(w, tarea);
        tarea.reloj = setTimeout(() => caido(w), ESPERA_MAXIMA_PASO);
        try { w.postMessage(tarea.mensaje); } catch (e) { clearTimeout(tarea.reloj); enCurso.delete(w); libres.push(w); tarea.mal(e); }
      }
    }
    function caido(w) {
      const tarea = enCurso.get(w);
      if (tarea) clearTimeout(tarea.reloj);
      enCurso.delete(w);
      w.terminate();
      const i = trabajadores.indexOf(w);
      if (i >= 0) trabajadores.splice(i, 1);
      const j = libres.indexOf(w);
      if (j >= 0) libres.splice(j, 1);
      const error = Object.assign(new Error('El hilo se cerró'), { name: 'InvalidStateError' });
      if (tarea) tarea.mal(error);
      if (!trabajadores.length) while (cola.length) cola.shift().mal(error);
    }
    for (const w of trabajadores) {
      w.onmessage = (e) => {
        const tarea = enCurso.get(w);
        if (!tarea) return;
        clearTimeout(tarea.reloj);
        enCurso.delete(w);
        libres.push(w);
        pasos.set(w, pasos.get(w) + 1);
        const d = e.data || {};
        if (d.tipo === 'hecho') tarea.ok(d);
        else tarea.mal(Object.assign(new Error(d.mensaje || 'Error en el hilo'), { name: d.nombre || 'Error', enPreparar: !!d.enPreparar }));
        despachar();
      };
      w.onerror = (e) => { if (e.preventDefault) e.preventDefault(); caido(w); };
      w.onmessageerror = () => caido(w);
    }
    return {
      tamano: trabajadores.length,
      sirve: (m) => !!m.hilo && !sinHilo.has(m.id) && trabajadores.length > 0,
      sinHilo,
      ejecutar(m, ficha, opciones, contexto) {
        return new Promise((ok, mal) => {
          cola.push({ ok, mal, mensaje: { tipo: 'paso', id: siguiente++, modulo: m.id, ficha, opciones, contexto } });
          despachar();
        });
      },
      pasosPorHilo: () => [...pasos.values()],
      cerrar() { for (const w of pasos.keys()) w.terminate(); },
    };
  }

  /**
   * Corre el flujo. archivosPorEntrada: un arreglo de archivos por entrada (con una sola entrada vale el arreglo
   * de archivos directo). Un tramo de pasos 'uno'/'abrir' lleva cada archivo (y lo que se abra de él) hasta el
   * próximo 'juntar', que espera a todos. Al final de las entradas, lo que salió de todas pasa junto al tramo común.
   * Varios archivos avanzan a la vez (ver «hilos»), pero lo que sale queda en el orden de los archivos, como si
   * hubieran ido de a uno. alAvanzar({ hechos, total, archivo, entrada, paso }) avisa el avance (el total es
   * aproximado si un paso abre un archivo en varios). Un archivo que falla no frena a los demás.
   * Devuelve { resultados, fallos, hilos: { cuantos, pasosPorHilo, enPagina } } (hilos: para medir y probar).
   * opciones.hilos: cuántos hilos abrir (0 = todo en la página); por defecto, según los núcleos del equipo.
   */
  async function ejecutar(flujo, archivos, { alAvanzar, hilos } = {}) {
    flujo = normalizar(flujo);
    const error = validar(flujo);
    if (error) throw Object.assign(new Error('Flujo inválido'), { codigo: error.motivo, entrada: error.entrada, paso: error.paso });
    const fallos = [];
    // Una lista de archivos (lo normal: la única zona de la página) se reparte por tipo; también se acepta ya
    // repartida (un arreglo por entrada).
    let archivosPorEntrada;
    if (archivos.length && Array.isArray(archivos[0])) archivosPorEntrada = archivos;
    else {
      const r = repartir(flujo, archivos);
      archivosPorEntrada = r.porEntrada;
      r.sinEntrada.forEach((a) => fallos.push({ nombre: a.name, entrada: null, codigo: 'tipo' }));
    }
    const usados = [...new Set([...flujo.entradas.flat(), ...flujo.comun].map((p) => modulo(p.id)))];
    const cuantos = hilos == null ? cuantosHilos() : Math.max(0, Math.floor(hilos));
    let grupo = null;
    try {
      // No más hilos que archivos, salvo que un paso abra cada archivo en varios (las páginas de un PDF).
      const archivosTotal = archivosPorEntrada.reduce((s, l) => s + ((l && l.length) || 0), 0);
      const necesarios = usados.some((m) => m.forma === 'abrir') ? cuantos : Math.min(cuantos, archivosTotal);
      if (necesarios && usados.some((m) => m.hilo)) grupo = await abrirHilos(necesarios);
      // Lo que va a la página se prepara ya (si falla, falla todo, como siempre). Lo que va a los hilos se prepara
      // allá; en la página solo si algún paso tiene que volver (ver ejecutarPaso).
      for (const m of usados) if (!grupo || !m.hilo) await prepararEnPagina(m);
      return await correrTodo(flujo, archivosPorEntrada, fallos, alAvanzar, grupo);
    } finally {
      if (grupo) grupo.cerrar();
      for (const m of usados) if (m.terminar) { try { await m.terminar(); } catch (e) { /* ya terminó */ } }
    }
  }

  const preparadosEnPagina = new Map();
  function prepararEnPagina(m) {
    if (!m.preparar) return Promise.resolve();
    if (!preparadosEnPagina.has(m.id)) {
      const p = Promise.resolve().then(() => m.preparar());
      p.catch(() => preparadosEnPagina.delete(m.id));   // el próximo intento vuelve a probar
      preparadosEnPagina.set(m.id, p);
    }
    return preparadosEnPagina.get(m.id);
  }

  async function correrTodo(flujo, archivosPorEntrada, fallos, alAvanzar, grupo) {
    // Los de la página: uno solo a la vez entre todos (candado), salvo los que declaran `simultaneos`.
    const candado = semaforo(1);
    const propios = new Map();
    const limiteDe = (m) => {
      if (!m.simultaneos) return candado;
      if (!propios.has(m.id)) propios.set(m.id, semaforo(Math.max(1, typeof m.simultaneos === 'function' ? m.simultaneos() : m.simultaneos)));
      return propios.get(m.id);
    };
    let enPagina = 0;
    const seguidas = new Map();   // fallas de entorno seguidas, por módulo
    async function ejecutarPaso(m, ficha, opciones, contexto) {
      if (grupo && grupo.sirve(m)) {
        try {
          const d = await grupo.ejecutar(m, ficha, opciones, contexto);
          seguidas.set(m.id, 0);
          return d.vacio ? undefined : d.r;
        } catch (e) {
          if (e.enPreparar) grupo.sinHilo.add(m.id);
          else if (!ERRORES_DE_ENTORNO.includes(e.name)) { seguidas.set(m.id, 0); throw e; }   // el archivo está mal: en la página fallaría igual
          else {
            const n = (seguidas.get(m.id) || 0) + 1;
            seguidas.set(m.id, n);
            if (n >= FALLAS_SEGUIDAS_MAX) grupo.sinHilo.add(m.id);   // no anda en el hilo: el resto va directo a la página
          }
          // Se repite en la página: mismo código, mismo resultado.
        }
      }
      return limiteDe(m)(async () => {
        await prepararEnPagina(m);
        enPagina++;
        return m.ejecutar(ficha, opciones, contexto);
      });
    }
    // Cuántos archivos avanzan a la vez: lo justo para que ningún hilo espere, sin tener el lote entero en memoria.
    const usados = [...new Set([...flujo.entradas.flat(), ...flujo.comun].map((p) => modulo(p.id)))];
    const enParalelo = grupo ? grupo.tamano + 2
      : Math.max(1, ...usados.map((m) => (m.simultaneos ? (typeof m.simultaneos === 'function' ? m.simultaneos() : m.simultaneos) : 1)));
    const admitir = semaforo(enParalelo);
    // Un tramo que abre cada archivo en varios (PDF a imágenes, dividir) ya reparte sus partes entre los hilos: de a
    // dos archivos, para no tener las páginas de muchos PDF a la vez en memoria.
    const admitirAbriendo = grupo ? semaforo(2) : admitir;

    const porEntrada = flujo.entradas.map((pasos, i) => {
      const aceptados = tiposDeEntrada(flujo, i);
      const fichas = [];
      for (const a of archivosPorEntrada[i] || []) {
        const tipo = tipoDeArchivo(a);
        if (!tipo || !aceptados.includes(tipo)) { fallos.push({ nombre: a.name, entrada: i, codigo: 'tipo' }); continue; }
        fichas.push({ nombre: a.name, blob: a, tipo, datos: {}, origen: a.name });
      }
      return fichas;
    });

    let hechos = 0;
    const total = flujo.entradas.reduce((s, pasos, i) => s + porEntrada[i].length * pasos.length, 0)
      + porEntrada.reduce((s, f) => s + f.length, 0) * flujo.comun.length;
    const avisar = (archivo, entrada, paso) => alAvanzar && alAvanzar({ hechos, total: Math.max(total, hechos), archivo, entrada, paso });

    // Los fallos se anotan con el orden del archivo, para listarlos al final como si todo hubiera ido de a uno.
    const anotar = (orden, fallo) => fallos.push(Object.assign(fallo, { _orden: orden }));

    // Lleva una ficha por los pasos [k, fin) que son 'uno' o 'abrir'; devuelve lo que sale (0, 1 o varias), en orden.
    async function recorrer(ficha, pasos, k, fin, entrada, orden) {
      if (k >= fin) return [ficha];
      avisar(ficha.origen, entrada, k);
      let r, fallo = false;
      try {
        r = await ejecutarPaso(modulo(pasos[k].id), ficha, pasos[k].opciones, { paso: k, entrada });
      } catch (e) {
        anotar(orden, { nombre: ficha.origen, entrada, paso: k, error: e.message || String(e) });
        r = null;
        fallo = true;
      }
      hechos++;
      await new Promise((ok) => setTimeout(ok, 0));   // deja respirar a la página (barra de avance)
      const lista = Array.isArray(r) ? r : r ? [r] : [];
      if (!lista.length && r !== undefined && !fallo) anotar(orden, { nombre: ficha.origen, entrada, paso: k, codigo: 'quitado' });
      // Lo que abrió un paso (las páginas de un PDF) también avanza a la vez: un solo PDF grande usa todos los hilos.
      const salen = await Promise.all(lista.map((x) => { x.origen = ficha.origen; x._orden = orden; return recorrer(x, pasos, k + 1, fin, entrada, orden); }));
      return salen.flat();
    }

    async function correr(pasos, fichas, entrada) {
      const inicio = (entrada === 'comun' ? flujo.entradas.length : entrada) * 1e7;
      let i = 0;
      while (i < pasos.length) {
        const m = modulo(pasos[i].id);
        if (m.forma === 'juntar') {
          avisar('', entrada, i);
          if (fichas.length) {
            try {
              await prepararEnPagina(m);
              const juntas = fichas;   // también con el candado: las entradas corren a la vez
              const r = await limiteDe(m)(() => m.ejecutar(juntas, pasos[i].opciones, { paso: i, entrada }));
              fichas = r ? [r] : [];
            } catch (e) {
              anotar(inicio + 1e7 - 1, { nombre: fichas.map((f) => f.origen).join(', '), entrada, paso: i, error: e.message || String(e) });
              fichas = [];
            }
          }
          hechos++;
          i++;
          continue;
        }
        let j = i;
        while (j < pasos.length && modulo(pasos[j].id).forma !== 'juntar') j++;
        const desde = i;
        const entrar = pasos.slice(desde, j).some((p) => modulo(p.id).forma === 'abrir') ? admitirAbriendo : admitir;
        // El orden de un fallo es el del archivo del que salió (viaja en _orden de etapa en etapa).
        const siguen = await Promise.all(fichas.map((f, n) => entrar(() => recorrer(f, pasos, desde, j, entrada, f._orden ?? inicio + n))));
        fichas = siguen.flat();
        i = j;
      }
      return fichas;
    }

    const salidas = await Promise.all(flujo.entradas.map((pasos, i) => correr(pasos, porEntrada[i], i)));
    let fichas = salidas.flat();
    if (flujo.comun.length) fichas = await correr(flujo.comun, fichas, 'comun');

    const nombres = new Set();
    const resultados = fichas.map((f) => {
      const nombre = f.tipo === 'pdf' ? base(f.nombre) + '.pdf' : f.tipo === 'texto' ? base(f.nombre) + '.txt' : f.nombre;
      return { nombre: nombreUnico(nombre, nombres), blob: f.blob, tipo: f.tipo, datos: f.datos, origen: f.origen };
    });
    // Orden estable: los de repartir (sin _orden) primero, después por archivo.
    const ordenados = fallos.map((f, i) => [f, i]).sort((a, b) => ((a[0]._orden ?? -1) - (b[0]._orden ?? -1)) || a[1] - b[1]).map(([f]) => f);
    ordenados.forEach((f) => delete f._orden);
    fallos.splice(0, fallos.length, ...ordenados);
    return {
      resultados, fallos,
      hilos: { cuantos: grupo ? grupo.pasosPorHilo().length : 0, pasosPorHilo: grupo ? grupo.pasosPorHilo() : [], enPagina },
    };
  }

  // ---------- guardar ----------
  function leerGuardadas() {
    try {
      const d = JSON.parse(localStorage.getItem(CLAVE) || '[]');
      return Array.isArray(d) ? d.filter((x) => x && typeof x.nombre === 'string') : [];
    } catch (e) { return []; }
  }
  function escribirGuardadas(lista) {
    try { localStorage.setItem(CLAVE, JSON.stringify(lista)); return true; } catch (e) { return false; }
  }
  // Las guardadas antes de 2026-11-10 tienen `cadena` (una lista); las nuevas, `flujo`.
  const guardadas = () => leerGuardadas().map((x) => ({ nombre: x.nombre, flujo: normalizar(x.flujo || x.cadena) }));
  function guardar(nombre, flujo) {
    nombre = String(nombre).trim().slice(0, 60);
    if (!nombre) return false;
    const lista = leerGuardadas().filter((x) => x.nombre !== nombre);
    lista.push({ nombre, flujo: normalizar(flujo) });
    return escribirGuardadas(lista);
  }
  function borrar(nombre) { return escribirGuardadas(leerGuardadas().filter((x) => x.nombre !== nombre)); }

  function exportar(nombre, flujo) {
    const f = normalizar(flujo);
    const datos = { formato: FORMATO, version: 2, nombre: String(nombre || '').slice(0, 60), entradas: f.entradas, comun: f.comun };
    return new Blob([JSON.stringify(datos, null, 2)], { type: 'application/json' });
  }

  // El archivo viene de afuera: solo se toman ids de módulos que existen y opciones dentro de lo permitido.
  // Versión 1 (una lista en `pasos`) y versión 2 (`entradas` + `comun`).
  function importar(texto) {
    let d;
    try { d = JSON.parse(texto); } catch (e) { throw Object.assign(new Error('JSON inválido'), { codigo: 'json' }); }
    if (!d || d.formato !== FORMATO) throw Object.assign(new Error('No es un flujo'), { codigo: 'formato' });
    let crudo;
    if (Array.isArray(d.pasos)) crudo = { entradas: [d.pasos], comun: [] };
    else if (Array.isArray(d.entradas)) crudo = { entradas: d.entradas.filter(Array.isArray), comun: Array.isArray(d.comun) ? d.comun : [] };
    else throw Object.assign(new Error('No es un flujo'), { codigo: 'formato' });
    const flujo = normalizar(crudo);
    const cuantos = (f) => f.entradas.reduce((s, p) => s + (Array.isArray(p) ? p.length : 0), 0) + (f.comun || []).length;
    if (!cuantos(flujo)) throw Object.assign(new Error('Sin pasos conocidos'), { codigo: 'vacio' });
    return { nombre: typeof d.nombre === 'string' ? d.nombre.slice(0, 60) : '', flujo, omitidos: cuantos(crudo) - cuantos(flujo) };
  }

  // ---------- utilidades para los módulos ----------
  const cargados = new Map();
  function cargarScript(src) {
    // Ya lo puso la página u otro cargador (archivo-a-pdf.js trae pdf-texto.js, zip-propio.js…): no repetirlo, que
    // un script cargado dos veces falla al volver a declarar sus constantes. Se compara la dirección completa.
    if (!cargados.has(src) && [...document.scripts].some((s) => s.src === src)) cargados.set(src, Promise.resolve());
    if (!cargados.has(src)) cargados.set(src, new Promise((ok, mal) => {
      const s = document.createElement('script');
      s.src = src; s.onload = ok;
      s.onerror = () => { cargados.delete(src); mal(new Error('No cargó ' + src)); };
      document.head.appendChild(s);
    }));
    return cargados.get(src);
  }
  const BASE_JS = new URL('./', document.currentScript.src).href;
  const rutaJs = (archivo) => BASE_JS + archivo;

  window.Flujos = {
    registrar, modulos, modulo, compatibles, tiposDeEntrada, tiposAlFinal, tiposTrasPasos, tiposAlComun, repartir, salidaDe, validar, ejecutar,
    tipoDeArchivo, normalizar, guardadas, guardar, borrar, exportar, importar, cargarScript, rutaJs, base,
  };
})();

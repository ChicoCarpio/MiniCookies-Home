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
//   forma:'uno'|'abrir'|'juntar', opciones:[{ id, tipo:'select'|'numero'|'casilla'|'texto', etiqueta:{es,en},
//   valores:[{ valor, es, en }], defecto, min, max, paso }], preparar() -> Promise (carga lo que necesita),
//   ejecutar(ficha | fichas, opciones, contexto) -> Promise<ficha | fichas | null>,
//   terminar() -> Promise (opcional: libera lo que preparar o ejecutar dejaron abierto, p. ej. el motor de OCR) }.
//   'uno' recibe una ficha y devuelve una (o null para quitarla); 'abrir' recibe una y devuelve varias (un PDF ->
//   sus páginas); 'juntar' recibe todas las que llegan y devuelve una.
// Todo corre en este navegador, archivo por archivo (un lote grande no ocupa la memoria de todos juntos).
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

  /**
   * Corre el flujo. archivosPorEntrada: un arreglo de archivos por entrada (con una sola entrada vale el arreglo
   * de archivos directo). Cada entrada procesa sus archivos de a uno: un tramo de pasos 'uno'/'abrir' lleva cada
   * archivo (y lo que se abra de él) hasta el próximo 'juntar', que espera a todos. Al final de las entradas, lo que
   * salió de todas pasa junto al tramo común. alAvanzar({ hechos, total, archivo, entrada, paso }) avisa el avance
   * (el total es aproximado si un paso abre un archivo en varios). Un archivo que falla no frena a los demás.
   */
  async function ejecutar(flujo, archivos, { alAvanzar } = {}) {
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
    try {
      for (const m of usados) if (m.preparar) await m.preparar();
      return await correrTodo(flujo, archivosPorEntrada, fallos, alAvanzar);
    } finally {
      for (const m of usados) if (m.terminar) { try { await m.terminar(); } catch (e) { /* ya terminó */ } }
    }
  }

  async function correrTodo(flujo, archivosPorEntrada, fallos, alAvanzar) {

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

    // Lleva una ficha por los pasos [k, fin) que son 'uno' o 'abrir'; devuelve lo que sale (0, 1 o varias).
    async function recorrer(ficha, pasos, k, fin, entrada) {
      if (k >= fin) return [ficha];
      avisar(ficha.origen, entrada, k);
      let r;
      try {
        r = await modulo(pasos[k].id).ejecutar(ficha, pasos[k].opciones, { paso: k, entrada });
      } catch (e) {
        fallos.push({ nombre: ficha.origen, entrada, paso: k, error: e.message || String(e) });
        r = null;
      }
      hechos++;
      await new Promise((ok) => setTimeout(ok, 0));   // deja respirar a la página (barra de avance)
      const lista = Array.isArray(r) ? r : r ? [r] : [];
      if (!lista.length && r !== undefined && !fallos.some((f) => f.nombre === ficha.origen && f.paso === k && f.entrada === entrada)) {
        fallos.push({ nombre: ficha.origen, entrada, paso: k, codigo: 'quitado' });
      }
      const salen = [];
      for (const x of lista) { x.origen = ficha.origen; salen.push(...await recorrer(x, pasos, k + 1, fin, entrada)); }
      return salen;
    }

    async function correr(pasos, fichas, entrada) {
      let i = 0;
      while (i < pasos.length) {
        const m = modulo(pasos[i].id);
        if (m.forma === 'juntar') {
          avisar('', entrada, i);
          if (fichas.length) {
            try {
              const r = await m.ejecutar(fichas, pasos[i].opciones, { paso: i, entrada });
              fichas = r ? [r] : [];
            } catch (e) {
              fallos.push({ nombre: fichas.map((f) => f.origen).join(', '), entrada, paso: i, error: e.message || String(e) });
              fichas = [];
            }
          }
          hechos++;
          i++;
          continue;
        }
        let j = i;
        while (j < pasos.length && modulo(pasos[j].id).forma !== 'juntar') j++;
        const siguen = [];
        for (const f of fichas) siguen.push(...await recorrer(f, pasos, i, j, entrada));
        fichas = siguen;
        i = j;
      }
      return fichas;
    }

    let fichas = [];
    for (let i = 0; i < flujo.entradas.length; i++) fichas.push(...await correr(flujo.entradas[i], porEntrada[i], i));
    if (flujo.comun.length) fichas = await correr(flujo.comun, fichas, 'comun');

    const nombres = new Set();
    const resultados = fichas.map((f) => {
      const nombre = f.tipo === 'pdf' ? base(f.nombre) + '.pdf' : f.tipo === 'texto' ? base(f.nombre) + '.txt' : f.nombre;
      return { nombre: nombreUnico(nombre, nombres), blob: f.blob, tipo: f.tipo, datos: f.datos, origen: f.origen };
    });
    return { resultados, fallos };
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

/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Motor de Flujos: encadenar herramientas del sitio. No dibuja nada: la página arma la interfaz.
//
//   Flujos.registrar(modulo)                       da de alta un módulo (ver flujos-modulos.js)
//   Flujos.modulos() / Flujos.modulo(id)           los módulos registrados
//   Flujos.compatibles(tipo)                       módulos que reciben ese tipo ('imagen', 'pdf'); sin tipo: todos
//   Flujos.salidaDe(cadena)                        tipo que entrega la cadena (null si está vacía)
//   Flujos.entradas(cadena)                        tipos de archivo que pueden recorrer la cadena entera
//   Flujos.validar(cadena)                       null si se puede correr, o { paso, motivo } ('vacia'|'modulo'|'tipo')
//   Flujos.ejecutar(cadena, archivos, { alAvanzar }) -> Promise<{ resultados, fallos }>
//   Flujos.tipoDeArchivo(archivo)                  'imagen' | 'pdf' | null
//   Flujos.normalizar(cadena)                      cadena limpia: módulos que existen y opciones válidas
//   Flujos.guardadas() / guardar(nombre, cadena) / borrar(nombre)       en este navegador (localStorage)
//   Flujos.exportar(nombre, cadena) -> Blob         archivo .json para guardar o pasar a otro equipo
//   Flujos.importar(texto) -> { nombre, cadena }    lanza un Error con `codigo` si el archivo no sirve
//
// Una cadena es [{ id, opciones }]. Lo que viaja de un paso al siguiente es una ficha de archivo:
//   { nombre, blob, tipo: 'imagen'|'pdf', datos: { codigos: [...], ... }, origen }
// `datos` se conserva de paso en paso: así el QR que se leyó en el paso 1 llega al PDF del paso 3.
//
// Un módulo: { id, icono, nombre:{es,en}, descripcion:{es,en}, entra:[tipos], sale:'imagen'|'pdf'|'igual',
//   forma:'uno'|'juntar', opciones:[{ id, tipo:'select'|'numero'|'casilla', etiqueta:{es,en},
//   valores:[{ valor, es, en }], defecto, min, max, paso }], preparar() -> Promise (carga lo que necesita),
//   ejecutar(ficha | fichas, opciones, contexto) -> Promise<ficha | fichas | null> }.
//   'uno' recibe una ficha y devuelve una (o null para quitarla del flujo); 'juntar' recibe todas y devuelve una.
// Todo corre en este navegador, archivo por archivo (un lote grande no ocupa la memoria de todos juntos).
(function () {
  const registro = new Map();
  const CLAVE = 'flujos-guardados';
  const FORMATO = 'minicookies-flujo';

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
    return null;
  }

  // Tipo que sale de un módulo cuando le entra `tipo` ('igual' = el mismo que entró).
  const salida = (m, tipo) => (m.sale === 'igual' ? tipo : m.sale);

  function compatibles(tipo) {
    return modulos().filter((m) => !tipo || m.entra.includes(tipo));
  }

  function salidaDe(cadena) {
    let tipo = null;
    for (const paso of cadena) {
      const m = modulo(paso.id);
      if (!m) return null;
      tipo = tipo === null ? (m.sale === 'igual' ? m.entra[0] : m.sale) : salida(m, tipo);
    }
    return tipo;
  }

  // Tipos de archivo que pueden recorrer la cadena entera (no solo el primer paso: «leer códigos» recibe PDF,
  // pero si después viene «recortar documento», un PDF no pasa).
  function entradas(cadena) {
    if (!cadena.length || cadena.some((p) => !modulo(p.id))) return [];
    return modulo(cadena[0].id).entra.filter((tipo) => {
      let t = tipo;
      for (const p of cadena) { const m = modulo(p.id); if (!m.entra.includes(t)) return false; t = salida(m, t); }
      return true;
    });
  }

  // El tipo de entrada de la cadena es el del primer paso; si recibe varios, vale cualquiera de ellos.
  function validar(cadena) {
    if (!cadena.length) return { paso: -1, motivo: 'vacia' };
    let tipos = null;
    for (let i = 0; i < cadena.length; i++) {
      const m = modulo(cadena[i].id);
      if (!m) return { paso: i, motivo: 'modulo' };
      if (tipos === null) tipos = m.entra.slice();
      const aceptados = tipos.filter((t) => m.entra.includes(t));
      if (!aceptados.length) return { paso: i, motivo: 'tipo' };
      tipos = [...new Set(aceptados.map((t) => salida(m, t)))];
    }
    return null;
  }

  // Opciones dentro de lo que el módulo permite; lo que falte o no sirva vuelve a su valor por defecto.
  function opcionesValidas(m, dadas) {
    const o = {};
    for (const op of m.opciones || []) {
      let v = dadas && Object.prototype.hasOwnProperty.call(dadas, op.id) ? dadas[op.id] : op.defecto;
      if (op.tipo === 'select' && !op.valores.some((x) => x.valor === v)) v = op.defecto;
      if (op.tipo === 'casilla') v = v === true || v === 'true';
      if (op.tipo === 'numero') {
        v = Number(v);
        if (!Number.isFinite(v)) v = op.defecto;
        v = Math.min(op.max ?? v, Math.max(op.min ?? v, v));
      }
      o[op.id] = v;
    }
    return o;
  }

  function normalizar(cadena) {
    return (Array.isArray(cadena) ? cadena : [])
      .filter((p) => p && typeof p.id === 'string' && registro.has(p.id))
      .slice(0, 20)
      .map((p) => ({ id: p.id, opciones: opcionesValidas(registro.get(p.id), p.opciones) }));
  }

  // ---------- ejecutar ----------
  function extension(tipo) { return tipo === 'pdf' ? '.pdf' : ''; }
  function base(nombre) { const i = nombre.lastIndexOf('.'); return i > 0 ? nombre.slice(0, i) : nombre; }

  function nombreUnico(nombre, usados) {
    let n = nombre, k = 2;
    while (usados.has(n)) { const i = nombre.lastIndexOf('.'); n = i > 0 ? `${nombre.slice(0, i)}-${k}${nombre.slice(i)}` : `${nombre}-${k}`; k++; }
    usados.add(n);
    return n;
  }

  /**
   * Corre la cadena sobre los archivos. Los tramos de pasos 'uno' se hacen archivo por archivo (cada uno sigue
   * su camino hasta el próximo 'juntar'); un 'juntar' espera a todos los que llegaron sanos.
   * alAvanzar({ hechos, total, archivo, paso }) avisa el progreso. Un archivo que falla no frena a los demás.
   */
  async function ejecutar(cadena, archivos, { alAvanzar } = {}) {
    cadena = normalizar(cadena);
    const error = validar(cadena);
    if (error) throw Object.assign(new Error('Cadena inválida'), { codigo: error.motivo, paso: error.paso });
    for (const p of cadena) { const m = modulo(p.id); if (m.preparar) await m.preparar(); }

    let fichas = [];
    const fallos = [];
    const aceptados = entradas(cadena);
    for (const a of archivos) {
      const tipo = tipoDeArchivo(a);
      if (!tipo || !aceptados.includes(tipo)) { fallos.push({ nombre: a.name, codigo: 'tipo' }); continue; }
      fichas.push({ nombre: a.name, blob: a, tipo, datos: {}, origen: a.name });
    }

    // Tramos: [i, j) de pasos 'uno' seguidos, o un paso 'juntar' solo.
    const total = fichas.length * cadena.length;
    let hechos = 0;
    const avisar = (ficha, paso) => alAvanzar && alAvanzar({ hechos, total, archivo: ficha ? ficha.origen : '', paso });
    let i = 0;
    while (i < cadena.length) {
      const m = modulo(cadena[i].id);
      if (m.forma === 'juntar') {
        avisar(null, i);
        if (fichas.length) {
          try {
            const r = await m.ejecutar(fichas, cadena[i].opciones, { paso: i });
            fichas = r ? [r] : [];
          } catch (e) {
            fallos.push({ nombre: fichas.map((f) => f.origen).join(', '), paso: i, error: e.message || String(e) });
            fichas = [];
          }
        }
        hechos += fichas.length;
        i++;
        continue;
      }
      let j = i;
      while (j < cadena.length && modulo(cadena[j].id).forma !== 'juntar') j++;
      const siguen = [];
      for (const ficha of fichas) {
        let actual = ficha;
        for (let k = i; k < j && actual; k++) {
          avisar(actual, k);
          try {
            const r = await modulo(cadena[k].id).ejecutar(actual, cadena[k].opciones, { paso: k });
            if (r) r.origen = ficha.origen;
            if (!r) fallos.push({ nombre: ficha.origen, paso: k, codigo: 'quitado' });
            actual = r;
          } catch (e) {
            fallos.push({ nombre: ficha.origen, paso: k, error: e.message || String(e) });
            actual = null;
          }
          hechos++;
          // Deja respirar a la página entre archivo y archivo (barra de avance, botón de cancelar).
          await new Promise((ok) => setTimeout(ok, 0));
        }
        if (actual) siguen.push(actual);
      }
      fichas = siguen;
      i = j;
    }

    const usados = new Set();
    const resultados = fichas.map((f) => {
      const nombre = f.tipo === 'pdf' ? base(f.nombre) + '.pdf' : f.nombre;
      return { nombre: nombreUnico(nombre + (nombre.includes('.') ? '' : extension(f.tipo)), usados), blob: f.blob, tipo: f.tipo, datos: f.datos, origen: f.origen };
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
  const guardadas = () => leerGuardadas().map((x) => ({ nombre: x.nombre, cadena: normalizar(x.cadena) }));
  function guardar(nombre, cadena) {
    nombre = String(nombre).trim().slice(0, 60);
    if (!nombre) return false;
    const lista = leerGuardadas().filter((x) => x.nombre !== nombre);
    lista.push({ nombre, cadena: normalizar(cadena) });
    return escribirGuardadas(lista);
  }
  function borrar(nombre) { return escribirGuardadas(leerGuardadas().filter((x) => x.nombre !== nombre)); }

  function exportar(nombre, cadena) {
    const datos = { formato: FORMATO, version: 1, nombre: String(nombre || '').slice(0, 60), pasos: normalizar(cadena) };
    return new Blob([JSON.stringify(datos, null, 2)], { type: 'application/json' });
  }

  // El archivo viene de afuera: solo se toman ids de módulos que existen y opciones dentro de lo permitido.
  function importar(texto) {
    let d;
    try { d = JSON.parse(texto); } catch (e) { throw Object.assign(new Error('JSON inválido'), { codigo: 'json' }); }
    if (!d || d.formato !== FORMATO || !Array.isArray(d.pasos)) throw Object.assign(new Error('No es un flujo'), { codigo: 'formato' });
    const cadena = normalizar(d.pasos);
    if (!cadena.length) throw Object.assign(new Error('Sin pasos conocidos'), { codigo: 'vacio' });
    return { nombre: typeof d.nombre === 'string' ? d.nombre.slice(0, 60) : '', cadena, omitidos: d.pasos.length - cadena.length };
  }

  // ---------- utilidades para los módulos ----------
  const cargados = new Map();
  function cargarScript(src) {
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
    registrar, modulos, modulo, compatibles, salidaDe, entradas, validar, ejecutar, tipoDeArchivo, normalizar,
    guardadas, guardar, borrar, exportar, importar, cargarScript, rutaJs, base,
  };
})();

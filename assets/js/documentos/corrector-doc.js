/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Corrector ortográfico de Documentos (Fase 6, etapa 6.8). Sin DOM: corre igual en Node. Usa nspell (MIT) con diccionarios Hunspell: español (elegido
// bajo MPL 1.1) e inglés (MIT y BSD), todo local; nada sale del equipo.
//
//   const c = CorrectorDoc.crear({ NSpell, leer: async (idioma) => ({ aff, dic }) });
//   await c.cargar('es')                          carga el diccionario (una vez)
//   c.errores('Una palabra mal escrita', 'es')    → [{ o, n, palabra }] (desplazamiento y largo de cada palabra que no está)
//   c.sugerencias('palabrra', 'es')               → ['palabra', …]
//   c.agregar('Pelicano') · c.ignorar('xyz')      al diccionario personal (c.personales()) · solo mientras dure la sesión
//
// Se saltan: palabras de una letra, con números, guiones bajos, arrobas o puntos y barras (direcciones), todo en mayúsculas (siglas) y las que ya se
// agregaron o ignoraron. Una palabra con guion se da por buena si lo son todas sus partes.
(function () {
  'use strict';
  const PALABRA = /[\p{L}\p{M}\p{N}_]+(?:['’.\-@/][\p{L}\p{M}\p{N}_]+)*/gu;

  const WEB = /(?:\b(?:https?|ftp):\/\/|\bwww\.)\S+/g;
  // un correo es un trozo sin espacios con una arroba (no al principio) y después un punto que no es el último carácter. Sin expresión regular a
  // propósito: «\S+@\S+\.\S+» tardaba un minuto en un párrafo de 10.000 caracteres como «a@a@a@…» (vuelve atrás en cada arroba)
  const esCorreo = (x) => { const i = x.indexOf('@', 1), j = x.lastIndexOf('.'); return i > 0 && j > i + 1 && j < x.length - 1; };

  function crear(opc) {
    const motores = new Map();      // idioma → NSpell
    const cargando = new Map();
    const personales = new Set(opc.personales || []);
    const ignoradas = new Set();
    const cache = new Map();        // idioma|palabra → bool

    async function cargar(idioma) {
      if (motores.has(idioma)) return motores.get(idioma);
      if (!cargando.has(idioma)) {
        cargando.set(idioma, (async () => {
          const d = await opc.leer(idioma);
          const m = opc.NSpell({ aff: d.aff, dic: d.dic });
          motores.set(idioma, m);
          return m;
        })().catch((e) => { cargando.delete(idioma); throw e; }));
      }
      return cargando.get(idioma);
    }
    const listo = (idioma) => motores.has(idioma);

    function buena(m, idioma, p) {
      const k = idioma + '|' + p;
      if (cache.has(k)) return cache.get(k);
      let ok = m.correct(p);
      if (!ok && p.includes('-')) ok = p.split('-').every((x) => !x || x.length < 2 || m.correct(x));
      if (!ok && /[’]/.test(p)) ok = m.correct(p.replace(/’/g, "'"));
      cache.set(k, ok);
      return ok;
    }
    function errores(texto, idioma) {
      const m = motores.get(idioma);
      if (!m || !texto) return [];
      // direcciones web y de correo: se tapan con espacios (mismo largo) para que no cuenten
      const tapar = (x) => ' '.repeat(x.length);
      texto = texto.replace(WEB, tapar).replace(/\S+/g, (x) => (esCorreo(x) ? tapar(x) : x));
      const sal = [];
      PALABRA.lastIndex = 0;
      let r;
      while ((r = PALABRA.exec(texto))) {
        const p = r[0];
        if (p.length < 2 || /[\d_@/.]/.test(p) || (p === p.toUpperCase() && p !== p.toLowerCase()) || personales.has(p) || ignoradas.has(p)) continue;
        if (/^['’]|['’]$/.test(p)) continue;
        if (!buena(m, idioma, p)) sal.push({ o: r.index, n: p.length, palabra: p });
      }
      return sal;
    }
    function sugerencias(palabra, idioma, max = 6) {
      const m = motores.get(idioma);
      return m ? m.suggest(palabra).slice(0, max) : [];
    }
    function agregar(palabra) { personales.add(palabra); const l = opc.alAgregar; if (l) l([...personales]); }
    function ignorar(palabra) { ignoradas.add(palabra); }
    return { cargar, listo, errores, sugerencias, agregar, ignorar, personales: () => [...personales] };
  }

  globalThis.CorrectorDoc = { crear };
})();

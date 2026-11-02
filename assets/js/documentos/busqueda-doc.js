/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Buscar en el texto de los párrafos de un documento (Fase 6, etapa 6.7). Sin DOM: corre igual en Node.
//
//   BusquedaDoc.buscar(parrafos, 'hola', { mayusculas: false, palabra: false }) → [{ b, o, n }]   (b: el párrafo, o: desplazamiento, n: largo)
//
// Como en Word, una coincidencia no cruza párrafos. Sin distinguir mayúsculas se compara letra por letra con su minúscula, y solo si esa minúscula
// ocupa una unidad (así los desplazamientos son los del texto original). Los saltos suaves (U+2028) se buscan con «\n».
(function () {
  'use strict';
  const SUAVE = String.fromCharCode(0x2028);
  const LETRA = /[\p{L}\p{N}_]/u;

  const plegar = (s) => {
    let sal = '';
    for (let i = 0; i < s.length; i++) { const m = s[i].toLowerCase(); sal += m.length === 1 ? m : s[i]; }
    return sal;
  };

  function buscar(parrafos, consulta, opc = {}) {
    const q0 = String(consulta || '').replace(/\n/g, SUAVE);
    if (!q0) return [];
    const q = opc.mayusculas ? q0 : plegar(q0);
    const sal = [];
    for (const b of parrafos) {
      const t = b.texto;
      if (t.length < q.length) continue;
      const h = opc.mayusculas ? t : plegar(t);
      for (let i = h.indexOf(q); i >= 0; i = h.indexOf(q, i + q.length)) {
        if (opc.palabra) {
          // ¿hay una letra pegada antes o después? (el carácter entero, no la mitad de un par sustituto)
          const ant = i > 0 ? Array.from(t.slice(Math.max(0, i - 2), i)).pop() : '';
          const des = Array.from(t.slice(i + q.length, i + q.length + 2))[0] || '';
          if ((ant && LETRA.test(ant)) || (des && LETRA.test(des))) continue;
        }
        sal.push({ b, o: i, n: q.length });
      }
    }
    return sal;
  }

  globalThis.BusquedaDoc = { buscar };
})();

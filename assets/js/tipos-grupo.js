/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Tipos de grupo del Editor avanzado (plan del rediseño, hallazgos/editores-rediseno.md, 3.3). Un grupo
// no es genérico: tiene un tipo que dice qué puede ir adentro y cómo se
// compone con lo de abajo. Cada tipo se define UNA vez en TIPOS; uno nuevo
// (un "Marco", "Variantes"…) se agrega ahí sin tocar los otros.
//
// Los elementos del árbol, por clase:
//   'ajuste'    capa de ajuste (luz, color, filtro…), sin contenido propio
//   'contenido' capa con texto o una imagen; solo vive como base de un objeto
//   'objeto' | 'ajustes' | 'carpeta'   grupos ({ tipo: 'grupo', clase })
//
// Cómo se compone cada tipo lo hace el editor (componerItem); acá está el
// "qué": si pasa a través, si se recorta a una forma y qué admite.

window.TiposGrupo = (function () {
  'use strict';

  const TIPOS = {
    // Una imagen o un texto con SUS ajustes (el "objeto inteligente" de
    // Photoshop). El contenido va fijo abajo (índice 0) y no se mueve de ahí;
    // los ajustes de arriba trabajan solo sobre él y no se salen de su forma.
    // Después el objeto entero se funde con lo de abajo con su opacidad, su
    // modo y su zona (qué parte del objeto se ve).
    objeto: { icono: '🖼', admite: ['ajuste', 'ajustes'], base: 'contenido', pasaATraves: false },
    // Varios ajustes juntos: afectan a todo lo de abajo (o a todo el objeto,
    // si están adentro de uno), o solo a su zona. Nunca un objeto.
    ajustes: { icono: '☀', admite: ['ajuste', 'ajustes'], pasaATraves: true },
    // Solo ordena: lo de adentro se comporta igual que afuera. Sirve para
    // plegar, ocultar, mover u opacar varias cosas juntas.
    carpeta: { icono: '📁', admite: ['ajuste', 'ajustes', 'objeto', 'carpeta'], pasaATraves: true },
  };

  function esGrupo(item) { return !!item && item.tipo === 'grupo'; }

  function claseDe(item) {
    if (esGrupo(item)) return TIPOS[item.clase] ? item.clase : 'carpeta';
    return item && item.contenido ? 'contenido' : 'ajuste';
  }

  function contiene(grupo, buscado) {
    if (!esGrupo(grupo)) return false;
    return grupo.hijos.some((h) => h === buscado || contiene(h, buscado));
  }

  // Por qué "item" no puede ir en la posición "indice" (0 = abajo) de la
  // lista de "padre" (null = la raíz), como clave de texto; null si puede.
  function motivoNoPuede(item, padre, indice) {
    const clase = claseDe(item);
    if (clase === 'contenido') return 'motivo-contenido-fijo';
    if (padre && (padre === item || contiene(item, padre))) return 'motivo-dentro-de-si';
    if (!padre) return null;
    const tipo = TIPOS[claseDe(padre)];
    if (!tipo.admite.includes(clase)) {
      if (padre.clase === 'objeto') return clase === 'objeto' ? 'motivo-objeto-en-objeto' : 'motivo-objeto-solo-ajustes';
      return 'motivo-ajustes-solo-ajustes';
    }
    if (tipo.base && indice < 1) return 'motivo-debajo-del-contenido';
    return null;
  }

  // ---- Proyectos guardados hasta la versión 6 (sin tipos) ----
  // Se convierten sobre el JSON, antes de armar los lienzos, y dan
  // EXACTAMENTE los mismos píxeles:
  //  - Una capa de texto o imagen pasa a ser un objeto: la capa queda como su
  //    contenido (sin ajustes, sin máscara, al 100 % y en Normal) y sus
  //    ajustes, si tenía, como una capa de ajuste adentro. La máscara, la
  //    opacidad y el modo pasan al objeto. Antes: máscara(recorte(ajustes(
  //    contenido))); ahora el objeto hace lo mismo en el mismo orden.
  //  - Un grupo "Pasar a través" pasa a ser de ajustes, o carpeta si tiene
  //    objetos adentro: ninguno de los dos aísla, como antes.
  //  - Un grupo aislado (versiones 5 y 6) no encaja en ningún tipo sin cambiar
  //    píxeles en los bordes: queda como carpeta con "aisladoV6", que el
  //    editor sigue componiendo como antes.
  // opciones: { version, esNeutro(ajustes), nuevoId() }
  function convertirDesdeV6(items, opciones) {
    const { version, esNeutro, nuevoId } = opciones;
    const aislado = (g) => !(g.modoFusion === 'paso' || (version < 5 && (!g.modoFusion || g.modoFusion === 'normal')));

    function convertir(c) {
      if (c.tipo === 'grupo' && TIPOS[c.clase]) return c; // ya tiene tipo
      if (c.tipo === 'grupo') {
        const hijos = (c.hijos || []).map(convertir);
        const g = { ...c, hijos };
        if (aislado(c)) {
          g.clase = 'carpeta';
          g.aisladoV6 = true;
        } else {
          g.clase = hijos.some((h) => h.tipo === 'grupo' && h.clase !== 'ajustes') ? 'carpeta' : 'ajustes';
          g.modoFusion = 'normal';
        }
        return g;
      }
      if (!c.contenido) return c;
      const contenido = {
        ...c, visible: true, opacidad: 100, modoFusion: 'normal',
        mascara: null, pincelSumar: null, pincelQuitar: null, ajustes: {},
      };
      const hijos = [contenido];
      if (c.ajustes && !esNeutro(c.ajustes)) {
        hijos.push({
          id: nuevoId(), nombre: c.nombre, visible: true, opacidad: 100, modoFusion: 'normal',
          mascara: null, pincelSumar: null, pincelQuitar: null, ajustes: c.ajustes,
        });
      }
      return {
        id: nuevoId(), tipo: 'grupo', clase: 'objeto', nombre: c.nombre, nombreAuto: c.nombreAuto || null,
        visible: c.visible !== false, opacidad: c.opacidad, modoFusion: c.modoFusion,
        mascara: c.mascara, pincelSumar: c.pincelSumar, pincelQuitar: c.pincelQuitar,
        plegado: hijos.length === 1, hijos,
      };
    }
    return items.map(convertir);
  }

  return { TIPOS, claseDe, motivoNoPuede, convertirDesdeV6 };
})();

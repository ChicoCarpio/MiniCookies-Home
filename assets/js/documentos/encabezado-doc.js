/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Encabezados y pies de página de Documentos (Fase 6, etapa 6.7): buscar el de una sección (heredado de las anteriores, como en Word), crear uno que no
// existe y abrirlo como un documento aparte que se edita con EdicionDoc y se muestra con MaquetadorDoc.
//
//   EncabezadoDoc.resolver(modelo, secIdx, 'encabezado'|'pie', 'default'|'first'|'even') → { ruta, seccion } | null
//   await EncabezadoDoc.crear(modelo, tipo, cual) → ruta            solo en la última sección; la referencia se escribe al guardar
//   await EncabezadoDoc.abrir(modelo, secIdx, tipo, cual, { crear }) → { vista, ruta, creada } | null
//   await EncabezadoDoc.guardar(vista, ed)                           escribe en el paquete el XML de la parte editada
//
// Requiere docx-modelo.js (modelo.editarParte) y paquete-editable.js.
(function () {
  'use strict';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

  // Una sección sin referencia propia usa la de la anterior; el «even» sin referencia usa el «default»
  function resolver(modelo, secIdx, tipo, cual) {
    const buscar = (c) => {
      for (let i = Math.min(secIdx, modelo.secciones.length - 1); i >= 0; i--) {
        const r = (modelo.secciones[i].referencias || []).find((x) => x.tipo === tipo && x.cual === c);
        const rel = r && modelo.rels.get(r.id);
        if (rel && !rel.externo) return { ruta: rel.destino, seccion: i };
      }
      return null;
    };
    return buscar(cual) || (cual === 'even' ? buscar('default') : null);
  }

  async function crear(modelo, tipo, cual) {
    if (modelo.crearParte) return modelo.crearParte(tipo, cual);   // .odt: el encabezado vive en la página maestra de styles.xml
    const nombre = tipo === 'encabezado' ? 'header' : 'footer', raiz = tipo === 'encabezado' ? 'hdr' : 'ftr';
    const paquete = modelo.paquete;
    let n = 1;
    while (paquete.existe(`word/${nombre}${n}.xml`)) n++;
    const ruta = `word/${nombre}${n}.xml`;
    paquete.escribir(ruta, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${raiz} xmlns:w="${W}" xmlns:r="${REL}"><w:p/></w:${raiz}>`);
    await paquete.asegurarTipo('/' + ruta, `application/vnd.openxmlformats-officedocument.wordprocessingml.${nombre}+xml`);
    const rid = await paquete.agregarRelacion(modelo.ruta, `${REL}/${nombre}`, `${nombre}${n}.xml`);
    modelo.rels.set(rid, { tipo: nombre, destino: ruta, externo: false });
    modelo.seccionFinal.referencias.push({ tipo, cual, id: rid });
    modelo.recalcularSecciones();
    (modelo.referenciasNuevas = modelo.referenciasNuevas || []).push({ tipo, cual, id: rid });
    return ruta;
  }

  // Abre el encabezado o pie de la sección. Si no existe y `crear` es verdadero, lo crea (solo se puede en la última sección o en un documento de
  // una sola sección, porque la referencia nueva se escribe en la última). Devuelve null si no hay.
  async function abrir(modelo, secIdx, tipo, cual, opc = {}) {
    let r = resolver(modelo, secIdx, tipo, cual), creada = false;
    if (!r) {
      if (!opc.crear || secIdx !== modelo.secciones.length - 1) return null;
      const ruta = await crear(modelo, tipo, cual);
      r = { ruta, seccion: secIdx };
      creada = true;
    }
    const vista = await modelo.editarParte(r.ruta, secIdx);
    return vista ? { vista, ruta: r.ruta, creada, heredada: r.seccion !== secIdx } : null;
  }

  async function guardar(vista, ed) {
    vista.paquete.escribir(vista.ruta, await ed.xmlGuardado());
  }

  globalThis.EncabezadoDoc = { resolver, crear, abrir, guardar };
})();

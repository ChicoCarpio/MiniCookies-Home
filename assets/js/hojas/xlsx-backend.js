/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Backend de .xlsx / .xlsm del libro editable (F3.4a, etapa 1 del plan de ODS). Es lo que LibroEditable necesita del formato: abrir el archivo en un
// modelo (LibroXlsx), decir qué celdas no se pueden editar, crear los estilos nuevos y escribir SOLO las celdas editadas sin perder nada. Lo neutro
// (celdas, operaciones, deshacer, motor de cálculo) se queda en libro-editable.js; un .ods tendrá su propio backend con la misma forma.
//
//   await XlsxBackend.reconoce(bytes) → bool        XlsxBackend.formato → 'xlsx'
//   await XlsxBackend.abrir(bytes) → modelo = {
//     libro, paquete, hojas, estilos, externos,        (LibroXlsx, PaqueteEditable, un modelo de hoja por hoja, EstilosLibro, vínculos a otros libros)
//     protegida(h, f, c) → bool,                        (encabezado de tabla o fórmula matricial: se rechaza la edición)
//     async guardarCeldas(porHoja) → Uint8Array,       (porHoja = [[h, [{ f, c, valor, estilo }]]]; escribe y devuelve el archivo)
//     destino() → { extension, mime }                  (cómo se descarga)
//   }
// Requiere zip-propio, xml-editar, xlsx-modelo (del sitio) y, de la suite, paquete-editable, motores/xlsx-escritor, hojas/calculo y hojas/estilos-libro.
(function () {
  const MAX_COLS = 16384;
  const clave = (f, c) => f * MAX_COLS + c;

  // Vínculos a otros libros (xl/externalLinks/externalLinkN.xml), en el orden de <externalReferences> del libro: los nombres de sus hojas y
  // las celdas guardadas → [{ hojas: [nombre], datos: Map(indiceHoja → { celdas: Map(f * MAX_COLS + c → valor), errorAlActualizar }) }]
  async function leerVinculosExternos(paquete) {
    const raizRels = await paquete.relaciones('');
    const principal = raizRels.find((r) => /\/officeDocument$/.test(r.tipo));
    if (!principal) return [];
    const xmlLibro = await paquete.leer(principal.destino);
    if (!xmlLibro || xmlLibro.indexOf('externalReference') < 0) return [];
    const raiz = XmlEditar.analizar(xmlLibro, { profundidad: 3 });
    const refs = XmlEditar.hijo(raiz, 'externalReferences');
    if (!refs) return [];
    const rels = await paquete.relaciones(principal.destino);
    const lista = [];
    for (const r of XmlEditar.hijos(refs, 'externalReference')) {
      const id = Object.keys(r.attrs).find((k) => /(^|:)id$/.test(k));
      const rel = id && rels.find((x) => x.id === r.attrs[id]);
      const xml = rel && !rel.externa ? await paquete.leer(rel.destino) : null;
      const vinculo = { hojas: [], datos: new Map() };
      lista.push(vinculo);
      if (!xml) continue;
      const libroExt = XmlEditar.hijo(XmlEditar.analizar(xml), 'externalBook');
      if (!libroExt) continue;
      const nombres = XmlEditar.hijo(libroExt, 'sheetNames');
      if (nombres) vinculo.hojas = XmlEditar.hijos(nombres, 'sheetName').map((x) => x.attrs.val || '');
      const conjunto = XmlEditar.hijo(libroExt, 'sheetDataSet');
      for (const sd of conjunto ? XmlEditar.hijos(conjunto, 'sheetData') : []) {
        const celdas = new Map();
        vinculo.datos.set(+sd.attrs.sheetId, { celdas, errorAlActualizar: sd.attrs.refreshError === '1' || sd.attrs.refreshError === 'true' });
        for (const fila of XmlEditar.hijos(sd, 'row')) for (const cel of XmlEditar.hijos(fila, 'cell')) {
          const ref = LibroXlsx.decodificarRango(cel.attrs.r || '');
          const v = XmlEditar.hijo(cel, 'v');
          if (!ref || !v || v.iniCierre < 0) continue;
          const texto = XmlEditar.decodificar(xml.slice(v.finApertura, v.iniCierre));
          const t = cel.attrs.t || 'n';
          celdas.set(ref.f0 * MAX_COLS + ref.c0, t === 'n' ? Number(texto) : t === 'b' ? texto === '1' || texto === 'true' : t === 'e' ? CalculoHojas.error(texto) : texto);
        }
      }
    }
    return lista;
  }

  // Matrices dinámicas de Excel 365: la celda ancla lleva cm="n" (metadata de celda) cuyo tipo es XLDAPR en xl/metadata.xml. Devuelve el
  // índice cm que marca «dinámica» (null si el libro no tiene ninguno), la ruta de metadata.xml y la del libro.
  async function leerMetadataDinamica(paquete) {
    const principal = (await paquete.relaciones('')).find((r) => /\/officeDocument$/.test(r.tipo));
    if (!principal) return { cm: null, ruta: null, rutaLibro: null };
    const rel = (await paquete.relaciones(principal.destino)).find((r) => /\/sheetMetadata$/.test(r.tipo) && !r.externa);
    const xml = rel ? await paquete.leer(rel.destino) : null;
    if (!xml) return { cm: null, ruta: rel ? rel.destino : null, rutaLibro: principal.destino, existe: false };
    const raiz = XmlEditar.analizar(xml, { profundidad: 4 });
    const tipos = XmlEditar.hijo(raiz, 'metadataTypes'), celdas = XmlEditar.hijo(raiz, 'cellMetadata');
    const nombres = tipos ? XmlEditar.hijos(tipos, 'metadataType').map((t) => t.attrs.name) : [];
    let cm = null;
    if (celdas) XmlEditar.hijos(celdas, 'bk').forEach((bk, k) => {
      const rc = XmlEditar.hijo(bk, 'rc');
      if (cm === null && rc && nombres[+rc.attrs.t - 1] === 'XLDAPR') cm = k + 1;
    });
    return { cm, ruta: rel.destino, rutaLibro: principal.destino, existe: true };
  }
  const METADATA_DINAMICA = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
    + '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>'
    + '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>'
    + '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>';

  async function abrir(bytes) {
    const libro = await LibroXlsx.abrir(bytes);
    const paquete = PaqueteEditable.desdeZip(libro.zip);
    const hojas = [];
    for (let i = 0; i < libro.hojas.length; i++) hojas[i] = await libro.cargarHoja(i);
    const estilos = await EstilosLibro.crear(libro, paquete);
    // nombre visible de cada tabla (displayName: el que usan las fórmulas; el modelo guarda name) y nombres de sus columnas
    for (const hoja of hojas) {
      if (!hoja || hoja.sinCeldas || !(hoja.tablas && hoja.tablas.length)) continue;
      for (const ruta of hoja.partes.tablas || []) {
        const xml = await paquete.leer(ruta);
        if (!xml) continue;
        const raiz = XmlEditar.analizar(xml, { profundidad: 3 });
        const ref = LibroXlsx.decodificarRango(raiz.attrs.ref || '');
        const t = ref && hoja.tablas.find((x) => x.ref.f0 === ref.f0 && x.ref.c0 === ref.c0 && x.ref.f1 === ref.f1 && x.ref.c1 === ref.c1);
        if (!t) continue;
        t.nombreVisible = raiz.attrs.displayName || t.nombre;
        const cols = XmlEditar.hijo(raiz, 'tableColumns');
        if (cols) t.nombresColumnas = XmlEditar.hijos(cols, 'tableColumn').map((x) => XmlEditar.decodificar(x.attrs.name || ''));
        t.calculadas = new Map();
        (xml.match(/<tableColumn\b[^>]*?(?:\/>|>[\s\S]*?<\/tableColumn>)/g) || []).forEach((e, k) => { const q = /<calculatedColumnFormula\b[^>]*>([\s\S]*?)<\/calculatedColumnFormula>/.exec(e); if (q) t.calculadas.set(k, XmlEditar.decodificar(q[1])); });
      }
    }
    const externos = await leerVinculosExternos(paquete);
    // anclas de matrices dinámicas (cm de XLDAPR) y las celdas que desbordan (las que no son el ancla guardan solo su valor)
    const meta = await leerMetadataDinamica(paquete);
    for (let i = 0; i < hojas.length && meta.cm !== null; i++) {
      const hoja = hojas[i];
      if (!hoja || hoja.sinCeldas || !libro.hojas[i].ruta) continue;
      const xml = await paquete.leer(libro.hojas[i].ruta);
      if (!xml || xml.indexOf(' cm="') < 0) continue;
      const re = /<c\b([^>]*\bcm="(\d+)"[^>]*)>/g;
      let m;
      while ((m = re.exec(xml))) {
        if (+m[2] !== meta.cm) continue;
        const r = /\br="([A-Za-z]{1,3}\d+)"/.exec(m[1]);
        const pos = r && LibroXlsx.decodificarRango(r[1]);
        if (!pos) continue;
        const C = hoja.celdas, k = C.indiceDe(pos.f0, pos.c0), fx = k >= 0 && C.formulas.get(k);
        if (!fx || fx.tipo !== 'array') continue;
        fx.dinamica = true;
        const rango = LibroXlsx.decodificarRango(fx.ref || '');
        if (!rango) continue;
        if (!hoja.derramadas) hoja.derramadas = new Set();
        for (let f = rango.f0; f <= rango.f1; f++) for (let c = rango.c0; c <= rango.c1; c++) if (f !== pos.f0 || c !== pos.c0) hoja.derramadas.add(clave(f, c));
      }
    }
    const protegidas = hojas.map((h) => (h.sinCeldas ? () => false : XlsxEscritor.celdasProtegidas(h)));

    // la primera vez que se guarda una matriz dinámica en un libro sin metadata.xml, se crea (como Excel); si el libro ya trae una
    // metadata sin XLDAPR, la matriz se guarda como las de Ctrl+Mayús+Intro (sin cm), que Excel abre igual pero con rango fijo
    const metaDinamica = { cm: meta.cm };
    async function asegurarMetadata() {
      if (metaDinamica.cm !== null || meta.existe || !meta.rutaLibro) return;
      const dir = meta.rutaLibro.replace(/[^/]*$/, '');
      paquete.escribir(dir + 'metadata.xml', METADATA_DINAMICA);
      await paquete.asegurarTipo('/' + dir + 'metadata.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml');
      await paquete.agregarRelacion(meta.rutaLibro, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata', 'metadata.xml');
      meta.existe = true; metaDinamica.cm = 1;
    }
    async function guardarCeldas(porHoja) {
      const conMatriz = porHoja.some(([, lista]) => lista.some((x) => x.valor && x.valor.matriz && !x.valor.fija));
      if (conMatriz) await asegurarMetadata();
      for (const [, lista] of porHoja) for (const x of lista) if (x.valor && x.valor.matriz && !x.valor.fija) x.valor.cm = metaDinamica.cm;
      for (const [h, lista] of porHoja) {
        const hoja = hojas[h];
        await XlsxEscritor.editarCeldas(libro, h, lista);
        libro._hojasCargadas.set(h, hoja);   // el modelo en memoria ya tiene los cambios: no se vuelve a leer
      }
      paquete.olvidarCambios();
      return paquete.generar();
    }

    return {
      libro, paquete, hojas, estilos, externos,
      protegida: (h, f, c) => protegidas[h](f, c),
      guardarCeldas,
      destino: () => (libro.tieneMacros
        ? { extension: 'xlsm', mime: 'application/vnd.ms-excel.sheet.macroEnabled.12' }
        : { extension: 'xlsx', mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
    };
  }

  globalThis.XlsxBackend = { formato: 'xlsx', reconoce: async () => true, abrir };   // el último de la lista: reconoce todo lo que no es de otro formato
})();

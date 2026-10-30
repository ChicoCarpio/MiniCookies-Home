/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Paquete editable de Office (OOXML) para los editores de la suite (núcleo, fase 1 del PLAN). Sin DOM: corre igual en Node.
//
// La regla de la suite es la IDA Y VUELTA SIN PÉRDIDA: lo que el editor no toca se guarda byte a byte como vino (zip-propio copia las
// entradas sin descomprimir ni recomprimir) y en lo que toca se cambia solo el fragmento editado (xml-editar). Este módulo junta las
// operaciones de paquete que necesitan los tres editores: leer y escribir partes, tipos de contenido, relaciones y comparar dos paquetes.
//
// Requiere zip-propio.js y xml-editar.js.
//
//   const p = await PaqueteEditable.abrir(bytes)      // o PaqueteEditable.desdeZip(zipPropio)
//   await p.leer('xl/styles.xml') → texto | null
//   p.escribir('xl/styles.xml', texto)                // la parte queda marcada como cambiada
//   p.existe(ruta) / p.rutas() / p.cambiadas()
//   await p.asegurarTipo('/xl/styles.xml', 'application/vnd…+xml')   // Override en [Content_Types].xml si falta
//   await p.relaciones('xl/workbook.xml') → [{ id, tipo, destino, externa }]
//   await p.agregarRelacion('xl/workbook.xml', tipoUri, destinoRelativo) → id
//   await p.generar() → Uint8Array (DEFLATE en las partes nuevas o cambiadas)
//   await PaqueteEditable.comparar(bytesA, bytesB) → { iguales: [...], distintas: [...], soloA: [...], soloB: [...] }
(function () {
  const X = () => globalThis.XmlEditar;
  const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

  // Ruta de la parte .rels de una parte ('xl/workbook.xml' → 'xl/_rels/workbook.xml.rels'; '' → '_rels/.rels')
  function rutaRels(parte) {
    const p = parte.replace(/^\//, '');
    const k = p.lastIndexOf('/');
    return (k < 0 ? '' : p.slice(0, k + 1)) + '_rels/' + p.slice(k + 1) + '.rels';
  }
  // Destino de una relación resuelto desde la carpeta de la parte de origen ('../media/a.png' desde 'xl/drawings/d.xml' → 'xl/media/a.png')
  function resolver(parte, destino) {
    if (destino.startsWith('/')) return destino.slice(1);
    const base = parte.replace(/^\//, '').split('/');
    base.pop();
    for (const trozo of destino.split('/')) {
      if (trozo === '..') base.pop();
      else if (trozo !== '.' && trozo !== '') base.push(trozo);
    }
    return base.join('/');
  }

  function desdeZip(zip) {
    const cambiadas = new Set();
    const leer = async (ruta) => { const e = zip.file(ruta); return e ? e.async('string') : null; };
    const leerBytes = async (ruta) => { const e = zip.file(ruta); return e ? e.async('uint8array') : null; };
    function escribir(ruta, datos) { zip.file(ruta, datos); cambiadas.add(ruta); }
    function quitar(ruta) { zip.remove(ruta); cambiadas.add(ruta); }

    async function asegurarTipo(parte, tipo) {
      const nombre = '/' + parte.replace(/^\//, '');
      let xml = await leer('[Content_Types].xml');
      if (xml === null) throw new Error('el paquete no tiene [Content_Types].xml');
      const raiz = X().analizar(xml, { profundidad: 2 });
      for (const h of raiz.hijos) if (h.local === 'Override' && (h.attrs.PartName || '').toLowerCase() === nombre.toLowerCase()) return false;
      const ext = (nombre.match(/\.([^./]+)$/) || [])[1];
      for (const h of raiz.hijos) if (h.local === 'Default' && ext && (h.attrs.Extension || '').toLowerCase() === ext.toLowerCase() && h.attrs.ContentType === tipo) return false;
      xml = X().aplicar(xml, [{ ini: raiz.iniCierre, fin: raiz.iniCierre, texto: `<Override PartName="${X().escAttr(nombre)}" ContentType="${X().escAttr(tipo)}"/>` }]);
      escribir('[Content_Types].xml', xml);
      return true;
    }

    async function relaciones(parte) {
      const xml = await leer(rutaRels(parte));
      if (xml === null) return [];
      const raiz = X().analizar(xml, { profundidad: 2 });
      return raiz.hijos.filter((h) => h.local === 'Relationship').map((h) => ({
        id: h.attrs.Id, tipo: h.attrs.Type || '', destino: h.attrs.TargetMode === 'External' ? h.attrs.Target : resolver(parte, h.attrs.Target || ''), externa: h.attrs.TargetMode === 'External',
      }));
    }
    async function agregarRelacion(parte, tipo, destino) {
      const ruta = rutaRels(parte);
      let xml = await leer(ruta);
      if (xml === null) xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${RELS_NS}"></Relationships>`;
      const raiz = X().analizar(xml, { profundidad: 2 });
      const usados = new Set(raiz.hijos.map((h) => h.attrs.Id));
      let n = raiz.hijos.length + 1;
      while (usados.has('rId' + n)) n++;
      const id = 'rId' + n;
      const nodo = `<Relationship Id="${id}" Type="${X().escAttr(tipo)}" Target="${X().escAttr(destino)}"/>`;
      if (raiz.iniCierre < 0) xml = X().aplicar(xml, [{ ini: raiz.ini, fin: raiz.fin, texto: xml.slice(raiz.ini, raiz.fin).replace(/\s*\/>$/, '>') + nodo + '</Relationships>' }]);
      else xml = X().aplicar(xml, [{ ini: raiz.iniCierre, fin: raiz.iniCierre, texto: nodo }]);
      escribir(ruta, xml);
      return id;
    }

    return {
      zip, leer, leerBytes, escribir, quitar, asegurarTipo, relaciones, agregarRelacion,
      existe: (ruta) => !!zip.file(ruta),
      rutas: () => Object.keys(zip.files).filter((r) => !zip.files[r].dir),
      cambiadas: () => [...cambiadas],
      olvidarCambios: () => cambiadas.clear(),
      generar: () => zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
    };
  }

  async function abrir(bytes) { return desdeZip(await globalThis.ZipPropio.loadAsync(bytes)); }

  // Qué entradas cambiaron entre dos paquetes (por contenido descomprimido). Para las pruebas de ida y vuelta.
  async function comparar(a, b) {
    const za = await globalThis.ZipPropio.loadAsync(a), zb = await globalThis.ZipPropio.loadAsync(b);
    const ra = Object.keys(za.files).filter((r) => !za.files[r].dir), rb = new Set(Object.keys(zb.files).filter((r) => !zb.files[r].dir));
    const r = { iguales: [], distintas: [], soloA: [], soloB: [] };
    for (const ruta of ra) {
      if (!rb.has(ruta)) { r.soloA.push(ruta); continue; }
      rb.delete(ruta);
      const x = await za.file(ruta).async('uint8array'), y = await zb.file(ruta).async('uint8array');
      let igual = x.length === y.length;
      for (let i = 0; igual && i < x.length; i++) if (x[i] !== y[i]) igual = false;
      (igual ? r.iguales : r.distintas).push(ruta);
    }
    r.soloB = [...rb];
    return r;
  }

  globalThis.PaqueteEditable = { abrir, desdeZip, comparar, rutaRels, resolver };
})();

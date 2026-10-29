/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Backend de .ods del libro editable (F3.4a). Misma forma que XlsxBackend (ver hojas/xlsx-backend.js): reconoce el archivo, lo abre en un modelo con la forma de
// LibroXlsx (ods-modelo.js), dice qué celdas no se pueden editar, crea estilos nuevos y guarda SOLO las celdas editadas en content.xml (ods-escritor-celdas.js);
// el resto del paquete (meta.xml, imágenes, objetos incrustados) no se toca; styles.xml, settings.xml, el manifiesto y `Object N/` solo los cambian los cambios de estructura, de
// página y de gráficos (ods-estructura.js, ods-pagina.js, ods-graficos.js).
// Requiere zip-propio, paquete-editable, xlsx-estilos, motores/ods-modelo, ods-estilos, ods-estilos-escritura, ods-escritor-celdas y, para las celdas protegidas,
// motores/xlsx-escritor (celdasProtegidas lee solo el modelo de la hoja).
(function () {
  const MIME_HOJA = 'application/vnd.oasis.opendocument.spreadsheet';

  // ¿Es una hoja de cálculo OpenDocument? El mimetype (primera entrada, sin comprimir) o, si el programa que lo escribió no lo puso primero, el cuerpo de content.xml.
  async function reconoce(bytes) {
    if (!bytes || bytes[0] !== 0x50 || bytes[1] !== 0x4b) return false;
    let zip;
    try { zip = await ZipPropio.loadAsync(bytes); } catch (e) { return false; }
    const mime = zip.file('mimetype');
    if (mime) return (await mime.async('string')).trim() === MIME_HOJA;
    const contenido = zip.file('content.xml');
    return !!contenido && (await contenido.async('string')).indexOf('<office:spreadsheet') >= 0;
  }

  async function abrir(bytes, opciones) {
    const o = opciones || {};
    const libro = await OdsModelo.abrir(bytes);
    const paquete = PaqueteEditable.desdeZip(libro.zip);
    const hojas = [];
    for (let i = 0; i < libro.hojas.length; i++) hojas[i] = await libro.cargarHoja(i);
    const protegidas = hojas.map((h) => XlsxEscritor.celdasProtegidas(h));
    // los gráficos (objetos incrustados): el motor los ve como los de Excel, en una vista del paquete; lo que se guarda es `paquete`
    if (globalThis.OdsGraficos) await OdsGraficos.preparar(libro, hojas, paquete);
    const tabla = libro._tablaXf;
    const contadores = { estilo: 0, numero: 0 };

    // estilos nuevos de celda (negrita, color, formato…): un xf más en el styles.xml virtual y, al guardar, un estilo automático en content.xml
    const memo = new Map();
    const estilos = {
      derivar(base, cambios) {
        const k = base + '|' + JSON.stringify(cambios);
        if (memo.has(k)) return memo.get(k);
        const i = tabla.derivar(base, cambios);
        libro.estilos = XlsxEstilos.leer(tabla.xml(), null);   // la rejilla y la cinta leen libro.estilos: se deja al día
        memo.set(k, i);
        return i;
      },
      propiedades(indice) {
        const x = libro.estilos.xf(indice);
        const color = (ref) => { const c = ref ? libro.estilos.color(ref) : null; return c ? c.slice(1).toUpperCase() : null; };
        const f = x.fuente || {};
        return {
          negrita: !!f.negrita, cursiva: !!f.cursiva, subrayado: f.subrayado || false, tachado: !!f.tachado, colorFuente: color(f.color), tamano: f.tam || 10, fuente: f.nombre || 'Liberation Sans',
          relleno: x.relleno && x.relleno.tipo === 'patron' && x.relleno.patron === 'solid' ? color(x.relleno.fg) : null,
          h: (x.alineacion && x.alineacion.h) || 'general', v: (x.alineacion && x.alineacion.v) || 'bottom', ajustar: !!(x.alineacion && x.alineacion.ajustar), formato: x.formato || 'General',
        };
      },
    };

    // Un guardado a la vez, como en xlsx: lo que se edite mientras tanto queda para el siguiente
    async function guardarCeldas(porHoja) {
      if (!porHoja.length) return paquete.generar();   // nada cambió (o se deshizo todo): el archivo sale como vino
      const texto = await paquete.leer('content.xml');
      if (texto === null) throw new Error('ods: falta content.xml');
      const r = OdsEscritorCeldas.guardar({ texto, porHoja, libro, hojas, tabla, contadores, estilos: libro._estilosOds, idioma: o.idioma });
      libro.avisos.formulasNoGuardadas = (libro.avisos.formulasNoGuardadas || 0) + r.avisos.formulasNoGuardadas;
      paquete.escribir('content.xml', r.texto);
      const bytes = await paquete.generar();
      paquete.olvidarCambios();
      return bytes;
    }

    return {
      libro, paquete, hojas, estilos, externos: [],
      protegida: (h, f, c) => protegidas[h](f, c),
      guardarCeldas,
      destino: () => ({ extension: 'ods', mime: MIME_HOJA }),
    };
  }

  globalThis.OdsBackend = { formato: 'ods', reconoce, abrir };
})();

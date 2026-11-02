/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Cambio de versión de hojas de cálculo: pasa un libro de cualquier formato que el sitio sabe leer (.xlsx/.xlsm, .xls, .ods, .xlsb, XML 2003) a .xlsx, .xls o .ods, conservando
// fórmulas, formatos de número, estilos, combinadas, anchos, altos, ocultos y nombres definidos. Todos los lectores devuelven el mismo contrato de celdas y todos los escritores lo
// reciben (ver XlsEscritor), así que cualquier origen sirve a cualquier destino.
//
//   await CambioVersion.detectar(bytes)                          → 'xlsx' | 'xls' | 'ods' | 'xlsb' | 'xml' | 'html' | null
//   await CambioVersion.convertir(bytes, 'xlsx' | 'xls' | 'ods', { formulas: true })
//        → { bytes: Uint8Array, origen, avisos: [{ tipo, cantidad, ejemplo? }] }
//
// formulas: false deja solo los valores (lo que se veía). Los avisos son lo que el formato de destino no puede llevar (ver XlsEscritor y OdsEscritor).
// Requiere: formatos-antiguos.js (con sus lectores), xlsx-contrato.js, xls-formulas.js, ods-formulas.js y los escritores que se usen.
(function () {
  const bytesIgual = (b, firma) => b.length >= firma.length && firma.every((x, k) => b[k] === x);

  async function detectar(bytes) {
    if (bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4B) {
      let zip;
      try { zip = await ZipPropio.loadAsync(bytes); } catch (e) { return null; }
      if (zip.file('xl/workbook.xml')) return 'xlsx';
    }
    return FormatosAntiguos.detectar(bytes);
  }

  const esCelda = (x) => x && typeof x === 'object' && !(x instanceof Date) && typeof x.error !== 'string';

  // Solo valores: se quitan las fórmulas (queda el último resultado, con su formato y estilo).
  function sinFormulas(hojas) {
    for (const h of hojas) for (const fila of h.filas) {
      if (!fila) continue;
      for (let c = 0; c < fila.length; c++) {
        const x = fila[c];
        if (!esCelda(x) || !('formula' in x)) continue;
        const { formula, matriz, ...resto } = x;
        let v = resto.valor;
        if (esCelda(v) && 'valor' in v) { resto.formato = resto.formato === undefined ? v.formato : resto.formato; resto.estilo = resto.estilo === undefined ? v.estilo : resto.estilo; v = v.valor; }
        resto.valor = v === undefined ? null : v;
        fila[c] = resto.formato === undefined && resto.estilo === undefined ? resto.valor : resto;
      }
    }
  }

  // La base de fechas 1904 (Mac antiguo) se pasa a la 1900 como hace Excel al abrir el libro, salvo en los formatos de solo horas.
  function a1900(hojas) {
    const mueve = (v, formato) => {
      const codigo = typeof formato === 'number' ? XlsxEstilos.FORMATOS_INTEGRADOS[formato] : formato;
      return typeof v === 'number' && codigo && XlsxFormato.esFecha(codigo) && XlsxFormato.tieneCalendario(codigo) ? v + 1462 : v;
    };
    for (const h of hojas) for (const fila of h.filas) {
      if (!fila) continue;
      for (let c = 0; c < fila.length; c++) {
        const x = fila[c];
        if (!esCelda(x)) continue;
        // el formato va al lado del valor, o dentro del resultado de una fórmula ({ formula, valor: { valor, formato } }, como la deja XlsLector)
        if (typeof x.valor === 'number' && x.formato) fila[c] = { ...x, valor: mueve(x.valor, x.formato) };
        else if (esCelda(x.valor) && typeof x.valor.valor === 'number' && (x.valor.formato || x.formato)) fila[c] = { ...x, valor: { ...x.valor, valor: mueve(x.valor.valor, x.valor.formato || x.formato) } };
      }
    }
  }

  async function leerContrato(bytes, origen, conFormulas) {
    if (origen === 'xlsx') return XlsxContrato.leer(bytes);
    if (origen === 'xls') return XlsLector.leer(bytes, { formulas: conFormulas });
    if (origen === 'ods') return OdsLector.leer(bytes, { formulas: conFormulas });
    return FormatosAntiguos.leer(bytes);
  }

  async function convertir(bytes, destino, opciones) {
    opciones = opciones || {};
    const conFormulas = opciones.formulas !== false;
    const origen = await detectar(bytes);
    if (!origen) throw new Error('el archivo no es una hoja de cálculo que se pueda leer');
    if (destino !== 'xlsx' && destino !== 'xls' && destino !== 'ods') throw new Error('formato de destino desconocido: ' + destino);
    const libro = await leerContrato(bytes, origen, conFormulas);   // antes que la comprobación de formato: un archivo con contraseña tiene que decir eso
    if (origen === destino) { const e = new Error('el archivo ya está en ese formato'); e.codigo = 'mismo-formato'; throw e; }
    if (!conFormulas) sinFormulas(libro.hojas);
    // los nombres locales llegan con el nombre de su hoja (xls) o ya con el índice: se dejan con el índice, que entienden los tres escritores
    const nombres = (libro.nombres || []).filter((n) => !n.integrado && n.formula).map((n) => {
      let hoja = n.hoja;
      if (typeof hoja === 'string') { const k = libro.hojas.findIndex((h) => h.nombre === hoja); hoja = k >= 0 ? k : undefined; }
      return { nombre: n.nombre, formula: n.formula, ...(hoja !== undefined && hoja !== null ? { hoja } : {}) };
    });
    const avisos = [];
    let salida;
    if (destino === 'xlsx') {
      if (libro.fecha1904) a1900(libro.hojas);
      salida = await XlsxEscritor.crearLibro(libro.hojas, { nombres: conFormulas ? nombres : [] });
    } else if (destino === 'xls') {
      const r = XlsEscritor.crearLibro(libro.hojas, { formulas: conFormulas, fecha1904: !!libro.fecha1904, nombres: conFormulas ? nombres : [] });
      salida = r.bytes; avisos.push(...r.avisos);
    } else {
      const r = await OdsEscritor.crearLibro(libro.hojas, { formulas: conFormulas, fecha1904: !!libro.fecha1904, nombres: conFormulas ? nombres : [] });
      salida = r.bytes; avisos.push(...r.avisos);
    }
    // las fórmulas que el lector dejó en su valor por los topes contra archivos hostiles (XlsLector)
    if (conFormulas && libro.formulasOmitidas) {
      const a = avisos.find((x) => x.tipo === 'formula-sin-llevar');
      if (a) a.cantidad += libro.formulasOmitidas; else avisos.push({ tipo: 'formula-sin-llevar', cantidad: libro.formulasOmitidas });
    }
    return { bytes: salida, origen, avisos };
  }

  globalThis.CambioVersion = { detectar, convertir };
})();

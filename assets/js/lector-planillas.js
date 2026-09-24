// Lectura de planillas compartida por las herramientas de Excel/CSV. Requiere
// SheetJS (XLSX) cargado antes.
//
// Los .xlsx/.xls/.ods se leen tal cual. Los CSV necesitan tratamiento aparte:
// - XLSX.read(bytes, {type:'array'}) interpreta un CSV UTF-8 sin BOM como
//   Latin-1 ("José" quedaba "JosÃ©"). Acá se decodifica el texto primero:
//   UTF-8 (o UTF-16 si tiene BOM), y si no es UTF-8 válido, Windows-1252 (el
//   CSV que guarda Excel en Windows en español).
// - Con separador ";" (Excel en español) los números vienen con coma decimal
//   y punto de miles: SheetJS leía "1.234,50" como 1.2345. Acá se lee todo
//   como texto (raw: true) y los números se convierten con el separador
//   decimal que corresponde al archivo. "00123" (códigos, cédulas) queda como
//   texto para no perder los ceros a la izquierda.

function esArchivoCsv(archivo) {
  return /\.(csv|tsv|txt)$/i.test(archivo.name) || archivo.type === 'text/csv';
}

function decodificarTextoPlanilla(bytes) {
  if (bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) return new TextDecoder('utf-8').decode(bytes.subarray(3));
  if (bytes[0] === 0xFF && bytes[1] === 0xFE) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
  if (bytes[0] === 0xFE && bytes[1] === 0xFF) return new TextDecoder('utf-16be').decode(bytes.subarray(2));
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch (e) {
    return new TextDecoder('windows-1252').decode(bytes);
  }
}

// Separador más frecuente en las primeras líneas, contando solo fuera de comillas.
function detectarSeparadorCsv(texto) {
  const muestra = texto.slice(0, 20000);
  const cuentas = { ',': 0, ';': 0, '\t': 0, '|': 0 };
  let entreComillas = false;
  let lineas = 0;
  for (const caracter of muestra) {
    if (caracter === '"') entreComillas = !entreComillas;
    else if (!entreComillas && caracter in cuentas) cuentas[caracter]++;
    else if (!entreComillas && caracter === '\n' && ++lineas >= 20) break;
  }
  return Object.keys(cuentas).reduce((a, b) => (cuentas[b] > cuentas[a] ? b : a), ',');
}

// Con ";" la coma es decimal (convención de Excel en español); con "," el
// punto. Con tabulador o "|" se decide por lo que más aparezca en los datos.
function detectarComaDecimal(texto, separador) {
  if (separador === ';') return true;
  if (separador === ',') return false;
  const conComa = (texto.match(/\d,\d/g) || []).length;
  const conPunto = (texto.match(/\d\.\d/g) || []).length;
  return conComa > conPunto;
}

// Devuelve el número que representa el texto, o null si no es un número
// "limpio" (así los códigos con ceros a la izquierda, fechas, etc. quedan como texto).
function interpretarNumeroTexto(texto, comaDecimal) {
  const valor = String(texto).trim();
  if (!valor) return null;
  const miles = comaDecimal ? '\\.' : ',';
  const decimal = comaDecimal ? ',' : '\\.';
  const patron = new RegExp(`^[-+]?(\\d{1,3}(${miles}\\d{3})+|\\d+)(${decimal}\\d+)?$`);
  if (!patron.test(valor)) return null;
  const sinSigno = valor.replace(/^[-+]/, '');
  if (/^0\d/.test(sinSigno)) return null;
  const normalizado = comaDecimal
    ? valor.replace(/\./g, '').replace(',', '.')
    : valor.replace(/,/g, '');
  const numero = Number(normalizado);
  return Number.isFinite(numero) ? numero : null;
}

function leerCsvComoLibro(texto) {
  const separador = detectarSeparadorCsv(texto);
  const comaDecimal = detectarComaDecimal(texto, separador);
  const libro = XLSX.read(texto, { type: 'string', raw: true, FS: separador });
  libro.SheetNames.forEach((nombre) => {
    const hoja = libro.Sheets[nombre];
    Object.keys(hoja).forEach((direccion) => {
      if (direccion[0] === '!') return;
      const celda = hoja[direccion];
      if (celda.t !== 's') return;
      const numero = interpretarNumeroTexto(celda.v, comaDecimal);
      if (numero !== null) {
        // Se conserva el texto original en "w": así las vistas previas
        // muestran "1.234,50" tal como venía en el archivo.
        celda.w = celda.v;
        celda.t = 'n';
        celda.v = numero;
      }
    });
  });
  libro.separadorCsv = separador;
  libro.comaDecimalCsv = comaDecimal;
  return libro;
}

// opciones: las mismas de XLSX.read (cellStyles, cellDates...), se aplican
// solo a los archivos que no son CSV.
async function leerLibroDesdeArchivo(archivo, opciones) {
  const bytes = new Uint8Array(await archivo.arrayBuffer());
  if (esArchivoCsv(archivo)) return leerCsvComoLibro(decodificarTextoPlanilla(bytes));
  return XLSX.read(bytes, Object.assign({ type: 'array' }, opciones));
}

// Copia una fila de celdas de una hoja de SheetJS a otra conservando el
// tipo y el formato de número de cada celda (fechas, monedas, porcentajes).
// Armar la hoja nueva con aoa_to_sheet a partir del texto formateado
// (raw: false) dejaba números y fechas como texto en el Excel resultante.
// Las fórmulas se conservan solo si la fila no cambia de lugar (si se mueve,
// sus referencias relativas apuntarían a otras celdas): si no, queda el valor.
// Las columnas columnaInicio..columnaFin del origen se escriben desde
// columnaDestino en adelante.
function copiarFilaDeHoja(origen, filaOrigen, destino, filaDestino, columnaInicio, columnaFin, columnaDestino) {
  const desplazamiento = (columnaDestino === undefined ? columnaInicio : columnaDestino) - columnaInicio;
  for (let c = columnaInicio; c <= columnaFin; c++) {
    const celda = origen[XLSX.utils.encode_cell({ r: filaOrigen, c })];
    if (!celda || celda.t === 'z') continue;
    const copia = { t: celda.t, v: celda.v };
    if (celda.z !== undefined) copia.z = celda.z;
    if (celda.w !== undefined) copia.w = celda.w;
    if (celda.f !== undefined && filaOrigen === filaDestino && desplazamiento === 0) copia.f = celda.f;
    destino[XLSX.utils.encode_cell({ r: filaDestino, c: c + desplazamiento })] = copia;
  }
}

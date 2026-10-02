/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Lectura de archivos de texto del usuario (.txt, .md, .csv...) con la
// codificación correcta. archivo.text() asume UTF-8: un .txt guardado en
// Windows como "ANSI" (Windows-1252) o como "Unicode" (UTF-16) mostraba los
// acentos rotos ("canciÃ³n", "�"). Misma detección que decodificarTextoPlanilla
// de lector-planillas.js, más UTF-16 sin BOM.
//
// leerTextoArchivo(archivo) → { texto, finDeLinea }
// - texto: siempre con saltos "\n" (así cada herramienta procesa igual un
//   archivo de Windows que uno de Mac/Linux);
// - finDeLinea: "\r\n" o "\n", el que usaba el archivo, para devolverlo igual
//   con conFinDeLinea(texto, finDeLinea).

function decodificarTexto(bytes) {
  if (bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) return new TextDecoder('utf-8').decode(bytes.subarray(3));
  if (bytes[0] === 0xFF && bytes[1] === 0xFE) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
  if (bytes[0] === 0xFE && bytes[1] === 0xFF) return new TextDecoder('utf-16be').decode(bytes.subarray(2));
  // UTF-16 sin BOM: texto latino con un byte 0 cada dos.
  const muestra = Math.min(bytes.length, 4000);
  let cerosPares = 0, cerosImpares = 0;
  for (let i = 0; i < muestra; i++) if (bytes[i] === 0) { if (i % 2) cerosImpares++; else cerosPares++; }
  if (muestra >= 4) {
    if (cerosImpares > muestra * 0.3 && cerosPares < muestra * 0.05) return new TextDecoder('utf-16le').decode(bytes);
    if (cerosPares > muestra * 0.3 && cerosImpares < muestra * 0.05) return new TextDecoder('utf-16be').decode(bytes);
  }
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch (e) {
    return new TextDecoder('windows-1252').decode(bytes);
  }
}

async function leerTextoArchivo(archivo) {
  const crudo = decodificarTexto(new Uint8Array(await archivo.arrayBuffer()));
  const conCrLf = (crudo.match(/\r\n/g) || []).length;
  const soloLf = (crudo.match(/(^|[^\r])\n/g) || []).length;
  return { texto: crudo.replace(/\r\n?/g, '\n'), finDeLinea: conCrLf > soloLf ? '\r\n' : '\n' };
}

function conFinDeLinea(texto, finDeLinea) {
  return finDeLinea === '\r\n' ? texto.replace(/\n/g, '\r\n') : texto;
}

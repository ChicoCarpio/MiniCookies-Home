/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Retoques de texto plano: la lógica de «Ordenar líneas», «Quitar líneas duplicadas o vacías», «Encabezado / pie de
// texto» y «Arreglar codificación», para que la página de cada herramienta y su módulo de Flujos (flujos-modulos.js)
// hagan lo mismo con el mismo código. No toca la página: también corre en los hilos de trabajo de Flujos.
//
// Todas reciben el texto con saltos «\n» (como lo entrega leerTextoArchivo de leer-texto.js) y lo devuelven igual.
//   TextoRetoques.ordenar(texto, { orden, ignorarMayusculas, idioma }) -> texto
//        orden: 'alfabetico-asc' | 'alfabetico-desc' | 'numerico-asc' | 'numerico-desc'
//   TextoRetoques.quitarLineas(texto, { vacias, duplicadas }) -> { texto, total, quedan, duplicadasQuitadas }
//   TextoRetoques.encabezadoPie(texto, encabezado, pie) -> texto
//   TextoRetoques.arreglarCodificacion(bytes) -> { texto, codificacion, conBom, utf16SinBom, mojibake, saltosMixtos }
//        bytes: Uint8Array del archivo. texto: ya corregido, con saltos «\n». codificacion: 'utf-8' | 'utf-16le' |
//        'utf-16be' | 'windows-1252' (la que se usó para leerlo). mojibake: cuántos caracteres dañados se repararon.
(function () {
  // El salto de línea final del archivo no es una línea más: se aparta y se vuelve a poner al final.
  function partir(texto) {
    const terminaEnSalto = texto.endsWith('\n');
    return { lineas: (terminaEnSalto ? texto.slice(0, -1) : texto).split('\n'), terminaEnSalto };
  }
  const unir = (lineas, terminaEnSalto) => lineas.join('\n') + (terminaEnSalto ? '\n' : '');

  // ---------- Ordenar líneas ----------
  // Número con el que EMPIEZA la línea (admite espacios y signo delante); si no
  // empieza con uno, null. Un punto o coma entre grupos de exactamente 3 cifras es
  // separador de miles ("1.200" → 1200); si no, es decimal ("3,5" → 3.5). Con
  // punto y coma a la vez, el último es el decimal ("1.234,50" → 1234.5).
  function numeroDeLinea(linea) {
    const m = linea.match(/^\s*([+-]?\d+(?:[.,]\d+)*)/);
    if (!m) return null;
    let s = m[1];
    const tienePunto = s.includes('.'), tieneComa = s.includes(',');
    if (tienePunto && tieneComa) {
      const decimal = s.lastIndexOf('.') > s.lastIndexOf(',') ? '.' : ',';
      const miles = decimal === '.' ? ',' : '.';
      s = s.split(miles).join('').replace(decimal, '.');
    } else if (tienePunto || tieneComa) {
      const partes = s.split(tienePunto ? '.' : ',');
      if (partes.slice(1).every((p) => p.length === 3)) s = partes.join('');
      else if (partes.length === 2) s = partes.join('.');
      else s = partes[0]; // "1.2.3": no es un número válido, se toma el primer grupo
    }
    const numero = parseFloat(s);
    return Number.isFinite(numero) ? numero : null;
  }

  function ordenar(texto, { orden, ignorarMayusculas, idioma }) {
    const { lineas, terminaEnSalto } = partir(texto);
    let ordenadas;
    if (String(orden).startsWith('alfabetico')) {
      // Orden de diccionario del idioma de la página: acentos y ñ en su
      // lugar ("árbol" junto a "arco", "ñandú" después de "nube"; antes iban
      // al final) y números como números ("tema 2" antes que "tema 10").
      const colador = new Intl.Collator(idioma || 'es', { numeric: true, sensitivity: ignorarMayusculas ? 'accent' : 'variant', caseFirst: 'upper' });
      ordenadas = [...lineas].sort(colador.compare);
      if (orden.endsWith('desc')) ordenadas.reverse();
    } else {
      const conNumero = [], sinNumero = [];
      lineas.forEach((linea) => {
        const numero = numeroDeLinea(linea);
        if (numero !== null) conNumero.push({ linea, numero });
        else sinNumero.push(linea);
      });
      conNumero.sort((a, b) => a.numero - b.numero);
      if (String(orden).endsWith('desc')) conNumero.reverse();
      ordenadas = [...conNumero.map((x) => x.linea), ...sinNumero];
    }
    return unir(ordenadas, terminaEnSalto);
  }

  // ---------- Quitar líneas duplicadas o vacías ----------
  function quitarLineas(texto, { vacias, duplicadas }) {
    const partes = partir(texto);
    let lineas = partes.lineas;
    const total = lineas.length;
    if (vacias) lineas = lineas.filter((l) => l.trim() !== '');
    let duplicadasQuitadas = 0;
    if (duplicadas) {
      const vistas = new Set();
      lineas = lineas.filter((linea) => {
        if (vistas.has(linea)) { duplicadasQuitadas++; return false; }
        vistas.add(linea);
        return true;
      });
    }
    return { texto: unir(lineas, partes.terminaEnSalto), total, quedan: lineas.length, duplicadasQuitadas };
  }

  // ---------- Encabezado / pie ----------
  function encabezadoPie(texto, encabezado, pie) {
    const partes = [];
    if (encabezado) partes.push(encabezado);
    partes.push(texto);
    if (pie) partes.push(pie);
    return partes.join('\n\n');
  }

  // ---------- Arreglar codificación ----------
  // Tabla inversa de Windows-1252 (carácter -> byte), para deshacer el
  // "mojibake": texto UTF-8 que en algún momento se leyó como Windows-1252 y
  // se volvió a guardar ("canción" convertido en "canciÃ³n", "’" en "â€™").
  const BYTE_WINDOWS_1252 = (() => {
    const mapa = new Map();
    const texto = new TextDecoder('windows-1252').decode(Uint8Array.from({ length: 256 }, (_, i) => i));
    [...texto].forEach((caracter, i) => mapa.set(caracter, i));
    return mapa;
  })();
  // Un carácter UTF-8 de 2 o 3 bytes leído como Windows-1252 se ve como una
  // letra "inicial" (Ã, Â, â...) seguida de 1 o 2 caracteres "de continuación"
  // (los que corresponden a los bytes 0x80-0xBF: ©, ³, €, ™, ƒ...).
  const PATRON_MOJIBAKE = (() => {
    const continuacion = new TextDecoder('windows-1252').decode(Uint8Array.from({ length: 64 }, (_, i) => 0x80 + i))
      .replace(/[\]\\^-]/g, '\\$&');
    return new RegExp(`[\\u00C2-\\u00DF][${continuacion}]|[\\u00E0-\\u00EF][${continuacion}]{2}`, 'g');
  })();
  const contarMojibake = (texto) => (texto.match(PATRON_MOJIBAKE) || []).length;

  // Intenta deshacer el mojibake (hasta 2 veces, por si se dañó dos veces).
  // Devuelve null si no hay nada que corregir o si la corrección no es segura.
  function corregirMojibake(texto) {
    let actual = texto;
    let corregido = false;
    for (let vuelta = 0; vuelta < 2 && contarMojibake(actual) > 0; vuelta++) {
      const bytes = [];
      for (const caracter of actual) {
        if (!BYTE_WINDOWS_1252.has(caracter)) return corregido ? actual : null;
        bytes.push(BYTE_WINDOWS_1252.get(caracter));
      }
      let candidato;
      try {
        candidato = new TextDecoder('utf-8', { fatal: true }).decode(Uint8Array.from(bytes));
      } catch (e) {
        return corregido ? actual : null;
      }
      if (contarMojibake(candidato) >= contarMojibake(actual)) break;
      actual = candidato;
      corregido = true;
    }
    return corregido ? actual : null;
  }

  // UTF-16 sin BOM: texto en alfabeto latino tiene un byte 0 cada dos.
  function pareceUtf16(bytes) {
    const muestra = Math.min(bytes.length, 4000);
    if (muestra < 4) return null;
    let cerosPares = 0, cerosImpares = 0;
    for (let i = 0; i < muestra; i++) {
      if (bytes[i] === 0) { if (i % 2 === 0) cerosPares++; else cerosImpares++; }
    }
    if (cerosImpares > muestra * 0.3 && cerosPares < muestra * 0.05) return 'utf-16le';
    if (cerosPares > muestra * 0.3 && cerosImpares < muestra * 0.05) return 'utf-16be';
    return null;
  }

  function arreglarCodificacion(bytes) {
    let texto, codificacion = 'utf-8', conBom = false, utf16SinBom = false;
    if (bytes.length >= 3 && bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) {
      // BOM UTF-8: EF BB BF al principio del archivo.
      conBom = true;
      bytes = bytes.slice(3);
    } else if (bytes.length >= 2 && ((bytes[0] === 0xFF && bytes[1] === 0xFE) || (bytes[0] === 0xFE && bytes[1] === 0xFF))) {
      // UTF-16 con BOM: lo que el Bloc de notas llama "Unicode".
      conBom = true;
      codificacion = bytes[0] === 0xFF ? 'utf-16le' : 'utf-16be';
      texto = new TextDecoder(codificacion).decode(bytes.slice(2));
    }
    if (texto === undefined) {
      const utf16 = pareceUtf16(bytes);
      if (utf16) {
        codificacion = utf16;
        utf16SinBom = true;
        texto = new TextDecoder(utf16).decode(bytes);
      } else {
        try {
          texto = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
        } catch (errorUtf8) {
          texto = new TextDecoder('windows-1252').decode(bytes);
          codificacion = 'windows-1252';
        }
      }
    }
    let mojibake = 0;
    const sinMojibake = corregirMojibake(texto);
    if (sinMojibake !== null) { mojibake = contarMojibake(texto); texto = sinMojibake; }
    const saltosMixtos = /\r\n/.test(texto) && /[^\r]\n|\r(?!\n)/.test(texto.replace(/\r\n/g, ''));
    return { texto: texto.replace(/\r\n/g, '\n').replace(/\r/g, '\n'), codificacion, conBom, utf16SinBom, mojibake, saltosMixtos };
  }

  window.TextoRetoques = { ordenar, quitarLineas, encabezadoPie, arreglarCodificacion };
})();

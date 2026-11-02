/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Texto de derecha a izquierda (Fase 6, etapa 6.4b). Sin DOM: corre igual en Node.
//
// Una versión práctica del algoritmo bidireccional de Unicode (UAX #9) para un párrafo de Word: hebreo, árabe, números dentro de ellos y signos.
// No lleva incrustaciones ni aislamientos explícitos (los controles U+202A a U+2069 se tratan como neutros): Word casi nunca los escribe, y la
// dirección de un párrafo la da su propiedad «de derecha a izquierda» (w:bidi), no el primer carácter fuerte.
//
//   Bidi.niveles(texto, baseDerecha)  → Uint8Array con el nivel de cada unidad UTF-16, o null si todo es de izquierda a derecha en un párrafo
//                                        de izquierda a derecha (no hay nada que reordenar)
//   Bidi.ordenVisual(niveles)         → índices de los elementos de una línea en orden visual (de izquierda a derecha de la pantalla)
//   Bidi.esDerecha(nivel)             → nivel impar
//
// Reglas: W1 a W7 (marcas, números tras letras árabes, separadores y terminadores de números), N1 y N2 (neutros entre letras de la misma
// dirección toman esa dirección; si no, la del párrafo), I1 e I2 (niveles), y L2 (invertir de los niveles más altos a los más bajos). L1 (los
// espacios del final de la línea y los tabuladores van a la dirección del párrafo) la aplica quien corta las líneas.
(function () {
  'use strict';
  const entre = (c, a, b) => c >= a && c <= b;
  const c16 = String.fromCharCode;
  // los bloques donde puede haber letras de derecha a izquierda: hebreo, árabe, sirio, thaana, nko, samaritano…, formas de presentación, y los planos altos
  const POSIBLE = new RegExp('[' + c16(0x590) + '-' + c16(0x8FF) + c16(0xFB1D) + '-' + c16(0xFDFF) + c16(0xFE70) + '-' + c16(0xFEFF) + c16(0x200F) + c16(0xD802) + c16(0xD803) + c16(0xD83A) + c16(0xD83B) + ']');
  const MARCA = /\p{Mn}/u, LETRA = /\p{L}/u, ESPACIO = /\s/u;

  // Tipo bidireccional de un punto de código (los que importan)
  function tipoDe(c, s) {
    if (c >= 0x30 && c <= 0x39) return 'EN';
    if (entre(c, 0x6F0, 0x6F9)) return 'EN';
    if (entre(c, 0x660, 0x669) || c === 0x66B || c === 0x66C) return 'AN';
    if (c === 0x9) return 'S';
    if (c === 0xA || c === 0xD || c === 0x2028 || c === 0x2029 || c === 0xC) return 'B';
    if (c === 0x2B || c === 0x2D) return 'ES';
    if (c === 0x23 || c === 0x24 || c === 0x25 || entre(c, 0xA2, 0xA5) || c === 0xB0 || c === 0xB1 || c === 0x66A || c === 0x2030 || c === 0x2031 || entre(c, 0x20A0, 0x20CF)) return 'ET';
    if (c === 0x2C || c === 0x2E || c === 0x2F || c === 0x3A || c === 0xA0 || c === 0x60C) return 'CS';
    if (entre(c, 0x590, 0x5FF) || entre(c, 0x7C0, 0x85F) || entre(c, 0xFB1D, 0xFB4F) || entre(c, 0x10800, 0x10FFF) || entre(c, 0x1E800, 0x1EFFF)) return MARCA.test(s) ? 'NSM' : 'R';
    if (entre(c, 0x600, 0x6FF) || entre(c, 0x750, 0x77F) || entre(c, 0x8A0, 0x8FF) || entre(c, 0xFB50, 0xFDFF) || entre(c, 0xFE70, 0xFEFF) || entre(c, 0x1EE00, 0x1EEFF)) return MARCA.test(s) ? 'NSM' : 'AL';
    if (c === 0x200F) return 'R';
    if (c === 0x200E) return 'L';
    if (MARCA.test(s)) return 'NSM';
    if (LETRA.test(s)) return 'L';
    if (ESPACIO.test(s)) return 'WS';
    return 'ON';
  }
  const esNI = (t) => t === 'B' || t === 'S' || t === 'WS' || t === 'ON';
  const fuerte = (t) => (t === 'L' ? 'L' : t === 'R' || t === 'EN' || t === 'AN' ? 'R' : null);

  function niveles(texto, baseDerecha) {
    const n = texto.length;
    if (!n) return null;
    // camino rápido: un párrafo de izquierda a derecha sin ningún carácter de escritura de derecha a izquierda (casi todos) no se analiza
    if (!baseDerecha && !POSIBLE.test(texto)) return null;
    const base = baseDerecha ? 1 : 0, sos = base ? 'R' : 'L';
    const t = new Array(n);
    let hayDerecha = false;
    for (let i = 0; i < n; i++) {
      const c = texto.codePointAt(i);
      const par = c > 0xFFFF;
      t[i] = tipoDe(c, par ? texto.slice(i, i + 2) : texto[i]);
      if (t[i] === 'R' || t[i] === 'AL' || t[i] === 'AN') hayDerecha = true;
      if (par) { t[i + 1] = t[i]; i++; }
    }
    if (!base && !hayDerecha) return null;
    // W1: una marca toma el tipo del carácter anterior
    let previo = sos;
    for (let i = 0; i < n; i++) { if (t[i] === 'NSM') t[i] = previo; else previo = t[i]; }
    // W2: un número europeo tras una letra árabe es árabe; W3: AL pasa a R
    let ultimo = sos;
    for (let i = 0; i < n; i++) {
      if (t[i] === 'L' || t[i] === 'R' || t[i] === 'AL') ultimo = t[i];
      else if (t[i] === 'EN' && ultimo === 'AL') t[i] = 'AN';
    }
    for (let i = 0; i < n; i++) if (t[i] === 'AL') t[i] = 'R';
    // W4: un separador solo entre dos números del mismo tipo
    for (let i = 1; i < n - 1; i++) {
      if (t[i] === 'ES' && t[i - 1] === 'EN' && t[i + 1] === 'EN') t[i] = 'EN';
      else if (t[i] === 'CS' && (t[i - 1] === 'EN' || t[i - 1] === 'AN') && t[i - 1] === t[i + 1]) t[i] = t[i - 1];
    }
    // W5: una serie de terminadores (%, $…) junto a un número europeo es parte de él
    for (let i = 0; i < n; i++) {
      if (t[i] !== 'ET') continue;
      let j = i; while (j < n && t[j] === 'ET') j++;
      if ((i > 0 && t[i - 1] === 'EN') || (j < n && t[j] === 'EN')) for (let k = i; k < j; k++) t[k] = 'EN';
      i = j - 1;
    }
    // W6: los demás separadores y terminadores son neutros
    for (let i = 0; i < n; i++) if (t[i] === 'ES' || t[i] === 'ET' || t[i] === 'CS') t[i] = 'ON';
    // W7: un número europeo tras una letra de izquierda a derecha es de izquierda a derecha
    ultimo = sos;
    for (let i = 0; i < n; i++) {
      if (t[i] === 'L' || t[i] === 'R') ultimo = t[i];
      else if (t[i] === 'EN' && ultimo === 'L') t[i] = 'L';
    }
    // N1 y N2: los neutros
    for (let i = 0; i < n; i++) {
      if (!esNI(t[i])) continue;
      let j = i; while (j < n && esNI(t[j])) j++;
      const antes = i === 0 ? sos : fuerte(t[i - 1]), despues = j >= n ? sos : fuerte(t[j]);
      const dir = antes && antes === despues ? antes : sos;
      for (let k = i; k < j; k++) t[k] = dir;
      i = j - 1;
    }
    // I1 e I2
    const sal = new Uint8Array(n);
    for (let i = 0; i < n; i++) {
      const x = t[i];
      sal[i] = base === 0 ? (x === 'L' ? 0 : x === 'R' ? 1 : 2) : (x === 'R' ? 1 : 2);
    }
    return sal;
  }

  // L2: el orden en que se ven, de izquierda a derecha, elementos con esos niveles
  function ordenVisual(niv) {
    const n = niv.length, idx = Array.from({ length: n }, (_, i) => i);
    let max = 0, minImpar = 255;
    for (const l of niv) { if (l > max) max = l; if ((l & 1) && l < minImpar) minImpar = l; }
    for (let nivel = max; nivel >= minImpar && nivel > 0; nivel--) {
      let i = 0;
      while (i < n) {
        if (niv[idx[i]] >= nivel) {
          let j = i; while (j < n && niv[idx[j]] >= nivel) j++;
          for (let a = i, b = j - 1; a < b; a++, b--) { const x = idx[a]; idx[a] = idx[b]; idx[b] = x; }
          i = j;
        } else i++;
      }
    }
    return idx;
  }

  globalThis.Bidi = { niveles, ordenVisual, esDerecha: (n) => (n & 1) === 1 };
})();

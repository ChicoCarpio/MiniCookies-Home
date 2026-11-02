/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Comparador de textos propio (reemplaza a jsdiff). Las herramientas
// "Comparar textos" y "Comparar Word" lo usan para marcar lo agregado y lo
// quitado:
//   const cambios = DiffPropio.diffWords(textoA, textoB);   // por palabra
//   const cambios = DiffPropio.diffLines(textoA, textoB);   // por línea
// Devuelve una lista de { value, added, removed, count } (count = cantidad de
// piezas del tramo). Es el algoritmo de Myers con el mismo recorrido que
// jsdiff 9, así que da el mismo resultado, incluida la forma de tratar los
// espacios: en diffWords un espacio pegado a un cambio no se marca como
// cambio si es igual en los dos textos.
//
// Solo está lo que el sitio usa: sin opciones, sin callback y sin segmentador
// de Intl.
//
// Protección que jsdiff no tenía: si dos textos son tan distintos que harían
// falta más de TOPE cambios, el cálculo exacto crece con el cuadrado de esa
// cantidad. Pasado el tope se devuelve el resultado grueso (lo común del
// principio y del final, y en medio todo quitado y todo agregado). Con textos
// de tamaño normal nunca se llega al tope.
(function () {
  'use strict';

  const TOPE = 5000;

  // Letras de "palabra": las de jsdiff (latinas con acentos, dígitos y _).
  const CARACTERES_PALABRA = 'a-zA-Z0-9_\\u{AD}\\u{C0}-\\u{D6}\\u{D8}-\\u{F6}\\u{F8}-\\u{2C6}\\u{2C8}-\\u{2D7}\\u{2DE}-\\u{2FF}\\u{1E00}-\\u{1EFF}';
  const RE_PIEZAS = new RegExp('[' + CARACTERES_PALABRA + ']+|\\s+|[^' + CARACTERES_PALABRA + ']', 'ug');
  const RE_ESPACIO = /\s/;

  // ---------- Myers ----------

  function agregarACamino(camino, agregado, quitado, avanceViejo) {
    const ultimo = camino.lastComponent;
    if (ultimo && ultimo.added === agregado && ultimo.removed === quitado) {
      return { oldPos: camino.oldPos + avanceViejo, lastComponent: { count: ultimo.count + 1, added: agregado, removed: quitado, previousComponent: ultimo.previousComponent } };
    }
    return { oldPos: camino.oldPos + avanceViejo, lastComponent: { count: 1, added: agregado, removed: quitado, previousComponent: ultimo } };
  }

  // Avanza por la diagonal mientras las piezas sean iguales.
  function extraerComun(camino, nuevas, viejas, diagonal, iguales) {
    const nLen = nuevas.length, vLen = viejas.length;
    let vPos = camino.oldPos, nPos = vPos - diagonal, comunes = 0;
    while (nPos + 1 < nLen && vPos + 1 < vLen && iguales(viejas[vPos + 1], nuevas[nPos + 1])) {
      nPos++; vPos++; comunes++;
    }
    if (comunes) camino.lastComponent = { count: comunes, previousComponent: camino.lastComponent, added: false, removed: false };
    camino.oldPos = vPos;
    return nPos;
  }

  // Arma la lista final (sin value todavía) a partir de la cadena de tramos.
  function desenlazar(ultimo) {
    const tramos = [];
    while (ultimo) {
      tramos.push(ultimo);
      const siguiente = ultimo.previousComponent;
      delete ultimo.previousComponent;
      ultimo = siguiente;
    }
    return tramos.reverse();
  }

  function ponerValores(tramos, nuevas, viejas, unir) {
    let nPos = 0, vPos = 0;
    for (const t of tramos) {
      if (t.removed) {
        t.value = unir(viejas.slice(vPos, vPos + t.count));
        vPos += t.count;
      } else {
        t.value = unir(nuevas.slice(nPos, nPos + t.count));
        nPos += t.count;
        if (!t.added) vPos += t.count;
      }
    }
    return tramos;
  }

  // Resultado grueso para cuando se pasa el tope.
  function resultadoGrueso(nuevas, viejas, iguales) {
    const nLen = nuevas.length, vLen = viejas.length;
    let pre = 0;
    while (pre < nLen && pre < vLen && iguales(viejas[pre], nuevas[pre])) pre++;
    let suf = 0;
    while (suf < nLen - pre && suf < vLen - pre && iguales(viejas[vLen - 1 - suf], nuevas[nLen - 1 - suf])) suf++;
    const tramos = [];
    if (pre) tramos.push({ count: pre, added: false, removed: false });
    if (vLen - pre - suf > 0) tramos.push({ count: vLen - pre - suf, added: false, removed: true });
    if (nLen - pre - suf > 0) tramos.push({ count: nLen - pre - suf, added: true, removed: false });
    if (suf) tramos.push({ count: suf, added: false, removed: false });
    return tramos;
  }

  function myers(viejas, nuevas, iguales, unir) {
    const nLen = nuevas.length, vLen = viejas.length;
    const maxEdicion = Math.min(nLen + vLen, DiffPropio.tope);
    const mejor = [{ oldPos: -1, lastComponent: undefined }];
    let nPos = extraerComun(mejor[0], nuevas, viejas, 0, iguales);
    if (mejor[0].oldPos + 1 >= vLen && nPos + 1 >= nLen) {
      return ponerValores(desenlazar(mejor[0].lastComponent), nuevas, viejas, unir);
    }
    let diagMin = -Infinity, diagMax = Infinity;
    for (let edicion = 1; edicion <= maxEdicion; edicion++) {
      for (let diag = Math.max(diagMin, -edicion); diag <= Math.min(diagMax, edicion); diag += 2) {
        let base;
        const quitar = mejor[diag - 1], agregar = mejor[diag + 1];
        if (quitar) mejor[diag - 1] = undefined;
        let puedeAgregar = false;
        if (agregar) {
          const nuevaPos = agregar.oldPos - diag;
          puedeAgregar = nuevaPos >= 0 && nuevaPos < nLen;
        }
        const puedeQuitar = quitar && quitar.oldPos + 1 < vLen;
        if (!puedeAgregar && !puedeQuitar) {
          mejor[diag] = undefined;
          continue;
        }
        if (!puedeQuitar || (puedeAgregar && quitar.oldPos < agregar.oldPos)) base = agregarACamino(agregar, true, false, 0);
        else base = agregarACamino(quitar, false, true, 1);
        nPos = extraerComun(base, nuevas, viejas, diag, iguales);
        if (base.oldPos + 1 >= vLen && nPos + 1 >= nLen) {
          return ponerValores(desenlazar(base.lastComponent), nuevas, viejas, unir);
        }
        mejor[diag] = base;
        if (base.oldPos + 1 >= vLen) diagMax = Math.min(diagMax, diag - 1);
        if (nPos + 1 >= nLen) diagMin = Math.max(diagMin, diag + 1);
      }
    }
    return ponerValores(resultadoGrueso(nuevas, viejas, iguales), nuevas, viejas, unir);
  }

  // ---------- Por línea ----------

  // Cada pieza es una línea con su salto de línea (\n o \r\n) incluido.
  function piezasLineas(texto) {
    const partes = texto.split(/(\n|\r\n)/);
    if (!partes[partes.length - 1]) partes.pop();
    const lineas = [];
    for (let i = 0; i < partes.length; i++) {
      if (i % 2) lineas[lineas.length - 1] += partes[i];
      else lineas.push(partes[i]);
    }
    return lineas.filter(Boolean);
  }

  function diffLines(textoViejo, textoNuevo) {
    return myers(piezasLineas(textoViejo), piezasLineas(textoNuevo), (a, b) => a === b, (p) => p.join(''));
  }

  // ---------- Por palabra ----------

  // Palabras, signos y espacios; el espacio queda pegado a la palabra que lo
  // sigue, y los espacios del principio forman su propia pieza.
  function piezasPalabras(texto) {
    const partes = texto.match(RE_PIEZAS) || [];
    const piezas = [];
    let anterior = null;
    for (const parte of partes) {
      if (RE_ESPACIO.test(parte)) {
        if (anterior == null) piezas.push(parte);
        else piezas.push(piezas.pop() + parte);
      } else if (anterior != null && RE_ESPACIO.test(anterior)) {
        if (piezas[piezas.length - 1] == anterior) piezas.push(piezas.pop() + parte);
        else piezas.push(anterior + parte);
      } else {
        piezas.push(parte);
      }
      anterior = parte;
    }
    return piezas.filter(Boolean);
  }

  function espaciosIniciales(texto) {
    const m = texto.match(/^\s*/);
    return m ? m[0] : '';
  }
  function espaciosFinales(texto) {
    let i = texto.length - 1;
    while (i >= 0 && texto[i].match(/\s/)) i--;
    return texto.substring(i + 1);
  }
  function prefijoComun(a, b) {
    let i = 0;
    while (i < a.length && i < b.length && a[i] == b[i]) i++;
    return a.slice(0, i);
  }
  function sufijoComun(a, b) {
    if (!a || !b || a[a.length - 1] != b[b.length - 1]) return '';
    let i = 0;
    while (i < a.length && i < b.length) {
      if (a[a.length - (i + 1)] != b[b.length - (i + 1)]) return a.slice(-i);
      i++;
    }
    return a.slice(-i);
  }
  function cambiarPrefijo(texto, viejo, nuevo) {
    if (texto.slice(0, viejo.length) != viejo) throw new Error('El texto no empieza con el prefijo esperado');
    return nuevo + texto.slice(viejo.length);
  }
  function cambiarSufijo(texto, viejo, nuevo) {
    if (!viejo) return texto + nuevo;
    if (texto.slice(-viejo.length) != viejo) throw new Error('El texto no termina con el sufijo esperado');
    return texto.slice(0, -viejo.length) + nuevo;
  }
  function quitarPrefijo(texto, prefijo) { return cambiarPrefijo(texto, prefijo, ''); }
  function quitarSufijo(texto, sufijo) { return cambiarSufijo(texto, sufijo, ''); }

  // Mayor tramo en que el final de a coincide con el principio de b (KMP).
  function mayorSolape(a, b) {
    let inicioA = 0, finB = b.length;
    if (a.length > b.length) inicioA = a.length - b.length;
    if (a.length < b.length) finB = a.length;
    const tabla = Array(finB);
    let k = 0;
    tabla[0] = 0;
    for (let j = 1; j < finB; j++) {
      if (b[j] == b[k]) tabla[j] = tabla[k];
      else tabla[j] = k;
      while (k > 0 && b[j] != b[k]) k = tabla[k];
      if (b[j] == b[k]) k++;
    }
    k = 0;
    for (let i = inicioA; i < a.length; i++) {
      while (k > 0 && a[i] != b[k]) k = tabla[k];
      if (a[i] == b[k]) k++;
    }
    return b.slice(0, k);
  }

  // Después de comparar, los espacios que rodean un cambio y son iguales en
  // los dos lados se devuelven al tramo sin cambios.
  function limpiarEspacios(tramoAntes, quitado, agregado, tramoDespues) {
    if (quitado && agregado) {
      const viejoIni = espaciosIniciales(quitado.value), viejoFin = espaciosFinales(quitado.value);
      const nuevoIni = espaciosIniciales(agregado.value), nuevoFin = espaciosFinales(agregado.value);
      if (tramoAntes) {
        const comun = prefijoComun(viejoIni, nuevoIni);
        tramoAntes.value = cambiarSufijo(tramoAntes.value, nuevoIni, comun);
        quitado.value = quitarPrefijo(quitado.value, comun);
        agregado.value = quitarPrefijo(agregado.value, comun);
      }
      if (tramoDespues) {
        const comun = sufijoComun(viejoFin, nuevoFin);
        tramoDespues.value = cambiarPrefijo(tramoDespues.value, nuevoFin, comun);
        quitado.value = quitarSufijo(quitado.value, comun);
        agregado.value = quitarSufijo(agregado.value, comun);
      }
    } else if (agregado) {
      if (tramoAntes) agregado.value = agregado.value.substring(espaciosIniciales(agregado.value).length);
      if (tramoDespues) tramoDespues.value = tramoDespues.value.substring(espaciosIniciales(tramoDespues.value).length);
    } else if (tramoAntes && tramoDespues) {
      const espDespues = espaciosIniciales(tramoDespues.value);
      const iniQuitado = espaciosIniciales(quitado.value), finQuitado = espaciosFinales(quitado.value);
      const delIni = prefijoComun(espDespues, iniQuitado);
      quitado.value = quitarPrefijo(quitado.value, delIni);
      const delFin = sufijoComun(quitarPrefijo(espDespues, delIni), finQuitado);
      quitado.value = quitarSufijo(quitado.value, delFin);
      tramoDespues.value = cambiarPrefijo(tramoDespues.value, espDespues, delFin);
      tramoAntes.value = cambiarSufijo(tramoAntes.value, espDespues, espDespues.slice(0, espDespues.length - delFin.length));
    } else if (tramoDespues) {
      const solape = mayorSolape(espaciosFinales(quitado.value), espaciosIniciales(tramoDespues.value));
      quitado.value = quitarSufijo(quitado.value, solape);
    } else if (tramoAntes) {
      const solape = mayorSolape(espaciosFinales(tramoAntes.value), espaciosIniciales(quitado.value));
      quitado.value = quitarPrefijo(quitado.value, solape);
    }
  }

  function diffWords(textoViejo, textoNuevo) {
    const cambios = myers(
      piezasPalabras(textoViejo), piezasPalabras(textoNuevo),
      // Dos piezas son iguales aunque difieran en los espacios de los bordes.
      (a, b) => a === b || a.trim() === b.trim(),
      // Al unir, se descartan los espacios del principio de cada pieza menos la primera.
      (piezas) => piezas.map((p, i) => (i == 0 ? p : p.replace(/^\s+/, ''))).join('')
    );
    let ultimoIgual = null, agregado = null, quitado = null;
    for (const cambio of cambios) {
      if (cambio.added) agregado = cambio;
      else if (cambio.removed) quitado = cambio;
      else {
        if (agregado || quitado) limpiarEspacios(ultimoIgual, quitado, agregado, cambio);
        ultimoIgual = cambio;
        agregado = null;
        quitado = null;
      }
    }
    if (agregado || quitado) limpiarEspacios(ultimoIgual, quitado, agregado, null);
    return cambios;
  }

  const DiffPropio = { diffWords, diffLines, tope: TOPE };
  globalThis.DiffPropio = DiffPropio;
})();

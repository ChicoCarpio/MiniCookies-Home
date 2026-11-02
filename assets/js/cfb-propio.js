/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Contenedor CFB propio ("Compound File Binary", [MS-CFB]): el sistema de archivos dentro de los .xls, .doc y .ppt de Office 97-2003.
// Un archivo CFB es una lista de sectores; una tabla (FAT) dice qué sector sigue a cuál, y un directorio da nombre a cada flujo.
// Los flujos de menos de 4096 bytes van en el "miniflujo" (sectores de 64 bytes, con su propia tabla, la MiniFAT). No depende del navegador.
//
//   const cfb = CfbPropio.leer(bytes, { danado: (detalle) => new Error(detalle) });   // danado es opcional: arma el error de un archivo roto
//   cfb.entradas   // [{ id, nombre, tipo: 1 almacén | 2 flujo | 5 raíz, inicio, tam, izq, der, hijo }]  (las entradas vacías no están)
//   cfb.flujo(entrada)   // Uint8Array con el contenido de un flujo
//
//   const bytes = CfbPropio.escribir([{ nombre: 'Workbook', bytes }, { nombre: 'ObjectPool', hijos: [...] }]);   // Uint8Array
//
// El escritor arma la versión 3 (sectores de 512 bytes, la que escribe Office 97-2003 y la que todo programa lee): FAT, DIFAT si la FAT pasa de
// 109 sectores, directorio como árbol rojo-negro válido, y miniflujo para los flujos chicos. Fechas y CLSID en cero.
(function () {
  const u16 = (b, p) => b[p] | (b[p + 1] << 8);
  const u32 = (b, p) => (b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24)) >>> 0;
  const FIN = 0xFFFFFFFE, LIBRE = 0xFFFFFFFF, SECT_FAT = 0xFFFFFFFD, SECT_DIFAT = 0xFFFFFFFC, SIN_ENTRADA = 0xFFFFFFFF;
  const FIRMA = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

  // ---- lector -------------------------------------------------------------------------------------------------------
  function leer(b, opciones) {
    const danado = (opciones && opciones.danado) || ((detalle) => Object.assign(new Error('CFB dañado: ' + detalle), { codigo: 'danado', detalle }));
    if (b.length < 512) throw danado('contenedor muy corto');
    const desp = u16(b, 0x1E), despMini = u16(b, 0x20);
    if (desp < 7 || desp > 16 || despMini > desp) throw danado('tamaño de sector');
    const tam = 1 << desp, tamMini = 1 << despMini;
    const nSect = Math.floor((b.length - 512) / tam) + 1;
    const sector = (id) => {
      const ini = (id + 1) * tam;
      if (id >= nSect + 1 || ini >= b.length) throw danado('sector fuera del archivo');
      return b.subarray(ini, Math.min(ini + tam, b.length));
    };
    // tabla de asignación (FAT): sus sectores los dice el encabezado y, si son más de 109, la cadena DIFAT. No puede haber más sectores de FAT que
    // sectores en el archivo: una DIFAT hostil que repite el mismo id armaba una FAT de gigas con un archivo de 10 MB.
    const idsFat = [];
    for (let k = 0; k < 109; k++) { const id = u32(b, 0x4C + k * 4); if (id < 0xFFFFFFFA) idsFat.push(id); }
    let dif = u32(b, 0x44), vueltas = 0;
    while (dif < 0xFFFFFFFA && vueltas++ < nSect && idsFat.length <= nSect) {
      const s = sector(dif);
      for (let k = 0; k < tam / 4 - 1; k++) { const id = u32(s, k * 4); if (id < 0xFFFFFFFA) idsFat.push(id); }
      dif = u32(s, tam - 4);
    }
    if (idsFat.length > nSect) throw danado('FAT más grande que el archivo');
    const fat = new Uint32Array(idsFat.length * (tam / 4));
    idsFat.forEach((id, k) => { const s = sector(id); for (let j = 0; j < tam / 4; j++) fat[k * (tam / 4) + j] = u32(s, j * 4); });
    // una cadena no pasa por más sectores que los que tiene el archivo (con un ciclo, la tabla entera: un flujo de cientos de megas)
    const cadena = (inicio, tabla, max) => {
      const ids = [];
      let id = inicio;
      while (id < 0xFFFFFFFA) {
        if (ids.length >= max || id >= tabla.length) throw danado('cadena de sectores');
        ids.push(id); id = tabla[id];
      }
      return ids;
    };
    const leerCadena = (inicio, tamano) => {
      const ids = cadena(inicio, fat, nSect), r = new Uint8Array(ids.length * tam);
      ids.forEach((id, k) => r.set(sector(id), k * tam));
      return tamano === undefined ? r : r.subarray(0, tamano);
    };
    // directorio
    const dir = leerCadena(u32(b, 0x30));
    // el nombre ocupa a lo sumo los 64 bytes de su campo: un largo mayor (archivo hostil o mal escrito) leía los campos siguientes como texto
    const entradas = [], utf16 = new TextDecoder('utf-16le');
    for (let p = 0; p + 128 <= dir.length; p += 128) {
      const nc = Math.min(u16(dir, p + 64), 64);
      if (!dir[p + 66]) continue;
      entradas.push({ id: p / 128, nombre: utf16.decode(dir.subarray(p, p + Math.max(0, nc - 2))).split('\0')[0], tipo: dir[p + 66], inicio: u32(dir, p + 116), tam: u32(dir, p + 120),
        izq: u32(dir, p + 68), der: u32(dir, p + 72), hijo: u32(dir, p + 76) });
    }
    const raiz = entradas.find((e) => e.tipo === 5);
    const corte = u32(b, 0x38);
    let miniFat = null, miniFlujo = null;
    const flujo = (e) => {
      if (e.tam < corte && raiz) {
        if (!miniFat) {
          const bytes = leerCadena(u32(b, 0x3C));
          miniFat = new Uint32Array(bytes.length / 4);
          for (let k = 0; k < miniFat.length; k++) miniFat[k] = u32(bytes, k * 4);
          miniFlujo = leerCadena(raiz.inicio, raiz.tam);
        }
        const ids = cadena(e.inicio, miniFat, Math.ceil(miniFlujo.length / tamMini)), r = new Uint8Array(ids.length * tamMini);
        ids.forEach((id, k) => { const ini = id * tamMini; if (ini + tamMini > miniFlujo.length + tamMini) throw danado('minisector'); r.set(miniFlujo.subarray(ini, ini + tamMini), k * tamMini); });
        return r.subarray(0, e.tam);
      }
      return leerCadena(e.inicio, e.tam);
    };
    return { entradas, flujo };
  }

  const esCfb = (b) => b && b.length >= 8 && FIRMA.every((x, i) => b[i] === x);

  // ---- escritor -----------------------------------------------------------------------------------------------------
  const TAM = 512, TAM_MINI = 64, CORTE = 4096, POR_SECTOR = TAM / 4;
  const NOMBRE_PROHIBIDO = /[\/\\:!]/;

  // Orden de los hermanos en el directorio ([MS-CFB] 2.6.4): primero el nombre más corto; a igual largo, carácter por carácter en mayúsculas.
  // La norma pasa a mayúsculas cada unidad UTF-16 por separado (conversión simple) y no toca los sustitutos: toUpperCase() sobre el texto entero
  // no sirve, porque 'ß' pasa a 'SS' (cambia el largo y el orden) y un par sustituto sí cambia. Si la mayúscula de una unidad no es una sola, queda igual.
  const mayusculas = (s) => { let r = ''; for (let i = 0; i < s.length; i++) { const c = s[i], m = c.toUpperCase(); r += m.length === 1 ? m : c; } return r; };
  function compararNombres(a, b) {
    if (a.length !== b.length) return a.length - b.length;
    const A = mayusculas(a), B = mayusculas(b);
    return A < B ? -1 : A > B ? 1 : 0;
  }

  function escribir(lista, opciones) {
    const nodos = [];   // [{ nombre, tipo, bytes, hijos: [índices], izq, der, hijo, color, inicio }]
    const raiz = { nombre: 'Root Entry', tipo: 5, hijos: [] };
    nodos.push(raiz);
    const agregar = (padre, items) => {
      const vistos = new Set();
      for (const it of items) {
        if (!it || typeof it.nombre !== 'string' || !it.nombre || it.nombre.length > 31 || NOMBRE_PROHIBIDO.test(it.nombre)) throw new Error('CFB: nombre de entrada no válido: ' + (it && it.nombre));
        const clave = mayusculas(it.nombre);
        if (vistos.has(clave)) throw new Error('CFB: nombre repetido: ' + it.nombre);
        vistos.add(clave);
        const n = it.hijos ? { nombre: it.nombre, tipo: 1, hijos: [] } : { nombre: it.nombre, tipo: 2, bytes: it.bytes instanceof Uint8Array ? it.bytes : new Uint8Array(it.bytes || 0) };
        padre.hijos.push(nodos.length); nodos.push(n);
        if (it.hijos) agregar(n, it.hijos);
      }
    };
    agregar(raiz, lista || []);
    for (const n of nodos) { n.izq = SIN_ENTRADA; n.der = SIN_ENTRADA; n.hijo = SIN_ENTRADA; n.color = 1; }
    // Árbol de hermanos: binario de búsqueda balanceado (raíz en el medio de la lista ordenada). Todos los niveles quedan llenos salvo el último;
    // ese último nivel va en rojo y el resto en negro, así todo camino tiene los mismos negros y ningún rojo tiene hijo rojo: es rojo-negro válido.
    for (const padre of nodos) {
      if (!padre.hijos || !padre.hijos.length) continue;
      const orden = padre.hijos.slice().sort((x, y) => compararNombres(nodos[x].nombre, nodos[y].nombre));
      const alto = Math.floor(Math.log2(orden.length));
      const armar = (ini, fin, prof) => {
        if (ini > fin) return SIN_ENTRADA;
        const m = (ini + fin) >> 1, n = nodos[orden[m]];
        n.color = prof === alto && orden.length !== (1 << (alto + 1)) - 1 ? 0 : 1;
        n.izq = armar(ini, m - 1, prof + 1); n.der = armar(m + 1, fin, prof + 1);
        return orden[m];
      };
      padre.hijo = armar(0, orden.length - 1, 0);
    }

    // Miniflujo: los flujos de menos de 4096 bytes, uno detrás del otro en sectores de 64 bytes.
    const flujos = nodos.filter((n) => n.tipo === 2);
    const chicos = flujos.filter((n) => n.bytes.length > 0 && n.bytes.length < CORTE), grandes = flujos.filter((n) => n.bytes.length >= CORTE);
    const miniFat = [];
    for (const n of chicos) {
      const k = Math.ceil(n.bytes.length / TAM_MINI);
      n.inicio = miniFat.length;
      for (let i = 0; i < k; i++) miniFat.push(i === k - 1 ? FIN : miniFat.length + 1);
    }
    for (const n of flujos) if (n.bytes.length === 0) n.inicio = FIN;
    const tamMiniFlujo = miniFat.length * TAM_MINI;

    // Sectores, en este orden: directorio, MiniFAT, miniflujo, flujos grandes; después la FAT y la DIFAT (su tamaño depende del total).
    const sectores = (n) => Math.ceil(n / TAM);
    const nDir = sectores(nodos.length * 128), nMiniFat = sectores(miniFat.length * 4), nMini = sectores(tamMiniFlujo);
    let nDatos = nDir + nMiniFat + nMini;
    for (const n of grandes) nDatos += sectores(n.bytes.length);
    let nFat = 0, nDifat = 0;
    for (;;) {   // la FAT también se cuenta a sí misma y a la DIFAT
      const fat = Math.ceil((nDatos + nFat + nDifat) / POR_SECTOR);
      const difat = fat > 109 ? Math.ceil((fat - 109) / (POR_SECTOR - 1)) : 0;
      if (fat === nFat && difat === nDifat) break;
      nFat = fat; nDifat = difat;
    }
    const total = nDatos + nFat + nDifat;
    if (total > 0x7FFFFFF0 / TAM) throw new Error('CFB: archivo demasiado grande');
    const fat = new Uint32Array(nFat * POR_SECTOR).fill(LIBRE);
    let siguiente = 0;
    const reservar = (k) => {   // k sectores seguidos, encadenados en la FAT; devuelve el primero (o FIN si k = 0)
      if (!k) return FIN;
      const ini = siguiente;
      for (let i = 0; i < k; i++) fat[ini + i] = i === k - 1 ? FIN : ini + i + 1;
      siguiente += k;
      return ini;
    };
    const iniDir = reservar(nDir), iniMiniFat = reservar(nMiniFat), iniMini = reservar(nMini);
    for (const n of grandes) n.inicio = reservar(sectores(n.bytes.length));
    const iniFat = siguiente;
    for (let i = 0; i < nFat; i++) fat[siguiente++] = SECT_FAT;
    const iniDifat = siguiente;
    for (let i = 0; i < nDifat; i++) fat[siguiente++] = SECT_DIFAT;
    raiz.inicio = nMini ? iniMini : FIN;

    const salida = new Uint8Array(TAM * (total + 1));
    const v = new DataView(salida.buffer);
    const pos = (id) => TAM * (id + 1);
    // encabezado
    salida.set(FIRMA, 0);
    v.setUint16(0x18, 0x003E, true); v.setUint16(0x1A, 3, true); v.setUint16(0x1C, 0xFFFE, true);
    v.setUint16(0x1E, 9, true); v.setUint16(0x20, 6, true);
    v.setUint32(0x2C, nFat, true); v.setUint32(0x30, iniDir, true); v.setUint32(0x38, CORTE, true);
    v.setUint32(0x3C, nMiniFat ? iniMiniFat : FIN, true); v.setUint32(0x40, nMiniFat, true);
    v.setUint32(0x44, nDifat ? iniDifat : FIN, true); v.setUint32(0x48, nDifat, true);
    for (let k = 0; k < 109; k++) v.setUint32(0x4C + k * 4, k < nFat ? iniFat + k : LIBRE, true);
    // DIFAT: los ids de los sectores de FAT que no entran en el encabezado; el último entero de cada sector apunta al siguiente
    for (let d = 0; d < nDifat; d++) {
      const p = pos(iniDifat + d);
      for (let k = 0; k < POR_SECTOR - 1; k++) { const i = 109 + d * (POR_SECTOR - 1) + k; v.setUint32(p + k * 4, i < nFat ? iniFat + i : LIBRE, true); }
      v.setUint32(p + TAM - 4, d === nDifat - 1 ? FIN : iniDifat + d + 1, true);
    }
    for (let i = 0; i < fat.length; i++) v.setUint32(pos(iniFat) + i * 4, fat[i], true);   // los sectores de FAT son seguidos
    // directorio (las entradas sobrantes del último sector quedan vacías: tipo 0, sin hermanos)
    for (let i = 0; i < nDir * 4; i++) {
      const p = pos(iniDir) + i * 128;
      const n = nodos[i];
      if (!n) { v.setUint32(p + 68, SIN_ENTRADA, true); v.setUint32(p + 72, SIN_ENTRADA, true); v.setUint32(p + 76, SIN_ENTRADA, true); continue; }
      for (let k = 0; k < n.nombre.length; k++) v.setUint16(p + k * 2, n.nombre.charCodeAt(k), true);
      v.setUint16(p + 64, (n.nombre.length + 1) * 2, true);
      salida[p + 66] = n.tipo; salida[p + 67] = n.color;
      v.setUint32(p + 68, n.izq, true); v.setUint32(p + 72, n.der, true); v.setUint32(p + 76, n.hijo, true);
      if (n.tipo === 2) { v.setUint32(p + 116, n.inicio, true); v.setUint32(p + 120, n.bytes.length, true); }
      else if (n.tipo === 5) { v.setUint32(p + 116, raiz.inicio, true); v.setUint32(p + 120, tamMiniFlujo, true); }
      // almacén: inicio y tamaño en 0
    }
    for (let i = 0; i < nMiniFat * POR_SECTOR; i++) v.setUint32(pos(iniMiniFat) + i * 4, i < miniFat.length ? miniFat[i] : LIBRE, true);
    for (const n of chicos) salida.set(n.bytes, pos(iniMini) + n.inicio * TAM_MINI);   // el miniflujo ocupa sectores seguidos
    for (const n of grandes) salida.set(n.bytes, pos(n.inicio));                        // y cada flujo grande también
    return salida;
  }

  globalThis.CfbPropio = { leer, escribir, esCfb, _compararNombres: compararNombres };
})();

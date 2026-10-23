/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Lector XML por eventos, ligero y tolerante, para las partes grandes de un
// libro de Excel (hojas de millones de celdas): DOMParser construye un árbol
// entero en memoria y con una hoja grande tarda segundos y gasta cientos de MB.
//
// No valida ni resuelve espacios de nombres: usa el nombre local de elementos y
// atributos (`x:row` → `row`, `r:id` → `id`), que es lo que necesita OOXML
// (los generadores usan prefijos distintos para los mismos elementos). Entiende
// etiquetas, atributos con comillas simples o dobles, texto, CDATA, comentarios,
// instrucciones de proceso y las entidades predefinidas y numéricas. No depende
// del navegador: corre igual en Node (las pruebas).
//
// Uso:
//   XmlLigero.recorrer(texto, {
//     abrir(nombre, atributos, vacio) { … },   // vacio = <x/>; igualmente llega cerrar() justo después
//     cerrar(nombre) { … },
//     texto(cadena) { … },                      // ya con las entidades resueltas
//   });
//   XmlLigero.sinEscapes(cadena)  → resuelve los "_x000D_" que Excel usa para caracteres de control

(function () {
  const ENTIDADES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

  function decodificar(s) {
    if (s.indexOf('&') < 0) return s;
    return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g, (todo, e) => {
      if (e.charCodeAt(0) === 35) {
        const codigo = e.charCodeAt(1) === 120 || e.charCodeAt(1) === 88 ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
        if (!(codigo >= 0 && codigo <= 0x10FFFF)) return todo;
        try { return String.fromCodePoint(codigo); } catch (error) { return todo; }
      }
      return Object.prototype.hasOwnProperty.call(ENTIDADES, e) ? ENTIDADES[e] : todo;
    });
  }

  // Excel escribe los caracteres de control (y el propio "_x") como _xHHHH_.
  function sinEscapes(s) {
    if (s.indexOf('_x') < 0) return s;
    return s.replace(/_x([0-9a-fA-F]{4})_/g, (todo, h) => String.fromCharCode(parseInt(h, 16)));
  }

  function local(nombre) {
    const dos = nombre.indexOf(':');
    return dos < 0 ? nombre : nombre.slice(dos + 1);
  }

  const esEsp = (k) => k === 32 || k === 10 || k === 13 || k === 9;

  function recorrer(texto, h) {
    const n = texto.length;
    const pila = [];
    let i = 0;
    if (n && texto.charCodeAt(0) === 0xFEFF) i = 1; // BOM
    while (i < n) {
      const lt = texto.indexOf('<', i);
      if (lt < 0) break;
      if (lt > i && h.texto && pila.length) {
        let t = texto.slice(i, lt);
        if (t.indexOf('\r') >= 0) t = t.replace(/\r\n?/g, '\n'); // la norma XML pide normalizar los saltos de línea
        h.texto(t.indexOf('&') < 0 ? t : decodificar(t));
      }
      const c = texto.charCodeAt(lt + 1);
      if (c === 47) { // </nombre>
        const gt = texto.indexOf('>', lt + 2);
        if (gt < 0) break;
        const nombre = pila.pop();
        if (h.cerrar && nombre !== undefined) h.cerrar(nombre);
        i = gt + 1;
        continue;
      }
      if (c === 33) { // <!-- -->, <![CDATA[ ]]>, <!DOCTYPE …>
        if (texto.startsWith('<!--', lt)) {
          const fin = texto.indexOf('-->', lt + 4);
          i = fin < 0 ? n : fin + 3;
        } else if (texto.startsWith('<![CDATA[', lt)) {
          const fin = texto.indexOf(']]>', lt + 9);
          const dato = texto.slice(lt + 9, fin < 0 ? n : fin);
          if (h.texto && pila.length) h.texto(dato);
          i = fin < 0 ? n : fin + 3;
        } else {
          const gt = texto.indexOf('>', lt);
          i = gt < 0 ? n : gt + 1;
        }
        continue;
      }
      if (c === 63) { // <?xml … ?>
        const fin = texto.indexOf('?>', lt + 2);
        i = fin < 0 ? n : fin + 2;
        continue;
      }
      // <nombre atributos… > o <nombre atributos… />
      let p = lt + 1;
      while (p < n) {
        const k = texto.charCodeAt(p);
        if (k === 62 || k === 47 || esEsp(texto.charCodeAt(p))) break;
        p++;
      }
      const nombre = local(texto.slice(lt + 1, p));
      const atributos = {};
      let vacio = false;
      for (;;) {
        while (p < n && esEsp(texto.charCodeAt(p))) p++;
        if (p >= n) break;
        const k = texto.charCodeAt(p);
        if (k === 62) { p++; break; }
        if (k === 47) { vacio = true; p = texto.indexOf('>', p); p = p < 0 ? n : p + 1; break; }
        const ini = p;
        while (p < n) {
          const q = texto.charCodeAt(p);
          if (q === 61 || q === 62 || q === 47 || esEsp(texto.charCodeAt(p))) break;
          p++;
        }
        const clave = texto.slice(ini, p);
        while (p < n && esEsp(texto.charCodeAt(p))) p++;
        if (texto.charCodeAt(p) !== 61) { atributos[local(clave)] = ''; continue; } // atributo sin valor
        p++;
        while (p < n && esEsp(texto.charCodeAt(p))) p++;
        const comilla = texto[p];
        if (comilla !== '"' && comilla !== "'") { // valor sin comillas: hasta el espacio o el cierre
          const ini2 = p;
          while (p < n && texto.charCodeAt(p) !== 62 && !esEsp(texto.charCodeAt(p))) p++;
          atributos[local(clave)] = texto.slice(ini2, p);
          continue;
        }
        const fin = texto.indexOf(comilla, p + 1);
        if (fin < 0) { p = n; break; }
        const valor = texto.slice(p + 1, fin);
        if (clave !== 'xmlns' && !clave.startsWith('xmlns:')) atributos[local(clave)] = valor.indexOf('&') < 0 ? valor : decodificar(valor);
        p = fin + 1;
      }
      if (h.abrir) h.abrir(nombre, atributos, vacio);
      if (vacio) { if (h.cerrar) h.cerrar(nombre); } else pila.push(nombre);
      i = p;
    }
  }

  // Árbol pequeño { n: nombre local, a: atributos, h: hijos (elementos o cadenas) } para las partes que se recorren varias veces (Word, OpenDocument).
  // Falla con un Error claro si el XML pasa los límites (elementos o niveles), para que un archivo hostil no agote la memoria.
  function arbol(texto, maxElementos = 3000000, maxProfundidad = 256) {
    const raiz = { n: '#', a: {}, h: [] };
    const pila = [raiz];
    let cuenta = 0;
    recorrer(texto, {
      abrir(n, a) {
        if (++cuenta > maxElementos) throw new Error('El documento tiene demasiados elementos y no se puede leer.');
        if (pila.length > maxProfundidad) throw new Error('El documento tiene demasiados niveles anidados y no se puede leer.');
        const e = { n, a, h: [] };
        pila[pila.length - 1].h.push(e);
        pila.push(e);
      },
      cerrar() { if (pila.length > 1) pila.pop(); },
      texto(t) { pila[pila.length - 1].h.push(t); },
    });
    return raiz.h.find((x) => typeof x !== 'string') || null;
  }

  globalThis.XmlLigero = { recorrer, decodificar, sinEscapes, arbol };
})();

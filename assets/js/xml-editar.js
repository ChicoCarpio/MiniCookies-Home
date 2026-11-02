/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Edición quirúrgica de XML para el escritor de libros de Excel (etapa 7). Complementa a xml-ligero.js: aquel solo lee por eventos y descarta dónde
// está cada cosa; este conserva la POSICIÓN de cada elemento en el texto, de modo que se puede insertar, reemplazar o borrar un fragmento y dejar
// todo lo demás exactamente como estaba (prefijos, espacios de nombres, mc:AlternateContent, extLst, comentarios, orden y espacios).
// No depende del navegador (nada de DOMParser ni XMLSerializer): corre igual en Node.
//
// Uso:
//   const raiz = XmlEditar.analizar(texto);                    // { nombre, local, attrs, ini, finApertura, iniCierre, fin, hijos, padre }
//   const fila = XmlEditar.hijo(XmlEditar.hijo(raiz, 'sheetData'), 'row');
//   texto = XmlEditar.aplicar(texto, [{ ini: fila.fin, fin: fila.fin, texto: '<row r="9"/>' }]);   // ediciones sobre el texto original
//   XmlEditar.analizar(texto, { profundidad: 3 })              // hojas enormes: solo se guardan los elementos hasta esa profundidad (la raíz es 1)
//   XmlEditar.esc(valor) / XmlEditar.escAttr(valor)            // escapes de texto y de atributo
//   XmlEditar.ponerAtributo(etiquetaApertura, 'r', '5')        // devuelve la etiqueta de apertura con ese atributo cambiado o añadido
//
// Las posiciones son índices del texto recibido: ini = el "<" de apertura, finApertura = justo después del ">" de apertura,
// iniCierre = el "<" de "</nombre>" (-1 si el elemento es vacío "<x/>"), fin = justo después de su última ">".
(function () {
  const MAX_PROFUNDIDAD = 2000;       // un XML hostil con millones de niveles no debe agotar la memoria
  const MAX_ELEMENTOS = 20000000;

  class ErrorXmlEditar extends Error {}

  const local = (nombre) => { const d = nombre.indexOf(':'); return d < 0 ? nombre : nombre.slice(d + 1); };
  const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
  function decodificar(s) {
    if (s.indexOf('&') < 0) return s;
    return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g, (todo, e) => {
      if (e.charCodeAt(0) === 35) {
        const c = e.charCodeAt(1) === 120 || e.charCodeAt(1) === 88 ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
        if (!(c >= 0 && c <= 0x10FFFF)) return todo;
        try { return String.fromCodePoint(c); } catch (error) { return todo; }
      }
      return Object.prototype.hasOwnProperty.call(ENT, e) ? ENT[e] : todo;
    });
  }

  // Los caracteres que XML 1.0 no admite (controles salvo tab, salto y retorno) se descartan: el archivo quedaría roto y Excel pediría repararlo.
  const SIN_CONTROL = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g;
  const esc = (v) => String(v).replace(SIN_CONTROL, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const escAttr = (v) => esc(v).replace(/"/g, '&quot;').replace(/\n/g, '&#10;').replace(/\r/g, '&#13;').replace(/\t/g, '&#9;');

  // El nombre solo puede empezar tras un espacio o una comilla: sin eso, una etiqueta con un texto largo sin "=" se reintentaba desde cada carácter
  // (cuadrático: 80.000 caracteres tardaban 3 s).
  const RE_ATTR = /(?<=^|[\s"'])([^\s=\/>"']+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

  // Fin de la etiqueta que empieza en `lt`: el ">" que no está dentro de comillas. -1 si no cierra.
  function finEtiqueta(t, lt) {
    let comilla = 0;
    for (let p = lt + 1; p < t.length; p++) {
      const k = t.charCodeAt(p);
      if (comilla) { if (k === comilla) comilla = 0; } else if (k === 34 || k === 39) comilla = k; else if (k === 62) return p;
    }
    return -1;
  }

  function analizar(texto, opciones) {
    const maxProf = (opciones && opciones.profundidad) || MAX_PROFUNDIDAD;
    const n = texto.length;
    const pila = [];       // todos los elementos abiertos (también los que no se guardan, para que las posiciones salgan bien)
    let raiz = null, cuenta = 0, i = 0;
    if (n && texto.charCodeAt(0) === 0xFEFF) i = 1;
    while (i < n) {
      const lt = texto.indexOf('<', i);
      if (lt < 0) break;
      const c = texto.charCodeAt(lt + 1);
      if (c === 33) { // comentario, CDATA, DOCTYPE
        let fin;
        if (texto.startsWith('<!--', lt)) { fin = texto.indexOf('-->', lt + 4); i = fin < 0 ? n : fin + 3; }
        else if (texto.startsWith('<![CDATA[', lt)) { fin = texto.indexOf(']]>', lt + 9); i = fin < 0 ? n : fin + 3; }
        else { fin = texto.indexOf('>', lt); i = fin < 0 ? n : fin + 1; }
        continue;
      }
      if (c === 63) { const fin = texto.indexOf('?>', lt + 2); i = fin < 0 ? n : fin + 2; continue; }
      if (c === 47) { // </nombre>
        const gt = texto.indexOf('>', lt + 2);
        if (gt < 0) throw new ErrorXmlEditar('XML sin cerrar');
        const nodo = pila.pop();
        if (!nodo) throw new ErrorXmlEditar('XML con cierre de más');
        if (nodo.guardado) { nodo.iniCierre = lt; nodo.fin = gt + 1; }
        i = gt + 1;
        continue;
      }
      const gt = finEtiqueta(texto, lt);
      if (gt < 0) throw new ErrorXmlEditar('XML sin cerrar');
      const vacio = texto.charCodeAt(gt - 1) === 47;
      const cuerpo = texto.slice(lt + 1, vacio ? gt - 1 : gt);
      let fn = 0;
      while (fn < cuerpo.length && !/\s/.test(cuerpo[fn])) fn++;
      const prof = pila.length + 1;
      if (prof > MAX_PROFUNDIDAD) throw new ErrorXmlEditar('XML demasiado anidado');
      const guardado = prof <= maxProf && (pila.length === 0 || pila[pila.length - 1].guardado);
      let nodo = { guardado, vacio };
      if (guardado) {
        if (++cuenta > MAX_ELEMENTOS) throw new ErrorXmlEditar('XML con demasiados elementos');
        const nombre = cuerpo.slice(0, fn);
        const attrs = {};
        if (fn < cuerpo.length) {
          RE_ATTR.lastIndex = 0;
          let m;
          const resto = cuerpo.slice(fn);
          while ((m = RE_ATTR.exec(resto))) attrs[m[1]] = decodificar(m[2] !== undefined ? m[2] : m[3]);
        }
        nodo = { nombre, local: local(nombre), attrs, ini: lt, finApertura: gt + 1, iniCierre: -1, fin: vacio ? gt + 1 : -1, hijos: [], padre: null, guardado, vacio };
        const padre = pila.length ? pila[pila.length - 1] : null;
        if (padre) { nodo.padre = padre; padre.hijos.push(nodo); } else if (!raiz) raiz = nodo;
        else throw new ErrorXmlEditar('XML con más de una raíz');
      }
      if (!vacio) pila.push(nodo);
      i = gt + 1;
    }
    if (pila.length) throw new ErrorXmlEditar('XML sin cerrar');
    if (!raiz) throw new ErrorXmlEditar('XML sin elementos');
    return raiz;
  }

  const hijo = (nodo, nombre) => { if (!nodo) return null; for (const h of nodo.hijos) if (h.local === nombre) return h; return null; };
  const hijos = (nodo, nombre) => (nodo ? nodo.hijos.filter((h) => h.local === nombre) : []);
  // El texto de dentro de un elemento, sin interpretar (con sus entidades).
  const interior = (texto, nodo) => (nodo.iniCierre < 0 ? '' : texto.slice(nodo.finApertura, nodo.iniCierre));
  // Atributo por nombre local (r:id → id): el primero que coincida.
  function attr(nodo, nombre) {
    if (!nodo) return undefined;
    if (Object.prototype.hasOwnProperty.call(nodo.attrs, nombre)) return nodo.attrs[nombre];
    for (const k of Object.keys(nodo.attrs)) if (local(k) === nombre && !k.startsWith('xmlns')) return nodo.attrs[k];
    return undefined;
  }

  // Aplica ediciones { ini, fin, texto } (posiciones del texto ORIGINAL) sin que se pisen entre sí: se ordenan y se aplican de atrás hacia adelante.
  function aplicar(texto, ediciones) {
    const e = ediciones.slice().sort((a, b) => a.ini - b.ini || a.fin - b.fin);
    for (let k = 1; k < e.length; k++) if (e[k].ini < e[k - 1].fin) throw new ErrorXmlEditar('ediciones que se pisan');
    const partes = [];
    let pos = 0;
    for (const x of e) { partes.push(texto.slice(pos, x.ini), x.texto); pos = x.fin; }
    partes.push(texto.slice(pos));
    return partes.join('');
  }

  // La etiqueta de apertura de un nodo, con ese atributo cambiado o añadido (valor === null lo quita). Respeta el resto tal cual.
  function ponerAtributo(etiqueta, nombre, valor) {
    const vacio = /\/\s*>$/.test(etiqueta);
    const cierre = vacio ? etiqueta.match(/\s*\/\s*>$/)[0] : '>';
    const cuerpo = etiqueta.slice(0, etiqueta.length - cierre.length);
    const re = new RegExp('(\\s)' + nombre.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s*=\\s*(?:"[^"]*"|\'[^\']*\')');
    if (re.test(cuerpo)) return cuerpo.replace(re, valor === null ? '' : (todo, esp) => esp + nombre + '="' + escAttr(valor) + '"') + cierre;
    if (valor === null) return etiqueta;
    return cuerpo + ' ' + nombre + '="' + escAttr(valor) + '"' + cierre;
  }

  globalThis.XmlEditar = { analizar, hijo, hijos, interior, attr, aplicar, ponerAtributo, esc, escAttr, decodificar, local, ErrorXmlEditar };
})();

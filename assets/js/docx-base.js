/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Base común de los módulos de Word propios (docx-lector.js y docx-maquetador.js):
// árbol del XML, ayudas para recorrerlo y lectura del paquete (.docx = zip con
// partes y relaciones). No depende del navegador: corre igual en Node.
// Requiere ZipPropio y XmlLigero.
//
// API (globalThis.DocxBase):
//   arbol(xml) → { n, a, h, t } | null      hijo(e, n)   lista(e, n)   activo(e)   buscar(e, n)
//   resolver(base, destino)  rutaRels(parte)  leerRels(zip, parte)  porTipo(rels, fin)  base64(bytes)  hrefSeguro(url)
//   abrirPaquete(datos) → { zip, principal, rels, cuerpo }   (datos: ArrayBuffer, Uint8Array o Blob)

(function () {
  'use strict';

  // ───────── XML → árbol ─────────
  function arbol(xml) {
    const raiz = { n: '#', a: {}, h: [], t: '' };
    const pila = [raiz];
    XmlLigero.recorrer(xml, {
      abrir(n, a) {
        const e = { n, a, h: [], t: '' };
        pila[pila.length - 1].h.push(e);
        pila.push(e);
      },
      cerrar() { if (pila.length > 1) pila.pop(); },
      texto(t) { pila[pila.length - 1].t += t; },
    });
    return raiz.h[0] || null;
  }
  const hijo = (e, n) => (e ? e.h.find((x) => x.n === n) || null : null);
  const lista = (e, n) => (e ? e.h.filter((x) => x.n === n) : []);
  const activo = (e) => {
    if (!e) return false;
    const v = e.a.val;
    return !(v === '0' || v === 'false' || v === 'off' || v === 'none');
  };
  function buscar(e, n) { // primer descendiente con ese nombre
    if (!e) return null;
    for (const x of e.h) {
      if (x.n === n) return x;
      const r = buscar(x, n);
      if (r) return r;
    }
    return null;
  }

  // ───────── Paquete ─────────
  function resolver(base, destino) {
    if (destino.startsWith('/')) return destino.slice(1);
    const partes = base.split('/').slice(0, -1);
    for (const p of destino.split('/')) {
      if (p === '..') partes.pop();
      else if (p !== '.' && p !== '') partes.push(p);
    }
    return partes.join('/');
  }
  function rutaRels(parte) {
    const i = parte.lastIndexOf('/');
    return parte.slice(0, i + 1) + '_rels/' + parte.slice(i + 1) + '.rels';
  }
  async function leerRels(zip, parte) {
    const f = zip.file(rutaRels(parte));
    const mapa = new Map();
    if (!f) return mapa;
    const raiz = arbol(await f.async('string'));
    for (const r of lista(raiz, 'Relationship')) {
      const externo = (r.a.TargetMode || '').toLowerCase() === 'external';
      mapa.set(r.a.Id, { tipo: r.a.Type || '', destino: externo ? r.a.Target : resolver(parte, r.a.Target || ''), externo });
    }
    return mapa;
  }
  // Bytes → base64 (para las imágenes como data:), por trozos: apply con millones de argumentos no cabe.
  function base64(bytes) {
    let s = '';
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  }

  // Un enlace del archivo no puede ser javascript:, vbscript: ni data: (el navegador ignora espacios y caracteres
  // de control en el esquema: "java	script:" vale como "javascript:").
  const hrefSeguro = (h) => (/^(javascript|vbscript|data):/i.test(String(h || '').replace(/[\x00-\x20]/g, '')) ? '' : h || '');
  const porTipo = (rels, fin) => { for (const r of rels.values()) if (r.tipo.endsWith(fin)) return r.destino; return null; };

  // Abre el .docx: busca la parte principal (por _rels/.rels), sus relaciones y el <w:body>.
  async function abrirPaquete(datos) {
    if (typeof Blob !== 'undefined' && datos instanceof Blob) datos = await datos.arrayBuffer();
    // Con formatos-word.js en la página, un .doc o un .odt se pasa a .docx antes de leerlo; un .docx pasa sin tocarse
    if (typeof FormatosWord !== 'undefined') datos = await FormatosWord.comoDocx(datos);
    const zip = await ZipPropio.loadAsync(datos);
    let principal = 'word/document.xml';
    const f = zip.file('_rels/.rels');
    if (f) {
      const raiz = arbol(await f.async('string'));
      const r = lista(raiz, 'Relationship').find((x) => /\/officeDocument$/.test(x.a.Type || ''));
      if (r && r.a.Target) principal = resolver('x', r.a.Target);
    }
    const docFile = zip.file(principal);
    if (!docFile) throw new Error('No es un documento de Word válido (falta word/document.xml).');
    const rels = await leerRels(zip, principal);
    const cuerpo = hijo(arbol(await docFile.async('string')), 'body');
    if (!cuerpo) throw new Error('El documento no tiene cuerpo.');
    return { zip, principal, rels, cuerpo };
  }

  globalThis.DocxBase = { arbol, hijo, lista, activo, buscar, resolver, rutaRels, leerRels, porTipo, base64, hrefSeguro, abrirPaquete };
})();

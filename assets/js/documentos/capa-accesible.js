/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Capa accesible de Documentos (Fase 6, etapa 6.3). Un lienzo no le dice nada a un lector de pantalla: esta capa es un árbol DOM oculto a la vista
// (pero visible para las tecnologías de apoyo) que sigue el modelo del documento en el orden de lectura: títulos (h1 a h6, por el nivel de esquema
// o por el nombre del estilo), párrafos, listas con viñetas y numeradas, tablas con su fila de encabezado, imágenes con su texto alternativo,
// enlaces, saltos de línea y las notas al pie. No depende del diseño de hojas: sigue el modelo, que es lo que lee la persona.
//
//   const capa = CapaAccesible.crear(contenedor, modelo);   // contenedor: el elemento donde se crea la capa (suele ser el del visor)
//   capa.actualizar();                                       // reconstruye (la edición actualizará solo los párrafos tocados)
//   capa.elemento · capa.destruir()
(function () {
  'use strict';
  const MAX_BLOQUES = 20000;   // un documento más largo se corta aquí (la capa no puede crecer sin límite)

  function crear(contenedor, modelo, opciones = {}) {
    const doc = contenedor.ownerDocument;
    // lo que la capa dice por su cuenta, en el idioma de la página (opciones.textos)
    const txt = Object.assign({ contenido: 'Contenido del documento', imagen: 'Imagen', dibujo: 'Dibujo', notas: 'Notas al pie' }, opciones.textos || {});
    const raiz = doc.createElement('div');
    raiz.className = 'vd-accesible';
    raiz.setAttribute('role', 'document');
    raiz.setAttribute('aria-label', opciones.titulo || txt.contenido);
    // oculto a la vista, no al lector de pantalla
    raiz.style.cssText = 'position:absolute;width:1px;height:1px;margin:-1px;padding:0;overflow:hidden;clip:rect(0 0 0 0);clip-path:inset(50%);white-space:nowrap;border:0';
    contenedor.appendChild(raiz);
    const m = modelo.m;

    const el = (nombre, texto) => { const e = doc.createElement(nombre); if (texto) e.textContent = texto; return e; };
    function nivelTitulo(b) {
      if (b.p.nivelEsquema !== undefined && b.p.nivelEsquema >= 0 && b.p.nivelEsquema <= 8) return Math.min(b.p.nivelEsquema + 1, 6);
      const nombre = b.estilo ? m.estilos.nombre(b.estilo) : '';
      const r = /^(?:heading|t[ií]tulo)\s*([1-9])$/i.exec(nombre || '');
      if (r) return Math.min(+r[1], 6);
      return /^(?:title|t[ií]tulo)$/i.test(nombre || '') ? 1 : 0;
    }
    // El contenido de un párrafo: texto, saltos, enlaces e imágenes
    function rellenar(e, b) {
      let enlace = null, ancla = null, actual = e;
      const cerrar = () => { enlace = null; actual = e; };
      for (const z of b.piezas) {
        if ((z.enlace || null) !== enlace) { cerrar(); if (z.enlace) { ancla = doc.createElement('a'); ancla.setAttribute('href', z.enlace); e.appendChild(ancla); actual = ancla; enlace = z.enlace; } }
        if (z.t === 'texto') actual.appendChild(doc.createTextNode(z.s));
        else if (z.t === 'tab') actual.appendChild(doc.createTextNode(' '));
        else if (z.t === 'br') actual.appendChild(doc.createElement('br'));
        else if (z.t === 'imagen' || z.t === 'dibujo' || z.t === 'cuadro') {
          const img = doc.createElement('span');
          img.setAttribute('role', 'img');
          img.setAttribute('aria-label', z.descr || (z.t === 'imagen' ? txt.imagen : txt.dibujo));
          actual.appendChild(img);
        }
      }
    }
    function parrafo(b) {
      const nivel = nivelTitulo(b);
      const e = doc.createElement(nivel ? 'h' + nivel : 'p');
      rellenar(e, b);
      if (!e.textContent && !e.querySelector('[role=img]')) e.setAttribute('aria-hidden', 'true');   // un párrafo vacío no se lee
      return e;
    }
    function tabla(b) {
      const t = doc.createElement('table');
      const encabezado = b.tbl && b.tbl.aspecto && b.tbl.aspecto.primeraFila;
      b.filas.forEach((f, i) => {
        const tr = doc.createElement('tr');
        for (const c of f.celdas) {
          if (c.omitir) continue;
          const td = doc.createElement(encabezado && i === 0 ? 'th' : 'td');
          if (encabezado && i === 0) td.setAttribute('scope', 'col');
          if (c.span > 1) td.setAttribute('colspan', String(c.span));
          if (c.filas > 1) td.setAttribute('rowspan', String(c.filas));
          bloques(td, c.bloques);
          tr.appendChild(td);
        }
        t.appendChild(tr);
      });
      return t;
    }
    // Bloques en orden; los párrafos de lista seguidos del mismo tipo forman una sola lista
    function bloques(destino, lista) {
      let listaActual = null, tipoLista = null;
      for (const b of lista) {
        if (b.tipo === 'tabla') { listaActual = null; destino.appendChild(tabla(b)); continue; }
        if (b.etiqueta && b.texto.replace(/[￼\f]/g, '').trim()) {
          const tipo = b.etiqueta.fmt === 'bullet' ? 'ul' : 'ol';
          if (!listaActual || tipoLista !== tipo) { listaActual = doc.createElement(tipo); tipoLista = tipo; destino.appendChild(listaActual); }
          const li = doc.createElement('li');
          rellenar(li, b);
          listaActual.appendChild(li);
          continue;
        }
        listaActual = null;
        destino.appendChild(parrafo(b));
      }
    }

    function actualizar() {
      raiz.textContent = '';
      bloques(raiz, modelo.bloques.slice(0, MAX_BLOQUES));
      if (modelo.notas && modelo.notas.size) {
        const s = doc.createElement('section');
        s.setAttribute('aria-label', txt.notas);
        for (const n of modelo.notas.values()) { const d = doc.createElement('div'); bloques(d, n.bloques); s.appendChild(d); }
        raiz.appendChild(s);
      }
    }
    actualizar();
    return { elemento: raiz, actualizar, destruir() { raiz.remove(); } };
  }

  globalThis.CapaAccesible = { crear };
})();

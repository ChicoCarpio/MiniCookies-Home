/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Cirugía del XML de un párrafo de .odt para EdicionDoc (Fase 6b, etapa 6b.3): el «dialecto» ODT del editor de texto. Sin DOM.
//
// EdicionDoc corta el XML del párrafo y cambia justo lo que se tocó; lo que cambia de un formato a otro es cómo se corta. En Word el texto vive en
// `<w:t>`; en ODF está suelto entre las etiquetas y los espacios se colapsan (ODF 1.2 §6.1.2). Por eso:
//   · el texto de un hueco (o de un `<text:s/>`) se vuelve a escribir ENTERO con `codificar` cada vez que se toca: un espacio suelto solo va literal
//     entre dos caracteres visibles; al principio, al final o en tandas va como `<text:s text:c="n"/>`;
//   · el espacio que sigue al lugar editado se vuelve a escribir si quedaría sin un carácter visible delante (si no, el lector lo descartaría);
//   · tras cada edición el párrafo se vuelve a leer y su texto tiene que ser el esperado; si no, la edición se RECHAZA (no se escribe nada).
// El formato de caracteres y de párrafo y los estilos con nombre (6b.4) los hace odt-formato.js, que este dialecto ofrece a EdicionDoc.
// Fuera de este dialecto no se editan párrafos protegidos (índices), saltos de página ni de sección, tablas ni imágenes: `puede(capacidad)` lo dice y
// EdicionDoc responde null.
//
// Requiere xml-editar.js y odt-formato.js.
(function () {
  'use strict';
  const X = () => globalThis.XmlEditar;
  const SIN_PARTIR = /^(?:note|annotation|ruby|frame|meta|meta-field|date|time|page-number|page-count|title|subject|author-name|variable-set|variable-get|user-defined|sequence|bookmark-ref|reference-ref)$/;

  class Rechazo extends Error { constructor(motivo) { super('edición rechazada: ' + motivo); this.rechazo = motivo; } }

  const prefijoDe = (raiz) => { const i = raiz.nombre.indexOf(':'); return i < 0 ? '' : raiz.nombre.slice(0, i + 1); };
  const visible = (c) => !!c && !/[ \t\n\r\u2028]/.test(c);

  // El texto plano de un párrafo a partir de sus piezas (el mismo que `b.texto`)
  function textoDe(piezas) {
    let t = '';
    for (const z of piezas) {
      if (z.t === 'texto') t += z.s; else if (z.t === 'tab') t += '\t'; else if (z.t === 'br') t += '\u2028';
      else if (z.t === 'pagina' || z.t === 'columna') t += '\f'; else if (z.t !== 'ancla') t += '￼';
    }
    return t;
  }

  // El texto `s` como contenido de un párrafo ODF. `antes`: el carácter que lo precede en el párrafo ('' si es el principio).
  // `finLiteral`: lo que sigue a `s` no empieza con un espacio suelto, así que un único espacio al final puede ir tal cual (si no, va como <text:s/>).
  function codificar(s, antes, pf, finLiteral) {
    let sal = '', previo = antes;
    const n = s.length;
    for (let i = 0; i < n;) {
      const c = s[i];
      if (c === ' ') {
        let j = i; while (j < n && s[j] === ' ') j++;
        const cuantos = j - i;
        const finDelTrozo = j >= n;
        if (cuantos === 1 && visible(previo) && (!finDelTrozo || finLiteral)) sal += ' ';
        else if (visible(previo) && cuantos > 1 && !(finDelTrozo && cuantos === 1)) {
          sal += ' ' + (cuantos - 1 === 1 ? `<${pf}s/>` : `<${pf}s ${pf}c="${cuantos - 1}"/>`);
        } else sal += cuantos === 1 ? `<${pf}s/>` : `<${pf}s ${pf}c="${cuantos}"/>`;
        previo = ' '; i = j; continue;
      }
      if (c === '\t') sal += `<${pf}tab/>`;
      else if (c === '\u2028') sal += `<${pf}line-break/>`;
      else {
        // un trozo de texto corrido (sin espacios, tabs ni saltos)
        let j = i; while (j < n && s[j] !== ' ' && s[j] !== '\t' && s[j] !== '\u2028') j++;
        sal += X().esc(s.slice(i, j)); previo = s[j - 1]; i = j; continue;
      }
      previo = c; i++;
    }
    return sal;
  }
  // `codificar` termina un trozo que acaba en espacio con `<text:s/>` (no con un espacio suelto): el que sigue puede empezar con otro suelto sin perderlo
  const editable = (z) => ((z.t === 'texto' && z.hueco && !z.opaco && !z.rara) || (z.t === 'texto' && z.espacio)) && !z.nota;
  const nodos = (raiz) => { const m = new Map(); const rec = (n) => { for (const h of n.hijos) { m.set(h.ini, h); rec(h); } }; rec(raiz); return m; };

  function ancestros(raiz, corte) {
    const sal = [];
    let n = raiz;
    for (;;) {
      const h = n.hijos.find((c) => c.ini < corte && corte < c.fin);
      if (!h) break;
      sal.push(h); n = h;
    }
    return sal;
  }

  function crear(modelo) {
    // ── quitar [desde, hasta) ──
    function edicionesQuitar(raiz, piezas, desde, hasta, texto) {
      const pf = prefijoDe(raiz), ed = [];
      const t = texto !== undefined ? texto : textoDe(piezas);
      const nuevo = t.slice(0, desde) + t.slice(hasta);
      const reescritos = new Set();
      for (const z of piezas) {
        if (z.hasta <= z.desde || z.hasta <= desde || z.desde >= hasta) continue;
        const a = Math.max(desde, z.desde) - z.desde, b = Math.min(hasta, z.hasta) - z.desde;
        if (editable(z)) {
          const resto = z.s.slice(0, a) + z.s.slice(b);
          const inicio = Math.min(z.desde, desde);
          ed.push({ ini: z.ini, fin: z.fin, texto: resto ? codificar(resto, nuevo[inicio - 1] || '', pf) : '' });
          reescritos.add(z);
        } else if (z.t === 'tab' || z.t === 'br' || z.t === 'campo' || z.t === 'imagen' || z.t === 'dibujo' || z.t === 'cuadro' || z.nota) ed.push({ ini: z.ini, fin: z.fin, texto: '' });
        else throw new Rechazo('el texto de un campo o de un índice no se edita aquí');
      }
      // el hueco que queda pegado a lo borrado pierde su carácter visible de delante: si empieza con espacio, se escribe con <text:s/>
      const sig = piezas.find((z) => z.desde === hasta && z.hasta > z.desde && editable(z) && !reescritos.has(z) && z.s[0] === ' ');
      if (sig && !visible(nuevo[desde - 1])) ed.push({ ini: sig.ini, fin: sig.fin, texto: codificar(sig.s, nuevo[desde - 1] || '', pf) });
      return ed;
    }

    // ── meter `s` en `pos` ──
    function edicionesInsertar(xml, raiz, piezas, pos, s, texto) {
      if (!s) return [];
      const pf = prefijoDe(raiz), t = texto !== undefined ? texto : textoDe(piezas);
      s = s.replace(/\f/g, '');
      if (!s) return [];
      let prev = null, sig = null;
      for (const z of piezas) {
        if (z.hasta <= z.desde) continue;
        if (z.desde < pos && z.hasta >= pos) prev = z;
        if (z.desde === pos && !sig) sig = z;
        if (z.desde > pos) break;
      }
      const ed = [];
      let objetivo = null, k = 0;
      if (prev && editable(prev)) { objetivo = prev; k = pos - prev.desde; }
      else if (sig && editable(sig)) { objetivo = sig; k = 0; }
      if (objetivo) {
        const nuevo = objetivo.s.slice(0, k) + s + objetivo.s.slice(k);
        ed.push({ ini: objetivo.ini, fin: objetivo.fin, texto: codificar(nuevo, t[objetivo.desde - 1] || '', pf) });
      } else {
        const donde = prev ? prev.fin : sig ? sig.ini : raiz.finApertura;
        ed.push({ ini: donde, fin: donde, texto: codificar(s, t[pos - 1] || '', pf) });
      }
      // el hueco siguiente que empieza con espacio necesita un carácter visible delante
      const ultimo = s[s.length - 1];
      const s2 = sig && sig !== objetivo && editable(sig) && sig.s[0] === ' ' ? sig : null;
      if (s2 && !visible(ultimo)) ed.push({ ini: s2.ini, fin: s2.fin, texto: codificar(s2.s, ultimo, pf) });
      return ed;
    }

    // ── partir en `pos` ──
    function partirXml(xml, raiz, piezas, pos, texto) {
      const pf = prefijoDe(raiz), t = texto !== undefined ? texto : textoDe(piezas);
      const util = piezas.filter((z) => z.hasta > z.desde);
      let prev = null;
      for (const z of util) if (z.desde < pos && z.hasta >= pos) prev = z;
      const sig = util.find((z) => z.desde >= pos);
      const ed = [];
      let corte, dentro = null;
      if (prev && editable(prev) && pos < prev.hasta) {
        const k = pos - prev.desde, A = prev.s.slice(0, k), B = prev.s.slice(k);
        const encA = codificar(A, t[prev.desde - 1] || '', pf), encB = codificar(B, '', pf);
        ed.push({ ini: prev.ini, fin: prev.fin, texto: encA + encB });
        corte = prev.ini + encA.length; dentro = prev;
      } else if (prev) corte = prev.fin;
      else corte = raiz.finApertura;
      // lo que queda al principio del párrafo nuevo no puede empezar con un espacio suelto
      const primero = dentro ? null : util.find((z) => z.desde >= pos);
      if (primero && editable(primero) && primero.s[0] === ' ') ed.push({ ini: primero.ini, fin: primero.fin, texto: codificar(primero.s, '', pf) });
      const xml2 = ed.length ? X().aplicar(xml, ed) : xml;
      const raiz2 = X().analizar(xml2);
      let pila = ancestros(raiz2, corte);
      for (let vuelta = 0; vuelta < 8; vuelta++) {
        const mala = pila.find((n) => SIN_PARTIR.test(n.local));
        if (!mala) break;
        corte = mala.fin; pila = ancestros(raiz2, corte);
      }
      const cierres = pila.slice().reverse().map((n) => `</${n.nombre}>`).join('');
      let tag = xml2.slice(raiz2.ini, raiz2.finApertura);
      for (const a of ['xml:id', pf + 'id']) tag = X().ponerAtributo(tag, a, null);
      const abrir = (n) => xml2.slice(n.ini, n.finApertura);
      const antes = xml2.slice(0, corte) + cierres + `</${raiz2.nombre}>`;
      const despues = tag + pila.map(abrir).join('') + xml2.slice(corte, raiz2.iniCierre) + `</${raiz2.nombre}>`;
      return { antes, despues };
    }

    // ── unir dos párrafos ──
    function unirXml(xa, ra, xb, rb, vacioA, vacioB) {
      const usoB = vacioA && !vacioB;
      const cuerpo = (x, r) => x.slice(r.finApertura, r.iniCierre);
      const base = usoB ? [xb, rb] : [xa, ra];
      let b = cuerpo(xb, rb);
      // el espacio suelto del principio de B queda detrás del final de A: si A termina sin carácter visible, se vuelve <text:s/>
      const pf = prefijoDe(ra);
      let piezasB = null;
      try { piezasB = modelo.releer(xb, {}).piezas; } catch (e) { piezasB = null; }
      const primero = piezasB && piezasB.find((z) => z.hasta > z.desde);
      const piezasA = modelo.releer(xa, {}).piezas, tA = textoDe(piezasA);
      if (primero && editable(primero) && primero.s[0] === ' ' && !visible(tA[tA.length - 1]) && !vacioA) {
        const rel = primero.ini - rb.finApertura;
        b = b.slice(0, rel) + codificar(primero.s, tA[tA.length - 1] || '', pf) + b.slice(primero.fin - rb.finApertura);
      }
      const inicio = base[0].slice(base[1].ini, base[1].finApertura);
      return { xml: inicio + cuerpo(xa, ra) + b + `</${base[1].nombre}>`, usoB };
    }

    // Cambia el estilo de un <text:p>/<text:h> (y lo vuelve p o h según el estilo); null = sin estilo
    function ponerEstilo(xml, id) {
      const raiz = X().analizar(xml), pf = prefijoDe(raiz);
      let ab = xml.slice(raiz.ini, raiz.finApertura);
      ab = X().ponerAtributo(ab, pf + 'style-name', id || null);
      const e = id ? modelo.est.resolverParrafo(id) : null;
      const esTitulo = !!(e && e.p.nivelEsquema !== undefined);
      let cierre = xml.slice(raiz.iniCierre, raiz.fin), nombre = raiz.nombre;
      if (raiz.local === 'h' && !esTitulo) { ab = X().ponerAtributo(ab.replace(/^<[^\s>/]+/, '<' + pf + 'p'), pf + 'outline-level', null); cierre = `</${pf}p>`; nombre = pf + 'p'; }
      else if (raiz.local === 'p' && esTitulo) { ab = ab.replace(/^<[^\s>/]+/, '<' + pf + 'h'); ab = X().ponerAtributo(ab, pf + 'outline-level', String(e.p.nivelEsquema + 1)); cierre = `</${pf}h>`; }
      return ab + xml.slice(raiz.finApertura, raiz.iniCierre) + cierre;
    }

    // ¿Se puede hacer este reemplazo? `bloques`: los párrafos que toca; `parte`: si parte o une párrafos. Devuelve el motivo del rechazo o null.
    function rechazar(bloques, parte) {
      if (bloques.some((b) => b.protegido)) return 'protegido';
      if (parte) {
        if (bloques.some((b) => b.lista)) return 'lista';
        if (bloques.some((b, i) => (b.abreSeccion || (b.p && b.p.saltoAntes)) && (i > 0 || parte === 'partir'))) return 'salto';
      }
      return null;
    }
    const CAPACIDADES = new Set(['texto', 'formato', 'estilos', 'listas', 'tablas', 'imagenes', 'campos', 'saltos']);
    let fmt = null;
    const F = () => fmt || (fmt = globalThis.OdtFormato.crear(modelo, { codificar, Rechazo, prefijoDe }));
    let tb = null;
    const T = () => tb || (tb = globalThis.OdtTabla.crear(modelo, { formato: F() }));
    const URI_DRAW = 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0', URI_SVG = 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0', URI_XLINK = 'http://www.w3.org/1999/xlink';
    const URI_TEXT = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0';
    const TIPOS = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp' };
    const imagenes = [];
    let manifiesto0;
    const redondo = (v) => Math.round(v * 100) / 100;
    // Las imágenes van a Pictures/ y se anotan en el manifiesto al guardar, solo las que el documento usa (en content.xml o en un encabezado de styles.xml).
    // La de una imagen deshecha se quita del paquete al guardar: LibreOffice no abre un .odt con un archivo que el manifiesto no nombra (medido 2026-10-31).
    async function imagenInline(bytes, ext, ancho, alto, descripcion, pf) {
      const paquete = modelo.paquete;
      let n = 1;
      const ocupada = (r) => paquete.existe(r) || imagenes.some((i) => i.ruta === r);
      while (ocupada('Pictures/imagen' + n + '.' + ext)) n++;
      const ruta = 'Pictures/imagen' + n + '.' + ext;
      paquete.escribir(ruta, bytes);
      imagenes.push({ ruta, tipo: TIPOS[ext], bytes });
      const desc = descripcion ? `<svg:desc>${X().esc(descripcion)}</svg:desc>` : '';
      return `<${pf}p><draw:frame xmlns:draw="${URI_DRAW}" xmlns:svg="${URI_SVG}" xmlns:xlink="${URI_XLINK}" draw:name="Imagen ${n}" ${pf}anchor-type="as-char" svg:width="${redondo(Math.min(ancho, 1584))}pt" svg:height="${redondo(Math.min(alto, 1584))}pt" draw:z-index="0">`
        + `<draw:image xlink:href="${ruta}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>${desc}</draw:frame></${pf}p>`;
    }
    async function alGuardarManifiesto(paquete, xmlFinal) {
      if (!imagenes.length) return;
      if (manifiesto0 === undefined) manifiesto0 = await paquete.leer('META-INF/manifest.xml');
      if (manifiesto0 === null) return;
      const estilos = (await paquete.leer('styles.xml')) || '';
      const usada = (i) => xmlFinal.includes('"' + i.ruta + '"') || estilos.includes('"' + i.ruta + '"');
      for (const i of imagenes) {
        if (!usada(i)) { if (paquete.existe(i.ruta)) paquete.quitar(i.ruta); }
        else if (!paquete.existe(i.ruta)) paquete.escribir(i.ruta, i.bytes);
      }
      const faltan = imagenes.filter((i) => usada(i) && !manifiesto0.includes('full-path="' + i.ruta + '"'));
      let nuevo = manifiesto0;
      if (faltan.length) {
        const r = X().analizar(manifiesto0), pm = r.nombre.slice(0, r.nombre.indexOf(':') + 1);
        nuevo = X().aplicar(manifiesto0, [{ ini: r.iniCierre, fin: r.iniCierre, texto: faltan.map((i) => `<${pm}file-entry ${pm}full-path="${i.ruta}" ${pm}media-type="${i.tipo}"/>`).join('') }]);
      }
      if ((await paquete.leer('META-INF/manifest.xml')) !== nuevo) paquete.escribir('META-INF/manifest.xml', nuevo);
    }
    let lst = null;
    const L = () => lst || (lst = globalThis.OdtListas.crear(modelo, { formato: F(), prefijoDe, Rechazo }));
    return {
      formato: 'odt', edicionesQuitar, edicionesInsertar, partirXml, unirXml, ponerEstilo, quitarLista: (xml) => xml, rechazar, puede: (c) => CAPACIDADES.has(c), Rechazo, textoDe,
      // formato (odt-formato.js)
      formatoCaracter: (Fp, desde, hasta, cambios, rBase) => F().formatoCaracter(Fp, desde, hasta, cambios, rBase),
      formatoParrafo: (xml, cambios) => F().formatoParrafo(xml, cambios),
      estiloConservando: (xml, id) => F().ponerEstilo(xml, id, { ponerEstilo }),
      enEstilos: (fn) => F().enEstilos(fn),
      estiloComun: (n) => F().estiloComun(n), estilosDisponibles: () => F().estilosDisponibles(), asegurarEstilo: (id) => F().asegurarEstilo(id), estiloPredeterminado: () => F().predeterminado(),
      finalizar: (xml) => (fmt ? fmt.finalizar(xml) : xml),
      alGuardar: async (paquete, xml) => { if (fmt) await fmt.alGuardar(paquete, xml); await alGuardarManifiesto(paquete, xml); },
      // tablas, imágenes, campos (odt-tabla.js)
      tabla: {
        nueva: (...a) => T().nueva(...a), insertarFila: (...a) => T().insertarFila(...a), eliminarFila: (...a) => T().eliminarFila(...a),
        insertarColumna: (...a) => T().insertarColumna(...a), eliminarColumna: (...a) => T().eliminarColumna(...a),
      },
      parrafoVacio: () => `<text:p xmlns:text="${URI_TEXT}"/>`,
      imagenInline, tiposImagen: TIPOS,
      estilosCreados: () => (fmt ? [...fmt.creados.values()] : []),
      campoXml: (nombre, pf) => `<${pf}p>` + (nombre === 'PAGE' ? `<${pf}page-number ${pf}select-page="current">1</${pf}page-number>` : `<${pf}page-count>1</${pf}page-count>`) + `</${pf}p>`,
      // listas (odt-listas.js)
      operacionLista: (op, bloques, frag, objetivo, extra) => L().aplicar(op, bloques, frag, objetivo, extra), tipoDeEntrada: (e) => L().tipoDe(e),
      formatoInterno: F,
    };
  }

  globalThis.OdtParrafo = { crear, codificar, textoDe, Rechazo };
})();

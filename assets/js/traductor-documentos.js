/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Traductor de documentos: separa el texto del formato, lo manda a traducir y rearma el documento con el texto traducido y el mismo formato.
// No traduce nada por sí mismo: recibe la función que traduce (TraductorHilos.traducir, o una falsa en las pruebas). Corre en Node y en el
// navegador (nada de DOM en el núcleo, en TXT ni en Markdown).
//
//   const salida = await TraductorDocumentos.traducirTexto(texto, { traducir, destino, progreso })        // .txt -> string
//   const salida = await TraductorDocumentos.traducirMarkdown(texto, { traducir, destino, progreso })     // .md  -> string
//        traducir(segmentos, destino) -> Promise<[texto]>   segmentos = [{ texto, contexto? }]
//        progreso(hechos, total) opcional
//
// Núcleo (lo usan todos los formatos):
//   unidad = [{ texto, formato }]   los tramos de un párrafo, en orden; formato = clave de texto ('' = sin formato especial)
//   1. Cada unidad se parte en ORACIONES (el modelo aprendió de oraciones; un párrafo largo pasaría de las 512 piezas). Cada oración se
//      traduce con la anterior como contexto (<ctx>; también la última del párrafo anterior), para que los términos se mantengan.
//   2. Cada oración pasa a texto con MARCAS: el formato que más letras tiene va sin marca; cada tramo con otro formato va entre <k>…</k>
//      (k = 1…8, uno por tramo). Los espacios del borde de un tramo quedan fuera de la marca.
//   3. La traducción se vuelve a partir en tramos por sus marcas. Si las marcas vienen rotas (falta o sobra una, o están cruzadas), PLAN B:
//      se traduce la oración sin marcas y queda toda con el formato principal. Nunca se pierde texto.
//   4. Los espacios del principio y del final de la unidad se conservan tal cual.
(function (raiz) {
  'use strict';

  const MAX_MARCAS = 8;
  const SIN_ESPACIO = new Set(['zh', 'ja']);             // idiomas de destino que juntan las oraciones sin espacio

  // ───────── Oraciones ─────────
  // Abreviaturas comunes en los idiomas latinos (sin el punto): un punto después de ellas no corta la oración.
  const ABREVIATURAS = new Set(('sr sra srta dr dra lic ing arq prof sto sta ud uds vd vds pág págs p pp cap art núm nº no etc ej '
    + 'mr mrs ms st vs e.g i.e cf al approx dept est fig inc ltd co corp jr sr ave blvd '
    + 'm mme mlle dr hr fr frau nr bzw usw vgl z.b d.h ca s sig dott ecc '
    + 'jan feb mar apr jun jul aug sep sept oct nov dec ene abr ago dic').split(' '));
  const FIN_CJK = /[。！？!?]/;

  // Posiciones (en el texto plano de la unidad) donde termina una oración: después del signo y de los espacios que siguen.
  function cortesDeOracion(texto) {
    const cortes = [];
    const re = /([.!?…]+)(["'”’»)\]]*)(\s+)(?=["'“‘«(¿¡]?[\p{Lu}\p{Lo}\p{N}])|([。！？]+)(["'”’」』)]*)/gu;
    let m;
    while ((m = re.exec(texto))) {
      const fin = m.index + m[0].length;
      if (fin >= texto.length) break;
      if (m[1] === '.') {
        // «Sr. Pérez», «p. ej. algo», «J. R. R. Tolkien», «3. Introducción»: no se corta.
        const antes = texto.slice(0, m.index).match(/([\p{L}.]+)$/u);
        const palabra = antes ? antes[1].toLowerCase().replace(/\.$/, '') : '';
        if (ABREVIATURAS.has(palabra) || /^\p{L}$/u.test(palabra)) continue;
        if (/^\s*\d+$/.test(texto.slice(0, m.index))) continue;          // «3. Introducción»: la numeración del principio
      }
      cortes.push(fin);
    }
    return cortes;
  }

  // Parte una unidad (tramos) en oraciones; un corte que cae dentro de un tramo lo divide en dos con el mismo formato.
  function partirEnOraciones(tramos) {
    const plano = tramos.map((t) => t.texto).join('');
    const cortes = cortesDeOracion(plano);
    if (!cortes.length) return [tramos];
    const oraciones = [[]];
    let pos = 0, c = 0;
    for (const t of tramos) {
      let texto = t.texto, ini = pos;
      while (c < cortes.length && cortes[c] > ini && cortes[c] < ini + texto.length) {
        const k = cortes[c] - ini;
        oraciones[oraciones.length - 1].push({ texto: texto.slice(0, k), formato: t.formato });
        oraciones.push([]);
        texto = texto.slice(k); ini += k; c++;
      }
      if (c < cortes.length && cortes[c] === ini + texto.length) {
        oraciones[oraciones.length - 1].push({ texto, formato: t.formato });
        oraciones.push([]); c++;
      } else if (texto) {
        oraciones[oraciones.length - 1].push({ texto, formato: t.formato });
      }
      pos += t.texto.length;
    }
    return oraciones.filter((o) => o.some((t) => t.texto.trim()));
  }

  // ───────── Marcas ─────────
  function juntar(tramos) {
    const r = [];
    for (const t of tramos) {
      if (!t.texto) continue;
      const u = r[r.length - 1];
      if (u && u.formato === t.formato) u.texto += t.texto; else r.push({ texto: t.texto, formato: t.formato });
    }
    return r;
  }

  // tramos de UNA oración -> { texto (con marcas), marcas: [formato de cada k], principal, inicio, fin } (inicio/fin: espacios del borde)
  // puedeSerPrincipal(formato): un literal (código de Markdown) no puede: se repone tal cual y duplicaría el texto.
  function armar(tramos, puedeSerPrincipal = () => true) {
    let ts = juntar(tramos);
    const plano = ts.map((t) => t.texto).join('');
    const inicio = plano.match(/^\s*/)[0], fin = plano.slice(inicio.length).match(/\s*$/)[0];
    const letras = new Map([['', 0]]);
    for (const t of ts) if (puedeSerPrincipal(t.formato)) letras.set(t.formato, (letras.get(t.formato) || 0) + t.texto.replace(/\s/g, '').length);
    const principal = [...letras].sort((a, b) => b[1] - a[1])[0][0];
    // Un tramo de solo espacios con otro formato (un espacio subrayado) pasa al formato principal.
    ts = juntar(ts.map((t) => (t.texto.trim() ? t : { texto: t.texto, formato: principal })));
    const marcas = [];
    const bordes = [];        // por marca: si los espacios de su borde eran de ella (antes de sacarlos fuera de la marca)
    let texto = '';
    for (const t of ts) {
      if (t.formato === principal) { texto += t.texto; continue; }
      const a = t.texto.match(/^\s*/)[0], z = t.texto.slice(a.length).match(/\s*$/)[0];
      const medio = t.texto.slice(a.length, t.texto.length - z.length);
      marcas.push(t.formato);
      bordes.push({ antes: !!a, despues: !!z });
      const k = marcas.length;
      texto += `${a}<${k}>${medio}</${k}>${z}`;
    }
    texto = texto.trim();
    if (marcas.length > MAX_MARCAS) return { texto: plano.trim(), marcas: [], bordes: [], principal, inicio, fin, sinMarcas: true };
    return { texto, marcas, bordes, principal, inicio, fin };
  }

  const RE_MARCA = /<(\/?)([1-8])>/g;
  const quitarMarcas = (s) => s.replace(RE_MARCA, '').replace(/[ \t]{2,}/g, ' ').trim();

  // Traducción con marcas -> tramos, o null si las marcas no cierran exactamente como se mandaron.
  function desarmar(salida, info) {
    const tramos = [];
    const abiertas = [];
    const vistas = new Set();
    let ultimo = 0, m;
    RE_MARCA.lastIndex = 0;
    const actual = (texto) => {
      const k = abiertas.length ? abiertas[abiertas.length - 1] : undefined;
      return { texto, formato: k ? info.marcas[k - 1] : info.principal, k };
    };
    while ((m = RE_MARCA.exec(salida))) {
      const k = Number(m[2]);
      if (k > info.marcas.length) return null;
      if (m.index > ultimo) tramos.push(actual(salida.slice(ultimo, m.index)));
      if (!m[1]) {
        if (vistas.has(k)) return null;
        vistas.add(k); abiertas.push(k);
      } else {
        if (abiertas[abiertas.length - 1] !== k) return null;
        abiertas.pop();
      }
      ultimo = m.index + m[0].length;
    }
    if (abiertas.length || vistas.size !== info.marcas.length) return null;
    if (ultimo < salida.length) tramos.push(actual(salida.slice(ultimo)));
    // Espacios pegados a las marcas («<1> rojo </1>»): afuera, con el formato principal.
    // Sin espacios dobles donde se juntan: «y » + « b» -> «y b».
    const r = [];
    const poner = (texto, formato, k) => {
      const u = r[r.length - 1];
      if (u && /\s$/.test(u.texto)) texto = texto.replace(/^\s+/, '');
      if (texto) r.push({ texto, formato, k });
    };
    for (const t of tramos) {
      if (t.k === undefined) { poner(t.texto, t.formato); continue; }
      const a = t.texto.match(/^\s*/)[0], z = t.texto.slice(a.length).match(/\s*$/)[0];
      if (a) poner(a, info.principal);
      if (t.texto.trim()) poner(t.texto.trim(), t.formato, t.k);
      if (z) poner(z, info.principal);
    }
    // Cada espacio pegado a una marca vuelve al formato que tenía en el original: si en «Leé <a>el capítulo</a> después» el enlace es el
    // principal, los espacios eran del texto normal (marcado) y no deben quedar subrayados dentro del enlace.
    const bordes = info.bordes || [];
    for (let i = 0; i < r.length; i++) {
      const t = r[i];
      if (t.k !== undefined) continue;
      const prev = r[i - 1], sig = r[i + 1];
      if (prev && prev.k !== undefined && bordes[prev.k - 1]?.despues) {
        const a = t.texto.match(/^\s*/)[0];
        if (a) { prev.texto += a; t.texto = t.texto.slice(a.length); }
      }
      if (sig && sig.k !== undefined && bordes[sig.k - 1]?.antes) {
        const z = t.texto.match(/\s*$/)[0];
        if (z) { sig.texto = z + sig.texto; t.texto = t.texto.slice(0, t.texto.length - z.length); }
      }
    }
    return juntar(r.filter((t) => t.texto).map(({ texto, formato }) => ({ texto, formato })));
  }

  // ───────── Traducir unidades ─────────
  // unidades: [[tramos]] -> [[tramos traducidos]]. Una unidad sin letras (solo espacios, números sueltos o signos) vuelve igual.
  async function traducirUnidades(unidades, { traducir, destino, progreso, puedeSerPrincipal }) {
    const SEP = SIN_ESPACIO.has(destino) ? '' : ' ';
    const trabajos = [];     // { u, info, contexto }
    // Contexto: la oración anterior, también la última del párrafo anterior (en un documento, los párrafos seguidos hablan de lo mismo).
    let anterior;
    const oracionesPorUnidad = unidades.map((tramos, u) => {
      const plano = tramos.map((t) => t.texto).join('');
      if (!/\p{L}/u.test(plano)) return null;
      const ors = partirEnOraciones(tramos).map((o) => armar(o, puedeSerPrincipal));
      ors.forEach((info) => {
        trabajos.push({ u, info, contexto: anterior && anterior.length <= 400 ? anterior : undefined });
        anterior = quitarMarcas(info.texto);
      });
      return ors;
    });
    const total = trabajos.length;
    let hechos = 0;
    const avisar = progreso ? () => progreso(++hechos, total) : () => {};
    const salidas = await traducir(trabajos.map((t) => ({ texto: t.info.texto, contexto: t.contexto })), destino, avisar);
    // Plan B para las que volvieron con las marcas rotas: de nuevo, sin marcas.
    const rotas = [];
    const resultado = trabajos.map((t, i) => {
      const tramos = t.info.marcas.length ? desarmar(salidas[i], t.info) : [{ texto: salidas[i], formato: t.info.principal }];
      if (!tramos) rotas.push(i);
      return tramos;
    });
    if (rotas.length) {
      const otra = await traducir(rotas.map((i) => ({ texto: quitarMarcas(trabajos[i].info.texto), contexto: trabajos[i].contexto })), destino, () => {});
      rotas.forEach((i, j) => { resultado[i] = [{ texto: otra[j], formato: trabajos[i].info.principal }]; });
    }
    // Rearmar cada unidad: oraciones juntas con el separador del idioma, y los espacios del borde originales.
    const porUnidad = unidades.map(() => []);
    trabajos.forEach((t, i) => porUnidad[t.u].push({ tramos: resultado[i], info: t.info }));
    const estadistica = { oraciones: total, planB: rotas.length };
    const final = unidades.map((tramos, u) => {
      if (!oracionesPorUnidad[u]) return tramos;
      const partes = porUnidad[u];
      const r = [];
      partes.forEach((p, i) => {
        const ts = p.tramos.map((t) => ({ ...t, texto: t.texto }));
        if (ts.length) { ts[0].texto = ts[0].texto.replace(/^\s+/, ''); ts[ts.length - 1].texto = ts[ts.length - 1].texto.replace(/\s+$/, ''); }
        if (i === 0) r.push({ texto: p.info.inicio, formato: p.info.principal });
        else r.push({ texto: SEP, formato: p.info.principal });
        r.push(...ts);
        if (i === partes.length - 1) r.push({ texto: p.info.fin, formato: p.info.principal });
      });
      return juntar(r);
    });
    final.estadistica = estadistica;
    return final;
  }

  // Envoltura de traducir(): el que llama da (segmentos, destino) -> textos; acá se le suma el aviso de progreso por segmento.
  function conProgreso(traducir) {
    return async (segmentos, destino, avisar) => {
      const r = await traducir(segmentos, destino, { progreso: avisar });
      return r;
    };
  }

  // ───────── TXT ─────────
  // Cada línea es una unidad, salvo los párrafos «cortados a mano» (líneas largas seguidas que no terminan en punto): esas se juntan en una
  // unidad y la traducción se vuelve a cortar al mismo ancho. Se conservan los saltos (\n o \r\n), las sangrías y las líneas vacías.
  function bloquesDeTexto(texto) {
    const salto = /\r\n/.test(texto) ? '\r\n' : '\n';
    const lineas = texto.split(/\r?\n/);
    const bloques = [];      // { lineas: [texto], sangria, ancho, junto }
    let i = 0;
    while (i < lineas.length) {
      const l = lineas[i];
      // Párrafo cortado a mano: 3 o más líneas seguidas de 40+ caracteres, la anterior sin terminar en signo de cierre.
      let j = i;
      while (j + 1 < lineas.length && lineas[j].trim().length >= 40 && lineas[j + 1].trim() && !/[.!?:;。！？]["'”’»)]*\s*$/.test(lineas[j])
             && lineas[j + 1].match(/^\s*/)[0] === lineas[i].match(/^\s*/)[0]) j++;
      if (j - i >= 2) {
        const grupo = lineas.slice(i, j + 1);
        bloques.push({ lineas: grupo, sangria: l.match(/^\s*/)[0], ancho: Math.max(...grupo.map((x) => x.length)), junto: true });
        i = j + 1;
      } else {
        bloques.push({ lineas: [l], junto: false });
        i++;
      }
    }
    return { bloques, salto };
  }

  function cortarAlAncho(texto, ancho, sangria) {
    const palabras = texto.split(/(\s+)/).filter((x) => x.length);
    const lineas = [];
    let actual = '';
    for (const p of palabras) {
      if (/^\s+$/.test(p)) { if (actual) actual += ' '; continue; }
      if (actual && (sangria + actual + p).length > ancho) { lineas.push(sangria + actual.trimEnd()); actual = ''; }
      actual += p;
    }
    if (actual) lineas.push(sangria + actual.trimEnd());
    return lineas;
  }

  async function traducirTexto(texto, { traducir, destino, progreso }) {
    const { bloques, salto } = bloquesDeTexto(String(texto));
    const unidades = bloques.map((b) => [{ texto: b.junto ? b.lineas.map((x) => x.trim()).join(' ') : b.lineas[0], formato: '' }]);
    const r = await traducirUnidades(unidades, { traducir: conProgreso(traducir), destino, progreso });
    const lineas = [];
    bloques.forEach((b, i) => {
      const t = r[i].map((x) => x.texto).join('');
      if (!b.junto) lineas.push(t);
      else lineas.push(...(SIN_ESPACIO.has(destino) ? [b.sangria + t.trim()] : cortarAlAncho(t.trim(), b.ancho, b.sangria)));
    });
    const salida = lineas.join(salto);
    return { texto: salida, estadistica: r.estadistica };
  }

  // ───────── Markdown ─────────
  // Se traduce solo el texto. Quedan intactos: bloques de código (``` y sangría de 4), tablas (solo se traducen las celdas), el marcado
  // de títulos, listas, citas y casillas, las direcciones de los enlaces e imágenes, el código en línea, el HTML y las referencias [x]: url.
  // El formato en línea viaja como marcas: **negrita**, *cursiva*, ~~tachado~~, `código` (que vuelve tal cual), [enlace](url).
  const RE_INLINE = /(\*\*\*|___)(?=\S)([\s\S]+?)(?<=\S)\1|(\*\*|__)(?=\S)([\s\S]+?)(?<=\S)\3|(\*|_)(?=\S)([\s\S]+?)(?<=\S)\5|~~(?=\S)([\s\S]+?)(?<=\S)~~|(`+)([\s\S]+?)\8|(!?)\[([^\]]*)\]\(([^)\s]+(?:\s+"[^"]*")?)\)|<[^>\n]+>|https?:\/\/[^\s<>()]+/g;

  // Texto en línea -> tramos. formato = JSON de { estilo, url, literal } para poder rearmar.
  function tramosEnLinea(texto) {
    const tramos = [];
    let ultimo = 0, m;
    RE_INLINE.lastIndex = 0;
    const pon = (t, f) => { if (t) tramos.push({ texto: t, formato: f }); };
    while ((m = RE_INLINE.exec(texto))) {
      pon(texto.slice(ultimo, m.index), '');
      if (m[1]) pon(m[2], JSON.stringify({ e: m[1] }));
      else if (m[3]) pon(m[4], JSON.stringify({ e: m[3] }));
      else if (m[5]) {
        // «_» dentro de una palabra (nombre_de_archivo) no es cursiva.
        if (m[5] === '_' && (/\w/.test(texto[m.index - 1] || '') || /\w/.test(texto[m.index + m[0].length] || ''))) { pon(m[0], ''); }
        else pon(m[6], JSON.stringify({ e: m[5] }));
      } else if (m[7]) pon(m[7], JSON.stringify({ e: '~~' }));
      else if (m[8]) pon(m[0], JSON.stringify({ literal: m[0] }));                          // código: vuelve tal cual
      else if (m[11] !== undefined) {
        if (m[10] === '!') pon(m[0], JSON.stringify({ literal: m[0] }));                    // imagen: igual (el texto alternativo, sin traducir)
        else pon(m[11] || ' ', JSON.stringify({ url: m[12] }));
      } else pon(m[0], JSON.stringify({ literal: m[0] }));                                  // HTML en línea y direcciones sueltas
      ultimo = m.index + m[0].length;
    }
    pon(texto.slice(ultimo), '');
    return tramos;
  }

  // Los espacios del borde de un tramo van fuera de los delimitadores: «_ suave_» no es cursiva en Markdown.
  function rearmarEnLinea(tramos) {
    return tramos.map((t) => {
      if (!t.formato) return t.texto;
      const f = JSON.parse(t.formato);
      if (f.literal !== undefined) return f.literal;
      const a = t.texto.match(/^\s*/)[0], z = t.texto.slice(a.length).match(/\s*$/)[0];
      const medio = t.texto.slice(a.length, t.texto.length - z.length);
      if (!medio) return t.texto;
      if (f.url !== undefined) return `${a}[${medio}](${f.url})${z}`;
      return `${a}${f.e}${medio}${f.e}${z}`;
    }).join('');
  }
  const noEsLiteral = (formato) => !formato || JSON.parse(formato).literal === undefined;

  // Los literales (código, direcciones, HTML) se mandan como marca con su texto; al volver, se reponen aunque el modelo los haya tocado.
  function reponerLiterales(original, traducidos) {
    return traducidos.map((t) => (t.formato && JSON.parse(t.formato).literal !== undefined ? { ...t, texto: JSON.parse(t.formato).literal } : t));
  }

  async function traducirMarkdown(texto, { traducir, destino, progreso }) {
    const salto = /\r\n/.test(texto) ? '\r\n' : '\n';
    const lineas = String(texto).split(/\r?\n/);
    const piezas = [];       // { tipo: 'fija', linea } | { tipo: 'texto', prefijo, contenido, sufijo } | { tipo: 'tabla', celdas, separadores }
    let enCodigo = null;
    for (let i = 0; i < lineas.length; i++) {
      const l = lineas[i];
      const valla = l.match(/^\s*(```+|~~~+)/);
      if (enCodigo) { piezas.push({ tipo: 'fija', linea: l }); if (valla && valla[1][0] === enCodigo[0] && valla[1].length >= enCodigo.length) enCodigo = null; continue; }
      if (valla) { enCodigo = valla[1]; piezas.push({ tipo: 'fija', linea: l }); continue; }
      if (!l.trim() || /^( {4}|\t)/.test(l) || /^\s*([-*_])(\s*\1){2,}\s*$/.test(l) || /^\s*\[[^\]]+\]:\s*\S/.test(l)
          || /^\s*<\/?[a-zA-Z][^>]*>\s*$/.test(l) || /^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)*\|?\s*$/.test(l)) {
        piezas.push({ tipo: 'fija', linea: l }); continue;
      }
      if (/^\s*\|.*\|\s*$/.test(l)) {
        // Fila de tabla: se traduce cada celda.
        const partes = l.split(/(?<!\\)\|/);
        piezas.push({ tipo: 'tabla', partes });
        continue;
      }
      const pre = l.match(/^(\s*(?:>\s*)*(?:#{1,6}\s+|[-*+]\s+(?:\[[ xX]\]\s+)?|\d+[.)]\s+(?:\[[ xX]\]\s+)?)?)/)[0];
      const resto = l.slice(pre.length);
      const cierre = pre.includes('#') ? (resto.match(/\s+#+\s*$/) || [''])[0] : '';
      piezas.push({ tipo: 'texto', prefijo: pre, contenido: resto.slice(0, resto.length - cierre.length), sufijo: cierre });
    }
    // Unidades: cada línea de texto y cada celda de tabla.
    const unidades = [], destinoDe = [];
    piezas.forEach((p, i) => {
      if (p.tipo === 'texto') { unidades.push(tramosEnLinea(p.contenido)); destinoDe.push([i]); }
      else if (p.tipo === 'tabla') p.partes.forEach((c, j) => { if (c.trim()) { unidades.push(tramosEnLinea(c)); destinoDe.push([i, j]); } });
    });
    const r = await traducirUnidades(unidades, { traducir: conProgreso(traducir), destino, progreso, puedeSerPrincipal: noEsLiteral });
    r.forEach((tramos, k) => {
      const [i, j] = destinoDe[k];
      const texto = rearmarEnLinea(reponerLiterales(unidades[k], tramos));
      if (j === undefined) piezas[i].contenido = texto; else piezas[i].partes[j] = texto;
    });
    const salida = piezas.map((p) => (p.tipo === 'fija' ? p.linea : p.tipo === 'tabla' ? p.partes.join('|') : p.prefijo + p.contenido + p.sufijo)).join(salto);
    return { texto: salida, estadistica: r.estadistica };
  }

  raiz.TraductorDocumentos = {
    traducirTexto, traducirMarkdown, traducirUnidades,
    _nucleo: { partirEnOraciones, cortesDeOracion, armar, desarmar, quitarMarcas, tramosEnLinea, rearmarEnLinea, bloquesDeTexto },
  };
})(typeof self !== 'undefined' ? self : globalThis);

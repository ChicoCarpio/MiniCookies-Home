/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Pegar como imagen (Pegar > Imagen en Excel): lo copiado se dibuja como se ve, sin la cuadrícula, y queda como una imagen anclada en la celda de destino,
// que se mueve con las celdas y no cambia de tamaño. Medido en Excel (pruebas/oraculo/excel-pegar-imagen.ps1): «Picture N», del tamaño del rango más un píxel.
// Excel la guarda en EMF; aquí va en PNG, como su «Pegar como imagen (PNG)». En .ods, un marco con la imagen anclado a la celda, como lo deja LibreOffice.
//   await PegadoImagen.dibujar(P, idioma) → { png, ancho, alto }      (P: lo copiado, PegadoEspecial.capturar; px a 96 dpi; necesita un canvas)
//   await PegadoImagen.agregar(bytes, formato, h, f, c, imagen) → bytes
(function () {
  const X = () => XmlEditar;
  const EMU_PX = 9525;
  const REL_DIBUJO = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
  const REL_IMAGEN = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

  async function dibujar(P, idioma) {
    const libro = P.ed.libro, hoja = P.ed.hojas[P.h];
    const r = { f0: P.filas[0], f1: P.filas[P.filas.length - 1], c0: P.r.c0, c1: Math.min(P.r.c1, P.r.c0 + P.celdas[0].length - 1) };
    await XlsxDibujo.cargarFuentes(libro);
    const g = XlsxDibujo.geometria(libro, hoja);
    const ancho = Math.max(1, g.x(r.c1 + 1) - g.x(r.c0) + 1), alto = Math.max(1, g.y(r.f1 + 1) - g.y(r.f0) + 1);
    if (ancho * alto > 4096 * 4096) throw new Error('grande');
    const lienzo = document.createElement('canvas');
    lienzo.width = ancho; lienzo.height = alto;
    const ctx = lienzo.getContext('2d');
    ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, ancho, alto);
    XlsxDibujo.dibujarRango(ctx, libro, hoja, g, r, { x: 0, y: 0, idioma: idioma === 'en' ? 'en' : 'es', cuadricula: false });
    const blob = await new Promise((res) => lienzo.toBlob(res, 'image/png'));
    return { png: new Uint8Array(await blob.arrayBuffer()), ancho, alto };
  }

  async function agregar(bytes, formato, h, f, c, imagen) {
    return formato === 'ods' ? agregarOds(bytes, h, f, c, imagen) : agregarXlsx(bytes, h, f, c, imagen);
  }
  const libreCon = (P, plantilla) => { let k = 1; while (P.existe(plantilla(k))) k++; return plantilla(k); };

  async function agregarXlsx(bytes, h, f, c, imagen) {
    const libro = await LibroXlsx.abrir(bytes), P = PaqueteEditable.desdeZip(libro.zip), ruta = libro.hojas[h].ruta;
    const media = libreCon(P, (k) => 'xl/media/image' + k + '.png');
    P.escribir(media, imagen.png);
    await P.asegurarTipo('/' + media, 'image/png');
    let rel = (await P.relaciones(ruta)).find((x) => x.tipo === REL_DIBUJO && !x.externa);
    let dibujo = rel ? rel.destino : null;
    if (!dibujo) {
      dibujo = libreCon(P, (k) => 'xl/drawings/drawing' + k + '.xml');
      P.escribir(dibujo, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>');
      await P.asegurarTipo('/' + dibujo, 'application/vnd.openxmlformats-officedocument.drawing+xml');
      const id = await P.agregarRelacion(ruta, REL_DIBUJO, '../' + dibujo.replace(/^xl\//, ''));
      const con = MoverNotas._conPrefijoR(await P.leer(ruta));
      P.escribir(ruta, EstructuraLibro._ponerEnHoja(con.xml, 'drawing', `<drawing ${con.p}:id="${id}"/>`));
    }
    const idImagen = await P.agregarRelacion(dibujo, REL_IMAGEN, '../media/' + media.split('/').pop());
    let xml = await P.leer(dibujo);
    const raiz = /<([A-Za-z_][\w.-]*:)?wsDr\b[^>]*>/.exec(xml);
    if (!raiz) throw new Error('el dibujo de la hoja no se entiende');
    const px = raiz[1] || '';
    // el prefijo de DrawingML (a:) y el de las relaciones (r:), los que ya declara el dibujo o los de siempre
    const pre = (ns, pref) => { const m = new RegExp('xmlns:([A-Za-z_][\\w.-]*)="' + ns.replace(/[.]/g, '\\.') + '"').exec(xml); if (m) return m[1]; xml = xml.replace(raiz[0], raiz[0].replace(/>$/, ` xmlns:${pref}="${ns}">`)); return pref; };
    const a = pre('http://schemas.openxmlformats.org/drawingml/2006/main', 'a');
    const rr = pre('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'r');
    let id = 1;
    for (const m of xml.matchAll(/<(?:[\w.-]+:)?cNvPr\b[^>]*?\bid="(\d+)"/g)) id = Math.max(id, +m[1] + 1);
    const n = (xml.match(/<(?:[\w.-]+:)?pic\b/g) || []).length + 1;
    const cx = imagen.ancho * EMU_PX, cy = imagen.alto * EMU_PX;
    const ancla = `<${px}oneCellAnchor><${px}from><${px}col>${c}</${px}col><${px}colOff>0</${px}colOff><${px}row>${f}</${px}row><${px}rowOff>0</${px}rowOff></${px}from><${px}ext cx="${cx}" cy="${cy}"/>`
      + `<${px}pic><${px}nvPicPr><${px}cNvPr id="${id}" name="Picture ${n}"/><${px}cNvPicPr><${a}:picLocks noChangeAspect="1"/></${px}cNvPicPr></${px}nvPicPr>`
      + `<${px}blipFill><${a}:blip ${rr}:embed="${idImagen}"/><${a}:stretch><${a}:fillRect/></${a}:stretch></${px}blipFill>`
      + `<${px}spPr><${a}:xfrm><${a}:off x="0" y="0"/><${a}:ext cx="${cx}" cy="${cy}"/></${a}:xfrm><${a}:prstGeom prst="rect"><${a}:avLst/></${a}:prstGeom></${px}spPr></${px}pic><${px}clientData/></${px}oneCellAnchor>`;
    xml = /<\/([A-Za-z_][\w.-]*:)?wsDr>\s*$/.test(xml) ? xml.replace(/<\/([A-Za-z_][\w.-]*:)?wsDr>\s*$/, (m) => ancla + m) : xml.replace(/\/>\s*$/, '>' + ancla + `</${px}wsDr>`);
    P.escribir(dibujo, xml);
    return libro.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  // .ods: el marco va dentro de la celda de destino (anclado a la celda, como LibreOffice), con la imagen en Pictures/ y su entrada en el manifiesto
  async function agregarOds(bytes, h, f, c, imagen) {
    const P = await PaqueteEditable.abrir(bytes);
    const nombre = libreCon(P, (k) => 'Pictures/imagen' + k + '.png');
    P.escribir(nombre, imagen.png);
    let man = await P.leer('META-INF/manifest.xml');
    if (man !== null) { man = man.replace(/<\/manifest:manifest>\s*$/, `<manifest:file-entry manifest:full-path="${nombre}" manifest:media-type="image/png"/></manifest:manifest>`); P.escribir('META-INF/manifest.xml', man); }
    const ancho = (imagen.ancho * 2.54 / 96).toFixed(3) + 'cm', alto = (imagen.alto * 2.54 / 96).toFixed(3) + 'cm';
    let texto = await P.leer('content.xml');
    const n = (texto.match(/<draw:frame\b/g) || []).length + 1;
    const marco = `<draw:frame draw:z-index="${n - 1}" draw:name="Imagen ${n}" svg:width="${ancho}" svg:height="${alto}" svg:x="0cm" svg:y="0cm"><draw:image xlink:href="${nombre}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"><text:p/></draw:image></draw:frame>`;
    texto = ponerMarcoOds(texto, h, f, c, marco);
    if (!/xmlns:xlink=/.test(texto.slice(0, texto.indexOf('>', texto.indexOf('<office:document-content'))))) texto = texto.replace(/<office:document-content\b/, '<office:document-content xmlns:xlink="http://www.w3.org/1999/xlink"');
    P.escribir('content.xml', texto);
    return P.generar();
  }
  // pone el marco en la celda (h, f, c) del content.xml: parte las filas y celdas repetidas para que esa celda quede sola
  function ponerMarcoOds(texto, h, f, c, marco) {
    const U = OdsEscritorCeldas._util;
    const raiz = X().analizar(texto, { profundidad: 5 });
    const hojaCalc = X().hijo(X().hijo(raiz, 'body'), 'spreadsheet');
    const tabla = X().hijos(hojaCalc, 'table')[h];
    if (!tabla) throw new Error('la hoja no existe');
    const filas = U.enumerarFilas(texto, tabla);
    const fila = filas.find((x) => x.f0 <= f && f < x.f0 + x.n);
    if (!fila) throw new Error('la celda está fuera de la hoja');
    // la fila sola
    const antes = f > fila.f0 ? U.copiaConRepeticion(texto, fila.nodo, 'table:number-rows-repeated', f - fila.f0) : '';
    const despues = fila.f0 + fila.n - f - 1 > 0 ? U.copiaConRepeticion(texto, fila.nodo, 'table:number-rows-repeated', fila.f0 + fila.n - f - 1) : '';
    const una = U.copiaConRepeticion(texto, fila.nodo, 'table:number-rows-repeated', 1);
    // la celda sola, con el marco adentro
    const r = X().analizar(una, { profundidad: 2 });
    let col = 0, nueva = null;
    for (const x of r.hijos) {
      if (x.local !== 'table-cell' && x.local !== 'covered-table-cell') continue;
      const k = U.repeticiones(x, 'table:number-columns-repeated');
      if (c >= col && c < col + k) {
        const izq = c - col, der = col + k - c - 1;
        const sola = U.copiaConRepeticion(una, x, 'table:number-columns-repeated', 1);
        const conMarco = /\/>$/.test(sola) ? sola.replace(/\/>$/, '>' + marco + '</table:' + x.local + '>') : sola.replace(/(<table:(?:covered-)?table-cell\b[^>]*>)/, (m) => m + marco);
        nueva = una.slice(0, x.ini) + (izq ? U.copiaConRepeticion(una, x, 'table:number-columns-repeated', izq) : '') + conMarco + (der ? U.copiaConRepeticion(una, x, 'table:number-columns-repeated', der) : '') + una.slice(x.fin);
        break;
      }
      col += k;
    }
    if (nueva === null) throw new Error('la celda está fuera de la hoja');
    return texto.slice(0, fila.nodo.ini) + antes + nueva + despues + texto.slice(fila.nodo.fin);
  }

  globalThis.PegadoImagen = { dibujar, agregar };
})();

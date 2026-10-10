/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Maquetado del texto de una forma de PowerPoint: corta las líneas, aplica
// interlineado, espacio entre párrafos, sangrías, viñetas y numeración,
// alineación (también justificado) y anclaje vertical, y devuelve cada trozo
// con su posición. Lo dibuja pptx-dibujo.js.
//
// Mide con el navegador (measureText), pero con fuentes de vendor/ que tienen
// los mismos anchos que las de Office (METRICAS, etapa 3): así el corte de
// líneas no depende de las fuentes de la PC. Se calibra contra los cortes de
// línea que da PowerPoint por COM (npm run comparar-pptx -- --motor-propio,
// "Cortes de línea"). Lo que todavía no está calibrado está marcado
// "[calibrar]". Necesita assets/js/vendor.js (rutaVendor) cargado antes.
//
// Uso: const m = PptxTexto.maquetar(texto, anchoPx, altoPx, pxPorEmu, ctx2d)
//   texto: { cuerpo, parrafos } de pptx-modelo.js; ancho/alto: el área de
//   texto ya sin márgenes. m: { lineas, alto }.

const PptxTexto = (() => {
  const EMU_POR_PT = 12700;
  // Altura de una línea con interlineado sencillo, en veces el tamaño. [calibrar]
  const ALTO_LINEA = 1.2;
  // Wingdings y Symbol: el carácter guardado es la posición en esa fuente.
  const SIMBOLOS = {
    wingdings: { v: '❖', 'Ø': '➢', '§': '▪', 'ü': '✔', q: '❑', n: '■', l: '●', 'Ÿ': '•', 'à': '➔', 'è': '➜', 'ê': '➢', 'ð': '⇒', o: '□', p: '◻', 'ª': '✦', w: '⬥', u: '◆', s: '⬧', t: '⧫', 'Ø': '➢', 'ö': '⬧', 'û': '✗', 'ý': '☒', 'þ': '☑' },
    symbol: { '·': '•', 'Þ': '⇒', '®': '→', 'à': '◊', '¨': '♣', '©': '♦', 'ª': '♥', '«': '♠', '-': '−', 'Ö': '√' },
  };
  const FUENTES_SIMBOLO = '"Segoe UI Symbol", "Segoe UI Emoji", sans-serif';

  // Fuentes que el navegador no tiene: la sustituta más parecida en Windows. [calibrar]
  const SUSTITUTAS = { aptos: 'Segoe UI', 'aptos display': 'Segoe UI', 'aptos narrow': 'Arial Narrow' };

  // Fuentes libres con los mismos anchos de letra que las de Office, servidas
  // desde vendor/: si la PC no tiene la original, el navegador usa esta y las
  // líneas se cortan igual que en PowerPoint. Comprobado el 2026-10-10 contra
  // las de Windows (anchos de 317 letras y una frase con kerning): Carlito =
  // Calibri y Liberation = Arial / Times New Roman / Courier New, idénticas;
  // Gelasio ≈ Georgia (la frase igual, 7 letras raras distintas). Caladea NO
  // coincide con Cambria (5 % más angosta): no se usa. Intos = Aptos, Aptos
  // Display y Aptos Narrow (anchos, kerning y alto idénticos a la Aptos de
  // Office en 11 estilos; la de Office viene de la nube, ninguna PC sin
  // Office la tiene). Intos sin versiones publicadas: fijada en el commit.
  // Estilos: r normal, b negrita, i cursiva, bi negrita cursiva.
  const METRICAS = {
    calibri: { familia: 'MC Carlito', r: rutaVendor('carlito/1.104/Carlito-Regular.ttf'), b: rutaVendor('carlito/1.104/Carlito-Bold.ttf'), i: rutaVendor('carlito/1.104/Carlito-Italic.ttf'), bi: rutaVendor('carlito/1.104/Carlito-BoldItalic.ttf') },
    arial: { familia: 'MC Liberation Sans', r: rutaVendor('liberation-fonts/2.1.5/LiberationSans-Regular.ttf'), b: rutaVendor('liberation-fonts/2.1.5/LiberationSans-Bold.ttf'), i: rutaVendor('liberation-fonts/2.1.5/LiberationSans-Italic.ttf'), bi: rutaVendor('liberation-fonts/2.1.5/LiberationSans-BoldItalic.ttf') },
    'times new roman': { familia: 'MC Liberation Serif', r: rutaVendor('liberation-fonts/2.1.5/LiberationSerif-Regular.ttf'), b: rutaVendor('liberation-fonts/2.1.5/LiberationSerif-Bold.ttf'), i: rutaVendor('liberation-fonts/2.1.5/LiberationSerif-Italic.ttf'), bi: rutaVendor('liberation-fonts/2.1.5/LiberationSerif-BoldItalic.ttf') },
    'courier new': { familia: 'MC Liberation Mono', r: rutaVendor('liberation-fonts/2.1.5/LiberationMono-Regular.ttf'), b: rutaVendor('liberation-fonts/2.1.5/LiberationMono-Bold.ttf'), i: rutaVendor('liberation-fonts/2.1.5/LiberationMono-Italic.ttf'), bi: rutaVendor('liberation-fonts/2.1.5/LiberationMono-BoldItalic.ttf') },
    aptos: { familia: 'MC Intos', r: rutaVendor('intos/fef9315/Intos-Regular.ttf'), b: rutaVendor('intos/fef9315/Intos-Bold.ttf'), i: rutaVendor('intos/fef9315/Intos-Italic.ttf'), bi: rutaVendor('intos/fef9315/Intos-BoldItalic.ttf') },
    'aptos display': { familia: 'MC Intos Display', r: rutaVendor('intos/fef9315/IntosDisplay-Regular.ttf'), b: rutaVendor('intos/fef9315/IntosDisplay-Bold.ttf'), i: rutaVendor('intos/fef9315/IntosDisplay-Italic.ttf'), bi: rutaVendor('intos/fef9315/IntosDisplay-BoldItalic.ttf') },
    'aptos narrow': { familia: 'MC Intos Narrow', r: rutaVendor('intos/fef9315/IntosNarrow-Regular.ttf'), b: rutaVendor('intos/fef9315/IntosNarrow-Bold.ttf'), i: rutaVendor('intos/fef9315/IntosNarrow-Italic.ttf'), bi: rutaVendor('intos/fef9315/IntosNarrow-BoldItalic.ttf') },
    // Variable (peso 400 a 700): el mismo archivo para normal y negrita.
    georgia: { familia: 'MC Gelasio', variable: true, r: rutaVendor('gelasio/1.008/Gelasio-VF.ttf'), i: rutaVendor('gelasio/1.008/Gelasio-Italic-VF.ttf') },
  };
  const cargadas = new Map();
  // Carga (una sola vez) los estilos que se van a usar. Hay que esperarla antes
  // de medir: si la fuente no terminó de cargar, el canvas mide con otra sin
  // avisar. Sin FontFace o si falla, queda la cadena de familia() como antes.
  // usos: [{ fuente, negrita, cursiva }].
  async function cargarFuentes(usos) {
    if (typeof FontFace === 'undefined' || !document.fonts) return;
    const pendientes = [];
    for (const u of usos) {
      const m = METRICAS[(u.fuente || 'Calibri').replace(/"/g, '').toLowerCase()];
      if (!m) continue;
      const clave = m.variable ? (u.cursiva ? 'i' : 'r') : (u.negrita ? 'b' : '') + (u.cursiva ? 'i' : '') || 'r';
      const id = m.familia + ':' + clave;
      if (!cargadas.has(id)) {
        const cara = new FontFace(m.familia, `url("${m[clave]}")`, {
          weight: m.variable ? '400 700' : clave.includes('b') ? '700' : '400', style: clave.includes('i') ? 'italic' : 'normal',
        });
        cargadas.set(id, cara.load().then((c) => { document.fonts.add(c); }).catch((e) => console.warn('No cargó la fuente', id, e)));
      }
      pendientes.push(cargadas.get(id));
    }
    await Promise.all(pendientes);
  }
  function familia(nombre) {
    const n = (nombre || 'Calibri').replace(/"/g, '');
    const sust = SUSTITUTAS[n.toLowerCase()];
    const metricas = METRICAS[n.toLowerCase()];
    return `"${n}"${metricas ? `, "${metricas.familia}"` : ''}${sust ? `, "${sust}"` : ''}, "Segoe UI", sans-serif`;
  }
  const css = (t, px) => `${t.cursiva ? 'italic ' : ''}${t.negrita ? 'bold ' : ''}${px.toFixed(3)}px ${familia(t.fuente)}`;
  const colorCss = (c) => c ? `rgba(${c.r},${c.g},${c.b},${c.a})` : null;
  function colorDeRelleno(r) {
    if (!r) return null;
    if (r.tipo === 'solido') return r.color;
    if (r.tipo === 'degradado') return r.paradas[0].color;
    if (r.tipo === 'trama') return r.fg;
    return null;
  }

  // --- Numeración ------------------------------------------------------------------
  function romano(n) {
    const t = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
    let s = ''; for (const [v, l] of t) while (n >= v) { s += l; n -= v; } return s;
  }
  function alfa(n) { let s = ''; while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); } return s; }
  function numerar(esquema, n) {
    const m = /^(arabic|romanUc|romanLc|alphaUc|alphaLc)(Period|ParenR|ParenBoth|Plain|Minus)?/.exec(esquema || '');
    let base = String(n);
    if (m) {
      if (m[1] === 'romanUc') base = romano(n); else if (m[1] === 'romanLc') base = romano(n).toLowerCase();
      else if (m[1] === 'alphaUc') base = alfa(n); else if (m[1] === 'alphaLc') base = alfa(n).toLowerCase();
      switch (m[2]) {
        case 'ParenR': return base + ')';
        case 'ParenBoth': return '(' + base + ')';
        case 'Plain': return base;
        case 'Minus': return '- ' + base + ' -';
        default: return base + '.';
      }
    }
    if (/^circleNum/.test(esquema) && n <= 20) return String.fromCharCode(0x2460 + n - 1);
    return n + '.';
  }

  // --- Maquetado -----------------------------------------------------------------------
  function maquetar(texto, ancho, alto, pxPorEmu, ctx) {
    const { cuerpo, parrafos } = texto;
    const ptPx = EMU_POR_PT * pxPorEmu;
    const auto = cuerpo.autoajuste.tipo === 'reducir' ? cuerpo.autoajuste : { escala: 1, reduccion: 0 };
    // Con fontScale, PowerPoint usa el tamaño reducido redondeado al punto
    // entero (medido: 18 pt × 92,5 % → 17; × 62,5 % → 11; 16 × 77,5 % → 12).
    const escalar = (tam) => auto.escala === 1 ? tam : Math.max(1, Math.round(tam * auto.escala));
    const sinAjuste = cuerpo.ajuste === 'none';
    const columnas = Math.max(1, cuerpo.columnas);
    const espacioCol = cuerpo.espacioColumnas * pxPorEmu;
    const anchoCol = columnas > 1 ? (ancho - espacioCol * (columnas - 1)) / columnas : ancho;
    const contadores = [];
    const lineas = [];
    let y = 0;

    // PowerPoint hace kerning solo con kern y desde ese tamaño (sin kern, nunca);
    // el canvas lo hace siempre si no se le dice.
    const kerning = (t, px) => t.kern > 0 && px / ptPx >= t.kern - 0.001 ? 'normal' : 'none';
    const medir = (t, s, px) => {
      ctx.fontKerning = kerning(t, px);
      ctx.font = css(t, px);
      ctx.letterSpacing = t.espaciado ? `${t.espaciado * ptPx}px` : '0px';
      return ctx.measureText(s).width;
    };
    const metricas = (t, px) => {
      ctx.font = css(t, px);
      const m = ctx.measureText('Hg');
      const asc = m.fontBoundingBoxAscent || px * 0.8, desc = m.fontBoundingBoxDescent || px * 0.2;
      return { asc, desc };
    };

    parrafos.forEach((p, ip) => {
      const visibles = p.trozos.filter((t) => !t.salto && t.texto);
      const vacio = !visibles.length;
      const tamDe = (t) => escalar(t.tamano) * (t.base ? 2 / 3 : 1);
      const primerTam = escalar((visibles[0] || p.fin).tamano);

      // Numeración: sigue en el mismo nivel; un nivel menor reinicia los mayores.
      contadores.length = Math.min(contadores.length, p.nivel + 1);
      let textoVineta = null;
      if (p.vineta && !vacio) {
        if (p.vineta.tipo === 'numero') {
          const c = contadores[p.nivel];
          const n = c && c.esquema === p.vineta.esquema ? c.n + 1 : p.vineta.inicio;
          contadores[p.nivel] = { esquema: p.vineta.esquema, n };
          textoVineta = numerar(p.vineta.esquema, n);
        } else if (p.vineta.tipo === 'caracter') textoVineta = p.vineta.caracter;
      } else if (!vacio) contadores[p.nivel] = null;

      // Espacio antes (no en el primer párrafo, salvo spcFirstLastPara).
      const espacio = (e) => e.pts != null ? e.pts * ptPx : e.pct * primerTam * ALTO_LINEA * ptPx; // [calibrar] base del %
      const antes = ip > 0 || cuerpo.primeroYUltimo ? espacio(p.antes) : 0;
      y += antes;

      // Trozos → palabras (con su estilo), para cortar.
      const piezas = [];
      for (const t of p.trozos) {
        if (t.salto) { piezas.push({ salto: true, t }); continue; }
        let s = t.texto;
        if (t.mayusculas === 'all' || t.mayusculas === 'small') s = s.toUpperCase();
        const px = tamDe(t) * ptPx * (t.mayusculas === 'small' ? 0.8 : 1);
        // Los tabuladores van sueltos (saltan a su tope); los demás espacios, juntos.
        for (const parte of s.split(/(\t|[^\S\t]+)/)) {
          if (!parte) continue;
          if (parte === '\t') { piezas.push({ tab: true, t, px }); continue; }
          piezas.push({ texto: parte, espacio: /^\s+$/.test(parte), t, px, ancho: medir(t, parte, px) });
        }
      }

      const margen = p.margen * pxPorEmu, sangria = p.sangria * pxPorEmu;
      // Viñeta: en marL + indent; el texto de la primera línea, en marL (o
      // después de la viñeta si no entra).
      let vineta = null;
      let inicioPrimera = margen + sangria;
      if (textoVineta != null) {
        const base = visibles[0] || p.fin;
        const px = (p.vineta.puntos != null ? escalar(p.vineta.puntos) : escalar(base.tamano) * (p.vineta.escala || 1)) * ptPx;
        let car = textoVineta, fam = familia(p.vineta.fuente || base.fuente);
        const tabla = SIMBOLOS[(p.vineta.fuente || '').toLowerCase()];
        if (tabla && p.vineta.tipo === 'caracter') { car = tabla[car] || car; fam = FUENTES_SIMBOLO; }
        else if (/wingdings|symbol|webdings/i.test(p.vineta.fuente || '')) fam = FUENTES_SIMBOLO;
        const fuenteCss = `${p.vineta.tipo === 'numero' && base.negrita ? 'bold ' : ''}${px.toFixed(3)}px ${p.vineta.tipo === 'numero' ? familia(base.fuente) : fam}`;
        ctx.font = fuenteCss; ctx.letterSpacing = '0px';
        const w = ctx.measureText(car).width;
        const x = margen + sangria;
        vineta = { texto: car, css: fuenteCss, color: colorCss(p.vineta.color || colorDeRelleno(base.relleno)), x, px };
        inicioPrimera = sangria < 0 ? Math.max(margen, x + w) : x + w + px * 0.25;
      }

      // Corte de líneas.
      const lineasP = [];
      let actual = null;
      const nueva = (primera) => {
        actual = { piezas: [], inicio: primera ? inicioPrimera : margen, ancho: 0, final: false, vineta: primera ? vineta : null };
        actual.limite = sinAjuste ? Infinity : anchoCol - actual.inicio;
        lineasP.push(actual);
      };
      nueva(true);
      const proximoTab = (x) => {
        const abs = actual.inicio + x;
        for (const tb of p.tabs) if (tb.pos * pxPorEmu > abs + 0.5) return tb.pos * pxPorEmu - actual.inicio;
        const paso = (p.tabulacion || 914400) * pxPorEmu;
        return (Math.floor(abs / paso) + 1) * paso - actual.inicio;
      };
      // Una palabra puede venir en varios trozos (un enlace, una letra en
      // negrita): solo se corta donde hay espacio, así que se mide entera.
      const esPalabra = (pz) => pz && !pz.salto && !pz.tab && !pz.espacio;
      piezas.forEach((pz, i) => {
        if (pz.salto) { actual.salto = pz.t; actual.final = true; nueva(false); return; }
        if (pz.tab) { const x = proximoTab(actual.ancho); actual.piezas.push({ ...pz, texto: '', ancho: x - actual.ancho }); actual.ancho = x; return; }
        if (pz.espacio) { actual.piezas.push(pz); actual.ancho += pz.ancho; return; }
        if (!esPalabra(piezas[i - 1])) {
          let palabra = 0;
          for (let k = i; esPalabra(piezas[k]); k++) palabra += piezas[k].ancho;
          const ocupado = anchoSinEspacioFinal(actual);
          if (actual.piezas.length && ocupado > 0 && actual.ancho + palabra > actual.limite + 0.01) nueva(false);
        }
        if (pz.ancho > actual.limite && !sinAjuste) {
          // Una palabra más ancha que la línea se corta por letras.
          let resto = pz.texto;
          while (resto) {
            let n = resto.length;
            while (n > 1 && actual.ancho + medir(pz.t, resto.slice(0, n), pz.px) > actual.limite) n--;
            const trozo = resto.slice(0, n);
            const w = medir(pz.t, trozo, pz.px);
            actual.piezas.push({ ...pz, texto: trozo, ancho: w }); actual.ancho += w;
            resto = resto.slice(n);
            if (resto) nueva(false);
          }
          return;
        }
        actual.piezas.push(pz); actual.ancho += pz.ancho;
      });
      actual.final = true;

      // Altura de cada línea y posición vertical.
      const interlineado = p.interlineado;
      lineasP.forEach((l, il) => {
        const conTexto = l.piezas.filter((pz) => !pz.espacio || l.piezas.length === 1);
        const fuentes = conTexto.length ? conTexto : [{ t: l.salto || p.fin, px: tamDe(l.salto || p.fin) * ptPx }];
        let tam = 0, asc = 0, desc = 0;
        for (const pz of fuentes) {
          const m = metricas(pz.t, pz.px);
          tam = Math.max(tam, pz.px); asc = Math.max(asc, m.asc); desc = Math.max(desc, m.desc);
        }
        if (l.vineta) tam = Math.max(tam, Math.min(l.vineta.px, tam || l.vineta.px));
        // El espacio de más (o de menos) del interlineado va arriba de la
        // línea; lo que baja de la línea de base queda fijo. [calibrar]
        const unaLinea = tam * ALTO_LINEA;
        const bajoLaBase = unaLinea * (desc / (asc + desc || 1));
        const altoLinea = interlineado.pts != null ? interlineado.pts * ptPx : unaLinea * Math.max(0, interlineado.pct - auto.reduccion);
        l.alto = altoLinea;
        l.arriba = y;
        l.base = y + altoLinea - bajoLaBase;
        y += altoLinea;
        l.parrafo = p; l.ultimaDelParrafo = il === lineasP.length - 1;
      });
      lineas.push(...lineasP);
      const despues = ip < parrafos.length - 1 || cuerpo.primeroYUltimo ? espacio(p.despues) : 0;
      y += despues;
      // PowerPoint cuenta estos espacios dentro de la caja de la primera y la
      // última línea del párrafo (TextRange2.Lines): para comparar igual.
      lineasP[0].antes = antes; lineasP[lineasP.length - 1].despues = despues;
    });

    // Alineación horizontal de cada línea.
    for (const l of lineas) {
      const alg = l.parrafo.alineacion;
      const usado = anchoSinEspacioFinal(l);
      const libre = l.limite === Infinity ? 0 : l.limite - usado;
      let x = l.inicio;
      let extraPorEspacio = 0;
      if (l.limite !== Infinity) {
        if (alg === 'ctr') x += libre / 2;
        else if (alg === 'r') x += libre;
        else if ((alg === 'just' && !l.final) || alg === 'dist') {
          const espacios = l.piezas.slice(0, ultimaNoEspacio(l) + 1).filter((pz) => pz.espacio).length;
          if (espacios) extraPorEspacio = libre / espacios;
        }
      } else if (alg === 'ctr' || alg === 'r') {
        // Sin ajuste: el texto se alinea sobre el ancho de la caja.
        const libreCaja = anchoCol - l.inicio - usado;
        x += alg === 'ctr' ? libreCaja / 2 : libreCaja;
      }
      l.trozos = [];
      for (const pz of l.piezas.slice(0, ultimaNoEspacio(l) + 1)) {
        if (!pz.espacio && pz.texto) {
          const t = pz.t;
          l.trozos.push({
            texto: pz.texto, x, ancho: pz.ancho, css: css(t, pz.px), px: pz.px,
            espaciado: t.espaciado ? t.espaciado * ptPx : 0, kerning: kerning(t, pz.px), color: colorCss(colorDeRelleno(t.relleno)),
            subrayado: t.subrayado, tachado: t.tachado, resaltado: colorCss(t.resaltado),
            desplazamiento: -t.base * escalar(t.tamano) * ptPx, sombra: t.sombra,
          });
        }
        x += pz.ancho + (pz.espacio ? extraPorEspacio : 0);
      }
    }

    // Columnas: las líneas pasan a la siguiente columna cuando no entran.
    if (columnas > 1) {
      let col = 0, desplazamiento = 0;
      for (const l of lineas) {
        if (l.base - desplazamiento > alto && col < columnas - 1 && l !== lineas[0]) { col++; desplazamiento = l.base - l.alto * 0.8; }
        l.base -= desplazamiento; l.arriba -= desplazamiento;
        const dx = col * (anchoCol + espacioCol);
        for (const t of l.trozos) t.x += dx;
        if (l.vineta) l.vineta.x += dx;
      }
      y = Math.min(y, alto);
    }
    return { lineas, alto: y };
  }
  function ultimaNoEspacio(l) { for (let i = l.piezas.length - 1; i >= 0; i--) if (!l.piezas[i].espacio) return i; return -1; }
  function anchoSinEspacioFinal(l) {
    let w = l.ancho;
    for (let i = l.piezas.length - 1; i >= 0 && l.piezas[i].espacio; i--) w -= l.piezas[i].ancho;
    return w;
  }

  return { maquetar, familia, colorCss, cargarFuentes };
})();

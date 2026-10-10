/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Dibuja en un canvas una diapositiva ya resuelta por pptx-modelo.js: fondo,
// formas (geometría de pptx-geometria.js), rellenos, líneas con guiones y
// puntas de flecha, imágenes (recorte, mosaico, efectos), grupos con su
// transformación, sombras y texto (maquetado por pptx-texto.js).
//
// Etapa 2 del motor propio. Lo que todavía es aproximado está marcado
// "[calibrar]" y se ajusta midiendo con npm run comparar-pptx.
//
// Uso: await PptxDibujo.dibujar(pres, diapositiva, canvas)  (canvas ya del
// tamaño final; la escala sale de canvas.width / pres.ancho). Devuelve
// { textos }: por cada forma con texto, su nombre y sus líneas (texto y caja
// en puntos de la diapositiva), para comparar los cortes de línea con
// PowerPoint (comparar-pptx.mjs).

const PptxDibujo = (() => {
  const EMU_POR_PX96 = 9525;
  const rgba = (c, alfa = 1) => `rgba(${c.r},${c.g},${c.b},${Math.max(0, Math.min(1, c.a * alfa))})`;

  async function dibujar(pres, diap, canvas) {
    const ctx = canvas.getContext('2d');
    const s = canvas.width / pres.ancho;
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.globalAlpha = 1;
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.imageSmoothingQuality = 'high';

    // Todas las imágenes se cargan antes de dibujar (el dibujo es síncrono).
    // Y las fuentes, antes de medir el texto.
    const rutas = new Set();
    const usos = [];
    const juntar = (r) => { if (r && r.tipo === 'imagen') rutas.add(r.ruta); };
    juntar(diap.fondo);
    (function recorrer(lista) {
      for (const el of lista) {
        juntar(el.relleno);
        if (el.imagen) juntar(el.imagen);
        if (el.texto) for (const p of el.texto.parrafos) usos.push(...p.trozos, p.fin);
        if (el.hijos) recorrer(el.hijos);
      }
    })(diap.elementos);
    const imagenes = new Map();
    await Promise.all([...[...rutas].map(async (r) => imagenes.set(r, await pres.imagen(r))), PptxTexto.cargarFuentes(usos.filter(Boolean))]);

    const entorno = { s, imagenes, fondo: diap.fondo, base: ctx.getTransform(), anchoDiap: pres.ancho * s, altoDiap: pres.alto * s, textos: [] };
    const todo = new Path2D(); todo.rect(0, 0, entorno.anchoDiap, entorno.altoDiap);
    if (diap.fondo) pintar(ctx, diap.fondo, todo, entorno.anchoDiap, entorno.altoDiap, entorno);

    const mapa = (x) => ({ x: x.x * s, y: x.y * s, w: x.w * s, h: x.h * s });
    for (const el of diap.elementos) dibujarElemento(ctx, el, mapa, entorno, null);
    return { textos: entorno.textos };
  }

  function aplicarTransformacion(ctx, caja, x) {
    ctx.translate(caja.x + caja.w / 2, caja.y + caja.h / 2);
    if (x.rot) ctx.rotate(x.rot * Math.PI / 180);
    if (x.flipH || x.flipV) ctx.scale(x.flipH ? -1 : 1, x.flipV ? -1 : 1);
    ctx.translate(-caja.w / 2, -caja.h / 2);
  }

  function dibujarElemento(ctx, el, mapa, entorno, rellenoGrupo) {
    if (!el.xfrm || el.tipo === 'tabla' || el.tipo === 'grafico') return;
    const caja = mapa(el.xfrm);
    ctx.save();
    aplicarTransformacion(ctx, caja, el.xfrm);
    try {
      if (el.tipo === 'grupo') {
        const h = el.xfrm.hijos || { x: el.xfrm.x, y: el.xfrm.y, w: el.xfrm.w, h: el.xfrm.h };
        const kx = h.w ? caja.w / h.w : 1, ky = h.h ? caja.h / h.h : 1;
        const mapaHijos = (x) => ({ x: (x.x - h.x) * kx, y: (x.y - h.y) * ky, w: x.w * kx, h: x.h * ky });
        const rg = el.relleno && el.relleno.tipo !== 'grupo' ? el.relleno : rellenoGrupo;
        for (const hijo of el.hijos) dibujarElemento(ctx, hijo, mapaHijos, entorno, rg);
      } else {
        dibujarForma(ctx, el, caja, entorno, rellenoGrupo);
      }
    } catch (e) {
      console.warn('No se pudo dibujar', el.nombre, e);
    }
    ctx.restore();
  }

  // --- Trazados ------------------------------------------------------------------------
  function aPath2D(cmds) {
    const p = new Path2D();
    for (const c of cmds) {
      if (c[0] === 'M') p.moveTo(c[1], c[2]);
      else if (c[0] === 'L') p.lineTo(c[1], c[2]);
      else if (c[0] === 'C') p.bezierCurveTo(c[1], c[2], c[3], c[4], c[5], c[6]);
      else if (c[0] === 'Z') p.closePath();
    }
    return p;
  }

  function conSombra(ctx, sombra, entorno, dibujo) {
    if (!sombra) { dibujo(); return; }
    ctx.save();
    const ang = sombra.direccion * Math.PI / 180;
    ctx.shadowColor = rgba(sombra.color);
    ctx.shadowBlur = sombra.desenfoque * entorno.s; // [calibrar]
    ctx.shadowOffsetX = Math.cos(ang) * sombra.distancia * entorno.s;
    ctx.shadowOffsetY = Math.sin(ang) * sombra.distancia * entorno.s;
    dibujo();
    ctx.restore();
  }

  function dibujarForma(ctx, el, caja, entorno, rellenoGrupo) {
    const w = caja.w, h = caja.h;
    const geo = el.geometria && el.geometria.def ? PptxGeometria.trazar(el.geometria.def, w, h, el.geometria.ajustes) : null;
    const rutas = geo ? geo.rutas.map((r) => ({ ...r, path: aPath2D(r.cmds) })) : [];
    let rell = el.relleno;
    if (rell && rell.tipo === 'grupo') rell = rellenoGrupo;
    const linea = el.linea;
    let sombraPendiente = el.sombra;

    if (el.tipo === 'imagen') {
      const recorte = new Path2D();
      if (rutas.length) for (const r of rutas) recorte.addPath(r.path); else recorte.rect(0, 0, w, h);
      if (rell) pintar(ctx, rell, recorte, w, h, entorno);
      conSombra(ctx, sombraPendiente, entorno, () => pintar(ctx, el.imagen, recorte, w, h, entorno));
      sombraPendiente = null;
    } else if (rell) {
      for (const r of rutas) {
        if (r.relleno === 'none') continue;
        conSombra(ctx, sombraPendiente, entorno, () => pintar(ctx, rell, r.path, w, h, entorno, r.relleno));
        sombraPendiente = null;
      }
    }
    if (linea) {
      for (const r of rutas) {
        if (!r.trazo) continue;
        conSombra(ctx, sombraPendiente, entorno, () => trazarLinea(ctx, linea, r, entorno));
        sombraPendiente = null;
      }
      puntas(ctx, linea, rutas.filter((r) => r.trazo), entorno);
    }
    if (el.texto) dibujarTexto(ctx, el, w, h, geo ? geo.texto : { l: 0, t: 0, r: w, b: h }, caja, entorno);
  }

  // --- Rellenos ----------------------------------------------------------------------------
  // modo: el "fill" del trazado (lighten, darken…): el mismo relleno más
  // claro u oscuro, como el lado de un cubo o el pliegue de una cinta.
  const MODOS = { lighten: ['#fff', 0.4], lightenLess: ['#fff', 0.2], darken: ['#000', 0.4], darkenLess: ['#000', 0.2] };

  function pintar(ctx, rell, path, w, h, entorno, modo = 'norm') {
    ctx.save();
    try {
      switch (rell.tipo) {
        case 'solido': ctx.fillStyle = rgba(rell.color); ctx.fill(path, 'evenodd'); break;
        case 'degradado': degradado(ctx, rell, path, w, h); break;
        case 'trama': trama(ctx, rell, path, entorno); break;
        case 'imagen': imagenEn(ctx, rell, path, w, h, entorno); break;
        case 'fondo': {
          // El fondo de la diapositiva, en sus coordenadas, solo dentro de la forma.
          ctx.clip(path, 'evenodd');
          ctx.setTransform(entorno.base);
          const todo = new Path2D(); todo.rect(0, 0, entorno.anchoDiap, entorno.altoDiap);
          if (entorno.fondo && entorno.fondo.tipo !== 'fondo') pintar(ctx, entorno.fondo, todo, entorno.anchoDiap, entorno.altoDiap, entorno);
          break;
        }
      }
      const m = MODOS[modo];
      if (m) { ctx.globalAlpha = m[1]; ctx.fillStyle = m[0]; ctx.fill(path, 'evenodd'); }
    } finally { ctx.restore(); }
  }

  function degradado(ctx, d, path, w, h) {
    if (d.ruta) {
      const f = d.foco;
      const fx = w * (f.l + (1 - f.r)) / 2, fy = h * (f.t + (1 - f.b)) / 2;
      if (d.ruta === 'circle') {
        const radio = Math.max(Math.hypot(fx, fy), Math.hypot(w - fx, fy), Math.hypot(fx, h - fy), Math.hypot(w - fx, h - fy));
        const g = ctx.createRadialGradient(fx, fy, 0, fx, fy, radio || 1);
        for (const p of paradasLineales(d.paradas)) g.addColorStop(p.pos, rgba(p.color));
        ctx.fillStyle = g; ctx.fill(path, 'evenodd');
        return;
      }
      // rect / shape: el color depende de la distancia al rectángulo del foco
      // en la dirección dominante. Se calcula chico y se estira (es suave).
      const lado = 128;
      const c = document.createElement('canvas'); c.width = lado; c.height = lado;
      const cx = c.getContext('2d'), img = cx.createImageData(lado, lado);
      const fl = w * f.l, ft = h * f.t, fr = w * (1 - f.r), fb = h * (1 - f.b);
      const tabla = rampa(d.paradas);
      for (let j = 0; j < lado; j++) for (let i = 0; i < lado; i++) {
        const x = (i + 0.5) * w / lado, y = (j + 0.5) * h / lado;
        const tx = x < fl ? (fl - x) / (fl || 1) : x > fr ? (x - fr) / ((w - fr) || 1) : 0;
        const ty = y < ft ? (ft - y) / (ft || 1) : y > fb ? (y - fb) / ((h - fb) || 1) : 0;
        const t = Math.min(1, Math.max(tx, ty));
        const k = Math.round(t * 255) * 4, o = (j * lado + i) * 4;
        img.data[o] = tabla[k]; img.data[o + 1] = tabla[k + 1]; img.data[o + 2] = tabla[k + 2]; img.data[o + 3] = tabla[k + 3];
      }
      cx.putImageData(img, 0, 0);
      ctx.clip(path, 'evenodd');
      ctx.drawImage(c, 0, 0, w, h);
      return;
    }
    // Lineal: el ángulo es el de las líneas de color, medido en el sentido de
    // las agujas del reloj. "scaled": el ángulo se deforma con la forma.
    const ang = d.angulo * Math.PI / 180;
    if (d.escalado) {
      ctx.clip(path, 'evenodd');
      ctx.scale(w || 1, h || 1);
      const L = Math.abs(Math.cos(ang)) + Math.abs(Math.sin(ang));
      const dx = Math.cos(ang) * L / 2, dy = Math.sin(ang) * L / 2;
      const g = ctx.createLinearGradient(0.5 - dx, 0.5 - dy, 0.5 + dx, 0.5 + dy);
      for (const p of paradasLineales(d.paradas)) g.addColorStop(p.pos, rgba(p.color));
      ctx.fillStyle = g; ctx.fillRect(0, 0, 1, 1);
      return;
    }
    const L = Math.abs(w * Math.cos(ang)) + Math.abs(h * Math.sin(ang));
    const dx = Math.cos(ang) * L / 2, dy = Math.sin(ang) * L / 2;
    const g = ctx.createLinearGradient(w / 2 - dx, h / 2 - dy, w / 2 + dx, h / 2 + dy);
    for (const p of paradasLineales(d.paradas)) g.addColorStop(p.pos, rgba(p.color));
    ctx.fillStyle = g; ctx.fill(path, 'evenodd');
  }
  // Cómo mezcla PowerPoint los colores de un degradado (medido el 2026-10-09
  // en comparar-pptx/corpus/05-rellenos, diapositivas 2 y 3):
  //  - con DOS paradas, en RGB lineal y con una curva suave (smoothstep):
  //    FFDC00→005AC8 a la mitad da 185,170,147 (así: 188,171,146; sRGB:
  //    128,155,100), y al 25 % 237,207,84 (así: 237,206,85);
  //  - con tres o más, en sRGB, como el canvas (4 paradas: igual punto a punto).
  // Para dos paradas, al canvas se le dan paradas intermedias ya mezcladas.
  const aLineal = (c) => { c /= 255; return c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4); };
  const aSrgb = (v) => Math.round(255 * (v <= 0.0031308 ? v * 12.92 : 1.055 * Math.pow(v, 1 / 2.4) - 0.055));
  const suave = (u) => u * u * (3 - 2 * u);
  function mezclar(a, b, u, dos) {
    if (!dos) {
      const m = (p, q) => Math.round(p + (q - p) * u);
      return { r: m(a.r, b.r), g: m(a.g, b.g), b: m(a.b, b.b), a: a.a + (b.a - a.a) * u };
    }
    const v = suave(u);
    const m = (p, q) => aSrgb(aLineal(p) + (aLineal(q) - aLineal(p)) * v);
    return { r: m(a.r, b.r), g: m(a.g, b.g), b: m(a.b, b.b), a: a.a + (b.a - a.a) * v };
  }
  function paradasLineales(paradas) {
    if (paradas.length !== 2) return paradas;
    const [a, b] = paradas, salida = [a];
    for (let k = 1; k <= 24; k++) salida.push({ pos: a.pos + (b.pos - a.pos) * k / 24, color: mezclar(a.color, b.color, k / 24, true) });
    return salida;
  }
  // 256 colores interpolados entre las paradas (RGBA plano).
  function rampa(paradas) {
    const t = new Uint8ClampedArray(256 * 4);
    for (let k = 0; k < 256; k++) {
      const x = k / 255;
      let a = paradas[0], b = paradas[paradas.length - 1];
      for (let i = 0; i < paradas.length - 1; i++) if (x >= paradas[i].pos && x <= paradas[i + 1].pos) { a = paradas[i]; b = paradas[i + 1]; break; }
      const u = b.pos > a.pos ? Math.min(1, Math.max(0, (x - a.pos) / (b.pos - a.pos))) : 0;
      const c = x <= paradas[0].pos ? paradas[0].color : x >= paradas[paradas.length - 1].pos ? paradas[paradas.length - 1].color : null;
      const col = c || mezclar(a.color, b.color, u, paradas.length === 2);
      t[k * 4] = col.r; t[k * 4 + 1] = col.g; t[k * 4 + 2] = col.b; t[k * 4 + 3] = Math.round(col.a * 255);
    }
    return t;
  }

  // Tramas (pattFill): dibujos de 8×8 puntos, una fila por byte (bit alto a
  // la izquierda; 1 = color de primer plano). Medidos de PowerPoint el
  // 2026-10-09 (comparar-pptx/corpus/05-rellenos, diapositiva 5, exportada a
  // 96 ppp): cada punto mide 1/96 de pulgada (escala con la diapositiva), la
  // trama arranca en el origen de la diapositiva y no gira con la forma. Las
  // seis primeras de la norma (horz…diagCross) no salen en el corpus: son las
  // rayas de un punto de siempre.
  const TRAMAS = {
    pct5: '8000000008000000', pct10: '8000080080000800', pct20: '8800220088002200', pct25: '8822882288228822',
    pct30: 'aa44aa11aa44aa11', pct40: 'aa55aa51aa55aa15', pct50: 'aa55aa55aa55aa55', pct60: 'ee55bb55ee55bb55',
    pct70: '77dd77dd77dd77dd', pct75: '77ffddff77ffddff', pct80: 'effffeffeffffeff', pct90: 'fffffff7ffffff7f',
    dkHorz: 'ffff0000ffff0000', dkVert: 'cccccccccccccccc', dkDnDiag: 'cc663399cc663399',
    dkUpDiag: '3366cc993366cc99', smCheck: '9966669999666699', trellis: 'ff66ff99ff66ff99',
    ltHorz: 'ff000000ff000000', ltVert: '8888888888888888', ltDnDiag: '8844221188442211',
    ltUpDiag: '1122448811224488', smGrid: 'ff888888ff888888', dotDmnd: '8000220008002200',
    wdDnDiag: 'c1e070381c0e0783', wdUpDiag: '83070e1c3870e0c1', dashUpDiag: '0000112244880000',
    dashDnDiag: '0000884422110000', narVert: '5555555555555555', narHorz: 'ff00ff00ff00ff00',
    dashVert: '8080808008080808', dashHorz: 'f00000000f000000', lgConfetti: 'b130031bd8c00c8d',
    lgGrid: 'ff80808080808080', horzBrick: 'ff808080ff080808', lgCheck: 'f0f0f0f00f0f0f0f',
    smConfetti: '8008400210012004', zigZag: '8142241881422418', solidDmnd: '10387cfe7c381000',
    diagBrick: '0102040818244281', openDmnd: '8244281028448201', plaid: 'aa55aa55f0f0f0f0',
    sphere: '77898f8f7798f8f8', weave: '8854224588142251', dotGrid: 'aa00800080008000',
    divot: '0010081000800180', shingle: '038448300c020101', wave: '001825c0001825c0', horz: 'ff00000000000000',
    vert: '8080808080808080', dnDiag: '8040201008040201', upDiag: '0102040810204080',
    cross: 'ff80808080808080', diagCross: '8142241818244281',
  };
  const lienzosTrama = new Map();
  function trama(ctx, t, path, entorno) {
    const clave = t.prst + rgba(t.fg) + rgba(t.bg);
    let c = lienzosTrama.get(clave);
    if (!c) {
      const bits = TRAMAS[t.prst] || TRAMAS.pct50;
      c = document.createElement('canvas'); c.width = 8; c.height = 8;
      const x = c.getContext('2d'), img = x.createImageData(8, 8);
      for (let j = 0; j < 8; j++) {
        const fila = parseInt(bits.slice(j * 2, j * 2 + 2), 16);
        for (let i = 0; i < 8; i++) {
          const col = (fila >> (7 - i)) & 1 ? t.fg : t.bg, o = (j * 8 + i) * 4;
          img.data[o] = col.r; img.data[o + 1] = col.g; img.data[o + 2] = col.b; img.data[o + 3] = Math.round(col.a * 255);
        }
      }
      x.putImageData(img, 0, 0);
      lienzosTrama.set(clave, c);
    }
    const patron = ctx.createPattern(c, 'repeat');
    // De las coordenadas de la forma a las de la diapositiva, y un punto de la
    // trama = 1/96 de pulgada.
    const k = EMU_POR_PX96 * entorno.s;
    patron.setTransform(ctx.getTransform().inverse().multiply(entorno.base).scale(k, k));
    ctx.imageSmoothingEnabled = false;
    ctx.fillStyle = patron; ctx.fill(path, 'evenodd');
  }

  // --- Imágenes -------------------------------------------------------------------------------
  const procesadas = new WeakMap();
  function conEfectos(img, efectos) {
    if (!efectos || !efectos.length) return img.fuente;
    let porImagen = procesadas.get(img);
    if (!porImagen) { porImagen = new Map(); procesadas.set(img, porImagen); }
    const clave = JSON.stringify(efectos);
    if (porImagen.has(clave)) return porImagen.get(clave);
    const c = document.createElement('canvas');
    c.width = img.ancho; c.height = img.alto;
    const x = c.getContext('2d', { willReadFrequently: true });
    x.drawImage(img.fuente, 0, 0, img.ancho, img.alto);
    const datos = x.getImageData(0, 0, c.width, c.height), d = datos.data;
    for (const e of efectos) {
      for (let k = 0; k < d.length; k += 4) {
        const y = 0.299 * d[k] + 0.587 * d[k + 1] + 0.114 * d[k + 2];
        switch (e.tipo) {
          case 'gris': d[k] = d[k + 1] = d[k + 2] = y; break;
          case 'bitono': d[k] = d[k + 1] = d[k + 2] = y >= e.umbral * 255 ? 255 : 0; break;
          case 'luz': for (let i = 0; i < 3; i++) { let v = d[k + i] / 255; v = (v - 0.5) * (1 + e.contraste) / (1 - Math.min(0.99, Math.max(0, e.contraste)) || 1) + 0.5 + e.brillo; d[k + i] = Math.max(0, Math.min(255, v * 255)); } break;
          case 'duotono': { const u = y / 255, [a, b] = e.colores; d[k] = a.r + (b.r - a.r) * u; d[k + 1] = a.g + (b.g - a.g) * u; d[k + 2] = a.b + (b.b - a.b) * u; break; }
          case 'cambioColor': if (Math.abs(d[k] - e.de.r) < 3 && Math.abs(d[k + 1] - e.de.g) < 3 && Math.abs(d[k + 2] - e.de.b) < 3) { d[k] = e.a.r; d[k + 1] = e.a.g; d[k + 2] = e.a.b; d[k + 3] = Math.round(e.a.a * 255); } break;
        }
      }
    }
    x.putImageData(datos, 0, 0);
    porImagen.set(clave, c);
    return c;
  }

  function imagenEn(ctx, r, path, w, h, entorno) {
    const img = entorno.imagenes.get(r.ruta);
    if (!img) return;
    const fuente = conEfectos(img, r.efectos);
    ctx.clip(path, 'evenodd');
    ctx.globalAlpha *= r.alfa;
    if (r.mosaico) {
      const m = r.mosaico, s = entorno.s;
      const tw = img.ancho / img.dpiX * 914400 * m.sx * s, th = img.alto / img.dpiY * 914400 * m.sy * s;
      const ax = /l$/.test(m.algn) || m.algn === 'l' ? 0 : /r$/.test(m.algn) || m.algn === 'r' ? w : w / 2;
      const ay = /^t/.test(m.algn) ? 0 : /^b/.test(m.algn) ? h : h / 2;
      const ox = ax - (ax === 0 ? 0 : ax === w ? tw : tw / 2) + m.tx * s, oy = ay - (ay === 0 ? 0 : ay === h ? th : th / 2) + m.ty * s;
      const patron = ctx.createPattern(fuente, 'repeat');
      patron.setTransform(new DOMMatrix([tw / img.ancho, 0, 0, th / img.alto, ox, oy]));
      ctx.fillStyle = patron;
      ctx.fillRect(0, 0, w, h);
      return;
    }
    const e = r.estirar || { l: 0, t: 0, r: 0, b: 0 };
    const dx = w * e.l, dy = h * e.t, dw = w * (1 - e.l - e.r), dh = h * (1 - e.t - e.b);
    const c = r.recorte || { l: 0, t: 0, r: 0, b: 0 };
    const sw = img.ancho * (1 - c.l - c.r), sh = img.alto * (1 - c.t - c.b);
    if (sw <= 0 || sh <= 0 || dw <= 0 || dh <= 0) return;
    // Con recorte "negativo" (relleno alrededor), la parte fuera de la imagen
    // queda vacía: se dibuja solo la parte que existe, en su lugar.
    const sx = img.ancho * c.l, sy = img.alto * c.t;
    const kx = dw / sw, ky = dh / sh;
    const x0 = Math.max(0, sx), y0 = Math.max(0, sy), x1 = Math.min(img.ancho, sx + sw), y1 = Math.min(img.alto, sy + sh);
    if (x1 <= x0 || y1 <= y0) return;
    ctx.drawImage(fuente, x0, y0, x1 - x0, y1 - y0, dx + (x0 - sx) * kx, dy + (y0 - sy) * ky, (x1 - x0) * kx, (y1 - y0) * ky);
  }

  // --- Líneas ------------------------------------------------------------------------------------
  const GUIONES = {
    solid: [], dot: [1, 3], dash: [4, 3], lgDash: [8, 3], dashDot: [4, 3, 1, 3], lgDashDot: [8, 3, 1, 3],
    lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDash: [3, 1], sysDot: [1, 1], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1],
  };
  const anchoLinea = (l, s) => Math.max(l.ancho * s, 1);
  function colorLinea(l) {
    const r = l.relleno;
    return r.tipo === 'solido' ? r.color : r.tipo === 'degradado' ? r.paradas[0].color : r.tipo === 'trama' ? r.fg : { r: 0, g: 0, b: 0, a: 1 };
  }
  function trazarLinea(ctx, l, ruta, entorno) {
    const lw = anchoLinea(l, entorno.s);
    ctx.save();
    ctx.strokeStyle = rgba(colorLinea(l));
    ctx.lineWidth = lw;
    ctx.lineCap = l.tapa === 'rnd' ? 'round' : l.tapa === 'sq' ? 'square' : 'butt';
    ctx.lineJoin = l.union === 'bevel' ? 'bevel' : l.union === 'miter' ? 'miter' : 'round';
    const g = Array.isArray(l.guion) ? l.guion : GUIONES[l.guion] || [];
    if (g.length) ctx.setLineDash(g.map((v) => v * lw));
    // Las líneas compuestas (doble, triple) salen simples por ahora: borrar el
    // hueco del medio borraría también lo que hay debajo. [etapa 5]
    ctx.stroke(ruta.path);
    ctx.restore();
  }

  // Puntas de flecha: en el comienzo (headEnd) del primer trazado abierto y en
  // el final (tailEnd) del último.
  const TAM_PUNTA = { sm: 2, med: 3, lg: 5 };
  function puntas(ctx, l, rutas, entorno) {
    if (!l.inicio && !l.fin) return;
    const abiertas = rutas.filter((r) => r.cmds.length > 1 && r.cmds[r.cmds.length - 1][0] !== 'Z');
    if (!abiertas.length) return;
    const lw = anchoLinea(l, entorno.s);
    const color = rgba(colorLinea(l));
    const puntos = (cmds) => {
      const pts = [];
      for (const c of cmds) if (c[0] === 'M' || c[0] === 'L') pts.push([c[1], c[2]]); else if (c[0] === 'C') pts.push([c[1], c[2]], [c[3], c[4]], [c[5], c[6]]);
      return pts;
    };
    const dibujarPunta = (p, tipo, desde) => {
      // desde: un punto anterior para la dirección (distinto de p).
      const ang = Math.atan2(p[1] - desde[1], p[0] - desde[0]);
      const largo = TAM_PUNTA[tipo.len] * lw, ancho = TAM_PUNTA[tipo.w] * lw;
      ctx.save();
      ctx.translate(p[0], p[1]); ctx.rotate(ang);
      ctx.fillStyle = color; ctx.strokeStyle = color; ctx.lineWidth = lw; ctx.setLineDash([]);
      ctx.beginPath();
      switch (tipo.tipo) {
        case 'triangle': ctx.moveTo(0, 0); ctx.lineTo(-largo, -ancho / 2); ctx.lineTo(-largo, ancho / 2); ctx.closePath(); ctx.fill(); break;
        case 'stealth': ctx.moveTo(0, 0); ctx.lineTo(-largo, -ancho / 2); ctx.lineTo(-largo * 0.6, 0); ctx.lineTo(-largo, ancho / 2); ctx.closePath(); ctx.fill(); break;
        case 'diamond': ctx.moveTo(largo / 2, 0); ctx.lineTo(0, -ancho / 2); ctx.lineTo(-largo / 2, 0); ctx.lineTo(0, ancho / 2); ctx.closePath(); ctx.fill(); break;
        case 'oval': ctx.ellipse(0, 0, largo / 2, ancho / 2, 0, 0, Math.PI * 2); ctx.fill(); break;
        case 'arrow': ctx.lineJoin = 'miter'; ctx.moveTo(-largo, -ancho / 2); ctx.lineTo(0, 0); ctx.lineTo(-largo, ancho / 2); ctx.stroke(); break;
      }
      ctx.restore();
    };
    const distinto = (pts, p, desdeFinal) => {
      const orden = desdeFinal ? [...pts].reverse() : pts;
      return orden.find((q) => Math.hypot(q[0] - p[0], q[1] - p[1]) > 0.01) || [p[0] - 1, p[1]];
    };
    if (l.inicio) { const pts = puntos(abiertas[0].cmds); const p = pts[0]; dibujarPunta(p, l.inicio, distinto(pts.slice(1), p, false)); }
    if (l.fin) { const pts = puntos(abiertas[abiertas.length - 1].cmds); const p = pts[pts.length - 1]; dibujarPunta(p, l.fin, distinto(pts.slice(0, -1), p, true)); }
  }

  // --- Texto ----------------------------------------------------------------------------------------
  function dibujarTexto(ctx, el, w, h, rect, caja, entorno) {
    const s = entorno.s;
    const cuerpo = el.texto.cuerpo;
    let r = { x: rect.l, y: rect.t, w: rect.r - rect.l, h: rect.b - rect.t };
    if (el.cajaTexto) {
      // El texto de un SmartArt va en su propio rectángulo (txXfrm).
      const kx = el.xfrm.w ? w / el.xfrm.w : 1, ky = el.xfrm.h ? h / el.xfrm.h : 1;
      r = { x: (el.cajaTexto.x - el.xfrm.x) * kx, y: (el.cajaTexto.y - el.xfrm.y) * ky, w: el.cajaTexto.w * kx, h: el.cajaTexto.h * ky };
    }
    ctx.save();
    // El texto no se espeja con la forma; si está volteada en vertical, queda
    // girado 180°.
    if (el.xfrm.flipH || el.xfrm.flipV) {
      ctx.translate(w / 2, h / 2); ctx.scale(el.xfrm.flipH ? -1 : 1, el.xfrm.flipV ? -1 : 1); ctx.translate(-w / 2, -h / 2);
      if (el.xfrm.flipV) { ctx.translate(w / 2, h / 2); ctx.rotate(Math.PI); ctx.translate(-w / 2, -h / 2); }
    }
    let giro = cuerpo.giro || 0;
    if (cuerpo.derecho) giro -= el.xfrm.rot || 0;
    if (el.cajaTexto && el.cajaTexto.rot) giro += el.cajaTexto.rot - (el.xfrm.rot || 0);
    const vertical = cuerpo.vertical === 'vert' || cuerpo.vertical === 'eaVert' || cuerpo.vertical === 'wordArtVert' ? 90 : cuerpo.vertical === 'vert270' ? -90 : 0;
    const centro = [r.x + r.w / 2, r.y + r.h / 2];
    if (giro || vertical) {
      ctx.translate(centro[0], centro[1]); ctx.rotate((giro + vertical) * Math.PI / 180); ctx.translate(-centro[0], -centro[1]);
      if (vertical) r = { x: centro[0] - r.h / 2, y: centro[1] - r.w / 2, w: r.h, h: r.w };
    }
    const m = cuerpo.margenes;
    const area = { x: r.x + m.l * s, y: r.y + m.t * s, w: Math.max(0, r.w - (m.l + m.r) * s), h: Math.max(0, r.h - (m.t + m.b) * s) };
    const maquetado = PptxTexto.maquetar(el.texto, area.w, area.h, s, ctx);
    let y0 = area.y;
    if (cuerpo.ancla === 'ctr') y0 += (area.h - maquetado.alto) / 2;
    else if (cuerpo.ancla === 'b') y0 += area.h - maquetado.alto;
    ctx.textBaseline = 'alphabetic';
    ctx.textAlign = 'left';
    registrarLineas(ctx, el, maquetado, area, y0, entorno);
    for (const l of maquetado.lineas) {
      const base = y0 + l.base;
      if (l.vineta && l.vineta.color) {
        ctx.font = l.vineta.css; ctx.letterSpacing = '0px'; ctx.fillStyle = l.vineta.color;
        ctx.fillText(l.vineta.texto, area.x + l.vineta.x, base);
      }
      for (const t of l.trozos) {
        const x = area.x + t.x, y = base + t.desplazamiento;
        if (t.resaltado) { ctx.fillStyle = t.resaltado; ctx.fillRect(x, y - t.px * 0.95, t.ancho, t.px * 1.2); }
        if (!t.color) continue;
        ctx.font = t.css; ctx.letterSpacing = t.espaciado ? `${t.espaciado}px` : '0px'; ctx.fillStyle = t.color;
        conSombra(ctx, t.sombra, entorno, () => ctx.fillText(t.texto, x, y));
        if (t.subrayado || t.tachado) {
          ctx.fillStyle = t.color;
          const grosor = Math.max(1, t.px * 0.06);
          if (t.subrayado) {
            const dobles = /dbl/.test(t.subrayado);
            ctx.fillRect(x, y + t.px * 0.12, t.ancho, grosor);
            if (dobles) ctx.fillRect(x, y + t.px * 0.12 + grosor * 2, t.ancho, grosor);
          }
          if (t.tachado) {
            ctx.fillRect(x, y - t.px * 0.3, t.ancho, grosor);
            if (t.tachado === 'dblStrike') ctx.fillRect(x, y - t.px * 0.3 - grosor * 2, t.ancho, grosor);
          }
        }
      }
    }
    ctx.restore();
  }

  // Las líneas tal como quedaron cortadas, en puntos de la diapositiva (la
  // caja es la del texto sin girar: vale para formas sin giro).
  function registrarLineas(ctx, el, maquetado, area, y0, entorno) {
    const m = ctx.getTransform();
    const aPt = 1 / (entorno.s * 12700);
    const punto = (x, y) => ({ x: (m.a * x + m.c * y + m.e) * aPt, y: (m.b * x + m.d * y + m.f) * aPt });
    const lineas = [];
    for (const l of maquetado.lineas) {
      const texto = l.piezas.map((pz) => pz.tab ? '\t' : pz.texto || '').join('').trim();
      if (!texto) continue;
      const a = l.trozos[0], z = l.trozos[l.trozos.length - 1];
      const p0 = punto(area.x + (a ? a.x : 0), y0 + l.arriba - (l.antes || 0));
      const p1 = punto(area.x + (z ? z.x + z.ancho : 0), y0 + l.arriba + l.alto + (l.despues || 0));
      lineas.push({ texto, x: p0.x, y: p0.y, w: p1.x - p0.x, h: p1.y - p0.y });
    }
    if (lineas.length) entorno.textos.push({ nombre: el.nombre, lineas });
  }

  return { dibujar };
})();

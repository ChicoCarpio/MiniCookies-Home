/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Máscaras de capa del Editor avanzado: dónde se aplica cada capa.
//
// Una capa tiene UNA máscara, como en Photoshop, armada en tres pasos:
//   1. Una base: toda la imagen, nada, un degradado (lineal o radial), un
//      color o un rango de luz (sombras, medios, luces).
//   2. El pincel encima: lo pintado en "sumar" se agrega a la base y lo
//      pintado en "quitar" se le resta. Son dos lienzos (alfa = cuánto), así
//      el pincel combina con cualquier base en vez de reemplazarla.
//   3. Invertir: da vuelta el resultado.
// El resultado es un lienzo cuyo alfa dice cuánto se aplica la capa en cada
// punto (blanco opaco = entera, transparente = nada), o null si es "toda la
// imagen" sin retoques (el caso común, sin costo).
//
// Color y luz se calculan píxel a píxel sobre la imagen original: van a una
// caché ("memo", una por capa) y solo se rehacen si cambia su configuración o
// la imagen. Con memo, además, se reusan los lienzos de un render al
// siguiente (crear lienzos por fotograma es lo que traba la vista previa);
// sin memo (la descarga a resolución completa) todo sale en lienzos nuevos.

window.MascarasCapa = (function () {
  'use strict';

  const BASES = ['toda', 'nada', 'degradado', 'color', 'luz'];
  // Rangos rápidos de "Luz", en 0..100 de luminosidad.
  const RANGOS_LUZ = { sombras: [0, 35], medios: [35, 65], luces: [65, 100] };

  function neutra() {
    return {
      base: 'toda',
      invertir: false,
      degradado: { tipo: 'lineal', a: null, b: null }, // a/b en 0..1 del ancho/alto
      color: null, // {r,g,b}
      tolerancia: 30,
      luz: { desde: RANGOS_LUZ.sombras[0], hasta: RANGOS_LUZ.sombras[1], suavidad: 15 },
    };
  }

  // Completa una máscara guardada con lo que le falte (proyectos viejos).
  function normalizar(m) {
    const n = neutra();
    if (!m) return n;
    return {
      base: BASES.includes(m.base) ? m.base : n.base,
      invertir: !!m.invertir,
      degradado: { ...n.degradado, ...(m.degradado || {}) },
      color: m.color || null,
      tolerancia: Number.isFinite(m.tolerancia) ? m.tolerancia : n.tolerancia,
      luz: { ...n.luz, ...(m.luz || {}) },
    };
  }

  function rgbAHsl(r, g, b) {
    r /= 255; g /= 255; b /= 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, s = 0; const l = (max + min) / 2;
    const d = max - min;
    if (d !== 0) {
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      if (max === r) h = ((g - b) / d + (g < b ? 6 : 0));
      else if (max === g) h = (b - r) / d + 2;
      else h = (r - g) / d + 4;
      h *= 60;
    }
    return { h, s, l };
  }

  function lienzo(ancho, alto, reusar) {
    const c = reusar || document.createElement('canvas');
    if (c.width !== ancho || c.height !== alto) { c.width = ancho; c.height = alto; }
    else c.getContext('2d').clearRect(0, 0, ancho, alto);
    return c;
  }

  // Recorre la imagen y escribe en el alfa lo que devuelve "alfa(r,g,b)".
  function porPixel(origen, ancho, alto, destino, alfa) {
    const ctx = destino.getContext('2d');
    ctx.drawImage(origen, 0, 0, ancho, alto);
    const datos = ctx.getImageData(0, 0, ancho, alto);
    const d = datos.data;
    for (let i = 0; i < d.length; i += 4) {
      const a = alfa(d[i], d[i + 1], d[i + 2]);
      d[i] = 255; d[i + 1] = 255; d[i + 2] = 255;
      d[i + 3] = Math.round(a * 255);
    }
    ctx.putImageData(datos, 0, 0);
    return destino;
  }

  // Por color: parecido en tono, y no demasiado lejos en saturación y luz.
  function dibujarColor(origen, color, tolerancia, ancho, alto, destino) {
    const objetivo = rgbAHsl(color.r, color.g, color.b);
    const radioTono = 8 + (tolerancia / 100) * 90;
    const radioSL = 0.15 + (tolerancia / 100) * 0.6;
    const suavizado = Math.max(4, radioTono * 0.25);
    return porPixel(origen, ancho, alto, destino, (r, g, b) => {
      const hsl = rgbAHsl(r, g, b);
      if (Math.abs(hsl.s - objetivo.s) > radioSL || Math.abs(hsl.l - objetivo.l) > radioSL) return 0;
      let difTono = Math.abs(hsl.h - objetivo.h);
      if (difTono > 180) difTono = 360 - difTono;
      if (difTono <= radioTono - suavizado) return 1;
      if (difTono >= radioTono) return 0;
      return 1 - (difTono - (radioTono - suavizado)) / suavizado;
    });
  }

  // Por luz: entera dentro de [desde, hasta] y se desvanece afuera a lo largo
  // de "suavidad" puntos (luminosidad percibida, 0..100).
  function dibujarLuz(origen, luz, ancho, alto, destino) {
    const desde = Math.min(luz.desde, luz.hasta), hasta = Math.max(luz.desde, luz.hasta);
    const suavidad = Math.max(0, luz.suavidad);
    return porPixel(origen, ancho, alto, destino, (r, g, b) => {
      const l = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 2.55;
      const fuera = l < desde ? desde - l : (l > hasta ? l - hasta : 0);
      if (!fuera) return 1;
      return suavidad ? Math.max(0, 1 - fuera / suavidad) : 0;
    });
  }

  // Degradado: entero en el punto donde empezó el arrastre (a) y nada en el
  // que terminó (b). Radial: a es el centro y b marca el radio. Sin trazar
  // todavía, se aplica entero (para que elegir "Degradado" no borre la capa).
  function dibujarDegradado(ctx, degradado, ancho, alto) {
    const { a, b, tipo } = degradado;
    if (!a || !b) { ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, ancho, alto); return; }
    const ax = a.x * ancho, ay = a.y * alto, bx = b.x * ancho, by = b.y * alto;
    const gradiente = tipo === 'radial'
      ? ctx.createRadialGradient(ax, ay, 0, ax, ay, Math.max(1, Math.hypot(bx - ax, by - ay)))
      : ctx.createLinearGradient(ax, ay, bx, by);
    gradiente.addColorStop(0, 'rgba(255,255,255,1)');
    gradiente.addColorStop(1, 'rgba(255,255,255,0)');
    ctx.fillStyle = gradiente;
    ctx.fillRect(0, 0, ancho, alto);
  }

  // Base de color o de luz, cacheada en memo mientras no cambie nada de lo
  // que la define.
  function basePorPixel(m, origen, ancho, alto, memo) {
    const firma = m.base === 'color' ? JSON.stringify([m.color, m.tolerancia]) : JSON.stringify(m.luz);
    const clave = m.base + firma;
    if (memo && memo.base && memo.baseClave === clave && memo.baseOrigen === origen
      && memo.base.width === ancho && memo.base.height === alto) return memo.base;
    const destino = lienzo(ancho, alto, memo ? memo.base : null);
    if (m.base === 'color') dibujarColor(origen, m.color, m.tolerancia, ancho, alto, destino);
    else dibujarLuz(origen, m.luz, ancho, alto, destino);
    if (memo) { memo.base = destino; memo.baseClave = clave; memo.baseOrigen = origen; }
    return destino;
  }

  // ¿La máscara deja la capa en toda la imagen sin retoques? Entonces no hace
  // falta ningún lienzo.
  function esTodaLaImagen(m, quitar) {
    return m.base === 'toda' && !m.invertir && !quitar;
  }

  // m: la máscara de la capa. origen: la imagen original (para color y luz).
  // sumar/quitar: lienzos del pincel, o null si no se pintó nada.
  function crear(m, origen, sumar, quitar, ancho, alto, memo = null) {
    m = normalizar(m);
    if (esTodaLaImagen(m, quitar)) return null;
    // Color sin color elegido todavía: como "toda" (igual que antes).
    const base = m.base === 'color' && !m.color ? 'toda' : m.base;

    const final = lienzo(ancho, alto, memo ? memo.final : null);
    if (memo) memo.final = final;
    const ctx = final.getContext('2d');
    if (base === 'toda') { ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, ancho, alto); }
    else if (base === 'degradado') dibujarDegradado(ctx, m.degradado, ancho, alto);
    else if (base === 'color' || base === 'luz') ctx.drawImage(basePorPixel({ ...m, base }, origen, ancho, alto, memo), 0, 0);
    // 'nada': queda vacía; solo cuenta lo pintado.

    if (sumar && base !== 'toda') ctx.drawImage(sumar, 0, 0, ancho, alto);
    if (quitar) {
      ctx.globalCompositeOperation = 'destination-out';
      ctx.drawImage(quitar, 0, 0, ancho, alto);
      ctx.globalCompositeOperation = 'source-over';
    }
    if (!m.invertir) return final;

    const invertida = lienzo(ancho, alto, memo ? memo.invertida : null);
    if (memo) memo.invertida = invertida;
    const ctxInv = invertida.getContext('2d');
    ctxInv.fillStyle = '#fff';
    ctxInv.fillRect(0, 0, ancho, alto);
    ctxInv.globalCompositeOperation = 'destination-out';
    ctxInv.drawImage(final, 0, 0);
    ctxInv.globalCompositeOperation = 'source-over';
    return invertida;
  }

  return { BASES, RANGOS_LUZ, neutra, normalizar, crear, esTodaLaImagen };
})();

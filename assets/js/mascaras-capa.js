/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Máscaras de capa del Editor avanzado: dónde se aplica cada capa.
//
// Desde el 2026-10-07 (proyecto versión 8) una máscara es una LISTA ORDENADA
// DE PIEZAS, como en Lightroom (plan del rediseño, hallazgos/editores-rediseno.md, 3.2). Cada pieza
// marca una zona (toda la imagen, lo pintado con un pincel, un degradado, un
// color, un rango de luz, lo que no es el fondo de un croma —verde, azul…—,
// o el primer plano o un objeto que separa la IA:
// recorte-sujeto.js y objeto-por-toque.js) y dice cómo se combina con lo que
// marcan las de abajo:
//   - sumar: se agrega (source-over);
//   - restar: se quita (destination-out);
//   - intersecar: queda solo lo que también marca esta (destination-in).
// Cada pieza se puede invertir y apagar, y la máscara entera se puede
// invertir al final. El resultado es un lienzo cuyo alfa dice cuánto se
// aplica la capa en cada punto (blanco opaco = entera, transparente = nada),
// o null si queda "toda la imagen" sin retoques (el caso común, sin costo).
//
// Antes (hasta la versión 7) era una base, el pincel para sumar y quitar, e
// invertir. desdeV7 lo convierte en piezas con el mismo resultado:
// (base ∪ sumado) − quitado = [base (sumar), pincel (sumar), pincel (restar)].
//
// Color, luz y la IA de una pieza pueden medir sobre la foto o, en un objeto
// (texto o imagen), sobre su contenido ("fuente": 'capa'). Entonces la pieza
// se calcula en las coordenadas del contenido y quien llama la coloca con la
// misma transformación ("colocar"): sigue al objeto si se mueve, escala o
// gira. Toda la imagen, el pincel y el degradado van en las coordenadas de la
// foto.
//
// Color, luz y la IA se calculan píxel a píxel: van a una caché por pieza
// ("memo", una por capa) y solo se rehacen si cambia su configuración o la
// imagen. Con memo, además, se reusan los lienzos de un render al siguiente
// (crear lienzos por fotograma es lo que traba la vista previa); sin memo (la
// descarga a resolución completa) todo sale en lienzos nuevos.

window.MascarasCapa = (function () {
  'use strict';

  const TIPOS = ['toda', 'pincel', 'degradado', 'color', 'luz', 'croma', 'sujeto', 'objeto'];
  const OPERACIONES = ['sumar', 'restar', 'intersecar'];
  // Las piezas que miden algo sobre una imagen (y por eso tienen "fuente").
  const MIDEN = ['color', 'luz', 'croma', 'sujeto', 'objeto'];
  // Rangos rápidos de "Luz", en 0..100 de luminosidad.
  const RANGOS_LUZ = { sombras: [0, 35], medios: [35, 65], luces: [65, 100] };
  const COMPOSICION = { sumar: 'source-over', restar: 'destination-out', intersecar: 'destination-in' };

  // Ids únicos en toda la página: el historial copia el lienzo de cada pincel
  // por el id de su pieza.
  let contadorPiezas = 0;
  function nuevoIdPieza() {
    contadorPiezas += 1;
    return 'pz-' + contadorPiezas.toString(36) + Math.random().toString(36).slice(2, 7);
  }

  // Una pieza nueva de un tipo, con su configuración de partida.
  function pieza(tipo, extra = {}) {
    const p = { id: nuevoIdPieza(), tipo, op: 'sumar', invertir: false, visible: true };
    if (tipo === 'pincel') { p.trazo = null; p._version = 0; } // trazo: lienzo a tamaño de trabajo, alfa = cuánto
    if (tipo === 'degradado') p.degradado = { tipo: 'lineal', a: null, b: null }; // a/b en 0..1 del ancho/alto
    if (tipo === 'color') { p.color = null; p.tolerancia = 30; } // color: {r,g,b}
    // color: el del fondo, {r,g,b}; recorte: cuánto se achica el borde, en
    // píxeles de la imagen original (ver recortarBorde).
    if (tipo === 'croma') { p.color = null; p.tolerancia = 25; p.suavidad = 10; p.recorte = 1; }
    if (tipo === 'luz') p.luz = { desde: RANGOS_LUZ.sombras[0], hasta: RANGOS_LUZ.sombras[1], suavidad: 15 };
    // [{x,y}] en 0..1: lo que se marcó; la pieza es la unión de esos objetos
    // (hasta el 2026-10-08 era uno solo, "punto").
    if (tipo === 'objeto') p.puntos = [];
    if (MIDEN.includes(tipo)) p.fuente = 'foto'; // 'foto' | 'capa'
    return Object.assign(p, extra);
  }

  function neutra() {
    return { piezas: [pieza('toda')], invertir: false };
  }

  function puntosValidos(lista) {
    return lista.filter((o) => o && Number.isFinite(o.x) && Number.isFinite(o.y)).map((o) => ({ x: o.x, y: o.y }));
  }

  // Completa una pieza guardada con lo que le falte. El lienzo del pincel
  // (trazo) pasa tal cual.
  function normalizarPieza(g) {
    const tipo = TIPOS.includes(g && g.tipo) ? g.tipo : 'toda';
    const n = pieza(tipo);
    const p = {
      ...n,
      id: typeof g.id === 'string' && g.id ? g.id : n.id,
      op: OPERACIONES.includes(g.op) ? g.op : 'sumar',
      invertir: !!g.invertir,
      visible: g.visible !== false,
    };
    if (tipo === 'pincel') { p.trazo = g.trazo || null; p._version = g._version || 0; }
    if (tipo === 'degradado') p.degradado = { ...n.degradado, ...(g.degradado || {}) };
    if (tipo === 'croma') {
      p.color = g.color || null;
      p.tolerancia = Number.isFinite(g.tolerancia) ? g.tolerancia : n.tolerancia;
      p.suavidad = Number.isFinite(g.suavidad) ? g.suavidad : n.suavidad;
      p.recorte = Number.isFinite(g.recorte) ? g.recorte : n.recorte;
    }
    if (tipo === 'color') {
      p.color = g.color || null;
      p.tolerancia = Number.isFinite(g.tolerancia) ? g.tolerancia : n.tolerancia;
    }
    if (tipo === 'luz') p.luz = { ...n.luz, ...(g.luz || {}) };
    if (tipo === 'objeto') p.puntos = puntosValidos(Array.isArray(g.puntos) ? g.puntos : [g.punto]);
    if (MIDEN.includes(tipo)) p.fuente = g.fuente === 'capa' ? 'capa' : 'foto';
    return p;
  }

  // Una máscara de la versión 8. Una que no trae piezas (la de un elemento
  // nuevo de un proyecto convertido, mascara: null) es la neutra.
  function normalizar(m) {
    if (!m || !Array.isArray(m.piezas)) return neutra();
    return { piezas: m.piezas.map(normalizarPieza), invertir: !!m.invertir };
  }

  // La máscara de la versión 7 (base + pincel + invertir) como piezas, con
  // el mismo resultado. sumar/quitar: los lienzos del pincel, o null.
  function desdeV7(v, sumar, quitar) {
    v = v || {};
    const base = ['toda', 'nada', 'degradado', 'color', 'luz', 'sujeto', 'fondo'].includes(v.base) ? v.base : 'toda';
    const fuente = v.fuente === 'capa' ? 'capa' : 'foto';
    const piezas = [];
    if (base === 'toda') piezas.push(pieza('toda'));
    else if (base === 'degradado') piezas.push(pieza('degradado', { degradado: { tipo: 'lineal', a: null, b: null, ...(v.degradado || {}) } }));
    else if (base === 'color') {
      piezas.push(pieza('color', { color: v.color || null, tolerancia: Number.isFinite(v.tolerancia) ? v.tolerancia : 30, fuente }));
    } else if (base === 'luz') {
      piezas.push(pieza('luz', { luz: { ...pieza('luz').luz, ...(v.luz || {}) }, fuente }));
    } else if (base === 'sujeto' || base === 'fondo') {
      // "Fondo" era todo menos el sujeto (o el objeto): la misma pieza, invertida.
      const puntos = puntosValidos([v.objeto]);
      piezas.push(puntos.length ? pieza('objeto', { puntos, fuente }) : pieza('sujeto', { fuente }));
      if (base === 'fondo') piezas[0].invertir = true;
    }
    // 'nada': sin pieza de base; solo cuenta lo pintado.
    if (sumar) piezas.push(pieza('pincel', { trazo: sumar }));
    if (quitar) piezas.push(pieza('pincel', { trazo: quitar, op: 'restar' }));
    return { piezas, invertir: !!v.invertir };
  }

  // Islas: grupos de píxeles conectados (4 vecinos) con alfa >= umbral. Para
  // cada píxel, el número de su isla (-1 si no es de ninguna), y el tamaño de
  // cada isla. La usa la limpieza de manchas del generador de stickers.
  function islas(datos, ancho, alto, umbral) {
    const total = ancho * alto;
    const visitado = new Uint8Array(total);
    const etiquetas = new Int32Array(total).fill(-1);
    const tamanos = [];
    const pila = new Int32Array(total);
    for (let inicio = 0; inicio < total; inicio++) {
      if (visitado[inicio] || datos[inicio * 4 + 3] < umbral) { visitado[inicio] = 1; continue; }
      const etiqueta = tamanos.length;
      let tope = 0;
      pila[tope++] = inicio;
      visitado[inicio] = 1;
      let tamano = 0;
      while (tope > 0) {
        const p = pila[--tope];
        etiquetas[p] = etiqueta;
        tamano++;
        const x = p % ancho;
        const y = (p / ancho) | 0;
        if (x > 0 && !visitado[p - 1] && datos[(p - 1) * 4 + 3] >= umbral) { visitado[p - 1] = 1; pila[tope++] = p - 1; }
        if (x < ancho - 1 && !visitado[p + 1] && datos[(p + 1) * 4 + 3] >= umbral) { visitado[p + 1] = 1; pila[tope++] = p + 1; }
        if (y > 0 && !visitado[p - ancho] && datos[(p - ancho) * 4 + 3] >= umbral) { visitado[p - ancho] = 1; pila[tope++] = p - ancho; }
        if (y < alto - 1 && !visitado[p + ancho] && datos[(p + ancho) * 4 + 3] >= umbral) { visitado[p + ancho] = 1; pila[tope++] = p + ancho; }
      }
      tamanos.push(tamano);
    }
    return { etiquetas, tamanos };
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

  // Croma: como en video, se compara el TONO del píxel con el del fondo, no
  // el brillo. En el plano Cb/Cr (YCbCr) el tono es la DIRECCIÓN y la
  // intensidad del color es el largo, que se achica con la sombra: por eso se
  // compara el ángulo, y así las sombras y los reflejos sobre el fondo verde
  // también son fondo. Un píxel casi sin color (gris, blanco, negro) nunca es
  // fondo. Marca lo que NO es el fondo (la persona, el objeto), como "Primer
  // plano": nada hasta "tolerancia" grados de diferencia, entera desde
  // "tolerancia" + "suavidad", y un borde suave en el medio.
  function cbcr(r, g, b) {
    return { cb: -0.168736 * r - 0.331264 * g + 0.5 * b, cr: 0.5 * r - 0.418688 * g - 0.081312 * b };
  }
  function dibujarCroma(origen, p, ancho, alto, destino) {
    const fondo = cbcr(p.color.r, p.color.g, p.color.b);
    const largoFondo = Math.max(1, Math.hypot(fondo.cb, fondo.cr));
    const anguloFondo = Math.atan2(fondo.cr, fondo.cb);
    const desde = p.tolerancia * 0.8, hasta = desde + Math.max(0.5, p.suavidad * 0.8);
    return porPixel(origen, ancho, alto, destino, (r, g, b) => {
      const c = cbcr(r, g, b);
      // Cuánto color tiene, comparado con el fondo: 0 por debajo del 15 %
      // (no es fondo), 1 desde el 30 %.
      const conColor = Math.min(1, Math.max(0, (Math.hypot(c.cb, c.cr) / largoFondo - 0.15) / 0.15));
      if (conColor <= 0) return 1;
      let dif = Math.abs(Math.atan2(c.cr, c.cb) - anguloFondo) * 180 / Math.PI;
      if (dif > 180) dif = 360 - dif;
      const lejos = dif <= desde ? 0 : (dif >= hasta ? 1 : (dif - desde) / (hasta - desde));
      return 1 - conColor * (1 - lejos);
    });
  }

  // Recortar el borde de lo marcado unos píxeles hacia adentro: el contorno de
  // una persona sobre un croma es mezcla de los dos colores (por el
  // suavizado y la compresión de la foto), su tono queda lejos del fondo y
  // quedaba como un hilo verde alrededor. Se AGRANDA el fondo (lo no marcado,
  // dibujado corrido en varias direcciones) en vez de achicar lo marcado: así
  // lo que toca el borde de la imagen (los pies de una persona) no se
  // recorta, porque lo que se corre fuera del lienzo no se dibuja.
  function recortarBorde(lienzoMarcado, radio) {
    const ancho = lienzoMarcado.width, alto = lienzoMarcado.height;
    const fondo = document.createElement('canvas');
    fondo.width = ancho; fondo.height = alto;
    const cf = fondo.getContext('2d');
    cf.fillStyle = '#fff';
    cf.fillRect(0, 0, ancho, alto);
    cf.globalCompositeOperation = 'destination-out';
    cf.drawImage(lienzoMarcado, 0, 0);
    const agrandado = document.createElement('canvas');
    agrandado.width = ancho; agrandado.height = alto;
    const ca = agrandado.getContext('2d');
    ca.drawImage(fondo, 0, 0);
    const pasos = Math.max(8, Math.ceil(radio * 4));
    for (const r of radio > 1.5 ? [radio / 2, radio] : [radio]) {
      for (let i = 0; i < pasos; i++) {
        const a = (i / pasos) * Math.PI * 2;
        ca.drawImage(fondo, Math.cos(a) * r, Math.sin(a) * r);
      }
    }
    const ctx = lienzoMarcado.getContext('2d');
    ctx.clearRect(0, 0, ancho, alto);
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, ancho, alto);
    ctx.globalCompositeOperation = 'destination-out';
    ctx.drawImage(agrandado, 0, 0);
    ctx.globalCompositeOperation = 'source-over';
    return lienzoMarcado;
  }

  // El color del fondo de un croma: el más repetido en los bordes de la
  // imagen, entre los que tienen color (un fondo verde o azul); si los bordes
  // no tienen color, el más repetido igual. Se mira una copia chica: leer la
  // imagen original con getImageData la pasa a la memoria común (ver el
  // gotero del Editor avanzado).
  function detectarCroma(origen) {
    const escala = Math.min(1, 256 / Math.max(origen.width, origen.height));
    const w = Math.max(2, Math.round(origen.width * escala)), h = Math.max(2, Math.round(origen.height * escala));
    const c = document.createElement('canvas');
    c.width = w; c.height = h;
    const ctx = c.getContext('2d', { willReadFrequently: true });
    ctx.drawImage(origen, 0, 0, w, h);
    const d = ctx.getImageData(0, 0, w, h).data;
    const grupos = new Map(); // tono redondeado -> { n, r, g, b, conColor }
    const anotar = (x, y) => {
      const i = (y * w + x) * 4;
      if (d[i + 3] < 128) return;
      const t = cbcr(d[i], d[i + 1], d[i + 2]);
      const clave = Math.round(t.cb / 10) + ',' + Math.round(t.cr / 10);
      const gr = grupos.get(clave) || { n: 0, r: 0, g: 0, b: 0, conColor: Math.hypot(t.cb, t.cr) > 20 };
      gr.n += 1; gr.r += d[i]; gr.g += d[i + 1]; gr.b += d[i + 2];
      grupos.set(clave, gr);
    };
    for (let x = 0; x < w; x++) { anotar(x, 0); anotar(x, h - 1); }
    for (let y = 1; y < h - 1; y++) { anotar(0, y); anotar(w - 1, y); }
    const todos = [...grupos.values()];
    if (!todos.length) return null;
    const conColor = todos.filter((g) => g.conColor);
    const elegido = (conColor.length ? conColor : todos).sort((a, b) => b.n - a.n)[0];
    return { r: Math.round(elegido.r / elegido.n), g: Math.round(elegido.g / elegido.n), b: Math.round(elegido.b / elegido.n) };
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

  // ¿Qué marca la máscara entera? Se sigue pieza por pieza: 'llena' (toda la
  // imagen), 'vacia' o 'parcial'. Solo 'toda' sabe lo que marca sin calcular.
  function estadoFinal(m) {
    let estado = 'vacia';
    for (const p of m.piezas) {
      if (!p.visible) continue;
      const marca = p.tipo !== 'toda' ? 'parcial' : (p.invertir ? 'nada' : 'todo');
      if (p.op === 'sumar') {
        if (marca === 'todo') estado = 'llena';
        else if (marca === 'parcial' && estado !== 'llena') estado = 'parcial';
      } else if (p.op === 'restar') {
        if (marca === 'todo') estado = 'vacia';
        else if (marca === 'parcial' && estado !== 'vacia') estado = 'parcial';
      } else if (marca === 'nada') estado = 'vacia';
      else if (marca === 'parcial' && estado !== 'vacia') estado = 'parcial';
    }
    return estado;
  }

  // ¿La máscara deja la capa en toda la imagen sin retoques? Entonces no hace
  // falta ningún lienzo.
  function esTodaLaImagen(m) {
    m = normalizar(m);
    return !m.invertir && estadoFinal(m) === 'llena';
  }

  // ¿Es la máscara de partida (una sola pieza "Toda la imagen")? Una pieza
  // nueva que suma la reemplaza, como antes reemplazaba a la base.
  function esSoloTodaLaImagen(m) {
    return !!m && !m.invertir && m.piezas.length === 1 && m.piezas[0].tipo === 'toda'
      && m.piezas[0].op === 'sumar' && !m.piezas[0].invertir && m.piezas[0].visible;
  }

  function invertirEn(origen, ancho, alto, reusar) {
    const c = lienzo(ancho, alto, reusar);
    const ctx = c.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, ancho, alto);
    ctx.globalCompositeOperation = 'destination-out';
    ctx.drawImage(origen, 0, 0, ancho, alto);
    ctx.globalCompositeOperation = 'source-over';
    return c;
  }

  // Lo que mide una pieza de color, luz o IA, a tamaño "ancho x alto" (el de
  // trabajo, o el del contenido si se mide sobre él), cacheado en su memo
  // mientras no cambie nada de lo que lo define.
  // escala: cuántos píxeles de "ancho x alto" mide un píxel de la imagen
  // original (el recorte del croma va en píxeles originales).
  function medicion(p, origen, ia, ancho, alto, cache, escala = 1) {
    const clave = p.tipo === 'color' ? JSON.stringify([p.color, p.tolerancia])
      : p.tipo === 'croma' ? 'croma' + JSON.stringify([p.color, p.tolerancia, p.suavidad, p.recorte])
      : (p.tipo === 'luz' ? JSON.stringify(p.luz) : p.tipo);
    const desde = p.tipo === 'sujeto' || p.tipo === 'objeto' ? ia : origen;
    if (cache && cache.base && cache.clave === clave && cache.desde === desde
      && cache.base.width === ancho && cache.base.height === alto) return cache.base;
    const destino = lienzo(ancho, alto, cache ? cache.base : null);
    if (p.tipo === 'color') dibujarColor(origen, p.color, p.tolerancia, ancho, alto, destino);
    else if (p.tipo === 'croma') {
      dibujarCroma(origen, p, ancho, alto, destino);
      // En píxeles de la imagen original: igual a cualquier resolución de la
      // vista previa y en la descarga.
      const radio = (p.recorte || 0) * escala;
      if (radio > 0.05) recortarBorde(destino, radio);
    }
    else if (p.tipo === 'luz') dibujarLuz(origen, p.luz, ancho, alto, destino);
    else destino.getContext('2d').drawImage(ia, 0, 0, ancho, alto);
    if (cache) Object.assign(cache, { base: destino, clave, desde });
    return destino;
  }

  // El lienzo de una pieza a tamaño final (ya invertida y colocada), para
  // combinarlo con las de abajo: { lienzo, escalar } (escalar: dibujarlo
  // estirado a ancho x alto, el del pincel), o null si no marca nada todavía
  // (sin pintar, o la IA sin terminar).
  function lienzoDePieza(p, ancho, alto, opciones, cache) {
    const reusar = (nombre) => (cache ? cache[nombre] : null);
    const guardar = (nombre, c) => { if (cache) cache[nombre] = c; return c; };
    if (p.tipo === 'pincel') {
      if (!p.trazo) return null;
      if (!p.invertir) return { lienzo: p.trazo, escalar: true };
      return { lienzo: guardar('invertida', invertirEn(p.trazo, ancho, alto, reusar('invertida'))) };
    }
    // Degradado sin trazar y color sin elegir: toda la imagen (así elegirlos
    // no borra la capa, como antes con la base).
    if (p.tipo === 'degradado' || ((p.tipo === 'color' || p.tipo === 'croma') && !p.color)) {
      const c = guardar('propio', lienzo(ancho, alto, reusar('propio')));
      const ctx = c.getContext('2d');
      if (p.tipo === 'degradado') dibujarDegradado(ctx, p.degradado, ancho, alto);
      else { ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, ancho, alto); }
      return { lienzo: p.invertir ? guardar('invertida', invertirEn(c, ancho, alto, reusar('invertida'))) : c };
    }
    // Color, luz y la IA: sobre la foto o sobre el contenido del objeto.
    const { origen = null, ia = null, colocar = null, escala = 1 } = opciones.medir ? opciones.medir(p) : {};
    if ((p.tipo === 'sujeto' || p.tipo === 'objeto') && !ia) return null;
    if ((p.tipo === 'color' || p.tipo === 'luz' || p.tipo === 'croma') && !origen) return null;
    const anchoBase = colocar ? origen.width : ancho, altoBase = colocar ? origen.height : alto;
    let base = medicion(p, origen, ia, anchoBase, altoBase, cache, escala);
    // Se invierte antes de colocarla: afuera del contenido no marca nada.
    if (p.invertir) base = guardar('invertida', invertirEn(base, anchoBase, altoBase, reusar('invertida')));
    if (!colocar) return { lienzo: base };
    return { lienzo: guardar('puesta', colocar(base, lienzo(ancho, alto, reusar('puesta')))) };
  }

  // m: la máscara. opciones:
  //   memo: caché de la capa (o null, para la descarga);
  //   medir(pieza) -> { origen, ia, colocar, escala }: sobre qué mide una
  //     pieza de color, luz, croma o IA (escala: píxeles de "origen" por
  //     píxel de la imagen original, 1 si no se dice). origen: la imagen
  //     (la foto a tamaño de trabajo, o el contenido a su tamaño); ia: el lienzo que dio la IA para esa pieza, o
  //     null si todavía no está; colocar(lienzo, destino): solo si se mide
  //     sobre el contenido, lo dibuja en destino transformado como él.
  function crear(m, ancho, alto, opciones = {}) {
    m = normalizar(m);
    if (esTodaLaImagen(m)) return null;
    const memo = opciones.memo || null;
    if (memo && !memo.piezas) memo.piezas = new Map();
    const final = lienzo(ancho, alto, memo ? memo.final : null);
    if (memo) memo.final = final;
    const ctx = final.getContext('2d');
    for (const p of m.piezas) {
      if (!p.visible) continue;
      if (p.tipo === 'toda') {
        const todo = !p.invertir;
        if (p.op === 'sumar' && todo) { ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, ancho, alto); }
        else if ((p.op === 'restar' && todo) || (p.op === 'intersecar' && !todo)) ctx.clearRect(0, 0, ancho, alto);
        continue;
      }
      let cache = null;
      if (memo) {
        cache = memo.piezas.get(p.id);
        if (!cache) { cache = {}; memo.piezas.set(p.id, cache); }
      }
      const resultado = lienzoDePieza(p, ancho, alto, opciones, cache);
      if (!resultado) {
        // Lo que todavía no marca nada: si tenía que intersecar, no queda nada.
        if (p.op === 'intersecar') ctx.clearRect(0, 0, ancho, alto);
        continue;
      }
      ctx.globalCompositeOperation = COMPOSICION[p.op];
      if (resultado.escalar) ctx.drawImage(resultado.lienzo, 0, 0, ancho, alto);
      else ctx.drawImage(resultado.lienzo, 0, 0);
      ctx.globalCompositeOperation = 'source-over';
    }
    if (!m.invertir) return final;
    const invertida = invertirEn(final, ancho, alto, memo ? memo.invertida : null);
    if (memo) memo.invertida = invertida;
    return invertida;
  }

  return {
    TIPOS, OPERACIONES, MIDEN, RANGOS_LUZ,
    pieza, neutra, normalizar, desdeV7, crear, esTodaLaImagen, esSoloTodaLaImagen, islas, detectarCroma,
  };
})();

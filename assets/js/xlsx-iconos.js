/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Iconos de los conjuntos de iconos del formato condicional de Excel (etapa 6), dibujados como vectores propios sobre una cuadrícula
// de 24 × 24 que se escala al tamaño pedido (16 px en pantalla, 12 pt impreso). Las formas son geometría simple (flechas, círculos, rombo,
// triángulo, símbolos, bandera, escalones, cuartos) hechas aquí; de Excel solo se tomaron medidas: qué forma lleva cada posición de cada
// conjunto y los colores (medidos en el PDF que exporta Excel). Sin DOM: solo usa la API de contexto 2D (Canvas o el lienzo de PDF).
//
// Uso: XlsxIconos.dibujar(ctx, '3Arrows', indice, x, y, tam)  (indice 0 = el más bajo). Devuelve false si el conjunto no se conoce.
// Los tres conjuntos de la extensión x14 de Excel 2010 (3Stars, 3Triangles, 5Boxes) también están: medidos en un libro hecho por Excel (36-x14).

(function () {
  const C = {
    rojo: ['#D65532', '#9E381C'], amarillo: ['#EAC282', '#A4802B'], verde: ['#68A490', '#326F5B'], gris: ['#808080', '#646262'],
    azul: ['#4A7DB1', '#335F8C'], grisClaro: ['#B3B3B3', '#727272'], negro: ['#505050', '#2E2E2E'], plata: ['#B1B1B1', '#6D6D6D'],
    rosa: ['#EDB2A3', '#C49284'], naranja: ['#D86344', '#B14629'],
  };
  const S3 = [C.rojo, C.amarillo, C.verde];

  // Flecha hacia arriba en la cuadrícula; `grados` la gira alrededor del centro.
  const FLECHA = [[12, 1.5], [22, 11.5], [15.2, 11.5], [15.2, 22.5], [8.8, 22.5], [8.8, 11.5], [2, 11.5]];

  function poligono(ctx, pts) { ctx.beginPath(); pts.forEach((p, i) => (i ? ctx.lineTo(p[0], p[1]) : ctx.moveTo(p[0], p[1]))); ctx.closePath(); }
  function relleno(ctx, col) { ctx.fillStyle = col[0]; ctx.fill(); ctx.strokeStyle = col[1]; ctx.lineWidth = 1; ctx.lineJoin = 'round'; ctx.stroke(); }
  function circulo(ctx, cx, cy, r, col) { ctx.beginPath(); ctx.arc(cx, cy, r, 0, Math.PI * 2); ctx.closePath(); relleno(ctx, col); }

  const flecha = (grados, col) => (ctx) => {
    ctx.save(); ctx.translate(12, 12); ctx.rotate(grados * Math.PI / 180); ctx.translate(-12, -12);
    poligono(ctx, FLECHA); relleno(ctx, col);
    ctx.restore();
  };
  const bandera = (col) => (ctx) => {
    ctx.fillStyle = '#707070'; ctx.fillRect(5, 1, 2.2, 22);
    poligono(ctx, [[7.5, 2], [21.5, 8.5], [7.5, 15]]); relleno(ctx, col);
  };
  const luz = (col) => (ctx) => circulo(ctx, 12, 12, 10.5, col);
  const luzOscura = (col) => (ctx) => {
    ctx.beginPath(); ctx.moveTo(4, 1.5); ctx.lineTo(20, 1.5); ctx.quadraticCurveTo(22.5, 1.5, 22.5, 4); ctx.lineTo(22.5, 20); ctx.quadraticCurveTo(22.5, 22.5, 20, 22.5);
    ctx.lineTo(4, 22.5); ctx.quadraticCurveTo(1.5, 22.5, 1.5, 20); ctx.lineTo(1.5, 4); ctx.quadraticCurveTo(1.5, 1.5, 4, 1.5); ctx.closePath();
    ctx.fillStyle = '#505050'; ctx.fill();
    circulo(ctx, 12, 12, 7.5, [col[0], col[0]]);
  };
  const rombo = (col) => (ctx) => { poligono(ctx, [[12, 1], [23, 12], [12, 23], [1, 12]]); relleno(ctx, col); };
  const triangulo = (col) => (ctx) => { poligono(ctx, [[12, 2], [23, 21], [1, 21]]); relleno(ctx, col); };

  // símbolos: cruz, exclamación y marca de verificación (en color o blancos sobre un círculo)
  const cruz = (ctx, color) => { ctx.strokeStyle = color; ctx.lineWidth = 4.2; ctx.lineCap = 'butt'; ctx.beginPath(); ctx.moveTo(6, 6); ctx.lineTo(18, 18); ctx.moveTo(18, 6); ctx.lineTo(6, 18); ctx.stroke(); };
  const exclama = (ctx, color) => { ctx.fillStyle = color; poligono(ctx, [[9.2, 3], [14.8, 3], [13.8, 14.5], [10.2, 14.5]]); ctx.fill(); ctx.beginPath(); ctx.arc(12, 18.6, 2.5, 0, Math.PI * 2); ctx.fill(); };
  const marca = (ctx, color) => { ctx.strokeStyle = color; ctx.lineWidth = 4; ctx.lineCap = 'butt'; ctx.lineJoin = 'miter'; ctx.beginPath(); ctx.moveTo(4.5, 12.5); ctx.lineTo(10, 18); ctx.lineTo(20, 5.5); ctx.stroke(); };
  const GLIFO = [cruz, exclama, marca];
  const simbolo = (col, k) => (ctx) => { circulo(ctx, 12, 12, 10.5, col); GLIFO[k](ctx, '#FFFFFF'); };
  const simboloSolo = (col, k) => (ctx) => GLIFO[k](ctx, col[0]);

  // valoración: escalera de cuatro escalones, `n` en azul y el resto en gris
  const escalera = (n) => (ctx) => {
    for (let k = 0; k < 4; k++) {
      const x = 1 + 5 * k, alto = 10 + 4 * k;
      ctx.fillStyle = k < n ? C.azul[0] : C.grisClaro[0];
      ctx.fillRect(x, 22.5 - alto, k === 3 ? 6 : 5.6, alto);
    }
  };
  // cuartos: círculo blanco con borde oscuro y `n` cuartos oscuros desde las 12 en punto, en el sentido del reloj
  const cuartos = (n) => (ctx) => {
    ctx.beginPath(); ctx.arc(12, 12, 10.5, 0, Math.PI * 2); ctx.closePath();
    ctx.fillStyle = '#FFFFFF'; ctx.fill();
    if (n > 0) {
      ctx.beginPath(); ctx.moveTo(12, 12); ctx.arc(12, 12, 10.5, -Math.PI / 2, -Math.PI / 2 + n * Math.PI / 2); ctx.closePath();
      ctx.fillStyle = C.negro[0]; ctx.fill();
    }
    ctx.beginPath(); ctx.arc(12, 12, 10.5, 0, Math.PI * 2); ctx.closePath();
    ctx.strokeStyle = C.negro[0]; ctx.lineWidth = 1.2; ctx.stroke();
  };

  // estrellas (x14): vacía, mitad izquierda dorada y llena
  const ESTRELLA = Array.from({ length: 10 }, (_, k) => { const r = k % 2 ? 5.3 : 11.5, a = -Math.PI / 2 + k * Math.PI / 5; return [12 + r * Math.cos(a), 12.4 + r * Math.sin(a)]; });
  const estrella = (n) => (ctx) => {
    const pintar = (col) => { poligono(ctx, ESTRELLA); relleno(ctx, col); };
    const mitad = (x0, col) => { ctx.save(); ctx.beginPath(); ctx.rect(x0, 0, 12, 24); ctx.clip(); pintar(col); ctx.restore(); };
    if (n === 2) pintar(C.amarillo); else if (n === 0) pintar(['#FFFFFF', C.grisClaro[1]]);
    else { mitad(0, C.amarillo); mitad(12, ['#FFFFFF', C.grisClaro[1]]); }
  };
  // triángulos (x14): hacia abajo rojo, raya amarilla y hacia arriba verde
  const trianguloAbajo = (ctx) => { poligono(ctx, [[1, 5.5], [23, 5.5], [12, 18.5]]); relleno(ctx, C.rojo); };
  const raya = (ctx) => { ctx.beginPath(); ctx.rect(1.5, 8.5, 21, 7); relleno(ctx, C.amarillo); };
  const trianguloArriba = (ctx) => { poligono(ctx, [[12, 5], [23, 18], [1, 18]]); relleno(ctx, C.verde); };
  // cajas (x14): cuadrícula de 2 × 2 con `n` cajas azules, de abajo hacia arriba y de izquierda a derecha, y el resto grises
  const cajas = (n) => (ctx) => {
    const orden = [[1.5, 12.5], [12.5, 12.5], [1.5, 1.5], [12.5, 1.5]]; // esquina de cada caja: inferior izquierda, inferior derecha, superior izquierda, superior derecha
    for (const llenas of [false, true]) {
      orden.forEach(([x, y], k) => {
        if ((k < n) !== llenas) return;
        const col = llenas ? C.azul : C.grisClaro;
        ctx.fillStyle = col[0]; ctx.fillRect(x, y, 10.5, 10.5);
        ctx.strokeStyle = col[1]; ctx.lineWidth = 1; ctx.strokeRect(x, y, 10.5, 10.5);
      });
    }
  };

  const FLECHAS3 = [180, 90, 0], FLECHAS4 = [180, 135, 45, 0], FLECHAS5 = [180, 135, 90, 45, 0];
  const colFlecha = (k, n) => (k === 0 ? C.rojo : k === n - 1 ? C.verde : C.amarillo);
  const flechas = (angulos, gris) => angulos.map((a, k) => flecha(a, gris ? C.gris : colFlecha(k, angulos.length)));

  const CONJUNTOS = {
    '3Arrows': flechas(FLECHAS3), '3ArrowsGray': flechas(FLECHAS3, true),
    '4Arrows': flechas(FLECHAS4), '4ArrowsGray': flechas(FLECHAS4, true),
    '5Arrows': flechas(FLECHAS5), '5ArrowsGray': flechas(FLECHAS5, true),
    '3Flags': S3.map(bandera),
    '3TrafficLights1': S3.map(luz),
    '3TrafficLights2': S3.map(luzOscura),
    '3Signs': [rombo(C.rojo), triangulo(C.amarillo), luz(C.verde)],
    '3Symbols': S3.map((c, k) => simbolo(c, k)),
    '3Symbols2': S3.map((c, k) => simboloSolo(c, k)),
    '4RedToBlack': [luz(C.negro), luz(C.plata), luz(C.rosa), luz(C.naranja)],
    '4TrafficLights': [luz(C.negro), luz(C.rojo), luz(C.amarillo), luz(C.verde)],
    '4Rating': [1, 2, 3, 4].map(escalera),
    '5Rating': [0, 1, 2, 3, 4].map(escalera),
    '5Quarters': [0, 1, 2, 3, 4].map(cuartos),
    '3Stars': [0, 1, 2].map(estrella),
    '3Triangles': [trianguloAbajo, raya, trianguloArriba],
    '5Boxes': [0, 1, 2, 3, 4].map(cajas),
  };
  function dibujar(ctx, conjunto, indice, x, y, tam) {
    const lista = CONJUNTOS[conjunto];
    if (!lista || !(tam > 0)) return false;
    const f = lista[Math.max(0, Math.min(lista.length - 1, indice | 0))];
    ctx.save();
    ctx.translate(x, y); ctx.scale(tam / 24, tam / 24);
    try { f(ctx); } finally { ctx.restore(); }
    return true;
  }

  globalThis.XlsxIconos = { dibujar, conjuntos: Object.keys(CONJUNTOS) };
})();

/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// El motor de filtros (procesar-filtros.js) en la placa de video, con WebGL2.
// Medido el 2026-10-05 con una foto de 24 MP: el bucle por píxel en
// JavaScript tardaba ~450 ms por capa y por fotograma; en un shader, el mismo
// trabajo lo hace la placa en paralelo y los píxeles nunca vuelven a
// JavaScript (leerlos con getImageData era casi la mitad del tiempo en la
// vista "Rápida").
//
// Acá solo está la parte de WebGL: el pipeline (qué se aplica y en qué orden,
// los números de cada deslizador, las cachés) lo decide procesar-filtros.js,
// que llama a estas tres pasadas y, si devuelven false, usa su bucle en
// JavaScript. El shader principal repite ese bucle paso por paso, con los
// mismos recortes a 0-255 y el mismo redondeo de un Uint8ClampedArray
// (roundEven): un cambio en uno va también en el otro.
//
// Necesita, cargado antes: ajuste-por-color.js (la tabla de pesos por color).
window.FiltrosGPU = (function () {
  let gl = null, lienzo = null, programas = null, ladoMaximo = 0;
  let roto = false, perdidas = 0;
  const texturas = []; // por unidad: { textura, ancho, alto }

  const VERTICES = `#version 300 es
void main() {
  // Un triángulo que cubre toda la vista.
  vec2 v = vec2(float((gl_VertexID << 1) & 2), float(gl_VertexID & 2));
  gl_Position = vec4(v * 2.0 - 1.0, 0.0, 1.0);
}`;

  // Común a las tres pasadas: coordenadas de la imagen (origen arriba a la
  // izquierda, como el índice del bucle en JS) y la aberración cromática, que
  // corre antes que todo y la usan la principal y la de zonas brillantes.
  const COMUN = `#version 300 es
precision highp float;
precision highp int;
precision highp sampler2D;
uniform ivec2 uTam;
uniform sampler2D uFuente;
uniform bool uAberracion, uAberracionMatiz;
uniform ivec2 uDesp;
uniform mat3 uMAberracion;
out vec4 salida;
const vec3 LUMA = vec3(0.299, 0.587, 0.114);

ivec2 pixel() { return ivec2(int(gl_FragCoord.x), uTam.y - 1 - int(gl_FragCoord.y)); }
vec4 leer(sampler2D t, ivec2 p) { return floor(texelFetch(t, p, 0) * 255.0 + 0.5); }
// Lo que hace guardar un número en un Uint8ClampedArray.
vec3 u8(vec3 v) { return roundEven(clamp(v, 0.0, 255.0)); }
vec3 c255(vec3 v) { return clamp(v, 0.0, 255.0); }

vec4 pixelEntrada(ivec2 p) {
  vec4 c = leer(uFuente, p);
  if (uAberracion) {
    vec3 cr = leer(uFuente, clamp(p + uDesp, ivec2(0), uTam - 1)).rgb;
    vec3 cb = leer(uFuente, clamp(p - uDesp, ivec2(0), uTam - 1)).rgb;
    if (uAberracionMatiz) {
      c.r = u8(uMAberracion * cr).r;
      c.b = u8(uMAberracion * cb).b;
    } else {
      c.r = cr.r;
      c.b = cb.b;
    }
  }
  return c;
}
`;

  const PRINCIPAL = COMUN + `
uniform sampler2D uRef, uClaridad, uNitidez, uDispersion, uPesos;
uniform float uExposicion, uContraste, uSombras, uReflejos, uBlancos, uNegros;
uniform bool uBalance, uLuces, uTono, uSaturar, uPorColor, uBn, uMezclaGris, uMatizDispersion;
uniform vec3 uK, uDelta, uW;
uniform mat3 uMTono, uMDispersion;
uniform float uSat, uIntensidad, uSumaCanales, uSepia, uInvertir, uAmpDispersion;
uniform float uClaridadAmp, uNitidezAmp, uVineta, uGrano, uIntensidadFiltro;
uniform vec2 uCentro;
uniform uint uSemilla;
uniform vec4 uTonoA, uTonoB, uSatA, uSatB, uLumA, uLumB, uMezclaA, uMezclaB;

float transicion(float desde, float hasta, float x) {
  float t = clamp((x - desde) / (hasta - desde), 0.0, 1.0);
  return t * t * (3.0 - 2.0 * t);
}

vec3 rgbAHsl(vec3 c) {
  vec3 n = c / 255.0;
  float mx = max(n.r, max(n.g, n.b)), mn = min(n.r, min(n.g, n.b));
  float h = 0.0, s = 0.0, l = (mx + mn) / 2.0, d = mx - mn;
  if (d != 0.0) {
    s = l > 0.5 ? d / (2.0 - mx - mn) : d / (mx + mn);
    if (mx == n.r) h = (n.g - n.b) / d + (n.g < n.b ? 6.0 : 0.0);
    else if (mx == n.g) h = (n.b - n.r) / d + 2.0;
    else h = (n.r - n.g) / d + 4.0;
    h *= 60.0;
  }
  return vec3(h, s, l);
}

float canal(float p, float q, float t) {
  if (t < 0.0) t += 1.0;
  if (t > 1.0) t -= 1.0;
  if (t < 1.0 / 6.0) return p + (q - p) * 6.0 * t;
  if (t < 0.5) return q;
  if (t < 2.0 / 3.0) return p + (q - p) * (2.0 / 3.0 - t) * 6.0;
  return p;
}

vec3 hslARgb(float h, float s, float l) {
  if (s == 0.0) return vec3(l * 255.0);
  float q = l < 0.5 ? l * (1.0 + s) : l + s - l * s;
  float p = 2.0 * l - q;
  float hn = h / 360.0;
  return vec3(canal(p, q, hn + 1.0 / 3.0), canal(p, q, hn), canal(p, q, hn - 1.0 / 3.0)) * 255.0;
}

// calcularPesos de ajuste-por-color.js: los 8 rangos en dos vec4 (verdes es
// el cuarto, a.w).
void calcularPesos(vec3 hsl, bool conOliva, out vec4 a, out vec4 b) {
  float comoVerde = 0.0;
  if (conOliva && hsl.z < 0.4 && hsl.x > 35.0 && hsl.x < 105.0) {
    float oscuridad = 1.0 - transicion(0.12, 0.4, hsl.z);
    float cercania = 1.0 - transicion(20.0, 35.0, abs(hsl.x - 70.0));
    comoVerde = oscuridad * cercania;
  }
  float piso = floor(hsl.x);
  int h0 = int(piso) % 360, h1 = (h0 + 1) % 360;
  float fraccion = hsl.x - piso;
  vec4 a0 = texelFetch(uPesos, ivec2(0, h0), 0), b0 = texelFetch(uPesos, ivec2(1, h0), 0);
  vec4 a1 = texelFetch(uPesos, ivec2(0, h1), 0), b1 = texelFetch(uPesos, ivec2(1, h1), 0);
  a = (a0 + (a1 - a0) * fraccion) * (1.0 - comoVerde);
  b = (b0 + (b1 - b0) * fraccion) * (1.0 - comoVerde);
  a.w += comoVerde;
}

// pesosDeDecision: devuelve el croma con el que se decide.
float pesosDeDecision(vec3 c, vec3 hsl, vec3 ref, bool conOliva, out vec4 a, out vec4 b) {
  float cromaPixel = (max(c.r, max(c.g, c.b)) - min(c.r, min(c.g, c.b))) / 255.0;
  calcularPesos(hsl, conOliva, a, b);
  if (all(equal(c, ref))) return cromaPixel;
  float diferencia = abs((c.r - c.g) - (ref.r - ref.g)) + abs((c.g - c.b) - (ref.g - ref.b));
  float confianza = 1.0 - transicion(6.0, 20.0, diferencia);
  if (confianza == 0.0) return cromaPixel;
  vec4 ra, rb;
  calcularPesos(rgbAHsl(ref), conOliva, ra, rb);
  a += (ra - a) * confianza;
  b += (rb - b) * confianza;
  float cromaRef = (max(ref.r, max(ref.g, ref.b)) - min(ref.r, min(ref.g, ref.b))) / 255.0;
  return cromaPixel + (cromaRef - cromaPixel) * confianza;
}

// AjustePorColor.aplicar
vec3 aplicarPorColor(vec3 c, vec3 ref) {
  vec3 hsl = rgbAHsl(c);
  vec4 a, b;
  float croma = pesosDeDecision(c, hsl, ref, true, a, b);
  if (croma <= 0.02) return c;
  float dTono = dot(a, uTonoA) + dot(b, uTonoB);
  float dSat = dot(a, uSatA) + dot(b, uSatB);
  float dLuz = dot(a, uLumA) + dot(b, uLumB);
  if (dTono == 0.0 && dSat == 0.0 && dLuz == 0.0) return c;
  float fuerza = transicion(0.02, 0.14, croma);
  float h = hsl.x + dTono * 0.3 * fuerza;
  if (h < 0.0) h += 360.0; else if (h >= 360.0) h -= 360.0;
  float s = clamp(hsl.y * (1.0 + (dSat / 100.0) * fuerza), 0.0, 1.0);
  float k = (dLuz / 100.0) * 0.5 * fuerza;
  float l = k > 0.0
    ? hsl.z + (1.0 - hsl.z) * k * transicion(0.0, 0.25, hsl.z)
    : hsl.z * (1.0 + k * transicion(0.0, 0.25, 1.0 - hsl.z));
  return hslARgb(h, s, l);
}

float hombro(float x) { return x <= 0.8 ? x : 0.8 + 0.2 * (1.0 - exp(-(x - 0.8) / 0.2)); }

// AjustePorColor.mezclarGris
float mezclarGris(vec3 c, float gris, vec3 ref) {
  vec4 a, b;
  float croma = pesosDeDecision(c, rgbAHsl(c), ref, false, a, b);
  if (croma <= 0.02) return gris;
  float d = dot(a, uMezclaA) + dot(b, uMezclaB);
  if (d == 0.0) return gris;
  float pasos = (d / 100.0) * 1.5 * transicion(0.02, 0.14, croma);
  float antes = gris / 255.0;
  float despues = antes * exp2(pasos);
  if (pasos <= 0.0) return despues * 255.0;
  return min(255.0, (antes + hombro(despues) - hombro(antes)) * 255.0);
}

// Ruido del grano (en JS es Math.random()): un hash por píxel y por llamada.
uint hash(uvec3 v) {
  v = v * 1664525u + 1013904223u;
  v.x += v.y * v.z; v.y += v.z * v.x; v.z += v.x * v.y;
  v ^= v >> 16u;
  v.x += v.y * v.z; v.y += v.z * v.x; v.z += v.x * v.y;
  return v.x;
}

void main() {
  ivec2 p = pixel();
  vec4 entrada = pixelEntrada(p);
  vec3 c = entrada.rgb;
  vec3 o = c;
  vec3 ref = (uPorColor || uMezclaGris) ? leer(uRef, p).rgb : vec3(0.0);

  c *= uExposicion;
  if (uBalance) c *= uK;
  c = c255(c);
  c = c255(uContraste * (c - 128.0) + 128.0);

  if (uLuces) {
    float lum = dot(c, LUMA);
    float mascaraSombra = clamp(1.0 - lum / 128.0, 0.0, 1.0);
    float mascaraReflejo = clamp((lum - 128.0) / 127.0, 0.0, 1.0);
    float extremoBlanco = max(0.0, (lum - 170.0) / 85.0), extremoNegro = max(0.0, (85.0 - lum) / 85.0);
    float delta = uSombras * mascaraSombra + uReflejos * mascaraReflejo
      + uBlancos * extremoBlanco * extremoBlanco + uNegros * extremoNegro * extremoNegro;
    c = c255(c + delta);
  }

  if (uTono) c = c255(uMTono * c);

  if (uSaturar) {
    float lum = dot(c, LUMA);
    float maxC = max(c.r, max(c.g, c.b)), minC = min(c.r, min(c.g, c.b));
    float saturacionActual = maxC > 0.0 ? (maxC - minC) / maxC : 0.0;
    float factorTotal = uSat * (1.0 + uIntensidad * (1.0 - saturacionActual) * 1.5);
    c = c255(lum + (c - lum) * factorTotal);
  }

  if (uPorColor) c = aplicarPorColor(c, ref);

  c = c255(c + uDelta);

  if (uBn) {
    float gris = dot(c, uW) / uSumaCanales;
    if (uMezclaGris) gris = mezclarGris(c, gris, ref);
    c = vec3(gris);
  } else if (uMezclaGris) {
    float lum = dot(c, LUMA);
    if (lum > 0.0) c = min(c * (mezclarGris(c, lum, ref) / lum), 255.0);
  }

  if (uSepia != 0.0) {
    vec3 t = vec3(dot(c, vec3(0.393, 0.769, 0.189)), dot(c, vec3(0.349, 0.686, 0.168)), dot(c, vec3(0.272, 0.534, 0.131)));
    c = c255(c + (t - c) * uSepia);
  }

  if (uInvertir != 0.0) c += (255.0 - c - c) * uInvertir;

  if (uAmpDispersion > 0.0) {
    vec3 b = leer(uDispersion, p).rgb;
    if (uMatizDispersion) b = c255(uMDispersion * b);
    vec3 pantalla = 255.0 - (255.0 - c) * (255.0 - b) / 255.0;
    c = c255(c + (pantalla - c) * uAmpDispersion);
  }

  if (uClaridadAmp != 0.0) c += (dot(o, LUMA) - dot(leer(uClaridad, p).rgb, LUMA)) * uClaridadAmp;
  if (uNitidezAmp != 0.0) c += (dot(o, LUMA) - dot(leer(uNitidez, p).rgb, LUMA)) * uNitidezAmp;

  if (uVineta != 0.0) {
    vec2 centro = vec2(uCentro.x == 0.0 ? 1.0 : uCentro.x, uCentro.y == 0.0 ? 1.0 : uCentro.y);
    vec2 d = (vec2(p) - uCentro) / centro;
    float t = (length(d) / sqrt(2.0) - 0.35) / 0.65;
    if (t > 0.0) {
      t = min(t, 1.0);
      float m = t * t * (3.0 - 2.0 * t) * uVineta * 0.85;
      if (m < 0.0) c *= 1.0 + m;
      else c += (255.0 - c) * m;
    }
  }

  if (uGrano != 0.0) {
    float azar = float(hash(uvec3(uvec2(p), uSemilla)) >> 8u) / 16777216.0;
    c += (azar - 0.5) * uGrano;
  }

  if (uIntensidadFiltro != 1.0) c = o + (c - o) * uIntensidadFiltro;

  salida = vec4(u8(c), entrada.a) / 255.0;
}`;

  // Zonas brillantes de la dispersión de luz (bloom), sobre la imagen ya con
  // la aberración: como en el bucle de JS.
  const BRILLANTE = COMUN + `
void main() {
  vec4 c = pixelEntrada(pixel());
  float factor = clamp((dot(c.rgb, LUMA) - 170.0) / 60.0, 0.0, 1.0);
  salida = vec4(u8(c.rgb * factor) / 255.0, 1.0);
}`;

  // "Intensidad del filtro": mezcla con la imagen sin ajustes, alfa incluido.
  const MEZCLA = COMUN + `
uniform sampler2D uResultado;
uniform float uFactor;
void main() {
  ivec2 p = pixel();
  vec4 o = leer(uFuente, p), r = leer(uResultado, p);
  salida = roundEven(clamp(o + (r - o) * uFactor, 0.0, 255.0)) / 255.0;
}`;

  // Unidad de textura de cada sampler.
  const UNIDADES = { uFuente: 0, uRef: 1, uClaridad: 2, uNitidez: 3, uDispersion: 4, uPesos: 5, uResultado: 1 };

  function compilar(tipo, fuente) {
    const s = gl.createShader(tipo);
    gl.shaderSource(s, fuente);
    gl.compileShader(s);
    if (!gl.getShaderParameter(s, gl.COMPILE_STATUS)) throw new Error(gl.getShaderInfoLog(s));
    return s;
  }

  function programa(fuenteFragmentos) {
    const p = gl.createProgram();
    gl.attachShader(p, compilar(gl.VERTEX_SHADER, VERTICES));
    gl.attachShader(p, compilar(gl.FRAGMENT_SHADER, fuenteFragmentos));
    gl.linkProgram(p);
    if (!gl.getProgramParameter(p, gl.LINK_STATUS)) throw new Error(gl.getProgramInfoLog(p));
    const ubicaciones = {};
    const u = (nombre) => {
      if (!(nombre in ubicaciones)) ubicaciones[nombre] = gl.getUniformLocation(p, nombre);
      return ubicaciones[nombre];
    };
    gl.useProgram(p);
    for (const [nombre, unidad] of Object.entries(UNIDADES)) if (u(nombre)) gl.uniform1i(u(nombre), unidad);
    return { p, u };
  }

  // Tabla de pesos por tono (360 × 8) como textura: fila = grado, dos texels
  // de 4 rangos cada uno.
  function texturaPesos() {
    const datos = new Float32Array(360 * 8);
    AjustePorColor.PESOS.forEach((pesos, h) => { for (let k = 0; k < 8; k++) datos[h * 8 + k] = pesos[k]; });
    const t = gl.createTexture();
    gl.activeTexture(gl.TEXTURE0 + UNIDADES.uPesos);
    gl.bindTexture(gl.TEXTURE_2D, t);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.NEAREST);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, gl.NEAREST);
    gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGBA32F, 2, 360, 0, gl.RGBA, gl.FLOAT, datos);
  }

  function olvidarContexto() {
    gl = null; lienzo = null; programas = null;
    texturas.length = 0;
  }

  function iniciar() {
    if (gl) return true;
    if (roto) return false;
    try {
      if (AjustePorColor.RANGOS.length !== 8 || AjustePorColor.INDICE_VERDES !== 3) throw new Error('el shader espera 8 rangos de color con verdes en el cuarto');
      lienzo = document.createElement('canvas');
      gl = lienzo.getContext('webgl2', {
        alpha: true, premultipliedAlpha: false, antialias: false, depth: false, stencil: false, preserveDrawingBuffer: false,
      });
      if (!gl) throw new Error('sin WebGL2');
      // Si la placa se reinicia, la próxima llamada arma un contexto nuevo;
      // después de 3 veces, se queda con el bucle en JavaScript.
      lienzo.addEventListener('webglcontextlost', (e) => {
        e.preventDefault();
        perdidas += 1;
        if (perdidas >= 3) roto = true;
        olvidarContexto();
      });
      ladoMaximo = Math.min(gl.getParameter(gl.MAX_TEXTURE_SIZE), gl.getParameter(gl.MAX_RENDERBUFFER_SIZE),
        ...gl.getParameter(gl.MAX_VIEWPORT_DIMS));
      programas = { principal: programa(PRINCIPAL), brillante: programa(BRILLANTE), mezcla: programa(MEZCLA) };
      texturaPesos();
      gl.pixelStorei(gl.UNPACK_FLIP_Y_WEBGL, false);
      gl.pixelStorei(gl.UNPACK_PREMULTIPLY_ALPHA_WEBGL, false);
      gl.pixelStorei(gl.UNPACK_COLORSPACE_CONVERSION_WEBGL, gl.NONE);
      return true;
    } catch (error) {
      console.warn('FiltrosGPU: se usa el motor en JavaScript.', error);
      roto = true;
      olvidarContexto();
      return false;
    }
  }

  function disponible(ancho, alto) {
    return ancho > 0 && alto > 0 && iniciar() && !gl.isContextLost() && ancho <= ladoMaximo && alto <= ladoMaximo;
  }

  // El lienzo de WebGL solo crece (cambiarle el tamaño en cada llamada, entre
  // miniaturas y vista previa, es caro); se dibuja en su esquina de abajo a
  // la izquierda. Si el navegador no le da el tamaño pedido, no se usa.
  function prepararLienzo(ancho, alto) {
    if (lienzo.width < ancho || lienzo.height < alto) {
      lienzo.width = Math.max(lienzo.width, ancho);
      lienzo.height = Math.max(lienzo.height, alto);
    }
    if (gl.drawingBufferWidth < ancho || gl.drawingBufferHeight < alto) {
      ladoMaximo = Math.min(ladoMaximo, Math.max(ancho, alto) - 1);
      return false;
    }
    gl.viewport(0, 0, ancho, alto);
    return true;
  }

  function subir(unidad, fuente) {
    let t = texturas[unidad];
    gl.activeTexture(gl.TEXTURE0 + unidad);
    if (!t) {
      t = texturas[unidad] = { textura: gl.createTexture(), ancho: 0, alto: 0 };
      gl.bindTexture(gl.TEXTURE_2D, t.textura);
      gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.NEAREST);
      gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, gl.NEAREST);
      gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_S, gl.CLAMP_TO_EDGE);
      gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_T, gl.CLAMP_TO_EDGE);
    } else {
      gl.bindTexture(gl.TEXTURE_2D, t.textura);
    }
    if (t.ancho === fuente.width && t.alto === fuente.height) {
      gl.texSubImage2D(gl.TEXTURE_2D, 0, 0, 0, gl.RGBA, gl.UNSIGNED_BYTE, fuente);
    } else {
      gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGBA, gl.RGBA, gl.UNSIGNED_BYTE, fuente);
      t.ancho = fuente.width; t.alto = fuente.height;
    }
  }

  function aberracion(prog, p) {
    const a = p.aberracion;
    gl.uniform1i(prog.u('uAberracion'), a ? 1 : 0);
    gl.uniform1i(prog.u('uAberracionMatiz'), a && a.matriz ? 1 : 0);
    if (a) gl.uniform2i(prog.u('uDesp'), a.dx, a.dy);
    if (a && a.matriz) gl.uniformMatrix3fv(prog.u('uMAberracion'), true, a.matriz);
  }

  // Dibuja y deja el resultado en "destino" (reemplaza todo, alfa incluido).
  function dibujarEn(destino, ancho, alto) {
    gl.drawArrays(gl.TRIANGLES, 0, 3);
    if (gl.isContextLost()) return false;
    const ctx = destino.getContext('2d');
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.globalCompositeOperation = 'copy';
    ctx.globalAlpha = 1;
    ctx.filter = 'none';
    ctx.drawImage(lienzo, 0, lienzo.height - alto, ancho, alto, 0, 0, ancho, alto);
    ctx.restore();
    return true;
  }

  // Corre una pasada; ante cualquier error devuelve false y el motor sigue
  // en JavaScript.
  function pasada(ancho, alto, armar, destino) {
    if (!disponible(ancho, alto)) return false;
    try {
      if (!prepararLienzo(ancho, alto)) return false;
      gl.uniform2i(armar(), ancho, alto);
      return dibujarEn(destino, ancho, alto);
    } catch (error) {
      console.warn('FiltrosGPU: falló una pasada, se usa el motor en JavaScript.', error);
      roto = true;
      olvidarContexto();
      return false;
    }
  }

  const vec4De = (lista, desde, clave) => [0, 1, 2, 3].map((k) => (clave ? lista[desde + k][clave] : lista[desde + k]) || 0);

  // procesarCanvas: "p" es prepararAjustes(); "refs" los canvases de
  // referencia que pida (color, claridad, nitidez, dispersion).
  function procesar(canvas, p, refs) {
    return pasada(p.ancho, p.alto, () => {
      const prog = programas.principal;
      const u = prog.u;
      gl.useProgram(prog.p);
      subir(UNIDADES.uFuente, canvas);
      if (refs.color) subir(UNIDADES.uRef, refs.color);
      if (refs.claridad) subir(UNIDADES.uClaridad, refs.claridad);
      if (refs.nitidez) subir(UNIDADES.uNitidez, refs.nitidez);
      if (refs.dispersion) subir(UNIDADES.uDispersion, refs.dispersion);
      aberracion(prog, p);
      gl.uniform1f(u('uExposicion'), p.factorExposicion);
      gl.uniform1i(u('uBalance'), p.necesitaBalance ? 1 : 0);
      gl.uniform3f(u('uK'), p.kR, p.kG, p.kB);
      gl.uniform1f(u('uContraste'), p.factorContraste);
      gl.uniform1i(u('uLuces'), p.amplitudSombras !== 0 || p.amplitudReflejos !== 0 || p.ampBlancos !== 0 || p.ampNegros !== 0 ? 1 : 0);
      gl.uniform1f(u('uSombras'), p.amplitudSombras);
      gl.uniform1f(u('uReflejos'), p.amplitudReflejos);
      gl.uniform1f(u('uBlancos'), p.ampBlancos);
      gl.uniform1f(u('uNegros'), p.ampNegros);
      gl.uniform1i(u('uTono'), p.mTono ? 1 : 0);
      if (p.mTono) gl.uniformMatrix3fv(u('uMTono'), true, p.mTono);
      gl.uniform1i(u('uSaturar'), p.factorSaturacion !== 1 || p.nIntensidad !== 0 ? 1 : 0);
      gl.uniform1f(u('uSat'), p.factorSaturacion);
      gl.uniform1f(u('uIntensidad'), p.nIntensidad);
      const porColor = Boolean(p.ajustePorColor && refs.color);
      gl.uniform1i(u('uPorColor'), porColor ? 1 : 0);
      if (porColor) {
        const r = p.ajustePorColor;
        gl.uniform4fv(u('uTonoA'), vec4De(r, 0, 'tono')); gl.uniform4fv(u('uTonoB'), vec4De(r, 4, 'tono'));
        gl.uniform4fv(u('uSatA'), vec4De(r, 0, 'saturacion')); gl.uniform4fv(u('uSatB'), vec4De(r, 4, 'saturacion'));
        gl.uniform4fv(u('uLumA'), vec4De(r, 0, 'luminancia')); gl.uniform4fv(u('uLumB'), vec4De(r, 4, 'luminancia'));
      }
      gl.uniform3f(u('uDelta'), p.deltaColorRojo, p.deltaColorVerde, p.deltaColorAzul);
      gl.uniform1i(u('uBn'), p.bn ? 1 : 0);
      gl.uniform3f(u('uW'), p.wr, p.wg, p.wb);
      gl.uniform1f(u('uSumaCanales'), p.sumaCanales);
      const mezclaGris = Boolean(p.mezclaGris && refs.color);
      gl.uniform1i(u('uMezclaGris'), mezclaGris ? 1 : 0);
      if (mezclaGris) {
        gl.uniform4fv(u('uMezclaA'), vec4De(p.mezclaGris, 0));
        gl.uniform4fv(u('uMezclaB'), vec4De(p.mezclaGris, 4));
      }
      gl.uniform1f(u('uSepia'), p.ampSepia);
      gl.uniform1f(u('uInvertir'), p.ampInvertir);
      gl.uniform1f(u('uAmpDispersion'), refs.dispersion ? p.ampDispersion : 0);
      gl.uniform1i(u('uMatizDispersion'), p.mMatizDispersion ? 1 : 0);
      if (p.mMatizDispersion) gl.uniformMatrix3fv(u('uMDispersion'), true, p.mMatizDispersion);
      gl.uniform1f(u('uClaridadAmp'), refs.claridad ? p.ampClaridad : 0);
      gl.uniform1f(u('uNitidezAmp'), refs.nitidez ? p.ampNitidez : 0);
      gl.uniform1f(u('uVineta'), p.ampVineta);
      gl.uniform2f(u('uCentro'), p.centroX, p.centroY);
      gl.uniform1f(u('uGrano'), p.ampGrano);
      gl.uniform1ui(u('uSemilla'), (Math.random() * 0xffffffff) >>> 0);
      gl.uniform1f(u('uIntensidadFiltro'), p.factorIntensidadFiltro);
      return u('uTam');
    }, canvas);
  }

  // Las zonas brillantes de la dispersión, en un canvas nuevo (o null).
  function brillante(canvas, p) {
    const destino = document.createElement('canvas');
    destino.width = p.ancho; destino.height = p.alto;
    const ok = pasada(p.ancho, p.alto, () => {
      const prog = programas.brillante;
      gl.useProgram(prog.p);
      subir(UNIDADES.uFuente, canvas);
      aberracion(prog, p);
      return prog.u('uTam');
    }, destino);
    return ok ? destino : null;
  }

  // canvas = base + (canvas − base) × factor, en los cuatro canales.
  function mezclar(base, canvas, factor) {
    return pasada(canvas.width, canvas.height, () => {
      const prog = programas.mezcla;
      gl.useProgram(prog.p);
      subir(UNIDADES.uFuente, base);
      subir(UNIDADES.uResultado, canvas);
      gl.uniform1i(prog.u('uAberracion'), 0);
      gl.uniform1f(prog.u('uFactor'), factor);
      return prog.u('uTam');
    }, canvas);
  }

  // El lado más grande (en px) que acepta la placa; 0 sin WebGL2. Para que el
  // editor explique por qué no ofrece una resolución.
  function ladoMaximoGpu() {
    return iniciar() ? ladoMaximo : 0;
  }

  return { disponible, ladoMaximo: ladoMaximoGpu, procesar, brillante, mezclar };
})();

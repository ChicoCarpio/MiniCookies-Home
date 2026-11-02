/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// LUT de color en formato .cube (Adobe/Resolve): una tabla que dice a qué color se cambia cada color. Lee el archivo de texto y lo aplica a un
// canvas. Los editores de imagen lo usan como un paso más del pipeline (ProcesarFiltros.renderizarFiltro, ajustes.lut).
//
//   LutCube.leer(texto)               → { titulo, tipo: '3d' | '1d', tamano, datos: Float32Array, min: [r,g,b], max: [r,g,b] }  (lanza Error con 'codigo')
//   LutCube.aplicar(canvas, lut, f)   → muta el canvas: original + (con LUT − original) × f, f de 0 a 2. No toca el canal alfa.
//
// Formato: líneas "TITLE", "DOMAIN_MIN/MAX r g b", "LUT_3D_SIZE n" (o "LUT_1D_SIZE n") y luego n³ (o n) filas "r g b" entre 0 y 1; en las de
// 3D el rojo es el que más rápido cambia. Se ignoran los comentarios (#) y las líneas desconocidas (por ejemplo LUT_3D_INPUT_RANGE no se
// usa; DOMAIN_MIN/MAX sí). Interpolación trilineal (3D) o lineal por canal (1D).
(function (global) {
  var TAMANO_3D_MAXIMO = 129;
  var TAMANO_1D_MAXIMO = 65536;

  function fallo(codigo, texto) {
    var error = new Error(texto || codigo);
    error.codigo = codigo;
    return error;
  }

  function tresNumeros(partes) {
    var v = [Number(partes[1]), Number(partes[2]), Number(partes[3])];
    return v.every(isFinite) && partes.length >= 4 ? v : null;
  }

  function leer(texto) {
    var lineas = String(texto).split(/\r\n|\r|\n/);
    var titulo = '', tipo = null, tamano = 0, min = [0, 0, 0], max = [1, 1, 1];
    var datos = null, n = 0, esperado = 0;
    for (var i = 0; i < lineas.length; i++) {
      var linea = lineas[i].replace(/#.*/, '').trim();
      if (!linea) continue;
      var partes = linea.split(/\s+/);
      var palabra = partes[0].toUpperCase();
      if (datos === null || n < esperado) {
        if (palabra === 'TITLE') { titulo = linea.replace(/^TITLE\s*/i, '').replace(/^"|"$/g, ''); continue; }
        if (palabra === 'DOMAIN_MIN' || palabra === 'DOMAIN_MAX') {
          var d = tresNumeros(partes);
          if (!d) throw fallo('dominio', 'DOMAIN inválido en la línea ' + (i + 1));
          if (palabra === 'DOMAIN_MIN') min = d; else max = d;
          continue;
        }
        if (palabra === 'LUT_3D_SIZE' || palabra === 'LUT_1D_SIZE') {
          if (tipo) throw fallo('formato', 'El archivo declara dos tablas');
          tipo = palabra === 'LUT_3D_SIZE' ? '3d' : '1d';
          tamano = Number(partes[1]);
          if (!Number.isInteger(tamano) || tamano < 2 || tamano > (tipo === '3d' ? TAMANO_3D_MAXIMO : TAMANO_1D_MAXIMO)) throw fallo('tamano', 'Tamaño de tabla no válido: ' + partes[1]);
          esperado = tipo === '3d' ? tamano * tamano * tamano : tamano;
          datos = new Float32Array(esperado * 3);
          continue;
        }
        if (/^[A-Z_]+[A-Z0-9_]*$/.test(palabra) && isNaN(Number(palabra))) continue;   // otra palabra clave que no usamos
      }
      var v = tresNumeros(['', partes[0], partes[1], partes[2]]);
      if (!v || partes.length < 3) throw fallo('fila', 'Fila no válida en la línea ' + (i + 1));
      if (!datos) throw fallo('formato', 'Hay colores antes de LUT_3D_SIZE o LUT_1D_SIZE');
      if (n >= esperado) throw fallo('filas', 'Sobran filas respecto del tamaño declarado');
      datos[n * 3] = v[0]; datos[n * 3 + 1] = v[1]; datos[n * 3 + 2] = v[2];
      n++;
    }
    if (!datos) throw fallo('formato', 'No es un archivo .cube (falta LUT_3D_SIZE o LUT_1D_SIZE)');
    if (n < esperado) throw fallo('filas', 'Faltan filas: hay ' + n + ' de ' + esperado);
    for (var k = 0; k < 3; k++) if (!(max[k] > min[k])) throw fallo('dominio', 'DOMAIN_MAX debe ser mayor que DOMAIN_MIN');
    return { titulo: titulo, tipo: tipo, tamano: tamano, datos: datos, min: min, max: max };
  }

  // Un solo canal: de 0..255 a la posición en la tabla (0..tamano−1), ya recortada.
  function posicion(valor255, min, max, tamano) {
    var t = (valor255 / 255 - min) / (max - min);
    t = t < 0 ? 0 : t > 1 ? 1 : t;
    return t * (tamano - 1);
  }

  function aplicar(canvas, lut, factor) {
    if (!lut || factor === 0) return canvas;
    if (factor === undefined) factor = 1;
    var ctx = canvas.getContext('2d');
    var imagen = ctx.getImageData(0, 0, canvas.width, canvas.height);
    var px = imagen.data, datos = lut.datos, N = lut.tamano, min = lut.min, max = lut.max;
    var NN = N * N, i, r, g, b, orR, orG, orB;
    if (lut.tipo === '3d') {
      for (i = 0; i < px.length; i += 4) {
        orR = px[i]; orG = px[i + 1]; orB = px[i + 2];
        var fr = posicion(orR, min[0], max[0], N), fg = posicion(orG, min[1], max[1], N), fb = posicion(orB, min[2], max[2], N);
        var r0 = Math.floor(fr), g0 = Math.floor(fg), b0 = Math.floor(fb);
        var r1 = r0 + 1 < N ? r0 + 1 : r0, g1 = g0 + 1 < N ? g0 + 1 : g0, b1 = b0 + 1 < N ? b0 + 1 : b0;
        var dr = fr - r0, dg = fg - g0, db = fb - b0;
        var i000 = (r0 + g0 * N + b0 * NN) * 3, i100 = (r1 + g0 * N + b0 * NN) * 3, i010 = (r0 + g1 * N + b0 * NN) * 3, i110 = (r1 + g1 * N + b0 * NN) * 3;
        var i001 = (r0 + g0 * N + b1 * NN) * 3, i101 = (r1 + g0 * N + b1 * NN) * 3, i011 = (r0 + g1 * N + b1 * NN) * 3, i111 = (r1 + g1 * N + b1 * NN) * 3;
        var c;
        for (c = 0; c < 3; c++) {
          var x00 = datos[i000 + c] + (datos[i100 + c] - datos[i000 + c]) * dr;
          var x10 = datos[i010 + c] + (datos[i110 + c] - datos[i010 + c]) * dr;
          var x01 = datos[i001 + c] + (datos[i101 + c] - datos[i001 + c]) * dr;
          var x11 = datos[i011 + c] + (datos[i111 + c] - datos[i011 + c]) * dr;
          var y0 = x00 + (x10 - x00) * dg, y1 = x01 + (x11 - x01) * dg;
          var salida = (y0 + (y1 - y0) * db) * 255;
          var original = c === 0 ? orR : c === 1 ? orG : orB;
          px[i + c] = original + (salida - original) * factor;   // Uint8ClampedArray recorta a 0..255 y redondea
        }
      }
    } else {
      for (i = 0; i < px.length; i += 4) {
        for (var k = 0; k < 3; k++) {
          var original1 = px[i + k];
          var f = posicion(original1, min[k], max[k], N);
          var a = Math.floor(f), a1 = a + 1 < N ? a + 1 : a;
          var s = (datos[a * 3 + k] + (datos[a1 * 3 + k] - datos[a * 3 + k]) * (f - a)) * 255;
          px[i + k] = original1 + (s - original1) * factor;
        }
      }
    }
    ctx.putImageData(imagen, 0, 0);
    return canvas;
  }

  global.LutCube = { leer: leer, aplicar: aplicar, TAMANO_3D_MAXIMO: TAMANO_3D_MAXIMO };
})(typeof window !== 'undefined' ? window : this);

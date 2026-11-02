/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Generador de contraseñas aleatorias, compartido por la herramienta
// "Generador de contraseñas" y por el administrador de contraseñas.
//
// Usa crypto.getRandomValues (nunca Math.random). Rechaza y vuelve a sortear
// las que tienen repetidos (aaa), secuencias (abc, 321) o fragmentos de
// contraseñas comunes.
//
// GeneradorContrasenas.generar(longitud, { mayusculas, minusculas, numeros, simbolos })
//   → texto. Si las opciones no sirven lanza un Error con `codigo`:
//   'sin-tipos' (no se eligió ningún tipo) o 'longitud' (con `minimo`: la
//   longitud mínima para lo elegido). Quien llama lo traduce.
// GeneradorContrasenas.fortaleza(texto) → 'debil' | 'media' | 'fuerte'
var GeneradorContrasenas = (function () {
  const DENYLIST = [
    'password', 'contrasena', 'contraseña', 'qwerty', 'asdfgh', 'zxcvbn', 'letmein',
    'admin', 'welcome', 'iloveyou', 'abc123', '1234567890', '0123456789',
  ];
  const MAX_INTENTOS = 200;

  // Entero aleatorio uniforme en [0, maximo). Con un simple "% maximo" los
  // primeros valores salen apenas más seguido que el resto (sesgo de módulo):
  // se descartan los números del tramo final que no alcanza para una vuelta
  // completa y se vuelve a sortear.
  function enteroAleatorio(maximo) {
    const limite = Math.floor(0x100000000 / maximo) * maximo;
    const valor = new Uint32Array(1);
    do { crypto.getRandomValues(valor); } while (valor[0] >= limite);
    return valor[0] % maximo;
  }

  function eleccionAleatoria(conjunto) {
    return conjunto[enteroAleatorio(conjunto.length)];
  }

  function barajar(arreglo) {
    for (let i = arreglo.length - 1; i > 0; i--) {
      const j = enteroAleatorio(i + 1);
      [arreglo[i], arreglo[j]] = [arreglo[j], arreglo[i]];
    }
    return arreglo;
  }

  function tieneRepetidos(contrasena, largoCorrida = 3) {
    for (let i = 0; i <= contrasena.length - largoCorrida; i++) {
      const trozo = contrasena.slice(i, i + largoCorrida);
      if (new Set(trozo).size === 1) return true;
    }
    return false;
  }

  function tieneSecuencia(contrasena, largoCorrida = 3) {
    for (let i = 0; i <= contrasena.length - largoCorrida; i++) {
      const codigos = [...contrasena.slice(i, i + largoCorrida)].map((c) => c.charCodeAt(0));
      const deltas = [];
      for (let j = 0; j < codigos.length - 1; j++) deltas.push(codigos[j + 1] - codigos[j]);
      if (deltas.every((d) => d === 1) || deltas.every((d) => d === -1)) return true;
    }
    return false;
  }

  function contienePatronProhibido(contrasena) {
    const minuscula = contrasena.toLowerCase();
    return DENYLIST.some((termino) => minuscula.includes(termino));
  }

  function esDebil(contrasena) {
    return tieneRepetidos(contrasena) || tieneSecuencia(contrasena) || contienePatronProhibido(contrasena);
  }

  function error(codigo, mensaje, extra) {
    return Object.assign(new Error(mensaje), { codigo }, extra);
  }

  function generar(longitud, opciones) {
    const o = opciones || {};
    const pools = [];
    if (o.mayusculas) pools.push('ABCDEFGHIJKLMNOPQRSTUVWXYZ');
    if (o.minusculas) pools.push('abcdefghijklmnopqrstuvwxyz');
    if (o.numeros) pools.push('0123456789');
    if (o.simbolos) pools.push('!@#$%^&*()-_=+[]{}?');
    if (pools.length === 0) throw error('sin-tipos', 'Elegí al menos un tipo de carácter.');
    if (!Number.isInteger(longitud)) throw error('longitud', 'El largo tiene que ser un número entero.', { minimo: pools.length });
    if (longitud < pools.length) {
      throw error('longitud', 'La longitud mínima para lo elegido es ' + pools.length + '.', { minimo: pools.length });
    }

    const todosLosCaracteres = pools.join('');
    for (let intento = 0; intento < MAX_INTENTOS; intento++) {
      let caracteres = pools.map((pool) => eleccionAleatoria(pool));
      while (caracteres.length < longitud) caracteres.push(eleccionAleatoria(todosLosCaracteres));
      caracteres = barajar(caracteres);
      const contrasena = caracteres.join('');
      if (!esDebil(contrasena)) return contrasena;
    }
    throw error('intentos', 'No se pudo generar una contraseña segura tras varios intentos, probá otras opciones.');
  }

  // Etiqueta simple para la interfaz (no es una medida de seguridad real):
  // largo y variedad de tipos de carácter.
  function fortaleza(contrasena) {
    const tipos = [/[A-ZÁÉÍÓÚÑ]/, /[a-záéíóúñ]/, /\d/, /[^A-Za-z0-9áéíóúñÁÉÍÓÚÑ]/]
      .filter((r) => r.test(contrasena)).length;
    if (contienePatronProhibido(contrasena)) return 'debil'; // "Password1234567!" no es fuerte aunque tenga de todo
    if (contrasena.length >= 20) return 'fuerte'; // una frase larga vale aunque sea de un solo tipo
    if (contrasena.length >= 16 && tipos >= 3) return 'fuerte';
    if (contrasena.length >= 10 && tipos >= 2) return 'media';
    return 'debil';
  }

  return { generar, fortaleza, esDebil };
})();

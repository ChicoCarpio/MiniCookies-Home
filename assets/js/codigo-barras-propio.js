/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Generador de códigos de barras propio (reemplaza a JsBarcode). Codifica
// CODE128 (modos A, B y C elegidos solos), EAN-13, UPC-A y CODE39, y los dibuja
// como SVG: las barras son rectángulos, el texto va debajo.
//
// Es un port fiel de lo que JsBarcode 3.12.3 hacía con esos cuatro formatos y
// con las opciones que usa el sitio, incluidas sus rarezas (por ejemplo, qué
// entradas rechaza CODE128 y cómo se centra el texto), para que el dibujo sea
// idéntico al de antes. Lo que no estaba en uso (EAN-8, ITF, MSI, codabar, la
// salida en canvas, etc.) no se porta.
//
// API:
//   CodigoBarras.dibujarSvg(svg, texto, { formato, ancho, alto, colorLinea, … });
//   CodigoBarras.codificar(formato, texto, opciones)   // partes: [{ data, text, options }]
//   CodigoBarras.FORMATOS                              // ['CODE128', 'EAN13', 'UPC', 'CODE39']
//
// Opciones (con su valor por omisión): formato 'CODE128', ancho 2 (px por
// módulo), alto 100, colorLinea '#000000', fondo '#ffffff', margen 10,
// mostrarTexto true, tamanoTexto 20, margenTexto 2, fuente 'monospace'.
//
// Si la entrada no es válida para el formato (largo o dígito de control
// incorrectos, caracteres que no existen en ese código) tira un Error con
// name 'EntradaNoValida' y las propiedades formato y entrada. No toca el SVG
// en ese caso. El texto del mensaje queda para depurar; lo que se le muestra a
// quien usa la herramienta lo arma la página con t().
(function () {
  'use strict';

  const NS_SVG = 'http://www.w3.org/2000/svg';
  const FORMATOS = ['CODE128', 'EAN13', 'UPC', 'CODE39'];

  function entradaNoValida(formato, entrada) {
    const error = new Error('"' + entrada + '" is not a valid input for ' + formato);
    error.name = 'EntradaNoValida';
    error.formato = formato;
    error.entrada = entrada;
    return error;
  }

  // ---------- CODE128 ----------

  // Patrón de barras (11 módulos, 13 el de parada) de cada símbolo 0..106.
  const BARRAS_128 = [
    '11011001100', '11001101100', '11001100110', '10010011000', '10010001100', '10001001100',
    '10011001000', '10011000100', '10001100100', '11001001000', '11001000100', '11000100100',
    '10110011100', '10011011100', '10011001110', '10111001100', '10011101100', '10011100110',
    '11001110010', '11001011100', '11001001110', '11011100100', '11001110100', '11101101110',
    '11101001100', '11100101100', '11100100110', '11101100100', '11100110100', '11100110010',
    '11011011000', '11011000110', '11000110110', '10100011000', '10001011000', '10001000110',
    '10110001000', '10001101000', '10001100010', '11010001000', '11000101000', '11000100010',
    '10110111000', '10110001110', '10001101110', '10111011000', '10111000110', '10001110110',
    '11101110110', '11010001110', '11000101110', '11011101000', '11011100010', '11011101110',
    '11101011000', '11101000110', '11100010110', '11101101000', '11101100010', '11100011010',
    '11101111010', '11001000010', '11110001010', '10100110000', '10100001100', '10010110000',
    '10010000110', '10000101100', '10000100110', '10110010000', '10110000100', '10011010000',
    '10011000010', '10000110100', '10000110010', '11000010010', '11001010000', '11110111010',
    '11000010100', '10001111010', '10100111100', '10010111100', '10010011110', '10111100100',
    '10011110100', '10011110010', '11110100100', '11110010100', '11110010010', '11011011110',
    '11011110110', '11110110110', '10101111000', '10100011110', '10001011110', '10111101000',
    '10111100010', '11110101000', '11110100010', '10111011110', '10111101110', '11101011110',
    '11110101110', '11010000100', '11010010000', '11010011100', '1100011101011',
  ];
  const INICIO_A = 103, INICIO_B = 104, INICIO_C = 105, MODULO_128 = 103, PARADA = 106, SHIFT = 98, FNC1 = 207;
  const MODO_A = 0, MODO_B = 1, MODO_C = 2;
  // Símbolo de inicio -> modo; y símbolo de cambio (101, 100, 99) -> modo.
  const MODO_DE_INICIO = { [INICIO_A]: MODO_A, [INICIO_B]: MODO_B, [INICIO_C]: MODO_C };
  const CAMBIO_A_MODO = { 101: MODO_A, 100: MODO_B, 99: MODO_C };
  // Caracteres de cada modo: los de control o de datos, más las funciones
  // (\xC8-\xCF). Las funciones \xCC-\xD2 son marcas internas de esta lógica.
  const CARS_A = '[\\x00-\\x5F\\xC8-\\xCF]';
  const CARS_B = '[\\x20-\\x7F\\xC8-\\xCF]';
  const CARS_C = '(\\xCF*[0-9]{2}\\xCF*)';
  const CAR_INICIO_A = String.fromCharCode(208);
  const CAR_INICIO_B = String.fromCharCode(209);
  const CAR_INICIO_C = String.fromCharCode(210);
  const VALIDOS_128 = /^[\x00-\x7F\xC8-\xD3]+$/;

  function largoModoA(s) { return s.match(new RegExp('^' + CARS_A + '*'))[0].length; }
  function largoModoB(s) { return s.match(new RegExp('^' + CARS_B + '*'))[0].length; }
  function tramoModoC(s) { return s.match(new RegExp('^' + CARS_C + '*'))[0]; }

  // Recorre el texto decidiendo dónde cambiar de modo; devuelve el mismo texto
  // con los caracteres de cambio (\xCC = a C, \xCD = a B, \xCE = a A) metidos.
  // Si encuentra un carácter que no está en ningún modo, tira (null) y quien
  // llama lo convierte en «entrada no válida».
  function elegirModos(texto, enA) {
    let resultado = '';
    for (;;) {
      const cars = enA ? CARS_A : CARS_B;
      const hastaC = texto.match(new RegExp('^(' + cars + '+?)(([0-9]{2}){2,})([^0-9]|$)'));
      if (hastaC) {
        const desdeC = elegirDesdeC(texto.substring(hastaC[1].length));
        return desdeC === null ? null : resultado + hastaC[1] + String.fromCharCode(204) + desdeC;
      }
      const tramo = texto.match(new RegExp('^' + cars + '+'));
      if (!tramo) return null;
      if (tramo[0].length === texto.length) return resultado + texto;
      resultado += tramo[0] + String.fromCharCode(enA ? 205 : 206);
      texto = texto.substring(tramo[0].length);
      enA = !enA;
    }
  }
  function elegirDesdeC(texto) {
    const tramoC = tramoModoC(texto);
    if (tramoC.length === texto.length) return texto;
    texto = texto.substring(tramoC.length);
    const enA = largoModoA(texto) >= largoModoB(texto);
    const resto = elegirModos(texto, enA);
    if (resto === null) return null;
    return tramoC + String.fromCharCode(enA ? 206 : 205) + resto;
  }
  function elegirInicio(texto) {
    let elegido;
    if (tramoModoC(texto).length >= 2) {
      const desdeC = elegirDesdeC(texto);
      if (desdeC === null) return null;
      elegido = CAR_INICIO_C + desdeC;
    } else {
      const enA = largoModoA(texto) > largoModoB(texto);
      const resto = elegirModos(texto, enA);
      if (resto === null) return null;
      elegido = (enA ? CAR_INICIO_A : CAR_INICIO_B) + resto;
    }
    // Un solo carácter entre dos cambios de modo se manda con SHIFT (\xCB).
    return elegido.replace(/[\xCD\xCE]([^])[\xCD\xCE]/, function (todo, car) { return String.fromCharCode(203) + car; });
  }

  function barra128(indice) {
    return BARRAS_128[indice] ? BARRAS_128[indice] : '';
  }

  // Valor del símbolo que corresponde al siguiente byte (o par de dígitos en C).
  function indiceDeByte(bytes, modo) {
    if (modo === MODO_A) {
      const b = bytes.shift();
      return b < 32 ? b + 64 : b - 32;
    }
    if (modo === MODO_B) return bytes.shift() - 32;
    return 10 * (bytes.shift() - 48) + bytes.shift() - 48;
  }

  function codificar128(texto) {
    const convertido = VALIDOS_128.test(texto) ? elegirInicio(texto) : texto;
    if (convertido === null) throw entradaNoValida('CODE128', texto);
    const datos = convertido.substring(1);
    if (!VALIDOS_128.test(datos)) throw entradaNoValida('CODE128', texto);
    const bytes = convertido.split('').map(function (c) { return c.charCodeAt(0); });
    const inicio = bytes.shift() - 105;
    let modo = MODO_DE_INICIO[inicio];
    // Texto con un primer carácter que no es de inicio: no se puede codificar.
    if (modo === undefined) throw entradaNoValida('CODE128', texto);

    let barras = barra128(inicio);
    let suma = 0;
    let posicion = 1;
    while (bytes.length) {
      let indice;
      if (bytes[0] >= 200) {
        indice = bytes.shift() - 105;
        const nuevo = CAMBIO_A_MODO[indice];
        if (nuevo !== undefined) {
          modo = nuevo;
        } else if ((modo === MODO_A || modo === MODO_B) && indice === SHIFT) {
          // El carácter que sigue se toma del otro modo, solo ese.
          bytes[0] = modo === MODO_A
            ? (bytes[0] > 95 ? bytes[0] - 96 : bytes[0])
            : (bytes[0] < 32 ? bytes[0] + 96 : bytes[0]);
        }
      } else {
        indice = indiceDeByte(bytes, modo);
      }
      barras += barra128(indice);
      suma += indice * posicion;
      posicion++;
    }
    barras += barra128((suma + inicio) % MODULO_128) + barra128(PARADA);
    // Se muestra el texto sin los caracteres de control ni los de cambio de modo.
    return [{ data: barras, text: datos.replace(/[^\x20-\x7E]/g, '') }];
  }

  // ---------- EAN-13 y UPC-A ----------

  const L = ['0001101', '0011001', '0010011', '0111101', '0100011', '0110001', '0101111', '0111011', '0110111', '0001011'];
  const G = ['0100111', '0110011', '0011011', '0100001', '0011101', '0111001', '0000101', '0010001', '0001001', '0010111'];
  const R = ['1110010', '1100110', '1101100', '1000010', '1011100', '1001110', '1010000', '1000100', '1001000', '1110100'];
  const TABLAS_EAN = { L: L, G: G, R: R };
  const GUARDA_LATERAL = '101';
  const GUARDA_CENTRAL = '01010';
  // Qué tabla (L o G) usa cada uno de los 6 dígitos de la izquierda, según el primer dígito.
  const ESTRUCTURA_EAN13 = ['LLLLLL', 'LLGLGG', 'LLGGLG', 'LLGGGL', 'LGLLGG', 'LGGLLG', 'LGGGLL', 'LGLGLG', 'LGLGGL', 'LGGLGL'];

  // Binario de una tanda de dígitos; `tablas` dice qué tabla usa cada dígito.
  function binarioEan(digitos, tablas) {
    let salida = '';
    for (let i = 0; i < digitos.length; i++) salida += TABLAS_EAN[tablas[i]][digitos[i]];
    return salida;
  }

  function digitoControlEan13(datos) {
    const suma = datos.substr(0, 12).split('').map(Number).reduce(function (acc, e, i) { return i % 2 ? acc + 3 * e : acc + e; }, 0);
    return (10 - suma % 10) % 10;
  }
  function digitoControlUpc(datos) {
    let suma = 0;
    let i;
    for (i = 1; i < 11; i += 2) suma += parseInt(datos[i]);
    for (i = 0; i < 11; i += 2) suma += 3 * parseInt(datos[i]);
    return (10 - suma % 10) % 10;
  }

  // El texto de los números más chico que 10 veces el ancho del módulo, y las
  // barras de guarda más altas que las demás para que el texto entre debajo.
  function medidasGuarda(o) {
    const tamano = o.tamanoTexto > 10 * o.ancho ? 10 * o.ancho : o.tamanoTexto;
    return { tamano: tamano, alturaGuarda: o.alto + tamano / 2 + o.margenTexto };
  }

  function codificarEan13(texto, o) {
    let datos = texto;
    if (/^[0-9]{12}$/.test(datos)) datos += digitoControlEan13(datos);
    if (!/^[0-9]{13}$/.test(datos) || +datos[12] !== digitoControlEan13(datos)) throw entradaNoValida('EAN13', texto);
    const m = medidasGuarda(o);
    const guarda = { alto: m.alturaGuarda };
    const numeros = { tamanoTexto: m.tamano };
    const partes = [
      { data: GUARDA_LATERAL, options: guarda },
      { data: binarioEan(datos.substr(1, 6), ESTRUCTURA_EAN13[datos[0]]), text: datos.substr(1, 6), options: numeros },
      { data: GUARDA_CENTRAL, options: guarda },
      { data: binarioEan(datos.substr(7, 6), 'RRRRRR'), text: datos.substr(7, 6), options: numeros },
      { data: GUARDA_LATERAL, options: guarda },
    ];
    // El primer dígito va fuera de las barras, a la izquierda.
    if (o.mostrarTexto) partes.unshift({ data: '000000000000', text: datos.substr(0, 1), options: { alineacion: 'left', tamanoTexto: m.tamano } });
    return partes;
  }

  function codificarUpc(texto, o) {
    let datos = texto;
    if (/^[0-9]{11}$/.test(datos)) datos += digitoControlUpc(datos);
    if (!/^[0-9]{12}$/.test(datos) || datos[11] != digitoControlUpc(datos)) throw entradaNoValida('UPC', texto);
    const m = medidasGuarda(o);
    const guarda = { alto: m.alturaGuarda };
    const numeros = { tamanoTexto: m.tamano };
    const partes = [];
    // Primer y último dígito van fuera de las barras.
    if (o.mostrarTexto) partes.push({ data: '00000000', text: datos.substr(0, 1), options: { alineacion: 'left', tamanoTexto: m.tamano } });
    partes.push({ data: GUARDA_LATERAL + binarioEan(datos[0], 'L'), options: guarda });
    partes.push({ data: binarioEan(datos.substr(1, 5), 'LLLLL'), text: datos.substr(1, 5), options: numeros });
    partes.push({ data: GUARDA_CENTRAL, options: guarda });
    partes.push({ data: binarioEan(datos.substr(6, 5), 'RRRRR'), text: datos.substr(6, 5), options: numeros });
    partes.push({ data: binarioEan(datos[11], 'R') + GUARDA_LATERAL, options: guarda });
    if (o.mostrarTexto) partes.push({ data: '00000000', text: datos.substr(11, 1), options: { alineacion: 'right', tamanoTexto: m.tamano } });
    return partes;
  }

  // ---------- CODE39 ----------

  const SIMBOLOS_39 = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', '-', '.', ' ', '$', '/', '+', '%', '*'];
  // Patrón de cada símbolo, en el mismo orden (el último es el asterisco de inicio y fin).
  const BARRAS_39 = [
    '101000111011101', '111010001010111', '101110001010111', '111011100010101', '101000111010111', '111010001110101',
    '101110001110101', '101000101110111', '111010001011101', '101110001011101', '111010100010111', '101110100010111',
    '111011101000101', '101011100010111', '111010111000101', '101110111000101', '101010001110111', '111010100011101',
    '101110100011101', '101011100011101', '111010101000111', '101110101000111', '111011101010001', '101011101000111',
    '111010111010001', '101110111010001', '101010111000111', '111010101110001', '101110101110001', '101011101110001',
    '111000101010111', '100011101010111', '111000111010101', '100010111010111', '111000101110101', '100011101110101',
    '100010101110111', '111000101011101', '100011101011101', '100010001000101', '100010001010001', '100010100010001',
    '101000100010001', '1000101110111010',
  ];

  function codificar39(texto) {
    const datos = texto.toUpperCase();
    if (!/^[0-9A-Z\-.\ $/+%]+$/.test(datos)) throw entradaNoValida('CODE39', texto);
    const asterisco = BARRAS_39[SIMBOLOS_39.indexOf('*')];
    let barras = asterisco;
    for (let i = 0; i < datos.length; i++) barras += BARRAS_39[SIMBOLOS_39.indexOf(datos[i])] + '0';
    return [{ data: barras + asterisco, text: datos }];
  }

  // ---------- Opciones y partes ----------

  function opcionesCompletas(opciones) {
    const o = Object.assign({
      formato: 'CODE128', ancho: 2, alto: 100, colorLinea: '#000000', fondo: '#ffffff',
      margen: 10, alineacion: 'center', mostrarTexto: true, tamanoTexto: 20, margenTexto: 2, fuente: 'monospace',
    }, opciones);
    // Margen 0 cuenta como «sin margen» (igual que antes); queda un solo valor para los cuatro lados.
    o.margenArriba = o.margenAbajo = o.margenIzquierda = o.margenDerecha = o.margen;
    return o;
  }

  // Cada parte trae data (0 y 1), text y las opciones que cambian para esa parte
  // (alto de las guardas, tamaño y alineación del texto de EAN y UPC: alto, tamanoTexto, alineacion).
  function codificar(formato, texto, opciones) {
    const o = opcionesCompletas(opciones);
    texto = '' + texto;
    let partes;
    switch (formato) {
      case 'CODE128': partes = codificar128(texto); break;
      case 'EAN13': partes = codificarEan13(texto, o); break;
      case 'UPC': partes = codificarUpc(texto, o); break;
      case 'CODE39': partes = codificar39(texto); break;
      default: throw new Error('Unknown barcode format: ' + formato);
    }
    return partes.map(function (p) { return { data: p.data || '', text: p.text || '', options: p.options || {} }; });
  }

  // ---------- Dibujo en SVG ----------

  function medidor() {
    if (typeof document === 'undefined') return null;
    return document.createElement('canvas').getContext('2d');
  }
  function anchoDeTexto(ctx, texto, o) {
    if (!ctx) return 0;
    ctx.font = ' ' + o.tamanoTexto + 'px ' + o.fuente;
    const medida = ctx.measureText(texto);
    return medida ? medida.width : 0;
  }

  function elemento(nombre) { return document.createElementNS(NS_SVG, nombre); }
  function rectangulo(x, y, ancho, alto, padre) {
    const r = elemento('rect');
    r.setAttribute('x', x);
    r.setAttribute('y', y);
    r.setAttribute('width', ancho);
    r.setAttribute('height', alto);
    padre.appendChild(r);
    return r;
  }

  // Dibuja el código en `svg` (le borra lo que tuviera). Tira EntradaNoValida
  // antes de tocarlo si el texto no sirve para el formato.
  function dibujarSvg(svg, texto, opciones) {
    const base = opcionesCompletas(opciones);
    const partes = codificar(base.formato, texto, base);
    const ctx = medidor();

    // Medidas de cada parte: el ancho es el mayor entre las barras y su texto;
    // si el texto es más ancho, las barras se centran.
    partes.forEach(function (p) {
      p.o = Object.assign({}, base, p.options);
      const anchoTexto = p.o.mostrarTexto ? anchoDeTexto(ctx, p.text, p.o) : 0;
      const anchoBarras = p.data.length * p.o.ancho;
      p.ancho = Math.ceil(Math.max(anchoTexto, anchoBarras));
      p.alto = p.o.alto + (p.o.mostrarTexto && p.text.length > 0 ? p.o.tamanoTexto + p.o.margenTexto : 0) + p.o.margenArriba + p.o.margenAbajo;
      p.relleno = 0;
      if (p.o.mostrarTexto && anchoBarras < anchoTexto) {
        const alineacion = p.o.alineacion;
        if (alineacion === 'center') p.relleno = Math.floor((anchoTexto - anchoBarras) / 2);
        else if (alineacion === 'right') p.relleno = Math.floor(anchoTexto - anchoBarras);
      }
    });
    const anchoTotal = partes.reduce(function (suma, p) { return suma + p.ancho; }, 0);
    const altoTotal = partes.reduce(function (mayor, p) { return p.alto > mayor ? p.alto : mayor; }, 0);
    const anchoSvg = anchoTotal + base.margenIzquierda + base.margenDerecha;

    while (svg.firstChild) svg.removeChild(svg.firstChild);
    svg.setAttribute('width', anchoSvg + 'px');
    svg.setAttribute('height', altoTotal + 'px');
    svg.setAttribute('x', '0px');
    svg.setAttribute('y', '0px');
    svg.setAttribute('viewBox', '0 0 ' + anchoSvg + ' ' + altoTotal);
    svg.setAttribute('xmlns', NS_SVG);
    svg.setAttribute('version', '1.1');
    if (base.fondo) rectangulo(0, 0, anchoSvg, altoTotal, svg).setAttribute('fill', base.fondo);

    let x = base.margenIzquierda;
    partes.forEach(function (p) {
      const o = p.o;
      const grupo = elemento('g');
      grupo.setAttribute('transform', 'translate(' + x + ', ' + o.margenArriba + ')');
      svg.appendChild(grupo);
      grupo.setAttribute('fill', o.colorLinea);

      // Barras: cada racha de 1 seguidos es un solo rectángulo.
      let rachas = 0;
      let posicion = 0;
      for (let i = 0; i < p.data.length; i++) {
        posicion = i * o.ancho + p.relleno;
        if (p.data[i] === '1') rachas++;
        else if (rachas > 0) {
          rectangulo(posicion - o.ancho * rachas, 0, o.ancho * rachas, o.alto, grupo);
          rachas = 0;
        }
      }
      if (rachas > 0) rectangulo(posicion - o.ancho * (rachas - 1), 0, o.ancho * rachas, o.alto, grupo);

      // Texto (también el vacío de las guardas, como antes).
      if (o.mostrarTexto) {
        const t = elemento('text');
        t.setAttribute('font-family', o.fuente);
        t.setAttribute('font-size', o.tamanoTexto);
        const y = o.alto + o.margenTexto + o.tamanoTexto;
        let tx;
        if (o.alineacion === 'left' || p.relleno > 0) {
          tx = 0;
          t.setAttribute('text-anchor', 'start');
        } else if (o.alineacion === 'right') {
          tx = p.ancho - 1;
          t.setAttribute('text-anchor', 'end');
        } else {
          tx = p.ancho / 2;
          t.setAttribute('text-anchor', 'middle');
        }
        t.setAttribute('x', tx);
        t.setAttribute('y', y);
        t.appendChild(document.createTextNode(p.text));
        grupo.appendChild(t);
      }
      x += p.ancho;
    });
  }

  globalThis.CodigoBarras = { FORMATOS: FORMATOS, codificar: codificar, dibujarSvg: dibujarSvg };
})();

/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Tokenizador propio del traductor de documentos: corta el texto en las piezas del vocabulario (SentencePiece unigram de 32.000 piezas,
// compartido por los 9 idiomas) exactamente como lo hace SentencePiece en Python al entrenar, y las vuelve a unir. Sin librerías.
//
//   const tok = TraductorPiezas.crear(json)   json = { piezas, puntajes, tipos } (piezas.json, de exportar-vocabulario.py)
//   tok.codificar('<2es>Hello world')  -> [ids]      (sin BOS/EOS: los agrega quien arma la entrada del modelo)
//   tok.decodificar([ids])             -> 'texto'    (ignora PAD/BOS/EOS)
//   tok.piezas('texto')                -> ['▁Hel', 'lo', …]   (para pruebas)
//   tok.id('<2es>'), tok.PAD, tok.UNK, tok.BOS, tok.EOS
//
// Reglas que imita (normalizer.cc y unigram_model.cc de SentencePiece, normalización «identity»):
//   1. Espacios: se quitan los del principio y del final y se juntan los repetidos (solo el espacio común U+0020); cada espacio pasa a «▁»
//      y se agrega un «▁» al principio (add_dummy_prefix).
//   2. Corte: el camino de mayor puntaje (Viterbi) sobre las piezas normales y las definidas por el usuario. Las definidas valen
//      largo × max_score − 0,1 (max_score arranca en FLT_MIN, como en el original); un carácter sin pieza vale min_score − 10 y es
//      «desconocido». Las sumas se hacen en float32 (Math.fround) y un empate lo gana el primer camino encontrado, como en C++.
//   3. Respaldo por bytes: un carácter desconocido se escribe como sus bytes UTF-8 (<0xE2><0x82>…), así nunca se pierde nada.
(function (raiz) {
  'use strict';

  const ESPACIO = '▁';
  const FLT_MIN = 1.1754943508222875e-38;
  const f = Math.fround;

  function crear(json) {
    const { piezas, puntajes, tipos } = json;
    if (!Array.isArray(piezas) || piezas.length !== puntajes.length || piezas.length !== tipos.length) {
      throw new Error('Vocabulario del traductor inválido');
    }
    const n = piezas.length;
    const puntaje = new Float32Array(puntajes);
    // Diccionario de piezas buscables (normales y definidas por el usuario), por texto; y el largo máximo en caracteres.
    const buscables = new Map();
    const porTexto = new Map();
    const deByte = new Int32Array(256).fill(-1);
    let minimo = Infinity, maximo = FLT_MIN, largoMax = 1;
    let PAD = -1, UNK = -1, BOS = -1, EOS = -1;
    for (let i = 0; i < n; i++) {
      const p = piezas[i], t = tipos[i];
      porTexto.set(p, i);
      if (t === 'n' || t === 'd') {
        buscables.set(p, i);
        const largo = [...p].length;
        if (largo > largoMax) largoMax = largo;
      }
      if (t === 'n') {
        if (puntaje[i] < minimo) minimo = puntaje[i];
        if (puntaje[i] > maximo) maximo = puntaje[i];
      } else if (t === 'u') UNK = i;
      else if (t === 'b') deByte[parseInt(p.slice(3, 5), 16)] = i;
      else if (t === 'c') {
        if (p === '<pad>') PAD = i; else if (p === '<s>') BOS = i; else if (p === '</s>') EOS = i;
      }
    }
    if (UNK < 0) throw new Error('El vocabulario no tiene pieza desconocida');
    const puntajeUnk = f(f(minimo) - 10);
    const definida = (i) => tipos[i] === 'd';

    function normalizar(texto) {
      const sinBordes = String(texto).replace(/ +/g, ' ').replace(/^ | $/g, '');
      return sinBordes ? ESPACIO + sinBordes.replace(/ /g, ESPACIO) : '';
    }

    // Viterbi en una pasada hacia adelante (como EncodeOptimized): mejor[k] = mejor camino que termina en el carácter k.
    function cortar(normal) {
      const chars = [...normal];
      const L = chars.length;
      const puntos = new Float32Array(L + 1);
      const desde = new Int32Array(L + 1).fill(-1);
      const ids = new Int32Array(L + 1);
      for (let i = 0; i < L; i++) {
        if (i > 0 && desde[i] < 0) continue; // inalcanzable (no pasa: siempre hay un camino de a un carácter)
        const base = puntos[i];
        let hayUno = false;
        let sub = '';
        const tope = Math.min(L, i + largoMax);
        for (let j = i; j < tope; j++) {
          sub += chars[j];
          const id = buscables.get(sub);
          if (id === undefined) continue;
          const largo = j - i + 1;
          const valor = definida(id) ? f(f(largo * maximo) - 0.1) : puntaje[id];
          const candidato = f(valor + base);
          const fin = j + 1;
          if (desde[fin] < 0 || candidato > puntos[fin]) {
            puntos[fin] = candidato; desde[fin] = i; ids[fin] = id;
          }
          if (largo === 1) hayUno = true;
        }
        if (!hayUno) {
          const candidato = f(puntajeUnk + base);
          const fin = i + 1;
          if (desde[fin] < 0 || candidato > puntos[fin]) {
            puntos[fin] = candidato; desde[fin] = i; ids[fin] = UNK;
          }
        }
      }
      const salida = [];
      for (let fin = L; fin > 0; fin = desde[fin]) salida.push([chars.slice(desde[fin], fin).join(''), ids[fin]]);
      return salida.reverse();
    }

    const utf8 = new TextEncoder();
    function codificar(texto) {
      const salida = [];
      for (const [trozo, id] of cortar(normalizar(texto))) {
        if (id !== UNK) { salida.push(id); continue; }
        for (const b of utf8.encode(trozo)) salida.push(deByte[b] >= 0 ? deByte[b] : UNK);
      }
      return salida;
    }

    function piezasDe(texto) {
      return codificar(texto).map((i) => piezas[i]);
    }

    // Como el decodificador de SentencePiece: mientras no salió nada, a cada pieza se le quita un «▁» del principio (el del prefijo).
    function decodificar(lista) {
      let texto = '';
      let bytes = [];
      const agregar = (s) => {
        if (!texto && s.startsWith(ESPACIO)) s = s.slice(1);
        texto += s;
      };
      const vaciarBytes = () => {
        if (bytes.length) { agregar(new TextDecoder('utf-8').decode(new Uint8Array(bytes))); bytes = []; }
      };
      for (const id of lista) {
        if (id < 0 || id >= n) continue;
        const t = tipos[id];
        if (t === 'c') continue;
        if (t === 'b') { bytes.push(parseInt(piezas[id].slice(3, 5), 16)); continue; }
        vaciarBytes();
        agregar(t === 'u' ? ' ⁇ ' : piezas[id]);
      }
      vaciarBytes();
      return texto.replace(/▁/g, ' ');
    }

    return {
      codificar, decodificar, piezas: piezasDe,
      id: (p) => (porTexto.has(p) ? porTexto.get(p) : UNK),
      PAD, UNK, BOS, EOS, tamano: n,
    };
  }

  raiz.TraductorPiezas = { crear };
})(typeof self !== 'undefined' ? self : globalThis);

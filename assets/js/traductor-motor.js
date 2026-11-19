/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Motor del traductor de documentos: traduce segmentos de texto con el modelo propio en dos piezas ONNX (codificador y paso del decodificador
// con caché, de exportar-onnx.py) sobre onnxruntime-web. No sabe de documentos ni de hilos: lo usan traductor-hilo.js y las pruebas.
//
//   const motor = await TraductorMotor.crear({ ort, codificador, decodificador, tokenizador })
//        codificador/decodificador: bytes (Uint8Array) o ruta de cada .onnx; tokenizador: TraductorPiezas.crear(json)
//   await motor.traducir(segmentos, destino, { haz = 4, alfa = 0.6 })
//        segmentos: [{ texto, contexto? }] (contexto = la frase anterior, para mantener los términos) -> [texto traducido]
//        destino: 'es', 'en', 'pt', 'fr', 'de', 'it', 'zh', 'ja', 'ko'
//   await motor.cerrar()
//
// Entrada del modelo (igual que al entrenar): «<2xx>» + [contexto + «<ctx>»] + texto, + EOS. Salida: BOS … EOS.
// haz = 1 es la búsqueda voraz y traduce todos los segmentos juntos en un lote (con relleno a la derecha); haz > 1 es búsqueda por haz,
// de a un segmento, con el puntaje normalizado por largo (((5 + n) / 6) ^ alfa, como GNMT).
(function (raiz) {
  'use strict';

  const IDIOMAS = ['de', 'en', 'es', 'fr', 'it', 'ja', 'ko', 'pt', 'zh'];
  const MAX_POS = 512;

  async function crear({ ort, codificador, decodificador, tokenizador, opcionesSesion = {} }) {
    const tok = tokenizador;
    const etiqueta = {};
    for (const i of IDIOMAS) {
      const id = tok.id(`<2${i}>`);
      if (id === tok.UNK) throw new Error(`El vocabulario no tiene la etiqueta <2${i}>`);
      etiqueta[i] = `<2${i}>`;
    }
    const cod = await ort.InferenceSession.create(codificador, opcionesSesion);
    const dec = await ort.InferenceSession.create(decodificador, opcionesSesion);

    const entrada = (s, destino) => {
      const texto = s.contexto ? `${etiqueta[destino]}${s.contexto}<ctx>${s.texto}` : `${etiqueta[destino]}${s.texto}`;
      return [...tok.codificar(texto), tok.EOS];
    };
    const tensorEnteros = (lista, forma) => new ort.Tensor('int64', BigInt64Array.from(lista, (x) => BigInt(x)), forma);

    async function codificar(listas) {
      const S = Math.max(...listas.map((l) => l.length));
      const plano = [];
      for (const l of listas) { plano.push(...l); for (let k = l.length; k < S; k++) plano.push(tok.PAD); }
      const r = await cod.run({ src: tensorEnteros(plano, [listas.length, S]) });
      return { cruzada: r.cruzada, mascara: r.mascara, S };
    }

    async function paso(piezas, pos, cruzada, mascara, pasado) {
      const r = await dec.run({ pieza: tensorEnteros(piezas, [piezas.length, 1]), pos: tensorEnteros([pos], [1]), cruzada, mascara, pasado });
      return { logprob: r.logprob.data, V: r.logprob.dims[1], presente: r.presente };
    }

    const pasadoVacio = (cruzada, B) => {
      const [L, , , H, , dk] = cruzada.dims;
      return new ort.Tensor('float32', new Float32Array(0), [L, 2, B, H, 0, dk]);
    };
    const limite = (S) => Math.min(256, Math.floor(S * 2.5) + 10, MAX_POS);

    // Voraz, en lote: todos los segmentos a la vez.
    async function voraz(listas) {
      const B = listas.length;
      const { cruzada, mascara, S } = await codificar(listas);
      let pasado = pasadoVacio(cruzada, B);
      let piezas = new Array(B).fill(tok.BOS);
      const salida = listas.map(() => []);
      const vivo = new Array(B).fill(true);
      for (let t = 0; t < limite(S); t++) {
        const r = await paso(piezas, t, cruzada, mascara, pasado);
        pasado.dispose?.();
        pasado = r.presente;
        for (let b = 0; b < B; b++) {
          let mejor = 0, valor = -Infinity;
          const base = b * r.V;
          for (let v = 0; v < r.V; v++) if (r.logprob[base + v] > valor) { valor = r.logprob[base + v]; mejor = v; }
          piezas[b] = mejor;
          if (!vivo[b]) continue;
          if (mejor === tok.EOS) vivo[b] = false; else salida[b].push(mejor);
        }
        if (!vivo.some(Boolean)) break;
      }
      pasado.dispose?.();
      cruzada.dispose?.(); mascara.dispose?.();
      return salida;
    }

    // Repite el lote de la atención cruzada para k haces: (L, 2, 1, H, S, dk) -> (L, 2, k, H, S, dk); la máscara (1,1,1,S) -> (k,1,1,S).
    function repetir(tensor, k, ejeLote) {
      const dims = tensor.dims.slice();
      const antes = dims.slice(0, ejeLote).reduce((a, b) => a * b, 1);
      const bloque = dims.slice(ejeLote + 1).reduce((a, b) => a * b, 1);
      const origen = tensor.data;
      const datos = new Float32Array(antes * k * bloque);
      for (let a = 0; a < antes; a++) for (let j = 0; j < k; j++) datos.set(origen.subarray(a * bloque, (a + 1) * bloque), (a * k + j) * bloque);
      dims[ejeLote] = k;
      return new ort.Tensor('float32', datos, dims);
    }

    // Reordena la caché del decodificador según de qué haz viene cada uno: (L, 2, B, H, T, dk), eje del lote = 2.
    function reordenar(tensor, origenes) {
      const [L, dos, B, H, T, dk] = tensor.dims;
      const bloque = H * T * dk;
      const datos = new Float32Array(L * dos * origenes.length * bloque);
      const o = tensor.data;
      for (let a = 0; a < L * dos; a++) {
        for (let j = 0; j < origenes.length; j++) {
          const desde = (a * B + origenes[j]) * bloque;
          datos.set(o.subarray(desde, desde + bloque), (a * origenes.length + j) * bloque);
        }
      }
      return new ort.Tensor('float32', datos, [L, dos, origenes.length, H, T, dk]);
    }

    async function porHaz(lista, k, alfa) {
      const { cruzada: c1, mascara: m1, S } = await codificar([lista]);
      const normal = (puntaje, n) => puntaje / Math.pow((5 + n) / 6, alfa);
      let haces = [{ ids: [], puntaje: 0 }];
      let cruzada = c1, mascara = m1;
      let pasado = pasadoVacio(cruzada, 1);
      const terminados = [];
      for (let t = 0; t < limite(S) && haces.length; t++) {
        if (cruzada.dims[2] !== haces.length) {
          cruzada = repetir(c1, haces.length, 2);
          mascara = repetir(m1, haces.length, 0);
        }
        const r = await paso(haces.map((h) => (h.ids.length ? h.ids[h.ids.length - 1] : tok.BOS)), t, cruzada, mascara, pasado);
        // Los k mejores de cada haz, y de todos ellos los k mejores en total.
        const candidatos = [];
        for (let b = 0; b < haces.length; b++) {
          const base = b * r.V;
          const mejores = [];
          for (let v = 0; v < r.V; v++) {
            const lp = r.logprob[base + v];
            if (mejores.length < k) { mejores.push([lp, v]); mejores.sort((x, y) => y[0] - x[0]); }
            else if (lp > mejores[k - 1][0]) { mejores[k - 1] = [lp, v]; mejores.sort((x, y) => y[0] - x[0]); }
          }
          for (const [lp, v] of mejores) candidatos.push({ b, v, puntaje: haces[b].puntaje + lp });
        }
        candidatos.sort((x, y) => y.puntaje - x.puntaje);
        const siguientes = [];
        const origenes = [];
        for (const c of candidatos) {
          if (c.v === tok.EOS) {
            terminados.push({ ids: haces[c.b].ids, final: normal(c.puntaje, haces[c.b].ids.length + 1) });
          } else if (siguientes.length < k) {
            siguientes.push({ ids: [...haces[c.b].ids, c.v], puntaje: c.puntaje });
            origenes.push(c.b);
          }
          if (siguientes.length >= k) break;
        }
        // Parar cuando ningún haz vivo puede superar al mejor terminado (con el puntaje actual, que solo baja).
        const mejorFin = terminados.reduce((a, x) => Math.max(a, x.final), -Infinity);
        const vivos = siguientes.filter((h) => normal(h.puntaje, limite(S)) > mejorFin || normal(h.puntaje, h.ids.length) > mejorFin);
        const presente = r.presente;
        if (!vivos.length || terminados.length >= k) { presente.dispose?.(); haces = []; break; }
        const nuevo = reordenar(presente, origenes);
        presente.dispose?.();
        pasado = nuevo;
        haces = siguientes;
      }
      for (const h of haces) terminados.push({ ids: h.ids, final: normal(h.puntaje, h.ids.length) });
      terminados.sort((x, y) => y.final - x.final);
      return terminados.length ? terminados[0].ids : [];
    }

    async function traducir(segmentos, destino, { haz = 4, alfa = 0.6 } = {}) {
      if (!etiqueta[destino]) throw new Error(`Idioma de destino no disponible: ${destino}`);
      const listas = segmentos.map((s) => entrada(s, destino));
      const ids = haz <= 1 ? await voraz(listas) : await (async () => {
        const r = [];
        for (const l of listas) r.push(await porHaz(l, haz, alfa));
        return r;
      })();
      return ids.map((x) => tok.decodificar(x));
    }

    return {
      traducir,
      idsVoraz: (segmentos, destino) => voraz(segmentos.map((s) => entrada(s, destino))),
      cerrar: async () => { await cod.release?.(); await dec.release?.(); },
    };
  }

  raiz.TraductorMotor = { crear, IDIOMAS };
})(typeof self !== 'undefined' ? self : globalThis);

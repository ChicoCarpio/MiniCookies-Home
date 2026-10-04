/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Lector de EMF (Enhanced Metafile de Windows) → canvas.
//
// Office guarda así las tablas y gráficos pegados desde Excel: una lista de
// órdenes de dibujo (líneas, rellenos, texto) que ningún navegador muestra.
// Se probó EMFJS (rtf.js) y solo dibujaba las líneas, sin texto ni rellenos:
// para una tabla, casi nada. Este lector cubre lo que usan esos dibujos:
//  - Sistema de coordenadas: modos de mapeo, ventana/viewport y transformación
//    del mundo.
//  - Objetos: lápices, pinceles y fuentes (también los de stock).
//  - Líneas, polilíneas, polígonos, rectángulos, elipses y trazados (paths).
//  - Rellenos BITBLT/PATBLT sin imagen (el fondo de las celdas) y mapas de bits
//    (BITBLT, STRETCHBLT, STRETCHDIBITS, SETDIBITSTODEVICE).
//  - Texto (EXTTEXTOUTW/A) con su fuente, color, alineación, giro y el ancho de
//    cada letra que guardó Windows.
//  - Recortes rectangulares y SAVEDC/RESTOREDC.
// Lo que no cubre se ignora (no rompe el dibujo). Los comentarios EMF+ también:
// los archivos "duales" traen además las órdenes EMF clásicas.
//
// Uso: const canvas = await Emf.dibujar(bytes, { escalaMaxima: 2, ladoMaximo: 3000 });

(function () {
  // Tamaños de los modos de mapeo fijos, en milímetros por unidad lógica.
  const MM_POR_UNIDAD = { 2: 0.1, 3: 0.01, 4: 0.254, 5: 0.0254, 6: 25.4 / 1440 };
  const STOCK = {
    0: { tipo: 'pincel', estilo: 0, color: '#ffffff' }, 1: { tipo: 'pincel', estilo: 0, color: '#c0c0c0' },
    2: { tipo: 'pincel', estilo: 0, color: '#808080' }, 3: { tipo: 'pincel', estilo: 0, color: '#404040' },
    4: { tipo: 'pincel', estilo: 0, color: '#000000' }, 5: { tipo: 'pincel', estilo: 1 },
    6: { tipo: 'lapiz', estilo: 0, ancho: 0, color: '#ffffff' }, 7: { tipo: 'lapiz', estilo: 0, ancho: 0, color: '#000000' },
    8: { tipo: 'lapiz', estilo: 5 }, 18: { tipo: 'pincel', estilo: 0, color: '#ffffff' }, 19: { tipo: 'lapiz', estilo: 0, ancho: 0, color: '#000000' },
  };
  const FUENTE_STOCK = { tipo: 'fuente', alto: -12, peso: 400, cursiva: false, subrayado: false, giro: 0, nombre: 'Arial' };

  const colorRef = (v, off) => `rgb(${v.getUint8(off)},${v.getUint8(off + 1)},${v.getUint8(off + 2)})`;

  function estadoInicial() {
    return {
      modo: 1, winOrg: [0, 0], winExt: [1, 1], vpOrg: [0, 0], vpExt: [1, 1], mundo: [1, 0, 0, 1, 0, 0],
      lapiz: STOCK[7], pincel: STOCK[0], fuente: FUENTE_STOCK,
      colorTexto: 'rgb(0,0,0)', colorFondo: 'rgb(255,255,255)', fondoOpaco: true, alineacion: 0, relleno: 'evenodd',
      actual: [0, 0], recortes: null,
    };
  }

  function copiarEstado(e) {
    return { ...e, winOrg: [...e.winOrg], winExt: [...e.winExt], vpOrg: [...e.vpOrg], vpExt: [...e.vpExt], mundo: [...e.mundo], actual: [...e.actual], recortes: e.recortes && e.recortes.map((r) => [...r]) };
  }

  function multiplicar(m, n) {
    return [m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1], m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3], m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5]];
  }

  // DIB (BITMAPINFOHEADER + píxeles) → canvas.
  async function dibACanvas(v, bytes, offBmi, cbBmi, offBits, cbBits) {
    const ancho = v.getInt32(offBmi + 4, true);
    let alto = v.getInt32(offBmi + 8, true);
    const deArribaAbajo = alto < 0;
    alto = Math.abs(alto);
    const bits = v.getUint16(offBmi + 14, true);
    const compresion = v.getUint32(offBmi + 16, true);
    const tamCabecera = v.getUint32(offBmi, true);
    const canvas = document.createElement('canvas');
    canvas.width = ancho; canvas.height = alto;
    const ctx = canvas.getContext('2d');
    if (compresion === 4 || compresion === 5) { // JPEG / PNG incrustado
      const mapa = await createImageBitmap(new Blob([bytes.subarray(offBits, offBits + cbBits)]));
      ctx.drawImage(mapa, 0, 0);
      return canvas;
    }
    let colores = v.getUint32(offBmi + 32, true);
    if (!colores && bits <= 8) colores = 1 << bits;
    const paleta = [];
    const offPaleta = offBmi + tamCabecera + (compresion === 3 && tamCabecera === 40 ? 12 : 0);
    for (let i = 0; i < colores && offPaleta + i * 4 + 3 < offBmi + cbBmi + 1024; i++) {
      paleta.push([v.getUint8(offPaleta + i * 4 + 2), v.getUint8(offPaleta + i * 4 + 1), v.getUint8(offPaleta + i * 4)]);
    }
    const datos = ctx.createImageData(ancho, alto);
    const fila = Math.floor((ancho * bits + 31) / 32) * 4;
    for (let y = 0; y < alto; y++) {
      const origen = offBits + (deArribaAbajo ? y : alto - 1 - y) * fila;
      for (let x = 0; x < ancho; x++) {
        let r = 0, g = 0, b = 0;
        if (bits === 24 || bits === 32) {
          const o = origen + x * (bits / 8);
          b = bytes[o]; g = bytes[o + 1]; r = bytes[o + 2];
        } else if (bits === 16) {
          const p = bytes[origen + x * 2] | (bytes[origen + x * 2 + 1] << 8);
          r = ((p >> 10) & 31) << 3; g = ((p >> 5) & 31) << 3; b = (p & 31) << 3;
        } else {
          const porByte = 8 / bits;
          const byte = bytes[origen + Math.floor(x / porByte)];
          const indice = (byte >> ((porByte - 1 - (x % porByte)) * bits)) & ((1 << bits) - 1);
          [r, g, b] = paleta[indice] || [0, 0, 0];
        }
        const d = (y * ancho + x) * 4;
        datos.data[d] = r; datos.data[d + 1] = g; datos.data[d + 2] = b; datos.data[d + 3] = 255;
      }
    }
    ctx.putImageData(datos, 0, 0);
    return canvas;
  }

  async function dibujar(entrada, opciones = {}) {
    const bytes = entrada instanceof Uint8Array ? entrada : new Uint8Array(entrada);
    const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    if (v.getUint32(0, true) !== 1 || v.getUint32(40, true) !== 0x464d4520) throw new Error('No es un archivo EMF.');
    const limites = [v.getInt32(8, true), v.getInt32(12, true), v.getInt32(16, true), v.getInt32(20, true)];
    const dispositivoPx = [v.getInt32(72, true), v.getInt32(76, true)];
    const dispositivoMm = [v.getInt32(80, true), v.getInt32(84, true)];
    const ancho = limites[2] - limites[0] + 1;
    const alto = limites[3] - limites[1] + 1;
    const escala = Math.min(opciones.escalaMaxima || 2, (opciones.ladoMaximo || 3000) / Math.max(ancho, alto));
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round(ancho * escala));
    canvas.height = Math.max(1, Math.round(alto * escala));
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    let e = estadoInicial();
    const pila = [];
    const objetos = [];
    let enTrazado = false;
    let trazado = null; // Path2D en coordenadas lógicas, entre BEGINPATH y ENDPATH

    // Lógicas → canvas.
    function matriz() {
      let sx = 1, sy = 1;
      if (e.modo === 7 || e.modo === 8) {
        sx = e.vpExt[0] / (e.winExt[0] || 1); sy = e.vpExt[1] / (e.winExt[1] || 1);
        if (e.modo === 7) { const s = Math.min(Math.abs(sx), Math.abs(sy)); sx = Math.sign(sx) * s; sy = Math.sign(sy) * s; }
      } else if (MM_POR_UNIDAD[e.modo]) {
        sx = (MM_POR_UNIDAD[e.modo] * dispositivoPx[0]) / (dispositivoMm[0] || 1);
        sy = -(MM_POR_UNIDAD[e.modo] * dispositivoPx[1]) / (dispositivoMm[1] || 1);
      }
      const pagina = [sx, 0, 0, sy, e.vpOrg[0] - e.winOrg[0] * sx, e.vpOrg[1] - e.winOrg[1] * sy];
      const lienzo = [escala, 0, 0, escala, -limites[0] * escala, -limites[1] * escala];
      return multiplicar(lienzo, multiplicar(pagina, e.mundo));
    }

    function conRecorte(dibujo) {
      ctx.save();
      if (e.recortes) {
        ctx.setTransform(1, 0, 0, 1, 0, 0);
        const zona = new Path2D();
        for (const [x0, y0, x1, y1] of e.recortes) zona.rect((x0 - limites[0]) * escala, (y0 - limites[1]) * escala, (x1 - x0) * escala, (y1 - y0) * escala);
        ctx.clip(zona);
      }
      const m = matriz();
      ctx.setTransform(m[0], m[1], m[2], m[3], m[4], m[5]);
      dibujo(m);
      ctx.restore();
    }

    function aplicarLapiz(m) {
      const l = e.lapiz;
      if (!l || l.estilo === 5) return false;
      const escalaM = Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])) || 1;
      ctx.strokeStyle = l.color;
      ctx.lineWidth = Math.max(l.ancho || 0, escala / escalaM); // ancho 0: un píxel del dispositivo
      const paso = escala / escalaM;
      ctx.setLineDash({ 1: [6 * paso, 3 * paso], 2: [1.5 * paso, 1.5 * paso], 3: [6 * paso, 2 * paso, 1.5 * paso, 2 * paso] }[l.estilo] || []);
      return true;
    }
    function aplicarPincel() {
      const p = e.pincel;
      if (!p || p.estilo === 1) return false;
      ctx.fillStyle = p.color;
      return true;
    }

    function dibujarForma(construir, rellenar) {
      if (enTrazado) { construir(trazado); return; }
      conRecorte((m) => {
        const forma = new Path2D();
        construir(forma);
        if (rellenar && aplicarPincel()) ctx.fill(forma, e.relleno);
        if (aplicarLapiz(m)) ctx.stroke(forma);
      });
    }

    function rellenarRect(x, y, w, h, color) {
      conRecorte(() => { ctx.fillStyle = color; ctx.fillRect(x, y, w, h); });
    }

    const puntos16 = (o, n) => Array.from({ length: n }, (_, i) => [v.getInt16(o + i * 4, true), v.getInt16(o + i * 4 + 2, true)]);
    const puntos32 = (o, n) => Array.from({ length: n }, (_, i) => [v.getInt32(o + i * 8, true), v.getInt32(o + i * 8 + 4, true)]);

    function figura(pts, cerrar, bezier, desdeActual) {
      return (f) => {
        let i = 0;
        if (desdeActual) f.moveTo(e.actual[0], e.actual[1]);
        else { f.moveTo(pts[0][0], pts[0][1]); i = 1; }
        if (bezier) for (; i + 2 < pts.length; i += 3) f.bezierCurveTo(pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1], pts[i + 2][0], pts[i + 2][1]);
        else for (; i < pts.length; i++) f.lineTo(pts[i][0], pts[i][1]);
        if (cerrar) f.closePath();
      };
    }

    function texto(o, ancho16) {
      // EMREXTTEXTOUTW: bounds(16) modo(4) escalaX(4) escalaY(4) EMRTEXT
      const ref = [v.getInt32(o + 36, true), v.getInt32(o + 40, true)];
      const n = v.getUint32(o + 44, true);
      const offTexto = v.getUint32(o + 48, true);
      const opciones = v.getUint32(o + 52, true);
      const rect = [v.getInt32(o + 56, true), v.getInt32(o + 60, true), v.getInt32(o + 64, true), v.getInt32(o + 68, true)];
      const offDx = v.getUint32(o + 72, true);
      let cadena = '';
      for (let i = 0; i < n; i++) cadena += ancho16 ? String.fromCharCode(v.getUint16(o + offTexto + i * 2, true)) : String.fromCharCode(v.getUint8(o + offTexto + i));
      const dx = offDx ? Array.from({ length: n }, (_, i) => v.getInt32(o + offDx + i * 4, true)) : null;
      if (opciones & 2) rellenarRect(rect[0], rect[1], rect[2] - rect[0], rect[3] - rect[1], e.colorFondo);
      if (!cadena.trim()) return;
      const f = e.fuente;
      const tamano = Math.abs(f.alto) || 12;
      const alinea = e.alineacion;
      const origen = alinea & 1 ? e.actual : ref; // TA_UPDATECP
      conRecorte((m) => {
        if (opciones & 4) { ctx.beginPath(); ctx.rect(rect[0], rect[1], rect[2] - rect[0], rect[3] - rect[1]); ctx.clip(); }
        ctx.translate(origen[0], origen[1]);
        // Letras derechas aunque el eje Y del archivo apunte hacia arriba.
        ctx.scale(Math.sign(m[0]) || 1, Math.sign(m[3]) || 1);
        if (f.giro) ctx.rotate((-f.giro / 10) * Math.PI / 180 * (Math.sign(m[3]) || 1));
        ctx.font = `${f.cursiva ? 'italic ' : ''}${f.peso >= 600 ? 'bold' : 'normal'} ${tamano}px "${f.nombre}", Arial, sans-serif`;
        ctx.fillStyle = e.colorTexto;
        ctx.textBaseline = (alinea & 24) === 24 ? 'alphabetic' : alinea & 8 ? 'bottom' : 'top';
        const total = dx ? dx.reduce((a, b) => a + b, 0) : ctx.measureText(cadena).width;
        let x = (alinea & 6) === 6 ? -total / 2 : alinea & 2 ? -total : 0;
        if (dx) {
          for (let i = 0; i < cadena.length; i++) { ctx.fillText(cadena[i], x, 0); x += dx[i]; }
        } else {
          ctx.fillText(cadena, x, 0);
        }
        if (f.subrayado) ctx.fillRect(x - total, tamano * 0.1, total, Math.max(1, tamano / 15));
      });
      if (alinea & 1 && dx) e.actual = [origen[0] + dx.reduce((a, b) => a + b, 0), origen[1]];
    }

    async function mapaDeBits(o, tipo) {
      let x, y, w, h, rop, offBmi, cbBmi, offBits, cbBits, sx = 0, sy = 0, sw, sh;
      if (tipo === 81) { // STRETCHDIBITS
        x = v.getInt32(o + 24, true); y = v.getInt32(o + 28, true); sx = v.getInt32(o + 32, true); sy = v.getInt32(o + 36, true);
        sw = v.getInt32(o + 40, true); sh = v.getInt32(o + 44, true);
        offBmi = v.getUint32(o + 48, true); cbBmi = v.getUint32(o + 52, true); offBits = v.getUint32(o + 56, true); cbBits = v.getUint32(o + 60, true);
        rop = v.getUint32(o + 68, true); w = v.getInt32(o + 72, true); h = v.getInt32(o + 76, true);
      } else if (tipo === 80) { // SETDIBITSTODEVICE
        x = v.getInt32(o + 24, true); y = v.getInt32(o + 28, true); sx = v.getInt32(o + 32, true); sy = v.getInt32(o + 36, true);
        w = sw = v.getInt32(o + 40, true); h = sh = v.getInt32(o + 44, true);
        offBmi = v.getUint32(o + 48, true); cbBmi = v.getUint32(o + 52, true); offBits = v.getUint32(o + 56, true); cbBits = v.getUint32(o + 60, true);
        rop = 0x00cc0020;
      } else { // BITBLT (76) / STRETCHBLT (77)
        x = v.getInt32(o + 24, true); y = v.getInt32(o + 28, true); w = v.getInt32(o + 32, true); h = v.getInt32(o + 36, true);
        rop = v.getUint32(o + 40, true); sx = v.getInt32(o + 44, true); sy = v.getInt32(o + 48, true);
        offBmi = v.getUint32(o + 84, true); cbBmi = v.getUint32(o + 88, true); offBits = v.getUint32(o + 92, true); cbBits = v.getUint32(o + 96, true);
        if (tipo === 77) { sw = v.getInt32(o + 100, true); sh = v.getInt32(o + 104, true); }
      }
      if (!cbBits) {
        // Sin imagen: un relleno con el pincel (así pinta Excel el fondo de las celdas).
        const color = rop === 0x00000042 ? '#000' : rop === 0x00ff0062 ? '#fff' : e.pincel && e.pincel.estilo !== 1 ? e.pincel.color : null;
        if (color) rellenarRect(x, y, w, h, color);
        return;
      }
      const imagen = await dibACanvas(v, bytes, o + offBmi, cbBmi, o + offBits, cbBits);
      conRecorte(() => {
        ctx.imageSmoothingQuality = 'high';
        ctx.drawImage(imagen, sx, sy, sw || imagen.width, sh || imagen.height, x, y, w, h);
      });
    }

    for (let o = 0; o + 8 <= bytes.length;) {
      const tipo = v.getUint32(o, true);
      const tam = v.getUint32(o + 4, true);
      if (tam < 8 || o + tam > bytes.length) break;
      const p = o + 8;
      try {
        switch (tipo) {
          case 9: e.winExt = [v.getInt32(p, true), v.getInt32(p + 4, true)]; break;
          case 10: e.winOrg = [v.getInt32(p, true), v.getInt32(p + 4, true)]; break;
          case 11: e.vpExt = [v.getInt32(p, true), v.getInt32(p + 4, true)]; break;
          case 12: e.vpOrg = [v.getInt32(p, true), v.getInt32(p + 4, true)]; break;
          case 17: e.modo = v.getUint32(p, true); break;
          case 18: e.fondoOpaco = v.getUint32(p, true) === 2; break;
          case 19: e.relleno = v.getUint32(p, true) === 2 ? 'nonzero' : 'evenodd'; break;
          case 22: e.alineacion = v.getUint32(p, true); break;
          case 24: e.colorTexto = colorRef(v, p); break;
          case 25: e.colorFondo = colorRef(v, p); break;
          case 27: e.actual = [v.getInt32(p, true), v.getInt32(p + 4, true)]; if (enTrazado) trazado.moveTo(e.actual[0], e.actual[1]); break;
          case 54: { // LINETO
            const a = [...e.actual];
            const b = [v.getInt32(p, true), v.getInt32(p + 4, true)];
            if (enTrazado) trazado.lineTo(b[0], b[1]);
            else dibujarForma((f) => { f.moveTo(a[0], a[1]); f.lineTo(b[0], b[1]); }, false);
            e.actual = b;
            break;
          }
          case 33: pila.push(copiarEstado(e)); break;
          case 34: { const n = v.getInt32(p, true); for (let i = 0; i < -n - 1 && pila.length > 1; i++) pila.pop(); if (pila.length) e = pila.pop(); break; }
          case 35: e.mundo = [v.getFloat32(p, true), v.getFloat32(p + 4, true), v.getFloat32(p + 8, true), v.getFloat32(p + 12, true), v.getFloat32(p + 16, true), v.getFloat32(p + 20, true)]; break;
          case 36: {
            const t = [v.getFloat32(p, true), v.getFloat32(p + 4, true), v.getFloat32(p + 8, true), v.getFloat32(p + 12, true), v.getFloat32(p + 16, true), v.getFloat32(p + 20, true)];
            const modo = v.getUint32(p + 24, true);
            e.mundo = modo === 1 ? [1, 0, 0, 1, 0, 0] : modo === 2 ? multiplicar(e.mundo, t) : modo === 3 ? multiplicar(t, e.mundo) : t;
            break;
          }
          case 37: { // SELECTOBJECT
            const id = v.getUint32(p, true);
            const stock = id & 0x7fffffff;
            const obj = id & 0x80000000 ? STOCK[stock] || ([10, 11, 12, 13, 14, 16, 17].includes(stock) ? FUENTE_STOCK : null) : objetos[id];
            if (obj && obj.tipo === 'lapiz') e.lapiz = obj;
            else if (obj && obj.tipo === 'pincel') e.pincel = obj;
            else if (obj && obj.tipo === 'fuente') e.fuente = obj;
            break;
          }
          case 38: objetos[v.getUint32(p, true)] = { tipo: 'lapiz', estilo: v.getUint32(p + 4, true) & 0xf, ancho: v.getInt32(p + 8, true), color: colorRef(v, p + 16) }; break;
          case 95: objetos[v.getUint32(p, true)] = { tipo: 'lapiz', estilo: v.getUint32(p + 20, true) & 0xf, ancho: v.getUint32(p + 24, true), color: colorRef(v, p + 32) }; break;
          case 39: objetos[v.getUint32(p, true)] = { tipo: 'pincel', estilo: v.getUint32(p + 4, true) === 1 ? 1 : 0, color: colorRef(v, p + 8) }; break;
          case 93: case 94: objetos[v.getUint32(p, true)] = { tipo: 'pincel', estilo: 0, color: '#999' }; break; // pincel de patrón: gris
          case 82: {
            let nombre = '';
            for (let i = 0; i < 32; i++) { const c = v.getUint16(p + 32 + i * 2, true); if (!c) break; nombre += String.fromCharCode(c); }
            objetos[v.getUint32(p, true)] = { tipo: 'fuente', alto: v.getInt32(p + 4, true), giro: v.getInt32(p + 12, true), peso: v.getInt32(p + 20, true), cursiva: !!v.getUint8(p + 24), subrayado: !!v.getUint8(p + 25), nombre: nombre || 'Arial' };
            break;
          }
          case 40: objetos[v.getUint32(p, true)] = undefined; break;
          case 30: { // INTERSECTCLIPRECT (lógicas → dispositivo con la matriz del momento)
            const m = matriz();
            const a = [v.getInt32(p, true), v.getInt32(p + 4, true)];
            const b = [v.getInt32(p + 8, true), v.getInt32(p + 12, true)];
            const aDisp = (q) => [(m[0] * q[0] + m[2] * q[1] + m[4]) / escala + limites[0], (m[1] * q[0] + m[3] * q[1] + m[5]) / escala + limites[1]];
            const [x0, y0] = aDisp(a); const [x1, y1] = aDisp(b);
            const nuevo = [Math.min(x0, x1), Math.min(y0, y1), Math.max(x0, x1), Math.max(y0, y1)];
            e.recortes = (e.recortes || [[-1e9, -1e9, 1e9, 1e9]]).map((r) => [Math.max(r[0], nuevo[0]), Math.max(r[1], nuevo[1]), Math.min(r[2], nuevo[2]), Math.min(r[3], nuevo[3])]);
            break;
          }
          case 75: { // EXTSELECTCLIPRGN (rectángulos en coordenadas de dispositivo)
            const cb = v.getUint32(p, true);
            const modo = v.getUint32(p + 4, true);
            if (!cb) { if (modo === 5) e.recortes = null; break; }
            const n = v.getUint32(p + 8 + 8, true);
            const rects = Array.from({ length: n }, (_, i) => { const q = p + 8 + 32 + i * 16; return [v.getInt32(q, true), v.getInt32(q + 4, true), v.getInt32(q + 8, true), v.getInt32(q + 12, true)]; });
            e.recortes = modo === 1 && e.recortes ? e.recortes.flatMap((r) => rects.map((s) => [Math.max(r[0], s[0]), Math.max(r[1], s[1]), Math.min(r[2], s[2]), Math.min(r[3], s[3])])) : rects;
            break;
          }
          case 43: { const [x0, y0, x1, y1] = [v.getInt32(p, true), v.getInt32(p + 4, true), v.getInt32(p + 8, true), v.getInt32(p + 12, true)]; dibujarForma((f) => f.rect(x0, y0, x1 - x0, y1 - y0), true); break; }
          case 44: { const [x0, y0, x1, y1] = [v.getInt32(p, true), v.getInt32(p + 4, true), v.getInt32(p + 8, true), v.getInt32(p + 12, true)]; const r = [v.getInt32(p + 16, true) / 2, v.getInt32(p + 20, true) / 2]; dibujarForma((f) => f.roundRect(x0, y0, x1 - x0, y1 - y0, Math.min(r[0], r[1])), true); break; }
          case 42: { const [x0, y0, x1, y1] = [v.getInt32(p, true), v.getInt32(p + 4, true), v.getInt32(p + 8, true), v.getInt32(p + 12, true)]; dibujarForma((f) => f.ellipse((x0 + x1) / 2, (y0 + y1) / 2, Math.abs(x1 - x0) / 2, Math.abs(y1 - y0) / 2, 0, 0, Math.PI * 2), true); break; }
          case 86: case 87: case 85: case 88: case 89: case 3: case 4: case 2: case 5: case 6: {
            const n = v.getUint32(p + 16, true);
            const corto = tipo >= 85;
            const pts = corto ? puntos16(p + 20, n) : puntos32(p + 20, n);
            const base = corto ? tipo - 83 : tipo; // 2 bezier, 3 polígono, 4 polilínea, 5 bezierTo, 6 polilíneaTo
            const desdeActual = base === 5 || base === 6;
            if (enTrazado) figura(pts, base === 3, base === 2 || base === 5, desdeActual)(trazado);
            else dibujarForma(figura(pts, base === 3, base === 2 || base === 5, desdeActual), base === 3);
            if (desdeActual || base === 4) e.actual = pts[pts.length - 1] || e.actual;
            break;
          }
          case 90: case 91: case 7: case 8: { // POLYPOLYLINE/POLYPOLYGON
            const poligonos = v.getUint32(p + 16, true);
            const corto = tipo >= 90;
            const conteos = Array.from({ length: poligonos }, (_, i) => v.getUint32(p + 24 + i * 4, true));
            let q = p + 24 + poligonos * 4;
            const cerrar = tipo === 91 || tipo === 8;
            dibujarForma((f) => {
              for (const c of conteos) {
                const pts = corto ? puntos16(q, c) : puntos32(q, c);
                q += c * (corto ? 4 : 8);
                figura(pts, cerrar, false, false)(f);
              }
            }, cerrar);
            break;
          }
          case 59: enTrazado = true; trazado = new Path2D(); break;
          case 60: enTrazado = false; break;
          case 61: if (trazado) trazado.closePath(); break;
          case 68: trazado = null; enTrazado = false; break;
          case 62: case 63: case 64: {
            const t = trazado;
            if (t) conRecorte((m) => { if (tipo !== 64 && aplicarPincel()) ctx.fill(t, e.relleno); if (tipo !== 62 && aplicarLapiz(m)) ctx.stroke(t); });
            trazado = null;
            break;
          }
          case 76: case 77: case 80: case 81: await mapaDeBits(p - 8, tipo); break;
          case 84: texto(p - 8, true); break;
          case 83: texto(p - 8, false); break;
          case 14: o = bytes.length; continue; // EOF
          default: break;
        }
      } catch {
        // Un registro que no se entiende no tira abajo el resto del dibujo.
      }
      o += tam;
    }
    return canvas;
  }

  window.Emf = { dibujar };
})();

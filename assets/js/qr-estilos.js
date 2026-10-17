/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Códigos QR con estilo: la parte opcional del generador de QR. Toma la matriz
// que arma qr-propio.js y la dibuja con módulos redondeados o puntos,
// ojos (los tres cuadros de las esquinas) redondeados, colores propios y un logo
// en el centro. Sin red ni librerías: todo se dibuja en un canvas.
//
// Lo que NO hace es tocar el codificador. Sin opciones (módulos cuadrados, ojos
// cuadrados, sin logo) delega en QRPropio.pintar y el resultado es el de
// siempre, píxel por píxel.
//
// Reglas que protegen la lectura (el QR tiene que seguir escaneándose):
//  - Con logo, el QR se arma con corrección H (30 %) y la placa del logo ocupa
//    como mucho el 22 % del ancho (menos del 5 % de los módulos). Se limpian
//    módulos enteros, alineados a la cuadrícula, y nunca los cuadros de
//    posición, las líneas de temporización, los patrones de alineación ni la
//    información de formato y de versión. Si en una versión no hay lugar (la de
//    alineación del centro cae justo debajo), se sube de versión.
//  - El código siempre es más oscuro que el fondo y el contraste mínimo es 4,5:1
//    (WCAG). Un par que no llega se corrige (oscureciendo el código y aclarando
//    el fondo) y se avisa en el resultado.
//  - Se conservan los 4 módulos de zona de silencio.
//
// API: QREstilos.crearConLogo(texto, fraccion?)  → QRPropio.crear en nivel H, con
//          .placa = { k, a, n } (lado de la placa en módulos, primer módulo)
//      QREstilos.dibujar(canvas, qr, { lado, estiloModulo, estiloOjo, oscuro,
//          claro, logo })  → { colores: { oscuro, claro, motivo }, logo, placa }
//      QREstilos.cargarLogo(archivo)  → Promise<canvas> o error con .codigo
//      QREstilos.derivarColores(canvas) → { oscuro, claro, motivo, paleta }
//      QREstilos.asegurarContraste(oscuro, claro) y QREstilos.contraste(a, b)
(function () {
  'use strict';

  const ESTILOS_MODULO = ['cuadrado', 'redondeado', 'puntos'];
  const ESTILOS_OJO = ['cuadrado', 'redondeado', 'circulo'];
  const CONTRASTE_MIN = 4.5;      // relación de luminancia mínima entre código y fondo
  const FRACCION_LOGO = 0.22;     // ancho máximo de la placa respecto del QR
  const LADO_LOGO = 512;          // el logo se reduce a esto antes de usarse
  const LADO_PALETA = 64;         // y a esto para sacarle los colores
  const MAX_ARCHIVO = 30 * 1024 * 1024;

  function fallo(codigo) {
    const e = new Error(codigo);
    e.codigo = codigo;
    return e;
  }

  // ---------------------------------------------------------------- color

  function aRgb(c) {
    if (Array.isArray(c)) return c.slice(0, 3);
    const m = /^#?([0-9a-f]{6})$/i.exec(String(c || '').trim());
    if (!m) return [0, 0, 0];
    const n = parseInt(m[1], 16);
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  }
  function aHex(rgb) {
    return '#' + rgb.map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('');
  }
  // Luminancia relativa de WCAG (0 negro, 1 blanco).
  function luminancia(c) {
    const [r, g, b] = aRgb(c).map((v) => {
      const s = v / 255;
      return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
    });
    return 0.2126 * r + 0.7152 * g + 0.0722 * b;
  }
  function contraste(a, b) {
    const la = luminancia(a), lb = luminancia(b);
    return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
  }
  function mezclar(c, destino, t) {
    return c.map((v, i) => Math.round(v + (destino[i] - v) * t));
  }

  // Devuelve un par { oscuro, claro } (en #rrggbb) que se puede escanear.
  // motivo: null si no hubo que tocarlo; 'invertido' si el código era más claro
  // que el fondo (se intercambian); 'contraste' si faltaba contraste (se
  // oscurece el código y se aclara el fondo, de a poco y alternando, para
  // moverse lo mínimo).
  function asegurarContraste(oscuro, claro, minimo) {
    const min = minimo || CONTRASTE_MIN;
    let fg = aRgb(oscuro).map(Math.round), bg = aRgb(claro).map(Math.round), motivo = null;
    if (luminancia(fg) > luminancia(bg)) { const x = fg; fg = bg; bg = x; motivo = 'invertido'; }
    if (contraste(fg, bg) < min) {
      motivo = motivo || 'contraste';
      for (let i = 0; i < 60 && contraste(fg, bg) < min; i++) {
        fg = mezclar(fg, [0, 0, 0], 0.08);
        if (contraste(fg, bg) >= min) break;
        bg = mezclar(bg, [255, 255, 255], 0.08);
      }
      if (contraste(fg, bg) < min) { fg = [0, 0, 0]; bg = [255, 255, 255]; }
    }
    return { oscuro: aHex(fg), claro: aHex(bg), motivo, contraste: contraste(fg, bg) };
  }

  // ------------------------------------------------------------- paleta

  // Píxeles opacos de la fuente reducida a LADO_PALETA como mucho.
  function pixelesOpacos(fuente) {
    const ancho = fuente.width || fuente.naturalWidth || 1;
    const alto = fuente.height || fuente.naturalHeight || 1;
    const e = Math.min(1, LADO_PALETA / Math.max(ancho, alto));
    const w = Math.max(1, Math.round(ancho * e)), h = Math.max(1, Math.round(alto * e));
    const c = document.createElement('canvas');
    c.width = w; c.height = h;
    const cx = c.getContext('2d', { willReadFrequently: true });
    cx.drawImage(fuente, 0, 0, w, h);
    const d = cx.getImageData(0, 0, w, h).data;
    const pix = [];
    // Lo casi transparente no cuenta: su color es el del fondo del archivo.
    for (let i = 0; i < d.length; i += 4) if (d[i + 3] >= 128) pix.push([d[i], d[i + 1], d[i + 2]]);
    return pix;
  }

  // Corte por la mediana: parte la caja de colores por su canal más extendido
  // hasta tener "cuantos" cajas; cada una aporta su color promedio.
  function cortarMediana(pix, cuantos) {
    if (!pix.length) return [];
    let cajas = [pix];
    while (cajas.length < cuantos) {
      let mejor = -1, puntaje = 0, canal = 0;
      cajas.forEach((caja, i) => {
        if (caja.length < 2) return;
        for (let ch = 0; ch < 3; ch++) {
          let mn = 255, mx = 0;
          for (const p of caja) { if (p[ch] < mn) mn = p[ch]; if (p[ch] > mx) mx = p[ch]; }
          const s = (mx - mn) * Math.sqrt(caja.length);
          if (mx - mn > 8 && s > puntaje) { puntaje = s; mejor = i; canal = ch; }
        }
      });
      if (mejor < 0) break;
      const caja = cajas[mejor].slice().sort((a, b) => a[canal] - b[canal]);
      const medio = caja.length >> 1;
      cajas.splice(mejor, 1, caja.slice(0, medio), caja.slice(medio));
    }
    return cajas.map((caja) => {
      const s = [0, 0, 0];
      for (const p of caja) { s[0] += p[0]; s[1] += p[1]; s[2] += p[2]; }
      return { rgb: s.map((v) => v / caja.length), n: caja.length };
    }).sort((a, b) => b.n - a.n);
  }

  function paleta(fuente, cuantos) {
    return cortarMediana(pixelesOpacos(fuente), cuantos || 8);
  }

  // Colores del QR a partir de la paleta del logo: el color dominante con color
  // (no el blanco, el negro ni los grises) como código, y el color claro más
  // presente (o una versión muy clara del código) como fondo. Después se asegura
  // el contraste. motivo: null | 'contraste' | 'invertido' | 'sin-color'.
  function derivarColores(fuente) {
    const pal = paleta(fuente, 8);
    let acento = null, mejor = 0;
    for (const p of pal) {
      const mx = Math.max(...p.rgb), mn = Math.min(...p.rgb);
      p.croma = (mx - mn) / 255;
      p.lum = luminancia(p.rgb);
      if (p.croma < 0.18 || mx < 40) continue; // gris, casi blanco o casi negro: no es un acento
      const puntaje = p.n * (0.3 + p.croma);
      if (puntaje > mejor) { mejor = puntaje; acento = p; }
    }
    const salida = { oscuro: '#000000', claro: '#ffffff', motivo: 'sin-color', paleta: pal.map((p) => aHex(p.rgb)) };
    if (!acento) return salida;
    const claros = pal.filter((p) => p !== acento && p.lum >= 0.7).sort((a, b) => b.lum - a.lum);
    const claro = claros.length ? claros[0].rgb : mezclar(acento.rgb, [255, 255, 255], 0.9);
    const r = asegurarContraste(acento.rgb, claro, CONTRASTE_MIN);
    salida.oscuro = r.oscuro;
    salida.claro = r.claro;
    salida.motivo = r.motivo;
    return salida;
  }

  // --------------------------------------------------------------- logo

  function imagenDesdeUrl(url) {
    return new Promise((ok, mal) => {
      const img = new Image();
      img.onload = () => ok(img);
      img.onerror = () => mal(new Error('decode'));
      img.src = url;
    });
  }

  // Lee un archivo de imagen (PNG, JPG, WebP, GIF, SVG…) y lo deja en un canvas
  // de hasta LADO_LOGO de lado, con la orientación EXIF aplicada. No lanza otra
  // cosa que errores con .codigo: 'vacio' | 'grande' | 'tipo' | 'ilegible'.
  async function cargarLogo(archivo) {
    if (!archivo) throw fallo('vacio');
    if (archivo.size > MAX_ARCHIVO) throw fallo('grande');
    if (archivo.type && !/^image\//i.test(archivo.type)) throw fallo('tipo');
    let fuente = null, ancho = 0, alto = 0;
    try {
      fuente = await createImageBitmap(archivo, { imageOrientation: 'from-image' });
      ancho = fuente.width; alto = fuente.height;
    } catch (e) {
      // SVG y algún formato raro no pasan por createImageBitmap: se prueba con <img>.
      const url = URL.createObjectURL(archivo);
      try {
        fuente = await imagenDesdeUrl(url);
        ancho = fuente.naturalWidth || 300;
        alto = fuente.naturalHeight || ancho;
      } catch (e2) {
        throw fallo('ilegible');
      } finally {
        URL.revokeObjectURL(url);
      }
    }
    try {
      if (!(ancho > 0 && alto > 0)) throw fallo('ilegible');
      const e = Math.min(1, LADO_LOGO / Math.max(ancho, alto));
      const lienzo = document.createElement('canvas');
      lienzo.width = Math.max(1, Math.round(ancho * e));
      lienzo.height = Math.max(1, Math.round(alto * e));
      const cx = lienzo.getContext('2d', { willReadFrequently: true });
      cx.imageSmoothingQuality = 'high';
      cx.drawImage(fuente, 0, 0, lienzo.width, lienzo.height);
      const d = cx.getImageData(0, 0, lienzo.width, lienzo.height).data;
      let visible = false;
      for (let i = 3; i < d.length; i += 4) if (d[i] >= 16) { visible = true; break; }
      if (!visible) throw fallo('vacio');
      return lienzo;
    } catch (err) {
      throw err.codigo ? err : fallo('ilegible');
    } finally {
      if (fuente && fuente.close) fuente.close();
    }
  }

  // ------------------------------------------------------- placa del logo

  // Qué módulos son de la estructura (no se pueden tapar) en la versión v.
  const cacheReservado = {};
  function reservado(v) {
    if (cacheReservado[v]) return cacheReservado[v];
    const n = 4 * v + 17;
    const m = new Uint8Array(n * n);
    const marcar = (f0, f1, c0, c1) => {
      for (let f = Math.max(0, f0); f <= Math.min(n - 1, f1); f++) {
        for (let c = Math.max(0, c0); c <= Math.min(n - 1, c1); c++) m[f * n + c] = 1;
      }
    };
    marcar(0, 8, 0, 8); marcar(0, 8, n - 8, n - 1); marcar(n - 8, n - 1, 0, 8); // ojos, separadores y formato
    marcar(6, 6, 0, n - 1); marcar(0, n - 1, 6, 6);                             // temporización
    if (v >= 7) { marcar(0, 5, n - 11, n - 9); marcar(n - 11, n - 9, 0, 5); }   // versión
    const pos = QRPropio.posicionesAlineacion(v);
    for (const f of pos) {
      for (const c of pos) {
        if ((f === 6 && c === 6) || (f === 6 && c === n - 7) || (f === n - 7 && c === 6)) continue;
        marcar(f - 2, f + 2, c - 2, c + 2);
      }
    }
    cacheReservado[v] = m;
    return m;
  }

  // La placa más grande (lado k impar, centrada) que entra en la versión v sin
  // tocar la estructura, de hasta "fraccion" del ancho. null si no hay lugar.
  function placaPara(v, fraccion) {
    const n = 4 * v + 17;
    const f = Math.min(fraccion || FRACCION_LOGO, FRACCION_LOGO);
    let k = Math.floor(f * n);
    if (k % 2 === 0) k--;
    if (k < 3) k = 3; // un pedido muy chico en un QR chico: la placa más pequeña posible
    const minimo = Math.max(3, Math.ceil(0.55 * f * n));
    const m = reservado(v);
    for (; k >= minimo; k -= 2) {
      const a = (n - k) / 2;
      let libre = true;
      for (let fila = a; libre && fila < a + k; fila++) {
        for (let col = a; col < a + k; col++) if (m[fila * n + col]) { libre = false; break; }
      }
      if (libre) return { k, a, n };
    }
    return null;
  }

  // QR en nivel H con lugar para el logo. Si la versión más chica donde cabe el
  // texto no tiene lugar para una placa útil, sube de versión. Lanza lo mismo
  // que QRPropio.crear si el texto no cabe.
  function crearConLogo(texto, fraccion) {
    for (let v = 1; v <= 40;) {
      const qr = QRPropio.crear(texto, { nivel: 'H', version: v });
      const placa = placaPara(qr.version, fraccion);
      if (placa) { qr.nivel = 'H'; qr.placa = placa; return qr; }
      v = qr.version + 1;
    }
    throw new Error('Too long data');
  }

  // Módulos que cubre la placa: los que tocan un cuadrado de esquinas redondeadas.
  function modulosLimpios(placa) {
    const { k, a, n } = placa;
    const limpio = new Uint8Array(n * n);
    const h = k / 2, rho = 0.25 * k, dentro = h - rho;
    let cuenta = 0;
    for (let f = a; f < a + k; f++) {
      for (let c = a; c < a + k; c++) {
        const dx = Math.max(Math.abs(c + 0.5 - (a + h)) - 0.5, 0) - dentro;
        const dy = Math.max(Math.abs(f + 0.5 - (a + h)) - 0.5, 0) - dentro;
        if (Math.hypot(Math.max(dx, 0), Math.max(dy, 0)) + Math.min(Math.max(dx, dy), 0) <= rho) {
          limpio[f * n + c] = 1;
          cuenta++;
        }
      }
    }
    return { limpio, cuenta };
  }

  // ------------------------------------------------------------- dibujo

  // Rectángulo con una esquina distinta cada una (arriba-izq., arriba-der., abajo-der., abajo-izq.).
  function rectRedondo(cx, x0, y0, x1, y1, a, b, c, d) {
    cx.moveTo(x0 + a, y0);
    cx.lineTo(x1 - b, y0);
    if (b) cx.arc(x1 - b, y0 + b, b, -Math.PI / 2, 0);
    cx.lineTo(x1, y1 - c);
    if (c) cx.arc(x1 - c, y1 - c, c, 0, Math.PI / 2);
    cx.lineTo(x0 + d, y1);
    if (d) cx.arc(x0 + d, y1 - d, d, Math.PI / 2, Math.PI);
    cx.lineTo(x0, y0 + a);
    if (a) cx.arc(x0 + a, y0 + a, a, Math.PI, Math.PI * 1.5);
    cx.closePath();
  }

  // Un ojo (cuadro de posición de 7 × 7 módulos): anillo de 1 módulo, hueco de
  // 5 × 5 y centro de 3 × 3. Los estilos redondean las curvas pero no cambian
  // las proporciones 1:1:3:1:1 que los lectores buscan.
  function dibujarOjo(cx, P, estilo, f0, c0) {
    const x0 = P(c0), y0 = P(f0), x1 = P(c0 + 7), y1 = P(f0 + 7);
    const hx0 = P(c0 + 1), hy0 = P(f0 + 1), hx1 = P(c0 + 6), hy1 = P(f0 + 6);
    const kx0 = P(c0 + 2), ky0 = P(f0 + 2), kx1 = P(c0 + 5), ky1 = P(f0 + 5);
    const u = (x1 - x0) / 7;
    cx.beginPath();
    if (estilo === 'cuadrado') {
      cx.rect(x0, y0, x1 - x0, y1 - y0);
      cx.rect(hx0, hy0, hx1 - hx0, hy1 - hy0);
      cx.rect(kx0, ky0, kx1 - kx0, ky1 - ky0);
    } else {
      const r1 = (estilo === 'circulo' ? 2.2 : 2) * u;
      rectRedondo(cx, x0, y0, x1, y1, r1, r1, r1, r1);
      const r2 = r1 - u;
      rectRedondo(cx, hx0, hy0, hx1, hy1, r2, r2, r2, r2);
      if (estilo === 'circulo') {
        cx.moveTo((kx0 + kx1) / 2 + 1.5 * u, (ky0 + ky1) / 2);
        cx.arc((kx0 + kx1) / 2, (ky0 + ky1) / 2, 1.5 * u, 0, Math.PI * 2);
        cx.closePath();
      } else {
        const r3 = 0.9 * u;
        rectRedondo(cx, kx0, ky0, kx1, ky1, r3, r3, r3, r3);
      }
    }
    cx.fill('evenodd');
  }

  function dibujar(canvas, qr, opciones) {
    const op = opciones || {};
    const lado = op.lado || 300;
    const estM = ESTILOS_MODULO.indexOf(op.estiloModulo) >= 0 ? op.estiloModulo : 'cuadrado';
    const estO = ESTILOS_OJO.indexOf(op.estiloOjo) >= 0 ? op.estiloOjo : 'cuadrado';
    const col = asegurarContraste(op.oscuro || '#000000', op.claro || '#ffffff');
    // El logo solo va en un QR armado con crearConLogo (nivel H y placa libre).
    const logo = op.logo && qr.placa && qr.nivel === 'H' ? op.logo : null;
    const res = { colores: col, logo: !!logo, placa: logo ? qr.placa : null };

    if (estM === 'cuadrado' && estO === 'cuadrado' && !logo) {
      QRPropio.pintar(canvas, qr, lado, { oscuro: col.oscuro, claro: col.claro });
      return res;
    }

    const n = qr.tamano;
    const margen = Math.ceil(4 * lado / n);
    const celda = lado / n;
    canvas.width = canvas.height = lado + margen * 2;
    const cx = canvas.getContext('2d');
    cx.fillStyle = col.claro;
    cx.fillRect(0, 0, canvas.width, canvas.height);
    // Los bordes de módulo se redondean al píxel, como en QRPropio.pintar: sin huecos entre vecinos.
    const P = (i) => margen + Math.round(i * celda);

    const esOjo = (f, c) => (f < 7 && c < 7) || (f < 7 && c >= n - 7) || (f >= n - 7 && c < 7);
    const limpios = logo ? modulosLimpios(qr.placa) : null;
    if (limpios) res.placa = Object.assign({}, qr.placa, { cuenta: limpios.cuenta, fraccion: limpios.cuenta / (n * n) });
    const datos = new Uint8Array(n * n);
    for (let f = 0; f < n; f++) {
      for (let c = 0; c < n; c++) {
        if (qr.esOscuro(f, c) && !esOjo(f, c) && !(limpios && limpios.limpio[f * n + c])) datos[f * n + c] = 1;
      }
    }

    cx.fillStyle = col.oscuro;
    cx.beginPath();
    for (let f = 0; f < n; f++) {
      const y0 = P(f), y1 = P(f + 1);
      for (let c = 0; c < n; c++) {
        if (!datos[f * n + c]) continue;
        const x0 = P(c), x1 = P(c + 1);
        const t = Math.min(x1 - x0, y1 - y0);
        if (estM === 'cuadrado') {
          cx.rect(x0, y0, x1 - x0, y1 - y0);
        } else if (estM === 'redondeado') {
          const r = 0.32 * t;
          rectRedondo(cx, x0, y0, x1, y1, r, r, r, r);
        } else if (estM === 'puntos') {
          const r = 0.46 * t;
          cx.moveTo((x0 + x1) / 2 + r, (y0 + y1) / 2);
          cx.arc((x0 + x1) / 2, (y0 + y1) / 2, r, 0, Math.PI * 2);
          cx.closePath();
        }
      }
    }
    cx.fill();

    dibujarOjo(cx, P, estO, 0, 0);
    dibujarOjo(cx, P, estO, 0, n - 7);
    dibujarOjo(cx, P, estO, n - 7, 0);

    if (logo) {
      const { k, a } = qr.placa;
      const pad = 0.35 * celda;
      const bx0 = P(a) + pad, by0 = P(a) + pad, bx1 = P(a + k) - pad, by1 = P(a + k) - pad;
      const bw = bx1 - bx0, bh = by1 - by0;
      const esc = Math.min(bw / logo.width, bh / logo.height);
      const dw = logo.width * esc, dh = logo.height * esc;
      cx.save();
      cx.beginPath();
      const r = Math.max(0, 0.25 * (P(a + k) - P(a)) - pad);
      rectRedondo(cx, bx0, by0, bx1, by1, r, r, r, r);
      cx.clip();
      cx.imageSmoothingQuality = 'high';
      cx.drawImage(logo, bx0 + (bw - dw) / 2, by0 + (bh - dh) / 2, dw, dh);
      cx.restore();
    }
    return res;
  }

  globalThis.QREstilos = {
    ESTILOS_MODULO, ESTILOS_OJO, CONTRASTE_MIN,
    crearConLogo, placaPara, dibujar, cargarLogo, paleta, derivarColores, asegurarContraste, contraste
  };
})();

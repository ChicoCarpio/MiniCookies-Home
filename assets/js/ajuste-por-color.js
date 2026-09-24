/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// "Ajustar por color": tono, saturación y luminancia para cada rango de color
// (como el panel HSL de Lightroom). Compartido por el Editor de color y el
// Editor avanzado.
//
// Ajustes: objeto { rojos: { tono, saturacion, luminancia }, naranjas: {...}, ... }
// con valores de -100 a 100, o null si no se tocó nada.
const AjustePorColor = (() => {
  // "centro" es el tono (en grados) donde el rango pesa 100%; entre dos
  // centros vecinos el peso pasa de uno al otro de forma suave, así que
  // mover los 8 igual equivale a un ajuste global, sin huecos ni saltos.
  // Son los mismos colores que los filtros "Solo X" del Editor de color.
  const RANGOS = [
    { clave: 'rojos', nombre: 'Rojos', centro: 0, muestra: '#e03131' },
    { clave: 'naranjas', nombre: 'Naranjas y marrones', centro: 25, muestra: '#f08c00' },
    { clave: 'amarillos', nombre: 'Amarillos', centro: 50, muestra: '#f5d000' },
    { clave: 'verdes', nombre: 'Verdes', centro: 118, muestra: '#2f9e44' },
    { clave: 'cian', nombre: 'Cian / turquesa', centro: 180, muestra: '#15aabf' },
    { clave: 'azules', nombre: 'Azules', centro: 215, muestra: '#1c7ed6' },
    { clave: 'morados', nombre: 'Morados', centro: 275, muestra: '#7048e8' },
    { clave: 'rosas', nombre: 'Rosas / magentas', centro: 325, muestra: '#d6336c' },
  ];
  const INDICE_VERDES = RANGOS.findIndex((r) => r.clave === 'verdes');

  function neutro() {
    const ajustes = {};
    RANGOS.forEach((r) => { ajustes[r.clave] = { tono: 0, saturacion: 0, luminancia: 0 }; });
    return ajustes;
  }

  function rangoModificado(a) {
    return Boolean(a) && (a.tono !== 0 || a.saturacion !== 0 || a.luminancia !== 0);
  }

  function enUso(ajustes) {
    return Boolean(ajustes) && RANGOS.some((r) => rangoModificado(ajustes[r.clave]));
  }

  // Para cada tono entero (0-359), cuánto pesa cada rango.
  const PESOS = (() => {
    const tabla = [];
    const n = RANGOS.length;
    for (let h = 0; h < 360; h++) {
      const pesos = new Array(n).fill(0);
      for (let k = 0; k < n; k++) {
        const desde = RANGOS[k].centro;
        let hasta = RANGOS[(k + 1) % n].centro;
        if (hasta <= desde) hasta += 360;
        const hh = h < desde ? h + 360 : h;
        if (hh >= desde && hh < hasta) {
          const t = (hh - desde) / (hasta - desde);
          const suave = t * t * (3 - 2 * t);
          pesos[k] = 1 - suave;
          pesos[(k + 1) % n] = suave;
          break;
        }
      }
      tabla.push(pesos);
    }
    return tabla;
  })();

  function transicion(desde, hasta, x) {
    const t = Math.max(0, Math.min(1, (x - desde) / (hasta - desde)));
    return t * t * (3 - 2 * t);
  }

  function rgbAHsl(r, g, b) {
    r /= 255; g /= 255; b /= 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, s = 0;
    const l = (max + min) / 2;
    const d = max - min;
    if (d !== 0) {
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
      else if (max === g) h = (b - r) / d + 2;
      else h = (r - g) / d + 4;
      h *= 60;
    }
    return { h, s, l };
  }

  function hslARgb(h, s, l) {
    if (s === 0) return [l * 255, l * 255, l * 255];
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    const canal = (t) => {
      if (t < 0) t += 1;
      if (t > 1) t -= 1;
      if (t < 1 / 6) return p + (q - p) * 6 * t;
      if (t < 1 / 2) return q;
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
      return p;
    };
    const hn = h / 360;
    return [canal(hn + 1 / 3) * 255, canal(hn) * 255, canal(hn - 1 / 3) * 255];
  }

  // Devuelve el arreglo de ajustes por rango (en el orden de RANGOS) listo
  // para aplicar() en el bucle de píxeles, o null si no hay nada que hacer.
  function preparar(ajustes) {
    return enUso(ajustes) ? RANGOS.map((r) => ajustes[r.clave] || { tono: 0, saturacion: 0, luminancia: 0 }) : null;
  }

  // Cuánto pertenece un píxel a cada rango: deja el resultado en PESOS_PIXEL
  // (arreglo reutilizado, para no crear uno por píxel). Compartido por
  // aplicar() y mezclarGris().
  const PESOS_PIXEL = new Float64Array(RANGOS.length);
  const PESOS_REFERENCIA = new Float64Array(RANGOS.length);
  // conOliva: aplicar la regla del verde oliva de abajo (la mezcla B/N no la
  // usa, ver mezclarGris). destino: arreglo donde se dejan los pesos.
  function calcularPesos(hsl, conOliva = true, destino = PESOS_PIXEL) {
    // El follaje oscuro (verde oliva) se ve verde aunque su tono caiga del
    // lado del amarillo: oscuro y entre 50° y 90° cuenta como verde del todo,
    // y la corrección se desvanece hacia 35° y 105° (así los marrones
    // oscuros, 20°-35°, siguen siendo naranjas/marrones). Se mezclan los
    // pesos hacia "verdes" en vez de correr el tono: correr el tono hasta
    // 118° dentro de una franja de 15° amplificaba cualquier variación y
    // dejaba manchas en las sombras con ruido.
    let comoVerde = 0;
    if (conOliva && hsl.l < 0.4 && hsl.h > 35 && hsl.h < 105) {
      const oscuridad = 1 - transicion(0.12, 0.4, hsl.l);
      const cercania = 1 - transicion(20, 35, Math.abs(hsl.h - 70));
      comoVerde = oscuridad * cercania;
    }
    // Pesos interpolados entre los dos grados enteros vecinos (sin escalones).
    const h0 = Math.floor(hsl.h) % 360, fraccion = hsl.h - Math.floor(hsl.h);
    const pesosA = PESOS[h0], pesosB = PESOS[(h0 + 1) % 360];
    for (let k = 0; k < pesosA.length; k++) {
      const pesoTono = pesosA[k] + (pesosB[k] - pesosA[k]) * fraccion;
      destino[k] = pesoTono * (1 - comoVerde) + (k === INDICE_VERDES ? comoVerde : 0);
    }
    return destino;
  }

  // Los colores muy apagados tienen un tono poco confiable: los ajustes por
  // color entran de a poco hasta un croma de 0.14 (con 0.1 dejaba contornos).
  function fuerzaSegunCroma(croma) {
    return transicion(0.02, 0.14, croma);
  }

  // Copia desenfocada de la imagen (radio 1.5% del lado menor, proporcional:
  // la vista previa reducida y la descarga a tamaño completo quedan iguales),
  // para decidir A QUÉ COLOR pertenece cada píxel. El tono de la piel cae
  // justo entre rojos (0°) y naranjas (25°) y varía de un poro a otro (más
  // el ruido de color del JPEG): decidido píxel por píxel, con esos dos
  // rangos en valores distintos la piel quedaba llena de manchas. El ajuste
  // se sigue aplicando al píxel real (no se pierde nitidez), y en los bordes
  // entre colores distintos manda el píxel (ver pesosDeDecision), así no
  // queda halo. Devuelve los datos RGBA, o null si el lienzo no tiene nada.
  function referenciaColor(canvas) {
    const ancho = canvas.width, alto = canvas.height;
    if (!ancho || !alto) return null;
    const copia = document.createElement('canvas');
    copia.width = ancho; copia.height = alto;
    const ctx = copia.getContext('2d');
    // Sin soporte de ctx.filter (Safari viejo) queda sin desenfocar: igual
    // que antes de este arreglo, no se rompe nada.
    ctx.filter = `blur(${Math.max(0.5, Math.min(ancho, alto) * 0.015)}px)`;
    ctx.drawImage(canvas, 0, 0);
    return ctx.getImageData(0, 0, ancho, alto).data;
  }

  // Pesos y croma con los que se decide el rango de color de un píxel.
  // Mezcla los del píxel con los de la referencia desenfocada según cuánto
  // se parecen sus colores (diferencias rojo−verde y verde−azul): en una zona
  // pareja, como la piel, manda la referencia (sin manchas); en un borde entre
  // colores distintos, como sombrero contra cielo, manda el píxel (sin halo).
  // Deja el croma en cromaDecision para no crear un objeto por píxel.
  let cromaDecision = 0;
  function pesosDeDecision(r, g, b, hsl, rr, gr, br, conOliva) {
    const cromaPixel = (Math.max(r, g, b) - Math.min(r, g, b)) / 255;
    const pesos = calcularPesos(hsl, conOliva);
    if (rr === undefined || (rr === r && gr === g && br === b)) {
      cromaDecision = cromaPixel;
      return pesos;
    }
    const diferencia = Math.abs((r - g) - (rr - gr)) + Math.abs((g - b) - (gr - br));
    const confianza = 1 - transicion(6, 20, diferencia);
    if (confianza === 0) {
      cromaDecision = cromaPixel;
      return pesos;
    }
    const pesosRef = calcularPesos(rgbAHsl(rr, gr, br), conOliva, PESOS_REFERENCIA);
    for (let k = 0; k < pesos.length; k++) pesos[k] += (pesosRef[k] - pesos[k]) * confianza;
    const cromaRef = (Math.max(rr, gr, br) - Math.min(rr, gr, br)) / 255;
    cromaDecision = cromaPixel + (cromaRef - cromaPixel) * confianza;
    return pesos;
  }

  // Aplica el ajuste a un píxel (canales 0-255) y devuelve [r, g, b].
  // rr, gr, br (opcionales): el mismo píxel en referenciaColor(), que ayuda
  // a decidir el rango de color (ver pesosDeDecision).
  function aplicar(r, g, b, preparado, rr, gr, br) {
    const hsl = rgbAHsl(r, g, b);
    const pesos = pesosDeDecision(r, g, b, hsl, rr, gr, br, true);
    const croma = cromaDecision;
    if (croma <= 0.02) return [r, g, b]; // los grises no tienen color: no se tocan
    let dTono = 0, dSat = 0, dLuz = 0;
    for (let k = 0; k < pesos.length; k++) {
      const peso = pesos[k];
      if (peso === 0) continue;
      const a = preparado[k];
      dTono += peso * a.tono;
      dSat += peso * a.saturacion;
      dLuz += peso * a.luminancia;
    }
    if (dTono === 0 && dSat === 0 && dLuz === 0) return [r, g, b];
    const fuerza = fuerzaSegunCroma(croma);
    let h = hsl.h + dTono * 0.3 * fuerza; // ±100 → ±30°
    if (h < 0) h += 360; else if (h >= 360) h -= 360;
    const s = Math.max(0, Math.min(1, hsl.s * (1 + (dSat / 100) * fuerza)));
    // Luminancia: se desvanece en los casi negros al aclarar (y en los casi
    // blancos al oscurecer). Ahí el tono es muy inestable —un punto más en un
    // canal lo mueve varios grados— y sumarles luz fija hacía que un píxel
    // casi negro quedara el triple de claro que su vecino: manchas y
    // contornos en las sombras con ruido (follaje, fondos desenfocados).
    const k = (dLuz / 100) * 0.5 * fuerza;
    const l = k > 0
      ? hsl.l + (1 - hsl.l) * k * transicion(0, 0.25, hsl.l)
      : hsl.l * (1 + k * transicion(0, 0.25, 1 - hsl.l));
    return hslARgb(h, s, l);
  }

  // --- Luz por color (en B/N es la "Mezcla B/N" de Lightroom) ---
  // valores: arreglo de -100 a 100 en el orden de RANGOS (o null). Recibe el
  // píxel a color y un nivel de luz (el gris del B/N, o la luminancia del
  // píxel a color) y lo aclara u oscurece según el color: p. ej. Azules en
  // −60 oscurece el cielo. A color, los editores aplican el cociente
  // resultado/luminancia como factor a los tres canales.
  function prepararMezclaGris(valores) {
    return valores && valores.some((v) => v !== 0) ? valores : null;
  }

  // rr, gr, br (opcionales): igual que en aplicar().
  function mezclarGris(r, g, b, gris, valores, rr, gr, br) {
    // Sin la regla del verde oliva: en gris, contar el oliva oscuro como verde
    // pero el claro como amarillo levantaba las sombras más que los medios y
    // aplanaba la textura del follaje (parches planos con Verdes en +60).
    const pesos = pesosDeDecision(r, g, b, rgbAHsl(r, g, b), rr, gr, br, false);
    const croma = cromaDecision;
    if (croma <= 0.02) return gris;
    let d = 0;
    for (let k = 0; k < pesos.length; k++) d += pesos[k] * valores[k];
    if (d === 0) return gris;
    // Como la exposición: ±100 → ±1.5 pasos (×2.8 / ×0.35). Multiplicar
    // respeta la textura; "empujar hacia el blanco" levantaba mucho más los
    // oscuros que los medios, invertía el orden entre vecinos y dejaba
    // manchas planas en el follaje con Verdes en +60. Al multiplicar, los
    // casi negros (tono poco confiable) casi no se mueven.
    const pasos = (d / 100) * 1.5 * fuerzaSegunCroma(croma);
    const antes = gris / 255;
    const despues = antes * Math.pow(2, pasos);
    if (pasos <= 0) return despues * 255;
    // Al aclarar, hombro suave: arriba de 0.8 se acerca al blanco sin
    // cortarse de golpe (si no, Naranjas en +60 quemaba las luces cálidas en
    // una mancha blanca). Se suma solo lo que el hombro agrega respecto del
    // gris original, así lo que no se aclara queda exactamente igual.
    const hombro = (x) => (x <= 0.8 ? x : 0.8 + 0.2 * (1 - Math.exp(-(x - 0.8) / 0.2)));
    return Math.min(255, (antes + hombro(despues) - hombro(antes)) * 255);
  }

  // --- Panel ---------------------------------------------------------------
  const CSS = `
    .ajuste-color-subtitulo { font-size: 0.8rem; font-weight: 600; margin: 14px 0 6px; padding-top: 10px; border-top: 1px solid var(--gris-borde); }
    .ajuste-color-chips { display: grid; grid-template-columns: repeat(8, 1fr); gap: 4px; margin-bottom: 4px; }
    .ajuste-color-chip { position: relative; aspect-ratio: 1; border-radius: 50%; border: 2px solid transparent; cursor: pointer; padding: 0; box-shadow: inset 0 0 0 2px var(--gris-fondo); }
    .ajuste-color-chip.activo { border-color: var(--texto); }
    .ajuste-color-chip.modificado::after { content: ''; position: absolute; right: -2px; top: -2px; width: 8px; height: 8px; border-radius: 50%; background: var(--azul); border: 1px solid var(--gris-fondo); }
    .ajuste-color-nombre { font-size: 0.78rem; color: var(--texto-suave); margin-bottom: 6px; display: flex; justify-content: space-between; align-items: center; gap: 6px; }
    .ajuste-color-nombre button { font-size: 0.72rem; background: none; border: none; color: var(--azul); cursor: pointer; padding: 0; }
    .ajuste-color-deslizadores { display: grid; gap: 8px; }
    .ajuste-color-nota { font-size: 0.72rem; color: var(--texto-suave); margin: 6px 0 0; }`;

  // opciones.leer(): ajustes actuales (o null); opciones.escribir(ajustes):
  // guarda los ajustes nuevos (siempre un objeto completo). Devuelve
  // { refrescar() } para volver a pintar el panel si los ajustes cambian
  // desde afuera (se eligió un filtro, otra capa, se restableció...).
  function crearPanel(contenedor, opciones) {
    if (!document.getElementById('estilosAjustePorColor')) {
      const estilo = document.createElement('style');
      estilo.id = 'estilosAjustePorColor';
      estilo.textContent = CSS;
      document.head.appendChild(estilo);
    }
    contenedor.innerHTML = `
      <div class="ajuste-color-subtitulo">Ajustar por color</div>
      <div class="ajuste-color-chips" role="group" aria-label="Color a ajustar"></div>
      <div class="ajuste-color-nombre"><span class="ajuste-color-rango"></span><button type="button">Restablecer este color</button></div>
      <div class="ajuste-color-deslizadores">
        <label class="campo-deslizador"><span class="etiqueta-deslizador"><span>Tono</span><span data-valor="tono">0</span></span><input type="range" data-ajuste="tono" min="-100" max="100" value="0"></label>
        <label class="campo-deslizador"><span class="etiqueta-deslizador"><span>Saturación</span><span data-valor="saturacion">0</span></span><input type="range" data-ajuste="saturacion" min="-100" max="100" value="0"></label>
        <label class="campo-deslizador"><span class="etiqueta-deslizador"><span>Luminancia</span><span data-valor="luminancia">0</span></span><input type="range" data-ajuste="luminancia" min="-100" max="100" value="0"></label>
      </div>
      <p class="ajuste-color-nota">Cambia solo los tonos de ese color. Por ejemplo, Saturación en −100 en todos los colores menos uno deja solo ese color.</p>`;
    const chips = contenedor.querySelector('.ajuste-color-chips');
    let elegido = RANGOS[0].clave;

    const actuales = () => opciones.leer() || neutro();

    function refrescar() {
      const ajustes = actuales();
      contenedor.querySelector('.ajuste-color-rango').textContent = RANGOS.find((r) => r.clave === elegido).nombre;
      contenedor.querySelectorAll('input[data-ajuste]').forEach((input) => {
        const n = ajustes[elegido][input.dataset.ajuste];
        input.value = n;
        contenedor.querySelector(`[data-valor="${input.dataset.ajuste}"]`).textContent = `${n > 0 ? '+' : ''}${n}`;
      });
      chips.querySelectorAll('.ajuste-color-chip').forEach((chip) => {
        chip.classList.toggle('activo', chip.dataset.rango === elegido);
        chip.classList.toggle('modificado', rangoModificado(ajustes[chip.dataset.rango]));
        chip.setAttribute('aria-pressed', chip.dataset.rango === elegido ? 'true' : 'false');
      });
    }

    RANGOS.forEach((rango) => {
      const chip = document.createElement('button');
      chip.type = 'button';
      chip.className = 'ajuste-color-chip';
      chip.dataset.rango = rango.clave;
      chip.title = rango.nombre;
      chip.setAttribute('aria-label', rango.nombre);
      chip.style.background = rango.muestra;
      chip.addEventListener('click', () => { elegido = rango.clave; refrescar(); });
      chips.appendChild(chip);
    });

    contenedor.querySelectorAll('input[data-ajuste]').forEach((input) => {
      input.addEventListener('input', () => {
        const ajustes = JSON.parse(JSON.stringify(actuales()));
        ajustes[elegido][input.dataset.ajuste] = Number(input.value);
        opciones.escribir(ajustes);
        refrescar();
      });
    });

    contenedor.querySelector('.ajuste-color-nombre button').addEventListener('click', () => {
      const ajustes = JSON.parse(JSON.stringify(actuales()));
      ajustes[elegido] = { tono: 0, saturacion: 0, luminancia: 0 };
      opciones.escribir(ajustes);
      refrescar();
    });

    refrescar();
    return { refrescar };
  }

  return { RANGOS, neutro, enUso, preparar, referenciaColor, aplicar, prepararMezclaGris, mezclarGris, crearPanel };
})();

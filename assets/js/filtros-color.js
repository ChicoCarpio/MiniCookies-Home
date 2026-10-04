/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Filtros del Editor de color (los "looks" con miniatura), compartidos con
// el Editor avanzado, que los aplica sobre la capa activa.
//
// Cada filtro es { cambios, capas, textura }:
// - cambios: valores de los deslizadores que se diferencian del neutro (cada
//   editor los completa con sus propios AJUSTES_NEUTROS);
// - capas: capas extra con máscara propia (viñeta, sombras, luces, un tono)
//   para lo que un solo juego de ajustes sobre toda la imagen no logra;
// - textura: { tipo, intensidad, orden } o null.
// Necesita assets/js/ajuste-por-color.js (porColor / apagarMenos).
const FiltrosColor = (() => {
  function filtro(cambios, extra = {}) {
    return { cambios, capas: extra.capas || [], textura: extra.textura || null };
  }

  // Capas extra: misma idea que las capas del Editor avanzado (ajustes propios
  // + una máscara que limita dónde se aplican), pero fijas dentro del filtro,
  // sin controles nuevos. Sirven para lo que un solo juego de ajustes sobre
  // toda la imagen no logra: virar distinto sombras y luces, una viñeta, dejar
  // a color solo un tono, o teñir una imagen que ya pasó a blanco y negro.
  const vineta = (fuerza, interior = 0.45) => ({ mascara: { tipo: 'vineta', interior }, ajustes: { exposicion: -fuerza } });
  const enSombras = (ajustes) => ({ mascara: { tipo: 'sombras' }, ajustes });
  const enLuces = (ajustes) => ({ mascara: { tipo: 'luces' }, ajustes });
  const encimaDeTodo = (ajustes) => ({ mascara: null, ajustes });
  // extra (opcional): { tonoOscuro, rangoOscuro, luzMinima } — ver generarMascara.
  const segunTono = (tono, rango, ajustes, invertir = false, extra = {}) => ({ mascara: { tipo: 'tono', tono, rango, invertir, ...extra }, ajustes });
  // "Solo X a color": pasa a blanco y negro todo MENOS ese tono.
  const soloTono = (tono, rango, extra) => segunTono(tono, rango, { bn: true }, true, extra);
  // El follaje con poca luz, a contraluz o con luz cálida de atardecer se ve
  // verde, pero su tono cae entre 50° y 80° (verde oliva, del lado del
  // amarillo): los filtros de verdes corren su centro hacia ese tono en los
  // píxeles oscuros, y los de amarillos no toman esos píxeles oscuros.
  const VERDES_OSCUROS = { tonoOscuro: 70, rangoOscuro: 34 };
  const AMARILLOS_CLAROS = { luzMinima: 0.2 };

  // "Ajustar por color" de un filtro: solo se declaran los rangos y valores que
  // cambian, el resto queda en 0. A diferencia de las capas extra, esto se ve
  // (y se puede retocar) en el panel "Ajustar por color" al elegir el filtro.
  // (Los filtros B/N no lo usan: para aclarar u oscurecer cada color en gris
  // está "Luz por color", mezclaRojos...)
  function porColor(cambios) {
    const ajustes = AjustePorColor.neutro();
    Object.entries(cambios).forEach(([rango, valores]) => Object.assign(ajustes[rango], valores));
    return ajustes;
  }
  // Baja la saturación de todos los rangos menos los indicados (que pueden
  // traer su propio ajuste): el resto de la imagen queda apagado, no gris.
  function apagarMenos(cuanto, destacados) {
    const cambios = {};
    AjustePorColor.RANGOS.forEach((r) => { cambios[r.clave] = destacados[r.clave] || { saturacion: -cuanto }; });
    return porColor(cambios);
  }

  // Calentar/enfriar se hace con el balance de rojo/verde/azul, no con "Girar
  // tono": girar el tono rota TODOS los colores (el pasto deja de ser verde),
  // no le da una dominante cálida o fría a la imagen.
  const FILTROS = {
    ninguno: filtro({}),

    // --- Blanco y negro (la mezcla de canales imita filtros de color de lente) ---
    grises: filtro({ bn: true, contraste: 10 }),
    bnClasico: filtro({ bn: true, contraste: 15, claridad: 10, grano: 10, canalRojo: 110, canalAzul: 85 }),
    bnCine: filtro({ bn: true, contraste: 25, sombras: 10, reflejos: -10, claridad: 15, canalRojo: 130, canalAzul: 70 }, { capas: [vineta(30)] }),
    bnAltaClave: filtro({ bn: true, exposicion: 25, contraste: -15, sombras: 25, reflejos: 5, canalRojo: 120 }),
    bnBajaClave: filtro({ bn: true, exposicion: -20, contraste: 35, sombras: -15, claridad: 15 }, { capas: [vineta(35)] }),
    // Filtro rojo: oscurece cielos azules y aclara pieles
    bnAltoContraste: filtro({ bn: true, contraste: 45, claridad: 25, canalRojo: 150, canalVerde: 90, canalAzul: 40,
      mezclaAzules: -30, mezclaCian: -20 }),
    bnSuave: filtro({ bn: true, contraste: -15, exposicion: 5, claridad: -15, dispersionIntensidad: 20, dispersionTamano: 70 }),
    // Filtro verde: aclara el follaje y oscurece pieles/labios
    bnFiltroVerde: filtro({ bn: true, contraste: 15, claridad: 10, canalRojo: 70, canalVerde: 160, canalAzul: 60,
      mezclaVerdes: 25, mezclaAmarillos: 10 }),
    // Piel clara y pareja: naranjas/rojos (tonos de piel) más claros antes de pasar a gris
    bnRetrato: filtro({ bn: true, contraste: 10, sombras: 10, claridad: -10,
      mezclaNaranjas: 30, mezclaRojos: 15, mezclaAmarillos: 5 }, { capas: [vineta(15, 0.55)] }),
    // Cielo casi negro con nubes blancas, sin tocar tanto el resto como el filtro rojo
    bnCieloDramatico: filtro({ bn: true, contraste: 25, claridad: 20, reflejos: -10,
      mezclaAzules: -60, mezclaCian: -40, mezclaVerdes: 10, mezclaNaranjas: 5 }),
    // Infrarrojo B/N: follaje casi blanco, cielo casi negro, halo suave
    infrarrojo: filtro({ bn: true, contraste: 25, grano: 10, canalRojo: 80, canalVerde: 200, canalAzul: 0, dispersionIntensidad: 30, dispersionTamano: 70 }, {
      capas: [segunTono(105, 55, { exposicion: 70 })],
    }),
    selenio: filtro({ bn: true, contraste: 20, claridad: 10 }, { capas: [enSombras({ colorRojo: 10, colorVerde: -4, colorAzul: 14 })] }),
    cianotipia: filtro({ bn: true, contraste: 10, exposicion: 5 }, { capas: [encimaDeTodo({ colorRojo: -60, colorVerde: -15, colorAzul: 35 })] }),
    rayosX: filtro({ bn: true, invertir: 100, contraste: 25, claridad: 20 }, { capas: [encimaDeTodo({ colorRojo: -20, colorVerde: 5, colorAzul: 25 })] }),

    // --- Sepia y película ---
    sepia: filtro({ sepia: 100, contraste: 5 }),
    sepiaSuave: filtro({ sepia: 50, contraste: -5, exposicion: 3 }),
    sepiaIntenso: filtro({ sepia: 100, contraste: 20, grano: 15 }, { capas: [vineta(30)] }),
    vintage: filtro({ exposicion: 5, contraste: -10, sombras: 15, reflejos: -10, saturacion: -20, grano: 15, sepia: 30 }, { capas: [vineta(20)] }),
    retro70: filtro({ sepia: 20, saturacion: -15, contraste: -15, sombras: 20, grano: 20, exposicion: 5, colorRojo: 10, colorAzul: -12 }),
    retro80: filtro({ saturacion: 30, contraste: 15, grano: 8, dispersionIntensidad: 20, dispersionMatiz: -20 }, {
      capas: [enSombras({ colorRojo: 6, colorAzul: 18 }), enLuces({ colorRojo: 14, colorVerde: -4, colorAzul: 6 })],
    }),
    polaroid: filtro({ sepia: 15, contraste: -10, saturacion: -15, sombras: 20, grano: 15, exposicion: 8, colorVerde: 5, colorAzul: -5 }, {
      capas: [enSombras({ colorVerde: 6, colorAzul: 12 })],
    }),
    lomografia: filtro({ contraste: 30, saturacion: 25, grano: 20, sombras: -15 }, {
      capas: [enLuces({ colorRojo: 8, colorVerde: 8, colorAzul: -10 }), vineta(45, 0.35)],
    }),
    peliculaAntigua: filtro({ sepia: 35, grano: 40, contraste: -5, exposicion: -5, sombras: 15, claridad: -10 }, { capas: [vineta(35)] }),
    // Película de diapositiva: colores densos y cálidos, negros profundos
    diapositiva: filtro({ saturacion: 20, intensidad: 10, contraste: 20, sombras: -5, colorRojo: 8, colorAzul: -6 }),
    // Negativo de tonos fríos: sombras verdosas/cian, luces apenas cálidas
    peliculaFria: filtro({ saturacion: 10, intensidad: 10, contraste: 10 }, {
      capas: [enSombras({ colorRojo: -4, colorVerde: 6, colorAzul: 8 }), enLuces({ colorRojo: 4, colorVerde: 2 })],
    }),
    // Revelado cruzado: sombras azules, luces amarillas, mucho contraste
    procesoCruzado: filtro({ contraste: 30, saturacion: 30 }, {
      capas: [enSombras({ colorRojo: -10, colorAzul: 25 }), enLuces({ colorRojo: 10, colorVerde: 15, colorAzul: -25 })],
    }),

    // --- Tonos de color / clima ---
    frio: filtro({ colorRojo: -12, colorVerde: 2, colorAzul: 15, saturacion: -5, contraste: 10 }),
    calido: filtro({ colorRojo: 15, colorVerde: 5, colorAzul: -15, saturacion: 8, exposicion: 5 }),
    veranoDorado: filtro({ colorRojo: 18, colorVerde: 8, colorAzul: -20, saturacion: 15, exposicion: 8, dispersionIntensidad: 20, dispersionTamano: 70 }),
    oceano: filtro({ colorRojo: -15, colorVerde: 5, colorAzul: 18, saturacion: 15, contraste: 10, claridad: 10,
      hsl: porColor({ cian: { saturacion: 25 }, azules: { tono: -20, saturacion: 15 } }) }),
    atardecer: filtro({ colorRojo: 20, colorAzul: -10, saturacion: 20, reflejos: -15, contraste: 10,
      hsl: porColor({ rojos: { saturacion: 15 }, naranjas: { saturacion: 25 }, amarillos: { tono: -25, saturacion: 15 } }) }, {
      capas: [enLuces({ colorRojo: 10, colorVerde: 5, colorAzul: -15 }), enSombras({ colorRojo: 6, colorAzul: 14 })],
    }),
    amanecer: filtro({ exposicion: 10, colorRojo: 12, colorVerde: 3, colorAzul: 5, saturacion: -5, contraste: -5, claridad: -5, dispersionIntensidad: 20, dispersionTamano: 65 }),
    lluvioso: filtro({ saturacion: -30, colorRojo: -6, colorAzul: 8, contraste: -5, exposicion: -5, claridad: -10 }),
    artico: filtro({ saturacion: -25, colorRojo: -12, colorVerde: 4, colorAzul: 18, exposicion: 12, contraste: 10, reflejos: -10 }),
    plateado: filtro({ saturacion: -45, colorAzul: 6, contraste: 15, claridad: 10, reflejos: 10 }),
    // Los verdes giran hacia amarillo/naranja (la capa gira más de lo que
    // permite "Ajustar por color", ±30°), y amarillos/naranjas se encienden
    otono: filtro({ colorRojo: 12, colorVerde: 4, colorAzul: -15, saturacion: 5, contraste: 10,
      hsl: porColor({ amarillos: { tono: -35, saturacion: 25 }, naranjas: { saturacion: 30 }, rojos: { saturacion: 15 }, azules: { saturacion: -20 } }) }, {
      capas: [segunTono(110, 50, { tono: -45, saturacion: 15 })],
    }),
    primavera: filtro({ saturacion: 10, intensidad: 20, exposicion: 8, contraste: -5, colorVerde: 3,
      hsl: porColor({ verdes: { tono: 15, saturacion: 15, luminancia: 10 }, amarillos: { tono: 15, luminancia: 10 }, rosas: { saturacion: 25, luminancia: 10 } }) }),
    niebla: filtro({ exposicion: 10, contraste: -35, saturacion: -20, sombras: 30, claridad: -25, dispersionIntensidad: 25, dispersionTamano: 90 }),

    // --- Paisaje y naturaleza (casi todo con "Ajustar por color") ---
    // Solo los azules del cielo se oscurecen y saturan
    cieloIntenso: filtro({ contraste: 5, intensidad: 10,
      hsl: porColor({ azules: { saturacion: 30, luminancia: -30 }, cian: { saturacion: 20, luminancia: -15 } }) }),
    // Agua turquesa: los azules giran hacia el cian y el cian se enciende
    marCaribe: filtro({ exposicion: 5, contraste: 5, intensidad: 10,
      hsl: porColor({ cian: { saturacion: 45, luminancia: 10 }, azules: { tono: -50, saturacion: 25, luminancia: 5 }, verdes: { tono: 30 } }) }),
    // Verdes (y los verde oliva) giran hacia un verde más frío y profundo
    verdesEsmeralda: filtro({ contraste: 5,
      hsl: porColor({ verdes: { tono: 60, saturacion: 10, luminancia: -15 }, amarillos: { saturacion: -10 } }) }),
    // Follaje apagado y oscuro, look "moody" de bosque
    bosqueSombrio: filtro({ contraste: 15, sombras: -10, reflejos: -10, colorAzul: 4,
      hsl: porColor({ verdes: { tono: 20, saturacion: -40, luminancia: -30 }, amarillos: { saturacion: -30, luminancia: -20 }, cian: { saturacion: -20 } }) }, { capas: [vineta(20)] }),
    // Flores que resaltan sobre un follaje más quieto
    floresVivas: filtro({ contraste: 5,
      hsl: porColor({ rojos: { saturacion: 30 }, rosas: { saturacion: 35 }, morados: { saturacion: 35 }, amarillos: { saturacion: 15 }, verdes: { saturacion: -15, luminancia: -10 } }) }),
    // Paleta terrosa: fríos casi apagados, verdes hacia oliva, naranjas cálidos
    tonosTierra: filtro({ contraste: -5, sombras: 10,
      hsl: porColor({ azules: { saturacion: -50 }, cian: { saturacion: -50 }, morados: { saturacion: -40 }, verdes: { tono: -40, saturacion: -35 }, naranjas: { saturacion: 10 } }) }),
    // Paisaje seco: sin verdes vivos, cálido y con cielo pálido
    desierto: filtro({ contraste: 10, exposicion: 5, colorRojo: 8, colorAzul: -8,
      hsl: porColor({ verdes: { tono: -60, saturacion: -50 }, amarillos: { tono: -20, saturacion: 10 }, naranjas: { saturacion: 20, luminancia: 5 }, azules: { saturacion: -25, luminancia: 15 } }) }),

    // --- Color selectivo (misma idea que "Selección por color" del Editor avanzado) ---
    // Los rangos se continúan uno con otro alrededor del círculo de tonos, para
    // que ningún color caiga en un hueco (antes los naranjas y marrones —piel,
    // pelaje, madera— no entraban en ninguno).
    soloRojos: filtro({}, { capas: [soloTono(0, 18)] }),
    // Marrones = naranjas oscuros: sin mínimo de luz, a diferencia de los amarillos.
    soloNaranjas: filtro({}, { capas: [soloTono(25, 16)] }),
    soloAmarillos: filtro({}, { capas: [soloTono(50, 22, AMARILLOS_CLAROS)] }),
    soloVerdes: filtro({}, { capas: [soloTono(118, 55, VERDES_OSCUROS)] }),
    soloCian: filtro({}, { capas: [soloTono(180, 25)] }),
    soloAzules: filtro({}, { capas: [soloTono(215, 35)] }),
    soloMorados: filtro({}, { capas: [soloTono(275, 30)] }),
    soloRosas: filtro({}, { capas: [soloTono(325, 28)] }),
    // Acentos: a diferencia de "Solo X", el resto no pasa a gris sino que queda
    // apagado, y el ajuste se puede retocar en "Ajustar por color".
    acentoRojo: filtro({ contraste: 10, hsl: apagarMenos(75, { rojos: { saturacion: 25 }, rosas: { saturacion: -30 }, naranjas: { saturacion: -40 } }) }),
    acentoAzul: filtro({ contraste: 10, hsl: apagarMenos(75, { azules: { saturacion: 30 }, cian: { saturacion: -20 } }) }),
    acentoCalido: filtro({ contraste: 5, hsl: apagarMenos(80, { rojos: { saturacion: 15 }, naranjas: { saturacion: 15 }, amarillos: { saturacion: 10 }, rosas: { saturacion: -30 } }) }),

    // --- Retrato y edición dramática ---
    technicolor: filtro({ saturacion: 40, contraste: 20, intensidad: 15 }, { capas: [enSombras({ colorVerde: 4, colorAzul: 10 })] }),
    dramatico: filtro({ exposicion: -5, contraste: 40, sombras: -15, reflejos: -15, saturacion: -10, intensidad: 20, claridad: 35 }, { capas: [vineta(30)] }),
    retrato: filtro({ exposicion: 5, contraste: 5, sombras: 15, reflejos: -10, intensidad: 10, claridad: -15, colorRojo: 3, colorAzul: -3,
      hsl: porColor({ naranjas: { saturacion: -10, luminancia: 15 }, rojos: { saturacion: -10, luminancia: 5 } }) }, { capas: [vineta(12, 0.55)] }),
    // Piel más clara y pareja (naranjas/rojos), el resto casi sin tocar
    pielLuminosa: filtro({ exposicion: 3, sombras: 10, claridad: -20, dispersionIntensidad: 15, dispersionTamano: 70,
      hsl: porColor({ naranjas: { tono: 5, saturacion: -15, luminancia: 30 }, rojos: { saturacion: -5, luminancia: 5 }, amarillos: { saturacion: -10 } }) }),
    // "Teal & orange" de cine: sombras verde azuladas, luces (pieles) anaranjadas
    cine: filtro({ contraste: 20, saturacion: 5, sombras: -10, claridad: 10,
      hsl: porColor({ azules: { tono: -40 }, verdes: { tono: 50, saturacion: -30 }, amarillos: { tono: -20, saturacion: -15 }, naranjas: { saturacion: 15, luminancia: 5 } }) }, {
      capas: [enSombras({ colorRojo: -18, colorVerde: 8, colorAzul: 28 }), enLuces({ colorRojo: 22, colorVerde: 6, colorAzul: -20 }), vineta(20)],
    }),
    altaModa: filtro({ contraste: 35, claridad: 25, saturacion: -15, reflejos: 10 }, { capas: [enSombras({ colorAzul: 6 })] }),
    pastel: filtro({ saturacion: -30, exposicion: 12, contraste: -20, sombras: 20, claridad: -10, colorRojo: 5, colorAzul: 5 }),
    desvanecido: filtro({ exposicion: 10, contraste: -25, sombras: 25, saturacion: -20 }),
    altoBrillo: filtro({ exposicion: 20, sombras: 20, reflejos: -5, contraste: -10, saturacion: 5 }),
    sombrasProfundas: filtro({ sombras: -30, reflejos: 10, contraste: 25, claridad: 20 }, { capas: [vineta(25)] }),
    // "Bleach bypass": poco color, mucho contraste y textura
    blanqueado: filtro({ saturacion: -50, contraste: 40, claridad: 30, reflejos: 10 }),
    // Negros lavados, look mate de revista
    mate: filtro({ sombras: 30, contraste: -10, saturacion: -10, colorAzul: 4 }),

    // --- Luz, neón y creativos ---
    resplandor: filtro({ exposicion: 5, reflejos: -10, saturacion: 5, intensidad: 10, dispersionIntensidad: 55, dispersionMatiz: 15, dispersionTamano: 65 }),
    neonRosa: filtro({ colorRojo: 20, colorVerde: -15, colorAzul: 20, saturacion: 40, contraste: 15, dispersionIntensidad: 40, dispersionMatiz: -30, dispersionTamano: 55 }),
    neonAzul: filtro({ colorRojo: -15, colorVerde: 5, colorAzul: 25, saturacion: 45, contraste: 15, dispersionIntensidad: 35, dispersionMatiz: 60, dispersionTamano: 60 }),
    cyberpunk: filtro({ contraste: 30, saturacion: 35, claridad: 15, dispersionIntensidad: 40, dispersionMatiz: -40, dispersionTamano: 70, aberracionIntensidad: 20 }, {
      capas: [enSombras({ colorRojo: -10, colorVerde: 10, colorAzul: 25 }), enLuces({ colorRojo: 25, colorVerde: -10, colorAzul: 15 })],
    }),
    glitch: filtro({ contraste: 25, saturacion: 35, grano: 20, aberracionIntensidad: 75 }),
    glamour: filtro({ exposicion: 5, claridad: -20, saturacion: -5, colorRojo: 5, dispersionIntensidad: 25, dispersionTamano: 70 }),
    sueno: filtro({ exposicion: 10, contraste: -15, saturacion: -10, dispersionIntensidad: 35, dispersionTamano: 80, claridad: -15 }),
    fantasma: filtro({ exposicion: 25, saturacion: -60, contraste: -10, colorAzul: 10, dispersionIntensidad: 30, dispersionTamano: 80 }, {
      textura: { tipo: 'direccional', intensidad: 12, orden: 'debajo' },
    }),
    terror: filtro({ saturacion: -40, colorRojo: -5, colorVerde: 12, contraste: 30, exposicion: -15, sombras: -15, grano: 15 }, { capas: [vineta(50, 0.35)] }),
    // "Noche americana": oscuro, azulado y con viñeta
    nocturno: filtro({ exposicion: -15, contraste: 20, saturacion: -20, colorRojo: -10, colorAzul: 18 }, { capas: [vineta(25)] }),
    vaporwave: filtro({ saturacion: 30, contraste: 10, dispersionIntensidad: 25, dispersionTamano: 65 }, {
      capas: [enSombras({ colorRojo: 10, colorAzul: 30 }), enLuces({ colorRojo: 25, colorVerde: -10, colorAzul: 20 })],
    }),

    // --- Artístico y otros ---
    acuarela: filtro({ saturacion: -15, contraste: -10, claridad: -20, desenfoque: 1, exposicion: 5 }),
    oleo: filtro({ contraste: 10, claridad: -25, saturacion: 10 }),
    invertir: filtro({ invertir: 100 }),
    neon: filtro({ saturacion: 60, intensidad: 40, contraste: 20, dispersionIntensidad: 30, dispersionMatiz: 40, dispersionTamano: 60 }),
  };

  // --- Version en ingles de los nombres de la grilla ---------------------
  // Viven aca y no en i18nComun porque solo los usan los dos editores, y
  // idioma.js se carga en las 119 paginas del sitio.
  const NOMBRES_EN = {
    ninguno: "No filter",
    grises: "Black and white",
    bnClasico: "B&W classic",
    bnCine: "B&W cinematic",
    bnAltaClave: "B&W high key",
    bnBajaClave: "B&W low key",
    bnAltoContraste: "B&W contrast",
    bnSuave: "B&W soft",
    bnFiltroVerde: "B&W green filter",
    bnRetrato: "B&W portrait",
    bnCieloDramatico: "B&W dramatic sky",
    infrarrojo: "Infrared",
    selenio: "Selenium",
    cianotipia: "Cyanotype",
    rayosX: "X-ray",
    sepia: "Sepia",
    sepiaSuave: "Soft sepia",
    sepiaIntenso: "Deep sepia",
    vintage: "Vintage",
    retro70: "Retro 70s",
    retro80: "Retro 80s",
    polaroid: "Polaroid",
    lomografia: "Lomography",
    peliculaAntigua: "Old film",
    diapositiva: "Slide film",
    peliculaFria: "Cool film",
    procesoCruzado: "Cross process",
    frio: "Cool",
    calido: "Warm",
    veranoDorado: "Golden summer",
    oceano: "Ocean",
    atardecer: "Sunset",
    amanecer: "Sunrise",
    lluvioso: "Rainy",
    artico: "Arctic",
    plateado: "Silver",
    otono: "Autumn",
    primavera: "Spring",
    niebla: "Fog",
    cieloIntenso: "Deep sky",
    marCaribe: "Caribbean sea",
    verdesEsmeralda: "Emerald greens",
    bosqueSombrio: "Dark forest",
    floresVivas: "Vivid flowers",
    tonosTierra: "Earth tones",
    desierto: "Desert",
    soloRojos: "Reds only",
    soloNaranjas: "Oranges and browns only",
    soloAmarillos: "Yellows only",
    soloVerdes: "Greens only",
    soloCian: "Cyan / teal only",
    soloAzules: "Blues only",
    soloMorados: "Purples only",
    soloRosas: "Pinks / magentas only",
    acentoRojo: "Red accent",
    acentoAzul: "Blue accent",
    acentoCalido: "Warm accent",
    technicolor: "Technicolor",
    dramatico: "Dramatic",
    retrato: "Portrait",
    pielLuminosa: "Glowing skin",
    cine: "Cinematic",
    altaModa: "High fashion",
    pastel: "Pastel",
    desvanecido: "Faded",
    altoBrillo: "High key",
    sombrasProfundas: "Shadows",
    blanqueado: "Bleached",
    mate: "Matte",
    resplandor: "Glow",
    neonRosa: "Pink neon",
    neonAzul: "Blue neon",
    cyberpunk: "Cyberpunk",
    glitch: "Glitch",
    glamour: "Glamour",
    sueno: "Dream",
    fantasma: "Ghost",
    terror: "Horror",
    nocturno: "Night",
    vaporwave: "Vaporwave",
    acuarela: "Watercolor",
    oleo: "Oil paint",
    invertir: "Invert",
    neon: "Neon",
  };
  const GRUPOS_EN = ["Black and white", "Sepia and film", "Color tones / weather", "Landscape and nature", "Selective color", "Portrait and dramatic looks", "Light, neon and creative", "Artistic and others"];
  const enIngles = () => (window.idiomaActual ? idiomaActual() === 'en' : false);

  // Grupos y nombres de la grilla de miniaturas ("Sin filtro" va primero, fuera de los grupos).
  const GRUPOS = [
    { titulo: "Blanco y negro", filtros: [
      ["grises", "Blanco y negro"],
      ["bnClasico", "B/N clásico"],
      ["bnCine", "B/N cine"],
      ["bnAltaClave", "B/N alta clave"],
      ["bnBajaClave", "B/N baja clave"],
      ["bnAltoContraste", "B/N contraste"],
      ["bnSuave", "B/N suave"],
      ["bnFiltroVerde", "B/N filtro verde"],
      ["bnRetrato", "B/N retrato"],
      ["bnCieloDramatico", "B/N cielo dramático"],
      ["infrarrojo", "Infrarrojo"],
      ["selenio", "Selenio"],
      ["cianotipia", "Cianotipia"],
      ["rayosX", "Rayos X"],
    ] },
    { titulo: "Sepia y película", filtros: [
      ["sepia", "Sepia"],
      ["sepiaSuave", "Sepia suave"],
      ["sepiaIntenso", "Sepia intenso"],
      ["vintage", "Vintage"],
      ["retro70", "Retro 70s"],
      ["retro80", "Retro 80s"],
      ["polaroid", "Polaroid"],
      ["lomografia", "Lomografía"],
      ["peliculaAntigua", "Película antigua"],
      ["diapositiva", "Diapositiva"],
      ["peliculaFria", "Película fría"],
      ["procesoCruzado", "Proceso cruzado"],
    ] },
    { titulo: "Tonos de color / clima", filtros: [
      ["frio", "Frío"],
      ["calido", "Cálido"],
      ["veranoDorado", "Verano dorado"],
      ["oceano", "Océano"],
      ["atardecer", "Atardecer"],
      ["amanecer", "Amanecer"],
      ["lluvioso", "Lluvioso"],
      ["artico", "Ártico"],
      ["plateado", "Plateado"],
      ["otono", "Otoño"],
      ["primavera", "Primavera"],
      ["niebla", "Niebla"],
    ] },
    { titulo: "Paisaje y naturaleza", filtros: [
      ["cieloIntenso", "Cielo intenso"],
      ["marCaribe", "Mar Caribe"],
      ["verdesEsmeralda", "Verdes esmeralda"],
      ["bosqueSombrio", "Bosque sombrío"],
      ["floresVivas", "Flores vivas"],
      ["tonosTierra", "Tonos tierra"],
      ["desierto", "Desierto"],
    ] },
    { titulo: "Color selectivo", filtros: [
      ["soloRojos", "Solo rojos"],
      ["soloNaranjas", "Solo naranjas y marrones"],
      ["soloAmarillos", "Solo amarillos"],
      ["soloVerdes", "Solo verdes"],
      ["soloCian", "Solo cian / turquesa"],
      ["soloAzules", "Solo azules"],
      ["soloMorados", "Solo morados"],
      ["soloRosas", "Solo rosas / magentas"],
      ["acentoRojo", "Acento rojo"],
      ["acentoAzul", "Acento azul"],
      ["acentoCalido", "Acento cálido"],
    ] },
    { titulo: "Retrato y edición dramática", filtros: [
      ["technicolor", "Technicolor"],
      ["dramatico", "Dramático"],
      ["retrato", "Retrato"],
      ["pielLuminosa", "Piel luminosa"],
      ["cine", "Cine"],
      ["altaModa", "Alta moda"],
      ["pastel", "Pastel"],
      ["desvanecido", "Desvanecido"],
      ["altoBrillo", "Alto brillo"],
      ["sombrasProfundas", "Sombras"],
      ["blanqueado", "Blanqueado"],
      ["mate", "Mate"],
    ] },
    { titulo: "Luz, neón y creativos", filtros: [
      ["resplandor", "Resplandor"],
      ["neonRosa", "Neón rosa"],
      ["neonAzul", "Neón azul"],
      ["cyberpunk", "Cyberpunk"],
      ["glitch", "Glitch"],
      ["glamour", "Glamour"],
      ["sueno", "Sueño"],
      ["fantasma", "Fantasma"],
      ["terror", "Terror"],
      ["nocturno", "Nocturno"],
      ["vaporwave", "Vaporwave"],
    ] },
    { titulo: "Artístico y otros", filtros: [
      ["acuarela", "Acuarela"],
      ["oleo", "Óleo"],
      ["invertir", "Invertir"],
      ["neon", "Neón"],
    ] },
  ];

  // ===== Máscaras de las capas extra =====
  // Devuelven un canvas blanco cuya transparencia marca dónde se aplica la
  // capa (misma convención que las máscaras del Editor avanzado). Se calculan
  // sobre la imagen SIN ajustes, así la zona elegida no depende de lo que
  // hagan los deslizadores.
  function rgbAHsl(r, g, b) {
    r /= 255; g /= 255; b /= 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, sat = 0; const l = (max + min) / 2;
    const d = max - min;
    if (d !== 0) {
      sat = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      if (max === r) h = ((g - b) / d + (g < b ? 6 : 0));
      else if (max === g) h = (b - r) / d + 2;
      else h = (r - g) / d + 4;
      h *= 60;
    }
    return { h, s: sat, l };
  }

  // Transición suave entre 0 y 1 (sin bordes duros en las máscaras).
  function transicion(desde, hasta, x) {
    const t = Math.max(0, Math.min(1, (x - desde) / (hasta - desde)));
    return t * t * (3 - 2 * t);
  }

  // Cuánto se aplica la capa en cada píxel (0-255, uno por píxel), calculado
  // sobre la imagen sin ajustes.
  function valoresMascara(canvasOrigen, mascara) {
    const ancho = canvasOrigen.width, alto = canvasOrigen.height;
    const d = canvasOrigen.getContext('2d').getImageData(0, 0, ancho, alto).data;
    const valores = new Uint8ClampedArray(ancho * alto);
    const cx = ancho / 2, cy = alto / 2, radioMaximo = Math.hypot(cx, cy);
    for (let y = 0; y < alto; y++) {
      for (let x = 0; x < ancho; x++) {
        const i = (y * ancho + x) * 4;
        let alfa;
        if (mascara.tipo === 'vineta') {
          alfa = transicion(mascara.interior, 1.05, Math.hypot(x - cx, y - cy) / radioMaximo);
        } else if (mascara.tipo === 'sombras' || mascara.tipo === 'luces') {
          const lum = (0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2]) / 255;
          alfa = mascara.tipo === 'sombras' ? 1 - transicion(0.2, 0.65, lum) : transicion(0.35, 0.8, lum);
        } else { // tono
          const hsl = rgbAHsl(d[i], d[i + 1], d[i + 2]);
          // tonoOscuro / rangoOscuro: en los píxeles oscuros el centro (y el
          // ancho del rango) se corren hacia esos valores (ver VERDES_OSCUROS;
          // el rango se angosta para no llegar a los marrones oscuros).
          const oscuridad = 1 - transicion(0.12, 0.4, hsl.l);
          let centro = mascara.tono;
          let rango = mascara.rango;
          if (mascara.tonoOscuro !== undefined) centro += (mascara.tonoOscuro - mascara.tono) * oscuridad;
          if (mascara.rangoOscuro !== undefined) rango += (mascara.rangoOscuro - mascara.rango) * oscuridad;
          let dif = Math.abs(hsl.h - centro);
          if (dif > 180) dif = 360 - dif;
          // Los grises y los casi negros/blancos no tienen un tono confiable.
          // Se mide el croma (diferencia entre el canal mayor y el menor) y no
          // la saturación HSL: así los colores oscuros pero con color real
          // (follaje en sombra) no quedan afuera.
          const croma = (Math.max(d[i], d[i + 1], d[i + 2]) - Math.min(d[i], d[i + 1], d[i + 2])) / 255;
          alfa = (1 - transicion(rango * 0.6, rango, dif))
            * transicion(0.03, 0.09, croma) * transicion(0.02, 0.07, hsl.l) * (1 - transicion(0.9, 0.98, hsl.l));
          if (mascara.luzMinima !== undefined) alfa *= transicion(mascara.luzMinima - 0.08, mascara.luzMinima + 0.08, hsl.l);
        }
        if (mascara.invertir) alfa = 1 - alfa;
        valores[y * ancho + x] = Math.round(alfa * 255);
      }
    }
    return valores;
  }

  // La misma máscara como canvas blanco con transparencia.
  function generarMascara(canvasOrigen, mascara) {
    const ancho = canvasOrigen.width, alto = canvasOrigen.height;
    const valores = valoresMascara(canvasOrigen, mascara);
    const salida = new ImageData(ancho, alto);
    const o = salida.data;
    for (let p = 0; p < valores.length; p++) {
      o[p * 4] = 255; o[p * 4 + 1] = 255; o[p * 4 + 2] = 255; o[p * 4 + 3] = valores[p];
    }
    const canvas = document.createElement('canvas');
    canvas.width = ancho; canvas.height = alto;
    canvas.getContext('2d').putImageData(salida, 0, 0);
    return canvas;
  }

  function copiarCanvas(canvas) {
    const copia = document.createElement('canvas');
    copia.width = canvas.width; copia.height = canvas.height;
    copia.getContext('2d').drawImage(canvas, 0, 0);
    return copia;
  }

  // Aplica las capas extra de un filtro sobre "canvas" (que ya tiene los
  // ajustes principales). "base" es la imagen SIN ajustes, sobre la que se
  // calculan las máscaras. procesar(canvas, ajustes) es el
  // procesarCanvas del motor (procesar-filtros.js), que completa los ajustes
  // con sus neutros.
  //
  // Cada capa se mezcla píxel a píxel según su máscara, sin tocar el alfa.
  // Antes se recortaba con la máscara y se dibujaba encima (source-over), y
  // eso suma la opacidad de las dos: una zona semitransparente de un PNG
  // quedaba más opaca (viñeta de "vintage": alfa 128 → 143), aun sin máscara.
  function aplicarCapasExtra(canvas, base, capas, procesar) {
    const ctx = canvas.getContext('2d');
    const ancho = canvas.width, alto = canvas.height;
    for (const capa of capas) {
      const procesado = copiarCanvas(canvas);
      procesar(procesado, capa.ajustes);
      const nuevo = procesado.getContext('2d').getImageData(0, 0, ancho, alto).data;
      const actual = ctx.getImageData(0, 0, ancho, alto);
      const d = actual.data;
      const valores = capa.mascara ? valoresMascara(base, capa.mascara) : null;
      for (let p = 0, i = 0; p < ancho * alto; p++, i += 4) {
        const m = valores ? valores[p] / 255 : 1;
        if (m === 0) continue;
        d[i] += (nuevo[i] - d[i]) * m;
        d[i + 1] += (nuevo[i + 1] - d[i + 1]) * m;
        d[i + 2] += (nuevo[i + 2] - d[i + 2]) * m;
      }
      ctx.putImageData(actual, 0, 0);
    }
  }

  // Llena "contenedor" con un botón por filtro (clase preset-miniatura, con
  // un <canvas> para la miniatura y data-preset = clave), agrupados.
  function crearGrilla(contenedor) {
    const boton = (clave) => `<button type="button" class="preset-miniatura" data-preset="${clave}" title="${nombre(clave)}"><canvas></canvas><span>${nombre(clave)}</span></button>`;
    contenedor.innerHTML = boton('ninguno') + GRUPOS.map((g, i) =>
      `<div class="titulo-grupo-presets">${tituloGrupo(i)}</div>` + g.filtros.map(([clave]) => boton(clave)).join('')).join('');
    contenedor.querySelector('[data-preset="ninguno"]').classList.add('activo');
    // Se reescriben los textos en vez de rehacer la grilla: rehacerla borraria
    // las miniaturas ya dibujadas en cada <canvas> y cual esta seleccionado.
    document.addEventListener('cambioIdioma', () => {
      contenedor.querySelectorAll('.preset-miniatura').forEach((b) => {
        const texto = nombre(b.dataset.preset);
        b.title = texto;
        b.querySelector('span').textContent = texto;
      });
      contenedor.querySelectorAll('.titulo-grupo-presets').forEach((d, i) => { d.textContent = tituloGrupo(i); });
    });
  }

  function tituloGrupo(indice) {
    return enIngles() ? GRUPOS_EN[indice] : GRUPOS[indice].titulo;
  }

  function nombre(clave) {
    if (enIngles() && NOMBRES_EN[clave]) return NOMBRES_EN[clave];
    if (clave === 'ninguno') return 'Sin filtro';
    for (const g of GRUPOS) for (const [k, n] of g.filtros) if (k === clave) return n;
    return clave;
  }

  return { FILTROS, GRUPOS, nombre, generarMascara, aplicarCapasExtra, crearGrilla };
})();

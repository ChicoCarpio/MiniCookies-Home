# Sistema de interfaz de Mini Cookies

Cómo se construye una página de este sitio, en escritorio y en celular.

Vive en la raíz del repo, al lado del código que describe. Reemplaza a la
antigua `GUIA-INTERFAZ-MOVIL.md`, que solo cubría el modo de pantalla completa
en celular y vivía fuera del repo. El cambio de nombre no es cosmético: el problema nunca fue
que faltara consejo sobre móvil, sino que **cada página se construía desde
cero**. Esta guía describe piezas que ya existen como código, no recomendaciones
a reimplementar.

> **Regla de oro.** Si dos páginas necesitan lo mismo, no se copia: se sube a la
> pieza compartida. Durante semanas la publicidad tapó el encabezado en tres
> herramientas porque la misma pantalla estaba escrita dos veces, y arreglar una
> no arreglaba la otra. La duplicación no es deuda futura: es el bug de hoy.

---

## 1. Por dónde empezar

Antes de escribir una línea, decidí **qué arquetipo** es la página. Son cinco y
cubren las 120 páginas del sitio. Cada uno tiene su esqueleto en la sección 4.

| Si la página… | Es el arquetipo |
|---|---|
| lista herramientas o categorías para elegir | **A — Menú** |
| recibe una entrada, hace algo y devuelve un resultado | **B — Herramienta simple** |
| deja manipular un contenido con varios grupos de ajustes | **C — Edición a pantalla completa** |
| muestra páginas/hojas/diapositivas para seleccionar | **D — Vista previa de documento** |
| es solo texto para leer | **E — Página de texto** |

La gran mayoría (unas 100) son **B**. Solo cuatro son **C**:
`herramientas/collage-imagenes.html`, `herramientas/Editor-de-color-de-imagen.html`,
`herramientas/editor-avanzado-imagen.html` y `herramientas/generador-stickers.html`. No conviertas una
herramienta simple en arquetipo C porque "quedaría lindo": C solo se justifica
cuando la persona alterna seguido entre grupos de ajustes mirando el mismo
contenido.

---

## 2. Tokens

Están en `:root` de `assets/css/estilos.css` y cambian solos con el tema claro /
oscuro. **Usá siempre el token, nunca el color literal**: un `#fff` escrito a
mano se ve bien en claro y queda ilegible en oscuro.

| Token | Para qué |
|---|---|
| `--azul` / `--azul-oscuro` | color de acento (es dorado, el nombre quedó de antes) |
| `--texto` / `--texto-suave` | texto principal / secundario |
| `--superficie` / `--superficie-alt` | fondo de tarjetas y paneles |
| `--gris-fondo` / `--gris-borde` | fondo de página y bordes |
| `--verde` | confirmación |
| `--radio` | radio de esquina estándar |
| `--aviso-*`, `--resultado-*`, `--activo-fondo` | fondos de aviso, resultado y estado activo |

El módulo de pantalla completa publica además `--app-alto-encabezado`, que mide
en vivo el alto del encabezado ya compactado. Es de solo lectura: lo escribe
`assets/js/app-movil.js`.

---

## 3. Las piezas

Todas viven en `assets/css/estilos.css` salvo donde se indique. El nombre de
clase **es la API**: si usás la clase, heredás el comportamiento en los dos
tamaños y las correcciones futuras.

### Estructura de página

| Pieza | Qué es | Contrato |
|---|---|---|
| `header.sitio` | encabezado del sitio | ver abajo |
| `footer.sitio` | pie con los enlaces legales | idéntico en las 120 páginas |
| `main` | columna central, ancho máximo y padding | va un único `<h1>` adentro |
| `.panel-herramienta` | caja donde vive la herramienta | |
| `.anuncio` | hueco de publicidad | ver sección 5 |
| `#fondoDecorativo` | marca de agua de fondo | `position: fixed` y recortado; **nunca** causa desbordes |

### Encabezado

La marca es **un solo enlace** con el logo y la palabra adentro, igual en las
120 páginas (solo cambia el `../` según la carpeta):

```html
<div class="marca"><a href="../index.html"><img src="../marca-logo/mini-cookies-logo-horizontal.png" alt="Mini Cookies" class="logo-marca"><div class="texto-marca" aria-hidden="true">Mini<span>Cookies</span></div></a></div>
```

El nombre accesible sale del `alt` del logo; la palabra lleva `aria-hidden`
porque repite lo mismo (y en pantalla angosta ni se ve). Nada de estilos en
línea: el enlace se arma en `estilos.css`.

Hay **una** definición. En pantalla angosta (≤900px) se compacta solo: se oculta
la palabra "MiniCookies", el logo baja a 22px, los enlaces se achican y centran,
y los botones de idioma y tema quedan al final. Entra en una fila de 46px tanto
con un enlace (una herramienta) como con tres (un menú). Abajo de ~360px los
enlaces pasan a un segundo renglón **enteros**, nunca partidos al medio.

No agregues reglas de encabezado en la página. Si algo no entra, se arregla en
`estilos.css` para todo el sitio.

> **Trampa documentada.** Para pisar el tamaño de `.boton-tema` dentro del media
> query hace falta `header.sitio .boton-tema`: la regla base está más abajo en el
> mismo archivo y, a igual especificidad, gana ella. Esos 16px de más eran lo que
> partía el encabezado en dos renglones.

### Entrada

| Pieza | Qué es |
|---|---|
| `.zona-arrastre` | caja de "arrastrá tu archivo acá"; lleva un `<input type="file" hidden>`. Se conecta con `assets/js/zona-arrastre.js` (ver abajo) |
| `.controles` | fila de controles; es flex y envuelve |
| `label.campo` | etiqueta + control, apilados |
| `.campo-deslizador` | variante con el valor a la derecha de la etiqueta |
| `.boton` | acción principal. Variantes (2026-10-12, administrador de contraseñas): `.boton.secundario` (contorno, para acciones que no compiten con la principal), `.boton.peligro` (contorno rojo, para borrar) y `.boton.chico` (más compacto, para botones dentro de filas de una lista) |
| `.boton-tema` | botón circular de icono (idioma, tema) |
| `.selector-direccion` | opciones "A → B / B → A" de una herramienta bidireccional; con `assets/js/selector-direccion.js` y un `data-panel-direccion="valor"` en cada panel, muestra solo el de la opción elegida, sin código propio |
| `.columnas` | dos columnas (entrada \| resultado) que se apilan abajo de 700px |

**Zona de arrastre: nunca se cablea a mano.** Se incluye
`assets/js/zona-arrastre.js` antes del script de la página y se llama:

```js
conectarZonaArrastre(zonaArrastre, entradaArchivo, (archivos) => { … });
```

`archivos` es siempre un `Array` de `File`, venga del selector o de lo soltado.
El módulo limpia el selector después de cada elección (si no, volver a elegir
el mismo archivo no hace nada) y hace **toda** `.zona-arrastre` de la página
usable con teclado: foco con Tab, Enter/Espacio para elegir, contorno visible.
Estaba copiado en 77 herramientas: 53 no limpiaban el selector y ninguna
funcionaba con teclado. Si una herramienta acepta un solo tipo o un solo
archivo, se filtra en la función que se le pasa, no reescribiendo el módulo.

> **Trampa documentada.** Un `<select>` toma el ancho de su opción más larga, y
> un hijo de un contenedor flex no se achica por debajo de su contenido. Por eso
> `.controles > *` y `label.campo` llevan `min-width: 0` y los controles
> `max-width: 100%`. Sin eso, una sola opción descriptiva ensancha la página
> entera en celular. Ya está resuelto: no lo reintroduzcas con estilos propios.

### Salida

| Pieza | Qué es |
|---|---|
| `.resultado` | caja de resultado; oculta hasta que se le agrega `.visible` |
| `a.boton` | un enlace con aspecto de botón (una descarga, los de la 404): `<a class="boton">`, sin estilos en línea |
| `.grid-miniaturas`, `.miniatura-pagina` | grilla de páginas/hojas seleccionables |
| `.contenedor-miniaturas`, `.selector-vista` | lista o grilla conmutable |
| `.lista-archivos` | lista de archivos cargados |
| `.tabla-ancha` | envoltorio para una tabla que no entra: se desplaza sola sin arrastrar la página |
| `.contenedor-tabla` + `table.tabla-datos` | tabla de datos de una planilla, con encabezado fijo y desplazamiento propio; `tr.fila-par` y `tr.fila-espaciadora` para filas alternadas y virtualización. La página solo ajusta lo suyo (altura, resaltados) |

### Menús

`.hero`, `.buscador`, `.bloque-categorias`, `.grid-categorias`,
`.tarjeta-categoria`, `.grid-herramientas`, `.tarjeta-herramienta`,
`.titulo-seccion`, `.indice-catalogo`, `.grupo-catalogo`, `.sin-resultados`,
`.aviso-privacidad`, `.estadisticas`.

### Componentes con JavaScript

| Pieza | Archivos | Para qué |
|---|---|---|
| Pantalla completa móvil | `assets/css/app-movil.css` + `assets/js/app-movil.js` | arquetipo C (sección 6) |
| Ayuda "?" | `assets/css/ayuda-tip.css` + `assets/js/ayuda-tip.js` | explicaciones secundarias en burbuja |
| Tema | `assets/js/tema.js` | claro / oscuro |
| Idioma | `assets/js/idioma.js` | es / en |
| Fondo decorativo | `assets/js/fondo-decorativo.js` | marca de agua |
| Descargas | `assets/js/descargar.js` | `descargarTexto`, `nombreConFecha` |
| Reordenar al tacto | `assets/js/reordenar-tactil.js` | arrastrar para reordenar |
| Ordenar archivos | `crearBarraOrden` + `conMover` de `assets/js/miniaturas.js` + `.barra-orden`, `.posicion-orden`, `.mover-tarjeta` (en `estilos.css`) + `orden-*` / `mover-*` (en `i18nComun`) | ordenar una lista de archivos enteros sin depender del arrastre (2026-10-09, Unir PDF). Las tarjetas se crean con `crearTarjetaMiniatura({ …, conMover: true, nombre })` y la página pone la barra con `crearBarraOrden(contenedor, alReordenar)`: número de posición, ◀ ▶ en cada tarjeta (en la vista Lista apuntan arriba/abajo), Nombre A → Z / Z → A (orden natural: "2" antes que "10") e Invertir. Renumera sola también al arrastrar, agregar o quitar, y devuelve el foco a ◀ ▶ para poder apretar varias veces con el teclado. Se suma a `habilitarArrastreReordenable`, no lo reemplaza. Para un archivo que tarda en prepararse, la tarjeta se crea en el acto con `cargando: true` (una barra en lugar de la miniatura, `.carga-tarjeta` / `.barra-carga`, `role="progressbar"`); `progresoTarjeta(tarjeta, fraccion)` la avanza (`null`: va y viene) y `completarTarjeta(tarjeta, { elementoVisual, etiqueta })` pone la miniatura. La barra se mueve con `transform`, que anima el compositor: sigue andando aunque la página esté ocupada |
| Hoja para medir un documento | `HtmlAPdf.crearLienzo(clase)` de `assets/js/html-a-pdf.js` + `.lienzo-pdf` (en `estilos.css`) | la hoja fuera de pantalla donde html-a-pdf.js mide un documento antes de escribir el PDF; con `'libro'`, A4 con letra de libro (EPUB). Quien la pide la saca con `remove()`. Antes cada página tenía su `<div id="lienzoPdf">` con sus estilos |
| Escapar y limpiar HTML | `assets/js/seguridad-html.js` | `escaparHtml` (única definición del sitio) y la limpieza de HTML de archivos del usuario |
| Zona de arrastre | `assets/js/zona-arrastre.js` | `conectarZonaArrastre` y acceso con teclado (sección 3, Entrada) |
| Selector de dirección | `assets/js/selector-direccion.js` | panel según la dirección elegida (`data-panel-direccion`) |
| Deshacer y rehacer | `assets/js/historial.js` + `.botones-historial` (en `editor-imagen.css`) + `title-deshacer`/`title-rehacer` (en `i18nComun`) | los tres editores de imagen (2026-10-07). La página escribe los dos botones `#botonDeshacer` y `#botonRehacer` dentro de un `<div class="botones-historial">`, debajo de la imagen. Después llama a `Historial.conectar({ capturar, restaurar, deshacer, rehacer, lienzo, activo })`: `capturar()` devuelve el estado con una `firma` (texto que lo resume) y `restaurar(paso)` vuelve a él. Con eso el módulo arma los pasos solo: después de cada clic, cambio, toque o tecla compara la firma. Un deslizador cuenta como un paso al soltarlo. Habilita los botones y conecta Ctrl+Z / Ctrl+Shift+Z / Ctrl+Y (fuera de los campos de texto) y el toque con dos dedos sobre `lienzo`. Los lienzos que cambian se copian con `Historial.crearCopiador()` solo cuando cambia su versión. Hay un tope de memoria (192 MB en celular, 512 MB en PC) y, pasado ese tope, se olvidan los pasos más viejos. Un cambio que llega tarde (fotos agregadas, un proyecto abierto) llama a `comprobar()` o a `reiniciar()` |
| Pantalla completa | `assets/js/pantalla-completa.js` + `.pantalla-completa-activa` / `.franja-anuncio-completa` (en `editor-imagen.css`) + `title-pantalla-completa` (en `i18nComun`) | los tres editores de imagen (2026-10-07). La página pone `#botonPantallaCompleta` (con `hidden`) en la fila de `.botones-historial` y llama a `PantallaCompleta.conectar({ contenedor, boton, activo })`. El contenedor es el mismo que usa app-movil: `[data-app-movil]` o el layout. Con eso el botón y la tecla F entran y salen, y Esc sale. Donde el navegador no lo permite (iPhone), el botón queda oculto. En pantalla completa solo se ve lo que está adentro del contenedor. Por eso, en PC, el anuncio pasa a una franja propia de 90 px debajo del editor y vuelve a su lugar al salir; se mueve con `AppMovil.moverSinRecargar` (la función de `app-movil.js`, para no recargar el anuncio). En celular ya está en la tira de app-movil. Todo lo que se abra por encima va en `PantallaCompleta.capa()` (`document.fullscreenElement` o `<body>`); `ayuda-tip.js` ya lo hace. El reparto de adentro es de cada página (el collage pone su propio alto de lienzo) |
| Atajos de teclado | `assets/js/atajos.js` | `Atajos.conectar({ 'ctrl+0': fn, ' ': { abajo, arriba } }, activo)`: una sola regla para todos los atajos, también los de deshacer. Nunca mientras se escribe en un campo de texto, y la barra espaciadora nunca sobre un botón o un enlace. Un par `{ abajo, arriba }` es para lo que se mantiene apretado (ver el original, mover la vista) y se suelta solo si la ventana pierde el foco |
| Galería de texturas | `assets/js/efectos-textura.js` | `EfectosTextura.crearGaleria(galeria)` llena `#modosTextura` con un botón por textura (`.miniatura-textura`, `data-textura`): "Ninguna", las de siempre (pixelado, puntos, ruido…) y, si la página cargó `efectos-mosaico.js`, los mosaicos y el cómic. Trae los nombres en es/en. Llamarla **antes** de conectar los clics de la galería. La usan los dos editores y el generador de stickers; antes los botones estaban escritos a mano en cada página |
| Mosaicos y cómic | `assets/js/paleta-color.js` + `assets/js/efectos-mosaico.js` (+ `vendor/imagetracerjs` para el cómic) | texturas de pixel art, cuadrícula, punto de cruz, cuentas, diamantes, abalorios, pop art y cómic (editores de imagen y generador de stickers). Son **texturas, no filtros**: no hay filtros predefinidos que las usen, y cada una se ajusta con sus propios controles. `EfectosMosaico.crearPanel(contenedor, { alCambiar, campoIntensidad })` arma esos controles (celdas, colores, tramado, separación o línea, grosor…) con `leer()`, `poner(opciones, tipo)` y `mostrar(tipo)`; con uno de estos tipos oculta `campoIntensidad`, el deslizador de las demás texturas, que no usan. Trae sus propias traducciones. Se aplican por `EfectosTextura.aplicarCompleto(…, opciones)` como cualquier textura; `EfectosTextura.esMosaico(tipo)` dice si dejan transparente lo que no es figura (las de siempre no: en un recorte hay que devolverles la transparencia) |
| Rejilla de hoja | `assets/css/rejilla-hoja.css` + `assets/js/rejilla-hoja.js` + `assets/js/motor-hojas.js` (+ `vendor.js`, `lector-hojas.js` y los módulos livianos del motor de Excel antes) | una hoja de Excel, CSV o formato antiguo como se ve al abrirla en Excel o Google Sheets (2026-10-15, plan `06-plan-interfaz-visor-e-impresion.md`; hoy solo «Ver Excel o CSV»). La página pone un contenedor **con alto propio** (la rejilla lo llena), hace `await MotorHojas.cargarDibujo()` (trae el motor de dibujo, sin pdf-lib) y llama a `RejillaHoja.crear(contenedor, { libro: await LectorHojas.abrir(archivo) })`. Garantiza: hoja dibujada por el motor en un Canvas nítido, desplazamiento nativo (también en celular), celda activa y rango (ratón, Mayús, arrastre, teclado, encabezados, combinadas enteras), cuadro de nombre, barra de contenido (fórmula con `=`), pestañas con color de pestaña (el color de una hoja que no se abrió aún no se conoce), barra de estado (Recuento, Suma, Promedio) y zoom 50–200 % (botones, deslizador, Ctrl + rueda, pellizco). Solo lectura. En celular un toque corto elige la celda y un deslizamiento mueve la hoja. Desde la fase A completa también: paneles inmovilizados del archivo, zoom del archivo al abrir, buscar (Ctrl+F; opción «Todo el libro»), ancho de columna (arrastrar el borde o doble clic, solo en la vista), menú «Hojas ocultas», asas de selección con el tacto y pantalla completa (`.rh-completa`, bloquea el desplazamiento de la página con `html.rh-bloqueo`). También «Ordenar y filtrar» (solo para mirar; `assets/js/rejilla-orden-filtro.js` tiene que cargarse antes): botón en la barra, botones en la fila de encabezado de los datos y menú con orden y casillas de valores; `ordenarYFiltrar(si)` en la API. Devuelve `{ mostrarHoja, seleccion, copiar, zoom, alCambiar, rectCelda, enfocar, pantallaCompleta, buscar, destruir }`. **No sustituye a la tabla:** el Canvas no lo lee un lector de pantalla, así que la página mantiene la tabla HTML detrás de «Ver como tabla»; si el motor de dibujo no carga, se ve solo la tabla. Los textos (`rh-…`) los trae el módulo en es/en y los de la página mandan. La hoja se dibuja blanca en los dos temas, por eso la selección usa colores fijos y no tokens. Excel → PDF la usará para elegir rangos |

Los dos primeros necesitan que incluyas **su CSS y su JS**. Olvidarse de uno es
un error silencioso: la página se ve, pero sin el comportamiento.

---

## 4. Los cuatro arquetipos

Todos comparten este marco. Lo que cambia es el contenido de `<main>`.

```html
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title data-i18n="titulo-pagina">… gratis — Mini Cookies</title>
  <meta name="description" data-i18n-attr="content:meta-descripcion" content="…">
  <link rel="icon" type="image/png" href="../marca-logo/minicookies_galleta.png">
  <link rel="stylesheet" href="../assets/css/estilos.css">
  <link rel="stylesheet" href="../assets/css/fondo-decorativo.css">
  <script>(function(){try{var t=localStorage.getItem('tema');if(t)document.documentElement.setAttribute('data-theme',t);}catch(e){}})();</script>
  <script>window.i18nPagina = { es: { … }, en: { … } };</script>
</head>
<body>
<div id="fondoDecorativo"></div>
<header class="sitio">…</header>
<main>…</main>
<footer class="sitio">…</footer>
<script src="../assets/js/tema.js"></script>
<script src="../assets/js/fondo-decorativo.js"></script>
<script src="../assets/js/idioma.js"></script>
</body>
</html>
```

El script del tema va **en el `<head>` y en línea**: si esperara a los scripts
del final, la página parpadearía en claro antes de pasar a oscuro.

### A — Menú

`index.html`, `todas.html` y las 9 categorías.

```html
<main>
  <div class="hero">
    <h1 data-i18n="hero-titulo">📄 Herramientas de PDF</h1>
    <div class="aviso-privacidad" data-i18n="aviso-privacidad">🔒 …</div>
  </div>
  <div class="anuncio" data-i18n="anuncio">Espacio publicitario</div>
  <div class="grid-herramientas" data-catalogo="pdf"></div>
</main>
```

**Las tarjetas no se escriben en la página.** Salen de `catalogo.json`, la única
lista de herramientas del sitio: qué herramientas lleva cada categoría, en qué
orden y con qué icono, más el título y la descripción de cada una en español e
inglés, escritos una sola vez. El build (`catalogo.mjs`) llena cada
`data-catalogo`, la lista del buscador de la portada (`{{CATALOGO_BUSCADOR}}`)
y los conteos, y falla si una herramienta de `herramientas/` falta en el
catálogo o no está en ninguna categoría. Por eso, abiertas directo del disco,
las páginas de menú se ven sin tarjetas: para verlas, `npm run preview`.

Antes eran tres copias a mano (categoría, `todas.html`, buscador) y se habían
separado: el buscador no encontraba el generador de stickers y `todas.html`
tenía otro orden y otros iconos.

El aviso de privacidad **es exclusivo de los menús** (ver sección 5).

### B — Herramienta simple

El 85% del sitio.

```html
<main>
  <h1 data-i18n="hero-titulo">Generador de códigos QR</h1>
  <p data-i18n="intro">Escribí un link o texto y generá un código QR.</p>

  <div class="panel-herramienta">
    <div class="zona-arrastre" id="zonaArrastre">
      <span data-i18n="zona-arrastre">📄 Arrastrá tu archivo acá…</span>
      <input type="file" id="entradaArchivo" hidden>
    </div>
    <div class="controles">
      <label class="campo">
        <span data-i18n="campo-x">Etiqueta</span>
        <select id="x">…</select>
      </label>
      <button class="boton" id="botonAccion" data-i18n="boton-accion">Hacer</button>
    </div>
    <div class="resultado" id="resultado">…</div>
  </div>

  <div class="anuncio" data-i18n="anuncio">Espacio publicitario</div>
</main>
```

Desplazamiento normal de página. **No** la conviertas en pantalla completa.

### C — Edición a pantalla completa

Ver sección 6: el esqueleto es el de B, más los marcadores del módulo.

Los **editores de imagen** comparten además `assets/css/editor-imagen.css`:
acordeón de ajustes, deslizadores, miniaturas de ajustes predefinidos y de
texturas, botones de modo y el panel lateral en pantalla angosta. Antes eran
43 reglas idénticas copiadas en cada `<style>`. En la página queda solo lo
propio (ancho, distribución, tamaño de miniaturas).

También ahí, **`.trabajando-ia`** (2026-10-04): el velo animado sobre TODA la
imagen mientras un modelo de IA se descarga o analiza, para que no parezca que
la página se trabó (pedido del usuario). Va dentro del panel del lienzo, que
necesita `position: relative`; no tapa los toques; sin animación con
`prefers-reduced-motion`. Marcado: `<div class="trabajando-ia" role="status"
aria-live="polite" hidden><span class="trabajando-ia-texto"></span></div>`.
Desde el 2026-10-08 lo maneja **`assets/js/trabajando-ia.js`**
(`TrabajandoIA.crear(velo)` → `empezar(texto)` → `{ texto(t), terminar() }`,
y `TrabajandoIA.textoDeDescarga`): un texto por trabajo (descargando →
analizando → buscando), el del más reciente, y se esconde cuando no queda
ninguno. Lo usan el Editor avanzado y el Generador de stickers. Una herramienta
nueva que corra un modelo sobre una imagen usa esta pieza, no un spinner propio
ni una barra de progreso aparte.

Y comparten el **motor de filtros**, `assets/js/procesar-filtros.js`
(`renderizarFiltro`: textura → ajustes → capas extra → intensidad → textura →
desenfoque). Lo usan el Editor de color, el Editor avanzado (una vez por capa,
con `procesarCapaEnCanvas` como adaptador) y el generador de stickers. Hasta el
2026-10-04 cada editor tenía su copia, y se habían separado: la mezcla de
"Intensidad del filtro" volvía más opacas las zonas semitransparentes en el
Editor de color y no en el avanzado. Tanto la intensidad como las capas
extra (`FiltrosColor.aplicarCapasExtra`) mezclan píxel a píxel y no tocan el
alfa: dibujar una capa encima con `source-over` suma opacidad (con "cine", un
píxel de alfa 128 llegaba a 229). Un ajuste o filtro nuevo se agrega ahí,
nunca en una página. Para la vista previa en vivo el motor recibe `scratch`
(canvases de repuesto) y `memo` (lo que depende solo de la imagen); quien llama
vacía el memo cuando cambia la imagen de entrada.

### Esqueleto de PC de los editores de imagen (2026-10-08)

Lo maneja `assets/js/esqueleto-editor.js` + la sección "Esqueleto de PC" de
`assets/css/editor-imagen.css` (plan del rediseño, fase 6). Con **901 px o más
y una foto cargada**, el editor ocupa toda la ventana debajo del encabezado del
sitio (que queda fijo arriba) y se arma así:

```
┌ barra superior: ↶ ↷ │ 👁 Original │ − 100% + ⤢ │ vista │ ⛶ │  [Descargar] ⋯ ┐
├──┬──────────────────┬────────────────────────────────────────────────────┤
│🗂│ Editando: …      │                                                    │
│☀️│ ☀️ LUZ           │                    escenario                       │
│🎨│ (una categoría)  ├────────────────────────────────────────────────────┤
│  │                  │      anuncio: franja propia de 90 px               │
└──┴──────────────────┴────────────────────────────────────────────────────┘
```

Debajo de 900 px el marco es de `app-movil.js` (sección 6), pero las
categorías, la barra y "Editando" siguen saliendo del esqueleto (ver
"Esqueleto de celular", abajo). Los dos conviven porque **el esqueleto no
mueve las categorías**.

**El contrato:**

| Marcador / opción | Dónde | Qué hace |
|---|---|---|
| `data-categoria="luz"` | cada bloque de controles (hijo directo de un panel del contenedor) | se ve solo con su categoría elegida; varios bloques pueden ser de la misma (Efectos = efectos + resplandor + dispersión) |
| `orden` | `EsqueletoEditor.conectar` | categorías de la barra de íconos, con `'|'` como separador. Nombre e ícono salen del registro `CATEGORIAS` del módulo: **iguales en los tres editores** |
| `barra` | ídem | los botones **reales** de la página (deshacer, rehacer, ⛶, descargar, el `<select>` de resolución, avisos cortos y los del menú "⋯") se **mudan** a la barra y vuelven a su lugar al pasar a celular. `original` y `zoom` pueden ser funciones: el módulo arma el botón |
| `editando` | ídem | arma "Editando: …" arriba de la columna (solo el Editor avanzado) |
| `alCambiarCategoria(cat)` | ídem | para lo que la página muda por su cuenta (los controles del pincel a "Borrar", la tira de fotos a "Fotos"). Con `null` el esqueleto se apagó: devolver todo |
| `data-sin-esqueleto` | cualquier elemento | no se ve en PC (lo que la barra ya resuelve, como "Mantené presionado para ver el original") |

Qué garantiza:

- **Una sola categoría a la vista.** Tocar la abierta cierra la columna (la
  imagen gana ancho). Teclas **1–9 y 0** en el orden de la barra; **Tab**
  esconde la columna **solo en pantalla completa** (afuera sigue recorriendo
  la página).
- **El anuncio en su franja**, de 90 px fijos y sin nada encima; se muda con
  `AppMovil.moverSinRecargar` y vuelve. Como ya está adentro del contenedor,
  `pantalla-completa.js` no lo toca.
- **Descargar desde la barra baja el archivo** (toca el enlace que arma el
  editor), solo con un clic de verdad: las pruebas tocan el botón y leen el
  enlace sin bajar nada.
- El menú "⋯" y la lista de "Editando" viven dentro del editor: se ven también
  en pantalla completa.

Trampas ya pagadas:

- **Al pasar de celular a PC, el esqueleto se arma en el turno siguiente**
  (`setTimeout`): `app-movil.js` primero devuelve el anuncio a su lugar, y si
  el esqueleto se lo llevaba antes, `app-movil` se lo sacaba de la franja. Al
  revés (PC → celular) se apaga enseguida, y solo devuelve el anuncio si
  todavía está en su franja.
- **Un contenedor con `align-items: start`** (el collage) dejaba la barra de
  íconos y la columna cortas y el collage desbordado: el esqueleto re-afirma
  `align-items: stretch`.
- **`.modos { display: flex }` le gana al atributo `hidden`**: un `.modos`
  que se esconde con `hidden` necesita su `[hidden] { display: none }`.
- **`typeof` sobre un `const` que todavía no se inicializó lanza error**: si
  una función de la página puede correr antes de conectar el esqueleto,
  declararlo con `let esqueleto = null` (o `var`) arriba.

`npm run revisar` comprueba, con el editor abierto a 1366×768, que el
escenario sea al menos el 50 % de la pantalla, que el anuncio mida 90 px sin
nada encima y que se vea una sola categoría; al terminar dice en qué páginas
lo revisó. `npm run probar` (`esqueleto de PC y de celular`) recorre los tres
editores, lo contextual y la ida y vuelta a 390 px.

### Esqueleto de celular (2026-10-08)

Plan del rediseño, fase 7: **lo mismo que en PC, en otro lugar**. Debajo de
900 px y con una foto cargada:

```
┌ encabezado del sitio ──────────┐
├ ↶ ↷ │ 👁 ⤢ ⛶        [⤓]  ⋯ ───┤  la barra de la PC (mismos botones reales)
│ Editando: Capa 1 ▾             │  solo el Editor avanzado
│            IMAGEN              │
├ [ anuncio 50 px ] ─────────────┤  franja propia, DEBAJO de la imagen
│ ☀️ la categoría abierta (hoja) │  la columna de la PC
│🗂 🔤 🎯 🧽 │ ☀️ 🎨 ◐ ✨ 🎞 🧱 → │  la barra de íconos de la PC, acostada
└────────────────────────────────┘
```

Cómo se reparten el trabajo:

- `app-movil.js` ve bloques con `data-categoria` y entra en **modo
  esqueleto**: no arma categorías de los acordeones ni de
  `data-app-categoria` (se ignoran en estos tres editores), ni la tira de
  acciones de abajo. Al encender pide `EsqueletoEditor.de(contenedor).movil`.
- El esqueleto arma la barra con los botones reales (`.ed-barra-celular`):
  quedan ↶ ↷, Original y Encajar solo con su ícono, ⛶, Descargar como botón
  ⤓ y "⋯". Lo que no entra en 390 px va al menú: el zoom (− % +, que no
  cierran el menú: el pellizco no puede ser la única forma), la resolución
  con su rótulo y los avisos. Descargar baja el archivo directo, igual que
  en PC.
- Abrir una categoría mueve **todos sus bloques** a la hoja, en su orden; el
  primero sin título y los demás con su nombre como subtítulo, como la
  columna. Se avisa al esqueleto (`movil.abrio`), que la deja como elegida y
  llama a `alCambiarCategoria`: "Borrar" se lleva los controles del pincel y
  "Fotos" la tira del collage, igual que en PC.
- `esqueleto.abrir(cat)` y `esqueleto.activo()` valen en las dos pantallas:
  en celular, abrir es abrir la hoja (agregar un texto abre "Texto e imagen";
  tocar una foto del collage abre "Fotos").
- La hoja arranca cerrada (la imagen tiene todo el alto); tocar la categoría
  abierta la cierra, como en PC la columna.
- `body.ed-esqueleto-movil` marca el modo. `[data-sin-esqueleto]` y la fila
  vacía de `.botones-historial` se esconden en las dos pantallas.

Trampa ya pagada: **el contenedor fijo hereda su margen de escritorio** (el
collage trae 16 px arriba) y dejaba una franja vacía entre el encabezado y la
barra; `app-movil.css` lo pone en 0.

### D — Vista previa de documento

Como B, con `.grid-miniaturas` / `.contenedor-miniaturas` dentro de
`.panel-herramienta`. Las miniaturas se generan con `assets/js/miniaturas.js`
(las 19 herramientas de PDF ya lo usan). Las diapositivas renderizadas de
PowerPoint van en `.vista-previa-diapositivas`.

### E — Página de texto

`privacidad.html`, `terminos.html`, `creditos.html`, `contacto.html` y `404.html`: el marco
común y prosa dentro de un `.panel-herramienta`, sin controles. No llevan
`.aviso-privacidad` (eso es de los menús) ni `.anuncio` (nunca lo tuvieron).

Dos detalles propios: `creditos.html` envuelve su tabla en `.tabla-ancha`, que es
lo que evita que se lleve la página entera a lo ancho en un teléfono, y
`contacto.html` tiene un formulario.

`404.html` es la excepción a las rutas relativas: Cloudflare la sirve en
cualquier dirección que no existe (por ejemplo `/herramientas/algo/otra`), así
que lleva `<base href="/">` y el build la pasa a `/en/` en su versión en inglés.
Tiene `noindex`, y por eso el build no le agrega SEO, no la pone en el sitemap
y no le inyecta AdSense (Google no admite anuncios en páginas de error).

Piezas: `table.tabla-creditos` (una fila nueva de créditos es `<tr><td>…`
sin estilos: antes cada celda llevaba su `style="padding…"`) y
`.texto-actualizacion` para la línea "Última actualización".

---

## 5. Reglas transversales

### Publicidad

Mostrar publicidad es un objetivo del proyecto (ver `../CLAUDE.md`,
"Monetización"), así que **toda página lleva su `.anuncio`, visible en
escritorio y en celular**.

- **Nunca ocupa el lugar del encabezado.** Es literalmente lo que pasaba: la capa
  de pantalla completa era `position: fixed; inset: 0`, tapaba el encabezado, y
  al abrir el editor en el teléfono lo primero que se veía era un anuncio donde
  debía estar el logo.
- **Nunca se esconde con `display: none` para ganar espacio.** Si no entra donde
  está, se le busca un hueco propio y acotado, no se lo borra de esa versión del
  sitio.
- **Siempre con alto fijo y `overflow: hidden`**, nunca `flex: 1` ni alto libre:
  un anuncio real de AdSense puede salir mucho más alto que el texto de prueba
  "Espacio publicitario" y empujar el resto de la pantalla.
- **En el fuente se escribe solo el hueco**, siempre igual:
  `<div class="anuncio" data-i18n="anuncio">Espacio publicitario</div>` (puede
  llevar `id`). El build (`ponerUnidadesDeAnuncio` en `build.mjs`) lo convierte
  en la unidad de AdSense de `config.adsenseSlot` del `package.json`: el texto
  pasa a un `<span class="anuncio-texto">` (porque `idioma.js` reemplaza el
  texto de todo `[data-i18n]` y borraba el anuncio) y al lado va el `<ins>`. Un
  hueco escrito de otra forma hace fallar el build. El texto se ve debajo hasta
  que Google llena el hueco. Alto del `<ins>`: 90 px en escritorio, 50 px en
  celular (entra en la tira de 52 px de los editores).
- **Moverlo de lugar recarga el anuncio** (es un `<iframe>`): `app-movil.js` lo
  mueve con `moveBefore`, que lo conserva, donde el navegador lo tiene.

### Privacidad

El mensaje "🔒 Tus archivos nunca se suben a ningún servidor…"
(`.aviso-privacidad`) va **solo en los menús**. Una herramienta individual no lo
repite, ni con palabras propias, en ningún texto visible. Sí puede mencionarlo su
`<meta name="description">`, que es para buscadores. Detalle completo y
excepciones en `../CLAUDE.md`.

### Idioma

Todo texto visible lleva `data-i18n`, y todo atributo traducible
`data-i18n-attr`. Las claves van en `window.i18nPagina` con `es` y `en`. Incluye
los textos que genera el JavaScript: estados, botones creados al vuelo y nombres
accesibles. Si una página no tiene inglés de verdad, no declara alternativa en
inglés — ver la skill `traducir-ingles`.

### Comentarios y detalles secundarios

Una aclaración de *por qué* o *cómo funciona por dentro* no va como texto fijo:
va detrás del botón "?" (`assets/css/ayuda-tip.css` + `assets/js/ayuda-tip.js`).
Lo que hace falta para **usar** el control —etiqueta, estado, valor, acción—
queda siempre visible.

### Accesibilidad

- Área táctil cómoda, separación suficiente y foco visible.
- El estado activo no se indica solo con color.
- Los iconos conservan nombre accesible; `title` es ayuda extra, nunca el único
  nombre.
- Ningún gesto es la única vía: zoom, selección y reordenamiento también tienen
  controles visibles.
- Mientras se procesa algo, los controles incompatibles se deshabilitan y vuelven
  a un estado usable al terminar o al fallar.

### Terceros

Cualquier librería nueva se registra en `creditos.html` en el mismo paso
(nombre, versión exacta, licencia verificada contra npm, enlace al repo). Si ya
está en la tabla, no hay nada que hacer — y reusar lo que ya está es lo
preferible.

---

## 6. Arquetipo C: pantalla completa en celular

Lo maneja `assets/js/app-movil.js` + `assets/css/app-movil.css`. Abajo de 900px,
y solo cuando hay contenido cargado, la herramienta se reorganiza así:

```
┌──────────────┐
│ mini🍪  cat  │  encabezado del sitio, compacto pero COMPLETO
├──────────────┤
│  publicidad  │  tira fina de 52px
├──────────────┤
│              │
│  ESCENARIO   │  lo que se edita; nada se le superpone
│              │
├──────────────┤
│   ajustes    │  la categoría abierta
│ ☀🎨◐✨ 🖼️↺ │  categorías y acciones
└──────────────┘
```

Arriba de 900px no pasa nada: queda la distribución de escritorio tal cual.

### El contrato

Vos marcás cuatro cosas en el HTML; el módulo hace el resto.

| Marcador | Dónde | Qué hace |
|---|---|---|
| `data-app-movil` | contenedor de la herramienta | lo convierte en la capa de pantalla completa |
| `data-app-escenario` | el elemento que se edita | recibe el espacio que sobra y nunca se tapa |
| `data-app-escenario="desplazable"` | ídem | además se desborda y se desplaza, para imagen con zoom propio |
| `data-app-categoria="nombre"` | cada grupo de ajustes | se convierte en un botón de la barra inferior |

```html
<div class="mi-layout" id="editor" style="display:none;" data-app-movil>
  <div data-app-escenario>
    <canvas id="lienzo"></canvas>
  </div>
  <div>
    <div class="bloque-opciones" data-app-categoria="estilo">
      <h3 data-i18n="h-estilo">Estilo</h3>
      …
    </div>
  </div>
</div>
```

Qué garantiza el módulo:

- **Se enciende y se apaga solo.** Vigila si el contenedor se vuelve visible: si
  tu herramienta hace `editor.style.display = ''` al cargar un archivo, con eso
  alcanza. No hay nada que llamar.
- **No duplica nada.** El panel elegido se *mueve* al cajón de abajo y vuelve a
  su lugar al cerrarlo o al pasar a pantalla ancha. Mover conserva los
  escuchadores: el código de la herramienta no se toca.
- **El nombre y el icono de la categoría** salen del título del grupo
  (`.acordeon-titulo` o `h3`). El icono se indexa por la clave `data-i18n` del
  título, **no por el texto visible**, porque el texto cambia de idioma. Dalo de
  alta en `ICONOS_POR_CLAVE` de `app-movil.js`; si falta, sale uno genérico y no
  se rompe nada.
- **Las columnas que no son el escenario se ocultan por descarte**, no por una
  lista de clases. Una herramienta nueva no tiene que registrar la suya.
- **El encabezado no se tapa:** la capa arranca en `--app-alto-encabezado`.
- **La publicidad va en su tira**, no en el lugar del encabezado.
- **Ningún botón sale dos veces.** La barra fija de abajo (🖼️ cambiar imagen,
  ↺ restablecer, ⤓ descargar) es un atajo solo para los botones que **no**
  están dentro de ninguna categoría; si la herramienta los tiene en una
  ("Acciones", "Exportar"), no se copian, y sin botones la barra no aparece.
  Hasta el 2026-10-02 se copiaban siempre y una marca por página
  (`data-app-sin-descarga-flotante`) tapaba solo "Descargar": en el editor de
  color salían los tres repetidos.

Opcional: `data-app-categorias-individuales` en un panel para que cada acordeón
de adentro sea su propia categoría en vez de viajar todo junto.

### Lo que sigue siendo tuyo

El módulo da el marco; el reparto **dentro** del escenario es de la página. No le
toca el `display` a propósito: hay escenarios que reparten su alto entre varias
partes (el collage pone el contador arriba y la tira de fotos abajo; el generador
de stickers pone el grosor del borde debajo del lienzo). Forzarlos a bloque es lo
que aplastó el lienzo del generador de stickers.

```css
@media (max-width: 900px) {
  body.app-activo [data-app-escenario] { display: flex; flex-direction: column; }
  body.app-activo .mi-lienzo { flex: 1 1 auto; min-height: 0; overflow: hidden; }
  /* Ya dentro del cajón: el título lo repite el botón de la categoría. */
  body.app-activo .app-hoja .bloque-opciones h3 { display: none; }
}
```

### Trampas ya pagadas

- **El contenedor trae su alineación de escritorio.** El `align-items: start` del
  collage dejaba la tira de publicidad como una cajita en la esquina; por eso la
  capa re-afirma `align-items: stretch`.
- **La tira de categorías se arrastra con eventos de puntero**, no con el
  desplazamiento nativo: con mouse en ventana angosta el nativo no funciona y las
  últimas categorías quedaban fuera de alcance. El puntero se captura recién
  cuando el gesto supera 6px; capturarlo antes desvía el `click` y tocar una
  categoría deja de abrirla.
- **Un `<a>` como contenedor flex no parte su texto**: adentro queda un ítem
  anónimo que no se encoge por debajo de su palabra más larga.
- **Un panel que en escritorio es un cajón lateral se lleva su CSS puesto.** Los
  paneles de Filtros y Capas traen `position: fixed` y un `translateX(-105%)` de
  cuando están plegados: movidos al cajón de abajo quedaban invisibles y fuera de
  la pantalla (el botón "+ Capa" terminaba en x = -344 y no se podía tocar).
  `app-movil.css` lo neutraliza con `.app-hoja > .panel-*`; si aparece otro panel
  de ese tipo, hay que incluirlo ahí.
- **Los botones de acción de abajo no duplican lógica**: pulsan los botones que
  ya existen en la página, buscados por id. Como el editor avanzado los nombra
  distinto (`botonDescargarAvanzado`), el módulo prueba varios nombres.
- **`moveBefore` solo mueve entre nodos que ya están en la página.** El anuncio
  se mueve con `moveBefore` para que AdSense no lo recargue, pero la tira de
  arriba se agregaba a la página *después*: Edge y Chrome tiraban
  `HierarchyRequestError`, `encender()` se cortaba a la mitad y los cuatro
  editores quedaban en modo celular sin categorías ni botones (2026-10-04, en
  producción unas horas). `revisar` no lo vio porque mira las páginas sin
  contenido cargado; ahora lo cubre una prueba de `npm run probar` que carga
  una imagen a 390px.

---

## 7. Verificación

```
npm run revisar
```

Compila el sitio y abre **las 120 páginas** dos veces —en un iframe de 390px
(celular) y en uno de 1366px (escritorio)—, comprobando en cada una
que el encabezado esté visible y sin tapar, que la publicidad no ocupe el borde
superior, que nada desborde a lo ancho y que haya un `<h1>`. Las herramientas de
arquetipo C se revisan dos veces: también con su editor abierto, que es el estado
donde aparecieron los bugs. Cuando algo falla, dice **qué elemento** lo causa.

**No lo conectes a `npm run build`**: ese comando lo corre Cloudflare Pages, en
Linux y sin Edge, así que fallaría siempre y no se podría publicar.

```
npm run probar
```

Complementa a `revisar`: en vez de mirar cómo se ve, **usa** las herramientas,
sin internet y en tiempo real (`probar-funciones.mjs`). Comprueba que toda zona
de arrastre de las 240 páginas responde al teclado, que elegir dos veces el mismo
archivo y soltar funcionan, que las librerías de `vendor/` cargan de verdad
(Excel → PDF sin caracteres sustituidos, PDF.js con su worker, OCR completo,
fuentes propias), el buscador en los dos idiomas, y que ninguna página pide un
archivo que no existe. Cada prueba verifica una señal concreta, no solo que no
hubo error: varias herramientas tienen respaldos silenciosos que esconden fallas.
Tampoco va conectado al build.

Detalles de la técnica —por qué hace falta el iframe, cómo medir el DOM— en la
memoria `verificar-vistas-headless`.

### Lo que el verificador todavía no mira

Esto sigue siendo a ojo:

- [ ] Que el escenario no se cruce con la bandeja de controles en ninguna
      categoría ni proporción.
- [ ] Que cada categoría entre sin desplazamiento interno.
- [ ] Orientación horizontal y pantallas muy bajas.
- [ ] Cambiar idioma y tema, incluyendo los controles generados por JavaScript.
- [ ] Carga, error, procesamiento y descarga con contenido real.
- [ ] Que no se dupliquen nodos ni escuchadores al cambiar de tamaño.

Probá con el **emulador del navegador, no con un teléfono**: los gestos con mouse
en ventana angosta son donde aparecen los problemas de este patrón.

---

## 8. Agregar una pieza nueva

1. ¿Existe algo parecido? Reusalo. Una variante con un modificador es mejor que
   una pieza nueva.
2. Si es nueva, va en `estilos.css` (o en su propio par CSS+JS si lleva
   comportamiento), nunca en el `<style>` de una página.
3. Documentala en la sección 3 con su contrato: qué asume de vos y qué te
   garantiza. Las trampas de `.panel-lienzo` existieron porque ese contrato nunca
   se escribió.
4. Si se puede comprobar sola, agregá la comprobación a `revisar-interfaz.mjs`.

---

## 9. Deudas conocidas

- **No hay catálogo visual de piezas.** Una página que las muestre todas en los
  dos temas y los dos anchos haría esta guía comprobable en vez de aspiracional.
- **Migración a piezas (2026-10-01/02):** A (`catalogo.json`), B (zona de
  arrastre, selector de dirección, columnas, tablas, `escaparHtml`), C
  (`app-movil.*` + `editor-imagen.css`) y D (`miniaturas.js`,
  `.vista-previa-diapositivas`) ya comparten su código. Queda CSS propio en
  varias herramientas, pero ya no hay bloques copiados entre páginas: lo que
  queda es de cada una. El E también: `.tabla-creditos` y `.texto-actualizacion`.
- `assets/js/app-movil.js` conserva en `ICONOS` un mapa por texto visible, como
  respaldo de grupos sin `data-i18n`. Es frágil por idioma; se puede quitar
  cuando todos los grupos tengan clave.

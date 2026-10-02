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
| `.zona-arrastre` | caja de "arrastrá tu archivo acá"; lleva un `<input type="file" hidden>` |
| `.controles` | fila de controles; es flex y envuelve |
| `label.campo` | etiqueta + control, apilados |
| `.campo-deslizador` | variante con el valor a la derecha de la etiqueta |
| `.boton` | acción principal |
| `.boton-tema` | botón circular de icono (idioma, tema) |

> **Trampa documentada.** Un `<select>` toma el ancho de su opción más larga, y
> un hijo de un contenedor flex no se achica por debajo de su contenido. Por eso
> `.controles > *` y `label.campo` llevan `min-width: 0` y los controles
> `max-width: 100%`. Sin eso, una sola opción descriptiva ensancha la página
> entera en celular. Ya está resuelto: no lo reintroduzcas con estilos propios.

### Salida

| Pieza | Qué es |
|---|---|
| `.resultado` | caja de resultado; oculta hasta que se le agrega `.visible` |
| `.grid-miniaturas`, `.miniatura-pagina` | grilla de páginas/hojas seleccionables |
| `.contenedor-miniaturas`, `.selector-vista` | lista o grilla conmutable |
| `.lista-archivos` | lista de archivos cargados |
| `.tabla-ancha` | envoltorio para una tabla que no entra: se desplaza sola sin arrastrar la página |

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
  <div class="grid-herramientas">
    <a class="tarjeta-herramienta" href="herramientas/unir-pdf.html">…</a>
  </div>
</main>
```

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

### D — Vista previa de documento

Como B, con `.grid-miniaturas` / `.contenedor-miniaturas` dentro de
`.panel-herramienta`. Las miniaturas se generan con `assets/js/miniaturas.js`.

### E — Página de texto

`privacidad.html`, `terminos.html`, `creditos.html` y `contacto.html`: el marco
común y prosa dentro de un `.panel-herramienta`, sin controles. No llevan
`.aviso-privacidad` (eso es de los menús) pero sí su `.anuncio`.

Dos detalles propios: `creditos.html` envuelve su tabla en `.tabla-ancha`, que es
lo que evita que se lleve la página entera a lo ancho en un teléfono, y
`contacto.html` tiene un formulario.

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

Opcionales: `data-app-sin-descarga-flotante` en el contenedor si tu herramienta
ya tiene su propia categoría de descarga (evita dos botones que dicen lo mismo),
y `data-app-categorias-individuales` en un panel para que cada acordeón de
adentro sea su propia categoría en vez de viajar todo junto.

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

- **El encabezado tiene dos marcados distintos**: 76 herramientas usan dos `<a>`
  separados (logo y palabra) y 29 usan uno solo que envuelve ambos. Se ven igual
  porque el CSS los iguala, pero conviene unificarlos.
- **El verificador solo mira celular.** Faltan las comprobaciones de escritorio.
- **No hay catálogo visual de piezas.** Una página que las muestre todas en los
  dos temas y los dos anchos haría esta guía comprobable en vez de aspiracional.
- **Los arquetipos A, B, D y E no están migrados a piezas**: hoy son convención,
  no código compartido. Solo el C lo está. Es la tarea de prioridad 1 nº 3 en
  `../PENDIENTES.md`.
- `assets/js/app-movil.js` conserva en `ICONOS` un mapa por texto visible, como
  respaldo de grupos sin `data-i18n`. Es frágil por idioma; se puede quitar
  cuando todos los grupos tengan clave.

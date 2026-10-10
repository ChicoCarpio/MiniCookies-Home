# comparar-pptx — el motor de diapositivas contra PowerPoint real

`npm run comparar-pptx` dibuja cada `.pptx` con el motor del sitio y lo
compara, diapositiva por diapositiva, con lo que exporta PowerPoint (por COM).
Solo corre en Windows con PowerPoint y Edge instalados; no va en `npm run build`.
Detalle y opciones: el encabezado de `comparar-pptx.mjs`.

| Carpeta / archivo | Qué es | ¿En git? |
|---|---|---|
| `corpus\` | Casos sintéticos por característica (formas, texto, autoajuste, viñetas, rellenos, líneas, efectos, tablas con los 74 estilos, grupos, imágenes, SmartArt, gráficos, patrón, tamaños) | sí |
| `generar-corpus.ps1` | Los vuelve a generar con PowerPoint (`-Solo '^03'` para uno) | sí |
| `corpus\estilos-tabla-resultados.csv` | Los 74 GUID de estilos de tabla, todos aceptados por PowerPoint | sí |
| `locales\` | Presentaciones reales que no se pueden publicar (también se miden) | **no** |
| `powerpoint.ps1` | Referencia: PNG por diapositiva + `referencia.json` (caja de cada forma y de cada **línea de texto** como la cortó PowerPoint) | sí |
| `linea-base.json` | Resultado aceptado; el comando falla si una diapositiva empeora (> 1 punto de % o > 0,01 de SSIM) | sí |
| `salida\` | Referencias en caché (por SHA-256) e `informe.html` lado a lado | no |
| `ecma\presetShapeDefinitions.xml` | Las formas predefinidas tal como vienen en ECMA-376 Parte 1, 5.ª ed. (no se publica) | sí |
| `convertir-geometrias.mjs` | Las pasa a `vendor/ecma-376-geometrias/5ed-2016/geometrias.js` (solo cambia el formato) | sí |

Flujo de cada cambio en el motor: `npm run comparar-pptx`, revisar
`salida\actual\informe.html`, y si mejoró, `npm run comparar-pptx -- --base`.
`--sin-preparar` mide sin `pptx-preparar.js` (cuánto aporta la etapa 0).
Desde el 2026-10-10 el predeterminado es el motor propio (`pptx-modelo.js` y
compañía); el informe marca las diapositivas que todavía dibuja PptxViewJS
(tablas, gráficos) y por qué. `--pptxviewjs` mide la librería sola (salida en
`salida\pptxviewjs\`, no cambia la línea base). Con el motor propio el informe
compara además los **cortes de línea** con los de PowerPoint (las líneas de
cada forma, por nombre); ojo: la caja de línea de PowerPoint incluye el espacio
antes/después del párrafo y los espacios del final de la línea.

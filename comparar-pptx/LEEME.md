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
| `powerpoint.ps1` | Referencia: PNG por diapositiva + `referencia.json` (caja de cada forma, de cada **línea de texto** como la cortó PowerPoint y, en los gráficos, de cada parte: trazado, título, leyenda y sus entradas, ejes con su escala, puntos y etiquetas) | sí |
| `linea-base.json` | Resultado aceptado; el comando falla si una diapositiva empeora (> 1 punto de % o > 0,01 de SSIM) | sí |
| `salida\` | Referencias en caché (por SHA-256) e `informe.html` lado a lado | no |
| `ecma\presetShapeDefinitions.xml` | Las formas predefinidas tal como vienen en ECMA-376 Parte 1, 5.ª ed. (no se publica) | sí |
| `convertir-geometrias.mjs` | Las pasa a `vendor/ecma-376-geometrias/5ed-2016/geometrias.js` (solo cambia el formato) | sí |

Flujo de cada cambio en el motor: `npm run comparar-pptx`, revisar
`salida\actual\informe.html`, y si mejoró, `npm run comparar-pptx -- --base`.
Desde el 2026-10-10 el predeterminado es el motor propio (`pptx-modelo.js` y
compañía) y desde el 2026-10-11 el único (PptxViewJS y `pptx-preparar.js` se
quitaron; el informe marca lo que el motor omitió). Con el motor propio el informe
compara además los **cortes de línea** con los de PowerPoint (las líneas de
cada forma, por nombre); ojo: la caja de línea de PowerPoint incluye el espacio
antes/después del párrafo y los espacios del final de la línea.
Desde el 2026-10-10 compara también el texto de cada **celda de tabla**
(`Tabla [fila,col]`). PowerPoint da esas líneas relativas a la celda, y la
referencia (versión 2) trae el alto real de cada fila para pasarlas a la
diapositiva; aun así, PowerPoint mide la línea desde adentro del margen
superior de la celda: un desfase vertical de exactamente `marT` (3,6 pt por
defecto) en todas las celdas es esa convención, no un error de alto de fila.

Desde el 2026-10-10 compara también los **gráficos** (referencia versión 4):
por cada gráfico, la diferencia en pt de la caja del área de trazado, del
título y de la leyenda, si la escala de cada eje es la de PowerPoint, y la
mediana de la diferencia de las cajas de los puntos y de las etiquetas de
datos. `--detalle-graficos` lo muestra gráfico por gráfico. Ojo con lo que
da COM: el título, la leyenda, los ejes, el trazado y las etiquetas vienen
relativos al área del gráfico (4 pt adentro) y los puntos relativos a la
forma; `powerpoint.ps1` los deja todos relativos a la forma. En un circular
con un sector separado, las cajas de los sectores que da COM no coinciden
con lo dibujado: no usarlas para calibrar. `corpus-graficos-variantes.pptx`
tiene las variantes que cambian el diseño (leyenda en cada lado, sin
título, negativos, miles, categorías largas, eje secundario, circular con
porcentajes, dispersión con líneas, tamaños).

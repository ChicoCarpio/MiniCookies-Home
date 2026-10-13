# comparar-xlsx — el motor de Excel contra Excel real

Etapa 1 del motor propio de Excel (plan: `Investigaciones\investigacion-motor-excel\01-plan.md`).
Corre solo en Windows con Excel instalado, por COM; no va en `npm run build`.
Excel de referencia: Microsoft 365 16.0 (build 20326), interfaz en español (LCID 3082), 96 dpi.

| Archivo | Qué es |
|---|---|
| `medir-base.ps1` → `medidas\base.json` | Ancho de dígito (MDW), ancho/alto por defecto y alto de fila por fuente; orden de los colores de tema; TintAndShade |
| `casos-formato.json` + `medir-formatos.ps1` → `medidas\formatos.json` | Oráculo del formato de número: 226 casos, texto exacto (`Range.Text`) y color de fuente |
| `medidas\formatos-legible.txt` | Los mismos casos en una línea cada uno |
| `medidas\verificar-tint.cjs`, `probar-variantes-tint.cjs` | Comparan la fórmula de la norma con lo que da Excel |

`medir-formatos.ps1` mide con separadores en-US (`-Separadores sistema` usa los de Windows). Los nombres de color de los códigos
se traducen ([Red] → [Rojo]) porque este Excel solo acepta los de su idioma; en el archivo van siempre en inglés.
Los resultados dependen del idioma de Excel (nombres de día y mes, "p. m.", VERDADERO, `#¡DIV/0!`): el motor propio los saca de
su tabla es/en, no de esta medición.

## Etapa 4: geometría y dibujo en Canvas
| Archivo | Qué es |
|---|---|
| `medir-alto-fila.ps1` → `medidas\alto-fila.json` | AutoFit de una línea por fuente, tamaño y variante (solo vale de 13 pt para arriba) |
| `medir-alto-lineas.ps1` → `medidas\alto-lineas.json` | AutoFit de 2 y 3 líneas: la diferencia es el alto de una línea de esa fuente (de ahí sale la tabla de `xlsx-geometria.js`) |
| `volcar-geometria.ps1` → `medidas\geometria\*.json` | `Left` de cada columna y `Top` de cada fila de cada libro del corpus, como los calcula Excel |
| `comparar-geometria.mjs` | Posiciones del motor contra esos volcados (sin Excel; ±1 px; entra en `npm run comparar-xlsx`) |
| `exportar-rango.ps1` | Imagen de un rango como la pinta Excel en pantalla (`Range.CopyPicture`) |
| `casos-dibujo.json` + `comparar-dibujo.mjs` | Lo anterior en Edge: dibuja cada caso con `xlsx-dibujo.js`, lo pone al lado de la imagen de Excel y mide los píxeles distintos (tolerante al suavizado). Deja `[Excel \| motor \| diferencias]` en `medidas\tmp\dibujo\salida\`. También mide la geometría con ajuste de texto y el rendimiento con un libro de 1.000.000 de celdas (`generar-grande.py`). `node comparar-xlsx\comparar-dibujo.mjs 3 5` corre solo esos casos. Necesita Excel y Edge |
| `tinta-por-celda.mjs`, `desfase-dibujo.mjs`, `resumen-tinta.sh` | Calibración del texto: caja de tinta por celda en Excel y en el motor, desfase global, resumen |
| `png.mjs` | Lector/escritor mínimo de PNG para esas herramientas |
Corpus nuevo: `corpus\23-openpyxl-dibujo.xlsx` (fuentes y tamaños, estilos, tramas, bordes, alineación, rotación, ajuste de texto, números en columnas angostas). Solo archivos sintéticos.

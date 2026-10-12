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

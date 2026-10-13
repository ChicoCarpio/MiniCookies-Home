# Primera medición del motor de Excel (etapa 1): lo que la norma no dice y Excel sí.
#   - ancho de dígito (MDW) y ancho/alto por defecto según la fuente del estilo Normal
#   - ancho en puntos de columnas con ancho conocido (para verificar las fórmulas de la norma)
#   - alto de fila por fuente y tamaño
#   - color RGB de cada color de tema (orden de los índices) y de TintAndShade
# Uso: powershell -ExecutionPolicy Bypass -File medir-base.ps1  → medidas\base.json
param([string]$Salida = (Join-Path $PSScriptRoot 'medidas\base.json'))
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$res = [ordered]@{ version = $xl.Version; build = $xl.Build; pantalla = $null; fuentes = @(); tema = @(); tint = @() }
try {
  Add-Type -AssemblyName System.Windows.Forms
  $g = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
  $res.pantalla = [ordered]@{ dpiX = $g.DpiX; dpiY = $g.DpiY }
  $g.Dispose()

  $libro = $xl.Workbooks.Add()
  $hoja = $libro.Worksheets.Item(1)
  $normal = $libro.Styles.Item('Normal')
  $res.fuenteInicial = [ordered]@{ nombre = $normal.Font.Name; tam = $normal.Font.Size }

  $combos = @(
    @('Calibri', 11), @('Calibri', 12), @('Calibri', 10), @('Arial', 10), @('Arial', 11),
    @('Times New Roman', 11), @('Calibri Light', 11), @('Arial Narrow', 11), @('Aptos Narrow', 11), @('Aptos', 11), @('Carlito', 11), @('Liberation Sans', 10)
  )
  foreach ($c in $combos) {
    $normal.Font.Name = $c[0]; $normal.Font.Size = $c[1]
    $fila = [ordered]@{ fuente = $c[0]; tam = $c[1]; aplicada = $normal.Font.Name }
    $fila.anchoEstandar = $hoja.StandardWidth        # caracteres
    $fila.altoEstandar = $hoja.StandardHeight         # puntos
    $hoja.Range('A1').EntireColumn.ColumnWidth = [double]$hoja.StandardWidth
    $fila.anchoColumnaPuntos = $hoja.Range('A1').Width
    $fila.altoFilaPuntos = $hoja.Range('A1').Height
    $barrido = [ordered]@{}
    foreach ($n in 0, 1, 2, 5, 8, 8.43, 8.7109375, 9.140625, 10, 20, 50, 100, 255) {
      $hoja.Range('B1').EntireColumn.ColumnWidth = [double]$n
      $barrido["$n"] = [ordered]@{ ancho = $hoja.Range('B1').EntireColumn.ColumnWidth; puntos = $hoja.Range('B1').Width }
    }
    $fila.barrido = $barrido
    $res.fuentes += $fila
  }
  $normal.Font.Name = [string]$res.fuenteInicial.nombre; $normal.Font.Size = [double]$res.fuenteInicial.tam

  # Orden de los colores de tema: ThemeColor 1..12 → RGB (xlThemeColor*: 1 dk1? se mide, no se supone).
  foreach ($i in 1..12) {
    $r = $hoja.Range('D1'); $r.Interior.ThemeColor = $i; $r.Interior.TintAndShade = 0
    $res.tema += [ordered]@{ indice = $i; bgr = [int]$r.Interior.Color; rgb = ('{0:X6}' -f ((([int]$r.Interior.Color -band 0xFF) -shl 16) -bor ([int]$r.Interior.Color -band 0xFF00) -bor (([int]$r.Interior.Color -shr 16) -band 0xFF))) }
  }
  # TintAndShade en pasos de 0,1 para los 12 colores (para comparar con la fórmula de la norma).
  foreach ($i in 1..12) {
    foreach ($t in -1.0, -0.75, -0.5, -0.25, -0.1, 0.1, 0.25, 0.5, 0.75, 1.0) {
      $r = $hoja.Range('D1'); $r.Interior.ThemeColor = $i; $r.Interior.TintAndShade = $t
      $c = [int]$r.Interior.Color
      $res.tint += [ordered]@{ tema = $i; tint = $t; rgb = ('{0:X6}' -f ((($c -band 0xFF) -shl 16) -bor ($c -band 0xFF00) -bor (($c -shr 16) -band 0xFF))) }
    }
  }
  $libro.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
New-Item -ItemType Directory -Force (Split-Path $Salida) | Out-Null
$res | ConvertTo-Json -Depth 8 | Set-Content -Encoding UTF8 $Salida
"OK → $Salida"

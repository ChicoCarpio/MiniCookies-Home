# Ancho y alto IMPRESOS de columnas y filas (etapa 1).
# En pantalla una columna de 8.43 caracteres mide 48 pt (64 px), pero en la paginación entran menos columnas
# de las que salen de ese cálculo. Aquí se halla el ancho de página mínimo (moviendo el margen derecho, o el inferior
# para las filas) con el que caben k columnas / filas, y de ahí el ancho impreso de cada una.
# Uso: powershell -ExecutionPolicy Bypass -File medir-ancho-impreso.ps1  → medidas\ancho-impreso.json
param([string]$Salida = (Join-Path $PSScriptRoot 'medidas\ancho-impreso.json'))
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$res = [ordered]@{ version = $xl.Version; impresora = $xl.ActivePrinter; columnas = @(); filas = @() }

# ¿Cuántas columnas hay en la primera página? (primer salto vertical − 1; sin saltos = todas)
function Primera-Pagina-Cols($hoja) {
  $null = $hoja.PageSetup.Pages.Count   # fuerza el cálculo de la paginación: sin esto los saltos salen vacíos
  if ($hoja.VPageBreaks.Count -gt 0) { return $hoja.VPageBreaks.Item(1).Location.Column - 1 }
  return 9999
}
function Primera-Pagina-Filas($hoja) {
  $null = $hoja.PageSetup.Pages.Count
  if ($hoja.HPageBreaks.Count -gt 0) { return $hoja.HPageBreaks.Item(1).Location.Row - 1 }
  return 9999
}

try {
  $libro = $xl.Workbooks.Add()
  $hoja = $libro.Worksheets.Item(1)
  $ps = $hoja.PageSetup
  $ps.PaperSize = 9; $ps.Orientation = 1; $ps.Zoom = 100
  $ps.LeftMargin = 0; $ps.TopMargin = 0; $ps.HeaderMargin = 0; $ps.FooterMargin = 0
  $anchoPagina = 595.3; $altoPagina = 841.9   # A4 en puntos (se comprueba abajo con el salto de página)

  foreach ($chars in 1, 2, 5, 8.43, 10, 20, 40) {
    $hoja.Cells.Clear() | Out-Null
    $hoja.Range('A1:GZ1').Value2 = 'x'
    $hoja.Range('A1:GZ1').EntireColumn.ColumnWidth = [double]$chars
    $pantalla = $hoja.Range('A1').Width
    $fila = [ordered]@{ caracteres = $chars; pantallaPuntos = $pantalla }
    foreach ($k in 1, 5, 12) {
      # Margen derecho de 0 a (ancho de página): a más margen, menos columnas entran. Se busca el mayor margen con k columnas.
      $lo = 0.0; $hi = $anchoPagina - 1.0
      for ($i = 0; $i -lt 22; $i++) {
        $m = ($lo + $hi) / 2
        $ps.RightMargin = $m
        $n = Primera-Pagina-Cols $hoja
        if ($n -ge $k) { $lo = $m } else { $hi = $m }
      }
      $fila["k$k"] = [Math]::Round($anchoPagina - $lo, 2)   # ancho útil mínimo para k columnas
    }
    $res.columnas += $fila
  }
  $ps.RightMargin = 0
  foreach ($alto in 6, 15, 30, 60) {
    $hoja.Cells.Clear() | Out-Null
    $hoja.Range('A1:A300').Value2 = 'x'
    $hoja.Range('A1:A300').EntireRow.RowHeight = [double]$alto
    $fila = [ordered]@{ altoPuntos = $alto; pantallaPuntos = $hoja.Range('A1').Height }
    foreach ($k in 1, 5, 12) {
      $lo = 0.0; $hi = $altoPagina - 1.0
      for ($i = 0; $i -lt 22; $i++) {
        $m = ($lo + $hi) / 2
        $ps.BottomMargin = $m
        $n = Primera-Pagina-Filas $hoja
        if ($n -ge $k) { $lo = $m } else { $hi = $m }
      }
      $fila["k$k"] = [Math]::Round($altoPagina - $lo, 2)
    }
    $res.filas += $fila
  }
  $libro.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
New-Item -ItemType Directory -Force (Split-Path $Salida) | Out-Null
$res | ConvertTo-Json -Depth 8 | Set-Content -Encoding UTF8 $Salida
"OK → $Salida"

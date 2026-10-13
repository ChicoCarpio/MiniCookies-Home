# Alto de fila automatico de una celda con 2 y 3 lineas (ajuste de texto) por fuente, tamano y variante (n, b, i, bi).
# La diferencia entre 3 y 2 lineas es el alto de una linea de esa fuente (el que Excel da a una fila de texto al abrir un libro sin altos
# guardados); no se usa el alto estandar del estilo Normal porque Excel no lo actualiza de inmediato al cambiarlo por COM.
# Un libro nuevo por fuente: Excel deja de aceptar formatos tras ~2000 combinaciones.
# Uso: powershell -ExecutionPolicy Bypass -File medir-alto-lineas.ps1 [-Salida medidas\alto-lineas.json]
param([string]$Salida = (Join-Path $PSScriptRoot 'medidas\alto-lineas.json'))
$ErrorActionPreference = 'Stop'
$fuentes = @('Calibri', 'Calibri Light', 'Arial', 'Times New Roman', 'Courier New', 'Verdana', 'Tahoma', 'Segoe UI', 'Consolas', 'Cambria', 'Georgia', 'Arial Narrow', 'Comic Sans MS', 'Trebuchet MS')
$variantes = @('n', 'b', 'i', 'bi')
$tamanos = @(6, 7, 8, 9, 10, 10.5, 11, 12, 13, 14, 15, 16, 18, 20, 22, 24, 26, 28, 32, 36, 40)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$resultado = [ordered]@{}
try {
  foreach ($f in $fuentes) {
    $l = $xl.Workbooks.Add(); $h = $l.Worksheets.Item(1)
    $h.Columns.Item(1).ColumnWidth = 255
    $fila = 1
    $porVariante = [ordered]@{}
    foreach ($v in $variantes) {
      $mapa = [ordered]@{}
      foreach ($t in $tamanos) {
        $alturas = @()
        foreach ($n in 2..3) {
          $c = $h.Cells.Item($fila, 1)
          $c.Value2 = (@('Xy') * $n) -join "`n"
          $c.WrapText = $true
          $c.Font.Name = $f; $c.Font.Size = [double]$t
          $c.Font.Bold = [bool]($v -match 'b'); $c.Font.Italic = [bool]($v -match 'i')
          $h.Rows.Item($fila).AutoFit() | Out-Null
          $alturas += [double]$h.Rows.Item($fila).RowHeight
          $fila++
        }
        $mapa[[string]$t] = $alturas
      }
      $porVariante[$v] = $mapa
    }
    $resultado[$f] = $porVariante
    $l.Close($false)
  }
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
$json = $resultado | ConvertTo-Json -Depth 5
[IO.File]::WriteAllText($Salida, $json, (New-Object Text.UTF8Encoding($true)))
"OK -> $Salida"

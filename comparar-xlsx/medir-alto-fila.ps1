# Alto de fila automatico (AutoFit) que Excel da a una fila con una sola celda de cada fuente, tamano y variante (n, b, i, bi).
# Hasta 12 pt el AutoFit no baja del alto estandar, asi que ahi se mide el alto estandar con esa fuente como estilo Normal.
# Estilo Normal de partida: Calibri 11 (MDW 7). Uso: powershell -ExecutionPolicy Bypass -File medir-alto-fila.ps1 [-Salida medidas\alto-fila.json]
param([string]$Salida = (Join-Path $PSScriptRoot 'medidas\alto-fila.json'))
$ErrorActionPreference = 'Stop'
$fuentes = @('Calibri', 'Calibri Light', 'Arial', 'Times New Roman', 'Courier New', 'Verdana', 'Tahoma', 'Segoe UI', 'Consolas', 'Cambria', 'Georgia', 'Arial Narrow', 'Comic Sans MS', 'Trebuchet MS')
$variantes = @('n', 'b', 'i', 'bi')
$tamanos = @(6, 7, 8, 9, 10, 10.5, 11, 12, 13, 14, 15, 16, 18, 20, 22, 24, 26, 28, 32, 36, 40, 48, 60, 72)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$resultado = [ordered]@{}
try {
  $l = $xl.Workbooks.Add()
  $h = $l.Worksheets.Item(1)
  $fila = 1
  foreach ($f in $fuentes) {
    $porVariante = [ordered]@{}
    foreach ($v in $variantes) {
      $mapa = [ordered]@{}
      foreach ($t in ($tamanos | Where-Object { $_ -gt 12 })) {
        $c = $h.Cells.Item($fila, 1)
        $c.Value2 = 'Xy'
        $c.Font.Name = $f
        $c.Font.Size = [double]$t
        $c.Font.Bold = ($v -match 'b')
        $c.Font.Italic = ($v -match 'i')
        $h.Rows.Item($fila).AutoFit() | Out-Null
        $mapa[[string]$t] = [double]$h.Rows.Item($fila).RowHeight
        $fila++
      }
      $porVariante[$v] = $mapa
    }
    $resultado[$f] = $porVariante
  }
  $normal = [ordered]@{}
  foreach ($f in $fuentes) {
    $porVariante = [ordered]@{}
    foreach ($v in $variantes) {
      $mapa = [ordered]@{}
      foreach ($t in ($tamanos | Where-Object { $_ -le 12 })) {
        $st = $l.Styles.Item('Normal')
        $st.Font.Name = $f
        $st.Font.Size = [double]$t
        $st.Font.Bold = ($v -match 'b')
        $st.Font.Italic = ($v -match 'i')
        $hn = $l.Worksheets.Add()
        $mapa[[string]$t] = [double]$hn.StandardHeight
        $hn.Delete()
      }
      $porVariante[$v] = $mapa
    }
    $normal[$f] = $porVariante
  }
  $resultado['__normal'] = $normal
  $l.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
$json = $resultado | ConvertTo-Json -Depth 5
[IO.File]::WriteAllText($Salida, $json, (New-Object Text.UTF8Encoding($true)))
"OK -> $Salida"

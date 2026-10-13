# Posicion (Left/Top, en puntos) de cada columna y fila de cada hoja visible, tal como Excel la calcula. Referencia de la geometria (etapa 4).
# Uso: powershell -ExecutionPolicy Bypass -File volcar-geometria.ps1 -Libro ruta.xlsx -Salida geometria.json [-Columnas 40] [-Filas 80]
param([Parameter(Mandatory)][string]$Libro, [Parameter(Mandatory)][string]$Salida, [int]$Columnas = 40, [int]$Filas = 80)
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$hojas = @()
try {
  $l = $xl.Workbooks.Open((Resolve-Path $Libro).Path, 0, $true)
  for ($i = 1; $i -le $l.Worksheets.Count; $i++) {
    $h = $l.Worksheets.Item($i)
    $u = $h.UsedRange
    $nc = [Math]::Min($Columnas, [Math]::Max(1, $u.Column + $u.Columns.Count + 2))
    $nf = [Math]::Min($Filas, [Math]::Max(1, $u.Row + $u.Rows.Count + 2))
    $cols = New-Object System.Collections.ArrayList
    for ($c = 1; $c -le ($nc + 1); $c++) { [void]$cols.Add([double]$h.Cells.Item(1, $c).Left) }
    $fil = New-Object System.Collections.ArrayList
    for ($f = 1; $f -le ($nf + 1); $f++) { [void]$fil.Add([double]$h.Cells.Item($f, 1).Top) }
    $hojas += [ordered]@{ indice = $i - 1; nombre = $h.Name; visible = ($h.Visible -eq -1); columnas = $cols; filas = $fil }
  }
  $l.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
$json = ConvertTo-Json -InputObject $hojas -Depth 5
[IO.File]::WriteAllText($Salida, $json, (New-Object Text.UTF8Encoding($true)))
"OK -> $Salida"

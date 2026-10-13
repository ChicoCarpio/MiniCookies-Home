# Volcado de las fórmulas de un libro tal como las ve Excel (Formula2), para compararlas con las del lector propio
# (sobre todo las compartidas, que el archivo guarda una sola vez). Salida: una línea por fórmula "índiceHoja<TAB>dirección<TAB>fórmula".
# Uso: powershell -ExecutionPolicy Bypass -File volcar-formulas.ps1 -Libro ruta.xlsx -Salida salida.tsv
param([Parameter(Mandatory)][string]$Libro, [Parameter(Mandatory)][string]$Salida)
$ErrorActionPreference = 'Stop'
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('en-US')
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
$lineas = New-Object System.Collections.Generic.List[string]
try {
  $l = $xl.Workbooks.Open((Resolve-Path $Libro).Path, 0, $true)
  for ($i = 1; $i -le $l.Worksheets.Count; $i++) {
    $h = $l.Worksheets.Item($i)
    try { $rng = $h.UsedRange.SpecialCells(-4123) } catch { continue }   # xlCellTypeFormulas
    foreach ($area in $rng.Areas) {
      $f = $area.Formula2
      if ($f -isnot [object[,]]) { $lineas.Add("$($i - 1)`t$($area.Address($false, $false))`t$f"); continue }
      $f0 = $area.Row; $c0 = $area.Column
      for ($r = 1; $r -le $area.Rows.Count; $r++) { for ($c = 1; $c -le $area.Columns.Count; $c++) {
        $t = $f[$r, $c]
        if ($t -is [string] -and $t.StartsWith('=')) { $lineas.Add("$($i - 1)`t$($h.Cells.Item($f0 + $r - 1, $c0 + $c - 1).Address($false, $false))`t$($t.Replace("`r", '').Replace("`n", '\n'))") }
      } }
    }
  }
  $l.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
[IO.File]::WriteAllLines($Salida, $lineas, (New-Object Text.UTF8Encoding($false)))
"OK: $($lineas.Count) fórmulas → $Salida"

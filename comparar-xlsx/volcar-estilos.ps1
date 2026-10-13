# Estilo de las celdas tal como lo ve Excel (fuente, relleno, formato de número, alineación, bordes), para compararlo
# con el que resuelve el lector propio (comparar-estilos.mjs). Recorre las celdas con contenido o estilo propio
# de las primeras filas de cada hoja visible.
# Uso: powershell -ExecutionPolicy Bypass -File volcar-estilos.ps1 -Libro ruta.xlsx -Salida salida.json [-MaxCeldas 150] [-MaxHojas 6]
param([Parameter(Mandatory)][string]$Libro, [Parameter(Mandatory)][string]$Salida, [int]$MaxCeldas = 150, [int]$MaxHojas = 6)
$ErrorActionPreference = 'Stop'
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('en-US')
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
$usarSistema = $xl.UseSystemSeparators; $decOrig = $xl.DecimalSeparator; $milOrig = $xl.ThousandsSeparator
$xl.UseSystemSeparators = $false; $xl.DecimalSeparator = '.'; $xl.ThousandsSeparator = ','
$celdas = New-Object System.Collections.Generic.List[object]
function Borde($c, $i) { $b = $c.Borders.Item($i); if ($b.LineStyle -eq -4142) { return $null } return [ordered]@{ estilo = [int]$b.LineStyle; peso = [int]$b.Weight; color = [int]$b.Color } }
try {
  $l = $xl.Workbooks.Open((Resolve-Path $Libro).Path, 0, $true)
  $hechas = 0
  for ($i = 1; $i -le $l.Worksheets.Count -and $hechas -lt $MaxHojas; $i++) {
    $h = $l.Worksheets.Item($i)
    if ($h.Visible -ne -1) { continue }
    $hechas++
    $usado = $h.UsedRange
    $n = 0
    foreach ($c in $usado.Cells) {
      if ($n -ge $MaxCeldas) { break }
      if ($c.MergeCells -and $c.Address() -ne $c.MergeArea.Cells.Item(1, 1).Address()) { continue }
      $n++
      $celdas.Add([ordered]@{
        hoja = $i - 1; dir = $c.Address($false, $false).Replace('$', '')
        fuente = [ordered]@{ nombre = [string]$c.Font.Name; tam = [double]$c.Font.Size; negrita = [bool]$c.Font.Bold; cursiva = [bool]$c.Font.Italic; subrayado = [int]$c.Font.Underline; tachado = [bool]$c.Font.Strikethrough; color = [int]$c.Font.Color }
        relleno = [ordered]@{ patron = [int]$c.Interior.Pattern; color = [int]$c.Interior.Color; colorPatron = [int]$c.Interior.PatternColor }
        formato = [string]$c.NumberFormat
        alineacion = [ordered]@{ h = [int]$c.HorizontalAlignment; v = [int]$c.VerticalAlignment; ajustar = [bool]$c.WrapText; sangria = [int]$c.IndentLevel; rotacion = [int]$c.Orientation; reducir = [bool]$c.ShrinkToFit }
        bordes = [ordered]@{ izq = (Borde $c 7); arr = (Borde $c 8); aba = (Borde $c 9); der = (Borde $c 10) }
        texto = [string]$c.Text
      })
    }
  }
  $l.Close($false)
} finally { $xl.DecimalSeparator = $decOrig; $xl.ThousandsSeparator = $milOrig; $xl.UseSystemSeparators = $usarSistema; $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
$celdas | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 $Salida
"OK: $($celdas.Count) celdas → $Salida"

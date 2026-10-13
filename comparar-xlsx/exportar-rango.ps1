# Imagen de un rango tal como Excel lo dibuja en pantalla (referencia para el dibujo en Canvas, etapa 4).
# Usa Range.CopyPicture y un gráfico auxiliar para guardar el PNG. Con -ConCuadricula Excel dibuja la cuadrícula.
# Uso: powershell -ExecutionPolicy Bypass -File exportar-rango.ps1 -Libro ruta.xlsx -Hoja 1 -Rango A1:H20 -Salida salida.png [-Zoom 100]
param([Parameter(Mandatory)][string]$Libro, [int]$Hoja = 1, [Parameter(Mandatory)][string]$Rango, [Parameter(Mandatory)][string]$Salida, [int]$Zoom = 100, [switch]$ConCuadricula)
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true; $xl.DisplayAlerts = $false
try {
  $l = $xl.Workbooks.Open((Resolve-Path $Libro).Path, 0, $true)
  $h = $l.Worksheets.Item($Hoja)
  $h.Activate()
  $xl.ActiveWindow.Zoom = $Zoom
  $xl.ActiveWindow.DisplayGridlines = [bool]$ConCuadricula
  $r = $h.Range($Rango)
  $r.CopyPicture(1, 2) | Out-Null      # xlScreen, xlBitmap
  $ancho = [double]$r.Width; $alto = [double]$r.Height
  $g = $h.ChartObjects().Add(0, 0, $ancho + 2, $alto + 2)
  $g.Activate() | Out-Null
  $g.Chart.ChartArea.Format.Line.Visible = 0
  $g.Chart.Paste() | Out-Null
  $g.Chart.Export((Join-Path (Resolve-Path (Split-Path $Salida)).Path (Split-Path $Salida -Leaf))) | Out-Null
  $g.Delete() | Out-Null
  $l.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
"OK → $Salida"

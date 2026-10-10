# Referencia de PowerPoint real para comparar-pptx.mjs: abre un .pptx en solo
# lectura (sin ventana), exporta cada diapositiva a PNG y escribe
# referencia.json con el tamaño de la diapositiva, la caja de cada forma y el
# texto y la caja de cada LÍNEA de texto tal como las cortó PowerPoint (la
# verdad para comparar el corte de líneas del motor propio). Con -Pdf exporta
# además el PDF de PowerPoint.
#
# Uso: powershell -ExecutionPolicy Bypass -File powerpoint.ps1 -Archivo x.pptx -Salida carpeta [-Pdf]
# Necesita Microsoft PowerPoint instalado. No modifica el archivo.
param(
  [Parameter(Mandatory = $true)][string]$Archivo,
  [Parameter(Mandatory = $true)][string]$Salida,
  [switch]$Pdf,
  [int]$Ancho = 1920
)
$ErrorActionPreference = 'Stop'
$Archivo = (Resolve-Path -LiteralPath $Archivo).Path
New-Item -ItemType Directory -Force -Path $Salida | Out-Null
$Salida = (Resolve-Path -LiteralPath $Salida).Path

# Si PowerPoint ya estaba abierto (por el usuario), no se cierra al final.
$yaEstaba = @(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
$app = New-Object -ComObject PowerPoint.Application
$p = $null

function Caja($o) {
  [ordered]@{ x = [math]::Round($o.Left, 2); y = [math]::Round($o.Top, 2); w = [math]::Round($o.Width, 2); h = [math]::Round($o.Height, 2) }
}

# Las líneas de un TextRange2, con el texto y la caja que PowerPoint les dio.
function Lineas($tr) {
  $lista = @()
  $n = $tr.Lines().Count
  for ($i = 1; $i -le $n; $i++) {
    $l = $tr.Lines($i, 1)
    $lista += [ordered]@{
      texto = $l.Text -replace "[\r\v\n]+$", ''
      x = [math]::Round($l.BoundLeft, 2); y = [math]::Round($l.BoundTop, 2)
      w = [math]::Round($l.BoundWidth, 2); h = [math]::Round($l.BoundHeight, 2)
    }
  }
  , $lista
}

# Lo que PowerPoint calculó de un gráfico, todo en puntos relativos a la
# esquina de la forma: área de trazado interior, título, leyenda, cada eje
# (caja y escala) y la caja de cada punto y de su etiqueta. La verdad para
# comparar el diseño automático del motor propio. Ojo (comprobado contra los
# píxeles de la exportación): COM da los puntos relativos a la forma, pero el
# título, la leyenda, los ejes, el trazado y las etiquetas relativos al área
# del gráfico (ChartArea.Left/Top, 4 pt): se les suma para que todo quede igual.
function Grafico($ch) {
  $dx = [double]$ch.ChartArea.Left; $dy = [double]$ch.ChartArea.Top
  $r2 = { param($v) [math]::Round([double]$v, 2) }
  $cj = { param($x, $y, $w, $h) [ordered]@{ x = (& $r2 ($x + $dx)); y = (& $r2 ($y + $dy)); w = (& $r2 $w); h = (& $r2 $h) } }
  $g = [ordered]@{ tipo = [int]$ch.ChartType }
  $pa = $ch.PlotArea
  $g.trazado = (& $cj $pa.InsideLeft $pa.InsideTop $pa.InsideWidth $pa.InsideHeight)
  if ($ch.HasTitle) { $t = $ch.ChartTitle; $g.titulo = (& $cj $t.Left $t.Top $t.Width $t.Height) }
  if ($ch.HasLegend) {
    $l = $ch.Legend; $g.leyenda = (& $cj $l.Left $l.Top $l.Width $l.Height)
    # Cada entrada y su muestra de color (para calibrar los anchos de la leyenda).
    $g.entradas = @()
    try {
      $n = $l.LegendEntries().Count
      for ($i = 1; $i -le $n; $i++) {
        $en = $l.LegendEntries($i); $k = $en.LegendKey
        $g.entradas += [ordered]@{ caja = (& $cj $en.Left $en.Top $en.Width $en.Height); muestra = (& $cj $k.Left $k.Top $k.Width $k.Height) }
      }
    } catch { }
  }
  $g.ejes = @()
  foreach ($grupo in 1, 2) {
    foreach ($tipo in 1, 2) {
      try { if (-not $ch.HasAxis($tipo, $grupo)) { continue }; $e = $ch.Axes($tipo, $grupo) } catch { continue }
      $eje = [ordered]@{ tipo = $tipo; grupo = $grupo; caja = (& $cj $e.Left $e.Top $e.Width $e.Height) }
      try { $eje.min = $e.MinimumScale; $eje.max = $e.MaximumScale; $eje.unidad = $e.MajorUnit } catch { }
      $g.ejes += $eje
    }
  }
  $g.series = @()
  $n = $ch.SeriesCollection().Count
  for ($k = 1; $k -le $n; $k++) {
    $se = $ch.SeriesCollection($k)
    $pts = @()
    $np = 0; try { $np = $se.Points().Count } catch { }
    for ($i = 1; $i -le $np; $i++) {
      $p = $se.Points($i)
      $q = [ordered]@{}
      try { $q.caja = (Caja $p) } catch { }
      if ($p.HasDataLabel) { $l = $p.DataLabel; $q.etiqueta = (& $cj $l.Left $l.Top $l.Width $l.Height); $q.etiqueta.texto = $l.Text }
      $pts += $q
    }
    $g.series += [ordered]@{ nombre = $se.Name; puntos = $pts }
  }
  $g
}

function Forma($s) {
  $f = [ordered]@{ nombre = $s.Name; tipo = [int]$s.Type; caja = (Caja $s); giro = [math]::Round($s.Rotation, 2) }
  try {
    if ($s.Type -eq 6) {
      # Grupo: sus formas, con su caja ya en coordenadas de la diapositiva.
      $f.hijos = @(); foreach ($h in $s.GroupItems) { $f.hijos += (Forma $h) }
    }
    if ($s.HasTextFrame -and $s.TextFrame2.HasText) { $f.lineas = (Lineas $s.TextFrame2.TextRange) }
    if ($s.HasChart) { $f.grafico = (Grafico $s.Chart) }
    if ($s.HasTable) {
      # Alto real de cada fila (crece con el texto) y ancho de cada columna, en
      # puntos: con ellos las líneas de una celda (relativas a la celda) pasan
      # a la diapositiva.
      $t = $s.Table; $f.celdas = @(); $f.filas = @(); $f.columnas = @()
      for ($r = 1; $r -le $t.Rows.Count; $r++) { $f.filas += [math]::Round($t.Rows($r).Height, 2) }
      for ($c = 1; $c -le $t.Columns.Count; $c++) { $f.columnas += [math]::Round($t.Columns($c).Width, 2) }
      for ($r = 1; $r -le $t.Rows.Count; $r++) {
        for ($c = 1; $c -le $t.Columns.Count; $c++) {
          $tf = $t.Cell($r, $c).Shape.TextFrame2
          if ($tf.HasText) { $f.celdas += [ordered]@{ fila = $r; col = $c; lineas = (Lineas $tf.TextRange) } }
        }
      }
    }
  } catch { $f.error = $_.Exception.Message }
  $f
}

try {
  # Open(ruta, ReadOnly, Untitled, WithWindow)
  $p = $app.Presentations.Open($Archivo, -1, 0, 0)
  $anchoPt = $p.PageSetup.SlideWidth; $altoPt = $p.PageSetup.SlideHeight
  $alto = [int][math]::Round($Ancho * $altoPt / $anchoPt)
  $ref = [ordered]@{
    version = 4
    archivo = [IO.Path]::GetFileName($Archivo)
    ancho = $anchoPt; alto = $altoPt; pxAncho = $Ancho; pxAlto = $alto
    powerpoint = $app.Version
    diapositivas = @()
  }
  for ($i = 1; $i -le $p.Slides.Count; $i++) {
    $d = $p.Slides($i)
    $png = Join-Path $Salida ('d{0:000}.png' -f $i)
    $d.Export($png, 'PNG', $Ancho, $alto)
    $formas = @(); foreach ($s in $d.Shapes) { $formas += (Forma $s) }
    $ref.diapositivas += [ordered]@{ n = $i; oculta = [bool]$d.SlideShowTransition.Hidden; formas = $formas }
  }
  if ($Pdf) { $p.SaveCopyAs((Join-Path $Salida 'powerpoint.pdf'), 32) }  # ppSaveAsPDF
  $json = $ref | ConvertTo-Json -Depth 12
  [IO.File]::WriteAllText((Join-Path $Salida 'referencia.json'), $json, (New-Object Text.UTF8Encoding $false))
  "OK $($p.Slides.Count) diapositivas"
} finally {
  if ($p) { $p.Close() }
  if (-not $yaEstaba -and $app.Presentations.Count -eq 0) { $app.Quit() }
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}

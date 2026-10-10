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

function Forma($s) {
  $f = [ordered]@{ nombre = $s.Name; tipo = [int]$s.Type; caja = (Caja $s); giro = [math]::Round($s.Rotation, 2) }
  try {
    if ($s.Type -eq 6) {
      # Grupo: sus formas, con su caja ya en coordenadas de la diapositiva.
      $f.hijos = @(); foreach ($h in $s.GroupItems) { $f.hijos += (Forma $h) }
    }
    if ($s.HasTextFrame -and $s.TextFrame2.HasText) { $f.lineas = (Lineas $s.TextFrame2.TextRange) }
    if ($s.HasTable) {
      $t = $s.Table; $f.celdas = @()
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

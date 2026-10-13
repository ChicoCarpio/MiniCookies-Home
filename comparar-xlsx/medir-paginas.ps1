# Paginación de Excel (etapa 1): dónde corta las páginas y qué factor usa en "ajustar a".
# Para cada configuración arma una hoja con filas de altos conocidos y columnas de anchos conocidos y lee
# HPageBreaks / VPageBreaks (posición y tipo), Pages.Count y PageSetup.Zoom.
# Uso: powershell -ExecutionPolicy Bypass -File medir-paginas.ps1  → medidas\paginas.json
param([string]$Salida = (Join-Path $PSScriptRoot 'medidas\paginas.json'))
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$res = [ordered]@{ version = $xl.Version; impresora = $xl.ActivePrinter; pruebas = @() }

function Leer-Saltos($hoja) {
  $h = @(); $v = @()
  $null = $hoja.PageSetup.Pages.Count
  for ($i = 1; $i -le $hoja.HPageBreaks.Count; $i++) { try { $b = $hoja.HPageBreaks.Item($i); $h += [ordered]@{ fila = $b.Location.Row; tipo = $b.Type } } catch { $h += [ordered]@{ fila = $null; tipo = $null } } }
  for ($i = 1; $i -le $hoja.VPageBreaks.Count; $i++) { try { $b = $hoja.VPageBreaks.Item($i); $v += [ordered]@{ col = $b.Location.Column; tipo = $b.Type } } catch { $v += [ordered]@{ col = $null; tipo = $null } } }
  return @{ h = $h; v = $v }
}

try {
  $libro = $xl.Workbooks.Add()
  $hoja = $libro.Worksheets.Item(1)
  # 120 filas de 15 pt y 24 columnas de 64 px (8.43), con un valor en cada celda.
  $hoja.Range('A1:X120').Value2 = 'x'
  $hoja.Range('A1:X120').Font.Size = 11
  $hoja.Range('A1:X120').EntireColumn.ColumnWidth = [double]8.43
  $hoja.Range('A1:X120').EntireRow.RowHeight = [double]15
  $res.impresora = $xl.ActivePrinter
  $tmp = Join-Path $env:TEMP 'mc-xlsx'; New-Item -ItemType Directory -Force $tmp | Out-Null

  $configs = @(
    @{ nombre = 'carta-vertical-100'; papel = 1; orient = 1; zoom = 100 },
    @{ nombre = 'a4-vertical-100'; papel = 9; orient = 1; zoom = 100 },
    @{ nombre = 'a4-horizontal-100'; papel = 9; orient = 2; zoom = 100 },
    @{ nombre = 'a4-vertical-50'; papel = 9; orient = 1; zoom = 50 },
    @{ nombre = 'a4-vertical-150'; papel = 9; orient = 1; zoom = 150 },
    @{ nombre = 'a4-vertical-85'; papel = 9; orient = 1; zoom = 85 },
    @{ nombre = 'a4-vertical-ajuste-1x0'; papel = 9; orient = 1; ancho = 1; alto = 0 },
    @{ nombre = 'a4-vertical-ajuste-1x1'; papel = 9; orient = 1; ancho = 1; alto = 1 },
    @{ nombre = 'a4-vertical-ajuste-2x2'; papel = 9; orient = 1; ancho = 2; alto = 2 },
    @{ nombre = 'a4-vertical-ajuste-0x2'; papel = 9; orient = 1; ancho = 0; alto = 2 },
    @{ nombre = 'carta-horizontal-ajuste-3x0'; papel = 1; orient = 2; ancho = 3; alto = 0 }
  )
  foreach ($c in $configs) {
    $ps = $hoja.PageSetup
    try {
      $ps.PaperSize = $c.papel; $ps.Orientation = $c.orient
      $ps.LeftMargin = $xl.InchesToPoints(0.7); $ps.RightMargin = $xl.InchesToPoints(0.7)
      $ps.TopMargin = $xl.InchesToPoints(0.75); $ps.BottomMargin = $xl.InchesToPoints(0.75)
      $ps.HeaderMargin = $xl.InchesToPoints(0.3); $ps.FooterMargin = $xl.InchesToPoints(0.3)
      if ($c.zoom) { $ps.Zoom = [int]$c.zoom }
      else {
        $ps.Zoom = $false
        if ($c.ancho -eq 0) { $ps.FitToPagesWide = $false } else { $ps.FitToPagesWide = [int]$c.ancho }
        if ($c.alto -eq 0) { $ps.FitToPagesTall = $false } else { $ps.FitToPagesTall = [int]$c.alto }
      }
      $paginas = $hoja.PageSetup.Pages.Count
      $s = Leer-Saltos $hoja
      $zoomLeido = $ps.Zoom
      # El oráculo final: el PDF que exporta Excel (páginas y tamaño de papel).
      $pdf = Join-Path $tmp ($c.nombre + '.pdf')
      $hoja.ExportAsFixedFormat(0, $pdf)
      $texto = [Text.Encoding]::GetEncoding(28591).GetString([IO.File]::ReadAllBytes($pdf))
      $paginasPdf = ([regex]::Matches($texto, '/Type\s*/Page[^s]')).Count
      $caja = [regex]::Match($texto, '/MediaBox\s*\[([^\]]+)\]').Groups[1].Value
      $res.pruebas += [ordered]@{ nombre = $c.nombre; paginas = $paginas; paginasPdf = $paginasPdf; mediaBox = $caja; zoomLeido = $zoomLeido; saltosH = $s.h; saltosV = $s.v }
    } catch { $res.pruebas += [ordered]@{ nombre = $c.nombre; error = $_.Exception.Message.Trim() } }
  }
  $libro.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
New-Item -ItemType Directory -Force (Split-Path $Salida) | Out-Null
$res | ConvertTo-Json -Depth 8 | Set-Content -Encoding UTF8 $Salida
"OK: $($res.pruebas.Count) configuraciones → $Salida"

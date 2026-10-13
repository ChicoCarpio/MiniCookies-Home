# Paginas de cada hoja visible tal como Excel las corta (etapa 5): saltos (HPageBreaks/VPageBreaks), Pages.Count,
# y del PDF que exporta Excel: paginas, tamano de hoja y, por pagina, el texto que contiene (referencia del contenido).
# Uso: powershell -ExecutionPolicy Bypass -File volcar-paginas.ps1 -Libro ruta.xlsx -Salida paginas.json
param([Parameter(Mandatory)][string]$Libro, [Parameter(Mandatory)][string]$Salida, [string]$CarpetaPdf = (Join-Path $env:TEMP 'mc-xlsx'))
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$hojas = @()
$tmp = $CarpetaPdf; New-Item -ItemType Directory -Force $tmp | Out-Null
try {
  $l = $xl.Workbooks.Open((Resolve-Path $Libro).Path, 0, $true)
  for ($i = 1; $i -le $l.Worksheets.Count; $i++) {
    $h = $l.Worksheets.Item($i)
    if ($h.Visible -ne -1) { continue }
    $ps = $h.PageSetup
    try {
      $paginas = $ps.Pages.Count
      $hs = @(); $vs = @()
      for ($k = 1; $k -le $h.HPageBreaks.Count; $k++) { try { $b = $h.HPageBreaks.Item($k); $hs += [ordered]@{ fila = $b.Location.Row; col = $b.Location.Column; tipo = $b.Type } } catch { } }
      for ($k = 1; $k -le $h.VPageBreaks.Count; $k++) { try { $b = $h.VPageBreaks.Item($k); $vs += [ordered]@{ fila = $b.Location.Row; col = $b.Location.Column; tipo = $b.Type } } catch { } }
      $pdf = Join-Path $tmp ("p$i.pdf")
      $h.ExportAsFixedFormat(0, $pdf)
      $texto = [Text.Encoding]::GetEncoding(28591).GetString([IO.File]::ReadAllBytes($pdf))
      $paginasPdf = ([regex]::Matches($texto, '/Type\s*/Page[^s]')).Count
      $caja = [regex]::Match($texto, '/MediaBox\s*\[([^\]]+)\]').Groups[1].Value
      $hojas += [ordered]@{ indice = $i - 1; nombre = $h.Name; paginas = $paginas; paginasPdf = $paginasPdf; mediaBox = $caja.Trim(); saltosH = $hs; saltosV = $vs; pdf = ("p$i.pdf") }
    } catch { $hojas += [ordered]@{ indice = $i - 1; nombre = $h.Name; error = $_.Exception.Message.Trim() } }
  }
  $l.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
$json = ConvertTo-Json -InputObject $hojas -Depth 6
[IO.File]::WriteAllText($Salida, $json, (New-Object Text.UTF8Encoding($true)))
"OK -> $Salida ($($hojas.Count) hojas)"

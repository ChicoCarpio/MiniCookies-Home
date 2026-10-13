# Tamaño de página (puntos) que exporta Excel para cada código paperSize (1-118).
# Se lee el MediaBox del PDF que genera ExportAsFixedFormat. Ojo: depende de lo que acepte la impresora
# activa ("Microsoft Print to PDF"); los códigos que no acepta salen como Letter o dan error.
# Uso: powershell -ExecutionPolicy Bypass -File medir-papel.ps1  → medidas\papel.json
param([string]$Salida = (Join-Path $PSScriptRoot 'medidas\papel.json'))
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$res = [ordered]@{ version = $xl.Version; impresora = $xl.ActivePrinter; papeles = @() }
try {
  $libro = $xl.Workbooks.Add()
  $hoja = $libro.Worksheets.Item(1)
  $hoja.Range('A1').Value2 = 'x'
  $tmp = Join-Path $env:TEMP 'mc-xlsx'; New-Item -ItemType Directory -Force $tmp | Out-Null
  foreach ($codigo in 1..118) {
    $fila = [ordered]@{ codigo = $codigo }
    try {
      $hoja.PageSetup.PaperSize = $codigo
      $fila.leido = $hoja.PageSetup.PaperSize
      $pdf = Join-Path $tmp "papel$codigo.pdf"
      $hoja.ExportAsFixedFormat(0, $pdf)
      $t = [Text.Encoding]::GetEncoding(28591).GetString([IO.File]::ReadAllBytes($pdf))
      $fila.mediaBox = [regex]::Match($t, '/MediaBox\s*\[([^\]]+)\]').Groups[1].Value.Trim()
      Remove-Item $pdf -ErrorAction SilentlyContinue
    } catch { $fila.error = $_.Exception.Message.Trim() }
    $res.papeles += $fila
  }
  $libro.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
New-Item -ItemType Directory -Force (Split-Path $Salida) | Out-Null
$res | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 $Salida
"OK → $Salida"

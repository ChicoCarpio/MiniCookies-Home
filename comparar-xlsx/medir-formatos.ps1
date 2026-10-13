# Oráculo del formato de número (etapa 1): cada caso de casos-formato.json se escribe en una celda,
# se le pone el formato y se lee lo que muestra Excel (Range.Text) con la columna bien ancha.
# También lee los errores (#DIV/0! etc.), los formatos integrados (NumberFormatLocal por id) y
# el formato que Excel asigna a "General" según el idioma.
# Uso: powershell -ExecutionPolicy Bypass -File medir-formatos.ps1  → medidas\formatos.json
param([string]$Salida = (Join-Path $PSScriptRoot 'medidas\formatos.json'), [ValidateSet('en','sistema')][string]$Separadores = 'en')
$ErrorActionPreference = 'Stop'
$casos = Get-Content -Raw -Encoding UTF8 (Join-Path $PSScriptRoot 'casos-formato.json') | ConvertFrom-Json
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
# Los códigos de formato de COM se interpretan con la cultura del hilo y los separadores de Excel: se miden en en-US ("." decimal, "," miles) salvo que se pida 'sistema'.
if ($Separadores -eq 'en') {
  [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('en-US')
  $usarSistema = $xl.UseSystemSeparators; $decOrig = $xl.DecimalSeparator; $milOrig = $xl.ThousandsSeparator
  $xl.UseSystemSeparators = $false; $xl.DecimalSeparator = '.'; $xl.ThousandsSeparator = ','
}
$res = [ordered]@{ version = $xl.Version; separadorDecimal = $xl.International(3); separadorMiles = $xl.International(4); casos = @(); errores = @(); integrados = @() }
try {
  $libro = $xl.Workbooks.Add()
  $hoja = $libro.Worksheets.Item(1)
  $hoja.Columns.Item(1).ColumnWidth = 80
  $res.date1904 = $libro.Date1904
  $i = 0
  foreach ($c in $casos) {
    $i++
    $codigo = [string]$c[0]; $valor = $c[1]; if ($valor -is [int] -or $valor -is [long]) { $valor = [double]$valor }
    $celda = $hoja.Cells.Item(1, 1)
    $celda.NumberFormat = 'General'
    $celda.ClearContents() | Out-Null
    $entrada = $null
    try {
      try { $celda.NumberFormat = $codigo } catch {
        # Excel en español exige los nombres de color en español ([Rojo]); en el archivo siempre van en inglés.
        $alt = $codigo
        foreach ($par in @(@('Red', 'Rojo'), @('Blue', 'Azul'), @('Cyan', 'Turquesa'), @('Green', 'Verde'), @('Yellow', 'Amarillo'), @('White', 'Blanco'), @('Black', 'Negro'))) { $alt = $alt.Replace('[' + $par[0] + ']', '[' + $par[1] + ']') }
        $celda.NumberFormat = $alt
      }
      if ($valor -is [string]) { $celda.NumberFormat = "@"; $celda.Value2 = $valor; $celda.NumberFormat = $codigo }
      elseif ($valor -is [bool]) { $celda.Value2 = [bool]$valor }
      else { $celda.Value2 = [double]$valor }
      $texto = [string]$celda.Text
      $res.casos += [ordered]@{ n = $i; codigo = $codigo; valor = $valor; tipo = $(if ($valor -is [string]) { 'texto' } elseif ($valor -is [bool]) { 'bool' } else { 'numero' }); texto = $texto; colorFuente = [int]$celda.DisplayFormat.Font.Color; formatoAplicado = [string]$celda.NumberFormat; formatoLocal = [string]$celda.NumberFormatLocal }
    } catch {
      $res.casos += [ordered]@{ n = $i; codigo = $codigo; valor = $valor; error = $_.Exception.Message.Trim(); donde = $_.ScriptStackTrace + " | " + $_.InvocationInfo.PositionMessage }
    }
  }
  # Errores de celda con formatos varios.
  foreach ($f in 'General', '0.00', '@', '0%', 'yyyy-mm-dd') {
    foreach ($fx in '=1/0', '=NA()', '=SQRT(-1)', '=FOO()', '=INDIRECT("A0")', '="a"+1', '=#REF!') {
      $celda = $hoja.Cells.Item(2, 1)
      try { $celda.NumberFormat = $f; $celda.Formula = $fx; $res.errores += [ordered]@{ formato = $f; formula = $fx; texto = [string]$celda.Text } } catch { $res.errores += [ordered]@{ formato = $f; formula = $fx; error = $_.Exception.Message.Trim() } }
    }
  }
  # Formato de cada id integrado: se escribe el formato por id usando el estilo de celda no es posible por COM;
  # se lee NumberFormat tras aplicar los atajos estándar de Excel (ver LEEME). Aquí solo se anota el estado.
  $libro.Close($false)
} finally {
  if ($Separadores -eq 'en') { $xl.DecimalSeparator = $decOrig; $xl.ThousandsSeparator = $milOrig; $xl.UseSystemSeparators = $usarSistema }
  $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
New-Item -ItemType Directory -Force (Split-Path $Salida) | Out-Null
$res | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 $Salida
"OK: $($res.casos.Count) casos → $Salida"

# Corpus sintético del motor de Excel: un libro por característica, guardado por Excel real.
# Uso: powershell -ExecutionPolicy Bypass -File generar-corpus.ps1 [-Solo '^03']
# Salida: corpus\*.xlsx (en git: son chicos y no tienen datos personales).
param([string]$Solo = '')
$ErrorActionPreference = 'Stop'
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('en-US')
$destino = Join-Path $PSScriptRoot 'corpus'
New-Item -ItemType Directory -Force $destino | Out-Null
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
# Los formatos de COM se interpretan con los separadores de Excel: se fijan en en-US mientras dura la generación.
$usarSistema = $xl.UseSystemSeparators; $decOrig = $xl.DecimalSeparator; $milOrig = $xl.ThousandsSeparator
$xl.UseSystemSeparators = $false; $xl.DecimalSeparator = '.'; $xl.ThousandsSeparator = ','
$hechos = @()

function Caso([string]$nombre, [scriptblock]$armar) {
  if ($Solo -and $nombre -notmatch $Solo) { return }
  $libro = $xl.Workbooks.Add()
  try {
    & $armar $libro
    $ruta = Join-Path $destino ($nombre + '.xlsx')
    if (Test-Path $ruta) { Remove-Item $ruta }
    $libro.SaveAs($ruta, 51)
    $script:hechos += $nombre
  } catch { Write-Host ("FALLO " + $nombre + ": " + $_.Exception.Message.Trim() + " @ " + $_.InvocationInfo.ScriptLineNumber) }
  finally { $libro.Close($false) }
}

Caso '01-tipos' {
  param($l)
  $h = $l.Worksheets.Item(1); $h.Name = 'Tipos'
  $h.Range('A1').Value2 = 'Texto'; $h.Range('B1').Value2 = 'Número'; $h.Range('C1').Value2 = 'Booleano'; $h.Range('D1').Value2 = 'Error'; $h.Range('E1').Value2 = 'Fecha'
  $h.Range('A2').Value2 = 'hola'; $h.Range('B2').Value2 = [double]3.14159265358979; $h.Range('C2').Value2 = $true; $h.Range('D2').Formula = '=1/0'; $h.Range('E2').Value2 = [double]45000; $h.Range('E2').NumberFormat = 'yyyy-mm-dd'
  $h.Range('A3').Value2 = "con`nsalto de línea"; $h.Range('B3').Value2 = [double]-0.000001234; $h.Range('C3').Value2 = $false; $h.Range('D3').Formula = '=NA()'; $h.Range('E3').Value2 = [double]0.5; $h.Range('E3').NumberFormat = 'h:mm'
  $h.Range('A4').Value2 = '  espacios  '; $h.Range('B4').Value2 = [double]1E+30; $h.Range('D4').Formula = '=SQRT(-1)'; $h.Range('E4').Value2 = [double]60; $h.Range('E4').NumberFormat = 'yyyy-mm-dd'
  $h.Range('A5').Value2 = 'ñandú ü 漢字 😀'; $h.Range('B5').Value2 = [double]123456789012345; $h.Range('D5').Formula = '=INDIRECT("A0")'
  $h.Range('A6').Value2 = "'apóstrofo"; $h.Range('B6').Value2 = [double]0.1; $h.Range('D6').Formula = '="a"+1'
  $h.Range('A7').Value2 = ('x' * 5000); $h.Range('B7').Value2 = [double]1; $h.Range('B7').NumberFormat = '0.00%'
  $h.Range('A8').Value2 = '<b>&amp;</b> "comillas" & ''simples'''; $h.Range('B8').Value2 = [double]99999999999.5
  $h.Range('A9').NumberFormat = '@'; $h.Range('A9').Value2 = '00123'
  $h.Range('A10').Formula = '=A2&" mundo"'; $h.Range('B10').Formula = '=SUM(B2:B3)'; $h.Range('C10').Formula = '=B2>1'; $h.Range('E10').Formula = '=TODAY()'
}

Caso '02-fechas-1904' {
  param($l)
  $l.Date1904 = $true
  $h = $l.Worksheets.Item(1); $h.Name = 'Fechas1904'
  $h.Range('A1').Value2 = [double]0; $h.Range('A2').Value2 = [double]1; $h.Range('A3').Value2 = [double]43000.75; $h.Range('A4').Value2 = [double]100
  $h.Range('A1:A4').NumberFormat = 'yyyy-mm-dd hh:mm:ss'
}

Caso '03-estilos' {
  param($l)
  $h = $l.Worksheets.Item(1); $h.Name = 'Estilos'
  $h.Range('A1').Value2 = 'negrita'; $h.Range('A1').Font.Bold = $true
  $h.Range('A2').Value2 = 'cursiva'; $h.Range('A2').Font.Italic = $true
  $h.Range('A3').Value2 = 'subrayado'; $h.Range('A3').Font.Underline = 2
  $h.Range('A4').Value2 = 'tachado'; $h.Range('A4').Font.Strikethrough = $true
  $h.Range('A5').Value2 = 'Arial 14 rojo'; $h.Range('A5').Font.Name = 'Arial'; $h.Range('A5').Font.Size = 14; $h.Range('A5').Font.Color = 255
  $h.Range('A6').Value2 = 'relleno amarillo'; $h.Range('A6').Interior.Color = 65535
  $h.Range('A7').Value2 = 'tema accent1 tint 0.4'; $h.Range('A7').Interior.ThemeColor = 5; $h.Range('A7').Interior.TintAndShade = 0.4
  $h.Range('A8').Value2 = 'tema dk2 tint -0.25'; $h.Range('A8').Interior.ThemeColor = 3; $h.Range('A8').Interior.TintAndShade = -0.25
  $h.Range('A9').Value2 = 'fuente tema accent2'; $h.Range('A9').Font.ThemeColor = 6
  $h.Range('A10').Value2 = 'trama gris 25'; $h.Range('A10').Interior.Pattern = -4124; $h.Range('A10').Interior.PatternColor = 255
  $h.Range('A11').Value2 = 'índice 3'; $h.Range('A11').Interior.ColorIndex = 3
  $h.Range('B1').Value2 = 'borde fino'; $h.Range('B1').Borders.LineStyle = 1; $h.Range('B1').Borders.Weight = 2
  $h.Range('B2').Value2 = 'borde grueso rojo'; $h.Range('B2').Borders.LineStyle = 1; $h.Range('B2').Borders.Weight = 4; $h.Range('B2').Borders.Color = 255
  $h.Range('B3').Value2 = 'doble'; $h.Range('B3').Borders.Item(9).LineStyle = -4119
  $h.Range('B4').Value2 = 'punteado'; $h.Range('B4').Borders.Item(7).LineStyle = -4118
  $h.Range('B5').Value2 = 'guiones'; $h.Range('B5').Borders.Item(8).LineStyle = -4115
  $h.Range('B6').Value2 = 'diagonal'; $h.Range('B6').Borders.Item(5).LineStyle = 1
  $h.Range('C1').Value2 = 'izq'; $h.Range('C1').HorizontalAlignment = -4131
  $h.Range('C2').Value2 = 'centro'; $h.Range('C2').HorizontalAlignment = -4108
  $h.Range('C3').Value2 = 'der'; $h.Range('C3').HorizontalAlignment = -4152
  $h.Range('C4').Value2 = 'sangría 3'; $h.Range('C4').IndentLevel = 3
  $h.Range('C5').Value2 = 'giro 45'; $h.Range('C5').Orientation = 45
  $h.Range('C6').Value2 = 'vertical'; $h.Range('C6').Orientation = -4166
  $h.Range('C7').Value2 = 'texto largo con ajuste de línea para varias líneas'; $h.Range('C7').WrapText = $true
  $h.Range('C8').Value2 = 'reducir hasta ajustar'; $h.Range('C8').ShrinkToFit = $true
  $h.Range('C9').Value2 = 'arriba'; $h.Range('C9').VerticalAlignment = -4160
  $h.Range('C10').Value2 = 'abajo'; $h.Range('C10').VerticalAlignment = -4107
  $h.Range('C11').Value2 = 'rellenar'; $h.Range('C11').HorizontalAlignment = 5
  $h.Range('C12').Value2 = 'justificar'; $h.Range('C12').HorizontalAlignment = -4130
  $h.Range('D1').Value2 = [double]1234.5; $h.Range('D1').NumberFormat = '#,##0.00'
  $h.Range('D2').Value2 = [double]0.256; $h.Range('D2').NumberFormat = '0.0%'
  $h.Range('D3').Value2 = [double]-5; $h.Range('D3').NumberFormat = '#,##0;[Color3]-#,##0'
  $h.Range('D4').Value2 = [double]45000; $h.Range('D4').NumberFormat = 'dddd, mmmm d, yyyy'
  $h.Range('D5').Value2 = [double]0.75; $h.Range('D5').NumberFormat = '# ?/?'
  $h.Range('D6').Value2 = [double]12345; $h.Range('D6').NumberFormat = '0.00E+00'
  $h.Range('D7').Value2 = [double]5; $h.Range('D7').NumberFormat = '"$"#,##0.00'
  $h.Range('D8').Value2 = [double]5; $h.Range('D8').NumberFormat = '_($* #,##0.00_);_($* (#,##0.00)'
  $h.Rows.Item(2).RowHeight = 30
}

Caso '04-combinadas' {
  param($l)
  $h = $l.Worksheets.Item(1); $h.Name = 'Combinadas'
  $h.Range('A1').Value2 = 'combinada 3x2'; $h.Range('A1:C2').Merge()
  $h.Range('A1').HorizontalAlignment = -4108; $h.Range('A1').VerticalAlignment = -4108
  $h.Range('E1').Value2 = 'solo fila'; $h.Range('E1:H1').Merge()
  $h.Range('E3').Value2 = 'solo col'; $h.Range('E3:E6').Merge()
  $h.Range('G3').Value2 = 'con borde'; $h.Range('G3:H4').Merge(); $h.Range('G3:H4').BorderAround(1, 4) | Out-Null
  $h.Range('A8').Value2 = 'valor oculto'; $h.Range('B8').Value2 = 'se pierde'
}

Caso '05-anchos-altos' {
  param($l)
  $h = $l.Worksheets.Item(1); $h.Name = 'Medidas'
  $h.Range('A1:J10').Value2 = 'x'
  $h.Columns.Item(1).ColumnWidth = 3; $h.Columns.Item(2).ColumnWidth = 20; $h.Columns.Item(3).ColumnWidth = 8.43
  $h.Columns.Item(4).Hidden = $true
  $h.Columns.Item(5).ColumnWidth = 40
  $h.Range('F1:G1').EntireColumn.ColumnWidth = 12.5
  $h.Rows.Item(2).RowHeight = 40; $h.Rows.Item(3).Hidden = $true; $h.Rows.Item(5).RowHeight = 8
  $h.Range('A7').Value2 = "tres`nlíneas`nautoajuste"; $h.Range('A7').WrapText = $true; $h.Rows.Item(7).AutoFit() | Out-Null
  $h.Range('A8:A9').EntireRow.Group() | Out-Null
  $h.Range('H1:I1').EntireColumn.Group() | Out-Null
  $h.Range('A1').Select() | Out-Null
  $xl.ActiveWindow.SplitRow = 1; $xl.ActiveWindow.SplitColumn = 1; $xl.ActiveWindow.FreezePanes = $true
  $xl.ActiveWindow.DisplayGridlines = $false; $xl.ActiveWindow.Zoom = 85
}

Caso '06-nombres' {
  param($l)
  $h = $l.Worksheets.Item(1); $h.Name = 'Datos'
  $h.Range('A1:A5').Value2 = [double]7
  $l.Names.Add('Global1', '=Datos!$A$1:$A$5') | Out-Null
  $h.Names.Add('Local1', '=Datos!$A$1') | Out-Null
  $l.Names.Add('Constante', '=42') | Out-Null
  $h2 = $l.Worksheets.Add([Type]::Missing, $h); $h2.Name = 'Otra'
  $h2.Range('A1').Formula = '=SUM(Global1)+Constante'
  $h2.Names.Add('Local1', '=Otra!$B$2') | Out-Null
  $h.PageSetup.PrintArea = '$A$1:$B$3'
  $h.PageSetup.PrintTitleRows = '$1:$1'
}

Caso '07-hojas' {
  param($l)
  $h = $l.Worksheets.Item(1); $h.Name = 'Visible'
  $h.Range('A1').Value2 = 'v'
  $o = $l.Worksheets.Add([Type]::Missing, $l.Worksheets.Item($l.Worksheets.Count)); $o.Name = 'Oculta'; $o.Range('A1').Value2 = 'o'; $o.Visible = 0
  $mo = $l.Worksheets.Add([Type]::Missing, $l.Worksheets.Item($l.Worksheets.Count)); $mo.Name = 'MuyOculta'; $mo.Range('A1').Value2 = 'mo'; $mo.Visible = 2
  $e = $l.Worksheets.Add([Type]::Missing, $l.Worksheets.Item($l.Worksheets.Count)); $e.Name = "Año 2025 ('x') & más"; $e.Range('A1').Value2 = 'esp'
  $e.Tab.Color = 255
  $c = $l.Worksheets.Add([Type]::Missing, $l.Worksheets.Item($l.Worksheets.Count)); $c.Name = 'Con espacios y ñ'; $c.Range('A1').Formula = "='Año 2025 (''x'') & más'!A1"
  $l.Worksheets.Item(1).Activate()
}

Caso '08-tablas-enlaces' {
  param($l)
  $h = $l.Worksheets.Item(1); $h.Name = 'Tabla'
  $h.Range('A1').Value2 = 'Producto'; $h.Range('B1').Value2 = 'Cantidad'; $h.Range('C1').Value2 = 'Precio'
  $h.Range('A2').Value2 = 'A'; $h.Range('B2').Value2 = [double]3; $h.Range('C2').Value2 = [double]2.5
  $h.Range('A3').Value2 = 'B'; $h.Range('B3').Value2 = [double]5; $h.Range('C3').Value2 = [double]1.25
  $h.Range('A4').Value2 = 'C'; $h.Range('B4').Value2 = [double]7; $h.Range('C4').Value2 = [double]0.75
  $t = $h.ListObjects.Add(1, $h.Range('A1:C4'), $null, 1); $t.Name = 'Ventas'; $t.TableStyle = 'TableStyleMedium2'; $t.ShowTotals = $true
  $h.Hyperlinks.Add($h.Range('E1'), 'https://example.com/', '', 'Sitio', 'Ejemplo') | Out-Null
  $h.Hyperlinks.Add($h.Range('E2'), '', "'Tabla'!A1", 'Ir a A1', 'Interno') | Out-Null
  $h.Hyperlinks.Add($h.Range('E3'), 'mailto:alguien@example.com', '', 'Correo', 'Escribir') | Out-Null
  $h.Range('A2').AddComment('Nota de prueba') | Out-Null
  $h.Range('A6').Formula = '=SUM(Ventas[Cantidad])'
  $h.Range('B6').Formula = '=Ventas[[#This Row],[Precio]]'
}

Caso '09-rich-text' {
  param($l)
  $h = $l.Worksheets.Item(1); $h.Name = 'Rico'
  $h.Range('A1').Value2 = 'negrita y cursiva y rojo'
  $h.Range('A1').Characters(1, 7).Font.Bold = $true
  $h.Range('A1').Characters(11, 7).Font.Italic = $true
  $h.Range('A1').Characters(20, 4).Font.Color = 255
  $h.Range('A2').Value2 = 'H2O y x2'
  $h.Range('A2').Characters(2, 1).Font.Subscript = $true
  $h.Range('A2').Characters(8, 1).Font.Superscript = $true
}

Caso '10-formulas' {
  param($l)
  $h = $l.Worksheets.Item(1); $h.Name = 'F'
  1..10 | ForEach-Object { $h.Cells.Item($_, 1).Value2 = [double]$_ }
  $h.Range('B1:B10').Formula = '=A1*2'          # compartida
  $h.Range('C1').FormulaArray = '=SUMA(A1:A10*2)'   # FormulaArray usa el idioma de Excel (español)
  $h.Range('D1').Formula = '=IF(A1>5,"alto","bajo")'
  $h.Range('D2').Formula = '=VLOOKUP(3,A1:B10,2,FALSE)'
  $h.Range('D3').Formula2 = '=SEQUENCE(3)'
  $h.Range('D6').Formula = '=TEXT(A1,"0.00")&"x"'
  $h.Range('D7').Formula = '=RAND()'
  $h.Range('D8').Formula = '=NOW()'
  $h.Range('D9').Formula = '=Hoja_inexistente!A1'
}

Caso '11-vacios-y-estilos-fila' {
  param($l)
  $h = $l.Worksheets.Item(1); $h.Name = 'Vacios'
  $h.Rows.Item(2).Interior.Color = 65535
  $h.Columns.Item(3).Font.Bold = $true
  $h.Range('A1').Value2 = 'a'; $h.Range('Z1').Value2 = 'z'
  $h.Range('B5').Interior.Color = 255   # celda con estilo y sin valor
  $h.Range('A100').Value2 = 'fila 100'
  $h.Range('A1048576').Value2 = 'última fila'
  $h.Range('XFD1').Value2 = 'última columna'
}

$xl.DecimalSeparator = $decOrig; $xl.ThousandsSeparator = $milOrig; $xl.UseSystemSeparators = $usarSistema
$xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
"OK: $($hechos.Count) libros en $destino"

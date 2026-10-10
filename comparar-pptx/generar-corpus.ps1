# Genera con PowerPoint real (COM) el corpus de presentaciones de prueba del arnés de fidelidad del motor de PowerPoint de Mini Cookies.
<#
.SYNOPSIS
  Crea en comparar-pptx\corpus\ un .pptx por tema (formas, texto, autoajuste,
  viñetas, rellenos, líneas, efectos, tablas, grupos, imágenes, SmartArt,
  gráficos, patrón y tamaños) usando Microsoft PowerPoint instalado.

.DESCRIPTION
  - Cada archivo se crea con Presentations.Add (sin ventana, salvo el de SmartArt,
    porque PowerPoint solo carga Application.SmartArtLayouts con una ventana).
  - Cada caso va en su propio try/catch: si una API falla se registra y se sigue.
  - Es determinista (sin fechas ni azar) y re-ejecutable: sobrescribe los .pptx.
  - Deja además corpus\estilos-tabla-resultados.csv (los 74 GUID de estilos de
    tabla: si PowerPoint los aceptó) y corpus\errores.txt.
  - No deja procesos POWERPNT/EXCEL propios colgados: al final cierra los que
    haya abierto esta ejecución (y solo esos).

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File comparar-pptx\generar-corpus.ps1
#>
[CmdletBinding()]
param([string]$Salida = '', [string]$Solo = '')   # -Solo: expresión regular para regenerar solo algunos archivos (p. ej. -Solo '^03')

$ErrorActionPreference = 'Stop'
$AquiDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Salida) { $Salida = Join-Path $AquiDir 'corpus' }
$null = New-Item -ItemType Directory -Force -Path $Salida
$Salida = (Resolve-Path $Salida).Path
$Recursos = Join-Path $Salida 'recursos'
$null = New-Item -ItemType Directory -Force -Path $Recursos

# ---------------------------------------------------------------- constantes
$msoTrue = -1   # msoTrue
$msoFalse = 0    # msoFalse
$script:Errores = New-Object System.Collections.ArrayList
$script:Resumen = New-Object System.Collections.ArrayList
$script:Notas   = New-Object System.Collections.ArrayList
$script:ArchivoActual = ''
$script:ResultadosEstilos = New-Object System.Collections.ArrayList

# Los 74 estilos de tabla (04-tablas-graficos-smartart.md, 4.1.3): grupo|acento|GUID
$EstilosTabla = @(
    'Themed Style 1|sin acento|{2D5ABB26-0587-4C30-8999-92F81FD0307C}',
    'Themed Style 1|Accent 1|{3C2FFA5D-87B4-456A-9821-1D502468CF0F}',
    'Themed Style 1|Accent 2|{284E427A-3D55-4303-BF80-6455036E1DE7}',
    'Themed Style 1|Accent 3|{69C7853C-536D-4A76-A0AE-DD22124D55A5}',
    'Themed Style 1|Accent 4|{775DCB02-9BB8-47FD-8907-85C794F793BA}',
    'Themed Style 1|Accent 5|{35758FB7-9AC5-4552-8A53-C91805E547FA}',
    'Themed Style 1|Accent 6|{08FB837D-C827-4EFA-A057-4D05807E0F7C}',
    'Themed Style 2|sin acento|{5940675A-B579-460E-94D1-54222C63F5DA}',
    'Themed Style 2|Accent 1|{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}',
    'Themed Style 2|Accent 2|{18603FDC-E32A-4AB5-989C-0864C3EAD2B8}',
    'Themed Style 2|Accent 3|{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}',
    'Themed Style 2|Accent 4|{E269D01E-BC32-4049-B463-5C60D7B0CCD2}',
    'Themed Style 2|Accent 5|{327F97BB-C833-4FB7-BDE5-3F7075034690}',
    'Themed Style 2|Accent 6|{638B1855-1B75-4FBE-930C-398BA8C253C6}',
    'Light Style 1|sin acento|{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}',
    'Light Style 1|Accent 1|{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}',
    'Light Style 1|Accent 2|{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}',
    'Light Style 1|Accent 3|{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}',
    'Light Style 1|Accent 4|{D27102A9-8310-4765-A935-A1911B00CA55}',
    'Light Style 1|Accent 5|{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}',
    'Light Style 1|Accent 6|{68D230F3-CF80-4859-8CE7-A43EE81993B5}',
    'Light Style 2|sin acento|{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}',
    'Light Style 2|Accent 1|{69012ECD-51FC-41F1-AA8D-1B2483CD663E}',
    'Light Style 2|Accent 2|{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}',
    'Light Style 2|Accent 3|{F2DE63D5-997A-4646-A377-4702673A728D}',
    'Light Style 2|Accent 4|{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}',
    'Light Style 2|Accent 5|{5A111915-BE36-4E01-A7E5-04B1672EAD32}',
    'Light Style 2|Accent 6|{912C8C85-51F0-491E-9774-3900AFEF0FD7}',
    'Light Style 3|sin acento|{616DA210-FB5B-4158-B5E0-FEB733F419BA}',
    'Light Style 3|Accent 1|{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}',
    'Light Style 3|Accent 2|{5DA37D80-6434-44D0-A028-1B22A696006F}',
    'Light Style 3|Accent 3|{8799B23B-EC83-4686-B30A-512413B5E67A}',
    'Light Style 3|Accent 4|{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}',
    'Light Style 3|Accent 5|{BDBED569-4797-4DF1-A0F4-6AAB3CD982D8}',
    'Light Style 3|Accent 6|{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}',
    'Medium Style 1|sin acento|{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}',
    'Medium Style 1|Accent 1|{B301B821-A1FF-4177-AEE7-76D212191A09}',
    'Medium Style 1|Accent 2|{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}',
    'Medium Style 1|Accent 3|{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}',
    'Medium Style 1|Accent 4|{1E171933-4619-4E11-9A3F-F7608DF75F80}',
    'Medium Style 1|Accent 5|{FABFCF23-3B69-468F-B69F-88F6DE6A72F2}',
    'Medium Style 1|Accent 6|{10A1B5D5-9B99-4C35-A422-299274C87663}',
    'Medium Style 2|sin acento|{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}',
    'Medium Style 2|Accent 1|{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}',
    'Medium Style 2|Accent 2|{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}',
    'Medium Style 2|Accent 3|{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}',
    'Medium Style 2|Accent 4|{00A15C55-8517-42AA-B614-E9B94910E393}',
    'Medium Style 2|Accent 5|{7DF18680-E054-41AD-8BC1-D1AEF772440D}',
    'Medium Style 2|Accent 6|{93296810-A885-4BE3-A3E7-6D5BEEA58F35}',
    'Medium Style 3|sin acento|{8EC20E35-A176-4012-BC5E-935CFFF8708E}',
    'Medium Style 3|Accent 1|{6E25E649-3F16-4E02-A733-19D2CDBF48F0}',
    'Medium Style 3|Accent 2|{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}',
    'Medium Style 3|Accent 3|{EB344D84-9AFB-497E-A393-DC336BA19D2E}',
    'Medium Style 3|Accent 4|{EB9631B5-78F2-41C9-869B-9F39066F8104}',
    'Medium Style 3|Accent 5|{74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6}',
    'Medium Style 3|Accent 6|{2A488322-F2BA-4B5B-9748-0D474271808F}',
    'Medium Style 4|sin acento|{D7AC3CCA-C797-4891-BE02-D94E43425B78}',
    'Medium Style 4|Accent 1|{69CF1AB2-1976-4502-BF36-3FF5EA218861}',
    'Medium Style 4|Accent 2|{8A107856-5554-42FB-B03E-39F5DBC370BA}',
    'Medium Style 4|Accent 3|{0505E3EF-67EA-436B-97B2-0124C06EBD24}',
    'Medium Style 4|Accent 4|{C4B1156A-380E-4F78-BDF5-A606A8083BF9}',
    'Medium Style 4|Accent 5|{22838BEF-8BB2-4498-84A7-C5851F593DF1}',
    'Medium Style 4|Accent 6|{16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4}',
    'Dark Style 1|sin acento|{E8034E78-7F5D-4C2E-B375-FC64B27BC917}',
    'Dark Style 1|Accent 1|{125E5076-3810-47DD-B79F-674D7AD40C01}',
    'Dark Style 1|Accent 2|{37CE84F3-28C3-443E-9E96-99CF82512B78}',
    'Dark Style 1|Accent 3|{D03447BB-5D67-496B-8E87-E561075AD55C}',
    'Dark Style 1|Accent 4|{E929F9F4-4A8F-4326-A1B4-22849713DDAB}',
    'Dark Style 1|Accent 5|{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}',
    'Dark Style 1|Accent 6|{AF606853-7671-496A-8E4F-DF71F8EC918B}',
    'Dark Style 2|sin acento|{5202B0CA-FC54-4496-8BCA-5EF66A818D29}',
    'Dark Style 2|Accent 1|{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}',
    'Dark Style 2|Accent 3|{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}',
    'Dark Style 2|Accent 5|{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}'
)

$ParrafoLargo = 'El veloz murciélago hindú comía feliz cardillo y kiwi mientras la cigüeña tocaba el saxofón detrás del palenque de paja. ' +
                'Quien madruga encuentra el café caliente, pero el que duerme sueña con montañas, ríos y pájaros de colores.'

# ---------------------------------------------------------------- ayudantes
function RGB([int]$r, [int]$g, [int]$b) { return ($r + 256 * $g + 65536 * $b) }

function Registrar([string]$caso, $err) {
    $msg = ($err.Exception.Message -replace '\s+', ' ').Trim()
    if ($err.InvocationInfo -and $err.InvocationInfo.ScriptLineNumber) { $msg += ' (línea ' + $err.InvocationInfo.ScriptLineNumber + ')' }
    [void]$script:Errores.Add([pscustomobject]@{ Archivo = $script:ArchivoActual; Caso = $caso; Error = $msg })
    Write-Host ("  [ERROR] {0}: {1}" -f $caso, $msg) -ForegroundColor Red
}

# Ejecuta un caso; si falla, lo registra y sigue.
function Caso([string]$nombre, [scriptblock]$bloque) {
    try { & $bloque } catch { Registrar $nombre $_ }
}

function Nota([string]$texto) {
    [void]$script:Notas.Add(("{0}: {1}" -f $script:ArchivoActual, $texto))
    Write-Host ("  [nota] " + $texto) -ForegroundColor Yellow
}

function Poner-Rotulo($d, [string]$texto) {
    $r = $d.Shapes.AddTextbox(1, 6, 2, 420, 20)
    $r.Name = 'Rotulo'
    $r.TextFrame.WordWrap = $msoFalse
    $r.TextFrame.TextRange.Text = $texto
    $r.TextFrame.TextRange.Font.Name = 'Arial'
    $r.TextFrame.TextRange.Font.Size = 12
    $r.TextFrame.TextRange.ParagraphFormat.Alignment = 1
}

# Diapositiva en blanco con su rótulo.
function Nueva-Diap($pres, [string]$rotulo, [int]$layout = 12) {
    $d = $pres.Slides.Add($pres.Slides.Count + 1, $layout)
    Poner-Rotulo $d $rotulo
    return $d
}

# Diapositiva con un diseño del patrón (índice de CustomLayouts) y su rótulo.
function Nueva-DiapDiseno($pres, [string]$rotulo, [int]$indiceDiseno) {
    $lay = $pres.SlideMaster.CustomLayouts.Item($indiceDiseno)
    $d = $pres.Slides.AddSlide($pres.Slides.Count + 1, $lay)
    Poner-Rotulo $d $rotulo
    return $d
}

function Caja($d, [int]$tipo, [double]$x, [double]$y, [double]$w, [double]$h, [string]$texto = '', [double]$tam = 10) {
    $s = $d.Shapes.AddShape($tipo, $x, $y, $w, $h)
    if ($texto -ne '') {
        $s.TextFrame.TextRange.Text = $texto
        $s.TextFrame.TextRange.Font.Size = $tam
    }
    return $s
}

function Texto($d, [double]$x, [double]$y, [double]$w, [double]$h, [string]$texto, [double]$tam = 12) {
    $s = $d.Shapes.AddTextbox(1, $x, $y, $w, $h)
    $s.TextFrame.TextRange.Text = $texto
    $s.TextFrame.TextRange.Font.Size = $tam
    return $s
}

function Etiqueta($d, [double]$x, [double]$y, [double]$w, [string]$texto) {
    $s = $d.Shapes.AddTextbox(1, $x, $y, $w, 14)
    $s.TextFrame.TextRange.Text = $texto
    $s.TextFrame.TextRange.Font.Name = 'Arial'
    $s.TextFrame.TextRange.Font.Size = 8
    $s.TextFrame.MarginTop = 0; $s.TextFrame.MarginBottom = 0
    return $s
}

function Palabras([int]$n) {
    $b = 'El veloz murciélago hindú comía feliz cardillo y kiwi mientras la cigüeña tocaba el saxofón detrás del palenque de paja'.Split(' ')
    $o = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $n; $i++) { [void]$o.Add($b[$i % $b.Count]) }
    return (($o -join ' ') + '.')
}

# Un gráfico con sus datos: categorías (texto o números) y series (nombre →
# valores). La tabla de datos se ajusta al rango exacto: si no, quedan filas
# de los datos de ejemplo de Excel (los de áreas traen cinco filas).
function Grafico($d, [int]$tipo, [double]$x, [double]$y, [double]$w, [double]$h, [string]$nombre, $categorias, $series, [string]$titulo = '') {
    $sh = $d.Shapes.AddChart2(-1, $tipo, $x, $y, $w, $h, $msoTrue)
    $sh.Name = $nombre
    $ch = $sh.Chart
    $wb = $null
    try {
        $ch.ChartData.ActivateChartDataWindow()
        $wb = $ch.ChartData.Workbook
        $ws = $wb.Worksheets.Item(1)
        $ws.Range('A1:Z60').ClearContents()
        $nombres = @($series.Keys)
        for ($c = 0; $c -lt $nombres.Count; $c++) { $ws.Cells.Item(1, $c + 2).Value2 = [string]$nombres[$c] }
        for ($r = 0; $r -lt $categorias.Count; $r++) {
            # Un número por Value2 falla en esta celda (el adaptador de COM la
            # toma como texto): va como fórmula, que Excel lee en inglés.
            $v = $categorias[$r]
            if ($v -is [string]) { $ws.Cells.Item($r + 2, 1).Value2 = $v }
            else { $ws.Cells.Item($r + 2, 1).Formula = [string]::Format([cultureinfo]::InvariantCulture, '{0}', $v) }
            for ($c = 0; $c -lt $nombres.Count; $c++) { $ws.Cells.Item($r + 2, $c + 2).Value2 = [double]$series[$nombres[$c]][$r] }
        }
        $rango = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($categorias.Count + 1, $nombres.Count + 1))
        try { $ch.SetSourceData("='" + $ws.Name + "'!" + $rango.Address()) } catch { Registrar ('origen de datos de ' + $nombre) $_ }
        try { $ch.Refresh() } catch { }
        Start-Sleep -Milliseconds 700   # que PowerPoint vuelque los datos al gráfico antes de cerrar la hoja
    } finally { if ($wb -ne $null) { try { $wb.Close() } catch { } } }
    if ($titulo) { $ch.HasTitle = $true; $ch.ChartTitle.Text = $titulo } else { $ch.HasTitle = $false }
    return $sh
}

# Garantiza que haya una instancia viva de PowerPoint (si se cayó, abre otra).
function Asegurar-PowerPoint {
    $vivo = $false
    if ($script:ppt -ne $null) { try { $null = $script:ppt.Presentations.Count; $vivo = $true } catch { $vivo = $false } }
    if (-not $vivo) {
        if ($script:ppt -ne $null) { Write-Host '  (PowerPoint dejó de responder: se abre otra instancia)' -ForegroundColor Yellow }
        $script:ppt = New-Object -ComObject PowerPoint.Application
        try { $script:ppt.DisplayAlerts = 1 } catch { }   # ppAlertsNone: sin diálogos
    }
}

# Crea (o recrea) un archivo: ejecuta $cuerpo con la presentación, guarda, cuenta y cierra.
# Si el archivo entero falla (p. ej. PowerPoint se cae), reintenta una vez con una instancia nueva.
function Nuevo-Archivo([string]$nombre, [scriptblock]$cuerpo, [int]$ancho = 960, [int]$alto = 540, [bool]$conVentana = $false) {
    if ($Solo -and $nombre -notmatch $Solo) { return }
    $script:ArchivoActual = $nombre
    Write-Host ("== " + $nombre) -ForegroundColor Cyan
    $ruta = Join-Path $Salida $nombre
    for ($intento = 1; $intento -le 2; $intento++) {
        if (Test-Path $ruta) { Remove-Item $ruta -Force }
        $nErr = $script:Errores.Count; $nNotas = $script:Notas.Count
        $pres = $null; $fallo = $false
        try {
            Asegurar-PowerPoint
            if ($conVentana) {
                $pres = $ppt.Presentations.Add($msoTrue)
                try { $ppt.WindowState = 2 } catch { }   # minimizada
            } else {
                $pres = $ppt.Presentations.Add($msoFalse)
            }
            if ($ancho -gt 0) {
                $pres.PageSetup.SlideWidth = $ancho
                $pres.PageSetup.SlideHeight = $alto
            }
            & $cuerpo $pres
            $pres.SaveAs($ruta, 24)
            $formas = 0
            foreach ($sl in $pres.Slides) { $formas += $sl.Shapes.Count }
            [void]$script:Resumen.Add([pscustomobject]@{ Archivo = $nombre; Diapositivas = $pres.Slides.Count; Formas = $formas; Tamano = "$($pres.PageSetup.SlideWidth)x$($pres.PageSetup.SlideHeight)" })
        } catch {
            $fallo = $true
            Registrar ('archivo completo (intento ' + $intento + ')') $_
        } finally {
            if ($pres -ne $null) { try { $pres.Close() } catch { } }
        }
        if (-not $fallo) { break }
        if ($intento -eq 1) {
            # descartar lo registrado en el intento fallido y reiniciar PowerPoint
            while ($script:Errores.Count -gt $nErr) { $script:Errores.RemoveAt($script:Errores.Count - 1) }
            while ($script:Notas.Count -gt $nNotas) { $script:Notas.RemoveAt($script:Notas.Count - 1) }
            [void]$script:Errores.Add([pscustomobject]@{ Archivo = $nombre; Caso = 'archivo completo (intento 1)'; Error = 'falló el primer intento; se reintentó con una instancia nueva de PowerPoint' })
            try { $script:ppt.Quit() } catch { }
            $script:ppt = $null
            if ($nombre -eq '08-tablas.pptx') { $script:ResultadosEstilos.Clear() }
        }
    }
}

# ---------------------------------------------------------------- imágenes de prueba
function Crear-Imagenes {
    Add-Type -AssemblyName System.Drawing
    # PNG con transparencia (esquinas transparentes, círculo con degradado, zona semitransparente)
    $bmp = New-Object System.Drawing.Bitmap 256, 256, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.Clear([System.Drawing.Color]::Transparent)
    $rect = New-Object System.Drawing.Rectangle 8, 8, 240, 240
    $br = New-Object System.Drawing.Drawing2D.LinearGradientBrush $rect, ([System.Drawing.Color]::FromArgb(255, 220, 60, 40)), ([System.Drawing.Color]::FromArgb(255, 40, 80, 220)), 45.0
    $g.FillEllipse($br, $rect)
    $semi = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(110, 255, 255, 255))
    $g.FillRectangle($semi, 60, 100, 136, 56)
    $fuente = New-Object System.Drawing.Font 'Arial', 28, ([System.Drawing.FontStyle]::Bold)
    $g.DrawString('PNG', $fuente, [System.Drawing.Brushes]::Black, 78, 108)
    $g.Dispose()
    $bmp.Save((Join-Path $Recursos 'prueba.png'), [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    # JPG (320x200) con franjas y texto
    $bmp = New-Object System.Drawing.Bitmap 320, 200
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $colores = @(@(230, 60, 60), @(240, 160, 40), @(240, 230, 60), @(60, 180, 80), @(60, 120, 220), @(140, 70, 190))
    for ($i = 0; $i -lt 6; $i++) {
        $c = $colores[$i]
        $b = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb($c[0], $c[1], $c[2]))
        $g.FillRectangle($b, $i * 54, 0, 54, 200)
    }
    $fuente = New-Object System.Drawing.Font 'Arial', 30, ([System.Drawing.FontStyle]::Bold)
    $g.DrawString('JPG', $fuente, [System.Drawing.Brushes]::White, 100, 70)
    $g.DrawRectangle([System.Drawing.Pens]::Black, 0, 0, 319, 199)
    $g.Dispose()
    $bmp.Save((Join-Path $Recursos 'prueba.jpg'), [System.Drawing.Imaging.ImageFormat]::Jpeg)
    $bmp.Dispose()
    # mosaico pequeño (40x40)
    $bmp = New-Object System.Drawing.Bitmap 40, 40
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear([System.Drawing.Color]::FromArgb(255, 250, 240, 200))
    $g.FillRectangle([System.Drawing.Brushes]::SteelBlue, 0, 0, 20, 20)
    $g.FillRectangle([System.Drawing.Brushes]::Tomato, 20, 20, 20, 20)
    $g.DrawLine([System.Drawing.Pens]::Black, 0, 0, 40, 40)
    $g.Dispose()
    $bmp.Save((Join-Path $Recursos 'mosaico.png'), [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}

# ================================================================ inicio
$procPrevios = @(Get-Process -Name POWERPNT, EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$ppt = $null
try {
    Crear-Imagenes
    $PNG = Join-Path $Recursos 'prueba.png'
    $JPG = Join-Path $Recursos 'prueba.jpg'
    $MOS = Join-Path $Recursos 'mosaico.png'

    $ppt = New-Object -ComObject PowerPoint.Application
    try { $ppt.DisplayAlerts = 1 } catch { }   # ppAlertsNone: sin diálogos
    Write-Host ("PowerPoint " + $ppt.Version)

    # ============================================================ 01 formas
    Nuevo-Archivo '01-formas.pptx' {
        param($pres)
        $d = $null; $enDiap = 0; $primero = 0; $saltadas = New-Object System.Collections.ArrayList
        for ($t = 1; $t -le 183; $t++) {
            if ($d -eq $null) { $d = Nueva-Diap $pres 'Formas'; $enDiap = 0; $primero = 0 }
            $c = $enDiap % 8; $f = [math]::Floor($enDiap / 8)
            try {
                $s = $d.Shapes.AddShape($t, 14 + $c * 117, 36 + $f * 100, 96, 64)
            } catch { [void]$saltadas.Add($t); continue }
            $s.Name = "Forma_$t"
            $s.TextFrame.TextRange.Text = "$t"
            $s.TextFrame.TextRange.Font.Size = 8
            if ($primero -eq 0) { $primero = $t }
            $enDiap++
            $d.Shapes.Item('Rotulo').TextFrame.TextRange.Text = "Formas (MsoAutoShapeType $primero–$t)"
            if ($enDiap -ge 40) { $d = $null }
        }
        if ($saltadas.Count -gt 0) { Nota ("tipos AutoShape rechazados por AddShape: " + ($saltadas -join ',')) }

        # Adjustments extremos
        $d = Nueva-Diap $pres 'Formas con Adjustments extremos'
        $casos = @(
            @(5, 'roundRect adj=0', @(0.0)), @(5, 'roundRect adj=0.5', @(0.5)), @(5, 'roundRect adj=1 (fuera de rango)', @(1.0)),
            @(33, 'flecha der. adj=0.05/0.2', @(0.05, 0.2)), @(33, 'flecha der. adj=1/1.5', @(1.0, 1.5)),
            @(92, 'estrella 5 adj=0', @(0.0)), @(92, 'estrella 5 adj=0.5', @(0.5)), @(94, 'estrella 24 adj=0.5', @(0.5)),
            @(105, 'llamada rect. lejos', @(-1.5, 2.0)), @(106, 'llamada redond. lejos', @(1.8, -1.2)), @(107, 'llamada ovalo', @(-1.0, 1.5)),
            @(52, 'chevron adj=0.9', @(0.9)), @(11, 'cruz adj=0.5', @(0.5)), @(18, 'donut adj=0.5', @(0.5)),
            @(25, 'arco 0/0.9', @(0.0, 0.9)), @(24, 'luna adj=0.9', @(0.9)), @(15, 'bisel adj=0.5', @(0.5)), @(21, 'corazon', @(0.5)),
            @(60, 'flecha circular', @(0.9, 0.1)), @(34, 'flecha izq. adj=0.9/0.1', @(0.9, 0.1))
        )
        $i = 0
        foreach ($cs in $casos) {
            $x = 20 + ($i % 5) * 185; $y = 50 + [math]::Floor($i / 5) * 118
            Caso ("Adjustments: " + $cs[1]) {
                $s = $d.Shapes.AddShape($cs[0], $x + 20, $y + 12, 120, 70)
                $s.Name = "Adj_$i"
                $n = $s.Adjustments.Count
                for ($k = 0; $k -lt $cs[2].Count; $k++) {
                    if (($k + 1) -le $n) { $s.Adjustments.Item($k + 1) = [single]$cs[2][$k] }
                }
                $null = Etiqueta $d $x ($y + 92) 170 ($cs[1] + " (n=$n)")
            }
            $i++
        }
    }

    # ============================================================ 02 texto
    Nuevo-Archivo '02-texto.pptx' {
        param($pres)
        # --- fuentes x tamaños
        $fuentes = @('Calibri', 'Arial', 'Times New Roman', 'Aptos', 'Segoe UI', 'Georgia')
        $d1 = Nueva-Diap $pres 'Fuentes 10/14/18 pt'
        $d2 = Nueva-Diap $pres 'Fuentes 28/44 pt'
        for ($i = 0; $i -lt $fuentes.Count; $i++) {
            $x = 14 + ($i % 3) * 315; $y = 34 + [math]::Floor($i / 3) * 250
            $fn = $fuentes[$i]
            Caso "Fuente $fn" {
                $s = Texto $d1 $x $y 300 100 "$fn 10: Ñandú áéíóú`r$fn 14: Ñandú áéíóú`r$fn 18: Ñandú áéíóú"
                $s.Name = "Fuente10-18_$fn"
                $s.TextFrame.TextRange.Font.Name = $fn
                $s.TextFrame.TextRange.Paragraphs(1).Font.Size = 10
                $s.TextFrame.TextRange.Paragraphs(2).Font.Size = 14
                $s.TextFrame.TextRange.Paragraphs(3).Font.Size = 18
                $s2 = Texto $d2 $x $y 300 160 "${fn}`rÑandú 28`rÑandú 44"
                $s2.Name = "Fuente28-44_$fn"
                $s2.TextFrame.TextRange.Font.Name = $fn
                $s2.TextFrame.TextRange.Paragraphs(1).Font.Size = 12
                $s2.TextFrame.TextRange.Paragraphs(2).Font.Size = 28
                $s2.TextFrame.TextRange.Paragraphs(3).Font.Size = 44
            }
        }
        # --- alineación
        $d = Nueva-Diap $pres 'Alineación: izquierda, centro, derecha, justificado'
        $al = @(@(1, 'Izquierda'), @(2, 'Centro'), @(3, 'Derecha'), @(4, 'Justificado'))
        for ($i = 0; $i -lt 4; $i++) {
            $a = $al[$i]
            Caso ("Alineación " + $a[1]) {
                $s = Caja $d 1 (20 + ($i % 2) * 470) (40 + [math]::Floor($i / 2) * 250) 450 230 ($a[1] + ': ' + $ParrafoLargo + ' ' + $ParrafoLargo) 14
                $s.TextFrame2.TextRange.ParagraphFormat.Alignment = $a[0]
                $s.TextFrame2.VerticalAnchor = 1
                $s.Fill.ForeColor.RGB = (RGB 245 245 245); $s.Line.ForeColor.RGB = (RGB 100 100 100)
                $s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
            }
        }
        # --- interlineado
        $d = Nueva-Diap $pres 'Interlineado 0.9 / 1.0 / 1.5 / exacto 30 pt'
        $il = @(@('0.9', -1, 0.9), @('1.0', -1, 1.0), @('1.5', -1, 1.5), @('exacto 30 pt', 0, 30))
        for ($i = 0; $i -lt 4; $i++) {
            $a = $il[$i]
            Caso ("Interlineado " + $a[0]) {
                $s = Caja $d 1 (20 + $i * 232) 40 222 440 ("Interlineado " + $a[0] + ': ' + $ParrafoLargo) 14
                $s.TextFrame2.VerticalAnchor = 1
                $pf = $s.TextFrame2.TextRange.ParagraphFormat
                $pf.LineRuleWithin = $a[1]
                $pf.SpaceWithin = [single]$a[2]
                $s.Fill.ForeColor.RGB = (RGB 250 250 235); $s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
            }
        }
        # --- espacio antes / después, sangrías
        $d = Nueva-Diap $pres 'Espacio antes/después y sangrías'
        Caso 'Espacio antes 12 pt / después 6 pt' {
            $s = Caja $d 1 20 40 450 230 "Espacio antes 12 pt y después 6 pt`rPárrafo dos con el mismo espaciado.`rPárrafo tres con el mismo espaciado." 14
            $s.TextFrame2.VerticalAnchor = 1
            $pf = $s.TextFrame2.TextRange.ParagraphFormat
            $pf.LineRuleBefore = 0; $pf.SpaceBefore = 12; $pf.LineRuleAfter = 0; $pf.SpaceAfter = 6
            $s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
        }
        Caso 'Espacio antes/después en líneas (0.5 / 1.0)' {
            $s = Caja $d 1 490 40 450 230 "Espacio antes 0.5 líneas y después 1 línea`rPárrafo dos.`rPárrafo tres." 14
            $s.TextFrame2.VerticalAnchor = 1
            $pf = $s.TextFrame2.TextRange.ParagraphFormat
            $pf.LineRuleBefore = -1; $pf.SpaceBefore = 0.5; $pf.LineRuleAfter = -1; $pf.SpaceAfter = 1
            $s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
        }
        Caso 'Sangría izquierda 36 pt' {
            $s = Caja $d 1 20 290 450 230 ("Sangría izquierda 36 pt: " + $ParrafoLargo) 14
            $s.TextFrame2.VerticalAnchor = 1
            $s.TextFrame2.TextRange.ParagraphFormat.LeftIndent = 36
            $s.TextFrame2.TextRange.ParagraphFormat.FirstLineIndent = 0
            $s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
        }
        Caso 'Sangría francesa 36 pt' {
            $s = Caja $d 1 490 290 450 105 ("Sangría francesa 36 pt: " + $ParrafoLargo) 14
            $s.TextFrame2.VerticalAnchor = 1
            $s.TextFrame2.TextRange.ParagraphFormat.LeftIndent = 36
            $s.TextFrame2.TextRange.ParagraphFormat.FirstLineIndent = -36
            $s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
        }
        Caso 'Sangría de primera línea +30 pt' {
            $s = Caja $d 1 490 405 450 115 ("Primera línea +30 pt: " + $ParrafoLargo) 14
            $s.TextFrame2.VerticalAnchor = 1
            $s.TextFrame2.TextRange.ParagraphFormat.LeftIndent = 0
            $s.TextFrame2.TextRange.ParagraphFormat.FirstLineIndent = 30
            $s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
        }

        # --- estilos de carácter
        $d = Nueva-Diap $pres 'Estilos de carácter'
        $est = @(
            @('Negrita', { param($f) $f.Bold = $msoTrue }), @('Cursiva', { param($f) $f.Italic = $msoTrue }),
            @('Negrita y cursiva', { param($f) $f.Bold = $msoTrue; $f.Italic = $msoTrue }),
            @('Tachado simple', { param($f) $f.Strike = 1 }), @('Tachado doble', { param($f) $f.Strike = 2 }),
            @('Superíndice (x2)', { param($f) $f.Superscript = $msoTrue }), @('Subíndice (H2O)', { param($f) $f.Subscript = $msoTrue }),
            @('Versalitas minúsculas', { param($f) $f.Smallcaps = $msoTrue }), @('Mayúsculas', { param($f) $f.Allcaps = $msoTrue }),
            @('Espaciado -2 pt', { param($f) $f.Spacing = -2 }), @('Espaciado +5 pt', { param($f) $f.Spacing = 5 }),
            @('Espaciado +12 pt', { param($f) $f.Spacing = 12 }),
            @('Color rojo', { param($f) $f.Fill.ForeColor.RGB = (RGB 220 30 30) }),
            @('Color azul tema (Accent 1)', { param($f) $f.Fill.ForeColor.ObjectThemeColor = 5 }),
            @('Color con aclarado', { param($f) $f.Fill.ForeColor.ObjectThemeColor = 6; $f.Fill.ForeColor.Brightness = 0.4 }),
            @('Resaltado amarillo', { param($f) $f.Highlight.RGB = (RGB 255 255 0) }),
            @('Kerning 14 + Negrita', { param($f) $f.Kerning = 14; $f.Bold = $msoTrue }),
            @('Contorno de texto', { param($f) $f.Line.Visible = $msoTrue; $f.Line.ForeColor.RGB = (RGB 200 0 0); $f.Line.Weight = 1 })
        )
        for ($i = 0; $i -lt $est.Count; $i++) {
            $e = $est[$i]
            Caso ("Estilo " + $e[0]) {
                $s = Texto $d (20 + ($i % 3) * 310) (40 + [math]::Floor($i / 3) * 62) 295 40 $e[0] 18
                $s.Name = "Estilo_$i"
                $accion = $e[1]; & $accion $s.TextFrame2.TextRange.Font
            }
        }
        # --- subrayados
        $d = Nueva-Diap $pres 'Subrayados (MsoTextUnderlineType 1–17)'
        $nombresSub = @('Palabras', 'Simple', 'Doble', 'Grueso', 'Punteado', 'Punt. grueso', 'Guiones', 'Guiones grueso', 'Guion largo', 'Guion largo grueso', 'Punto-raya', 'Punto-raya grueso', 'Punto-punto-raya', 'P-P-raya grueso', 'Ondulado', 'Ondulado grueso', 'Ondulado doble')
        for ($i = 0; $i -lt 17; $i++) {
            $n = $i + 1
            Caso ("Subrayado tipo $n") {
                $s = Texto $d (20 + ($i % 3) * 310) (40 + [math]::Floor($i / 3) * 62) 295 40 ("$n " + $nombresSub[$i] + ' abc') 18
                $s.TextFrame2.TextRange.Font.UnderlineStyle = $n
                if ($n -ge 15) { $s.TextFrame2.TextRange.Font.UnderlineColor.RGB = (RGB 200 0 0) }
            }
        }
        # --- tamaños mixtos, colores en el mismo párrafo, salto manual
        $d = Nueva-Diap $pres 'Tamaños y colores mixtos, salto de línea manual'
        Caso 'Tamaños mixtos en un párrafo' {
            $s = Texto $d 20 40 920 100 'Pequeño mediano GRANDE enorme y otra vez pequeño' 12
            $tr = $s.TextFrame2.TextRange
            $tr.Characters(9, 7).Font.Size = 20
            $tr.Characters(17, 6).Font.Size = 36
            $tr.Characters(24, 6).Font.Size = 54
        }
        Caso 'Colores mixtos en un párrafo' {
            $s = Texto $d 20 160 920 50 'Rojo, verde, azul y negro en la misma línea' 24
            $tr = $s.TextFrame2.TextRange
            $tr.Characters(1, 4).Font.Fill.ForeColor.RGB = (RGB 220 0 0)
            $tr.Characters(7, 5).Font.Fill.ForeColor.RGB = (RGB 0 150 0)
            $tr.Characters(14, 4).Font.Fill.ForeColor.RGB = (RGB 0 0 220)
        }
        Caso 'Salto de línea manual (vertical tab)' {
            $s = Texto $d 20 240 450 120 ("Línea uno" + [char]11 + "línea dos tras salto manual" + [char]11 + "línea tres`rPárrafo nuevo tras retorno") 18
            $s.Name = 'SaltoManual'
        }
        Caso 'Texto con tabulaciones' {
            $s = Texto $d 490 240 450 120 ("Col A`tCol B`tCol C`r1`t2`t3`rxx`tyy`tzz") 18
        }
        # --- orientación y rotación
        $d = Nueva-Diap $pres 'Texto vertical (Orientation) y rotado'
        $ori = @(@(2, 'Arriba (2)'), @(3, 'Abajo (3)'), @(5, 'Apilado (5)'), @(4, 'Vertical Este-Asia (4)'))
        for ($i = 0; $i -lt $ori.Count; $i++) {
            $o = $ori[$i]
            Caso ("Orientación " + $o[1]) {
                $s = Caja $d 1 (20 + $i * 150) 40 130 300 ("Texto " + $o[1] + ' Ñandú') 18
                $s.TextFrame2.Orientation = $o[0]
            }
        }
        Caso 'Forma rotada 15 con texto' {
            $s = Caja $d 1 640 60 200 80 'Rotada 15 grados' 18; $s.Rotation = 15
        }
        Caso 'Forma rotada 90 con texto' {
            $s = Caja $d 1 640 180 200 60 'Rotada 90 grados' 18; $s.Rotation = 90
        }
        Caso 'Cuadro de texto rotado 330' {
            $s = Texto $d 640 360 250 40 'Cuadro de texto rotado 330' 18; $s.Rotation = 330
        }
        # --- columnas
        $d = Nueva-Diap $pres 'Dos y tres columnas'
        Caso 'Dos columnas' {
            $s = Caja $d 1 20 40 450 300 ($ParrafoLargo + ' ' + $ParrafoLargo + ' ' + $ParrafoLargo) 14
            $s.TextFrame2.VerticalAnchor = 1
            $s.TextFrame2.Column.Number = 2; $s.TextFrame2.Column.Spacing = 24
            $s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
        }
        Caso 'Tres columnas' {
            $s = Caja $d 1 490 40 450 300 ($ParrafoLargo + ' ' + $ParrafoLargo + ' ' + $ParrafoLargo + ' ' + $ParrafoLargo) 14
            $s.TextFrame2.VerticalAnchor = 1
            $s.TextFrame2.Column.Number = 3; $s.TextFrame2.Column.Spacing = 12
            $s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
        }
        # --- anclaje y márgenes
        $d = Nueva-Diap $pres 'Anclaje vertical/horizontal y márgenes'
        $anc = @(@(1, 'Arriba'), @(3, 'Medio'), @(4, 'Abajo'))
        for ($i = 0; $i -lt 3; $i++) {
            $a = $anc[$i]
            Caso ("Anclaje " + $a[1]) {
                $s = Caja $d 1 (20 + $i * 310) 40 295 150 ("Anclaje " + $a[1]) 18
                $s.TextFrame2.VerticalAnchor = $a[0]
            }
        }
        Caso 'Anclaje centrado horizontal (AnchorHorizontal)' {
            $s = Caja $d 1 20 210 295 150 'Ancla horizontal centrada, medio' 18
            $s.TextFrame2.VerticalAnchor = 3; $s.TextFrame2.HorizontalAnchor = 2
            $s.TextFrame2.TextRange.ParagraphFormat.Alignment = 1
        }
        Caso 'Márgenes internos 0 / 40' {
            $s = Caja $d 1 330 210 295 150 'Márgenes izq 40, der 0, arriba 30, abajo 0' 18
            $s.TextFrame2.MarginLeft = 40; $s.TextFrame2.MarginRight = 0; $s.TextFrame2.MarginTop = 30; $s.TextFrame2.MarginBottom = 0
            $s.TextFrame2.VerticalAnchor = 1; $s.TextFrame2.TextRange.ParagraphFormat.Alignment = 1
        }
        Caso 'Sin márgenes' {
            $s = Caja $d 1 640 210 295 150 'Todos los márgenes en 0' 18
            $s.TextFrame2.MarginLeft = 0; $s.TextFrame2.MarginRight = 0; $s.TextFrame2.MarginTop = 0; $s.TextFrame2.MarginBottom = 0
            $s.TextFrame2.VerticalAnchor = 1; $s.TextFrame2.TextRange.ParagraphFormat.Alignment = 1
        }
        # --- sin ajuste de línea
        $d = Nueva-Diap $pres 'Sin ajuste de línea (WordWrap off)'
        Caso 'WordWrap apagado, forma estrecha' {
            $s = Caja $d 1 20 50 200 80 'Este texto largo no se ajusta a la forma y se sale por los lados' 18
            $s.TextFrame2.WordWrap = $msoFalse; $s.TextFrame2.AutoSize = 0
        }
        Caso 'WordWrap apagado, alineado a la izquierda' {
            $s = Caja $d 1 20 180 200 80 'Texto largo sin ajuste alineado a la izquierda de la forma' 18
            $s.TextFrame2.WordWrap = $msoFalse; $s.TextFrame2.AutoSize = 0; $s.TextFrame2.TextRange.ParagraphFormat.Alignment = 1
        }
        Caso 'WordWrap apagado, alineado a la derecha' {
            $s = Caja $d 1 500 50 200 80 'Texto largo sin ajuste alineado a la derecha de la forma' 18
            $s.TextFrame2.WordWrap = $msoFalse; $s.TextFrame2.AutoSize = 0; $s.TextFrame2.TextRange.ParagraphFormat.Alignment = 3
        }
        Caso 'Cuadro de texto sin ajuste (crece a lo ancho)' {
            $s = Texto $d 20 320 200 30 'Cuadro de texto sin ajuste que crece hacia la derecha' 18
            $s.TextFrame2.WordWrap = $msoFalse
        }
        # --- desbordes
        $d = Nueva-Diap $pres 'Párrafos muy largos que desbordan la forma'
        $muyLargo = ((1..8 | ForEach-Object { $ParrafoLargo }) -join "`r")
        $anc2 = @(@(1, 'arriba'), @(3, 'medio'), @(4, 'abajo'))
        for ($i = 0; $i -lt 3; $i++) {
            $a = $anc2[$i]
            Caso ("Desborde anclaje " + $a[1]) {
                $s = Caja $d 1 (20 + $i * 310) 150 295 120 ($a[1] + ': ' + $muyLargo) 12
                $s.TextFrame2.AutoSize = 0; $s.TextFrame2.VerticalAnchor = $a[0]
                $s.Fill.ForeColor.RGB = (RGB 255 240 240); $s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
            }
        }
        Caso 'Desborde sin anclaje definido, texto enorme' {
            $s = Caja $d 1 20 400 600 60 ('44 pt desborda: ' + $ParrafoLargo) 44
            $s.TextFrame2.AutoSize = 0
        }
    }

    # ============================================================ 03 autoajuste
    Nuevo-Archivo '03-autoajuste.pptx' {
        param($pres)
        $vol = @(15, 25, 35, 45, 55, 70, 100, 160)
        $d = Nueva-Diap $pres 'Reducir texto al desbordar: volúmenes 15–160 palabras (18 pt)'
        for ($i = 0; $i -lt $vol.Count; $i++) {
            Caso ("Autoajuste texto, " + $vol[$i] + " palabras") {
                $s = Caja $d 1 (14 + ($i % 4) * 233) (40 + [math]::Floor($i / 4) * 245) 225 230 '' 18
                $s.Name = "Reducir_$($vol[$i])"
                $s.TextFrame2.VerticalAnchor = 1
                $s.TextFrame2.TextRange.Text = (Palabras $vol[$i])
                $s.TextFrame2.TextRange.Font.Size = 18
                $s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
                $s.TextFrame2.AutoSize = 2   # al final: asignar la fuente después apaga el autoajuste
                $s.Fill.ForeColor.RGB = (RGB 235 245 255)
            }
        }
        $d = Nueva-Diap $pres 'Reducir texto: varios párrafos con espaciado (lnSpcReduction)'
        $pv = @(3, 5, 7, 9, 11, 14)
        for ($i = 0; $i -lt $pv.Count; $i++) {
            Caso ("Autoajuste texto, " + $pv[$i] + " párrafos") {
                $s = Caja $d 1 (14 + ($i % 3) * 312) (40 + [math]::Floor($i / 3) * 245) 300 230 '' 16
                $s.Name = "ReducirPars_$($pv[$i])"
                $s.TextFrame2.VerticalAnchor = 1
                $txt = (1..$pv[$i] | ForEach-Object { "Párrafo ${_}: " + (Palabras 9) }) -join "`r"
                $s.TextFrame2.TextRange.Text = $txt
                $s.TextFrame2.TextRange.Font.Size = 16
                $s.TextFrame2.TextRange.ParagraphFormat.SpaceAfter = 6
                $s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
                $s.TextFrame2.AutoSize = 2   # al final: asignar la fuente después apaga el autoajuste
                $s.Fill.ForeColor.RGB = (RGB 235 255 235)
            }
        }
        $d = Nueva-Diap $pres 'Ajustar forma al texto (msoAutoSizeShapeToFitText)'
        $vol2 = @(5, 20, 40, 80)
        for ($i = 0; $i -lt $vol2.Count; $i++) {
            Caso ("Forma al texto, " + $vol2[$i] + " palabras") {
                $s = Caja $d 1 (14 + $i * 233) 50 225 40 '' 14
                $s.Name = "FormaAlTexto_$($vol2[$i])"
                $s.TextFrame2.WordWrap = $msoTrue
                $s.TextFrame2.VerticalAnchor = 1
                $s.TextFrame2.TextRange.Text = (Palabras $vol2[$i])
                $s.TextFrame2.TextRange.Font.Size = 14
                $s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
                $s.TextFrame2.AutoSize = 1   # al final: asignar la fuente después apaga el autoajuste
                $s.Fill.ForeColor.RGB = (RGB 255 250 220)
            }
        }
        Caso 'Cuadro de texto que crece (AutoSize ShapeToFit) en una línea' {
            $s = Texto $d 14 400 300 30 'Texto corto' 24
            $s.TextFrame2.AutoSize = 1; $s.TextFrame2.WordWrap = $msoFalse
        }
        $d = Nueva-Diap $pres 'Sin ajuste con desborde (msoAutoSizeNone)'
        for ($i = 0; $i -lt 3; $i++) {
            Caso ("Sin ajuste, desborde " + $i) {
                $s = Caja $d 1 (14 + $i * 312) 50 300 120 '' 18
                $s.Name = "SinAjuste_$i"
                $s.TextFrame2.AutoSize = 0
                $s.TextFrame2.VerticalAnchor = (1, 3, 4)[$i]
                $s.TextFrame2.TextRange.Text = (Palabras 90)
                $s.TextFrame2.TextRange.Font.Size = 18
                $s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
                $s.Fill.ForeColor.RGB = (RGB 255 235 235)
            }
        }
    } 960 540 $true   # con ventana (minimizada): sin ella PowerPoint no calcula fontScale

    # ============================================================ 04 viñetas
    Nuevo-Archivo '04-vinetas.pptx' {
        param($pres)
        # Lista de 4 elementos en un cuadro de texto; $cfg configura la viñeta de cada párrafo.
        function Lista($d, [double]$x, [double]$y, [double]$w, [double]$h, [string]$titulo, [string[]]$lineas, [scriptblock]$cfg) {
            $s = Caja $d 1 $x $y $w $h '' 14
            $s.TextFrame.TextRange.Text = (($titulo + '|') + ($lineas -join '|')) -replace '\|', "`r"
            $s.TextFrame.VerticalAnchor = 1
            $s.TextFrame.TextRange.ParagraphFormat.Alignment = 1
            $s.TextFrame.TextRange.Font.Color.RGB = 0
            $s.Fill.ForeColor.RGB = (RGB 250 250 250)
            $s.TextFrame.TextRange.Paragraphs(1).Font.Bold = $msoTrue
            $s.TextFrame.TextRange.Paragraphs(1).ParagraphFormat.Bullet.Visible = $msoFalse
            for ($i = 2; $i -le $s.TextFrame.TextRange.Paragraphs().Count; $i++) {
                $p = $s.TextFrame.TextRange.Paragraphs($i)
                & $cfg $p $i
            }
            for ($n = 1; $n -le 5; $n++) {
                $s.TextFrame.Ruler.Levels($n).FirstMargin = ($n - 1) * 27
                $s.TextFrame.Ruler.Levels($n).LeftMargin = ($n - 1) * 27 + 22
            }
            return $s
        }
        $d = Nueva-Diap $pres 'Viñetas de caracteres'
        $chars = @(@('Bullet U+2022', 8226, 'Arial'), @('Wingdings v (rombos)', 118, 'Wingdings'), @('Wingdings § (cuadrado)', 167, 'Wingdings'),
            @('Wingdings Ø (flecha)', 216, 'Wingdings'), @('Guion U+2013', 8211, 'Arial'), @('Check Wingdings 252', 252, 'Wingdings'))
        for ($i = 0; $i -lt $chars.Count; $i++) {
            $c = $chars[$i]
            Caso ("Viñeta " + $c[0]) {
                $null = Lista $d (14 + ($i % 3) * 312) (40 + [math]::Floor($i / 3) * 240) 300 220 $c[0] @('Primer elemento', 'Segundo elemento', 'Tercer elemento con texto largo que ocupa más de una línea del cuadro') {
                    param($p, $k)
                    $b = $p.ParagraphFormat.Bullet
                    $b.Visible = $msoTrue; $b.Type = 1; $b.UseTextFont = $msoFalse; $b.Font.Name = $c[2]; $b.Character = $c[1]
                }
            }
        }
        $d = Nueva-Diap $pres 'Viñeta con color y tamaño propios'
        Caso 'Viñeta roja 150%' {
            $null = Lista $d 14 40 300 220 'Roja, 150 %' @('Uno', 'Dos', 'Tres') {
                param($p, $k)
                $b = $p.ParagraphFormat.Bullet
                $b.Visible = $msoTrue; $b.Type = 1; $b.Character = 8226; $b.UseTextColor = $msoFalse; $b.Font.Color.RGB = (RGB 220 0 0); $b.RelativeSize = 1.5
            }
        }
        Caso 'Viñeta azul 60%, texto verde' {
            $null = Lista $d 326 40 300 220 'Azul, 60 %, texto verde' @('Uno', 'Dos', 'Tres') {
                param($p, $k)
                $b = $p.ParagraphFormat.Bullet
                $p.Font.Color.RGB = (RGB 0 140 0)
                $b.Visible = $msoTrue; $b.Type = 1; $b.Character = 8226; $b.UseTextColor = $msoFalse; $b.Font.Color.RGB = (RGB 0 0 220); $b.RelativeSize = 0.6
            }
        }
        Caso 'Viñeta con fuente distinta (Times)' {
            $null = Lista $d 638 40 300 220 'Fuente Times New Roman, tamaño 24' @('Uno', 'Dos', 'Tres') {
                param($p, $k)
                $b = $p.ParagraphFormat.Bullet
                $p.Font.Size = 24
                $b.Visible = $msoTrue; $b.Type = 1; $b.UseTextFont = $msoFalse; $b.Font.Name = 'Times New Roman'; $b.Character = 9632
            }
        }
        $d = Nueva-Diap $pres 'Numeración: estilos'
        $est = @(@('1.  (ArabicPeriod=3)', 3), @('1)  (ArabicParenRight=2)', 2), @('a.  (AlphaLCPeriod=0)', 0), @('A.  (AlphaUCPeriod=1)', 1),
            @('i.  (RomanLCPeriod=6)', 6), @('I.  (RomanUCPeriod=7)', 7), @('(1) (ArabicParenBoth=12)', 12), @('a) (AlphaLCParenRight=9)', 9), @('(i) (RomanLCParenBoth=4)', 4))
        for ($i = 0; $i -lt $est.Count; $i++) {
            $e = $est[$i]
            Caso ("Numeración " + $e[0]) {
                $null = Lista $d (14 + ($i % 3) * 312) (36 + [math]::Floor($i / 3) * 165) 300 155 $e[0] @('Primero', 'Segundo', 'Tercero', 'Cuarto') {
                    param($p, $k)
                    $b = $p.ParagraphFormat.Bullet
                    $b.Visible = $msoTrue; $b.Type = 2; $b.Style = $e[1]; $b.StartValue = 1
                }
            }
        }
        $d = Nueva-Diap $pres 'Numeración: StartValue, párrafos vacíos y niveles'
        Caso 'StartValue = 5' {
            $null = Lista $d 14 40 300 200 'Empieza en 5' @('Quinto', 'Sexto', 'Séptimo') {
                param($p, $k)
                $b = $p.ParagraphFormat.Bullet
                $b.Visible = $msoTrue; $b.Type = 2; $b.Style = 3; $b.StartValue = 5
            }
        }
        Caso 'StartValue = 10 con letras (k.)' {
            $null = Lista $d 326 40 300 200 'Letras desde la 11.ª (k.)' @('Letra k', 'Letra l', 'Letra m') {
                param($p, $k)
                $b = $p.ParagraphFormat.Bullet
                $b.Visible = $msoTrue; $b.Type = 2; $b.Style = 0; $b.StartValue = 11
            }
        }
        Caso 'Párrafos vacíos en medio de una lista numerada' {
            $s = Lista $d 638 40 300 200 'Vacíos intercalados' @('Uno', '', 'Dos', '', '', 'Tres') {
                param($p, $k)
                $b = $p.ParagraphFormat.Bullet
                $b.Visible = $msoTrue; $b.Type = 2; $b.Style = 3; $b.StartValue = 1
            }
            $s.Name = 'NumeradaConVacios'
        }
        Caso 'Cinco niveles de viñeta (IndentLevel 1–5)' {
            $car = @(8226, 8211, 9642, 9702, 187)
            $s = Lista $d 14 260 450 255 'Niveles 1 a 5' @('Nivel 1', 'Nivel 2', 'Nivel 3', 'Nivel 4', 'Nivel 5', 'Otra vez nivel 3', 'Otra vez nivel 1') {
                param($p, $k)
                $niveles = @(1, 2, 3, 4, 5, 3, 1)
                $p.IndentLevel = $niveles[$k - 2]
                $b = $p.ParagraphFormat.Bullet
                $b.Visible = $msoTrue; $b.Type = 1; $b.Character = $car[$niveles[$k - 2] - 1]
            }
            $s.Name = 'CincoNiveles'
        }
        Caso 'Niveles con numeración mixta' {
            $s = Lista $d 480 260 450 255 'Numeración jerárquica' @('Capítulo', 'Sección a', 'Sección b', 'Subsección i', 'Subsección ii', 'Capítulo 2') {
                param($p, $k)
                $niveles = @(1, 2, 2, 3, 3, 1)
                $nv = $niveles[$k - 2]
                $p.IndentLevel = $nv
                $b = $p.ParagraphFormat.Bullet
                $b.Visible = $msoTrue; $b.Type = 2
                $b.Style = (3, 0, 6)[$nv - 1]
                $b.StartValue = 1
            }
            $s.Name = 'NumeracionMixta'
        }
        # --- marcadores de contenido de un diseño (viñetas heredadas del patrón)
        Caso 'Diseño Título y objetos: viñetas heredadas' {
            $d = Nueva-DiapDiseno $pres 'Marcador de contenido con viñetas del patrón' 2
            $d.Shapes.Placeholders.Item(1).TextFrame.TextRange.Text = 'Viñetas heredadas del patrón'
            $cuerpo = $d.Shapes.Placeholders.Item(2)
            $cuerpo.TextFrame.TextRange.Text = "Nivel 1: primer punto`rNivel 2: subpunto`rNivel 3: sub-subpunto`rNivel 4: más profundo`rNivel 5: lo más profundo`rVuelve al nivel 1"
            $niv = @(1, 2, 3, 4, 5, 1)
            for ($i = 1; $i -le 6; $i++) { $cuerpo.TextFrame.TextRange.Paragraphs($i).IndentLevel = $niv[$i - 1] }
        }
        Caso 'Diseño Dos objetos: dos marcadores con viñetas' {
            $d = Nueva-DiapDiseno $pres 'Dos contenidos con viñetas del patrón' 4
            $d.Shapes.Placeholders.Item(1).TextFrame.TextRange.Text = 'Dos contenidos'
            $d.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "Izquierda uno`rIzquierda dos`rIzquierda tres con un texto largo para que ocupe más de una línea"
            $d.Shapes.Placeholders.Item(3).TextFrame.TextRange.Text = "Derecha uno`rDerecha dos`rDerecha tres"
            $d.Shapes.Placeholders.Item(3).TextFrame.TextRange.Paragraphs(2).IndentLevel = 2
        }
        Caso 'Marcador con viñeta apagada y viñeta propia en un párrafo' {
            $d = Nueva-DiapDiseno $pres 'Marcador: anulaciones locales de viñeta' 2
            $d.Shapes.Placeholders.Item(1).TextFrame.TextRange.Text = 'Anulaciones locales'
            $cuerpo = $d.Shapes.Placeholders.Item(2)
            $cuerpo.TextFrame.TextRange.Text = "Con viñeta heredada`rSin viñeta (apagada)`rViñeta propia Wingdings`rNumerada local`rNumerada local 2"
            $cuerpo.TextFrame.TextRange.Paragraphs(2).ParagraphFormat.Bullet.Visible = $msoFalse
            $b = $cuerpo.TextFrame.TextRange.Paragraphs(3).ParagraphFormat.Bullet
            $b.Visible = $msoTrue; $b.Type = 1; $b.UseTextFont = $msoFalse; $b.Font.Name = 'Wingdings'; $b.Character = 118
            foreach ($k in 4, 5) { $b = $cuerpo.TextFrame.TextRange.Paragraphs($k).ParagraphFormat.Bullet; $b.Visible = $msoTrue; $b.Type = 2; $b.Style = 3; $b.StartValue = 1 }
        }
    }

    # ============================================================ 05 rellenos
    Nuevo-Archivo '05-rellenos.pptx' {
        param($pres)
        # --- sólido con transparencia
        $d = Nueva-Diap $pres 'Relleno sólido con transparencia 0–90 %'
        $fondo = Caja $d 1 20 50 920 200 ''
        $fondo.Fill.ForeColor.RGB = (RGB 40 40 40)
        for ($i = 0; $i -lt 10; $i++) {
            $x = 24 + $i * 92
            $b = Caja $d 1 $x 70 80 160 ''; $b.Fill.ForeColor.RGB = (RGB 255 255 255); $b.Line.Visible = $msoFalse
            $b.Height = 70; $b.Top = 60
            Caso ("Transparencia " + ($i * 10)) {
                $s = Caja $d 1 $x 80 80 150 ("$($i * 10)%") 12
                $s.Fill.Solid(); $s.Fill.ForeColor.RGB = (RGB 220 30 30); $s.Fill.Transparency = $i / 10.0
                $s.Line.ForeColor.RGB = (RGB 0 0 0)
            }
        }
        Caso 'Transparencia en la línea' {
            $s = Caja $d 1 20 280 200 100 'Línea 60 % transp.' 12
            $s.Line.Weight = 10; $s.Line.ForeColor.RGB = (RGB 0 0 200); $s.Line.Transparency = 0.6
            $s.Fill.ForeColor.RGB = (RGB 255 220 120)
        }
        # --- degradados lineales por ángulo
        $d = Nueva-Diap $pres 'Degradados lineales: ángulos 0–315'
        for ($i = 0; $i -lt 8; $i++) {
            $ang = $i * 45
            Caso ("Degradado lineal $ang") {
                $s = Caja $d 1 (14 + ($i % 4) * 233) (40 + [math]::Floor($i / 4) * 150) 225 140 "$ang grados" 12
                $s.Fill.TwoColorGradient(1, 1)
                $s.Fill.ForeColor.RGB = (RGB 255 220 0); $s.Fill.BackColor.RGB = (RGB 0 90 200)
                $s.Fill.GradientAngle = $ang
            }
        }
        Caso 'Estilos TwoColorGradient 1–7 variante 1' {
            Nota "msoGradientFromTitle (6) no es válido para formas: se omite el estilo 6 de TwoColorGradient"
            foreach ($g in 1, 2, 3, 4, 5, 7) {
                Caso ("TwoColorGradient estilo $g") {
                    $s = Caja $d 1 (14 + (($g - 1) % 4) * 233) (350 + [math]::Floor(($g - 1) / 4) * 90) 100 80 "estilo $g" 10
                    $s.Fill.TwoColorGradient($g, 1)
                    $s.Fill.ForeColor.RGB = (RGB 255 255 255); $s.Fill.BackColor.RGB = (RGB 200 30 60)
                }
            }
        }
        # --- radial / ruta y varias paradas
        $d = Nueva-Diap $pres 'Degradados radial, ruta (rectangular, forma) y varias paradas'
        for ($i = 0; $i -lt $rad.Count; $i++) {
            $r = $rad[$i]
            Caso ("Degradado " + $r[0]) {
                $s = Caja $d 1 (14 + $i * 233) 40 225 140 $r[0] 12
                $s.Fill.TwoColorGradient($r[1], $r[2])
                $s.Fill.ForeColor.RGB = (RGB 255 255 255); $s.Fill.BackColor.RGB = (RGB 0 120 60)
            }
        }
        Caso 'Elipse con degradado de ruta (círculo)' {
            $s = Caja $d 9 14 200 225 140 'Elipse radial' 12
            $s.Fill.TwoColorGradient(7, 1); $s.Fill.ForeColor.RGB = (RGB 255 240 120); $s.Fill.BackColor.RGB = (RGB 220 60 0)
        }
        Caso 'Cuatro paradas con colores y posiciones' {
            $s = Caja $d 1 247 200 225 140 '4 paradas' 12
            $s.Fill.TwoColorGradient(1, 1); $s.Fill.ForeColor.RGB = (RGB 255 0 0); $s.Fill.BackColor.RGB = (RGB 0 0 255)
            $s.Fill.GradientStops.Insert((RGB 255 255 0), 0.33, 0, 2)
            $s.Fill.GradientStops.Insert((RGB 0 200 0), 0.66, 0, 3)
        }
        Caso 'Paradas con transparencia distinta' {
            $fnd = Caja $d 1 480 200 225 140 ''; $fnd.Fill.Patterned(21); $fnd.Fill.ForeColor.RGB = (RGB 180 180 180); $fnd.Fill.BackColor.RGB = (RGB 255 255 255)
            $s = Caja $d 1 480 200 225 140 'Transp. 0 / 50 / 90 %' 12
            $s.Fill.TwoColorGradient(1, 1); $s.Fill.ForeColor.RGB = (RGB 0 0 255); $s.Fill.BackColor.RGB = (RGB 0 0 255)
            $s.Fill.GradientStops.Insert((RGB 0 0 255), 0.5, 0.5, 2)
            $s.Fill.GradientStops.Item(1).Transparency = 0.0
            $s.Fill.GradientStops.Item(3).Transparency = 0.9
        }
        Caso 'Parada con color de tema' {
            $s = Caja $d 1 713 200 225 140 'Paradas con tema' 12
            $s.Fill.TwoColorGradient(1, 1)
            $s.Fill.GradientStops.Item(1).Color.ObjectThemeColor = 5
            $s.Fill.GradientStops.Item(2).Color.ObjectThemeColor = 6
        }
        Caso 'Degradado con ángulo 30 y 3 paradas giradas' {
            $s = Caja $d 1 14 360 225 150 'Ángulo 30, rotada 10' 12
            $s.Fill.TwoColorGradient(1, 1); $s.Fill.ForeColor.RGB = (RGB 255 100 100); $s.Fill.BackColor.RGB = (RGB 100 100 255)
            $s.Fill.GradientStops.Insert((RGB 255 255 255), 0.5, 0, 2)
            $s.Fill.GradientAngle = 30; $s.Rotation = 10
        }
        # --- degradados preestablecidos
        $d = Nueva-Diap $pres 'Degradados preestablecidos (PresetGradient 1–24)'
        for ($i = 1; $i -le 24; $i++) {
            Caso ("PresetGradient $i") {
                $s = Caja $d 1 (14 + (($i - 1) % 6) * 155) (40 + [math]::Floor(($i - 1) / 6) * 120) 145 105 "preset $i" 10
                $s.Fill.PresetGradient(1, 1, $i)
            }
        }
        # --- tramas
        $d = Nueva-Diap $pres 'Tramas: MsoPatternType 1–48'
        for ($p = 1; $p -le 48; $p++) {
            Caso ("Trama $p") {
                $x = 14 + (($p - 1) % 8) * 117; $y = 38 + [math]::Floor(($p - 1) / 8) * 85
                $s = Caja $d 1 $x $y 100 58 ''
                $s.Fill.Patterned($p)
                $s.Fill.ForeColor.RGB = (RGB 20 60 160); $s.Fill.BackColor.RGB = (RGB 255 240 200)
                $s.Line.ForeColor.RGB = (RGB 0 0 0); $s.Line.Weight = 0.5
                $null = Etiqueta $d $x ($y + 60) 100 "trama $p"
            }
        }
        # --- imagen
        $d = Nueva-Diap $pres 'Relleno de imagen: estirado, mosaico, con escala y desplazamiento'
        Caso 'Imagen estirada' {
            $s = Caja $d 1 14 40 225 150 'estirada' 12; $s.Fill.UserPicture($PNG)
        }
        Caso 'Imagen JPG estirada en elipse' {
            $s = Caja $d 9 247 40 225 150 '' 12; $s.Fill.UserPicture($JPG)
        }
        Caso 'Mosaico automático' {
            $s = Caja $d 1 480 40 225 150 'mosaico' 12; $s.Fill.UserTextured($MOS)
        }
        Caso 'Mosaico con escala 200 % y desplazamiento' {
            $s = Caja $d 1 713 40 225 150 'mosaico 200 %' 12; $s.Fill.UserTextured($MOS)
            $s.Fill.TextureHorizontalScale = 2; $s.Fill.TextureVerticalScale = 2
            $s.Fill.TextureOffsetX = 10; $s.Fill.TextureOffsetY = 10
        }
        Caso 'Mosaico alineado a la derecha-abajo' {
            $s = Caja $d 1 14 210 225 150 'mosaico alineado' 12; $s.Fill.UserTextured($MOS)
            $s.Fill.TextureAlignment = 8   # msoTextureBottomRight (base 0)
        }
        Caso 'Imagen con transparencia 50 %' {
            $f = Caja $d 1 247 210 225 150 ''; $f.Fill.ForeColor.RGB = (RGB 0 0 0)
            $s = Caja $d 1 247 210 225 150 'imagen 50 %' 12; $s.Fill.UserPicture($PNG); $s.Fill.Transparency = 0.5
        }
        Caso 'Imagen de relleno de la diapositiva (fondo)' {
            $d.FollowMasterBackground = $msoFalse
            $d.Background.Fill.UserTextured($MOS)
        }
        # --- colores de tema con aclarar/oscurecer
        $d = Nueva-Diap $pres 'Colores de tema: Accent 1–6 con Brightness -0.5 … 0.8'
        $br = @(-0.5, -0.25, 0, 0.4, 0.8)
        for ($a = 0; $a -lt 6; $a++) {
            for ($b = 0; $b -lt 5; $b++) {
                Caso ("Tema accent$($a + 1) brillo $($br[$b])") {
                    $s = Caja $d 1 (14 + $b * 186) (38 + $a * 80) 176 70 ("accent$($a + 1)  $($br[$b])") 10
                    $s.Fill.Solid(); $s.Fill.ForeColor.ObjectThemeColor = 5 + $a; $s.Fill.ForeColor.Brightness = [single]$br[$b]
                }
            }
        }
        $d = Nueva-Diap $pres 'Colores de tema: Texto/Fondo con TintAndShade'
        $tc = @(@(13, 'Text1'), @(14, 'Background1'), @(15, 'Text2'), @(16, 'Background2'), @(1, 'Dark1'), @(2, 'Light1'), @(11, 'Hyperlink'), @(12, 'FollowedHyperlink'))
        for ($i = 0; $i -lt $tc.Count; $i++) {
            Caso ("Tema " + $tc[$i][1]) {
                $s = Caja $d 1 (14 + ($i % 4) * 233) (40 + [math]::Floor($i / 4) * 120) 225 100 ($tc[$i][1]) 12
                $s.Fill.ForeColor.ObjectThemeColor = $tc[$i][0]
                $s.Line.ForeColor.RGB = (RGB 0 0 0)
            }
        }
        Caso 'TintAndShade ±0.5 sobre Accent 2' {
            foreach ($j in 0, 1) {
                $s = Caja $d 1 (14 + $j * 233) 300 225 100 ("TintAndShade " + (0.5, -0.5)[$j]) 12
                $s.Fill.ForeColor.ObjectThemeColor = 6; $s.Fill.ForeColor.TintAndShade = [single](0.5, -0.5)[$j]
            }
        }
    }

    # ============================================================ 06 líneas
    Nuevo-Archivo '06-lineas.pptx' {
        param($pres)
        $d = Nueva-Diap $pres 'Líneas: grosores 0.25 – 20 pt'
        $gr = @(0.25, 0.5, 1, 2, 4, 8, 12, 20)
        for ($i = 0; $i -lt $gr.Count; $i++) {
            Caso ("Grosor " + $gr[$i]) {
                $l = $d.Shapes.AddLine(30, 60 + $i * 56, 600, 60 + $i * 56)
                $l.Line.Weight = [single]$gr[$i]; $l.Line.ForeColor.RGB = (RGB 20 20 20)
                $null = Texto $d 620 (48 + $i * 56) 200 24 ("grosor " + $gr[$i] + ' pt') 12
            }
        }
        $d = Nueva-Diap $pres 'Líneas: los 12 estilos de guion (DashStyle 1–12), 1.5 pt y 6 pt'
        $nomD = @('1 Sólido', '2 SquareDot', '3 RoundDot', '4 Dash', '5 DashDot', '6 DashDotDot', '7 LongDash', '8 LongDashDot', '9 LongDashDotDot', '10 SysDash', '11 SysDot', '12 SysDashDot')
        for ($i = 0; $i -lt 12; $i++) {
            Caso ("DashStyle " + ($i + 1)) {
                $y = 56 + $i * 40
                $l = $d.Shapes.AddLine(20, $y, 330, $y); $l.Line.Weight = 1.5; $l.Line.DashStyle = $i + 1; $l.Line.ForeColor.RGB = (RGB 0 0 0)
                $l2 = $d.Shapes.AddLine(350, $y, 660, $y); $l2.Line.Weight = 6; $l2.Line.DashStyle = $i + 1; $l2.Line.ForeColor.RGB = (RGB 0 0 160)
                $null = Texto $d 680 ($y - 12) 260 22 $nomD[$i] 12
            }
        }
        $d = Nueva-Diap $pres 'Líneas: estilos compuestos (Line.Style 1–5) a 8 y 12 pt'
        $nomE = @('1 Simple', '2 ThinThin', '3 ThinThick', '4 ThickThin', '5 ThickBetweenThin')
        for ($i = 0; $i -lt 5; $i++) {
            Caso ("Line.Style " + ($i + 1)) {
                $y = 70 + $i * 85
                $l = $d.Shapes.AddLine(20, $y, 330, $y); $l.Line.Weight = 8; $l.Line.Style = $i + 1; $l.Line.ForeColor.RGB = (RGB 160 0 0)
                $l2 = $d.Shapes.AddLine(350, $y, 660, $y); $l2.Line.Weight = 12; $l2.Line.Style = $i + 1; $l2.Line.ForeColor.RGB = (RGB 0 100 0)
                $null = Texto $d 680 ($y - 12) 260 22 $nomE[$i] 12
                $r = Caja $d 1 800 ($y - 25) 60 50 ''
                $r.Fill.Visible = $msoFalse; $r.Line.Weight = 8; $r.Line.Style = $i + 1; $r.Line.ForeColor.RGB = (RGB 0 0 160)
            }
        }
        # --- puntas: 5 tipos x 9 tamaños (inicio y fin, intercalados)
        $tipos = @(@(2, 'Triangle'), @(3, 'Open'), @(4, 'Stealth'), @(5, 'Diamond'), @(6, 'Oval'))
        $d1 = Nueva-Diap $pres 'Puntas de flecha: tipos 2–6 x largo 1–3 x ancho 1–3 (grosor 3)'
        for ($t = 0; $t -lt 5; $t++) {
            for ($k = 0; $k -lt 9; $k++) {
                $lg = [int]([math]::Floor($k / 3)) + 1; $an = [int]($k % 3) + 1
                $idx = $t * 9 + $k
                Caso ("Punta " + $tipos[$t][1] + " largo $lg ancho $an") {
                    $col = $idx % 3; $fila = [math]::Floor($idx / 3)
                    $x = 20 + $col * 310; $y = 50 + $fila * 31
                    $l = $d1.Shapes.AddLine($x, $y, $x + 200, $y)
                    $l.Line.Weight = 3; $l.Line.ForeColor.RGB = (RGB 20 20 20)
                    $l.Line.EndArrowheadStyle = $tipos[$t][0]; $l.Line.EndArrowheadLength = $lg; $l.Line.EndArrowheadWidth = $an
                    $l.Line.BeginArrowheadStyle = $tipos[($t + 1) % 5][0]; $l.Line.BeginArrowheadLength = $lg; $l.Line.BeginArrowheadWidth = $an
                    $null = Etiqueta $d1 ($x + 205) ($y - 6) 100 ("fin " + $tipos[$t][1] + " L$lg A$an")
                }
            }
        }
        $d = Nueva-Diap $pres 'Puntas de flecha con distintos grosores (0.5, 1, 2, 6, 10 pt) y colores'
        $gg = @(0.5, 1, 2, 6, 10)
        for ($i = 0; $i -lt 5; $i++) {
            Caso ("Punta con grosor " + $gg[$i]) {
                $y = 70 + $i * 80
                $l = $d.Shapes.AddLine(30, $y, 400, $y); $l.Line.Weight = [single]$gg[$i]; $l.Line.ForeColor.RGB = (RGB 200 0 0)
                $l.Line.EndArrowheadStyle = 2; $l.Line.BeginArrowheadStyle = 6
                $l2 = $d.Shapes.AddLine(450, $y, 820, $y); $l2.Line.Weight = [single]$gg[$i]; $l2.Line.ForeColor.RGB = (RGB 0 100 200); $l2.Line.DashStyle = 4
                $l2.Line.EndArrowheadStyle = 4; $l2.Line.EndArrowheadLength = 3; $l2.Line.EndArrowheadWidth = 3; $l2.Line.BeginArrowheadStyle = 5
                $null = Texto $d 830 ($y - 12) 120 24 ("$($gg[$i]) pt") 12
            }
        }
        # --- conectores
        $d = Nueva-Diap $pres 'Conectores: recto, de codo y curvo (libres y pegados a formas)'
        Caso 'Conectores libres' {
            $tp = @(@(1, 'Recto'), @(2, 'Codo'), @(3, 'Curvo'))
            for ($i = 0; $i -lt 3; $i++) {
                $c = $d.Shapes.AddConnector($tp[$i][0], 30, 70 + $i * 70, 300, 110 + $i * 70)
                $c.Line.Weight = 2.5; $c.Line.ForeColor.RGB = (RGB 0 0 0); $c.Line.EndArrowheadStyle = 2
                $null = Texto $d 310 (80 + $i * 70) 150 22 $tp[$i][1] 12
            }
        }
        Caso 'Conectores pegados a formas (codo y curvo)' {
            $a = Caja $d 1 500 60 110 50 'A' 14; $b = Caja $d 1 760 200 110 50 'B' 14
            $c = Caja $d 1 500 330 110 50 'C' 14; $e = Caja $d 9 760 400 110 50 'D' 14
            $c1 = $d.Shapes.AddConnector(2, 0, 0, 10, 10); $c1.ConnectorFormat.BeginConnect($a, 4); $c1.ConnectorFormat.EndConnect($b, 1)
            $c1.Line.EndArrowheadStyle = 2; $c1.Line.Weight = 2
            $c2 = $d.Shapes.AddConnector(3, 0, 0, 10, 10); $c2.ConnectorFormat.BeginConnect($c, 4); $c2.ConnectorFormat.EndConnect($e, 2)
            $c2.Line.EndArrowheadStyle = 2; $c2.Line.Weight = 2; $c2.Line.ForeColor.RGB = (RGB 200 0 0)
            $c3 = $d.Shapes.AddConnector(1, 0, 0, 10, 10); $c3.ConnectorFormat.BeginConnect($a, 3); $c3.ConnectorFormat.EndConnect($c, 1)
            $c3.Line.EndArrowheadStyle = 3; $c3.Line.DashStyle = 4; $c3.Line.Weight = 2
        }
        Caso 'Conector de codo con ajuste de punto medio' {
            $c = $d.Shapes.AddConnector(2, 30, 330, 400, 480)
            $c.Line.Weight = 2.5; $c.Line.ForeColor.RGB = (RGB 0 120 0)
            if ($c.Adjustments.Count -gt 0) { $c.Adjustments.Item(1) = 0.2 }
        }
        Caso 'Línea con ancho de contorno en formas (cuadrado con 3 estilos)' {
            $r = Caja $d 1 30 250 120 50 'Doble' 12; $r.Line.Weight = 6; $r.Line.Style = 2
            $r = Caja $d 1 180 250 120 50 'Guion' 12; $r.Line.Weight = 3; $r.Line.DashStyle = 4
            $r = Caja $d 1 330 250 120 50 'Sin línea' 12; $r.Line.Visible = $msoFalse
        }
    }

    # ============================================================ 07 efectos
    Nuevo-Archivo '07-efectos.pptx' {
        param($pres)
        function Forma-Efecto($d, [int]$i, [string]$etiqueta, [int]$cols = 5, [int]$tipo = 5) {
            $x = 24 + ($i % $cols) * (920 / $cols); $y = 44 + [math]::Floor($i / $cols) * 165
            $s = $d.Shapes.AddShape($tipo, $x, $y, 120, 80)
            $s.Fill.ForeColor.RGB = (RGB 70 130 200); $s.Line.Visible = $msoFalse
            $null = Etiqueta $d $x ($y + 120) 170 $etiqueta
            return $s
        }
        $d = Nueva-Diap $pres 'Sombra exterior: distancia, desenfoque, color, transparencia'
        $sh = @(@(0, 0, 0, 0), @(4, 4, 0, 0), @(4, 4, 6, 0), @(8, 8, 12, 0), @(12, 12, 20, 0.5), @(-6, 6, 8, 0.3), @(0, 10, 15, 0.6), @(20, 0, 30, 0.2), @(3, 3, 3, 0.8), @(10, -10, 5, 0))
        for ($i = 0; $i -lt $sh.Count; $i++) {
            $p = $sh[$i]
            Caso ("Sombra exterior " + ($p -join '/')) {
                $s = Forma-Efecto $d $i ("X$($p[0]) Y$($p[1]) blur$($p[2]) transp$($p[3])")
                $s.Shadow.Visible = $msoTrue; $s.Shadow.Style = 2
                $s.Shadow.OffsetX = [single]$p[0]; $s.Shadow.OffsetY = [single]$p[1]; $s.Shadow.Blur = [single]$p[2]; $s.Shadow.Transparency = [single]$p[3]
                $s.Shadow.ForeColor.RGB = (RGB 0 0 0)
            }
        }
        $d = Nueva-Diap $pres 'Sombra interior y sombras de color'
        $in = @(@(3, 3, 4, 0.3), @(6, 6, 10, 0.2), @(-5, 5, 8, 0.5), @(0, 0, 15, 0.3))
        for ($i = 0; $i -lt $in.Count; $i++) {
            $p = $in[$i]
            Caso ("Sombra interior " + ($p -join '/')) {
                $s = Forma-Efecto $d $i ("interior X$($p[0]) Y$($p[1]) blur$($p[2])")
                $s.Shadow.Visible = $msoTrue; $s.Shadow.Style = 1
                $s.Shadow.OffsetX = [single]$p[0]; $s.Shadow.OffsetY = [single]$p[1]; $s.Shadow.Blur = [single]$p[2]; $s.Shadow.Transparency = [single]$p[3]
                $s.Shadow.ForeColor.RGB = (RGB 0 0 0)
            }
        }
        Caso 'Sombra exterior roja' {
            $s = Forma-Efecto $d 5 'exterior roja 6/6 blur 5'
            $s.Shadow.Visible = $msoTrue; $s.Shadow.Style = 2; $s.Shadow.OffsetX = 6; $s.Shadow.OffsetY = 6; $s.Shadow.Blur = 5
            $s.Shadow.ForeColor.RGB = (RGB 220 0 0); $s.Shadow.Transparency = 0.2
        }
        Caso 'Sombra con tamaño 120 % (Size)' {
            $s = Forma-Efecto $d 6 'sombra Size 120'
            $s.Shadow.Visible = $msoTrue; $s.Shadow.Style = 2; $s.Shadow.OffsetX = 0; $s.Shadow.OffsetY = 12; $s.Shadow.Blur = 10; $s.Shadow.Size = 120
            $s.Shadow.Transparency = 0.5
        }
        Caso 'Sombra con inclinación (Obscured/RotateWithShape)' {
            $s = Forma-Efecto $d 7 'sombra con RotateWithShape, rot 30'
            $s.Rotation = 30
            $s.Shadow.Visible = $msoTrue; $s.Shadow.OffsetX = 8; $s.Shadow.OffsetY = 8; $s.Shadow.Blur = 6; $s.Shadow.RotateWithShape = $msoTrue
        }
        $d = Nueva-Diap $pres 'Resplandor (Glow) y bordes suaves (SoftEdge)'
        $gl = @(@(5, 0.0, 255, 200, 0), @(10, 0.0, 0, 200, 255), @(20, 0.4, 255, 80, 80), @(40, 0.6, 80, 255, 80))
        for ($i = 0; $i -lt $gl.Count; $i++) {
            $p = $gl[$i]
            Caso ("Glow radio " + $p[0]) {
                $s = Forma-Efecto $d $i ("glow radio $($p[0]) transp $($p[1])")
                $s.Glow.Radius = $p[0]; $s.Glow.Transparency = $p[1]; $s.Glow.Color.RGB = (RGB $p[2] $p[3] $p[4])
            }
        }
        $se = @(2.5, 5, 10, 25, 50)
        for ($i = 0; $i -lt $se.Count; $i++) {
            Caso ("SoftEdge radio " + $se[$i]) {
                $s = Forma-Efecto $d (5 + $i) ("soft edge radio $($se[$i])")
                $s.SoftEdge.Radius = [single]$se[$i]
            }
        }
        Caso 'SoftEdge tipo preestablecido 1–6' {
            for ($k = 1; $k -le 3; $k++) {
                Caso ("SoftEdge.Type $k") {
                    $s = Forma-Efecto $d (10 + $k - 1) ("SoftEdge.Type $k")
                    $s.SoftEdge.Type = $k
                }
            }
        }
        Caso 'Glow + sombra combinados' {
            $s = Forma-Efecto $d 13 'glow 10 + sombra'
            $s.Glow.Radius = 10; $s.Glow.Color.RGB = (RGB 255 220 0)
            $s.Shadow.Visible = $msoTrue; $s.Shadow.OffsetX = 5; $s.Shadow.OffsetY = 5; $s.Shadow.Blur = 5
        }
        $d = Nueva-Diap $pres 'Reflejo: tipos 1–9'
        for ($k = 1; $k -le 9; $k++) {
            Caso ("Reflection.Type $k") {
                $s = Forma-Efecto $d ($k - 1) ("Reflection.Type $k")
                $s.Reflection.Type = $k
            }
        }
        Caso 'Reflejo personalizado (blur, offset, tamaño, transparencia)' {
            $s = Forma-Efecto $d 9 'reflejo personalizado'
            $s.Reflection.Type = 1
            $s.Reflection.Blur = 4; $s.Reflection.Offset = 8; $s.Reflection.Size = 60; $s.Reflection.Transparency = 0.5
        }
        $d = Nueva-Diap $pres 'Bisel 3D: tipos 2–13 (MsoBevelType)'
        for ($k = 2; $k -le 13; $k++) {
            Caso ("Bisel superior tipo $k") {
                $s = Forma-Efecto $d ($k - 2) ("BevelTopType $k") 6
                $s.ThreeD.BevelTopType = $k; $s.ThreeD.BevelTopDepth = 8; $s.ThreeD.BevelTopInset = 8
                $s.Fill.ForeColor.RGB = (RGB 200 120 60)
            }
        }
        Caso 'Bisel superior e inferior con material' {
            $s = Forma-Efecto $d 12 'bisel sup+inf, metal' 6
            $s.ThreeD.BevelTopType = 3; $s.ThreeD.BevelTopDepth = 10; $s.ThreeD.BevelTopInset = 10
            $s.ThreeD.BevelBottomType = 3; $s.ThreeD.BevelBottomDepth = 6; $s.ThreeD.BevelBottomInset = 6
            $s.ThreeD.PresetMaterial = 5
        }
        Caso 'Extrusión (profundidad) con rotación 3D' {
            $s = Forma-Efecto $d 13 'extrusión 30 pt + rot X 30' 6
            $s.ThreeD.Visible = $msoTrue; $s.ThreeD.Depth = 30; $s.ThreeD.RotationX = 30; $s.ThreeD.RotationY = 20
        }
        Caso 'Bisel con relieve y contorno' {
            $s = Forma-Efecto $d 14 'bisel + contorno de relieve' 6
            $s.ThreeD.BevelTopType = 12; $s.ThreeD.BevelTopDepth = 6; $s.ThreeD.BevelTopInset = 6
            $s.ThreeD.ContourWidth = 4; $s.ThreeD.ContourColor.RGB = (RGB 255 255 0)
        }
        # --- efectos sobre el texto
        $d = Nueva-Diap $pres 'Efectos en el texto (sin relleno de forma)'
        function Texto-Efecto($d, [int]$i, [string]$texto) {
            $s = $d.Shapes.AddTextbox(1, 20 + ($i % 2) * 470, 44 + [math]::Floor($i / 2) * 110, 450, 90)
            $s.TextFrame2.TextRange.Text = $texto
            $s.TextFrame2.TextRange.Font.Size = 36; $s.TextFrame2.TextRange.Font.Bold = $msoTrue
            $s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = (RGB 40 100 180)
            return $s
        }
        Caso 'Texto con sombra exterior' {
            $s = Texto-Efecto $d 0 'Sombra de texto'
            $f = $s.TextFrame2.TextRange.Font
            $f.Shadow.Visible = $msoTrue; $f.Shadow.OffsetX = 4; $f.Shadow.OffsetY = 4; $f.Shadow.Blur = 4; $f.Shadow.Transparency = 0.4; $f.Shadow.ForeColor.RGB = (RGB 0 0 0)
        }
        Caso 'Texto con resplandor' {
            $s = Texto-Efecto $d 1 'Resplandor de texto'
            $f = $s.TextFrame2.TextRange.Font
            $f.Glow.Radius = 10; $f.Glow.Color.RGB = (RGB 255 200 0); $f.Glow.Transparency = 0.3
        }
        Caso 'Texto con reflejo' {
            $s = Texto-Efecto $d 2 'Reflejo de texto'
            $s.TextFrame2.TextRange.Font.Reflection.Type = 3
        }
        Caso 'Texto con contorno y relleno degradado' {
            $s = Texto-Efecto $d 3 'Degradado de texto'
            $f = $s.TextFrame2.TextRange.Font
            $f.Fill.TwoColorGradient(1, 1); $f.Fill.ForeColor.RGB = (RGB 255 0 0); $f.Fill.BackColor.RGB = (RGB 0 0 255)
            $f.Line.Visible = $msoTrue; $f.Line.ForeColor.RGB = (RGB 0 0 0); $f.Line.Weight = 1
        }
        Caso 'Texto con bisel 3D' {
            $s = Texto-Efecto $d 4 'Bisel de texto'
            $s.TextFrame2.ThreeD.BevelTopType = 3; $s.TextFrame2.ThreeD.BevelTopDepth = 4; $s.TextFrame2.ThreeD.BevelTopInset = 4
        }
        Caso 'Texto con bordes suaves (SoftEdgeFormat)' {
            $s = Texto-Efecto $d 5 'Bordes suaves'
            $s.TextFrame2.TextRange.Font.SoftEdgeFormat = 3
        }
        Caso 'Texto con sombra interior' {
            $s = Texto-Efecto $d 6 'Sombra interior'
            $f = $s.TextFrame2.TextRange.Font
            $f.Shadow.Visible = $msoTrue; $f.Shadow.Style = 1; $f.Shadow.OffsetX = 3; $f.Shadow.OffsetY = 3; $f.Shadow.Blur = 4
        }
        Caso 'Texto con relleno de trama' {
            $s = Texto-Efecto $d 7 'Trama de texto'
            $f = $s.TextFrame2.TextRange.Font
            $f.Fill.Patterned(15); $f.Fill.ForeColor.RGB = (RGB 200 0 0); $f.Fill.BackColor.RGB = (RGB 255 255 255)
        }
        Caso 'Forma con texto: sombra en la forma y en el texto' {
            $s = Caja $d 5 20 340 450 100 'Forma con sombra' 28
            $s.Fill.ForeColor.RGB = (RGB 255 240 200)
            $s.Shadow.Visible = $msoTrue; $s.Shadow.OffsetX = 6; $s.Shadow.OffsetY = 6; $s.Shadow.Blur = 6
        }
        Caso 'Forma sin relleno con sombra (la sombra va al texto)' {
            $s = Caja $d 5 490 340 450 100 'Sin relleno, con sombra' 28
            $s.Fill.Visible = $msoFalse; $s.Line.Visible = $msoFalse
            $s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = (RGB 0 100 0)
            $s.Shadow.Visible = $msoTrue; $s.Shadow.OffsetX = 4; $s.Shadow.OffsetY = 4; $s.Shadow.Blur = 3
        }
    }

    # ============================================================ 08 tablas
    Nuevo-Archivo '08-tablas.pptx' {
        param($pres)
        function Llenar-Tabla($tbl, [int]$filas, [int]$cols, [double]$tam = 10) {
            for ($r = 1; $r -le $filas; $r++) {
                for ($c = 1; $c -le $cols; $c++) {
                    $tx = $tbl.Cell($r, $c).Shape.TextFrame.TextRange
                    if ($r -eq 1) { $tx.Text = "Col $c" } else { $tx.Text = "f$($r - 1)c$c" }
                    $tx.Font.Size = $tam
                }
                $tbl.Rows($r).Height = 20
            }
        }
        # --- los 74 estilos
        $d = $null; $enDiap = 0; $hoja = 0
        foreach ($linea in $EstilosTabla) {
            $partes = $linea.Split('|')
            $grupo = $partes[0]; $acento = $partes[1]; $guid = $partes[2]
            $res = [pscustomobject]@{ Grupo = $grupo; Acento = $acento; Guid = $guid; Aceptado = $false; IdLeido = ''; NombrePowerPoint = ''; Error = '' }
            if ($d -eq $null -or $enDiap -ge 12) { $hoja++; $d = Nueva-Diap $pres "Estilos de tabla (hoja $hoja)"; $enDiap = 0 }
            $x = 14 + ($enDiap % 4) * 233; $y = 30 + [math]::Floor($enDiap / 4) * 165
            try {
                $sf = $d.Shapes.AddTable(4, 4, $x, $y + 16, 225, 88)
                $sf.Name = "Tabla_" + ($guid -replace '[{}]', '')
                $tbl = $sf.Table
                Llenar-Tabla $tbl 4 4 9
                $tbl.ApplyStyle($guid, $msoTrue)
                $res.IdLeido = $tbl.Style.Id
                $res.NombrePowerPoint = $tbl.Style.Name
                $res.Aceptado = ($tbl.Style.Id -eq $guid)
                if (-not $res.Aceptado) { $res.Error = 'ApplyStyle no falló pero Style.Id no coincide' }
                $null = Etiqueta $d $x $y 225 ("$grupo - $acento")
            } catch {
                $res.Error = ($_.Exception.Message -replace '\s+', ' ').Trim()
            }
            [void]$script:ResultadosEstilos.Add($res)
            if (-not $res.Aceptado) { Registrar ("Estilo de tabla $grupo $acento $guid") ([pscustomobject]@{ Exception = [pscustomobject]@{ Message = $res.Error } }) }
            $enDiap++
        }
        # control negativo: un GUID inventado
        Caso 'Control: GUID inexistente' {
            $d = Nueva-Diap $pres 'Control: ApplyStyle con un GUID inexistente'
            $sf = $d.Shapes.AddTable(4, 4, 20, 50, 300, 90); $tbl = $sf.Table
            Llenar-Tabla $tbl 4 4 10
            $malo = '{00000000-1111-2222-3333-444444444444}'
            $r = [pscustomobject]@{ Grupo = 'CONTROL'; Acento = 'inexistente'; Guid = $malo; Aceptado = $false; IdLeido = ''; NombrePowerPoint = ''; Error = '' }
            try { $tbl.ApplyStyle($malo, $msoTrue); $r.IdLeido = $tbl.Style.Id; $r.NombrePowerPoint = $tbl.Style.Name; $r.Aceptado = ($tbl.Style.Id -eq $malo); $r.Error = 'sin excepción' }
            catch { $r.Error = ($_.Exception.Message -replace '\s+', ' ').Trim() }
            [void]$script:ResultadosEstilos.Add($r)
            Nota ("control GUID inexistente -> aceptado=" + $r.Aceptado + "; Style.Id leído='" + $r.IdLeido + "'; " + $r.Error)
        }
        # --- celdas combinadas
        $d = Nueva-Diap $pres 'Tablas: celdas combinadas (Merge)'
        Caso 'Merge horizontal, vertical y en bloque' {
            $sf = $d.Shapes.AddTable(5, 5, 20, 50, 560, 250); $tbl = $sf.Table
            Llenar-Tabla $tbl 5 5 12
            $tbl.Cell(1, 1).Merge($tbl.Cell(1, 3))
            $tbl.Cell(2, 1).Merge($tbl.Cell(4, 1))
            $tbl.Cell(3, 3).Merge($tbl.Cell(4, 4))
            $tbl.Cell(5, 2).Merge($tbl.Cell(5, 5))
            $tbl.Cell(1, 1).Shape.TextFrame.TextRange.Text = 'Combinada 1x3'
            $tbl.Cell(2, 1).Shape.TextFrame.TextRange.Text = 'Combinada 3x1 vertical'
            $tbl.Cell(3, 3).Shape.TextFrame.TextRange.Text = 'Bloque 2x2'
            $tbl.Cell(5, 2).Shape.TextFrame.TextRange.Text = 'Combinada 1x4'
        }
        Caso 'Combinada con Split posterior' {
            $sf = $d.Shapes.AddTable(3, 3, 600, 50, 340, 150); $tbl = $sf.Table
            Llenar-Tabla $tbl 3 3 12
            $tbl.Cell(1, 1).Merge($tbl.Cell(2, 2))
            $tbl.Cell(3, 1).Split(1, 2)
        }
        # --- bordes y rellenos por celda
        $d = Nueva-Diap $pres 'Tablas: bordes y relleno por celda'
        Caso 'Bordes distintos en cada celda' {
            $sf = $d.Shapes.AddTable(3, 3, 20, 50, 440, 210); $tbl = $sf.Table
            Llenar-Tabla $tbl 3 3 12
            for ($r = 1; $r -le 3; $r++) {
                for ($c = 1; $c -le 3; $c++) {
                    $cel = $tbl.Cell($r, $c)
                    $cel.Shape.Fill.ForeColor.RGB = (RGB (255 - $r * 30) (255 - $c * 30) 220)
                    foreach ($b in 1, 2, 3, 4) {
                        $bd = $cel.Borders($b)
                        $bd.Visible = $msoTrue
                        $bd.ForeColor.RGB = (RGB ($r * 80) ($c * 80) 100)
                        $bd.Weight = [single]($r + $c)
                        $bd.DashStyle = (1, 4, 3)[($r + $c) % 3]
                    }
                }
            }
        }
        Caso 'Bordes diagonales' {
            $sf = $d.Shapes.AddTable(2, 2, 480, 50, 300, 120); $tbl = $sf.Table
            Llenar-Tabla $tbl 2 2 12
            foreach ($b in 5, 6) { $bd = $tbl.Cell(1, 1).Borders($b); $bd.Visible = $msoTrue; $bd.ForeColor.RGB = (RGB 200 0 0); $bd.Weight = 2 }
        }
        Caso 'Relleno de celda: sólido, degradado, transparente' {
            $sf = $d.Shapes.AddTable(2, 3, 20, 290, 600, 100); $tbl = $sf.Table
            Llenar-Tabla $tbl 2 3 12
            $tbl.Cell(1, 1).Shape.Fill.Solid(); $tbl.Cell(1, 1).Shape.Fill.ForeColor.RGB = (RGB 255 200 0)
            $tbl.Cell(1, 2).Shape.Fill.TwoColorGradient(1, 1); $tbl.Cell(1, 2).Shape.Fill.ForeColor.RGB = (RGB 255 0 0); $tbl.Cell(1, 2).Shape.Fill.BackColor.RGB = (RGB 0 0 255)
            $tbl.Cell(1, 3).Shape.Fill.Visible = $msoFalse
            $tbl.Cell(2, 1).Shape.Fill.ForeColor.ObjectThemeColor = 6; $tbl.Cell(2, 1).Shape.Fill.ForeColor.Brightness = 0.6
            $tbl.Cell(2, 2).Shape.Fill.Solid(); $tbl.Cell(2, 2).Shape.Fill.ForeColor.RGB = (RGB 0 150 0); $tbl.Cell(2, 2).Shape.Fill.Transparency = 0.5
        }
        Caso 'Texto de celda: alineación, anclaje, márgenes, color' {
            $sf = $d.Shapes.AddTable(2, 3, 20, 410, 600, 100); $tbl = $sf.Table
            Llenar-Tabla $tbl 2 3 12
            $tbl.Rows(2).Height = 60
            $c = $tbl.Cell(2, 1).Shape.TextFrame; $c.TextRange.ParagraphFormat.Alignment = 3; $c.VerticalAnchor = 4
            $c = $tbl.Cell(2, 2).Shape.TextFrame; $c.TextRange.ParagraphFormat.Alignment = 2; $c.VerticalAnchor = 3; $c.TextRange.Font.Bold = $msoTrue
            $c = $tbl.Cell(2, 3).Shape.TextFrame; $c.MarginLeft = 20; $c.MarginTop = 15; $c.TextRange.Font.Color.RGB = (RGB 200 0 0)
        }
        # --- fila que crece
        $d = Nueva-Diap $pres 'Tablas: texto que obliga a crecer la fila'
        Caso 'Fila que crece por texto largo' {
            $sf = $d.Shapes.AddTable(3, 3, 20, 50, 440, 90); $tbl = $sf.Table
            Llenar-Tabla $tbl 3 3 12
            $tbl.Cell(2, 2).Shape.TextFrame.TextRange.Text = $ParrafoLargo
            $sf.Name = 'FilaQueCrece'
            Nota ("tabla 'FilaQueCrece': alto declarado tras el texto = " + $sf.Height)
        }
        Caso 'Fila con varios párrafos y fuente grande' {
            $sf = $d.Shapes.AddTable(2, 2, 480, 50, 440, 60); $tbl = $sf.Table
            Llenar-Tabla $tbl 2 2 10
            $tbl.Cell(2, 1).Shape.TextFrame.TextRange.Text = "Uno`rDos`rTres`rCuatro"
            $tbl.Cell(2, 1).Shape.TextFrame.TextRange.Font.Size = 24
        }
        Caso 'Columnas de ancho distinto' {
            $sf = $d.Shapes.AddTable(3, 4, 20, 330, 700, 90); $tbl = $sf.Table
            Llenar-Tabla $tbl 3 4 12
            $tbl.Columns(1).Width = 60; $tbl.Columns(2).Width = 300; $tbl.Columns(3).Width = 100; $tbl.Columns(4).Width = 240
        }
        # --- banderas del estilo
        $d = Nueva-Diap $pres 'Tablas: FirstRow / HorizBanding / FirstCol / LastRow / LastCol / VertBanding'
        $flags = @(
            @('Por defecto', @{ }), @('FirstRow apagado', @{ FirstRow = $msoFalse }), @('HorizBanding apagado', @{ HorizBanding = $msoFalse }),
            @('FirstCol prendido', @{ FirstCol = $msoTrue }), @('LastRow prendido', @{ LastRow = $msoTrue }), @('LastCol prendido', @{ LastCol = $msoTrue }),
            @('VertBanding prendido', @{ VertBanding = $msoTrue }), @('Todo apagado', @{ FirstRow = $msoFalse; HorizBanding = $msoFalse }),
            @('Todo prendido', @{ FirstRow = $msoTrue; HorizBanding = $msoTrue; FirstCol = $msoTrue; LastRow = $msoTrue; LastCol = $msoTrue; VertBanding = $msoTrue })
        )
        for ($i = 0; $i -lt $flags.Count; $i++) {
            $fl = $flags[$i]
            Caso ("Banderas: " + $fl[0]) {
                $x = 14 + ($i % 3) * 312; $y = 36 + [math]::Floor($i / 3) * 165
                $null = Etiqueta $d $x $y 300 $fl[0]
                $sf = $d.Shapes.AddTable(4, 4, $x, $y + 16, 300, 88); $tbl = $sf.Table
                Llenar-Tabla $tbl 4 4 10
                foreach ($k in $fl[1].Keys) { $tbl.$k = $fl[1][$k] }
            }
        }
        $d = Nueva-Diap $pres 'Tablas: banderas con estilos Medium 4 (Accent 2) y Dark 1'
        $guidDos = ($EstilosTabla | Where-Object { $_ -like 'Medium Style 4|Accent 2|*' }).Split('|')[2]
        $guidDark = ($EstilosTabla | Where-Object { $_ -like 'Dark Style 1|Accent 3|*' }).Split('|')[2]
        $k = 0
        foreach ($g in $guidDos, $guidDark) {
            foreach ($cfg in @(@{ FirstCol = $msoTrue }, @{ LastRow = $msoTrue; LastCol = $msoTrue }, @{ VertBanding = $msoTrue; HorizBanding = $msoFalse })) {
                Caso ("Banderas con estilo $g #$k") {
                    $x = 14 + ($k % 3) * 312; $y = 40 + [math]::Floor($k / 3) * 165
                    $sf = $d.Shapes.AddTable(4, 4, $x, $y, 300, 88); $tbl = $sf.Table
                    Llenar-Tabla $tbl 4 4 10
                    $tbl.ApplyStyle($g, $msoTrue)
                    foreach ($key in $cfg.Keys) { $tbl.$key = $cfg[$key] }
                }
                $k++
            }
        }
    }

    # ============================================================ 09 grupos y transformaciones
    Nuevo-Archivo '09-grupos-transformaciones.pptx' {
        param($pres)
        $d = Nueva-Diap $pres 'Formas rotadas y volteadas'
        $rot = @(0, 15, 45, 90, 135, 180, 270, 330)
        for ($i = 0; $i -lt $rot.Count; $i++) {
            Caso ("Rotación " + $rot[$i]) {
                $s = Caja $d 33 (60 + ($i % 4) * 225) (70 + [math]::Floor($i / 4) * 130) 130 60 ("rot " + $rot[$i]) 12
                $s.Rotation = $rot[$i]
            }
        }
        $fl = @(@('Sin voltear', $false, $false), @('Flip H', $true, $false), @('Flip V', $false, $true), @('Flip H+V', $true, $true))
        for ($i = 0; $i -lt 4; $i++) {
            Caso ("Volteo " + $fl[$i][0]) {
                $s = Caja $d 33 (60 + $i * 225) 340 130 70 ($fl[$i][0]) 12
                if ($fl[$i][1]) { $s.Flip(0) }
                if ($fl[$i][2]) { $s.Flip(1) }
            }
        }
        $d = Nueva-Diap $pres 'Rotación + volteo + texto'
        Caso 'Rotada 30 y volteada H con texto' {
            $s = Caja $d 33 60 80 200 100 'rot 30 + flipH' 14; $s.Rotation = 30; $s.Flip(0)
        }
        Caso 'Rotada 60 y volteada V con texto' {
            $s = Caja $d 33 350 80 200 100 'rot 60 + flipV' 14; $s.Rotation = 60; $s.Flip(1)
        }
        Caso 'Triángulo rectángulo volteado H+V rotado 20' {
            $s = Caja $d 8 650 80 200 100 'tri rect' 14; $s.Flip(0); $s.Flip(1); $s.Rotation = 20
        }
        Caso 'Línea volteada y rotada' {
            $l = $d.Shapes.AddLine(100, 300, 300, 400); $l.Flip(0); $l.Rotation = 25; $l.Line.Weight = 4; $l.Line.EndArrowheadStyle = 2
        }
        Caso 'Cuadro de texto con texto rotado y volteado' {
            $s = Texto $d 450 330 300 40 'Cuadro de texto rot 345' 20; $s.Rotation = 345
        }
        # --- grupos
        $d = Nueva-Diap $pres 'Grupos, anidados, escalados y rotados'
        Caso 'Grupo simple de 3 formas' {
            $a = Caja $d 1 40 60 80 50 'A' 12; $a.Name = 'g1a'
            $b = Caja $d 9 90 100 80 50 'B' 12; $b.Name = 'g1b'
            $c = Caja $d 5 60 140 80 50 'C' 12; $c.Name = 'g1c'
            $g = $d.Shapes.Range([object[]]@('g1a', 'g1b', 'g1c')).Group(); $g.Name = 'Grupo1'
        }
        Caso 'Grupo dentro de grupo (2 niveles)' {
            $a = Caja $d 1 260 60 70 40 'a' 12; $a.Name = 'g2a'
            $b = Caja $d 1 340 60 70 40 'b' 12; $b.Name = 'g2b'
            $g = $d.Shapes.Range([object[]]@('g2a', 'g2b')).Group(); $g.Name = 'Interno'
            $c = Caja $d 9 260 120 150 50 'c' 12; $c.Name = 'g2c'
            $g2 = $d.Shapes.Range([object[]]@('Interno', 'g2c')).Group(); $g2.Name = 'Externo'
        }
        Caso 'Grupo con tres niveles' {
            $a = Caja $d 1 460 60 50 30 '1' 10; $a.Name = 'g3a'
            $b = Caja $d 1 520 60 50 30 '2' 10; $b.Name = 'g3b'
            $g = $d.Shapes.Range([object[]]@('g3a', 'g3b')).Group(); $g.Name = 'N3'
            $c = Caja $d 1 460 100 110 30 '3' 10; $c.Name = 'g3c'
            $g2 = $d.Shapes.Range([object[]]@('N3', 'g3c')).Group(); $g2.Name = 'N2'
            $e = Caja $d 9 460 140 110 30 '4' 10; $e.Name = 'g3e'
            $g3 = $d.Shapes.Range([object[]]@('N2', 'g3e')).Group(); $g3.Name = 'N1'
        }
        Caso 'Grupo escalado de forma no uniforme con hijos rotados' {
            $a = Caja $d 1 40 260 60 60 'rot' 12; $a.Rotation = 30; $a.Name = 'g4a'
            $b = Caja $d 9 120 270 60 40 'rot' 12; $b.Rotation = 60; $b.Name = 'g4b'
            $c = Caja $d 33 60 330 100 40 'rot' 12; $c.Rotation = 15; $c.Name = 'g4c'
            $g = $d.Shapes.Range([object[]]@('g4a', 'g4b', 'g4c')).Group(); $g.Name = 'GrupoEscalado'
            $g.LockAspectRatio = $msoFalse
            $g.Width = $g.Width * 2.2; $g.Height = $g.Height * 0.7
        }
        Caso 'Grupo rotado 30 grados' {
            $a = Caja $d 1 500 280 80 50 'x' 12; $a.Name = 'g5a'
            $b = Caja $d 9 600 300 80 50 'y' 12; $b.Name = 'g5b'
            $c = Caja $d 5 540 350 90 40 'z' 12; $c.Rotation = 20; $c.Name = 'g5c'
            $g = $d.Shapes.Range([object[]]@('g5a', 'g5b', 'g5c')).Group(); $g.Name = 'GrupoRotado'
            $g.Rotation = 30
        }
        Caso 'Grupo volteado H con texto' {
            $a = Caja $d 33 720 280 100 50 'flip' 12; $a.Name = 'g6a'
            $b = Caja $d 9 740 340 60 50 'g' 12; $b.Name = 'g6b'
            $g = $d.Shapes.Range([object[]]@('g6a', 'g6b')).Group(); $g.Name = 'GrupoVolteado'
            $g.Flip(0)
        }
        $d = Nueva-Diap $pres 'Grupo anidado, escalado y rotado a la vez'
        Caso 'Grupo anidado escalado y rotado' {
            $a = Caja $d 1 100 100 60 40 'a' 12; $a.Rotation = 20; $a.Name = 'h1a'
            $b = Caja $d 1 180 100 60 40 'b' 12; $b.Name = 'h1b'
            $g = $d.Shapes.Range([object[]]@('h1a', 'h1b')).Group(); $g.Name = 'Hijo'
            $c = Caja $d 9 100 160 140 60 'c' 12; $c.Name = 'h1c'
            $g2 = $d.Shapes.Range([object[]]@('Hijo', 'h1c')).Group(); $g2.Name = 'Padre'
            $g2.LockAspectRatio = $msoFalse
            $g2.Width = $g2.Width * 1.8; $g2.Height = $g2.Height * 1.3
            $g2.Rotation = 340
        }
        Caso 'Grupo con línea, conector y cuadro de texto' {
            $l = $d.Shapes.AddLine(500, 100, 650, 180); $l.Name = 'h2l'; $l.Line.Weight = 3
            $t = Texto $d 520 190 150 30 'texto' 14; $t.Name = 'h2t'
            $e = Caja $d 9 680 100 100 60 'o' 12; $e.Name = 'h2e'
            $g = $d.Shapes.Range([object[]]@('h2l', 'h2t', 'h2e')).Group(); $g.Name = 'Mixto'
            $g.Rotation = 10
        }
    }

    # ============================================================ 10 imágenes
    Nuevo-Archivo '10-imagenes.pptx' {
        param($pres)
        function Imagen($d, [string]$ruta, [double]$x, [double]$y, [double]$w, [double]$h, [string]$etq) {
            $p = $d.Shapes.AddPicture($ruta, $msoFalse, $msoTrue, $x, $y, $w, $h)
            $null = Etiqueta $d $x ($y + $h + 2) 200 $etq
            return $p
        }
        $d = Nueva-Diap $pres 'Imágenes: PNG con transparencia, JPG, tamaños'
        $fondo = Caja $d 1 14 36 300 160 ''; $fondo.Fill.Patterned(21); $fondo.Fill.ForeColor.RGB = (RGB 200 200 200); $fondo.Fill.BackColor.RGB = (RGB 255 255 255)
        Caso 'PNG con transparencia sobre trama' { $null = Imagen $d $PNG 40 44 140 140 'PNG con transparencia' }
        Caso 'JPG' { $null = Imagen $d $JPG 340 44 240 150 'JPG 320x200' }
        Caso 'PNG a tamaño original (sin w/h)' {
            $p = $d.Shapes.AddPicture($PNG, $msoFalse, $msoTrue, 620, 44)
            $null = Etiqueta $d 620 ($p.Top + $p.Height + 2) 200 ("PNG tamaño original: $([math]::Round($p.Width))x$([math]::Round($p.Height))")
        }
        Caso 'JPG deformado (proporción distinta)' { $null = Imagen $d $JPG 40 260 300 60 'JPG deformado 300x60' }
        Caso 'PNG vinculado vs. incrustado (LinkToFile=0, Save=-1)' { $null = Imagen $d $PNG 400 230 120 120 'PNG incrustado' }
        # --- recorte
        $d = Nueva-Diap $pres 'Imágenes: recorte (CropLeft/Top/Right/Bottom)'
        $rec = @(@('Sin recorte', 0, 0, 0, 0), @('Recorte izq 40', 40, 0, 0, 0), @('Recorte arriba 40 y abajo 40', 0, 40, 0, 40), @('Recorte 30 en los 4 lados', 30, 30, 30, 30), @('Recorte asimétrico 20/60/10/5', 20, 60, 10, 5))
        for ($i = 0; $i -lt $rec.Count; $i++) {
            $r = $rec[$i]
            Caso ("Recorte " + $r[0]) {
                $p = Imagen $d $JPG (14 + $i * 188) 50 178 112 $r[0]
                $p.LockAspectRatio = $msoFalse
                $p.PictureFormat.CropLeft = $r[1]; $p.PictureFormat.CropTop = $r[2]; $p.PictureFormat.CropRight = $r[3]; $p.PictureFormat.CropBottom = $r[4]
            }
        }
        Caso 'Recorte con PictureFormat.Crop (desplazar/escalar la imagen dentro del marco)' {
            $p = Imagen $d $JPG 14 230 240 150 'Crop.PictureWidth 400, offset'
            $p.PictureFormat.Crop.PictureWidth = 400; $p.PictureFormat.Crop.PictureHeight = 250
            $p.PictureFormat.Crop.PictureOffsetX = -30; $p.PictureFormat.Crop.PictureOffsetY = -20
        }
        Caso 'Recorte con forma + recorte rectangular' {
            $p = Imagen $d $PNG 300 230 150 150 'Recorte 20 + elipse'
            $p.PictureFormat.CropLeft = 20; $p.PictureFormat.CropRight = 20; $p.AutoShapeType = 9
        }
        # --- recorte con forma
        $d = Nueva-Diap $pres 'Imágenes: recortadas a una forma (AutoShapeType)'
        $fm = @(@(9, 'Elipse'), @(5, 'Rect. redondeado'), @(92, 'Estrella 5'), @(21, 'Corazón'), @(7, 'Rombo'), @(52, 'Chevron'), @(10, 'Hexágono?'), @(106, 'Llamada redondeada'))
        for ($i = 0; $i -lt $fm.Count; $i++) {
            Caso ("Imagen con forma " + $fm[$i][1]) {
                $p = Imagen $d $JPG (14 + ($i % 4) * 233) (40 + [math]::Floor($i / 4) * 190) 200 125 $fm[$i][1]
                $p.AutoShapeType = $fm[$i][0]
            }
        }
        # --- rotación, volteo
        $d = Nueva-Diap $pres 'Imágenes: rotadas, volteadas y con borde'
        Caso 'Rotada 20' { $p = Imagen $d $JPG 40 70 200 125 'rotada 20'; $p.Rotation = 20 }
        Caso 'Volteada H' { $p = Imagen $d $JPG 300 70 200 125 'flip H'; $p.Flip(0) }
        Caso 'Volteada V' { $p = Imagen $d $JPG 560 70 200 125 'flip V'; $p.Flip(1) }
        Caso 'Rotada 90 y volteada H' { $p = Imagen $d $JPG 780 60 130 85 'rot 90 + flip H'; $p.Rotation = 90; $p.Flip(0) }
        Caso 'Con borde de 4 pt' {
            $p = Imagen $d $JPG 40 300 200 125 'borde 4 pt rojo'
            $p.Line.Visible = $msoTrue; $p.Line.Weight = 4; $p.Line.ForeColor.RGB = (RGB 200 0 0)
        }
        Caso 'Con borde discontinuo y recorte elíptico' {
            $p = Imagen $d $JPG 300 300 200 125 'borde guiones + elipse'
            $p.AutoShapeType = 9; $p.Line.Visible = $msoTrue; $p.Line.Weight = 3; $p.Line.DashStyle = 4; $p.Line.ForeColor.RGB = (RGB 0 0 200)
        }
        Caso 'Con sombra' {
            $p = Imagen $d $JPG 560 300 200 125 'sombra'
            $p.Shadow.Visible = $msoTrue; $p.Shadow.OffsetX = 6; $p.Shadow.OffsetY = 6; $p.Shadow.Blur = 6
        }
        # --- recolor, transparencia
        $d = Nueva-Diap $pres 'Imágenes: color, brillo/contraste y transparencia'
        $ct = @(@('Automático (1)', 1), @('Escala de grises (2)', 2), @('Blanco y negro (3)', 3), @('Marca de agua (4)', 4))
        for ($i = 0; $i -lt $ct.Count; $i++) {
            Caso ("ColorType " + $ct[$i][0]) {
                $p = Imagen $d $JPG (14 + $i * 233) 40 220 138 $ct[$i][0]
                $p.PictureFormat.ColorType = $ct[$i][1]
            }
        }
        Caso 'Brillo 0.75' { $p = Imagen $d $JPG 14 220 220 138 'Brightness 0.75'; $p.PictureFormat.Brightness = 0.75 }
        Caso 'Contraste 0.9' { $p = Imagen $d $JPG 247 220 220 138 'Contrast 0.9'; $p.PictureFormat.Contrast = 0.9 }
        Caso 'Brillo 0.3 y contraste 0.2' { $p = Imagen $d $JPG 480 220 220 138 'Brillo 0.3 / Contraste 0.2'; $p.PictureFormat.Brightness = 0.3; $p.PictureFormat.Contrast = 0.2 }
        Caso 'Color transparente (TransparentBackground, color blanco)' {
            $f = Caja $d 1 713 220 220 138 ''; $f.Fill.ForeColor.RGB = (RGB 255 220 120)
            $p = Imagen $d $PNG 723 225 120 120 'color transparente'
            $p.PictureFormat.TransparentBackground = $msoTrue; $p.PictureFormat.TransparencyColor = (RGB 255 255 255)
        }
        Nota "PictureFormat.Transparency no existe en el modelo COM de PowerPoint: la transparencia de imagen solo se cubre con TransparencyColor y con el alfa del PNG"
        Caso 'Transparencia vía relleno de la forma de imagen (Fill.Transparency)' {
            $f = Caja $d 1 247 400 220 110 ''; $f.Fill.ForeColor.RGB = (RGB 40 40 40)
            $p = Imagen $d $PNG 247 400 110 110 'Fill.Transparency 0.5'
            $p.Fill.Transparency = 0.5
        }
        Caso 'Texto alternativo y título' {
            $p = Imagen $d $PNG 500 400 100 100 'AlternativeText'
            $p.AlternativeText = 'Círculo de prueba con degradado'; $p.Title = 'Imagen de prueba'
        }
    }

    # ============================================================ 11 SmartArt (necesita ventana)
    Nuevo-Archivo '11-smartart.pptx' {
        param($pres)
        $mapa = @{ }
        $n = $ppt.SmartArtLayouts.Count
        if ($n -eq 0) { throw 'Application.SmartArtLayouts está vacío (PowerPoint no lo cargó)' }
        for ($i = 1; $i -le $n; $i++) {
            $l = $ppt.SmartArtLayouts.Item($i)
            $mapa[($l.Id -replace '^.*/layout/', '')] = $i
        }
        Nota ("SmartArtLayouts disponibles: $n")
        $diseños = @(
            @('default', 'Lista de bloques básica'), @('vList2', 'Lista vertical de viñetas'), @('hList1', 'Lista horizontal de viñetas'), @('lProcess2', 'Lista agrupada'),
            @('process1', 'Proceso básico'), @('chevron1', 'Proceso cheurón'), @('vProcess5', 'Proceso escalonado'), @('cycle2', 'Ciclo básico'),
            @('cycle1', 'Ciclo de texto'), @('radial1', 'Radial básico'), @('hierarchy1', 'Jerarquía'), @('orgChart1', 'Organigrama'),
            @('pyramid1', 'Pirámide básica'), @('matrix1', 'Matriz en mosaico'), @('venn1', 'Venn básico'), @('target1', 'Objetivo básico')
        )
        $d = $null
        for ($i = 0; $i -lt $diseños.Count; $i++) {
            $dz = $diseños[$i]
            if (($i % 4) -eq 0) { $d = Nueva-Diap $pres ("SmartArt " + ($i + 1) + '–' + ($i + 4)) }
            Caso ("SmartArt " + $dz[0]) {
                if (-not $mapa.ContainsKey($dz[0])) { throw "no existe el diseño $($dz[0])" }
                $x = 14 + ($i % 2) * 468; $y = 36 + [math]::Floor(($i % 4) / 2) * 252
                $sh = $d.Shapes.AddSmartArt($ppt.SmartArtLayouts.Item($mapa[$dz[0]]), $x, $y, 456, 236)
                $sh.Name = "SmartArt_$($dz[0])"
                $sa = $sh.SmartArt
                $total = $sa.AllNodes.Count
                for ($k = 1; $k -le $total; $k++) {
                    $sa.AllNodes.Item($k).TextFrame2.TextRange.Text = "Nodo $k de $($dz[1])"
                }
                $null = Etiqueta $d $x ($y + 237) 456 ($dz[0] + " - nodos iniciales: $total")
            }
        }
        Caso 'SmartArt de jerarquía con nodos hijos añadidos' {
            $d = Nueva-Diap $pres 'SmartArt: jerarquía con nodos añadidos y colores'
            $sh = $d.Shapes.AddSmartArt($ppt.SmartArtLayouts.Item($mapa['hierarchy1']), 14, 40, 920, 400)
            $sh.Name = 'Jerarquia_extendida'
            $sa = $sh.SmartArt
            $sa.AllNodes.Item(1).TextFrame2.TextRange.Text = 'Dirección'
            $h1 = $sa.AllNodes.Item(1).AddNode(5)   # debajo
            $h1.TextFrame2.TextRange.Text = 'Área A'
            $h2 = $sa.AllNodes.Item(1).AddNode(5)
            $h2.TextFrame2.TextRange.Text = 'Área B'
            $h3 = $h1.AddNode(5); $h3.TextFrame2.TextRange.Text = 'Equipo A1'
            Nota ("jerarquía extendida: nodos finales = " + $sa.AllNodes.Count)
        }
        Caso 'SmartArt con estilo de color y de diseño distintos' {
            $d = Nueva-Diap $pres 'SmartArt: Color y QuickStyle'
            $sh = $d.Shapes.AddSmartArt($ppt.SmartArtLayouts.Item($mapa['process1']), 14, 40, 920, 250)
            $sh.Name = 'Proceso_color'
            $sa = $sh.SmartArt
            for ($k = 1; $k -le $sa.AllNodes.Count; $k++) { $sa.AllNodes.Item($k).TextFrame2.TextRange.Text = "Paso $k" }
            try { $sa.Color = $ppt.SmartArtColors.Item(5) } catch { Registrar 'SmartArt Color' $_ }
            try { $sa.QuickStyle = $ppt.SmartArtQuickStyles.Item(4) } catch { Registrar 'SmartArt QuickStyle' $_ }
        }
    } 960 540 $true

    # ============================================================ 12 gráficos
    Nuevo-Archivo '12-graficos.pptx' {
        param($pres)
        $tipos = @(
            @('Columnas agrupadas', 51), @('Columnas apiladas', 52), @('Barras agrupadas', 57), @('Líneas con marcadores', 65),
            @('Circular', 5), @('Anillo', -4120), @('Áreas', 1), @('Dispersión', -4169),
            @('Columnas 100 % apiladas', 53), @('Barras apiladas', 58), @('Líneas', 4), @('Áreas apiladas', 76)
        )
        $series = [ordered]@{ Norte = @(4.3, 2.5, 3.5, 4.5); Centro = @(2.4, 4.4, 1.8, 2.8); Sur = @(2, 2, 3, 5) }
        $d = $null
        for ($i = 0; $i -lt $tipos.Count; $i++) {
            $t = $tipos[$i]
            if (($i % 4) -eq 0) { $d = Nueva-Diap $pres ("Gráficos " + ($i + 1) + '–' + ($i + 4)) }
            Caso ("Gráfico " + $t[0]) {
                $x = 14 + ($i % 2) * 468; $y = 36 + [math]::Floor(($i % 4) / 2) * 252
                $cats = if ($t[1] -eq -4169) { @(1, 2, 3, 4) } else { @('Trim 1', 'Trim 2', 'Trim 3', 'Trim 4') }
                $sh = Grafico $d $t[1] $x $y 456 244 "Grafico_$i" $cats $series $t[0]
                $ch = $sh.Chart
                $ch.HasLegend = $true; $ch.Legend.Position = -4107
                $ch.SeriesCollection(1).HasDataLabels = $true
                Nota ("gráfico '" + $t[0] + "': series = " + $ch.SeriesCollection().Count + ", puntos = " + $ch.SeriesCollection(1).Points().Count) | Out-Null
            }
        }
    } 960 540 $true

    # ============================================================ 15 gráficos: variantes
    # Lo que cambia el diseño automático y la escala: posición de la leyenda,
    # sin título, negativos, miles, categorías largas, eje secundario,
    # circular con porcentajes, dispersión con líneas, tamaños.
    Nuevo-Archivo '15-graficos-variantes.pptx' {
        param($pres)
        $trim = @('Trim 1', 'Trim 2', 'Trim 3', 'Trim 4')
        $tres = [ordered]@{ Norte = @(4.3, 2.5, 3.5, 4.5); Centro = @(2.4, 4.4, 1.8, 2.8); Sur = @(2, 2, 3, 5) }
        $casos = @(
            @{ n = 'Leyenda a la derecha'; t = 51; f = { param($ch) $ch.HasLegend = $true; $ch.Legend.Position = -4152 } },
            @{ n = 'Leyenda arriba'; t = 4; f = { param($ch) $ch.HasLegend = $true; $ch.Legend.Position = -4160 } },
            @{ n = 'Leyenda a la izquierda'; t = 57; f = { param($ch) $ch.HasLegend = $true; $ch.Legend.Position = -4131 } },
            @{ n = ''; t = 51; f = { param($ch) $ch.HasLegend = $false } },
            @{ n = 'Negativos'; t = 51; s = [ordered]@{ Saldo = @(-3.2, 1.5, -0.8, 4.6); Meta = @(1, 2, -2, 3) } },
            @{ n = 'Miles'; t = 51; s = [ordered]@{ Ventas = @(12500, 31800, 27400, 48900) }; f = { param($ch) $ch.Axes(2).TickLabels.NumberFormat = '#,##0' } },
            @{ n = 'Categorías largas'; t = 51; c = @('Enero de 2025', 'Febrero de 2025', 'Marzo de 2025', 'Abril de 2025', 'Mayo de 2025', 'Junio de 2025', 'Julio de 2025', 'Agosto de 2025'); s = [ordered]@{ Visitas = @(120, 340, 290, 410, 380, 520, 610, 580) } },
            @{ n = 'Eje secundario'; t = 51; s = [ordered]@{ Unidades = @(120, 340, 290, 410); Margen = @(0.12, 0.31, 0.22, 0.4) }; f = { param($ch) $s2 = $ch.SeriesCollection(2); $s2.ChartType = 65; $s2.AxisGroup = 2; $ch.HasLegend = $true; $ch.Legend.Position = -4107 } },
            @{ n = 'Circular con porcentajes'; t = 5; s = [ordered]@{ Partes = @(45, 25, 20, 10) }; f = { param($ch) $se = $ch.SeriesCollection(1); $se.HasDataLabels = $true; $dl = $se.DataLabels(); $dl.ShowPercentage = $true; $dl.ShowValue = $false; $dl.ShowCategoryName = $true; $se.Points(1).Explosion = 12; $ch.HasLegend = $false } },
            @{ n = 'Anillo'; t = -4120; s = [ordered]@{ Partes = @(45, 25, 20, 10) }; f = { param($ch) $ch.HasLegend = $true; $ch.Legend.Position = -4152 } },
            @{ n = 'Dispersión con líneas'; t = 72; c = @(0.5, 1.7, 2.2, 3.9); s = [ordered]@{ A = @(1.2, 3.4, 2.1, 4.8); B = @(0.4, 1.1, 3.3, 2.9) } },
            @{ n = 'Decimales chicos'; t = 65; s = [ordered]@{ Tasa = @(0.12, 0.27, 0.19, 0.47) } }
        )
        $d = $null
        for ($i = 0; $i -lt $casos.Count; $i++) {
            $k = $casos[$i]
            if (($i % 4) -eq 0) { $d = Nueva-Diap $pres ("Variantes " + ($i + 1) + '–' + ($i + 4)) }
            Caso ("Gráfico " + $k.n) {
                $x = 14 + ($i % 2) * 468; $y = 36 + [math]::Floor(($i % 4) / 2) * 252
                $cats = if ($k.c) { $k.c } else { $trim }
                $ser = if ($k.s) { $k.s } else { $tres }
                $sh = Grafico $d $k.t $x $y 456 244 "Variante_$i" $cats $ser $k.n
                if ($k.f) { & $k.f $sh.Chart }
            }
        }
        $d = Nueva-Diap $pres 'Tamaños'
        Caso 'Gráfico chico' { $sh = Grafico $d 52 14 36 240 160 'Chico' $trim $tres 'Chico'; $sh.Chart.HasLegend = $true; $sh.Chart.Legend.Position = -4107 }
        Caso 'Gráfico grande' { $sh = Grafico $d 51 270 36 676 470 'Grande' $trim $tres 'Grande'; $sh.Chart.HasLegend = $true; $sh.Chart.Legend.Position = -4107 }
    } 960 540 $true

    # ============================================================ 13 patrón y diseño
    Nuevo-Archivo '13-patron-y-diseno.pptx' {
        param($pres)
        $m = $pres.SlideMaster
        Caso 'Fondo de patrón y formas propias en el patrón' {
            $m.Background.Fill.Patterned(21)
            $m.Background.Fill.ForeColor.RGB = (RGB 225 235 245); $m.Background.Fill.BackColor.RGB = (RGB 255 255 255)
            $b = $m.Shapes.AddShape(1, 0, 500, 960, 10); $b.Name = 'FranjaPatron'; $b.Fill.ForeColor.RGB = (RGB 20 60 120); $b.Line.Visible = $msoFalse
            $o = $m.Shapes.AddShape(9, 900, 6, 50, 50); $o.Name = 'LogoPatron'; $o.Fill.ForeColor.RGB = (RGB 230 120 20); $o.Line.Visible = $msoFalse
            $o.TextFrame.TextRange.Text = 'MC'; $o.TextFrame.TextRange.Font.Size = 14
        }
        Caso 'Viñetas del patrón: nivel 1 Wingdings §, nivel 2 numerado' {
            $niv = $m.TextStyles.Item(3).Levels   # ppBodyStyle = 3 (1 = predeterminado, 2 = título)
            $b1 = $niv.Item(1).ParagraphFormat.Bullet
            $b1.Visible = $msoTrue; $b1.Type = 1; $b1.UseTextFont = $msoFalse; $b1.Font.Name = 'Wingdings'; $b1.Character = 167
            $b2 = $niv.Item(2).ParagraphFormat.Bullet
            $b2.Visible = $msoTrue; $b2.Type = 2; $b2.Style = 3; $b2.StartValue = 1
        }
        Caso 'Diseño 4 (Dos objetos): forma propia y fondo propio' {
            $lay = $m.CustomLayouts.Item(4)
            $lay.FollowMasterBackground = $msoFalse
            $lay.Background.Fill.ForeColor.RGB = (RGB 255 245 220)
            $s = $lay.Shapes.AddShape(5, 330, 480, 300, 20); $s.Name = 'PildoraDiseno'; $s.Fill.ForeColor.RGB = (RGB 160 40 40); $s.Line.Visible = $msoFalse
        }
        Caso 'Diseño 5 (Comparación): degradado de fondo' {
            $lay = $m.CustomLayouts.Item(5)
            $lay.FollowMasterBackground = $msoFalse
            $lay.Background.Fill.TwoColorGradient(1, 1)
            $lay.Background.Fill.ForeColor.RGB = (RGB 240 255 240); $lay.Background.Fill.BackColor.RGB = (RGB 190 230 190)
        }
        # Llena los marcadores de una diapositiva según su tipo.
        function Llenar-Marcadores($d, [string]$prefijo) {
            $n = 0
            foreach ($sh in $d.Shapes) {
                if ($sh.Type -ne 14) { continue }
                $pt = $sh.PlaceholderFormat.Type
                if ($pt -eq 1 -or $pt -eq 3) { $sh.TextFrame.TextRange.Text = "$prefijo - título" }
                elseif ($pt -eq 4) { $sh.TextFrame.TextRange.Text = "$prefijo - subtítulo" }
                elseif ($pt -eq 2 -or $pt -eq 7) {
                    $n++
                    $sh.TextFrame.TextRange.Text = "Contenido $n primer nivel`rSegundo nivel de $n`rTercer nivel de $n`rOtro primer nivel $n"
                    $sh.TextFrame.TextRange.Paragraphs(2).IndentLevel = 2
                    $sh.TextFrame.TextRange.Paragraphs(3).IndentLevel = 3
                }
            }
        }
        $diseños = @(@(1, 'Título'), @(2, 'Título y contenido'), @(3, 'Encabezado de sección'), @(4, 'Dos contenidos'), @(5, 'Comparación'), @(6, 'Solo título'), @(7, 'En blanco'))
        foreach ($dz in $diseños) {
            Caso ("Diapositiva con diseño " + $dz[1]) {
                $d = Nueva-DiapDiseno $pres ("Diseño " + $dz[0] + ': ' + $dz[1]) $dz[0]
                Llenar-Marcadores $d $dz[1]
                $hf = $d.HeadersFooters
                $hf.Footer.Visible = $msoTrue; $hf.Footer.Text = 'Pie de página de prueba'
                $hf.SlideNumber.Visible = $msoTrue
                $hf.DateAndTime.Visible = $msoTrue; $hf.DateAndTime.UseFormat = $msoFalse; $hf.DateAndTime.Text = '9 de octubre de 2026'
            }
        }
        Caso 'Diapositiva con pie apagado y fondo propio' {
            $d = Nueva-DiapDiseno $pres 'Diseño 2 con fondo propio y sin pie' 2
            Llenar-Marcadores $d 'Fondo propio'
            $d.FollowMasterBackground = $msoFalse
            $d.Background.Fill.ForeColor.RGB = (RGB 255 230 230)
            $d.DisplayMasterShapes = $msoFalse
            $d.HeadersFooters.Footer.Visible = $msoFalse; $d.HeadersFooters.SlideNumber.Visible = $msoFalse; $d.HeadersFooters.DateAndTime.Visible = $msoFalse
        }
        Caso 'Diapositiva con fondo de patrón propio y formas del patrón ocultas' {
            $d = Nueva-Diap $pres 'En blanco con fondo de trama propia'
            $d.FollowMasterBackground = $msoFalse
            $d.Background.Fill.Patterned(10); $d.Background.Fill.ForeColor.RGB = (RGB 180 200 220); $d.Background.Fill.BackColor.RGB = (RGB 255 255 255)
            $s = Caja $d 1 100 100 300 100 'Formas del patrón visibles o no' 16
        }
    }

    # ============================================================ 14 tamaños
    function Cuerpo-Tamano($pres, [string]$nombre) {
        $w = $pres.PageSetup.SlideWidth; $h = $pres.PageSetup.SlideHeight
        $d = Nueva-Diap $pres "Tamaño: $nombre ($w x $h pt)"
        $r = Caja $d 1 1 1 ($w - 2) ($h - 2) ''
        $r.Fill.Visible = $msoFalse; $r.Line.Weight = 3; $r.Line.ForeColor.RGB = (RGB 200 0 0)
        $l1 = $d.Shapes.AddLine(0, 0, $w, $h); $l1.Line.ForeColor.RGB = (RGB 0 0 200)
        $l2 = $d.Shapes.AddLine($w, 0, 0, $h); $l2.Line.ForeColor.RGB = (RGB 0 0 200)
        $c = Caja $d 9 ($w / 2 - 60) ($h / 2 - 40) 120 80 "$nombre`r$w x $h" 12
        $t = Texto $d 20 ($h - 40) ($w - 40) 30 "Texto al pie de la diapositiva $nombre" 14
        $d2 = Nueva-DiapDiseno $pres "$nombre - diseño Título y contenido" 2
        $d2.Shapes.Placeholders.Item(1).TextFrame.TextRange.Text = "Título en $nombre"
        $d2.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "Primer punto`rSegundo punto`rTercer punto"
        $d2.HeadersFooters.Footer.Visible = $msoTrue; $d2.HeadersFooters.Footer.Text = 'Pie'
        $d2.HeadersFooters.SlideNumber.Visible = $msoTrue
    }
    Nuevo-Archivo '14a-tamano-4x3.pptx' {
        param($pres)
        $pres.PageSetup.SlideSize = 1    # ppSlideSizeOnScreen (4:3, 720x540)
        Cuerpo-Tamano $pres '4:3'
    } 0 0
    Nuevo-Archivo '14b-tamano-a4.pptx' {
        param($pres)
        $pres.PageSetup.SlideSize = 3    # ppSlideSizeA4Paper
        Cuerpo-Tamano $pres 'A4'
    } 0 0
    Nuevo-Archivo '14c-tamano-personalizado.pptx' {
        param($pres)
        $pres.PageSetup.SlideWidth = 500
        $pres.PageSetup.SlideHeight = 700
        Cuerpo-Tamano $pres 'Personalizado vertical 500x700'
    } 0 0

} catch {
    Registrar 'script' $_
} finally {
    if ($ppt -ne $null) {
        try { $ppt.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) } catch { }
    }
}

# ================================================================ verificación: reabrir y contar
Write-Host ''
Write-Host '== Verificación (reabrir cada archivo)' -ForegroundColor Cyan
$ppt2 = $null
$verif = New-Object System.Collections.ArrayList
try {
    $ppt2 = New-Object -ComObject PowerPoint.Application
    foreach ($f in (Get-ChildItem $Salida -Filter '*.pptx' | Sort-Object Name)) {
        $o = [pscustomobject]@{ Archivo = $f.Name; KB = [math]::Round($f.Length / 1KB, 1); Diapositivas = 0; Formas = 0; Abre = $false }
        try {
            $p = $ppt2.Presentations.Open($f.FullName, $msoTrue, $msoFalse, $msoFalse)   # solo lectura, sin título, sin ventana
            $o.Diapositivas = $p.Slides.Count
            foreach ($sl in $p.Slides) { $o.Formas += $sl.Shapes.Count }
            $o.Abre = $true
            $p.Close()
        } catch { $o.Abre = $false; Registrar ("Reabrir " + $f.Name) $_ }
        [void]$verif.Add($o)
    }
} catch { Registrar 'verificación' $_ }
finally {
    if ($ppt2 -ne $null) { try { $ppt2.Quit() } catch { }; try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt2) } catch { } }
}

# ================================================================ cierre de procesos propios
# PowerPoint tarda unos segundos en terminar tras Quit(): esperar hasta 20 s antes de matar nada.
for ($espera = 0; $espera -lt 20; $espera++) {
    if (@(Get-Process -Name POWERPNT, EXCEL -ErrorAction SilentlyContinue | Where-Object { $procPrevios -notcontains $_.Id }).Count -eq 0) { break }
    Start-Sleep -Seconds 1
}
$nuevos = @(Get-Process -Name POWERPNT, EXCEL -ErrorAction SilentlyContinue | Where-Object { $procPrevios -notcontains $_.Id })
foreach ($p in $nuevos) {
    Write-Host ("Cerrando proceso propio colgado: " + $p.ProcessName + " " + $p.Id) -ForegroundColor Yellow
    try { Stop-Process -Id $p.Id -Force } catch { }
}

# ================================================================ resultados
if ($script:ResultadosEstilos.Count -gt 0) { $script:ResultadosEstilos | Export-Csv -Path (Join-Path $Salida 'estilos-tabla-resultados.csv') -NoTypeInformation -Encoding UTF8 }
$errTxt = @($script:Errores | ForEach-Object { "{0} | {1} | {2}" -f $_.Archivo, $_.Caso, $_.Error })
Set-Content -Path (Join-Path $Salida 'errores.txt') -Value $errTxt -Encoding UTF8
Set-Content -Path (Join-Path $Salida 'notas.txt') -Value @($script:Notas) -Encoding UTF8

Write-Host ''
Write-Host '== RESUMEN' -ForegroundColor Green
$verif | Format-Table -AutoSize | Out-String | Write-Host
Write-Host ("Estilos de tabla: aceptados " + @($script:ResultadosEstilos | Where-Object { $_.Aceptado -and $_.Grupo -ne 'CONTROL' }).Count + " de " + $EstilosTabla.Count)
Write-Host ("Errores registrados: " + $script:Errores.Count)
$script:Errores | Format-Table -AutoSize -Wrap | Out-String -Width 250 | Write-Host

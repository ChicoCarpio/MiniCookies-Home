# Libro sintetico de la etapa 5 (paginacion): una hoja por configuracion de impresion.
# Uso: py comparar-xlsx/generar-corpus-impresion.py  -> comparar-xlsx/corpus/24-impresion.xlsx
import os
import random
from openpyxl import Workbook
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.page import PageMargins

salida = os.path.join(os.path.dirname(__file__), 'corpus', '24-impresion.xlsx')
wb = Workbook()
wb.remove(wb.active)


def hoja(nombre, filas=110, cols=26, anchos=None, altos=None, papel=9, orient='portrait', escala=None, margenes=None):
    ws = wb.create_sheet(nombre)
    rnd = random.Random(nombre)
    for f in range(1, filas + 1):
        for c in range(1, cols + 1):
            ws.cell(f, c, f'T{f}' if (f + c) % 3 else f * c)
    for c in range(1, cols + 1):
        w = anchos(c) if anchos else 8.43
        if w is not None:
            ws.column_dimensions[ws.cell(1, c).column_letter].width = w
    if altos:
        for f in range(1, filas + 1):
            h = altos(f)
            if h is not None:
                ws.row_dimensions[f].height = h
    ws.page_setup.paperSize = papel
    ws.page_setup.orientation = orient
    if escala:
        ws.page_setup.scale = escala
    m = margenes or (0.7, 0.7, 0.75, 0.75, 0.3, 0.3)
    ws.page_margins = PageMargins(left=m[0], right=m[1], top=m[2], bottom=m[3], header=m[4], footer=m[5])
    return ws


def ajuste(ws, ancho, alto):
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = ancho
    ws.page_setup.fitToHeight = alto


# 1-3: papel y orientacion
hoja('Base')
hoja('Horizontal', orient='landscape')
hoja('Carta', papel=1)
hoja('A3', papel=8)
hoja('Legal', papel=5)
hoja('B4', papel=12)
# 4: escalas
for e in (40, 70, 130, 200):
    hoja(f'Escala{e}', escala=e)
# 5: ancho y alto variados
rnd = random.Random(7)
anch = [rnd.choice([3, 5, 8.43, 12, 20, 35, 60]) for _ in range(40)]
alt = [rnd.choice([6, 12.75, 15, 15, 15, 24, 30, 48]) for _ in range(160)]
hoja('Variado', filas=160, cols=40, anchos=lambda c: anch[c - 1], altos=lambda f: alt[f - 1])
hoja('VariadoH', filas=160, cols=40, anchos=lambda c: anch[c - 1], altos=lambda f: alt[f - 1], orient='landscape', papel=1)
# 6: ajustar a
for n, m in ((1, 0), (1, 1), (2, 2), (0, 2), (2, 1), (3, 3), (5, 0), (0, 1), (4, 4)):
    ws = hoja(f'Ajuste{n}x{m}', papel=1); ajuste(ws, n, m)
ws = hoja('Ajuste2x2A4'); ajuste(ws, 2, 2)
ws = hoja('Ajuste3x0H', orient='landscape', papel=1); ajuste(ws, 3, 0)
ws = hoja('AjusteVariado', filas=160, cols=40, anchos=lambda c: anch[c - 1], altos=lambda f: alt[f - 1], papel=1); ajuste(ws, 2, 0)
ws = hoja('AjusteVariado2', filas=160, cols=40, anchos=lambda c: anch[c - 1], altos=lambda f: alt[f - 1], papel=1, orient='landscape'); ajuste(ws, 3, 4)
# 7: saltos manuales
ws = hoja('SaltosManuales')
for f in (20, 70):
    ws.row_breaks.append(Break(id=f))
for c in (4, 15):
    ws.col_breaks.append(Break(id=c))
ws = hoja('SaltosAjuste', papel=1); ajuste(ws, 2, 3)
ws.row_breaks.append(Break(id=30)); ws.col_breaks.append(Break(id=8))
# 8: area de impresion
ws = hoja('Area'); ws.print_area = 'B3:M60'
ws = hoja('AreaMultiple'); ws.print_area = 'A1:E20,H5:K30'
ws = hoja('AreaChica'); ws.print_area = 'A1:C5'
# 9: titulos
ws = hoja('Titulos'); ws.print_title_rows = '1:2'; ws.print_title_cols = 'A:A'
ws = hoja('TitulosFilas'); ws.print_title_rows = '1:3'
ws = hoja('TitulosArea'); ws.print_title_rows = '2:2'; ws.print_area = 'A1:Z100'
# 10: margenes, centrado, cuadricula, encabezados
hoja('MargenesAngostos', margenes=(0.25, 0.25, 0.75, 0.75, 0.3, 0.3))
hoja('MargenesAmplios', margenes=(1.5, 1.5, 1.5, 1.5, 0.5, 0.5))
ws = hoja('Centrado'); ws.print_options.horizontalCentered = True; ws.print_options.verticalCentered = True
ws = hoja('Cuadricula'); ws.print_options.gridLines = True; ws.print_options.headings = True
ws = hoja('EncabezadoPie')
ws.oddHeader.center.text = 'Titulo &P de &N'; ws.oddFooter.left.text = '&F'; ws.oddFooter.right.text = '&D'
ws = hoja('EncabezadoGrande', margenes=(0.7, 0.7, 0.75, 0.75, 0.3, 0.3))
ws.oddHeader.center.text = 'Grande'; ws.oddHeader.center.size = 36
# 11: ocultas
ws = hoja('Ocultas')
for f in range(10, 25):
    ws.row_dimensions[f].hidden = True
for c in (3, 4, 5, 9):
    ws.column_dimensions[ws.cell(1, c).column_letter].hidden = True
# 12: orden de paginas
ws = hoja('SobreLuegoAbajo'); ws.page_setup.pageOrder = 'overThenDown'
# 13: celdas combinadas que cruzan el corte
ws = hoja('Combinadas'); ws.merge_cells('H40:L60'); ws.merge_cells('A44:C52')
# 14: letra grande -> filas altas
ws = hoja('LetraGrande')
from openpyxl.styles import Font
for f in range(1, 80):
    for c in range(1, 12):
        ws.cell(f, c).font = Font(name='Arial', size=18)
os.makedirs(os.path.dirname(salida), exist_ok=True)
wb.save(salida)
print('OK', salida, len(wb.sheetnames), 'hojas')

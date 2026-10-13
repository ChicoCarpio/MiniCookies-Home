# Libro de calibracion de la etapa 5 (en Carta): encabezados y pies, encabezados de fila/columna, titulos repetidos, centrado y cuadricula impresa.
# Uso: py comparar-xlsx/generar-corpus-encabezados.py -> comparar-xlsx/corpus/27-impresion-extras.xlsx
import os
from openpyxl import Workbook
from openpyxl.worksheet.page import PageMargins
from openpyxl.styles import Font

salida = os.path.join(os.path.dirname(__file__), 'corpus', '27-impresion-extras.xlsx')
wb = Workbook()
wb.remove(wb.active)


def hoja(nombre, filas=6, cols=3, margenes=(0.7, 0.7, 0.75, 0.75, 0.3, 0.3)):
    ws = wb.create_sheet(nombre)
    ws.page_setup.paperSize = 1
    ws.page_margins = PageMargins(left=margenes[0], right=margenes[1], top=margenes[2], bottom=margenes[3], header=margenes[4], footer=margenes[5])
    for f in range(1, filas + 1):
        for c in range(1, cols + 1):
            ws.cell(f, c, f * 10 + c)
    return ws


ws = hoja('HF')
ws.oddHeader.left.text = 'IzqH'; ws.oddHeader.center.text = 'CentroH'; ws.oddHeader.right.text = 'DerH'
ws.oddFooter.left.text = 'IzqP'; ws.oddFooter.center.text = 'Pagina &P de &N'; ws.oddFooter.right.text = '&F'
ws = hoja('HFTam')
ws.oddHeader.center.text = 'Grande'; ws.oddHeader.center.size = 24
ws.oddFooter.center.text = 'Chico'; ws.oddFooter.center.size = 8
ws = hoja('HFFuente')
ws.oddHeader.center.text = 'Arial negrita'; ws.oddHeader.center.font = 'Arial,Bold'
ws.oddFooter.center.text = 'Times'; ws.oddFooter.center.font = 'Times New Roman,Regular'
ws = hoja('HFMargen', margenes=(0.7, 0.7, 1.2, 1.2, 0.8, 0.8))
ws.oddHeader.center.text = 'Margen'; ws.oddFooter.center.text = 'Margen'
for n in (9, 99, 999):
    ws = hoja(f'Enc{n}', filas=n, cols=3)
    ws.print_options.headings = True
    ws.print_options.gridLines = True
ws = hoja('Cuad', filas=4, cols=3); ws.print_options.gridLines = True
ws = hoja('Titulos', filas=130, cols=3); ws.print_title_rows = '1:2'; ws.print_options.gridLines = True
ws = hoja('TitulosCol', filas=3, cols=40); ws.print_title_cols = 'A:B'; ws.print_options.gridLines = True
ws = hoja('Centrado', filas=4, cols=3); ws.print_options.horizontalCentered = True; ws.print_options.verticalCentered = True; ws.print_options.gridLines = True
ws = hoja('CentradoH', filas=4, cols=3); ws.print_options.horizontalCentered = True; ws.print_options.gridLines = True
ws = hoja('Bordes', filas=4, cols=3)
from openpyxl.styles import Border, Side
for f in range(1, 5):
    for c in range(1, 4):
        ws.cell(f, c).border = Border(left=Side(style='thin'), right=Side(style='medium'), top=Side(style='thin'), bottom=Side(style='thick'))
os.makedirs(os.path.dirname(salida), exist_ok=True)
wb.save(salida)
print('OK', salida, len(wb.sheetnames), 'hojas')

# Libro de calibracion de la etapa 5: hojas Carta con cuadricula impresa para leer, en el PDF de Excel, donde cae cada columna y cada fila.
# Uso: py comparar-xlsx/generar-corpus-calibre.py -> comparar-xlsx/corpus/25-impresion-calibre.xlsx
import os
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.page import PageMargins

salida = os.path.join(os.path.dirname(__file__), 'corpus', '25-impresion-calibre.xlsx')
wb = Workbook()
wb.remove(wb.active)


def base(nombre, escala=100):
    ws = wb.create_sheet(nombre)
    ws.page_setup.paperSize = 1
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.scale = escala
    ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.25, bottom=0.25, header=0.1, footer=0.1)
    ws.print_options.gridLines = True
    return ws


ANCHOS = [1, 2, 3, 5, 6.5, 8.43, 10, 12, 15, 20, 25, 30, 40, 50, 70, 100]
ALTOS = [3, 6, 9, 12, 12.75, 15, 18, 20, 24, 30, 36, 48, 60, 90]
for escala in (100, 50, 200, 77):
    ws = base(f'Cols{escala}', escala)
    for i, w in enumerate(ANCHOS, 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
        for f in range(1, 6):
            ws.cell(f, i, 1)
    ws = base(f'Filas{escala}', escala)
    for i, h in enumerate(ALTOS, 1):
        ws.row_dimensions[i].height = h
        for c in range(1, 4):
            ws.cell(i, c, 1)
# filas sin alto guardado: Excel las calcula con la letra
FUENTES = [('Calibri', 8), ('Calibri', 11), ('Calibri', 14), ('Calibri', 20), ('Calibri', 36), ('Arial', 10), ('Arial', 14), ('Times New Roman', 12), ('Courier New', 10), ('Verdana', 9), ('Tahoma', 8), ('Segoe UI', 12)]
ws = base('FilasAuto')
for i, (n, t) in enumerate(FUENTES, 1):
    for c in range(1, 4):
        ws.cell(i, c, 'Ag').font = Font(name=n, size=t)
os.makedirs(os.path.dirname(salida), exist_ok=True)
wb.save(salida)
print('OK', salida, len(wb.sheetnames), 'hojas')

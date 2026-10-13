# Libro de la etapa 5 para medir como elige Excel la escala de "ajustar a N x M paginas": una fila de n columnas (solo manda el ancho)
# y una columna de n filas (solo manda el alto), en Carta vertical.
# Uso: py comparar-xlsx/generar-corpus-ajuste.py -> comparar-xlsx/corpus/26-impresion-ajuste.xlsx
import os
from openpyxl import Workbook
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties

salida = os.path.join(os.path.dirname(__file__), 'corpus', '26-impresion-ajuste.xlsx')
wb = Workbook()
wb.remove(wb.active)


def hoja(nombre, n, m):
    ws = wb.create_sheet(nombre)
    ws.page_setup.paperSize = 1
    ws.page_margins = PageMargins(left=0.7, right=0.7, top=0.75, bottom=0.75, header=0.3, footer=0.3)
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = n
    ws.page_setup.fitToHeight = m
    return ws


for cols in (8, 11, 13, 14, 16, 18, 20, 22, 24, 26, 28, 30, 34, 36, 40, 44, 50, 60, 90):
    for N in (1, 2, 3, 4, 5, 6):
        ws = hoja(f'A{cols}c{N}', N, 0)
        for c in range(1, cols + 1):
            ws.cell(1, c, 'x')
            ws.cell(2, c, 1)
for filas in (50, 110, 200):
    for M in (1, 2, 3):
        ws = hoja(f'F{filas}f{M}', 0, M)
        for f in range(1, filas + 1):
            ws.cell(f, 1, 'x')
            ws.cell(f, 2, f)
os.makedirs(os.path.dirname(salida), exist_ok=True)
wb.save(salida)
print('OK', salida, len(wb.sheetnames), 'hojas')

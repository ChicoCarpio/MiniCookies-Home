# Libro grande sintetico para medir el visor (sin Excel): 100.000 filas x 10 columnas (1.000.000 de celdas). Uso: py generar-grande.py salida.xlsx
import sys
import openpyxl
wb = openpyxl.Workbook(write_only=True)
ws = wb.create_sheet('Grande')
for i in range(100000):
    ws.append(['fila %d' % i, i, i * 1.5, 'texto %d' % (i % 97), i % 7 == 0, i * 3.25, 'ñ%d' % i, i * 2, i / 7, 'fin'])
wb.save(sys.argv[1])

# Libros de OTROS productores (openpyxl y pandas), para comprobar que el lector acepta lo que no escribe Excel:
# fórmulas sin valor guardado, cadenas en línea, estilos mínimos, sin sharedStrings, etc.
# Uso: py comparar-xlsx/generar-corpus-otros.py   → corpus\2x-*.xlsx
import datetime
import os
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.comments import Comment
from openpyxl.workbook.defined_name import DefinedName

destino = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'corpus')
os.makedirs(destino, exist_ok=True)

# 20: variado, con fórmulas SIN valor calculado (openpyxl no calcula)
wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Datos'
ws.append(['Nombre', 'Cantidad', 'Precio', 'Total'])
for i in range(2, 12):
    ws.append(['Producto %d' % i, i, i * 1.5, '=B%d*C%d' % (i, i)])
ws['F1'] = datetime.datetime(2024, 2, 29, 13, 45)
ws['F2'] = datetime.date(2024, 12, 31)
ws['F3'] = datetime.time(8, 30)
ws['G1'] = 'texto con\nsalto'
ws['G2'] = True
ws['G3'] = '=SUM(B2:B11)'
ws['H1'].font = Font(bold=True, italic=True, color='FF0000', name='Arial', size=14)
ws['H1'] = 'estilo'
ws['H2'].fill = PatternFill('solid', fgColor='FFFF00')
ws['H2'].border = Border(left=Side(style='thin'), right=Side(style='medium', color='0000FF'), top=Side(style='dashed'), bottom=Side(style='double'))
ws['H3'].alignment = Alignment(horizontal='center', vertical='top', wrap_text=True, indent=2, text_rotation=45)
ws['H3'] = 'alineado'
ws['I1'].number_format = '#,##0.00'
ws['I1'] = 1234.5
ws['I2'].number_format = '0.0%'
ws['I2'] = 0.256
ws['I3'].number_format = 'dd/mm/yyyy'
ws['I3'] = datetime.date(2025, 3, 15)
ws.merge_cells('K1:M2')
ws['K1'] = 'combinada'
ws.column_dimensions['A'].width = 25
ws.column_dimensions['B'].hidden = True
ws.row_dimensions[3].height = 40
ws.freeze_panes = 'B2'
ws.sheet_properties.tabColor = '1072BA'
ws.auto_filter.ref = 'A1:D11'
dv = DataValidation(type='list', formula1='"a,b,c"', allow_blank=True)
ws.add_data_validation(dv)
dv.add('N1:N5')
ws.conditional_formatting.add('B2:B11', CellIsRule(operator='greaterThan', formula=['5'], fill=PatternFill('solid', bgColor='FFC7CE')))
ws.conditional_formatting.add('C2:C11', ColorScaleRule(start_type='min', start_color='F8696B', end_type='max', end_color='63BE7B'))
ws['A2'].comment = Comment('un comentario', 'autor')
ws['A3'].hyperlink = 'https://example.com'
ws['A4'].hyperlink = "#'Datos'!A1"
ws2 = wb.create_sheet('Oculta')
ws2.sheet_state = 'hidden'
ws2['A1'] = 'secreto'
ws3 = wb.create_sheet('Otra hoja con espacios')
ws3['A1'] = "='Datos'!A1"
wb.defined_names['Total'] = DefinedName('Total', attr_text="Datos!$D$2:$D$11")
wb.save(os.path.join(destino, '20-openpyxl-variado.xlsx'))

# 21: openpyxl solo lectura de valores (celdas inlineStr, sin estilos) — write_only
wb = openpyxl.Workbook(write_only=True)
w = wb.create_sheet('Plano')
w.append(['a', 1, 2.5, None, True])
w.append(['b', None, 'ñ', datetime.datetime(2020, 1, 1), '=A1'])
wb.save(os.path.join(destino, '21-openpyxl-write-only.xlsx'))

# 22: pandas
import pandas as pd
df = pd.DataFrame({'texto': ['x', 'y', None], 'entero': [1, 2, 3], 'real': [1.5, None, 3.25], 'fecha': pd.to_datetime(['2024-01-01', '2024-06-15', None]), 'logico': [True, False, True]})
with pd.ExcelWriter(os.path.join(destino, '22-pandas.xlsx'), engine='openpyxl') as w:
    df.to_excel(w, sheet_name='df', index=True)
print('ok')

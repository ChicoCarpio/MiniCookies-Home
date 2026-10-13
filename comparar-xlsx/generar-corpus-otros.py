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
# 23: geometria y dibujo (etapa 4). openpyxl no guarda alto de fila salvo que se pida: Excel ajusta las filas al abrir.
wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Geometria'
tam = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36]
for i, t in enumerate(tam, start=1):
    ws.cell(row=i, column=1, value='Calibri %d' % t).font = Font(name='Calibri', size=t)
    ws.cell(row=i, column=2, value='Arial %d' % t).font = Font(name='Arial', size=t)
    ws.cell(row=i, column=3, value='Times %d' % t).font = Font(name='Times New Roman', size=t)
    ws.cell(row=i, column=4, value='Courier %d' % t).font = Font(name='Courier New', size=t)
ws.row_dimensions[14].height = 30
ws.row_dimensions[16].hidden = True
for i, a in enumerate([4, 8.43, 12, 20, 33.7, 50, 1, 0.5], start=6):
    ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = a
ws.column_dimensions['M'].hidden = True
ws['A14'] = 'alto 30'
ws['A15'] = 'antes de oculta'
ws['A17'] = 'despues de oculta'

# Estilos de celda: rellenos, patrones, fuentes y colores
es = wb.create_sheet('Estilos')
es.column_dimensions['A'].width = 18
es.column_dimensions['B'].width = 14
for i, (n, b, it, u, tc) in enumerate([('normal', 0, 0, None, '000000'), ('negrita', 1, 0, None, '000000'), ('cursiva', 0, 1, None, '000000'), ('subrayado', 0, 0, 'single', '000000'),
                                        ('doble subrayado', 0, 0, 'double', '000000'), ('negrita cursiva', 1, 1, None, '000000'), ('tachado', 0, 0, None, 'C00000'), ('rojo', 0, 0, None, 'FF0000'),
                                        ('azul', 0, 0, None, '0000FF'), ('verde sobre amarillo', 0, 0, None, '006100')], start=1):
    c = es.cell(row=i, column=1, value=n)
    c.font = Font(bold=bool(b), italic=bool(it), underline=u, strike=(n == 'tachado'), color=tc)
    if n.startswith('verde'):
        c.fill = PatternFill('solid', fgColor='FFFF00')
patrones = ['solid', 'gray125', 'gray0625', 'darkGray', 'mediumGray', 'lightGray', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis',
            'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis']
for i, p in enumerate(patrones, start=1):
    c = es.cell(row=i, column=3, value=p)
    c.fill = PatternFill(patternType=p, fgColor='2F5597', bgColor='FFE699')
for i, col in enumerate(['FF0000', '00B050', '0070C0', 'FFC000', '7030A0', 'D9D9D9', '000000'], start=1):
    es.cell(row=i, column=5).fill = PatternFill('solid', fgColor=col)

# Bordes
bo = wb.create_sheet('Bordes')
estilos = ['thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
for i, e in enumerate(estilos, start=2):
    s = Side(style=e, color='000000')
    bo.cell(row=i * 2, column=2, value=e).border = Border(left=s, right=s, top=s, bottom=s)
    bo.cell(row=i * 2, column=4).border = Border(bottom=Side(style=e, color='C00000'))
    bo.cell(row=i * 2, column=6).border = Border(left=Side(style=e, color='0000FF'), diagonal=Side(style='thin', color='00B050'), diagonalUp=True, diagonalDown=True)
bo.column_dimensions['B'].width = 22

# Alineacion, ajuste de texto, sangria, rotacion
al = wb.create_sheet('Alineacion')
for c in range(1, 5):
    al.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 18
horiz = ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed']
vert = ['top', 'center', 'bottom']
for i, h in enumerate(horiz, start=1):
    for j, v in enumerate(vert, start=1):
        c = al.cell(row=i * 2, column=j, value='%s %s' % (h, v))
        c.alignment = Alignment(horizontal=h, vertical=v)
        al.row_dimensions[i * 2].height = 30
al['E2'] = 'texto largo que se ajusta al ancho de la celda y ocupa varias lineas'
al['E2'].alignment = Alignment(wrap_text=True)
al['F2'] = 'sangria 3'
al['F2'].alignment = Alignment(horizontal='left', indent=3)
for k, r in enumerate([0, 45, 90, 135, 180], start=0):
    c = al.cell(row=4 + k * 3, column=5, value='rot %d' % r)
    c.alignment = Alignment(text_rotation=r)
    al.row_dimensions[4 + k * 3].height = 60
al['G2'] = 'reducir hasta ajustar a la celda'
al['G2'].alignment = Alignment(shrink_to_fit=True)
al.column_dimensions['G'].width = 12
al['A20'] = 'Desborde a la derecha de un texto largo sobre celdas vacias'
al['A22'] = 'No desborda'
al['B22'] = 'bloquea'
al['D24'] = 'derecha: desborda hacia la izquierda'
al['D24'].alignment = Alignment(horizontal='right')
al['D26'] = 'centrado: desborda a ambos lados'
al['D26'].alignment = Alignment(horizontal='center')
al['A28'] = 1234567890123
al['B28'] = 'numero que no cabe se ve como ###'
al.merge_cells('A30:C32')
al['A30'] = 'combinada centrada'
al['A30'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
al['A30'].fill = PatternFill('solid', fgColor='DDEBF7')
ro = wb.create_sheet('Rotacion')
for c in range(1, 4):
    ro.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 11.4
fila = 1
for ang in [45, 90, 135, 180]:
    for v in ['top', 'center', 'bottom']:
        ro.row_dimensions[fila].height = 60
        for j, h in enumerate(['left', 'center', 'right'], start=1):
            ro.cell(row=fila, column=j, value='Ab %d' % ang).alignment = Alignment(horizontal=h, vertical=v, text_rotation=ang)
        fila += 1
tr = wb.create_sheet('Tramas')
for i, pt in enumerate(patrones):
    for rep in range(3):
        tr.cell(row=i * 3 + rep + 1, column=1).fill = PatternFill(patternType=pt, fgColor='000000', bgColor='FFFFFF')
        tr.cell(row=i * 3 + rep + 1, column=2).fill = PatternFill(patternType=pt, fgColor='FF0000', bgColor='FFFF00')
aj = wb.create_sheet('Ajuste')
for col, ancho in zip('ABC', [12, 20, 30]):
    aj.column_dimensions[col].width = ancho
frases = ['corto', 'un texto un poco mas largo que una linea', 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore',
          'palabrasuperlargasinespaciosquenocabeenlacolumnadeterminada y otra', 'linea uno\nlinea dos\nlinea tres', 'ñandú café, canción — acentos y símbolos € £ ¥ © ® ™']
for i, t in enumerate(frases, start=1):
    for j, col in enumerate('ABC'):
        c = aj['%s%d' % (col, i)]
        c.value = t
        c.alignment = Alignment(wrap_text=True, vertical=['top', 'center', 'bottom'][j])
        if i % 2 == 0:
            c.font = Font(name=['Arial', 'Times New Roman', 'Calibri'][j], size=[9, 12, 14][j])
aj['A8'] = 'linea uno\nlinea dos (sin ajuste)'
aj.merge_cells('B8:C9')
aj['B8'] = 'combinada con ajuste y bastante texto para ocupar mas de una linea dentro de la combinada'
aj['B8'].alignment = Alignment(wrap_text=True, vertical='top')

nu = wb.create_sheet('NumerosAngostos')
anchos = [4, 6, 8, 8.43, 10, 12, 16]
valores = [(123456789012, 'General'), (0.000012345, 'General'), (1234.5678, 'General'), (12345678.9, 'General'), (-123456789, 'General'), (0.1234567890123, 'General'),
           (45000, 'dd/mm/yyyy'), (0.7291666, 'h:mm:ss'), (0.256, '0.00%'), (1234567.891, '#,##0.00'), (1234567.891, '0.000'), (9.99999999, 'General'), (100000000, 'General')]
for j, a in enumerate(anchos, start=1):
    nu.column_dimensions[openpyxl.utils.get_column_letter(j)].width = a
    for i, (v, f) in enumerate(valores, start=1):
        c = nu.cell(row=i, column=j, value=v)
        c.number_format = f
ai = wb.create_sheet('Aislado')
ai['B2'].fill = PatternFill('solid', fgColor='FF0000')
ai['D4'].fill = PatternFill('solid', fgColor='0000FF')
ai['B2'] = 'x'
wb.save(os.path.join(destino, '23-openpyxl-dibujo.xlsx'))

print('ok')

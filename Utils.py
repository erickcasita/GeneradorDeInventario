from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, numbers
import datetime,locale
def headers():
    wb = Workbook()
    ws = wb.active
    ws.sheet_view.showGridLines = False
    ws['B2'] = "Reporte de Inventario Físico de Líquido"
    ws['B2'].font = Font(color="000000", size="14", name="Comic Sans MS", underline="single")
    ws['B2'].alignment = Alignment(horizontal='general',vertical='center')
    ws.row_dimensions[2].height = 47.40
    #Unions Cells for date
    ws.merge_cells('H2:K2')
    locale.setlocale(locale.LC_ALL, 'es_ES.utf8')
    date = datetime.datetime.strftime(datetime.datetime.now(),'%A, %d de %B del %Y')
    ws['H2'] = date
    ws['H2'].number_format = numbers.FORMAT_DATE_TIME5
    ws['H2'].font = Font(color="000000", size="12", name="Arial", bold=True)
    ws['H2'].alignment = Alignment(horizontal='center')
    #Unions Cell for  headers "Almacén"
    ws.merge_cells('B3:D3')
    ws['B3'] = "ALMACÉN"
    ws['B3'].font = Font(color="FFFFFF", size="13.5", name="MS Sans Serif")
    ws['B3'].alignment = Alignment(horizontal='center',vertical="center")
    #ws['B3'].fill = PatternFill(bgColor="FFC7CE", fill_type = "solid")
    ws.row_dimensions[3].height = 47.40
    ws.column_dimensions['B'].width = 53.67
    #Unions Cell for Headers "San Andres" and "Juan Diaz Covarrubias" and "Totales"
    ws.merge_cells('E3:G3')
    ws.merge_cells('K3:N3')
    ws.merge_cells('R3:T3')
    ws['E3'] = "San Andrés"
    ws['E3'].font = Font(color="FFFFFF", size="11", name="Comic Sans MS")
    ws['E3'].alignment = Alignment(horizontal='center',vertical="center")
    ws['K3'] = "Juan Díaz Covarrubias"
    ws['K3'].font = Font(color="FFFFFF", size="11", name="Comic Sans MS")
    ws['K3'].alignment = Alignment(horizontal='center',vertical="center")
    ws['R3'] = "TOTALES"
    ws['R3'].font = Font(color="FFFFFF", size="11", name="Comic Sans MS")
    ws['R3'].alignment = Alignment(horizontal='center',vertical="center")
    #Header "Stock Min and Stock Max SAT and JDC"
    ws['I3'] = "Stock Mín"
    ws['J3'] = "Stock MAX"
    ws['P3'] = "Stock Mín"
    ws['Q3'] = "Stock MAX"
    ws['I3'].font = Font(color="FFFF00", size="10", name="Comic Sans MS")
    ws['I3'].alignment = Alignment(wrap_text=True,horizontal='center',vertical="center")
    ws['J3'].font = Font(color="FFFF00", size="10", name="Comic Sans MS")
    ws['J3'].alignment = Alignment(wrap_text=True,horizontal='center',vertical="center")
    ws['P3'].font = Font(color="FFFF00", size="10", name="Comic Sans MS")
    ws['P3'].alignment = Alignment(wrap_text=True,horizontal='center',vertical="center")
    ws['Q3'].font = Font(color="FFFF00", size="10", name="Comic Sans MS")
    ws['Q3'].alignment = Alignment(wrap_text=True,horizontal='center',vertical="center")
    for i in  range(1, 21):
        ws.cell(3,i).fill = PatternFill(fgColor="222b35", fill_type = "solid")
    
    wb.save("test.xlsx")
    
from Mark_Sheet import SheetMarker
import Const_txt_excel as CET
import Const_Balise_Excel as CEB
from openpyxl.styles import Font,Alignment,PatternFill,Side,Border

FORMAT_COMPTA = "0,00 €"
class InitSheet:
    def __init__(self, ws):
        self.ws = ws
        # self.set_marker()
        # self.mks = SheetMarker(ws)  # Assurez-vous que SheetMarker est défini correctement
        self.write_start()
        self.write_bilan()
        self.write_income()
        self.write_loss()
        self.write_detail()


    def write_start(self):
        cell_date_b = self.ws.cell(1, 1)
        cell_start_b = self.ws.cell(2, 1)
        cell_sold_st = self.ws.cell(2,2)
        cell_date_txt = self.ws.cell(1,2) 
        
        font_style = Font(size=14, bold=True)
        cell_date_b.font = font_style
        cell_start_b.font = font_style
        cell_sold_st.font = font_style
        cell_date_txt.font = font_style
        
        cell_sold_st.number_format = FORMAT_COMPTA
        
        cell_date_b.alignment = Alignment(horizontal='center', vertical='center')
        cell_start_b.alignment = Alignment(horizontal='center', vertical='center')
        cell_sold_st.alignment = Alignment(horizontal='right', vertical='center')
        cell_date_txt.alignment = Alignment(horizontal='center', vertical='center')
        
        
        cell_date_b.value = CEB.BILAN_DATE
        cell_date_b.number_format = f';;;"{CET.INIT_DATE}"'
        cell_start_b.value = CEB.BILAN_ST_SOLD
        cell_start_b.number_format = f';;;"{CET.INIT_SOLD_ST}"'
        
    def write_bilan(self):
        cell_title_b  = self.ws.cell(4,1)
        cell_title_ex = self.ws.cell(4,2)
        cell_title_re = self.ws.cell(4,3)
        cell_val_ex   = self.ws.cell(5,2)
        cell_val_re   = self.ws.cell(5,3)
        cell_valid_b  = self.ws.cell(3,4)
        
        cell_title_b.font = Font(size=16, bold=True)
        cell_title_ex.font = Font(size=16, bold=True)
        cell_title_re.font = Font(size=16, bold=True)
        cell_val_ex.font = Font(size=14)
        cell_val_re.font = Font(size=14)
        cell_valid_b.font = Font(size=14)
        
        cell_title_b.alignment = Alignment(horizontal='center', vertical='center')
        cell_title_ex.alignment = Alignment(horizontal='center', vertical='center')
        cell_title_re.alignment = Alignment(horizontal='center', vertical='center')
        cell_val_ex.alignment = Alignment(horizontal='right', vertical='center')
        cell_val_re.alignment = Alignment(horizontal='right', vertical='center')
        cell_valid_b.alignment = Alignment(horizontal='center', vertical='center')
        
        cell_val_ex.number_format = FORMAT_COMPTA
        cell_val_re.number_format = FORMAT_COMPTA
        
        fill_color1 = PatternFill(start_color='83CCEB', end_color='83CCEB', fill_type='solid')
        fill_color2 = PatternFill(start_color='C0E6F5', end_color='C0E6F5', fill_type='solid')
        cell_title_b.fill = fill_color1
        cell_title_ex.fill = fill_color1
        cell_title_re.fill = fill_color1
        cell_val_ex.fill = fill_color2
        cell_val_re.fill = fill_color2

        cell_title_b.border = Border(left=Side(style='thick'),right=Side(style='thin'),top=Side(style='thick'),bottom=Side(style='thick'))
        cell_title_ex.border = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thick'),bottom=Side(style='thin'))
        cell_title_re.border = Border(left=Side(style='thin'),right=Side(style='thick'),top=Side(style='thick'),bottom=Side(style='thin'))
        cell_val_ex.border = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thick'))
        cell_val_re.border = Border(left=Side(style='thin'),right=Side(style='thick'),top=Side(style='thin'),bottom=Side(style='thick'))

        self.ws.merge_cells(start_row=4, start_column=1, end_row=5, end_column=1)
        
        cell_title_b.value = CEB.BILAN_ED_SOLD
        cell_title_b.number_format = f';;;"{CET.INIT_TITLE_BILAN}"'
        cell_valid_b.value = CEB.BILAN_ED_VAL
        cell_valid_b.number_format = f';;;"{CET.INIT_VALID}"'
        
        cell_title_ex.value = CET.INIT_TITLE_EXPECTED
        cell_title_re.value = CET.INIT_TITLE_REAL

    def write_income(self):
        cell_title_b = self.ws.cell(7,1)
        cell_title_tot = self.ws.cell(8,1)
        cell_val_ex = self.ws.cell(8,2)
        cell_val_re = self.ws.cell(8,3)
        cell_title_name = self.ws.cell(10,1)
        cell_title_ex = self.ws.cell(10,2)
        cell_title_re = self.ws.cell(10,3)

        cell_title_b.font = Font(size=16, bold=True)
        cell_title_tot.font = Font(size=16, bold=True)
        cell_title_name.font = Font(size=14)
        cell_title_ex.font = Font(size=14)
        cell_title_re.font = Font(size=14)
        cell_val_ex.font = Font(size=14)
        cell_val_re.font = Font(size=14)
        
        cell_title_b.alignment = Alignment(horizontal='center', vertical='center')
        cell_title_tot.alignment = Alignment(horizontal='center', vertical='center')
        cell_title_name.alignment = Alignment(horizontal='center', vertical='center')
        cell_title_ex.alignment = Alignment(horizontal='center', vertical='center')
        cell_title_re.alignment = Alignment(horizontal='center', vertical='center')
        cell_val_ex.alignment = Alignment(horizontal='right', vertical='center')
        cell_val_re.alignment = Alignment(horizontal='right', vertical='center')
        
        cell_val_ex.number_format = FORMAT_COMPTA
        cell_val_re.number_format = FORMAT_COMPTA
        
        fill_color = PatternFill(start_color='C1F0C8', end_color='C1F0C8', fill_type='solid')
        cell_title_b.fill = PatternFill(start_color='83E28E', end_color='83E28E', fill_type='solid')
        cell_title_tot.fill = fill_color
        cell_title_name.fill = fill_color
        cell_title_ex.fill = fill_color
        cell_title_re.fill = fill_color
        cell_val_ex.fill = fill_color
        cell_val_re.fill = fill_color

        cell_title_b.border = Border(left=Side(style='thick'),right=Side(style='thick'),top=Side(style='thick'),bottom=Side(style='thick'))
        cell_title_tot.border = Border(left=Side(style='thick'),right=Side(style='thin'),top=Side(style='thick'),bottom=Side(style='thick'))
        cell_title_name.border = Border(left=Side(style='thick'),right=Side(style='thin'),top=Side(style='thick'),bottom=Side(style='thick'))
        cell_title_ex.border = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thick'),bottom=Side(style='thick'))
        cell_title_re.border = Border(left=Side(style='thin'),right=Side(style='thick'),top=Side(style='thick'),bottom=Side(style='thick'))
        cell_val_ex.border = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thick'),bottom=Side(style='thick'))
        cell_val_re.border = Border(left=Side(style='thin'),right=Side(style='thick'),top=Side(style='thick'),bottom=Side(style='thick'))

        self.ws.merge_cells(start_row=7, start_column=1, end_row=7, end_column=3)
        
        cell_title_b.value = CEB.INCOME
        cell_title_b.number_format = f';;;"{CET.INIT_TITLE_IN}"'
        cell_title_tot.value = CET.INIT_TITLE_TOT
        cell_title_name.value = CET.INIT_TITLE_NAME
        cell_title_ex.value = CET.INIT_TITLE_EXPECTED
        cell_title_re.value = CET.INIT_TITLE_REAL
    
    def write_loss(self):
        cell_title_b = self.ws.cell(7,5)
        cell_title_tot = self.ws.cell(8,5)
        cell_val_ex = self.ws.cell(8,6)
        cell_val_re = self.ws.cell(8,7)
        cell_title_name = self.ws.cell(10,5)
        cell_title_ex = self.ws.cell(10,6)
        cell_title_re = self.ws.cell(10,7)

        cell_title_b.font = Font(size=16, bold=True)
        cell_title_tot.font = Font(size=16, bold=True)
        cell_title_name.font = Font(size=14)
        cell_title_ex.font = Font(size=14)
        cell_title_re.font = Font(size=14)
        cell_val_ex.font = Font(size=14)
        cell_val_re.font = Font(size=14)
        
        cell_title_b.alignment = Alignment(horizontal='center', vertical='center')
        cell_title_tot.alignment = Alignment(horizontal='center', vertical='center')
        cell_title_name.alignment = Alignment(horizontal='center', vertical='center')
        cell_title_ex.alignment = Alignment(horizontal='center', vertical='center')
        cell_title_re.alignment = Alignment(horizontal='center', vertical='center')
        cell_val_ex.alignment = Alignment(horizontal='right', vertical='center')
        cell_val_re.alignment = Alignment(horizontal='right', vertical='center')
        
        cell_val_ex.number_format = FORMAT_COMPTA
        cell_val_re.number_format = FORMAT_COMPTA
        
        fill_color = PatternFill(start_color='FBE2D5', end_color='FBE2D5', fill_type='solid')
        cell_title_b.fill = PatternFill(start_color='F7C7AC', end_color='F7C7AC', fill_type='solid')
        cell_title_tot.fill = fill_color
        cell_title_name.fill = fill_color
        cell_title_ex.fill = fill_color
        cell_title_re.fill = fill_color
        cell_val_ex.fill = fill_color
        cell_val_re.fill = fill_color

        cell_title_b.border = Border(left=Side(style='thick'),right=Side(style='thick'),top=Side(style='thick'),bottom=Side(style='thick'))
        cell_title_tot.border = Border(left=Side(style='thick'),right=Side(style='thin'),top=Side(style='thick'),bottom=Side(style='thick'))
        cell_title_name.border = Border(left=Side(style='thick'),right=Side(style='thin'),top=Side(style='thick'),bottom=Side(style='thick'))
        cell_title_ex.border = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thick'),bottom=Side(style='thick'))
        cell_title_re.border = Border(left=Side(style='thin'),right=Side(style='thick'),top=Side(style='thick'),bottom=Side(style='thick'))
        cell_val_ex.border = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thick'),bottom=Side(style='thick'))
        cell_val_re.border = Border(left=Side(style='thin'),right=Side(style='thick'),top=Side(style='thick'),bottom=Side(style='thick'))

        self.ws.merge_cells(start_row=7, start_column=5, end_row=7, end_column=7)
        
        cell_title_b.value = CEB.LOSS
        cell_title_b.number_format = f';;;"{CET.INIT_TITLE_LOS}"'
        cell_title_tot.value = CET.INIT_TITLE_TOT
        cell_title_name.value = CET.INIT_TITLE_NAME
        cell_title_ex.value = CET.INIT_TITLE_EXPECTED
        cell_title_re.value = CET.INIT_TITLE_REAL

    def write_detail(self):
        cell_detail_b = self.ws.cell(26,1)

        cell_detail_b.font = Font(size=11)
        cell_detail_b.style = "Accent1"
        cell_detail_b.alignment = Alignment(horizontal='center', vertical='center')

        # cell_detail_b.fill = PatternFill(start_color='F7C7AC', end_color='F7C7AC', fill_type='solid')

        # cell_detail_b.border = Border(left=Side(style='thick'),right=Side(style='thick'),top=Side(style='thick'),bottom=Side(style='thick'))

        self.ws.merge_cells(start_row=26, start_column=1, end_row=26, end_column=11)
        
        cell_detail_b.value = CEB.DETAIL
        cell_detail_b.number_format = f';;;"{CET.INIT_DETAIL}"'

    def write_validity(self):
        pass

if __name__ == "__main__":
    import openpyxl
    from Styles import StyleCell
    
    path = r"D:\One Drive\OneDrive\Bureau\Classeur1.xlsx"
    workbook = openpyxl.load_workbook(path)
    StyleCell(workbook)
    ws = workbook.active
    InitSheet(ws)
    workbook.save(path)
    
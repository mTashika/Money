import Const_txt_excel as CET
import Const_Balise_Excel as CEB
from openpyxl.styles import Font,Alignment,PatternFill,Side,Border
from Protection import unprotect_cell,unprotect_range,protect_ws

FORMAT_COMPTA = "0.00 €"
class InitSheet:
    def __init__(self, ws):
        self.ws = ws
        self.set_big_border()
        self.set_column_width()
        self.set_line_height()
        self.write_start()
        self.write_bilan()
        self.write_income()
        self.write_loss()
        self.write_detail()
        self.write_validity()
        self.set_protection()

    def set_big_border(self):
        for col in range(1, 11):
            cell = self.ws.cell(row=10, column=col)
            cell.border =  Border(bottom=Side(border_style='dotted', color='000000'))
        for row in range(1, 10):
            cell = self.ws.cell(row=row, column=11)
            cell.border =  Border(right=Side(border_style='dotted', color='000000'))
        
        cell = self.ws.cell(row=10, column=11)
        cell.border =  Border(bottom=Side(border_style='dotted', color='000000'),right=Side(border_style='dotted', color='000000'))

    def set_column_width(self):
        self.ws.column_dimensions['A'].width = 15
        self.ws.column_dimensions['B'].width = 17
        self.ws.column_dimensions['C'].width = 24
        self.ws.column_dimensions['D'].width = 12
        self.ws.column_dimensions['E'].width = 10
        self.ws.column_dimensions['F'].width = 15
        self.ws.column_dimensions['G'].width = 15
        self.ws.column_dimensions['H'].width = 15
        self.ws.column_dimensions['I'].width = 15
        self.ws.column_dimensions['J'].width = 15
        self.ws.column_dimensions['K'].width = 15
        
    def set_line_height(self):
        self.ws.row_dimensions[1].height = 20
        self.ws.row_dimensions[2].height = 20
        self.ws.row_dimensions[3].height = 30
        self.ws.row_dimensions[4].height = 24
        self.ws.row_dimensions[5].height = 21
        self.ws.row_dimensions[6].height = 15
        self.ws.row_dimensions[7].height = 21
        self.ws.row_dimensions[8].height = 20
        self.ws.row_dimensions[9].height = 15
        self.ws.row_dimensions[10].height = 20
        
    def write_start(self):
        cell_date_b = self.ws.cell(1, 1)
        cell_start_b = self.ws.cell(2, 1)
        cell_sold_st = self.ws.cell(2,2)
        cell_date_txt = self.ws.cell(1,2) 
        
        font_style = Font(size=14, bold=True)
        cell_date_b.font = font_style
        cell_start_b.font = font_style
        cell_sold_st.font = font_style
        cell_date_txt.font = Font(size=13, bold=True)
        
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
        cell_val_ex.value = '=B2+B8-F8'

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
        cell_title_b       = self.ws.cell(1,8)
        cell_title_ex1     = self.ws.cell(2,8)
        cell_title_ex2     = self.ws.cell(3,8)
        cell_title_id      = self.ws.cell(4,8)
        cell_title_re      = self.ws.cell(5,8)
        cell_title_path    = self.ws.cell(2,9)
        cell_title_date_ex = self.ws.cell(3,9)
        cell_title_date_id = self.ws.cell(4,9)
        cell_title_date_re = self.ws.cell(5,9)
        cell_val_ex_path   = self.ws.cell(2,10)
        cell_val_ex_date   = self.ws.cell(3,10)
        cell_val_id_date   = self.ws.cell(4,10)
        cell_val_re_date   = self.ws.cell(5,10)
        
        cell_title_b.font = Font(size=9, color="FFFFFF")
        cell_title_ex1.font = Font(size=9, color="FFFFFF")
        cell_title_ex2.font = Font(size=9, color="FFFFFF")
        cell_title_id.font = Font(size=9, color="FFFFFF")
        cell_title_re.font = Font(size=9, color="FFFFFF")
        cell_title_path.font = Font(size=9, color="FFFFFF")
        cell_title_date_ex.font = Font(size=9, color="FFFFFF")
        cell_title_date_id.font = Font(size=9, color="FFFFFF")
        cell_title_date_re.font = Font(size=9, color="FFFFFF")
        cell_val_ex_path.font = Font(size=9, color="FFFFFF")
        cell_val_ex_date.font = Font(size=9, color="FFFFFF")
        cell_val_id_date.font = Font(size=9, color="FFFFFF")
        cell_val_re_date.font = Font(size=9, color="FFFFFF")
        
        cell_title_b.alignment = Alignment(horizontal='center', vertical='center')
        cell_title_ex1.alignment = Alignment(horizontal='left', vertical='center')
        cell_title_ex2.alignment = Alignment(horizontal='left', vertical='center')
        cell_title_id.alignment = Alignment(horizontal='left', vertical='center')
        cell_title_re.alignment = Alignment(horizontal='left', vertical='center')
        cell_title_path.alignment = Alignment(horizontal='left', vertical='center')
        cell_title_date_ex.alignment = Alignment(horizontal='left', vertical='center')
        cell_title_date_id.alignment = Alignment(horizontal='left', vertical='center')
        cell_title_date_re.alignment = Alignment(horizontal='left', vertical='center')
        cell_val_ex_path.alignment = Alignment(horizontal='left', vertical='center')
        cell_val_ex_date.alignment = Alignment(horizontal='left', vertical='center')
        cell_val_id_date.alignment = Alignment(horizontal='left', vertical='center')
        cell_val_re_date.alignment = Alignment(horizontal='left', vertical='center')
        
        # fill_color = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        # cell_title_b.fill = fill_color
        # cell_title_ex1.fill = fill_color
        # cell_title_ex2.fill = fill_color
        # cell_title_id.fill = fill_color
        # cell_title_re.fill = fill_color
        # cell_title_path.fill = fill_color
        # cell_title_date_ex.fill = fill_color
        # cell_title_date_id.fill = fill_color
        # cell_title_date_re.fill = fill_color
        # cell_val_ex_path.fill = fill_color
        # cell_val_ex_date.fill = fill_color
        # cell_val_id_date.fill = fill_color
        # cell_val_re_date.fill = fill_color
        
        # cell_title_b.border = Border(left=Side(style=None),right=Side(style=None),top=Side(style=None),bottom=Side(style=None))
        # cell_title_ex1.border = Border(left=Side(style=None),right=Side(style=None),top=Side(style=None),bottom=Side(style=None))
        # cell_title_ex2.border = Border(left=Side(style=None),right=Side(style=None),top=Side(style=None),bottom=Side(style=None))
        # cell_title_id.border = Border(left=Side(style=None),right=Side(style=None),top=Side(style=None),bottom=Side(style=None))
        # cell_title_re.border = Border(left=Side(style=None),right=Side(style=None),top=Side(style=None),bottom=Side(style=None))
        # cell_title_path.border = Border(left=Side(style=None),right=Side(style=None),top=Side(style=None),bottom=Side(style=None))
        # cell_title_date_ex.border = Border(left=Side(style=None),right=Side(style=None),top=Side(style=None),bottom=Side(style=None))
        # cell_title_date_id.border = Border(left=Side(style=None),right=Side(style=None),top=Side(style=None),bottom=Side(style=None))
        # cell_title_date_re.border = Border(left=Side(style=None),right=Side(style=None),top=Side(style=None),bottom=Side(style=None))
        # cell_val_ex_path.border = Border(left=Side(style=None),right=Side(style=None),top=Side(style=None),bottom=Side(style=None))
        # cell_val_ex_date.border = Border(left=Side(style=None),right=Side(style=None),top=Side(style=None),bottom=Side(style=None))
        # cell_val_id_date.border = Border(left=Side(style=None),right=Side(style=None),top=Side(style=None),bottom=Side(style=None))
        # cell_val_re_date.border = Border(left=Side(style=None),right=Side(style=None),top=Side(style=None),bottom=Side(style=None))
        
        self.ws.merge_cells(start_row=1, start_column=8, end_row=1, end_column=10)
        
        cell_title_b.value = CEB.VALIDITY
        cell_title_b.number_format = f';;;"{CET.INIT_TITLE_VALIDITY}"'
        cell_title_ex1.value = CET.INIT_TITLE_EXTRACTION
        cell_title_ex2.value = CET.INIT_TITLE_EXTRACTION
        cell_title_id.value = CET.INIT_TITLE_IDENTIFICATION
        cell_title_re.value = CET.INIT_TITLE_REALISATION
        cell_title_path.value = CET.INIT_TITLE_PATH
        cell_title_date_ex.value = CET.INIT_TITLE_DATE
        cell_title_date_id.value = CET.INIT_TITLE_DATE
        cell_title_date_re.value = CET.INIT_TITLE_DATE

    def set_protection(self):
        # protect all the sheet
        protect_ws(self.ws)
        
        # unprotect certain part
        cell_val_ex = self.ws.cell(8,2)
        unprotect_cell(cell_val_ex)
        cell_val_ex = self.ws.cell(8,6)
        unprotect_cell(cell_val_ex)
        
        unprotect_range(self.ws,11,25,1,11)
        unprotect_range(self.ws,1,100,12,100)
    
if __name__ == "__main__":
    import openpyxl
    from Styles import StyleCell
    
    path = r"D:\One Drive\OneDrive\Bureau\Classeur1.xlsx"
    workbook = openpyxl.load_workbook(path)
    StyleCell(workbook)
    ws = workbook.active
    InitSheet(ws)
    workbook.save(path)
    
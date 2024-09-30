import Const as C
import Const_Balise_Excel as BE
from openpyxl.worksheet.datavalidation import DataValidation
from Tools import clear_zone
from Protection import set_protection

class WriteExtraction():
    def __init__(self,sheet,extraction,markers):
        self.extraction = extraction
        self.fo = self.extraction.financial_operations
        self.mks = markers
        self.ws = sheet
        self.supress_old_ext()
        self.write()
        self.up_mrks()
        self.set_cells_format()
        self.write_validation()
        self.write_date()
        self.write_sold()
        set_protection(self.ws,self.mks)

    def supress_old_ext(self):
        clear_zone(self.ws,self.mks.B_DETAIL_LINE_ST,self.ws.max_row+1,self.mks.B_ID_C1_COL,self.mks.B_REAL_COL_ED)
        
    def write(self):
        for row,fo in enumerate(self.fo,start=self.mks.B_DETAIL_LINE_ST):
            self.ws.cell(row=row,column=self.mks.B_EXT_NAME_COL).value=fo.name
            self.ws.cell(row=row,column=self.mks.B_EXT_VALUE_COL).value=fo.value
        self.mks.B_ext_ed_line = row + C.EXT_LINE_ADD
    
    def up_mrks(self):
        self.mks.b_ext_sold_ed_est = [self.mks.B_ext_ed_line+1,self.mks.B_ID_C2_COL]
        self.mks.b_ext_sold_ed_true = [self.mks.B_ext_ed_line+2,self.mks.B_ID_C2_COL]
        self.mks.b_ext_verif = [self.mks.b_ext_sold_ed_est[0],self.mks.B_EXT_NAME_COL]
    
    def set_cells_format(self):
        dv1 = DataValidation(type="list", formula1=f"={C.TOOL_SHEET_NAME}!$A:$A")
        dv2 = DataValidation(type="list", formula1=f"={C.TOOL_SHEET_NAME}!$B:$B")
        self.ws.add_data_validation(dv2)
        self.ws.add_data_validation(dv1)
        for row in range(self.mks.B_DETAIL_LINE_ST,self.mks.B_ext_ed_line+1):
            # color
            self.ws.cell(row=row,column=self.mks.B_EXT_NAME_COL).style = "Normal"
            self.ws.cell(row=row,column=self.mks.B_EXT_VALUE_COL).style = "Note"
            cell_cat1 = self.ws.cell(row=row,column=self.mks.B_ID_C1_COL)
            cell_cat2 =  self.ws.cell(row=row,column=self.mks.B_ID_C2_COL)
            cell_cat1.style = "style_cat1"
            cell_cat2.style = "style_cat2"
            # Data validation
            dv1.add(cell_cat1)
            dv2.add(cell_cat2)
    
    def write_validation(self):
        # estimated
        self.ws.cell(row=self.mks.b_ext_sold_ed_est[0],column=self.mks.b_ext_sold_ed_est[1]).value = self.extraction.sold_est
        cell_txt_est = self.ws.cell(row=self.mks.b_ext_sold_ed_est[0],column=self.mks.b_ext_sold_ed_est[1]-1)
        cell_txt_est.value = BE.EXT_VAL
        cell_txt_est.number_format  = f';;;"{C.EXT_VALID_TXT_SOLD_EST}"'
        # validation
        cell_val = self.ws.cell(row=self.mks.b_ext_verif[0],column=self.mks.b_ext_verif[1])
        cell_val.value = self.extraction.validation
        cell_val.style = "Good"
    
    def write_date(self):
        date = f"{self.extraction.date_st} -> {self.extraction.date_ed}"       
        self.ws.cell(row=self.mks.B_BILAN_DATE[0],column=self.mks.B_BILAN_DATE[1]).value = date

    def write_sold(self):
        self.ws.cell(row=self.mks.B_BILAN_ST_SOLD[0],column=self.mks.B_BILAN_ST_SOLD[1]).value = self.extraction.sold_st
        self.ws.cell(row=self.mks.B_BILAN_ED_SOLDRE[0],column=self.mks.B_BILAN_ED_SOLDRE[1]).value = self.extraction.sold_ed
    
    
    def is_valid(self):
        return True
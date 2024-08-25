import CONST as C
from Extract_from_pdf import PDFExtraction
from Tools import protect_ws,unprotect_range,unprotect_cell


class WriteFo:
    def __init__(self,fo,sheet,S):
        self.ws = sheet
        self.fo = fo
        self.month_start_row = C.MONTH_ROW_TITLE+C.START_MONTH[0]
        self.month_end_row = self.month_start_row + len(self.fo)+C.NB_MORE_FO_MONTH
        
        self.styles = S
        self.init_excel_month()
        self.write_fo()
        self.protect_sheet()
        
        

    def init_excel_month(self):
        self.init_cell_size_month()
        self.init_color_month()
        self.init_alignement_month()
        # self.init_boder_month()
        self.init_value_month()
        
    def init_cell_size_month(self):
        # column width
        self.ws.column_dimensions['A'].width = 34
        self.ws.column_dimensions['B'].width = 16
        self.ws.column_dimensions['C'].width = 26
        self.ws.column_dimensions['D'].width = 11
        self.ws.column_dimensions['E'].width = 2
        self.ws.column_dimensions['F'].width = 20
        self.ws.column_dimensions['G'].width = 20
        self.ws.column_dimensions['H'].width = 20
        self.ws.column_dimensions['I'].width = 20
        self.ws.column_dimensions['J'].width = 15
        self.ws.column_dimensions['K'].width = 15
        
        # merge
        self.ws.merge_cells(start_row=C.START_MONTH[0], start_column=C.START_MONTH[1],
                            end_row=C.START_MONTH[0]-1+C.MONTH_ROW_TITLE, end_column=C.START_MONTH[1]+3)
        
        self.ws.merge_cells(start_row=self.month_end_row , start_column=C.START_MONTH[1]+2,
                            end_row=self.month_end_row+1, end_column=C.START_MONTH[1]+2)
        self.ws.merge_cells(start_row=self.month_end_row, start_column=C.START_MONTH[1]+3,
                            end_row=self.month_end_row+1, end_column=C.START_MONTH[1]+3)
    def init_color_month(self):
        self.ws.cell(row=C.START_MONTH[0],column=C.START_MONTH[1]).style='Accent1'
        
        self.ws.cell(row=self.month_end_row,column=C.START_MONTH[1]).style='Normal'
        self.ws.cell(row=self.month_end_row,column=C.START_MONTH[1]+1).style='Input'
        self.ws.cell(row=self.month_end_row,column=C.START_MONTH[1]+2).style='Neutral'
        self.ws.cell(row=self.month_end_row,column=C.START_MONTH[1]+3).style='Neutral'
        
        self.ws.cell(row=self.month_end_row+1,column=C.START_MONTH[1]).style='Output'
        self.ws.cell(row=self.month_end_row+1,column=C.START_MONTH[1]+1).style='Output'
        for row in range(self.month_start_row,self.month_end_row):
            self.ws.cell(row=row,column=1).style = self.styles.style_cat1
            self.ws.cell(row=row,column=2).style = self.styles.style_cat2
            self.ws.cell(row=row,column=3).style = "Normal"
            self.ws.cell(row=row,column=4).style = "Note"
    def init_alignement_month(self):
        self.ws.cell(row=C.START_MONTH[0],column=C.START_MONTH[1]).alignment=C.ALIGN_CENTER
        self.ws.cell(row=self.month_end_row,column=C.START_MONTH[1]+2).alignment=C.ALIGN_CENTER
        self.ws.cell(row=self.month_end_row,column=C.START_MONTH[1]+3).alignment=C.ALIGN_CENTER
    # def init_boder_month(self):
    def init_value_month(self):
        self.ws.cell(row=C.START_MONTH[0],column=C.START_MONTH[1]).value=f'{C.BALISE_NEW_MONTH}{self.ws.title}'
        self.ws.cell(row=self.month_end_row,column=C.START_MONTH[1]).value = C.MONTH_SOLD_BALISE
        self.ws.cell(row=self.month_end_row,column=C.START_MONTH[1]+2).value = C.MONTH_VERIF
        self.ws.cell(row=self.month_end_row+1,column=C.START_MONTH[1]).value = C.MONTH_SOLD_EST

    def write_fo(self):
        for row,fo in enumerate(self.fo,start=self.month_start_row):
            self.ws.cell(row=row,column=C.START_MONTH[1]+2).value=fo.name
            self.ws.cell(row=row,column=C.START_MONTH[1]+3).value=fo.value
            
    def protect_sheet(self):
        protect_ws(self.ws)
        unprotect_range(self.ws,f'A{self.month_start_row}:D{self.month_end_row-1}')
        unprotect_cell(self.ws,f'B{self.month_end_row}')
        


        
if __name__ == '__main__':
    pass
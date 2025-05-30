"Read the PADf and extract information"
import pdfplumber
from Ext_LPB_2024 import ExtractionLbp2024
from tkinter import messagebox
from Const import EXT_VALID_ERR_UNIT
from Tools import custom_sort_fincancial_operation

class PDFExtraction():
    def __init__(self,file_path):
        self.file_path = file_path
        self.text = ''
        
        self.financial_operations = None
        self.month,self.year = None,None
        self.sold_st,self.sold_ed = None,None
        
        self.read_pdf()
        self.extract()
        self.add_and_cut()
        self.sort()
        self.lenght = len(self.financial_operations)
        self.validation_extraction()
        
    def read_pdf(self):
        try:
            with pdfplumber.open(self.file_path) as pdf:
                for page in pdf.pages:
                    self.text += page.extract_text()
        except FileNotFoundError:
            messagebox.showerror("Error", "Invalid PDF Path")
            
    def extract(self):
        self.financial_operations,self.month,self.year,self.sold_st,self.sold_ed,self.date_st,self.date_ed,self.tot_op = ExtractionLbp2024(self.text).get_value()
    
    def print(self):
        print(f'Date: {self.month} {self.year}')
        print(f'Sold Start: {self.sold_st}')
        print(f'Sold End: {self.sold_ed}')
        for op in self.financial_operations:
            op.print()
        print(f'Total: {self.tot_op}')

    def add_and_cut(self):
        'manage the financial operations to group them'
        for op in self.financial_operations:
            for opp in self.financial_operations:
                if op.num != opp.num and op.name == opp.name:
                    op.value = round(op.value + opp.value,3)
                    self.financial_operations.remove(opp)
        
    def sort(self):
        self.financial_operations = sorted(self.financial_operations, key=custom_sort_fincancial_operation)
    def validation_extraction(self):
        self.sold_est = self.sold_st + self.tot_op
        self.sold_diff = self.sold_est-self.sold_ed
        if self.sold_diff > EXT_VALID_ERR_UNIT:
            self.validation = False
        else:
            self.validation = True

if __name__ == '__main__':
    PATH = r"D:\One Drive\OneDrive\Documents\.Perso\[5] Finance\Compte\Releves\2024\releve_CCP2142397R038_20240126.pdf"
    extraction = PDFExtraction(PATH)
    extraction.print()




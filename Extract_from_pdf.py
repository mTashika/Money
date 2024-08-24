import pdfplumber
from Extraction_LPB_2024 import ExtractionLbp2024

class PDFExtraction():
    def __init__(self,file_path):
        self.file_path = file_path
        self.text = ''
        
        self.financial_operations = None
        self.month,self.year = None,None
        
        self.read_pdf()
        self.extract()
        self.add_and_cut()
        
    def read_pdf(self):
        with pdfplumber.open(self.file_path) as pdf:
            for page in pdf.pages:
                self.text += page.extract_text()
        # print(self.text)
                
    def extract(self):
        self.financial_operations,self.month,self.year = ExtractionLbp2024(self.text).get_value()
    
    def print(self):
        print(f'Date: {self.month} {self.year}')
        for op in self.financial_operations:
            op.print()
        
            
        
    def add_and_cut(self):
        'manage the financial operations to group them'
        for op in self.financial_operations:
            for opp in self.financial_operations:
                if op.num != opp.num and op.name == opp.name:
                    op.value = round(op.value + opp.value,3)
                    self.financial_operations.remove(opp)

if __name__ == '__main__':
    PATH = r"D:\One Drive\OneDrive\Bureau\releve_CCP2142397R038_20240328.pdf"
    extraction = PDFExtraction(PATH)
    extraction.print()
    
    




""" Read the PAD and extract information """
from Ext_LPB_2024 import ExtractionLbp2024
from Const import EXT_VALID_ERR_UNIT

class PDFExtraction():
    def __init__(self,pdf_path,flg_account=None):
        self.pdf_path = pdf_path
        self.flg_account = flg_account
        
        self.financial_operations = None
        self.month,self.year = None,None
        self.sold_st,self.sold_ed = None,None

        self.extract()
        self.add_and_cut()
        self.sort()
        self.lenght = len(self.financial_operations)
        self.validation_extraction()
            
    def extract(self):
        """Extract the data from the pdf using the right extraction method"""
        self.financial_operations,self.month,self.year,self.sold_st,self.sold_ed,self.date_st,self.date_ed,self.tot_op = ExtractionLbp2024(self.pdf_path,self.flg_account).get_value()
    
    def print(self):
        print(f'Date: {self.month} {self.year}')
        print(f'Sold Start: {self.sold_st}')
        print(f'Sold End: {self.sold_ed}')
        for op in self.financial_operations:
            op.print()
        print(f'Total: {self.tot_op}')

    def add_and_cut(self):
        """Manage the financial operations to group them"""
        for op in self.financial_operations:
            for opp in self.financial_operations:
                if op.num != opp.num and op.name == opp.name: # if the current financial operation "opp" has the same same but is not "op"
                    op.value = op.value + opp.value
                    self.financial_operations.remove(opp)
            op.value = round(op.value,2)
        
    def sort(self):
        self.financial_operations = sorted(self.financial_operations, key=custom_sort_fincancial_operation)

    def validation_extraction(self):
        """ Valid or not the estimation"""
        self.sold_est = self.sold_st + self.tot_op
        self.sold_diff = self.sold_est-self.sold_ed
        if self.sold_diff > EXT_VALID_ERR_UNIT:
            self.validation = False
        else:
            self.validation = True

def custom_sort_fincancial_operation(fo):
    """ On renvoie un tuple : le premier élément détermine l'ordre (positif ou négatif),
    le second élément est la valeur elle-même pour le tri interne"
    """
    value = fo.value
    return (1 if value < 0 else 0, -value if value >= 0 else value)

if __name__ == '__main__':
    PATH = r"C:/Users/mcast/OneDrive/Bureau/releve_CCP2142397R038_20241227.pdf"
    extraction = PDFExtraction(PATH,"FR75 2004 1010 0721 4239 7R03 830")
    extraction.print()




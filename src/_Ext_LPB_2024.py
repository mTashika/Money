"Fonction et classe d'extraction des données depuis un pdf la banque postal (marche en 2024)"
import re
from Tools import FinancialOperation,remove_accent
import Const as C

BALISE_CREDIT = ['VIREMENT DE','VIREMENT INSTANTANE DE','CREDIT',"REMISE COMMERCIALE D'AGIOS"]
BALISE_BUY    = ['ACHAT CB','PRELEVEMENT DE ','COTISATION TRIMESTRIELLE',
                'VIREMENT INSTANTANE A','VIREMENT POUR','MINIMUM FORFAITAIRE TRIMESTRIEL',
                'AVANTAGE SEUIL DE NON-PERCEPTION','COMMISSION','RETRAIT']
BALISE_NONE = ['INTERETS ACQUIS POUR']

BALISE_ST_CCP = 'Compte Courant Postal'
BALISE_ED_CCP = 'Comptes'
BUY_VALUE = C.EXT_BUY_VALUE
CRED_VALUE = C.EXT_CRED_VALUE
NONE_VALUE = C.EXT_NONE_VALUE
INTER_LINE = ' --- '

def is_line_transaction_one(line):
        pattern = r"\b\d{2}/\d{2}\b\s+[A-Z\s]+\s+\d{1,3}(?: \d{3})*,\d{2}"
        if re.match(pattern, line):
            return True
        else:
            return False

class ExtractionLbp2024:
    def __init__(self,text):
        self.text = text
        self.financial_operation = []
        self.init_balise()
        self.change_text()
        self.get_matching_lines()
        self.extract_lpb_2024()
        self.set_sign()
        self.get_date()
        self.get_sold()
    
    def init_balise(self):
        # Balise for the Positive (+) transaction
        self.B_CREDIT = BALISE_CREDIT
        # Balise for the negative (-) transaction
        self.B_BUY = BALISE_BUY
        self.BALISE_NONE = BALISE_NONE
        # Balise for the date
        self.B_DATE = 'Situation de vos comptes\nau '
        " Balise for the start and end Sold"
        self.B_SOLD_ST = 'Ancien solde au'
        self.B_SOLD_ED = 'Nouveau solde au'
        
    def change_text(self):
        self.text = self.text.replace('Ø', "e")
        self.text = self.text.replace('(cid:201)','E')
        self.text = remove_accent(self.text)
        self.text = self.text.replace('ß','u')
        self.text = self.text.replace('*',' ')
        self.text = self.text.replace('’', "'")
    
    def get_value(self):
        return self.financial_operation,self.month,self.year,self.sold_st,self.sold_ed,self.date_st,self.date_ed,self.tot_op
        
        
    def get_balise(self,line):
        'Return the balise name of the line'
        for balise_buy in self.B_BUY:
                if balise_buy in line:
                    return balise_buy
        for balise_cred in self.B_CREDIT:
                if balise_cred in line:
                    return balise_cred
        for balise_none in self.BALISE_NONE:
                if balise_none in line:
                    return balise_none
        raise Exception(f'Error\nLine: {line}\nLine does not have any balise') # HERE
    
    def get_balise_type(self,balise_name):
        if balise_name in self.B_BUY:
            return BUY_VALUE
        elif balise_name in self.B_CREDIT:
            return CRED_VALUE
        elif balise_name in self.BALISE_NONE:
            return NONE_VALUE
        else:
            raise Exception('Invalide balise')
        
    def get_pattern(self,balise_name,balise_type):
        'return the pattern for the specifique balise to take the description and the amount'
        if balise_type == NONE_VALUE:
            return None,None
        
        pattern_date = rf"\d{{2}}\.\d{{2}}\.\d{{2}}"
        pattern_special_chars = rf'[^\w\s.]'
        pattern_amount_detection = r'(?<!\d)(?:\d{1,3}(?: \d{3})*|\d+)(?:,\d{2})(?!\d)'
        pattern_name = rf'{re.escape(balise_name)}\s*([A-Za-z0-9.\s]*?)(?=\s*{pattern_amount_detection}|{pattern_special_chars}|{INTER_LINE}|{pattern_date})'

        pattern_amount = rf'(?<=\s)(\d+(?: \d{{3}})*\,\d{{2}})(?=\s*{re.escape(INTER_LINE)})'

        if balise_name == 'VIREMENT INSTANTANE DE':
            pattern_name = rf'{re.escape(INTER_LINE)}\s*([A-Za-z]+\s+[A-Za-z]+)'
        return pattern_name,pattern_amount
    
    def update_description(self,name,balise):
        'specific case for the description (name)'
        if balise == 'RETRAIT':
            name = balise +' '+name
        if name == '' or name == ' ' or name == '  ':
            name = balise
        return name
    
    def get_matching_lines(self):
        lines = self.text.splitlines()
        self.matching_lines = []
        f_ccp = False
        for i in range(len(lines)):
            if BALISE_ST_CCP in lines[i]:
                f_ccp = True
            if BALISE_ED_CCP in lines[i]:
                f_ccp = False
            if f_ccp:
                if re.match(r"^\d{2}/\d{2} ", lines[i]):
                    l = lines[i]
                    l+= INTER_LINE
                    # Ajouter la ligne suivante si elle existe
                    if i + 1 < len(lines):
                        if not is_line_transaction_one(lines[i+1]):
                            l += lines[i+1]
                    self.matching_lines.append(l)
        # for l in self.matching_lines:
        #     print(l)
    
    def get_date(self):
        pattern_date = rf'{re.escape(self.B_DATE)}\s*\d{{2}}\s*([a-zA-Z]+)\s(\d{{4}})'
        match = re.search(pattern_date,self.text)
        self.month,self.year = match.group(1),match.group(2)
        pattern_date_st = rf"{re.escape(self.B_SOLD_ST)}\s(\d{{2}})/(\d{{2}})/\d{{4}}"
        match_st = re.search(pattern_date_st,self.text)
        self.date_st = f"{match_st.group(1)}/{match_st.group(2)}"
        pattern_date_ed = rf"{re.escape(self.B_SOLD_ED)}\s(\d{{2}})/(\d{{2}})/\d{{4}}"
        match_ed = re.search(pattern_date_ed,self.text)
        self.date_ed = f"{match_ed.group(1)}/{match_ed.group(2)}"

    def get_sold(self):
        pattern_sold_st = rf'{re.escape(self.B_SOLD_ST)}\s+\d{{2}}/\d{{2}}/\d{{4}}\s+(\d{{1,3}}(?:[ \d]*\d)?\,\d{{2}})'
        match_st = re.search(pattern_sold_st,self.text)
        pattern_sold_ed = rf'{re.escape(self.B_SOLD_ED)}\s+\d{{2}}/\d{{2}}/\d{{4}}\s+(\d{{1,3}}(?:[ \d]*\d)?\,\d{{2}})'
        match_ed = re.search(pattern_sold_ed,self.text)
        self.sold_st = round(float(match_st.group(1).replace(',', '.').replace(' ', '')),2)
        self.sold_ed = round(float(match_ed.group(1).replace(',', '.').replace(' ', '')),2)
        
        self.tot_op = 0
        for op in self.financial_operation:
            self.tot_op+=op.value
        self.tot_op = round(self.tot_op,2)
            
        if not round(self.sold_st + self.tot_op,0) == round(self.sold_ed,0):
            if round(-self.sold_st + self.tot_op,0) == round(self.sold_ed,0):
                self.sold_st = -self.sold_st
            elif round(self.sold_st + self.tot_op,0) == round(-self.sold_ed,0):
                self.sold_ed = -self.sold_ed
            elif round(-self.sold_st + self.tot_op,0) == round(-self.sold_ed,0):
                self.sold_st = -self.sold_st
                self.sold_ed = -self.sold_ed
    
    def extract_lpb_2024(self):
        '''
        Extract information from a pdf with "la banque postal" convention
        '''
        for matching_line in self.matching_lines:
            balise_name = self.get_balise(matching_line)
            balise_type = self.get_balise_type(balise_name)
            pattern_name,pattern_amount = self.get_pattern(balise_name,balise_type)
            if pattern_name is not None and pattern_amount is not None:
                match_name = re.search(pattern_name, matching_line)
                match_amout = re.search(pattern_amount, matching_line)
                
                if match_name and match_amout:
                    name = self.update_description(match_name.group(1),balise_name)
                    value = round(float(match_amout.group(1).replace(',', '.').replace(' ', '')),2)
                    op = FinancialOperation(num=len(self.financial_operation)+1,
                                            type=balise_type,
                                            name=name,
                                            value=value
                                            )
                    self.financial_operation.append(op)
                else:
                    raise Exception (f'No match for matching_line: {matching_line}\n{match_name}\n{match_amout}')
            else:
                print(f'None Balise : {balise_name}')
            
    def set_sign(self):
        for fo in self.financial_operation:
            if fo.type == BUY_VALUE:
                fo.value = -fo.value
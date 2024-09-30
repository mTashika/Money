"Fonction et classe d'extraction des données depuis un pdf la banque postal (marche en 2024)"
import re
from Tools import FinancialOperation
import Const as C

BALISE_CREDIT = ['VIREMENT DE ','VIREMENT INSTANTANE DE ','CREDIT']
BALISE_BUY    = ['ACHAT CB ','PRELEVEMENT DE ','COTISATION TRIMESTRIELLE DE ',
                'VIREMENT INSTANTANE A ','VIREMENT POUR']

BUY_VALUE = C.EXT_BUY_VALUE
CRED_VALUE = C.EXT_CRED_VALUE
INTER_LINE = ' --- '

class ExtractionLbp2024:
    def __init__(self,text):
        self.text = text
        self.FinOp = []
        self.init_balise()
        
    def get_value(self):
        self.get_matching_lines()
        self.extract_lpb_2024()
        self.set_sign()
        self.get_date()
        self.get_sold()
        return self.FinOp,self.month,self.year,self.sold_st,self.sold_ed,self.date_st,self.date_ed
        
    def init_balise(self):
        # Balise for the Positive (+) transaction
        self.B_CREDIT = BALISE_CREDIT
        # Balise for the negative (-) transaction
        self.B_BUY = BALISE_BUY
        # Balise for the date
        self.B_DATE = 'Situation de vos comptes\nau '
        " Balise for the start and end Sold"
        self.B_SOLD_ST = 'Ancien solde au'
        self.B_SOLD_ED = 'Nouveau solde au'
        
    def get_balise(self,line):
        'Return the balise name of the line'
        for balise_buy in self.B_BUY:
                if balise_buy in line:
                    return balise_buy
        for balise_cred in self.B_CREDIT:
                if balise_cred in line:
                    return balise_cred
        raise Exception(f'Error\nLine: {line}\nLine does not have any balise')
    
    def get_balise_type(self,balise_name):
        if balise_name in self.B_BUY:
            return BUY_VALUE
        elif balise_name in self.B_CREDIT:
            return CRED_VALUE
        else:
            raise Exception('Invalide balise')
        
    def get_pattern(self,balise_name):
        'return the pattern for the specifique balise to take the description and the amount'
        pattern_date = rf"\d{{2}}\.\d{{2}}\.\d{{2}}"
        pattern_special_chars = rf'[^\w\s.]'
        pattern_amount_g = rf'\d{{1,3}}(?: \d{{3}})*,\d{{2}}'
        
        pattern_name   = rf'{re.escape(balise_name)}\s*([A-Za-z\s.]+?)(?=\d|{pattern_special_chars}|{INTER_LINE}|{pattern_amount_g}|{pattern_date})'
        pattern_amount = rf'({pattern_amount_g})(?={INTER_LINE})'

        if balise_name == 'VIREMENT INSTANTANE DE ':
            pattern_name = rf'{re.escape(INTER_LINE)}\s*([A-Za-z]+\s+[A-Za-z]+)'
        return pattern_name,pattern_amount
    
    def update_description(self,name):
        'specific case for the description (name)'
        return name
    
    def get_matching_lines(self):
        lines = self.text.splitlines()
        self.matching_lines = []
        for i in range(len(lines)):
            if re.match(r"^\d{2}/\d{2} ", lines[i]):
                l = lines[i]
                l+= INTER_LINE
                # Ajouter la ligne suivante si elle existe
                if i + 1 < len(lines):
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
        
    
    def extract_lpb_2024(self):
        '''
        Extract information from a pdf with "la banque postal" convention
        '''
        for matching_line in self.matching_lines:
            balise_name = self.get_balise(matching_line)
            balise_type = self.get_balise_type(balise_name)
            pattern_name,pattern_amount = self.get_pattern(balise_name)
            match_name = re.search(pattern_name, matching_line)
            match_amout = re.search(pattern_amount, matching_line)
            
            if match_name and match_amout:
                name = self.update_description(match_name.group(1))
                value = round(float(match_amout.group(1).replace(',', '.').replace(' ', '')),2)
                op = FinancialOperation(num=len(self.FinOp)+1,
                                        type=balise_type,
                                        name=name,
                                        value=value
                                        )
                self.FinOp.append(op)
            else:
                raise Exception(f'No match for matching_line: {matching_line}\n{match_name}\n{match_amout}')
            
    def set_sign(self):
        for fo in self.FinOp:
            if fo.type == BUY_VALUE:
                fo.value = -fo.value
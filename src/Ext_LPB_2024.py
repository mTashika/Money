import fitz
import re
from Tools import FinancialOperation,month_number_to_name
YTOL = 3
XMIN = 350
XENDNEGATIF = 470
YENDPAGE = 750
YSTARTPAGE = 160
XTOTAL = 290

B_REMOVE_TXT = {"ACHAT CB"}
B_REMOVE_LINE = {"CARTE NUMERO","CARTE NO"}
B_TAKE_NAME_NEXT_LINE = {"VIREMENT INSTANTANE","VIREMENT POUR"}
B_FIRST_LINE_IS_ENOUGH = {"VIREMENT DE","PRELEVEMENT DE","ACHAT CB"}

class ExtractionLbp2024:
    def __init__(self,PATH,flg_account=None):
        self.path = PATH
        self.FLAG_ACCOUNT = "Compte Courant Postal" if flg_account is None else flg_account
        self.tab_lines = []
        self.ref_start=r'Ancien solde au'
        self.ref_end=r'Nouveau solde au'
        self.id_st,self.id_ed = None,None
        self.Fo_lines = []
        self.fo_txts = []
        self.ret_tab = []
        self.financial_operation = []
        self.month = ""
        self.year = ""
        self.sold_st = ""
        self.sold_ed = ""
        self.date_st = ""
        self.date_ed = ""
        self.tot_op = 0

        self.extract_text_and_positions()
        self.separate_in_lines()
        self.find_start_end()
        self.txt_line_opti = self.tab_lines[self.id_st:self.id_ed+1]
        self.group_FO()
        self.get_text_fo_lines()
        self.get_value_and_description()
        self.build_FO()
        self.get_date()
        self.get_sold()

        del self.tab_lines,self.tab_word_position,self.Fo_lines,self.fo_txts

    def extract_text_and_positions(self):
        doc = fitz.open(self.path)  # Open the PDF file
        word_positions = []  # List to hold words and their positions
        all_txt = ""
        
        for page_num, page in enumerate(doc):
            words = page.get_text("words")
            for word in words:
                x0, y0, x1, y1, text, _, _, _ = word  # Unpack the word and coordinates
                # Append the word along with its position and the page number for sorting
                word_positions.append([x0, y0, text, page_num])
        
        # Sort word positions by page number, then by y0 (top-to-bottom), then by x0 (left-to-right)
        word_positions.sort(key=lambda w: (w[3], w[1], w[0]))
        
        # Create the text in the correct reading order
        all_txt = " ".join([word[2] for word in word_positions])
        
        self.tab_word_position = word_positions
        return all_txt.strip() 

    def separate_in_lines(self):
        tol = YTOL
        idw_c = 0
        tmp_line = [self.tab_word_position[idw_c]]
        l_c = self.tab_word_position[idw_c][1]
        idw_c += 1
        while idw_c < len(self.tab_word_position):
            w = self.tab_word_position[idw_c]
            if "Ancien" in w:
                a = 1

            if abs(w[1] - l_c) <= tol:
                tmp_line.append(w)
            else:
                tmp_line.sort(key=lambda w: w[0])
                self.tab_lines.append(tmp_line)
                tmp_line = [w]
                l_c = w[1]
            idw_c += 1
        self.tab_lines.append(tmp_line)

    def find_start_end(self):
        """"Search for FLAG_ACCOUNT, once found, search for start and end references and return the line number"""
        flg_account = False
        for i,line in enumerate(self.tab_lines):
            txt = ""
            for w in line:
                txt += w[2] + " "
            if not flg_account and self.FLAG_ACCOUNT in txt: 
                flg_account = True # FLAG_ACCOUNT found
            if flg_account:
                if re.match(self.ref_start, txt, re.IGNORECASE) and self.id_st is None:
                    self.id_st = i
                elif re.match(self.ref_end, txt, re.IGNORECASE) and self.id_ed is None:
                    self.id_ed = i
                    break

    def group_FO(self):
        """Group the lines for each Financial Operation"""
        tmp_l = []
        for line in self.txt_line_opti:
            if line[0][1] > YENDPAGE or line[0][1] < YSTARTPAGE:
                continue
            if line[0][0] > XTOTAL: # if case total
                self.Fo_lines.append(tmp_l)
                tmp_l = []
                break
            if bool(re.match(r"^\d{2}/\d{2}$", line[0][2])): # fist word with good syntax
                if tmp_l != []: # case if new Fo line (save the previous and init for the next)
                    self.Fo_lines.append(tmp_l)
                    tmp_l = []
                if line[-1][0]>=XMIN: # last word with good x (on the right side of the page) -> line in fo
                    tmp_l.append(line)
            elif tmp_l != []:
                tmp_l.append(line)

    def get_text_fo_lines(self):
        for lines in self.Fo_lines:
            txt = []
            val_txt = ""
            for i,line in enumerate(lines): # the first word is XX/XX
                tmp_txt = ""
                for j,w in enumerate(line):
                    if i == 0:
                        if w[0] > XMIN: # it is the value
                            if val_txt == "":
                                if w[0] < XENDNEGATIF:
                                    val_txt = "-"
                            val_txt += w[2]
                        elif j != 0:
                            tmp_txt += w[2] + " "
                    else:
                        tmp_txt += w[2] + " "
                txt.append(tmp_txt.strip())
            txt.append(val_txt)
            self.fo_txts.append(txt)

    def get_value_and_description(self):
        val = 0
        # Get Description #
        desc = self.get_fo_txt_description()
        # Get value and final Tab #
        for i,fo_txt in enumerate(self.fo_txts):
            val = round(float(fo_txt[-1].replace(",",".")),2)
            self.ret_tab.append([desc[i],val])
    
    def get_fo_txt_description(self):
        """Get the financial operation description and associated value"""
        patt_rmv_date = r"\b\d{2}\.\d{2}\.\d{2}\b"  # Matches 'XX.XX.XX'
        ret_txt = []
        for fo_txt in self.fo_txts:
            ## First Line ##
            # Remove date in the first line
            fo_txt[0] = re.sub(patt_rmv_date, "", fo_txt[0]).strip()
            # Remove number
            fo_txt[0] = re.sub(r"[^a-zA-Z ]", "", fo_txt[0])
            ## Second Line ##
            # Remove second line if special Case
            for B in B_REMOVE_LINE:
                if B in fo_txt[1]:
                    del fo_txt[1]
            ## Build final txt ##
            desc = None
            # First line is enough #
            for B in B_FIRST_LINE_IS_ENOUGH:
                if B in fo_txt[0]:
                    desc = fo_txt[0]
            # Take the name in the next section #
            for B in B_TAKE_NAME_NEXT_LINE:
                if B in fo_txt[0]:
                    name = " ".join(re.findall(r'\b[a-zA-Z]{3,}\b', fo_txt[1])[:2]) 
                    desc = fo_txt[0] + " " + name
            # Default Case #
            if desc is None:
                desc = " ".join(fo_txt[:-1]) # by default take all the description
            # Remove Balise in final description
            for B in B_REMOVE_TXT:
                if B in desc:
                    desc = desc.replace(B,"")
            ret_txt.append(desc.strip()) # add the final description and the value
        return ret_txt
    
    def build_FO(self):
        for fo in self.ret_tab:
            op = FinancialOperation(num=len(self.financial_operation)+1,
                                                type=-1 if fo[1]< 0 else 1,
                                                name=fo[0],
                                                value=fo[1]
                                                )
            self.financial_operation.append(op)
        # Tot Operation #
        for op in self.financial_operation:
            self.tot_op+=op.value
        self.tot_op = round(self.tot_op,2)
    def get_value(self):
        return self.financial_operation,self.month,self.year,self.sold_st,self.sold_ed,self.date_st,self.date_ed,self.tot_op

    def get_sold(self):
        val_st_txt = ""
        val_ed_txt = ""
        # START #
        for w in self.txt_line_opti[0]:
            if w[0] > XMIN:
                if val_st_txt == "" and w[0] < XENDNEGATIF:
                    if w[0] > XENDNEGATIF:
                        val_st_txt = "-"
                val_st_txt +=  w[2]
        self.sold_st = round(float(val_st_txt.replace(",",".")),2)
        # END #
        for w in self.txt_line_opti[-1]:
            if w[0] > XMIN:
                if val_ed_txt == "" and w[0] < XENDNEGATIF:
                    if w[0] > XENDNEGATIF:
                        val_ed_txt = "-"
                val_ed_txt +=  w[2]
        self.sold_ed = round(float(val_ed_txt.replace(",",".")),2)

        if round(self.sold_st + self.tot_op) != round(self.sold_ed):
            print("Warning : self.sold_st + self.tot_op != self.sold_ed")

    def get_date(self):
        d = self.txt_line_opti[-1][3][2].split("/")
        self.month = month_number_to_name(d[1])
        self.year = d[2]
        self.date_st = self.txt_line_opti[0][3][2].split("/")[0] + "/" + self.txt_line_opti[0][3][2].split("/")[1]
        self.date_ed = self.txt_line_opti[-1][3][2].split("/")[1] + "/" + self.txt_line_opti[-1][3][2].split("/")[0]

if __name__ == "__main__":
    PATH = r"C:/Users/mcast/OneDrive/Bureau/releve_CCP2142397R038_20241227.pdf"
    a = ExtractionLbp2024(PATH)

    a=1
    

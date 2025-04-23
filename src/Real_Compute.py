import Const as C
from Tools import check_cell_value,evaluate_excel_formula

class DataRealisation():
    def __init__(self,cat1,cat2,fo_type,refund,name,value):
        self.cat1 = cat1 # nom de la cat 1 (peux etre None)
        self.cat2 = cat2
        self.fo_type = fo_type
        self.refund = refund
        self.name = name
        self.value = value # value in the pdf

        if self.cat2 is None:
            self.cat2 = ""
        if self.refund is None:
            self.refund = 0
        if self.name is None:
            self.name = ""
        if self.value is None:
            self.value = 0


class MonthRealisation():
    def __init__(self,sheet,markers,real_type):
        self.ws = sheet
        self.mks = markers
        self.real_type = real_type
        self.cats1=[]
        
        self.tot = 0 # sum of every charges
        self.tot_ex = 0 # sum of every negative charges
        self.tot_inc = 0 # sum of every positive charges=
        
        self.line_real = []
        self.extract_data_realisation()
        self.filter_lines()
        self.update_tot()
        self.cat_diff_signs()
        self.create_cats1()
        self.sort_cats1()
        # delete attribut
        del self.ws, self.mks
        

    def extract_data_realisation(self):
        """ Extract all the data for each financial operation """
        for i in range(self.mks.B_DETAIL_LINE_ST,self.mks.B_ext_ed_line+1):
            if check_cell_value(self.ws.cell(row=i, column=self.mks.B_ID_C1_COL).value) == 0:
                cat1 = self.ws.cell(row=i, column=self.mks.B_ID_C1_COL).value
                cat2 = self.ws.cell(row=i, column=self.mks.B_ID_C2_COL).value
                fo_type = self.ws.cell(row=i, column=self.mks.B_ID_TYPE_COL).value
                refund = self.ws.cell(row=i, column=self.mks.B_ID_REFUND_COL).value
                if check_cell_value(refund) == 0:
                    refund = evaluate_excel_formula(refund,self.ws)
                    if check_cell_value(refund) not in [1,2]:
                        refund = None
                name = self.ws.cell(row=i, column=self.mks.B_EXT_NAME_COL).value
                val = self.ws.cell(row=i, column=self.mks.B_EXT_VALUE_COL).value

                data_line = DataRealisation(cat1,cat2,fo_type,refund,name,val)
                self.line_real.append(data_line)
        self.sold_st_month_pdf = self.ws.cell(row=self.mks.B_BILAN_ST_SOLD[0], column=self.mks.B_BILAN_ST_SOLD[1]).value
    
    def filter_lines(self):
        if self.real_type == "Monthly":
            self.line_real = [l for l in self.line_real if l.fo_type == C.VALID_MONTHLY]
            for l in self.line_real:
                l.value += l.refund

        elif self.real_type == "Pdf":
            self.line_real = [l for l in self.line_real if (l.fo_type == C.VALID_MONTHLY or l.fo_type == C.VALID_DEFFERED)]

        elif self.real_type == "Intern":
            self.line_real = [l for l in self.line_real if (l.fo_type == C.VALID_INTERN)]

    def update_tot(self):
        """ Compute the total """
        for l in self.line_real:
            euro = l.value
            if isinstance(euro,(int,float)) and (check_cell_value(euro)==1 or check_cell_value(euro)==2):
                self.tot+=euro
                if euro > 0:
                    self.tot_inc+=euro
                else:
                    self.tot_ex+=euro
        
    def cat_diff_signs(self):
        """ Change the name of categories 1 that have both + and - value. """
        c1_done = [] # name of the categorie 1 done
        for l in self.line_real:
            c1 = l.cat1
            values = []
            if c1 not in c1_done and not c1 is None:
                # Get all the values of the Cat1
                for l in self.line_real:
                    if l.cat1 == c1: # Good Categorie 1
                        values.append(l.value)
                if any(x > 0 for x in values) and any(x < 0 for x in values): # if different sign 
                    nplus = c1 + " +"
                    nminus = c1 + " -"
                    for l in self.line_real:
                        if l.cat1 == c1: # Good Categorie 1
                            if l.value > 0:
                                l.cat1 = nplus
                            else:
                                l.cat1 = nminus
                    c1_done.append(nplus)
                    c1_done.append(nminus)
                c1_done.append(c1)

    def create_cats1(self):
        """ Create the Categories 1"""
        dic_cats1 = {}
        for l in self.line_real:
            if l.cat1 is not None:
                if l.cat1 not in dic_cats1:
                    dic_cats1[l.cat1] = [l]
                else:
                    dic_cats1[l.cat1].append(l)
        # create categories 
        for _,line_list in dic_cats1.items():
            self.cats1.append(Cat1(line_list,self.tot_ex,self.tot_inc))
    
    def sort_cats1(self):
        self.cats1 = sorted(self.cats1, key=lambda cat: ((cat.tot) < 0, cat.tot))


class Cat1():
    def __init__(self,line_real,exp_tot,inc_tot):
        self.inc_tot_all = inc_tot # sum of every positive charges except the 'Intern'
        self.exp_tot_all = exp_tot # sum of every negative charges except the 'Intern'
        self.real_line = line_real
        
        self.tot = 0 # sum of every charge in this categorie
        self.tot_p = 0 # pourcentage of the categorie compare to inc_tot_all or exp_tot_all
        self.cats2=[]
        self.name = self.real_line[0].cat1
        
        self.get_tot()
        self.create_cats2()
        self.nb_cat2 = len(self.cats2)
    def get_tot(self):
        """ Generate the tot and the % of the categorie (/sum pdf total - Intern). """
        # tot of the categorie in money unit
        for l in self.real_line:
            self.tot += l.value
        # tot of the categorie in pourcentage
        if self.tot > 0:
            self.tot_p = self.tot/self.inc_tot_all
        elif self.tot < 0:
            self.tot_p = self.tot/self.exp_tot_all
        else:
            self.tot_p = 0

    def create_cats2(self):
        """ Create the Categorie 2 """
        dic_cats2 = {} # Goup the same cat2 together
        for l in self.real_line:
            if l.cat2 not in dic_cats2:
                dic_cats2[l.cat2] = [l]
            else:
                dic_cats2[l.cat2].append(l)
        for _,fo_line in dic_cats2.items():
            self.cats2.append(Cat2(fo_line,self.tot))

class Cat2():
    def __init__(self,real_line,tot_cat1):
        self.real_line = real_line
        self.tot_cat1 = tot_cat1
        
        self.tot = 0
        self.tot_p = 0
        self.name = self.real_line[0].cat2
        self.get_tot()
        
    def get_tot(self):
        # tot
        for l in self.real_line:
            self.tot += l.value
        # tot_p
        if self.tot_cat1 !=0:
            self.tot_p = abs(self.tot/self.tot_cat1)
        else:
            self.tot_p = 0
        self.tot_p = round(self.tot_p,4)





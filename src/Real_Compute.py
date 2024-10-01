import Const as C
from Tools import check_cell_value

class DataRealisation():
    def __init__(self,cat1,cat2,name,value):
        self.cat1 = cat1
        self.cat2 = cat2
        self.name = name
        self.value = value
        
class MonthRealisation():
    def __init__(self,sheet,markers):
        self.ws = sheet
        self.mks = markers
        self.cats1=[]
        
        self.tot = 0
        self.tot_ex = 0
        self.tot_inc = 0
        self.tot_real = 0
        self.tot_ex_real = 0
        self.tot_inc_real = 0
        
        self.sold_st_month_pdf = 0
        self.tot_sold_real = 0
        
        self.d_realisation = []
        self.extract_data_realisation()
        self.update_tot()
        self.create_cats1()
        self.sort_cats1()

    def extract_data_realisation(self):
        for i in range(self.mks.B_DETAIL_LINE_ST,self.mks.B_ext_ed_line+1):
            if check_cell_value(self.ws.cell(row=i, column=self.mks.B_ID_C1_COL).value) == 0:
                cat1 = self.ws.cell(row=i, column=self.mks.B_ID_C1_COL).value
                cat2 = self.ws.cell(row=i, column=self.mks.B_ID_C2_COL).value
                name = self.ws.cell(row=i, column=self.mks.B_EXT_NAME_COL).value
                val = self.ws.cell(row=i, column=self.mks.B_EXT_VALUE_COL).value
                d = DataRealisation(cat1,cat2,name,val)
                self.d_realisation.append(d)
        self.sold_st_month_pdf = self.ws.cell(row=self.mks.B_BILAN_ST_SOLD[0], column=self.mks.B_BILAN_ST_SOLD[1]).value
        
    def update_tot(self):
        for val in self.d_realisation:
            euro = val.value
            ok_real = True if val.cat1 not in C.NOT_REAL_ACTION else False
            if isinstance(euro,(int,float)) and (check_cell_value(euro)==1 or check_cell_value(euro)==2):
                self.tot+=euro
                if ok_real:
                    self.tot_real+=euro
                if euro > 0:
                    self.tot_inc+=euro
                    if ok_real:
                        self.tot_inc_real+=euro
                else:
                    self.tot_ex+=euro
                    if ok_real:
                        self.tot_ex_real+=euro
        self.tot_sold_real = self.sold_st_month_pdf + self.tot_real
        
    def create_cats1(self):
        dic_cats1 = {}
        for d_real in self.d_realisation:
            if d_real.cat1 is not None:
                if d_real.cat1 not in dic_cats1:
                    dic_cats1[d_real.cat1] = [d_real]
                else:
                    dic_cats1[d_real.cat1].append(d_real)
        # create categories 
        for _,dic_line in dic_cats1.items():
            self.cats1.append(Cat1(dic_line,self.tot_ex_real,self.tot_inc_real))
    
    def sort_cats1(self):
        self.cats1 = sorted(self.cats1, key=lambda cat: (cat.tot < 0, cat.tot))

class Cat1():
    def __init__(self,data,exp_tot,inc_tot):
        self.inc_tot = inc_tot
        self.exp_tot = exp_tot
        self.data = data
        
        #
        self.tot = None
        self.tot_p = None
        self.disp_p = True
        self.cats2=[]
        self.name = self.data[0].cat1
        #
        if self.name in C.NOT_REAL_ACTION:
            self.disp_p = False
            
        self.get_tot()
        self.create_cats2()
        self.nb_cat2 = sum(1 for cat2 in self.cats2 if cat2.name is not None)

    def get_tot(self):
        # tot
        self.tot = 0
        for d in self.data:
            if d.value is None:
                d.value = 0
            self.tot+=d.value
        # tot_p
        if self.tot > 0:
            if self.inc_tot !=0:
                self.tot_p = self.tot/self.inc_tot
            else:
                self.tot_p = ""
        elif self.tot < 0:
            if self.exp_tot !=0:
                self.tot_p = self.tot/self.exp_tot
            else:
                self.tot_p = ""
        else:
            self.tot_p = 0

    def create_cats2(self):
        dic_cats2 = {}
        for data in self.data:
            if data.cat2 is not None:
                if data.cat2 not in dic_cats2:
                    dic_cats2[data.cat2] = [data]
                else:
                    dic_cats2[data.cat2].append(data)
                    
        # create categories 
        for _,dic_line in dic_cats2.items():
            self.cats2.append(Cat2(dic_line,self.tot))

class Cat2():
    def __init__(self,data,tot_cat1):
        self.data = data
        self.tot_cat1 = tot_cat1
        
        #
        self.tot = None
        self.tot_p = None
        self.name = self.data[0].cat2
        #
        
        self.get_tot()
        
    def get_tot(self):
        # tot
        self.tot = 0
        for d in self.data:
            self.tot+=d.value
        # tot_p
        if self.tot_cat1 !=0:
            self.tot_p = self.tot/self.tot_cat1
        else:
            self.tot_p = ""





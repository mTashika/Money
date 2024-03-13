import pandas as pd
import numpy as np

FLAG_M = '#'
FLAG_BAL = '>'
NOT_REAL_ACTION = ['virement moi','test']

def verif_nan(val):
    return ~(isinstance(val, (int,float)) and np.isnan(val))

class Sheet():
    def __init__(self,path,sheet = None):
        self.file_path = path
        self.sheet_name = sheet
        self.data = None
        self.cell_sold_init = [1,0]
        self.c_months = 0
        self.col_cat1 = self.c_months
        self.col_cat2 = self.col_cat1+1
        self.c_note = self.col_cat2+1
        self.c_val = self.c_note+1
        self.c_w1 = self.c_val+1
        self.c_w2 = self.c_w1+1
        self.c_w3 = self.c_w2+1
        self.c_w4 = self.c_w3+1
        self.c_w5 = self.c_w4+1
        
        self.months_obj = []
        
        self.read_excel()
        self.verif_bal_init()
        self.create_months()
        self.write_in_sheet()

    def read_excel(self):
        try:
            if self.sheet_name:
                self.data = pd.read_excel(self.file_path, sheet_name=self.sheet_name, header=None)
            else:
                self.data = pd.read_excel(self.file_path, header=None)
        except Exception as e:
            print("Error reading Excel file:", e)

    def get_all_months(self):
        self.dic_m = {}
        for row, val in self.data.iterrows():
            v = val[self.c_months]
            if isinstance(v,str):
                if FLAG_M in v:
                    m_txt = v.replace(FLAG_M, '')
                    if m_txt not in self.dic_m:
                        self.dic_m[m_txt] = row
            
    def verif_bal_init(self):
        val = self.data.iloc[self.cell_sold_init[0],self.cell_sold_init[1]]
        if not isinstance(val,(int,float)) or not verif_nan(val):
            raise Exception ('Balance value not at the expected spot')
        
    def get_bal_prev(self):
        if self.months_obj == []:
            return self.data.iloc[self.cell_sold_init[0],self.cell_sold_init[1]]
        return self.months_obj[-1].bal_usr
        
    def create_months(self):
        self.get_all_months()         
        def get_row_end_month(c_name,dic):
            dic= list(dic.items())
            for i, (key, _) in enumerate(dic):
                if key == c_name:
                    if i + 1 >= len(dic):
                        return None
                    else:
                        return dic[i + 1][1]
            
        for month, row_s in self.dic_m.items():
            row_e = get_row_end_month(month,self.dic_m)
            m_data = self.data.iloc[row_s:row_e,:]
            
            self.months_obj.append(Month(m_data,self.get_bal_prev()))

    def write_in_sheet(self):
        
        for month in self.months_obj:
            row = month.data.index.to_list()[0]
            self.data.at[row,self.c_w1] = month.txt_m_tot
            self.data.at[row,self.c_w2] = month.txt_m_exp_tot
            self.data.at[row,self.c_w3] = month.txt_m_inc_tot
            self.data.at[row,self.c_w4] = month.txt_m_exp_real
            self.data.at[row,self.c_w5] = month.txt_m_inc_real
            row +=1
            self.data.at[row,self.c_w1] = month.tot
            self.data.at[row,self.c_w2] = month.tot_exp
            self.data.at[row,self.c_w3] = month.tot_inc
            self.data.at[row,self.c_w4] = month.tot_exp_real
            self.data.at[row,self.c_w5] = month.tot_inc_real
            
            self.data.at[month.row_bal+1,self.col_cat1] = month.txt_m_bal_real
            self.data.at[month.row_bal+1,self.col_cat2] = month.bal_pred
            self.data.at[month.row_bal,self.c_note] = month.txt_m_verif
            self.data.at[month.row_bal,self.c_val] = month.verif
            
            month.write_cat1(self.data,month.data.index.to_list()[0]+3,self.c_w1)
            
        self.data.to_excel(self.file_path, sheet_name=self.sheet_name,index=False, header=False)

        


class Month():
    def __init__(self,data,bal_prev):
        self.data=data
        self.name = data.iloc[0,0].replace(FLAG_M, '')
        self.cat1=[]
        self.tot = 0
        self.tot_exp = 0
        self.tot_inc = 0
        self.tot_exp_real = 0
        self.tot_inc_real = 0
        self.bal_pred = 0
        self.bal_usr = 0
        self.bal_prev = bal_prev
        
        # txt
        self.txt_m_tot = 'Total transaction'
        self.txt_m_exp_tot = 'dépense totale'
        self.txt_m_inc_tot = 'revenue totale'
        self.txt_m_exp_real= 'dépense réelle'
        self.txt_m_inc_real = 'revenue réel'
        self.txt_m_bal_real= 'solde fin du mois :'
        self.txt_m_verif='Verif'
        
        self.row_s = None
        self.find_start()
        
        self.d_cat1 = self.data.loc[self.row_s:,0]
        self.d_val = self.data.loc[self.row_s:,3]
        
        
        self.update_tot()
        self.update_bal()
        self.create_cat1()

    def find_start(self):
        for row,val in self.data.iterrows():
            if isinstance(val[0],str):
                if not FLAG_M in val[0]:
                    self.row_s = row
                    return
                
    def update_tot(self):
        for _,val in self.data.iterrows():
            euro = val[3]
            ok_real = True if val[0]not in NOT_REAL_ACTION else None
            if isinstance(euro,(int,float)) and verif_nan(euro):
                self.tot+=euro
                if euro > 0:
                    self.tot_inc+=euro
                    self.tot_inc_real+=euro if ok_real else None
                    
                else:
                    self.tot_exp+=euro
                    self.tot_exp_real+=euro if ok_real else None
    
    def update_bal(self):
        for row,val in self.data.iterrows():
            if isinstance(val[0],str):
                if FLAG_BAL in val[0]:
                    self.bal_usr = val[1]
                    self.row_bal = row
                    break
        self.bal_pred = self.bal_prev + self.tot
        self.verif_bal()
        
    def verif_bal(self):
        marg = 10 # %
        if np.abs(self.bal_usr*(1-marg/100)) <= np.abs(self.bal_pred) <= np.abs(self.bal_usr*(1+marg/100)):
            self.verif = True
        else:
            self.verif = False
            
    def create_cat1(self):
        # get all categories with the dataframe corresponding to it
        dic_cat1 = {}
        for id,val in self.d_cat1.items():
            if verif_nan(val):
                if val not in dic_cat1:
                    dic_cat1[val] = self.data.loc[id]
                else:
                    dic_cat1[val]=pd.concat([dic_cat1[val], self.data.loc[id]], axis=1)
                    
        # create categories 
        for _,df in dic_cat1.items():
            self.cat1.append(Cat(df.transpose(),self.tot_exp_real,self.tot_inc_real))
            
    def write_cat1(self,data,row_s,col_s):
        for cat in self.cat1:
            cat.write(data,row_s,col_s)
            if len(cat.data.shape) == 1:
                add_row = 2
            else:
                add_row = len(cat.cat2)+2
            row_s+=add_row
            

class Cat():
    def __init__(self,data,exp_tot,inc_tot):
        self.inc_tot = inc_tot
        self.exp_tot = exp_tot
        self.data = data
        self.txt_total = 'Total'
        self.disp_p = True
        self.tot_p = None
        self.cat2=[]
        
        if len(self.data.shape) == 1:
            self.name = self.data[0]
            self.d_cat2 = self.data.loc[1]
            self.d_euro = self.data.loc[3]
            self.tot_euro = self.d_euro
            
            if self.name in NOT_REAL_ACTION:
                self.disp_p = False
            if self.disp_p:
                if self.tot_euro>0:
                    ref = self.inc_tot
                else:
                    ref = self.exp_tot
                self.tot_p = self.tot_euro*100/ref
                
            if verif_nan(self.d_cat2):
                self.txt_total = self.d_cat2
            return
        
        self.name = self.data.iloc[0,0]
        if self.name in NOT_REAL_ACTION:
            self.disp_p = False
        self.d_cat2 = self.data.loc[:,1]
        self.d_euro = self.data.loc[:,3]
        self.tot_euro = 0
        self.tot_p = None
        self.get_tot()
        self.create_cat2()
        
        
    def get_tot(self):
        for id,val in self.d_euro.items():
            self.tot_euro+=val
        
        if self.disp_p:
            if self.tot_euro>0:
                ref = self.inc_tot
            else:
                ref = self.exp_tot
            self.tot_p = self.tot_euro*100/ref
            
    def create_cat2(self):
        # check if multiple cat2
        if self.d_cat2.nunique(dropna=False) == 1:  # only 1 sub categorie
            self.txt_total = self.d_cat2.iloc[0]
        else:
            # get all categories with the dataframe corresponding to it
            dic_cat2 = {}
            for id,val in self.d_cat2.items():
                if verif_nan(val):
                    if val not in dic_cat2:
                        dic_cat2[val] = self.data.loc[id]
                    else:
                        dic_cat2[val]=pd.concat([dic_cat2[val], self.data.loc[id]], axis=1)
                        
            # create categories 
            for _,df in dic_cat2.items():
                self.cat2.append(Cat2(df.transpose(),self.tot_euro,self.disp_p))
        
    def write(self,data,row_s,col_s):
        data.at[row_s,col_s] = self.name
        col_s+=1
        for cat2 in self.cat2:
            cat2.write(data,row_s,col_s)
            row_s+=1
        data.at[row_s,col_s] = self.txt_total
        data.at[row_s,col_s+1] = self.tot_euro
        data.at[row_s,col_s+2] = self.tot_p if self.disp_p else None
        
        
class Cat2():
    def __init__(self,data,tot_cat,disp_p):
        print(data)
        self.data = data
        self.tot_cat = tot_cat
        self.tot_p = None
        self.disp_p = disp_p
        
        if len(self.data.shape) == 1:
            self.name = self.data[1]
            self.d_euro = self.data.loc[3]
            self.tot_euro = self.d_euro
            if self.disp_p:
                self.tot_p = self.tot_euro*100/self.tot_cat
            return
        
        self.name = self.data.iloc[0,1]
        self.d_euro = self.data.loc[:,3]
        self.tot_euro = 0
        self.tot_p = 0
        
        self.get_tot()
        
    def get_tot(self):
        for id,val in self.d_euro.items():
            self.tot_euro+=val
        if self.disp_p:
            self.tot_p = self.tot_euro*100/self.tot_cat
            
            
    def write(self,data,row_s,col_s):
        data.at[row_s,col_s] = self.name
        data.at[row_s,col_s+1] = self.tot_euro
        data.at[row_s,col_s+2] = self.tot_p if self.disp_p else None

if __name__=='__main__':
    excel_path = 'exemple.xlsx'
    sheet_name = 'Feuil2'
    
    sheet = Sheet(excel_path,sheet_name)

                
    print('FINISH')

    
    

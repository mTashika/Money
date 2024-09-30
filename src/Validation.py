from Tools import check_cell_value
from Const import EXT_BILAN_TXT_VALID
from openpyxl.styles import Alignment

class Valid():
    def __init__(self,ws,mks):
        real_income = 0
        real_loss = 0
        for row in range(mks.B_IN_ST_LINE,mks.B_in_ed_line+1):
            val = ws.cell(row=row,column=mks.B_IN_RE_COL).value
            if val is not None:
                real_income += val
        for row in range(mks.B_LOS_ST_LINE,mks.B_los_ed_line+1):
            val = ws.cell(row=row,column=mks.B_LOS_RE_COL).value
            if val is not None:
                real_loss += val
            
        sold_start = ws.cell(row=mks.B_BILAN_ST_SOLD[0],column=mks.B_BILAN_ST_SOLD[1]).value
        sold_end = ws.cell(row=mks.B_BILAN_ED_SOLDRE[0],column=mks.B_BILAN_ED_SOLDRE[1]).value
        cell_val = ws.cell(row=mks.B_BILAN_ED_VAL[0],column=mks.B_BILAN_ED_VAL[1])
        if round(sold_start + real_income + real_loss,0) == round(sold_end,0):
            cell_val.style = 'Good'
        elif (not ws.cell(row=mks.b_ext_verif[0],column=mks.b_ext_verif[1]).value or 
            check_cell_value(real_loss) not in [1,2] or
            check_cell_value(real_income) not in [1,2]):
            cell_val.style = 'Neutral'
        else:
            cell_val.style = 'Bad'
            
        cell_val.number_format  = f';;;"{EXT_BILAN_TXT_VALID}"'
        cell_val.alignment=Alignment(horizontal='center', vertical='center')
    
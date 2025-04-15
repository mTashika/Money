from Const import PROTECTION_PASSWORD
from Mark_Sheet import SheetMarker

    # for row in ws.iter_rows(min_col=1, max_col=1, max_row=100):
    #     for cell in row:
    #         cell.protection = Protection(locked=False)


def set_protection(ws):
    mks = SheetMarker(ws)
    # protect all the sheet
    protect_ws(ws)
    
    # unprotect certain part
    cell_val_ex_in = ws.cell(mks.B_IN_TOT_EX[0],mks.B_IN_TOT_EX[1])
    unprotect_cell(cell_val_ex_in)
    cell_val_ex_los = ws.cell(mks.B_LOS_TOT_EX[0],mks.B_LOS_TOT_EX[1])
    unprotect_cell(cell_val_ex_los)
    
    unprotect_range(ws,mks.B_LOS_IN_ST_LINE,mks.B_LOS_IN_ED_LINE_CLEAR,mks.B_IN_NAME_COL,mks.B_REAL_COL_ED)
    unprotect_range(ws,mks.B_LOS_IN_ST_LINE,mks.B_LOS_IN_ED_LINE_CLEAR,mks.B_LOS_NAME_COL,mks.B_REAL_COL_ED)
    unprotect_range(ws,mks.B_BILAN_DATE[0],100,mks.B_REAL_COL_ED+1,100)
    unprotect_range(ws,mks.B_DETAIL_LINE_ST,mks.B_ext_ed_line,mks.B_ID_C1_COL,mks.B_ID_REFUND_COL)
    
def protect_wb(wb):
    wb.security.lockStructure = True

def protect_ws(ws):
    ws.protection.sheet = True
    ws.protection.password = PROTECTION_PASSWORD
    ws.protection.formatCells = False
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.protection.insertColumns = False
    ws.protection.insertRows = False

def protect_range(ws,min_row, max_row, min_col, max_col):
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            protect_cell(cell)
    
def unprotect_range(ws, min_row, max_row, min_col, max_col):
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            unprotect_cell(cell)

def unprotect_cell(cell):
    cell.protection = cell.protection.copy(locked=False)

def protect_cell(cell):
    cell.protection = cell.protection.copy(locked=True)
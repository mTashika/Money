from Const import PROTECTION_PASSWORD

def set_protection(ws,mks):
    # protect all the sheet
    protect_ws(ws)
    
    # unprotect certain part
    cell_val_ex = ws.cell(mks.B_IN_TOT_EX[0],mks.B_IN_TOT_EX[1])
    unprotect_cell(cell_val_ex)
    cell_val_ex = ws.cell(mks.B_LOS_TOT_EX[0],mks.B_LOS_TOT_EX[1])
    unprotect_cell(cell_val_ex)
    
    unprotect_range(ws,mks.B_IN_ST_LINE,mks.B_DETAIL_LINE_ST-2,mks.B_IN_NAME_COL,mks.B_REAL_COL_ED)
    unprotect_range(ws,mks.B_LOS_ST_LINE,mks.B_DETAIL_LINE_ST-2,mks.B_LOS_NAME_COL,mks.B_REAL_COL_ED)
    unprotect_range(ws,mks.B_BILAN_DATE[0],100,mks.B_VALID_EXT_PATH[1]+2,100)
    unprotect_range(ws,mks.B_DETAIL_LINE_ST,mks.B_ext_ed_line,mks.B_ID_C1_COL,mks.B_ID_C2_COL)
    
def protect_wb(wb):
    wb.security.lockStructure = True
def protect_ws(ws):
    ws.protection.enableSelection = 'unlocked'
    ws.protection.formatCells = True
    ws.protection.set_password(PROTECTION_PASSWORD)
    ws.protection.sheet = True
    

def protect_range(ws,min_row, max_row, min_col, max_col):
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.protection = cell.protection.copy(locked=True)
    
def unprotect_range(ws, min_row, max_row, min_col, max_col):
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.protection = cell.protection.copy(locked=False)
    
def unprotect_cell(cell):
    cell.protection = cell.protection.copy(locked=False)
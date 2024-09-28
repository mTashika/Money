from Mark_Sheet import SheetMarker
from Ext_PDF_extraction import PDFExtraction
from Ext_Write import WriteExtraction
from Real_Compute import MonthRealisation
from Real_Write import WriteRealisation
from Validation import Valid
from Styles import StyleCell
from datetime import datetime
import Const

if __name__ == "__main__":
    import openpyxl
    path = r"D:\One Drive\OneDrive\Bureau\Classeur1.xlsx"
    workbook = openpyxl.load_workbook(path)
    StyleCell(workbook)
    ws = workbook.active
    mks = SheetMarker(ws)
    PATH_EXTRACTION = r"D:\One Drive\OneDrive\Bureau\relevejuin.pdf"
    extraction = PDFExtraction(PATH_EXTRACTION)
    w = WriteExtraction(ws,extraction,mks)
    if w.is_valid:
        ws.cell(row=mks.B_VALID_EXT_PATH[0],column=mks.B_VALID_EXT_PATH[1]).value = PATH_EXTRACTION
        ws.cell(row=mks.B_VALID_EXT_DATE[0],column=mks.B_VALID_EXT_DATE[1]).value = datetime.now()
    
    # realisation = MonthRealisation(ws,mks)
    # WriteRealisation(ws,mks,realisation)
    # ws.cell(row=mks.B_VALID_REAL_DATE[0],column=mks.B_VALID_REAL_DATE[1]).value = datetime.now()
    # Valid(ws,mks)
    
    
    workbook.save(path)
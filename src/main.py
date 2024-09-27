from Mark_Sheet import SheetMarker
from Ext_PDF_extraction import PDFExtraction
from Ext_Write import WriteExtraction
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
    mks.print()
    PATH_EXTRACTION = r"D:\One Drive\OneDrive\Bureau\releve_CCP2142397R038_20240328.pdf"
    extraction = PDFExtraction(PATH_EXTRACTION)
    w = WriteExtraction(ws,extraction,mks)
    if w.is_valid:
        ws.cell(row=mks.B_VALID_EXT_PATH[0],column=mks.B_VALID_EXT_PATH[1]).value = PATH_EXTRACTION
        ws.cell(row=mks.B_VALID_EXT_DATE[0],column=mks.B_VALID_EXT_DATE[1]).value = datetime.now()
    
    workbook.save(path)
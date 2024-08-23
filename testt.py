import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment

# Create a new workbook and select the active worksheet
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "#Juin"

# Fill colors
header_fill = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")
category_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
description_fill = PatternFill(start_color="EAD1DC", end_color="EAD1DC", fill_type="solid")
amount_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
verification_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")

# Header styling
ws['A1'].value = "#Juin"
ws['A1'].fill = header_fill
ws.merge_cells('A1:E1')
ws['A1'].alignment = Alignment(horizontal='center')

# Categories and Descriptions styling
for row in range(2, 26):
    ws[f'A{row}'].fill = category_fill
    ws[f'B{row}'].fill = description_fill
    ws[f'C{row}'].fill = description_fill
    ws[f'D{row}'].fill = amount_fill

# Final row (solde fiche de paye and solde estimé)
ws['A26'].value = ">solde fiche de paye (fin du mois)"
ws['C26'].fill = amount_fill
ws['D26'].fill = verification_fill
ws['E26'].fill = verification_fill

ws['A27'].value = "Solde estimé :"
ws['C27'].fill = amount_fill

# Save the workbook
wb.save("styled_table.xlsx")
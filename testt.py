import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

# Charger le fichier Excel
wb = openpyxl.load_workbook(r"D:\One Drive\OneDrive\Bureau\Classeur1.xlsx")

# Accéder aux feuilles
sheet_source = wb['Tools']
sheet_target = wb['Sheet1']

# Récupérer toutes les valeurs de la colonne d'une autre feuille
column_source = sheet_source['A']  # Par exemple, si la colonne est 'A'
values = [cell.value for cell in column_source if cell.value is not None]

# Créer une chaîne de valeurs pour la validation de données
value_list = ",".join(map(str, values))

# Créer un objet de validation de données
dv = DataValidation(type="list", formula1="=Tools!$A:$A")

# Définir la cellule où vous souhaitez ajouter la validation
cell = sheet_target['M20']  # Par exemple, la cellule 'B2'
sheet_target.add_data_validation(dv)

# Ajouter la validation à la cellule cible
dv.add(cell)

# Sauvegarder le fichier
wb.save(r"D:\One Drive\OneDrive\Bureau\Classeur1.xlsx")

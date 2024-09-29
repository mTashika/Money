import Const as C
import unicodedata
from openpyxl.styles import Protection
from openpyxl.utils import range_boundaries,get_column_letter
import numpy as np

def check_cell_value(cell_value):
    if cell_value is None:
        return -1
    if isinstance(cell_value, str):
        return 0
    if isinstance(cell_value, int):
        return 1
    if isinstance(cell_value, float):
        return 2
    
def clear_zone(ws,minrow,maxrow,mincol,maxcol):        
        mcr_coord_list = [mcr.coord for mcr in ws.merged_cells.ranges]
        for mcr in mcr_coord_list:
            min_col, min_row, max_col, max_row = range_boundaries(mcr)
            if min_col>=mincol and max_col<=maxcol and min_row>=minrow and max_row<=maxrow:
                ws.unmerge_cells(mcr)
        for row in range(minrow, maxrow + 1):
            for col in range(mincol, maxcol + 1):
                cell = ws.cell(row=row, column=col)
                cell.value = None
                cell.style = 'Normal'
def unmerge_cells_by_coords(ws, start_row, start_col, end_row, end_col):
    start_cell = get_column_letter(start_col) + str(start_row)
    end_cell = get_column_letter(end_col) + str(end_row)
    ws.unmerge_cells(f'{start_cell}:{end_cell}')

def generate_table(t_size, val_max):
    table = []
    for i in range(t_size):
        table.append((i % val_max)+1)  # Génère un nombre de 1 à x en boucle
    return table


def remove_accents(input_str):
    # Normalize the string to decompose accented characters into base characters + accents
    normalized_str = unicodedata.normalize('NFD', input_str)
    # Remove the accents
    without_accents = ''.join(char for char in normalized_str if unicodedata.category(char) != 'Mn')
    return without_accents

def center_window(win, width, height):
        # Obtenir la taille de l'écran
        screen_width = win.winfo_screenwidth()
        screen_height = win.winfo_screenheight()

        # Calculer les coordonnées x et y pour centrer la fenêtre
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2

        # Définir la taille et la position de la fenêtre
        win.geometry(f"{width}x{height}+{x}+{y}")
        
class FinancialOperation:
    '''
    Classe used for the data extraction from a PDF
    '''
    def __init__(self,num,type,name,value,unit='€'):
        self.num = num
        self.type = type
        self.name=name
        self.value = value
        self.unit = unit
    def print(self):
        print(f'{self.num}\t{self.name}: {self.value}{self.unit} ({self.type})')
def get_previous_month(current_month):
    """
    Returns the previous month given the current month.

    :param current_month: Name of the current month (e.g., 'July').
    :return: Name of the previous month (e.g., 'June').
    """
    # List of months in order
    months = C.MONTH_FR
    current_month = current_month.lower()
    current_month = remove_accents(current_month)
    
    # Find the index of the current month
    index_month = months.index(current_month)
    
    # Calculate the index of the previous month
    index_previous_month = (index_month - 1) % 12
    month_prev = months[index_previous_month].lower()
    month_prev = remove_accents(month_prev)
    return month_prev

def column_to_number(column):
    """
    Converts an Excel column letter to its numeric index (1 = A, 2 = B, 3 = C, etc.).
    
    Args:
    column (str): Letter(s) representing the column (e.g., "A", "B", "C", "AA").
    
    Returns:
    int: Numeric index of the column.
    """
    column = column.upper()  # Ensure the letter is in uppercase
    number = 0
    for letter in column:
        number = number * 26 + (ord(letter) - ord('A') + 1)
    return number

def write_list_to_excel(sheet,lst,start):
    """
    Writes two lists to columns in an Excel file, starting from specified coordinates.

    :param lst: The first list to write to the first column.
    :param file_name: The name of the Excel file to create or overwrite.
    """
    if isinstance(start[1],str):
        start[1] = column_to_number(start[1])
    for row_index, value in enumerate(lst, start=start[0]):
        cell = sheet.cell(row=row_index, column=start[1])
        cell.value = value
        
def is_smallest_month(month_list, month):
    # Define a dictionary to map month names to their order
    month_order = {
    'janvier': 1, 'fevrier': 2, 'mars': 3, 'avril': 4,
    'mai': 5, 'juin': 6, 'juillet': 7, 'aout': 8,
    'septembre': 9, 'octobre': 10, 'novembre': 11, 'decembre': 12
    }
    month = month.lower()
    month = remove_accents(month)
    
    month_list_lower = [m.lower() for m in month_list]
    month_list_lower = [remove_accents(m) for m in month_list]
    valid_month_list = [m for m in month_list_lower if m in month_order]
    
    # Check if the provided month is valid
    if month not in month_order:
        return False
    if not valid_month_list:
        return False
    
    smallest_month = min(valid_month_list, key=lambda m: month_order[m])
    # Determine if the provided month is the smallest among valid months
    return month_order[month] == month_order[smallest_month]

def find_valid_row(ws):
    for row in range(1, ws.max_row + 1):
        cell_value = ws.cell(row=row, column=1).value
        if cell_value and C.BALISE_SOLD_MONTH in str(cell_value):
            return row

def update_init_sold(wb,sheet_name):
    'Update le sold dans tool sheet si le moi ajouter est le plus petit'
    ws_names = wb.sheetnames
    if len(ws_names) >1:
        if is_smallest_month(ws_names,sheet_name):
            print('is smallest month')
            ws_init = wb[C.INIT_SHEET_NAME]
            ws_init[C.INIT_SHEET_MONTH_CELL[1]] = get_previous_month(sheet_name)
            ws_init[C.INIT_SHEET_SOLD_CELL[1]] = None

def protect_ws(ws):
    for row in ws.iter_rows():
            for cell in row:
                cell.protection = Protection(locked=True)
    ws.protection.sheet = True
    
def protect_range(ws,min_row, max_row, min_col, max_col):
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.protection = cell.protection.copy(locked=True)
    ws.protection.sheet = True
    
def unprotect_range(ws, min_row, max_row, min_col, max_col):
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.protection = cell.protection.copy(locked=False)
    ws.protection.sheet = True
    
def unprotect_cell(cell):
    cell.protection = cell.protection.copy(locked=False)
    
def is_file_closed(file_path):
    try:
        # Essayer d'ouvrir le fichier en mode exclusif
        with open(file_path, 'r+'):
            pass
        return True
    except IOError:
        # Si une exception est levée, le fichier est probablement encore ouvert
        return False
def verif_nan(val):
    return ~(isinstance(val, (int,float)) and np.isnan(val))
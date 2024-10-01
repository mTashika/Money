import unicodedata
from openpyxl.utils import range_boundaries,get_column_letter
from datetime import datetime

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
def remove_accent(input_str):
    return ''.join(
        c for c in unicodedata.normalize('NFD', input_str)
        if unicodedata.category(c) != 'Mn'
    )
    
def normalize_string(input_str):
    # Convert to lowercase
    lower_str = input_str.lower()
    # Remove accents
    normalized_str = remove_accent(lower_str)
    return normalized_str

def center_window(win, width, height):
        # Obtenir la taille de l'écran
        screen_width = win.winfo_screenwidth()
        screen_height = win.winfo_screenheight()

        # Calculer les coordonnées x et y pour centrer la fenêtre
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2

        # Définir la taille et la position de la fenêtre
        win.geometry(f"{width}x{height}+{x}+{y}")
        
def get_current_month():
    return datetime.now().strftime("%B")
def get_current_year():
    return datetime.now().strftime("%Y")

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

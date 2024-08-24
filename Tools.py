import CONST as C
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
    
    # Find the index of the current month
    index_month = months.index(current_month.capitalize())
    
    # Calculate the index of the previous month
    index_previous_month = (index_month - 1) % 12
    
    # Return the previous month
    return months[index_previous_month]

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
        

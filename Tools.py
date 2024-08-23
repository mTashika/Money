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
    months = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ]
    
    # Find the index of the current month
    index_month = months.index(current_month.capitalize())
    
    # Calculate the index of the previous month
    index_previous_month = (index_month - 1) % 12
    
    # Return the previous month
    return months[index_previous_month]

import openpyxl

def write_list_to_excel(sheet,lst,start):
    """
    Writes two lists to columns in an Excel file, starting from specified coordinates.

    :param lst: The first list to write to the first column.
    :param file_name: The name of the Excel file to create or overwrite.
    """
    for row_index, value in enumerate(lst, start=start[0]):
        cell = sheet.cell(row=row_index, column=start[1])
        cell.value = value
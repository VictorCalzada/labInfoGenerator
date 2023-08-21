import openpyxl
import xlrd

from openpyxl.utils.exceptions import InvalidFileException

def read_table_op(file:str, sheet_name:str, rows:list, cols:list):
    """Lectura de tablas de excel con formato .xlsx"""
    wb = openpyxl.load_workbook(file)
    sh = wb[sheet_name]
    
    cols = cols_to_num(cols = cols)

    table = []
    for row in sh.iter_rows(min_row=rows[0], max_row=rows[1], min_col=cols[0] + 1, max_col=cols[1] + 1, values_only=True, ):
        table.append(row)

    return table

def read_table_xl(file: str, sheet_name: str, rows: list, cols: list):
    """Lectura de tablas de excel con formato .xls"""
    book = xlrd.open_workbook(file)
    sh = book.sheet_by_name(sheet_name)
    
    cols = cols_to_num(cols = cols)
    table = []
    for i in range(rows[0], rows[1]):
        row = []
        for j in range(cols[0], cols[1] + 1): # Hay que sumerle uno por que el range no coge el último número del rango
            row.append(sh.cell_value(rowx = i, colx = j))

        table.append(row)

    return table

def cols_to_num(cols: list):
    """Paso de columnas con letra a número"""
    letter_to_number = {
    'A': 0, 'B': 1, 'C': 2, 'D': 3, 'E': 4,
    'F': 5, 'G': 6, 'H': 7, 'I': 8, 'J': 9,
    'K': 10, 'L': 11, 'M': 12, 'N': 13, 'O': 14,
    'P': 15, 'Q': 16, 'R': 17, 'S': 18, 'T': 19,
    'U': 20, 'V': 21, 'W': 22, 'X': 23, 'Y': 24, 'Z': 25
}
    return [letter_to_number[cols[0]], letter_to_number[cols[1]]]


def read_table(file: str, sheet_name: str, rows: list, cols: list):
    """Lectura de tablas excel independientemente de la extension (.xls, .xlsx)"""
    try:
        return read_table_op(file, sheet_name, rows, cols)
    except InvalidFileException:
        return read_table_xl(file, sheet_name, rows, cols)


# from docx import Document

def add_row(table, contenido: list):
    size = len(table.table.rows), len(table.table.row_cells(0))
    if len(contenido)!=size[1]:
        raise ListasLongitudDiferenteError(len(contenido), size[1]) 
     

class ListasLongitudDiferenteError(Exception):
    def __init__(self, longitud1, longitud2):
        self.longitud1 = longitud1
        self.longitud2 = longitud2
        super().__init__(f"Las listas tienen longitudes diferentes: {longitud1} y {longitud2}")


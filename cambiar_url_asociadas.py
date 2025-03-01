import openpyxl

# Ruta del archivo de Excel
EXCEL_FILE = "Ranking_TorredeEboli.xlsx"
SHEET_NAME = "Hoja1"  # Cambia esto si el nombre es distinto

# Abrir el archivo de Excel y la hoja
wb = openpyxl.load_workbook(EXCEL_FILE)
sheet = wb[SHEET_NAME]

# Obtener todos los valores
data = list(sheet.values)

# Convertir datos en un diccionario
headers = data[0]  # Primera fila como encabezados
rows = data[1:]    # Resto de filas

# Índices de columnas relevantes
col_fide_id = headers.index("FIDE ID")

# URL base de la FIDE
url_base = "https://ratings.fide.com/profile/"

# Iterar sobre cada fila (sin la cabecera)
for i, row in enumerate(rows):
    fide_id = str(int(row[col_fide_id]))  # Obtener el ID de FIDE
    url = url_base + fide_id

    # Actualizar la celda de FIDE ID para que sea un hipervínculo
    cell = sheet.cell(row=i + 2, column=col_fide_id + 1)  # "+2" porque la primera fila es la cabecera
    cell.hyperlink = url
    cell.value = fide_id
    cell.style = "Hyperlink"

# Guardar el archivo de Excel
wb.save(EXCEL_FILE)

print("Proceso completado. URLs actualizadas en el archivo de Excel.")
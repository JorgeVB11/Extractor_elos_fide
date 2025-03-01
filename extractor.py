import openpyxl
import requests
from bs4 import BeautifulSoup
import time

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
col_elo = headers.index("ELO FIDE")

# URL base de la FIDE
url_base = "https://ratings.fide.com/profile/"

# Lista para nuevos valores de Elo
elos = []

# Iterar sobre cada fila (sin la cabecera)
for row in rows:
    fide_id = str(row[col_fide_id])  # Obtener el ID de FIDE
    url = url_base + fide_id

    # Hacer solicitud HTTP con un user-agent
    headers = {'User-Agent': 'Mozilla/5.0'}
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')
        div_standard = soup.find("div", class_="profile-standart profile-game")

        if div_standard:
            elo = div_standard.find("p").text.strip()
            print(f"Jugador {fide_id}: Elo {elo}")
            elos.append(elo)
        else:
            print(f"Jugador {fide_id}: No se encontró Elo")
            elos.append("N/A")
    else:
        print(f"Error al acceder a {url}")
        elos.append("Error")

    time.sleep(2)  # Evitar bloqueos

# Escribir los nuevos valores en el archivo de Excel
for i, elo in enumerate(elos):
    sheet.cell(row=i + 2, column=col_elo + 1, value=elo)  # "+2" porque la primera fila es la cabecera

# Guardar el archivo de Excel
wb.save(EXCEL_FILE)

print("Proceso completado. Datos actualizados en el archivo de Excel.")
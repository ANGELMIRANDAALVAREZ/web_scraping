

import openpyxl
import requests
from bs4 import BeautifulSoup
import subprocess


def check_dimples(url):
    try:
        response = requests.get(url)
        response.raise_for_status()

        # Parsear el contenido HTML de la página
        soup = BeautifulSoup(response.text, 'html.parser')

        # Buscar la clase 'dimples-section'
        section = soup.find(class_='dimples-section')

        # Retornar 1 si se encontró la clase, de lo contrario 0
        return 1 if section else 0
    except requests.RequestException as e:
        print(f"Error al obtener la URL {url}: {e}")
        return 0


# Cargar el libro de Excel
workbook = openpyxl.load_workbook('/Users/angel1/Desktop/flags_esteticos_inventario.xlsx')


sheet = workbook.active

# Iterar
for row in range(2, 5819):  #rango de la columna (va de la celda 2 a la 5819)
    # Obtener el URL
    url = sheet[f'A{row}'].value

    # Verificar si la página tiene la clase 'dimples-section'
    sheet[f'H{row}'] = check_dimples(url)  # Guardar el resultado en la columna h

# Guarda el libro de Excel modificado
workbook.save('/Users/angel1/Desktop/STOCKS ID.xlsx')

# Abre el archivo de Excel con la aplicación predeterminada en macOS
subprocess.call(['open', '/Users/angel1/Desktop/STOCKS ID.xlsx'])

print("Proceso completado, archivo de Excel sido actualizado.")

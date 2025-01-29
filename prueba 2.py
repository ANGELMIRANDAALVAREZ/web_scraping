import openpyxl
import requests
from bs4 import BeautifulSoup
import subprocess
from tqdm import tqdm

def extract_picture_count(soup):
    # Buscar todas las etiquetas <picture> dentro de la sección 'drago-vip-dimples'
    dimples_section = soup.select_one('dimples-section')
    if dimples_section:
        pictures = dimples_section.find_all('picture')
        return len(pictures)
    return 0

def check_dimples_and_count_pictures(url):
    try:
        response = requests.get(url, timeout=300)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        picture_count = extract_picture_count(soup)
        return (1, picture_count) if picture_count else (0, 0)
    except requests.Timeout:
        print(f"Tiempo de espera agotado para la URL {url}")
        return (3, 0)
    except requests.RequestException as e:
        print(f"Error con la URL {url}: {e}")
        return (0, 0)

# Cargar el libro de Excel
workbook = openpyxl.load_workbook('/Users/angel1/Desktop/Prueba_2.xlsx')
sheet = workbook.active

# Crear una barra de progreso para las URLs
for row in tqdm(range(2, 1951), desc="Escrapeando URLs"):
    url = sheet[f'A{row}'].value
    dimple_check, photo_count = check_dimples_and_count_pictures(url)
    sheet[f'H{row}'] = dimple_check
    sheet[f'I{row}'] = photo_count if dimple_check else 0

    # Guardar el libro de Excel cada 20 registros escrapeados
    if row % 5 == 0 or row == 1950:
        workbook.save('/Users/angel1/Desktop/Prueba_2.xlsx')
        print(f"Progreso guardado en la fila {row}.")

# Guardar los cambios finales en el libro de Excel
workbook.save('/Users/angel1/Desktop/Prueba_2.xlsx')

# Abrir el archivo de Excel con la aplicación predeterminada
subprocess.call(['open', '/Users/angel1/Desktop/Prueba_2.xlsx'])

print("Proceso completado. El archivo de Excel ha sido actualizado y abierto.")

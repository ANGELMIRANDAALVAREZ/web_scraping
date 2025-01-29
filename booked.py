import openpyxl
import requests
from bs4 import BeautifulSoup
import subprocess
from tqdm import tqdm  # Necesitas instalar tqdm si no lo tienes

def check_dimples(url):
    try:
        # Realizar la petición con un tiempo límite de 300 segundos (5 minutos)
        response = requests.get(url, timeout=300)
        response.raise_for_status()

        # Analizar el contenido HTML de la página
        soup = BeautifulSoup(response.text, 'html.parser')

        # Buscar la clase 'dimples-section'
        section = soup.find(class_='dimples-section')

        # Devolver 1 si se encuentra la clase, 0 en caso contrario
        return 1 if section else 0
    except requests.Timeout:
        # Imprimir el mensaje de tiempo agotado y devolver 3
        print(f"Tiempo de espera agotado para la URL {url}")
        return 3
    except requests.RequestException as e:
        # Imprimir cualquier otro tipo de error HTTP
        print(f"Error con la URL {url}: {e}")
        return 0

# Cargar el libro de Excel
workbook = openpyxl.load_workbook('/Users/angel1/Desktop/flags_esteticos_inventario_reservas.xlsx')
sheet = workbook.active

# Crear una barra de progreso para las URLs
for row in tqdm(range(2, 1951), desc="Escrapeando URLs"):  # Actualizado a 1949 URLs
    # Obtener el URL de la celda
    url = sheet[f'A{row}'].value

    # Verificar y registrar si la página tiene la clase 'dimples-section'
    sheet[f'H{row}'] = check_dimples(url)

    # Guardar el libro de Excel cada 200 registros escrapeados
    if row % 20 == 0 or row == 1950:  # Asegurarse de guardar al final del proceso también
        workbook.save('/Users/angel1/Desktop/flags_esteticos_inventario_reservas.xlsx')
        print(f"Progreso guardado en la fila {row}.")

# Guardar los cambios finales en el libro de Excel
workbook.save('/Users/angel1/Desktop/flags_esteticos_inventario_reservas.xlsx')

# Abrir el archivo de Excel con la aplicación predeterminada
subprocess.call(['open', '/Users/angel1/Desktop/flags_esteticos_inventario_reservas.xlsx'])

print("Proceso completado. El archivo de Excel ha sido actualizado y abierto.")


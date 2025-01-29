from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl
from tqdm import tqdm
import time

def get_image_counter_selenium(url, driver, attempt=1):
    try:
        # Ir a la URL
        driver.get(url)

        # Esperar a que la página inicial se cargue completamente
        wait = WebDriverWait(driver, 10)
        if "Lo sentimos, el auto que buscabas ya no está disponible, pero encontramos este que es muy similar." in driver.page_source:
            return "auto no disponible"
        wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')
        driver.execute_script("document.body.style.zoom='5%'")

        # En el segundo intento, desplazarse al final de la página
        if attempt == 2:
            total_height = driver.execute_script("return document.body.scrollHeight")
            driver.execute_script(f"window.scrollTo(0, {total_height});")
            driver.execute_script("document.body.style.zoom='5%'")
            # En el segundo intento, desplazarse al final de la página
        if attempt == 3:
            total_height = driver.execute_script("return document.body.scrollHeight")
            driver.execute_script(f"window.scrollTo(0, {total_height/2});")
            driver.execute_script("document.body.style.zoom='5%'")

        # Esperar un tiempo adicional para que los elementos se carguen después del desplazamiento
        time.sleep(15)  # Espera 15 segundos

        # Localizar los elementos 'counter'
        image_counters = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'span.counter')))

        # Asegurarse de que se encontraron suficientes elementos
        if len(image_counters) < 2:
            print("No se encontraron suficientes elementos 'span.counter'.")
            raise Exception("Elementos insuficientes")  # Forzar reintento

        # Obtener el texto del segundo elemento
        return image_counters[1].get_attribute('textContent').strip()
    except Exception as e:
        print(f"Error al obtener contador de imágenes para la URL {url}: {e}")
        if attempt < 2:
            print("Reintentando...")
            return get_image_counter_selenium(url, driver, attempt + 1)
        return 'Error al obtener el contador'
chrome_options = Options()
# chrome_options.add_argument("--headless") # Descomenta esta línea para ejecutar sin abrir el navegador

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
workbook = openpyxl.load_workbook('/Users/angel1/Desktop/Prueba_1_5000.xlsx')
sheet = workbook.active
for row in tqdm(range(5000, sheet.max_row + 1), desc="Obteniendo contadores de imágenes"):
    # Comprueba si la celda de la columna I está vacía o contiene un error
    existing_value = sheet[f'I{row}'].value
    condition = sheet[f'H{row}'].value

    # Solo procede si la condición de la columna H es 1
    if condition == 1 and (existing_value is None or existing_value == 'Error al obtener el contador'):
        url = sheet[f'A{row}'].value
        image_counter_text = get_image_counter_selenium(url, driver)

        # Ahora image_counter_text está definido y puedes usarlo en tus condiciones
        if image_counter_text == "auto no disponible":
            print(f"Fila {row}: Auto no disponible.")
            sheet[f'I{row}'] = image_counter_text
        elif image_counter_text == 'Error al obtener el contador':
            print(f"Fila {row}: Error al obtener el contador (3 intentos).")
        else:
            print(f"Fila {row}: Imperfecciones obtenidas.")
            sheet[f'I{row}'] = image_counter_text
    else:
        # Manejo de la condición cuando H es 0 y la celda I está vacía o tiene un error
        if existing_value is None or existing_value == 'Error al obtener el contador':
            sheet[f'I{row}'] = "sin imperfecciones"
            print(f"Fila {row}: Cambio a sin imperfecciones.")
        else:
            # Si la celda I ya tiene un valor y H es 0, se asume que no se necesita cambio
            print(f"Fila {row}: Celda sin cambios.")

    # Guarda el progreso después de cada fila
    workbook.save('/Users/angel1/Desktop/Prueba_1_5000.xlsx')
    print(f"Progreso guardado en la fila {row}.")

# Cierra el navegador y guarda los cambios finales en el libro de Excel
driver.quit()
workbook.save('/Users/angel1/Desktop/Prueba_1_5000.xlsx')
print("Proceso completado. El archivo de Excel ha sido actualizado.")


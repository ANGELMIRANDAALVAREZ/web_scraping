import os
import keyboard
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from concurrent.futures import ThreadPoolExecutor
import openpyxl
from tqdm import tqdm
import time


def get_image_counter_selenium(url, driver, attempt=1):
    try:
        # Ir a la URL
        driver.get(url)

        # Esperar a que la página inicial se cargue completamente
        wait = WebDriverWait(driver, 40)
        wait.until(lambda d: d.execute_script('return document.readyState') == 'complete')
        driver.execute_script("document.body.style.zoom='5%'")

        # En el segundo intento, desplazarse al final de la página
        if attempt == 2:
            total_height = driver.execute_script("return document.body.scrollHeight")
            driver.execute_script(f"window.scrollTo(0, {total_height});")
            driver.execute_script("document.body.style.zoom='5%'")
            time.sleep(15)

        # Esperar un tiempo adicional para que los elementos se carguen después del desplazamiento
        time.sleep(5)  # Espera 5 segundos

        image_counters = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'span.label')))

        # Asegurarse de que se encontraron suficientes elementos
        if len(image_counters) < 2:
            print("No se encontraron suficientes elementos 'span.label'.")
            return "sin dimples"
            raise Exception("Elementos insuficientes")  # Forzar reintento

        # Obtener el texto del segundo elemento
        return image_counters[1].get_attribute('textContent').strip()
    except Exception as e:
        print(f"Error al obtener contador de imágenes para la URL {url}: {e}")
        if attempt < 2:
            print("Reintentando...")
            return get_image_counter_selenium(url, driver, attempt + 1)
        return 'Error al obtener el contador'
def process_urls(start_row, end_row, input_file_path, output_file_path):
    chrome_options = Options()
    #chrome_options.add_argument("--headless")  # ejecutar sin abrir el navegador
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    # Cargar el archivo de entrada
    workbook = openpyxl.load_workbook(input_file_path)
    sheet = workbook.active

    # Crear un nuevo libro de trabajo para el archivo de salida
    output_workbook = openpyxl.Workbook()
    output_sheet = output_workbook.active

    for row in tqdm(range(start_row, end_row + 1), desc=f"Procesando filas {start_row} a {end_row}"):
        url = sheet[f'A{row}'].value
        image_counter_text = get_image_counter_selenium(url, driver)

        if image_counter_text == 'Error al obtener el contador':
            output_sheet[f'E{row}'] = "sin dimples" #aqui cambiamos D por E para probar introducir a otra columna
        else:
            output_sheet[f'E{row}'] = image_counter_text #SE CAMBIO D POR E

        # Guardar el progreso cada 10 URLs
        if (row - start_row + 1) % 10 == 0:
            output_workbook.save(output_file_path)
            print(f"Progreso guardado en la fila {row}.")

    # Guardar y cerrar el navegador al final
    output_workbook.save(output_file_path)
    driver.quit()

def main():
    input_file_path = '/Users/angel1/Desktop/Ops.xlsx'
    base_output_path = '/Users/angel1/Desktop/ops2'
    start_at_row = 1  # Fila de inicio
    end_at_row = 10  # Fila de finalización
    total_rows = end_at_row - start_at_row + 1  # Total de filas a procesar
    threads = 1  # Número de hilos
    rows_per_thread = total_rows // threads  # Filas por hilo

    with ThreadPoolExecutor(max_workers=threads) as executor:
        futures = []
        for i in range(threads):
            start_row = start_at_row + i * rows_per_thread
            end_row = start_row + rows_per_thread - 1
            if i == threads - 1:  # Ajuste para el último segmento
                end_row = end_at_row

            output_file_path = os.path.join(base_output_path, f'output_thread_{i}.xlsx')
            shutil.copy(input_file_path, output_file_path)

            futures.append(executor.submit(process_urls, start_row, end_row, input_file_path, output_file_path))

        # Esperar a que todos los hilos terminen
        for future in futures:
            future.result()

if __name__ == "__main__":
    main()


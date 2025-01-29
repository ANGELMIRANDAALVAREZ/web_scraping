import pandas as pd
import os
import openpyxl
import subprocess
from openpyxl.utils.dataframe import dataframe_to_rows

# Directorio donde están los archivos Excel
directorio = '/Users/angel1/Desktop/Respaldo 2 ops_python'

# Lista de archivos en el directorio
archivos = [archivo for archivo in os.listdir(directorio) if archivo.endswith('.xlsx')]

# Asegurarse de que hay archivos para procesar
if not archivos:
    raise ValueError("No se encontraron archivos Excel en el directorio.")

# Cargar el primer archivo en el libro de trabajo combinado
primer_archivo = os.path.join(directorio, archivos[0])
wb = openpyxl.load_workbook(primer_archivo)
ws = wb.active

# Procesar los archivos restantes
for archivo in archivos[1:]:
    try:
        # Cargar el archivo Excel
        df = pd.read_excel(os.path.join(directorio, archivo), header=None, engine='openpyxl')

        # Escribir los datos en el libro de trabajo combinado
        for index, row in df.iterrows():
            # Escribir solo si la celda no está vacía
            if pd.notna(row[3]):
                ws.cell(row=index+1, column=4, value=row[3])
    except Exception as e:
        print(f"Error al leer el archivo: {archivo}. Error: {e}")
from pync import Notifier

def send_notification(title, message):
    Notifier.notify(message, title=title)

# Llama a esta función al final de tu script
send_notification("Python Script", "El script ha terminado de ejecutarse")
# Guardar el libro de trabajo combinado en un nuevo archivo Excel
wb.save('/Users/angel1/Desktop/merge_total.xlsx')


#subprocess.run(["shortcuts", "run", "Python"])
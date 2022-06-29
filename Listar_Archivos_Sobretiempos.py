import os
import pandas as pd
from pandas import ExcelWriter
from openpyxl import load_workbook
from openpyxl import Workbook

archivos = os.listdir('C:/Users/HB75117/Documents/NOVEDADES/Sobretiempos') # Ruta de la carpeta que queremos listar

df = pd.DataFrame({'Nombre': archivos,})# Estructurando el Archivo en Excel.
df = df[['Nombre']] # Construyendo la primera columna llamada nombre
writer = ExcelWriter('C:/Users/HB75117/Documents/NOVEDADES/Consolidado Sobretiempos.xlsx') # Escribiendo el Archivo Excel.
df.to_excel(writer, 'Informe', index=False) # Nombrando la pestaña
writer.save() # Salvando el Archivo
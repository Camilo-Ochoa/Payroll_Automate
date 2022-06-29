from fileinput import filename
import pandas as pd
import os


archivos = os.listdir('C:/Users/HB75117/Documents/NOVEDADES/Sobretiempos')

for i in range(1,len(archivos)):
    A=archivos[i]
    doc1=(f'C:/Users/HB75117/Documents/NOVEDADES/Sobretiempos/{A}')   #Ruta del archivo
    df = pd.read_excel(
    doc1,
    sheet_name='Formato Reporte',
    usecols=[0,8])
    final_name = df['Unnamed: 8'].values[1]
    print(final_name)
    newname=(f'C:/Users/HB75117/Documents/NOVEDADES/Sobretiempos/{final_name}.xlsb')   #Nuevo nombre como se llamara
    os.rename(doc1,newname) #Renombrado final
import os
from openpyxl import load_workbook
from openpyxl import Workbook

archivos = os.listdir('C:/Users/HB75117/Documents/NOVEDADES/BONOS')    #Listar Archivos y agrega variable archivos

for cont in range(0,len(archivos)):
    A=archivos[cont]            #Archivos en secuencia desde el primero hasta el ultimo.
    doc1=(f'C:/Users/HB75117/Documents/NOVEDADES/BONOS/{A}')   #Ruta del archivo
    try:
        wb = load_workbook(doc1, data_only = True) # carga el archivo
        ws = wb.active
        sh = wb["COP"] #llama la hoja del workbook
        ptsDiff = (sh['A1'].value) #valor de la celda (Nombre de Persona)                                                   
        newname=(f'C:/Users/HB75117/Documents/NOVEDADES/BONOS/{ptsDiff}.xlsm')   #Nuevo nombre como se llamara
        os.rename(doc1,newname) #Renombrado final
        print(newname) #mostrar el nuevo archivo
    except ValueError:
        continue
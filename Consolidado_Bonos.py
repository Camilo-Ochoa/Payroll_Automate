import os
import pandas as pd
from pandas import ExcelWriter
from openpyxl import load_workbook
from openpyxl import Workbook
from numpy import *

ptsDiff=[]
SAP=[]
Bonos=[]
Valor_COP=[]
item=[]
Desde=[]
Hasta=[]

for n in range (0,105): # Solo crear los numeros de item // Ej: si tengo 105 archivos mi rango seria (0,105)
    item.append(n)

for n in range (0,105): # Solo crear los numeros de item // Ej: si tengo 105 archivos mi rango seria (0,105)
    Desde.append('06/01/2022') # Solo crear las fechas desde
    
for n in range (0,105): # Solo crear los numeros de item // Ej: si tengo 105 archivos mi rango seria (0,105)
    Hasta.append('06/30/2022') # Solo crear las fechas hasta

archivos = os.listdir('C:/Users/HB75117/Documents/NOVEDADES/BONOS')
for cont in range(0,len(archivos)):
    A=archivos[cont]            #Archivos en secuencia desde el primero hasta el ultimo.
    doc1=(f'C:/Users/HB75117/Documents/NOVEDADES/BONOS/{A}')   #Ruta del archivo
    print(doc1)
    wb = load_workbook(doc1, data_only = True) # carga el archivo
    ws = wb.active
    sh = wb["COP"] #llama la hoja del workbook

    ptsDiff.append((sh['C7'].value)) #valor de la celda (Nombre de Persona)
    SAP.append((sh['K5'].value)) #otro valor de SAP celda "Numero SAP".
    Bonos.append((sh['H30'].value)) #otro valor de Bonos celda. Dias de Bonos.
    Valor_COP.append((sh['L30'].value)) #otro valor de Valor Total celda. Valor Total de Bonos.


df = pd.DataFrame({'Item': item,
                    'SAP':SAP,
                    'Nombre':ptsDiff,
                    'Desde':Desde,
                    'Hasta':Hasta,
                    'Bonos':Bonos,
                    'Valor COP':Valor_COP})
df = df[['Item', 'SAP', 'Nombre','Desde','Hasta','Bonos','Valor COP']]
writer = ExcelWriter('C:/Users/HB75117/Documents/NOVEDADES/BONOS/ZZZ---Consolidado Final.xlsx')
df.to_excel(writer, 'Informe', index=False)
writer.save()
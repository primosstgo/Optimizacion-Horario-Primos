import pandas as pd
import numpy as np

# Leer el archivo Excel
archivo = pd.ExcelFile('Disponibilidad horaria primos SJ 2023-1.xlsx')
Horarios_primos=[] #arreglo que guardará arreglos de la forma ['nombre',[horarios]]

def float_to_str(lista):
    """
    Esta funcion recorre una lista y cambia los valores de cada posicion a cero
    si el elemento es un flotante, esto porque en las casillas vacias las toma como
    'nan' con tipo flotante (en la hoja de Felipe Rojas no lleno con ceros y dejo vacio por eso implemente esto)
        
    """
    for j in range(0,len(lista)):
            if isinstance(lista[j],float):
                lista[j]= 0 
    
# Recorrer las hojas del archivo, empezando desde la segunda hoja
for nombre_hoja in archivo.sheet_names[2:]:                  #borras el 3 y agarra todas las hojas hasta el final
    # Leer la hoja y obtener el nombre del Primo
    hoja = archivo.parse(nombre_hoja)
    Nombre_Primo = hoja.iloc[0,9]
    Matriz_Disponibilidad=[]
    Matriz_Preferencia=[]
    #print(type(Nombre_Primo))
    for i in range(2,7):
        Disponibilidad = []
        Preferencia = []
        
        Disponibilidad.extend(hoja.iloc[1:5,i].to_numpy().tolist())# turnos desde el bloque 1-8
        Disponibilidad.extend(hoja.iloc[6:9,i].to_numpy().tolist())# turnos desde el bloque 9-14        
        Matriz_Disponibilidad.append(Disponibilidad)
        float_to_str(Disponibilidad)
        
        Preferencia.extend(hoja.iloc[13:17,i].to_numpy().tolist())# turnos desde el bloque 1-8
        Preferencia.extend(hoja.iloc[18:21,i].to_numpy().tolist())# turnos desde el bloque 9-14    
        Matriz_Preferencia.append(Preferencia)
        float_to_str(Preferencia)
       
    for i in range(len(Matriz_Preferencia)):
        for j in range(len(Matriz_Preferencia[i])):
            if Matriz_Preferencia[i][j]:
                Matriz_Disponibilidad[i][j] += 1
        
    #datos_primo = [Nombre_Primo,Matriz_Disponibilidad,Matriz_Preferencia]
    datos_primo = [Nombre_Primo,Matriz_Disponibilidad]
    Horarios_primos.append(datos_primo)

    # Imprimir los datos de la celdas y el nombre de la hoja
    #print(nombre_hoja,':',Nombre_Primo,Matriz_Disponibilidad,Matriz_Preferencia")


""" Debería retornar:

    [['Francisca Daniela Romero Gonzalez', [[0, 0, 0, 0, 0, 1, 0], [0, 1, 0, 1, 1, 1, 0], [0, 0, 1, 1, 1, 1, 0], [0, 1, 1, 1, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0]]]]
    
    en la 'lista de listas' Matriz_Disponibilidad se accede a los valores de la siguiente forma:
    
    Matriz_Disponibilidad[Dia][Horario]

    el rango de Dia es [0-4] (lunes,martes,miercoles,jueves,viernes)
    el rango de Horario es [0-6]( 1-2, 3-4,....,13-14)
    
    

"""

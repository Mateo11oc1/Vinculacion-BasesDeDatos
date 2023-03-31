
import pandas
import math
"""Devuelve la columna como un diccionario de acuerdo a los parametroe"""
def leerColumna(archivo: str, hoja: int, columna: int):
    archivo = pandas.read_excel(archivo, sheet_name=hoja)
    lista = archivo.iloc[7:,columna].tolist()
    columna = {"atractor": lista[0], "numAtractores": lista[2], "tamanio": lista[3:6], "horario": lista[6:11], "dias": lista[12:22]}
    return columna

"""validar que la suma de atractores no sea mayor o 
menor al numero de atractores"""
def validarSuma(columna):
    if columna["numAtractores"] == sum(x for x in columna["tamanio"] if not math.isnan(x)):
        return True
    else:
        return False

columna = leerColumna("../04. Forumularios digitalizados grupo 4.xlsx", 51, 21)
# validarSuma(columna)
print(validarSuma(columna))


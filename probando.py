import openpyxl
import pandas as pd
import math

workbook = openpyxl.load_workbook("../04. Forumularios digitalizados grupo 4.xlsx")

"""Devuelve la columna como un diccionario de acuerdo a los parametros"""
def leerColumna(leido,  columna: int):

    lista = leido.iloc[7:,columna].tolist()
    #[7:,columna] nos da la celda desde la fila 7 hasta la ultima fila,  columna 9 y le hace lista
    columna = {"atractor": lista[0], "numAtractores": lista[2], "tamanio": lista[3:6], "jornada": lista[6:11], "dias": lista[12:22]}
    return columna

def modificarCampoNAtractores(columna: dict):
    #si el numero de atractores es nulo
    print("Hola pata")
    if math.isnan(columna["numAtractores"]):
        hoja=workbook.worksheets[9]
        print(sum(x for x in columna['tamanio'] if not math.isnan(x)))
        hoja.cell(row=1, column=9).value=sum(x for x in columna['tamanio'] if not math.isnan(x))
        workbook.save("../04. Forumularios digitalizados grupo 4.xlsx")
    else:
        return True



leido=pd.read_excel("../04. Forumularios digitalizados grupo 4.xlsx", sheet_name=9)
columna = leerColumna(leido, 8)

modificarCampoNAtractores(columna)
print(workbook)
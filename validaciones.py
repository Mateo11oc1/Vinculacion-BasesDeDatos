
import pandas
import math

"""Devuelve la columna como un diccionario de acuerdo a los parametros"""
def leerColumna(archivo: str, hoja: int, columna: int):
    archivo = pandas.read_excel(archivo, sheet_name = hoja)
    lista = archivo.iloc[7:,columna].tolist()
    columna = {"atractor": lista[0], "numAtractores": lista[2], "tamanio": lista[3:6], "horario": lista[6:11], "dias": lista[12:22]}
    return columna

"""validar que la suma de atractores no sea mayor o 
menor al numero de atractores colocado"""
def validarSuma(columna):
    try:
        if columna["numAtractores"] == sum(x for x in columna["tamanio"] if not math.isnan(x)):
            return True
        else:
            return False
    except TypeError:
        return False
    

"""Validad que solo sean numero y no letras"""
def validarCaracteres(columna: dict):
    if not isinstance(columna["numAtractores"], int):
        return False
    
    for i in list(columna.values())[2:]:
        for j in i:
            if isinstance(j, str) or (not math.isnan(j) and not isinstance(j, int)):
                return False
    
    return True
    
"""Validar que en caso de haber datos en las filas inferiores
, el campo de numero de atractores no sea nulo"""
def validarNaN(columna: dict):
    """Valida que el detalle, las filas desde el tamanio
        no contengan datos"""
    def validarDetalle(columna:dict):
        for i in list(columna.values())[2:]:
            for j in i:
                if not isinstance(j, int):
                    return False
                elif not math.isnan(j):
                    return False
                else: 
                    return True

    if math.isnan(columna["numAtractores"]) and not validarDetalle(columna):
        return False
    else:
        return True
        

columna = leerColumna("../04. Forumularios digitalizados grupo 4.xlsx", 9, 9)

print(validarSuma(columna))

print(validarCaracteres(columna))

print(validarNaN(columna))

#hola pajaro con cola :)
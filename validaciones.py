
import pandas
import math

"""Devuelve la columna como un diccionario de acuerdo a los parametros"""
def leerColumna(archivo: str, hoja: int, columna: int):
    archivo = pandas.read_excel(archivo, sheet_name = hoja)
    lista = archivo.iloc[7:,columna].tolist()
    #[7:,columna] nos da la celda desde la fila 7 hasta la ultima fila,  columna 9 y le hace lista
    columna = {"atractor": lista[0], "numAtractores": lista[2], "tamanio": lista[3:6], "jornada": lista[6:11], "dias": lista[12:22]}
    return columna

"""validar que la suma de atractores no sea mayor o 
menor al numero de atractores colocado"""
def validarSuma(columna):
    try:
        #si es que hay algo en la celda se valida que el numero de atractores sea igual a la suma de los tamanos
        if columna["numAtractores"] == sum(x for x in columna["tamanio"] if not math.isnan(x)):
            return True
        else:
            return False
    except TypeError:
        #la excepecion se da si el dato no es un numero(error de tipos)
        return False
    
"""Validar que la suma de todos los datos en días de atención no sea menor al numero de atractores"""
def validarSumaDias(columna:dict):
    #si el numero de atractores es mayor a la suma de los valores de dias
    if columna["numAtractores"] > sum(x for x in columna['dias'] if not math.isnan(x)):
        return False
    else:
        return True
"""Validar que la suma de todos los datos en jornada no sea menor al numero de atractores"""
def validarSumaJornada(columna:dict):

    #si el numero de atractores es mayor a la suma de los valores del jornada
    if columna["numAtractores"] > sum(x for x in columna['jornada'] if not math.isnan(x)):
        return False
    else:
        return True


"""Validar que solo sean numero y no letras"""
def validarCaracteres(columna: dict):
    #si no es un numero entero
    if not isinstance(columna["numAtractores"], int):
        return False
    #se recorre la lista de valores de la columna desde 2 en adelante porque va desde el tamanio
    for i in list(columna.values())[2:]:
        #se recorre cada lista, porque hay la lista tamanio, jornada, dias
        for j in i:
            #si el valor de la celda es un string o esta vacio o es un numero decimal
            if isinstance(j, str) or (not math.isnan(j) and not isinstance(j, int)):
                return False
    
    return True


#sirve para validar que no se ingresen numeros muy grandes o muy pequenios o de punto flotante
#incluye a los numeros negativos
def validarExtremos(columna:dict):
    for i in list(columna.values())[2:]:
        for j in i:
            if not math.isnan(j):
                if j>1000 or j<1 or isinstance(j,float):
                    return False
    return True

"""Validar que en caso de haber datos en las filas inferiores
, el campo de numero de atractores no sea nulo"""
def validarNaN(columna: dict):
 #  Valida que el detalle, las filas desde el tamanio hacaia abajo no contengan datos

    def validarDetalle(columna:dict):
        #recorre la columna desde la fila 2 hacia abajo
        for i in list(columna.values())[2:]:
            for j in i:
                #si el valor de la celda no es entero, retorna falso
                if not isinstance(j, int):
                    return False
                #si no esta vacio retorna falso
                elif not math.isnan(j):
                    return False
                else: 
                    return True

    #si el numero de atractores es nulo y el detalle no esta vacio
    if math.isnan(columna["numAtractores"]) and not validarDetalle(columna):
        return False
    else:
        return True
        

columna = leerColumna("../04. Forumularios digitalizados grupo 4.xlsx", 9, 9)


print(validarSuma(columna))
print(validarCaracteres(columna))
print(validarNaN(columna))

print(validarExtremos(columna))
print(validarSumaJornada(columna))
print(validarSumaDias(columna))


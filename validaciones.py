import os
import glob
import pandas
import time
import math
import logging
class Validaciones:
    
    def __init__(self):
        self.carpetaExcel = "../"
        self.archivos_excel = self.leerCarpeta() 

    #Obtengo una lista de todos lo archivos excel
    def leerCarpeta(self):
        # Obtener todos los archivos en la carpeta que tengan la extensión .xlsx
        archivos_excel = (archivo for archivo in  glob.glob(os.path.join(self.carpetaExcel, '*.xlsx')) if not os.path.basename(archivo).startswith("~$"))
        #Se filtran los archivos temporales
        return archivos_excel
        
    #Devuelve la columna como un diccionario de acuerdo a los parametros
    def leerColumna(self):
        self.listaColumnas = []
        
        #Recorro todos los archivos
        for i in self.archivos_excel:
            #Leo todas las hojas de una vez del documento
            leido = pandas.read_excel(i, sheet_name = None)
            numHoja = 0
            #Recorro cada hoja
            for j in leido.values():
                #Recorro cada columna
                #shape[1] nos da el numero de columnas de la hoja
                for h in range(2, j.shape[1]):
                    #Desde la fila 7 en adelante
                    lista = j.iloc[7:, h].values.tolist()
                    columna = {"atractor": lista[0], "numAtractores":lista[2], "tamanio": lista[3:6], "jornada": lista[6:11], 
                            "dias": lista[12:22], "numColumna": h, "hoja": numHoja, "archivo": i[3:], "vacia":False, "listaErrores":[]}
                    columna = self.validarColVacia(columna) #Se valida que la columna esta vacia al leer
                    print(columna)
                    self.listaColumnas.append(columna)
                    
                numHoja += 1
        x=4
                
    def validarColVacia(self, columna: dict):
        def listaVacia(lista):
            for i in lista:
                if not math.isnan(i):
                    return False
            return True
        #Si hay un TypeError automaticamente no esta vacia
        try:
            #Si todo es NaN, la columna esta vacia
            if math.isnan(columna["numAtractores"]) and listaVacia(columna["tamanio"]) and listaVacia(columna["jornada"]) and listaVacia(columna["dias"]):
                columna["vacia"] = True
                return columna, True
            else:
                columna["vacia"] = False
                return columna, False
        except TypeError:
            columna["vacia"] = False
            return columna, False
        
    #Validar que solo sean numero y no letras
    def validarCaracteres(self, columna: dict):
        #si no es un numero entero
        if not isinstance(columna["numAtractores"], int):
            #Esto reporta en consola como si fuera un error
            logging.error("Numero de atractores no es un entero")
            return False
        #se recorre la lista de valores de la columna desde 2 en adelante porque va desde el tamanio
        for i in list(columna.values())[2:]:
            #se recorre cada lista, porque hay la lista tamanio, jornada, dias
            for j in i:
                #si el valor de la celda es un string, esta vacio o es un numero decimal
                if isinstance(j, str) or (not math.isnan(j) and not isinstance(j, int)):
                    logging.error("El valor no es un numero")
                    return False
        return True
    
    
    
    def validar(self, columna: dict):
        pass
        # columna = {}
        # if self.validarColVacia(columna)[1]:
        #     pass
        # else:
        #     pass
        
        # return columna


validaciones = Validaciones()
validaciones.leerColumna()
# validaciones.validar()

import os
import glob
import pandas
import time
import math

class Validaciones:
    
    def __init__(self):
        self.carpetaExcel = "../"
        self.archivos_excel = self.leerCarpeta() 

    #Obtengo una lista de todos lo archivos excel
    def leerCarpeta(self):
        # Obtener todos los archivos en la carpeta que tengan la extensión .xlsx
        archivos_excel = glob.glob(os.path.join(self.carpetaExcel, '*.xlsx'))

        #Filtras los archivos temporales
        return [archivo for archivo in archivos_excel if not os.path.basename(archivo).startswith("~$")]
        
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
                for h in range(2, j.shape[1]):
                    #Desde la columna 7 en adelante
                    lista = j.iloc[7:, h].tolist()
                    columna = {"atractor": lista[0], "numAtractores":lista[2], "tamanio": lista[3:6], "jornada": lista[6:11], 
                            "dias": lista[12:22], "numColumna": h, "hoja": numHoja, "archivo": i[3:], "vacia":False, "listaErrores":[]}
                    self.listaColumnas.append(columna)
                    
                numHoja += 1
                
                
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
                return columna
            else:
                columna["vacia"] = False
                return columna
        except TypeError:
            columna["vacia"] = False
            return columna
        
    def validar(self):
        for i in range(len(self.listaColumnas)):
            print("Hola: ", self.listaColumnas[i])
            self.listaColumnas[i] = self.validarColVacia(self.listaColumnas[i])
            print(self.listaColumnas[i])
        


validaciones = Validaciones()
validaciones.leerColumna()
validaciones.validar()

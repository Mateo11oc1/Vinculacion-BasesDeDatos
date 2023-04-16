import os
import glob
import pandas
import time

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
        lista1 = []
        
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
                    columna = {"atractor": lista[0], "numAtractores":lista[2], "tamanio": lista[3:6], "jornada": lista[6:11], "dias": lista[12:22], "numColumna": h, "hoja": numHoja}
                    lista1.append(columna)
                    print(columna)
                    
                numHoja += 1
validaciones = Validaciones()
validaciones.leerColumna()
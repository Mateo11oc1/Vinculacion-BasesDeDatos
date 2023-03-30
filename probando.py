import xlrd
import pandas as pd


archivoPandas = pd.read_excel("../04. Forumularios digitalizados grupo 4.xlsx", sheet_name="1")

print(archivoPandas.iloc[6, 8]) #la fila 9 es donde dice atractores
# print(archivoPandas)

# a2 = xlrd.open_workbook_xls("../prueba.xlsx")
# print(a2)
#Autora: Ana Lopez 
#Este codigo es para poder probar si se instalo correctamente la libreria openpy xls
#Usar este comando para instalar la biblioteca desde la terminal: pip install openpyxl
from openpyxl import Workbook

libro = Workbook()
libro.save("prueba_openpyxl.xlsx")
print("✅ Archivo Excel creado exitosamente")
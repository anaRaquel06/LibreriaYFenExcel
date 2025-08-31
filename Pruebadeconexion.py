"""
@author: Ana López
"""
#Pruba de conexion 
# Necesitamos intalar una biblioteca nueva en la terminal que se debe instalar con 
#pip install mysql-connector-python 
#este codigo imprime todas las filas de nuestro codigo 
import mysql.connector #es la que instalamos anteriormente   
import pandas as pd 

#creamos una conexion 
#poner usuario y contraseña en usser y pasword 
conexion= mysql.connector.connect(user='root',password='1234',
                                  host='localhost',  # PUERTO
                                  database='importar',# Nombre de la Base de datos
                                  port=3306) 
print("✅ Conexión exitosa:", conexion)

# Crear cursor y ejecutar consulta
cursor = conexion.cursor()
query = "SELECT * FROM historico_2024"
cursor.execute(query)
resultados = cursor.fetchall()

# Obtener nombres de columnas
columnas = [col[0] for col in cursor.description]

# Convertir a DataFrame
df_2024 = pd.DataFrame(resultados, columns=columnas)

# Mostrar todas las filas del DataFrame
pd.set_option('display.max_rows', None)
print(df_2024)

# Cerrar conexión
conexion.close()
print("🔒 Conexión cerrada correctamente.")


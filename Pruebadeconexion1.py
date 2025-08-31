"""
@author: Ana López
"""
# Requiere instalar la biblioteca con:
# pip install mysql-connector-python

import mysql.connector
import pandas as pd

# Crear conexión a la base de datos
conexion = mysql.connector.connect(
    user='root',
    password='1234',
    host='localhost',
    database='importar',
    port=3306
)

print("✅ Conexión exitosa:", conexion)

# Crear cursor y ejecutar consulta
cursor = conexion.cursor()
query = "SELECT * FROM historico_2024"
cursor.execute(query)
resultados = cursor.fetchall()

# Obtener nombres de columnas las estamos sacando 
columnas = [col[0] for col in cursor.description]

# Convertir a DataFrame
df_2024 = pd.DataFrame(resultados, columns=columnas)

# Mostrar solo las primeras filas (por defecto pandas muestra 5)
print(df_2024.head())

# Cerrar conexión
conexion.close()
print("🔒 Conexión cerrada correctamente.")
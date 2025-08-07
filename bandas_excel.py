# Autor: David Rey 
# Modificación: Ana López 
# Mi idea es que el archivo Excel "historico_ult_mes.xlsx" se pueda usar 
# en los demás códigos en lugar de Yahoo Finance 
# Hice la prueba con el código que más me familiaricé (gracias a David): 
# Las Bandas de Bollinger con unas pequeñas modificaciones

import pandas as pd #traje pandas para poder leer nuestro archivo excel
import matplotlib.pyplot as plt

def bandas_bollinger_desde_excel(tickers, nombre_archivo, window=20):
    if isinstance(tickers, str):
        tickers = [tickers]

    # Con esta linea vamos a poder leer el archivo 
    df = pd.read_excel(nombre_archivo)
    df['Fecha'] = pd.to_datetime(df['Fecha'])

    datos = {}

    for ticker in tickers:
        data = df[df['Ticker'] == ticker].sort_values('Fecha').copy()

        if data.empty: #en caso de no encontrar los tickets se avisara 
            print(f"⚠️ No hay datos para {ticker}") #otras opciones por si quieren personalizar 🚨🚧⛔📢💣 
            continue

        # Renombramos por que necesitabamos la columna cierre 
        if 'Cierre' in data.columns and 'Close' not in data.columns:
            data['Close'] = data['Cierre'] #

        # Usar la fecha como índice para gráficas
        data.set_index('Fecha', inplace=True)

        # Calcular medias y desviaciones
        data['MA20'] = data['Close'].rolling(window=window).mean()
        data['std_dev'] = data['Close'].rolling(window=window).std()

        # Calcular bandas
        data['UpperBand'] = data['MA20'] + 2 * data['std_dev']
        data['MiddleBand'] = data['MA20']
        data['LowerBand'] = data['MA20'] - 2 * data['std_dev']

        datos[ticker] = data

        # Graficar
        plt.figure(figsize=(10, 6))
        plt.plot(data.index, data['Close'], label='Precio de cierre', linewidth=1.5)
        plt.plot(data.index, data['UpperBand'], linestyle='--', linewidth=1, label='Banda Superior', color="purple")
        plt.plot(data.index, data['MiddleBand'], linestyle='--', linewidth=1, label='Banda Media', color="black")
        plt.plot(data.index, data['LowerBand'], linestyle='--', linewidth=1, label='Banda Inferior', color="purple")

        plt.title(f'Bandas de Bollinger - {ticker}')
        plt.xlabel('Fecha')
        plt.ylabel('Precio')
        plt.legend()
        plt.tight_layout()
        plt.show()

    return datos

# Ejemplo de uso:
tickers = ['TSLA', 'GOOGL', 'AMZN']
nombre_archivo = "historico_ult_mes.xlsx"
bandas_bollinger_desde_excel(tickers, nombre_archivo)


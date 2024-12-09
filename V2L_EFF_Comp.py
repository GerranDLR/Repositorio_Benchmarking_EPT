import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os

# Ruta del archivo Excel y carpeta de destino
ruta_excel = r"C:\Users\SNX6774\OneDrive - Nissan Motor Corporation\Escritorio\ETAS tratado\KONA_EFF_Comparation.xlsx"
ruta_salida = r"C:\Users\SNX6774\OneDrive - Nissan Motor Corporation\Escritorio\ETAS tratado\comparacion_eff"

# Crear carpeta de salida si no existe
os.makedirs(ruta_salida, exist_ok=True)

# Leer el archivo Excel
df = pd.read_excel(ruta_excel, engine='openpyxl')

# Eliminar la columna A (no necesaria)
df = df.iloc[:, 1:]  # Nos quedamos con las columnas de eficiencia y la columna B de potencia

# Obtener los valores de la columna B (potencia en kW) como eje X
x_values = df.iloc[:, 0]  # La columna B es la primera columna ahora
x_label = "Power (kW)"

# Asegurarnos de que 0.1 kW esté en el conjunto de datos si no está presente
if 0.1 not in x_values.values:
    x_values = np.append([0.1], x_values)  # Agregar 0.1 si no está
    # Insertamos una fila con NaN para las eficiencias asociadas al 0.1 kW
    empty_row = np.nan * np.ones(df.shape[1])
    df.loc[-1] = empty_row
    df.index = df.index + 1  # Reindexar
    df.loc[0, df.columns[1:]] = np.nan  # Establecer NaN para las eficiencias

# Convertir eficiencias a porcentaje (multiplicamos por 100)
df.iloc[:, 1:] *= 100  # Convertir todas las eficiencias a porcentaje

# Crear el gráfico
plt.figure(figsize=(10, 6))

# Colores personalizados para las curvas (pares de colores similares)
colores = [
    ['#1f77b4', '#aec7e8'],  # Dos tonos de azul
    ['#2ca02c', '#98df8a'],  # Dos tonos de verde
    ['#9467bd', '#c5b0d5']   # Dos tonos de violeta
]

# Graficar las curvas de las columnas de eficiencia (desde la columna C hacia adelante)
for i, col in enumerate(df.columns[1:]):  # Excluir la columna B (potencia)
    color_par = colores[i // 2]  # Asignar el par de colores basado en el índice
    
    # Excluir la fila 4 (índice 3) al graficar la curva
    if col in ['G', 'H']:  # Especificamos las columnas G y H
        # Graficar las curvas excluyendo la fila 4 (índice 3)
        plt.plot(x_values.drop(3), df[col].drop(3), marker='o', label=col, color=color_par[i % 2])  
        
        # Agregar los puntos de la fila 4 como puntos rojos (se visualizan de forma independiente)
        plt.scatter(x_values.iloc[3], df[col].iloc[3], color='red', s=100, zorder=5, label=f'{col} Fila 4')
    else:
        # Para las demás columnas, graficamos como una curva normal
        plt.plot(x_values, df[col], marker='o', label=col, color=color_par[i % 2])

# Ajustar límites y etiquetas del gráfico
plt.xlim(min(x_values), max(x_values))  # El eje X se ajusta automáticamente a los valores de la columna B
plt.ylim(0, 100)  # Rango de eficiencia de 0% a 100%

# Ajustar el eje X para que vaya de 0,5 en 0,5 kW
ticks_x = np.arange(0, max(x_values) + 0.5, 0.5)  # Comenzamos en 0.1 kW
plt.xticks(ticks_x)  # Establecer los ticks del eje X

# Dividir el eje Y en 10 intervalos
ticks_y = np.linspace(0, 100, 10)  # Crear 10 intervalos para el eje Y
plt.yticks(ticks_y)  # Establecer los ticks del eje Y

# Etiquetas y título
plt.xlabel(x_label)
plt.ylabel("Efficiency (%)")
plt.title("Vehicle Efficiency Comparison")

# Agregar leyenda y moverla a la parte inferior derecha
plt.legend(title="Vehicle", loc='lower right', bbox_to_anchor=(1.0, 0.0))

# Mostrar cuadrícula ajustada a los puntos del gráfico
plt.grid(True, which='both', axis='both', linestyle='--', color='gray', alpha=0.7)

# Ajustar el espacio alrededor del gráfico para separar las curvas del eje Y
plt.subplots_adjust(left=0.1, right=0.9, top=0.9, bottom=0.1)

# Guardar el gráfico en la carpeta correspondiente
nombre_archivo = "Efficiency Comparison Between Vehicles.png"
plt.savefig(os.path.join(ruta_salida, nombre_archivo), bbox_inches='tight')

# Mostrar el gráfico
plt.show()

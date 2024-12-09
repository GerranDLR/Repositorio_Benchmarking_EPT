import os
import pandas as pd
import matplotlib.pyplot as plt

# Ruta del archivo Excel y carpeta de salida
ruta_excel = r"C:\Users\SNX6774\OneDrive - Nissan Motor Corporation\Escritorio\ETAS tratado\ETAS_Kona_EXCEL_V2L_POT.xlsx"
carpeta_salida = r"C:\Users\SNX6774\OneDrive - Nissan Motor Corporation\Escritorio\ETAS tratado\Pot_DCDC"

# Crear carpeta de salida si no existe
os.makedirs(carpeta_salida, exist_ok=True)

# Leer el archivo Excel y obtener las hojas
excel = pd.ExcelFile(ruta_excel)
hojas = excel.sheet_names

# Procesar cada hoja
for hoja in hojas:
    # Leer los datos de la hoja
    df = excel.parse(sheet_name=hoja, skiprows=2)  # Saltar las dos primeras filas
    df.columns = ['Tiempo', 'Corriente', 'Voltaje', 'Potencia']  # Asignar nombres a las columnas

    # Filtrar datos para el rango de tiempo deseado (150 a 175)
    df_filtrado = df[(df['Tiempo'] >= 150) & (df['Tiempo'] <= 175)]

    # Calcular estadísticas dentro del rango
    max_corriente, mean_corriente = df_filtrado['Corriente'].max(), df_filtrado['Corriente'].mean()
    max_voltaje, mean_voltaje = df_filtrado['Voltaje'].max(), df_filtrado['Voltaje'].mean()
    max_potencia, mean_potencia = df_filtrado['Potencia'].max(), df_filtrado['Potencia'].mean()

    # Crear el gráfico
    fig, ax1 = plt.subplots(figsize=(12, 8))  # Ajustar tamaño para incluir leyenda y estadísticas

    # Gráfico del eje Y izquierdo (Corriente y Voltaje)
    ax1.plot(df['Tiempo'], df['Corriente'], label='Current (A)', color='blue', linewidth=1.5)
    ax1.plot(df['Tiempo'], df['Voltaje'], label='Voltage (V)', color='green', linewidth=1.5)
    ax1.set_xlabel('Time (s)')
    ax1.set_ylabel('Current / Voltage', color='black')
    ax1.tick_params(axis='y', labelcolor='black')
    ax1.grid()

    # Crear un segundo eje Y para la Potencia
    ax2 = ax1.twinx()
    ax2.plot(df['Tiempo'], df['Potencia'], label='Power', color='red', linewidth=1.5)
    ax2.set_ylabel('Power', color='red')
    ax2.tick_params(axis='y', labelcolor='red')

    # Ajustar dinámicamente el rango del eje Y derecho con un margen del 30%
    potencia_max = df['Potencia'].max()
    potencia_min = df['Potencia'].min()
    margen = 0.90 * max(abs(potencia_max), abs(potencia_min))  # Margen del 30% en ambos lados
    ax2.set_ylim(potencia_min - margen, potencia_max + margen)

    # Preparar texto para las estadísticas
    stats_text = (
        f"DCDC_Current (A) (Blue): Máx = {max_corriente:.2f}, Average = {mean_corriente:.2f}\n"
        f"DCDC_Voltage (V) (Green): Máx = {max_voltaje:.2f}, Average = {mean_voltaje:.2f}\n"
        f"DCDC_Power (W) (Red): Máx = {max_potencia:.2f}, Average = {mean_potencia:.2f}"
    )

    # Agregar estadísticas como parte del gráfico
    plt.figtext(0.5, 0.02, stats_text, wrap=True, horizontalalignment='center', fontsize=10, color='black')

    # Agregar leyenda unificada dentro del gráfico
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper center', bbox_to_anchor=(0.5, -0.1), ncol=3, frameon=False)

    # Título del gráfico
    plt.title(f'Test graph: {hoja}')

    # Guardar el gráfico en la carpeta de salida
    ruta_grafico = os.path.join(carpeta_salida, f'{hoja}.png')
    plt.tight_layout(rect=[0, 0.15, 1, 1])  # Deja espacio para estadísticas debajo del gráfico
    plt.savefig(ruta_grafico)
    plt.close()

print(f"Gráficos generados y guardados en {carpeta_salida}")

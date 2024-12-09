import pandas as pd 
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import os
import numpy as np
from matplotlib.patches import Arc
def plot_and_save(df, signals, output_folder, title_prefix, x_col='Time', y_range=None, colors=None, ylabel='Valor'):
    """Genera gráficos para las señales especificadas y guarda los gráficos en la carpeta indicada."""
    file_name = os.path.splitext(os.path.basename(df_file_path))[0]
    output_subfolder = os.path.join(output_folder, file_name)
    os.makedirs(output_subfolder, exist_ok=True)

    fig_width, fig_height = 8.28, 2.5  # Tamaño del gráfico

    if len(signals) > 1:
        plt.figure(figsize=(fig_width, fig_height))
        ax1 = plt.gca()
        handles, labels = [], []

        for signal in signals:
            color = colors.get(signal, None) if colors else None
            if signal == 'AveIrms1':  # Eje secundario
                ax2 = ax1.twinx()
                line2, = ax2.plot(df[x_col], df[signal], label=signal, color='blue')
                ax2.set_ylabel('AveIrms1', color='blue')
                ax2.tick_params(axis='y', labelcolor='blue')
                ax2.set_ylim(0, 20)
                handles.append(line2)
                labels.append(signal)
            else:
                line1, = ax1.plot(df[x_col], df[signal], label=signal, color=color)
                y_label_color = 'green' if signal == 'AveUrms1' else 'black'
                ax1.set_ylabel(ylabel, color=y_label_color)
                ax1.tick_params(axis='y', labelcolor=y_label_color)
                handles.append(line1)
                labels.append(signal)

        if 'AveUrms1' in signals:
            ax1.set_ylim(0, 240)
            ax1.yaxis.set_major_locator(ticker.MultipleLocator(60))

        ax1.set_xlabel('Time (s)')
        plt.title(title_prefix)

        # Configuración de la cuadrícula
        ax1.grid(True, which='major', axis='both')  # Cuadrícula horizontal y vertical en el eje principal
        if 'AveIrms1' in signals:
            ax2.grid(False)  # Evitar cuadrícula en el eje secundario

        ax1.xaxis.set_major_locator(ticker.MaxNLocator(integer=True))
        ax1.xaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{int(x)}'))

        plt.legend(handles=handles, labels=labels, loc='upper left', bbox_to_anchor=(1, 1))
        plt.savefig(os.path.join(output_subfolder, f'{title_prefix}.png'), bbox_inches='tight')
        plt.close()
    else:
        for signal in signals:
            plt.figure(figsize=(fig_width, fig_height))
            color = colors.get(signal, None) if colors else None
            plt.plot(df[x_col], df[signal], label=signal, color=color)
            plt.xlabel('Time')
            plt.ylabel(signal if ylabel is None else ylabel)
            y_label_color = 'green' if signal == 'AveUrms1' else 'black'
            plt.ylabel(signal, color=y_label_color)
            plt.tick_params(axis='y', labelcolor=y_label_color)
            plt.title(f'{title_prefix}: {signal}')
            plt.grid(True, which='major', axis='both')  # Mostrar ambas cuadrículas
            ax = plt.gca()
            ax.xaxis.set_major_locator(ticker.MaxNLocator(integer=True))
            if y_range:
                plt.ylim(y_range)
            plt.legend(loc='upper left', bbox_to_anchor=(1, 1))
            plt.savefig(os.path.join(output_subfolder, f'{title_prefix}_{signal}.png'), bbox_inches='tight')
            plt.close()




def plot_fasorial(df, df_file_path, output_folder, title_prefix):  
    """Genera una gráfica fasorial de los vectores de potencia y corriente y la guarda en la carpeta indicada."""

    # Crear carpeta para los gráficos del archivo CSV
    file_name = os.path.splitext(os.path.basename(df_file_path))[0]
    output_subfolder = os.path.join(output_folder, file_name)
    os.makedirs(output_subfolder, exist_ok=True)

    # Obtener los valores máximos de P y S
    P_max = df['AveP1'].max()
    S_max = df['AveS1'].max()

    # Calcular el ángulo en radianes usando la fórmula cos(θ) = P / S
    cos_theta = P_max / S_max
    angle_rad = np.arccos(cos_theta)  # Ángulo en radianes
    angle_deg = np.degrees(angle_rad)  # Ángulo en grados

    # Definir módulos para vectores
    V_mag = P_max  # Módulo del voltaje (proporcional a P)
    I_mag = S_max  # Módulo de la corriente (proporcional a S)

    # Coordenadas de los vectores en el plano
    V_x = V_mag  # Voltaje en el eje real
    V_y = 0

    I_x = I_mag * np.cos(angle_rad)  # Componente en el eje real
    I_y = I_mag * np.sin(angle_rad)  # Componente en el eje imaginario

    # Graficar el gráfico vectorial
    plt.figure(figsize=(8, 8))

    # Añadir el círculo unitario como cuadrícula
    circle = plt.Circle((0, 0), max(V_mag, I_mag), color='black', fill=False, linestyle='-', linewidth=1)
    plt.gca().add_patch(circle)

    # Graficar los vectores
    plt.quiver(0, 0, V_x, V_y, angles='xy', scale_units='xy', scale=1, color='green', label='Voltage (V)', linewidth=2)
    plt.quiver(0, 0, I_x, I_y, angles='xy', scale_units='xy', scale=1, color='blue', label='Current (A)', linewidth=2)

    # Añadir el factor de potencia (FP) como un label, a la derecha, a la altura de la leyenda
    plt.text(max(V_mag, I_mag) * 1.3, 1.4 * max(V_mag, I_mag), f'FP = {cos_theta:.4f}', color='purple', fontsize=12, va='center', ha='left')

    # Dibujar el arco que representa el ángulo entre los vectores
    arc_radius = 0.5 * max(V_mag, I_mag)  # Radio del arco
    arc = Arc((0, 0), 2 * arc_radius, 2 * arc_radius, theta1=0, theta2=angle_deg, color='red', lw=2, linestyle='--')
    plt.gca().add_patch(arc)

    # Colocar el texto del ángulo en el gráfico (con símbolo °)
    plt.text(0.5 * arc_radius, 0.5 * arc_radius, f'{angle_deg:.2f}°', color='red', fontsize=12, ha='center')

    # **Nueva parte: Obtener el valor máximo de THD desde las gráficas MaxU1_Group1 y MaxU1_Group2**
    # Definir las columnas de MaxU1_Group1 y MaxU1_Group2
    columns_group1 = ['MaxU1(2)', 'MaxU1(4)', 'MaxU1(6)', 'MaxU1(8)', 'MaxU1(10)']
    columns_group2 = ['MaxU1(3)', 'MaxU1(5)', 'MaxU1(7)', 'MaxU1(9)']

    # Asegurarse de que las columnas existen en el DataFrame
    df[columns_group1] = df[columns_group1].apply(pd.to_numeric, errors='coerce')
    df[columns_group2] = df[columns_group2].apply(pd.to_numeric, errors='coerce')

    # Encontrar el valor máximo entre las dos gráficas
    max_value_group1 = df[columns_group1].max().max()  # Máximo entre las columnas de Group1
    max_value_group2 = df[columns_group2].max().max()  # Máximo entre las columnas de Group2

    # El valor de THD será el máximo entre ambos grupos
    max_thd_value = max(max_value_group1, max_value_group2)

    # Agregar el label de THD al gráfico
    plt.text(max(V_mag, I_mag) * 1.3, 1.3 * max(V_mag, I_mag), f'THD = {max_thd_value:.2f}%', color='orange', fontsize=12, va='center', ha='left')

    # Ajustes visuales
    # Quitar el recuadro negro
    plt.gca().set_facecolor('white')  # Eliminar fondo negro
    plt.axis('off')  # Desactivar los ejes

    # Configurar el aspecto del gráfico (círculo)
    plt.gca().set_aspect('equal', adjustable='box')

    # Título y etiquetas
    plt.title(f'Fasorial: Voltage & Current - Angle: {angle_deg:.2f}°', va='bottom', fontsize=14)
    
    # Aquí colocamos los valores de los ejes al lado
    plt.text(1.05 * max(V_mag, I_mag), 0, '(0° - 360°) Real axis ', color='black', fontsize=12, ha='left', va='center')
    # Mover el eje Y un poco más hacia arriba
    plt.text(0, 1.2 * max(V_mag, I_mag), 'Imaginary axis (90°)', color='black', fontsize=12, ha='center', va='bottom')

    # Añadir la leyenda debajo del label de THD
    plt.legend(loc='upper right', fontsize=12, bbox_to_anchor=(1.2, 0.9))

    # Añadir la cuadrícula circular (en la parte interna del gráfico)
    num_ticks = 8  # Número de divisiones de la cuadrícula
    ticks = np.linspace(0, 1.2 * max(V_mag, I_mag), num_ticks)
    plt.xticks(ticks)
    plt.yticks(ticks)

    # Añadir círculos concéntricos en la cuadrícula
    for radius in ticks:
        circle_inner = plt.Circle((0, 0), radius, color='gray', fill=False, linestyle='--', linewidth=0.5)
        plt.gca().add_patch(circle_inner)

    # Añadir líneas radiales para los cuadrantes
    for angle in np.linspace(0, 360, 8, endpoint=False):
        radian = np.radians(angle)
        plt.plot([0, max(V_mag, I_mag) * np.cos(radian)], [0, max(V_mag, I_mag) * np.sin(radian)], color='gray', linestyle='--', linewidth=0.5)

    # Etiquetas para los ángulos en los ejes
    plt.text(0, max(V_mag, I_mag) * 1.05, '90°', color='black', fontsize=12, ha='center', va='bottom')
    plt.text(-max(V_mag, I_mag) * 1.05, 0, '180°', color='black', fontsize=12, ha='right', va='center')
    plt.text(0, -max(V_mag, I_mag) * 1.05, '270°', color='black', fontsize=12, ha='center', va='top')

    # Guardar la gráfica fasorial
    plt.savefig(os.path.join(output_subfolder, f'{title_prefix}_Fasorial.png'), bbox_inches='tight')
    plt.close()

def main(input_folder, output_folder):
    """Lee archivos CSV de la carpeta de entrada y genera gráficos."""
    # Obtener lista de archivos CSV en la carpeta de entrada
    for filename in os.listdir(input_folder):
        if filename.endswith('.csv'):
            global df_file_path
            df_file_path = os.path.join(input_folder, filename)
            
            # Leer el archivo CSV en un DataFrame de pandas
            df = pd.read_csv(df_file_path)
            
            # Colores específicos para las señales
            colors_AveP_S_Q_PF = {
                'AveP1': 'red',
                'AveS1': 'black',
                'AveQ1': 'yellow',
                'AvePF1': 'cyan'
            }
            
            colors_AveU_I = {
                'AveUrms1': 'green',
                'AveIrms1': 'blue'
            }
            # Graficar AveUrms1 y AveIrms1 con colores específicos y un eje secundario para AveIrms1
            plot_and_save(df, ['AveUrms1', 'AveIrms1'], output_folder, 'AveU_I', colors=colors_AveU_I, ylabel='Voltage (V)')
            
            # Graficar AveP1, AveS1, AveQ1, AvePF1 en un solo gráfico con colores específicos y Power (W, VA, VAR) en la etiqueta del eje y
            plot_and_save(df, ['AveP1', 'AveS1', 'AveQ1', 'AvePF1'], output_folder, 'AveP_S_Q_PF', colors=colors_AveP_S_Q_PF, ylabel='Power (W, VA, VAR)')

            # Graficar MaxUthd1 y MaxU1(2) a MaxU1(10) con y_range (0, 1.5)
            plot_and_save(df, ['MaxUthd1'] + [f'MaxU1({i})' for i in range(2, 11)], output_folder, 'MaxU', y_range=(0, 1.5), ylabel='THD (%)')

            # Graficar MaxU1(2), MaxU1(4), MaxU1(6), MaxU1(8), MaxU1(10) con y_range automático
            plot_and_save(df, ['MaxU1(2)', 'MaxU1(4)', 'MaxU1(6)', 'MaxU1(8)', 'MaxU1(10)'], output_folder, 'MaxU1_Group1', ylabel='THD (%)')

            # Graficar MaxU1(3), MaxU1(5), MaxU1(7), MaxU1(9) con y_range automático
            plot_and_save(df, ['MaxU1(3)', 'MaxU1(5)', 'MaxU1(7)', 'MaxU1(9)'], output_folder, 'MaxU1_Group2', ylabel='THD (%)')
            # Graficar la gráfica fasorial
            plot_fasorial(df, df_file_path, output_folder, 'Fasorial')  # Aquí incluimos df_file_path


# Llamar a la función principal para procesar los archivos CSV
input_folder = r'C:\Users\SNX6774\OneDrive - Nissan Motor Corporation\Escritorio\Home_matrix\V2L\Potencias_harmonicos_v2l\Archivos_csv'
output_folder = r'C:\Users\SNX6774\OneDrive - Nissan Motor Corporation\Escritorio\Home_matrix\V2L\Potencias_harmonicos_v2l\Graficos'
main(input_folder, output_folder)

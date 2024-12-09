import pandas as pd 
import matplotlib.pyplot as plt
import numpy as np
import os
from PIL import Image

def extract_data_from_csv(csv_file):
    df = pd.read_csv(csv_file, header=None, names=['Value1', 'Value2'])
    df['Value1'] = pd.to_numeric(df['Value1'], errors='coerce')
    df['Value2'] = pd.to_numeric(df['Value2'], errors='coerce')
    print(f"Datos leídos del archivo CSV: {csv_file}")
    print(df)
    return df

def plot_data(df, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    length = len(df)
    time = np.arange(length)  # Crear un eje de tiempo basado en el número de muestras
    
    # Cálculos para el gráfico de Value1
    max_value1 = np.max(df['Value1'])
    min_value1 = np.min(df['Value1'])
    mean_max_value1 = np.mean([max_value1])
    
    # Gráfico de Value1
    fig, ax = plt.subplots(figsize=(6.28, 2.039))
    ax.plot(time, df['Value1'], color='green')
    ax.set_title('Voltage waveform', fontsize=8)
    ax.set_xlabel('Time (s)', fontsize=8)
    ax.set_ylabel('Voltage (V)', fontsize=8)
    ax.axhline(y=0, color='black', linewidth=0.8)
    ax.tick_params(axis='x', which='both', bottom=False, top=False, labelbottom=False)
    ax.tick_params(axis='y', labelsize=6)
    ax.yaxis.set_major_locator(plt.MaxNLocator(8))
    
    textstr = f'Max: {format_decimal(max_value1, 1)}V\nMin: {format_decimal(min_value1, 1)}V\nMean Max: {format_decimal(mean_max_value1, 3)}V'
    props = dict(boxstyle='round', facecolor='wheat', alpha=0.5)
    ax.text(0.02, 0.05, textstr, transform=ax.transAxes, fontsize=5.8, verticalalignment='bottom', bbox=props)
    
    logo_path = r'C:\Users\SNX6774\OneDrive - Nissan Motor Corporation\Escritorio\Home_matrix\Logo_Nissan.png'
    logo = Image.open(logo_path)
    logo_resized = logo.resize((90, 50))
    fig.figimage(logo_resized, xo=fig.bbox.xmax - -920, yo=fig.bbox.ymax - -260, alpha=0.5, zorder=1)
    
    voltage_waveform_path = os.path.join(output_dir, 'voltage_wave.png')
    plt.savefig(voltage_waveform_path, dpi=300)
    plt.close(fig)
    
    # Cálculos para el gráfico de Value2
    max_value2 = np.max(df['Value2'])
    min_value2 = np.min(df['Value2'])
    mean_max_value2 = np.mean([max_value2])
    
    # Gráfico de Value2
    fig, ax = plt.subplots(figsize=(6.28, 2.039))
    ax.plot(time, df['Value2'], color='blue')
    ax.set_title('Current waveform', fontsize=8)
    ax.set_xlabel('Time (s)', fontsize=8)
    ax.set_ylabel('Current (A)', fontsize=8)
    ax.axhline(y=0, color='black', linewidth=0.8)
    ax.tick_params(axis='x', which='both', bottom=False, top=False, labelbottom=False)
    ax.tick_params(axis='y', labelsize=6)
    ax.yaxis.set_major_locator(plt.MaxNLocator(8))
    
    textstr = f'Max: {format_decimal(max_value2, 1)}A\nMin: {format_decimal(min_value2, 1)}A\nMean Max: {format_decimal(mean_max_value2, 3)}A'
    props = dict(boxstyle='round', facecolor='wheat', alpha=0.5)
    ax.text(0.02, 0.05, textstr, transform=ax.transAxes, fontsize=5.8, verticalalignment='bottom', bbox=props)
    
    fig.figimage(logo_resized, xo=fig.bbox.xmax - -920, yo=fig.bbox.ymax - -260, alpha=0.5, zorder=1)
    
    current_waveform_path = os.path.join(output_dir, 'current_wave.png')
    plt.savefig(current_waveform_path, dpi=300)
    plt.close(fig)
    
   # Gráfico combinado: superposición de corriente y voltaje con doble eje Y
    fig, ax1 = plt.subplots(figsize=(6.28, 2.039))

    # Eje izquierdo (Tensión)
    ax1.plot(time, df['Value1'], color='green', label='Voltage (V)')
    ax1.set_xlabel('Time (s)', fontsize=8)
    ax1.set_ylabel('Voltage (V)', fontsize=8, color='green')
    ax1.tick_params(axis='y', labelcolor='green', labelsize=6)
    ax1.axhline(y=0, color='black', linewidth=0.8)
    ax1.tick_params(axis='x', which='both', bottom=False, top=False, labelbottom=False)
    ax1.yaxis.set_major_locator(plt.MaxNLocator(8))

    # Eje derecho (Corriente)
    ax2 = ax1.twinx()
    ax2.plot(time, df['Value2'], color='blue', label='Current (A)', alpha=0.7)
    ax2.set_ylabel('Current (A)', fontsize=8, color='blue')
    ax2.tick_params(axis='y', labelcolor='blue', labelsize=6)
    ax2.set_ylim(-25, 25)  # Establecer el rango de corriente de 0 a 25

    # Título y leyenda
    fig.suptitle('Voltage and Current Waveforms with Phase Shift', fontsize=8)
    fig.legend(loc='upper right', fontsize=6)

    # Agregar logo
    fig.figimage(logo_resized, xo=fig.bbox.xmax - -920, yo=fig.bbox.ymax - -260, alpha=0.5, zorder=1)

    combined_waveform_path = os.path.join(output_dir, 'combined_waveform.png')
    plt.savefig(combined_waveform_path, dpi=300)
    plt.close(fig)
        
    return max_value1, max_value2

def format_decimal(number, decimal_places):
    return f"{number:.{decimal_places}f}"

def sort_and_save_to_excel(df, output_file):
    df_sorted = df[['Value1', 'Value2']].sort_values(by=['Value1'], ascending=False)
    excel_file = output_file
    
    try:
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            df_sorted.to_excel(writer, index=False, sheet_name='Sorted_Data')
        print(f'Archivo Excel generado: {excel_file}')
    except PermissionError:
        print(f"Permission denied: {excel_file}. Por favor, cierra el archivo si está abierto y vuelve a intentarlo.")

def main():
    directory = r'C:\Users\SNX6774\OneDrive - Nissan Motor Corporation\Escritorio\Home_matrix\V2L\Potencia _V2L\Archivos_CSV'
    
    for filename in os.listdir(directory):
        if filename.endswith('.csv'):
            csv_file = os.path.join(directory, filename)
            base_name = os.path.splitext(os.path.basename(csv_file))[0]
            main_dir = os.path.join(directory, base_name)
            output_dir = os.path.join(main_dir, 'Image_waveform')
            output_file = os.path.join(main_dir, 'sorted_values.xlsx')
            
            if not os.path.exists(main_dir):
                os.makedirs(main_dir)
            
            df = extract_data_from_csv(csv_file)
            max_value1, max_value2 = plot_data(df, output_dir)
            sort_and_save_to_excel(df, output_file)
    
            mean_max_value = np.mean([max_value1, max_value2])
            print(f'Mean of max values for {filename}: {mean_max_value:.3f}V/A')

if __name__ == "__main__":
    main()

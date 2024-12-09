import os
import pandas as pd
import matplotlib.pyplot as plt
import logging

# Configuración de logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def create_output_folder(base_path, folder_name):
    """Crea una carpeta de salida si no existe."""
    folder_path = os.path.join(base_path, folder_name)
    os.makedirs(folder_path, exist_ok=True)
    return folder_path

def process_sheet(df, sheet_name, output_folder):
    """Genera gráficos para una hoja específica."""
    sheet_folder = create_output_folder(output_folder, sheet_name)

    # Valores medio y máximo de las curvas
    mean_current_lv = df["ADS1_CH1"].mean()
    max_current_lv = df["ADS1_CH1"].max()
    mean_voltage_lv = df["ADS2_CH3"].mean()
    max_voltage_lv = df["ADS2_CH3"].max()
    mean_dcdc_current = df["ADS1_CH2"].mean()
    max_dcdc_current = df["ADS1_CH2"].max()
    mean_v2l_current = df["ADS1_CH3"].mean()
    max_v2l_current = df["ADS1_CH3"].max()

    # Gráfico 1: ADS1_CH1 vs ADS2_CH3
    generate_plot(
        x=df["Tiempo"], 
        y_list=[df["ADS1_CH1"], df["ADS2_CH3"]],
        labels=["Current (LV)", "Voltage (LV)"],
        colors=["blue", "green"],
        title=f"Graph 1: Current & Voltage vs Time\nMean Current: {mean_current_lv:.2f} A | Max Current: {max_current_lv:.2f} A\nMean Voltage: {mean_voltage_lv:.2f} V | Max Voltage: {max_voltage_lv:.2f} V",
        xlabel="Time (s)",
        ylabel="LV Battery",
        save_path=os.path.join(sheet_folder, "Grafico1.png")
    )

    # Gráfico 2: ADS1_CH1 vs ADS1_CH2
    generate_plot(
        x=df["Tiempo"], 
        y_list=[df["ADS1_CH1"], df["ADS1_CH2"]],
        labels=["Current (LV)", "DCDC_OUT Current"],
        colors=["blue", "darkblue"],
        title=f"Graph 2: LV & DCDC_OUT Current vs Time\nMean LV Current: {mean_current_lv:.2f} A | Max LV Current: {max_current_lv:.2f} A\nMean DCDC Current: {mean_dcdc_current:.2f} A | Max DCDC Current: {max_dcdc_current:.2f} A",
        xlabel="Time (s)",
        ylabel="Current",
        save_path=os.path.join(sheet_folder, "Grafico2.png")
    )

    # Gráfico 3: ADS1_CH3 con eje Y dinámico
    min_y = df["ADS1_CH3"].min()
    max_y = df["ADS1_CH3"].max()
    margin = (max_y - min_y) * 0.1  # Margen del 10% para el rango

    generate_plot(
        x=df["Tiempo"], 
        y_list=[df["ADS1_CH3"]],
        labels=["V2L Adapter"],
        colors=["blue"],
        title=f"Graph 3: V2L Current vs Time\nMean V2L Current: {mean_v2l_current:.2f} A | Max V2L Current: {max_v2l_current:.2f} A",
        xlabel="Time (s)",
        ylabel="V2L Current",
        save_path=os.path.join(sheet_folder, "Grafico3.png"),
        y_limits=(min_y - margin, max_y + margin)  # Eje Y dinámico con margen
    )

def generate_plot(x, y_list, labels, colors, title, xlabel, ylabel, save_path, y_limits=None):
    """Genera un gráfico y lo guarda en un archivo."""
    plt.figure()
    for y, label, color in zip(y_list, labels, colors):
        plt.plot(x, y, label=label, color=color)
    plt.title(title)
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    if y_limits:
        plt.ylim(y_limits)  # Configurar rango del eje Y
    plt.legend()
    plt.grid()
    plt.savefig(save_path)
    plt.close()

def main():
    # Ruta del archivo Excel
    excel_path = r"C:\Users\SNX6774\OneDrive - Nissan Motor Corporation\Escritorio\ETAS tratado\Etas_kona_excel_V2L.xlsx"

    # Ruta de salida
    output_base_path = create_output_folder(
        r"C:\Users\SNX6774\OneDrive - Nissan Motor Corporation\Escritorio\ETAS tratado",
        "señales tratadas"
    )

    # Verifica si el archivo Excel existe
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"El archivo Excel no existe: {excel_path}")

    # Leer todas las hojas del archivo Excel
    excel_data = pd.ExcelFile(excel_path)

    for sheet_name in excel_data.sheet_names:
        logging.info(f"Procesando hoja: {sheet_name}")
        df = excel_data.parse(sheet_name, skiprows=2)
        
        # Asegúrate de que la hoja no esté vacía
        if df.empty:
            logging.warning(f"La hoja '{sheet_name}' está vacía. Se omite.")
            continue

        # Renombrar las columnas
        df.columns = ["Tiempo", "ADS1_CH1", "ADS1_CH2", "ADS1_CH3", "ADS1_CH4", "ADS2_CH3"]
        
        # Convertir "Tiempo" a valores numéricos y eliminar filas con NaN
        df["Tiempo"] = pd.to_numeric(df["Tiempo"], errors="coerce")
        df.dropna(subset=["Tiempo"], inplace=True)

        # Procesar la hoja actual
        process_sheet(df, sheet_name, output_base_path)

    logging.info("Gráficos generados y guardados en las carpetas correspondientes.")

if __name__ == "__main__":
    main()

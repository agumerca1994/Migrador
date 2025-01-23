import os
import pandas as pd
from datetime import datetime

def obtener_accountnames(filepath):
    df = pd.read_excel(filepath, usecols=[1])  # Columna B es el índice 1
    accountnames = df.iloc[1:, 0].dropna().tolist()  # Desde la fila 2
    return accountnames

def obtener_datos_sim(filepath):
    df = pd.read_excel(filepath)
    return df.iloc[:, 0].tolist(), df.iloc[:, 1].tolist(), df.iloc[:, 2].tolist(), df.iloc[:, 3].tolist()

def generar_archivo_migracion(accountnames, sim_data, output_path):
    sim_product_type, sim_type, esim_form_factor, packaging_size = sim_data

    # Crear listas para cada columna del DataFrame de salida
    accountnames_expanded = []
    sim_product_type_expanded = []
    sim_type_expanded = []
    esim_form_factor_expanded = []
    packaging_size_expanded = []

    # Expandir los datos SIM para cada accountname
    for accountname in accountnames:
        accountnames_expanded.extend([accountname] * len(sim_product_type))
        sim_product_type_expanded.extend(sim_product_type)
        sim_type_expanded.extend(sim_type)
        esim_form_factor_expanded.extend(esim_form_factor)
        packaging_size_expanded.extend(packaging_size)

    data = {
        'Account Name (M)': accountnames_expanded,
        'SIM Product Type Name (M)': sim_product_type_expanded,
        'SIM Type (M)': sim_type_expanded,
        'eSIM Type / Form Factor (M)': esim_form_factor_expanded,
        'Minimum Packaging Size (M)': packaging_size_expanded,
        'Comment (O)': [''] * len(accountnames_expanded)
    }

    df_output = pd.DataFrame(data)
    df_output.to_excel(output_path, index=False)

def main():
    import sys
    if len(sys.argv) > 1:
        ruta_archivo_seleccionado = sys.argv[1]  # Archivo seleccionado por la interfaz
    else:
        print("Por favor, proporcione la ruta del archivo seleccionado desde la interfaz de usuario como argumento.")
        return

    directorio_base = "C:\\Migracion"
    
    # Extraer solo el nombre del archivo sin la ruta completa
    excel_file_name = os.path.splitext(os.path.basename(ruta_archivo_seleccionado))[0]
    
    directorio_roles = os.path.join(directorio_base, "SIM Product Type", "Template de migracion", excel_file_name)
    directorio_sim = os.path.join(directorio_base, "SIM Product Type")

    if not os.path.exists(directorio_roles):
        os.makedirs(directorio_roles)

    accountnames = obtener_accountnames(ruta_archivo_seleccionado)

    archivo_sim = os.path.join(directorio_sim, "bd.xlsx")
    if not os.path.exists(archivo_sim):
        print(f"El archivo {archivo_sim} no existe.")
        return

    sim_data = obtener_datos_sim(archivo_sim)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_output = f"SIM_Product_Type_{timestamp}.xlsx"
    ruta_output = os.path.join(directorio_roles, nombre_output)

    generar_archivo_migracion(accountnames, sim_data, ruta_output)
    print(f"Archivo generado exitosamente en: {ruta_output}")

if __name__ == "__main__":
    main()
import os
import pandas as pd
from datetime import datetime

def generar_archivo_pm(ruta_archivo_seleccionado):
    # Leer AccountName desde el archivo seleccionado
    datos_seleccionados = pd.read_excel(ruta_archivo_seleccionado)
    account_names = datos_seleccionados.iloc[1:, 1].dropna().tolist()  # Columna B desde fila 2

    # Leer datos del archivo Roles.xlsx
    ruta_bd = f"../Price Model/bd.xlsx"
    datos_pm = pd.read_excel(ruta_bd)

    # Preparar el nuevo DataFrame
    columnas = ["Account (M)", "Price Model Type (M)", "Name (M)", "Currency (M)", "Zone (M)", "External ID (O)", "Service Plan Type (M)", "Unit (M)", "GL Code (O)", "Price Rating Type (M)", "Rate per sim Flat (M)", "Price Rating Type (M)", "Tiered (M)", "Price (M)"]
    nuevo_datos = []

    for account in account_names:
        for _, fila in datos_pm.iterrows():
            nuevo_datos.append({
                "Account (M)": account,
                "Price Model Type (M)": fila["Price Model Type"],
                "Name (M)": account,
                "Currency (M)": fila["Currency"],
                "Zone (M)": fila["Zone"],
                "Service Plan Type (M)": fila["Service Plan Type"],
                "Unit (M)": fila["Unit"],
                "Price Rating Type (M)": fila["Price Rating Type1"],
                "Price Rating Type (M)": fila["Price Rating Type2"],
                "Tiered (M)": fila["Tiered1"],
                "Tiered (M)": fila["Tiered2"],
                "Price (M)": fila["Price"]
            })

    nuevo_df = pd.DataFrame(nuevo_datos, columns=columnas)

    # Crear el nuevo archivo con timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    # Extraer solo el nombre del archivo sin la ruta completa
    excel_file_name = os.path.splitext(os.path.basename(ruta_archivo_seleccionado))[0]
    
    # Directorio de salida
    directorio_salida = f"../Price Model/Template de migracion/{excel_file_name}/"
    os.makedirs(directorio_salida, exist_ok=True)
    
    archivo_salida = os.path.join(directorio_salida, f"Price_Model_{timestamp}.xlsx")
    nuevo_df.to_excel(archivo_salida, index=False)

    print(f"Archivo generado exitosamente: {archivo_salida}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        ruta_archivo = sys.argv[1]
        generar_archivo_pm(ruta_archivo)
    else:
        print("Por favor, proporcione la ruta del archivo seleccionado desde la interfaz de usuario como argumento.")
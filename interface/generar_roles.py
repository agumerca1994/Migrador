import os
import pandas as pd
from datetime import datetime

def generar_archivo_roles(ruta_archivo_seleccionado):
    # Leer AccountName desde el archivo seleccionado
    datos_seleccionados = pd.read_excel(ruta_archivo_seleccionado)
    account_names = datos_seleccionados.iloc[1:, 1].dropna().tolist()  # Columna B desde fila 2

    # Leer datos del archivo Roles.xlsx
    ruta_roles = r'C:\\Migracion\\Roles\\Roles.xlsx'
    datos_roles = pd.read_excel(ruta_roles)

    # Preparar el nuevo DataFrame
    columnas = ["Account", "Role Name", "Role Description", "Module", "Sub-Module", "Privilegios", "Ver"]
    nuevo_datos = []

    for account in account_names:
        for _, fila in datos_roles.iterrows():
            nuevo_datos.append({
                "Account": account,
                "Role Name": fila["Role Name"],
                "Role Description": fila["Role Description"],
                "Module": fila["Module"],
                "Sub-Module": fila["Sub-Module"],
                "Privilegios": fila["Privilegios"],
                "Ver": fila["Ver"]
            })

    nuevo_df = pd.DataFrame(nuevo_datos, columns=columnas)

    # Crear el nuevo archivo con timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    # Extraer solo el nombre del archivo sin la ruta completa
    excel_file_name = os.path.splitext(os.path.basename(ruta_archivo_seleccionado))[0]
    
    # Directorio de salida
    directorio_salida = rf'C:\\Migracion\\Roles\\Template de migracion\\{excel_file_name}'
    os.makedirs(directorio_salida, exist_ok=True)
    
    archivo_salida = os.path.join(directorio_salida, f"Roles_{timestamp}.xlsx")
    nuevo_df.to_excel(archivo_salida, index=False)

    print(f"Archivo generado exitosamente: {archivo_salida}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        ruta_archivo = sys.argv[1]
        generar_archivo_roles(ruta_archivo)
    else:
        print("Por favor, proporcione la ruta del archivo seleccionado desde la interfaz de usuario como argumento.")
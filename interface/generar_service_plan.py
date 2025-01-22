import os
import pandas as pd
from datetime import datetime
import sys

# Función para obtener nombres de cuentas desde un archivo Excel
def obtener_accountnames(filepath):
    df = pd.read_excel(filepath, usecols=[1])  # Columna B es el índice 1
    accountnames = df.iloc[1:, 0].dropna().tolist()  # Desde la fila 2
    return accountnames

# Función para obtener datos de la base de datos desde un archivo Excel
def obtener_datos_bd(filepath):
    df = pd.read_excel(filepath)
    return df

# Función para generar el archivo de migración
def generar_archivo_migracion(accountnames, bd_data, output_path):
    # Verificar si las columnas necesarias existen en el archivo
    columnas_requeridas = ['Description (M)', 'External ID (M)', 'Service Plan Type (O)', 'Service Types']
    for columna in columnas_requeridas:
        if columna not in bd_data.columns:
            print(f"Advertencia: La columna '{columna}' no existe en el archivo bd.xlsx. Usando valores vacíos.")
            bd_data[columna] = ''

    # Crear listas para cada columna del DataFrame de salida
    accounts = []
    names = []
    descriptions = []
    external_ids = []
    service_plan_types = []
    service_types = []

    # Expandir los datos de la base de datos para cada accountname
    for accountname in accountnames:
        for index, row in bd_data.iterrows():
            # Dividir valores separados por comas en Service Plan Type (O) y Service Types
            service_plan_type_list = str(row['Service Plan Type (O)']).split(',') if pd.notna(row['Service Plan Type (O)']) else ['']
            service_type_list = str(row['Service Types']).split(',') if pd.notna(row['Service Types']) else ['']

            # Crear combinaciones para cada cuenta
            for service_plan_type in service_plan_type_list:
                if service_type_list == ['']:  # Si Service Types está vacío, mantenerlo vacío
                    accounts.append(accountname)
                    names.append('')
                    descriptions.append(row['Description (M)'])
                    external_ids.append(row['External ID (M)'])
                    service_plan_types.append(service_plan_type.strip())
                    service_types.append('')
                else:
                    for service_type in service_type_list:
                        accounts.append(accountname)
                        names.append('')
                        descriptions.append(row['Description (M)'])
                        external_ids.append(row['External ID (M)'])
                        service_plan_types.append(service_plan_type.strip())
                        service_types.append(service_type.strip())

    # Crear el DataFrame de salida
    data = {
        'Account (M)': accounts,
        'Name (M)': names,
        'Description (M)': descriptions,
        'External ID (M)': external_ids,
        'Service Plan Type (O)': service_plan_types,
        'Service Types': service_types
    }

    df_output = pd.DataFrame(data)
    df_output.to_excel(output_path, index=False)

# Main
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python script_service_plan.py <ruta_archivo_seleccionado>")
        sys.exit(1)

    ruta_archivo_seleccionado = sys.argv[1]
    
    if not os.path.exists(ruta_archivo_seleccionado):
        print(f"El archivo especificado no existe: {ruta_archivo_seleccionado}")
        sys.exit(1)

    directorio_base = "C:\\Migracion"
    directorio_template = os.path.join(directorio_base, "Service Plan", "Template de migracion")
    directorio_bd = os.path.join(directorio_base, "Service Plan")

    if not os.path.exists(directorio_template):
        os.makedirs(directorio_template)

    accountnames = obtener_accountnames(ruta_archivo_seleccionado)

    archivo_bd = os.path.join(directorio_bd, "bd.xlsx")
    if not os.path.exists(archivo_bd):
        print(f"El archivo {archivo_bd} no existe.")
        sys.exit(1)

    bd_data = obtener_datos_bd(archivo_bd)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_output = f"Service_Plan_{timestamp}.xlsx"
    ruta_output = os.path.join(directorio_template, nombre_output)

    generar_archivo_migracion(accountnames, bd_data, ruta_output)
    print(f"Archivo generado exitosamente en: {ruta_output}")

import requests
import pandas as pd
import os
import sys
from datetime import datetime

archivo_generado_account_plan = ""

def obtener_account_ids(excel_file):
    """
    Lee los account_ids desde la columna 1 (A) del archivo Excel proporcionado.
    """
    try:
        data = pd.read_excel(excel_file, usecols=[0], header=0)  # Leer solo la columna A
        account_ids = data.iloc[1:, 0].dropna().astype(float).tolist()  # Ignorar la primera fila y filas vacías
        return [int(account_id) for account_id in account_ids]
    except Exception as e:
        print(f"Error al leer los IDs desde el archivo '{excel_file}': {e}")
        return []

def consultar_api_y_generar_excel(excel_file):
    """
    Consulta la API usando los account_ids y genera un archivo Excel con el resultado.
    """
    global archivo_generado_account_plan

    # Configuración de la API
    base_url = "https://restapi1.jasper.com/rws/api/v1/accounts/"
    token = "Basic YWJ1c3RhbWFudGU6ODhiZGNjZTQtZGYwOS00MTIyLThiNjgtMTcxZDM1N2EzZTdl"
    headers_api = {
        "Authorization": token,
        "Accept": "Application/json"
    }

    # Valores fijos
    fixed_values = {
        "SIM State for total Lines (M)": "Activated",
        "Contract Start Date (M)": "2024-12-01 00:00:00",
        "Contract Duration (M)": 9999999999,
        "Recurring Charge": "Yes",
        "Proration Required": "No",
        "Provide Discount?": "No"
    }

    # Cabeceras del archivo Excel
    columns = [
        "Account (M)", "Name (M)", "Currency (M)", "Description (M)", "Legacy Plan Id (O)",
        "SIM State for total Lines (M)", "External ID (O)", "Contract Start Date (M)",
        "Contract Duration (M)", "One Time Charge", "Recurring Charge", "Proration Required",
        "Provide Discount?", "Discount Type"
    ]

    # Directorio de salida
    #output_directory = r"C:\\Migracion\\Account Plan\\Tempate de migracion"
    #os.makedirs(output_directory, exist_ok=True)

    # Extraer solo el nombre del archivo sin la ruta completa
    excel_file_name = os.path.splitext(os.path.basename(excel_file))[0]

    # Directorio de salida
    output_directory = rf"C:\\Migracion\\Account Plan\\Tempate de migracion\\{excel_file_name}"
    os.makedirs(output_directory, exist_ok=True)

    # Obtener account_ids
    account_ids = obtener_account_ids(excel_file)
    if not account_ids:
        print("No se encontraron account_ids válidos en el archivo seleccionado.")
        return

    # Lista para almacenar los resultados
    result_data = []

    print(f"Iniciando la consulta de {len(account_ids)} account_ids...")

    # Iterar sobre los account_ids y consultar la API
    for idx, account_id in enumerate(account_ids, start=1):
        print(f"Procesando account_id {account_id} ({idx}/{len(account_ids)})...")
        try:
            response = requests.get(f"{base_url}{account_id}", headers=headers_api)

            if response.status_code == 200:
                # Procesar la respuesta JSON
                data = response.json()
                account_name = data.get("accountName", "")
                prefixed_name = f"Plan_account_{account_name}"
                account_data = {
                    "Account (M)": account_name,
                    "Name (M)": prefixed_name,
                    "Currency (M)": data.get("currency", ""),
                    "Description (M)": prefixed_name,  # Ahora igual al campo "Name (M)"
                    "Legacy Plan Id (O)": data.get("accountId", ""),
                    "SIM State for total Lines (M)": fixed_values["SIM State for total Lines (M)"],
                    "External ID (O)": data.get("externalId", ""),
                    "Contract Start Date (M)": fixed_values["Contract Start Date (M)"],
                    "Contract Duration (M)": fixed_values["Contract Duration (M)"],
                    "One Time Charge": "",
                    "Recurring Charge": fixed_values["Recurring Charge"],
                    "Proration Required": fixed_values["Proration Required"],
                    "Provide Discount?": fixed_values["Provide Discount?"],
                    "Discount Type": ""
                }
                result_data.append(account_data)
            else:
                print(f"Error al consultar el account_id {account_id}: {response.status_code}")
        except requests.exceptions.RequestException as e:
            print(f"Error de conexión al consultar el account_id {account_id}: {e}")

    # Guardar los datos en un archivo Excel
    if result_data:
        df = pd.DataFrame(result_data, columns=columns)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        archivo_generado_account_plan = os.path.join(output_directory, f"Account_Plan_{timestamp}.xlsx")
        df.to_excel(archivo_generado_account_plan, index=False)
        print(archivo_generado_account_plan)  # Solo el nombre sin texto adicional
    else:
        print("No se encontraron datos para guardar.")

  

# Ejecución del script
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python generar_account_plan.py <ruta_del_archivo.xlsx>")
    else:
        consultar_api_y_generar_excel(sys.argv[1])

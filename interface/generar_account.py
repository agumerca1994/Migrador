import sys
import os
import requests
import pandas as pd
from datetime import datetime

# Función para consultar la API y procesar el archivo seleccionado
def consultar_api_accounts(input_file):
    """
    Consulta la API de Jasper con los account_ids de un archivo Excel,
    y guarda el resultado completo de la API con los campos específicos en el archivo Excel.
    """
    # Directorio de salida
    #output_directory = r"C:\Migracion\Accounts\Template de migracion"

    # Extraer solo el nombre del archivo sin la ruta completa
    excel_file_name = os.path.splitext(os.path.basename(input_file))[0]

    # Directorio de salida
    output_directory = f"../Accounts/Template de migracion/{excel_file_name}/"
    os.makedirs(output_directory, exist_ok=True)

    # Token de autenticación
    token = "Basic YWJ1c3RhbWFudGU6ODhiZGNjZTQtZGYwOS00MTIyLThiNjgtMTcxZDM1N2EzZTdl"

    # Cabecera que se requiere en el archivo de salida
    headers = [
        "Parent Account (M)", "Account Name (M)", "Account Vertical (M)", "Country (M)", "Currency (M)",
        "Account State (M)", "Billing Flag (M)", "Bill Date (M)", "Frequency (M)", "Tax ID (O)", "Area Code (O)",
        "Enterprise ID (M)", "Region ID (M)", "Legacy BAN (M)", "Provider Id (O)", "SMSR Id (O)", "ICCID Manager (O)",
        "Profile Type (O)", "SIM Manufacturer (M)", "Do you want to suspend all the SIMs on Account Suspension? (M)",
        "In which account statement do you want to Terminate all your SIMs? (M)", "Billing Contact Name (M)",
        "Billing Address(M)", "Billing Email Address (M)", "Billing Secondary Email Address (M)", "Billing Primary Phone (M)",
        "Billing Secondary Phone (O)", "Contact Name (M)", "Last Name (M)", "Primary Phone (M)", "Secondary Phone (O)",
        "Email Address (M)", "Address (M)"
    ]

    # Valores fijos para algunos campos
    fixed_values = {
        "Parent Account (M)": "Claro Argentina",
        "Country (M)": "Argentina",
        "Account State (M)": "Activated",
        "Billing Flag (M)": "Yes",
        "Bill Date (M)": "8",
        "Frequency (M)": "Monthly",
        "Region ID (M)": "1",
        "SIM Manufacturer (M)": "GEMALTO",
        "Do you want to suspend all the SIMs on Account Suspension? (M)": "Allowed",
        "In which account statement do you want to Terminate all your SIMs? (M)": "Terminate"
    }

    try:
        # Leer los account_id desde la columna A (columna 1 = A) desde la fila 2
        data = pd.read_excel(input_file, usecols=[0], header=0)
        account_ids = data.iloc[1:, 0].dropna().astype(float).tolist()
        account_ids = [int(account_id) for account_id in account_ids]

        if not os.path.exists(output_directory):
            os.makedirs(output_directory)

        # URL base de la API y los headers
        base_url = "https://restapi1.jasper.com/rws/api/v1/accounts/"

        headers_api = {
            "Authorization": f"{token}",
            "Content-Type": "application/json"
        }

        result_data = []
        total_account_ids = len(account_ids)
        print(f"Iniciando la consulta de {total_account_ids} account_ids...")

        for idx, account_id in enumerate(account_ids, start=1):
            print(f"Procesando account_id {account_id} ({idx}/{total_account_ids})...")
            response = requests.get(f"{base_url}{account_id}", headers=headers_api)

            if response.status_code == 200:
                account_data = {
                    "Parent Account (M)": fixed_values["Parent Account (M)"],
                    "Account Name (M)": response.json().get("accountName", ""),
                    "Account Vertical (M)": response.json().get("industryVertical", ""),
                    "Country (M)": fixed_values["Country (M)"],
                    "Currency (M)": response.json().get("currency", ""),
                    "Account State (M)": response.json().get("status", ""),
                    "Billing Flag (M)": fixed_values["Billing Flag (M)"],
                    "Bill Date (M)": fixed_values["Bill Date (M)"],
                    "Frequency (M)": fixed_values["Frequency (M)"],
                    "Tax ID (O)": response.json().get("taxId", ""),
                    "Area Code (O)": response.json().get("areaCode", ""),
                    "Enterprise ID (M)": response.json().get("operatorAccountId", ""),
                    "Region ID (M)": fixed_values["Region ID (M)"],
                    "Legacy BAN (M)": response.json().get("operatorAccountId", ""),
                    "Provider Id (O)": response.json().get("providerId", ""),
                    "SMSR Id (O)": response.json().get("smsrId", ""),
                    "ICCID Manager (O)": response.json().get("iccidManager", ""),
                    "Profile Type (O)": response.json().get("profileType", ""),
                    "SIM Manufacturer (M)": fixed_values["SIM Manufacturer (M)"],
                    "Do you want to suspend all the SIMs on Account Suspension? (M)": fixed_values["Do you want to suspend all the SIMs on Account Suspension? (M)"],
                    "In which account statement do you want to Terminate all your SIMs? (M)": fixed_values["In which account statement do you want to Terminate all your SIMs? (M)"],
                    "Billing Contact Name (M)": response.json().get("billingContact", {}).get("firstName", ""),
                    "Billing Address(M)": response.json().get("billingAddr", {}).get("addr1", ""),
                    "Billing Email Address (M)": response.json().get("billingContact", {}).get("email", ""),
                    "Billing Secondary Email Address (M)": response.json().get("billingSecondaryEmailAddress", ""),
                    "Billing Primary Phone (M)": response.json().get("billingContact", {}).get("phone", ""),
                    "Billing Secondary Phone (O)": response.json().get("billingSecondaryPhone", ""),
                    "Contact Name (M)": response.json().get("primaryContact", {}).get("firstName", ""),
                    "Last Name (M)": response.json().get("primaryContact", {}).get("lastName", ""),
                    "Primary Phone (M)": response.json().get("primaryContact", {}).get("phone", ""),
                    "Secondary Phone (O)": response.json().get("secondaryPhone", ""),
                    "Email Address (M)": response.json().get("primaryContact", {}).get("email", ""),
                    "Address (M)": response.json().get("ppuAddress", {}).get("addr1", ""),
                }

                result_data.append(account_data)
            else:
                print(f"Error al consultar el account_id {account_id}: {response.status_code}")

        # Guardar los resultados en un archivo Excel
        if result_data:
            result_df = pd.DataFrame(result_data, columns=headers)  # Aseguramos que las columnas estén en el orden adecuado
            # Definir el timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            # Nombre del archivo con timestamp
            timestamp_filename = f"Account_{timestamp}.xlsx"
            output_path = os.path.join(output_directory, timestamp_filename)
            result_df.to_excel(output_path, index=False)
            print(f"Los resultados se han guardado en '{output_path}'.")
        else:
            print("No se encontraron datos para guardar.")

    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")
        return ""

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python script.py <archivo>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = consultar_api_accounts(input_file)

    # Imprimir el nombre del archivo generado como salida
    if output_file:
        print(output_file)

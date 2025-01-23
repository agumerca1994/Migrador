import json
import pandas as pd
import requests
import openpyxl
from openpyxl.styles import Alignment, Font
from datetime import datetime
import os
import sys

# Función para obtener datos de un usuario desde la API
def fetch_user_data(user_id, headers):
    url = f"https://restapi1.jasper.com/rws/api/v1/users/{user_id}"
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        data = response.json()

        # Filtrar los campos no deseados
        fields_to_remove = [
            "userId", "accountId", "operatorName", "operatorId", "status",
            "userLocked", "accessType", "language", "customerName", "customerId",
            "customerGroup", "accountGroup", "lastLogin", "lastPasswordResetDate",
            "passwordExpirationInDays", "liveUpdateEnabled", "dateAdded", "dateModified"
        ]
        return {key: value for key, value in data.items() if key not in fields_to_remove}
    else:
        print(f"Error al obtener datos para el usuario {user_id}: {response.status_code} - {response.text}")
        return None

# Función para exportar datos a un archivo Excel
def export_to_excel(users_data, file_name, user_selected_file):
    if not file_name.endswith(".xlsx"):
        file_name += ".xlsx"

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "API Data"

    # Ordenar y renombrar los campos restantes
    ordered_fields = [
        ("User Type (M)", "User Type (M)"),
        ("User Category (M)", "User Category (M)"),
        ("accountName", "Account (M)"),
        ("username", "User Name (M)"),
        ("firstName", "First Name (M)"),
        ("lastName", "Last Name (M)"),
        ("phone", "Primary Phone (M)"),
        ("Secondary Phone", "Secondary Phone"),
        ("email", "Email Address (M)"),
        ("email", "Confirm Email Address (M)"),
        ("roleName", "Role (M)"),
        ("User Lock (M)", "User Lock (M)"),
        ("Country (M)", "Country (M)"),
        ("timeZone", "Time Zone (M)"),
    ]

    fixed_values = {
        "User Type (M)": "Platform User",
        "User Category (M)": "Normal User",
        "Role (M)": "ACCOUNTADMIN",
        "User Lock (M)": "No",
        "Country (M)": "Argentina",
        "Time Zone (M)": "Buenos Aires Georgetown"
    }

    # Escribir encabezados
    for col_num, (_, header) in enumerate(ordered_fields, start=1):
        cell = sheet.cell(row=1, column=col_num, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    # Escribir datos de los usuarios
    for row_num, user_data in enumerate(users_data, start=2):
        for col_num, (key, header) in enumerate(ordered_fields, start=1):
            value = fixed_values.get(header, user_data.get(key, ""))
            sheet.cell(row=row_num, column=col_num, value=value)

    # Ajustar el ancho de las columnas
    for column in sheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = max_length + 2
        sheet.column_dimensions[column_letter].width = adjusted_width

    # Definir el directorio de destino basado en el archivo seleccionado por el usuario
    directory = os.path.join(r"C:\Migracion\Users\Template de migracion", os.path.splitext(user_selected_file)[0])
    os.makedirs(directory, exist_ok=True)

    # Guardar el archivo Excel en el directorio con timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"Users_{timestamp}.xlsx"
    file_path = os.path.join(directory, file_name)

    workbook.save(file_path)
    print(f"Datos exportados exitosamente a {file_path}")

# Inicio del script
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Uso: python script.py <nombre_del_archivo_seleccionado_por_el_usuario> <directorio_de_salida>")
        sys.exit(1)

    user_selected_file = sys.argv[1]
    output_folder = sys.argv[2]
    
    print(f"Archivo seleccionado por el usuario: {user_selected_file}")
    print(f"Directorio de salida: {output_folder}")

    # Leer el archivo Excel "resultados_user_ids.xlsx"
    input_file = r"C:\Migracion\Users\users_id.xlsx"
    try:
        data = pd.read_excel(input_file, usecols=[1], header=0)  # Leer solo la columna B
        user_ids = data.iloc[1:, 0].dropna().astype(str).str.strip().tolist()  # Desde fila 2 hasta la última con datos

        headers = {
            "Accept": "application/json",
            "Authorization": "Basic YWJ1c3RhbWFudGU6ODhiZGNjZTQtZGYwOS00MTIyLThiNjgtMTcxZDM1N2EzZTdl"
        }

        # Obtener datos de los usuarios
        users_data = [fetch_user_data(user_id, headers) for user_id in user_ids]
        users_data = [data for data in users_data if data is not None]  # Filtrar errores

        if users_data:
            export_to_excel(users_data, user_selected_file, output_folder)  # El nombre del archivo se generará automáticamente
        else:
            print("No se obtuvieron datos válidos para exportar.")
    except FileNotFoundError:
        print(f"El archivo '{input_file}' no se encontró. Verifique el nombre y la ubicación del archivo.")
    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")
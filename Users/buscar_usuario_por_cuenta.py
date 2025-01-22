import json
import pandas as pd
import os

def buscar_users_ids():
    """
    Busca archivos Excel en el directorio 'C:\\Migracion', permite al usuario seleccionar uno,
    y busca coincidencias en el archivo 'users.txt'.
    Guarda el resultado en un archivo Excel con las columnas 'AccountName' y 'user_id'.
    """
    # Directorio de búsqueda
    directory = r"C:\Migracion"

    try:
        # Listar archivos Excel en el directorio
        files = [file for file in os.listdir(directory) if file.endswith(".xlsx")]
        
        if not files:
            print(f"No se encontraron archivos Excel en el directorio '{directory}'.")
            return

        # Mostrar los archivos disponibles
        print("Archivos disponibles:")
        for idx, file in enumerate(files, start=1):
            print(f"{idx}. {file}")

        # Mensaje adicional al usuario
        print("\nIndique el archivo base según la ventana de migración que corresponda:")

        # Solicitar al usuario seleccionar un archivo
        while True:
            try:
                choice = int(input(f"Seleccione el número del archivo (1-{len(files)}): ").strip())
                if 1 <= choice <= len(files):
                    selected_file = files[choice - 1]
                    break
                else:
                    print("Por favor, ingrese un número válido.")
            except ValueError:
                print("Entrada no válida. Intente de nuevo.")

        # Ruta completa del archivo seleccionado
        excel_file = os.path.join(directory, selected_file)

        # Leer los accountNames desde la columna B (columna 1 = A, columna 2 = B) desde la fila 2
        data = pd.read_excel(excel_file, usecols=[1], header=0)  # Leer solo la columna B
        account_names = data.iloc[1:, 0].dropna().astype(str).str.strip().tolist()  # Desde fila 2 hasta la última con datos

        # Leer el archivo 'users.txt' que contiene los usuarios
        with open('users.txt', 'r') as archivo:
            users_data = json.load(archivo)

        # Buscar los userIds correspondientes a los accountNames ingresados
        matched_data = []
        for user in users_data:
            if user["accountName"] in account_names:
                matched_data.append({"AccountName": user["accountName"], "user_id": user["userId"]})

        # Duplicar accountName si tiene más de un user_id asociado
        final_data = []
        for account_name in account_names:
            associated_users = [entry for entry in matched_data if entry["AccountName"] == account_name]
            final_data.extend(associated_users)

        # Si encontramos coincidencias, guardamos los resultados en un archivo Excel
        if final_data:
            result_df = pd.DataFrame(final_data)
            result_df.to_excel("users_id.xlsx", index=False)
            print(f"Los resultados se han guardado en 'users_id.xlsx'.")
        else:
            print("No se encontraron coincidencias para los accountNames proporcionados.")

    except FileNotFoundError:
        print(f"El directorio '{directory}' o algún archivo no se encontró.")
    except IOError as e:
        print(f"Error al leer un archivo: {e}")
    except json.JSONDecodeError as e:
        print(f"Error al procesar el archivo JSON: {e}")
    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")

# Ejemplo de uso
if __name__ == "__main__":
    buscar_users_ids()

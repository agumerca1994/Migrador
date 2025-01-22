import requests
import json

def guardar_usuarios_en_archivo():
    """
    Consulta la API iterando entre páginas hasta la última página con datos y guarda el resultado en un archivo.
    Solo se guardan los campos accountName y userId.
    Si se recibe un error 400 por una página vacía, termina la iteración y guarda los resultados.
    """
    headers = {
        "Accept": "application/json",
        "Authorization": "Basic YWJ1c3RhbWFudGU6ODhiZGNjZTQtZGYwOS00MTIyLThiNjgtMTcxZDM1N2EzZTdl"
    }

    todos_los_usuarios = []
    page_number = 1

    while True:
        # Construimos la URL completa con los parámetros
        url = f"https://restapi1.jasper.com/rws/api/v1/users?pageSize=50&pageNumber={page_number}"

        try:
            response = requests.get(url, headers=headers)
            # Verificamos si la respuesta es un error 400 (página vacía)
            if response.status_code == 400:
                print(f"Error 400 recibido en la página {page_number}. Finalizando la iteración.")
                break  # Termina la iteración si hay un error 400
            response.raise_for_status()  # Verifica que no haya errores HTTP (otros códigos de error)
            data = response.json()

            # Verificamos si hay usuarios en la respuesta
            users = data.get('users', [])
            if users:
                # Filtramos los campos accountName y userId de cada usuario
                usuarios_filtrados = [
                    {"accountName": user["accountName"], "userId": user["userId"]}
                    for user in users
                ]
                todos_los_usuarios.extend(usuarios_filtrados)
                page_number += 1  # Incrementamos el número de página
            else:
                break  # Salimos del bucle si no hay más datos
        except requests.exceptions.RequestException as e:
            print(f"Error al procesar la página {page_number}: {e}")
            break

    # Guardamos los datos en 'users.txt'
    try:
        with open('users.txt', 'w') as archivo:
            json.dump(todos_los_usuarios, archivo, indent=4)
        print(f"Datos guardados correctamente en 'users.txt'.")
    except IOError as e:
        print(f"Error al guardar el archivo: {e}")

# Ejemplo de uso
if __name__ == "__main__":
    guardar_usuarios_en_archivo()
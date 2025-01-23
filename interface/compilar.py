import os
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import sys

# Función para mostrar la ventana emergente y obtener la selección del usuario
def seleccionar_carpetas(subfolders, output_folder, selected_file_name):
    def on_submit():
        selected_folders = [subfolders[i] for i in range(len(subfolders)) if var_list[i].get()]
        root.destroy()
        output_file = process_folders(selected_folders, output_folder, selected_file_name)
        if output_file:
            print(f"Archivo generado: {output_file}")
            global output_folder_path, output_path
            output_folder_path = output_file
            output_path = output_folder_path
            #label.config(text=f"Archivo generado: {output_path}")
        else:
            print("No se generó ningún archivo.")
            #label.config(text="No se generó ningún archivo.")

    def toggle_select_all():
        new_state = not all(var.get() for var in var_list)
        for var in var_list:
            var.set(new_state)
        update_toggle_btn_text()

    def update_toggle_btn_text():
        if all(var.get() for var in var_list):
            toggle_btn.config(text="Deseleccionar todo", fg="red")
        else:
            toggle_btn.config(text="Seleccionar todo", fg="black")

    root = tk.Tk()
    root.title("Seleccionar Carpetas")

    # Centrar la ventana en la pantalla y ajustar el tamaño
    window_width = 600
    window_height = 400
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_top = int(screen_height / 2 - window_height / 2)
    position_right = int(screen_width / 2 - window_width / 2)
    root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')

    # Configurar estilos
    style = ttk.Style()
    style.configure("TButton", font=("Helvetica", 12), padding=10)
    style.configure("TCheckbutton", font=("Helvetica", 12))

    var_list = []

    # Botón para seleccionar/deseleccionar todo
    toggle_btn = tk.Button(root, text="Seleccionar todo", command=toggle_select_all, font=("Helvetica", 12))
    toggle_btn.pack(anchor='w', pady=(20, 10), padx=20)

    # Frame para los checkboxes con margen superior
    frame = ttk.Frame(root)
    frame.pack(anchor='w', pady=(10, 0), padx=20)

    for folder in subfolders:
        var = tk.BooleanVar()
        chk = ttk.Checkbutton(frame, text=folder, variable=var, style="TCheckbutton", command=update_toggle_btn_text)
        chk.pack(anchor='w')
        var_list.append(var)

# Configurar estilos
    style = ttk.Style()
    style.configure("TButton", font=("Helvetica", 12), padding=(10, 10), anchor="center")

    # Crear el botón de "Continuar" con el estilo actualizado
    submit_btn = ttk.Button(root, text="Continuar", command=on_submit, style="TButton")
    submit_btn.pack(pady=(10, 20), ipadx=10, ipady=10)

    root.mainloop()

# Función para procesar las carpetas seleccionadas
def process_folders(selected_folders, output_folder, selected_file_name):
    # Ruta base
    base_path = (f"../")

    # Crear la carpeta base si no existe
    if not os.path.exists(base_path):
        os.makedirs(base_path)

    # Crear la carpeta de salida si no existe
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Nombre del archivo final  
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(output_folder, f"template_de_migracion_{timestamp}.xlsx")

    sheet_added = False

    # Crear un libro de Excel
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for subfolder in selected_folders:
            folder_path = os.path.join(base_path, subfolder, "Template de migracion", selected_file_name)
            
            # Verificar si existe el directorio "Template de migracion"
            if os.path.exists(folder_path):
                # Obtener la lista de archivos .xlsx
                xlsx_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]
                
                if xlsx_files:
                    # Ajustar nombres de archivos según subcarpeta
                    if subfolder == "Account Plan":
                        xlsx_files = [f for f in xlsx_files if f.startswith("Account_Plan_")]
                    
                    # Obtener el archivo más reciente
                    if xlsx_files:  # Asegurarse de que la lista no esté vacía después del filtro
                        xlsx_files.sort(key=lambda f: os.path.getmtime(os.path.join(folder_path, f)), reverse=True)
                        latest_file = xlsx_files[0]
                        print(f"Archivo más reciente en {subfolder}: {latest_file}")

                        # Leer el archivo y agregarlo al libro final
                        try:
                            df = pd.read_excel(os.path.join(folder_path, latest_file), engine='openpyxl')
                            sheet_name = subfolder[:31]  # Limitar a 31 caracteres para nombres de hoja en Excel
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                            sheet_added = True
                        except Exception as e:
                            print(f"Advertencia: Error procesando {latest_file} en {subfolder}: {e}")
                else:
                    print(f"Advertencia: No se encontraron archivos .xlsx en {folder_path}")
            else:
                print(f"Advertencia: Carpeta no encontrada: {subfolder}")

    if not sheet_added:
        print("Advertencia: No se añadieron hojas al archivo Excel porque no se encontraron datos válidos.")
        os.remove(output_file)  # Eliminar el archivo vacío
        return None

    return output_file

# Función para ejecutar el merge y guardar el archivo en el directorio especificado
def ejecutar_merge(selected_file):
    selected_file_name = os.path.splitext(os.path.basename(selected_file))[0]
    output_folder_path = os.path.join(f"../Template de migracion/", selected_file_name)

    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)

    subfolders = [f.name for f in os.scandir(f"../") if f.is_dir() and f.name not in ['.git', 'interface', 'Template de migracion', 'ejemplo']]
    
    seleccionar_carpetas(subfolders, output_folder_path, selected_file_name)

# Punto de entrada principal
if __name__ == "__main__":
    global output_folder_path, output_path
    output_folder_path = None
    
    if len(sys.argv) > 1:  # Verifica si se pasó un archivo como argumento
        selected_file = sys.argv[1]
        ejecutar_merge(selected_file)
        
        output_path = output_folder_path
        
        print(f"La ruta del archivo generado es: {output_path}")
        
    else:
        print("Error: No se proporcionó un archivo base. Usa: python script.py <ruta_del_archivo>")
        sys.exit(1)
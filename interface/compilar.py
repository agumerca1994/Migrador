import os
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook

# Ruta base
base_path = r"C:\\Migracion"
output_folder = os.path.join(base_path, "Template de migracion")

# Crear la carpeta "Template de migracion" si no existe
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Nombre del archivo final
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
output_file = os.path.join(output_folder, f"template_de_migracion_{timestamp}.xlsx")

# Obtener la lista de carpetas dentro del directorio base
subfolders = [f.name for f in os.scandir(base_path) if f.is_dir()]

# Crear un libro de Excel
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    for subfolder in subfolders:
        folder_path = os.path.join(base_path, subfolder, "Template de migracion")
        
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
                    except Exception as e:
                        print(f"Error procesando {latest_file} en {subfolder}: {e}")
        else:
            print(f"Advertencia: Carpeta no encontrada: {subfolder}")

# Ajustar el ancho de las columnas al contenido en cada hoja
wb = load_workbook(output_file)
for sheet in wb.sheetnames:
    ws = wb[sheet]
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # Obtener la letra de la columna
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width

wb.save(output_file)

print(f"Archivo consolidado generado en: {output_file}")

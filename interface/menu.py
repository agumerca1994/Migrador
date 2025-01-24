import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkinter import PhotoImage
import subprocess
import threading
import os
import queue

# Variables globales
selected_file = None
output_folder = None
message_queue = queue.Queue()

# Función para seleccionar un archivo
# Permite al usuario seleccionar un archivo .xlsx desde el directorio especificado
def seleccionar_archivo():
    global selected_file, output_folder
    directorio = (f"../")
    archivo = filedialog.askopenfilename(
        initialdir=directorio,
        title="Seleccionar archivo según ventana de migración que corresponda",
        filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
    )
    if archivo:
        selected_file = archivo
        label_file.config(text=f"Archivo cargado: {os.path.basename(archivo)}")
        for button in buttons:
            button.config(state=tk.NORMAL)
        button_select_file_merge.config(state=tk.NORMAL)  # Habilitar el botón "Merge"
        
        # Crear carpeta de salida si no existe
        output_folder_name = os.path.splitext(os.path.basename(archivo))[0]
        output_folder_path = os.path.join(f"/Template de migracion/", output_folder_name)
        
        if not os.path.exists(output_folder_path):
            os.makedirs(output_folder_path)
        
        output_folder = output_folder_path
        
        print(f"Carpeta de salida: {output_folder}")
    else:
        messagebox.showwarning("Sin selección", "No se seleccionó ningún archivo.")

# Función para ejecutar un script en un hilo separado
# Actualiza la barra de progreso mientras se ejecuta el script
def ejecutar_generar_script(script_name, label):
    global selected_file, output_folder
    if not selected_file:
        messagebox.showwarning("Archivo no seleccionado", "Por favor, selecciona un archivo primero.")
        return

    progress_bar.pack(fill=tk.X, pady=10)
    hilo = threading.Thread(target=_ejecutar_script, args=(script_name, label, selected_file))
    hilo.start()

# Función auxiliar para ejecutar el script y capturar su salida
def _ejecutar_script(script_name, label, selected_file):
    global output_folder
    try:
        message_queue.put(("start_progress",))
        script_path = os.path.join(f"/interface", script_name)

        if not os.path.exists(script_path):
            message_queue.put(("error", f"El script {script_name} no se encuentra en la ruta especificada."))
            return

        result = subprocess.run(["python", script_path, selected_file, output_folder], capture_output=True, text=True)
        output = result.stdout.strip()
        error = result.stderr.strip()

        message_queue.put(("update_output", output, error, label, os.path.basename(output) if output else "Ninguno"))

    except Exception as e:
        message_queue.put(("error", f"Error al ejecutar el script: {e}"))
    finally:
        message_queue.put(("stop_progress",))

# Función para procesar mensajes entre hilos
# Maneja las actualizaciones de la interfaz desde otros hilos
def procesar_mensajes():
    while not message_queue.empty():
        msg = message_queue.get()

        if msg[0] == "start_progress":
            progress_bar.start()
        elif msg[0] == "stop_progress":
            progress_bar.stop()
        elif msg[0] == "error":
            messagebox.showerror("Error", msg[1])
        elif msg[0] == "update_output":
            output, error, label, file_name = msg[1], msg[2], msg[3], msg[4]
            output_text.delete(1.0, tk.END)
            output_text.insert(tk.END, output)
            if error:
                output_text.insert(tk.END, f"\nErrores:\n{error}")
            label.config(text=f"Archivo generado: {file_name}")

    ventana.after(100, procesar_mensajes)

# Configuración de la ventana principal
ventana = tk.Tk()
ventana.title("Migrador de datos JASPER a CLARO CONNECT")
ventana.state("zoomed")

# Obtener la ruta del directorio actual del script
current_dir = os.path.dirname(__file__)

# Construir la ruta relativa al archivo de imagen
file_path = os.path.join(current_dir, "recursos", "migrador3.png")

# Cambiar el ícono de la ventana (asegúrate de que "icono.ico" esté en la misma carpeta)
#file_path = (f"/interface/recursos/migrador3.png")
icono = tk.PhotoImage(file=file_path)  # Usa .png o .ico
ventana.iconphoto(True, icono)

# Marco principal que contiene dos columnas
frame_main = tk.Frame(ventana, padx=10, pady=10)
frame_main.pack(fill=tk.BOTH, expand=True)

# Sección para seleccionar archivo (en la parte superior)
frame_file = tk.Frame(frame_main)
frame_file.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

# Marco para las columnas (debajo de la sección de selección de archivo)
frame_columns = tk.Frame(frame_main)
frame_columns.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)

# Columna izquierda
frame_left = tk.Frame(frame_columns, width=200)
frame_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

# Columna derecha
frame_right = tk.Frame(frame_columns, width=200)
frame_right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

# Construir la ruta relativa al archivo de imagen
img_path = os.path.join(current_dir, "recursos", "migrador5.png")

# Agregar imagen al lado izquierdo del botón seleccionar archivo
try:
    #img_path = (f"/interface/recursos/migrador5.png")  # Cambiar a PNG
    icon_image = PhotoImage(file=img_path)
    label_icon = tk.Label(frame_file, image=icon_image, anchor="w")  # Anchor para alineación izquierda
    label_icon.image = icon_image  # Mantener referencia para evitar el recolector de basura
    label_icon.pack(side=tk.RIGHT, padx=5)
except Exception as e:
    output_text.insert(tk.END, f"Error al cargar la imagen: {e}\n")  # type: ignore # Mostrar error en el área de salida

# Crear un Label con texto en negrita y alineado a la izquierda
label_select_file = tk.Label(frame_file, text="Seleccione un archivo según la ventana de migración que corresponda", font=("Arial", 12, "bold"), anchor="w")
label_select_file.pack(side=tk.TOP, anchor="w", pady=5)  # Coloca el texto arriba del botón y lo alinea a la izquierda

button_select_file = tk.Button(frame_file, text="Seleccionar archivo", font=("Arial", 12), command=seleccionar_archivo)
button_select_file.pack(side=tk.LEFT, padx=5)

label_file = tk.Label(frame_file, text="Archivo cargado: Ninguno", font=("Arial", 10))
label_file.pack(side=tk.LEFT, padx=5)

# Sección para botones de generación de templates
frame_template = tk.LabelFrame(frame_left, text="Generar plantilla de migración", padx=10, pady=10, font=("Arial", 12))
frame_template.pack(fill=tk.X, pady=10)

# Lista para almacenar botones
buttons = []

# Función para crear botones con etiquetas asociadas
def create_button(text, script_name):
    button = tk.Button(
        frame_template, text=text, font=("Arial", 10),
        command=lambda: ejecutar_generar_script(script_name, labels[text]),
        state=tk.DISABLED
    )
    button.pack(fill=tk.X, pady=2)
    buttons.append(button)
    label = tk.Label(frame_template, text="Archivo generado: Ninguno", font=("Arial", 10))
    label.pack(fill=tk.X, pady=2)
    return label

# Crear botones para cada script y sus etiquetas
labels = {
    "Account Plan": create_button("Account Plan", "generar_account_plan.py"),
    "Account": create_button("Account", "generar_account.py"),
    "APN": create_button("APN", "generar_apn.py"),
    "Roles": create_button("Roles", "generar_roles.py"),
    "SIM Product Type": create_button("SIM Product Type", "generar_sim_pdt.py"),
    "Users": create_button("Users", "generar_users.py"),
    "Zones Management": create_button("Zones Management", "generar_zones.py"),
    "Service Plan": create_button("Service Plan", "generar_service_plan.py"),
    "Price Model": create_button("Price Model", "generar_pm.py"),
    "Device Plan": create_button("Device Plan", "generar_service_plan.py"),
}

# Sección de merge
frame_merge = tk.LabelFrame(frame_right, text="Generar merge de migración", padx=10, pady=10, font=("Arial", 12))
frame_merge.pack(fill=tk.X, pady=10)

def ejecutar_merge():
    global selected_file, output_folder
    if not selected_file:
        messagebox.showwarning("Archivo no seleccionado", "Por favor, selecciona un archivo primero.")
        return

    progress_bar.pack(fill=tk.X, pady=10)
    hilo = threading.Thread(target=_ejecutar_merge_script)
    hilo.start()

def _ejecutar_merge_script():
    global selected_file, output_folder
    try:
        message_queue.put(("start_progress",))
        #script_path = os.path.join(f"/interface/", "compilar.py")
        script_path = os.path.join(current_dir, "compilar.py")

        if not os.path.exists(script_path):
            message_queue.put(("error", "El script compilar.py no se encuentra en la ruta especificada."))
            return

        result = subprocess.run(["python", script_path, selected_file], capture_output=True, text=True)
        output = result.stdout.strip()
        error = result.stderr.strip()

        message_queue.put(("update_output", output, error, label_file_merge, os.path.basename(output) if output else "Ninguno"))

    except Exception as e:
        message_queue.put(("error", f"Error al ejecutar el script: {e}"))
    finally:
        message_queue.put(("stop_progress",))

button_select_file_merge = tk.Button(frame_merge, text="Merge", font=("Arial", 12), command=ejecutar_merge, state=tk.DISABLED)  # Inicialmente deshabilitado
button_select_file_merge.pack(side=tk.LEFT, padx=5)

label_file_merge = tk.Label(frame_merge, text="Archivo generado: Ninguno", font=("Arial", 10))
label_file_merge.pack(side=tk.LEFT, padx=5)

# Sección de barra de progreso
frame_progress = tk.LabelFrame(frame_right, text="Progreso", padx=10, pady=10, font=("Arial", 12))
frame_progress.pack(fill=tk.X, pady=10)
progress_bar = ttk.Progressbar(frame_progress, mode="indeterminate")
progress_bar.pack(fill=tk.X, pady=10)

# Área de texto para mostrar la salida de los scripts
output_text = tk.Text(frame_right, wrap=tk.WORD, height=20, width=80)
output_text.pack(pady=10)

# Botón para salir de la aplicación
button_exit = tk.Button(frame_right, text="Salir", font=("Arial", 12), command=ventana.quit)
button_exit.pack(pady=10)

# Procesar mensajes de hilos secundarios
ventana.after(100, procesar_mensajes)

# Iniciar el bucle principal de la interfaz
ventana.mainloop()

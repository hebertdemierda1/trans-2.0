import os
import shutil
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk  # Importa ttk para el ComboBox

# Función para seleccionar el archivo Excel
def seleccionar_archivo_excel():
    archivo_excel = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
    archivo_excel_entry.delete(0, tk.END)
    archivo_excel_entry.insert(0, archivo_excel)

    # Actualizar las hojas disponibles en el ComboBox
    actualizar_hojas_disponibles(archivo_excel)

# Función para actualizar las hojas disponibles en el ComboBox
def actualizar_hojas_disponibles(archivo_excel):
    try:
        hojas_excel = pd.ExcelFile(archivo_excel).sheet_names
        hoja_excel_combo['values'] = hojas_excel
        hoja_excel_combo.set(hojas_excel[0])  # Establece el valor predeterminado

    except pd.errors.EmptyDataError:
        messagebox.showerror("Error", "El archivo Excel está vacío.")
        hoja_excel_combo['values'] = []

# Función para seleccionar la carpeta de origen
def seleccionar_carpeta_origen():
    carpeta_origen = filedialog.askdirectory()
    carpeta_origen_entry.delete(0, tk.END)
    carpeta_origen_entry.insert(0, carpeta_origen)

# Función para seleccionar la carpeta de destino
def seleccionar_carpeta_destino():
    carpeta_destino = filedialog.askdirectory()
    carpeta_destino_entry.delete(0, tk.END)
    carpeta_destino_entry.insert(0, carpeta_destino)

# Función para mover archivos PDF
def mover_archivos():
    carpeta_origen = carpeta_origen_entry.get()
    carpeta_destino = carpeta_destino_entry.get()
    archivo_excel = archivo_excel_entry.get()

    # Verificar si se proporcionó un archivo Excel
    if not archivo_excel:
        messagebox.showerror("Error", "Por favor, seleccione un archivo Excel.")
        return

    hoja_excel = hoja_excel_combo.get()  # Obtener el valor seleccionado del ComboBox

    try:
        df = pd.read_excel(archivo_excel, sheet_name=hoja_excel)
        
        # Tomar la primera columna del DataFrame como nombres de archivo
        nombres_archivos = df.iloc[:, 0].tolist()

        archivos_faltantes = []
        archivos_repetidos = []

        for nombre_archivo in nombres_archivos:
            # Convertir a minúsculas antes de comparar
            nombre_archivo = str(nombre_archivo).lower()

            origen = os.path.join(carpeta_origen, nombre_archivo)
            destino = os.path.join(carpeta_destino, nombre_archivo)

            # Verificar si el nombre del archivo termina con ".pdf" o ".PDF"
            if os.path.exists(origen) and (nombre_archivo.endswith(".pdf") or nombre_archivo.endswith(".PDF")):
                if not os.path.exists(destino):
                    shutil.move(origen, destino)
                else:
                    archivos_repetidos.append(nombre_archivo)
            else:
                archivos_faltantes.append(nombre_archivo)

        mensaje = "Operación completada."

        if archivos_faltantes:
            mensaje += f"\nArchivos faltantes: {len(archivos_faltantes)}"

        if archivos_repetidos:
            mensaje += f"\nArchivos repetidos: {len(archivos_repetidos)}"

        resultado_label.config(text=mensaje)
        if archivos_faltantes or archivos_repetidos:
            messagebox.showwarning("Advertencia", mensaje)

    except pd.errors.EmptyDataError:
        messagebox.showerror("Error", "El archivo Excel está vacío o no contiene datos en la hoja seleccionada.")
    except Exception as e:
        resultado_label.config(text=f"Error: {str(e)}")
        messagebox.showerror("Error", f"Se produjo un error: {str(e)}")

# Crear la ventana de la GUI
ventana = tk.Tk()
ventana.title("Mover Archivos PDF")

# Configurar el tamaño de fuente
fuente_grande = ('Arial', 12)

# Colores
color_fondo = "#EDEDED"
color_boton = "#008CBA"
color_texto = "#FFFFFF"

# Configurar colores de fondo
ventana.configure(bg=color_fondo)

# Etiqueta y entrada para el archivo Excel
archivo_excel_label = tk.Label(ventana, text="Archivo Excel:", font=fuente_grande, bg=color_fondo)
archivo_excel_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
archivo_excel_entry = tk.Entry(ventana, font=fuente_grande)
archivo_excel_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
seleccionar_archivo_excel_button = tk.Button(ventana, text="Seleccionar Archivo", font=fuente_grande, bg=color_boton, fg=color_texto, command=seleccionar_archivo_excel)
seleccionar_archivo_excel_button.grid(row=0, column=2, padx=10, pady=5)

# Etiqueta y entrada para la carpeta de origen
carpeta_origen_label = tk.Label(ventana, text="Carpeta de origen:", font=fuente_grande, bg=color_fondo)
carpeta_origen_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
carpeta_origen_entry = tk.Entry(ventana, font=fuente_grande)
carpeta_origen_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
seleccionar_origen_button = tk.Button(ventana, text="Seleccionar Carpeta", font=fuente_grande, bg=color_boton, fg=color_texto, command=seleccionar_carpeta_origen)
seleccionar_origen_button.grid(row=1, column=2, padx=10, pady=5)

# Etiqueta y entrada para la carpeta de destino
carpeta_destino_label = tk.Label(ventana, text="Carpeta de destino:", font=fuente_grande, bg=color_fondo)
carpeta_destino_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
carpeta_destino_entry = tk.Entry(ventana, font=fuente_grande)
carpeta_destino_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
seleccionar_destino_button = tk.Button(ventana, text="Seleccionar Carpeta", font=fuente_grande, bg=color_boton, fg=color_texto, command=seleccionar_carpeta_destino)
seleccionar_destino_button.grid(row=2, column=2, padx=10, pady=5)

# ComboBox para seleccionar la hoja de Excel
hoja_excel_label = tk.Label(ventana, text="Hoja de Excel:", font=fuente_grande, bg=color_fondo)
hoja_excel_label.grid(row=3, column=0, padx=10, pady=5, sticky="w")

hoja_excel_combo = ttk.Combobox(ventana, font=fuente_grande)
hoja_excel_combo.grid(row=3, column=1, padx=10, pady=5, sticky="ew")

# Botón para mover archivos
mover_button = tk.Button(ventana, text="Mover Archivos", font=fuente_grande, bg=color_boton, fg=color_texto, command=mover_archivos)
mover_button.grid(row=4, column=1, padx=10, pady=10)

# Etiqueta para mostrar el resultado
resultado_label = tk.Label(ventana, text="", wraplength=500, font=fuente_grande, bg=color_fondo)
resultado_label.grid(row=5, column=0, columnspan=3, padx=10, pady=10, sticky="w")

# Configurar opciones de expansión
ventana.columnconfigure(1, weight=1)
archivo_excel_entry.grid(sticky="ew")
carpeta_origen_entry.grid(sticky="ew")
carpeta_destino_entry.grid(sticky="ew")
hoja_excel_combo.grid(sticky="ew")
resultado_label.grid(sticky="ew")

ventana.mainloop()

import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from PIL import Image, ImageTk
import os
import sys

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # Para empaquetar con PyInstaller o similar
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def process_file(file_path):
    try:
        # Cargar el archivo Excel con la primera fila como encabezados
        df = pd.read_excel(file_path, header=0)

        # Eliminar la segunda y tercera columna si existen
        if df.shape[1] > 2:
            df.drop(df.columns[[1, 2]], axis=1, inplace=True)

        # Elimina la fila de Proveedor menos la primera    

        def eliminar_filas_omit_first(df, palabra):
            count = 0
            indices_a_eliminar = []
            for i, row in df.iterrows():
                if row.astype(str).str.contains(palabra, case=False, na=False).any():
                    count += 1
                    if count > 1:
                        indices_a_eliminar.append(i)
            df.drop(indices_a_eliminar, inplace=True)

        eliminar_filas_omit_first(df, "Proveedor")

        # Elimina Filas debajo de Total

        def eliminar_filas_bajo(df, palabra):
            try:
                # Buscar la fila que contiene la palabra (ignorando mayúsculas/minúsculas)
                indice_fila_palabra = df.index[df.astype(str).str.contains(palabra, case=False, na=False).any()].tolist()[0]
            except IndexError:
                # Si no se encuentra la palabra, se emite una advertencia y se devuelve el DataFrame original
                print(f"Advertencia: La palabra '{palabra}' no se encontró en el DataFrame.")
                return df

            # Eliminar filas desde la fila con la palabra hasta el final del DataFrame
            df.drop(df.index[indice_fila_palabra + 1:], inplace=True)

            eliminar_filas_bajo(df,'Total de Pagos')


        # Eliminar filas que contienen palabras específicas, excluyendo la primera fila (encabezados)
        keywords = ["Fecha de Asiento", "Subtotal", "MANANTIAL DEL SILENCIO", "Página", "Valore Pagados", "Filtros Seleccionados", "Fecha Inicial"]
        for keyword in keywords:
            df = df[~df.apply(lambda row: row.astype(str).str.contains(keyword).any(), axis=1)]

        # Guardar el archivo procesado
        output_path = file_path.replace(".xlsx", "_procesado.xlsx")
        df.to_excel(output_path, index=False)
        messagebox.showinfo("Proceso completado", f"El archivo ha sido procesado y guardado como:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {e}")

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        process_file(file_path)

# Crear la ventana principal
root = tk.Tk()
root.title("CONVERSOR ARCHIVO CM")
root.geometry("500x400")

try:
    # Cargar la imagen de fondo
    background_image_path = resource_path('.\Hotel.jpg')
    if not os.path.exists(background_image_path):
        raise FileNotFoundError(f"No se encontró el archivo de imagen: {background_image_path}")
    background_image = Image.open(background_image_path)
    background_image = background_image.resize((500, 400))
    background_photo = ImageTk.PhotoImage(background_image)
except Exception as e:
    messagebox.showerror("Error al cargar la imagen", f"Ocurrió un error al cargar la imagen de fondo: {e}")
    background_photo = None

# Crear un Canvas para la imagen de fondo
canvas = tk.Canvas(root, width=500, height=400)
canvas.pack(fill="both", expand=True)

if background_photo:
    # Añadir la imagen de fondo al Canvas
    canvas.create_image(0, 0, image=background_photo, anchor='nw')

# Crear el botón de selección de archivo sobre el Canvas
select_button = tk.Button(root, text="SELECCIONAR ARCHIVOS", command=select_file, bg='#000814', fg='white')
select_button_window = canvas.create_window(250, 200, anchor='center', window=select_button)

# Iniciar la aplicación
root.mainloop()
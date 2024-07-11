import pandas as pd
import os
import sys
from tkinter import Tk, Button, filedialog, Label, Canvas
from PIL import Image, ImageTk
from tkinter.font import Font

# Función para obtener la ruta absoluta de la imagen
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Función para seleccionar el archivo y procesarlo para "SUMA DE SALDOS"
def seleccionar_archivo():
    file_path = filedialog.askopenfilename(title="Seleccione el archivo Excel", filetypes=[("Excel files", "*.xls;*.xlsx")])
    if not file_path:
        print("No se seleccionó ningún archivo.")
    else:
        df = pd.read_excel(file_path, engine='xlrd')
        df = df.iloc[4:]
        columnas_a_eliminar = ['C', 'D', 'J']
        columnas_existentes_a_eliminar = [col for col in columnas_a_eliminar if col in df.columns]
        df.drop(columns=columnas_existentes_a_eliminar, inplace=True)
        palabras_clave = ["Contabilidad", "MANANTIAL DEL SILENCIO", "Balancete Provisório", "Moeda", "Totales", "Unnamed", "Saldo Anterior"]
        pattern = '|'.join(palabras_clave)
        df = df[~df.apply(lambda row: row.astype(str).str.contains(pattern, case=False, na=False).any(), axis=1)]

        def eliminar_filas_omit_first(df, palabra):
            count = 0
            indices_a_eliminar = []
            for i, row in df.iterrows():
                if row.astype(str).str.contains(palabra, case=False, na=False).any():
                    count += 1
                    if count > 1:
                        indices_a_eliminar.append(i)
            df.drop(indices_a_eliminar, inplace=True)

        eliminar_filas_omit_first(df, "Cód. Cuenta")
        df.dropna(axis=1, how='all', inplace=True)
        df.dropna(axis=0, how='all', inplace=True)
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        new_file_path = os.path.join(os.path.dirname(file_path), f'{file_name}.xlsx')
        df.to_excel(new_file_path, index=False)
        print(f"Proceso completado. Archivo limpio guardado como '{new_file_path}'.")

        # Mostrar mensaje de éxito
        label_exito.config(text="PROCESADO CON EXITO", fg='green')

# Función para mostrar el botón de seleccionar archivo
def mostrar_boton_seleccionar(tipo_proceso):
    boton_seleccionar.config(command=seleccionar_archivo if tipo_proceso == 'suma_saldos' else seleccionar_archivo_generico)
    boton_seleccionar.place(relx=0.5, rely=0.5, anchor='center')
    boton_atras.place(relx=0.5, rely=0.9, anchor='center')
    boton_suma_saldos.place_forget()
    boton_mayores.place_forget()
    boton_comprobantes.place_forget()

# Función para volver a la pantalla inicial
def ir_atras():
    boton_suma_saldos.place(relx=0.15, rely=0.5, anchor='center')
    boton_mayores.place(relx=0.5, rely=0.5, anchor='center')
    boton_comprobantes.place(relx=0.80, rely=0.5, anchor='center')
    boton_seleccionar.place_forget()
    boton_atras.place_forget()
    label_exito.config(text="")

# Función genérica para seleccionar archivo para otros procesos (MAYORES, COMPROBANTES)
def seleccionar_archivo_generico():
    file_path = filedialog.askopenfilename(title="Seleccione el archivo", filetypes=[("Excel files", "*.xls;*.xlsx")])
    if file_path:
        print(f"Archivo '{file_path}' seleccionado para el proceso genérico.")
        # Aquí se podría agregar el código específico para cada proceso genérico (MAYORES, COMPROBANTES)
        label_exito.config(text=f"Archivo '{os.path.basename(file_path)}' seleccionado para proceso genérico.", fg='blue')

# Crear la ventana principal
root = Tk()
root.title("CONVERSOR ARCHIVO CM")
root.geometry("500x400")

# Cargar la imagen de fondo
background_image_path = resource_path('C:\Python\Codigos\CM-Python/Hotel.jpg')
background_image = Image.open(background_image_path)
background_image = background_image.resize((500, 400))
background_photo = ImageTk.PhotoImage(background_image)

# Crear un Canvas para la imagen de fondo
canvas = Canvas(root, width=500, height=400)
canvas.pack(fill="both", expand=True)

# Añadir la imagen de fondo al Canvas
canvas.create_image(0, 0, image=background_photo, anchor='nw')

# Crear el botón de selección de archivo sobre el Canvas
boton_seleccionar = Button(root, text="SELECCIONAR ARCHIVO", bg='#000814', fg='white')

# Crear los botones "SUMA DE SALDOS", "MAYORES" y "COMPROBANTES"
font_montserrat = Font(family="Montserrat", size=12)

boton_suma_saldos = Button(root, text="SUMA DE SALDOS", command=lambda: mostrar_boton_seleccionar('suma_saldos'), bg='#8ecae6', fg='black', font=font_montserrat)
boton_suma_saldos.place(relx=0.20, rely=0.5, anchor='center')

boton_mayores = Button(root, text="MAYOR", command=lambda: mostrar_boton_seleccionar('mayores'), bg='#219ebc', fg='white', font=font_montserrat)
boton_mayores.place(relx=0.5, rely=0.5, anchor='center')

boton_comprobantes = Button(root, text="COMPROBANTES", command=lambda: mostrar_boton_seleccionar('comprobantes'), bg='#bde0fe', fg='black', font=font_montserrat)
boton_comprobantes.place(relx=0.80, rely=0.5, anchor='center')

# Crear el botón "Atrás"
boton_atras = Button(root, text="Atrás", command=ir_atras, bg='#03045e', fg='white', font=font_montserrat)

# Crear la etiqueta de éxito
label_exito = Label(root, text="", bg='white', fg='black', font=('Helvetica', 12))
label_exito_window = canvas.create_window(200, 250, anchor='center', window=label_exito)

# Iniciar el loop de la aplicación
root.mainloop()

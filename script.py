import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import PhotoImage
from tkinter import Label
from PIL import Image, ImageTk
import openpyxl
from tkinter import simpledialog
import pandas as pd
import subprocess

font_tuple = ("Times New Roman", 16, "bold")

image_dir = 'imagenes'
image_paths = [os.path.join(image_dir, file) for file in os.listdir(image_dir) if file.endswith('.png') or file.endswith('.jpg')]
current_image_index = 0

selected_excel = ""  # Variable para almacenar el nombre del archivo seleccionado

def open_file():
    global selected_excel
    file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
    if file_path:
        selected_excel = file_path
        messagebox.showinfo("Información", "El archivo ha sido seleccionado correctamente")
        # Abre el archivo Excel usando pandas para mostrar el contenido
        df = pd.read_excel(file_path)
        print(df)  # Hacer algo con los datos del archivo

def create_file():
    # Pide al usuario que ingrese el nombre del archivo
    name = simpledialog.askstring("Nombre del archivo", "Ingrese el nombre del archivo:")
    if name:
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=name)
        if file_path:
            # Crear un nuevo archivo Excel usando openpyxl
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Datos"
            workbook.save(filename=file_path)
            messagebox.showinfo("Información", "El archivo se guardó de forma correcta")

def register_excel():
    global selected_excel
    file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
    if file_path:
        selected_excel = file_path
        messagebox.showinfo("Información", "Se ha seleccionado el archivo de Excel correctamente")
        # Ejecutar el otro script con la selección del archivo
        subprocess.Popen(["python", "script2.py", selected_excel])

def open_qr_window():
    global current_image_index

    def show_image():
        global current_image_index
        image_path = image_paths[current_image_index]
        image = Image.open(image_path)
        image = image.resize((200, 200), Image.LANCZOS)
        image = ImageTk.PhotoImage(image)
        image_label.configure(image=image)
        image_label.image = image

    def next_image():
        global current_image_index
        current_image_index = (current_image_index + 1) % len(image_paths)
        show_image()

    def prev_image():
        global current_image_index
        current_image_index = (current_image_index - 1) % len(image_paths)
        show_image()

    qr_window = tk.Toplevel(root)
    qr_window.title("Qrs")
    qr_window.geometry("500x300")

    qr_window.qr_background_image = PhotoImage(file="ventana1.png")
    qr_background_label = Label(qr_window, image=qr_window.qr_background_image)
    qr_background_label.place(x=0, y=0, relwidth=1, relheight=1)

    image_label = Label(qr_window)
    image_label.place(relx=0.5, rely=0.4, anchor='center')

    prev_button_widget = tk.Button(qr_window, text="Anterior", command=prev_image)
    next_button_widget = tk.Button(qr_window, text="Siguiente", command=next_image)
    back_button = tk.Button(qr_window, text="Volver", command=qr_window.destroy)

    prev_button_widget.configure(bg='lightblue')
    next_button_widget.configure(bg='lightgreen')
    back_button.configure(bg='lightyellow')

    prev_button_widget.configure(font=font_tuple)
    next_button_widget.configure(font=font_tuple)
    back_button.configure(font=font_tuple)

    next_button_widget.place(relx=0.9, rely=0.8, anchor="se")
    prev_button_widget.place(relx=0.1, rely=0.8, anchor="sw")
    back_button.place(relx=0.5, rely=0.98, anchor="s")

    show_image()  # Mostrar la primera imagen al abrir la ventana

root = tk.Tk()
root.geometry("500x300")
root.title("Reconocimiento de QR")

background_image = PhotoImage(file="qr.png")
background_label = Label(root, image=background_image)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

open_button = tk.Button(root, text="Abrir archivo", command=open_file)
create_button = tk.Button(root, text="Crear archivo", command=create_file)
qr_button = tk.Button(root, text="QRs", command=open_qr_window)
register_button = tk.Button(root, text="Registrar", command=register_excel)

open_button.configure(bg='lightblue')
create_button.configure(bg='lightgreen')
qr_button.configure(bg='lightyellow')
register_button.configure(bg='#E6E6FA', fg='black')  # Cambiar color de fondo a lila claro y color de texto a negro

open_button.configure(font=font_tuple)
create_button.configure(font=font_tuple)
qr_button.configure(font=font_tuple)
register_button.configure(font=font_tuple)

open_button.pack(anchor='w', padx=(187, 0), pady=40)
create_button.pack(anchor='center', padx=(10, 0))
qr_button.pack(anchor='e', padx=(0, 220), pady=40)
register_button.place(relx=0.98, rely=0.95, anchor='se')

root.mainloop()

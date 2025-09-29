import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
from codigo import funcion_principal

# Configuración inicial
ctk.set_appearance_mode("light")  # "light" o "dark"
ctk.set_default_color_theme("blue")  #  "green", "dark-blue"

# ---------------------- FUNCIONES ----------------------
def seleccionar_pdf():
    archivo = filedialog.askopenfilename(
        title="Seleccionar PDF",
        filetypes=[("Archivos PDF", "*.pdf")]
    )
    if archivo:
        entry_pdf.delete(0, "end")
        entry_pdf.insert(0, archivo)

def seleccionar_directorio():
    carpeta = filedialog.askdirectory(title="Seleccionar carpeta de salida")
    if carpeta:
        entry_dir.delete(0, "end")
        entry_dir.insert(0, carpeta)

def ejecutar():
    pdf_path = entry_pdf.get()
    output_dir = entry_dir.get()
    fecha = entry_fecha.get()
    numero_archivo = entry_num.get()

    if not pdf_path or not output_dir or not fecha or not numero_archivo:
        messagebox.showerror("Error", "Todos los campos son obligatorios.")
        return
    
    try:
        buffer_zip = funcion_principal(pdf_path, output_dir, fecha, numero_archivo)

        zip_filename = os.path.join(output_dir, f"PROVEEDORES_{fecha}_{numero_archivo}.zip")
        with open(zip_filename, "wb") as f:
            f.write(buffer_zip.getvalue())

        messagebox.showinfo("Éxito", f"Archivo ZIP generado:\n{zip_filename}")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error:\n{e}")

# ---------------------- INTERFAZ ----------------------
app = ctk.CTk()
app.title("The Ultimate Separador")
app.geometry("600x400")

frame = ctk.CTkFrame(app, corner_radius=15)
frame.pack(padx=20, pady=20, fill="both", expand=True)

# Campos
label_pdf = ctk.CTkLabel(frame, text="Seleccionar PDF:")
label_pdf.grid(row=0, column=0, sticky="w", padx=10, pady=10)
entry_pdf = ctk.CTkEntry(frame, width=300)
entry_pdf.grid(row=0, column=1, padx=10)
btn_pdf = ctk.CTkButton(frame, text="Buscar", command=seleccionar_pdf)
btn_pdf.grid(row=0, column=2, padx=10)

label_dir = ctk.CTkLabel(frame, text="Carpeta de salida:")
label_dir.grid(row=1, column=0, sticky="w", padx=10, pady=10)
entry_dir = ctk.CTkEntry(frame, width=300)
entry_dir.grid(row=1, column=1, padx=10)
btn_dir = ctk.CTkButton(frame, text="Buscar", command=seleccionar_directorio)
btn_dir.grid(row=1, column=2, padx=10)

label_fecha = ctk.CTkLabel(frame, text="Fecha:")
label_fecha.grid(row=2, column=0, sticky="w", padx=10, pady=10)
entry_fecha = ctk.CTkEntry(frame, width=150)
entry_fecha.grid(row=2, column=1, sticky="w", padx=10)

label_num = ctk.CTkLabel(frame, text="Numero de archivo:")
label_num.grid(row=3, column=0, sticky="w", padx=10, pady=10)
entry_num = ctk.CTkEntry(frame, width=150)
entry_num.grid(row=3, column=1, sticky="w", padx=10)

# Botón ejecutar
btn_ejecutar = ctk.CTkButton(frame, text="Generar ZIP", command=ejecutar, fg_color="green", hover_color="darkgreen")
btn_ejecutar.grid(row=4, column=0, columnspan=3, pady=20)

app.mainloop()

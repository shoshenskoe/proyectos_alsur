
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox
import PyPDF2
from codigo import funcion_principal  #importamos el archivo .py llamado codigo que contiene todas las funciones que definimos

# ---------- Interfaz Gráfica ----------
class PDFSplitterApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("The ultimate separador de pagos")
        self.geometry("600x400")
        ctk.set_appearance_mode("dark")   # "light" o "dark"
        ctk.set_default_color_theme("blue")  # también: "green", "dark-blue"

        # Variables
        self.pdf_path = None
        self.fecha_var = tk.StringVar()
        self.numero_var = tk.StringVar()

        # Widgets
        self.create_widgets()

    def create_widgets(self):
        title = ctk.CTkLabel(self, text="Dividir PDF y guardar como zip", font=("Arial", 22, "bold"))
        title.pack(pady=20)

        # Botón para seleccionar PDF
        self.pdf_button = ctk.CTkButton(self, text="Seleccionar PDF", command=self.seleccionar_pdf)
        self.pdf_button.pack(pady=10)

        # Entrada de fecha
        fecha_label = ctk.CTkLabel(self, text="Fecha (texto):")
        fecha_label.pack()
        fecha_entry = ctk.CTkEntry(self, textvariable=self.fecha_var, width=300)
        fecha_entry.pack(pady=5)

        # Entrada de número de archivo
        numero_label = ctk.CTkLabel(self, text="Numero de archivo (texto):")
        numero_label.pack()
        numero_entry = ctk.CTkEntry(self, textvariable=self.numero_var, width=300)
        numero_entry.pack(pady=5)

        # Botón para generar ZIP
        self.zip_button = ctk.CTkButton(self, text="Generar ZIP", command=self.generar_zip)
        self.zip_button.pack(pady=20)

    def seleccionar_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("Archivo pdf", "*.pdf")])
        if file_path:
            self.pdf_path = file_path
            messagebox.showinfo("Archivo seleccionado", f"Has seleccionado:\n{file_path}")

    def generar_zip(self):
        if not self.pdf_path:
            messagebox.showwarning("Error", "Debes seleccionar un PDF primero.")
            return

        fecha = self.fecha_var.get().strip()
        numero = self.numero_var.get().strip()

        if not fecha or not numero:
            messagebox.showwarning("Error", "Debes ingresar la fecha y el numero de archivo.")
            return

        try:
            zip_buffer = funcion_principal(self.pdf_path, fecha, numero)

            # Guardar ZIP en disco
            save_path = filedialog.asksaveasfilename(
                defaultextension=".zip",
                filetypes=[("ZIP files", "*.zip")],
                initialfile=f"PROVEEDORES_{fecha}_{numero}.zip"
            )

            if save_path:
                with open(save_path, "wb") as f:
                    f.write(zip_buffer.getvalue())
                messagebox.showinfo("Éxito", f"ZIP guardado en:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error:\n{str(e)}")

# Ejecutar la app
if __name__ == "__main__":
    app = PDFSplitterApp()
    app.mainloop()

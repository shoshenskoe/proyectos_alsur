# interfaz.py
import customtkinter as ctk
from tkinter import filedialog, messagebox
from codigo_modificado import main

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Generador de poliza SiVale")
        self.geometry("700x450")
        self.resizable(False, False)

        self.excel_path = ctk.StringVar()
        self.referencia = ctk.StringVar()
        self.dic_nombre_cc = {}
        self.dic_util_cc = {}

        self.crear_widgets()

    def crear_widgets(self):
        ctk.CTkLabel(self, text="Seleccione el archivo Excel:", font=("Arial", 15, "bold")).pack(pady=15)
        ctk.CTkButton(self, text="Examinar", command=self.seleccionar_excel).pack()
        ctk.CTkLabel(self, textvariable=self.excel_path, wraplength=600, font=("Arial", 12)).pack(pady=10)

        ctk.CTkLabel(self, text="Ingrese la referencia:", font=("Arial", 15, "bold")).pack(pady=10)
        ctk.CTkEntry(self, textvariable=self.referencia, width=250).pack(pady=5)

        ctk.CTkButton(self, text="Generar archivo", command=self.ejecutar_main, width=200, height=40).pack(pady=25)

    def seleccionar_excel(self):
        archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if archivo:
            self.excel_path.set(archivo)

    def pedir_valores(self, lista, titulo):
        dic = {}
        ventana = ctk.CTkToplevel(self)
        ventana.title(titulo)
        ventana.geometry("400x400")

        ctk.CTkLabel(ventana, text=f"Ingrese los valores para {titulo}", font=("Arial", 14)).pack(pady=10)
        entries = {}
        for item in lista:
            frame = ctk.CTkFrame(ventana)
            frame.pack(pady=5)
            ctk.CTkLabel(frame, text=item, width=200, anchor="w").pack(side="left")
            e = ctk.CTkEntry(frame, width=120)
            e.pack(side="right")
            entries[item] = e

        def guardar():
            for k, v in entries.items():
                dic[k] = v.get().upper()
            ventana.destroy()

        ctk.CTkButton(ventana, text="Guardar", command=guardar).pack(pady=10)
        ventana.wait_window()
        return dic

    def ejecutar_main(self):
        if not self.excel_path.get() or not self.referencia.get():
            messagebox.showerror("Error", "Debe seleccionar un archivo y una referencia.")
            return

        try:
            buffer, faltantes_cc, faltantes_util = main(self.excel_path.get(), self.referencia.get())

            if faltantes_cc:
                self.dic_nombre_cc = self.pedir_valores(faltantes_cc, "Centros de Costos Faltantes")
                buffer, _, faltantes_util = main(
                    self.excel_path.get(), self.referencia.get(),
                    diccionario_nombre_cc=self.dic_nombre_cc
                )

            if faltantes_util:
                self.dic_util_cc = self.pedir_valores(faltantes_util, "Utilitarios Faltantes")
                buffer, _, _ = main(
                    self.excel_path.get(), self.referencia.get(),
                    diccionario_nombre_cc=self.dic_nombre_cc,
                    diccionario_util_cc=self.dic_util_cc
                )

            ruta_guardar = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
            if ruta_guardar:
                with open(ruta_guardar, "wb") as f:
                    f.write(buffer.getbuffer())
                messagebox.showinfo("Exito", f"Archivo guardado en:\n{ruta_guardar}")

        except Exception as e:
            messagebox.showerror("Error", f"Error:\n{e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()

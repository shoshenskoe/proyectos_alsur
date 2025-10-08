
import customtkinter as ctk
import pandas as pd
from customtkinter import filedialog
from codigo_1 import obtener_dfsucio
from codigo_1 import obtener_df
from codigo_1 import crear_tabla_con_cc_vacia
from codigo_1  import nombres_faltantes
from codigo_1 import hacer_verficiacion_v2
from codigo_1 import crear_segunda_tabla_din
from codigo_1 import enlazar_con_utilitario
from codigo_1 import obtener_faltantes_utilitario
from codigo_1 import main_gui



class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Generador de poliza si vale")
        self.geometry("600x450")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.main_frame = ctk.CTkFrame(self, corner_radius=15)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)

        # --- Widgets ---
        self.title_label = ctk.CTkLabel(self.main_frame, text="Generador de poliza Si Vale", font=ctk.CTkFont(size=24, weight="bold"))
        self.title_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.file_button = ctk.CTkButton(self.main_frame, text="Seleccionar Archivo Excel", command=self.select_file)
        self.file_button.grid(row=1, column=0, padx=40, pady=10, sticky="ew")

        self.file_label = ctk.CTkLabel(self.main_frame, text="Ningún archivo seleccionado", text_color="gray")
        self.file_label.grid(row=2, column=0, padx=40, pady=(0, 20))

        self.ref_entry = ctk.CTkEntry(self.main_frame, placeholder_text="Ingrese la referencia")
        self.ref_entry.grid(row=3, column=0, padx=40, pady=10, sticky="ew")

        self.run_button = ctk.CTkButton(self.main_frame, text="Generar poliza", command=self.run_main_logic)
        self.run_button.grid(row=4, column=0, padx=40, pady=20, sticky="ew")
        
        self.download_button = ctk.CTkButton(self.main_frame, text="Descargar poliza", command=self.download_file, state="disabled")
        self.download_button.grid(row=5, column=0, padx=40, pady=10, sticky="ew")

        self.status_label = ctk.CTkLabel(self.main_frame, text="")
        self.status_label.grid(row=6, column=0, padx=20, pady=(10, 20))

        self.excel_path = ""
        self.excel_buffer = None

    def select_file(self):
        self.excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.excel_path:
            filename = self.excel_path.split('/')[-1]
            self.file_label.configure(text=filename, text_color="white")
            self.status_label.configure(text="")
            self.download_button.configure(state="disabled")

    def run_main_logic(self):
        if not self.excel_path:
            self.status_label.configure(text="Error: Por favor, selecciona un archivo de Excel.", text_color="red")
            return
        
        referencia = self.ref_entry.get()
        if not referencia:
            self.status_label.configure(text="Error: Por favor, ingresa una referencia.", text_color="red")
            return

        self.status_label.configure(text="Procesando...", text_color="white")
        self.update_idletasks() # Actualiza la UI para mostrar el mensaje

        try:
            # --- Primera Verificación: Centros de Costos Faltantes ---
            dfsucio_pre = obtener_dfsucio(self.excel_path)
            df_pre = obtener_df(dfsucio_pre)
            tabla_incompleta_pre = crear_tabla_con_cc_vacia(df_pre)
            nombres_faltantes_list = nombres_faltantes(tabla_incompleta_pre)
            nombres_faltantes_list = [n for n in nombres_faltantes_list if n != "Total"]

            diccionario_nombre_cc = {}
            if nombres_faltantes_list:
                popup = InputPopup(self, "Centros de Costo Faltantes", nombres_faltantes_list, "CC")
                self.wait_window(popup)
                if popup.result is None: # Si el usuario cierra la ventana
                    self.status_label.configure(text="Proceso cancelado por el usuario.", text_color="orange")
                    return
                diccionario_nombre_cc = popup.result

            # --- Segunda Verificación: Utilitarios Faltantes ---
            path_cc_utilitario = "https://docs.google.com/spreadsheets/d/1gnfLiD1arrr5G7seQi85-f3Cd5n7_miS/export?format=csv&gid=1471990202"
            utilitario_pre = pd.read_csv(path_cc_utilitario)
            df_completo_temp, _ = hacer_verficiacion_v2(tabla_incompleta_pre, diccionario_nombre_cc, pd.read_excel("https://docs.google.com/spreadsheets/d/1Iy68cztYlqI6fLjE8l4s2D32BQLzUZG9/export?format=xlsx"))
            df4_pre = crear_segunda_tabla_din(df_completo_temp)
            df_parapoliza_pre = enlazar_con_utilitario(df4_pre, utilitario_pre)
            utilitarios_faltantes_list = obtener_faltantes_utilitario(df_parapoliza_pre)
            utilitarios_faltantes_list = [u for u in utilitarios_faltantes_list if u != "E913"]


            diccionario_util_cc = {}
            if utilitarios_faltantes_list:
                popup = InputPopup(self, "Utilitarios Faltantes", utilitarios_faltantes_list, "Utilitario")
                self.wait_window(popup)
                if popup.result is None: # Si el usuario cierra la ventana
                    self.status_label.configure(text="Proceso cancelado por el usuario.", text_color="orange")
                    return
                diccionario_util_cc = popup.result

            # --- Ejecución Principal ---
            self.excel_buffer = main_gui(
                path_archivo_excel=self.excel_path,
                Referencia=referencia,
                diccionario_nombre_cc=diccionario_nombre_cc,
                diccionario_util_cc=diccionario_util_cc
            )

            self.status_label.configure(text="¡Proceso completado con éxito!", text_color="lightgreen")
            self.download_button.configure(state="normal")

        except Exception as e:
            self.status_label.configure(text=f"Error: {e}", text_color="red")
            self.download_button.configure(state="disabled")

    def download_file(self):
        if self.excel_buffer:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="Poliza_Generada.xlsx"
            )
            if file_path:
                with open(file_path, "wb") as f:
                    f.write(self.excel_buffer.getbuffer())
                self.status_label.configure(text=f"Archivo guardado en: {file_path}", text_color="white")

class InputPopup(ctk.CTkToplevel):
    def __init__(self, parent, title, item_list, input_label):
        super().__init__(parent)
        self.title(title)
        self.geometry("400x500")
        self.transient(parent) # Mantener la ventana emergente sobre la principal
        self.grab_set() # Bloquear la ventana principal
        self.protocol("WM_DELETE_WINDOW", self.on_cancel) # Manejar el cierre de la ventana

        self.result = None
        self.entries = {}

        main_frame = ctk.CTkScrollableFrame(self, label_text=f"Por favor, ingrese los {input_label}s faltantes")
        main_frame.pack(expand=True, fill="both", padx=10, pady=10)

        for i, item in enumerate(item_list):
            label = ctk.CTkLabel(main_frame, text=item)
            label.grid(row=i, column=0, padx=10, pady=5, sticky="w")
            entry = ctk.CTkEntry(main_frame)
            entry.grid(row=i, column=1, padx=10, pady=5, sticky="ew")
            self.entries[item] = entry
        
        main_frame.grid_columnconfigure(1, weight=1)

        submit_button = ctk.CTkButton(self, text="Aceptar", command=self.on_submit)
        submit_button.pack(pady=10)

    def on_submit(self):
        self.result = {item: entry.get() for item, entry in self.entries.items()}
        self.destroy()

    def on_cancel(self):
        self.result = None # Indicar que se canceló
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
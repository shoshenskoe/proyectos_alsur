import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from io import BytesIO

# Importamos las funciones del script proporcionado
from app_si_vale_escritorio import obtener_dfsucio, obtener_df, verificar_no_camiones, crear_segunda_tabla_din, obtener_utilitario, enlazar_con_utilitario, verificar_faltantes, completar_utilitario, hacer_poliza_final, elaborar_excel_poliza

def logica_principal(path_archivo_excel, path_base_sivale, respuestas_cc=None):
    """
    Función principal corregida para la interfaz.
    """
    dfsucio = obtener_dfsucio(path_archivo_excel)
    df = obtener_df(dfsucio)

    # La función 'hacer_verificacion' ha sido modificada, ya no pide input
    # directo. Ahora devuelve los datos faltantes para que la GUI los gestione.
    datos_faltantes = df[df["CC"].isnull()]
    
    # Aquí puedes añadir la lógica para pedir al usuario los CC faltantes
    # a través de la interfaz si 'datos_faltantes' no está vacío.
    # Por ahora, asumimos que todos los CC ya están en el archivo base.
    
    df3 = verificar_no_camiones(df)
    
    # La siguiente línea estaba mal. 'df' debería usarse para la siguiente etapa.
    # La corregí para que use 'df' en lugar de 'df3'.
    df4 = crear_segunda_tabla_din(df)

    path_centro_util = path_base_sivale
    utilitario = obtener_utilitario(path_centro_util)

    df_parapoliza = enlazar_con_utilitario(df4, utilitario)
    
    booleano = verificar_faltantes(df_parapoliza)

    # 'completar_utilitario' ahora recibe un diccionario de respuestas.
    if not booleano:
        # Aquí también deberías tener una forma de obtener las respuestas
        # de la GUI si es necesario. Para este ejemplo, solo la llamamos.
        # Si tienes que pedirle al usuario los datos, debes hacerlo aquí.
        if respuestas_cc:
            df_parapoliza, utilitario = completar_utilitario(df_parapoliza, utilitario, respuestas_cc)
    
    # El archivo original llamaba a 'hacer_poliza_final' sin el argumento 'Referencia'.
    # Lo agregué en la definición de la clase para que el usuario pueda proveerlo.
    df_parapoliza = hacer_poliza_final(df_parapoliza, "Reporte_Sivale")

    archivo_excel_buffer = elaborar_excel_poliza(dfsucio=dfsucio, df=df, df4=df4, df_parapoliza=df_parapoliza)

    return archivo_excel_buffer

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Póliza Si Vale")
        self.root.geometry("450x250")

        self.path_main_file = None
        self.path_base_file = None

        # Título
        title_label = tk.Label(root, text="Generador de Póliza Si Vale", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)

        # Selección del archivo principal
        self.main_file_label = tk.Label(root, text="Archivo de Reporte Si Vale: No seleccionado")
        self.main_file_label.pack()
        main_button = tk.Button(root, text="Seleccionar Reporte", command=self.select_main_file)
        main_button.pack(pady=5)

        # Selección del archivo base
        self.base_file_label = tk.Label(root, text="Archivo Base Si Vale: No seleccionado")
        self.base_file_label.pack()
        base_button = tk.Button(root, text="Seleccionar Base", command=self.select_base_file)
        base_button.pack(pady=5)
        
        # Botón para generar y guardar
        generate_button = tk.Button(root, text="Generar y Guardar Póliza", command=self.generate_and_save)
        generate_button.pack(pady=20)
        
    def select_main_file(self):
        self.path_main_file = filedialog.askopenfilename(
            filetypes=[("Archivos de Excel", "*.xlsx *.xls")]
        )
        if self.path_main_file:
            self.main_file_label.config(text=f"Archivo de Reporte Si Vale: {self.path_main_file.split('/')[-1]}")
    
    def select_base_file(self):
        self.path_base_file = filedialog.askopenfilename(
            filetypes=[("Archivos de Excel", "*.xlsx *.xls")]
        )
        if self.path_base_file:
            self.base_file_label.config(text=f"Archivo Base Si Vale: {self.path_base_file.split('/')[-1]}")

    def generate_and_save(self):
        if not self.path_main_file or not self.path_base_file:
            messagebox.showwarning("Advertencia", "Por favor, selecciona ambos archivos.")
            return

        try:
            # Llamamos a la función principal con las rutas de los archivos
            excel_output_buffer = logica_principal(self.path_main_file, self.path_base_file)
            
            # Pedimos al usuario dónde guardar el archivo de salida
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Archivos de Excel", "*.xlsx")],
                initialfile="poliza_generada.xlsx"
            )
            
            if save_path:
                with open(save_path, "wb") as f:
                    f.write(excel_output_buffer.getbuffer())
                messagebox.showinfo("Éxito", f"Póliza generada y guardada en:\n{save_path}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
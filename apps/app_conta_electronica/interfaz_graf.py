import tkinter as tk
from tkinter import filedialog, messagebox
from script_principal import procesamiento_archivos
import os

def seleccionar_archivo():
    path = filedialog.askopenfilename(
        title="Selecciona el archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if path:
        entry_archivo.delete(0, tk.END)
        entry_archivo.insert(0, path)

def ejecutar():
    excel_path = entry_archivo.get().strip()
    anio = entry_anio.get().strip()
    mes = entry_mes.get().strip()

    if not excel_path or not anio or not mes:
        messagebox.showerror("Error", "Completa archivo, año y mes.")
        return

    try:
        m_int = int(mes)
        if m_int < 1 or m_int > 12:
            raise ValueError("Mes debe estar entre 1 y 12.")
        int(anio)
    except Exception as e:
        messagebox.showerror("Error", f"Revisa año/mes: {e}")
        return

    try:
        excel_balance, nombre_balance_mes, arbol, arbol_nombre, string_concatenado, nombre_string = \
            procesamiento_archivos(excel_path, anio, m_int)

        carpeta = filedialog.askdirectory(title="Selecciona carpeta para guardar resultados")
        if not carpeta:
            return

        # Excel
        with open(os.path.join(carpeta, nombre_balance_mes), "wb") as f:
            f.write(excel_balance.getvalue())

        # XML
        with open(os.path.join(carpeta, arbol_nombre), "wb") as f:
            f.write(arbol.getvalue())

        # TXT (string concatenado)
        with open(os.path.join(carpeta, nombre_string), "w", encoding="utf-8") as f:
            f.write(string_concatenado.getvalue())

        messagebox.showinfo("Éxito", "Archivos generados correctamente.")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {e}")

# --- Interfaz ---
root = tk.Tk()
root.title("Procesamiento de Archivos")
root.geometry("520x270")

tk.Label(root, text="Archivo Excel:").pack(pady=4)
entry_archivo = tk.Entry(root, width=60); entry_archivo.pack(pady=2)
tk.Button(root, text="Seleccionar...", command=seleccionar_archivo).pack(pady=4)

tk.Label(root, text="Año:").pack(pady=2)
entry_anio = tk.Entry(root, width=12); entry_anio.pack(pady=2)

tk.Label(root, text="Mes (1-12):").pack(pady=2)
entry_mes = tk.Entry(root, width=12); entry_mes.pack(pady=2)

tk.Button(root, text="Ejecutar", command=ejecutar).pack(pady=12)

root.mainloop()

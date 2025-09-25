import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from conciliaciones import ejecucion_programa

def seleccionar_archivo_banco():
    ruta = filedialog.askopenfilename(filetypes=[("Archivos PDF", "*.pdf")])
    if ruta:
        entry_banco.delete(0, tk.END)
        entry_banco.insert(0, ruta)

def seleccionar_archivo_baan():
    ruta = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx;*.xls")])
    if ruta:
        entry_baan.delete(0, tk.END)
        entry_baan.insert(0, ruta)

def ejecutar():
    ruta_banco = entry_banco.get()
    ruta_baan = entry_baan.get()
    if not ruta_banco or not ruta_baan:
        messagebox.showwarning("Error", "Debes seleccionar ambos archivos.")
        return
    
    try:
        (lista_cargos_bancarios,
         lista_cargos_conta,
         lista_abonos_bancarios,
         lista_abonos_conta,
         cargos_bancarios_extraidos,
         abonos_bancarios_extraidos) = ejecucion_programa(ruta_banco, ruta_baan)

        # Resumen
        resumen = ""
        resumen += f"CARGOS DEL BANCO NO RECONOCIDOS:\n  Cant. Mov: {len(lista_cargos_bancarios)}\n  Suma: {sum(lista_cargos_bancarios):,.2f}\n\n"
        resumen += f"NUESTROS CARGOS NO RECONOCIDOS:\n  Cant. Mov: {len(lista_cargos_conta)}\n  Suma: {sum(lista_cargos_conta):,.2f}\n\n"
        resumen += f"ABONOS DEL BANCO NO RECONOCIDOS:\n  Cant. Mov: {len(lista_abonos_bancarios)}\n  Suma: {sum(lista_abonos_bancarios):,.2f}\n\n"
        resumen += f"NUESTROS ABONOS NO RECONOCIDOS:\n  Cant. Mov: {len(lista_abonos_conta)}\n  Suma: {sum(lista_abonos_conta):,.2f}\n\n"
        resumen += "=====VERIFICAR QUE ESTO COINCIDA============\n"
        resumen += f"Cargos bancarios extraidos :\n  Cant. Mov: {len(cargos_bancarios_extraidos)}\n  Suma: {sum(cargos_bancarios_extraidos):,.2f}\n\n"
        resumen += f"Abonos bancarios extraidos:\n  Cant. Mov: {len(abonos_bancarios_extraidos)}\n  Suma: {sum(abonos_bancarios_extraidos):,.2f}\n\n"

        text_resumen.delete("1.0", tk.END)
        text_resumen.insert(tk.END, resumen)

        # Listas completas
        text_cargos_banco.delete("1.0", tk.END)
        text_cargos_banco.insert(tk.END, lista_cargos_bancarios)

        text_cargos_conta.delete("1.0", tk.END)
        text_cargos_conta.insert(tk.END, lista_cargos_conta)

        text_abonos_banco.delete("1.0", tk.END)
        text_abonos_banco.insert(tk.END, lista_abonos_bancarios)

        text_abonos_conta.delete("1.0", tk.END)
        text_abonos_conta.insert(tk.END, lista_abonos_conta)

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un problema:\n{str(e)}")

# Función auxiliar para crear cuadros con scrollbar
def crear_texto_scroll(parent, width=90, height=5):
    frame = tk.Frame(parent)
    text_widget = tk.Text(frame, width=width, height=height, wrap="word")
    scroll = tk.Scrollbar(frame, command=text_widget.yview)
    text_widget.configure(yscrollcommand=scroll.set)
    text_widget.pack(side="left", fill="both", expand=True)
    scroll.pack(side="right", fill="y")
    frame.pack(fill="both", expand=True, padx=10, pady=5)
    return text_widget

# Interfaz
root = tk.Tk()
root.title("Conciliació¿on bancaria")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

tk.Label(frame, text="Estado de cuenta (PDF):").grid(row=0, column=0, sticky="w")
entry_banco = tk.Entry(frame, width=50)
entry_banco.grid(row=0, column=1)
tk.Button(frame, text="Buscar", command=seleccionar_archivo_banco).grid(row=0, column=2, padx=5)

tk.Label(frame, text="Archivo del Baan (Excel):").grid(row=1, column=0, sticky="w")
entry_baan = tk.Entry(frame, width=50)
entry_baan.grid(row=1, column=1)
tk.Button(frame, text="Buscar", command=seleccionar_archivo_baan).grid(row=1, column=2, padx=5)

tk.Button(root, text="Ejecutar Conciliacion", command=ejecutar).pack(pady=10)

# Crear Notebook (pestañas)
notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill="both", padx=10, pady=10)

# Pestaña Resumen
frame_resumen = tk.Frame(notebook)
notebook.add(frame_resumen, text="Resumen")
text_resumen = tk.Text(frame_resumen, width=90, height=30, wrap="word")
text_resumen.pack(padx=10, pady=10, fill="both", expand=True)

# Pestaña Listas
frame_listas = tk.Frame(notebook)
notebook.add(frame_listas, text="Listado de movimientos")

# Cada lista en su propio cuadro con scrollbar
tk.Label(frame_listas, text="Cargos del banco no reconocidos:").pack(anchor="w")
text_cargos_banco = crear_texto_scroll(frame_listas)

tk.Label(frame_listas, text="Nuestros cargos no reconocidos:").pack(anchor="w")
text_cargos_conta = crear_texto_scroll(frame_listas)

tk.Label(frame_listas, text="Abonos del banco no reconocidos:").pack(anchor="w")
text_abonos_banco = crear_texto_scroll(frame_listas)

tk.Label(frame_listas, text="Nuestros abonos no reconocidos:").pack(anchor="w")
text_abonos_conta = crear_texto_scroll(frame_listas)

root.mainloop()

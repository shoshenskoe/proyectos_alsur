import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from conciliaciones import ejecucion_programa

# ===================== FUNCIONES =====================
def seleccionar_archivo_banco():
    ruta = filedialog.askopenfilename(filetypes=[("Archivos PDF", "*.pdf")])
    if ruta:
        entry_banco.delete(0, "end")
        entry_banco.insert(0, ruta)

def seleccionar_archivo_baan():
    ruta = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx;*.xls")])
    if ruta:
        entry_baan.delete(0, "end")
        entry_baan.insert(0, ruta)

def ejecutar():
    ruta_banco = entry_banco.get()
    ruta_baan = entry_baan.get()
    if not ruta_banco or not ruta_baan:
        messagebox.showwarning("⚠️ Error", "Debes seleccionar ambos archivos.")
        return
    
    try:
        (lista_cargos_bancarios,
         lista_cargos_conta,
         lista_abonos_bancarios,
         lista_abonos_conta,
         cargos_bancarios_extraidos,
         abonos_bancarios_extraidos) = ejecucion_programa(ruta_banco, ruta_baan)

        resumen = ""
        resumen += f"💳 CARGOS DEL BANCO NO RECONOCIDOS:\n  Cant. Mov: {len(lista_cargos_bancarios)}\n  Suma: {sum(lista_cargos_bancarios):,.2f}\n\n"
        resumen += f"📕 NUESTROS CARGOS NO RECONOCIDOS:\n  Cant. Mov: {len(lista_cargos_conta)}\n  Suma: {sum(lista_cargos_conta):,.2f}\n\n"
        resumen += f"💰 ABONOS DEL BANCO NO RECONOCIDOS:\n  Cant. Mov: {len(lista_abonos_bancarios)}\n  Suma: {sum(lista_abonos_bancarios):,.2f}\n\n"
        resumen += f"📗 NUESTROS ABONOS NO RECONOCIDOS:\n  Cant. Mov: {len(lista_abonos_conta)}\n  Suma: {sum(lista_abonos_conta):,.2f}\n\n"
        resumen += "=============================\n"
        resumen += f"🔎 Cargos bancarios extraídos :\n  Cant. Mov: {len(cargos_bancarios_extraidos)}\n  Suma: {sum(cargos_bancarios_extraidos):,.2f}\n\n"
        resumen += f"🔎 Abonos bancarios extraídos:\n  Cant. Mov: {len(abonos_bancarios_extraidos)}\n  Suma: {sum(abonos_bancarios_extraidos):,.2f}\n\n"

        text_resumen.delete("1.0", "end")
        text_resumen.insert("end", resumen)

        text_cargos_banco.delete("1.0", "end")
        text_cargos_banco.insert("end", lista_cargos_bancarios)

        text_cargos_conta.delete("1.0", "end")
        text_cargos_conta.insert("end", lista_cargos_conta)

        text_abonos_banco.delete("1.0", "end")
        text_abonos_banco.insert("end", lista_abonos_bancarios)

        text_abonos_conta.delete("1.0", "end")
        text_abonos_conta.insert("end", lista_abonos_conta)

    except Exception as e:
        messagebox.showerror("❌ Error", f"Ocurrió un problema:\n{str(e)}")

def crear_texto_scroll(parent, height=6):
    frame = tb.Frame(parent, bootstyle="secondary")
    text_widget = tb.ScrolledText(frame, height=height, font=("Consolas", 10))
    text_widget.pack(fill="both", expand=True)
    frame.pack(fill="both", expand=True, padx=10, pady=5)
    return text_widget

# ===================== INTERFAZ =====================
app = tb.Window(themename="superhero")  # puedes probar: flatly, darkly, cyborg, journal...
app.title("✨ Conciliación Bancaria ✨")
app.geometry("1000x700")

# Encabezado
header = tb.Label(app, text="📊 Sistema de Conciliación Bancaria",
                  bootstyle="inverse-primary", font=("Segoe UI", 18, "bold"), anchor="center")
header.pack(fill="x", pady=5)

# Frame selección de archivos
frame = tb.Labelframe(app, text="Selecciona los archivos", bootstyle="primary")
frame.pack(padx=15, pady=15, fill="x")

tb.Label(frame, text="📄 Estado de cuenta (PDF):").grid(row=0, column=0, sticky="w", pady=5)
entry_banco = tb.Entry(frame, width=70)
entry_banco.grid(row=0, column=1, padx=5)
tb.Button(frame, text="Buscar", bootstyle="info", command=seleccionar_archivo_banco).grid(row=0, column=2, padx=5)

tb.Label(frame, text="📊 Archivo del Baan (Excel):").grid(row=1, column=0, sticky="w", pady=5)
entry_baan = tb.Entry(frame, width=70)
entry_baan.grid(row=1, column=1, padx=5)
tb.Button(frame, text="Buscar", bootstyle="info", command=seleccionar_archivo_baan).grid(row=1, column=2, padx=5)

tb.Button(app, text="🚀 Ejecutar Conciliación", bootstyle="success-outline", command=ejecutar).pack(pady=15)

# Notebook
notebook = tb.Notebook(app, bootstyle="primary")
notebook.pack(expand=True, fill="both", padx=15, pady=10)

# Pestaña Resumen
frame_resumen = tb.Frame(notebook)
notebook.add(frame_resumen, text="📑 Resumen")
text_resumen = tb.ScrolledText(frame_resumen, font=("Segoe UI", 11))
text_resumen.pack(padx=10, pady=10, fill="both", expand=True)

# Pestaña Listas
frame_listas = tb.Frame(notebook)
notebook.add(frame_listas, text="📋 Movimientos")

tb.Label(frame_listas, text="💳 Cargos del banco no reconocidos:", font=("Segoe UI", 11, "bold")).pack(anchor="w")
text_cargos_banco = crear_texto_scroll(frame_listas)

tb.Label(frame_listas, text="📕 Nuestros cargos no reconocidos:", font=("Segoe UI", 11, "bold")).pack(anchor="w")
text_cargos_conta = crear_texto_scroll(frame_listas)

tb.Label(frame_listas, text="💰 Abonos del banco no reconocidos:", font=("Segoe UI", 11, "bold")).pack(anchor="w")
text_abonos_banco = crear_texto_scroll(frame_listas)

tb.Label(frame_listas, text="📗 Nuestros abonos no reconocidos:", font=("Segoe UI", 11, "bold")).pack(anchor="w")
text_abonos_conta = crear_texto_scroll(frame_listas)

app.mainloop()

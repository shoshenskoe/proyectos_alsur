import customtkinter as ctk
from tkinter import filedialog, messagebox
from conciliaciones import ejecucion_programa

# ===================== CONFIG GLOBAL =====================
ctk.set_appearance_mode("system")   # "light", "dark" o "system"
ctk.set_default_color_theme("blue") # también: "green", "dark-blue"

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

        text_resumen.delete("0.0", "end")
        text_resumen.insert("end", resumen)

        text_cargos_banco.delete("0.0", "end")
        text_cargos_banco.insert("end", lista_cargos_bancarios)

        text_cargos_conta.delete("0.0", "end")
        text_cargos_conta.insert("end", lista_cargos_conta)

        text_abonos_banco.delete("0.0", "end")
        text_abonos_banco.insert("end", lista_abonos_bancarios)

        text_abonos_conta.delete("0.0", "end")
        text_abonos_conta.insert("end", lista_abonos_conta)

    except Exception as e:
        messagebox.showerror("❌ Error", f"Ocurrió un problema:\n{str(e)}")

def crear_texto_scroll(parent, height=8):
    text_widget = ctk.CTkTextbox(parent, height=height*20, font=("Consolas", 12))
    text_widget.pack(fill="both", expand=True, pady=5)
    return text_widget

# ===================== INTERFAZ =====================
app = ctk.CTk()
app.title("✨ Conciliación Bancaria ✨")
app.geometry("1000x700")

# Encabezado
header = ctk.CTkLabel(app, text="📊 Sistema de Conciliación Bancaria",
                      font=("Segoe UI", 20, "bold"))
header.pack(pady=15)

# Frame selección de archivos
frame = ctk.CTkFrame(app, corner_radius=12)
frame.pack(padx=20, pady=15, fill="x")

label1 = ctk.CTkLabel(frame, text="📄 Estado de cuenta (PDF):", font=("Segoe UI", 12))
label1.grid(row=0, column=0, sticky="w", pady=10, padx=10)
entry_banco = ctk.CTkEntry(frame, width=500)
entry_banco.grid(row=0, column=1, padx=10)
btn_banco = ctk.CTkButton(frame, text="Buscar", command=seleccionar_archivo_banco)
btn_banco.grid(row=0, column=2, padx=10)

label2 = ctk.CTkLabel(frame, text="📊 Archivo del Baan (Excel):", font=("Segoe UI", 12))
label2.grid(row=1, column=0, sticky="w", pady=10, padx=10)
entry_baan = ctk.CTkEntry(frame, width=500)
entry_baan.grid(row=1, column=1, padx=10)
btn_baan = ctk.CTkButton(frame, text="Buscar", command=seleccionar_archivo_baan)
btn_baan.grid(row=1, column=2, padx=10)

btn_run = ctk.CTkButton(app, text="🚀 Ejecutar Conciliación", fg_color="green",
                        hover_color="#0b6623", command=ejecutar, height=40, width=250)
btn_run.pack(pady=20)

# Tabs
tabview = ctk.CTkTabview(app, width=900, height=450, corner_radius=12)
tabview.pack(expand=True, fill="both", padx=20, pady=10)

tab_resumen = tabview.add("📑 Resumen")
tab_listas = tabview.add("📋 Movimientos")

# Resumen
text_resumen = ctk.CTkTextbox(tab_resumen, font=("Segoe UI", 12))
text_resumen.pack(fill="both", expand=True, padx=10, pady=10)

# Listas
ctk.CTkLabel(tab_listas, text="💳 Cargos del banco no reconocidos:", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=10)
text_cargos_banco = crear_texto_scroll(tab_listas)

ctk.CTkLabel(tab_listas, text="📕 Nuestros cargos no reconocidos:", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=10)
text_cargos_conta = crear_texto_scroll(tab_listas)

ctk.CTkLabel(tab_listas, text="💰 Abonos del banco no reconocidos:", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=10)
text_abonos_banco = crear_texto_scroll(tab_listas)

ctk.CTkLabel(tab_listas, text="📗 Nuestros abonos no reconocidos:", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=10)
text_abonos_conta = crear_texto_scroll(tab_listas)

# ===================== LOOP =====================
app.mainloop()

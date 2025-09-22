# --- PASO 1: Importar las herramientas que necesitamos ---
import tkinter as tk
from tkinter import filedialog, messagebox

# Importamos la función principal de tu otro archivo
from codigo_reclasificacion import reclasificacion

# --- PASO 2: Definir qué hacen los botones ---

def seleccionar_archivo_ingresos():
    """Esta función se activa con el primer botón.
       Abre una ventana para que elijas el archivo de ingresos."""
    # Abre el explorador de archivos para seleccionar un .xlsx
    ruta_archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
    # Si el usuario selecciona un archivo, lo ponemos en la caja de texto
    if ruta_archivo:
        caja_texto_ingresos.delete(0, tk.END)  # Borra el texto anterior
        caja_texto_ingresos.insert(0, ruta_archivo) # Escribe la nueva ruta

def seleccionar_archivo_gastos():
    """Esta función se activa con el segundo botón.
       Abre una ventana para que elijas el archivo de gastos."""
    ruta_archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
    if ruta_archivo:
        caja_texto_gastos.delete(0, tk.END)
        caja_texto_gastos.insert(0, ruta_archivo)

def procesar():
    """Esta es la función principal. Se activa con el botón 'Procesar'.
       Recoge los datos, ejecuta tu código y guarda el resultado."""
    
    # Obtenemos las rutas y textos de las cajas
    ingresos = caja_texto_ingresos.get()
    gastos = caja_texto_gastos.get()
    mes_nombre = caja_texto_mes.get()
    mes_num = caja_texto_mes_numero.get()

    # Revisamos si algún campo está vacío
    if not ingresos or not gastos or not mes_nombre or not mes_num:
        messagebox.showwarning("Atención", "Todos los campos son obligatorios.")
        return # Detenemos la función si falta algo

    # Si todo está bien, le decimos al usuario que estamos trabajando
    etiqueta_estado.config(text="Procesando... por favor espera.")
    ventana.update() # Forzamos la actualización de la ventana

    try:
        # Ejecutamos tu función de reclasificación
        archivo_en_memoria, nombre_sugerido = reclasificacion(ingresos, gastos, mes_nombre, mes_num)

        # Preguntamos al usuario dónde quiere guardar el archivo nuevo
        ruta_para_guardar = filedialog.asksaveasfilename(
            initialfile=nombre_sugerido, # Le sugerimos un nombre
            defaultextension=".xlsx",
            filetypes=[("Archivos de Excel", "*.xlsx")]
        )

        # Si el usuario eligió una ubicación, guardamos el archivo
        if ruta_para_guardar:
            with open(ruta_para_guardar, "wb") as f:
                f.write(archivo_en_memoria.getbuffer())
            messagebox.showinfo("Éxito", f"¡Proceso completado!\nArchivo guardado en: {ruta_para_guardar}")
        
    except Exception as e:
        # Si algo sale mal en tu script, mostramos un error
        messagebox.showerror("Error", f"Ocurrió un error inesperado:\n{e}")

    finally:
        # Al final, limpiamos el mensaje de estado
        etiqueta_estado.config(text="")


# --- PASO 3: Crear la ventana y los elementos (widgets) ---

# Creamos la ventana principal
ventana = tk.Tk()
ventana.title("Herramienta de Reclasificación")
ventana.geometry("550x300") # Ancho x Alto

# Creamos un marco para organizar los elementos
marco = tk.Frame(ventana, padx=15, pady=15)
marco.pack()

# --- Elementos para el archivo de INGRESOS ---
etiqueta_ingresos = tk.Label(marco, text="Archivo de Ingresos:")
etiqueta_ingresos.grid(row=0, column=0, sticky="w", pady=5)

caja_texto_ingresos = tk.Entry(marco, width=50)
caja_texto_ingresos.grid(row=0, column=1, padx=5)

boton_ingresos = tk.Button(marco, text="Seleccionar", command=seleccionar_archivo_ingresos)
boton_ingresos.grid(row=0, column=2)

# --- Elementos para el archivo de GASTOS ---
etiqueta_gastos = tk.Label(marco, text="Archivo de Gastos:")
etiqueta_gastos.grid(row=1, column=0, sticky="w", pady=5)

caja_texto_gastos = tk.Entry(marco, width=50)
caja_texto_gastos.grid(row=1, column=1, padx=5)

boton_gastos = tk.Button(marco, text="Seleccionar", command=seleccionar_archivo_gastos)
boton_gastos.grid(row=1, column=2)

# --- Elementos para el MES y NÚMERO DE MES ---
etiqueta_mes = tk.Label(marco, text="Mes (ej. Enero):")
etiqueta_mes.grid(row=2, column=0, sticky="w", pady=5)

caja_texto_mes = tk.Entry(marco, width=50)
caja_texto_mes.grid(row=2, column=1, padx=5)

etiqueta_mes_numero = tk.Label(marco, text="Número de Mes (ej. 01):")
etiqueta_mes_numero.grid(row=3, column=0, sticky="w", pady=5)

caja_texto_mes_numero = tk.Entry(marco, width=50)
caja_texto_mes_numero.grid(row=3, column=1, padx=5)

# --- Botón para PROCESAR ---
boton_procesar = tk.Button(marco, text="▶️ Procesar Reclasificación", bg="green", fg="white", font=("Helvetica", 10, "bold"), command=procesar)
boton_procesar.grid(row=4, column=1, pady=20, sticky="ew") # sticky="ew" hace que se estire a lo ancho

# --- Etiqueta de ESTADO ---
etiqueta_estado = tk.Label(marco, text="", fg="blue")
etiqueta_estado.grid(row=5, column=1)

# --- PASO 4: Iniciar la aplicación ---
# Esto mantiene la ventana abierta y esperando a que el usuario haga algo
ventana.mainloop()
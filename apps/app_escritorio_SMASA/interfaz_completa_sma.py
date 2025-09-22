import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
from io import BytesIO
import threading

#funciones auxiliares

def excel_limpieza_ingresos_gastos (df) :
  df.rename(columns={'cont.        ':'cuenta', '     ':'subn', '                                ':'Descripcion',
        '                    ':'Saldoapertura', '                Debe':'Debe', '               Haber':'Haber',
        '                Debe.1':'Debe1', '               Haber.1':'Haber1', 'Unnamed: 8':'Ingreso'}, inplace= True)
  df=df.query("Descripcion.isnull()== False")
  df["cuenta"]=df["cuenta"].astype("str")
  df=df[['cuenta', 'Descripcion', 'Saldoapertura', 'Debe', 'Haber',
        'Debe1', 'Haber1', 'Ingreso']]
  dflimpio=df.query(' cuenta.str.startswith("5") | cuenta.str.startswith("6")  ' )
  dflimpio=dflimpio.reset_index(drop= True)
  dfisr=df.copy()
  dfisr["cuenta"]=df["cuenta"].str.replace(" ","")
  ISR=dfisr.query('cuenta== "66010101IR01" ')
  ISR=ISR.reset_index()
  ISR=ISR.iloc[0,6]
  PTU=dfisr.query('cuenta== "66010101IR02" ')
  PTU=PTU.reset_index()
  PTU=PTU.iloc[0,6]
  return dflimpio ,  ISR, PTU

def excel_tabla_cc_sino (df2, mes) :
  df2.rename(columns={'Cta. cont.   ': "Cuenta", '      ': "Sub", '                                   ': "Descripcion",
        '                    ':"Saldoapertura", '               Debe ':"Debe", '              Haber ':"Haber",
        '               Debe .1':"Debe1", '              Haber .1':"Haber1",
        '    Saldo de cierre':"Saldocierre"}, inplace= True)
  df2["Cuenta"]=df2["Cuenta"].str.replace("'","")
  df2=df2.query('not ( Cuenta.str.startswith("-")  | Cuenta.str.startswith("Total"))')
  df2.reset_index(drop= True, inplace= True)
  Bodega=[]
  a=""
  for i in range(len(df2)):
    if df2.iloc[i,0].startswith("6"):
      Bodega.append(a)
    elif df2.iloc[i,0].startswith("5"):
       Bodega.append(a)
    else:
      a=df2.iloc[i,0]
      Bodega.append(a)

  df2["cc"]=Bodega
  df2["Cuenta"]=df2["Cuenta"]+"_"+df2["Descripcion"]
  df2=df2.query(' Cuenta.isnull()= =False')
  df2.reset_index(drop= True,inplace= True)
  sino=[]
  for i in range(len(df2)):
    if df2.iloc[i,0].replace(" ","").startswith("641001010001_SUELDOSYPRESTACIONES"):
      sino.append("sueldo")
    else:
      sino.append("extra")
  df2["sino"]=sino
  df3=df2.copy()
  df3=df3[["cc","sino","Cuenta","Debe1"]]
  df3.rename(columns={"Debe1":mes}, inplace= True)
  df3[mes]=df3[mes].replace('                    ',"0")
  for i in range(len(df3)):
    if str(df3.loc[i, mes]).strip().endswith('-'):
      a=str(df3.loc[i, mes]).strip()
      df3.loc[i, mes]= float(a[:-2])*-1
  df3[mes]=df3[mes].astype("float64")
  df3=df3.loc[df3[mes] !=0]
  return df3

def tabla_din_1 ( dataframe_cc_sino, mes ):
  tabladin=pd.pivot_table( dataframe_cc_sino , index="cc", columns="sino", values= mes, aggfunc="sum")
  tabladin.reset_index(inplace=True)
  tabladin.fillna(0, inplace= True)
  return tabladin

def tabla_din_isr( tabla_dinamica_1 , ISR, PTU):
  tabladin2=tabla_dinamica_1.copy()
  tabladin2["isr-ptu"]= tabla_dinamica_1["extra"]*(ISR+PTU)/(tabladin2["extra"].sum())
  tabladin2["Total de extras"]=tabladin2["extra"]+tabladin2["isr-ptu"]
  tabladin2["Total general"]= tabladin2["Total de extras"]+tabladin2["sueldo"]
  S=pd.Series({"cc":"Total","extra":tabladin2["extra"].sum(),"sueldo":tabladin2["sueldo"].sum(), "isr-ptu":tabladin2["isr-ptu"].sum(),"Total de extras":tabladin2["Total de extras"].sum(),"Total general":tabladin2["Total general"].sum()})
  tabladin2=pd.concat([tabladin2, S.to_frame().T], ignore_index= True)
  return tabladin2

def tabla_din_3( tabla_dinamica_2):
  tabladin3=tabla_dinamica_2[["extra", "isr-ptu", "Total de extras", "sueldo", "Total general"]].astype(float)
  tabladin3=tabladin3.round(0)
  tabladin3["cc"]=tabla_dinamica_2["cc"]
  tabladin3=tabladin3[["cc","extra", "isr-ptu", "Total de extras", "sueldo", "Total general"]]
  return tabladin3

def consolidar (tabla_din_3, tabla_utilitario):
  df_parapoliza=tabla_din_3.copy()
  df_parapoliza["cc"]=df_parapoliza["cc"].str.strip()
  df_parapoliza[["cc","unidad"]]=df_parapoliza["cc"].str.split(" 0 ",expand= True)
  df_parapoliza.drop(columns=["unidad"], inplace= True)
  df_parapoliza.rename(columns={'cc':'CC'}, inplace= True)
  df_parapoliza['CC']=df_parapoliza['CC'].str.strip()
  df_parapoliza=pd.merge(df_parapoliza,tabla_utilitario, on=["CC"], how="left")
  df_parapoliza['CC']=df_parapoliza['CC'].replace("Total", "")
  return df_parapoliza



def generador_poliza_final (poliza_utilitario, Referencia):
  CF_sueldo='64913101CF01'
  GA_sueldo='64913101GA01'
  GV_sueldo='64913101GV01'
  Total_sueldo= '240112900026'
  CF_extra='64913101CF04'
  GA_extra='64913101GA04'
  GV_extra='64913101GV04'
  Total_extra= '240112900026'
  df_parapoliza_sueldo=poliza_utilitario[['CC','UTILITARIO' ,'sueldo']].copy()
  df_parapoliza_extra=poliza_utilitario[['CC','UTILITARIO' , 'Total de extras']].copy()
  df_parapoliza_sueldo.insert(0, 'plantilla', ['CAM' for i in range(len(df_parapoliza_sueldo))] )
  df_parapoliza_sueldo.insert(1, 'cons', [i+1 for i in range(len(df_parapoliza_sueldo))] )
  df_parapoliza_sueldo.insert(2, 'comp', ['100' for i in range(len(df_parapoliza_sueldo))] )
  cta=[]
  for i in range(len(df_parapoliza_sueldo)):
    CC= df_parapoliza_sueldo.loc[i, 'CC']
    if not  CC.startswith('E')  and  CC != "D981" and CC !="":
      cta.append(CF_sueldo)
    elif CC.startswith('E') :
      cta.append(GA_sueldo)
    elif CC=="D981" and CC !="" :
      cta.append(GV_sueldo)
    elif  CC =="" :
      cta.append(Total_sueldo)
  df_parapoliza_sueldo.insert(3, 'cta', cta)
  df_parapoliza_sueldo.insert(6, 'nada', [np.nan for i in range(len(df_parapoliza_sueldo))] )
  df_parapoliza_sueldo.insert(7, 'nada1', [np.nan for i in range(len(df_parapoliza_sueldo))] )
  df_parapoliza_sueldo.insert(8, 'debe/haber', [1 if df_parapoliza_sueldo.loc[i, 'cta'] .startswith('6') or df_parapoliza_sueldo.loc[i, 'cta'] .startswith('1')  else 2 for i in range(len(df_parapoliza_sueldo))] )
  df_parapoliza_sueldo["ref"]= [ 'SMASA ' + Referencia+ " -24 Sueldos"  for i in range(len(df_parapoliza_sueldo))]
  df_parapoliza_sueldo["CC"]=df_parapoliza_sueldo["CC"].replace("", np.nan)
  df_parapoliza_sueldo.rename(columns={'sueldo':'total'}, inplace= True)
  df_parapoliza_extra.insert(0, 'plantilla', ['CAM' for i in range(len(df_parapoliza_extra))] )
  df_parapoliza_extra.insert(1, 'cons', [i+1 for i in range(len(df_parapoliza_extra))] )
  df_parapoliza_extra.insert(2, 'comp', ['100' for i in range(len(df_parapoliza_extra))] )
  cta=[]
  for i in range(len(df_parapoliza_extra)):
    CC= df_parapoliza_extra.loc[i, 'CC']
    if not  CC.startswith('E')  and  CC != "D981" and CC !="":
      cta.append(CF_extra)
    elif CC.startswith('E') :
      cta.append(GA_extra)
    elif CC=="D981" and CC !="" :
      cta.append(GV_extra)
    elif  CC =="" :
      cta.append(Total_extra)
  df_parapoliza_extra.insert(3, 'cta', cta)
  df_parapoliza_extra.insert(6, 'nada', [np.nan for i in range(len(df_parapoliza_extra))] )
  df_parapoliza_extra.insert(7, 'nada1', [np.nan for i in range(len(df_parapoliza_extra))] )
  df_parapoliza_extra.insert(8, 'debe/haber', [1 if df_parapoliza_extra.loc[i, 'cta'] .startswith('6') or df_parapoliza_extra.loc[i, 'cta'] .startswith('1')  else 2 for i in range(len(df_parapoliza_extra))] )
  df_parapoliza_extra["ref"]= [ 'SMASA ' + Referencia+ " -24 Extras"  for i in range(len(df_parapoliza_extra))]
  df_parapoliza_extra["CC"]=df_parapoliza_extra["CC"].replace("", np.nan)
  df_parapoliza_extra.rename(columns={'Total de extras':'total'}, inplace= True)
  df_parapoliza = pd.concat([df_parapoliza_sueldo, df_parapoliza_extra], ignore_index=True).copy()
  df_parapoliza['cons']= [ i+1 for i in range(len(df_parapoliza))]
  return df_parapoliza


#Interfaz
class UtilitarioDialog(tk.Toplevel):
    """
    Esta clase crea la ventana emergente que pide al usuario
    los datos del centro utilitario que faltan.
    """
    def __init__(self, parent, datos_faltantes):
        super().__init__(parent)
        self.transient(parent)
        self.title("Completar Datos Faltantes")
        self.parent = parent
        self.datos_faltantes = datos_faltantes
        self.nuevos_utilitarios = {}
        self.current_index = 0

        # Widgets
        self.label_info = ttk.Label(self, text="")
        self.label_info.pack(padx=20, pady=10)
        
        self.entry_utilitario = ttk.Entry(self)
        self.entry_utilitario.pack(padx=20, pady=5)
        
        self.submit_button = ttk.Button(self, text="Siguiente", command=self.submit_and_next)
        self.submit_button.pack(pady=10)

        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.ask_next()

    def ask_next(self):
        """Muestra la siguiente CC que necesita un utilitario."""
        if self.current_index < len(self.datos_faltantes):
            cc_actual = self.datos_faltantes.iloc[self.current_index]["CC"]
            self.label_info.config(text=f"Introduce el centro utilitario para CC: {cc_actual}")
            self.entry_utilitario.focus_set()
        else:
            self.finish()

    def submit_and_next(self):
        """Guarda el dato introducido y pasa al siguiente."""
        cc_actual = self.datos_faltantes.iloc[self.current_index]["CC"]
        uti_ingresado = self.entry_utilitario.get().upper().strip()
        
        if not uti_ingresado:
            messagebox.showwarning("Dato Vacío", "Por favor, introduce un valor para el utilitario.", parent=self)
            return

        self.nuevos_utilitarios[cc_actual] = uti_ingresado
        self.entry_utilitario.delete(0, tk.END)
        self.current_index += 1
        self.ask_next()
        
    def finish(self):
        """Cierra la ventana cuando se han completado todos los datos."""
        self.destroy()

    def cancel(self):
        """Se activa si el usuario cierra la ventana."""
        if messagebox.askyesno("Confirmar", "¿Seguro que quieres cancelar el proceso?", parent=self):
            self.nuevos_utilitarios = None  # Marcar como cancelado
            self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Costo de SMASA")
        self.geometry("500x300")

        # Variables para almacenar las rutas de los archivos
        self.path_secuencia = tk.StringVar()
        self.path_dimension = tk.StringVar()

        # --- Creación de Widgets ---
        frame = ttk.Frame(self, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        # Entradas de texto
        ttk.Label(frame, text="Mes (ej: enero):").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.mes_entry = ttk.Entry(frame)
        self.mes_entry.grid(row=0, column=1, sticky=tk.EW, pady=2)
        
        ttk.Label(frame, text="Referencia (ej: ene):").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.ref_entry = ttk.Entry(frame)
        self.ref_entry.grid(row=1, column=1, sticky=tk.EW, pady=2)

        # Botones para seleccionar archivos
        ttk.Label(frame, text="Archivo Ban Secuencia:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.secuencia_label = ttk.Label(frame, text="No seleccionado", foreground="grey")
        self.secuencia_label.grid(row=2, column=1, sticky=tk.EW, padx=(0, 5))
        ttk.Button(frame, text="Buscar...", command=lambda: self.select_file(self.path_secuencia, self.secuencia_label)).grid(row=2, column=2)

        ttk.Label(frame, text="Archivo Ban Dimensión:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.dimension_label = ttk.Label(frame, text="No seleccionado", foreground="grey")
        self.dimension_label.grid(row=3, column=1, sticky=tk.EW, padx=(0, 5))
        ttk.Button(frame, text="Buscar...", command=lambda: self.select_file(self.path_dimension, self.dimension_label)).grid(row=3, column=2)

        # Botón para generar el reporte y etiqueta de estado
        self.generate_button = ttk.Button(frame, text="Generar Excel", command=self.start_process_thread)
        self.generate_button.grid(row=4, column=0, columnspan=3, pady=20)
        
        self.status_label = ttk.Label(frame, text="Listo para empezar")
        self.status_label.grid(row=5, column=0, columnspan=3, pady=10)

        frame.columnconfigure(1, weight=1)

    def select_file(self, path_variable, label_widget):
        """Abre un diálogo para seleccionar un archivo Excel."""
        filepath = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos de Excel", "*.xlsx *.xls")]
        )
        if filepath:
            path_variable.set(filepath)
            # Muestra solo el nombre del archivo, no la ruta completa
            filename = filepath.split("/")[-1]
            label_widget.config(text=filename, foreground="black")
    
    def start_process_thread(self):
        """
        Inicia el procesamiento en un hilo separado para no congelar la GUI.
        """
        # Validaciones de entrada
        if not all([self.mes_entry.get(), self.ref_entry.get(), self.path_secuencia.get(), self.path_dimension.get()]):
            messagebox.showerror("Error", "Todos los campos son obligatorios.")
            return

        self.generate_button.config(state=tk.DISABLED)
        self.status_label.config(text="Procesando, por favor espera...")
        
        # El hilo ejecutará la función 'run_full_process'
        thread = threading.Thread(target=self.run_full_process)
        thread.start()

    def run_full_process(self):
        """
        Esta función contiene la lógica principal del script original.
        Se ejecuta en un hilo secundario.
        """
        try:
            # 1. Obtener datos de la GUI
            mes = self.mes_entry.get().lower()
            referencia = self.ref_entry.get().upper()
            path_secuencia = self.path_secuencia.get()
            path_dimension = self.path_dimension.get()

            # 2. Ejecutar la lógica del script original paso a paso
            df = pd.read_excel(path_secuencia, skiprows=[0, 1, 2, 3, 4, 5, 6, 7, 8, 10])
            dflimpio, ISR, PTU = excel_limpieza_ingresos_gastos(df)

            df2 = pd.read_excel(path_dimension, skiprows=[0, 1, 2, 3, 4, 5, 6, 7, 8, 10])
            df3 = excel_tabla_cc_sino(df2, mes)

            tabladin = tabla_din_1(df3, mes)
            tabladin2 = tabla_din_isr(tabladin, ISR, PTU)
            tabladin3 = tabla_din_3(tabladin2)

            enlace = "https://docs.google.com/spreadsheets/d/1WAfIZJde_M2sYM7SIOyWCPvsMgJoyqVu/export?format=xlsx"

            utilitario = pd.read_excel(enlace)

            df_parapoliza_prototipo = consolidar(tabladin3, utilitario)
            
            # --- Lógica de la ventana emergente ---
            datos_faltantes = df_parapoliza_prototipo.query('UTILITARIO.isnull() == True and CC != ""')
            df_parapoliza_consolidado = df_parapoliza_prototipo.copy()

            if not datos_faltantes.empty:
                dialog = UtilitarioDialog(self, datos_faltantes)
                self.wait_window(dialog) # Espera a que la ventana emergente se cierre

                nuevos_datos = dialog.nuevos_utilitarios
                if nuevos_datos is None: # Si el usuario canceló
                    self.process_finished("Proceso cancelado por el usuario.", success=False)
                    return
                
                # Actualizar el dataframe con los datos ingresados
                for cc, uti in nuevos_datos.items():
                    df_parapoliza_consolidado.loc[df_parapoliza_consolidado['CC'] == cc, 'UTILITARIO'] = uti
                    
                    # Opcional: añadir a la tabla de utilitarios para futuras ejecuciones
                    nuevo_registro = pd.DataFrame([{'CC': cc, 'UTILITARIO': uti}])
                    utilitario = pd.concat([utilitario, nuevo_registro], ignore_index=True)
                    utilitario.to_excel("/content/drive/MyDrive/Costo SMASA cc y su utilitario/CC y su utilitario.xlsx.xlsx", index= False)
                    
            # --- Continuar con el resto del proceso ---
            df_parapoliza = generador_poliza_final(df_parapoliza_consolidado, referencia)
            dfsucio = pd.read_excel(path_secuencia)

            # 3. Guardar el archivo final
            archivo_excel_buffer = BytesIO()
            with pd.ExcelWriter(archivo_excel_buffer, engine='xlsxwriter') as writer:
                dfsucio.to_excel(writer, sheet_name="Sheet1", index=False)
                dflimpio.to_excel(writer, sheet_name="ctas 5-6", index=False)
                df3.to_excel(writer, sheet_name="ctas x cc 6", index=False)
                tabladin.to_excel(writer, sheet_name="poliza", index=False)
                tabladin3.to_excel(writer, sheet_name="poliza", index=False, startcol=6)
                df_parapoliza.to_excel(writer, sheet_name="tfgld013", index=False, header=False)
            
            archivo_excel_buffer.seek(0)
            
            # Pedir al usuario dónde guardar el archivo
            nombre_archivo_sugerido = f"costo empresa {mes} 2024.xlsx"
            ruta_guardado = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=nombre_archivo_sugerido,
                filetypes=[("Archivos de Excel", "*.xlsx")]
            )

            if ruta_guardado:
                with open(ruta_guardado, 'wb') as f:
                    f.write(archivo_excel_buffer.getbuffer())
                self.process_finished(f"Archivo guardado en:\n{ruta_guardado}", success=True)
            else:
                self.process_finished("Guardado cancelado por el usuario.", success=False)

        except Exception as e:
            self.process_finished(f"Ocurrió un error:\n{e}", success=False)
            
    def process_finished(self, message, success=True):
        """Actualiza la GUI cuando el proceso termina (con exito o error)."""
        if success:
            self.status_label.config(text=message, foreground="green")
        else:
            self.status_label.config(text=message, foreground="red")
            messagebox.showerror("Error en el Proceso", message)
            
        self.generate_button.config(state=tk.NORMAL)


if __name__ == "__main__":
    app = App()
    app.mainloop()
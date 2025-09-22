"""
Este script crea una interfaz gráfica para un script de Python. El usuario puede procesar
facturas en formato XML que se encuentren dentro de carpetas comprimidas (ZIP).
Extrae los datos, los consolida en un único archivo Excel y añade una columna
para identificar la carpeta de origen de cada factura.

El script original puede consultarse : https://colab.research.google.com/drive/17J2vx4iTqhi3FFAgDBa1u26j6I9Rb-2v?usp=sharing

Posee cualquier limitación derivada del script original. Las limitaciones de la interfaz no incluye algunos casos de uso extravagantes y el manejo
de excepciones y errores es generico.


Pagina para aprender sobre interfaces graficas y el uso de tkinter en general: https://realpython.com/python-gui-tkinter/

Para crear .exe instalar PyInstaller y despues usar el siguiente comando : python -m PyInstaller --windowed --onefile interfaz_facturas.py
Usamos windowed para que no se abra una ventana de consola al ejecutar el .exe y onefile para que se genere un solo archivo .exe
El .exe generado se encuentra en la carpeta dist
"""

# --- Importación de librerías necesarias ---
import tkinter as tk
from tkinter import filedialog, messagebox
import xml.etree.ElementTree as ET
import pandas as pd
import os
import zipfile
import tempfile # Para crear carpetas temporales que se borran solas
import shutil   # Utilidad para manejo de archivos, aunque no se use directamente aquí

def procesar_archivos_xml():
    """
    Función principal que se ejecuta al pulsar el botón.
    Gestiona la selección de archivos ZIP, su procesamiento y el guardado del reporte final.
    """
    # Abre una ventana para que el usuario seleccione uno o más archivos .zip
    rutas_zip = filedialog.askopenfilenames(
        title="Selecciona las carpetas ZIP a procesar",
        filetypes=[("Archivos ZIP", "*.zip")]
    )

    # Si el usuario no selecciona nada, termina la función.
    if not rutas_zip:
        messagebox.showinfo("Información", "No se seleccionó ningún archivo.")
        return

    # Lista para guardar los datos (en formato DataFrame) de cada archivo ZIP.
    todos_los_datos = []

    # Bucle para procesar cada archivo ZIP seleccionado por el usuario.
    for ruta_zip_actual in rutas_zip:
        # Obtenemos el nombre del archivo ZIP para usarlo como referencia.
        nombre_zip = os.path.basename(ruta_zip_actual)
        
        # Crea un directorio temporal para extraer los archivos de forma segura.
        # Este directorio se eliminará automáticamente al finalizar.
        with tempfile.TemporaryDirectory() as directorio_temporal:
            try:
                # Extrae el contenido del ZIP en la carpeta temporal.
                with zipfile.ZipFile(ruta_zip_actual, 'r') as archivo_zip:
                    archivo_zip.extractall(directorio_temporal)
            except zipfile.BadZipFile:
                messagebox.showerror("Error", f"No se pudo procesar {nombre_zip}. El archivo podría estar dañado.")
                continue # Salta al siguiente archivo ZIP.

            # Busca todos los archivos .xml de forma recursiva (dentro de todas las subcarpetas).
            archivos_xml_encontrados = []
            for directorio_raiz, _, archivos in os.walk(directorio_temporal):
                for archivo in archivos:
                    if archivo.lower().endswith('.xml'):
                        archivos_xml_encontrados.append(os.path.join(directorio_raiz, archivo))
            
            # Si no hay archivos XML, no hay nada que procesar en este ZIP.
            if not archivos_xml_encontrados:
                continue

            # Procesa la lista de archivos XML y la añade a nuestros datos generales.
            dataframe_del_zip = extraer_datos_de_un_zip(archivos_xml_encontrados, nombre_zip)
            todos_los_datos.append(dataframe_del_zip)

    # Si después de revisar todos los ZIPs no se extrajo ningún dato, informa al usuario.
    if not todos_los_datos:
        messagebox.showinfo("Información", "No se encontraron facturas válidas en los archivos seleccionados.")
        return

    # Combina los datos de todos los ZIPs en un único DataFrame.
    dataframe_final = pd.concat(todos_los_datos, ignore_index=True)

    # Pide al usuario la ubicación y el nombre para guardar el archivo Excel.
    ruta_guardado = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Archivos Excel", "*.xlsx")],
        title="Guardar el reporte como"
    )

    # Si el usuario cancela la ventana de guardado, termina la operación.
    if not ruta_guardado:
        messagebox.showinfo("Información", "Operación de guardado cancelada.")
        return

    # Intenta guardar el DataFrame en un archivo Excel.
    try:
        dataframe_final.to_excel(ruta_guardado, index=False)
        messagebox.showinfo("Éxito", f"Archivo guardado correctamente en:\n{ruta_guardado}")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al guardar el archivo: {e}")


def extraer_datos_de_un_zip(lista_archivos_xml, nombre_zip_origen):
    """
    Recibe una lista de rutas de archivos XML y el nombre del ZIP de donde provienen.
    Extrae los datos de cada XML y los devuelve en un DataFrame de pandas.
    """

    # Define las columnas para el DataFrame que almacenará los datos.
    columnas = ['FACTURA', 'FECHA', 'RFC', 'NOMBRE', 'SUBTOTAL', 'IVA', 'RETENCIÓN', 'TOTAL', 'UUID', "REFERENCIA",  'ORIGEN_CARPETA']
    datos_xml_df = pd.DataFrame(columns=columnas)

    # Define los "espacios de nombres" (namespaces) del XML para encontrar los datos correctamente.
    espacios_nombres = {
        'cfdi': 'http://www.sat.gob.mx/cfd/4',
        'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital',
    }

    # Recorre cada ruta de archivo XML.
    for ruta_xml in lista_archivos_xml:
        try:
            # Parsea (lee e interpreta) el archivo XML.
            arbol = ET.parse(ruta_xml)
            raiz = arbol.getroot()

            # Busca el nodo del Timbre Fiscal Digital, que contiene datos clave como el UUID.
            complemento = raiz.find('.//tfd:TimbreFiscalDigital', espacios_nombres)
            if complemento is None:
                continue # Si no tiene timbre, no es una factura válida.

            # Procesa solo facturas de Ingreso ('I') o Egreso ('E') que tengan impuestos.
            if (raiz.get('TipoDeComprobante') in ['I', 'E']) and (raiz.find('cfdi:Impuestos', espacios_nombres) is not None):
                emisor = raiz.find('cfdi:Emisor', espacios_nombres)
                
                # Extrae Serie, Folio y UUID.
                serie = str(raiz.get('Serie'))
                folio = str(raiz.get('Folio'))
                valor_uuid = str(complemento.get("UUID"))

                # Construye el número de factura o usa los últimos 12 dígitos del UUID.
                if folio != "None":
                    factura = f"{serie} {folio}" if serie != "None" else str(folio)
                else:
                    factura = valor_uuid[-12:]

                # Extrae impuestos de forma segura (si no existen, pone none).
                nodo_impuestos = raiz.find('cfdi:Impuestos', espacios_nombres)
                iva = nodo_impuestos.get('TotalImpuestosTrasladados')
                retencion = nodo_impuestos.get('TotalImpuestosRetenidos' )

                # Crea una nueva fila (Serie de pandas) con todos los datos extraídos usando un diccionario.
                nueva_fila = pd.Series( {
                    "FACTURA": factura,
                    "FECHA": complemento.get("FechaTimbrado"),
                    "RFC": emisor.get("Rfc"),
                    "NOMBRE": emisor.get("Nombre"),
                    'SUBTOTAL': raiz.get('SubTotal'),
                    'IVA': iva,
                    'RETENCIÓN': retencion,
                    "TOTAL": raiz.get("Total"),
                    "UUID": valor_uuid,

                    # Añade el nombre del ZIP de origen para identificar de dónde proviene la factura.
                    "ORIGEN_CARPETA": nombre_zip_origen, # Añade el nombre del ZIP de origen.

                    # Añade la referencia de la factura como una combinación del RFC, los primeros 8 caracteres del UUID y el nombre del emisor.
                    #"REFERENCIA": str( emisor.get("Rfc") ) + "*" + valor_uuid[0:8] + "*" + str ( emisor.get("Nombre") ) 
                    "REFERENCIA": str( emisor.get("Rfc") ) + "*" + valor_uuid[0:8] + "*" + str ( nombre_zip_origen ).partition(".zip")[0]
                } )
                # Añade la nueva fila al DataFrame.
                datos_xml_df = pd.concat([datos_xml_df, nueva_fila.to_frame().T], ignore_index=True)
        
        except ET.ParseError:
            # Si un XML está mal formado, lo ignora y avisa en la consola.
            print(f"Aviso: No se pudo leer {os.path.basename(ruta_xml)}. El archivo puede estar corrupto.")
            continue

    # Si se extrajeron datos, se limpian y ordenan.
    if not datos_xml_df.empty:
        # Convierte la columna TOTAL a número.
        datos_xml_df['TOTAL'] = pd.to_numeric(datos_xml_df['TOTAL'], errors='coerce').fillna(0)
        # Convierte la columna SUBTOTAL a número.
        datos_xml_df['SUBTOTAL'] = pd.to_numeric(datos_xml_df['SUBTOTAL'] )
        # Convierte la columna IVA a número.
        datos_xml_df['IVA'] = pd.to_numeric(datos_xml_df['IVA'] )
        # Convierte la columna RETENCIÓN a número.
        datos_xml_df['RETENCIÓN'] = pd.to_numeric(datos_xml_df['RETENCIÓN'] )
                                                                                                                   

        # Convierte la FECHA a formato de fecha para poder ordenarla.
        datos_xml_df['FECHA'] = pd.to_datetime(datos_xml_df['FECHA'])
        # Ordena los registros por fecha.
        datos_xml_df = datos_xml_df.sort_values(by='FECHA').reset_index(drop=True)
        # Vuelve a formatear la fecha a día/mes/año para el dataframe final.
        datos_xml_df['FECHA'] = datos_xml_df['FECHA'].dt.strftime('%d/%m/%Y')

    return datos_xml_df # regresamos el dataframe de la funcion extraer_datos_de_un_zip y culminamos la función.


# --- Creación de la Interfaz Gráfica (GUI) ---
# Este bloque solo se ejecuta si el script es el archivo principal.
if __name__ == "__main__":
    # Crea la ventana principal de la aplicación.
    ventana_principal = tk.Tk()
    ventana_principal.title("Procesador de Facturas XML")
    ventana_principal.geometry("400x200") # Tamaño inicial de la ventana

    # Crea un marco (frame) para organizar los elementos dentro de la ventana.
    marco_principal = tk.Frame(ventana_principal, padx=20, pady=20)
    marco_principal.pack(expand=True, fill=tk.BOTH)

    # Crea una etiqueta de texto con instrucciones para el usuario.
    etiqueta_instrucciones = tk.Label(marco_principal, text="Haz clic en el botón para seleccionar los archivos ZIP que deseas procesar.", wraplength=350)
    etiqueta_instrucciones.pack(pady=(0, 20))

    # Crea el botón que, al ser presionado, llamará a la función 'procesar_archivos_xml'.
    boton_procesar = tk.Button(marco_principal, text="Seleccionar y Procesar Archivos ZIP", command=procesar_archivos_xml, bg="#2E8B57", fg="white", font=("Helvetica", 10, "bold"))
    boton_procesar.pack(pady=10, ipady=10, fill=tk.X)

    # Inicia el bucle principal de la aplicación, que la mantiene visible y esperando acciones del usuario.
    ventana_principal.mainloop()
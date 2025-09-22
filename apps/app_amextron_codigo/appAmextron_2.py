# -*- coding: utf-8 -*-
"""
Procesador de facturas Amex

Este código proporciona una sencilla interfaz gráfica para el script desarrollado como Amextron 
disponible en : 
https://drive.google.com/file/d/1vcxJZntbQDFeMNvzsRgWX8bchwKHQtSV/view?usp=drive_link

Se permite: 
1. Proporcionar un nombre y un mes. 
2. Seleccionar una carpeta comprimida en formato .zip que contiene las facturas en formato pdf y xml
3. Se procesan ambos formatos siguiendo las instrucciones del script Amextron: extrae datos de los archivos XML, renombra los pdf  según su UUID (en caso de 
encontrar una coincidencia con esa frase) y se genera un reporte en Excel. 
4. Guarda el archivo Excel generado, los archivos XML originales y los pdf renombrados. Ignora cualquier otro archivo contenido en la carpeta original comprimida
en formato .zip

Posee cualquier limitación derivada del script original Amextron. Las limitaciones de la interfaz no incluye algunos casos de uso extravagantes y el manejo
de excepciones y errores es generico.


Pagina para aprender sobre interfaces graficas y el uso de tkinter en general: https://realpython.com/python-gui-tkinter/

Para crear .exe instalar PyInstaller y despues usar el siguiente comando : python -m PyInstaller --windowed --onefile appAmextron_2.py
Usamos windowed para que no se abra una ventana de consola al ejecutar el .exe y onefile para que se genere un solo archivo .exe
El .exe generado se encuentra en la carpeta dist
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import pdfplumber
import shutil
import threading
from pathlib import Path

class App:
    """
    La aplicacion principal que maneja y administra la GUI
    """

    #self es un argumento que debe pasarse a la funcion
    def __init__(self, ventanaOrigen):
        self.root = ventanaOrigen
        self.root.title("Procesador de facturas Amex")
        self.root.geometry("600x550")
        self.root.configure(bg="#f0f0f0")

        self.directorioZip = None

        # frame principal
        main_frame = tk.Frame(self.root, padx=15, pady=15, bg="#f0f0f0")
        main_frame.pack(expand=True, fill=tk.BOTH)

        # Nombre y mes (datos) de la factura
        fields_frame = tk.LabelFrame(main_frame, text="Datos de la factura", padx=10, pady=10, bg="#f0f0f0", font=("Helvetica", 12))
        fields_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(fields_frame, text="Nombre:", bg="#f0f0f0", font=("Helvetica", 10)).grid(row=0, column=0, sticky="w", pady=5, padx=5)
        self.personaje_entry = tk.Entry(fields_frame, width=40, font=("Helvetica", 10))
        self.personaje_entry.grid(row=0, column=1, sticky="ew", pady=5, padx=5)

        tk.Label(fields_frame, text="Mes:", bg="#f0f0f0", font=("Helvetica", 10)).grid(row=1, column=0, sticky="w", pady=5, padx=5)
        self.mes_entry = tk.Entry(fields_frame, width=40, font=("Helvetica", 10))
        self.mes_entry.grid(row=1, column=1, sticky="ew", pady=5, padx=5)
        
        fields_frame.grid_columnconfigure(1, weight=1)

        # seleccion de archivos
        file_frame = tk.Frame(main_frame, bg="#f0f0f0")
        file_frame.pack(fill=tk.X, pady=10)

        self.select_button = tk.Button(file_frame, text="1. Seleccionar Archivo .ZIP", command=self.select_zip, font=("Helvetica", 10, "bold"), bg="#007bff", fg="white", relief=tk.FLAT, padx=10)
        self.select_button.pack(side=tk.LEFT)

        self.file_label = tk.Label(file_frame, text="Ningún archivo seleccionado", bg="#e9ecef", fg="#495057", padx=10, anchor="w")
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))

        # boton de procesamiento
        self.process_button = tk.Button(main_frame, text="2. Iniciar Procesamiento", command=self.start_processing_thread, font=("Helvetica", 12, "bold"), bg="#28a745", fg="white", relief=tk.FLAT, state=tk.DISABLED, pady=5)
        self.process_button.pack(fill=tk.X, pady=10)

        # frame del registro de actividad (log frame)
        log_frame = tk.LabelFrame(main_frame, text="Registro de Actividad", padx=10, pady=10, bg="#f0f0f0", font=("Helvetica", 12))
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.status_log = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state='disabled', font=("Courier New", 9))
        self.status_log.pack(fill=tk.BOTH, expand=True)

    def update_status(self, message):
        """agrega los mensajes al registro de actividad en la interfaz grafica GUI (log )"""
        self.status_log.config(state='normal')
        self.status_log.insert(tk.END, message + "\n")
        self.status_log.config(state='disabled')
        self.status_log.see(tk.END)
        self.root.update_idletasks()

    def select_zip(self):
        """Abre un cuadro de dialogo para seleccionar un zip"""
        path = filedialog.askopenfilename(
            title="Seleccione el archivo ZIP con las facturas",
            filetypes=[("Zip files", "*.zip")]
        )
        if path:
            self.directorioZip = path
            self.file_label.config(text=os.path.basename(path))
            self.update_status(f"Archivo seleccionado: {path}")
            self.process_button.config(state=tk.NORMAL) # habilita el boton de procesado
        else:
            self.file_label.config(text="Ningún archivo seleccionado")
            self.process_button.config(state=tk.DISABLED)

    def start_processing_thread(self):
        """
        valida inputs y empieza el procesamiento en una ejecucion diferente al del codigo amextron
        para separarlo de la interfaz 
        """
        personaje = self.personaje_entry.get().strip().upper()
        mes = self.mes_entry.get().strip().upper()

        if not personaje or not mes:
            messagebox.showerror("Error de Entrada", "Los campos '´Persona' y 'Mes' no pueden estar vacíos.")
            return
        
        if not self.directorioZip:
            messagebox.showerror("Error de Entrada", "Por favor, seleccione un archivo ZIP para procesar.")
            return

        # desabilita los botones de la ventana para evitar que el usuario añada mas carpetas durante el procesamiento
        self.process_button.config(state=tk.DISABLED)
        self.select_button.config(state=tk.DISABLED)

        # corre la logica principal del programa (el de colab Amextron) en un thread 
        processing_thread = threading.Thread(
            target=self.process_files,
            args=(personaje, mes, self.directorioZip)
        )
        processing_thread.start()

    def process_files(self, personaje, mes, directorioZip):
        """
        La logica del programa Amextron corre aqui. Es el codigo en Colab basicamente
        """
        directorioTemporal = None
        try:
            # crea un directorio temporal para procesar
            directorioTemporal = Path(f"./temp_{personaje}_{mes}")
            if directorioTemporal.exists():
                shutil.rmtree(directorioTemporal) # Clean up old temp dir if it exists
            directorioTemporal.mkdir()
            self.update_status(f"Directorio temporal creado en: {directorioTemporal.resolve()}")

            # extrae el archivo zip
            extract_path = directorioTemporal / personaje
            self.update_status(f"Extrayendo '{Path(directorioZip).name}' a {extract_path}...")
            with zipfile.ZipFile(directorioZip, 'r') as zip_ref:
                zip_ref.extractall(extract_path)
            self.update_status("✅ Extracción completada.")

            # inicializa el dataframe
            df = pd.DataFrame(columns=['UUID', "RFC del emisor", 'Subtotal', 'IVA', 'Total', "Tipo", "Nombre"])
            
            # Process XML files
            self.update_status("Procesando archivos XML...")
            nombrederaices = {
                'cfdi': 'http://www.sat.gob.mx/cfd/4',
                'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital'
            }
            
            all_files = list(os.walk(extract_path))
            for dirpath, _, filenames in all_files:
                for item in filenames:
                    if item.lower().endswith(".xml"):
                        xml_path = Path(dirpath) / item
                        try:
                            tree = ET.parse(xml_path)
                            root = tree.getroot()
                            
                            complemento = root.find('.//tfd:TimbreFiscalDigital', nombrederaices)
                            if complemento is None:
                                self.update_status(f"  - Advertencia: No se encontró TimbreFiscalDigital en {item}")
                                continue
                            
                            uuid = complemento.get("UUID")
                            emisor = root.find('cfdi:Emisor', nombrederaices)
                            
                            impuestos = root.find('cfdi:Impuestos', nombrederaices)
                            iva = "0"
                            if impuestos is not None and impuestos.get('TotalImpuestosTrasladados'):
                                iva = impuestos.get('TotalImpuestosTrasladados')
                            
                            new_row = pd.Series({
                                "UUID": uuid,
                                "Total": root.get("Total"),
                                "RFC del emisor": emisor.get("Nombre") if emisor is not None else "N/A",
                                'Subtotal': root.get('SubTotal'),
                                'IVA': iva,
                                "Tipo": root.get('TipoDeComprobante'),
                                "Nombre": item
                            })
                            df = pd.concat([df, new_row.to_frame().T], ignore_index=True)
                            
                            # Rename XML file
                            os.rename(xml_path, xml_path.parent / f"{uuid}.xml")

                        except ET.ParseError:
                            self.update_status(f"  - Error: No se pudo parsear el XML {item}. Omitiendo.")
                        except Exception as e:
                            self.update_status(f"  - Error procesando {item}: {e}")

            self.update_status(f"✅ {len(df)} facturas XML procesadas.")

            # convierte columnas
            for col in ['Subtotal', 'IVA', 'Total']:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            # procesa y renombra los archivos pdf
            self.update_status("Renombrando archivos PDF según UUID...")
            for dirpath, _, filenames in all_files:
                for archivo in filenames:
                    if archivo.lower().endswith('.pdf'):
                        pdf_path = Path(dirpath) / archivo
                        renamed = False
                        try:
                            with pdfplumber.open(pdf_path) as pdf:
                                for uuid_to_find in df['UUID'].tolist():
                                    for page in pdf.pages:
                                        text = page.extract_text()
                                        if text and uuid_to_find in text:
                                            new_name = pdf_path.parent / f"{uuid_to_find}.pdf"
                                            os.rename(pdf_path, new_name)
                                            self.update_status(f"  - Renombrado: {archivo} -> {new_name.name}")
                                            renamed = True
                                            break
                                    if renamed:
                                        break
                        except Exception as e:
                            self.update_status(f"  - Error procesando PDF {archivo}: {e}")
            self.update_status("✅ Procesamiento de PDF completado.")

            # Crea el archivo Excel que se mostrara
            nombreExcel = f"{mes}_{personaje}.xlsx"
            directorioExcel = directorioTemporal / nombreExcel
            self.update_status(f"Creando archivo Excel: {nombreExcel}...")
            df.to_excel(directorioExcel, index=False)
            self.update_status("✅ Archivo Excel creado.")

            # Le pregunta al usuario donde guaradara el archivo final
            output_zip_name = f"{personaje}_{mes}.zip"
            save_path = filedialog.asksaveasfilename(
                defaultextension=".zip",
                initialfile=output_zip_name,
                filetypes=[("Zip files", "*.zip")],
                title="Guardar archivo Zip procesado"
            )

            if not save_path:
                self.update_status("Operación cancelada por el usuario.")
                return

            # Crea el archivo zip final que se almacenara
            self.update_status(f"Creando archivo ZIP final en: {save_path}")
            with zipfile.ZipFile(save_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                # Add processed files from the extraction folder
                for root_dir, _, files_in_dir in os.walk(extract_path):
                    for file in files_in_dir:
                        file_path = Path(root_dir) / file
                        arcname = file_path.relative_to(extract_path)
                        zipf.write(file_path, arcname=arcname)
                # Add the excel file to the root of the zip
                zipf.write(directorioExcel, arcname=nombreExcel)
            
            self.update_status("--------------------------------------------------")
            self.update_status(" PROCESO COMPLETADO CON ÉXITO ")
            self.update_status(f"Archivo final guardado en: {save_path}")
            messagebox.showinfo("Éxito", f"Proceso completado.\nEl archivo ha sido guardado en:\n{save_path}")

        #manejo de expeciones
        except Exception as e:
            mensajeError = f"ERROR INESPERADO: {e}"
            self.update_status(mensajeError)
            messagebox.showerror( f"Ocurrió un error durante el proceso:\n{e}")
        finally:
            # limpia el directorio temporal
            if directorioTemporal and directorioTemporal.exists():
                shutil.rmtree(directorioTemporal)
                self.update_status(f"Directorio temporal '{directorioTemporal.name}' eliminado.")
            # Habilita de nuevo los botones de la ventana
            self.process_button.config(state=tk.NORMAL)
            self.select_button.config(state=tk.NORMAL)


if __name__ == "__main__":
    #funcion main, la ejecucion inicia en este punto por ser la funcion  main
    #los siguientes comandos son usuales en el manejo de la biblioteca tkinter
    root = tk.Tk()
    app = App(root)
    root.mainloop()

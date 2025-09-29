import PyPDF2
import re
import os
import zipfile
import io

# Función para extraer el titular de la cuenta de una página
def extract_titular(text : str) -> str :
    match = re.search(r'UXILIAR DEL  Titular de la cuenta:\s*([A-Z\s]+)', text)
    if match:
        return match.group(1).strip()
    match = re.search(r'SA DE CV  Titular de la cuenta:\s*([A-Z\s]+)', text)
    if match:
        return match.group(1).strip()
    match = re.search(r'\nCREDITO  Titular de la cuenta:\s*([A-Z\s]+)', text)
    if match:
        return match.group(1).strip()
    match = re.search(r'\nCV  Titular de la cuenta:\s*([A-Z\s]+)', text)
    if match:
        return match.group(1).strip()
    match = re.search(r'DEL Titular de la cuenta:\s*([A-Z\s]+)', text)
    if match:
        return match.group(1).strip()


def separar_paginas ( pdf_path : str, output_dir : str, fecha : str ) -> list[io.BytesIO]:
    
    # almacenamos los pdf en una lista
    pdf_en_memoria_buffer = []

    #creamos un elemento en memoria (buffer)
    output_buffer = io.BytesIO()
    
    # Leer el PDF y separar por paginas
    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)

        for i, page in enumerate(reader.pages):
            text = page.extract_text()
            titular = extract_titular(text)
            if not titular:
                titular = f"Desconocido_{i+1}"
            else:
                titular = '_'.join(titular.split())

            output_filename = f"{fecha}_{titular}_{i+1}.pdf"
            output_path = os.path.join(output_dir, output_filename)

            # Crear nuevo PDF con la página extraída
            writer = PyPDF2.PdfWriter()
            writer.add_page(page)

            writer.write( output_buffer )
            output_buffer.seek(0) 

            pdf_en_memoria_buffer.append( (output_buffer, output_filename) )

    

    return pdf_en_memoria_buffer
        

def funcion_principal(pdf_path: str, output_dir: str, fecha: str, numero_archivo: str) -> io.BytesIO:
    lista_pdf_buffer = separar_paginas(pdf_path=pdf_path, output_dir=output_dir, fecha=fecha)

    # Nombre del archivo ZIP
    zip_filename = f"PROVEEDORES_{fecha}_{numero_archivo}.zip"

    # Creamos el ZIP en memoria
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for elemento_buffer, nombre in lista_pdf_buffer:
            elemento_buffer.seek(0)
            zipf.writestr(nombre, elemento_buffer.read())  
    zip_buffer.seek(0)  # Regresar el puntero al inicio

    return zip_buffer

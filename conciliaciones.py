import fitz  # PyMuPDF se importa como 'fitz'
import re
import pandas as pd

"""
    Extrae los montos de los cargos de un estado de cuenta en PDF utilizando PyMuPDF (fitz).
    Localiza la columna 'CARGOS' por coordenadas y extrae los números correspondientes.

    input:
        ruta_pdf: La ruta del archivo PDF del estado de cuenta.

    output:
        Una lista de números de punto flotante que representan los montos de los cargos.
"""
def extraer_cargos_con_pymupdf(ruta_pdf: str) -> list[float]:
    todos_los_cargos = []
    
    # Abrimos el documento con fitz.open()
    with fitz.open(ruta_pdf) as doc:
        # Iteramos a través de cada página del documento
        for i, page in enumerate(doc):
            print(f"Procesando Página {i + 1}...")

            # page.get_text("words") devuelve una lista de tuplas.
            # Cada tupla contiene (x0, y0, x1, y1, "texto", block_no, line_no, word_no)
            palabras_en_pagina = page.get_text("words")
            
            # Buscamos el encabezado "CARGOS" para usarlo como referencia a la hora de ubicar la columna
            #de los cargos
            encabezado_info = [palabra for palabra in palabras_en_pagina if palabra[4].upper() == 'CARGOS']
            
            # Si encontramos el encabezado en la página
            if encabezado_info:

                if ( i== 0 ) :
                    encabezado = encabezado_info[-1]
                else:
                    encabezado = encabezado_info[0]
                
                # Extraemos las coordenadas del encabezado de la tupla
                columna_x0 = encabezado[0]
                columna_x1 = encabezado[2]
                borde_inferior_encabezado = encabezado[3]
                
                # Iteramos sobre todas las palabras de la página nuevamente
                for palabra_tupla in palabras_en_pagina:
                    # Extraemos la información de la tupla de la palabra
                    px0, py0, px1, py1, texto_palabra = palabra_tupla[:5]
                    
                    # Verificamos tres condiciones 
                    centro_x_palabra = (px0 + px1) / 2
                    esta_debajo_encabezado = py0 > borde_inferior_encabezado
                    esta_en_columna = columna_x0 <= centro_x_palabra <= columna_x1
                    
                    if esta_debajo_encabezado and esta_en_columna:
                        # Limpiamos el texto para convertirlo a número
                        texto_limpio = texto_palabra.replace(',', '')
                        
                        if re.match(r'^\d+(\.\d{1,2})?$', texto_limpio):
                            try:
                                valor_cargo = float(texto_limpio)
                                todos_los_cargos.append(valor_cargo)
                            except ValueError:
                                continue
                                
    print(f"\nExtracción completada. Se encontraron {len(todos_los_cargos)} cargos.")
    return todos_los_cargos
"""
    regresa dos listas, las columnas de cargos y creditos que provienen del sistema BAN

    input:
        ruta_pdf: La ruta del archivo PDF del documento del Ban

    output:
        una tupla con dos listas. Una lista son los abonos y otro los cargos
"""

def procesar_baan(ruta_pdf: str) -> tuple[list]:

    documento = pd.read_excel(ruta_pdf, skiprows= 15 )


    abonos = documento.iloc[:, 10].dropna()
    cargos = documento.iloc[:,11].dropna()

    abonos = abonos.astype(str)
    cargos = cargos.astype(str)

    abonos = abonos[ abonos.str.strip() != "" ]
    cargos= cargos [ cargos.str.strip() != "" ]

    cargos = cargos.to_list()
    abonos = abonos.to_list()

    for indice in range(len(cargos)):
        cargos[indice] =cargos[indice].strip()

    for indice in range(len(abonos)):
        abonos[indice] =abonos[indice].strip()
    
    for elemento in cargos:
        if elemento.endswith("-"):  # revisamos si termina con un menos
            elemento = -float(elemento[:-1])  # 
        else:
            elemento = float(elemento)

    for elemento in abonos:
        if elemento.endswith("-"):  # revisamos si termina con un menos
            elemento = -float(elemento[:-1])  # take all except the last char, convert to float, and negate
        else:
            elemento = float(elemento)


    
    cargos_resultado = [ -float(elemento[:-1]) if  elemento.endswith("-") else float(elemento) for elemento in cargos]
    abonos_resultado = [ -float(elemento[:-1]) if  elemento.endswith("-") else float(elemento) for elemento in abonos]

    return abonos_resultado,cargos_resultado







####abonos

def extraer_abonos_con_pymupdf(ruta_pdf: str, epsilon:float ) -> list[float]:
    """
    Extrae los montos de los abonos de un estado de cuenta en PDF utilizando PyMuPDF (fitz).
    Localiza la columna 'ABONOS' por coordenadas y extrae los números correspondientes.

    Args:
        ruta_pdf: La ruta del archivo PDF del estado de cuenta.

    Returns:
        Una lista de números de punto flotante que representan los montos de los abonos.
    """
    todos_los_abonos = []
    
    # Abrimos el documento con fitz.open()
    with fitz.open(ruta_pdf) as doc:
        # Iteramos a través de cada página del documento
        for i, page in enumerate(doc):
            print(f"Procesando Página {i + 1}...")

            # Obtenemos todas las palabras de la página con sus coordenadas.
            palabras_en_pagina = page.get_text("words")
            
            # ¡CAMBIO CLAVE! Ahora buscamos "ABONOS" como encabezado.
            encabezado_info = [p for p in palabras_en_pagina if p[4].upper() == 'ABONOS']
            
            # Si encontramos el encabezado "ABONOS" en la página
            if encabezado_info:

                if ( i== 0 ) :
                    encabezado = encabezado_info[-1]
                else:
                    encabezado = encabezado_info[0]
                
                
                # Extraemos las coordenadas del encabezado
                columna_x0 = encabezado[0]
                columna_x1 = encabezado[2]
                borde_inferior_encabezado = encabezado[3]
                
                # Iteramos sobre todas las palabras para encontrar las que están en la columna de abonos
                for palabra_tupla in palabras_en_pagina:
                    px0, py0, px1, py1, texto_palabra = palabra_tupla[:5]
                    
                    centro_x_palabra = (px0 + px1) / 2
                    esta_debajo_encabezado = py0 > borde_inferior_encabezado
                    esta_en_columna = columna_x0- epsilon <= centro_x_palabra <= columna_x1+ epsilon
                    
                    if esta_debajo_encabezado and esta_en_columna:
                        texto_limpio = texto_palabra.replace(',', '')
                        
                        if re.match(r'^\d+(\.\d{1,2})?$', texto_limpio):
                            try:
                                valor_abono = float(texto_limpio)
                                todos_los_abonos.append(valor_abono)
                            except ValueError:
                                continue
                                
    print(f"\nExtracción completada. Se encontraron {len(todos_los_abonos)} abonos.")
    return todos_los_abonos




# Ejecutar script

# Ruta al archivo PDF
#ruta_excel_banco = r"C:\Users\SALCIDOA\Downloads\SERMEX 3108.pdf"

#ruta_excel_baan = r"C:\Users\SALCIDOA\Downloads\excel.xlsx"

def ejecucion_programa(ruta_excel_banco: str, ruta_excel_baan: str) -> list:

    #obtenemos la lista de los abonos y conta del ban
    lista_abonos_conta, lista_cargos_conta = procesar_baan(ruta_excel_baan)

    # Llama a la funcion que usa PyMuPDF
    cargos_bancarios_extraidos = extraer_cargos_con_pymupdf(ruta_excel_banco)

    # Llamamos a la función para extraer los abonos
    delta = 0.001 #los abonos llevan una delta por lo problematico de su ubicacion
    abonos_bancarios_extraidos = extraer_abonos_con_pymupdf(ruta_excel_banco,   delta)


    #creamos una copia de las listas de cargos y abonos extraidos para trabajar con ellas y  
    #no modificar las listas originales
    lista_cargos_bancarios = cargos_bancarios_extraidos.copy()
    
    lista_abonos_bancarios = abonos_bancarios_extraidos.copy()



    ###cargos
    #ahora revisamos que cargos estan en el banco pero no en los cargos de contabilidad

    for elemento in lista_cargos_conta:
        if elemento in lista_cargos_bancarios:
            lista_cargos_bancarios.remove( elemento )

    #revisamos que elementos estan en los cargos de contabilidad pero en el banco
    for elemento in cargos_bancarios_extraidos:
        if elemento in lista_cargos_conta:
            lista_cargos_conta.remove( elemento )


 
    #####abonos
    #revisamos que abonos estan en el banco pero no en la contabilidad

    for elemento in lista_abonos_conta:
        if elemento in lista_abonos_bancarios:
            lista_abonos_bancarios.remove(elemento)

    #revisamos los abonos que estan en contabilidad pero no en el banco

    for elemento in abonos_bancarios_extraidos:
        if elemento in lista_abonos_conta:
            lista_abonos_conta.remove(elemento)

    return lista_cargos_bancarios,lista_cargos_conta, lista_abonos_bancarios, lista_abonos_conta



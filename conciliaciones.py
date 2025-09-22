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
ruta_pdf= r"C:\Users\SALCIDOA\Downloads\excel.xlsx"

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




    patron = r"\d+-"

    return abonos,cargos

# Ejecutar script

# 1. Ruta al archivo PDF
archivo_pdf = r"C:\Users\SALCIDOA\Downloads\SERMEX 3108.pdf"

ruta_pdf = r"C:\Users\SALCIDOA\Downloads\SERMEX 3108.pdf"

# 2. Llama a la nueva función que usa PyMuPDF
cargos_bancarios_extraidos = extraer_cargos_con_pymupdf(archivo_pdf)


print(len(cargos_bancarios_extraidos))

# 4. Calcula la suma total para verificar
total_cargos = sum(cargos_bancarios_extraidos)
print(f"\nSuma de todos los cargos extraídos: {total_cargos:,.2f}")


###procesamos los pagos que unicamente estan de uno u otro lado

lista_contable = pd.read_csv( r"C:\Users\SALCIDOA\Downloads\lista_cargos_ban.csv", header= None)

lista_cargos_conta = lista_contable.iloc[:,0].to_list()

lista_bancarios = cargos_bancarios_extraidos


for elemento in lista_cargos_conta:
    if elemento in lista_bancarios:
        lista_bancarios.remove( elemento )

print("Cargos del banco no reconocidos: ", lista_bancarios)

#reiniciamos la lista de cargos bancarios 
lista_bancarios = cargos_bancarios_extraidos

for elemento in cargos_bancarios_extraidos:
    if elemento in lista_cargos_conta:
        lista_cargos_conta.remove( elemento )

print("Nuestros cargos no reconocidos: ", lista_cargos_conta )





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

# Ejecucion

# 1. Ruta al archivo PDF
archivo_pdf = r"c:\Users\SALCIDOA\Downloads\conciliacion\SERMEX 3108.pdf"

# 2. Llama a la función para extraer los abonos
delta = 0.001
abonos_extraidos = extraer_abonos_con_pymupdf(archivo_pdf,   delta)


len(abonos_extraidos)
# 4. Calcula la suma total para verificar con el documento
total_abonos = sum(abonos_extraidos)
# [cite_start]Compara este total con el "TOTAL IMPORTE ABONOS" de 67,282,311.56 que aparece en la página 52 del PDF [cite: 826]
print(f"\nSuma de todos los abonos extraídos: {total_abonos:,.2f}")


lista_abonos_ban = pd.read_csv(r"C:\Users\SALCIDOA\Downloads\abonos_ban.csv", header=None)
lista_abonos_conta = lista_abonos_ban.iloc[:,0].to_list()




## revisamos los que estan de mas en una y otra lista

lista_abonos_banco = abonos_extraidos.copy()

for elemento in lista_abonos_conta:
    if elemento in lista_abonos_banco:
        lista_abonos_banco.remove(elemento)

print("Creditos del banco no correspondidos : " , lista_abonos_banco)
print("Total: " , len(lista_abonos_banco))

#reiniciamos la lista
lista_abonos_banco = abonos_extraidos.copy()

for elemento in lista_abonos_banco:
    if elemento in lista_abonos_conta:
        lista_abonos_conta.remove(elemento)

print("Nuestros creditos no correspondidos: " , lista_abonos_conta)
print("Total: " , len(lista_abonos_conta))

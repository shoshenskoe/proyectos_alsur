import fitz  # PyMuPDF
import re

def extraer_cargos_con_pymupdf(ruta_pdf: str) -> list[tuple[str, float]]:
    cargos_con_fecha = []
    
    with fitz.open(ruta_pdf) as doc:
        for i, page in enumerate(doc):
            palabras_en_pagina = page.get_text("words")
            
            # Localizar encabezados
            encabezado_cargos = [w for w in palabras_en_pagina if w[4].upper() == 'CARGOS']
            encabezado_fecha  = [w for w in palabras_en_pagina if w[4].upper() == 'OPER' ]
            
            if encabezado_cargos and encabezado_fecha:
                # Tomamos las coordenadas de las columnas

                if (i == 0):
                    col_x0_cargos, col_x1_cargos = encabezado_cargos[-1][0], encabezado_cargos[-1][2]
                    col_x0_fecha, col_x1_fecha   = encabezado_fecha[-1][0], encabezado_fecha[0][2]
                    borde_inferior_encabezado    = encabezado_cargos[-1][3]
                else:
                    col_x0_cargos, col_x1_cargos = encabezado_cargos[0][0], encabezado_cargos[0][2]

                    col_x0_fecha, col_x1_fecha   = encabezado_fecha[0][0], encabezado_fecha[0][2]
                    borde_inferior_encabezado    = encabezado_cargos[0][3]
                
                # Recorremos todas las palabras
                for px0, py0, px1, py1, texto, *_ in palabras_en_pagina:
                    centro_x = (px0 + px1) / 2
                    if py0 > borde_inferior_encabezado and col_x0_cargos <= centro_x <= col_x1_cargos:
                        texto_limpio = texto.replace(',', '')
                        if re.match(r'^\d+(\.\d{1,2})?$', texto_limpio):
                            try:
                                cargo = float(texto_limpio)
                                # Buscar la fecha en la misma línea (misma coordenada Y aproximada)
                                fecha_candidatos = [
                                    t for (fx0, fy0, fx1, fy1, t, *_) in palabras_en_pagina
                                    if  (py1 <= (fy0+fy1)/2 <=py0)  and ( col_x0_fecha-0.001 <= (fx0+fx1)/2 <= (col_x1_fecha + 0.1) )
                                ]

                                fecha = fecha_candidatos[0] if fecha_candidatos else None
                                cargos_con_fecha.append((fecha, cargo))
                            except ValueError:
                                continue
    return cargos_con_fecha

# Ejemplo de uso
ruta_pdf = r"C:\Users\SALCIDOA\Downloads\SERMEX 3108.pdf"

cargos_con_fechas = extraer_cargos_con_pymupdf(ruta_pdf)

if cargos_con_fechas:
    print("Se extrajeron los siguientes cargos y fechas:")
    for fecha, cargo in cargos_con_fechas:
        print(f"Fecha: {fecha}, Cargo: ${cargo:,.2f}")
else:
    print("No se encontraron cargos o fechas en el documento.")

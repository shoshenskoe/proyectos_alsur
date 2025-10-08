
import pandas as pd
import numpy as np
from io import BytesIO

# funciones auxiliares


def obtener_dfsucio(excel_path):
    dfsucio = pd.read_excel(excel_path, skiprows=list(range(16)))
    return dfsucio

def obtener_df(df_sucio):
    df = df_sucio[(df_sucio["IVA"] != 0) & (df_sucio["Importe"] != 0) & (df_sucio["Cargo"] != 0)].copy()
    df = df.reset_index(drop=True)
    df = df[["Nombre\nEmpleado", "Cargo", "Importe", "IVA"]]
    df = df.rename(columns={'Nombre\nEmpleado': "Nombre Empleado"})
    df = df.groupby(['Nombre Empleado']).sum().reset_index()
    total_row = pd.Series({
        "Nombre Empleado": "Total",
        "Importe": df.Importe.sum(),
        "Cargo": df.Cargo.sum(),
        "IVA": df.IVA.sum()
    })
    df = pd.concat([df, total_row.to_frame().T], ignore_index=True)
    return df

def crear_tabla_con_cc_vacia(df):
    enlace = "https://docs.google.com/spreadsheets/d/1Iy68cztYlqI6fLjE8l4s2D32BQLzUZG9/export?format=xlsx"
    df2 = pd.read_excel(enlace)
    df = pd.merge(df, df2, on='Nombre Empleado', how='left')
    df = df.reindex(columns=["CC", "Nombre Empleado", "Cargo", "Importe", "IVA"])
    return df

def nombres_faltantes(df_sin_cc):
    datos_faltantes = df_sin_cc[df_sin_cc["CC"].isnull()]
    lista_nombres = datos_faltantes["Nombre Empleado"].tolist()
    return lista_nombres

def hacer_verficiacion_v2(df, dic_nombre_cc, df_empleados_cc):
    df_actualizado = df.copy()
    for nombre, centro in dic_nombre_cc.items():
        if nombre in df_actualizado['Nombre Empleado'].values:
            df_actualizado.loc[df_actualizado['Nombre Empleado'] == nombre, 'CC'] = centro.upper()
            if nombre not in df_empleados_cc['Nombre Empleado'].values:
                nuevo_empleado = pd.DataFrame([{'CC': centro.upper(), 'Nombre Empleado': nombre}])
                df_empleados_cc = pd.concat([df_empleados_cc, nuevo_empleado], ignore_index=True)
    return df_actualizado, df_empleados_cc


def obterner_df_no_camiones(df3):
    camiones = ['TRAILER 001', 'TRAILER 002', 'TRAILER 004', 'TRAILER 05']
    df3 = df3[~df3['Nombre Empleado'].isin(camiones)]
    return df3.reset_index(drop=True)

def crear_segunda_tabla_din(df):
    df_copy = df.copy()
    df_copy[["Cargo", "Importe", "IVA"]] = df_copy[["Cargo", "Importe", "IVA"]].astype("float64")
    df_copy["CC"] = df_copy["CC"].str.upper()
    df4 = df_copy.groupby("CC")[["Cargo", "Importe", "IVA"]].sum().reset_index()
    total_row = pd.Series({
        "CC": "Total",
        "Cargo": df4.Cargo.sum(),
        "Importe": df4.Importe.sum(),
        "IVA": df4.IVA.sum()
    })
    df4 = pd.concat([df4, total_row.to_frame().T], ignore_index=True)
    df4[["Cargo", "IVA", "Importe"]] = df4[["Cargo", "IVA", "Importe"]].astype('float').round(0)
    return df4

def enlazar_con_utilitario(df_parapoliza, utilitario):
    df_parapoliza['CC'] = df_parapoliza['CC'].str.strip()
    df_parapoliza = pd.merge(df_parapoliza, utilitario, on=["CC"], how="left")
    df_parapoliza['CC'] = df_parapoliza['CC'].replace("Total", "E913")
    return df_parapoliza

def obtener_faltantes_utilitario(df_parapoliza):
    datos_faltantes = df_parapoliza[df_parapoliza["UTILITARIO"].isnull()]
    return datos_faltantes["CC"].tolist() if not datos_faltantes.empty else []

def completar_utilitario(df_parapoliza, utilitario, diccionario):
    df_actualizado = df_parapoliza.copy()
    for cc, cod_util in diccionario.items():
        if cc in df_actualizado['CC'].values:
            df_actualizado.loc[df_actualizado['CC'] == cc, "UTILITARIO"] = cod_util.upper()
            if cc not in utilitario['CC'].values:
                nuevo_utilitario = pd.DataFrame([{'CC': cc, 'UTILITARIO': cod_util.upper()}])
                utilitario = pd.concat([utilitario, nuevo_utilitario], ignore_index=True)
    return df_actualizado, utilitario

def hacer_poliza_final(df_parapoliza, Referencia: str):
    CF = '64911801CF01'
    GA = '64911801GA01'
    GV = '64911801GV01'
    IVA_por_acreditable = '140104010002'
    Prov_consumo_Sivale = '240112900007'

    df_poliza = df_parapoliza[['CC', 'UTILITARIO', 'Importe', 'IVA', 'Cargo']].copy()
    
    if not df_poliza.empty:
      df_poliza.loc[df_poliza.index[-1], 'Importe'] = df_poliza.loc[df_poliza.index[-1], 'Cargo']

      iva_row = pd.Series({"CC": "", "UTILITARIO": np.nan, 'Importe': df_poliza.loc[df_poliza.index[-1], 'IVA']})
      df_poliza = pd.concat([df_poliza, iva_row.to_frame().T], ignore_index=True)
    
    df_poliza.insert(0, 'plantilla', 'SVA')
    df_poliza.insert(1, 'cons', range(1, len(df_poliza) + 1))
    df_poliza.insert(2, 'comp', '100')
    
    cta = []
    for i in range(len(df_poliza)):
        CC = df_poliza.loc[i, 'CC']
        if pd.notna(CC) and not CC.startswith('E') and CC != "D981" and CC != "":
            cta.append(CF)
        elif pd.notna(CC) and CC.startswith('E') and CC != "E913":
            cta.append(GA)
        elif CC == "D981":
            cta.append(GV)
        elif CC == "E913":
            cta.append(Prov_consumo_Sivale)
        else: # CC is "" or NaN
            cta.append(IVA_por_acreditable)
    df_poliza.insert(3, 'cta', cta)
    
    df_poliza.insert(6, 'nada', np.nan)
    df_poliza.insert(7, 'nada1', np.nan)
    
    debe_haber = [1 if str(c).startswith('6') or str(c).startswith('1') else 2 for c in df_poliza['cta']]
    df_poliza.insert(8, 'debe/haber', debe_haber)
    
    df_poliza.drop(columns=['IVA', 'Cargo'], inplace=True)
    df_poliza["ref"] = f'Consumo Si Vale {Referencia}'
    df_poliza["CC"] = df_poliza["CC"].replace("", np.nan)
    
    return df_poliza

def elaborar_excel_poliza(dfsucio, df, df4, df_parapoliza):
    archivo_buffer = BytesIO()
    with pd.ExcelWriter(archivo_buffer, engine='xlsxwriter') as writer:
        dfsucio.to_excel(writer, sheet_name="Original")
        df.to_excel(writer, sheet_name="TablaDin1", index=False)
        df4.to_excel(writer, sheet_name="TablaDin2", index=False)
        df_parapoliza.to_excel(writer, sheet_name="tfgld013", index=False, header=True)
    archivo_buffer.seek(0)
    return archivo_buffer

# funcion principal 
def main_gui(path_archivo_excel, Referencia, diccionario_nombre_cc, diccionario_util_cc):
    dfsucio = obtener_dfsucio(path_archivo_excel)
    df_inicial = obtener_df(dfsucio)
    
    tabla_incompleta = crear_tabla_con_cc_vacia(df_inicial)

    path_centro_util = "https://docs.google.com/spreadsheets/d/1Iy68cztYlqI6fLjE8l4s2D32BQLzUZG9/export?format=xlsx"
    centros_costos = pd.read_excel(path_centro_util)

    # El df_empleados_cc es el df con los empleados que tienen cc vacias
    df_completo, df_empleados_cc_actualizado = hacer_verficiacion_v2(
        df=tabla_incompleta, 
        dic_nombre_cc=diccionario_nombre_cc, 
        df_empleados_cc=centros_costos
    )

    df3 = obterner_df_no_camiones(df_empleados_cc_actualizado)
    # guardar_en_drive(df3) # La función original está comentada
    
    df4 = crear_segunda_tabla_din(df_completo)

    path_cc_utilitario = "https://docs.google.com/spreadsheets/d/1gnfLiD1arrr5G7seQi85-f3Cd5n7_miS/export?format=csv&gid=1471990202"
    utilitario = pd.read_csv(path_cc_utilitario)

    df_parapoliza = enlazar_con_utilitario(df4, utilitario)

    df_parapoliza, utilitario_actualizado = completar_utilitario(
        df_parapoliza=df_parapoliza,
        utilitario=utilitario,
        diccionario=diccionario_util_cc
    )
    # guardar_en_drive(utilitario_actualizado) # La función original está comentada

    df_poliza_final = hacer_poliza_final(df_parapoliza, Referencia=Referencia)
    
    archivo_excel_buffer = elaborar_excel_poliza(dfsucio, df_completo, df4, df_poliza_final)
    
    return archivo_excel_buffer


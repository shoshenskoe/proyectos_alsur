
import pandas as pd
import numpy as np
from io import BytesIO


def obtener_dfsucio(excel_path):
    dfsucio = pd.read_excel(    excel_path, skiprows=[0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15]     )
    return dfsucio


def obtener_df(df_sucio):
   
    df= df_sucio[(df_sucio["IVA"]!=0) & (df_sucio["Importe"]!=0) & (df_sucio["Cargo"]!=0)].copy()
    df=df.reset_index()
    df=df.drop('index',axis=1)
    df=df[["Nombre\nEmpleado","Cargo","Importe","IVA"]]
    df=df.rename(columns={'Nombre\nEmpleado':"Nombre Empleado"})

    """Hasta aquí ya tenemos la tabla lista y limpia para seguir con el proceso que marca el manual de ángel."""

    # Proceso que imita la primera  tabla dinamica

    df=df.groupby(['Nombre Empleado']).sum()
    df=df.reset_index()
    #df.drop(columns=['No.Trx', 'Precio Unitario ', 'Litros'], inplace= True )
    #df=df.append({"Nombre Empleado":"Total","Cargo":df.Cargo.sum(),"IVA":df.IVA.sum(), "Importe": df.Importe.sum()}, ignore_index=True)
    S=pd.Series({"Nombre Empleado":"Total", "Importe": df.Importe.sum(),"Cargo":df.Cargo.sum(),"IVA":df.IVA.sum()})
    df=pd.concat([df, S.to_frame().T], ignore_index=True)

    return df


def crear_tabla_con_cc_vacia(df):
    enlace = "https://docs.google.com/spreadsheets/d/1Iy68cztYlqI6fLjE8l4s2D32BQLzUZG9/export?format=xlsx"
    df2 = pd.read_excel(enlace)
    df = pd.merge(df, df2, on='Nombre Empleado', how='left')
    df = df.reindex(["CC", "Nombre Empleado", "Cargo", "Importe", "IVA"], axis=1)
    return df


def nombres_faltantes(df_sin_cc):
    datos_faltantes = df_sin_cc[df_sin_cc["CC"].isnull()]
    return datos_faltantes["Nombre Empleado"].tolist()


def hacer_verficiacion_v2(df, dic_nombre_cc, df_empleados_cc):
    for i in range(len(df)):
        nombre = df.loc[i, "Nombre Empleado"]
        if nombre in dic_nombre_cc:
            centro = dic_nombre_cc[nombre]
            df.loc[i, "CC"] = centro.upper()
            nuevo = pd.Series({"CC": centro, "Nombre Empleado": nombre})
            df_empleados_cc = pd.concat([df_empleados_cc, nuevo.to_frame().T], ignore_index=True)
    return df, df_empleados_cc


def obterner_df_no_camiones(df3):
    camiones = ['TRAILER 001', 'TRAILER 002', 'TRAILER 004', 'TRAILER 05']
    df3 = df3[~df3['Nombre Empleado'].isin(camiones)].reset_index(drop=True)
    return df3


def crear_segunda_tabla_din(df):
    df[["Cargo", "Importe", "IVA"]] = df[["Cargo", "Importe", "IVA"]].astype("float64")
    df["CC"] = df["CC"].str.upper()
    df4 = df.groupby("CC").sum().reset_index()
    total = pd.Series({"CC": "Total", "Cargo": df4.Cargo.sum(), "Importe": df4.Importe.sum(), "IVA": df4.IVA.sum()})
    df4 = pd.concat([df4, total.to_frame().T], ignore_index=True)
    df4 = df4.round(0)
    return df4


def enlazar_con_utilitario(df_parapoliza, utilitario):
    df_parapoliza['CC'] = df_parapoliza['CC'].str.strip()
    df_parapoliza = pd.merge(df_parapoliza, utilitario, on=["CC"], how="left")
    df_parapoliza['CC'] = df_parapoliza['CC'].replace("Total", "E913")
    return df_parapoliza


def obtener_faltantes_utilitario(df_parapoliza):
    datos_faltantes = df_parapoliza[df_parapoliza["UTILITARIO"].isnull()]
    return [] if datos_faltantes.empty else datos_faltantes["CC"].tolist()


def completar_utilitario(df_parapoliza, utilitario, diccionario):
    datos_faltantes = df_parapoliza[df_parapoliza["UTILITARIO"].isnull()].reset_index()
    for i in range(len(datos_faltantes) - 1):
        cc = datos_faltantes["CC"][i]
        if cc != "E913":
            indice = int(datos_faltantes["index"][i])
            cod_util = diccionario.get(cc, "").upper()
            df_parapoliza.loc[indice, "UTILITARIO"] = cod_util
            renglon = pd.Series({"CC": cc, "UTILITARIO": cod_util}).to_frame().T
            utilitario = pd.concat([utilitario, renglon], ignore_index=True)
    return df_parapoliza, utilitario


def hacer_poliza_final(df_parapoliza, Referencia: str):
    CF = '64911801CF01'
    GA = '64911801GA01'
    GV = '64911801GV01'
    IVA = '140104010002'
    PROV = '240112900007'

    df_parapoliza = df_parapoliza[['CC', 'UTILITARIO', 'Importe', 'IVA', 'Cargo']]
    df_parapoliza.loc[len(df_parapoliza) - 1, 'Importe'] = df_parapoliza.loc[len(df_parapoliza) - 1, 'Cargo']
    extra = pd.Series({"CC": "", "UTILITARIO": np.nan, 'Importe': df_parapoliza.loc[len(df_parapoliza) - 1, 'IVA']})
    df_parapoliza = pd.concat([df_parapoliza, extra.to_frame().T], ignore_index=True)

    df_parapoliza.insert(0, 'plantilla', ['SVA'] * len(df_parapoliza))
    df_parapoliza.insert(1, 'cons', range(1, len(df_parapoliza) + 1))
    df_parapoliza.insert(2, 'comp', ['100'] * len(df_parapoliza))

    ctas = []
    for _, row in df_parapoliza.iterrows():
        cc = row['CC']
        if not cc:
            ctas.append(IVA)
        elif cc == "E913":
            ctas.append(PROV)
        elif cc == "D981":
            ctas.append(GV)
        elif cc.startswith('E'):
            ctas.append(GA)
        else:
            ctas.append(CF)

    df_parapoliza.insert(3, 'cta', ctas)
    df_parapoliza.insert(6, 'nada', [np.nan] * len(df_parapoliza))
    df_parapoliza.insert(7, 'nada1', [np.nan] * len(df_parapoliza))
    df_parapoliza.insert(8, 'debe/haber', [1 if x.startswith(('6', '1')) else 2 for x in ctas])
    df_parapoliza.drop(columns=['IVA', 'Cargo'], inplace=True)
    df_parapoliza["ref"] = ['Consumo Si Vale ' + Referencia] * len(df_parapoliza)
    df_parapoliza["CC"].replace("", np.nan, inplace=True)
    return df_parapoliza



def elaborar_excel_poliza(dfsucio, df, df4, df_parapoliza):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer,  engine='openpyxl') as writer:
        dfsucio.to_excel(writer, sheet_name="Sheet1")
        df.to_excel(writer, sheet_name="Hoja2", index=False, startcol=7)
        df4.to_excel(writer, sheet_name="Hoja3", index=False)
        df_parapoliza.to_excel(writer, sheet_name="tfgld013", index=False, header=True)
    buffer.seek(0)
    return buffer



#path_archivo_excel = r"C:\Users\SALCIDOA\Downloads\archivo_para_probar_si_vale.xlsx"
#Referencia = "sept_2025"

def main(path_archivo_excel, Referencia: str,
         diccionario_nombre_cc: dict[str, str] | None = None,
         diccionario_util_cc: dict[str, str] | None = None):

    df_sucio = obtener_dfsucio(path_archivo_excel)
    df = obtener_df(df_sucio)
    df.reset_index(inplace=True)

    tabla_incompleta = crear_tabla_con_cc_vacia(df)
    nombre_empleados_sin_cc = nombres_faltantes(tabla_incompleta)[:-1]

    centros_costos = pd.read_excel(
        "https://docs.google.com/spreadsheets/d/1Iy68cztYlqI6fLjE8l4s2D32BQLzUZG9/export?format=xlsx"
    )

    #diccionario_nombre_cc = {nombre: "BZ01" for nombre in nombre_empleados_sin_cc}
    #obtenemos el df con los cc completos
    # El df_empleados_cc es el df con los empleados que tienen cc vacias
    if nombre_empleados_sin_cc and diccionario_nombre_cc is not None:
        df, df_empleados_cc = hacer_verficiacion_v2(tabla_incompleta, diccionario_nombre_cc, centros_costos)
    else:
        df_empleados_cc = centros_costos

    #df3 guarda el utilitario con los cc y los nombres de 
    # los empleados que hacen falta y no son camiones
    df3 = obterner_df_no_camiones(df_empleados_cc)

    #guardar_en_drive( df3 )

    ##iniciamos segunda tabla dinamica y poliza

    # Segunda tabla dinámica y poliza
    df4 = crear_segunda_tabla_din(df)

    #leemos el csv en Drive que relaciona el centro de costos y el utilitario
    #lo convertimos en un dataframe de pandas
    utilitario = pd.read_csv(
        "https://docs.google.com/spreadsheets/d/1gnfLiD1arrr5G7seQi85-f3Cd5n7_miS/export?format=csv&gid=1471990202"
    )

    
    #relacionamos con un left join los dataframe df4 que es la tabla dinamica y 
    #la tabla del utilitario. d4 son totales de los centros de costos
    df_parapoliza = enlazar_con_utilitario(df4, utilitario)

    #df_parapoliza contiene ahora utilitario, cc , cargo, importe e IVA 
    #el utilitario puede ser vacio NaN y es necesario completarlos
    #quitamos de la lista anterior el CC E913 por corresponder al total
    #y no tener un utilitario asociado. Se elimina la ultima entrada
    lista_utilitario_faltantes = obtener_faltantes_utilitario(df_parapoliza)[:-1]

    #si la lista de los utilitarios falntaes
    #no es vacia se completa el diccionario

    

    #diccionario_util_cc = {cc: "IN101" for cc in lista_utilitario_faltantes}
    if len( lista_utilitario_faltantes)!=0  and len( diccionario_util_cc )!=0 :
        #se completa la poliza y el utilitario con los cc y los nombres de los empleados que el usuario debe completar
        #con el diccionario diccionario_usuario1
        df_parapoliza, utilitario = completar_utilitario(df_parapoliza, utilitario, diccionario_util_cc)

    #guardar_en_drive( utilitario )

    df_poliza = hacer_poliza_final(df_parapoliza, Referencia)

     # creamos un objeto de buffer
    archivo_excel_buffer = BytesIO()

    #creamos el archivo excel con las 4 hojas como se indica en google
    #colab
    archivo_excel_buffer = elaborar_excel_poliza(df_sucio, df, df4, df_poliza)

    archivo_excel_buffer.seek(0)
    return archivo_excel_buffer, nombre_empleados_sin_cc, lista_utilitario_faltantes



#with open(r"C:\Users\SALCIDOA\Downloads\algo.xlsx", 'wb') as f:
#            f.write(archivo_excel_buffer.getvalue())


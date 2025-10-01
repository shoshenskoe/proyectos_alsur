import pandas as pd
import numpy as np
from io import BytesIO



#input : path de excel
#limpia algunos renglones
#ouput: dataframe pandas 

def obtener_dfsucio(excel_path):
    dfsucio = pd.read_excel(excel_path, skiprows=[0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15])

    return dfsucio
    
def obtener_df ( df_sucio ):

    df=df_sucio[(df["IVA"]!=0) & (df_sucio["Importe"]!=0) & (df_sucio["Cargo"]!=0)].copy()
    df=df.reset_index()
    df=df.drop('index',axis=1)
    df=df[["Nombre\nEmpleado","Cargo","Importe","IVA"]]
    df=df.rename(columns={'Nombre\nEmpleado':"Nombre Empleado"})

    """Hasta aquí ya tenemos la tabla lista y limpia para seguir con el proceso que marca el manual de ángel."""

    # Proceso que imita la primera  tabla dinámica

    df=df.groupby(['Nombre Empleado']).sum()
    df=df.reset_index()
    #df.drop(columns=['No.Trx', 'Precio Unitario ', 'Litros'], inplace= True )
    #df=df.append({"Nombre Empleado":"Total","Cargo":df.Cargo.sum(),"IVA":df.IVA.sum(), "Importe": df.Importe.sum()}, ignore_index=True)
    S=pd.Series({"Nombre Empleado":"Total", "Importe": df.Importe.sum(),"Cargo":df.Cargo.sum(),"IVA":df.IVA.sum()})
    df=pd.concat([df, S.to_frame().T], ignore_index=True)

    enlace= "https://docs.google.com/spreadsheets/d/1Iy68cztYlqI6fLjE8l4s2D32BQLzUZG9/edit?usp=sharing&ouid=111113060171554295483&rtpof=true&sd=true"
    df2=pd.read_excel(enlace)

    # @title Texto de título predeterminado
    df=pd.merge( df, df2, on='Nombre Empleado', how='left')

    df=df.reindex(["CC","Nombre Empleado","Cargo","Importe","IVA"],axis=1)


    return df


#####aqui empezamos el proceso de verificacion , es un esbozo!!!!!
def nombres_faltantes( df_faltantes):

    lista_nombres = df_faltantes["Nombre Empleado"].tolist()

    return lista_nombres

def rellenar_cc( df_faltantes, lista_nombres, lista_cc):

    return 0


##################aqui termina el esbozo

def hacer_verficacion(df, dic_nombre_cc=None):

    # verificacion

    datos_faltantes=df[df["CC"].isnull()]

    if ( not datos_faltantes.empty):
        
        datos_faltantes=datos_faltantes.reset_index()
        #datos_faltantes

        for i in range (len(datos_faltantes)-1):
            j=datos_faltantes["index"][i]
            cc=input("Introduzca el CC de "+str(datos_faltantes["Nombre Empleado"][i]))
            centro=cc.upper()
            df.loc[j,"CC"]=centro
            S=pd.Series({"CC":centro,"Nombre Empleado":datos_faltantes["Nombre Empleado"][i]})
            df3=pd.concat([df3, S.to_frame().T], ignore_index=True)



    return df, df3


###asumimos que el diccionario es proporcionado por el usuario
##y no es nulo 
#input: df, diccionario con nombre y cc y faltantes; df_empleados_cc es la relacion
#de empleados y cc que se extrae del archivo de Google drive
##output: df con cc rellenados y df_empleados_cc el df con los faltantes
def hacer_verficiacion_v2(df, dic_nombre_cc: dict[str, str], df_empleados_cc):
    for i in range(len(df)):
        nombre = df.loc[i, "Nombre Empleado"]
        if nombre in dic_nombre_cc:
            centro = dic_nombre_cc[nombre] #pedimos el valor de la llave nombre
            df.loc[i, "CC"] = centro.upper()

            #ahora anadimos a df3 para crear una tabla con los faltantes
            S=pd.Series({"CC":centro,"Nombre Empleado": nombre })
            df_empleados_cc = pd.concat([df_empleados_cc, S.to_frame().T], ignore_index=True)
            
    return df, df_empleados_cc

def verificar_no_camiones(df3):

    camiones= ['TRAILER 001','TRAILER 002', 'TRAILER 004', 'TRAILER 05' ]
    df3=df3[~ df3['Nombre Empleado'].isin(camiones) ]
    df3=df3.reset_index(drop=True)
    #df3.to_excel("/content/drive/MyDrive/Si Vale Gasolina Empleados y cc/Base si vale gasolina.xlsx", index= False)

    #df3.to_excel("/content/drive/MyDrive/Si Vale Gasolina Empleados y cc/Base si vale gasolina.xlsx", index= False)

    return df3

def crear_segunda_tabla_din( df ):


    df[["Cargo","Importe","IVA"]]=df[["Cargo","Importe","IVA"]].astype("float64")
    #df=df.round(0)

    df4=df.copy()
    df4["CC"]=df4["CC"].str.upper()
    df4=df4.groupby("CC").sum()


    df4=df4.reset_index()
    #df4=df4.append({"CC":"Total","Cargo":df4.Cargo.sum(),"Importe": df4.Importe.sum(),"IVA":df4.IVA.sum()}, ignore_index=True)
    S=pd.Series({"CC":"Total","Cargo":df4.Cargo.sum(),"Importe": df4.Importe.sum(),"IVA":df4.IVA.sum()}) # con esto resume la tabla en el ultimo renglon
    df4=pd.concat([df4, S.to_frame().T], ignore_index=True) #anade al dataframe el renglon resumen
    df4["Cargo"]=df4.Cargo.astype('float')
    df4["IVA"]=df4.IVA.astype('float')
    df4["Importe"]=df4.Importe.astype('float')
    df4= df4.round(0)
    df4=df4.drop(['Nombre Empleado'],axis=1)

    return df4


def obtener_utilitario(enlace_path):

    utilitario= pd.read_excel(enlace_path)


    return utilitario

def enlazar_con_utilitario( df_parapoliza, utilitario):
    df_parapoliza['CC']=df_parapoliza['CC'].str.strip()
    df_parapoliza=pd.merge(df_parapoliza,utilitario, on=["CC"], how="left")

    df_parapoliza['CC']=df_parapoliza['CC'].replace("Total", "E913")

    return df_parapoliza


def verificar_faltantes (df_parapoliza):

    datos_faltantes=df_parapoliza[df_parapoliza["UTILITARIO"].isnull()]
    boleano = False
    if ( datos_faltantes.empty):
        boleano = True

    return boleano


###esta funcion pide los datos faltantes y los
# anade al df de poliza y al df de utilitario

def completar_utilitario1 ( df_parapoliza, utilitario) :
    
    datos_faltantes=df_parapoliza[df_parapoliza["UTILITARIO"].isnull()]
    datos_faltantes=datos_faltantes.reset_index()
    #datos_faltantes

    for i in range (len(datos_faltantes)-1):

        if datos_faltantes["CC"][i]!="E913":

            j=datos_faltantes["index"][i]
            
            uti=input("Introduzca el centro utilitario de "+str(datos_faltantes["CC"][i]) )

            uti=uti.upper()

            df_parapoliza.loc[j,"UTILITARIO"]=uti #aqui rellena el df para poliza

            S=pd.Series({"CC":datos_faltantes["CC"][i],"UTILITARIO": uti })

            utilitario=pd.concat([utilitario, S.to_frame().T], ignore_index=True)
            enlace = "https://drive.google.com/drive/folders/172IVSmCSHNfAzATjhLq641FAnWdN2PFr?usp=sharing"
            utilitario.to_excel(enlace, index= False)



    return df_parapoliza



def completar_utilitario(df_parapoliza, utilitario, respuestas: dict):
    """
    Completa el campo 'UTILITARIO' en df_parapoliza usando valores provistos por el usuario.

    Args:
        df_parapoliza: DataFrame que contiene la columna "UTILITARIO" y "CC".
        utilitario: DataFrame donde se acumulan las correspondencias CC-UTILITARIO.
        respuestas: dict con clave=CC y valor=utilitario elegido por el usuario.
    
    Returns:
        df_parapoliza actualizado y utilitario actualizado.
    """
    datos_faltantes = df_parapoliza[df_parapoliza["UTILITARIO"].isnull()].reset_index()

    for i in range(len(datos_faltantes)):
        cc = datos_faltantes.loc[i, "CC"]
        if cc != "E913":
            j = datos_faltantes.loc[i, "index"]

            if cc not in respuestas:
                raise ValueError(f"No se proporcionó utilitario para {cc}")

            uti = respuestas[cc].upper()

            # Actualizamos el df original
            df_parapoliza.loc[j, "UTILITARIO"] = uti

            # Guardamos en tabla auxiliar
            S = pd.Series({"CC": cc, "UTILITARIO": uti})
            utilitario = pd.concat([utilitario, S.to_frame().T], ignore_index=True)

            enlace = "https://drive.google.com/drive/folders/172IVSmCSHNfAzATjhLq641FAnWdN2PFr?usp=sharing"
            utilitario.to_excel(enlace, index= False)


    return df_parapoliza, utilitario












def hacer_poliza_final( df_parapoliza, Referencia): 

    # Empecemos a definr las ctas contables que usaremos de acuerdo al catalogo de cuentas

    CF='64911801CF01'  # Centro de costos A,B,C Y D941  a D950

    GA='64911801GA01'  # Centro de costos E

    GV='64911801GV01'  # Centro de costo D981

    IVA_por_acreditable= '140104010002'  # Sin centro de costo

    Prov_consumo_Sivale= '240112900007' # Para el Centro de costo E913


    df_parapoliza=df_parapoliza[['CC','UTILITARIO' ,'Importe', 'IVA', 'Cargo']]

    #ajusta un importe
    #el del ultimo renglon
    df_parapoliza.loc[len(df_parapoliza)-1, 'Importe'] = df_parapoliza.loc[len(df_parapoliza)-1, 'Cargo']


    S=pd.Series({"CC":"","UTILITARIO": np.nan, 'Importe': df_parapoliza.loc[len(df_parapoliza)-1, 'IVA'] })
    #anade un renglon


    df_parapoliza=pd.concat([df_parapoliza, S.to_frame().T], ignore_index=True)

    df_parapoliza.insert(0, 'plantilla', ['SVA' for i in range(len(df_parapoliza))] )
    df_parapoliza.insert(1, 'cons', [i+1 for i in range(len(df_parapoliza))] )
    df_parapoliza.insert(2, 'comp', ['100' for i in range(len(df_parapoliza))] )


    cta=[]

    for i in range(len(df_parapoliza)):

        CC= df_parapoliza.loc[i, 'CC']

        if not  CC.startswith('E')  and  CC != "D981" and CC !="":
            cta.append(CF)

        elif CC.startswith('E') and CC!="E913" :
            cta.append(GA)

        elif CC=="D981" and CC !="" :
            cta.append(GV)

        elif CC=="E913" and CC !="" :

            cta.append(Prov_consumo_Sivale)

        elif CC== "":

            cta.append(IVA_por_acreditable)

    df_parapoliza.insert(3, 'cta', cta)

    df_parapoliza.insert(6, 'nada', [np.nan for i in range(len(df_parapoliza))] )

    df_parapoliza.insert(7, 'nada1', [np.nan for i in range(len(df_parapoliza))] )

    df_parapoliza.insert(8, 'debe/haber', [1 if df_parapoliza.loc[i, 'cta'] .startswith('6') or df_parapoliza.loc[i, 'cta'] .startswith('1')  else 2 for i in range(len(df_parapoliza))] )

    df_parapoliza.drop(columns=['IVA','Cargo'], inplace= True)

    df_parapoliza["ref"]= [ 'Consumo Si Vale '+ Referencia for i in range(len(df_parapoliza))]

    df_parapoliza["CC"]=df_parapoliza["CC"].replace("", np.nan)

    return df_parapoliza

#input: dataframes pandas cuyos nombres indican su elaboracion previamente
#output: un archivo buffer que puede ser utilizado por la interfaz
def elaborar_excel_poliza(dfsucio, df, df4, df_parapoliza ):

    # creamos un objeto de buffer
    archivo_buffer = BytesIO()

    with pd.ExcelWriter(archivo_buffer) as writer:
        dfsucio.to_excel(writer,sheet_name="Sheet1")
        #dftd1.to_excel(writer, sheet_name="Hoja2",index=False, startcol=0)
        df.to_excel(writer, sheet_name="Hoja2", index= False, startcol=7)
        df4.to_excel(writer,sheet_name="Hoja3", index= False)
        df_parapoliza.to_excel(writer,sheet_name="tfgld013", index= False, header= False)
    # movemos al encabezado de la memoria para lectura
    archivo_buffer.seek(0)

    return archivo_buffer

def logica_principal( path_archivo_excel, path_base_sivale ):

    dfsucio = obtener_dfsucio(path_archivo_excel)

    df= obtener_df(dfsucio)

    """# verificación"""

    df = hacer_verficacion(df)

    df3 = verificar_no_camiones(df)

    # Segunda tabla dinámica y poliza

    df4 = crear_segunda_tabla_din(df)
    
    path_centro_util = "https://docs.google.com/spreadsheets/d/1gnfLiD1arrr5G7seQi85-f3Cd5n7_miS/edit?usp=sharing&ouid=111113060171554295483&rtpof=true&sd=true"
    utilitario = obtener_utilitario(path_centro_util)

    df_parapoliza = enlazar_con_utilitario(df4, utilitario)

    booleano = verificar_faltantes(df_parapoliza)

    if (booleano== True):
        df_parapoliza,_ = completar_utilitario(df_parapoliza,utilitario)

    df_parapoliza = hacer_poliza_final(df_parapoliza)

    archivo_excel_buffer = elaborar_excel_poliza(dfsucio= dfsucio, 
                          df= df, 
                          df_parapoliza= df_parapoliza)

    return archivo_excel_buffer

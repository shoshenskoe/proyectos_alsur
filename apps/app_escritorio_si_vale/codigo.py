import pandas as pd
import numpy as np
from io import BytesIO



#input : path de excel
#limpia algunos renglones
#ouput: dataframe pandas 
def obtener_dfsucio(excel_path):
    dfsucio = pd.read_excel(excel_path, skiprows=[0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15])

    return dfsucio

#imita una tabla dinamica de excel. Trae empleado, cargo, importe e iva
#input: df sucio   
def obtener_df ( df_sucio ):

    df=df_sucio[(df_sucio["IVA"]!=0) & (df_sucio["Importe"]!=0) & (df_sucio["Cargo"]!=0)].copy()
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

    # @title Texto de título predeterminado
    df=pd.merge( df, df2, on='Nombre Empleado', how='left')

    df=df.reindex(["CC","Nombre Empleado","Cargo","Importe","IVA"],axis=1)



    return df


#####aqui empezamos el proceso de verificacion , es un esbozo!!!!!
def nombres_faltantes( df_sin_cc):
    datos_faltantes=df_sin_cc[df_sin_cc["CC"].isnull()]
    lista_nombres = datos_faltantes["Nombre Empleado"].tolist()

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
            #pedimos el valor de la llave nombre
            centro = dic_nombre_cc[nombre] 
            #rellenamos el df original con lo que nos dio el usuario
            df.loc[i, "CC"] = centro.upper()

            #ahora anadimos a df3 para crear una tabla con los faltantes
            S=pd.Series({"CC":centro,"Nombre Empleado": nombre })
            df_empleados_cc = pd.concat([df_empleados_cc, S.to_frame().T], ignore_index=True)
            
    return df, df_empleados_cc



def obterner_df_no_camiones(df3):

    camiones= ['TRAILER 001','TRAILER 002', 'TRAILER 004', 'TRAILER 05' ]
    df3=df3[~ df3['Nombre Empleado'].isin(camiones) ]
    df3=df3.reset_index(drop=True)
    return df3


##esta funcion tiene que ser mas elaborada
##esto es solamente un esbozo
def guardar_en_drive( df3 ):
   
    enlace = r"https://drive.google.com/drive/folders/172IVSmCSHNfAzATjhLq641FAnWdN2PFr?usp=sharing/Base si vale gasolina_prueba.xlsx"
    #df3.to_excel(enlace, index= False)
    return None


#####empieza segunda tabla dinamica y poliza
#esta tabla es una tabla dinamica que agrupa por cc
#muestra cargo, importe e iva
#el ultimo renglon es el total
def crear_segunda_tabla_din( df ):


    df[["Cargo","Importe","IVA"]]=df[["Cargo","Importe","IVA"]].astype("float64")
    #df=df.round(0)

    df["CC"]=df["CC"].str.upper()
    
     #aqui se agrupa por cc y se elimina totAL de df
    df4=df.groupby("CC").sum()


    df4=df4.reset_index()
    #df4=df4.append({"CC":"Total","Cargo":df4.Cargo.sum(),"Importe": df4.Importe.sum(),"IVA":df4.IVA.sum()}, ignore_index=True)
    # con lo siguiente se resume la tabla en el ultimo renglon
    S=pd.Series({"CC":"Total","Cargo":df4.Cargo.sum(),"Importe": df4.Importe.sum(),"IVA":df4.IVA.sum()})
    df4=pd.concat([df4, S.to_frame().T], ignore_index=True) #anade al dataframe el renglon resumen
    df4["Cargo"]=df4.Cargo.astype('float')
    df4["IVA"]=df4.IVA.astype('float')
    df4["Importe"]=df4.Importe.astype('float')
    df4= df4.round(0)
    df4=df4.drop(['Nombre Empleado'],axis=1)

    return df4



##aqui ya tenemo el df para poliza
#ubicamos el utilitario de acuerdo al cc
#input: df de poliza y df de utilitario
#output: df de poliza con utilitario

def enlazar_con_utilitario( df_parapoliza, utilitario):
    df_parapoliza['CC']=df_parapoliza['CC'].str.strip()
    df_parapoliza=pd.merge(df_parapoliza,utilitario, on=["CC"], how="left")

    #reemplazamos un CC especial que sera el total
    df_parapoliza['CC']=df_parapoliza['CC'].replace("Total", "E913")

    return df_parapoliza




def obtener_faltantes_utilitario ( df_parapoliza):
    datos_faltantes=df_parapoliza[df_parapoliza["UTILITARIO"].isnull()]

    if ( datos_faltantes.empty):
        return []
    else:
        lista_cc = datos_faltantes["CC"].tolist()
        return lista_cc


###esta funcion pide los datos faltantes y los
# anade al df de poliza y al df de utilitario

def completar_utilitario ( df_parapoliza, utilitario, diccionario=None) :
    
    datos_faltantes = df_parapoliza[df_parapoliza["UTILITARIO"].isnull()]
    datos_faltantes = datos_faltantes.reset_index()
    #datos_faltantes

    #recorremos los datos faltantes 
    # menos el ultimo renglon que es el total
    for i in range (len(datos_faltantes)-1):

        if datos_faltantes["CC"][i]!="E913":

            indice_df_parapoliza = int( datos_faltantes["index"][i] )
            
            cod_util = diccionario.get( datos_faltantes["CC"][i], None)


            cod_util=cod_util.upper()


            ####### aqui rellena el df para poliza, revisar el indice!!
             #aqui rellena el df para poliza
            df_parapoliza.loc[indice_df_parapoliza,"UTILITARIO"]=cod_util

            renglon=pd.Series({"CC":datos_faltantes["CC"][i],"UTILITARIO": cod_util })
            
            renglon = renglon.to_frame().T

            utilitario=pd.concat( [utilitario, renglon] , ignore_index=True)

    return df_parapoliza, utilitario






######funcion para crear la poliza final

def hacer_poliza_final( df_parapoliza, Referencia : str ): 

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


#path_archivo_excel = r"C:\Users\SALCIDOA\Downloads\archivo_para_probar_si_vale.xlsx"


def main( path_archivo_excel , Referencia:str ):

    dfsucio = obtener_dfsucio(path_archivo_excel)

    #la funcion obtener df imita una tabla dinamica de excel
    df= obtener_df(dfsucio)
    df.reset_index(inplace=True)

    #es la misma tabla que la primera pero posiblemente con cc vacias
    tabla_incompleta = crear_tabla_con_cc_vacia(df)

    #obtenemos los nombres de los empleados que tienen cc vacias
    nombre_empleados_sin_cc = nombres_faltantes(tabla_incompleta)

    #quitamos el total de la lista nombre_empleados_sin_cc
    nombre_empleados_sin_cc = nombre_empleados_sin_cc[:-1]
    
    #almacenamos los centros de costos en un dataframe
    #posiblemente este paso sea redundante
    #  con la funcion crear_tabla_con_cc_vacia
    path_centro_util = "https://docs.google.com/spreadsheets/d/1Iy68cztYlqI6fLjE8l4s2D32BQLzUZG9/export?format=xlsx"
    centros_costos = pd.read_excel(path_centro_util)

    #si la lista de los empleados no es vacia se completa el diccionario
    if nombre_empleados_sin_cc:
        diccionario_nombre_cc = {nombre: "" for nombre in nombre_empleados_sin_cc}
        
        for empleado in nombre_empleados_sin_cc:
            cc = input(f"Introduzca el CC de {empleado}: ")
            diccionario_nombre_cc[empleado] = str(cc).upper()


        #obtenemos el df con los cc completos
        # El df_empleados_cc es el df con los empleados que tienen cc vacias
        df, df_empleados_cc = hacer_verficiacion_v2(df=tabla_incompleta, 
                                                    dic_nombre_cc=diccionario_nombre_cc, 
                                                    df_empleados_cc= centros_costos)
    


    #df3 guarda el utilitario con los cc y los nombres de 
    # los empleados que hacen falta y no son camiones
    df3 = obterner_df_no_camiones ( df_empleados_cc )

    #guardar_en_drive( df3 )

    ##iniciamos segunda tabla dinamica y poliza

    # Segunda tabla dinámica y poliza

    df4 = crear_segunda_tabla_din(df)


    #leemos el csv en Drive que relaciona el centro de costos y el utilitario
    #lo convertimos en un dataframe de pandas
    path_cc_utilitario = "https://docs.google.com/spreadsheets/d/1gnfLiD1arrr5G7seQi85-f3Cd5n7_miS/export?format=csv&gid=1471990202"
    utilitario = pd.read_csv(path_cc_utilitario)

    #relacionamos con un left join los dataframe df4 que es la tabla dinamica y 
    #la tabla del utilitario. d4 son totales de los centros de costos
    df_parapoliza = enlazar_con_utilitario(df4, utilitario)

    #df_parapoliza contiene ahora utilitario, cc , cargo, importe e IVA 
    #el utilitario puede ser vacio NaN y es necesario completarlos

    lista_utilitario_faltantes = obtener_faltantes_utilitario(df_parapoliza)

    #quitamos de la lista anterior el CC E913 por corresponder al total
    #y no tener un utilitario asociado. Se elimina la ultima entrada
    lista_utilitario_faltantes = lista_utilitario_faltantes[:-1]



     #si la lista de los utilitarios faltantes
     #no es vacia se completa el diccionario
    if lista_utilitario_faltantes:
        diccionario_util_cc = {CC: "" for CC in lista_utilitario_faltantes}
        
        for cc in diccionario_util_cc:
            util = input(f"Introduzca el utilitario de {cc}: ")
            diccionario_util_cc[cc] = str(util).upper()

        #se completa la poliza y el utilitario con los cc y los nombres de los empleados que el usuario debe completar
        #con el diccionario diccionario_usuario1
        df_parapoliza , utilitario= completar_utilitario ( df_parapoliza = df_parapoliza, 
                                                        utilitario=utilitario, 
                                                        diccionario= diccionario_util_cc )

    
    #guardar_en_drive( utilitario )

    df_poliza = hacer_poliza_final(df_parapoliza, Referencia=Referencia)

     # creamos un objeto de buffer
    archivo_excel_buffer = BytesIO()

    #creamos el archivo excel con las 4 hojas como se indica en google
    #colab
    archivo_excel_buffer = elaborar_excel_poliza(dfsucio, df, df4, df_poliza)

    archivo_excel_buffer.seek(0)
    
    return archivo_excel_buffer
import pandas as pd
import numpy as np
import requests
from io import BytesIO #para guardar el Excel como bytes


def reclasificacion ( ingresos_df , gastos_df, mes, mes_numero ):

    mes = str(mes)
    mes_numero = str(mes_numero)

    #necesitamos estos id para guardor los xlsx que estan en drive
    #se obtienen viendo el enlace en el navegador
    sheet_id_bodegas = "1q5UA81-N5K8UoCg7HrPpyI99neQwgCBA"
    sheet_id_cruces = "1GlPczUnK8TfYn_9pooMSMIdei4MvooCD"

    # generamos el enlace necesario
    enlace_bodegas = f"https://docs.google.com/spreadsheets/d/{sheet_id_bodegas}/export?format=xlsx"
    enlace_cruce  = f"https://docs.google.com/spreadsheets/d/{sheet_id_cruces}/export?format=xlsx"
    
    #leemos el excel guardado en esas direcciones web con Pandas 
    #bodegas_df = pd.read_excel(enlace_bodegas)
    #cruces_df = pd.read_excel(enlace_cruces)

    dfbaan=pd.read_excel(enlace_cruce)

    df = gastos_df.drop( [0,1,2,3,4,5,6,7,8,10] )
    #df=pd.read_excel( ingresos_xlsx ,skiprows=[0,1,2,3,4,5,6,7,8,10])
    df.rename(columns={'Cta. cont.   ':"Cuenta", '      ':"Sub", '                                   ':"Descripcion",
       '                    ':"Saldoapertura", '               Debe ':"Debe", '              Haber ':"Haber",
       '               Debe .1':"Debe1", '              Haber .1':"Haber1",
       '    Saldo de cierre':"Saldocierre"}, inplace= True) #Cuando pones "rename(columns)" haces un diccionario
    
    # Empecemos limpiando las cuentas

    df=df.query(' not  ( Cuenta.str.startswith("-") | Cuenta.str.startswith("Total")  )').copy()

    df["Cuenta"]=df["Cuenta"].str.replace("'","")

    df["Descripcion"]=df["Descripcion"].str.strip()

    df["Cuenta"]=df["Cuenta"].str.strip()

    df["Debe1"]=df["Debe1"].replace('                    ',"0")
    #df["Debe1"]=df["Debe1"].str.replace(' ',"")
    df["Haber1"]=df["Haber1"].replace('                    ',"0")
    #df["Haber1"]=df["Haber1"].str.replace(' ',"")

    # Codigo para

    a=""
    Bodega=[]

    for i in range(len(df)):
        if df.iloc[i,0].startswith("5") or df.iloc[i,0].startswith("6"):
            Bodega.append(a)
        else:
            a=df.iloc[i,0]
            Bodega.append(a)

    df["Bodega"]=Bodega


    df=df.query(' Descripcion.isnull()== False').copy()

    for i in range(len(df)):
        J=str(df.iloc[i,6])
        J=J.replace(" ","")
        if J.endswith("-"):
            df.iloc[i,6]= -1*float(J[0:len(J)-1])

    for i in range(len(df)):
        J=str(df.iloc[i,7])
        J=J.replace(" ","")
        if J.endswith("-"):
            df.iloc[i,7]= -1*float(J[0:len(J)-1])

    df["Debe1"]=df["Debe1"].astype("float64")
    df["Haber1"]=df["Haber1"].astype("float64")
    df["Bodega"].str.split(" 0 ",expand= True)

    df["Bodega"]=df["Bodega"].str.strip()
    df[["CC","bodega"]]=df["Bodega"].str.split(" 0 ",expand= True)
    df["CC"]=df["CC"].str.replace(" ","")
    df["bodega"]=df["bodega"].str.strip()
    df["ingresos"]=-1*df["Debe1"]+df["Haber1"]

    df["CUENTA"]=df["Cuenta"]+"_"+df["Descripcion"]

    df=df.query(' ingresos> 0').reset_index().copy()

    df["Bodega"]=df["CC"]+"_"+df["bodega"]
    #df["Bodega"]=df["Bodega"].str.strip()

    df1=df[["CC","CUENTA","ingresos"]]

    """ Realicemos una tabla dináica por Bodega"""

    name="ingresos "+ mes + ".xlsx"
    df1.groupby("CC").sum().reset_index().to_excel(name, index= False)
    df_agrupado = df1.groupby("CC").sum().reset_index()

    # Lipiemos los gastos de bz01-bz15
    #df_gasto=pd.read_excel(gastos_df)

    df_gasto = gastos_df
    df = gastos_df.drop( [0,1,2,3,4,5,6,7,8,10] )
    #df=pd.read_excel( gastos_df ,skiprows=[0,1,2,3,4,5,6,7,8,10])
    df.rename(columns={'Cta. cont.   ':"Cuenta", '      ':"Sub", '                                   ':"Descripcion",
        '                    ':"Saldoapertura", '               Debe ':"Debe", '              Haber ':"Haber",
        '               Debe .1':"Debe1", '              Haber .1':"Haber1",
        '    Saldo de cierre':"Saldocierre"}, inplace= True)
        

     # Empecemos limpiando las cuentas

    df=df.query(' not  ( Cuenta.str.startswith("-") | Cuenta.str.startswith("Total")  )').copy()

    df["Cuenta"]=df["Cuenta"].str.replace("'","")

    df["Descripcion"]=df["Descripcion"].str.strip()

    df["Cuenta"]=df["Cuenta"].str.strip()

    df["Debe1"]=df["Debe1"].replace('                    ',"0")
    #df["Debe1"]=df["Debe1"].str.replace(' ',"")
    df["Haber1"]=df["Haber1"].replace('                    ',"0")
    #df["Haber1"]=df["Haber1"].str.replace(' ',"")

    a=""
    Bodega=[]

    for i in range(len(df)):
        if df.iloc[i,0].startswith("5") or df.iloc[i,0].startswith("6"):
            Bodega.append(a)
        else:
            a=df.iloc[i,0]
            Bodega.append(a)

    df["Bodega"]=Bodega


    #df=df.query(' Descripcion.isnull()== False').copy()


    for i in range(len(df)):
        J=str(df.iloc[i,6])
        J=J.replace(" ","")
    if J.endswith("-"):
        df.iloc[i,6]= -1*float(J[0:len(J)-1])

    for i in range(len(df)):
        J=str(df.iloc[i,7])
        J=J.replace(" ","")
    if J.endswith("-"):
        df.iloc[i,7]= -1*float(J[0:len(J)-1])

    df["Debe1"]=df["Debe1"].astype("float64")
    df["Haber1"]=df["Haber1"].astype("float64")

    df["Bodega"]=df["Bodega"].str.strip()
    df[["CC","bodega"]]=df["Bodega"].str.split(" 0 ",expand= True)
    df["CC"]=df["CC"].str.replace(" ","")
    df["bodega"]=df["bodega"].str.strip()

    name2= "Gastos "+ mes + " bz.xlsx"
    df[["CC", "Cuenta","Descripcion","Saldocierre"]].to_excel(name2, index= False)
    dfgasto=df[["CC", "Cuenta","Descripcion","Saldocierre"]].copy()

    """# Validamos que todas las bodegas estén"""

    df=pd.read_excel(name)

    df2=pd.read_excel( enlace_cruce, sheet_name='Nuevos datos')
    df2["Cuenta"]=df2['Cuenta'].str.upper()
    columnas=df2.Cuenta.unique()
    #df2.columns
    df2

    df.query(' not (  CC in @columnas ) ')

    # Prorrateo

    ingresos=pd.read_excel(name)

    ingresos

    gasto=pd.read_excel(name2)
    gasto.CC.unique()

    """De las bodegas anteriores nos damos cuenta cuáles no traen gastos como es el caso de la BZ06, BZ14. En esta situación, en el código siguiente tachamos la bodega con un #

    Además recordemos que tenemos bodegas con bdegas supervisadas
    """

    ctas=pd.read_excel( enlace_cruce , sheet_name='Nuevos datos')
    ctas["Cuenta"]=ctas['Cuenta'].str.upper()

    #ctas['Bodega']=ctas['Bodega'].replace('BZ09             0   REGIONAL HAB CHIHUAHUA','BZ11             0   REGIONAL ADM DURANGO')

    bodega=['BZ01             0   REGIONAL HAB CULIACAN',
        'BZ02             0   REGIONAL HAB GUADALAJARA',
        'BZ03             0   REGIONAL HAB VERACRUZ',
        'BZ04             0   REGIONAL HAB CHIAPAS',
        #'BZ05             0   REGIONAL ADM ZACATECAS'  bodegas en supervision
        #'BZ06             0   REGIONAL HAB PUEBLA', NO DEBE DE EXISTIR MOVIMIENTOOOOOS , SI TIENE HACEMOS EL PRORRATEO MANUAL
        'BZ07             0   REGIONAL HAB GUANAJUATO',
        'BZ08             0   REGIONAL ADM MONTERREY',
        #'BZ09             0   REGIONAL HAB CHIHUAHUA'  bodegas en supervision
        'BZ10             0   REGIONAL HAB SONORA',
        #'BZ11             0   REGIONAL ADM DURANGO',   AQUI HAY CHUECO, EL COSTO DEBE DE IR EN LA BZ09.
        'BZ12             0   REGIONAL HAB SLP',
        'BZ13             0   REGIONAL HAB PENINSULAR',
        'BZ14             0   REGIONAL HAB MORELOS'
        'BZ15            0   REGIONAL HAB ESTADO MEX']

    ctas[["CC","basura",]]=ctas["Bodega "].str.split(" 0 ", expand= True)


    CC=['BZ01','BZ02','BZ03','BZ04', 'BZ07','BZ08','BZ10','BZ12','BZ13','BZ15']

    # Parte para los cambios chuechos

    ctas=pd.read_excel(enlace_cruce, sheet_name='Nuevos datos')
    ctas["Cuenta"]=ctas['Cuenta'].str.upper()

    ctas['Bodega ']=ctas['Bodega '].replace('BZ09             0   REGIONAL HAB CHIHUAHUA','BZ11             0   REGIONAL ADM DURANGO')

    bodega=['BZ11             0   REGIONAL ADM DURANGO']
    ctas[["CC","basura",]]=ctas["Bodega "].str.split(" 0 ", expand= True)


    CC=['BZ11']

    #

    # la parte de cambbios chuechos no  va a aaqii!!!!!!!

    Misdatos=[]
    Epsilon=0
    for bodega_aponderar in bodega:


        tablas=ctas.loc[ctas["Bodega "]== bodega_aponderar]

        tablas  # Extraemos la porcion de ctas asociadas a cada bodega

        # En tablas puede que entre en accion un if para las bodegas supervisadas

        Cuenta=tablas.Cuenta.values # Volvemos en np.array las bodegas que e tocan a cada BZ


        Cuenta=Cuenta.tolist()   # Convertimos en listas de python

        Ingreso=ingresos.query(' CC in @Cuenta  ')
        
        # Add a check to see if Ingreso is empty
        if Ingreso.empty:
            print(f"Warning: No matching income data found for bodega: {bodega_aponderar}")
            continue # Skip to the next bodega if no income data is found

        Ingreso["Agrupar"]=[ 1 for i in range(len(Ingreso))]

        Ingreso  

        prorrateo=pd.pivot(Ingreso, columns= "CC", values="ingresos", index= "Agrupar").reset_index(drop= True)

        # Creamos un df donde las filas son los ingresos de la tabla

        Total=prorrateo.iloc[0].sum()

        peso=prorrateo.iloc[0]/Total

        prorrateo.loc[len(prorrateo.index)]=peso

        prorrateo.reset_index(inplace= True)



        gastos=gasto.loc[gasto["CC"]==CC[Epsilon]] # Empezamos a trabajar con las BZ'S

        gastos["Abs"]=gastos["Saldocierre"].abs()

        gastos=gastos.sort_values(by=["Abs"])

        gastos=gastos.query('Saldocierre.isnull()== False').reset_index()  # Nos desasemos de los nulos



        matriz1=[gastos.iloc[j,4] for j in range(len(gastos)) ]  # Obtenemos los valores de los gastos



        matriz1=np.array(matriz1)

        matriz1=matriz1.reshape(-1,1)  # Volvemos los valores una matriz columna 1xn


        Pesos=peso.values.reshape(1,-1)

        Pesos  # LO volvemos una matriz nx1

        resultado=np.dot(matriz1,Pesos)

        prorrateo.drop(columns='index', inplace= True)


        for k in range(len(resultado)):
            prorrateo.loc[len(prorrateo.index)]=resultado[k]


        # Peguemos las ctas

        gastos["Saldocierre"]=-1*gastos["Saldocierre"]

        Labodega=[]
        Lacuenta=[]
        Ladesc=[]
        Leimporter=[]
        for i in range(len(prorrateo)):
            if i ==0 or i ==1:
                Labodega.append(np.nan)
                Lacuenta.append(np.nan)
                Ladesc.append(np.nan)
                Leimporter.append(np.nan)
            else:

                i=i-2
                Labodega.append(gastos.loc[i,"CC"])
                Lacuenta.append(gastos.loc[i,"Cuenta"])
                Ladesc.append(gastos.loc[i,"Descripcion"])
                Leimporter.append(gastos.loc[i,"Saldocierre"])

        prorrateo.insert(0,'Bodega', Labodega)
        prorrateo.insert(1,'Cuenta', Lacuenta)
        prorrateo.insert(2,'Descripcion', Ladesc)
        prorrateo.insert(3,'Saldocierre', Leimporter)

        # Codigo para quitar los pesos a los valores <1200
        Pesos_max=prorrateo.iloc[1].values

        Indice_maximo=np.where(Pesos_max == Pesos_max.max())

        Indice_maximo= Indice_maximo[0][0]

        for i in range(prorrateo.shape[1]-4):
            i= i+4
            if i== Indice_maximo:
                for j in range(len(prorrateo)):
                    if (j!=0 or j != 1) and abs(prorrateo.loc[j,"Saldocierre"])<=1200:
                        prorrateo.iloc[j, i]=-1*prorrateo.loc[j,"Saldocierre"]

            else:
                for j in range(len(prorrateo)):
                    if (j!=0 or j != 1) and abs(prorrateo.loc[j,"Saldocierre"])<=1200:
                        prorrateo.iloc[j, i]= 0



        Misdatos.append(prorrateo)

        Epsilon+=1

    # Trabajemos con las bodegas que tienen en supervisión
    #resagados=[
    #      'BZ05             0   REGIONAL ADM ZACATECAS',
    #
    #      'BZ09             0   REGIONAL HAB CHIHUAHUA',
    #    ]

    #CC=['BZ05','BZ09']


    resagados=['BZ11             0   REGIONAL ADM DURANGO']
    CC=['BZ11']

    Epsilon=0

    for bodega_aponderar in resagados:

        tablas=ctas.loc[ctas["Bodega "]== bodega_aponderar]

        Cuenta=tablas.Cuenta.values # Volvemos en np.array las bodegas que e tocan a cada BZ


        Cuenta=Cuenta.tolist()   # Convertimos en listas de python

        Ingreso=ingresos.query(' CC in @Cuenta  ')

        Ingreso["Agrupar"]=[ 1 for i in range(len(Ingreso))]

        Num_bodegas_con_igreso= len(Ingreso["CC"].values)

        prorrateo=pd.pivot(Ingreso, columns= "CC", values="ingresos", index= "Agrupar").reset_index(drop= True)



        # Creamos un df donde las filas son los ingresos de la tabla

        if bodega_aponderar.startswith("BZ05"):

            a=2

            prorrateo["A44G"]=[np.nan]

            prorrateo["BM56"]=[np.nan]

        else :

            a=3

            prorrateo["BA58"]=[np.nan]

            prorrateo["BA26"]=[np.nan]

            prorrateo["BM90"]=[np.nan]

        Reparto_de_peso= a+ Num_bodegas_con_igreso

        peso=[ 1/Reparto_de_peso for i in range(Reparto_de_peso)]

        prorrateo.loc[len(prorrateo.index)]=peso

        prorrateo.reset_index(inplace= True, drop= True)


        gastos=gasto.loc[gasto["CC"]==CC[Epsilon]] # Empezamos a trabajar con las BZ'S

        gastos["Abs"]=gastos["Saldocierre"].abs()

        gastos=gastos.sort_values(by=["Abs"])

        gastos=gastos.query('Saldocierre.isnull()== False').reset_index()  # Nos desasemos de los nulos



        Saldo=gastos["Saldocierre"].values.tolist()  # Obtenemos los valores de los gastos

        Saldo=[ Saldo[i]* (1/Reparto_de_peso) for i in range(len(Saldo))]

        J=0

        for saldo in Saldo:

            J=len(prorrateo)+J

            for columna in prorrateo.columns:
                prorrateo.loc[J, columna]=saldo

            J+=1


        # Peguemos las ctas

        gastos["Saldocierre"]=-1*gastos["Saldocierre"]

        Labodega=[]
        Lacuenta=[]
        Ladesc=[]
        Leimporter=[]
        for i in range(len(prorrateo)):
            if i ==0 or i ==1:
                Labodega.append(np.nan)
                Lacuenta.append(np.nan)
                Ladesc.append(np.nan)
                Leimporter.append(np.nan)
            else:

                i=i-2
                Labodega.append(gastos.loc[i,"CC"])
                Lacuenta.append(gastos.loc[i,"Cuenta"])
                Ladesc.append(gastos.loc[i,"Descripcion"])
                Leimporter.append(gastos.loc[i,"Saldocierre"])
        prorrateo.insert(0,'Bodega', Labodega)
        prorrateo.insert(1,'Cuenta', Lacuenta)
        prorrateo.insert(2,'Descripcion', Ladesc)
        prorrateo.insert(3,'Saldocierre', Leimporter)

        prorrateo.reset_index(drop= True, inplace= True)

        Misdatos.append(prorrateo)

        Epsilon+=1


    for delta in range(len(Misdatos)):

        prueba= Misdatos[delta]

        minimo = prueba.iloc[1,].idxmin()

        prueba2=prueba.iloc[:2].reset_index(drop=True)
        prueba1 =prueba.iloc[2:].reset_index(drop=True)

        prueba1=prueba1.round(2)

        prueba1=prueba1.drop(columns=minimo)

        prueba1[minimo]= [ -prueba1.iloc[i,3:].sum()   for i in range(len(prueba1))]

        Misdatos[delta]=pd.concat([prueba2, prueba1], ignore_index=True)

    # tfgld generalizado

    RBZ=[]
    La_bodega=[]
    cuenta2=[]
    debehaber=[]
    Importe2=[]


    for DataFrame in Misdatos:

        Bodegas2=DataFrame.columns[0:1].tolist() + DataFrame.columns[4:].tolist()

        for bodega in Bodegas2:
            if bodega== "Bodega":
                ccentro = DataFrame.Bodega.unique()[1]

            for i in range(len(DataFrame)-2):
                i=i+2
                RBZ.append("RBZ")
                cuenta2.append(DataFrame.loc[i,'Cuenta' ])
                debehaber.append(1)
                Importe2.append( DataFrame.loc[i,'Saldocierre' ])
                La_bodega.append( ccentro)
            else:
                for i in range(len(DataFrame)-2):
                    i=i+2
                    RBZ.append("RBZ")
                    cuenta2.append(DataFrame.loc[i,'Cuenta' ])
                    debehaber.append(2)
                    Importe2.append( DataFrame.loc[i,bodega ])
                    La_bodega.append(bodega)

    tfgld=pd.DataFrame({'RBZ': RBZ, 'Cuenta':cuenta2,  'Bodega':La_bodega ,'debehaber':debehaber, 'Importe':Importe2})
    tfgld=tfgld.query('Importe != 0')

    tfgld.insert(1, "cons",[i+1 for i in range(len(tfgld))])
    tfgld.insert(2, "Compañia",[100 for i in range(len(tfgld))])
    tfgld.insert(5, "in",["in2101" for i in range(len(tfgld))])
    tfgld.insert(6, "g",[np.nan for i in range(len(tfgld))])
    tfgld.insert(7, "h",[np.nan for i in range(len(tfgld))])
    tfgld["debehaber"]=[1 for i in range(len(tfgld))]
    tfgld["ref"]=["reclasif BZ "+ mes + " 2024" for i in range(len(tfgld))]

    inicio=0
    Guardarcomo= mes_numero+ " reclasificacion bhz "  + mes +  " 2025"  + '.xlsx'

    # Guardamos cada DataFrame en una hoja distinta dentro del Excel
    with pd.ExcelWriter(Guardarcomo) as writer:
        dfbaan.to_excel(writer,sheet_name='Ingresos baan',index= False)
        df_agrupado.to_excel(writer,sheet_name='Ingresos limpios',index= False)
        df_gasto.to_excel(writer,sheet_name='Gasto BAAN',index= False)
        dfgasto.to_excel(writer,sheet_name='Gasto limpio',index= False)
        tfgld.to_excel(writer,sheet_name='tfgld013',index= False,header=False)
        for df in Misdatos:
            df.to_excel(writer,sheet_name='Prorrateo', startrow=inicio,index= False)
            inicio+= df.shape[0]+25  # dejamos 25 filas en blanco entre cada bloque

    # Creamos un buffer en memoria (archivo virtual) y usamos la biblioteca io
    archivo_memoria = BytesIO()

    # Regresamos el puntero al inicio del archivo en memoria
    archivo_memoria.seek(0)

    # Construimos el nombre dinámico del archivo
    nombre_archivo = f"{mes_numero} reclasificacion bhz {mes} 2025.xlsx"

    # Devolvemos el archivo en memoria y el nombre del archivo
    return archivo_memoria, nombre_archivo




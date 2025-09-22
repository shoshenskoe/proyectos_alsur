
def procesamiento_archivos (mes, Referencia, archivo1_path, archivo2_path):

  mes = mes.lower()
  meses = {
      "enero": 1,
      "febrero": 2,
      "marzo": 3,
      "abril": 4,
      "mayo": 5,
      "junio": 6,
      "julio": 7,
      "agosto": 8,
      "septiembre": 9,
      "octubre": 10,
      "noviembre": 11,
      "diciembre": 12
  }

  """# ISR- PTU
  """

  """La limpieza del archivo que obtendremos cosiste en extraer las cuentas 5-6 que son ingresos y gastos"""

  pd.set_option('display.float_format', '{:.2f}'.format)
  df=pd.read_excel(archivo1_path,skiprows=[0,1,2,3,4,5,6,7,8,10])
  df.rename(columns={'cont.        ':'cuenta', '     ':'subn', '                                ':'Descripcion',
        '                    ':'Saldoapertura', '                Debe':'Debe', '               Haber':'Haber',
        '                Debe.1':'Debe1', '               Haber.1':'Haber1', 'Unnamed: 8':'Ingreso'}, inplace= True)
  df=df.query("Descripcion.isnull()== False")
  

  # Extraemos los puros ingresos:

  df["cuenta"]=df["cuenta"].astype("str")

  df=df[['cuenta', 'Descripcion', 'Saldoapertura', 'Debe', 'Haber',
        'Debe1', 'Haber1', 'Ingreso']]
  dflimpio=df.query(' cuenta.str.startswith("5") | cuenta.str.startswith("6")  ' )

  dflimpio=dflimpio.reset_index(drop= True)

  dfisr=df.copy()
  dfisr["cuenta"]=df["cuenta"].str.replace(" ","")
  ISR=dfisr.query('cuenta== "66010101IR01" ')
  ISR=ISR.reset_index()
  ISR=ISR.iloc[0,6]
  PTU=dfisr.query('cuenta== "66010101IR02" ')
  PTU=PTU.reset_index()
  PTU=PTU.iloc[0,6]

  ################se regresa PTU e ISR

  #########################

  """# Gastos y poliza

  Ahora, en baan de SMA sacamos el balance de comprobación (dimnesión/ cuenta contable) como a continuación se especifíca:

  Procedemos a limpiar la información. El objetivo es obtener una tabla con cc, si/no, cuenta, gasto del mes

  """

  nombrearchivo2=mes+" 2"+".xlsx"

  df2=pd.read_excel(archivo2_path,skiprows=[0,1,2,3,4,5,6,7,8,10])
  df2_sucio=pd.read_excel(nombrearchivo2)
  df2.rename(columns={'Cta. cont.   ': "Cuenta", '      ': "Sub", '                                   ': "Descripcion",
        '                    ':"Saldoapertura", '               Debe ':"Debe", '              Haber ':"Haber",
        '               Debe .1':"Debe1", '              Haber .1':"Haber1",
        '    Saldo de cierre':"Saldocierre"}, inplace= True)
  df2["Cuenta"]=df2["Cuenta"].str.replace("'","")
  df2=df2.query('not ( Cuenta.str.startswith("-")  | Cuenta.str.startswith("Total"))')
  df2.reset_index(drop= True, inplace= True)


  Bodega=[]
  a=""
  for i in range(len(df2)):
    if df2.iloc[i,0].startswith("6"):
      Bodega.append(a)
    else:
      a=df2.iloc[i,0]
      Bodega.append(a)
  df2["cc"]=Bodega


  df2["Cuenta"]=df2["Cuenta"]+"_"+df2["Descripcion"]

  df2=df2.query(' Cuenta.isnull()= =False')
  df2.reset_index(drop= True,inplace= True)


  sino=[]
  for i in range(len(df2)):
    #variable=df2.iloc[i,0].replace(" ","")
    if df2.iloc[i,0].replace(" ","").startswith("641001010001_SUELDOSYPRESTACIONES"):
      sino.append("sueldo")
    else:
      sino.append("extra")
  df2["sino"]=sino


  df3=df2.copy()
  df3=df3[["cc","sino","Cuenta","Debe1"]]

  df3.rename(columns={"Debe1":mes}, inplace= True)



  df3[mes]=df3[mes].replace('                    ',"0")

  for i in range(len(df3)):

    if str(df3.loc[i, mes]).strip().endswith('-'):
      a=str(df3.loc[i, mes]).strip()
      df3.loc[i, mes]= float(a[:-2])*-1

  df3[mes]=df3[mes].astype("float64")

  df3=df3.loc[df3[mes] !=0]


  """Hacemos una tabla dinamica por CC, sumando extras y sueldos
  """

  tabladin=pd.pivot_table(df3, index="cc", columns="sino", values= mes, aggfunc="sum")
  tabladin.reset_index(inplace=True)
  tabladin.fillna(0, inplace= True)

  #Ahora con la info de ISR, PTU extraida al principio, crearemos la siguiente tabla:

  #donde la columna isr-ptu se obtiene por medio de:
  #isr-ptu =  (extras)* (\frac{ISR+PTU}{ total \enspace de \enspace  Extras })


  tabladin2=tabladin.copy()
  tabladin2["isr-ptu"]=tabladin["extra"]*(ISR+PTU)/(tabladin2["extra"].sum())
  tabladin2["Total de extras"]=tabladin2["extra"]+tabladin2["isr-ptu"]
  tabladin2["Total general"]= tabladin2["Total de extras"]+tabladin2["sueldo"]


  S=pd.Series({"cc":"Total","extra":tabladin2["extra"].sum(),"sueldo":tabladin2["sueldo"].sum(), "isr-ptu":tabladin2["isr-ptu"].sum(),"Total de extras":tabladin2["Total de extras"].sum(),"Total general":tabladin2["Total general"].sum()})

  tabladin2=pd.concat([tabladin2, S.to_frame().T], ignore_index= True)

  tabladin3=tabladin2[["extra", "isr-ptu", "Total de extras", "sueldo", "Total general"]].astype(float)

  tabladin3=tabladin3.round(0)

  tabladin3["cc"]=tabladin2["cc"]

  tabladin3=tabladin3[["cc","extra", "isr-ptu", "Total de extras", "sueldo", "Total general"]]


  """Se van a registrar en las cuentas:
  - 64913101CF01, 64913101GA01 el sueldo en debe. Recordemos que los CC con C van en CF y con E en GA
  - el total de sueldo va en 240112900026 en haber.
  - las cuentas 64913101CF04, 64913101GA04 se usan para el total de extra de sueldos
  - la cuenta 240112900026 es el total de extras.
  """

  # Empecemos a definr las ctas contables que usaremos de acuerdo al catalogo de cuentas

  # Los sueldos van:


  CF_sueldo='64913101CF01'  # Centro de costos A,B,C Y D941  a D950

  GA_sueldo='64913101GA01'  # Centro de costos E

  GV_sueldo='64913101GV01'  # Centro de costo D981

  Total_sueldo= '240112900026'  # Sin centro de costo


  # Los extras van en

  CF_extra='64913101CF04'  # Centro de costos A,B,C Y D941  a D950

  GA_extra='64913101GA04'  # Centro de costos E

  GV_extra='64913101GV04'  # Centro de costo D981

  Total_extra= '240112900026'

  df_parapoliza=tabladin3.copy()

  df_parapoliza["cc"]=df_parapoliza["cc"].str.strip()

  df_parapoliza[["cc","unidad"]]=df_parapoliza["cc"].str.split(" 0 ",expand= True)



  df_parapoliza.drop(columns=["unidad"], inplace= True)

  df_parapoliza.rename(columns={'cc':'CC'}, inplace= True)

  df_parapoliza['CC']=df_parapoliza['CC'].str.strip()

  # Definimos nuestros CC con su Centro utilitario

  enlace= "https://docs.google.com/spreadsheets/d/15UP1JVbwgoljDTdNfn7BkFcrs7p1zFJs/export?format=xlsx"
  utilitario= pd.read_excel(enlace)
  df_parapoliza=pd.merge(df_parapoliza,utilitario, on=["CC"], how="left")

  df_parapoliza['CC']=df_parapoliza['CC'].replace("Total", "")

  """# Verificamos centros utilitarios"""

  datos_faltantes=df_parapoliza.query(' UTILITARIO.isnull()== True and CC!= ""')
  if (not datos_faltantes.empty):
    print( "Datos faltantesen CC :\n"+ str(datos_faltantes))
  else:
    print("Todo ok")

  datos_faltantes=datos_faltantes.reset_index()
  #datos_faltantes

  for i in range (len(datos_faltantes)):

    if datos_faltantes["CC"][i]!="E913":

      j=datos_faltantes["index"][i]

      uti=input("Introduzca el centro utilitario de "+str(datos_faltantes["CC"][i]))

      uti=uti.upper()

      df_parapoliza.loc[j,"UTILITARIO"]=uti

      S=pd.Series({"CC":datos_faltantes["CC"][i],"UTILITARIO": uti })

      utilitario=pd.concat([utilitario, S.to_frame().T], ignore_index=True)

  enlace = "https://docs.google.com/spreadsheets/d/1WAfIZJde_M2sYM7SIOyWCPvsMgJoyqVu/edit?usp=sharing&ouid=111113060171554295483&rtpof=true&sd=true"
  utilitario.to_excel(enlace, index= False)

  """# Generador de poliza y archivo final"""

  

  df_parapoliza_sueldo=df_parapoliza[['CC','UTILITARIO' ,'sueldo']].copy()
  df_parapoliza_extra=df_parapoliza[['CC','UTILITARIO' , 'Total de extras']].copy()


  df_parapoliza_sueldo.insert(0, 'plantilla', ['CAM' for i in range(len(df_parapoliza_sueldo))] )
  df_parapoliza_sueldo.insert(1, 'cons', [i+1 for i in range(len(df_parapoliza_sueldo))] )
  df_parapoliza_sueldo.insert(2, 'comp', ['100' for i in range(len(df_parapoliza_sueldo))] )


  cta=[]

  for i in range(len(df_parapoliza_sueldo)):

    CC= df_parapoliza_sueldo.loc[i, 'CC']

    if not  CC.startswith('E')  and  CC != "D981" and CC !="":
      cta.append(CF_sueldo)

    elif CC.startswith('E') :
      cta.append(GA_sueldo)

    elif CC=="D981" and CC !="" :
      cta.append(GV_sueldo)

    elif  CC =="" :

      cta.append(Total_sueldo)


  df_parapoliza_sueldo.insert(3, 'cta', cta)

  df_parapoliza_sueldo.insert(6, 'nada', [np.nan for i in range(len(df_parapoliza_sueldo))] )

  df_parapoliza_sueldo.insert(7, 'nada1', [np.nan for i in range(len(df_parapoliza_sueldo))] )

  df_parapoliza_sueldo.insert(8, 'debe/haber', [1 if df_parapoliza_sueldo.loc[i, 'cta'] .startswith('6') or df_parapoliza_sueldo.loc[i, 'cta'] .startswith('1')  else 2 for i in range(len(df_parapoliza_sueldo))] )



  df_parapoliza_sueldo["ref"]= [ 'SMASA ' + Referencia+ " -24 Sueldos"  for i in range(len(df_parapoliza_sueldo))]

  df_parapoliza_sueldo["CC"]=df_parapoliza_sueldo["CC"].replace("", np.nan)

  df_parapoliza_sueldo.rename(columns={'sueldo':'total'}, inplace= True)


  df_parapoliza_extra.insert(0, 'plantilla', ['CAM' for i in range(len(df_parapoliza_extra))] )
  df_parapoliza_extra.insert(1, 'cons', [i+1 for i in range(len(df_parapoliza_extra))] )
  df_parapoliza_extra.insert(2, 'comp', ['100' for i in range(len(df_parapoliza_extra))] )


  cta=[]

  for i in range(len(df_parapoliza_extra)):

    CC= df_parapoliza_extra.loc[i, 'CC']

    if not  CC.startswith('E')  and  CC != "D981" and CC !="":
      cta.append(CF_extra)

    elif CC.startswith('E') :
      cta.append(GA_extra)

    elif CC=="D981" and CC !="" :
      cta.append(GV_extra)

    elif  CC =="" :

      cta.append(Total_extra)


  df_parapoliza_extra.insert(3, 'cta', cta)

  df_parapoliza_extra.insert(6, 'nada', [np.nan for i in range(len(df_parapoliza_extra))] )

  df_parapoliza_extra.insert(7, 'nada1', [np.nan for i in range(len(df_parapoliza_extra))] )

  df_parapoliza_extra.insert(8, 'debe/haber', [1 if df_parapoliza_extra.loc[i, 'cta'] .startswith('6') or df_parapoliza_extra.loc[i, 'cta'] .startswith('1')  else 2 for i in range(len(df_parapoliza_extra))] )



  df_parapoliza_extra["ref"]= [ 'SMASA ' + Referencia+ " -24 Extras"  for i in range(len(df_parapoliza_extra))]

  df_parapoliza_extra["CC"]=df_parapoliza_extra["CC"].replace("", np.nan)

  df_parapoliza_extra.rename(columns={'Total de extras':'total'}, inplace= True)

  
  df_parapoliza = pd.concat([df_parapoliza_sueldo, df_parapoliza_extra], ignore_index=True).copy()
  df_parapoliza['cons']= [ i+1 for i in range(len(df_parapoliza))]

  dfsucio= pd.read_excel(archivo1_path)
  dfsucio2= pd.read_excel(archivo2_path)
  numero=meses.get(mes)

  # creamos un objeto de tipo BytesIO
  buffer_salida = BytesIO()

  # escribimos el excel en el buffer 
  with pd.ExcelWriter(buffer_salida, engine="xlsxwriter") as writer:
    dfsucio.to_excel(writer, sheet_name="Sheet1", index=False)
    dflimpio.to_excel(writer, sheet_name="ctas 5-6", index=False)
    df3.to_excel(writer, sheet_name="ctas x cc 6", index=False)
    tabladin.to_excel(writer, sheet_name="poliza", index=False)
    tabladin3.to_excel(writer, sheet_name="poliza", index=False, startcol=6)
    df_parapoliza.to_excel(writer, sheet_name="tfgld013", index=False, header=False)

  # reseteamos al inicio 
  buffer_salida.seek(0)

  return buffer_salida

  #en buffer_salida esta almacenado el Excel deseado. El nombre del archivo se lo encargamos a la aplicacion que haga uso
  #de la funcion, por ejemplo, al usar la siguiente funcion 
  #Guardarcomo=str(numero)+" "+ "costo empresa"+" "+mes+" "+"2024"+".xlsx"
  #with open(guardarcomo, "wb") as f:
  #   f.write(buffer_salida.getbuffer())




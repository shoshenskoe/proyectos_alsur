
import pandas as pd
import io
import xml.etree.ElementTree as ET

MES_LETRA = {
    '1':'ENERO','2':'FEBRERO','3':'MARZO','4':'ABRIL','5':'MAYO','6':'JUNIO',
    '7':'JULIO','8':'AGOSTO','9':'SEPTIEMBRE','10':'OCTUBRE','11':'NOVIEMBRE','12':'DICIEMBRE'
}

def procesar_trial(excel_path, mes, anio):
    mes = str(mes)
    anio = str(anio)

    m = mes[1] if mes.startswith('0') else mes
    m_int = int(m)

    tb = pd.read_excel(excel_path)
    tb.columns = tb.columns.str.replace(' ', '')

    if m_int == 1:
        tbm = tb[[MES_LETRA[str(m_int)], 'SALDOINICIAL', 'DEBE', 'HABER', 'SALDOFINAL']]
    else:
        tbm = tb[[MES_LETRA[str(m_int)], f'SALDOINICIAL.{m_int-1}', f'DEBE.{m_int-1}', f'HABER.{m_int-1}', f'SALDOFINAL.{m_int-1}']]

    tbm = tbm.set_axis(['CUENTA','SALDO INICIAL','DEBE','HABER','SALDO FINAL'], axis=1)

    # Excel "limpio"
    excel_balance_mes_buffer = io.BytesIO()
    tbm.to_excel(excel_balance_mes_buffer, index=False)
    excel_balance_mes_buffer.seek(0)

    # Columnas para XML
    bolsa = ['<BCE:Ctas NumCta="','"SaldoIni="','"Debe="','"Haber="','"SaldoFin="','"/>']
    for i, col in enumerate(bolsa):
        tbm.insert(loc=2*i, column=col, value=col)

    tbm.rename(columns={
        '<BCE:Ctas NumCta="':'','CUENTA':'NumCta','"SaldoIni="':'','SALDO INICIAL':'SaldoIni',
        '"Debe="':'','DEBE':'Debe','"Haber="':'','HABER':'Haber','"SaldoFin="':'','SALDO FINAL':'SaldoFin','"/>':''
    }, inplace=True)

    # Excel con columnas para XML
    trial_balanza_columnas_buffer = io.BytesIO()
    tbm.to_excel(trial_balanza_columnas_buffer, index=False)
    trial_balanza_columnas_buffer.seek(0)

    # String concatenado
    tbm['total'] = tbm.astype(str).agg(''.join, axis=1)
    string_concatenado_buffer = io.StringIO()
    tbm['total'].to_csv(string_concatenado_buffer, index=False, header=False)
    string_concatenado_buffer.seek(0)

    return excel_balance_mes_buffer, trial_balanza_columnas_buffer, string_concatenado_buffer


def convert_xlsx_to_xml(xlsx_buffer, anio, mes):
    anio = str(anio)
    mes = str(mes)

    BCE_NAMESPACE = "http://www.sat.gob.mx/esquemas/ContabilidadE/1_3/BalanzaComprobacion"
    ET.register_namespace("BCE", BCE_NAMESPACE)

    root_attributes = {
        "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
        "xsi:schemaLocation": "http://www.sat.gob.mx/esquemas/ContabilidadE/1_3/BalanzaComprobacion http://www.sat.gob.mx/esquemas/ContabilidadE/1_3/BalanzaComprobacion/BalanzaComprobacion_1_3.xsd",
        "Version": "1.3",
        "RFC": "ASC951228IM0",
        "Mes": mes,
        "Anio": anio,
        "TipoEnvio": "N",
    }

    root = ET.Element(f"{{{BCE_NAMESPACE}}}Balanza", attrib=root_attributes)

    xlsx_buffer.seek(0)
    df = pd.read_excel(xlsx_buffer)

    for _, row in df.iterrows():
        num_cta  = "" if pd.isna(row.get("NumCta")) else str(row["NumCta"])
        saldo_ini = "" if pd.isna(row.get("SaldoIni")) else f'{row["SaldoIni"]:.2f}'
        debe      = "" if pd.isna(row.get("Debe")) else f'{row["Debe"]:.2f}'
        haber     = "" if pd.isna(row.get("Haber")) else f'{row["Haber"]:.2f}'
        saldo_fin = "" if pd.isna(row.get("SaldoFin")) else f'{row["SaldoFin"]:.2f}'

        ET.SubElement(root, f"{{{BCE_NAMESPACE}}}Ctas", {
            "NumCta": num_cta, "SaldoIni": saldo_ini, "Debe": debe, "Haber": haber, "SaldoFin": saldo_fin
        })

    tree = ET.ElementTree(root)
    try:
        ET.indent(tree, space="  ", level=0)  # disponible en Python 3.9+
    except AttributeError:
        pass

    arbol_buffer = io.BytesIO()
    tree.write(arbol_buffer, encoding="utf-8", xml_declaration=True)
    arbol_buffer.seek(0)
    return arbol_buffer


def procesamiento_archivos(excel_path, anio, mes):
    excel_balance, excel_balance_colum, string_concatenado = procesar_trial(excel_path, mes, anio)
    arbol = convert_xlsx_to_xml(excel_balance_colum, anio, mes)

    mes_txt = MES_LETRA[str(int(mes))]
    anio = str(anio)

    nombre_balance_mes = f'BALANZA {mes_txt}  {anio}.xlsx'
    arbol_nombre = f'ASC951228IM0{anio}{mes_txt}BN.xml'
    nombre_string = f'Columna_concatenada{mes_txt} {anio}.txt'

    return (excel_balance, nombre_balance_mes,
            arbol, arbol_nombre,
            string_concatenado, nombre_string)

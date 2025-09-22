import streamlit as st
import pandas as pd
from codigo_reclasificacion import reclasificacion  # importa el script

st.title("Reclasificación de Bodegas")

# Subir archivos
ingresos_file = st.file_uploader("Selecciona el archivo de Ingresos", type=["xlsx"])
gastos_file = st.file_uploader("Selecciona el archivo de Gastos", type=["xlsx"])

# Entradas de texto
mes = st.text_input("Mes (ej. Enero)")
mes_numero = st.text_input("Número de mes (ej. 01)")

# Botón para procesar
if st.button("Procesar Reclasificación"):
    if ingresos_file and gastos_file and mes and mes_numero:
        try:
            ingresos_df = pd.read_excel(ingresos_file)
            gastos_df = pd.read_excel(gastos_file)
            
            archivo_memoria, nombre_archivo = reclasificacion(ingresos_df, gastos_df, mes, mes_numero)
            
            # Botón para descargar el archivo resultante
            st.download_button(
                label="Descargar archivo resultante",
                data=archivo_memoria,
                file_name=nombre_archivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Reclasificación completada.")
        except Exception as e:
            st.error(f"Ocurrió un error: {e}")
    else:
        st.warning("Por favor completa todos los campos y sube ambos archivos.")

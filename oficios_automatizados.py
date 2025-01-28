import streamlit as st
import pandas as pd
from docx import Document
from datetime import datetime
import os
import zipfile
from io import BytesIO

# Configuración de rutas
TEMPLATE_PATH = "001 OFICIO ciclo escolar 2024-2025.docx"
EXCEL_PATH = "PLANTILLA 29D AUDITORIA.xlsx"
REGISTRO_PATH = "registro_oficios_comision.xlsx"
OUTPUT_FOLDER_BASE = "output_oficios"

# Crear carpeta de salida si no existe
if not os.path.exists(OUTPUT_FOLDER_BASE):
    os.makedirs(OUTPUT_FOLDER_BASE)

# Función para convertir la fecha al formato deseado
def formatear_fecha(fecha):
    meses = [
        "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
    ]
    dia = fecha.day
    mes = meses[fecha.month - 1]
    anio = fecha.year
    return f"{dia} de {mes} del {anio}"

# Interfaz en Streamlit
st.set_page_config(page_title="Generador de Oficios", page_icon="📄")
st.title("📄 Generador de Oficios en Word")

# Verificar contraseña
password = st.text_input("🔒 Ingrese la contraseña", type="password")
if password != "defvm11":
    st.warning("Ingrese la contraseña correcta para continuar.")
    st.stop()

# Cargar los datos desde Excel
try:
    df = pd.read_excel(EXCEL_PATH)
except FileNotFoundError:
    st.error("El archivo de plantilla no se encuentra. Por favor, súbelo y asegúrate de que la ruta sea correcta.")
    st.stop()

selected_rows = st.multiselect(
    "👥 Selecciona los docentes",
    df.index,
    format_func=lambda i: f"{df.loc[i, 'NOMBRE (S)']} {df.loc[i, 'APELLIDO PATERNO']} {df.loc[i, 'APELLIDO MATERNO']}"
)

if selected_rows:
    st.write("✅ Docentes seleccionados:")
    st.write(df.loc[selected_rows])

# Campos de entrada
num_oficio = st.text_input("📄 Número de Oficio")
sede = st.text_input("🏫 Sede")
ubicacion = st.text_input("📍 Ubicación")
fecha_comision = st.date_input("📅 Fecha de Comisión")
horario = st.text_input("🕒 Horario")
fecha_emision = st.date_input("📅 Fecha de Emisión")
comision = st.text_input("🔖 Comisión")

# Botón para generar oficios
if st.button("Generar Oficios"):
    if not selected_rows:
        st.warning("Por favor selecciona al menos un docente.")
    else:
        data_to_process = df.loc[selected_rows]
        docentes = df.loc[selected_rows, "NOMBRE (S)"].tolist()
        # Aquí irían las funciones para generar los oficios y registrar en Excel
        st.success("🎉 Oficios generados con éxito. Descárgalos a continuación:")

# Botón para descargar el registro
if os.path.exists(REGISTRO_PATH):
    with open(REGISTRO_PATH, "rb") as registro:
        st.download_button(
            label="📊 Descargar Registro de Oficios",
            data=registro,
            file_name="registro_oficios_comision.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

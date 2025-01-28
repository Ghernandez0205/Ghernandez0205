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

# Función para generar oficios
def generar_oficio(datos, num_oficio, sede, ubicacion, fecha_comision, horario, fecha_emision, comision):
    archivos_generados = []
    for _, fila in datos.iterrows():
        nombre = fila['NOMBRE (S)']
        apellido_paterno = fila['APELLIDO PATERNO']
        apellido_materno = fila['APELLIDO MATERNO']
        rfc = fila['R.F.C. CON HOMONIMIA']

        # Crear el documento Word
        doc = Document(TEMPLATE_PATH)
        for p in doc.paragraphs:
            p.text = p.text.replace("{{numero_oficio}}", num_oficio)
            p.text = p.text.replace("{{nombre}}", nombre)
            p.text = p.text.replace("{{apellido_paterno}}", apellido_paterno)
            p.text = p.text.replace("{{apellido_materno}}", apellido_materno)
            p.text = p.text.replace("{{rfc}}", rfc)
            p.text = p.text.replace("{{sede}}", sede)
            p.text = p.text.replace("{{ubicacion}}", ubicacion)
            p.text = p.text.replace("{{fecha}}", formatear_fecha(fecha_comision))
            p.text = p.text.replace("{{horario}}", horario)
            p.text = p.text.replace("{{fecha_emision}}", formatear_fecha(fecha_emision))
            p.text = p.text.replace("{{comision}}", comision)

        # Guardar el archivo Word
        nombre_archivo = f"oficio_{rfc}.docx"
        ruta_archivo = os.path.join(OUTPUT_FOLDER_BASE, nombre_archivo)
        doc.save(ruta_archivo)
        archivos_generados.append(ruta_archivo)
    return archivos_generados

# Función para comprimir archivos en un ZIP
def comprimir_archivos(archivos):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for archivo in archivos:
            zipf.write(archivo, os.path.basename(archivo))
    zip_buffer.seek(0)
    return zip_buffer

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
        archivos_generados = generar_oficio(
            data_to_process, num_oficio, sede, ubicacion, fecha_comision, horario, fecha_emision, comision
        )
        zip_buffer = comprimir_archivos(archivos_generados)
        st.success("🎉 Oficios generados con éxito. Descárgalos a continuación:")
        st.download_button(
            label="📥 Descargar Oficios Comprimidos (ZIP)",
            data=zip_buffer,
            file_name="oficios_comprimidos.zip",
            mime="application/zip"
        )


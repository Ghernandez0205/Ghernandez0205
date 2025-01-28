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

# Función para generar archivos Word individuales
def generar_oficios(data, num_oficio, sede, ubicacion, fecha_comision, horario, fecha_emision, comision):
    docx_files = []
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    output_folder = os.path.join(OUTPUT_FOLDER_BASE, f'Oficios_{timestamp}')
    os.makedirs(output_folder)

    for index, row in data.iterrows():
        nombre = row["NOMBRE (S)"]
        apellido_paterno = row["APELLIDO PATERNO"]
        apellido_materno = row["APELLIDO MATERNO"]
        rfc = row["R.F.C. CON HOMONIMIA"]

        # Crear el documento Word desde la plantilla
        doc = Document(TEMPLATE_PATH)

        # Reemplazar texto en la plantilla
        for para in doc.paragraphs:
            para.text = para.text.replace("numero_oficio", num_oficio)
            para.text = para.text.replace("nombre", nombre)
            para.text = para.text.replace("apellido_paterno", apellido_paterno)
            para.text = para.text.replace("apellido_materno", apellido_materno)
            para.text = para.text.replace("rfc", rfc)
            para.text = para.text.replace("sede", sede)
            para.text = para.text.replace("ubicacion", ubicacion)
            para.text = para.text.replace("fecha", formatear_fecha(fecha_comision))
            para.text = para.text.replace("horario", horario)
            para.text = para.text.replace("fecha_de_emision", formatear_fecha(fecha_emision))
            para.text = para.text.replace("comision", comision)

        # Guardar el archivo Word
        output_docx = os.path.join(output_folder, f'oficio_{rfc}.docx')
        doc.save(output_docx)
        docx_files.append(output_docx)

    return docx_files

# Función para registrar los oficios generados en Excel
def registrar_oficios(fecha_comision, docentes, actividad, fecha_emision):
    if os.path.exists(REGISTRO_PATH):
        registro_df = pd.read_excel(REGISTRO_PATH)
    else:
        registro_df = pd.DataFrame(columns=["Fecha de Comisión", "Docentes", "Actividad", "Fecha de Emisión"])

    nueva_entrada = {
        "Fecha de Comisión": fecha_comision,
        "Docentes": ", ".join(docentes),
        "Actividad": actividad,
        "Fecha de Emisión": fecha_emision,
    }
    registro_df = pd.concat([registro_df, pd.DataFrame([nueva_entrada])], ignore_index=True)

    registro_df.to_excel(REGISTRO_PATH, index=False)

# Función para comprimir los archivos en un ZIP
def comprimir_archivos(archivos):
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w") as zipf:
        for archivo in archivos:
            zipf.write(archivo, os.path.basename(archivo))
    buffer.seek(0)
    return buffer

# Interfaz en Streamlit
st.set_page_config(page_title="Generador de Oficios", page_icon="📄")
st.title("📄 Generador de Oficios en Word")

# Agregar una imagen de encabezado
st.image("encabezado.png", use_column_width=True)

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
        docx_files = generar_oficios(
            data_to_process,
            num_oficio,
            sede,
            ubicacion,
            fecha_comision,
            horario,
            fecha_emision,
            comision,
        )
        registrar_oficios(
            formatear_fecha(fecha_comision),
            docentes,
            comision,
            formatear_fecha(fecha_emision),
        )
        zip_buffer = comprimir_archivos(docx_files)
        st.success("🎉 Oficios generados con éxito. Descárgalos a continuación:")
        st.download_button(
            label="📂 Descargar Oficios (ZIP)",
            data=zip_buffer,
            file_name="Oficios_Generados.zip",
            mime="application/zip",
        )

# Botón para descargar el registro
if os.path.exists(REGISTRO_PATH):
    with open(REGISTRO_PATH, "rb") as registro:
        st.download_button(
            label="📊 Descargar Registro de Oficios",
            data=registro,
            file_name="registro_oficios_comision.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

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
OUTPUT_FOLDER_BASE = "output_oficios"

# Crear carpeta de salida si no existe
if not os.path.exists(OUTPUT_FOLDER_BASE):
    os.makedirs(OUTPUT_FOLDER_BASE)

# Función para generar archivos Word individuales y convertirlos a PDF
def generar_oficios_y_convertir(data, num_oficio, sede, ubicacion, fecha, horario, fecha_emision, comision):
    pdf_files = []
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
            para.text = para.text.replace("fecha", fecha)
            para.text = para.text.replace("horario", horario)
            para.text = para.text.replace("fecha_emision", fecha_emision)
            para.text = para.text.replace("comision", comision)

        # Guardar el archivo Word
        output_docx = os.path.join(output_folder, f'oficio_{rfc}.docx')
        doc.save(output_docx)

        # Convertir Word a PDF usando LibreOffice
        output_pdf = output_docx.replace(".docx", ".pdf")
        os.system(f'libreoffice --headless --convert-to pdf "{output_docx}" --outdir "{output_folder}"')
        pdf_files.append(output_pdf)

    return pdf_files

# Función para comprimir los archivos en un ZIP
def comprimir_archivos(archivos):
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w") as zipf:
        for archivo in archivos:
            zipf.write(archivo, os.path.basename(archivo))
    buffer.seek(0)
    return buffer

# Interfaz en Streamlit
st.title("Generador de Oficios en PDF con Imágenes Preservadas")

# Verificar contraseña
password = st.text_input("Ingrese la contraseña", type="password")
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
    "Selecciona los docentes",
    df.index,
    format_func=lambda i: f"{df.loc[i, 'NOMBRE (S)']} {df.loc[i, 'APELLIDO PATERNO']} {df.loc[i, 'APELLIDO MATERNO']}"
)

if selected_rows:
    st.write("Docentes seleccionados:")
    st.write(df.loc[selected_rows])

# Campos de entrada
num_oficio = st.text_input("Número de Oficio")
sede = st.text_input("Sede")
ubicacion = st.text_input("Ubicación")
fecha = st.date_input("Fecha de Comisión")
horario = st.text_input("Horario")
fecha_emision = st.text_input("Fecha de Emisión (por ejemplo: 15 de enero del 2025)", placeholder="15 de enero del 2025")
comision = st.text_input("Comisión")

# Botón para generar oficios en PDF
if st.button("Generar Oficios en PDF"):
    if not selected_rows:
        st.warning("Por favor selecciona al menos un docente.")
    elif not fecha_emision.strip():
        st.warning("Por favor ingresa una fecha de emisión válida.")
    else:
        data_to_process = df.loc[selected_rows]
        pdf_files = generar_oficios_y_convertir(
            data_to_process,
            num_oficio,
            sede,
            ubicacion,
            fecha.strftime("%d/%m/%Y"),
            horario,
            fecha_emision.strip(),
            comision,
        )
        zip_buffer = comprimir_archivos(pdf_files)
        st.success("Oficios generados con éxito. Descárgalos a continuación:")
        st.download_button(
            label="Descargar Oficios (ZIP)",
            data=zip_buffer,
            file_name="Oficios_Generados.zip",
            mime="application/zip",
        )

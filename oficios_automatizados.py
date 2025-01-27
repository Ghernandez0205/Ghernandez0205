import streamlit as st
import pandas as pd
from docx import Document
from pypandoc import convert
from PyPDF2 import PdfMerger
import os
from datetime import datetime

# Ruta a la plantilla y al archivo Excel
TEMPLATE_PATH = "001 OFICIO ciclo escolar 2024-2025.docx"
EXCEL_PATH = "PLANTILLA 29D AUDITORIA.xlsx"
OUTPUT_FOLDER_BASE = "output_oficios"

# Crear carpeta de salida si no existe
if not os.path.exists(OUTPUT_FOLDER_BASE):
    os.makedirs(OUTPUT_FOLDER_BASE)

# Función para generar los oficios
def generar_oficio(data, num_oficio, sede, ubicacion, fecha, horario, fecha_emision, comision):
    pdf_files = []
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    output_folder = os.path.join(OUTPUT_FOLDER_BASE, f'Oficios_{timestamp}')
    os.makedirs(output_folder)
    
    for index, row in data.iterrows():
        nombre = row["NOMBRE (S)"]
        apellido_paterno = row["APELLIDO PATERNO"]
        apellido_materno = row["APELLIDO MATERNO"]
        rfc = row["R.F.C. CON HOMONIMIA"]

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

        # Guardar archivo .docx y convertir a PDF
        output_docx = os.path.join(output_folder, f'oficio_{rfc}.docx')
        doc.save(output_docx)
        convert(output_docx)

        output_pdf = output_docx.replace(".docx", ".pdf")
        pdf_files.append(output_pdf)

    # Combinar PDFs
    merger = PdfMerger()
    for pdf in pdf_files:
        merger.append(pdf)
    combined_pdf_path = os.path.join(output_folder, f"Oficios_Combinados_{timestamp}.pdf")
    merger.write(combined_pdf_path)
    merger.close()

    return combined_pdf_path

# Interfaz en Streamlit
st.title("Generador de Oficios Automatizados")

# Verificar contraseña
password = st.text_input("Ingrese la contraseña", type="password")
if password != "defvm11":
    st.warning("Ingrese la contraseña correcta para continuar.")
    st.stop()

# Cargar los datos desde Excel
df = pd.read_excel(EXCEL_PATH)
selected_rows = st.multiselect("Selecciona los docentes", df.index, format_func=lambda i: f"{df.loc[i, 'NOMBRE (S)']} {df.loc[i, 'APELLIDO PATERNO']} {df.loc[i, 'APELLIDO MATERNO']}")

if selected_rows:
    st.write("Docentes seleccionados:")
    st.write(df.loc[selected_rows])

# Campos de entrada
num_oficio = st.text_input("Número de Oficio")
sede = st.text_input("Sede")
ubicacion = st.text_input("Ubicación")
fecha = st.date_input("Fecha de Comisión")
horario = st.text_input("Horario")
fecha_emision = st.date_input("Fecha de Emisión")
comision = st.text_input("Comisión")

# Botón para generar oficios
if st.button("Generar Oficios"):
    if not selected_rows:
        st.warning("Por favor selecciona al menos un docente.")
    else:
        data_to_process = df.loc[selected_rows]
        result_pdf = generar_oficio(
            data_to_process,
            num_oficio,
            sede,
            ubicacion,
            fecha.strftime("%d/%m/%Y"),
            horario,
            fecha_emision.strftime("%d/%m/%Y"),
            comision,
        )
        st.success("Oficios generados con éxito. Descárgalos a continuación:")
        st.download_button("Descargar Oficios Combinados", open(result_pdf, "rb"), file_name="Oficios_Combinados.pdf")

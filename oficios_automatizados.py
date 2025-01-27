import streamlit as st
import pandas as pd
from docx import Document
from datetime import datetime
import os
from io import BytesIO

# Configuración de rutas
TEMPLATE_PATH = "001 OFICIO ciclo escolar 2024-2025.docx"
EXCEL_PATH = "PLANTILLA 29D AUDITORIA.xlsx"
OUTPUT_FOLDER_BASE = "output_oficios"

# Crear carpeta de salida si no existe
if not os.path.exists(OUTPUT_FOLDER_BASE):
    os.makedirs(OUTPUT_FOLDER_BASE)

# Función para generar un archivo Word combinado
def generar_oficio_combinado(data, num_oficio, sede, ubicacion, fecha, horario, fecha_emision, comision):
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    combined_doc = Document()

    for index, row in data.iterrows():
        nombre = row["NOMBRE (S)"]
        apellido_paterno = row["APELLIDO PATERNO"]
        apellido_materno = row["APELLIDO MATERNO"]
        rfc = row["R.F.C. CON HOMONIMIA"]

        # Cargar la plantilla
        template_doc = Document(TEMPLATE_PATH)

        # Reemplazar texto en la plantilla
        for para in template_doc.paragraphs:
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

        # Agregar el contenido del archivo al documento combinado
        for element in template_doc.element.body:
            combined_doc.element.body.append(element)

        # Agregar una página en blanco entre oficios
        combined_doc.add_page_break()

    # Guardar el archivo Word combinado en memoria
    output = BytesIO()
    combined_doc.save(output)
    output.seek(0)
    return output

# Interfaz en Streamlit
st.title("Generador de Oficios Combinados en Word")

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
fecha_emision = st.date_input("Fecha de Emisión")
comision = st.text_input("Comisión")

# Botón para generar oficios combinados
if st.button("Generar Oficios Combinados"):
    if not selected_rows:
        st.warning("Por favor selecciona al menos un docente.")
    else:
        data_to_process = df.loc[selected_rows]
        combined_doc = generar_oficio_combinado(
            data_to_process,
            num_oficio,
            sede,
            ubicacion,
            fecha.strftime("%d/%m/%Y"),
            horario,
            fecha_emision.strftime("%d/%m/%Y"),
            comision,
        )
        st.success("Oficios combinados generados con éxito. Descárgalos a continuación:")
        st.download_button(
            label="Descargar Oficios Combinados (Word)",
            data=combined_doc,
            file_name="Oficios_Combinados.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

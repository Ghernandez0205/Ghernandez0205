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
HISTORIAL_PATH = "historial_oficios.xlsx"
OUTPUT_FOLDER_BASE = "output_oficios"

# Crear carpeta de salida si no existe
if not os.path.exists(OUTPUT_FOLDER_BASE):
    os.makedirs(OUTPUT_FOLDER_BASE)

# Función para generar oficios
def generar_oficio(data, num_oficio, sede, ubicacion, fecha_comision, horario, mes_emision, comision):
    archivos_generados = []
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    output_folder = os.path.join(OUTPUT_FOLDER_BASE, f'Oficios_{timestamp}')
    os.makedirs(output_folder)

    for _, fila in data.iterrows():
        nombre = fila['NOMBRE (S)']
        apellido_paterno = fila['APELLIDO PATERNO']
        apellido_materno = fila['APELLIDO MATERNO']
        rfc = fila['R.F.C. CON HOMONIMIA']

        # Crear el documento Word desde la plantilla
        doc = Document(TEMPLATE_PATH)

        # Reemplazar texto en la plantilla
        for p in doc.paragraphs:
            p.text = p.text.replace("mes", mes_emision)
            p.text = p.text.replace("fecha", fecha_comision.strftime('%d de %B del %Y'))
            p.text = p.text.replace("numero_oficio", num_oficio)
            p.text = p.text.replace("nombre", nombre)
            p.text = p.text.replace("apellido_paterno", apellido_paterno)
            p.text = p.text.replace("apellido_materno", apellido_materno)
            p.text = p.text.replace("rfc", rfc)
            p.text = p.text.replace("sede", sede)
            p.text = p.text.replace("ubicacion", ubicacion)
            p.text = p.text.replace("horario", horario)
            p.text = p.text.replace("comision", comision)

        # Guardar el archivo Word
        nombre_archivo = f"oficio_{rfc}.docx"
        ruta_archivo = os.path.join(output_folder, nombre_archivo)
        doc.save(ruta_archivo)
        archivos_generados.append(ruta_archivo)
    return archivos_generados

# Función para comprimir los archivos en un ZIP
def comprimir_archivos(archivos):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for archivo in archivos:
            zipf.write(archivo, os.path.basename(archivo))
    zip_buffer.seek(0)
    return zip_buffer

# Función para actualizar el historial de oficios
def actualizar_historial(data, num_oficio, comision):
    historial_df = pd.DataFrame()
    if os.path.exists(HISTORIAL_PATH):
        historial_df = pd.read_excel(HISTORIAL_PATH)
    
    nuevo_historial = pd.DataFrame({
        "Número Consecutivo": [len(historial_df) + i + 1 for i in range(len(data))],
        "Nombre": data['NOMBRE (S)'].values,
        "Apellido Paterno": data['APELLIDO PATERNO'].values,
        "Apellido Materno": data['APELLIDO MATERNO'].values,
        "Número de Oficio": [num_oficio] * len(data),
        "Actividad": [comision] * len(data)
    })
    
    historial_df = pd.concat([historial_df, nuevo_historial], ignore_index=True)
    historial_df.to_excel(HISTORIAL_PATH, index=False)

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
mes_emision = st.text_input("📆 Mes de Emisión")
horario = st.text_input("🕒 Horario")
comision = st.text_input("🔖 Comisión")

# Botón para generar oficios
if st.button("Generar Oficios"):
    if not selected_rows:
        st.warning("Por favor selecciona al menos un docente.")
    else:
        data_to_process = df.loc[selected_rows]
        archivos_generados = generar_oficio(
            data_to_process, num_oficio, sede, ubicacion, fecha_comision, horario, mes_emision, comision
        )
        actualizar_historial(data_to_process, num_oficio, comision)
        zip_buffer = comprimir_archivos(archivos_generados)
        st.success("🎉 Oficios generados con éxito. Descárgalos a continuación:")
        st.download_button(
            label="📥 Descargar Oficios Comprimidos (ZIP)",
            data=zip_buffer,
            file_name="oficios_comprimidos.zip",
            mime="application/zip"
        )
        st.download_button(
            label="📥 Descargar Historial de Oficios",
            data=open(HISTORIAL_PATH, "rb"),
            file_name="historial_oficios.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

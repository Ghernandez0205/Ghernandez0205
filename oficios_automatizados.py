
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
IMAGE_PATH = r"C:\Users\sup11\Ghernandez0205\genshn impact.png"  # Ruta completa de la imagen

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

# Cargar la imagen desde la ruta completa
st.image(IMAGE_PATH, use_column_width=True)

# Verificar contraseña
password = st.text_input("🔒 Ingrese la contraseña", type="password")
if password != "defvm11":
    st.warning("Ingrese la contraseña correcta para continuar.")
    st.stop()

# Resto del código (como ya corregido anteriormente)

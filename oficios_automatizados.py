import os
import pandas as pd
from docx import Document
from tkinter import *
from tkinter import ttk, messagebox
from datetime import datetime
from docx2pdf import convert
from PyPDF2 import PdfMerger

# Ruta a los archivos
template_path = r"C:\Users\sup11\OneDrive\Attachments\Documentos\Interfaces de phyton\oficios automatizados Lap\001 OFICIO ciclo escolar 2024-2025.docx"
excel_path = r"C:\Users\sup11\OneDrive\Attachments\Documentos\Interfaces de phyton\oficios automatizados Lap\PLANTILLA 29D AUDITORIA.xlsx"
output_folder_base = r"C:\Users\sup11\OneDrive\Attachments\Documentos\Interfaces de phyton\oficios automatizados Lap\Oficios"

# Verificar que los archivos existen
if not os.path.exists(template_path):
    raise FileNotFoundError(f"Plantilla no encontrada: {template_path}")

if not os.path.exists(excel_path):
    raise FileNotFoundError(f"Archivo Excel no encontrado: {excel_path}")

# Crear carpeta de salida si no existe
if not os.path.exists(output_folder_base):
    os.makedirs(output_folder_base)

# Cargar los datos desde Excel
df = pd.read_excel(excel_path)
df_filtered = df[['NOMBRE (S)', 'APELLIDO PATERNO', 'APELLIDO MATERNO', 'R.F.C. CON HOMONIMIA']]

# Función para generar los oficios y combinarlos en un PDF
def generar_oficio():
    try:
        # Capturar datos desde la interfaz
        num_oficio = entry_oficio.get()
        sede = entry_sede.get()
        ubicacion = entry_ubicacion.get()
        fecha = entry_fecha.get()
        horario = entry_horario.get()
        fecha_emision = entry_fecha_emision.get()
        comision = entry_comision.get()

        # Verificar que todos los campos estén completos
        if not (num_oficio and sede and ubicacion and fecha and horario and fecha_emision and comision):
            messagebox.showwarning("Advertencia", "Por favor, complete todos los campos.")
            return

        # Verificar que al menos un docente esté seleccionado
        selected_items = tree.selection()
        if not selected_items:
            messagebox.showwarning("Advertencia", "Por favor, seleccione al menos un docente.")
            return

        # Crear una carpeta única para guardar los oficios generados
        timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        output_folder = os.path.join(output_folder_base, f'Oficios_{timestamp}')
        os.makedirs(output_folder)

        # Lista para almacenar los PDFs generados
        pdf_files = []

        # Crear un DataFrame para guardar los datos de los docentes seleccionados
        selected_data = []

        # Generar un oficio para cada docente seleccionado
        for item in selected_items:
            docente = tree.item(item, 'values')
            nombre = docente[0]
            apellido_paterno = docente[1]
            apellido_materno = docente[2]
            rfc = docente[3]

            # Cargar el documento Word
            doc = Document(template_path)

            # Reemplazar manualmente en el documento
            for para in doc.paragraphs:
                para.text = para.text.replace('numero_oficio', num_oficio)
                para.text = para.text.replace('nombre', nombre)
                para.text = para.text.replace('apellido_paterno', apellido_paterno)
                para.text = para.text.replace('apellido_materno', apellido_materno)
                para.text = para.text.replace('rfc', rfc)
                para.text = para.text.replace('sede', sede)
                para.text = para.text.replace('ubicacion', ubicacion)
                para.text = para.text.replace('fecha.', fecha_emision)  # Fecha de emisión para encabezado
                para.text = para.text.replace('fecha', fecha)  # Fecha de la comisión
                para.text = para.text.replace('horario', horario)
                para.text = para.text.replace('comision', comision)

            # Guardar el documento en formato .docx
            output_docx = os.path.join(output_folder, f'oficio_{rfc}.docx')
            doc.save(output_docx)

            # Convertir el .docx a PDF
            convert(output_docx)

            # Obtener la ruta del archivo PDF generado
            output_pdf = output_docx.replace('.docx', '.pdf')
            pdf_files.append(output_pdf)

            # Agregar datos al DataFrame
            selected_data.append({
                'Nombre': nombre,
                'Apellido Paterno': apellido_paterno,
                'Apellido Materno': apellido_materno,
                'RFC': rfc,
                'Comisión': comision,
                'No. de Oficio': num_oficio
            })

        # Combinar todos los PDFs en un solo archivo
        merger = PdfMerger()
        for pdf in pdf_files:
            merger.append(pdf)

        # Guardar el PDF combinado
        combined_pdf_path = os.path.join(output_folder, f'Oficios_Combinados_{timestamp}.pdf')
        merger.write(combined_pdf_path)
        merger.close()

        # Guardar los datos seleccionados en un archivo Excel
        excel_output_path = os.path.join(output_folder, f'Datos_Seleccionados_{timestamp}.xlsx')
        pd.DataFrame(selected_data).to_excel(excel_output_path, index=False)

        messagebox.showinfo("Éxito", f"Oficios generados y combinados en {combined_pdf_path}\n"
                                     f"Datos guardados en {excel_output_path}")

    except Exception as e:
        messagebox.showerror("Error", f"Error durante la generación de los oficios: {str(e)}")

# Interfaz gráfica con tkinter
root = Tk()
root.title("Generación de Oficios - Selección de Docentes")

# Crear la tabla para mostrar los datos del Excel
tree = ttk.Treeview(root, columns=('Nombre', 'Apellido Paterno', 'Apellido Materno', 'RFC'), show='headings')
tree.heading('Nombre', text='Nombre')
tree.heading('Apellido Paterno', text='Apellido Paterno')
tree.heading('Apellido Materno', text='Apellido Materno')
tree.heading('RFC', text='RFC')

# Insertar los datos en la tabla desde el archivo Excel
for index, row in df_filtered.iterrows():
    tree.insert('', END, values=(row['NOMBRE (S)'], row['APELLIDO PATERNO'], row['APELLIDO MATERNO'], row['R.F.C. CON HOMONIMIA']))

# Posicionar la tabla en la ventana
tree.pack(fill=BOTH, expand=True)

# Añadir campos de entrada para los datos adicionales
frame = Frame(root)
frame.pack(pady=20)

Label(frame, text="No. de Oficio:").grid(row=0, column=0, padx=5, pady=5)
entry_oficio = Entry(frame)
entry_oficio.grid(row=0, column=1, padx=5, pady=5)

Label(frame, text="Sede:").grid(row=1, column=0, padx=5, pady=5)
entry_sede = Entry(frame)
entry_sede.grid(row=1, column=1, padx=5, pady=5)

Label(frame, text="Ubicación:").grid(row=2, column=0, padx=5, pady=5)
entry_ubicacion = Entry(frame)
entry_ubicacion.grid(row=2, column=1, padx=5, pady=5)

Label(frame, text="Fecha:").grid(row=3, column=0, padx=5, pady=5)
entry_fecha = Entry(frame)
entry_fecha.grid(row=3, column=1, padx=5, pady=5)

Label(frame, text="Horario:").grid(row=4, column=0, padx=5, pady=5)
entry_horario = Entry(frame)
entry_horario.grid(row=4, column=1, padx=5, pady=5)

Label(frame, text="Fecha de Emisión:").grid(row=5, column=0, padx=5, pady=5)
entry_fecha_emision = Entry(frame)
entry_fecha_emision.grid(row=5, column=1, padx=5, pady=5)

Label(frame, text="Comisión:").grid(row=6, column=0, padx=5, pady=5)
entry_comision = Entry(frame)
entry_comision.grid(row=6, column=1, padx=5, pady=5)

# Botón para generar el oficio
Button(frame, text="Generar Oficios", command=generar_oficio).grid(row=7, columnspan=2, pady=10)

# Ejecutar la interfaz gráfica
root.mainloop()

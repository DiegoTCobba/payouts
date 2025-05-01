import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re

st.title("📄🔍 Filtrado y resaltado de Excel desde PDF")

# Subir archivos
pdf_file = st.file_uploader("Sube un archivo PDF", type="pdf")
excel_file = st.file_uploader("Sube un archivo Excel", type=["xlsx"])

if pdf_file and excel_file:
    # Extraer texto del PDF
    text = ""
    with fitz.open(stream=pdf_file.read(), filetype="pdf") as doc:
        for page in doc:
            text += page.get_text()

    # Buscar números de documento (ajustar expresión según tu formato)
    numeros_documento = re.findall(r'\d{6,}', text)
    numeros_documento = list(set(numeros_documento))  # eliminar duplicados

    st.success(f"Números de documento encontrados: {len(numeros_documento)}")
    st.write(numeros_documento)

    # Leer Excel con pandas (previa visualización)
    df = pd.read_excel(excel_file)
    st.subheader("Vista previa del Excel:")
    st.dataframe(df)

    # Reposicionar puntero para openpyxl
    output = io.BytesIO()
    excel_file.seek(0)
    wb = load_workbook(excel_file)
    ws = wb.active

    # Estilo de resaltado amarillo
    fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Resaltar coincidencias
    for row in ws.iter_rows():
        for cell in row:
            if str(cell.value) in numeros_documento:
                cell.fill = fill

    # Guardar archivo en memoria
    wb.save(output)
    output.seek(0)

    # Botón de descarga
    st.download_button(
        label="📥 Descargar Excel con resaltado",
        data=output,
        file_name="resaltado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

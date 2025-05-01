import streamlit as st
import pandas as pd
import fitz
import io


st.title("Filtrado y resaltado de Excel desde PDF")

# Subir archivos
pdf_file = st.file_uploader("Sube un archivo PDF", type="pdf")
excel_file = st.file_uploader("Sube un archivo Excel", type=["xlsx"])

if pdf_file and excel_file:
    # Extraer texto del PDF
     with fitz.open(stream=pdf_file.read(), filetype="pdf") as doc:
        text = ""
        for page in doc:
            text += page.get_text()

    # Extraer los números de documento
    import re
    numeros_documento = re.findall(r"\d{6,9}\d", text)
    numeros_documento = list(set(numeros_documento))  # eliminar duplicados

    st.write("Números de documento encontrados:", numeros_documento)

    # Leer el Excel
    df = pd.read_excel(excel_file)
    st.write("Vista previa del Excel:")
    st.dataframe(df)

    # Buscar en todas las columnas si algún valor coincide
    output = io.BytesIO()
    excel_file.seek(0)
    wb = load_workbook(excel_file)
    ws = wb.active

    # Estilo de resaltado
    fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Resaltar celdas con coincidencias
    for row in ws.iter_rows():
        for cell in row:
            if str(cell.value) in numeros_documento:
                cell.fill = fill

    # Guardar a BytesIO y permitir descarga
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 Descargar Excel resaltado",
        data=output,
        file_name="resaltado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

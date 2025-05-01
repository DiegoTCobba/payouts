import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re

st.title("📄🔍 Filtrado y resaltado de clientes a rechazar")

# Subir archivos
pdf_file = st.file_uploader("Sube un archivo PDF", type="pdf")
excel_file = st.file_uploader("Sube un archivo Excel", type=["xlsx"])

if pdf_file and excel_file:
    # Extraer texto del PDF
    text = ""
    with fitz.open(stream=pdf_file.read(), filetype="pdf") as doc:
        for page in doc:
            text += page.get_text()

    # Eliminar posibles números de cuenta como 380-75148297-0-31
    cuentas_posibles = re.findall(r'\b[\d]{2,4}-[\d]{5,10}-\d{1,2}-\d{1,3}\b', text)
    for cuenta in cuentas_posibles:
        text = text.replace(cuenta, '')

    # Buscar números de documento puros de 6 o más dígitos (sin símbolos ni guiones)
    numeros_documento = re.findall(r'\b\d{6,}\b', text)
    numeros_documento = list(set(numeros_documento))  # eliminar duplicados

    # Leer Excel con pandas (previa visualización)
    df = pd.read_excel(excel_file)

    # Reposicionar puntero para openpyxl
    output = io.BytesIO()
    excel_file.seek(0)
    wb = load_workbook(excel_file)
    ws = wb.active

    # Estilo de resaltado amarillo
    fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Ocultar columnas específicas en el archivo Excel
    columnas_a_ocultar = ['B', 'C', 'E', 'F', 'G', 'H', 'J', 'K', 'L', 'N', 'O', 'P', 'R']
    for col in columnas_a_ocultar:
        ws.column_dimensions[col].hidden = True

    # Resaltar coincidencias
    for row in ws.iter_rows():
        for cell in row:
            if str(cell.value) in numeros_documento:
                cell.fill = fill

    # Eliminar filas cuya celda en la columna A no sea amarilla
    filas_a_eliminar = []
    for row_idx, row in enumerate(ws.iter_rows(), start=1):  # Iterar con índice de fila
        # Verificar si la celda de la columna A tiene fondo amarillo
        if row[0].fill != fill:  # Si la celda de la columna A no es amarilla
            filas_a_eliminar.append(row_idx)

    # Eliminar filas no amarillas en la columna A
    for idx in reversed(filas_a_eliminar):  # Reverso para evitar problemas al eliminar
        ws.delete_rows(idx)

    # Guardar archivo en memoria
    wb.save(output)
    output.seek(0)

    # Convertir hoja activa a DataFrame (después de resaltar)
    data = ws.values
    columns = next(data)
    df_resaltado = pd.DataFrame(data, columns=columns)

    # Limpiar el DataFrame, convirtiendo todos los valores a cadenas
    df_visible_cleaned = df_resaltado.applymap(str)  # Convertir todos los valores a cadenas

    # Ocultar columnas específicas por letra
    columnas_a_ocultar = ['B', 'C', 'E', 'F', 'G', 'H', 'J', 'K', 'L', 'N', 'O', 'P', 'R']
    # Convertir letras a índices (0-based)
    letras_a_indices = [ord(c) - ord('A') for c in columnas_a_ocultar]

    # Eliminar del DataFrame
    columnas_visibles_

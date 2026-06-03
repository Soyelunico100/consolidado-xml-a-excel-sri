from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path

import streamlit as st

import consolidado_xml_a_excel as converter


BASE_DIR = Path(__file__).resolve().parent


def count_files(folder, pattern):
    return len(list(folder.rglob(pattern))) if folder.exists() else 0


st.set_page_config(page_title="XML SRI a Excel", layout="centered")

st.title("XML SRI a Excel")
st.write("Coloca tus XML en `entrada_xml` y pulsa procesar.")

cols = st.columns(4)
cols[0].metric("XML entrada", count_files(converter.XML_FOLDER, "*.xml"))
cols[1].metric("Excel salida", count_files(converter.SALIDA_EXCEL_DIR, "*.xlsx"))
cols[2].metric("Procesados", count_files(converter.PROCESADOS_DIR, "*.xml"))
cols[3].metric("Errores", count_files(converter.ERRORES_DIR, "*.xml"))

st.code(f"Entrada XML: {converter.XML_FOLDER}\nSalida Excel: {converter.SALIDA_EXCEL_DIR}", language="text")

mover_xml = st.checkbox("Mover XML a procesados/errores al terminar", value=True)

if st.button("Procesar XML", type="primary"):
    converter.MOVER_XML_PROCESADOS = mover_xml
    output = StringIO()
    try:
        with redirect_stdout(output):
            converter.main()
    except Exception as exc:
        st.error(f"Ocurrio un error: {exc}")
        st.code(output.getvalue(), language="text")
    else:
        st.success("Proceso terminado.")
        st.code(output.getvalue(), language="text")

st.divider()
st.write("Los XML reales y Excel generados quedan ignorados por `.gitignore` para no subir datos sensibles a GitHub.")

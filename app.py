import streamlit as st
import pandas as pd
from datetime import datetime
import io

from scripts.reporte_general import procesar_reporte_general
from scripts.script2 import procesar_script2

# Logo y título
col1, col2 = st.columns([1, 5])
with col1:
    st.image("mi_logo.png", use_container_width=True)
with col2:
    st.title("Generador de Reportes Automático")

# --- SUBIR ARCHIVO ---
st.subheader("📥 Subí tu archivo CSV")

if "uploaded_file" not in st.session_state:
    st.session_state.uploaded_file = None

uploaded = st.file_uploader("Seleccioná el archivo", type="csv", label_visibility="collapsed")

if uploaded:
    st.session_state.uploaded_file = uploaded

# Mostrar archivo subido con botón para eliminar
if st.session_state.uploaded_file:
    with st.container():
        st.markdown(f"**Archivo cargado:** `{st.session_state.uploaded_file.name}`", unsafe_allow_html=True)
        if st.button("❌ Eliminar archivo", key="delete_file"):
            st.session_state.uploaded_file = None
            st.experimental_rerun()

# --- ELEGIR SCRIPT ---
st.subheader("⚙️ Elegí el tipo de procesamiento")
opcion = st.selectbox("Seleccioná el script a ejecutar:", ["", "Reporte General", "Script 2"])

# --- EJECUTAR SCRIPT ---
if st.session_state.uploaded_file and opcion:
    if st.button("▶️ Ejecutar script"):
        df = pd.read_csv(st.session_state.uploaded_file)

        nombre_reporte = f"reporte_general_{datetime.today().strftime('%Y-%m-%d')}.xlsx"
        output = io.BytesIO()

        if opcion == "Reporte General":
            procesar_reporte_general(df, output)

        elif opcion == "Script 2":
            procesar_script2(df, output)

        output.seek(0)
        st.success("✅ Reporte generado con éxito")
        st.download_button(
            label="📥 Descargar reporte",
            data=output,
            file_name=nombre_reporte,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

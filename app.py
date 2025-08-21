import io
from datetime import datetime

import pandas as pd
import streamlit as st

from scripts.reporte_general import procesar_reporte_general
from scripts.script2 import procesar_script2

# ---------- Config ----------
st.set_page_config(
    page_title="Generador de Reportes Automático",
    page_icon="📊",
    layout="wide",
)

# CSS externo
try:
    with open("assets/styles.css", "r", encoding="utf-8") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
except FileNotFoundError:
    pass  # si no tenés aún la carpeta assets/styles.css, no rompe

# ---------- Header ----------
col1, col2 = st.columns([1, 5])
with col1:
    st.image("mi_logo.png", use_container_width=True)
with col2:
    st.markdown("<h1 style='margin-bottom:0.2rem'>Generador de Reportes Automático</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:#475569;margin-top:0'>Subí tu CSV, elegí el script y descargá el Excel listo.</p>", unsafe_allow_html=True)

# ---------- Subir archivo ----------
st.subheader("📥 Subí tu archivo CSV")

if "uploaded_file" not in st.session_state:
    st.session_state.uploaded_file = None

uploaded = st.file_uploader("Seleccioná el archivo", type="csv", label_visibility="collapsed")

if uploaded:
    st.session_state.uploaded_file = uploaded

# Mostrar archivo subido con botón para eliminar
if st.session_state.uploaded_file:
    with st.container():
        st.markdown(f"**Archivo cargado:** `{st.session_state.uploaded_file.name}`")
        if st.button("❌ Eliminar archivo", key="delete_file"):
            st.session_state.uploaded_file = None
            st.experimental_rerun()

# ---------- Elegir script ----------
st.subheader("⚙️ Elegí el tipo de procesamiento")
opcion = st.selectbox(
    "Seleccioná el script a ejecutar:",
    ["", "Reporte General", "Script 2"],
    index=0
)

# ---------- Ejecutar script ----------
if st.session_state.uploaded_file and opcion and opcion != "":
    if st.button("▶️ Ejecutar script", type="primary", use_container_width=True):
        df = pd.read_csv(st.session_state.uploaded_file)

        nombre_reporte = f"reporte_general_{datetime.today().strftime('%Y-%m-%d')}.xlsx"
        output = io.BytesIO()

        progress = st.progress(0)
        st.toast("Procesando…")

        # Procesar según opción
        if opcion == "Reporte General":
            procesar_reporte_general(df, output)
        elif opcion == "Script 2":
            procesar_script2(df, output)

        progress.progress(85)
        output.seek(0)
        progress.progress(100)

        st.markdown('<div class="success-card" style="margin-top:0.75rem;">', unsafe_allow_html=True)
        st.success("✅ Reporte generado con éxito")
        st.download_button(
            label="📥 Descargar reporte",
            data=output,
            file_name=nombre_reporte,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# FIX para Python 3.13 + Windows (Playwright / Streamlit)
# =========================================================
# Evita "NotImplementedError" al crear subprocesos en Windows.
import sys, asyncio
if sys.platform.startswith("win"):
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
        print("[INFO] Usando WindowsProactorEventLoopPolicy para asyncio (compatible con Playwright).")
    except Exception as e:
        print(f"[WARN] No pude cambiar event loop policy: {e}")
# =========================================================
# FIN FIX asyncio
# =========================================================
from scripts.ui_panels import (
    render_tab_todas_ordenes,
    render_tab_nubes_terceros,
    render_tab_bajas,
)



import os
import re
import numpy as np
from io import BytesIO
from datetime import datetime, date
from pathlib import Path

import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, JsCode

# Importa tu script de generación y el downloader de Superset
from scripts.reporte_general import procesar_reporte_general
from scripts.superset_downloader import download_superset_csvs



def _estilos_tabs():
    st.markdown("""
    <style>
    /* CONTENEDOR DE TABS */
    section[data-testid="stTabs"] [data-baseweb="tab-list"]{
        gap: 2.4rem !important;           /* separa los tabs */
    }

    /* BOTÓN DEL TAB (wrapper) */
    section[data-testid="stTabs"] button[role="tab"]{
        padding: 14px 22px !important;    /* área clickeable */
        border-radius: 10px !important;
    }

    /* TEXTO DEL TAB */
    section[data-testid="stTabs"] button[role="tab"] p{
        font-size: 1.6rem !important;     /* ▲ más grande */
        line-height: 1.5 !important;
        font-weight: 800 !important;      /* bien negrita */
        color: #0f172a !important;
        margin: 0 !important;
        color: #0f172a;
    }

    /* TAB ACTIVO (texto y botón) */
    section[data-testid="stTabs"] button[role="tab"][aria-selected="true"] p{
        color: #791a5a !important;        /* violeta marca */
    }

    /* SUBRAYADO DEL TAB ACTIVO */
    section[data-testid="stTabs"] [data-baseweb="tab-highlight"]{
        height: 4px !important;
        background: #791a5a !important;
        border-radius: 3px !important;
    }

    /* fallback para versiones antiguas (estructura distinta) */
    div.stTabs [data-baseweb="tab"] p{
        font-size: 1.6rem !important;
        font-weight: 800 !important;
    }
    </style>
    """, unsafe_allow_html=True)





# =========================
# CONFIGURACIÓN BÁSICA
# =========================
st.set_page_config(page_title="Generador de reportes automático", layout="wide")
os.makedirs("outputs", exist_ok=True)

# Estado para logs/descargas de Superset
if "superset_logs" not in st.session_state:
    st.session_state.superset_logs = []
if "superset_results" not in st.session_state:
    st.session_state.superset_results = []
if "usar_descargados" not in st.session_state:
    st.session_state.usar_descargados = False
if "csvs_seleccionados" not in st.session_state:
    st.session_state.csvs_seleccionados = []


# =========================
# FUNCIONES AUXILIARES
# =========================
def nombre_salida() -> str:
    """Nombre de archivo con la fecha actual."""
    return f"reporte_general_{datetime.now().strftime('%d-%m-%Y')}.xlsx"


def leer_fuente(archivo) -> pd.DataFrame:
    """Lee CSV o Excel detectando separador automáticamente (CSV) y usando openpyxl para XLSX."""
    nombre = (archivo.name if hasattr(archivo, "name") else str(archivo)).lower()
    if nombre.endswith(".csv"):
        try:
            return pd.read_csv(archivo, sep=None, engine="python")
        except Exception:
            # fallback por separadores comunes
            try_order = [",", ";", "\t", "|"]
            for s in try_order:
                try:
                    if hasattr(archivo, "seek"):
                        archivo.seek(0)
                    return pd.read_csv(archivo, sep=s, engine="c", low_memory=False)
                except Exception:
                    continue
            raise
    return pd.read_excel(archivo, engine="openpyxl")


def leer_fuentes_csv_multiples(rutas: list[str]) -> pd.DataFrame:
    """
    Lee varios CSVs y concatena por columnas (outer join de columnas).
    Intenta detectar separador; agrega columna __ORIGEN con el nombre del archivo.
    """
    frames = []
    for ruta in rutas:
        try:
            # Detección rápida de separador
            try:
                df = pd.read_csv(ruta, sep=None, engine="python", low_memory=False)
            except Exception:
                ok = False
                for s in [",", ";", "\t", "|"]:
                    try:
                        df = pd.read_csv(ruta, sep=s, engine="c", low_memory=False)
                        ok = True
                        break
                    except Exception:
                        continue
                if not ok:
                    raise
            df["__ORIGEN"] = Path(ruta).name
            frames.append(df)
        except Exception as e:
            st.warning(f"No pude leer {ruta}: {e}")
    if not frames:
        raise RuntimeError("No se pudo leer ninguno de los CSV seleccionados.")
    # Concat robusto (alineación de columnas)
    df_final = pd.concat(frames, axis=0, ignore_index=True, sort=False)
    return df_final


def cargar_hoja_todas_las_ordenes() -> tuple[pd.DataFrame | None, str]:
    """Carga la hoja 'TODAS LAS ORDENES' del Excel más reciente (memoria o disco) y limpia auxiliares heredadas."""
    bytes_guardados = st.session_state.get("excel_bytes")
    nombre_guardado = st.session_state.get("excel_name")
    if bytes_guardados:
        try:
            xls = pd.ExcelFile(BytesIO(bytes_guardados), engine="openpyxl")
            hoja = "TODAS LAS ORDENES" if "TODAS LAS ORDENES" in xls.sheet_names else xls.sheet_names[0]
            df = pd.read_excel(xls, sheet_name=hoja, engine="openpyxl")
            aux_cols = [c for c in df.columns if c.startswith("_FECHA_") or c.endswith("_DISPLAY")]
            df = df.drop(columns=aux_cols, errors="ignore")
            return df, f"{nombre_guardado} (memoria)"
        except Exception as e:
            return None, f"No pude leer desde memoria: {e}"

    try:
        ruta_default = os.path.join("outputs", nombre_salida())
        if os.path.exists(ruta_default):
            xls = pd.ExcelFile(ruta_default, engine="openpyxl")
            hoja = "TODAS LAS ORDENES" if "TODAS LAS ORDENES" in xls.sheet_names else xls.sheet_names[0]
            df = pd.read_excel(xls, sheet_name=hoja, engine="openpyxl")
            aux_cols = [c for c in df.columns if c.startswith("_FECHA_") or c.endswith("_DISPLAY")]
            df = df.drop(columns=aux_cols, errors="ignore")
            return df, f"{os.path.basename(ruta_default)} (disco)"
        return None, "No encontré un reporte del día en /outputs."
    except Exception as e:
        return None, f"No pude leer desde disco: {e}"


def normalizar_estado_series(df: pd.DataFrame) -> pd.Series:
    """Devuelve una serie de estado en minúsculas ('completed', 'inprogress' o '')."""
    if "ESTADO" in df.columns:
        est = df["ESTADO"].astype(str).str.strip().str.lower()
    elif "Order Status" in df.columns:
        est = df["Order Status"].astype(str).str.strip().str.lower()
    else:
        est = pd.Series([""] * len(df))
    return est


def contar_estados(df: pd.DataFrame) -> tuple[int, int, int]:
    """(total, completas, en_progreso)."""
    est = normalizar_estado_series(df)
    total = len(df)
    completas = (est == "completed").sum()
    en_progreso = (est == "inprogress").sum()
    return total, int(completas), int(en_progreso)


# ------- Parser de fecha con desambiguación por fila -------
_ddmmaa = re.compile(r"^\s*(\d{1,2})[/-](\d{1,2})[/-](\d{2})\s*$")


def _try_make_date(Y: int, m: int, d: int):
    try:
        return date(Y, m, d)
    except ValueError:
        return None


def _split_ddmmaa(val, dias_ref: int | None = None):
    """
    Devuelve (Y, m, d, yy) a partir de un valor de fecha.
      - datetime -> respetar.
      - texto 'N-N-YY' o 'N/N/YY': reglas de desambiguación.
      - último recurso: to_datetime(dayfirst=True)
    """
    if pd.isna(val):
        return None, None, None, None

    if isinstance(val, (pd.Timestamp, datetime)):
        Y = val.year
        m = val.month
        d = val.day
        return Y, m, d, Y % 100

    s = str(val).strip()
    mobj = _ddmmaa.match(s)
    if mobj:
        a = int(mobj.group(1))
        b = int(mobj.group(2))
        yy = int(mobj.group(3))
        Y = 2000 + yy
        hoy = date.today()

        if a > 12 and b <= 12:
            dt = _try_make_date(Y, b, a)
            if dt is None:
                return None, None, None, None
            return Y, dt.month, dt.day, yy
        if a <= 12 and b > 12:
            dt = _try_make_date(Y, a, b)
            if dt is None:
                return None, None, None, None
            return Y, dt.month, dt.day, yy

        dt_ddmm = _try_make_date(Y, b, a)
        dt_mmdd = _try_make_date(Y, a, b)
        if dt_ddmm is None and dt_mmdd is None:
            return None, None, None, None

        if dias_ref is not None and pd.notna(dias_ref):
            best = None
            if dt_ddmm:
                diff1 = abs((hoy - dt_ddmm).days - int(dias_ref))
                best = ("ddmm", diff1, dt_ddmm)
            if dt_mmdd:
                diff2 = abs((hoy - dt_mmdd).days - int(dias_ref))
                if best is None or diff2 < best[1]:
                    best = ("mmdd", diff2, dt_mmdd)
            dt = best[2]
            return Y, dt.month, dt.day, yy

        if dt_ddmm:
            if (hoy - dt_ddmm).days >= -1:
                dt = dt_ddmm
            elif dt_mmdd:
                dt = dt_mmdd
            else:
                dt = dt_ddmm
        else:
            dt = dt_mmdd

        return Y, dt.month, dt.day, yy

    t = pd.to_datetime(s, dayfirst=True, errors="coerce")
    if pd.isna(t):
        return None, None, None, None
    return t.year, t.month, t.day, t.year % 100


def a_iso_y_display(series: pd.Series, dias_ref_series: pd.Series | None = None) -> tuple[pd.Series, pd.Series]:
    """Devuelve (ISO 'YYYY-MM-DD', DISPLAY 'DD-MM-AA') usando desambiguación por fila."""
    iso_vals, disp_vals = [], []
    for i, v in enumerate(series):
        dias_ref = None
        if dias_ref_series is not None:
            try:
                dias_ref = dias_ref_series.iloc[i]
            except Exception:
                dias_ref = None
        Y, m, d, yy = _split_ddmmaa(v, dias_ref=dias_ref)
        if Y is None:
            iso_vals.append("")
            disp_vals.append("")
        else:
            iso_vals.append(f"{Y:04d}-{m:02d}-{d:02d}")
            disp_vals.append(f"{d:02d}-{m:02d}-{yy:02d}")
    return pd.Series(iso_vals, index=series.index), pd.Series(disp_vals, index=series.index)


# =========================
# ENCABEZADO
# =========================
st.title("Generador de reportes automático")

# =========================
# --- DESCARGA SUPERSET ---
st.subheader("Descarga de CSV (Superset)")

with st.expander("Opciones de descarga", expanded=True):
    col1, col2 = st.columns([3, 2])
    with col1:
        dashboard_url = st.text_input(
            "URL del dashboard (permalink con filtros)",
            value="",
            placeholder="https://…/superset/dashboard/…?permalink_key=…  (o explore/p/<key>/)",
        )
        
        dest_root = st.text_input(
            "Carpeta destino",
            value=str(Path.home() / "Downloads" / "superset_csv")
        )
    with col2:
        key_user = st.text_input("Usuario Keycloak", value="")
        key_pass = st.text_input("Contraseña Keycloak", value="", type="password")

    # ⬇️ valores ocultos (defaults)
    MAX_PANELES_DEFAULT = 0      # 0 = todos
    PANEL_TIMEOUT_DEFAULT = 25   # segundos

    run_dl = st.button("⬇️ Descargar CSVs del dashboard", type="primary")

# zona de logs en vivo
log_box = st.empty()
def _log(msg: str):
    st.session_state.superset_logs.append(msg)
    log_box.code("\n".join(st.session_state.superset_logs[-150:]), language="text")

if run_dl:
    st.session_state.superset_logs = []
    st.session_state.superset_results = []

    if not dashboard_url.strip() or not key_user.strip() or not key_pass:
        _log("❌ Completá URL, usuario y contraseña de Keycloak.")
    else:
        day_folder = Path(dest_root) / datetime.now().strftime("%Y-%m-%d")
        _log(f"🚀 Descargando a: {day_folder.resolve()}")

        files = download_superset_csvs(
            dashboard_url=dashboard_url.strip(),
            download_dir=day_folder,
            keycloak_user=key_user.strip(),
            keycloak_pass=key_pass.strip(),
            max_panels=MAX_PANELES_DEFAULT,          # oculto
            panel_timeout=PANEL_TIMEOUT_DEFAULT,     # oculto
            headless=False,
            log=_log,
        )
        st.session_state.superset_results = [str(p) for p in files]



# listado de resultados (si hay) + switch para usarlos en el reporte
if st.session_state.superset_results:
    st.success(f"Descargados: {len(st.session_state.superset_results)} CSV")
    # selección múltiple
    opciones = st.session_state.superset_results
    seleccion = st.multiselect(
        "Elegí qué CSVs usar para el reporte (puede ser más de uno):",
        options=opciones,
        default=opciones,  # por defecto, todos
        format_func=lambda p: Path(p).name,
    )
    st.session_state.csvs_seleccionados = seleccion

    st.toggle(
        "Usar CSVs descargados como fuente del reporte",
        key="usar_descargados",
        value=st.session_state.usar_descargados,
        help="Si está activo, el botón '▶️ Ejecutar y mostrar' usa estos CSVs en lugar de un archivo subido.",
    )

st.divider()


# =========================
# LAYOUT SUPERIOR: ARCHIVO | SCRIPT
# =========================
c_archivo, c_script = st.columns([3, 2], gap="large")

with c_archivo:
    st.subheader("Archivo")
    if "archivo_cargado" not in st.session_state:
        st.session_state.archivo_cargado = None

    if st.session_state.archivo_cargado is None:
        archivo = st.file_uploader("Drag and drop o Browse files", type=["csv", "xlsx"], key="uploader")
        if archivo is not None:
            st.session_state.archivo_cargado = archivo
            st.success(f"Archivo '{archivo.name}' cargado.")
    else:
        st.info(f"Archivo listo: **{st.session_state.archivo_cargado.name}**")
        if st.button("Borrar archivo y subir otro", type="secondary"):
            st.session_state.archivo_cargado = None
            st.experimental_rerun()

with c_script:
    st.subheader("Script")
    script_opciones = {"Reporte general": "reporte_general.py"}
    elegido = st.selectbox("Elegí el script a ejecutar", list(script_opciones.keys()))
    st.caption(f"Seleccionado: **{script_opciones[elegido]}**")

st.divider()

# =========================
# BOTÓN: EJECUTAR Y MOSTRAR
# =========================
ejecutar = st.button("▶️ Ejecutar y mostrar", type="primary")

if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None
    st.session_state.excel_name = None

if ejecutar:
    try:
        buffer = BytesIO()

        # Fuente: CSVs descargados (si así lo elegiste y hay selección válida)
        if st.session_state.usar_descargados:
            rutas = [p for p in (st.session_state.csvs_seleccionados or []) if os.path.exists(p)]
            if not rutas:
                st.error("Activaste 'Usar CSVs descargados', pero no hay CSVs seleccionados disponibles.")
                st.stop()
            df = leer_fuentes_csv_multiples(rutas)
            st.info(f"Usando {len(rutas)} CSV(s) descargados como fuente ({len(df)} filas).")
        else:
            # Fuente: archivo subido
            if st.session_state.archivo_cargado is None:
                st.error("Subí un archivo (CSV/XLSX) o activa 'Usar CSVs descargados'.")
                st.stop()
            df = leer_fuente(st.session_state.archivo_cargado)

        # Ejecutar el pipeline
        procesar_reporte_general(df, buffer)
        buffer.seek(0)

        # Guardar a disco + sesión
        nombre = nombre_salida()
        ruta = os.path.join("outputs", nombre)
        with open(ruta, "wb") as f:
            f.write(buffer.getbuffer())

        st.session_state.excel_bytes = buffer.getvalue()
        st.session_state.excel_name = nombre
        st.success(f"Reporte generado: **{nombre}**")
    except Exception as e:
        st.error(f"Ocurrió un error al ejecutar el script: {e}")

# =========================
# DESCARGA DEL EXCEL
# =========================
if st.session_state.excel_bytes:
    st.download_button(
        "📥 Descargar Excel generado",
        data=st.session_state.excel_bytes,
        file_name=st.session_state.excel_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# =========================
# VISUALIZACIÓN
# =========================
_estilos_tabs()
st.header("Visualización")
tabs = st.tabs(["Todas las Órdenes", "Nubes de terceros", "Bajas"])

with tabs[0]:
    st.subheader("Todas las Órdenes")
    df_all, origen = cargar_hoja_todas_las_ordenes()
    if df_all is None or df_all.empty:
        st.info("Generá el reporte o asegurate de tener el Excel del día en /outputs.")
    else:
        st.caption(f"Fuente: **{origen}**")
        render_tab_todas_ordenes(df_all)

with tabs[1]:
    st.subheader("Nubes de terceros")
    df_all, origen = cargar_hoja_todas_las_ordenes()
    if df_all is None or df_all.empty:
        st.info("Generá el reporte o asegurate de tener el Excel del día en /outputs.")
    else:
        st.caption(f"Fuente: **{origen}**")
        render_tab_nubes_terceros(df_all)

with tabs[2]:
    st.subheader("Bajas")
    df_all, origen = cargar_hoja_todas_las_ordenes()
    if df_all is None or df_all.empty:
        st.info("Generá el reporte o asegurate de tener el Excel del día en /outputs.")
    else:
        st.caption(f"Fuente: **{origen}**")
        render_tab_bajas(df_all) 
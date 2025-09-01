import io
import base64
import re
from datetime import datetime
from typing import Dict, Any, List, Optional

import numpy as np
import pandas as pd
import streamlit as st

# Gráficos (Plotly opcional)
try:
    import plotly.express as px
    HAS_PLOTLY = True
except Exception:
    HAS_PLOTLY = False

from scripts.reporte_general import procesar_reporte_general
from scripts.script2 import procesar_script2


# ================== Config básica ==================
st.set_page_config(
    page_title="Generador de Reportes Automático",
    page_icon="📊",
    layout="wide",
)

# ====== CSS compacto (reduce márgenes / paddings) ======
COMPACT_CSS = """
<style>
.block-container { padding-top: 1.4rem; padding-bottom: 0.8rem; }
section[data-testid="stSidebar"] .block-container { padding-top: 1rem; }
.block-container > div:first-child { margin-top: .25rem; }
div[data-testid="stVerticalBlock"] > div:has(> .stMarkdown) { margin: .2rem 0; }
.stButton > button, .stDownloadButton > button { padding: .45rem .9rem; }
.stSelectbox, .stFileUploader, .stMultiSelect { margin-bottom: .4rem; }
div[data-baseweb="select"] > div { min-height: 38px; }

/* topbar visual */
.topbar { margin-top: .2rem; margin-bottom: .25rem; }
.topbar-inner{ display:flex; align-items:center; justify-content:space-between; }
.topbar-left{ display:flex; align-items:center; gap:.6rem;}
.topbar-logo{ height:28px; }
.title{ font-weight:600; font-size:1.05rem; line-height:1.2; }
.subtitle{ color:#6b7280; font-size:.85rem; }
.topbar-divider{ height:1px; background:#e5e7eb; margin:.4rem 0 .6rem; }
.badge{ display:inline-flex; align-items:center; gap:.35rem; background:#f3f4f6; border:1px solid #e5e7eb; border-radius:999px; padding:.15rem .6rem; font-size:.8rem; color:#374151;}
</style>
"""
st.markdown(COMPACT_CSS, unsafe_allow_html=True)

# ================== Logo opcional ==================
def get_base64_of_bin_file(path: str) -> Optional[str]:
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return None

logo_b64 = get_base64_of_bin_file("mi_logo.png")


# ================== Estado ==================
if "uploaded_file" not in st.session_state:
    st.session_state.uploaded_file = None

if "history" not in st.session_state:
    # cada item: {"filename","payload","script","when","views","df_raw"}
    st.session_state.history = []

# Fuente de la verdad de meses seleccionados
if "selected_months" not in st.session_state:
    st.session_state.selected_months: List[str] = []

# Valor “pendiente” para forzar el widget en el próximo rerun
if "pending_months_value" not in st.session_state:
    st.session_state.pending_months_value: Optional[List[str]] = None

MAX_HISTORY = 5


# ================== Topbar ==================
st.markdown(
    f"""
    <div class="topbar">
      <div class="topbar-inner">
        <div class="topbar-left">
          {'<img src="data:image/png;base64,' + logo_b64 + '" class="topbar-logo" />' if logo_b64 else ''}
          <div>
            <div class="title">Generador de Reportes Automático</div>
            <div class="subtitle">Panel compacto · Subí CSV, elegí script y corré el reporte</div>
          </div>
        </div>
        <div class="topbar-right">
          <span class="badge">UI compacta</span>
        </div>
      </div>
      <div class="topbar-divider"></div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ================== Helpers de columnas/meses ==================
VIS_TABS_ORDER = [
    "Órdenes por PM",          # 1ª pestaña
    "Órdenes por oferta",
    "Estados",
    "Órdenes por categoría",
    "Nubes de terceros",       # nueva pestaña
    "Evolución mensual",
]

MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}

def _col(df: pd.DataFrame, *names: str) -> Optional[str]:
    """Devuelve el nombre real de la primera columna que matchee alguno de los nombres."""
    lower_map = {str(c).lower(): c for c in df.columns}
    for n in names:
        c = lower_map.get(str(n).lower())
        if c is not None:
            return c
    return None

def _month_label(dt: pd.Timestamp) -> str:
    if pd.isna(dt):
        return ""
    return f"{MESES_ES.get(dt.month, dt.strftime('%B')).upper()} {dt.year}"

def _available_months(df: pd.DataFrame) -> List[str]:
    col_fcre = _col(df, "FECHA DE CREACION", "Order Creation Date", "order_creation_date")
    col_fact = _col(df, "FECHA DE ACTIVACION", "Fecha Activación")
    fecha_col = col_fcre or col_fact
    if not fecha_col:
        return []
    fechas = pd.to_datetime(df[fecha_col], errors="coerce")
    meses = fechas.dt.to_period("M").dropna().unique()
    meses_sorted = sorted([pd.Period(m, freq="M") for m in meses])
    return [_month_label(pd.Timestamp(m.start_time)) for m in meses_sorted]

def _filter_by_months(df: pd.DataFrame, selected: List[str]) -> pd.DataFrame:
    if not selected:
        return df
    col_fcre = _col(df, "FECHA DE CREACION", "Order Creation Date", "order_creation_date")
    col_fact = _col(df, "FECHA DE ACTIVACION", "Fecha Activación")
    fecha_col = col_fcre or col_fact
    if not fecha_col:
        return df
    fechas = pd.to_datetime(df[fecha_col], errors="coerce")
    etiquetas = fechas.apply(_month_label)
    return df[etiquetas.isin(selected)].copy()

# Patrones de estados (flexibles ES/EN)
# Después (no-capturantes)
RE_IN_PROGRESS = re.compile(r"\b(?:in\s*progres+s?|in\s*progress|en\s*progres?o|en\s*proceso)\b", re.I)
RE_COMPLETED   = re.compile(r"(?:completad|complete|closed|finaliz)", re.I)


def _normalize_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    """
    Renombra columnas al español que espera reporte_general.py si vienen con otros nombres.
    No altera el df original (devuelve una copia).
    """
    aliases = {
        "SUSCRIPCION": ["SUSCRIPCION", "Subscription", "subscription"],
        "INTERACCION": ["INTERACCION", "Interaction", "interaction"],
        "CATEGORIA": ["CATEGORIA", "Order Category", "order category", "order_category"],
        "ESTADO": ["ESTADO", "Order Status", "status", "order status"],
        "OFERTA": ["OFERTA", "Main Offer", "offer", "main offer", "main_offer"],
        "RESPONSABLE": ["RESPONSABLE", "Responsible", "PM", "Project Manager", "owner", "project manager"],
        "FECHA DE CREACION": [
            "FECHA DE CREACION", "Order Creation Date", "order creation date",
            "order_creation_date", "fecha de creacion", "fecha_creacion"
        ],
        "FECHA DE ACTIVACION": [
            "FECHA DE ACTIVACION", "Fecha Activación", "Activation Date",
            "activation date", "activation_date", "fecha de activacion", "fecha_activacion"
        ],
    }
    lower_map = {str(c).lower(): c for c in df.columns}
    colmap = {}
    for target, names in aliases.items():
        for n in names:
            src = lower_map.get(n.lower())
            if src:
                colmap[src] = target
                break
    return df.rename(columns=colmap).copy()

# --- Detección de nube (Huawei/GCP/Azure) ---
CLOUD_KEYWORDS = {
    "HUAWEI": r"huawei",
    "GCP": r"\bgcp\b|google\s*cloud|google\b",
    "AZURE": r"azure",
}
CLOUD_CAND_COLS = [
    "OFERTA", "CATEGORIA", "PRODUCTO", "PRODUCT", "SERVICE", "SERVICIO",
    "Main Offer", "Offer", "Producto", "Servicio"
]

def _present_text_columns(df: pd.DataFrame) -> List[str]:
    """Devuelve columnas reales (presentes) a partir de la lista de candidatas."""
    cols: List[str] = []
    for c in CLOUD_CAND_COLS:
        real = _col(df, c)
        if real and real not in cols:
            cols.append(real)
    if not cols:
        cols = list(df.columns)
    return cols

def _infer_cloud_provider(df: pd.DataFrame) -> pd.Series:
    """Devuelve Serie con 'HUAWEI'/'GCP'/'AZURE'/None buscando keywords en columnas candidatas."""
    present = _present_text_columns(df)
    text = df[present].astype(str).agg(" ".join, axis=1).str.lower()
    choices = [text.str.contains(pat, regex=True, na=False) for pat in CLOUD_KEYWORDS.values()]
    return np.select(choices, list(CLOUD_KEYWORDS.keys()), default=None)


# ================== Generación de vistas (con filtro de meses) ==================
def build_auto_views(df: pd.DataFrame):
    """Genera las vistas/pestañas aplicando el filtro de meses. Por defecto: último mes. Excluye 'Deactivation'."""
    meses_disponibles = _available_months(df)

    # Si no hay selección previa, setear último mes
    if meses_disponibles and not st.session_state.selected_months:
        st.session_state.selected_months = [meses_disponibles[-1]]
        st.session_state.pending_months_value = st.session_state.selected_months.copy()

    df_f = _filter_by_months(df, st.session_state.selected_months)

    # columnas útiles
    col_sub = _col(df_f, "SUSCRIPCION", "Subscription")
    col_int = _col(df_f, "INTERACCION", "Interaction")
    col_cat = _col(df_f, "CATEGORIA", "Order Category")
    col_est = _col(df_f, "ESTADO", "Order Status")
    col_ofe = _col(df_f, "OFERTA", "Main Offer")
    col_res = _col(df_f, "RESPONSABLE", "Responsible", "PM", "Project Manager")
    col_fcre = _col(df_f, "FECHA DE CREACION", "Order Creation Date", "order_creation_date")
    col_fact = _col(df_f, "FECHA DE ACTIVACION", "Fecha Activación")

    # base limpia (dedup + sin Deactivation)
    base = df_f.copy()
    if col_sub and col_int:
        base = base.drop_duplicates(subset=[col_sub, col_int])
    if col_cat:
        base = base[~base[col_cat].astype(str).str.lower().eq("deactivation")]

    views: Dict[str, Dict[str, Any]] = {}

    # 1) Órdenes por PM
    if col_res and col_est:
        estados = base[col_est].astype(str)

        base_prog = base[estados.str.contains(RE_IN_PROGRESS, na=False)].copy()
        por_res_prog = (
            base_prog.groupby(col_res, as_index=False)
                .size().rename(columns={"size": "CANTIDAD"})
                .sort_values("CANTIDAD", ascending=False)
                .head(20)
        )

        base_comp = base[estados.str.contains(RE_COMPLETED, na=False)].copy()
        por_res_comp = (
            base_comp.groupby(col_res, as_index=False)
                .size().rename(columns={"size": "CANTIDAD"})
                .sort_values("CANTIDAD", ascending=False)
                .head(20)
        )

        views["Órdenes por PM"] = {
            "df": None,
            "charts": [
                {"type": "bar", "x": col_res, "y": "CANTIDAD", "title": "Órdenes activas (In Progress)", "data": por_res_prog, "layout": "half"},
                {"type": "bar", "x": col_res, "y": "CANTIDAD", "title": "Órdenes completadas", "data": por_res_comp, "layout": "half"},
            ],
        }
    else:
        views["Órdenes por PM"] = {"df": pd.DataFrame(), "charts": []}

    # 2) Órdenes por oferta
    if col_ofe:
        por_oferta = (
            base.groupby(col_ofe, as_index=False)
                .size().rename(columns={"size": "CANTIDAD"})
                .sort_values("CANTIDAD", ascending=False)
        )
        views["Órdenes por oferta"] = {
            "df": por_oferta,
            "charts": [{"type": "bar", "x": col_ofe, "y": "CANTIDAD", "title": "Órdenes por oferta"}],
        }
    else:
        views["Órdenes por oferta"] = {"df": pd.DataFrame(), "charts": []}

    # 3) Estados
    if col_est:
        por_estado = (
            base.groupby(col_est, as_index=False)
                .size().rename(columns={"size": "CANTIDAD"})
                .sort_values("CANTIDAD", ascending=False)
        )
        views["Estados"] = {
            "df": por_estado,
            "charts": [{"type": "pie", "x": col_est, "y": "CANTIDAD", "title": "Distribución por estado"}],
        }
    else:
        views["Estados"] = {"df": pd.DataFrame(), "charts": []}

    # 4) Órdenes por categoría
    if col_cat:
        por_cat = (
            base.groupby(col_cat, as_index=False)
                .size().rename(columns={"size": "CANTIDAD"})
                .sort_values("CANTIDAD", ascending=False)
        )
        views["Órdenes por categoría"] = {
            "df": por_cat,
            "charts": [{"type": "bar", "x": col_cat, "y": "CANTIDAD", "title": "Órdenes por categoría"}],
        }
    else:
        views["Órdenes por categoría"] = {"df": pd.DataFrame(), "charts": []}

    # 5) Nubes de terceros (Huawei/GCP/Azure) – SOLO gráfico (sin tabla)
    fecha_col = col_fcre or col_fact
    clouds = _infer_cloud_provider(base)
    base_clouds = base.copy()
    base_clouds["CLOUD"] = clouds
    base_clouds = base_clouds[base_clouds["CLOUD"].isin(["HUAWEI", "GCP", "AZURE"])]

    if fecha_col and not base_clouds.empty:
        tmp = base_clouds.copy()
        tmp["MES"] = pd.to_datetime(tmp[fecha_col], errors="coerce").dt.to_period("M").astype(str)
        serie = (
            tmp.groupby(["MES", "CLOUD"], as_index=False)
               .size().rename(columns={"size": "CANTIDAD"})
               .sort_values(["MES", "CLOUD"])
        )
        # df=None -> NO se muestra tabla en la pestaña
        views["Nubes de terceros"] = {
            "df": None,
            "charts": [
                {"type": "line", "x": "MES", "y": "CANTIDAD", "color": "CLOUD", "title": "Órdenes por mes y nube", "data": serie}
            ],
        }
    else:
        views["Nubes de terceros"] = {"df": pd.DataFrame(), "charts": []}

    # 6) Evolución mensual
    if fecha_col:
        tmp_all = base.copy()
        tmp_all["MES"] = pd.to_datetime(tmp_all[fecha_col], errors="coerce").dt.to_period("M").astype(str)
        por_mes = (
            tmp_all.groupby("MES", as_index=False)
               .size().rename(columns={"size": "CANTIDAD"})
               .sort_values("MES")
        )
        views["Evolución mensual"] = {
            "df": por_mes,
            "charts": [
                {"type": "line", "x": "MES", "y": "CANTIDAD", "title": "Órdenes por mes"},
                {"type": "area", "x": "MES", "y": "CANTIDAD", "title": "Acumulado mensual"},
            ],
        }
    else:
        views["Evolución mensual"] = {"df": pd.DataFrame(), "charts": []}

    ordered = {k: views.get(k, {"df": pd.DataFrame(), "charts": []}) for k in VIS_TABS_ORDER}
    return ordered, meses_disponibles


def render_views(views: Dict[str, Dict[str, Any]]) -> None:
    """Renderiza las pestañas definidas en VIS_TABS_ORDER con tablas y gráficos."""
    tabs = st.tabs(VIS_TABS_ORDER)
    for tab, name in zip(tabs, VIS_TABS_ORDER):
        with tab:
            v = views.get(name, {"df": pd.DataFrame(), "charts": []})
            df_default = v.get("df")
            charts: List[Dict[str, Any]] = v.get("charts", [])

            # Solo mostramos tabla si df_default no es None/empty.
            if df_default is not None and isinstance(df_default, pd.DataFrame) and not df_default.empty:
                st.dataframe(df_default, width="stretch")

            if charts and HAS_PLOTLY:
                half_charts = [c for c in charts if c.get("layout") == "half"]
                other_charts = [c for c in charts if c.get("layout") != "half"]

                for i in range(0, len(half_charts), 2):
                    c1, c2 = st.columns(2)
                    for c, holder in zip(half_charts[i:i+2], (c1, c2)):
                        with holder:
                            _plot_chart(c, df_default, height=300)

                for cfg in other_charts:
                    _plot_chart(cfg, df_default, height=420)
            elif charts and not HAS_PLOTLY:
                st.info("Instalá 'plotly' para ver los gráficos (pip install plotly).")

def _plot_chart(cfg: Dict[str, Any], df_default: Optional[pd.DataFrame], height: int = 420):
    ctype = cfg.get("type", "bar")
    x = cfg.get("x"); y = cfg.get("y"); color = cfg.get("color", None)
    title = cfg.get("title", "")
    df = cfg.get("data", df_default)
    if df is None or df.empty:
        st.caption("Sin datos para mostrar.")
        return
    if ctype == "bar":
        fig = px.bar(df, x=x, y=y, color=color, title=title)
    elif ctype == "line":
        fig = px.line(df, x=x, y=y, color=color, title=title)
    elif ctype == "area":
        fig = px.area(df, x=x, y=y, color=color, title=title)
    elif ctype == "pie":
        fig = px.pie(df, names=x, values=y if isinstance(y, str) else None, title=title)
    else:
        fig = px.bar(df, x=x, y=y, color=color, title=title)
    fig.update_layout(
        margin=dict(l=10, r=10, t=40, b=10),
        height=height,
        legend=dict(orientation="h", yanchor="bottom", y=-0.2),
    )
    st.plotly_chart(fig, width="stretch")


def build_excel_from_views(views: Dict[str, Dict[str, Any]]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        for sheet_name in VIS_TABS_ORDER:
            v = views.get(sheet_name, {})
            df = v.get("df")
            if df is None or df.empty:
                continue
            safe = str(sheet_name)[:31]
            df.to_excel(writer, index=False, sheet_name=safe)
    buf.seek(0)
    return buf.read()


# ================== PANEL COMPACTO: subir/ejecutar (dos columnas) ==================
c_left, c_right = st.columns(2)

with c_left:
    st.subheader("📥 Archivo")
    uploaded = st.file_uploader("Seleccioná el CSV", type="csv", label_visibility="collapsed")
    if uploaded:
        st.session_state.uploaded_file = uploaded

with c_right:
    st.subheader("⚙️ Script")
    opcion = st.selectbox("Tipo de procesamiento", ["", "Reporte General", "Script 2"], index=0, label_visibility="collapsed")
    if st.session_state.uploaded_file:
        st.markdown(f"<span class='badge'>Archivo: {st.session_state.uploaded_file.name}</span>", unsafe_allow_html=True)
        if st.button("❌ Eliminar archivo", key="delete_file", type="secondary", width="stretch"):
            st.session_state.uploaded_file = None
            st.session_state.selected_months = []
            st.session_state.pending_months_value = []
            st.rerun()

# Botón centrado abajo (depende de ambas columnas)
col1, col2, col3 = st.columns([3, 1, 3])
with col2:
    run_clicked = st.button("▶️ Ejecutar y mostrar", type="primary", width="stretch")


# ================== Ejecutar y VISUALIZAR ==================
if st.session_state.uploaded_file and opcion and opcion != "" and run_clicked:
    df_input = pd.read_csv(st.session_state.uploaded_file)

    # Excel (se mantiene)
    out_excel = io.BytesIO()
    if opcion == "Reporte General":
        try:
            df_excel = _normalize_for_excel(df_input)  # normalizo SOLO para el Excel
            procesar_reporte_general(df_excel, out_excel)
        except TypeError:
            pass
    elif opcion == "Script 2":
        procesar_script2(df_input, out_excel)
    out_excel.seek(0)
    payload = out_excel.read()

    # Vistas iniciales (setea mes por defecto si aplica)
    views, _ = build_auto_views(df_input)

    # Guardar en historial
    nombre_reporte = f"reporte_general_{datetime.today().strftime('%Y-%m-%d')}.xlsx"
    st.session_state.history.insert(0, {
        "filename": nombre_reporte,
        "payload": payload if payload else build_excel_from_views(views),
        "script": opcion,
        "when": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "views": views,
        "df_raw": df_input.copy(),
    })
    st.session_state.history = st.session_state.history[:MAX_HISTORY]


# ================== Historial + Filtro de meses (FORM) + Visualización ==================
st.subheader("🗂️ Historial & Filtro")

if not st.session_state.history:
    st.caption("Todavía no generaste reportes en esta sesión.")
else:
    # fila compacta: selector + acciones
    col_h1, col_h2 = st.columns([3, 2])
    with col_h1:
        labels = [f"{i+1}. {h['filename']} · {h['script']} · {h['when']}" for i, h in enumerate(st.session_state.history)]
        sel = st.selectbox("Reporte", options=list(range(len(labels))), format_func=lambda i: labels[i], label_visibility="collapsed")
    with col_h2:
        c_dl, c_clr = st.columns(2)
        current = st.session_state.history[sel]
        with c_dl:
            st.download_button(
                label="📥 Descargar Excel",
                data=current["payload"],
                file_name=current["filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_history_{sel}",
                width="stretch",
            )
        with c_clr:
            if st.button("🧹 Limpiar historial", width="stretch"):
                st.session_state.history = []
                st.session_state.selected_months = []
                st.session_state.pending_months_value = []
                st.rerun()

# ======= Filtro de meses =======
st.markdown("##### 🗓️ Filtro de meses")
df_for_filter = current["df_raw"]
meses_disponibles = _available_months(df_for_filter)

# Bootstrap: si no hay selección previa, por defecto el último mes (si existe)
if not st.session_state.get("selected_months"):
    st.session_state["selected_months"] = [meses_disponibles[-1]] if meses_disponibles else []

with st.form("months_form", clear_on_submit=False):
    # NO usamos key. Tomamos el default desde session_state y luego usamos el valor devuelto por el widget.
    sel = st.multiselect(
        "Seleccioná uno o más meses",
        options=meses_disponibles,
        default=st.session_state["selected_months"],
    )

    col_f1, col_f2, col_f3, col_f4 = st.columns([1, 1, 1, 1])
    apply_btn  = col_f1.form_submit_button("✅ Aplicar meses")
    ultimo_btn = col_f2.form_submit_button("📅 Último mes")
    clear_btn  = col_f3.form_submit_button("🗑️ Limpiar selección")
    all_btn    = col_f4.form_submit_button("📆 Todos los meses")

# Acciones del form (prioridad: limpiar > último > todos > aplicar)
if clear_btn:
    st.session_state["selected_months"] = []
    st.rerun()
elif ultimo_btn:
    last = [meses_disponibles[-1]] if meses_disponibles else []
    st.session_state["selected_months"] = last
    st.rerun()
elif all_btn:
    st.session_state["selected_months"] = meses_disponibles.copy()
    st.rerun()
elif apply_btn:
    st.session_state["selected_months"] = sel

st.divider()

# === Reconstruir vistas con el filtro actual (ojo: FUERA del form y FUERA de los if/elif) ===
views_filtered, _ = build_auto_views(df_for_filter)

st.markdown("### 📊 Visualización")
render_views(views_filtered)



    
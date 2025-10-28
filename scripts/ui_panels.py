# scripts/ui_panels.py
from __future__ import annotations
from pathlib import Path
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

from .dates import a_iso_y_display
from .grid import build_date_comparators, configure_common_grid
from .superset_downloader import download_superset_csvs


# =========================================================
# Helpers
# =========================================================
def normalizar_estado_series(df: pd.DataFrame) -> pd.Series:
    """
    Devuelve la columna de estado en minúsculas.
    Prioriza 'ESTADO' y, si no está, usa 'Order Status'.
    """
    if "ESTADO" in df.columns:
        s = df["ESTADO"]
    elif "Order Status" in df.columns:
        s = df["Order Status"]
    else:
        return pd.Series([""] * len(df), index=df.index, dtype="object")
    return s.astype(str).str.strip().str.lower()


def dias_habiles_entre(creacion_iso: pd.Series, activacion_iso: pd.Series | None = None) -> pd.Series:
    """
    Días hábiles (lun–vie) entre fecha de creación y:
      - fecha de activación (si existe), o
      - hoy (si no hay activación).
    No incluye feriados; sólo excluye fines de semana.
    """
    fc = pd.to_datetime(creacion_iso, errors="coerce")
    if activacion_iso is not None:
        fa = pd.to_datetime(activacion_iso, errors="coerce")
    else:
        fa = pd.Series(pd.NaT, index=fc.index)

    hoy = pd.Timestamp.today().normalize()
    end = fa.fillna(hoy)

    start_np = fc.dt.date.values.astype("datetime64[D]")
    end_np = end.dt.date.values.astype("datetime64[D]")

    mask_valid = (~pd.isna(fc)) & (~pd.isna(end))
    out = np.full(len(fc), np.nan)
    out[mask_valid.values] = np.busday_count(start_np[mask_valid.values], end_np[mask_valid.values])

    return pd.Series(out, index=fc.index).astype("Int64")


# =========================================================
# Panel de descarga desde Superset (UI simple)
# =========================================================
def render_superset_download_panel():
    """
    Panel minimal para descargar CSVs desde Superset.
    No muestra filtro por título, máx. paneles ni timeout.
    Usa defaults internos: title_filter_regex="", max_panels=0, panel_timeout=25.
    Devuelve (lista_de_csvs_descargados, usar_descargados_bool).
    """
    if "superset_logs" not in st.session_state:
        st.session_state.superset_logs = []
    if "superset_results" not in st.session_state:
        st.session_state.superset_results = []
    if "usar_descargados" not in st.session_state:
        st.session_state.usar_descargados = False
    if "csvs_seleccionados" not in st.session_state:
        st.session_state.csvs_seleccionados = []

    st.subheader("Descarga de CSV (Superset)")

    with st.expander("Opciones de descarga", expanded=True):
        col1, col2 = st.columns([3, 2])
        with col1:
            dashboard_url = st.text_input(
                "URL del dashboard o explore (permalink)",
                value="",
                placeholder="https://…/superset/dashboard/… o /superset/explore/…",
            )
            dest_root = st.text_input(
                "Carpeta destino",
                value=str(Path.home() / "Downloads" / "superset_csv"),
            )
        with col2:
            key_user = st.text_input("Usuario Keycloak", value="")
            key_pass = st.text_input("Contraseña Keycloak", value="", type="password")

        run_dl = st.button("⬇️ Descargar CSVs", type="primary", key="btn_dl_csv")

    # zona de logs en vivo
    log_box = st.empty()

    def _log(msg: str):
        st.session_state.superset_logs.append(msg)
        log_box.code("\n".join(st.session_state.superset_logs[-200:]), language="text")

    if run_dl and not st.session_state.get("dl_running", False):
        st.session_state.dl_running = True  # guard anti doble click
        try:
            st.session_state.superset_logs = []
            st.session_state.superset_results = []

            if not dashboard_url.strip():
                _log("❌ Faltó la URL del dashboard o explore.")
            elif not key_user.strip() or not key_pass:
                _log("❌ Completá usuario y contraseña de Keycloak.")
            else:
                day_folder = Path(dest_root) / datetime.now().strftime("%Y-%m-%d")
                _log(f"🚀 Iniciando descarga a: {day_folder.resolve()}")

                files = download_superset_csvs(
                    dashboard_url=dashboard_url.strip(),
                    download_dir=day_folder,
                    keycloak_user=key_user.strip(),
                    keycloak_pass=key_pass,
                    title_filter_regex="",  # sin filtro
                    max_panels=0,           # todos
                    panel_timeout=25,       # 25s
                    headless=True,          # invisible
                    log=_log,
                )
                st.session_state.superset_results = [str(p) for p in files]
        finally:
            st.session_state.dl_running = False

    # listado de resultados (si hay)
    if st.session_state.superset_results:
        st.success(f"Descargados: {len(st.session_state.superset_results)} CSV")
        opciones = st.session_state.superset_results
        seleccion = st.multiselect(
            "Elegí qué CSVs usar para el reporte (puede ser más de uno):",
            options=opciones,
            default=opciones,
            format_func=lambda p: Path(p).name,
        )
        st.session_state.csvs_seleccionados = seleccion

        st.toggle(
            "Usar CSVs descargados como fuente del reporte",
            key="usar_descargados",
            value=st.session_state.usar_descargados,
            help="Si está activo, el botón '▶️ Ejecutar y mostrar' usa estos CSVs en lugar de un archivo subido.",
        )

    # Devuelve a app.py los paths y el toggle
    return st.session_state.csvs_seleccionados, st.session_state.usar_descargados


# =========================================================
# KPIs superiores (botones)
# =========================================================
def render_top_kpis(df_all: pd.DataFrame) -> str:
    est = normalizar_estado_series(df_all)
    total = len(df_all)
    completas = int((est == "completed").sum())
    prog = int((est == "inprogress").sum())

    if "filtro_estado" not in st.session_state:
        st.session_state.filtro_estado = "todas"

    b1, b2, b3 = st.columns(3)
    with b1:
        if st.button(
            f"Todas ({total})",
            type="primary" if st.session_state.filtro_estado == "todas" else "secondary",
        ):
            st.session_state.filtro_estado = "todas"
            st.rerun()
    with b2:
        if st.button(
            f"Completas ({completas})",
            type="primary" if st.session_state.filtro_estado == "completed" else "secondary",
        ):
            st.session_state.filtro_estado = "completed"
            st.rerun()
    with b3:
        if st.button(
            f"En progreso ({prog})",
            type="primary" if st.session_state.filtro_estado == "inprogress" else "secondary",
        ):
            st.session_state.filtro_estado = "inprogress"
            st.rerun()

    st.caption("Tip: clic en encabezados para ordenar globalmente.")
    return st.session_state.filtro_estado


# =========================================================
# Tab: Todas las Órdenes
# =========================================================
def render_tab_todas_ordenes(df_all: pd.DataFrame):
    st.subheader("Todas las Órdenes")
    if df_all is None or df_all.empty:
        st.info("Generá el reporte o asegurate de tener el Excel del día en /outputs.")
        return

    filtro = render_top_kpis(df_all)

    df_show = df_all.copy()
    if "FECHA DE CREACION" in df_show.columns:
        iso_crea, disp_crea = a_iso_y_display(df_show["FECHA DE CREACION"])
        df_show["_FECHA_CREACION_ISO"] = iso_crea
        df_show["FECHA_DE_CREACION_DISPLAY"] = disp_crea
    if "FECHA DE ACTIVACION" in df_show.columns:
        iso_act, disp_act = a_iso_y_display(df_show["FECHA DE ACTIVACION"])
        df_show["_FECHA_ACTIVACION_ISO"] = iso_act
        df_show["FECHA_DE_ACTIVACION_DISPLAY"] = disp_act

    # DÍAS ABIERTA (hábiles) → si existe fecha de creación
    if "_FECHA_CREACION_ISO" in df_show.columns:
        activ_col = df_show["_FECHA_ACTIVACION_ISO"] if "_FECHA_ACTIVACION_ISO" in df_show.columns else None
        df_show["DIAS ABIERTA"] = dias_habiles_entre(df_show["_FECHA_CREACION_ISO"], activ_col)

    # filtro KPI
    est = df_show.get("ESTADO", pd.Series([""] * len(df_show))).astype(str).str.lower()
    if filtro == "completed":
        df_show = df_show[est == "completed"].copy()
    elif filtro == "inprogress":
        df_show = df_show[est == "inprogress"].copy()

    columnas = [
        "CATEGORIA",
        "OFERTA",
        "SUSCRIPCION",
        "RESPONSABLE",
        "NOMBRE DEL CLIENTE",
        "INTERACCION",
        "FECHA_DE_CREACION_DISPLAY",
        "FECHA_DE_ACTIVACION_DISPLAY",
        "DIAS ABIERTA",
        "_FECHA_CREACION_ISO",
        "_FECHA_ACTIVACION_ISO",
    ]
    cols_existentes = [c for c in columnas if c in df_show.columns]
    df_show = df_show[cols_existentes].copy()
    df_show.insert(0, "#", range(1, len(df_show) + 1))

    cmp_crea, cmp_act = build_date_comparators()
    gb = GridOptionsBuilder.from_dataframe(df_show)
    configure_common_grid(gb)
    gb.configure_column("#", sortable=False, filter=False, pinned="left", width=90)
    gb.configure_column(
        "FECHA_DE_CREACION_DISPLAY",
        header_name="FECHA\nDE CREACION",
        comparator=cmp_crea,
        wrapHeaderText=True,
        autoHeaderHeight=True,
        width=145,
    )
    gb.configure_column(
        "FECHA_DE_ACTIVACION_DISPLAY",
        header_name="FECHA\nDE ACTIVACION",
        comparator=cmp_act,
        wrapHeaderText=True,
        autoHeaderHeight=True,
        width=145,
    )
    gb.configure_column("_FECHA_CREACION_ISO", hide=True)
    gb.configure_column("_FECHA_ACTIVACION_ISO", hide=True)

    # 👉 Mostrar página de 100 filas y expandir sin scroll interno
    gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=100)
    grid_options = gb.build()
    grid_options["domLayout"] = "autoHeight"
    grid_options["rowHeight"] = 34  # opcional

    AgGrid(
        df_show,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.NO_UPDATE,
        fit_columns_on_grid_load=True,
        theme="balham",
        height=0,  # autoHeight => sin scroll interno
        enable_enterprise_modules=False,
        allow_unsafe_jscode=True,
    )


# =========================================================
# Tab: Nubes de terceros (gráfico + detalle)
# =========================================================
def render_tab_nubes_terceros(df_all: pd.DataFrame) -> None:
    st.subheader("Nubes de terceros")

    base = df_all.copy()

    # --- Fechas base
    if "FECHA DE CREACION" in base.columns:
        iso_crea, disp_crea = a_iso_y_display(base["FECHA DE CREACION"])
        base["_FECHA_CREACION_ISO"] = iso_crea
        base["FECHA_DE_CREACION_DISPLAY"] = disp_crea
        base["_FECHA_CREACION_DT"] = pd.to_datetime(base["_FECHA_CREACION_ISO"], errors="coerce")
    else:
        base["_FECHA_CREACION_ISO"] = ""
        base["FECHA_DE_CREACION_DISPLAY"] = ""
        base["_FECHA_CREACION_DT"] = pd.NaT

    # --- Normalizaciones
    estado_norm = normalizar_estado_series(base)
    categoria_norm = base["CATEGORIA"].astype(str).str.strip()

    # Filtro: Completed + SalesOrder + Infraestructura como Servicio + (GCP/Huawei/Azure)
    oferta_norm = base["OFERTA"].astype(str)
    mask_nube = (
        oferta_norm.str.contains("INFRAESTRUCTURA COMO SERVICIO", case=False, na=False)
        & (
            oferta_norm.str.contains("GCP", case=False, na=False)
            | oferta_norm.str.contains("HUAWEI", case=False, na=False)
            | oferta_norm.str.contains("AZURE", case=False, na=False)
        )
    )
    mask_total = (estado_norm == "completed") & (categoria_norm == "SalesOrder") & mask_nube
    df_cloud = base.loc[mask_total].copy()

    # Etiqueta de nube
    def _tag_nube(s: str) -> str:
        s = str(s).upper()
        if "GCP" in s:
            return "GCP"
        if "HUAWEI" in s:
            return "Huawei"
        if "AZURE" in s:
            return "Azure"
        return "Otra"

    df_cloud["NUBE"] = base["OFERTA"].astype(str).map(_tag_nube)

    # Serie por mes (conteo por NUBE)
    df_cloud["MES"] = df_cloud["_FECHA_CREACION_DT"].dt.to_period("M").dt.to_timestamp()
    serie = (
        df_cloud.dropna(subset=["MES"])
        .groupby(["MES", "NUBE"], as_index=False)
        .size()
        .rename(columns={"size": "CANTIDAD"})
        .sort_values(["MES", "NUBE"])
    )

    # Gráfico (Altair)
    try:
        import altair as alt

        # Colores fijos por nube
        color_scale = alt.Scale(
            domain=["GCP", "Huawei", "Azure"],
            range=["#e63946", "#2a9d8f", "#457b9d"]
        )

        line = alt.Chart(serie).mark_line(interpolate="monotone").encode(
            x=alt.X("yearmonth(MES):T", title="Mes", axis=alt.Axis(format="%b %Y", labelAngle=0)),
            y=alt.Y("CANTIDAD:Q", title="Órdenes completadas"),
            color=alt.Color("NUBE:N", title="Nube", scale=color_scale),
            tooltip=[
                alt.Tooltip("yearmonth(MES):T", title="Mes", format="%b %Y"),
                alt.Tooltip("NUBE:N", title="Nube"),
                alt.Tooltip("CANTIDAD:Q", title="Cantidad"),
            ],
        ).properties(width="container", height=320)

        pts = alt.Chart(serie).mark_circle(size=64, opacity=0.9).encode(
            x=alt.X("yearmonth(MES):T", axis=alt.Axis(format="%b %Y", labelAngle=0)),
            y="CANTIDAD:Q",
            color=alt.Color("NUBE:N", scale=color_scale, legend=None),
            tooltip=[
                alt.Tooltip("yearmonth(MES):T", title="Mes", format="%b %Y"),
                alt.Tooltip("NUBE:N", title="Nube"),
                alt.Tooltip("CANTIDAD:Q", title="Cantidad"),
            ],
        )

        st.altair_chart(line + pts, use_container_width=True)

    except Exception as e:
        st.info(f"No se pudo renderizar Altair ({e}). Muestro tabla abajo.")

    # ================== Detalle (selector + AgGrid) ==================
    st.divider()
    st.markdown("**Órdenes consideradas**")

    # Elegir “Mes seleccionado” o “Todos los meses”
    modo = st.radio(
        "Mostrar",
        options=("Mes seleccionado", "Todos los meses"),
        index=0,
        horizontal=True,
        key="modo_detalle_nubes",
    )

    if modo == "Mes seleccionado":
        meses_opts = (
            df_cloud.dropna(subset=["_FECHA_CREACION_DT"])["_FECHA_CREACION_DT"]
            .dt.to_period("M")
            .dt.to_timestamp()
            .drop_duplicates()
            .sort_values()
        )
        default_mes = meses_opts.max() if not meses_opts.empty else None

        mes_sel = st.selectbox(
            "Mes",
            options=list(meses_opts),
            index=(len(meses_opts) - 1) if default_mes is not None else 0,
            format_func=lambda d: pd.Timestamp(d).strftime("%Y-%m"),
            key="detalle_mes_nubes",
        )

        df_detalle = df_cloud[
            df_cloud["_FECHA_CREACION_DT"].dt.to_period("M").dt.to_timestamp() == mes_sel
        ].copy()
    else:
        df_detalle = df_cloud.copy()

    # Fechas display de activación si faltan
    if "FECHA DE ACTIVACION" in df_detalle.columns and "FECHA_DE_ACTIVACION_DISPLAY" not in df_detalle.columns:
        iso_act, disp_act = a_iso_y_display(df_detalle["FECHA DE ACTIVACION"])
        df_detalle["_FECHA_ACTIVACION_ISO"] = iso_act
        df_detalle["FECHA_DE_ACTIVACION_DISPLAY"] = disp_act

    # Días hábiles abiertos (creación → hoy)
    if "_FECHA_CREACION_ISO" in df_detalle.columns:
        fc_dt = pd.to_datetime(df_detalle["_FECHA_CREACION_ISO"], errors="coerce")
        hoy = pd.Timestamp.today().normalize()
        df_detalle["DIAS ABIERTA"] = [
            int(np.busday_count(start.date(), hoy.date())) if pd.notnull(start) else None
            for start in fc_dt
        ]

    columnas_objetivo = [
        "NUBE",
        "CATEGORIA",
        "OFERTA",
        "SUSCRIPCION",
        "RESPONSABLE",
        "NOMBRE DEL CLIENTE",
        "INTERACCION",
        "FECHA_DE_CREACION_DISPLAY",
        "FECHA_DE_ACTIVACION_DISPLAY",
        "DIAS ABIERTA",
        "_FECHA_CREACION_ISO",
        "_FECHA_ACTIVACION_ISO",
    ]
    cols_existentes = [c for c in columnas_objetivo if c in df_detalle.columns]
    df_show = df_detalle[cols_existentes].copy()
    df_show.insert(0, "#", range(1, len(df_show) + 1))

    cmp_crea, cmp_act = build_date_comparators()
    gb = GridOptionsBuilder.from_dataframe(df_show)
    configure_common_grid(gb)
    gb.configure_column("#", sortable=False, filter=False, pinned="left", width=90)
    gb.configure_column(
        "FECHA_DE_CREACION_DISPLAY",
        header_name="FECHA\nDE CREACION",
        comparator=cmp_crea,
        wrapHeaderText=True,
        autoHeaderHeight=True,
        width=145,
    )
    gb.configure_column(
        "FECHA_DE_ACTIVACION_DISPLAY",
        header_name="FECHA\nDE ACTIVACION",
        comparator=cmp_act,
        wrapHeaderText=True,
        autoHeaderHeight=True,
        width=145,
    )
    gb.configure_column("_FECHA_CREACION_ISO", hide=True)
    gb.configure_column("_FECHA_ACTIVACION_ISO", hide=True)

    gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=100)
    grid_options = gb.build()
    grid_options["domLayout"] = "autoHeight"
    grid_options["rowHeight"] = 34

    AgGrid(
        df_show,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.NO_UPDATE,
        fit_columns_on_grid_load=True,
        theme="balham",
        height=0,
        enable_enterprise_modules=False,
        allow_unsafe_jscode=True,
    )


# =========================================================
# Tab: BAJAS (Deactivation)
# =========================================================
def render_tab_bajas(df_all: pd.DataFrame) -> None:
    """
    Vista 'Bajas (Deactivation)':
    - Filtra CATEGORIA == 'Deactivation'
    - Bucketiza OFERTA en: Virtual CPU, IPLAN Cloud Premium, IPLAN Cloud, Virtual Datacenter, GCP, AZURE, HUAWEI (sin 'Otra')
    - Filtro por oferta (pills/multiselect) que afecta gráfico y tabla
    - Gráfico de BARRAS horizontales (sin dona)
    - Tabla detalle con DIAS ABIERTA (hábiles) desde creación → hoy
    """
    st.subheader("Bajas (Deactivation)")

    if df_all is None or df_all.empty:
        st.info("No hay datos para mostrar.")
        return

    base = df_all.copy()

    # Fechas base (creación)
    if "FECHA DE CREACION" in base.columns:
        iso_crea, disp_crea = a_iso_y_display(base["FECHA DE CREACION"])
        base["_FECHA_CREACION_ISO"] = iso_crea
        base["FECHA_DE_CREACION_DISPLAY"] = disp_crea
        base["_FECHA_CREACION_DT"] = pd.to_datetime(base["_FECHA_CREACION_ISO"], errors="coerce")
    else:
        base["_FECHA_CREACION_ISO"] = ""
        base["FECHA_DE_CREACION_DISPLAY"] = ""
        base["_FECHA_CREACION_DT"] = pd.NaT

    # Solo Deactivation
    categoria_norm = base["CATEGORIA"].astype(str).str.strip().str.lower()
    df_bajas = base.loc[categoria_norm.eq("deactivation")].copy()
    if df_bajas.empty:
        st.info("No hay bajas (Deactivation) para mostrar.")
        return

    # Buckets de oferta
    def _bucket_oferta(s: str) -> str:
        u = str(s).upper()
        if "VIRTUAL CPU" in u:
            return "Virtual CPU"
        if "IPLAN CLOUD PREMIUM" in u:
            return "IPLAN Cloud Premium"
        if "IPLAN CLOUD" in u:
            return "IPLAN Cloud"
        if "VIRTUAL DATACENTER" in u:
            return "Virtual Datacenter"
        if "GCP" in u:
            return "GCP"
        if "AZURE" in u:
            return "AZURE"
        if "HUAWEI" in u:
            return "HUAWEI"
        return "Otra"

    df_bajas["OFERTA_BUCKET"] = df_bajas["OFERTA"].apply(_bucket_oferta)

    # Serie agregada (sin "Otra")
    orden_buckets = ["Virtual CPU", "IPLAN Cloud Premium", "IPLAN Cloud",
                     "Virtual Datacenter", "GCP", "AZURE", "HUAWEI"]
    color_map = {
        "Virtual CPU": "#a8dadc",
        "IPLAN Cloud Premium": "#1d3557",
        "IPLAN Cloud": "#457b9d",
        "Virtual Datacenter": "#8d99ae",
        "GCP": "#e63946",
        "AZURE": "#118ab2",
        "HUAWEI": "#2a9d30",
    }

    serie = (
        df_bajas.assign(OFERTA_BUCKET=df_bajas["OFERTA_BUCKET"].astype(str))
                .query("OFERTA_BUCKET in @orden_buckets")
                .groupby("OFERTA_BUCKET", as_index=False)
                .size()
                .rename(columns={"size": "CANTIDAD"})
                .sort_values("OFERTA_BUCKET",
                             key=lambda s: pd.Categorical(s, categories=orden_buckets, ordered=True))
    )

    # ---------- Filtro UI (pills o multiselect) ----------
    try:
        sel_buckets = st.pills(
            "Filtrar por oferta (opcional)",
            options=orden_buckets,
            selection_mode="multi",
            default=orden_buckets,
            key="bajas_pills",
        )
    except Exception:
        sel_buckets = st.multiselect(
            "Filtrar por oferta (opcional)",
            options=orden_buckets,
            default=orden_buckets,
            key="bajas_multiselect",
        )

    sel_buckets = list(sel_buckets) if sel_buckets else orden_buckets

    # Aplicar filtro al gráfico y al detalle
    serie_plot = serie[serie["OFERTA_BUCKET"].isin(sel_buckets)].copy()
    df_bajas_filtrado = df_bajas[df_bajas["OFERTA_BUCKET"].isin(sel_buckets)].copy()

    # ---------- Gráfico: SOLO BARRAS ----------
    try:
        import altair as alt

        color_scale = alt.Scale(
            domain=orden_buckets,
            range=[color_map[b] for b in orden_buckets],
        )

        total_bajas = int(serie_plot["CANTIDAD"].sum()) if not serie_plot.empty else 0
        if total_bajas > 0:
            serie_plot = serie_plot.assign(
                PORC=lambda d: (d["CANTIDAD"] / total_bajas) * 100,
                LABEL=lambda d: d.apply(lambda r: f'{int(r["CANTIDAD"])} ({r["PORC"]:.1f}%)', axis=1),
            )
        else:
            serie_plot = serie_plot.assign(PORC=0.0, LABEL="0 (0.0%)")

        h = 36 * max(len(serie_plot), 5) + 40
        bars = alt.Chart(serie_plot).mark_bar().encode(
            y=alt.Y("OFERTA_BUCKET:N", sort="-x", title="Oferta"),
            x=alt.X("CANTIDAD:Q", title="Bajas"),
            color=alt.Color("OFERTA_BUCKET:N", scale=color_scale, legend=None),
            tooltip=[
                alt.Tooltip("OFERTA_BUCKET:N", title="Oferta"),
                alt.Tooltip("CANTIDAD:Q", title="Cantidad"),
                alt.Tooltip("PORC:Q", title="% del total", format=".1f")
            ]
        ).properties(width="container", height=h)

        labels = bars.mark_text(align="left", dx=6).encode(text="LABEL:N")
        st.altair_chart(bars + labels, use_container_width=True)

    except Exception as e:
        st.info(f"No se pudo renderizar Altair ({e}). Muestro tabla abajo).")

    st.divider()
    st.subheader("Detalle de bajas")

    # DÍAS ABIERTA (hábiles) desde creación → hoy
    if "_FECHA_CREACION_ISO" in df_bajas_filtrado.columns:
        fc_dt = pd.to_datetime(df_bajas_filtrado["_FECHA_CREACION_ISO"], errors="coerce")
        hoy = pd.Timestamp.today().normalize()
        df_bajas_filtrado["DIAS ABIERTA"] = [
            int(np.busday_count(start.date(), hoy.date())) if pd.notnull(start) else None
            for start in fc_dt
        ]

    # Fechas display de activación si existen
    if "FECHA DE ACTIVACION" in df_bajas_filtrado.columns and "FECHA_DE_ACTIVACION_DISPLAY" not in df_bajas_filtrado.columns:
        iso_act, disp_act = a_iso_y_display(df_bajas_filtrado["FECHA DE ACTIVACION"])
        df_bajas_filtrado["_FECHA_ACTIVACION_ISO"] = iso_act
        df_bajas_filtrado["FECHA_DE_ACTIVACION_DISPLAY"] = disp_act
    else:
        df_bajas_filtrado["_FECHA_ACTIVACION_ISO"] = df_bajas_filtrado.get("_FECHA_ACTIVACION_ISO", "")
        df_bajas_filtrado["FECHA_DE_ACTIVACION_DISPLAY"] = df_bajas_filtrado.get("FECHA_DE_ACTIVACION_DISPLAY", "")

    # Tabla
    columnas_objetivo = [
        "OFERTA_BUCKET",
        "OFERTA",
        "SUSCRIPCION",
        "RESPONSABLE",
        "NOMBRE DEL CLIENTE",
        "INTERACCION",
        "FECHA_DE_CREACION_DISPLAY",
        "FECHA_DE_ACTIVACION_DISPLAY",
        "DIAS ABIERTA",
        "_FECHA_CREACION_ISO",
        "_FECHA_ACTIVACION_ISO",
    ]
    cols_existentes = [c for c in columnas_objetivo if c in df_bajas_filtrado.columns]
    df_show = df_bajas_filtrado[cols_existentes].copy()
    df_show.insert(0, "#", range(1, len(df_show) + 1))

    gb = GridOptionsBuilder.from_dataframe(df_show)
    configure_common_grid(gb)
    gb.configure_column("#", sortable=False, filter=False, pinned="left", width=80)
    gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=100)
    if "_FECHA_CREACION_ISO" in df_show.columns:
        gb.configure_column("_FECHA_CREACION_ISO", hide=True)
    if "_FECHA_ACTIVACION_ISO" in df_show.columns:
        gb.configure_column("_FECHA_ACTIVACION_ISO", hide=True)

    grid_options = gb.build()
    grid_options["domLayout"] = "autoHeight"
    grid_options["rowHeight"] = 34

    AgGrid(
        df_show,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.NO_UPDATE,
        fit_columns_on_grid_load=True,
        theme="balham",
        height=0,
        enable_enterprise_modules=False,
        allow_unsafe_jscode=True,
    )

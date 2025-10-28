import pandas as pd

# ---------------- Utilidades de fechas (robustas a múltiples formatos) ----------------
def _parse_mixed_datetime(series: pd.Series) -> pd.Series:
    """
    Intenta parsear fechas en varios formatos sin ambigüedad:
      1) ISO con hora:  YYYY-MM-DD HH:MM:SS
      2) ISO sin hora:  YYYY-MM-DD
      3) Formatos día/mes: DD-MM-YY/AAAA o DD/MM/YY/AAAA (dayfirst=True)
    Retorna datetime64[ns] (NaT donde no se pudo).
    """
    s = pd.to_datetime(series, format="%Y-%m-%d %H:%M:%S", errors="coerce")
    mask = s.isna()
    if mask.any():
        s.loc[mask] = pd.to_datetime(series[mask], format="%Y-%m-%d", errors="coerce")
    mask = s.isna()
    if mask.any():
        # Último intento tolerante para entradas tipo DD-MM-AAAA, DD/MM/AA, etc.
        s.loc[mask] = pd.to_datetime(series[mask], dayfirst=True, errors="coerce")
    return s

def _dt_to_ddmmaa_text(s: pd.Series) -> pd.Series:
    """Convierte datetime a texto DD-MM-AA; NaT -> cadena vacía."""
    out = s.dt.strftime("%d-%m-%y")
    return out.where(~s.isna(), "")


def procesar_reporte_general(df: pd.DataFrame, output):
    # ---------------- Renombrar columnas ----------------
    df.columns = [col.strip() for col in df.columns]
    df = df.rename(columns={
        "Order Status": "ESTADO",
        "Order Creation Date": "FECHA DE CREACION",
        "Responsible": "RESPONSABLE",
        "Nombre Cliente": "NOMBRE DEL CLIENTE",
        "Main Offer": "OFERTA",
        "Subscription": "SUSCRIPCION",
        "Interaction": "INTERACCION",
        "Order Category": "CATEGORIA",
        "Modelo Comercial": "MODELO COMERCIAL",
        "Ejecutivo": "EJECUTIVO",
        "Fecha Activación": "FECHA DE ACTIVACION",
    })

    # ---------------- Eliminar columnas innecesarias ----------------
    columnas_a_eliminar = [
        "Order ID", "Party Role ID", "Mail Contacto Técnico", "Instalation Address",
        "Nombre Elemento", "Monto", "Moneda", "Tipo de Precio", "Delta",
        "Fecha Agendamiento", "Motivo Reprogramación", "Motivo", "Segmento",
        "Fecha Cancelación", "Current Phase",
    ]
    df = df.drop(columns=[c for c in columnas_a_eliminar if c in df.columns], errors="ignore")

    # ---------------- Parseo de fechas sin ambigüedad ----------------
    if "FECHA DE ACTIVACION" in df.columns:
        df["FECHA DE ACTIVACION"] = _parse_mixed_datetime(df["FECHA DE ACTIVACION"])
    if "FECHA DE CREACION" in df.columns:
        df["FECHA DE CREACION"] = _parse_mixed_datetime(df["FECHA DE CREACION"])

    # ---------------- Días abierta ----------------
    hoy = pd.Timestamp.today().normalize()
    if "FECHA DE CREACION" in df.columns:
        fc = pd.to_datetime(df["FECHA DE CREACION"], errors="coerce")
        df["DIAS ABIERTA"] = (hoy - fc).dt.days
    else:
        df["DIAS ABIERTA"] = pd.NA

    # ---------------- Eliminar duplicados ----------------
    # Clave compuesta (según tu preferencia): SUSCRIPCION + INTERACCION (si existe)
    if "INTERACCION" in df.columns:
        df = df.drop_duplicates(subset=["SUSCRIPCION", "INTERACCION"])
    else:
        df = df.drop_duplicates(subset=["SUSCRIPCION"])

    # ---------------- Separar hojas ----------------
    # Abiertas: todo lo que no está "Completed"
    df_abiertas = df[df.get("ESTADO", "") != "Completed"].copy()

    # Top 20 + Abiertas: InProgress y no Deactivation, ordenadas por días abiertas
    df_top_20_abiertas = df_abiertas[
        (df_abiertas.get("ESTADO", "") == "InProgress") &
        (df_abiertas.get("CATEGORIA", "") != "Deactivation")
    ].sort_values(by="DIAS ABIERTA", ascending=False).head(20).copy()

    if "FECHA DE ACTIVACION" in df_top_20_abiertas.columns:
        df_top_20_abiertas = df_top_20_abiertas.drop(columns=["FECHA DE ACTIVACION"])

    # Bajas no completadas
    df_bajas = df[
        (df.get("CATEGORIA", "") == "Deactivation") &
        (df.get("ESTADO", "") != "Completed")
    ].copy()

    columnas_bajas = [
        "ESTADO", "CATEGORIA", "FECHA DE CREACION", "OFERTA", "SUSCRIPCION",
        "RESPONSABLE", "NOMBRE DEL CLIENTE", "INTERACCION", "MODELO COMERCIAL", "DIAS ABIERTA",
    ]
    df_bajas = df_bajas[[c for c in columnas_bajas if c in df_bajas.columns]].sort_values(
        by="DIAS ABIERTA", ascending=False
    )

    # Activaciones (completadas y SalesOrder)
    df_activaciones = df[
        (df.get("ESTADO", "") == "Completed") &
        (df.get("CATEGORIA", "") == "SalesOrder")
    ].copy()

    # ---------------- MES en español (para "ACTIVACIONES POR MODELO") ----------------
    # Evitamos depender de nombres en inglés; usamos el número de mes.
    meses_es = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
                "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

    if "FECHA DE CREACION" in df_activaciones.columns:
        fcrea = pd.to_datetime(df_activaciones["FECHA DE CREACION"], errors="coerce")
        df_activaciones["__YEAR"] = fcrea.dt.year
        df_activaciones["__MONTH"] = fcrea.dt.month
        df_activaciones["MES"] = (
            df_activaciones["__MONTH"].apply(lambda m: meses_es[m - 1] if pd.notna(m) and m >= 1 and m <= 12 else "")
            + " "
            + df_activaciones["__YEAR"].astype("Int64").astype(str)
        ).where(~fcrea.isna(), pd.NA)
    else:
        df_activaciones["MES"] = pd.NA

    # Lista de meses ordenada (YYYY-MM desc)
    if "FECHA DE CREACION" in df_activaciones.columns:
        fcrea = pd.to_datetime(df_activaciones["FECHA DE CREACION"], errors="coerce")
        orden_mes = fcrea.dt.to_period("M").astype("period[M]")
        df_activaciones["__ORDEN_MES"] = orden_mes
        meses = (
            df_activaciones.loc[~df_activaciones["MES"].isna(), ["MES", "__ORDEN_MES"]]
            .drop_duplicates()
            .sort_values("__ORDEN_MES", ascending=False)["MES"]
            .tolist()
        )
    else:
        meses = []

    # Órdenes completadas (para hoja "ORDENES ACTIVADAS")
    df_activadas = df[df.get("ESTADO", "") == "Completed"].copy()

    # ---------------- Exportar a Excel (fechas como TEXTO "DD-MM-AA") ----------------
    def df_fechas_a_texto(df_src: pd.DataFrame) -> pd.DataFrame:
        """
        Devuelve una copia con columnas 'FECHA*' convertidas a texto DD-MM-AA.
        No usa dayfirst=True para parseo principal; asume que ya están en datetime
        y sólo cae a un parseo laxo si hubiera strings colados.
        """
        df_out = df_src.copy()
        for col in df_out.columns:
            if "FECHA" in col.upper():
                # Si ya es datetime -> formateo directo; si no, intento parseo tolerante
                if pd.api.types.is_datetime64_any_dtype(df_out[col]):
                    s = df_out[col]
                else:
                    s = _parse_mixed_datetime(df_out[col])
                df_out[col] = _dt_to_ddmmaa_text(s).astype(object)  # escribir como texto
        return df_out

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_numbers": False}},  # evita auto-conversión a número
    ) as writer:

        def export_and_autofit(df_export: pd.DataFrame, sheet_name: str):
            df_txt = df_fechas_a_texto(df_export)
            df_txt.to_excel(writer, sheet_name=sheet_name, index=False, na_rep="")
            ws = writer.sheets[sheet_name]
            for i, col in enumerate(df_txt.columns):
                try:
                    maxlen = int(df_txt[col].astype(str).map(len).max())
                except Exception:
                    maxlen = len(col)
                ws.set_column(i, i, max(maxlen, len(col)) + 4)

        # Hojas principales
        export_and_autofit(df, "TODAS LAS ORDENES")
        export_and_autofit(df_abiertas, "ORDENES ABIERTAS")
        export_and_autofit(df_top_20_abiertas, "TOP 20 + ABIERTAS")
        export_and_autofit(df_bajas, "BAJAS")
        export_and_autofit(df_activadas, "ORDENES ACTIVADAS")

        # ---------------- Hoja: ACTIVACIONES POR MODELO ----------------
        workbook  = writer.book
        worksheet = workbook.add_worksheet("ACTIVACIONES POR MODELO")
        writer.sheets["ACTIVACIONES POR MODELO"] = worksheet

        startrow = 0
        for mes in meses:
            bloque = df_activaciones[df_activaciones["MES"] == mes].copy()

            # Asegurar fechas como texto en el bloque exportado
            bloque_txt = df_fechas_a_texto(bloque)

            # Pivot por OFERTA x MODELO COMERCIAL (conteo de SUSCRIPCION)
            tabla_mes = pd.pivot_table(
                bloque_txt,
                index="OFERTA",
                columns="MODELO COMERCIAL",
                values="SUSCRIPCION",
                aggfunc="count",
                fill_value=0,
                margins=True,
                margins_name="Suma total",
            )

            # Título del bloque
            worksheet.write(startrow, 0, mes)

            # Escribir la tabla
            tabla_mes_reset = tabla_mes.reset_index()
            tabla_mes_reset.to_excel(
                writer,
                sheet_name="ACTIVACIONES POR MODELO",
                startrow=startrow + 1,
                index=False
            )

            # Ajustes de ancho
            worksheet.set_column(0, 0, 48)
            worksheet.set_column(1, 10, 13)

            # Avanzar fila para siguiente bloque (+4 de separación)
            startrow += len(tabla_mes_reset) + 4

    output.seek(0)

import pandas as pd

def procesar_reporte_general(df, output):
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
    df = df.drop(columns=[c for c in columnas_a_eliminar if c in df.columns])

    # ---------------- Parseo de fechas para cálculos ----------------
    if "FECHA DE ACTIVACION" in df.columns:
        df["FECHA DE ACTIVACION"] = pd.to_datetime(df["FECHA DE ACTIVACION"], errors="coerce")
    if "FECHA DE CREACION" in df.columns:
        df["FECHA DE CREACION"] = pd.to_datetime(df["FECHA DE CREACION"], errors="coerce")

    # ---------------- Días abierta ----------------
    hoy = pd.Timestamp.today().normalize()
    if "FECHA DE CREACION" in df.columns:
        df["DIAS ABIERTA"] = (hoy - df["FECHA DE CREACION"]).dt.days
    else:
        df["DIAS ABIERTA"] = pd.NA

    # ---------------- Eliminar duplicados ----------------
    if "INTERACCION" in df.columns:
        df = df.drop_duplicates(subset=["SUSCRIPCION", "INTERACCION"])
    else:
        df = df.drop_duplicates(subset=["SUSCRIPCION"])

    # ---------------- Separar hojas ----------------
    df_abiertas = df[df["ESTADO"] != "Completed"]

    df_top_20_abiertas = df_abiertas[
        (df_abiertas["ESTADO"] == "InProgress") &
        (df_abiertas["CATEGORIA"] != "Deactivation")
    ].sort_values(by="DIAS ABIERTA", ascending=False).head(20).copy()

    if "FECHA DE ACTIVACION" in df_top_20_abiertas.columns:
        df_top_20_abiertas = df_top_20_abiertas.drop(columns=["FECHA DE ACTIVACION"])

    df_bajas = df[
        (df["CATEGORIA"] == "Deactivation") &
        (df["ESTADO"] != "Completed")
    ].copy()

    columnas_bajas = [
        "ESTADO", "CATEGORIA", "FECHA DE CREACION", "OFERTA", "SUSCRIPCION",
        "RESPONSABLE", "NOMBRE DEL CLIENTE", "INTERACCION", "MODELO COMERCIAL", "DIAS ABIERTA",
    ]
    df_bajas = df_bajas[[c for c in columnas_bajas if c in df_bajas.columns]].sort_values(
        by="DIAS ABIERTA", ascending=False
    )

    df_activaciones = df[
        (df["ESTADO"] == "Completed") &
        (df["CATEGORIA"] == "SalesOrder")
    ].copy()

    # ---------------- MES en español (para hoja "ACTIVACIONES POR MODELO") ----------------
    meses_es = {
        "January": "ENERO", "February": "FEBRERO", "March": "MARZO",
        "April": "ABRIL", "May": "MAYO", "June": "JUNIO",
        "July": "JULIO", "August": "AGOSTO", "September": "SEPTIEMBRE",
        "October": "OCTUBRE", "November": "NOVIEMBRE", "December": "DICIEMBRE",
    }

    if "FECHA DE CREACION" in df_activaciones.columns:
        df_activaciones["FECHA DE CREACION"] = pd.to_datetime(df_activaciones["FECHA DE CREACION"], errors="coerce")
        df_activaciones["MES"] = df_activaciones["FECHA DE CREACION"].dt.strftime("%B %Y")
        df_activaciones["MES"] = df_activaciones["MES"].apply(
            lambda x: f'{meses_es.get(x.split()[0], x.split()[0])} {x.split()[1]}' if isinstance(x, str) else x
        )
    else:
        df_activaciones["MES"] = pd.NA

    meses = df_activaciones["MES"].dropna().unique()
    meses = sorted(
        meses,
        key=lambda x: pd.to_datetime(
            x.replace("ENERO", "January").replace("FEBRERO", "February").replace("MARZO", "March")
             .replace("ABRIL", "April").replace("MAYO", "May").replace("JUNIO", "June")
             .replace("JULIO", "July").replace("AGOSTO", "August").replace("SEPTIEMBRE", "September")
             .replace("OCTUBRE", "October").replace("NOVIEMBRE", "November").replace("DICIEMBRE", "December"),
            format="%B %Y", errors="coerce",
        ),
        reverse=True,
    )

    df_activadas = df[df["ESTADO"] == "Completed"].copy()

    # ---------------- Exportar a Excel (fechas como TEXTO "DD-MM-AA") ----------------
    DATE_FMT = "%d-%m-%y"

    def df_fechas_a_texto(df_src: pd.DataFrame) -> pd.DataFrame:
        """Devuelve una copia con columnas 'FECHA*' convertidas a string DD-MM-AA (Excel no las reinterpreta)."""
        df_out = df_src.copy()
        for col in df_out.columns:
            if "FECHA" in col.upper():
                s = pd.to_datetime(df_out[col], errors="coerce")
                txt = s.dt.strftime(DATE_FMT)
                txt = txt.where(~s.isna(), "")
                df_out[col] = txt.astype(object)   # asegura que pandas escriba como texto
        return df_out

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_numbers": False}},  # evita auto-conversión
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
            bloque = df_activaciones[df_activaciones["MES"] == mes]
            bloque_txt = df_fechas_a_texto(bloque)

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

            worksheet.write(startrow, 0, mes)
            tabla_mes.reset_index().to_excel(
                writer, sheet_name="ACTIVACIONES POR MODELO", startrow=startrow + 1, index=False
            )
            startrow += len(tabla_mes) + 4
            worksheet.set_column(0, 0, 48)
            worksheet.set_column(1, 10, 13)

    output.seek(0)

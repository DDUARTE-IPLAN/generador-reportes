import streamlit as st
import pandas as pd
from datetime import datetime
import io

col1, col2 = st.columns([1, 5])

with col1:
    st.image("mi_logo.png", use_container_width=True)

with col2:
    st.title("Generador de Reportes Automático")

uploaded_file = st.file_uploader("Subí tu archivo CSV", type="csv")

if uploaded_file is not None:
    nombre_reporte = f"reporte_general_{datetime.today().strftime('%Y-%m-%d')}.xlsx"
    df = pd.read_csv(uploaded_file)

    # Renombrar columnas
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
        "Fecha Activación": "FECHA DE ACTIVACION"
    })

    columnas_a_eliminar = [
        "Order ID", "Party Role ID", "Mail Contacto Técnico", "Instalation Address",
        "Nombre Elemento", "Monto", "Moneda", "Tipo de Precio", "Delta",
        "Fecha Agendamiento", "Motivo Reprogramación", "Motivo", "Segmento", "Fecha Cancelación", "Current Phase"
    ]
    df = df.drop(columns=[col for col in columnas_a_eliminar if col in df.columns])

    df["FECHA DE ACTIVACION"] = pd.to_datetime(df["FECHA DE ACTIVACION"], errors="coerce")
    df["FECHA DE CREACION"] = pd.to_datetime(df["FECHA DE CREACION"], errors="coerce")
    hoy = pd.Timestamp.today().normalize()
    df["DIAS ABIERTA"] = (hoy - df["FECHA DE CREACION"]).dt.days

    if "INTERACCION" in df.columns:
        df = df.drop_duplicates(subset=["SUSCRIPCION", "INTERACCION"])
    else:
        df = df.drop_duplicates(subset=["SUSCRIPCION"])

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
        "RESPONSABLE", "NOMBRE DEL CLIENTE", "INTERACCION", "MODELO COMERCIAL", "DIAS ABIERTA"
    ]

    df_bajas = df_bajas[[col for col in columnas_bajas if col in df_bajas.columns]]
    df_bajas = df_bajas.sort_values(by="DIAS ABIERTA", ascending=False)

    df_activaciones = df[
        (df["ESTADO"] == "Completed") &
        (df["CATEGORIA"] == "SalesOrder")
    ].copy()

    meses_es = {
        'January': 'ENERO', 'February': 'FEBRERO', 'March': 'MARZO',
        'April': 'ABRIL', 'May': 'MAYO', 'June': 'JUNIO',
        'July': 'JULIO', 'August': 'AGOSTO', 'September': 'SEPTIEMBRE',
        'October': 'OCTUBRE', 'November': 'NOVIEMBRE', 'December': 'DICIEMBRE'
    }

    df_activaciones["MES"] = df_activaciones["FECHA DE CREACION"].dt.strftime("%B %Y")
    df_activaciones["MES"] = df_activaciones["MES"].apply(
        lambda x: f'{meses_es.get(x.split()[0], x.split()[0])} {x.split()[1]}' if isinstance(x, str) else x
    )

    meses = df_activaciones["MES"].dropna().unique()
    meses = sorted(meses, key=lambda x: pd.to_datetime(
        x.replace('ENERO', 'January').replace('FEBRERO', 'February').replace('MARZO', 'March')
         .replace('ABRIL', 'April').replace('MAYO', 'May').replace('JUNIO', 'June')
         .replace('JULIO', 'July').replace('AGOSTO', 'August').replace('SEPTIEMBRE', 'September')
         .replace('OCTUBRE', 'October').replace('NOVIEMBRE', 'November').replace('DICIEMBRE', 'December'),
        format='%B %Y', errors='coerce'
    ), reverse=True)

    df_activadas = df[df["ESTADO"] == "Completed"].copy()

    # 🔔 Formatear fechas en todas las hojas antes de exportar
    for df_tmp in [df, df_abiertas, df_top_20_abiertas, df_bajas, df_activadas]:
        if "FECHA DE CREACION" in df_tmp.columns:
            df_tmp["FECHA DE CREACION"] = pd.to_datetime(df_tmp["FECHA DE CREACION"], errors="coerce").dt.strftime("%d/%m/%Y")
        if "FECHA DE ACTIVACION" in df_tmp.columns:
            df_tmp["FECHA DE ACTIVACION"] = pd.to_datetime(df_tmp["FECHA DE ACTIVACION"], errors="coerce").dt.strftime("%d/%m/%Y")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Función auxiliar para ajustar ancho automáticamente
        def export_and_autofit(df_export, sheet_name):
            df_export.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            for i, col in enumerate(df_export.columns):
                column_len = max(df_export[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, column_len)

        export_and_autofit(df, "TODAS LAS ORDENES")
        export_and_autofit(df_abiertas, "ORDENES ABIERTAS")
        export_and_autofit(df_top_20_abiertas, "TOP 20 + ABIERTAS")
        export_and_autofit(df_bajas, "BAJAS")
        export_and_autofit(df_activadas, "ORDENES ACTIVADAS")

        # ACTIVACIONES POR MODELO queda igual (sin autofit especial)
        workbook = writer.book
        worksheet = workbook.add_worksheet("ACTIVACIONES POR MODELO")
        writer.sheets["ACTIVACIONES POR MODELO"] = worksheet

        startrow = 0
        for mes in meses:
            bloque = df_activaciones[df_activaciones["MES"] == mes]
            tabla_mes = pd.pivot_table(
                bloque,
                index="OFERTA",
                columns="MODELO COMERCIAL",
                values="SUSCRIPCION",
                aggfunc="count",
                fill_value=0,
                margins=True,
                margins_name="Suma total"
            )

            worksheet.write(startrow, 0, mes)
            tabla_mes.reset_index().to_excel(writer, sheet_name="ACTIVACIONES POR MODELO", startrow=startrow + 1, index=False)
            startrow += len(tabla_mes) + 4

            worksheet.set_column(0, 0, 48)  
            worksheet.set_column(1, 10, 13)  
        

    output.seek(0)
    st.success("✅ Reporte generado con éxito")
    st.download_button(
        label="📥 Descargar reporte",
        data=output,
        file_name=nombre_reporte,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("📥 Subí un archivo CSV para comenzar.")

import pandas as pd

def procesar_script2(df, output):
    # Asegurarse de que las fechas estén bien formateadas
    if "FECHA DE CREACION" in df.columns:
        df["FECHA DE CREACION"] = pd.to_datetime(df["FECHA DE CREACION"], errors="coerce").dt.strftime("%d/%m/%Y")
    if "FECHA DE ACTIVACION" in df.columns:
        df["FECHA DE ACTIVACION"] = pd.to_datetime(df["FECHA DE ACTIVACION"], errors="coerce").dt.strftime("%d/%m/%Y")

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name="TODAS LAS ORDENES", index=False)

        # Autoajustar columnas
        worksheet = writer.sheets["TODAS LAS ORDENES"]
        for i, col in enumerate(df.columns):
            column_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, column_len)

from __future__ import annotations
import os
from io import BytesIO
from datetime import datetime
from pathlib import Path
import pandas as pd
import streamlit as st

def nombre_salida() -> str:
    return f"reporte_general_{datetime.now().strftime('%d-%m-%Y')}.xlsx"

def leer_fuente(archivo) -> pd.DataFrame:
    nombre = (archivo.name if hasattr(archivo, "name") else str(archivo)).lower()
    if nombre.endswith(".csv"):
        try:
            return pd.read_csv(archivo, sep=None, engine="python")
        except Exception:
            for s in [",", ";", "\t", "|"]:
                try:
                    if hasattr(archivo, "seek"):
                        archivo.seek(0)
                    return pd.read_csv(archivo, sep=s, engine="c", low_memory=False)
                except Exception:
                    continue
            raise
    return pd.read_excel(archivo, engine="openpyxl")

def leer_fuentes_csv_multiples(rutas: list[str]) -> pd.DataFrame:
    frames = []
    for ruta in rutas:
        try:
            try:
                df = pd.read_csv(ruta, sep=None, engine="python", low_memory=False)
            except Exception:
                ok = False
                for s in [",", ";", "\t", "|"]:
                    try:
                        df = pd.read_csv(ruta, sep=s, engine="c", low_memory=False)
                        ok = True; break
                    except Exception:
                        pass
                if not ok: raise
            df["__ORIGEN"] = Path(ruta).name
            frames.append(df)
        except Exception as e:
            st.warning(f"No pude leer {ruta}: {e}")
    if not frames:
        raise RuntimeError("No se pudo leer ninguno de los CSV seleccionados.")
    return pd.concat(frames, axis=0, ignore_index=True, sort=False)

def cargar_hoja_todas_las_ordenes() -> tuple[pd.DataFrame | None, str]:
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

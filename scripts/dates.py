import re
from datetime import date, datetime
import pandas as pd

_ddmmaa = re.compile(r"^\s*(\d{1,2})[/-](\d{1,2})[/-](\d{2})\s*$")

def _try_make_date(Y: int, m: int, d: int):
    from datetime import date
    try: return date(Y, m, d)
    except ValueError: return None

def _split_ddmmaa(val, dias_ref: int | None = None):
    if pd.isna(val): return None, None, None, None
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.year, val.month, val.day, val.year % 100
    s = str(val).strip()
    mobj = _ddmmaa.match(s)
    if mobj:
        a, b, yy = int(mobj.group(1)), int(mobj.group(2)), int(mobj.group(3))
        Y = 2000 + yy; hoy = date.today()
        if a > 12 and b <= 12:
            dt = _try_make_date(Y, b, a); 
            return (Y, dt.month, dt.day, yy) if dt else (None,)*4
        if a <= 12 and b > 12:
            dt = _try_make_date(Y, a, b);
            return (Y, dt.month, dt.day, yy) if dt else (None,)*4
        dt_ddmm = _try_make_date(Y, b, a)
        dt_mmdd = _try_make_date(Y, a, b)
        if not dt_ddmm and not dt_mmdd: return (None,)*4
        if dias_ref is not None and pd.notna(dias_ref):
            best = None
            if dt_ddmm:
                diff1 = abs((hoy - dt_ddmm).days - int(dias_ref)); best = ("ddmm", diff1, dt_ddmm)
            if dt_mmdd:
                diff2 = abs((hoy - dt_mmdd).days - int(dias_ref))
                if best is None or diff2 < best[1]: best = ("mmdd", diff2, dt_mmdd)
            dt = best[2]; return Y, dt.month, dt.day, yy
        if dt_ddmm:
            dt = dt_ddmm if (hoy - dt_ddmm).days >= -1 else (dt_mmdd or dt_ddmm)
        else:
            dt = dt_mmdd
        return Y, dt.month, dt.day, yy
    t = pd.to_datetime(s, dayfirst=True, errors="coerce")
    if pd.isna(t): return (None,)*4
    return t.year, t.month, t.day, t.year % 100

def a_iso_y_display(series: pd.Series, dias_ref_series: pd.Series | None = None):
    iso_vals, disp_vals = [], []
    for i, v in enumerate(series):
        dias_ref = dias_ref_series.iloc[i] if dias_ref_series is not None and i < len(dias_ref_series) else None
        Y, m, d, yy = _split_ddmmaa(v, dias_ref=dias_ref)
        if Y is None:
            iso_vals.append(""); disp_vals.append("")
        else:
            iso_vals.append(f"{Y:04d}-{m:02d}-{d:02d}")
            disp_vals.append(f"{d:02d}-{m:02d}-{yy:02d}")
    return pd.Series(iso_vals, index=series.index), pd.Series(disp_vals, index=series.index)

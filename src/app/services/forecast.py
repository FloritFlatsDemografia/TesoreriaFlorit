import pandas as pd
from pandas.tseries.offsets import BusinessDay
from datetime import date, datetime
from dateutil.relativedelta import relativedelta

def _to_date(x):
    if pd.isna(x):
        return None
    if isinstance(x, date):
        return x
    try:
        return pd.to_datetime(x).date()
    except Exception:
        return None

def _adjust_weekend(d: date, rule: str) -> date:
    # 5=Sat 6=Sun
    if d.weekday() < 5:
        return d
    if "ANT" in rule:
        # move back to Friday
        while d.weekday() >= 5:
            d = d - pd.Timedelta(days=1)
        return d
    # default SIG HABIL: move forward to Monday
    while d.weekday() >= 5:
        d = d + pd.Timedelta(days=1)
    return d

def _add_business_days(d: date, n: int) -> date:
    if n <= 0:
        return d
    return (pd.Timestamp(d) + BusinessDay(n)).date()

def _parse_valor_fecha(regla_fecha: str, valor_raw):
    regla = (regla_fecha or "").upper().strip()
    if regla == "FECHA_FIJA":
        d = _to_date(valor_raw)
        if d is None:
            raise ValueError(f"FECHA_FIJA requiere una fecha válida en VALOR_FECHA. Valor: {valor_raw}")
        return ("fecha_fija", d)
    if regla == "ULTIMO_HABIL":
        return ("ultimo_habil", None)
    # DIA_MES default
    # valor_raw might be a datetime (Excel day=1 stored as 2026-01-01)
    if isinstance(valor_raw, (pd.Timestamp, datetime, date)):
        day = pd.to_datetime(valor_raw).day
    else:
        day = int(float(valor_raw))
    if day < 1 or day > 31:
        raise ValueError(f"DIA_MES requiere día 1-31. Valor: {valor_raw}")
    return ("dia_mes", day)

def _month_last_business_day(y, m):
    d = date(y, m, 1) + relativedelta(months=1, days=-1)
    return _adjust_weekend(d, "ANT_HABIL")

def generate_recurrent_events(catalog_df: pd.DataFrame, start_date: date, horizon_months: int = 12) -> pd.DataFrame:
    start = pd.Timestamp(start_date).date()
    end = (pd.Timestamp(start_date) + relativedelta(months=horizon_months)).date()

    rows = []
    for _, r in catalog_df.iterrows():
        concepto = str(r["concepto"]).strip()
        if concepto == "" or concepto.lower() == "nan":
            continue
        tipo = str(r["tipo"]).upper().strip()
        dept = str(r["departamento"]).upper().strip()
        periodicidad = str(r["periodicidad"]).upper().strip()
        regla_fecha = str(r["regla_fecha"]).upper().strip()
        valor_raw = r["valor_fecha_raw"]
        lag = int(r.get("lag", 0) or 0)
        ajuste = str(r.get("ajuste_finde", "SIG HABIL")).upper().strip()
        importe = r.get("importe", None)
        if pd.isna(importe):
            # allow missing; skip for now (user can fill later)
            importe = 0.0
        else:
            importe = float(importe)

        kind, parsed = _parse_valor_fecha(regla_fecha, valor_raw)

        # Determine schedule dates
        if periodicidad == "PUNTUAL":
            # single event
            if kind == "fecha_fija":
                d = parsed
            elif kind == "ultimo_habil":
                d = _month_last_business_day(start.year, start.month)
            else:
                d = date(start.year, start.month, parsed)
            d = _adjust_weekend(d, ajuste)
            d = _add_business_days(d, lag)
            if start <= d <= end:
                rows.append((d, concepto, tipo, dept, importe, r))
            continue

        if periodicidad == "ANUAL":
            if kind == "fecha_fija":
                d0 = parsed
                # generate yearly occurrences within range
                y = start.year
                while True:
                    d = date(y, d0.month, d0.day)
                    d = _adjust_weekend(d, ajuste)
                    d = _add_business_days(d, lag)
                    if d > end:
                        break
                    if d >= start:
                        rows.append((d, concepto, tipo, dept, importe, r))
                    y += 1
            else:
                # annual on day-of-month in same month as start? fallback: start month
                m = start.month
                y = start.year
                while True:
                    if kind == "ultimo_habil":
                        d = _month_last_business_day(y, m)
                    else:
                        d = date(y, m, parsed)
                    d = _adjust_weekend(d, ajuste)
                    d = _add_business_days(d, lag)
                    if d > end:
                        break
                    if d >= start:
                        rows.append((d, concepto, tipo, dept, importe, r))
                    y += 1
            continue

        if periodicidad in ("MENSUAL","TRIMESTRAL"):
            step = 1 if periodicidad == "MENSUAL" else 3
            cur = date(start.year, start.month, 1)
            while cur <= end:
                y, m = cur.year, cur.month
                if kind == "ultimo_habil":
                    d = _month_last_business_day(y, m)
                else:
                    # clamp day to last day if month shorter
                    day = parsed
                    last = (pd.Timestamp(date(y,m,1)) + relativedelta(months=1, days=-1)).date().day
                    day = min(day, last)
                    d = date(y, m, day)
                d = _adjust_weekend(d, ajuste)
                d = _add_business_days(d, lag)
                if start <= d <= end:
                    rows.append((d, concepto, tipo, dept, importe, r))
                cur = (pd.Timestamp(cur) + relativedelta(months=step)).date()
            continue

        # default: treat as mensual
        cur = date(start.year, start.month, 1)
        while cur <= end:
            y, m = cur.year, cur.month
            if kind == "ultimo_habil":
                d = _month_last_business_day(y, m)
            else:
                last = (pd.Timestamp(date(y,m,1)) + relativedelta(months=1, days=-1)).date().day
                day = min(parsed, last)
                d = date(y, m, day)
            d = _adjust_weekend(d, ajuste)
            d = _add_business_days(d, lag)
            if start <= d <= end:
                rows.append((d, concepto, tipo, dept, importe, r))
            cur = (pd.Timestamp(cur) + relativedelta(months=1)).date()

    if not rows:
        return pd.DataFrame(columns=["fecha","concepto","tipo","departamento","importe","cobro","pago"])

    out_rows=[]
    for d, concepto, tipo, dept, importe, r in rows:
        cobro = importe if tipo == "INGRESO" else 0.0
        pago = importe if tipo == "GASTO" else 0.0
        out_rows.append({
            "fecha": d,
            "concepto": concepto,
            "tipo": tipo,
            "departamento": dept,
            "importe": float(importe),
            "cobro": float(cobro),
            "pago": float(pago),
            "naturaleza": r.get("naturaleza", None),
            "periodicidad": r.get("periodicidad", None),
            "iva_en_factura": r.get("iva_en_factura", None),
            "iva_pct": r.get("iva_pct", None),
            "iva_sentido": r.get("iva_sentido", None),
            "tratamiento_iva": r.get("tratamiento_iva", None),
            "impuesto_tipo": r.get("impuesto_tipo", None),
            "modelo": r.get("modelo", None),
        })
    df = pd.DataFrame(out_rows).sort_values(["fecha","concepto"]).reset_index(drop=True)
    return df

def consolidate_and_compute_balance(events_df: pd.DataFrame, import_df: pd.DataFrame | None, saldo_inicial: float) -> pd.DataFrame:
    frames = [events_df.copy()]
    if import_df is not None and len(import_df):
        imp = import_df.copy()
        # normalize
        if "importe" in imp.columns and ("cobro" not in imp.columns and "pago" not in imp.columns):
            imp["cobro"] = imp.apply(lambda r: float(r["importe"]) if str(r.get("tipo","")).upper()=="INGRESO" else 0.0, axis=1)
            imp["pago"] = imp.apply(lambda r: float(r["importe"]) if str(r.get("tipo","")).upper()=="GASTO" else 0.0, axis=1)
        for c in ["departamento","tipo","concepto"]:
            if c not in imp.columns:
                imp[c] = None
        imp["fecha"] = pd.to_datetime(imp["fecha"]).dt.date
        imp["importe"] = (imp.get("importe", imp["cobro"]+imp["pago"])).astype(float)
        frames.append(imp[events_df.columns.intersection(imp.columns)].copy() if len(events_df.columns) else imp.copy())

    df = pd.concat(frames, ignore_index=True, sort=False)
    df = df.dropna(subset=["fecha"]).sort_values(["fecha","tipo","concepto"]).reset_index(drop=True)

    # Ensure cobro/pago columns
    if "cobro" not in df.columns:
        df["cobro"] = 0.0
    if "pago" not in df.columns:
        df["pago"] = 0.0
    df["cobro"] = pd.to_numeric(df["cobro"], errors="coerce").fillna(0.0)
    df["pago"] = pd.to_numeric(df["pago"], errors="coerce").fillna(0.0)

    df["saldo"] = saldo_inicial + (df["cobro"] - df["pago"]).cumsum()
    return df

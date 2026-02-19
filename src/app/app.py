import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from datetime import datetime, date, timedelta

# -----------------------------
# Helpers
# -----------------------------
def norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def normalize_pagado(v) -> bool:
    if v is None:
        return False
    if isinstance(v, float) and np.isnan(v):
        return False
    s = str(v).strip().lower()
    return s in {"✓", "✅", "x", "si", "sí", "true", "1", "pagado", "ok"}

def to_date_safe(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        d = pd.to_datetime(v, dayfirst=True, errors="coerce")
        if pd.isna(d):
            return None
        return d.date()
    except Exception:
        return None

def next_business_day(d: date) -> date:
    while d.weekday() >= 5:
        d += timedelta(days=1)
    return d

def euro_fmt(x: float) -> str:
    s = f"{x:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".") + " €"

# -----------------------------
# Lectura hoja BANCOS
# -----------------------------
def read_bancos_values(excel_file) -> dict:
    wb = openpyxl.load_workbook(excel_file, data_only=True)

    if "BANCOS" not in wb.sheetnames:
        raise ValueError("No existe la hoja 'BANCOS' en el Excel.")

    ws = wb["BANCOS"]

    mapping = {}
    for r in range(1, ws.max_row + 1):
        k = ws.cell(r, 1).value
        v = ws.cell(r, 2).value
        if k is None:
            continue
        mapping[str(k).strip().upper()] = v

    def get_num(key: str):
        v = mapping.get(key)
        if v is None:
            return None
        try:
            return float(v)
        except Exception:
            return None

    total_bancos = get_num("TOTAL BANCOS")
    suplidos = get_num("CUENTA SUPLIDOS")
    efectivo = get_num("CUENTA DE EFECTIVO")

    if total_bancos is None:
        raise ValueError("No he podido leer un número en 'TOTAL BANCOS' (hoja BANCOS).")

    return {
        "total_bancos": total_bancos,
        "cuenta_suplidos": suplidos,
        "cuenta_efectivo": efectivo,
    }

# -----------------------------
# Construcción series
# -----------------------------
def build_series(df: pd.DataFrame, saldo_inicial: float, horizon_months: int):
    df = norm_cols(df)

    required = ["TIPO", "IMPORTE PRONOSTICADO", "IMPORTE REAL", "VALOR_FECHA"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas obligatorias en CATALOGO_RECURRENTE: {missing}")

    # Columnas opcionales
    if "Pagado" not in df.columns:
        df["Pagado"] = ""
    if "Fecha" not in df.columns:
        df["Fecha"] = pd.NaT
    if "LAG" not in df.columns:
        df["LAG"] = 0
    if "AJUSTE FINDE" not in df.columns:
        df["AJUSTE FINDE"] = ""

    # Normalizaciones
    df["TIPO"] = df["TIPO"].astype(str).str.strip().str.upper()
    df["IMPORTE PRONOSTICADO"] = pd.to_numeric(df["IMPORTE PRONOSTICADO"], errors="coerce").fillna(0.0)
    df["IMPORTE REAL"] = pd.to_numeric(df["IMPORTE REAL"], errors="coerce").fillna(0.0)

    df["Pagado_norm"] = df["Pagado"].apply(normalize_pagado)
    df["Fecha_pago"] = df["Fecha"].apply(to_date_safe)

    df["Fecha_plan_base"] = pd.to_datetime(df["VALOR_FECHA"], errors="coerce").dt.date
    df["LAG"] = pd.to_numeric(df["LAG"], errors="coerce").fillna(0).astype(int)
    df["AJUSTE FINDE"] = df["AJUSTE FINDE"].astype(str).str.strip().str.upper()

    # Índice diario
    today = date.today()
    start = date(today.year, today.month, 1)
    end = (pd.Timestamp(start) + pd.DateOffset(months=horizon_months)).date()
    idx = pd.date_range(start=start, end=end, freq="D")

    # PRONOSTICADO: solo pendientes
    pending = df[~df["Pagado_norm"]].copy()

    def calc_plan_date(row):
        d = row["Fecha_plan_base"]
        if d is None or pd.isna(d):
            return None
        d2 = d + timedelta(days=int(row["LAG"]))
        if row["AJUSTE FINDE"] in {"SIG HABIL", "SIGUIENTE HABIL", "SIG_HABIL", "NEXT_BUSINESS"}:
            d2 = next_business_day(d2)
        return d2

    pending["Fecha_plan"] = pending.apply(calc_plan_date, axis=1)
    pending["signed_plan"] = np.where(
        pending["TIPO"].eq("INGRESO"),
        pending["IMPORTE PRONOSTICADO"],
        -pending["IMPORTE PRONOSTICADO"],
    )

    plan_flows = (
        pending.dropna(subset=["Fecha_plan"])
        .groupby("Fecha_plan")["signed_plan"]
        .sum()
    )

    plan_daily = pd.Series(0.0, index=idx)
    for d, v in plan_flows.items():
        ts = pd.Timestamp(d)
        if ts in plan_daily.index:
            plan_daily.loc[ts] += float(v)

    saldo_plan = (saldo_inicial + plan_daily.cumsum()).rename("PRONOSTICADO")

    # REAL: solo pagados con fecha
    paid = df[df["Pagado_norm"] & df["Fecha_pago"].notna()].copy()
    paid["signed_real"] = np.where(
        paid["TIPO"].eq("INGRESO"),
        paid["IMPORTE REAL"],
        -paid["IMPORTE REAL"],
    )

    real_flows = paid.groupby("Fecha_pago")["signed_real"].sum()

    real_daily = pd.Series(0.0, index=idx)
    for d, v in real_flows.items():
        ts = pd.Timestamp(d)
        if ts in real_daily.index:
            real_daily.loc[ts] += float(v)

    saldo_real = (saldo_inicial + real_daily.cumsum()).rename("REAL")

    return saldo_plan, saldo_real

# -----------------------------
# UI
# -----------------------------
st.set_page_config(page_title="Tesorería", layout="wide")
st.title("APP Tesorería")

uploaded = st.file_uploader("Sube tu Excel de Tesorería", type=["xlsx"])
if not uploaded:
    st.info("Sube el Excel para cargar CATALOGO_RECURRENTE y BANCOS.")
    st.stop()

# Leer BANCOS
bancos = read_bancos_values(uploaded)

c1, c2, c3 = st.columns(3)
c1.metric("TOTAL BANCOS (saldo inicial)", euro_fmt(bancos["total_bancos"]))
c2.metric("CUENTA SUPLIDOS (línea fija)", euro_fmt(bancos["cuenta_suplidos"]) if bancos["cuenta_suplidos"] is not None else "No definido")
c3.metric("CUENTA DE EFECTIVO (línea fija)", euro_fmt(bancos["cuenta_efectivo"]) if bancos["cuenta_efectivo"] is not None else "No definido")

horizon_months = st.slider("Horizonte forecast (meses)", 1, 24, 12)

# Leer catálogo
df = pd.read_excel(uploaded, sheet_name="CATALOGO_RECURRENTE")

# Series
saldo_plan, saldo_real = build_series(df, bancos["total_bancos"], horizon_months)

# DataFrame para gráfica
chart_df = pd.concat([saldo_plan, saldo_real], axis=1)

# Líneas fijas (constantes) si existen
if bancos["cuenta_suplidos"] is not None:
    chart_df["CUENTA SUPLIDOS"] = bancos["cuenta_suplidos"]
if bancos["cuenta_efectivo"] is not None:
    chart_df["CUENTA DE EFECTIVO"] = bancos["cuenta_efectivo"]

st.subheader("Evolución de saldo: Real vs Pronosticado")
st.line_chart(chart_df)

with st.expander("Comprobación Pagado/Fecha"):
    df_dbg = norm_cols(df)
    if "Pagado" in df_dbg.columns:
        df_dbg["Pagado_norm"] = df_dbg["Pagado"].apply(normalize_pagado)
    if "Fecha" in df_dbg.columns:
        df_dbg["Fecha_pago"] = df_dbg["Fecha"].apply(to_date_safe)

    cols_show = [c for c in ["GENERAL", "TIPO", "IMPORTE PRONOSTICADO", "IMPORTE REAL", "Pagado", "Fecha"] if c in df_dbg.columns]
    st.dataframe(df_dbg[cols_show].head(100))

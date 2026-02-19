import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from datetime import datetime, date, timedelta
import matplotlib.pyplot as plt

# -----------------------------
# Utilidades
# -----------------------------
def _normalize_pagado(v) -> bool:
    """Acepta ✓ / ✅ / X / TRUE / 1 / SI / etc."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return False
    s = str(v).strip().lower()
    return s in {"✓", "✅", "x", "si", "sí", "true", "1", "pagado", "ok"}

def _to_date(v):
    if pd.isna(v) or v is None:
        return None
    if isinstance(v, (datetime, date)):
        return v.date() if isinstance(v, datetime) else v
    # intenta parse
    try:
        return pd.to_datetime(v, dayfirst=True).date()
    except Exception:
        return None

def _next_business_day(d: date) -> date:
    # lunes=0..domingo=6
    while d.weekday() >= 5:
        d += timedelta(days=1)
    return d

def read_bancos_values(excel_file) -> dict:
    """
    Lee hoja BANCOS y devuelve:
    - total_bancos
    - cuenta_suplidos
    - cuenta_efectivo
    """
    wb = openpyxl.load_workbook(excel_file, data_only=True)
    if "BANCOS" not in wb.sheetnames:
        raise ValueError("No existe la hoja 'BANCOS' en el Excel.")

    ws = wb["BANCOS"]

    # Lee pares (A,B) y mapea por etiqueta
    mapping = {}
    for r in range(1, ws.max_row + 1):
        k = ws.cell(r, 1).value
        v = ws.cell(r, 2).value
        if k is None:
            continue
        k_norm = str(k).strip().upper()
        mapping[k_norm] = v

    def get_num(key):
        v = mapping.get(key)
        if v is None:
            return None
        try:
            return float(v)
        except Exception:
            return None

    total_bancos = get_num("TOTAL BANCOS")
    cuenta_suplidos = get_num("CUENTA SUPLIDOS")
    cuenta_efectivo = get_num("CUENTA DE EFECTIVO")

    if total_bancos is None:
        raise ValueError("No he podido leer el valor numérico de 'TOTAL BANCOS' en la hoja BANCOS.")

    # Si no existen las cuentas, las dejamos en None
    return {
        "total_bancos": total_bancos,
        "cuenta_suplidos": cuenta_suplidos,
        "cuenta_efectivo": cuenta_efectivo
    }

def build_daily_series(df: pd.DataFrame, start_balance: float, horizon_months: int):
    """
    Construye series diarias de:
    - Saldo pronosticado (solo movimientos NO pagados; usa IMPORTE PRONOSTICADO y VALOR_FECHA)
    - Saldo real (solo movimientos pagados con Fecha; usa IMPORTE REAL y Fecha)
    Asume columnas:
      GENERAL , Pagado, Fecha, TIPO, IMPORTE PRONOSTICADO, IMPORTE REAL, VALOR_FECHA, LAG, AJUSTE FINDE
    """
    today = date.today()

    # Horizonte
    start = date(today.year, today.month, 1)
    # fin = start + horizon_months meses aprox
    end = (pd.Timestamp(start) + pd.DateOffset(months=horizon_months)).date()

    days = pd.date_range(start=start, end=end, freq="D")
    idx = pd.Index(days)

    # Normaliza columnas esperadas
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]

    # nombres esperados (en tu excel hay "GENERAL " con espacio final)
    col_general = "GENERAL" if "GENERAL" in df.columns else "GENERAL"
    # realmente tu archivo trae "GENERAL" ya tras strip, queda "GENERAL"
    # ok.

    # Pagado
    if "Pagado" not in df.columns:
        df["Pagado"] = False
    df["Pagado_norm"] = df["Pagado"].apply(_normalize_pagado)

    # Fecha pago real
    if "Fecha" not in df.columns:
        df["Fecha"] = pd.NaT
    df["Fecha_pago"] = df["Fecha"].apply(_to_date)

    # Fecha base plan
    if "VALOR_FECHA" not in df.columns:
        raise ValueError("Falta la columna 'VALOR_FECHA' en CATALOGO_RECURRENTE.")
    df["Fecha_plan"] = pd.to_datetime(df["VALOR_FECHA"], errors="coerce").dt.date

    # LAG (días)
    if "LAG" not in df.columns:
        df["LAG"] = 0
    df["LAG"] = pd.to_numeric(df["LAG"], errors="coerce").fillna(0).astype(int)

    # AJUSTE FINDE
    if "AJUSTE FINDE" not in df.columns:
        df["AJUSTE FINDE"] = ""
    df["AJUSTE FINDE"] = df["AJUSTE FINDE"].astype(str).str.strip().str.upper()

    # TIPO: INGRESO/GASTO
    if "TIPO" not in df.columns:
        raise ValueError("Falta la columna 'TIPO' (INGRESO/GASTO).")
    df["TIPO"] = df["TIPO"].astype(str).str.strip().str.upper()

    # Importes
    for c in ["IMPORTE PRONOSTICADO", "IMPORTE REAL"]:
        if c not in df.columns:
            raise ValueError(f"Falta la columna '{c}'.")
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    # -----------------------------
    # Flujos pronosticados (pendientes)
    # -----------------------------
    pending = df[~df["Pagado_norm"]].copy()

    # Construye fecha efectiva plan: Fecha_plan + LAG y ajuste finde
    def _calc_plan_date(row):
        d = row["Fecha_plan"]
        if d is None or pd.isna(d):
            return None
        d2 = d + timedelta(days=int(row["LAG"]))
        if row["AJUSTE FINDE"] in {"SIG HABIL", "SIGUIENTE HABIL", "SIG_HABIL"}:
            d2 = _next_business_day(d2)
        return d2

    pending["Fecha_efectiva_plan"] = pending.apply(_calc_plan_date, axis=1)

    # signo según tipo
    pending["Importe_plan_signed"] = np.where(
        pending["TIPO"].eq("INGRESO"),
        pending["IMPORTE PRONOSTICADO"],
        -pending["IMPORTE PRONOSTICADO"]
    )

    plan_flows = (
        pending.dropna(subset=["Fecha_efectiva_plan"])
        .groupby("Fecha_efectiva_plan")["Importe_plan_signed"]
        .sum()
    )

    plan_daily = pd.Series(0.0, index=idx)
    for d, v in plan_flows.items():
        ts = pd.Timestamp(d)
        if ts in plan_daily.index:
            plan_daily.loc[ts] += float(v)

    saldo_plan = (start_balance + plan_daily.cumsum()).rename("Saldo pronosticado")

    # -----------------------------
    # Flujos reales (pagados)
    # -----------------------------
    paid = df[df["Pagado_norm"] & df["Fecha_pago"].notna()].copy()
    paid["Importe_real_signed"] = np.where(
        paid["TIPO"].eq("INGRESO"),
        paid["IMPORTE REAL"],
        -paid["IMPORTE REAL"]
    )

    real_flows = paid.groupby("Fecha_pago")["Importe_real_signed"].sum()

    real_daily = pd.Series(0.0, index=idx)
    for d, v in real_flows.items():
        ts = pd.Timestamp(d)
        if ts in real_daily.index:
            real_daily.loc[ts] += float(v)

    saldo_real = (start_balance + real_daily.cumsum()).rename("Saldo real")

    return saldo_plan, saldo_real

# -----------------------------
# UI
# -----------------------------
st.set_page_config(page_title="Tesorería", layout="wide")
st.title("APP Tesorería")

uploaded = st.file_uploader("Sube tu Excel de Tesorería", type=["xlsx"])

if uploaded is None:
    st.info("Sube el Excel para cargar la hoja CATALOGO_RECURRENTE y BANCOS.")
    st.stop()

# Lee BANCOS
try:
    bancos = read_bancos_values(uploaded)
except Exception as e:
    st.error(f"Error leyendo hoja BANCOS: {e}")
    st.stop()

# Muestra saldo actual automáticamente (sin input manual)
col1, col2, col3 = st.columns(3)
col1.metric("TOTAL BANCOS (saldo inicial)", f"{bancos['total_bancos']:,.2f} €".replace(",", "X").replace(".", ",").replace("X", "."))

if bancos["cuenta_suplidos"] is not None:
    col2.metric("CUENTA SUPLIDOS (línea fija)", f"{bancos['cuenta_suplidos']:,.2f} €".replace(",", "X").replace(".", ",").replace("X", "."))
else:
    col2.metric("CUENTA SUPLIDOS (línea fija)", "No definido")

if bancos["cuenta_efectivo"] is not None:
    col3.metric("CUENTA DE EFECTIVO (línea fija)", f"{bancos['cuenta_efectivo']:,.2f} €".replace(",", "X").replace(".", ",").replace("X", "."))
else:
    col3.metric("CUENTA DE EFECTIVO (línea fija)", "No definido")

# Horizonte
horizon_months = st.slider("Horizonte forecast (meses)", min_value=1, max_value=24, value=12)

# Lee catálogo
try:
    df_cat = pd.read_excel(uploaded, sheet_name="CATALOGO_RECURRENTE")
except Exception as e:
    st.error(f"Error leyendo CATALOGO_RECURRENTE: {e}")
    st.stop()

# Construye series
try:
    saldo_plan, saldo_real = build_daily_series(df_cat, bancos["total_bancos"], horizon_months)
except Exception as e:
    st.error(f"Error construyendo series: {e}")
    st.stop()

# Plot
fig, ax = plt.subplots()
ax.plot(saldo_plan.index, saldo_plan.values, label="Pronosticado")
ax.plot(saldo_real.index, saldo_real.values, label="Real")

# Líneas fijas
if bancos["cuenta_suplidos"] is not None:
    ax.axhline(bancos["cuenta_suplidos"], linestyle="--", linewidth=1.5, color="orange", label="Cuenta suplidos")
if bancos["cuenta_efectivo"] is not None:
    ax.axhline(bancos["cuenta_efectivo"], linestyle="--", linewidth=1.5, color="red", label="Cuenta efectivo")

ax.set_title("Evolución saldo: Real vs Pronosticado")
ax.set_xlabel("Fecha")
ax.set_ylabel("€")
ax.legend()
ax.grid(True, alpha=0.2)

st.pyplot(fig)

# Vista rápida de pendientes/pagados (opcional)
with st.expander("Ver movimientos (pendientes vs pagados)"):
    df_view = df_cat.copy()
    df_view.columns = [c.strip() for c in df_view.columns]
    if "Pagado" in df_view.columns:
        df_view["Pagado_norm"] = df_view["Pagado"].apply(_normalize_pagado)
        st.write("Pendientes")
        st.dataframe(df_view[~df_view["Pagado_norm"]].head(50))
        st.write("Pagados")
        st.dataframe(df_view[df_view["Pagado_norm"]].head(50))
    else:
        st.dataframe(df_view.head(50))

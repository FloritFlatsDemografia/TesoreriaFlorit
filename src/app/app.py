import re
from datetime import date
from io import BytesIO

import pandas as pd
import streamlit as st
import altair as alt
import openpyxl

# -----------------------------
# Config
# -----------------------------
st.set_page_config(page_title="APP Tesorería", layout="wide")
st.title("APP Tesorería — Dashboard")

# -----------------------------
# Helpers
# -----------------------------
MIN_REQUIRED = [
    "GENERAL", "TIPO", "DEPARTAMENTO",
    "NATURALEZA", "PERIODICIDAD",
    "REGLA_FECHA", "VALOR_FECHA",
    "LAG", "AJUSTE FINDE"
]

def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df

def find_header_row(df_raw: pd.DataFrame) -> int | None:
    for i in range(min(80, len(df_raw))):
        row = df_raw.iloc[i].astype(str).str.strip().str.upper().tolist()
        if "GENERAL" in row and "TIPO" in row:
            return i
    return None

def is_pagado(v) -> bool:
    """Acepta ✓ / ✅ / X / SI / TRUE / 1 / OK / PAGADO (case-insensitive)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return False
    s = str(v).strip().lower()
    return s in {"✓", "✅", "x", "si", "sí", "true", "1", "ok", "pagado", "y", "yes"}

def eur(x):
    try:
        return f"{float(x):,.2f} €"
    except Exception:
        return ""

def color_saldo(v):
    try:
        v = float(v)
    except Exception:
        return ""
    if v > 0:
        return "color: green; font-weight: 700;"
    if v < 0:
        return "color: red; font-weight: 700;"
    return ""

def estado_cobro_pago(tipo: str, pagado_bool: bool) -> str:
    t = (tipo or "").strip().upper()
    if not pagado_bool:
        return "Pendiente"
    if t == "INGRESO":
        return "Cobrado"
    if t == "GASTO":
        return "Pagado"
    return "Pagado"

def style_estado_cell(val: str):
    v = str(val or "").strip().upper()
    if v == "PENDIENTE":
        return "background-color: #FFE6CC; font-weight: 700;"  # naranja clarito
    if v == "COBRADO":
        return "background-color: #C6EFCE; font-weight: 700;"  # verde
    if v == "PAGADO":
        return "background-color: #FFC7CE; font-weight: 700;"  # rojo
    return ""

# -----------------------------
# Leer hoja BANCOS
# -----------------------------
def read_bancos_from_excel(uploaded_file) -> dict:
    """
    Lee hoja 'BANCOS' y devuelve:
      - total_bancos
      - cuenta_suplidos (opcional)
      - cuenta_efectivo (opcional)
    """
    uploaded_file.seek(0)
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    sheet_name = None
    for name in wb.sheetnames:
        if str(name).strip().upper() == "BANCOS":
            sheet_name = name
            break
    if sheet_name is None:
        raise ValueError("No encuentro la hoja 'BANCOS' en el Excel.")

    ws = wb[sheet_name]

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
        raise ValueError("En hoja BANCOS no puedo leer un número en 'TOTAL BANCOS' (columna €).")

    return {
        "total_bancos": total_bancos,
        "cuenta_suplidos": suplidos,
        "cuenta_efectivo": efectivo
    }

# -----------------------------
# Leer catálogo
# -----------------------------
def read_catalog_from_excel(uploaded_file) -> pd.DataFrame:
    uploaded_file.seek(0)
    raw = pd.read_excel(uploaded_file, sheet_name=0, header=None, engine="openpyxl")
    header_idx = find_header_row(raw)
    if header_idx is None:
        raise ValueError("No encuentro la fila de cabecera (debe contener 'GENERAL' y 'TIPO').")

    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, sheet_name=0, header=header_idx, engine="openpyxl")
    df = normalize_cols(df)

    missing = [c for c in MIN_REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas requeridas: {missing}. Columnas detectadas: {list(df.columns)}")

    # PRORRATEO (opcional)
    if "PRORRATEO" not in df.columns:
        df["PRORRATEO"] = ""
    df["PRORRATEO"] = df["PRORRATEO"].astype(str).str.strip().str.upper()

    # PAGADO + FECHA (opcionales)
    if "PAGADO" not in df.columns:
        df["PAGADO"] = ""
    if "FECHA" in df.columns:
        df["FECHA_PAGO"] = pd.to_datetime(df["FECHA"], errors="coerce").dt.normalize()
    else:
        df["FECHA_PAGO"] = pd.NaT
    df["PAGADO_BOOL"] = df["PAGADO"].apply(is_pagado)

    # Importes
    if "IMPORTE PRONOSTICADO" not in df.columns:
        if "IMPORTE_PRONOSTICADO" in df.columns:
            df["IMPORTE PRONOSTICADO"] = df["IMPORTE_PRONOSTICADO"]
        else:
            raise ValueError("Falta columna: 'IMPORTE PRONOSTICADO'")

    if "IMPORTE REAL" not in df.columns:
        if "IMPORTE_REAL" in df.columns:
            df["IMPORTE REAL"] = df["IMPORTE_REAL"]
        else:
            df["IMPORTE REAL"] = 0.0

    # HASTA (opcional)
    if "HASTA" not in df.columns:
        df["HASTA"] = pd.NaT
    df["HASTA"] = pd.to_datetime(df["HASTA"], errors="coerce").dt.normalize()

    df = df.dropna(how="all").copy()

    for c in ["GENERAL", "TIPO", "DEPARTAMENTO", "NATURALEZA", "PERIODICIDAD", "REGLA_FECHA", "AJUSTE FINDE"]:
        df[c] = df[c].astype(str).str.strip()

    df["TIPO"] = df["TIPO"].str.upper()
    df["DEPARTAMENTO"] = df["DEPARTAMENTO"].str.upper()
    df["NATURALEZA"] = df["NATURALEZA"].str.upper()
    df["PERIODICIDAD"] = df["PERIODICIDAD"].str.upper()
    df["REGLA_FECHA"] = df["REGLA_FECHA"].str.upper()
    df["AJUSTE FINDE"] = df["AJUSTE FINDE"].str.upper()

    df["IMPORTE_PRON"] = pd.to_numeric(df["IMPORTE PRONOSTICADO"], errors="coerce").fillna(0.0)
    df["IMPORTE_REAL"] = pd.to_numeric(df["IMPORTE REAL"], errors="coerce").fillna(0.0)

    df["LAG"] = pd.to_numeric(df["LAG"], errors="coerce").fillna(0).astype(int)

    def to_day_of_month(v):
        if pd.isna(v):
            return None
        if isinstance(v, pd.Timestamp):
            return int(v.day)
        if hasattr(v, "day"):
            try:
                return int(v.day)
            except Exception:
                pass
        s = str(v).strip()
        if re.fullmatch(r"\d{1,2}", s):
            d = int(s)
            if 1 <= d <= 31:
                return d
        return None

    df["DIA_MES"] = df["VALOR_FECHA"].apply(to_day_of_month)
    df["FECHA_FIJA"] = pd.to_datetime(df["VALOR_FECHA"], errors="coerce").dt.normalize()

    return df

def next_business_day(d: pd.Timestamp) -> pd.Timestamp:
    if d.weekday() == 5:
        return d + pd.Timedelta(days=2)
    if d.weekday() == 6:
        return d + pd.Timedelta(days=1)
    return d

def add_business_days(d: pd.Timestamp, n: int) -> pd.Timestamp:
    cur = d
    step = 1 if n >= 0 else -1
    remaining = abs(n)
    while remaining > 0:
        cur = cur + pd.Timedelta(days=step)
        if cur.weekday() < 5:
            remaining -= 1
    return cur

def months_step_from_periodicidad(periodicidad: str) -> int:
    p = (periodicidad or "").upper().strip()
    if p == "MENSUAL":
        return 1
    if p in ("BIMESTRAL", "BIMENSUAL"):
        return 2
    if p == "TRIMESTRAL":
        return 3
    if p == "SEMESTRAL":
        return 6
    return 1

def generate_events_from_catalog(catalog: pd.DataFrame, start_date: pd.Timestamp, months_horizon: int) -> pd.DataFrame:
    """
    Soporta PRORRATEO=DIARIO para filas MENSUALES/BIMESTRALES/...:
    - En vez de 1 evento al mes, genera 1 evento por día del mes (días naturales),
      repartiendo IMPORTE_PRON e IMPORTE_REAL entre los días del mes.
    """
    end_date = (start_date + pd.offsets.MonthBegin(months_horizon + 1)).normalize()
    rows = []

    for _, r in catalog.iterrows():
        periodicidad = str(r.get("PERIODICIDAD", "")).upper().strip()
        regla = str(r.get("REGLA_FECHA", "")).upper().strip()
        ajuste = str(r.get("AJUSTE FINDE", "")).upper().strip()
        lag = int(r.get("LAG", 0))
        prorrateo = str(r.get("PRORRATEO", "")).upper().strip()

        hasta = r.get("HASTA", pd.NaT)
        hasta = pd.Timestamp(hasta).normalize() if not pd.isna(hasta) else pd.NaT

        def apply_adjustments(d: pd.Timestamp) -> pd.Timestamp:
            if ajuste == "SIG_HABIL":
                d = next_business_day(d)
            if lag:
                d = add_business_days(d, lag)
            return d

        def within_limits_base(d_base: pd.Timestamp) -> bool:
            if d_base < start_date or d_base > end_date:
                return False
            if not pd.isna(hasta) and d_base > hasta:
                return False
            return True

        def within_limits_adjusted(d_adj: pd.Timestamp) -> bool:
            return (d_adj >= start_date) and (d_adj <= end_date)

        def add_if_valid(d_base: pd.Timestamp):
            d_base = pd.Timestamp(d_base).normalize()
            if within_limits_base(d_base):
                d_adj = apply_adjustments(d_base)
                if within_limits_adjusted(d_adj):
                    rows.append((d_adj, r))

        def add_prorrateo_diario_for_month(d_base: pd.Timestamp):
            d_base = pd.Timestamp(d_base).normalize()
            y, m = d_base.year, d_base.month
            month_start = pd.Timestamp(year=y, month=m, day=1).normalize()
            month_end = (month_start + pd.offsets.MonthEnd(0)).normalize()
            days_in_month = int((month_end - month_start).days) + 1
            if days_in_month <= 0:
                return

            imp_pron_day = float(r.get("IMPORTE_PRON", 0.0)) / days_in_month
            imp_real_day = float(r.get("IMPORTE_REAL", 0.0)) / days_in_month

            d = month_start
            while d <= month_end:
                if d < start_date or d > end_date:
                    d += pd.Timedelta(days=1)
                    continue
                if not pd.isna(hasta) and d > hasta:
                    d += pd.Timedelta(days=1)
                    continue

                rr = r.copy()
                rr["IMPORTE_PRON"] = imp_pron_day
                rr["IMPORTE_REAL"] = imp_real_day
                rows.append((d.normalize(), rr))
                d += pd.Timedelta(days=1)

        # PUNTUAL
        if periodicidad in ("PUNTUAL", "ONE-OFF", "ONEOFF"):
            if pd.isna(r.get("FECHA_FIJA")):
                continue
            add_if_valid(pd.Timestamp(r["FECHA_FIJA"]).normalize())
            continue

        # ANUAL
        if periodicidad == "ANUAL":
            if pd.isna(r.get("FECHA_FIJA")):
                continue
            base = pd.Timestamp(r["FECHA_FIJA"]).normalize()
            year = start_date.year
            while True:
                candidate = pd.Timestamp(year=year, month=base.month, day=base.day)
                if candidate > end_date:
                    break
                add_if_valid(candidate.normalize())
                year += 1
            continue

        # SEMANAL
        if periodicidad == "SEMANAL":
            anchor = r.get("FECHA_FIJA")
            if pd.isna(anchor):
                continue
            anchor = pd.Timestamp(anchor).normalize()

            if not pd.isna(hasta):
                stop = min(hasta, end_date)
            else:
                month_start = anchor.replace(day=1)
                stop = (month_start + pd.offsets.MonthEnd(0)).normalize()
                stop = min(stop, end_date)

            d = anchor
            while d <= stop:
                add_if_valid(d.normalize())
                d = d + pd.Timedelta(days=7)
            continue

        # MENSUAL / BIMESTRAL / TRIMESTRAL / SEMESTRAL
        if periodicidad in ("MENSUAL", "BIMESTRAL", "BIMENSUAL", "TRIMESTRAL", "SEMESTRAL"):
            step = months_step_from_periodicidad(periodicidad)

            if not pd.isna(r.get("FECHA_FIJA")):
                base_date = pd.Timestamp(r["FECHA_FIJA"]).normalize()
                anchor_day = base_date.day
            else:
                day = r.get("DIA_MES")
                if not day or pd.isna(day):
                    continue
                y0, m0 = start_date.year, start_date.month
                last_day0 = (pd.Timestamp(year=y0, month=m0, day=1) + pd.offsets.MonthEnd(0)).day
                base_date = pd.Timestamp(year=y0, month=m0, day=min(int(day), int(last_day0))).normalize()
                anchor_day = base_date.day

            current = base_date

            while current <= end_date:
                y, m = current.year, current.month

                if regla in ("DIA_MES", "FECHA_FIJA"):
                    last_day = (pd.Timestamp(year=y, month=m, day=1) + pd.offsets.MonthEnd(0)).day
                    d_base = pd.Timestamp(year=y, month=m, day=min(int(anchor_day), int(last_day))).normalize()

                elif regla == "ULTIMO_HABIL":
                    d_base = (pd.Timestamp(year=y, month=m, day=1) + pd.offsets.MonthEnd(0)).normalize()
                    if d_base.weekday() >= 5:
                        d_base = next_business_day(d_base).normalize()
                else:
                    d_base = None

                if d_base is not None:
                    if prorrateo == "DIARIO":
                        if within_limits_base(d_base):
                            add_prorrateo_diario_for_month(d_base)
                    else:
                        add_if_valid(d_base)

                current = (current + pd.DateOffset(months=step)).normalize()

    if not rows:
        return pd.DataFrame(columns=[
            "FECHA", "CONCEPTO", "TIPO", "DEPARTAMENTO", "IMPORTE_PRON", "IMPORTE_REAL", "NATURALEZA",
            "PAGADO_BOOL", "FECHA_PAGO", "PRORRATEO"
        ])

    out = pd.DataFrame([{
        "FECHA": d,
        "CONCEPTO": rr["GENERAL"],
        "TIPO": rr["TIPO"],
        "DEPARTAMENTO": rr["DEPARTAMENTO"],
        "IMPORTE_PRON": float(rr.get("IMPORTE_PRON", 0.0)),
        "IMPORTE_REAL": float(rr.get("IMPORTE_REAL", 0.0)),
        "NATURALEZA": rr.get("NATURALEZA", ""),
        "PAGADO_BOOL": bool(rr.get("PAGADO_BOOL", False)),
        "FECHA_PAGO": rr.get("FECHA_PAGO", pd.NaT),
        "PRORRATEO": str(rr.get("PRORRATEO", "")).upper().strip(),
    } for d, rr in rows])

    out["FECHA_PAGO"] = pd.to_datetime(out["FECHA_PAGO"], errors="coerce").dt.normalize()
    return out.sort_values("FECHA").reset_index(drop=True)

def compute_balance_from_amount(df: pd.DataFrame, starting_balance: float, amount_col: str) -> pd.DataFrame:
    df = df.copy()
    df["COBROS"] = df.apply(lambda x: x[amount_col] if x["TIPO"] == "INGRESO" else 0.0, axis=1)
    df["PAGOS"] = df.apply(lambda x: x[amount_col] if x["TIPO"] == "GASTO" else 0.0, axis=1)
    df["NETO"] = df["COBROS"] - df["PAGOS"]
    df["SALDO"] = starting_balance + df["NETO"].cumsum()
    return df

# -----------------------------
# Sidebar inputs
# -----------------------------
st.sidebar.header("Inputs")
saldo_fecha = st.sidebar.date_input("Fecha del saldo (hoy)", value=date.today())
months_horizon = st.sidebar.slider("Horizonte forecast (meses)", min_value=1, max_value=36, value=12)
dedupe_exact = st.sidebar.checkbox("Eliminar duplicados exactos (red de seguridad)", value=True)
uploaded = st.sidebar.file_uploader("Sube el Excel de catálogo (xlsx)", type=["xlsx"])

# -----------------------------
# Main flow
# -----------------------------
if not uploaded:
    st.info("Sube tu Excel para generar el dashboard.")
    st.stop()

# Leer BANCOS -> saldo inicial + líneas fijas
try:
    bancos = read_bancos_from_excel(uploaded)
except Exception as e:
    st.error(f"Error leyendo hoja BANCOS: {e}")
    st.stop()

saldo_hoy = float(bancos["total_bancos"])
cuenta_suplidos = bancos.get("cuenta_suplidos", None)
cuenta_efectivo = bancos.get("cuenta_efectivo", None)

# Mostrar en sidebar como info
st.sidebar.header("Bancos (desde Excel)")
st.sidebar.metric("TOTAL BANCOS (saldo inicial)", eur(saldo_hoy))
if cuenta_suplidos is not None:
    st.sidebar.metric("CUENTA SUPLIDOS", eur(cuenta_suplidos))
if cuenta_efectivo is not None:
    st.sidebar.metric("CUENTA DE EFECTIVO", eur(cuenta_efectivo))

# Leer catálogo
try:
    catalog = read_catalog_from_excel(uploaded)
except Exception as e:
    st.error(f"Error leyendo catálogo: {e}")
    st.stop()

start_ts = pd.Timestamp(saldo_fecha).normalize()
generated = generate_events_from_catalog(catalog=catalog, start_date=start_ts, months_horizon=months_horizon)

if dedupe_exact and not generated.empty:
    generated = generated.drop_duplicates(
        subset=["FECHA", "CONCEPTO", "TIPO", "DEPARTAMENTO", "IMPORTE_PRON", "IMPORTE_REAL", "PAGADO_BOOL", "FECHA_PAGO", "PRORRATEO"],
        keep="first"
    )

if generated.empty:
    st.warning("No se generaron movimientos (revisa PERIODICIDAD / REGLA_FECHA / VALOR_FECHA / HASTA).")
    st.dataframe(catalog.head(50), use_container_width=True)
    st.stop()

# -----------------------------
# Filtros base
# -----------------------------
st.sidebar.header("Filtros base")
deptos = sorted(generated["DEPARTAMENTO"].dropna().unique().tolist())
tipos = sorted(generated["TIPO"].dropna().unique().tolist())

sel_deptos = st.sidebar.multiselect("Departamento", options=deptos, default=deptos)
sel_tipos = st.sidebar.multiselect("Tipo", options=tipos, default=tipos)

base_filtered = generated[
    generated["DEPARTAMENTO"].isin(sel_deptos) &
    generated["TIPO"].isin(sel_tipos)
].copy().sort_values("FECHA").reset_index(drop=True)

# -----------------------------
# PRON vs REAL
# -----------------------------
# PRON = solo pendientes (no pagados)
pron_df = base_filtered[~base_filtered["PAGADO_BOOL"]].copy().sort_values("FECHA").reset_index(drop=True)

# REAL = TODOS (pagados + pendientes) con IMPORTE_REAL
real_df = base_filtered.copy()
real_df["FECHA_EFECTIVA_REAL"] = real_df["FECHA_PAGO"].fillna(real_df["FECHA"])
real_df["FECHA"] = pd.to_datetime(real_df["FECHA_EFECTIVA_REAL"], errors="coerce").dt.normalize()
real_df = real_df.sort_values("FECHA").reset_index(drop=True)

# Consolidación
consolidado_pron = compute_balance_from_amount(pron_df, saldo_hoy, "IMPORTE_PRON")
consolidado_real = compute_balance_from_amount(real_df, saldo_hoy, "IMPORTE_REAL")

# Fila saldo inicial
base_row = pd.DataFrame([{
    "FECHA": start_ts,
    "CONCEPTO": "SALDO BANCOS TOTAL",
    "TIPO": "SALDO",
    "DEPARTAMENTO": "",
    "IMPORTE_PRON": 0.0,
    "IMPORTE_REAL": 0.0,
    "NATURALEZA": "SALDO",
    "PAGADO_BOOL": False,
    "FECHA_PAGO": pd.NaT,
    "PRORRATEO": "",
    "COBROS": 0.0,
    "PAGOS": 0.0,
    "NETO": 0.0,
    "SALDO": saldo_hoy
}])

consolidado_pron2 = pd.concat([base_row, consolidado_pron], ignore_index=True)
consolidado_real2 = pd.concat([base_row, consolidado_real], ignore_index=True)

# -----------------------------
# Buscador + rango fechas
# -----------------------------
st.sidebar.header("Búsqueda y rango (solo visualización)")
q = st.sidebar.text_input("Buscar concepto", value="").strip()

min_d = min(consolidado_pron2["FECHA"].min(), consolidado_real2["FECHA"].min()).date()
max_d = max(consolidado_pron2["FECHA"].max(), consolidado_real2["FECHA"].max()).date()

date_range = st.sidebar.date_input(
    "Rango de fechas",
    value=(min_d, max_d),
    min_value=min_d,
    max_value=max_d
)

if isinstance(date_range, tuple) and len(date_range) == 2:
    d_from, d_to = date_range
else:
    d_from, d_to = min_d, max_d

# -----------------------------
# KPIs
# -----------------------------
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.metric("Saldo inicial (TOTAL BANCOS)", eur(saldo_hoy))
with c2:
    st.metric("Neto periodo (PRON, pendientes)", eur(consolidado_pron["NETO"].sum() if not consolidado_pron.empty else 0.0))
with c3:
    st.metric("Saldo final (PRON, pendientes)", eur(consolidado_pron["SALDO"].iloc[-1] if not consolidado_pron.empty else saldo_hoy))
with c4:
    st.metric("Saldo final (REAL, estimado/ejecutado)", eur(consolidado_real["SALDO"].iloc[-1] if not consolidado_real.empty else saldo_hoy))

# -----------------------------
# Gráfico diario (PRON vs REAL) + líneas fijas
# -----------------------------
st.subheader("Evolución de saldo — diario (Pronosticado vs Real)")

d_pron = consolidado_pron2[["FECHA", "SALDO"]].copy()
d_pron["FECHA"] = pd.to_datetime(d_pron["FECHA"]).dt.normalize()
d_pron = d_pron.groupby("FECHA", as_index=False)["SALDO"].last().rename(columns={"SALDO": "SALDO_PRON"})

d_real = consolidado_real2[["FECHA", "SALDO"]].copy()
d_real["FECHA"] = pd.to_datetime(d_real["FECHA"]).dt.normalize()
d_real = d_real.groupby("FECHA", as_index=False)["SALDO"].last().rename(columns={"SALDO": "SALDO_REAL"})

daily = pd.merge(d_pron, d_real, on="FECHA", how="outer").sort_values("FECHA")

all_days = pd.date_range(start=daily["FECHA"].min(), end=daily["FECHA"].max(), freq="D")
daily = daily.set_index("FECHA").reindex(all_days).rename_axis("FECHA").reset_index()
daily["SALDO_PRON"] = daily["SALDO_PRON"].ffill().fillna(saldo_hoy)
daily["SALDO_REAL"] = daily["SALDO_REAL"].ffill().fillna(saldo_hoy)

if cuenta_suplidos is not None:
    daily["LINEA_SUPLIDOS"] = float(cuenta_suplidos)
if cuenta_efectivo is not None:
    daily["LINEA_EFECTIVO"] = float(cuenta_efectivo)

zoom_start = pd.Timestamp(d_from)
zoom_end = pd.Timestamp(d_to)
daily_zoom = daily[(daily["FECHA"] >= zoom_start) & (daily["FECHA"] <= zoom_end)].copy()

value_vars = ["SALDO_PRON", "SALDO_REAL"]
series_map = {
    "SALDO_PRON": "Pronosticado (pendiente)",
    "SALDO_REAL": "Real (estimado/ejecutado)"
}
if "LINEA_EFECTIVO" in daily_zoom.columns:
    value_vars.append("LINEA_EFECTIVO")
    series_map["LINEA_EFECTIVO"] = "Cuenta efectivo"
if "LINEA_SUPLIDOS" in daily_zoom.columns:
    value_vars.append("LINEA_SUPLIDOS")
    series_map["LINEA_SUPLIDOS"] = "Cuenta suplidos"

plot_df = daily_zoom.melt(
    id_vars=["FECHA"],
    value_vars=value_vars,
    var_name="SERIE",
    value_name="SALDO"
)
plot_df["SERIE"] = plot_df["SERIE"].map(series_map)

domain = ["Cuenta efectivo", "Cuenta suplidos", "Pronosticado (pendiente)", "Real (estimado/ejecutado)"]
range_ = ["#FF8C00", "#D62728", "#7DB9FF", "#0B2E8A"]  # naranja, rojo, azul claro, azul oscuro

chart = (
    alt.Chart(plot_df)
    .mark_line()
    .encode(
        x=alt.X("FECHA:T", title="Fecha"),
        y=alt.Y("SALDO:Q", title="Saldo"),
        color=alt.Color("SERIE:N", title="", scale=alt.Scale(domain=domain, range=range_)),
        tooltip=[
            alt.Tooltip("FECHA:T", title="Fecha"),
            alt.Tooltip("SERIE:N", title="Serie"),
            alt.Tooltip("SALDO:Q", title="Saldo", format=",.2f"),
        ],
    )
    .properties(height=340)
)
st.altair_chart(chart, use_container_width=True)

# -----------------------------
# Movimientos PRON (pendientes)
# -----------------------------
st.subheader("Movimientos (formato tesorería) — PRON (pendientes)")

mov_pron = consolidado_pron2.copy()
mov_pron = mov_pron[(mov_pron["FECHA"].dt.date >= d_from) & (mov_pron["FECHA"].dt.date <= d_to)].copy()
if q:
    mov_pron = mov_pron[mov_pron["CONCEPTO"].astype(str).str.contains(q, case=False, na=False)].copy()

mov_pron["VTO. PAGO"] = mov_pron["FECHA"].dt.strftime("%d-%m-%y")
mov_pron["COBRADO/PAGADO"] = mov_pron.apply(
    lambda r: estado_cobro_pago(r.get("TIPO", ""), bool(r.get("PAGADO_BOOL", False))),
    axis=1
)

mov_pron_out = mov_pron[["VTO. PAGO", "CONCEPTO", "COBRADO/PAGADO", "COBROS", "PAGOS", "SALDO"]].copy()

styled_pron = (
    mov_pron_out.style
    .applymap(style_estado_cell, subset=["COBRADO/PAGADO"])
    .applymap(color_saldo, subset=["SALDO"])
    .format({"COBROS": eur, "PAGOS": eur, "SALDO": eur})
)
st.dataframe(styled_pron, use_container_width=True)

# -----------------------------
# Movimientos REAL (incluye pendientes + pagados)
# -----------------------------
st.subheader("Movimientos (formato tesorería) — REAL (estimado/ejecutado)")

mov_real = consolidado_real2.copy()
mov_real = mov_real[(mov_real["FECHA"].dt.date >= d_from) & (mov_real["FECHA"].dt.date <= d_to)].copy()
if q:
    mov_real = mov_real[mov_real["CONCEPTO"].astype(str).str.contains(q, case=False, na=False)].copy()

mov_real["VTO. PAGO"] = mov_real["FECHA"].dt.strftime("%d-%m-%y")
mov_real["COBRADO/PAGADO"] = mov_real.apply(
    lambda r: estado_cobro_pago(r.get("TIPO", ""), bool(r.get("PAGADO_BOOL", False))),
    axis=1
)

mov_real_out = mov_real[["VTO. PAGO", "CONCEPTO", "COBRADO/PAGADO", "COBROS", "PAGOS", "SALDO"]].copy()

styled_real = (
    mov_real_out.style
    .applymap(style_estado_cell, subset=["COBRADO/PAGADO"])
    .applymap(color_saldo, subset=["SALDO"])
    .format({"COBROS": eur, "PAGOS": eur, "SALDO": eur})
)
st.dataframe(styled_real, use_container_width=True)

# -----------------------------
# Resumen mensual (según lo visible) - PRON
# -----------------------------
st.subheader("Resumen mensual (según lo visible) — PRON")

tmp = mov_pron.copy()
tmp["MES"] = tmp["FECHA"].dt.to_period("M").astype(str)

monthly_visible = tmp.groupby("MES", as_index=False).agg(
    COBROS=("COBROS", "sum"),
    PAGOS=("PAGOS", "sum"),
    NETO=("NETO", "sum"),
)

glob = consolidado_pron2.copy()
glob["MES"] = glob["FECHA"].dt.to_period("M").astype(str)
monthly_global_close = glob.groupby("MES", as_index=False).agg(
    SALDO_CIERRE=("SALDO", "last")
)

monthly = monthly_visible.merge(monthly_global_close, on="MES", how="left")
styled_month = monthly.style.format({c: eur for c in monthly.columns if c != "MES"})
st.dataframe(styled_month, use_container_width=True)

# -----------------------------
# Exportar a Excel (lo visible)
# -----------------------------
st.subheader("Exportar (lo visible)")

export_pron = mov_pron_out.copy()
export_real = mov_real_out.copy()
export_month = monthly.copy()

def to_numeric_safe(df0: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    df1 = df0.copy()
    for c in cols:
        if c in df1.columns:
            df1[c] = pd.to_numeric(df1[c], errors="coerce")
    return df1

export_pron = to_numeric_safe(export_pron, ["COBROS", "PAGOS", "SALDO"])
export_real = to_numeric_safe(export_real, ["COBROS", "PAGOS", "SALDO"])
export_month = to_numeric_safe(export_month, ["COBROS", "PAGOS", "NETO", "SALDO_CIERRE"])

def build_excel_bytes(df_pron: pd.DataFrame, df_real: pd.DataFrame, df_month: pd.DataFrame) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df_pron.to_excel(writer, index=False, sheet_name="Mov_PRON")
        df_real.to_excel(writer, index=False, sheet_name="Mov_REAL")
        df_month.to_excel(writer, index=False, sheet_name="Resumen_PRON")
    bio.seek(0)
    return bio.read()

xlsx_bytes = build_excel_bytes(export_pron, export_real, export_month)

st.download_button(
    "Descargar Excel (XLSX)",
    data=xlsx_bytes,
    file_name="tesoreria_export.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

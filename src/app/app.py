import re
from datetime import date
from io import BytesIO

import pandas as pd
import streamlit as st
import altair as alt

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

def read_catalog_from_excel(uploaded_file) -> pd.DataFrame:
    raw = pd.read_excel(uploaded_file, sheet_name=0, header=None, engine="openpyxl")
    header_idx = find_header_row(raw)
    if header_idx is None:
        raise ValueError("No encuentro la fila de cabecera (debe contener 'GENERAL' y 'TIPO').")

    df = pd.read_excel(uploaded_file, sheet_name=0, header=header_idx, engine="openpyxl")
    df = normalize_cols(df)

    # Requeridas mínimas
    missing = [c for c in MIN_REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas requeridas: {missing}. Columnas detectadas: {list(df.columns)}")

    # Importe pronosticado / real
    # Permitimos varias formas de nombre
    if "IMPORTE PRONOSTICADO" not in df.columns and "IMPORTE_PRONOSTICADO" in df.columns:
        df["IMPORTE PRONOSTICADO"] = df["IMPORTE_PRONOSTICADO"]
    if "IMPORTE REAL" not in df.columns and "IMPORTE_REAL" in df.columns:
        df["IMPORTE REAL"] = df["IMPORTE_REAL"]

    if "IMPORTE PRONOSTICADO" not in df.columns:
        raise ValueError("Falta columna: 'IMPORTE PRONOSTICADO'")
    if "IMPORTE REAL" not in df.columns:
        # si no existe, la creamos en 0 para no romper nada
        df["IMPORTE REAL"] = 0.0

    # HASTA opcional
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

    # Importes numéricos
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
    end_date = (start_date + pd.offsets.MonthBegin(months_horizon + 1)).normalize()
    rows = []

    for _, r in catalog.iterrows():
        periodicidad = str(r.get("PERIODICIDAD", "")).upper().strip()
        regla = str(r.get("REGLA_FECHA", "")).upper().strip()
        ajuste = str(r.get("AJUSTE FINDE", "")).upper().strip()
        lag = int(r.get("LAG", 0))

        hasta = r.get("HASTA", pd.NaT)
        hasta = pd.Timestamp(hasta).normalize() if not pd.isna(hasta) else pd.NaT

        def apply_adjustments(d: pd.Timestamp) -> pd.Timestamp:
            if ajuste == "SIG_HABIL":
                d = next_business_day(d)
            if lag:
                d = add_business_days(d, lag)
            return d

        def within_limits(d: pd.Timestamp) -> bool:
            if d < start_date or d > end_date:
                return False
            if not pd.isna(hasta) and d > hasta:
                return False
            return True

        # PUNTUAL
        if periodicidad in ("PUNTUAL", "ONE-OFF", "ONEOFF"):
            if pd.isna(r.get("FECHA_FIJA")):
                continue
            d = pd.Timestamp(r["FECHA_FIJA"]).normalize()
            d = apply_adjustments(d)
            if within_limits(d):
                rows.append((d, r))
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
                d = apply_adjustments(candidate)
                if within_limits(d):
                    rows.append((d, r))
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
                dd = apply_adjustments(d)
                if within_limits(dd):
                    rows.append((dd, r))
                d = d + pd.Timedelta(days=7)
            continue

        # POR MESES
        if periodicidad in ("MENSUAL", "BIMESTRAL", "BIMENSUAL", "TRIMESTRAL", "SEMESTRAL"):
            step = months_step_from_periodicidad(periodicidad)

            if not pd.isna(r.get("FECHA_FIJA")):
                base_date = pd.Timestamp(r["FECHA_FIJA"]).normalize()
            else:
                day = r.get("DIA_MES")
                if not day or pd.isna(day):
                    continue
                y, m = start_date.year, start_date.month
                last_day = (pd.Timestamp(year=y, month=m, day=1) + pd.offsets.MonthEnd(0)).day
                base_date = pd.Timestamp(year=y, month=m, day=min(int(day), int(last_day))).normalize()

            anchor_day = base_date.day
            current = base_date

            while current <= end_date:
                y, m = current.year, current.month

                if regla == "DIA_MES":
                    last_day = (pd.Timestamp(year=y, month=m, day=1) + pd.offsets.MonthEnd(0)).day
                    d = pd.Timestamp(year=y, month=m, day=min(int(anchor_day), int(last_day))).normalize()
                elif regla == "ULTIMO_HABIL":
                    d = pd.Timestamp(year=y, month=m, day=1) + pd.offsets.MonthEnd(0)
                    d = next_business_day(d) if d.weekday() >= 5 else d
                    d = d.normalize()
                elif regla == "FECHA_FIJA":
                    last_day = (pd.Timestamp(year=y, month=m, day=1) + pd.offsets.MonthEnd(0)).day
                    d = pd.Timestamp(year=y, month=m, day=min(int(anchor_day), int(last_day))).normalize()
                else:
                    d = None

                if d is not None:
                    d = apply_adjustments(d)
                    if within_limits(d):
                        rows.append((d, r))

                current = (current + pd.DateOffset(months=step)).normalize()

    if not rows:
        return pd.DataFrame(columns=["FECHA", "CONCEPTO", "TIPO", "DEPARTAMENTO", "IMPORTE_PRON", "IMPORTE_REAL", "NATURALEZA"])

    out = pd.DataFrame([{
        "FECHA": d,
        "CONCEPTO": rr["GENERAL"],
        "TIPO": rr["TIPO"],
        "DEPARTAMENTO": rr["DEPARTAMENTO"],
        "IMPORTE_PRON": float(rr["IMPORTE_PRON"]),
        "IMPORTE_REAL": float(rr["IMPORTE_REAL"]),
        "NATURALEZA": rr.get("NATURALEZA", "")
    } for d, rr in rows])

    return out.sort_values("FECHA").reset_index(drop=True)

def compute_balance_dual(df: pd.DataFrame, starting_balance: float) -> pd.DataFrame:
    df = df.copy()

    # Pronosticado
    df["COBROS_PRON"] = df.apply(lambda x: x["IMPORTE_PRON"] if x["TIPO"] == "INGRESO" else 0.0, axis=1)
    df["PAGOS_PRON"] = df.apply(lambda x: x["IMPORTE_PRON"] if x["TIPO"] == "GASTO" else 0.0, axis=1)
    df["NETO_PRON"] = df["COBROS_PRON"] - df["PAGOS_PRON"]
    df["SALDO_PRON"] = starting_balance + df["NETO_PRON"].cumsum()

    # Real
    df["COBROS_REAL"] = df.apply(lambda x: x["IMPORTE_REAL"] if x["TIPO"] == "INGRESO" else 0.0, axis=1)
    df["PAGOS_REAL"] = df.apply(lambda x: x["IMPORTE_REAL"] if x["TIPO"] == "GASTO" else 0.0, axis=1)
    df["NETO_REAL"] = df["COBROS_REAL"] - df["PAGOS_REAL"]
    df["SALDO_REAL"] = starting_balance + df["NETO_REAL"].cumsum()

    return df

# -----------------------------
# Formatting helpers
# -----------------------------
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

# -----------------------------
# Sidebar inputs
# -----------------------------
st.sidebar.header("Inputs")
saldo_fecha = st.sidebar.date_input("Fecha del saldo (hoy)", value=date.today())
saldo_hoy = st.sidebar.number_input("Saldo actual en banco (€)", min_value=-1e12, max_value=1e12, value=0.0, step=100.0)
months_horizon = st.sidebar.slider("Horizonte forecast (meses)", min_value=1, max_value=36, value=12)

dedupe_exact = st.sidebar.checkbox("Eliminar duplicados exactos (red de seguridad)", value=True)
uploaded = st.sidebar.file_uploader("Sube el Excel de catálogo (xlsx)", type=["xlsx"])

# -----------------------------
# Main flow
# -----------------------------
if not uploaded:
    st.info("Sube tu Excel para generar el dashboard.")
    st.stop()

try:
    catalog = read_catalog_from_excel(uploaded)
except Exception as e:
    st.error(f"Error leyendo catálogo: {e}")
    st.stop()

start_ts = pd.Timestamp(saldo_fecha).normalize()
generated = generate_events_from_catalog(catalog=catalog, start_date=start_ts, months_horizon=months_horizon)

if dedupe_exact and not generated.empty:
    generated = generated.drop_duplicates(subset=["FECHA", "CONCEPTO", "TIPO", "DEPARTAMENTO", "IMPORTE_PRON", "IMPORTE_REAL"], keep="first")

if generated.empty:
    st.warning("No se generaron movimientos (revisa PERIODICIDAD / REGLA_FECHA / VALOR_FECHA / HASTA).")
    st.dataframe(catalog.head(50), use_container_width=True)
    st.stop()

# -----------------------------
# Filtros base (afectan al cálculo real)
# -----------------------------
st.sidebar.header("Filtros base (afectan al saldo real)")
deptos = sorted(generated["DEPARTAMENTO"].dropna().unique().tolist())
tipos = sorted(generated["TIPO"].dropna().unique().tolist())

sel_deptos = st.sidebar.multiselect("Departamento", options=deptos, default=deptos)
sel_tipos = st.sidebar.multiselect("Tipo", options=tipos, default=tipos)

base_filtered = generated[
    generated["DEPARTAMENTO"].isin(sel_deptos) &
    generated["TIPO"].isin(sel_tipos)
].copy()

base_filtered = base_filtered.sort_values("FECHA").reset_index(drop=True)

# Saldo dual (pron y real)
consolidado = compute_balance_dual(base_filtered, float(saldo_hoy))

# Fila inicial saldo
base_row = pd.DataFrame([{
    "FECHA": start_ts,
    "CONCEPTO": "SALDO BANCOS TOTAL",
    "TIPO": "SALDO",
    "DEPARTAMENTO": "",
    "IMPORTE_PRON": 0.0,
    "IMPORTE_REAL": 0.0,
    "NATURALEZA": "SALDO",
    "COBROS_PRON": 0.0, "PAGOS_PRON": 0.0, "NETO_PRON": 0.0, "SALDO_PRON": float(saldo_hoy),
    "COBROS_REAL": 0.0, "PAGOS_REAL": 0.0, "NETO_REAL": 0.0, "SALDO_REAL": float(saldo_hoy),
}])

consolidado2 = pd.concat([base_row, consolidado], ignore_index=True)
consolidado2["_ORD"] = 1
consolidado2.loc[0, "_ORD"] = 0
consolidado2 = consolidado2.sort_values(["FECHA", "_ORD"]).drop(columns=["_ORD"]).reset_index(drop=True)

# -----------------------------
# Buscador + rango fechas (solo visualización)
# -----------------------------
st.sidebar.header("Búsqueda y rango (solo visualización)")
q = st.sidebar.text_input("Buscar concepto", value="").strip()

min_d = consolidado2["FECHA"].min().date()
max_d = consolidado2["FECHA"].max().date()

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

view_df = consolidado2.copy()
view_df = view_df[(view_df["FECHA"].dt.date >= d_from) & (view_df["FECHA"].dt.date <= d_to)].copy()

if q:
    view_df = view_df[view_df["CONCEPTO"].astype(str).str.contains(q, case=False, na=False)].copy()

view_df = view_df.sort_values("FECHA").reset_index(drop=True)

# -----------------------------
# KPIs (pronosticado vs real)
# -----------------------------
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.metric("Saldo inicial (hoy)", eur(saldo_hoy))
with c2:
    st.metric("Neto periodo (PRON)", eur(consolidado["NETO_PRON"].sum()))
with c3:
    st.metric("Saldo final (PRON)", eur(consolidado["SALDO_PRON"].iloc[-1]))
with c4:
    st.metric("Saldo final (REAL)", eur(consolidado["SALDO_REAL"].iloc[-1]))

# -----------------------------
# Gráfico diario doble (PRON vs REAL) + zoom al rango visible
# -----------------------------
st.subheader("Evolución de saldo — diario (Pronosticado vs Real)")

# Serie diaria con 2 saldos
daily = consolidado2[["FECHA", "SALDO_PRON", "SALDO_REAL"]].copy()
daily["FECHA"] = pd.to_datetime(daily["FECHA"]).dt.normalize()
daily = daily.groupby("FECHA", as_index=False).agg(
    SALDO_PRON=("SALDO_PRON", "last"),
    SALDO_REAL=("SALDO_REAL", "last"),
)

all_days = pd.date_range(start=daily["FECHA"].min(), end=daily["FECHA"].max(), freq="D")
daily = daily.set_index("FECHA").reindex(all_days).rename_axis("FECHA").reset_index()

daily["SALDO_PRON"] = daily["SALDO_PRON"].ffill().fillna(float(saldo_hoy))
daily["SALDO_REAL"] = daily["SALDO_REAL"].ffill().fillna(float(saldo_hoy))

# Zoom
zoom_start = pd.Timestamp(d_from) - pd.Timedelta(days=1)
zoom_end = pd.Timestamp(d_to)
daily_zoom = daily[(daily["FECHA"] >= zoom_start) & (daily["FECHA"] <= zoom_end)].copy()
daily_zoom = daily_zoom[daily_zoom["FECHA"] >= pd.Timestamp(d_from)].copy()

# Long format para Altair
plot_df = daily_zoom.melt(
    id_vars=["FECHA"],
    value_vars=["SALDO_PRON", "SALDO_REAL"],
    var_name="SERIE",
    value_name="SALDO"
)

plot_df["SERIE"] = plot_df["SERIE"].map({
    "SALDO_PRON": "Pronosticado",
    "SALDO_REAL": "Real"
})

chart = (
    alt.Chart(plot_df)
    .mark_line()
    .encode(
        x=alt.X("FECHA:T", title="Fecha"),
        y=alt.Y("SALDO:Q", title="Saldo"),
        color=alt.Color("SERIE:N", title=""),
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
# Movimientos formato tesorería (mostramos PRON; y añadimos REAL aparte)
# -----------------------------
st.subheader("Movimientos (formato tesorería)")

mov = view_df.copy()
mov["VTO. PAGO"] = mov["FECHA"].dt.strftime("%d-%m-%y")

mov_out = mov[[
    "VTO. PAGO", "CONCEPTO",
    "COBROS_PRON", "PAGOS_PRON", "SALDO_PRON",
    "COBROS_REAL", "PAGOS_REAL", "SALDO_REAL"
]].copy()

mov_out = mov_out.rename(columns={
    "COBROS_PRON": "COBROS (PRON)",
    "PAGOS_PRON": "PAGOS (PRON)",
    "SALDO_PRON": "SALDO (PRON)",
    "COBROS_REAL": "COBROS (REAL)",
    "PAGOS_REAL": "PAGOS (REAL)",
    "SALDO_REAL": "SALDO (REAL)",
})

styled_mov = (
    mov_out.style
    .applymap(color_saldo, subset=["SALDO (PRON)", "SALDO (REAL)"])
    .format({
        "COBROS (PRON)": eur, "PAGOS (PRON)": eur, "SALDO (PRON)": eur,
        "COBROS (REAL)": eur, "PAGOS (REAL)": eur, "SALDO (REAL)": eur,
    })
)

st.dataframe(styled_mov, use_container_width=True)

# -----------------------------
# Resumen mensual (visible) PRON y REAL
# -----------------------------
st.subheader("Resumen mensual (según lo visible)")

tmp = view_df.copy()
tmp["MES"] = tmp["FECHA"].dt.to_period("M").astype(str)

monthly = tmp.groupby("MES", as_index=False).agg(
    COBROS_PRON=("COBROS_PRON", "sum"),
    PAGOS_PRON=("PAGOS_PRON", "sum"),
    NETO_PRON=("NETO_PRON", "sum"),
    SALDO_CIERRE_PRON=("SALDO_PRON", "last"),
    COBROS_REAL=("COBROS_REAL", "sum"),
    PAGOS_REAL=("PAGOS_REAL", "sum"),
    NETO_REAL=("NETO_REAL", "sum"),
    SALDO_CIERRE_REAL=("SALDO_REAL", "last"),
)

monthly = monthly.rename(columns={
    "COBROS_PRON": "COBROS (PRON)",
    "PAGOS_PRON": "PAGOS (PRON)",
    "NETO_PRON": "NETO (PRON)",
    "SALDO_CIERRE_PRON": "SALDO_CIERRE (PRON)",
    "COBROS_REAL": "COBROS (REAL)",
    "PAGOS_REAL": "PAGOS (REAL)",
    "NETO_REAL": "NETO (REAL)",
    "SALDO_CIERRE_REAL": "SALDO_CIERRE (REAL)",
})

styled_month = monthly.style.format({c: eur for c in monthly.columns if c != "MES"})
st.dataframe(styled_month, use_container_width=True)

# -----------------------------
# Exportar a Excel (visible)
# -----------------------------
st.subheader("Exportar (lo visible)")

export_mov = mov_out.copy()
for c in export_mov.columns:
    if c not in ["VTO. PAGO", "CONCEPTO"]:
        export_mov[c] = pd.to_numeric(export_mov[c], errors="coerce")

export_month = monthly.copy()
for c in export_month.columns:
    if c != "MES":
        export_month[c] = pd.to_numeric(export_month[c], errors="coerce")

def build_excel_bytes(df_mov: pd.DataFrame, df_month: pd.DataFrame) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df_mov.to_excel(writer, index=False, sheet_name="Movimientos")
        df_month.to_excel(writer, index=False, sheet_name="Resumen_mensual")
    bio.seek(0)
    return bio.read()

xlsx_bytes = build_excel_bytes(export_mov, export_month)

st.download_button(
    "Descargar Excel (XLSX)",
    data=xlsx_bytes,
    file_name="tesoreria_export.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

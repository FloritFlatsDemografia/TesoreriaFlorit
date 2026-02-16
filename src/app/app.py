import re
from datetime import date
import pandas as pd
import streamlit as st

# -----------------------------
# Config
# -----------------------------
st.set_page_config(page_title="APP Tesorería", layout="wide")
st.title("APP Tesorería — Dashboard")

# -----------------------------
# Helpers
# -----------------------------
REQUIRED_COLS = [
    "GENERAL", "TIPO", "DEPARTAMENTO", "IMPORTE",
    "NATURALEZA", "PERIODICIDAD",
    "REGLA_FECHA", "VALOR_FECHA",
    "LAG", "AJUSTE FINDE", "PERIODO_SERVICIO",
    "IVA_APLICA", "IMPUESTO_TIPO", "MODELO"
]

def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df

def find_header_row(df_raw: pd.DataFrame) -> int | None:
    """
    Encuentra la fila donde están los headers (busca 'GENERAL' y 'TIPO').
    """
    for i in range(min(30, len(df_raw))):
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

    # Recorta a columnas existentes (al menos las mínimas)
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    # Permitimos que falten algunas de IVA/impuesto si aún no las usas
    allowed_missing = {"IVA_APLICA", "IMPUESTO_TIPO", "MODELO"}
    hard_missing = [c for c in missing if c not in allowed_missing]
    if hard_missing:
        raise ValueError(f"Faltan columnas requeridas: {hard_missing}. Columnas detectadas: {list(df.columns)}")

    # Limpieza básica
    df = df.dropna(how="all").copy()
    df["GENERAL"] = df["GENERAL"].astype(str).str.strip()
    df["TIPO"] = df["TIPO"].astype(str).str.strip().str.upper()
    df["DEPARTAMENTO"] = df["DEPARTAMENTO"].astype(str).str.strip().str.upper()
    df["NATURALEZA"] = df["NATURALEZA"].astype(str).str.strip().str.upper()
    df["PERIODICIDAD"] = df["PERIODICIDAD"].astype(str).str.strip().str.upper()
    df["REGLA_FECHA"] = df["REGLA_FECHA"].astype(str).str.strip().str.upper()

    # IMPORTE numérico
    df["IMPORTE"] = pd.to_numeric(df["IMPORTE"], errors="coerce").fillna(0.0)

    # LAG numérico
    df["LAG"] = pd.to_numeric(df["LAG"], errors="coerce").fillna(0).astype(int)

    # VALOR_FECHA: puede venir como fecha o como número (día)
    # Guardamos también DIA_MES derivado si procede
    def to_day_of_month(v):
        if pd.isna(v):
            return None
        if isinstance(v, (pd.Timestamp,)):
            return int(v.day)
        # Excel a veces trae datetime64
        if hasattr(v, "day") and hasattr(v, "month"):
            try:
                return int(v.day)
            except Exception:
                pass
        # número / string
        s = str(v).strip()
        if re.fullmatch(r"\d{1,2}", s):
            d = int(s)
            if 1 <= d <= 31:
                return d
        return None

    df["DIA_MES"] = df["VALOR_FECHA"].apply(to_day_of_month)

    # Para FECHA_FIJA, intentamos parsear fecha completa
    df["FECHA_FIJA"] = pd.to_datetime(df["VALOR_FECHA"], errors="coerce")

    return df

def next_business_day(d: pd.Timestamp) -> pd.Timestamp:
    # simplificado: sábado/domingo -> lunes
    if d.weekday() == 5:
        return d + pd.Timedelta(days=2)
    if d.weekday() == 6:
        return d + pd.Timedelta(days=1)
    return d

def add_business_days(d: pd.Timestamp, n: int) -> pd.Timestamp:
    # suma n días hábiles (sin festivos, solo finde)
    cur = d
    step = 1 if n >= 0 else -1
    remaining = abs(n)
    while remaining > 0:
        cur = cur + pd.Timedelta(days=step)
        if cur.weekday() < 5:
            remaining -= 1
    return cur

def generate_events_from_catalog(
    catalog: pd.DataFrame,
    start_date: pd.Timestamp,
    months_horizon: int
) -> pd.DataFrame:
    """
    Genera movimientos a futuro desde catálogo (MENSUAL/ANUAL/PUNTUAL).
    REGLA_FECHA:
      - DIA_MES: usa DIA_MES
      - FECHA_FIJA: usa FECHA_FIJA
      - ULTIMO_HABIL: último hábil del mes
    Ajuste finde: SIG_HABIL (si cae finde, siguiente hábil)
    Lag: suma días hábiles
    """
    end_date = (start_date + pd.offsets.MonthBegin(months_horizon + 1))  # algo holgado
    rows = []

    for _, r in catalog.iterrows():
        periodicidad = str(r.get("PERIODICIDAD", "")).upper()
        regla = str(r.get("REGLA_FECHA", "")).upper()
        ajuste = str(r.get("AJUSTE FINDE", "")).upper()
        lag = int(r.get("LAG", 0))
        tipo = str(r.get("TIPO", "")).upper()

        def apply_adjustments(d: pd.Timestamp) -> pd.Timestamp:
            if ajuste == "SIG_HABIL":
                d = next_business_day(d)
            if lag:
                d = add_business_days(d, lag)
            return d

        if periodicidad in ("PUNTUAL", "ONE-OFF", "ONEOFF"):
            if regla != "FECHA_FIJA" or pd.isna(r.get("FECHA_FIJA")):
                continue
            d = pd.Timestamp(r["FECHA_FIJA"]).normalize()
            d = apply_adjustments(d)
            if start_date <= d <= end_date:
                rows.append((d, r))

        elif periodicidad == "ANUAL":
            if regla == "FECHA_FIJA" and not pd.isna(r.get("FECHA_FIJA")):
                base = pd.Timestamp(r["FECHA_FIJA"]).normalize()
                # generar años dentro del rango
                year = start_date.year
                while pd.Timestamp(year=year, month=base.month, day=base.day) <= end_date:
                    d = pd.Timestamp(year=year, month=base.month, day=base.day)
                    d = apply_adjustments(d)
                    if start_date <= d <= end_date:
                        rows.append((d, r))
                    year += 1
            else:
                # si anual pero sin fecha fija, no generamos
                continue

        else:
            # Tratamos todo lo demás como MENSUAL por defecto
            # Genera meses dentro del horizonte
            for m in range(months_horizon + 1):
                month_start = (start_date + pd.offsets.MonthBegin(m)).normalize()
                year = month_start.year
                month = month_start.month

                if regla == "DIA_MES":
                    day = r.get("DIA_MES")
                    if not day or pd.isna(day):
                        continue
                    # cap day al último día del mes
                    last_day = (pd.Timestamp(year=year, month=month, day=1) + pd.offsets.MonthEnd(0)).day
                    day = int(min(int(day), int(last_day)))
                    d = pd.Timestamp(year=year, month=month, day=day)

                elif regla == "ULTIMO_HABIL":
                    d = pd.Timestamp(year=year, month=month, day=1) + pd.offsets.MonthEnd(0)
                    d = next_business_day(d) if d.weekday() >= 5 else d

                elif regla == "FECHA_FIJA" and not pd.isna(r.get("FECHA_FIJA")):
                    # Si alguien metió fecha fija pero periodicidad mensual, usamos día+mes del año actual
                    base = pd.Timestamp(r["FECHA_FIJA"])
                    last_day = (pd.Timestamp(year=year, month=month, day=1) + pd.offsets.MonthEnd(0)).day
                    d = pd.Timestamp(year=year, month=month, day=min(base.day, last_day))
                else:
                    continue

                d = apply_adjustments(d)
                if start_date <= d <= end_date:
                    rows.append((d, r))

    if not rows:
        return pd.DataFrame(columns=["FECHA", "CONCEPTO", "TIPO", "DEPARTAMENTO", "IMPORTE", "NATURALEZA"])

    out = pd.DataFrame([{
        "FECHA": d,
        "CONCEPTO": rr["GENERAL"],
        "TIPO": rr["TIPO"],
        "DEPARTAMENTO": rr["DEPARTAMENTO"],
        "IMPORTE": float(rr["IMPORTE"]),
        "NATURALEZA": rr.get("NATURALEZA", "")
    } for d, rr in rows])

    out = out.sort_values("FECHA").reset_index(drop=True)
    return out

def compute_balance(df: pd.DataFrame, starting_balance: float) -> pd.DataFrame:
    df = df.copy()
    df["COBROS"] = df.apply(lambda x: x["IMPORTE"] if x["TIPO"] == "INGRESO" else 0.0, axis=1)
    df["PAGOS"] = df.apply(lambda x: x["IMPORTE"] if x["TIPO"] == "GASTO" else 0.0, axis=1)
    df["NETO"] = df["COBROS"] - df["PAGOS"]
    df["SALDO"] = starting_balance + df["NETO"].cumsum()
    return df

# -----------------------------
# Sidebar inputs
# -----------------------------
st.sidebar.header("Inputs")
saldo_fecha = st.sidebar.date_input("Fecha del saldo (hoy)", value=date.today())
saldo_hoy = st.sidebar.number_input("Saldo actual en banco (€)", min_value=-1e12, max_value=1e12, value=0.0, step=100.0)
months_horizon = st.sidebar.slider("Horizonte forecast (meses)", min_value=1, max_value=36, value=12)

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

generated = generate_events_from_catalog(
    catalog=catalog,
    start_date=start_ts,
    months_horizon=months_horizon
)

if generated.empty:
    st.warning("No se generaron movimientos (revisa PERIODICIDAD / REGLA_FECHA / VALOR_FECHA).")
    st.dataframe(catalog.head(50), use_container_width=True)
    st.stop()

# Filtros
st.sidebar.header("Filtros")
deptos = sorted(generated["DEPARTAMENTO"].dropna().unique().tolist())
tipos = sorted(generated["TIPO"].dropna().unique().tolist())

sel_deptos = st.sidebar.multiselect("Departamento", options=deptos, default=deptos)
sel_tipos = st.sidebar.multiselect("Tipo", options=tipos, default=tipos)

filtered = generated[
    generated["DEPARTAMENTO"].isin(sel_deptos) &
    generated["TIPO"].isin(sel_tipos)
].copy()

filtered = filtered.sort_values("FECHA").reset_index(drop=True)
consolidado = compute_balance(filtered, float(saldo_hoy))

# -----------------------------
# Dashboard
# -----------------------------
c1, c2, c3 = st.columns(3)
with c1:
    st.metric("Saldo inicial (hoy)", f"{saldo_hoy:,.2f} €")
with c2:
    st.metric("Neto periodo", f"{consolidado['NETO'].sum():,.2f} €")
with c3:
    st.metric("Saldo final forecast", f"{consolidado['SALDO'].iloc[-1]:,.2f} €")

st.subheader("Evolución de saldo")
saldo_series = consolidado[["FECHA", "SALDO"]].set_index("FECHA")
st.line_chart(saldo_series)

st.subheader("Resumen mensual")
consolidado["MES"] = consolidado["FECHA"].dt.to_period("M").astype(str)
monthly = consolidado.groupby("MES", as_index=False).agg(
    COBROS=("COBROS", "sum"),
    PAGOS=("PAGOS", "sum"),
    NETO=("NETO", "sum"),
)
st.dataframe(monthly, use_container_width=True)

st.subheader("Movimientos consolidados")
st.dataframe(consolidado, use_container_width=True)

# Export
st.subheader("Exportar")
export_df = consolidado.copy()
export_df["FECHA"] = export_df["FECHA"].dt.date
csv_bytes = export_df.to_csv(index=False).encode("utf-8")
st.download_button(
    "Descargar consolidado (CSV)",
    data=csv_bytes,
    file_name="movimientos_consolidado.csv",
    mime="text/csv"
)

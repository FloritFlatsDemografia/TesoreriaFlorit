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

# Columnas mínimas realmente necesarias para generar eventos
MIN_REQUIRED = [
    "GENERAL", "TIPO", "DEPARTAMENTO", "IMPORTE",
    "NATURALEZA", "PERIODICIDAD",
    "REGLA_FECHA", "VALOR_FECHA",
    "LAG", "AJUSTE FINDE"
]

# Alias para compatibilidad entre nombres (tu Excel puede variar)
COL_ALIASES = {
    "IVA_APLICA": ["IVA_APLICA", "IVA_EN_FACTURA"],
    "IMPUESTO_TIPO": ["IMPUESTO_TIPO", "IVA_%", "IVA_PORCENTAJE", "IVA"],
    "MODELO": ["MODELO", "IVA_SENTIDO"],
    "TRATAMIENTO_IVA": ["TRATAMIENTO_IVA", "TRATAMINETO_IVA", "TRATAMIENTO IVA"],
    "PERIODO_SERVICIO": ["PERIODO_SERVICIO", "PERIODO_SERV.", "PERIODO"]
}

def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df

def coalesce_column(df: pd.DataFrame, target: str) -> pd.DataFrame:
    """
    Si target no existe, intenta mapearlo desde alias conocidos.
    """
    target_u = target.strip().upper()
    if target_u in df.columns:
        return df
    aliases = COL_ALIASES.get(target_u, [])
    for a in aliases:
        a_u = a.strip().upper()
        if a_u in df.columns:
            df[target_u] = df[a_u]
            return df
    # Si no existe, no hacemos nada
    return df

def find_header_row(df_raw: pd.DataFrame) -> int | None:
    """
    Encuentra la fila donde están los headers (busca 'GENERAL' y 'TIPO').
    """
    for i in range(min(40, len(df_raw))):
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

    # Compatibilidad con variaciones de nombres
    for target in ["IVA_APLICA", "IMPUESTO_TIPO", "MODELO", "TRATAMIENTO_IVA", "PERIODO_SERVICIO"]:
        df = coalesce_column(df, target)

    # Comprobación de columnas mínimas
    missing = [c for c in MIN_REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas requeridas: {missing}. Columnas detectadas: {list(df.columns)}")

    # Limpieza básica
    df = df.dropna(how="all").copy()

    def clean_str(col):
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype(str).str.strip()
        return df

    df = clean_str("GENERAL")
    df = clean_str("TIPO")
    df = clean_str("DEPARTAMENTO")
    df = clean_str("NATURALEZA")
    df = clean_str("PERIODICIDAD")
    df = clean_str("REGLA_FECHA")
    df = clean_str("AJUSTE FINDE")

    df["TIPO"] = df["TIPO"].str.upper()
    df["DEPARTAMENTO"] = df["DEPARTAMENTO"].str.upper()
    df["NATURALEZA"] = df["NATURALEZA"].str.upper()
    df["PERIODICIDAD"] = df["PERIODICIDAD"].str.upper()
    df["REGLA_FECHA"] = df["REGLA_FECHA"].str.upper()

    # IMPORTE numérico (pronóstico)
    df["IMPORTE"] = pd.to_numeric(df["IMPORTE"], errors="coerce").fillna(0.0)

    # LAG numérico
    df["LAG"] = pd.to_numeric(df["LAG"], errors="coerce").fillna(0).astype(int)

    # VALOR_FECHA: puede venir como fecha o como número (día)
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
    # por defecto: mensual
    return 1

def generate_events_from_catalog(
    catalog: pd.DataFrame,
    start_date: pd.Timestamp,
    months_horizon: int
) -> pd.DataFrame:
    """
    Genera movimientos a futuro desde catálogo:
    - PUNTUAL: FECHA_FIJA
    - ANUAL: FECHA_FIJA (mismo día/mes cada año)
    - SEMANAL: REGLA_FECHA=DIA_SEMANA y VALOR_FECHA=fecha ancla (p.ej primer viernes del mes)
              Genera cada 7 días, pero SOLO dentro del MES de esa ancla (evita que enero siga en febrero).
    - Mensuales (MENSUAL/BIMESTRAL/TRIMESTRAL/SEMESTRAL): según regla DIA_MES/ULTIMO_HABIL/FECHA_FIJA

    Ajuste finde: SIG_HABIL
    Lag: suma días hábiles
    """
    end_date = (start_date + pd.offsets.MonthBegin(months_horizon + 1)).normalize()
    rows = []

    for _, r in catalog.iterrows():
        periodicidad = str(r.get("PERIODICIDAD", "")).upper().strip()
        regla = str(r.get("REGLA_FECHA", "")).upper().strip()
        ajuste = str(r.get("AJUSTE FINDE", "")).upper().strip()
        lag = int(r.get("LAG", 0))

        def apply_adjustments(d: pd.Timestamp) -> pd.Timestamp:
            if ajuste == "SIG_HABIL":
                d = next_business_day(d)
            if lag:
                d = add_business_days(d, lag)
            return d

        # ---------- PUNTUAL ----------
        if periodicidad in ("PUNTUAL", "ONE-OFF", "ONEOFF"):
            if regla != "FECHA_FIJA" or pd.isna(r.get("FECHA_FIJA")):
                continue
            d = pd.Timestamp(r["FECHA_FIJA"]).normalize()
            d = apply_adjustments(d)
            if start_date <= d <= end_date:
                rows.append((d, r))
            continue

        # ---------- ANUAL ----------
        if periodicidad == "ANUAL":
            if regla == "FECHA_FIJA" and not pd.isna(r.get("FECHA_FIJA")):
                base = pd.Timestamp(r["FECHA_FIJA"]).normalize()
                year = start_date.year
                while True:
                    candidate = pd.Timestamp(year=year, month=base.month, day=base.day)
                    if candidate > end_date:
                        break
                    d = apply_adjustments(candidate)
                    if start_date <= d <= end_date:
                        rows.append((d, r))
                    year += 1
            continue

        # ---------- SEMANAL ----------
        if periodicidad == "SEMANAL":
            # Necesitamos una fecha ancla (VALOR_FECHA como fecha)
            anchor = r.get("FECHA_FIJA")
            if pd.isna(anchor):
                # si no hay fecha válida, no generamos
                continue

            anchor = pd.Timestamp(anchor).normalize()

            # Generar solo dentro del mes de la ancla (esto evita duplicidades por tener una fila por mes)
            month_start = anchor.replace(day=1)
            month_end = (month_start + pd.offsets.MonthEnd(0)).normalize()

            d = anchor
            while d <= month_end:
                dd = apply_adjustments(d)
                if start_date <= dd <= end_date:
                    rows.append((dd, r))
                d = d + pd.Timedelta(days=7)

            continue

        # ---------- PERIODICIDADES MENSUALES (incluye bimestral, etc.) ----------
        step = months_step_from_periodicidad(periodicidad)

        for m in range(0, months_horizon + 1, step):
            month_start = (start_date + pd.offsets.MonthBegin(m)).normalize()
            year = month_start.year
            month = month_start.month

            if regla == "DIA_MES":
                day = r.get("DIA_MES")
                if not day or pd.isna(day):
                    continue
                last_day = (pd.Timestamp(year=year, month=month, day=1) + pd.offsets.MonthEnd(0)).day
                day = int(min(int(day), int(last_day)))
                d = pd.Timestamp(year=year, month=month, day=day)

            elif regla == "ULTIMO_HABIL":
                d = pd.Timestamp(year=year, month=month, day=1) + pd.offsets.MonthEnd(0)
                d = next_business_day(d) if d.weekday() >= 5 else d

            elif regla == "FECHA_FIJA" and not pd.isna(r.get("FECHA_FIJA")):
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

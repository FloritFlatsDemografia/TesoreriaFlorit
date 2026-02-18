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

MIN_REQUIRED = [
    "GENERAL", "TIPO", "DEPARTAMENTO",
    "NATURALEZA", "PERIODICIDAD",
    "REGLA_FECHA", "VALOR_FECHA",
    "LAG", "AJUSTE FINDE"
]

COL_ALIASES = {
    "IMPORTE": ["IMPORTE", "IMPORTE PRONOSTICADO", "IMPORTE_PRONOSTICADO"],
    "HASTA": ["HASTA", "FECHA_FIN", "FIN", "HASTA_FECHA"],
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
    target_u = target.strip().upper()
    if target_u in df.columns:
        return df
    for a in COL_ALIASES.get(target_u, []):
        a_u = a.strip().upper()
        if a_u in df.columns:
            df[target_u] = df[a_u]
            return df
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

    for target in ["IMPORTE", "HASTA", "IVA_APLICA", "IMPUESTO_TIPO", "MODELO", "TRATAMIENTO_IVA", "PERIODO_SERVICIO"]:
        df = coalesce_column(df, target)

    missing = [c for c in MIN_REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas requeridas: {missing}. Columnas detectadas: {list(df.columns)}")

    if "IMPORTE" not in df.columns:
        raise ValueError(f"Falta columna requerida: 'IMPORTE'. Columnas detectadas: {list(df.columns)}")

    df = df.dropna(how="all").copy()

    # Normaliza strings
    for c in ["GENERAL", "TIPO", "DEPARTAMENTO", "NATURALEZA", "PERIODICIDAD", "REGLA_FECHA", "AJUSTE FINDE"]:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].astype(str).str.strip()

    df["TIPO"] = df["TIPO"].str.upper()
    df["DEPARTAMENTO"] = df["DEPARTAMENTO"].str.upper()
    df["NATURALEZA"] = df["NATURALEZA"].str.upper()
    df["PERIODICIDAD"] = df["PERIODICIDAD"].str.upper()
    df["REGLA_FECHA"] = df["REGLA_FECHA"].str.upper()
    df["AJUSTE FINDE"] = df["AJUSTE FINDE"].str.upper()

    # IMPORTE numérico (forecast)
    df["IMPORTE"] = pd.to_numeric(df["IMPORTE"], errors="coerce").fillna(0.0)

    # LAG numérico
    df["LAG"] = pd.to_numeric(df["LAG"], errors="coerce").fillna(0).astype(int)

    # DIA_MES a partir de VALOR_FECHA si es día o fecha
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

    # FECHA_FIJA (si VALOR_FECHA es parseable como fecha)
    df["FECHA_FIJA"] = pd.to_datetime(df["VALOR_FECHA"], errors="coerce").dt.normalize()

    # HASTA
    if "HASTA" in df.columns:
        df["HASTA"] = pd.to_datetime(df["HASTA"], errors="coerce").dt.normalize()
    else:
        df["HASTA"] = pd.NaT

    return df

def next_business_day(d: pd.Timestamp) -> pd.Timestamp:
    # sáb/dom -> lunes
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
    Genera eventos hasta end_date (horizonte) y respeta HASTA.
    Fix importante: MENSUAL/BIMESTRAL/TRIMESTRAL/SEMESTRAL se anclan a FECHA_FIJA del concepto,
    no al start_date, para evitar duplicados/solapes.
    """
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

        # ---------- PUNTUAL ----------
        if periodicidad in ("PUNTUAL", "ONE-OFF", "ONEOFF"):
            # usa FECHA_FIJA si VALOR_FECHA es parseable, aunque REGLA_FECHA ponga DIA_MES
            if pd.isna(r.get("FECHA_FIJA")):
                continue
            d = pd.Timestamp(r["FECHA_FIJA"]).normalize()
            d = apply_adjustments(d)
            if within_limits(d):
                rows.append((d, r))
            continue

        # ---------- ANUAL ----------
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

        # ---------- SEMANAL ----------
        if periodicidad == "SEMANAL":
            anchor = r.get("FECHA_FIJA")
            if pd.isna(anchor):
                continue
            anchor = pd.Timestamp(anchor).normalize()

            if not pd.isna(hasta):
                stop = min(hasta, end_date)
            else:
                # si no hay HASTA, por defecto solo dentro del mes del ancla (para filas "ENERO", "FEBRERO", etc.)
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

        # ---------- PERIODICIDADES POR MESES (FIX: anclado a FECHA_FIJA) ----------
        if periodicidad in ("MENSUAL", "BIMESTRAL", "BIMENSUAL", "TRIMESTRAL", "SEMESTRAL"):
            step = months_step_from_periodicidad(periodicidad)

            # ancla: FECHA_FIJA (que viene de VALOR_FECHA)
            if pd.isna(r.get("FECHA_FIJA")) and regla != "DIA_MES":
                # sin ancla no podemos generar
                continue

            # base_date:
            # - si FECHA_FIJA existe, manda
            # - si no, construimos una base usando el día (DIA_MES) y el mes/año de start_date para iniciar
            if not pd.isna(r.get("FECHA_FIJA")):
                base_date = pd.Timestamp(r["FECHA_FIJA"]).normalize()
            else:
                day = r.get("DIA_MES")
                if not day or pd.isna(day):
                    continue
                # inicia en el mes de start_date
                y, m = start_date.year, start_date.month
                last_day = (pd.Timestamp(year=y, month=m, day=1) + pd.offsets.MonthEnd(0)).day
                base_date = pd.Timestamp(year=y, month=m, day=min(int(day), int(last_day))).normalize()

            current = base_date

            # avanza hasta cubrir el horizonte
            while current <= end_date:
                # construir fecha del evento según regla
                if regla == "DIA_MES":
                    # usa el día de base_date y ajusta al último día del mes
                    y, m = current.year, current.month
                    day = base_date.day  # importante: queda fijado por el ancla
                    last_day = (pd.Timestamp(year=y, month=m, day=1) + pd.offsets.MonthEnd(0)).day
                    d = pd.Timestamp(year=y, month=m, day=min(int(day), int(last_day))).normalize()

                elif regla == "ULTIMO_HABIL":
                    d = pd.Timestamp(year=current.year, month=current.month, day=1) + pd.offsets.MonthEnd(0)
                    d = next_business_day(d) if d.weekday() >= 5 else d
                    d = d.normalize()

                elif regla == "FECHA_FIJA":
                    # mensual con fecha fija -> usa el día del ancla en cada mes
                    y, m = current.year, current.month
                    day = base_date.day
                    last_day = (pd.Timestamp(year=y, month=m, day=1) + pd.offsets.MonthEnd(0)).day
                    d = pd.Timestamp(year=y, month=m, day=min(int(day), int(last_day))).normalize()

                else:
                    # si no se reconoce, no generamos
                    d = None

                if d is not None:
                    d = apply_adjustments(d)
                    if within_limits(d):
                        rows.append((d, r))

                current = (current + pd.DateOffset(months=step)).normalize()

            continue

        # ---------- Default: tratar como MENSUAL (compatibilidad) ----------
        # Si alguien mete "MENSUAL" mal escrito u otra cosa, intentamos no romper:
        # pero mejor forzar en Excel.
        continue

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

    return out.sort_values("FECHA").reset_index(drop=True)

def compute_balance(df: pd.DataFrame, starting_balance: float) -> pd.DataFrame:
    df = df.copy()
    df["COBROS"] = df.apply(lambda x: x["IMPORTE"] if x["TIPO"] == "INGRESO" else 0.0, axis=1)
    df["PAGOS"] = df.apply(lambda x: x["IMPORTE"] if x["TIPO"] == "GASTO" else 0.0, axis=1)
    df["NETO"] = df["COBROS"] - df["PAGOS"]
    df["SALDO"] = starting_balance + df["NETO"].cumsum()
    return df

# --- Extensión ingresos año anterior (semanal viernes) limitada por end_date ---
def extend_ingresos_from_previous_year_weekly(generated: pd.DataFrame, growth: float, end_date: pd.Timestamp) -> pd.DataFrame:
    gen = generated.copy()
    gen["YEAR"] = gen["FECHA"].dt.year
    gen["MONTH"] = gen["FECHA"].dt.month

    years = sorted(gen["YEAR"].dropna().unique().tolist())
    if len(years) < 1:
        return generated

    base_year = years[0]
    target_year = base_year + 1

    base_ing = gen[(gen["TIPO"] == "INGRESO") & (gen["YEAR"] == base_year)]
    if base_ing.empty:
        return generated

    monthly_base = base_ing.groupby(["MONTH", "DEPARTAMENTO"], as_index=False)["IMPORTE"].sum()
    monthly_base["IMPORTE"] = monthly_base["IMPORTE"] * float(growth)

    rows = []
    for _, rr in monthly_base.iterrows():
        month = int(rr["MONTH"])
        dept = rr["DEPARTAMENTO"]
        total_mes = float(rr["IMPORTE"])

        start = pd.Timestamp(year=target_year, month=month, day=1)
        end = (start + pd.offsets.MonthEnd(0)).normalize()
        fridays = pd.date_range(start, end, freq="W-FRI")
        n_viernes = len(fridays)
        if n_viernes == 0:
            continue

        importe_viernes = total_mes / n_viernes

        for d in fridays:
            d = pd.Timestamp(d).normalize()
            if d > end_date:
                continue
            rows.append({
                "FECHA": d,
                "CONCEPTO": f"INGRESO AUTO {target_year} (base {base_year})",
                "TIPO": "INGRESO",
                "DEPARTAMENTO": dept,
                "IMPORTE": float(importe_viernes),
                "NATURALEZA": "AUTO"
            })

    if not rows:
        return generated

    gen2 = pd.concat([generated, pd.DataFrame(rows)], ignore_index=True)
    return gen2.sort_values("FECHA").reset_index(drop=True)

# -----------------------------
# Sidebar inputs
# -----------------------------
st.sidebar.header("Inputs")
saldo_fecha = st.sidebar.date_input("Fecha del saldo (hoy)", value=date.today())
saldo_hoy = st.sidebar.number_input("Saldo actual en banco (€)", min_value=-1e12, max_value=1e12, value=0.0, step=100.0)
months_horizon = st.sidebar.slider("Horizonte forecast (meses)", min_value=1, max_value=36, value=12)

extend_ingresos = st.sidebar.checkbox("Extender INGRESOS usando año anterior (AUTO)", value=True)
growth = st.sidebar.number_input("Factor crecimiento año extendido", min_value=0.0, max_value=3.0, value=1.00, step=0.01)

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
end_ts = (start_ts + pd.offsets.MonthBegin(months_horizon + 1)).normalize()

generated = generate_events_from_catalog(
    catalog=catalog,
    start_date=start_ts,
    months_horizon=months_horizon
)

# Extensión automática (limitada por end_ts)
if extend_ingresos and not generated.empty:
    generated = extend_ingresos_from_previous_year_weekly(generated, growth, end_ts)

# Dedupe exacto (por si el Excel trae filas repetidas o por seguridad)
if dedupe_exact and not generated.empty:
    generated = generated.drop_duplicates(subset=["FECHA", "CONCEPTO", "TIPO", "DEPARTAMENTO", "IMPORTE"], keep="first")

if generated.empty:
    st.warning("No se generaron movimientos (revisa PERIODICIDAD / REGLA_FECHA / VALOR_FECHA / HASTA).")
    st.dataframe(catalog.head(50), use_container_width=True)
    st.stop()

# -----------------------------
# Filtros
# -----------------------------
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

# -----------------------------
# Vista "tipo Excel": VTO. PAGO / CONCEPTO / COBROS / PAGOS / SALDO / PREVISION MES
# -----------------------------
st.subheader("Movimientos (formato tesorería)")

view = consolidado.copy()
view["VTO. PAGO"] = view["FECHA"].dt.strftime("%d-%m-%y")

# Previsión del mes = NETO total del mes (cobros - pagos)
tmp = view.copy()
tmp["MES"] = tmp["FECHA"].dt.to_period("M").astype(str)
monthly_forecast = tmp.groupby("MES", as_index=False).agg(PREVISION_MES=("NETO", "sum"))

view["MES"] = view["FECHA"].dt.to_period("M").astype(str)
view = view.merge(monthly_forecast, on="MES", how="left")

view_out = view[["VTO. PAGO", "CONCEPTO", "COBROS", "PAGOS", "SALDO", "PREVISION_MES"]].copy()
st.dataframe(view_out, use_container_width=True)

# -----------------------------
# Resumen mensual
# -----------------------------
st.subheader("Resumen mensual")
consolidado["MES"] = consolidado["FECHA"].dt.to_period("M").astype(str)
monthly = consolidado.groupby("MES", as_index=False).agg(
    COBROS=("COBROS", "sum"),
    PAGOS=("PAGOS", "sum"),
    NETO=("NETO", "sum"),
)
st.dataframe(monthly, use_container_width=True)

# -----------------------------
# Export
# -----------------------------
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

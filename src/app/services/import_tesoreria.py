import pandas as pd

REQUIRED_COLS = ["VTO. PAGO", "CONCEPTO", "COBROS", "PAGOS"]

def load_tesoreria_excel(uploaded_file) -> pd.DataFrame:
    """
    Lee un Excel .xls/.xlsx y devuelve un DF normalizado.

    Normaliza:
    - Convierte VTO. PAGO a datetime
    - Asegura COBROS/PAGOS numéricos
    - Ordena por fecha
    - Recalcula SALDO de forma consistente (SALDO_CALCULADO)
    - Si existe SALDO en el Excel, calcula DIF_SALDO
    - PREVISIÓN: si no existe, se iguala a SALDO
    """
    name = getattr(uploaded_file, "name", "").lower()

    if name.endswith(".xls"):
        df = pd.read_excel(uploaded_file, engine="xlrd")
    else:
        df = pd.read_excel(uploaded_file, engine="openpyxl")

    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(
            f"Faltan columnas requeridas: {missing}. "
            f"Columnas encontradas: {list(df.columns)}"
        )

    df = df.dropna(how="all").copy()

    df["VTO. PAGO"] = pd.to_datetime(df["VTO. PAGO"], errors="coerce")
    df["COBROS"] = pd.to_numeric(df["COBROS"], errors="coerce").fillna(0.0)
    df["PAGOS"] = pd.to_numeric(df["PAGOS"], errors="coerce").fillna(0.0)
    df["CONCEPTO"] = df["CONCEPTO"].astype(str).fillna("")

    df = df.sort_values("VTO. PAGO", ascending=True).reset_index(drop=True)

    df["SALDO_CALCULADO"] = (df["COBROS"] - df["PAGOS"]).cumsum()

    if "SALDO" in df.columns:
        df["SALDO"] = pd.to_numeric(df["SALDO"], errors="coerce")
        df["DIF_SALDO"] = (df["SALDO"] - df["SALDO_CALCULADO"]).round(2)
    else:
        df["SALDO"] = df["SALDO_CALCULADO"]
        df["DIF_SALDO"] = 0.0

    if "PREVISIÓN" not in df.columns:
        df["PREVISIÓN"] = df["SALDO"]

    return df

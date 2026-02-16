import pandas as pd
import re
from datetime import datetime, date
from typing import Tuple, Optional, Dict, Any

REQUIRED = [
    "concepto",
    "tipo",
    "departamento",
    "importe",
    "naturaleza",
    "periodicidad",
    "regla_fecha",
    "valor_fecha",
    "lag",
    "ajuste_finde",
    "periodo_servicio",
]

ALIASES = {
    "concepto": ["concepto", "general", "descripcion", "descripción"],
    "tipo": ["tipo"],
    "departamento": ["departamento", "depto"],
    "importe": ["importe", "cuantia", "cantidad"],
    "naturaleza": ["naturaleza"],
    "periodicidad": ["periodicidad", "periodico", "periódico"],
    "regla_fecha": ["regla_fecha", "regla fecha"],
    "valor_fecha": ["valor_fecha", "valor fecha", "fecha ingreso o gasto", "dia o fecha base"],
    "lag": ["lag"],
    "ajuste_finde": ["ajuste finde", "ajuste_finde"],
    "periodo_servicio": ["periodo_servicio", "periodo servicio", "mes", "mes prestado"],
    # optional
    "iva_en_factura": ["iva_en_factura", "iva en factura"],
    "iva_pct": ["iva_%", "iva_pct", "iva %"],
    "iva_sentido": ["iva_sentido", "iva sentido"],
    "tratamiento_iva": ["tratamineto_iva", "tratamiento_iva", "tratamiento iva", "tratamineto iva"],
    "impuesto_tipo": ["impuesto_tipo", "impuesto tipo"],
    "modelo": ["modelo", "modelo_impuesto", "modelo impuesto"],
}

def _normalize(s: str) -> str:
    s = str(s).strip().lower()
    s = re.sub(r"\s+", "_", s)
    return s

def _map_columns(cols):
    norm = {_normalize(c): c for c in cols}
    mapping = {}
    for canonical, alts in ALIASES.items():
        for a in alts:
            na = _normalize(a)
            if na in norm:
                mapping[canonical] = norm[na]
                break
    return mapping

def _find_header_row(df: pd.DataFrame) -> int:
    # Heurística: fila donde aparece "GENERAL" en primera col o "TIPO" en segunda
    for i in range(min(len(df), 50)):
        row = df.iloc[i].astype(str).str.strip().str.upper().tolist()
        if len(row) >= 2 and (row[0] == "GENERAL" or row[0] == "CONCEPTO"):
            return i
        if "TIPO" in row and ("DEPARTAMENTO" in row or "DEPTO" in row):
            return i
    raise ValueError("No encuentro la fila de cabecera del catálogo (esperaba una fila con 'GENERAL'/'CONCEPTO' y 'TIPO').")

def _read_catalog(sheet_df: pd.DataFrame) -> pd.DataFrame:
    hdr = _find_header_row(sheet_df)
    headers = sheet_df.iloc[hdr].tolist()
    data = sheet_df.iloc[hdr+1:].copy()
    data.columns = headers
    # drop fully empty rows
    data = data.dropna(how="all")
    # keep until first row where 'TIPO' empty and 'GENERAL' empty
    first_col = headers[0]
    if first_col in data.columns:
        data = data[~(data[first_col].isna() & data.get("TIPO", pd.Series([None]*len(data))).isna())]
    # normalize columns
    colmap = _map_columns(data.columns)
    # validate required
    missing = [c for c in REQUIRED if c not in colmap]
    if missing:
        raise ValueError(f"Faltan columnas requeridas en catálogo: {missing}. Columnas detectadas: {list(data.columns)}")
    out = pd.DataFrame()
    for k, col in colmap.items():
        out[k] = data[col]
    # type cleanup
    out["concepto"] = out["concepto"].astype(str).str.strip()
    out["tipo"] = out["tipo"].astype(str).str.strip().str.upper()
    out["departamento"] = out["departamento"].astype(str).str.strip().str.upper()
    out["naturaleza"] = out["naturaleza"].astype(str).str.strip().str.upper()
    out["periodicidad"] = out["periodicidad"].astype(str).str.strip().str.upper()
    out["regla_fecha"] = out["regla_fecha"].astype(str).str.strip().str.upper()
    # valor_fecha can be date or number
    out["valor_fecha_raw"] = out["valor_fecha"]
    out["lag"] = pd.to_numeric(out["lag"], errors="coerce").fillna(0).astype(int)
    out["ajuste_finde"] = out["ajuste_finde"].astype(str).str.strip().str.upper()
    out["periodo_servicio"] = out["periodo_servicio"].astype(str).str.strip().str.upper()

    out["importe"] = pd.to_numeric(out["importe"], errors="coerce")
    # Optional columns
    for opt in ["iva_en_factura","iva_pct","iva_sentido","tratamiento_iva","impuesto_tipo","modelo"]:
        if opt in colmap:
            out[opt] = data[colmap[opt]]
        else:
            out[opt] = None
    return out.reset_index(drop=True)

def _read_import(xl: pd.ExcelFile) -> Optional[pd.DataFrame]:
    # Search for likely sheets
    for name in xl.sheet_names:
        if _normalize(name) in ["import_turistico","import_movimientos","movimientos_import","import"]:
            df = pd.read_excel(xl, sheet_name=name)
            df = df.dropna(how="all")
            if not len(df):
                return None
            # expected columns: Fecha, Concepto, Tipo, Importe (tolerant)
            cols = [_normalize(c) for c in df.columns]
            # map
            def find(colname_opts):
                for o in colname_opts:
                    if _normalize(o) in cols:
                        return df.columns[cols.index(_normalize(o))]
                return None
            c_fecha = find(["fecha","vto_pago","vto. pago"])
            c_conc = find(["concepto","concpeto","descripcion","general"])
            c_tipo = find(["tipo"])
            c_imp = find(["importe","cobro","pago","importe_total"])
            if c_fecha is None or c_conc is None:
                return None
            out = pd.DataFrame()
            out["fecha"] = pd.to_datetime(df[c_fecha], errors="coerce").dt.date
            out["concepto"] = df[c_conc].astype(str)
            if c_tipo is not None:
                out["tipo"] = df[c_tipo].astype(str).str.upper()
            else:
                out["tipo"] = "UNKNOWN"
            # if import has cobro/pago separate:
            if find(["cobro"]) is not None or find(["pago"]) is not None:
                cob_col = find(["cobro"])
                pag_col = find(["pago"])
                out["cobro"] = pd.to_numeric(df[cob_col], errors="coerce").fillna(0.0) if cob_col else 0.0
                out["pago"] = pd.to_numeric(df[pag_col], errors="coerce").fillna(0.0) if pag_col else 0.0
            else:
                out["importe"] = pd.to_numeric(df[c_imp], errors="coerce").fillna(0.0) if c_imp else 0.0
            out["departamento"] = "TURISTICO"
            return out.dropna(subset=["fecha"])
    return None

def _read_params(xl: pd.ExcelFile) -> Dict[str, Any]:
    for name in xl.sheet_names:
        if _normalize(name) in ["parametros","parameters","config"]:
            df = pd.read_excel(xl, sheet_name=name, header=None)
            df = df.dropna(how="all")
            params = {}
            for _, r in df.iterrows():
                k = str(r.iloc[0]).strip().lower()
                v = r.iloc[1] if len(r) > 1 else None
                if "saldo" in k:
                    params["saldo_inicial"] = float(v)
                if "fecha" in k and "inicio" in k:
                    try:
                        params["start_date"] = pd.to_datetime(v).date()
                    except Exception:
                        pass
            return params
    return {}

def load_workbook_tables(uploaded_file) -> Tuple[pd.DataFrame, Optional[pd.DataFrame], Dict[str, Any]]:
    # load workbook
    name = getattr(uploaded_file, "name", "").lower()
    if name.endswith(".xls"):
        xl = pd.ExcelFile(uploaded_file, engine="xlrd")
    else:
        xl = pd.ExcelFile(uploaded_file, engine="openpyxl")

    # Catalog: choose first sheet by default
    catalog_sheet = xl.sheet_names[0]
    sheet_df = pd.read_excel(xl, sheet_name=catalog_sheet, header=None)
    catalog_df = _read_catalog(sheet_df)

    import_df = _read_import(xl)
    params = _read_params(xl)
    return catalog_df, import_df, params

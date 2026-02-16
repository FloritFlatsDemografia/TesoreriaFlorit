import pandas as pd
from io import BytesIO

def to_excel_bytes(consolidated: pd.DataFrame, events_df: pd.DataFrame, catalog_df: pd.DataFrame, import_df: pd.DataFrame | None):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        consolidated.to_excel(writer, sheet_name="MOVIMIENTOS_CONSOLIDADO", index=False)
        events_df.to_excel(writer, sheet_name="RECURSIVOS_GENERADOS", index=False)
        catalog_df.to_excel(writer, sheet_name="CATALOGO_NORMALIZADO", index=False)
        if import_df is not None and len(import_df):
            import_df.to_excel(writer, sheet_name="IMPORT_ORIGINAL", index=False)
    return output.getvalue()

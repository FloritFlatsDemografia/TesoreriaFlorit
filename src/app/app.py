import streamlit as st
import pandas as pd
from datetime import date

from services.io_excel import load_workbook_tables
from services.forecast import generate_recurrent_events, consolidate_and_compute_balance

st.set_page_config(page_title="APP Tesorería", layout="wide")
st.title("APP Tesorería — MVP")

uploaded = st.file_uploader("Sube el Excel (catálogo + import opcional)", type=["xlsx","xls"])

colA, colB, colC = st.columns(3)
with colA:
    horizon_months = st.number_input("Horizonte (meses)", min_value=1, max_value=36, value=12, step=1)
with colB:
    start_date = st.date_input("Fecha inicio forecast", value=date.today())
with colC:
    saldo_inicial = st.number_input("Saldo inicial (€)", value=0.0, step=1000.0, format="%.2f")

if not uploaded:
    st.info("Sube un Excel para empezar.")
    st.stop()

try:
    catalog_df, import_df, params = load_workbook_tables(uploaded)

    # Si hay PARAMETROS en el excel, respetamos si el usuario no tocó inputs (o si saldo=0)
    if params:
        if params.get("start_date") and start_date == date.today():
            start_date = params["start_date"]
        if params.get("saldo_inicial") is not None and saldo_inicial == 0.0:
            saldo_inicial = float(params["saldo_inicial"])

    st.subheader("1) Catálogo recurrente (detectado)")
    st.dataframe(catalog_df, use_container_width=True, height=240)

    if import_df is not None and len(import_df):
        st.subheader("2) Import (movimientos reales) — opcional")
        st.dataframe(import_df, use_container_width=True, height=220)

    events_df = generate_recurrent_events(catalog_df, start_date=start_date, horizon_months=horizon_months)

    st.subheader("3) Recurrentes generados (calendario)")
    st.dataframe(events_df, use_container_width=True, height=240)

    consolidated = consolidate_and_compute_balance(
        events_df=events_df,
        import_df=import_df,
        saldo_inicial=saldo_inicial
    )

    st.subheader("4) Consolidado + saldo")
    st.dataframe(consolidated, use_container_width=True, height=420)

    # KPIs
    k1,k2,k3,k4 = st.columns(4)
    with k1:
        st.metric("Filas", f"{len(consolidated)}")
    with k2:
        st.metric("Ingresos", f"{consolidated['cobro'].sum():,.2f} €")
    with k3:
        st.metric("Gastos", f"{consolidated['pago'].sum():,.2f} €")
    with k4:
        st.metric("Saldo final", f"{consolidated['saldo'].iloc[-1]:,.2f} €" if len(consolidated) else f"{saldo_inicial:,.2f} €")

    # Export
    xlsx_bytes = None
    from services.exporter import to_excel_bytes
    xlsx_bytes = to_excel_bytes(consolidated, events_df, catalog_df, import_df)

    st.download_button(
        "Descargar Excel consolidado",
        data=xlsx_bytes,
        file_name="tesoreria_consolidado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

except Exception as e:
    st.error(f"Error: {e}")
    st.exception(e)

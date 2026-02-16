import streamlit as st

from services.import_tesoreria import load_tesoreria_excel

st.set_page_config(page_title="APP Tesorería", layout="wide")
st.title("APP Tesorería — Importación")

uploaded = st.file_uploader("Sube el Excel de tesorería (.xls o .xlsx)", type=["xls", "xlsx"])

if uploaded:
    try:
        df = load_tesoreria_excel(uploaded)
        st.success("Archivo importado y normalizado.")

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Filas", len(df))
        with col2:
            st.metric("Cobros (suma)", f"{df['COBROS'].sum():,.2f}")
        with col3:
            st.metric("Pagos (suma)", f"{df['PAGOS'].sum():,.2f}")

        if df["DIF_SALDO"].abs().max() > 0.01:
            st.warning("Hay diferencias entre SALDO del Excel y SALDO_CALCULADO (revisar).")

        st.subheader("Vista previa")
        st.dataframe(
            df[["VTO. PAGO", "CONCEPTO", "COBROS", "PAGOS", "SALDO", "SALDO_CALCULADO", "DIF_SALDO", "PREVISIÓN"]],
            use_container_width=True
        )

        st.download_button(
            "Descargar normalizado (CSV)",
            data=df.to_csv(index=False).encode("utf-8"),
            file_name="tesoreria_normalizada.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"Error al importar: {e}")
else:
    st.info("Sube el Excel para empezar.")

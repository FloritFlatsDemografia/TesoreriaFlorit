# Decisiones

- **Entrypoint**: Streamlit (`src/app/app.py`).
- **Importación**: se soporta `.xls` (engine `xlrd`) y `.xlsx` (engine `openpyxl`).
- **Normalización**:
  - `VTO. PAGO` -> datetime
  - `COBROS`, `PAGOS` -> numérico (NaN -> 0)
  - Orden por fecha ascendente
- **Saldo**:
  - `SALDO_CALCULADO` = cumsum(COBROS - PAGOS)
  - Si existe `SALDO` en el Excel, se compara y se crea `DIF_SALDO`.
  - `PREVISIÓN` se rellena si falta (igual a `SALDO`).

Pendiente (siguientes iteraciones):
- saldo inicial configurable
- generación de cobros recurrentes (remesas 25 y 3 + día hábil + lag)
- visualización de forecast y mínimos de caja

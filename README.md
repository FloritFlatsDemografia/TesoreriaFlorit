# APP TESORERIA (scaffold)

Este repositorio contiene un primer "vertical slice" de la APP Tesorería:
- Importación de Excel de tesorería (.xls / .xlsx)
- Validación de columnas
- Normalización de tipos
- Re-cálculo de saldo
- Vista previa + export CSV

## Ejecutar en local

```bash
pip install -r requirements.txt
streamlit run src/app/app.py
```

## Formato de entrada (Excel)

Hoja: `TESORERÍA` (o primera hoja)

Columnas mínimas requeridas:
- `VTO. PAGO` (fecha)
- `CONCEPTO` (texto)
- `COBROS` (numérico)
- `PAGOS` (numérico)

Columnas opcionales:
- `SALDO` (si existe, se compara con el saldo calculado)
- `PREVISIÓN` (si no existe, se rellena = SALDO)

## Estructura

- `src/app/app.py`: entrypoint Streamlit
- `src/app/services/import_tesoreria.py`: carga/normaliza el Excel
- `configs/`: configuración (reservado)
- `docs/`: documentación (reservado)
- `scripts/`: utilidades (reservado)

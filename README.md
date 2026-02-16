# APP TESORERÍA (MVP Streamlit)

## Ejecutar en local
```bash
pip install -r requirements.txt
streamlit run src/app/app.py
```

## Qué hace
- Carga un Excel con:
  - Catálogo recurrente (como el de *Tesoreria Excel Nacho.xlsx*)
  - Opcional: hoja `IMPORT_TURISTICO` (movimientos reales importados del software)
  - Opcional: hoja `PARAMETROS` (Saldo inicial y fecha inicio)
- Genera el calendario futuro de recurrentes (mensual/anual/puntual) para un horizonte configurable.
- Consolida import + recurrentes y calcula el saldo acumulado.
- Exporta a Excel el consolidado.

## Convención mínima (Catálogo)
Columnas esperadas (cabecera):
- CONCEPTO / GENERAL
- TIPO (INGRESO/GASTO)
- DEPARTAMENTO
- IMPORTE
- NATURALEZA
- PERIODICIDAD (MENSUAL/ANUAL/PUNTUAL/TRIMESTRAL)
- REGLA_FECHA (DIA_MES/FECHA_FIJA/ULTIMO_HABIL)
- VALOR_FECHA (día 1-31 o fecha YYYY-MM-DD)
- LAG (días hábiles)
- AJUSTE FINDE (SIG HABIL / ANT HABIL)
- PERIODO_SERVICIO (MES ACTUAL / MES ANTERIOR / MES SIGUIENTE)
- IVA_EN_FACTURA, IVA_%, IVA_SENTIDO, TRATAMINETO_IVA (opcionales)

La app es tolerante a algunos nombres; valida y te avisa si falta algo.

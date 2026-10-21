# Conciliación Contable Multiorigen

    streamlit run app.py      # ejecutar la app
    pytest                    # tests del núcleo (no requieren Streamlit)

## Estructura

    app.py                  Entrada: sidebar + despacho a la página elegida
    config.py               ERP_CONFIGS, tolerancias y constantes
    core/                   Lógica de negocio (sin Streamlit)
      cleaning.py             safe_float, clean_cuit, clean_entity_key, parse_monto_txt
      io_utils.py             decodificar_texto, validar_formato_archivo
      columns.py              ColumnasMayor / ColumnasExterno, obtener_monto_mayor
      preparacion.py          montos y claves de cruce
      matching.py             estrategias de cruce (1a1, lote CUIT, lote razón social)
      reconciler.py           orquesta todo: procesar_archivos()
    parsers/                Archivos de entrada → DataFrames
      percepciones.py         txt IIBB, padrón de proveedores
      autodealer.py           Mayor crudo AutoDealer → xlsx
    reports/                Salida Excel
      styles.py / excel_helpers.py / excel_report.py
    ui/                     Streamlit (una página por módulo)
      components.py, retenciones_page.py, percepciones_page.py, formatear_page.py

## Cómo extender

- **Nuevo ERP**: agregar una entrada en `ERP_CONFIGS` (config.py). Nada más.
- **Nueva estrategia de cruce**: escribir la función en `core/matching.py` y sumarla en `ejecutar_cruce` (core/reconciler.py).
- **Nuevo módulo/pantalla**: crear `ui/<nombre>_page.py` con `TITULO` y `render()` y agregarlo a `PAGINAS` en app.py.

"""Componentes de Streamlit reutilizables entre páginas."""
import pandas as pd
import streamlit as st

from config import MIME_XLSX


def mostrar_mapeo_erp(cfg, modo_total=False):
    """
    Muestra en la sidebar el mapeo de columnas activo del ERP.
    modo_total=True: variante del módulo IIBB (cruza por la columna de total con signo).
    """
    st.markdown("### 📌 Mapeo activo")
    st.caption(f"**Asiento:** `{cfg.get('asiento')}`")
    st.caption(f"**Fecha:** `{cfg.get('fecha')}`")
    st.caption(f"**Detalle / Concepto:** `{cfg.get('detalle')}`")
    if modo_total:
        st.caption(f"**Ref Secundario:** `{cfg.get('ref_secundaria')}`")
    else:
        st.caption(f"**Referencia:** `{cfg.get('referencia')}`")
        if cfg.get("ref_secundaria"):
            st.caption(f"**Ref Secundario:** `{cfg.get('ref_secundaria')}`")
    st.caption(f"**Razón Social:** `{cfg.get('entidad')}`")
    st.caption(f"**CUIT:** `{cfg.get('cuit')}`")
    if modo_total:
        st.caption(f"**Monto de cruce:** `{cfg.get('saldo')}`")
    else:
        st.caption(f"**Debe:** `{cfg.get('debe')}` | **Haber:** `{cfg.get('haber')}`")


def mostrar_metricas(stats, etiqueta_externo):
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Conciliados 1 a 1", stats["cnt_1a1"])
    m2.metric("Conciliados por Lote", stats["cnt_lote"])
    m3.metric(f"Pendientes {etiqueta_externo}", stats["cnt_arca"])
    m4.metric("Pendientes Mayor", stats["cnt_mayor"])


def boton_descarga_xlsx(label, data, file_name):
    st.download_button(
        label=label,
        data=data,
        file_name=file_name,
        mime=MIME_XLSX,
        use_container_width=True,
    )


def _tabla_o_mensaje(registros, msg_vacio, es_exito=False):
    if registros:
        st.dataframe(pd.DataFrame(registros), use_container_width=True)
    elif es_exito:
        st.success(msg_vacio)
    else:
        st.info(msg_vacio)


def mostrar_resultados(stats, etiqueta_externo, msg_sin_pend_externo):
    """Previsualización en 4 pestañas del resultado del cruce."""
    st.markdown("### 🔍 Previsualización de Resultados")
    tab1, tab2, tab3, tab4 = st.tabs([
        "Exactas 1a1", "Por Lote / CUIT", f"Pendientes {etiqueta_externo}", "Pendientes Mayor"
    ])
    with tab1:
        _tabla_o_mensaje(stats["conciliadas_1a1"], "No se encontraron registros 1 a 1 exactos.")
    with tab2:
        _tabla_o_mensaje(stats["conciliadas_lote"], "No se encontraron registros acumulados por lote.")
    with tab3:
        _tabla_o_mensaje(stats["pendientes_arca"], msg_sin_pend_externo, es_exito=True)
    with tab4:
        _tabla_o_mensaje(stats["pendientes_mayor"], "🎉 ¡No hay registros pendientes en el Mayor!", es_exito=True)


def ejecutar_con_manejo_de_errores(funcion):
    """Ejecuta `funcion` y muestra los errores de forma uniforme en la UI."""
    try:
        funcion()
    except ValueError as ve:
        st.error(f"⚠️ {str(ve)}", icon="🛑")
    except Exception as error:
        st.error(f"❌ Ocurrió un error inesperado al procesar los archivos: {str(error)}", icon="❌")

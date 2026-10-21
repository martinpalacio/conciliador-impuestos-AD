"""Página: Conciliador de Retenciones (Mayor vs MisRetenciones ARCA)."""
import streamlit as st

from config import ERP_CONFIGS
from core.reconciler import procesar_archivos
from ui.components import (
    boton_descarga_xlsx, ejecutar_con_manejo_de_errores, mostrar_mapeo_erp,
    mostrar_metricas, mostrar_resultados,
)

TITULO = "🔄 Conciliador de Retenciones"


def _sidebar():
    with st.sidebar:
        st.header("⚙️ Configuración del Cruce")
        erp = st.radio(
            "Selecciona el ERP del Mayor:",
            options=list(ERP_CONFIGS.keys()),
            help="Adapta dinámicamente las referencias de los encabezados según el sistema contable emisor.",
        )
        st.markdown("---")
        mostrar_mapeo_erp(ERP_CONFIGS[erp])
    return erp


def _procesar(file_mayor, file_arca, erp):
    excel_bytes, stats = procesar_archivos(file_mayor, file_arca, erp)

    st.success("✅ ¡Conciliación completada con éxito!")
    mostrar_metricas(stats, "ARCA")
    boton_descarga_xlsx(
        "📥 Descargar Reporte Conciliado en Excel", excel_bytes, f"Conciliacion_Retenciones_{erp}.xlsx"
    )
    mostrar_resultados(stats, "ARCA", "🎉 ¡No hay retenciones pendientes en ARCA!")


def render():
    st.title("📊 Conciliador Contable de Retenciones (ARCA vs ERP)")
    st.markdown(
        "Selecciona el **ERP de origen**, carga los reportes de **Mayor** y **Mis Retenciones (ARCA)**, "
        "y genera el cruce automático con validación por **CUIT y Razón Social**."
    )
    st.warning(
        "⚠️ **ATENCIÓN:** Solo se admiten archivos en formato moderno **.xlsx**. Si tus archivos son `.xls` "
        "(muy común en ARCA/AFIP), ábrelos en Excel y guárdalos como `.xlsx` antes de subirlos.",
        icon="⚠️",
    )

    erp = _sidebar()

    col1, col2 = st.columns(2)
    with col1:
        file_mayor = st.file_uploader(f"Subir Mayor (.xlsx) ({erp})", type=["xlsx"], key="mayor")
    with col2:
        file_arca = st.file_uploader("Subir MisRetenciones (.xlsx) (ARCA)", type=["xlsx"], key="arca")

    if file_mayor and file_arca:
        st.info(f"💡 Listo para realizar el cruce utilizando la estructura de **{erp}**.")

        if st.button("⚙️ Procesar Conciliación", type="primary", use_container_width=True):
            with st.spinner("Procesando retenciones y aplicando cruce por CUIT..."):
                ejecutar_con_manejo_de_errores(lambda: _procesar(file_mayor, file_arca, erp))

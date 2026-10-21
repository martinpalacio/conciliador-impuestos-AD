"""Página: Formateador del Mayor crudo de AutoDealer."""
import streamlit as st

from parsers.autodealer import parse_autodealer_xls
from ui.components import boton_descarga_xlsx

TITULO = "✨ Formatear Mayor (XLS)"


def render():
    st.title("✨ Formateador Automático de Mayor (AutoDealer)")
    st.markdown(
        "Sube tu archivo crudo **.xls** exportado del sistema. "
        "El motor inteligente extraerá números de comprobante, entidades sociales de longitud variable y montos, "
        "generando un **Excel tabulado (.xlsx)** listo para el conciliador."
    )

    file_xls = st.file_uploader("📂 Sube el archivo crudo del Mayor", type=["xls", "xlsx", "txt", "prn"])
    if not file_xls:
        return

    st.info("💡 Archivo leído. Listo para ser interpretado y estructurado.")
    if st.button("🪄 Ejecutar Formateo", type="primary", use_container_width=True):
        with st.spinner("Desglosando registros del mayor y limpiando columnas..."):
            try:
                excel_formateado = parse_autodealer_xls(file_xls)
                st.success("✅ ¡Archivo procesado y estructurado con éxito!")
                boton_descarga_xlsx(
                    "📥 Descargar Mayor Formateado (.xlsx)", excel_formateado, "Mayor_AutoDealer_Formateado.xlsx"
                )
                st.markdown(
                    "⚠️ *Recomendación: Verifica que el formato descargado sea el esperado "
                    "(especialmente la separación DEBE/HABER/SALDO) antes de importarlo al conciliador.*"
                )
            except Exception as e:
                st.error(f"❌ Ocurrió un error al procesar el archivo: {str(e)}")

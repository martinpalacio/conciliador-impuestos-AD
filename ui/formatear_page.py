"""Página: Formateador del Mayor crudo de AutoDealer."""
import streamlit as st

from parsers.autodealer import formatear_mayor
from reports.mayor_excel import generar_excel_mayor
from ui.components import boton_descarga_xlsx

TITULO = "✨ Formatear Mayor (XLS)"


def _fmt(n):
    return "—" if n is None else f"$ {n:,.2f}"


def mostrar_control(control):
    """Resumen contable: la prueba de que el formateo no perdió ni alteró importes."""
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Movimientos", control["movimientos"])
    c2.metric("Total Debe", _fmt(control["total_debe"]))
    c3.metric("Total Haber", _fmt(control["total_haber"]))
    c4.metric("Saldo final ERP", _fmt(control["saldo_final"]))

    if control["ok"] is True:
        st.success(
            "✅ **Control OK:** saldo inicial + Debe − Haber = saldo final del ERP "
            f"({_fmt(control['saldo_calculado'])})."
        )
    elif control["ok"] is False:
        st.error(
            f"❌ **El control NO cierra:** calculado {_fmt(control['saldo_calculado'])} vs saldo final del ERP "
            f"{_fmt(control['saldo_final'])} (diferencia {_fmt(control['diferencia'])}). "
            "No uses este archivo sin revisar el origen."
        )
    else:
        st.warning("⚠️ El archivo no trae 'SALDO FINAL': no se pudo controlar contra el ERP.")

    if control["sin_tres_partes"]:
        st.info(
            f"ℹ️ {control['sin_tres_partes']} movimiento(s) tienen un concepto sin el formato "
            "«comprobante / referencia / entidad» (ej. asiento de apertura, DDJJ IVA): DETALLE y ENTIDAD "
            "llevan el texto completo."
        )


def mostrar_resultado(resultado):
    if resultado.metodo == "columnas":
        st.success("✅ ¡Archivo procesado por columnas (lectura exacta)!")
        mostrar_control(resultado.control)
    else:
        st.warning(
            "⚠️ El archivo no tiene el formato de columnas del export nativo, por lo que se interpretó "
            "**como texto**. En este modo los importes pueden quedar mal asignados y no hay control "
            "contra el saldo final: **verificá DEBE/HABER/SALDO** antes de usarlo."
        )

    boton_descarga_xlsx(
        "📥 Descargar Mayor Formateado (.xlsx)",
        generar_excel_mayor(resultado.df),
        "Mayor_AutoDealer_Formateado.xlsx",
    )
    st.markdown("### 🔍 Previsualización")
    st.dataframe(resultado.df, use_container_width=True)


def render():
    st.title("✨ Formateador Automático de Mayor (AutoDealer)")
    st.markdown(
        "Sube tu archivo crudo **.xls** exportado del sistema (MayorPRN). "
        "Se genera un **Excel tabulado (.xlsx)** listo para el conciliador, con un **control contable** "
        "contra el saldo inicial y final del ERP."
    )

    file_xls = st.file_uploader("📂 Sube el archivo crudo del Mayor", type=["xls", "xlsx", "txt", "prn"])
    if not file_xls:
        return

    st.info("💡 Archivo leído. Listo para ser interpretado y estructurado.")
    if st.button("🪄 Ejecutar Formateo", type="primary", use_container_width=True):
        with st.spinner("Desglosando registros del mayor y limpiando columnas..."):
            try:
                mostrar_resultado(formatear_mayor(file_xls))
            except Exception as e:
                st.error(f"❌ Ocurrió un error al procesar el archivo: {str(e)}")

"""Página: Conciliador de Percepciones IIBB (txt del sistema vs Mayor Bejerman)."""
import io

import streamlit as st

from config import ERP_CONFIGS, MIME_XLSX, SIN_RAZON_SOCIAL
from core.reconciler import procesar_archivos
from parsers.percepciones import (
    agregar_razon_social, cargar_padron_proveedores, construir_df_motor, parse_percepciones_txt,
)
from ui.components import (
    boton_descarga_xlsx, ejecutar_con_manejo_de_errores, mostrar_mapeo_erp,
    mostrar_metricas, mostrar_resultados,
)

TITULO = "🧾 Conciliador de Percepciones IIBB"
ERP_FIJO = "Bejerman"


def _sidebar():
    with st.sidebar:
        st.header("⚙️ Configuración del Cruce")
        st.caption(f"ERP del Mayor: **{ERP_FIJO}** (fijo en este módulo)")
        st.markdown("---")
        mostrar_mapeo_erp(ERP_CONFIGS[ERP_FIJO], modo_total=True)


def _leer_txt_percepciones(file_txt):
    """Parsea el .txt cargado. Devuelve el DataFrame o None (mostrando el error en pantalla)."""
    try:
        df_txt = parse_percepciones_txt(file_txt.getvalue())
    except Exception as e:
        st.error(f"❌ No se pudo interpretar el .txt de percepciones: {e}")
        return None

    if df_txt.empty:
        st.error("❌ No se encontraron líneas con el formato esperado en el .txt de percepciones.")
        return None
    return df_txt


def _selector_jurisdiccion(df_txt):
    resumen = df_txt.groupby("Jurisdicción")["Monto"].agg(["count", "sum"])

    def _fmt(j):
        return f"{j}  —  {int(resumen.loc[j, 'count'])} registros  —  $ {resumen.loc[j, 'sum']:,.2f}"

    jurisdiccion = st.selectbox(
        "¿Qué jurisdicción necesitás conciliar?",
        options=sorted(resumen.index.tolist()),
        format_func=_fmt,
    )

    with st.expander("👁️ Ver el .txt ya formateado (todas las jurisdicciones)"):
        st.dataframe(df_txt, use_container_width=True)
        buf = io.BytesIO()
        df_txt.to_excel(buf, index=False, sheet_name="Percepciones")
        buf.seek(0)
        st.download_button(
            "📥 Descargar .txt formateado (.xlsx)",
            data=buf,
            file_name="Percepciones_Formateadas.xlsx",
            mime=MIME_XLSX,
        )
    return jurisdiccion


def _procesar(file_mayor, df_txt, file_prov, jurisdiccion):
    padron = cargar_padron_proveedores(file_prov.getvalue())
    df_j = df_txt[df_txt["Jurisdicción"] == jurisdiccion].copy()
    df_j = agregar_razon_social(df_j, padron)
    df_motor = construir_df_motor(df_j)

    excel_bytes, stats = procesar_archivos(
        file_mayor,
        None,
        ERP_FIJO,
        df_arca_override=df_motor,
        origen_label="IIBB",
        concepto="PERCEPCIONES IIBB",
        subtitulo=f"JURISDICCIÓN {jurisdiccion}",
        priorizar_total_mayor=True,
    )

    st.success(f"✅ ¡Conciliación de la jurisdicción {jurisdiccion} completada!")

    sin_rs = df_j[df_j["Razón Social"] == SIN_RAZON_SOCIAL]["CUIT"].unique().tolist()
    if sin_rs:
        st.warning(
            f"⚠️ {len(sin_rs)} CUIT del .txt no figuran en el listado de proveedores "
            f"(se informan como '{SIN_RAZON_SOCIAL}'): " + ", ".join(sin_rs)
        )

    mostrar_metricas(stats, "IIBB")
    boton_descarga_xlsx(
        "📥 Descargar Reporte Conciliado en Excel",
        excel_bytes,
        f"Conciliacion_Percepciones_IIBB_{jurisdiccion}.xlsx",
    )
    mostrar_resultados(stats, "IIBB", "🎉 ¡No hay percepciones pendientes en el listado!")


def render():
    st.title("🧾 Conciliador de Percepciones IIBB (Listado del sistema vs Mayor Bejerman)")
    st.markdown(
        "Carga el **Mayor de Bejerman** (una sola jurisdicción por reporte), el **.txt de percepciones** del mes "
        "y el **listado de CUIT Proveedores**. Luego elegí la jurisdicción a conciliar."
    )
    st.warning(
        "⚠️ El Mayor debe estar en formato **.xlsx**. El cruce usa la columna **ImpTotal_Loc** "
        "(las NC van en negativo).",
        icon="⚠️",
    )

    _sidebar()

    c1, c2, c3 = st.columns(3)
    with c1:
        file_mayor = st.file_uploader("Subir Mayor Bejerman (.xlsx)", type=["xlsx"], key="mayor_iibb")
    with c2:
        file_txt = st.file_uploader("Subir listado de percepciones (.txt)", type=["txt"], key="txt_iibb")
    with c3:
        file_prov = st.file_uploader("Subir CUIT Proveedores (.txt)", type=["txt"], key="prov_iibb")

    df_txt = _leer_txt_percepciones(file_txt) if file_txt else None
    jurisdiccion = _selector_jurisdiccion(df_txt) if df_txt is not None else None

    if file_mayor and df_txt is not None and file_prov and jurisdiccion:
        st.info(f"💡 Listo para conciliar la jurisdicción **{jurisdiccion}** contra el Mayor de Bejerman.")

        if st.button("⚙️ Procesar Conciliación", type="primary", use_container_width=True):
            with st.spinner("Procesando percepciones y aplicando cruce..."):
                ejecutar_con_manejo_de_errores(
                    lambda: _procesar(file_mayor, df_txt, file_prov, jurisdiccion)
                )

# Dependencias: pip install streamlit pandas openpyxl lxml html5lib xlrd
# Ejecutar con:  streamlit run app.py
"""Punto de entrada. Solo configura la página y despacha al módulo elegido."""
import streamlit as st

from ui import formatear_page, percepciones_page, retenciones_page

st.set_page_config(
    page_title="Conciliación Contable Multiorigen",
    page_icon="📊",
    layout="wide",
)

# Registro de módulos: para sumar uno nuevo, crear ui/<nombre>_page.py con TITULO y render()
# y agregarlo a esta lista.
PAGINAS = [retenciones_page, percepciones_page, formatear_page]


def main():
    st.sidebar.title("🛠️ Herramientas")
    titulos = {p.TITULO: p for p in PAGINAS}
    elegido = st.sidebar.radio("Selecciona un módulo:", list(titulos.keys()))
    st.sidebar.markdown("---")
    titulos[elegido].render()


if __name__ == "__main__":
    main()

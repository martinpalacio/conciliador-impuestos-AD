"""Lectura y validación de archivos de entrada."""
import pandas as pd


def decodificar_texto(raw):
    """Decodifica bytes de un txt: UTF-8 y, si falla, Latin-1 (típico de Bejerman/ERPs)."""
    for enc in ("utf-8-sig", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1", errors="ignore")


def validar_formato_archivo(uploaded_file):
    """
    Verifica que el archivo sea un Excel válido (.xlsx).
    Si es .xls o no se puede leer, lanza un error amigable.
    """
    filename = uploaded_file.name.lower()

    if filename.endswith(".xls"):
        raise ValueError(
            f"El archivo '{uploaded_file.name}' está en formato '.xls' (formato antiguo de Excel o archivo web de AFIP/ARCA). "
            f"Por favor, abre el archivo en Excel y guárdalo usando la opción 'Guardar como...' eligiendo el formato "
            f"'Libro de Excel (*.xlsx)' antes de subirlo aquí."
        )

    # Verificamos si realmente es un xlsx intentando leer una pequeña parte
    try:
        pd.read_excel(uploaded_file, nrows=1)
    except Exception:
        raise ValueError(
            f"El archivo '{uploaded_file.name}' parece estar corrupto o no es un formato de Excel válido (.xlsx). "
            f"Si lo descargaste de ARCA/AFIP, por favor ábrelo en Excel y guárdalo como 'Libro de Excel (*.xlsx)'."
        )

    # Reseteamos el puntero del archivo después de la prueba de lectura
    uploaded_file.seek(0)
    return uploaded_file

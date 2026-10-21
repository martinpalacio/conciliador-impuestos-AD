"""Parser del Mayor crudo de AutoDealer (.xls / HTML / texto plano) → Excel tabulado."""
import io
import re

import pandas as pd

from core.io_utils import decodificar_texto

_RE_INICIO_LINEA = re.compile(r"^\d+\s+\d{2}/\d{2}/\d{4}")
_RE_REFERENCIA = re.compile(r"^[A-Z]{2}\s\d+/\d")
_RE_DETALLE = re.compile(r"^[A-Za-z0-9]+\-[A-Za-z0-9\-]+")
# Flexible: soporta formatos con comas, puntos o reconstruidos por pandas
_RE_MONTO = re.compile(r"\b(?:\d{1,3}(?:[.,]\d{3})+|\d+)[.,]\d{1,2}\b")
_RE_FECHA_ISO = re.compile(r"^\d{4}-\d{2}-\d{2}( \d{2}:\d{2}:\d{2})?$")


# --- 1. Lectura: cualquier formato de entrada → lista de líneas de texto -------------

def _detectar_tipo(raw):
    """El tipo se detecta por el contenido real (firma de bytes), no por la extensión."""
    es_excel_binario = raw[:4] == b"\xd0\xcf\x11\xe0" or raw[:2] == b"PK"  # .xls (OLE2) o .xlsx (zip)
    inicio = raw[:2048].lower()
    es_html = (not es_excel_binario) and (b"<table" in inicio or b"<html" in inicio)
    return es_excel_binario, es_html


def _leer_dataframe_crudo(raw, es_html):
    if es_html:
        return pd.read_html(io.StringIO(decodificar_texto(raw)))[0]
    try:
        # xlrd lee el formato viejo .xls
        return pd.read_excel(io.BytesIO(raw), header=None, dtype=str, engine="xlrd")
    except Exception:
        # Si no era .xls, probamos el motor por defecto (.xlsx)
        return pd.read_excel(io.BytesIO(raw), header=None, dtype=str)


def _normalizar_valor(val):
    """Devuelve el valor como texto limpio (fechas ISO → dd/mm/YYYY) o None si está vacío."""
    if not (pd.notna(val) and str(val).strip() and str(val).lower() != "nan"):
        return None
    v_str = str(val).strip()
    if _RE_FECHA_ISO.match(v_str):
        try:
            v_str = pd.to_datetime(v_str).strftime("%d/%m/%Y")
        except Exception:
            pass
    return v_str


def _dataframe_a_lineas(df_raw):
    """Reconstruye cada fila del dataframe como una línea de texto para el motor de Regex."""
    lines = []
    for _, row in df_raw.iterrows():
        row_vals = [v for v in (_normalizar_valor(x) for x in row.values) if v is not None]
        if row_vals:
            lines.append(" ".join(row_vals))
    return lines


def leer_lineas(raw):
    """Convierte el contenido crudo (Excel, HTML o texto plano) en una lista de líneas."""
    es_excel_binario, es_html = _detectar_tipo(raw)
    if es_excel_binario or es_html:
        return _dataframe_a_lineas(_leer_dataframe_crudo(raw, es_html))
    # Reporte de texto plano (muy común en ERPs viejos que exportan a .xls / .prn)
    return decodificar_texto(raw).splitlines()


# --- 2. Parseo: una línea de texto → un registro --------------------------------------

def _asignar_montos(montos):
    """Acomoda DEBE, HABER, SALDO según la cantidad de montos extraídos."""
    if len(montos) >= 3:
        return montos[-3], montos[-2], montos[-1]
    if len(montos) == 2:
        # Comportamiento común del PRN: asume DEBE y SALDO
        return montos[0], "", montos[1]
    if len(montos) == 1:
        return "", "", montos[0]
    return "", "", ""


def parsear_linea(line):
    """Interpreta una línea del Mayor. Devuelve un dict o None si no es un asiento."""
    line = line.strip()
    # Validamos que la línea empiece con un Asiento y una Fecha sí o sí
    if not _RE_INICIO_LINEA.search(line):
        return None

    asiento = re.search(r"^(\d+)", line).group(1)
    fecha = re.search(r"(\d{2}/\d{2}/\d{4})", line).group(1)
    resto = line.replace(asiento, "", 1).replace(fecha, "", 1).strip()

    # Referencia (Ej: PV 12132/0, US 12187/1)
    referencia = ""
    m_ref = _RE_REFERENCIA.search(resto)
    if m_ref:
        referencia = m_ref.group(0)
        resto = resto.replace(referencia, "", 1).strip()

    # Detalle / Recibo (Ej: RC-X-0100-00104458)
    detalle = ""
    m_det = _RE_DETALLE.search(resto)
    if m_det:
        detalle = m_det.group(0)
        resto = resto.replace(detalle, "", 1).strip()

    # Montos; se quitan del texto para aislar la Entidad y el Centro de Costo
    montos = _RE_MONTO.findall(resto)
    for m in montos:
        resto = resto.replace(m, "", 1)
    resto = resto.strip()

    # Centro de Costo (última palabra pegada al final) vs Entidad
    centro_costo = ""
    entidad = resto
    m_cc = re.search(r"\s+([A-Za-z]+)$", resto)
    if m_cc:
        centro_costo = m_cc.group(1)
        entidad = resto[:m_cc.start()].strip()
    entidad = re.sub(r"\s+", " ", entidad).strip()

    debe, haber, saldo = _asignar_montos(montos)

    return {
        "ASIENTO": asiento,
        "FECHA": fecha,
        "Referencia": referencia,
        "DETALLE": detalle,
        "REF": referencia.split("/")[0] if referencia else "",  # REF corta (sin lo posterior a la barra)
        "ENTIDAD": entidad,
        "DEBE": debe,
        "HABER": haber,
        "SALDO": saldo,
        "CENTRO DE COSTO": centro_costo,
    }


# --- 3. Salida --------------------------------------------------------------------------

def parse_autodealer_xls(file_bytes):
    """
    Interpreta un archivo .xls (reporte de texto disfrazado de Excel, HTML disfrazado
    o Excel binario real) y devuelve un BytesIO con el Mayor tabulado en .xlsx.
    """
    file_bytes.seek(0)
    raw = file_bytes.read()
    file_bytes.seek(0)

    registros = [r for r in (parsear_linea(l) for l in leer_lineas(raw)) if r is not None]
    df = pd.DataFrame(registros)

    output = io.BytesIO()
    df.to_excel(output, index=False, sheet_name="Mayor Formateado")
    output.seek(0)
    return output

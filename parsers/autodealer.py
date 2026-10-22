"""
Parser del Mayor crudo de AutoDealer → DataFrame tabulado + controles contables.

Dos caminos, elegidos por el CONTENIDO del archivo (no por la extensión):
  1) Por columnas (preferido): el export nativo (.xls binario / HTML) ya trae una columna por dato.
     Se leen directamente, sin aplanar a texto: es exacto y permite validar contra el saldo final.
  2) Por texto (fallback): reportes de texto plano (.prn / .txt) o layouts desconocidos, con regex.
"""
import io
import re
from dataclasses import dataclass, field

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


def leer_lineas(raw, df_raw=None):
    """Convierte el contenido crudo (Excel, HTML o texto plano) en una lista de líneas."""
    if df_raw is None:
        df_raw = _cargar_dataframe(raw)
    if df_raw is not None:
        return _dataframe_a_lineas(df_raw)
    # Reporte de texto plano (muy común en ERPs viejos que exportan a .xls / .prn)
    return decodificar_texto(raw).splitlines()


def _cargar_dataframe(raw):
    """DataFrame crudo (sin encabezados, todo texto) si el archivo es Excel/HTML; None si es texto plano."""
    es_excel_binario, es_html = _detectar_tipo(raw)
    if es_excel_binario or es_html:
        return _leer_dataframe_crudo(raw, es_html)
    return None


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


# --- 3. Camino por columnas (export nativo del ERP) ----------------------------------

COLUMNAS_SALIDA = [
    "ASIENTO", "FECHA", "Referencia", "DETALLE", "REF", "ENTIDAD",
    "DEBE", "HABER", "SALDO", "CENTRO DE COSTO",
]
_SEPARADOR_CONCEPTO = " / "          # con espacios: no rompe textos como "DJ IVA 02/2026"
_RE_ASIENTO = r"\d+(\.0+)?"
_RE_FECHA = r"(\d{2}/\d{2}/\d{4}|\d{4}-\d{2}-\d{2}.*)"

# Posiciones del export nativo (no trae encabezados)
_C_ASIENTO, _C_FECHA, _C_REFERENCIA, _C_CONCEPTO, _C_DEBE, _C_HABER, _C_SALDO_ERP, _C_CC = range(8)


def _filas_movimiento(df_raw):
    """Filas que son movimientos: asiento numérico en la col. 0 y fecha en la col. 1."""
    if df_raw.shape[1] < 8:
        return df_raw.iloc[0:0]
    c0 = df_raw[_C_ASIENTO].astype(str).str.strip()
    c1 = df_raw[_C_FECHA].astype(str).str.strip()
    return df_raw[c0.str.fullmatch(_RE_ASIENTO) & c1.str.fullmatch(_RE_FECHA)]


def _es_export_nativo(movs, umbral=0.9):
    """
    Confirma que el layout por posición es el esperado: en (casi) todos los movimientos el saldo del ERP
    es numérico y hay un importe en Debe o Haber. Evita aplicar el mapeo por columnas a tablas con otro diseño.
    """
    if len(movs) == 0:
        return False
    saldo_ok = movs[_C_SALDO_ERP].map(lambda v: _a_numero(v) is not None)
    importe_ok = movs[_C_DEBE].map(lambda v: _a_numero(v) is not None) | movs[_C_HABER].map(
        lambda v: _a_numero(v) is not None)
    return bool((saldo_ok & importe_ok).mean() >= umbral)


def _a_numero(val):
    """float redondeado a 2 decimales, o None si está vacío. (No usa safe_float: aquí el punto es decimal.)"""
    if val is None or pd.isna(val) or not str(val).strip():
        return None
    try:
        return round(float(str(val).strip()), 2)
    except ValueError:
        return None


def _a_texto(val):
    if val is None or pd.isna(val) or not str(val).strip() or str(val).lower() == "nan":
        return None
    return str(val).strip()


def dividir_concepto(concepto):
    """
    "RC-X-0100-00104458 / PV 12132/0 / SAN CRISTOBAL, Soc. Mutual"  →  (DETALLE, REF, ENTIDAD)
      DETALLE = comprobante, REF = referencia sin el sufijo "/n", ENTIDAD = resto.
    Si el concepto no tiene 3 partes (apertura, DDJJ IVA, etc.), DETALLE y ENTIDAD llevan el texto completo.
    """
    texto = (concepto or "").strip()
    partes = [p.strip() for p in texto.split(_SEPARADOR_CONCEPTO)]
    if len(partes) >= 3:
        detalle, ref, entidad = partes[0], partes[1], _SEPARADOR_CONCEPTO.join(partes[2:])
        return detalle, ref.split("/")[0].strip(), entidad, True
    return texto, "", texto, False


def _saldo_de(df_raw, etiqueta):
    """Valor de la fila 'SALDO INICIAL' / 'SALDO FINAL' del reporte (col. concepto → col. saldo)."""
    marca = df_raw[_C_CONCEPTO].astype(str).str.strip().str.upper().str.startswith(etiqueta)
    for val in df_raw.loc[marca, _C_SALDO_ERP]:
        n = _a_numero(val)
        if n is not None:
            return n
    return None


def parsear_columnas(df_raw):
    """Export nativo → (DataFrame de salida, saldo inicial del ERP, saldo final del ERP, filas sin 3 partes)."""
    movs = _filas_movimiento(df_raw)
    registros, sin_tres_partes = [], 0

    for _, r in movs.iterrows():
        detalle, ref, entidad, tiene_3 = dividir_concepto(_a_texto(r[_C_CONCEPTO]))
        sin_tres_partes += 0 if tiene_3 else 1
        debe, haber = _a_numero(r[_C_DEBE]), _a_numero(r[_C_HABER])
        registros.append({
            "ASIENTO": int(float(str(r[_C_ASIENTO]).strip())),
            "FECHA": _normalizar_valor(r[_C_FECHA]),
            "Referencia": _a_texto(r[_C_REFERENCIA]),
            "DETALLE": detalle or None,
            "REF": ref or None,
            "ENTIDAD": entidad or None,
            "DEBE": debe,
            "HABER": haber,
            "SALDO": round((debe or 0) - (haber or 0), 2),   # por movimiento, no el acumulado del ERP
            "CENTRO DE COSTO": _a_texto(r[_C_CC]),
        })

    df = pd.DataFrame(registros, columns=COLUMNAS_SALIDA)
    return df, _saldo_de(df_raw, "SALDO INICIAL"), _saldo_de(df_raw, "SALDO FINAL"), sin_tres_partes


def calcular_control(df, saldo_inicial, saldo_final, sin_tres_partes=0):
    """
    Control contable: saldo inicial + Σ Debe − Σ Haber debe igualar el saldo final del ERP.
    `ok` es None si el reporte no trae saldo final para comparar.
    """
    total_debe = round(float(df["DEBE"].fillna(0).sum()), 2)
    total_haber = round(float(df["HABER"].fillna(0).sum()), 2)
    calculado = round((saldo_inicial or 0) + total_debe - total_haber, 2)
    diferencia = None if saldo_final is None else round(calculado - saldo_final, 2)
    return {
        "movimientos": len(df),
        "total_debe": total_debe,
        "total_haber": total_haber,
        "saldo_inicial": saldo_inicial,
        "saldo_final": saldo_final,
        "saldo_calculado": calculado,
        "diferencia": diferencia,
        "ok": None if diferencia is None else abs(diferencia) < 0.01,
        "sin_tres_partes": sin_tres_partes,
    }


# --- 4. API pública ---------------------------------------------------------------------

@dataclass
class ResultadoMayor:
    df: pd.DataFrame
    metodo: str                      # "columnas" (exacto, con control) | "texto" (regex, sin control)
    control: dict = field(default_factory=dict)


def formatear_mayor(file_bytes):
    """Lee el Mayor crudo y devuelve un ResultadoMayor (DataFrame + método usado + control contable)."""
    file_bytes.seek(0)
    raw = file_bytes.read()
    file_bytes.seek(0)

    df_raw = _cargar_dataframe(raw)
    if df_raw is not None and _es_export_nativo(_filas_movimiento(df_raw)):
        df, ini, fin, sin_3 = parsear_columnas(df_raw)
        return ResultadoMayor(df, "columnas", calcular_control(df, ini, fin, sin_3))

    registros = [r for r in (parsear_linea(l) for l in leer_lineas(raw, df_raw)) if r is not None]
    return ResultadoMayor(pd.DataFrame(registros), "texto")


def parse_autodealer_xls(file_bytes):
    """Compatibilidad: devuelve directamente el BytesIO con el .xlsx formateado."""
    from reports.mayor_excel import generar_excel_mayor
    return generar_excel_mayor(formatear_mayor(file_bytes).df)
